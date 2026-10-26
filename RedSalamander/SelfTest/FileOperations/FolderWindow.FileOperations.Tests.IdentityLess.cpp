// I35 S4 (Specs/Plans/Done/FileOperations_CopyMoveUxAndReliabilityReview_2026-10-11.md, section 7):
// a Local location that opens an object but exposes no file identity (the NAS of report 3) deletes it
// by name with corroboration behind one explicit consent (D-7.3); an object it cannot open is never
// deleted by name; a Delete of a leftover such a location cannot identify is handed to that consent
// instead of the touch guard, while every other touch stays refused; and owned-stage cleanup keeps its
// own budget under the synchronous-I/O cancel watch (R7.4). A delete by name root is respected by every
// task admitted after it (queued or not), a folder's facts include its last-write time, another live
// task that takes the root refuses the delete up to the moment it runs, a failed open is never treated
// as identity-less, and an archive cleanup never deletes by name (the S4 red-team fixes).

namespace
{
using IdentityLessTask                                  = FolderWindow::FileOperationState::Task;
using IdentityLessPrompt                                = IdentityLessTask::ConflictPromptState;
using IdentityLessBucket                                = IdentityLessTask::ConflictBucket;
using IdentityLessAction                                = IdentityLessTask::ConflictAction;
using IdentityLessSeverity                              = FolderWindow::FileOperationState::DiagnosticSeverity;
using IdentityLessSummary                               = FolderWindow::FileOperationState::CompletedTaskSummary;
using IdentityLessDiagnostic                            = FolderWindow::FileOperationState::TaskDiagnosticEntry;
constexpr auto kIdentityLessWait                        = std::chrono::seconds(30);
constexpr std::wstring_view kIdentityLessRecycleFailEnv = L"REDSALAMANDER_FILEOPS_RECYCLE_FAIL_PATH";
constexpr std::wstring_view kIdentityLessStageDelayRoot = L"REDSALAMANDER_FILEOPS_STAGE_CLEANUP_DELAY_ROOT";
constexpr std::wstring_view kIdentityLessStageDelayFire = L"REDSALAMANDER_FILEOPS_STAGE_CLEANUP_DELAY_FIRED";
// The Local provider's BindObject fails the no-follow open of exactly this path with ERROR_NOT_SUPPORTED.
constexpr std::wstring_view kIdentityLessBindOpenNotSupportedEnv = L"REDSALAMANDER_FILEOPS_BIND_OPEN_NOT_SUPPORTED_PATH";
// The Local provider pauses a delete by name of exactly this root between the check of its listing and the
// first deletion: it sets ENTERED to 1 and waits until RELEASE is 1.
constexpr std::wstring_view kIdentityLessByNamePauseRootEnv    = L"REDSALAMANDER_FILEOPS_BYNAME_DELETE_PAUSE_ROOT";
constexpr std::wstring_view kIdentityLessByNamePauseEnteredEnv = L"REDSALAMANDER_FILEOPS_BYNAME_DELETE_PAUSE_ENTERED";
constexpr std::wstring_view kIdentityLessByNamePauseReleaseEnv = L"REDSALAMANDER_FILEOPS_BYNAME_DELETE_PAUSE_RELEASE";
// A test-only entry budget for the listing a delete by name pins.
constexpr std::wstring_view kIdentityLessByNameEntryBudgetEnv = L"REDSALAMANDER_FILEOPS_BYNAME_ENTRY_BUDGET";
// How long the provider pause waits for RELEASE (milliseconds; the case scales it with the timeout multiplier).
constexpr std::wstring_view kIdentityLessByNamePauseTimeoutEnv = L"REDSALAMANDER_FILEOPS_BYNAME_DELETE_PAUSE_TIMEOUT_MS";
// The listing of a delete by name reports no file index (as on a redirector without id listings).
constexpr std::wstring_view kIdentityLessByNameListWithoutIndexEnv = L"REDSALAMANDER_FILEOPS_BYNAME_LIST_WITHOUT_INDEX";
// A test-only number of folders a delete by name holds at once (its tree then goes in chunks).
constexpr std::wstring_view kIdentityLessByNameHoldBudgetEnv = L"REDSALAMANDER_FILEOPS_BYNAME_HOLD_BUDGET";

[[nodiscard]] std::optional<IdentityLessSummary> FindIdentityLessSummary(FolderWindow::FileOperationState* operations, uint64_t taskId)
{
    std::vector<IdentityLessSummary> summaries;
    operations->CollectCompletedTasks(summaries);
    const auto found = std::ranges::find(summaries, taskId, &IdentityLessSummary::taskId);
    if (found == summaries.end())
        return std::nullopt;
    return *found;
}

[[nodiscard]] size_t CountIdentityLessIssues(const IdentityLessSummary& summary,
                                             std::wstring_view category,
                                             IdentityLessSeverity severity,
                                             HRESULT status) noexcept
{
    return static_cast<size_t>(std::ranges::count_if(summary.issueDiagnostics, [&](const IdentityLessDiagnostic& issue) noexcept {
        return issue.category == category && issue.severity == severity && issue.status == status;
    }));
}

// A retained user note of the task, whatever its status (a Recycle escalation note keeps the Recycle's).
[[nodiscard]] size_t CountIdentityLessNotes(const IdentityLessSummary& summary, std::wstring_view category, UINT messageId) noexcept
{
    return static_cast<size_t>(std::ranges::count_if(summary.issueDiagnostics, [&](const IdentityLessDiagnostic& issue) noexcept {
        return issue.category == category && issue.severity == IdentityLessSeverity::Info && issue.userText.messageId == messageId;
    }));
}

// Technical entries (an untyped Info such as a refused delete by name) stay in the live log only.
[[nodiscard]] size_t CountLiveIdentityLessDiagnostics(
    FolderWindow::FileOperationState* operations, uint64_t taskId, std::wstring_view category, HRESULT status, std::wstring_view messageFragment = {})
{
    std::vector<IdentityLessDiagnostic> entries;
    operations->CollectDiagnostics(entries);
    return static_cast<size_t>(std::ranges::count_if(entries,
                                                     [&](const IdentityLessDiagnostic& entry) noexcept
    {
        return entry.taskId == taskId && entry.category == category && entry.status == status &&
               (messageFragment.empty() || entry.message.find(messageFragment) != std::wstring::npos);
    }));
}

[[nodiscard]] std::wstring DescribeIdentityLessIssues(const IdentityLessSummary& summary)
{
    std::wstring text;
    for (const IdentityLessDiagnostic& issue : summary.issueDiagnostics)
    {
        if (! text.empty())
            text += L"; ";
        text += std::format(
            L"{}:{}:0x{:08X}:{}", issue.category, static_cast<unsigned int>(issue.severity), static_cast<unsigned long>(issue.status), issue.message);
    }
    return text;
}

[[nodiscard]] std::wstring DescribeIdentityLessPrompt(const std::optional<IdentityLessPrompt>& prompt)
{
    if (! prompt.has_value())
        return L"<no question>";
    std::wstring actions;
    for (size_t index = 0u; index < prompt->actionCount; ++index)
        actions += std::format(L"{}{}", index == 0u ? L"" : L",", static_cast<unsigned int>(prompt->actions[index]));
    return std::format(L"bucket={} status=0x{:08X} deferred={} actions=[{}] default={} escape={} applyToAll={} note='{}'",
                       static_cast<unsigned int>(prompt->bucket),
                       static_cast<unsigned long>(prompt->status),
                       prompt->deferredConsent,
                       actions,
                       static_cast<unsigned int>(prompt->defaultAction),
                       static_cast<unsigned int>(prompt->escapeAction),
                       prompt->applyToAllEligible,
                       prompt->consentNote);
}

// The permanent-delete card of a root deleted by name: one explicit consent, "Delete by name" or
// Cancel (the default and Escape action), nothing to skip and no Apply to all.
[[nodiscard]] bool IsByNameCard(const IdentityLessPrompt& prompt) noexcept
{
    return prompt.deferredConsent && prompt.bucket == IdentityLessBucket::PermanentDeleteByNameConfirmation && prompt.actionCount == 2u &&
           PromptHasAction(prompt, IdentityLessAction::PermanentDelete) && PromptHasAction(prompt, IdentityLessAction::Cancel) &&
           prompt.defaultAction == IdentityLessAction::Cancel && prompt.escapeAction == IdentityLessAction::Cancel && ! prompt.applyToAllEligible;
}

// A Recycle escalation to a delete by name: Cancel default, "Delete by name", Skip, and only for a
// multi-item task the explicit "Delete all remaining by name".
[[nodiscard]] bool IsByNameEscalation(const IdentityLessPrompt& prompt, bool multiItem) noexcept
{
    return prompt.deferredConsent && prompt.bucket == IdentityLessBucket::RecycleFailedByName && PromptHasAction(prompt, IdentityLessAction::PermanentDelete) &&
           PromptHasAction(prompt, IdentityLessAction::Cancel) && PromptHasAction(prompt, IdentityLessAction::Skip) &&
           PromptHasAction(prompt, IdentityLessAction::PermanentDeleteAll) == multiItem && ! PromptHasAction(prompt, IdentityLessAction::SkipAll) &&
           prompt.defaultAction == IdentityLessAction::Cancel && ! prompt.applyToAllEligible;
}

[[nodiscard]] bool HasStageLeafIn(const std::filesystem::path& folder) noexcept
{
    std::error_code error;
    for (std::filesystem::directory_iterator entry(folder, error), end; ! error && entry != end; entry.increment(error))
    {
        if (entry->path().filename().native().find(L".rs_tmp_") != std::wstring::npos)
            return true;
    }
    return false;
}

// Owns what an identity-less case changes (auto-accept, prompt override, queue mode, injections,
// pause points, settings) and the records of the tasks it starts.
class IdentityLessCaseScope final
{
public:
    IdentityLessCaseScope(SelfTestState& state, FolderWindow::FileOperationState* operations) noexcept : _state(state), _operations(operations)
    {
        auto& settings  = EnsureFileOperationsSettingsForSelfTest();
        _verify         = settings.verifyAfterCopy;
        _infoEnabled    = settings.diagnosticsInfoEnabled;
        _autoAccept     = HostGetAutoAcceptPrompts();
        _autoDismiss    = operations->GetAutoDismissSuccess();
        _queueMode      = operations->GetQueueNewTasks();
        _originalPopup  = operations->GetPopupHwndForSelfTest();
        _promptOverride = HostGetTestPromptResultOverride();
        // Admission auto-accept would answer the by-name consent (permanentDeleteTestOverride).
        HostSetAutoAcceptPrompts(false);
        HostClearTestPromptResultOverride();
        settings.diagnosticsInfoEnabled = true;
        operations->SetAutoDismissSuccess(false);
        operations->ApplyQueueMode(false);
    }
    IdentityLessCaseScope(const IdentityLessCaseScope&)            = delete;
    IdentityLessCaseScope& operator=(const IdentityLessCaseScope&) = delete;
    ~IdentityLessCaseScope()
    {
        ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest();
        SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(false);
        ReleaseFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest();
        SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(false);
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
        // Released first: a delete paused by the provider seam must not outlive the case.
        static_cast<void>(SetEnvironmentVariableW(kIdentityLessByNamePauseReleaseEnv.data(), L"1"));
        for (const std::wstring_view name : {kSelfTestEnvBindIdentityUnsupportedPath,
                                             kIdentityLessRecycleFailEnv,
                                             kIdentityLessStageDelayRoot,
                                             kIdentityLessStageDelayFire,
                                             kIdentityLessBindOpenNotSupportedEnv,
                                             kIdentityLessByNamePauseRootEnv,
                                             kIdentityLessByNamePauseTimeoutEnv,
                                             kIdentityLessByNameListWithoutIndexEnv,
                                             kIdentityLessByNameHoldBudgetEnv,
                                             kIdentityLessByNameEntryBudgetEnv})
            static_cast<void>(SetEnvironmentVariableW(name.data(), nullptr));
        const bool quiet = CancelOutstandingSelfTestOperations(_operations);
        if (quiet)
        {
            for (const uint64_t taskId : _tasks)
            {
                _operations->DismissCompletedTask(taskId);
                _operations->DebugRemoveDiagnosticsForTask(taskId);
                _state.completedTasks.erase(taskId);
            }
        }
        auto& settings                  = EnsureFileOperationsSettingsForSelfTest();
        settings.verifyAfterCopy        = _verify;
        settings.diagnosticsInfoEnabled = _infoEnabled;
        _operations->ApplyQueueMode(_queueMode);
        _operations->SetAutoDismissSuccess(_autoDismiss);
        HostSetAutoAcceptPrompts(_autoAccept);
        if (_promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(_promptOverride);
        static_cast<void>(_operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
        if (quiet && ! _originalPopup && _operations->GetPopupHwndForSelfTest())
            _operations->DebugResetPopupForSelfTest();
    }

    // Every path at or below `location` binds as a redirector without FileIdInfo does (Unsupported).
    [[nodiscard]] static bool MakeIdentityLess(const std::filesystem::path& location) noexcept
    {
        return SetEnvironmentVariableW(kSelfTestEnvBindIdentityUnsupportedPath.data(), location.c_str()) != FALSE;
    }

    // Starts a Delete; `admissionAnswer` answers a host prompt at admission (the touch guard's) and
    // `waitForOthers` queues it behind running tasks as the Queue execution mode does.
    [[nodiscard]] std::optional<uint64_t> StartDelete(std::vector<std::filesystem::path> paths,
                                                      bool recycle,
                                                      HostPromptResult admissionAnswer = HOST_PROMPT_RESULT_NONE,
                                                      bool waitForOthers               = false)
    {
        if (admissionAnswer != HOST_PROMPT_RESULT_NONE)
            HostSetTestPromptResultOverride(admissionAnswer);
        const auto admitted = StartFileOperationAndGetId(_operations,
                                                         FILESYSTEM_DELETE,
                                                         FolderWindow::Pane::Left,
                                                         std::nullopt,
                                                         _state.fsLocal,
                                                         std::move(paths),
                                                         {},
                                                         // As the pane's Delete: recursive, to the Recycle Bin or not.
                                                         recycle ? static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN)
                                                                 : FILESYSTEM_FLAG_RECURSIVE,
                                                         waitForOthers,
                                                         0u,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
        HostClearTestPromptResultOverride();
        if (admitted.has_value())
            _tasks.push_back(admitted.value());
        return admitted;
    }

    void Track(uint64_t taskId)
    {
        _tasks.push_back(taskId);
    }

    // The task's actionable question, or nothing once it ended or the wait ran out. A question this scope
    // already answered stays published until the worker takes the answer; it is never returned again.
    [[nodiscard]] std::optional<IdentityLessPrompt> WaitForQuestion(uint64_t taskId, std::wstring_view what)
    {
        std::optional<IdentityLessPrompt> prompt;
        static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            prompt          = TryGetConflictPromptCopy(_operations->FindTask(taskId));
            const auto seen = _answeredGenerations.find(taskId);
            if (prompt.has_value() && seen != _answeredGenerations.end() && seen->second == prompt->generation)
                prompt.reset();
            return prompt.has_value() || _state.completedTasks.contains(taskId);
        },
            {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = what}));
        return prompt;
    }

    // Answers as the card does: an action under More (Delete all remaining by name) is submitted with
    // its overflow placement and no Apply-to-all state.
    [[nodiscard]] bool Answer(uint64_t taskId, const IdentityLessPrompt& prompt, IdentityLessAction action) noexcept
    {
        IdentityLessTask* const task = _operations->FindTask(taskId);
        const auto overflowEnd =
            prompt.overflowActions.begin() + static_cast<std::ptrdiff_t>(std::min(prompt.overflowActionCount, prompt.overflowActions.size()));
        const bool overflow = std::find(prompt.overflowActions.begin(), overflowEnd, action) != overflowEnd;
        const bool accepted = task != nullptr && task->SubmitConflictDecision(prompt.generation,
                                                                              action,
                                                                              overflow ? std::nullopt : std::optional<bool>(false),
                                                                              overflow ? FileOperations::ConflictActionPlacement::Overflow
                                                                                       : FileOperations::ConflictActionPlacement::Published);
        if (accepted)
            _answeredGenerations.insert_or_assign(taskId, prompt.generation);
        return accepted;
    }

    // The task's end, or nothing when it did not end in time.
    [[nodiscard]] const CompletedTaskInfo* WaitForEnd(uint64_t taskId, std::wstring_view what)
    {
        const auto ended = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return _state.completedTasks.contains(taskId) && _operations->FindTask(taskId) == nullptr;
        }, {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = what});
        const auto found = _state.completedTasks.find(taskId);
        return ended.conditionMet && found != _state.completedTasks.end() ? &found->second : nullptr;
    }

    // Whether the task is waiting to enter its operation (its queue place, an overlap edge or the interlock).
    [[nodiscard]] bool WaitUntilWaitingToEnter(uint64_t taskId, std::wstring_view what)
    {
        const auto waiting = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            IdentityLessTask* const task = _operations->FindTask(taskId);
            return (task != nullptr && task->IsWaitingInQueue()) || _state.completedTasks.contains(taskId);
        },
            {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = what});
        IdentityLessTask* const task = _operations->FindTask(taskId);
        return waiting.conditionMet && task != nullptr && task->IsWaitingInQueue() && ! _state.completedTasks.contains(taskId);
    }

    // Ends a task that may still ask questions (a refused item asks what to do next): each one is
    // answered Cancel and recorded in `asked`.
    [[nodiscard]] const CompletedTaskInfo* CancelQuestionsUntilEnd(uint64_t taskId, std::wstring_view what, std::vector<IdentityLessPrompt>* asked = nullptr)
    {
        for (size_t round = 0u; round < 4u; ++round)
        {
            const auto question = WaitForQuestion(taskId, what);
            if (! question.has_value())
                break;
            if (asked != nullptr)
                asked->push_back(question.value());
            if (! Answer(taskId, question.value(), IdentityLessAction::Cancel))
                break;
            static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                const auto next = TryGetConflictPromptCopy(_operations->FindTask(taskId));
                return ! next.has_value() || next->generation != question->generation || _state.completedTasks.contains(taskId);
            },
                {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = what}));
        }
        return WaitForEnd(taskId, what);
    }

private:
    SelfTestState& _state;
    FolderWindow::FileOperationState* _operations = nullptr;
    std::vector<uint64_t> _tasks;
    std::unordered_map<uint64_t, uint64_t> _answeredGenerations; // task id -> generation of its last answered question
    HWND _originalPopup              = nullptr;
    HostPromptResult _promptOverride = HOST_PROMPT_RESULT_NONE;
    bool _verify                     = false;
    bool _infoEnabled                = true;
    bool _autoAccept                 = false;
    bool _autoDismiss                = false;
    bool _queueMode                  = false;
};

// A running task T (a Recycle the Recycle Bin refuses, waiting on its question, on a drive-letter path that
// shares no identity domain with the loopback share) and a delete by name D of `target`, answered and
// waiting behind T in the queue (queue mode on).
struct QueuedByNameDelete final
{
    uint64_t running = 0u;
    IdentityLessPrompt runningQuestion{};
    uint64_t deletion = 0u;
};

// Empty when the fixture is in place, else why not.
[[nodiscard]] std::wstring StartQueuedByNameDelete(IdentityLessCaseScope& scope,
                                                   const std::filesystem::path& hold,
                                                   const std::filesystem::path& target,
                                                   QueuedByNameDelete& out)
{
    out = {};
    if (SetEnvironmentVariableW(kIdentityLessRecycleFailEnv.data(), hold.c_str()) == FALSE)
        return L"the Recycle failure injection could not be armed";
    const auto running = scope.StartDelete({hold}, true, HOST_PROMPT_RESULT_NONE, true);
    if (! running.has_value())
        return L"the running task was not admitted";
    const auto runningQuestion = scope.WaitForQuestion(running.value(), L"queued delete running task");
    if (! runningQuestion.has_value() || runningQuestion->bucket != IdentityLessBucket::RecycleFailed)
        return std::format(L"the running task did not wait on its Recycle question: {}", DescribeIdentityLessPrompt(runningQuestion));
    out.running         = running.value();
    out.runningQuestion = runningQuestion.value();
    const auto deletion = scope.StartDelete({target}, false, HOST_PROMPT_RESULT_NONE, true);
    if (! deletion.has_value())
        return L"the delete by name was not admitted";
    out.deletion    = deletion.value();
    const auto card = scope.WaitForQuestion(out.deletion, L"queued delete card");
    if (! card.has_value() || ! IsByNameCard(card.value()) || ! scope.Answer(out.deletion, card.value(), IdentityLessAction::PermanentDelete))
        return std::format(L"the delete by name did not ask its card: {}", DescribeIdentityLessPrompt(card));
    if (! scope.WaitUntilWaitingToEnter(out.deletion, L"queued delete in the queue"))
        return L"the delete by name did not wait behind the running task";
    return {};
}

// Ends T (Cancel on its question), then lets D run to its end; D's questions are answered Cancel.
[[nodiscard]] const CompletedTaskInfo* FinishQueuedByNameDelete(IdentityLessCaseScope& scope,
                                                                const QueuedByNameDelete& fixture,
                                                                std::vector<IdentityLessPrompt>* deleteQuestions = nullptr)
{
    if (! scope.Answer(fixture.running, fixture.runningQuestion, IdentityLessAction::Cancel) ||
        scope.WaitForEnd(fixture.running, L"queued delete running task end") == nullptr)
        return nullptr;
    static_cast<void>(SetEnvironmentVariableW(kIdentityLessRecycleFailEnv.data(), nullptr));
    return scope.CancelQuestionsUntilEnd(fixture.deletion, L"queued delete end", deleteQuestions);
}

// The overlap question of a task that works in a folder another task deletes by name: it names that
// delete with the by-name problem, defaults to Queue after these tasks and offers no Run at the same time.
[[nodiscard]] bool IsByNameOverlapQuestion(const IdentityLessPrompt& prompt, uint64_t deletion, FileOperations::SameHostOverlapProblem problem) noexcept
{
    return prompt.bucket == IdentityLessBucket::SameHostOverlap && prompt.overlapTaskId == deletion && prompt.overlapProblem == static_cast<uint8_t>(problem) &&
           prompt.defaultAction == IdentityLessAction::Proceed && PromptHasAction(prompt, IdentityLessAction::Proceed) &&
           PromptHasAction(prompt, IdentityLessAction::Cancel) && ! PromptHasAction(prompt, IdentityLessAction::RunConcurrently);
}

[[nodiscard]] std::wstring DescribeOverlapQuestion(const std::optional<IdentityLessPrompt>& prompt)
{
    if (! prompt.has_value())
        return L"<no question>";
    return std::format(
        L"{} overlapTask={} problem={}", DescribeIdentityLessPrompt(prompt), prompt->overlapTaskId, static_cast<unsigned int>(prompt->overlapProblem));
}

// Opens the file the way another program holding it for reading does: no write or delete sharing.
[[nodiscard]] wil::unique_hfile HoldForReadingOnly(const std::filesystem::path& file) noexcept
{
    return wil::unique_hfile(
        CreateFileW(file.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS, nullptr));
}

[[nodiscard]] bool AppendIdentityLessBytes(const std::filesystem::path& file, size_t bytes) noexcept
{
    wil::unique_hfile handle(
        CreateFileW(file.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, 0u, nullptr));
    const std::string payload(bytes, 'g');
    DWORD written = 0u;
    return handle && WriteFile(handle.get(), payload.data(), static_cast<DWORD>(payload.size()), &written, nullptr) != FALSE &&
           written == static_cast<DWORD>(payload.size());
}

// Moves the last-write time one hour back without changing the size.
[[nodiscard]] bool ShiftIdentityLessLastWrite(const std::filesystem::path& file) noexcept
{
    wil::unique_hfile handle(CreateFileW(file.c_str(),
                                         FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES,
                                         FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                         nullptr,
                                         OPEN_EXISTING,
                                         0u,
                                         nullptr));
    FILETIME creation{};
    FILETIME access{};
    FILETIME write{};
    if (! handle || GetFileTime(handle.get(), &creation, &access, &write) == FALSE)
        return false;
    ULARGE_INTEGER ticks{};
    ticks.LowPart  = write.dwLowDateTime;
    ticks.HighPart = write.dwHighDateTime;
    ticks.QuadPart -= 36'000'000'000ull;
    write.dwLowDateTime  = ticks.LowPart;
    write.dwHighDateTime = ticks.HighPart;
    return SetFileTime(handle.get(), nullptr, nullptr, &write) != FALSE;
}
} // namespace

// R7.3 (D-7.3), plan 7.3 `Delete_IdentityLessLocationByName`: on the loopback alias with the bind
// injection active (the drive-letter spelling when the share is unreachable), a permanent Delete asks
// one "Delete by name" consent; Cancel deletes nothing, Delete by name removes the item and records it
// as by name; a size or last-write change after the answer refuses with ERROR_REVISION_MISMATCH and
// deletes nothing; a planted leftover is handed to that consent (no guard prompt, the leftover warning
// in the consent); a Recycle the location refuses escalates to the same route behind its own consent.
[[nodiscard]] bool RunDeleteIdentityLessLocationByNameChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Identity-less delete by name: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-by-name";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    // Fixtures are written and checked through the drive letter; the operations run on `location`.
    const bool viaAlias                  = uncStatus == LoopbackShareStatus::Reachable;
    const std::filesystem::path location = viaAlias ? LoopbackShareAlias(root) : root;
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location))
        return reject(L"the identity-less injection could not be configured");
    AppendLog(std::format(L"Identity-less delete by name: running on {} ('{}').", viaAlias ? L"the loopback alias" : L"the drive letter", location.native()));

    const std::wstring checkedNote  = LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BY_NAME_CHECKED);
    const std::wstring leftoverNote = LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_POSSIBLE_LEFTOVER);
    const std::wstring byNameLabel  = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_DELETE_BY_NAME);
    const std::wstring allLabel     = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_DELETE_ALL_BY_NAME);
    if (checkedNote.empty() || leftoverNote.empty() || byNameLabel.empty() || allLabel.empty() ||
        FileOperations::FormatConflictActionLabel(
            IdentityLessAction::PermanentDelete, IdentityLessBucket::PermanentDeleteByNameConfirmation, FILESYSTEM_DELETE) != byNameLabel ||
        FileOperations::FormatConflictActionLabel(IdentityLessAction::PermanentDelete, IdentityLessBucket::RecycleFailedByName, FILESYSTEM_DELETE) !=
            byNameLabel ||
        FileOperations::FormatConflictActionLabel(IdentityLessAction::PermanentDeleteAll, IdentityLessBucket::RecycleFailedByName, FILESYSTEM_DELETE) !=
            allLabel)
        return reject(L"the by-name consent texts or action labels are missing");

    // 1. Cancel on the one by-name consent deletes nothing, and no touch-guard prompt was asked.
    const std::filesystem::path plain = root / L"plain.bin";
    if (! WriteFilledTestFile(plain, 4096u, 0x5Au))
        return reject(L"plain fixture unavailable");
    HostResetTestPromptRequestCount();
    const auto canceledTask = scope.StartDelete({location / plain.filename()}, false);
    if (! canceledTask.has_value())
        return reject(L"the by-name Delete was not admitted");
    const auto cancelCard = scope.WaitForQuestion(canceledTask.value(), L"identity-less cancel card");
    if (! cancelCard.has_value() || ! IsByNameCard(cancelCard.value()) || cancelCard->consentNote.find(checkedNote) == std::wstring::npos ||
        cancelCard->consentNote.find(leftoverNote) != std::wstring::npos)
        return reject(std::format(L"the by-name card contract does not hold: {}", DescribeIdentityLessPrompt(cancelCard)));
    if (! scope.Answer(canceledTask.value(), cancelCard.value(), IdentityLessAction::Cancel))
        return reject(L"Cancel on the by-name card was rejected");
    const CompletedTaskInfo* canceled = scope.WaitForEnd(canceledTask.value(), L"identity-less canceled Delete");
    if (canceled == nullptr || canceled->hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || ! FileSizeEquals(plain, 4096u) || HostGetTestPromptRequestCount() != 0u)
        return reject(std::format(L"Cancel on the by-name card did not leave the file: hr=0x{:08X} exists={} prompts={}",
                                  canceled != nullptr ? static_cast<unsigned long>(canceled->hr) : 0ul,
                                  std::filesystem::exists(plain),
                                  HostGetTestPromptRequestCount()));

    // 2. "Delete by name" removes it; the task records one by-name note and counts no problem.
    const auto removedTask = scope.StartDelete({location / plain.filename()}, false);
    if (! removedTask.has_value())
        return reject(L"the second by-name Delete was not admitted");
    const auto removeCard = scope.WaitForQuestion(removedTask.value(), L"identity-less delete card");
    if (! removeCard.has_value() || ! IsByNameCard(removeCard.value()) ||
        ! scope.Answer(removedTask.value(), removeCard.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"Delete by name could not be answered: {}", DescribeIdentityLessPrompt(removeCard)));
    const CompletedTaskInfo* removed = scope.WaitForEnd(removedTask.value(), L"identity-less Delete by name");
    const auto removedSummary        = FindIdentityLessSummary(operations, removedTask.value());
    if (removed == nullptr || removed->hr != S_OK || std::filesystem::exists(plain) || ! removedSummary.has_value() || removedSummary->warningCount != 0u ||
        removedSummary->errorCount != 0u ||
        CountIdentityLessIssues(removedSummary.value(), L"delete.permanent.byName", IdentityLessSeverity::Info, S_OK) != 1u ||
        ! std::ranges::any_of(removedSummary->issueDiagnostics,
                              [](const IdentityLessDiagnostic& issue) noexcept { return issue.userText.messageId == IDS_FILEOPS_NOTE_DELETED_BY_NAME; }))
        return reject(std::format(L"Delete by name did not remove the file with one by-name note: hr=0x{:08X} exists={} issues=[{}]",
                                  removed != nullptr ? static_cast<unsigned long>(removed->hr) : 0ul,
                                  std::filesystem::exists(plain),
                                  removedSummary.has_value() ? DescribeIdentityLessIssues(removedSummary.value()) : std::wstring(L"<no summary>")));

    // 3. A size change after the answer, on the handle that deletes: ERROR_REVISION_MISMATCH, nothing deleted.
    const std::filesystem::path grown = root / L"grown.bin";
    if (! WriteFilledTestFile(grown, 4096u, 0x11u))
        return reject(L"size-change fixture unavailable");
    SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(true);
    const auto grownTask = scope.StartDelete({location / grown.filename()}, false);
    if (! grownTask.has_value())
        return reject(L"the size-change Delete was not admitted");
    const auto grownCard = scope.WaitForQuestion(grownTask.value(), L"identity-less size-change card");
    if (! grownCard.has_value() || ! IsByNameCard(grownCard.value()) ||
        ! scope.Answer(grownTask.value(), grownCard.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the size-change card could not be answered: {}", DescribeIdentityLessPrompt(grownCard)));
    const auto paused = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return HasFileOpsPermanentDeleteBeforeLiveOutputGuardPauseEnteredForSelfTest() || state.completedTasks.contains(grownTask.value());
    }, {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = L"identity-less size-change pause"});
    const bool grewInWindow =
        paused.conditionMet && HasFileOpsPermanentDeleteBeforeLiveOutputGuardPauseEnteredForSelfTest() && AppendIdentityLessBytes(grown, 16u);
    ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest();
    SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(false);
    if (! grewInWindow)
        return reject(L"the size could not be changed between the answer and the delete");
    // The refused item asks what to do next (its known non-commit); Cancel ends the task.
    const auto refusedQuestion = scope.WaitForQuestion(grownTask.value(), L"identity-less size-change refusal");
    if (refusedQuestion.has_value() && (refusedQuestion->status != HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) ||
                                        PromptHasAction(refusedQuestion.value(), IdentityLessAction::PermanentDelete) ||
                                        ! scope.Answer(grownTask.value(), refusedQuestion.value(), IdentityLessAction::Cancel)))
        return reject(std::format(L"the refused delete by name asked the wrong question: {}", DescribeIdentityLessPrompt(refusedQuestion)));
    const CompletedTaskInfo* grownEnd = scope.WaitForEnd(grownTask.value(), L"identity-less size-change Delete");
    const auto grownSummary           = FindIdentityLessSummary(operations, grownTask.value());
    // The refusal names both results: the open succeeded and the compared facts differed.
    const std::wstring mismatchText =
        std::format(L"openHr=0x00000000 identityHr=0x{:08X}", static_cast<unsigned long>(HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH)));
    const size_t refusals =
        CountLiveIdentityLessDiagnostics(operations, grownTask.value(), L"delete.permanent.byName", HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), mismatchText);
    if (grownEnd == nullptr || SUCCEEDED(grownEnd->hr) || ! FileSizeEquals(grown, 4096u + 16u) || ! grownSummary.has_value() || refusals != 1u ||
        CountIdentityLessNotes(grownSummary.value(), L"delete.permanent.byName", IDS_FILEOPS_NOTE_DELETED_BY_NAME) != 0u)
        return reject(std::format(L"a size change after the answer was not refused with nothing deleted: hr=0x{:08X} exists={} refusals={} issues=[{}]",
                                  grownEnd != nullptr ? static_cast<unsigned long>(grownEnd->hr) : 0ul,
                                  std::filesystem::exists(grown),
                                  refusals,
                                  grownSummary.has_value() ? DescribeIdentityLessIssues(grownSummary.value()) : std::wstring(L"<no summary>")));

    // 4. A last-write change while the card waits: the re-check after the answer refuses, nothing deleted.
    const std::filesystem::path touched = root / L"touched.bin";
    if (! WriteFilledTestFile(touched, 4096u, 0x22u))
        return reject(L"last-write fixture unavailable");
    const auto touchedTask = scope.StartDelete({location / touched.filename()}, false);
    if (! touchedTask.has_value())
        return reject(L"the last-write Delete was not admitted");
    const auto touchedCard = scope.WaitForQuestion(touchedTask.value(), L"identity-less last-write card");
    if (! touchedCard.has_value() || ! IsByNameCard(touchedCard.value()) || ! ShiftIdentityLessLastWrite(touched) ||
        ! scope.Answer(touchedTask.value(), touchedCard.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the last-write card could not be answered after the change: {}", DescribeIdentityLessPrompt(touchedCard)));
    const CompletedTaskInfo* touchedEnd = scope.WaitForEnd(touchedTask.value(), L"identity-less last-write Delete");
    const auto touchedSummary           = FindIdentityLessSummary(operations, touchedTask.value());
    if (touchedEnd == nullptr || touchedEnd->hr != HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) || ! FileSizeEquals(touched, 4096u) ||
        ! touchedSummary.has_value() ||
        CountIdentityLessIssues(
            touchedSummary.value(), L"delete.confirmation.identityChanged", IdentityLessSeverity::Error, HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH)) != 1u)
        return reject(std::format(L"a last-write change before the answer was not refused: hr=0x{:08X} exists={} issues=[{}]",
                                  touchedEnd != nullptr ? static_cast<unsigned long>(touchedEnd->hr) : 0ul,
                                  std::filesystem::exists(touched),
                                  touchedSummary.has_value() ? DescribeIdentityLessIssues(touchedSummary.value()) : std::wstring(L"<no summary>")));

    // 5. A planted leftover goes to the same consent: no touch-guard prompt, the leftover warning in it.
    const std::filesystem::path leftover = root / L"movie.mkv.rs_tmp_9D380E53676EAB4315AE57F418F5049D_0";
    if (! WriteFilledTestFile(leftover, 8192u, 0x33u))
        return reject(L"leftover fixture unavailable");
    HostResetTestPromptRequestCount();
    const auto leftoverTask = scope.StartDelete({location / leftover.filename()}, false);
    if (! leftoverTask.has_value() || HostGetTestPromptRequestCount() != 0u)
        return reject(std::format(
            L"the leftover Delete was refused or asked the touch guard: admitted={} prompts={}", leftoverTask.has_value(), HostGetTestPromptRequestCount()));
    const auto leftoverCard = scope.WaitForQuestion(leftoverTask.value(), L"identity-less leftover card");
    if (! leftoverCard.has_value() || ! IsByNameCard(leftoverCard.value()) || leftoverCard->consentNote.find(checkedNote) == std::wstring::npos ||
        leftoverCard->consentNote.find(leftoverNote) == std::wstring::npos ||
        ! scope.Answer(leftoverTask.value(), leftoverCard.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the leftover card did not carry the leftover warning: {}", DescribeIdentityLessPrompt(leftoverCard)));
    const CompletedTaskInfo* leftoverEnd = scope.WaitForEnd(leftoverTask.value(), L"identity-less leftover Delete");
    if (leftoverEnd == nullptr || leftoverEnd->hr != S_OK || std::filesystem::exists(leftover) || HostGetTestPromptRequestCount() != 0u)
        return reject(std::format(L"the leftover was not deleted by name: hr=0x{:08X} exists={}",
                                  leftoverEnd != nullptr ? static_cast<unsigned long>(leftoverEnd->hr) : 0ul,
                                  std::filesystem::exists(leftover)));

    // 6. A refused Recycle escalates by name behind its own consent (one item: no "Delete all remaining").
    const std::filesystem::path recycled = root / L"recycled.bin";
    if (! WriteFilledTestFile(recycled, 4096u, 0x44u) ||
        SetEnvironmentVariableW(kIdentityLessRecycleFailEnv.data(), (location / recycled.filename()).c_str()) == FALSE)
        return reject(L"Recycle escalation fixture unavailable");
    const auto recycleTask = scope.StartDelete({location / recycled.filename()}, true);
    if (! recycleTask.has_value())
        return reject(L"the identity-less Recycle was not admitted");
    const auto escalation = scope.WaitForQuestion(recycleTask.value(), L"identity-less Recycle escalation");
    if (! escalation.has_value() || ! IsByNameEscalation(escalation.value(), false) || escalation->consentNote.find(checkedNote) == std::wstring::npos ||
        escalation->consentNote.find(leftoverNote) != std::wstring::npos ||
        ! scope.Answer(recycleTask.value(), escalation.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the Recycle escalation by name contract does not hold: {}", DescribeIdentityLessPrompt(escalation)));
    const CompletedTaskInfo* recycleEnd = scope.WaitForEnd(recycleTask.value(), L"identity-less Recycle escalation");
    const auto recycleSummary           = FindIdentityLessSummary(operations, recycleTask.value());
    if (recycleEnd == nullptr || recycleEnd->hr != S_OK || std::filesystem::exists(recycled) || ! recycleSummary.has_value() ||
        CountIdentityLessNotes(recycleSummary.value(), L"delete.recycle.byName", IDS_FILEOPS_NOTE_RECYCLE_ESCALATED_BY_NAME) != 1u)
        return reject(std::format(L"the Recycle escalation by name did not remove the item: hr=0x{:08X} exists={} issues=[{}]",
                                  recycleEnd != nullptr ? static_cast<unsigned long>(recycleEnd->hr) : 0ul,
                                  std::filesystem::exists(recycled),
                                  recycleSummary.has_value() ? DescribeIdentityLessIssues(recycleSummary.value()) : std::wstring(L"<no summary>")));

    // 7. Two refused Recycles: one question offers "Delete all remaining by name", which covers both.
    const std::filesystem::path firstRecycled  = root / L"first.bin";
    const std::filesystem::path secondRecycled = root / L"second.bin";
    const std::wstring failList = std::format(L"{}|{}", (location / firstRecycled.filename()).native(), (location / secondRecycled.filename()).native());
    if (! WriteFilledTestFile(firstRecycled, 2048u, 0x55u) || ! WriteFilledTestFile(secondRecycled, 2048u, 0x66u) ||
        SetEnvironmentVariableW(kIdentityLessRecycleFailEnv.data(), failList.c_str()) == FALSE)
        return reject(L"multi-item Recycle escalation fixture unavailable");
    const auto bulkTask = scope.StartDelete({location / firstRecycled.filename(), location / secondRecycled.filename()}, true);
    if (! bulkTask.has_value())
        return reject(L"the identity-less two-item Recycle was not admitted");
    const auto bulkEscalation = scope.WaitForQuestion(bulkTask.value(), L"identity-less two-item escalation");
    if (! bulkEscalation.has_value() || ! IsByNameEscalation(bulkEscalation.value(), true) ||
        ! scope.Answer(bulkTask.value(), bulkEscalation.value(), IdentityLessAction::PermanentDeleteAll))
        return reject(std::format(L"the two-item escalation did not offer Delete all remaining by name: {}", DescribeIdentityLessPrompt(bulkEscalation)));
    const CompletedTaskInfo* bulkEnd = scope.WaitForEnd(bulkTask.value(), L"identity-less two-item escalation");
    const auto bulkSummary           = FindIdentityLessSummary(operations, bulkTask.value());
    // One question for both items: the second escalation used the task-scoped grant.
    const size_t bulkQuestions =
        bulkSummary.has_value()
            ? static_cast<size_t>(std::ranges::count_if(bulkSummary->issueDiagnostics,
                                                        [](const IdentityLessDiagnostic& issue) noexcept { return issue.category == L"item.conflict.prompt"; }))
            : 0u;
    if (bulkEnd == nullptr || bulkEnd->hr != S_OK || bulkQuestions != 1u || std::filesystem::exists(firstRecycled) || std::filesystem::exists(secondRecycled) ||
        ! bulkSummary.has_value() || CountIdentityLessNotes(bulkSummary.value(), L"delete.recycle.byName", IDS_FILEOPS_NOTE_RECYCLE_ESCALATED_BY_NAME) != 2u)
        return reject(
            std::format(L"Delete all remaining by name did not cover both items with one question: hr=0x{:08X} prompts={} first={} second={} issues=[{}]",
                        bulkEnd != nullptr ? static_cast<unsigned long>(bulkEnd->hr) : 0ul,
                        bulkQuestions,
                        std::filesystem::exists(firstRecycled),
                        std::filesystem::exists(secondRecycled),
                        bulkSummary.has_value() ? DescribeIdentityLessIssues(bulkSummary.value()) : std::wstring(L"<no summary>")));

    // 8. A Recycle of a leftover: the guard warns (OK/Cancel, no receipt), the escalation repeats the warning.
    const std::filesystem::path recycledLeftover = root / L"clip.mp4.rs_tmp_0123456789ABCDEF0123456789ABCDEF_1";
    if (! WriteFilledTestFile(recycledLeftover, 4096u, 0x77u) ||
        SetEnvironmentVariableW(kIdentityLessRecycleFailEnv.data(), (location / recycledLeftover.filename()).c_str()) == FALSE)
        return reject(L"Recycle leftover fixture unavailable");
    HostResetTestPromptRequestCount();
    const auto leftoverRecycleTask = scope.StartDelete({location / recycledLeftover.filename()}, true, HOST_PROMPT_RESULT_OK);
    HostPromptDebugSnapshot guardPrompt{};
    if (! leftoverRecycleTask.has_value() || HostGetTestPromptRequestCount() != 1u || ! HostGetTestLastPromptDebugSnapshot(guardPrompt) ||
        guardPrompt.presentation != HOST_PROMPT_PRESENTATION_ARTIFACT_TOUCH || guardPrompt.buttons != HOST_PROMPT_BUTTONS_OK_CANCEL)
        return reject(std::format(L"the Recycle of a leftover was not warned about once with OK/Cancel: admitted={} prompts={} presentation={}",
                                  leftoverRecycleTask.has_value(),
                                  HostGetTestPromptRequestCount(),
                                  static_cast<unsigned int>(guardPrompt.presentation)));
    const auto leftoverEscalation = scope.WaitForQuestion(leftoverRecycleTask.value(), L"identity-less leftover escalation");
    if (! leftoverEscalation.has_value() || ! IsByNameEscalation(leftoverEscalation.value(), false) ||
        leftoverEscalation->consentNote.find(leftoverNote) == std::wstring::npos ||
        ! scope.Answer(leftoverRecycleTask.value(), leftoverEscalation.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the leftover escalation did not repeat the leftover warning: {}", DescribeIdentityLessPrompt(leftoverEscalation)));
    const CompletedTaskInfo* leftoverRecycleEnd = scope.WaitForEnd(leftoverRecycleTask.value(), L"identity-less leftover escalation");
    if (leftoverRecycleEnd == nullptr || leftoverRecycleEnd->hr != S_OK || std::filesystem::exists(recycledLeftover))
        return reject(std::format(L"the Recycle leftover was not deleted by name: hr=0x{:08X} exists={}",
                                  leftoverRecycleEnd != nullptr ? static_cast<unsigned long>(leftoverRecycleEnd->hr) : 0ul,
                                  std::filesystem::exists(recycledLeftover)));

    Debug::Perf::Emit(L"FileOps.SelfTest.IdentityLessDeleteByName", viaAlias ? L"loopback-alias" : L"drive-letter", 0u, 8u, viaAlias ? 1u : 0u, S_OK);
    return true;
}

// R7.1 remainder and R7.3: a root the location opened at the pin but another program then holds for
// reading only fails its delete open with a sharing violation reported with both results and Retry,
// never another route; a root the delete bind cannot open (Indeterminate) is never offered by name;
// and a root another live task holds is refused.
[[nodiscard]] bool RunDeleteIdentityLessSharingAndInterlockChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Identity-less sharing and interlock: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-sharing";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(root) : root;
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location))
        return reject(L"the identity-less injection could not be configured");
    const std::wstring sharingText = std::format(L"openHr=0x{:08X}", static_cast<unsigned long>(HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION)));

    // 1. The pin opened the file; a reader without delete sharing arrives while the card waits. The
    //    delete open then fails: both results are reported, the item offers Retry and no other route,
    //    and Retry after the reader left deletes it by name.
    const std::filesystem::path held = root / L"held-after-probe.bin";
    if (! WriteFilledTestFile(held, 4096u, 0x1Bu))
        return reject(L"probe-then-holder fixture unavailable");
    const auto heldTask = scope.StartDelete({location / held.filename()}, false);
    if (! heldTask.has_value())
        return reject(L"the probe-then-holder Delete was not admitted");
    const auto heldCard = scope.WaitForQuestion(heldTask.value(), L"probe-then-holder card");
    if (! heldCard.has_value() || ! IsByNameCard(heldCard.value()))
        return reject(std::format(L"the probe-then-holder Delete did not pin by name: {}", DescribeIdentityLessPrompt(heldCard)));
    wil::unique_hfile reader = HoldForReadingOnly(held);
    if (! reader || ! scope.Answer(heldTask.value(), heldCard.value(), IdentityLessAction::PermanentDelete))
        return reject(L"the reader could not hold the file, or Delete by name was rejected");
    const auto sharing = scope.WaitForQuestion(heldTask.value(), L"probe-then-holder sharing question");
    std::vector<IdentityLessDiagnostic> liveDiagnostics;
    operations->CollectDiagnostics(liveDiagnostics);
    const bool reportedBoth = std::ranges::any_of(liveDiagnostics,
                                                  [&](const IdentityLessDiagnostic& entry) noexcept
    {
        return entry.taskId == heldTask.value() && entry.category == L"delete.permanent.byName" &&
               entry.status == HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) && entry.message.find(sharingText) != std::wstring::npos &&
               entry.message.find(L"identityHr=") != std::wstring::npos;
    });
    if (! sharing.has_value() || sharing->bucket != IdentityLessBucket::SharingViolation || sharing->status != HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) ||
        ! PromptHasAction(sharing.value(), IdentityLessAction::Retry) || PromptHasAction(sharing.value(), IdentityLessAction::PermanentDelete) ||
        ! reportedBoth || ! FileSizeEquals(held, 4096u))
        return reject(std::format(L"the sharing failure after the probe was not reported with both results and Retry: reported={} exists={} {}",
                                  reportedBoth,
                                  std::filesystem::exists(held),
                                  DescribeIdentityLessPrompt(sharing)));
    reader.reset();
    if (! scope.Answer(heldTask.value(), sharing.value(), IdentityLessAction::Retry))
        return reject(L"Retry on the sharing question was rejected");
    const CompletedTaskInfo* heldEnd = scope.WaitForEnd(heldTask.value(), L"probe-then-holder Retry");
    if (heldEnd == nullptr || heldEnd->hr != S_OK || std::filesystem::exists(held))
        return reject(std::format(L"Retry after the reader left did not delete by name: hr=0x{:08X} exists={}",
                                  heldEnd != nullptr ? static_cast<unsigned long>(heldEnd->hr) : 0ul,
                                  std::filesystem::exists(held)));

    // 2. A root the delete bind cannot open is Indeterminate: never offered by name, both results logged.
    const std::filesystem::path locked = root / L"locked.bin";
    if (! WriteFilledTestFile(locked, 4096u, 0x2Cu))
        return reject(L"Indeterminate fixture unavailable");
    wil::unique_hfile lockedReader = HoldForReadingOnly(locked);
    if (! lockedReader)
        return reject(L"the reader could not hold the Indeterminate fixture");
    const auto lockedTask = scope.StartDelete({location / locked.filename()}, false);
    if (! lockedTask.has_value())
        return reject(L"the Indeterminate Delete was not admitted");
    const auto lockedQuestion = scope.WaitForQuestion(lockedTask.value(), L"Indeterminate Delete");
    if (lockedQuestion.has_value())
        return reject(std::format(L"an Indeterminate root asked a question: {}", DescribeIdentityLessPrompt(lockedQuestion)));
    const CompletedTaskInfo* lockedEnd = scope.WaitForEnd(lockedTask.value(), L"Indeterminate Delete");
    const auto lockedSummary           = FindIdentityLessSummary(operations, lockedTask.value());
    const bool lockedLogged            = lockedSummary.has_value() && std::ranges::any_of(lockedSummary->issueDiagnostics,
                                                                                          [&](const IdentityLessDiagnostic& issue) noexcept
    {
        return issue.category == L"preparation.deleteRoot" && issue.status == HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) &&
               issue.message.find(sharingText) != std::wstring::npos && issue.message.find(L"identityHr=not queried") != std::wstring::npos;
    });
    lockedReader.reset();
    if (lockedEnd == nullptr || lockedEnd->hr != HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) || ! lockedLogged || ! FileSizeEquals(locked, 4096u))
        return reject(std::format(L"an Indeterminate root was not refused with both results: hr=0x{:08X} logged={} exists={} issues=[{}]",
                                  lockedEnd != nullptr ? static_cast<unsigned long>(lockedEnd->hr) : 0ul,
                                  lockedLogged,
                                  std::filesystem::exists(locked),
                                  lockedSummary.has_value() ? DescribeIdentityLessIssues(lockedSummary.value()) : std::wstring(L"<no summary>")));

    // 3. Another live task holds the path (a Recycle waiting on its escalation question): a delete by
    //    name of it is refused at Preparing; the holder's question is then canceled, nothing deleted.
    const std::filesystem::path shared = root / L"shared.bin";
    if (! WriteFilledTestFile(shared, 4096u, 0x3Du) ||
        SetEnvironmentVariableW(kIdentityLessRecycleFailEnv.data(), (location / shared.filename()).c_str()) == FALSE)
        return reject(L"interlock fixture unavailable");
    const auto holderTask = scope.StartDelete({location / shared.filename()}, true);
    if (! holderTask.has_value())
        return reject(L"the holding Recycle was not admitted");
    const auto holderQuestion = scope.WaitForQuestion(holderTask.value(), L"interlock holder question");
    if (! holderQuestion.has_value() || ! IsByNameEscalation(holderQuestion.value(), false))
        return reject(std::format(L"the holding Recycle did not wait on its escalation: {}", DescribeIdentityLessPrompt(holderQuestion)));
    const auto refusedTask = scope.StartDelete({location / shared.filename()}, false);
    if (! refusedTask.has_value())
        return reject(L"the second Delete was not admitted");
    const auto refusedQuestion = scope.WaitForQuestion(refusedTask.value(), L"interlock refused Delete");
    if (refusedQuestion.has_value())
        return reject(std::format(L"a delete by name of a held path asked a question: {}", DescribeIdentityLessPrompt(refusedQuestion)));
    const CompletedTaskInfo* refusedEnd = scope.WaitForEnd(refusedTask.value(), L"interlock refused Delete");
    const auto refusedSummary           = FindIdentityLessSummary(operations, refusedTask.value());
    if (refusedEnd == nullptr || refusedEnd->hr != HRESULT_FROM_WIN32(ERROR_BUSY) || ! refusedSummary.has_value() ||
        CountIdentityLessIssues(refusedSummary.value(), L"delete.byName.heldByAnotherTask", IdentityLessSeverity::Error, HRESULT_FROM_WIN32(ERROR_BUSY)) !=
            1u ||
        ! FileSizeEquals(shared, 4096u))
        return reject(std::format(L"a delete by name of a path another task holds was not refused: hr=0x{:08X} exists={} issues=[{}]",
                                  refusedEnd != nullptr ? static_cast<unsigned long>(refusedEnd->hr) : 0ul,
                                  std::filesystem::exists(shared),
                                  refusedSummary.has_value() ? DescribeIdentityLessIssues(refusedSummary.value()) : std::wstring(L"<no summary>")));
    if (! scope.Answer(holderTask.value(), holderQuestion.value(), IdentityLessAction::Cancel))
        return reject(L"Cancel on the holder's escalation was rejected");
    const CompletedTaskInfo* holderEnd = scope.WaitForEnd(holderTask.value(), L"interlock holder");
    if (holderEnd == nullptr || holderEnd->hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || ! FileSizeEquals(shared, 4096u))
        return reject(std::format(L"the holder's canceled escalation changed the file: hr=0x{:08X} exists={}",
                                  holderEnd != nullptr ? static_cast<unsigned long>(holderEnd->hr) : 0ul,
                                  std::filesystem::exists(shared)));
    Debug::Perf::Emit(L"FileOps.SelfTest.IdentityLessSharingAndInterlock", L"sharing=retry;indeterminate=refused;held=refused", 0u, 3u, 0u, S_OK);
    return true;
}

// R7.3/R7.6: the touch-guard partition. A Delete item that its location opened but cannot identify
// (any provider without object binding: FTP, SFTP, SCP, S3, MTP, Microsoft Drive; or the Local by-name
// route) is never a receipt item: a permanent Delete's card asks for it and a Recycle is warned about
// without a receipt. Every other touch of such an item, and any item whose open failed, stays refused,
// so registering the provider staging names (R7.6) makes none of their leftovers impossible to delete
// and none of them movable.
[[nodiscard]] bool RunFileOpsIdentityLessGuardPartitionChecks(SelfTestState& /*state*/) noexcept
{
    namespace Artifacts = FileOperationArtifacts;
    const auto reject   = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Identity-less guard partition: {}", reason));
        return false;
    };

    struct ShapeCase final
    {
        std::wstring_view leaf;
        bool possible;
        bool mayHoldOriginal;
    };
    constexpr std::array kShapes{
        ShapeCase{L".rs-bak-0123456789abcdef0123456789abcdef-report.docx", true, true},
        ShapeCase{L".rs-bak-0123456789ABCDEF0123456789ABCDEF", true, true},
        ShapeCase{L".redsalamander-upload-0123456789abcdef0123456789abcdef", true, false},
        ShapeCase{L".redsalamander-writer-0123456789ABCDEF0123456789ABCDEF", true, false},
        ShapeCase{L".redsalamander-rollback-0123456789abcdef0123456789abcdef", true, true},
        ShapeCase{L".rs_case_tmp_0123456789abcdef0123456789abcdef", true, true},
        ShapeCase{L".report.docx.rs-mtp-overwrite-0123456789abcdef0123456789abcdef.tmp", true, false},
        ShapeCase{L".report.docx.rs-mtp-create-0123456789abcdef0123456789abcdef.tmp", true, false},
        ShapeCase{L"movie.mkv.rs_tmp_9D380E53676EAB4315AE57F418F5049D_0", true, false},
        ShapeCase{L"report.docx.rs_bak_{01234567-89AB-CDEF-0123-456789ABCDEF}", true, true},
        // Near misses stay ordinary.
        ShapeCase{L".rs-bak-0123456789abcdef-report.docx", false, false},
        ShapeCase{L".rs-bak-0123456789abcdef0123456789abcdefreport.docx", false, false},
        ShapeCase{L".redsalamander-upload-0123456789abcdef0123456789abcdeg", false, false},
        ShapeCase{L".redsalamander-upload-0123456789abcdef0123456789abcdef.txt", false, false},
        ShapeCase{L"redsalamander-upload-0123456789abcdef0123456789abcdef", false, false},
        ShapeCase{L".report.docx.rs-mtp-overwrite-0123456789abcdef.tmp", false, false},
        ShapeCase{L".report.docx.rs-mtp-create-0123456789abcdef0123456789abcdef.bin", false, false},
        ShapeCase{L"notes.rs_case_tmp_0123", false, false},
        // The case-only rename's temporary name is the whole leaf: no prefix, no suffix, 32 hex digits.
        ShapeCase{L"notes.rs_case_tmp_0123456789abcdef0123456789abcdef", false, false},
        ShapeCase{L".rs_case_tmp_0123456789abcdef0123456789abcdef.txt", false, false},
        ShapeCase{L".rs_case_tmp_0123456789abcdef0123456789abcdeg", false, false},
        ShapeCase{L".rs_case_tmp_0123456789abcdef0123456789abcde", false, false},
    };
    for (const ShapeCase& shape : kShapes)
    {
        if (Artifacts::HasPossibleArtifactName(shape.leaf) != shape.possible || Artifacts::HasPossibleOriginalArtifactName(shape.leaf) != shape.mayHoldOriginal)
            return reject(std::format(L"name shape '{}' classified possible={} original={}, expected {} {}",
                                      shape.leaf,
                                      Artifacts::HasPossibleArtifactName(shape.leaf),
                                      Artifacts::HasPossibleOriginalArtifactName(shape.leaf),
                                      shape.possible,
                                      shape.mayHoldOriginal));
    }

    const FileSystemPathIdentity slashPaths{
        .pathTextStableIdentity = true,
        .componentComparison    = FileSystemPathComponentComparison::OrdinalCaseSensitive,
        .preferredSeparator     = L'/',
        .acceptedSeparators     = L"/",
        .casePreserving         = true,
        .caseOnlyRename         = FileSystemPathCaseOnlyRename::Supported,
    };
    const FileSystemPathIdentity win32Paths{
        .pathTextStableIdentity = true,
        .componentComparison    = FileSystemPathComponentComparison::OrdinalIgnoreCase,
        .preferredSeparator     = L'\\',
        .acceptedSeparators     = L"\\/",
        .casePreserving         = true,
        .caseOnlyRename         = FileSystemPathCaseOnlyRename::Supported,
    };
    const Artifacts::Endpoint ftp{
        .pluginId = L"builtin/file-system-curl", .instanceId = L"host/default", .profileId = L"curl-ftp", .rootId = L"ftp://127.0.0.1"};
    const Artifacts::Endpoint local{
        .pluginId = L"builtin/file-system", .instanceId = L"host/default", .profileId = L"local-win32-smb", .rootId = L"\\\\nas\\share"};
    // A provider without object binding: BindObjectAuthority reports Unsupported (E_NOINTERFACE).
    const Artifacts::IdentityProbe noBinding{Artifacts::IdentityProbeOutcome::IdentityUnsupported, S_OK, E_NOINTERFACE};
    const Artifacts::IdentityProbe noFileId{Artifacts::IdentityProbeOutcome::IdentityUnsupported, S_OK, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)};
    const Artifacts::IdentityProbe sharingViolation{Artifacts::IdentityProbeOutcome::OpenFailed, HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), S_FALSE};
    const auto candidate = [](const Artifacts::Endpoint& endpoint,
                              const FileSystemPathIdentity& paths,
                              std::wstring path,
                              Artifacts::IdentityProbe probe,
                              Artifacts::DeleteHandoff handoff)
    {
        return Artifacts::Candidate{
            .endpoint      = endpoint,
            .pathIdentity  = paths,
            .path          = std::move(path),
            .probeState    = Artifacts::ProbeState::Indeterminate,
            .identityProbe = probe,
            .deleteHandoff = handoff,
        };
    };
    const std::wstring ftpLeftover = L"/pub/.redsalamander-upload-0123456789abcdef0123456789abcdef";

    // 1. A permanent Delete of an FTP leftover: no guard item at all (the provider's card asks).
    const std::array permanentFtp{candidate(ftp, slashPaths, ftpLeftover, noBinding, Artifacts::DeleteHandoff::PermanentCard)};
    Artifacts::TouchGuardRequest request{};
    if (Artifacts::BuildTouchGuardRequest(permanentFtp, request) != S_FALSE || ! request.items.empty() || request.containsPossible)
        return reject(std::format(L"a permanent Delete of an FTP leftover entered the guard ({} items)", request.items.size()));

    // 2. A Recycle-style handoff is warned about and accepted without a receipt.
    const std::array recycleFtp{candidate(ftp, slashPaths, ftpLeftover, noBinding, Artifacts::DeleteHandoff::RecycleWarn)};
    Artifacts::TouchGuardReceipt receipt{};
    if (Artifacts::BuildTouchGuardRequest(recycleFtp, request) != S_OK || request.items.size() != 1u || ! request.items.front().handedOff ||
        request.items.front().revalidatable || ! request.allItemsRevalidatable || Artifacts::AcceptTouchGuard(request, receipt) != S_OK ||
        ! receipt.items.empty())
        return reject(L"a handed-off Recycle item was not warned about without a receipt");

    // 3. Every other touch of the same leftover (a Move, an overwrite) stays refused.
    const std::array moveFtp{candidate(ftp, slashPaths, ftpLeftover, noBinding, Artifacts::DeleteHandoff::None)};
    if (Artifacts::BuildTouchGuardRequest(moveFtp, request) != S_OK || request.items.size() != 1u || request.items.front().handedOff ||
        request.allItemsRevalidatable || Artifacts::AcceptTouchGuard(request, receipt) != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || ! receipt.items.empty())
        return reject(L"a Move of an unidentifiable leftover was not refused");

    // 4. An item whose open failed is never handed off, whatever its Delete route.
    const std::array lockedLocal{candidate(
        local, win32Paths, L"\\\\nas\\share\\movie.mkv.rs_tmp_9D380E53676EAB4315AE57F418F5049D_0", sharingViolation, Artifacts::DeleteHandoff::PermanentCard)};
    if (Artifacts::BuildTouchGuardRequest(lockedLocal, request) != S_OK || request.items.size() != 1u || request.items.front().handedOff ||
        Artifacts::AcceptTouchGuard(request, receipt) != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
        return reject(L"an item whose open failed was handed to a by-name route");

    // 5. Mixed: a bound leftover keeps its receipt and the worker revalidates only that item; an item
    //    handed off is filtered, a handed-off item that turned bindable is a mismatch.
    Artifacts::Candidate bound = candidate(local,
                                           win32Paths,
                                           L"\\\\nas\\share\\clip.mp4.rs_tmp_0123456789ABCDEF0123456789ABCDEF_1",
                                           Artifacts::IdentityProbe{Artifacts::IdentityProbeOutcome::Bound, S_OK, S_OK},
                                           Artifacts::DeleteHandoff::None);
    bound.probeState           = Artifacts::ProbeState::Present;
    bound.currentIdentity      = Artifacts::Identity{
        .objectId      = {std::byte{0x01}, std::byte{0x02}, std::byte{0x03}, std::byte{0x04}},
        .revisionId    = {std::byte{0x10}},
        .pathProfileId = L"local-win32-smb",
        .kind          = FILESYSTEM_BOUND_REGULAR_FILE,
    };
    const Artifacts::Candidate handedOff =
        candidate(local, win32Paths, L"\\\\nas\\share\\movie.mkv.rs_tmp_9D380E53676EAB4315AE57F418F5049D_0", noFileId, Artifacts::DeleteHandoff::RecycleWarn);
    const std::array mixed{bound, handedOff};
    if (Artifacts::BuildTouchGuardRequest(mixed, request) != S_OK || request.items.size() != 2u || ! request.allItemsRevalidatable ||
        Artifacts::AcceptTouchGuard(request, receipt) != S_OK || receipt.items.size() != 1u || Artifacts::RevalidateTouchGuard(receipt, mixed) != S_OK)
        return reject(L"a mixed request did not keep the bound item's receipt and revalidate it alone");
    Artifacts::Candidate replaced          = bound;
    replaced.currentIdentity.revisionId[0] = std::byte{0x11};
    const std::array replacedMix{replaced, handedOff};
    if (Artifacts::RevalidateTouchGuard(receipt, replacedMix) != HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH))
        return reject(L"a replaced receipt item revalidated");
    Artifacts::Candidate nowBindable     = handedOff;
    nowBindable.identityProbe            = Artifacts::IdentityProbe{Artifacts::IdentityProbeOutcome::Bound, S_OK, S_OK};
    nowBindable.probeState               = Artifacts::ProbeState::Present;
    nowBindable.currentIdentity          = bound.currentIdentity;
    nowBindable.currentIdentity.objectId = {std::byte{0x09}};
    const std::array turnedBindable{bound, nowBindable};
    if (Artifacts::RevalidateTouchGuard(receipt, turnedBindable) != HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH))
        return reject(L"an item handed off at admission that became bindable was not a mismatch");

    // 6. A forged request cannot smuggle an identity-less item into a receipt.
    Artifacts::TouchGuardRequest forged{};
    forged.containsPossible      = true;
    forged.allItemsRevalidatable = true;
    forged.items.push_back(Artifacts::TouchGuardItem{.classification = Artifacts::Classification::Possible,
                                                     .endpoint       = ftp,
                                                     .pathIdentity   = slashPaths,
                                                     .path           = ftpLeftover,
                                                     .identityProbe  = noBinding,
                                                     .revalidatable  = true,
                                                     .handedOff      = true});
    if (Artifacts::AcceptTouchGuard(forged, receipt) != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
        return reject(L"a handed-off item claiming an identity was accepted");
    forged.items.front().revalidatable = false;
    forged.items.front().identityProbe = sharingViolation;
    if (Artifacts::AcceptTouchGuard(forged, receipt) != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
        return reject(L"a handed-off item whose open failed was accepted");

    Debug::Perf::Emit(
        L"FileOps.SelfTest.IdentityLessGuardPartition", L"shapes;permanent=card;recycle=warn;move=refused;openFailed=refused", 0u, kShapes.size(), 6u, S_OK);
    return true;
}

// I35 R7.5: the leftover listing's subfolder walk, run on the test thread: it lists the Possible names
// below the folder, never enters a link or a leftover folder, reports its folder, item and depth bounds
// each by its own flag, and ends at once when stopped.
[[nodiscard]] bool RunFileOpsLeftoverCleanupWalkBoundsChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Leftover cleanup walk bounds: {}", reason));
        return false;
    };
    if (! state.fsLocal)
        return reject(L"Local provider unavailable");
    const std::filesystem::path root    = state.tempRoot / L"leftover-walk";
    const std::filesystem::path outside = state.tempRoot / L"leftover-walk-outside";
    if (! RecreateEmptyDirectory(root / L"sub" / L"deeper") || ! RecreateEmptyDirectory(outside))
        return reject(L"case root unavailable");
    const auto removeRoots                  = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SelfTest::RemoveAll(root));
        static_cast<void>(SelfTest::RemoveAll(outside));
    });
    const std::wstring tail                 = L".rs_tmp_00112233445566778899AABBCCDDEEFF_0";
    const std::filesystem::path top         = root / (L"a.mkv" + tail);
    const std::filesystem::path middle      = root / L"sub" / (L"b.bin" + tail);
    const std::filesystem::path deep        = root / L"sub" / L"deeper" / (L"c.bin" + tail);
    const std::filesystem::path leftoverDir = root / L"folder.rs_tmp_00112233445566778899AABBCCDDEEFF_1";
    const std::filesystem::path inner       = leftoverDir / (L"inner.bin" + tail);
    const std::filesystem::path beyondLink  = outside / (L"e.bin" + tail);
    if (! WriteFilledTestFile(top, 512u, 0x71u) || ! WriteFilledTestFile(middle, 512u, 0x72u) || ! WriteFilledTestFile(deep, 512u, 0x73u) ||
        ! RecreateEmptyDirectory(leftoverDir) || ! WriteFilledTestFile(inner, 512u, 0x74u) || ! WriteFilledTestFile(beyondLink, 512u, 0x75u) ||
        ! TryCreateJunction(root / L"link", outside))
        return reject(L"fixtures (including the junction) unavailable");

    const auto scan = [&](size_t maxFolders, size_t maxRows, size_t maxDepth, bool cancel, FolderWindow::LeftoverScanDebugResult& result) noexcept
    { return FolderWindow::DebugScanLeftoversForSelfTest(state.fsLocal.get(), root, maxFolders, maxRows, maxDepth, cancel, result); };
    const auto listed = [](const FolderWindow::LeftoverScanDebugResult& result, const std::filesystem::path& path) noexcept
    { return std::ranges::find(result.rows, path) != result.rows.end(); };
    const auto describe = [](const FolderWindow::LeftoverScanDebugResult& result)
    {
        std::wstring text = std::format(
            L"folderLimit={} rowLimit={} depthLimit={} canceled={} rows=[", result.folderLimit, result.rowLimit, result.depthLimit, result.canceled);
        for (const std::filesystem::path& row : result.rows)
            text += row.filename().native() + L";";
        return text + L"]";
    };

    // 1. Unbounded: the three leftovers and the leftover folder (listed, not entered); nothing beyond the link.
    FolderWindow::LeftoverScanDebugResult full{};
    if (FAILED(scan(4096u, 4096u, 32u, false, full)) || full.rows.size() != 4u || ! listed(full, top) || ! listed(full, middle) || ! listed(full, deep) ||
        ! listed(full, leftoverDir) || listed(full, inner) ||
        std::ranges::any_of(full.rows, [&](const std::filesystem::path& row) noexcept { return row.filename() == beyondLink.filename(); }) ||
        full.folderLimit || full.rowLimit || full.depthLimit || full.canceled)
        return reject(std::format(L"the unbounded walk listed the wrong rows: {}", describe(full)));

    // 2. Depth: two levels reach `sub` but not `deeper`, and say so.
    FolderWindow::LeftoverScanDebugResult shallow{};
    if (FAILED(scan(4096u, 4096u, 2u, false, shallow)) || shallow.rows.size() != 3u || listed(shallow, deep) || ! listed(shallow, middle) ||
        ! shallow.depthLimit || shallow.folderLimit || shallow.rowLimit)
        return reject(std::format(L"the depth bound was not reported: {}", describe(shallow)));

    // 3. Folders: one folder read, the rest reported as not searched.
    FolderWindow::LeftoverScanDebugResult oneFolder{};
    if (FAILED(scan(1u, 4096u, 32u, false, oneFolder)) || oneFolder.rows.size() != 2u || ! listed(oneFolder, top) || ! listed(oneFolder, leftoverDir) ||
        ! oneFolder.folderLimit || oneFolder.rowLimit || oneFolder.depthLimit)
        return reject(std::format(L"the folder bound was not reported: {}", describe(oneFolder)));

    // 4. Items: the first one listed, the item bound reported (not the folder bound).
    FolderWindow::LeftoverScanDebugResult oneRow{};
    if (FAILED(scan(4096u, 1u, 32u, false, oneRow)) || oneRow.rows.size() != 1u || ! oneRow.rowLimit || oneRow.folderLimit)
        return reject(std::format(L"the item bound was not reported: {}", describe(oneRow)));

    // 5. Stopped before the first folder: nothing listed, canceled.
    FolderWindow::LeftoverScanDebugResult stopped{};
    if (FAILED(scan(4096u, 4096u, 32u, true, stopped)) || ! stopped.rows.empty() || ! stopped.canceled)
        return reject(std::format(L"a stopped walk did not end at once: {}", describe(stopped)));

    Debug::Perf::Emit(L"FileOps.SelfTest.LeftoverCleanupWalkBounds", L"full;depth;folders;items;stop", 0u, 5u, 0u, S_OK);
    return true;
}

// R7.4 (plan 7.3): a bridge Copy to the loopback alias whose owned-stage removal takes one second (a
// cancelable synchronous read stands in for a slow SMB disposition) under a requested cancel. The
// cancel watch must leave that cleanup alone for its budget: the stage is removed, its disposition is
// Removed and no bridge.stage.abort* diagnostic is recorded.
[[nodiscard]] bool RunFileOpsOwnedStageCleanupBudgetChecks(SelfTestState& state, std::wstring& skipped) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Owned-stage cleanup budget: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");
    if (uncStatus == LoopbackShareStatus::Unreachable)
    {
        skipped = L"The loopback administrative share is not reachable.";
        AppendLog(L"Owned-stage cleanup budget: the loopback administrative share is not reachable; environment-gated case skipped.");
        return true;
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"owned-stage-cleanup-budget";
    if (! RecreateEmptyDirectory(root / L"src") || ! RecreateEmptyDirectory(root / L"dst"))
        return reject(L"case root unavailable");
    const auto removeRoot                   = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path source      = root / L"src" / L"payload.bin";
    const std::filesystem::path destination = LoopbackShareAlias(root / L"dst");
    if (! WriteFilledTestFile(source, 8u * 1024u * 1024u, 0x3Cu))
        return reject(L"source fixture unavailable");
    IdentityLessCaseScope scope(state, operations);
    // Verify after copy routes a Local Copy through the host bridge and its owned stage.
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = true;
    static_cast<void>(SetEnvironmentVariableW(kIdentityLessStageDelayFire.data(), nullptr));
    if (SetEnvironmentVariableW(kIdentityLessStageDelayRoot.data(), destination.c_str()) == FALSE)
        return reject(L"the stage-cleanup delay could not be armed");

    const auto copyTask = StartFileOperationAndGetId(operations,
                                                     FILESYSTEM_COPY,
                                                     FolderWindow::Pane::Left,
                                                     FolderWindow::Pane::Right,
                                                     state.fsLocal,
                                                     {source},
                                                     destination,
                                                     FILESYSTEM_FLAG_NONE,
                                                     false,
                                                     512u * 1024u,
                                                     FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                     false,
                                                     state.fsLocal);
    if (! copyTask.has_value())
        return reject(L"the bridge Copy was not admitted");
    scope.Track(copyTask.value());
    const auto staged            = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return HasStageLeafIn(root / L"dst") || state.completedTasks.contains(copyTask.value());
    }, {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = L"bridge owned stage"});
    IdentityLessTask* const task = operations->FindTask(copyTask.value());
    // Only the host bridge names its owned stage `.rs_tmp_`: seeing one proves the route under test.
    if (! staged.conditionMet || state.completedTasks.contains(copyTask.value()) || task == nullptr)
        return reject(L"the bridge Copy never showed its owned stage while running");
    task->RequestCancel();
    const CompletedTaskInfo* ended = scope.WaitForEnd(copyTask.value(), L"canceled bridge Copy");
    const auto summary             = FindIdentityLessSummary(operations, copyTask.value());
    const std::wstring fired       = GetEnvVarTrimmed(kIdentityLessStageDelayFire);
    const bool abortRecorded       = summary.has_value() && std::ranges::any_of(summary->issueDiagnostics, [](const IdentityLessDiagnostic& issue) noexcept {
        return issue.category.starts_with(L"bridge.stage.abort");
    });
    const auto* item               = ended != nullptr && ended->sourceItemResults.size() == 1u && ended->sourceItemResults.front().has_value()
                                         ? &ended->sourceItemResults.front().value()
                                         : nullptr;
    if (ended == nullptr || ended->hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || item == nullptr ||
        item->ownedStageDisposition != FileOperations::OwnedStageDisposition::Removed || fired != L"1" || abortRecorded || HasStageLeafIn(root / L"dst") ||
        ! summary.has_value() || summary->retainedOwnedStageCount != 0u || summary->unknownOwnedStageCount != 0u)
        return reject(std::format(L"the delayed stage cleanup did not survive the cancel watch: hr=0x{:08X} stage={} fired='{}' abortRecorded={} "
                                  L"leftover={} issues=[{}]",
                                  ended != nullptr ? static_cast<unsigned long>(ended->hr) : 0ul,
                                  item != nullptr ? static_cast<unsigned int>(item->ownedStageDisposition) : 255u,
                                  fired,
                                  abortRecorded,
                                  HasStageLeafIn(root / L"dst"),
                                  summary.has_value() ? DescribeIdentityLessIssues(summary.value()) : std::wstring(L"<no summary>")));
    Debug::Perf::Emit(L"FileOps.SelfTest.OwnedStageCleanupBudget", L"delay=1000ms;budget=2000ms;stage=Removed", 0u, 1u, 0u, S_OK);
    return true;
}

// Plan 7.3 `Delete_NasFileIdInfoUnsupportedByName` on the registered share (path.smb.primary): the
// location exposes no FileIdInfo, so BindObject is Unsupported and the corroboration facts carry the
// 64-bit file index the share reports (with its volume serial, corroboration only; the plan expected
// a zero serial, the share reports one through GetFileInformationByHandle, so the serial is logged and
// not asserted); the file is deleted by name behind its consent. Only a unique per-run child is
// written; R7.7: the share's @Recycle may keep what is deleted, so only the visible tree is asserted.
[[nodiscard]] bool RunDeleteNasFileIdInfoUnsupportedByNameChecks(SelfTestState& state, std::wstring& skipped) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"NAS delete by name: {}", reason));
        return false;
    };
    const std::wstring smbRootText = GetEnvVarTrimmed(L"REDSALAMANDER_SELFTEST_SMB_ROOT");
    if (smbRootText.empty())
    {
        skipped = L"path.smb.primary is not configured (REDSALAMANDER_SELFTEST_SMB_ROOT).";
        AppendLog(L"NAS delete by name: skipped (path.smb.primary is not configured).");
        return true;
    }
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path child =
        std::filesystem::path(smbRootText) / std::format(L"identity-less-{}-{:X}", GetCurrentProcessId(), static_cast<unsigned long long>(GetTickCount64()));
    if (! RecreateEmptyDirectory(child))
        return reject(std::format(L"per-run child '{}' unavailable", child.native()));
    const auto removeChild = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(child)); });
    IdentityLessCaseScope scope(state, operations);
    const std::filesystem::path victim = child / L"nas-by-name.bin";
    if (! WriteFilledTestFile(victim, 64u * 1024u, 0x6Du))
        return reject(L"NAS fixture unavailable");

    wil::com_ptr<IFileSystemObjectBinding> binding;
    wil::com_ptr<IFileSystemIdentityDelete> identityDelete;
    if (FAILED(state.fsLocal->QueryInterface(__uuidof(IFileSystemObjectBinding), binding.put_void())) || ! binding ||
        FAILED(state.fsLocal->QueryInterface(__uuidof(IFileSystemIdentityDelete), identityDelete.put_void())) || ! identityDelete)
        return reject(L"the Local provider lacks object binding or its delete by name");
    wil::com_ptr<IFileSystemBoundObject> boundObject;
    const HRESULT bindHr =
        binding->BindObject(victim.c_str(), static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA), boundObject.put());
    FileSystemDeleteIdentity corroboration{};
    corroboration.sizeBytes = sizeof(corroboration);
    const HRESULT resolveHr = identityDelete->ResolveDeleteIdentity(victim.c_str(), nullptr, &corroboration);
    const std::wstring_view facts(corroboration.identity, ::wcsnlen(corroboration.identity, std::size(corroboration.identity)));
    if (bindHr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || boundObject || resolveHr != S_OK || facts.find(L"|i=") == std::wstring_view::npos ||
        facts.find(L"|v=") == std::wstring_view::npos)
        return reject(std::format(L"the share did not behave as the probed NAS: bind=0x{:08X} resolve=0x{:08X} facts='{}'",
                                  static_cast<unsigned long>(bindHr),
                                  static_cast<unsigned long>(resolveHr),
                                  facts));

    const auto deleteTask = scope.StartDelete({victim}, false);
    if (! deleteTask.has_value())
        return reject(L"the NAS Delete was not admitted");
    const auto card = scope.WaitForQuestion(deleteTask.value(), L"NAS by-name card");
    if (! card.has_value() || ! IsByNameCard(card.value()) || ! scope.Answer(deleteTask.value(), card.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the NAS Delete did not ask its by-name consent: {}", DescribeIdentityLessPrompt(card)));
    const CompletedTaskInfo* ended = scope.WaitForEnd(deleteTask.value(), L"NAS Delete by name");
    const auto summary             = FindIdentityLessSummary(operations, deleteTask.value());
    if (ended == nullptr || ended->hr != S_OK || std::filesystem::exists(victim) || ! summary.has_value() ||
        CountIdentityLessIssues(summary.value(), L"delete.permanent.byName", IdentityLessSeverity::Info, S_OK) != 1u)
        return reject(std::format(L"the NAS file was not deleted by name: hr=0x{:08X} visible={} issues=[{}]",
                                  ended != nullptr ? static_cast<unsigned long>(ended->hr) : 0ul,
                                  std::filesystem::exists(victim),
                                  summary.has_value() ? DescribeIdentityLessIssues(summary.value()) : std::wstring(L"<no summary>")));
    AppendLog(std::format(L"NAS delete by name: '{}' deleted by name; corroboration '{}'.", victim.native(), facts));

    // Round 2 (pinned tree): a small folder on the share pins its whole listing (digest and counts: two files,
    // one folder, no link); its entries are then compared, one by one, on the share's own handles and the
    // folder is deleted by name.
    const std::filesystem::path tree = child / L"nas-by-name-tree";
    if (! RecreateEmptyDirectory(tree / L"sub") || ! WriteFilledTestFile(tree / L"a.bin", 8u * 1024u, 0x6Eu) ||
        ! WriteFilledTestFile(tree / L"sub" / L"b.bin", 4u * 1024u, 0x6Fu))
        return reject(L"NAS folder fixture unavailable");
    FileSystemDeleteIdentity treeFacts{};
    treeFacts.sizeBytes         = sizeof(treeFacts);
    const HRESULT treeResolveHr = identityDelete->ResolveDeleteIdentity(tree.c_str(), nullptr, &treeFacts);
    const std::wstring treeText(treeFacts.identity, ::wcsnlen(treeFacts.identity, std::size(treeFacts.identity)));
    if (treeResolveHr != S_OK || treeFacts.isDirectory == FALSE || treeText.find(L"|m=") == std::wstring::npos ||
        treeText.find(L"|n=2.1.0") == std::wstring::npos)
        return reject(std::format(L"the NAS folder did not pin its listing: resolve=0x{:08X} facts='{}'", static_cast<unsigned long>(treeResolveHr), treeText));
    const auto treeTask = scope.StartDelete({tree}, false);
    if (! treeTask.has_value())
        return reject(L"the NAS folder Delete was not admitted");
    const auto treeCard = scope.WaitForQuestion(treeTask.value(), L"NAS folder by-name card");
    if (! treeCard.has_value() || ! IsByNameCard(treeCard.value()) || ! scope.Answer(treeTask.value(), treeCard.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the NAS folder Delete did not ask its by-name consent: {}", DescribeIdentityLessPrompt(treeCard)));
    const CompletedTaskInfo* treeEnded = scope.CancelQuestionsUntilEnd(treeTask.value(), L"NAS folder Delete by name");
    const auto treeSummary             = FindIdentityLessSummary(operations, treeTask.value());
    if (treeEnded == nullptr || treeEnded->hr != S_OK || std::filesystem::exists(tree))
        return reject(std::format(L"the NAS folder was not deleted by name: hr=0x{:08X} visible={} issues=[{}]",
                                  treeEnded != nullptr ? static_cast<unsigned long>(treeEnded->hr) : 0ul,
                                  std::filesystem::exists(tree),
                                  treeSummary.has_value() ? DescribeIdentityLessIssues(treeSummary.value()) : std::wstring(L"<no summary>")));
    AppendLog(std::format(L"NAS delete by name: folder '{}' deleted by name; corroboration '{}'.", tree.native(), treeText));
    Debug::Perf::Emit(L"FileOps.SelfTest.NasDeleteByName", L"bind=unsupported;index-corroborated;deleted=by-name;tree=pinned", 0u, 2u, 0u, S_OK);
    return true;
}

// I35 S4 (red-team HIGH, rounds 1 and 2): with queue mode on and a task running, a Delete by name of a
// shared folder is answered and queued; a Move of a local folder into that shared folder then starts in
// parallel. The Move sees the queued delete root by path: its question names the delete, says that the
// delete could remove what it writes there, defaults to Queue after these tasks and offers no Run at the
// same time; queued after it, the Move runs only once the delete ended, and the moved data survives.
[[nodiscard]] bool RunDeleteIdentityLessQueuedRootOverlapChecks(SelfTestState& state, std::wstring& skipped) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Identity-less queued root overlap: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");
    if (uncStatus == LoopbackShareStatus::Unreachable)
    {
        // The running task must hold another identity domain (the drive letter) than the share, so that
        // only its queue place, not the conservative interlock, keeps the delete waiting.
        skipped = L"The loopback administrative share is not reachable.";
        AppendLog(L"Identity-less queued root overlap: the loopback administrative share is not reachable; environment-gated case skipped.");
        return true;
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"identity-less-queued-overlap";
    if (! RecreateEmptyDirectory(root / L"share" / L"old") || ! RecreateEmptyDirectory(root / L"local" / L"docs"))
        return reject(L"case root unavailable");
    const auto removeRoot                    = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location     = LoopbackShareAlias(root / L"share");
    const std::filesystem::path shared       = root / L"share" / L"old"; // the shared folder, through the drive letter
    const std::filesystem::path docs         = root / L"local" / L"docs";
    const std::filesystem::path payload      = docs / L"doc.bin";
    const std::filesystem::path movedPayload = shared / L"docs" / L"doc.bin";
    const std::filesystem::path keep         = shared / L"keep.bin";
    const std::filesystem::path hold         = root / L"local" / L"hold.bin";
    if (! WriteFilledTestFile(keep, 2048u, 0x41u) || ! WriteFilledTestFile(payload, 8192u, 0x42u) || ! WriteFilledTestFile(hold, 1024u, 0x43u))
        return reject(L"fixtures unavailable");
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location))
        return reject(L"the identity-less injection could not be configured");
    operations->ApplyQueueMode(true);

    QueuedByNameDelete fixture{};
    if (const std::wstring failure = StartQueuedByNameDelete(scope, hold, location / L"old", fixture); ! failure.empty())
        return reject(failure);

    // B: a Move of a local folder into the shared folder, run in parallel (not queued).
    const auto move = StartFileOperationAndGetId(operations,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {docs},
                                                 location / L"old",
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
    if (! move.has_value())
        return reject(L"the Move into the shared folder was not admitted");
    scope.Track(move.value());
    const auto advisory = scope.WaitForQuestion(move.value(), L"queued overlap Move question");
    const bool advised =
        advisory.has_value() && IsByNameOverlapQuestion(advisory.value(), fixture.deletion, FileOperations::SameHostOverlapProblem::DeletedByNameWritten);
    bool queuedBehindDelete = false;
    if (advisory.has_value() && advisory->bucket == IdentityLessBucket::SameHostOverlap)
    {
        // Queue after these tasks, the default, whatever else the question offered.
        if (! scope.Answer(move.value(), advisory.value(), IdentityLessAction::Proceed))
            return reject(L"the Move's overlap question could not be answered");
        queuedBehindDelete               = scope.WaitUntilWaitingToEnter(move.value(), L"queued overlap Move behind the delete");
        IdentityLessTask* const moveTask = operations->FindTask(move.value());
        queuedBehindDelete               = queuedBehindDelete && moveTask != nullptr && moveTask->IsOverlapQueued();
    }
    else if (! advisory.has_value())
    {
        // The defect of round 1: the Move ran without a question; let it finish before T ends.
        static_cast<void>(scope.WaitForEnd(move.value(), L"queued overlap unadvised Move"));
    }
    const bool moveEndedBeforeDelete = state.completedTasks.contains(move.value()) && ! state.completedTasks.contains(fixture.deletion);

    // T ends; D runs, and B after it when B queued behind it.
    std::vector<IdentityLessPrompt> deleteQuestions;
    const CompletedTaskInfo* deleteEnd = FinishQueuedByNameDelete(scope, fixture, &deleteQuestions);
    const CompletedTaskInfo* moveEnd   = scope.CancelQuestionsUntilEnd(move.value(), L"queued overlap Move");

    const bool sourceKept  = FileSizeEquals(payload, 8192u);
    const bool moved       = FileSizeEquals(movedPayload, 8192u);
    const HRESULT deleteHr = deleteEnd != nullptr ? deleteEnd->hr : E_PENDING;
    const HRESULT moveHr   = moveEnd != nullptr ? moveEnd->hr : E_PENDING;
    const std::wstring observed =
        std::format(L"advised={} queuedBehindDelete={} moveEndedBeforeDelete={} deleteHr=0x{:08X} moveHr=0x{:08X} sourceKept={} moved={} keep={} "
                    L"deleteQuestions={} question=[{}]",
                    advised,
                    queuedBehindDelete,
                    moveEndedBeforeDelete,
                    static_cast<unsigned long>(deleteHr),
                    static_cast<unsigned long>(moveHr),
                    sourceKept,
                    moved,
                    std::filesystem::exists(keep),
                    deleteQuestions.size(),
                    DescribeOverlapQuestion(advisory));
    AppendLog(std::format(L"Identity-less queued root overlap: {}", observed));
    // B waited for the delete: the delete removed only what it was consented for, the moved data survives once.
    if (! advised || ! queuedBehindDelete || moveEndedBeforeDelete || deleteEnd == nullptr || moveEnd == nullptr || deleteHr != S_OK ||
        std::filesystem::exists(keep) || sourceKept == moved)
        return reject(std::format(L"a task admitted after a queued delete by name did not respect its root: {}", observed));
    Debug::Perf::Emit(L"FileOps.SelfTest.IdentityLessQueuedRootOverlap", L"queued=after-delete;run-concurrently=not-offered", 0u, 1u, 0u, S_OK);
    return true;
}

// I35 S4 round 2 (pinned tree): the shared folder reached through another spelling (here the drive letter
// of the loopback share's folder) cannot be compared by path, so a Move into a subfolder there runs while
// the delete by name waits in the queue. The delete deletes only the tree it pinned: the moved entries change
// its listing, so it refuses (or stops partway), and the moved data survives.
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeOtherSpellingChecks(SelfTestState& state, std::wstring& skipped) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, other spelling: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");
    if (uncStatus == LoopbackShareStatus::Unreachable)
    {
        skipped = L"The loopback administrative share is not reachable.";
        AppendLog(L"Pinned tree, other spelling: the loopback administrative share is not reachable; environment-gated case skipped.");
        return true;
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"identity-less-other-spelling";
    if (! RecreateEmptyDirectory(root / L"share" / L"old" / L"Sub") || ! RecreateEmptyDirectory(root / L"local" / L"docs"))
        return reject(L"case root unavailable");
    const auto removeRoot                    = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location     = LoopbackShareAlias(root / L"share");
    const std::filesystem::path shared       = root / L"share" / L"old";
    const std::filesystem::path keep         = shared / L"keep.bin";
    const std::filesystem::path payload      = root / L"local" / L"docs" / L"doc.bin";
    const std::filesystem::path movedPayload = shared / L"Sub" / L"docs" / L"doc.bin";
    const std::filesystem::path hold         = root / L"local" / L"hold.bin";
    if (! WriteFilledTestFile(keep, 2048u, 0x51u) || ! WriteFilledTestFile(shared / L"Sub" / L"sub.bin", 1024u, 0x52u) ||
        ! WriteFilledTestFile(payload, 8192u, 0x53u) || ! WriteFilledTestFile(hold, 1024u, 0x54u))
        return reject(L"fixtures unavailable");
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location))
        return reject(L"the identity-less injection could not be configured");
    operations->ApplyQueueMode(true);

    QueuedByNameDelete fixture{};
    if (const std::wstring failure = StartQueuedByNameDelete(scope, hold, location / L"old", fixture); ! failure.empty())
        return reject(failure);

    // B: a Move into the drive-letter spelling of the delete's subfolder; nothing can relate the two paths.
    const auto move = StartFileOperationAndGetId(operations,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {payload.parent_path()},
                                                 shared / L"Sub",
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
    if (! move.has_value())
        return reject(L"the Move into the other spelling was not admitted");
    scope.Track(move.value());
    std::vector<IdentityLessPrompt> moveQuestions;
    const CompletedTaskInfo* moveEnd = scope.CancelQuestionsUntilEnd(move.value(), L"other-spelling Move", &moveQuestions);
    const bool movedFirst            = moveEnd != nullptr && moveEnd->hr == S_OK && FileSizeEquals(movedPayload, 8192u);

    std::vector<IdentityLessPrompt> deleteQuestions;
    const CompletedTaskInfo* deleteEnd = FinishQueuedByNameDelete(scope, fixture, &deleteQuestions);
    const HRESULT deleteHr             = deleteEnd != nullptr ? deleteEnd->hr : E_PENDING;
    const bool movedIntact             = FileSizeEquals(movedPayload, 8192u);
    const std::wstring observed        = std::format(L"movedFirst={} moveQuestions={} deleteHr=0x{:08X} deleteQuestions={} movedIntact={} keep={}",
                                                     movedFirst,
                                                     moveQuestions.size(),
                                                     static_cast<unsigned long>(deleteHr),
                                                     deleteQuestions.size(),
                                                     movedIntact,
                                                     std::filesystem::exists(keep));
    AppendLog(std::format(L"Pinned tree, other spelling: {}", observed));
    if (! movedFirst || ! moveQuestions.empty() || deleteEnd == nullptr || SUCCEEDED(deleteHr) || ! movedIntact)
        return reject(std::format(L"the delete by name removed data moved in through another spelling: {}", observed));
    Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeOtherSpelling", L"moved-first;delete-refused", 0u, 1u, 0u, S_OK);
    return true;
}

// I35 S4 round 2 (pinned tree): Pack writes its archive directly, outside the File Operations admission,
// the overlap question and the interlock. An archive packed into a subfolder of a folder whose delete by
// name waits in the queue changes the pinned listing and survives that delete.
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeArchiveWriterChecks(SelfTestState& state, std::wstring& skipped) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, archive writer: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");
    if (uncStatus == LoopbackShareStatus::Unreachable)
    {
        skipped = L"The loopback administrative share is not reachable.";
        AppendLog(L"Pinned tree, archive writer: the loopback administrative share is not reachable; environment-gated case skipped.");
        return true;
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"identity-less-archive-writer";
    if (! RecreateEmptyDirectory(root / L"share" / L"old" / L"Sub") || ! RecreateEmptyDirectory(root / L"local" / L"pack-src"))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location = LoopbackShareAlias(root / L"share");
    const std::filesystem::path shared   = root / L"share" / L"old";
    const std::filesystem::path archive  = shared / L"Sub" / L"packed.zip";
    const std::filesystem::path packSrc  = root / L"local" / L"pack-src";
    const std::filesystem::path hold     = root / L"local" / L"hold.bin";
    if (! WriteFilledTestFile(shared / L"keep.bin", 2048u, 0x61u) || ! WriteFilledTestFile(shared / L"Sub" / L"sub.bin", 1024u, 0x62u) ||
        ! WriteFilledTestFile(packSrc / L"packme.txt", 333u, 0x6Du) || ! WriteFilledTestFile(hold, 1024u, 0x63u))
        return reject(L"fixtures unavailable");
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location))
        return reject(L"the identity-less injection could not be configured");
    operations->ApplyQueueMode(true);

    QueuedByNameDelete fixture{};
    if (const std::wstring failure = StartQueuedByNameDelete(scope, hold, location / L"old", fixture); ! failure.empty())
        return reject(failure);

    // Pack, from the left pane, writes its archive into the delete's subfolder (drive-letter spelling).
    const std::optional<std::filesystem::path> leftBefore = folder->GetCurrentPath(FolderWindow::Pane::Left);
    const std::wstring leftPluginBefore(folder->GetFileSystemPluginId(FolderWindow::Pane::Left));
    const auto restorePane = wil::scope_exit([&]() noexcept
    {
        folder->DebugClearArchiveCommandOptionsForTest();
        static_cast<void>(folder->SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
            folder->SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
    });
    if (FAILED(folder->SetFileSystemPluginForPane(FolderWindow::Pane::Left, kPluginIdLocal)))
        return reject(L"the left pane could not use the Local provider");
    folder->SetFolderPath(FolderWindow::Pane::Left, packSrc);
    const auto selected = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        folder->SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"packme.txt"; });
        return folder->DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u;
    },
        {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = L"archive writer selection"});
    if (! selected.conditionMet)
        return reject(L"the file to pack could not be selected in the left pane");
    FolderWindow::ArchiveCommandDebugOptions packOptions{};
    packOptions.archivePath = archive;
    packOptions.overwrite   = true;
    folder->DebugSetNextArchiveCommandOptionsForTest(packOptions);
    folder->CommandPack(FolderWindow::Pane::Left);
    const std::optional<FolderWindow::ArchiveCommandDebugResult> packResult = folder->DebugGetLastArchiveCommandResultForTest();
    if (! packResult.has_value() || FAILED(packResult->hr) || ! std::filesystem::exists(archive))
        return reject(std::format(L"Pack did not write its archive into the folder: hr=0x{:08X} exists={}",
                                  packResult.has_value() ? static_cast<unsigned long>(packResult->hr) : 0ul,
                                  std::filesystem::exists(archive)));
    std::error_code sizeError;
    const uintmax_t archiveBytes = std::filesystem::file_size(archive, sizeError);

    std::vector<IdentityLessPrompt> deleteQuestions;
    const CompletedTaskInfo* deleteEnd = FinishQueuedByNameDelete(scope, fixture, &deleteQuestions);
    const HRESULT deleteHr             = deleteEnd != nullptr ? deleteEnd->hr : E_PENDING;
    std::error_code afterError;
    const bool archiveIntact = ! sizeError && std::filesystem::exists(archive) && std::filesystem::file_size(archive, afterError) == archiveBytes;
    const std::wstring observed =
        std::format(L"deleteHr=0x{:08X} deleteQuestions={} archiveIntact={}", static_cast<unsigned long>(deleteHr), deleteQuestions.size(), archiveIntact);
    AppendLog(std::format(L"Pinned tree, archive writer: {}", observed));
    if (deleteEnd == nullptr || SUCCEEDED(deleteHr) || ! archiveIntact)
        return reject(std::format(L"the delete by name removed an archive Pack wrote into its tree: {}", observed));
    Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeArchiveWriter", L"pack-into-subfolder;delete-refused", 0u, 1u, 0u, S_OK);
    return true;
}

// I35 S4 round 2: while the by-name card waits, the delete is already visible to later tasks (its
// registration is published before the card): a Copy into its folder is asked, with the by-name problem and
// no Run at the same time. Data added deep in the tree while the card waits changes the pinned listing, so
// the answer's re-check refuses with ERROR_REVISION_MISMATCH and nothing is deleted.
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeCardPendingChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, card pending: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-card-pending";
    if (! RecreateEmptyDirectory(root / L"share" / L"old" / L"Sub") || ! RecreateEmptyDirectory(root / L"local"))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path share    = root / L"share";
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(share) : share;
    const std::filesystem::path shared   = share / L"old";
    const std::filesystem::path incoming = root / L"local" / L"incoming.bin";
    const std::filesystem::path deep     = shared / L"Sub" / L"deep.bin";
    if (! WriteFilledTestFile(shared / L"a.bin", 2048u, 0x71u) || ! WriteFilledTestFile(shared / L"Sub" / L"b.bin", 1024u, 0x72u) ||
        ! WriteFilledTestFile(incoming, 4096u, 0x73u))
        return reject(L"fixtures unavailable");
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location))
        return reject(L"the identity-less injection could not be configured");

    const auto deletion = scope.StartDelete({location / L"old"}, false);
    if (! deletion.has_value())
        return reject(L"the delete by name was not admitted");
    const auto card = scope.WaitForQuestion(deletion.value(), L"card-pending card");
    if (! card.has_value() || ! IsByNameCard(card.value()))
        return reject(std::format(L"the delete by name did not ask its card: {}", DescribeIdentityLessPrompt(card)));

    // 1. A Copy into the folder while the card waits is asked about the delete; Don't start ends it.
    const auto copy = StartFileOperationAndGetId(operations,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {incoming},
                                                 location / L"old" / L"Sub",
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
    if (! copy.has_value())
        return reject(L"the Copy into the folder was not admitted");
    scope.Track(copy.value());
    const auto copyQuestion = scope.WaitForQuestion(copy.value(), L"card-pending Copy question");
    const bool copyAdvised  = copyQuestion.has_value() &&
                              IsByNameOverlapQuestion(copyQuestion.value(), deletion.value(), FileOperations::SameHostOverlapProblem::DeletedByNameWritten);
    if (copyQuestion.has_value() && copyQuestion->bucket == IdentityLessBucket::SameHostOverlap)
        static_cast<void>(scope.Answer(copy.value(), copyQuestion.value(), IdentityLessAction::Cancel));
    const CompletedTaskInfo* copyEnd = scope.CancelQuestionsUntilEnd(copy.value(), L"card-pending Copy");
    const bool copiedIn              = std::filesystem::exists(shared / L"Sub" / incoming.filename());

    // 2. Data added deep in the tree while the card waits; then Delete by name.
    if (! WriteFilledTestFile(deep, 512u, 0x74u) || ! scope.Answer(deletion.value(), card.value(), IdentityLessAction::PermanentDelete))
        return reject(L"the deep change or the answer could not be made");
    const CompletedTaskInfo* deleteEnd = scope.CancelQuestionsUntilEnd(deletion.value(), L"card-pending delete");
    const auto summary                 = FindIdentityLessSummary(operations, deletion.value());
    const bool refused =
        deleteEnd != nullptr && deleteEnd->hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) && summary.has_value() &&
        CountIdentityLessIssues(
            summary.value(), L"delete.confirmation.identityChanged", IdentityLessSeverity::Error, HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH)) == 1u;
    const bool kept             = FileSizeEquals(shared / L"a.bin", 2048u) && FileSizeEquals(shared / L"Sub" / L"b.bin", 1024u) && FileSizeEquals(deep, 512u);
    const std::wstring observed = std::format(L"copyAdvised={} copyHr=0x{:08X} copiedIn={} deleteHr=0x{:08X} refused={} kept={} question=[{}]",
                                              copyAdvised,
                                              copyEnd != nullptr ? static_cast<unsigned long>(copyEnd->hr) : 0ul,
                                              copiedIn,
                                              deleteEnd != nullptr ? static_cast<unsigned long>(deleteEnd->hr) : 0ul,
                                              refused,
                                              kept,
                                              DescribeOverlapQuestion(copyQuestion));
    AppendLog(std::format(L"Pinned tree, card pending: {}", observed));
    if (! copyAdvised || copyEnd == nullptr || copyEnd->hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || copiedIn || ! refused || ! kept)
        return reject(std::format(L"the delete by name was not visible while its card waited, or deleted a tree that changed: {}", observed));
    Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeCardPending", L"later-task-asked;deep-change-refused", 0u, 2u, 0u, S_OK);
    return true;
}

// Arms the provider pause of a delete by name of exactly `providerRoot`: the next one stops after the check of its
// listing, before its first deletion, until RELEASE is set (or the scaled deadline passes).
[[nodiscard]] bool ArmIdentityLessByNamePause(const std::filesystem::path& providerRoot) noexcept
{
    static_cast<void>(SetEnvironmentVariableW(kIdentityLessByNamePauseEnteredEnv.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kIdentityLessByNamePauseReleaseEnv.data(), nullptr));
    const auto timeoutMs = std::chrono::duration_cast<std::chrono::milliseconds>(SelfTest::Scale(kIdentityLessWait)).count();
    return SetEnvironmentVariableW(kIdentityLessByNamePauseTimeoutEnv.data(), std::to_wstring(timeoutMs).c_str()) != FALSE &&
           SetEnvironmentVariableW(kIdentityLessByNamePauseRootEnv.data(), providerRoot.c_str()) != FALSE;
}

// Waits until the armed provider pause holds the delete (or the task ended without reaching it).
[[nodiscard]] bool WaitIdentityLessByNamePause(SelfTestState& state, uint64_t taskId, std::wstring_view what) noexcept
{
    const auto paused = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return GetEnvVarTrimmed(kIdentityLessByNamePauseEnteredEnv) == L"1" || state.completedTasks.contains(taskId);
    }, {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = what});
    return paused.conditionMet && GetEnvVarTrimmed(kIdentityLessByNamePauseEnteredEnv) == L"1";
}

[[nodiscard]] std::wstring DescribeItemReceipt(const std::optional<FileSystemItemMutationResult>& receipt)
{
    if (! receipt.has_value())
        return L"<none>";
    return std::format(L"known={} committed={} present={}", receipt->outcomeKnown, receipt->mutationCommitted, receipt->originalStillPresent);
}

[[nodiscard]] bool IsPartialTreeReceipt(const std::optional<FileSystemItemMutationResult>& receipt) noexcept
{
    return receipt.has_value() && receipt->outcomeKnown == TRUE && receipt->mutationCommitted == TRUE && receipt->originalStillPresent == TRUE;
}

[[nodiscard]] bool IsKeptReceipt(const std::optional<FileSystemItemMutationResult>& receipt) noexcept
{
    return receipt.has_value() && receipt->outcomeKnown == TRUE && receipt->mutationCommitted == FALSE && receipt->originalStillPresent == TRUE;
}

// I35 S4 round 2 (pinned tree): a file another program writes into the tree after the delete's last listing
// (a provider pause between that listing's check and the first deletion) is not in the listing, so it
// survives: the delete removes the listed entries, cannot remove the folder that still holds the new file,
// and reports a partial delete (the committed, original-present receipt), never "kept" for everything.
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeConcurrentWriterChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, concurrent writer: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-concurrent-writer";
    if (! RecreateEmptyDirectory(root / L"old" / L"Sub"))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(root) : root;
    const std::filesystem::path shared   = root / L"old";
    const std::filesystem::path late     = shared / L"Sub" / L"late.bin";
    if (! WriteFilledTestFile(shared / L"a.bin", 2048u, 0x81u) || ! WriteFilledTestFile(shared / L"Sub" / L"b.bin", 1024u, 0x82u))
        return reject(L"fixtures unavailable");
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location))
        return reject(L"the identity-less injection could not be configured");
    if (! ArmIdentityLessByNamePause(location / L"old"))
        return reject(L"the provider pause could not be armed");

    const auto deletion = scope.StartDelete({location / L"old"}, false);
    if (! deletion.has_value())
        return reject(L"the delete by name was not admitted");
    const auto card = scope.WaitForQuestion(deletion.value(), L"concurrent writer card");
    if (! card.has_value() || ! IsByNameCard(card.value()) || ! scope.Answer(deletion.value(), card.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the delete by name did not ask its card: {}", DescribeIdentityLessPrompt(card)));
    const auto paused       = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return GetEnvVarTrimmed(kIdentityLessByNamePauseEnteredEnv) == L"1" || state.completedTasks.contains(deletion.value());
    }, {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = L"concurrent writer pause"});
    const bool enteredPause = paused.conditionMet && GetEnvVarTrimmed(kIdentityLessByNamePauseEnteredEnv) == L"1";
    const bool wroteLate    = enteredPause && WriteFilledTestFile(late, 777u, 0x83u);
    static_cast<void>(SetEnvironmentVariableW(kIdentityLessByNamePauseReleaseEnv.data(), L"1"));
    if (! enteredPause || ! wroteLate)
        return reject(std::format(L"the delete never paused between its listing and its deletions (entered={} wrote={})", enteredPause, wroteLate));

    std::vector<IdentityLessPrompt> deleteQuestions;
    const CompletedTaskInfo* deleteEnd = scope.CancelQuestionsUntilEnd(deletion.value(), L"concurrent writer delete", &deleteQuestions);
    std::vector<IdentityLessDiagnostic> entries;
    operations->CollectDiagnostics(entries);
    // The item failed: one Error that names how much of the listing was deleted (2 of its 3 entries).
    const size_t partialNotes   = static_cast<size_t>(std::ranges::count_if(entries,
                                                                            [&](const IdentityLessDiagnostic& entry) noexcept
    {
        return entry.taskId == deletion.value() && entry.category == L"delete.permanent.byName" && entry.severity == IdentityLessSeverity::Error &&
               entry.message.find(L"deleting 2 of the 3 items") != std::wstring::npos;
    }));
    const auto receipt          = TakeFileOpsItemReceiptForSelfTest(deletion.value(), 0u);
    const bool lateKept         = FileSizeEquals(late, 777u);
    const bool listedRemoved    = ! std::filesystem::exists(shared / L"a.bin") && ! std::filesystem::exists(shared / L"Sub" / L"b.bin");
    const bool foldersKept      = std::filesystem::is_directory(shared) && std::filesystem::is_directory(shared / L"Sub");
    const std::wstring observed = std::format(L"deleteHr=0x{:08X} questions={} lateKept={} listedRemoved={} foldersKept={} partialNotes={} receipt=[{}]",
                                              deleteEnd != nullptr ? static_cast<unsigned long>(deleteEnd->hr) : 0ul,
                                              deleteQuestions.size(),
                                              lateKept,
                                              listedRemoved,
                                              foldersKept,
                                              partialNotes,
                                              DescribeItemReceipt(receipt));
    AppendLog(std::format(L"Pinned tree, concurrent writer: {}", observed));
    if (deleteEnd == nullptr || SUCCEEDED(deleteEnd->hr) || ! lateKept || ! listedRemoved || ! foldersKept || partialNotes == 0u ||
        ! IsPartialTreeReceipt(receipt))
        return reject(std::format(L"a file written after the listing was not kept, or the partial delete was not reported: {}", observed));
    Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeConcurrentWriter", L"late-file-kept;partial", 0u, 1u, 0u, S_OK);
    return true;
}

// I35 S4 round 2: a folder holding more entries than the by-name budget (a test-only budget of 2 here) is
// not deleted by name: no card, the task fails at Preparing with a plain reason, nothing is deleted.
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeBudgetChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, budget: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-budget";
    if (! RecreateEmptyDirectory(root / L"big"))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(root) : root;
    for (unsigned int index = 0u; index < 4u; ++index)
    {
        if (! WriteFilledTestFile(root / L"big" / std::format(L"f{}.bin", index), 512u, static_cast<uint8_t>(0x90u + index)))
            return reject(L"fixtures unavailable");
    }
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location) || SetEnvironmentVariableW(kIdentityLessByNameEntryBudgetEnv.data(), L"2") == FALSE)
        return reject(L"the identity-less injection or the test budget could not be configured");

    const auto deletion = scope.StartDelete({location / L"big"}, false);
    if (! deletion.has_value())
        return reject(L"the over-budget Delete was not admitted");
    const auto question = scope.WaitForQuestion(deletion.value(), L"over-budget Delete");
    if (question.has_value())
    {
        static_cast<void>(scope.CancelQuestionsUntilEnd(deletion.value(), L"over-budget Delete question"));
        return reject(std::format(L"a folder over the by-name budget was offered a delete by name: {}", DescribeIdentityLessPrompt(question)));
    }
    const CompletedTaskInfo* deleteEnd = scope.WaitForEnd(deletion.value(), L"over-budget Delete");
    const auto summary                 = FindIdentityLessSummary(operations, deletion.value());
    const bool reasoned = summary.has_value() && std::ranges::any_of(summary->issueDiagnostics, [](const IdentityLessDiagnostic& issue) noexcept {
        return issue.category == L"preparation.deleteRoot.byName" && issue.status == HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_QUOTA) &&
               issue.userText.messageId != 0u;
    });
    bool kept           = true;
    for (unsigned int index = 0u; index < 4u; ++index)
        kept = kept && FileSizeEquals(root / L"big" / std::format(L"f{}.bin", index), 512u);
    if (deleteEnd == nullptr || deleteEnd->hr != HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_QUOTA) || ! reasoned || ! kept)
        return reject(std::format(L"a folder over the by-name budget was not refused with its reason: hr=0x{:08X} reasoned={} kept={} issues=[{}]",
                                  deleteEnd != nullptr ? static_cast<unsigned long>(deleteEnd->hr) : 0ul,
                                  reasoned,
                                  kept,
                                  summary.has_value() ? DescribeIdentityLessIssues(summary.value()) : std::wstring(L"<no summary>")));
    Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeBudget", L"over-budget-refused", 0u, 1u, 0u, S_OK);
    return true;
}

// I35 S4 round 3: a listed entry that changes, or disappears, between the delete's last listing and its own
// deletion (the provider pause) stops the delete with ERROR_REVISION_MISMATCH before anything is deleted: the
// rest of the tree is kept, the stop says nothing was deleted, and the item's receipt is a known non-commit.
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeChangedEntryChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, changed entry: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-changed-entry";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(root) : root;
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location))
        return reject(L"the identity-less injection could not be configured");

    for (const bool missing : {false, true})
    {
        const std::wstring_view variant  = missing ? L"missing" : L"changed";
        const std::filesystem::path tree = root / variant;
        const std::filesystem::path file = tree / L"Sub" / L"b.bin";
        if (! RecreateEmptyDirectory(tree / L"Sub") || ! WriteFilledTestFile(file, 2048u, 0xA1u) || ! ArmIdentityLessByNamePause(location / variant))
            return reject(std::format(L"{}: fixture or provider pause unavailable", variant));
        const auto deletion = scope.StartDelete({location / variant}, false);
        if (! deletion.has_value())
            return reject(std::format(L"{}: the delete by name was not admitted", variant));
        const auto card = scope.WaitForQuestion(deletion.value(), L"changed entry card");
        if (! card.has_value() || ! IsByNameCard(card.value()) || ! scope.Answer(deletion.value(), card.value(), IdentityLessAction::PermanentDelete))
            return reject(std::format(L"{}: the delete by name did not ask its card: {}", variant, DescribeIdentityLessPrompt(card)));
        const bool paused = WaitIdentityLessByNamePause(state, deletion.value(), L"changed entry pause");
        bool changed      = false;
        std::error_code ec;
        if (paused)
            changed = missing ? std::filesystem::remove(file, ec) : WriteFilledTestFile(file, 999u, 0xA2u);
        static_cast<void>(SetEnvironmentVariableW(kIdentityLessByNamePauseReleaseEnv.data(), L"1"));
        if (! paused || ! changed)
            return reject(std::format(L"{}: the delete never paused before its deletions, or the entry could not be changed (paused={})", variant, paused));
        const CompletedTaskInfo* deleteEnd = scope.CancelQuestionsUntilEnd(deletion.value(), L"changed entry delete");
        const auto receipt                 = TakeFileOpsItemReceiptForSelfTest(deletion.value(), 0u);
        const size_t refusals              = CountLiveIdentityLessDiagnostics(
            operations, deletion.value(), L"delete.permanent.byName", HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), L"nothing was deleted");
        const bool kept             = std::filesystem::is_directory(tree / L"Sub") && (missing ? ! std::filesystem::exists(file) : FileSizeEquals(file, 999u));
        const std::wstring observed = std::format(L"{}: deleteHr=0x{:08X} refusals={} kept={} receipt=[{}]",
                                                  variant,
                                                  deleteEnd != nullptr ? static_cast<unsigned long>(deleteEnd->hr) : 0ul,
                                                  refusals,
                                                  kept,
                                                  DescribeItemReceipt(receipt));
        AppendLog(std::format(L"Pinned tree, changed entry: {}", observed));
        if (deleteEnd == nullptr || SUCCEEDED(deleteEnd->hr) || refusals == 0u || ! kept || ! IsKeptReceipt(receipt))
            return reject(std::format(L"a listed entry that changed after the listing did not stop the delete with nothing deleted: {}", observed));
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeChangedEntry", L"changed;missing", 0u, 2u, 0u, S_OK);
    return true;
}

// I35 S4 round 3: a refused Recycle escalated to a delete by name stops on a file written into the folder after
// its last listing: the listed files are deleted, the folder and the new file are kept, and the stop is reported
// as a partial delete (an Error naming N of M listed items) with the committed, original-present receipt.
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeEscalationPartialChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, escalation partial: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-escalation-partial";
    if (! RecreateEmptyDirectory(root / L"old" / L"Sub"))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(root) : root;
    const std::filesystem::path shared   = root / L"old";
    const std::filesystem::path late     = shared / L"Sub" / L"late.bin";
    if (! WriteFilledTestFile(shared / L"a.bin", 2048u, 0xB1u) || ! WriteFilledTestFile(shared / L"Sub" / L"b.bin", 1024u, 0xB2u))
        return reject(L"fixtures unavailable");
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location) ||
        SetEnvironmentVariableW(kIdentityLessRecycleFailEnv.data(), (location / L"old").c_str()) == FALSE || ! ArmIdentityLessByNamePause(location / L"old"))
        return reject(L"the identity-less, Recycle-failure or pause injection could not be configured");

    const auto recycle = scope.StartDelete({location / L"old"}, true);
    if (! recycle.has_value())
        return reject(L"the Recycle was not admitted");
    const auto escalation = scope.WaitForQuestion(recycle.value(), L"escalation partial question");
    if (! escalation.has_value() || ! IsByNameEscalation(escalation.value(), false) ||
        ! scope.Answer(recycle.value(), escalation.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the refused Recycle did not ask its delete by name: {}", DescribeIdentityLessPrompt(escalation)));
    const bool paused  = WaitIdentityLessByNamePause(state, recycle.value(), L"escalation partial pause");
    const bool wroteIt = paused && WriteFilledTestFile(late, 777u, 0xB3u);
    static_cast<void>(SetEnvironmentVariableW(kIdentityLessByNamePauseReleaseEnv.data(), L"1"));
    if (! paused || ! wroteIt)
        return reject(std::format(L"the escalated delete never paused between its listing and its deletions (paused={} wrote={})", paused, wroteIt));

    const CompletedTaskInfo* end = scope.CancelQuestionsUntilEnd(recycle.value(), L"escalation partial end");
    const auto receipt           = TakeFileOpsItemReceiptForSelfTest(recycle.value(), 0u);
    std::vector<IdentityLessDiagnostic> entries;
    operations->CollectDiagnostics(entries);
    const size_t partialNotes   = static_cast<size_t>(std::ranges::count_if(entries,
                                                                            [&](const IdentityLessDiagnostic& entry) noexcept
    {
        return entry.taskId == recycle.value() && entry.category == L"delete.recycle.byName" && entry.severity == IdentityLessSeverity::Error &&
               entry.message.find(L"deleting 2 of the 3 items") != std::wstring::npos;
    }));
    const bool lateKept         = FileSizeEquals(late, 777u);
    const bool listedRemoved    = ! std::filesystem::exists(shared / L"a.bin") && ! std::filesystem::exists(shared / L"Sub" / L"b.bin");
    const bool foldersKept      = std::filesystem::is_directory(shared / L"Sub");
    const std::wstring observed = std::format(L"hr=0x{:08X} lateKept={} listedRemoved={} foldersKept={} partialNotes={} receipt=[{}]",
                                              end != nullptr ? static_cast<unsigned long>(end->hr) : 0ul,
                                              lateKept,
                                              listedRemoved,
                                              foldersKept,
                                              partialNotes,
                                              DescribeItemReceipt(receipt));
    AppendLog(std::format(L"Pinned tree, escalation partial: {}", observed));
    if (end == nullptr || SUCCEEDED(end->hr) || ! lateKept || ! listedRemoved || ! foldersKept || partialNotes == 0u || ! IsPartialTreeReceipt(receipt))
        return reject(std::format(L"the escalated delete did not report its partial delete: {}", observed));
    Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeEscalationPartial", L"late-file-kept;partial", 0u, 1u, 0u, S_OK);
    return true;
}

// I35 S4 round 3 (hardening): a listed folder swapped for a junction after the delete's last listing never
// redirects the delete, even with a listing that reports no file index (the per-entry facts are then the only
// check): the file the junction leads to, forged with the listed file's name, size, times and attributes,
// survives. Local spelling: NTFS follows the junction for any open below it.
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeLinkSwapChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, link swap: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root    = state.tempRoot / L"identity-less-link-swap";
    const std::filesystem::path tree    = root / L"old";
    const std::filesystem::path sub     = tree / L"Sub";
    const std::filesystem::path listed  = sub / L"b.bin";
    const std::filesystem::path outside = root / L"outside";
    const std::filesystem::path decoy   = outside / L"b.bin";
    const std::filesystem::path parked  = root / L"parked";
    if (! RecreateEmptyDirectory(sub) || ! RecreateEmptyDirectory(outside) || ! RecreateEmptyDirectory(parked) || ! WriteFilledTestFile(listed, 1024u, 0xC1u) ||
        ! WriteFilledTestFile(decoy, 1024u, 0xC2u))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(RemoveDirectoryW(sub.c_str())); // the junction, never its target
        static_cast<void>(SelfTest::RemoveAll(root));
    });
    // The decoy carries the listed file's times, so only an identity could tell the two apart.
    {
        wil::unique_hfile source(
            CreateFileW(listed.c_str(), FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, 0, nullptr));
        wil::unique_hfile target(
            CreateFileW(decoy.c_str(), FILE_WRITE_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, 0, nullptr));
        FILETIME creation{};
        FILETIME access{};
        FILETIME write{};
        if (! source || ! target || GetFileTime(source.get(), &creation, &access, &write) == FALSE ||
            SetFileTime(target.get(), &creation, &access, &write) == FALSE)
            return reject(L"the decoy could not take the listed file's times");
    }
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(root) || SetEnvironmentVariableW(kIdentityLessByNameListWithoutIndexEnv.data(), L"1") == FALSE ||
        ! ArmIdentityLessByNamePause(tree))
        return reject(L"the identity-less, index-less listing or pause injection could not be configured");

    const auto deletion = scope.StartDelete({tree}, false);
    if (! deletion.has_value())
        return reject(L"the delete by name was not admitted");
    const auto card = scope.WaitForQuestion(deletion.value(), L"link swap card");
    if (! card.has_value() || ! IsByNameCard(card.value()) || ! scope.Answer(deletion.value(), card.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the delete by name did not ask its card: {}", DescribeIdentityLessPrompt(card)));
    const bool paused  = WaitIdentityLessByNamePause(state, deletion.value(), L"link swap pause");
    const bool swapped = paused && MoveFileExW(sub.c_str(), (parked / L"Sub").c_str(), 0) != FALSE && TryCreateJunction(sub, outside);
    static_cast<void>(SetEnvironmentVariableW(kIdentityLessByNamePauseReleaseEnv.data(), L"1"));
    if (! paused || ! swapped)
        return reject(std::format(L"the listed folder could not be swapped for a junction at the pause (paused={} swapped={})", paused, swapped));

    const CompletedTaskInfo* deleteEnd = scope.CancelQuestionsUntilEnd(deletion.value(), L"link swap delete");
    const bool decoyKept               = FileSizeEquals(decoy, 1024u);
    const bool originalKept            = FileSizeEquals(parked / L"Sub" / L"b.bin", 1024u);
    const std::wstring observed        = std::format(
        L"deleteHr=0x{:08X} decoyKept={} originalKept={}", deleteEnd != nullptr ? static_cast<unsigned long>(deleteEnd->hr) : 0ul, decoyKept, originalKept);
    AppendLog(std::format(L"Pinned tree, link swap: {}", observed));
    if (deleteEnd == nullptr || SUCCEEDED(deleteEnd->hr) || ! decoyKept || ! originalKept)
        return reject(std::format(L"a junction swapped in after the listing redirected the delete: {}", observed));
    Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeLinkSwap", L"junction-refused", 0u, 1u, 0u, S_OK);
    return true;
}

// I35 S4 round 3: a sub-folder the location refuses to open with ERROR_NOT_SUPPORTED while the folder is listed
// for a delete by name is an open failure (ERROR_OPEN_FAILED, refused at Preparing, nothing deleted), never the
// identity-delete contract's "the location can identify this object".
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeSubfolderOpenChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, sub-folder open: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-subfolder-open";
    if (! RecreateEmptyDirectory(root / L"old" / L"Sub") || ! WriteFilledTestFile(root / L"old" / L"a.bin", 512u, 0xD1u) ||
        ! WriteFilledTestFile(root / L"old" / L"Sub" / L"b.bin", 512u, 0xD2u))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(root) : root;
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location) ||
        SetEnvironmentVariableW(kIdentityLessBindOpenNotSupportedEnv.data(), (location / L"old" / L"Sub").c_str()) == FALSE)
        return reject(L"the identity-less or open-refusal injection could not be configured");

    const auto deletion = scope.StartDelete({location / L"old"}, false);
    if (! deletion.has_value())
        return reject(L"the Delete was not admitted");
    const auto question = scope.WaitForQuestion(deletion.value(), L"sub-folder open refusal");
    if (question.has_value())
    {
        static_cast<void>(scope.CancelQuestionsUntilEnd(deletion.value(), L"sub-folder open refusal question"));
        return reject(std::format(L"a folder whose sub-folder could not be opened was offered a delete by name: {}", DescribeIdentityLessPrompt(question)));
    }
    const CompletedTaskInfo* end = scope.WaitForEnd(deletion.value(), L"sub-folder open refusal");
    const auto summary           = FindIdentityLessSummary(operations, deletion.value());
    const bool reasoned          = summary.has_value() && std::ranges::any_of(summary->issueDiagnostics, [](const IdentityLessDiagnostic& issue) noexcept {
        return issue.category == L"preparation.deleteRoot.byName" && issue.status == HRESULT_FROM_WIN32(ERROR_OPEN_FAILED);
    });
    const bool kept              = FileSizeEquals(root / L"old" / L"a.bin", 512u) && FileSizeEquals(root / L"old" / L"Sub" / L"b.bin", 512u);
    if (end == nullptr || end->hr != HRESULT_FROM_WIN32(ERROR_OPEN_FAILED) || ! reasoned || ! kept)
        return reject(
            std::format(L"a sub-folder refused with ERROR_NOT_SUPPORTED was not reported as an open failure: hr=0x{:08X} reasoned={} kept={} issues=[{}]",
                        end != nullptr ? static_cast<unsigned long>(end->hr) : 0ul,
                        reasoned,
                        kept,
                        summary.has_value() ? DescribeIdentityLessIssues(summary.value()) : std::wstring(L"<no summary>")));
    Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeSubfolderOpen", L"open-failed", 0u, 1u, 0u, S_OK);
    return true;
}

// I35 S4 round 3: a Recycle of a folder over the by-name bounds keeps the reason its pin failed; once the Recycle
// is refused, the escalation asks nothing and says that the folder holds too many items to check by name.
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeRecycleOverBudgetChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, Recycle over budget: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-recycle-budget";
    if (! RecreateEmptyDirectory(root / L"big"))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(root) : root;
    for (unsigned int index = 0u; index < 4u; ++index)
    {
        if (! WriteFilledTestFile(root / L"big" / std::format(L"f{}.bin", index), 512u, static_cast<uint8_t>(0xE0u + index)))
            return reject(L"fixtures unavailable");
    }
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location) || SetEnvironmentVariableW(kIdentityLessByNameEntryBudgetEnv.data(), L"2") == FALSE ||
        SetEnvironmentVariableW(kIdentityLessRecycleFailEnv.data(), (location / L"big").c_str()) == FALSE)
        return reject(L"the identity-less, budget or Recycle-failure injection could not be configured");

    const auto recycle = scope.StartDelete({location / L"big"}, true);
    if (! recycle.has_value())
        return reject(L"the Recycle was not admitted");
    const auto question = scope.WaitForQuestion(recycle.value(), L"Recycle over budget");
    if (question.has_value())
    {
        static_cast<void>(scope.CancelQuestionsUntilEnd(recycle.value(), L"Recycle over budget question"));
        return reject(std::format(L"a Recycle over the by-name bounds asked a question: {}", DescribeIdentityLessPrompt(question)));
    }
    const CompletedTaskInfo* end = scope.WaitForEnd(recycle.value(), L"Recycle over budget");
    const auto summary           = FindIdentityLessSummary(operations, recycle.value());
    const bool reasoned          = summary.has_value() && std::ranges::any_of(summary->issueDiagnostics,
                                                                              [](const IdentityLessDiagnostic& issue) noexcept
    {
        return issue.category == L"delete.recycle.byName" && issue.status == HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_QUOTA) &&
               issue.userText.messageId == IDS_FILEOPS_NOTE_BY_NAME_TOO_MANY_ITEMS;
    });
    bool kept                    = true;
    for (unsigned int index = 0u; index < 4u; ++index)
        kept = kept && FileSizeEquals(root / L"big" / std::format(L"f{}.bin", index), 512u);
    if (end == nullptr || SUCCEEDED(end->hr) || ! reasoned || ! kept)
        return reject(std::format(L"the refused Recycle did not keep the reason of its pin: hr=0x{:08X} reasoned={} kept={} issues=[{}]",
                                  end != nullptr ? static_cast<unsigned long>(end->hr) : 0ul,
                                  reasoned,
                                  kept,
                                  summary.has_value() ? DescribeIdentityLessIssues(summary.value()) : std::wstring(L"<no summary>")));
    Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeRecycleOverBudget", L"reason-kept", 0u, 1u, 0u, S_OK);
    return true;
}

// I35 S4 round 3: another live task that holds the folder of a delete by name and will not wait for it (its
// overlap question decided, or running) refuses the delete with heldByAnotherTask, at the re-check after the
// card's answer and at the re-check right before the delete. The holder is a published registration that reads
// inside the folder.
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeHolderChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, holders: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-holders";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(root) : root;
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location))
        return reject(L"the identity-less injection could not be configured");

    for (const bool atExecution : {false, true})
    {
        const std::wstring_view stage = atExecution ? L"execution" : L"recheck";
        if (! RecreateEmptyDirectory(root / stage) || ! WriteFilledTestFile(root / stage / L"x.bin", 2048u, 0xF1u))
            return reject(std::format(L"{}: fixture unavailable", stage));
        if (atExecution)
            SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(true);
        else
            SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(true);
        const auto deletion = scope.StartDelete({location / stage}, false);
        if (! deletion.has_value())
            return reject(std::format(L"{}: the delete by name was not admitted", stage));
        const auto card = scope.WaitForQuestion(deletion.value(), L"holder card");
        if (! card.has_value() || ! IsByNameCard(card.value()) || ! scope.Answer(deletion.value(), card.value(), IdentityLessAction::PermanentDelete))
            return reject(std::format(L"{}: the delete by name did not ask its card: {}", stage, DescribeIdentityLessPrompt(card)));
        const auto paused = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            return (atExecution ? HasFileOpsPermanentDeleteBeforeLiveOutputGuardPauseEnteredForSelfTest()
                                : HasFileOpsPermanentDeleteBeforeRecheckPauseEnteredForSelfTest()) ||
                   state.completedTasks.contains(deletion.value());
        },
            {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = L"holder pause"});

        // The holder reads x.bin inside the folder through the delete's own endpoint; its question decided
        // nothing about the delete (no Queue after), so it holds the path.
        std::unique_ptr<IdentityLessTask> holder;
        IdentityLessTask* const deleteTask = operations->FindTask(deletion.value());
        const auto deleteScopes            = deleteTask != nullptr ? deleteTask->LoadMutationInterlockScopes() : nullptr;
        if (paused.conditionMet && deleteScopes)
        {
            const auto byName = std::ranges::find_if(
                *deleteScopes, [](const FileOperations::MutationInterlockScope& candidate) noexcept { return candidate.byNameCorroborated; });
            if (byName != deleteScopes->end())
            {
                FileOperations::MutationInterlockScope reading = *byName;
                reading.access                                 = FileOperations::MutationInterlockAccess::ReadSource;
                reading.byNameCorroborated                     = false;
                reading.pinnedDeleteIdentity.reset();
                reading.providerPath = (std::filesystem::path(byName->providerPath) / L"x.bin").native();
                holder               = std::make_unique<IdentityLessTask>(*operations);
                holder->_taskId      = operations->AllocateLateQueueOrderKey();
                holder->_admissionKey.store(holder->_taskId, std::memory_order_release);
                holder->InstallPreparedMutationInterlockScopes({reading}, false);
                operations->PublishPreparedMutationInterlock(*holder);
                FolderWindow::FileOperationState::MutationInterlockRelations decided{};
                decided.advisoryDecided = true;
                operations->PublishMutationInterlockDecision(*holder, decided);
            }
        }
        const uint64_t holderKey = holder ? holder->_admissionKey.load(std::memory_order_acquire) : 0u;
        if (atExecution)
        {
            ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest();
            SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(false);
        }
        else
        {
            ReleaseFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest();
            SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(false);
        }
        const CompletedTaskInfo* end = scope.CancelQuestionsUntilEnd(deletion.value(), L"holder delete");
        if (holderKey != 0u)
            operations->WithdrawPreparedMutationInterlock(holderKey);
        holder.reset();
        const size_t refused =
            CountLiveIdentityLessDiagnostics(operations, deletion.value(), L"delete.byName.heldByAnotherTask", HRESULT_FROM_WIN32(ERROR_BUSY));
        const bool kept = FileSizeEquals(root / stage / L"x.bin", 2048u);
        // Refused at Preparing the task ends ERROR_BUSY; refused at the delete, the item's question was answered Cancel.
        if (holderKey == 0u || end == nullptr || (atExecution ? SUCCEEDED(end->hr) : end->hr != HRESULT_FROM_WIN32(ERROR_BUSY)) || refused == 0u || ! kept)
            return reject(std::format(L"{}: a holder that will not wait for the delete did not refuse it: holder={} hr=0x{:08X} refused={} kept={}",
                                      stage,
                                      holderKey != 0u,
                                      end != nullptr ? static_cast<unsigned long>(end->hr) : 0ul,
                                      refused,
                                      kept));
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeHolders", L"recheck;execution", 0u, 2u, 0u, S_OK);
    return true;
}

// I35 S4 round 3 (M1): a delete by name of a tree of a few thousand entries, on the loopback alias, again there
// with a test-only budget of four held folders (each top folder held alone, its sub-folders in chunks of four),
// and, where path.smb.primary is configured, on a per-run child of the registered share. The log and the perf record keep
// the time from the start to the card (the pin listing) and from the answer to the end (the re-check listing,
// the execution listing and the deletions); the case asserts only that the whole tree is deleted.
[[nodiscard]] bool RunDeleteIdentityLessPinnedTreeScaleChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Pinned tree, scale: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    constexpr unsigned int kTopFolders     = 6u;
    constexpr unsigned int kSubFolders     = 10u;
    constexpr unsigned int kFilesPerFolder = 40u;
    constexpr uint64_t kEntries            = kTopFolders + kTopFolders * kSubFolders + kTopFolders * kSubFolders * kFilesPerFolder;
    const auto buildTree                   = [&](const std::filesystem::path& tree) noexcept
    {
        if (! RecreateEmptyDirectory(tree))
            return false;
        for (unsigned int top = 0u; top < kTopFolders; ++top)
        {
            for (unsigned int sub = 0u; sub < kSubFolders; ++sub)
            {
                const std::filesystem::path folderPath = tree / std::format(L"t{}", top) / std::format(L"s{}", sub);
                if (! RecreateEmptyDirectory(folderPath))
                    return false;
                for (unsigned int file = 0u; file < kFilesPerFolder; ++file)
                {
                    if (! WriteFilledTestFile(folderPath / std::format(L"f{}.bin", file), 64u, static_cast<uint8_t>(file)))
                        return false;
                }
            }
        }
        return true;
    };
    const auto measure = [&](std::wstring_view where, const std::filesystem::path& tree, const std::filesystem::path& providerTree) noexcept -> bool
    {
        IdentityLessCaseScope scope(state, operations);
        const auto started  = std::chrono::steady_clock::now();
        const auto deletion = scope.StartDelete({providerTree}, false);
        if (! deletion.has_value())
            return reject(std::format(L"{}: the delete by name was not admitted", where));
        const auto card     = scope.WaitForQuestion(deletion.value(), L"scale card");
        const auto answered = std::chrono::steady_clock::now();
        if (! card.has_value() || ! IsByNameCard(card.value()) || ! scope.Answer(deletion.value(), card.value(), IdentityLessAction::PermanentDelete))
            return reject(std::format(L"{}: the delete by name did not ask its card: {}", where, DescribeIdentityLessPrompt(card)));
        const CompletedTaskInfo* end = scope.CancelQuestionsUntilEnd(deletion.value(), L"scale delete");
        const auto ended             = std::chrono::steady_clock::now();
        const auto cardMs            = std::chrono::duration_cast<std::chrono::milliseconds>(answered - started).count();
        const auto runMs             = std::chrono::duration_cast<std::chrono::milliseconds>(ended - answered).count();
        const bool deleted           = ! std::filesystem::exists(tree);
        AppendLog(std::format(L"Pinned tree, scale ({}): entries={} card={} ms run={} ms hr=0x{:08X} deleted={}",
                              where,
                              kEntries,
                              cardMs,
                              runMs,
                              end != nullptr ? static_cast<unsigned long>(end->hr) : 0ul,
                              deleted));
        Debug::Perf::Emit(L"FileOps.SelfTest.PinnedTreeScale",
                          where,
                          static_cast<uint64_t>(runMs) * 1000u,
                          kEntries,
                          static_cast<uint64_t>(cardMs),
                          end != nullptr ? end->hr : E_FAIL);
        if (end == nullptr || end->hr != S_OK || ! deleted)
            return reject(std::format(
                L"{}: the tree was not deleted by name (hr=0x{:08X} deleted={})", where, end != nullptr ? static_cast<unsigned long>(end->hr) : 0ul, deleted));
        return true;
    };

    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");
    const std::filesystem::path root = state.tempRoot / L"identity-less-scale";
    const auto removeRoot            = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    if (! buildTree(root / L"tree"))
        return reject(L"local tree unavailable");
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(root) : root;
    if (! IdentityLessCaseScope::MakeIdentityLess(location) ||
        ! measure(uncStatus == LoopbackShareStatus::Reachable ? L"loopback" : L"local", root / L"tree", location / L"tree"))
        return false;
    // The same tree with a hold budget of 4 folders: no top folder's subtree (11 folders) fits, so each is held
    // alone and its sub-folders go in chunks of four (the case scope clears the budget and the location).
    if (! buildTree(root / L"tree") || ! IdentityLessCaseScope::MakeIdentityLess(location) ||
        SetEnvironmentVariableW(kIdentityLessByNameHoldBudgetEnv.data(), L"4") == FALSE)
        return reject(L"local tree for the chunked delete unavailable");
    if (! measure(L"chunks", root / L"tree", location / L"tree"))
        return false;
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvBindIdentityUnsupportedPath.data(), nullptr));

    // The registered NAS share (path.smb.primary) is identity-less by itself.
    const std::wstring smbRootText = GetEnvVarTrimmed(L"REDSALAMANDER_SELFTEST_SMB_ROOT");
    if (smbRootText.empty())
    {
        AppendLog(L"Pinned tree, scale: the NAS part is skipped (path.smb.primary is not configured).");
        return true;
    }
    const std::filesystem::path child = std::filesystem::path(smbRootText) /
                                        std::format(L"identity-less-scale-{}-{:X}", GetCurrentProcessId(), static_cast<unsigned long long>(GetTickCount64()));
    const auto removeChild            = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(child)); });
    if (! buildTree(child / L"tree"))
        return reject(std::format(L"NAS tree under '{}' unavailable", child.native()));
    return measure(L"nas", child / L"tree", child / L"tree");
}

// I35 S4 (red-team HIGH c and b, round 2): a folder deleted by name carries its last-write time in its
// facts, so a child added after the pin, while the card waits or between the answer and the delete,
// refuses with ERROR_REVISION_MISMATCH and deletes nothing. A task admitted after the answer, before the
// re-check or while the delete is about to run, sees the published delete: it is asked about it, cannot
// run at the same time, and its pending question does not refuse the delete.
[[nodiscard]] bool RunDeleteIdentityLessFolderChangeAndHolderChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Identity-less folder change and holders: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-folder-holders";
    if (! RecreateEmptyDirectory(root / L"share") || ! RecreateEmptyDirectory(root / L"copy-out"))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path share    = root / L"share";
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(share) : share;
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location))
        return reject(L"the identity-less injection could not be configured");
    const std::wstring heldCategory = L"delete.byName.heldByAnotherTask";

    // 1. A child added while the card waits: the re-check after the answer refuses, nothing deleted.
    const std::filesystem::path cardFolder = share / L"child-on-card";
    if (! RecreateEmptyDirectory(cardFolder) || ! WriteFilledTestFile(cardFolder / L"a.bin", 2048u, 0x51u))
        return reject(L"card-time fixture unavailable");
    const auto cardTask = scope.StartDelete({location / cardFolder.filename()}, false);
    if (! cardTask.has_value())
        return reject(L"the card-time Delete was not admitted");
    const auto cardQuestion = scope.WaitForQuestion(cardTask.value(), L"card-time child card");
    if (! cardQuestion.has_value() || ! IsByNameCard(cardQuestion.value()) || ! WriteFilledTestFile(cardFolder / L"added.bin", 1024u, 0x52u) ||
        ! scope.Answer(cardTask.value(), cardQuestion.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the card-time folder could not be changed while its card waited: {}", DescribeIdentityLessPrompt(cardQuestion)));
    const CompletedTaskInfo* cardEnd = scope.WaitForEnd(cardTask.value(), L"card-time child Delete");
    const auto cardSummary           = FindIdentityLessSummary(operations, cardTask.value());
    if (cardEnd == nullptr || cardEnd->hr != HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) || ! FileSizeEquals(cardFolder / L"a.bin", 2048u) ||
        ! FileSizeEquals(cardFolder / L"added.bin", 1024u) || ! cardSummary.has_value() ||
        CountIdentityLessIssues(
            cardSummary.value(), L"delete.confirmation.identityChanged", IdentityLessSeverity::Error, HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH)) != 1u)
        return reject(std::format(L"a child added while the card waited was not refused: hr=0x{:08X} a={} added={} issues=[{}]",
                                  cardEnd != nullptr ? static_cast<unsigned long>(cardEnd->hr) : 0ul,
                                  std::filesystem::exists(cardFolder / L"a.bin"),
                                  std::filesystem::exists(cardFolder / L"added.bin"),
                                  cardSummary.has_value() ? DescribeIdentityLessIssues(cardSummary.value()) : std::wstring(L"<no summary>")));

    // 2. A child added between the answer and the delete: the deleting handle refuses before the walk.
    const std::filesystem::path lateFolder = share / L"child-before-delete";
    if (! RecreateEmptyDirectory(lateFolder) || ! WriteFilledTestFile(lateFolder / L"a.bin", 2048u, 0x53u))
        return reject(L"delete-time fixture unavailable");
    SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(true);
    const auto lateTask = scope.StartDelete({location / lateFolder.filename()}, false);
    if (! lateTask.has_value())
        return reject(L"the delete-time Delete was not admitted");
    const auto lateCard = scope.WaitForQuestion(lateTask.value(), L"delete-time child card");
    if (! lateCard.has_value() || ! IsByNameCard(lateCard.value()) || ! scope.Answer(lateTask.value(), lateCard.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the delete-time card could not be answered: {}", DescribeIdentityLessPrompt(lateCard)));
    const auto latePaused    = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return HasFileOpsPermanentDeleteBeforeLiveOutputGuardPauseEnteredForSelfTest() || state.completedTasks.contains(lateTask.value());
    }, {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = L"delete-time child pause"});
    const bool addedInWindow = latePaused.conditionMet && HasFileOpsPermanentDeleteBeforeLiveOutputGuardPauseEnteredForSelfTest() &&
                               WriteFilledTestFile(lateFolder / L"added.bin", 1024u, 0x54u);
    ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest();
    SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(false);
    if (! addedInWindow)
        return reject(L"a child could not be added between the answer and the delete");
    std::vector<IdentityLessPrompt> lateQuestions;
    const CompletedTaskInfo* lateEnd = scope.CancelQuestionsUntilEnd(lateTask.value(), L"delete-time child Delete", &lateQuestions);
    const std::wstring mismatchText =
        std::format(L"openHr=0x00000000 identityHr=0x{:08X}", static_cast<unsigned long>(HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH)));
    const size_t lateRefusals =
        CountLiveIdentityLessDiagnostics(operations, lateTask.value(), L"delete.permanent.byName", HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), mismatchText);
    const bool lateQuestionsSafe = std::ranges::none_of(
        lateQuestions, [](const IdentityLessPrompt& prompt) noexcept { return PromptHasAction(prompt, IdentityLessAction::PermanentDelete); });
    if (lateEnd == nullptr || SUCCEEDED(lateEnd->hr) || lateRefusals != 1u || ! lateQuestionsSafe || ! FileSizeEquals(lateFolder / L"a.bin", 2048u) ||
        ! FileSizeEquals(lateFolder / L"added.bin", 1024u))
        return reject(std::format(L"a child added between the answer and the delete was not refused on the deleting handle: hr=0x{:08X} "
                                  L"refusals={} questions={} a={} added={}",
                                  lateEnd != nullptr ? static_cast<unsigned long>(lateEnd->hr) : 0ul,
                                  lateRefusals,
                                  lateQuestions.size(),
                                  std::filesystem::exists(lateFolder / L"a.bin"),
                                  std::filesystem::exists(lateFolder / L"added.bin")));

    // 3. A task admitted between the answer and its re-check (a Recycle of a child): the delete is already
    //    published, so the Recycle is asked about it (by-name problem, no Run at the same time) and cannot
    //    touch the folder first; its pending question does not refuse the delete, which removes the folder.
    const std::filesystem::path recheckFolder = share / L"held-after-answer";
    if (! RecreateEmptyDirectory(recheckFolder) || ! WriteFilledTestFile(recheckFolder / L"x.bin", 2048u, 0x55u) ||
        SetEnvironmentVariableW(kIdentityLessRecycleFailEnv.data(), (location / recheckFolder.filename() / L"x.bin").c_str()) == FALSE)
        return reject(L"re-check fixture unavailable");
    SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(true);
    const auto recheckTask = scope.StartDelete({location / recheckFolder.filename()}, false);
    if (! recheckTask.has_value())
        return reject(L"the re-check Delete was not admitted");
    const auto recheckCard = scope.WaitForQuestion(recheckTask.value(), L"re-check card");
    if (! recheckCard.has_value() || ! IsByNameCard(recheckCard.value()) ||
        ! scope.Answer(recheckTask.value(), recheckCard.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the re-check card could not be answered: {}", DescribeIdentityLessPrompt(recheckCard)));
    const auto recheckPaused = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return HasFileOpsPermanentDeleteBeforeRecheckPauseEnteredForSelfTest() || state.completedTasks.contains(recheckTask.value());
    }, {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = L"re-check pause"});
    std::optional<uint64_t> recheckLater;
    std::optional<IdentityLessPrompt> recheckLaterQuestion;
    if (recheckPaused.conditionMet && HasFileOpsPermanentDeleteBeforeRecheckPauseEnteredForSelfTest())
    {
        recheckLater = scope.StartDelete({location / recheckFolder.filename() / L"x.bin"}, true);
        if (recheckLater.has_value())
            recheckLaterQuestion = scope.WaitForQuestion(recheckLater.value(), L"re-check later task question");
    }
    ReleaseFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest();
    SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(false);
    const bool recheckLaterAsked =
        recheckLaterQuestion.has_value() &&
        IsByNameOverlapQuestion(recheckLaterQuestion.value(), recheckTask.value(), FileOperations::SameHostOverlapProblem::DeletedByNameWritten);
    const CompletedTaskInfo* recheckEnd = scope.WaitForEnd(recheckTask.value(), L"re-check Delete");
    // R7.5: the leftover listing leaves out the rows of a folder a live task works in: here the later task,
    // whose question still waits once the delete (itself published from its card on) has ended.
    const uint64_t workingInHeldFolder =
        recheckLater.has_value() ? operations->FindLiveTaskWorkingIn(kPluginIdLocal, (location / recheckFolder.filename()).native()) : 0u;
    const uint64_t workingInOtherFolder = operations->FindLiveTaskWorkingIn(kPluginIdLocal, (location / L"not-held").native());
    if (recheckLater.has_value())
        static_cast<void>(scope.CancelQuestionsUntilEnd(recheckLater.value(), L"re-check later task end"));
    static_cast<void>(SetEnvironmentVariableW(kIdentityLessRecycleFailEnv.data(), nullptr));
    if (! recheckLater.has_value() || ! recheckLaterAsked || recheckEnd == nullptr || recheckEnd->hr != S_OK || std::filesystem::exists(recheckFolder))
        return reject(std::format(L"a task admitted after the answer was not asked about the delete, or blocked it: question=[{}] deleteHr=0x{:08X} "
                                  L"folder={}",
                                  DescribeOverlapQuestion(recheckLaterQuestion),
                                  recheckEnd != nullptr ? static_cast<unsigned long>(recheckEnd->hr) : 0ul,
                                  std::filesystem::exists(recheckFolder)));
    if (workingInHeldFolder != recheckLater.value() || workingInOtherFolder != 0u)
        return reject(std::format(L"the live-task folder check of the leftover listing was wrong: held={} (expected {}) other={}",
                                  workingInHeldFolder,
                                  recheckLater.value(),
                                  workingInOtherFolder));

    // 4. A task admitted while the delete is about to run (a Copy out of the folder): asked about the delete
    //    (it reads there; no Run at the same time), it waits for its answer and the delete proceeds.
    const std::filesystem::path runFolder = share / L"held-at-delete";
    if (! RecreateEmptyDirectory(runFolder) || ! WriteFilledTestFile(runFolder / L"y.bin", 2048u, 0x56u))
        return reject(L"delete-time fixture unavailable");
    SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(true);
    const auto runTask = scope.StartDelete({location / runFolder.filename()}, false);
    if (! runTask.has_value())
        return reject(L"the delete-time Delete was not admitted");
    const auto runCard = scope.WaitForQuestion(runTask.value(), L"delete-time card");
    if (! runCard.has_value() || ! IsByNameCard(runCard.value()) || ! scope.Answer(runTask.value(), runCard.value(), IdentityLessAction::PermanentDelete))
        return reject(std::format(L"the delete-time card could not be answered: {}", DescribeIdentityLessPrompt(runCard)));
    const auto runPaused = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return HasFileOpsPermanentDeleteBeforeLiveOutputGuardPauseEnteredForSelfTest() || state.completedTasks.contains(runTask.value());
    }, {.timeout = SelfTest::Scale(kIdentityLessWait), .operationName = L"delete-time pause"});
    std::optional<uint64_t> copyOut;
    std::optional<IdentityLessPrompt> copyOutQuestion;
    if (runPaused.conditionMet && HasFileOpsPermanentDeleteBeforeLiveOutputGuardPauseEnteredForSelfTest())
    {
        copyOut = StartFileOperationAndGetId(operations,
                                             FILESYSTEM_COPY,
                                             FolderWindow::Pane::Left,
                                             FolderWindow::Pane::Right,
                                             state.fsLocal,
                                             {location / runFolder.filename() / L"y.bin"},
                                             root / L"copy-out",
                                             FILESYSTEM_FLAG_NONE,
                                             false,
                                             0u,
                                             FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                             false,
                                             state.fsLocal);
        if (copyOut.has_value())
        {
            scope.Track(copyOut.value());
            copyOutQuestion = scope.WaitForQuestion(copyOut.value(), L"delete-time Copy question");
        }
    }
    ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest();
    SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(false);
    const bool copyOutAsked = copyOutQuestion.has_value() &&
                              IsByNameOverlapQuestion(copyOutQuestion.value(), runTask.value(), FileOperations::SameHostOverlapProblem::DeletedByNameRead);
    const CompletedTaskInfo* runEnd = scope.CancelQuestionsUntilEnd(runTask.value(), L"delete-time Delete");
    if (copyOut.has_value())
        static_cast<void>(scope.CancelQuestionsUntilEnd(copyOut.value(), L"delete-time Copy end"));
    if (! copyOut.has_value() || ! copyOutAsked || runEnd == nullptr || runEnd->hr != S_OK || std::filesystem::exists(runFolder))
        return reject(std::format(L"a task admitted while the delete was about to run was not asked about it, or blocked it: question=[{}] "
                                  L"deleteHr=0x{:08X} folder={}",
                                  DescribeOverlapQuestion(copyOutQuestion),
                                  runEnd != nullptr ? static_cast<unsigned long>(runEnd->hr) : 0ul,
                                  std::filesystem::exists(runFolder)));
    Debug::Perf::Emit(L"FileOps.SelfTest.IdentityLessFolderChangeAndHolders", L"card;delete;later-task-asked;running-task-asked", 0u, 4u, 0u, S_OK);
    return true;
}

// I35 S4 (red-team MEDIUM 2, LOW 6): an open the location refuses with ERROR_NOT_SUPPORTED is a failed
// open (Indeterminate), never an identity-less object: the Delete is refused at Preparing with both
// results and its leftover is refused by the touch guard instead of being handed to a by-name card. A
// Pack or Unpack cleanup (a consent captured before the task) never deletes by name: it keeps the
// refusal of a root it cannot identify, and nothing is deleted.
[[nodiscard]] bool RunDeleteIdentityLessOpenFailureAndArchiveOriginChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Identity-less open failure and archive cleanup: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");

    const std::filesystem::path root = state.tempRoot / L"identity-less-open-archive";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot                = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path location = uncStatus == LoopbackShareStatus::Reachable ? LoopbackShareAlias(root) : root;
    IdentityLessCaseScope scope(state, operations);
    if (! IdentityLessCaseScope::MakeIdentityLess(location))
        return reject(L"the identity-less injection could not be configured");

    // 1. The delete bind's open fails with ERROR_NOT_SUPPORTED: refused at Preparing as a failed open.
    const std::filesystem::path refused = root / L"open-refused.bin";
    if (! WriteFilledTestFile(refused, 4096u, 0x61u) ||
        SetEnvironmentVariableW(kIdentityLessBindOpenNotSupportedEnv.data(), (location / refused.filename()).c_str()) == FALSE)
        return reject(L"open-failure fixture unavailable");
    const auto refusedTask = scope.StartDelete({location / refused.filename()}, false);
    if (! refusedTask.has_value())
        return reject(L"the open-failure Delete was not admitted");
    const auto refusedQuestion = scope.WaitForQuestion(refusedTask.value(), L"open-failure Delete");
    if (refusedQuestion.has_value())
    {
        static_cast<void>(scope.CancelQuestionsUntilEnd(refusedTask.value(), L"open-failure Delete question"));
        return reject(std::format(L"an open that failed with ERROR_NOT_SUPPORTED was offered a question: {}", DescribeIdentityLessPrompt(refusedQuestion)));
    }
    const CompletedTaskInfo* refusedEnd = scope.WaitForEnd(refusedTask.value(), L"open-failure Delete");
    const auto refusedSummary           = FindIdentityLessSummary(operations, refusedTask.value());
    const std::wstring openFailedText   = std::format(L"openHr=0x{:08X}", static_cast<unsigned long>(HRESULT_FROM_WIN32(ERROR_OPEN_FAILED)));
    const bool refusedLogged            = refusedSummary.has_value() && std::ranges::any_of(refusedSummary->issueDiagnostics,
                                                                                            [&](const IdentityLessDiagnostic& issue) noexcept
    {
        return issue.category == L"preparation.deleteRoot" && issue.message.find(openFailedText) != std::wstring::npos &&
               issue.message.find(L"identityHr=not queried") != std::wstring::npos;
    });
    if (refusedEnd == nullptr || refusedEnd->hr != HRESULT_FROM_WIN32(ERROR_OPEN_FAILED) || ! refusedLogged || ! FileSizeEquals(refused, 4096u))
        return reject(std::format(L"an open that failed with ERROR_NOT_SUPPORTED was not refused as a failed open: hr=0x{:08X} logged={} exists={} "
                                  L"issues=[{}]",
                                  refusedEnd != nullptr ? static_cast<unsigned long>(refusedEnd->hr) : 0ul,
                                  refusedLogged,
                                  std::filesystem::exists(refused),
                                  refusedSummary.has_value() ? DescribeIdentityLessIssues(refusedSummary.value()) : std::wstring(L"<no summary>")));

    // 2. A leftover whose open fails the same way is refused by the touch guard, never handed off.
    const std::filesystem::path leftover = root / L"movie.mkv.rs_tmp_00112233445566778899AABBCCDDEEFF_0";
    if (! WriteFilledTestFile(leftover, 4096u, 0x62u) ||
        SetEnvironmentVariableW(kIdentityLessBindOpenNotSupportedEnv.data(), (location / leftover.filename()).c_str()) == FALSE)
        return reject(L"open-failure leftover fixture unavailable");
    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    uint64_t leftoverTaskId  = 0u;
    const HRESULT leftoverHr = operations->AdmitOperation(FILESYSTEM_DELETE,
                                                          FolderWindow::Pane::Left,
                                                          std::nullopt,
                                                          state.fsLocal,
                                                          {location / leftover.filename()},
                                                          {},
                                                          FILESYSTEM_FLAG_RECURSIVE,
                                                          false,
                                                          0u,
                                                          FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                          false,
                                                          nullptr,
                                                          &leftoverTaskId);
    HostClearTestPromptResultOverride();
    if (leftoverTaskId != 0u)
        scope.Track(leftoverTaskId);
    HostPromptDebugSnapshot leftoverPrompt{};
    const bool leftoverPromptSeen = HostGetTestLastPromptDebugSnapshot(leftoverPrompt);
    if (leftoverHr != FolderWindow::kFileOperationRefusedAlreadyReported || leftoverTaskId != 0u || HostGetTestPromptRequestCount() != 1u ||
        ! leftoverPromptSeen || leftoverPrompt.presentation != HOST_PROMPT_PRESENTATION_ARTIFACT_BLOCKED || ! FileSizeEquals(leftover, 4096u))
        return reject(std::format(L"a leftover whose open failed with ERROR_NOT_SUPPORTED was handed off instead of refused: hr=0x{:08X} task={} "
                                  L"prompts={} presentation={} exists={}",
                                  static_cast<unsigned long>(leftoverHr),
                                  leftoverTaskId,
                                  HostGetTestPromptRequestCount(),
                                  static_cast<unsigned int>(leftoverPrompt.presentation),
                                  std::filesystem::exists(leftover)));
    static_cast<void>(SetEnvironmentVariableW(kIdentityLessBindOpenNotSupportedEnv.data(), nullptr));

    // 3. Pack and Unpack cleanups carry a consent captured before the task: never by name.
    for (const FileOperations::DeleteOrigin origin : {FileOperations::DeleteOrigin::PackCleanup, FileOperations::DeleteOrigin::UnpackCleanup})
    {
        const bool pack                    = origin == FileOperations::DeleteOrigin::PackCleanup;
        const std::filesystem::path victim = root / (pack ? L"packed.bin" : L"unpacked.bin");
        if (! WriteFilledTestFile(victim, 4096u, pack ? 0x63u : 0x64u))
            return reject(L"archive cleanup fixture unavailable");
        uint64_t cleanupTaskId  = 0u;
        const HRESULT cleanupHr = operations->AdmitOperation(FILESYSTEM_DELETE,
                                                             FolderWindow::Pane::Left,
                                                             std::nullopt,
                                                             state.fsLocal,
                                                             {location / victim.filename()},
                                                             {},
                                                             FILESYSTEM_FLAG_RECURSIVE,
                                                             false,
                                                             0u,
                                                             FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                             false,
                                                             nullptr,
                                                             &cleanupTaskId,
                                                             {},
                                                             {},
                                                             {},
                                                             {},
                                                             std::nullopt,
                                                             origin,
                                                             FileOperations::ConsentKind::ArchiveDeleteAfter);
        if (FAILED(cleanupHr) || cleanupTaskId == 0u)
            return reject(std::format(L"the {} cleanup was not admitted (hr=0x{:08X})", pack ? L"Pack" : L"Unpack", static_cast<unsigned long>(cleanupHr)));
        scope.Track(cleanupTaskId);
        const auto cleanupQuestion = scope.WaitForQuestion(cleanupTaskId, L"archive cleanup Delete");
        if (cleanupQuestion.has_value())
            static_cast<void>(scope.CancelQuestionsUntilEnd(cleanupTaskId, L"archive cleanup Delete question"));
        const CompletedTaskInfo* cleanupEnd = scope.WaitForEnd(cleanupTaskId, L"archive cleanup Delete");
        const auto cleanupSummary           = FindIdentityLessSummary(operations, cleanupTaskId);
        const bool byNameNoted =
            cleanupSummary.has_value() && CountIdentityLessNotes(cleanupSummary.value(), L"delete.permanent.byName", IDS_FILEOPS_NOTE_DELETED_BY_NAME) != 0u;
        if (cleanupQuestion.has_value() || cleanupEnd == nullptr || cleanupEnd->hr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || byNameNoted ||
            ! FileSizeEquals(victim, 4096u))
            return reject(std::format(L"the {} cleanup of an item the location cannot identify was not refused: question={} hr=0x{:08X} byName={} "
                                      L"exists={} issues=[{}]",
                                      pack ? L"Pack" : L"Unpack",
                                      DescribeIdentityLessPrompt(cleanupQuestion),
                                      cleanupEnd != nullptr ? static_cast<unsigned long>(cleanupEnd->hr) : 0ul,
                                      byNameNoted,
                                      std::filesystem::exists(victim),
                                      cleanupSummary.has_value() ? DescribeIdentityLessIssues(cleanupSummary.value()) : std::wstring(L"<no summary>")));
    }
    Debug::Perf::Emit(
        L"FileOps.SelfTest.IdentityLessOpenFailureAndArchiveOrigins", L"open=refused;leftover=blocked;pack=refused;unpack=refused", 0u, 4u, 0u, S_OK);
    return true;
}
