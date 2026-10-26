// I35 S5 (Specs/Plans/Done/FileOperations_CopyMoveUxAndReliabilityReview_2026-10-11.md, section 8): the
// decision mirror. A published decision is rendered on its task card and mirrored over the pane that shows
// its folder; one engine decision, two renderers, the first accepted answer wins. These cases drive real
// tasks through the production prompt route and answer them on the mirror, on the card, by Escape and by
// Cancel all. They stand in the keyboard focus, the foreground window and the key message time (no
// activation, no real focus change, no taskbar flash).

namespace
{
using MirrorTask                = FolderWindow::FileOperationState::Task;
using MirrorAction              = MirrorTask::ConflictAction;
using MirrorBucket              = MirrorTask::ConflictBucket;
using MirrorSnapshot            = FolderWindow::FileOperationState::DecisionMirrorDebugSnapshot;
using MirrorSummary             = FolderWindow::FileOperationState::CompletedTaskSummary;
using MirrorOverlay             = RedSalamander::Ui::AlertOverlayModelessDebugSnapshot;
constexpr auto kMirrorWait      = std::chrono::seconds(15);
constexpr LONG kMirrorKeyLateMs = static_cast<LONG>(RedSalamander::Ui::kModelessKeyGuardTypeaheadMs) + 100;

[[nodiscard]] std::optional<MirrorSummary> FindMirrorCompletedSummary(FolderWindow::FileOperationState* operations, uint64_t taskId)
{
    std::vector<MirrorSummary> summaries;
    operations->CollectCompletedTasks(summaries);
    const auto found = std::ranges::find_if(summaries, [taskId](const MirrorSummary& summary) noexcept { return summary.taskId == taskId; });
    if (found == summaries.end())
        return std::nullopt;
    return *found;
}

[[nodiscard]] MirrorSnapshot MirrorState(FolderWindow::FileOperationState* operations) noexcept
{
    MirrorSnapshot snapshot{};
    static_cast<void>(operations->DebugGetDecisionMirrorSnapshot(snapshot));
    return snapshot;
}

[[nodiscard]] MirrorOverlay OverlayState(HWND overlay) noexcept
{
    MirrorOverlay snapshot{};
    static_cast<void>(RedSalamander::Ui::DebugGetAlertOverlayModelessSnapshot(overlay, snapshot));
    return snapshot;
}

[[nodiscard]] std::wstring StatusDecisionText(FolderWindow* folder, FolderWindow::Pane pane)
{
    FolderWindow::FolderWindowPaneStatusBarDebugSnapshot snapshot{};
    if (! folder->DebugGetPaneStatusBarSnapshot(pane, snapshot))
        return {};
    return snapshot.decisionText;
}

[[nodiscard]] bool FileExistsForMirror(const std::filesystem::path& path) noexcept
{
    return GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES;
}

[[nodiscard]] bool MirrorHasButton(const MirrorOverlay& overlay, uint32_t id) noexcept
{
    return std::ranges::find(overlay.buttonIds, id) != overlay.buttonIds.end();
}

[[nodiscard]] bool MirrorLocalPanes(FolderWindow* folder) noexcept
{
    return (folder->GetFileSystemPluginId(FolderWindow::Pane::Left) == kPluginIdLocal ||
            SUCCEEDED(folder->SetFileSystemPluginForPane(FolderWindow::Pane::Left, kPluginIdLocal))) &&
           (folder->GetFileSystemPluginId(FolderWindow::Pane::Right) == kPluginIdLocal ||
            SUCCEEDED(folder->SetFileSystemPluginForPane(FolderWindow::Pane::Right, kPluginIdLocal)));
}

// A key the user presses `afterArmMs` after the mirror's key guard was armed (it took the focus, or the focus
// came back to it; negative: typed before it). The key guard reads the message time, which a sent message
// does not carry, so the test stands it in.
void PressMirrorKey(HWND overlay, WPARAM vk, LONG afterArmMs, bool repeat = false) noexcept
{
    const DWORD armedTick = OverlayState(overlay).guardArmedTick;
    RedSalamander::Ui::DebugSetAlertOverlayKeyMessageTimeOverride(armedTick + static_cast<DWORD>(afterArmMs));
    static_cast<void>(SendMessageW(overlay, WM_KEYDOWN, vk, repeat ? (static_cast<LPARAM>(1) << 30) : 0));
    RedSalamander::Ui::DebugSetAlertOverlayKeyMessageTimeOverride(std::nullopt);
}

// Sets or clears a key in the calling thread's keyboard state (a modifier held for a chord).
[[nodiscard]] bool SetMirrorKeyDown(int virtualKey, bool down) noexcept
{
    std::array<BYTE, 256> keyboard{};
    if (GetKeyboardState(keyboard.data()) == FALSE)
        return false;
    BYTE& state = keyboard[static_cast<size_t>(virtualKey)];
    state       = down ? static_cast<BYTE>(state | 0x80u) : static_cast<BYTE>(state & 0x7Fu);
    return SetKeyboardState(keyboard.data()) != FALSE;
}

// Owns the settings, prompt, focus, foreground and pane state a mirror case changes, and every task it
// starts. By default nothing has the focus and no window is the foreground window, so the mirror never
// takes the focus or flashes unless a step says so.
class DecisionMirrorCaseScope final
{
public:
    DecisionMirrorCaseScope(SelfTestState& state, FolderWindow* folder, FolderWindow::FileOperationState* operations) noexcept
        : _state(state),
          _folder(folder),
          _operations(operations)
    {
        _mirrorEnabled                                                  = EnsureFileOperationsSettingsForSelfTest().mirrorDecisionsInPane;
        _autoDismiss                                                    = operations->GetAutoDismissSuccess();
        _autoAccept                                                     = HostGetAutoAcceptPrompts();
        _promptOverride                                                 = HostGetTestPromptResultOverride();
        _originalPopup                                                  = operations->GetPopupHwndForSelfTest();
        _originalLeft                                                   = folder->GetCurrentPath(FolderWindow::Pane::Left);
        _originalRight                                                  = folder->GetCurrentPath(FolderWindow::Pane::Right);
        EnsureFileOperationsSettingsForSelfTest().mirrorDecisionsInPane = true;
        operations->SetAutoDismissSuccess(false);
        // A real decision on the card: neither the run's auto-accept nor an override answers it.
        HostSetAutoAcceptPrompts(false);
        HostClearTestPromptResultOverride();
        RedSalamander::Ui::DebugSetAlertOverlayFocusOverride(std::optional<HWND>(nullptr));
        operations->DebugSetDecisionMirrorForegroundOverride(std::optional<HWND>(nullptr));
    }
    DecisionMirrorCaseScope(const DecisionMirrorCaseScope&)            = delete;
    DecisionMirrorCaseScope& operator=(const DecisionMirrorCaseScope&) = delete;
    ~DecisionMirrorCaseScope()
    {
        if (const HWND root = GetAncestor(_state.mainWindow, GA_ROOT); root && IsWindowEnabled(root) == FALSE)
            EnableWindow(root, TRUE);
        _operations->DebugSetDecisionMirrorReconcilePaused(false);
        _operations->DebugFailNextDecisionMirrorShows(0u);
        static_cast<void>(SetMirrorKeyDown(VK_CONTROL, false));
        const bool quiet = CancelOutstandingSelfTestOperations(_operations);
        for (const uint64_t taskId : _tasks)
        {
            if (quiet)
            {
                _operations->DismissCompletedTask(taskId);
                _operations->DebugRemoveDiagnosticsForTask(taskId);
                _state.completedTasks.erase(taskId);
            }
        }
        RedSalamander::Ui::DebugSetAlertOverlayFocusOverride(std::nullopt);
        RedSalamander::Ui::DebugSetAlertOverlayKeyMessageTimeOverride(std::nullopt);
        _operations->DebugSetDecisionMirrorForegroundOverride(std::nullopt);
        EnsureFileOperationsSettingsForSelfTest().mirrorDecisionsInPane = _mirrorEnabled;
        _operations->SetAutoDismissSuccess(_autoDismiss);
        static_cast<void>(_operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
        HostSetAutoAcceptPrompts(_autoAccept);
        if (_promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(_promptOverride);
        if (_originalLeft.has_value())
            _folder->SetFolderPath(FolderWindow::Pane::Left, _originalLeft.value());
        if (_originalRight.has_value())
            _folder->SetFolderPath(FolderWindow::Pane::Right, _originalRight.value());
        if (quiet && ! _originalPopup && _operations->GetPopupHwndForSelfTest())
            _operations->DebugResetPopupForSelfTest();
        if (! quiet)
            _state.retainTempRootOnFailure = true;
    }

    void Track(uint64_t taskId)
    {
        _tasks.push_back(taskId);
    }

    // A permanent Delete asks its confirmation on the card before any change (C1).
    [[nodiscard]] uint64_t StartPermanentDelete(FolderWindow::FileOperationState* operations, const std::filesystem::path& item)
    {
        const auto admitted = StartFileOperationAndGetId(
            operations, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, _state.fsLocal, {item}, {}, FILESYSTEM_FLAG_RECURSIVE, false);
        if (! admitted.has_value())
            return 0u;
        if (operations == _operations)
            Track(admitted.value());
        return admitted.value();
    }

    [[nodiscard]] uint64_t StartPermanentDelete(const std::filesystem::path& item)
    {
        return StartPermanentDelete(_operations, item);
    }

    [[nodiscard]] uint64_t StartCopy(std::vector<std::filesystem::path> items, const std::filesystem::path& destination)
    {
        const auto admitted = StartFileOperationAndGetId(_operations,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         _state.fsLocal,
                                                         std::move(items),
                                                         destination,
                                                         FILESYSTEM_FLAG_NONE,
                                                         false,
                                                         0u,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         _state.fsLocal);
        if (! admitted.has_value())
            return 0u;
        Track(admitted.value());
        return admitted.value();
    }

    [[nodiscard]] bool ShowPane(FolderWindow::Pane pane, const std::filesystem::path& folder, std::wstring_view itemName) noexcept
    {
        _folder->SetFolderPath(pane, folder);
        return RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept
        { return CurrentPanePathEquals(_folder, pane, folder) && (itemName.empty() || _folder->DebugHasItemDisplayName(pane, itemName)); },
                                                             {.timeout = SelfTest::Scale(kMirrorWait), .operationName = L"decision mirror pane listing"})
            .conditionMet;
    }

    // The task's actionable published prompt (optionally of one class, optionally newer than a generation).
    [[nodiscard]] std::optional<MirrorTask::ConflictPromptState> WaitPrompt(FolderWindow::FileOperationState* operations,
                                                                            uint64_t taskId,
                                                                            std::optional<MirrorBucket> bucket,
                                                                            uint64_t afterGeneration              = 0u,
                                                                            const std::function<void()>& eachPump = {}) noexcept
    {
        std::optional<MirrorTask::ConflictPromptState> prompt;
        const auto asked = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            if (eachPump)
                eachPump();
            prompt = TryGetConflictPromptCopy(operations->FindTask(taskId));
            return prompt.has_value() && prompt->buttonsPublishable && prompt->generation > afterGeneration;
        },
            {.timeout = SelfTest::Scale(kMirrorWait), .operationName = L"decision mirror prompt"});
        if (! asked.conditionMet || ! prompt.has_value() || (bucket.has_value() && prompt->bucket != bucket.value()))
            return std::nullopt;
        return prompt;
    }

    [[nodiscard]] std::optional<MirrorTask::ConflictPromptState> WaitPrompt(uint64_t taskId, MirrorBucket bucket) noexcept
    {
        return WaitPrompt(_operations, taskId, bucket);
    }

    [[nodiscard]] bool WaitMirror(FolderWindow::FileOperationState* operations,
                                  const std::function<bool(const MirrorSnapshot&)>& condition,
                                  std::wstring_view name) noexcept
    {
        return RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept { return condition(MirrorState(operations)); },
                                                             {.timeout = SelfTest::Scale(kMirrorWait), .operationName = name})
            .conditionMet;
    }

    [[nodiscard]] bool WaitMirror(const std::function<bool(const MirrorSnapshot&)>& condition, std::wstring_view name) noexcept
    {
        return WaitMirror(_operations, condition, name);
    }

    // Pumps through at least two more presentation passes (each one reconciles the mirror).
    void Settle() noexcept
    {
        const uint64_t start = MirrorState(_operations).reconcileCount;
        static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return MirrorState(_operations).reconcileCount >= start + 2u;
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"decision mirror settle"}));
    }

    [[nodiscard]] std::optional<MirrorSummary> WaitCompleted(FolderWindow::FileOperationState* operations, uint64_t taskId) noexcept
    {
        const auto done = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return _state.completedTasks.contains(taskId) && operations->FindTask(taskId) == nullptr;
        }, {.timeout = SelfTest::Scale(kMirrorWait), .operationName = L"decision mirror task completion"});
        return done.conditionMet ? FindMirrorCompletedSummary(operations, taskId) : std::nullopt;
    }

    [[nodiscard]] std::optional<MirrorSummary> WaitCompleted(uint64_t taskId) noexcept
    {
        return WaitCompleted(_operations, taskId);
    }

    [[nodiscard]] bool WaitIndicator(std::wstring_view expected, std::wstring_view name) noexcept
    {
        return RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept
        { return StatusDecisionText(_folder, FolderWindow::Pane::Left) == expected && StatusDecisionText(_folder, FolderWindow::Pane::Right) == expected; },
                                                             {.timeout = SelfTest::Scale(kMirrorWait), .operationName = name})
            .conditionMet;
    }

    // R8.12: who announced the prompt (true: the mirror, false: the card).
    [[nodiscard]] bool WaitAnnouncer(uint64_t taskId, uint64_t generation, bool byMirror) noexcept
    {
        return RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept
        { return _operations->DebugGetPromptAnnouncer(taskId, generation) == std::optional<bool>{byMirror}; },
                                                             {.timeout = SelfTest::Scale(kMirrorWait), .operationName = L"decision mirror announcement"})
            .conditionMet;
    }

private:
    SelfTestState& _state;
    FolderWindow* _folder                         = nullptr;
    FolderWindow::FileOperationState* _operations = nullptr;
    std::vector<uint64_t> _tasks;
    std::optional<std::filesystem::path> _originalLeft;
    std::optional<std::filesystem::path> _originalRight;
    HWND _originalPopup              = nullptr;
    bool _mirrorEnabled              = true;
    bool _autoDismiss                = false;
    bool _autoAccept                 = false;
    HostPromptResult _promptOverride = HOST_PROMPT_RESULT_NONE;
};

[[nodiscard]] bool SubmitOnCard(FolderWindow::FileOperationState* operations, uint64_t taskId, uint64_t generation, MirrorAction action) noexcept
{
    MirrorTask* task = operations->FindTask(taskId);
    return task != nullptr && task->SubmitConflictDecision(generation, action, std::nullopt, FileOperations::ConflictActionPlacement::Primary);
}

[[nodiscard]] bool MirrorCaseReady(SelfTestState& state,
                                   FolderWindow*& folder,
                                   FolderWindow::FileOperationState*& operations,
                                   RedSalamander::TestSupport::ScopedWindowActivationBlocker& blocker) noexcept
{
    folder     = TryGetFolderWindow(state.mainWindow);
    operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return false;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return false;
    return MirrorLocalPanes(folder);
}
} // namespace

// R8.1-R8.9, R8.12, R8.13 and the key guard on a permanent Delete's card confirmation (a task-wide question
// before any change).
[[nodiscard]] bool RunFileOpsDecisionMirrorChecks(SelfTestState& state) noexcept
{
    using FileOps     = FolderWindow::FileOperationState;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Decision mirror: {}", reason));
        return false;
    };
    FolderWindow* folder = nullptr;
    FileOps* operations  = nullptr;
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! MirrorCaseReady(state, folder, operations, blocker))
        return reject(L"owner, Local panes, the no-activation guard or the initial worker quiet point unavailable");

    const std::filesystem::path root = state.tempRoot / L"decision-mirror";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot   = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const auto sourceFolder = root / L"source";
    const auto otherFolder  = root / L"elsewhere";
    if (! RecreateEmptyDirectory(sourceFolder) || ! RecreateEmptyDirectory(otherFolder) || ! WriteTestFile(otherFolder / L"other.bin", 64u))
        return reject(L"fixture folders unavailable");
    DecisionMirrorCaseScope scope(state, folder, operations);
    const HWND leftView  = folder->GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND rightView = folder->GetFolderViewHwnd(FolderWindow::Pane::Right);
    const HWND ownerRoot = GetAncestor(state.mainWindow, GA_ROOT);
    if (! leftView || IsWindowVisible(leftView) == FALSE || ! rightView || ! ownerRoot)
        return reject(L"the pane views are not available, so no pane can host the mirror");
    const uint32_t cancelId        = FileOps::DebugDecisionMirrorPrimaryButtonId(MirrorAction::Cancel);
    const uint32_t deleteId        = FileOps::DebugDecisionMirrorPrimaryButtonId(MirrorAction::PermanentDelete);
    const std::wstring deleteTitle = LoadStringResource(nullptr, IDS_FILEOP_OPERATION_DELETE);
    const std::wstring pendingOne  = FormatStringResource(nullptr, IDS_FMT_FILEOPS_DECISION_PENDING, 1ul);
    const HRESULT canceledHr       = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    const auto startDelete         = [&](std::wstring_view name, uint64_t& taskId) -> std::optional<MirrorTask::ConflictPromptState>
    {
        const auto item = sourceFolder / name;
        taskId          = 0u;
        if (! WriteTestFile(item, 2048u) || ! scope.ShowPane(FolderWindow::Pane::Left, sourceFolder, name))
            return std::nullopt;
        taskId = scope.StartPermanentDelete(item);
        return taskId != 0u ? scope.WaitPrompt(taskId, MirrorBucket::PermanentDeleteConfirmation) : std::nullopt;
    };
    const auto shownFor = [&](uint64_t taskId, std::wstring_view what) noexcept
    { return scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return value.visible && value.taskId == taskId; }, what); };
    const auto canceledOnCard = [&](uint64_t taskId, uint64_t generation) noexcept
    {
        const auto done = SubmitOnCard(operations, taskId, generation, MirrorAction::Cancel) ? scope.WaitCompleted(taskId) : std::nullopt;
        return done.has_value() && done->resultHr == canceledHr && done->warningCount == 0u;
    };

    // 1. The mirror appears over the pane that shows the folder and renders the card's question with the
    //    published actions in the published order (Cancel the default, Escape bound to the prompt's own
    //    escape action); both status bars show one pending decision. It did not take the focus, so the
    //    card announces the prompt (R8.12). An answer on the mirror decides once. The card's synchronization
    //    runs first here (the mirror's reconcile is held back): it defers the announcement, announces nothing,
    //    and announces once the reconcile decided.
    if (! scope.ShowPane(FolderWindow::Pane::Right, otherFolder, L"other.bin"))
        return reject(L"right pane listing unavailable");
    operations->DebugSetDecisionMirrorReconcilePaused(true);
    const uint64_t defersBefore = MirrorState(operations).cardDeferCount;
    uint64_t firstTask          = 0u;
    const auto firstPrompt      = startDelete(L"first.bin", firstTask);
    if (! firstPrompt.has_value())
    {
        operations->DebugSetDecisionMirrorReconcilePaused(false);
        return reject(L"the permanent Delete did not publish its card confirmation");
    }
    const bool deferred =
        scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return value.cardDeferCount > defersBefore; }, L"decision mirror card defers");
    const std::optional<bool> announcerWhileDeferred = operations->DebugGetPromptAnnouncer(firstTask, firstPrompt->generation);
    const bool mirrorShownWhileDeferred              = MirrorState(operations).visible;
    operations->DebugSetDecisionMirrorReconcilePaused(false);
    if (! deferred || announcerWhileDeferred.has_value() || mirrorShownWhileDeferred)
        return reject(std::format(L"the card synchronized before the mirror's reconcile did not defer the announcement (deferred={} announced={} shown={})",
                                  deferred,
                                  announcerWhileDeferred.has_value(),
                                  mirrorShownWhileDeferred));
    if (! shownFor(firstTask, L"decision mirror shown"))
        return reject(L"the mirror did not appear over the pane that shows the deleted item's folder");
    MirrorSnapshot mirror = MirrorState(operations);
    MirrorOverlay overlay = OverlayState(mirror.overlay);
    const std::wstring expectedQuestion =
        FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFIRM_PERMANENT_DELETE, firstPrompt->consentDetail, firstPrompt->sourcePath);
    if (mirror.pane != FolderWindow::Pane::Left || mirror.anchor != leftView || overlay.anchor != leftView || mirror.generation != firstPrompt->generation ||
        ! mirror.windowVisible || overlay.title != deleteTitle || overlay.rootUiaName != deleteTitle ||
        overlay.message.find(expectedQuestion) == std::wstring::npos || overlay.buttonIds.size() != 2u || overlay.buttonIds[0] != cancelId ||
        overlay.buttonIds[1] != deleteId || overlay.primaryButtonId != std::optional<uint32_t>{cancelId} ||
        overlay.focusedButtonId != std::optional<uint32_t>{cancelId} ||
        overlay.escapeButtonId != std::optional<uint32_t>{FileOps::DebugDecisionMirrorEscapeButtonId()} || ! overlay.keysRequireFocus || mirror.tookFocus ||
        mirror.announced || overlay.keysGuarded)
        return reject(std::format(L"the mirror did not render the card's decision: pane={} anchor={} title='{}' uia='{}' buttons={} default={} "
                                  L"focused={} escape={} keysRequireFocus={} tookFocus={} announced={} guarded={}",
                                  static_cast<int>(mirror.pane),
                                  mirror.anchor == leftView,
                                  overlay.title,
                                  overlay.rootUiaName,
                                  overlay.buttonIds.size(),
                                  overlay.primaryButtonId.value_or(0u),
                                  overlay.focusedButtonId.value_or(0u),
                                  overlay.escapeButtonId.value_or(0u),
                                  overlay.keysRequireFocus,
                                  mirror.tookFocus,
                                  mirror.announced,
                                  overlay.keysGuarded));
    if (! scope.WaitAnnouncer(firstTask, firstPrompt->generation, false))
        return reject(L"the card did not announce a decision the mirror showed without the focus (R8.12)");
    if (mirror.pendingCount != 1u || ! scope.WaitIndicator(pendingOne, L"decision indicator shown"))
        return reject(std::format(L"the status bars do not show one pending decision (pending={}, left='{}', right='{}')",
                                  mirror.pendingCount,
                                  StatusDecisionText(folder, FolderWindow::Pane::Left),
                                  StatusDecisionText(folder, FolderWindow::Pane::Right)));
    FolderWindow::FolderWindowPaneStatusBarDebugSnapshot leftBar{};
    if (! folder->DebugGetPaneStatusBarSnapshot(FolderWindow::Pane::Left, leftBar) || leftBar.windowText != pendingOne ||
        (leftBar.visible && leftBar.decisionPartWidthPx <= 0))
        return reject(L"the status bar's UI Automation name or painted segment does not carry the pending decision");
    const uint64_t acceptedBefore = mirror.acceptedCount;
    if (! RedSalamander::Ui::DebugInvokeAlertOverlayButton(mirror.overlay, deleteId))
        return reject(L"the mirror's Permanently delete button could not be pressed");
    mirror = MirrorState(operations);
    if (mirror.visible || mirror.acceptedCount != acceptedBefore + 1u)
        return reject(L"an answer on the mirror was not accepted, or the mirror stayed on screen after it");
    if (SubmitOnCard(operations, firstTask, firstPrompt->generation, MirrorAction::Cancel))
        return reject(L"the card's answer to a prompt the mirror already decided was accepted");
    const auto firstDone = scope.WaitCompleted(firstTask);
    if (! firstDone.has_value() || firstDone->resultHr != S_OK || firstDone->warningCount != 0u || FileExistsForMirror(sourceFolder / L"first.bin"))
        return reject(L"the decision taken on the mirror did not run the Delete");
    if (! scope.WaitIndicator({}, L"decision indicator cleared"))
        return reject(L"the pending-decision segment stayed after the last decision was answered");

    // 2. Answered on the card: the mirror disappears and the task follows the card's answer.
    uint64_t secondTask     = 0u;
    const auto secondPrompt = startDelete(L"second.bin", secondTask);
    if (! secondPrompt.has_value() || ! shownFor(secondTask, L"decision mirror second"))
        return reject(L"the second Delete's decision did not reach the mirror");
    if (! SubmitOnCard(operations, secondTask, secondPrompt->generation, MirrorAction::Cancel))
        return reject(L"the card's Cancel was not accepted");
    if (! scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return ! value.visible; }, L"decision mirror hides after the card's answer"))
        return reject(L"the mirror stayed after the card answered its decision");
    const auto secondDone = scope.WaitCompleted(secondTask);
    if (! secondDone.has_value() || secondDone->resultHr != canceledHr || secondDone->warningCount != 0u || ! FileExistsForMirror(sourceFolder / L"second.bin"))
        return reject(L"the card's Cancel did not end the Delete as canceled with the file kept");

    // 3. The pane navigates away: the mirror hides and the card keeps the decision; back again, it returns.
    //    Closing the popup leaves the status-bar segment, whose click opens the popup again. Cancel all while
    //    the mirror is shown hides it and ends the task canceled with no warning.
    uint64_t thirdTask     = 0u;
    const auto thirdPrompt = startDelete(L"third.bin", thirdTask);
    if (! thirdPrompt.has_value() || ! shownFor(thirdTask, L"decision mirror third"))
        return reject(L"the third Delete's decision did not reach the mirror");
    if (! scope.ShowPane(FolderWindow::Pane::Left, otherFolder, L"other.bin") ||
        ! scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return ! value.visible; }, L"decision mirror hides on navigation"))
        return reject(L"the mirror stayed over a pane that no longer shows the decision's folder");
    if (! TryGetConflictPromptCopy(operations->FindTask(thirdTask)).has_value())
        return reject(L"navigating the pane away resolved the card's decision");
    if (! scope.ShowPane(FolderWindow::Pane::Left, sourceFolder, L"third.bin") || ! shownFor(thirdTask, L"decision mirror returns"))
        return reject(L"the mirror did not return when the pane showed the folder again");
    const auto popupShown = [&]() noexcept
    {
        const HWND popup = operations->GetPopupHwndForSelfTest();
        return popup != nullptr && IsWindowVisible(popup) != FALSE;
    };
    if (! RedSalamander::TestSupport::PumpMessagesUntil(popupShown, {.timeout = SelfTest::Scale(kMirrorWait), .operationName = L"decision popup revealed"})
              .conditionMet)
        return reject(L"the pending decision did not reveal the File Operations popup");
    SendMessageW(operations->GetPopupHwndForSelfTest(), WM_CLOSE, 0, 0);
    scope.Settle();
    if (popupShown() || StatusDecisionText(folder, FolderWindow::Pane::Left) != pendingOne || ! MirrorState(operations).visible)
        return reject(L"closing the popup with a decision pending hid the decision from the main window");
    if (FolderWindow::FolderWindowPaneStatusBarDebugSnapshot bar{}; folder->DebugGetPaneStatusBarSnapshot(FolderWindow::Pane::Left, bar) && bar.visible)
    {
        if (! folder->DebugClickPaneStatusBarDecision(FolderWindow::Pane::Left) ||
            ! RedSalamander::TestSupport::PumpMessagesUntil(popupShown,
                                                            {.timeout = SelfTest::Scale(kMirrorWait), .operationName = L"decision indicator opens the popup"})
                  .conditionMet)
            return reject(L"a click on the pending-decision segment did not open the File Operations popup");
    }
    else
    {
        AppendLog(L"Decision mirror: the left status bar is hidden; the segment click was not exercised.");
    }
    if (! MirrorState(operations).visible)
        return reject(L"the mirror was not shown when Cancel all was requested");
    operations->CancelAll();
    if (! scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return ! value.visible; }, L"decision mirror hides on Cancel all"))
        return reject(L"the mirror stayed on screen after Cancel all");
    const auto thirdDone = scope.WaitCompleted(thirdTask);
    if (! thirdDone.has_value() || thirdDone->resultHr != canceledHr || thirdDone->warningCount != 0u || thirdDone->errorCount != 0u ||
        ! FileExistsForMirror(sourceFolder / L"third.bin"))
        return reject(L"Cancel all with the mirror shown did not end the task canceled with no warning and the file kept");
    if (! scope.WaitIndicator({}, L"decision indicator after Cancel all"))
        return reject(L"the pending-decision segment stayed after Cancel all");

    // 3b. Closing the mirror only hides it (never a decision), and that prompt is not mirrored again.
    uint64_t closeTask     = 0u;
    const auto closePrompt = startDelete(L"close.bin", closeTask);
    if (! closePrompt.has_value() || ! shownFor(closeTask, L"decision mirror to close"))
        return reject(L"the Delete to close on the mirror did not reach it");
    if (! RedSalamander::Ui::DebugInvokeAlertOverlayClose(MirrorState(operations).overlay))
        return reject(L"the mirror's close affordance could not be used");
    if (MirrorState(operations).visible || ! TryGetConflictPromptCopy(operations->FindTask(closeTask)).has_value())
        return reject(L"closing the mirror decided the prompt or left the mirror on screen");
    scope.Settle();
    if (MirrorState(operations).visible || StatusDecisionText(folder, FolderWindow::Pane::Left) != pendingOne)
        return reject(L"a closed mirror came back for the same prompt, or the pending segment disappeared");
    if (! canceledOnCard(closeTask, closePrompt->generation))
        return reject(L"the closed mirror's prompt could not be answered on the card");
    if (! scope.WaitIndicator({}, L"decision indicator before the focused mirror"))
        return reject(L"the pending-decision segment stayed after the closed prompt was answered");

    // 4. Shown in the focused pane of the foreground window: the mirror takes the focus, announces the card's
    //    text itself (R8.12) and flashes the popup's taskbar button once (the foreground then is the mirror,
    //    a window the owner owns). Tab and arrows work immediately without a click or Ctrl+Tab. Merely
    //    moving focus never answers, repeat is ignored, and Escape gives focus back without deleting.
    operations->DebugEnsurePopupVisibleForSelfTest();
    RedSalamander::Ui::DebugSetAlertOverlayFocusOverride(leftView);
    operations->DebugSetDecisionMirrorForegroundOverride(ownerRoot);
    const uint64_t flashBefore = MirrorState(operations).flashCount;
    uint64_t fourthTask        = 0u;
    const auto fourthPrompt    = startDelete(L"fourth.bin", fourthTask);
    if (! fourthPrompt.has_value() || ! shownFor(fourthTask, L"decision mirror focused"))
        return reject(L"the fourth Delete's decision did not reach the mirror");
    scope.Settle();
    mirror  = MirrorState(operations);
    overlay = OverlayState(mirror.overlay);
    FileOperations::ConflictPromptTextFacts facts{};
    facts.operation                         = FILESYSTEM_DELETE;
    facts.bucket                            = static_cast<uint8_t>(MirrorBucket::PermanentDeleteConfirmation);
    facts.sourcePath                        = fourthPrompt->sourcePath;
    facts.consentDetail                     = fourthPrompt->consentDetail;
    const std::wstring expectedAnnouncement = FileOperations::BuildConflictPromptAnnouncement(facts);
    if (! mirror.tookFocus || RedSalamander::Ui::AlertOverlayFocusWindow() != mirror.overlay || ! overlay.restoresFocus || overlay.restoreFocus != leftView ||
        overlay.keysGuarded || overlay.guardsKeyboardFocus || overlay.focusedButtonId != std::optional<uint32_t>{cancelId})
        return reject(std::format(L"the mirror did not take immediately usable focus on Cancel (took={} restore={} guarded={})",
                                  mirror.tookFocus,
                                  overlay.restoreFocus == leftView,
                                  overlay.keysGuarded));
    if (! mirror.announced || overlay.lastAnnouncement != expectedAnnouncement ||
        operations->DebugGetPromptAnnouncer(fourthTask, fourthPrompt->generation) != std::optional<bool>{true})
        return reject(std::format(
            L"the mirror that took the focus did not announce the card's text itself (announced={} text='{}')", mirror.announced, overlay.lastAnnouncement));
    if (mirror.flashCount != flashBefore + 1u || mirror.foreground != mirror.overlay)
        return reject(std::format(L"the popup's taskbar button flashed {} time(s) for one new pending set while the mirror held the foreground",
                                  mirror.flashCount - flashBefore));
    PressMirrorKey(mirror.overlay, VK_TAB, 0);
    if (OverlayState(mirror.overlay).focusedButtonId != std::optional<uint32_t>{deleteId})
        return reject(L"Tab immediately after show did not reach the next button without a click or Ctrl+Tab");
    const bool shiftDown = SetMirrorKeyDown(VK_SHIFT, true);
    PressMirrorKey(mirror.overlay, VK_TAB, 0);
    const bool shiftReleased = SetMirrorKeyDown(VK_SHIFT, false);
    if (! shiftDown || ! shiftReleased || OverlayState(mirror.overlay).focusedButtonId != std::optional<uint32_t>{cancelId})
        return reject(L"Shift+Tab immediately after show did not return to Cancel");
    PressMirrorKey(mirror.overlay, VK_TAB, 0);
    PressMirrorKey(mirror.overlay, VK_LEFT, 0);
    if (OverlayState(mirror.overlay).focusedButtonId != std::optional<uint32_t>{cancelId})
        return reject(L"Left immediately after show did not return to Cancel");
    PressMirrorKey(mirror.overlay, VK_RIGHT, 0);
    if (OverlayState(mirror.overlay).focusedButtonId != std::optional<uint32_t>{deleteId})
        return reject(L"Right immediately after show did not reach the next button");
    PressMirrorKey(mirror.overlay, VK_ESCAPE, kMirrorKeyLateMs, true);
    overlay = OverlayState(mirror.overlay);
    if (! MirrorState(operations).visible || overlay.focusedButtonId != std::optional<uint32_t>{deleteId} || overlay.keysGuarded ||
        ! TryGetConflictPromptCopy(operations->FindTask(fourthTask)).has_value())
        return reject(L"navigation or an auto-repeated Escape answered the decision");
    PressMirrorKey(mirror.overlay, VK_ESCAPE, 0);
    if (MirrorState(operations).visible || RedSalamander::Ui::AlertOverlayFocusWindow() != leftView || MirrorState(operations).foreground != ownerRoot)
        return reject(L"Escape on the focused mirror did not answer and give the focus back to the pane");
    const auto fourthDone = scope.WaitCompleted(fourthTask);
    if (! fourthDone.has_value() || fourthDone->resultHr != canceledHr || ! FileExistsForMirror(sourceFolder / L"fourth.bin"))
        return reject(L"Escape on the mirror did not submit the prompt's escape action (Cancel)");

    // 4b. The keyboard way out. An auto-repeated F6 does nothing; a fresh F6 immediately sends
    //     the focus back to the pane with nothing answered, and the mirror stays
    //     shown without taking the focus again, also when it returns after the pane left the folder. A focus
    //     handed back to it without a click is immediately usable; Tab reaches the destructive button without pressing anything, and
    //     F6 leaves again.
    uint64_t keysTask     = 0u;
    const auto keysPrompt = startDelete(L"keys.bin", keysTask);
    if (! keysPrompt.has_value() || ! shownFor(keysTask, L"decision mirror keyboard way out"))
        return reject(L"the Delete for the keyboard way out did not reach the mirror");
    mirror = MirrorState(operations);
    if (! mirror.tookFocus || OverlayState(mirror.overlay).keysGuarded)
        return reject(L"the mirror did not take immediately usable focus for the keyboard way out");
    const uint64_t submitsBefore = mirror.submitCount;
    PressMirrorKey(mirror.overlay, VK_F6, kMirrorKeyLateMs, true);
    if (RedSalamander::Ui::AlertOverlayFocusWindow() != mirror.overlay || MirrorState(operations).focusReturnCount != mirror.focusReturnCount)
        return reject(L"an auto-repeated F6 left the mirror");
    PressMirrorKey(mirror.overlay, VK_F6, 0);
    const MirrorSnapshot returned = MirrorState(operations);
    if (RedSalamander::Ui::AlertOverlayFocusWindow() != leftView || ! returned.visible || returned.taskId != keysTask ||
        returned.submitCount != submitsBefore || returned.focusReturnCount != mirror.focusReturnCount + 1u || ! returned.focusDeclined ||
        returned.foreground != ownerRoot || ! TryGetConflictPromptCopy(operations->FindTask(keysTask)).has_value())
        return reject(L"F6 on the mirror did not send the focus back to the pane with nothing answered and the mirror still shown");
    scope.Settle();
    if (RedSalamander::Ui::AlertOverlayFocusWindow() != leftView || ! MirrorState(operations).visible)
        return reject(L"the mirror took the focus again, or left, after F6 sent the focus back to the pane");
    if (! scope.ShowPane(FolderWindow::Pane::Left, otherFolder, L"other.bin") ||
        ! scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return ! value.visible; }, L"decision mirror hides for the keyboard way out"))
        return reject(L"the mirror stayed over a pane that left the folder");
    RedSalamander::Ui::DebugSetAlertOverlayFocusOverride(leftView);
    if (! scope.ShowPane(FolderWindow::Pane::Left, sourceFolder, L"keys.bin") || ! shownFor(keysTask, L"decision mirror back without the focus"))
        return reject(L"the mirror did not return when the pane showed the folder again");
    mirror = MirrorState(operations);
    if (mirror.tookFocus || RedSalamander::Ui::AlertOverlayFocusWindow() != leftView)
        return reject(L"the mirror took the focus again for a prompt the user had sent the focus away from");
    RedSalamander::Ui::DebugSetAlertOverlayFocusOverride(mirror.overlay);
    static_cast<void>(SendMessageW(mirror.overlay, WM_SETFOCUS, reinterpret_cast<WPARAM>(leftView), 0));
    if (OverlayState(mirror.overlay).keysGuarded)
        return reject(L"a focus handed back to the mirror required a click or Ctrl+Tab");
    PressMirrorKey(mirror.overlay, VK_TAB, 0);
    overlay = OverlayState(mirror.overlay);
    if (overlay.focusedButtonId != std::optional<uint32_t>{deleteId} || MirrorState(operations).submitCount != submitsBefore ||
        ! TryGetConflictPromptCopy(operations->FindTask(keysTask)).has_value())
        return reject(L"Tab on returned focus did not reach the next button without pressing anything");
    PressMirrorKey(mirror.overlay, VK_F6, kMirrorKeyLateMs);
    if (RedSalamander::Ui::AlertOverlayFocusWindow() != leftView || ! TryGetConflictPromptCopy(operations->FindTask(keysTask)).has_value())
        return reject(L"F6 on the unguarded mirror did not send the focus back to the pane with nothing answered");
    if (! canceledOnCard(keysTask, keysPrompt->generation))
        return reject(L"the keyboard way out's Delete could not be canceled from the card");

    // 4c. The show fails once for a prompt the mirror would announce (it takes the focus): the announcement is
    //     not lost. Exactly one surface announces it, and the mirror only from a show that succeeded.
    operations->DebugFailNextDecisionMirrorShows(1u);
    const uint64_t failuresBefore = MirrorState(operations).showFailureCount;
    uint64_t failTask             = 0u;
    const auto failPrompt         = startDelete(L"failed-show.bin", failTask);
    if (! failPrompt.has_value() || ! shownFor(failTask, L"decision mirror after a failed show"))
        return reject(L"the mirror did not show the prompt after one failed show");
    mirror                                  = MirrorState(operations);
    const std::optional<bool> failAnnouncer = operations->DebugGetPromptAnnouncer(failTask, failPrompt->generation);
    FileOperations::ConflictPromptTextFacts failFacts{};
    failFacts.operation     = FILESYSTEM_DELETE;
    failFacts.bucket        = static_cast<uint8_t>(MirrorBucket::PermanentDeleteConfirmation);
    failFacts.sourcePath    = failPrompt->sourcePath;
    failFacts.consentDetail = failPrompt->consentDetail;
    const bool mirrorSpoke  = mirror.announced && OverlayState(mirror.overlay).lastAnnouncement == FileOperations::BuildConflictPromptAnnouncement(failFacts);
    if (mirror.showFailureCount != failuresBefore + 1u || ! mirror.tookFocus || ! failAnnouncer.has_value() || (failAnnouncer.value() && ! mirrorSpoke) ||
        (! failAnnouncer.value() && mirror.announced))
        return reject(std::format(L"a failed show lost the prompt's announcement (failures={} announcer={} mirror announced={})",
                                  mirror.showFailureCount - failuresBefore,
                                  failAnnouncer.has_value() ? (failAnnouncer.value() ? L"mirror" : L"card") : L"none",
                                  mirror.announced));
    if (! canceledOnCard(failTask, failPrompt->generation))
        return reject(L"the failed-show Delete could not be canceled from the card");

    // 4d. The user moves the focus to the other pane after the mirror took it: the mirror forgets where the
    //     focus came from. When the focus comes back to the mirror later (the user clicks it), answering on
    //     the card does not take the focus back to the pane the user had left.
    RedSalamander::Ui::DebugSetAlertOverlayFocusOverride(leftView);
    uint64_t staleTask     = 0u;
    const auto stalePrompt = startDelete(L"stale.bin", staleTask);
    if (! stalePrompt.has_value() || ! shownFor(staleTask, L"decision mirror focused again"))
        return reject(L"the Delete for the stale focus check did not reach the mirror");
    mirror = MirrorState(operations);
    if (! mirror.tookFocus)
        return reject(L"the mirror did not take the focus for the stale focus check");
    RedSalamander::Ui::DebugSetAlertOverlayFocusOverride(rightView);
    static_cast<void>(SendMessageW(mirror.overlay, WM_KILLFOCUS, reinterpret_cast<WPARAM>(rightView), 0));
    if (OverlayState(mirror.overlay).restoresFocus)
        return reject(L"the mirror kept its focus-return target after the user moved the focus to another pane");
    RedSalamander::Ui::DebugSetAlertOverlayFocusOverride(mirror.overlay);
    if (! canceledOnCard(staleTask, stalePrompt->generation) || RedSalamander::Ui::AlertOverlayFocusWindow() == leftView)
        return reject(L"answering on the card took the focus back to the pane the user had left");
    RedSalamander::Ui::DebugSetAlertOverlayFocusOverride(rightView);

    // 5. A foreground decision also takes focus from the other pane. Enter immediately chooses safe Cancel
    //    and restores that pane, rather than leaving the visible dialog inert.
    uint64_t fifthTask     = 0u;
    const auto fifthPrompt = startDelete(L"fifth.bin", fifthTask);
    if (! fifthPrompt.has_value() || ! shownFor(fifthTask, L"decision mirror unfocused"))
        return reject(L"the fifth Delete's decision did not reach the mirror");
    mirror = MirrorState(operations);
    if (! mirror.tookFocus || ! mirror.announced || RedSalamander::Ui::AlertOverlayFocusWindow() != mirror.overlay || OverlayState(mirror.overlay).keysGuarded)
        return reject(L"the decision shown from the other foreground pane did not take immediately usable focus");
    PressMirrorKey(mirror.overlay, VK_RETURN, 0);
    const auto fifthDone = scope.WaitCompleted(fifthTask);
    if (! fifthDone.has_value() || fifthDone->resultHr != canceledHr || MirrorState(operations).visible ||
        RedSalamander::Ui::AlertOverlayFocusWindow() != rightView || ! FileExistsForMirror(sourceFolder / L"fifth.bin"))
        return reject(L"immediate Enter did not cancel safely and restore the other pane's focus");
    RedSalamander::Ui::DebugSetAlertOverlayFocusOverride(std::optional<HWND>(nullptr));
    operations->DebugSetDecisionMirrorForegroundOverride(std::optional<HWND>(nullptr));

    // 6. The setting turned off while a mirror is shown hides it at the next pass; the status-bar segment
    //    stays; turned on again, the mirror returns. Off from the start, nothing appears over the pane.
    uint64_t settingTask     = 0u;
    const auto settingPrompt = startDelete(L"setting.bin", settingTask);
    if (! settingPrompt.has_value() || ! shownFor(settingTask, L"decision mirror before the setting"))
        return reject(L"the Delete for the setting check did not reach the mirror");
    EnsureFileOperationsSettingsForSelfTest().mirrorDecisionsInPane = false;
    scope.Settle();
    if (MirrorState(operations).visible || StatusDecisionText(folder, FolderWindow::Pane::Left) != pendingOne)
        return reject(L"turning the setting off left the mirror shown or removed the pending segment");
    EnsureFileOperationsSettingsForSelfTest().mirrorDecisionsInPane = true;
    if (! shownFor(settingTask, L"decision mirror after the setting"))
        return reject(L"the mirror did not return when the setting was turned on again");
    if (! canceledOnCard(settingTask, settingPrompt->generation))
        return reject(L"the setting check's Delete could not be canceled from the card");
    EnsureFileOperationsSettingsForSelfTest().mirrorDecisionsInPane = false;
    uint64_t offTask                                                = 0u;
    const auto offPrompt                                            = startDelete(L"off.bin", offTask);
    if (! offPrompt.has_value())
        return reject(L"the Delete with the setting off did not publish its decision");
    scope.Settle();
    if (MirrorState(operations).visible || MirrorState(operations).pendingCount != 1u || ! scope.WaitIndicator(pendingOne, L"indicator with the setting off"))
        return reject(L"with the setting off the mirror appeared, or the pending-decision segment did not");
    if (! canceledOnCard(offTask, offPrompt->generation))
        return reject(L"the Delete with the setting off could not be canceled from the card");
    EnsureFileOperationsSettingsForSelfTest().mirrorDecisionsInPane = true;

    // 7. A dialog of the owner runs a modal loop (the owner window is disabled): the mirror leaves the screen
    //    and returns once the owner accepts input again.
    uint64_t modalTask     = 0u;
    const auto modalPrompt = startDelete(L"modal.bin", modalTask);
    if (! modalPrompt.has_value() || ! shownFor(modalTask, L"decision mirror before the modal"))
        return reject(L"the Delete for the modal check did not reach the mirror");
    EnableWindow(ownerRoot, FALSE);
    scope.Settle();
    const MirrorSnapshot underModal = MirrorState(operations);
    EnableWindow(ownerRoot, TRUE);
    if (underModal.visible || underModal.windowVisible)
        return reject(L"the mirror stayed above a modal dialog of its owner window");
    if (! shownFor(modalTask, L"decision mirror after the modal"))
        return reject(L"the mirror did not return when the owner window accepted input again");
    if (! canceledOnCard(modalTask, modalPrompt->generation))
        return reject(L"the modal check's Delete could not be canceled from the card");

    AppendLog(L"Decision mirror: shown over the folder's pane with the card's question and actions (the card defers to the reconcile, then "
              L"announces), answered once on the mirror, hidden after the card's answer, on navigation, on Cancel all and on close (not shown "
              L"again), status-bar segment kept when the popup closed and opening it, immediate keyboard focus with the mirror's own "
              L"announcement and one flash, Tab/Shift+Tab and arrows without a click, safe Enter and Escape, F6 back to the pane without an answer "
              L"and no focus taken again, usable focus return, an announcement kept across a failed show, stale focus return forgotten, inert keys "
              L"without focus, the setting and a modal owner hide it.");
    return true;
}

// R8.2/R8.6 on a same-host overlap consent: the engine's default (Queue after these tasks) and Escape (Don't
// start) are the mirror's Enter and Escape; Escape ends that task before any change.
[[nodiscard]] bool RunFileOpsDecisionMirrorOverlapChecks(SelfTestState& state) noexcept
{
    using FileOps     = FolderWindow::FileOperationState;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Decision mirror overlap: {}", reason));
        return false;
    };
    FolderWindow* folder = nullptr;
    FileOps* operations  = nullptr;
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! MirrorCaseReady(state, folder, operations, blocker))
        return reject(L"owner, Local panes, the no-activation guard or the initial worker quiet point unavailable");

    const std::filesystem::path root = state.tempRoot / L"decision-mirror-overlap";
    const auto holdSource            = root / L"hold-src" / L"collision.bin";
    const auto holdDestination       = root / L"hold-dst" / L"collision.bin";
    const auto elsewhere             = root / L"elsewhere";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(elsewhere) || ! WriteFilledTestFile(holdSource, 4096u, 0x11u) ||
        ! WriteFilledTestFile(holdDestination, 8192u, 0x22u) || ! WriteTestFile(elsewhere / L"other.bin", 64u))
        return reject(L"fixture unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    DecisionMirrorCaseScope scope(state, folder, operations);
    const HWND leftView = folder->GetFolderViewHwnd(FolderWindow::Pane::Left);
    if (! leftView || IsWindowVisible(leftView) == FALSE)
        return reject(L"the left pane's view is not visible");
    // Neither pane shows the Copy's destination, so the Copy's own question is never mirrored here.
    if (! scope.ShowPane(FolderWindow::Pane::Left, holdSource.parent_path(), L"collision.bin") ||
        ! scope.ShowPane(FolderWindow::Pane::Right, elsewhere, L"other.bin"))
        return reject(L"pane listings unavailable");

    // A Copy waiting on its own name-collision question is active and reads the source; a Delete of that
    // source then asks the same-host overlap question before it changes anything.
    const uint64_t copyTask = scope.StartCopy({holdSource}, holdDestination.parent_path());
    if (copyTask == 0u)
        return reject(L"the Copy did not start");
    if (! scope.WaitPrompt(operations, copyTask, std::nullopt).has_value())
        return reject(L"the Copy did not ask about the existing destination file");

    RedSalamander::Ui::DebugSetAlertOverlayFocusOverride(leftView);
    operations->DebugSetDecisionMirrorForegroundOverride(GetAncestor(state.mainWindow, GA_ROOT));
    const auto deleteTask = StartFileOperationAndGetId(operations,
                                                       FILESYSTEM_DELETE,
                                                       FolderWindow::Pane::Left,
                                                       std::nullopt,
                                                       state.fsLocal,
                                                       {holdSource},
                                                       {},
                                                       static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN),
                                                       false);
    if (! deleteTask.has_value())
        return reject(L"the overlapping Delete did not start");
    scope.Track(deleteTask.value());
    const auto prompt = scope.WaitPrompt(deleteTask.value(), MirrorBucket::SameHostOverlap);
    if (! prompt.has_value() || prompt->defaultAction != MirrorAction::Proceed || prompt->escapeAction != MirrorAction::Cancel ||
        prompt->overlapTaskId != copyTask)
        return reject(L"the Delete did not ask the same-host overlap question naming the Copy, with Queue after (default) and Don't start (Escape)");
    if (! scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return value.visible && value.taskId == deleteTask.value(); },
                           L"decision mirror overlap shown"))
        return reject(L"the overlap question did not reach the mirror over the source's pane");
    const MirrorSnapshot mirror = MirrorState(operations);
    const MirrorOverlay overlay = OverlayState(mirror.overlay);
    const uint32_t proceedId    = FileOps::DebugDecisionMirrorPrimaryButtonId(MirrorAction::Proceed);
    if (overlay.buttonIds.empty() || overlay.buttonIds.front() != proceedId || overlay.primaryButtonId != std::optional<uint32_t>{proceedId} ||
        overlay.focusedButtonId != std::optional<uint32_t>{proceedId} ||
        ! MirrorHasButton(overlay, FileOps::DebugDecisionMirrorPrimaryButtonId(MirrorAction::Cancel)) ||
        overlay.buttonLabels.front() != FileOperations::FormatConflictActionLabel(MirrorAction::Proceed, MirrorBucket::SameHostOverlap, FILESYSTEM_DELETE) ||
        ! mirror.tookFocus || overlay.keysGuarded || overlay.guardsKeyboardFocus || mirror.pendingCount != 2u)
        return reject(std::format(L"the mirror did not show the engine's overlap actions with Queue after as the focused default (buttons={} "
                                  L"tookFocus={} guarded={} pending={})",
                                  overlay.buttonIds.size(),
                                  mirror.tookFocus,
                                  overlay.keysGuarded,
                                  mirror.pendingCount));
    PressMirrorKey(mirror.overlay, VK_ESCAPE, kMirrorKeyLateMs);
    const auto deleted = scope.WaitCompleted(deleteTask.value());
    if (! deleted.has_value() || deleted->resultHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || deleted->warningCount != 0u || ! FileExistsForMirror(holdSource))
        return reject(L"Escape on the mirror did not answer Don't start before any change");
    const auto copyPrompt = TryGetConflictPromptCopy(operations->FindTask(copyTask));
    if (const MirrorTask* copy = operations->FindTask(copyTask); ! copy || copy->CancelRequestedDurable() || ! copyPrompt.has_value())
        return reject(L"Escape on the Delete's mirror reached the Copy or its question");
    AppendLog(L"Decision mirror overlap: Queue after is the focused default, Escape answers Don't start for that task only.");
    return true;
}

// R8.10: a pre-task prompt anchors on the pane the request came from (its view), not on the whole window.
[[nodiscard]] bool RunFileOpsDecisionMirrorIngressChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Decision mirror ingress: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    if (! MirrorLocalPanes(folder))
        return reject(L"Local panes unavailable");
    const std::filesystem::path root = state.tempRoot / L"decision-mirror-ingress";
    const auto source                = root / L"src" / L"payload.bin";
    const auto destination           = root / L"dst";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(destination) || ! WriteTestFile(source, 1024u))
        return reject(L"fixture unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    DecisionMirrorCaseScope scope(state, folder, operations);
    if (! scope.ShowPane(FolderWindow::Pane::Left, source.parent_path(), L"payload.bin") || ! scope.ShowPane(FolderWindow::Pane::Right, destination, {}))
        return reject(L"pane listings unavailable");
    const HWND leftView = folder->GetFolderViewHwnd(FolderWindow::Pane::Left);
    if (! leftView || IsWindowVisible(leftView) == FALSE)
        return reject(L"the left pane's view is not visible");

    const auto lastPromptAnchoredOnLeft = [&](std::wstring_view what) noexcept
    {
        HostPromptDebugSnapshot prompt{};
        const bool anchored = HostGetTestLastPromptDebugSnapshot(prompt) && prompt.scope == HOST_ALERT_SCOPE_WINDOW && prompt.targetWindow == leftView;
        AppendLog(std::format(L"Decision mirror ingress: {} prompt scope={} anchored on the left pane={}", what, static_cast<int>(prompt.scope), anchored));
        return anchored;
    };
    // An explicit Copy confirmation, declined: nothing starts.
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    uint64_t prompts = HostGetTestPromptRequestCount();
    const auto copy  = StartFileOperationAndGetId(operations,
                                                  FILESYSTEM_COPY,
                                                  FolderWindow::Pane::Left,
                                                  FolderWindow::Pane::Right,
                                                  state.fsLocal,
                                                  {source},
                                                  destination,
                                                  FILESYSTEM_FLAG_NONE,
                                                  false,
                                                  0u,
                                                  FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                  true,
                                                  state.fsLocal);
    if (copy.has_value() || HostGetTestPromptRequestCount() != prompts + 1u || ! lastPromptAnchoredOnLeft(L"Copy confirmation"))
        return reject(L"the explicit Copy confirmation was not one prompt anchored on the originating pane");
    // A confirmed Delete to the Recycle Bin, declined.
    prompts          = HostGetTestPromptRequestCount();
    const auto erase = StartFileOperationAndGetId(operations,
                                                  FILESYSTEM_DELETE,
                                                  FolderWindow::Pane::Left,
                                                  std::nullopt,
                                                  state.fsLocal,
                                                  {source},
                                                  {},
                                                  static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN),
                                                  false,
                                                  0u,
                                                  FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                  true);
    if (erase.has_value() || HostGetTestPromptRequestCount() != prompts + 1u || ! lastPromptAnchoredOnLeft(L"Delete confirmation"))
        return reject(L"the Delete confirmation was not one prompt anchored on the originating pane");
    if (! FileExistsForMirror(source) || operations->HasActiveOperations())
        return reject(L"a declined pre-task prompt started work");

    // A recursive Change Attributes over a folder holding a possible leftover asks its touch confirmation from
    // its worker; the prompt covers the pane the command came from, and declined, it changes nothing.
    const auto attributesRoot   = root / L"attributes";
    const auto attributesFolder = attributesRoot / L"folder";
    const auto leftover         = attributesFolder / L"payload.rs_ren_0123456789abcdef0123456789abcdef";
    if (! RecreateEmptyDirectory(attributesRoot) || ! RecreateEmptyDirectory(attributesFolder) || ! WriteTestFile(leftover, 64u) ||
        ! scope.ShowPane(FolderWindow::Pane::Left, attributesRoot, L"folder"))
        return reject(L"the Change Attributes fixture is unavailable");
    folder->SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](const std::wstring_view name) noexcept { return name == L"folder"; }, true);
    if (folder->DebugGetSelectedCount(FolderWindow::Pane::Left) != 1u)
        return reject(L"the Change Attributes folder could not be selected");
    FolderWindow::ChangeAttributesOptions attributes{};
    attributes.hidden                = FolderWindow::AttributeChangeState::Set;
    attributes.includeSubdirectories = true;
    folder->DebugSetNextChangeAttributesOptions(attributes);
    const auto clearAttributes = wil::scope_exit([&]() noexcept { folder->DebugSetNextChangeAttributesOptions(std::nullopt); });
    prompts                    = HostGetTestPromptRequestCount();
    folder->CommandChangeAttributes(FolderWindow::Pane::Left);
    std::optional<FolderWindow::ChangeAttributesReport> attributesReport;
    const bool attributesDone      = RedSalamander::TestSupport::PumpMessagesUntil(
                                         [&]() noexcept
    {
        attributesReport = folder->DebugGetLastChangeAttributesReport();
        return attributesReport.has_value();
    },
                                    {.timeout = SelfTest::Scale(kMirrorWait), .operationName = L"decision mirror ingress Change Attributes"})
                                         .conditionMet;
    const DWORD leftoverAttributes = GetFileAttributesW(leftover.c_str());
    if (! attributesDone || HostGetTestPromptRequestCount() != prompts + 1u || ! lastPromptAnchoredOnLeft(L"Change Attributes leftover touch") ||
        attributesReport->itemsProcessed != 0u || leftoverAttributes == INVALID_FILE_ATTRIBUTES || (leftoverAttributes & FILE_ATTRIBUTE_HIDDEN) != 0u)
        return reject(L"the recursive Change Attributes leftover confirmation was not one prompt anchored on the originating pane, or it changed an item");
    AppendLog(L"Decision mirror ingress: the Copy and Delete confirmations and a recursive Change Attributes leftover confirmation (asked from "
              L"its worker) anchor on the originating pane's view.");
    return true;
}

// R8.4/R8.8: a Compare Directories window has its own mirror over its own panes, independent of the main
// window's; closing that window with its mirror shown tears the mirror down with no decision taken.
[[nodiscard]] bool RunFileOpsDecisionMirrorCompareWindowChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Decision mirror Compare window: {}", reason));
        return false;
    };
    FolderWindow* folder                         = nullptr;
    FolderWindow::FileOperationState* operations = nullptr;
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! MirrorCaseReady(state, folder, operations, blocker) || folder->IsFileOperationsExitPending())
        return reject(L"owner, Local panes, the no-activation guard or an idle operation state unavailable");
    if (GetCompareDirectoriesWindowHandle() != nullptr)
        return reject(L"a Compare Directories window was already open");

    const std::filesystem::path root = state.tempRoot / L"decision-mirror-compare";
    const auto left                  = root / L"left";
    const auto right                 = root / L"right";
    const auto mainItem              = left / L"main.bin";
    const auto compareItem           = left / L"compare.bin";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(left) || ! RecreateEmptyDirectory(right) || ! WriteTestFile(mainItem, 1024u) ||
        ! WriteTestFile(compareItem, 1024u) || ! WriteTestFile(right / L"right.bin", 64u))
        return reject(L"fixture unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    DecisionMirrorCaseScope scope(state, folder, operations);
    HWND compareWindow          = nullptr;
    FolderWindow* compareFolder = nullptr;
    uint64_t compareTask        = 0u;
    const auto closeCompare     = wil::scope_exit([&]() noexcept
    {
        if (compareWindow == nullptr || IsWindow(compareWindow) == FALSE)
            return;
        if (auto* compareOperations = compareFolder != nullptr ? TryGetFileOps(compareFolder) : nullptr)
        {
            static_cast<void>(CancelOutstandingSelfTestOperations(compareOperations));
            if (compareTask != 0u)
            {
                compareOperations->DismissCompletedTask(compareTask);
                compareOperations->DebugRemoveDiagnosticsForTask(compareTask);
            }
        }
        const HWND window = compareWindow;
        PostMessageW(window, WM_CLOSE, 0, 0);
        static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return IsWindow(window) == FALSE;
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"decision mirror Compare close"}));
    });
    if (! scope.ShowPane(FolderWindow::Pane::Left, left, L"compare.bin") || ! scope.ShowPane(FolderWindow::Pane::Right, right, L"right.bin"))
        return reject(L"pane listings unavailable");
    SendMessageW(state.mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    const auto opened = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        compareWindow = GetCompareDirectoriesWindowHandle();
        return compareWindow != nullptr;
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"decision mirror Compare open"});
    compareFolder = opened.conditionMet ? DebugGetCompareDirectoriesFolderWindowForWindow(compareWindow) : nullptr;
    if (! compareFolder)
        return reject(L"Compare Directories did not open for the two panes");
    SendMessageW(compareWindow, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_RESCAN, 0), 0);
    const auto compared = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        CompareDirectoriesRunDebugSnapshot snapshot{};
        return DebugGetCompareDirectoriesRunSnapshotForWindow(compareWindow, snapshot) && snapshot.compareStarted && snapshot.compareActive &&
               ! snapshot.compareRunPending && compareFolder->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"compare.bin");
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(15)), .operationName = L"decision mirror Compare run"});
    auto* compareOperations = compared.conditionMet ? TryGetFileOps(compareFolder) : nullptr;
    if (! compareOperations)
        return reject(L"the Compare Directories run did not settle on the two folders");
    const HWND compareView = compareFolder->GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND mainView    = folder->GetFolderViewHwnd(FolderWindow::Pane::Left);
    if (! compareView || IsWindowVisible(compareView) == FALSE || ! mainView || IsWindowVisible(mainView) == FALSE)
        return reject(L"a left pane view is not visible");
    compareOperations->DebugSetDecisionMirrorForegroundOverride(std::optional<HWND>(nullptr));

    // One permanent Delete in each engine: each mirror sits over its own window's pane.
    compareTask             = scope.StartPermanentDelete(compareOperations, compareItem);
    const uint64_t mainTask = scope.StartPermanentDelete(mainItem);
    if (compareTask == 0u || mainTask == 0u)
        return reject(L"the two Deletes did not start");
    MirrorSnapshot compareMirror{};
    MirrorSnapshot mainMirror{};
    const auto bothShown = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        compareMirror = MirrorState(compareOperations);
        mainMirror    = MirrorState(operations);
        return compareMirror.visible && compareMirror.taskId == compareTask && mainMirror.visible && mainMirror.taskId == mainTask;
    },
        {.timeout = SelfTest::Scale(kMirrorWait), .operationName = L"decision mirror Compare both shown"});
    if (! bothShown.conditionMet || compareMirror.anchor != compareView || mainMirror.anchor != mainView || compareMirror.overlay == mainMirror.overlay ||
        compareMirror.overlay == nullptr)
        return reject(L"the Compare window and the main window did not each mirror their own decision over their own pane");

    // Close the Compare window, accepting its cancel-and-close consent: its mirror goes with it, with no
    // decision taken, while the main window's mirror stays.
    const HWND compareOverlay = compareMirror.overlay;
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    PostMessageW(compareWindow, WM_CLOSE, 0, 0);
    const auto closed = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return IsWindow(compareWindow) == FALSE;
    }, {.timeout = SelfTest::Scale(kMirrorWait), .operationName = L"decision mirror Compare closed"});
    HostClearTestPromptResultOverride();
    if (closed.conditionMet)
    {
        compareFolder = nullptr;
        compareWindow = nullptr;
    }
    if (! closed.conditionMet || IsWindow(compareOverlay) != FALSE || ! FileExistsForMirror(compareItem))
        return reject(L"closing the Compare window left its mirror window, or a decision deleted its item");
    const MirrorSnapshot mainAfter = MirrorState(operations);
    if (! mainAfter.visible || mainAfter.taskId != mainTask || IsWindow(mainAfter.overlay) == FALSE)
        return reject(L"closing the Compare window took the main window's mirror with it");
    if (! RedSalamander::Ui::DebugInvokeAlertOverlayButton(mainAfter.overlay,
                                                           FolderWindow::FileOperationState::DebugDecisionMirrorPrimaryButtonId(MirrorAction::Cancel)))
        return reject(L"the main mirror's Cancel could not be pressed");
    const auto mainDone = scope.WaitCompleted(mainTask);
    if (! mainDone.has_value() || mainDone->resultHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || ! FileExistsForMirror(mainItem))
        return reject(L"the main window's mirror did not decide its own task");
    AppendLog(L"Decision mirror Compare window: independent mirrors per window; the Compare window's mirror is torn down with it.");
    return true;
}

// D-8.1 and the shared Apply-to-all flag on a Copy with several name collisions (one question per item): a run
// of consecutive prompts flashes the popup's taskbar button once and keeps the status segment between them;
// Apply to all on the mirror is the engine's flag the card's checkbox edits, and the answer carries it.
[[nodiscard]] bool RunFileOpsDecisionMirrorConsecutiveChecks(SelfTestState& state) noexcept
{
    using FileOps     = FolderWindow::FileOperationState;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Decision mirror consecutive prompts: {}", reason));
        return false;
    };
    FolderWindow* folder = nullptr;
    FileOps* operations  = nullptr;
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! MirrorCaseReady(state, folder, operations, blocker))
        return reject(L"owner, Local panes, the no-activation guard or the initial worker quiet point unavailable");

    const std::filesystem::path root = state.tempRoot / L"decision-mirror-consecutive";
    const auto skipSource            = root / L"skip-src";
    const auto skipDestination       = root / L"skip-dst";
    const auto allSource             = root / L"all-src";
    const auto allDestination        = root / L"all-dst";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(skipSource) || ! RecreateEmptyDirectory(skipDestination) ||
        ! RecreateEmptyDirectory(allSource) || ! RecreateEmptyDirectory(allDestination))
        return reject(L"fixture folders unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    std::vector<std::filesystem::path> skipItems;
    std::vector<std::filesystem::path> allItems;
    for (const wchar_t* name : {L"a.bin", L"b.bin", L"c.bin"})
    {
        if (! WriteFilledTestFile(skipSource / name, 1024u, 0x31u) || ! WriteFilledTestFile(skipDestination / name, 2048u, 0x32u) ||
            ! WriteFilledTestFile(allSource / name, 1024u, 0x33u) || ! WriteFilledTestFile(allDestination / name, 2048u, 0x34u))
            return reject(L"fixture files unavailable");
        skipItems.push_back(skipSource / name);
        allItems.push_back(allSource / name);
    }
    DecisionMirrorCaseScope scope(state, folder, operations);
    const HWND ownerRoot = GetAncestor(state.mainWindow, GA_ROOT);
    if (! scope.ShowPane(FolderWindow::Pane::Left, skipSource, L"a.bin") || ! scope.ShowPane(FolderWindow::Pane::Right, skipDestination, L"a.bin"))
        return reject(L"pane listings unavailable");
    const std::wstring_view pendingEmpty{};
    if (! scope.WaitIndicator(pendingEmpty, L"no decision pending before the run"))
        return reject(L"a pending-decision segment was left from before the run");

    // 1. Three questions in a row, each answered Skip on the mirror over the destination pane: one flash for
    //    the run, and the status segment never goes blank between two questions.
    operations->DebugEnsurePopupVisibleForSelfTest();
    operations->DebugSetDecisionMirrorForegroundOverride(ownerRoot);
    const uint64_t flashBefore = MirrorState(operations).flashCount;
    const uint64_t skipTask    = scope.StartCopy(skipItems, skipDestination);
    if (skipTask == 0u)
        return reject(L"the Copy with three collisions did not start");
    bool segmentBlinked     = false;
    uint64_t lastGeneration = 0u;
    const uint32_t skipId   = FileOps::DebugDecisionMirrorPrimaryButtonId(MirrorAction::Skip);
    for (int question = 0; question < 3; ++question)
    {
        const auto prompt = scope.WaitPrompt(operations,
                                             skipTask,
                                             MirrorBucket::RegularFileExists,
                                             lastGeneration,
                                             [&]()
        {
            if (question > 0 && StatusDecisionText(folder, FolderWindow::Pane::Right).empty())
                segmentBlinked = true;
        });
        if (! prompt.has_value())
            return reject(std::format(L"question {} of the Copy was not asked", question + 1));
        lastGeneration = prompt->generation;
        if (! scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return value.visible && value.generation == prompt->generation; },
                               L"decision mirror consecutive question"))
            return reject(std::format(L"question {} did not reach the mirror over the destination pane", question + 1));
        if (MirrorState(operations).pane != FolderWindow::Pane::Right)
            return reject(L"a Copy's question was not mirrored over the pane that shows its destination folder");
        if (! RedSalamander::Ui::DebugInvokeAlertOverlayButton(MirrorState(operations).overlay, skipId))
            return reject(L"the mirror's Skip could not be pressed");
    }
    const auto skipDone = scope.WaitCompleted(skipTask);
    if (! skipDone.has_value() || skipDone->warningCount != 0u || skipDone->errorCount != 0u)
        return reject(L"the Copy whose three questions were answered Skip did not end with skipped items only");
    const uint64_t flashes = MirrorState(operations).flashCount - flashBefore;
    if (flashes != 1u || segmentBlinked)
        return reject(
            std::format(L"a run of three questions flashed {} time(s) (expected 1) or blanked the status segment between them ({})", flashes, segmentBlinked));
    if (! scope.WaitIndicator(pendingEmpty, L"no decision pending after the run"))
        return reject(L"the pending-decision segment stayed after the run ended");

    // 2. Apply to all: the mirror's option is the engine's flag. Turned on over the pane, the card sees it;
    //    turned off on the card, the toggle posts a presentation pass and the mirror shows it in place (no
    //    new presentation); turned on again and answered Keep both on the mirror, the remaining collisions
    //    take that answer without asking.
    if (! scope.ShowPane(FolderWindow::Pane::Left, allSource, L"a.bin") || ! scope.ShowPane(FolderWindow::Pane::Right, allDestination, L"a.bin"))
        return reject(L"pane listings for Apply to all unavailable");
    operations->DebugSetDecisionMirrorForegroundOverride(std::optional<HWND>(nullptr));
    const uint64_t allTask = scope.StartCopy(allItems, allDestination);
    const auto allPrompt   = allTask != 0u ? scope.WaitPrompt(operations, allTask, MirrorBucket::RegularFileExists) : std::nullopt;
    if (! allPrompt.has_value() || ! allPrompt->applyToAllEligible)
        return reject(L"the Copy did not ask an Apply-to-all eligible question");
    if (! scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return value.visible && value.generation == allPrompt->generation; },
                           L"decision mirror apply to all"))
        return reject(L"the Apply-to-all question did not reach the mirror");
    const HWND allOverlay  = MirrorState(operations).overlay;
    const uint32_t applyId = FileOps::DebugDecisionMirrorApplyToAllOptionId();
    const auto optionValue = [&]() noexcept -> std::optional<uint64_t>
    {
        const MirrorOverlay snapshot = OverlayState(allOverlay);
        for (size_t index = 0; index < snapshot.optionIds.size() && index < snapshot.optionValues.size(); ++index)
            if (snapshot.optionIds[index] == applyId)
                return snapshot.optionValues[index];
        return std::nullopt;
    };
    const auto engineChecked = [&]() noexcept
    {
        const auto prompt = TryGetConflictPromptCopy(operations->FindTask(allTask));
        return prompt.has_value() && prompt->applyToAllChecked;
    };
    if (optionValue() != std::optional<uint64_t>{0u} || engineChecked())
        return reject(L"the mirror's Apply to all did not start from the card's unchecked state");
    if (! RedSalamander::Ui::DebugInvokeAlertOverlayOption(allOverlay, applyId) || ! engineChecked())
        return reject(L"Apply to all turned on over the pane did not reach the engine flag the card shows");
    // The card's toggle (the same engine call its checkbox makes) posts a presentation pass: drain the passes
    // already queued, toggle, and the pass is waiting before anything else runs.
    const uint64_t presentationBefore = OverlayState(allOverlay).presentationGeneration;
    const HWND folderWindow           = folder->GetHwnd();
    MSG queued{};
    while (PeekMessageW(&queued, folderWindow, WndMsg::kFileOperationPresentationChanged, WndMsg::kFileOperationPresentationChanged, PM_REMOVE) != FALSE)
        static_cast<void>(DispatchMessageW(&queued));
    if (MirrorTask* task = operations->FindTask(allTask))
        task->ToggleConflictApplyToAllChecked(allPrompt->generation);
    const bool passPosted =
        PeekMessageW(&queued, folderWindow, WndMsg::kFileOperationPresentationChanged, WndMsg::kFileOperationPresentationChanged, PM_NOREMOVE) != FALSE;
    const bool followed = scope.WaitMirror([&](const MirrorSnapshot&) noexcept { return optionValue() == std::optional<uint64_t>{0u}; },
                                           L"decision mirror follows the card's Apply to all");
    if (! passPosted || ! followed || engineChecked() || OverlayState(allOverlay).presentationGeneration != presentationBefore)
        return reject(std::format(L"Apply to all turned off on the card was not shown in place on the mirror at once (pass posted={} followed={} "
                                  L"re-presented={})",
                                  passPosted,
                                  followed,
                                  OverlayState(allOverlay).presentationGeneration != presentationBefore));
    if (! RedSalamander::Ui::DebugInvokeAlertOverlayOption(allOverlay, applyId) || ! engineChecked())
        return reject(L"Apply to all turned on again over the pane did not reach the engine flag");
    if (! RedSalamander::Ui::DebugInvokeAlertOverlayButton(allOverlay, FileOps::DebugDecisionMirrorPrimaryButtonId(MirrorAction::KeepBoth)))
        return reject(L"the mirror's Keep both could not be pressed");
    const auto allDone      = scope.WaitCompleted(allTask);
    size_t destinationFiles = 0u;
    std::error_code listError;
    for (const auto& entry : std::filesystem::directory_iterator(allDestination, listError))
    {
        static_cast<void>(entry);
        ++destinationFiles;
    }
    if (! allDone.has_value() || allDone->resultHr != S_OK || allDone->warningCount != 0u || listError || destinationFiles != 6u)
        return reject(std::format(L"Keep both with Apply to all on the mirror did not keep all three collisions (files={}, hr=0x{:08X})",
                                  destinationFiles,
                                  allDone.has_value() ? static_cast<unsigned long>(allDone->resultHr) : 0ul));
    AppendLog(L"Decision mirror consecutive prompts: one flash and a steady status segment for a run of three questions; Apply to all on the "
              L"mirror is the card's engine flag in both directions (the card's toggle posts a pass, shown in place) and the answer applied it "
              L"to the remaining collisions.");
    return true;
}

// R8.6: one mirror at a time, the oldest pending decision first (queue order), and a shown decision stays put
// while it remains eligible (no jump to an older decision that becomes eligible later).
[[nodiscard]] bool RunFileOpsDecisionMirrorQueueOrderChecks(SelfTestState& state) noexcept
{
    using FileOps     = FolderWindow::FileOperationState;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Decision mirror queue order: {}", reason));
        return false;
    };
    FolderWindow* folder = nullptr;
    FileOps* operations  = nullptr;
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! MirrorCaseReady(state, folder, operations, blocker))
        return reject(L"owner, Local panes, the no-activation guard or the initial worker quiet point unavailable");

    const std::filesystem::path root = state.tempRoot / L"decision-mirror-order";
    const auto olderFolder           = root / L"older";
    const auto newerFolder           = root / L"newer";
    const auto sharedFolder          = root / L"shared";
    const auto elsewhere             = root / L"elsewhere";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(olderFolder) || ! RecreateEmptyDirectory(newerFolder) ||
        ! RecreateEmptyDirectory(sharedFolder) || ! RecreateEmptyDirectory(elsewhere) || ! WriteTestFile(olderFolder / L"older.bin", 512u) ||
        ! WriteTestFile(newerFolder / L"newer.bin", 512u) || ! WriteTestFile(sharedFolder / L"first.bin", 512u) ||
        ! WriteTestFile(sharedFolder / L"second.bin", 512u) || ! WriteTestFile(elsewhere / L"other.bin", 64u))
        return reject(L"fixture unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    DecisionMirrorCaseScope scope(state, folder, operations);
    const uint32_t cancelId = FileOps::DebugDecisionMirrorPrimaryButtonId(MirrorAction::Cancel);

    // 1. Keep current: the newer task's decision is the only one shown in a pane; once the older task's folder
    //    is shown too, the mirror stays on the newer decision until it is answered, then moves to the older.
    if (! scope.ShowPane(FolderWindow::Pane::Left, newerFolder, L"newer.bin") || ! scope.ShowPane(FolderWindow::Pane::Right, elsewhere, L"other.bin"))
        return reject(L"pane listings unavailable");
    const uint64_t olderTask = scope.StartPermanentDelete(olderFolder / L"older.bin");
    const uint64_t newerTask = scope.StartPermanentDelete(newerFolder / L"newer.bin");
    const auto olderPrompt   = olderTask != 0u ? scope.WaitPrompt(olderTask, MirrorBucket::PermanentDeleteConfirmation) : std::nullopt;
    const auto newerPrompt   = newerTask != 0u ? scope.WaitPrompt(newerTask, MirrorBucket::PermanentDeleteConfirmation) : std::nullopt;
    if (! olderPrompt.has_value() || ! newerPrompt.has_value())
        return reject(L"the two Deletes did not both ask their confirmation");
    if (! scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return value.visible && value.taskId == newerTask; }, L"decision mirror newer"))
        return reject(L"the newer task's decision, the only one shown in a pane, was not mirrored");
    if (! scope.ShowPane(FolderWindow::Pane::Right, olderFolder, L"older.bin"))
        return reject(L"the right pane did not show the older task's folder");
    scope.Settle();
    if (const MirrorSnapshot kept = MirrorState(operations); ! kept.visible || kept.taskId != newerTask || kept.pendingCount != 2u || kept.evaluatedCount != 2u)
        return reject(L"the mirror jumped away from the decision it showed when an older decision became eligible");
    if (! SubmitOnCard(operations, newerTask, newerPrompt->generation, MirrorAction::Cancel) || ! scope.WaitCompleted(newerTask).has_value())
        return reject(L"the newer task could not be answered on the card");
    if (! scope.WaitMirror([&](const MirrorSnapshot& value) noexcept {
        return value.visible && value.taskId == olderTask && value.pane == FolderWindow::Pane::Right;
    }, L"decision mirror older"))
        return reject(L"the mirror did not move to the older decision over the pane that shows it");
    // The prompts a reconcile evaluated are the live pending set (R8.12's card deferral), never a history.
    if (const size_t evaluated = MirrorState(operations).evaluatedCount; evaluated != 1u)
        return reject(std::format(L"the reconcile still counts an answered prompt among the evaluated ones ({} evaluated, 1 pending)", evaluated));
    if (! RedSalamander::Ui::DebugInvokeAlertOverlayButton(MirrorState(operations).overlay, cancelId) || ! scope.WaitCompleted(olderTask).has_value())
        return reject(L"the older task could not be answered on the mirror");

    // 2. FIFO: two decisions about the same folder are pending before the pane shows it; the mirror shows the
    //    task queued first, then the other once the first is answered.
    if (! scope.ShowPane(FolderWindow::Pane::Left, elsewhere, L"other.bin") || ! scope.ShowPane(FolderWindow::Pane::Right, elsewhere, L"other.bin"))
        return reject(L"pane listings for the FIFO check unavailable");
    const uint64_t firstTask  = scope.StartPermanentDelete(sharedFolder / L"first.bin");
    const uint64_t secondTask = scope.StartPermanentDelete(sharedFolder / L"second.bin");
    const auto firstPrompt    = firstTask != 0u ? scope.WaitPrompt(firstTask, MirrorBucket::PermanentDeleteConfirmation) : std::nullopt;
    const auto secondPrompt   = secondTask != 0u ? scope.WaitPrompt(secondTask, MirrorBucket::PermanentDeleteConfirmation) : std::nullopt;
    if (! firstPrompt.has_value() || ! secondPrompt.has_value())
        return reject(L"the two Deletes in one folder did not both ask their confirmation");
    if (! scope.ShowPane(FolderWindow::Pane::Left, sharedFolder, L"first.bin") ||
        ! scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return value.visible; }, L"decision mirror FIFO first"))
        return reject(L"neither decision reached the mirror once the pane showed their folder");
    if (MirrorState(operations).taskId != firstTask)
        return reject(L"the mirror did not show the task queued first");
    if (! RedSalamander::Ui::DebugInvokeAlertOverlayButton(MirrorState(operations).overlay, cancelId) || ! scope.WaitCompleted(firstTask).has_value())
        return reject(L"the first task could not be answered on the mirror");
    if (! scope.WaitMirror([&](const MirrorSnapshot& value) noexcept { return value.visible && value.taskId == secondTask; }, L"decision mirror FIFO second"))
        return reject(L"the mirror did not move to the next decision once the first was answered");
    if (! RedSalamander::Ui::DebugInvokeAlertOverlayButton(MirrorState(operations).overlay, cancelId) || ! scope.WaitCompleted(secondTask).has_value())
        return reject(L"the second task could not be answered on the mirror");
    AppendLog(L"Decision mirror queue order: a shown decision stays until answered; the oldest pending decision is mirrored next.");
    return true;
}

// R8.13, F1-F4, F11 and the key guard at the overlay level: the modeless presentation the mirror is built on.
[[nodiscard]] bool RunFileOpsAlertOverlayModelessChecks(SelfTestState& state) noexcept
{
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
    {
        AppendLog(L"Alert overlay modeless: no-activation guard unavailable");
        return false;
    }
    RedSalamander::Ui::AlertOverlayModelessSelfTestResult result{};
    if (! RedSalamander::Ui::DebugRunAlertOverlayModelessSelfTest(state.mainWindow, result))
    {
        AppendLog(L"Alert overlay modeless: the fixture could not run");
        return false;
    }
    uint32_t failures  = 0u;
    const auto require = [&](bool condition, std::wstring_view what) noexcept
    {
        if (! condition)
        {
            ++failures;
            AppendLog(std::format(L"Alert overlay modeless: {}", what));
        }
    };
    require(result.takeFocusCapturesAnchor, L"a modeless overlay shown to take the focus did not record the anchor as the focus to give back");
    require(result.takeFocusMovesFocus, L"a modeless overlay shown to take the focus did not take it");
    require(result.hideRestoresFocus, L"hiding the focused modeless overlay did not give the focus back to the anchor");
    require(result.unfocusedEscapeIgnored, L"an unfocused modeless overlay acted on a routed Escape");
    require(result.unfocusedEnterIgnored, L"an unfocused modeless overlay acted on Enter");
    require(result.focusedEscapeRunsBinding, L"Escape on the focused overlay did not run only its escape binding");
    require(result.closeRunsOnCloseNotBinding, L"the close glyph ran the escape binding instead of onClose");
    require(result.firstShowAnnounces, L"the first show did not raise the owner's UI Automation announcement");
    require(result.followsRootMove, L"the overlay did not follow a move of the anchor's top-level window");
    require(result.hidesWithHiddenAnchor && result.showsWithShownAnchor, L"the overlay did not hide and show with its anchor");
    require(result.destroyedAnchorHides && result.destroyedAnchorReported, L"a destroyed pane anchor left the overlay visible or its owner uninformed");
    require(result.guardKeepsFocusOnDefault, L"a guarded focus let Tab, an arrow or Space move the focus or press a button");
    require(result.guardIgnoresRepeat, L"a guarded overlay acted on an auto-repeated key");
    require(result.guardIgnoresTypeahead, L"a guarded overlay acted on a key typed before or just after it took the focus");
    require(result.guardIgnoresHeldKey, L"a guarded overlay acted on an Escape held since before the show, or not after its release");
    require(result.guardEnterSubmitsDefault, L"Enter on a guarded overlay did not submit only the default");
    require(result.clickLiftsGuard, L"a click on the overlay did not lift the key guard");
    require(result.unguardedTabReachesClose, L"Tab on an unguarded modeless overlay did not reach the close glyph");
    require(result.repeatIgnoredWhenUnguarded, L"an unguarded modeless overlay acted on an auto-repeated key");
    require(result.ownedMenuKeepsReturn, L"the overlay forgot its focus-return target when its own menu took the focus");
    require(result.focusLossForgetsReturn, L"the overlay kept its focus-return target after the focus went to another window");
    require(result.optionChangeReported && result.optionValueSetInPlace,
            L"an option change did not reach the owner, or the owner's value was not shown in place");
    require(result.hiddenReturnFallsBackToAnchor, L"a hidden focus-return target did not fall back to the alert's anchor");
    require(result.blockingHideFallsBackToAnchor,
            L"a blocking alert whose captured focus window was hidden did not give the focus to its own anchor when it hid");
    require(result.closeMessageForwardedToOwner, L"Alt+F4 on the modeless overlay did not go to its owner's top-level window");
    require(result.f6IgnoresRepeatTypeaheadAndHeld, L"a guarded overlay acted on an F6 held since the show, typed within the typeahead window or repeated");
    require(result.f6ReturnsFocusToPane, L"a fresh F6 did not hand the focus back to the pane without answering, the overlay staying shown");
    require(result.shiftF6ReturnsFocusToPane, L"Shift+F6 did not hand the focus back to the pane like F6");
    require(result.ctrlTabIgnoresRepeatTypeaheadAndHeld,
            L"a guarded overlay acted on a Ctrl+Tab typed within the typeahead window, repeated, or with its Ctrl held since the focus arrived");
    require(result.ctrlTabLiftsGuardToDefault, L"a fresh Ctrl+Tab did not lift the guard with the focus ring on the default");
    require(result.tabAfterCtrlTabLeavesDefault, L"Tab after Ctrl+Tab did not move the focus ring off the default");
    require(result.focusGainArmsGuard, L"a focus handed to the overlay later without a click did not arm the guard and pin the focus ring to the default");
    require(result.ownedMenuFocusReturnKeepsGuard, L"a focus coming back from the overlay's own menu armed the guard again");
    require(result.plainOverlayKeepsRepeatAndClose, L"a non-modal alert that does not require the focus lost its repeated-key action or its own Alt+F4 close");
    require(result.trackingMoveKeepsModelessZOrder, L"a move of the tracked top-level window raised the modeless overlay in the z-order");
    require(result.trackingMoveRaisesBlocking, L"a move of the tracked top-level window did not raise the blocking alert again");
    require(result.foregroundRequests == 0u, std::format(L"the modeless overlay requested the foreground {} time(s)", result.foregroundRequests));
    if (failures != 0u)
        return false;
    AppendLog(L"Alert overlay modeless: focus taken and returned, guarded keys (repeat, typeahead, held keys, Tab/arrows/Space inert, Enter to "
              L"the default), click lifts the guard, F6/Shift+F6 back to the pane and Ctrl+Tab lifting the guard (fresh only), a focus gain "
              L"guarded again and one back from its own menu not, Tab reaches close, focus loss forgets the return, options reported and set in "
              L"place, close is not Escape, Alt+F4 to the owner, a non-modal alert keeps its repeat and close, the blocking-hide fallback, "
              L"first-show announcement, root and visibility tracking with the z-order kept (modeless) or raised (blocking), destroyed pane "
              L"anchor.");
    return true;
}
