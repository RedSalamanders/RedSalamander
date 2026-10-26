// Included by the FileOps runner: exact Local deletion truth and live task counters, with no activation.
[[nodiscard]] bool RunFileOpsPermanentDeleteProgressChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Permanent Delete completion progress: {}", reason));
        return false;
    };
    if (! state.fileOps || ! state.fsLocal)
        return reject(L"Local provider or task owner unavailable");
    wil::com_ptr<IFileSystemObjectBinding> binding;
    if (! state.fsLocal.try_query_to(binding.put()))
        return reject(L"Local exact-object binding unavailable");
    const auto fixture = state.tempRoot / L"permanent-delete-completion-progress";
    if (! RecreateEmptyDirectory(fixture))
        return reject(L"owned fixture unavailable");
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(fixture)); });
    const auto newTask = [&](const std::filesystem::path& source) noexcept
    {
        auto task            = std::make_unique<Task>(*state.fileOps);
        task->_operation     = FILESYSTEM_DELETE;
        task->_sourcePaths   = {source};
        task->_executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        task->BeginDiscovery();
        task->SkipDiscovery();
        return task;
    };
    // Abort before the Shell can run: only the actual provider's route-certification report is
    // exercised. Recycle cannot certify exact removal counters without per-object observations.
    const auto recycleRetained = fixture / L"recycle-not-started.bin";
    if (! WriteFilledTestFile(recycleRetained, 5u, 0x44u))
        return reject(L"owned canceled Recycle fixture unavailable");
    auto recycleTask = newTask(recycleRetained);
    FileSystemOptions recycleOptions{};
    recycleTask->InitializeFileSystemOptions(recycleOptions, nullptr);
    recycleTask->RequestCancel();
    const HRESULT recycleHr = [&]() noexcept
    {
        recycleTask->_dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
        const auto callbackScope = wil::scope_exit([&]() noexcept { recycleTask->_dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
        return state.fsLocal->DeleteItem(recycleRetained.c_str(), FILESYSTEM_FLAG_USE_RECYCLE_BIN, &recycleOptions, nullptr, nullptr);
    }();
    if (recycleHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || recycleTask->_completedWorkCountsKnown.load(std::memory_order_acquire) ||
        recycleTask->_completedWorkFileCount.load(std::memory_order_acquire) != 0u || GetFileAttributesW(recycleRetained.c_str()) == INVALID_FILE_ATTRIBUTES)
        return reject(L"Recycle certified unobserved removal counts or crossed its canceled operation control");
    constexpr FileSystemBindFlags deleteBinding =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_DELETE);
    const auto partialRoot = fixture / L"readonly-subtree";
    const auto ordinary    = partialRoot / L"ordinary.bin";
    const auto readonly    = partialRoot / L"readonly.bin";
    if (! WriteFilledTestFile(ordinary, 17u, 0x31u) || ! WriteFilledTestFile(readonly, 29u, 0x92u) ||
        ! SetFileAttributesW(readonly.c_str(), FILE_ATTRIBUTE_READONLY))
        return reject(L"readonly fixture unavailable");
    auto partialTask = newTask(partialRoot);
    FileSystemOptions options{};
    partialTask->InitializeFileSystemOptions(options, nullptr);
    wil::com_ptr<IFileSystemBoundObject> retained;
    if (FAILED(binding->BindObject(partialRoot.c_str(), deleteBinding, retained.put())) || ! retained)
        return reject(L"exact selected directory authority unavailable");
    FileSystemConditionalMutationResult receipt{};
    receipt.sizeBytes            = sizeof(receipt);
    const auto recursiveContinue = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    const auto startedAt         = std::chrono::steady_clock::now();
    const HRESULT partialHr      = [&]() noexcept
    {
        partialTask->_dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
        const auto callbackScope = wil::scope_exit([&]() noexcept { partialTask->_dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
        return retained->DeleteIfUnchanged(recursiveContinue, &options, &receipt);
    }();
    if (partialHr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) || receipt.outcomeKnown != TRUE || receipt.mutationCommitted != FALSE ||
        receipt.originalStillPresent != TRUE || ! partialTask->_discoverySkipped.load(std::memory_order_acquire) ||
        ! partialTask->_completedWorkCountsKnown.load(std::memory_order_acquire) ||
        partialTask->_completedWorkFileCount.load(std::memory_order_acquire) != 1u ||
        partialTask->_completedWorkFolderCount.load(std::memory_order_acquire) != 0u ||
        partialTask->_completedWorkBytes.load(std::memory_order_acquire) != 17u || GetFileAttributesW(ordinary.c_str()) != INVALID_FILE_ATTRIBUTES)
        return reject(L"partial root did not retain exact one-file removal progress while counting was skipped");
    const DWORD retainedAttributes = GetFileAttributesW(readonly.c_str());
    if (retainedAttributes == INVALID_FILE_ATTRIBUTES || (retainedAttributes & FILE_ATTRIBUTE_READONLY) == 0u)
        return reject(L"unconfirmed read-only child was removed or its attributes changed");
    receipt                   = {sizeof(receipt), FALSE, FALSE, TRUE};
    const HRESULT confirmedHr = [&]() noexcept
    {
        partialTask->_dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
        const auto callbackScope = wil::scope_exit([&]() noexcept { partialTask->_dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
        return retained->DeleteIfUnchanged(static_cast<FileSystemFlags>(recursiveContinue | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY), &options, &receipt);
    }();
    if (confirmedHr != S_OK || receipt.outcomeKnown != TRUE || receipt.mutationCommitted != TRUE || receipt.originalStillPresent != FALSE ||
        GetFileAttributesW(partialRoot.c_str()) != INVALID_FILE_ATTRIBUTES || partialTask->_completedWorkFileCount.load(std::memory_order_acquire) != 2u ||
        partialTask->_completedWorkFolderCount.load(std::memory_order_acquire) != 1u || partialTask->_completedWorkBytes.load(std::memory_order_acquire) != 46u)
        return reject(L"explicit read-only modifier did not delete the retained subtree with exact cumulative committed counters");
    retained.reset();

    // Refused disposition after clearing readonly must restore it; no object has completed.
    const auto nonempty = fixture / L"nonempty-readonly";
    const auto survivor = nonempty / L"survivor.bin";
    if (! WriteFilledTestFile(survivor, 13u, 0xA6u) || ! SetFileAttributesW(nonempty.c_str(), FILE_ATTRIBUTE_READONLY))
        return reject(L"nonempty readonly fixture unavailable");
    auto refusedTask = newTask(nonempty);
    refusedTask->InitializeFileSystemOptions(options, nullptr);
    const HRESULT refusedHr = [&]() noexcept
    {
        refusedTask->_dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
        const auto callbackScope = wil::scope_exit([&]() noexcept { refusedTask->_dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
        return state.fsLocal->DeleteItem(nonempty.c_str(), FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY, &options, nullptr, nullptr);
    }();
    const DWORD restoredAttributes = GetFileAttributesW(nonempty.c_str());
    if (refusedHr != HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY) || restoredAttributes == INVALID_FILE_ATTRIBUTES ||
        (restoredAttributes & FILE_ATTRIBUTE_READONLY) == 0u || GetFileAttributesW(survivor.c_str()) == INVALID_FILE_ATTRIBUTES ||
        refusedTask->_completedWorkFileCount.load(std::memory_order_acquire) != 0u ||
        refusedTask->_completedWorkFolderCount.load(std::memory_order_acquire) != 0u || refusedTask->_completedWorkBytes.load(std::memory_order_acquire) != 0u)
        return reject(L"refused nonempty directory removal changed readonly state or fabricated completed work");

    // A committed child stays counted when cancellation stops the next mutation. This raw observer
    // requests cancellation through the independent operation-control channel, never by its HRESULT.
    const auto canceledRoot = fixture / L"cancel-after-child";
    if (! WriteFilledTestFile(canceledRoot / L"a.bin", 7u, 0x11u) || ! WriteFilledTestFile(canceledRoot / L"b.bin", 7u, 0x22u) ||
        ! WriteFilledTestFile(canceledRoot / L"c.bin", 7u, 0x33u))
        return reject(L"cancel fixture unavailable");
    auto canceledTask = newTask(canceledRoot);
    struct CancelObserver final : IFileSystemFileWorkCallback
    {
        Task& task;
        explicit CancelObserver(Task& owner) noexcept : task(owner)
        {
        }
        CancelObserver(const CancelObserver&)            = delete;
        CancelObserver(CancelObserver&&)                 = delete;
        CancelObserver& operator=(const CancelObserver&) = delete;
        CancelObserver& operator=(CancelObserver&&)      = delete;
        HRESULT STDMETHODCALLTYPE FileSystemReportFileWork(const FileSystemFileWorkProgress* progress, void* cookie) noexcept override
        {
            const HRESULT hr = task.FileSystemReportFileWork(progress, cookie);
            if (progress != nullptr && progress->phase == FileSystemFileWorkPhase::ObjectCompleted)
                task.RequestCancel();
            return FAILED(hr) ? hr : E_ABORT; // ignored: observations cannot undo a committed deletion
        }
    } observer(*canceledTask);
    canceledTask->InitializeFileSystemOptions(options, nullptr);
    options.fileWorkCallback = &observer;
    if (FAILED(binding->BindObject(canceledRoot.c_str(), deleteBinding, retained.put())) || ! retained)
        return reject(L"cancel exact authority unavailable");
    receipt                  = {sizeof(receipt), FALSE, FALSE, TRUE};
    const HRESULT canceledHr = [&]() noexcept
    {
        canceledTask->_dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
        const auto callbackScope = wil::scope_exit([&]() noexcept { canceledTask->_dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
        return retained->DeleteIfUnchanged(FILESYSTEM_FLAG_RECURSIVE, &options, &receipt);
    }();
    if (canceledHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || receipt.outcomeKnown != TRUE || receipt.mutationCommitted != FALSE ||
        receipt.originalStillPresent != TRUE || canceledTask->_completedWorkFileCount.load(std::memory_order_acquire) != 1u ||
        canceledTask->_completedWorkFolderCount.load(std::memory_order_acquire) != 0u ||
        canceledTask->_completedWorkBytes.load(std::memory_order_acquire) != 7u || GetFileAttributesW(canceledRoot.c_str()) == INVALID_FILE_ATTRIBUTES)
        return reject(L"cancel did not preserve exactly the committed child and retained root facts");
    retained.reset();

    const auto copySource      = fixture / L"copy-source";
    const auto copyDestination = fixture / L"copy-destination";
    std::error_code ec;
    std::filesystem::create_directories(copySource / L"empty", ec);
    if (ec || ! WriteFilledTestFile(copySource / L"nested" / L"leaf.bin", 11u, 0x52u) || ! WriteFilledTestFile(copySource / L"root.bin", 19u, 0x39u))
        return reject(L"copy directory completion fixture unavailable");
    auto copiedTask            = std::make_unique<Task>(*state.fileOps);
    copiedTask->_operation     = FILESYSTEM_COPY;
    copiedTask->_sourcePaths   = {copySource};
    copiedTask->_executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
    copiedTask->BeginDiscovery();
    copiedTask->SkipDiscovery();
    copiedTask->InitializeFileSystemOptions(options, nullptr);
    const HRESULT copyHr = [&]() noexcept
    {
        copiedTask->_dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
        const auto callbackScope = wil::scope_exit([&]() noexcept { copiedTask->_dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
        return state.fsLocal->CopyItem(copySource.c_str(), copyDestination.c_str(), FILESYSTEM_FLAG_RECURSIVE, &options, nullptr, nullptr);
    }();
    if (copyHr != S_OK || ! copiedTask->_copiedFileCountKnown.load(std::memory_order_acquire) ||
        copiedTask->_copiedFileCount.load(std::memory_order_acquire) != 2u || ! copiedTask->_copiedFolderCountKnown.load(std::memory_order_acquire) ||
        copiedTask->_copiedFolderCount.load(std::memory_order_acquire) != 3u)
        return reject(L"native copy did not count two published files and three completed folders after Skip counting");
    // No link creation: synthetic delivery exercises the same reducer used by both Native and
    // bridge literal directory-link publication, including duplicate and late terminal reports.
    auto linkTask            = std::make_unique<Task>(*state.fileOps);
    linkTask->_operation     = FILESYSTEM_COPY;
    linkTask->_sourcePaths   = {copySource};
    linkTask->_executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
    linkTask->BeginDiscovery();
    FileSystemFileWorkProgress link{};
    link.sizeBytes = sizeof(link);
    link.phase     = FileSystemFileWorkPhase::ScopeSupported;
    link.kind      = FileSystemFileWorkKind::Directory;
    if (FAILED(linkTask->FileSystemReportFileWork(&link, nullptr)))
        return reject(L"directory-link reducer scope rejected");
    link.kind             = FileSystemFileWorkKind::DirectoryLink;
    link.phase            = FileSystemFileWorkPhase::Begin;
    link.sourcePath       = copySource.c_str();
    link.logicalLeafId    = 1u;
    link.progressStreamId = FILESYSTEM_PRODUCER_PROGRESS_STREAM_ID;
    link.sizeKnown        = TRUE;
    if (FAILED(linkTask->FileSystemReportFileWork(&link, nullptr)))
        return reject(L"directory-link reducer begin rejected");
    link.phase = FileSystemFileWorkPhase::Published;
    if (FAILED(linkTask->FileSystemReportFileWork(&link, nullptr)) || FAILED(linkTask->FileSystemReportFileWork(&link, nullptr)) ||
        ! linkTask->_copiedFolderCountKnown.load(std::memory_order_acquire) || linkTask->_copiedFolderCount.load(std::memory_order_acquire) != 1u ||
        linkTask->_copiedFileCount.load(std::memory_order_acquire) != 0u)
        return reject(L"directory-link publication was omitted or counted twice/as a regular file");
    link.phase       = FileSystemFileWorkPhase::Terminal;
    link.disposition = FileSystemFileWorkDisposition::Completed;
    if (FAILED(linkTask->FileSystemReportFileWork(&link, nullptr)) || linkTask->_copiedFolderCount.load(std::memory_order_acquire) != 1u)
        return reject(L"directory-link terminal changed its already-counted publication");
    link           = {};
    link.sizeBytes = sizeof(link);
    link.phase     = FileSystemFileWorkPhase::ScopeSupported;
    link.kind      = FileSystemFileWorkKind::Directory;
    if (FAILED(linkTask->FileSystemReportFileWork(&link, nullptr)) || linkTask->_copiedFolderCountKnown.load(std::memory_order_acquire))
        return reject(L"replaying a directory-link scope retained count certainty without stable leaf identity");
    auto malformedDirectoryTask          = std::make_unique<Task>(*state.fileOps);
    malformedDirectoryTask->_operation   = FILESYSTEM_COPY;
    malformedDirectoryTask->_sourcePaths = {copySource};
    malformedDirectoryTask->BeginDiscovery();
    if (FAILED(malformedDirectoryTask->FileSystemReportFileWork(&link, nullptr)))
        return reject(L"ordinary-directory validation scope rejected");
    link.phase         = FileSystemFileWorkPhase::Begin;
    link.logicalLeafId = 1u;
    link.sourcePath    = copySource.c_str();
    if (malformedDirectoryTask->FileSystemReportFileWork(&link, nullptr) != E_INVALIDARG ||
        malformedDirectoryTask->_copiedFileCountKnown.load(std::memory_order_acquire) ||
        malformedDirectoryTask->_copiedFolderCountKnown.load(std::memory_order_acquire))
        return reject(L"ordinary directory entered the copied-file lifecycle without withdrawing count certainty");
    link.phase = FileSystemFileWorkPhase::Published;
    if (malformedDirectoryTask->FileSystemReportFileWork(&link, nullptr) != E_INVALIDARG ||
        malformedDirectoryTask->_copiedFileCount.load(std::memory_order_acquire) != 0u ||
        malformedDirectoryTask->_copiedFolderCount.load(std::memory_order_acquire) != 0u)
        return reject(L"ordinary directory lifecycle publication fabricated copied files/folders");
    // A native permanent Delete reports both committed-object observations and legacy cumulative
    // progress. Only the observer counts completed work while discovery remains open.
    auto metricsTask            = newTask(copySource);
    metricsTask->_executionMode = FolderWindow::FileOperationState::ExecutionMode::BulkItems;
    FileSystemFileWorkProgress completion{};
    completion.sizeBytes = sizeof(completion);
    completion.phase     = FileSystemFileWorkPhase::ScopeSupported;
    completion.kind      = FileSystemFileWorkKind::Directory;
    if (FAILED(metricsTask->FileSystemReportFileWork(&completion, nullptr)))
        return reject(L"native Delete metric scope rejected");
    completion.phase          = FileSystemFileWorkPhase::ObjectCompleted;
    completion.kind           = FileSystemFileWorkKind::RegularFile;
    completion.disposition    = FileSystemFileWorkDisposition::Completed;
    completion.sourcePath     = copySource.c_str();
    completion.sizeKnown      = TRUE;
    completion.totalBytes     = 31u;
    completion.completedBytes = 31u;
    if (FAILED(metricsTask->FileSystemReportFileWork(&completion, nullptr)))
        return reject(L"native Delete metric object rejected");
    metricsTask->_dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_acq_rel);
    const auto leaveCallback = wil::scope_exit([&]() noexcept { metricsTask->_dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_acq_rel); });
    if (FAILED(metricsTask->FileSystemProgress(FILESYSTEM_DELETE, 0u, 1u, 0u, 31u, copySource.c_str(), nullptr, 0u, 0u, nullptr, 1u, nullptr)) ||
        metricsTask->_discoveryCompletedBytesWhileOpen.load(std::memory_order_acquire) != 31u ||
        metricsTask->_discoveryCompletedMutationsWhileOpen.load(std::memory_order_acquire) != 1u)
        return reject(L"legacy cumulative Delete progress counted an observed object twice");
    completion           = {};
    completion.sizeBytes = sizeof(completion);
    completion.phase     = FileSystemFileWorkPhase::CountsUnavailable;
    if (FAILED(metricsTask->FileSystemReportFileWork(&completion, nullptr)) ||
        FAILED(metricsTask->FileSystemProgress(FILESYSTEM_DELETE, 0u, 2u, 0u, 62u, copySource.c_str(), nullptr, 0u, 0u, nullptr, 1u, nullptr)) ||
        metricsTask->_discoveryCompletedBytesWhileOpen.load(std::memory_order_acquire) != 62u ||
        metricsTask->_discoveryCompletedMutationsWhileOpen.load(std::memory_order_acquire) != 2u)
        return reject(L"unsupported Delete observer suppressed the legacy progress metric fallback");
    auto malformedLinkTask = newTask(copySource);
    completion             = {};
    completion.sizeBytes   = sizeof(completion);
    completion.phase       = FileSystemFileWorkPhase::ScopeSupported;
    completion.kind        = FileSystemFileWorkKind::Directory;
    if (FAILED(malformedLinkTask->FileSystemReportFileWork(&completion, nullptr)))
        return reject(L"file-link removal validation scope rejected");
    completion.phase          = FileSystemFileWorkPhase::ObjectCompleted;
    completion.kind           = FileSystemFileWorkKind::FileLink;
    completion.disposition    = FileSystemFileWorkDisposition::Completed;
    completion.sourcePath     = copySource.c_str();
    completion.sizeKnown      = TRUE;
    completion.totalBytes     = 1u;
    completion.completedBytes = 1u;
    if (malformedLinkTask->FileSystemReportFileWork(&completion, nullptr) != E_INVALIDARG ||
        malformedLinkTask->_completedWorkCountsKnown.load(std::memory_order_acquire) ||
        malformedLinkTask->_completedWorkFileCount.load(std::memory_order_acquire) != 0u ||
        malformedLinkTask->_completedWorkBytes.load(std::memory_order_acquire) != 0u)
        return reject(L"malformed link-object payload fabricated deleted bytes or retained count certainty");
    Debug::Perf::Emit(L"FileOps.SelfTest.PermanentDeleteCompletionProgressUs",
                      L"readonly-confirmation,partial,retry,cancel",
                      Debug::Perf::ElapsedUs(startedAt),
                      3u,
                      53u,
                      S_OK);
    AppendLog(L"Permanent Delete completion progress: partial=1 file/17 B; confirmed=2 files/1 folder/46 B; canceled=1 file/7 B; refused=0. Readonly retained "
              L"on refusal.");
    return true;
}

[[nodiscard]] bool RunFileOpsReadonlyDeleteUserConsentChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Readonly Delete user consent: {}", reason));
        return false;
    };
    auto* operations = state.fileOps;
    if (! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"Local owner quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground = GetForegroundWindow();
    const auto root       = state.tempRoot / L"readonly-delete-user-consent";
    const auto source     = root / L"ordinary-parent";
    const auto readonly   = source / L"readonly.bin";
    if (! RecreateEmptyDirectory(root) || ! WriteFilledTestFile(source / L"ordinary.bin", 17u, 0x37u) || ! WriteFilledTestFile(readonly, 29u, 0xA3u) ||
        ! SetFileAttributesW(readonly.c_str(), FILE_ATTRIBUTE_READONLY))
        return reject(L"owned normal-parent/readonly-child fixture unavailable");
    const bool autoDismiss    = operations->GetAutoDismissSuccess();
    const bool autoAccept     = HostGetAutoAcceptPrompts();
    const auto promptOverride = HostGetTestPromptResultOverride();
    uint64_t taskId           = 0u;
    bool quiet                = true;
    const auto cleanup        = wil::scope_exit([&]() noexcept
    {
        quiet = CancelOutstandingSelfTestOperations(operations);
        if (quiet && taskId != 0u)
        {
            const auto retired = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                return ! operations->FindTask(taskId);
            }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"readonly Delete fixture retirement"});
            quiet              = retired.conditionMet;
        }
        if (quiet)
        {
            operations->DismissCompletedTask(taskId);
            operations->DebugRemoveDiagnosticsForTask(taskId);
            state.completedTasks.erase(taskId);
            operations->DebugResetPopupForSelfTest();
            static_cast<void>(SelfTest::RemoveAll(root));
        }
        else
            AppendLog(L"Readonly Delete fixture retained because its worker did not reach a quiet point.");
        operations->SetAutoDismissSuccess(autoDismiss);
        HostSetAutoAcceptPrompts(autoAccept);
        if (promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(promptOverride);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
    });
    operations->SetAutoDismissSuccess(false);
    HostSetAutoAcceptPrompts(false);
    HostClearTestPromptResultOverride();
    const auto admitted = StartFileOperationAndGetId(
        operations, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {source}, {}, FILESYSTEM_FLAG_RECURSIVE, false);
    if (! admitted.has_value())
        return reject(L"normal permanent Delete admission failed");
    taskId = admitted.value();
    std::optional<Task::ConflictPromptState> prompt;
    const auto initial = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        prompt = TryGetConflictPromptCopy(operations->FindTask(taskId));
        return prompt.has_value();
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"readonly Delete initial confirmation"});
    auto* task = operations->FindTask(taskId);
    if (! initial.conditionMet || ! task || prompt.value().bucket != Task::ConflictBucket::PermanentDeleteConfirmation ||
        ! task->SubmitConflictDecision(prompt.value().generation, Task::ConflictAction::PermanentDelete, false))
        return reject(L"initial permanent Delete confirmation unavailable/rejected");
    const uint64_t initialGeneration = prompt.value().generation;
    task                             = nullptr;
    const auto readonlyQuestion      = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        prompt = TryGetConflictPromptCopy(operations->FindTask(taskId));
        return prompt.has_value() && prompt.value().generation != initialGeneration;
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"readonly subtree explicit authorization"});
    task                      = operations->FindTask(taskId);
    const DWORD beforeConsent = GetFileAttributesW(readonly.c_str());
    if (! readonlyQuestion.conditionMet || ! task || prompt.value().sourcePath != source.native() ||
        ! PromptHasAction(prompt.value(), Task::ConflictAction::DeleteReadOnly) || prompt.value().applyToAllEligible ||
        beforeConsent == INVALID_FILE_ATTRIBUTES || (beforeConsent & FILE_ATTRIBUTE_READONLY) == 0u)
        return reject(L"actual readonly-child evidence did not produce an explicit exact-root nonrepeatable DeleteReadOnly action");
    task->SkipDiscovery();
    if (! task->SubmitConflictDecision(prompt.value().generation, Task::ConflictAction::DeleteReadOnly, false))
        return reject(L"explicit DeleteReadOnly answer was rejected");
    task                 = nullptr;
    const auto completed = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return state.completedTasks.contains(taskId) && ! operations->FindTask(taskId);
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"readonly Delete completion"});
    if (! completed.conditionMet || GetFileAttributesW(source.c_str()) != INVALID_FILE_ATTRIBUTES || GetForegroundWindow() != foreground)
        return reject(L"confirmed subtree did not disappear at quiet point or the foreground changed");
    const auto completion = state.completedTasks.find(taskId);
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
    operations->CollectCompletedTasks(summaries);
    const auto summary = std::ranges::find_if(summaries, [taskId](const auto& value) noexcept { return value.taskId == taskId; });
    if (completion == state.completedTasks.end() || completion->second.hr != S_OK || summary == summaries.end() || summary->resultHr != S_OK ||
        ! summary->completedWorkCountsKnown || summary->completedWorkFileCount != 2u || summary->completedWorkFolderCount != 1u ||
        summary->completedWorkBytes != 46u || ! summary->discoverySkipped)
        return reject(L"completed task lost successful exact 2-file/1-folder/46-byte removal accounting");
    AppendLog(L"Readonly Delete user consent: normal parent/readonly child; separate initial permanent and explicit readonly confirmations; retained before "
              L"answer; exact completed2 files/1 folder/46 B; foreground preserved.");
    return true;
}

[[nodiscard]] bool RunFileOpsBridgeCompletedFolderChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Bridge completed folders: {}", reason));
        return false;
    };
    auto* operations = state.fileOps;
    if (! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"Local owner quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground  = GetForegroundWindow();
    const auto root        = state.tempRoot / L"bridge-completed-folders";
    const auto source      = root / L"source" / L"tree";
    const bool verify      = EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy;
    const bool autoDismiss = operations->GetAutoDismissSuccess();
    uint64_t taskId        = 0u;
    const auto restore     = wil::scope_exit([&]() noexcept
    {
        SetFileOpsBridgeDirectoryRestorePauseForSelfTest({});
        const bool quiet = CancelOutstandingSelfTestOperations(operations);
        if (quiet)
        {
            if (taskId != 0u)
            {
                operations->DismissCompletedTask(taskId);
                operations->DebugRemoveDiagnosticsForTask(taskId);
                state.completedTasks.erase(taskId);
            }
            static_cast<void>(SelfTest::RemoveAll(root));
        }
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = verify;
        operations->SetAutoDismissSuccess(autoDismiss);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
    });
    operations->SetAutoDismissSuccess(false);
    if (! RecreateEmptyDirectory(root) || ! WriteFilledTestFile(source / L"root.bin", 19u, 0x61u) ||
        ! WriteFilledTestFile(source / L"nested" / L"leaf.bin", 11u, 0xD3u))
        return reject(L"owned bridge source unavailable");
    std::error_code ec;
    std::filesystem::create_directories(source / L"empty", ec);
    if (ec)
        return reject(L"owned empty source folder unavailable");
    for (const bool sequential : {false, true})
    {
        const auto destination = root / (sequential ? L"sequential" : L"parallel");
        if (! RecreateEmptyDirectory(destination))
            return reject(L"owned bridge destination unavailable");
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = sequential;
        SetFileOpsBridgeDirectoryRestorePauseForSelfTest((destination / L"tree" / L"nested").native());
        const auto admitted = StartFileOperationAndGetId(operations,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         {source},
                                                         destination,
                                                         FILESYSTEM_FLAG_RECURSIVE,
                                                         false,
                                                         0u,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsLocal);
        if (! admitted.has_value())
            return reject(L"forced bridge Copy admission failed");
        taskId          = admitted.value();
        const auto held = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return HasFileOpsBridgeDirectoryRestorePauseEnteredForSelfTest() || state.completedTasks.contains(taskId);
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"bridge folder completion hold"});
        if (! held.conditionMet || ! HasFileOpsBridgeDirectoryRestorePauseEnteredForSelfTest())
            return reject(L"forced bridge did not reach its real post-child folder boundary");
        ReleaseFileOpsBridgeDirectoryRestorePauseForSelfTest();
        const auto completed = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return state.completedTasks.contains(taskId) && ! operations->FindTask(taskId);
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"bridge completed folder counters"});
        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        operations->CollectCompletedTasks(summaries);
        const auto summary            = std::ranges::find_if(summaries, [taskId](const auto& value) noexcept { return value.taskId == taskId; });
        const auto info               = state.completedTasks.find(taskId);
        const bool rootPayloadMatches = FilesEqualBytes(source / L"root.bin", destination / L"tree" / L"root.bin");
        const bool leafPayloadMatches = FilesEqualBytes(source / L"nested" / L"leaf.bin", destination / L"tree" / L"nested" / L"leaf.bin");
        const HWND actualForeground   = GetForegroundWindow();
        if (! completed.conditionMet || info == state.completedTasks.end() || info->second.hr != S_OK ||
            (! sequential && info->second.bridgeDirectoryEnsureCount != 3u) || summary == summaries.end() || ! summary->copiedFileCountKnown ||
            summary->copiedFileCount != 2u || ! summary->copiedFolderCountKnown || summary->copiedFolderCount != 3u || ! rootPayloadMatches ||
            ! leafPayloadMatches || actualForeground != foreground)
            return reject(std::format(L"{}: completed={} info={} hr=0x{:08X} ensure={} summary={} filesKnown={} files={} foldersKnown={} folders={} "
                                      L"rootPayload={} leafPayload={} foreground=0x{:X} expected=0x{:X}",
                                      sequential ? L"sequential" : L"parallel",
                                      completed.conditionMet,
                                      info != state.completedTasks.end(),
                                      static_cast<uint32_t>(info != state.completedTasks.end() ? info->second.hr : E_PENDING),
                                      info != state.completedTasks.end() ? info->second.bridgeDirectoryEnsureCount : 0u,
                                      summary != summaries.end(),
                                      summary != summaries.end() && summary->copiedFileCountKnown,
                                      summary != summaries.end() ? summary->copiedFileCount : 0u,
                                      summary != summaries.end() && summary->copiedFolderCountKnown,
                                      summary != summaries.end() ? summary->copiedFolderCount : 0u,
                                      rootPayloadMatches,
                                      leafPayloadMatches,
                                      reinterpret_cast<uintptr_t>(actualForeground),
                                      reinterpret_cast<uintptr_t>(foreground)));
        operations->DismissCompletedTask(taskId);
        operations->DebugRemoveDiagnosticsForTask(taskId);
        state.completedTasks.erase(taskId);
        taskId = 0u;
        SetFileOpsBridgeDirectoryRestorePauseForSelfTest({});
    }
    AppendLog(L"Bridge completed folders: actual parallel and verified sequential routes each completed2 files/3 folders, retained summary counts, exact "
              L"payloads and foreground preserved.");
    return true;
}

[[nodiscard]] bool RunFileOpsSelectedReadonlyDeleteUserConsentChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Selected readonly Delete consent: {}", reason));
        return false;
    };
    auto* operations = state.fileOps;
    if (! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"Local owner quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground = GetForegroundWindow();
    const auto root       = state.tempRoot / L"selected-readonly-delete-consent";
    const auto file       = root / L"readonly.bin";
    const auto directory  = root / L"readonly-empty-folder";
    if (! RecreateEmptyDirectory(root) || ! WriteFilledTestFile(file, 17u, 0xC3u))
        return reject(L"owned readonly root fixture unavailable");
    std::error_code ec;
    std::filesystem::create_directory(directory, ec);
    if (ec || ! SetFileAttributesW(file.c_str(), FILE_ATTRIBUTE_READONLY) || ! SetFileAttributesW(directory.c_str(), FILE_ATTRIBUTE_READONLY))
        return reject(L"owned readonly root attributes unavailable");
    const bool autoDismiss    = operations->GetAutoDismissSuccess();
    const bool autoAccept     = HostGetAutoAcceptPrompts();
    const auto promptOverride = HostGetTestPromptResultOverride();
    uint64_t taskId           = 0u;
    const auto cleanup        = wil::scope_exit([&]() noexcept
    {
        if (CancelOutstandingSelfTestOperations(operations))
        {
            if (taskId != 0u)
            {
                operations->DismissCompletedTask(taskId);
                operations->DebugRemoveDiagnosticsForTask(taskId);
                state.completedTasks.erase(taskId);
            }
            static_cast<void>(SelfTest::RemoveAll(root));
        }
        operations->SetAutoDismissSuccess(autoDismiss);
        HostSetAutoAcceptPrompts(autoAccept);
        if (promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(promptOverride);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
    });
    operations->SetAutoDismissSuccess(false);
    HostSetAutoAcceptPrompts(false);
    HostClearTestPromptResultOverride();
    const auto admitted = StartFileOperationAndGetId(
        operations, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {file, directory}, {}, FILESYSTEM_FLAG_RECURSIVE, false);
    if (! admitted.has_value())
        return reject(L"two selected readonly roots were not admitted");
    taskId                      = admitted.value();
    uint64_t answeredGeneration = 0u;
    size_t initialConsents      = 0u;
    std::vector<std::wstring> confirmedReadonlyPaths;
    bool unexpected      = false;
    const auto completed = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        if (auto* task = operations->FindTask(taskId))
        {
            const auto prompt = TryGetConflictPromptCopy(task);
            if (prompt.has_value() && ! prompt->metadataLoading && prompt->buttonsPublishable && prompt->generation != answeredGeneration)
            {
                answeredGeneration = prompt->generation;
                if (prompt->bucket == Task::ConflictBucket::PermanentDeleteConfirmation)
                {
                    ++initialConsents;
                    unexpected = ! task->SubmitConflictDecision(prompt->generation, Task::ConflictAction::PermanentDelete, false);
                }
                else if (PromptHasAction(prompt.value(), Task::ConflictAction::DeleteReadOnly))
                {
                    const bool selectedRoot = prompt->sourcePath == file.native() || prompt->sourcePath == directory.native();
                    const DWORD attributes  = GetFileAttributesW(prompt->sourcePath.c_str());
                    unexpected              = ! selectedRoot || prompt->applyToAllEligible || attributes == INVALID_FILE_ATTRIBUTES ||
                                              (attributes & FILE_ATTRIBUTE_READONLY) == 0u ||
                                              std::ranges::find(confirmedReadonlyPaths, prompt->sourcePath) != confirmedReadonlyPaths.end();
                    if (! unexpected)
                    {
                        confirmedReadonlyPaths.push_back(prompt->sourcePath);
                        task->SkipDiscovery();
                        unexpected = ! task->SubmitConflictDecision(prompt->generation, Task::ConflictAction::DeleteReadOnly, false);
                    }
                }
                else
                    unexpected = true;
            }
        }
        return unexpected || (state.completedTasks.contains(taskId) && ! operations->FindTask(taskId));
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(12)), .operationName = L"selected readonly file/folder explicit consent"});
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
    operations->CollectCompletedTasks(summaries);
    const auto summary    = std::ranges::find_if(summaries, [taskId](const auto& value) noexcept { return value.taskId == taskId; });
    const auto completion = state.completedTasks.find(taskId);
    if (! completed.conditionMet || unexpected || initialConsents == 0u || confirmedReadonlyPaths.size() != 2u || completion == state.completedTasks.end() ||
        completion->second.hr != S_OK || summary == summaries.end() || ! summary->completedWorkCountsKnown || summary->completedWorkFileCount != 1u ||
        summary->completedWorkFolderCount != 1u || summary->completedWorkBytes != 17u || GetFileAttributesW(file.c_str()) != INVALID_FILE_ATTRIBUTES ||
        GetFileAttributesW(directory.c_str()) != INVALID_FILE_ATTRIBUTES || GetForegroundWindow() != foreground)
        return reject(L"each selected readonly file/empty folder did not require its own explicit consent with exact final removal counts");
    AppendLog(L"Selected readonly Delete consent: selected file and empty folder each retained readonly until separate explicit answers; initial permanent "
              L"consent alone did not delete either root; no grant leaked between roots; final1 file/1 folder/17 B.");
    return true;
}
