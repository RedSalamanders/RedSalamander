// Included from FolderWindow.FileOperations.Tests.Runner.cpp under ENABLE_TESTS.
// Admission/durable-owner contracts and private local execution witnesses; no foreground input.
[[nodiscard]] bool ValidateClipboardReadyStaleToken(SelfTestState& fixture,
                                                    SelfTest::CaseState& state,
                                                    const std::wstring_view sourcePath,
                                                    const std::wstring_view destinationFolder) noexcept
{
    if (! fixture.folderWindow || ! fixture.fileOps || ! fixture.fsDummy)
        return state.Require(false, L"Clipboard-ready token witness requires the existing FolderWindow and Dummy provider fixture.");
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> inheritedCompleted;
    std::vector<FolderWindow::InformationalTaskUpdate> inheritedInformational;
    fixture.fileOps->CollectCompletedTasks(inheritedCompleted);
    fixture.fileOps->CollectInformationalTasks(inheritedInformational);
    if (! state.Require(! fixture.fileOps->HasActiveOperations() && ! fixture.fileOps->GetPopupHwndForSelfTest() && inheritedCompleted.empty() &&
                            inheritedInformational.empty(),
                        std::format(L"Clipboard-ready token witness must start without inherited work, retained entries, or an application popup "
                                    L"(active={}, popup={}, completed={}, informational={}).",
                                    fixture.fileOps->HasActiveOperations(),
                                    fixture.fileOps->GetPopupHwndForSelfTest() != nullptr,
                                    inheritedCompleted.size(),
                                    inheritedInformational.size())))
        return false;

    // Accepted sequences remain protected after completion. Each invocation models a new
    // clipboard publication without changing the person's real clipboard.
    static std::atomic<uint32_t> nextSequence{0x7F01A001u};
    const uint32_t sequence = nextSequence.fetch_add(1u, std::memory_order_relaxed);
    const auto barrierCalls = std::make_shared<std::atomic<uint32_t>>(0u);
    const auto admit        = [&](HRESULT& admissionStatus) noexcept
    {
        return StartFileOperationAndGetId(fixture.fileOps,
                                          FILESYSTEM_MOVE,
                                          FolderWindow::Pane::Right,
                                          FolderWindow::Pane::Left,
                                          fixture.fsDummy,
                                          {std::filesystem::path(sourcePath)},
                                          std::filesystem::path(destinationFolder),
                                          FILESYSTEM_FLAG_NONE,
                                          false,
                                          0u,
                                          FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                          false,
                                          fixture.fsDummy,
                                          {},
                                          sequence,
                                          [barrierCalls]() noexcept
        {
            barrierCalls->fetch_add(1u, std::memory_order_release);
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED); // Release the worker without consuming or mutating.
        },
                                          {},
                                          &admissionStatus);
    };
    HRESULT admissionStatus = E_UNEXPECTED;
    const auto taskId       = admit(admissionStatus);
    if (! state.Require(
            taskId.has_value(),
            std::format(L"Clipboard-ready token witness must publish a pending clipboard Move admission (hr={:08X}).", static_cast<uint32_t>(admissionStatus))))
        return false;

    // This fixture admitted the only task into an idle state. Join before pumping: the real
    // completion handler removes and destroys Task. Dismiss only this task's retained result
    // through production cleanup, so the next family inherits neither work nor its popup.
    const auto finishOwnedTask = [&](uint64_t ownedTaskId, bool requireExpectedResult) noexcept -> bool
    {
        if (auto* ownedTask = fixture.fileOps->FindTask(ownedTaskId))
        {
            if (! ownedTask->_taskFinished.load(std::memory_order_acquire))
                ownedTask->RequestCancel();
            if (ownedTask->_thread.joinable())
                ownedTask->_thread.join();
        }
        const auto drained = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return fixture.fileOps->FindTask(ownedTaskId) == nullptr;
        }, {.timeout = SelfTest::Scale(std::chrono::milliseconds(5'000)), .operationName = L"Clipboard-ready fixture completion drain"});
        if (! drained.conditionMet)
            AppendLog(drained.timeoutDiagnostic);
        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completed;
        fixture.fileOps->CollectCompletedTasks(completed);
        const bool resultApplied = ! requireExpectedResult || std::ranges::any_of(completed, [&](const auto& result) noexcept {
            return result.taskId == ownedTaskId && result.resultHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        });
        fixture.fileOps->DismissCompletedTask(ownedTaskId);
        fixture.fileOps->DebugRemoveDiagnosticsForTask(ownedTaskId);
        fixture.fileOps->CollectCompletedTasks(completed);
        return drained.conditionMet && resultApplied && ! fixture.fileOps->HasActiveOperations() && ! fixture.fileOps->GetPopupHwndForSelfTest() &&
               std::ranges::none_of(completed, [&](const auto& result) noexcept { return result.taskId == ownedTaskId; });
    };
    auto cleanupTask = wil::scope_exit([&]() noexcept
    {
        if (! finishOwnedTask(taskId.value(), false))
        {
            AppendLog(L"Clipboard-ready fixture cleanup failed after an earlier assertion; retained state must not be treated as idle.");
            state.Require(false, L"Clipboard-ready fixture cleanup failed after an earlier assertion.");
        }
    });
    auto* const task = fixture.fileOps->FindTask(taskId.value());
    if (! state.Require(task != nullptr, L"Clipboard-ready token witness must find its published task."))
        return false;

    const ULONGLONG readyDeadline = GetTickCount64() + 5'000ull;
    while (! task->_debugWorkerStartGateWaiting.load(std::memory_order_acquire) && GetTickCount64() < readyDeadline)
        Sleep(1u);
    if (! state.Require(task->_debugWorkerStartGateWaiting.load(std::memory_order_acquire) &&
                            ! task->_clipboardMoveBarrierCompleted.load(std::memory_order_acquire) && ! task->_workerReleased.load(std::memory_order_acquire),
                        L"The clipboard Move worker must have posted its correlated ready payload and remain blocked on admission."))
    {
        task->RequestCancel();
        return false;
    }

    // Retire a real posted token through the same payload registry used by the FolderWindow
    // message handler. Replaying its now-unresolvable nonzero LPARAM with this pending task's ID
    // models an old/duplicate ready message arriving during a later admission.
    wil::unique_hwnd tokenWindow(CreateWindowExW(0u, L"STATIC", L"", 0u, 0, 0, 0, 0, HWND_MESSAGE, nullptr, GetModuleHandleW(nullptr), nullptr));
    if (tokenWindow)
        InitPostedPayloadWindow(tokenWindow.get());
    auto drainTokenWindow   = wil::scope_exit([&]() noexcept
    {
        if (tokenWindow)
            static_cast<void>(DrainPostedPayloadsForWindow(tokenWindow.get()));
    });
    auto stalePayload       = std::make_unique<FolderWindow::FileOperationState::ClipboardMoveReadyPayload>();
    stalePayload->taskId    = taskId.value();
    const UINT readyMessage = static_cast<UINT>(WndMsg::kFileOperationClipboardMoveReady);
    MSG staleMessage{};
    const bool stalePosted =
        tokenWindow &&
        PostMessagePayload(tokenWindow.get(), WndMsg::kFileOperationClipboardMoveReady, static_cast<WPARAM>(taskId.value()), std::move(stalePayload)) &&
        PeekMessageW(&staleMessage, tokenWindow.get(), readyMessage, readyMessage, PM_REMOVE) != FALSE &&
        DrainPostedPayloadsForWindow(tokenWindow.get()) == 1u && staleMessage.lParam != 0 &&
        ! TakeMessagePayload<FolderWindow::FileOperationState::ClipboardMoveReadyPayload>(staleMessage.hwnd, staleMessage.message, staleMessage.lParam);
    if (! state.Require(stalePosted, L"The clipboard-ready witness must retire a nonzero posted payload token before replay."))
    {
        task->RequestCancel();
        return false;
    }

    const HWND folderHwnd       = fixture.folderWindow->GetHwnd();
    const LRESULT staleDispatch = SendMessageW(folderHwnd, WndMsg::kFileOperationClipboardMoveReady, static_cast<WPARAM>(taskId.value()), staleMessage.lParam);
    const bool staleIgnored     = staleDispatch == 0 && ! task->_clipboardMoveBarrierCompleted.load(std::memory_order_acquire) &&
                                  ! task->_workerReleased.load(std::memory_order_acquire) && barrierCalls->load(std::memory_order_acquire) == 0u &&
                                  static_cast<bool>(task->_preWorkerReleaseBarrier);
    if (! state.Require(staleIgnored, L"An invalid nonzero token must not complete or cancel the matching pending clipboard admission."))
    {
        task->RequestCancel();
        return false;
    }

    // Deliver the worker's actual queued payload through the real HWND route. Its deliberately
    // failing test barrier proves the admission can still complete normally without mutation.
    MSG readyMessageRecord{};
    const bool readyQueued = PeekMessageW(&readyMessageRecord, folderHwnd, readyMessage, readyMessage, PM_REMOVE) != FALSE &&
                             static_cast<uint64_t>(readyMessageRecord.wParam) == taskId.value() && readyMessageRecord.lParam != 0;
    if (readyQueued)
        static_cast<void>(DispatchMessageW(&readyMessageRecord));

    const ULONGLONG completionDeadline = GetTickCount64() + 5'000ull;
    while (! task->_taskFinished.load(std::memory_order_acquire) && GetTickCount64() < completionDeadline)
        Sleep(1u);
    const bool validReadyCompleted = readyQueued && task->_taskFinished.load(std::memory_order_acquire) &&
                                     barrierCalls->load(std::memory_order_acquire) == 1u &&
                                     task->_clipboardMoveBarrierCompleted.load(std::memory_order_acquire) &&
                                     task->_clipboardMoveConsumptionStatus.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) &&
                                     ! task->_clipboardMoveConsumed.load(std::memory_order_acquire) &&
                                     task->_resultHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    if (! state.Require(validReadyCompleted, L"The valid ready token must still reach the barrier once and fail before mutation."))
        task->RequestCancel();
    const bool cleanupCompleted = finishOwnedTask(taskId.value(), validReadyCompleted);
    cleanupTask.release();
    state.Require(cleanupCompleted, L"Clipboard-ready fixture must apply its completion, dismiss its own result, and leave no work or popup.");
    if (! validReadyCompleted || ! cleanupCompleted)
        return false;

    HRESULT replayStatus = E_UNEXPECTED;
    const auto replay    = admit(replayStatus);
    const bool replayRejected =
        ! replay.has_value() && replayStatus == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && barrierCalls->load(std::memory_order_acquire) == 1u;
    const bool replayCleaned = ! replay.has_value() || finishOwnedTask(replay.value(), false);
    return state.Require(replayRejected && replayCleaned,
                         L"Replaying an accepted clipboard sequence must remain rejected after its owned result is dismissed.");
}

[[nodiscard]] bool ValidateQueuedAdmissionContracts(SelfTestState& fixture, SelfTest::CaseState& state) noexcept
{
    using State                         = FolderWindow::FileOperationState;
    const SelfTest::TestSandbox sandbox = SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::FileOperations, L"queued-admission-contracts");
    if (! state.Require(sandbox.IsValid() && fixture.fileOps && fixture.fsDummy, L"Queued admission contracts require the governed sandbox and Dummy fixture."))
        return false;
    // Predicate probes can publish diagnostics and presentation refreshes. Give them
    // a separate engine/owner so the subsequent real clipboard witness starts idle.
    Common::Settings::Settings privateSettings{};
    privateSettings.fileOperations.emplace();
    auto privateOwner = std::make_unique<FolderWindow>();
    privateOwner->SetSettings(&privateSettings);
    State isolatedOperations(*privateOwner);
    auto* const privateOperations              = &isolatedOperations;
    const std::filesystem::path root           = sandbox.root;
    const std::filesystem::path breadcrumbRoot = root / L"breadcrumb-state";
    FileOperationMoveBreadcrumb::SetRootForSelfTest(breadcrumbRoot);
    const auto restoreRoot = wil::scope_exit([]() noexcept { FileOperationMoveBreadcrumb::ClearRootForSelfTest(); });

    constexpr auto sourceFolder      = L"/queued-admission-contracts-source";
    constexpr auto sourcePath        = L"/queued-admission-contracts-source/single.txt";
    constexpr auto destinationFolder = L"/queued-admission-contracts-destination";
    const auto cleanupDummy          = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(fixture.fsDummy->DeleteItem(sourceFolder, FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr));
        static_cast<void>(fixture.fsDummy->DeleteItem(destinationFolder, FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr));
    });
    // These private fixture roots can survive an interrupted earlier process; seed from a known state.
    static_cast<void>(fixture.fsDummy->DeleteItem(sourceFolder, FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr));
    static_cast<void>(fixture.fsDummy->DeleteItem(destinationFolder, FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr));
    if (! state.Require(EnsureDummyFolderExists(fixture.fsDummy.get(), sourceFolder) && EnsureDummyFolderExists(fixture.fsDummy.get(), destinationFolder) &&
                            DummyWriteTextFile(fixture.fsDummy.get(), sourcePath, "unchanged-source"),
                        L"Captured admission fixture must create its provider-owned paths."))
        return false;

    State::OperationConstruction construction{};
    construction.operation                            = FILESYSTEM_MOVE;
    construction.sourcePane                           = FolderWindow::Pane::Right;
    construction.destinationPane                      = FolderWindow::Pane::Left;
    construction.fileSystem                           = fixture.fsDummy;
    construction.destinationFileSystem                = fixture.fsDummy;
    construction.sourcePaths                          = {std::filesystem::path(sourcePath)};
    construction.destinationFolder                    = destinationFolder;
    construction.sourcePluginId                       = kPluginIdDummy;
    construction.sourcePluginShortId                  = L"dummy";
    construction.sourceInstanceContext                = L"captured-context";
    construction.destinationPluginId                  = kPluginIdDummy;
    construction.destinationPluginShortId             = L"dummy";
    construction.destinationInstanceContext           = L"captured-context";
    construction.options.verifyAfterCopy              = true;
    construction.options.executionMode                = FileOperations::ExecutionMode::Queue;
    construction.options.bandwidthLimitBytesPerSecond = std::nullopt; // Captured unlimited, never reread defaults.
    construction.moveClipboardSequence                = 731u;         // Provenance only: constructing an envelope cannot consume it.
    std::vector<State::Task*> beforeTasks;
    privateOperations->CollectTasks(beforeTasks);
    State::OperationAdmission admission{};
    const HRESULT buildHr = State::BuildOperationAdmission(construction, admission);
    const auto* transfer  = admission.plans.size() == 1u ? std::get_if<FileOperations::TransferPlan>(&admission.plans.front()) : nullptr;
    if (! state.Require(buildHr == S_OK && transfer && transfer->strategy == FileOperations::OperationStrategy::Native &&
                            transfer->sourceEndpoint.instanceId == L"host/context/captured-context" &&
                            transfer->destinationEndpoint.instanceId == L"host/context/captured-context" &&
                            transfer->destination.providerFolderPath == destinationFolder && transfer->selectedItems.size() == 1u &&
                            transfer->selectedItems.front().providerPath == sourcePath && transfer->options.verifyAfterCopy &&
                            transfer->options.executionMode == FileOperations::ExecutionMode::Queue &&
                            ! transfer->options.bandwidthLimitBytesPerSecond.has_value() && transfer->moveClipboardSequence.has_value() &&
                            transfer->moveClipboardSequence->windowsSequenceNumber == 731u && ! admission.preWorkerReleaseBarrier &&
                            admission.sourcePane == FolderWindow::Pane::Right && admission.destinationPane == FolderWindow::Pane::Left,
                        L"Private admission must preserve captured provider, options and clipboard provenance without publication."))
        return false;
    State::OperationAdmission noBarrier = admission;
    uint64_t refusedTaskId              = 999u;
    state.Require(privateOperations->StartOperation(std::move(noBarrier), &refusedTaskId) == E_INVALIDARG && refusedTaskId == 0u,
                  L"A constructed clipboard-Move envelope cannot publish without its initial worker-release barrier.");
    State::TransferDestinationNameFailure nameFailure;
    std::atomic<bool> nameCancelled{false};
    state.Require(State::ValidateTransferDestinationNames(admission.plans, fixture.fsDummy.get(), FILESYSTEM_MOVE, {}, nameCancelled, nameFailure) == S_OK &&
                      nameFailure.rejectedLeaf.empty(),
                  L"Private destination-name validation must accept the provider's valid name without task diagnostics.");
    auto badNamePlans                                                             = admission.plans;
    std::get<FileOperations::TransferPlan>(badNamePlans.front()).explicitMappings = {{0u, L"/queued-admission-contracts-destination/bad:name.txt"}};
    state.Require(State::ValidateTransferDestinationNames(badNamePlans, fixture.fsDummy.get(), FILESYSTEM_MOVE, {}, nameCancelled, nameFailure) ==
                          HRESULT_FROM_WIN32(ERROR_INVALID_NAME) &&
                      nameFailure.rejectedLeaf == L"bad:name.txt" && nameFailure.sourcePath == sourcePath && nameFailure.destinationFolder == destinationFolder,
                  L"A rejected candidate must return its exact destination leaf, source and parent without publishing diagnostics.");
    state.Require(State::ValidateTransferDestinationNames(admission.plans, fixture.fsDummy.get(), FILESYSTEM_MOVE, {}, nameCancelled, nameFailure) == S_OK &&
                      nameFailure.rejectedLeaf.empty() && nameFailure.sourcePath.empty() && nameFailure.destinationFolder.empty(),
                  L"A subsequent valid candidate must clear all previous name-refusal facts.");
    nameCancelled.store(true, std::memory_order_release);
    nameFailure.rejectedLeaf = L"stale-refusal";
    state.Require(State::ValidateTransferDestinationNames(admission.plans, fixture.fsDummy.get(), FILESYSTEM_MOVE, {}, nameCancelled, nameFailure) ==
                          HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
                      nameFailure.rejectedLeaf.empty(),
                  L"Cancellation must not return an earlier candidate's destination-name refusal.");
    nameCancelled.store(false, std::memory_order_release);
    std::stop_source nameStop;
    nameStop.request_stop();
    state.Require(
        State::ValidateTransferDestinationNames(admission.plans, fixture.fsDummy.get(), FILESYSTEM_MOVE, nameStop.get_token(), nameCancelled, nameFailure) ==
            HRESULT_FROM_WIN32(ERROR_CANCELLED),
        L"Worker stop must prevent private destination-name validation from continuing.");
    std::vector<FileOperations::MutationInterlockScope> scopes;
    if (! state.Require(State::BuildTransferInterlockScopes(admission.plans, fixture.fsDummy.get(), fixture.fsDummy.get(), FILESYSTEM_MOVE, scopes) == S_OK &&
                            scopes.size() == 2u &&
                            std::ranges::count_if(
                                scopes, [](const auto& scope) noexcept { return scope.access == FileOperations::MutationInterlockAccess::WriteSource; }) == 1 &&
                            std::ranges::count_if(scopes,
                                                  [](const auto& scope) noexcept
    { return scope.access == FileOperations::MutationInterlockAccess::PublishDestination; }) == 1,
                        L"Private preparation must bind both the Native Move source and selected destination scopes."))
        return false;
    // Exercise the actual scheduler and live-output consent predicates with immutable private
    // registrations. Public task identity stays fixed while the proposed admission changes.
    auto candidate             = std::make_unique<State::Task>(*privateOperations);
    auto peer                  = std::make_unique<State::Task>(*privateOperations);
    const auto artifactReceipt = [](std::wstring_view name)
    {
        FileOperationArtifacts::TouchGuardReceipt receipt;
        FileOperationArtifacts::TouchGuardItem item;
        item.path = std::filesystem::path(name);
        receipt.items.push_back(std::move(item));
        return receipt;
    };
    const auto receiptNames = [](const auto& receipt, std::wstring_view name) noexcept
    { return static_cast<bool>(receipt) && receipt->items.size() == 1u && receipt->items.front().path.native() == name; };
    candidate->StoreInitialArtifactTouchReceipt(artifactReceipt(L"live-original"));
    State::Task viewFixture(*privateOperations);
    auto viewConstruction = construction;
    viewConstruction.destinationFileSystem.reset();
    viewConstruction.resolvedItems = {FolderWindow::ResolvedFileOperationItem{
        .sourcePath      = sourcePath,
        .destinationPath = std::filesystem::path(destinationFolder) / L"single.txt",
    }};
    State::OperationAdmission mappedViewAdmission;
    if (! state.Require(State::BuildOperationAdmission(viewConstruction, mappedViewAdmission) == S_OK,
                        L"The coherent-view fixture requires a canonical explicit-mapping plan."))
        return false;
    auto retainedPlans = std::make_shared<const FileOperations::FileOperationPlanGroup>(mappedViewAdmission.plans);
    State::AdmissionDestinationInputs originalDestination;
    originalDestination.folder          = destinationFolder;
    originalDestination.pluginId        = kPluginIdDummy;
    originalDestination.pluginShortId   = L"dummy";
    originalDestination.instanceContext = construction.destinationInstanceContext;
    originalDestination.resolvedItems   = mappedViewAdmission.resolvedItems;
    viewFixture.InitializeAdmissionInputs(retainedPlans, std::move(originalDestination), artifactReceipt(L"view-original"));
    const auto originalView                = viewFixture.LoadSelectedAdmission();
    viewConstruction.destinationFileSystem = fixture.fsDummy;
    viewConstruction.resolvedItems.clear();
    State::OperationAdmission plainViewAdmission;
    if (! state.Require(State::BuildOperationAdmission(viewConstruction, plainViewAdmission) == S_OK,
                        L"The coherent-view fixture requires a canonical plain-destination plan."))
        return false;
    const auto plainPlans = std::make_shared<const FileOperations::FileOperationPlanGroup>(plainViewAdmission.plans);
    State::AdmissionDestinationInputs explicitSameProvider;
    explicitSameProvider.fileSystem      = fixture.fsDummy;
    explicitSameProvider.folder          = destinationFolder;
    explicitSameProvider.pluginId        = kPluginIdDummy;
    explicitSameProvider.pluginShortId   = L"dummy";
    explicitSameProvider.instanceContext = construction.destinationInstanceContext;
    viewFixture.InitializeAdmissionInputs(plainPlans, std::move(explicitSameProvider), artifactReceipt(L"view-second"));
    const auto secondView = viewFixture.LoadSelectedAdmission();
    state.Require(
        originalView->plans == retainedPlans && secondView->plans == plainPlans && ! originalView->destination->fileSystem &&
            secondView->destination->fileSystem.get() == fixture.fsDummy.get() && originalView->destination->folder == destinationFolder &&
            secondView->destination->folder == destinationFolder && originalView->destination->resolvedItems.size() == 1u &&
            secondView->destination->resolvedItems.empty() &&
            std::get<FileOperations::TransferPlan>(originalView->plans->front()).explicitMappings.size() == 1u &&
            std::get<FileOperations::TransferPlan>(secondView->plans->front()).explicitMappings.empty() &&
            receiptNames(originalView->artifactTouchReceipt, L"view-original") && receiptNames(secondView->artifactTouchReceipt, L"view-second") &&
            ! originalView->interlock && ! secondView->interlock && ! originalView->preparation && ! secondView->preparation,
        L"Private initial views retain coherent typed mappings and receipt while preserving null versus explicit same-provider semantics without authority.");
    auto resealedPlans = std::make_shared<const FileOperations::FileOperationPlanGroup>(plainViewAdmission.plans);
    viewFixture.StorePlans(resealedPlans);
    const auto resealedView = viewFixture.LoadSelectedAdmission();
    state.Require(resealedView->plans == resealedPlans && secondView->plans == plainPlans && resealedView->destination == secondView->destination &&
                      resealedView->artifactTouchReceipt == secondView->artifactTouchReceipt,
                  L"Initial plan resealing replaces only the small selected view and preserves destination ownership without copying mappings.");
    candidate->_batchRenameArtifactPromptNextRequestId    = 2u;
    candidate->_batchRenameArtifactPromptPendingRequestId = 2u;
    candidate->CompleteBatchRenameArtifactPrompt(1u, S_OK, artifactReceipt(L"stale"));
    candidate->CompleteBatchRenameArtifactPrompt(0u, S_OK, artifactReceipt(L"uncorrelated"));
    state.Require(candidate->IsBatchRenameArtifactPromptPending(2u) && ! candidate->_batchRenameArtifactPromptResultReceipt.has_value() &&
                      candidate->_batchRenameArtifactPromptStatus == E_PENDING &&
                      receiptNames(candidate->LoadSelectedAdmission()->artifactTouchReceipt, L"live-original"),
                  L"Stale or uncorrelated artifact replies cannot complete a newer prompt or install a live receipt.");
    candidate->CompleteBatchRenameArtifactPrompt(2u, E_ACCESSDENIED, artifactReceipt(L"refused"));
    candidate->CompleteBatchRenameArtifactPrompt(2u, S_OK, artifactReceipt(L"duplicate"));
    state.Require(candidate->_batchRenameArtifactPromptCompleted && candidate->_batchRenameArtifactPromptStatus == E_ACCESSDENIED &&
                      ! candidate->_batchRenameArtifactPromptResultReceipt.has_value() &&
                      receiptNames(candidate->LoadSelectedAdmission()->artifactTouchReceipt, L"live-original"),
                  L"Failed artifact preparation retains the installed receipt and rejects a duplicate successful reply.");
    candidate->_batchRenameArtifactPromptNextRequestId    = 3u;
    candidate->_batchRenameArtifactPromptPendingRequestId = 3u;
    candidate->_batchRenameArtifactPromptCompleted        = false;
    candidate->_batchRenameArtifactPromptStatus           = E_PENDING;
    candidate->CompleteBatchRenameArtifactPrompt(2u, S_OK, artifactReceipt(L"late-previous-request"));
    state.Require(candidate->_batchRenameArtifactPromptStatus == E_PENDING && ! candidate->_batchRenameArtifactPromptCompleted &&
                      ! candidate->_batchRenameArtifactPromptResultReceipt.has_value(),
                  L"A late reply for a completed request must leave the next pending artifact request untouched.");
    candidate->CompleteBatchRenameArtifactPrompt(3u, S_OK, artifactReceipt(L"candidate"));
    state.Require(candidate->_batchRenameArtifactPromptCompleted && candidate->_batchRenameArtifactPromptStatus == S_OK &&
                      receiptNames(candidate->_batchRenameArtifactPromptResultReceipt, L"candidate") &&
                      receiptNames(candidate->LoadSelectedAdmission()->artifactTouchReceipt, L"live-original"),
                  L"A correlated success stays private until the admission owner commits it.");
    std::optional<FileOperationArtifacts::TouchGuardReceipt> preparedArtifact = artifactReceipt(L"previous-output");
    state.Require(candidate->PrepareArtifactTouchGuard(admission.plans, fixture.fsDummy.get(), fixture.fsDummy.get(), preparedArtifact) == S_OK &&
                      ! preparedArtifact.has_value() && receiptNames(candidate->LoadSelectedAdmission()->artifactTouchReceipt, L"live-original"),
                  L"Ordinary paths need no prompt and clear only the private output receipt.");
    auto malformedArtifactPlans = admission.plans;
    std::get<FileOperations::TransferPlan>(malformedArtifactPlans.front()).sourceEndpoint.pathIdentity.reset();
    preparedArtifact = artifactReceipt(L"previous-output");
    state.Require(
        candidate->PrepareArtifactTouchGuard(malformedArtifactPlans, fixture.fsDummy.get(), fixture.fsDummy.get(), preparedArtifact) == E_INVALIDARG &&
            receiptNames(preparedArtifact, L"previous-output") && receiptNames(candidate->LoadSelectedAdmission()->artifactTouchReceipt, L"live-original"),
        L"Failed artifact capture preserves both the caller's output and the installed live receipt.");
    candidate->_cancelled.store(true, std::memory_order_release);
    state.Require(candidate->PrepareArtifactTouchGuard(admission.plans, fixture.fsDummy.get(), fixture.fsDummy.get(), preparedArtifact) ==
                          HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
                      receiptNames(preparedArtifact, L"previous-output"),
                  L"Cancellation leaves the private artifact output unchanged.");
    candidate->_cancelled.store(false, std::memory_order_release);
    candidate->_batchRenameArtifactPromptCompleted = false;
    candidate->FailBatchRenameArtifactPromptWithoutPayload(S_OK);
    state.Require(candidate->_batchRenameArtifactPromptStatus == E_UNEXPECTED && ! candidate->_batchRenameArtifactPromptResultReceipt.has_value() &&
                      receiptNames(candidate->LoadSelectedAdmission()->artifactTouchReceipt, L"live-original"),
                  L"The payload-less fallback is failure-only and cannot grant artifact consent.");
    candidate->_taskId = 42u;
    candidate->_admissionKey.store(42u, std::memory_order_release);
    // Dummy supplies conservative scope identity. Real local objects prove that
    // coexisting admissions independently retain their bound destination parents.
    const auto localSourceRoot  = root / L"owned-source";
    const auto localDestination = root / L"owned-destination";
    const auto localReplacement = root / L"owned-replacement";
    const auto localSource      = localSourceRoot / L"single.txt";
    if (! state.Require(fixture.fsLocal && RecreateEmptyDirectory(localSourceRoot) && RecreateEmptyDirectory(localDestination) &&
                            RecreateEmptyDirectory(localReplacement) && SelfTest::WriteTextFile(localSource, "unchanged-local-source"),
                        L"Owned-admission fixtures require two real local destinations and an unchanged selected source."))
        return false;
    auto ownedConstruction                     = construction;
    ownedConstruction.fileSystem               = fixture.fsLocal;
    ownedConstruction.destinationFileSystem    = fixture.fsLocal;
    ownedConstruction.sourcePaths              = {localSource};
    ownedConstruction.destinationFolder        = localDestination;
    ownedConstruction.sourcePluginId           = kPluginIdLocal;
    ownedConstruction.destinationPluginId      = kPluginIdLocal;
    ownedConstruction.sourcePluginShortId      = L"file";
    ownedConstruction.destinationPluginShortId = L"file";
    ownedConstruction.sourceInstanceContext.clear();
    ownedConstruction.destinationInstanceContext.clear();
    State::OperationAdmission ownedAdmission;
    State::MutationInterlockScopes ownedScopes;
    if (! state.Require(
            State::BuildOperationAdmission(ownedConstruction, ownedAdmission) == S_OK &&
                State::BuildTransferInterlockScopes(ownedAdmission.plans, fixture.fsLocal.get(), fixture.fsLocal.get(), FILESYSTEM_MOVE, ownedScopes) == S_OK,
            L"Local ownership fixture must use canonical admission construction and bound scope preparation."))
        return false;
    auto movableScopes       = ownedScopes;
    const auto* movedStorage = movableScopes.data();
    candidate->InstallPreparedMutationInterlockScopes(std::move(movableScopes), false);
    const auto originalCandidateOwner = candidate->LoadMutationInterlockAdmission();
    state.Require(originalCandidateOwner && originalCandidateOwner->scopes->data() == movedStorage,
                  L"Freezing preparation must move the authority vector without copying its storage.");
    State::MutationInterlockAdmission candidateAdmission = *originalCandidateOwner;
    peer->_taskId                                        = 41u;
    peer->_admissionKey.store(41u, std::memory_order_release);
    peer->InstallPreparedMutationInterlockScopes(ownedScopes, false);
    const auto originalPeerOwner = peer->LoadMutationInterlockAdmission();
    std::array<State::ActiveMutationInterlock, 1> originalPeer{{{.taskId = 41u, .task = peer.get(), .admission = originalPeerOwner}}};
    auto replacementInput                     = *originalPeerOwner;
    replacementInput.key                      = 43u;
    auto replacementConstruction              = ownedConstruction;
    replacementConstruction.destinationFolder = localReplacement;
    State::OperationAdmission replacementAdmission;
    State::MutationInterlockScopes changedScopes;
    if (! state.Require(State::BuildOperationAdmission(std::move(replacementConstruction), replacementAdmission) == S_OK &&
                            State::BuildTransferInterlockScopes(
                                replacementAdmission.plans, fixture.fsLocal.get(), fixture.fsLocal.get(), FILESYSTEM_MOVE, changedScopes) == S_OK,
                        L"Replacement scopes must bind the canonical alternate destination, not edit a retained scope pathname."))
        return false;
    replacementInput.scopes  = std::make_shared<const State::MutationInterlockScopes>(std::move(changedScopes));
    const auto isPublication = [](const FileOperations::MutationInterlockScope& scope) noexcept
    { return scope.access == FileOperations::MutationInterlockAccess::PublishDestination; };
    const auto originalPublication    = std::ranges::find_if(*originalPeerOwner->scopes, isPublication);
    const auto replacementPublication = std::ranges::find_if(*replacementInput.scopes, isPublication);
    if (! state.Require(
            originalPublication != originalPeerOwner->scopes->end() && replacementPublication != replacementInput.scopes->end() && originalPublication->root &&
                replacementPublication->root && originalPublication->root->retained.authority.boundObject &&
                replacementPublication->root->retained.authority.boundObject && originalPublication->root->retained.providerPath == localDestination.native() &&
                replacementPublication->root->retained.providerPath == localReplacement.native() && originalPublication->root != replacementPublication->root,
            L"Coexisting registrations must retain separately bound original and replacement destination parents."))
        return false;
    std::array<State::ActiveMutationInterlock, 1> replacementPeer{
        {{.taskId = 41u, .task = peer.get(), .admission = std::make_shared<const State::MutationInterlockAdmission>(std::move(replacementInput))}}};
    uint64_t comparisons = 0u;
    state.Require(! State::MutationInterlocksPermit(42u, candidateAdmission, originalPeer, comparisons) && comparisons > 0u,
                  L"Overlapping immutable registrations must block without explicit consent.");
    candidateAdmission.relations.concurrentKeys[0] = 41u;
    candidateAdmission.relations.concurrentCount   = 1u;
    state.Require(State::MutationInterlocksPermit(42u, candidateAdmission, originalPeer, comparisons) &&
                      ! State::MutationInterlocksPermit(42u, candidateAdmission, replacementPeer, comparisons) &&
                      State::HasConcurrentPublisherConsent(candidateAdmission, originalPeer.front()) &&
                      ! State::HasConcurrentPublisherConsent(candidateAdmission, replacementPeer.front()),
                  L"An old Run grant cannot authorize a replacement registration with the same public task identity.");
    candidateAdmission.relations.concurrentKeys[0] = 43u;
    state.Require(State::MutationInterlocksPermit(42u, candidateAdmission, replacementPeer, comparisons) &&
                      State::HasConcurrentPublisherConsent(candidateAdmission, replacementPeer.front()),
                  L"Fresh consent names the replacement admission explicitly.");
    candidateAdmission.relations.concurrentCount = 0u;
    auto reciprocalInput                         = *originalPeerOwner;
    reciprocalInput.relations.concurrentKeys[0]  = 42u;
    reciprocalInput.relations.concurrentCount    = 1u;
    auto reciprocalPeer                          = originalPeer;
    reciprocalPeer.front().admission             = std::make_shared<const State::MutationInterlockAdmission>(reciprocalInput);
    state.Require(State::HasConcurrentPublisherConsent(candidateAdmission, reciprocalPeer.front()) &&
                      ! State::HasConcurrentPublisherConsent(candidateAdmission, originalPeer.front()) &&
                      reciprocalPeer.front().admission->scopes == originalPeerOwner->scopes,
                  L"A decided snapshot retains scopes while an older captured snapshot keeps its original decision.");
    candidateAdmission.key = 44u;
    state.Require(! State::HasConcurrentPublisherConsent(candidateAdmission, reciprocalPeer.front()),
                  L"A publisher's reciprocal grant names the exact recipient admission.");
    reciprocalInput.relations.concurrentKeys[0] = 44u;
    reciprocalPeer.front().admission            = std::make_shared<const State::MutationInterlockAdmission>(reciprocalInput);
    state.Require(State::HasConcurrentPublisherConsent(candidateAdmission, reciprocalPeer.front()) &&
                      ! State::HasConcurrentPublisherConsent(candidateAdmission, replacementPeer.front()),
                  L"Reciprocal consent is owned by the captured registration, never borrowed from its Task.");
    candidateAdmission.relations.advisoryDecided    = true;
    candidateAdmission.relations.predecessorKeys[0] = 41u;
    candidateAdmission.relations.predecessorCount   = 1u;
    std::vector<State::ActiveMutationInterlock> published{originalPeer.front(), replacementPeer.front()};
    state.Require(State::HasPublishedOverlapPredecessor(candidateAdmission, published, {}) &&
                      State::HasPublishedOverlapPredecessor(candidateAdmission, {}, published),
                  L"Queue-after recognizes the exact original admission in either registry.");
    State::EraseMutationInterlockAdmission(published, 41u);
    state.Require(published.size() == 1u && published.front().admission->key == 43u &&
                      ! State::HasPublishedOverlapPredecessor(candidateAdmission, published, {}) &&
                      originalPeerOwner->scopes->data() != replacementPeer.front().admission->scopes->data() &&
                      originalPeerOwner->scopes->back().providerPath == ownedScopes.back().providerPath,
                  L"Exact-key withdrawal releases the original predecessor while preserving replacement and retained original authority storage.");
    candidateAdmission.relations.predecessorKeys[0] = 43u;
    state.Require(State::HasPublishedOverlapPredecessor(candidateAdmission, published, {}),
                  L"A renewed Queue-after choice names the replacement registration.");
    published = {originalPeer.front(), replacementPeer.front()};
    State::EraseMutationInterlockAdmission(published, 43u);
    state.Require(published.size() == 1u && published.front().admission == originalPeerOwner,
                  L"Refusing a private replacement must withdraw only its key and retain the original snapshot exactly.");
    std::shared_ptr<const FileOperations::PreparationSnapshot> facts;
    if (! state.Require(State::BuildPreparationFacts(501u, FILESYSTEM_MOVE, admission.plans, scopes, 42u, facts) == S_OK && facts && facts->taskId == 501u &&
                            facts->operation == FILESYSTEM_MOVE && facts->selectedRootCount == 1u && facts->scopes.size() == 2u &&
                            facts->strategies.size() == 1u && facts->strategies.front().strategy == FileOperations::OperationStrategy::Native &&
                            facts->selectedRootReadinessUs >= 42u,
                        L"Private preparation facts must describe the supplied admission and scopes without a Task observer."))
        return false;
    auto invalidPlans = admission.plans;
    invalidPlans.emplace_back(FileOperations::DeletePlan{});
    const auto* originalScopeStorage = scopes.data();
    const auto originalScopeRoot     = scopes.front().root;
    const auto originalFacts         = facts;
    state.Require(State::BuildTransferInterlockScopes(invalidPlans, fixture.fsDummy.get(), fixture.fsDummy.get(), FILESYSTEM_MOVE, scopes) == E_INVALIDARG &&
                      scopes.data() == originalScopeStorage && scopes.size() == 2u && scopes.front().root == originalScopeRoot &&
                      State::BuildPreparationFacts(501u, FILESYSTEM_MOVE, {}, scopes, 42u, facts) == E_UNEXPECTED && facts == originalFacts,
                  L"Failed candidate preparation must leave previously accepted scope storage and facts intact.");
    {
        // runtime-queue#6: a peer whose destination change is being prepared holds two prepared
        // registrations (its admitted key and its candidate key). The advice still offers Run at
        // the same time, names both keys, and counts and lists the peer once.
        auto twinAdmitted         = std::make_unique<State::Task>(*privateOperations);
        auto twinCandidate        = std::make_unique<State::Task>(*privateOperations);
        const uint64_t twinTaskId = privateOperations->AllocateLateQueueOrderKey();
        twinAdmitted->_taskId     = twinTaskId;
        twinCandidate->_taskId    = twinTaskId;
        twinAdmitted->_admissionKey.store(twinTaskId, std::memory_order_release);
        twinCandidate->_admissionKey.store(privateOperations->AllocateLateQueueOrderKey(), std::memory_order_release);
        twinAdmitted->InstallPreparedMutationInterlockScopes(ownedScopes, false);
        twinCandidate->InstallPreparedMutationInterlockScopes(ownedScopes, false);
        privateOperations->PublishPreparedMutationInterlock(*twinAdmitted);
        privateOperations->PublishPreparedMutationInterlock(*twinCandidate);
        const uint64_t twinAdmittedKey           = twinAdmitted->_admissionKey.load(std::memory_order_acquire);
        const uint64_t twinCandidateKey          = twinCandidate->_admissionKey.load(std::memory_order_acquire);
        const auto withdrawTwins                 = wil::scope_exit([&]() noexcept
        {
            privateOperations->WithdrawPreparedMutationInterlock(twinAdmittedKey);
            privateOperations->WithdrawPreparedMutationInterlock(twinCandidateKey);
        });
        State::MutationInterlockAdmission asking = *originalCandidateOwner;
        asking.key                               = privateOperations->AllocateLateQueueOrderKey();
        const auto twinAdvice                    = privateOperations->FindSameHostOverlap(42u, asking);
        const auto twinKeysEnd                   = twinAdvice.admissionKeys.begin() + static_cast<std::ptrdiff_t>(twinAdvice.relationCount);
        state.Require(twinAdvice.problem != FileOperations::SameHostOverlapProblem::None && twinAdvice.relationCount == 2u &&
                          std::ranges::find(twinAdvice.admissionKeys.begin(), twinKeysEnd, twinAdmittedKey) != twinKeysEnd &&
                          std::ranges::find(twinAdvice.admissionKeys.begin(), twinKeysEnd, twinCandidateKey) != twinKeysEnd &&
                          twinAdvice.CanRunConcurrently() && twinAdvice.taskCount == 1u && twinAdvice.taskIdCount == 1u &&
                          twinAdvice.taskIds[0] == twinTaskId && twinAdvice.taskIds[1] == 0u,
                      std::format(L"A peer with two registrations must keep Run at the same time and be listed once (relations={}, tasks={}, listed={}, "
                                  L"run={}, second={}).",
                                  twinAdvice.relationCount,
                                  twinAdvice.taskCount,
                                  twinAdvice.taskIdCount,
                                  twinAdvice.CanRunConcurrently(),
                                  twinAdvice.taskIds[1]));
    }
    {
        // I35 S4 round 3: no three-task queue cycle. D (the oldest) prepares slowly; B and C publish first and
        // C queues after B; once D published, B queues after D. D must not name C (C waits on D through B), and a
        // decision naming C anyway is published without that edge, so the three never wait on each other.
        const auto makePeer = [&]()
        {
            auto peer     = std::make_unique<State::Task>(*privateOperations);
            peer->_taskId = privateOperations->AllocateLateQueueOrderKey();
            peer->_admissionKey.store(peer->_taskId, std::memory_order_release);
            peer->InstallPreparedMutationInterlockScopes(ownedScopes, false);
            return peer;
        };
        auto taskD               = makePeer();
        auto taskB               = makePeer();
        auto taskC               = makePeer();
        const uint64_t keyD      = taskD->_admissionKey.load(std::memory_order_acquire);
        const uint64_t keyB      = taskB->_admissionKey.load(std::memory_order_acquire);
        const uint64_t keyC      = taskC->_admissionKey.load(std::memory_order_acquire);
        const auto withdrawPeers = wil::scope_exit([&]() noexcept
        {
            privateOperations->WithdrawPreparedMutationInterlock(keyD);
            privateOperations->WithdrawPreparedMutationInterlock(keyB);
            privateOperations->WithdrawPreparedMutationInterlock(keyC);
        });
        const auto decide        = [&](State::Task& peer, std::initializer_list<uint64_t> predecessors) noexcept
        {
            State::MutationInterlockRelations relations{};
            relations.advisoryDecided = true;
            for (const uint64_t key : predecessors)
                relations.predecessorKeys[relations.predecessorCount++] = key;
            privateOperations->PublishMutationInterlockDecision(peer, relations);
        };
        privateOperations->PublishPreparedMutationInterlock(*taskB);
        privateOperations->PublishPreparedMutationInterlock(*taskC);
        decide(*taskC, {keyB});
        privateOperations->PublishPreparedMutationInterlock(*taskD);
        decide(*taskB, {keyD});
        const auto adviceD    = privateOperations->FindSameHostOverlap(*taskD);
        const auto adviceDEnd = adviceD.admissionKeys.begin() + static_cast<std::ptrdiff_t>(adviceD.relationCount);
        const bool namesC     = std::ranges::find(adviceD.admissionKeys.begin(), adviceDEnd, keyC) != adviceDEnd;
        state.Require(! namesC,
                      std::format(L"A task must not queue after a newer peer that waits on it through another (relations={}, names C={}).",
                                  adviceD.relationCount,
                                  namesC));
        decide(*taskD, {keyC});
        const auto admissionD = taskD->LoadMutationInterlockAdmission();
        bool keptCycle        = false;
        if (admissionD)
        {
            const size_t count = (std::min)(admissionD->relations.predecessorCount, admissionD->relations.predecessorKeys.size());
            for (size_t index = 0u; index < count; ++index)
                keptCycle = keptCycle || admissionD->relations.predecessorKeys[index] == keyC;
        }
        state.Require(admissionD && admissionD->relations.advisoryDecided && ! keptCycle,
                      L"A decision naming a peer that waits on this task through another is published without that edge.");
    }
    if (! ValidateClipboardReadyStaleToken(fixture, state, sourcePath, destinationFolder))
        return false;

    return state.failure.empty();
}
