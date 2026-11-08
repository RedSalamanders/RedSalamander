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

    const ULONGLONG readyDeadline = GetTickCount64() + static_cast<ULONGLONG>(SelfTest::Scale(std::chrono::milliseconds{5'000}).count());
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

    const ULONGLONG completionDeadline = GetTickCount64() + static_cast<ULONGLONG>(SelfTest::Scale(std::chrono::milliseconds{5'000}).count());
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

// This raw operation-control gate is specific to holding a provider metadata borrow;
// the shared test helpers do not expose a synchronous IFileSystemOperationControl gate.
struct DirectoryWitnessBorrowControl final : IFileSystemOperationControl
{
    wil::unique_event_nothrow entered;
    wil::unique_event_nothrow release;

    HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void*) noexcept override
    {
        if (! abort)
            return E_POINTER;
        *abort = FALSE;
        entered.SetEvent();
        if (WaitForSingleObject(release.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(10000u))) != WAIT_OBJECT_0)
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        // The held read exists only to exercise pin ownership. Cancel after release
        // so this witness does not depend on backup-stream inspection permissions.
        *abort = TRUE;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void*) noexcept override
    {
        if (! mode)
            return E_POINTER;
        *mode = FILESYSTEM_DISCOVERY_AHEAD;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress*, void*) noexcept override
    {
        return S_OK;
    }
};

[[nodiscard]] bool ValidateLocalDirectoryPinHandoff(const wil::com_ptr<IFileSystem>& fileSystem,
                                                    const std::filesystem::path& root,
                                                    SelfTest::CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr auto readFlags   = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    constexpr auto strictFlags = static_cast<FileSystemBindFlags>(readFlags | FILESYSTEM_BIND_DELETE);
    constexpr auto mixedFlags  = static_cast<FileSystemBindFlags>(strictFlags | FILESYSTEM_BIND_RENAME);
    const auto directory       = root / L"handoff-directory";
    const auto otherDirectory  = root / L"handoff-other-directory";
    const auto regular         = root / L"handoff-file.bin";
    if (! state.Require(fileSystem && RecreateEmptyDirectory(directory) && RecreateEmptyDirectory(otherDirectory) &&
                            SelfTest::WriteTextFile(regular, "handoff"),
                        L"Directory pin handoff requires its own synthetic files."))
        return false;
    FileOperations::QualifiedEndpoint endpoint;
    if (! state.Require(
            FileOperations::TryQualifyEndpointForSelfTest(fileSystem, directory.native(), FILESYSTEM_DELETE, kPluginIdLocal, L"host/default", endpoint),
            L"Directory handoff must use the acquired sandbox's qualified Local path profile."))
        return false;
    const auto bind = [&](const std::filesystem::path& path, FileSystemBindFlags flags) noexcept
    { return FileOperations::BindObjectAuthority(fileSystem.get(), path.native(), endpoint.profileId, flags); };
    auto witness = bind(directory, readFlags);
    if (! state.Require(witness.state == FileOperations::ObjectBindingState::Bound && witness.authority.boundObject,
                        L"Directory handoff requires a metadata-only witness."))
        return false;
    wil::com_ptr<IFileSystemBoundMetadata> metadata;
    if (! state.Require(SUCCEEDED(witness.authority.boundObject->QueryInterface(IID_PPV_ARGS(metadata.addressof()))) && metadata,
                        L"The metadata witness must expose bound metadata."))
        return false;
    const auto stillReadable = [&]() noexcept
    {
        FileSystemBasicInformation basic{};
        basic.sizeBytes = sizeof(basic);
        return witness.authority.boundObject->GetBasicInformation(&basic) == S_OK;
    };
    const auto rejectReceiver = [&](const std::filesystem::path& path, FileSystemBindFlags flags, HRESULT expected) noexcept
    {
        auto receiver = bind(path, flags);
        wil::com_ptr<IFileSystemBoundObjectPinHandoff> handoff;
        return receiver.state == FileOperations::ObjectBindingState::Bound && receiver.authority.boundObject &&
               SUCCEEDED(receiver.authority.boundObject->QueryInterface(IID_PPV_ARGS(handoff.addressof()))) && handoff &&
               handoff->TakeDirectoryReadPinFrom(witness.authority.boundObject.get()) == expected && stillReadable();
    };
    if (! state.Require(rejectReceiver(otherDirectory, strictFlags, HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH)) &&
                            rejectReceiver(regular, strictFlags, E_ACCESSDENIED) && rejectReceiver(directory, mixedFlags, E_ACCESSDENIED),
                        L"Mismatched, regular-file, and mixed-grant receivers must reject without retiring the witness."))
        return false;

    auto strict = bind(directory, strictFlags);
    wil::com_ptr<IFileSystemBoundObjectPinHandoff> handoff;
    if (! state.Require(strict.state == FileOperations::ObjectBindingState::Bound && strict.authority.boundObject &&
                            SUCCEEDED(strict.authority.boundObject->QueryInterface(IID_PPV_ARGS(handoff.addressof()))) && handoff,
                        L"The exact strict directory receiver must expose the handoff capability."))
        return false;
    FileSystemBoundObjectSnapshot before{};
    before.sizeBytes = sizeof(before);
    BOOL same        = FALSE;
    if (! state.Require(witness.authority.boundObject->GetSnapshot(&before) == S_OK && before.objectId && before.objectIdBytes > 0u &&
                            strict.authority.boundObject->IsSameObject(witness.authority.boundObject.get(), &same) == S_OK && same,
                        L"The witness and strict receiver must initially name the same exact directory."))
        return false;
    const std::vector<std::byte> identity(static_cast<const std::byte*>(before.objectId),
                                          static_cast<const std::byte*>(before.objectId) + before.objectIdBytes);

    DirectoryWitnessBorrowControl control;
    if (! state.Require(SUCCEEDED(control.entered.create(wil::EventOptions::ManualReset)) && SUCCEEDED(control.release.create(wil::EventOptions::ManualReset)),
                        L"Directory witness read gate events could not be created."))
        return false;
    FileSystemOptions options{};
    options.sizeBytes        = sizeof(options);
    options.operationControl = &control;
    std::atomic<HRESULT> readHr{E_PENDING};
    std::jthread reader;
    try
    {
        reader = std::jthread([&]() noexcept
        {
            FileSystemMetadataSnapshot snapshot{};
            snapshot.sizeBytes = sizeof(snapshot);
            readHr.store(metadata->GetMetadataSnapshot(&options, &snapshot), std::memory_order_release);
        });
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::system_error& error)
    {
        // Thread creation is the throwing boundary; no borrow exists on failure.
        Debug::Error(L"Directory witness reader could not start: {}", error.code().value());
        return state.Require(false, L"The directory witness reader could not start.");
    }
    auto joinReader = wil::scope_exit([&]() noexcept
    {
        control.release.SetEvent();
        if (reader.joinable())
            reader.join();
    });
    if (! state.Require(WaitForSingleObject(control.entered.get(), static_cast<DWORD>(SelfTest::Scale(5s).count())) == WAIT_OBJECT_0 &&
                            handoff->TakeDirectoryReadPinFrom(witness.authority.boundObject.get()) == HRESULT_FROM_WIN32(ERROR_BUSY) && stillReadable(),
                        L"An active metadata borrow must refuse handoff without retiring the witness."))
        return false;
    control.release.SetEvent();
    reader.join();
    joinReader.release();
    if (! state.Require(readHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
                            handoff->TakeDirectoryReadPinFrom(witness.authority.boundObject.get()) == S_OK,
                        L"Handoff must succeed after the metadata borrow quiesces."))
        return false;
    FileSystemBoundObjectSnapshot after{};
    after.sizeBytes = sizeof(after);
    FileSystemBasicInformation retiredBasic{};
    retiredBasic.sizeBytes = sizeof(retiredBasic);
    FileSystemMetadataSnapshot retiredMetadata{};
    retiredMetadata.sizeBytes = sizeof(retiredMetadata);
    same                      = FALSE;
    if (! state.Require(witness.authority.boundObject->GetSnapshot(&after) == S_OK && after.objectId && after.objectIdBytes == identity.size() &&
                            std::memcmp(after.objectId, identity.data(), identity.size()) == 0 &&
                            strict.authority.boundObject->IsSameObject(witness.authority.boundObject.get(), &same) == S_OK && same &&
                            witness.authority.boundObject->GetBasicInformation(&retiredBasic) == HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE) &&
                            metadata->GetMetadataSnapshot(nullptr, &retiredMetadata) == HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE),
                        L"Retired witnesses preserve immutable identity but reject further live metadata reads."))
        return false;

    // Exercise legacy disposition while the retired witness COM objects stay alive.
    handoff.reset();
    strict.authority.boundObject.reset();
    wil::unique_handle deletion(
        CreateFileW(directory.c_str(), DELETE, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT, nullptr));
    FILE_DISPOSITION_INFO disposition{};
    disposition.DeleteFile = TRUE;
    if (! state.Require(deletion && SetFileInformationByHandle(deletion.get(), FileDispositionInfo, &disposition, sizeof(disposition)) != FALSE,
                        L"The retired metadata COM witness must not block legacy directory disposition."))
        return false;
    deletion.reset();
    return state.Require(CreateDirectoryW(directory.c_str(), nullptr) != FALSE,
                         L"The exact directory name must be reusable while retired witness COM references remain alive.");
}

[[nodiscard]] bool ValidateDeferredDirectoryDeleteIdentity(FolderWindow::FileOperationState& operations,
                                                           const wil::com_ptr<IFileSystem>& fileSystem,
                                                           const std::filesystem::path& root,
                                                           const FolderWindow::FileOperationState::MutationInterlockScopes& independentScopes,
                                                           SelfTest::CaseState& state) noexcept
{
    using State           = FolderWindow::FileOperationState;
    const auto selected   = root / L"queued-directory-delete";
    const auto moved      = root / L"queued-directory-delete-original";
    const auto peer       = root / L"queued-directory-delete-peer";
    const auto peerMoved  = root / L"queued-directory-delete-peer-moved";
    const auto leaseMoved = root / L"confirmed-directory-delete-moved";
    if (! state.Require(SelfTest::RemoveAll(moved) && SelfTest::RemoveAll(peerMoved) && SelfTest::RemoveAll(leaseMoved) && RecreateEmptyDirectory(selected) &&
                            RecreateEmptyDirectory(peer) && SelfTest::WriteTextFile(selected / L"original.txt", "original-directory-content") &&
                            SelfTest::WriteTextFile(peer / L"peer.txt", "unchanged-peer-content"),
                        L"Deferred directory Delete requires its own synthetic root."))
        return false;
    const auto makeTask = [&](uint64_t id, bool includePeer, bool prepare = true) -> std::unique_ptr<State::Task>
    {
        State::OperationConstruction construction{};
        construction.operation           = FILESYSTEM_DELETE;
        construction.fileSystem          = fileSystem;
        construction.sourcePaths         = includePeer ? std::vector<std::filesystem::path>{peer, selected} : std::vector<std::filesystem::path>{selected};
        construction.flags               = FILESYSTEM_FLAG_RECURSIVE;
        construction.sourcePluginId      = kPluginIdLocal;
        construction.sourcePluginShortId = L"file";
        State::OperationAdmission admission;
        if (! state.Require(State::BuildOperationAdmission(std::move(construction), admission) == S_OK,
                            L"Directory Delete must use canonical Local admission construction."))
            return {};
        auto task     = std::make_unique<State::Task>(operations);
        task->_taskId = id;
        task->_admissionKey.store(id, std::memory_order_release);
        task->_operation      = FILESYSTEM_DELETE;
        task->_fileSystem     = fileSystem;
        task->_sourcePaths    = includePeer ? std::vector<std::filesystem::path>{peer, selected} : std::vector<std::filesystem::path>{selected};
        task->_sourcePluginId = kPluginIdLocal;
        task->_flags          = FILESYSTEM_FLAG_RECURSIVE;
        task->_permanentDeleteConfirmationPending = true;
        task->_permanentDeleteTestOverride        = 1;
        task->InitializeAdmissionInputs(std::make_shared<const FileOperations::FileOperationPlanGroup>(std::move(admission.plans)), {});
        if (prepare && ! state.Require(task->PrepareMutationInterlockScopes() == S_OK && task->LoadMutationInterlockAdmission()->hasDeferredDirectoryDeletes,
                                       L"Queued Local directory deletion retains metadata identity without a strict mutation lease."))
            return {};
        return task;
    };

    {
        auto refused = makeTask(0xD1EA202611010004ull, false, false);
        if (! refused)
            return false;
        try
        {
            const RedSalamander::TestSupport::ScopedEnvironmentVariable unsupported(L"REDSALAMANDER_FILEOPS_BIND_IDENTITY_UNSUPPORTED_PATH", selected.native());
            const RedSalamander::TestSupport::ScopedEnvironmentVariable deleteDenied(L"REDSALAMANDER_FILEOPS_BIND_DELETE_ACCESS_DENIED_PATH",
                                                                                     selected.native());
            const HRESULT refusedHr = refused->PrepareMutationInterlockScopes();
            const auto scopes       = refused->LoadMutationInterlockScopes();
            std::error_code ec;
            if (! state.Require(refusedHr == E_ACCESSDENIED &&
                                    (! scopes || std::ranges::none_of(*scopes, [](const auto& scope) noexcept { return scope.byNameCorroborated; })) &&
                                    std::filesystem::exists(selected / L"original.txt", ec) && ! ec,
                                L"Identity-less metadata must not grant a by-name route when the strict Delete open is denied."))
                return false;
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
        catch (const std::system_error& error)
        {
            // The shared environment guard reports setup failure by exception in this noexcept witness.
            Debug::Error(L"Directory Delete denial fixture could not set its scoped injection: {}", error.code().value());
            return state.Require(false, L"The scoped directory Delete denial injection could not be installed.");
        }
    }

    auto changed = makeTask(0xD1EA202611010001ull, true);
    if (! changed)
        return false;
    if (! state.Require(MoveFileExW(selected.c_str(), moved.c_str(), MOVEFILE_WRITE_THROUGH) != FALSE && RecreateEmptyDirectory(selected) &&
                            SelfTest::WriteTextFile(selected / L"replacement.txt", "replacement-directory-content"),
                        L"An unconfirmed queued directory keeps a nonblocking identity witness while names can change."))
        return false;
    const HRESULT changedHr = changed->AcquireDeferredDirectoryDeleteAuthoritiesForExecution();
    std::error_code ec;
    if (! state.Require(changedHr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) && changed->_deferredDirectoryDeleteBindings.empty() &&
                            std::filesystem::exists(moved / L"original.txt", ec) && ! ec && std::filesystem::exists(selected / L"replacement.txt", ec) &&
                            ! ec && MoveFileExW(peer.c_str(), peerMoved.c_str(), MOVEFILE_WRITE_THROUGH) != FALSE &&
                            std::filesystem::exists(peerMoved / L"peer.txt", ec) && ! ec,
                        L"A queued directory replacement must be refused before touching either tree and release partial strict leases."))
        return false;
    changed.reset();

    auto confirmed = makeTask(0xD1EA202611010002ull, false);
    if (! confirmed)
        return false;
    State::MutationInterlockRelations relations{};
    relations.advisoryDecided = true;
    operations.PublishMutationInterlockDecision(*confirmed, relations);
    operations.PublishPreparedMutationInterlock(*confirmed);
    const auto key = confirmed->LoadMutationInterlockAdmission()->key;
    std::shared_ptr<const State::SelectedAdmissionView> accepted;
    const auto releaseReservation = wil::scope_exit([&]() noexcept
    {
        confirmed->_deferredDirectoryDeleteBindings.clear();
        if (accepted)
            operations.LeaveOperation(*confirmed, accepted);
        operations.WithdrawPreparedMutationInterlock(key);
    });
    auto entered                  = operations.EnterOperation(*confirmed, {});
    accepted                      = std::move(entered.accepted);
    if (! state.Require(entered.outcome == State::OperationEntryOutcome::Entered && confirmed->HasReservedOperation() && ! confirmed->HasEnteredOperation() &&
                            confirmed->GetLifecyclePhase() == FileOperations::TaskLifecyclePhase::PreparingExecution &&
                            confirmed->AcquireDeferredDirectoryDeleteAuthoritiesForExecution() == S_OK,
                        L"Directory confirmation reserves overlapping admission and acquires the checked strict lease before execution."))
        return false;
    auto readScope                    = accepted->interlock->scopes->front();
    readScope.access                  = FileOperations::MutationInterlockAccess::ReadSource;
    readScope.exactDeleteAuthority    = false;
    readScope.deferredDirectoryDelete = false;
    State::MutationInterlockAdmission reader{};
    reader.key                            = 0xD1EA202611010003ull;
    reader.scopes                         = std::make_shared<const State::MutationInterlockScopes>(State::MutationInterlockScopes{std::move(readScope)});
    reader.isTransfer                     = true;
    reader.relations.advisoryDecided      = true;
    reader.relations.concurrentKeys[0]    = accepted->interlock->key;
    reader.relations.concurrentCount      = 1u;
    auto directory                        = *accepted->interlock;
    directory.relations.concurrentKeys[0] = reader.key;
    directory.relations.concurrentCount   = 1u;
    std::array<State::ActiveMutationInterlock, 1> directoryActive{{{.taskId = confirmed->_taskId, .admission = accepted->interlock}}};
    std::array<State::ActiveMutationInterlock, 1> readerActive{
        {{.taskId = reader.key, .admission = std::make_shared<const State::MutationInterlockAdmission>(reader)}}};
    uint64_t comparisons = 0u;
    const auto advice    = operations.FindSameHostOverlap(reader.key, reader);
    if (! state.Require(! State::MutationInterlocksPermit(reader.key, reader, directoryActive, comparisons) &&
                            ! State::MutationInterlocksPermit(confirmed->_taskId, directory, readerActive, comparisons) && ! advice.CanRunConcurrently(),
                        L"Directory Delete and overlapping reads serialize in both directions despite stale concurrent grants and offer no Run bypass."))
        return false;
    reader.scopes = std::make_shared<const State::MutationInterlockScopes>(independentScopes);
    if (! state.Require(State::MutationInterlocksPermit(reader.key, reader, directoryActive, comparisons),
                        L"A directory reservation must not block unrelated scopes in the same Local identity domain."))
        return false;
    // Readiness is frozen before the queued task obtains its reservation. The receipt-only
    // update must still publish consent without reopening general plan mutation.
    confirmed->_selectedRootReadinessStatus.store(S_OK, std::memory_order_release);
    confirmed->_selectedRootReadinessComplete.store(true, std::memory_order_release);
    if (! state.Require(confirmed->RunPermanentDeleteConfirmation() == S_OK, L"Deferred directory confirmation must succeed behind its strict lease."))
        return false;
    const auto consented = confirmed->LoadSelectedAdmission();
    const auto* deletion = consented->plans->size() == 1u ? std::get_if<FileOperations::DeletePlan>(&consented->plans->front()) : nullptr;
    if (! state.Require(consented->interlock == accepted->interlock && deletion && deletion->initialConsent.has_value() &&
                            deletion->initialConsent.value().taskNonce == confirmed->_taskId &&
                            deletion->initialConsent.value().kind == FileOperations::ConsentKind::PermanentDelete,
                        L"A reserved directory confirmation publishes its exact task receipt while preserving the active immutable interlock."))
        return false;
    const BOOL movedWhileConfirmed = MoveFileExW(selected.c_str(), leaseMoved.c_str(), MOVEFILE_WRITE_THROUGH);
    const DWORD moveError          = GetLastError();
    if (! state.Require(movedWhileConfirmed == FALSE && moveError == ERROR_SHARING_VIOLATION,
                        L"The confirmed directory keeps its strict pathname pin until deletion or cancellation."))
        return false;
    confirmed->_deferredDirectoryDeleteBindings.clear();
    if (! state.Require(MoveFileExW(selected.c_str(), leaseMoved.c_str(), MOVEFILE_WRITE_THROUGH) != FALSE,
                        L"Releasing the owned strict lease allows the directory to move before a successor is admitted."))
        return false;
    Debug::Perf::EmitCounter(L"FileOps.SelfTest.DeferredDirectoryDeleteIdentity");
    return true;
}

[[nodiscard]] bool ValidateQueuedAdmissionContracts(SelfTestState& fixture, SelfTest::CaseState& state) noexcept
{
    using State                         = FolderWindow::FileOperationState;
    const SelfTest::TestSandbox sandbox = SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::FileOperations, L"queued-admission-contracts");
    if (! state.Require(sandbox.IsValid() && fixture.fileOps && fixture.fsDummy, L"Queued admission contracts require the governed sandbox and Dummy fixture."))
        return false;
    FileOperations::QualifiedEndpoint sandboxEndpoint;
    if (! state.Require(fixture.fsLocal && FileOperations::TryQualifyEndpointForSelfTest(
                                               fixture.fsLocal, sandbox.root.native(), FILESYSTEM_DELETE, kPluginIdLocal, L"host/default", sandboxEndpoint),
                        L"Queued admission must inspect the acquired sandbox's Local identity support."))
        return false;
    const auto sandboxIdentity =
        FileOperations::BindObjectAuthority(fixture.fsLocal.get(),
                                            sandbox.root.native(),
                                            sandboxEndpoint.profileId,
                                            static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA));
    if (sandboxIdentity.state == FileOperations::ObjectBindingState::Unsupported)
        return state.Skip(
            L"The acquired test sandbox cannot supply the persistent no-follow identities required by Local admission and pin-handoff witnesses.");
    if (! state.Require(sandboxIdentity.state == FileOperations::ObjectBindingState::Bound,
                        L"The acquired sandbox must open successfully before Local identity admission can be tested."))
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
    auto candidate = std::make_unique<State::Task>(*privateOperations);
    auto peer      = std::make_unique<State::Task>(*privateOperations);
    // General consent-key ownership remains applicable to Rename. Two Copy/Move
    // admissions now serialize regardless of consent; the shared-transfer controls
    // cover that policy independently of these registration-replacement checks.
    candidate->_operation      = FILESYSTEM_RENAME;
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
    state.Require(originalCandidateOwner && ! originalCandidateOwner->isTransfer && originalCandidateOwner->scopes->data() == movedStorage,
                  L"Freezing the Rename consent fixture must move the authority vector without copying its storage.");
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
    if (! ValidateDeferredDirectoryDeleteIdentity(*privateOperations, fixture.fsLocal, root, ownedScopes, state))
        return false;
    if (! ValidateLocalDirectoryPinHandoff(fixture.fsLocal, root, state))
        return false;
    if (! ValidateClipboardReadyStaleToken(fixture, state, sourcePath, destinationFolder))
        return false;

    return state.failure.empty();
}
