// Included by FolderWindow.FileOperations.Tests.Runner.cpp; the reducer matrix is deterministic and
// filesystem-free. Task callback/finalizer and provider witnesses are separated below.

#include <array>
#include <atomic>
#include <initializer_list>
#include <limits>
#include <thread>

[[nodiscard]] bool RunFileOperationLocalAncestryProbeSelfTest(SelfTestState& state) noexcept
{
    const auto root = state.tempRoot / L"local-ancestry-probe";
    if (! RecreateEmptyDirectory(root / L"real" / L"tree" / L"inner") || ! RecreateEmptyDirectory(root / L"outside" / L"inner"))
        return false;
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    if (! TryCreateJunction(root / L"alias", root / L"real"))
        return false;
    const auto identity = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem();
    const auto source   = root / L"alias" / L"tree";
    const auto probe    = [&](const std::filesystem::path& destination, bool cancel = false) noexcept
    { return FileOperations::ProbeLocalDirectoryTransferAncestry(source.native(), destination.native(), identity, [cancel]() noexcept { return cancel; }); };
    const auto inside = [](const FileOperations::LocalDirectoryAncestryProbeResult& observation) noexcept
    { return observation.destinationInsideSource && observation.status == HRESULT_FROM_WIN32(ERROR_INVALID_NAME); };
    const auto outside = [](const FileOperations::LocalDirectoryAncestryProbeResult& observation) noexcept
    { return ! observation.destinationInsideSource && observation.status == S_OK; };
    const auto unanswered = [](const FileOperations::LocalDirectoryAncestryProbeResult& observation, HRESULT expectedStatus) noexcept
    { return ! observation.destinationInsideSource && observation.status == expectedStatus; };
    bool passed = inside(probe(root / L"real" / L"tree" / L"inner" / L"tree")) && inside(probe(root / L"real" / L"tree")) &&
                  inside(probe(root / L"real" / L"tree" / L"missing-parent" / L"tree")) && outside(probe(root / L"outside" / L"inner" / L"tree")) &&
                  unanswered(probe(root / L"outside" / L"inner" / L"tree", true), HRESULT_FROM_WIN32(ERROR_CANCELLED));
    // The source's interior link points outward. A non-follow transfer cannot reach that
    // marker by traversal and must not be rejected as a recursive self-copy.
    if (! TryCreateJunction(root / L"real" / L"tree" / L"outside", root / L"outside"))
        return false;
    passed = passed && outside(probe(root / L"outside" / L"inner" / L"tree"));
    if (! SelfTest::WriteTextFile(root / L"blocked-parent", "fixture"))
        return false;
    passed = passed && unanswered(probe(root / L"blocked-parent" / L"tree"), HRESULT_FROM_WIN32(ERROR_DIRECTORY));
    // Exercise extended input and an ordinary spelling beyond MAX_PATH in the same sandbox.
    auto deep = root / L"real" / L"tree";
    for (unsigned int level = 0u; level < 6u; ++level)
        deep /= L"synthetic-long-directory-component-0123456789";
    std::error_code error;
    std::filesystem::create_directories(std::filesystem::path(Common::Paths::ToExtendedWin32Path(deep.native())), error);
    if (error)
        return false;
    passed = passed && inside(probe(deep / L"missing" / L"tree")) &&
             inside(FileOperations::ProbeLocalDirectoryTransferAncestry(
                 Common::Paths::ToExtendedWin32Path(source.native()), Common::Paths::ToExtendedWin32Path((deep / L"tree").native()), identity, []() noexcept {
        return false;
    }));
    const auto fenceParent = root / L"control-fence";
    if (! RecreateEmptyDirectory(fenceParent))
        return false;
    bool retirementAttempted = false;
    bool retirementRefused   = false;
    std::filesystem::path fencedControl;
    const auto fenced = FileOperations::ProbeLocalDirectoryTransferAncestry(source.native(),
                                                                            (fenceParent / L"target").native(),
                                                                            identity,
                                                                            [&]() noexcept
    {
        if (! retirementAttempted)
        {
            std::error_code listError;
            std::filesystem::directory_iterator entry(fenceParent, listError);
            const std::filesystem::directory_iterator end;
            for (; ! listError && entry != end; entry.increment(listError))
            {
                if (! entry->path().filename().native().starts_with(L"inside-source-check."))
                    continue;
                fencedControl              = entry->path();
                const std::wstring control = Common::Paths::ToExtendedWin32Path(fencedControl.native());
                const BOOL deleted         = DeleteFileW(control.c_str());
                const DWORD deleteError    = deleted ? ERROR_SUCCESS : GetLastError();
                const std::wstring renamed = control + L".renamed";
                const BOOL moved           = MoveFileW(control.c_str(), renamed.c_str());
                const DWORD moveError      = moved ? ERROR_SUCCESS : GetLastError();
                retirementAttempted        = true;
                retirementRefused          = ! deleted && ! moved && deleteError == ERROR_SHARING_VIOLATION && moveError == ERROR_SHARING_VIOLATION;
                break;
            }
        }
        return false;
    });
    std::error_code fencedExistsError;
    passed = passed && outside(fenced) && retirementAttempted && retirementRefused && ! std::filesystem::exists(fencedControl, fencedExistsError) &&
             ! fencedExistsError;

    const auto movedParent   = root / L"control-parent";
    const auto renamedParent = root / L"control-parent-renamed";
    if (! RecreateEmptyDirectory(movedParent))
        return false;
    bool parentAttempted  = false;
    bool parentMoved      = false;
    DWORD parentMoveError = ERROR_SUCCESS;
    const auto moved      = FileOperations::ProbeLocalDirectoryTransferAncestry(source.native(),
                                                                                (movedParent / L"target").native(),
                                                                                identity,
                                                                                [&]() noexcept
    {
        if (! parentAttempted)
        {
            std::error_code listError;
            std::filesystem::directory_iterator entry(movedParent, listError);
            const std::filesystem::directory_iterator end;
            for (; ! listError && entry != end; entry.increment(listError))
            {
                if (! entry->path().filename().native().starts_with(L"inside-source-check."))
                    continue;
                const std::wstring from = Common::Paths::ToExtendedWin32Path(movedParent.native());
                const std::wstring to   = Common::Paths::ToExtendedWin32Path(renamedParent.native());
                parentMoved             = MoveFileW(from.c_str(), to.c_str()) != FALSE;
                parentMoveError         = parentMoved ? ERROR_SUCCESS : GetLastError();
                parentAttempted         = true;
                break;
            }
        }
        return false;
    });
    passed = passed && parentAttempted &&
             (parentMoved ? unanswered(moved, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) || unanswered(moved, HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
                          : outside(moved) && (parentMoveError == ERROR_SHARING_VIOLATION || parentMoveError == ERROR_ACCESS_DENIED));
    const auto missingSource = FileOperations::ProbeLocalDirectoryTransferAncestry(
        (root / L"missing-source").native(), (fenceParent / L"target").native(), identity, []() noexcept { return false; });
    passed = passed && unanswered(missingSource, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
    // Drive the real task preflight too: no provider IDs, different source spelling, and a
    // destination subtree that does not yet exist. Refusal precedes parent creation.
    if (! state.fileOps || ! state.fsLocal)
        return false;
    const auto missingFolder = root / L"real" / L"tree" / L"not-created";
    RedSalamander::TestSupport::ScopedEnvironmentVariable noIdentity(kSelfTestEnvBindIdentityUnsupportedPath, root.native());
    FileOperations::QualifiedEndpoint endpoint{.pluginId                 = L"builtin/file-system",
                                               .instanceId               = L"host/default",
                                               .profileId                = L"local",
                                               .rootId                   = L"fixture-ancestry-volume",
                                               .pathIdentity             = identity,
                                               .verificationHostReadback = true,
                                               .cancellationRouteClass   = FileOperations::CancellationRouteClass::Bounded};
    FileOperations::TransferPlan transfer{};
    transfer.intent                         = FileOperations::TransferIntent::Move;
    transfer.strategy                       = FileOperations::OperationStrategy::Managed;
    transfer.sourceEndpoint                 = endpoint;
    transfer.destinationEndpoint            = endpoint;
    transfer.selectedItems                  = {{.providerPath = source.native()}};
    transfer.destination.providerFolderPath = missingFolder.native();
    using State                             = FolderWindow::FileOperationState;
    State::Task task(*state.fileOps);
    task._operation                = FILESYSTEM_MOVE;
    task._executionMode            = State::ExecutionMode::PerItem;
    task._fileSystem               = state.fsLocal;
    task._sourcePaths              = {source};
    task._sourcePluginId           = endpoint.pluginId;
    task._sourcePluginShortId      = L"file";
    task._sourceInstanceContext    = endpoint.instanceId;
    task._sourcePathAttributesHint = {FILE_ATTRIBUTE_DIRECTORY};
    task.InitializeAdmissionInputs(std::make_shared<const FileOperations::FileOperationPlanGroup>(
                                       FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}}),
                                   {.fileSystem      = state.fsLocal,
                                    .folder          = missingFolder,
                                    .pluginId        = endpoint.pluginId,
                                    .pluginShortId   = L"file",
                                    .instanceContext = endpoint.instanceId});
    task.InitializeSourceItemResultBuilders();
    const auto diagnosticCleanup = wil::scope_exit([&]() noexcept { state.fileOps->DebugRemoveDiagnosticsForTask(task._taskId); });
    const HRESULT prepareHr      = task.PrepareMutationInterlockScopes();
    const HRESULT executeHr      = FAILED(prepareHr) ? prepareHr : task.ExecuteOperation();
    std::error_code missingError;
    passed = passed && executeHr == HRESULT_FROM_WIN32(ERROR_INVALID_NAME) && ! std::filesystem::exists(missingFolder, missingError) && ! missingError;
    wil::com_ptr<IFileSystemIO> io;
    if (! state.fsLocal.try_query_to(io.put()))
        return false;
    {
        const auto realView                   = root / L"cross-view-real";
        const auto aliasView                  = root / L"cross-view-alias";
        const auto viewSource                 = aliasView / L"album";
        const auto existingNested             = realView / L"album" / L"inside" / L"album";
        const auto missingNestedParent        = realView / L"album" / L"not-created";
        const auto distinctParent             = root / L"cross-view-distinct";
        constexpr std::string_view kViewBytes = "bound-cross-view-source";
        if (! RecreateEmptyDirectory(existingNested) || ! RecreateEmptyDirectory(distinctParent) || ! TryCreateJunction(aliasView, realView) ||
            ! WriteFileTextFsIo(io, realView / L"album" / L"track.bin", std::string(kViewBytes)))
            return false;
        RedSalamander::TestSupport::ScopedEnvironmentVariable restoreIdentity(kSelfTestEnvBindIdentityUnsupportedPath, std::nullopt);
        RedSalamander::TestSupport::ScopedEnvironmentVariable namespaceView(L"REDSALAMANDER_FILEOPS_BIND_IDENTITY_NAMESPACE_PATH", aliasView.native());
        auto sourceEndpoint        = endpoint;
        auto destinationEndpoint   = endpoint;
        sourceEndpoint.rootId      = L"fixture-cross-view-source";
        destinationEndpoint.rootId = L"fixture-cross-view-destination";
        auto ready                 = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                  state.fsLocal.get(),
                                                                                  sourceEndpoint,
                                                                                  destinationEndpoint,
                                                                                  FileOperations::TransferIntent::Copy,
                                                                                  viewSource.native(),
                                                                                  (realView / L"album").native());
        bool sameViewObject        = true;
        bool sameViewRevision      = false;
        const HRESULT crossCheckHr =
            ready.destination.has_value()
                ? FileOperations::CrossCheckBoundObjectIdentity(ready.source, ready.destination.value().authority, sameViewObject, sameViewRevision)
                : E_UNEXPECTED;
        constexpr FileSystemBindFlags viewBindFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
        auto rebound =
            FileOperations::RevalidateObjectAuthority(state.fsLocal.get(), viewSource.native(), sourceEndpoint.profileId, viewBindFlags, ready.source);
        if (ready.state != FileOperations::TransferSafetyState::Ready || ! ready.source.boundObject || ! ready.destination.has_value() ||
            ! ready.destination.value().authority.boundObject || FAILED(crossCheckHr) || sameViewObject ||
            rebound.state != FileOperations::ObjectRevalidationState::Same)
            return false;
        // Drop this observation's pins before exercising the real task, as a remote view
        // may enforce directory-handle sharing independently of its reported ID namespace.
        ready                  = {};
        rebound                = {};
        const auto runViewTask = [&](const std::filesystem::path& destinationParent) noexcept
        {
            FileOperations::TransferPlan viewPlan{};
            viewPlan.intent                         = FileOperations::TransferIntent::Copy;
            viewPlan.strategy                       = FileOperations::OperationStrategy::Copy;
            viewPlan.sourceEndpoint                 = sourceEndpoint;
            viewPlan.destinationEndpoint            = destinationEndpoint;
            viewPlan.selectedItems                  = {{.providerPath = viewSource.native()}};
            viewPlan.destination.providerFolderPath = destinationParent.native();
            State::Task viewTask(*state.fileOps);
            viewTask._operation                = FILESYSTEM_COPY;
            viewTask._executionMode            = State::ExecutionMode::PerItem;
            viewTask._fileSystem               = state.fsLocal;
            viewTask._sourcePaths              = {viewSource};
            viewTask._sourcePluginId           = sourceEndpoint.pluginId;
            viewTask._sourcePluginShortId      = L"file";
            viewTask._sourceInstanceContext    = sourceEndpoint.instanceId;
            viewTask._sourcePathAttributesHint = {FILE_ATTRIBUTE_DIRECTORY};
            viewTask.InitializeAdmissionInputs(std::make_shared<const FileOperations::FileOperationPlanGroup>(
                                                   FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(viewPlan)}}),
                                               {.fileSystem      = state.fsLocal,
                                                .folder          = destinationParent,
                                                .pluginId        = destinationEndpoint.pluginId,
                                                .pluginShortId   = L"file",
                                                .instanceContext = destinationEndpoint.instanceId});
            viewTask.InitializeSourceItemResultBuilders();
            const HRESULT prepared = viewTask.PrepareMutationInterlockScopes();
            return FAILED(prepared) ? prepared : viewTask.ExecuteOperation();
        };
        const HRESULT nestedHr = runViewTask(missingNestedParent);
        std::error_code nestedExistsError;
        const bool nestedParentExists = std::filesystem::exists(missingNestedParent, nestedExistsError);
        const HRESULT distinctHr      = runViewTask(distinctParent);
        std::string sourceBytes;
        std::string distinctBytes;
        passed = passed && nestedHr == HRESULT_FROM_WIN32(ERROR_INVALID_NAME) && ! nestedParentExists && ! nestedExistsError && distinctHr == S_OK &&
                 ReadFileTextFsIo(io, viewSource / L"track.bin", sourceBytes) && sourceBytes == kViewBytes &&
                 ReadFileTextFsIo(io, distinctParent / L"album" / L"track.bin", distinctBytes) && distinctBytes == kViewBytes;

        // Differing Local roots do not turn negative namespace IDs into destructive
        // distinctness. Exercise the real Managed admission/prompt, never a copied predicate.
        const auto distinctFileParent = distinctParent / L"file-collision";
        if (! RecreateEmptyDirectory(distinctFileParent) || ! WriteFileTextFsIo(io, distinctFileParent / L"track.bin", "existing-destination"))
            return false;
        for (const unsigned int control : {0u, 1u, 2u})
        {
            const bool sameEntry           = control == 0u;
            const bool identitylessCleanup = control == 2u;
            const auto selectedFile        = viewSource / L"track.bin";
            const auto targetParent        = sameEntry ? realView / L"album" : distinctFileParent;
            const auto targetFile          = targetParent / selectedFile.filename();
            RedSalamander::TestSupport::ScopedEnvironmentVariable sourceIdentity(
                kSelfTestEnvBindIdentityUnsupportedPath, identitylessCleanup ? std::optional<std::wstring_view>(selectedFile.native()) : std::nullopt);
            if (sameEntry)
            {
                RedSalamander::TestSupport::ScopedEnvironmentVariable comparableIds(L"REDSALAMANDER_FILEOPS_BIND_IDENTITY_NAMESPACE_PATH", std::nullopt);
                const auto positive = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                   state.fsLocal.get(),
                                                                                   sourceEndpoint,
                                                                                   destinationEndpoint,
                                                                                   FileOperations::TransferIntent::Move,
                                                                                   selectedFile.native(),
                                                                                   targetFile.native());
                if (positive.state != FileOperations::TransferSafetyState::SameObject)
                    return false;
            }
            auto fileGuard        = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                 state.fsLocal.get(),
                                                                                 sourceEndpoint,
                                                                                 destinationEndpoint,
                                                                                 FileOperations::TransferIntent::Move,
                                                                                 selectedFile.native(),
                                                                                 targetFile.native());
            bool sameFileObject   = true;
            bool sameFileRevision = false;
            const HRESULT fileCheck =
                fileGuard.destination.has_value()
                    ? FileOperations::CrossCheckBoundObjectIdentity(fileGuard.source, fileGuard.destination.value().authority, sameFileObject, sameFileRevision)
                    : E_UNEXPECTED;
            if (identitylessCleanup ? fileGuard.state != FileOperations::TransferSafetyState::Unsupported || fileGuard.source.boundObject
                                    : fileGuard.state != FileOperations::TransferSafetyState::Ready || ! fileGuard.source.boundObject ||
                                          ! fileGuard.destination.has_value() || FAILED(fileCheck) || sameFileObject)
                return false;
            fileGuard = {};

            FileOperations::TransferPlan filePlan{};
            filePlan.intent                         = FileOperations::TransferIntent::Move;
            filePlan.strategy                       = FileOperations::OperationStrategy::Managed;
            filePlan.sourceEndpoint                 = sourceEndpoint;
            filePlan.destinationEndpoint            = destinationEndpoint;
            filePlan.selectedItems                  = {{.providerPath = selectedFile.native()}};
            filePlan.destination.providerFolderPath = targetParent.native();
            State::Task fileTask(*state.fileOps);
            fileTask._operation                = FILESYSTEM_MOVE;
            fileTask._executionMode            = State::ExecutionMode::PerItem;
            fileTask._fileSystem               = state.fsLocal;
            fileTask._sourcePaths              = {selectedFile};
            fileTask._sourcePluginId           = sourceEndpoint.pluginId;
            fileTask._sourcePluginShortId      = L"file";
            fileTask._sourceInstanceContext    = sourceEndpoint.instanceId;
            fileTask._sourcePathAttributesHint = {FILE_ATTRIBUTE_NORMAL};
            fileTask._perItemMaxConcurrency    = 1u;
            fileTask.InitializeAdmissionInputs(std::make_shared<const FileOperations::FileOperationPlanGroup>(
                                                   FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(filePlan)}}),
                                               {.fileSystem      = state.fsLocal,
                                                .folder          = targetParent,
                                                .pluginId        = destinationEndpoint.pluginId,
                                                .pluginShortId   = L"file",
                                                .instanceContext = destinationEndpoint.instanceId});
            fileTask.InitializeSourceItemResultBuilders();
            const auto fileDiagnostics = wil::scope_exit([&]() noexcept { state.fileOps->DebugRemoveDiagnosticsForTask(fileTask._taskId); });
            if (FAILED(fileTask.PrepareMutationInterlockScopes()))
                return false;
            using Action = State::Task::ConflictAction;
            std::atomic<uint32_t> promptCount{0u};
            std::atomic<bool> discardOffered{false};
            std::atomic<bool> ordinaryActions{true};
            HRESULT fileHr           = E_PENDING;
            const auto fileStartedAt = std::chrono::steady_clock::now();
            {
                const std::jthread responder([&](std::stop_token stop) noexcept
                {
                    uint64_t answeredGeneration = 0u;
                    const auto deadline         = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::seconds(5));
                    while (! stop.stop_requested() && std::chrono::steady_clock::now() < deadline)
                    {
                        if (const auto prompt = TryGetConflictPromptCopy(&fileTask);
                            prompt.has_value() && prompt->generation != answeredGeneration && ! prompt->metadataLoading && prompt->buttonsPublishable)
                        {
                            const auto end     = prompt->actions.begin() + static_cast<std::ptrdiff_t>(prompt->actionCount);
                            const auto offered = [&](Action action) noexcept { return std::find(prompt->actions.begin(), end, action) != end; };
                            discardOffered.store(offered(Action::DiscardSource), std::memory_order_release);
                            ordinaryActions.store(offered(Action::Overwrite) && offered(Action::KeepBoth) && offered(Action::Skip) && offered(Action::Cancel),
                                                  std::memory_order_release);
                            // Alias controls always Cancel, even on a vulnerable baseline. Only the distinct-file
                            // identityless control replaces bytes, then exercises the actual cleanup helper.
                            if (fileTask.SubmitConflictDecision(prompt->generation, identitylessCleanup ? Action::Overwrite : Action::Cancel, false))
                            {
                                answeredGeneration = prompt->generation;
                                promptCount.fetch_add(1u, std::memory_order_release);
                            }
                        }
                        Sleep(2);
                    }
                    if (! stop.stop_requested())
                        fileTask.RequestCancel();
                });
                fileHr = fileTask.ExecuteOperation();
            }
            const uint64_t fileElapsedUs  = Debug::Perf::ElapsedUs(fileStartedAt);
            bool retainedAfterPublication = true;
            if (identitylessCleanup)
            {
                static_cast<void>(fileTask.FinalizeTypedItemResults(fileHr));
                const std::scoped_lock resultsLock(fileTask._sourceItemStatusMutex);
                retainedAfterPublication = fileTask._sourceItemResultBuilders.size() == 1u && fileTask._sourceItemResultBuilders.front().terminal.has_value();
                if (retainedAfterPublication)
                {
                    const auto& terminal     = fileTask._sourceItemResultBuilders.front().terminal.value();
                    retainedAfterPublication = terminal.publication == FileOperations::PublicationState::Published &&
                                               terminal.sourceDisposition == FileOperations::SourceDisposition::Retained &&
                                               terminal.completion == FileOperations::ItemCompletion::Completed && SUCCEEDED(terminal.status);
                }
            }
            std::string retainedSource;
            std::string retainedDestination;
            const bool filePassed =
                (identitylessCleanup ? SUCCEEDED(fileHr) : fileHr == HRESULT_FROM_WIN32(ERROR_CANCELLED)) && retainedAfterPublication &&
                promptCount.load(std::memory_order_acquire) == 1u && ! discardOffered.load(std::memory_order_acquire) &&
                ordinaryActions.load(std::memory_order_acquire) && fileTask._alreadyThereSourceRemovedCount.load(std::memory_order_acquire) == 0u &&
                ReadFileTextFsIo(io, selectedFile, retainedSource) && retainedSource == kViewBytes && ReadFileTextFsIo(io, targetFile, retainedDestination) &&
                retainedDestination == (sameEntry || identitylessCleanup ? kViewBytes : std::string_view("existing-destination"));
            Debug::Perf::EmitDurationUs(L"FileOps.SelfTest.LocalCrossRootDistinctness",
                                        identitylessCleanup ? L"identityless-replace" : (sameEntry ? L"same-entry" : L"different-entry"),
                                        fileElapsedUs,
                                        promptCount.load(std::memory_order_acquire),
                                        discardOffered.load(std::memory_order_acquire) ? 1u : 0u,
                                        filePassed ? S_OK : E_FAIL);
            passed = filePassed && passed;
        }
    }
    for (const bool unsupported : {true, false})
    {
        for (const FileSystemOperation operation : {FILESYSTEM_COPY, FILESYSTEM_MOVE})
        {
            const auto caseRoot               = root / std::format(L"same-path-{}-{}", unsupported, static_cast<unsigned int>(operation));
            const auto selected               = caseRoot / L"album";
            const auto originalFile           = selected / L"track.bin";
            constexpr std::string_view kBytes = "same-path-folder-must-survive";
            if (! RecreateEmptyDirectory(selected) || ! WriteFileTextFsIo(io, originalFile, std::string(kBytes)))
                return false;
            RedSalamander::TestSupport::ScopedEnvironmentVariable identitySupport(
                kSelfTestEnvBindIdentityUnsupportedPath, unsupported ? std::optional<std::wstring_view>(selected.native()) : std::nullopt);
            FileOperations::TransferPlan samePathPlan{};
            samePathPlan.intent         = operation == FILESYSTEM_COPY ? FileOperations::TransferIntent::Copy : FileOperations::TransferIntent::Move;
            samePathPlan.strategy       = operation == FILESYSTEM_COPY ? FileOperations::OperationStrategy::Copy : FileOperations::OperationStrategy::Managed;
            samePathPlan.sourceEndpoint = endpoint;
            samePathPlan.destinationEndpoint            = endpoint;
            samePathPlan.selectedItems                  = {{.providerPath = selected.native()}};
            samePathPlan.destination.providerFolderPath = caseRoot.native();
            State::Task samePathTask(*state.fileOps);
            samePathTask._operation                = operation;
            samePathTask._executionMode            = State::ExecutionMode::PerItem;
            samePathTask._fileSystem               = state.fsLocal;
            samePathTask._sourcePaths              = {selected};
            samePathTask._sourcePluginId           = endpoint.pluginId;
            samePathTask._sourcePluginShortId      = L"file";
            samePathTask._sourceInstanceContext    = endpoint.instanceId;
            samePathTask._sourcePathAttributesHint = {FILE_ATTRIBUTE_DIRECTORY};
            samePathTask.InitializeAdmissionInputs(std::make_shared<const FileOperations::FileOperationPlanGroup>(
                                                       FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(samePathPlan)}}),
                                                   {.fileSystem      = state.fsLocal,
                                                    .folder          = caseRoot,
                                                    .pluginId        = endpoint.pluginId,
                                                    .pluginShortId   = L"file",
                                                    .instanceContext = endpoint.instanceId});
            samePathTask.InitializeSourceItemResultBuilders();
            const HRESULT samePrepareHr = samePathTask.PrepareMutationInterlockScopes();
            const HRESULT sameExecuteHr = FAILED(samePrepareHr) ? samePrepareHr : samePathTask.ExecuteOperation();
            std::string originalBytes;
            std::string copiedBytes;
            std::wstring sibling;
            const HRESULT siblingHr      = Common::Paths::BuildUniqueSiblingPathCandidate(selected.native(), true, 2u, sibling);
            const HRESULT expectedMoveHr = HRESULT_FROM_WIN32(unsupported ? ERROR_ALREADY_EXISTS : ERROR_INVALID_NAME);
            const bool preserved         = ReadFileTextFsIo(io, originalFile, originalBytes) && originalBytes == kBytes;
            std::error_code siblingExistsError;
            const bool siblingExists = std::filesystem::exists(std::filesystem::path(sibling), siblingExistsError);
            const bool correct       = operation == FILESYSTEM_COPY
                                           ? sameExecuteHr == S_OK && SUCCEEDED(siblingHr) &&
                                                 ReadFileTextFsIo(io, std::filesystem::path(sibling) / L"track.bin", copiedBytes) && copiedBytes == kBytes
                                           : sameExecuteHr == expectedMoveHr && ! siblingExists && ! siblingExistsError;
            if (! preserved || ! correct)
                AppendLog(std::format(L"Same-path folder unsupported={} operation={} prepare=0x{:08X} execute=0x{:08X} preserved={} correct={}",
                                      unsupported,
                                      static_cast<unsigned int>(operation),
                                      static_cast<unsigned long>(samePrepareHr),
                                      static_cast<unsigned long>(sameExecuteHr),
                                      preserved,
                                      correct));
            passed = passed && preserved && correct;
        }
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.LocalAncestryProbe",
                      L"alias;missing-parent;outward-link;long-path;cancel;invalid-control-parent;control-retirement;control-parent-rename;missing-source",
                      0u,
                      21u,
                      0u,
                      passed ? S_OK : E_FAIL);
    return passed;
}

[[nodiscard]] bool RunFileOperationSharedTransferSelfTests(SelfTestState& state, std::wstring& failure) noexcept
{
    using State = FolderWindow::FileOperationState;
    if (! LoadPlugins(state) || ! state.fsLocal)
        return false;
    wil::com_ptr<IFileSystemIO> io;
    if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(io.addressof()))) || ! io)
        return false;
    const auto root = state.tempRoot / L"shared-parallel-transfers";
    if (! RecreateEmptyDirectory(root))
        return false;
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    struct Scenario final
    {
        uint32_t limit;
        size_t jobs;
        size_t filesPerJob;
        FileSystemOperation operation;
        bool cancelLast;
        bool directoryIdsOnly = false;
    };
    size_t ordinal = 0u;
    for (const Scenario scenario : {Scenario{1u, 4u, 1u, FILESYSTEM_COPY, false},
                                    Scenario{4u, 4u, 1u, FILESYSTEM_COPY, false},
                                    Scenario{4u, 5u, 1u, FILESYSTEM_COPY, true},
                                    Scenario{4u, 2u, 4u, FILESYSTEM_COPY, false},
                                    Scenario{4u, 4u, 1u, FILESYSTEM_MOVE, false, true},
                                    Scenario{4u, 4u, 1u, FILESYSTEM_MOVE, false, false}})
    {
        const auto caseRoot    = root / std::format(L"case-{}", ++ordinal);
        const auto destination = caseRoot / L"studio-nas-movies";
        if (! RecreateEmptyDirectory(destination))
            return false;
        std::vector<std::filesystem::path> sources;
        std::vector<std::string> expected;
        for (size_t index = 0u; index < scenario.jobs * scenario.filesPerJob; ++index)
        {
            sources.push_back(caseRoot / L"exports" / std::format(L"Moonrise-{}.bin", index));
            expected.emplace_back();
            if (! WriteTestFile(sources.back(), 4u * 1024u * 1024u + index) || ! ReadFileTextFsIo(io, sources.back(), expected.back()))
                return false;
        }
        RedSalamander::TestSupport::ScopedEnvironmentVariable unsupportedIds(
            kSelfTestEnvBindIdentityUnsupportedPath, scenario.directoryIdsOnly ? std::nullopt : std::optional<std::wstring_view>(destination.native()));
        RedSalamander::TestSupport::ScopedEnvironmentVariable unsupportedDirectoryIds(
            L"REDSALAMANDER_FILEOPS_BIND_DIRECTORY_IDENTITY_UNSUPPORTED_PATH",
            scenario.directoryIdsOnly ? std::optional<std::wstring_view>(destination.native()) : std::nullopt);
        if (! State::DebugRunSharedTransfersForSelfTest(
                state.fsLocal.get(), sources, destination, scenario.operation, scenario.limit, scenario.filesPerJob, scenario.cancelLast, failure))
            return false;
        size_t published = 0u;
        for (size_t index = 0u; index < sources.size(); ++index)
        {
            const auto target = destination / sources[index].filename();
            if (GetFileAttributesW(target.c_str()) == INVALID_FILE_ATTRIBUTES)
            {
                if (! scenario.cancelLast)
                {
                    failure = L"Every successful shared transfer must publish its distinct new destination.";
                    return false;
                }
                continue;
            }
            ++published;
            std::string actual;
            const bool sourcePresent = GetFileAttributesW(sources[index].c_str()) != INVALID_FILE_ATTRIBUTES;
            if (! ReadFileTextFsIo(io, target, actual) || actual != expected[index] || sourcePresent != (scenario.operation == FILESYSTEM_COPY))
            {
                failure = L"Shared transfers must preserve exact file content, keep Copy sources, and remove only successfully published Move sources.";
                return false;
            }
        }
        if (published != sources.size() - (scenario.cancelLast ? scenario.filesPerJob : 0u))
        {
            failure = L"Canceling a queued transfer must omit exactly its unpublished files.";
            return false;
        }
        AppendLog(std::format(L"SharedTransfers: limit={} jobs={} filesPerJob={} operation={} cancelQueued={} directoryIdsOnly={} published={} passed=true",
                              scenario.limit,
                              scenario.jobs,
                              scenario.filesPerJob,
                              static_cast<unsigned int>(scenario.operation),
                              scenario.cancelLast,
                              scenario.directoryIdsOnly,
                              published));
    }
    return true;
}

namespace FileOperationItemResultReducerSelfTestInternal
{
[[nodiscard]] bool EnsureFolder(IFileSystem* fileSystem, const std::wstring_view path) noexcept
{
    if (fileSystem == nullptr || path.empty())
        return false;
    wil::com_ptr<IFileSystemDirectoryOperations> directoryOperations;
    if (FAILED(fileSystem->QueryInterface(IID_PPV_ARGS(directoryOperations.addressof()))) || ! directoryOperations)
        return false;
    const std::wstring pathText(path);
    const HRESULT hr = directoryOperations->CreateDirectory(pathText.c_str());
    return SUCCEEDED(hr) || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
}

[[nodiscard]] bool WriteTextFile(IFileSystem* fileSystem, const std::wstring_view path, const std::string_view text) noexcept
{
    if (fileSystem == nullptr || path.empty() || text.size() > std::numeric_limits<unsigned long>::max())
        return false;
    wil::com_ptr<IFileSystemIO> fileSystemIo;
    if (FAILED(fileSystem->QueryInterface(IID_PPV_ARGS(fileSystemIo.addressof()))) || ! fileSystemIo)
        return false;
    const std::wstring pathText(path);
    wil::com_ptr<IFileWriter> writer;
    if (FAILED(fileSystemIo->CreateFileWriter(pathText.c_str(), FILESYSTEM_FLAG_NONE, writer.put())) || ! writer)
        return false;
    unsigned long written = 0u;
    return (text.empty() ||
            (SUCCEEDED(writer->Write(text.data(), static_cast<unsigned long>(text.size()), &written)) && written == static_cast<unsigned long>(text.size()))) &&
           SUCCEEDED(writer->Commit());
}
} // namespace FileOperationItemResultReducerSelfTestInternal

// Exercise the actual Local Native-Move producer -> Task callback -> dense finalizer seam
// when the admitted source disappears before the provider's first source read. ExecuteOperation
// is intentionally not involved: its final exact-source revalidation rejects an absent source
// before reaching the provider, so this fixture isolates the producer receipt contract. No rename
// was issued, so the item is a known failure with the real error (I30 fs-fileops#3, owner Q19).
[[nodiscard]] bool RunFileOperationNativeMoveMissingSourceTaskSelfTest(FolderWindow::FileOperationState& fileOperations,
                                                                       const wil::com_ptr<IFileSystem>& localFileSystem,
                                                                       const std::filesystem::path& testRoot) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;
    if (! localFileSystem || testRoot.empty())
        return false;

    const uint64_t runToken                 = (static_cast<uint64_t>(GetCurrentProcessId()) << 32u) ^ GetTickCount64();
    const std::filesystem::path root        = testRoot / std::format(L"native-move-missing-source-receipt-{:016X}", runToken);
    const std::filesystem::path source      = root / L"source.bin";
    const std::filesystem::path destination = root / L"destination";
    const std::filesystem::path moved       = destination / source.filename();
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(destination) || ! WriteTestFile(source, 4096u))
    {
        AppendLog(L"Native-Move missing-source fixture: failed to create its private Local paths.");
        const bool removed = SelfTest::RemoveAll(root);
        std::error_code cleanupError;
        const bool remains = std::filesystem::exists(root, cleanupError);
        if (! removed || cleanupError || remains)
            AppendLog(std::format(
                L"Native-Move missing-source fixture setup cleanup failed (removed={}, remains={}, error={}).", removed, remains, cleanupError.value()));
        return false;
    }
    bool cleanupVerified = false;
    const auto cleanup   = wil::scope_exit([&]() noexcept
    {
        if (! cleanupVerified)
        {
            const bool removed = SelfTest::RemoveAll(root);
            std::error_code cleanupError;
            const bool remains = std::filesystem::exists(root, cleanupError);
            if (! removed || cleanupError || remains)
                AppendLog(std::format(
                    L"Native-Move missing-source fixture fallback cleanup failed (removed={}, remains={}, error={}).", removed, remains, cleanupError.value()));
        }
    });

    FileOperations::QualifiedEndpoint endpoint{
        .pluginId               = L"builtin/file-system",
        .instanceId             = L"host/default",
        .profileId              = L"local",
        .rootId                 = L"local-test-volume",
        .pathIdentity           = FileSystemPathIdentity{.pathTextStableIdentity = true,
                                                         .componentComparison    = FileSystemPathComponentComparison::OrdinalIgnoreCase,
                                                         .preferredSeparator     = L'\\',
                                                         .acceptedSeparators     = L"\\/",
                                                         .casePreserving         = true,
                                                         .caseOnlyRename         = FileSystemPathCaseOnlyRename::Supported},
        .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
    };
    FileOperations::TransferPlan plan{};
    plan.intent                         = FileOperations::TransferIntent::Move;
    plan.strategy                       = FileOperations::OperationStrategy::Native;
    plan.sourceEndpoint                 = endpoint;
    plan.destinationEndpoint            = endpoint;
    plan.selectedItems                  = {{.providerPath = source.native()}};
    plan.destination.providerFolderPath = destination.native();
    auto plans                          = std::make_shared<const FileOperations::FileOperationPlanGroup>(
        FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(plan)}});

    auto taskStorage = std::make_unique<Task>(fileOperations);
    Task& task       = *taskStorage;
    task.StorePlans(plans);
    task._operation             = FILESYSTEM_MOVE;
    task._executionMode         = FolderWindow::FileOperationState::ExecutionMode::BulkItems;
    task._fileSystem            = localFileSystem;
    task._sourcePaths           = {source};
    task._taskId                = 0xF000000000000000ull | (runToken & 0x0FFFFFFFFFFFFFFFull);
    task._sourcePluginId        = endpoint.pluginId;
    task._sourcePluginShortId   = L"file";
    task._sourceInstanceContext = endpoint.instanceId;
    task._perItemMaxConcurrency = 1u;
    task.InitializeAdmissionInputs(plans,
                                   {.fileSystem      = localFileSystem,
                                    .folder          = destination,
                                    .pluginId        = endpoint.pluginId,
                                    .pluginShortId   = L"file",
                                    .instanceContext = endpoint.instanceId});
    task._sourcePathAttributesHint = {GetFileAttributesW(source.c_str())};
    task.InitializeSourceItemResultBuilders();
    const auto removeFixtureDiagnostics = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });
    if (task._sourcePathAttributesHint.front() == INVALID_FILE_ATTRIBUTES)
    {
        AppendLog(L"Native-Move missing-source fixture: source attributes were not captured before admission.");
        return false;
    }

    task.MarkSourceItemsMutationPossible();
    if (! DeleteFileW(MakeExtendedPathForSelfTest(source).c_str()))
    {
        AppendLog(std::format(L"Native-Move missing-source fixture: could not remove the source after the mutation gate (error {}).", GetLastError()));
        return false;
    }

    FileSystemOptions options{};
    options.sizeBytes              = sizeof(options);
    options.moveMode               = FILESYSTEM_MOVE_NATIVE_ONLY;
    options.operationControl       = &task;
    options.operationControlCookie = nullptr;
    HRESULT operationHr            = E_UNEXPECTED;
    {
        task._dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
        const auto callbackScope = wil::scope_exit([&]() noexcept { task._dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
        operationHr              = localFileSystem->MoveItem(source.c_str(), moved.c_str(), FILESYSTEM_FLAG_NONE, &options, &task, nullptr);
    }
    const HRESULT finalHr = task.FinalizeTypedItemResults(operationHr);

    const size_t builderCount   = task._sourceItemResultBuilders.size();
    const bool hasSingleBuilder = builderCount == 1u;
    const auto* builder         = hasSingleBuilder ? std::addressof(task._sourceItemResultBuilders.front()) : nullptr;
    const FileOperations::FileOperationItemResult* item =
        builder != nullptr && builder->terminal.has_value() ? std::addressof(builder->terminal.value()) : nullptr;
    const std::optional<FileSystemItemMutationResult>* receipt = builder != nullptr ? std::addressof(builder->mutation) : nullptr;
    const HRESULT callbackStatus                               = builder != nullptr ? builder->status.value_or(S_OK) : E_UNEXPECTED;
    std::error_code ec;
    const bool sourceAbsent = ! std::filesystem::exists(source, ec) && ! ec;
    ec.clear();
    const bool destinationAbsent     = ! std::filesystem::exists(moved, ec) && ! ec;
    const bool providerMissingSource = operationHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || operationHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    const bool matches = hasSingleBuilder && providerMissingSource && finalHr == operationHr && item != nullptr &&
                         item->completion == FileOperations::ItemCompletion::Failed && item->publication == FileOperations::PublicationState::NotPublished &&
                         item->sourceDisposition == FileOperations::SourceDisposition::Retained && item->status == operationHr &&
                         task._itemCompletedCallbackCount.load(std::memory_order_acquire) == 1u && callbackStatus == operationHr && receipt != nullptr &&
                         receipt->has_value() && receipt->value().outcomeKnown == TRUE && receipt->value().mutationCommitted == FALSE &&
                         receipt->value().originalStillPresent == TRUE && task._perf.conflictPromptCount == 0u && sourceAbsent && destinationAbsent;
    if (! matches)
    {
        AppendLog(std::format(L"Native-Move missing-source mismatch: operation=0x{:08X} final=0x{:08X} callback=0x{:08X} callbacks={} builders={} item={} "
                              L"itemStatus=0x{:08X} receipt={}/{}/{} prompts={} sourceAbsent={} destinationAbsent={}",
                              static_cast<unsigned long>(operationHr),
                              static_cast<unsigned long>(finalHr),
                              static_cast<unsigned long>(callbackStatus),
                              task._itemCompletedCallbackCount.load(std::memory_order_acquire),
                              builderCount,
                              item ? static_cast<unsigned>(item->completion) : UINT_MAX,
                              item ? static_cast<unsigned long>(item->status) : 0ul,
                              receipt != nullptr && receipt->has_value() ? static_cast<int>(receipt->value().outcomeKnown) : -1,
                              receipt != nullptr && receipt->has_value() ? static_cast<int>(receipt->value().mutationCommitted) : -1,
                              receipt != nullptr && receipt->has_value() ? static_cast<int>(receipt->value().originalStillPresent) : -1,
                              task._perf.conflictPromptCount,
                              sourceAbsent,
                              destinationAbsent));
    }
    else
    {
        AppendLog(L"Local missing-source producer and finalizer: provider callback reported a known no-commit; final result is a Failed item with the "
                  L"real error, no prompt and no destination mutation.");
    }
    fileOperations.DebugRemoveDiagnosticsForTask(task._taskId);
    std::error_code cleanupError;
    const bool removed    = SelfTest::RemoveAll(root);
    const bool rootAbsent = ! std::filesystem::exists(root, cleanupError) && ! cleanupError;
    cleanupVerified       = rootAbsent;
    if (! removed || ! rootAbsent)
    {
        AppendLog(std::format(
            L"Native-Move missing-source fixture checked cleanup failed (removed={}, rootAbsent={}, error={}).", removed, rootAbsent, cleanupError.value()));
        return false;
    }
    return matches;
}

// A queued Move retains its admission while another actor removes the selected source. Exercise
// the production per-item guard, finalizer and completed projection without a worker or UI input.
[[nodiscard]] bool RunFileOperationQueuedMissingSourceTaskSelfTest(FolderWindow::FileOperationState& fileOperations,
                                                                   const wil::com_ptr<IFileSystem>& localFileSystem,
                                                                   const std::filesystem::path& testRoot) noexcept
{
    using State          = FolderWindow::FileOperationState;
    const uint64_t token = (static_cast<uint64_t>(GetCurrentProcessId()) << 32u) ^ GetTickCount64();
    const auto root      = testRoot / std::format(L"queued-move-missing-source-{:016X}", token);
    if (! localFileSystem || ! RecreateEmptyDirectory(root))
        return false;
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    size_t ordinal     = 0u;
    bool passed        = true;
    for (const auto strategy : {FileOperations::OperationStrategy::Native, FileOperations::OperationStrategy::Managed})
    {
        for (const bool directory : {false, true})
        {
            for (unsigned int variant = 0u; variant < 3u; ++variant)
            {
                const bool missingSource          = variant != 2u;
                const bool unsupportedDestination = variant != 0u;
                if (! missingSource && strategy == FileOperations::OperationStrategy::Native)
                    continue;
                const auto caseRoot    = root / std::format(L"case-{}", ++ordinal);
                const auto source      = caseRoot / (directory ? L"source-folder" : L"source.bin");
                const auto destination = caseRoot / L"destination";
                if (! RecreateEmptyDirectory(caseRoot) || ! RecreateEmptyDirectory(destination) ||
                    ! (directory ? RecreateEmptyDirectory(source) : WriteTestFile(source, 4096u)))
                    return false;
                RedSalamander::TestSupport::ScopedEnvironmentVariable identitySupport(
                    kSelfTestEnvBindIdentityUnsupportedPath, unsupportedDestination ? std::optional<std::wstring_view>(destination.native()) : std::nullopt);
                FileOperations::QualifiedEndpoint endpoint{
                    .pluginId               = L"builtin/file-system",
                    .instanceId             = L"host/default",
                    .profileId              = L"local",
                    .rootId                 = L"local-test-volume",
                    .pathIdentity           = FileSystemPathIdentity{.pathTextStableIdentity = true,
                                                                     .componentComparison    = FileSystemPathComponentComparison::OrdinalIgnoreCase,
                                                                     .preferredSeparator     = L'\\',
                                                                     .acceptedSeparators     = L"\\/",
                                                                     .casePreserving         = true,
                                                                     .caseOnlyRename         = FileSystemPathCaseOnlyRename::Supported},
                    .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
                };
                FileOperations::TransferPlan transfer{};
                transfer.intent                         = FileOperations::TransferIntent::Move;
                transfer.strategy                       = strategy;
                transfer.sourceEndpoint                 = endpoint;
                transfer.destinationEndpoint            = endpoint;
                transfer.selectedItems                  = {{.providerPath = source.native()}};
                transfer.destination.providerFolderPath = destination.native();
                const auto plans                        = std::make_shared<const FileOperations::FileOperationPlanGroup>(
                    FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}});
                auto storage = std::make_unique<State::Task>(fileOperations);
                auto& task   = *storage;
                task._taskId = 0xD000000000000000ull | ((token + ordinal) & 0x0FFFFFFFFFFFFFFFull);
                task._admissionKey.store(task._taskId, std::memory_order_release);
                task._operation                = FILESYSTEM_MOVE;
                task._executionMode            = State::ExecutionMode::PerItem;
                task._fileSystem               = localFileSystem;
                task._sourcePaths              = {source};
                task._sourcePluginId           = endpoint.pluginId;
                task._sourcePluginShortId      = L"file";
                task._sourceInstanceContext    = endpoint.instanceId;
                task._sourcePathAttributesHint = {GetFileAttributesW(source.c_str())};
                task._perItemMaxConcurrency    = 1u;
                task.InitializeAdmissionInputs(plans,
                                               {.fileSystem      = localFileSystem,
                                                .folder          = destination,
                                                .pluginId        = endpoint.pluginId,
                                                .pluginShortId   = L"file",
                                                .instanceContext = endpoint.instanceId});
                task.InitializeSourceItemResultBuilders();
                const auto diagnosticsCleanup = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });
                if (FAILED(task.PrepareMutationInterlockScopes()) || (missingSource && ! SelfTest::RemoveAll(source)))
                    return false;
                // Exercise the retained-source guard used by a replacement admission as well:
                // positive absence must not be rewritten as a generic changed-source failure.
                if (missingSource && unsupportedDestination)
                    task._admissionKey.store(task._taskId + ordinal, std::memory_order_release);
                const HRESULT executeHr = task.ExecuteOperation();
                const HRESULT finalHr   = task.FinalizeTypedItemResults(executeHr);
                task._resultHr.store(finalHr, std::memory_order_release);
                const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);
                std::error_code sourceError;
                std::error_code destinationError;
                const bool sourceAbsent      = ! std::filesystem::exists(source, sourceError) && ! sourceError;
                const bool destinationAbsent = ! std::filesystem::exists(destination / source.filename(), destinationError) && ! destinationError;
                const bool missingError      = executeHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || executeHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
                const auto& item             = task._sourceItemResultBuilders.front().terminal;
                const auto missingText       = LoadStringResource(nullptr, IDS_FILEOPS_DIAG_SOURCE_MISSING);
                const bool matches =
                    (missingSource ? missingError : executeHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)) && finalHr == executeHr &&
                    summary.publishedItemCount == 0u && summary.removedSourceCount == 0u && summary.retainedSourceCount == (missingSource ? 0u : 1u) &&
                    summary.unknownSourceCount == 0u && summary.missingSourceCount == (missingSource ? 1u : 0u) &&
                    summary.retainedSourcePaths.empty() == missingSource && summary.indeterminateItemCount == 0u && item.has_value() &&
                    item->sourceDisposition == (missingSource ? FileOperations::SourceDisposition::Missing : FileOperations::SourceDisposition::Retained) &&
                    item->completion == FileOperations::ItemCompletion::Failed && (summary.lastDiagnosticMessage == missingText) == missingSource &&
                    sourceAbsent == missingSource && destinationAbsent && task._perf.conflictPromptCount == 0u &&
                    task._itemCompletedCallbackCount.load(std::memory_order_acquire) == 0u;
                AppendLog(std::format(
                    L"Queued Move missing source: strategy={} folder={} missing={} identityUnsupported={} execute=0x{:08X} final=0x{:08X} kept={} unknown={} "
                    L"sourceAbsent={} destinationAbsent={} passed={}",
                    static_cast<unsigned>(strategy),
                    directory,
                    missingSource,
                    unsupportedDestination,
                    static_cast<unsigned long>(executeHr),
                    static_cast<unsigned long>(finalHr),
                    summary.retainedSourceCount,
                    summary.unknownSourceCount,
                    sourceAbsent,
                    destinationAbsent,
                    matches));
                passed = matches && passed;
            }
        }
    }
    return passed;
}

// Reproduce an existing-file Managed Move to a location without FileIdInfo. The guard must
// retain both byte streams and explain the collision in the card and Issues, including after
// an aggregate task.result and a live language change. Negative controls must not invent a collision.
[[nodiscard]] bool RunFileOperationUnverifiableDestinationCollisionSelfTest(FolderWindow::FileOperationState& fileOperations,
                                                                            const wil::com_ptr<IFileSystem>& localFileSystem,
                                                                            const std::filesystem::path& testRoot) noexcept
{
    using State = FolderWindow::FileOperationState;
    wil::com_ptr<IFileSystemIO> io;
    const auto root = testRoot / L"unverifiable-destination-collision";
    if (! localFileSystem || ! localFileSystem.try_query_to(io.put()) || ! RecreateEmptyDirectory(root))
        return false;
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui->language.empty() && g_settings.ui->language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui->language};
    const auto restoreLanguage = wil::scope_exit([&]() noexcept { static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage)); });
    bool passed                = true;
    for (unsigned int variant = 0u; variant < 5u; ++variant)
    {
        const bool destinationExists = variant != 3u;
        const bool sourceUnsupported = variant == 4u;
        const bool expectCollision   = destinationExists && ! sourceUnsupported;
        const auto caseRoot          = root / std::format(L"case-{}", variant);
        const auto source            = caseRoot / L"source" / L"file.bin";
        const auto destinationFolder = caseRoot / L"destination";
        const auto destination       = destinationFolder / source.filename();
        const std::string sourceBytes(97u, 's');
        const std::string destinationBytes(97u, variant == 1u ? 's' : 'd');
        if (! RecreateEmptyDirectory(source.parent_path()) || ! RecreateEmptyDirectory(destinationFolder) || ! WriteFileTextFsIo(io, source, sourceBytes) ||
            (destinationExists && ! WriteFileTextFsIo(io, destination, destinationBytes)))
            return false;
        const auto injectionPath = sourceUnsupported ? source : variant == 0u || variant == 1u ? destination : destinationFolder;
        RedSalamander::TestSupport::ScopedEnvironmentVariable identitySupport(kSelfTestEnvBindIdentityUnsupportedPath, injectionPath.native());
        FileOperations::QualifiedEndpoint endpoint{
            .pluginId               = L"builtin/file-system",
            .instanceId             = L"host/default",
            .profileId              = L"local",
            .rootId                 = L"local-test-volume",
            .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };
        FileOperations::TransferPlan transfer{};
        transfer.intent                         = FileOperations::TransferIntent::Move;
        transfer.strategy                       = FileOperations::OperationStrategy::Managed;
        transfer.sourceEndpoint                 = endpoint;
        transfer.destinationEndpoint            = endpoint;
        transfer.selectedItems                  = {{.providerPath = source.native()}};
        transfer.destination.providerFolderPath = destinationFolder.native();
        const auto plans                        = std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}});
        State::Task task(fileOperations);
        task._operation                = FILESYSTEM_MOVE;
        task._executionMode            = State::ExecutionMode::PerItem;
        task._fileSystem               = localFileSystem;
        task._sourcePaths              = {source};
        task._sourcePluginId           = endpoint.pluginId;
        task._sourcePluginShortId      = L"file";
        task._sourceInstanceContext    = endpoint.instanceId;
        task._sourcePathAttributesHint = {GetFileAttributesW(source.c_str())};
        task._perItemMaxConcurrency    = 1u;
        task.InitializeAdmissionInputs(plans,
                                       {.fileSystem      = localFileSystem,
                                        .folder          = destinationFolder,
                                        .pluginId        = endpoint.pluginId,
                                        .pluginShortId   = L"file",
                                        .instanceContext = endpoint.instanceId});
        task.InitializeSourceItemResultBuilders();
        const auto diagnosticsCleanup = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });
        if (FAILED(task.PrepareMutationInterlockScopes()))
            return false;
        const auto started      = GetTickCount64();
        const HRESULT executeHr = task.ExecuteOperation();
        const HRESULT finalHr   = task.FinalizeTypedItemResults(executeHr);
        task._resultHr.store(finalHr, std::memory_order_release);
        task.LogDiagnostic(State::DiagnosticSeverity::Error, finalHr, L"task.result", L"Aggregate failure.");
        unsigned long warnings = 0u;
        unsigned long errors   = 0u;
        std::wstring runningNote;
        fileOperations.CollectTaskDiagnosticSnapshot(task._taskId, warnings, errors, runningNote);
        // I63 B5: the trailing task.result restates the item's own Error and is not counted again.
        const bool restatedResultCountedOnce = expectCollision ? errors == 1u : errors <= 1u;
        const auto summary                   = fileOperations.DebugRecordCompletedTaskForSelfTest(task);
        Debug::Perf::EmitValue(L"FileOps.SelfTest.UnverifiableDestinationCollisionMs", GetTickCount64() - started, finalHr);
        std::string actualSource;
        std::string actualDestination;
        std::error_code error;
        const bool destinationPreserved = destinationExists ? ReadFileTextFsIo(io, destination, actualDestination) && actualDestination == destinationBytes
                                                            : ! std::filesystem::exists(destination, error) && ! error;
        const bool matches =
            executeHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) && finalHr == executeHr && ReadFileTextFsIo(io, source, actualSource) &&
            actualSource == sourceBytes && destinationPreserved && summary.publishedItemCount == 0u && summary.removedSourceCount == 0u &&
            summary.retainedSourceCount == 1u && summary.unknownSourceCount == 0u && task._perf.conflictPromptCount == 0u &&
            task._itemCompletedCallbackCount.load(std::memory_order_acquire) == 0u && restatedResultCountedOnce &&
            (summary.lastDiagnosticMessage == LoadStringResource(nullptr, IDS_FILEOPS_DIAG_DESTINATION_EXISTS_UNVERIFIABLE)) == expectCollision &&
            (! expectCollision || (runningNote == summary.lastDiagnosticMessage && summary.lastProblem.has_value() &&
                                   summary.lastProblem.value().userText.messageId == IDS_FILEOPS_DIAG_DESTINATION_EXISTS_UNVERIFIABLE));
        AppendLog(std::format(L"Unverifiable destination collision: case={} exists={} sameBytes={} sourceUnsupported={} hr=0x{:08X} preserved={} "
                              L"errors={} passed={} note='{}'",
                              variant,
                              destinationExists,
                              sourceBytes == destinationBytes,
                              sourceUnsupported,
                              static_cast<unsigned long>(finalHr),
                              destinationPreserved && actualSource == sourceBytes,
                              errors,
                              matches,
                              summary.lastDiagnosticMessage));
        passed = matches && passed;
        if (expectCollision)
        {
            for (const std::wstring_view culture : {L"fr-FR", L"en-US"})
            {
                if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, std::wstring(culture)})))
                    return false;
                const std::wstring expected = LoadStringResource(nullptr, IDS_FILEOPS_DIAG_DESTINATION_EXISTS_UNVERIFIABLE);
                const bool localized        = ! expected.empty() && State::FormatCompletedLastNote(summary) == expected && summary.lastProblem.has_value() &&
                                              State::FormatIssueProblemText(summary.lastProblem.value()) == expected;
                AppendLog(std::format(
                    L"Unverifiable destination collision: case={} culture={} completed/Issues={} expected='{}'", variant, culture, localized, expected));
                passed = localized && passed;
            }
        }
    }
    return passed;
}

// I63 B3: the engine's comparison of the incoming item with the existing one, and the words the
// card and the decision mirror share for it.
namespace ConflictComparisonContract
{
using FileOperations::CompareConflictItems;
using FileOperations::ConflictDateRelation;
using FileOperations::ConflictSizeRelation;
inline constexpr __int64 kSecond = 10'000'000ll;
inline constexpr __int64 kHour   = 3'600ll * kSecond;
inline constexpr __int64 kBase   = 134'000'000'000'000'000ll;
static_assert(CompareConflictItems(true, 5u, kBase, true, 5u, kBase).date == ConflictDateRelation::Same);
static_assert(CompareConflictItems(true, 5u, kBase, true, 5u, kBase).size == ConflictSizeRelation::Same);
// Two seconds is the coarsest stored timestamp: inside it the dates are the same, in either direction.
static_assert(CompareConflictItems(true, 5u, kBase + 2 * kSecond, true, 5u, kBase).date == ConflictDateRelation::Same);
static_assert(CompareConflictItems(true, 5u, kBase, true, 5u, kBase + 2 * kSecond).date == ConflictDateRelation::Same);
static_assert(CompareConflictItems(true, 5u, kBase + 3 * kSecond, true, 5u, kBase).date == ConflictDateRelation::Newer);
static_assert(CompareConflictItems(true, 5u, kBase, true, 5u, kBase + 3 * kSecond).date == ConflictDateRelation::Older);
// One hour apart is its own class only with equal sizes; otherwise it is an ordinary newer or older.
static_assert(CompareConflictItems(true, 5u, kBase + kHour + kSecond, true, 5u, kBase).date == ConflictDateRelation::OneHourApart);
static_assert(CompareConflictItems(true, 5u, kBase, true, 5u, kBase + kHour - kSecond).date == ConflictDateRelation::OneHourApart);
static_assert(CompareConflictItems(true, 6u, kBase + kHour, true, 5u, kBase).date == ConflictDateRelation::Newer);
static_assert(CompareConflictItems(false, 0u, kBase + kHour, true, 5u, kBase).date == ConflictDateRelation::Newer);
static_assert(CompareConflictItems(true, 5u, kBase + 2 * kHour, true, 5u, kBase).date == ConflictDateRelation::Newer);
// Unknown facts stay unknown; nothing is guessed.
static_assert(CompareConflictItems(true, 5u, 0, true, 5u, kBase).date == ConflictDateRelation::Unknown);
static_assert(CompareConflictItems(false, 0u, kBase, true, 5u, kBase).size == ConflictSizeRelation::Unknown);
static_assert(CompareConflictItems(true, 9u, kBase, true, 5u, kBase).size == ConflictSizeRelation::Larger);
static_assert(CompareConflictItems(true, 3u, kBase, true, 5u, kBase).size == ConflictSizeRelation::Smaller);
} // namespace ConflictComparisonContract

[[nodiscard]] bool RunFileOperationConflictComparisonTextSelfTest() noexcept
{
    using FileOperations::ConflictDateRelation;
    using FileOperations::ConflictSizeRelation;
    const std::wstring newer   = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_CMP_NEWER);
    const std::wstring older   = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_CMP_OLDER);
    const std::wstring larger  = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_CMP_LARGER);
    const std::wstring smaller = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_CMP_SMALLER);
    const std::wstring sameDay = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_CMP_SAME_DATE);
    const std::wstring sameLen = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_CMP_SAME_SIZE);
    const std::wstring sameAll = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_CMP_SAME_SIZE_AND_DATE);
    const std::wstring oneHour = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_CMP_ONE_HOUR_APART);
    if (newer.empty() || older.empty() || larger.empty() || smaller.empty() || sameDay.empty() || sameLen.empty() || sameAll.empty() || oneHour.empty())
        return false;
    const auto has = [](const std::wstring& text, const std::wstring& word) noexcept { return text.find(word) != std::wstring::npos; };

    FileOperations::ConflictPromptTextFacts facts{};
    facts.operation       = FILESYSTEM_MOVE;
    facts.bucket          = static_cast<uint8_t>(FileOperations::ConflictClass::RegularFileExists);
    facts.sourceItem      = {.available = true, .sizeKnown = true, .sizeBytes = 4096u, .lastWriteTime = ConflictComparisonContract::kBase};
    facts.destinationItem = {.available = true, .sizeKnown = true, .sizeBytes = 1024u, .lastWriteTime = ConflictComparisonContract::kBase};
    const auto relate     = [&](const ConflictDateRelation date, const ConflictSizeRelation size) noexcept
    {
        facts.dateRelation = static_cast<uint8_t>(date);
        facts.sizeRelation = static_cast<uint8_t>(size);
    };
    const std::wstring sizeText = FormatBytesCompact(4096u);

    // Each role names its own relation; the existing item's words are the inverse of the incoming item's.
    relate(ConflictDateRelation::Newer, ConflictSizeRelation::Larger);
    const std::wstring incoming = FileOperations::BuildConflictPromptItemText(facts, false);
    const std::wstring existing = FileOperations::BuildConflictPromptItemText(facts, true);
    const bool roles = has(incoming, sizeText) && has(incoming, newer) && has(incoming, larger) && ! has(incoming, older) && ! has(incoming, smaller) &&
                       has(existing, older) && has(existing, smaller) && ! has(existing, newer) && ! has(existing, larger) &&
                       FileOperations::BuildConflictPromptFactText(facts).empty();

    // Same size and date is said once for both items, and never as "identical".
    relate(ConflictDateRelation::Same, ConflictSizeRelation::Same);
    const std::wstring sameIncoming = FileOperations::BuildConflictPromptItemText(facts, false);
    const bool same = FileOperations::BuildConflictPromptFactText(facts) == sameAll && has(sameIncoming, sizeText) && ! has(sameIncoming, sameDay) &&
                      ! has(sameIncoming, sameLen) && ! has(FileOperations::BuildConflictPromptItemText(facts, true), sameLen);

    relate(ConflictDateRelation::OneHourApart, ConflictSizeRelation::Same);
    const bool hour =
        FileOperations::BuildConflictPromptFactText(facts) == oneHour && ! has(FileOperations::BuildConflictPromptItemText(facts, false), sameLen);

    // A partial comparison says only what is known.
    relate(ConflictDateRelation::Same, ConflictSizeRelation::Smaller);
    const std::wstring mixed = FileOperations::BuildConflictPromptItemText(facts, false);
    relate(ConflictDateRelation::Older, ConflictSizeRelation::Unknown);
    const std::wstring partial = FileOperations::BuildConflictPromptItemText(facts, false);
    const bool partials        = has(mixed, sameDay) && has(mixed, smaller) && has(partial, older) && ! has(partial, larger) && ! has(partial, smaller) &&
                                 ! has(partial, sameLen) && FileOperations::BuildConflictPromptFactText(facts).empty();

    // A consent's own facts keep their line, and an unknown item has no line at all.
    relate(ConflictDateRelation::Same, ConflictSizeRelation::Same);
    facts.factItemCountKnown = true;
    facts.factItemCount      = 3u;
    const bool consentKept   = FileOperations::BuildConflictPromptFactText(facts) != sameAll;
    facts.sourceItem         = {};
    const bool unknownItem   = FileOperations::BuildConflictPromptItemText(facts, false).empty();

    const bool passed = roles && same && hour && partials && consentKept && unknownItem;
    AppendLog(std::format(L"Conflict comparison text: roles={} same={} hour={} partials={} consentKept={} unknownItem={} passed={}",
                          roles,
                          same,
                          hour,
                          partials,
                          consentKept,
                          unknownItem,
                          passed));
    return passed;
}

// I63 slice B: a Managed Move onto an existing file at a location that opens items but cannot
// identify them is a name collision the user decides, not a refusal. The decision is raised before
// any byte is read and shows both items' facts; Skip keeps both byte streams, Keep both publishes
// beside the existing item and completes the Move, Cancel changes nothing. Replace (I63 B2) publishes
// over the occupant the user was shown through one replacing rename. When that occupant changed after
// the decision was shown nothing is replaced and the decision is raised again on the current occupant;
// when another program keeps the occupant in use the rename is refused and both files stay.
[[nodiscard]] bool RunFileOperationIdentityLessCollisionDecisionSelfTest(FolderWindow::FileOperationState& fileOperations,
                                                                         const wil::com_ptr<IFileSystem>& localFileSystem,
                                                                         const std::filesystem::path& testRoot) noexcept
{
    using State  = FolderWindow::FileOperationState;
    using Action = State::Task::ConflictAction;
    wil::com_ptr<IFileSystemIO> io;
    const auto root = testRoot / L"identity-less-collision-decision";
    if (! localFileSystem || ! localFileSystem.try_query_to(io.put()) || ! RecreateEmptyDirectory(root))
        return false;
    const auto cleanup     = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    bool passed            = true;
    unsigned int caseIndex = 0u;
    struct Variant final
    {
        bool wholeFolder      = false; // the whole destination folder cannot identify items (a share without file IDs), or only the existing item
        Action decision       = Action::Skip;
        bool occupantReadOnly = false;
        bool occupantChanges  = false; // the occupant is rewritten after the decision is shown and before it is answered
        bool occupantInUse    = false; // after the decision is shown, another program opens the occupant and does not share its deletion
    };
    static constexpr std::array kVariants{
        Variant{false, Action::Skip},
        Variant{false, Action::KeepBoth},
        Variant{false, Action::Cancel},
        Variant{false, Action::Overwrite},
        Variant{false, Action::ReplaceReadOnly, true},
        Variant{false, Action::Overwrite, false, true},
        Variant{false, Action::Overwrite, false, false, true},
        Variant{true, Action::Skip},
        Variant{true, Action::KeepBoth},
        Variant{true, Action::Cancel},
        Variant{true, Action::Overwrite},
        Variant{true, Action::ReplaceReadOnly, true},
        Variant{true, Action::Overwrite, false, true},
        Variant{true, Action::Overwrite, false, false, true},
    };
    for (const Variant& variant : kVariants)
    {
        const bool wholeFolder       = variant.wholeFolder;
        const Action decision        = variant.decision;
        const auto caseRoot          = root / std::format(L"case-{}", caseIndex++);
        const auto source            = caseRoot / L"source" / L"harbor.bin";
        const auto destinationFolder = caseRoot / L"destination";
        const auto destination       = destinationFolder / source.filename();
        const std::string sourceBytes(97u, 's');
        const std::string destinationBytes(131u, 'd');
        const std::string changedBytes(177u, 'x');
        if (! RecreateEmptyDirectory(source.parent_path()) || ! RecreateEmptyDirectory(destinationFolder) || ! WriteFileTextFsIo(io, source, sourceBytes) ||
            ! WriteFileTextFsIo(io, destination, destinationBytes))
            return false;
        if (variant.occupantReadOnly && SetFileAttributesW(destination.c_str(), FILE_ATTRIBUTE_READONLY) == FALSE)
            return false;
        // A refused replacement must not leave a read-only fixture behind for the root cleanup.
        const auto writableCleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SetFileAttributesW(destination.c_str(), FILE_ATTRIBUTE_NORMAL)); });
        // Opened by the responder once the decision is shown: an occupant already in use when the
        // question is built is not offered for replacement at all, as on a location with identity.
        wil::unique_hfile inUse;
        const auto injectionPath = wholeFolder ? destinationFolder : destination;
        RedSalamander::TestSupport::ScopedEnvironmentVariable identitySupport(kSelfTestEnvBindIdentityUnsupportedPath, injectionPath.native());
        FileOperations::QualifiedEndpoint endpoint{
            .pluginId               = L"builtin/file-system",
            .instanceId             = L"host/default",
            .profileId              = L"local",
            .rootId                 = L"local-test-volume",
            .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };
        // The Local destination proves what it published; the decision route requires that proof.
        endpoint.verificationHostReadback = true;
        FileOperations::TransferPlan transfer{};
        transfer.intent                         = FileOperations::TransferIntent::Move;
        transfer.strategy                       = FileOperations::OperationStrategy::Managed;
        transfer.sourceEndpoint                 = endpoint;
        transfer.destinationEndpoint            = endpoint;
        transfer.selectedItems                  = {{.providerPath = source.native()}};
        transfer.destination.providerFolderPath = destinationFolder.native();
        const auto plans                        = std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}});
        State::Task task(fileOperations);
        task._operation                = FILESYSTEM_MOVE;
        task._executionMode            = State::ExecutionMode::PerItem;
        task._fileSystem               = localFileSystem;
        task._sourcePaths              = {source};
        task._sourcePluginId           = endpoint.pluginId;
        task._sourcePluginShortId      = L"file";
        task._sourceInstanceContext    = endpoint.instanceId;
        task._sourcePathAttributesHint = {GetFileAttributesW(source.c_str())};
        task._perItemMaxConcurrency    = 1u;
        task.InitializeAdmissionInputs(plans,
                                       {.fileSystem      = localFileSystem,
                                        .folder          = destinationFolder,
                                        .pluginId        = endpoint.pluginId,
                                        .pluginShortId   = L"file",
                                        .instanceContext = endpoint.instanceId});
        task.InitializeSourceItemResultBuilders();
        const auto diagnosticsCleanup = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });
        if (FAILED(task.PrepareMutationInterlockScopes()))
            return false;

        // The executor blocks on its decision: answer it from a second thread. A prompt that does not
        // offer the planned decision is answered with Cancel so the case cannot hang. A decision raised
        // again (changed occupant) or a refusal raised to the user (occupant in use) is answered Skip.
        std::mutex seenMutex;
        std::optional<State::Task::ConflictPromptState> seenPrompt;
        std::atomic<unsigned int> answered{0u};
        std::atomic<bool> decisionOffered{true};
        std::atomic<bool> followUpAnswerOffered{true};
        std::atomic<bool> occupantChanged{false};
        std::atomic<unsigned int> followUpPrompts{0u};
        std::atomic<uint64_t> followUpDestinationSize{0u};
        HRESULT executeHr = E_PENDING;
        {
            const std::jthread responder([&](const std::stop_token stop) noexcept
            {
                uint64_t answeredGeneration = 0u;
                while (! stop.stop_requested())
                {
                    if (const auto prompt = TryGetConflictPromptCopy(&task); prompt.has_value() && prompt->generation != answeredGeneration)
                    {
                        const bool followUp   = answered.load(std::memory_order_acquire) != 0u;
                        const Action wanted   = followUp && (variant.occupantChanges || variant.occupantInUse) ? Action::Skip : decision;
                        const auto offeredEnd = prompt->actions.begin() + static_cast<std::ptrdiff_t>(prompt->actionCount);
                        const bool offered    = std::find(prompt->actions.begin(), offeredEnd, wanted) != offeredEnd;
                        if (followUp)
                        {
                            followUpPrompts.fetch_add(1u, std::memory_order_acq_rel);
                            followUpDestinationSize.store(prompt->destinationMetadata.sizeKnown ? prompt->destinationMetadata.sizeBytes : 0u,
                                                          std::memory_order_release);
                        }
                        {
                            const std::scoped_lock lock(seenMutex);
                            if (! seenPrompt.has_value())
                                seenPrompt = prompt;
                        }
                        if (! offered)
                            (followUp ? followUpAnswerOffered : decisionOffered).store(false, std::memory_order_release);
                        if (variant.occupantInUse && ! followUp && ! inUse)
                        {
                            inUse.reset(CreateFileW(
                                destination.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
                        }
                        if (variant.occupantChanges && ! followUp && ! occupantChanged.load(std::memory_order_acquire))
                        {
                            // Another writer replaces the occupant's content while the user reads the question.
                            const wil::unique_hfile changed(
                                CreateFileW(destination.c_str(), GENERIC_WRITE, 0u, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
                            DWORD written = 0u;
                            occupantChanged.store(
                                changed && WriteFile(changed.get(), changedBytes.data(), static_cast<DWORD>(changedBytes.size()), &written, nullptr) != FALSE &&
                                    written == changedBytes.size(),
                                std::memory_order_release);
                        }
                        if (task.SubmitConflictDecision(prompt->generation, offered ? wanted : Action::Cancel, false))
                        {
                            answeredGeneration = prompt->generation;
                            answered.fetch_add(1u, std::memory_order_acq_rel);
                        }
                    }
                    Sleep(2);
                }
            });
            executeHr = task.ExecuteOperation();
        }
        const bool occupantHeld = inUse.is_valid();
        inUse.reset();
        const HRESULT finalHr = task.FinalizeTypedItemResults(executeHr);
        task._resultHr.store(finalHr, std::memory_order_release);
        const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);

        std::string actualSource;
        std::string actualDestination;
        std::error_code error;
        const bool sourceKept             = ReadFileTextFsIo(io, source, actualSource) && actualSource == sourceBytes;
        const bool sourceGone             = ! std::filesystem::exists(source, error) && ! error;
        const bool destinationRead        = ReadFileTextFsIo(io, destination, actualDestination);
        const bool destinationPreserved   = destinationRead && actualDestination == destinationBytes;
        const bool destinationReplaced    = destinationRead && actualDestination == sourceBytes;
        const bool destinationChanged     = destinationRead && actualDestination == changedBytes;
        const DWORD destinationAttributes = GetFileAttributesW(destination.c_str());
        const bool destinationWritable    = destinationAttributes != INVALID_FILE_ATTRIBUTES && (destinationAttributes & FILE_ATTRIBUTE_READONLY) == 0u;
        // Keep both publishes the source's bytes under a new name beside the existing item.
        size_t destinationEntries = 0u;
        size_t keptBothEntries    = 0u;
        size_t leftoverEntries    = 0u;
        for (std::filesystem::directory_iterator it(destinationFolder, error), end; ! error && it != end; it.increment(error))
        {
            ++destinationEntries;
            const std::wstring leaf = it->path().filename().native();
            if (FileOperationArtifacts::HasPossibleArtifactName(leaf))
            {
                ++leftoverEntries;
                continue;
            }
            std::string bytes;
            if (leaf != destination.filename().native() && ReadFileTextFsIo(io, it->path(), bytes) && bytes == sourceBytes)
                ++keptBothEntries;
        }

        const auto shown = [&]() noexcept
        {
            const std::scoped_lock lock(seenMutex);
            return seenPrompt;
        }();
        const auto offers = [&](const Action action) noexcept
        {
            if (! shown.has_value())
                return false;
            const auto offeredEnd = shown->actions.begin() + static_cast<std::ptrdiff_t>(shown->actionCount);
            return std::find(shown->actions.begin(), offeredEnd, action) != offeredEnd;
        };
        // Both items' facts are on the decision. Replace is offered because the Local provider can
        // publish over the occupant the user is shown; a read-only occupant needs its own grant, and a
        // link is never replaced this way.
        const bool promptFacts =
            shown.has_value() && shown->sourceMetadata.available && shown->sourceMetadata.sizeKnown && shown->sourceMetadata.sizeBytes == sourceBytes.size() &&
            shown->sourceMetadata.lastWriteTime != 0 && shown->destinationMetadata.available && shown->destinationMetadata.sizeKnown &&
            shown->destinationMetadata.sizeBytes == destinationBytes.size() && shown->destinationMetadata.lastWriteTime != 0 &&
            shown->comparison.size == FileOperations::ConflictSizeRelation::Smaller && shown->comparison.date != FileOperations::ConflictDateRelation::Unknown;
        // Delete source (I63 B4) is offered here as well: the location cannot say by identity that the two
        // files differ, so the same-file probe answers.
        const bool discardOffered = offers(Action::DiscardSource);
        const bool promptActions =
            discardOffered &&
            (variant.occupantReadOnly ? offers(Action::ReplaceReadOnly) && offers(Action::Skip) && offers(Action::Cancel) && ! offers(Action::ReplaceLink)
                                      : offers(Action::KeepBoth) && offers(Action::Skip) && offers(Action::Cancel) && offers(Action::Overwrite) &&
                                            ! offers(Action::ReplaceReadOnly) && ! offers(Action::ReplaceLink));
        // One decision per collision. A changed occupant raises exactly one more, on the occupant that
        // is there now; a refused rename raises the refusal once, with Skip offered.
        const unsigned int answeredCount = answered.load(std::memory_order_acquire);
        const unsigned int followUpCount = followUpPrompts.load(std::memory_order_acquire);
        const bool onePrompt =
            decisionOffered.load(std::memory_order_acquire) &&
            (variant.occupantChanges ? answeredCount == 2u && task._perf.conflictPromptCount == 2u && followUpAnswerOffered.load(std::memory_order_acquire) &&
                                           followUpDestinationSize.load(std::memory_order_acquire) == changedBytes.size()
             : variant.occupantInUse ? answeredCount == 2u && occupantHeld && followUpAnswerOffered.load(std::memory_order_acquire)
                                     : answeredCount == 1u && task._perf.conflictPromptCount == 1u);

        bool outcome = false;
        if (decision == Action::Skip)
        {
            outcome = sourceKept && destinationPreserved && destinationEntries == 1u && summary.publishedItemCount == 0u && summary.removedSourceCount == 0u &&
                      summary.retainedSourceCount == 1u && summary.unknownSourceCount == 0u && executeHr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        else if (decision == Action::KeepBoth)
        {
            outcome = sourceGone && destinationPreserved && destinationEntries == 2u && keptBothEntries == 1u && summary.publishedItemCount == 1u &&
                      summary.removedSourceCount == 1u && summary.retainedSourceCount == 0u && summary.unknownSourceCount == 0u && SUCCEEDED(finalHr);
        }
        else if (decision == Action::Cancel)
        {
            outcome = sourceKept && destinationPreserved && destinationEntries == 1u && summary.publishedItemCount == 0u && summary.removedSourceCount == 0u &&
                      executeHr == HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        else if (variant.occupantChanges)
        {
            // The occupant is no longer the one the user decided on: nothing is replaced, and the
            // decision raised again (answered Skip) leaves the other writer's content and the source.
            outcome = occupantChanged.load(std::memory_order_acquire) && sourceKept && destinationChanged && destinationEntries == 1u &&
                      summary.publishedItemCount == 0u && summary.removedSourceCount == 0u && summary.retainedSourceCount == 1u &&
                      summary.unknownSourceCount == 0u;
        }
        else if (variant.occupantInUse)
        {
            // The replacing rename is refused: the occupant and the source are untouched, the stage is
            // removed, and nothing is written over the occupant instead.
            outcome = sourceKept && destinationPreserved && destinationEntries == 1u && summary.publishedItemCount == 0u && summary.removedSourceCount == 0u &&
                      summary.unknownSourceCount == 0u;
        }
        else
        {
            // Replace and Replace read-only: the occupant now holds the source's bytes and the Move completed.
            outcome = sourceGone && destinationReplaced && destinationWritable && destinationEntries == 1u && summary.publishedItemCount == 1u &&
                      summary.removedSourceCount == 1u && summary.retainedSourceCount == 0u && summary.unknownSourceCount == 0u && SUCCEEDED(finalHr);
        }
        const bool matches = onePrompt && promptFacts && promptActions && outcome && leftoverEntries == 0u;
        AppendLog(std::format(L"Identity-less collision decision: wholeFolder={} decision={} readOnly={} changes={} inUse={} execute=0x{:08X} final=0x{:08X} "
                              L"prompts={} answered={} followUps={} followUpOffered={} offered={} facts={} actions={} discard={} sourceKept={} sourceGone={} "
                              L"destinationPreserved={} "
                              L"replaced={} changed={} "
                              L"entries={} keptBoth={} leftovers={} published={} removed={} retained={} unknown={} passed={}",
                              wholeFolder,
                              static_cast<unsigned int>(decision),
                              variant.occupantReadOnly,
                              variant.occupantChanges,
                              variant.occupantInUse,
                              static_cast<unsigned long>(executeHr),
                              static_cast<unsigned long>(finalHr),
                              static_cast<uint64_t>(task._perf.conflictPromptCount),
                              answeredCount,
                              followUpCount,
                              followUpAnswerOffered.load(std::memory_order_acquire),
                              decisionOffered.load(std::memory_order_acquire),
                              promptFacts,
                              promptActions,
                              discardOffered,
                              sourceKept,
                              sourceGone,
                              destinationPreserved,
                              destinationReplaced,
                              destinationChanged,
                              destinationEntries,
                              keptBothEntries,
                              leftoverEntries,
                              summary.publishedItemCount,
                              summary.removedSourceCount,
                              summary.retainedSourceCount,
                              summary.unknownSourceCount,
                              matches));
        passed = matches && passed;
    }
    return passed;
}

// I63 B4: Delete source. A Managed Move of one file onto an existing file at an ordinary destination (a location
// that identifies its items). The user answers Delete source from the prompt's More list: the destination keeps its
// bytes, nothing is copied or staged, the source goes to the Recycle Bin, and the item is reported as already there
// with its source removed, not as skipped. Apply to all is remembered only when the two files have the same size and
// date. The Local provider's Recycle route is the real one; the runner purges the Recycle Bin items it journals for
// originals inside the marked test sandbox.
[[nodiscard]] bool RunFileOperationDiscardSourceSelfTest(FolderWindow::FileOperationState& fileOperations,
                                                         const wil::com_ptr<IFileSystem>& localFileSystem,
                                                         const std::filesystem::path& testRoot) noexcept
{
    using State  = FolderWindow::FileOperationState;
    using Action = State::Task::ConflictAction;
    wil::com_ptr<IFileSystemIO> io;
    const auto root = testRoot / L"discard-source-decision";
    if (! localFileSystem || ! localFileSystem.try_query_to(io.put()) || ! RecreateEmptyDirectory(root))
        return false;
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    struct Variant final
    {
        bool sameSizeAndDate = false; // the destination has the source's size and last-write time, or another size
        bool applyToAll      = false; // the answer ticks Apply to all
    };
    static constexpr std::array kVariants{Variant{true, false}, Variant{false, false}, Variant{true, true}, Variant{false, true}};
    bool passed            = true;
    unsigned int caseIndex = 0u;
    for (const Variant& variant : kVariants)
    {
        const auto caseRoot          = root / std::format(L"case-{}", caseIndex++);
        const auto source            = caseRoot / L"source" / L"harbor.bin";
        const auto destinationFolder = caseRoot / L"destination";
        const auto destination       = destinationFolder / source.filename();
        const std::string sourceBytes(97u, 's');
        const std::string destinationBytes(variant.sameSizeAndDate ? 97u : 131u, 'd');
        std::error_code error;
        if (! RecreateEmptyDirectory(source.parent_path()) || ! RecreateEmptyDirectory(destinationFolder) || ! WriteFileTextFsIo(io, source, sourceBytes) ||
            ! WriteFileTextFsIo(io, destination, destinationBytes))
            return false;
        if (variant.sameSizeAndDate)
        {
            // The destination gets the source's last-write time, so the two files are the same size and date.
            const auto sourceTime = std::filesystem::last_write_time(source, error);
            if (error)
                return false;
            std::filesystem::last_write_time(destination, sourceTime, error);
            if (error)
                return false;
        }
        FileOperations::QualifiedEndpoint endpoint{
            .pluginId               = L"builtin/file-system",
            .instanceId             = L"host/default",
            .profileId              = L"local",
            .rootId                 = L"local-test-volume",
            .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };
        endpoint.verificationHostReadback = true;
        FileOperations::TransferPlan transfer{};
        transfer.intent                         = FileOperations::TransferIntent::Move;
        transfer.strategy                       = FileOperations::OperationStrategy::Managed;
        transfer.sourceEndpoint                 = endpoint;
        transfer.destinationEndpoint            = endpoint;
        transfer.selectedItems                  = {{.providerPath = source.native()}};
        transfer.destination.providerFolderPath = destinationFolder.native();
        const auto plans                        = std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}});
        State::Task task(fileOperations);
        task._operation                = FILESYSTEM_MOVE;
        task._executionMode            = State::ExecutionMode::PerItem;
        task._fileSystem               = localFileSystem;
        task._sourcePaths              = {source};
        task._sourcePluginId           = endpoint.pluginId;
        task._sourcePluginShortId      = L"file";
        task._sourceInstanceContext    = endpoint.instanceId;
        task._sourcePathAttributesHint = {GetFileAttributesW(source.c_str())};
        task._perItemMaxConcurrency    = 1u;
        task.InitializeAdmissionInputs(plans,
                                       {.fileSystem      = localFileSystem,
                                        .folder          = destinationFolder,
                                        .pluginId        = endpoint.pluginId,
                                        .pluginShortId   = L"file",
                                        .instanceContext = endpoint.instanceId});
        task.InitializeSourceItemResultBuilders();
        const auto diagnosticsCleanup = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });
        if (FAILED(task.PrepareMutationInterlockScopes()))
            return false;

        // The executor blocks on its decision: answer it from a second thread. A prompt that does not offer
        // Delete source is answered Cancel so the case cannot hang.
        std::mutex seenMutex;
        std::optional<State::Task::ConflictPromptState> seenPrompt;
        std::atomic<unsigned int> answered{0u};
        std::atomic<bool> decisionOffered{true};
        HRESULT executeHr = E_PENDING;
        {
            const std::jthread responder([&](const std::stop_token stop) noexcept
            {
                uint64_t answeredGeneration = 0u;
                while (! stop.stop_requested())
                {
                    if (const auto prompt = TryGetConflictPromptCopy(&task); prompt.has_value() && prompt->generation != answeredGeneration)
                    {
                        const auto offeredEnd = prompt->actions.begin() + static_cast<std::ptrdiff_t>(prompt->actionCount);
                        const bool offered    = std::find(prompt->actions.begin(), offeredEnd, Action::DiscardSource) != offeredEnd;
                        {
                            const std::scoped_lock lock(seenMutex);
                            if (! seenPrompt.has_value())
                                seenPrompt = prompt;
                        }
                        if (! offered)
                            decisionOffered.store(false, std::memory_order_release);
                        if (task.SubmitConflictDecision(prompt->generation, offered ? Action::DiscardSource : Action::Cancel, variant.applyToAll))
                        {
                            answeredGeneration = prompt->generation;
                            answered.fetch_add(1u, std::memory_order_acq_rel);
                        }
                    }
                    Sleep(2);
                }
            });
            executeHr = task.ExecuteOperation();
        }
        const HRESULT finalHr = task.FinalizeTypedItemResults(executeHr);
        task._resultHr.store(finalHr, std::memory_order_release);
        const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);

        std::string actualDestination;
        const bool sourceGone           = ! std::filesystem::exists(source, error) && ! error;
        const bool destinationPreserved = ReadFileTextFsIo(io, destination, actualDestination) && actualDestination == destinationBytes;
        size_t destinationEntries       = 0u;
        for (std::filesystem::directory_iterator it(destinationFolder, error), end; ! error && it != end; it.increment(error))
            ++destinationEntries;

        const auto shown = [&]() noexcept
        {
            const std::scoped_lock lock(seenMutex);
            return seenPrompt;
        }();
        const auto listed = [](const auto& actions, size_t count, const Action action) noexcept
        {
            const auto end = actions.begin() + static_cast<std::ptrdiff_t>(std::min(count, actions.size()));
            return std::find(actions.begin(), end, action) != end;
        };
        // Offered once, under More (never a primary button, never the default or the Escape action), and the
        // published comparison says what the two files are.
        const bool promptActions = shown.has_value() && listed(shown->actions, shown->actionCount, Action::DiscardSource) &&
                                   listed(shown->overflowActions, shown->overflowActionCount, Action::DiscardSource) &&
                                   ! listed(shown->primaryActions, shown->primaryActionCount, Action::DiscardSource) &&
                                   shown->defaultAction == Action::Cancel && shown->escapeAction == Action::Cancel;
        const bool promptFacts =
            shown.has_value() && shown->sourceMetadata.available && shown->sourceMetadata.sizeKnown && shown->sourceMetadata.sizeBytes == sourceBytes.size() &&
            shown->destinationMetadata.available && shown->destinationMetadata.sizeKnown && shown->destinationMetadata.sizeBytes == destinationBytes.size() &&
            (variant.sameSizeAndDate
                 ? shown->comparison.date == FileOperations::ConflictDateRelation::Same && shown->comparison.size == FileOperations::ConflictSizeRelation::Same
                 : shown->comparison.size == FileOperations::ConflictSizeRelation::Smaller);
        const bool onePrompt =
            decisionOffered.load(std::memory_order_acquire) && answered.load(std::memory_order_acquire) == 1u && task._perf.conflictPromptCount == 1u;
        // Apply to all is remembered only for the same size and date.
        const size_t cachedDecisions =
            static_cast<size_t>(std::ranges::count_if(task._conflictArbiter.decisionCache, [](const auto& entry) noexcept { return entry.has_value(); }));
        const bool cachedDiscard = std::ranges::any_of(task._conflictArbiter.decisionCache,
                                                       [](const auto& entry) noexcept { return entry.has_value() && entry->action == Action::DiscardSource; });
        const bool rememberExpected = variant.sameSizeAndDate && variant.applyToAll;
        const bool remembered       = rememberExpected ? cachedDecisions == 1u && cachedDiscard : cachedDecisions == 0u;
        const bool outcome          = sourceGone && destinationPreserved && destinationEntries == 1u && executeHr == S_OK && finalHr == S_OK &&
                                      summary.publishedItemCount == 0u && summary.removedSourceCount == 1u && summary.retainedSourceCount == 0u &&
                                      summary.unknownSourceCount == 0u && summary.skippedItemCount == 0u && summary.failedItemCount == 0u &&
                                      summary.alreadyThereSourceRemovedCount == 1u &&
                                      summary.resultSummary == LoadStringResource(nullptr, IDS_FILEOPS_RESULT_ALREADY_THERE_SOURCE_REMOVED) &&
                                      ! task._observedSkipAction.load(std::memory_order_acquire);
        const bool matches          = onePrompt && promptActions && promptFacts && remembered && outcome;
        AppendLog(std::format(L"Delete source decision: sameSizeAndDate={} applyToAll={} execute=0x{:08X} final=0x{:08X} prompts={} answered={} offered={} "
                              L"actions={} facts={} remembered={} cached={} sourceGone={} destinationPreserved={} entries={} published={} removed={} "
                              L"retained={} unknown={} skipped={} failed={} alreadyThere={} passed={}",
                              variant.sameSizeAndDate,
                              variant.applyToAll,
                              static_cast<unsigned long>(executeHr),
                              static_cast<unsigned long>(finalHr),
                              static_cast<uint64_t>(task._perf.conflictPromptCount),
                              answered.load(std::memory_order_acquire),
                              decisionOffered.load(std::memory_order_acquire),
                              promptActions,
                              promptFacts,
                              remembered,
                              cachedDecisions,
                              sourceGone,
                              destinationPreserved,
                              destinationEntries,
                              summary.publishedItemCount,
                              summary.removedSourceCount,
                              summary.retainedSourceCount,
                              summary.unknownSourceCount,
                              summary.skippedItemCount,
                              summary.failedItemCount,
                              summary.alreadyThereSourceRemovedCount,
                              matches));
        passed = matches && passed;
    }
    return passed;
}

// I63 G7 (D-14): the power API behind the keep-awake request, replaced by a counting fake so the case depends on nothing
// about the machine's power policy. The create function hands out a real event handle (the request owns and closes it).
struct KeepAwakeSelfTestApi final
{
    uint32_t createCalls = 0u;
    uint32_t setCalls    = 0u;
    uint32_t clearCalls  = 0u;
    bool failCreate      = false;
    bool failSet         = false;
    bool failClear       = false;
    HANDLE lastCreated   = nullptr; // borrowed observation only; PowerRequest owns and closes the event

    [[nodiscard]] static wil::unique_handle Create(void* context, wchar_t* reason) noexcept
    {
        auto* const self = static_cast<KeepAwakeSelfTestApi*>(context);
        ++self->createCalls;
        if (self->failCreate || reason == nullptr || *reason == L'\0')
        {
            SetLastError(ERROR_ACCESS_DENIED);
            return {};
        }
        self->lastCreated = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        return wil::unique_handle(self->lastCreated);
    }
    [[nodiscard]] static bool Set(void* context, HANDLE request) noexcept
    {
        auto* const self = static_cast<KeepAwakeSelfTestApi*>(context);
        ++self->setCalls;
        if (self->failSet)
            SetLastError(ERROR_NOT_SUPPORTED);
        return ! self->failSet && request != nullptr;
    }
    [[nodiscard]] static bool Clear(void* context, HANDLE request) noexcept
    {
        auto* const self = static_cast<KeepAwakeSelfTestApi*>(context);
        ++self->clearCalls;
        if (self->failClear)
            SetLastError(ERROR_GEN_FAILURE);
        return ! self->failClear && request != nullptr;
    }
    [[nodiscard]] FileOperations::PowerRequestApi Api() noexcept
    {
        FileOperations::PowerRequestApi api{};
        api.create  = &Create;
        api.set     = &Set;
        api.clear   = &Clear;
        api.context = this;
        return api;
    }
};

// I63 slice F (F1): stable stream rows in the engine, on tasks that are never admitted (no worker, no I/O), as in
// RunFolderFileWorkReducerControls. The byte-report reducer runs with fixed ticks: a stream keeps its slot while a lower
// stream ends and its entry is compacted away or removed, the lowest free slot goes to the next stream, a finished stream
// holds its slot for the 300 ms linger, and a stream that replaced an evicted one takes its slot. The lifecycle reducer:
// a new stream that reuses a finished entry takes the lowest free slot, and a stream that begins its next file keeps its
// slot. What this does not prove: the popup's drawing of the slots (its snapshot copies the field; the rows are drawn
// from it by ResolveVisibleStreamRows).
[[nodiscard]] bool RunFileOperationStreamSlotSelfTest(FolderWindow::FileOperationState& owner) noexcept
{
    using Task                 = FolderWindow::FileOperationState::Task;
    constexpr size_t kNoStream = Task::kMaxInFlightFiles + 1u;

    unsigned int step = 0u;
    const auto fail   = [&step](const wchar_t* what) noexcept
    {
        AppendLog(std::format(L"Stream slot self-test failed at step {}: {}", step, what));
        return false;
    };
    const auto slotOf = [](Task& subject, uint64_t stream) noexcept -> size_t
    {
        std::scoped_lock lock(subject._inFlightFilesMutex);
        for (size_t index = 0u; index < subject._inFlightFileCount; ++index)
            if (subject._inFlightFiles[index].progressStreamId == stream)
                return subject._inFlightFiles[index].streamSlot;
        return kNoStream;
    };
    const auto slotsDistinct = [](Task& subject) noexcept
    {
        std::scoped_lock lock(subject._inFlightFilesMutex);
        uint32_t seen = 0u;
        for (size_t index = 0u; index < subject._inFlightFileCount; ++index)
        {
            const size_t slot = subject._inFlightFiles[index].streamSlot;
            if (slot >= Task::kMaxInFlightFiles || (seen & (uint32_t{1u} << slot)) != 0u)
                return false;
            seen |= uint32_t{1u} << slot;
        }
        return true;
    };

    // Byte reports (UpdateInFlightFileProgress), every file 100 bytes, at fixed ticks.
    {
        Task streams(owner);
        const int cookieAnchor   = 0;
        const void* const cookie = &cookieAnchor;
        const auto report        = [&streams, cookie](uint64_t stream, uint64_t completed, ULONGLONG tick) noexcept
        {
            const std::wstring path = std::format(L"take-{:02}.wav", stream);
            FolderWindowFileOperationsStateInternal::UpdateInFlightFileProgress(streams, cookie, stream, path.c_str(), 100u, completed, tick, false);
        };
        ++step;
        report(1u, 10u, 1'000u);
        report(2u, 10u, 1'000u);
        report(3u, 10u, 1'000u);
        if (slotOf(streams, 1u) != 0u || slotOf(streams, 2u) != 1u || slotOf(streams, 3u) != 2u)
            return fail(L"the first streams take slots 0, 1 and 2");
        ++step;
        report(1u, 100u, 1'100u); // stream 1 completes
        report(2u, 20u, 1'500u);  // its entry expires 300 ms after completion and the array is compacted
        if (slotOf(streams, 1u) != kNoStream || slotOf(streams, 2u) != 1u || slotOf(streams, 3u) != 2u)
            return fail(L"a stream keeps its slot when a lower stream ends");
        ++step;
        report(4u, 10u, 1'500u);
        report(5u, 10u, 1'500u);
        if (slotOf(streams, 4u) != 0u || slotOf(streams, 5u) != 3u)
            return fail(L"the lowest free slot goes to the next stream");
        ++step;
        report(3u, 100u, 1'600u); // stream 3 completes
        report(6u, 10u, 1'700u);  // 100 ms later its finished row still holds slot 2
        if (slotOf(streams, 3u) != 2u || slotOf(streams, 6u) != 4u)
            return fail(L"a finished stream holds its slot for the linger");
        ++step;
        report(7u, 10u, 2'000u); // 400 ms after: stream 3 has expired and slot 2 is free
        if (slotOf(streams, 3u) != kNoStream || slotOf(streams, 7u) != 2u)
            return fail(L"a slot is free after the linger");
        ++step;
        report(2u, 30u, 2'100u);
        report(4u, 20u, 2'100u);
        FolderWindowFileOperationsStateInternal::RemoveInFlightFileBySourcePath(streams, L"take-04.wav");
        if (slotOf(streams, 4u) != kNoStream || slotOf(streams, 2u) != 1u || slotOf(streams, 5u) != 3u || slotOf(streams, 6u) != 4u ||
            slotOf(streams, 7u) != 2u || ! slotsDistinct(streams))
            return fail(L"removing a stream moves no other stream");
        ++step;
        for (uint64_t stream = 10u; stream < 22u; ++stream)
            report(stream, 10u, 2'200u);
        if (slotOf(streams, 10u) != 0u || slotOf(streams, 11u) != 5u || slotOf(streams, 21u) != 15u || ! slotsDistinct(streams))
            return fail(L"sixteen streams take sixteen distinct slots");
        ++step;
        report(30u, 10u, 2'300u); // the array is full: the stream with the oldest report (5, at 1500) is evicted
        if (slotOf(streams, 5u) != kNoStream || slotOf(streams, 30u) != 3u || ! slotsDistinct(streams))
            return fail(L"a stream that replaced an evicted one takes its slot");
    }

    // Lifecycle reports (FileSystemReportFileWork). These read the real tick, so only rules that do not depend on the
    // linger are checked: a reused finished entry is excluded from the slots held, and a stream's own slot is never taken.
    {
        Task lifecycle(owner);
        lifecycle._executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        lifecycle._fileWorkScopes.resize(1u);
        Task::PerItemCallbackCookie itemCookie(0u);
        FileSystemFileWorkProgress work{};
        work.sizeBytes = sizeof(work);
        work.phase     = FileSystemFileWorkPhase::ScopeSupported;
        ++step;
        if (FAILED(lifecycle.FileSystemReportFileWork(&work, &itemCookie)))
            return fail(L"the scope report was refused");
        const auto send = [&lifecycle, &itemCookie, &work](uint64_t leaf, uint64_t stream, FileSystemFileWorkPhase phase) noexcept
        {
            const std::wstring path = std::format(L"mix-{:02}.flac", leaf);
            work                    = {};
            work.sizeBytes          = sizeof(work);
            work.phase              = phase;
            work.kind               = FileSystemFileWorkKind::RegularFile;
            work.logicalLeafId      = leaf;
            work.progressStreamId   = stream;
            work.sourcePath         = path.c_str();
            work.sizeKnown          = TRUE;
            work.totalBytes         = 100u;
            work.completedBytes     = phase == FileSystemFileWorkPhase::Begin ? 0u : 100u;
            work.disposition = phase == FileSystemFileWorkPhase::Terminal ? FileSystemFileWorkDisposition::Completed : FileSystemFileWorkDisposition::None;
            work.status      = S_OK;
            return SUCCEEDED(lifecycle.FileSystemReportFileWork(&work, &itemCookie));
        };
        ++step;
        if (! send(1u, 0u, FileSystemFileWorkPhase::Begin) || ! send(2u, 1u, FileSystemFileWorkPhase::Begin) || ! send(3u, 2u, FileSystemFileWorkPhase::Begin))
            return fail(L"a lifecycle Begin was refused");
        if (slotOf(lifecycle, 0u) != 0u || slotOf(lifecycle, 1u) != 1u || slotOf(lifecycle, 2u) != 2u)
            return fail(L"lifecycle streams take slots 0, 1 and 2");
        ++step;
        if (! send(1u, 0u, FileSystemFileWorkPhase::Terminal) || ! send(4u, 3u, FileSystemFileWorkPhase::Begin))
            return fail(L"a lifecycle report was refused (reuse)");
        // Stream 3 reuses stream 0's finished entry and takes the lowest slot no other stream holds.
        if (slotOf(lifecycle, 0u) != kNoStream || slotOf(lifecycle, 3u) != 0u || slotOf(lifecycle, 1u) != 1u || slotOf(lifecycle, 2u) != 2u)
            return fail(L"a new stream in a finished entry takes the lowest free slot");
        ++step;
        if (! send(2u, 1u, FileSystemFileWorkPhase::Terminal) || ! send(5u, 1u, FileSystemFileWorkPhase::Begin))
            return fail(L"a lifecycle report was refused (next file)");
        if (slotOf(lifecycle, 1u) != 1u || ! slotsDistinct(lifecycle))
            return fail(L"a stream that begins its next file keeps its slot");
    }

    // Both reducers on one array. A finished leaf's linger ends while another stream still copies, a byte report
    // compacts the array, and a third stream begins in the entry the compaction vacated. That entry must hold nothing
    // of its earlier occupant: read as a live leaf this Begin replaces, it would withdraw the task's file counts for good.
    {
        Task compacted(owner);
        compacted._executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        compacted._fileWorkScopes.resize(1u);
        Task::PerItemCallbackCookie itemCookie(0u);
        FileSystemFileWorkProgress work{};
        work.sizeBytes = sizeof(work);
        work.phase     = FileSystemFileWorkPhase::ScopeSupported;
        ++step;
        if (FAILED(compacted.FileSystemReportFileWork(&work, &itemCookie)))
            return fail(L"the scope report was refused (compaction)");
        const auto send = [&compacted, &itemCookie, &work](uint64_t leaf, uint64_t stream, FileSystemFileWorkPhase phase) noexcept
        {
            const std::wstring path = std::format(L"mix-{:02}.flac", leaf);
            work                    = {};
            work.sizeBytes          = sizeof(work);
            work.phase              = phase;
            work.kind               = FileSystemFileWorkKind::RegularFile;
            work.logicalLeafId      = leaf;
            work.progressStreamId   = stream;
            work.sourcePath         = path.c_str();
            work.sizeKnown          = TRUE;
            work.totalBytes         = 100u;
            work.completedBytes     = phase == FileSystemFileWorkPhase::Begin ? 0u : 100u;
            work.disposition = phase == FileSystemFileWorkPhase::Terminal ? FileSystemFileWorkDisposition::Completed : FileSystemFileWorkDisposition::None;
            work.status      = S_OK;
            return SUCCEEDED(compacted.FileSystemReportFileWork(&work, &itemCookie));
        };
        ++step;
        if (! send(1u, 0u, FileSystemFileWorkPhase::Begin) || ! send(1u, 0u, FileSystemFileWorkPhase::Terminal) ||
            ! send(2u, 1u, FileSystemFileWorkPhase::Begin) || ! compacted._copiedFileCountKnown.load(std::memory_order_acquire))
            return fail(L"a lifecycle report was refused, or the counts were not known (compaction)");
        ++step;
        {
            std::scoped_lock lock(compacted._inFlightFilesMutex);
            compacted._inFlightFiles[0].lastUpdateTick = 1u; // long ago: the finished leaf's linger is over
        }
        FolderWindowFileOperationsStateInternal::UpdateInFlightFileProgress(compacted, &itemCookie, 1u, L"mix-02.flac", 100u, 50u, GetTickCount64(), false);
        {
            std::scoped_lock lock(compacted._inFlightFilesMutex);
            if (compacted._inFlightFileCount != 1u || compacted._inFlightFiles[0].logicalLeafId != 2u || compacted._inFlightFiles[1].explicitLifecycle ||
                ! compacted._inFlightFiles[1].sourcePath.empty())
                return fail(L"the byte report compacts the array and leaves the vacated entry empty");
        }
        ++step;
        if (! send(3u, 2u, FileSystemFileWorkPhase::Begin) || ! compacted._copiedFileCountKnown.load(std::memory_order_acquire))
            return fail(L"a stream that begins in a vacated entry keeps the task's file counts known");
        {
            std::scoped_lock lock(compacted._inFlightFilesMutex);
            if (compacted._inFlightFileCount != 2u || compacted._inFlightFiles[1].logicalLeafId != 3u)
                return fail(L"the new stream took the vacated entry");
        }
    }
    return true;
}

// I63 G7 (D-14): keep the computer awake while a task runs. The pure predicate table first, then the process-wide request
// through the fake power API. Real tasks are not driven through these states here: Task::UpdateKeepAwake only reads
// atomics the task already keeps and hands them to the lease, so the lease is driven with synthetic facts for two tasks.
// What this does not prove: that each hook in the task code fires (admission, leaving the operation, pause, queue pause,
// the wait for an answer, cancel, finish).
[[nodiscard]] bool RunFileOperationKeepAwakeSelfTest(FolderWindow::FileOperationState& owner) noexcept
{
    using FileOperations::KeepAwakeFacts;
    using FileOperations::KeepAwakeLease;
    using FileOperations::PowerRequest;
    using FileOperations::TaskNeedsMachineAwake;

    unsigned int step = 0u;
    const auto fail   = [&step](const wchar_t* what) noexcept
    {
        AppendLog(std::format(L"Keep awake self-test failed at step {}: {}", step, what));
        return false;
    };
    const auto facts = [](bool admitted, bool finished, bool stopping, bool userPaused, bool awaitingAnswer) noexcept
    {
        KeepAwakeFacts value{};
        value.admitted       = admitted;
        value.finished       = finished;
        value.stopping       = stopping;
        value.userPaused     = userPaused;
        value.awaitingAnswer = awaitingAnswer;
        return value;
    };

    // The predicate: admitted, finished, stopping, userPaused, awaitingAnswer.
    struct Row final
    {
        KeepAwakeFacts input;
        bool awake;
    };
    const std::array<Row, 12u> rows{{
        {facts(false, false, false, false, false), false}, // only queued
        {facts(true, false, false, false, false), true},   // running
        {facts(true, false, false, true, false), false},   // paused by the user
        {facts(true, false, false, false, true), false},   // waiting for an answer
        {facts(true, false, false, true, true), false},
        {facts(true, false, true, false, false), true},  // stopping
        {facts(true, false, true, true, false), true},   // stopping although it was paused
        {facts(true, false, true, false, true), true},   // stopping although it was waiting
        {facts(true, true, false, false, false), false}, // finished
        {facts(true, true, true, false, false), false},
        {facts(false, false, true, false, false), false}, // a stop request on a task that never started working
        {facts(false, false, false, true, true), false},
    }};
    for (const Row& row : rows)
    {
        ++step;
        if (TaskNeedsMachineAwake(row.input) != row.awake)
            return fail(L"predicate table");
    }
    for (unsigned int mask = 0u; mask < 32u; ++mask)
    {
        ++step;
        const KeepAwakeFacts value = facts((mask & 1u) != 0u, (mask & 2u) != 0u, (mask & 4u) != 0u, (mask & 8u) != 0u, (mask & 16u) != 0u);
        const bool expected        = ! value.admitted || value.finished ? false : value.stopping ? true : ! (value.userPaused || value.awaitingAnswer);
        if (TaskNeedsMachineAwake(value) != expected)
            return fail(L"predicate exhaustive");
    }

    const KeepAwakeFacts queued   = facts(false, false, false, false, false);
    const KeepAwakeFacts running  = facts(true, false, false, false, false);
    const KeepAwakeFacts awaiting = facts(true, false, false, false, true);
    const KeepAwakeFacts paused   = facts(true, false, false, true, false);
    const KeepAwakeFacts stopping = facts(true, false, true, true, false);
    const KeepAwakeFacts finished = facts(true, true, false, false, false);
    const auto apply    = [](KeepAwakeLease& lease, const KeepAwakeFacts& current) noexcept { lease.Update([&current]() noexcept { return current; }); };
    const auto counters = [](const PowerRequest& request, uint32_t holders, uint32_t acquires, uint32_t releases, uint32_t failures, bool held) noexcept
    {
        const PowerRequest::Counters actual = request.GetCounters();
        return actual.holders == holders && actual.acquireTransitions == acquires && actual.releaseTransitions == releases && actual.failures == failures &&
               actual.held == held;
    };

    // Two tasks and one request.
    {
        KeepAwakeSelfTestApi fake;
        PowerRequest request(fake.Api());
        KeepAwakeLease first(request);
        KeepAwakeLease second(request);

        ++step;
        apply(first, queued);
        apply(second, queued);
        if (! counters(request, 0u, 0u, 0u, 0u, false) || fake.createCalls != 0u)
            return fail(L"queued tasks acquire nothing");
        ++step;
        apply(first, running);
        if (! counters(request, 1u, 1u, 0u, 0u, true) || fake.createCalls != 1u || fake.setCalls != 1u)
            return fail(L"the first running task acquires once");
        ++step;
        apply(second, running);
        if (! counters(request, 2u, 1u, 0u, 0u, true) || fake.setCalls != 1u)
            return fail(L"a second running task adds no acquire");
        ++step;
        apply(first, finished);
        apply(first, finished);
        if (! counters(request, 1u, 1u, 0u, 0u, true) || fake.clearCalls != 0u)
            return fail(L"one task ending keeps the request, and Update is idempotent");
        ++step;
        apply(second, awaiting);
        if (! counters(request, 0u, 1u, 1u, 0u, false) || fake.clearCalls != 1u)
            return fail(L"the only task waiting for an answer releases");
        ++step;
        apply(second, running);
        if (! counters(request, 1u, 2u, 1u, 0u, true) || fake.createCalls != 1u || fake.setCalls != 2u)
            return fail(L"the answer acquires again and reuses the handle");
        ++step;
        apply(second, paused);
        if (! counters(request, 0u, 2u, 2u, 0u, false) || fake.clearCalls != 2u)
            return fail(L"a user pause releases");
        ++step;
        apply(second, running);
        if (! counters(request, 1u, 3u, 2u, 0u, true))
            return fail(L"resume acquires");
        ++step;
        apply(second, stopping);
        if (! counters(request, 1u, 3u, 2u, 0u, true))
            return fail(L"stopping while paused keeps the request");
        ++step;
        apply(second, finished);
        if (! counters(request, 0u, 3u, 3u, 0u, false) || fake.clearCalls != 3u)
            return fail(L"the last task ending releases");
        ++step;
        {
            KeepAwakeLease scoped(request);
            apply(scoped, running);
            if (! counters(request, 1u, 4u, 3u, 0u, true) || ! scoped.Held())
                return fail(L"a lease acquires");
        }
        if (! counters(request, 0u, 4u, 4u, 0u, false))
            return fail(L"destroying a lease that holds releases");
        ++step;
        request.Release();
        if (! counters(request, 0u, 4u, 4u, 0u, false))
            return fail(L"a Release without an Acquire is ignored");
    }

    // The power API refuses: nothing throws, nothing is left half counted, and the next attempt tries again.
    {
        KeepAwakeSelfTestApi fake;
        PowerRequest request(fake.Api());
        KeepAwakeLease lease(request);

        ++step;
        fake.failCreate = true;
        apply(lease, running);
        if (! counters(request, 1u, 1u, 0u, 1u, false) || fake.setCalls != 0u || ! lease.Held())
            return fail(L"a refused create is counted and not held");
        if (request.GetCounters().lastError != ERROR_ACCESS_DENIED || ! request.GetCounters().failureLogged)
            return fail(L"the exact create error is retained and the one diagnostic is recorded");
        ++step;
        apply(lease, finished);
        if (! counters(request, 0u, 1u, 1u, 1u, false) || fake.clearCalls != 0u)
            return fail(L"release after a refused create clears nothing");
        ++step;
        fake.failCreate = false;
        fake.failSet    = true;
        apply(lease, running);
        if (! counters(request, 1u, 2u, 1u, 2u, false) || fake.createCalls != 2u || fake.setCalls != 1u)
            return fail(L"a refused set is counted and not held");
        if (request.GetCounters().lastError != ERROR_NOT_SUPPORTED || ! request.GetCounters().failureLogged)
            return fail(L"the exact set error is retained without resetting the diagnostic guard");
        ++step;
        apply(lease, finished);
        if (! counters(request, 0u, 2u, 2u, 2u, false) || fake.clearCalls != 0u)
            return fail(L"release after a refused set clears nothing");
        ++step;
        fake.failSet = false;
        apply(lease, running);
        if (! counters(request, 1u, 3u, 2u, 2u, true) || fake.createCalls != 2u || fake.setCalls != 2u)
            return fail(L"the next attempt succeeds with the handle created before");
        ++step;
        fake.failClear = true;
        apply(lease, finished);
        if (! counters(request, 0u, 3u, 3u, 3u, false) || fake.clearCalls != 1u)
            return fail(L"a refused clear is counted and the request is no longer claimed");
        DWORD handleFlags = 0u;
        if (request.GetCounters().lastError != ERROR_GEN_FAILURE || GetHandleInformation(fake.lastCreated, &handleFlags) != FALSE ||
            GetLastError() != ERROR_INVALID_HANDLE)
            return fail(L"a refused clear preserves its error and closes the kernel request handle");
        ++step;
        fake.failClear = false;
        apply(lease, running);
        if (! counters(request, 1u, 4u, 3u, 3u, true) || fake.createCalls != 3u || fake.setCalls != 3u)
            return fail(L"an Acquire after failed clear creates a fresh request instead of reusing leaked kernel state");
        apply(lease, finished);
        if (! counters(request, 0u, 4u, 4u, 3u, false) || fake.clearCalls != 2u)
            return fail(L"the replacement request clears normally with unchanged failure count");
    }

    // A set refused for the first holder is tried again for the next one, while the first still holds its share.
    {
        KeepAwakeSelfTestApi fake;
        PowerRequest request(fake.Api());
        KeepAwakeLease first(request);
        KeepAwakeLease second(request);

        ++step;
        fake.failSet = true;
        apply(first, running);
        if (! counters(request, 1u, 1u, 0u, 1u, false) || fake.setCalls != 1u)
            return fail(L"the first holder's refused set is counted and not held");
        ++step;
        fake.failSet = false;
        apply(second, running);
        if (! counters(request, 2u, 1u, 0u, 1u, true) || fake.createCalls != 1u || fake.setCalls != 2u)
            return fail(L"a second holder tries the refused set again");
        ++step;
        apply(first, finished);
        if (! counters(request, 1u, 1u, 0u, 1u, true) || fake.clearCalls != 0u)
            return fail(L"the request stays while one holder remains");
        ++step;
        apply(second, finished);
        if (! counters(request, 0u, 1u, 1u, 1u, false) || fake.clearCalls != 1u)
            return fail(L"the last holder releases the request the retry obtained");
    }
    // Deadline arithmetic is checked without a two-minute sleep.
    ++step;
    if (FileOperations::AwaitingAnswerWaitMilliseconds(0u) != FileOperations::kAwaitingAnswerQuietMs ||
        FileOperations::AwaitingAnswerWaitMilliseconds(FileOperations::kAwaitingAnswerQuietMs - 1u) != 1u ||
        FileOperations::AwaitingAnswerWaitMilliseconds(FileOperations::kAwaitingAnswerQuietMs) != INFINITE ||
        FileOperations::AwaitingAnswerWaitMilliseconds(7u, 10u) != 3u)
        return fail(L"quiet wait uses one remaining deadline and becomes event-only at the boundary");

    using Task   = FolderWindow::FileOperationState::Task;
    using Action = Task::ConflictAction;
    enum class WaitControl : uint8_t
    {
        Preanswered,
        Precancelled,
        QuietCancel,
        QuietResume,
        StopOnly,
        Spurious
    };
    for (const WaitControl control : {WaitControl::Preanswered,
                                      WaitControl::Precancelled,
                                      WaitControl::QuietCancel,
                                      WaitControl::QuietResume,
                                      WaitControl::StopOnly,
                                      WaitControl::Spurious})
    {
        ++step;
        Task task(owner);
        // No operation slot, provider, real power request, owner posts or diagnostic publication.
        task._state = nullptr;
        task._dbgConflictQuietPeriodMs.store(25u, std::memory_order_relaxed);
        if (! task._conflictArbiter.decisionEvent || ! task._keepAwakeWorkEvent)
            return fail(L"owned conflict wait events could not be created");
        {
            std::scoped_lock lock(task._conflictArbiter.mutex);
            auto& prompt              = task._conflictArbiter.prompt;
            prompt.active             = true;
            prompt.generation         = 1u;
            prompt.buttonsPublishable = true;
            prompt.actions[0]         = Action::Skip;
            prompt.actions[1]         = Action::Cancel;
            prompt.actionCount        = 2u;
        }
        const auto answer = [&]() noexcept
        {
            std::scoped_lock lock(task._conflictArbiter.mutex);
            return FileOperations::ConflictBroker::TryAcceptLocked(
                task._conflictArbiter, 1u, Action::Skip, false, FileOperations::ConflictBroker::Placement::Published);
        };
        if (control == WaitControl::Preanswered || control == WaitControl::Precancelled)
        {
            if (! answer())
                return fail(L"the before-wait decision could not be published");
            if (control == WaitControl::Precancelled)
                task.RequestCancel();
        }
        if (control == WaitControl::Spurious)
            static_cast<void>(SetEvent(task._conflictArbiter.decisionEvent.get()));

        wil::unique_event_nothrow done;
        done.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
        if (! done)
            return fail(L"owned completion event could not be created");
        std::atomic<Action> result{Action::None};
        std::jthread worker([&]() noexcept
        {
            const auto decision = FolderWindowFileOperationsStateInternal::WaitForConflictDecision(task, nullptr, Task::ConflictDecisionScope{}, nullptr);
            result.store(decision.first, std::memory_order_release);
            static_cast<void>(SetEvent(done.get()));
        });
        const auto drain = wil::scope_exit([&]() noexcept
        {
            task.RequestCancel();
            if (worker.joinable())
                worker.join();
        });
        if (control != WaitControl::Preanswered && control != WaitControl::Precancelled)
        {
            const auto quiet = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                return task._dbgConflictDecisionInfiniteWaitCount.load(std::memory_order_acquire) >= 1u;
            }, {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"private conflict wait reaches quiet deadline"});
            if (! quiet.conditionMet || ! task._awaitingAnswerQuiet.load(std::memory_order_acquire) || WaitForSingleObject(done.get(), 0u) != WAIT_TIMEOUT)
                return fail(L"quiet conflict wait must remain asleep without accepting a spurious decision");
            if (control == WaitControl::QuietResume)
            {
                task.NoteKeepAwakeWork(GetTickCount64());
                const auto quietAgain = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                    return task._dbgConflictDecisionInfiniteWaitCount.load(std::memory_order_acquire) >= 2u;
                }, {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"resumed stream rearms private conflict quiet deadline"});
                if (! quietAgain.conditionMet || ! task._awaitingAnswerQuiet.load(std::memory_order_acquire) || ! answer())
                    return fail(L"resumed work must rearm the deadline and retain its later decision");
            }
            else if (control == WaitControl::Spurious)
            {
                if (! answer())
                    return fail(L"a real decision after a spurious wake must still be accepted");
            }
            else if (control == WaitControl::StopOnly)
                task._stopSource.request_stop(); // No RequestCancel/decision-event signal: the stop callback must wake it.
            else
                task.RequestCancel();
        }
        if (WaitForSingleObject(done.get(), 5'000u) != WAIT_OBJECT_0)
            return fail(L"decision or durable cancellation must wake the event-only wait");
        worker.join();
        const Action expected =
            control == WaitControl::Preanswered || control == WaitControl::QuietResume || control == WaitControl::Spurious ? Action::Skip : Action::Cancel;
        const uint32_t calls = task._dbgConflictDecisionWaitCount.load(std::memory_order_acquire);
        if (result.load(std::memory_order_acquire) != expected || task._awaitingUserDecision.load(std::memory_order_acquire) ||
            calls > (control == WaitControl::QuietResume ? 6u
                     : control == WaitControl::Spurious  ? 3u
                                                         : 2u))
            return fail(L"event waits preserve exact decisions/cancellation with bounded deadline wakes");
        Debug::Perf::Emit(L"FileOps.Power.ConflictWaitCalls", L"control,calls", 0u, static_cast<uint64_t>(control), calls, S_OK);
    }
    return true;
}

// I63 B6: the bulk policies at a collision. A Copy of several files onto a folder that already holds each of them. The
// first collision whose pair matches a policy is answered with it from the prompt's More list; every later file whose
// pair matches is then decided without a question, and every other pair still asks. A policy skip leaves the
// destination's bytes and creates no stage; a policy replacement is an ordinary Replace. Nothing is remembered as an
// Apply-to-all answer. The three cases: Skip all with the same size and date, Replace all older, and both in one task.
[[nodiscard]] bool RunFileOperationBulkCollisionPoliciesSelfTest(FolderWindow::FileOperationState& fileOperations,
                                                                 const wil::com_ptr<IFileSystem>& localFileSystem,
                                                                 const std::filesystem::path& testRoot) noexcept
{
    using State  = FolderWindow::FileOperationState;
    using Action = State::Task::ConflictAction;
    wil::com_ptr<IFileSystemIO> io;
    const auto root = testRoot / L"bulk-collision-policies";
    if (! localFileSystem || ! localFileSystem.try_query_to(io.put()) || ! RecreateEmptyDirectory(root))
        return false;
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });

    // How the existing file compares with the incoming one. Dates differ by one day, far beyond the two-second
    // tolerance and the one-hour class.
    enum class CollisionPair : uint8_t
    {
        SameSizeAndDate,
        OtherSize,        // same date; the existing file is larger
        DestinationOlder, // same size; the incoming file is newer
        DestinationNewer, // same size; the incoming file is older
    };
    struct Scenario final
    {
        const wchar_t* name     = L"";
        bool chooseSkipSame     = false;        // the first matching prompt is answered Skip all with the same size and date
        bool chooseReplaceOlder = false;        // the first matching prompt is answered Replace all older
        Action laterOtherAnswer = Action::Skip; // every other prompt after the first one, which is answered Skip
        std::vector<CollisionPair> pairs;
    };
    const std::array<Scenario, 3u> scenarios{{
        {L"skip-same",
         true,
         false,
         Action::Overwrite,
         {CollisionPair::SameSizeAndDate,
          CollisionPair::SameSizeAndDate,
          CollisionPair::SameSizeAndDate,
          CollisionPair::OtherSize,
          CollisionPair::SameSizeAndDate,
          CollisionPair::SameSizeAndDate,
          CollisionPair::SameSizeAndDate,
          CollisionPair::OtherSize,
          CollisionPair::SameSizeAndDate,
          CollisionPair::SameSizeAndDate,
          CollisionPair::SameSizeAndDate,
          CollisionPair::SameSizeAndDate}},
        {L"replace-older",
         false,
         true,
         Action::Skip,
         {CollisionPair::DestinationOlder,
          CollisionPair::DestinationNewer,
          CollisionPair::DestinationOlder,
          CollisionPair::SameSizeAndDate,
          CollisionPair::DestinationOlder,
          CollisionPair::DestinationNewer,
          CollisionPair::DestinationOlder,
          CollisionPair::SameSizeAndDate}},
        {L"both",
         true,
         true,
         Action::Skip,
         {CollisionPair::SameSizeAndDate,
          CollisionPair::DestinationOlder,
          CollisionPair::SameSizeAndDate,
          CollisionPair::DestinationOlder,
          CollisionPair::DestinationNewer,
          CollisionPair::SameSizeAndDate,
          CollisionPair::DestinationOlder}},
    }};
    constexpr uint64_t kBaseTicks = 133'500'000'000'000'000ull;
    constexpr uint64_t kDayTicks  = 864'000'000'000ull;
    const auto setLastWriteTime   = [](const std::filesystem::path& path, uint64_t ticks) noexcept
    {
        const wil::unique_hfile file(CreateFileW(path.c_str(),
                                                 FILE_WRITE_ATTRIBUTES,
                                                 FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                 nullptr,
                                                 OPEN_EXISTING,
                                                 FILE_ATTRIBUTE_NORMAL,
                                                 nullptr));
        FILETIME time{};
        time.dwLowDateTime  = static_cast<DWORD>(ticks & 0xFFFF'FFFFull);
        time.dwHighDateTime = static_cast<DWORD>(ticks >> 32u);
        return file.is_valid() && SetFileTime(file.get(), nullptr, nullptr, &time) != FALSE;
    };
    const std::string sourceBytes(64u, 's');
    const std::string sameSizeBytes(64u, 'd');
    const std::string otherSizeBytes(96u, 'd');

    bool passed            = true;
    unsigned int caseIndex = 0u;
    for (const Scenario& scenario : scenarios)
    {
        const auto caseRoot          = root / std::format(L"case-{}", caseIndex++);
        const auto sourceFolder      = caseRoot / L"source";
        const auto destinationFolder = caseRoot / L"destination";
        if (! RecreateEmptyDirectory(sourceFolder) || ! RecreateEmptyDirectory(destinationFolder))
            return false;
        std::vector<std::filesystem::path> sources;
        for (size_t index = 0u; index < scenario.pairs.size(); ++index)
        {
            const CollisionPair pair        = scenario.pairs[index];
            const auto source               = sourceFolder / std::format(L"item-{:02}.bin", index);
            const auto destination          = destinationFolder / source.filename();
            const uint64_t destinationTicks = pair == CollisionPair::DestinationOlder   ? kBaseTicks - kDayTicks
                                              : pair == CollisionPair::DestinationNewer ? kBaseTicks + kDayTicks
                                                                                        : kBaseTicks;
            if (! WriteFileTextFsIo(io, source, sourceBytes) ||
                ! WriteFileTextFsIo(io, destination, pair == CollisionPair::OtherSize ? otherSizeBytes : sameSizeBytes) ||
                ! setLastWriteTime(source, kBaseTicks) || ! setLastWriteTime(destination, destinationTicks))
                return false;
            sources.push_back(source);
        }

        FileOperations::QualifiedEndpoint endpoint{
            .pluginId               = L"builtin/file-system",
            .instanceId             = L"host/default",
            .profileId              = L"local",
            .rootId                 = L"local-test-volume",
            .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };
        endpoint.verificationHostReadback = true;
        FileOperations::TransferPlan transfer{};
        transfer.intent              = FileOperations::TransferIntent::Copy;
        transfer.strategy            = FileOperations::OperationStrategy::Copy;
        transfer.sourceEndpoint      = endpoint;
        transfer.destinationEndpoint = endpoint;
        for (const auto& source : sources)
            transfer.selectedItems.emplace_back().providerPath = source.native();
        transfer.destination.providerFolderPath = destinationFolder.native();
        const auto plans                        = std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}});
        State::Task task(fileOperations);
        task._operation             = FILESYSTEM_COPY;
        task._executionMode         = State::ExecutionMode::PerItem;
        task._fileSystem            = localFileSystem;
        task._sourcePluginId        = endpoint.pluginId;
        task._sourcePluginShortId   = L"file";
        task._sourceInstanceContext = endpoint.instanceId;
        task._perItemMaxConcurrency = 1u;
        for (const auto& source : sources)
        {
            task._sourcePaths.push_back(source);
            task._sourcePathAttributesHint.push_back(GetFileAttributesW(source.c_str()));
        }
        task.InitializeAdmissionInputs(plans,
                                       {.fileSystem      = localFileSystem,
                                        .folder          = destinationFolder,
                                        .pluginId        = endpoint.pluginId,
                                        .pluginShortId   = L"file",
                                        .instanceContext = endpoint.instanceId});
        task.InitializeSourceItemResultBuilders();
        const auto diagnosticsCleanup = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });
        if (FAILED(task.PrepareMutationInterlockScopes()))
            return false;

        // The executor blocks on each decision: answer from a second thread. Every published prompt is checked: each
        // policy is offered exactly when the prompt's own pair matches it, under More, never the default or Escape.
        std::atomic<unsigned int> answered{0u};
        std::atomic<bool> offersMatchPairs{true};
        HRESULT executeHr = E_PENDING;
        {
            const std::jthread responder([&](const std::stop_token stop) noexcept
            {
                const auto listed = [](const auto& actions, size_t count, const Action action) noexcept
                {
                    const auto end = actions.begin() + static_cast<std::ptrdiff_t>(std::min(count, actions.size()));
                    return std::find(actions.begin(), end, action) != end;
                };
                uint64_t answeredGeneration = 0u;
                bool skipSameChosen         = false;
                bool replaceOlderChosen     = false;
                unsigned int otherAnswers   = 0u;
                while (! stop.stop_requested())
                {
                    if (const auto prompt = TryGetConflictPromptCopy(&task);
                        prompt.has_value() && prompt->generation != answeredGeneration && ! prompt->metadataLoading && prompt->buttonsPublishable)
                    {
                        const bool sameSizeAndDate    = prompt->comparison.date == FileOperations::ConflictDateRelation::Same &&
                                                        prompt->comparison.size == FileOperations::ConflictSizeRelation::Same;
                        const bool incomingNewer      = prompt->comparison.date == FileOperations::ConflictDateRelation::Newer;
                        const bool offersSkipSame     = listed(prompt->actions, prompt->actionCount, Action::SkipSameSizeAndDate);
                        const bool offersReplaceOlder = listed(prompt->actions, prompt->actionCount, Action::ReplaceOlder);
                        if (offersSkipSame != sameSizeAndDate || offersReplaceOlder != incomingNewer ||
                            listed(prompt->primaryActions, prompt->primaryActionCount, Action::SkipSameSizeAndDate) ||
                            listed(prompt->primaryActions, prompt->primaryActionCount, Action::ReplaceOlder) || prompt->defaultAction != Action::Cancel ||
                            prompt->escapeAction != Action::Cancel)
                            offersMatchPairs.store(false, std::memory_order_release);
                        const bool takeSkipSame = offersSkipSame && scenario.chooseSkipSame && ! skipSameChosen;
                        const bool takeReplace  = ! takeSkipSame && offersReplaceOlder && scenario.chooseReplaceOlder && ! replaceOlderChosen;
                        const Action answer     = takeSkipSame         ? Action::SkipSameSizeAndDate
                                                  : takeReplace        ? Action::ReplaceOlder
                                                  : otherAnswers == 0u ? Action::Skip
                                                                       : scenario.laterOtherAnswer;
                        // An answer the prompt does not offer is a failure, answered Cancel so the case cannot hang.
                        const bool answerOffered = listed(prompt->actions, prompt->actionCount, answer);
                        if (! answerOffered)
                            offersMatchPairs.store(false, std::memory_order_release);
                        if (task.SubmitConflictDecision(prompt->generation, answerOffered ? answer : Action::Cancel, false))
                        {
                            answeredGeneration = prompt->generation;
                            answered.fetch_add(1u, std::memory_order_acq_rel);
                            if (takeSkipSame)
                                skipSameChosen = true;
                            else if (takeReplace)
                                replaceOlderChosen = true;
                            else
                                ++otherAnswers;
                        }
                    }
                    Sleep(2);
                }
            });
            executeHr = task.ExecuteOperation();
        }
        const HRESULT finalHr = task.FinalizeTypedItemResults(executeHr);
        task._resultHr.store(finalHr, std::memory_order_release);
        const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);

        // What each pair must have become, in item order.
        size_t expectedPrompts      = 0u;
        size_t expectedSkipSame     = 0u;
        size_t expectedReplaceOlder = 0u;
        size_t expectedReplaced     = 0u;
        size_t expectedSkipped      = 0u;
        size_t otherPrompts         = 0u;
        bool skipSameSeen           = false;
        bool replaceOlderSeen       = false;
        bool bytesMatch             = true;
        for (size_t index = 0u; index < scenario.pairs.size(); ++index)
        {
            const CollisionPair pair = scenario.pairs[index];
            bool replaced            = false;
            if (pair == CollisionPair::SameSizeAndDate && scenario.chooseSkipSame)
            {
                ++expectedSkipSame;
                if (! skipSameSeen)
                {
                    skipSameSeen = true;
                    ++expectedPrompts;
                }
            }
            else if (pair == CollisionPair::DestinationOlder && scenario.chooseReplaceOlder)
            {
                ++expectedReplaceOlder;
                replaced = true;
                if (! replaceOlderSeen)
                {
                    replaceOlderSeen = true;
                    ++expectedPrompts;
                }
            }
            else
            {
                ++expectedPrompts;
                replaced = otherPrompts != 0u && scenario.laterOtherAnswer == Action::Overwrite;
                ++otherPrompts;
            }
            if (replaced)
                ++expectedReplaced;
            else
                ++expectedSkipped;
            std::string actualSource;
            std::string actualDestination;
            const std::string& destinationBefore = pair == CollisionPair::OtherSize ? otherSizeBytes : sameSizeBytes;
            bytesMatch                           = bytesMatch && ReadFileTextFsIo(io, sources[index], actualSource) && actualSource == sourceBytes &&
                                                   ReadFileTextFsIo(io, destinationFolder / sources[index].filename(), actualDestination) &&
                                                   actualDestination == (replaced ? sourceBytes : destinationBefore);
        }
        size_t destinationEntries = 0u;
        size_t leftoverEntries    = 0u;
        std::error_code error;
        for (std::filesystem::directory_iterator it(destinationFolder, error), end; ! error && it != end; it.increment(error))
        {
            ++destinationEntries;
            const std::wstring leaf = it->path().filename().native();
            if (FileOperationArtifacts::HasPossibleArtifactName(leaf))
                ++leftoverEntries;
        }
        const size_t cachedDecisions =
            static_cast<size_t>(std::ranges::count_if(task._conflictArbiter.decisionCache, [](const auto& entry) noexcept { return entry.has_value(); }));
        const bool grants  = (task._conflictArbiter.skipSameSizeAndDateGrantGeneration != 0u) == scenario.chooseSkipSame &&
                             (task._conflictArbiter.replaceOlderGrantGeneration != 0u) == scenario.chooseReplaceOlder;
        const bool prompts = answered.load(std::memory_order_acquire) == expectedPrompts && task._perf.conflictPromptCount == expectedPrompts;
        const bool counts  = task._skippedSameSizeAndDateCount.load(std::memory_order_acquire) == expectedSkipSame &&
                             task._replacedOlderCount.load(std::memory_order_acquire) == expectedReplaceOlder &&
                             summary.skippedSameSizeAndDateCount == expectedSkipSame && summary.replacedOlderCount == expectedReplaceOlder &&
                             summary.skippedItemCount == expectedSkipped && summary.publishedItemCount == expectedReplaced && summary.failedItemCount == 0u;
        const bool matches = prompts && offersMatchPairs.load(std::memory_order_acquire) && counts && grants && cachedDecisions == 0u && bytesMatch &&
                             destinationEntries == scenario.pairs.size() && leftoverEntries == 0u && executeHr != HRESULT_FROM_WIN32(ERROR_CANCELLED);
        AppendLog(std::format(L"Bulk collision policies {}: execute=0x{:08X} final=0x{:08X} prompts={}/{} answered={} offers={} skippedSame={}/{} "
                              L"replacedOlder={}/{} skipped={}/{} published={}/{} failed={} cached={} grants={} bytes={} entries={} leftovers={} passed={}",
                              scenario.name,
                              static_cast<unsigned long>(executeHr),
                              static_cast<unsigned long>(finalHr),
                              static_cast<uint64_t>(task._perf.conflictPromptCount),
                              expectedPrompts,
                              answered.load(std::memory_order_acquire),
                              offersMatchPairs.load(std::memory_order_acquire),
                              summary.skippedSameSizeAndDateCount,
                              expectedSkipSame,
                              summary.replacedOlderCount,
                              expectedReplaceOlder,
                              summary.skippedItemCount,
                              expectedSkipped,
                              summary.publishedItemCount,
                              expectedReplaced,
                              summary.failedItemCount,
                              cachedDecisions,
                              grants,
                              bytesMatch,
                              destinationEntries,
                              leftoverEntries,
                              matches));
        passed = matches && passed;
    }
    return passed;
}

// I63 D-5: Delete source where the source's location has no Recycle Bin. The Local provider's self-test seam
// REDSALAMANDER_FILEOPS_RECYCLE_UNAVAILABLE_DIRECTORY refuses, before any change, every Recycle of an item directly inside
// one folder, as a location without a Recycle Bin does. The task then asks one question on the card and keeps the answer
// for the rest of the task. Four cases: Permanently delete over two files (asked once; both sources gone, both
// destinations untouched); Keep source over two files (asked once; everything stays); a source rewritten while the
// question is open (kept); a destination replaced, while the question is open, by a second name of the source itself, a
// hard link (kept: the source is never deleted).
[[nodiscard]] bool RunFileOperationDiscardSourceWithoutRecycleBinSelfTest(FolderWindow::FileOperationState& fileOperations,
                                                                          const wil::com_ptr<IFileSystem>& localFileSystem,
                                                                          const std::filesystem::path& testRoot) noexcept
{
    using State                                              = FolderWindow::FileOperationState;
    using Action                                             = State::Task::ConflictAction;
    constexpr std::wstring_view kRecycleUnavailableDirectory = L"REDSALAMANDER_FILEOPS_RECYCLE_UNAVAILABLE_DIRECTORY";
    wil::com_ptr<IFileSystemIO> io;
    const auto root = testRoot / L"discard-source-no-recycle-bin";
    if (! localFileSystem || ! localFileSystem.try_query_to(io.put()) || ! RecreateEmptyDirectory(root))
        return false;
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });

    // The card and the decision mirror ask it in the same words; Skip reads Keep source only on this Move question.
    {
        FileOperations::ConflictPromptTextFacts questionFacts{};
        questionFacts.operation = FILESYSTEM_MOVE;
        questionFacts.bucket    = static_cast<uint8_t>(FileOperations::ConflictClass::RecycleFailed);
        questionFacts.status    = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        const bool words =
            FileOperations::BuildConflictPromptQuestion(questionFacts) == LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_MOVE_SOURCE_NO_RECYCLE_BIN) &&
            FileOperations::FormatConflictActionLabel(Action::Skip, FileOperations::ConflictClass::RecycleFailed, FILESYSTEM_MOVE) ==
                LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_KEEP_SOURCE) &&
            FileOperations::FormatConflictActionLabel(Action::Skip, FileOperations::ConflictClass::RecycleFailed, FILESYSTEM_DELETE) ==
                LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_SKIP);
        if (! words)
        {
            AppendLog(L"Delete source without a Recycle Bin: the question's words or its Keep source button are wrong.");
            return false;
        }
    }

    enum class Shape : uint8_t
    {
        DeletePermanently,   // two files, answered Permanently delete
        KeepSources,         // two files, answered Keep source
        SourceChanged,       // the source is rewritten while the question is open
        DestinationIsSource, // the destination becomes a second name of the source while the question is open
        CancelQuestion,      // the question is answered Cancel: the task is canceled and every source stays
    };
    struct Scenario final
    {
        const wchar_t* name = L"";
        Shape shape         = Shape::DeletePermanently;
    };
    static constexpr std::array kScenarios{
        Scenario{L"delete-permanently", Shape::DeletePermanently},
        Scenario{L"keep-source", Shape::KeepSources},
        Scenario{L"source-changed", Shape::SourceChanged},
        Scenario{L"destination-is-source", Shape::DestinationIsSource},
        Scenario{L"cancel-question", Shape::CancelQuestion},
    };
    constexpr uint64_t kTicks   = 133'500'000'000'000'000ull;
    const auto setLastWriteTime = [](const std::filesystem::path& path, uint64_t ticks) noexcept
    {
        const wil::unique_hfile file(CreateFileW(path.c_str(),
                                                 FILE_WRITE_ATTRIBUTES,
                                                 FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                 nullptr,
                                                 OPEN_EXISTING,
                                                 FILE_ATTRIBUTE_NORMAL,
                                                 nullptr));
        FILETIME time{};
        time.dwLowDateTime  = static_cast<DWORD>(ticks & 0xFFFF'FFFFull);
        time.dwHighDateTime = static_cast<DWORD>(ticks >> 32u);
        return file.is_valid() && SetFileTime(file.get(), nullptr, nullptr, &time) != FALSE;
    };
    const std::string sourceBytes(64u, 's');
    const std::string destinationBytes(64u, 'd');
    const std::string changedBytes(101u, 'x');
    const auto holds = [&io](const std::filesystem::path& path, const std::string& expected) noexcept
    {
        std::string actual;
        return ReadFileTextFsIo(io, path, actual) && actual == expected;
    };

    bool passed            = true;
    unsigned int caseIndex = 0u;
    for (const Scenario& scenario : kScenarios)
    {
        const bool twoFiles     = scenario.shape == Shape::DeletePermanently || scenario.shape == Shape::KeepSources || scenario.shape == Shape::CancelQuestion;
        const auto caseRoot     = root / std::format(L"case-{}", caseIndex++);
        const auto sourceFolder = caseRoot / L"source";
        const auto destinationFolder = caseRoot / L"destination";
        if (! RecreateEmptyDirectory(sourceFolder) || ! RecreateEmptyDirectory(destinationFolder))
            return false;
        // Same size and same date, so a Delete source given for all also covers the second file.
        std::vector<std::filesystem::path> sources;
        const size_t fileCount = twoFiles ? 2u : 1u;
        for (size_t index = 0u; index < fileCount; ++index)
        {
            const auto source      = sourceFolder / std::format(L"reel-{:02}.bin", index);
            const auto destination = destinationFolder / source.filename();
            if (! WriteFileTextFsIo(io, source, sourceBytes) || ! WriteFileTextFsIo(io, destination, destinationBytes) || ! setLastWriteTime(source, kTicks) ||
                ! setLastWriteTime(destination, kTicks))
                return false;
            sources.push_back(source);
        }
        const auto firstDestination = destinationFolder / sources.front().filename();

        RedSalamander::TestSupport::ScopedEnvironmentVariable recycleUnavailable(kRecycleUnavailableDirectory, sourceFolder.native());
        FileOperations::QualifiedEndpoint endpoint{
            .pluginId               = L"builtin/file-system",
            .instanceId             = L"host/default",
            .profileId              = L"local",
            .rootId                 = L"local-test-volume",
            .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };
        endpoint.verificationHostReadback = true;
        FileOperations::TransferPlan transfer{};
        transfer.intent              = FileOperations::TransferIntent::Move;
        transfer.strategy            = FileOperations::OperationStrategy::Managed;
        transfer.sourceEndpoint      = endpoint;
        transfer.destinationEndpoint = endpoint;
        for (const auto& source : sources)
            transfer.selectedItems.emplace_back().providerPath = source.native();
        transfer.destination.providerFolderPath = destinationFolder.native();
        const auto plans                        = std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}});
        State::Task task(fileOperations);
        task._operation             = FILESYSTEM_MOVE;
        task._executionMode         = State::ExecutionMode::PerItem;
        task._fileSystem            = localFileSystem;
        task._sourcePluginId        = endpoint.pluginId;
        task._sourcePluginShortId   = L"file";
        task._sourceInstanceContext = endpoint.instanceId;
        task._perItemMaxConcurrency = 1u;
        for (const auto& source : sources)
        {
            task._sourcePaths.push_back(source);
            task._sourcePathAttributesHint.push_back(GetFileAttributesW(source.c_str()));
        }
        task.InitializeAdmissionInputs(plans,
                                       {.fileSystem      = localFileSystem,
                                        .folder          = destinationFolder,
                                        .pluginId        = endpoint.pluginId,
                                        .pluginShortId   = L"file",
                                        .instanceContext = endpoint.instanceId});
        task.InitializeSourceItemResultBuilders();
        const auto diagnosticsCleanup = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });
        if (FAILED(task.PrepareMutationInterlockScopes()))
            return false;

        // The executor blocks on each decision: answer from a second thread. The collision is answered Delete source (for
        // all when there are two files); the question is answered as the case says, after the case's change, if any. An
        // answer that is not offered is a failure, answered Cancel so the case cannot hang.
        std::atomic<unsigned int> collisionAnswers{0u};
        std::atomic<unsigned int> questionAnswers{0u};
        std::atomic<bool> answersOffered{true};
        std::atomic<bool> questionShape{true};
        std::atomic<bool> intervened{false};
        HRESULT executeHr = E_PENDING;
        {
            const std::jthread responder([&](const std::stop_token stop) noexcept
            {
                const auto listed = [](const auto& actions, size_t count, const Action action) noexcept
                {
                    const auto end = actions.begin() + static_cast<std::ptrdiff_t>(std::min(count, actions.size()));
                    return std::find(actions.begin(), end, action) != end;
                };
                uint64_t answeredGeneration = 0u;
                while (! stop.stop_requested())
                {
                    if (const auto prompt = TryGetConflictPromptCopy(&task);
                        prompt.has_value() && prompt->generation != answeredGeneration && ! prompt->metadataLoading && prompt->buttonsPublishable)
                    {
                        const bool question = prompt->deferredConsent && prompt->bucket == FileOperations::ConflictClass::RecycleFailed;
                        Action answer       = Action::DiscardSource;
                        bool applyToAll     = twoFiles;
                        if (question)
                        {
                            // Permanently delete, Keep source and Cancel; Cancel is the default and the Escape action, and
                            // there is no task-wide button and no Apply to all box: the one answer already covers the task.
                            if (! listed(prompt->actions, prompt->actionCount, Action::PermanentDelete) ||
                                ! listed(prompt->actions, prompt->actionCount, Action::Skip) ||
                                ! listed(prompt->actions, prompt->actionCount, Action::Cancel) ||
                                listed(prompt->actions, prompt->actionCount, Action::PermanentDeleteAll) || prompt->applyToAllEligible ||
                                prompt->defaultAction != Action::Cancel || prompt->escapeAction != Action::Cancel)
                                questionShape.store(false, std::memory_order_release);
                            if (scenario.shape == Shape::SourceChanged && ! intervened.load(std::memory_order_acquire))
                            {
                                const wil::unique_hfile rewritten(
                                    CreateFileW(sources.front().c_str(), GENERIC_WRITE, 0u, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
                                DWORD written = 0u;
                                intervened.store(
                                    rewritten &&
                                        WriteFile(rewritten.get(), changedBytes.data(), static_cast<DWORD>(changedBytes.size()), &written, nullptr) != FALSE &&
                                        written == changedBytes.size(),
                                    std::memory_order_release);
                            }
                            else if (scenario.shape == Shape::DestinationIsSource && ! intervened.load(std::memory_order_acquire))
                            {
                                // Both names lie inside this case's sandbox folder; the guarded helper checks it.
                                std::error_code linkError;
                                intervened.store(DeleteFileW(firstDestination.c_str()) != FALSE &&
                                                     Common::Testing::CreateSandboxedHardLink(firstDestination, sources.front(), linkError),
                                                 std::memory_order_release);
                            }
                            answer     = scenario.shape == Shape::KeepSources      ? Action::Skip
                                         : scenario.shape == Shape::CancelQuestion ? Action::Cancel
                                                                                   : Action::PermanentDelete;
                            applyToAll = false;
                        }
                        if (! listed(prompt->actions, prompt->actionCount, answer))
                        {
                            answersOffered.store(false, std::memory_order_release);
                            answer     = Action::Cancel;
                            applyToAll = false;
                        }
                        if (task.SubmitConflictDecision(prompt->generation, answer, applyToAll))
                        {
                            answeredGeneration = prompt->generation;
                            (question ? questionAnswers : collisionAnswers).fetch_add(1u, std::memory_order_acq_rel);
                        }
                    }
                    Sleep(2);
                }
            });
            executeHr = task.ExecuteOperation();
        }
        const HRESULT finalHr = task.FinalizeTypedItemResults(executeHr);
        task._resultHr.store(finalHr, std::memory_order_release);
        const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);
        Action remembered  = Action::None;
        {
            const std::scoped_lock lock(task._conflictArbiter.mutex);
            remembered = task._conflictArbiter.sourceNoRecycleBinAnswer;
        }

        bool outcome = false;
        switch (scenario.shape)
        {
            case Shape::DeletePermanently:
            {
                bool all = true;
                for (const auto& source : sources)
                {
                    std::error_code error;
                    all = all && ! std::filesystem::exists(source, error) && ! error && holds(destinationFolder / source.filename(), destinationBytes);
                }
                outcome = all && remembered == Action::PermanentDelete && summary.alreadyThereSourceRemovedCount == 2u && summary.removedSourceCount == 2u &&
                          summary.publishedItemCount == 0u && summary.failedItemCount == 0u && finalHr == S_OK;
                break;
            }
            case Shape::KeepSources:
            {
                bool all = true;
                for (const auto& source : sources)
                    all = all && holds(source, sourceBytes) && holds(destinationFolder / source.filename(), destinationBytes);
                outcome = all && remembered == Action::Skip && summary.alreadyThereSourceRemovedCount == 0u && summary.removedSourceCount == 0u &&
                          summary.failedItemCount == 0u && executeHr != HRESULT_FROM_WIN32(ERROR_CANCELLED);
                break;
            }
            case Shape::SourceChanged:
                outcome = intervened.load(std::memory_order_acquire) && holds(sources.front(), changedBytes) && holds(firstDestination, destinationBytes) &&
                          remembered == Action::PermanentDelete && summary.alreadyThereSourceRemovedCount == 0u && summary.removedSourceCount == 0u &&
                          summary.failedItemCount == 0u && executeHr != HRESULT_FROM_WIN32(ERROR_CANCELLED);
                break;
            case Shape::DestinationIsSource:
                outcome = intervened.load(std::memory_order_acquire) && holds(sources.front(), sourceBytes) && holds(firstDestination, sourceBytes) &&
                          remembered == Action::PermanentDelete && summary.alreadyThereSourceRemovedCount == 0u && summary.removedSourceCount == 0u &&
                          summary.failedItemCount == 0u && executeHr != HRESULT_FROM_WIN32(ERROR_CANCELLED);
                break;
            case Shape::CancelQuestion:
            {
                // Cancel is the task's answer: it is remembered, the task itself is canceled (not only this item), the
                // second file never asks again, and nothing is removed.
                bool all = true;
                for (const auto& source : sources)
                    all = all && holds(source, sourceBytes) && holds(destinationFolder / source.filename(), destinationBytes);
                outcome = all && remembered == Action::Cancel && task._cancelled.load(std::memory_order_acquire) &&
                          executeHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && summary.alreadyThereSourceRemovedCount == 0u && summary.removedSourceCount == 0u;
                break;
            }
        }
        // One collision decision (the second file takes the remembered Delete source) and one question per task.
        const bool prompts = answersOffered.load(std::memory_order_acquire) && questionShape.load(std::memory_order_acquire) &&
                             questionAnswers.load(std::memory_order_acquire) == 1u && collisionAnswers.load(std::memory_order_acquire) == 1u &&
                             task._perf.consentPromptCount == 1u && task._perf.conflictPromptCount == 1u;
        const bool matches = prompts && outcome;
        AppendLog(std::format(L"Delete source without a Recycle Bin {}: execute=0x{:08X} final=0x{:08X} collisions={} questions={} consents={} offered={} "
                              L"questionShape={} intervened={} remembered={} alreadyThere={} removed={} retained={} skipped={} failed={} passed={}",
                              scenario.name,
                              static_cast<unsigned long>(executeHr),
                              static_cast<unsigned long>(finalHr),
                              collisionAnswers.load(std::memory_order_acquire),
                              questionAnswers.load(std::memory_order_acquire),
                              static_cast<uint64_t>(task._perf.consentPromptCount),
                              answersOffered.load(std::memory_order_acquire),
                              questionShape.load(std::memory_order_acquire),
                              intervened.load(std::memory_order_acquire),
                              static_cast<unsigned int>(remembered),
                              summary.alreadyThereSourceRemovedCount,
                              summary.removedSourceCount,
                              summary.retainedSourceCount,
                              summary.skippedItemCount,
                              summary.failedItemCount,
                              matches));
        passed = matches && passed;
    }
    return passed;
}

// I63 B6.5 (D-13): the collision policy a task may carry from its start. A Copy of four files onto a folder that already
// holds each of them: one pair with the same size and date, one whose existing file is older, one whose existing file is
// newer and one of another size. Ask asks four times (each answered Skip) and keeps both items' facts on the Issues
// records (B3.3); Skip and Keep both decide every pair without a question; Replace older replaces the one older file and
// asks for the three others. The value is seeded before the task runs, as admission does.
[[nodiscard]] bool RunFileOperationPresetCollisionPolicySelfTest(FolderWindow::FileOperationState& fileOperations,
                                                                 const wil::com_ptr<IFileSystem>& localFileSystem,
                                                                 const std::filesystem::path& testRoot) noexcept
{
    using State  = FolderWindow::FileOperationState;
    using Action = State::Task::ConflictAction;
    using Preset = FileOperations::PresetCollisionPolicy;
    wil::com_ptr<IFileSystemIO> io;
    const auto root = testRoot / L"preset-collision-policy";
    if (! localFileSystem || ! localFileSystem.try_query_to(io.put()) || ! RecreateEmptyDirectory(root))
        return false;
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });

    enum class CollisionPair : uint8_t
    {
        SameSizeAndDate,
        DestinationOlder, // same size; the incoming file is newer
        DestinationNewer, // same size; the incoming file is older
        OtherSize,        // same date; the existing file is larger
    };
    static constexpr std::array kPairs{
        CollisionPair::SameSizeAndDate, CollisionPair::DestinationOlder, CollisionPair::DestinationNewer, CollisionPair::OtherSize};
    struct Scenario final
    {
        const wchar_t* name = L"";
        Preset preset       = Preset::Ask;
    };
    static constexpr std::array kScenarios{
        Scenario{L"ask", Preset::Ask},
        Scenario{L"skip", Preset::Skip},
        Scenario{L"replace-older", Preset::ReplaceOlder},
        Scenario{L"keep-both", Preset::KeepBoth},
    };
    constexpr uint64_t kBaseTicks = 133'500'000'000'000'000ull;
    constexpr uint64_t kDayTicks  = 864'000'000'000ull;
    const auto setLastWriteTime   = [](const std::filesystem::path& path, uint64_t ticks) noexcept
    {
        const wil::unique_hfile file(CreateFileW(path.c_str(),
                                                 FILE_WRITE_ATTRIBUTES,
                                                 FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                 nullptr,
                                                 OPEN_EXISTING,
                                                 FILE_ATTRIBUTE_NORMAL,
                                                 nullptr));
        FILETIME time{};
        time.dwLowDateTime  = static_cast<DWORD>(ticks & 0xFFFF'FFFFull);
        time.dwHighDateTime = static_cast<DWORD>(ticks >> 32u);
        return file.is_valid() && SetFileTime(file.get(), nullptr, nullptr, &time) != FALSE;
    };
    const std::string sourceBytes(64u, 's');
    const std::string sameSizeBytes(64u, 'd');
    const std::string otherSizeBytes(96u, 'd');

    bool passed            = true;
    unsigned int caseIndex = 0u;
    for (const Scenario& scenario : kScenarios)
    {
        const auto caseRoot          = root / std::format(L"case-{}", caseIndex++);
        const auto sourceFolder      = caseRoot / L"source";
        const auto destinationFolder = caseRoot / L"destination";
        if (! RecreateEmptyDirectory(sourceFolder) || ! RecreateEmptyDirectory(destinationFolder))
            return false;
        std::vector<std::filesystem::path> sources;
        for (size_t index = 0u; index < kPairs.size(); ++index)
        {
            const CollisionPair pair        = kPairs[index];
            const auto source               = sourceFolder / std::format(L"track-{:02}.bin", index);
            const auto destination          = destinationFolder / source.filename();
            const uint64_t destinationTicks = pair == CollisionPair::DestinationOlder   ? kBaseTicks - kDayTicks
                                              : pair == CollisionPair::DestinationNewer ? kBaseTicks + kDayTicks
                                                                                        : kBaseTicks;
            if (! WriteFileTextFsIo(io, source, sourceBytes) ||
                ! WriteFileTextFsIo(io, destination, pair == CollisionPair::OtherSize ? otherSizeBytes : sameSizeBytes) ||
                ! setLastWriteTime(source, kBaseTicks) || ! setLastWriteTime(destination, destinationTicks))
                return false;
            sources.push_back(source);
        }

        FileOperations::QualifiedEndpoint endpoint{
            .pluginId               = L"builtin/file-system",
            .instanceId             = L"host/default",
            .profileId              = L"local",
            .rootId                 = L"local-test-volume",
            .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };
        endpoint.verificationHostReadback = true;
        FileOperations::TransferPlan transfer{};
        transfer.intent              = FileOperations::TransferIntent::Copy;
        transfer.strategy            = FileOperations::OperationStrategy::Copy;
        transfer.sourceEndpoint      = endpoint;
        transfer.destinationEndpoint = endpoint;
        for (const auto& source : sources)
            transfer.selectedItems.emplace_back().providerPath = source.native();
        transfer.destination.providerFolderPath = destinationFolder.native();
        const auto plans                        = std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}});
        State::Task task(fileOperations);
        task._operation             = FILESYSTEM_COPY;
        task._executionMode         = State::ExecutionMode::PerItem;
        task._fileSystem            = localFileSystem;
        task._sourcePluginId        = endpoint.pluginId;
        task._sourcePluginShortId   = L"file";
        task._sourceInstanceContext = endpoint.instanceId;
        task._perItemMaxConcurrency = 1u;
        for (const auto& source : sources)
        {
            task._sourcePaths.push_back(source);
            task._sourcePathAttributesHint.push_back(GetFileAttributesW(source.c_str()));
        }
        task.InitializeAdmissionInputs(plans,
                                       {.fileSystem      = localFileSystem,
                                        .folder          = destinationFolder,
                                        .pluginId        = endpoint.pluginId,
                                        .pluginShortId   = L"file",
                                        .instanceContext = endpoint.instanceId});
        task.InitializeSourceItemResultBuilders();
        task.SeedPresetCollisionPolicy(scenario.preset);
        const auto diagnosticsCleanup = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });
        if (FAILED(task.PrepareMutationInterlockScopes()))
            return false;

        // Every question that is still asked is answered Skip; a question without Skip is answered Cancel so the case
        // cannot hang.
        std::atomic<unsigned int> answered{0u};
        std::atomic<bool> skipOffered{true};
        HRESULT executeHr = E_PENDING;
        {
            const std::jthread responder([&](const std::stop_token stop) noexcept
            {
                uint64_t answeredGeneration = 0u;
                while (! stop.stop_requested())
                {
                    if (const auto prompt = TryGetConflictPromptCopy(&task);
                        prompt.has_value() && prompt->generation != answeredGeneration && ! prompt->metadataLoading && prompt->buttonsPublishable)
                    {
                        const auto offeredEnd = prompt->actions.begin() + static_cast<std::ptrdiff_t>(prompt->actionCount);
                        const bool offered    = std::find(prompt->actions.begin(), offeredEnd, Action::Skip) != offeredEnd;
                        if (! offered)
                            skipOffered.store(false, std::memory_order_release);
                        if (task.SubmitConflictDecision(prompt->generation, offered ? Action::Skip : Action::Cancel, false))
                        {
                            answeredGeneration = prompt->generation;
                            answered.fetch_add(1u, std::memory_order_acq_rel);
                        }
                    }
                    Sleep(2);
                }
            });
            executeHr = task.ExecuteOperation();
        }
        const HRESULT finalHr = task.FinalizeTypedItemResults(executeHr);
        task._resultHr.store(finalHr, std::memory_order_release);
        const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);

        size_t expectedPrompts  = 0u;
        size_t expectedReplaced = 0u;
        size_t expectedSkipped  = 0u;
        size_t expectedKeptBoth = 0u;
        bool bytesMatch         = true;
        for (size_t index = 0u; index < kPairs.size(); ++index)
        {
            const CollisionPair pair = kPairs[index];
            const bool replaced      = scenario.preset == Preset::ReplaceOlder && pair == CollisionPair::DestinationOlder;
            const bool keptBoth      = scenario.preset == Preset::KeepBoth;
            const bool asked         = scenario.preset == Preset::Ask || (scenario.preset == Preset::ReplaceOlder && ! replaced);
            expectedPrompts += asked ? 1u : 0u;
            expectedReplaced += replaced ? 1u : 0u;
            expectedKeptBoth += keptBoth ? 1u : 0u;
            expectedSkipped += (! replaced && ! keptBoth) ? 1u : 0u;
            std::string actualSource;
            std::string actualDestination;
            const std::string& destinationBefore = pair == CollisionPair::OtherSize ? otherSizeBytes : sameSizeBytes;
            bytesMatch                           = bytesMatch && ReadFileTextFsIo(io, sources[index], actualSource) && actualSource == sourceBytes &&
                                                   ReadFileTextFsIo(io, destinationFolder / sources[index].filename(), actualDestination) &&
                                                   actualDestination == (replaced ? sourceBytes : destinationBefore);
        }
        size_t destinationEntries = 0u;
        size_t keptBothEntries    = 0u;
        size_t leftoverEntries    = 0u;
        std::error_code error;
        for (std::filesystem::directory_iterator it(destinationFolder, error), end; ! error && it != end; it.increment(error))
        {
            ++destinationEntries;
            const std::wstring leaf = it->path().filename().native();
            if (FileOperationArtifacts::HasPossibleArtifactName(leaf))
            {
                ++leftoverEntries;
                continue;
            }
            const bool original =
                std::ranges::any_of(sources, [&leaf](const std::filesystem::path& source) noexcept { return source.filename().native() == leaf; });
            std::string bytes;
            if (! original && ReadFileTextFsIo(io, it->path(), bytes) && bytes == sourceBytes)
                ++keptBothEntries;
        }
        // I63 B3.3: an asked collision keeps both items' facts on its Issues records (the question and the answer).
        size_t factRecords = 0u;
        for (const auto& entry : summary.issueDiagnostics)
        {
            if (! entry.conflictPrompt.has_value() || ! entry.conflictPrompt->itemFacts.has_value())
                continue;
            const auto& observed = entry.conflictPrompt->itemFacts.value();
            if (observed.source.sizeKnown && observed.source.sizeBytes == sourceBytes.size() && observed.destination.sizeKnown &&
                observed.source.lastWriteTime == static_cast<__int64>(kBaseTicks))
                ++factRecords;
        }
        const bool seeded  = task._conflictArbiter.presetPolicy == scenario.preset &&
                             (task._conflictArbiter.replaceOlderGrantGeneration == State::Task::ConflictArbiter::kPresetPolicyGeneration) ==
                                 (scenario.preset == Preset::ReplaceOlder);
        const bool prompts = skipOffered.load(std::memory_order_acquire) && answered.load(std::memory_order_acquire) == expectedPrompts &&
                             task._perf.conflictPromptCount == expectedPrompts;
        const bool counts  = summary.skippedItemCount == expectedSkipped && summary.publishedItemCount == expectedReplaced + expectedKeptBoth &&
                             summary.failedItemCount == 0u && summary.skippedSameSizeAndDateCount == 0u && summary.replacedOlderCount == expectedReplaced;
        const bool entries = destinationEntries == kPairs.size() + expectedKeptBoth && keptBothEntries == expectedKeptBoth && leftoverEntries == 0u;
        const bool facts   = scenario.preset != Preset::Ask || factRecords >= kPairs.size();
        const bool matches = seeded && prompts && counts && entries && facts && bytesMatch && executeHr != HRESULT_FROM_WIN32(ERROR_CANCELLED);
        AppendLog(std::format(L"Preset collision policy {}: execute=0x{:08X} final=0x{:08X} prompts={}/{} skipOffered={} skipped={}/{} published={}/{} "
                              L"replacedOlder={} failed={} entries={} keptBoth={} leftovers={} factRecords={} seeded={} bytes={} passed={}",
                              scenario.name,
                              static_cast<unsigned long>(executeHr),
                              static_cast<unsigned long>(finalHr),
                              answered.load(std::memory_order_acquire),
                              expectedPrompts,
                              skipOffered.load(std::memory_order_acquire),
                              summary.skippedItemCount,
                              expectedSkipped,
                              summary.publishedItemCount,
                              expectedReplaced + expectedKeptBoth,
                              summary.replacedOlderCount,
                              summary.failedItemCount,
                              destinationEntries,
                              keptBothEntries,
                              leftoverEntries,
                              factRecords,
                              seeded,
                              bytesMatch,
                              matches));
        passed = matches && passed;
    }
    return passed;
}

// I63 B3.3: the retained comparison renders the decision's own lines, and the Issues projection keeps them on the row its
// question and its answer share, in either arrival order.
[[nodiscard]] bool RunFileOperationConflictObservedFactsSelfTest() noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Collision facts on the Issues row: {}", reason));
        return false;
    };
    constexpr __int64 kNewer = 133'500'000'000'000'000ll;
    constexpr __int64 kOlder = 133'400'000'000'000'000ll;
    FileOperations::ConflictObservedComparison observed{};
    observed.source      = {.sizeKnown = true, .sizeBytes = 2'048u, .lastWriteTime = kNewer};
    observed.destination = {.sizeKnown = true, .sizeBytes = 1'024u, .lastWriteTime = kOlder};
    observed.comparison  = FileOperations::CompareConflictItems(true, 2'048u, kNewer, true, 1'024u, kOlder);
    FileOperations::ConflictPromptTextFacts facts{};
    facts.operation       = FILESYSTEM_COPY;
    facts.bucket          = static_cast<uint8_t>(FileOperations::ConflictClass::RegularFileExists);
    facts.sourceItem      = {.available = true, .isDirectory = false, .sizeKnown = true, .sizeBytes = 2'048u, .lastWriteTime = kNewer};
    facts.destinationItem = {.available = true, .isDirectory = false, .sizeKnown = true, .sizeBytes = 1'024u, .lastWriteTime = kOlder};
    facts.dateRelation    = static_cast<uint8_t>(observed.comparison.date);
    facts.sizeRelation    = static_cast<uint8_t>(observed.comparison.size);
    const FileOperations::ConflictObservedTexts differ = FileOperations::BuildConflictObservedTexts(FILESYSTEM_COPY, observed);
    if (differ.source.empty() || differ.destination.empty() || differ.source != FileOperations::BuildConflictPromptItemText(facts, false) ||
        differ.destination != FileOperations::BuildConflictPromptItemText(facts, true) || ! differ.summary.empty())
        return reject(L"the role lines differ from the decision's");
    observed.destination                             = observed.source;
    observed.comparison                              = FileOperations::CompareConflictItems(true, 2'048u, kNewer, true, 2'048u, kNewer);
    const FileOperations::ConflictObservedTexts same = FileOperations::BuildConflictObservedTexts(FILESYSTEM_COPY, observed);
    if (same.summary != LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_CMP_SAME_SIZE_AND_DATE))
        return reject(L"equal size and date did not read as one summary line");

    using Diagnostic = FolderWindow::FileOperationState::TaskDiagnosticEntry;
    Diagnostic question{};
    question.localTime                 = SYSTEMTIME{2026, 11, 3, 11, 9, 30, 0, 0};
    question.taskId                    = 31u;
    question.operation                 = FILESYSTEM_COPY;
    question.severity                  = FolderWindow::FileOperationState::DiagnosticSeverity::Info;
    question.status                    = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    question.category                  = L"item.conflict.prompt";
    question.message                   = L"Conflict prompt shown for item.";
    question.conflictPrompt            = FileOperations::ConflictPromptObservation{9u, FileOperations::ConflictClass::RegularFileExists, 0u};
    question.conflictPrompt->itemFacts = observed;
    Diagnostic answer                  = question;
    answer.category                    = L"item.conflict.accepted";
    answer.acceptedDecision            = FileOperations::AcceptedConflictDecision{.taskId           = question.taskId,
                                                                                  .promptGeneration = 9u,
                                                                                  .sourceItemIndex  = 0u,
                                                                                  .action           = FileOperations::ConflictAction::Skip,
                                                                                  .scope            = {.bucket = FileOperations::ConflictClass::RegularFileExists}};
    FileOperationsIssues::Text labels{};
    labels.infoSeverity      = L"Information";
    labels.unknownDecision   = L"Unknown decision";
    labels.automaticDecision = L"Automatic";
    labels.unknownResult     = L"Unknown result";
    labels.completedResult   = L"Completed";
    labels.skippedResult     = L"Skipped";
    labels.canceledResult    = L"Canceled";
    labels.failedResult      = L"Failed";
    for (const bool answerFirst : {false, true})
    {
        std::array<FileOperationsIssues::Observation, 2u> observations{};
        for (auto& observation : observations)
        {
            observation.timeText      = L"2026-11-11 09:30:00.000";
            observation.operationText = L"Copy";
            observation.statusText    = L"0x80070050";
        }
        auto& questionObservation                = observations[answerFirst ? 1u : 0u];
        auto& answerObservation                  = observations[answerFirst ? 0u : 1u];
        questionObservation.issue                = &question;
        questionObservation.sourceFactsText      = same.source;
        questionObservation.destinationFactsText = same.destination;
        questionObservation.comparisonText       = same.summary;
        answerObservation.issue                  = &answer;
        answerObservation.formattedDecisionText  = L"Skip";
        FileOperationsIssues::Projection projection;
        if (! projection.Update(observations, labels, 1u) || projection.Rows().size() != 1u)
            return reject(L"the question and its answer did not join one row");
        const auto& row = projection.Rows().front();
        if (row.sourceFactsText != same.source || row.destinationFactsText != same.destination || row.comparisonText != same.summary || ! row.hasConflictFacts)
            return reject(answerFirst ? L"the row lost the facts when the answer arrived first" : L"the row lost the facts when the answer joined it");
    }
    return true;
}

// I63 B1.5: a folder onto an existing folder at a location that opens items but cannot identify them. The two folders
// merge, as at a location that identifies its items: the folder itself raises no decision, and each child that
// collides raises its own (answered Skip here). A Copy merges and copies the new child; a Move merges too, keeps a
// source folder that still holds a skipped child (not an Error) and removes one the merge emptied. A destination
// folder that is a link stays refused before anything is touched.
[[nodiscard]] bool RunFileOperationIdentityLessFolderMergeSelfTest(FolderWindow::FileOperationState& fileOperations,
                                                                   const wil::com_ptr<IFileSystem>& localFileSystem,
                                                                   const std::filesystem::path& testRoot) noexcept
{
    using State  = FolderWindow::FileOperationState;
    using Action = State::Task::ConflictAction;
    wil::com_ptr<IFileSystemIO> io;
    const auto root = testRoot / L"identity-less-folder-merge";
    if (! localFileSystem || ! localFileSystem.try_query_to(io.put()) || ! RecreateEmptyDirectory(root))
        return false;
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    enum class Shape : uint8_t
    {
        CopyMergeSkip, // Copy: one child collides and is skipped, the other child is new
        MoveMergeSkip, // the same as a Move: the source folder keeps the skipped child
        MoveMergeNew,  // Move: no child collides, so the merge empties the source folder
        MoveOntoLink,  // Move: the existing destination folder is a link
    };
    struct Scenario final
    {
        const wchar_t* name = L"";
        Shape shape         = Shape::CopyMergeSkip;
    };
    static constexpr std::array kScenarios{
        Scenario{L"copy-merge-skip", Shape::CopyMergeSkip},
        Scenario{L"move-merge-skip", Shape::MoveMergeSkip},
        Scenario{L"move-merge-new", Shape::MoveMergeNew},
        Scenario{L"move-onto-link", Shape::MoveOntoLink},
    };
    const std::string sourceBytes(73u, 's');
    const std::string existingBytes(149u, 'd');
    const std::filesystem::path firstTrack(L"track-01.bin");
    const std::filesystem::path secondTrack(L"track-02.bin");
    const std::filesystem::path cover(L"cover.bin");
    const auto holds = [&io](const std::filesystem::path& path, const std::string& expected) noexcept
    {
        std::string actual;
        return ReadFileTextFsIo(io, path, actual) && actual == expected;
    };
    const auto absent = [](const std::filesystem::path& path) noexcept
    {
        std::error_code existsError;
        return ! std::filesystem::exists(path, existsError) && ! existsError;
    };
    const auto countEntries = [](const std::filesystem::path& folder, size_t& artifactEntries) noexcept
    {
        size_t listed   = 0u;
        artifactEntries = 0u;
        std::error_code listError;
        for (std::filesystem::directory_iterator it(folder, listError), end; ! listError && it != end; it.increment(listError))
        {
            ++listed;
            if (FileOperationArtifacts::HasPossibleArtifactName(it->path().filename().native()))
                ++artifactEntries;
        }
        return listed;
    };

    bool passed            = true;
    unsigned int caseIndex = 0u;
    for (const Scenario& scenario : kScenarios)
    {
        const bool move              = scenario.shape != Shape::CopyMergeSkip;
        const bool collides          = scenario.shape == Shape::CopyMergeSkip || scenario.shape == Shape::MoveMergeSkip;
        const bool ontoLink          = scenario.shape == Shape::MoveOntoLink;
        const auto caseRoot          = root / std::format(L"case-{}", caseIndex++);
        const auto sourceFolder      = caseRoot / L"source" / L"album";
        const auto destinationFolder = caseRoot / L"destination";
        const auto mergedFolder      = destinationFolder / sourceFolder.filename();
        const auto linkTarget        = caseRoot / L"elsewhere";
        if (! RecreateEmptyDirectory(sourceFolder) || ! RecreateEmptyDirectory(destinationFolder) ||
            ! WriteFileTextFsIo(io, sourceFolder / firstTrack, sourceBytes) || ! WriteFileTextFsIo(io, sourceFolder / secondTrack, sourceBytes))
            return false;
        if (ontoLink)
        {
            // The junction and its target both lie inside this case's sandbox folder.
            if (! RecreateEmptyDirectory(linkTarget) || ! TryCreateJunction(mergedFolder, linkTarget))
                return false;
        }
        else if (! RecreateEmptyDirectory(mergedFolder) || ! WriteFileTextFsIo(io, mergedFolder / (collides ? firstTrack : cover), existingBytes))
        {
            return false;
        }
        // The whole destination folder cannot identify its items (a share without file IDs); the source can.
        RedSalamander::TestSupport::ScopedEnvironmentVariable identitySupport(kSelfTestEnvBindIdentityUnsupportedPath, destinationFolder.native());
        FileOperations::QualifiedEndpoint endpoint{
            .pluginId               = L"builtin/file-system",
            .instanceId             = L"host/default",
            .profileId              = L"local",
            .rootId                 = L"local-test-volume",
            .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };
        // The Local destination proves what it publishes; the merge route requires that proof.
        endpoint.verificationHostReadback = true;
        FileOperations::TransferPlan transfer{};
        transfer.intent                         = move ? FileOperations::TransferIntent::Move : FileOperations::TransferIntent::Copy;
        transfer.strategy                       = move ? FileOperations::OperationStrategy::Managed : FileOperations::OperationStrategy::Copy;
        transfer.sourceEndpoint                 = endpoint;
        transfer.destinationEndpoint            = endpoint;
        transfer.selectedItems                  = {{.providerPath = sourceFolder.native()}};
        transfer.destination.providerFolderPath = destinationFolder.native();
        const auto plans                        = std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}});
        State::Task task(fileOperations);
        task._operation                = move ? FILESYSTEM_MOVE : FILESYSTEM_COPY;
        task._executionMode            = State::ExecutionMode::PerItem;
        task._fileSystem               = localFileSystem;
        task._sourcePaths              = {sourceFolder};
        task._sourcePluginId           = endpoint.pluginId;
        task._sourcePluginShortId      = L"file";
        task._sourceInstanceContext    = endpoint.instanceId;
        task._sourcePathAttributesHint = {GetFileAttributesW(sourceFolder.c_str())};
        task._perItemMaxConcurrency    = 1u;
        task.InitializeAdmissionInputs(plans,
                                       {.fileSystem      = localFileSystem,
                                        .folder          = destinationFolder,
                                        .pluginId        = endpoint.pluginId,
                                        .pluginShortId   = L"file",
                                        .instanceContext = endpoint.instanceId});
        task.InitializeSourceItemResultBuilders();
        const auto diagnosticsCleanup = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });
        if (FAILED(task.PrepareMutationInterlockScopes()))
            return false;

        // The executor blocks on each decision: answer from a second thread. Every decision is a colliding child's and
        // is answered Skip; a decision that does not offer Skip is answered Cancel so the case cannot hang.
        std::atomic<unsigned int> answered{0u};
        std::atomic<bool> skipOffered{true};
        HRESULT executeHr = E_PENDING;
        {
            const std::jthread responder([&](const std::stop_token stop) noexcept
            {
                uint64_t answeredGeneration = 0u;
                while (! stop.stop_requested())
                {
                    if (const auto prompt = TryGetConflictPromptCopy(&task);
                        prompt.has_value() && prompt->generation != answeredGeneration && ! prompt->metadataLoading && prompt->buttonsPublishable)
                    {
                        const auto offeredEnd = prompt->actions.begin() + static_cast<std::ptrdiff_t>(prompt->actionCount);
                        const bool offered    = std::find(prompt->actions.begin(), offeredEnd, Action::Skip) != offeredEnd;
                        if (! offered)
                            skipOffered.store(false, std::memory_order_release);
                        if (task.SubmitConflictDecision(prompt->generation, offered ? Action::Skip : Action::Cancel, false))
                        {
                            answeredGeneration = prompt->generation;
                            answered.fetch_add(1u, std::memory_order_acq_rel);
                        }
                    }
                    Sleep(2);
                }
            });
            executeHr = task.ExecuteOperation();
        }
        const HRESULT finalHr = task.FinalizeTypedItemResults(executeHr);
        task._resultHr.store(finalHr, std::memory_order_release);
        const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);
        const auto& item   = task._sourceItemResultBuilders.front().terminal;

        size_t mergedArtifacts      = 0u;
        const size_t mergedEntries  = countEntries(ontoLink ? linkTarget : mergedFolder, mergedArtifacts);
        const unsigned int prompts  = answered.load(std::memory_order_acquire);
        const bool sourceFirstKept  = holds(sourceFolder / firstTrack, sourceBytes);
        const bool sourceSecondKept = holds(sourceFolder / secondTrack, sourceBytes);
        bool outcome                = false;
        switch (scenario.shape)
        {
            case Shape::CopyMergeSkip:
                // Merged: the existing child keeps its bytes, the new child is copied, the source is untouched.
                outcome = prompts == 1u && sourceFirstKept && sourceSecondKept && holds(mergedFolder / firstTrack, existingBytes) &&
                          holds(mergedFolder / secondTrack, sourceBytes) && mergedEntries == 2u && summary.failedItemCount == 0u;
                break;
            case Shape::MoveMergeSkip:
                // Merged: the new child moved; the skipped child and its source folder stay, and the item is no Error.
                outcome = prompts == 1u && sourceFirstKept && absent(sourceFolder / secondTrack) && holds(mergedFolder / firstTrack, existingBytes) &&
                          holds(mergedFolder / secondTrack, sourceBytes) && mergedEntries == 2u && summary.failedItemCount == 0u && item.has_value() &&
                          item->completion != FileOperations::ItemCompletion::Failed && item->sourceDisposition == FileOperations::SourceDisposition::Retained;
                break;
            case Shape::MoveMergeNew:
                // Merged without a question: both children moved and the emptied source folder is gone.
                outcome = prompts == 0u && absent(sourceFolder) && holds(mergedFolder / cover, existingBytes) &&
                          holds(mergedFolder / firstTrack, sourceBytes) && holds(mergedFolder / secondTrack, sourceBytes) && mergedEntries == 3u &&
                          summary.failedItemCount == 0u && summary.removedSourceCount == 1u && SUCCEEDED(finalHr);
                break;
            case Shape::MoveOntoLink:
            {
                // Refused before anything is touched: the link and its empty target stay, the source keeps both children.
                const DWORD linkAttributes = GetFileAttributesW(mergedFolder.c_str());
                outcome = prompts == 0u && executeHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) && sourceFirstKept && sourceSecondKept && mergedEntries == 0u &&
                          linkAttributes != INVALID_FILE_ATTRIBUTES && (linkAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u &&
                          summary.publishedItemCount == 0u && summary.removedSourceCount == 0u;
                break;
            }
        }
        const bool matches =
            outcome && skipOffered.load(std::memory_order_acquire) && mergedArtifacts == 0u && executeHr != HRESULT_FROM_WIN32(ERROR_CANCELLED);
        AppendLog(std::format(L"Identity-less folder merge {}: execute=0x{:08X} final=0x{:08X} prompts={} skipOffered={} sourceFirstKept={} "
                              L"sourceSecondKept={} entries={} artifacts={} published={} removed={} retained={} failed={} passed={}",
                              scenario.name,
                              static_cast<unsigned long>(executeHr),
                              static_cast<unsigned long>(finalHr),
                              prompts,
                              skipOffered.load(std::memory_order_acquire),
                              sourceFirstKept,
                              sourceSecondKept,
                              mergedEntries,
                              mergedArtifacts,
                              summary.publishedItemCount,
                              summary.removedSourceCount,
                              summary.retainedSourceCount,
                              summary.failedItemCount,
                              matches));
        passed = matches && passed;
    }
    return passed;
}

// I63 B7: a Managed Move whose source is at a location that opens items but cannot identify them. The bridge has the
// provider read the source's facts before the copy opens it, and removes the source by name after the destination's
// publication proof only while those facts are unchanged; otherwise the source is kept and the item reports "Copied;
// source kept", which is not an Error. A source folder is removed only when the Move emptied it. A source carrying an
// alternate data stream is kept, because this route does not copy streams. A source whose open fails keeps its
// refusal. Where an existing destination file is replaced, at an ordinary destination or at one that cannot identify
// items either, the same-file probe answers before the source is removed. A collision answered Skip keeps both files.
// A source that is the destination file itself under another name (a junction) is never removed.
[[nodiscard]] bool RunFileOperationIdentityLessSourceMoveSelfTest(FolderWindow::FileOperationState& fileOperations,
                                                                  const wil::com_ptr<IFileSystem>& localFileSystem,
                                                                  const std::filesystem::path& testRoot) noexcept
{
    using State  = FolderWindow::FileOperationState;
    using Action = State::Task::ConflictAction;
    wil::com_ptr<IFileSystemIO> io;
    const auto root = testRoot / L"identity-less-source-move";
    if (! localFileSystem || ! localFileSystem.try_query_to(io.put()) || ! RecreateEmptyDirectory(root))
        return false;
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    // The cleanup pause is process-wide: it never stays armed after this test.
    const auto disarmPause = wil::scope_exit([]() noexcept
    {
        ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
    });
    enum class Shape : uint8_t
    {
        FileUnchanged,           // removed after the publication proof
        FileChanged,             // grows between the copy and its removal: kept
        FileWithStream,          // carries an alternate data stream this route does not copy: kept
        FileOpenFails,           // the location refuses the open: the refusal stays
        FolderEmptied,           // a folder the Move empties: removed
        FolderGainsItem,         // a folder that gains an item before its removal: kept with that item
        ReplaceWithoutIds,       // the destination cannot identify items either and holds another file: Replace, probe, removed
        ReplaceWithoutIdsLong,   // the same corroboration beyond MAX_PATH
        PreauthorizedWithoutIds, // ALLOW_OVERWRITE still obtains observed-facts consent when binding is unsupported
        ReplaceWithIds,          // an ordinary destination holds another file: Replace, probe, removed
        SkipExisting,            // the same collision answered Skip: both files stay
        SameFileTwoNames,        // the source folder is a junction to the destination folder: one file, two names; never removed
        SharingUnenforced,       // Replace where the location does not enforce sharing: the probe cannot answer, so the source is kept
    };
    struct Scenario final
    {
        const wchar_t* name = L"";
        Shape shape         = Shape::FileUnchanged;
    };
    static constexpr std::array kScenarios{
        Scenario{L"file-unchanged", Shape::FileUnchanged},
        Scenario{L"file-changed", Shape::FileChanged},
        Scenario{L"file-with-stream", Shape::FileWithStream},
        Scenario{L"file-open-fails", Shape::FileOpenFails},
        Scenario{L"folder-emptied", Shape::FolderEmptied},
        Scenario{L"folder-gains-item", Shape::FolderGainsItem},
        Scenario{L"replace-without-ids", Shape::ReplaceWithoutIds},
        Scenario{L"replace-without-ids-long", Shape::ReplaceWithoutIdsLong},
        Scenario{L"preauthorized-without-ids", Shape::PreauthorizedWithoutIds},
        Scenario{L"replace-with-ids", Shape::ReplaceWithIds},
        Scenario{L"skip-existing", Shape::SkipExisting},
        Scenario{L"same-file-two-names", Shape::SameFileTwoNames},
        Scenario{L"sharing-unenforced", Shape::SharingUnenforced},
    };
    const std::string sourceBytes(97u, 's');
    const std::string occupantBytes(131u, 'd');
    const std::string appendedBytes(29u, 'a');
    const auto holds = [&io](const std::filesystem::path& path, const std::string& expected) noexcept
    {
        std::string actual;
        return ReadFileTextFsIo(io, path, actual) && actual == expected;
    };
    const auto absent = [](const std::filesystem::path& path) noexcept
    {
        std::error_code existsError;
        return ! std::filesystem::exists(path, existsError) && ! existsError;
    };
    // Another writer appends to the file: its size and last-write time change.
    const auto append = [](const std::filesystem::path& path, const std::string& bytes) noexcept
    {
        const wil::unique_hfile file(CreateFileW(
            path.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
        DWORD written = 0u;
        return file.is_valid() && WriteFile(file.get(), bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr) != FALSE && written == bytes.size();
    };
    // A named stream on the file, as a downloaded file's Mark of the Web would be.
    const auto addStream = [](const std::filesystem::path& path) noexcept
    {
        const std::wstring streamPath = path.native() + L":note";
        const wil::unique_hfile stream(CreateFileW(streamPath.c_str(), GENERIC_WRITE, 0u, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        constexpr std::string_view kNote = "note";
        DWORD written                    = 0u;
        return stream.is_valid() && WriteFile(stream.get(), kNote.data(), static_cast<DWORD>(kNote.size()), &written, nullptr) != FALSE &&
               written == kNote.size();
    };

    bool passed            = true;
    unsigned int caseIndex = 0u;
    for (const Scenario& scenario : kScenarios)
    {
        const bool folder           = scenario.shape == Shape::FolderEmptied || scenario.shape == Shape::FolderGainsItem;
        const bool unenforced       = scenario.shape == Shape::SharingUnenforced;
        const bool noDestinationIds = scenario.shape == Shape::ReplaceWithoutIds || scenario.shape == Shape::ReplaceWithoutIdsLong ||
                                      scenario.shape == Shape::PreauthorizedWithoutIds || unenforced;
        const bool replace          = noDestinationIds || scenario.shape == Shape::ReplaceWithIds;
        const bool sameFile         = scenario.shape == Shape::SameFileTwoNames;
        const bool collides         = replace || sameFile || scenario.shape == Shape::SkipExisting;
        const bool pauseAtCleanup   = scenario.shape == Shape::FileChanged || scenario.shape == Shape::FolderGainsItem;
        const auto caseRoot         = root / std::format(L"case-{}", caseIndex++);
        auto sourceFolder           = caseRoot / L"source";
        auto destinationFolder      = caseRoot / L"destination";
        if (scenario.shape == Shape::ReplaceWithoutIdsLong)
        {
            for (unsigned int level = 0u; level < 6u; ++level)
            {
                sourceFolder /= L"source-component-012345678901234567890123456789";
                destinationFolder /= L"destination-component-012345678901234567890123";
            }
        }
        const auto selected      = folder ? sourceFolder / L"album" : sourceFolder / L"harbor.bin";
        const auto sourceFile    = folder ? selected / L"track-01.bin" : selected;
        const auto publishedFile = folder ? destinationFolder / selected.filename() / sourceFile.filename() : destinationFolder / selected.filename();
        const auto addedFile     = selected / L"added.bin";
        if (sameFile)
        {
            // One file in the destination folder; the source folder is a junction to that folder, inside this case's
            // sandbox. The source name and the destination name then reach the same file.
            if (! RecreateEmptyDirectory(caseRoot) || ! RecreateEmptyDirectory(destinationFolder) || ! WriteFileTextFsIo(io, publishedFile, sourceBytes) ||
                ! TryCreateJunction(sourceFolder, destinationFolder))
                return false;
        }
        else if (! RecreateEmptyDirectory(folder ? selected : sourceFolder) || ! RecreateEmptyDirectory(destinationFolder) ||
                 ! WriteFileTextFsIo(io, sourceFile, sourceBytes) || (collides && ! WriteFileTextFsIo(io, publishedFile, occupantBytes)) ||
                 (scenario.shape == Shape::FileWithStream && ! addStream(sourceFile)))
            return false;
        // Only the source's location cannot identify items; for one Replace case neither location can.
        RedSalamander::TestSupport::ScopedEnvironmentVariable identitySupport(kSelfTestEnvBindIdentityUnsupportedPath,
                                                                              (noDestinationIds ? caseRoot : sourceFolder).native());
        // The probe's control answers as a location that admits a second open of a held file would.
        SetFileOpsSameFileProbeSharingUnenforcedForSelfTest(unenforced);
        const auto restoreSharing = wil::scope_exit([]() noexcept { SetFileOpsSameFileProbeSharingUnenforcedForSelfTest(false); });
        RedSalamander::TestSupport::ScopedEnvironmentVariable openRefused(
            L"REDSALAMANDER_FILEOPS_BIND_OPEN_NOT_SUPPORTED_PATH",
            scenario.shape == Shape::FileOpenFails ? std::optional<std::wstring_view>(selected.native()) : std::nullopt);
        FileOperations::QualifiedEndpoint endpoint{
            .pluginId               = L"builtin/file-system",
            .instanceId             = L"host/default",
            .profileId              = L"local",
            .rootId                 = L"local-test-volume",
            .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };
        // The Local destination proves what it publishes; the source-side route requires that proof.
        endpoint.verificationHostReadback = true;
        FileOperations::TransferPlan transfer{};
        transfer.intent                         = FileOperations::TransferIntent::Move;
        transfer.strategy                       = FileOperations::OperationStrategy::Managed;
        transfer.sourceEndpoint                 = endpoint;
        transfer.destinationEndpoint            = endpoint;
        transfer.selectedItems                  = {{.providerPath = selected.native()}};
        transfer.destination.providerFolderPath = destinationFolder.native();
        const auto plans                        = std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}});
        State::Task task(fileOperations);
        task._operation                     = FILESYSTEM_MOVE;
        task._flags                         = scenario.shape == Shape::PreauthorizedWithoutIds ? FILESYSTEM_FLAG_ALLOW_OVERWRITE : FILESYSTEM_FLAG_NONE;
        task._executionMode                 = State::ExecutionMode::PerItem;
        task._fileSystem                    = localFileSystem;
        task._sourcePaths                   = {selected};
        task._sourcePluginId                = endpoint.pluginId;
        task._sourcePluginShortId           = L"file";
        task._sourceInstanceContext         = endpoint.instanceId;
        const std::wstring extendedSelected = Common::Paths::ToExtendedWin32Path(selected.native());
        task._sourcePathAttributesHint      = {GetFileAttributesW(extendedSelected.c_str())};
        task._perItemMaxConcurrency         = 1u;
        task.InitializeAdmissionInputs(plans,
                                       {.fileSystem      = localFileSystem,
                                        .folder          = destinationFolder,
                                        .pluginId        = endpoint.pluginId,
                                        .pluginShortId   = L"file",
                                        .instanceContext = endpoint.instanceId});
        task.InitializeSourceItemResultBuilders();
        const auto diagnosticsCleanup = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });
        // A source whose open fails may already be refused while the task prepares; that is the same refusal.
        const HRESULT prepareHr = task.PrepareMutationInterlockScopes();
        if (FAILED(prepareHr) && scenario.shape != Shape::FileOpenFails && ! sameFile)
            return false;

        // The deterministic "changed in between" seam: the bridge stops just before it removes the first source by name.
        if (pauseAtCleanup)
            SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(true, 30'000ull);
        // A second thread answers decisions (Replace or Skip for the collision cases, Cancel for any other, which is
        // unexpected) and, at the pause, changes the source and releases the bridge.
        const Action wantedAnswer = replace || sameFile ? Action::Overwrite : Action::Skip;
        std::atomic<unsigned int> answered{0u};
        std::atomic<bool> answerOffered{true};
        std::atomic<bool> pauseEntered{false};
        std::atomic<bool> changedInPause{false};
        HRESULT executeHr = prepareHr;
        if (SUCCEEDED(prepareHr))
        {
            const std::jthread responder([&](const std::stop_token stop) noexcept
            {
                uint64_t answeredGeneration = 0u;
                while (! stop.stop_requested())
                {
                    if (pauseAtCleanup && ! pauseEntered.load(std::memory_order_acquire) && HasFileOpsBridgeMoveSourceCleanupPauseEnteredForSelfTest())
                    {
                        const bool changed =
                            scenario.shape == Shape::FileChanged ? append(sourceFile, appendedBytes) : WriteFileTextFsIo(io, addedFile, sourceBytes);
                        changedInPause.store(changed, std::memory_order_release);
                        pauseEntered.store(true, std::memory_order_release);
                        ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
                    }
                    if (const auto prompt = TryGetConflictPromptCopy(&task);
                        prompt.has_value() && prompt->generation != answeredGeneration && ! prompt->metadataLoading && prompt->buttonsPublishable)
                    {
                        const auto offeredEnd = prompt->actions.begin() + static_cast<std::ptrdiff_t>(prompt->actionCount);
                        const bool offered    = std::find(prompt->actions.begin(), offeredEnd, wantedAnswer) != offeredEnd;
                        if (collides && ! offered)
                            answerOffered.store(false, std::memory_order_release);
                        if (task.SubmitConflictDecision(prompt->generation, collides && offered ? wantedAnswer : Action::Cancel, false))
                        {
                            answeredGeneration = prompt->generation;
                            answered.fetch_add(1u, std::memory_order_acq_rel);
                        }
                    }
                    Sleep(2);
                }
            });
            executeHr = task.ExecuteOperation();
        }
        if (pauseAtCleanup)
        {
            ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
            SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
        }
        const HRESULT finalHr = task.FinalizeTypedItemResults(executeHr);
        task._resultHr.store(finalHr, std::memory_order_release);
        const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);

        const unsigned int prompts = answered.load(std::memory_order_acquire);
        const bool paused          = pauseEntered.load(std::memory_order_acquire) && changedInPause.load(std::memory_order_acquire);
        bool outcome               = false;
        switch (scenario.shape)
        {
            case Shape::FileUnchanged:
                outcome = prompts == 0u && absent(sourceFile) && holds(publishedFile, sourceBytes) && summary.publishedItemCount == 1u &&
                          summary.removedSourceCount == 1u && summary.retainedSourceCount == 0u && summary.failedItemCount == 0u && SUCCEEDED(finalHr);
                break;
            case Shape::FileChanged:
                outcome = prompts == 0u && paused && holds(sourceFile, sourceBytes + appendedBytes) && holds(publishedFile, sourceBytes) &&
                          summary.publishedItemCount == 1u && summary.removedSourceCount == 0u && summary.retainedSourceCount == 1u &&
                          summary.failedItemCount == 0u;
                break;
            case Shape::FileWithStream:
                outcome = prompts == 0u && holds(sourceFile, sourceBytes) && holds(publishedFile, sourceBytes) && summary.publishedItemCount == 1u &&
                          summary.removedSourceCount == 0u && summary.retainedSourceCount == 1u && summary.failedItemCount == 0u;
                break;
            case Shape::FileOpenFails:
                outcome = prompts == 0u && FAILED(executeHr) && executeHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) && holds(sourceFile, sourceBytes) &&
                          absent(publishedFile) && summary.publishedItemCount == 0u && summary.removedSourceCount == 0u;
                break;
            case Shape::FolderEmptied:
                outcome = prompts == 0u && absent(selected) && holds(publishedFile, sourceBytes) && summary.removedSourceCount == 1u &&
                          summary.retainedSourceCount == 0u && summary.failedItemCount == 0u && SUCCEEDED(finalHr);
                break;
            case Shape::FolderGainsItem:
                outcome = prompts == 0u && paused && absent(sourceFile) && holds(addedFile, sourceBytes) && holds(publishedFile, sourceBytes) &&
                          absent(destinationFolder / selected.filename() / addedFile.filename()) && summary.removedSourceCount == 0u &&
                          summary.retainedSourceCount == 1u && summary.failedItemCount == 0u;
                break;
            case Shape::ReplaceWithoutIds:
            case Shape::ReplaceWithoutIdsLong:
            case Shape::PreauthorizedWithoutIds:
            case Shape::ReplaceWithIds:
                outcome = prompts == 1u && answerOffered.load(std::memory_order_acquire) && absent(sourceFile) && holds(publishedFile, sourceBytes) &&
                          summary.publishedItemCount == 1u && summary.removedSourceCount == 1u && summary.failedItemCount == 0u && SUCCEEDED(finalHr);
                break;
            case Shape::SkipExisting:
                // Both files stay as they were: nothing is published and no source is removed.
                outcome = prompts == 1u && answerOffered.load(std::memory_order_acquire) && holds(sourceFile, sourceBytes) &&
                          holds(publishedFile, occupantBytes) && summary.publishedItemCount == 0u && summary.removedSourceCount == 0u &&
                          summary.failedItemCount == 0u;
                break;
            case Shape::SharingUnenforced:
                // The replacement is made; the source stays, because nothing proved it is another file than the one replaced.
                outcome = prompts == 1u && answerOffered.load(std::memory_order_acquire) && holds(sourceFile, sourceBytes) &&
                          holds(publishedFile, sourceBytes) && summary.publishedItemCount == 1u && summary.removedSourceCount == 0u &&
                          summary.retainedSourceCount == 1u && summary.failedItemCount == 0u;
                break;
            case Shape::SameFileTwoNames:
                // The one file is still there under both names with its bytes, whichever way the task ended (a refused
                // replacement or a replacement followed by the same-file answer): no source is ever removed.
                outcome = holds(publishedFile, sourceBytes) && holds(sourceFile, sourceBytes) && summary.removedSourceCount == 0u;
                break;
        }
        std::vector<State::TaskDiagnosticEntry> retentionDiagnostics;
        fileOperations.CollectDiagnostics(retentionDiagnostics);
        const bool retentionReasonsLocalized = std::ranges::all_of(retentionDiagnostics, [&](const State::TaskDiagnosticEntry& entry) noexcept {
            return entry.taskId != task._taskId || entry.category != L"bridge.move.sourceRetained" || entry.userText.messageId != 0u;
        });
        const bool matches                   = outcome && (pauseAtCleanup ? paused : true) && retentionReasonsLocalized;
        AppendLog(std::format(L"Identity-less source Move {}: execute=0x{:08X} final=0x{:08X} prompts={} paused={} published={} removed={} retained={} "
                              L"failed={} passed={}",
                              scenario.name,
                              static_cast<unsigned long>(executeHr),
                              static_cast<unsigned long>(finalHr),
                              prompts,
                              paused,
                              summary.publishedItemCount,
                              summary.removedSourceCount,
                              summary.retainedSourceCount,
                              summary.failedItemCount,
                              matches));
        passed = matches && passed;
    }
    return passed;
}

// Exercise the real Local Copy executor and typed finalizer for successful empty-directory
// merges. Existing destinations must remain untouched while the empty source is retained;
// the new-destination control proves the fixture reaches the publication path as well.
[[nodiscard]] bool RunFileOperationLocalEmptyDirectoryCopyNoOpSelfTest(FolderWindow::FileOperationState& fileOperations,
                                                                       const wil::com_ptr<IFileSystem>& localFileSystem,
                                                                       IInformations* localInfo,
                                                                       const std::filesystem::path& testRoot) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;
    if (! localFileSystem || localInfo == nullptr || testRoot.empty())
        return false;

    std::string originalConfiguration;
    if (! BackupPluginConfiguration(localInfo, originalConfiguration))
    {
        AppendLog(L"Local empty-directory Copy fixture: could not capture the Local provider configuration.");
        return false;
    }
    bool configurationRestored      = false;
    const auto restoreConfiguration = wil::scope_exit([&]() noexcept
    {
        if (! configurationRestored && ! SetPluginConfiguration(localInfo, originalConfiguration))
            AppendLog(L"Local empty-directory Copy fixture: failed to restore the exact Local provider configuration.");
    });

    const uint64_t runToken          = (static_cast<uint64_t>(GetCurrentProcessId()) << 32u) ^ GetTickCount64();
    const std::filesystem::path root = testRoot / std::format(L"local-empty-copy-noop-{:016X}", runToken);
    bool fixtureCleanupVerified      = false;
    const auto removeFixture         = wil::scope_exit([&]() noexcept
    {
        if (fixtureCleanupVerified)
            return;
        const bool removed = SelfTest::RemoveAll(root);
        std::error_code cleanupError;
        const bool rootAbsent = ! std::filesystem::exists(root, cleanupError) && ! cleanupError;
        if (! removed || ! rootAbsent)
            AppendLog(std::format(L"Local empty-directory Copy fixture fallback cleanup failed (removed={}, rootAbsent={}, error={}).",
                                  removed,
                                  rootAbsent,
                                  cleanupError.value()));
    });

    const auto exactlyContains = [](const std::filesystem::path& directory, const std::initializer_list<std::wstring_view> expectedChildren) noexcept
    {
        std::error_code ec;
        // The runner's owned path may already approach MAX_PATH before FindFirst adds its
        // search suffix. Observe it through the same extended-path helper as fixture creation.
        std::filesystem::directory_iterator it(std::filesystem::path(MakeExtendedPathForSelfTest(directory)), ec);
        if (ec)
        {
            AppendLog(std::format(L"Local empty-directory Copy inventory: '{}' could not be opened ({}).", directory.native(), ec.value()));
            return false;
        }
        const std::filesystem::directory_iterator end;
        size_t actualCount = 0u;
        while (it != end)
        {
            const std::wstring name = it->path().filename().native();
            if (std::find(expectedChildren.begin(), expectedChildren.end(), name) == expectedChildren.end())
            {
                AppendLog(std::format(L"Local empty-directory Copy inventory: unexpected '{}' in '{}'.", name, directory.native()));
                return false;
            }
            ++actualCount;
            it.increment(ec);
            if (ec)
                return false;
        }
        if (actualCount != expectedChildren.size())
            AppendLog(std::format(
                L"Local empty-directory Copy inventory: '{}' contains {} items, expected {}.", directory.native(), actualCount, expectedChildren.size()));
        return actualCount == expectedChildren.size();
    };

    uint64_t caseOrdinal = 0u;
    const auto runCase   = [&](const unsigned int concurrency,
                               const std::wstring_view label,
                               const bool sourceHasEmptyChild,
                               const bool existingDestination,
                               const bool destinationHasSentinel,
                               const bool directProvider) noexcept
    {
        const std::string configuration = std::format(R"json({{"concurrencyMode":"manual","copyMoveMaxConcurrency":{}}})json", concurrency);
        if (! SetPluginConfiguration(localInfo, configuration))
        {
            AppendLog(std::format(L"Local empty-directory Copy {}: could not set concurrency cap {}.", label, concurrency));
            return false;
        }

        const std::filesystem::path caseRoot          = root / std::format(L"{}-{}-cap{}", directProvider ? L"provider" : L"executor", label, concurrency);
        const std::filesystem::path sourceParent      = caseRoot / L"source-parent";
        const std::filesystem::path destinationParent = caseRoot / L"destination-parent";
        const std::filesystem::path source            = sourceParent / L"empty-tree";
        const std::filesystem::path destination       = destinationParent / L"empty-tree";
        const std::filesystem::path sourceChild       = source / L"empty-child";
        const std::filesystem::path destinationChild  = destination / L"empty-child";
        const std::filesystem::path sentinel          = destination / L"sentinel.txt";
        if (! RecreateEmptyDirectory(source) || ! RecreateEmptyDirectory(destinationParent) || (sourceHasEmptyChild && ! RecreateEmptyDirectory(sourceChild)) ||
            (existingDestination && ! RecreateEmptyDirectory(destination)) ||
            (existingDestination && sourceHasEmptyChild && ! RecreateEmptyDirectory(destinationChild)) ||
            (destinationHasSentinel &&
             ! FileOperationItemResultReducerSelfTestInternal::WriteTextFile(localFileSystem.get(), sentinel.native(), "preserve-me")))
        {
            AppendLog(std::format(L"Local empty-directory Copy {}: failed to create private source/destination paths.", label));
            return false;
        }

        FileOperations::QualifiedEndpoint sourceEndpoint{};
        FileOperations::QualifiedEndpoint destinationEndpoint{};
        if (! FileOperations::TryQualifyEndpointForSelfTest(
                localFileSystem, source.native(), FILESYSTEM_COPY, L"builtin/file-system", L"host/default", sourceEndpoint) ||
            ! FileOperations::TryQualifyEndpointForSelfTest(
                localFileSystem, destinationParent.native(), FILESYSTEM_COPY, L"builtin/file-system", L"host/default", destinationEndpoint))
        {
            AppendLog(std::format(L"Local empty-directory Copy {}: Local endpoints did not qualify.", label));
            return false;
        }

        FileOperations::TransferPlan transfer{};
        transfer.intent                         = FileOperations::TransferIntent::Copy;
        transfer.strategy                       = FileOperations::OperationStrategy::Copy;
        transfer.sourceEndpoint                 = std::move(sourceEndpoint);
        transfer.destinationEndpoint            = std::move(destinationEndpoint);
        transfer.selectedItems                  = {{.providerPath = source.native()}};
        transfer.destination.providerFolderPath = destinationParent.native();
        const auto plans                        = std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}});

        auto taskStorage = std::make_unique<Task>(fileOperations);
        Task& task       = *taskStorage;
        task.StorePlans(plans);
        task._operation = FILESYSTEM_COPY;
        task._executionMode =
            directProvider ? FolderWindow::FileOperationState::ExecutionMode::BulkItems : FolderWindow::FileOperationState::ExecutionMode::PerItem;
        task._fileSystem            = localFileSystem;
        task._sourcePaths           = {source};
        task._flags                 = FILESYSTEM_FLAG_RECURSIVE;
        task._taskId                = 0xE000000000000000ull | ((runToken + (++caseOrdinal * 2u) + (directProvider ? 1u : 0u)) & 0x0FFFFFFFFFFFFFFFull);
        task._sourcePluginId        = L"builtin/file-system";
        task._sourcePluginShortId   = L"file";
        task._sourceInstanceContext = L"host/default";
        task.InitializeAdmissionInputs(plans,
                                       {.fileSystem      = localFileSystem,
                                        .folder          = destinationParent,
                                        .pluginId        = L"builtin/file-system",
                                        .pluginShortId   = L"file",
                                        .instanceContext = L"host/default"});
        task._sourcePathAttributesHint = {GetFileAttributesW(source.c_str())};
        task.InitializeSourceItemResultBuilders();
        const auto removeFixtureDiagnostics = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });
        if (task._sourcePathAttributesHint.front() == INVALID_FILE_ATTRIBUTES)
        {
            AppendLog(std::format(L"Local empty-directory Copy {}: source attributes were unavailable before admission.", label));
            return false;
        }

        HRESULT operationHr = E_UNEXPECTED;
        if (directProvider)
        {
            task.MarkSourceItemsMutationPossible();
            FileSystemOptions options{};
            options.sizeBytes              = sizeof(FileSystemOptions);
            options.copyMoveMaxConcurrency = concurrency;
            options.operationControl       = &task;
            options.operationControlCookie = nullptr;
            task._dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
            const auto callbackScope = wil::scope_exit([&]() noexcept { task._dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
            operationHr              = localFileSystem->CopyItem(source.c_str(), destination.c_str(), FILESYSTEM_FLAG_RECURSIVE, &options, &task, nullptr);
        }
        else
        {
            operationHr = task.ExecuteOperation();
        }
        const HRESULT finalHr       = task.FinalizeTypedItemResults(operationHr);
        const bool hasSingleBuilder = task._sourceItemResultBuilders.size() == 1u;
        const auto* builder         = hasSingleBuilder ? std::addressof(task._sourceItemResultBuilders.front()) : nullptr;
        const auto* item            = builder != nullptr && builder->terminal.has_value() ? std::addressof(builder->terminal.value()) : nullptr;
        const auto* receipt         = builder != nullptr && builder->mutation.has_value() ? std::addressof(builder->mutation.value()) : nullptr;
        const bool noop             = existingDestination;
        std::error_code fsError;
        const bool sourceExists = std::filesystem::is_directory(source, fsError) && ! fsError;
        if (fsError)
        {
            AppendLog(std::format(L"Local empty-directory Copy {}: checking the source directory failed ({}).", label, fsError.value()));
            return false;
        }
        const bool destinationExists = std::filesystem::is_directory(destination, fsError) && ! fsError;
        if (fsError)
        {
            AppendLog(std::format(L"Local empty-directory Copy {}: checking the destination directory failed ({}).", label, fsError.value()));
            return false;
        }
        const bool sourceShape =
            sourceExists && (sourceHasEmptyChild ? exactlyContains(source, {L"empty-child"}) && exactlyContains(sourceChild, {}) : exactlyContains(source, {}));
        bool destinationShape = false;
        if (destinationExists)
        {
            if (sourceHasEmptyChild)
            {
                destinationShape =
                    destinationHasSentinel ? exactlyContains(destination, {L"empty-child", L"sentinel.txt"}) : exactlyContains(destination, {L"empty-child"});
                destinationShape = destinationShape && exactlyContains(destinationChild, {});
            }
            else
            {
                destinationShape = destinationHasSentinel ? exactlyContains(destination, {L"sentinel.txt"}) : exactlyContains(destination, {});
            }
        }
        std::string sentinelText;
        wil::com_ptr<IFileSystemIO> localIo;
        const bool sentinelPreserved =
            ! destinationHasSentinel || (SUCCEEDED(localFileSystem->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) && localIo &&
                                         ReadFileTextFsIo(localIo, sentinel.native(), sentinelText) && sentinelText == "preserve-me");
        const unsigned int expectedCallbacks = directProvider ? 1u : 0u;
        const FileOperations::PublicationState expectedPublication =
            noop ? FileOperations::PublicationState::NotPublished : FileOperations::PublicationState::Published;
        const FileOperations::OwnedStageDisposition expectedStage =
            noop ? FileOperations::OwnedStageDisposition::NotCreated : FileOperations::OwnedStageDisposition::Published;
        const bool receiptMatch = directProvider ? receipt != nullptr && receipt->outcomeKnown == TRUE && receipt->mutationCommitted == (noop ? FALSE : TRUE) &&
                                                       receipt->originalStillPresent == TRUE &&
                                                       receipt->ownedStageDisposition ==
                                                           (noop ? FileSystemOwnedStageDisposition::NotCreated : FileSystemOwnedStageDisposition::Published)
                                                 : receipt == nullptr;
        const bool axesMatch    = item != nullptr && item->completion == FileOperations::ItemCompletion::Completed && item->status == S_OK &&
                                  item->publication == expectedPublication && item->sourceDisposition == FileOperations::SourceDisposition::Retained &&
                                  item->ownedStageDisposition == expectedStage && receiptMatch;
        const bool matches      = hasSingleBuilder && operationHr == S_OK && finalHr == S_OK &&
                                  task._itemCompletedCallbackCount.load(std::memory_order_acquire) == expectedCallbacks &&
                                  (directProvider || task._perItemMaxConcurrencyBudget == concurrency) && axesMatch && task._perf.conflictPromptCount == 0u &&
                                  sourceShape && destinationShape && sentinelPreserved;
        if (! matches)
        {
            AppendLog(std::format(L"Local empty-directory Copy {} {} cap{} mismatch: operation=0x{:08X} final=0x{:08X} callbackCount={} budget={} "
                                  L"builders={} status=0x{:08X} completion={} publication={} sourceDisposition={} stage={} prompts={} "
                                  L"source={} destination={} sentinel={} receiptKnown={} committed={} originalPresent={} receiptStage={}",
                                  directProvider ? L"provider" : L"executor",
                                  label,
                                  concurrency,
                                  static_cast<unsigned long>(operationHr),
                                  static_cast<unsigned long>(finalHr),
                                  task._itemCompletedCallbackCount.load(std::memory_order_acquire),
                                  task._perItemMaxConcurrencyBudget,
                                  task._sourceItemResultBuilders.size(),
                                  static_cast<unsigned long>(item ? item->status : E_UNEXPECTED),
                                  item ? static_cast<unsigned>(item->completion) : UINT_MAX,
                                  item ? static_cast<unsigned>(item->publication) : UINT_MAX,
                                  item ? static_cast<unsigned>(item->sourceDisposition) : UINT_MAX,
                                  item ? static_cast<unsigned>(item->ownedStageDisposition) : UINT_MAX,
                                  task._perf.conflictPromptCount,
                                  sourceShape,
                                  destinationShape,
                                  sentinelPreserved,
                                  receipt != nullptr && receipt->outcomeKnown == TRUE,
                                  receipt != nullptr && receipt->mutationCommitted == TRUE,
                                  receipt != nullptr && receipt->originalStillPresent == TRUE,
                                  receipt ? static_cast<unsigned>(receipt->ownedStageDisposition) : UINT_MAX));
        }
        return matches;
    };

    bool casesPassed = true;
    for (const unsigned int concurrency : {1u, 4u})
    {
        for (const bool directProvider : {false, true})
        {
            casesPassed = runCase(concurrency, L"existing-empty", false, true, false, directProvider) && casesPassed;
            casesPassed = runCase(concurrency, L"existing-sentinel", false, true, true, directProvider) && casesPassed;
            casesPassed = runCase(concurrency, L"existing-nested", true, true, false, directProvider) && casesPassed;
            casesPassed = runCase(concurrency, L"existing-nested-sentinel", true, true, true, directProvider) && casesPassed;
            casesPassed = runCase(concurrency, L"new-destination", true, false, false, directProvider) && casesPassed;
        }
    }

    configurationRestored = SetPluginConfiguration(localInfo, originalConfiguration);
    if (! configurationRestored)
        AppendLog(L"Local empty-directory Copy fixture: failed to restore the exact Local provider configuration.");
    const bool removed = SelfTest::RemoveAll(root);
    std::error_code cleanupError;
    const bool rootAbsent  = ! std::filesystem::exists(root, cleanupError) && ! cleanupError;
    fixtureCleanupVerified = rootAbsent;
    if (! removed || ! rootAbsent)
        AppendLog(std::format(
            L"Local empty-directory Copy fixture checked cleanup failed (removed={}, rootAbsent={}, error={}).", removed, rootAbsent, cleanupError.value()));
    return casesPassed && configurationRestored && removed && rootAbsent;
}

[[nodiscard]] bool RunFileOperationItemResultReducerSelfTest() noexcept
{
    using FileOperationItemResultReducer::Classification;
    using FileOperationItemResultReducer::ExecutionPhase;
    using FileOperationItemResultReducer::Input;
    using FileOperationItemResultReducer::ReceiptState;
    using FileOperationItemResultReducer::Reduce;
    using FileOperations::ItemCompletion;
    using FileOperations::OwnedStageDisposition;
    using FileOperations::PublicationState;
    using FileOperations::SourceDisposition;

    const auto expect = [](const Input& input,
                           const PublicationState publication,
                           const SourceDisposition source,
                           const ItemCompletion completion,
                           const HRESULT status,
                           const OwnedStageDisposition stage = OwnedStageDisposition::NotApplicable) noexcept
    {
        const Classification actual = Reduce(input);
        const bool matches          = actual.publication == publication && actual.sourceDisposition == source && actual.completion == completion &&
                                      actual.status == status && actual.ownedStageDisposition == stage;
        if (! matches)
            AppendLog(std::format(L"Receipt matrix mismatch: op={} phase={} receipt={} hr=0x{:08X}; expected pub/source/completion/stage={}/{}/{}/{} "
                                  L"hr=0x{:08X}; actual={}/{}/{}/{} hr=0x{:08X}",
                                  static_cast<unsigned>(input.operation),
                                  static_cast<unsigned>(input.phase),
                                  static_cast<unsigned>(input.receiptState),
                                  static_cast<unsigned long>(input.status),
                                  static_cast<unsigned>(publication),
                                  static_cast<unsigned>(source),
                                  static_cast<unsigned>(completion),
                                  static_cast<unsigned>(stage),
                                  static_cast<unsigned long>(status),
                                  static_cast<unsigned>(actual.publication),
                                  static_cast<unsigned>(actual.sourceDisposition),
                                  static_cast<unsigned>(actual.completion),
                                  static_cast<unsigned>(actual.ownedStageDisposition),
                                  static_cast<unsigned long>(actual.status)));
        return matches;
    };
    const auto operationFamily = [&](const FileSystemOperation operation) noexcept
    {
        Input input{};
        input.operation             = operation;
        input.phase                 = ExecutionPhase::Preparing;
        input.receiptState          = ReceiptState::Missing;
        input.status                = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        input.cancellationRequested = true;
        if (! expect(input, PublicationState::NotAttempted, SourceDisposition::Retained, ItemCompletion::Canceled, HRESULT_FROM_WIN32(ERROR_CANCELLED)))
            return false;
        // Before the gate the status alone decides completion, even under a requested cancel: a
        // root's own setup failure stays Failed, never Unknown. The dense finalizer supplies the
        // cancellation status to a root stopped before dispatch so a sibling's failure is not its own.
        input.status = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        if (! expect(input, PublicationState::NotAttempted, SourceDisposition::Retained, ItemCompletion::Failed, HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)))
            return false;
        input.status = HRESULT_FROM_WIN32(ERROR_CANCELLED);

        input.receiptState          = ReceiptState::Malformed;
        input.ownedStageDisposition = OwnedStageDisposition::Unknown;
        if (! expect(input, PublicationState::NotAttempted, SourceDisposition::Retained, ItemCompletion::Canceled, HRESULT_FROM_WIN32(ERROR_CANCELLED)))
            return false;
        input.receiptState          = ReceiptState::ExplicitUnknown;
        input.ownedStageDisposition = OwnedStageDisposition::NotApplicable;
        if (! expect(input,
                     operation == FILESYSTEM_DELETE ? PublicationState::NotAttempted : PublicationState::Unknown,
                     operation == FILESYSTEM_COPY ? SourceDisposition::Retained : SourceDisposition::Unknown,
                     ItemCompletion::Indeterminate,
                     HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE)))
            return false;
        input.receiptState = ReceiptState::Missing;

        input.phase  = ExecutionPhase::MutationPossible;
        input.status = S_OK;
        if (operation == FILESYSTEM_COPY)
        {
            if (! expect(input, PublicationState::Unknown, SourceDisposition::Retained, ItemCompletion::Canceled, HRESULT_FROM_WIN32(ERROR_CANCELLED)))
                return false;
        }
        else if (operation == FILESYSTEM_DELETE)
        {
            if (! expect(input, PublicationState::NotAttempted, SourceDisposition::Unknown, ItemCompletion::Canceled, HRESULT_FROM_WIN32(ERROR_CANCELLED)))
                return false;
        }
        else if (! expect(input, PublicationState::Unknown, SourceDisposition::Unknown, ItemCompletion::Canceled, HRESULT_FROM_WIN32(ERROR_CANCELLED)))
        {
            return false;
        }

        input.cancellationRequested = false;
        input.status                = HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED);
        if (! expect(input,
                     operation == FILESYSTEM_DELETE ? PublicationState::NotAttempted : PublicationState::Unknown,
                     operation == FILESYSTEM_COPY ? SourceDisposition::Retained : SourceDisposition::Unknown,
                     ItemCompletion::Failed,
                     HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED)))
            return false;

        return true;
    };

    if (! operationFamily(FILESYSTEM_COPY) || ! operationFamily(FILESYSTEM_MOVE) || ! operationFamily(FILESYSTEM_DELETE))
        return false;

    // S_FALSE is only a skip when no receipt contradicts that interpretation. Known receipt truth
    // remains governed by the mutation and HRESULT, as in the existing callback reducer.
    Input knownSkipStatus{};
    knownSkipStatus.operation            = FILESYSTEM_COPY;
    knownSkipStatus.phase                = ExecutionPhase::MutationPossible;
    knownSkipStatus.receiptState         = ReceiptState::Known;
    knownSkipStatus.status               = S_FALSE;
    knownSkipStatus.mutationCommitted    = true;
    knownSkipStatus.originalStillPresent = true;
    knownSkipStatus.userSkipped          = false;
    if (! expect(knownSkipStatus, PublicationState::Published, SourceDisposition::Retained, ItemCompletion::Completed, S_FALSE))
        return false;
    knownSkipStatus.mutationCommitted = false;
    if (! expect(knownSkipStatus, PublicationState::NotPublished, SourceDisposition::Retained, ItemCompletion::Failed, E_UNEXPECTED))
        return false;

    // Delete must remove the selected root to complete, even when a provider reports
    // a committed mutation. Exercise failure, cancel, Skip and later task-cancel controls
    // independently from the retained storage fact.
    Input retainedDelete{};
    retainedDelete.operation            = FILESYSTEM_DELETE;
    retainedDelete.phase                = ExecutionPhase::MutationPossible;
    retainedDelete.receiptState         = ReceiptState::Known;
    retainedDelete.mutationCommitted    = true;
    retainedDelete.originalStillPresent = true;
    for (const bool laterCancel : {false, true})
    {
        retainedDelete.cancellationRequested = laterCancel;
        for (const HRESULT successfulStatus : {S_OK, S_FALSE})
        {
            retainedDelete.status = successfulStatus;
            if (! expect(retainedDelete, PublicationState::NotAttempted, SourceDisposition::Retained, ItemCompletion::Failed, E_UNEXPECTED))
                return false;
        }
        for (const HRESULT partialFailure : {E_FAIL, HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)})
        {
            retainedDelete.status = partialFailure;
            if (! expect(retainedDelete, PublicationState::NotAttempted, SourceDisposition::Retained, ItemCompletion::Failed, partialFailure))
                return false;
        }
        for (const HRESULT cancellation : {E_ABORT, HRESULT_FROM_WIN32(ERROR_CANCELLED)})
        {
            retainedDelete.status = cancellation;
            if (! expect(
                    retainedDelete, PublicationState::NotAttempted, SourceDisposition::Retained, ItemCompletion::Canceled, HRESULT_FROM_WIN32(ERROR_CANCELLED)))
                return false;
        }
        retainedDelete.status      = S_OK;
        retainedDelete.userSkipped = true;
        if (! expect(retainedDelete, PublicationState::NotAttempted, SourceDisposition::Retained, ItemCompletion::Skipped, S_FALSE))
            return false;
        retainedDelete.userSkipped          = false;
        retainedDelete.originalStillPresent = false;
        if (! expect(retainedDelete, PublicationState::NotAttempted, SourceDisposition::Removed, ItemCompletion::Completed, S_OK))
            return false;
        retainedDelete.originalStillPresent = true;
    }

    Input emptyDirectoryMerge{};
    emptyDirectoryMerge.operation            = FILESYSTEM_COPY;
    emptyDirectoryMerge.phase                = ExecutionPhase::MutationPossible;
    emptyDirectoryMerge.receiptState         = ReceiptState::Known;
    emptyDirectoryMerge.status               = S_OK;
    emptyDirectoryMerge.mutationCommitted    = false;
    emptyDirectoryMerge.originalStillPresent = true;
    for (const OwnedStageDisposition cleanStage : {OwnedStageDisposition::NotApplicable, OwnedStageDisposition::NotCreated, OwnedStageDisposition::Removed})
    {
        emptyDirectoryMerge.ownedStageDisposition = cleanStage;
        if (! expect(emptyDirectoryMerge, PublicationState::NotPublished, SourceDisposition::Retained, ItemCompletion::Completed, S_OK, cleanStage))
            return false;
    }
    // An ordinary S_FALSE without explicit Skip evidence stays invalid; a surviving or
    // uncertain private stage also prevents the successful no-op interpretation.
    emptyDirectoryMerge.status                = S_FALSE;
    emptyDirectoryMerge.ownedStageDisposition = OwnedStageDisposition::NotCreated;
    if (! expect(emptyDirectoryMerge,
                 PublicationState::NotPublished,
                 SourceDisposition::Retained,
                 ItemCompletion::Failed,
                 E_UNEXPECTED,
                 OwnedStageDisposition::NotCreated))
        return false;
    emptyDirectoryMerge.status = S_OK;
    for (const OwnedStageDisposition debtStage : {OwnedStageDisposition::Retained, OwnedStageDisposition::Unknown, OwnedStageDisposition::RetainedIncomplete})
    {
        emptyDirectoryMerge.ownedStageDisposition = debtStage;
        if (! expect(emptyDirectoryMerge,
                     PublicationState::NotPublished,
                     SourceDisposition::Retained,
                     debtStage == OwnedStageDisposition::Retained ? ItemCompletion::Failed : ItemCompletion::Indeterminate,
                     debtStage == OwnedStageDisposition::Retained ? E_UNEXPECTED : HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE),
                     debtStage))
            return false;
    }
    emptyDirectoryMerge.ownedStageDisposition = OwnedStageDisposition::NotCreated;
    emptyDirectoryMerge.userSkipped           = true;
    if (! expect(emptyDirectoryMerge,
                 PublicationState::NotPublished,
                 SourceDisposition::Retained,
                 ItemCompletion::Skipped,
                 S_FALSE,
                 OwnedStageDisposition::NotCreated))
        return false;
    emptyDirectoryMerge.userSkipped           = false;
    emptyDirectoryMerge.status                = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    emptyDirectoryMerge.cancellationRequested = true;
    if (! expect(emptyDirectoryMerge,
                 PublicationState::NotPublished,
                 SourceDisposition::Retained,
                 ItemCompletion::Canceled,
                 HRESULT_FROM_WIN32(ERROR_CANCELLED),
                 OwnedStageDisposition::NotCreated))
        return false;
    // A later task cancel cannot rewrite a completed known no-op receipt.
    emptyDirectoryMerge.status = S_OK;
    if (! expect(emptyDirectoryMerge,
                 PublicationState::NotPublished,
                 SourceDisposition::Retained,
                 ItemCompletion::Completed,
                 S_OK,
                 OwnedStageDisposition::NotCreated))
        return false;
    emptyDirectoryMerge.cancellationRequested = false;
    for (const FileSystemOperation destructiveOperation : {FILESYSTEM_MOVE, FILESYSTEM_DELETE})
    {
        emptyDirectoryMerge.operation = destructiveOperation;
        if (! expect(emptyDirectoryMerge,
                     destructiveOperation == FILESYSTEM_DELETE ? PublicationState::NotAttempted : PublicationState::NotPublished,
                     SourceDisposition::Retained,
                     ItemCompletion::Failed,
                     E_UNEXPECTED,
                     OwnedStageDisposition::NotCreated))
            return false;
    }

    // This exact adapter entry point is used by storeQualifiedItemResult. Its explicit Skip
    // booleans classify the terminal decision while preserving a known commit or non-commit.
    // Cover both the direct and partial-child Skip signals, and ensure an explicit unknown receipt
    // still outranks that decision.
    knownSkipStatus.mutationCommitted          = true;
    const Classification committedExplicitSkip = FileOperationItemResultReducer::ReduceProviderReceipt(knownSkipStatus, true, false);
    const Classification committedPartialSkip  = FileOperationItemResultReducer::ReduceProviderReceipt(knownSkipStatus, false, true);
    if (committedExplicitSkip.completion != ItemCompletion::Skipped || committedExplicitSkip.publication != PublicationState::Published ||
        committedExplicitSkip.sourceDisposition != SourceDisposition::Retained || committedExplicitSkip.status != S_FALSE ||
        committedPartialSkip.completion != ItemCompletion::Skipped || committedPartialSkip.publication != PublicationState::Published ||
        committedPartialSkip.sourceDisposition != SourceDisposition::Retained || committedPartialSkip.status != S_FALSE)
        return false;

    knownSkipStatus.mutationCommitted            = false;
    const Classification uncommittedExplicitSkip = FileOperationItemResultReducer::ReduceProviderReceipt(knownSkipStatus, true, false);
    const Classification uncommittedPartialSkip  = FileOperationItemResultReducer::ReduceProviderReceipt(knownSkipStatus, false, true);
    if (uncommittedExplicitSkip.completion != ItemCompletion::Skipped || uncommittedExplicitSkip.publication != PublicationState::NotPublished ||
        uncommittedExplicitSkip.sourceDisposition != SourceDisposition::Retained || uncommittedExplicitSkip.status != S_FALSE ||
        uncommittedPartialSkip.completion != ItemCompletion::Skipped || uncommittedPartialSkip.publication != PublicationState::NotPublished ||
        uncommittedPartialSkip.sourceDisposition != SourceDisposition::Retained || uncommittedPartialSkip.status != S_FALSE)
        return false;

    knownSkipStatus.receiptState                = ReceiptState::ExplicitUnknown;
    const Classification unknownReceiptWithSkip = FileOperationItemResultReducer::ReduceProviderReceipt(knownSkipStatus, true, false);
    if (unknownReceiptWithSkip.completion != ItemCompletion::Indeterminate || unknownReceiptWithSkip.publication != PublicationState::Unknown ||
        unknownReceiptWithSkip.sourceDisposition != SourceDisposition::Retained || unknownReceiptWithSkip.status != HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE))
        return false;

    // I30 x-datasafety#5: the executor's own status decides a consumed descendant Skip. Exact
    // success after a child Skip is partial work even though a Copy-only or source-kept Move later
    // reports it as S_FALSE; an executor S_FALSE is a whole-root Skip and PARTIAL_COPY already marks
    // the partial merge. Without an observed Skip nothing is promoted.
    if (! FileOperationItemResultReducer::ConsumedDescendantSkip(S_OK, true) || FileOperationItemResultReducer::ConsumedDescendantSkip(S_FALSE, true) ||
        FileOperationItemResultReducer::ConsumedDescendantSkip(HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), true) ||
        FileOperationItemResultReducer::ConsumedDescendantSkip(S_OK, false) || FileOperationItemResultReducer::ConsumedDescendantSkip(E_FAIL, true))
        return false;

    // The qualified executor uses this exact adapter after a provider returns
    // ERROR_PARTIAL_COPY with an explicit descendant Skip and no mutation receipt.
    // Skip remains terminal, but cannot establish which storage changes committed.
    for (const FileSystemOperation operation : {FILESYSTEM_COPY, FILESYSTEM_MOVE, FILESYSTEM_DELETE})
    {
        for (const bool laterCancellation : {false, true})
        {
            Input partialSkip{};
            partialSkip.operation             = operation;
            partialSkip.phase                 = ExecutionPhase::MutationPossible;
            partialSkip.receiptState          = ReceiptState::Missing;
            partialSkip.status                = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            partialSkip.cancellationRequested = laterCancellation;
            const auto reduction              = FileOperationItemResultReducer::ReduceProviderReceipt(partialSkip, false, true);
            if (reduction.completion != ItemCompletion::Skipped || reduction.status != S_FALSE ||
                reduction.publication != (operation == FILESYSTEM_DELETE ? PublicationState::NotAttempted : PublicationState::Unknown) ||
                reduction.sourceDisposition != (operation == FILESYSTEM_COPY ? SourceDisposition::Retained : SourceDisposition::Unknown))
                return false;
            for (const ReceiptState receipt : {ReceiptState::ExplicitUnknown, ReceiptState::Malformed})
            {
                partialSkip.receiptState = receipt;
                const auto unknown       = FileOperationItemResultReducer::ReduceProviderReceipt(partialSkip, false, true);
                if (unknown.completion != ItemCompletion::Indeterminate || unknown.status != HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE))
                    return false;
            }
        }
    }

    const HRESULT canceled   = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    const HRESULT incomplete = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
    // Structurally valid receipts can still violate the shared mutation contract. Neither Skip
    // nor cancellation (including a later task cancel) can turn their missing truth into Retained.
    for (const FileSystemOperation operation : {FILESYSTEM_COPY, FILESYSTEM_MOVE, FILESYSTEM_DELETE})
    {
        for (const bool sourceAbsent : {false, true})
        {
            for (const HRESULT status : {S_OK, S_FALSE, E_FAIL, canceled})
            {
                for (const bool laterCancel : {false, true})
                {
                    for (const unsigned skipFlags : {0u, 1u, 2u})
                    {
                        Input contradiction{};
                        contradiction.operation             = operation;
                        contradiction.phase                 = ExecutionPhase::MutationPossible;
                        contradiction.receiptState          = ReceiptState::Known;
                        contradiction.status                = status;
                        contradiction.originalStillPresent  = ! sourceAbsent;
                        contradiction.ownedStageDisposition = sourceAbsent ? OwnedStageDisposition::NotApplicable : OwnedStageDisposition::Published;
                        contradiction.cancellationRequested = laterCancel;
                        contradiction.userSkipped           = skipFlags == 1u;
                        contradiction.partiallySkipped      = skipFlags == 2u;
                        if (! expect(contradiction,
                                     operation == FILESYSTEM_DELETE ? PublicationState::NotAttempted : PublicationState::Unknown,
                                     operation == FILESYSTEM_COPY ? SourceDisposition::Retained : SourceDisposition::Unknown,
                                     ItemCompletion::Indeterminate,
                                     incomplete,
                                     OwnedStageDisposition::Unknown))
                            return false;

                        contradiction.phase = ExecutionPhase::Preparing;
                        const bool skipped  = skipFlags != 0u;
                        if (! expect(contradiction,
                                     PublicationState::NotAttempted,
                                     SourceDisposition::Retained,
                                     skipped              ? ItemCompletion::Skipped
                                     : status == canceled ? ItemCompletion::Canceled
                                                          : ItemCompletion::Failed,
                                     skipped              ? S_FALSE
                                     : status == canceled ? canceled
                                     : SUCCEEDED(status)  ? E_UNEXPECTED
                                                          : status))
                            return false;
                    }
                }
            }
        }
    }
    for (const ReceiptState receipt : {ReceiptState::ExplicitUnknown, ReceiptState::Malformed})
    {
        Input input{};
        input.operation             = FILESYSTEM_MOVE;
        input.phase                 = ExecutionPhase::MutationPossible;
        input.receiptState          = receipt;
        input.status                = canceled;
        input.cancellationRequested = true;
        input.ownedStageDisposition = receipt == ReceiptState::Malformed ? OwnedStageDisposition::Unknown : OwnedStageDisposition::NotApplicable;
        if (! expect(input, PublicationState::Unknown, SourceDisposition::Unknown, ItemCompletion::Indeterminate, incomplete, input.ownedStageDisposition))
            return false;
    }

    Input lateTaskCancel{};
    lateTaskCancel.operation             = FILESYSTEM_DELETE;
    lateTaskCancel.phase                 = ExecutionPhase::MutationPossible;
    lateTaskCancel.receiptState          = ReceiptState::Missing;
    lateTaskCancel.status                = S_FALSE;
    lateTaskCancel.cancellationRequested = true;
    lateTaskCancel.userSkipped           = true;
    if (! expect(lateTaskCancel, PublicationState::NotAttempted, SourceDisposition::Retained, ItemCompletion::Skipped, S_FALSE))
        return false;

    // A known committed mutation preserves exact storage truth while cancellation remains the
    // terminal cause. Delete has no publication axis; its source receipt carries the commit.
    for (const FileSystemOperation operation : {FILESYSTEM_COPY, FILESYSTEM_MOVE, FILESYSTEM_DELETE})
    {
        Input input{};
        input.operation             = operation;
        input.phase                 = ExecutionPhase::MutationPossible;
        input.receiptState          = ReceiptState::Known;
        input.status                = canceled;
        input.cancellationRequested = true;
        input.mutationCommitted     = true;
        input.originalStillPresent  = false;
        if (! expect(input,
                     operation == FILESYSTEM_DELETE ? PublicationState::NotAttempted : PublicationState::Published,
                     SourceDisposition::Removed,
                     ItemCompletion::Canceled,
                     canceled))
            return false;

        input.originalStillPresent = true;
        if (! expect(input,
                     operation == FILESYSTEM_DELETE ? PublicationState::NotAttempted : PublicationState::Published,
                     SourceDisposition::Retained,
                     ItemCompletion::Canceled,
                     canceled))
            return false;

        input.mutationCommitted = false;
        if (! expect(input,
                     operation == FILESYSTEM_DELETE ? PublicationState::NotAttempted : PublicationState::NotPublished,
                     SourceDisposition::Retained,
                     ItemCompletion::Canceled,
                     canceled))
            return false;
    }

    // A valid known non-commit with a provider-owned stage that may remain is still indeterminate;
    // publication truth and cleanup truth are independent.
    Input cleanupDebt{};
    cleanupDebt.operation             = FILESYSTEM_COPY;
    cleanupDebt.phase                 = ExecutionPhase::MutationPossible;
    cleanupDebt.receiptState          = ReceiptState::Known;
    cleanupDebt.status                = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    cleanupDebt.ownedStageDisposition = OwnedStageDisposition::RetainedIncomplete;
    if (! expect(cleanupDebt,
                 PublicationState::NotPublished,
                 SourceDisposition::Retained,
                 ItemCompletion::Indeterminate,
                 incomplete,
                 OwnedStageDisposition::RetainedIncomplete))
        return false;

    // Missing-root receipts are reduced independently: a previous committed Copy cannot leak
    // into a later selected root whose terminal callback is absent at cancellation.
    Input earlierReceipt{};
    earlierReceipt.operation            = FILESYSTEM_COPY;
    earlierReceipt.phase                = ExecutionPhase::MutationPossible;
    earlierReceipt.receiptState         = ReceiptState::Known;
    earlierReceipt.status               = S_OK;
    earlierReceipt.mutationCommitted    = true;
    earlierReceipt.originalStillPresent = true;
    if (! expect(earlierReceipt, PublicationState::Published, SourceDisposition::Retained, ItemCompletion::Completed, S_OK))
        return false;
    earlierReceipt.receiptState          = ReceiptState::Missing;
    earlierReceipt.status                = canceled;
    earlierReceipt.cancellationRequested = true;
    earlierReceipt.mutationCommitted     = false;
    earlierReceipt.originalStillPresent  = true;
    if (! expect(earlierReceipt, PublicationState::Unknown, SourceDisposition::Retained, ItemCompletion::Canceled, canceled))
        return false;

    // Mixed task roots retain independent terminal truth instead of collapsing cancellation or
    // unknown publication from one item into another item's receipt.
    std::array<Classification, 3> mixed{};
    Input mixedInput{};
    mixedInput.operation             = FILESYSTEM_COPY;
    mixedInput.phase                 = ExecutionPhase::MutationPossible;
    mixedInput.receiptState          = ReceiptState::Known;
    mixedInput.status                = S_OK;
    mixedInput.mutationCommitted     = true;
    mixedInput.originalStillPresent  = true;
    mixed[0]                         = Reduce(mixedInput);
    mixedInput.operation             = FILESYSTEM_MOVE;
    mixedInput.receiptState          = ReceiptState::Missing;
    mixedInput.status                = canceled;
    mixedInput.cancellationRequested = true;
    mixedInput.mutationCommitted     = false;
    mixed[1]                         = Reduce(mixedInput);
    mixedInput.operation             = FILESYSTEM_DELETE;
    mixedInput.receiptState          = ReceiptState::ExplicitUnknown;
    mixedInput.cancellationRequested = false;
    mixed[2]                         = Reduce(mixedInput);
    return mixed[0].completion == ItemCompletion::Completed && mixed[0].publication == PublicationState::Published &&
           mixed[1].completion == ItemCompletion::Canceled && mixed[1].publication == PublicationState::Unknown &&
           mixed[1].sourceDisposition == SourceDisposition::Unknown && mixed[2].completion == ItemCompletion::Indeterminate &&
           mixed[2].publication == PublicationState::NotAttempted && mixed[2].sourceDisposition == SourceDisposition::Unknown && mixed[2].status == incomplete;
}

// Exercises the production callback receipt replacement and task-level dense finalizer through
// BulkItems Tasks. The provider callback is driven synchronously so the test needs no UI, worker,
// filesystem mutation, or timing assumption.
// I35 S4 round 3: a receipt that proved part of an item committed stays committed for that item: a later call for
// the same item that reports nothing deleted (a refused retry) never turns it back into "nothing changed".
[[nodiscard]] bool RunFileOperationStickyCommittedReceiptSelfTest(FolderWindow::FileOperationState& fileOperations) noexcept
{
    using Task              = FolderWindow::FileOperationState::Task;
    const uint64_t runToken = (static_cast<uint64_t>(GetCurrentProcessId()) << 32u) ^ GetTickCount64();
    auto taskStorage        = std::make_unique<Task>(fileOperations);
    Task& task              = *taskStorage;
    task._operation         = FILESYSTEM_DELETE;
    task._executionMode     = FolderWindow::FileOperationState::ExecutionMode::PerItem;
    task._sourcePaths       = {std::filesystem::path(L"C:\\sticky-committed-receipt\\folder")};
    task._taskId            = 0xF100000000000000ull | (runToken & 0x00FFFFFFFFFFFFFFull);
    task.InitializeSourceItemResultBuilders();
    const auto removeFixtureDiagnostics = wil::scope_exit([&]() noexcept { fileOperations.DebugRemoveDiagnosticsForTask(task._taskId); });

    Task::PerItemCallbackCookie cookie{0u};
    const FileSystemItemMutationResult partial{sizeof(FileSystemItemMutationResult), TRUE, TRUE, TRUE};
    const FileSystemItemMutationResult refused{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
    {
        task._dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
        const auto callbackScope = wil::scope_exit([&]() noexcept { task._dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
        static_cast<void>(task.FileSystemItemCompleted(
            FILESYSTEM_DELETE, 0u, task._sourcePaths.front().c_str(), nullptr, HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY), &partial, nullptr, &cookie));
        static_cast<void>(task.FileSystemItemCompleted(
            FILESYSTEM_DELETE, 0u, task._sourcePaths.front().c_str(), nullptr, HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), &refused, nullptr, &cookie));
    }
    const std::optional<FileSystemItemMutationResult>& receipt = task._sourceItemResultBuilders.front().mutation;
    const bool sticky = receipt.has_value() && receipt->outcomeKnown == TRUE && receipt->mutationCommitted == TRUE && receipt->originalStillPresent == TRUE;
    if (! sticky)
        AppendLog(std::format(L"Sticky committed receipt: a refused second call left known={} committed={} present={}.",
                              receipt.has_value() ? static_cast<int>(receipt->outcomeKnown) : -1,
                              receipt.has_value() ? static_cast<int>(receipt->mutationCommitted) : -1,
                              receipt.has_value() ? static_cast<int>(receipt->originalStillPresent) : -1));
    return sticky;
}

[[nodiscard]] bool RunFileOperationCanceledMissingReceiptTaskSelfTest(FolderWindow::FileOperationState& fileOperations) noexcept
{
    using Task     = FolderWindow::FileOperationState::Task;
    const auto run = [&](const FileSystemOperation operation, const bool sendCommittedThenNull, const bool omitTerminalCallback) noexcept
    {
        auto taskStorage    = std::make_unique<Task>(fileOperations);
        Task& task          = *taskStorage;
        task._operation     = operation;
        task._executionMode = FolderWindow::FileOperationState::ExecutionMode::BulkItems;
        task._sourcePaths   = {std::filesystem::path(operation == FILESYSTEM_MOVE ? L"C:\\fileops-reducer-move.bin" : L"C:\\fileops-reducer-delete.bin")};

        FileOperations::FileOperationPlan plan;
        if (operation == FILESYSTEM_MOVE)
        {
            FileOperations::TransferPlan transfer{};
            transfer.intent                         = FileOperations::TransferIntent::Move;
            transfer.strategy                       = FileOperations::OperationStrategy::Native;
            transfer.selectedItems                  = {{.providerPath = L"/fileops-reducer/source.bin"}};
            transfer.destination.providerFolderPath = L"/fileops-reducer/destination";
            plan                                    = FileOperations::FileOperationPlan{std::move(transfer)};
        }
        else
        {
            FileOperations::DeletePlan deletion{};
            deletion.selectedItems = {{.providerPath = L"/fileops-reducer/delete.bin"}};
            plan                   = FileOperations::FileOperationPlan{std::move(deletion)};
        }
        task.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(FileOperations::FileOperationPlanGroup{std::move(plan)}));
        task.InitializeSourceItemResultBuilders();
        task.MarkSourceItemsMutationPossible();

        HRESULT priorHr            = S_OK;
        HRESULT canceledCallbackHr = S_OK;
        bool priorReceiptObserved  = false;
        if (sendCommittedThenNull || ! omitTerminalCallback)
        {
            task._dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
            if (sendCommittedThenNull)
            {
                const FileSystemItemMutationResult committed{
                    sizeof(FileSystemItemMutationResult), TRUE, TRUE, FALSE, FileSystemOwnedStageDisposition::Published};
                priorHr              = task.FileSystemItemCompleted(operation, 0u, nullptr, nullptr, S_OK, &committed, nullptr, nullptr);
                priorReceiptObserved = task._sourceItemResultBuilders.front().mutation.has_value();
            }
            task.RequestCancel();
            if (! omitTerminalCallback)
            {
                canceledCallbackHr =
                    task.FileSystemItemCompleted(operation, 0u, nullptr, nullptr, HRESULT_FROM_WIN32(ERROR_CANCELLED), nullptr, nullptr, nullptr);
            }
            task._dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed);
        }
        else
        {
            task.RequestCancel();
        }

        const HRESULT finalHr = task.FinalizeTypedItemResults(S_OK);
        const FileOperations::FileOperationItemResult* item =
            task._sourceItemResultBuilders.size() == 1u && task._sourceItemResultBuilders.front().terminal.has_value()
                ? std::addressof(task._sourceItemResultBuilders.front().terminal.value())
                : nullptr;
        const FileSystemRouteContract::MutationClassification recovery =
            FileSystemRouteContract::ClassifyFailedMutation(true, HRESULT_FROM_WIN32(ERROR_CANCELLED), nullptr);
        const bool expectedAxes =
            item != nullptr && item->completion == FileOperations::ItemCompletion::Canceled && item->status == HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
            (operation == FILESYSTEM_MOVE
                 ? item->publication == FileOperations::PublicationState::Unknown && item->sourceDisposition == FileOperations::SourceDisposition::Unknown
                 : item->publication == FileOperations::PublicationState::NotAttempted &&
                       item->sourceDisposition == FileOperations::SourceDisposition::Unknown);
        return (! sendCommittedThenNull || (priorHr == S_OK && priorReceiptObserved && ! task._sourceItemResultBuilders.front().mutation.has_value())) &&
               (omitTerminalCallback || canceledCallbackHr == HRESULT_FROM_WIN32(ERROR_CANCELLED)) && finalHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
               expectedAxes && recovery == FileSystemRouteContract::MutationClassification::Indeterminate;
    };

    for (const FileSystemOperation operation : {FILESYSTEM_MOVE, FILESYSTEM_DELETE})
    {
        if (! run(operation, true, false) || ! run(operation, false, false) || ! run(operation, false, true))
            return false;
    }
    return true;
}

// Exercise per-root Skip and contradictory statuses/storage receipts through the actual
// callback and dense aggregate reducer, rather than validating only the pure classifier.
[[nodiscard]] bool RunFileOperationReceiptStatusTaskSelfTest(FolderWindow::FileOperationState& fileOperations) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;
    for (const bool committed : {false, true})
    {
        for (const HRESULT callbackStatus : {S_OK, S_FALSE})
        {
            auto taskStorage    = std::make_unique<Task>(fileOperations);
            Task& task          = *taskStorage;
            task._operation     = FILESYSTEM_COPY;
            task._executionMode = FolderWindow::FileOperationState::ExecutionMode::BulkItems;
            task._sourcePaths   = {L"C:\\fileops-receipt-status.bin"};
            FileOperations::TransferPlan transfer{};
            transfer.intent                         = FileOperations::TransferIntent::Copy;
            transfer.strategy                       = FileOperations::OperationStrategy::Native;
            transfer.selectedItems                  = {{.providerPath = L"/fileops-receipt-status/source.bin"}};
            transfer.destination.providerFolderPath = L"/fileops-receipt-status/destination";
            task.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
                FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(transfer)}}));
            task.InitializeSourceItemResultBuilders();
            task.MarkSourceItemsMutationPossible();
            const FileSystemItemMutationResult receipt{
                sizeof(FileSystemItemMutationResult), TRUE, committed ? TRUE : FALSE, TRUE, FileSystemOwnedStageDisposition::NotApplicable};
            task._dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
            const HRESULT observed = task.FileSystemItemCompleted(FILESYSTEM_COPY, 0u, nullptr, nullptr, callbackStatus, &receipt, nullptr, nullptr);
            task._dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed);
            const HRESULT expectedStatus  = callbackStatus;
            const auto expectedCompletion = callbackStatus == S_FALSE               ? FileOperations::ItemCompletion::Skipped
                                            : (committed || callbackStatus == S_OK) ? FileOperations::ItemCompletion::Completed
                                                                                    : FileOperations::ItemCompletion::Failed;
            const HRESULT aggregate       = task.FinalizeTypedItemResults(callbackStatus);
            const auto& terminal          = task._sourceItemResultBuilders.front().terminal;
            // The callback acknowledges delivery; its return value does not echo the provider's
            // item status. That status is retained in the dense result and task aggregate.
            if (observed != S_OK || ! terminal || aggregate != expectedStatus || terminal.value().status != expectedStatus ||
                terminal.value().completion != expectedCompletion || terminal.value().sourceDisposition != FileOperations::SourceDisposition::Retained ||
                terminal.value().publication != (committed ? FileOperations::PublicationState::Published : FileOperations::PublicationState::NotPublished))
            {
                AppendLog(std::format(L"Receipt task mismatch: committed={} callback=0x{:08X} acknowledgement=0x{:08X} aggregate=0x{:08X}",
                                      committed,
                                      static_cast<unsigned long>(callbackStatus),
                                      static_cast<unsigned long>(observed),
                                      static_cast<unsigned long>(aggregate)));
                return false;
            }
        }
    }
    for (const FileSystemOperation operation : {FILESYSTEM_COPY, FILESYSTEM_MOVE, FILESYSTEM_DELETE})
    {
        // The third variant controls valid no-commit truth. The fourth proves that the
        // production callback/finalizer rejects successful Delete with a retained root.
        for (const unsigned variant : {0u, 1u, 2u, 3u})
        {
            if (variant == 3u && operation != FILESYSTEM_DELETE)
                continue;
            for (const HRESULT callbackStatus : {S_OK, S_FALSE, E_FAIL, HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), HRESULT_FROM_WIN32(ERROR_CANCELLED)})
            {
                for (const bool laterCancel : {false, true})
                {
                    auto taskStorage    = std::make_unique<Task>(fileOperations);
                    Task& task          = *taskStorage;
                    task._operation     = operation;
                    task._executionMode = FolderWindow::FileOperationState::ExecutionMode::BulkItems;
                    task._sourcePaths   = {L"C:\\fileops-receipt-contradiction.bin"};
                    FileOperations::FileOperationPlan plan;
                    if (operation == FILESYSTEM_DELETE)
                    {
                        FileOperations::DeletePlan deletion{};
                        deletion.selectedItems = {{.providerPath = L"/fileops-receipt-contradiction/source.bin"}};
                        plan                   = FileOperations::FileOperationPlan{std::move(deletion)};
                    }
                    else
                    {
                        FileOperations::TransferPlan transfer{};
                        transfer.intent        = operation == FILESYSTEM_MOVE ? FileOperations::TransferIntent::Move : FileOperations::TransferIntent::Copy;
                        transfer.strategy      = FileOperations::OperationStrategy::Native;
                        transfer.selectedItems = {{.providerPath = L"/fileops-receipt-contradiction/source.bin"}};
                        transfer.destination.providerFolderPath = L"/fileops-receipt-contradiction/destination";
                        plan                                    = FileOperations::FileOperationPlan{std::move(transfer)};
                    }
                    task.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(FileOperations::FileOperationPlanGroup{std::move(plan)}));
                    task.InitializeSourceItemResultBuilders();
                    task.MarkSourceItemsMutationPossible();
                    const FileSystemItemMutationResult receipt{sizeof(FileSystemItemMutationResult),
                                                               TRUE,
                                                               variant == 3u ? TRUE : FALSE,
                                                               variant == 0u ? FALSE : TRUE,
                                                               variant == 1u ? FileSystemOwnedStageDisposition::Published
                                                                             : FileSystemOwnedStageDisposition::NotApplicable};
                    const bool contradictory = variant < 2u;
                    const auto recovery      = FileSystemRouteContract::ClassifyFailedMutation(true, callbackStatus, &receipt);
                    if ((recovery == FileSystemRouteContract::MutationClassification::ContractViolation) != contradictory)
                        return false;
                    task._dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
                    const HRESULT acknowledgement = task.FileSystemItemCompleted(operation, 0u, nullptr, nullptr, callbackStatus, &receipt, nullptr, nullptr);
                    task._dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed);
                    if (laterCancel)
                        task.RequestCancel();
                    const HRESULT aggregate        = task.FinalizeTypedItemResults(callbackStatus);
                    const auto& terminal           = task._sourceItemResultBuilders.front().terminal;
                    const bool skipped             = callbackStatus == S_FALSE;
                    const bool canceled            = callbackStatus == HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    const bool successfulCopyNoop  = operation == FILESYSTEM_COPY && variant == 2u && callbackStatus == S_OK;
                    const HRESULT expectedStatus   = contradictory               ? HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE)
                                                     : skipped                   ? S_FALSE
                                                     : canceled                  ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                     : successfulCopyNoop        ? S_OK
                                                     : SUCCEEDED(callbackStatus) ? E_UNEXPECTED
                                                                                 : callbackStatus;
                    const auto expectedCompletion  = contradictory        ? FileOperations::ItemCompletion::Indeterminate
                                                     : skipped            ? FileOperations::ItemCompletion::Skipped
                                                     : canceled           ? FileOperations::ItemCompletion::Canceled
                                                     : successfulCopyNoop ? FileOperations::ItemCompletion::Completed
                                                                          : FileOperations::ItemCompletion::Failed;
                    const auto expectedPublication = operation == FILESYSTEM_DELETE ? FileOperations::PublicationState::NotAttempted
                                                     : contradictory                ? FileOperations::PublicationState::Unknown
                                                                                    : FileOperations::PublicationState::NotPublished;
                    const auto expectedSource      = contradictory && operation != FILESYSTEM_COPY ? FileOperations::SourceDisposition::Unknown
                                                                                                   : FileOperations::SourceDisposition::Retained;
                    const auto expectedStage =
                        contradictory ? FileOperations::OwnedStageDisposition::Unknown : FileOperations::OwnedStageDisposition::NotApplicable;
                    if (acknowledgement != S_OK || ! terminal || aggregate != expectedStatus || terminal.value().status != expectedStatus ||
                        terminal.value().completion != expectedCompletion || terminal.value().publication != expectedPublication ||
                        terminal.value().sourceDisposition != expectedSource || terminal.value().ownedStageDisposition != expectedStage)
                    {
                        AppendLog(std::format(
                            L"Receipt contradiction task mismatch: op={} variant={} callback=0x{:08X} laterCancel={} aggregate=0x{:08X} expected=0x{:08X}",
                            static_cast<unsigned>(operation),
                            variant,
                            static_cast<unsigned long>(callbackStatus),
                            laterCancel,
                            static_cast<unsigned long>(aggregate),
                            static_cast<unsigned long>(expectedStatus)));
                        return false;
                    }
                }
            }
        }
    }
    return true;
}

// The Dummy provider's native Move and bulk Delete mutate before entering its existing
// cancellation-aware bandwidth wait, then return cancellation before FileSystemItemCompleted.
// This runs that real provider/ExecuteOperation path with a host-side 1 B/s test limit; the
// dedicated provider namespace is cleaned and no provider configuration is changed.
[[nodiscard]] bool RunFileOperationCanceledProviderNullReceiptSelfTest(FolderWindow::FileOperationState& fileOperations,
                                                                       const wil::com_ptr<IFileSystem>& dummyFileSystem,
                                                                       const FileOperations::QualifiedEndpoint& dummyEndpoint,
                                                                       const std::wstring_view runToken) noexcept
{
    if (! dummyFileSystem || runToken.empty())
        return false;

    using Task = FolderWindow::FileOperationState::Task;
    wil::com_ptr<IFileSystemIO> io;
    if (FAILED(dummyFileSystem->QueryInterface(IID_PPV_ARGS(io.addressof()))) || ! io)
        return false;

    const std::wstring root            = L"/fileops-canceled-null-receipt-" + std::wstring(runToken);
    const std::wstring moveFolder      = root + L"/move";
    const std::wstring moveDestination = moveFolder + L"/destination";
    const std::wstring moveSource      = moveFolder + L"/source.bin";
    const std::wstring movedSource     = moveDestination + L"/source.bin";
    const std::wstring deleteFolder    = root + L"/delete";
    const std::wstring deleteSource    = deleteFolder + L"/source.bin";
    const auto cleanup =
        wil::scope_exit([&]() noexcept { static_cast<void>(dummyFileSystem->DeleteItem(root.c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr)); });
    static_cast<void>(dummyFileSystem->DeleteItem(root.c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr));
    if (! FileOperationItemResultReducerSelfTestInternal::EnsureFolder(dummyFileSystem.get(), root) ||
        ! FileOperationItemResultReducerSelfTestInternal::EnsureFolder(dummyFileSystem.get(), moveFolder) ||
        ! FileOperationItemResultReducerSelfTestInternal::EnsureFolder(dummyFileSystem.get(), moveDestination) ||
        ! FileOperationItemResultReducerSelfTestInternal::EnsureFolder(dummyFileSystem.get(), deleteFolder) ||
        ! FileOperationItemResultReducerSelfTestInternal::WriteTextFile(dummyFileSystem.get(), moveSource, "move-payload") ||
        ! FileOperationItemResultReducerSelfTestInternal::WriteTextFile(dummyFileSystem.get(), deleteSource, "delete-payload"))
    {
        AppendLog(L"Canceled-provider fixture: failed to create its provider namespace or payload files.");
        return false;
    }

    const auto exists = [&](const std::wstring& path) noexcept
    {
        unsigned long attributes = 0u;
        return SUCCEEDED(io->GetAttributes(path.c_str(), &attributes));
    };
    const auto executeAndCancelAfterMutation = [&](Task& task, const auto& mutationVisible) noexcept
    {
        // Direct executor fixtures must cross the same gate as ThreadMain before entering
        // the provider; receipt-less completion must not be classified as never started.
        task.MarkSourceItemsMutationPossible();
        std::atomic<HRESULT> operationHr{E_PENDING};
        std::atomic<HRESULT> finalHr{E_PENDING};
        std::atomic<bool> finished{false};
        std::jthread worker([&]() noexcept
        {
            [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast();
            const HRESULT hr             = task.ExecuteOperation();
            operationHr.store(hr, std::memory_order_release);
            finalHr.store(task.FinalizeTypedItemResults(hr), std::memory_order_release);
            finished.store(true, std::memory_order_release);
        });

        bool observedMutation    = false;
        const ULONGLONG deadline = GetTickCount64() + static_cast<ULONGLONG>(SelfTest::Scale(std::chrono::milliseconds{10'000}).count());
        while (! finished.load(std::memory_order_acquire) && GetTickCount64() < deadline)
        {
            if (mutationVisible())
            {
                observedMutation = true;
                task.RequestCancel();
                break;
            }
            Sleep(5u);
        }
        if (! observedMutation)
            task.RequestCancel();
        if (worker.joinable())
            worker.join();

        const FileOperations::FileOperationItemResult* item =
            task._sourceItemResultBuilders.size() == 1u && task._sourceItemResultBuilders.front().terminal.has_value()
                ? std::addressof(task._sourceItemResultBuilders.front().terminal.value())
                : nullptr;
        const bool axesMatch =
            item != nullptr && item->completion == FileOperations::ItemCompletion::Canceled && item->status == HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
            (task._operation == FILESYSTEM_MOVE
                 ? item->publication == FileOperations::PublicationState::Unknown && item->sourceDisposition == FileOperations::SourceDisposition::Unknown
                 : item->publication == FileOperations::PublicationState::NotAttempted &&
                       item->sourceDisposition == FileOperations::SourceDisposition::Unknown);
        const bool matches =
            observedMutation && operationHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
            finalHr.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED) && axesMatch &&
            task._itemCompletedCallbackCount.load(std::memory_order_acquire) == 0u && ! task._sourceItemResultBuilders.front().mutation.has_value() &&
            FileSystemRouteContract::ClassifyFailedMutation(true, item->status, nullptr) == FileSystemRouteContract::MutationClassification::Indeterminate;
        if (! matches)
            AppendLog(std::format(L"Canceled-provider mismatch: op={} mutationObserved={} operation=0x{:08X} aggregate=0x{:08X} callbacks={} "
                                  L"receiptPresent={} axesMatch={} pub/source/completion={}/{}/{} item=0x{:08X}",
                                  static_cast<unsigned>(task._operation),
                                  observedMutation,
                                  static_cast<unsigned long>(operationHr.load(std::memory_order_acquire)),
                                  static_cast<unsigned long>(finalHr.load(std::memory_order_acquire)),
                                  task._itemCompletedCallbackCount.load(std::memory_order_acquire),
                                  task._sourceItemResultBuilders.front().mutation.has_value(),
                                  axesMatch,
                                  item ? static_cast<unsigned>(item->publication) : UINT_MAX,
                                  item ? static_cast<unsigned>(item->sourceDisposition) : UINT_MAX,
                                  item ? static_cast<unsigned>(item->completion) : UINT_MAX,
                                  static_cast<unsigned long>(item ? item->status : E_PENDING)));
        else
            AppendLog(std::format(L"Canceled-provider witness: op={} mutationObserved=true callbacks=0 receiptPresent=false "
                                  L"pub/source/completion={}/{}/{} item=0x{:08X}",
                                  static_cast<unsigned>(task._operation),
                                  static_cast<unsigned>(item->publication),
                                  static_cast<unsigned>(item->sourceDisposition),
                                  static_cast<unsigned>(item->completion),
                                  static_cast<unsigned long>(item->status)));
        if (! matches)
            return false;

        // The task ends canceled, but its root's outcome is unknown: the production completed
        // summary says so and carries a warning, so auto-dismiss never removes the card silently.
        task._taskId = 0x7E57'CA11'0000'0000ull + static_cast<uint64_t>(task._operation);
        task._resultHr.store(finalHr.load(std::memory_order_acquire), std::memory_order_release);
        fileOperations.DebugRemoveDiagnosticsForTask(task._taskId);
        const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);
        fileOperations.DebugRemoveDiagnosticsForTask(task._taskId);
        const bool summaryMatches = summary.resultHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && summary.indeterminateItemCount == 1u &&
                                    summary.resultSummary == LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN) && summary.warningCount == 1u &&
                                    summary.errorCount == 0u &&
                                    summary.lastDiagnosticMessage == LoadStringResource(nullptr, IDS_FILEOPS_DIAG_CANCELED_OUTCOME_UNKNOWN) &&
                                    ! IsAutoDismissableFileOperationCompletion(summary.resultHr, summary.warningCount, summary.errorCount);
        AppendLog(std::format(L"Canceled-provider summary: op={} hr=0x{:08X} indeterminate={} warnings={} errors={} summary='{}' message='{}' {}",
                              static_cast<unsigned>(task._operation),
                              static_cast<unsigned long>(summary.resultHr),
                              summary.indeterminateItemCount,
                              summary.warningCount,
                              summary.errorCount,
                              summary.resultSummary,
                              summary.lastDiagnosticMessage,
                              summaryMatches ? L"kept" : L"MISMATCH"));
        return summaryMatches;
    };

    FileOperations::TransferPlan movePlan{};
    movePlan.intent                         = FileOperations::TransferIntent::Move;
    movePlan.strategy                       = FileOperations::OperationStrategy::Native;
    movePlan.sourceEndpoint                 = dummyEndpoint;
    movePlan.destinationEndpoint            = dummyEndpoint;
    movePlan.selectedItems                  = {{.providerPath = moveSource}};
    movePlan.destination.providerFolderPath = moveDestination;
    auto moveTaskStorage                    = std::make_unique<Task>(fileOperations);
    Task& moveTask                          = *moveTaskStorage;
    moveTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
        FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(movePlan)}}));
    moveTask._operation             = FILESYSTEM_MOVE;
    moveTask._executionMode         = FolderWindow::FileOperationState::ExecutionMode::PerItem;
    moveTask._fileSystem            = dummyFileSystem;
    moveTask._sourcePaths           = {std::filesystem::path(moveSource)};
    moveTask._perItemMaxConcurrency = 1u;
    moveTask.InitializeAdmissionInputs(moveTask.LoadPlans(), {.fileSystem = dummyFileSystem, .folder = std::filesystem::path(moveDestination)});
    moveTask._sourcePathAttributesHint = {FILE_ATTRIBUTE_NORMAL};
    moveTask.InitializeSourceItemResultBuilders();
    moveTask.SetDesiredSpeedLimit(1u);
    if (! executeAndCancelAfterMutation(moveTask, [&]() noexcept { return ! exists(moveSource) && exists(movedSource); }))
        return false;

    FileOperations::DeletePlan deletePlan{};
    deletePlan.endpoint      = dummyEndpoint;
    deletePlan.mode          = FileOperations::DeleteMode::Recycle;
    deletePlan.selectedItems = {{.providerPath = deleteSource}};
    auto deleteTaskStorage   = std::make_unique<Task>(fileOperations);
    Task& deleteTask         = *deleteTaskStorage;
    deleteTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
        FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(deletePlan)}}));
    deleteTask._operation     = FILESYSTEM_DELETE;
    deleteTask._executionMode = FolderWindow::FileOperationState::ExecutionMode::BulkItems;
    deleteTask._fileSystem    = dummyFileSystem;
    deleteTask._sourcePaths   = {std::filesystem::path(deleteSource)};
    deleteTask._flags         = FILESYSTEM_FLAG_USE_RECYCLE_BIN;
    deleteTask.InitializeAdmissionInputs(deleteTask.LoadPlans(), {.fileSystem = dummyFileSystem, .folder = std::filesystem::path(deleteFolder)});
    deleteTask._sourcePathAttributesHint = {FILE_ATTRIBUTE_NORMAL};
    deleteTask.InitializeSourceItemResultBuilders();
    deleteTask.SetDesiredSpeedLimit(1u);
    return executeAndCancelAfterMutation(deleteTask, [&]() noexcept { return ! exists(deleteSource); });
}

// The per-item Copy/Move/Delete executor crosses the mutation gate one selected root at a time,
// when it dispatches that root; bulk and rename executors keep the task-wide gate. Through the
// gate that ThreadMain enters and the dense finalizer, a root never dispatched before a cancel, or
// before a sibling's failure stops the run, is untouched (NotAttempted/Retained/Canceled), while a
// dispatched root with no receipt stays cautious and a root's own terminal truth is kept. A setup
// failure before any dispatch, with no cancel, still fails every root.
[[nodiscard]] bool RunFileOperationUndispatchedRootGateTaskSelfTest(FolderWindow::FileOperationState& fileOperations) noexcept
{
    using Task          = FolderWindow::FileOperationState::Task;
    using ExecutionMode = FolderWindow::FileOperationState::ExecutionMode;
    using FileOperations::ItemCompletion;
    using FileOperations::PublicationState;
    using FileOperations::SourceDisposition;

    enum class Stop : uint8_t
    {
        UserCancel,        // the user/host canceled while root 1 was in flight
        FailureCancelTask, // root 1 failed; the concurrent scheduler's CancelTask policy requested cancel
        FailureSequential, // root 1 failed; the sequential loop stopped without a cancel request
        SetupFailure,      // ExecuteOperation failed before dispatching any root, with no cancel
    };

    const HRESULT canceled     = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    const HRESULT accessDenied = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    const auto run             = [&](const FileSystemOperation operation, const ExecutionMode mode, const Stop stop) noexcept
    {
        auto taskStorage    = std::make_unique<Task>(fileOperations);
        Task& task          = *taskStorage;
        task._operation     = operation;
        task._executionMode = mode;
        task._sourcePaths   = {std::filesystem::path(L"C:\\fileops-gate-0.bin"),
                               std::filesystem::path(L"C:\\fileops-gate-1.bin"),
                               std::filesystem::path(L"C:\\fileops-gate-2.bin"),
                               std::filesystem::path(L"C:\\fileops-gate-3.bin")};
        task.InitializeSourceItemResultBuilders();
        task.EnterSourceItemsMutationGate();
        const bool perRoot = mode == ExecutionMode::PerItem && operation != FILESYSTEM_RENAME;
        if (perRoot && stop != Stop::SetupFailure)
        {
            // QualifiedItemPolicy::Process marks each root it dispatches; roots 2 and 3 never are.
            task.MarkSourceItemMutationPossible(0u);
            task.MarkSourceItemMutationPossible(1u);
        }

        FileOperations::FileOperationItemResult finished{};
        finished.sourceIndex       = 0u;
        finished.finalSourcePath   = task._sourcePaths[0].native();
        finished.publication       = operation == FILESYSTEM_DELETE ? PublicationState::NotAttempted : PublicationState::Published;
        finished.sourceDisposition = operation == FILESYSTEM_COPY ? SourceDisposition::Retained : SourceDisposition::Removed;
        finished.completion        = ItemCompletion::Completed;
        finished.verification      = FileOperations::VerificationState::NotApplicable;
        finished.status            = S_OK;
        FileOperations::FileOperationItemResult failed{};
        failed.sourceIndex       = 1u;
        failed.finalSourcePath   = task._sourcePaths[1].native();
        failed.publication       = operation == FILESYSTEM_DELETE ? PublicationState::NotAttempted : PublicationState::NotPublished;
        failed.sourceDisposition = SourceDisposition::Retained;
        failed.completion        = ItemCompletion::Failed;
        failed.verification      = FileOperations::VerificationState::NotApplicable;
        failed.status            = accessDenied;
        const bool failure       = stop == Stop::FailureCancelTask || stop == Stop::FailureSequential;
        const bool stored        = stop == Stop::SetupFailure || (task.StoreTypedItemResult(finished) && (! failure || task.StoreTypedItemResult(failed)));

        if (stop == Stop::UserCancel || stop == Stop::FailureCancelTask)
            task.RequestCancel();
        const HRESULT finalHr = task.FinalizeTypedItemResults(stop == Stop::UserCancel ? canceled : accessDenied);

        const PublicationState unknownPublication = operation == FILESYSTEM_DELETE ? PublicationState::NotAttempted : PublicationState::Unknown;
        const SourceDisposition unknownSource     = operation == FILESYSTEM_COPY ? SourceDisposition::Retained : SourceDisposition::Unknown;
        const auto expectItem = [&](size_t index, PublicationState publication, SourceDisposition source, ItemCompletion completion, HRESULT status) noexcept
        {
            const std::optional<FileOperations::FileOperationItemResult>& item =
                index < task._sourceItemResultBuilders.size() ? task._sourceItemResultBuilders[index].terminal : std::nullopt;
            const bool matches = item.has_value() && item->publication == publication && item->sourceDisposition == source && item->completion == completion &&
                                 item->status == status;
            if (! matches)
                AppendLog(std::format(L"Mutation gate mismatch: op={} mode={} stop={} root={}; expected pub/source/completion={}/{}/{} hr=0x{:08X}; "
                                      L"actual={}/{}/{} hr=0x{:08X}",
                                      static_cast<unsigned>(operation),
                                      static_cast<unsigned>(mode),
                                      static_cast<unsigned>(stop),
                                      index,
                                      static_cast<unsigned>(publication),
                                      static_cast<unsigned>(source),
                                      static_cast<unsigned>(completion),
                                      static_cast<unsigned long>(status),
                                      item ? static_cast<unsigned>(item->publication) : UINT_MAX,
                                      item ? static_cast<unsigned>(item->sourceDisposition) : UINT_MAX,
                                      item ? static_cast<unsigned>(item->completion) : UINT_MAX,
                                      static_cast<unsigned long>(item ? item->status : E_PENDING)));
            return matches;
        };
        // One finished root plus canceled or failed roots make the task partial, never
        // indeterminate; a setup failure fails every root.
        const bool taskHrMatches = stop == Stop::SetupFailure ? FAILED(finalHr) && finalHr != HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE)
                                                              : finalHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (! stored || ! taskHrMatches)
            AppendLog(std::format(L"Mutation gate mismatch: op={} mode={} stop={} stored={} task hr=0x{:08X}",
                                  static_cast<unsigned>(operation),
                                  static_cast<unsigned>(mode),
                                  static_cast<unsigned>(stop),
                                  stored,
                                  static_cast<unsigned long>(finalHr)));
        if (stop == Stop::SetupFailure)
        {
            bool passed = stored && taskHrMatches;
            for (size_t index = 0u; passed && index < task._sourcePaths.size(); ++index)
                passed = expectItem(index, PublicationState::NotAttempted, SourceDisposition::Retained, ItemCompletion::Failed, accessDenied);
            return passed;
        }
        bool passed = stored && taskHrMatches && expectItem(0u, finished.publication, finished.sourceDisposition, ItemCompletion::Completed, S_OK) &&
                      (failure ? expectItem(1u, failed.publication, SourceDisposition::Retained, ItemCompletion::Failed, accessDenied)
                               : expectItem(1u, unknownPublication, unknownSource, ItemCompletion::Canceled, canceled));
        for (size_t index = 2u; passed && index < task._sourcePaths.size(); ++index)
        {
            // Never dispatched: another root's failure status is never this root's own.
            passed = perRoot ? expectItem(index, PublicationState::NotAttempted, SourceDisposition::Retained, ItemCompletion::Canceled, canceled)
                             : expectItem(index, unknownPublication, unknownSource, ItemCompletion::Canceled, canceled);
        }
        if (passed && perRoot && failure && operation != FILESYSTEM_DELETE)
        {
            // A root failure stopped the run, nobody canceled: the production completed summary names
            // the stop with its counts, never "Canceled" or plain success, and the never-reached
            // roots kept their sources without counting as a Move downgrade.
            task._taskId = 0x7E57'57A7'0000'0000ull + (static_cast<uint64_t>(operation) << 8u) + static_cast<uint64_t>(stop);
            task._resultHr.store(finalHr, std::memory_order_release);
            fileOperations.DebugRemoveDiagnosticsForTask(task._taskId);
            const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);
            fileOperations.DebugRemoveDiagnosticsForTask(task._taskId);
            const std::wstring expected =
                FormatStringResource(nullptr,
                                     operation == FILESYSTEM_COPY ? IDS_FMT_FILEOPS_RESULT_STOPPED_COPIED : IDS_FMT_FILEOPS_RESULT_STOPPED_MOVED,
                                     1ul,
                                     static_cast<unsigned long>(task._sourcePaths.size()));
            passed = summary.resultSummary == expected && summary.indeterminateItemCount == 0u && summary.canceledRetainedSourceCount == 2u &&
                     summary.retainedSourceCount == (operation == FILESYSTEM_COPY ? 4u : 3u) &&
                     ! IsAutoDismissableFileOperationCompletion(summary.resultHr, summary.warningCount, summary.errorCount);
            if (! passed)
                AppendLog(std::format(L"Mutation gate summary mismatch: op={} stop={} summary='{}' expected='{}' indeterminate={} retained={} "
                                      L"canceledRetained={} warnings={} errors={}",
                                      static_cast<unsigned>(operation),
                                      static_cast<unsigned>(stop),
                                      summary.resultSummary,
                                      expected,
                                      summary.indeterminateItemCount,
                                      summary.retainedSourceCount,
                                      summary.canceledRetainedSourceCount,
                                      summary.warningCount,
                                      summary.errorCount));
        }
        return passed;
    };

    for (const FileSystemOperation operation : {FILESYSTEM_COPY, FILESYSTEM_MOVE, FILESYSTEM_DELETE})
    {
        for (const Stop stop : {Stop::UserCancel, Stop::FailureCancelTask, Stop::FailureSequential, Stop::SetupFailure})
        {
            if (! run(operation, ExecutionMode::PerItem, stop))
                return false;
        }
    }
    return run(FILESYSTEM_DELETE, ExecutionMode::BulkItems, Stop::UserCancel) && run(FILESYSTEM_RENAME, ExecutionMode::PerItem, Stop::UserCancel);
}

// I30 model#6 (owner default): a dispatched root that consumed a descendant Skip without a receipt
// keeps its unknown axes, but the Skip is its terminal cause. Like a Canceled item, it does not make
// the task "Outcome unknown": the task completes with skipped items (S_FALSE).
[[nodiscard]] bool RunFileOperationSkippedUnknownAxesTaskSelfTest(FolderWindow::FileOperationState& fileOperations) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;
    using FileOperationItemResultReducer::ExecutionPhase;
    using FileOperationItemResultReducer::ReceiptState;
    for (const FileSystemOperation operation : {FILESYSTEM_MOVE, FILESYSTEM_DELETE})
    {
        auto taskStorage    = std::make_unique<Task>(fileOperations);
        Task& task          = *taskStorage;
        task._operation     = operation;
        task._executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        task._sourcePaths   = {std::filesystem::path(L"C:\\fileops-skipped-unknown-axes.bin")};
        task.InitializeSourceItemResultBuilders();
        task.EnterSourceItemsMutationGate();
        task.MarkSourceItemMutationPossible(0u);

        // The qualified executor stores exactly this reduction for a partial child Skip with no receipt.
        FileOperationItemResultReducer::Input input{};
        input.operation      = operation;
        input.phase          = ExecutionPhase::MutationPossible;
        input.receiptState   = ReceiptState::Missing;
        input.status         = S_FALSE;
        const auto reduction = FileOperationItemResultReducer::ReduceProviderReceipt(input, false, true);
        FileOperations::FileOperationItemResult skipped{};
        skipped.sourceIndex       = 0u;
        skipped.finalSourcePath   = task._sourcePaths[0].native();
        skipped.publication       = reduction.publication;
        skipped.sourceDisposition = reduction.sourceDisposition;
        skipped.completion        = reduction.completion;
        skipped.status            = reduction.status;
        skipped.verification      = FileOperations::VerificationState::NotApplicable;
        const bool stored         = task.StoreTypedItemResult(skipped);
        const HRESULT finalHr     = task.FinalizeTypedItemResults(S_OK);

        const std::optional<FileOperations::FileOperationItemResult>& item = task._sourceItemResultBuilders.front().terminal;
        const bool axesKept =
            item.has_value() && item->completion == FileOperations::ItemCompletion::Skipped && item->status == S_FALSE &&
            item->sourceDisposition == FileOperations::SourceDisposition::Unknown &&
            item->publication == (operation == FILESYSTEM_DELETE ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::Unknown);

        task._taskId = 0x7E57'5C1F'0000'0000ull + static_cast<uint64_t>(operation);
        task._resultHr.store(finalHr, std::memory_order_release);
        fileOperations.DebugRemoveDiagnosticsForTask(task._taskId);
        const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);
        fileOperations.DebugRemoveDiagnosticsForTask(task._taskId);
        const bool summaryMatches = summary.resultHr == S_FALSE && summary.indeterminateItemCount == 0u && summary.unknownSourceCount == 1u &&
                                    summary.unknownPublicationItemCount == (operation == FILESYSTEM_DELETE ? 0u : 1u) &&
                                    summary.resultSummary == LoadStringResource(nullptr, IDS_FILEOPS_RESULT_SKIPPED);
        AppendLog(std::format(L"Skipped unknown-axes task: op={} stored={} final=0x{:08X} axesKept={} summaryHr=0x{:08X} indeterminate={} "
                              L"unknownSource={} unknownPublication={} summary='{}' {}",
                              static_cast<unsigned>(operation),
                              stored,
                              static_cast<unsigned long>(finalHr),
                              axesKept,
                              static_cast<unsigned long>(summary.resultHr),
                              summary.indeterminateItemCount,
                              summary.unknownSourceCount,
                              summary.unknownPublicationItemCount,
                              summary.resultSummary,
                              stored && finalHr == S_FALSE && axesKept && summaryMatches ? L"ok" : L"MISMATCH"));
        if (! stored || finalHr != S_FALSE || ! axesKept || ! summaryMatches)
            return false;
    }
    return true;
}

// I30 x-arch#3: a partial Move's task.result text is built from the typed source dispositions. A
// fully moved root next to a failed one must not be described as "source preserved".
[[nodiscard]] bool RunFileOperationPartialMoveResultMessageSelfTest(FolderWindow::FileOperationState& fileOperations) noexcept
{
    using Task          = FolderWindow::FileOperationState::Task;
    auto taskStorage    = std::make_unique<Task>(fileOperations);
    Task& task          = *taskStorage;
    task._operation     = FILESYSTEM_MOVE;
    task._executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
    task._sourcePaths   = {std::filesystem::path(L"C:\\fileops-partial-move-0.bin"), std::filesystem::path(L"C:\\fileops-partial-move-1.bin")};
    task.InitializeSourceItemResultBuilders();
    task.EnterSourceItemsMutationGate();
    task.MarkSourceItemMutationPossible(0u);
    task.MarkSourceItemMutationPossible(1u);

    FileOperations::FileOperationItemResult moved{};
    moved.sourceIndex       = 0u;
    moved.finalSourcePath   = task._sourcePaths[0].native();
    moved.publication       = FileOperations::PublicationState::Published;
    moved.sourceDisposition = FileOperations::SourceDisposition::Removed;
    moved.completion        = FileOperations::ItemCompletion::Completed;
    moved.verification      = FileOperations::VerificationState::NotRequested;
    moved.status            = S_OK;
    FileOperations::FileOperationItemResult failed{};
    failed.sourceIndex       = 1u;
    failed.finalSourcePath   = task._sourcePaths[1].native();
    failed.publication       = FileOperations::PublicationState::NotPublished;
    failed.sourceDisposition = FileOperations::SourceDisposition::Retained;
    failed.completion        = FileOperations::ItemCompletion::Failed;
    failed.verification      = FileOperations::VerificationState::NotRequested;
    failed.status            = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    const bool stored        = task.StoreTypedItemResult(moved) && task.StoreTypedItemResult(failed);
    const HRESULT finalHr    = task.FinalizeTypedItemResults(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
    const std::wstring text  = task.FormatPartialMoveResultMessage(1ul, 2ul, 4096u, 8192u);
    const bool matches       = stored && finalHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) && text.find(L"source preserved") == std::wstring::npos &&
                               text.find(L"moved=1, kept=1, unknown=0") != std::wstring::npos;
    AppendLog(std::format(L"Partial Move task.result text: stored={} final=0x{:08X} text='{}' {}",
                          stored,
                          static_cast<unsigned long>(finalHr),
                          text,
                          matches ? L"ok" : L"MISMATCH"));
    return matches;
}

// Prove two schedules without equating visible destination bytes with terminal completion. One
// worker finishes two roots before its blocker; two workers start directly at two blockers. The
// live builder witnesses prove the completed prefix and occupied slots before Cancel. An open
// parallel prompt can hold a published sibling's cancellation callback, so a timed settle cannot
// establish that sibling's completion. Both schedules must leave later roots untouched.
[[nodiscard]] bool RunFileOpsPerItemCancelUndispatchedRootsChecks(SelfTestState& state,
                                                                  const FileSystemOperation operation,
                                                                  const unsigned int concurrency) noexcept
{
    using Task                    = FolderWindow::FileOperationState::Task;
    using Summary                 = FolderWindow::FileOperationState::CompletedTaskSummary;
    const bool move               = operation == FILESYSTEM_MOVE;
    const std::wstring_view label = move ? L"Move" : L"Copy";
    constexpr size_t kRootCount   = 8u;
    const size_t firstBlocker     = concurrency == 1u ? 2u : 0u;
    const size_t lastBlocker      = firstBlocker + concurrency - 1u;
    const auto reject             = [&](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Per-item cancel {} workers={}: {}", label, concurrency, reason));
        return false;
    };

    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    wil::com_ptr<IFileSystemIO> io;
    if ((concurrency != 1u && concurrency != 2u) || (operation != FILESYSTEM_COPY && ! move) || ! folder || ! operations || ! state.fsLocal ||
        ! state.infoLocal || ! state.fsLocal.try_query_to(io.put()) || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"Local endpoint or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const HWND originalPopup = operations->GetPopupHwndForSelfTest();
    const auto root          = state.tempRoot / (move ? L"per-item-cancel-move" : L"per-item-cancel-copy");
    const auto sourceRoot    = root / L"source";
    const auto destination   = root / L"destination";

    std::string originalLocalConfiguration;
    if (! BackupPluginConfiguration(state.infoLocal.get(), originalLocalConfiguration))
        return reject(L"Local configuration backup unavailable");
    const bool localConfigDirty = state.localConfigDirty;
    const std::wstring localPluginId(kPluginIdLocal);
    const auto configEntry = g_settings.plugins.configurationByPluginId.find(localPluginId);
    std::optional<Common::Settings::JsonValue> originalSettingsConfiguration;
    if (configEntry != g_settings.plugins.configurationByPluginId.end())
        originalSettingsConfiguration = configEntry->second;
    const bool verify          = EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy;
    const bool autoDismiss     = operations->GetAutoDismissSuccess();
    const bool queueNewTasks   = operations->GetQueueNewTasks();
    const bool autoAccept      = HostGetAutoAcceptPrompts();
    const auto promptOverride  = HostGetTestPromptResultOverride();
    const auto restoreSettings = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), originalLocalConfiguration));
        state.localConfigDirty = localConfigDirty;
        if (originalSettingsConfiguration.has_value())
            g_settings.plugins.configurationByPluginId[localPluginId] = originalSettingsConfiguration.value();
        else
            g_settings.plugins.configurationByPluginId.erase(localPluginId);
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = verify;
        HostSetAutoAcceptPrompts(autoAccept);
        if (promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(promptOverride);
        operations->SetQueueNewTasks(queueNewTasks);
        operations->SetAutoDismissSuccess(autoDismiss);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
    });
    std::optional<uint64_t> taskId;
    bool passed = false;
    // Join the worker before callbacks, configuration or owned trees unwind; keep a failed tree.
    auto drain = wil::scope_exit([&]() noexcept
    {
        if (auto* live = taskId.has_value() ? operations->FindTask(taskId.value()) : nullptr)
            live->RequestCancel();
        const bool quiet = CancelOutstandingSelfTestOperations(operations);
        if (! quiet)
            operations->Shutdown();
        if (taskId.has_value())
        {
            operations->DismissCompletedTask(taskId.value());
            operations->DebugRemoveDiagnosticsForTask(taskId.value());
            state.completedTasks.erase(taskId.value());
        }
        if (! originalPopup && operations->GetPopupHwndForSelfTest())
            operations->DebugResetPopupForSelfTest();
        if (passed && quiet)
            static_cast<void>(SelfTest::RemoveAll(root));
        else
            state.retainTempRootOnFailure = true;
    });

    const char* configuration = concurrency == 1u ? R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1,"deleteMaxConcurrency":1})json"
                                                  : R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":2,"deleteMaxConcurrency":1})json";
    if (! SetPluginConfiguration(state.infoLocal.get(), configuration))
        return reject(L"explicit Local worker configuration unavailable");
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    HostSetAutoAcceptPrompts(false);
    HostClearTestPromptResultOverride();
    operations->SetAutoDismissSuccess(false);
    operations->SetQueueNewTasks(false);

    std::vector<std::filesystem::path> sources;
    sources.reserve(kRootCount);
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(sourceRoot) || ! RecreateEmptyDirectory(destination))
        return reject(L"owned source and destination folders unavailable");
    for (size_t index = 0u; index < kRootCount; ++index)
    {
        sources.push_back(sourceRoot / std::format(L"root-{}.bin", index));
        if (! WriteFilledTestFile(sources.back(), 4096u + index, static_cast<unsigned char>(0x40u + index)))
            return reject(L"source payload unavailable");
    }
    for (size_t index = firstBlocker; index <= lastBlocker; ++index)
    {
        if (! WriteFilledTestFile(destination / sources[index].filename(), 777u, 0xEEu))
            return reject(L"pre-existing destination blocker unavailable");
    }
    const auto hasBytes = [&](const std::filesystem::path& path, size_t bytes, unsigned char value) noexcept
    {
        std::string text;
        return ReadFileTextFsIo(io, path, text) && text == std::string(bytes, static_cast<char>(value));
    };
    const auto absent = [](const std::filesystem::path& path) noexcept
    {
        const DWORD attributes = GetFileAttributesW(path.c_str());
        const DWORD error      = attributes == INVALID_FILE_ATTRIBUTES ? GetLastError() : ERROR_SUCCESS;
        return attributes == INVALID_FILE_ATTRIBUTES && (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND);
    };

    taskId = StartFileOperationAndGetId(operations,
                                        operation,
                                        FolderWindow::Pane::Left,
                                        FolderWindow::Pane::Right,
                                        state.fsLocal,
                                        sources,
                                        destination,
                                        FILESYSTEM_FLAG_NONE,
                                        false,
                                        0u,
                                        FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                        false,
                                        state.fsLocal);
    if (! taskId.has_value())
        return reject(L"actual per-item admission failed");

    // Copy only the completed prefix while locked; no live builder or Task borrow crosses a pump.
    std::array<std::optional<FileOperations::FileOperationItemResult>, kRootCount> completedBeforeCancel;
    std::optional<Task::ConflictPromptState> prompt;
    const auto held = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        auto* live = operations->FindTask(taskId.value());
        prompt     = TryGetConflictPromptCopy(live);
        if (! live || ! prompt.has_value())
            return false;
        std::scoped_lock lock(live->_sourceItemStatusMutex);
        if (live->_sourceItemResultBuilders.size() != kRootCount)
            return false;
        for (size_t index = 0u; index < kRootCount; ++index)
        {
            const auto& builder = live->_sourceItemResultBuilders[index];
            if (index < firstBlocker)
            {
                if (! builder.terminal.has_value() || builder.terminal->completion != FileOperations::ItemCompletion::Completed ||
                    builder.terminal->status != S_OK)
                    return false;
            }
            else
            {
                const auto expectedPhase = index <= lastBlocker ? Task::SourceItemExecutionPhase::MutationPossible : Task::SourceItemExecutionPhase::Preparing;
                if (builder.phase != expectedPhase || builder.terminal.has_value())
                    return false;
            }
        }
        for (size_t index = 0u; index < firstBlocker; ++index)
            completedBeforeCancel[index] = live->_sourceItemResultBuilders[index].terminal;
        return true;
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(15)), .operationName = L"per-item cancel first conflict prompt"});
    if (! held.conditionMet || ! prompt.has_value() || prompt->bucket != Task::ConflictBucket::RegularFileExists ||
        ! PromptHasAction(prompt.value(), Task::ConflictAction::Cancel) ||
        (prompt->destinationPath != (destination / sources[firstBlocker].filename()).native() &&
         prompt->destinationPath != (destination / sources[lastBlocker].filename()).native()))
        return reject(L"completed-prefix, occupied-worker and untouched-root witnesses did not precede the blocker prompt");
    AppendLog(std::format(L"Per-item cancel {} workers={}: before Cancel completed={} held={} preparing={}",
                          label,
                          concurrency,
                          firstBlocker,
                          concurrency,
                          kRootCount - lastBlocker - 1u));
    auto* task = operations->FindTask(taskId.value());
    if (! task || ! task->SubmitConflictDecision(prompt->generation, Task::ConflictAction::Cancel, false))
        return reject(L"the Cancel answer at the conflict prompt was rejected");
    task = nullptr; // Never retain a Task borrow across the next message pump.

    const auto quiet = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        if (auto* live = operations->FindTask(taskId.value()); live && TryGetConflictPromptCopy(live).has_value())
            live->RequestCancel();
        return state.completedTasks.contains(taskId.value()) && ! operations->FindTask(taskId.value()) && ! operations->HasActiveOperations();
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(30)), .operationName = L"per-item cancel completion quiet point"});
    const auto completion = state.completedTasks.find(taskId.value());
    if (! quiet.conditionMet || completion == state.completedTasks.end())
        return reject(L"the canceled task did not complete");
    const auto& info = completion->second;
    for (size_t index = 0u; index < info.sourceItemResults.size(); ++index)
    {
        const auto& item = info.sourceItemResults[index];
        AppendLog(std::format(L"Per-item cancel {} root={} pub/source/completion={}/{}/{} hr=0x{:08X}",
                              label,
                              index,
                              item ? static_cast<unsigned>(item->publication) : UINT_MAX,
                              item ? static_cast<unsigned>(item->sourceDisposition) : UINT_MAX,
                              item ? static_cast<unsigned>(item->completion) : UINT_MAX,
                              static_cast<unsigned long>(item ? item->status : E_PENDING)));
    }
    // Roots that finished before the cut make the canceled task partial, never indeterminate.
    const HRESULT expectedStatus = HRESULT_FROM_WIN32(firstBlocker > 0u ? ERROR_PARTIAL_COPY : ERROR_CANCELLED);
    if (info.hr != expectedStatus || info.configuredMaxConcurrency != concurrency || info.sourceItemResults.size() != kRootCount)
        return reject(std::format(L"task hr=0x{:08X} concurrency={} roots={}; expected 0x{:08X}, {} workers and {} roots",
                                  static_cast<unsigned long>(info.hr),
                                  info.configuredMaxConcurrency,
                                  info.sourceItemResults.size(),
                                  static_cast<unsigned long>(expectedStatus),
                                  concurrency,
                                  kRootCount));

    size_t untouchedCount = 0u;
    for (size_t index = 0u; index < kRootCount; ++index)
    {
        const auto& item           = info.sourceItemResults[index];
        const auto destinationLeaf = destination / sources[index].filename();
        const unsigned char value  = static_cast<unsigned char>(0x40u + index);
        if (! item.has_value())
            return reject(std::format(L"root {} has no terminal result", index));
        bool rootPassed = false;
        if (index < firstBlocker)
        {
            // Dispatched and finished: the root keeps its own receipt.
            const auto& before = completedBeforeCancel[index];
            rootPassed = before.has_value() && before->completion == item->completion && before->status == item->status &&
                         before->publication == item->publication && before->sourceDisposition == item->sourceDisposition &&
                         before->verification == item->verification && before->ownedStageDisposition == item->ownedStageDisposition &&
                         item->completion == FileOperations::ItemCompletion::Completed && item->publication == FileOperations::PublicationState::Published &&
                         item->sourceDisposition == (move ? FileOperations::SourceDisposition::Removed : FileOperations::SourceDisposition::Retained) &&
                         hasBytes(destinationLeaf, 4096u + index, value) && (move ? absent(sources[index]) : hasBytes(sources[index], 4096u + index, value));
        }
        else if (index <= lastBlocker)
        {
            // Held at (or behind) the prompt: canceled with known storage axes.
            rootPassed = item->completion == FileOperations::ItemCompletion::Canceled && item->publication != FileOperations::PublicationState::Unknown &&
                         item->publication != FileOperations::PublicationState::Published &&
                         item->sourceDisposition == FileOperations::SourceDisposition::Retained && hasBytes(destinationLeaf, 777u, 0xEEu) &&
                         hasBytes(sources[index], 4096u + index, value);
        }
        else
        {
            // Never dispatched: untouched, its source retained and no destination written.
            rootPassed = item->completion == FileOperations::ItemCompletion::Canceled && item->publication == FileOperations::PublicationState::NotAttempted &&
                         item->sourceDisposition == FileOperations::SourceDisposition::Retained && item->status == HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
                         absent(destinationLeaf) && hasBytes(sources[index], 4096u + index, value);
            untouchedCount += rootPassed ? 1u : 0u;
        }
        if (! rootPassed)
            return reject(std::format(L"root {} result or storage contradicts its dispatch", index));
    }

    std::vector<Summary> summaries;
    operations->CollectCompletedTasks(summaries);
    const auto summary = std::ranges::find_if(summaries, [&](const Summary& candidate) noexcept { return candidate.taskId == taskId.value(); });
    if (summary == summaries.end())
        return reject(L"completed summary unavailable");
    const size_t retainedExpected = move ? kRootCount - firstBlocker : kRootCount;
    const bool retainedListed     = std::ranges::all_of(sources.begin() + static_cast<std::ptrdiff_t>(move ? firstBlocker : 0u),
                                                        sources.end(),
                                                        [&](const std::filesystem::path& source) noexcept
    { return std::ranges::find(summary->retainedSourcePaths, source) != summary->retainedSourcePaths.end(); });
    // A cancel after some roots finished names the cancel and its counts, never plain success.
    const std::wstring canceledSummary = firstBlocker > 0u
                                             ? FormatStringResource(nullptr,
                                                                    move ? IDS_FMT_FILEOPS_RESULT_CANCELED_MOVED : IDS_FMT_FILEOPS_RESULT_CANCELED_COPIED,
                                                                    static_cast<unsigned long>(firstBlocker),
                                                                    static_cast<unsigned long>(kRootCount))
                                             : std::wstring{};
    if (summary->resultSummary != canceledSummary || summary->unknownPublicationItemCount != 0u || summary->unknownSourceCount != 0u ||
        summary->indeterminateItemCount != 0u || ! summary->unknownSourcePaths.empty() || summary->publishedItemCount != firstBlocker ||
        summary->removedSourceCount != (move ? firstBlocker : 0u) || summary->retainedSourceCount != retainedExpected || ! retainedListed ||
        summary->canceledRetainedSourceCount != kRootCount - firstBlocker || untouchedCount != kRootCount - lastBlocker - 1u)
        return reject(std::format(L"summary '{}' (expected '{}') unknownPublication={} unknownSource={} indeterminate={} published={} removed={} retained={} "
                                  L"canceledRetained={} listed={}",
                                  summary->resultSummary,
                                  canceledSummary,
                                  summary->unknownPublicationItemCount,
                                  summary->unknownSourceCount,
                                  summary->indeterminateItemCount,
                                  summary->publishedItemCount,
                                  summary->removedSourceCount,
                                  summary->retainedSourceCount,
                                  summary->canceledRetainedSourceCount,
                                  retainedListed));

    AppendLog(std::format(L"Per-item cancel {}: task={} workers={} finished={} held={} untouched={} summary='{}'",
                          label,
                          taskId.value(),
                          info.configuredMaxConcurrency,
                          firstBlocker,
                          concurrency,
                          untouchedCount,
                          summary->resultSummary));
    passed = true;
    return true;
}
