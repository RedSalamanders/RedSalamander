// Included by FolderWindow.FileOperations.Tests.Runner.cpp, not compiled separately.
// One real local batch-rename hop per round. The test pause is after the worker's
// final-slot preflight, so the destination appears only between preflight and mutation.

[[nodiscard]] HRESULT PopulateBatchRenameExactResultProviderFacts(IFileSystem* fileSystem, BatchRenameExecutionOp& operation) noexcept
{
    if (! fileSystem || operation.originalSource.empty() || operation.finalLeaf.empty())
        return E_INVALIDARG;

    const FileSystemPathIdentity identity = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem();
    std::wstring parentPath;
    std::wstring sourceLeaf;
    if (! TryGetFileSystemParentPath(identity, operation.originalSource.native(), parentPath) ||
        ! TryGetFileSystemLeafName(identity, operation.originalSource.native(), sourceLeaf))
        return E_INVALIDARG;
    const std::optional<std::wstring> parentKey = TryMakePathKey(identity, parentPath);
    if (! parentKey.has_value())
        return E_INVALIDARG;

    const auto source = FileSystemRouteContract::QueryChildNameContract(fileSystem, parentPath, sourceLeaf, FILESYSTEM_RENAME, L"builtin/file-system");
    const auto final  = FileSystemRouteContract::QueryChildNameContract(fileSystem, parentPath, operation.finalLeaf, FILESYSTEM_RENAME, L"builtin/file-system");
    const auto statusFor = [](const FileSystemRouteContract::ChildNameContractResult& result) noexcept
    {
        if (result.state == FileSystemRouteContract::QueryState::Available && result.nameStatus == FILESYSTEM_CHILD_NAME_INVALID &&
            FAILED(result.failureStatus))
            return result.failureStatus;
        return FAILED(result.status) ? result.status : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    };
    if (source.state != FileSystemRouteContract::QueryState::Available || source.nameStatus != FILESYSTEM_CHILD_NAME_VALID || source.collisionKey.empty())
        return statusFor(source);
    if (final.state != FileSystemRouteContract::QueryState::Available || final.nameStatus != FILESYSTEM_CHILD_NAME_VALID || final.joinedPath.empty() ||
        final.collisionKey.empty())
        return statusFor(final);

    operation.providerFinalPath          = std::filesystem::path(final.joinedPath);
    operation.providerParentKey          = parentKey.value();
    operation.providerSourceCollisionKey = source.collisionKey;
    operation.providerFinalCollisionKey  = final.collisionKey;
    return S_OK;
}

[[nodiscard]] TickResult TickBatchRenameExactResultCarrier(SelfTestState& state) noexcept
{
    if (state.step != SelfTestState::Step::FileOps_BatchRenameExactResultCarrier)
        return std::nullopt;

    using Task              = FolderWindow::FileOperationState::Task;
    const auto releasePause = []() noexcept
    {
        ReleaseFileOpsBatchRenameBeforeMutationPauseForSelfTest();
        SetFileOpsBatchRenameBeforeMutationPauseForSelfTest(false);
    };
    const auto failAndRelease = [&](std::wstring_view message) noexcept -> TickResult
    {
        // Fail immediately runs cleanup and waits for workers. Unblock the worker first.
        releasePause();
        Fail(message);
        return true;
    };

    if (! state.fileOps || ! state.folderWindow)
        return failAndRelease(L"Exact batch-rename result requires the local FileOps setup.");
    if (state.stepState == 0u)
        state.batchRenameExactPaneFileSystem = state.folderWindow->GetFileSystem(FolderWindow::Pane::Left);
    const wil::com_ptr<IFileSystem>& paneFileSystem = state.batchRenameExactPaneFileSystem;
    if (! paneFileSystem)
        return failAndRelease(L"Exact batch-rename result lost the left pane's filesystem instance.");
    constexpr std::array<std::wstring_view, 7> roundNames{{
        L"replace",
        L"skip",
        L"receipt-unknown-outcome",
        L"receipt-no-commit-no-original",
        L"receipt-committed-original-retained",
        L"receipt-invalid-boolean",
        L"receipt-short-prefix",
    }};
    constexpr std::array<FileOpsBatchRenameReceiptFaultForSelfTest, 7> receiptFaults{{
        FileOpsBatchRenameReceiptFaultForSelfTest::None,
        FileOpsBatchRenameReceiptFaultForSelfTest::None,
        FileOpsBatchRenameReceiptFaultForSelfTest::UnknownOutcome,
        FileOpsBatchRenameReceiptFaultForSelfTest::NoCommitNoOriginal,
        FileOpsBatchRenameReceiptFaultForSelfTest::CommittedOriginalRetained,
        FileOpsBatchRenameReceiptFaultForSelfTest::InvalidBoolean,
        FileOpsBatchRenameReceiptFaultForSelfTest::ShortPrefix,
    }};
    if (state.stepState >= 4u * roundNames.size())
        return failAndRelease(L"Exact batch-rename result state moved past all seven rounds.");

    const size_t round                     = state.stepState / 4u;
    const bool skipRound                   = round == 1u;
    const bool faultRound                  = round >= 2u;
    const uint32_t phase                   = state.stepState % 4u;
    const std::filesystem::path root       = state.tempRoot / L"batch-rename-exact-result" / std::wstring(roundNames[round]);
    const std::filesystem::path sourcePath = root / L"original.txt";
    const std::filesystem::path finalPath  = root / L"renamed.txt";
    const std::string sourceBytes          = std::format("source-round-{}-926", round);
    const std::string occupantBytes        = std::format("occupant-round-{}-178", round);
    const Task::ConflictAction action      = skipRound ? Task::ConflictAction::Skip : Task::ConflictAction::Overwrite;
    const uint64_t taskId                  = state.batchRenameExactCurrentTaskId;

    wil::com_ptr<IFileSystemIO> io;
    if (FAILED(paneFileSystem->QueryInterface(IID_PPV_ARGS(io.addressof()))) || ! io)
        return failAndRelease(L"Left-pane local provider did not supply file I/O for exact batch-rename fixture.");

    if (phase == 0u)
    {
        if (round == 0u)
        {
            state.batchRenameExactAutoAcceptBefore = HostGetAutoAcceptPrompts();
            HostResetPromptShutdown();
            HostSetAutoAcceptPrompts(false);
        }
        SetFileOpsBatchRenameReceiptFaultForSelfTest(FileOpsBatchRenameReceiptFaultForSelfTest::None);
        if (state.fileOps->HasActiveOperations())
            return failAndRelease(L"Exact batch-rename fixture requires no prior active operation.");
        std::error_code ec;
        std::filesystem::create_directories(root, ec);
        if (ec)
            return failAndRelease(L"Could not create owned exact batch-rename fixture directory.");
        const bool sourceAlreadyExists = std::filesystem::exists(sourcePath, ec);
        if (ec || sourceAlreadyExists)
            return failAndRelease(L"Exact batch-rename source path was not initially absent.");
        const bool finalAlreadyExists = std::filesystem::exists(finalPath, ec);
        if (ec || finalAlreadyExists || ! WriteFileTextFsIo(io, sourcePath, sourceBytes))
            return failAndRelease(L"Could not prepare an absent final slot and a distinct source file.");

        BatchRenameExecutionOp operation{};
        operation.originalSource              = sourcePath;
        operation.finalLeaf                   = finalPath.filename().native();
        const HRESULT factsHr                 = PopulateBatchRenameExactResultProviderFacts(paneFileSystem.get(), operation);
        const FileSystemPathIdentity identity = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem();
        if (factsHr != S_OK || ! EquivalentPath(identity, operation.providerFinalPath.native(), finalPath.native()))
            return failAndRelease(L"Local provider did not attest exact batch-rename source/final namespace facts.");

        std::vector<BatchRenameExecutionOp> operations;
        operations.push_back(std::move(operation));
        state.batchRenameExactPromptGeneration = 0u;
        SetFileOpsBatchRenameBeforeMutationPauseForSelfTest(true);
        uint64_t admittedTaskId = 0u;
        const HRESULT admitHr   = state.fileOps->AdmitBatchRename(FolderWindow::Pane::Left, paneFileSystem, std::move(operations), {}, {}, &admittedTaskId);
        if (admitHr != S_OK || admittedTaskId == 0u)
            return failAndRelease(L"Exact batch-rename operation did not admit a published task.");
        state.batchRenameExactCurrentTaskId = admittedTaskId;
        state.markerTick                    = GetTickCount64();
        ++state.stepState;
        return false;
    }

    if (taskId == 0u)
        return failAndRelease(L"Exact batch-rename round lost its published task ID.");

    if (phase == 1u)
    {
        if (! HasFileOpsBatchRenameBeforeMutationPauseEnteredForSelfTest())
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (state.completedTasks.contains(taskId) || (nowTick >= state.markerTick && nowTick - state.markerTick > SelfTest::ScaleTimeout(10'000ull)))
                return failAndRelease(L"Batch-rename worker failed to pause after successful final-slot preflight.");
            return false;
        }
        std::error_code ec;
        const bool sourceExists = std::filesystem::exists(sourcePath, ec);
        if (ec)
            return failAndRelease(L"Could not inspect source before post-preflight race injection.");
        const bool finalExists = std::filesystem::exists(finalPath, ec);
        if (ec || ! sourceExists || finalExists || ! WriteFileTextFsIo(io, finalPath, occupantBytes))
            return failAndRelease(L"Post-preflight race did not start with source present and final slot absent.");
        std::string actualOccupant;
        if (! ReadFileTextFsIo(io, finalPath, actualOccupant) || actualOccupant != occupantBytes)
            return failAndRelease(L"Post-preflight race occupant bytes were not retained before worker release.");
        releasePause();
        ++state.stepState;
        return false;
    }

    if (phase == 2u)
    {
        Task* task        = state.fileOps->FindTask(taskId);
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (state.completedTasks.contains(taskId) || (nowTick >= state.markerTick && nowTick - state.markerTick > SelfTest::ScaleTimeout(30'000ull)))
                return failAndRelease(L"Batch-rename race completed without an actionable exact collision prompt.");
            return false;
        }
        if (! prompt->sourceMetadata.available || ! prompt->destinationMetadata.available)
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (nowTick >= state.markerTick && nowTick - state.markerTick > SelfTest::ScaleTimeout(30'000ull))
                return failAndRelease(L"Batch-rename collision prompt metadata did not become actionable.");
            return false;
        }
        const FileSystemPathIdentity identity = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem();
        if (prompt->generation == 0u || ! EquivalentPath(identity, prompt->sourcePath, sourcePath.native()) ||
            ! EquivalentPath(identity, prompt->destinationPath, finalPath.native()) || prompt->status != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) ||
            ! PromptHasAction(prompt.value(), action))
            return failAndRelease(L"Batch-rename collision prompt lacked exact generation, paths, metadata or action.");

        state.batchRenameExactPromptGeneration = prompt->generation;
        if (faultRound)
            SetFileOpsBatchRenameReceiptFaultForSelfTest(receiptFaults[round]);
        task->DebugSubmitCurrentConflictDecision(action, false);
        ++state.stepState;
        return false;
    }

    const auto completed = state.completedTasks.find(taskId);
    if (completed == state.completedTasks.end())
    {
        if (const auto prompt = TryGetConflictPromptCopy(state.fileOps->FindTask(taskId));
            prompt.has_value() && prompt->generation != state.batchRenameExactPromptGeneration)
            return failAndRelease(L"Exact batch-rename mutation surfaced a second conflict prompt.");
        const ULONGLONG nowTick = GetTickCount64();
        if (nowTick >= state.markerTick && nowTick - state.markerTick > SelfTest::ScaleTimeout(60'000ull))
            return failAndRelease(L"Exact batch-rename task did not complete after conflict decision.");
        return false;
    }
    if (Task* task = state.fileOps->FindTask(taskId); TryGetConflictPromptCopy(task).has_value())
        return failAndRelease(L"Exact batch-rename task completed with another unresolved prompt.");
    const HRESULT expectedTaskHr = faultRound ? HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) : (skipRound ? S_FALSE : S_OK);
    if (completed->second.hr != expectedTaskHr)
        return failAndRelease(L"Exact batch-rename task returned the wrong Replace/Skip aggregate status.");

    std::error_code ec;
    const bool sourceExists = std::filesystem::exists(sourcePath, ec);
    if (ec)
        return failAndRelease(L"Could not inspect source after exact batch-rename completion.");
    const bool finalExists = std::filesystem::exists(finalPath, ec);
    std::string actualSource;
    std::string actualFinal;
    if (ec || ! finalExists || ! ReadFileTextFsIo(io, finalPath, actualFinal) || (! skipRound && (sourceExists || actualFinal != sourceBytes)) ||
        (skipRound && (! sourceExists || ! ReadFileTextFsIo(io, sourcePath, actualSource) || actualSource != sourceBytes || actualFinal != occupantBytes)))
        return failAndRelease(L"Exact batch-rename Replace/Skip disk bytes or source retention were wrong.");

    std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> diagnostics;
    state.fileOps->CollectDiagnostics(diagnostics);
    std::optional<FileOperations::AcceptedConflictDecision> accepted;
    std::optional<FileOperations::ConflictResultObservation> result;
    size_t acceptedCount = 0u;
    size_t resultCount   = 0u;
    for (const auto& entry : diagnostics)
    {
        if (entry.taskId != taskId)
            continue;
        if (entry.acceptedDecision.has_value())
        {
            ++acceptedCount;
            accepted = entry.acceptedDecision;
        }
        if (entry.category == L"item.conflict.result")
        {
            ++resultCount;
            result = entry.conflictResult;
        }
    }
    const FileOperations::PublicationState expectedPublication =
        faultRound ? FileOperations::PublicationState::Unknown
                   : (skipRound ? FileOperations::PublicationState::NotPublished : FileOperations::PublicationState::Published);
    const FileOperations::SourceDisposition expectedSourceDisposition =
        faultRound ? FileOperations::SourceDisposition::Unknown
                   : (skipRound ? FileOperations::SourceDisposition::Retained : FileOperations::SourceDisposition::Removed);
    const FileOperations::ItemCompletion expectedCompletion =
        faultRound ? FileOperations::ItemCompletion::Indeterminate
                   : (skipRound ? FileOperations::ItemCompletion::Skipped : FileOperations::ItemCompletion::Completed);
    // An exact user Skip is S_FALSE on the typed item (I30 model#4), as for every other operation.
    const HRESULT expectedResultStatus = faultRound ? HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) : (skipRound ? S_FALSE : S_OK);
    if (acceptedCount != 1u || resultCount != 1u || ! accepted.has_value() || ! result.has_value() || accepted->taskId != taskId ||
        accepted->promptGeneration != state.batchRenameExactPromptGeneration || accepted->promptGeneration == 0u || accepted->action != action ||
        accepted->applyToAll || result->reference.taskId != accepted->taskId || result->reference.promptGeneration != accepted->promptGeneration ||
        result->publication != expectedPublication || result->sourceDisposition != expectedSourceDisposition || result->completion != expectedCompletion ||
        result->verification != FileOperations::VerificationState::NotApplicable || result->status != expectedResultStatus)
        return failAndRelease(L"Exact batch-rename result lacked the accepted prompt reference or truthful outcome axes.");

    if (faultRound)
    {
        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        const auto summary = std::ranges::find_if(summaries, [&](const auto& value) noexcept { return value.taskId == taskId; });
        if (summary == summaries.end() || summary->resultHr != HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) || summary->publishedItemCount != 0u ||
            summary->unknownPublicationItemCount != 1u || summary->removedSourceCount != 0u || summary->retainedSourceCount != 0u ||
            summary->unknownSourceCount != 1u || summary->indeterminateItemCount != 1u)
            return failAndRelease(L"Faulted receipt must reach completed history as unknown publication/source and indeterminate result.");

        FileOperationsPopupInternal::TaskSnapshot popup{};
        if (! TryGetPopupTaskSnapshot(state.fileOps, taskId, popup) || ! popup.finished || popup.resultHr != HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) ||
            ! popup.hasIndeterminateResult || popup.publishedItemCount != 0u || popup.unknownPublicationItemCount != 1u || popup.removedSourceCount != 0u ||
            popup.retainedSourceCount != 0u || popup.unknownSourceCount != 1u)
            return failAndRelease(L"Faulted receipt must project unknown terminal facts into the actual popup snapshot.");
    }

    std::filesystem::path reportPath;
    if (! state.fileOps->ExportTaskIssuesReport(taskId, &reportPath, false))
        return failAndRelease(L"Exact batch-rename issue report export failed.");
    wil::unique_handle report(CreateFileW(
        reportPath.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    LARGE_INTEGER reportSize{};
    if (! report || ! GetFileSizeEx(report.get(), &reportSize) || reportSize.QuadPart <= 0 || reportSize.QuadPart > 1024 * 1024 ||
        reportSize.QuadPart % sizeof(wchar_t) != 0)
        return failAndRelease(L"Exact batch-rename issue report was not bounded UTF-16.");
    std::wstring reportText(static_cast<size_t>(reportSize.QuadPart) / sizeof(wchar_t), L'\0');
    DWORD bytesRead = 0u;
    if (! ReadFile(report.get(), reportText.data(), static_cast<DWORD>(reportSize.QuadPart), &bytesRead, nullptr) ||
        bytesRead != static_cast<DWORD>(reportSize.QuadPart) ||
        reportText.find(FileOperations::FormatAcceptedConflictDecisionDiagnostic(accepted.value())) == std::wstring::npos ||
        reportText.find(FileOperations::FormatConflictResultDiagnostic(result.value())) == std::wstring::npos)
        return failAndRelease(L"Exact batch-rename report omitted the accepted decision or mutation result.");

    if (round + 1u < roundNames.size())
    {
        state.batchRenameExactCurrentTaskId = 0u;
        state.stepState                     = static_cast<uint32_t>((round + 1u) * 4u);
        return false;
    }
    releasePause();
    SetFileOpsBatchRenameReceiptFaultForSelfTest(FileOpsBatchRenameReceiptFaultForSelfTest::None);
    if (state.batchRenameExactAutoAcceptBefore.has_value())
    {
        HostResetPromptShutdown();
        HostSetAutoAcceptPrompts(state.batchRenameExactAutoAcceptBefore.value());
        state.batchRenameExactAutoAcceptBefore.reset();
    }
    state.batchRenameExactCurrentTaskId = 0u;
    state.batchRenameExactPaneFileSystem.reset();
    state.batchRenameExactPromptGeneration = 0u;
    RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
    NextStep(state, SelfTestState::Step::FileOps_BatchRenameNeverRunRowsNotAttempted);
    return false;
}

// I30 state-c#1. One real local Batch Rename of a.txt -> b.txt and b.txt -> c.txt per round.
// a.txt -> b.txt waits for b.txt to be vacated, so it is scheduled in the second layer. A c.txt
// occupant appears after the worker's preflight, and the answer at its prompt (Replace with an
// unknown receipt, Skip or Cancel) stops the run before a.txt is touched. The row the run never
// reached must be NotAttempted with its source retained, never the stopping row's outcome.
[[nodiscard]] bool RunFileOpsBatchRenameNeverRunRowChecks(SelfTestState& state) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;
    struct Round final
    {
        std::wstring_view name;
        Task::ConflictAction action                       = Task::ConflictAction::Cancel;
        FileOpsBatchRenameReceiptFaultForSelfTest fault   = FileOpsBatchRenameReceiptFaultForSelfTest::None;
        FileOperations::ItemCompletion stopperCompletion  = FileOperations::ItemCompletion::Failed;
        FileOperations::ItemCompletion neverRunCompletion = FileOperations::ItemCompletion::Failed;
        HRESULT neverRunStatus                            = E_PENDING;
    };
    const std::array<Round, 3> rounds{{
        {L"unknown-receipt",
         Task::ConflictAction::Overwrite,
         FileOpsBatchRenameReceiptFaultForSelfTest::UnknownOutcome,
         FileOperations::ItemCompletion::Indeterminate,
         FileOperations::ItemCompletion::Failed,
         HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED)},
        {L"skip",
         Task::ConflictAction::Skip,
         FileOpsBatchRenameReceiptFaultForSelfTest::None,
         FileOperations::ItemCompletion::Skipped,
         FileOperations::ItemCompletion::Failed,
         HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED)},
        {L"cancel",
         Task::ConflictAction::Cancel,
         FileOpsBatchRenameReceiptFaultForSelfTest::None,
         FileOperations::ItemCompletion::Canceled,
         FileOperations::ItemCompletion::Canceled,
         HRESULT_FROM_WIN32(ERROR_CANCELLED)},
    }};

    auto* operations = state.fileOps;
    if (! operations || ! state.folderWindow || ! CancelOutstandingSelfTestOperations(operations))
    {
        AppendLog(L"Batch Rename never-run rows: local FileOps setup or initial worker quiet point unavailable");
        return false;
    }
    const wil::com_ptr<IFileSystem> paneFileSystem = state.folderWindow->GetFileSystem(FolderWindow::Pane::Left);
    wil::com_ptr<IFileSystemIO> io;
    if (! paneFileSystem || FAILED(paneFileSystem->QueryInterface(IID_PPV_ARGS(io.addressof()))) || ! io)
    {
        AppendLog(L"Batch Rename never-run rows: left-pane local provider or file I/O unavailable");
        return false;
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
    {
        AppendLog(L"Batch Rename never-run rows: no-activation guard unavailable");
        return false;
    }
    const bool autoAccept  = HostGetAutoAcceptPrompts();
    const bool autoDismiss = operations->GetAutoDismissSuccess();
    const auto restore     = wil::scope_exit([&]() noexcept
    {
        HostResetPromptShutdown();
        HostSetAutoAcceptPrompts(autoAccept);
        operations->SetAutoDismissSuccess(autoDismiss);
    });
    HostResetPromptShutdown();
    HostSetAutoAcceptPrompts(false);
    operations->SetAutoDismissSuccess(false);

    const FileSystemPathIdentity identity = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem();
    for (size_t roundIndex = 0u; roundIndex < rounds.size(); ++roundIndex)
    {
        const Round& round = rounds[roundIndex];
        const auto reject  = [&](std::wstring_view reason) noexcept
        {
            AppendLog(std::format(L"Batch Rename never-run rows {}: {}", round.name, reason));
            return false;
        };
        const std::filesystem::path root  = state.tempRoot / L"batch-rename-never-run" / std::wstring(round.name);
        const std::filesystem::path aPath = root / L"a.txt";
        const std::filesystem::path bPath = root / L"b.txt";
        const std::filesystem::path cPath = root / L"c.txt";
        const std::string aBytes          = std::format("alpha-{}-311", roundIndex);
        const std::string bBytes          = std::format("bravo-{}-542", roundIndex);
        const std::string occupantBytes   = std::format("occupant-{}-907", roundIndex);

        std::optional<uint64_t> taskId;
        bool passed = false;
        // Release the worker before draining; join it before the owned tree unwinds; keep a failed tree.
        auto drain = wil::scope_exit([&]() noexcept
        {
            ReleaseFileOpsBatchRenameBeforeMutationPauseForSelfTest();
            SetFileOpsBatchRenameBeforeMutationPauseForSelfTest(false);
            SetFileOpsBatchRenameReceiptFaultForSelfTest(FileOpsBatchRenameReceiptFaultForSelfTest::None);
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
            if (passed && quiet)
                static_cast<void>(SelfTest::RemoveAll(root));
            else
                state.retainTempRootOnFailure = true;
        });

        std::error_code ec;
        std::filesystem::create_directories(root, ec);
        if (ec || std::filesystem::exists(aPath, ec) || std::filesystem::exists(bPath, ec) || std::filesystem::exists(cPath, ec) || ec)
            return reject(L"owned fixture folder unavailable or not initially empty");
        if (! WriteFileTextFsIo(io, aPath, aBytes) || ! WriteFileTextFsIo(io, bPath, bBytes))
            return reject(L"source files unavailable");

        std::vector<BatchRenameExecutionOp> renames(2u);
        renames[0].originalSource = aPath;
        renames[0].finalLeaf      = bPath.filename().native();
        renames[1].originalSource = bPath;
        renames[1].finalLeaf      = cPath.filename().native();
        for (BatchRenameExecutionOp& rename : renames)
        {
            if (PopulateBatchRenameExactResultProviderFacts(paneFileSystem.get(), rename) != S_OK)
                return reject(L"local provider did not attest the source/final namespace facts");
        }

        SetFileOpsBatchRenameBeforeMutationPauseForSelfTest(true);
        uint64_t admittedTaskId = 0u;
        const HRESULT admitHr   = operations->AdmitBatchRename(FolderWindow::Pane::Left, paneFileSystem, std::move(renames), {}, {}, &admittedTaskId);
        if (admitHr != S_OK || admittedTaskId == 0u)
            return reject(std::format(L"admission failed hr=0x{:08X}", static_cast<unsigned long>(admitHr)));
        taskId = admittedTaskId;

        const auto paused = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return HasFileOpsBatchRenameBeforeMutationPauseEnteredForSelfTest() || state.completedTasks.contains(admittedTaskId);
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"batch rename never-run preflight pause"});
        if (! paused.conditionMet || ! HasFileOpsBatchRenameBeforeMutationPauseEnteredForSelfTest())
            return reject(L"the worker did not pause after its final-slot preflight");
        if (! WriteFileTextFsIo(io, cPath, occupantBytes))
            return reject(L"the post-preflight c.txt occupant could not be written");
        ReleaseFileOpsBatchRenameBeforeMutationPauseForSelfTest();
        SetFileOpsBatchRenameBeforeMutationPauseForSelfTest(false);

        std::optional<Task::ConflictPromptState> prompt;
        const auto prompted = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            prompt = TryGetConflictPromptCopy(operations->FindTask(admittedTaskId));
            return state.completedTasks.contains(admittedTaskId) ||
                   (prompt.has_value() && prompt->sourceMetadata.available && prompt->destinationMetadata.available);
        },
            {.timeout = SelfTest::Scale(std::chrono::seconds(30)), .operationName = L"batch rename never-run first-layer prompt"});
        if (! prompted.conditionMet || ! prompt.has_value() || prompt->generation == 0u || ! EquivalentPath(identity, prompt->sourcePath, bPath.native()) ||
            ! EquivalentPath(identity, prompt->destinationPath, cPath.native()) || prompt->status != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) ||
            ! PromptHasAction(prompt.value(), round.action))
            return reject(L"the first layer (b.txt -> c.txt) did not raise an actionable exact collision prompt");

        SetFileOpsBatchRenameReceiptFaultForSelfTest(round.fault);
        auto* task = operations->FindTask(admittedTaskId);
        if (! task || ! task->SubmitConflictDecision(prompt->generation, round.action, false))
            return reject(L"the answer at the first-layer prompt was rejected");
        task = nullptr; // Never retain a Task borrow across the next message pump.

        const auto finished   = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return state.completedTasks.contains(admittedTaskId) && ! operations->HasActiveOperations();
        }, {.timeout = SelfTest::Scale(std::chrono::seconds(60)), .operationName = L"batch rename never-run completion"});
        const auto completion = state.completedTasks.find(admittedTaskId);
        if (! finished.conditionMet || completion == state.completedTasks.end())
            return reject(L"the task did not complete after the first-layer answer");
        const auto& info = completion->second;

        const FileOperations::FileOperationItemResult* neverRun = nullptr;
        const FileOperations::FileOperationItemResult* stopper  = nullptr;
        for (const auto& item : info.sourceItemResults)
        {
            if (! item.has_value())
                continue;
            AppendLog(std::format(L"Batch Rename never-run rows {} source={} pub/source/completion={}/{}/{} hr=0x{:08X}",
                                  round.name,
                                  item->finalSourcePath,
                                  static_cast<unsigned>(item->publication),
                                  static_cast<unsigned>(item->sourceDisposition),
                                  static_cast<unsigned>(item->completion),
                                  static_cast<unsigned long>(item->status)));
            if (EquivalentPath(identity, item->finalSourcePath, aPath.native()))
                neverRun = &item.value();
            else if (EquivalentPath(identity, item->finalSourcePath, bPath.native()))
                stopper = &item.value();
        }
        AppendLog(std::format(L"Batch Rename never-run rows {} task hr=0x{:08X}", round.name, static_cast<unsigned long>(info.hr)));
        if (info.sourceItemResults.size() != 2u || ! neverRun || ! stopper)
            return reject(L"the task did not report one terminal result for each of its two rows");
        if (stopper->completion != round.stopperCompletion)
            return reject(L"the first-layer row did not keep its own outcome");
        if (neverRun->publication != FileOperations::PublicationState::NotAttempted ||
            neverRun->sourceDisposition != FileOperations::SourceDisposition::Retained || neverRun->completion != round.neverRunCompletion ||
            neverRun->status != round.neverRunStatus)
            return reject(L"the second-layer row the run never reached was not NotAttempted/Retained with its own status");
        // The task result still names what stopped the run; only the never-run row stops inheriting it.
        const bool unknownStop = round.fault == FileOpsBatchRenameReceiptFaultForSelfTest::UnknownOutcome;
        if (unknownStop ? info.hr != HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) : (info.hr == S_OK || info.hr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE)))
            return reject(L"the task result did not report the stopping row's outcome");

        std::string actualA;
        std::string actualB;
        std::string actualC;
        const bool replaced = round.action == Task::ConflictAction::Overwrite;
        if (! ReadFileTextFsIo(io, aPath, actualA) || actualA != aBytes || ! ReadFileTextFsIo(io, cPath, actualC) ||
            actualC != (replaced ? bBytes : occupantBytes) ||
            (replaced ? std::filesystem::exists(bPath, ec) : (! ReadFileTextFsIo(io, bPath, actualB) || actualB != bBytes)))
            return reject(L"disk state does not show a.txt untouched and only the first layer applied");
        passed = true;
    }
    return true;
}
