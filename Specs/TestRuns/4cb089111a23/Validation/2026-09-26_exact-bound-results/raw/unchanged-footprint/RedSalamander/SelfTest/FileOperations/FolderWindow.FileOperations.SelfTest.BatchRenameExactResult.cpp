// Included by FolderWindow.FileOperations.SelfTest.cpp, not compiled separately.
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
        task->SubmitConflictDecision(action, false);
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
    const HRESULT expectedResultStatus = faultRound ? HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) : (skipRound ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK);
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
    NextStep(state, SelfTestState::Step::FileOps_CopyMergeIntoExistingFolder);
    return false;
}
