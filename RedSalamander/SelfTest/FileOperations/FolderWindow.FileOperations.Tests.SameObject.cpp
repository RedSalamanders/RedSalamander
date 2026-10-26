[[nodiscard]] bool RunFileOpsSameObjectBytePreservationChecks(SelfTestState& state) noexcept
{
    using State       = FolderWindow::FileOperationState;
    using Task        = State::Task;
    auto* folder      = TryGetFolderWindow(state.mainWindow);
    auto* operations  = TryGetFileOps(folder);
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Same-object byte preservation: {}", reason));
        return false;
    };
    if (! folder || ! operations || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner or initial worker quiet point unavailable");

    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground            = GetForegroundWindow();
    const HWND originalPopup         = operations->GetPopupHwndForSelfTest();
    const auto sourceFileSystem      = folder->GetFileSystem(FolderWindow::Pane::Left);
    const auto destinationFileSystem = folder->GetFileSystem(FolderWindow::Pane::Right);
    wil::com_ptr<IFileSystemIO> sourceIo;
    wil::com_ptr<IFileSystemIO> destinationIo;
    if (! sourceFileSystem || ! destinationFileSystem || ! sourceFileSystem.try_query_to(sourceIo.put()) ||
        ! destinationFileSystem.try_query_to(destinationIo.put()) || folder->GetFileSystemPluginId(FolderWindow::Pane::Left) != kPluginIdLocal ||
        folder->GetFileSystemPluginId(FolderWindow::Pane::Right) != kPluginIdLocal)
        return reject(L"actual Local pane filesystems unavailable");

    const bool autoDismiss     = operations->GetAutoDismissSuccess();
    const bool queueNewTasks   = operations->GetQueueNewTasks();
    const bool autoAccept      = HostGetAutoAcceptPrompts();
    const auto promptOverride  = HostGetTestPromptResultOverride();
    const auto restoreSettings = [&]() noexcept
    {
        HostSetAutoAcceptPrompts(autoAccept);
        if (promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(promptOverride);
        operations->SetQueueNewTasks(queueNewTasks);
        operations->SetAutoDismissSuccess(autoDismiss);
        return operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u);
    };
    auto restore = wil::scope_exit([&]() noexcept { static_cast<void>(restoreSettings()); });
    HostSetAutoAcceptPrompts(false);
    HostClearTestPromptResultOverride();
    operations->SetAutoDismissSuccess(false);
    operations->SetQueueNewTasks(false);

    const auto root = state.tempRoot / L"same-object-byte-preservation";
    if (! RecreateEmptyDirectory(root))
        return reject(L"owned root unavailable");
    std::vector<uint64_t> ownedTaskIds;
    const auto cleanup = [&]() noexcept
    {
        // No callback borrows this stack. Still drain normal workers before removing
        // the filesystem fixture or restoring prompt/settings policy.
        if (! CancelOutstandingSelfTestOperations(operations))
        {
            AppendLog(L"Same-object cleanup: worker quiet point failed; owned root retained.");
            return false;
        }
        for (const auto taskId : ownedTaskIds)
        {
            operations->DismissCompletedTask(taskId);
            operations->DebugRemoveDiagnosticsForTask(taskId);
            state.completedTasks.erase(taskId);
        }
        if (! originalPopup && operations->GetPopupHwndForSelfTest())
            operations->DebugResetPopupForSelfTest();
        return SelfTest::RemoveAll(root);
    };
    auto drain = wil::scope_exit([&]() noexcept { static_cast<void>(cleanup()); });

    constexpr FileSystemBindFlags bindFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_CONTENT | FILESYSTEM_BIND_READ_METADATA);
    // Distinct payloads prevent an empty or uniform shared fixture from hiding truncation.
    // All paths and reference bytes belong to this case, under the governed tempRoot.
    for (unsigned int scenario = 0u; scenario < 4u; ++scenario)
    {
        const bool hardLink     = scenario >= 2u;
        const bool samePathCopy = scenario == 0u;
        const auto operation    = scenario % 2u == 0u ? FILESYSTEM_COPY : FILESYSTEM_MOVE;
        const auto caseRoot     = root / std::format(L"case-{}", scenario);
        const auto source       = caseRoot / L"src" / L"source.bin";
        const auto destination  = hardLink ? caseRoot / L"dest" / L"source.bin" : source;
        const auto reference    = caseRoot / L"reference" / L"source.bin";
        std::wstring siblingText;
        if (samePathCopy && FAILED(Common::Paths::BuildUniqueSiblingPathCandidate(source.native(), false, 2u, siblingText)))
            return reject(L"canonical KeepBoth sibling name unavailable");
        const std::filesystem::path sibling(siblingText);
        const size_t payloadSize = 4099u + 257u * scenario;
        const auto payloadValue  = static_cast<unsigned char>(0xA7u - 0x31u * scenario);
        std::error_code hardLinkEc;
        if (! WriteFilledTestFile(source, payloadSize, payloadValue) || ! WriteFilledTestFile(reference, payloadSize, payloadValue) ||
            (hardLink &&
             (! SelfTest::EnsureDirectory(destination.parent_path()) || ! Common::Testing::CreateSandboxedHardLink(destination, source, hardLinkEc))))
            return reject(L"case-owned source/reference or hard-link alias unavailable");
        if (! FilesEqualBytes(source, reference) || ! FilesEqualBytes(destination, reference))
            return reject(L"before-operation exact byte references differ");

        FileOperations::QualifiedEndpoint sourceEndpoint;
        FileOperations::QualifiedEndpoint destinationEndpoint;
        if (! FileOperations::TryQualifyEndpointForSelfTest(
                sourceFileSystem, source.native(), operation, folder->GetFileSystemPluginId(FolderWindow::Pane::Left), L"host/default", sourceEndpoint) ||
            ! FileOperations::TryQualifyEndpointForSelfTest(destinationFileSystem,
                                                            destination.native(),
                                                            operation,
                                                            folder->GetFileSystemPluginId(FolderWindow::Pane::Right),
                                                            L"host/default",
                                                            destinationEndpoint) ||
            sourceEndpoint.profileId.empty() || destinationEndpoint.profileId.empty() || sourceEndpoint.profileId != destinationEndpoint.profileId)
            return reject(L"actual Local path profiles did not qualify");
        const auto sourceBefore = FileOperations::BindObjectAuthority(sourceFileSystem.get(), source.native(), sourceEndpoint.profileId, bindFlags);
        const auto destinationBefore =
            FileOperations::BindObjectAuthority(destinationFileSystem.get(), destination.native(), destinationEndpoint.profileId, bindFlags);
        bool sameObject   = false;
        bool sameRevision = false;
        if (sourceBefore.state != FileOperations::ObjectBindingState::Bound || destinationBefore.state != FileOperations::ObjectBindingState::Bound ||
            FAILED(sourceBefore.status) || FAILED(destinationBefore.status) || sourceBefore.authority.kind != FILESYSTEM_BOUND_REGULAR_FILE ||
            destinationBefore.authority.kind != FILESYSTEM_BOUND_REGULAR_FILE || sourceBefore.authority.identity.objectId.empty() ||
            sourceBefore.authority.identity.objectId != destinationBefore.authority.identity.objectId ||
            FAILED(FileOperations::CrossCheckBoundObjectIdentity(sourceBefore.authority, destinationBefore.authority, sameObject, sameRevision)) ||
            ! sameObject || ! sameRevision)
            return reject(L"provider-bound source/destination IDs do not prove the same file before admission");

        std::string originalBytes;
        if (! ReadFileTextFsIo(sourceIo, source, originalBytes) || originalBytes.size() != payloadSize ||
            ! std::ranges::all_of(originalBytes, [payloadValue](char value) noexcept { return static_cast<unsigned char>(value) == payloadValue; }))
            return reject(L"exact seeded bytes unavailable before operation");
        Common::Crypto::Blake3Hasher originalHasher;
        originalHasher.Update(std::as_bytes(std::span<const char>(originalBytes.data(), originalBytes.size())));
        const auto originalDigest = originalHasher.Finalize();

        // Inventory only the closed fixture's known directories, without traversing
        // any discovered child. Equality rejects stage siblings and extra subtrees.
        using InventoryEntry = std::pair<std::filesystem::path, std::filesystem::file_type>;
        const std::vector<std::filesystem::path> directories =
            hardLink ? std::vector<std::filesystem::path>{caseRoot, source.parent_path(), destination.parent_path(), reference.parent_path()}
                     : std::vector<std::filesystem::path>{caseRoot, source.parent_path(), reference.parent_path()};
        const auto inventory = [&](std::vector<InventoryEntry>& entries) noexcept
        {
            entries.clear();
            std::error_code ec;
            for (const auto& directory : directories)
            {
                const auto status = std::filesystem::symlink_status(directory, ec);
                if (ec || std::filesystem::is_symlink(status) || ! std::filesystem::is_directory(status))
                    return false;
                std::filesystem::directory_iterator entry(directory, std::filesystem::directory_options::none, ec);
                if (ec)
                    return false;
                for (; entry != std::filesystem::directory_iterator{}; entry.increment(ec))
                {
                    if (ec || entries.size() >= 6u)
                        return false;
                    const auto entryStatus = entry->symlink_status(ec);
                    if (ec || std::filesystem::is_symlink(entryStatus) ||
                        (! std::filesystem::is_directory(entryStatus) && ! std::filesystem::is_regular_file(entryStatus)))
                        return false;
                    entries.emplace_back(entry->path().lexically_relative(caseRoot), entryStatus.type());
                }
                if (ec)
                    return false;
            }
            std::ranges::sort(entries);
            return true;
        };
        std::vector<InventoryEntry> expectedInventory{
            {source.parent_path().lexically_relative(caseRoot), std::filesystem::file_type::directory},
            {reference.parent_path().lexically_relative(caseRoot), std::filesystem::file_type::directory},
            {source.lexically_relative(caseRoot), std::filesystem::file_type::regular},
            {reference.lexically_relative(caseRoot), std::filesystem::file_type::regular},
        };
        if (hardLink)
        {
            expectedInventory.emplace_back(destination.parent_path().lexically_relative(caseRoot), std::filesystem::file_type::directory);
            expectedInventory.emplace_back(destination.lexically_relative(caseRoot), std::filesystem::file_type::regular);
        }
        std::ranges::sort(expectedInventory);
        std::vector<InventoryEntry> beforeInventory;
        if (! inventory(beforeInventory) || beforeInventory != expectedInventory)
            return reject(L"before-operation closed inventory differs from the owned fixture");

        uint64_t taskId           = 0u;
        const HRESULT admissionHr = operations->AdmitOperation(operation,
                                                               FolderWindow::Pane::Left,
                                                               FolderWindow::Pane::Right,
                                                               sourceFileSystem,
                                                               {source},
                                                               destination.parent_path(),
                                                               FILESYSTEM_FLAG_NONE,
                                                               false,
                                                               0u,
                                                               State::ExecutionMode::PerItem,
                                                               false,
                                                               destinationFileSystem,
                                                               &taskId);
        if (taskId != 0u)
            ownedTaskIds.push_back(taskId);
        if (! hardLink && ! samePathCopy)
        {
            if (! FAILED(admissionHr) || taskId != 0u || operations->HasActiveOperations())
                return reject(L"same-folder Move was not refused before task publication");
        }
        else
        {
            if (admissionHr != S_OK || taskId == 0u)
                return reject(L"same-object Copy/Move must reach normal worker admission");
            if (auto* admittedTask = operations->FindTask(taskId))
            {
                // Keep even a sub-500ms successful Copy's actual completed summary.
                // This affects presentation retention only, never execution or naming.
                admittedTask->_presentationState.store(Task::TaskPresentationState::Presented, std::memory_order_release);
            }
            bool promptRejected         = false;
            uint64_t acceptedGeneration = 0u;
            const auto quiet            = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                auto* task        = operations->FindTask(taskId);
                const auto prompt = TryGetConflictPromptCopy(task);
                if (task && prompt.has_value() && prompt.value().generation != acceptedGeneration)
                {
                    // Consume only a current-generation Skip if the normal route
                    // surfaces a conflict. The terminal assertions below still
                    // require the established KeepBoth success or SameObject refusal.
                    if (! PromptHasAction(prompt.value(), Task::ConflictAction::Skip) ||
                        ! task->SubmitConflictDecision(prompt.value().generation, Task::ConflictAction::Skip, false))
                        promptRejected = true;
                    else
                        acceptedGeneration = prompt.value().generation;
                }
                return promptRejected || (! operations->HasActiveOperations() && ! operations->FindTask(taskId) && state.completedTasks.contains(taskId));
            },
                {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"same-object normal worker completion and quiet point"});
            if (! quiet.conditionMet || promptRejected || operations->HasActiveOperations() || operations->FindTask(taskId))
                return reject(L"normal conflict decision or bounded worker quiet point failed");
            const auto completed         = state.completedTasks.find(taskId);
            const HRESULT expectedStatus = samePathCopy ? S_OK : HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            if (completed == state.completedTasks.end() || completed->second.hr != expectedStatus || completed->second.sourceItemResults.size() != 1u ||
                ! completed->second.sourceItemResults.front().has_value())
                return reject(L"actual completed task did not retain its KeepBoth success or SameObject refusal and single typed result");
            const auto& item = completed->second.sourceItemResults.front().value();
            if (item.sourceIndex != 0u || item.status != expectedStatus ||
                item.completion != (samePathCopy ? FileOperations::ItemCompletion::Completed : FileOperations::ItemCompletion::Failed) ||
                item.publication != (samePathCopy ? FileOperations::PublicationState::Published : FileOperations::PublicationState::NotAttempted) ||
                item.sourceDisposition != FileOperations::SourceDisposition::Retained ||
                item.ownedStageDisposition !=
                    (samePathCopy ? FileOperations::OwnedStageDisposition::Published : FileOperations::OwnedStageDisposition::NotApplicable) ||
                item.finalSourcePath != source.native() || item.finalDestinationPath != (samePathCopy ? sibling.native() : destination.native()))
                return reject(L"typed result lost Published/Retained KeepBoth success or pre-mutation Failed/NotAttempted/Retained refusal truth");
            std::vector<State::TaskDiagnosticEntry> diagnostics;
            operations->CollectDiagnostics(diagnostics);
            if (! samePathCopy && ! std::ranges::any_of(diagnostics, [taskId](const auto& entry) noexcept {
                return entry.taskId == taskId && entry.category == L"identity.sameObject" && entry.status == HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }))
                return reject(L"normal worker did not record its actual SameObject refusal");
            std::vector<State::CompletedTaskSummary> summaries;
            operations->CollectCompletedTasks(summaries);
            const auto summary = std::ranges::find_if(summaries, [taskId](const auto& value) noexcept { return value.taskId == taskId; });
            if (summary == summaries.end() || summary->resultHr != completed->second.hr || summary->publishedItemCount != (samePathCopy ? 1u : 0u) ||
                summary->unknownPublicationItemCount != 0u || summary->removedSourceCount != 0u || summary->retainedSourceCount != 1u ||
                summary->unknownSourceCount != 0u || summary->retainedOwnedStageCount != 0u || summary->unknownOwnedStageCount != 0u ||
                summary->indeterminateItemCount != 0u || (! samePathCopy && summary->errorCount == 0u))
                return reject(L"actual completed summary lost KeepBoth success or SameObject refusal axes");
            if (! samePathCopy && summary->lastDiagnosticMessage != LoadStringResource(nullptr, IDS_FILEOPS_DIAG_SAME_OBJECT))
                return reject(L"actual SameObject refusal showed a path-syntax error instead of the same-item explanation");
        }

        auto expectedAfterInventory = beforeInventory;
        if (samePathCopy)
        {
            expectedAfterInventory.emplace_back(sibling.lexically_relative(caseRoot), std::filesystem::file_type::regular);
            std::ranges::sort(expectedAfterInventory);
            if (! FilesEqualBytes(sibling, reference))
                return reject(L"normal same-path Copy did not publish the exact canonical KeepBoth sibling bytes");
            const auto siblingBinding =
                FileOperations::BindObjectAuthority(destinationFileSystem.get(), sibling.native(), destinationEndpoint.profileId, bindFlags);
            bool siblingIsOriginal      = true;
            bool siblingHasSameRevision = false;
            if (siblingBinding.state != FileOperations::ObjectBindingState::Bound || FAILED(siblingBinding.status) ||
                FAILED(FileOperations::CrossCheckBoundObjectIdentity(
                    sourceBefore.authority, siblingBinding.authority, siblingIsOriginal, siblingHasSameRevision)) ||
                siblingIsOriginal || siblingBinding.authority.identity.objectId == sourceBefore.authority.identity.objectId)
                return reject(L"KeepBoth sibling did not bind as a distinct published file");
        }
        std::vector<InventoryEntry> afterInventory;
        if (! inventory(afterInventory) || afterInventory != expectedAfterInventory || ! FilesEqualBytes(source, reference) ||
            ! FilesEqualBytes(destination, reference))
            return reject(L"operation changed source/destination/reference bytes or left unexpected stage/KeepBoth siblings");
        std::vector<std::pair<std::filesystem::path, wil::com_ptr<IFileSystemIO>>> retainedPaths{{source, sourceIo}, {destination, destinationIo}};
        if (samePathCopy)
            retainedPaths.emplace_back(sibling, destinationIo);
        for (const auto& [path, io] : retainedPaths)
        {
            std::string retainedBytes;
            if (! ReadFileTextFsIo(io, path, retainedBytes) || retainedBytes != originalBytes)
                return reject(L"after-operation exact seeded bytes changed");
            Common::Crypto::Blake3Hasher retainedHasher;
            retainedHasher.Update(std::as_bytes(std::span<const char>(retainedBytes.data(), retainedBytes.size())));
            const auto retainedDigest = retainedHasher.Finalize();
            if (retainedDigest != originalDigest)
                return reject(L"after-operation canonical BLAKE3 changed");
            AppendLog(
                std::format(L"Same-object case={} operation={} hardLink={} taskId={} path={} bytes={} before BLAKE3={} after BLAKE3={}; exact bytes retained",
                            scenario,
                            static_cast<unsigned int>(operation),
                            hardLink,
                            taskId,
                            path.native(),
                            retainedBytes.size(),
                            Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(originalDigest)),
                            Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(retainedDigest))));
        }
        const auto sourceAfter =
            FileOperations::RevalidateObjectAuthority(sourceFileSystem.get(), source.native(), sourceEndpoint.profileId, bindFlags, sourceBefore.authority);
        const auto destinationAfter = FileOperations::RevalidateObjectAuthority(
            destinationFileSystem.get(), destination.native(), destinationEndpoint.profileId, bindFlags, destinationBefore.authority);
        sameObject   = false;
        sameRevision = false;
        if (sourceAfter.state != FileOperations::ObjectRevalidationState::Same || destinationAfter.state != FileOperations::ObjectRevalidationState::Same ||
            FAILED(sourceAfter.status) || FAILED(destinationAfter.status) ||
            sourceAfter.current.identity.objectId != sourceBefore.authority.identity.objectId ||
            destinationAfter.current.identity.objectId != destinationBefore.authority.identity.objectId ||
            FAILED(FileOperations::CrossCheckBoundObjectIdentity(sourceAfter.current, destinationAfter.current, sameObject, sameRevision)) || ! sameObject ||
            ! sameRevision || GetForegroundWindow() != foreground)
            return reject(L"after-operation bound IDs/revision or primary foreground changed");
        AppendLog(std::format(
            L"Same-object case={} admission=0x{:08X} taskId={} profile={}: original provider IDs unchanged; closed inventory exact; KeepBoth={}; worker quiet",
            scenario,
            static_cast<unsigned long>(admissionHr),
            taskId,
            sourceEndpoint.profileId,
            samePathCopy));
    }
    if (! cleanup())
        return reject(L"owned worker/fixture cleanup did not reach its quiet point");
    drain.release();
    if (! restoreSettings())
        return reject(L"restored settings did not reach their save quiet point");
    restore.release();
    return GetForegroundWindow() == foreground;
}
