// Included by the canonical FileOps selftest translation unit. This bounded
// fixture deliberately observes only its fixed directories, rejects every
// reparse entry, and closes observation handles before the exclusive lock.
// File IDs below are preservation witnesses, never mutation authority.
[[nodiscard]] bool RunFileOpsOverwritePreservationChecksImpl(SelfTestState& state)
{
    using State                         = FolderWindow::FileOperationState;
    using Task                          = State::Task;
    using Env                           = RedSalamander::TestSupport::ScopedEnvironmentVariable;
    constexpr std::wstring_view kSource = L"REDSALAMANDER_FILEOPS_STAGE_WRITE_FAIL_SOURCE";
    constexpr std::wstring_view kFired  = L"REDSALAMANDER_FILEOPS_STAGE_WRITE_FAIL_FIRED";
    constexpr std::wstring_view kStage  = L"REDSALAMANDER_FILEOPS_STAGE_WRITE_FAIL_STAGE";
    constexpr std::wstring_view kBytes  = L"REDSALAMANDER_FILEOPS_STAGE_WRITE_FAIL_BYTES";
    constexpr std::wstring_view kDigest = L"REDSALAMANDER_FILEOPS_STAGE_WRITE_FAIL_BLAKE3";
    const auto reject                   = [&](std::wstring_view reason) noexcept
    {
        state.retainTempRootOnFailure = true;
        AppendLog(std::format(L"D4 overwrite preservation: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground    = GetForegroundWindow();
    const HWND originalPopup = operations->GetPopupHwndForSelfTest();
    const auto sourceFs      = folder->GetFileSystem(FolderWindow::Pane::Left);
    const auto destinationFs = folder->GetFileSystem(FolderWindow::Pane::Right);
    wil::com_ptr<IFileSystemIO> sourceIo;
    wil::com_ptr<IFileSystemIO> destinationIo;
    if (! sourceFs || ! destinationFs || ! sourceFs.try_query_to(sourceIo.put()) || ! destinationFs.try_query_to(destinationIo.put()) ||
        folder->GetFileSystemPluginId(FolderWindow::Pane::Left) != kPluginIdLocal || folder->GetFileSystemPluginId(FolderWindow::Pane::Right) != kPluginIdLocal)
        return reject(L"actual Local endpoints unavailable");

    const bool autoDismiss     = operations->GetAutoDismissSuccess();
    const bool queueNewTasks   = operations->GetQueueNewTasks();
    const bool autoAccept      = HostGetAutoAcceptPrompts();
    const auto promptOverride  = HostGetTestPromptResultOverride();
    const bool verify          = EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy;
    const auto restoreSettings = [&]() noexcept
    {
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = verify;
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
    // This selects Local provider Copy and its overwrite stage loop. Requested
    // content verification would select a different host/bridge transfer route.
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;

    const auto root = state.tempRoot / L"overwrite-byte-preservation";
    if (! RecreateEmptyDirectory(root))
        return reject(L"owned root unavailable");
    std::vector<uint64_t> ownedTaskIds;
    const auto cleanup = [&]() noexcept
    {
        if (! CancelOutstandingSelfTestOperations(operations))
        {
            AppendLog(L"D4 cleanup: workers not quiet; owned fixture retained.");
            return false;
        }
        if (state.retainTempRootOnFailure)
        {
            AppendLog(std::format(L"D4 cleanup: workers quiet; failed witness diagnostics and owned tree retained: {}", root.native()));
            return true;
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

    struct Proof final
    {
        FILE_ID_INFO id{};
        std::string bytes;
        Common::Crypto::Blake3Digest digest{};
    };
    const auto capture = [&](const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, Proof& proof) noexcept
    {
        wil::unique_handle handle(CreateFileW(path.c_str(),
                                              FILE_READ_ATTRIBUTES,
                                              FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                              nullptr,
                                              OPEN_EXISTING,
                                              FILE_FLAG_OPEN_REPARSE_POINT,
                                              nullptr));
        FILE_ATTRIBUTE_TAG_INFO tag{};
        if (! handle || ! GetFileInformationByHandleEx(handle.get(), FileAttributeTagInfo, &tag, sizeof(tag)) ||
            (tag.FileAttributes & (FILE_ATTRIBUTE_REPARSE_POINT | FILE_ATTRIBUTE_DIRECTORY)) != 0u ||
            ! GetFileInformationByHandleEx(handle.get(), FileIdInfo, &proof.id, sizeof(proof.id)) || ! ReadFileTextFsIo(io, path, proof.bytes))
            return false;
        Common::Crypto::Blake3Hasher hasher;
        hasher.Update(std::as_bytes(std::span<const char>(proof.bytes.data(), proof.bytes.size())));
        proof.digest = hasher.Finalize();
        return true;
    };
    const auto sameProof = [](const Proof& left, const Proof& right) noexcept
    {
        return left.bytes == right.bytes && left.digest == right.digest && left.id.VolumeSerialNumber == right.id.VolumeSerialNumber &&
               std::memcmp(left.id.FileId.Identifier, right.id.FileId.Identifier, sizeof(left.id.FileId.Identifier)) == 0;
    };
    const auto logProof = [](unsigned int scenario, std::wstring_view phase, const std::filesystem::path& path, const Proof& proof) noexcept
    {
        AppendLog(std::format(L"D4 case={} phase={} path={} bytes={} volume={} FILE_ID_INFO={} BLAKE3={}",
                              scenario,
                              phase,
                              path.native(),
                              proof.bytes.size(),
                              proof.id.VolumeSerialNumber,
                              Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(std::as_bytes(std::span(proof.id.FileId.Identifier)))),
                              Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(proof.digest))));
    };

    // Negative controls reach the same real stage/overwrite route. A successful
    // copy must not consume an absent, nonmatching, or already consumed arm;
    // a complete short/empty write has no partial-write cut to inject.
    const auto retireBeforeEnvironmentRestore = [&]() noexcept
    {
        bool quiet = CancelOutstandingSelfTestOperations(operations);
        if (quiet)
        {
            const auto retired = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                return ! operations->HasActiveOperations() &&
                       std::ranges::all_of(ownedTaskIds, [&](uint64_t taskId) noexcept { return ! operations->FindTask(taskId); });
            }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"D4 exact worker retirement before environment restore"});
            quiet              = retired.conditionMet;
        }
        if (! quiet)
        {
            // Failure-only terminal fallback: initial admission required an idle
            // state, so every worker here belongs to this fixture. Shutdown joins
            // task/reaper workers before scoped hook values or files can unwind.
            // This state cannot admit later work; the aggregate remains failed.
            AppendLog(L"D4 bounded retirement failed; hard shutdown joins fixture workers before environment restoration.");
            operations->Shutdown();
        }
    };
    const auto successfulOverwrite = [&](const std::filesystem::path& source, const std::filesystem::path& destination) noexcept
    {
        const auto admitted = StartFileOperationAndGetId(operations,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         sourceFs,
                                                         {source},
                                                         destination.parent_path(),
                                                         FILESYSTEM_FLAG_NONE,
                                                         false);
        if (! admitted.has_value())
            return false;
        const uint64_t taskId = admitted.value();
        ownedTaskIds.push_back(taskId);
        if (auto* task = operations->FindTask(taskId))
            task->_presentationState.store(Task::TaskPresentationState::Presented, std::memory_order_release);
        unsigned int decisions      = 0u;
        uint64_t acceptedGeneration = 0u;
        bool invalidPrompt          = false;
        const auto quiet            = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            auto* task        = operations->FindTask(taskId);
            const auto prompt = TryGetConflictPromptCopy(task);
            if (task && prompt.has_value() && prompt.value().generation != acceptedGeneration)
            {
                if (decisions != 0u || prompt.value().bucket != Task::ConflictBucket::RegularFileExists ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite) ||
                    ! task->SubmitConflictDecision(prompt.value().generation, Task::ConflictAction::Overwrite, false))
                {
                    invalidPrompt = true;
                    return true;
                }
                ++decisions;
                acceptedGeneration = prompt.value().generation;
            }
            return ! operations->HasActiveOperations() && ! operations->FindTask(taskId) && state.completedTasks.contains(taskId);
        },
            {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"D4 negative-control actual overwrite completion"});
        if (! quiet.conditionMet || invalidPrompt || decisions != 1u || operations->HasActiveOperations() || operations->FindTask(taskId))
            return false;
        const auto completed = state.completedTasks.find(taskId);
        if (completed == state.completedTasks.end() || completed->second.hr != S_OK || completed->second.sourceItemResults.size() != 1u ||
            ! completed->second.sourceItemResults.front().has_value())
            return false;
        const auto& item = completed->second.sourceItemResults.front().value();
        if (item.status != S_OK || item.completion != FileOperations::ItemCompletion::Completed ||
            item.publication != FileOperations::PublicationState::Published || item.sourceDisposition != FileOperations::SourceDisposition::Retained ||
            item.finalSourcePath != source.native() || item.finalDestinationPath != destination.native())
            return false;
        AppendLog(std::format(L"D4 successful negative-control task={} actual publication={} source={} final={}",
                              taskId,
                              static_cast<unsigned int>(item.publication),
                              static_cast<unsigned int>(item.sourceDisposition),
                              destination.native()));
        return true;
    };

    for (unsigned int scenario = 0u; scenario < 2u; ++scenario)
    {
        const bool locked                  = scenario == 0u;
        const auto caseRoot                = root / (locked ? L"locked" : L"stage-write-failure");
        const auto source                  = caseRoot / L"src" / L"payload.bin";
        const auto destination             = caseRoot / L"dst" / L"payload.bin";
        const auto sourceReference         = caseRoot / L"reference" / L"source.bin";
        const auto destinationReference    = caseRoot / L"reference" / L"original.bin";
        constexpr size_t kSourceBytes      = 2u * 1024u * 1024u + 4099u;
        constexpr size_t kDestinationBytes = 2053u;
        if (! WriteFilledTestFile(source, kSourceBytes, 0xA7u) || ! WriteFilledTestFile(destination, kDestinationBytes, 0x3Cu) ||
            ! WriteFilledTestFile(sourceReference, kSourceBytes, 0xA7u) || ! WriteFilledTestFile(destinationReference, kDestinationBytes, 0x3Cu))
            return reject(L"distinct source/occupant/reference bytes unavailable");
        Proof sourceBefore;
        Proof destinationBefore;
        if (! capture(sourceIo, source, sourceBefore) || ! capture(destinationIo, destination, destinationBefore) ||
            sourceBefore.bytes.size() != kSourceBytes || destinationBefore.bytes.size() != kDestinationBytes ||
            ! std::ranges::all_of(sourceBefore.bytes, [](char value) noexcept { return static_cast<unsigned char>(value) == 0xA7u; }) ||
            ! std::ranges::all_of(destinationBefore.bytes, [](char value) noexcept { return static_cast<unsigned char>(value) == 0x3Cu; }) ||
            ! FilesEqualBytes(source, sourceReference) || ! FilesEqualBytes(destination, destinationReference))
            return reject(L"pre-admission bytes/hash/FILE_ID_INFO witnesses unavailable");
        logProof(scenario, L"before", source, sourceBefore);
        logProof(scenario, L"before", destination, destinationBefore);

        // Closed inventory: observe only the four known directories, reject
        // every reparse point, and never recurse into discovered child names.
        using Entry          = std::pair<std::filesystem::path, DWORD>;
        const auto inventory = [&](std::wstring_view phase, std::vector<Entry>& entries) noexcept
        {
            entries.clear();
            std::error_code ec;
            for (const auto& directory : std::array{caseRoot, source.parent_path(), destination.parent_path(), sourceReference.parent_path()})
            {
                const DWORD attributes = GetFileAttributesW(directory.c_str());
                if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
                    return false;
                std::filesystem::directory_iterator entry(directory, std::filesystem::directory_options::none, ec);
                if (ec)
                    return false;
                for (; entry != std::filesystem::directory_iterator{}; entry.increment(ec))
                {
                    if (ec || entries.size() >= 7u)
                        return false;
                    const DWORD childAttributes = GetFileAttributesW(entry->path().c_str());
                    if (childAttributes == INVALID_FILE_ATTRIBUTES || (childAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
                        return false;
                    entries.emplace_back(entry->path().lexically_relative(caseRoot), childAttributes & FILE_ATTRIBUTE_DIRECTORY);
                }
                if (ec)
                    return false;
            }
            std::ranges::sort(entries);
            for (const auto& [relative, kind] : entries)
                AppendLog(std::format(L"D4 case={} inventory={} exact-name={} kind={}", scenario, phase, relative.native(), kind));
            return true;
        };
        std::vector<Entry> expected{{L"src", FILE_ATTRIBUTE_DIRECTORY},
                                    {L"dst", FILE_ATTRIBUTE_DIRECTORY},
                                    {L"reference", FILE_ATTRIBUTE_DIRECTORY},
                                    {L"src/payload.bin", 0u},
                                    {L"dst/payload.bin", 0u},
                                    {L"reference/source.bin", 0u},
                                    {L"reference/original.bin", 0u}};
        for (auto& [path, kind] : expected)
        {
            static_cast<void>(kind);
            path.make_preferred();
        }
        std::ranges::sort(expected);
        std::vector<Entry> beforeInventory;
        if (! inventory(L"before", beforeInventory) || beforeInventory != expected)
            return reject(L"pre-admission exact closed inventory differs");

        // Declare environment owners before drain-on-exit: no arm/receipt is
        // restored while an owned worker can still read or write it.
        Env sourceArm(kSource, locked ? std::nullopt : std::optional<std::wstring_view>(source.native()));
        Env fired(kFired, std::nullopt);
        Env stage(kStage, std::nullopt);
        Env bytes(kBytes, std::nullopt);
        Env digest(kDigest, std::nullopt);
        wil::unique_handle lock;
        auto quietBeforeEnvironmentRestore = wil::scope_exit([&]() noexcept { retireBeforeEnvironmentRestore(); });
        // The observation handles above are already closed. Keeping a bound
        // destination authority here would itself prevent a share-zero open.
        if (locked)
        {
            lock.reset(CreateFileW(destination.c_str(), GENERIC_READ, 0u, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT, nullptr));
            if (! lock)
                return reject(L"real case-owned exclusive destination lock unavailable");
            // Probe the same actual publication authority the normal issue callback
            // requests. Keep the lock held through consent and worker retirement.
            wil::com_ptr<IFileSystemObjectBinding> binding;
            wil::com_ptr<IFileSystemBoundObject> refusedDestination;
            constexpr FileSystemBindFlags conflictFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA |
                                                                                           FILESYSTEM_BIND_RENAME | FILESYSTEM_BIND_PUBLICATION);
            const bool bindingAvailable                 = destinationFs.try_query_to(binding.put());
            const HRESULT refusal =
                bindingAvailable && binding ? binding->BindObject(destination.c_str(), conflictFlags, refusedDestination.put()) : E_NOINTERFACE;
            AppendLog(std::format(L"D4 locked actual BindObject flags=0x{:08X} status=0x{:08X} authority={} destination={}",
                                  static_cast<unsigned int>(conflictFlags),
                                  static_cast<unsigned long>(refusal),
                                  static_cast<bool>(refusedDestination),
                                  destination.native()));
            if (refusal != HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) || refusedDestination)
                return reject(L"locked destination did not refuse actual publication authority with sharing violation");
        }
        // Do not pass destinationFs here: an explicit destination filesystem
        // selects the host bridge even for Local->Local. The native Local route
        // uses sourceFs for both real paths, like routine same-provider Copy.
        const auto taskIdResult = StartFileOperationAndGetId(operations,
                                                             FILESYSTEM_COPY,
                                                             FolderWindow::Pane::Left,
                                                             FolderWindow::Pane::Right,
                                                             sourceFs,
                                                             {source},
                                                             destination.parent_path(),
                                                             FILESYSTEM_FLAG_NONE,
                                                             false);
        if (! taskIdResult.has_value())
            return reject(L"normal production admission failed");
        const uint64_t taskId = taskIdResult.value();
        ownedTaskIds.push_back(taskId);
        if (auto* task = operations->FindTask(taskId))
            task->_presentationState.store(Task::TaskPresentationState::Presented, std::memory_order_release);
        uint64_t acceptedGeneration     = 0u;
        unsigned int overwriteDecisions = 0u;
        bool lockedRefusalPrompt        = false;
        bool writeFailurePrompt         = false;
        bool producerCallbackProved     = false;
        bool invalidPrompt              = false;
        const auto quiet                = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            auto* task        = operations->FindTask(taskId);
            const auto prompt = TryGetConflictPromptCopy(task);
            if (task && prompt.has_value() && prompt.value().generation != acceptedGeneration)
            {
                const auto& current         = prompt.value();
                Task::ConflictAction action = Task::ConflictAction::Skip;
                AppendLog(std::format(L"D4 observed prompt case={} task={} generation={} bucket={} status=0x{:08X} source={} destination={} action-count={}",
                                      scenario,
                                      taskId,
                                      current.generation,
                                      static_cast<unsigned int>(current.bucket),
                                      static_cast<unsigned long>(current.status),
                                      current.sourcePath,
                                      current.destinationPath,
                                      current.actionCount));
                for (size_t index = 0u; index < (std::min)(current.actionCount, current.actions.size()); ++index)
                    AppendLog(std::format(L"D4 observed offered-action index={} action={}", index, static_cast<unsigned int>(current.actions[index])));
                if (current.sourcePath != source.native() || current.destinationPath != destination.native())
                {
                    invalidPrompt = true;
                    return true;
                }
                if (locked && ! lockedRefusalPrompt && overwriteDecisions == 0u && current.bucket == Task::ConflictBucket::RegularFileExists &&
                    current.status == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && ! PromptHasAction(current, Task::ConflictAction::Overwrite) &&
                    PromptHasAction(current, Task::ConflictAction::Skip))
                {
                    // Actual replacement authority was refused while the lock is held.
                    // Follow the normal offered Skip path; never submit withheld Overwrite.
                    lockedRefusalPrompt = true;
                }
                else if (! locked && ! writeFailurePrompt && overwriteDecisions == 0u && current.bucket == Task::ConflictBucket::RegularFileExists &&
                         PromptHasAction(current, Task::ConflictAction::Overwrite))
                {
                    action = Task::ConflictAction::Overwrite;
                    ++overwriteDecisions;
                }
                else if (! locked && overwriteDecisions == 1u && ! writeFailurePrompt && current.status == HRESULT_FROM_WIN32(ERROR_WRITE_FAULT) &&
                         PromptHasAction(current, Task::ConflictAction::Skip))
                {
                    writeFailurePrompt = true;
                    // This is the host's provider-returned prompt, after Local's
                    // real FileSystemItemCompleted callback. Observe that real
                    // producer receipt before Skip seals the final item status.
                    std::scoped_lock lockResults(task->_sourceItemStatusMutex);
                    if (task->_sourceItemResultBuilders.size() == 1u)
                    {
                        const auto& builder = task->_sourceItemResultBuilders.front();
                        if (builder.status == HRESULT_FROM_WIN32(ERROR_WRITE_FAULT) && builder.mutation.has_value())
                        {
                            const auto& receipt    = builder.mutation.value();
                            producerCallbackProved = FileSystemRouteContract::IsValidItemMutationResultPrefix(receipt) &&
                                                     FileSystemItemMutationResultHasOwnedStageDisposition(receipt) && receipt.outcomeKnown == TRUE &&
                                                     receipt.mutationCommitted == FALSE && receipt.originalStillPresent == TRUE &&
                                                     receipt.ownedStageDisposition == FileSystemOwnedStageDisposition::Removed;
                        }
                    }
                    if (! producerCallbackProved)
                    {
                        invalidPrompt = true;
                        return true;
                    }
                }
                else
                {
                    invalidPrompt = true;
                    return true;
                }
                AppendLog(std::format(L"D4 case={} task={} normal prompt generation={} bucket={} status=0x{:08X} action={}",
                                      scenario,
                                      taskId,
                                      current.generation,
                                      static_cast<unsigned int>(current.bucket),
                                      static_cast<unsigned long>(current.status),
                                      static_cast<unsigned int>(action)));
                if (! task->SubmitConflictDecision(current.generation, action, false))
                    invalidPrompt = true;
                else
                    acceptedGeneration = current.generation;
            }
            return invalidPrompt || (! operations->HasActiveOperations() && ! operations->FindTask(taskId) && state.completedTasks.contains(taskId));
        },
            {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"D4 actual overwrite worker completion and quiet point"});
        if (! quiet.conditionMet || invalidPrompt || operations->HasActiveOperations() || operations->FindTask(taskId))
            return reject(L"normal consent or worker/completion quiet point failed");
        // Release only the fixture-owned exclusive handle before any after reads.
        lock.reset();
        const auto completed     = state.completedTasks.find(taskId);
        const HRESULT expectedHr = S_FALSE;
        if (completed != state.completedTasks.end())
        {
            AppendLog(std::format(L"D4 observed completed task={} case={} status=0x{:08X} typed-count={}",
                                  taskId,
                                  scenario,
                                  static_cast<unsigned long>(completed->second.hr),
                                  completed->second.sourceItemResults.size()));
            for (const auto& observed : completed->second.sourceItemResults)
            {
                if (observed.has_value())
                    AppendLog(std::format(L"D4 observed typed source-index={} status=0x{:08X} publication={} verification={} source={} stage={} completion={} "
                                          L"final-source={} final-destination={}",
                                          observed->sourceIndex,
                                          static_cast<unsigned long>(observed->status),
                                          static_cast<unsigned int>(observed->publication),
                                          static_cast<unsigned int>(observed->verification),
                                          static_cast<unsigned int>(observed->sourceDisposition),
                                          static_cast<unsigned int>(observed->ownedStageDisposition),
                                          static_cast<unsigned int>(observed->completion),
                                          observed->finalSourcePath,
                                          observed->finalDestinationPath));
            }
        }
        if (completed == state.completedTasks.end() || completed->second.hr != expectedHr || completed->second.sourceItemResults.size() != 1u ||
            ! completed->second.sourceItemResults.front().has_value() ||
            (locked ? ! lockedRefusalPrompt || overwriteDecisions != 0u || writeFailurePrompt
                    : overwriteDecisions != 1u || ! writeFailurePrompt || ! producerCallbackProved))
            return reject(L"terminal task lost locked refusal or explicit post-write Skip");
        const auto& item = completed->second.sourceItemResults.front().value();
        if (item.sourceIndex != 0u || item.status != expectedHr || item.sourceDisposition != FileOperations::SourceDisposition::Retained ||
            item.publication != FileOperations::PublicationState::NotPublished || item.completion != FileOperations::ItemCompletion::Skipped ||
            item.ownedStageDisposition != (locked ? FileOperations::OwnedStageDisposition::NotCreated : FileOperations::OwnedStageDisposition::Removed) ||
            item.verification != FileOperations::VerificationState::NotRequested || item.finalSourcePath != source.native() ||
            item.finalDestinationPath != destination.native())
            return reject(L"normal producer/callback/reducer result axes disagree with storage truth");
        std::vector<State::CompletedTaskSummary> summaries;
        operations->CollectCompletedTasks(summaries);
        const auto summary = std::ranges::find_if(summaries, [taskId](const auto& value) noexcept { return value.taskId == taskId; });
        if (summary == summaries.end() || summary->resultHr != expectedHr || summary->publishedItemCount != 0u || summary->unknownPublicationItemCount != 0u ||
            summary->removedSourceCount != 0u || summary->retainedSourceCount != 1u || summary->unknownSourceCount != 0u ||
            summary->retainedOwnedStageCount != 0u || summary->unknownOwnedStageCount != 0u || summary->indeterminateItemCount != 0u)
            return reject(L"normal completed summary lost zero-publication/source-retention/stage-cleanup axes");
        const auto readEnv = [](std::wstring_view name) { return RedSalamander::TestSupport::GetEnvironmentString(name); };
        if (locked)
        {
            if (! readEnv(kFired).empty() || ! readEnv(kStage).empty() || ! readEnv(kBytes).empty() || ! readEnv(kDigest).empty())
                return reject(L"locked guard refusal unexpectedly reached the real stage-write hook");
        }
        else
        {
            // Accept the actual successful first write's count (bounded by its
            // 1MiB request). GetContentProof reads the real retained owned stage.
            const std::wstring observedBytesText  = readEnv(kBytes);
            const std::string observedBytesNarrow = Common::Strings::Utf8FromUtf16ReplacingInvalid(observedBytesText);
            uint64_t observedStageBytes           = 0u;
            const auto parsed = std::from_chars(observedBytesNarrow.data(), observedBytesNarrow.data() + observedBytesNarrow.size(), observedStageBytes);
            if (parsed.ec != std::errc{} || parsed.ptr != observedBytesNarrow.data() + observedBytesNarrow.size() || observedStageBytes == 0u ||
                observedStageBytes > 1024u * 1024u || observedStageBytes >= sourceBefore.bytes.size())
                return reject(L"actual partial stage byte count is not a bounded positive prefix");
            Common::Crypto::Blake3Hasher prefixHasher;
            prefixHasher.Update(std::as_bytes(std::span<const char>(sourceBefore.bytes.data(), static_cast<size_t>(observedStageBytes))));
            const auto prefixDigest = prefixHasher.Finalize();
            const std::filesystem::path stagePath(readEnv(kStage));
            const std::wstring observedDigest = readEnv(kDigest);
            const DWORD stageAttributes       = GetFileAttributesW(stagePath.c_str());
            const DWORD stageError            = stageAttributes == INVALID_FILE_ATTRIBUTES ? GetLastError() : ERROR_SUCCESS;
            if (readEnv(kFired) != L"1" || ! readEnv(kSource).empty() || observedBytesText != std::format(L"{}", observedStageBytes) ||
                observedDigest != Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(prefixDigest)) || stagePath.empty() ||
                stagePath.parent_path() != destination.parent_path() || stagePath == destination || stagePath == source ||
                stageAttributes != INVALID_FILE_ATTRIBUTES || (stageError != ERROR_FILE_NOT_FOUND && stageError != ERROR_PATH_NOT_FOUND))
                return reject(L"one-shot actual partial stage path/bytes/BLAKE3 or actual cleanup witness failed");
            AppendLog(std::format(L"D4 task={} real partial stage={} observedBytes={} BLAKE3={} failure=ERROR_WRITE_FAULT; stage removed by normal cleanup",
                                  taskId,
                                  stagePath.native(),
                                  observedStageBytes,
                                  observedDigest));
        }
        Proof sourceAfter;
        Proof destinationAfter;
        std::vector<Entry> afterInventory;
        const bool capturedSource      = capture(sourceIo, source, sourceAfter);
        const bool capturedDestination = capture(destinationIo, destination, destinationAfter);
        const bool capturedInventory   = inventory(L"after-observed", afterInventory);
        AppendLog(std::format(L"D4 observed after capture source={} destination={} inventory={}", capturedSource, capturedDestination, capturedInventory));
        if (capturedSource)
            logProof(scenario, L"after-observed", source, sourceAfter);
        if (capturedDestination)
            logProof(scenario, L"after-observed", destination, destinationAfter);
        if (! capturedSource || ! capturedDestination || ! sameProof(sourceBefore, sourceAfter) || ! sameProof(destinationBefore, destinationAfter) ||
            ! FilesEqualBytes(source, sourceReference) || ! FilesEqualBytes(destination, destinationReference) || ! capturedInventory ||
            afterInventory != beforeInventory || GetForegroundWindow() != foreground)
            return reject(L"source/original exact bytes/hash/file IDs, closed no-follow inventory or foreground changed");
        logProof(scenario, L"after", source, sourceAfter);
        logProof(scenario, L"after", destination, destinationAfter);
        AppendLog(std::format(L"D4 case={} task={} hr=0x{:08X} publication={} verification={} source={} stage={} completion={}; normal worker quiet",
                              scenario,
                              taskId,
                              static_cast<unsigned long>(item.status),
                              static_cast<unsigned int>(item.publication),
                              static_cast<unsigned int>(item.verification),
                              static_cast<unsigned int>(item.sourceDisposition),
                              static_cast<unsigned int>(item.ownedStageDisposition),
                              static_cast<unsigned int>(item.completion)));
        if (! locked)
        {
            const auto receiptBeforeReplay = std::array{readEnv(kFired), readEnv(kStage), readEnv(kBytes), readEnv(kDigest)};
            if (! successfulOverwrite(source, destination))
                return reject(L"same-source replay with the consumed one-shot arm did not publish normally");
            Proof replaySource;
            Proof replayDestination;
            std::vector<Entry> replayInventory;
            if (! capture(sourceIo, source, replaySource) || ! capture(destinationIo, destination, replayDestination) ||
                ! sameProof(sourceBefore, replaySource) || replayDestination.bytes != sourceBefore.bytes || replayDestination.digest != sourceBefore.digest ||
                ! inventory(L"after-successful-replay", replayInventory) || replayInventory != beforeInventory || ! readEnv(kSource).empty() ||
                receiptBeforeReplay != std::array{readEnv(kFired), readEnv(kStage), readEnv(kBytes), readEnv(kDigest)})
                return reject(L"same-source replay changed the old failure receipt, source bytes or closed inventory");
            logProof(scenario, L"successful-replay", destination, replayDestination);
        }
        quietBeforeEnvironmentRestore.release();
    }

    for (unsigned int control = 0u; control < 4u; ++control)
    {
        constexpr std::array<std::wstring_view, 4> kNames{{L"unarmed", L"nonmatching-arm", L"short-complete-write", L"empty-file"}};
        const auto caseRoot      = root / kNames[control];
        const auto source        = caseRoot / L"src" / L"payload.bin";
        const auto destination   = caseRoot / L"dst" / L"payload.bin";
        const auto reference     = caseRoot / L"ref" / L"original-source.bin";
        const size_t sourceBytes = control == 3u ? 0u : (control == 2u ? 4099u : 2u * 1024u * 1024u + 4099u);
        if (! WriteFilledTestFile(source, sourceBytes, 0xA7u) || ! WriteFilledTestFile(reference, sourceBytes, 0xA7u) ||
            ! WriteFilledTestFile(destination, 2053u, 0x3Cu))
            return reject(L"negative-control distinct fixture bytes unavailable");
        Proof sourceBefore;
        Proof destinationBefore;
        if (! capture(sourceIo, source, sourceBefore) || ! capture(destinationIo, destination, destinationBefore) || sourceBefore.bytes.size() != sourceBytes ||
            destinationBefore.bytes.size() != 2053u ||
            ! std::ranges::all_of(sourceBefore.bytes, [](char value) noexcept { return static_cast<unsigned char>(value) == 0xA7u; }) ||
            ! std::ranges::all_of(destinationBefore.bytes, [](char value) noexcept { return static_cast<unsigned char>(value) == 0x3Cu; }) ||
            sourceBefore.bytes == destinationBefore.bytes || ! FilesEqualBytes(source, reference))
            return reject(L"negative-control original byte/hash witnesses unavailable");
        const std::wstring arm = control == 0u ? std::wstring{} : (control == 1u ? (caseRoot / L"different-source.bin").native() : source.native());
        Env sourceArm(kSource, arm.empty() ? std::nullopt : std::optional<std::wstring_view>(arm));
        Env fired(kFired, std::nullopt);
        Env stage(kStage, std::nullopt);
        Env bytes(kBytes, std::nullopt);
        Env digest(kDigest, std::nullopt);
        auto quietBeforeEnvironmentRestore = wil::scope_exit([&]() noexcept { retireBeforeEnvironmentRestore(); });
        if (! successfulOverwrite(source, destination))
            return reject(L"negative-control complete overwrite failed");
        Proof sourceAfter;
        Proof destinationAfter;
        if (! capture(sourceIo, source, sourceAfter) || ! capture(destinationIo, destination, destinationAfter) || ! sameProof(sourceBefore, sourceAfter) ||
            destinationAfter.bytes != sourceBefore.bytes || destinationAfter.digest != sourceBefore.digest || ! FilesEqualBytes(source, reference) ||
            GetForegroundWindow() != foreground)
            return reject(L"negative-control successful publication or source preservation differs");
        const auto readEnv = [](std::wstring_view name) { return RedSalamander::TestSupport::GetEnvironmentString(name); };
        if (readEnv(kSource) != arm || ! readEnv(kFired).empty() || ! readEnv(kStage).empty() || ! readEnv(kBytes).empty() || ! readEnv(kDigest).empty())
            return reject(L"absent/nonmatching/complete-write arm unexpectedly consumed or published a failure receipt");
        // Only enumerate the three known child directories and their root.
        // Unexpected temporary siblings or reparse entries fail immediately.
        std::vector<std::filesystem::path> observed;
        std::error_code ec;
        for (const auto& directory : std::array{caseRoot, source.parent_path(), destination.parent_path(), reference.parent_path()})
        {
            const DWORD attributes = GetFileAttributesW(directory.c_str());
            if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
                return reject(L"negative-control owned directory was replaced");
            std::filesystem::directory_iterator entry(directory, std::filesystem::directory_options::none, ec);
            if (ec)
                return reject(L"negative-control inventory cannot enumerate");
            for (; entry != std::filesystem::directory_iterator{}; entry.increment(ec))
            {
                if (ec || observed.size() >= 6u)
                    return reject(L"negative-control inventory exceeded its closed bound");
                const DWORD childAttributes = GetFileAttributesW(entry->path().c_str());
                if (childAttributes == INVALID_FILE_ATTRIBUTES || (childAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
                    return reject(L"negative-control inventory contains a replaced/reparse entry");
                observed.push_back(entry->path().lexically_relative(caseRoot));
            }
            if (ec)
                return reject(L"negative-control inventory stopped before its end");
        }
        std::vector<std::filesystem::path> expected{L"src", L"dst", L"ref", L"src/payload.bin", L"dst/payload.bin", L"ref/original-source.bin"};
        for (auto& path : expected)
            path.make_preferred();
        std::ranges::sort(expected);
        std::ranges::sort(observed);
        if (observed != expected)
            return reject(L"negative-control exact closed inventory contains a stage or missing path");
        logProof(control + 2u, kNames[control], destination, destinationAfter);
        quietBeforeEnvironmentRestore.release();
    }
    if (! cleanup())
        return reject(L"owned worker/fixture cleanup failed");
    drain.release();
    if (! restoreSettings())
        return reject(L"restored settings save did not become quiet");
    restore.release();
    return GetForegroundWindow() == foreground;
}

[[nodiscard]] bool RunFileOpsOverwritePreservationChecks(SelfTestState& state) noexcept
{
    // ScopedEnvironmentVariable can report setup errors with system_error. This
    // named boundary logs once and fails; allocation failure remains fatal.
    try
    {
        return RunFileOpsOverwritePreservationChecksImpl(state);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::system_error& error)
    {
        Debug::Error(L"D4 overwrite fixture environment setup failed: {}", Utf16FromUtf8ForFileOpsSelfTest(error.what()));
        return false;
    }
}
