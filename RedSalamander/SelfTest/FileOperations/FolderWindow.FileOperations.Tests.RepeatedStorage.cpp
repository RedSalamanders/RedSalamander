// Included by the canonical selftest TU. Each invocation admits two actual
// intentional non-clipboard commands with distinct IDs. Queue-after is real
// consent; repeated Copy/Move is not treated as an invalid duplicate command.
[[nodiscard]] bool RunFileOpsRepeatedNonClipboardStorageChecksImpl(SelfTestState& state, FileSystemOperation operation, std::wstring& unavailableReason)
{
    using State = FolderWindow::FileOperationState;
    using Task  = State::Task;
    unavailableReason.clear();
    const bool move   = operation == FILESYSTEM_MOVE;
    const auto reject = [&](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"D11 repeated non-clipboard {}: {}", move ? L"Move" : L"Copy", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    wil::com_ptr<IFileSystemIO> io;
    if ((operation != FILESYSTEM_COPY && ! move) || ! folder || ! operations || ! state.fsLocal || ! state.infoLocal ||
        ! state.fsLocal.try_query_to(io.put()) || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"actual Local endpoints or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground    = GetForegroundWindow();
    const HWND originalPopup = operations->GetPopupHwndForSelfTest();
    const auto root          = state.tempRoot / (move ? L"repeat-nonclipboard-move" : L"repeat-nonclipboard-copy");
    std::optional<std::filesystem::path> alternate;
    if (move)
    {
        std::wstring detail;
        alternate = TryCreateAlternateWritableVolumeSelfTestRoot(state.tempRoot, detail);
        if (! alternate.has_value())
        {
            unavailableReason = detail.empty() ? L"alternate-volume-unavailable" : detail;
            AppendLog(std::format(L"D11 Move unavailable: {}; no qualification claimed", unavailableReason));
            return false;
        }
    }
    std::array<std::optional<uint64_t>, 2> taskIds{};
    bool storageWitnessPassed = false;
    const auto cleanup        = [&]() noexcept
    {
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
        if (! CancelOutstandingSelfTestOperations(operations))
        {
            AppendLog(L"D11 cleanup: workers not quiet; owned trees retained.");
            return false;
        }
        if (! storageWitnessPassed)
        {
            // Preserve failed actual bytes/results instead of erasing them while
            // unwinding. Workers are still drained before callbacks/settings.
            state.retainTempRootOnFailure = true;
            AppendLog(std::format(L"D11 failed witness retained primary={} alternate={}; no storage/diagnostic dismissal",
                                  root.native(),
                                  alternate.has_value() ? alternate.value().native() : std::wstring{}));
            return true;
        }
        for (const auto& taskId : taskIds)
        {
            if (taskId.has_value())
            {
                operations->DismissCompletedTask(taskId.value());
                operations->DebugRemoveDiagnosticsForTask(taskId.value());
                state.completedTasks.erase(taskId.value());
            }
        }
        if (! originalPopup && operations->GetPopupHwndForSelfTest())
            operations->DebugResetPopupForSelfTest();
        const bool primaryRemoved   = SelfTest::RemoveAll(root);
        const bool alternateRemoved = ! alternate.has_value() || SelfTest::RemoveAll(alternate.value());
        if (alternate.has_value() && alternateRemoved)
            PruneEmptyAlternateVolumeSandboxParents(alternate.value());
        return primaryRemoved && alternateRemoved;
    };
    auto drain = wil::scope_exit([&]() noexcept { static_cast<void>(cleanup()); });
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
    const auto restoreSettings = [&]() noexcept
    {
        const bool restored    = SetPluginConfiguration(state.infoLocal.get(), originalLocalConfiguration);
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
        return restored && operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u);
    };
    auto restore = wil::scope_exit([&]() noexcept { static_cast<void>(restoreSettings()); });
    if (! SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1,"deleteMaxConcurrency":1})json"))
        return reject(L"explicit serial configuration unavailable");
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = true;
    HostSetAutoAcceptPrompts(false);
    HostClearTestPromptResultOverride();
    operations->SetAutoDismissSuccess(false);
    operations->SetQueueNewTasks(false);
    const auto sourceRoot           = root / L"source";
    const auto destinationRoot      = (move ? alternate.value() : root) / L"destination";
    const auto source               = sourceRoot / L"payload.bin";
    const auto destination          = destinationRoot / source.filename();
    const auto sentinel             = destinationRoot / L"sentinel.bin";
    constexpr size_t kPayloadBytes  = 256u * 1024u + 17u;
    constexpr size_t kSentinelBytes = 2053u;
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(destinationRoot) || ! WriteFilledTestFile(source, kPayloadBytes, 0xA7u) ||
        ! WriteFilledTestFile(sentinel, kSentinelBytes, 0x3Cu))
        return reject(L"owned exact storage fixture unavailable");
    const auto sourceVolume      = TryGetVolumeGuidPathForPath(sourceRoot);
    const auto destinationVolume = TryGetVolumeGuidPathForPath(destinationRoot);
    if (! sourceVolume.has_value() || ! destinationVolume.has_value() || (move && sourceVolume.value() == destinationVolume.value()))
        return reject(L"actual volume identity unavailable or Move is not cross-volume");
    AppendLog(std::format(L"D11 {} actual source-volume={} destination-volume={} source={} destination={}",
                          move ? L"Move" : L"Copy",
                          sourceVolume.value(),
                          destinationVolume.value(),
                          source.native(),
                          destination.native()));
    struct Proof final
    {
        FILE_ID_INFO id{};
        std::string bytes;
        Common::Crypto::Blake3Digest digest{};
    };
    const auto capture = [&](const std::filesystem::path& path, Proof& proof) noexcept
    {
        wil::unique_handle handle(CreateFileW(path.c_str(),
                                              FILE_READ_ATTRIBUTES,
                                              FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                              nullptr,
                                              OPEN_EXISTING,
                                              FILE_FLAG_OPEN_REPARSE_POINT,
                                              nullptr));
        FILE_ATTRIBUTE_TAG_INFO tag{};
        FILE_STANDARD_INFO standard{};
        if (! handle || ! GetFileInformationByHandleEx(handle.get(), FileAttributeTagInfo, &tag, sizeof(tag)) ||
            (tag.FileAttributes & (FILE_ATTRIBUTE_REPARSE_POINT | FILE_ATTRIBUTE_DIRECTORY)) != 0u ||
            ! GetFileInformationByHandleEx(handle.get(), FileStandardInfo, &standard, sizeof(standard)) || standard.EndOfFile.QuadPart < 0 ||
            static_cast<uint64_t>(standard.EndOfFile.QuadPart) > kPayloadBytes ||
            ! GetFileInformationByHandleEx(handle.get(), FileIdInfo, &proof.id, sizeof(proof.id)) || ! ReadFileTextFsIo(io, path, proof.bytes))
            return false;
        Common::Crypto::Blake3Hasher hasher;
        hasher.Update(std::as_bytes(std::span<const char>(proof.bytes.data(), proof.bytes.size())));
        proof.digest = hasher.Finalize();
        return true;
    };
    const auto sameBytes = [](const Proof& left, const Proof& right) noexcept { return left.bytes == right.bytes && left.digest == right.digest; };
    const auto sameProof = [&](const Proof& left, const Proof& right) noexcept
    {
        return sameBytes(left, right) && left.id.VolumeSerialNumber == right.id.VolumeSerialNumber &&
               std::memcmp(left.id.FileId.Identifier, right.id.FileId.Identifier, sizeof(left.id.FileId.Identifier)) == 0;
    };
    const auto logProof = [](std::wstring_view phase, const std::filesystem::path& path, const Proof& proof) noexcept
    {
        AppendLog(std::format(L"D11 phase={} path={} bytes={} volume={} FILE_ID_INFO={} BLAKE3={}",
                              phase,
                              path.native(),
                              proof.bytes.size(),
                              proof.id.VolumeSerialNumber,
                              Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(std::as_bytes(std::span(proof.id.FileId.Identifier)))),
                              Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(proof.digest))));
    };
    // A fixed, closed inventory over known directories only. Check every child
    // with no recursion and reject all reparse entries, including directories.
    using Entry          = std::pair<std::filesystem::path, DWORD>;
    const auto inventory = [&](std::wstring_view phase, const std::filesystem::path& directory, std::vector<Entry>& entries) noexcept
    {
        entries.clear();
        const DWORD attributes = GetFileAttributesW(directory.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
            return false;
        std::error_code ec;
        std::filesystem::directory_iterator entry(directory, std::filesystem::directory_options::none, ec);
        for (; ! ec && entry != std::filesystem::directory_iterator{}; entry.increment(ec))
        {
            const DWORD childAttributes = GetFileAttributesW(entry->path().c_str());
            if (childAttributes == INVALID_FILE_ATTRIBUTES || (childAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u || entries.size() >= 2u)
                return false;
            entries.emplace_back(entry->path().filename(), childAttributes & FILE_ATTRIBUTE_DIRECTORY);
        }
        if (ec)
            return false;
        std::ranges::sort(entries);
        for (const auto& [name, kind] : entries)
            AppendLog(std::format(L"D11 inventory={} directory={} exact-name={} kind={}", phase, directory.native(), name.native(), kind));
        return true;
    };
    const std::vector<Entry> rootExpected = move ? std::vector<Entry>{{L"source", FILE_ATTRIBUTE_DIRECTORY}}
                                                 : std::vector<Entry>{{L"destination", FILE_ATTRIBUTE_DIRECTORY}, {L"source", FILE_ATTRIBUTE_DIRECTORY}};
    const std::vector<Entry> alternateExpected{{L"destination", FILE_ATTRIBUTE_DIRECTORY}};
    const std::vector<Entry> sourceExpected{{L"payload.bin", 0u}};
    const std::vector<Entry> emptyExpected;
    const std::vector<Entry> destinationBeforeExpected{{L"sentinel.bin", 0u}};
    const std::vector<Entry> destinationPublishedExpected{{L"payload.bin", 0u}, {L"sentinel.bin", 0u}};
    const auto closedTrees = [&](std::wstring_view phase, bool published, bool sourceRemoved) noexcept
    {
        std::vector<Entry> entries;
        return inventory(phase, root, entries) && entries == rootExpected &&
               (! move || (inventory(phase, alternate.value(), entries) && entries == alternateExpected)) && inventory(phase, sourceRoot, entries) &&
               entries == (sourceRemoved ? emptyExpected : sourceExpected) && inventory(phase, destinationRoot, entries) &&
               entries == (published ? destinationPublishedExpected : destinationBeforeExpected);
    };
    const auto missing = [](const std::filesystem::path& path) noexcept
    {
        const DWORD attributes = GetFileAttributesW(path.c_str());
        const DWORD error      = attributes == INVALID_FILE_ATTRIBUTES ? GetLastError() : ERROR_SUCCESS;
        return attributes == INVALID_FILE_ATTRIBUTES && (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND);
    };
    Proof sourceBefore;
    Proof sentinelBefore;
    if (! capture(source, sourceBefore) || sourceBefore.bytes != std::string(kPayloadBytes, '\xA7') || ! capture(sentinel, sentinelBefore) ||
        sentinelBefore.bytes != std::string(kSentinelBytes, '\x3C') || ! missing(destination) || ! closedTrees(L"before", false, false))
        return reject(L"exact original bytes/hash/IDs and pre-admission inventories differ");
    logProof(L"before", source, sourceBefore);
    logProof(L"before", sentinel, sentinelBefore);
    std::array<std::atomic<unsigned int>, 2> callbackCounts{};
    std::array<std::atomic<HRESULT>, 2> callbackResults{};
    std::array<std::atomic<ULONGLONG>, 2> callbackTicks{};
    for (auto& result : callbackResults)
        result.store(E_PENDING, std::memory_order_relaxed);
    const uint64_t subscription = folder->AddFileOperationCompletedCallback([&](const FolderWindow::FileOperationCompletedEvent& event) noexcept
    {
        for (size_t index = 0u; index < taskIds.size(); ++index)
        {
            if (taskIds[index].has_value() && event.taskId == taskIds[index].value())
            {
                callbackResults[index].store(event.hr, std::memory_order_release);
                callbackTicks[index].store(GetTickCount64(), std::memory_order_release);
                callbackCounts[index].fetch_add(1u, std::memory_order_acq_rel);
            }
        }
    });
    const auto unsubscribe      = wil::scope_exit([&]() noexcept { folder->RemoveFileOperationCompletedCallback(subscription); });
    auto quietBeforeRestore     = wil::scope_exit([&]() noexcept
    {
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
        if (! CancelOutstandingSelfTestOperations(operations))
        {
            AppendLog(L"D11 worker quiet failed before callback/configuration restoration.");
            // This isolated fixture starts from an idle owner. Join its workers
            // before callback storage, configuration or owned trees can unwind.
            operations->Shutdown();
        }
        const auto retired = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            return ! operations->HasActiveOperations() &&
                   std::ranges::all_of(taskIds, [&](const auto& id) noexcept { return ! id.has_value() || ! operations->FindTask(id.value()); });
        },
            {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"D11 failure cleanup task retirement"});
        if (! retired.conditionMet)
            operations->Shutdown();
    });
    SetFileOpsLiveOutputPublishedPauseForSelfTest(true);
    taskIds[0] = StartFileOperationAndGetId(operations,
                                            operation,
                                            FolderWindow::Pane::Left,
                                            FolderWindow::Pane::Right,
                                            state.fsLocal,
                                            {source},
                                            destinationRoot,
                                            FILESYSTEM_FLAG_NONE,
                                            false,
                                            0u,
                                            State::ExecutionMode::PerItem,
                                            false,
                                            state.fsLocal);
    if (! taskIds[0].has_value())
        return reject(L"first actual non-clipboard admission failed");
    if (auto* first = operations->FindTask(taskIds[0].value()))
        first->_presentationState.store(Task::TaskPresentationState::Presented, std::memory_order_release);
    const auto publicationGate = RedSalamander::TestSupport::PumpMessagesUntil([]() noexcept {
        return HasFileOpsLiveOutputPublishedPauseEnteredForSelfTest();
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"D11 first actual host bridge publication"});
    const ULONGLONG gateTick   = GetTickCount64();
    auto* first                = operations->FindTask(taskIds[0].value());
    if (! publicationGate.conditionMet || ! first || ! first->HasEnteredOperation() || first->_clipboardMoveAdmission ||
        first->_effectiveConcurrencyBudget.load(std::memory_order_acquire) != 1u)
        return reject(L"first actual publication gate or non-clipboard ownership unavailable");
    const auto firstPlans = first->LoadPlans();
    if (! firstPlans || firstPlans->size() != 1u)
        return reject(L"first actual immutable plan unavailable");
    const auto* transfer = std::get_if<FileOperations::TransferPlan>(&firstPlans->front());
    if (! transfer || transfer->intent != (move ? FileOperations::TransferIntent::Move : FileOperations::TransferIntent::Copy) ||
        transfer->strategy != (move ? FileOperations::OperationStrategy::Managed : FileOperations::OperationStrategy::Copy) ||
        ! transfer->options.verifyAfterCopy || transfer->moveClipboardSequence.has_value() || transfer->sourceEndpoint.pluginId != kPluginIdLocal ||
        transfer->destinationEndpoint.pluginId != kPluginIdLocal || (move && transfer->sourceEndpoint.rootId == transfer->destinationEndpoint.rootId))
        return reject(L"first plan did not select the real verified Local host bridge without clipboard admission");
    Proof sourceAtGate;
    Proof destinationAtGate;
    Proof sentinelAtGate;
    if (! capture(source, sourceAtGate) || ! sameProof(sourceBefore, sourceAtGate) || ! capture(destination, destinationAtGate) ||
        ! sameBytes(sourceBefore, destinationAtGate) || ! capture(sentinel, sentinelAtGate) || ! sameProof(sentinelBefore, sentinelAtGate) ||
        ! closedTrees(L"first-publication-gate", true, false))
        return reject(L"first publication did not preserve exact source, destination and sentinel storage witnesses");
    logProof(L"first-publication-gate", source, sourceAtGate);
    logProof(L"first-publication-gate", destination, destinationAtGate);
    logProof(L"first-publication-gate", sentinel, sentinelAtGate);
    taskIds[1] = StartFileOperationAndGetId(operations,
                                            operation,
                                            FolderWindow::Pane::Left,
                                            FolderWindow::Pane::Right,
                                            state.fsLocal,
                                            {source},
                                            destinationRoot,
                                            FILESYSTEM_FLAG_NONE,
                                            true,
                                            0u,
                                            State::ExecutionMode::PerItem,
                                            false,
                                            state.fsLocal);
    if (! taskIds[1].has_value() || taskIds[1] == taskIds[0])
        return reject(L"second intentional admission did not create a distinct real task ID");
    if (auto* second = operations->FindTask(taskIds[1].value()))
        second->_presentationState.store(Task::TaskPresentationState::Presented, std::memory_order_release);
    unsigned int queueDecisions = 0u;
    uint64_t queueGeneration    = 0u;
    bool invalidPrompt          = false;
    const auto secondQueued     = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        auto* second      = operations->FindTask(taskIds[1].value());
        const auto prompt = TryGetConflictPromptCopy(second);
        if (second && prompt.has_value() && prompt->generation != queueGeneration)
        {
            if (queueDecisions != 0u || prompt->bucket != Task::ConflictBucket::SameHostOverlap || prompt->sourcePath != source.native() ||
                prompt->destinationPath != destinationRoot.native() || ! PromptHasAction(prompt.value(), Task::ConflictAction::Proceed) ||
                ! second->SubmitConflictDecision(prompt->generation, Task::ConflictAction::Proceed, false))
                invalidPrompt = true;
            else
            {
                queueGeneration = prompt->generation;
                ++queueDecisions;
                AppendLog(std::format(L"D11 second-task={} real SameHostOverlap Queue-after accepted generation={}", taskIds[1].value(), queueGeneration));
            }
        }
        return invalidPrompt || (second && second->IsOverlapQueued() && second->IsWaitingInQueue() && ! second->HasEnteredOperation());
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(5)), .operationName = L"D11 second actual task Queue-after disposition"});
    first        = operations->FindTask(taskIds[0].value());
    auto* second = operations->FindTask(taskIds[1].value());
    Proof destinationWhileQueued;
    Proof sentinelWhileQueued;
    if (! secondQueued.conditionMet || invalidPrompt || queueDecisions != 1u || ! first || ! first->HasEnteredOperation() || ! second ||
        second->_clipboardMoveAdmission || ! second->IsOverlapQueued() || ! second->IsWaitingInQueue() || second->HasEnteredOperation() ||
        second->HasStarted() || second->_discoveredTotalBytes.load(std::memory_order_acquire) != 0u ||
        second->_hasPublishedLiveOutput.load(std::memory_order_acquire) || GetTickCount64() - gateTick >= 8'000ull ||
        callbackCounts[0].load(std::memory_order_acquire) != 0u || callbackCounts[1].load(std::memory_order_acquire) != 0u || ! capture(source, sourceAtGate) ||
        ! sameProof(sourceBefore, sourceAtGate) || ! capture(destination, destinationWhileQueued) || ! sameProof(destinationAtGate, destinationWhileQueued) ||
        ! capture(sentinel, sentinelWhileQueued) || ! sameProof(sentinelBefore, sentinelWhileQueued) || ! closedTrees(L"second-queued-gate", true, false) ||
        GetForegroundWindow() != foreground)
        return reject(L"second task entered, discovered, published, completed or changed storage while first real publication was held");
    logProof(L"second-queued-gate", source, sourceAtGate);
    logProof(L"second-queued-gate", destination, destinationWhileQueued);
    logProof(L"second-queued-gate", sentinel, sentinelWhileQueued);
    const auto secondPlans     = second->LoadPlans();
    const auto* secondTransfer = secondPlans && secondPlans->size() == 1u ? std::get_if<FileOperations::TransferPlan>(&secondPlans->front()) : nullptr;
    if (! secondTransfer || secondTransfer->moveClipboardSequence.has_value() || secondTransfer->strategy != transfer->strategy ||
        secondTransfer->options.executionMode != FileOperations::ExecutionMode::Queue)
        return reject(L"second command did not retain its independent non-clipboard queued plan");
    AppendLog(std::format(L"D11 first-task={} second-task={} distinct=1 first-entered=1 second-entered=0 second-queued=1 actual-published-path={}",
                          taskIds[0].value(),
                          taskIds[1].value(),
                          destination.native()));
    ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
    SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
    unsigned int existsSkipDecisions = 0u;
    uint64_t existsGeneration        = 0u;
    const auto completionQuiet       = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        auto* liveSecond  = operations->FindTask(taskIds[1].value());
        const auto prompt = TryGetConflictPromptCopy(liveSecond);
        if (liveSecond && prompt.has_value() && prompt->generation != existsGeneration)
        {
            if (move || existsSkipDecisions != 0u || prompt->bucket != Task::ConflictBucket::RegularFileExists || prompt->sourcePath != source.native() ||
                prompt->destinationPath != destination.native() || ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip) ||
                ! liveSecond->SubmitConflictDecision(prompt->generation, Task::ConflictAction::Skip, false))
                invalidPrompt = true;
            else
            {
                existsGeneration = prompt->generation;
                ++existsSkipDecisions;
                AppendLog(std::format(L"D11 second Copy task={} real RegularFileExists Skip accepted generation={}", taskIds[1].value(), existsGeneration));
            }
        }
        return invalidPrompt || (! operations->HasActiveOperations() && ! operations->FindTask(taskIds[0].value()) && ! liveSecond &&
                                 state.completedTasks.contains(taskIds[0].value()) && state.completedTasks.contains(taskIds[1].value()));
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(30)), .operationName = L"D11 actual second attempt and two completion quiet points"});
    if (! completionQuiet.conditionMet || invalidPrompt || operations->HasActiveOperations() || operations->FindTask(taskIds[0].value()) ||
        operations->FindTask(taskIds[1].value()) || existsSkipDecisions != (move ? 0u : 1u))
        return reject(L"normal queued execution, Exists/Skip or missing-source failure did not reach two quiet completions");
    const auto firstCompleted  = state.completedTasks.find(taskIds[0].value());
    const auto secondCompleted = state.completedTasks.find(taskIds[1].value());
    if (firstCompleted == state.completedTasks.end() || secondCompleted == state.completedTasks.end())
        return reject(L"two distinct normal completion payloads unavailable");
    const HRESULT secondHr = secondCompleted->second.hr;
    if (firstCompleted->second.hr != S_OK || (! move && secondHr != S_FALSE) ||
        (move && secondHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && secondHr != HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND)))
        return reject(L"first success and second real Skip/missing-source failure status differ");
    for (size_t index = 0u; index < taskIds.size(); ++index)
    {
        const auto& completed = index == 0u ? firstCompleted->second : secondCompleted->second;
        if (callbackCounts[index].load(std::memory_order_acquire) != 1u || callbackResults[index].load(std::memory_order_acquire) != completed.hr ||
            completed.sourceItemResults.size() != 1u || ! completed.sourceItemResults.front().has_value())
            return reject(L"each real task did not retain exactly one matching callback and typed item");
        const auto& item     = completed.sourceItemResults.front().value();
        const bool firstItem = index == 0u;
        AppendLog(std::format(L"D11 observed task={} invocation={} status=0x{:08X} strategy={} publication={} verification={} source={} stage={} completion={} "
                              L"sourcePath={} destinationPath={}",
                              taskIds[index].value(),
                              index,
                              static_cast<unsigned long>(item.status),
                              static_cast<unsigned int>(item.strategy),
                              static_cast<unsigned int>(item.publication),
                              static_cast<unsigned int>(item.verification),
                              static_cast<unsigned int>(item.sourceDisposition),
                              static_cast<unsigned int>(item.ownedStageDisposition),
                              static_cast<unsigned int>(item.completion),
                              item.finalSourcePath,
                              item.finalDestinationPath));
        if (item.sourceIndex != 0u || item.status != completed.hr || item.strategy != transfer->strategy ||
            item.publication != (firstItem ? FileOperations::PublicationState::Published
                                           : (move ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::NotPublished)) ||
            item.verification != (firstItem ? FileOperations::VerificationState::Verified
                                            : (move ? FileOperations::VerificationState::NotRequested : FileOperations::VerificationState::NotApplicable)) ||
            item.sourceDisposition != (move ? (firstItem ? FileOperations::SourceDisposition::Removed : FileOperations::SourceDisposition::Missing)
                                            : FileOperations::SourceDisposition::Retained) ||
            item.ownedStageDisposition !=
                (firstItem ? FileOperations::OwnedStageDisposition::Published
                           : (move ? FileOperations::OwnedStageDisposition::NotApplicable : FileOperations::OwnedStageDisposition::NotCreated)) ||
            item.completion != (firstItem ? FileOperations::ItemCompletion::Completed
                                          : (move ? FileOperations::ItemCompletion::Failed : FileOperations::ItemCompletion::Skipped)) ||
            item.finalSourcePath != source.native() || item.finalDestinationPath != destination.native())
            return reject(L"normal separate typed task results disagree with actual first publication and second attempt");
        AppendLog(std::format(
            L"D11 task={} invocation={} callback-count=1 callback-tick={} status=0x{:08X} publication={} verification={} source={} stage={} completion={}",
            taskIds[index].value(),
            index,
            callbackTicks[index].load(std::memory_order_acquire),
            static_cast<unsigned long>(item.status),
            static_cast<unsigned int>(item.publication),
            static_cast<unsigned int>(item.verification),
            static_cast<unsigned int>(item.sourceDisposition),
            static_cast<unsigned int>(item.ownedStageDisposition),
            static_cast<unsigned int>(item.completion)));
    }
    Proof sourceAfter;
    Proof destinationAfter;
    Proof sentinelAfter;
    if ((move ? ! missing(source) : ! capture(source, sourceAfter) || ! sameProof(sourceBefore, sourceAfter)) || ! capture(destination, destinationAfter) ||
        ! sameProof(destinationAtGate, destinationAfter) || ! capture(sentinel, sentinelAfter) || ! sameProof(sentinelBefore, sentinelAfter) ||
        ! closedTrees(L"after-two-quiet-completions", true, move) || GetForegroundWindow() != foreground)
        return reject(L"repeated actual command damaged first published occupant, source disposition, sentinel, exact trees or foreground");
    if (! move)
        logProof(L"after-two-quiet-completions", source, sourceAfter);
    logProof(L"after-two-quiet-completions", destination, destinationAfter);
    logProof(L"after-two-quiet-completions", sentinel, sentinelAfter);
    std::vector<State::CompletedTaskSummary> summaries;
    operations->CollectCompletedTasks(summaries);
    for (size_t index = 0u; index < taskIds.size(); ++index)
    {
        const auto summary = std::ranges::find_if(summaries, [&](const auto& value) noexcept { return value.taskId == taskIds[index].value(); });
        if (summary == summaries.end() || summary->clipboardMoveAdmission || summary->clipboardMoveConsumed || summary->unknownPublicationItemCount != 0u ||
            summary->unknownSourceCount != 0u || summary->unknownOwnedStageCount != 0u || summary->retainedOwnedStageCount != 0u ||
            summary->indeterminateItemCount != 0u || summary->publishedItemCount != (index == 0u ? 1u : 0u) ||
            summary->verifiedItemCount != (index == 0u ? 1u : 0u) || summary->removedSourceCount != (index == 0u && move ? 1u : 0u) ||
            summary->retainedSourceCount != (move ? 0u : 1u) || summary->missingSourceCount != (move && index == 1u ? 1u : 0u))
            return reject(L"normal summaries imply clipboard deduplication or lose independent task/storage axes");
    }
    storageWitnessPassed = true;
    if (! cleanup())
        return reject(L"owned task/tree cleanup failed after quiet");
    drain.release();
    quietBeforeRestore.release();
    if (! restoreSettings())
        return reject(L"restored Local configuration/settings save failed");
    restore.release();
    return GetForegroundWindow() == foreground;
}

[[nodiscard]] bool RunFileOpsRepeatedNonClipboardStorageChecks(SelfTestState& state, FileSystemOperation operation, std::wstring& unavailableReason) noexcept
{
    try
    {
        return RunFileOpsRepeatedNonClipboardStorageChecksImpl(state, operation, unavailableReason);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::system_error& error)
    {
        Debug::Error(L"D11 repeated storage fixture setup failed: {}", Utf16FromUtf8ForFileOpsSelfTest(error.what()));
        return false;
    }
}
