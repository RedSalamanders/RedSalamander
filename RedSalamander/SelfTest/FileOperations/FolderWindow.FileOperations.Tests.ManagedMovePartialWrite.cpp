// Included by the canonical selftest TU. Two explicitly selected regular leaves
// form closed flat trees. This case does not claim recursive selected-root or
// atomic-tree Move coverage. File IDs witness preservation, never mutation authority.
[[nodiscard]] bool RunFileOpsManagedMovePartialWriteChecksImpl(SelfTestState& state, std::wstring& unavailableReason)
{
    using State                              = FolderWindow::FileOperationState;
    using Task                               = State::Task;
    using Env                                = RedSalamander::TestSupport::ScopedEnvironmentVariable;
    constexpr std::wstring_view kSource      = L"REDSALAMANDER_FILEOPS_MANAGED_STAGE_WRITE_FAIL_SOURCE";
    constexpr std::wstring_view kFired       = L"REDSALAMANDER_FILEOPS_MANAGED_STAGE_WRITE_FAIL_FIRED";
    constexpr std::wstring_view kDestination = L"REDSALAMANDER_FILEOPS_MANAGED_STAGE_WRITE_FAIL_DESTINATION";
    constexpr std::wstring_view kStage       = L"REDSALAMANDER_FILEOPS_MANAGED_STAGE_WRITE_FAIL_STAGE";
    constexpr std::wstring_view kBytes       = L"REDSALAMANDER_FILEOPS_MANAGED_STAGE_WRITE_FAIL_BYTES";
    constexpr std::wstring_view kDigest      = L"REDSALAMANDER_FILEOPS_MANAGED_STAGE_WRITE_FAIL_BLAKE3";
    unavailableReason.clear();
    const auto reject = [&state](std::wstring_view reason) noexcept
    {
        state.retainTempRootOnFailure = true;
        AppendLog(std::format(L"D2 real managed Move: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    wil::com_ptr<IFileSystemIO> io;
    if (! folder || ! operations || ! state.fsLocal || ! state.infoLocal || ! state.fsLocal.try_query_to(io.put()) ||
        ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"actual Local endpoint or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground    = GetForegroundWindow();
    const HWND originalPopup = operations->GetPopupHwndForSelfTest();
    const auto root          = state.tempRoot / L"managed-move-partial-stage-write";
    std::wstring alternateDetail;
    const auto alternate = TryCreateAlternateWritableVolumeSelfTestRoot(state.tempRoot, alternateDetail);
    if (! alternate.has_value())
    {
        // This bounded qualification fixture cannot silently turn a missing
        // independent authorized volume into a passing D2 result.
        AppendLog(std::format(L"D2 unavailable: {}; no qualification claimed", alternateDetail));
        unavailableReason = alternateDetail.empty() ? L"alternate-volume-unavailable" : alternateDetail;
        return false;
    }
    std::optional<uint64_t> ownedTaskId;
    const auto cleanup = [&]() noexcept
    {
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
        if (! CancelOutstandingSelfTestOperations(operations))
        {
            state.retainTempRootOnFailure = true;
            AppendLog(L"D2 cleanup: workers not quiet; owned fixture retained.");
            return false;
        }
        if (state.retainTempRootOnFailure)
        {
            AppendLog(std::format(L"D2 failed witness retained after quiet: primary={} alternate={}; no storage/diagnostic dismissal",
                                  root.native(),
                                  alternate.value().native()));
            return true;
        }
        if (ownedTaskId.has_value())
        {
            operations->DismissCompletedTask(ownedTaskId.value());
            operations->DebugRemoveDiagnosticsForTask(ownedTaskId.value());
            state.completedTasks.erase(ownedTaskId.value());
        }
        if (! originalPopup && operations->GetPopupHwndForSelfTest())
            operations->DebugResetPopupForSelfTest();
        const bool primaryRemoved   = SelfTest::RemoveAll(root);
        const bool alternateRemoved = SelfTest::RemoveAll(alternate.value());
        if (alternateRemoved)
            PruneEmptyAlternateVolumeSandboxParents(alternate.value());
        return primaryRemoved && alternateRemoved;
    };
    auto drain                   = wil::scope_exit([&]() noexcept { static_cast<void>(cleanup()); });
    const auto sourceVolume      = TryGetVolumeGuidPathForPath(state.tempRoot);
    const auto destinationVolume = TryGetVolumeGuidPathForPath(alternate.value());
    if (! sourceVolume.has_value() || ! destinationVolume.has_value() || sourceVolume.value() == destinationVolume.value())
        return reject(L"runtime GUID identity does not prove independent volumes");
    AppendLog(std::format(L"D2 actual source-volume={} destination-volume={} source-owned-root={} destination-owned-root={}",
                          sourceVolume.value(),
                          destinationVolume.value(),
                          root.native(),
                          alternate.value().native()));

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
        const bool configurationRestored = SetPluginConfiguration(state.infoLocal.get(), originalLocalConfiguration);
        state.localConfigDirty           = localConfigDirty;
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
        return configurationRestored && operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u);
    };
    auto restore = wil::scope_exit([&]() noexcept { static_cast<void>(restoreSettings()); });
    if (! SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1,"deleteMaxConcurrency":1})json"))
        return reject(L"explicit Local serial budget unavailable");
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = true;
    HostSetAutoAcceptPrompts(false);
    HostClearTestPromptResultOverride();
    operations->SetAutoDismissSuccess(false);
    operations->SetQueueNewTasks(false);
    const auto sourceRoot        = root / L"source";
    const auto destinationRoot   = alternate.value() / L"destination";
    const auto first             = sourceRoot / L"first.bin";
    const auto second            = sourceRoot / L"second.bin";
    const auto firstDestination  = destinationRoot / first.filename();
    const auto secondDestination = destinationRoot / second.filename();
    const auto sentinel          = destinationRoot / L"sentinel.bin";
    constexpr size_t kFirstBytes = 131'077u;
    // Exceeds the production bridge's 16MiB maximum buffer, including
    // adaptive provider hints. The first actual write cannot finish this file.
    constexpr size_t kSecondBytes   = 16u * 1024u * 1024u + 4099u;
    constexpr size_t kSentinelBytes = 2053u;
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(destinationRoot) || ! WriteFilledTestFile(first, kFirstBytes, 0x31u) ||
        ! WriteFilledTestFile(second, kSecondBytes, 0xA7u) || ! WriteFilledTestFile(sentinel, kSentinelBytes, 0x5Cu))
        return reject(L"owned flat trees and distinct exact byte patterns unavailable");
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
            static_cast<uint64_t>(standard.EndOfFile.QuadPart) > kSecondBytes ||
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
        AppendLog(std::format(L"D2 phase={} path={} bytes={} volume={} FILE_ID_INFO={} BLAKE3={}",
                              phase,
                              path.native(),
                              proof.bytes.size(),
                              proof.id.VolumeSerialNumber,
                              Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(std::as_bytes(std::span(proof.id.FileId.Identifier)))),
                              Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(proof.digest))));
    };
    // Only these two fixed directories are enumerated. Every entry must be a
    // regular file; no discovered directory or reparse entry is followed.
    const auto inventory = [&](std::wstring_view phase, const std::filesystem::path& directory, std::vector<std::filesystem::path>& names) noexcept
    {
        names.clear();
        const DWORD attributes = GetFileAttributesW(directory.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
            return false;
        std::error_code ec;
        std::filesystem::directory_iterator entry(directory, std::filesystem::directory_options::none, ec);
        for (; ! ec && entry != std::filesystem::directory_iterator{}; entry.increment(ec))
        {
            const DWORD childAttributes = GetFileAttributesW(entry->path().c_str());
            if (childAttributes == INVALID_FILE_ATTRIBUTES || (childAttributes & (FILE_ATTRIBUTE_REPARSE_POINT | FILE_ATTRIBUTE_DIRECTORY)) != 0u ||
                names.size() >= 3u)
                return false;
            names.push_back(entry->path().filename());
        }
        if (ec)
            return false;
        std::ranges::sort(names);
        for (const auto& name : names)
            AppendLog(std::format(L"D2 inventory={} directory={} exact-file={}", phase, directory.native(), name.native()));
        return true;
    };
    const auto missing = [](const std::filesystem::path& path) noexcept
    {
        const DWORD attributes = GetFileAttributesW(path.c_str());
        const DWORD error      = attributes == INVALID_FILE_ATTRIBUTES ? GetLastError() : ERROR_SUCCESS;
        return attributes == INVALID_FILE_ATTRIBUTES && (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND);
    };
    const auto singleChildDirectory = [](std::wstring_view phase, const std::filesystem::path& parent, const std::filesystem::path& expectedChild) noexcept
    {
        const DWORD parentAttributes = GetFileAttributesW(parent.c_str());
        if (parentAttributes == INVALID_FILE_ATTRIBUTES || (parentAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u ||
            (parentAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
            return false;
        std::error_code ec;
        size_t count = 0u;
        std::filesystem::directory_iterator entry(parent, std::filesystem::directory_options::none, ec);
        for (; ! ec && entry != std::filesystem::directory_iterator{}; entry.increment(ec))
        {
            const DWORD attributes = GetFileAttributesW(entry->path().c_str());
            if (entry->path() != expectedChild || attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u ||
                (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u || ++count != 1u)
                return false;
            AppendLog(std::format(L"D2 inventory={} directory={} exact-directory={}", phase, parent.native(), entry->path().filename().native()));
        }
        return ! ec && count == 1u;
    };
    Proof firstBefore;
    Proof secondBefore;
    Proof sentinelBefore;
    std::vector<std::filesystem::path> sourceNames;
    std::vector<std::filesystem::path> destinationNames;
    const std::vector<std::filesystem::path> expectedSourceBefore{L"first.bin", L"second.bin"};
    const std::vector<std::filesystem::path> expectedDestinationBefore{L"sentinel.bin"};
    const std::vector<std::filesystem::path> expectedSourceAfter{L"second.bin"};
    const std::vector<std::filesystem::path> expectedDestinationAfter{L"first.bin", L"sentinel.bin"};
    if (! capture(first, firstBefore) || ! capture(second, secondBefore) || ! capture(sentinel, sentinelBefore) ||
        firstBefore.bytes != std::string(kFirstBytes, '\x31') || secondBefore.bytes != std::string(kSecondBytes, '\xA7') ||
        sentinelBefore.bytes != std::string(kSentinelBytes, '\x5C') || firstBefore.id.VolumeSerialNumber == sentinelBefore.id.VolumeSerialNumber ||
        ! inventory(L"before", sourceRoot, sourceNames) || sourceNames != expectedSourceBefore || ! inventory(L"before", destinationRoot, destinationNames) ||
        destinationNames != expectedDestinationBefore || ! singleChildDirectory(L"before", root, sourceRoot) ||
        ! singleChildDirectory(L"before", alternate.value(), destinationRoot))
        return reject(L"exact pre-admission bytes, hashes, IDs and closed inventories differ");
    logProof(L"before", first, firstBefore);
    logProof(L"before", second, secondBefore);
    logProof(L"before", sentinel, sentinelBefore);

    Env sourceArm(kSource, std::optional<std::wstring_view>(second.native()));
    Env fired(kFired, std::nullopt);
    Env observedDestination(kDestination, std::nullopt);
    Env observedStage(kStage, std::nullopt);
    Env observedBytes(kBytes, std::nullopt);
    Env observedDigest(kDigest, std::nullopt);
    std::atomic<unsigned int> callbackCount{0u};
    std::atomic<HRESULT> callbackResult{E_PENDING};
    const uint64_t subscription = folder->AddFileOperationCompletedCallback([&](const FolderWindow::FileOperationCompletedEvent& event) noexcept
    {
        if (ownedTaskId.has_value() && event.taskId == ownedTaskId.value())
        {
            callbackResult.store(event.hr, std::memory_order_release);
            callbackCount.fetch_add(1u, std::memory_order_acq_rel);
        }
    });
    const auto unsubscribe      = wil::scope_exit([&]() noexcept { folder->RemoveFileOperationCompletedCallback(subscription); });
    // This guard executes before environment/configuration restoration on every
    // return. Cancellation is emergency cleanup only, never the failure stimulus.
    auto quietBeforeRestore = wil::scope_exit([&]() noexcept
    {
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
        bool quiet = CancelOutstandingSelfTestOperations(operations);
        if (quiet && ownedTaskId.has_value())
        {
            const auto retired = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                return ! operations->HasActiveOperations() && ! operations->FindTask(ownedTaskId.value());
            }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"D2 exact retirement before environment restoration"});
            quiet              = retired.conditionMet;
        }
        if (! quiet)
        {
            // Failure-only terminal fallback. The fixture started with an idle
            // operation state; join its workers before scoped hook values and
            // provider settings can be restored. The run remains failed.
            AppendLog(L"D2 bounded retirement failed; hard shutdown joins fixture workers before environment restoration.");
            operations->Shutdown();
        }
    });
    SetFileOpsLiveOutputPublishedPauseForSelfTest(true);
    const auto taskIdResult = StartFileOperationAndGetId(operations,
                                                         FILESYSTEM_MOVE,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         {first, second},
                                                         destinationRoot,
                                                         FILESYSTEM_FLAG_CONTINUE_ON_ERROR,
                                                         false,
                                                         0u,
                                                         State::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsLocal);
    if (! taskIdResult.has_value())
        return reject(L"normal real Local cross-volume Move admission failed");
    ownedTaskId = taskIdResult.value();
    if (auto* task = operations->FindTask(ownedTaskId.value()))
        task->_presentationState.store(Task::TaskPresentationState::Presented, std::memory_order_release);
    const auto gate          = RedSalamander::TestSupport::PumpMessagesUntil([]() noexcept {
        return HasFileOpsLiveOutputPublishedPauseEnteredForSelfTest();
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"D2 first actual managed Move publication"});
    const ULONGLONG gateTick = GetTickCount64();
    auto* task               = operations->FindTask(ownedTaskId.value());
    if (! gate.conditionMet || ! task || task->_effectiveConcurrencyBudget.load(std::memory_order_acquire) != 1u || task->_perItemMaxConcurrencyBudget != 1u ||
        task->_perItemMaxConcurrency != 1u || task->_sourcePaths != std::vector<std::filesystem::path>{first, second} ||
        TryGetConflictPromptCopy(task).has_value())
        return reject(L"first publication gate or explicit production serial budget unavailable");
    const auto plans     = task->LoadPlans();
    size_t selectedCount = 0u;
    if (! plans)
        return reject(L"actual immutable transfer plans unavailable");
    for (const auto& plan : *plans)
    {
        const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan);
        if (! transfer || transfer->intent != FileOperations::TransferIntent::Move || transfer->strategy != FileOperations::OperationStrategy::Managed ||
            ! transfer->options.verifyAfterCopy || transfer->sourceEndpoint.pluginId != kPluginIdLocal ||
            transfer->destinationEndpoint.pluginId != kPluginIdLocal || transfer->sourceEndpoint.rootId.empty() ||
            transfer->destinationEndpoint.rootId.empty() || transfer->sourceEndpoint.rootId == transfer->destinationEndpoint.rootId)
            return reject(L"actual plans are not verified real Local managed Move across distinct roots");
        selectedCount += transfer->selectedItems.size();
        AppendLog(std::format(L"D2 prepared source-root={} destination-root={} strategy=Managed verify=On selected={}",
                              transfer->sourceEndpoint.rootId,
                              transfer->destinationEndpoint.rootId,
                              transfer->selectedItems.size()));
    }
    if (selectedCount != 2u || ! inventory(L"first-published-gate", sourceRoot, sourceNames) || sourceNames != expectedSourceBefore ||
        ! inventory(L"first-published-gate", destinationRoot, destinationNames) || destinationNames != expectedDestinationAfter ||
        ! singleChildDirectory(L"first-published-gate", root, sourceRoot) ||
        ! singleChildDirectory(L"first-published-gate", alternate.value(), destinationRoot))
        return reject(L"first publication did not leave the exact two-source/one-published flat trees");
    // Derive the actual published leaf from the retained gate inventory.
    const auto actualFirstPublished = destinationRoot / destinationNames.front();
    Proof firstAtGate;
    Proof secondAtGate;
    Proof destinationAtGate;
    Proof sentinelAtGate;
    if (actualFirstPublished != firstDestination || ! capture(first, firstAtGate) || ! sameProof(firstBefore, firstAtGate) || ! capture(second, secondAtGate) ||
        ! sameProof(secondBefore, secondAtGate) || ! capture(actualFirstPublished, destinationAtGate) || ! sameBytes(firstBefore, destinationAtGate) ||
        ! capture(sentinel, sentinelAtGate) || ! sameProof(sentinelBefore, sentinelAtGate) || ! missing(secondDestination) ||
        GetForegroundWindow() != foreground || GetTickCount64() - gateTick >= 5'000ull || ! capture(first, firstAtGate) ||
        ! sameProof(firstBefore, firstAtGate))
        return reject(L"actual first publish byte/ID proof or still-present source witnesses failed at bounded gate");
    logProof(L"first-published-gate", first, firstAtGate);
    logProof(L"first-published-gate", second, secondAtGate);
    logProof(L"first-published-gate", actualFirstPublished, destinationAtGate);
    logProof(L"first-published-gate", sentinel, sentinelAtGate);
    AppendLog(std::format(
        L"D2 task={} actual-first-published-path={} serial-budget=1 source-removal=not-yet-at-gate", ownedTaskId.value(), actualFirstPublished.native()));
    ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
    SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
    bool unexpectedPrompt     = false;
    const auto completedQuiet = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        auto* live       = operations->FindTask(ownedTaskId.value());
        unexpectedPrompt = live && TryGetConflictPromptCopy(live).has_value();
        return unexpectedPrompt || (! operations->HasActiveOperations() && ! live && state.completedTasks.contains(ownedTaskId.value()));
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(30)), .operationName = L"D2 actual stage failure and normal completion quiet point"});
    if (! completedQuiet.conditionMet || unexpectedPrompt || operations->HasActiveOperations() || operations->FindTask(ownedTaskId.value()))
        return reject(L"normal continue-on-error or worker/completion quiet point failed");
    const auto completed          = state.completedTasks.find(ownedTaskId.value());
    const HRESULT expectedFailure = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    if (completed == state.completedTasks.end() || completed->second.hr != expectedFailure || completed->second.configuredMaxConcurrency != 1u ||
        callbackCount.load(std::memory_order_acquire) != 1u || callbackResult.load(std::memory_order_acquire) != expectedFailure ||
        completed->second.sourceItemResults.size() != 2u || ! completed->second.sourceItemResults[0].has_value() ||
        ! completed->second.sourceItemResults[1].has_value())
        return reject(L"normal task completion lost two separate typed results or failure");
    for (size_t index = 0u; index < 2u; ++index)
    {
        const auto& item     = completed->second.sourceItemResults[index].value();
        const bool published = index == 0u;
        AppendLog(std::format(L"D2 observed task={} index={} source-index={} strategy={} status=0x{:08X} publication={} verification={} source={} stage={} "
                              L"completion={} from={} to={}",
                              ownedTaskId.value(),
                              index,
                              item.sourceIndex,
                              static_cast<unsigned int>(item.strategy),
                              static_cast<unsigned long>(item.status),
                              static_cast<unsigned int>(item.publication),
                              static_cast<unsigned int>(item.verification),
                              static_cast<unsigned int>(item.sourceDisposition),
                              static_cast<unsigned int>(item.ownedStageDisposition),
                              static_cast<unsigned int>(item.completion),
                              item.finalSourcePath,
                              item.finalDestinationPath));
        if (item.sourceIndex != index || item.strategy != FileOperations::OperationStrategy::Managed || item.status != (published ? S_OK : expectedFailure) ||
            item.completion != (published ? FileOperations::ItemCompletion::Completed : FileOperations::ItemCompletion::Failed) ||
            item.publication != (published ? FileOperations::PublicationState::Published : FileOperations::PublicationState::NotPublished) ||
            item.verification != (published ? FileOperations::VerificationState::Verified : FileOperations::VerificationState::NotApplicable) ||
            item.sourceDisposition != (published ? FileOperations::SourceDisposition::Removed : FileOperations::SourceDisposition::Retained) ||
            item.ownedStageDisposition != (published ? FileOperations::OwnedStageDisposition::Published : FileOperations::OwnedStageDisposition::Removed) ||
            item.finalSourcePath != (published ? first.native() : second.native()) ||
            item.finalDestinationPath != (published ? firstDestination.native() : secondDestination.native()))
            return reject(L"publication/verification/source/stage/completion axes disagree with separate leaf storage truth");
        AppendLog(std::format(L"D2 task={} source-index={} status=0x{:08X} publication={} verification={} source={} stage={} completion={}",
                              ownedTaskId.value(),
                              index,
                              static_cast<unsigned long>(item.status),
                              static_cast<unsigned int>(item.publication),
                              static_cast<unsigned int>(item.verification),
                              static_cast<unsigned int>(item.sourceDisposition),
                              static_cast<unsigned int>(item.ownedStageDisposition),
                              static_cast<unsigned int>(item.completion)));
    }
    std::vector<State::CompletedTaskSummary> summaries;
    operations->CollectCompletedTasks(summaries);
    const auto summary = std::ranges::find_if(summaries, [&](const auto& value) noexcept { return value.taskId == ownedTaskId.value(); });
    if (summary == summaries.end() || summary->resultHr != expectedFailure || summary->publishedItemCount != 1u || summary->verifiedItemCount != 1u ||
        summary->removedSourceCount != 1u || summary->retainedSourceCount != 1u || summary->unknownPublicationItemCount != 0u ||
        summary->unknownSourceCount != 0u || summary->retainedOwnedStageCount != 0u || summary->unknownOwnedStageCount != 0u ||
        summary->indeterminateItemCount != 0u || summary->effectiveConcurrencyBudget != 1u ||
        ! std::ranges::any_of(summary->issueDiagnostics,
                              [&](const auto& issue) noexcept
    {
        return issue.status == HRESULT_FROM_WIN32(ERROR_WRITE_FAULT) && issue.category == L"bridge.selftest.managedStageWrite.fail" &&
               issue.sourcePath == second.native() && issue.destinationPath == secondDestination.native();
    }))
        return reject(L"normal completed summary or actual StageWrite error path lost truthful axes");
    const auto readEnv                = [](std::wstring_view name) noexcept { return GetEnvVarTrimmed(name); };
    const std::wstring stageBytesText = readEnv(kBytes);
    const std::string narrow          = Common::Strings::Utf8FromUtf16ReplacingInvalid(stageBytesText);
    uint64_t stageBytes               = 0u;
    const auto parsed                 = std::from_chars(narrow.data(), narrow.data() + narrow.size(), stageBytes);
    if (parsed.ec != std::errc{} || parsed.ptr != narrow.data() + narrow.size() || stageBytes == 0u || stageBytes > 16u * 1024u * 1024u ||
        stageBytes >= kSecondBytes)
        return reject(L"actual persisted stage prefix was not positive and incomplete");
    Common::Crypto::Blake3Hasher prefixHasher;
    prefixHasher.Update(std::as_bytes(std::span<const char>(secondBefore.bytes.data(), static_cast<size_t>(stageBytes))));
    const auto prefixDigest = prefixHasher.Finalize();
    const std::filesystem::path stagePath(readEnv(kStage));
    if (readEnv(kFired) != L"1" || ! readEnv(kSource).empty() || readEnv(kDestination) != secondDestination.native() ||
        readEnv(kDigest) != Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(prefixDigest)) || stagePath.empty() ||
        stagePath.parent_path() != destinationRoot || stagePath == secondDestination || stagePath == firstDestination || ! missing(stagePath))
        return reject(L"one-shot exact source/destination actual stage proof or normal stage removal failed");
    Proof secondAfter;
    Proof destinationAfter;
    Proof sentinelAfter;
    if (! missing(first) || ! missing(secondDestination) || ! capture(second, secondAfter) || ! sameProof(secondBefore, secondAfter) ||
        ! capture(firstDestination, destinationAfter) || ! sameProof(destinationAtGate, destinationAfter) || ! capture(sentinel, sentinelAfter) ||
        ! sameProof(sentinelBefore, sentinelAfter) || ! inventory(L"after-quiet", sourceRoot, sourceNames) || sourceNames != expectedSourceAfter ||
        ! inventory(L"after-quiet", destinationRoot, destinationNames) || destinationNames != expectedDestinationAfter ||
        ! singleChildDirectory(L"after-quiet", root, sourceRoot) || ! singleChildDirectory(L"after-quiet", alternate.value(), destinationRoot) ||
        GetForegroundWindow() != foreground)
        return reject(L"published occupant, unpublished source, sentinel, exact final trees or foreground changed");
    logProof(L"after-quiet", second, secondAfter);
    logProof(L"after-quiet", firstDestination, destinationAfter);
    logProof(L"after-quiet", sentinel, sentinelAfter);
    AppendLog(
        std::format(L"D2 task={} actual-stage={} prefix-bytes={} BLAKE3={} injected=ERROR_WRITE_FAULT normal-task=ERROR_PARTIAL_COPY stage=Removed quiet=1",
                    ownedTaskId.value(),
                    stagePath.native(),
                    stageBytes,
                    readEnv(kDigest)));
    if (! cleanup())
        return reject(L"owned tree cleanup did not complete after normal worker quiet");
    drain.release();
    quietBeforeRestore.release();
    if (! restoreSettings())
        return reject(L"restored configuration/settings save failed");
    restore.release();
    return GetForegroundWindow() == foreground;
}

[[nodiscard]] bool RunFileOpsManagedMovePartialWriteChecks(SelfTestState& state, std::wstring& unavailableReason) noexcept
{
    try
    {
        return RunFileOpsManagedMovePartialWriteChecksImpl(state, unavailableReason);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::system_error& error)
    {
        Debug::Error(L"D2 managed Move fixture setup failed: {}", Utf16FromUtf8ForFileOpsSelfTest(error.what()));
        return false;
    }
}
