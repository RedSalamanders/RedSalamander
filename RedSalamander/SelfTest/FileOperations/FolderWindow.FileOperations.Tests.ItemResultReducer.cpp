// Included by FolderWindow.FileOperations.Tests.Runner.cpp; the reducer matrix is deterministic and
// filesystem-free. Task callback/finalizer and provider witnesses are separated below.

#include <array>
#include <atomic>
#include <initializer_list>
#include <limits>
#include <thread>

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
        const auto summary = fileOperations.DebugRecordCompletedTaskForSelfTest(task);
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
            task._itemCompletedCallbackCount.load(std::memory_order_acquire) == 0u &&
            (summary.lastDiagnosticMessage == LoadStringResource(nullptr, IDS_FILEOPS_DIAG_DESTINATION_EXISTS_UNVERIFIABLE)) == expectCollision &&
            (! expectCollision || (runningNote == summary.lastDiagnosticMessage && summary.lastProblem.has_value() &&
                                   summary.lastProblem.value().userText.messageId == IDS_FILEOPS_DIAG_DESTINATION_EXISTS_UNVERIFIABLE));
        AppendLog(std::format(L"Unverifiable destination collision: case={} exists={} sameBytes={} sourceUnsupported={} hr=0x{:08X} preserved={} "
                              L"passed={} note='{}'",
                              variant,
                              destinationExists,
                              sourceBytes == destinationBytes,
                              sourceUnsupported,
                              static_cast<unsigned long>(finalHr),
                              destinationPreserved && actualSource == sourceBytes,
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
        const ULONGLONG deadline = GetTickCount64() + 10'000ull;
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
