// I35 S2 (Specs/Plans/Done/FileOperations_CopyMoveUxAndReliabilityReview_2026-10-11.md, section 6):
// a directory root's discovery closes when its listing ends, never after its last byte, on every
// walker (Local parallel and serial, host bridge parallel and sequential). Post-order directory
// metadata still lands after the last child, and a Move never removes a child of a directory whose
// listing is still open.

namespace
{
constexpr auto kClosureWait               = std::chrono::seconds(90);
constexpr uint64_t kClosureFileBytes      = 1024u * 1024u;
constexpr unsigned long kClosureFlatFiles = 8u;
constexpr uint64_t kClosureSpeedLimit     = 4u * 1024u * 1024u; // 8 MiB takes about two seconds
constexpr std::string_view kLocalConcurrency4 =
    R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json";
constexpr std::string_view kLocalConcurrency1 =
    R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json";

// Owns the settings a closure case changes and every task it starts.
class ClosureCaseScope final
{
public:
    ClosureCaseScope(SelfTestState& state, FolderWindow::FileOperationState* operations) noexcept : _state(state), _operations(operations)
    {
        _verify        = EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy;
        _autoDismiss   = operations->GetAutoDismissSuccess();
        _originalPopup = operations->GetPopupHwndForSelfTest();
        operations->SetAutoDismissSuccess(false);
    }
    ClosureCaseScope(const ClosureCaseScope&)            = delete;
    ClosureCaseScope& operator=(const ClosureCaseScope&) = delete;
    ~ClosureCaseScope()
    {
        if (! _closed)
            static_cast<void>(Close());
    }
    [[nodiscard]] bool Close() noexcept
    {
        _closed          = true;
        const bool quiet = CancelOutstandingSelfTestOperations(_operations);
        for (const uint64_t taskId : _tasks)
        {
            if (quiet)
            {
                _operations->DismissCompletedTask(taskId);
                _operations->DebugRemoveDiagnosticsForTask(taskId);
                _state.completedTasks.erase(taskId);
            }
        }
        if (! quiet)
            return false;
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = _verify;
        _operations->SetAutoDismissSuccess(_autoDismiss);
        const bool saveSettled = _operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u);
        if (quiet && ! _originalPopup && _operations->GetPopupHwndForSelfTest())
            _operations->DebugResetPopupForSelfTest();
        return saveSettled;
    }

    // Runs one transfer under the closure speed limit and returns its completion facts. `onPump` sees
    // the running task between message pumps (sampling); `stopAfter` cancels the task once it returns
    // true, and the canceled task's facts are returned.
    [[nodiscard]] std::optional<CompletedTaskInfo> Run(
        FileSystemOperation operation,
        const wil::com_ptr<IFileSystem>& destinationFs,
        std::vector<std::filesystem::path> sources,
        const std::filesystem::path& destination,
        uint64_t speedLimit                                                          = kClosureSpeedLimit,
        const std::function<void(FolderWindow::FileOperationState::Task&)>& onPump   = {},
        const std::function<bool()>& stopAfter                                       = {},
        const wil::com_ptr<IFileSystem>& sourceFs                                    = {},
        const std::function<bool(FolderWindow::FileOperationState::Task&)>& onPrompt = {},
        FolderWindow::FileOperationState::ExecutionMode executionMode                = FolderWindow::FileOperationState::ExecutionMode::PerItem,
        FileSystemFlags extraFlags                                                   = FILESYSTEM_FLAG_NONE,
        size_t maxMessagesPerPump                                                    = 1024u,
        bool forceHostBridge                                                         = false)
    {
        const wil::com_ptr<IFileSystem>& source = sourceFs ? sourceFs : _state.fsLocal;
        const bool localSource                  = source.get() == _state.fsLocal.get();
        // The same provider instance on both sides is the provider's own route (no destination file
        // system); a different destination provider is the host bridge.
        const wil::com_ptr<IFileSystem> bridgeDestination =
            ! forceHostBridge && destinationFs.get() == source.get() ? wil::com_ptr<IFileSystem>{} : destinationFs;
        const auto admitted = StartFileOperationAndGetId(_operations,
                                                         operation,
                                                         localSource ? FolderWindow::Pane::Left : FolderWindow::Pane::Right,
                                                         localSource ? FolderWindow::Pane::Right : FolderWindow::Pane::Left,
                                                         source,
                                                         std::move(sources),
                                                         destination,
                                                         static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | extraFlags),
                                                         false,
                                                         speedLimit,
                                                         executionMode,
                                                         false,
                                                         bridgeDestination);
        if (! admitted.has_value())
            return std::nullopt;
        const uint64_t taskId = admitted.value();
        _tasks.push_back(taskId);
        bool stopRequested = false;
        const auto done    = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            FolderWindow::FileOperationState::Task* const running = _operations->FindTask(taskId);
            if (TryGetConflictPromptCopy(running).has_value() && ! (onPrompt && onPrompt(*running)))
                return true; // an unexpected question: report it rather than wait for the timeout
            if (running != nullptr && onPump)
                onPump(*running);
            if (running != nullptr && stopAfter && ! stopRequested && stopAfter())
            {
                stopRequested = true;
                running->RequestCancel();
            }
            return _state.completedTasks.contains(taskId) && running == nullptr;
        },
            {.timeout = SelfTest::Scale(kClosureWait), .maxMessagesPerPump = maxMessagesPerPump, .operationName = L"discovery closure transfer"});
        if (! done.conditionMet || ! _state.completedTasks.contains(taskId))
            return std::nullopt;
        return _state.completedTasks.at(taskId);
    }

    [[nodiscard]] uint64_t LastTaskId() const noexcept
    {
        return _tasks.empty() ? 0u : _tasks.back();
    }

private:
    SelfTestState& _state;
    FolderWindow::FileOperationState* _operations = nullptr;
    std::vector<uint64_t> _tasks;
    HWND _originalPopup = nullptr;
    bool _verify        = false;
    bool _autoDismiss   = false;
    bool _closed        = false;
};

// Seeds `count` files of kClosureFileBytes under `folder`.
[[nodiscard]] bool SeedClosureFiles(const std::filesystem::path& folder, unsigned long count, std::wstring_view prefix = L"file") noexcept
{
    for (unsigned long index = 0u; index < count; ++index)
    {
        if (! WriteFilledTestFile(
                folder / std::format(L"{}-{:02}.bin", prefix, index), static_cast<size_t>(kClosureFileBytes), static_cast<unsigned char>(0x41u + index)))
            return false;
    }
    return true;
}

// R6.1/R6.8: closed exactly on the expected totals, before the last byte, and never grew afterwards.
[[nodiscard]] std::wstring DescribeLateClosure(std::wstring_view name,
                                               const std::optional<CompletedTaskInfo>& info,
                                               uint64_t expectedBytes,
                                               unsigned long expectedFiles)
{
    if (! info.has_value())
        return std::format(L"{}: the transfer did not complete (or asked an unexpected question).", name);
    const auto describe = [&]()
    {
        return std::format(L"hr=0x{:08X} closed={} bytes={}/{} files={}/{} bytesWhileOpen={} mutationsWhileOpen={} growthAfterClose={}",
                           static_cast<unsigned long>(info->hr),
                           info->discoveryClosed,
                           info->discoveredTotalBytes,
                           expectedBytes,
                           info->discoveredFiles,
                           expectedFiles,
                           info->discoveryBytesCompletedWhileOpen,
                           info->discoveryMutationsCompletedWhileOpen,
                           info->discoveryGrowthAfterCloseCount);
    };
    AppendLog(std::format(L"Discovery closure: {} {}", name, describe()));
    if (FAILED(info->hr) || ! info->discoveryClosed || info->discoveredTotalBytes != expectedBytes || info->discoveredFiles != expectedFiles ||
        info->discoveryGrowthAfterCloseCount != 0u)
        return std::format(L"{}: wrong result or totals ({}).", name, describe());
    // The listing of a few entries takes milliseconds; the throttled copy takes seconds. A walker
    // that closes after its copies drained shows every byte completed while open.
    if (info->discoveryBytesCompletedWhileOpen * 2u >= info->discoveredTotalBytes)
        return std::format(L"{}: discovery closed only after the copy was half done or finished ({}).", name, describe());
    return {};
}
} // namespace

// Reducer edge controls complement the real provider copies below. These tasks are never admitted
// and own no workers: the reports exercise the production observer without fabricating I/O evidence.
[[nodiscard]] bool RunFolderFileWorkReducerControls(FolderWindow::FileOperationState& owner, const std::filesystem::path& root) noexcept
{
    using Task = FolderWindow::FileOperationState::Task;
    Task task(owner);
    task._executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
    task._fileWorkScopes.resize(2u);
    Task::PerItemCallbackCookie first(0u);
    Task::PerItemCallbackCookie second(1u);
    const std::wstring source = (root / L"reducer-first.bin").wstring();
    const std::wstring other  = (root / L"reducer-second.bin").wstring();
    FileSystemFileWorkProgress report{};
    report.sizeBytes = sizeof(report);
    report.phase     = FileSystemFileWorkPhase::ScopeSupported;
    bool passed      = SUCCEEDED(task.FileSystemReportFileWork(&report, &first)) && ! task._copiedFileCountKnown.load(std::memory_order_acquire);
    passed           = SUCCEEDED(task.FileSystemReportFileWork(&report, &second)) && task._copiedFileCountKnown.load(std::memory_order_acquire) && passed;
    const auto send  = [&](Task::PerItemCallbackCookie& cookie,
                           uint64_t id,
                           uint64_t stream,
                           const std::wstring& path,
                           FileSystemFileWorkPhase phase,
                           uint64_t size,
                           FileSystemFileWorkDisposition disposition = FileSystemFileWorkDisposition::None) noexcept
    {
        report.phase            = phase;
        report.kind             = stream == FILESYSTEM_PRODUCER_PROGRESS_STREAM_ID ? FileSystemFileWorkKind::FileLink : FileSystemFileWorkKind::RegularFile;
        report.logicalLeafId    = id;
        report.progressStreamId = stream;
        report.sourcePath       = path.c_str();
        report.sizeKnown        = TRUE;
        report.totalBytes       = size;
        report.completedBytes   = phase == FileSystemFileWorkPhase::Begin ? 0u : size;
        report.disposition      = disposition;
        report.status           = disposition == FileSystemFileWorkDisposition::Failed     ? E_FAIL
                                  : disposition == FileSystemFileWorkDisposition::Skipped  ? S_FALSE
                                  : disposition == FileSystemFileWorkDisposition::Canceled ? HRESULT_FROM_WIN32(ERROR_CANCELLED)
                                                                                           : S_OK;
        return SUCCEEDED(task.FileSystemReportFileWork(&report, &cookie));
    };
    // An empty file publishes once. A repeated Begin resets attempt bytes, never publication truth.
    passed = send(first, 1u, 0u, source, FileSystemFileWorkPhase::Begin, 0u) && passed;
    passed = send(first, 1u, 0u, source, FileSystemFileWorkPhase::Published, 0u) && passed;
    passed = send(first, 1u, 0u, source, FileSystemFileWorkPhase::Begin, 0u) && passed;
    passed = send(first, 1u, 0u, source, FileSystemFileWorkPhase::Published, 0u) && passed;
    passed = send(first, 1u, 0u, source, FileSystemFileWorkPhase::Terminal, 0u, FileSystemFileWorkDisposition::Failed) &&
             task._copiedFileCount.load(std::memory_order_acquire) == 1u && passed;
    passed = send(first, 2u, 0u, other, FileSystemFileWorkPhase::Begin, 100u) && passed;
    passed = send(first, 1u, 0u, source, FileSystemFileWorkPhase::Terminal, 0u, FileSystemFileWorkDisposition::Completed) &&
             task._inFlightFiles[0].logicalLeafId == 2u && ! task._inFlightFiles[0].terminal && passed;
    passed = send(first, 2u, 0u, other, FileSystemFileWorkPhase::Terminal, 100u, FileSystemFileWorkDisposition::Skipped) &&
             task._copiedFileCount.load(std::memory_order_acquire) == 1u && passed;
    passed = send(first, 3u, 0u, source, FileSystemFileWorkPhase::Begin, 100u) && passed;
    passed = send(first, 3u, 0u, source, FileSystemFileWorkPhase::Terminal, 100u, FileSystemFileWorkDisposition::Canceled) &&
             task._copiedFileCount.load(std::memory_order_acquire) == 1u && passed;
    // Sixteen ordinary workers coexist with two selected-root producer leaves, neither evicted.
    for (uint64_t index = 0u; index < 16u; ++index)
        passed = send(first, 10u + index, index, source, FileSystemFileWorkPhase::Begin, 100u) && passed;
    passed           = send(first, 50u, FILESYSTEM_PRODUCER_PROGRESS_STREAM_ID, source, FileSystemFileWorkPhase::Begin, 0u) && passed;
    passed           = send(second, 50u, FILESYSTEM_PRODUCER_PROGRESS_STREAM_ID, other, FileSystemFileWorkPhase::Begin, 0u) && passed;
    passed           = send(first, 50u, FILESYSTEM_PRODUCER_PROGRESS_STREAM_ID, source, FileSystemFileWorkPhase::Published, 0u) && passed;
    passed           = send(second, 50u, FILESYSTEM_PRODUCER_PROGRESS_STREAM_ID, other, FileSystemFileWorkPhase::Published, 0u) &&
                       task._copiedFileCount.load(std::memory_order_acquire) == 3u && task._inFlightFileCount == 16u &&
                       task._copiedFileCountKnown.load(std::memory_order_acquire) && passed;
    report           = {};
    report.sizeBytes = sizeof(report);
    report.phase     = FileSystemFileWorkPhase::CountsUnavailable;
    passed           = SUCCEEDED(task.FileSystemReportFileWork(&report, &first)) && ! task._copiedFileCountKnown.load(std::memory_order_acquire) && passed;
    report.phase     = FileSystemFileWorkPhase::ScopeSupported;
    passed           = SUCCEEDED(task.FileSystemReportFileWork(&report, &first)) && ! task._copiedFileCountKnown.load(std::memory_order_acquire) && passed;

    // Bulk owns one declared public-call scope regardless of the number of selected roots.
    Task bulk(owner);
    bulk._executionMode = FolderWindow::FileOperationState::ExecutionMode::BulkItems;
    bulk._fileWorkScopes.resize(1u);
    passed = SUCCEEDED(bulk.FileSystemReportFileWork(&report, &first)) && bulk._copiedFileCountKnown.load(std::memory_order_acquire) && passed;
    bulk._discoveryClosed.store(true, std::memory_order_release);
    {
        std::scoped_lock lock(bulk._inFlightFilesMutex);
        bulk.PublishFileWorkAvailabilityLocked();
    }
    passed = bulk._discoveredBytesKnown.load(std::memory_order_acquire) && bulk._exactDiscoveredBytes.load(std::memory_order_acquire) == 0u && passed;
    // Mixed omitted listing sizes require all known reader sizes and their independent exact sum.
    bulk._discoveredFileCount.store(2u, std::memory_order_release);
    bulk._discoveredTotalBytes.store(5u, std::memory_order_release);
    report.phase         = FileSystemFileWorkPhase::Begin;
    report.kind          = FileSystemFileWorkKind::RegularFile;
    report.logicalLeafId = 1u;
    report.sourcePath    = source.c_str();
    report.sizeKnown     = TRUE;
    report.totalBytes    = 100u;
    passed               = SUCCEEDED(bulk.FileSystemReportFileWork(&report, &first)) && ! bulk._discoveredBytesKnown.load(std::memory_order_acquire) && passed;
    report.phase         = FileSystemFileWorkPhase::Terminal;
    report.disposition   = FileSystemFileWorkDisposition::Skipped;
    passed               = SUCCEEDED(bulk.FileSystemReportFileWork(&report, &first)) && passed;
    report.phase         = FileSystemFileWorkPhase::Begin;
    report.disposition   = FileSystemFileWorkDisposition::None;
    report.logicalLeafId = 2u;
    report.sourcePath    = other.c_str();
    report.totalBytes    = 0u;
    passed               = SUCCEEDED(bulk.FileSystemReportFileWork(&report, &first)) && bulk._discoveredBytesKnown.load(std::memory_order_acquire) &&
                           bulk._exactDiscoveredBytes.load(std::memory_order_acquire) == 100u && passed;
    report.sizeKnown     = 2;
    passed               = bulk.FileSystemReportFileWork(&report, &first) == E_INVALIDARG && ! bulk._copiedFileCountKnown.load(std::memory_order_acquire) &&
                           ! bulk._discoveredBytesKnown.load(std::memory_order_acquire) && passed;
    for (unsigned int variant = 0u; variant < 3u; ++variant)
    {
        const bool changeKind    = variant == 1u;
        const bool certifiedSize = variant == 2u;
        Task malformed(owner);
        malformed._fileWorkScopes.resize(1u);
        malformed._discoveredTotalBytes.store(100u, std::memory_order_release);
        report                 = {};
        report.sizeBytes       = sizeof(report);
        report.phase           = FileSystemFileWorkPhase::ScopeSupported;
        report.scopeBytesKnown = certifiedSize ? TRUE : FALSE;
        passed = SUCCEEDED(malformed.FileSystemReportFileWork(&report, &first)) && malformed._copiedFileCountKnown.load(std::memory_order_acquire) && passed;
        report.phase         = FileSystemFileWorkPhase::Begin;
        report.kind          = FileSystemFileWorkKind::RegularFile;
        report.logicalLeafId = 1u;
        report.sourcePath    = source.c_str();
        report.sizeKnown     = TRUE;
        report.totalBytes    = 100u;
        passed               = SUCCEEDED(malformed.FileSystemReportFileWork(&report, &first)) && passed;
        report.kind          = changeKind ? FileSystemFileWorkKind::DirectoryLink : FileSystemFileWorkKind::RegularFile;
        report.phase         = changeKind ? FileSystemFileWorkPhase::Finalizing : static_cast<FileSystemFileWorkPhase>(UINT32_MAX);
        passed = malformed.FileSystemReportFileWork(&report, &first) == E_INVALIDARG && ! malformed._copiedFileCountKnown.load(std::memory_order_acquire) &&
                 malformed._discoveredBytesKnown.load(std::memory_order_acquire) == certifiedSize &&
                 (! certifiedSize || malformed._exactDiscoveredBytes.load(std::memory_order_acquire) == 100u) && passed;
        report = {};
        report.sizeBytes       = sizeof(report);
        report.phase           = FileSystemFileWorkPhase::ScopeSupported;
        report.scopeBytesKnown = certifiedSize ? TRUE : FALSE;
        passed = SUCCEEDED(malformed.FileSystemReportFileWork(&report, &first)) && ! malformed._copiedFileCountKnown.load(std::memory_order_acquire) && passed;
    }
    FileOperationsPopupInternal::TaskSnapshot projection{};
    projection.operation          = FILESYSTEM_COPY;
    projection.started            = true;
    projection.operationStartTick = 1u;
    projection.discoveryClosed    = true;
    projection.totalBytes         = 100u; // partial listing metadata, without certification
    projection.completedBytes     = 50u;
    projection.totalItems         = 1u; // selected roots cannot substitute for child work
    FileOperationsPopupInternal::TaskProgressPresentationDebugSnapshot projected{};
    DebugResolveFileOperationsWholeTaskProgress(projection, projected);
    passed                          = ! projected.determinate && ! DebugFileOperationsTaskHasKnownCompactProgress(projection) && passed;
    projection.totalBytes           = 0u;
    projection.completedBytes       = 0u;
    projection.discoveredBytesKnown = true;
    projection.copiedFileCountKnown = true;
    projection.discoveredFileCount  = 4u;
    projection.copiedFileCount      = 1u;
    DebugResolveFileOperationsWholeTaskProgress(projection, projected);
    passed = projected.determinate && std::abs(projected.fraction - 0.25) < 0.001 && DebugFileOperationsTaskHasKnownCompactProgress(projection) && passed;
    projection.verificationRequested = true;
    DebugResolveFileOperationsWholeTaskProgress(projection, projected);
    passed = ! projected.determinate && ! DebugFileOperationsTaskHasKnownCompactProgress(projection) && passed;
    AppendLog(std::format(L"FolderFileWorkReducerControls: "
                          L"allRoots,bulk,empty,skip,cancel,retry,lateTerminal,parallelProducers,replay,sizeSum,invalidBOOL,invalidEnum,immutableKind={}",
                          passed));
    return passed;
}

// I37: these are real provider/bridge copies, sampled at the normal 100 ms presentation cadence.
// Record every independent witness so the baseline's count failure does not hide its meter failure.
[[nodiscard]] bool RunDiscoveryScopeFolderLiveProgressChecks(SelfTestState& state) noexcept
{
    using namespace FileOperationsPopupInternal;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Folder live progress: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.fsDummy || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, providers or initial worker quiet point unavailable");
    if (! RunFolderFileWorkReducerControls(*operations, state.tempRoot))
        return reject(L"logical-leaf reducer controls failed");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const bool previousFooter   = operations->GetPopupFooterOnly();
    const bool previousDensity  = operations->GetPopupCompactDensity();
    const auto previousBufferKB = EnsureFileOperationsSettingsForSelfTest().crossFsBridgeBufferSizeKB;
    std::string previousConfiguration;
    if (! BackupPluginConfiguration(state.infoLocal.get(), previousConfiguration) || ! SetPluginConfiguration(state.infoLocal.get(), previousConfiguration) ||
        ! BackupPluginConfiguration(state.infoLocal.get(), previousConfiguration))
        return reject(L"Local configuration backup unavailable");
    bool settingsRestored      = false;
    const auto restoreSettings = [&]() noexcept
    {
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        if (! CancelOutstandingSelfTestOperations(operations))
        {
            AppendLog(L"Folder live progress cleanup: worker quiet point failed; preserve fixture and configuration.");
            return false;
        }
        operations->SetPopupFooterOnly(previousFooter);
        operations->SetPopupCompactDensity(previousDensity);
        EnsureFileOperationsSettingsForSelfTest().crossFsBridgeBufferSizeKB = previousBufferKB;
        const bool configurationRestored                                    = SetPluginConfiguration(state.infoLocal.get(), previousConfiguration);
        const bool saveSettled                                              = operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u);
        std::string restoredConfiguration;
        settingsRestored = configurationRestored && saveSettled && BackupPluginConfiguration(state.infoLocal.get(), restoredConfiguration) &&
                           restoredConfiguration == previousConfiguration && operations->GetPopupFooterOnly() == previousFooter &&
                           operations->GetPopupCompactDensity() == previousDensity &&
                           EnsureFileOperationsSettingsForSelfTest().crossFsBridgeBufferSizeKB == previousBufferKB;
        return settingsRestored;
    };
    const auto restore = wil::scope_exit([&]() noexcept
    {
        if (! settingsRestored)
            static_cast<void>(restoreSettings());
    });
    operations->SetPopupFooterOnly(false);
    operations->SetPopupCompactDensity(false);
    EnsureFileOperationsSettingsForSelfTest().crossFsBridgeBufferSizeKB = 8u * 1024u;
    const auto root                                                     = state.tempRoot / L"folder-live-progress";
    const auto source                                                   = root / L"tree";
    const auto output = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::FileOperations, L"folder-copy-progress");
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(output) || ! WriteFilledTestFile(source / L"first.bin", 32u * 1024u * 1024u, 0x31u) ||
        ! WriteFilledTestFile(source / L"nested" / L"second.bin", 32u * 1024u * 1024u, 0x52u) ||
        ! WriteFilledTestFile(source / L"nested" / L"deep" / L"third.bin", 32u * 1024u * 1024u, 0x73u) || ! WriteFilledTestFile(source / L"empty.bin", 0u, 0u))
        return reject(L"the independent nested fixture or artifact folder could not be created");
    const auto removeRoot                        = wil::scope_exit([&]() noexcept
    {
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        if (CancelOutstandingSelfTestOperations(operations))
            static_cast<void>(SelfTest::RemoveAll(root));
    });
    constexpr uint64_t kExpectedBytes            = 96u * 1024u * 1024u;
    constexpr unsigned long kExpectedFiles       = 4u;
    constexpr unsigned long kExpectedDirectories = 3u;
    constexpr uint64_t kSpeed                    = 16u * 1024u * 1024u;
    ClosureCaseScope scope(state, operations);
    std::vector<std::wstring> defects;
    const auto record = [&](bool passed, std::wstring_view route, std::wstring_view assertion)
    {
        AppendLog(std::format(L"FolderLiveProgressAssertion: route={} assertion={} passed={}", route, assertion, passed));
        if (! passed)
            defects.push_back(std::format(L"{}: {}", route, assertion));
    };
    const auto color = [](const D2D1_COLOR_F& value) noexcept
    {
        return RGB(static_cast<BYTE>(std::clamp(value.r, 0.0f, 1.0f) * 255.0f),
                   static_cast<BYTE>(std::clamp(value.g, 0.0f, 1.0f) * 255.0f),
                   static_cast<BYTE>(std::clamp(value.b, 0.0f, 1.0f) * 255.0f));
    };
    const auto run = [&](std::wstring_view route,
                         const wil::com_ptr<IFileSystem>& destinationFs,
                         const std::filesystem::path& destination,
                         bool verify,
                         bool overwrite,
                         bool slowFinalization)
    {
        // A fixed delay alone can end between the 650 ms observation and a slow screenshot.
        // Hold real nonempty commits until the capture finishes; the provider also honors cancel
        // and a hard 30 s bound. No renderer/lifecycle facts are substituted.
        const RedSalamander::TestSupport::ScopedEnvironmentVariable holdCommit(L"REDSALAMANDER_FILEOPS_HOLD_NONEMPTY_COMMIT",
                                                                               slowFinalization ? std::optional<std::wstring>(L"1") : std::nullopt);
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = verify;
        bool totalsSeen                                           = false;
        bool countRowSeen                                         = false;
        bool countOracleRejectsIncorrectLabels                    = false;
        bool countCaptureAttempted                                = false;
        unsigned countCaptureAttempts                             = 0u;
        bool countInkSeen                                         = false;
        bool totalSizeRowSeen                                     = false;
        bool earlyCaptured                                        = false;
        bool lateCaptured                                         = false;
        bool finalizationObserved                                 = false;
        bool finalizationMeter                                    = false;
        bool knownTotalRetained                                   = false;
        const bool naturalUpdates                                 = route == L"local-serial";
        bool popupEnsured                                         = false;
        uint64_t capturePaintGeneration                           = 0u;
        uint64_t earlyPaintGeneration                             = 0u;
        uint64_t latePaintGeneration                              = 0u;
        std::wstring sampledPath;
        std::wstring finalizationPath;
        std::array<std::wstring, 3> attemptedPaths{};
        size_t sampleAttempts          = 0u;
        uint64_t sampledTotal          = 0u;
        uint64_t finalizationTotal     = 0u;
        ULONGLONG sampledCompleteTick  = 0u;
        ULONGLONG lastSample           = 0u;
        uint32_t lastFinalizationPhase = UINT32_MAX;
        bool lastFinalizationFull      = false;
        bool lastFinalizationFound     = false;
        RECT pixelRegion{};
        PopupLayoutDebugSnapshot::PaintedFileProgress sampledPaint{};
        PopupLayoutDebugSnapshot::PaintedFileProgress finalizationPaint{};
        const auto onPump = [&](FolderWindow::FileOperationState::Task& running)
        {
            const ULONGLONG now = GetTickCount64();
            if (now - lastSample < 100u)
                return;
            lastSample = now;
            if (! naturalUpdates || ! popupEnsured)
            {
                operations->DebugEnsurePopupVisibleForSelfTest();
                popupEnsured = true;
            }
            const HWND popup = operations->GetPopupHwndForSelfTest();
            if (! popup)
                return;
            if (! naturalUpdates)
                RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
            TaskSnapshot task{};
            PopupLayoutDebugSnapshot painted{};
            painted.taskId = running._taskId;
            if (! DebugGetFileOperationsPopupTaskSnapshot(popup, running._taskId, task) || ! DebugGetFileOperationsPopupLayoutSnapshot(popup, painted) ||
                task.conflict.active || task.finished)
                return;
            if (slowFinalization && ! finalizationPath.empty())
            {
                bool found = false;
                for (size_t index = 0u; index < task.inFlightFileCount; ++index)
                {
                    const auto& file = task.inFlightFiles[index];
                    if (file.sourcePath != finalizationPath)
                        continue;
                    found                = true;
                    const uint32_t phase = static_cast<uint32_t>(file.phase);
                    const bool full      = file.totalBytes > 0u && file.completedBytes >= file.totalBytes;
                    if (! lastFinalizationFound || phase != lastFinalizationPhase || full != lastFinalizationFull)
                        AppendLog(std::format(L"FolderLiveProgressLifecycle: route={} file={} phase={} explicit={} terminal={} bytes={}/{} elapsed={}",
                                              route,
                                              finalizationPath,
                                              phase,
                                              file.explicitLifecycle,
                                              file.terminal,
                                              file.completedBytes,
                                              file.totalBytes,
                                              sampledCompleteTick == 0u ? 0u : now - sampledCompleteTick));
                    lastFinalizationPhase = phase;
                    lastFinalizationFull  = full;
                }
                if (! found && lastFinalizationFound)
                    AppendLog(std::format(L"FolderLiveProgressLifecycle: route={} file={} absent=true previousPhase={} full={} elapsed={}",
                                          route,
                                          finalizationPath,
                                          lastFinalizationPhase,
                                          lastFinalizationFull,
                                          sampledCompleteTick == 0u ? 0u : now - sampledCompleteTick));
                lastFinalizationFound = found;
            }
            const auto noteCapturePaint = [&]() noexcept
            {
                PopupLayoutDebugSnapshot afterCapture{};
                afterCapture.taskId = running._taskId;
                if (DebugGetFileOperationsPopupLayoutSnapshot(popup, afterCapture))
                    capturePaintGeneration = afterCapture.taskPaintGeneration;
            };
            totalsSeen = totalsSeen || (task.discoveryClosed && task.discoveredFileCount == kExpectedFiles &&
                                        task.discoveredDirectoryCount == kExpectedDirectories && task.discoveredTotalBytes == kExpectedBytes);
            if (task.discoveryClosed)
            {
                const std::wstring expectedSize = FormatStringResource(nullptr,
                                                                       IDS_FMT_FILEOPS_TRANSFER_AMOUNT,
                                                                       FormatBytesCompact(painted.taskPaintedTransferCompletedBytes),
                                                                       FormatBytesCompact(kExpectedBytes));
                // Completion numerators belong to this paint generation; denominators are the
                // fixture's independently known tree, never copied out of the formatted label.
                const auto& counts = painted.taskPaintedRecursiveCounts;
                const auto expectedCounts =
                    counts.filesCopiedKnown && counts.foldersCopiedKnown
                        ? FormatStringResource(
                              nullptr, IDS_FMT_FILEOPS_FOLDER_FILE_PROGRESS, counts.foldersCopied, kExpectedDirectories, counts.filesCopied, kExpectedFiles)
                    : counts.filesCopiedKnown
                        ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_FOUND_FOLDERS_FILE_PROGRESS, kExpectedDirectories, counts.filesCopied, kExpectedFiles)
                        : FormatStringResource(nullptr, IDS_FMT_FILEOPS_FOLDER_FILE_TOTALS, kExpectedDirectories, kExpectedFiles);
                const auto matchesCounts = [&](std::wstring_view ink, std::wstring_view descriptor, std::wstring_view name, std::wstring_view help) noexcept
                { return ink == expectedCounts && descriptor == expectedCounts && name == expectedCounts && help == expectedCounts; };
                auto* renderer = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup, GWLP_USERDATA));
                std::wstring countName;
                std::wstring countHelp;
                const bool countAccessible = renderer && renderer->DebugHostedRegionAccessibleTextForSelfTest(
                                                             std::format(L"FileOperations.Task.{}.Progress.Counts", task.taskId), countName, countHelp);
                const bool correctCounts   = totalsSeen && counts.discoveryClosed && counts.filesFound == kExpectedFiles &&
                                             counts.foldersFound == kExpectedDirectories && counts.filesCopied <= kExpectedFiles &&
                                             counts.foldersCopied <= kExpectedDirectories && countAccessible &&
                                             matchesCounts(painted.taskPaintedDiscoveryText, painted.taskCountsText, countName, countHelp);
                countRowSeen               = countRowSeen || correctCounts;
                if (correctCounts)
                {
                    const auto incorrect = FormatStringResource(nullptr,
                                                                IDS_FMT_FILEOPS_FOLDER_FILE_PROGRESS,
                                                                counts.foldersCopied,
                                                                kExpectedDirectories + 1u,
                                                                counts.filesCopied,
                                                                kExpectedFiles + 1u);
                    countOracleRejectsIncorrectLabels =
                        ! matchesCounts(incorrect, incorrect, incorrect, incorrect) && ! matchesCounts(expectedCounts, expectedCounts, L"", expectedCounts);
                }
                if (countRowSeen && ! countCaptureAttempted)
                {
                    ++countCaptureAttempts;
                    RECT client{};
                    // Reveal the count row if a very small viewport has scrolled it out of view.
                    // Reveal the owned scroll body and then restore the transfer view for the
                    // independent per-file pixel witnesses; no foreground input is involved.
                    bool detailsRevealed = true;
                    for (unsigned reveal = 0u; reveal < 16u; ++reveal)
                    {
                        const auto& rowBounds = painted.taskPaintedDiscoveryRect;
                        const auto& viewport  = painted.listViewportRect;
                        if (std::floor(rowBounds.top) >= std::ceil(viewport.top) && std::ceil(rowBounds.bottom) <= std::floor(viewport.bottom))
                            break;
                        // Completed cards can follow the running card: the list's bottom is not
                        // necessarily this counter. Reveal its actual painted rectangle instead.
                        SendMessageW(popup, WM_VSCROLL, rowBounds.top < viewport.top ? SB_LINEUP : SB_LINEDOWN, 0);
                        RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, painted))
                        {
                            detailsRevealed = false;
                            break;
                        }
                    }
                    const auto row = painted.taskPaintedDiscoveryRect;
                    const RECT countBounds{static_cast<LONG>(std::floor(row.left)),
                                           static_cast<LONG>(std::floor(row.top)),
                                           static_cast<LONG>(std::ceil(row.right)),
                                           static_cast<LONG>(std::ceil(row.bottom))};
                    const bool inside = detailsRevealed && GetClientRect(popup, &client) != FALSE && countBounds.left >= client.left &&
                                        countBounds.top >= client.top && countBounds.right <= client.right && countBounds.bottom <= client.bottom &&
                                        countBounds.right > countBounds.left && countBounds.bottom > countBounds.top;
                    std::array<RedSalamander::TestSupport::WindowScreenshotColorProbe, 1> probes{{{.clientBounds  = countBounds,
                                                                                                   .expectedColor = color(painted.taskPaintedDiscoveryColor),
                                                                                                   .tolerance     = 12u,
                                                                                                   .minimumMatchingPixels = 8u}}};
                    const HRESULT captured =
                        inside ? RedSalamander::TestSupport::SaveWindowScreenshot(popup, output / std::format(L"{}-counts.png", route), probes) : E_INVALIDARG;
                    countInkSeen          = inside && SUCCEEDED(captured);
                    countCaptureAttempted = countInkSeen || countCaptureAttempts >= 8u;
                    SendMessageW(popup, WM_VSCROLL, SB_TOP, 0);
                    RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    static_cast<void>(DebugGetFileOperationsPopupLayoutSnapshot(popup, painted));
                    noteCapturePaint();
                    AppendLog(std::format(L"FolderLiveProgressFrame: route={} stage=counts task={} text={} inside={} hr=0x{:08X} inkPixels={}",
                                          route,
                                          task.taskId,
                                          painted.taskPaintedDiscoveryText,
                                          inside,
                                          static_cast<unsigned long>(captured),
                                          probes[0].matchedPixels));
                }
                totalSizeRowSeen = totalSizeRowSeen || painted.taskPaintedTransferText == expectedSize;
            }
            // The natural route must observe a later ordinary WM_PAINT, not the repaint a screenshot
            // may request. Its visible records were produced before this pump's capture calls.
            if (naturalUpdates && painted.taskPaintGeneration <= capturePaintGeneration)
                return;
            for (size_t index = 0u; index < painted.taskPaintedFileProgressCount; ++index)
            {
                const auto& file               = painted.taskPaintedFileProgress[index];
                bool previousStillTransferring = false;
                for (size_t currentIndex = 0u; currentIndex < task.inFlightFileCount; ++currentIndex)
                {
                    const auto& current       = task.inFlightFiles[currentIndex];
                    previousStillTransferring = previousStillTransferring || (current.sourcePath == sampledPath && current.completedBytes < sampledTotal);
                }
                const bool newAttempt = ! lateCaptured && (! earlyCaptured || ! previousStillTransferring) && sampleAttempts < attemptedPaths.size() &&
                                        std::ranges::find(attemptedPaths, file.sourcePath) == attemptedPaths.end();
                if (newAttempt && task.discoveryClosed && file.totalBytes >= 2u * 1024u * 1024u && file.completedBytes * 10u >= file.totalBytes &&
                    file.completedBytes * 10u <= file.totalBytes * 3u)
                {
                    if (earlyCaptured)
                        AppendLog(
                            std::format(L"FolderLiveProgressSampling: route={} attempt={} file={} missedLateWindow=true", route, sampleAttempts, sampledPath));
                    attemptedPaths[sampleAttempts++] = file.sourcePath;
                    sampledPath                      = file.sourcePath;
                    sampledTotal                     = file.totalBytes;
                    sampledPaint                     = file;
                    const float middle               = (file.track.left + file.track.right) * 0.5f;
                    pixelRegion                      = {static_cast<LONG>(middle - 2.0f),
                                                        static_cast<LONG>(file.track.top + 2.0f),
                                                        static_cast<LONG>(middle + 2.0f),
                                                        static_cast<LONG>(file.track.bottom - 1.0f)};
                    std::array<RedSalamander::TestSupport::WindowScreenshotColorProbe, 1> probes{
                        {{.clientBounds = pixelRegion, .expectedColor = color(file.trackColor), .tolerance = 12u, .minimumMatchingPixels = 4u}}};
                    const auto captureName =
                        sampleAttempts == 1u ? std::format(L"{}-early.png", route) : std::format(L"{}-attempt{}-early.png", route, sampleAttempts);
                    const HRESULT captured = RedSalamander::TestSupport::SaveWindowScreenshot(popup, output / captureName, probes);
                    const auto trackColor  = color(file.trackColor);
                    const auto fillColor   = color(file.fillColor);
                    const int distance     = std::max({std::abs(static_cast<int>(GetRValue(trackColor)) - static_cast<int>(GetRValue(fillColor))),
                                                       std::abs(static_cast<int>(GetGValue(trackColor)) - static_cast<int>(GetGValue(fillColor))),
                                                       std::abs(static_cast<int>(GetBValue(trackColor)) - static_cast<int>(GetBValue(fillColor)))});
                    earlyCaptured          = SUCCEEDED(captured) && distance > 24;
                    earlyPaintGeneration   = earlyCaptured ? painted.taskPaintGeneration : 0u;
                    noteCapturePaint();
                    if (earlyCaptured && finalizationPath.empty())
                    {
                        finalizationPath  = sampledPath;
                        finalizationTotal = sampledTotal;
                        finalizationPaint = sampledPaint;
                    }
                    AppendLog(std::format(L"FolderLiveProgressFrame: route={} stage=early task={} file={} bytes={}/{} hr=0x{:08X} pixels={} dpi={}",
                                          route,
                                          task.taskId,
                                          sampledPath,
                                          file.completedBytes,
                                          file.totalBytes,
                                          static_cast<unsigned long>(captured),
                                          probes[0].matchedPixels,
                                          GetDpiForWindow(popup)));
                    break;
                }
                if (earlyCaptured && ! lateCaptured && file.sourcePath == sampledPath && file.totalBytes == sampledTotal &&
                    file.completedBytes * 10u >= file.totalBytes * 7u && file.completedBytes < file.totalBytes)
                {
                    const float middle = (file.track.left + file.track.right) * 0.5f;
                    const RECT laterRegion{static_cast<LONG>(middle - 2.0f),
                                           static_cast<LONG>(file.track.top + 2.0f),
                                           static_cast<LONG>(middle + 2.0f),
                                           static_cast<LONG>(file.track.bottom - 1.0f)};
                    const bool stableRegion = EqualRect(&pixelRegion, &laterRegion) != FALSE;
                    sampledPaint            = file;
                    std::array<RedSalamander::TestSupport::WindowScreenshotColorProbe, 1> probes{
                        {{.clientBounds = laterRegion, .expectedColor = color(file.fillColor), .tolerance = 12u, .minimumMatchingPixels = 4u}}};
                    const auto captureName =
                        sampleAttempts == 1u ? std::format(L"{}-late.png", route) : std::format(L"{}-attempt{}-late.png", route, sampleAttempts);
                    const HRESULT captured = RedSalamander::TestSupport::SaveWindowScreenshot(popup, output / captureName, probes);
                    lateCaptured           = SUCCEEDED(captured);
                    latePaintGeneration    = lateCaptured ? painted.taskPaintGeneration : 0u;
                    noteCapturePaint();
                    AppendLog(std::format(L"FolderLiveProgressFrame: route={} stage=late task={} file={} bytes={}/{} hr=0x{:08X} pixels={} sameRegion={}",
                                          route,
                                          task.taskId,
                                          sampledPath,
                                          file.completedBytes,
                                          file.totalBytes,
                                          static_cast<unsigned long>(captured),
                                          probes[0].matchedPixels,
                                          stableRegion));
                }
            }
            if (! finalizationPath.empty() && slowFinalization)
            {
                bool finalizingPhaseActive = false;
                for (size_t index = 0u; index < task.inFlightFileCount; ++index)
                {
                    const auto& file = task.inFlightFiles[index];
                    if (file.sourcePath == finalizationPath && file.totalBytes == finalizationTotal && file.completedBytes >= file.totalBytes &&
                        ! file.terminal && (! file.explicitLifecycle || file.phase == FileSystemFileWorkPhase::Finalizing))
                    {
                        finalizingPhaseActive = true;
                        if (sampledCompleteTick == 0u)
                            sampledCompleteTick = now;
                    }
                }
                if (sampledCompleteTick != 0u && now - sampledCompleteTick >= 650u && ! finalizationObserved)
                {
                    finalizationObserved = true;
                    for (size_t index = 0u; index < painted.taskPaintedFileProgressCount; ++index)
                    {
                        const auto& file = painted.taskPaintedFileProgress[index];
                        if (file.sourcePath == finalizationPath)
                        {
                            finalizationMeter  = finalizingPhaseActive;
                            knownTotalRetained = finalizingPhaseActive && file.totalBytes == finalizationTotal && file.determinate;
                            finalizationPaint  = file;
                        }
                    }
                    const float middle = (finalizationPaint.track.left + finalizationPaint.track.right) * 0.5f;
                    std::array<RedSalamander::TestSupport::WindowScreenshotColorProbe, 1> probes{
                        {{.clientBounds          = {static_cast<LONG>(middle - 2.0f),
                                                    static_cast<LONG>(finalizationPaint.track.top + 2.0f),
                                                    static_cast<LONG>(middle + 2.0f),
                                                    static_cast<LONG>(finalizationPaint.track.bottom - 1.0f)},
                          .expectedColor         = color(finalizationPaint.fillColor),
                          .tolerance             = 12u,
                          .minimumMatchingPixels = 4u}}};
                    const HRESULT captured = RedSalamander::TestSupport::SaveWindowScreenshot(popup, output / std::format(L"{}-finalizing.png", route), probes);
                    const bool released    = SetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_HOLD_NONEMPTY_COMMIT", nullptr) != FALSE;
                    finalizationMeter      = finalizationMeter && SUCCEEDED(captured);
                    finalizationMeter      = finalizationMeter && released;
                    noteCapturePaint();
                    AppendLog(std::format(L"FolderLiveProgressFrame: route={} stage=finalizing task={} file={} elapsed={} rows={} meter={} knownSize={} "
                                          L"activePhase={} hr=0x{:08X}",
                                          route,
                                          task.taskId,
                                          finalizationPath,
                                          now - sampledCompleteTick,
                                          painted.taskPaintedFileProgressCount,
                                          finalizationMeter,
                                          knownTotalRetained,
                                          finalizingPhaseActive,
                                          static_cast<unsigned long>(captured)));
                    Debug::Perf::Emit(L"FileOps.SelfTest.FinalizationCaptureLease",
                                      route,
                                      (now - sampledCompleteTick) * 1000u,
                                      650u,
                                      30'000u,
                                      finalizationMeter && knownTotalRetained ? S_OK : E_FAIL);
                }
            }
        };
        const auto onPrompt = [&](FolderWindow::FileOperationState::Task& running)
        {
            const auto prompt = TryGetConflictPromptCopy(&running);
            if (! overwrite || ! prompt.has_value())
                return false;
            return running.SubmitConflictDecision(prompt->generation, FolderWindow::FileOperationState::Task::ConflictAction::Overwrite, true);
        };
        const auto result = scope.Run(FILESYSTEM_COPY,
                                      destinationFs,
                                      {source},
                                      destination,
                                      kSpeed,
                                      onPump,
                                      {},
                                      {},
                                      onPrompt,
                                      FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                      FILESYSTEM_FLAG_NONE,
                                      8u);
        record(result.has_value() && SUCCEEDED(result.value().hr) && result.value().discoveredFiles == kExpectedFiles &&
                   result.value().discoveredTotalBytes == kExpectedBytes,
               route,
               L"successful real copy with exact final file/byte totals");
        record(result.has_value() && result.value().copiedFileCountKnown && result.value().copiedFileCount == kExpectedFiles,
               route,
               L"actual publication counts every child once, including the empty file");
        record(result.has_value() && result.value().progressCompletedBytes == kExpectedBytes,
               route,
               L"written-byte total includes every child rather than saturating at the first child size");
        record(totalsSeen, route, L"exact child-file, folder and byte totals close during copy");
        record(countRowSeen, route, L"child-file total remains painted after discovery closes");
        record(countOracleRejectsIncorrectLabels, route, L"live count oracle rejects wrong totals and missing accessible text");
        record(countInkSeen, route, L"closed child-file count has visible text ink inside the popup client");
        record(totalSizeRowSeen, route, L"total byte size remains painted after discovery closes");
        record(earlyCaptured && lateCaptured, route, L"the same real file advances track pixels to fill pixels");
        if (naturalUpdates)
        {
            AppendLog(std::format(L"FolderLiveProgressNaturalPaint: route={} file={} earlyGeneration={} lateGeneration={} forcedRedraw=false",
                                  route,
                                  sampledPath,
                                  earlyPaintGeneration,
                                  latePaintGeneration));
            record(earlyCaptured && lateCaptured && earlyPaintGeneration > 0u && latePaintGeneration > earlyPaintGeneration,
                   route,
                   L"ordinary automatic paints advance the same real file independently of capture repaints");
        }
        if (slowFinalization)
        {
            record(finalizationObserved && finalizationMeter, route, L"fully written file meter persists during slow finalization");
            record(knownTotalRetained, route, L"known individual-file total survives finalization");
        }
        wil::com_ptr<IFileSystemIO> destinationIo;
        bool contentMatches = SUCCEEDED(destinationFs->QueryInterface(__uuidof(IFileSystemIO), destinationIo.put_void()));
        constexpr std::array<std::wstring_view, 4> names{{L"first.bin", L"nested/second.bin", L"nested/deep/third.bin", L"empty.bin"}};
        for (const auto name : names)
        {
            std::string copied;
            const unsigned char expected = name == L"first.bin" ? 0x31u : name == L"nested/second.bin" ? 0x52u : name == L"empty.bin" ? 0u : 0x73u;
            const size_t bytes           = name == L"empty.bin" ? 0u : 32u * 1024u * 1024u;
            const auto destinationFile   = destination / L"tree" / name;
            contentMatches = ReadFileTextFsIo(destinationIo, destinationFile, copied) && copied.size() == bytes &&
                             std::ranges::all_of(copied, [&](char value) { return static_cast<unsigned char>(value) == expected; }) && contentMatches;
        }
        record(contentMatches, route, L"independent destination content including the empty child");
        return CancelOutstandingSelfTestOperations(operations);
    };
    const auto runOpenTraversal = [&]()
    {
        // A sequential step lists at most four lookahead directories. Even the two steps before
        // the first payload leave some of these 32 nonempty siblings unlisted, regardless of order.
        constexpr unsigned long kWideFiles       = 32u;
        constexpr unsigned long kWideDirectories = 33u;
        constexpr uint64_t kWideLeafBytes        = 2u * 1024u * 1024u;
        constexpr uint64_t kWideBytes            = kWideFiles * kWideLeafBytes;
        constexpr uint64_t kWideSpeed            = 2u * 1024u * 1024u;
        constexpr std::wstring_view route        = L"bridge-open-discovery";
        const auto wideSource                    = root / L"open-tree";
        const auto wideDestination               = std::filesystem::path(L"/folder-live-progress/open");
        if (! RecreateEmptyDirectory(wideSource))
            return false;
        for (unsigned int index = 0u; index < kWideFiles; ++index)
        {
            if (! WriteFilledTestFile(wideSource / std::format(L"branch-{0:02}", index) / L"payload.bin", kWideLeafBytes, static_cast<BYTE>(0x40u + index)))
                return false;
        }
        struct OpenPrefix final
        {
            unsigned long files;
            unsigned long directories;
            uint64_t bytes;
        };
        constexpr auto prefixes = []() noexcept
        {
            std::array<OpenPrefix, kWideFiles - 1u> values{};
            for (unsigned long index = 0u; index < values.size(); ++index)
                values[index] = {index + 1u, kWideDirectories, (index + 1u) * kWideLeafBytes};
            return values;
        }();
        std::array<bool, prefixes.size()> attempted{};
        bool openInkSeen                                          = false;
        bool closureSeen                                          = false;
        ULONGLONG lastSample                                      = 0u;
        uint64_t lastOpenFiles                                    = UINT64_MAX;
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
        const auto onPump                                         = [&](FolderWindow::FileOperationState::Task& running)
        {
            const ULONGLONG now = GetTickCount64();
            if (now - lastSample < 100u)
                return;
            lastSample = now;
            operations->DebugEnsurePopupVisibleForSelfTest();
            const HWND popup = operations->GetPopupHwndForSelfTest();
            if (! popup)
                return;
            RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
            TaskSnapshot task{};
            PopupLayoutDebugSnapshot painted{};
            painted.taskId = running._taskId;
            if (! DebugGetFileOperationsPopupTaskSnapshot(popup, running._taskId, task) || ! DebugGetFileOperationsPopupLayoutSnapshot(popup, painted) ||
                task.conflict.active || task.finished)
                return;
            closureSeen = closureSeen || (task.discoveryClosed && task.discoveredFileCount == kWideFiles && task.discoveredDirectoryCount == kWideDirectories &&
                                          task.discoveredTotalBytes == kWideBytes);
            if (task.discoveryClosed || openInkSeen)
                return;
            if (lastOpenFiles != task.discoveredFileCount)
            {
                lastOpenFiles = task.discoveredFileCount;
                AppendLog(std::format(L"FolderLiveProgressOpen: files={} directories={} bytes={} text={} mode=sequential siblings=32",
                                      task.discoveredFileCount,
                                      task.discoveredDirectoryCount,
                                      task.discoveredTotalBytes,
                                      painted.taskPaintedDiscoveryText));
            }
            for (size_t index = 0u; index < prefixes.size(); ++index)
            {
                const auto& prefix = prefixes[index];
                if (attempted[index] || task.discoveredFileCount != prefix.files || task.discoveredDirectoryCount != prefix.directories ||
                    task.discoveredTotalBytes != prefix.bytes)
                    continue;
                const auto expected = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNTING_FOUND, prefix.directories, prefix.files);
                if (painted.taskPaintedDiscoveryText != expected)
                    continue;
                attempted[index] = true;
                RECT client{};
                const auto row = painted.taskPaintedDiscoveryRect;
                const RECT bounds{static_cast<LONG>(std::floor(row.left)),
                                  static_cast<LONG>(std::floor(row.top)),
                                  static_cast<LONG>(std::ceil(row.right)),
                                  static_cast<LONG>(std::ceil(row.bottom))};
                const bool inside = GetClientRect(popup, &client) != FALSE && bounds.left >= client.left && bounds.top >= client.top &&
                                    bounds.right <= client.right && bounds.bottom <= client.bottom && bounds.right > bounds.left && bounds.bottom > bounds.top;
                std::array<RedSalamander::TestSupport::WindowScreenshotColorProbe, 1> probes{
                    {{.clientBounds = bounds, .expectedColor = color(painted.taskPaintedDiscoveryColor), .tolerance = 12u, .minimumMatchingPixels = 8u}}};
                const HRESULT captured =
                    inside ? RedSalamander::TestSupport::SaveWindowScreenshot(popup, output / std::format(L"bridge-open-prefix-{}.png", index), probes)
                           : E_INVALIDARG;
                TaskSnapshot afterTask{};
                PopupLayoutDebugSnapshot afterPaint{};
                afterPaint.taskId       = running._taskId;
                const bool remainedOpen = DebugGetFileOperationsPopupTaskSnapshot(popup, running._taskId, afterTask) &&
                                          DebugGetFileOperationsPopupLayoutSnapshot(popup, afterPaint) && ! afterTask.discoveryClosed &&
                                          afterTask.discoveredFileCount == prefix.files && afterTask.discoveredDirectoryCount == prefix.directories &&
                                          afterTask.discoveredTotalBytes == prefix.bytes && afterPaint.taskPaintedDiscoveryText == expected;
                openInkSeen             = inside && SUCCEEDED(captured) && remainedOpen;
                AppendLog(
                    std::format(L"FolderLiveProgressOpenCapture: prefix={} files={} directories={} bytes={} inside={} hr=0x{:08X} inkPixels={} remainedOpen={}",
                                index,
                                prefix.files,
                                prefix.directories,
                                prefix.bytes,
                                inside,
                                static_cast<unsigned long>(captured),
                                probes[0].matchedPixels,
                                remainedOpen));
                break;
            }
        };
        const auto result = scope.Run(FILESYSTEM_COPY,
                                      state.fsDummy,
                                      {wideSource},
                                      wideDestination,
                                      kWideSpeed,
                                      onPump,
                                      {},
                                      {},
                                      {},
                                      FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                      FILESYSTEM_FLAG_NONE,
                                      8u);
        record(openInkSeen, route, L"real open traversal paints independently defined positive count/size prefixes with visible ink");
        record(closureSeen, route, L"wide traversal closes on exact 32-file/33-folder/64-MiB totals while copying");
        record(result.has_value() && SUCCEEDED(result.value().hr) && result.value().discoveredFiles == kWideFiles &&
                   result.value().discoveredDirectories == kWideDirectories && result.value().discoveredTotalBytes == kWideBytes &&
                   result.value().progressCompletedBytes == kWideBytes && result.value().copiedFileCountKnown && result.value().copiedFileCount == kWideFiles,
               route,
               L"wide final discovery, written bytes and child publication totals are exact");
        wil::com_ptr<IFileSystemIO> destinationIo;
        bool contentMatches = SUCCEEDED(state.fsDummy->QueryInterface(__uuidof(IFileSystemIO), destinationIo.put_void()));
        for (unsigned int index = 0u; index < kWideFiles; ++index)
        {
            std::string copied;
            const auto destinationFile = wideDestination / L"open-tree" / std::format(L"branch-{0:02}", index) / L"payload.bin";
            const auto expected        = static_cast<unsigned char>(0x40u + index);
            contentMatches = ReadFileTextFsIo(destinationIo, destinationFile, copied) && copied.size() == kWideLeafBytes &&
                             std::ranges::all_of(copied, [&](char value) { return static_cast<unsigned char>(value) == expected; }) && contentMatches;
        }
        record(contentMatches, route, L"independent content of every wide-tree payload");
        return CancelOutstandingSelfTestOperations(operations);
    };
    // Only the existing provider slow-commit hook is used: the renderer and real callbacks are unchanged.
    using Env = RedSalamander::TestSupport::ScopedEnvironmentVariable;
    try
    {
        Env slowCommit(L"REDSALAMANDER_FILEOPS_SLOW_COMMIT_MS", std::wstring(L"1500"));
        if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency1) || ! RecreateEmptyDirectory(root / L"local-serial"))
            return reject(L"serial Local configuration/destination unavailable");
        if (! run(L"local-serial", state.fsLocal, root / L"local-serial", false, false, true) ||
            ! run(L"local-overwrite", state.fsLocal, root / L"local-serial", false, true, true))
            return reject(L"a serial/overwrite worker did not reach its quiet point");
        if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4) || ! RecreateEmptyDirectory(root / L"local-parallel"))
            return reject(L"parallel Local configuration/destination unavailable");
        if (! run(L"local-parallel", state.fsLocal, root / L"local-parallel", false, false, true))
            return reject(L"a parallel worker did not reach its quiet point");
        const std::wstring dummyRoot = L"/folder-live-progress";
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot + L"/parallel") ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot + L"/verify") || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot + L"/open"))
            return reject(L"bridge destinations unavailable");
        if (! run(L"bridge-parallel", state.fsDummy, std::filesystem::path(dummyRoot + L"/parallel"), false, false, false))
            return reject(L"a bridge worker did not reach its quiet point");
        if (! run(L"bridge-verify", state.fsDummy, std::filesystem::path(dummyRoot + L"/verify"), true, false, false))
            return reject(L"a bridge worker did not reach its quiet point");
        if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency1))
            return reject(L"sequential open-traversal configuration unavailable");
        if (! runOpenTraversal())
            return reject(L"the separate wide sequential bridge worker did not reach its quiet point");
        if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4))
            return reject(L"parallel configuration restoration after open traversal failed");
        const auto emptyTree = root / L"empty-tree";
        if (! RecreateEmptyDirectory(emptyTree) || ! RecreateEmptyDirectory(root / L"empty-local") ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot + L"/empty"))
            return reject(L"empty directory fixture unavailable");
        const auto emptyLocal   = scope.Run(FILESYSTEM_COPY, state.fsLocal, {emptyTree}, root / L"empty-local", 0u);
        const auto emptyBridge  = scope.Run(FILESYSTEM_COPY, state.fsDummy, {emptyTree}, std::filesystem::path(dummyRoot + L"/empty"), 0u);
        const auto emptyCorrect = [](const std::optional<CompletedTaskInfo>& info)
        {
            return info.has_value() && SUCCEEDED(info.value().hr) && info.value().discoveryClosed && info.value().discoveredFiles == 0u &&
                   info.value().discoveredTotalBytes == 0u && info.value().copiedFileCountKnown && info.value().copiedFileCount == 0u &&
                   info.value().discoveredBytesKnown;
        };
        record(emptyCorrect(emptyLocal), L"empty-local", L"real empty root closes on known zero files, copied files and payload bytes");
        record(emptyCorrect(emptyBridge), L"empty-bridge", L"real empty bridge root closes on known zero files, copied files and payload bytes");

        // The native slow-commit knob does not delay Dummy. Reuse the canonical real publication
        // pause for the small bridge file: no alternate callbacks or renderer are selected.
        const auto smallSource = root / L"small-tree";
        if (! WriteFilledTestFile(smallSource / L"known-small.bin", 1024u * 1024u, 0x64u) ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot + L"/small") ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot + L"/small/small-tree"))
            return reject(L"small bridge fixture unavailable");
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
        SetFileOpsLiveOutputPublishedPauseForSelfTest(true);
        const auto releasePublication = wil::scope_exit([]() noexcept { ReleaseFileOpsLiveOutputPublishedPauseForSelfTest(); });
        bool publicationObserved      = false;
        bool smallMeterKnown          = false;
        bool smallFrameCaptured       = false;
        bool smallPublishedFact       = false;
        bool heldFileContent          = false;
        bool heldFilePublishing       = false;
        ULONGLONG publicationTick     = 0u;
        const auto smallResult        = scope.Run(FILESYSTEM_COPY,
                                                  state.fsDummy,
                                                  {smallSource},
                                                  std::filesystem::path(dummyRoot + L"/small"),
                                                  kSpeed,
                                                  [&](FolderWindow::FileOperationState::Task& running)
        {
            if (publicationObserved || ! HasFileOpsLiveOutputPublishedPauseEnteredForSelfTest())
                return;
            const auto now = GetTickCount64();
            if (publicationTick == 0u)
                publicationTick = now;
            if (now - publicationTick < 650u)
                return;
            publicationObserved = true;
            operations->DebugEnsurePopupVisibleForSelfTest();
            const HWND popup = operations->GetPopupHwndForSelfTest();
            PopupLayoutDebugSnapshot painted{};
            painted.taskId = running._taskId;
            RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
            if (popup && DebugGetFileOperationsPopupLayoutSnapshot(popup, painted))
            {
                TaskSnapshot live{};
                if (DebugGetFileOperationsPopupTaskSnapshot(popup, running._taskId, live))
                {
                    for (size_t index = 0u; index < live.inFlightFileCount; ++index)
                    {
                        const auto& file   = live.inFlightFiles[index];
                        heldFilePublishing = heldFilePublishing || (file.sourcePath == (smallSource / L"known-small.bin").wstring() && file.explicitLifecycle &&
                                                                    ! file.terminal && file.phase == FileSystemFileWorkPhase::Publishing);
                    }
                }
                smallPublishedFact =
                    running._copiedFileCountKnown.load(std::memory_order_acquire) && running._copiedFileCount.load(std::memory_order_acquire) == 1u;
                wil::com_ptr<IFileSystemIO> heldIo;
                std::string heldContent;
                heldFileContent = SUCCEEDED(state.fsDummy->QueryInterface(__uuidof(IFileSystemIO), heldIo.put_void())) &&
                                  ReadFileTextFsIo(heldIo, std::filesystem::path(dummyRoot + L"/small/small-tree/known-small.bin"), heldContent) &&
                                  heldContent.size() == 1024u * 1024u &&
                                  std::ranges::all_of(heldContent, [](char value) { return value == static_cast<char>(0x64u); });
                for (size_t index = 0u; index < painted.taskPaintedFileProgressCount; ++index)
                {
                    const auto& file = painted.taskPaintedFileProgress[index];
                    if (file.sourcePath != (smallSource / L"known-small.bin").wstring())
                        continue;
                    smallMeterKnown    = file.determinate && file.totalBytes == 1024u * 1024u && file.completedBytes == file.totalBytes;
                    const float middle = (file.fill.left + file.fill.right) * 0.5f;
                    std::array<RedSalamander::TestSupport::WindowScreenshotColorProbe, 1> probes{{{.clientBounds = {static_cast<LONG>(middle - 2.0f),
                                                                                                                    static_cast<LONG>(file.fill.top + 2.0f),
                                                                                                                    static_cast<LONG>(middle + 2.0f),
                                                                                                                    static_cast<LONG>(file.fill.bottom - 1.0f)},
                                                                                                   .expectedColor         = color(file.fillColor),
                                                                                                   .tolerance             = 12u,
                                                                                                   .minimumMatchingPixels = 4u}}};
                    smallFrameCaptured = SUCCEEDED(RedSalamander::TestSupport::SaveWindowScreenshot(popup, output / L"bridge-small-publishing.png", probes));
                    AppendLog(std::format(L"FolderLiveProgressFrame: route=bridge-small stage=publishing task={} file={} bytes={}/{} determinate={} pixels={}",
                                          running._taskId,
                                          file.sourcePath,
                                          file.completedBytes,
                                          file.totalBytes,
                                          file.determinate,
                                          probes[0].matchedPixels));
                }
                AppendLog(std::format(L"FolderLiveProgressSmallPublication: task={} rows={} knownMeter={} publishedCount={} countKnown={} heldContent={}",
                                      running._taskId,
                                      painted.taskPaintedFileProgressCount,
                                      smallMeterKnown,
                                      running._copiedFileCount.load(std::memory_order_acquire),
                                      running._copiedFileCountKnown.load(std::memory_order_acquire),
                                      heldFileContent));
                if (! smallFrameCaptured)
                    static_cast<void>(RedSalamander::TestSupport::SaveWindowScreenshot(popup, output / L"bridge-small-publishing-diagnostic.png"));
            }
            ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        });
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        record(smallResult.has_value() && SUCCEEDED(smallResult.value().hr) && smallResult.value().discoveredFiles == 1u &&
                   smallResult.value().discoveredTotalBytes == 1024u * 1024u,
               L"bridge-small",
               L"real known-size small-file copy completes with exact totals");
        record(publicationObserved && smallMeterKnown && smallFrameCaptured,
               L"bridge-small",
               L"known size and fully written meter remain visible during Publishing");
        record(smallPublishedFact && heldFileContent && heldFilePublishing,
               L"bridge-small",
               L"actual file publication is counted before terminal and independently readable while held in its Publishing phase");
        wil::com_ptr<IFileSystemIO> smallDestinationIo;
        std::string smallContent;
        const bool smallContentMatches =
            SUCCEEDED(state.fsDummy->QueryInterface(__uuidof(IFileSystemIO), smallDestinationIo.put_void())) &&
            ReadFileTextFsIo(smallDestinationIo, std::filesystem::path(dummyRoot + L"/small/small-tree/known-small.bin"), smallContent) &&
            smallContent.size() == 1024u * 1024u && std::ranges::all_of(smallContent, [](char value) { return value == static_cast<char>(0x64u); });
        record(smallContentMatches, L"bridge-small", L"independent exact small-file destination content");
        // Real owned-stage publication followed by a provider-style committed-error HRESULT.
        const auto committedErrorFolder    = root / L"committed-error";
        const auto committedErrorFile      = committedErrorFolder / L"known-small.bin";
        const std::wstring committedSource = dummyRoot + L"/small/small-tree/known-small.bin";
        if (! RecreateEmptyDirectory(committedErrorFolder) || ! smallContentMatches ||
            ! SetFileOpsBridgeCommittedErrorPublicationForSelfTest(committedSource, committedErrorFile))
            return reject(L"exact owned committed-error fixture could not be armed");
        const auto clearCommittedError = wil::scope_exit([]() noexcept { static_cast<void>(TakeAndClearFileOpsBridgeCommittedErrorPublicationForSelfTest()); });
        const auto committedError      = scope.Run(FILESYSTEM_COPY,
                                                   state.fsLocal,
                                                   {std::filesystem::path(committedSource)},
                                                   committedErrorFolder,
                                                   0u,
                                                   {},
                                                   {},
                                                   state.fsDummy,
                                                   {},
                                                   FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                   FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        if (! CancelOutstandingSelfTestOperations(operations))
            return reject(L"committed-error workers did not become quiet before hook restoration");
        const bool faultConsumed = TakeAndClearFileOpsBridgeCommittedErrorPublicationForSelfTest();
        const bool faultCleared  = ! TakeAndClearFileOpsBridgeCommittedErrorPublicationForSelfTest();
        std::string committedContent;
        wil::com_ptr<IFileSystemIO> committedIo;
        const bool committedContentMatches = SUCCEEDED(state.fsLocal->QueryInterface(__uuidof(IFileSystemIO), committedIo.put_void())) &&
                                             ReadFileTextFsIo(committedIo, committedErrorFile, committedContent) && committedContent == smallContent;
        const auto publishedFailure        = [](const std::optional<CompletedTaskInfo>& result)
        {
            return result.has_value() && result.value().sourceItemResults.size() == 1u && result.value().sourceItemResults[0].has_value() &&
                   result.value().sourceItemResults[0].value().publication == FileOperations::PublicationState::Published &&
                   result.value().sourceItemResults[0].value().completion == FileOperations::ItemCompletion::Failed;
        };
        const auto logFailure = [&](std::wstring_view route, const std::optional<CompletedTaskInfo>& result, bool content, uint64_t attempts)
        {
            const auto* typed = result.has_value() && result.value().sourceItemResults.size() == 1u && result.value().sourceItemResults[0].has_value()
                                    ? &result.value().sourceItemResults[0].value()
                                    : nullptr;
            AppendLog(std::format(L"FolderLiveProgressPublishedFailure: route={} result={} hr=0x{:08X} copied={} countKnown={} content={} attempts={} roots={} "
                                  L"publication={} completion={}",
                                  route,
                                  result.has_value(),
                                  static_cast<unsigned long>(result.has_value() ? result.value().hr : E_PENDING),
                                  result.has_value() ? result.value().copiedFileCount : 0u,
                                  result.has_value() && result.value().copiedFileCountKnown,
                                  content,
                                  attempts,
                                  result.has_value() ? result.value().sourceItemResults.size() : 0u,
                                  typed ? static_cast<unsigned int>(typed->publication) : UINT32_MAX,
                                  typed ? static_cast<unsigned int>(typed->completion) : UINT32_MAX));
        };
        logFailure(L"bridge-committed-error", committedError, committedContentMatches, faultConsumed ? 1u : 0u);
        record(committedError.has_value() && FAILED(committedError.value().hr) && committedError.value().copiedFileCountKnown &&
                   committedError.value().copiedFileCount == 1u && committedContentMatches && faultConsumed && faultCleared && publishedFailure(committedError),
               L"bridge-committed-error",
               L"real committed file remains counted and readable when publication returns failure");
        if (! CancelOutstandingSelfTestOperations(operations))
            return reject(L"committed-error workers did not become quiet");
        // Real atomic-final Commit succeeded, but the subsequent destination-size proof fails.
        const std::wstring proofRoot = dummyRoot + L"/proof-failure";
        if (! EnsureDummyFolderExists(state.fsDummy.get(), proofRoot))
            return reject(L"post-publication proof fixture unavailable");
        static_cast<void>(TakeFileOpsBridgeReportWrongDestinationSizeAttemptsForSelfTest());
        SetFileOpsBridgeReportWrongDestinationSizeForSelfTest(1u);
        const auto clearProofFault = wil::scope_exit([]() noexcept { SetFileOpsBridgeReportWrongDestinationSizeForSelfTest(0u); });
        const auto proofFailure    = scope.Run(FILESYSTEM_COPY,
                                               state.fsDummy,
                                               {smallSource / L"known-small.bin"},
                                               std::filesystem::path(proofRoot),
                                               0u,
                                               {},
                                               {},
                                               {},
                                               {},
                                               FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                               FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        if (! CancelOutstandingSelfTestOperations(operations))
            return reject(L"proof-failure workers did not become quiet before hook restoration");
        const auto proofFaults = TakeFileOpsBridgeReportWrongDestinationSizeAttemptsForSelfTest();
        SetFileOpsBridgeReportWrongDestinationSizeForSelfTest(0u);
        std::string proofContent;
        const bool proofContentMatches =
            ReadFileTextFsIo(smallDestinationIo, std::filesystem::path(proofRoot + L"/known-small.bin"), proofContent) && proofContent == smallContent;
        logFailure(L"bridge-proof-failure", proofFailure, proofContentMatches, proofFaults);
        record(proofFailure.has_value() && proofFailure.value().hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) && proofFailure.value().copiedFileCountKnown &&
                   proofFailure.value().copiedFileCount == 1u && proofContentMatches && proofFaults == 1u && publishedFailure(proofFailure),
               L"bridge-proof-failure",
               L"real atomic-final publication stays counted and readable after a failing size proof");
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::system_error& error)
    {
        // Environment setup is a throwing test boundary; no transfer is admitted on setup failure.
        Debug::Error(L"Folder live progress environment setup failed: {}", error.code().value());
        return reject(L"slow-commit environment scope unavailable");
    }
    for (const auto& defect : defects)
        AppendLog(std::format(L"Folder live progress: FAILED: {}", defect));
    const bool quietAndRestored      = scope.Close();
    const bool outerSettingsRestored = quietAndRestored && restoreSettings();
    return defects.empty() && quietAndRestored && outerSettingsRestored;
}

[[nodiscard]] bool RunDiscoveryScopeMultipleFoldersLiveProgressChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Multiple folder live progress: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"initial worker quiet point or Local provider unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    std::string previousConfiguration;
    if (! BackupPluginConfiguration(state.infoLocal.get(), previousConfiguration) || ! SetPluginConfiguration(state.infoLocal.get(), previousConfiguration) ||
        ! BackupPluginConfiguration(state.infoLocal.get(), previousConfiguration))
        return reject(L"configuration backup unavailable");
    bool configurationRestored      = false;
    const auto restoreConfiguration = [&]() noexcept
    {
        if (! CancelOutstandingSelfTestOperations(operations))
            return false;
        std::string readback;
        configurationRestored = SetPluginConfiguration(state.infoLocal.get(), previousConfiguration) &&
                                BackupPluginConfiguration(state.infoLocal.get(), readback) && readback == previousConfiguration;
        return configurationRestored;
    };
    const auto restore = wil::scope_exit([&]() noexcept
    {
        if (! configurationRestored)
            static_cast<void>(restoreConfiguration());
    });
    const auto root    = state.tempRoot / L"multiple-folder-live-progress";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(root / L"out") || ! SeedClosureFiles(root / L"one", 4u, L"one") ||
        ! SeedClosureFiles(root / L"two" / L"nested", 4u, L"two"))
        return reject(L"two-root fixture unavailable");
    std::error_code linkError;
    bool semanticLinksAvailable = true;
    for (const auto& name : {L"one", L"two"})
    {
        const auto target = root / name / (std::wstring(name) == L"one" ? L"one-00.bin" : L"nested/two-00.bin");
        const auto link   = root / name / (std::wstring(name) + L"-link.bin");
        if (! Common::Testing::CreateSandboxedSymbolicLink(link, target, SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE, linkError))
        {
            if (linkError.value() != ERROR_PRIVILEGE_NOT_HELD)
                return reject(std::format(L"sandboxed file-link fixture failed: {}", linkError.value()));
            semanticLinksAvailable = false;
            AppendLog(std::format(L"MultipleFolderLiveProgressSubcase: source={} status=skipped reason=ERROR_PRIVILEGE_NOT_HELD(1314)", link.wstring()));
        }
    }
    if (! semanticLinksAvailable)
    {
        if (! SelfTest::RemoveAll(root / L"one" / L"one-link.bin") || ! SelfTest::RemoveAll(root / L"two" / L"two-link.bin"))
            return reject(L"partial owned link fixture could not be removed");
        AppendLog(L"MultipleFolderLiveProgressSubcase: PerItemSemanticLinks=skipped NativeBulkSemanticLinks=skipped reason=Windows file-link creation "
                  L"privilege unavailable; regular-file I/O remains required.");
    }
    const unsigned long expectedFiles = semanticLinksAvailable ? 10u : 8u;
    const auto removeRoot             = wil::scope_exit([&]() noexcept
    {
        if (CancelOutstandingSelfTestOperations(operations))
            static_cast<void>(SelfTest::RemoveAll(root));
    });
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4))
        return reject(L"parallel Local configuration unavailable");
    ClosureCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    bool allRoutesCorrect                                     = true;
    for (const auto mode : {FolderWindow::FileOperationState::ExecutionMode::PerItem})
    {
        const auto destination = root / (mode == FolderWindow::FileOperationState::ExecutionMode::PerItem ? L"out" : L"out-bulk");
        if (! RecreateEmptyDirectory(destination))
            return reject(L"multiple-root destination unavailable");
        uint64_t previousBytes  = 0u;
        uint64_t previousCopied = 0u;
        bool monotonic          = true;
        bool closedDuringCopy   = false;
        bool boundedRows        = true;
        bool bothRootsVisible   = false;
        const auto result       = scope.Run(FILESYSTEM_COPY,
                                            state.fsLocal,
                                            {root / L"one", root / L"two"},
                                            destination,
                                            kClosureSpeedLimit,
                                            [&](FolderWindow::FileOperationState::Task& running)
        {
            const uint64_t bytes  = running._publishedProgressCompletedBytes.load(std::memory_order_acquire);
            const uint64_t copied = running._copiedFileCount.load(std::memory_order_acquire);
            monotonic             = monotonic && bytes >= previousBytes && bytes <= 8u * kClosureFileBytes && copied >= previousCopied && copied <= 10u;
            previousBytes         = bytes;
            previousCopied        = copied;
            closedDuringCopy      = closedDuringCopy || (running._discoveryClosed.load(std::memory_order_acquire) && bytes < 8u * kClosureFileBytes);
            std::scoped_lock lock(running._inFlightFilesMutex);
            boundedRows    = boundedRows && running._inFlightFileCount <= 16u;
            bool oneActive = false;
            bool twoActive = false;
            for (size_t index = 0u; index < running._inFlightFileCount; ++index)
            {
                const auto& file = running._inFlightFiles[index];
                if (file.explicitLifecycle && ! file.terminal && file.kind == FileSystemFileWorkKind::RegularFile)
                {
                    oneActive = oneActive || file.sourcePath.starts_with((root / L"one").wstring());
                    twoActive = twoActive || file.sourcePath.starts_with((root / L"two").wstring());
                }
            }
            bothRootsVisible = bothRootsVisible || (oneActive && twoActive);
        },
                                      {},
                                      {},
                                      {},
                                      mode);
        bool contentMatches     = true;
        for (unsigned long index = 0u; index < 4u; ++index)
        {
            contentMatches =
                FilesEqualBytes(root / L"one" / std::format(L"one-{:02}.bin", index), destination / L"one" / std::format(L"one-{:02}.bin", index)) &&
                FilesEqualBytes(root / L"two" / L"nested" / std::format(L"two-{:02}.bin", index),
                                destination / L"two" / L"nested" / std::format(L"two-{:02}.bin", index)) &&
                contentMatches;
        }
        for (const auto& name : {L"one", L"two"})
        {
            if (! semanticLinksAvailable)
                break;
            const auto linkName     = std::wstring(name) + L"-link.bin";
            const auto sourceTarget = std::filesystem::read_symlink(root / name / linkName, linkError);
            const bool sourceValid  = ! linkError;
            const auto copiedTarget = std::filesystem::read_symlink(destination / name / linkName, linkError);
            contentMatches          = contentMatches && sourceValid && ! linkError && sourceTarget == copiedTarget;
        }
        const bool exact = result.has_value() && SUCCEEDED(result.value().hr) && result.value().discoveredFiles == expectedFiles &&
                           result.value().discoveredTotalBytes == 8u * kClosureFileBytes && result.value().copiedFileCountKnown &&
                           result.value().copiedFileCount == expectedFiles;
        AppendLog(std::format(L"MultipleFolderLiveProgress: mode={} exact={} monotonic={} closedDuringCopy={} contents={} boundedRows={} bothRootsVisible={}",
                              static_cast<unsigned int>(mode),
                              exact,
                              monotonic,
                              closedDuringCopy,
                              contentMatches,
                              boundedRows,
                              bothRootsVisible));
        allRoutesCorrect = allRoutesCorrect && exact && monotonic && closedDuringCopy && contentMatches && boundedRows && bothRootsVisible;
    }
    // The product intentionally rejects Bulk Copy before I/O. This is Native API qualification on
    // the same real fixture, with its one public-call observer scope and a non-admitted reducer.
    const auto bulkDestination = root / L"out-bulk-api";
    if (! RecreateEmptyDirectory(bulkDestination))
        return reject(L"Native bulk destination unavailable");
    using Task = FolderWindow::FileOperationState::Task;
    Task bulkTask(*operations);
    bulkTask._executionMode = FolderWindow::FileOperationState::ExecutionMode::BulkItems;
    bulkTask._sourcePaths   = {root / L"one", root / L"two"};
    bulkTask.BeginDiscovery();
    struct BulkObserver final : IFileSystemFileWorkCallback
    {
        Task& task;
        std::wstring one;
        std::wstring two;
        std::atomic<bool> bothRoots{false};
        std::atomic<bool> valid{true};
        BulkObserver(Task& ownerTask, const std::filesystem::path& fixture)
            : task(ownerTask),
              one((fixture / L"one").wstring()),
              two((fixture / L"two").wstring())
        {
        }
        HRESULT STDMETHODCALLTYPE FileSystemReportFileWork(const FileSystemFileWorkProgress* progress, void* cookie) noexcept override
        {
            if (FAILED(task.FileSystemReportFileWork(progress, cookie)))
                valid.store(false, std::memory_order_release);
            std::scoped_lock lock(task._inFlightFilesMutex);
            bool firstActive  = false;
            bool secondActive = false;
            for (size_t index = 0u; index < task._inFlightFileCount; ++index)
            {
                const auto& file = task._inFlightFiles[index];
                if (! file.terminal && file.kind == FileSystemFileWorkKind::RegularFile)
                {
                    firstActive  = firstActive || file.sourcePath.starts_with(one);
                    secondActive = secondActive || file.sourcePath.starts_with(two);
                }
            }
            if (firstActive && secondActive)
                bothRoots.store(true, std::memory_order_release);
            return S_OK;
        }
    } observer(bulkTask, root);
    DirectApiDiscoveryRecorder discovery;
    FileOpsRecursiveProgressRecorder progress;
    FileSystemOptions options{};
    options.sizeBytes                    = sizeof(options);
    options.bandwidthLimitBytesPerSecond = kClosureSpeedLimit;
    options.copyMoveMaxConcurrency       = 4u;
    options.operationControl             = &discovery;
    options.fileWorkCallback             = &observer;
    const std::wstring firstRoot         = (root / L"one").wstring();
    const std::wstring secondRoot        = (root / L"two").wstring();
    const std::array<const wchar_t*, 2> sources{{firstRoot.c_str(), secondRoot.c_str()}};
    const HRESULT bulkStatus = state.fsLocal->CopyItems(
        sources.data(), static_cast<unsigned long>(sources.size()), bulkDestination.c_str(), FILESYSTEM_FLAG_RECURSIVE, &options, &progress, nullptr);
    bool bulkContent = true;
    for (unsigned long index = 0u; index < 4u; ++index)
        bulkContent = FilesEqualBytes(root / L"one" / std::format(L"one-{:02}.bin", index), bulkDestination / L"one" / std::format(L"one-{:02}.bin", index)) &&
                      FilesEqualBytes(root / L"two" / L"nested" / std::format(L"two-{:02}.bin", index),
                                      bulkDestination / L"two" / L"nested" / std::format(L"two-{:02}.bin", index)) &&
                      bulkContent;
    for (const auto& name : {L"one", L"two"})
    {
        if (! semanticLinksAvailable)
            break;
        const auto linkName     = std::wstring(name) + L"-link.bin";
        const auto sourceTarget = std::filesystem::read_symlink(root / name / linkName, linkError);
        const bool sourceValid  = ! linkError;
        const auto copiedTarget = std::filesystem::read_symlink(bulkDestination / name / linkName, linkError);
        bulkContent             = bulkContent && sourceValid && ! linkError && sourceTarget == copiedTarget;
    }
    const bool bulkExact =
        SUCCEEDED(bulkStatus) && observer.valid.load(std::memory_order_acquire) && observer.bothRoots.load(std::memory_order_acquire) &&
        bulkTask._copiedFileCountKnown.load(std::memory_order_acquire) && bulkTask._copiedFileCount.load(std::memory_order_acquire) == expectedFiles &&
        DescribeDiscoveryScopeDefect(
            L"Native CopyItems(two folders; explicit semantic-link prerequisite)", discovery, {8u * kClosureFileBytes, expectedFiles, std::nullopt})
            .empty();
    AppendLog(std::format(L"MultipleFolderLiveProgress: NativeBulkApi exact={} content={} copied={} bothRoots={}",
                          bulkExact,
                          bulkContent,
                          bulkTask._copiedFileCount.load(std::memory_order_acquire),
                          observer.bothRoots.load(std::memory_order_acquire)));
    allRoutesCorrect                = allRoutesCorrect && bulkExact && bulkContent;
    const bool quietAndRestored     = scope.Close();
    const bool configurationMatches = quietAndRestored && restoreConfiguration();
    return allRoutesCorrect && quietAndRestored && configurationMatches;
}

// R6.1/R6.5/R6.6/R6.8: flat folders and a tree close at listing end on every walker.
[[nodiscard]] bool RunDiscoveryScopeFolderClosesBeforeLastFileChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Discovery closure: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.fsDummy || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local/Dummy providers or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"discovery-closure";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const auto flat       = root / L"flat";
    const auto tree       = root / L"tree";
    if (! SeedClosureFiles(flat, kClosureFlatFiles) || ! SeedClosureFiles(tree / L"a", 2u) || ! SeedClosureFiles(tree / L"a" / L"deep", 2u) ||
        ! SeedClosureFiles(tree / L"b", 2u) || ! SeedClosureFiles(tree, 2u, L"top"))
        return reject(L"fixtures unavailable");
    constexpr uint64_t kFlatBytes = kClosureFlatFiles * kClosureFileBytes;
    constexpr uint64_t kTreeBytes = 8u * kClosureFileBytes;
    ClosureCaseScope scope(state, operations);
    std::vector<std::wstring> defects;
    const auto record = [&](std::wstring defect)
    {
        if (! defect.empty())
            defects.push_back(std::move(defect));
    };

    // Local parallel walker (4 workers) and Local serial walker (1).
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4))
        return reject(L"Local concurrency 4 could not be configured");
    if (! RecreateEmptyDirectory(root / L"out-local-parallel") || ! RecreateEmptyDirectory(root / L"out-local-tree"))
        return reject(L"Local destinations unavailable");
    record(DescribeLateClosure(
        L"Local parallel flat", scope.Run(FILESYSTEM_COPY, state.fsLocal, {flat}, root / L"out-local-parallel"), kFlatBytes, kClosureFlatFiles));
    record(DescribeLateClosure(L"Local parallel tree", scope.Run(FILESYSTEM_COPY, state.fsLocal, {tree}, root / L"out-local-tree"), kTreeBytes, 8u));
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency1))
        return reject(L"Local concurrency 1 could not be configured");
    if (! RecreateEmptyDirectory(root / L"out-local-serial") || ! RecreateEmptyDirectory(root / L"out-local-serial-tree"))
        return reject(L"Local serial destinations unavailable");
    record(DescribeLateClosure(
        L"Local serial flat", scope.Run(FILESYSTEM_COPY, state.fsLocal, {flat}, root / L"out-local-serial"), kFlatBytes, kClosureFlatFiles));
    record(DescribeLateClosure(L"Local serial tree", scope.Run(FILESYSTEM_COPY, state.fsLocal, {tree}, root / L"out-local-serial-tree"), kTreeBytes, 8u));

    // Host bridge to the Dummy provider: the parallel producer, then the sequential walker that
    // Verify On selects.
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4))
        return reject(L"Local concurrency 4 could not be restored");
    const std::wstring dummyRoot = L"/discovery-closure";
    if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot + L"/parallel") ||
        ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot + L"/sequential") ||
        ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot + L"/sequential-tree"))
        return reject(L"Dummy destinations unavailable");
    record(DescribeLateClosure(L"bridge parallel flat",
                               scope.Run(FILESYSTEM_COPY, state.fsDummy, {flat}, std::filesystem::path(dummyRoot + L"/parallel")),
                               kFlatBytes,
                               kClosureFlatFiles));
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = true;
    record(DescribeLateClosure(L"bridge sequential flat (Verify on)",
                               scope.Run(FILESYSTEM_COPY, state.fsDummy, {flat}, std::filesystem::path(dummyRoot + L"/sequential")),
                               kFlatBytes,
                               kClosureFlatFiles));
    record(DescribeLateClosure(L"bridge sequential tree (Verify on)",
                               scope.Run(FILESYSTEM_COPY, state.fsDummy, {tree}, std::filesystem::path(dummyRoot + L"/sequential-tree")),
                               kTreeBytes,
                               8u));
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;

    if (! defects.empty())
    {
        for (const std::wstring& defect : defects)
            AppendLog(std::format(L"Discovery closure: FAILED: {}", defect));
        return false;
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.DiscoveryClosure.BeforeLastFile", L"local-parallel,local-serial,bridge-parallel,bridge-sequential", 0u, 7u, 0u, S_OK);
    return true;
}

// R6.1 (D-6.1): a folder wider than the queue lead closes once its last entry is listed, while the
// throttled copy still has well over a queue lead of files to go.
[[nodiscard]] bool RunDiscoveryScopeWideFolderClosesWithinQueueLeadChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Discovery closure wide folder: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    constexpr unsigned long kWideFiles = 600u;
    constexpr size_t kWideFileBytes    = 16u * 1024u;
    const std::filesystem::path root   = state.tempRoot / L"discovery-closure-wide";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(root / L"out"))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    for (unsigned long index = 0u; index < kWideFiles; ++index)
    {
        if (! WriteFilledTestFile(root / L"wide" / std::format(L"entry-{:04}.bin", index), kWideFileBytes, static_cast<unsigned char>(index & 0xFFu)))
            return reject(L"wide fixture unavailable");
    }
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4))
        return reject(L"Local concurrency 4 could not be configured");
    ClosureCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    const auto info                                           = scope.Run(FILESYSTEM_COPY, state.fsLocal, {root / L"wide"}, root / L"out");
    if (! info.has_value() || FAILED(info->hr) || ! info->discoveryClosed || info->discoveredFiles != kWideFiles ||
        info->discoveredTotalBytes != static_cast<uint64_t>(kWideFiles) * kWideFileBytes || info->discoveryGrowthAfterCloseCount != 0u)
        return reject(L"the wide copy failed or closed on wrong totals");
    AppendLog(std::format(L"Discovery closure wide folder: files={} mutationsWhileOpen={} bytesWhileOpen={}/{} queueMax={}",
                          info->discoveredFiles,
                          info->discoveryMutationsCompletedWhileOpen,
                          info->discoveryBytesCompletedWhileOpen,
                          info->discoveredTotalBytes,
                          info->discoveryMaxQueueDepth));
    // Closure follows the last listed entry, bounded by the queue lead (128-256 entries ahead): at
    // least a quarter of a lead of files must still have been waiting when it closed.
    if (info->discoveryMutationsCompletedWhileOpen + 32u > kWideFiles || info->discoveryBytesCompletedWhileOpen >= info->discoveredTotalBytes)
        return reject(L"the wide folder closed only near its last file");
    return true;
}

// R6.2: a created directory's timestamps are restored after its last child completed, deepest
// first, on the Local parallel walker and on the host bridge (a Dummy tree copied to Local: only a
// destination with object binding restores created-directory metadata; a destination without it
// records one I35 R5.4 note instead).
[[nodiscard]] bool RunDiscoveryScopeNestedTreeMetadataChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Discovery closure metadata: FAILED: {}", reason));
        return false;
    };
    auto* folder                        = TryGetFolderWindow(state.mainWindow);
    auto* operations                    = TryGetFileOps(folder);
    wil::com_ptr<IFileSystemIO> localIo = state.fsLocal ? state.fsLocal.query<IFileSystemIO>() : nullptr;
    wil::com_ptr<IFileSystemIO> dummyIo = state.fsDummy ? state.fsDummy.query<IFileSystemIO>() : nullptr;
    if (! folder || ! operations || ! localIo || ! dummyIo || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local/Dummy IO or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"discovery-closure-metadata";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(root / L"out"))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const auto source     = root / L"meta";
    const std::array<std::filesystem::path, 4> directories{{source, source / L"a", source / L"a" / L"deep", source / L"b"}};
    if (! SeedClosureFiles(source / L"a" / L"deep", 2u) || ! SeedClosureFiles(source / L"a", 1u) || ! SeedClosureFiles(source / L"b", 2u) ||
        ! SeedClosureFiles(source, 1u, L"top"))
        return reject(L"tree fixture unavailable");
    // Distinct old timestamps per directory, set deepest first so no later write moves a parent's.
    constexpr int64_t kDay = 24ll * 60ll * 60ll * 10'000'000ll;
    std::array<int64_t, 4> expectedLastWrite{};
    for (size_t reverse = directories.size(); reverse > 0u; --reverse)
    {
        const size_t index = reverse - 1u;
        FileSystemBasicInformation basic{};
        basic.sizeBytes = sizeof(basic);
        if (FAILED(localIo->GetFileBasicInformation(directories[index].c_str(), &basic)))
            return reject(L"source directory metadata unreadable");
        basic.lastWriteTime -= static_cast<int64_t>(index + 3u) * kDay;
        basic.creationTime -= static_cast<int64_t>(index + 3u) * kDay;
        if (FAILED(localIo->SetFileBasicInformation(directories[index].c_str(), &basic)))
            return reject(L"source directory metadata could not be seeded");
        expectedLastWrite[index] = basic.lastWriteTime;
    }
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4))
        return reject(L"Local concurrency 4 could not be configured");
    ClosureCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;

    const auto checkTree = [&](std::wstring_view name, IFileSystemIO* io, const std::wstring& destinationRoot) -> std::wstring
    {
        const std::array<std::wstring, 4> relative{{L"", L"\\a", L"\\a\\deep", L"\\b"}};
        for (size_t index = 0u; index < relative.size(); ++index)
        {
            std::wstring path = destinationRoot + L"\\meta" + relative[index];
            if (! destinationRoot.empty() && destinationRoot.front() == L'/')
                std::ranges::replace(path, L'\\', L'/');
            FileSystemBasicInformation basic{};
            basic.sizeBytes = sizeof(basic);
            if (FAILED(io->GetFileBasicInformation(path.c_str(), &basic)))
                return std::format(L"{}: {} unreadable", name, path);
            if (basic.lastWriteTime != expectedLastWrite[index])
                return std::format(L"{}: {} lastWrite={} expected={}", name, path, basic.lastWriteTime, expectedLastWrite[index]);
        }
        return {};
    };
    const auto local = scope.Run(FILESYSTEM_COPY, state.fsLocal, {source}, root / L"out", 0u);
    if (! local.has_value() || FAILED(local->hr))
        return reject(L"the Local tree copy failed");
    if (const std::wstring defect = checkTree(L"Local parallel", localIo.get(), (root / L"out").native()); ! defect.empty())
        return reject(defect);

    // The same tree in the Dummy provider, its directory times set deepest first after its files.
    const std::wstring dummyRoot                       = L"/discovery-closure-metadata";
    const std::array<std::wstring, 4> dummyDirectories = {{dummyRoot + L"/meta", dummyRoot + L"/meta/a", dummyRoot + L"/meta/a/deep", dummyRoot + L"/meta/b"}};
    const std::array<std::wstring_view, 4> dummyFileLeaf = {{L"top.bin", L"a.bin", L"deep.bin", L"b.bin"}};
    if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
        return reject(L"Dummy source root unavailable");
    for (size_t index = 0u; index < dummyDirectories.size(); ++index)
    {
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyDirectories[index]) ||
            ! DummyWriteTextFile(state.fsDummy.get(), dummyDirectories[index] + L"/" + std::wstring(dummyFileLeaf[index]), "metadata child", true))
            return reject(L"Dummy source tree unavailable");
    }
    std::array<int64_t, 4> expectedBridgeLastWrite{};
    for (size_t reverse = dummyDirectories.size(); reverse > 0u; --reverse)
    {
        const size_t index = reverse - 1u;
        FileSystemBasicInformation basic{};
        basic.sizeBytes = sizeof(basic);
        if (FAILED(dummyIo->GetFileBasicInformation(dummyDirectories[index].c_str(), &basic)))
            return reject(L"Dummy source directory metadata unreadable");
        basic.lastWriteTime -= static_cast<int64_t>(index + 7u) * kDay;
        basic.creationTime -= static_cast<int64_t>(index + 7u) * kDay;
        if (FAILED(dummyIo->SetFileBasicInformation(dummyDirectories[index].c_str(), &basic)))
            return reject(L"Dummy source directory metadata could not be seeded");
        expectedBridgeLastWrite[index] = basic.lastWriteTime;
    }
    if (! RecreateEmptyDirectory(root / L"out-bridge"))
        return reject(L"bridge destination unavailable");
    const auto bridge =
        scope.Run(FILESYSTEM_COPY, state.fsLocal, {std::filesystem::path(dummyDirectories[0])}, root / L"out-bridge", 0u, {}, {}, state.fsDummy);
    if (! bridge.has_value() || FAILED(bridge->hr))
        return reject(L"the bridge tree copy failed");
    expectedLastWrite = expectedBridgeLastWrite;
    if (const std::wstring defect = checkTree(L"bridge (Dummy to Local)", localIo.get(), (root / L"out-bridge").native()); ! defect.empty())
        return reject(defect);
    AppendLog(L"Discovery closure metadata: every created directory kept its source timestamps on the Local parallel walker and the bridge.");
    return true;
}

// R6.6 invariant: a Managed Move (bridge walker with exact source cleanup) lists a directory
// completely before it removes any of its children, so no child is skipped whatever the name order.
// The Move crosses roots so that it is Managed (inside one root it is one native rename that lists
// nothing): from the loopback administrative share (SMB enumeration) when it is reachable, and from
// the authorized alternate fixed volume when there is one, to the case root. Both ends hold exact
// authority there, so the source cleanup really runs. A source without file identities, such as the
// NAS of I35 section 2, keeps its sources by design (D-7.3 never moves by name) and cannot witness
// the removal order. With neither root available the case is skipped, never passed.
[[nodiscard]] bool RunDiscoveryScopeManagedMoveListingChecks(SelfTestState& state, std::wstring& unavailableReason) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Discovery closure Managed Move: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"discovery-closure-move";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    // Names whose removal during an open listing would shift later entries: upper/lower case and a
    // name that sorts between them.
    const std::array<std::wstring_view, 6> names{{L"A.bin", L"a-1.bin", L"B.bin", L"b-1.bin", L"_first.bin", L"~last.bin"}};
    ClosureCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;

    // Seeds `album` under `seedParent`, moves it as `moveParent / album` (the same folder, possibly by
    // another path) to `destination`, and returns the defect, if any.
    const auto witness = [&](std::wstring_view where,
                             const std::filesystem::path& seedParent,
                             const std::filesystem::path& moveParent,
                             const std::filesystem::path& destination) -> std::wstring
    {
        const auto seeded = seedParent / L"album";
        for (size_t index = 0u; index < names.size(); ++index)
        {
            if (! WriteFilledTestFile(seeded / names[index], 64u * 1024u + index, static_cast<unsigned char>(0x30u + index)) ||
                ! WriteFilledTestFile(seeded / L"sub" / names[index], 32u * 1024u + index, static_cast<unsigned char>(0x60u + index)))
                return std::format(L"{}: Move fixture unavailable", where);
        }
        if (! RecreateEmptyDirectory(destination))
            return std::format(L"{}: destination unavailable", where);
        const auto moved = scope.Run(FILESYSTEM_MOVE, state.fsLocal, {moveParent / L"album"}, destination, 0u);
        if (! moved.has_value() || moved->hr != S_OK || ! moved->discoveryClosed || moved->discoveredFiles != 2u * names.size())
        {
            // Name the task's own diagnostics so a partial outcome says which child or step it was.
            std::wstring details;
            std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> diagnostics;
            operations->CollectDiagnostics(diagnostics);
            for (const auto& entry : diagnostics)
            {
                if (entry.taskId == scope.LastTaskId())
                    details +=
                        std::format(L" [{} 0x{:08X} {} '{}']", entry.category, static_cast<unsigned long>(entry.status), entry.message, entry.sourcePath);
            }
            if (moved.has_value())
            {
                for (const auto& status : moved->sourceItemStatuses)
                    details += status.has_value() ? std::format(L" item=0x{:08X}", static_cast<unsigned long>(status.value())) : std::wstring(L" item=-");
            }
            return std::format(L"{}: the Managed Move did not complete cleanly (hr=0x{:08X} closed={} files={} completedFiles={} completedFolders={}){}",
                               where,
                               moved.has_value() ? static_cast<unsigned long>(moved->hr) : 0ul,
                               moved.has_value() && moved->discoveryClosed,
                               moved.has_value() ? moved->discoveredFiles : 0ul,
                               moved.has_value() ? moved->completedFiles : 0ul,
                               moved.has_value() ? moved->completedFolders : 0ul,
                               details);
        }
        std::error_code ec;
        if (std::filesystem::exists(seeded, ec) || ec)
            return std::format(L"{}: the moved source folder still exists", where);
        for (size_t index = 0u; index < names.size(); ++index)
        {
            if (! FileSizeEquals(destination / L"album" / names[index], 64u * 1024u + index) ||
                ! FileSizeEquals(destination / L"album" / L"sub" / names[index], 32u * 1024u + index))
                return std::format(L"{}: {} was not moved", where, names[index]);
        }
        return {};
    };

    std::vector<std::wstring> witnessed;
    std::vector<std::wstring> unavailable;
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");
    if (uncStatus == LoopbackShareStatus::Unreachable)
    {
        unavailable.push_back(L"loopback-share-unreachable");
    }
    else
    {
        if (const std::wstring defect = witness(L"loopback share", root / L"unc", LoopbackShareAlias(root / L"unc"), root / L"moved-unc"); ! defect.empty())
            return reject(defect);
        witnessed.push_back(L"loopback-share");
    }
    std::wstring alternateDetail;
    const std::optional<std::filesystem::path> alternate = TryCreateAlternateWritableVolumeSelfTestRoot(state.tempRoot, alternateDetail);
    if (! alternate.has_value())
    {
        unavailable.push_back(alternateDetail.empty() ? std::wstring(L"alternate-volume-unavailable") : alternateDetail);
    }
    else
    {
        const auto alternateParent = alternate.value() / L"discovery-closure-move";
        const auto removeAlternate = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(alternateParent)); });
        if (! RecreateEmptyDirectory(alternateParent))
            return reject(L"alternate volume: case folder unavailable");
        if (const std::wstring defect = witness(L"alternate volume", alternateParent, alternateParent, root / L"moved-alt"); ! defect.empty())
            return reject(defect);
        witnessed.push_back(L"alternate-volume");
    }
    if (witnessed.empty())
    {
        for (const std::wstring& reason : unavailable)
            unavailableReason += (unavailableReason.empty() ? L"" : L";") + reason;
        AppendLog(std::format(L"Discovery closure Managed Move: no second root ({}); witness not run.", unavailableReason));
        return false;
    }
    std::wstring where;
    for (const std::wstring& location : witnessed)
        where += (where.empty() ? L"" : L",") + location;
    AppendLog(std::format(L"Discovery closure Managed Move: every child moved and the source folder was removed ({}).", where));
    Debug::Perf::Emit(L"FileOps.SelfTest.DiscoveryClosure.ManagedMoveWitness", where, 0u, static_cast<uint64_t>(witnessed.size()), 0u, S_OK);
    return true;
}

// R6.3: a directory whose listing fails after entries were counted, and is retried, counts every entry
// exactly once. The root lists 300 files, a subfolder and 10 more files; the throttled copy keeps the
// queue full, so the subfolder is held by name when the listing fails after 305 entries. The retry
// replays the counted entries without counting them again and still counts the held subfolder's
// subtree, which the failed attempt never walked.
[[nodiscard]] bool RunDiscoveryScopeRetryDoesNotGrowAfterCloseChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Discovery closure retry: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    constexpr unsigned long kLeadFiles = 300u;
    constexpr unsigned long kSubFiles  = 5u;
    constexpr unsigned long kTailFiles = 10u;
    constexpr uint64_t kFileBytes      = 4u * 1024u;
    constexpr unsigned long kAllFiles  = kLeadFiles + kSubFiles + kTailFiles;
    const std::filesystem::path root   = state.tempRoot / L"discovery-closure-retry";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(root / L"out"))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const auto source     = root / L"retry";
    std::vector<std::filesystem::path> expected;
    expected.reserve(kAllFiles);
    for (unsigned long index = 0u; index < kLeadFiles; ++index)
        expected.push_back(std::filesystem::path(std::format(L"a-{:03}.bin", index)));
    for (unsigned long index = 0u; index < kSubFiles; ++index)
        expected.push_back(std::filesystem::path(L"m-sub") / std::format(L"s-{}.bin", index));
    for (unsigned long index = 0u; index < kTailFiles; ++index)
        expected.push_back(std::filesystem::path(std::format(L"n-{:03}.bin", index)));
    for (size_t index = 0u; index < expected.size(); ++index)
    {
        if (! WriteFilledTestFile(source / expected[index], static_cast<size_t>(kFileBytes), static_cast<unsigned char>(index & 0xFFu)))
            return reject(L"fixture unavailable");
    }
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4))
        return reject(L"Local concurrency 4 could not be configured");

    constexpr wchar_t kFailDirectory[] = L"REDSALAMANDER_FILEOPS_LIST_FAIL_DIRECTORY";
    constexpr wchar_t kFailAfter[]     = L"REDSALAMANDER_FILEOPS_LIST_FAIL_AFTER";
    const auto disarm                  = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kFailDirectory, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kFailAfter, nullptr));
    });
    // 300 files, the subfolder and 4 tail files are handled when the listing fails.
    if (SetEnvironmentVariableW(kFailAfter, L"305") == FALSE || SetEnvironmentVariableW(kFailDirectory, source.c_str()) == FALSE)
        return reject(L"listing-failure injection could not be armed");

    ClosureCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    using Task                                                = FolderWindow::FileOperationState::Task;
    uint64_t answeredGeneration                               = 0u;
    unsigned long retryAnswers                                = 0u;
    // The listing failure gets Retry; the retry queues again the files the failed attempt already
    // queued, and their second copies are skipped once for all (the first copies completed).
    const auto answer = [&](Task& task) noexcept -> bool
    {
        const auto prompt = TryGetConflictPromptCopy(&task);
        if (! prompt.has_value())
            return true;
        if (prompt->generation == answeredGeneration)
            return true; // the answer is on its way
        if (prompt->bucket == Task::ConflictBucket::SharingViolation && PromptHasAction(prompt.value(), Task::ConflictAction::Retry))
        {
            answeredGeneration = prompt->generation;
            ++retryAnswers;
            task.DebugSubmitCurrentConflictDecision(Task::ConflictAction::Retry, false);
            return true;
        }
        if (PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
        {
            answeredGeneration = prompt->generation;
            task.DebugSubmitCurrentConflictDecision(Task::ConflictAction::Skip, true);
            return true;
        }
        return false;
    };
    const auto info = scope.Run(FILESYSTEM_COPY, state.fsLocal, {source}, root / L"out", 256u * 1024u, {}, {}, {}, answer);
    if (! info.has_value())
        return reject(L"the copy did not complete (or asked an unexpected question)");
    // The injection clears its directory when it fires.
    if (GetEnvironmentVariableW(kFailDirectory, nullptr, 0u) != 0u)
        return reject(L"the injected listing failure never fired");
    // The walk counts the selected folder itself and its one subfolder.
    if (! info->discoveryClosed || info->discoveredFiles != kAllFiles || info->discoveredDirectories != 2u ||
        info->discoveredTotalBytes != kAllFiles * kFileBytes || info->discoveryGrowthAfterCloseCount != 0u)
        return reject(std::format(L"discovery counted closed={} files={}/{} dirs={}/2 bytes={}/{} growthAfterClose={}",
                                  info->discoveryClosed,
                                  info->discoveredFiles,
                                  kAllFiles,
                                  info->discoveredDirectories,
                                  info->discoveredTotalBytes,
                                  kAllFiles * kFileBytes,
                                  info->discoveryGrowthAfterCloseCount));
    for (const std::filesystem::path& relative : expected)
    {
        if (! FileSizeEquals(root / L"out" / L"retry" / relative, kFileBytes))
            return reject(std::format(L"{} was not copied", relative.native()));
    }
    AppendLog(std::format(L"Discovery closure retry: hr=0x{:08X} files={} dirs={} growthAfterClose=0 after a retried listing (Retry answers: {}).",
                          static_cast<unsigned long>(info->hr),
                          info->discoveredFiles,
                          info->discoveredDirectories,
                          retryAnswers));

    // Part B: a subfolder walked inline whose listing fails and is skipped stays skipped when its
    // parent's listing fails later and is retried: the retry does not walk it again, so nothing is
    // copied that discovery did not count. The subfolder fails after its first file, the parent after
    // the subfolder and 50 files; the first question is answered Skip and the second Retry.
    const auto sourceB                  = root / L"retry-skip";
    constexpr unsigned long kPartBFiles = 100u;
    for (unsigned long index = 0u; index < 3u; ++index)
    {
        if (! WriteFilledTestFile(
                sourceB / L"a-sub" / std::format(L"s-{}.bin", index), static_cast<size_t>(kFileBytes), static_cast<unsigned char>(0x30u + index)))
            return reject(L"part B fixture unavailable");
    }
    for (unsigned long index = 0u; index < kPartBFiles; ++index)
    {
        if (! WriteFilledTestFile(sourceB / std::format(L"f-{:03}.bin", index), static_cast<size_t>(kFileBytes), static_cast<unsigned char>(index & 0xFFu)))
            return reject(L"part B fixture unavailable");
    }
    if (! RecreateEmptyDirectory(root / L"out-b"))
        return reject(L"part B destination unavailable");
    const std::wstring failDirectories = (sourceB / L"a-sub").native() + L"|" + sourceB.native();
    if (SetEnvironmentVariableW(kFailAfter, L"1|51") == FALSE || SetEnvironmentVariableW(kFailDirectory, failDirectories.c_str()) == FALSE)
        return reject(L"part B listing-failure injection could not be armed");
    unsigned long sharingAnswers = 0u;
    answeredGeneration           = 0u;
    const auto answerB           = [&](Task& task) noexcept -> bool
    {
        const auto prompt = TryGetConflictPromptCopy(&task);
        if (! prompt.has_value())
            return true;
        if (prompt->generation == answeredGeneration)
            return true; // the answer is on its way
        if (prompt->bucket == Task::ConflictBucket::SharingViolation && PromptHasAction(prompt.value(), Task::ConflictAction::Retry) &&
            PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
        {
            answeredGeneration = prompt->generation;
            task.DebugSubmitCurrentConflictDecision(sharingAnswers++ == 0u ? Task::ConflictAction::Skip : Task::ConflictAction::Retry, false);
            return true;
        }
        if (PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
        {
            answeredGeneration = prompt->generation;
            task.DebugSubmitCurrentConflictDecision(Task::ConflictAction::Skip, true);
            return true;
        }
        return false;
    };
    const auto infoB = scope.Run(FILESYSTEM_COPY, state.fsLocal, {sourceB}, root / L"out-b", 256u * 1024u, {}, {}, {}, answerB);
    if (! infoB.has_value())
        return reject(L"part B: the copy did not complete (or asked an unexpected question)");
    if (GetEnvironmentVariableW(kFailDirectory, nullptr, 0u) != 0u || sharingAnswers != 2u)
        return reject(std::format(L"part B: the two listing failures did not both ask (answers {})", sharingAnswers));
    // The skipped subfolder counted (and copied) its first file only; the root and the subfolder are the
    // two directories.
    if (! infoB->discoveryClosed || infoB->discoveredFiles != kPartBFiles + 1u || infoB->discoveredDirectories != 2u ||
        infoB->discoveryGrowthAfterCloseCount != 0u)
        return reject(std::format(L"part B: discovery counted closed={} files={}/{} dirs={}/2 growthAfterClose={}",
                                  infoB->discoveryClosed,
                                  infoB->discoveredFiles,
                                  kPartBFiles + 1u,
                                  infoB->discoveredDirectories,
                                  infoB->discoveryGrowthAfterCloseCount));
    std::error_code ec;
    if (std::filesystem::exists(root / L"out-b" / L"retry-skip" / L"a-sub" / L"s-1.bin", ec) ||
        std::filesystem::exists(root / L"out-b" / L"retry-skip" / L"a-sub" / L"s-2.bin", ec))
        return reject(L"part B: the retry walked the skipped subfolder again");
    AppendLog(std::format(
        L"Discovery closure retry: part B kept the skipped subfolder skipped (files={} dirs={}).", infoB->discoveredFiles, infoB->discoveredDirectories));
    return true;
}

// R6.1: a subfolder held by name whose own listing fails, and is skipped, is no longer pending, so the
// scope closes then, not after the files held behind it are queued. The root lists 130 files (the queue
// lead fills), the subfolder, which fails after its first file, and 300 more files.
[[nodiscard]] bool RunDiscoveryScopeHeldFolderFailureStillClosesChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Discovery closure held failure: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    constexpr unsigned long kLeadFiles = 130u;
    constexpr unsigned long kTailFiles = 300u;
    constexpr uint64_t kFileBytes      = 16u * 1024u;
    constexpr unsigned long kAllFiles  = kLeadFiles + 1u + kTailFiles; // the subfolder counted its first file
    const std::filesystem::path root   = state.tempRoot / L"discovery-closure-held-failure";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(root / L"out"))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const auto source     = root / L"held";
    for (unsigned long index = 0u; index < kLeadFiles; ++index)
    {
        if (! WriteFilledTestFile(source / std::format(L"a-{:03}.bin", index), static_cast<size_t>(kFileBytes), static_cast<unsigned char>(index & 0xFFu)))
            return reject(L"fixture unavailable");
    }
    for (unsigned long index = 0u; index < 3u; ++index)
    {
        if (! WriteFilledTestFile(
                source / L"m-sub" / std::format(L"s-{}.bin", index), static_cast<size_t>(kFileBytes), static_cast<unsigned char>(0x30u + index)))
            return reject(L"fixture unavailable");
    }
    for (unsigned long index = 0u; index < kTailFiles; ++index)
    {
        if (! WriteFilledTestFile(source / std::format(L"n-{:03}.bin", index), static_cast<size_t>(kFileBytes), static_cast<unsigned char>(index & 0xFFu)))
            return reject(L"fixture unavailable");
    }
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4))
        return reject(L"Local concurrency 4 could not be configured");

    constexpr wchar_t kFailDirectory[] = L"REDSALAMANDER_FILEOPS_LIST_FAIL_DIRECTORY";
    constexpr wchar_t kFailAfter[]     = L"REDSALAMANDER_FILEOPS_LIST_FAIL_AFTER";
    const auto disarm                  = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kFailDirectory, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kFailAfter, nullptr));
    });
    if (SetEnvironmentVariableW(kFailAfter, L"1") == FALSE || SetEnvironmentVariableW(kFailDirectory, (source / L"m-sub").c_str()) == FALSE)
        return reject(L"listing-failure injection could not be armed");

    ClosureCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    using Task                                                = FolderWindow::FileOperationState::Task;
    uint64_t answeredGeneration                               = 0u;
    unsigned long skipAnswers                                 = 0u;
    const auto answer                                         = [&](Task& task) noexcept -> bool
    {
        const auto prompt = TryGetConflictPromptCopy(&task);
        if (! prompt.has_value())
            return true;
        if (prompt->generation == answeredGeneration)
            return true; // the answer is on its way
        if (prompt->bucket == Task::ConflictBucket::SharingViolation && PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
        {
            answeredGeneration = prompt->generation;
            ++skipAnswers;
            task.DebugSubmitCurrentConflictDecision(Task::ConflictAction::Skip, false);
            return true;
        }
        return false;
    };
    const auto info = scope.Run(FILESYSTEM_COPY, state.fsLocal, {source}, root / L"out", 1024u * 1024u, {}, {}, {}, answer);
    if (! info.has_value())
        return reject(L"the copy did not complete (or asked an unexpected question)");
    if (GetEnvironmentVariableW(kFailDirectory, nullptr, 0u) != 0u || skipAnswers != 1u)
        return reject(std::format(L"the subfolder's listing failure did not ask once (answers {})", skipAnswers));
    constexpr uint64_t kTotalBytes = kAllFiles * kFileBytes;
    if (! info->discoveryClosed || info->discoveredFiles != kAllFiles || info->discoveredDirectories != 2u || info->discoveredTotalBytes != kTotalBytes ||
        info->discoveryGrowthAfterCloseCount != 0u)
        return reject(std::format(L"discovery counted closed={} files={}/{} dirs={}/2 bytes={}/{} growthAfterClose={}",
                                  info->discoveryClosed,
                                  info->discoveredFiles,
                                  kAllFiles,
                                  info->discoveredDirectories,
                                  info->discoveredTotalBytes,
                                  kTotalBytes,
                                  info->discoveryGrowthAfterCloseCount));
    // Closed once the subfolder was skipped: only the few files queued before it were done by then, never
    // most of the 300 held behind it.
    if (info->discoveryBytesCompletedWhileOpen >= kTotalBytes / 4u)
        return reject(std::format(L"the scope closed only after {} of {} bytes were copied", info->discoveryBytesCompletedWhileOpen, kTotalBytes));
    AppendLog(std::format(L"Discovery closure held failure: closed after {} of {} bytes, files={} dirs={}.",
                          info->discoveryBytesCompletedWhileOpen,
                          kTotalBytes,
                          info->discoveredFiles,
                          info->discoveredDirectories));
    return true;
}

// R6.7: a worker's failure that stops the walk stays that failure, never a cancellation. The failing
// file is early in a folder wider than the queue lead, so the producer is still queueing work when
// the worker fails (the host returns ERROR_NOT_ENOUGH_MEMORY to the worker as a task stop).
[[nodiscard]] bool RunDiscoveryScopeWorkerFailureIsNotCanceledChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Discovery closure worker failure: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    constexpr unsigned long kFiles   = 600u;
    const std::filesystem::path root = state.tempRoot / L"discovery-closure-worker-failure";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(root / L"out"))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    for (unsigned long index = 0u; index < kFiles; ++index)
    {
        if (! WriteFilledTestFile(root / L"wide" / std::format(L"entry-{:04}.bin", index), 4u * 1024u, static_cast<unsigned char>(index & 0xFFu)))
            return reject(L"wide fixture unavailable");
    }
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4))
        return reject(L"Local concurrency 4 could not be configured");
    constexpr wchar_t kFailSource[] = L"REDSALAMANDER_FILEOPS_COPY_FAIL_SOURCE";
    constexpr wchar_t kFailStatus[] = L"REDSALAMANDER_FILEOPS_COPY_FAIL_STATUS";
    const std::wstring failing      = (root / L"wide" / L"entry-0003.bin").native();
    const auto clearInjection       = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kFailSource, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kFailStatus, nullptr));
    });
    if (! SetEnvironmentVariableW(kFailStatus, L"0x80070008") || ! SetEnvironmentVariableW(kFailSource, failing.c_str()))
        return reject(L"the copy failure injection could not be armed");

    ClosureCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    const auto info                                           = scope.Run(FILESYSTEM_COPY, state.fsLocal, {root / L"wide"}, root / L"out", 512u * 1024u);
    if (! info.has_value())
        return reject(L"the copy did not complete (or asked an unexpected question)");
    const bool fired = GetEnvironmentVariableW(kFailSource, nullptr, 0u) == 0u;
    AppendLog(std::format(L"Discovery closure worker failure: fired={} hr=0x{:08X} closed={} growthAfterClose={}",
                          fired,
                          static_cast<unsigned long>(info->hr),
                          info->discoveryClosed,
                          info->discoveryGrowthAfterCloseCount));
    if (! fired)
        return reject(L"the injected copy failure never fired");
    if (info->hr != HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY))
        return reject(std::format(L"the walk ended 0x{:08X}, not with the worker's failure 0x80070008", static_cast<unsigned long>(info->hr)));
    if (info->discoveryGrowthAfterCloseCount != 0u)
        return reject(L"discovery grew after it closed");
    return true;
}

// DISC-3 (R6.5): the discovery-ahead reservation ends with a flat folder's listing, so the copy runs
// at the configured concurrency instead of half of it for its whole duration.
[[nodiscard]] bool RunDiscoveryScopeFlatFolderUsesFullConcurrencyChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Discovery closure concurrency: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"discovery-closure-concurrency";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(root / L"out"))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    // Each file spans several 1 MiB pump chunks, so a file that is copying reports partial progress; a
    // one-chunk file's only report already says it is complete.
    constexpr uint64_t kConcurrencyFileBytes = 4u * kClosureFileBytes;
    for (unsigned long index = 0u; index < kClosureFlatFiles; ++index)
    {
        if (! WriteFilledTestFile(
                root / L"flat" / std::format(L"file-{:02}.bin", index), static_cast<size_t>(kConcurrencyFileBytes), static_cast<unsigned char>(0x41u + index)))
            return reject(L"flat fixture unavailable");
    }
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4))
        return reject(L"Local concurrency 4 could not be configured");
    ClosureCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    // Files actively copying: started and not complete. A finished file's entry lingers briefly with its
    // total reached and does not count; a stalled one expires from the table.
    size_t maxInFlight = 0u;
    const auto info    = scope.Run(FILESYSTEM_COPY,
                                   state.fsLocal,
                                   {root / L"flat"},
                                   root / L"out",
                                   8u * 1024u * 1024u,
                                   [&](FolderWindow::FileOperationState::Task& task)
    {
        std::scoped_lock lock(task._inFlightFilesMutex);
        size_t active = 0u;
        for (size_t index = 0u; index < task._inFlightFileCount && index < task._inFlightFiles.size(); ++index)
        {
            const auto& file = task._inFlightFiles[index];
            if (file.totalBytes != 0u && file.completedBytes < file.totalBytes)
                ++active;
        }
        maxInFlight = (std::max)(maxInFlight, active);
    });
    if (! info.has_value() || FAILED(info->hr))
        return reject(L"the flat copy failed");
    AppendLog(std::format(L"Discovery closure concurrency: maxInFlight={} configured=4 bytesWhileOpen={}/{}",
                          maxInFlight,
                          info->discoveryBytesCompletedWhileOpen,
                          info->discoveredTotalBytes));
    if (maxInFlight < 4u)
        return reject(std::format(L"the copy ran at most {} files at once with concurrency 4", maxInFlight));
    return true;
}

// R6.10: a root still open while its copy completes work, with no discovery report for two seconds,
// records one Information note; the note never closes discovery. One worker copies a throttled file
// while the producer waits for the queue lead, so no new item is found for several seconds.
[[nodiscard]] bool RunDiscoveryScopeListingSilenceNotedChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Discovery closure silence: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    constexpr unsigned long kDirectories = 300u;
    const std::filesystem::path root     = state.tempRoot / L"discovery-closure-silence";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(root / L"out"))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    for (unsigned long index = 0u; index < kDirectories; ++index)
    {
        if (! WriteFilledTestFile(root / L"tree" / std::format(L"dir-{:03}", index) / L"file.bin", 128u * 1024u, static_cast<unsigned char>(index & 0xFFu)))
            return reject(L"tree fixture unavailable");
    }
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency1))
        return reject(L"Local concurrency 1 could not be configured");
    const auto restoreConcurrency = wil::scope_exit([&]() noexcept { static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4)); });
    ClosureCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    const ULONGLONG started                                   = GetTickCount64();
    bool noteWhileOpen                                        = false;
    const auto info                                           = scope.Run(FILESYSTEM_COPY,
                                                                          state.fsLocal,
                                                                          {root / L"tree"},
                                                                          root / L"out",
                                                                          32u * 1024u,
                                                                          [&](FolderWindow::FileOperationState::Task& task)
    {
        if (task._discoverySilenceRecorded.load(std::memory_order_acquire) && ! task._discoveryClosed.load(std::memory_order_acquire))
            noteWhileOpen = true;
    },
                                [&]() noexcept { return noteWhileOpen || GetTickCount64() - started > 20'000u; });
    if (! info.has_value())
        return reject(L"the throttled copy did not end after its cancel (or asked an unexpected question)");
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
    operations->CollectCompletedTasks(summaries);
    const auto summary = std::ranges::find_if(summaries, [&](const auto& candidate) noexcept { return candidate.taskId == scope.LastTaskId(); });
    if (summary == summaries.end())
        return reject(L"the canceled copy left no summary");
    const auto notes = std::ranges::count_if(summary->issueDiagnostics, [](const auto& issue) noexcept {
        return issue.category == L"discovery.silence" && issue.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Info;
    });
    AppendLog(std::format(L"Discovery closure silence: noteWhileOpen={} notes={} elapsedMs={}", noteWhileOpen, notes, GetTickCount64() - started));
    if (! noteWhileOpen || notes != 1)
        return reject(std::format(L"expected one listing-paused note while discovery was open: whileOpen={} notes={}", noteWhileOpen, notes));
    return true;
}

// D-6.2: a folder far wider than the queue lead closes once it is listed (names held ahead of the
// queue within the 16 MiB name budget), long before its throttled copy ends. Before the count-ahead
// the scope stayed open until only a queue lead of entries remained (about 9,700 of 10,000 copied).
[[nodiscard]] bool RunDiscoveryScopeTenThousandSmallFilesChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Discovery closure 10,000 files: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.infoLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    constexpr unsigned long kFiles   = 10'000u;
    constexpr size_t kFileBytes      = 1024u;
    const std::filesystem::path root = state.tempRoot / L"discovery-closure-ten-thousand";
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(root / L"out"))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    for (unsigned long index = 0u; index < kFiles; ++index)
    {
        if (! WriteFilledTestFile(root / L"wide" / std::format(L"small-{:05}.bin", index), kFileBytes, static_cast<unsigned char>(index & 0xFFu)))
            return reject(L"wide fixture unavailable");
    }
    if (! SetPluginConfiguration(state.infoLocal.get(), kLocalConcurrency4))
        return reject(L"Local concurrency 4 could not be configured");
    ClosureCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    // Throttled so the copy takes far longer than the listing; stopped once discovery has closed.
    const ULONGLONG started = GetTickCount64();
    const auto info         = scope.Run(FILESYSTEM_COPY,
                                        state.fsLocal,
                                        {root / L"wide"},
                                        root / L"out",
                                        256u * 1024u,
                                        {},
                                        [&]() noexcept
    {
        FolderWindow::FileOperationState::Task* const task = operations->FindTask(scope.LastTaskId());
        return (task != nullptr && task->_discoveryClosed.load(std::memory_order_acquire)) || GetTickCount64() - started > 30'000u;
    });
    if (! info.has_value())
        return reject(L"the throttled copy did not end after its stop (or asked an unexpected question)");
    AppendLog(std::format(L"Discovery closure 10,000 files: closed={} files={} bytes={} mutationsWhileOpen={} bytesWhileOpen={} growthAfterClose={}",
                          info->discoveryClosed,
                          info->discoveredFiles,
                          info->discoveredTotalBytes,
                          info->discoveryMutationsCompletedWhileOpen,
                          info->discoveryBytesCompletedWhileOpen,
                          info->discoveryGrowthAfterCloseCount));
    if (! info->discoveryClosed || info->discoveredFiles != kFiles || info->discoveredTotalBytes != static_cast<uint64_t>(kFiles) * kFileBytes)
        return reject(L"the wide folder did not close on its complete totals");
    if (info->discoveryGrowthAfterCloseCount != 0u)
        return reject(L"discovery grew after it closed");
    if (info->discoveryMutationsCompletedWhileOpen >= 1'000u)
        return reject(std::format(L"the folder closed only after {} of {} files were copied", info->discoveryMutationsCompletedWhileOpen, kFiles));
    return true;
}

// Deterministic production rendering, with captures from the owned test HWND. The provider
// completion observers have separate storage tests; these snapshots isolate the UI contract.
[[nodiscard]] bool RunFileOpsSkipCountingPresentationChecks(SelfTestState& state) noexcept
{
    using namespace FileOperationsPopupInternal;
    using Action      = FileOperations::ConflictAction;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Skip-counting presentation: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || operations->HasActiveOperations())
        return reject(L"requires an idle File Operations state");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"the no-activation guard could not be acquired");
    const auto output = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::FileOperations, L"skip-counting-presentation");
    if (output.empty() || ! SelfTest::EnsureDirectory(output))
        return reject(L"the authorized capture directory could not be created");
    std::ofstream manifest(output / L"scenarios.tsv", std::ios::binary);
    manifest << "file\tscenario\tculture\twidthDip\tdpi\tfixture\tcompileTime\n";
    if (! manifest)
        return reject(L"the capture provenance manifest could not be opened");
    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui->language.empty() && g_settings.ui->language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui->language};
    const auto previousTheme    = folder->GetTheme();
    const bool previousFooter   = operations->GetPopupFooterOnly();
    const bool previousDensity  = operations->GetPopupCompactDensity();
    const bool previousThroughput = operations->GetPopupThroughputExpanded();
    const auto restore          = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage));
        folder->ApplyTheme(previousTheme);
        operations->SetPopupFooterOnly(previousFooter);
        operations->SetPopupCompactDensity(previousDensity);
        operations->SetPopupThroughputExpanded(previousThroughput);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
    });
    auto theme                  = ResolveAppTheme(ThemeMode::Dark, L"SkipCountingPresentation");
    theme.reducedMotionOverride = true;
    folder->ApplyTheme(theme);
    operations->SetPopupFooterOnly(false);
    operations->SetPopupCompactDensity(false);
    operations->SetPopupThroughputExpanded(true);
    MONITORINFO monitor{};
    monitor.cbSize = sizeof(monitor);
    if (! GetMonitorInfoW(MonitorFromWindow(state.mainWindow, MONITOR_DEFAULTTONEAREST), &monitor))
        return reject(L"the test window's work area was unavailable");
    const std::array<std::wstring_view, 9u> names{{L"copy-counts",
                                                   L"copy-files-only",
                                                   L"copy-written-only",
                                                   L"delete-counts",
                                                   L"delete-counts-unavailable",
                                                   L"disk-full",
                                                   L"choice-cache-full",
                                                   L"read-only-folder-delete",
                                                   L"disk-full-unqualified"}};
    size_t captures    = 0u;
    const auto started = std::chrono::steady_clock::now();
    for (const std::wstring_view culture : {std::wstring_view(L"fr-FR"), std::wstring_view(L"en-US")})
    {
        if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, std::wstring(culture)})))
            return reject(L"the requested UI language could not be applied");
        // These consent actions must describe whether MOVE removes the encrypted original,
        // and whether retaining it still publishes a copy at the destination.
        FileOperations::ConflictPromptTextFacts efs{};
        efs.bucket                      = static_cast<uint8_t>(FileOperations::ConflictClass::EfsPlaintext);
        efs.operation                   = FILESYSTEM_COPY;
        const std::wstring copyQuestion = FileOperations::BuildConflictPromptQuestion(efs);
        const std::wstring copyLabel = FileOperations::FormatConflictActionLabel(Action::Proceed, FileOperations::ConflictClass::EfsPlaintext, FILESYSTEM_COPY);
        efs.operation                = FILESYSTEM_MOVE;
        const std::wstring moveQuestion               = FileOperations::BuildConflictPromptQuestion(efs);
        const bool french                             = culture == L"fr-FR";
        const std::wstring_view expectedMoveLabel     = french ? L"Déplacer sans chiffrement" : L"Move without encryption";
        const std::wstring_view expectedCopyKeepLabel = french ? L"Copier et conserver la source" : L"Copy and keep source";
        if (copyQuestion.empty() || copyQuestion != LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_EFS_PLAINTEXT) || moveQuestion.empty() ||
            moveQuestion == copyQuestion || moveQuestion != LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_EFS_MOVE_PLAINTEXT) ||
            FileOperations::BuildConflictPromptAnnouncement(efs).find(moveQuestion) == std::wstring::npos || copyLabel.empty() ||
            copyLabel != LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_PLAINTEXT) ||
            FileOperations::FormatConflictActionLabel(Action::Proceed, FileOperations::ConflictClass::EfsPlaintext, FILESYSTEM_MOVE) != expectedMoveLabel ||
            copyLabel == expectedMoveLabel)
            return reject(std::format(L"{} encrypted MOVE consent did not distinguish source removal from copying", culture));
        for (const auto bucket : {FileOperations::ConflictClass::EfsPlaintext, FileOperations::ConflictClass::MetadataLoss})
        {
            if (FileOperations::FormatConflictActionLabel(Action::RetainSource, bucket, FILESYSTEM_MOVE) != expectedCopyKeepLabel ||
                FileOperations::FormatConflictActionLabel(Action::RetainSource, bucket, FILESYSTEM_COPY) !=
                    LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_KEEP_SOURCE))
                return reject(std::format(L"{} retain-source consent hid that a destination copy is still created", culture));
        }
        for (const int widthDip : {480, 760})
        {
            for (size_t scene = 0u; scene < names.size(); ++scene)
            {
                TaskSnapshot task{};
                task.taskId                   = 0xF26A202610150100ull + scene;
                task.operation                = scene == 3u || scene == 4u || scene == 7u ? FILESYSTEM_DELETE : FILESYSTEM_COPY;
                task.started                  = true;
                task.operationStartTick       = GetTickCount64() - 1'000u;
                task.discoverySkipped         = true;
                task.discoveryClosed          = false;
                task.discoveredFileCount      = 9'999u; // unrelated found work must not masquerade as completed work
                task.discoveredDirectoryCount = 999u;
                task.currentSourcePath        = L"C:\\Fixture\\nested\\source.bin";
                task.currentDestinationPath   = L"D:\\Fixture\\nested\\source.bin";
                task.copiedFileCount          = 7u;
                task.copiedFileCountKnown     = scene != 2u;
                task.copiedFolderCount        = 2u;
                task.copiedFolderCountKnown   = scene == 0u;
                task.completedBytes           = 16u * 1024u * 1024u;
                task.completedWorkFileCount   = 93u;
                task.completedWorkFolderCount = 4u;
                task.completedWorkBytes       = 1024u * 1024u * 1024u;
                task.completedWorkCountsKnown = scene == 3u;
                if (scene == 0u && widthDip == 760)
                {
                    task.discoveryClosed      = true;
                    task.discoveredBytesKnown = true;
                    task.totalBytes           = 32u * 1024u * 1024u;
                }
                std::wstring expected;
                if (scene == 0u)
                    expected = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COPIED_FILES_FOLDERS_WRITTEN, 7u, 2u, FormatBytesCompact(task.completedBytes));
                else if (scene == 1u)
                    expected = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COPIED_FILES_SO_FAR, 7u, FormatBytesCompact(task.completedBytes));
                else if (scene == 2u)
                    expected = FormatStringResource(nullptr, IDS_FMT_FILEOPS_WRITTEN_BYTES, FormatBytesCompact(task.completedBytes));
                else if (scene == 3u)
                    expected = FormatStringResource(nullptr, IDS_FMT_FILEOPS_DELETED_FILES_FOLDERS_BYTES, 93u, 4u, FormatBytesCompact(task.completedWorkBytes));
                else if (scene == 4u)
                    expected = LoadStringResource(nullptr, IDS_FILEOPS_DELETING_WITHOUT_COUNTS);
                else
                {
                    task.conflict.active                        = true;
                    task.conflict.generation                    = 1u;
                    task.conflict.bucket                        = static_cast<uint8_t>(scene == 5u || scene == 8u ? FileOperations::ConflictClass::DiskFull
                                                                                       : scene == 7u ? FileOperations::ConflictClass::ReadOnlyRegularFile
                                                                                                     : FileOperations::ConflictClass::AccessDenied);
                    task.conflict.status                        = HRESULT_FROM_WIN32(scene == 5u || scene == 8u ? ERROR_DISK_FULL : ERROR_ACCESS_DENIED);
                    task.conflict.sourcePath                    = task.currentSourcePath;
                    task.conflict.destinationPath               = task.currentDestinationPath;
                    task.conflict.sourceMetadata.available      = true;
                    task.conflict.destinationMetadata.available = true;
                    if (scene == 6u)
                        task.conflict.repeatUnavailableReason = static_cast<uint8_t>(FileOperations::ConflictRepeatUnavailableReason::CacheFull);
                    if (scene == 8u)
                        task.conflict.repeatUnavailableReason = static_cast<uint8_t>(FileOperations::ConflictRepeatUnavailableReason::UnqualifiedDestination);
                    if (scene == 7u)
                    {
                        task.conflict.sourceIsDirectory          = true;
                        task.conflict.sourceMetadata.isDirectory = true;
                        task.conflict.sourcePath                 = L"C:\\Fixture\\read-only-folder";
                        task.conflict.destinationPath.clear();
                    }
                }
                auto lifetime = std::make_shared<int>(0);
                wil::unique_hwnd popup(FileOperationsPopup::Create(operations, folder, state.mainWindow, lifetime));
                const auto closePopup = wil::scope_exit([&]() noexcept
                {
                    lifetime.reset();
                    popup.reset();
                });
                auto* renderer        = popup ? reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA)) : nullptr;
                if (! renderer)
                    return reject(L"the private production renderer could not be created");
                renderer->DebugSetVisualScenario({task}, false, true);
                if (! SetWindowPos(popup.get(), nullptr, monitor.rcWork.left + 12, monitor.rcWork.top + 12, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE))
                    return reject(L"the popup could not be placed on its owner's display");
                const UINT dpi = GetDpiForWindow(popup.get());
                if (! SetWindowPos(popup.get(),
                                   nullptr,
                                   monitor.rcWork.left + 12,
                                   monitor.rcWork.top + 12,
                                   MulDiv(widthDip, static_cast<int>(dpi), 96),
                                   MulDiv(720, static_cast<int>(dpi), 96),
                                   SWP_NOZORDER | SWP_NOACTIVATE))
                    return reject(L"the popup's exact native-DPI width could not be established");
                ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                PopupLayoutDebugSnapshot painted{};
                painted.taskId = task.taskId;
                if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), painted))
                    return reject(L"the painted layout could not be observed");
                if (scene >= 5u && ! painted.conflictCancelRightmost)
                    return reject(L"Cancel was not physically rightmost among the decision actions");
                if (scene < 5u)
                {
                    const std::wstring operation =
                        LoadStringResource(nullptr, task.operation == FILESYSTEM_COPY ? IDS_FILEOP_OPERATION_COPY : IDS_FILEOP_OPERATION_DELETE);
                    if (painted.taskStatusHeaderPaintedText != operation || ! painted.taskPaintedDiscoveryText.empty() ||
                        painted.taskDiscoveryIndicatorVisible || painted.taskPaintedTransferText != expected ||
                        painted.taskHostedProgressIndeterminate == task.discoveryClosed || ! painted.taskPaintedBodyBoundsValid ||
                        painted.taskPaintedBodyBottomPx + 10.0f > painted.taskPaintedActionRowRect.top ||
                        painted.taskPaintedTransferRect.bottom > painted.taskPaintedBodyBottomPx)
                        return reject(std::format(L"{} {} at {} DIP retained duplicate counts or clipped its work summary", culture, names[scene], widthDip));
                    if (scene == 3u && (painted.hostedGraphControlCount != 1u || ! painted.graphCurrentBandwidthLabelVisible ||
                                        painted.graphCurrentBandwidthBytesPerSecond != 24.0 ||
                                        painted.graphPlotRect.bottom <= painted.graphPlotRect.top ||
                                        painted.graphPlotRect.bottom > painted.taskPaintedActionRowRect.top))
                        return reject(L"delete item-throughput graph/current marker was missing or overlapped the action row");
                    if (scene == 3u)
                    {
                        const float expandedHeight = painted.taskPaintedCardRect.bottom - painted.taskPaintedCardRect.top;
                        const PopupSelfTestInvoke toggle{.kind = PopupHitTest::Kind::TaskToggleThroughput, .taskId = task.taskId};
                        if (! DebugInvokeFileOperationsPopup(popup.get(), toggle))
                            return reject(L"delete throughput disclosure could not be collapsed");
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                        PopupLayoutDebugSnapshot collapsed{};
                        collapsed.taskId = task.taskId;
                        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), collapsed) || collapsed.taskThroughputExpanded ||
                            collapsed.taskPaintedCardRect.bottom - collapsed.taskPaintedCardRect.top >= expandedHeight - static_cast<float>(MulDiv(32, static_cast<int>(dpi), 96)) ||
                            collapsed.taskPaintedBodyBottomPx + 10.0f > collapsed.taskPaintedActionRowRect.top)
                            return reject(L"collapsed delete throughput retained expanded graph space or clipped the body");
                        if (! DebugInvokeFileOperationsPopup(popup.get(), toggle))
                            return reject(L"delete throughput disclosure could not be restored");
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    }
                    if (scene == 4u && painted.taskInlineSpeedRowVisible)
                        return reject(L"a delete route without committed counters showed a fabricated zero rate");
                    std::wstring accessible;
                    std::wstring help;
                    if (! renderer->DebugHostedRegionAccessibleTextForSelfTest(
                            std::format(L"FileOperations.Task.{}.WorkSummary", task.taskId), accessible, help) ||
                        accessible != expected)
                        return reject(L"the one completed-work summary was not exposed with its full accessible text");
                }
                else if (scene == 5u)
                {
                    const std::array<uint8_t, 3u> actions{
                        {static_cast<uint8_t>(Action::Retry), static_cast<uint8_t>(Action::SkipAll), static_cast<uint8_t>(Action::Cancel)}};
                    if (painted.conflictPrimaryActionCount != actions.size() ||
                        ! std::ranges::equal(std::span(painted.conflictPrimaryActions.data(), actions.size()), actions) || painted.conflictMoreVisible ||
                        painted.conflictApplyToAllVisible || ! painted.conflictPromptVisibleWithoutClipping || ! painted.conflictPrimaryLabelsFit)
                        return reject(L"Disk full did not present Retry, Skip all and Cancel directly without a redundant checkbox or More menu");
                }
                else if (scene == 6u)
                {
                    TaskSnapshot projected{};
                    if (! DebugGetFileOperationsPopupTaskSnapshot(popup.get(), task.taskId, projected))
                        return reject(L"the cache-full prompt could not be observed");
                    FileOperations::ConflictPromptTextFacts facts{};
                    facts.bucket                  = projected.conflict.bucket;
                    facts.status                  = projected.conflict.status;
                    facts.repeatUnavailableReason = projected.conflict.repeatUnavailableReason;
                    if (painted.conflictApplyToAllVisible || projected.conflict.skipAllEligible ||
                        FileOperations::BuildConflictPromptQuestion(facts).find(LoadStringResource(nullptr, IDS_FILEOPS_REPEAT_CACHE_FULL)) ==
                            std::wstring::npos)
                        return reject(L"the one-item fallback did not explain that its remembered-choice cache was full");
                }
                else if (scene == 7u)
                {
                    const uint8_t deleteAction = static_cast<uint8_t>(Action::DeleteReadOnly);
                    const auto offered         = std::span(painted.conflictPrimaryActions.data(), painted.conflictPrimaryActionCount);
                    const std::wstring prefix  = std::format(L"FileOperations.Task.{}.Conflict.", task.taskId);
                    std::wstring question;
                    std::wstring help;
                    std::wstring deleteText;
                    std::wstring deleteName;
                    const std::wstring deleteId = std::format(L"FileOperations.Action.{}.{}.{}.{}",
                                                              static_cast<uint32_t>(PopupHitTest::Kind::TaskConflictAction),
                                                              task.taskId,
                                                              static_cast<uint32_t>(Action::DeleteReadOnly),
                                                              task.conflict.generation);
                    if (std::ranges::find(offered, deleteAction) == offered.end() || ! painted.conflictApplyToAllVisible || painted.conflictMoreVisible ||
                        ! painted.conflictPrimaryLabelsFit || ! painted.conflictPromptVisibleWithoutClipping ||
                        ! renderer->DebugHostedButtonTextForSelfTest(deleteId, deleteText, deleteName) ||
                        deleteText != LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_DELETE_READONLY) || deleteName != deleteText ||
                        ! renderer->DebugHostedRegionAccessibleTextForSelfTest(prefix + L"Question", question, help) ||
                        question.find(LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_READONLY_DELETE_FOLDER)) == std::wstring::npos)
                        return reject(L"the read-only folder delete did not state its scope with a direct, fully visible exact-consent action");
                }
                else
                {
                    const std::array<uint8_t, 2u> actions{{static_cast<uint8_t>(Action::Retry), static_cast<uint8_t>(Action::Cancel)}};
                    const std::wstring questionId = std::format(L"FileOperations.Task.{}.Conflict.Question", task.taskId);
                    std::wstring question;
                    std::wstring help;
                    if (painted.conflictPrimaryActionCount != actions.size() ||
                        ! std::ranges::equal(std::span(painted.conflictPrimaryActions.data(), actions.size()), actions) || painted.conflictMoreVisible ||
                        painted.conflictApplyToAllVisible || ! painted.conflictPromptVisibleWithoutClipping || ! painted.conflictPrimaryLabelsFit ||
                        ! renderer->DebugHostedRegionAccessibleTextForSelfTest(questionId, question, help) ||
                        question.find(LoadStringResource(nullptr, IDS_FILEOPS_SKIP_ALL_UNQUALIFIED)) == std::wstring::npos ||
                        question.find(LoadStringResource(nullptr, IDS_FILEOPS_DISK_FULL_SKIP_ALL_NOTE)) != std::wstring::npos)
                        return reject(L"disk full without a qualified destination offered Skip all or omitted the safe Retry/Cancel explanation");
                }
                if (widthDip == 480)
                {
                    const std::wstring file = std::format(L"{}-{}.png", culture, names[scene]);
                    std::array<RedSalamander::TestSupport::WindowScreenshotColorProbe, 1u> probes{};
                    std::span<RedSalamander::TestSupport::WindowScreenshotColorProbe> activeProbes;
                    if (scene < 5u)
                    {
                        const auto& bounds              = painted.taskPaintedTransferRect;
                        const auto& color               = painted.taskPaintedTransferColor;
                        probes[0].clientBounds          = {static_cast<LONG>(std::floor(bounds.left)),
                                                           static_cast<LONG>(std::floor(bounds.top)),
                                                           static_cast<LONG>(std::ceil(bounds.right)),
                                                           static_cast<LONG>(std::ceil(bounds.bottom))};
                        probes[0].expectedColor         = RGB(static_cast<BYTE>(std::clamp(color.r, 0.0f, 1.0f) * 255.0f),
                                                              static_cast<BYTE>(std::clamp(color.g, 0.0f, 1.0f) * 255.0f),
                                                              static_cast<BYTE>(std::clamp(color.b, 0.0f, 1.0f) * 255.0f));
                        probes[0].tolerance             = 12u;
                        probes[0].minimumMatchingPixels = 8u;
                        activeProbes                    = probes;
                    }
                    const HRESULT capture = RedSalamander::TestSupport::SaveWindowScreenshot(popup.get(), output / file, activeProbes);
                    if (FAILED(capture))
                        return reject(std::format(L"{} {} owned-window capture failed: 0x{:08X}", culture, names[scene], static_cast<unsigned long>(capture)));
                    manifest << Common::Strings::Utf8FromUtf16ReplacingInvalid(file) << '\t' << Common::Strings::Utf8FromUtf16ReplacingInvalid(names[scene])
                             << '\t' << Common::Strings::Utf8FromUtf16ReplacingInvalid(culture) << '\t' << widthDip << '\t' << dpi
                             << "\tactual-application-deterministic-snapshot\t" << __DATE__ << ' ' << __TIME__ << '\n';
                    ++captures;
                }
            }
        }
    }
    const auto elapsed = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());
    Debug::Perf::Emit(L"FileOps.SelfTest.SkipCountingPresentationUs", L"candidate-only;owned-window-captures", elapsed, names.size() * 4u, captures, S_OK);
    AppendLog(std::format(
        L"Skip-counting presentation: {} scenarios, fr-FR/en-US, 480/760 DIP, {} owned-window captures, candidateUs={}", names.size(), captures, elapsed));
    return true;
}
