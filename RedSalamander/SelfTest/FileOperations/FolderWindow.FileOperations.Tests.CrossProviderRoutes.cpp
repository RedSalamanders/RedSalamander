// I35 S7 (Specs/Plans/Done/FileOperations_CopyMoveUxAndReliabilityReview_2026-10-11.md, section 11):
// cross-provider routes the provider-matrix lens found broken.
// XP-01: a new file copied to an MTP device publishes through the device writer (MTP has no object
//        binding, so the host has no owned stage of its own there).
// XP-04: requested verification never removes a destination's only publication route, whether or
//        not the confirmation is shown; a writer that proves its own digest is a verification route.
// XP-10: a source that reports no file sizes fails each file with its own Error, and the task names
//        that cause once as an Information note.
// XP-09: a Copy between two roots of one Local instance runs the provider's own CopyItem and
//        publishes the exact bytes.

namespace
{
constexpr auto kCrossProviderWait = std::chrono::seconds(30);

[[nodiscard]] std::optional<FolderWindow::FileOperationState::CompletedTaskSummary> FindCrossProviderSummary(FolderWindow::FileOperationState* operations,
                                                                                                             uint64_t taskId)
{
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
    operations->CollectCompletedTasks(summaries);
    const auto found = std::ranges::find_if(summaries, [taskId](const auto& summary) noexcept { return summary.taskId == taskId; });
    if (found == summaries.end())
        return std::nullopt;
    return *found;
}

[[nodiscard]] std::wstring DescribeCrossProviderDiagnostics(const FolderWindow::FileOperationState::CompletedTaskSummary& summary)
{
    std::wstring text;
    for (const auto& issue : summary.issueDiagnostics)
    {
        if (! text.empty())
            text += L"; ";
        text += std::format(L"{}:{}:0x{:08X}", issue.category, static_cast<unsigned int>(issue.severity), static_cast<unsigned long>(issue.status));
    }
    return text;
}

[[nodiscard]] std::vector<std::wstring> ListCrossProviderChildNames(IFileSystem* fs, std::wstring_view path)
{
    std::vector<std::wstring> names;
    wil::com_ptr<IFilesInformation> files;
    const std::wstring pathText(path);
    FileInfo* head = nullptr;
    if (! fs || FAILED(fs->ReadDirectoryInfo(pathText.c_str(), files.put())) || ! files || FAILED(files->GetBuffer(&head)) || head == nullptr)
        return names;
    for (FileInfo* entry = head; entry != nullptr;)
    {
        if (entry->FileNameSize >= sizeof(wchar_t))
        {
            std::wstring name(entry->FileName, entry->FileName + entry->FileNameSize / sizeof(wchar_t));
            if (name != L"." && name != L"..")
                names.push_back(std::move(name));
        }
        if (entry->NextEntryOffset == 0u)
            break;
        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
    }
    return names;
}

// Runs one admitted transfer from `source` to completion and returns its summary; the task is dismissed
// by the caller. A Local source runs from the left pane, any other source from the right one.
[[nodiscard]] std::optional<FolderWindow::FileOperationState::CompletedTaskSummary> RunCrossProviderTransferFrom(
    SelfTestState& state,
    FolderWindow::FileOperationState* operations,
    std::vector<uint64_t>& tasks,
    const wil::com_ptr<IFileSystem>& source,
    const wil::com_ptr<IFileSystem>& destination,
    std::vector<std::filesystem::path> sources,
    const std::filesystem::path& destinationFolder,
    bool requireConfirmation,
    FileSystemFlags flags = FILESYSTEM_FLAG_RECURSIVE)
{
    const bool localSource = source.get() == state.fsLocal.get();
    const auto admitted    = StartFileOperationAndGetId(operations,
                                                        FILESYSTEM_COPY,
                                                        localSource ? FolderWindow::Pane::Left : FolderWindow::Pane::Right,
                                                        localSource ? FolderWindow::Pane::Right : FolderWindow::Pane::Left,
                                                        source,
                                                        std::move(sources),
                                                        destinationFolder,
                                                        flags,
                                                        false,
                                                        0u,
                                                        FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                        requireConfirmation,
                                                        destination);
    if (! admitted.has_value())
        return std::nullopt;
    const uint64_t taskId = admitted.value();
    tasks.push_back(taskId);
    const auto done = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        if (TryGetConflictPromptCopy(operations->FindTask(taskId)).has_value())
            return true; // an unexpected question: report it rather than wait for the timeout
        return state.completedTasks.contains(taskId) && operations->FindTask(taskId) == nullptr;
    },
        {.timeout = SelfTest::Scale(kCrossProviderWait), .operationName = L"cross-provider transfer completion"});
    if (const auto prompt = TryGetConflictPromptCopy(operations->FindTask(taskId)); prompt.has_value())
        AppendLog(std::format(L"Cross-provider transfer asked an unexpected question: bucket={} status=0x{:08X} source='{}'",
                              static_cast<unsigned int>(prompt->bucket),
                              static_cast<unsigned long>(prompt->status),
                              prompt->sourcePath));
    if (! done.conditionMet || ! state.completedTasks.contains(taskId))
        return std::nullopt;
    return FindCrossProviderSummary(operations, taskId);
}

// Runs one admitted transfer from the Local provider; see RunCrossProviderTransferFrom.
[[nodiscard]] std::optional<FolderWindow::FileOperationState::CompletedTaskSummary> RunCrossProviderTransfer(SelfTestState& state,
                                                                                                             FolderWindow::FileOperationState* operations,
                                                                                                             std::vector<uint64_t>& tasks,
                                                                                                             const wil::com_ptr<IFileSystem>& destination,
                                                                                                             std::vector<std::filesystem::path> sources,
                                                                                                             const std::filesystem::path& destinationFolder,
                                                                                                             bool requireConfirmation,
                                                                                                             FileSystemFlags flags = FILESYSTEM_FLAG_RECURSIVE)
{
    return RunCrossProviderTransferFrom(
        state, operations, tasks, state.fsLocal, destination, std::move(sources), destinationFolder, requireConfirmation, flags);
}

class CrossProviderCaseScope final
{
public:
    CrossProviderCaseScope(SelfTestState& state, FolderWindow::FileOperationState* operations) noexcept : _state(state), _operations(operations)
    {
        _verify        = EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy;
        _autoDismiss   = operations->GetAutoDismissSuccess();
        _originalPopup = operations->GetPopupHwndForSelfTest();
        operations->SetAutoDismissSuccess(false);
    }
    CrossProviderCaseScope(const CrossProviderCaseScope&)            = delete;
    CrossProviderCaseScope& operator=(const CrossProviderCaseScope&) = delete;
    ~CrossProviderCaseScope()
    {
        const bool quiet = CancelOutstandingSelfTestOperations(_operations);
        for (const uint64_t taskId : tasks)
        {
            if (quiet)
            {
                _operations->DismissCompletedTask(taskId);
                _operations->DebugRemoveDiagnosticsForTask(taskId);
                _state.completedTasks.erase(taskId);
            }
        }
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = _verify;
        _operations->SetAutoDismissSuccess(_autoDismiss);
        static_cast<void>(_operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
        if (quiet && ! _originalPopup && _operations->GetPopupHwndForSelfTest())
            _operations->DebugResetPopupForSelfTest();
    }

    std::vector<uint64_t> tasks;

private:
    SelfTestState& _state;
    FolderWindow::FileOperationState* _operations = nullptr;
    HWND _originalPopup                           = nullptr;
    bool _verify                                  = false;
    bool _autoDismiss                             = false;
};
} // namespace

[[nodiscard]] bool RunFileOpsCrossProviderMtpNewFileChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"XP-01 MTP new file: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    using CreateMtpForSelfTestFunc = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const char*, void**) noexcept;
    wil::unique_hmodule module;
    FARPROC createAddress = nullptr;
    HRESULT hr            = SelfTest::LoadMtpPluginSelfTestExport("RedSalamanderMtpCreateForSelfTest", module, createAddress);
    if (FAILED(hr) || createAddress == nullptr)
        return reject(std::format(L"fake-MTP factory unavailable (0x{:08X})", static_cast<unsigned long>(hr)));
#pragma warning(push)
#pragma warning(disable : 4191) // The named self-test export fixes the typed factory ABI.
    const auto createMtp = reinterpret_cast<CreateMtpForSelfTestFunc>(createAddress);
#pragma warning(pop)
    FactoryOptions factoryOptions{};
    factoryOptions.debugLevel = DEBUG_LEVEL_NONE;
    wil::com_ptr<IFileSystem> mtp;
    hr = createMtp(__uuidof(IFileSystem), &factoryOptions, GetHostServices(), R"json({})json", mtp.put_void());
    wil::com_ptr<IInformations> mtpInfo;
    wil::com_ptr<IFileSystemInitialize> mtpInitialize;
    wil::com_ptr<IFileSystemIO> mtpIo;
    wil::com_ptr<IFileSystemDirectoryOperations> mtpDirectories;
    if (SUCCEEDED(hr) && mtp)
        hr = mtp->QueryInterface(IID_PPV_ARGS(mtpInfo.addressof()));
    if (SUCCEEDED(hr))
        hr = mtpInfo->SetConfiguration(R"json({"readOnly":false,"commandTimeoutMs":5000})json");
    if (SUCCEEDED(hr))
        hr = mtp->QueryInterface(IID_PPV_ARGS(mtpInitialize.addressof()));
    if (SUCCEEDED(hr))
        hr = mtpInitialize->Initialize(L"/", nullptr);
    if (SUCCEEDED(hr))
        hr = mtp->QueryInterface(IID_PPV_ARGS(mtpIo.addressof()));
    if (SUCCEEDED(hr))
        hr = mtp->QueryInterface(IID_PPV_ARGS(mtpDirectories.addressof()));
    if (FAILED(hr) || ! mtp || ! mtpIo || ! mtpDirectories)
        return reject(std::format(L"fake MTP could not be initialized (0x{:08X})", static_cast<unsigned long>(hr)));

    const std::filesystem::path root = state.tempRoot / L"xp01-mtp-new-file";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot             = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path file  = root / L"photo-copy.bin";
    const std::filesystem::path tree  = root / L"album";
    const std::filesystem::path child = tree / L"inner.bin";
    if (! WriteFilledTestFile(file, 65'537u, 0x5Au) || ! WriteFilledTestFile(child, 4'099u, 0x3Cu))
        return reject(L"local sources unavailable");
    constexpr std::wstring_view kMtpDestination = L"/Fake Phone/Internal Storage/xp01";
    if (FAILED(mtpDirectories->CreateDirectory(std::wstring(kMtpDestination).c_str())))
        return reject(L"fake-MTP destination folder could not be created");

    CrossProviderCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    const auto summary = RunCrossProviderTransfer(state, operations, scope.tasks, mtp, {file, tree}, std::filesystem::path(kMtpDestination), false);
    if (! summary.has_value())
        return reject(L"the Local to MTP Copy did not complete (or asked an unexpected question)");
    std::string fileBytes;
    std::string childBytes;
    std::string expectedFile;
    std::string expectedChild;
    const bool read             = ReadFileTextFsIo(mtpIo, std::filesystem::path(kMtpDestination) / L"photo-copy.bin", fileBytes) &&
                                  ReadFileTextFsIo(mtpIo, std::filesystem::path(kMtpDestination) / L"album" / L"inner.bin", childBytes) &&
                                  ReadFileTextFsIo(state.fsLocal.query<IFileSystemIO>(), file, expectedFile) &&
                                  ReadFileTextFsIo(state.fsLocal.query<IFileSystemIO>(), child, expectedChild);
    const auto destinationNames = ListCrossProviderChildNames(mtp.get(), kMtpDestination);
    const bool leftover         = std::ranges::any_of(destinationNames, [](const std::wstring& name) noexcept {
        return name.find(L".rs-mtp-") != std::wstring::npos || name.find(L".rs_tmp_") != std::wstring::npos;
    });
    if (summary->resultHr != S_OK || summary->errorCount != 0u || ! read || fileBytes != expectedFile || childBytes != expectedChild || leftover ||
        std::ranges::any_of(summary->issueDiagnostics, [](const auto& issue) noexcept { return issue.category == L"bridge.stage.authorityUnavailable"; }))
        return reject(std::format(L"new files did not publish on the device: hr=0x{:08X} errors={} read={} leftover={} names={} diagnostics=[{}]",
                                  static_cast<unsigned long>(summary->resultHr),
                                  summary->errorCount,
                                  read,
                                  leftover,
                                  destinationNames.size(),
                                  DescribeCrossProviderDiagnostics(summary.value())));
    Debug::Perf::Emit(L"FileOps.SelfTest.CrossProviderMtpNewFile", L"local-to-mtp=new-file+new-folder;no-temp-left", 0u, 2u, 0u, S_OK);
    return true;
}

[[nodiscard]] bool RunFileOpsVerifyAvailabilityChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"XP-04 verification availability: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.fsDummy || ! state.infoDummy || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local/Dummy providers or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    constexpr std::string_view kDummyDefault =
        R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json";
    constexpr std::string_view kDummyWithoutProof =
        R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0","writerProof":false})json";
    const auto restoreDummy          = wil::scope_exit([&]() noexcept { static_cast<void>(SetPluginConfiguration(state.infoDummy.get(), kDummyDefault)); });
    const std::filesystem::path root = state.tempRoot / L"xp04-verify-availability";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const auto source     = root / L"payload.bin";
    if (! WriteFilledTestFile(source, 32'771u, 0x6Du))
        return reject(L"local source unavailable");
    CrossProviderCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = true;

    // 1. No proof at all (an identity-less destination without a writer digest), Verify on, no
    //    confirmation: the copy publishes; verification is simply not requested.
    if (! SetPluginConfiguration(state.infoDummy.get(), kDummyWithoutProof))
        return reject(L"Dummy could not drop its writer proof");
    const std::wstring withoutProof = L"/xp04/without-proof";
    if (! EnsureDummyFolderExists(state.fsDummy.get(), L"/xp04") || ! EnsureDummyFolderExists(state.fsDummy.get(), withoutProof))
        return reject(L"Dummy destination unavailable");
    const auto unproved = RunCrossProviderTransfer(state, operations, scope.tasks, state.fsDummy, {source}, std::filesystem::path(withoutProof), false);
    if (! unproved.has_value() || unproved->resultHr != S_OK || unproved->errorCount != 0u || unproved->verifiedItemCount != 0u ||
        std::ranges::any_of(unproved->issueDiagnostics, [](const auto& issue) noexcept {
        return issue.category == L"bridge.stage.authorityUnavailable" || issue.category == L"bridge.verification.unavailable";
    }))
        return reject(std::format(L"Verify on a destination without any proof refused to publish: hr=0x{:08X} errors={} verified={} diagnostics=[{}]",
                                  unproved.has_value() ? static_cast<unsigned long>(unproved->resultHr) : 0ul,
                                  unproved.has_value() ? unproved->errorCount : 0ul,
                                  unproved.has_value() ? unproved->verifiedItemCount : 0ul,
                                  unproved.has_value() ? DescribeCrossProviderDiagnostics(unproved.value()) : std::wstring{}));

    // 2. A writer that proves its digest, behind the confirmation: Verify stays on and the item is verified.
    if (! SetPluginConfiguration(state.infoDummy.get(), kDummyDefault))
        return reject(L"Dummy could not restore its writer proof");
    const std::wstring withProof = L"/xp04/with-proof";
    if (! EnsureDummyFolderExists(state.fsDummy.get(), withProof))
        return reject(L"Dummy proof destination unavailable");
    const auto proved = RunCrossProviderTransfer(state, operations, scope.tasks, state.fsDummy, {source}, std::filesystem::path(withProof), true);
    if (! proved.has_value() || proved->resultHr != S_OK || proved->verifiedItemCount != 1u)
        return reject(std::format(L"a writer-digest destination lost Verify at the confirmation: hr=0x{:08X} verified={} diagnostics=[{}]",
                                  proved.has_value() ? static_cast<unsigned long>(proved->resultHr) : 0ul,
                                  proved.has_value() ? proved->verifiedItemCount : 0ul,
                                  proved.has_value() ? DescribeCrossProviderDiagnostics(proved.value()) : std::wstring{}));
    Debug::Perf::Emit(L"FileOps.SelfTest.VerifyAvailability", L"no-proof=copied-unverified;writer-digest=verified-after-confirmation", 0u, 2u, 0u, S_OK);
    return true;
}

[[nodiscard]] SelfTest::SelfTestCaseResult RunFileOpsCrossRootLocalCopyChecks(SelfTestState& state,
                                                                              std::optional<LoopbackShareStatus> probeOverride = std::nullopt) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept -> SelfTest::SelfTestCaseResult
    {
        AppendLog(std::format(L"XP-09 cross-root Local Copy: {}", reason));
        return {.status = SelfTest::SelfTestCaseResult::Status::failed, .reason = std::wstring(reason)};
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    // Test-only injection exercises the actual result path without creating or exposing a share.
    const LoopbackShareStatus uncStatus = probeOverride.has_value() ? probeOverride.value() : ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");
    if (uncStatus == LoopbackShareStatus::Unreachable)
    {
        AppendLog(L"XP-09 cross-root Local Copy: the loopback administrative share is not reachable; environment-gated case skipped.");
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = L"Loopback administrative share is unreachable; no cross-root qualification."};
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"xp09-cross-root";
    if (! RecreateEmptyDirectory(root / L"src") || ! RecreateEmptyDirectory(root / L"dst"))
        return reject(L"case root unavailable");
    const auto removeRoot                   = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::filesystem::path source      = root / L"src" / L"cross-root.bin";
    const std::filesystem::path destination = LoopbackShareAlias(root / L"dst");
    if (! WriteFilledTestFile(source, 96u * 1024u + 5u, 0x4Bu))
        return reject(L"local source unavailable");
    FileOperations::QualifiedEndpoint sourceEndpoint{};
    FileOperations::QualifiedEndpoint destinationEndpoint{};
    if (! FileOperations::TryQualifyEndpointForSelfTest(
            state.fsLocal, source.native(), FILESYSTEM_COPY, L"builtin/file-system", L"host/default", sourceEndpoint) ||
        ! FileOperations::TryQualifyEndpointForSelfTest(
            state.fsLocal, destination.native(), FILESYSTEM_COPY, L"builtin/file-system", L"host/default", destinationEndpoint) ||
        sourceEndpoint.rootId.empty() || sourceEndpoint.rootId == destinationEndpoint.rootId)
        return reject(L"the volume and its share alias must qualify as two roots of the Local provider");

    CrossProviderCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    // Same plugin and instance context: no destination file system, the provider's own CopyItem route.
    const auto summary = RunCrossProviderTransfer(state, operations, scope.tasks, wil::com_ptr<IFileSystem>{}, {source}, destination, false);
    if (! summary.has_value() || summary->resultHr != S_OK || summary->publishedItemCount != 1u)
        return reject(
            std::format(L"the cross-root Copy did not publish (hr=0x{:08X})", summary.has_value() ? static_cast<unsigned long>(summary->resultHr) : 0ul));
    const auto info = state.completedTasks.find(scope.tasks.back());
    std::string copied;
    std::string expected;
    const bool read = ReadFileTextFsIo(state.fsLocal.query<IFileSystemIO>(), root / L"dst" / L"cross-root.bin", copied) &&
                      ReadFileTextFsIo(state.fsLocal.query<IFileSystemIO>(), source, expected);
    if (info == state.completedTasks.end() || info->second.bridgeFileAdmissionCount != 0u || ! read || copied != expected)
        return reject(std::format(L"the Copy did not run the provider route with exact bytes (bridgeAdmissions={} read={})",
                                  info == state.completedTasks.end() ? 0u : info->second.bridgeFileAdmissionCount,
                                  read));
    return {.status = SelfTest::SelfTestCaseResult::Status::passed};
}

// Use the real archive reader and File Operations bridge, rather than a second stream-copy
// implementation. Each destination is owned by this case; the UNC spelling aliases only its sandbox.
[[nodiscard]] SelfTest::SelfTestCaseResult RunFileOpsArchiveNestedCopyChecks(SelfTestState& state, bool network) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept -> SelfTest::SelfTestCaseResult
    {
        AppendLog(std::format(L"Archive nested Copy: {}", reason));
        return {.status = SelfTest::SelfTestCaseResult::Status::failed, .reason = std::wstring(reason)};
    };
    auto* operations = TryGetFileOps(TryGetFolderWindow(state.mainWindow));
    if (! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    if (network)
    {
        const auto uncStatus = ProbeLoopbackShare(state.tempRoot);
        if (uncStatus == LoopbackShareStatus::Refused)
            return reject(L"sandbox guard refused the UNC alias");
        if (uncStatus == LoopbackShareStatus::Unreachable)
            return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = L"Loopback administrative share is unreachable; no SMB qualification."};
    }
    const std::filesystem::path root = state.tempRoot / (network ? L"archive-copy-unc" : L"archive-copy-local");
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    std::error_code ec;
    const auto repository = Common::Testing::FindTestRepositoryRoot(ec);
    if (ec || repository.empty())
        return reject(L"fixture repository unavailable");
    const auto fixtures = repository / L"Plugins" / L"FileSystem7z" / L"Tests";
    for (const auto* name : {L"Tests.zip", L"Tests.7z"})
    {
        if (! std::filesystem::copy_file(fixtures / name, root / name, ec) || ec)
            return reject(std::format(L"archive fixture copy failed: {} ({})", name, ec.value()));
    }
    // Mount the first RAR volume only after all four siblings are inside the sandbox.
    for (unsigned int volume = 1u; volume <= 4u; ++volume)
    {
        const std::wstring name = std::format(L"test_read_format_rar_multivolume.part{:04}.rar", volume);
        if (! std::filesystem::copy_file(fixtures / L"MultipartRar" / name, root / name, ec) || ec)
            return reject(std::format(L"RAR volume fixture copy failed: {} ({})", name, ec.value()));
    }

    wil::unique_hmodule module;
    FARPROC createAddress = nullptr;
    if (FAILED(SelfTest::LoadPluginSelfTestExport(kPluginId7z, "RedSalamanderCreate", module, createAddress)) || ! createAddress)
        return reject(L"archive factory unavailable");
    wil::com_ptr<IFileSystem> archive;
    FactoryOptions factoryOptions{};
    factoryOptions.debugLevel = DEBUG_LEVEL_NONE;
    using CreateFn            = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);
#pragma warning(push)
#pragma warning(disable : 4191) // The named plugin factory fixes the typed ABI.
    const HRESULT created =
        reinterpret_cast<CreateFn>(createAddress)(__uuidof(IFileSystem), &factoryOptions, GetHostServices(), kPluginId7z.data(), archive.put_void());
#pragma warning(pop)
    wil::com_ptr<IFileSystemInitialize> initialize;
    wil::com_ptr<IFileSystemIO> archiveIo;
    if (FAILED(created) || ! archive || ! archive.try_query_to(initialize.put()) || ! archive.try_query_to(archiveIo.put()))
        return reject(L"archive instance or reader interfaces unavailable");
    CrossProviderCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;

    struct ArchiveFixture
    {
        std::wstring_view name;
        std::wstring_view rootPath;
        std::wstring_view rootLeaf;
        uint64_t rootBytes;
        std::wstring_view nestedFolder;
        std::wstring_view nestedLeaf;
        uint64_t nestedBytes;
    };
    const std::array<ArchiveFixture, 3> archives = {{
        {L"Tests.zip", L"/Tests/README.txt", L"README.txt", 1521u, L"/Tests/folders with text", L"ascii.txt", 869u},
        {L"Tests.7z", L"/Tests/README.txt", L"README.txt", 1521u, L"/Tests/folders with text", L"ascii.txt", 869u},
        {L"test_read_format_rar_multivolume.part0001.rar", L"/LibarchiveAddingTest.html", L"LibarchiveAddingTest.html", 20111u, L"/testdir", L"test.txt", 20u},
    }};
    for (const auto& fixture : archives)
    {
        const auto name = fixture.name;
        if (initialize->Initialize((root / name).c_str(), nullptr) != S_OK)
            return reject(std::format(L"archive mount failed: {}", name));
        const std::wstring canonicalNested = std::format(L"{}/{}", fixture.nestedFolder, fixture.nestedLeaf);
        std::wstring backslashNested       = canonicalNested;
        std::replace(backslashNested.begin(), backslashNested.end(), L'/', L'\\');
        const std::array<std::filesystem::path, 4> sources = {
            fixture.rootPath,
            canonicalNested,
            std::filesystem::path(fixture.nestedFolder) / fixture.nestedLeaf, // Exact FolderView Windows join.
            backslashNested,
        };
        for (size_t index = 0u; index < sources.size(); ++index)
        {
            const auto destination = root / std::format(L"{}-{}", name, index);
            if (! RecreateEmptyDirectory(destination))
                return reject(L"destination unavailable");
            const auto providerDestination = network ? LoopbackShareAlias(destination) : destination;
            if (providerDestination.empty())
                return reject(L"destination alias refused");
            const std::wstring canonical = index == 0u ? std::wstring(fixture.rootPath) : canonicalNested;
            const std::wstring leaf      = std::wstring(index == 0u ? fixture.rootLeaf : fixture.nestedLeaf);
            std::string expected;
            if (! ReadFileTextFsIo(archiveIo, canonical, expected) || expected.size() != (index == 0u ? fixture.rootBytes : fixture.nestedBytes))
            {
                // Failure-only probe preserves the reader's HRESULT instead of hiding it behind
                // the shared whole-file helper's boolean result.
                wil::com_ptr<IFileReader> probe;
                const HRESULT openHr = archiveIo->CreateFileReader(canonical.c_str(), probe.put());
                uint64_t size        = 0u;
                const HRESULT sizeHr = probe ? probe->GetSize(&size) : E_POINTER;
                std::array<char, 16> bytes{};
                unsigned long read   = 0u;
                const HRESULT readHr = probe ? probe->Read(bytes.data(), static_cast<unsigned long>(bytes.size()), &read) : E_POINTER;
                return reject(std::format(L"canonical archive contents unavailable: archive={} path={} observed={} open=0x{:08X} sizeHr=0x{:08X} "
                                          L"size={} readHr=0x{:08X} read={}",
                                          name,
                                          canonical,
                                          expected.size(),
                                          static_cast<unsigned long>(openHr),
                                          static_cast<unsigned long>(sizeHr),
                                          size,
                                          static_cast<unsigned long>(readHr),
                                          read));
            }

            // Preparation must still reject invalid leaves, including a whole relative path passed
            // as a name. A correct source contract, rather than weakened validation, permits Copy.
            const auto sourceRoute      = FileSystemRouteContract::Query(archive.get(), sources[index].native(), FILESYSTEM_COPY, kPluginId7z);
            const auto destinationRoute = FileSystemRouteContract::Query(state.fsLocal.get(), providerDestination.native(), FILESYSTEM_COPY, kPluginIdLocal);
            if (! sourceRoute.snapshot.pathIdentity.has_value() || ! destinationRoute.snapshot.pathIdentity.has_value())
                return reject(L"source or destination path identity unavailable");
            for (const std::wstring_view invalidLeaf : {L"CON", L"trailing-dot.", LR"(folder\ascii.txt)", L"folder/ascii.txt"})
            {
                const auto invalid = FileSystemRouteContract::QueryChildNameContract(
                    state.fsLocal.get(), providerDestination.native(), invalidLeaf, FILESYSTEM_COPY, kPluginIdLocal);
                if (invalid.state != FileSystemRouteContract::QueryState::Available || invalid.nameStatus != FILESYSTEM_CHILD_NAME_INVALID ||
                    ! invalid.joinedPath.empty())
                    return reject(L"destination name validation was bypassed");
            }
            for (const std::wstring_view invalidLeaf : {L"CON", L"trailing-dot."})
            {
                FileOperations::TransferPlan rejected;
                rejected.sourceEndpoint.pathIdentity      = sourceRoute.snapshot.pathIdentity;
                rejected.destinationEndpoint.pathIdentity = destinationRoute.snapshot.pathIdentity;
                rejected.destinationEndpoint.pluginId     = kPluginIdLocal;
                rejected.destination.providerFolderPath   = providerDestination.native();
                rejected.selectedItems.push_back({.providerPath = std::format(L"/Tests/{}", invalidLeaf)});
                const FileOperations::FileOperationPlanGroup plans{std::move(rejected)};
                FolderWindow::FileOperationState::TransferDestinationNameFailure failure;
                const std::atomic<bool> cancelled{false};
                const HRESULT hr =
                    FolderWindow::FileOperationState::ValidateTransferDestinationNames(plans, state.fsLocal.get(), FILESYSTEM_COPY, {}, cancelled, failure);
                if (hr != HRESULT_FROM_WIN32(invalidLeaf == L"CON" ? ERROR_BAD_DEVICE : ERROR_INVALID_NAME) || failure.rejectedLeaf != invalidLeaf ||
                    failure.destinationFolder != providerDestination.native() ||
                    ! ListCrossProviderChildNames(state.fsLocal.get(), destination.native()).empty())
                    return reject(L"Preparing did not refuse the invalid leaf before mutation");
            }
            const auto started = std::chrono::steady_clock::now();
            const auto summary =
                RunCrossProviderTransferFrom(state, operations, scope.tasks, archive, state.fsLocal, {sources[index]}, providerDestination, false);
            std::string copied;
            const auto children = ListCrossProviderChildNames(state.fsLocal.get(), destination.native());
            if (! summary.has_value() || summary->resultHr != S_OK || summary->errorCount != 0u || summary->publishedItemCount != 1u ||
                ! ReadFileTextFsIo(state.fsLocal.query<IFileSystemIO>(), destination / leaf, copied) || copied != expected || children.size() != 1u ||
                children.front() != leaf)
                return reject(std::format(L"{} source '{}' did not publish exactly '{}' (hr=0x{:08X}, diagnostics=[{}])",
                                          name,
                                          sources[index].native(),
                                          leaf,
                                          summary.has_value() ? static_cast<unsigned long>(summary->resultHr) : 0ul,
                                          summary.has_value() ? DescribeCrossProviderDiagnostics(summary.value()) : std::wstring{}));
            Debug::Perf::Emit(L"FileOps.SelfTest.ArchiveNestedCopy",
                              std::format(L"target={};archive={};spelling={}", network ? L"unc" : L"local", name, index),
                              Debug::Perf::ElapsedUs(started),
                              expected.size(),
                              1u,
                              S_OK);
        }
    }
    return {.status = SelfTest::SelfTestCaseResult::Status::passed};
}

[[nodiscard]] bool RunFileOpsSizelessSourceNoteChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"XP-10 sizeless source: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.fsDummy || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local/Dummy providers or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    // The source is a Dummy folder: the bridge opens a Dummy reader through IFileSystemIO, which the
    // self-test decorator wraps (a Local source opens its reader from a bound object instead).
    const std::filesystem::path root = state.tempRoot / L"xp10-sizeless-source";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot        = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    const std::wstring dummyRoot = L"/xp10-sizeless-source";
    if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
        return reject(L"Dummy source folder unavailable");
    std::vector<std::filesystem::path> sources;
    for (unsigned int index = 0u; index < 3u; ++index)
    {
        const std::wstring path = std::format(L"{}/listing-{}.bin", dummyRoot, index);
        if (! DummyWriteTextFile(state.fsDummy.get(), path, std::format("sizeless-source-{}", index), true))
            return reject(L"Dummy sources unavailable");
        sources.emplace_back(path);
    }

    CrossProviderCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    // Every source reader of this task reports that its size is unavailable, as a sizeless FTP server does.
    SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(3u, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    static_cast<void>(TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest());
    const auto disarm            = wil::scope_exit([]() noexcept
    {
        SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0u);
        static_cast<void>(TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest());
    });
    const auto summary           = RunCrossProviderTransferFrom(state,
                                                                operations,
                                                                scope.tasks,
                                                                state.fsDummy,
                                                                state.fsLocal,
                                                                sources,
                                                                root,
                                                                false,
                                                                static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR));
    const unsigned long injected = TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest();
    if (! summary.has_value())
        return reject(L"the Dummy to Local Copy did not complete (or asked an unexpected question)");
    const auto count = [&](std::wstring_view category, FolderWindow::FileOperationState::DiagnosticSeverity severity) noexcept
    {
        return std::ranges::count_if(summary->issueDiagnostics,
                                     [&](const auto& issue) noexcept { return issue.category == category && issue.severity == severity; });
    };
    const auto fileErrors = count(L"bridge.integrity.sourceSizeUnknown", FolderWindow::FileOperationState::DiagnosticSeverity::Error);
    const auto serverNote = count(L"bridge.integrity.sourceSizeUnknown.server", FolderWindow::FileOperationState::DiagnosticSeverity::Info);
    const auto names      = ListCrossProviderChildNames(state.fsLocal.get(), root.native());
    if (injected != 3u || fileErrors != 3 || serverNote != 1 || summary->publishedItemCount != 0u || ! names.empty() || summary->errorCount < 3u)
        return reject(std::format(L"expected three per-file Errors, one server note and nothing published: injected={} errors={} notes={} published={} "
                                  L"names={} diagnostics=[{}]",
                                  injected,
                                  fileErrors,
                                  serverNote,
                                  summary->publishedItemCount,
                                  names.size(),
                                  DescribeCrossProviderDiagnostics(summary.value())));
    return true;
}
