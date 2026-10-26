// Included by the canonical FileOperations selftest TU; not a separate TU.
// Policy-specific closed seed, two known directories, large source and old occupant.
// The stage is excluded ONLY by its exact owned-stage proof path while parked.
namespace OwnedShutdownFixture
{
constexpr size_t kSourceBytes        = 16u * 1024u * 1024u + 4099u;
constexpr std::string_view kOldBytes = "D10 original occupant: preserve these distinct old bytes";
constexpr std::string_view kSentinel = "D10 unrelated destination sentinel";
struct Entry final
{
    std::filesystem::path relative;
    bool directory = false;
    std::string bytes;
    Common::Crypto::Blake3Digest digest{};
    bool operator==(const Entry&) const noexcept = default;
};
[[nodiscard]] bool Capture(const std::filesystem::path& data,
                           const wil::com_ptr<IFileSystemIO>& io,
                           std::vector<Entry>& out,
                           const std::filesystem::path& permittedStage = {}) noexcept
{
    out.clear();
    const std::array<std::pair<std::filesystem::path, bool>, 6> expected{
        {{L".", true}, {L"src", true}, {L"src/payload.bin", false}, {L"dst", true}, {L"dst/payload.bin", false}, {L"dst/sentinel.bin", false}}};
    for (const auto& directory : std::array{data, data / L"src", data / L"dst"})
    {
        const DWORD attributes = GetFileAttributesW(directory.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
            return false;
        if (directory == data)
            out.push_back(Entry{L".", true, {}, {}});
        std::error_code ec;
        std::filesystem::directory_iterator it(directory, std::filesystem::directory_options::none, ec);
        if (ec)
            return false;
        for (; it != std::filesystem::directory_iterator{}; it.increment(ec))
        {
            if (ec || out.size() >= 7u)
                return false;
            const auto path   = it->path();
            const DWORD attrs = GetFileAttributesW(path.c_str());
            if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
                return false;
            if (! permittedStage.empty() && path == permittedStage)
            {
                if ((attrs & FILE_ATTRIBUTE_DIRECTORY) != 0u || path.parent_path() != data / L"dst")
                    return false;
                continue;
            }
            Entry entry{};
            entry.relative  = path.lexically_relative(data);
            entry.directory = (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0u;
            const auto seed =
                std::ranges::find_if(expected, [&](const auto& value) noexcept { return value.first == entry.relative && value.second == entry.directory; });
            if (seed == expected.end())
                return false;
            if (! entry.directory)
            {
                uint64_t size = 0u;
                if (! GetFileSizeFsIo(io, path, size) || size > kSourceBytes || ! ReadFileTextFsIo(io, path, entry.bytes) || size != entry.bytes.size())
                    return false;
                Common::Crypto::Blake3Hasher hash;
                hash.Update(std::as_bytes(std::span(entry.bytes.data(), entry.bytes.size())));
                entry.digest = hash.Finalize();
            }
            out.emplace_back(std::move(entry));
        }
        if (ec)
            return false;
    }
    std::ranges::sort(out, [](const Entry& a, const Entry& b) noexcept { return a.relative.native() < b.relative.native(); });
    return out.size() == expected.size();
}
[[nodiscard]] bool SeedMatches(const std::vector<Entry>& entries) noexcept
{
    const auto get = [&](const std::filesystem::path& relative) noexcept -> const Entry*
    {
        const auto it = std::ranges::find(entries, relative, &Entry::relative);
        return it == entries.end() ? nullptr : &*it;
    };
    const auto* source   = get(L"src/payload.bin");
    const auto* old      = get(L"dst/payload.bin");
    const auto* sentinel = get(L"dst/sentinel.bin");
    return source && old && sentinel && source->bytes.size() == kSourceBytes &&
           std::ranges::all_of(source->bytes, [](char value) noexcept { return static_cast<unsigned char>(value) == 0xA7u; }) && old->bytes == kOldBytes &&
           sentinel->bytes == kSentinel;
}
struct Child final
{
    unsigned int phase = 0u;
    uint64_t taskId = 0u, subscription = 0u, deadline = 0u, parkedTick = 0u, promptCount = 0u;
    uint32_t heartbeats = 0u, decisions = 0u;
    uint64_t acceptedGeneration = 0u;
    HWND foreground             = nullptr;
    HWND main                   = nullptr;
    std::filesystem::path root, data, source, destination;
    std::vector<Entry> before;
    HostPromptResult priorOverride = HOST_PROMPT_RESULT_NONE;
    bool priorAccept               = false;
    bool failureRequested          = false;
    bool completionObserved        = false;
    HRESULT completionHr           = E_PENDING;
    bool strategyLogged            = false;
};
std::shared_ptr<Child> g_child;
[[nodiscard]] bool ChildTick(SelfTestState& state) noexcept
{
    using namespace OwnedShutdownSelfTest;
    using State = FolderWindow::FileOperationState;
    using Task  = State::Task;
    using namespace RedSalamander::TestSupport;
    if (! g_child)
        g_child = std::make_shared<Child>();
    const auto child = g_child;
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    const auto fail  = [&](std::wstring_view reason) noexcept
    {
        RecordFailure(reason);
        AppendLog(std::format(L"D10 owned child failure; preserving external tree: {}", reason));
        ReleaseStagePause();
        if (operations)
            operations->CancelAll();
        HostSetAutoAcceptPrompts(false);
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
        if (! child->failureRequested)
        {
            child->failureRequested = true;
            static_cast<void>(PostMessageW(state.mainWindow, WM_CLOSE, 0u, 0));
        }
        return false; // never ordinary suite cleanup of externally supplied evidence
    };
    if (! folder || ! operations || ! IsWindow(state.mainWindow))
        return fail(L"owned main/folder/operation owner unavailable");
    if (child->deadline != 0u && GetTickCount64() > child->deadline)
        return fail(L"owned close controller deadline");
    if (child->phase == 0u)
    {
        child->main           = state.mainWindow;
        child->root           = GetEnvironmentString(kCaseRoot);
        child->data           = child->root / L"data";
        child->source         = child->data / L"src" / L"payload.bin";
        child->destination    = child->data / L"dst" / L"payload.bin";
        const auto runId      = GetEnvironmentString(kTestRunIdEnvironmentVariable);
        const auto nonce      = GetEnvironmentString(kNonce);
        const auto parentText = Common::Strings::TryUtf8FromUtf16Strict(GetEnvironmentString(kParentPid));
        const auto parentPid  = parentText ? Common::Json::ParseIntegerString<DWORD>(*parentText) : std::nullopt;
        if (! parentPid || FAILED(ArmAuthority(state.mainWindow, SelfTest::GetSelfTestOptions().caseFilter, child->root, runId, nonce, *parentPid)))
            return fail(L"private exact-child close authority was refused");
        std::error_code ec;
        const auto base         = ResolveTestSandboxBase(ec);
        const auto childRun     = base / L"runs" / runId;
        const auto settingsPath = Common::Settings::GetSettingsPath(L"RedSalamander");
        if (ec || settingsPath.empty() || ! Common::Testing::IsSameOrDescendantTestSandboxPath(settingsPath, childRun))
            return fail(L"child settings escaped the isolated marked run");
        const auto sourceFs      = folder->GetFileSystem(FolderWindow::Pane::Left);
        const auto destinationFs = folder->GetFileSystem(FolderWindow::Pane::Right);
        wil::com_ptr<IFileSystemIO> io;
        if (! sourceFs || ! destinationFs || folder->GetFileSystemPluginId(FolderWindow::Pane::Left) != kPluginIdLocal ||
            folder->GetFileSystemPluginId(FolderWindow::Pane::Right) != kPluginIdLocal || ! sourceFs.try_query_to(io.put()) ||
            ! Capture(child->data, io, child->before) || ! SeedMatches(child->before))
            return fail(L"actual Local panes/external seed bytes/hash/inventory unavailable");
        child->foreground = GetForegroundWindow();
        child->deadline   = GetTickCount64() + SelfTest::ScaleTimeout(30'000u);
        operations->SetAutoDismissSuccess(false);
        operations->SetQueueNewTasks(false);
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = true; // captured normal verified Copy route, then normal teardown saves isolated settings
        if (FAILED(ArmStagePause(child->source, child->destination)))
            return fail(L"owned exact-source/destination stage pause arm failed");
        child->subscription    = folder->AddFileOperationCompletedCallback([child](const FolderWindow::FileOperationCompletedEvent& event) noexcept
        {
            if (event.taskId != child->taskId)
                return;
            const bool hasItem = event.itemOutcomes.size() == 1u;
            bool axes          = false;
            if (hasItem)
            {
                const auto& item = event.itemOutcomes.front();
                axes = item.sourceIndex == 0u && item.status == HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
                       item.publication == FileOperations::PublicationState::NotPublished &&
                       item.sourceDisposition == FileOperations::SourceDisposition::Retained && item.completion == FileOperations::ItemCompletion::Canceled &&
                       item.ownedStageDisposition == FileOperations::OwnedStageDisposition::Removed && item.finalSourcePath == child->source &&
                       item.finalDestinationPath == child->destination;
            }
            RecordCompletion(event.hr, axes);
            child->completionHr       = event.hr;
            child->completionObserved = true;
            if (auto* owner = TryGetFolderWindow(child->main))
                owner->RemoveFileOperationCompletedCallback(child->subscription);
            child->subscription = 0u;
        });
        const HRESULT admitted = operations->AdmitOperation(FILESYSTEM_COPY,
                                                            FolderWindow::Pane::Left,
                                                            FolderWindow::Pane::Right,
                                                            sourceFs,
                                                            {child->source},
                                                            child->destination.parent_path(),
                                                            FILESYSTEM_FLAG_NONE,
                                                            false,
                                                            0u,
                                                            State::ExecutionMode::PerItem,
                                                            false,
                                                            destinationFs,
                                                            &child->taskId);
        if (admitted != S_OK || child->taskId == 0u)
            return fail(std::format(L"actual verified Copy admission failed 0x{:08X}", static_cast<unsigned long>(admitted)));
        if (auto* task = operations->FindTask(child->taskId))
            task->_presentationState.store(Task::TaskPresentationState::Presented, std::memory_order_release);
        child->phase = 1u;
        return false; // all provider/IO borrows end before the next tick
    }
    if (child->phase == 1u)
    {
        auto* task        = operations->FindTask(child->taskId);
        const auto prompt = TryGetConflictPromptCopy(task);
        if (task && prompt && prompt->generation != child->acceptedGeneration)
        {
            if (child->decisions != 0u || prompt->bucket != Task::ConflictBucket::RegularFileExists ||
                ! PromptHasAction(*prompt, Task::ConflictAction::Overwrite) ||
                ! task->SubmitConflictDecision(prompt->generation, Task::ConflictAction::Overwrite, false))
                return fail(L"ordinary one-shot existing-occupant Overwrite consent mismatch");
            child->acceptedGeneration = prompt->generation;
            ++child->decisions;
        }
        task                = nullptr;
        const auto strategy = DebugGetPreparedTransferStrategyForSelfTest(child->taskId);
        if (strategy && ! child->strategyLogged)
        {
            AppendLog(std::format(L"D10 actual prepared strategy={} task={}", static_cast<unsigned int>(strategy.value()), child->taskId));
            child->strategyLogged = true;
        }
        const auto stage = GetStageProof();
        if (! stage.entered)
        {
            if (child->completionObserved || ! operations->HasActiveOperations())
            {
                AppendLog(std::format(L"D10 terminal before bridge stage cut: strategy={} stageBytes={} completion=0x{:08X}",
                                      strategy ? static_cast<int>(strategy.value()) : -1,
                                      stage.bytes,
                                      static_cast<unsigned long>(child->completionHr)));
                return fail(L"task completed before positive incomplete bridge stage write");
            }
            return false;
        }
        if (! strategy || strategy.value() != FileOperations::OperationStrategy::Copy)
            return fail(L"owned shutdown cut requires the actual verified bridge Copy strategy");
        if (stage.status != S_OK || stage.released || stage.returned || stage.bytes == 0u || stage.bytes >= kSourceBytes ||
            stage.source != child->source.native() || stage.destination != child->destination.native() ||
            std::filesystem::path(stage.stage).parent_path() != child->destination.parent_path() || child->decisions != 1u)
            return fail(L"positive incomplete owned-stage write proof mismatch");
        const auto sourceFs = folder->GetFileSystem(FolderWindow::Pane::Left);
        wil::com_ptr<IFileSystemIO> io;
        std::vector<Entry> parked;
        if (! sourceFs || ! sourceFs.try_query_to(io.put()) || ! Capture(child->data, io, parked, stage.stage) || parked != child->before)
            return fail(L"source/old occupant/sentinel changed at the owned-stage write cut");
        const auto source = std::ranges::find(child->before, std::filesystem::path(L"src/payload.bin"), &Entry::relative);
        Common::Crypto::Blake3Hasher expectedPrefix;
        expectedPrefix.Update(std::as_bytes(std::span(source->bytes.data(), static_cast<size_t>(stage.bytes))));
        if (expectedPrefix.Finalize() != stage.digest)
            return fail(L"actual persisted stage prefix differs from original source bytes");
        operations->DebugEnsurePopupVisibleForSelfTest();
        child->priorAccept   = HostGetAutoAcceptPrompts();
        child->priorOverride = HostGetTestPromptResultOverride();
        child->promptCount   = HostGetTestPromptRequestCount();
        HostSetAutoAcceptPrompts(false);
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
        child->parkedTick = GetTickCount64();
        child->phase      = 2u;
        if (! PostMessageW(state.mainWindow, WM_CLOSE, 0u, 0))
            return fail(L"actual first WM_CLOSE post failed");
        return false;
    }
    if (child->phase == 2u)
    {
        const auto stage = GetStageProof();
        FileOperationsPopupInternal::TaskSnapshot snapshot{};
        HostPromptDebugSnapshot exitPrompt{};
        if (! stage.entered || stage.released || stage.returned || ! operations->HasActiveOperations() || ! operations->FindTask(child->taskId) ||
            GetForegroundWindow() != child->foreground)
            return fail(L"main closed/task drained/stage released/foreground changed while parked");
        if (! DebugGetFileOperationsPopupTaskSnapshot(operations->GetPopupHwndForSelfTest(), child->taskId, snapshot) || ! snapshot.cancellationRequested ||
            snapshot.finished || snapshot.statusKind != FileOperationsPopupInternal::TaskSnapshot::StatusKind::Stopping)
            return false; // wait for actual production cancellation presentation, never fabricate it
        if (HostGetTestPromptRequestCount() != child->promptCount + 1u || ! HostGetTestLastPromptDebugSnapshot(exitPrompt) ||
            exitPrompt.buttons != HOST_PROMPT_BUTTONS_OK_CANCEL || exitPrompt.defaultResult != HOST_PROMPT_RESULT_CANCEL ||
            exitPrompt.title != LoadStringResource(nullptr, IDS_CAPTION_FILEOPS_EXIT) ||
            exitPrompt.message != LoadStringResource(nullptr, IDS_MSG_FILEOPS_CANCEL_ALL_EXIT))
            return fail(L"normal CancelAll close consent was bypassed or duplicated");
        ++child->heartbeats;
        const uint64_t elapsed = GetTickCount64() - child->parkedTick;
        if (child->heartbeats < 10u || elapsed < 300u)
            return false;
        RecordParkedProof(true, child->heartbeats, elapsed, true, true);
        HostSetAutoAcceptPrompts(child->priorAccept);
        if (child->priorOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(child->priorOverride);
        ReleaseStagePause();
        child->phase = 3u;
        // Completion and deferred WM_CLOSE now own teardown. The callback owns
        // child state; no stack, Task, FileOperationState or provider borrow survives.
        return false;
    }
    return false; // actual deferred WM_CLOSE/WM_DESTROY must end this child process
}

[[nodiscard]] bool ParentImpl(SelfTestState& state)
{
    using namespace OwnedShutdownSelfTest;
    using namespace RedSalamander::TestSupport;
    const auto reject = [&state](std::wstring_view reason) noexcept
    {
        state.retainTempRootOnFailure = true;
        AppendLog(std::format(L"D10 parent FAILED; preserve owned child evidence: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"parent quiet owner unavailable");
    const HWND foreground = GetForegroundWindow();
    const HWND parentMain = state.mainWindow;
    std::error_code ec;
    const auto base = ResolveTestSandboxBase(ec);
    if (ec || base.empty())
        return reject(L"marker-owned isolated run base unavailable");
    static uint64_t serial = 0u;
    // The native fallback run grammar ends in the real owner PID and tick.
    const auto runId    = std::format(L"d10-c{}-{}-{}", ++serial, GetCurrentProcessId(), GetTickCount64());
    const auto nonce    = runId;
    const auto childRun = base / L"runs" / runId;
    const auto root     = childRun / L"scratch" / L"owned-production-close";
    const auto data     = root / L"data";
    if (! Common::Testing::IsAuthorizedTestSandboxPath(childRun, {}, false, ec) || std::filesystem::exists(childRun, ec) || ec ||
        ! SelfTest::EnsureDirectory(data / L"src") || ! SelfTest::EnsureDirectory(data / L"dst"))
        return reject(L"fresh owned child run unavailable");
    wil::unique_handle childRunHandle(CreateFileW(childRun.c_str(),
                                                  FILE_READ_ATTRIBUTES,
                                                  FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                  nullptr,
                                                  OPEN_EXISTING,
                                                  FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                                                  nullptr));
    FILE_ID_INFO childRunId{};
    if (! childRunHandle || ! GetFileInformationByHandleEx(childRunHandle.get(), FileIdInfo, &childRunId, sizeof(childRunId)))
        return reject(L"exact externally owned child run identity unavailable");
    const auto archiveGroup = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::FileOperations, L"d10-owned-children");
    const auto archiveRun   = archiveGroup / runId;
    const auto parentRunId  = GetEnvironmentString(kTestRunIdEnvironmentVariable);
    if (! IsSafeTestRunId(parentRunId) || ! Common::Testing::IsSameOrDescendantTestSandboxPath(archiveGroup, base / L"runs" / parentRunId / L"artifacts") ||
        ! Common::Testing::IsAuthorizedTestSandboxPath(archiveRun, {}, false, ec))
        return reject(L"parent artifact archive escaped its exact governed run");
    // Negative authority control: the parent has no private child role and must
    // retain the ordinary selftest close policy even for the exact selected case.
    if (UsesProductionRoute(state.mainWindow) ||
        ArmAuthority(state.mainWindow, SelfTest::GetSelfTestOptions().caseFilter, root, runId, nonce, GetCurrentProcessId()) != E_ACCESSDENIED ||
        UsesProductionRoute(state.mainWindow))
        return reject(L"ordinary parent was able to arm child-only production close");
    const auto sourceFs = folder->GetFileSystem(FolderWindow::Pane::Left);
    wil::com_ptr<IFileSystemIO> io;
    if (! sourceFs || folder->GetFileSystemPluginId(FolderWindow::Pane::Left) != kPluginIdLocal || ! sourceFs.try_query_to(io.put()) ||
        ! WriteFilledTestFile(data / L"src/payload.bin", kSourceBytes, 0xA7u) || ! WriteFileTextFsIo(io, data / L"dst/payload.bin", std::string(kOldBytes)) ||
        ! WriteFileTextFsIo(io, data / L"dst/sentinel.bin", std::string(kSentinel)))
        return reject(L"externally owned distinct seed staging failed");
    std::vector<Entry> before;
    if (! Capture(data, io, before) || ! SeedMatches(before))
        return reject(L"exact external seed inventory/bytes/Blake3 unavailable");
    for (const auto& entry : before)
        if (! entry.directory)
            AppendLog(std::format(L"D10 external BEFORE path={} bytes={} BLAKE3={}",
                                  entry.relative.native(),
                                  entry.bytes.size(),
                                  Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(entry.digest))));
    wil::unique_cotaskmem_string executable;
    const HRESULT executableHr = wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, executable);
    if (FAILED(executableHr) || ! executable)
        return reject(L"owned executable path unavailable");
    struct Job final
    {
        std::atomic<bool> done{false};
        std::optional<ChildProcessResult> result;
    };
    const auto job = std::make_shared<Job>();
    std::stop_source cancel;
    ChildProcessOptions options{.executablePath       = executable.get(),
                                .arguments            = {L"--fileops-selftest", L"--selftest-no-activate", L"--selftest-case=FileOps_OwnedProductionShutdown"},
                                .timeout              = SelfTest::Scale(std::chrono::seconds(45)),
                                .cancellationToken    = cancel.get_token(),
                                .environmentOverrides = {{std::wstring(kTestRootEnvironmentVariable), base.native()},
                                                         {std::wstring(kTestRunIdEnvironmentVariable), runId},
                                                         {kRole, L"owned-child-v1"},
                                                         {kCaseRoot, root.native()},
                                                         {kNonce, nonce},
                                                         {kParentPid, std::to_wstring(GetCurrentProcessId())}}};
    // Parent environment/settings remain unchanged while its UI pumps. The owned
    // worker retains options/output storage and the runner's kill-on-close Job.
    std::jthread worker([job, options = std::move(options)]() noexcept
    {
        job->result = RunChildProcess(options);
        job->done.store(true, std::memory_order_release);
    });
    const auto drain  = wil::scope_exit([&]() noexcept
    {
        cancel.request_stop();
        if (worker.joinable())
            worker.join();
    });
    const auto exited = PumpMessagesUntil([job]() noexcept { return job->done.load(std::memory_order_acquire); },
                                          {.timeout = SelfTest::Scale(std::chrono::seconds(50)), .operationName = L"D10 owned child real process exit"});
    if (! exited.conditionMet)
        return reject(exited.timeoutDiagnostic);
    worker.join();
    if (! job->result.has_value())
        return reject(L"joined contained runner returned no result");
    const auto& result      = job->result.value();
    bool archiveAttempted   = false;
    bool archiveSucceeded   = false;
    const auto archiveChild = [&]() noexcept
    {
        if (archiveAttempted)
            return archiveSucceeded;
        archiveAttempted = true;
        // Completed is the contained runner's observed natural process exit.
        // Never move a timeout/killed/wait-failed or possibly live child's tree.
        if (! result.Completed() || ! Common::Testing::IsAuthorizedTestSandboxPath(childRun, {}, false, ec) ||
            ! Common::Testing::IsAuthorizedTestSandboxPath(archiveRun, {}, false, ec) || std::filesystem::exists(archiveRun, ec) || ec ||
            ! SelfTest::EnsureDirectory(archiveGroup))
            return false;
        wil::unique_handle current(CreateFileW(childRun.c_str(),
                                               FILE_READ_ATTRIBUTES,
                                               FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                               nullptr,
                                               OPEN_EXISTING,
                                               FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                                               nullptr));
        wil::unique_handle targetParent(CreateFileW(archiveGroup.c_str(),
                                                    FILE_READ_ATTRIBUTES,
                                                    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                    nullptr,
                                                    OPEN_EXISTING,
                                                    FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                                                    nullptr));
        FILE_ID_INFO currentId{}, targetId{};
        if (! current || ! targetParent || ! GetFileInformationByHandleEx(current.get(), FileIdInfo, &currentId, sizeof(currentId)) ||
            ! GetFileInformationByHandleEx(targetParent.get(), FileIdInfo, &targetId, sizeof(targetId)) ||
            currentId.VolumeSerialNumber != childRunId.VolumeSerialNumber || targetId.VolumeSerialNumber != childRunId.VolumeSerialNumber ||
            std::memcmp(&currentId.FileId, &childRunId.FileId, sizeof(currentId.FileId)) != 0 ||
            ! Common::Testing::IsExistingTestSandboxPathReparseFree(childRun, ec) || ! Common::Testing::IsExistingTestSandboxPathReparseFree(archiveGroup, ec))
            return false;
        // One same-volume rename retains every byte/receipt without traversing
        // descendants, and removes the sibling from stale-run GC/audit targets.
        std::filesystem::rename(childRun, archiveRun, ec);
        archiveSucceeded = ! ec;
        AppendLog(std::format(L"D10 child archive: exited=true original={} retained={} success={} error={}",
                              childRun.native(),
                              archiveRun.native(),
                              archiveSucceeded,
                              ec.value()));
        return archiveSucceeded;
    };
    const auto retainChild = wil::scope_exit([&]() noexcept
    {
        if (! archiveChild())
        {
            state.retainTempRootOnFailure = true;
            AppendLog(std::format(L"D10 child archive unavailable; retain original governed run for exact recovery: {}", childRun.native()));
        }
    });
    // WM_CLOSE can destroy a parent during the pump. Reacquire before any owner
    // dereference or further provider I/O; no old FolderWindow/state borrow wins.
    folder                     = IsWindow(parentMain) ? TryGetFolderWindow(parentMain) : nullptr;
    operations                 = TryGetFileOps(folder);
    const auto currentSourceFs = folder ? folder->GetFileSystem(FolderWindow::Pane::Left) : wil::com_ptr<IFileSystem>{};
    if (! folder || ! operations || ! currentSourceFs || ! currentSourceFs.try_query_to(io.put()))
        return reject(L"parent owner/provider disappeared while waiting for child exit");
    AppendLog(std::format(L"D10 child result pid={} status={} exit={} elapsedMs={} run={}",
                          result.processId,
                          static_cast<unsigned int>(result.status),
                          result.exitCode,
                          result.elapsed.count(),
                          childRun.native()));
    // Keep bounded launch/wait/output failure evidence outside the closed data
    // inventory, after the runner has ended and joined its output drains.
    const auto processReceipt = std::format("pid={}\nstatus={}\nexit={}\nwin32Error={}\nelapsedMs={}\nstdoutTruncated={}\nstderrTruncated={}\ndiagnostic={}\n",
                                            result.processId,
                                            static_cast<unsigned int>(result.status),
                                            result.exitCode,
                                            result.win32Error,
                                            result.elapsed.count(),
                                            result.stdoutTruncated,
                                            result.stderrTruncated,
                                            Common::Strings::Utf8FromUtf16ReplacingInvalid(result.diagnostic));
    if (! WriteFileTextFsIo(io, root / L"child-process.txt", processReceipt) || ! WriteFileTextFsIo(io, root / L"child-stdout.txt", result.stdoutBytes) ||
        ! WriteFileTextFsIo(io, root / L"child-stderr.txt", result.stderrBytes))
        return reject(L"bounded child process/output evidence retention failed");
    // Only actual root process completion can reach filesystem qualification.
    // Timed-out/killed/wait-failed children remain failures with scratch retained.
    if (! result.Completed() || result.exitCode != 0u || result.stdoutTruncated || result.stderrTruncated)
        return reject(result.diagnostic.empty() ? L"owned child exited without a clean natural result" : result.diagnostic);
    std::string receipt;
    if (! ReadFileTextFsIo(io, root / L"shutdown-receipt.txt", receipt) || receipt.size() > 64u * 1024u)
        return reject(L"normal close/destroy/quiet receipt missing");
    const auto line = [&](std::string_view key, std::string_view value) noexcept
    {
        const auto expected = std::format("{}={}\n", key, value);
        return receipt.starts_with(expected) || receipt.find("\n" + expected) != std::string::npos;
    };
    if (! line("schema", "1") || ! line("nonce", Common::Strings::Utf8FromUtf16ReplacingInvalid(nonce)) || ! line("pid", std::to_string(result.processId)) ||
        ! line("parentPid", std::to_string(GetCurrentProcessId())) || ! line("passed", "1") || ! line("normalTeardown", "1") ||
        ! line("quietBeforeDestroy", "1") || ! line("completion", "0x800704C7") || ! line("completionAxes", "1") || ! line("completionCount", "1") ||
        ! line("settingsSave", "0x00000000") || ! line("settingsSaveOwned", "1"))
        return reject(L"actual natural close receipt failed its identity/production-route/quiet axes");
    std::vector<Entry> after;
    if (! Capture(data, io, after) || after != before)
        return reject(L"source/old destination/sentinel/closed tree bytes or Blake3 changed after real child exit");
    for (const auto& entry : after)
        if (! entry.directory)
            AppendLog(std::format(L"D10 external AFTER pid={} processExited=true path={} bytes={} BLAKE3={}",
                                  result.processId,
                                  entry.relative.native(),
                                  entry.bytes.size(),
                                  Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(entry.digest))));
    if (GetForegroundWindow() != foreground || operations->HasActiveOperations())
        return reject(L"parent foreground/quiet point changed");
    if (! archiveChild())
        return reject(L"whole exited-child evidence could not join the parent governed artifact packet");
    AppendLog(std::format(L"D10 QUALIFIED actual WM_CLOSE/normal SaveAppSettings/WM_DESTROY/process exit; exact external tree retained: {}",
                          (archiveRun / L"scratch" / L"owned-production-close").native()));
    return true;
}
[[nodiscard]] bool Parent(SelfTestState& state) noexcept
{
    try
    {
        return ParentImpl(state);
    }
    catch (const std::system_error& error)
    {
        // Parent selftest boundary: failed runner-thread creation must preserve
        // the externally owned seed and report failure rather than terminate.
        AppendLog(std::format(L"D10 parent runner setup FAILED; evidence retained: {}", Utf16FromUtf8ForFileOpsSelfTest(error.what())));
        return false;
    }
}
} // namespace OwnedShutdownFixture
