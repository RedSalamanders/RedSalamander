// Included by the FileOps runner. Owned storage only; no activation, UIA or input.
[[nodiscard]] bool RunFileOpsDeleteParentPinChecks(SelfTestState& state) noexcept
{
    using Env         = RedSalamander::TestSupport::ScopedEnvironmentVariable;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Delete parent pin: {}", reason));
        return false;
    };
    if (! state.fsLocal || ! state.infoLocal)
        return reject(L"Local provider unavailable");
    std::string configuration;
    if (! BackupPluginConfiguration(state.infoLocal.get(), configuration))
        return reject(L"Local configuration backup failed");
    const auto restore = wil::scope_exit([&]() noexcept { static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), configuration)); });
    const auto fixture = state.tempRoot / L"delete-parent-pins";
    if (! RecreateEmptyDirectory(fixture))
        return reject(L"owned fixture unavailable");
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(fixture)); });
    wil::com_ptr<IFileSystemObjectBinding> binding;
    if (! state.fsLocal.try_query_to(binding.put()))
        return reject(L"Local binding unavailable");
    for (const unsigned int concurrency : {1u, 4u})
        for (const bool bound : {false, true})
            for (const bool cancel : {false, true})
            {
                if (! SetPluginConfiguration(state.infoLocal.get(), std::format("{{\"deleteMaxConcurrency\":{}}}", concurrency)))
                    return reject(L"delete concurrency configuration failed");
                const auto tree                 = fixture / L"tree";
                constexpr uint32_t nestedLevels = 32u;
                auto parent                     = tree / L"parent";
                for (uint32_t depth = 0u; depth < nestedLevels; ++depth)
                    parent /= L"d";
                const auto renamed  = fixture / L"renamed-parent";
                const auto sentinel = fixture / L"outside.bin";
                if (! WriteFilledTestFile(parent / L"a.bin", 17u, 0x31u) || ! WriteFilledTestFile(parent / L"b.bin", 29u, 0x92u) ||
                    ! WriteFilledTestFile(tree / L"sibling" / L"c.bin", 11u, 0x51u) || ! WriteFilledTestFile(sentinel, 31u, 0x77u))
                    return reject(L"owned tree staging failed");
                const std::wstring eventPrefix = std::format(L"Local\\RedSalamander.ParentPin.{}.{}", GetCurrentProcessId(), GetTickCount64());
                const std::wstring enteredName = eventPrefix + L".entered";
                const std::wstring releaseName = eventPrefix + L".release";
                wil::unique_handle entered(CreateEventW(nullptr, TRUE, FALSE, enteredName.c_str()));
                wil::unique_handle release(CreateEventW(nullptr, TRUE, FALSE, releaseName.c_str()));
                wil::unique_handle done(CreateEventW(nullptr, TRUE, FALSE, nullptr));
                if (! entered || ! release || ! done)
                    return reject(L"owned events unavailable");
                const std::wstring extendedParent  = Common::Paths::ToExtendedWin32Path(parent.native());
                const std::wstring extendedRenamed = Common::Paths::ToExtendedWin32Path(renamed.native());
                Env parentEnv(L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_PARENT", extendedParent);
                Env enteredEnv(L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_ENTERED_EVENT", enteredName);
                Env releaseEnv(L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_RELEASE_EVENT", releaseName);
                Env removedEnv(L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_REMOVED_PATH", std::nullopt);
                auto task          = std::make_unique<FolderWindow::FileOperationState::Task>(*state.fileOps);
                task->_operation   = FILESYSTEM_DELETE;
                task->_sourcePaths = {tree};
                task->BeginDiscovery();
                task->SkipDiscovery();
                FileSystemOptions options{};
                task->InitializeFileSystemOptions(options, nullptr);
                wil::com_ptr<IFileSystemBoundObject> retained;
                if (bound && FAILED(binding->BindObject(
                                 tree.c_str(),
                                 static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_DELETE),
                                 retained.put())))
                    return reject(L"selected-root binding failed");
                std::atomic<HRESULT> result{E_PENDING};
                const auto started = std::chrono::steady_clock::now();
                std::jthread worker([&]() noexcept
                {
                    const auto initialized = wil::CoInitializeEx(COINIT_MULTITHREADED);
                    static_cast<void>(initialized);
                    task->_dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
                    const auto callback = wil::scope_exit([&]() noexcept { task->_dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
                    FileSystemConditionalMutationResult receipt{};
                    receipt.sizeBytes = sizeof(receipt);
                    const HRESULT hr  = bound ? retained->DeleteIfUnchanged(FILESYSTEM_FLAG_RECURSIVE, &options, &receipt)
                                              : state.fsLocal->DeleteItem(tree.c_str(), FILESYSTEM_FLAG_RECURSIVE, &options, nullptr, nullptr);
                    result.store(hr, std::memory_order_release);
                    static_cast<void>(SetEvent(done.get()));
                });
                const auto ensureRelease = wil::scope_exit([&]() noexcept { static_cast<void>(SetEvent(release.get())); });
                if (WaitForSingleObject(entered.get(), 10'000u) != WAIT_OBJECT_0)
                    return reject(L"walk did not reach the post-listing committed-child pause");
                const bool renamedDuringWalk = MoveFileExW(extendedParent.c_str(), extendedRenamed.c_str(), 0u) != FALSE;
                const DWORD renameError      = renamedDuringWalk ? ERROR_SUCCESS : GetLastError();
                wil::unique_handle writer(CreateFileW(parent.c_str(),
                                                      GENERIC_WRITE,
                                                      FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                      nullptr,
                                                      OPEN_EXISTING,
                                                      FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                                      nullptr));
                if (renamedDuringWalk || (renameError != ERROR_SHARING_VIOLATION && renameError != ERROR_ACCESS_DENIED) || writer)
                    return reject(L"a traversed parent could be renamed or reopened for in-place write");
                if (cancel)
                    task->RequestCancel();
                static_cast<void>(SetEvent(release.get()));
                if (WaitForSingleObject(done.get(), 10'000u) != WAIT_OBJECT_0)
                    return reject(L"delete did not reach callback quiescence after release");
                worker.join();
                const HRESULT hr = result.load(std::memory_order_acquire);
                retained.reset();
                if ((! cancel && hr != S_OK) || (cancel && hr != HRESULT_FROM_WIN32(ERROR_CANCELLED)))
                    return reject(std::format(L"unexpected delete result 0x{:08X}", static_cast<unsigned long>(hr)));
                if (cancel)
                {
                    // Scanner/notification cleanup can lag callback quiescence briefly. Retry only
                    // sharing/access contention, bounded by the canonical pump, and retain the
                    // actual failure instead of diagnosing every rename failure as a leaked pin.
                    const auto renameAfterJoin = [&](const std::wstring& from, const std::wstring& to) noexcept
                    {
                        DWORD error           = ERROR_SUCCESS;
                        bool renamedAfterJoin = false;
                        const auto quiet      = RedSalamander::TestSupport::PumpMessagesUntil(
                            [&]() noexcept
                        {
                            renamedAfterJoin = MoveFileExW(from.c_str(), to.c_str(), 0u) != FALSE;
                            error            = renamedAfterJoin ? ERROR_SUCCESS : GetLastError();
                            return renamedAfterJoin || (error != ERROR_SHARING_VIOLATION && error != ERROR_ACCESS_DENIED);
                        },
                            {.timeout = SelfTest::Scale(std::chrono::seconds(2)), .operationName = L"released delete-parent pin"});
                        if (! quiet.conditionMet || ! renamedAfterJoin)
                            return reject(std::format(
                                L"post-join rename failed: bound={} configuredConcurrency={} error={} from=[{}] to=[{}]", bound, concurrency, error, from, to));
                        return true;
                    };
                    // A parallel child may have committed its directory removal before the
                    // root observes cancellation. Probe the fixture ancestor, which survives
                    // either partial outcome and cannot move while a descendant pin is live.
                    const std::wstring fixturePath  = Common::Paths::ToExtendedWin32Path(fixture.native());
                    const std::wstring releasedPath = Common::Paths::ToExtendedWin32Path((state.tempRoot / L"delete-parent-pins-released").native());
                    if (! renameAfterJoin(fixturePath, releasedPath) || ! renameAfterJoin(releasedPath, fixturePath))
                        return false;
                }
                else if (GetFileAttributesW(tree.c_str()) != INVALID_FILE_ATTRIBUTES)
                    return reject(L"ordinary recursive deletion retained its root");
                std::string sentinelBytes;
                if (! ReadFileTextFsIo(state.fsLocal.query<IFileSystemIO>(), sentinel, sentinelBytes) ||
                    sentinelBytes != std::string(31u, static_cast<char>(0x77)))
                    return reject(L"the outside-tree sentinel changed");
                Debug::Perf::Emit(
                    L"FileOps.SelfTest.DeleteParentPin",
                    std::format(L"bound={} configuredConcurrency={} effectiveConcurrency={} cancel={} nestedLevels={}",
                                bound,
                                concurrency,
                                bound ? 1u : concurrency,
                                cancel,
                                nestedLevels),
                    static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count()),
                    bound ? 1u : concurrency,
                    nestedLevels,
                    hr);
                if (! SelfTest::RemoveAll(tree))
                    return reject(L"case storage cleanup failed");
            }
    return true;
}

[[nodiscard]] bool RunFileOpsCrossRootSkipContractChecks(SelfTestState& state) noexcept
{
    const auto unreachable = RunFileOpsCrossRootLocalCopyChecks(state, LoopbackShareStatus::Unreachable);
    const auto refused     = RunFileOpsCrossRootLocalCopyChecks(state, LoopbackShareStatus::Refused);
    return unreachable.status == SelfTest::SelfTestCaseResult::Status::skipped && ! unreachable.reason.empty() &&
           refused.status == SelfTest::SelfTestCaseResult::Status::failed && ! refused.reason.empty();
}
