case SelfTestState::Step::FileOps_CrossVolumeMovePartialFailureStatus:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"FileOps_CrossVolumeMovePartialFailureStatus timed out.");
        return true;
    }

    const std::filesystem::path srcRoot  = state.tempRoot / L"phase4-partial-move-source";
    const std::filesystem::path dstRoot  = state.tempRoot / L"phase4-partial-move-destination";
    const std::filesystem::path srcFile  = srcRoot / L"locked.bin";
    const std::filesystem::path dstFile  = dstRoot / srcRoot.filename() / srcFile.filename();
    const std::filesystem::path dstItem  = dstRoot / srcRoot.filename();

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(dstItem) ||
            ! WriteTestFile(srcFile, 256ull * 1024ull))
        {
            Fail(L"Partial move status test failed to prepare source/destination folders.");
            return true;
        }
        if (! SetFileAttributesW(srcFile.c_str(), FILE_ATTRIBUTE_READONLY))
        {
            Fail(L"Partial move status test failed to make the source file read-only.");
            return true;
        }

        if (! state.forceMoveCopyFallbackEnvBackedUp)
        {
            state.forceMoveCopyFallbackEnvOriginal    = GetEnvVarTrimmed(kSelfTestEnvForceMoveCopyFallback);
            state.forceMoveCopyFallbackEnvHadOriginal = ! state.forceMoveCopyFallbackEnvOriginal.empty();
            state.forceMoveCopyFallbackEnvBackedUp    = true;
        }
        if (! SetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(), L"1"))
        {
            Fail(L"Partial move status test failed to enable forced move-copy fallback.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        state.taskA                 = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_MOVE, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {srcRoot}, dstRoot, flags, false);
        if (! state.taskA.has_value())
        {
            Fail(L"Partial move status test failed to start the move task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (state.forceMoveCopyFallbackEnvBackedUp)
        {
            static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(),
                                                      state.forceMoveCopyFallbackEnvHadOriginal ? state.forceMoveCopyFallbackEnvOriginal.c_str() : nullptr));
            state.forceMoveCopyFallbackEnvBackedUp    = false;
            state.forceMoveCopyFallbackEnvHadOriginal = false;
            state.forceMoveCopyFallbackEnvOriginal.clear();
        }

        const HRESULT expectedPartial = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (completed->second.hr != expectedPartial)
        {
            Fail(std::format(L"Partial move status expected ERROR_PARTIAL_COPY, got 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        std::error_code ec;
        if (! std::filesystem::exists(srcFile, ec) || ec)
        {
            Fail(L"Partial move status test did not preserve the locked source file.");
            return true;
        }
        ec.clear();
        if (! std::filesystem::exists(dstFile, ec) || ec)
        {
            Fail(L"Partial move status test did not leave the locked destination copy.");
            return true;
        }

        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        const auto summaryIt = std::find_if(summaries.begin(), summaries.end(), [&](const auto& summary) noexcept {
            return state.taskA.has_value() && summary.taskId == state.taskA.value();
        });
        if (summaryIt == summaries.end())
        {
            Fail(L"Partial move status test could not find the completed task summary.");
            return true;
        }

        bool foundExplicitIssue = false;
        for (const auto& issue : summaryIt->issueDiagnostics)
        {
            if (issue.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Warning &&
                ContainsIgnoreCase(issue.message, L"source preserved") && ContainsIgnoreCase(issue.message, L"partial copy left"))
            {
                foundExplicitIssue = true;
                break;
            }
        }
        if (! foundExplicitIssue)
        {
            Fail(L"Partial move status test did not find a 'source preserved / partial copy left' issue row.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.CrossVolumeMovePartialFailureStatus",
                          L"shape=debug-forced-copy-delete-rollback-failure",
                          completed->second.progressCompletedBytes,
                          summaryIt->issueDiagnostics.size(),
                          1,
                          completed->second.hr);

        static_cast<void>(SetFileAttributesW(srcFile.c_str(), FILE_ATTRIBUTE_NORMAL));
        static_cast<void>(SetFileAttributesW(dstFile.c_str(), FILE_ATTRIBUTE_NORMAL));
        NextStep(state, SelfTestState::Step::FileOps_ReparseRetargetDestinationContainment);
        return false;
    }

    return false;
}

case SelfTestState::Step::FileOps_ReparseRetargetDestinationContainment:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        Fail(L"FileOps_ReparseRetargetDestinationContainment timed out.");
        return true;
    }

    static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"copyReparse"})json"));

    const std::filesystem::path srcRoot       = state.tempRoot / L"phase4-retarget-src";
    const std::filesystem::path dstRoot       = state.tempRoot / L"phase4-retarget-dst";
    const std::filesystem::path insideTarget  = srcRoot / L"inside-target";
    const std::filesystem::path outsideTarget = state.tempRoot / L"phase4-retarget-outside";
    const std::filesystem::path insideLink    = srcRoot / L"inside-link";
    const std::filesystem::path outsideLink   = srcRoot / L"outside-link";

    if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(insideTarget) ||
        ! RecreateEmptyDirectory(outsideTarget))
    {
        Fail(L"Reparse containment test failed to prepare directories.");
        return true;
    }
    if (! WriteTestFile(insideTarget / L"inside.bin", 64) || ! WriteTestFile(outsideTarget / L"outside.bin", 64))
    {
        Fail(L"Reparse containment test failed to seed target files.");
        return true;
    }
    if (! TryCreateJunction(insideLink, insideTarget) || ! TryCreateJunction(outsideLink, outsideTarget))
    {
        Fail(L"Reparse containment test failed to create source junctions.");
        return true;
    }

    const std::filesystem::path copiedRoot = dstRoot / srcRoot.filename();
    const HRESULT copyHr =
        state.fsLocal->CopyItem(srcRoot.c_str(), copiedRoot.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), nullptr, nullptr, nullptr);
    if (FAILED(copyHr))
    {
        Fail(std::format(L"Reparse containment copy failed: 0x{:08X}.", static_cast<unsigned long>(copyHr)));
        return true;
    }

    const auto copiedInsideTarget = TryGetDirectoryReparseTargetAbsolute(copiedRoot / insideLink.filename());
    if (! copiedInsideTarget.has_value())
    {
        Fail(L"Reparse containment test could not read copied in-tree link target.");
        return true;
    }

    const std::wstring expectedInsideTarget = NormalizePathForCompare((copiedRoot / insideTarget.filename()).wstring());
    if (copiedInsideTarget.value() != expectedInsideTarget)
    {
        Fail(std::format(L"Copied in-tree reparse target escaped or was not retargeted. expected='{}' actual='{}'.",
                         expectedInsideTarget,
                         copiedInsideTarget.value()));
        return true;
    }

    const auto copiedOutsideTarget = TryGetDirectoryReparseTargetAbsolute(copiedRoot / outsideLink.filename());
    if (! copiedOutsideTarget.has_value())
    {
        Fail(L"Reparse containment test could not read copied out-of-tree link target.");
        return true;
    }

    const std::wstring expectedOutsideTarget = NormalizePathForCompare(std::filesystem::absolute(outsideTarget).wstring());
    if (copiedOutsideTarget.value() != expectedOutsideTarget)
    {
        Fail(std::format(L"Copied out-of-tree reparse target should not be retargeted. expected='{}' actual='{}'.",
                         expectedOutsideTarget,
                         copiedOutsideTarget.value()));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.ReparseRetargetDestinationContainment",
                      L"shape=in-tree-retarget-plus-out-of-tree-preserve",
                      2,
                      CountFilesRecursive(copiedRoot),
                      1,
                      S_OK);

    NextStep(state, SelfTestState::Step::FileOps_DeleteToctouSwapGuard);
    return false;
}

case SelfTestState::Step::FileOps_DeleteToctouSwapGuard:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        Fail(L"FileOps_DeleteToctouSwapGuard timed out.");
        return true;
    }

    static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","deleteMaxConcurrency":8})json"));

    const std::filesystem::path deleteRoot = state.tempRoot / L"phase4-toctou-delete-root";
    const std::filesystem::path victimDir  = deleteRoot / L"victim-dir";
    const std::filesystem::path targetRoot = state.tempRoot / L"phase4-toctou-outside-target";
    const std::filesystem::path targetFile = targetRoot / L"must-survive.bin";
    const std::filesystem::path sibling    = deleteRoot / L"sibling.bin";

    if (! RecreateEmptyDirectory(deleteRoot) || ! RecreateEmptyDirectory(victimDir) || ! RecreateEmptyDirectory(targetRoot))
    {
        Fail(L"Delete TOCTOU guard test failed to prepare directories.");
        return true;
    }
    if (! WriteTestFile(targetFile, 128) || ! WriteTestFile(sibling, 128))
    {
        Fail(L"Delete TOCTOU guard test failed to seed files.");
        return true;
    }

    if (! SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapPath.data(), victimDir.c_str()) ||
        ! SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapTarget.data(), targetRoot.c_str()))
    {
        Fail(L"Delete TOCTOU guard test failed to arm the debug swap hook.");
        return true;
    }

    const HRESULT deleteHr =
        state.fsLocal->DeleteItem(deleteRoot.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), nullptr, nullptr, nullptr);
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapTarget.data(), nullptr));
    if (FAILED(deleteHr))
    {
        Fail(std::format(L"Delete TOCTOU guard test delete failed: 0x{:08X}.", static_cast<unsigned long>(deleteHr)));
        return true;
    }

    std::error_code ec;
    if (std::filesystem::exists(deleteRoot, ec) || ec)
    {
        Fail(L"Delete TOCTOU guard test did not remove the requested root.");
        return true;
    }

    ec.clear();
    if (! std::filesystem::exists(targetFile, ec) || ec)
    {
        Fail(L"Delete TOCTOU guard removed an out-of-tree target through a swapped junction.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.DeleteToctouSwapGuard",
                      L"shape=dir-to-reparse-swap-during-parallel-flatten",
                      1,
                      CountFilesRecursive(targetRoot),
                      1,
                      S_OK);

    NextStep(state, SelfTestState::Step::Phase5_PreCalcSettingsApplied);
    return false;
}

case SelfTestState::Step::Phase10_PermanentDelete:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Phase10_PermanentDelete timed out.");
        return true;
    }

    const std::filesystem::path delDir     = state.tempRoot / L"perm-delete";
    const std::filesystem::path delFile    = delDir / L"perm.bin";
    const std::filesystem::path cancelFile = delDir / L"cancelled-perm.bin";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(delDir))
        {
            Fail(L"Failed to reset perm-delete directory.");
            return true;
        }

        if (! WriteTestFile(cancelFile, 1024))
        {
            Fail(L"Failed to write permanent-delete cancellation test file.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        HostResetTestPromptRequestCount();
        {
            HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
            const auto clearPromptOverride                   = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });
            const std::optional<std::uint64_t> cancelledTask = StartFileOperationAndGetId(state.fileOps,
                                                                                          FILESYSTEM_DELETE,
                                                                                          FolderWindow::Pane::Left,
                                                                                          std::nullopt,
                                                                                          state.fsLocal,
                                                                                          {cancelFile},
                                                                                          {},
                                                                                          flags,
                                                                                          false,
                                                                                          0,
                                                                                          FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                                          false);
            if (cancelledTask.has_value())
            {
                Fail(L"Permanent delete without Recycle Bin started even though its confirmation prompt was cancelled.");
                return true;
            }
        }

        if (HostGetTestPromptRequestCount() != 1u)
        {
            Fail(std::format(L"Permanent delete cancellation expected one confirmation prompt, got {}.", HostGetTestPromptRequestCount()));
            return true;
        }

        std::error_code cancelExistsEc;
        if (! std::filesystem::exists(cancelFile, cancelExistsEc) || cancelExistsEc)
        {
            Fail(L"Cancelled permanent delete removed the source file.");
            return true;
        }

        if (! WriteTestFile(delFile, 4096))
        {
            Fail(L"Failed to write perm-delete test file.");
            return true;
        }

        HostResetTestPromptRequestCount();
        {
            HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
            const auto clearPromptOverride = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });
            state.taskA                    = StartFileOperationAndGetId(state.fileOps,
                                                                        FILESYSTEM_DELETE,
                                                                        FolderWindow::Pane::Left,
                                                                        std::nullopt,
                                                                        state.fsLocal,
                                                                        {delFile},
                                                                        {},
                                                                        flags,
                                                                        false,
                                                                        0,
                                                                        FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                        false);
        }
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start confirmed permanent-delete task.");
            return true;
        }

        if (HostGetTestPromptRequestCount() != 1u)
        {
            Fail(std::format(L"Confirmed permanent delete expected one confirmation prompt, got {}.", HostGetTestPromptRequestCount()));
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        auto* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (task && (task->_flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0)
        {
            Fail(L"Permanent delete task unexpectedly used Recycle Bin flag.");
            return true;
        }

        state.stepState = 2;
    }

    const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (it == state.completedTasks.end())
    {
        return false;
    }

    if (FAILED(it->second.hr))
    {
        Fail(std::format(L"Permanent delete task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
        return true;
    }

    std::error_code ec;
    if (std::filesystem::exists(delFile, ec))
    {
        Fail(L"Permanent delete task did not remove the source file.");
        return true;
    }

    // Validate file-root pre-calc contract on local filesystem (S_OK + fileCount=1).
    const std::filesystem::path localSizeFile = state.tempRoot / L"size-root-file.bin";
    constexpr uint64_t kLocalSizeBytes        = 12'345ull;
    if (! WriteTestFile(localSizeFile, kLocalSizeBytes))
    {
        Fail(L"Failed to create local size-root file.");
        return true;
    }

    wil::com_ptr<IFileSystemDirectoryOperations> localDirOps;
    if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localDirOps.addressof()))) || ! localDirOps)
    {
        Fail(L"Local filesystem does not expose IFileSystemDirectoryOperations.");
        return true;
    }

    FileSystemDirectorySizeResult localSizeResult{};
    localSizeResult.sizeBytes = sizeof(FileSystemDirectorySizeResult);
    const HRESULT localSizeHr = localDirOps->GetDirectorySize(localSizeFile.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, &localSizeResult);
    if (FAILED(localSizeHr) || FAILED(localSizeResult.status))
    {
        Fail(std::format(L"Local file-root GetDirectorySize failed: hr=0x{:08X} status=0x{:08X}.",
                         static_cast<unsigned long>(localSizeHr),
                         static_cast<unsigned long>(localSizeResult.status)));
        return true;
    }

    if (localSizeResult.totalBytes != kLocalSizeBytes || localSizeResult.fileCount != 1ull || localSizeResult.directoryCount != 0ull)
    {
        Fail(std::format(L"Local file-root GetDirectorySize mismatch: bytes={} files={} dirs={}.",
                         localSizeResult.totalBytes,
                         localSizeResult.fileCount,
                         localSizeResult.directoryCount));
        return true;
    }

    // Validate file-root pre-calc contract on dummy filesystem (S_OK + fileCount=1).
    wil::com_ptr<IFileSystemDirectoryOperations> dummyDirOps;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyDirOps.addressof()))) || ! dummyDirOps)
    {
        Fail(L"Dummy filesystem does not expose IFileSystemDirectoryOperations.");
        return true;
    }

    const std::wstring dummyFolder = (! state.dummyPaths.empty()) ? state.dummyPaths.front() : L"/";
    wil::com_ptr<IFilesInformation> dummyInfo;
    if (FAILED(state.fsDummy->ReadDirectoryInfo(dummyFolder.c_str(), dummyInfo.addressof())) || ! dummyInfo)
    {
        Fail(L"Failed to enumerate dummy folder for file-root size test.");
        return true;
    }

    FileInfo* dummyEntry          = nullptr;
    unsigned long dummyBufferSize = 0;
    if (FAILED(dummyInfo->GetBuffer(&dummyEntry)) || FAILED(dummyInfo->GetBufferSize(&dummyBufferSize)) || dummyEntry == nullptr ||
        dummyBufferSize < sizeof(FileInfo))
    {
        Fail(L"Dummy folder enumeration returned no entries for file-root size test.");
        return true;
    }

    std::wstring dummyFilePath;
    {
        const std::byte* base = reinterpret_cast<const std::byte*>(dummyEntry);
        const std::byte* end  = base + dummyBufferSize;
        const FileInfo* cur   = dummyEntry;

        while (cur != nullptr)
        {
            if ((cur->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
            {
                const size_t nameChars = static_cast<size_t>(cur->FileNameSize) / sizeof(wchar_t);
                std::wstring_view name(cur->FileName, nameChars);

                dummyFilePath = dummyFolder;
                if (dummyFilePath.empty())
                {
                    dummyFilePath = L"/";
                }
                if (! dummyFilePath.empty() && dummyFilePath.back() != L'/' && dummyFilePath.back() != L'\\')
                {
                    dummyFilePath.push_back(L'/');
                }
                dummyFilePath.append(name);
                break;
            }

            if (cur->NextEntryOffset == 0)
            {
                break;
            }

            if (cur->NextEntryOffset < sizeof(FileInfo))
            {
                break;
            }

            const std::byte* next = reinterpret_cast<const std::byte*>(cur) + cur->NextEntryOffset;
            if (next < base || next + sizeof(FileInfo) > end)
            {
                break;
            }

            cur = reinterpret_cast<const FileInfo*>(next);
        }
    }

    if (dummyFilePath.empty())
    {
        Fail(L"Dummy folder did not provide a file entry for file-root size test.");
        return true;
    }

    FileSystemDirectorySizeResult dummySizeResult{};
    dummySizeResult.sizeBytes = sizeof(FileSystemDirectorySizeResult);
    const HRESULT dummySizeHr = dummyDirOps->GetDirectorySize(dummyFilePath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, &dummySizeResult);
    if (FAILED(dummySizeHr) || FAILED(dummySizeResult.status))
    {
        Fail(std::format(L"Dummy file-root GetDirectorySize failed: path={} hr=0x{:08X} status=0x{:08X}.",
                         dummyFilePath,
                         static_cast<unsigned long>(dummySizeHr),
                         static_cast<unsigned long>(dummySizeResult.status)));
        return true;
    }

    if (dummySizeResult.fileCount != 1ull || dummySizeResult.directoryCount != 0ull)
    {
        Fail(std::format(L"Dummy file-root GetDirectorySize mismatch: bytes={} files={} dirs={}.",
                         dummySizeResult.totalBytes,
                         dummySizeResult.fileCount,
                         dummySizeResult.directoryCount));
        return true;
    }

    // Validate recycle-bin delete failure returns specific per-item error (not generic E_FAIL).
    const std::filesystem::path recycleLocked = state.tempRoot / std::format(L"recyclebin-locked-{}.bin", GetTickCount64());
    if (! WriteTestFile(recycleLocked, 1024))
    {
        Fail(std::format(L"Failed to create recycle-bin locked test file (err={}).", GetLastError()));
        return true;
    }

    wil::unique_handle lockHandle(CreateFileW(recycleLocked.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! lockHandle)
    {
        Fail(L"Failed to open recycle-bin locked test file handle.");
        return true;
    }

    const HRESULT recycleHr = state.fsLocal->DeleteItem(recycleLocked.c_str(), FILESYSTEM_FLAG_USE_RECYCLE_BIN, nullptr, nullptr, nullptr);
    if (SUCCEEDED(recycleHr))
    {
        Fail(L"Recycle-bin locked-file delete unexpectedly succeeded.");
        return true;
    }

    if (recycleHr == E_FAIL || recycleHr == E_UNEXPECTED || recycleHr == HRESULT_FROM_WIN32(ERROR_GEN_FAILURE))
    {
        Fail(std::format(L"Recycle-bin locked-file delete returned generic HRESULT: 0x{:08X}.", static_cast<unsigned long>(recycleHr)));
        return true;
    }

    lockHandle.reset();
    ec.clear();
    if (! std::filesystem::exists(recycleLocked, ec))
    {
        Fail(L"Recycle-bin locked-file test unexpectedly removed the source file.");
        return true;
    }

    ec.clear();
    static_cast<void>(std::filesystem::remove(recycleLocked, ec));

    NextStep(state, SelfTestState::Step::Phase11_CrossFileSystemBridge);
    return false;
}
case SelfTestState::Step::Phase11_CrossFileSystemBridge:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Phase11_CrossFileSystemBridge timed out.");
        return true;
    }

    const std::filesystem::path srcDir   = state.tempRoot / L"bridge-src";
    const std::filesystem::path dstDir   = state.tempRoot / L"bridge-roundtrip";
    const std::filesystem::path moveDir  = state.tempRoot / L"bridge-move-src";
    const std::filesystem::path moveFile = moveDir / L"move.bin";

    const std::wstring dummyCopyRoot                = L"/bridge-copy";
    const std::wstring dummyMoveRoot                = L"/bridge-move";
    const std::wstring dummyCancelRoot              = L"/bridge-cancel";
    constexpr size_t kBridgeCancelFileBytes         = 8ull * 1024ull * 1024ull;
    constexpr uint64_t kBridgeCancelSpeedLimitBytes = 1ull * 1024ull * 1024ull;
    constexpr size_t kBridgeConcurrencyFileBytes    = 2ull * 1024ull * 1024ull;
    constexpr uint64_t kBridgeConcurrencySpeedLimit = 1ull * 1024ull * 1024ull;
    constexpr int kBridgeConcurrencyFileCount       = 4;

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir) || ! RecreateEmptyDirectory(moveDir))
        {
            Fail(L"Failed to reset bridge test directories.");
            return true;
        }

        std::error_code ec;
        std::filesystem::create_directories(srcDir / L"sub", ec);
        if (ec)
        {
            Fail(L"Failed to create bridge-src directory structure.");
            return true;
        }

        if (! WriteTestFile(srcDir / L"a.bin", 128) || ! WriteTestFile(srcDir / L"sub" / L"b.bin", 4096))
        {
            Fail(L"Failed to write bridge-src test files.");
            return true;
        }

        if (! WriteTestFile(moveFile, 2048))
        {
            Fail(L"Failed to write bridge-move-src test file.");
            return true;
        }

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyCopyRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyMoveRoot))
        {
            Fail(L"Failed to create dummy folders for cross-filesystem bridge tests.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_COPY,
                                                                 FolderWindow::Pane::Left,
                                                                 FolderWindow::Pane::Right,
                                                                 state.fsLocal,
                                                                 {srcDir},
                                                                 std::filesystem::path(dummyCopyRoot),
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                 false,
                                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start cross-filesystem copy (local -> dummy).");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }
        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Cross-filesystem copy (local -> dummy) failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        const std::filesystem::path dummySource = std::filesystem::path(dummyCopyRoot) / L"bridge-src";

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskB                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_COPY,
                                                                 FolderWindow::Pane::Right,
                                                                 FolderWindow::Pane::Left,
                                                                 state.fsDummy,
                                                                 {dummySource},
                                                                 dstDir,
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                 false,
                                                                 state.fsLocal);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start cross-filesystem copy (dummy -> local).");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto it = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }
        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Cross-filesystem copy (dummy -> local) failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        std::error_code ec;
        const std::filesystem::path outRoot = dstDir / L"bridge-src";
        const auto aSize                    = std::filesystem::file_size(outRoot / L"a.bin", ec);
        if (ec || aSize != 128)
        {
            Fail(L"Cross-filesystem roundtrip: a.bin missing or wrong size.");
            return true;
        }
        ec.clear();
        const auto bSize = std::filesystem::file_size(outRoot / L"sub" / L"b.bin", ec);
        if (ec || bSize != 4096)
        {
            Fail(L"Cross-filesystem roundtrip: b.bin missing or wrong size.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_NONE);
        state.taskC                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_MOVE,
                                                                 FolderWindow::Pane::Left,
                                                                 FolderWindow::Pane::Right,
                                                                 state.fsLocal,
                                                                 {moveFile},
                                                                 std::filesystem::path(dummyMoveRoot),
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                 false,
                                                                 state.fsDummy);
        if (! state.taskC.has_value())
        {
            Fail(L"Failed to start cross-filesystem move (local -> dummy).");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    using Task = FolderWindow::FileOperationState::Task;

    if (state.stepState == 3)
    {
        const auto it = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Cross-filesystem move (local -> dummy) failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        std::error_code ec;
        if (std::filesystem::exists(moveFile, ec))
        {
            Fail(L"Cross-filesystem move did not remove the source file.");
            return true;
        }

        const std::filesystem::path overwriteFile = srcDir / L"a.bin";
        if (! WriteTestFile(overwriteFile, 512))
        {
            Fail(L"Failed to update a.bin for overwrite prompt test.");
            return true;
        }

        const std::wstring dummyOverwriteFolder = std::wstring(dummyCopyRoot) + L"/bridge-src";

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_NONE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_COPY,
                                                                 FolderWindow::Pane::Left,
                                                                 FolderWindow::Pane::Right,
                                                                 state.fsLocal,
                                                                 {overwriteFile},
                                                                 std::filesystem::path(dummyOverwriteFolder),
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                 false,
                                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start overwrite prompt test copy (local -> dummy).");
            return true;
        }

        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        Task* task        = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }

        if (prompt->bucket != Task::ConflictBucket::Exists)
        {
            Fail(L"Cross-filesystem overwrite test did not produce an Exists prompt.");
            return true;
        }

        if (! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
        {
            Fail(L"Cross-filesystem overwrite test prompt did not offer Overwrite.");
            return true;
        }

        task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Cross-filesystem overwrite test copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> dummyIo;
        if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
        {
            Fail(L"Dummy filesystem does not support IFileSystemIO for bridge validation.");
            return true;
        }

        const std::wstring dummyMovedPath = std::wstring(dummyMoveRoot) + L"/move.bin";
        unsigned long attrs               = 0;
        if (FAILED(dummyIo->GetAttributes(dummyMovedPath.c_str(), &attrs)))
        {
            Fail(L"Cross-filesystem move: destination file not found in dummy filesystem.");
            return true;
        }

        const std::wstring dummyOverwrittenPath = std::wstring(dummyCopyRoot) + L"/bridge-src/a.bin";

        wil::com_ptr<IFileReader> reader;
        const HRESULT hrReader = dummyIo->CreateFileReader(dummyOverwrittenPath.c_str(), reader.addressof());
        if (FAILED(hrReader) || ! reader)
        {
            Fail(L"Cross-filesystem overwrite test: failed to open destination file in dummy filesystem.");
            return true;
        }

        uint64_t sizeBytes = 0;
        if (FAILED(reader->GetSize(&sizeBytes)) || sizeBytes != 512ull)
        {
            Fail(L"Cross-filesystem overwrite test: destination file size mismatch.");
            return true;
        }

        wil::com_ptr<IFileSystemIO> localIo;
        if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
        {
            Fail(L"Local filesystem does not support IFileSystemIO for metadata validation.");
            return true;
        }

        FileSystemBasicInformation sourceBasic{};
        sourceBasic.sizeBytes                     = sizeof(FileSystemBasicInformation);
        const std::filesystem::path overwriteFile = srcDir / L"a.bin";
        if (FAILED(localIo->GetFileBasicInformation(overwriteFile.c_str(), &sourceBasic)))
        {
            Fail(L"Cross-filesystem metadata test: failed to query source file basic information.");
            return true;
        }

        FileSystemBasicInformation destinationBasic{};
        destinationBasic.sizeBytes = sizeof(FileSystemBasicInformation);
        if (FAILED(dummyIo->GetFileBasicInformation(dummyOverwrittenPath.c_str(), &destinationBasic)))
        {
            Fail(L"Cross-filesystem metadata test: failed to query destination file basic information.");
            return true;
        }

        if (sourceBasic.lastWriteTime != destinationBasic.lastWriteTime || sourceBasic.creationTime != destinationBasic.creationTime)
        {
            Fail(L"Cross-filesystem metadata test: destination timestamps did not match source.");
            return true;
        }

        const char* propsJson = nullptr;
        const HRESULT hrProps = dummyIo->GetItemProperties(dummyMovedPath.c_str(), &propsJson);
        if (FAILED(hrProps) || ! propsJson || propsJson[0] == '\0')
        {
            Fail(L"GetItemProperties returned no JSON for dummy filesystem item.");
            return true;
        }

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyCancelRoot))
        {
            Fail(L"Failed to create dummy folder for bridge cancel atomicity test.");
            return true;
        }

        const std::filesystem::path cancelSource = srcDir / L"bridge_cancel.bin";
        if (! WriteTestFile(cancelSource, kBridgeCancelFileBytes))
        {
            Fail(L"Failed to write bridge cancel atomicity test file.");
            return true;
        }

        const std::wstring dummyCancelPath = std::wstring(dummyCancelRoot) + L"/bridge_cancel.bin";
        unsigned long cancelAttrs          = 0;
        if (SUCCEEDED(dummyIo->GetAttributes(dummyCancelPath.c_str(), &cancelAttrs)))
        {
            const HRESULT hrDel =
                state.fsDummy->DeleteItem(dummyCancelPath.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_NONE), nullptr, nullptr, nullptr);
            if (FAILED(hrDel))
            {
                Fail(L"Failed to clear existing bridge cancel test file in dummy filesystem.");
                return true;
            }
        }

        const FileSystemFlags cancelFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_NONE);
        state.taskB                       = StartFileOperationAndGetId(state.fileOps,
                                                                       FILESYSTEM_COPY,
                                                                       FolderWindow::Pane::Left,
                                                                       FolderWindow::Pane::Right,
                                                                       state.fsLocal,
                                                                       {cancelSource},
                                                                       std::filesystem::path(dummyCancelRoot),
                                                                       cancelFlags,
                                                                       false,
                                                                       kBridgeCancelSpeedLimitBytes,
                                                                       FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                       false,
                                                                       state.fsDummy);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start bridge cancel atomicity copy test (local -> dummy).");
            return true;
        }

        state.markerTick = 0;
        state.stepState  = 6;
        return false;
    }

    if (state.stepState == 6)
    {
        const auto it = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (it != state.completedTasks.end())
        {
            Fail(L"Bridge cancel atomicity copy completed before cancel could be requested.");
            return true;
        }

        FolderWindow::FileOperationState::Task* task = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
        if (task && task->HasEnteredOperation())
        {
            uint64_t completedBytes = 0;
            {
                std::scoped_lock lock(task->_progressMutex);
                completedBytes = task->_progressItemCompletedBytes;
            }

            if (state.markerTick == 0)
            {
                state.markerTick = nowTick;
            }

            // Cancel as soon as we see any progress, or after a short delay to ensure the transfer has started.
            if (completedBytes != 0 || (nowTick >= state.markerTick && (nowTick - state.markerTick) > 500ull))
            {
                task->RequestCancel();
                state.stepState = 7;
            }
        }
        return false;
    }

    if (state.stepState == 7)
    {
        const auto it = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT hr = it->second.hr;
        if (hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) && hr != E_ABORT)
        {
            Fail(std::format(L"Bridge cancel atomicity copy unexpectedly completed: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> dummyIo;
        if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
        {
            Fail(L"Dummy filesystem does not support IFileSystemIO for bridge cancel atomicity validation.");
            return true;
        }

        const std::wstring dummyCancelPath = std::wstring(dummyCancelRoot) + L"/bridge_cancel.bin";
        unsigned long cancelAttrs          = 0;
        if (SUCCEEDED(dummyIo->GetAttributes(dummyCancelPath.c_str(), &cancelAttrs)))
        {
            Fail(L"Bridge cancel atomicity test produced a final destination file on cancel.");
            return true;
        }

        const std::filesystem::path cancelSource = srcDir / L"bridge_cancel.bin";
        std::error_code ec;
        if (! std::filesystem::exists(cancelSource, ec))
        {
            Fail(L"Bridge cancel atomicity test unexpectedly removed the source file.");
            return true;
        }

        std::vector<std::filesystem::path> concurrencySources;
        concurrencySources.reserve(static_cast<size_t>(kBridgeConcurrencyFileCount));
        for (int i = 0; i < kBridgeConcurrencyFileCount; ++i)
        {
            const std::filesystem::path file = srcDir / std::format(L"bridge_conc_{:02}.bin", i);
            if (! WriteTestFile(file, kBridgeConcurrencyFileBytes))
            {
                Fail(L"Failed to write bridge concurrency test file.");
                return true;
            }
            concurrencySources.push_back(file);
        }

        const FileSystemFlags bridgeConcurrencyFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskC                                  = StartFileOperationAndGetId(state.fileOps,
                                                                                  FILESYSTEM_COPY,
                                                                                  FolderWindow::Pane::Left,
                                                                                  FolderWindow::Pane::Right,
                                                                                  state.fsLocal,
                                                                                  std::move(concurrencySources),
                                                                                  std::filesystem::path(dummyCopyRoot),
                                                                                  bridgeConcurrencyFlags,
                                                                                  false,
                                                                                  kBridgeConcurrencySpeedLimit,
                                                                                  FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                                  false,
                                                                                  state.fsDummy);
        if (! state.taskC.has_value())
        {
            Fail(L"Failed to start bridge concurrency copy test.");
            return true;
        }

        state.markerTick = 0;
        state.stepState  = 8;
        return false;
    }

    if (state.stepState == 8)
    {
        FolderWindow::FileOperationState::Task* task = state.taskC.has_value() ? state.fileOps->FindTask(state.taskC.value()) : nullptr;
        if (task && task->HasStarted())
        {
            unsigned int maxConc = 0;
            size_t inFlight      = 0;
            {
                std::scoped_lock lock(task->_progressMutex);
                maxConc  = task->_perItemMaxConcurrency;
                inFlight = task->_perItemInFlightCallCount;
            }

            if (maxConc <= 1u)
            {
                Fail(L"Bridge per-item concurrency expected >1, but task max concurrency is 1.");
                return true;
            }

            if (inFlight > 1u)
            {
                state.markerTick = (std::numeric_limits<ULONGLONG>::max)();
            }
            else if (state.markerTick == 0)
            {
                state.markerTick = nowTick;
            }
            else if (state.markerTick != (std::numeric_limits<ULONGLONG>::max)() && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
            {
                Fail(L"Bridge per-item concurrency expected >1 in-flight calls but did not observe them.");
                return true;
            }
        }

        const auto it = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Bridge concurrency copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> dummyIo;
        if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
        {
            Fail(L"Dummy filesystem does not support IFileSystemIO for bridge concurrency validation.");
            return true;
        }

        for (int i = 0; i < kBridgeConcurrencyFileCount; ++i)
        {
            const std::wstring dummyPath = std::format(L"{}/bridge_conc_{:02}.bin", dummyCopyRoot, i);
            unsigned long attrs          = 0;
            if (FAILED(dummyIo->GetAttributes(dummyPath.c_str(), &attrs)))
            {
                Fail(L"Bridge concurrency output file missing in dummy filesystem.");
                return true;
            }
        }

        NextStep(state, SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase11_BridgeSingleFolderParallelCopyInFlightLines timed out.");
        return true;
    }

    const std::filesystem::path srcDir = state.tempRoot / L"bridge-singlefolder-src";
    const std::wstring dummyRoot       = L"/bridge-singlefolder";

    constexpr int kFileCount          = 12;
    constexpr size_t kFileBytes       = 2ull * 1024ull * 1024ull;
    constexpr uint64_t kSpeedLimitBps = 1ull * 1024ull * 1024ull;

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcDir))
        {
            Fail(L"Failed to reset bridge single-folder source directory.");
            return true;
        }

        for (int i = 0; i < kFileCount; ++i)
        {
            const std::filesystem::path file = srcDir / std::format(L"sf_{:02}.bin", i);
            if (! WriteTestFile(file, kFileBytes))
            {
                Fail(L"Failed to write bridge single-folder test file.");
                return true;
            }
        }

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
        {
            Fail(L"Failed to create dummy folder for bridge single-folder test.");
            return true;
        }

        const FileSystemFlags flags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcDir},
                                                 std::filesystem::path(dummyRoot),
                                                 flags,
                                                 false,
                                                 kSpeedLimitBps,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start bridge single-folder copy test.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    const bool completed                         = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();

    if (state.stepState == 1)
    {
        if (! task || ! task->HasStarted())
        {
            return false;
        }

        state.markerTick = nowTick;
        state.stepState  = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (task)
        {
            size_t inFlightCount     = 0;
            size_t inFlightItemCalls = 0;
            {
                std::scoped_lock lock(task->_progressMutex);
                inFlightCount     = task->_inFlightFileCount;
                inFlightItemCalls = task->_perItemInFlightCallCount;
            }

            if (inFlightItemCalls != 1u)
            {
                Fail(L"Bridge single-folder test expected exactly one top-level in-flight call.");
                return true;
            }

            if (inFlightCount <= 1u)
            {
                if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                {
                    Fail(L"Bridge single-folder test expected >1 in-flight entries but did not observe them.");
                    return true;
                }
                return false;
            }
        }

        state.stepState = 3;
        return false;
    }

    if (! completed)
    {
        return false;
    }

    const auto it = state.completedTasks.find(state.taskA.value());
    if (it == state.completedTasks.end())
    {
        return false;
    }

    if (FAILED(it->second.hr))
    {
        Fail(std::format(L"Bridge single-folder copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
    {
        Fail(L"Dummy filesystem does not support IFileSystemIO for bridge single-folder validation.");
        return true;
    }

    const std::wstring dummyProbe = std::format(L"{}/{}/sf_{:02}.bin", dummyRoot, srcDir.filename().wstring(), 0);
    unsigned long attrs           = 0;
    if (FAILED(dummyIo->GetAttributes(dummyProbe.c_str(), &attrs)))
    {
        Fail(L"Bridge single-folder output file missing in dummy filesystem.");
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines);
    return false;
}
case SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        SetFileOpsBridgeProducerDelayForSelfTest(0);
        Fail(L"Phase11_BridgeMultiFolderParallelCopyInFlightLines timed out.");
        return true;
    }

    const std::filesystem::path srcDir = state.tempRoot / L"bridge-multifolder-src";
    const std::wstring dummyRoot       = L"/bridge-multifolder";

    constexpr int kFolderCount        = 6;
    constexpr int kFilesPerFolder     = 8;
    constexpr size_t kFileBytes       = 2ull * 1024ull * 1024ull;
    constexpr uint64_t kSpeedLimitBps = 1ull * 1024ull * 1024ull;
    constexpr unsigned int kProducerDelayMs = 20u;

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcDir))
        {
            Fail(L"Failed to reset bridge multi-folder source directory.");
            return true;
        }

        for (int folderIndex = 0; folderIndex < kFolderCount; ++folderIndex)
        {
            const std::filesystem::path folder = srcDir / std::format(L"mf_{}", folderIndex);
            std::error_code ec;
            std::filesystem::create_directories(folder, ec);
            if (ec)
            {
                Fail(L"Failed to create bridge multi-folder subdirectory.");
                return true;
            }

            for (int fileIndex = 0; fileIndex < kFilesPerFolder; ++fileIndex)
            {
                const std::filesystem::path file = folder / std::format(L"mf_{:02}.bin", fileIndex);
                if (! WriteTestFile(file, kFileBytes))
                {
                    Fail(L"Failed to write bridge multi-folder test file.");
                    return true;
                }
            }
        }

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
        {
            Fail(L"Failed to create dummy folder for bridge multi-folder test.");
            return true;
        }

        std::vector<std::filesystem::path> sources;
        sources.reserve(static_cast<size_t>(kFolderCount));
        for (int folderIndex = 0; folderIndex < kFolderCount; ++folderIndex)
        {
            sources.push_back(srcDir / std::format(L"mf_{}", folderIndex));
        }

        const FileSystemFlags flags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

        SetFileOpsBridgeProducerDelayForSelfTest(kProducerDelayMs);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 std::move(sources),
                                                 std::filesystem::path(dummyRoot),
                                                 flags,
                                                 false,
                                                 kSpeedLimitBps,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start bridge multi-folder copy test.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    const bool completed                         = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();

    if (state.stepState == 1)
    {
        if (! task || ! task->HasStarted())
        {
            return false;
        }

        state.markerTick = nowTick;
        state.stepState  = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (task)
        {
            size_t inFlightCount     = 0;
            size_t inFlightItemCalls = 0;
            unsigned int budget      = 0;
            {
                std::scoped_lock lock(task->_progressMutex);
                inFlightCount     = task->_inFlightFileCount;
                inFlightItemCalls = task->_perItemInFlightCallCount;
                budget            = task->_perItemMaxConcurrencyBudget;
            }

            if (budget <= 1u)
            {
                Fail(L"Bridge multi-folder test expected within-folder budget >1, but task budget is 1.");
                return true;
            }

            if (inFlightItemCalls < 2u)
            {
                if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                {
                    Fail(L"Bridge multi-folder test expected >=2 top-level in-flight calls but did not observe them.");
                    return true;
                }
                return false;
            }

            if (inFlightCount <= inFlightItemCalls)
            {
                if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                {
                    Fail(L"Bridge multi-folder test expected within-folder parallelism (in-flight lines > in-flight calls) but did not observe it.");
                    return true;
                }
                return false;
            }
        }

        state.stepState = 3;
        return false;
    }

    if (! completed)
    {
        return false;
    }

    const auto it = state.completedTasks.find(state.taskA.value());
    if (it == state.completedTasks.end())
    {
        return false;
    }

    if (FAILED(it->second.hr))
    {
        SetFileOpsBridgeProducerDelayForSelfTest(0);
        Fail(std::format(L"Bridge multi-folder copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
        return true;
    }

    SetFileOpsBridgeProducerDelayForSelfTest(0);

    Debug::Perf::Emit(L"FileOps.SelfTest.BridgeWideShallowDirectoryEnsures",
                      std::format(L"folders={} filesPerFolder={} producerDelayMs={}", kFolderCount, kFilesPerFolder, kProducerDelayMs),
                      0u,
                      it->second.bridgeDirectoryEnsureCount,
                      it->second.bridgeAdmissionMaxQueueDepth,
                      it->second.hr);
    Debug::Perf::Emit(L"FileOps.SelfTest.BridgeWideShallowEarlyFileStarts",
                      std::format(L"folders={} filesPerFolder={} producerDelayMs={}", kFolderCount, kFilesPerFolder, kProducerDelayMs),
                      0u,
                      it->second.bridgeEarlyFileStartCount,
                      it->second.bridgeFileAdmissionCount,
                      it->second.hr);

    if (it->second.bridgeFileAdmissionCount == 0)
    {
        Fail(L"Bridge wide-shallow test did not observe any file admission counters.");
        return true;
    }
    if (it->second.bridgeEarlyFileStartCount == 0)
    {
        Fail(L"Bridge wide-shallow test expected file copying to start before directory producer completion.");
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
    {
        Fail(L"Dummy filesystem does not support IFileSystemIO for bridge multi-folder validation.");
        return true;
    }

    const std::wstring dummyProbe = std::format(L"{}/mf_{}/mf_{:02}.bin", dummyRoot, 0, 0);
    unsigned long attrs           = 0;
    if (FAILED(dummyIo->GetAttributes(dummyProbe.c_str(), &attrs)))
    {
        Fail(L"Bridge multi-folder output file missing in dummy filesystem.");
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase11_BridgePipelineDummyToDummyPerf);
    return false;
}
case SelfTestState::Step::Phase11_BridgePipelineDummyToDummyPerf:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        SetFileOpsBridgePipelineModeForSelfTest(FileOpsBridgePipelineMode::Default);
        if (state.infoDummy && ! state.dummyConfigSnapshot.empty())
        {
            static_cast<void>(SetPluginConfiguration(state.infoDummy.get(), state.dummyConfigSnapshot));
            state.dummyConfigSnapshot.clear();
        }
        Fail(L"Phase11_BridgePipelineDummyToDummyPerf timed out.");
        return true;
    }

    constexpr int kFileCount                         = 4;
    constexpr uint64_t kFileBytes                    = 32ull * 1024ull * 1024ull;
    constexpr unsigned long kStreamChunkLatencyMs    = 30;
    constexpr uint32_t kConfiguredBridgeBufferSizeKB = 4096u;
    constexpr uint32_t kExpectedBridgeBufferSizeKB   = 8192u;
    const std::wstring dummySourceRoot               = L"/bridge-pipeline-src";
    const std::wstring dummySourceFolder             = dummySourceRoot + L"/dataset";
    const std::wstring dummyDestinationRoot          = L"/bridge-pipeline-dst";

    const auto restoreBridgePerfState = [&]() noexcept
    {
        SetFileOpsBridgePipelineModeForSelfTest(FileOpsBridgePipelineMode::Default);
        if (state.infoDummy && ! state.dummyConfigSnapshot.empty())
        {
            static_cast<void>(SetPluginConfiguration(state.infoDummy.get(), state.dummyConfigSnapshot));
            state.dummyConfigSnapshot.clear();
        }
    };

    wil::com_ptr<IFileSystemIO> dummyIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
    {
        restoreBridgePerfState();
        Fail(L"Dummy filesystem does not support IFileSystemIO for bridge pipeline validation.");
        return true;
    }

    const FileSystemFlags flags =
        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

    if (state.stepState == 0)
    {
        state.bridgePipelineBaselineUs   = 0;
        state.bridgePipelineCandidateUs  = 0;
        state.bridgePipelineRunStartTick = 0;
        state.dummyConfigSnapshot.clear();

        if (! BackupPluginConfiguration(state.infoDummy.get(), state.dummyConfigSnapshot))
        {
            restoreBridgePerfState();
            Fail(L"Failed to snapshot dummy configuration for bridge pipeline perf test.");
            return true;
        }

        const std::string config = std::format(
            "{{\"maxChildrenPerDirectory\":42,\"maxDepth\":10,\"seed\":42,\"latencyMs\":0,\"streamChunkLatencyMs\":{},\"virtualSpeedLimit\":\"0\"}}",
            kStreamChunkLatencyMs);
        if (! SetPluginConfiguration(state.infoDummy.get(), config))
        {
            restoreBridgePerfState();
            Fail(L"Failed to apply dummy stream chunk latency for bridge pipeline perf test.");
            return true;
        }

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummySourceRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummySourceFolder) ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), dummyDestinationRoot))
        {
            restoreBridgePerfState();
            Fail(L"Failed to create dummy folders for bridge pipeline perf test.");
            return true;
        }

        for (int fileIndex = 0; fileIndex < kFileCount; ++fileIndex)
        {
            const std::filesystem::path filePath = std::filesystem::path(dummySourceFolder) / std::format(L"bp_{:02}.bin", fileIndex);
            if (! WritePatternFileFsIo(dummyIo, filePath, kFileBytes))
            {
                restoreBridgePerfState();
                Fail(L"Failed to seed dummy source files for bridge pipeline perf test.");
                return true;
            }
        }

        SetFileOpsBridgePipelineModeForSelfTest(FileOpsBridgePipelineMode::Disabled);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummySourceFolder)},
                                                 std::filesystem::path(dummyDestinationRoot),
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            restoreBridgePerfState();
            Fail(L"Failed to start baseline dummy->dummy bridge pipeline copy.");
            return true;
        }

        state.bridgePipelineRunStartTick = nowTick;
        state.stepState                  = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (task && task->HasStarted())
        {
            unsigned int budget                 = 0;
            unsigned long configuredBufferBytes = 0;
            {
                std::scoped_lock lock(task->_progressMutex);
                budget = task->_perItemMaxConcurrencyBudget;
            }
            configuredBufferBytes                         = task->_crossFsBridgeBufferBytes;
            const unsigned long resolvedBridgeBufferBytes = task->_resolvedCrossFsBridgeBufferBytes.load(std::memory_order_acquire);

            if (budget <= 1u)
            {
                restoreBridgePerfState();
                Fail(L"Bridge pipeline perf test expected within-folder bridge budget > 1.");
                return true;
            }

            if (configuredBufferBytes != static_cast<unsigned long>(kConfiguredBridgeBufferSizeKB * 1024u))
            {
                restoreBridgePerfState();
                Fail(std::format(L"Bridge pipeline perf test expected configured task bridge buffer {} bytes but observed {}.",
                                 static_cast<unsigned long>(kConfiguredBridgeBufferSizeKB * 1024u),
                                 configuredBufferBytes));
                return true;
            }

            if (resolvedBridgeBufferBytes != 0 && resolvedBridgeBufferBytes != static_cast<unsigned long>(kExpectedBridgeBufferSizeKB * 1024u))
            {
                restoreBridgePerfState();
                Fail(std::format(L"Bridge pipeline perf test expected resolved bridge buffer {} bytes but observed {}.",
                                 static_cast<unsigned long>(kExpectedBridgeBufferSizeKB * 1024u),
                                 resolvedBridgeBufferBytes));
                return true;
            }
        }

        const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }
        if (FAILED(it->second.hr))
        {
            restoreBridgePerfState();
            Fail(std::format(L"Baseline dummy->dummy bridge pipeline copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        state.bridgePipelineBaselineUs = (state.bridgePipelineRunStartTick != 0 && it->second.completionTick >= state.bridgePipelineRunStartTick)
                                             ? static_cast<uint64_t>(it->second.completionTick - state.bridgePipelineRunStartTick) * 1000ull
                                             : 0ull;

        SetFileOpsBridgePipelineModeForSelfTest(FileOpsBridgePipelineMode::Enabled);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummySourceFolder)},
                                                 std::filesystem::path(dummyDestinationRoot),
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskB.has_value())
        {
            restoreBridgePerfState();
            Fail(L"Failed to start candidate dummy->dummy bridge pipeline copy.");
            return true;
        }

        state.bridgePipelineRunStartTick = nowTick;
        state.stepState                  = 2;
        return false;
    }

    const auto it = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
    if (it == state.completedTasks.end())
    {
        return false;
    }
    if (FAILED(it->second.hr))
    {
        restoreBridgePerfState();
        Fail(std::format(L"Candidate dummy->dummy bridge pipeline copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
        return true;
    }

    state.bridgePipelineCandidateUs = (state.bridgePipelineRunStartTick != 0 && it->second.completionTick >= state.bridgePipelineRunStartTick)
                                          ? static_cast<uint64_t>(it->second.completionTick - state.bridgePipelineRunStartTick) * 1000ull
                                          : 0ull;

    const std::wstring detail = std::format(L"files={} fileBytes={} chunkLatencyMs={} configuredBufferKB={} resolvedBufferKB={} source={} destination={}",
                                            kFileCount,
                                            kFileBytes,
                                            kStreamChunkLatencyMs,
                                            kConfiguredBridgeBufferSizeKB,
                                            kExpectedBridgeBufferSizeKB,
                                            dummySourceFolder,
                                            dummyDestinationRoot);
    Debug::Perf::Emit(L"FileOps.SelfTest.BridgePipelineBaseline",
                      detail,
                      state.bridgePipelineBaselineUs,
                      static_cast<uint64_t>(kFileCount) * kFileBytes,
                      kStreamChunkLatencyMs,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.BridgePipelineCandidate",
                      detail,
                      state.bridgePipelineCandidateUs,
                      static_cast<uint64_t>(kFileCount) * kFileBytes,
                      kStreamChunkLatencyMs,
                      S_OK);
    const uint64_t improvementUs =
        (state.bridgePipelineBaselineUs > state.bridgePipelineCandidateUs) ? (state.bridgePipelineBaselineUs - state.bridgePipelineCandidateUs) : 0ull;
    Debug::Perf::Emit(
        L"FileOps.SelfTest.BridgePipelineImprovement", detail, improvementUs, state.bridgePipelineBaselineUs, state.bridgePipelineCandidateUs, S_OK);

    AppendLog(std::format(L"Phase11_BridgePipelineDummyToDummyPerf baseline={}us candidate={}us latencyMs={} configuredBufferKB={} resolvedBufferKB={}",
                          state.bridgePipelineBaselineUs,
                          state.bridgePipelineCandidateUs,
                          kStreamChunkLatencyMs,
                          kConfiguredBridgeBufferSizeKB,
                          kExpectedBridgeBufferSizeKB));

    const std::wstring dummyProbe = std::format(L"{}/dataset/bp_{:02}.bin", dummyDestinationRoot, 0);
    unsigned long attrs           = 0;
    if (FAILED(dummyIo->GetAttributes(dummyProbe.c_str(), &attrs)))
    {
        restoreBridgePerfState();
        Fail(L"Bridge pipeline perf output file missing in dummy filesystem.");
        return true;
    }

    if (state.bridgePipelineBaselineUs == 0 || state.bridgePipelineCandidateUs == 0)
    {
        restoreBridgePerfState();
        Fail(L"Bridge pipeline perf test captured zero duration.");
        return true;
    }

    const uint64_t improvementThresholdUs = (state.bridgePipelineBaselineUs * 80ull) / 100ull;
    if (state.bridgePipelineCandidateUs > improvementThresholdUs)
    {
        restoreBridgePerfState();
        Fail(std::format(L"Bridge pipeline perf improvement too small: baseline={}us candidate={}us threshold={}us.",
                         state.bridgePipelineBaselineUs,
                         state.bridgePipelineCandidateUs,
                         improvementThresholdUs));
        return true;
    }

    restoreBridgePerfState();
    NextStep(state, SelfTestState::Step::Phase11_ConnectionOverridePrecedence);
    return false;
}
case SelfTestState::Step::Phase11_ConnectionOverridePrecedence:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase11_ConnectionOverridePrecedence timed out.");
        return true;
    }

    constexpr unsigned int kBaselineConcurrency   = 4u;
    constexpr unsigned int kCandidateConcurrency  = 2u;
    constexpr int kFileCount                      = 16;
    constexpr size_t kFileBytes                   = 4ull * 1024ull * 1024ull;
    constexpr unsigned long kStreamChunkLatencyMs = 30u;

    const std::filesystem::path srcDir         = state.tempRoot / L"conn-override-precedence-src";
    const std::wstring connRoot                = L"/conn-override-precedence";
    const std::wstring resolvedBaselineFolder  = connRoot + L"/baseline";
    const std::wstring resolvedCandidateFolder = connRoot + L"/candidate";
    const auto makeDestinationFolder           = [&](std::wstring_view leaf) noexcept -> std::wstring
    { return std::format(L"/@conn:{}/{}", state.connOverrideProfileName, leaf); };

    const auto restoreOverridePerfState = [&]() noexcept
    {
        if (state.infoDummy && ! state.connOverrideDummyConfigSnapshot.empty())
        {
            static_cast<void>(SetPluginConfiguration(state.infoDummy.get(), state.connOverrideDummyConfigSnapshot));
            state.connOverrideDummyConfigSnapshot.clear();
        }
        RemoveConnectionProfileByName(state.connOverrideProfileName);
        state.connOverrideProfileName.clear();
    };

    const auto applyConnectionOverride = [&](std::optional<unsigned int> overrideConcurrency, std::wstring_view label) noexcept -> bool
    {
        if (! g_settings.connections)
        {
            g_settings.connections = Common::Settings::ConnectionsSettings{};
        }

        RemoveConnectionProfileByName(state.connOverrideProfileName);

        Common::Settings::ConnectionProfile profile{};
        profile.id                  = NewGuidString();
        profile.name                = state.connOverrideProfileName;
        profile.pluginId            = std::wstring(kPluginIdDummy);
        profile.initialPath         = connRoot;
        profile.requireWindowsHello = false;
        if (overrideConcurrency.has_value())
        {
            profile.extra = MakeJsonObjectWithUIntMembers({{"copyMoveMaxConcurrency", static_cast<uint64_t>(*overrideConcurrency)}});
        }

        g_settings.connections->items.push_back(profile);
        if (! EnsureDummyFolderExists(state.fsDummy.get(), profile.initialPath) || ! EnsureDummyFolderExists(state.fsDummy.get(), resolvedBaselineFolder) ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), resolvedCandidateFolder))
        {
            Fail(std::format(L"Failed to prepare dummy folders for {}.", label));
            return false;
        }

        return true;
    };

    const auto seedSourceDir = [&]() noexcept -> bool
    {
        if (! RecreateEmptyDirectory(srcDir))
        {
            Fail(L"Failed to reset @conn precedence source directory.");
            return false;
        }

        for (int i = 0; i < kFileCount; ++i)
        {
            const std::filesystem::path file = srcDir / std::format(L"precedence_{:02}.bin", i);
            if (! WriteTestFile(file, kFileBytes))
            {
                Fail(std::format(L"Failed to write @conn precedence source file {}.", file.native()));
                return false;
            }
        }

        return true;
    };

    const auto startCopy = [&](std::wstring_view destinationFolder, std::optional<std::uint64_t>& taskSlot, std::wstring_view label) noexcept -> bool
    {
        std::vector<std::filesystem::path> sources = CollectFiles(srcDir, 512u);
        if (sources.size() != static_cast<size_t>(kFileCount))
        {
            Fail(std::format(L"{} expected {} source files, got {}.", label, kFileCount, sources.size()));
            return false;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        taskSlot                    = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_COPY,
                                                                 FolderWindow::Pane::Left,
                                                                 FolderWindow::Pane::Right,
                                                                 state.fsLocal,
                                                                 std::move(sources),
                                                                 std::filesystem::path(destinationFolder),
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                 false,
                                                                 state.fsDummy);
        if (! taskSlot.has_value())
        {
            Fail(std::format(L"Failed to start {}.", label));
            return false;
        }

        return true;
    };

    const auto observeConfigured = [&](FolderWindow::FileOperationState::Task* task, unsigned int& configured, ULONGLONG& runStartTick) noexcept
    {
        if (! task || ! task->HasStarted())
        {
            return;
        }

        if (runStartTick == 0)
        {
            runStartTick = nowTick;
        }

        std::scoped_lock lock(task->_progressMutex);
        configured = (std::max)(configured, task->_perItemMaxConcurrency);
    };

    const auto finalizeCopy = [&](const std::optional<std::uint64_t>& taskSlot,
                                  const std::wstring& resolvedDestinationFolder,
                                  uint64_t& durationOut,
                                  unsigned int expectedConfigured,
                                  unsigned int observedConfigured,
                                  std::wstring_view label) noexcept -> int
    {
        if (! taskSlot.has_value())
        {
            Fail(std::format(L"{} was never started.", label));
            return -1;
        }

        const auto completionIt = state.completedTasks.find(taskSlot.value());
        if (completionIt == state.completedTasks.end())
        {
            return 0;
        }

        const CompletedTaskInfo& completion = completionIt->second;
        if (FAILED(completion.hr))
        {
            Fail(std::format(L"{} failed: hr=0x{:08X}.", label, static_cast<unsigned long>(completion.hr)));
            return -1;
        }

        if (! completion.started)
        {
            Fail(std::format(L"{} never reported started.", label));
            return -1;
        }

        wil::com_ptr<IFileSystemIO> dummyIo;
        if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
        {
            Fail(L"Dummy filesystem does not support IFileSystemIO for @conn precedence validation.");
            return -1;
        }

        for (int i = 0; i < kFileCount; ++i)
        {
            const std::wstring probe = std::format(L"{}/precedence_{:02}.bin", resolvedDestinationFolder, i);
            unsigned long attrs      = 0;
            if (FAILED(dummyIo->GetAttributes(probe.c_str(), &attrs)))
            {
                Fail(std::format(L"{} output file missing: {}.", label, probe));
                return -1;
            }
        }

        durationOut = (state.connOverridePerfRunStartTick != 0 && completion.completionTick >= state.connOverridePerfRunStartTick)
                          ? static_cast<uint64_t>(completion.completionTick - state.connOverridePerfRunStartTick) * 1000ull
                          : 0ull;
        if (durationOut == 0)
        {
            Fail(std::format(L"{} captured zero duration.", label));
            return -1;
        }

        if (observedConfigured != expectedConfigured)
        {
            Fail(std::format(L"{} expected configured concurrency {} but observed {}.", label, expectedConfigured, observedConfigured));
            return -1;
        }

        return 1;
    };

    const auto verifyConnectionOverridePopup = [&](const std::optional<std::uint64_t>& taskSlot,
                                                   bool& observed,
                                                   unsigned int& resolvedOut,
                                                   unsigned int& appliedOut,
                                                   std::wstring_view label) noexcept -> int
    {
        if (observed)
        {
            return 1;
        }
        if (! taskSlot.has_value())
        {
            Fail(std::format(L"{} was never started for popup diagnostics.", label));
            return -1;
        }

        FileOperationsPopupInternal::TaskSnapshot popupSnapshot{};
        if (! TryGetPopupTaskSnapshot(state.fileOps, taskSlot.value(), popupSnapshot))
        {
            return 0;
        }

        if (! popupSnapshot.autoConcurrencyUsed || popupSnapshot.autoTunedConcurrency != kBaselineConcurrency ||
            popupSnapshot.effectiveConcurrencyBudget != kCandidateConcurrency)
        {
            Fail(std::format(L"{} popup diagnostics mismatch: autoUsed={} resolved={} applied={}.",
                             label,
                             popupSnapshot.autoConcurrencyUsed,
                             popupSnapshot.autoTunedConcurrency,
                             popupSnapshot.effectiveConcurrencyBudget));
            return -1;
        }

        observed    = true;
        resolvedOut = popupSnapshot.autoTunedConcurrency;
        appliedOut  = popupSnapshot.effectiveConcurrencyBudget;
        return 1;
    };

    if (state.stepState == 0)
    {
        state.taskA.reset();
        state.taskB.reset();
        state.connOverridePerfRunStartTick           = 0;
        state.connOverridePerfBaselineUs             = 0;
        state.connOverridePerfCandidateUs            = 0;
        state.connOverridePerfBaselineConfigured     = 0;
        state.connOverridePerfCandidateConfigured    = 0;
        state.connOverridePerfCandidatePopupObserved = false;
        state.connOverridePerfCandidatePopupResolved = 0;
        state.connOverridePerfCandidatePopupApplied  = 0;
        state.connOverrideDummyConfigSnapshot.clear();

        if (state.connOverrideProfileName.empty())
        {
            state.connOverrideProfileName = MakeUniqueConnectionProfileName(L"FileOpsSelfTestDummyPrecedence");
        }

        if (! BackupPluginConfiguration(state.infoDummy.get(), state.connOverrideDummyConfigSnapshot))
        {
            Fail(L"Failed to snapshot dummy configuration for @conn precedence validation.");
            return true;
        }

        const std::string dummyConfig =
            std::format(R"json({{"maxChildrenPerDirectory":42,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":{},"virtualSpeedLimit":"0"}})json",
                        kStreamChunkLatencyMs);
        if (! SetPluginConfiguration(state.infoDummy.get(), dummyConfig))
        {
            restoreOverridePerfState();
            Fail(L"Failed to apply deterministic dummy configuration for @conn precedence validation.");
            return true;
        }

        if (! seedSourceDir() || ! applyConnectionOverride(std::nullopt, L"@conn precedence baseline") ||
            ! startCopy(makeDestinationFolder(L"baseline"), state.taskA, L"@conn precedence baseline"))
        {
            restoreOverridePerfState();
            return true;
        }

        state.connOverridePerfRunStartTick = 0;
        state.stepState                    = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        observeConfigured(state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr,
                          state.connOverridePerfBaselineConfigured,
                          state.connOverridePerfRunStartTick);

        const int baselineStatus = finalizeCopy(state.taskA,
                                                resolvedBaselineFolder,
                                                state.connOverridePerfBaselineUs,
                                                kBaselineConcurrency,
                                                state.connOverridePerfBaselineConfigured,
                                                L"@conn precedence baseline");
        if (baselineStatus < 0)
        {
            restoreOverridePerfState();
            return true;
        }
        if (baselineStatus == 0)
        {
            return false;
        }

        if (! applyConnectionOverride(kCandidateConcurrency, L"@conn precedence candidate") ||
            ! startCopy(makeDestinationFolder(L"candidate"), state.taskB, L"@conn precedence candidate"))
        {
            restoreOverridePerfState();
            return true;
        }

        state.connOverridePerfRunStartTick = 0;
        state.stepState                    = 2;
        return false;
    }

    observeConfigured(state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr,
                      state.connOverridePerfCandidateConfigured,
                      state.connOverridePerfRunStartTick);
    const int candidatePopupStatus = verifyConnectionOverridePopup(state.taskB,
                                                                   state.connOverridePerfCandidatePopupObserved,
                                                                   state.connOverridePerfCandidatePopupResolved,
                                                                   state.connOverridePerfCandidatePopupApplied,
                                                                   L"@conn precedence candidate");
    if (candidatePopupStatus < 0)
    {
        restoreOverridePerfState();
        return true;
    }

    const int candidateStatus = finalizeCopy(state.taskB,
                                             resolvedCandidateFolder,
                                             state.connOverridePerfCandidateUs,
                                             kCandidateConcurrency,
                                             state.connOverridePerfCandidateConfigured,
                                             L"@conn precedence candidate");
    if (candidateStatus < 0)
    {
        restoreOverridePerfState();
        return true;
    }
    if (candidateStatus == 0)
    {
        return false;
    }
    if (candidatePopupStatus == 0)
    {
        restoreOverridePerfState();
        Fail(L"@conn precedence candidate completed before popup diagnostics could be observed.");
        return true;
    }

    const std::wstring detail = std::format(L"fileCount={} fileBytes={} chunkLatencyMs={} baselineConfigured={} candidateConfigured={} "
                                            L"baselineOverride=inherit candidateOverride={} popupResolved={} popupApplied={}",
                                            kFileCount,
                                            kFileBytes,
                                            kStreamChunkLatencyMs,
                                            state.connOverridePerfBaselineConfigured,
                                            state.connOverridePerfCandidateConfigured,
                                            kCandidateConcurrency,
                                            state.connOverridePerfCandidatePopupResolved,
                                            state.connOverridePerfCandidatePopupApplied);
    Debug::Perf::Emit(
        L"FileOps.SelfTest.ConnectionOverrideInherit", detail, state.connOverridePerfBaselineUs, state.connOverridePerfBaselineConfigured, 0, S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.ConnectionOverrideApplied",
                      detail,
                      state.connOverridePerfCandidateUs,
                      state.connOverridePerfCandidateConfigured,
                      kCandidateConcurrency,
                      S_OK);
    const uint64_t durationDeltaUs = (state.connOverridePerfCandidateUs >= state.connOverridePerfBaselineUs)
                                         ? (state.connOverridePerfCandidateUs - state.connOverridePerfBaselineUs)
                                         : (state.connOverridePerfBaselineUs - state.connOverridePerfCandidateUs);
    Debug::Perf::Emit(L"FileOps.SelfTest.ConnectionOverrideDurationDelta",
                      detail,
                      durationDeltaUs,
                      state.connOverridePerfBaselineUs,
                      state.connOverridePerfCandidateUs,
                      S_OK);

    AppendLog(std::format(L"Phase11_ConnectionOverridePrecedence baseline={}us candidate={}us configured={} -> {}",
                          state.connOverridePerfBaselineUs,
                          state.connOverridePerfCandidateUs,
                          state.connOverridePerfBaselineConfigured,
                          state.connOverridePerfCandidateConfigured));

    restoreOverridePerfState();
    NextStep(state, SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate);
    return false;
}
case SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase11_ConnectionOverrideGlobalGate timed out.");
        return true;
    }

    const std::filesystem::path srcDir        = state.tempRoot / L"conn-gate-src";
    constexpr int kFileCount                  = 8;
    constexpr size_t kFileBytes               = 2ull * 1024ull * 1024ull;
    constexpr uint64_t kSpeedLimitBps         = 1ull * 1024ull * 1024ull;
    constexpr size_t kGlobalCopyMoveMax       = 2u;
    constexpr uint64_t kMinimumGateDurationUs = 24ull * 1000ull * 1000ull;

    const auto countActiveStreams = [](const FolderWindow::FileOperationState::Task& task) noexcept -> size_t { return task._perItemInFlightCallCount; };

    if (state.stepState == 0)
    {
        if (state.connOverrideProfileName.empty())
        {
            state.connOverrideProfileName = MakeUniqueConnectionProfileName(L"FileOpsSelfTestDummyGate");
        }

        if (! g_settings.connections)
        {
            g_settings.connections = Common::Settings::ConnectionsSettings{};
        }

        RemoveConnectionProfileByName(state.connOverrideProfileName);

        Common::Settings::ConnectionProfile profile{};
        profile.id                  = NewGuidString();
        profile.name                = state.connOverrideProfileName;
        profile.pluginId            = std::wstring(kPluginIdDummy);
        profile.initialPath         = L"/conn-gate";
        profile.requireWindowsHello = false;
        profile.extra               = MakeJsonObjectWithUIntMembers({{"copyMoveMaxConcurrency", static_cast<uint64_t>(kGlobalCopyMoveMax)}});

        g_settings.connections->items.push_back(profile);

        if (! EnsureDummyFolderExists(state.fsDummy.get(), profile.initialPath))
        {
            Fail(L"Failed to create dummy initialPath for @conn global gate test.");
            return true;
        }

        const std::wstring destinationFolderA = std::format(L"/@conn:{}/gateA", state.connOverrideProfileName);
        const std::wstring destinationFolderB = std::format(L"/@conn:{}/gateB", state.connOverrideProfileName);
        const std::wstring resolvedA          = std::format(L"{}/gateA", profile.initialPath);
        const std::wstring resolvedB          = std::format(L"{}/gateB", profile.initialPath);
        if (! EnsureDummyFolderExists(state.fsDummy.get(), resolvedA) || ! EnsureDummyFolderExists(state.fsDummy.get(), resolvedB))
        {
            Fail(L"Failed to create dummy destination folders for @conn global gate test.");
            return true;
        }

        if (! RecreateEmptyDirectory(srcDir))
        {
            Fail(L"Failed to reset @conn global gate test source directory.");
            return true;
        }

        std::vector<std::filesystem::path> sourcesA;
        std::vector<std::filesystem::path> sourcesB;
        sourcesA.reserve(static_cast<size_t>(kFileCount));
        sourcesB.reserve(static_cast<size_t>(kFileCount));
        for (int i = 0; i < kFileCount; ++i)
        {
            const std::filesystem::path file = srcDir / std::format(L"gate_{:02}.bin", i);
            if (! WriteTestFile(file, kFileBytes))
            {
                Fail(L"Failed to write @conn global gate test file.");
                return true;
            }
            sourcesA.push_back(file);
            sourcesB.push_back(file);
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

        state.connGateMaxActiveCopyStreams = 0;
        state.connGateObservedSaturation   = false;

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 std::move(sourcesA),
                                                 std::filesystem::path(destinationFolderA),
                                                 flags,
                                                 false,
                                                 kSpeedLimitBps,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start @conn global gate copy task A.");
            return true;
        }

        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 std::move(sourcesB),
                                                 std::filesystem::path(destinationFolderB),
                                                 flags,
                                                 false,
                                                 kSpeedLimitBps,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start @conn global gate copy task B.");
            return true;
        }

        state.markerTick = nowTick;
        state.stepState  = 1;
        return false;
    }

    FolderWindow::FileOperationState::Task* taskA = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    FolderWindow::FileOperationState::Task* taskB = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
    const bool completedA                         = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();
    const bool completedB                         = state.taskB.has_value() && state.completedTasks.find(state.taskB.value()) != state.completedTasks.end();

    if (state.stepState == 1)
    {
        if (! taskA || ! taskB || ! taskA->HasStarted() || ! taskB->HasStarted())
        {
            return false;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (taskA && taskB)
        {
            size_t activeA = 0;
            size_t activeB = 0;
            {
                std::scoped_lock lock(taskA->_progressMutex, taskB->_progressMutex);
                activeA = countActiveStreams(*taskA);
                activeB = countActiveStreams(*taskB);
            }

            const size_t totalActive           = activeA + activeB;
            state.connGateMaxActiveCopyStreams = (std::max)(state.connGateMaxActiveCopyStreams, totalActive);
            if (totalActive == kGlobalCopyMoveMax)
            {
                state.connGateObservedSaturation = true;
            }
        }

        if (! completedA || ! completedB)
        {
            return false;
        }

        const auto itA = state.completedTasks.find(state.taskA.value());
        const auto itB = state.completedTasks.find(state.taskB.value());
        if (itA == state.completedTasks.end() || itB == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itA->second.hr) || FAILED(itB->second.hr))
        {
            Fail(std::format(L"@conn global gate copy failed: A=0x{:08X} B=0x{:08X}.",
                             static_cast<unsigned long>(itA->second.hr),
                             static_cast<unsigned long>(itB->second.hr)));
            return true;
        }

        const ULONGLONG completionTick = (std::max)(itA->second.completionTick, itB->second.completionTick);
        const uint64_t gateDurationUs =
            (state.markerTick != 0 && completionTick >= state.markerTick) ? static_cast<uint64_t>(completionTick - state.markerTick) * 1000ull : 0ull;
        const std::wstring detail = std::format(L"fileCountPerTask={} fileBytes={} speedLimitBps={} maxObservedActive={} observedSaturation={}",
                                                kFileCount,
                                                kFileBytes,
                                                kSpeedLimitBps,
                                                state.connGateMaxActiveCopyStreams,
                                                state.connGateObservedSaturation ? 1 : 0);
        Debug::Perf::Emit(L"FileOps.SelfTest.ConnectionOverrideGlobalGateDuration",
                          detail,
                          gateDurationUs,
                          static_cast<uint64_t>(kGlobalCopyMoveMax),
                          state.connGateMaxActiveCopyStreams,
                          S_OK);

        if (gateDurationUs < kMinimumGateDurationUs)
        {
            Fail(std::format(L"@conn global gate duration too short: duration={}us threshold={}us maxObservedActive={}.",
                             gateDurationUs,
                             kMinimumGateDurationUs,
                             state.connGateMaxActiveCopyStreams));
            return true;
        }

        RemoveConnectionProfileByName(state.connOverrideProfileName);
        state.connOverrideProfileName.clear();

        NextStep(state, SelfTestState::Step::Phase11_ConnectionOverrideClamp);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase11_ConnectionOverrideClamp:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase11_ConnectionOverrideClamp timed out.");
        return true;
    }

    const std::filesystem::path srcDir = state.tempRoot / L"conn-override-src";
    constexpr int kFileCount           = 6;
    constexpr size_t kFileBytes        = 2ull * 1024ull * 1024ull;
    constexpr uint64_t kSpeedLimitBps  = 1ull * 1024ull * 1024ull;

    if (state.stepState == 0)
    {
        if (state.connOverrideProfileName.empty())
        {
            state.connOverrideProfileName = MakeUniqueConnectionProfileName(L"FileOpsSelfTestDummy");
        }

        if (! g_settings.connections)
        {
            g_settings.connections = Common::Settings::ConnectionsSettings{};
        }

        RemoveConnectionProfileByName(state.connOverrideProfileName);

        Common::Settings::ConnectionProfile profile{};
        profile.id                  = NewGuidString();
        profile.name                = state.connOverrideProfileName;
        profile.pluginId            = std::wstring(kPluginIdDummy);
        profile.initialPath         = L"/conn-selftest";
        profile.requireWindowsHello = false;
        profile.extra               = MakeJsonObjectWithUIntMembers({{"copyMoveMaxConcurrency", 1ull}, {"deleteMaxConcurrency", 1ull}});

        g_settings.connections->items.push_back(profile);

        if (! EnsureDummyFolderExists(state.fsDummy.get(), profile.initialPath))
        {
            Fail(L"Failed to create dummy initialPath for @conn override test.");
            return true;
        }

        const std::wstring destinationFolder         = std::format(L"/@conn:{}/copy", state.connOverrideProfileName);
        const std::wstring resolvedDestinationFolder = std::format(L"{}/copy", profile.initialPath);
        if (! EnsureDummyFolderExists(state.fsDummy.get(), resolvedDestinationFolder))
        {
            Fail(L"Failed to create dummy destination folder for @conn override test.");
            return true;
        }

        wil::com_ptr<IFileSystemIO> dummyIo;
        if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
        {
            Fail(L"Dummy filesystem does not support IFileSystemIO for @conn override test validation.");
            return true;
        }

        {
            unsigned long attrs  = 0;
            const HRESULT hrAttr = dummyIo->GetAttributes(resolvedDestinationFolder.c_str(), &attrs);
            if (FAILED(hrAttr) || (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
            {
                Fail(std::format(L"@conn override preflight: resolved destination folder missing (hr=0x{:08X}).", static_cast<unsigned long>(hrAttr)));
                return true;
            }
        }

        {
            unsigned long attrs  = 0;
            const HRESULT hrAttr = dummyIo->GetAttributes(destinationFolder.c_str(), &attrs);
            if (FAILED(hrAttr) || (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
            {
                Fail(std::format(L"@conn override preflight: @conn destination folder did not resolve (hr=0x{:08X}).", static_cast<unsigned long>(hrAttr)));
                return true;
            }
        }

        if (! RecreateEmptyDirectory(srcDir))
        {
            Fail(L"Failed to reset @conn override test source directory.");
            return true;
        }

        std::vector<std::filesystem::path> sources;
        sources.reserve(static_cast<size_t>(kFileCount));
        for (int i = 0; i < kFileCount; ++i)
        {
            const std::filesystem::path file = srcDir / std::format(L"ovr_{:02}.bin", i);
            if (! WriteTestFile(file, kFileBytes))
            {
                Fail(L"Failed to write @conn override test file.");
                return true;
            }
            sources.push_back(file);
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 std::move(sources),
                                                 std::filesystem::path(destinationFolder),
                                                 flags,
                                                 false,
                                                 kSpeedLimitBps,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start @conn override copy test.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (! task || ! task->HasStarted())
        {
            return false;
        }

        unsigned int budget  = 0;
        unsigned int maxConc = 0;
        {
            std::scoped_lock lock(task->_progressMutex);
            budget  = task->_perItemMaxConcurrencyBudget;
            maxConc = task->_perItemMaxConcurrency;
        }

        if (budget != 1u || maxConc != 1u)
        {
            Fail(std::format(L"@conn override copy clamp expected 1, got budget={} max={}.", budget, maxConc));
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }
        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"@conn override copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        std::vector<std::filesystem::path> deletePaths;
        deletePaths.reserve(static_cast<size_t>(kFileCount));
        for (int i = 0; i < kFileCount; ++i)
        {
            deletePaths.push_back(std::filesystem::path(std::format(L"/@conn:{}/copy/ovr_{:02}.bin", state.connOverrideProfileName, i)));
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_NONE);
        state.taskB                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_DELETE,
                                                                 FolderWindow::Pane::Right,
                                                                 std::nullopt,
                                                                 state.fsDummy,
                                                                 std::move(deletePaths),
                                                                 {},
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start @conn override delete test.");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        FolderWindow::FileOperationState::Task* task = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
        if (! task || ! task->HasStarted())
        {
            return false;
        }

        unsigned int budget  = 0;
        unsigned int maxConc = 0;
        {
            std::scoped_lock lock(task->_progressMutex);
            budget  = task->_perItemMaxConcurrencyBudget;
            maxConc = task->_perItemMaxConcurrency;
        }

        if (budget != 1u || maxConc != 1u)
        {
            Fail(std::format(L"@conn override delete clamp expected 1, got budget={} max={}.", budget, maxConc));
            return true;
        }

        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        const auto it = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }
        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"@conn override delete failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        RemoveConnectionProfileByName(state.connOverrideProfileName);
        state.connOverrideProfileName.clear();

        NextStep(state, SelfTestState::Step::Phase12_ReparsePointPolicy);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase12_ReparsePointPolicy:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        const auto completed = [&](const std::optional<std::uint64_t>& taskId) noexcept -> bool
        {
            if (! taskId.has_value())
            {
                return false;
            }
            return state.completedTasks.find(taskId.value()) != state.completedTasks.end();
        };

        const bool aDone = completed(state.taskA);
        const bool bDone = completed(state.taskB);
        const bool cDone = completed(state.taskC);

        bool promptActive = false;
        if (state.fileOps && state.taskA.has_value())
        {
            if (auto* task = state.fileOps->FindTask(state.taskA.value()))
            {
                promptActive = TryGetConflictPromptCopy(task).has_value();
            }
        }

        Fail(std::format(L"Phase12_ReparsePointPolicy timed out (stepState={} taskA={} doneA={} taskB={} doneB={} taskC={} doneC={} promptActive={}).",
                         state.stepState,
                         state.taskA.value_or(0ull),
                         aDone ? 1 : 0,
                         state.taskB.value_or(0ull),
                         bDone ? 1 : 0,
                         state.taskC.value_or(0ull),
                         cDone ? 1 : 0,
                         promptActive ? 1 : 0));
        return true;
    }

    const std::filesystem::path srcDir                = state.tempRoot / L"reparse-src";
    const std::filesystem::path dstDir                = state.tempRoot / L"reparse-dst";
    const std::filesystem::path moveSrc               = state.tempRoot / L"reparse-move-src";
    const std::filesystem::path moveDst               = state.tempRoot / L"reparse-move-dst";
    const std::filesystem::path delDir                = state.tempRoot / L"reparse-delete";
    const std::filesystem::path targetDir             = state.tempRoot / L"reparse-target";
    const std::filesystem::path targetFile            = targetDir / L"keep.bin";
    const std::filesystem::path bridgeMoveRootReparse = state.tempRoot / L"bridge-move-root-link";
    const std::filesystem::path bridgeCopyRootReparse = state.tempRoot / L"bridge-copy-root-link";

    const std::wstring dummyBridgeMoveRoot = L"/bridge-reparse-move";
    const std::wstring dummyBridgeCopyRoot = L"/bridge-reparse-copy";

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"copyReparse"})json"));

        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir) || ! RecreateEmptyDirectory(moveSrc) || ! RecreateEmptyDirectory(moveDst) ||
            ! RecreateEmptyDirectory(delDir) || ! RecreateEmptyDirectory(targetDir))
        {
            Fail(L"Failed to reset reparse test directories.");
            return true;
        }

        std::error_code ec;
        static_cast<void>(std::filesystem::remove_all(bridgeMoveRootReparse, ec));
        ec.clear();
        static_cast<void>(std::filesystem::remove_all(bridgeCopyRootReparse, ec));

        if (! WriteTestFile(srcDir / L"seed.bin", 128) || ! WriteTestFile(moveSrc / L"moved.bin", 96) || ! WriteTestFile(targetFile, 256))
        {
            Fail(L"Failed to write reparse test files.");
            return true;
        }

        // Create a junction loop inside the tree: srcDir\\loop -> srcDir.
        const std::filesystem::path loop = srcDir / L"loop";
        if (! TryCreateJunction(loop, srcDir))
        {
            Fail(L"Failed to create junction loop for reparse copy test.");
            return true;
        }
        if (! TryDenyListDirectoryToEveryone(loop))
        {
            Fail(L"Failed to apply protected junction ACL for reparse copy test.");
            return true;
        }

        const std::filesystem::path linkToTarget = srcDir / L"linkToTarget";
        if (! TryCreateJunction(linkToTarget, targetDir))
        {
            Fail(L"Failed to create out-of-tree junction for reparse copy test.");
            return true;
        }

        const std::filesystem::path moveLink = moveSrc / L"toTarget";
        if (! TryCreateJunction(moveLink, targetDir))
        {
            Fail(L"Failed to create move reparse link.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        state.taskA                 = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {srcDir}, dstDir, flags, false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start reparse copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Reparse copy task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        const std::filesystem::path copiedLoop = dstDir / srcDir.filename() / L"loop";
        const auto tag                         = TryGetReparseTag(copiedLoop);
        if (! tag.has_value() || (tag.value() != IO_REPARSE_TAG_MOUNT_POINT && tag.value() != IO_REPARSE_TAG_SYMLINK))
        {
            Fail(L"Reparse copy did not recreate loop as a directory reparse point.");
            return true;
        }

        const auto copiedLoopTarget = TryGetDirectoryReparseTargetAbsolute(copiedLoop);
        if (! copiedLoopTarget.has_value())
        {
            Fail(L"Reparse copy could not read copied loop target.");
            return true;
        }

        const std::wstring expectedLoopTarget = NormalizePathForCompare((dstDir / srcDir.filename()).wstring());
        if (copiedLoopTarget.value() != expectedLoopTarget)
        {
            Fail(std::format(L"Reparse copy loop target mismatch. expected='{}' actual='{}'.", expectedLoopTarget, copiedLoopTarget.value()));
            return true;
        }

        const std::filesystem::path copiedOutOfTree = dstDir / srcDir.filename() / L"linkToTarget";
        const auto copiedOutTarget                  = TryGetDirectoryReparseTargetAbsolute(copiedOutOfTree);
        if (! copiedOutTarget.has_value())
        {
            Fail(L"Reparse copy could not read copied out-of-tree junction target.");
            return true;
        }

        const std::wstring expectedOutTarget = NormalizePathForCompare(std::filesystem::absolute(targetDir).wstring());
        if (copiedOutTarget.value() != expectedOutTarget)
        {
            Fail(std::format(L"Reparse copy out-of-tree target mismatch. expected='{}' actual='{}'.", expectedOutTarget, copiedOutTarget.value()));
            return true;
        }

        const FileSystemFlags moveFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        state.taskB                     = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_MOVE, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {moveSrc}, moveDst, moveFlags, false);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start local move reparse task.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto itMove = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (itMove == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itMove->second.hr))
        {
            Fail(std::format(L"Local move reparse task failed: 0x{:08X}.", static_cast<unsigned long>(itMove->second.hr)));
            return true;
        }

        std::error_code ec;
        if (std::filesystem::exists(moveSrc, ec))
        {
            Fail(L"Local move reparse task did not remove source directory.");
            return true;
        }

        const std::filesystem::path movedLink = moveDst / moveSrc.filename() / L"toTarget";
        const auto movedTarget                = TryGetDirectoryReparseTargetAbsolute(movedLink);
        if (! movedTarget.has_value())
        {
            Fail(L"Local move reparse task did not preserve moved link.");
            return true;
        }

        const std::wstring expectedMoveTarget = NormalizePathForCompare(std::filesystem::absolute(targetDir).wstring());
        if (movedTarget.value() != expectedMoveTarget)
        {
            Fail(std::format(L"Local move reparse target mismatch. expected='{}' actual='{}'.", expectedMoveTarget, movedTarget.value()));
            return true;
        }

        wil::com_ptr<IFileSystemIO> localIo;
        if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
        {
            Fail(L"Local filesystem does not support IFileSystemIO for metadata validation.");
            return true;
        }

        const std::filesystem::path metadataSrcRoot = state.tempRoot / L"reparse-metadata-src";
        const std::filesystem::path metadataDstRoot = state.tempRoot / L"reparse-metadata-dst";
        if (! RecreateEmptyDirectory(metadataSrcRoot) || ! RecreateEmptyDirectory(metadataDstRoot) || ! WriteTestFile(metadataSrcRoot / L"seed.bin", 64))
        {
            Fail(L"Failed to prepare directory metadata copy test.");
            return true;
        }

        FileSystemBasicInformation sourceDirBasic{};
        sourceDirBasic.sizeBytes = sizeof(FileSystemBasicInformation);
        if (FAILED(localIo->GetFileBasicInformation(metadataSrcRoot.c_str(), &sourceDirBasic)))
        {
            Fail(L"Directory metadata test: failed to query source directory basic information.");
            return true;
        }

        constexpr int64_t kMetadataTickOffset = 3ll * 24ll * 60ll * 60ll * 10'000'000ll;
        sourceDirBasic.creationTime -= kMetadataTickOffset;
        sourceDirBasic.lastWriteTime -= kMetadataTickOffset;
        sourceDirBasic.attributes |= FILE_ATTRIBUTE_READONLY;
        if (FAILED(localIo->SetFileBasicInformation(metadataSrcRoot.c_str(), &sourceDirBasic)))
        {
            Fail(L"Directory metadata test: failed to seed source directory basic information.");
            return true;
        }

        FileSystemBasicInformation seededSourceDirBasic{};
        seededSourceDirBasic.sizeBytes = sizeof(FileSystemBasicInformation);
        if (FAILED(localIo->GetFileBasicInformation(metadataSrcRoot.c_str(), &seededSourceDirBasic)))
        {
            Fail(L"Directory metadata test: failed to re-query seeded source directory basic information.");
            return true;
        }

        const std::filesystem::path metadataCopyPath = metadataDstRoot / metadataSrcRoot.filename();
        const HRESULT metadataCopyHr                 = state.fsLocal->CopyItem(
            metadataSrcRoot.c_str(), metadataCopyPath.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), nullptr, nullptr, nullptr);
        if (FAILED(metadataCopyHr))
        {
            Fail(std::format(L"Directory metadata test: local copy failed: 0x{:08X}.", static_cast<unsigned long>(metadataCopyHr)));
            return true;
        }

        const std::filesystem::path metadataCopiedDir = metadataCopyPath;
        FileSystemBasicInformation copiedDirBasic{};
        copiedDirBasic.sizeBytes = sizeof(FileSystemBasicInformation);
        if (FAILED(localIo->GetFileBasicInformation(metadataCopiedDir.c_str(), &copiedDirBasic)))
        {
            Fail(L"Directory metadata test: failed to query copied directory basic information.");
            return true;
        }

        if (copiedDirBasic.creationTime != seededSourceDirBasic.creationTime || copiedDirBasic.lastWriteTime != seededSourceDirBasic.lastWriteTime)
        {
            Fail(std::format(L"Directory metadata test: copied directory timestamps did not match source (srcCreation={} dstCreation={} srcWrite={} "
                             L"dstWrite={} srcAttr=0x{:08X} dstAttr=0x{:08X}).",
                             seededSourceDirBasic.creationTime,
                             copiedDirBasic.creationTime,
                             seededSourceDirBasic.lastWriteTime,
                             copiedDirBasic.lastWriteTime,
                             static_cast<unsigned long>(seededSourceDirBasic.attributes),
                             static_cast<unsigned long>(copiedDirBasic.attributes)));
            return true;
        }

        if ((copiedDirBasic.attributes & FILE_ATTRIBUTE_READONLY) == 0)
        {
            Fail(L"Directory metadata test: copied directory did not preserve read-only attribute.");
            return true;
        }

        seededSourceDirBasic.attributes &= ~FILE_ATTRIBUTE_READONLY;
        static_cast<void>(localIo->SetFileBasicInformation(metadataSrcRoot.c_str(), &seededSourceDirBasic));
        copiedDirBasic.attributes &= ~FILE_ATTRIBUTE_READONLY;
        static_cast<void>(localIo->SetFileBasicInformation(metadataCopiedDir.c_str(), &copiedDirBasic));

        const std::filesystem::path rollbackRoot   = state.tempRoot / L"reparse-move-rollback";
        const std::filesystem::path rollbackTarget = rollbackRoot / L"target";
        const std::filesystem::path rollbackSrc    = rollbackRoot / L"src";
        const std::filesystem::path rollbackDst    = rollbackRoot / L"dst";
        if (! RecreateEmptyDirectory(rollbackRoot) || ! RecreateEmptyDirectory(rollbackTarget) || ! RecreateEmptyDirectory(rollbackSrc) ||
            ! RecreateEmptyDirectory(rollbackDst) || ! WriteTestFile(rollbackTarget / L"payload.bin", 32))
        {
            Fail(L"Failed to prepare reparse move rollback test.");
            return true;
        }

        const std::filesystem::path rollbackLink = rollbackSrc / L"junction";
        if (! TryCreateJunction(rollbackLink, rollbackTarget))
        {
            Fail(L"Failed to create reparse move rollback source junction.");
            return true;
        }

        wil::unique_handle rollbackHandle(CreateFileW(rollbackLink.c_str(),
                                                      FILE_READ_ATTRIBUTES,
                                                      FILE_SHARE_READ | FILE_SHARE_WRITE,
                                                      nullptr,
                                                      OPEN_EXISTING,
                                                      FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                                      nullptr));
        if (! rollbackHandle)
        {
            Fail(L"Failed to open rollback lock handle for reparse move test.");
            return true;
        }

        const std::filesystem::path rollbackDestLink = rollbackDst / rollbackLink.filename();
        const HRESULT rollbackMoveHr                 = state.fsLocal->MoveItem(
            rollbackLink.c_str(), rollbackDestLink.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_NONE), nullptr, nullptr, nullptr);
        rollbackHandle.reset();
        if (rollbackMoveHr == S_OK)
        {
            if (GetFileAttributesW(rollbackLink.c_str()) != INVALID_FILE_ATTRIBUTES)
            {
                Fail(L"Reparse move lock test reported success but left the source junction behind.");
                return true;
            }

            if (GetFileAttributesW(rollbackDestLink.c_str()) == INVALID_FILE_ATTRIBUTES)
            {
                Fail(L"Reparse move lock test reported success but did not materialize the destination junction.");
                return true;
            }

            const auto movedRollbackTarget = TryGetDirectoryReparseTargetAbsolute(rollbackDestLink);
            if (! movedRollbackTarget.has_value())
            {
                Fail(L"Reparse move lock test reported success but the destination junction target could not be read.");
                return true;
            }

            const std::wstring expectedRollbackTarget = NormalizePathForCompare(std::filesystem::absolute(rollbackTarget).wstring());
            if (movedRollbackTarget.value() != expectedRollbackTarget)
            {
                Fail(std::format(L"Reparse move lock test target mismatch. expected='{}' actual='{}'.", expectedRollbackTarget, movedRollbackTarget.value()));
                return true;
            }
        }
        else if (rollbackMoveHr != HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) && rollbackMoveHr != HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED))
        {
            Fail(std::format(L"Reparse move lock test returned unexpected hr: 0x{:08X}.", static_cast<unsigned long>(rollbackMoveHr)));
            return true;
        }
        else if (GetFileAttributesW(rollbackLink.c_str()) == INVALID_FILE_ATTRIBUTES)
        {
            Fail(L"Reparse move rollback test removed the source junction.");
            return true;
        }
        else if (GetFileAttributesW(rollbackDestLink.c_str()) != INVALID_FILE_ATTRIBUTES)
        {
            Fail(L"Reparse move rollback test left a destination junction behind.");
            return true;
        }

        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"skip"})json"));

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyBridgeMoveRoot))
        {
            Fail(L"Failed to prepare dummy root for bridge move reparse test.");
            return true;
        }

        if (! TryCreateJunction(bridgeMoveRootReparse, targetDir))
        {
            Fail(L"Failed to create bridge move root reparse source.");
            return true;
        }

        const FileSystemFlags bridgeMoveFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskC                           = StartFileOperationAndGetId(state.fileOps,
                                                                           FILESYSTEM_MOVE,
                                                                           FolderWindow::Pane::Left,
                                                                           FolderWindow::Pane::Right,
                                                                           state.fsLocal,
                                                                           {bridgeMoveRootReparse},
                                                                           std::filesystem::path(dummyBridgeMoveRoot),
                                                                           bridgeMoveFlags,
                                                                           false,
                                                                           0,
                                                                           FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                           false,
                                                                           state.fsDummy);
        if (! state.taskC.has_value())
        {
            Fail(L"Failed to start bridge move reparse task.");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        const auto itBridgeMove = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
        if (itBridgeMove == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT expectedPartial = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (itBridgeMove->second.hr != expectedPartial)
        {
            Fail(std::format(L"Bridge move reparse expected partial (0x{:08X}) but got 0x{:08X}.",
                             static_cast<unsigned long>(expectedPartial),
                             static_cast<unsigned long>(itBridgeMove->second.hr)));
            return true;
        }

        std::error_code ec;
        if (! std::filesystem::exists(bridgeMoveRootReparse, ec))
        {
            Fail(L"Bridge move reparse skipped item but source link was removed.");
            return true;
        }

        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"copyReparse"})json"));

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyBridgeCopyRoot))
        {
            Fail(L"Failed to prepare dummy root for bridge copy unsupported test.");
            return true;
        }

        if (! TryCreateJunction(bridgeCopyRootReparse, targetDir))
        {
            Fail(L"Failed to create bridge copy root reparse source.");
            return true;
        }

        const FileSystemFlags bridgeCopyFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                           = StartFileOperationAndGetId(state.fileOps,
                                                                           FILESYSTEM_COPY,
                                                                           FolderWindow::Pane::Left,
                                                                           FolderWindow::Pane::Right,
                                                                           state.fsLocal,
                                                                           {bridgeCopyRootReparse},
                                                                           std::filesystem::path(dummyBridgeCopyRoot),
                                                                           bridgeCopyFlags,
                                                                           false,
                                                                           0,
                                                                           FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                           false,
                                                                           state.fsDummy);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start bridge copy unsupported reparse task.");
            return true;
        }

        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        using Task = FolderWindow::FileOperationState::Task;

        Task* task        = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }

        if (! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
        {
            Fail(L"Bridge copy unsupported reparse prompt did not offer Skip.");
            return true;
        }

        if (prompt->bucket != Task::ConflictBucket::UnsupportedReparse)
        {
            Fail(L"Bridge copy unsupported reparse prompt did not classify as UnsupportedReparse bucket.");
            return true;
        }

        task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        const auto itBridgeCopy = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (itBridgeCopy == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT expectedPartial = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (itBridgeCopy->second.hr != expectedPartial)
        {
            Fail(std::format(L"Bridge copy unsupported reparse expected partial (0x{:08X}) but got 0x{:08X}.",
                             static_cast<unsigned long>(expectedPartial),
                             static_cast<unsigned long>(itBridgeCopy->second.hr)));
            return true;
        }

        const std::filesystem::path linkToTarget = delDir / L"linkToTarget";
        if (! TryCreateJunction(linkToTarget, targetDir))
        {
            Fail(L"Failed to create junction for reparse delete test.");
            return true;
        }

        const FileSystemFlags deleteFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        state.taskB                       = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {delDir}, {}, deleteFlags, false);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start reparse delete task.");
            return true;
        }

        state.stepState = 6;
        return false;
    }

    if (state.stepState == 6)
    {
        const auto itDelete = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (itDelete == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itDelete->second.hr))
        {
            Fail(std::format(L"Reparse delete task failed: 0x{:08X}.", static_cast<unsigned long>(itDelete->second.hr)));
            return true;
        }

        std::error_code ec;
        if (std::filesystem::exists(delDir, ec))
        {
            Fail(L"Reparse delete task did not remove the source directory.");
            return true;
        }

        ec.clear();
        if (! std::filesystem::exists(targetFile, ec))
        {
            Fail(L"Reparse delete task removed the junction target (should remain).");
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase13_PostMortemDiagnostics);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase13_PostMortemDiagnostics:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        Fail(L"Phase13_PostMortemDiagnostics timed out.");
        return true;
    }

    if (! state.fileOps)
    {
        Fail(L"Phase13_PostMortemDiagnostics missing file operation state.");
        return true;
    }

    if (state.stepState == 0)
    {
        const std::filesystem::path diagnosticSrc = state.tempRoot / L"phase13-diagnostics-src";
        const std::filesystem::path diagnosticDst = state.tempRoot / L"phase13-diagnostics-dst";
        if (! RecreateEmptyDirectory(diagnosticSrc) || ! RecreateEmptyDirectory(diagnosticDst))
        {
            Fail(L"Phase13_PostMortemDiagnostics could not create diagnostic seed folders.");
            return true;
        }

        if (! WriteTestFile(diagnosticSrc / L"ok.bin", 64))
        {
            Fail(L"Phase13_PostMortemDiagnostics could not create diagnostic seed source file.");
            return true;
        }

        const FileSystemFlags diagnosticFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        state.taskA                           = StartFileOperationAndGetId(state.fileOps,
                                                                           FILESYSTEM_COPY,
                                                                           FolderWindow::Pane::Left,
                                                                           FolderWindow::Pane::Right,
                                                                           state.fsLocal,
                                                                           {diagnosticSrc / L"ok.bin", diagnosticSrc / L"missing.bin"},
                                                                           diagnosticDst,
                                                                           diagnosticFlags,
                                                                           false);
        if (! state.taskA.has_value())
        {
            Fail(L"Phase13_PostMortemDiagnostics could not start diagnostic seed copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        const auto itDiagnostic = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (itDiagnostic == state.completedTasks.end())
        {
            return false;
        }

        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        const auto summaryIt = std::find_if(
            summaries.begin(), summaries.end(), [&](const auto& summary) noexcept { return state.taskA.has_value() && summary.taskId == state.taskA.value(); });
        if (summaryIt == summaries.end())
        {
            Fail(L"Phase13_PostMortemDiagnostics could not find the diagnostic seed task summary.");
            return true;
        }

        const bool hasDiagnostics = summaryIt->warningCount > 0 || summaryIt->errorCount > 0;
        if (FAILED(summaryIt->resultHr) && ! hasDiagnostics)
        {
            Fail(std::format(L"Phase13_PostMortemDiagnostics task {} failed without warning/error diagnostics.", summaryIt->taskId));
            return true;
        }

        if (! hasDiagnostics)
        {
            Fail(L"Phase13_PostMortemDiagnostics expected at least one completed summary with diagnostics.");
            return true;
        }

        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(L"RedSalamander");
        if (settingsPath.empty())
        {
            Fail(L"Phase13_PostMortemDiagnostics could not resolve settings path.");
            return true;
        }

        const std::filesystem::path settingsDir = settingsPath.parent_path();
        const std::filesystem::path logsDir     = settingsDir.parent_path().empty() ? (settingsDir / L"Logs") : (settingsDir.parent_path() / L"Logs");
        std::error_code ec;
        bool foundLogFile = false;
        for (std::filesystem::directory_iterator it(logsDir, ec), end; ! ec && it != end; it.increment(ec))
        {
            const auto& de = *it;
            if (! de.is_regular_file(ec))
            {
                continue;
            }

            const std::wstring fileName = de.path().filename().wstring();
            if (fileName.rfind(L"FileOperations-", 0) != 0 || de.path().extension().wstring() != L".jsonl")
            {
                continue;
            }

            const auto size = de.file_size(ec);
            if (! ec && size > 0)
            {
                foundLogFile = true;
                break;
            }
        }

        if (! foundLogFile)
        {
            Fail(L"Phase13_PostMortemDiagnostics did not find persisted file operation diagnostics logs.");
            return true;
        }

        std::filesystem::path issuesReportPath;
        if (! state.fileOps->ExportTaskIssuesReport(summaryIt->taskId, &issuesReportPath, false))
        {
            Fail(L"Phase13_PostMortemDiagnostics could not export task issues report.");
            return true;
        }

        if (issuesReportPath.empty())
        {
            Fail(L"Phase13_PostMortemDiagnostics exported issues report path is empty.");
            return true;
        }

        ec.clear();
        if (! std::filesystem::exists(issuesReportPath, ec) || ec)
        {
            Fail(L"Phase13_PostMortemDiagnostics exported issues report file does not exist.");
            return true;
        }

        const auto reportSize = std::filesystem::file_size(issuesReportPath, ec);
        if (ec || reportSize == 0)
        {
            Fail(L"Phase13_PostMortemDiagnostics exported issues report file is empty.");
            return true;
        }

        const std::filesystem::path autoDismissSrc = state.tempRoot / L"phase13-auto-dismiss-src";
        const std::filesystem::path autoDismissDst = state.tempRoot / L"phase13-auto-dismiss-dst";
        if (! RecreateEmptyDirectory(autoDismissSrc) || ! RecreateEmptyDirectory(autoDismissDst))
        {
            Fail(L"Phase13_PostMortemDiagnostics could not create auto-dismiss test folders.");
            return true;
        }

        if (! WriteTestFile(autoDismissSrc / L"auto1.bin", 64))
        {
            Fail(L"Phase13_PostMortemDiagnostics could not create auto-dismiss test source file.");
            return true;
        }

        state.fileOps->SetAutoDismissSuccess(true);

        const FileSystemFlags copyFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        state.taskC                     = StartFileOperationAndGetId(state.fileOps,
                                                                     FILESYSTEM_COPY,
                                                                     FolderWindow::Pane::Left,
                                                                     FolderWindow::Pane::Right,
                                                                     state.fsLocal,
                                                                     {autoDismissSrc / L"auto1.bin"},
                                                                     autoDismissDst,
                                                                     copyFlags,
                                                                     false);
        if (! state.taskC.has_value())
        {
            Fail(L"Phase13_PostMortemDiagnostics could not start auto-dismiss enabled copy.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto itAutoDismissOn = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
        if (itAutoDismissOn == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itAutoDismissOn->second.hr))
        {
            Fail(std::format(L"Phase13_PostMortemDiagnostics auto-dismiss enabled copy failed: 0x{:08X}.",
                             static_cast<unsigned long>(itAutoDismissOn->second.hr)));
            return true;
        }

        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        for (const auto& summary : summaries)
        {
            if (summary.taskId == state.taskC.value())
            {
                Fail(L"Phase13_PostMortemDiagnostics auto-dismiss enabled task was not auto-dismissed.");
                return true;
            }
        }

        const std::filesystem::path autoDismissSrc = state.tempRoot / L"phase13-auto-dismiss-src";
        const std::filesystem::path autoDismissDst = state.tempRoot / L"phase13-auto-dismiss-dst";
        if (! WriteTestFile(autoDismissSrc / L"auto2.bin", 64))
        {
            Fail(L"Phase13_PostMortemDiagnostics could not create second auto-dismiss test source file.");
            return true;
        }

        state.fileOps->SetAutoDismissSuccess(false);

        const FileSystemFlags copyFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        state.taskA                     = StartFileOperationAndGetId(state.fileOps,
                                                                     FILESYSTEM_COPY,
                                                                     FolderWindow::Pane::Left,
                                                                     FolderWindow::Pane::Right,
                                                                     state.fsLocal,
                                                                     {autoDismissSrc / L"auto2.bin"},
                                                                     autoDismissDst,
                                                                     copyFlags,
                                                                     false);
        if (! state.taskA.has_value())
        {
            Fail(L"Phase13_PostMortemDiagnostics could not start auto-dismiss disabled copy.");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        const auto itAutoDismissOff = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (itAutoDismissOff == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itAutoDismissOff->second.hr))
        {
            Fail(std::format(L"Phase13_PostMortemDiagnostics auto-dismiss disabled copy failed: 0x{:08X}.",
                             static_cast<unsigned long>(itAutoDismissOff->second.hr)));
            return true;
        }

        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        bool foundRetained = false;
        for (const auto& summary : summaries)
        {
            if (summary.taskId == state.taskA.value())
            {
                foundRetained = true;
                break;
            }
        }

        if (! foundRetained)
        {
            Fail(L"Phase13_PostMortemDiagnostics auto-dismiss disabled task was unexpectedly removed.");
            return true;
        }

        // Enabling auto-dismiss should immediately remove already-completed success tasks.
        state.fileOps->SetAutoDismissSuccess(true);
        summaries.clear();
        state.fileOps->CollectCompletedTasks(summaries);
        for (const auto& summary : summaries)
        {
            if (summary.taskId == state.taskA.value())
            {
                Fail(L"Phase13_PostMortemDiagnostics enabling auto-dismiss did not remove the existing success task.");
                return true;
            }
        }

        // Auto-dismiss should also apply to canceled tasks.
        if (! state.fsDummy || state.dummyPaths.empty())
        {
            Fail(L"Phase13_PostMortemDiagnostics missing FileSystemDummy for auto-dismiss cancellation test.");
            return true;
        }

        if (! EnsureDummyFolderExists(state.fsDummy.get(), L"/dest-auto-cancel"))
        {
            Fail(L"Phase13_PostMortemDiagnostics could not create dummy destination folder for cancellation test.");
            return true;
        }

        const FileSystemFlags cancelFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                                         FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        state.taskB                       = StartFileOperationAndGetId(state.fileOps,
                                                                       FILESYSTEM_COPY,
                                                                       FolderWindow::Pane::Left,
                                                                       FolderWindow::Pane::Right,
                                                                       state.fsDummy,
                                                                       {std::filesystem::path(state.dummyPaths.front())},
                                                                       std::filesystem::path(L"/dest-auto-cancel"),
                                                                       cancelFlags,
                                                                       false);
        if (! state.taskB.has_value())
        {
            Fail(L"Phase13_PostMortemDiagnostics could not start cancelable dummy copy task.");
            return true;
        }

        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());
        if (! taskB)
        {
            return false;
        }

        if (taskB->_preCalcInProgress.load(std::memory_order_acquire) || taskB->HasEnteredOperation() || taskB->HasStarted())
        {
            taskB->RequestCancel();
            state.stepState = 5;
        }
        return false;
    }

    if (state.stepState == 5)
    {
        const auto itCancel = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (itCancel == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT hrCancel = itCancel->second.hr;
        if (hrCancel != HRESULT_FROM_WIN32(ERROR_CANCELLED) && hrCancel != E_ABORT)
        {
            Fail(std::format(L"Phase13_PostMortemDiagnostics expected cancelled task hr, got 0x{:08X}.", static_cast<unsigned long>(hrCancel)));
            return true;
        }

        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        for (const auto& summary : summaries)
        {
            if (summary.taskId == state.taskB.value())
            {
                Fail(L"Phase13_PostMortemDiagnostics cancelled task was not auto-dismissed.");
                return true;
            }
        }

        NextStep(state, SelfTestState::Step::Phase14_PopupHostLifetimeGuard);
        return false;
    }

    return false;
}
