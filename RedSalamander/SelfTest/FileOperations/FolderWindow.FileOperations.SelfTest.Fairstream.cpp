    // Operation Fairstream Phase 1 cases (data-loss correctness).
    // Included into the selftest switch in FolderWindow.FileOperations.SelfTest.cpp.

case SelfTestState::Step::Riptide_MoveDistinctSameSizeFilePreservesSource:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    const auto restoreForcedFallbackEnv = [&]() noexcept
    {
        if (state.forceMoveCopyFallbackEnvBackedUp)
        {
            static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(),
                                                      state.forceMoveCopyFallbackEnvHadOriginal ? state.forceMoveCopyFallbackEnvOriginal.c_str() : nullptr));
            state.forceMoveCopyFallbackEnvBackedUp = false;
            state.forceMoveCopyFallbackEnvHadOriginal = false;
            state.forceMoveCopyFallbackEnvOriginal.clear();
        }
    };

    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        restoreForcedFallbackEnv();
        Fail(L"Riptide_MoveDistinctSameSizeFilePreservesSource timed out.");
        return true;
    }

    constexpr size_t kFileBytes = 16 * 1024;
    const std::filesystem::path srcRoot = state.tempRoot / L"riptide-same-size-move-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"riptide-same-size-move-dst";
    const std::filesystem::path refRoot = state.tempRoot / L"riptide-same-size-move-ref";
    const std::filesystem::path srcFoo = srcRoot / L"Foo";
    const std::filesystem::path dstFoo = dstRoot / L"Foo";
    const std::filesystem::path srcFile = srcFoo / L"same.bin";
    const std::filesystem::path dstFile = dstFoo / L"same.bin";
    const std::filesystem::path srcRef = refRoot / L"source.bin";
    const std::filesystem::path dstRef = refRoot / L"destination.bin";

    const auto sourceStillIntact = [&]() noexcept { return FileSizeEquals(srcFile, kFileBytes) && FilesEqualBytes(srcFile, srcRef); };
    const auto destinationStillIntact = [&]() noexcept { return FileSizeEquals(dstFile, kFileBytes) && FilesEqualBytes(dstFile, dstRef); };

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(refRoot))
        {
            Fail(L"Riptide same-size move test failed to reset directories.");
            return true;
        }
        if (! WriteFilledTestFile(srcFile, kFileBytes, 0x31) || ! WriteFilledTestFile(srcRef, kFileBytes, 0x31) ||
            ! WriteFilledTestFile(dstFile, kFileBytes, 0xA7) || ! WriteFilledTestFile(dstRef, kFileBytes, 0xA7))
        {
            Fail(L"Riptide same-size move test failed to seed distinct source/destination files.");
            return true;
        }

        std::error_code ec;
        const auto sharedWriteTime = std::filesystem::file_time_type::clock::now() - std::chrono::seconds(4);
        std::filesystem::last_write_time(srcFile, sharedWriteTime, ec);
        if (ec)
        {
            Fail(L"Riptide same-size move test failed to set source mtime.");
            return true;
        }
        ec.clear();
        std::filesystem::last_write_time(dstFile, sharedWriteTime, ec);
        if (ec)
        {
            Fail(L"Riptide same-size move test failed to set destination mtime.");
            return true;
        }

        if (! state.forceMoveCopyFallbackEnvBackedUp)
        {
            state.forceMoveCopyFallbackEnvOriginal = GetEnvVarTrimmed(kSelfTestEnvForceMoveCopyFallback);
            state.forceMoveCopyFallbackEnvHadOriginal = ! state.forceMoveCopyFallbackEnvOriginal.empty();
            state.forceMoveCopyFallbackEnvBackedUp = true;
        }
        if (! SetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(), L"1"))
        {
            restoreForcedFallbackEnv();
            Fail(L"Riptide same-size move test failed to enable forced move-copy fallback.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFoo},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            restoreForcedFallbackEnv();
            Fail(L"Riptide same-size move test failed to start the move task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::Exists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
                {
                    restoreForcedFallbackEnv();
                    Fail(L"Riptide same-size move test expected an Exists prompt with Skip.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.markerTick = nowTick;
                state.stepState = 2;
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        restoreForcedFallbackEnv();
        if (! sourceStillIntact())
        {
            Fail(std::format(L"Riptide same-size move completed without a conflict and deleted or changed the source (hr=0x{:08X}, prompts={}).",
                             static_cast<unsigned long>(completed->second.hr),
                             completed->second.conflictPromptCount));
            return true;
        }

        Fail(std::format(L"Riptide same-size move completed without the required conflict prompt (hr=0x{:08X}, prompts={}).",
                         static_cast<unsigned long>(completed->second.hr),
                         completed->second.conflictPromptCount));
        return true;
    }

    if (state.stepState == 2)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        restoreForcedFallbackEnv();
        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(
                std::format(L"Riptide same-size move expected ERROR_PARTIAL_COPY after Skip, got 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 1u)
        {
            Fail(std::format(L"Riptide same-size move expected exactly 1 prompt, saw {}.", completed->second.conflictPromptCount));
            return true;
        }
        if (! sourceStillIntact())
        {
            Fail(L"Riptide same-size move did not preserve the skipped source file.");
            return true;
        }
        if (! destinationStillIntact())
        {
            Fail(L"Riptide same-size move changed the skipped destination file.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.RiptideSameSizeMovePreserved", L"shape=same-size-mtime-distinct-bytes", 1, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Riptide_ReparseMoveRollbackKeepsOverwrittenDestination);
        return false;
    }

    return false;
}
case SelfTestState::Step::Riptide_ReparseMoveRollbackKeepsOverwrittenDestination:
{
    const ULONGLONG nowTick = GetTickCount64();

    const auto restoreFailEnv = [&]() noexcept
    {
        if (state.reparseMoveSourceDeleteFailEnvBackedUp)
        {
            static_cast<void>(
                SetEnvironmentVariableW(kSelfTestEnvReparseMoveSourceDeleteFailPath.data(),
                                        state.reparseMoveSourceDeleteFailEnvHadOriginal ? state.reparseMoveSourceDeleteFailEnvOriginal.c_str() : nullptr));
            state.reparseMoveSourceDeleteFailEnvBackedUp = false;
            state.reparseMoveSourceDeleteFailEnvHadOriginal = false;
            state.reparseMoveSourceDeleteFailEnvOriginal.clear();
        }
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvReparseMoveSourceDeleteFailFired.data(), nullptr));
    };

    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        restoreFailEnv();
        Fail(L"Riptide_ReparseMoveRollbackKeepsOverwrittenDestination timed out.");
        return true;
    }

    const std::filesystem::path targetDir = state.tempRoot / L"riptide-reparse-rollback-target";
    const std::filesystem::path srcRoot = state.tempRoot / L"riptide-reparse-rollback-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"riptide-reparse-rollback-dst";
    const std::filesystem::path srcLink = srcRoot / L"Link";
    const std::filesystem::path dstLink = dstRoot / L"Link";
    const std::filesystem::path targetFile = targetDir / L"target-sentinel.bin";

    const auto sourceLinkStillIntact = [&]() noexcept
    {
        const auto target = TryGetDirectoryReparseTargetAbsolute(srcLink);
        return target.has_value() && NormalizePathForCompare(target.value()) == NormalizePathForCompare(targetDir.wstring());
    };
    const auto destinationLinkStillIntact = [&]() noexcept
    {
        const auto target = TryGetDirectoryReparseTargetAbsolute(dstLink);
        return target.has_value() && NormalizePathForCompare(target.value()) == NormalizePathForCompare(targetDir.wstring());
    };

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(targetDir) || ! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Riptide reparse rollback test failed to reset directories.");
            return true;
        }
        if (! WriteTestFile(targetFile, 4 * 1024) || ! WriteTestFile(dstLink, 2 * 1024))
        {
            Fail(L"Riptide reparse rollback test failed to seed target/destination sentinels.");
            return true;
        }
        if (! TryCreateJunction(srcLink, targetDir))
        {
            Fail(L"Riptide reparse rollback test failed to create the source junction.");
            return true;
        }

        if (! state.reparseMoveSourceDeleteFailEnvBackedUp)
        {
            state.reparseMoveSourceDeleteFailEnvOriginal = GetEnvVarTrimmed(kSelfTestEnvReparseMoveSourceDeleteFailPath);
            state.reparseMoveSourceDeleteFailEnvHadOriginal = ! state.reparseMoveSourceDeleteFailEnvOriginal.empty();
            state.reparseMoveSourceDeleteFailEnvBackedUp = true;
        }
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvReparseMoveSourceDeleteFailFired.data(), nullptr));
        if (! SetEnvironmentVariableW(kSelfTestEnvReparseMoveSourceDeleteFailPath.data(), srcLink.c_str()))
        {
            restoreFailEnv();
            Fail(L"Riptide reparse rollback test failed to enable source-delete failure injection.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        FileOpsRecursiveProgressRecorder callback{};
        const HRESULT hr = state.fsLocal->MoveItem(srcLink.c_str(), dstLink.c_str(), flags, nullptr, &callback, nullptr);

        const bool injectionFired = GetEnvVarTrimmed(kSelfTestEnvReparseMoveSourceDeleteFailFired) == L"1";
        restoreFailEnv();

        if (! injectionFired)
        {
            Fail(L"Riptide reparse rollback test did not trigger the source-delete failure injection.");
            return true;
        }
        if (! sourceLinkStillIntact())
        {
            Fail(L"Riptide reparse rollback did not preserve the source junction after injected delete failure.");
            return true;
        }
        if (! destinationLinkStillIntact())
        {
            Fail(L"Riptide reparse rollback deleted the overwritten destination after source delete failure.");
            return true;
        }
        if (! FileSizeEquals(targetFile, 4 * 1024))
        {
            Fail(L"Riptide reparse rollback damaged the out-of-tree junction target sentinel.");
            return true;
        }
        if (hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Riptide reparse rollback expected ERROR_PARTIAL_COPY after source delete failure, got 0x{:08X}.",
                             static_cast<unsigned long>(hr)));
            return true;
        }
        if (callback.completedCount != 1u)
        {
            Fail(std::format(L"Riptide reparse rollback expected exactly 1 item completion callback, saw {}.", callback.completedCount));
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.RiptideReparseRollbackPreservedDestination", L"shape=junction-over-existing-file-source-delete-fail", 1, 0, 0, hr);
        NextStep(state, SelfTestState::Step::Fairstream_MoveFallbackPreservesUncopiedNewFiles);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_MoveFallbackPreservesUncopiedNewFiles:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"Fairstream_MoveFallbackPreservesUncopiedNewFiles timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-move-window-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-move-window-dst";
    const std::filesystem::path srcFoo = srcRoot / L"Foo";
    const std::filesystem::path dstFoo = dstRoot / L"Foo";
    const std::filesystem::path srcDone = srcFoo / L"aaa_done";
    const std::filesystem::path dstDone = dstFoo / L"aaa_done";
    const std::filesystem::path srcSeed = srcDone / L"seed.bin";
    const std::filesystem::path dstSeed = dstDone / L"seed.bin";
    const std::filesystem::path srcBig = srcFoo / L"big.bin";
    const std::filesystem::path dstBig = dstFoo / L"big.bin";
    const std::filesystem::path srcNew = srcDone / L"zz_new.bin";
    const std::filesystem::path dstNew = dstDone / L"zz_new.bin";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Move-window test failed to reset directories.");
            return true;
        }
        if (! WriteTestFile(srcSeed, 4 * 1024) || ! WriteTestFile(srcBig, 4ull * 1024ull * 1024ull))
        {
            Fail(L"Move-window test failed to seed the source tree.");
            return true;
        }

        if (! state.forceMoveCopyFallbackEnvBackedUp)
        {
            state.forceMoveCopyFallbackEnvOriginal = GetEnvVarTrimmed(kSelfTestEnvForceMoveCopyFallback);
            state.forceMoveCopyFallbackEnvHadOriginal = ! state.forceMoveCopyFallbackEnvOriginal.empty();
            state.forceMoveCopyFallbackEnvBackedUp = true;
        }
        if (! SetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(), L"1"))
        {
            Fail(L"Move-window test failed to enable forced move-copy fallback.");
            return true;
        }

        // Throttle to ~512 KB/s so the 4 MB transfer leaves a multi-second copy window.
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFoo},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 512ull * 1024ull,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Move-window test failed to start the move task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        const bool alreadyCompleted = state.taskA.has_value() && state.completedTasks.contains(state.taskA.value());

        std::error_code ec;
        const bool copyUnderway = std::filesystem::exists(dstBig, ec) && ! ec && std::filesystem::exists(dstSeed, ec) && ! ec;
        if (alreadyCompleted)
        {
            Fail(L"Move-window test completed before the new file could be injected (copy window too small).");
            return true;
        }
        if (! copyUnderway)
        {
            return false;
        }

        if (state.markerTick == 0)
        {
            state.markerTick = nowTick;
            return false;
        }
        if ((nowTick - state.markerTick) < 300ull)
        {
            return false;
        }

        // aaa_done was fully enumerated and copied; a file created in it now exists only in the source.
        if (! WriteTestFile(srcNew, 2 * 1024))
        {
            Fail(L"Move-window test failed to inject the new source file.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
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
            state.forceMoveCopyFallbackEnvBackedUp = false;
            state.forceMoveCopyFallbackEnvHadOriginal = false;
            state.forceMoveCopyFallbackEnvOriginal.clear();
        }

        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Move-window test expected ERROR_PARTIAL_COPY (source preserved), got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        std::error_code ec;
        if (! std::filesystem::exists(srcNew, ec) || ec)
        {
            Fail(L"Move-window test PERMANENTLY LOST the file created during the copy window.");
            return true;
        }
        ec.clear();
        if (std::filesystem::exists(dstNew, ec))
        {
            Fail(L"Move-window test unexpectedly copied the injected file.");
            return true;
        }
        ec.clear();
        if (! FileSizeEquals(dstBig, 4ull * 1024ull * 1024ull) || ! FileSizeEquals(dstSeed, 4 * 1024))
        {
            Fail(L"Move-window test destination tree is incomplete.");
            return true;
        }
        if (std::filesystem::exists(srcBig, ec) && ! ec)
        {
            Fail(L"Move-window test did not delete the verified-copied source file.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamMoveWindowPreserved", L"shape=new-file-during-fallback-copy", 1, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_MoveMergeReadonlyDestinationFolder);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_MoveMergeReadonlyDestinationFolder:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Fairstream_MoveMergeReadonlyDestinationFolder timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-ro-merge-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-ro-merge-dst";
    const std::filesystem::path srcFoo = srcRoot / L"Foo";
    const std::filesystem::path dstFoo = dstRoot / L"Foo";

    if (state.stepState == 0)
    {
        static_cast<void>(SetFileAttributesW(dstFoo.c_str(), FILE_ATTRIBUTE_NORMAL));
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Readonly-merge test failed to reset directories.");
            return true;
        }
        if (! WriteTestFile(srcFoo / L"a.bin", 8 * 1024) || ! WriteTestFile(dstFoo / L"keep.bin", 2 * 1024))
        {
            Fail(L"Readonly-merge test failed to seed trees.");
            return true;
        }
        if (! SetFileAttributesW(dstFoo.c_str(), FILE_ATTRIBUTE_READONLY | FILE_ATTRIBUTE_DIRECTORY))
        {
            Fail(L"Readonly-merge test failed to mark the destination folder read-only.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFoo},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Readonly-merge test failed to start the move task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                Fail(L"Readonly-merge move raised a conflict prompt for a read-only destination folder merge.");
                return true;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        const DWORD attributes = GetFileAttributesW(dstFoo.c_str());
        const auto restoreAttributes = wil::scope_exit([&]() noexcept { static_cast<void>(SetFileAttributesW(dstFoo.c_str(), FILE_ATTRIBUTE_NORMAL)); });

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Readonly-merge move failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 0u)
        {
            Fail(std::format(L"Readonly-merge move expected zero prompts, saw {}.", completed->second.conflictPromptCount));
            return true;
        }
        if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_READONLY) == 0)
        {
            Fail(L"Readonly-merge move stripped the destination folder's read-only attribute.");
            return true;
        }

        std::error_code ec;
        if (std::filesystem::exists(srcFoo, ec) && ! ec)
        {
            Fail(L"Readonly-merge move left the source folder behind.");
            return true;
        }
        if (! FileSizeEquals(dstFoo / L"a.bin", 8 * 1024) || ! FileSizeEquals(dstFoo / L"keep.bin", 2 * 1024))
        {
            Fail(L"Readonly-merge move destination integrity check failed.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamReadonlyMergePromptCount", L"", completed->second.conflictPromptCount, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_OverwriteGrantIsOneShot);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_OverwriteGrantIsOneShot:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"Fairstream_OverwriteGrantIsOneShot timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-oneshot-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-oneshot-dst";
    const std::filesystem::path srcFoo = srcRoot / L"Foo";
    const std::filesystem::path dstFoo = dstRoot / L"Foo";

    // Eight colliding children across four workers force every worker to process more than one
    // conflict: a per-worker sticky grant would collapse the prompt count to the worker count.
    constexpr unsigned int kCollidingFileCount = 8u;
    const auto collidingName = [](unsigned int index) noexcept { return std::format(L"f{}.bin", index); };

    const auto seedDestination = [&]() noexcept -> bool
    {
        if (! RecreateEmptyDirectory(dstRoot))
        {
            return false;
        }
        for (unsigned int index = 1; index <= kCollidingFileCount; ++index)
        {
            if (! WriteTestFile(dstFoo / collidingName(index), 1024))
            {
                return false;
            }
        }
        return true;
    };

    const auto destinationMatchesSource = [&]() noexcept -> bool
    {
        for (unsigned int index = 1; index <= kCollidingFileCount; ++index)
        {
            if (! FilesEqualBytes(srcFoo / collidingName(index), dstFoo / collidingName(index)))
            {
                return false;
            }
        }
        return true;
    };

    const auto startCopyTask = [&]() noexcept -> std::optional<std::uint64_t>
    {
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        return StartFileOperationAndGetId(state.fileOps,
                                          FILESYSTEM_COPY,
                                          FolderWindow::Pane::Left,
                                          FolderWindow::Pane::Right,
                                          state.fsLocal,
                                          {srcFoo},
                                          dstRoot,
                                          flags,
                                          false,
                                          0,
                                          FolderWindow::FileOperationState::ExecutionMode::PerItem);
    };

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! seedDestination())
        {
            Fail(L"One-shot grant test failed to reset directories.");
            return true;
        }
        for (unsigned int index = 1; index <= kCollidingFileCount; ++index)
        {
            if (! WriteTestFile(srcFoo / collidingName(index), (8u + index) * 1024u))
            {
                Fail(L"One-shot grant test failed to seed the source tree.");
                return true;
            }
        }

        state.taskA = startCopyTask();
        if (! state.taskA.has_value())
        {
            Fail(L"One-shot grant test failed to start the first copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        // Answer every Exists prompt with a plain Overwrite (apply-to-all unchecked). Each
        // colliding child must raise its own prompt: a single answer never authorizes more.
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::Exists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"One-shot grant test expected an Exists prompt with Overwrite.");
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"One-shot grant copy failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != kCollidingFileCount)
        {
            Fail(std::format(L"One-shot grant test expected exactly {} prompts (one per colliding file), saw {}.",
                             kCollidingFileCount,
                             completed->second.conflictPromptCount));
            return true;
        }
        if (! destinationMatchesSource())
        {
            Fail(L"One-shot grant test destination integrity check failed after per-file overwrites.");
            return true;
        }

        // Control run: Apply-to-all checked on the first prompt must remain a single prompt.
        if (! seedDestination())
        {
            Fail(L"One-shot grant test failed to reseed the destination for the apply-to-all run.");
            return true;
        }
        state.taskB = startCopyTask();
        if (! state.taskB.has_value())
        {
            Fail(L"One-shot grant test failed to start the apply-to-all copy task.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (Task* task = state.fileOps && state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, true);
                return false;
            }
        }

        const auto completed = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Apply-to-all overwrite copy failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 1u)
        {
            Fail(std::format(L"Apply-to-all overwrite expected exactly 1 prompt, saw {}.", completed->second.conflictPromptCount));
            return true;
        }
        if (! destinationMatchesSource())
        {
            Fail(L"Apply-to-all overwrite destination integrity check failed.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamOneShotPromptCounts", L"", kCollidingFileCount, 1, 0, S_OK);
        NextStep(state, SelfTestState::Step::Fairstream_ConflictPromptsSerializedUnderParallelism);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_ConflictPromptsSerializedUnderParallelism:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Fairstream_ConflictPromptsSerializedUnderParallelism timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-serialize-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-serialize-dst";
    const std::filesystem::path srcFoo = srcRoot / L"Foo";
    const std::filesystem::path dstFoo = dstRoot / L"Foo";
    constexpr std::array<size_t, 4> kSourceBytes{{
        10u * 1024u,
        11u * 1024u,
        12u * 1024u,
        13u * 1024u,
    }};
    constexpr std::array<size_t, 4> kDestinationBytes{{
        1024u,
        2048u,
        3072u,
        4096u,
    }};
    const auto conflictName = [](size_t index) { return std::format(L"d{}.bin", index + 1u); };
    const auto conflictIndexForLeaf = [&](std::wstring_view leaf) noexcept -> std::optional<size_t>
    {
        for (size_t index = 0; index < kSourceBytes.size(); ++index)
        {
            if (leaf == conflictName(index))
            {
                return index;
            }
        }
        return std::nullopt;
    };
    const auto shouldOverwriteConflict = [](size_t index) noexcept { return index == 0u || index == 2u; };

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Prompt-serialization test failed to reset directories.");
            return true;
        }

        for (size_t index = 0; index < kSourceBytes.size(); ++index)
        {
            const std::wstring name = conflictName(index);
            if (! WriteTestFile(srcFoo / name, kSourceBytes[index]) || ! WriteTestFile(dstFoo / name, kDestinationBytes[index]))
            {
                Fail(L"Prompt-serialization test failed to seed colliding trees.");
                return true;
            }
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFoo},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Prompt-serialization test failed to start the copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        // Parallel workers hit all four collisions near-simultaneously. Decisions are deliberately
        // mixed and keyed by the prompted child, so a misrouted answer changes that child's final
        // byte count instead of hiding behind four identical Skip answers.
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::Exists)
                {
                    Fail(L"Prompt-serialization test expected only Exists prompts.");
                    return true;
                }

                std::wstring_view promptLeaf = prompt->destinationPath;
                if (const size_t slash = promptLeaf.find_last_of(L"\\/"); slash != std::wstring_view::npos)
                {
                    promptLeaf = promptLeaf.substr(slash + 1);
                }

                const std::optional<size_t> conflictIndex = conflictIndexForLeaf(promptLeaf);
                if (! conflictIndex.has_value())
                {
                    Fail(std::format(L"Prompt-serialization test saw an unexpected conflict destination '{}'.", prompt->destinationPath));
                    return true;
                }

                const Task::ConflictAction action =
                    shouldOverwriteConflict(conflictIndex.value()) ? Task::ConflictAction::Overwrite : Task::ConflictAction::Skip;
                task->SubmitConflictDecision(action, false);
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Prompt-serialization test expected ERROR_PARTIAL_COPY after mixed Overwrite/Skip decisions, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 4u)
        {
            Fail(std::format(L"Prompt-serialization test expected exactly 4 prompts, saw {}.", completed->second.conflictPromptCount));
            return true;
        }

        for (size_t index = 0; index < kSourceBytes.size(); ++index)
        {
            const std::wstring name = conflictName(index);
            const size_t expectedBytes = shouldOverwriteConflict(index) ? kSourceBytes[index] : kDestinationBytes[index];
            if (! FileSizeEquals(dstFoo / name, expectedBytes))
            {
                Fail(std::format(L"Prompt-serialization test: '{}' ended with the wrong byte count for its own routed decision.", name));
                return true;
            }
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamPromptSerialization", L"", completed->second.conflictPromptCount, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_ReparseReplaceNonEmptyDirRequiresConsent);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_ReparseReplaceNonEmptyDirRequiresConsent:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Fairstream_ReparseReplaceNonEmptyDirRequiresConsent timed out.");
        return true;
    }

    const std::filesystem::path targetDir = state.tempRoot / L"fairstream-junction-target";
    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-junction-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-junction-dst";
    const std::filesystem::path srcFoo = srcRoot / L"Foo";
    const std::filesystem::path srcLink = srcFoo / L"Link";
    const std::filesystem::path dstFoo = dstRoot / L"Foo";
    const std::filesystem::path dstLink = dstFoo / L"Link";
    const std::filesystem::path dstInner = dstLink / L"inner.bin";

    const auto seedTrees = [&]() noexcept -> bool
    {
        if (! RecreateEmptyDirectory(targetDir) || ! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            return false;
        }
        if (! WriteTestFile(targetDir / L"t.bin", 1024) || ! WriteTestFile(dstInner, 2 * 1024))
        {
            return false;
        }
        std::error_code ec;
        std::filesystem::create_directories(srcFoo, ec);
        if (ec)
        {
            return false;
        }
        return TryCreateJunction(srcLink, targetDir);
    };

    const auto startCopyTask = [&]() noexcept -> std::optional<std::uint64_t>
    {
        // Operation-wide overwrite is exactly the dangerous grant: it must still not authorize
        // replacing a non-empty directory with a reparse point without its own prompt.
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        return StartFileOperationAndGetId(state.fileOps,
                                          FILESYSTEM_COPY,
                                          FolderWindow::Pane::Left,
                                          FolderWindow::Pane::Right,
                                          state.fsLocal,
                                          {srcFoo},
                                          dstRoot,
                                          flags,
                                          false,
                                          0,
                                          FolderWindow::FileOperationState::ExecutionMode::PerItem);
    };

    if (state.stepState == 0)
    {
        if (! seedTrees())
        {
            Fail(L"Reparse-consent test failed to seed trees (junction creation requires no admin; check temp root).");
            return true;
        }

        state.taskA = startCopyTask();
        if (! state.taskA.has_value())
        {
            Fail(L"Reparse-consent test failed to start the first copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::NonEmptyDirectory || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Reparse-consent test expected a NonEmptyDirectory prompt offering Overwrite.");
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Reparse-consent Skip run expected ERROR_PARTIAL_COPY, got 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 1u)
        {
            Fail(std::format(L"Reparse-consent Skip run expected exactly 1 prompt, saw {}.", completed->second.conflictPromptCount));
            return true;
        }

        std::error_code ec;
        if (! std::filesystem::exists(dstInner, ec) || ec)
        {
            Fail(L"Reparse-consent Skip run DESTROYED the non-empty destination directory.");
            return true;
        }

        // Second run: an explicit Overwrite answer for this path authorizes the replacement.
        state.taskB = startCopyTask();
        if (! state.taskB.has_value())
        {
            Fail(L"Reparse-consent test failed to start the second copy task.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (Task* task = state.fileOps && state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
                return false;
            }
        }

        const auto completed = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Reparse-consent Overwrite run failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        const auto linkTarget = TryGetDirectoryReparseTargetAbsolute(dstLink);
        if (! linkTarget.has_value() || NormalizePathForCompare(linkTarget.value()) != NormalizePathForCompare(targetDir.wstring()))
        {
            Fail(L"Reparse-consent Overwrite run did not replace the directory with the junction.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamReparseConsent", L"", 1, 1, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate);
        return false;
    }

    return false;
}
case SelfTestState::Step::Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate timed out.");
        return true;
    }

    const std::filesystem::path targetDir = state.tempRoot / L"floodgate-bucket-target";
    const std::filesystem::path srcRoot = state.tempRoot / L"floodgate-bucket-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"floodgate-bucket-dst";
    const std::filesystem::path srcPlain = srcRoot / L"plain.txt";
    const std::filesystem::path dstPlain = dstRoot / L"plain.txt";
    const std::filesystem::path srcLink = srcRoot / L"RiskLink";
    const std::filesystem::path dstLink = dstRoot / L"RiskLink";
    const std::filesystem::path dstInner = dstLink / L"inner.bin";

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(),
                                                 R"json({"reparsePointPolicy":"copyReparse","concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        if (! RecreateEmptyDirectory(targetDir) || ! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(dstLink))
        {
            Fail(L"Conflict-bucket apply-to-all test failed to prepare directories.");
            return true;
        }

        if (! WriteFilledTestFile(targetDir / L"sentinel.bin", 4096, 0xA1) || ! WriteFilledTestFile(srcPlain, 2048, 0x5A) ||
            ! WriteFilledTestFile(dstPlain, 1024, 0x22) || ! WriteFilledTestFile(dstInner, 512, 0x39))
        {
            Fail(L"Conflict-bucket apply-to-all test failed to seed files.");
            return true;
        }

        if (! TryCreateJunction(srcLink, targetDir))
        {
            Fail(L"Conflict-bucket apply-to-all test failed to create source junction.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcPlain, srcLink},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Conflict-bucket apply-to-all test failed to start copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::Exists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Conflict-bucket apply-to-all test expected the first prompt to be a plain Exists overwrite.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, true);
                state.stepState = 2;
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed != state.completedTasks.end())
        {
            Fail(std::format(L"Conflict-bucket apply-to-all test completed before the first prompt: 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        return false;
    }

    if (state.stepState == 2)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::NonEmptyDirectory || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite) ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
                {
                    Fail(L"Conflict-bucket apply-to-all test expected a separate NonEmptyDirectory prompt after Apply to all.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.stepState = 3;
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed != state.completedTasks.end())
        {
            Fail(std::format(L"Conflict-bucket apply-to-all test completed without prompting for the riskier bucket: 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        return false;
    }

    if (state.stepState == 3)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Conflict-bucket apply-to-all test expected ERROR_PARTIAL_COPY after skipping the directory conflict, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 2u)
        {
            Fail(std::format(L"Conflict-bucket apply-to-all test expected exactly 2 prompts, saw {}.", completed->second.conflictPromptCount));
            return true;
        }
        if (! FilesEqualBytes(srcPlain, dstPlain))
        {
            Fail(L"Conflict-bucket apply-to-all test did not apply the cached Exists overwrite to the plain file.");
            return true;
        }

        std::error_code ec;
        if (! std::filesystem::exists(dstInner, ec) || ec)
        {
            Fail(L"Conflict-bucket apply-to-all test destroyed the skipped non-empty destination directory.");
            return true;
        }
        const DWORD dstAttributes = GetFileAttributesW(dstLink.c_str());
        if (dstAttributes == INVALID_FILE_ATTRIBUTES || WI_IsFlagSet(dstAttributes, FILE_ATTRIBUTE_REPARSE_POINT) ||
            ! WI_IsFlagSet(dstAttributes, FILE_ATTRIBUTE_DIRECTORY))
        {
            Fail(L"Conflict-bucket apply-to-all test converted the skipped destination into a reparse point.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateConflictBucketApplyAll", L"", 2, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent);
        return false;
    }

    return false;
}
case SelfTestState::Step::Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent timed out.");
        return true;
    }

    const std::filesystem::path targetDir = state.tempRoot / L"riptide-empty-dir-reparse-target";
    const std::filesystem::path srcRoot = state.tempRoot / L"riptide-empty-dir-reparse-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"riptide-empty-dir-reparse-dst";
    const std::filesystem::path refRoot = state.tempRoot / L"riptide-empty-dir-reparse-ref";
    const std::filesystem::path srcFoo = srcRoot / L"Foo";
    const std::filesystem::path srcLink = srcFoo / L"Link";
    const std::filesystem::path srcSibling = srcFoo / L"sibling.bin";
    const std::filesystem::path dstFoo = dstRoot / L"Foo";
    const std::filesystem::path dstLink = dstFoo / L"Link";
    const std::filesystem::path dstSibling = dstFoo / L"sibling.bin";
    const std::filesystem::path targetSentinel = targetDir / L"sentinel.bin";
    const std::filesystem::path targetRef = refRoot / L"sentinel.bin";

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"copyReparse"})json"));

        if (! RecreateEmptyDirectory(targetDir) || ! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) ||
            ! RecreateEmptyDirectory(refRoot) || ! RecreateEmptyDirectory(dstLink))
        {
            Fail(L"Empty-dir reparse consent test failed to prepare directories.");
            return true;
        }

        std::error_code ec;
        std::filesystem::create_directories(srcFoo, ec);
        if (ec)
        {
            Fail(L"Empty-dir reparse consent test failed to create source folder.");
            return true;
        }

        if (! WriteFilledTestFile(targetSentinel, 4096, 0xC3) || ! WriteFilledTestFile(targetRef, 4096, 0xC3) || ! WriteFilledTestFile(srcSibling, 8192, 0x5D))
        {
            Fail(L"Empty-dir reparse consent test failed to seed files.");
            return true;
        }

        if (! TryCreateJunction(srcLink, targetDir))
        {
            Fail(L"Empty-dir reparse consent test failed to create source junction.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFoo},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Empty-dir reparse consent test failed to start copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::Exists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip) ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Empty-dir reparse consent test expected an Exists prompt offering Skip and Overwrite.");
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Empty-dir reparse consent Skip run expected ERROR_PARTIAL_COPY, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 1u)
        {
            Fail(std::format(L"Empty-dir reparse consent Skip run expected exactly 1 prompt, saw {}.", completed->second.conflictPromptCount));
            return true;
        }

        const DWORD dstLinkAttributes = ::GetFileAttributesW(dstLink.c_str());
        if (dstLinkAttributes == INVALID_FILE_ATTRIBUTES || (dstLinkAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0 ||
            (dstLinkAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0)
        {
            Fail(L"Empty-dir reparse consent Skip run converted the destination real directory into a reparse point.");
            return true;
        }

        if (TryGetDirectoryReparseTargetAbsolute(dstLink).has_value())
        {
            Fail(L"Empty-dir reparse consent Skip run left the destination resolving as a junction.");
            return true;
        }

        if (! FilesEqualBytes(targetSentinel, targetRef))
        {
            Fail(L"Empty-dir reparse consent Skip run changed the out-of-tree junction target sentinel.");
            return true;
        }

        if (! FilesEqualBytes(srcSibling, dstSibling))
        {
            Fail(L"Empty-dir reparse consent Skip run did not preserve/copy the non-conflicting sibling bytes.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.RiptideEmptyDirReparseConsent", L"", completed->second.conflictPromptCount, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Riptide_ReparseMoveSourceNeverUsesDirectoryMerge);
        return false;
    }

    return false;
}
case SelfTestState::Step::Riptide_ReparseMoveSourceNeverUsesDirectoryMerge:
{
    const std::filesystem::path root = state.tempRoot / L"riptide-reparse-move-merge";
    const std::filesystem::path targetDir = root / L"target";
    const std::filesystem::path srcRoot = root / L"src";
    const std::filesystem::path dstRoot = root / L"dst";
    const std::filesystem::path refRoot = root / L"ref";
    const std::filesystem::path srcLink = srcRoot / L"Link";
    const std::filesystem::path dstLink = dstRoot / L"Link";
    const std::filesystem::path targetSentinel = targetDir / L"sentinel.bin";
    const std::filesystem::path targetRef = refRoot / L"sentinel.bin";

    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(targetDir) || ! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstLink) ||
        ! RecreateEmptyDirectory(refRoot))
    {
        Fail(L"Reparse move merge test failed to prepare directories.");
        return true;
    }

    if (! WriteFilledTestFile(targetSentinel, 4096, 0xA4) || ! WriteFilledTestFile(targetRef, 4096, 0xA4))
    {
        Fail(L"Reparse move merge test failed to seed target sentinel.");
        return true;
    }

    if (! TryCreateJunction(srcLink, targetDir))
    {
        Fail(L"Reparse move merge test failed to create source junction.");
        return true;
    }

    if (! SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"followTargets"})json"))
    {
        Fail(L"Reparse move merge test failed to set followTargets policy.");
        return true;
    }
    auto restoreReparsePolicy = wil::scope_exit([&]() noexcept
    { static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"copyReparse"})json")); });

    const HRESULT moveHr =
        state.fsLocal->MoveItem(srcLink.c_str(), dstLink.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), nullptr, nullptr, nullptr);
    if (moveHr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
    {
        Fail(std::format(L"Reparse move merge test expected ERROR_ALREADY_EXISTS, got 0x{:08X}.", static_cast<unsigned long>(moveHr)));
        return true;
    }

    if (! FilesEqualBytes(targetSentinel, targetRef))
    {
        Fail(L"Reparse move merge test moved or changed data through the source junction target.");
        return true;
    }

    std::error_code ec;
    if (std::filesystem::exists(dstLink / L"sentinel.bin", ec))
    {
        Fail(L"Reparse move merge test renamed target data into the destination real directory.");
        return true;
    }

    const DWORD srcLinkAttributes = ::GetFileAttributesW(srcLink.c_str());
    if (srcLinkAttributes == INVALID_FILE_ATTRIBUTES || (srcLinkAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0)
    {
        Fail(L"Reparse move merge test removed the source junction after a rejected merge.");
        return true;
    }

    const DWORD dstLinkAttributes = ::GetFileAttributesW(dstLink.c_str());
    if (dstLinkAttributes == INVALID_FILE_ATTRIBUTES || (dstLinkAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0 ||
        (dstLinkAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0)
    {
        Fail(L"Reparse move merge test did not preserve the destination as a real directory.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.RiptideReparseMoveNoMerge", L"policy=followTargets", 1, 0, 0, moveHr);
    NextStep(state, SelfTestState::Step::Fairstream_MovedTreeRetargetsInternalLinks);
    return false;
}
case SelfTestState::Step::Fairstream_MovedTreeRetargetsInternalLinks:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Fairstream_MovedTreeRetargetsInternalLinks timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-move-links-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-move-links-dst";
    const std::filesystem::path srcTree = srcRoot / L"Tree";
    const std::filesystem::path srcData = srcTree / L"Data";
    const std::filesystem::path srcLink = srcTree / L"LinkToData";
    const std::filesystem::path dstTree = dstRoot / L"Tree";
    const std::filesystem::path dstData = dstTree / L"Data";
    const std::filesystem::path dstLink = dstTree / L"LinkToData";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Link-retarget move test failed to reset directories.");
            return true;
        }
        if (! WriteTestFile(srcData / L"d.bin", 4 * 1024))
        {
            Fail(L"Link-retarget move test failed to seed the source tree.");
            return true;
        }
        if (! TryCreateJunction(srcLink, srcData))
        {
            Fail(L"Link-retarget move test failed to create the intra-tree junction.");
            return true;
        }

        if (! state.forceMoveCopyFallbackEnvBackedUp)
        {
            state.forceMoveCopyFallbackEnvOriginal = GetEnvVarTrimmed(kSelfTestEnvForceMoveCopyFallback);
            state.forceMoveCopyFallbackEnvHadOriginal = ! state.forceMoveCopyFallbackEnvOriginal.empty();
            state.forceMoveCopyFallbackEnvBackedUp = true;
        }
        if (! SetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(), L"1"))
        {
            Fail(L"Link-retarget move test failed to enable forced move-copy fallback.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcTree},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Link-retarget move test failed to start the move task.");
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
            state.forceMoveCopyFallbackEnvBackedUp = false;
            state.forceMoveCopyFallbackEnvHadOriginal = false;
            state.forceMoveCopyFallbackEnvOriginal.clear();
        }

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Link-retarget move failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        const auto linkTarget = TryGetDirectoryReparseTargetAbsolute(dstLink);
        if (! linkTarget.has_value())
        {
            Fail(L"Link-retarget move did not produce a junction at the destination.");
            return true;
        }
        if (NormalizePathForCompare(linkTarget.value()) != NormalizePathForCompare(dstData.wstring()))
        {
            Fail(std::format(L"Link-retarget move left the junction dangling: target='{}' expected='{}'.", linkTarget.value(), dstData.wstring()));
            return true;
        }

        std::error_code ec;
        if (std::filesystem::exists(srcTree, ec) && ! ec)
        {
            Fail(L"Link-retarget move left the source tree behind.");
            return true;
        }
        if (! FileSizeEquals(dstData / L"d.bin", 4 * 1024))
        {
            Fail(L"Link-retarget move destination integrity check failed.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamMoveLinkRetarget", L"", 1, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_CrossFsConcurrentMoveUsesBridge);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_CrossFsConcurrentMoveUsesBridge:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Fairstream_CrossFsConcurrentMoveUsesBridge timed out.");
        return true;
    }

    const std::filesystem::path localDir = state.tempRoot / L"fairstream-crossfs-src";
    const std::filesystem::path returnDir = state.tempRoot / L"fairstream-crossfs-return";
    const std::filesystem::path refDir = state.tempRoot / L"fairstream-crossfs-ref";
    const std::wstring dummyRoot = L"/fairstream-crossfs-move";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(localDir) || ! RecreateEmptyDirectory(returnDir) || ! RecreateEmptyDirectory(refDir))
        {
            Fail(L"Cross-FS concurrent move test failed to reset directories.");
            return true;
        }
        if (! WriteTestFile(localDir / L"m1.bin", 8 * 1024) || ! WriteTestFile(localDir / L"m2.bin", 12 * 1024))
        {
            Fail(L"Cross-FS concurrent move test failed to seed files.");
            return true;
        }

        std::error_code ec;
        std::filesystem::copy_file(localDir / L"m1.bin", refDir / L"m1.bin", ec);
        if (! ec)
        {
            std::filesystem::copy_file(localDir / L"m2.bin", refDir / L"m2.bin", ec);
        }
        if (ec)
        {
            Fail(L"Cross-FS concurrent move test failed to snapshot reference copies.");
            return true;
        }

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
        {
            Fail(L"Cross-FS concurrent move test failed to create the dummy destination folder.");
            return true;
        }

        // Two items + PerItem mode drives the concurrent path; the move must run as bridge
        // copy + source delete, never as a foreign-namespace MoveItem on the local plugin.
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localDir / L"m1.bin", localDir / L"m2.bin"},
                                                 std::filesystem::path(dummyRoot),
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            Fail(L"Cross-FS concurrent move test failed to start the outbound move task.");
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

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Cross-FS concurrent outbound move failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        std::error_code ec;
        if (std::filesystem::exists(localDir / L"m1.bin", ec) || std::filesystem::exists(localDir / L"m2.bin", ec))
        {
            Fail(L"Cross-FS concurrent move left source files behind after a successful move.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyRoot) / L"m1.bin", std::filesystem::path(dummyRoot) / L"m2.bin"},
                                                 returnDir,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskB.has_value())
        {
            Fail(L"Cross-FS concurrent move test failed to start the return move task.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto completed = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Cross-FS concurrent return move failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        if (! FilesEqualBytes(refDir / L"m1.bin", returnDir / L"m1.bin") || ! FilesEqualBytes(refDir / L"m2.bin", returnDir / L"m2.bin"))
        {
            Fail(L"Cross-FS concurrent move round-trip integrity check failed.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamCrossFsConcurrentMove", L"shape=2-item-roundtrip", 2, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_RerunCopyIsResumeAware);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_RerunCopyIsResumeAware:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"Fairstream_RerunCopyIsResumeAware timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-rerun-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-rerun-dst";
    const std::filesystem::path srcFoo = srcRoot / L"Foo";
    const std::filesystem::path dstFoo = dstRoot / L"Foo";

    const auto startCopyTask = [&]() noexcept -> std::optional<std::uint64_t>
    {
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        return StartFileOperationAndGetId(state.fileOps,
                                          FILESYSTEM_COPY,
                                          FolderWindow::Pane::Left,
                                          FolderWindow::Pane::Right,
                                          state.fsLocal,
                                          {srcFoo},
                                          dstRoot,
                                          flags,
                                          false,
                                          0,
                                          FolderWindow::FileOperationState::ExecutionMode::PerItem);
    };

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Re-run resume test failed to reset directories.");
            return true;
        }
        if (! WriteTestFile(srcFoo / L"a.bin", 8 * 1024) || ! WriteTestFile(srcFoo / L"b.bin", 12 * 1024) ||
            ! WriteTestFile(srcFoo / L"nested" / L"c.bin", 16 * 1024))
        {
            Fail(L"Re-run resume test failed to seed the source tree.");
            return true;
        }

        state.taskA = startCopyTask();
        if (! state.taskA.has_value())
        {
            Fail(L"Re-run resume test failed to start the initial copy.");
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
        if (FAILED(completed->second.hr) || completed->second.conflictPromptCount != 0u)
        {
            Fail(L"Re-run resume test initial copy did not complete cleanly.");
            return true;
        }

        // Re-running the identical copy must not flood the user with 'already exists' prompts:
        // unchanged destinations are recognized as artifacts of the earlier pass.
        state.taskB = startCopyTask();
        if (! state.taskB.has_value())
        {
            Fail(L"Re-run resume test failed to start the re-run copy.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (Task* task = state.fileOps && state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr)
        {
            if (TryGetConflictPromptCopy(task).has_value())
            {
                Fail(L"Re-run resume test raised a prompt for an identical, already-copied file.");
                return true;
            }
        }

        const auto completed = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Re-run resume copy should report host-visible skipped items: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 0u)
        {
            Fail(std::format(L"Re-run resume copy expected zero prompts for identical files, saw {}.", completed->second.conflictPromptCount));
            return true;
        }

        // A genuinely changed child must still prompt — resume must never auto-overwrite content.
        if (! WriteTestFile(srcFoo / L"b.bin", 20 * 1024))
        {
            Fail(L"Re-run resume test failed to modify the changed child.");
            return true;
        }

        state.taskA = startCopyTask();
        if (! state.taskA.has_value())
        {
            Fail(L"Re-run resume test failed to start the changed-child copy.");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                const std::wstring expected = NormalizePathForCompare((srcFoo / L"b.bin").wstring());
                if (NormalizePathForCompare(prompt->sourcePath) != expected)
                {
                    Fail(std::format(L"Re-run resume test prompted for the wrong child: '{}'.", prompt->sourcePath));
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Re-run resume changed-child copy should report its byte-identical skipped siblings: 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 1u)
        {
            Fail(std::format(L"Re-run resume changed-child copy expected exactly 1 prompt, saw {}.", completed->second.conflictPromptCount));
            return true;
        }
        if (! FilesEqualBytes(srcFoo / L"b.bin", dstFoo / L"b.bin"))
        {
            Fail(L"Re-run resume changed-child copy did not refresh the changed destination.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamRerunResume", L"", 0, 1, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_MoveFallbackSameSizeCollisionPrompts);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_MoveFallbackSameSizeCollisionPrompts:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Fairstream_MoveFallbackSameSizeCollisionPrompts timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-same-size-collision-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-same-size-collision-dst";
    const std::filesystem::path srcFile = srcRoot / L"collide.bin";
    const std::filesystem::path dstFile = dstRoot / L"collide.bin";
    constexpr size_t kBytes = 24ull * 1024ull;

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Same-size collision test failed to reset directories.");
            return true;
        }
        if (! WriteTestFile(srcFile, kBytes) || ! WriteTestFile(dstFile, kBytes))
        {
            Fail(L"Same-size collision test failed to seed source/destination files.");
            return true;
        }

        wil::unique_handle sourceHandle(CreateFileW(srcFile.c_str(),
                                                    FILE_READ_ATTRIBUTES,
                                                    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                    nullptr,
                                                    OPEN_EXISTING,
                                                    FILE_ATTRIBUTE_NORMAL,
                                                    nullptr));
        if (! sourceHandle)
        {
            Fail(L"Same-size collision test failed to reopen the source file.");
            return true;
        }

        FILETIME sourceWrite{};
        if (! GetFileTime(sourceHandle.get(), nullptr, nullptr, &sourceWrite))
        {
            Fail(L"Same-size collision test failed to capture the source timestamp.");
            return true;
        }

        wil::unique_handle destinationHandle(CreateFileW(dstFile.c_str(),
                                                         GENERIC_READ | GENERIC_WRITE | FILE_WRITE_ATTRIBUTES,
                                                         FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                         nullptr,
                                                         OPEN_EXISTING,
                                                         FILE_ATTRIBUTE_NORMAL,
                                                         nullptr));
        if (! destinationHandle)
        {
            Fail(L"Same-size collision test failed to reopen the destination file.");
            return true;
        }

        LARGE_INTEGER zero{};
        if (! SetFilePointerEx(destinationHandle.get(), zero, nullptr, FILE_BEGIN))
        {
            Fail(L"Same-size collision test failed to seek the destination file.");
            return true;
        }

        const unsigned char differentByte = 0xC3u;
        DWORD written = 0;
        if (! WriteFile(destinationHandle.get(), &differentByte, 1u, &written, nullptr) || written != 1u)
        {
            Fail(L"Same-size collision test failed to mutate the destination bytes.");
            return true;
        }
        if (! SetFileTime(destinationHandle.get(), nullptr, nullptr, &sourceWrite))
        {
            Fail(L"Same-size collision test failed to align the destination timestamp.");
            return true;
        }
        destinationHandle.reset();

        if (FilesEqualBytes(srcFile, dstFile))
        {
            Fail(L"Same-size collision test failed to create byte-different same-size files.");
            return true;
        }

        if (! state.forceMoveCopyFallbackEnvBackedUp)
        {
            state.forceMoveCopyFallbackEnvOriginal = GetEnvVarTrimmed(kSelfTestEnvForceMoveCopyFallback);
            state.forceMoveCopyFallbackEnvHadOriginal = ! state.forceMoveCopyFallbackEnvOriginal.empty();
            state.forceMoveCopyFallbackEnvBackedUp = true;
        }
        if (! SetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(), L"1"))
        {
            Fail(L"Same-size collision test failed to enable forced move-copy fallback.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFile},
                                                 dstRoot,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Same-size collision test failed to start the move task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                const std::wstring expectedSource = NormalizePathForCompare(srcFile.wstring());
                if (prompt->bucket != Task::ConflictBucket::Exists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Same-size collision test expected an Exists prompt offering Overwrite.");
                    return true;
                }
                if (NormalizePathForCompare(prompt->sourcePath) != expectedSource)
                {
                    Fail(std::format(L"Same-size collision test prompted for the wrong source path: '{}'.", prompt->sourcePath));
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.stepState = 2;
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed != state.completedTasks.end())
        {
            if (state.forceMoveCopyFallbackEnvBackedUp)
            {
                static_cast<void>(
                    SetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(),
                                            state.forceMoveCopyFallbackEnvHadOriginal ? state.forceMoveCopyFallbackEnvOriginal.c_str() : nullptr));
                state.forceMoveCopyFallbackEnvBackedUp = false;
                state.forceMoveCopyFallbackEnvHadOriginal = false;
                state.forceMoveCopyFallbackEnvOriginal.clear();
            }

            Fail(std::format(L"Same-size collision move completed before raising a conflict prompt (hr=0x{:08X}, prompts={}).",
                             static_cast<unsigned long>(completed->second.hr),
                             completed->second.conflictPromptCount));
            return true;
        }

        return false;
    }

    if (state.stepState == 2)
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
            state.forceMoveCopyFallbackEnvBackedUp = false;
            state.forceMoveCopyFallbackEnvHadOriginal = false;
            state.forceMoveCopyFallbackEnvOriginal.clear();
        }

        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Same-size collision move expected ERROR_PARTIAL_COPY after Skip, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 1u)
        {
            Fail(std::format(L"Same-size collision move expected exactly 1 prompt, saw {}.", completed->second.conflictPromptCount));
            return true;
        }

        std::error_code ec;
        if (! std::filesystem::exists(srcFile, ec) || ec)
        {
            Fail(L"Same-size collision move deleted the source after Skip.");
            return true;
        }
        ec.clear();
        if (! std::filesystem::exists(dstFile, ec) || ec)
        {
            Fail(L"Same-size collision move removed the existing destination.");
            return true;
        }
        if (! FileSizeEquals(srcFile, kBytes) || ! FileSizeEquals(dstFile, kBytes))
        {
            Fail(L"Same-size collision move changed file sizes unexpectedly.");
            return true;
        }
        if (FilesEqualBytes(srcFile, dstFile))
        {
            Fail(L"Same-size collision move overwrote the existing destination despite Skip.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamSameSizeCollisionPrompt", L"", completed->second.conflictPromptCount, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_GraphBandsFairUnderParallelCopy);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_GraphBandsFairUnderParallelCopy:
{
    FileOperationsPopupInternal::PopupLayoutDebugSnapshot graph{};
    if (! DebugBuildFileOperationsGraphFairnessHistorySnapshot(graph))
    {
        Fail(std::format(L"Graph fairness deterministic sampler failed (multi={}, single={}, distinct={}, min={:.3f}, max={:.3f}, accumCalls={}, pending={}, "
                         L"maxStreams={}).",
                         graph.graphMultiHueBucketCount,
                         graph.graphSingleHueBucketCount,
                         graph.graphDistinctHueCount,
                         graph.graphMinHueShare,
                         graph.graphMaxHueShare,
                         graph.graphDebugAccumulateCalls,
                         graph.graphDebugLastPending,
                         graph.graphDebugMaxStreams));
        return true;
    }

    if (graph.graphDebugMaxStreams != 4u)
    {
        Fail(std::format(L"Graph fairness deterministic sampler expected 4 concurrent streams, saw {}.", graph.graphDebugMaxStreams));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamGraphBandFairness",
                      L"shape=4-equal-streams-deterministic-history",
                      graph.graphMultiHueBucketCount,
                      static_cast<uint64_t>(graph.graphMinHueShare * 1000.0),
                      static_cast<uint64_t>(graph.graphMaxHueShare * 1000.0),
                      S_OK);
    NextStep(state, SelfTestState::Step::Fairstream_MergeMoveRenamesChildren);
    return false;
}
case SelfTestState::Step::Fairstream_MergeMoveRenamesChildren:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Fairstream_MergeMoveRenamesChildren timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-mergemove-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-mergemove-dst";
    const std::filesystem::path srcTree = srcRoot / L"tree";
    const std::filesystem::path dstTree = dstRoot / L"tree";

    const std::array<std::filesystem::path, 3> sourceFiles = {
        srcTree / L"dirA" / L"file1.bin",
        srcTree / L"dirA" / L"sub" / L"file2.bin",
        srcTree / L"file3.bin",
    };
    const std::array<std::filesystem::path, 3> destinationFiles = {
        dstTree / L"dirA" / L"file1.bin",
        dstTree / L"dirA" / L"sub" / L"file2.bin",
        dstTree / L"file3.bin",
    };

    const auto fileIdOf = [](const std::filesystem::path& path) noexcept -> uint64_t
    {
        wil::unique_hfile handle(CreateFileW(path.c_str(),
                                             FILE_READ_ATTRIBUTES,
                                             FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                             nullptr,
                                             OPEN_EXISTING,
                                             FILE_FLAG_BACKUP_SEMANTICS,
                                             nullptr));
        if (! handle)
        {
            return 0;
        }
        BY_HANDLE_FILE_INFORMATION info{};
        if (! GetFileInformationByHandle(handle.get(), &info))
        {
            return 0;
        }
        return (static_cast<uint64_t>(info.nFileIndexHigh) << 32) | static_cast<uint64_t>(info.nFileIndexLow);
    };

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(srcTree / L"dirA" / L"sub") ||
            ! RecreateEmptyDirectory(dstTree / L"dirA"))
        {
            Fail(L"Merge-move rename test failed to prepare directories.");
            return true;
        }
        if (! WriteTestFile(sourceFiles[0], 256 * 1024) || ! WriteTestFile(sourceFiles[1], 64 * 1024) || ! WriteTestFile(sourceFiles[2], 128 * 1024) ||
            ! WriteTestFile(dstTree / L"keep.bin", 8 * 1024))
        {
            Fail(L"Merge-move rename test failed to seed files.");
            return true;
        }

        for (size_t i = 0; i < sourceFiles.size(); ++i)
        {
            state.fairstreamMergeFileIds[i] = fileIdOf(sourceFiles[i]);
            if (state.fairstreamMergeFileIds[i] == 0)
            {
                Fail(L"Merge-move rename test failed to capture source file ids.");
                return true;
            }
        }

        // No conflict answerer runs in this case: any prompt would hang the task and time the
        // case out, so completion alone proves the merge stayed prompt-free.
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcTree},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Merge-move rename test failed to start the move task.");
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

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Merge-move rename task failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        std::error_code ec;
        if (std::filesystem::exists(srcTree, ec) || ec)
        {
            Fail(L"Merge-move rename test left the source tree behind.");
            return true;
        }

        ec.clear();
        if (! std::filesystem::exists(dstTree / L"keep.bin", ec) || ec)
        {
            Fail(L"Merge-move rename test lost the pre-existing destination child.");
            return true;
        }

        for (size_t i = 0; i < destinationFiles.size(); ++i)
        {
            const uint64_t destinationId = fileIdOf(destinationFiles[i]);
            if (destinationId == 0)
            {
                Fail(std::format(L"Merge-move rename test: destination child {} is missing.", destinationFiles[i].wstring()));
                return true;
            }
            if (destinationId != state.fairstreamMergeFileIds[i])
            {
                Fail(std::format(L"Merge-move rename test: child {} was copied (file id changed), expected a rename.", destinationFiles[i].wstring()));
                return true;
            }
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamMergeMoveRename", L"shape=same-volume-merge-3-files", 3, 1, 0, S_OK);
        NextStep(state, SelfTestState::Step::Fairstream_JunctionMergeTargetRequiresConsent);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_JunctionMergeTargetRequiresConsent:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Fairstream_JunctionMergeTargetRequiresConsent timed out.");
        return true;
    }

    // The junction sits one level BELOW the top-level merge target: the host's top-level
    // pre-flight only probes the root, so the child merge in the engine is the only guard.
    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-junctionmerge-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-junctionmerge-dst";
    const std::filesystem::path linkTarget = state.tempRoot / L"fairstream-junctionmerge-target";
    const std::filesystem::path srcPayload = srcRoot / L"payload";
    const std::filesystem::path dstPayload = dstRoot / L"payload";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcPayload / L"sub") || ! RecreateEmptyDirectory(dstPayload) || ! RecreateEmptyDirectory(linkTarget))
        {
            Fail(L"Junction-merge test failed to prepare directories.");
            return true;
        }
        if (! WriteTestFile(srcPayload / L"sub" / L"inner.bin", 4 * 1024) || ! WriteTestFile(linkTarget / L"sentinel.bin", 1024))
        {
            Fail(L"Junction-merge test failed to seed files.");
            return true;
        }
        if (! TryCreateJunction(dstPayload / L"sub", linkTarget))
        {
            Fail(L"Junction-merge test failed to create the nested destination junction.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcPayload},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Junction-merge test failed to start the copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        // Grant Overwrite so the operation gets past the host's top-level pre-flight and into
        // the engine's child merge — that is where the junction guard lives.
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Junction-merge test copy failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        // The junction target must never receive merged children (write-through = data lands
        // outside the visible destination tree).
        std::error_code ec;
        if (std::filesystem::exists(linkTarget / L"inner.bin", ec))
        {
            Fail(L"Junction-merge test: merge wrote through the junction into its target.");
            return true;
        }

        ec.clear();
        if (! std::filesystem::exists(linkTarget / L"sentinel.bin", ec) || ec)
        {
            Fail(L"Junction-merge test: the junction target's own content was lost.");
            return true;
        }

        // With Overwrite granted, the engine must have REPLACED the junction with a real
        // directory and merged into that.
        const DWORD subAttributes = ::GetFileAttributesW((dstPayload / L"sub").c_str());
        if (subAttributes == INVALID_FILE_ATTRIBUTES || (subAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0 || (subAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0)
        {
            Fail(L"Junction-merge test: the destination junction was not replaced by a real directory.");
            return true;
        }

        ec.clear();
        if (! std::filesystem::exists(dstPayload / L"sub" / L"inner.bin", ec) || ec)
        {
            Fail(L"Junction-merge test: merged child is missing from the replaced directory.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamJunctionMergeConsent", L"", completed->second.conflictPromptCount, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_CopyIntoSelfAliasRejected);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_CopyIntoSelfAliasRejected:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        Fail(L"Fairstream_CopyIntoSelfAliasRejected timed out.");
        return true;
    }

    // dstAlias is a junction pointing INSIDE the source tree: the textual overlap test cannot
    // see it, so without canonicalization the copy recurses into itself.
    const std::filesystem::path root = state.tempRoot / L"fairstream-selfcopy";
    const std::filesystem::path srcTree = root / L"tree";
    const std::filesystem::path inside = srcTree / L"inside";
    const std::filesystem::path dstAlias = root / L"alias";

    if (! RecreateEmptyDirectory(inside))
    {
        Fail(L"Copy-into-self alias test failed to prepare directories.");
        return true;
    }
    if (! WriteTestFile(srcTree / L"file.bin", 1024))
    {
        Fail(L"Copy-into-self alias test failed to seed files.");
        return true;
    }
    if (! TryCreateJunction(dstAlias, inside))
    {
        Fail(L"Copy-into-self alias test failed to create the alias junction.");
        return true;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
    const std::optional<uint64_t> taskId = StartFileOperationAndGetId(state.fileOps,
                                                                      FILESYSTEM_COPY,
                                                                      FolderWindow::Pane::Left,
                                                                      FolderWindow::Pane::Right,
                                                                      state.fsLocal,
                                                                      {srcTree},
                                                                      dstAlias,
                                                                      flags,
                                                                      false,
                                                                      0,
                                                                      FolderWindow::FileOperationState::ExecutionMode::PerItem);
    if (taskId.has_value())
    {
        // Contain the runaway recursion before failing the case.
        if (FolderWindow::FileOperationState::Task* task = state.fileOps ? state.fileOps->FindTask(taskId.value()) : nullptr)
        {
            task->RequestCancel();
        }
        Fail(L"Copy-into-self alias test: the overlap guard accepted a destination aliased inside the source.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamCopyIntoSelfAlias", L"shape=junction-alias-into-source", 1, 0, 0, S_OK);
    NextStep(state, SelfTestState::Step::Fairstream_StorageKindProbed);
    return false;
}
case SelfTestState::Step::Fairstream_StorageKindProbed:
{
    const ULONGLONG nowTick = GetTickCount64();
    const auto clearAutoConcurrencyOverride = []() noexcept { SetFileOpsAutoConcurrencyOverrideForSelfTest(false, 1u, FILESYSTEM_STORAGE_UNKNOWN); };
    const auto cancelStorageClampTask = [&]() noexcept
    {
        if (state.taskA.has_value())
        {
            if (FolderWindow::FileOperationState::Task* task = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
            {
                task->RequestCancel();
            }
        }
    };

    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        clearAutoConcurrencyOverride();
        cancelStorageClampTask();
        Fail(L"Fairstream_StorageKindProbed timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-storage-clamp-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-storage-clamp-dst";
    const std::filesystem::path srcDir = srcRoot / L"payload";
    const std::filesystem::path nestedDir = srcDir / L"nested";
    constexpr int kFileCount = 4;
    constexpr size_t kFileBytes = 2ull * 1024ull * 1024ull;

    const auto fileName = [](int index) noexcept { return std::format(L"storage_clamp_{:02}.bin", index); };

    if (state.stepState == 0)
    {
        FileSystemStorageCharacteristics characteristics{};
        characteristics.sizeBytes = sizeof(characteristics);
        const std::wstring probePath = state.tempRoot.wstring();
        const HRESULT hr = state.fsLocal->GetStorageCharacteristics(probePath.c_str(), &characteristics);
        if (FAILED(hr))
        {
            Fail(std::format(L"Storage-kind probe failed: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }

        // A physical local volume must resolve to a real medium, not UNKNOWN.
        if (characteristics.storageKind == FILESYSTEM_STORAGE_UNKNOWN)
        {
            Fail(L"Storage-kind probe returned UNKNOWN for a physical local volume.");
            return true;
        }

        const bool rotational = (characteristics.flags & FILESYSTEM_STORAGE_FLAG_ROTATIONAL) != 0;
        if (rotational && characteristics.preferredCopyMoveConcurrency > 2u)
        {
            Fail(std::format(L"Seek-penalty media must clamp copy/move concurrency (kind={}, concurrency={}).",
                             characteristics.storageKind,
                             characteristics.preferredCopyMoveConcurrency));
            return true;
        }
        if (! rotational && characteristics.preferredCopyMoveConcurrency < 2u)
        {
            Fail(std::format(L"Non-rotational media should allow parallel transfers (kind={}, concurrency={}).",
                             characteristics.storageKind,
                             characteristics.preferredCopyMoveConcurrency));
            return true;
        }

        if (! RecreateEmptyDirectory(nestedDir) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Storage-kind clamp test failed to reset directories.");
            return true;
        }

        for (int i = 0; i < kFileCount; ++i)
        {
            if (! WriteTestFile(nestedDir / fileName(i), kFileBytes))
            {
                Fail(L"Storage-kind clamp test failed to seed source files.");
                return true;
            }
        }

        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"auto","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"copyReparse","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));
        SetFileOpsAutoConcurrencyOverrideForSelfTest(true, 1u, FILESYSTEM_STORAGE_HDD);
        state.storageClampMaxInFlight = 0;

        const FileSystemFlags flags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcDir},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 1ull * 1024ull * 1024ull,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            clearAutoConcurrencyOverride();
            Fail(L"Storage-kind clamp test failed to start the copy task.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamStorageKind",
                          probePath,
                          characteristics.storageKind,
                          characteristics.preferredCopyMoveConcurrency,
                          characteristics.flags,
                          S_OK);
        state.stepState = 1;
        return false;
    }

    if (state.taskA.has_value())
    {
        if (FolderWindow::FileOperationState::Task* task = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            std::scoped_lock lock(task->_inFlightFilesMutex);
            state.storageClampMaxInFlight = std::max(state.storageClampMaxInFlight, task->_inFlightFileCount);
        }
    }

    if (state.storageClampMaxInFlight > 1u)
    {
        clearAutoConcurrencyOverride();
        cancelStorageClampTask();
        Fail(std::format(L"Storage-kind clamp was not propagated into recursive copy: observed {} in-flight files with host budget 1.",
                         state.storageClampMaxInFlight));
        return true;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    clearAutoConcurrencyOverride();

    if (FAILED(completed->second.hr))
    {
        Fail(std::format(L"Storage-kind clamp copy failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    if (completed->second.configuredMaxConcurrency != 1u)
    {
        Fail(std::format(L"Storage-kind clamp expected host configured concurrency 1 but observed {}.", completed->second.configuredMaxConcurrency));
        return true;
    }

    if (state.storageClampMaxInFlight == 0u)
    {
        Fail(L"Storage-kind clamp copy completed before any in-flight file sample was observed.");
        return true;
    }

    const std::filesystem::path dstNested = dstRoot / srcDir.filename() / nestedDir.filename();
    for (int i = 0; i < kFileCount; ++i)
    {
        const std::filesystem::path copied = dstNested / fileName(i);
        if (! FileSizeEquals(copied, kFileBytes))
        {
            Fail(std::format(L"Storage-kind clamp test: {} did not copy intact.", copied.wstring()));
            return true;
        }
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamStorageClamp",
                      L"host-budget=1 recursive-copy",
                      completed->second.configuredMaxConcurrency,
                      state.storageClampMaxInFlight,
                      kFileCount,
                      S_OK);
    NextStep(state, SelfTestState::Step::Fairstream_ParallelDeleteContinuesPastLockedChild);
    return false;
}
case SelfTestState::Step::Fairstream_ParallelDeleteContinuesPastLockedChild:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        state.holdOpenHandle.reset();
        Fail(L"Fairstream_ParallelDeleteContinuesPastLockedChild timed out.");
        return true;
    }

    const std::filesystem::path root = state.tempRoot / L"fairstream-pardel-continue";
    const std::filesystem::path lockedFile = root / L"sub1" / L"locked.bin";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(root / L"sub1") || ! RecreateEmptyDirectory(root / L"sub2" / L"nested"))
        {
            Fail(L"Parallel-delete continue test failed to prepare directories.");
            return true;
        }

        bool seeded = WriteTestFile(lockedFile, 1024);
        for (int index = 1; seeded && index <= 5; ++index)
        {
            seeded = WriteTestFile(root / L"sub1" / std::format(L"a{}.bin", index), 1024) &&
                     WriteTestFile(root / L"sub2" / std::format(L"b{}.bin", index), 1024) &&
                     WriteTestFile(root / L"sub2" / L"nested" / std::format(L"c{}.bin", index), 1024);
        }
        if (! seeded)
        {
            Fail(L"Parallel-delete continue test failed to seed files.");
            return true;
        }

        // Exclusive open (no sharing): both enumeration-time and delete-time opens fail on it.
        state.holdOpenHandle.reset(CreateFileW(lockedFile.c_str(), GENERIC_READ, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! state.holdOpenHandle)
        {
            Fail(L"Parallel-delete continue test failed to lock the sentinel file.");
            return true;
        }

        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","deleteMaxConcurrency":8})json"));

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {root},
                                                 {},
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            state.holdOpenHandle.reset();
            Fail(L"Parallel-delete continue test failed to start the delete task.");
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

        state.holdOpenHandle.reset();

        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Parallel-delete continue test expected ERROR_PARTIAL_COPY, got 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        // The locked file (and its ancestor chain) survives; EVERYTHING else must be gone.
        std::error_code ec;
        if (! std::filesystem::exists(lockedFile, ec) || ec)
        {
            Fail(L"Parallel-delete continue test: the locked sentinel was deleted.");
            return true;
        }
        ec.clear();
        if (std::filesystem::exists(root / L"sub2", ec))
        {
            Fail(L"Parallel-delete continue test: unrelated subtree sub2 was not deleted.");
            return true;
        }
        for (int index = 1; index <= 5; ++index)
        {
            ec.clear();
            if (std::filesystem::exists(root / L"sub1" / std::format(L"a{}.bin", index), ec))
            {
                Fail(std::format(L"Parallel-delete continue test: sibling a{}.bin survived next to the locked file.", index));
                return true;
            }
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamParallelDeleteContinue", L"", 1, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_BridgePerFileConflictSkips);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_BridgePerFileConflictSkips:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Fairstream_BridgePerFileConflictSkips timed out.");
        return true;
    }

    // Cross-filesystem COPY (dummy -> local) of a DIRECTORY whose child collides with a
    // pre-existing local file. The host only sees the top-level directory item (no collision),
    // so the CHILD conflict is the bridge's job: it must raise a per-file conflict, honor Skip
    // without touching the existing child, copy the non-colliding sibling, and end PARTIAL.
    const std::wstring dummyDir = L"/fairstream-bridge-conflict-src";
    const std::wstring dummyCollide = dummyDir + L"/collide.txt";
    const std::wstring dummyKeep = dummyDir + L"/keep.txt";
    const std::filesystem::path localDest = state.tempRoot / L"fairstream-bridge-conflict-dst";
    const std::filesystem::path destSubdir = localDest / L"fairstream-bridge-conflict-src";
    const std::filesystem::path destCollide = destSubdir / L"collide.txt";
    const std::filesystem::path destKeep = destSubdir / L"keep.txt";
    constexpr size_t kSentinelBytes = 7u;

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyDir) || ! DummyWriteTextFile(state.fsDummy.get(), dummyCollide, "dummy-collide-content") ||
            ! DummyWriteTextFile(state.fsDummy.get(), dummyKeep, "dummy-keep-content"))
        {
            Fail(L"Bridge per-file conflict test failed to seed the dummy source tree.");
            return true;
        }
        // Pre-create only the colliding child under the merge destination; keep.txt must NOT
        // exist yet (it is the proof a sibling still copies past the skipped child).
        if (! RecreateEmptyDirectory(destSubdir) || ! WriteTestFile(destCollide, kSentinelBytes))
        {
            Fail(L"Bridge per-file conflict test failed to seed the colliding local child.");
            return true;
        }

        // Pin single-concurrency on the destination plugin so the bridge processes children
        // deterministically (within-folder budget clamps to the min of both endpoints): only
        // collide.txt can produce a conflict, regardless of leftover config from earlier cases.
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyDir)},
                                                 localDest,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            Fail(L"Bridge per-file conflict test failed to start the cross-FS copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::Exists)
                {
                    Fail(std::format(L"Bridge per-file conflict test expected an Exists conflict for the colliding destination file; bucket={} status=0x{:08X} "
                                     L"source='{}' destination='{}'.",
                                     static_cast<int>(prompt->bucket),
                                     static_cast<unsigned long>(prompt->status),
                                     prompt->sourcePath,
                                     prompt->destinationPath));
                    return true;
                }

                // The conflict must name the colliding CHILD (collide.txt), not the whole
                // top-level directory: that is the bridge's per-file conflict, not a
                // whole-directory host conflict. A path-prefix bug or whole-dir fail-closed
                // would surface the directory leaf here instead.
                std::wstring_view promptLeaf = prompt->destinationPath;
                if (const size_t slash = promptLeaf.find_last_of(L"\\/"); slash != std::wstring_view::npos)
                {
                    promptLeaf = promptLeaf.substr(slash + 1);
                }
                if (promptLeaf != L"collide.txt")
                {
                    Fail(std::format(
                        L"Bridge per-file conflict named '{}' instead of the colliding child 'collide.txt' (whole-directory conflict, not per-file).",
                        std::wstring(promptLeaf)));
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.markerTick = nowTick;
                state.stepState = 2;
                return false;
            }
        }

        const auto completedBeforePrompt = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completedBeforePrompt != state.completedTasks.end())
        {
            Fail(std::format(L"Bridge per-file conflict test completed before the expected child prompt: 0x{:08X}.",
                             static_cast<unsigned long>(completedBeforePrompt->second.hr)));
            return true;
        }
        return false;
    }

    if (state.stepState == 2)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > SelfTest::ScaleTimeout(5'000ull))
                {
                    Fail(std::format(
                        L"Bridge per-file conflict test still had an active prompt after Skip: bucket={} status=0x{:08X} source='{}' destination='{}'.",
                        static_cast<int>(prompt->bucket),
                        static_cast<unsigned long>(prompt->status),
                        prompt->sourcePath,
                        prompt->destinationPath));
                    return true;
                }
                return false;
            }
        }

        state.markerTick = 0;
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                Fail(std::format(L"Bridge per-file conflict test raised a second prompt after Skip: bucket={} status=0x{:08X} destination='{}'.",
                                 static_cast<int>(prompt->bucket),
                                 static_cast<unsigned long>(prompt->status),
                                 prompt->destinationPath));
                return true;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (completed->second.conflictPromptCount == 0u)
        {
            Fail(L"Bridge per-file conflict test: the bridge failed closed without raising a per-file conflict.");
            return true;
        }
        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Bridge per-file conflict test expected ERROR_PARTIAL_COPY after Skip, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        // Skip must leave the pre-existing colliding child byte-for-byte untouched.
        if (! FileSizeEquals(destCollide, kSentinelBytes))
        {
            Fail(L"Bridge per-file conflict test: the skipped child was overwritten.");
            return true;
        }

        // The non-colliding sibling must still have copied past the skipped child: this is what
        // distinguishes a per-CHILD bridge conflict from a whole-directory host conflict.
        std::error_code ec;
        if (! std::filesystem::exists(destKeep, ec) || ec)
        {
            Fail(L"Bridge per-file conflict test: the non-colliding sibling did not copy past the skipped child.");
            return true;
        }
        if (! FileSizeEquals(destKeep, std::string_view("dummy-keep-content").size()))
        {
            Fail(L"Bridge per-file conflict test: the copied sibling has the wrong byte count.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamBridgePerFileConflictSkip", L"", completed->second.conflictPromptCount, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Riptide_BridgeNestedDirVsFileSkipContinuesSiblings);
        return false;
    }

    return false;
}
case SelfTestState::Step::Riptide_BridgeNestedDirVsFileSkipContinuesSiblings:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Riptide_BridgeNestedDirVsFileSkipContinuesSiblings timed out.");
        return true;
    }

    const std::wstring dummyRoot = L"/riptide-bridge-dir-file-skip-src";
    const std::wstring dummyBlockedDir = dummyRoot + L"/blocked";
    const std::wstring dummyBlockedChild = dummyBlockedDir + L"/inside.txt";
    const std::wstring dummySibling = dummyRoot + L"/sibling.txt";
    const std::filesystem::path localDest = state.tempRoot / L"riptide-bridge-dir-file-skip-dst";
    const std::filesystem::path destRoot = localDest / L"riptide-bridge-dir-file-skip-src";
    const std::filesystem::path destBlocker = destRoot / L"blocked";
    const std::filesystem::path destSibling = destRoot / L"sibling.txt";
    constexpr size_t kBlockerBytes = 77u;
    constexpr std::string_view kSiblingBytes = "sibling-ok";

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyBlockedDir) ||
            ! DummyWriteTextFile(state.fsDummy.get(), dummyBlockedChild, "blocked-child") ||
            ! DummyWriteTextFile(state.fsDummy.get(), dummySibling, kSiblingBytes))
        {
            Fail(L"Riptide bridge dir-vs-file test failed to seed the dummy source tree.");
            return true;
        }

        if (! RecreateEmptyDirectory(destRoot) || ! WriteTestFile(destBlocker, kBlockerBytes))
        {
            Fail(L"Riptide bridge dir-vs-file test failed to seed the local blocking file.");
            return true;
        }

        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyRoot)},
                                                 localDest,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            Fail(L"Riptide bridge dir-vs-file test failed to start the cross-FS copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::Exists)
                {
                    Fail(std::format(L"Riptide bridge dir-vs-file test expected Exists, got bucket={} status=0x{:08X} destination='{}'.",
                                     static_cast<int>(prompt->bucket),
                                     static_cast<unsigned long>(prompt->status),
                                     prompt->destinationPath));
                    return true;
                }

                if (! EqualsIgnoreCase(prompt->destinationPath, destBlocker.wstring()))
                {
                    Fail(std::format(
                        L"Riptide bridge dir-vs-file test expected child blocker prompt '{}', got '{}'.", destBlocker.wstring(), prompt->destinationPath));
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.stepState = 2;
                return false;
            }
        }

        const auto completedBeforePrompt = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completedBeforePrompt != state.completedTasks.end())
        {
            Fail(std::format(L"Riptide bridge dir-vs-file test completed before the expected child prompt: 0x{:08X}.",
                             static_cast<unsigned long>(completedBeforePrompt->second.hr)));
            return true;
        }
        return false;
    }

    if (state.stepState == 2)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > SelfTest::ScaleTimeout(5'000ull))
                {
                    Fail(std::format(
                        L"Riptide bridge dir-vs-file test still had an active prompt after Skip: bucket={} status=0x{:08X} source='{}' destination='{}'.",
                        static_cast<int>(prompt->bucket),
                        static_cast<unsigned long>(prompt->status),
                        prompt->sourcePath,
                        prompt->destinationPath));
                    return true;
                }
                return false;
            }
        }

        state.markerTick = 0;
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                Fail(std::format(L"Riptide bridge dir-vs-file test raised a second prompt after Skip: bucket={} status=0x{:08X} destination='{}'.",
                                 static_cast<int>(prompt->bucket),
                                 static_cast<unsigned long>(prompt->status),
                                 prompt->destinationPath));
                return true;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Riptide bridge dir-vs-file test expected ERROR_PARTIAL_COPY after Skip, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (! FileSizeEquals(destBlocker, kBlockerBytes))
        {
            Fail(L"Riptide bridge dir-vs-file test: skipped blocking file was modified.");
            return true;
        }
        if (! FileSizeEquals(destSibling, kSiblingBytes.size()))
        {
            Fail(L"Riptide bridge dir-vs-file test: sibling did not copy after skipped directory blocker.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.RiptideBridgeDirVsFileSkip", L"", completed->second.conflictPromptCount, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Riptide_BridgeDirectoryOverReadOnlyFileRequiresReplaceReadOnly);
        return false;
    }

    return false;
}
case SelfTestState::Step::Riptide_BridgeDirectoryOverReadOnlyFileRequiresReplaceReadOnly:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    const std::wstring dummyRoot = L"/riptide-bridge-readonly-dir-src";
    const std::wstring dummyChild = dummyRoot + L"/child.txt";
    const std::filesystem::path localDest = state.tempRoot / L"riptide-bridge-readonly-dir-dst";
    const std::filesystem::path readOnlyPath = localDest / L"riptide-bridge-readonly-dir-src";
    const std::filesystem::path copiedChild = readOnlyPath / L"child.txt";
    constexpr std::string_view kChildBytes = "readonly-child";

    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
        Fail(L"Riptide_BridgeDirectoryOverReadOnlyFileRequiresReplaceReadOnly timed out.");
        return true;
    }

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! DummyWriteTextFile(state.fsDummy.get(), dummyChild, kChildBytes))
        {
            Fail(L"Riptide bridge read-only directory test failed to seed the dummy source tree.");
            return true;
        }

        if (! RecreateEmptyDirectory(localDest) || ! WriteTestFile(readOnlyPath, 5u))
        {
            Fail(L"Riptide bridge read-only directory test failed to seed the local read-only blocker.");
            return true;
        }
        if (! SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_READONLY))
        {
            Fail(L"Riptide bridge read-only directory test failed to mark the blocker read-only.");
            return true;
        }

        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyRoot)},
                                                 localDest,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
            Fail(L"Riptide bridge read-only directory test failed to start the cross-FS copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::ReadOnly || ! PromptHasAction(prompt.value(), Task::ConflictAction::ReplaceReadOnly))
                {
                    static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
                    Fail(
                        std::format(L"Riptide bridge read-only directory test expected ReplaceReadOnly prompt, got bucket={} status=0x{:08X} destination='{}'.",
                                    static_cast<int>(prompt->bucket),
                                    static_cast<unsigned long>(prompt->status),
                                    prompt->destinationPath));
                    return true;
                }

                if (! EqualsIgnoreCase(prompt->destinationPath, readOnlyPath.wstring()))
                {
                    static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
                    Fail(std::format(
                        L"Riptide bridge read-only directory test expected blocker prompt '{}', got '{}'.", readOnlyPath.wstring(), prompt->destinationPath));
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::ReplaceReadOnly, false);
                state.markerTick = nowTick;
                state.stepState = 2;
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed != state.completedTasks.end())
        {
            static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
            Fail(std::format(L"Riptide bridge read-only directory test completed before ReplaceReadOnly prompt: 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        return false;
    }

    if (state.stepState == 2)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > SelfTest::ScaleTimeout(5'000ull))
                {
                    static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
                    Fail(std::format(L"Riptide bridge read-only directory test still had an active prompt after ReplaceReadOnly: bucket={} status=0x{:08X} "
                                     L"source='{}' destination='{}'.",
                                     static_cast<int>(prompt->bucket),
                                     static_cast<unsigned long>(prompt->status),
                                     prompt->sourcePath,
                                     prompt->destinationPath));
                    return true;
                }
                return false;
            }
        }

        state.markerTick = 0;
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
                Fail(std::format(
                    L"Riptide bridge read-only directory test raised a second prompt after ReplaceReadOnly: bucket={} status=0x{:08X} destination='{}'.",
                    static_cast<int>(prompt->bucket),
                    static_cast<unsigned long>(prompt->status),
                    prompt->destinationPath));
                return true;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(completed->second.hr))
        {
            static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
            Fail(std::format(L"Riptide bridge read-only directory test expected success after ReplaceReadOnly, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        std::error_code ec;
        if (! std::filesystem::is_directory(readOnlyPath, ec) || ec)
        {
            static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
            Fail(L"Riptide bridge read-only directory test did not replace the blocking file with a directory.");
            return true;
        }
        if (! FileSizeEquals(copiedChild, kChildBytes.size()))
        {
            Fail(L"Riptide bridge read-only directory test did not copy the source child after replacement.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.RiptideBridgeReadOnlyDirReplace", L"", completed->second.conflictPromptCount, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial);
        return false;
    }

    return false;
}
case SelfTestState::Step::Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    const std::wstring dummyRoot = L"/riptide-bridge-create-race-src";
    const std::wstring dummyChild = dummyRoot + L"/child.txt";
    const std::filesystem::path localDest = state.tempRoot / L"riptide-bridge-create-race-dst";
    const std::filesystem::path racedPath = localDest / L"riptide-bridge-create-race-src";
    const std::filesystem::path copiedChild = racedPath / L"child.txt";
    constexpr std::string_view kChildBytes = "race-child";
    constexpr const wchar_t* kRaceEnv = L"REDSALAMANDER_FILEOPS_BRIDGE_CREATE_DIRECTORY_RACE_PATH";

    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        static_cast<void>(SetEnvironmentVariableW(kRaceEnv, nullptr));
        Fail(L"Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial timed out.");
        return true;
    }

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! DummyWriteTextFile(state.fsDummy.get(), dummyChild, kChildBytes))
        {
            Fail(L"Riptide bridge create-directory race test failed to seed the dummy source tree.");
            return true;
        }

        if (! RecreateEmptyDirectory(localDest))
        {
            Fail(L"Riptide bridge create-directory race test failed to seed the local destination parent.");
            return true;
        }

        static_cast<void>(SetEnvironmentVariableW(kRaceEnv, racedPath.c_str()));
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyRoot)},
                                                 localDest,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            static_cast<void>(SetEnvironmentVariableW(kRaceEnv, nullptr));
            Fail(L"Riptide bridge create-directory race test failed to start the cross-FS copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                static_cast<void>(SetEnvironmentVariableW(kRaceEnv, nullptr));
                if (prompt->bucket != Task::ConflictBucket::Exists)
                {
                    Fail(std::format(L"Riptide bridge create-directory race test expected Exists, got bucket={} status=0x{:08X} destination='{}'.",
                                     static_cast<int>(prompt->bucket),
                                     static_cast<unsigned long>(prompt->status),
                                     prompt->destinationPath));
                    return true;
                }
                if (! EqualsIgnoreCase(prompt->destinationPath, racedPath.wstring()))
                {
                    Fail(std::format(
                        L"Riptide bridge create-directory race test expected raced path prompt '{}', got '{}'.", racedPath.wstring(), prompt->destinationPath));
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.markerTick = nowTick;
                state.stepState = 2;
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed != state.completedTasks.end())
        {
            static_cast<void>(SetEnvironmentVariableW(kRaceEnv, nullptr));
            Fail(std::format(L"Riptide bridge create-directory race test completed before conflict prompt: 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        return false;
    }

    if (state.stepState == 2)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > SelfTest::ScaleTimeout(5'000ull))
                {
                    static_cast<void>(SetEnvironmentVariableW(kRaceEnv, nullptr));
                    Fail(std::format(L"Riptide bridge create-directory race test still had an active prompt after Skip: bucket={} status=0x{:08X} source='{}' "
                                     L"destination='{}'.",
                                     static_cast<int>(prompt->bucket),
                                     static_cast<unsigned long>(prompt->status),
                                     prompt->sourcePath,
                                     prompt->destinationPath));
                    return true;
                }
                return false;
            }
        }

        state.markerTick = 0;
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                static_cast<void>(SetEnvironmentVariableW(kRaceEnv, nullptr));
                Fail(std::format(L"Riptide bridge create-directory race test raised a second prompt after Skip: bucket={} status=0x{:08X} destination='{}'.",
                                 static_cast<int>(prompt->bucket),
                                 static_cast<unsigned long>(prompt->status),
                                 prompt->destinationPath));
                return true;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(std::format(L"Riptide bridge create-directory race test expected ERROR_PARTIAL_COPY after Skip, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (! FileSizeEquals(racedPath, 0u))
        {
            Fail(L"Riptide bridge create-directory race test expected the raced blocker file to remain empty.");
            return true;
        }
        std::error_code ec;
        if (std::filesystem::exists(copiedChild, ec) && ! ec)
        {
            Fail(L"Riptide bridge create-directory race test copied a child under a skipped raced file.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.RiptideBridgeCreateDirRace", L"", completed->second.conflictPromptCount, 0, 0, completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_SaturationConcurrentCopiesMakeProgress);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_SaturationConcurrentCopiesMakeProgress:
{
    // Fairstream 5E. Several sizeable directory trees copied together in one PerItem task each
    // drive their own recursive-copy job. The scheduler now dispatches queue items as short work
    // items (not long-lived consumer loops), so the concurrent jobs share the worker pool and all
    // make progress instead of the first tree pinning every worker until it finishes. The proof is
    // liveness + correctness + credibility: the task must complete, at least two distinct source
    // trees must appear in-flight in the same sample window, and each recursive copy must emit
    // robust dispatch metrics where dispatches materially exceed the fixed worker concurrency.
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"Fairstream_SaturationConcurrentCopiesMakeProgress timed out (a concurrent copy job stalled).");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-saturation-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-saturation-dst";
    constexpr unsigned int kTreeCount = 4u; // distinct top-level dirs => distinct copy jobs
    constexpr unsigned int kFilesPerTree = 32u;
    constexpr unsigned int kExpectedConcurrency = 4u;
    constexpr size_t kFileBytes = 64ull * 1024ull;

    const auto treeName = [](unsigned int index) { return std::format(L"tree{}", index); };
    const auto fileName = [](unsigned int index) { return std::format(L"file{}.bin", index); };
    const auto countDistinctInFlightTrees = [&](const FileOperationsPopupInternal::TaskSnapshot& snapshot)
    {
        std::array<bool, kTreeCount> seenTrees{};
        size_t distinctTrees = 0;
        const size_t count = std::min(snapshot.inFlightFileCount, snapshot.inFlightFiles.size());
        for (size_t i = 0; i < count; ++i)
        {
            const std::wstring& sourcePath = snapshot.inFlightFiles[i].sourcePath;
            for (unsigned int t = 0; t < kTreeCount; ++t)
            {
                if (! seenTrees[t] && sourcePath.find(treeName(t)) != std::wstring::npos)
                {
                    seenTrees[t] = true;
                    ++distinctTrees;
                }
            }
        }
        return distinctTrees;
    };

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Saturation test failed to reset directories.");
            return true;
        }

        std::vector<PerfMetricSample> dispatchMetrics;
        static_cast<void>(TryReadPerfMetricSamples("FileOps.CopyRecursiveParallel.WorkItemDispatches", dispatchMetrics));
        state.fairstreamSaturationDispatchMetricBaseline = dispatchMetrics.size();
        state.fairstreamSaturationMaxDistinctInFlightTrees = 0;

        // Pin local concurrency so the dispatch proof has a stable denominator: each recursive
        // copy must show dispatches substantially above four, not merely eventual completion.
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"copyReparse","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));

        std::vector<std::filesystem::path> sources;
        for (unsigned int t = 0; t < kTreeCount; ++t)
        {
            const std::filesystem::path tree = srcRoot / treeName(t);
            if (! RecreateEmptyDirectory(tree))
            {
                Fail(L"Saturation test failed to seed a source tree.");
                return true;
            }
            for (unsigned int f = 0; f < kFilesPerTree; ++f)
            {
                if (! WriteTestFile(tree / fileName(f), kFileBytes))
                {
                    Fail(L"Saturation test failed to seed source files.");
                    return true;
                }
            }
            sources.push_back(tree);
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 std::move(sources),
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 1ull * 1024ull * 1024ull,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Saturation test failed to start the concurrent directory copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (state.taskA.has_value())
        {
            FileOperationsPopupInternal::TaskSnapshot snapshot{};
            if (TryGetPopupTaskSnapshot(state.fileOps, state.taskA.value(), snapshot))
            {
                state.fairstreamSaturationMaxDistinctInFlightTrees =
                    std::max(state.fairstreamSaturationMaxDistinctInFlightTrees, countDistinctInFlightTrees(snapshot));
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Saturation concurrent copy failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        // Every file of every tree must have arrived intact: proof that no concurrent job was
        // starved to a zero-progress stall and that the dynamic scheduler dropped nothing.
        for (unsigned int t = 0; t < kTreeCount; ++t)
        {
            for (unsigned int f = 0; f < kFilesPerTree; ++f)
            {
                const std::filesystem::path copied = dstRoot / treeName(t) / fileName(f);
                if (! FileSizeEquals(copied, kFileBytes))
                {
                    Fail(std::format(L"Saturation test: {} did not copy intact (expected {} bytes).", copied.wstring(), kFileBytes));
                    return true;
                }
            }
        }

        if (state.fairstreamSaturationMaxDistinctInFlightTrees < 2u)
        {
            Fail(std::format(
                L"Saturation test never observed at least two distinct source trees in flight in the same sample window (max distinct source trees={}).",
                state.fairstreamSaturationMaxDistinctInFlightTrees));
            return true;
        }

        std::vector<PerfMetricSample> dispatchMetrics;
        if (! TryReadPerfMetricSamples("FileOps.CopyRecursiveParallel.WorkItemDispatches", dispatchMetrics) ||
            dispatchMetrics.size() < state.fairstreamSaturationDispatchMetricBaseline + kTreeCount)
        {
            Fail(std::format(L"Saturation test expected at least {} new FileOps.CopyRecursiveParallel.WorkItemDispatches samples, baseline={}, total={}.",
                             kTreeCount,
                             state.fairstreamSaturationDispatchMetricBaseline,
                             dispatchMetrics.size()));
            return true;
        }

        size_t robustDispatchSamples = 0;
        for (size_t i = state.fairstreamSaturationDispatchMetricBaseline; i < dispatchMetrics.size(); ++i)
        {
            const PerfMetricSample& sample = dispatchMetrics[i];
            if (sample.hr == S_OK && sample.value0 >= kFilesPerTree && sample.value1 == kExpectedConcurrency && sample.durationUs >= kFilesPerTree &&
                sample.durationUs >= static_cast<uint64_t>(kExpectedConcurrency * 3u))
            {
                ++robustDispatchSamples;
            }
        }

        if (robustDispatchSamples < kTreeCount)
        {
            Fail(
                std::format(L"Saturation test expected at least {} robust dispatch metric samples after baseline, saw {}.", kTreeCount, robustDispatchSamples));
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamSaturationConcurrentCopies",
                          L"shape=4-trees-dispatches-exceed-concurrency",
                          kTreeCount,
                          state.fairstreamSaturationMaxDistinctInFlightTrees,
                          robustDispatchSamples,
                          S_OK);
        NextStep(state, SelfTestState::Step::Fairstream_EarlyAdmissionOverlapsPreCalc);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_EarlyAdmissionOverlapsPreCalc:
{
    // Fairstream 5F. Pre-calculation now runs CONCURRENTLY with the transfer (early admission):
    // bytes start moving before the recursive size scan finishes. The deterministic proof is the
    // engine latching _transferStartedBeforePreCalcComplete when a transfer progress callback fires
    // while pre-calc is still in progress — impossible in the old serial pre-calc-then-execute model
    // (RED there: pre-calc always finished before the first transfer callback). Plus: the operation
    // must still complete correctly with the totals reconciled once pre-calc publishes them.
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"Fairstream_EarlyAdmissionOverlapsPreCalc timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-earlyadmit-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-earlyadmit-dst";
    // Wide tree (many entries) so single-worker pre-calc enumeration runs long enough to overlap the
    // transfer; modest per-file bytes keep the run quick.
    constexpr unsigned int kDirs = 16u;
    constexpr unsigned int kFilesPerDir = 16u;
    constexpr size_t kFileBytes = 8u * 1024u;
    constexpr unsigned int kDirectorySizeDelayMs = 5u;

    const auto dirName = [](unsigned int i) noexcept { return std::format(L"d{}", i); };
    const auto fileName = [](unsigned int i) noexcept { return std::format(L"f{}.bin", i); };

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Early-admission test failed to reset directories.");
            return true;
        }
        for (unsigned int d = 0; d < kDirs; ++d)
        {
            const std::filesystem::path dir = srcRoot / dirName(d);
            if (! RecreateEmptyDirectory(dir))
            {
                Fail(L"Early-admission test failed to seed a source directory.");
                return true;
            }
            for (unsigned int f = 0; f < kFilesPerDir; ++f)
            {
                if (! WriteTestFile(dir / fileName(f), kFileBytes))
                {
                    Fail(L"Early-admission test failed to seed source files.");
                    return true;
                }
            }
        }

        const std::string localConfig = std::format(
            R"json({{"concurrencyMode":"auto","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"directorySizeDelayMs":{}}})json",
            kDirectorySizeDelayMs);
        if (! SetPluginConfiguration(state.infoLocal.get(), localConfig))
        {
            Fail(L"Early-admission test failed to apply delayed local pre-calc configuration.");
            return true;
        }

        // Enable pre-calc with a SINGLE worker and a debug-only local enumeration delay: the
        // transfer must start while the scan is still running, giving a deterministic overlap window.
        g_settings.fileOperations->preCalcEnabled = true;
        g_settings.fileOperations->preCalcMaxWorkers = 1u;

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcRoot},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            if (! state.localConfigOriginal.empty())
            {
                static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), state.localConfigOriginal));
            }
            Fail(L"Early-admission test failed to start the copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    if (! state.localConfigOriginal.empty() && ! SetPluginConfiguration(state.infoLocal.get(), state.localConfigOriginal))
    {
        Fail(L"Early-admission test failed to restore local plugin configuration.");
        return true;
    }

    if (FAILED(completed->second.hr))
    {
        Fail(std::format(L"Early-admission copy failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    // The transfer must have overlapped pre-calc (this is the whole point of 5F).
    if (! completed->second.earlyAdmissionTransferObserved)
    {
        Fail(std::format(L"Early admission did not engage: no transfer progress fired while pre-calc was still running "
                         L"(preCalcUs={} completedBytes={} expectedBytes={} configuredConcurrency={} delayMs={}).",
                         completed->second.preCalcDurationUs,
                         completed->second.progressCompletedBytes,
                         static_cast<uint64_t>(kDirs) * kFilesPerDir * kFileBytes,
                         completed->second.configuredMaxConcurrency,
                         kDirectorySizeDelayMs));
        return true;
    }

    // Pre-calc must still have completed, so the estimated totals reconciled to real ones.
    if (! completed->second.preCalcCompleted)
    {
        Fail(L"Early-admission run finished but pre-calc never completed (totals would not reconcile).");
        return true;
    }

    // Correctness: every file of every directory copied intact (no item dropped during the overlap).
    // Copying the source directory into dstRoot nests it under its own leaf name.
    const std::wstring srcLeaf = srcRoot.filename().wstring();
    const uint64_t expectedBytes = static_cast<uint64_t>(kDirs) * kFilesPerDir * kFileBytes;
    for (unsigned int d = 0; d < kDirs; ++d)
    {
        for (unsigned int f = 0; f < kFilesPerDir; ++f)
        {
            const std::filesystem::path copied = dstRoot / srcLeaf / dirName(d) / fileName(f);
            if (! FileSizeEquals(copied, kFileBytes))
            {
                Fail(std::format(L"Early-admission test: {} did not copy intact.", copied.wstring()));
                return true;
            }
        }
    }

    // Totals reconciled: all bytes accounted for at completion (completed == the real total).
    if (completed->second.progressCompletedBytes < expectedBytes)
    {
        Fail(std::format(
            L"Early-admission totals did not reconcile: completedBytes={} < expected={}.", completed->second.progressCompletedBytes, expectedBytes));
        return true;
    }

    Debug::Perf::Emit(
        L"FileOps.SelfTest.FairstreamEarlyAdmission", L"", completed->second.progressCompletedBytes, completed->second.preCalcTotalBytes, expectedBytes, S_OK);
    NextStep(state, SelfTestState::Step::Riptide_EarlyAdmissionThreadStartFailureFallsBack);
    return false;
}
case SelfTestState::Step::Riptide_EarlyAdmissionThreadStartFailureFallsBack:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 90'000ull))
    {
        Fail(L"Riptide_EarlyAdmissionThreadStartFailureFallsBack timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"riptide-earlyadmit-thread-fail-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"riptide-earlyadmit-thread-fail-dst";
    constexpr unsigned int kDirs = 4u;
    constexpr unsigned int kFilesPerDir = 4u;
    constexpr size_t kFileBytes = 4u * 1024u;
    const auto dirName = [](unsigned int index) noexcept { return std::format(L"d{}", index); };
    const auto fileName = [](unsigned int index) noexcept { return std::format(L"f{}.bin", index); };

    if (state.stepState == 0)
    {
        SetFileOpsPreCalcThreadStartFailureForSelfTest(false);
        static_cast<void>(TakeFileOpsPreCalcThreadStartAttemptsForSelfTest());

        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Riptide early-admission thread-failure test failed to reset directories.");
            return true;
        }

        for (unsigned int d = 0; d < kDirs; ++d)
        {
            const std::filesystem::path dir = srcRoot / dirName(d);
            if (! RecreateEmptyDirectory(dir))
            {
                Fail(L"Riptide early-admission thread-failure test failed to seed a source directory.");
                return true;
            }
            for (unsigned int f = 0; f < kFilesPerDir; ++f)
            {
                if (! WriteTestFile(dir / fileName(f), kFileBytes))
                {
                    Fail(L"Riptide early-admission thread-failure test failed to seed source files.");
                    return true;
                }
            }
        }

        g_settings.fileOperations->preCalcEnabled = true;
        g_settings.fileOperations->preCalcMaxWorkers = 1u;
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"auto"})json"));

        SetFileOpsPreCalcThreadStartFailureForSelfTest(true);
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcRoot},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            SetFileOpsPreCalcThreadStartFailureForSelfTest(false);
            Fail(L"Riptide early-admission thread-failure test failed to start the copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    SetFileOpsPreCalcThreadStartFailureForSelfTest(false);
    const unsigned long attempts = TakeFileOpsPreCalcThreadStartAttemptsForSelfTest();
    if (attempts != 1u)
    {
        Fail(std::format(L"Riptide early-admission thread-failure test expected one pre-calc thread-start attempt, got {}.", attempts));
        return true;
    }

    if (FAILED(completed->second.hr))
    {
        Fail(std::format(L"Riptide early-admission thread-failure fallback copy failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    if (! completed->second.preCalcCompleted || completed->second.preCalcWorkerCountUsed == 0u)
    {
        Fail(std::format(L"Riptide early-admission fallback did not run serial pre-calc (completed={} workers={}).",
                         completed->second.preCalcCompleted ? 1 : 0,
                         completed->second.preCalcWorkerCountUsed));
        return true;
    }

    if (completed->second.earlyAdmissionTransferObserved)
    {
        Fail(L"Riptide early-admission fallback still transferred while pre-calc was in progress.");
        return true;
    }

    const std::wstring srcLeaf = srcRoot.filename().wstring();
    for (unsigned int d = 0; d < kDirs; ++d)
    {
        for (unsigned int f = 0; f < kFilesPerDir; ++f)
        {
            const std::filesystem::path copied = dstRoot / srcLeaf / dirName(d) / fileName(f);
            if (! FileSizeEquals(copied, kFileBytes))
            {
                Fail(std::format(L"Riptide early-admission thread-failure fallback: {} did not copy intact.", copied.wstring()));
                return true;
            }
        }
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.RiptideEarlyAdmissionThreadStartFailureFallback",
                      L"",
                      completed->second.progressCompletedBytes,
                      completed->second.preCalcTotalBytes,
                      attempts,
                      S_OK);
    NextStep(state, SelfTestState::Step::Riptide_EarlyAdmissionDisabledDoesNotStartThread);
    return false;
}
case SelfTestState::Step::Riptide_EarlyAdmissionDisabledDoesNotStartThread:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        Fail(L"Riptide_EarlyAdmissionDisabledDoesNotStartThread timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"riptide-earlyadmit-disabled-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"riptide-earlyadmit-disabled-dst";
    constexpr size_t kFileBytes = 6u * 1024u;

    if (state.stepState == 0)
    {
        SetFileOpsPreCalcThreadStartFailureForSelfTest(false);
        static_cast<void>(TakeFileOpsPreCalcThreadStartAttemptsForSelfTest());

        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Riptide early-admission disabled test failed to reset directories.");
            return true;
        }
        if (! WriteTestFile(srcRoot / L"payload.bin", kFileBytes))
        {
            Fail(L"Riptide early-admission disabled test failed to seed the source file.");
            return true;
        }

        g_settings.fileOperations->preCalcEnabled = false;
        g_settings.fileOperations->preCalcMaxWorkers = 1u;
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"auto"})json"));

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcRoot},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Riptide early-admission disabled test failed to start the copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    const unsigned long attempts = TakeFileOpsPreCalcThreadStartAttemptsForSelfTest();
    if (attempts != 0u)
    {
        Fail(std::format(L"Disabled pre-calc COPY still attempted an early-admission pre-calc thread (attempts={}).", attempts));
        return true;
    }

    if (FAILED(completed->second.hr))
    {
        Fail(std::format(L"Riptide early-admission disabled copy failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    if (completed->second.preCalcCompleted || completed->second.preCalcWorkerCountUsed != 0u || completed->second.preCalcTotalBytes != 0u)
    {
        Fail(std::format(L"Disabled pre-calc COPY still ran pre-calc (completed={} workers={} bytes={}).",
                         completed->second.preCalcCompleted ? 1 : 0,
                         completed->second.preCalcWorkerCountUsed,
                         completed->second.preCalcTotalBytes));
        return true;
    }

    const std::filesystem::path copied = dstRoot / srcRoot.filename() / L"payload.bin";
    if (! FileSizeEquals(copied, kFileBytes))
    {
        Fail(std::format(L"Riptide early-admission disabled copy: {} did not copy intact.", copied.wstring()));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.RiptideEarlyAdmissionDisabledNoThreadStart",
                      L"",
                      completed->second.progressCompletedBytes,
                      completed->second.preCalcTotalBytes,
                      attempts,
                      S_OK);
    NextStep(state, SelfTestState::Step::Riptide_LiveFinishedSnapshotCarriesDiagnostics);
    return false;
}
case SelfTestState::Step::Riptide_LiveFinishedSnapshotCarriesDiagnostics:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
        SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
        Fail(L"Riptide_LiveFinishedSnapshotCarriesDiagnostics timed out.");
        return true;
    }

    if (! state.fileOps)
    {
        Fail(L"Riptide_LiveFinishedSnapshotCarriesDiagnostics missing file operation state.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"riptide-live-finished-diag-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"riptide-live-finished-diag-dst";
    const std::filesystem::path srcFile = srcRoot / L"payload.bin";
    const std::filesystem::path dstFile = dstRoot / L"payload.bin";
    constexpr size_t kFileBytes = 8u * 1024u;

    if (state.stepState == 0)
    {
        ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
        SetFileOpsPostFinishedCompletionPauseForSelfTest(false);

        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Riptide live-finished diagnostics test failed to reset directories.");
            return true;
        }
        if (! WriteTestFile(srcFile, kFileBytes))
        {
            Fail(L"Riptide live-finished diagnostics test failed to seed the source file.");
            return true;
        }

        state.fileOps->SetAutoDismissSuccess(false);
        SetFileOpsPostFinishedCompletionPauseForSelfTest(true);

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        state.taskA = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {srcFile}, dstRoot, flags, false);
        if (! state.taskA.has_value())
        {
            ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
            SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
            Fail(L"Riptide live-finished diagnostics test failed to start the copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        state.fileOps->RecordTaskDiagnostic(state.taskA.value(),
                                            FILESYSTEM_COPY,
                                            FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                            S_OK,
                                            L"selftest.live-finished",
                                            L"Selftest warning carried by live finished snapshot.",
                                            srcFile.wstring(),
                                            dstFile.wstring());
        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (! HasFileOpsPostFinishedCompletionPauseEnteredForSelfTest())
        {
            return false;
        }

        FileOperationsPopupInternal::TaskSnapshot snapshot{};
        if (! TryGetPopupTaskSnapshot(state.fileOps, state.taskA.value(), snapshot))
        {
            if (state.markerTick == 0)
            {
                state.markerTick = nowTick;
            }
            if (nowTick >= state.markerTick && (nowTick - state.markerTick) > SelfTest::ScaleTimeout(2'000ull))
            {
                ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
                SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
                Fail(L"Riptide live-finished diagnostics test could not read the paused popup snapshot.");
                return true;
            }
            return false;
        }

        if (! snapshot.finished || FAILED(snapshot.resultHr))
        {
            ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
            SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
            Fail(std::format(L"Riptide live-finished diagnostics test expected a paused successful finished row (finished={} hr=0x{:08X}).",
                             snapshot.finished ? 1 : 0,
                             static_cast<unsigned long>(snapshot.resultHr)));
            return true;
        }

        if (snapshot.warningCount != 1u || snapshot.errorCount != 0u)
        {
            ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
            SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
            Fail(std::format(
                L"Riptide live-finished diagnostics test expected live warning/error counts 1/0, got {}/{}.", snapshot.warningCount, snapshot.errorCount));
            return true;
        }

        if (snapshot.statusKind != FileOperationsPopupInternal::TaskSnapshot::StatusKind::Partial)
        {
            ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
            SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
            Fail(L"Riptide live-finished diagnostics test expected succeeded-with-warning live row to resolve Partial, not Done.");
            return true;
        }

        if (snapshot.lastDiagnosticMessage.empty())
        {
            ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
            SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
            Fail(L"Riptide live-finished diagnostics test expected the live row to carry the last diagnostic message.");
            return true;
        }

        ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
        SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
        state.stepState = 3;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    if (FAILED(completed->second.hr))
    {
        Fail(std::format(L"Riptide live-finished diagnostics copy failed after release: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    if (! FileSizeEquals(dstFile, kFileBytes))
    {
        Fail(std::format(L"Riptide live-finished diagnostics copy did not produce the expected destination file: {}.", dstFile.wstring()));
        return true;
    }

    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
    state.fileOps->CollectCompletedTasks(summaries);
    const auto summaryIt = std::find_if(
        summaries.begin(), summaries.end(), [&](const auto& summary) noexcept { return state.taskA.has_value() && summary.taskId == state.taskA.value(); });
    if (summaryIt == summaries.end())
    {
        Fail(L"Riptide live-finished diagnostics test could not find the completed summary.");
        return true;
    }
    if (summaryIt->warningCount != 1u || summaryIt->errorCount != 0u)
    {
        Fail(std::format(L"Riptide live-finished diagnostics completed summary expected warning/error counts 1/0, got {}/{}.",
                         summaryIt->warningCount,
                         summaryIt->errorCount));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.RiptideLiveFinishedSnapshotDiagnostics",
                      L"",
                      summaryIt->warningCount,
                      summaryIt->errorCount,
                      completed->second.progressCompletedBytes,
                      S_OK);
    NextStep(state, SelfTestState::Step::Riptide_BridgeSequentialContinueOnErrorCopiesSiblings);
    return false;
}
case SelfTestState::Step::Riptide_BridgeSequentialContinueOnErrorCopiesSiblings:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest());
        Fail(L"Riptide_BridgeSequentialContinueOnErrorCopiesSiblings timed out.");
        return true;
    }

    const std::wstring dummyRoot = L"/riptide-bridge-seq-coe-src";
    const std::filesystem::path localDest = state.tempRoot / L"riptide-bridge-seq-coe-dst";
    const std::filesystem::path destRoot = localDest / L"riptide-bridge-seq-coe-src";
    constexpr std::array<std::wstring_view, 4> kFileNames{{
        L"alpha.txt",
        L"bravo.txt",
        L"charlie.txt",
        L"delta.txt",
    }};
    constexpr std::string_view kCopiedBytes = "copy-survived";

    if (state.stepState == 0)
    {
        SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest());
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! RecreateEmptyDirectory(localDest))
        {
            Fail(L"Riptide sequential bridge continue-on-error test failed to reset source/destination roots.");
            return true;
        }

        for (const std::wstring_view fileName : kFileNames)
        {
            const std::wstring dummyFile = dummyRoot + L"/" + std::wstring(fileName);
            if (! DummyWriteTextFile(state.fsDummy.get(), dummyFile, kCopiedBytes))
            {
                Fail(std::format(L"Riptide sequential bridge continue-on-error test failed to seed dummy file '{}'.", dummyFile));
                return true;
            }
        }

        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));
        SetFileOpsBridgeFailNextFileCopiesForSelfTest(1);

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyRoot)},
                                                 localDest,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
            static_cast<void>(TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest());
            Fail(L"Riptide sequential bridge continue-on-error test failed to start the cross-FS copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    const unsigned long attempts = TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest();
    SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);

    if (attempts != 1u)
    {
        Fail(std::format(L"Riptide sequential bridge continue-on-error test expected exactly one injected file-copy failure, got {}.", attempts));
        return true;
    }

    if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        Fail(std::format(L"Riptide sequential bridge continue-on-error test expected ERROR_PARTIAL_COPY, got 0x{:08X}.",
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    unsigned int copiedCount = 0;
    for (const std::wstring_view fileName : kFileNames)
    {
        if (FileSizeEquals(destRoot / std::wstring(fileName), kCopiedBytes.size()))
        {
            ++copiedCount;
        }
    }

    if (copiedCount != static_cast<unsigned int>(kFileNames.size() - 1u))
    {
        Fail(std::format(L"Riptide sequential bridge continue-on-error test expected {} copied siblings after one failed child, got {}.",
                         static_cast<unsigned int>(kFileNames.size() - 1u),
                         copiedCount));
        return true;
    }

    Debug::Perf::Emit(
        L"FileOps.SelfTest.RiptideBridgeSequentialContinueOnError", L"", copiedCount, attempts, completed->second.progressCompletedBytes, completed->second.hr);
    NextStep(state, SelfTestState::Step::Causeway_BridgeRejectsHostileChildNames);
    return false;
}
case SelfTestState::Step::Causeway_BridgeRejectsHostileChildNames:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        SetFileOpsBridgeInjectHostileChildNamesForSelfTest(false);
        static_cast<void>(TakeFileOpsBridgeInjectHostileChildNameAttemptsForSelfTest());
        Fail(L"Causeway_BridgeRejectsHostileChildNames timed out.");
        return true;
    }

    constexpr std::array<std::wstring_view, 8> kPlaceholderNames{{
        L"placeholder-00-long-safe-name.txt",
        L"placeholder-01-long-safe-name.txt",
        L"placeholder-02-long-safe-name.txt",
        L"placeholder-03-long-safe-name.txt",
        L"placeholder-04-long-safe-name.txt",
        L"placeholder-05-long-safe-name.txt",
        L"placeholder-06-long-safe-name.txt",
        L"placeholder-07-long-safe-name.txt",
    }};
    constexpr std::wstring_view kValidSibling = L"z-valid-sibling.txt";
    constexpr std::string_view kPayload        = "causeway-hostile-child-name-payload";
    const unsigned int scenario                = static_cast<unsigned int>(state.stepState / 2);
    const FileSystemOperation operation        = scenario == 0u ? FILESYSTEM_COPY : FILESYSTEM_MOVE;
    const std::wstring rootLeaf                = std::format(L"causeway-hostile-{}-src", scenario == 0u ? L"copy" : L"move");
    const std::wstring dummyRoot               = L"/" + rootLeaf;
    const std::wstring dummyEscapeFile         = L"/escape.txt";
    const std::filesystem::path localDest      = state.tempRoot / std::format(L"causeway-hostile-{}-dst", scenario == 0u ? L"copy" : L"move");
    const std::filesystem::path destinationRoot = localDest / rootLeaf;
    const std::filesystem::path escapedDestination = localDest / L"escape.txt";

    if (scenario >= 2u)
    {
        NextStep(state, SelfTestState::Step::Causeway_BridgeProviderOutputContracts);
        return false;
    }

    if ((state.stepState % 2) == 0)
    {
        SetFileOpsBridgeInjectHostileChildNamesForSelfTest(false);
        static_cast<void>(TakeFileOpsBridgeInjectHostileChildNameAttemptsForSelfTest());
        g_settings.fileOperations->preCalcEnabled = false;
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));

        const FileSystemFlags dummyCleanupFlags =
            static_cast<FileSystemFlags>(static_cast<uint32_t>(FILESYSTEM_FLAG_RECURSIVE) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY));
        static_cast<void>(state.fsDummy->DeleteItem(dummyRoot.c_str(), dummyCleanupFlags, nullptr, nullptr, nullptr));
        static_cast<void>(state.fsDummy->DeleteItem(dummyEscapeFile.c_str(), dummyCleanupFlags, nullptr, nullptr, nullptr));

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! DummyWriteTextFile(state.fsDummy.get(), dummyEscapeFile, kPayload, true) ||
            ! RecreateEmptyDirectory(localDest))
        {
            Fail(L"Causeway hostile-child-name test could not reset its source/destination.");
            return true;
        }
        for (const std::wstring_view placeholder : kPlaceholderNames)
        {
            if (! DummyWriteTextFile(state.fsDummy.get(), dummyRoot + L"/" + std::wstring(placeholder), kPayload, true))
            {
                Fail(L"Causeway hostile-child-name test could not seed a provider placeholder entry.");
                return true;
            }
        }
        if (! DummyWriteTextFile(state.fsDummy.get(), dummyRoot + L"/" + std::wstring(kValidSibling), kPayload, true))
        {
            Fail(L"Causeway hostile-child-name test could not seed its valid sibling.");
            return true;
        }

        SetFileOpsBridgeInjectHostileChildNamesForSelfTest(true);
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 operation,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyRoot)},
                                                 localDest,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            SetFileOpsBridgeInjectHostileChildNamesForSelfTest(false);
            Fail(L"Causeway hostile-child-name bridge operation did not start.");
            return true;
        }

        ++state.stepState;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    const unsigned long injectionAttempts = TakeFileOpsBridgeInjectHostileChildNameAttemptsForSelfTest();
    SetFileOpsBridgeInjectHostileChildNamesForSelfTest(false);
    if (injectionAttempts != 1u || completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        Fail(std::format(L"Causeway hostile-child-name {} expected one injection and ERROR_PARTIAL_COPY; got attempts={} hr=0x{:08X}.",
                         scenario == 0u ? L"COPY" : L"MOVE",
                         injectionAttempts,
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }
    if (! FileSizeEquals(destinationRoot / std::wstring(kValidSibling), kPayload.size()))
    {
        Fail(L"Causeway hostile-child-name test did not copy the valid sibling after rejecting malicious entries.");
        return true;
    }

    std::error_code ec;
    if (std::filesystem::exists(escapedDestination, ec) || ec)
    {
        Fail(L"Causeway hostile-child-name test observed an escaped destination write.");
        return true;
    }
    for (std::filesystem::recursive_directory_iterator it(localDest, ec), end; ! ec && it != end; it.increment(ec))
    {
        if (it->path().filename().native().find(L".rs_tmp_") != std::wstring::npos)
        {
            Fail(L"Causeway hostile-child-name test left a staging file behind.");
            return true;
        }
    }
    if (ec)
    {
        Fail(L"Causeway hostile-child-name test could not inspect the destination tree.");
        return true;
    }

    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
    state.fileOps->CollectCompletedTasks(summaries);
    const auto summary = std::find_if(summaries.begin(), summaries.end(), [&](const auto& value) noexcept {
        return state.taskA.has_value() && value.taskId == state.taskA.value();
    });
    if (summary == summaries.end())
    {
        Fail(L"Causeway hostile-child-name test could not find its completed summary.");
        return true;
    }
    const size_t rejectedNameCount = static_cast<size_t>(std::count_if(summary->issueDiagnostics.begin(), summary->issueDiagnostics.end(), [](const auto& issue) noexcept {
        return issue.category == L"bridge.source.invalidChildName";
    }));
    constexpr size_t kExpectedRejectedEntries = 7u; // Six invalid components plus the second name in the case-colliding pair.
    if (rejectedNameCount != kExpectedRejectedEntries)
    {
        Fail(std::format(L"Causeway hostile-child-name test expected {} rejected entries, got {}.", kExpectedRejectedEntries, rejectedNameCount));
        return true;
    }
    if (std::filesystem::exists(destinationRoot / L"Case.txt", ec) || std::filesystem::exists(destinationRoot / L"case.txt", ec) || ec)
    {
        Fail(L"Causeway hostile-child-name test unexpectedly wrote a member of the case-colliding pair.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.CausewayBridgeHostileChildNames",
                      scenario == 0u ? L"COPY" : L"MOVE",
                      rejectedNameCount,
                      injectionAttempts,
                      completed->second.progressCompletedBytes,
                      completed->second.hr);
    ++state.stepState;
    state.taskA.reset();
    return false;
}
case SelfTestState::Step::Causeway_BridgeProviderOutputContracts:
{
    constexpr std::array<std::wstring_view, 4> kScenarioNames{{L"read-overreport", L"premature-eof", L"write-underconsume", L"write-overreport"}};
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        SetFileOpsBridgeOverReportNextReadForSelfTest(0);
        SetFileOpsBridgePrematureEofNextReadForSelfTest(0);
        SetFileOpsBridgeUnderConsumeNextWriteForSelfTest(0);
        SetFileOpsBridgeOverReportNextWriteForSelfTest(0);
        Fail(L"Causeway_BridgeProviderOutputContracts timed out.");
        return true;
    }

    const unsigned int scenario = static_cast<unsigned int>(state.stepState / 2);
    if (scenario >= kScenarioNames.size())
    {
        NextStep(state, SelfTestState::Step::Causeway_BridgeSchedulingAndResourceContracts);
        return false;
    }

    const std::wstring dummyRoot = L"/causeway-provider-contract-src";
    const std::wstring dummyFile = dummyRoot + L"/payload.bin";
    const std::filesystem::path localDest = state.tempRoot / std::format(L"causeway-provider-contract-{}", kScenarioNames[scenario]);
    constexpr std::string_view kSourceContents = "causeway-provider-contract-payload";

    const auto resetHooks = []() noexcept
    {
        SetFileOpsBridgeOverReportNextReadForSelfTest(0);
        SetFileOpsBridgePrematureEofNextReadForSelfTest(0);
        SetFileOpsBridgeUnderConsumeNextWriteForSelfTest(0);
        SetFileOpsBridgeOverReportNextWriteForSelfTest(0);
    };
    const auto takeAttempts = [scenario]() noexcept -> unsigned long
    {
        switch (scenario)
        {
            case 0: return TakeFileOpsBridgeOverReportNextReadAttemptsForSelfTest();
            case 1: return TakeFileOpsBridgePrematureEofNextReadAttemptsForSelfTest();
            case 2: return TakeFileOpsBridgeUnderConsumeNextWriteAttemptsForSelfTest();
            case 3: return TakeFileOpsBridgeOverReportNextWriteAttemptsForSelfTest();
            default: return 0u;
        }
    };

    if ((state.stepState % 2) == 0)
    {
        resetHooks();
        static_cast<void>(takeAttempts());
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, kSourceContents, true) ||
            ! RecreateEmptyDirectory(localDest))
        {
            Fail(std::format(L"Causeway provider-contract scenario '{}' could not seed its source/destination.", kScenarioNames[scenario]));
            return true;
        }

        switch (scenario)
        {
            case 0: SetFileOpsBridgeOverReportNextReadForSelfTest(1); break;
            case 1: SetFileOpsBridgePrematureEofNextReadForSelfTest(1); break;
            case 2: SetFileOpsBridgeUnderConsumeNextWriteForSelfTest(1); break;
            case 3: SetFileOpsBridgeOverReportNextWriteForSelfTest(1); break;
            default: break;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyFile)},
                                                 localDest,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            resetHooks();
            static_cast<void>(takeAttempts());
            Fail(std::format(L"Causeway provider-contract scenario '{}' did not start.", kScenarioNames[scenario]));
            return true;
        }

        ++state.stepState;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    const unsigned long attempts = takeAttempts();
    resetHooks();
    if (attempts != 1u)
    {
        Fail(std::format(L"Causeway provider-contract scenario '{}' expected one injected provider violation, got {}.", kScenarioNames[scenario], attempts));
        return true;
    }
    if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        Fail(std::format(L"Causeway provider-contract scenario '{}' expected ERROR_PARTIAL_COPY, got 0x{:08X}.",
                         kScenarioNames[scenario],
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    std::error_code ec;
    if (! std::filesystem::is_empty(localDest, ec) || ec)
    {
        Fail(std::format(L"Causeway provider-contract scenario '{}' left a final or staging file behind.", kScenarioNames[scenario]));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.CausewayBridgeProviderOutputContract", std::wstring(kScenarioNames[scenario]), attempts, 0u, 0u, completed->second.hr);
    ++state.stepState;
    state.taskA.reset();
    return false;
}
case SelfTestState::Step::Causeway_BridgeSchedulingAndResourceContracts:
{
    constexpr uint64_t kBridgeBudgetCeilingBytes = 256ull * 1024ull * 1024ull;
    constexpr unsigned int kFilesPerDirectory    = 10u;
    constexpr unsigned int kDirectoryCount       = 2u;
    const ULONGLONG nowTick                       = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Causeway_BridgeSchedulingAndResourceContracts timed out.");
        return true;
    }

    if (state.stepState == 0)
    {
        if (! RunFileOpsPerItemSchedulerNestedSaturationSelfTestForSelfTest(*state.fileOps))
        {
            Fail(L"Causeway nested scheduler saturation probe failed.");
            return true;
        }
        constexpr std::array<DWORD, 6> kTransientErrors{{
            ERROR_TIMEOUT,
            ERROR_SEM_TIMEOUT,
            ERROR_CONNECTION_ABORTED,
            ERROR_CONNECTION_REFUSED,
            ERROR_BAD_NET_RESP,
            ERROR_UNEXP_NET_ERR,
        }};
        for (const DWORD error : kTransientErrors)
        {
            if (! IsFileOpsCircuitBreakerTransientErrorForSelfTest(error))
            {
                Fail(std::format(L"Causeway circuit-breaker classifier rejected transient error {}.", error));
                return true;
            }
        }
        state.stepState = 1;
    }

    const unsigned int scenario = static_cast<unsigned int>((state.stepState - 1u) / 2u);
    if (scenario >= 2u)
    {
        NextStep(state, SelfTestState::Step::Causeway_BridgeFileReparsePolicy);
        return false;
    }

    const FileSystemOperation operation = scenario == 0u ? FILESYSTEM_COPY : FILESYSTEM_MOVE;
    const std::wstring rootLeaf         = std::format(L"causeway-resource-{}-src", scenario == 0u ? L"copy" : L"move");
    const std::wstring dummyRoot        = L"/" + rootLeaf;
    const std::wstring dummyNested      = dummyRoot + L"/nested";
    const std::filesystem::path localDest = state.tempRoot / std::format(L"causeway-resource-{}-dst", scenario == 0u ? L"copy" : L"move");

    if (((state.stepState - 1u) % 2u) == 0u)
    {
        if (! g_settings.fileOperations.has_value())
        {
            g_settings.fileOperations.emplace();
        }
        g_settings.fileOperations->preCalcEnabled                = true;
        g_settings.fileOperations->crossFsBridgeBufferSizeKB     = 2048u;
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":16,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"copyReparse","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":20,"streamChunkLatencyMs":50,"virtualSpeedLimit":"0"})json"));

        const FileSystemFlags cleanupFlags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        static_cast<void>(state.fsDummy->DeleteItem(dummyRoot.c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyNested) ||
            ! RecreateEmptyDirectory(localDest))
        {
            Fail(L"Causeway scheduling/resource test could not reset its source and destination roots.");
            return true;
        }

        constexpr std::string_view kPayload = "causeway-bounded-bridge-buffer-payload";
        for (unsigned int i = 0u; i < kFilesPerDirectory; ++i)
        {
            if (! DummyWriteTextFile(state.fsDummy.get(), dummyRoot + std::format(L"/root-{:02}.bin", i), kPayload, true) ||
                ! DummyWriteTextFile(state.fsDummy.get(), dummyNested + std::format(L"/nested-{:02}.bin", i), kPayload, true))
            {
                Fail(L"Causeway scheduling/resource test could not seed its source files.");
                return true;
            }
        }

        ResetFileOpsBridgeBufferBudgetPeakForSelfTest();
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 operation,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyRoot)},
                                                 localDest,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            Fail(L"Causeway scheduling/resource bridge operation did not start.");
            return true;
        }

        ++state.stepState;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    if (FAILED(completed->second.hr))
    {
        Fail(std::format(L"Causeway scheduling/resource {} failed with 0x{:08X}.",
                         scenario == 0u ? L"COPY" : L"MOVE",
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }
    if (! completed->second.preCalcSuppressedForHighMetadataCrossFs || completed->second.preCalcCompleted)
    {
        Fail(L"Causeway high-metadata cross-filesystem operation did not suppress the redundant pre-calculation scan.");
        return true;
    }

    const uint64_t expectedEnumerations = scenario == 0u ? kDirectoryCount : 2u * kDirectoryCount;
    if (completed->second.bridgeSourceDirectoryEnumerationCount != expectedEnumerations)
    {
        Fail(std::format(L"Causeway {} expected {} source directory listings, got {}.",
                         scenario == 0u ? L"COPY" : L"MOVE",
                         expectedEnumerations,
                         completed->second.bridgeSourceDirectoryEnumerationCount));
        return true;
    }

    const uint64_t peakBudgetBytes = GetFileOpsBridgeBufferBudgetPeakForSelfTest();
    if (peakBudgetBytes == 0u || peakBudgetBytes > kBridgeBudgetCeilingBytes)
    {
        Fail(std::format(L"Causeway bridge buffer budget peak {} bytes violated ceiling {} bytes.", peakBudgetBytes, kBridgeBudgetCeilingBytes));
        return true;
    }

    const std::wstring_view operationName = scenario == 0u ? L"COPY" : L"MOVE";
    Debug::Perf::Emit(L"FileOps.SelfTest.CausewayBridgeBufferPeakBytes",
                      operationName,
                      0u,
                      peakBudgetBytes,
                      kBridgeBudgetCeilingBytes,
                      completed->second.hr);
    Debug::Perf::Emit(L"FileOps.SelfTest.CausewaySourceDirectoryEnumerations",
                      operationName,
                      0u,
                      completed->second.bridgeSourceDirectoryEnumerationCount,
                      expectedEnumerations,
                      completed->second.hr);
    ++state.stepState;
    state.taskA.reset();
    return false;
}
case SelfTestState::Step::Causeway_BridgeFileReparsePolicy:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        SetFileOpsBridgeInjectFileReparseForSelfTest(0);
        SetFileOpsBridgeReparsePolicyOverrideForSelfTest(FileOpsBridgeReparsePolicyOverride::None);
        Fail(L"Causeway_BridgeFileReparsePolicy timed out.");
        return true;
    }

    constexpr std::wstring_view kReparseFile = L"reparse-file.bin";
    constexpr std::wstring_view kValidFile   = L"valid-sibling.bin";
    constexpr std::string_view kPayload      = "causeway-reparse-file-target-bytes";
    const unsigned int scenario              = state.stepState < 2u ? 0u : 1u;
    const std::wstring rootLeaf              = std::format(L"causeway-file-reparse-{}-src", scenario == 0u ? L"skip" : L"copy");
    const std::wstring dummyRoot             = L"/" + rootLeaf;
    const std::filesystem::path localDest    = state.tempRoot / std::format(L"causeway-file-reparse-{}-dst", scenario == 0u ? L"skip" : L"copy");
    const std::filesystem::path copiedRoot   = localDest / rootLeaf;

    const auto resetHooks = []() noexcept
    {
        SetFileOpsBridgeInjectFileReparseForSelfTest(0);
        SetFileOpsBridgeReparsePolicyOverrideForSelfTest(FileOpsBridgeReparsePolicyOverride::None);
    };

    if (state.stepState == 0u || state.stepState == 2u)
    {
        resetHooks();
        static_cast<void>(TakeFileOpsBridgeInjectFileReparseAttemptsForSelfTest());
        g_settings.fileOperations->preCalcEnabled = false;
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"copyReparse","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        const FileSystemFlags cleanupFlags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        static_cast<void>(state.fsDummy->DeleteItem(dummyRoot.c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) ||
            ! DummyWriteTextFile(state.fsDummy.get(), dummyRoot + L"/" + std::wstring(kReparseFile), kPayload, true) ||
            ! DummyWriteTextFile(state.fsDummy.get(), dummyRoot + L"/" + std::wstring(kValidFile), kPayload, true) || ! RecreateEmptyDirectory(localDest))
        {
            resetHooks();
            Fail(L"Causeway file-reparse test could not seed its source and destination.");
            return true;
        }

        SetFileOpsBridgeInjectFileReparseForSelfTest(1u);
        SetFileOpsBridgeReparsePolicyOverrideForSelfTest(scenario == 0u ? FileOpsBridgeReparsePolicyOverride::Skip
                                                                        : FileOpsBridgeReparsePolicyOverride::CopyReparse);
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyRoot)},
                                                 localDest,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            resetHooks();
            Fail(L"Causeway file-reparse bridge copy did not start.");
            return true;
        }
        ++state.stepState;
        return false;
    }

    using Task = FolderWindow::FileOperationState::Task;
    Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    if (scenario == 1u && state.stepState == 3u)
    {
        const auto prompt = TryGetConflictPromptCopy(task);
        if (prompt.has_value())
        {
            if (prompt->bucket != Task::ConflictBucket::UnsupportedReparse || ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
            {
                resetHooks();
                Fail(L"Causeway file CopyReparse failure did not use the UnsupportedReparse conflict bucket.");
                return true;
            }
            task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
            state.stepState = 4u;
            return false;
        }
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    const unsigned long injectionAttempts = TakeFileOpsBridgeInjectFileReparseAttemptsForSelfTest();
    resetHooks();
    if (injectionAttempts != 1u || completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        Fail(std::format(L"Causeway file-reparse {} expected one injection and ERROR_PARTIAL_COPY; attempts={} hr=0x{:08X}.",
                         scenario == 0u ? L"Skip" : L"CopyReparse",
                         injectionAttempts,
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    std::error_code ec;
    if (std::filesystem::exists(copiedRoot / std::wstring(kReparseFile), ec) || ec ||
        ! FileSizeEquals(copiedRoot / std::wstring(kValidFile), kPayload.size()))
    {
        Fail(L"Causeway file-reparse policy copied target bytes or failed to continue with the valid sibling.");
        return true;
    }

    if (scenario == 1u)
    {
        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        const auto summary = std::find_if(summaries.begin(), summaries.end(), [&](const auto& value) noexcept {
            return state.taskA.has_value() && value.taskId == state.taskA.value();
        });
        const bool sawUnsupported = summary != summaries.end() &&
                                    std::any_of(summary->issueDiagnostics.begin(), summary->issueDiagnostics.end(), [](const auto& issue) noexcept {
                                        return issue.category == L"bridge.reparse.unsupported" && issue.status == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                                    });
        if (! sawUnsupported)
        {
            Fail(L"Causeway file CopyReparse scenario did not retain the ERROR_NOT_SUPPORTED diagnostic.");
            return true;
        }
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.CausewayBridgeFileReparsePolicy",
                      scenario == 0u ? L"Skip" : L"CopyReparse",
                      injectionAttempts,
                      0u,
                      kPayload.size(),
                      completed->second.hr);
    state.taskA.reset();
    if (scenario == 0u)
    {
        state.stepState = 2u;
        return false;
    }
    NextStep(state, SelfTestState::Step::Causeway_BridgeFailureStatusAndPausedReader);
    return false;
}
case SelfTestState::Step::Causeway_BridgeFailureStatusAndPausedReader:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest());
        Fail(L"Causeway_BridgeFailureStatusAndPausedReader timed out.");
        return true;
    }

    constexpr std::wstring_view kDummyRoot = L"/causeway-worker-failure-src";
    const std::filesystem::path localDest  = state.tempRoot / L"causeway-worker-failure-dst";
    if (state.stepState == 0u)
    {
        if (! RunFileOpsBridgePausedReaderStopSelfTestForSelfTest(*state.fileOps))
        {
            Fail(L"Causeway paused pipeline reader did not honor the writer-side stop wakeup.");
            return true;
        }

        SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest());
        g_settings.fileOperations->preCalcEnabled = false;
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"copyReparse","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":10,"streamChunkLatencyMs":20,"virtualSpeedLimit":"0"})json"));
        const FileSystemFlags cleanupFlags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        static_cast<void>(state.fsDummy->DeleteItem(std::wstring(kDummyRoot).c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), kDummyRoot) || ! RecreateEmptyDirectory(localDest))
        {
            Fail(L"Causeway worker-failure propagation test could not reset its roots.");
            return true;
        }
        constexpr std::string_view kPayload = "causeway-worker-failure-payload";
        for (unsigned int i = 0u; i < 8u; ++i)
        {
            if (! DummyWriteTextFile(state.fsDummy.get(), std::wstring(kDummyRoot) + std::format(L"/payload-{:02}.bin", i), kPayload, true))
            {
                Fail(L"Causeway worker-failure propagation test could not seed its source files.");
                return true;
            }
        }

        SetFileOpsBridgeFailNextFileCopiesForSelfTest(1u);
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsDummy,
                                                 {std::filesystem::path(kDummyRoot)},
                                                 localDest,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
            Fail(L"Causeway worker-failure propagation bridge copy did not start.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }

    using Task = FolderWindow::FileOperationState::Task;
    Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    if (state.stepState == 1u)
    {
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
            if (completed != state.completedTasks.end())
            {
                SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
                Fail(std::format(L"Causeway worker failure completed without exposing its real status; hr=0x{:08X}.",
                                 static_cast<unsigned long>(completed->second.hr)));
                return true;
            }
            return false;
        }
        if (prompt->status != HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) || prompt->status == HRESULT_FROM_WIN32(ERROR_CANCELLED) ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
        {
            SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
            Fail(std::format(L"Causeway worker failure was not preserved (prompt status=0x{:08X}).", static_cast<unsigned long>(prompt->status)));
            return true;
        }
        task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
        state.stepState = 2u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    const unsigned long attempts = TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest();
    SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
    if (attempts != 1u || completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        Fail(std::format(L"Causeway worker-failure propagation expected one failure and final partial status; attempts={} hr=0x{:08X}.",
                         attempts,
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.CausewayFailureStatusAndPausedReader", L"", attempts, 0u, 0u, completed->second.hr);
    NextStep(state, SelfTestState::Step::Floodgate_CrossFsCopyGetSizeFailureRefusesCommit);
    return false;
}
case SelfTestState::Step::Floodgate_CrossFsCopyGetSizeFailureRefusesCommit:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest());
        Fail(L"Floodgate_CrossFsCopyGetSizeFailureRefusesCommit timed out.");
        return true;
    }

    const std::wstring dummyRoot = L"/floodgate-bridge-copy-getsize-src";
    const std::wstring dummyFile = dummyRoot + L"/payload.txt";
    const std::filesystem::path localDest = state.tempRoot / L"floodgate-bridge-copy-getsize-dst";
    const std::filesystem::path localDestFile = localDest / L"payload.txt";
    constexpr std::string_view kSourceContents = "source-must-not-commit-copy-without-size";

    if (state.stepState == 0)
    {
        SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest());
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, kSourceContents) ||
            ! RecreateEmptyDirectory(localDest))
        {
            Fail(L"Floodgate bridge COPY GetSize failure test failed to seed source/destination.");
            return true;
        }

        SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(1);
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyFile)},
                                                 localDest,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0);
            static_cast<void>(TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest());
            Fail(L"Floodgate bridge COPY GetSize failure test failed to start the cross-FS copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    const unsigned long attempts = TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest();
    SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0);

    if (attempts != 1u)
    {
        Fail(std::format(L"Floodgate bridge COPY GetSize failure test expected one injected source-size failure, got {}.", attempts));
        return true;
    }
    if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        Fail(std::format(L"Floodgate bridge COPY GetSize failure test expected ERROR_PARTIAL_COPY, got 0x{:08X}.",
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
    {
        Fail(L"Floodgate bridge COPY GetSize failure test could not query dummy IFileSystemIO.");
        return true;
    }

    std::string sourceText;
    if (! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyFile), sourceText) || sourceText != kSourceContents)
    {
        Fail(L"Floodgate bridge COPY GetSize failure test did not preserve the source bytes.");
        return true;
    }

    std::error_code ec;
    if (std::filesystem::exists(localDestFile, ec) || ec)
    {
        Fail(L"Floodgate bridge COPY GetSize failure test unexpectedly created the destination file after source-size failure.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateBridgeCopyGetSizeFailure", L"", attempts, 0, 0, completed->second.hr);
    NextStep(state, SelfTestState::Step::Floodgate_CrossFsMoveGetSizeFailurePreservesSource);
    return false;
}
case SelfTestState::Step::Floodgate_CrossFsMoveGetSizeFailurePreservesSource:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest());
        Fail(L"Floodgate_CrossFsMoveGetSizeFailurePreservesSource timed out.");
        return true;
    }

    const std::wstring dummyRoot = L"/floodgate-bridge-getsize-src";
    const std::wstring dummyFile = dummyRoot + L"/payload.txt";
    const std::filesystem::path localDest = state.tempRoot / L"floodgate-bridge-getsize-dst";
    const std::filesystem::path localDestFile = localDest / L"payload.txt";
    constexpr std::string_view kSourceContents = "source-must-survive-getsize-failure";

    if (state.stepState == 0)
    {
        SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest());
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, kSourceContents) ||
            ! RecreateEmptyDirectory(localDest))
        {
            Fail(L"Floodgate bridge GetSize failure test failed to seed source/destination.");
            return true;
        }

        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));
        SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(1);

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyFile)},
                                                 localDest,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0);
            static_cast<void>(TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest());
            Fail(L"Floodgate bridge GetSize failure test failed to start the cross-FS move.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    const unsigned long attempts = TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest();
    SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0);

    if (attempts != 1u)
    {
        Fail(std::format(L"Floodgate bridge GetSize failure test expected one injected source-size failure, got {}.", attempts));
        return true;
    }
    if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        Fail(
            std::format(L"Floodgate bridge GetSize failure test expected ERROR_PARTIAL_COPY, got 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
    {
        Fail(L"Floodgate bridge GetSize failure test could not query dummy IFileSystemIO.");
        return true;
    }

    std::string sourceText;
    if (! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyFile), sourceText) || sourceText != kSourceContents)
    {
        Fail(L"Floodgate bridge GetSize failure test did not preserve the source bytes.");
        return true;
    }

    std::error_code ec;
    if (std::filesystem::exists(localDestFile, ec) || ec)
    {
        Fail(L"Floodgate bridge GetSize failure test unexpectedly created the destination file after source-size failure.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateBridgeMoveGetSizeFailure", L"", attempts, 0, 0, completed->second.hr);
    NextStep(state, SelfTestState::Step::Floodgate_CrossFsMoveCleanupDetectsDestinationCorruption);
    return false;
}
case SelfTestState::Step::Floodgate_CrossFsMoveCleanupDetectsDestinationCorruption:
{
    constexpr const wchar_t* kMutateDestinationPathEnv = L"REDSALAMANDER_FILEOPS_BRIDGE_MUTATE_DESTINATION_BEFORE_MOVE_CLEANUP_PATH";
    constexpr const wchar_t* kMutateDestinationPayloadEnv = L"REDSALAMANDER_FILEOPS_BRIDGE_MUTATE_DESTINATION_BEFORE_MOVE_CLEANUP_PAYLOAD";
    const auto clearMutationEnv = []() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kMutateDestinationPathEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kMutateDestinationPayloadEnv, nullptr));
    };

    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        clearMutationEnv();
        static_cast<void>(TakeFileOpsBridgeMutateDestinationBeforeMoveCleanupAttemptsForSelfTest());
        Fail(L"Floodgate_CrossFsMoveCleanupDetectsDestinationCorruption timed out.");
        return true;
    }

    const std::wstring dummyRoot = L"/floodgate-crossfs-file-cleanup";
    const std::wstring dummyFile = dummyRoot + L"/payload.txt";
    const std::filesystem::path localDest = state.tempRoot / L"floodgate-crossfs-file-cleanup-dst";
    const std::filesystem::path localDestFile = localDest / L"payload.txt";
    constexpr std::string_view kSourceContents = "source-before-move";
    constexpr std::string_view kCorruptContents = "corrupt-after-move";

    wil::com_ptr<IFileSystemIO> dummyIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
    {
        clearMutationEnv();
        static_cast<void>(TakeFileOpsBridgeMutateDestinationBeforeMoveCleanupAttemptsForSelfTest());
        Fail(L"Floodgate bridge file MOVE cleanup corruption test could not query dummy IFileSystemIO.");
        return true;
    }

    wil::com_ptr<IFileSystemIO> localIo;
    if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
    {
        clearMutationEnv();
        static_cast<void>(TakeFileOpsBridgeMutateDestinationBeforeMoveCleanupAttemptsForSelfTest());
        Fail(L"Floodgate bridge file MOVE cleanup corruption test could not query local IFileSystemIO.");
        return true;
    }

    if (state.stepState == 0)
    {
        clearMutationEnv();
        static_cast<void>(TakeFileOpsBridgeMutateDestinationBeforeMoveCleanupAttemptsForSelfTest());
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        const FileSystemFlags dummyCleanupFlags =
            static_cast<FileSystemFlags>(static_cast<uint32_t>(FILESYSTEM_FLAG_RECURSIVE) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY));
        static_cast<void>(state.fsDummy->DeleteItem(dummyRoot.c_str(), dummyCleanupFlags, nullptr, nullptr, nullptr));

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, kSourceContents) ||
            ! RecreateEmptyDirectory(localDest))
        {
            clearMutationEnv();
            Fail(L"Floodgate bridge file MOVE cleanup corruption test failed to seed source/destination.");
            return true;
        }

        const std::wstring localDestFileText = ToPluginPathText(localDestFile);
        const std::wstring corruptPayload(kCorruptContents.begin(), kCorruptContents.end());
        if (! SetEnvironmentVariableW(kMutateDestinationPathEnv, localDestFileText.c_str()) ||
            ! SetEnvironmentVariableW(kMutateDestinationPayloadEnv, corruptPayload.c_str()))
        {
            clearMutationEnv();
            Fail(L"Floodgate bridge file MOVE cleanup corruption test failed to arm destination mutation.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyFile)},
                                                 localDest,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            clearMutationEnv();
            static_cast<void>(TakeFileOpsBridgeMutateDestinationBeforeMoveCleanupAttemptsForSelfTest());
            Fail(L"Floodgate bridge file MOVE cleanup corruption test failed to start the cross-FS move.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    clearMutationEnv();
    const unsigned long mutationAttempts = TakeFileOpsBridgeMutateDestinationBeforeMoveCleanupAttemptsForSelfTest();
    if (mutationAttempts != 1u)
    {
        Fail(std::format(L"Floodgate bridge file MOVE cleanup corruption test expected one destination mutation, got {}.", mutationAttempts));
        return true;
    }

    if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        Fail(std::format(L"Floodgate bridge file MOVE cleanup corruption test expected ERROR_PARTIAL_COPY, got 0x{:08X}.",
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    std::string text;
    if (! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyFile), text) || text != kSourceContents)
    {
        Fail(L"Floodgate bridge file MOVE cleanup corruption test did not preserve the source bytes.");
        return true;
    }
    if (! ReadFileTextFsIo(localIo, localDestFile, text) || text != kCorruptContents)
    {
        Fail(L"Floodgate bridge file MOVE cleanup corruption test did not retain the mutated destination bytes.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateCrossFsFileMoveCleanupCorruption", L"", 0, completed->second.progressCompletedBytes, 0, completed->second.hr);
    NextStep(state, SelfTestState::Step::Floodgate_CrossFsDirectoryMoveCleanupPreservesChangedSource);
    return false;
}
case SelfTestState::Step::Floodgate_CrossFsDirectoryMoveCleanupPreservesChangedSource:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
        ReleaseFileOpsBridgeMoveManifestTakePauseForSelfTest();
        SetFileOpsBridgeMoveManifestTakePauseForSelfTest(false);
        Fail(L"Floodgate_CrossFsDirectoryMoveCleanupPreservesChangedSource timed out.");
        return true;
    }

    const std::filesystem::path sourceRoot = state.tempRoot / L"floodgate-crossfs-dir-cleanup-src";
    const std::filesystem::path sourceTree = sourceRoot / L"tree";
    const std::filesystem::path sourceLevel = sourceTree / L"level-one";
    const std::filesystem::path stableFile = sourceLevel / L"stable.txt";
    const std::wstring dummyRoot = L"/floodgate-crossfs-dir-cleanup";
    const std::wstring dummyTree = dummyRoot + L"/tree";
    const std::wstring dummyStable = dummyTree + L"/level-one/stable.txt";
    constexpr std::string_view kStableBefore = "stable-before-cleanup";

    wil::com_ptr<IFileSystemIO> localIo;
    if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
    {
        ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
        Fail(L"Floodgate bridge directory MOVE cleanup test could not query local IFileSystemIO.");
        return true;
    }

    if (state.stepState == 0)
    {
        ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
        ReleaseFileOpsBridgeMoveManifestTakePauseForSelfTest();
        SetFileOpsBridgeMoveManifestTakePauseForSelfTest(false);
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        const FileSystemFlags dummyCleanupFlags =
            static_cast<FileSystemFlags>(static_cast<uint32_t>(FILESYSTEM_FLAG_RECURSIVE) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY));
        static_cast<void>(state.fsDummy->DeleteItem(dummyRoot.c_str(), dummyCleanupFlags, nullptr, nullptr, nullptr));

        std::error_code ec;
        std::filesystem::remove_all(sourceRoot, ec);
        ec.clear();
        std::filesystem::create_directories(sourceLevel, ec);
        if (ec || ! std::filesystem::exists(sourceLevel, ec) || ec || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) ||
            ! WriteFileTextFsIo(localIo, stableFile, kStableBefore))
        {
            ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
            SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
            ReleaseFileOpsBridgeMoveManifestTakePauseForSelfTest();
            SetFileOpsBridgeMoveManifestTakePauseForSelfTest(false);
            Fail(L"Floodgate bridge directory MOVE cleanup test failed to seed source/destination.");
            return true;
        }

        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(true);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceTree},
                                                 std::filesystem::path(dummyRoot),
                                                 FILESYSTEM_FLAG_RECURSIVE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
            SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
            ReleaseFileOpsBridgeMoveManifestTakePauseForSelfTest();
            SetFileOpsBridgeMoveManifestTakePauseForSelfTest(false);
            Fail(L"Floodgate bridge directory MOVE cleanup test failed to start the cross-FS move.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (! HasFileOpsBridgeMoveSourceCleanupPauseEnteredForSelfTest())
        {
            const auto completedBeforePause = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
            if (completedBeforePause != state.completedTasks.end())
            {
                ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
                SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
                ReleaseFileOpsBridgeMoveManifestTakePauseForSelfTest();
                SetFileOpsBridgeMoveManifestTakePauseForSelfTest(false);
                Fail(std::format(L"Floodgate bridge directory MOVE cleanup test completed before cleanup pause: 0x{:08X}.",
                                 static_cast<unsigned long>(completedBeforePause->second.hr)));
                return true;
            }
            return false;
        }

        const uint64_t peakEntries = GetFileOpsBridgeMoveManifestPeakEntriesForSelfTest();
        const uint64_t currentEntries = GetFileOpsBridgeMoveManifestCurrentEntriesForSelfTest();
        if (peakEntries < 3u || currentEntries != peakEntries)
        {
            ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
            SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
            Fail(std::format(L"Floodgate bridge directory MOVE manifest did not retain the copied tree before cleanup. peak={}, current={}.",
                             peakEntries,
                             currentEntries));
            return true;
        }

        FileSystemBasicInformation sourceRootBasic{};
        sourceRootBasic.sizeBytes = sizeof(FileSystemBasicInformation);
        const std::wstring sourceTreeText = ToPluginPathText(sourceTree);
        if (FAILED(localIo->GetFileBasicInformation(sourceTreeText.c_str(), &sourceRootBasic)))
        {
            ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
            SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
            Fail(L"Floodgate bridge directory MOVE cleanup test failed to query source-root basic information.");
            return true;
        }
        sourceRootBasic.lastWriteTime -= 10'000'000ll;
        if (FAILED(localIo->SetFileBasicInformation(sourceTreeText.c_str(), &sourceRootBasic)))
        {
            ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
            SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
            Fail(L"Floodgate bridge directory MOVE cleanup test failed to mutate only source-root basic information.");
            return true;
        }

        SetFileOpsBridgeMoveManifestTakePauseForSelfTest(true);
        ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (! HasFileOpsBridgeMoveManifestTakePauseEnteredForSelfTest())
        {
            return false;
        }

        const uint64_t peakEntries = GetFileOpsBridgeMoveManifestPeakEntriesForSelfTest();
        const uint64_t currentEntries = GetFileOpsBridgeMoveManifestCurrentEntriesForSelfTest();
        if (currentEntries == 0u || currentEntries >= peakEntries)
        {
            ReleaseFileOpsBridgeMoveManifestTakePauseForSelfTest();
            SetFileOpsBridgeMoveManifestTakePauseForSelfTest(false);
            Fail(std::format(L"Floodgate bridge MOVE manifest did not shrink during cleanup. peak={}, current={}.", peakEntries, currentEntries));
            return true;
        }

        ReleaseFileOpsBridgeMoveManifestTakePauseForSelfTest();
        state.stepState = 3;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
    SetFileOpsBridgeMoveManifestTakePauseForSelfTest(false);

    if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        Fail(std::format(L"Floodgate bridge directory MOVE cleanup test expected ERROR_PARTIAL_COPY, got 0x{:08X}.",
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
    {
        Fail(L"Floodgate bridge directory MOVE cleanup test could not query dummy IFileSystemIO.");
        return true;
    }

    std::string text;
    if (ReadFileTextFsIo(localIo, stableFile, text))
    {
        Fail(L"Floodgate bridge directory MOVE cleanup test left the unchanged source file behind.");
        return true;
    }
    std::error_code ec;
    if (! std::filesystem::exists(sourceTree, ec) || ec)
    {
        Fail(L"Floodgate bridge directory MOVE cleanup test removed the source shell whose basic information changed.");
        return true;
    }
    if (std::filesystem::exists(sourceLevel, ec) || ec)
    {
        Fail(L"Floodgate bridge directory MOVE cleanup test did not remove the matching child directory.");
        return true;
    }

    if (! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyStable), text) || text != kStableBefore)
    {
        Fail(L"Floodgate bridge directory MOVE cleanup test did not copy the stable file to the destination.");
        return true;
    }
    if (GetFileOpsBridgeMoveManifestCurrentEntriesForSelfTest() != 0u)
    {
        Fail(std::format(L"Floodgate bridge directory MOVE cleanup left {} manifest entries after completion.",
                         GetFileOpsBridgeMoveManifestCurrentEntriesForSelfTest()));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateCrossFsDirectoryMoveCleanup",
                      L"root-basic-info-mutated",
                      0,
                      GetFileOpsBridgeMoveManifestPeakEntriesForSelfTest(),
                      GetFileOpsBridgeMoveManifestCurrentEntriesForSelfTest(),
                      completed->second.hr);
    NextStep(state, SelfTestState::Step::Floodgate_LocalWriterOverwriteIsStaged);
    return false;
}
case SelfTestState::Step::Floodgate_LocalWriterOverwriteIsStaged:
{
    wil::com_ptr<IFileSystemIO> localIo;
    if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
    {
        Fail(L"Floodgate local overwrite writer test could not query local IFileSystemIO.");
        return true;
    }

    const std::filesystem::path targetPath = state.tempRoot / L"floodgate-local-overwrite-writer.txt";
    const std::filesystem::path siblingTempPrefix = targetPath.parent_path() / std::filesystem::path(targetPath.filename().wstring() + L".~rs-write-");
    std::error_code ec;
    std::filesystem::remove(targetPath, ec);

    constexpr std::string_view kOriginal = "original-preserved";
    constexpr std::string_view kPartial = "partial-new-bytes";
    constexpr std::string_view kReplacement = "replacement-bytes";
    if (! WriteFileTextFsIo(localIo, targetPath, kOriginal))
    {
        Fail(L"Floodgate local overwrite writer test failed to seed the original file.");
        return true;
    }

    {
        wil::com_ptr<IFileWriter> writer;
        const std::wstring pathText = ToPluginPathText(targetPath);
        const HRESULT createHr = localIo->CreateFileWriter(pathText.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.put());
        if (FAILED(createHr) || ! writer)
        {
            Fail(std::format(L"Floodgate local overwrite writer test failed to open overwrite writer. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
            return true;
        }

        unsigned long written = 0;
        const HRESULT writeHr = writer->Write(kPartial.data(), static_cast<unsigned long>(kPartial.size()), &written);
        if (FAILED(writeHr) || written != static_cast<unsigned long>(kPartial.size()))
        {
            Fail(std::format(L"Floodgate local overwrite writer test failed to write partial bytes. hr=0x{:08X}", static_cast<unsigned long>(writeHr)));
            return true;
        }
    }

    std::string afterAbort;
    if (! ReadFileTextFsIo(localIo, targetPath, afterAbort) || afterAbort != kOriginal)
    {
        Fail(L"Floodgate local overwrite writer abort did not preserve the original destination bytes.");
        return true;
    }

    {
        wil::com_ptr<IFileWriter> writer;
        const std::wstring pathText = ToPluginPathText(targetPath);
        const HRESULT createHr = localIo->CreateFileWriter(pathText.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.put());
        if (FAILED(createHr) || ! writer)
        {
            Fail(std::format(L"Floodgate local overwrite writer test failed to open commit writer. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
            return true;
        }

        unsigned long written = 0;
        const HRESULT writeHr = writer->Write(kReplacement.data(), static_cast<unsigned long>(kReplacement.size()), &written);
        if (FAILED(writeHr) || written != static_cast<unsigned long>(kReplacement.size()))
        {
            Fail(std::format(L"Floodgate local overwrite writer test failed to write replacement bytes. hr=0x{:08X}", static_cast<unsigned long>(writeHr)));
            return true;
        }

        const HRESULT commitHr = writer->Commit();
        if (FAILED(commitHr))
        {
            Fail(std::format(L"Floodgate local overwrite writer commit failed. hr=0x{:08X}", static_cast<unsigned long>(commitHr)));
            return true;
        }
    }

    std::string afterCommit;
    if (! ReadFileTextFsIo(localIo, targetPath, afterCommit) || afterCommit != kReplacement)
    {
        Fail(L"Floodgate local overwrite writer commit did not replace the destination bytes.");
        return true;
    }

    const DWORD writerCommitAttributes = ::GetFileAttributesW(targetPath.c_str());
    if (writerCommitAttributes == INVALID_FILE_ATTRIBUTES ||
        (writerCommitAttributes & (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY | FILE_ATTRIBUTE_READONLY)) != 0u)
    {
        Fail(L"Floodgate local overwrite writer leaked staged temp attributes to the final file.");
        return true;
    }

    for (const auto& entry : std::filesystem::directory_iterator(targetPath.parent_path(), ec))
    {
        if (ec)
        {
            break;
        }

        const std::wstring entryPath = entry.path().wstring();
        if (entryPath.starts_with(siblingTempPrefix.wstring()))
        {
            Fail(L"Floodgate local overwrite writer left a staged temp file behind.");
            return true;
        }
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateLocalWriterOverwriteIsStaged", L"", kOriginal.size(), kReplacement.size(), 0, S_OK);
    NextStep(state, SelfTestState::Step::Floodgate_LocalCopyOverwriteIsStaged);
    return false;
}
case SelfTestState::Step::Floodgate_LocalCopyOverwriteIsStaged:
{
    wil::com_ptr<IFileSystemIO> localIo;
    if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
    {
        Fail(L"Floodgate local copy overwrite test could not query local IFileSystemIO.");
        return true;
    }

    const std::filesystem::path sourcePath = state.tempRoot / L"floodgate-local-copy-overwrite-source.txt";
    const std::filesystem::path targetPath = state.tempRoot / L"floodgate-local-copy-overwrite-target.txt";
    const std::filesystem::path siblingTempPrefix = targetPath.parent_path() / std::filesystem::path(targetPath.filename().wstring() + L".rs_copy_tmp_");

    constexpr std::string_view kOriginal = "copy-original-preserved";
    constexpr std::string_view kReplacement = "copy-replacement-bytes";

    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailFired.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailFired.data(), nullptr));
    auto clearStagedCopyEnv = wil::scope_exit([]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailPath.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailFired.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailPath.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailFired.data(), nullptr));
    });

    std::error_code ec;
    std::filesystem::remove(sourcePath, ec);
    ec.clear();
    std::filesystem::remove(targetPath, ec);

    if (! WriteFileTextFsIo(localIo, sourcePath, kReplacement) || ! WriteFileTextFsIo(localIo, targetPath, kOriginal))
    {
        Fail(L"Floodgate local copy overwrite test failed to seed source/destination.");
        return true;
    }

    const DWORD sourceAttributes = ::GetFileAttributesW(sourcePath.c_str());
    if (sourceAttributes == INVALID_FILE_ATTRIBUTES || ::SetFileAttributesW(sourcePath.c_str(), sourceAttributes | FILE_ATTRIBUTE_READONLY) == 0)
    {
        Fail(L"Floodgate local copy overwrite test failed to mark the source read-only.");
        return true;
    }

    if (::SetFileAttributesW(targetPath.c_str(), FILE_ATTRIBUTE_READONLY) == 0)
    {
        Fail(L"Floodgate local copy overwrite test failed to mark the destination read-only.");
        return true;
    }

    if (! SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailPath.data(), targetPath.c_str()))
    {
        Fail(L"Floodgate local copy overwrite test failed to set the staged-copy promote failure hook.");
        return true;
    }

    const FileSystemFlags overwriteFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
    const HRESULT failHr = state.fsLocal->CopyItem(sourcePath.c_str(), targetPath.c_str(), overwriteFlags, nullptr, nullptr, nullptr);
    if (failHr != HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED))
    {
        Fail(std::format(L"Floodgate local copy overwrite injected failure expected ERROR_ACCESS_DENIED, got 0x{:08X}.", static_cast<unsigned long>(failHr)));
        return true;
    }

    if (GetEnvVarTrimmed(kSelfTestEnvStagedCopyPromoteFailFired) != L"1")
    {
        Fail(L"Floodgate local copy overwrite test did not fire the staged-copy promote failure hook.");
        return true;
    }

    std::string afterFailure;
    if (! ReadFileTextFsIo(localIo, targetPath, afterFailure) || afterFailure != kOriginal)
    {
        Fail(L"Floodgate local copy overwrite failure did not preserve original destination bytes.");
        return true;
    }

    const DWORD afterFailureAttributes = ::GetFileAttributesW(targetPath.c_str());
    if (afterFailureAttributes == INVALID_FILE_ATTRIBUTES || (afterFailureAttributes & FILE_ATTRIBUTE_READONLY) == 0u)
    {
        Fail(L"Floodgate local copy overwrite failure did not preserve the destination read-only attribute.");
        return true;
    }

    for (const auto& entry : std::filesystem::directory_iterator(targetPath.parent_path(), ec))
    {
        if (ec)
        {
            break;
        }

        const std::wstring entryPath = entry.path().wstring();
        if (entryPath.starts_with(siblingTempPrefix.wstring()))
        {
            Fail(L"Floodgate local copy overwrite failure left a staged temp file behind.");
            return true;
        }
    }

    static_cast<void>(::SetFileAttributesW(targetPath.c_str(), afterFailureAttributes & ~FILE_ATTRIBUTE_READONLY));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailFired.data(), nullptr));

    if (! SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailPath.data(), targetPath.c_str()))
    {
        Fail(L"Floodgate local copy overwrite test failed to set the final-attributes failure hook.");
        return true;
    }

    const HRESULT finalAttributesFailHr = state.fsLocal->CopyItem(sourcePath.c_str(), targetPath.c_str(), overwriteFlags, nullptr, nullptr, nullptr);
    if (finalAttributesFailHr != HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED))
    {
        Fail(std::format(L"Floodgate local copy overwrite final-attributes failure expected ERROR_ACCESS_DENIED, got 0x{:08X}.",
                         static_cast<unsigned long>(finalAttributesFailHr)));
        return true;
    }
    if (GetEnvVarTrimmed(kSelfTestEnvFinalAttributesFailFired) != L"1")
    {
        Fail(L"Floodgate local copy overwrite test did not fire the final-attributes failure hook.");
        return true;
    }

    const DWORD afterFinalAttributesFailure = ::GetFileAttributesW(targetPath.c_str());
    if (afterFinalAttributesFailure == INVALID_FILE_ATTRIBUTES ||
        ::SetFileAttributesW(targetPath.c_str(), afterFinalAttributesFailure & ~FILE_ATTRIBUTE_READONLY) == 0 ||
        ! WriteFileTextFsIo(localIo, targetPath, kOriginal, true) || ::SetFileAttributesW(targetPath.c_str(), FILE_ATTRIBUTE_READONLY) == 0)
    {
        Fail(L"Floodgate local copy overwrite test failed to reseed after the committed final-attributes failure.");
        return true;
    }
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailFired.data(), nullptr));

    const HRESULT successHr = state.fsLocal->CopyItem(sourcePath.c_str(), targetPath.c_str(), overwriteFlags, nullptr, nullptr, nullptr);
    if (FAILED(successHr))
    {
        Fail(std::format(L"Floodgate local copy overwrite success path failed: 0x{:08X}.", static_cast<unsigned long>(successHr)));
        return true;
    }

    std::string afterSuccess;
    if (! ReadFileTextFsIo(localIo, targetPath, afterSuccess) || afterSuccess != kReplacement)
    {
        Fail(L"Floodgate local copy overwrite success path did not replace the destination bytes.");
        return true;
    }

    const DWORD afterSuccessAttributes = ::GetFileAttributesW(targetPath.c_str());
    if (afterSuccessAttributes == INVALID_FILE_ATTRIBUTES || (afterSuccessAttributes & FILE_ATTRIBUTE_READONLY) == 0u ||
        (afterSuccessAttributes & (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY)) != 0u)
    {
        Fail(L"Floodgate local copy overwrite success path did not preserve source read-only without temp attributes.");
        return true;
    }
    static_cast<void>(::SetFileAttributesW(targetPath.c_str(), afterSuccessAttributes & ~FILE_ATTRIBUTE_READONLY));
    static_cast<void>(::SetFileAttributesW(sourcePath.c_str(), sourceAttributes));

    ec.clear();
    for (const auto& entry : std::filesystem::directory_iterator(targetPath.parent_path(), ec))
    {
        if (ec)
        {
            break;
        }

        const std::wstring entryPath = entry.path().wstring();
        if (entryPath.starts_with(siblingTempPrefix.wstring()))
        {
            Fail(L"Floodgate local copy overwrite success path left a staged temp file behind.");
            return true;
        }
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateLocalCopyOverwriteIsStaged", L"", kOriginal.size(), kReplacement.size(), 0, S_OK);
    NextStep(state, SelfTestState::Step::Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker);
    return false;
}
case SelfTestState::Step::Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker:
{
    using PfnRunDebugSelfTests = HRESULT(__stdcall*)(unsigned int*, unsigned int*);

    std::array<wchar_t, MAX_PATH> exePath{};
    const DWORD exePathLength = GetModuleFileNameW(nullptr, exePath.data(), static_cast<DWORD>(exePath.size()));
    if (exePathLength == 0 || exePathLength >= exePath.size())
    {
        Fail(L"Riptide shared scheduler shutdown test failed to resolve the executable path.");
        return true;
    }

    const std::filesystem::path dllPath =
        std::filesystem::path(std::wstring_view(exePath.data(), exePathLength)).parent_path() / L"Plugins" / L"FileSystem.dll";
    wil::unique_hmodule module(LoadLibraryExW(dllPath.c_str(), nullptr, 0));
    if (! module)
    {
        Fail(std::format(L"Riptide shared scheduler shutdown test failed to load FileSystem.dll from '{}'.", dllPath.wstring()));
        return true;
    }

    const FARPROC runDebugSelfTestsProc = GetProcAddress(module.get(), "RedSalamanderFileSystemDebugSelfTests");
#pragma warning(push)
#pragma warning(disable : 4191) // Win32 exports are resolved as FARPROC; validate null before calling the typed debug-only export.
    const auto runDebugSelfTests = reinterpret_cast<PfnRunDebugSelfTests>(runDebugSelfTestsProc);
#pragma warning(pop)
    if (runDebugSelfTests == nullptr)
    {
        Fail(L"Riptide shared scheduler shutdown test could not resolve RedSalamanderFileSystemDebugSelfTests.");
        return true;
    }

    unsigned int passed = 0;
    unsigned int failed = 0;
    const HRESULT hr = runDebugSelfTests(&passed, &failed);
    if (FAILED(hr) || failed != 0u)
    {
        Fail(std::format(L"Riptide shared scheduler shutdown test expected FileSystem debug selftests to pass, got hr=0x{:08X}, passed={}, failed={}.",
                         static_cast<unsigned long>(hr),
                         passed,
                         failed));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.RiptideSharedSchedulerShutdownQuietPoint", L"", passed, failed, 0u, hr);
    NextStep(state, SelfTestState::Step::Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker);
    return false;
}
case SelfTestState::Step::Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker:
{
    if (state.fileOps == nullptr)
    {
        Fail(L"Riptide host scheduler shutdown test has no FileOperationState.");
        return true;
    }

    if (! RunFileOpsPerItemSchedulerShutdownQuietPointSelfTestForSelfTest(*state.fileOps))
    {
        Fail(L"Riptide host scheduler shutdown test observed shutdown/WaitJob returning before the worker callback exited.");
        return true;
    }
    if (! RunFileOpsBridgeDirectoryBufferValidationSelfTestForSelfTest())
    {
        Fail(L"Riptide bridge directory buffer validation selftest failed.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.RiptideHostSchedulerShutdownQuietPoint", L"", 1u, 0u, 0u, S_OK);
    NextStep(state, SelfTestState::Step::Phase5_PreCalcSettingsApplied);
    return false;
}
