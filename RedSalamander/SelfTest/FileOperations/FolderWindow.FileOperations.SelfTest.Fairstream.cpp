#if defined(FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_A)

    // Operation Fairstream Phase 1 cases (data-loss correctness).
    // Included into the selftest switch in FolderWindow.FileOperations.SelfTest.cpp.

case SelfTestState::Step::Riptide_MoveDistinctSameSizeFilePreservesSource:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();

    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Riptide_MoveDistinctSameSizeFilePreservesSource timed out.");
        return true;
    }

    constexpr size_t kFileBytes = 16 * 1024;
    const std::filesystem::path srcRoot = state.tempRoot / L"riptide-same-size-move-src";
    const std::filesystem::path refRoot = state.tempRoot / L"riptide-same-size-move-ref";
    const std::filesystem::path srcFile = srcRoot / L"same.bin";
    const std::filesystem::path srcRef = refRoot / L"source.bin";
    const std::wstring dummyRoot = L"/riptide-same-size-move-dst";
    const std::wstring dummyFile = dummyRoot + L"/same.bin";
    const std::string destinationBytes(kFileBytes, static_cast<char>(0xA7));

    const auto sourceStillIntact = [&]() noexcept { return FileSizeEquals(srcFile, kFileBytes) && FilesEqualBytes(srcFile, srcRef); };
    const auto destinationStillIntact = [&]() noexcept
    {
        wil::com_ptr<IFileSystemIO> io;
        wil::com_ptr<IFileReader> reader;
        if (! state.fsDummy || FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(io.addressof()))) || ! io ||
            FAILED(io->CreateFileReader(dummyFile.c_str(), reader.addressof())) || ! reader)
        {
            return false;
        }

        uint64_t sizeBytes = 0u;
        if (FAILED(reader->GetSize(&sizeBytes)) || sizeBytes != destinationBytes.size())
        {
            return false;
        }
        std::vector<char> actual(destinationBytes.size());
        unsigned long bytesRead = 0u;
        return SUCCEEDED(reader->Read(actual.data(), static_cast<unsigned long>(actual.size()), &bytesRead)) && bytesRead == actual.size() &&
               std::ranges::equal(actual, destinationBytes);
    };

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(refRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
        {
            Fail(L"Riptide same-size move test failed to reset directories.");
            return true;
        }
        if (! WriteFilledTestFile(srcFile, kFileBytes, 0x31) || ! WriteFilledTestFile(srcRef, kFileBytes, 0x31) ||
            ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, destinationBytes, true))
        {
            Fail(L"Riptide same-size move test failed to seed distinct source/destination files.");
            return true;
        }

        const FileSystemFlags flags = FILESYSTEM_FLAG_NONE;
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFile},
                                                 std::filesystem::path(dummyRoot),
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
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
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
                {
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

        if (completed->second.hr != S_FALSE)
        {
            Fail(
                std::format(L"Riptide same-size move expected S_FALSE after intentional Skip, got 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
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
        NextStep(state, SelfTestState::Step::Riptide_ReparseNativeMoveRelocatesLinkObject);
        return false;
    }

    return false;
}
case SelfTestState::Step::Riptide_ReparseNativeMoveRelocatesLinkObject:
{
    const std::filesystem::path targetDir = state.tempRoot / L"riptide-reparse-native-target";
    const std::filesystem::path srcRoot = state.tempRoot / L"riptide-reparse-native-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"riptide-reparse-native-dst";
    const std::filesystem::path srcLink = srcRoot / L"Link";
    const std::filesystem::path dstFile = dstRoot / L"Link";
    const std::filesystem::path targetFile = targetDir / L"target-sentinel.bin";

    if (! RecreateEmptyDirectory(targetDir) || ! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) ||
        ! WriteTestFile(targetFile, 4 * 1024) || ! WriteTestFile(dstFile, 2 * 1024) || ! TryCreateJunction(srcLink, targetDir))
    {
        Fail(L"Riptide reparse native-only fixture setup failed.");
        return true;
    }

    FileSystemOptions options{};
    options.sizeBytes = sizeof(options);
    options.moveMode  = FILESYSTEM_MOVE_NATIVE_ONLY;
    FileOpsRecursiveProgressRecorder callback{};
    const HRESULT hr = state.fsLocal->MoveItem(srcLink.c_str(),
                                                dstFile.c_str(),
                                                static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                &options,
                                                &callback,
                                                nullptr);

    std::error_code ec;
    const bool sourceStillPresent = std::filesystem::exists(srcLink, ec);
    if (hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) || callback.issueCount != 1u ||
        callback.lastIssueStatus != HRESULT_FROM_WIN32(ERROR_DATATYPE_MISMATCH) || ec || ! sourceStillPresent ||
        ! FileSizeEquals(dstFile, 2 * 1024) || ! FileSizeEquals(targetFile, 4 * 1024))
    {
        Fail(std::format(L"Native-only link-on-file Move must report a typed mismatch, honor Skip, and preserve both objects "
                         L"(hr=0x{:08X}, issue=0x{:08X}, count={}).",
                         static_cast<unsigned long>(hr),
                         static_cast<unsigned long>(callback.lastIssueStatus),
                         callback.issueCount));
        return true;
    }
    if (callback.completedCount != 1u)
    {
        Fail(std::format(L"Native-only reparse Move expected exactly 1 item completion callback, saw {}.", callback.completedCount));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.RiptideReparseNativeRelocation",
                      L"shape=junction-on-file-type-mismatch;source-retained",
                      1,
                      0,
                      0,
                      hr);
    NextStep(state, SelfTestState::Step::Fairstream_ManagedMovePreservesUncopiedNewFiles);
    return false;
}

case SelfTestState::Step::Fairstream_ManagedMovePreservesUncopiedNewFiles:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
        Fail(L"Fairstream_ManagedMovePreservesUncopiedNewFiles timed out.");
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
    const std::filesystem::path srcNew = srcDone / L"zz_new.bin";
    const std::filesystem::path dstNew = dstDone / L"zz_new.bin";

    if (state.stepState == 0)
    {
        // An existing destination directory makes this a rename merge: the seed's directory relocates
        // by one rename, and the pause fires before the emptied source root is removed.
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstFoo))
        {
            Fail(L"Move-window test failed to reset directories.");
            return true;
        }
        if (! WriteTestFile(srcSeed, 4 * 1024))
        {
            Fail(L"Move-window test failed to seed the source tree.");
            return true;
        }

        // Pause before the source root's exact cleanup. Its child directory has already been
        // relocated, so a child injected under the old path deterministically exists only in the
        // source and must make the cleanup retain it.
        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(true);
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
            SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
            Fail(L"Move-window test failed to start the move task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        const bool alreadyCompleted = state.taskA.has_value() && state.completedTasks.contains(state.taskA.value());

        if (alreadyCompleted)
        {
            SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
            Fail(L"Rename merge completed without entering the deterministic pre-cleanup pause.");
            return true;
        }
        if (! HasFileOpsBridgeMoveSourceCleanupPauseEnteredForSelfTest())
        {
            return false;
        }

        std::error_code injectEc;
        std::filesystem::create_directories(srcDone, injectEc);
        if (injectEc || ! WriteTestFile(srcNew, 2 * 1024))
        {
            ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
            SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
            Fail(L"Move-window test failed to inject the new source file.");
            return true;
        }

        ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
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

        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Rename merge with an unselected new child expected S_FALSE (Moved; source folder kept), got 0x{:08X}.",
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
        if (! FileSizeEquals(dstSeed, 4 * 1024))
        {
            Fail(L"Move-window test destination tree is incomplete.");
            return true;
        }
        if (std::filesystem::exists(srcSeed, ec) && ! ec)
        {
            Fail(L"Move-window test did not relocate the source file.");
            return true;
        }

        Debug::Perf::Emit(
            L"FileOps.SelfTest.FairstreamMoveWindowPreserved", L"shape=new-file-before-exact-cleanup", 1, 0, 0, completed->second.hr);
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
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstFoo))
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
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
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
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists)
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

        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Prompt-serialization test expected intentional-Skip S_FALSE after mixed Overwrite/Skip decisions, got 0x{:08X}.",
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
    const std::filesystem::path dstLinkSibling = dstFoo / L"Link (2)";
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
                if (prompt->bucket != Task::ConflictBucket::TypeMismatch || ! PromptHasAction(prompt.value(), Task::ConflictAction::KeepBoth) ||
                    PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Reparse-consent test expected a one-shot type-mismatch prompt offering Keep both but not Overwrite.");
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

        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Reparse-consent Skip run expected intentional-Skip S_FALSE, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
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

        // Second run: the user may preserve both objects, but a type mismatch never authorizes replacing
        // the existing directory with a link object.
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
                if (prompt->bucket != Task::ConflictBucket::TypeMismatch || ! PromptHasAction(prompt.value(), Task::ConflictAction::KeepBoth) ||
                    PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Reparse-consent Keep Both run did not preserve the typed type-mismatch action surface.");
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::KeepBoth, false);
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
            Fail(std::format(L"Reparse-consent Keep Both run failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        std::error_code destinationEc;
        if (! std::filesystem::exists(dstInner, destinationEc) || destinationEc)
        {
            Fail(L"Reparse-consent Keep Both run changed the existing non-empty directory.");
            return true;
        }

        const auto linkTarget = TryGetDirectoryReparseTargetAbsolute(dstLinkSibling);
        if (! linkTarget.has_value() || NormalizePathForCompare(linkTarget.value()) != NormalizePathForCompare(targetDir.wstring()))
        {
            std::wstring observedNames;
            std::error_code enumerateEc;
            for (std::filesystem::directory_iterator iterator(dstFoo, enumerateEc), end; ! enumerateEc && iterator != end;
                 iterator.increment(enumerateEc))
            {
                if (! observedNames.empty())
                {
                    observedNames.append(L", ");
                }
                observedNames.append(iterator->path().filename().wstring());
            }
            Fail(std::format(L"Reparse-consent Keep Both run did not publish the junction under a sibling name: expected='{}' actual='{}' entries=[{}].",
                             targetDir.wstring(),
                             linkTarget.value_or(L"<missing>"),
                             observedNames));
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
    const std::filesystem::path srcOpposite = srcRoot / L"Opposite";
    const std::filesystem::path dstOpposite = dstRoot / L"Opposite";

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(),
                                                 R"json({"reparsePointPolicy":"preserve","concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        if (! RecreateEmptyDirectory(targetDir) || ! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) ||
            ! RecreateEmptyDirectory(dstLink) || ! RecreateEmptyDirectory(srcOpposite))
        {
            Fail(L"Conflict-bucket apply-to-all test failed to prepare directories.");
            return true;
        }

        if (! WriteFilledTestFile(targetDir / L"sentinel.bin", 4096, 0xA1) || ! WriteFilledTestFile(srcPlain, 2048, 0x5A) ||
            ! WriteFilledTestFile(dstPlain, 1024, 0x22) || ! WriteFilledTestFile(dstInner, 512, 0x39) ||
            ! WriteFilledTestFile(srcOpposite / L"child.bin", 768, 0x71) || ! WriteFilledTestFile(dstOpposite, 256, 0x17))
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
                                                 {srcPlain, srcLink, srcOpposite},
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
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
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
                if (prompt->bucket != Task::ConflictBucket::TypeMismatch || ! PromptHasAction(prompt.value(), Task::ConflictAction::KeepBoth) ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip) || ! PromptHasAction(prompt.value(), Task::ConflictAction::SkipAll) ||
                    PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite) || ! prompt->applyToAllEligible)
                {
                    Fail(L"Conflict-bucket apply-to-all test expected a separately scoped link-to-directory type-mismatch prompt.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::SkipAll, false);
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
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::TypeMismatch || ! prompt->sourceMetadata.isDirectory ||
                    prompt->destinationMetadata.isDirectory || prompt->destinationMetadata.isLink ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip) || PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Conflict-bucket apply-to-all test expected the inverse directory-to-file type mismatch to prompt separately.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.stepState = 4;
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed != state.completedTasks.end())
        {
            Fail(std::format(L"Conflict-bucket apply-to-all test completed without prompting for the inverse type mismatch: 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        return false;
    }

    if (state.stepState == 4)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Conflict-bucket apply-to-all test expected intentional-Skip S_FALSE after skipping the typed conflicts, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 3u)
        {
            Fail(std::format(L"Conflict-bucket apply-to-all test expected exactly three class/orientation-scoped prompts, saw {}.",
                             completed->second.conflictPromptCount));
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
        const DWORD oppositeSourceAttributes      = GetFileAttributesW(srcOpposite.c_str());
        const DWORD oppositeDestinationAttributes = GetFileAttributesW(dstOpposite.c_str());
        if (oppositeSourceAttributes == INVALID_FILE_ATTRIBUTES || ! WI_IsFlagSet(oppositeSourceAttributes, FILE_ATTRIBUTE_DIRECTORY) ||
            oppositeDestinationAttributes == INVALID_FILE_ATTRIBUTES || WI_IsFlagSet(oppositeDestinationAttributes, FILE_ATTRIBUTE_DIRECTORY))
        {
            Fail(L"Conflict-bucket apply-to-all test changed the inverse directory-to-file mismatch after Skip.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateConflictBucketApplyAll", L"", 3, 0, 0, completed->second.hr);
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
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"preserve"})json"));

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
                if (prompt->bucket != Task::ConflictBucket::TypeMismatch || ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip) ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::KeepBoth) ||
                    PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Empty-dir reparse consent test expected a type-mismatch prompt offering Skip and Keep Both but not Overwrite.");
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

        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Empty-dir reparse consent Skip run expected S_FALSE, got 0x{:08X}.",
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
        Fail(L"Reparse move merge test failed to load the legacy followTargets migration fixture.");
        return true;
    }
    auto restoreReparsePolicy = wil::scope_exit([&]() noexcept
    { static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"preserve"})json")); });

    const HRESULT moveHr =
        state.fsLocal->MoveItem(srcLink.c_str(), dstLink.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), nullptr, nullptr, nullptr);
    if (moveHr != HRESULT_FROM_WIN32(ERROR_DATATYPE_MISMATCH))
    {
        Fail(std::format(L"Reparse move merge test expected typed mismatch, got 0x{:08X}.", static_cast<unsigned long>(moveHr)));
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

    Debug::Perf::Emit(L"FileOps.SelfTest.RiptideReparseMoveNoMerge", L"policy=legacyFollowMigratedToSkip", 1, 0, 0, moveHr);
    NextStep(state, SelfTestState::Step::Fairstream_MovedTreeKeepsLiteralLinks);
    return false;
}
case SelfTestState::Step::Fairstream_MovedTreeKeepsLiteralLinks:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"Fairstream_MovedTreeKeepsLiteralLinks timed out.");
        return true;
    }

    // C3: Preserve copies every link literally. A moved or copied tree's junctions keep their stored
    // target text, so an absolute in-tree junction keeps naming the source location (and dangles
    // after a Move), a Keep Both rename of a nested target never rewrites the links naming it, a
    // skipped target leaves a literal link next to a retained source target, and a wide tree of
    // links copies with no link queue or ceiling.
    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-move-links-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-move-links-dst";
    const std::filesystem::path srcTree = srcRoot / L"Tree";
    const std::filesystem::path dstTree = dstRoot / L"Tree";
    const std::filesystem::path srcBackwardTarget = srcTree / L"00-backward-target";
    const std::filesystem::path srcBackwardLink = srcTree / L"10-backward-link";
    const std::filesystem::path srcForwardLink = srcTree / L"20-forward-link";
    const std::filesystem::path srcForwardTarget = srcTree / L"30-forward-target";
    const std::filesystem::path srcCycleA = srcTree / L"40-cycle-a";
    const std::filesystem::path srcCycleB = srcTree / L"50-cycle-b";
    const std::filesystem::path dstBackwardTarget = dstTree / srcBackwardTarget.filename();
    const std::filesystem::path dstBackwardLink = dstTree / srcBackwardLink.filename();
    const std::filesystem::path dstForwardLink = dstTree / srcForwardLink.filename();
    const std::filesystem::path dstForwardCollision = dstTree / srcForwardTarget.filename();
    const std::filesystem::path dstForwardTarget = dstTree / L"30-forward-target (2)";
    const std::filesystem::path dstCycleA = dstTree / srcCycleA.filename();
    const std::filesystem::path dstCycleB = dstTree / srcCycleB.filename();
    const std::filesystem::path skipSrcRoot = state.tempRoot / L"fairstream-move-links-retained-src";
    const std::filesystem::path skipDstRoot = state.tempRoot / L"fairstream-move-links-retained-dst";
    const std::filesystem::path skipSrcTree = skipSrcRoot / L"Tree";
    const std::filesystem::path skipDstTree = skipDstRoot / L"Tree";
    const std::filesystem::path skipSrcLink = skipSrcTree / L"10-forward-link";
    const std::filesystem::path skipSrcTarget = skipSrcTree / L"20-target";
    const std::filesystem::path skipDstLink = skipDstTree / skipSrcLink.filename();
    const std::filesystem::path skipDstTargetCollision = skipDstTree / skipSrcTarget.filename();
    const std::filesystem::path copySrcRoot = state.tempRoot / L"fairstream-copy-links-src";
    const std::filesystem::path copyDstRoot = state.tempRoot / L"fairstream-copy-links-dst";
    const std::filesystem::path copySrcTree = copySrcRoot / L"Tree";
    const std::filesystem::path copyDstTree = copyDstRoot / L"Tree";
    const std::filesystem::path copySrcLink = copySrcTree / L"10-forward-link";
    const std::filesystem::path copySrcTarget = copySrcTree / L"20-target";
    const std::filesystem::path copyDstLink = copyDstTree / copySrcLink.filename();
    const std::filesystem::path copyDstTargetCollision = copyDstTree / copySrcTarget.filename();
    const std::filesystem::path copyDstTarget = copyDstTree / L"20-target (2)";
    const std::filesystem::path wideSrcRoot = state.tempRoot / L"fairstream-copy-wide-links-src";
    const std::filesystem::path wideDstRoot = state.tempRoot / L"fairstream-copy-wide-links-dst";
    const std::filesystem::path wideSrcTree = wideSrcRoot / L"Tree";
    const std::filesystem::path wideDstTree = wideDstRoot / L"Tree";
    const std::filesystem::path wideTarget = wideSrcTree / L"target";
    constexpr unsigned int kWideLinkCount = 4'200u; // above the retired 4,096-entry link queue

    const auto linkNames = [](const std::filesystem::path& link, const std::filesystem::path& expectedTarget) noexcept -> bool
    {
        const auto target = TryGetDirectoryReparseTargetAbsolute(link);
        return target.has_value() && NormalizePathForCompare(target.value()) == NormalizePathForCompare(expectedTarget.wstring());
    };
    const auto wideLinkName = [](unsigned int index) { return std::format(L"link-{:05}", index); };

    if (state.stepState == 0)
    {
        // Existing dstTree selects the host Managed merge route.
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstTree))
        {
            Fail(L"Literal-link move test failed to reset directories.");
            return true;
        }
        if (! WriteTestFile(srcBackwardTarget / L"back.bin", 4 * 1024) ||
            ! WriteTestFile(srcForwardTarget / L"forward.bin", 6 * 1024) ||
            ! WriteTestFile(dstForwardCollision, 2 * 1024))
        {
            Fail(L"Literal-link move test failed to seed the source tree.");
            return true;
        }
        if (! TryCreateJunction(srcBackwardLink, srcBackwardTarget) ||
            ! TryCreateJunction(srcForwardLink, srcForwardTarget) ||
            ! TryCreateJunction(srcCycleA, srcCycleB) ||
            ! TryCreateJunction(srcCycleB, srcCycleA))
        {
            Fail(L"Literal-link move test failed to create its backward, forward, and cycle junctions.");
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
            Fail(L"Literal-link move test failed to start the move task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        auto* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            if (state.taskA.has_value() && state.completedTasks.contains(state.taskA.value()))
            {
                Fail(L"Literal-link move completed without the nested target Keep Both conflict.");
                return true;
            }
            return false;
        }
        if (! OrdinalString::EqualsNoCasePath(std::filesystem::path(prompt->destinationPath), dstForwardCollision) ||
            prompt->bucket != FolderWindow::FileOperationState::Task::ConflictBucket::TypeMismatch ||
            ! PromptHasAction(prompt.value(), FolderWindow::FileOperationState::Task::ConflictAction::KeepBoth) ||
            PromptHasAction(prompt.value(), FolderWindow::FileOperationState::Task::ConflictAction::Overwrite))
        {
            Fail(L"Literal-link move did not expose a non-overwrite TypeMismatch/Keep Both choice for the nested target component collision.");
            return true;
        }
        task->SubmitConflictDecision(FolderWindow::FileOperationState::Task::ConflictAction::KeepBoth, false);
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

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Literal-link move failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (! linkNames(dstBackwardLink, srcBackwardTarget) || ! linkNames(dstForwardLink, srcForwardTarget) ||
            ! linkNames(dstCycleA, srcCycleB) || ! linkNames(dstCycleB, srcCycleA))
        {
            Fail(L"Literal-link move did not keep every junction's stored target text (backward, forward, and cycle links must still name the source tree).");
            return true;
        }

        std::error_code ec;
        if (std::filesystem::exists(srcTree, ec) && ! ec)
        {
            Fail(L"Literal-link move left the source tree behind.");
            return true;
        }
        if (! FileSizeEquals(dstBackwardTarget / L"back.bin", 4 * 1024) ||
            ! FileSizeEquals(dstForwardTarget / L"forward.bin", 6 * 1024) ||
            ! FileSizeEquals(dstForwardCollision, 2 * 1024))
        {
            Fail(L"Literal-link move destination integrity check failed.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamMoveLiteralLinks",
                          L"backward=1 forwardKeepBoth=1 cycleLinks=2 rewritten=0",
                          1,
                          completed->second.conflictPromptCount,
                          0,
                          completed->second.hr);
        state.taskA.reset();
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (! RecreateEmptyDirectory(skipSrcRoot) || ! RecreateEmptyDirectory(skipDstTree) ||
            ! WriteTestFile(skipSrcTarget / L"payload.bin", 3 * 1024) ||
            ! WriteTestFile(skipDstTargetCollision, 1 * 1024) ||
            ! TryCreateJunction(skipSrcLink, skipSrcTarget))
        {
            Fail(L"Literal-link skipped-target test failed to seed its source/destination tree.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {skipSrcTree},
                                                 skipDstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Literal-link skipped-target test failed to start the move task.");
            return true;
        }
        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        auto* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            if (state.taskA.has_value() && state.completedTasks.contains(state.taskA.value()))
            {
                Fail(L"Literal-link skipped-target test completed without its target collision prompt.");
                return true;
            }
            return false;
        }
        if (! OrdinalString::EqualsNoCasePath(std::filesystem::path(prompt->destinationPath), skipDstTargetCollision) ||
            ! PromptHasAction(prompt.value(), FolderWindow::FileOperationState::Task::ConflictAction::Skip))
        {
            Fail(L"Literal-link skipped-target test did not expose the target collision.");
            return true;
        }
        task->SubmitConflictDecision(FolderWindow::FileOperationState::Task::ConflictAction::Skip, false);
        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        // The link publishes literally (it still names the source target) and its exact source is
        // cleaned; the skipped target and its parent stay in the source; nothing asks a second time.
        if (completed->second.hr != S_FALSE || completed->second.conflictPromptCount != 1u ||
            ! linkNames(skipDstLink, skipSrcTarget) ||
            GetFileAttributesW(skipSrcTarget.c_str()) == INVALID_FILE_ATTRIBUTES ||
            GetFileAttributesW(skipSrcLink.c_str()) != INVALID_FILE_ATTRIBUTES ||
            ! FileSizeEquals(skipDstTargetCollision, 1 * 1024))
        {
            Fail(std::format(L"Literal-link skipped-target move was not partial with a literal link and a retained source target: hr=0x{:08X} prompts={}.",
                             static_cast<unsigned long>(completed->second.hr),
                             completed->second.conflictPromptCount));
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamMoveLiteralLinkSkippedTarget",
                          L"target=skipped link=literal source-target=retained",
                          completed->second.conflictPromptCount,
                          1,
                          0,
                          completed->second.hr);
        state.taskA.reset();
        state.stepState = 6;
        return false;
    }

    if (state.stepState == 6)
    {
        if (! RecreateEmptyDirectory(copySrcRoot) || ! RecreateEmptyDirectory(copyDstTree) ||
            ! WriteTestFile(copySrcTarget / L"payload.bin", 5 * 1024) ||
            ! WriteTestFile(copyDstTargetCollision, 2 * 1024) ||
            ! TryCreateJunction(copySrcLink, copySrcTarget))
        {
            Fail(L"Literal-link Local Copy test failed to seed its source/destination tree.");
            return true;
        }
        if (! SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":2,"reparsePointPolicy":"preserve"})json"))
        {
            Fail(L"Literal-link Local Copy test failed to configure parallel Preserve.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {copySrcTree},
                                                 copyDstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Literal-link Local Copy test failed to start the task.");
            return true;
        }
        state.stepState = 7;
        return false;
    }

    if (state.stepState == 7)
    {
        auto* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            if (state.taskA.has_value() && state.completedTasks.contains(state.taskA.value()))
            {
                Fail(L"Literal-link Local Copy completed without the nested target Keep Both conflict.");
                return true;
            }
            return false;
        }
        if (! OrdinalString::EqualsNoCasePath(std::filesystem::path(prompt->destinationPath), copyDstTargetCollision) ||
            prompt->bucket != FolderWindow::FileOperationState::Task::ConflictBucket::TypeMismatch ||
            ! PromptHasAction(prompt.value(), FolderWindow::FileOperationState::Task::ConflictAction::KeepBoth) ||
            PromptHasAction(prompt.value(), FolderWindow::FileOperationState::Task::ConflictAction::Overwrite))
        {
            Fail(L"Literal-link Local Copy did not expose TypeMismatch/Keep Both for the nested target component.");
            return true;
        }
        task->SubmitConflictDecision(FolderWindow::FileOperationState::Task::ConflictAction::KeepBoth, false);
        state.stepState = 8;
        return false;
    }

    if (state.stepState == 8)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        if (FAILED(completed->second.hr) || ! linkNames(copyDstLink, copySrcTarget) ||
            ! FileSizeEquals(copyDstTarget / L"payload.bin", 5 * 1024) ||
            ! FileSizeEquals(copyDstTargetCollision, 2 * 1024) ||
            ! FileSizeEquals(copySrcTarget / L"payload.bin", 5 * 1024) ||
            GetFileAttributesW(copySrcLink.c_str()) == INVALID_FILE_ATTRIBUTES)
        {
            Fail(std::format(L"Literal-link Local Copy did not preserve the source and keep the link's stored text beside the Keep Both target: hr=0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamCopyLiteralLinks",
                          L"parallel=2 forwardKeepBoth=1 rewritten=0",
                          1,
                          completed->second.conflictPromptCount,
                          0,
                          completed->second.hr);
        state.taskA.reset();
        state.stepState = 9;
        return false;
    }

    if (state.stepState == 9)
    {
        // No link queue, no 4,096-entry link ceiling: a wide tree of junctions copies whole.
        if (! RecreateEmptyDirectory(wideSrcRoot) || ! RecreateEmptyDirectory(wideDstRoot) ||
            ! WriteTestFile(wideTarget / L"payload.bin", 1 * 1024))
        {
            Fail(L"Literal-link wide Copy test failed to seed its source tree.");
            return true;
        }
        for (unsigned int index = 0u; index < kWideLinkCount; ++index)
        {
            if (! TryCreateJunction(wideSrcTree / wideLinkName(index), wideTarget))
            {
                Fail(std::format(L"Literal-link wide Copy test failed to create junction {}.", index));
                return true;
            }
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {wideSrcTree},
                                                 wideDstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Literal-link wide Copy test failed to start the task.");
            return true;
        }
        state.stepState = 10;
        return false;
    }

    if (state.stepState == 10)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Literal-link wide Copy failed: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        unsigned int copiedLinks = 0u;
        bool sampledLiteral      = true;
        for (unsigned int index = 0u; index < kWideLinkCount; ++index)
        {
            const std::filesystem::path copiedLink = wideDstTree / wideLinkName(index);
            const DWORD attributes                 = GetFileAttributesW(copiedLink.c_str());
            if (attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
            {
                ++copiedLinks;
            }
            if (index % 700u == 0u && ! linkNames(copiedLink, wideTarget))
            {
                sampledLiteral = false;
            }
        }
        if (copiedLinks != kWideLinkCount || ! sampledLiteral || ! FileSizeEquals(wideDstTree / L"target" / L"payload.bin", 1 * 1024))
        {
            Fail(std::format(L"Literal-link wide Copy did not publish every junction literally: copied={} of {} sampledLiteral={}.",
                             copiedLinks,
                             kWideLinkCount,
                             sampledLiteral ? 1 : 0));
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamCopyWideLiteralLinks",
                          L"junctions=4200 linkQueue=none",
                          kWideLinkCount,
                          copiedLinks,
                          0,
                          completed->second.hr);
        state.taskA.reset();
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
        // R3-2: this case covers Copy-only Move into a destination without content proof.
        static_cast<void>(SetPluginConfiguration(state.infoDummy.get(), R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0","writerProof":false})json"));
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

        // Two items + PerItem mode drives the concurrent path. With no qualified bound-delete
        // strategy, the admitted Move must use CopyOnly and retain both source objects.
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

        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Cross-FS concurrent outbound Move expected CopyOnly S_FALSE, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        if (! FilesEqualBytes(refDir / L"m1.bin", localDir / L"m1.bin") || ! FilesEqualBytes(refDir / L"m2.bin", localDir / L"m2.bin"))
        {
            Fail(L"Cross-FS concurrent outbound CopyOnly Move did not preserve the local source files.");
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

        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Cross-FS concurrent return Move expected CopyOnly S_FALSE, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        if (! FilesEqualBytes(refDir / L"m1.bin", returnDir / L"m1.bin") || ! FilesEqualBytes(refDir / L"m2.bin", returnDir / L"m2.bin"))
        {
            Fail(L"Cross-FS concurrent move round-trip integrity check failed.");
            return true;
        }

        wil::com_ptr<IFileSystemIO> dummyIo;
        unsigned long dummyAttributes = 0u;
        if (! state.fsDummy || FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo ||
            FAILED(dummyIo->GetAttributes((dummyRoot + L"/m1.bin").c_str(), &dummyAttributes)) ||
            FAILED(dummyIo->GetAttributes((dummyRoot + L"/m2.bin").c_str(), &dummyAttributes)))
        {
            Fail(L"Cross-FS concurrent return CopyOnly Move removed a Dummy source object.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamCrossFsConcurrentMove",
                          L"shape=2-item-copy-only-roundtrip-source-kept",
                          2,
                          2,
                          2,
                          completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_RerunCopyIdenticalCollisionPrompts);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_RerunCopyIdenticalCollisionPrompts:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"Fairstream_RerunCopyIdenticalCollisionPrompts timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-rerun-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-rerun-dst";
    const std::filesystem::path srcFoo = srcRoot / L"Foo";
    const std::filesystem::path dstFoo = dstRoot / L"Foo";
    const std::filesystem::path srcFile = srcFoo / L"payload.bin";
    const std::filesystem::path dstFile = dstFoo / L"payload.bin";

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
        if (! WriteTestFile(srcFile, 12 * 1024))
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

        // Content equality is not a data-choice authority. An identical destination must use the
        // same typed collision surface as every other existing regular file.
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
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists ||
                    NormalizePathForCompare(prompt->sourcePath) != NormalizePathForCompare(srcFile.wstring()) ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip) ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Re-run identical copy did not expose the exact file through the typed Exists conflict surface.");
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.stepState = 3;
                return false;
            }
        }

        const auto completed = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (completed != state.completedTasks.end())
        {
            Fail(std::format(L"Re-run identical copy completed without a conflict decision (hr=0x{:08X}).",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        return false;
    }

    if (state.stepState == 3)
    {
        const auto completed = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Re-run identical copy should report the explicit Skip: 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 1u || ! FilesEqualBytes(srcFile, dstFile))
        {
            Fail(std::format(L"Re-run identical copy expected one prompt and unchanged bytes, saw {} prompt(s).",
                             completed->second.conflictPromptCount));
            return true;
        }

        if (! WriteTestFile(srcFile, 20 * 1024))
        {
            Fail(L"Re-run collision test failed to modify the source file.");
            return true;
        }

        state.taskA = startCopyTask();
        if (! state.taskA.has_value())
        {
            Fail(L"Re-run resume test failed to start the changed-child copy.");
            return true;
        }

        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                const std::wstring expected = NormalizePathForCompare(srcFile.wstring());
                if (NormalizePathForCompare(prompt->sourcePath) != expected)
                {
                    Fail(std::format(L"Re-run resume test prompted for the wrong child: '{}'.", prompt->sourcePath));
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
                state.stepState = 5;
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed != state.completedTasks.end())
        {
            Fail(std::format(L"Changed re-run copy completed without a conflict decision (hr=0x{:08X}).",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        return false;
    }

    if (state.stepState == 5)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Changed re-run copy failed after explicit Overwrite: 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.conflictPromptCount != 1u)
        {
            Fail(std::format(L"Re-run resume changed-child copy expected exactly 1 prompt, saw {}.", completed->second.conflictPromptCount));
            return true;
        }
        if (! FilesEqualBytes(srcFile, dstFile))
        {
            Fail(L"Changed re-run copy did not refresh the destination after explicit Overwrite.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamRerunCollision",
                          L"identical=prompt+skip;changed=prompt+overwrite;hash-authority=none",
                          2u,
                          2u,
                          0u,
                          completed->second.hr);
        NextStep(state, SelfTestState::Step::Fairstream_MoveSameSizeCollisionPrompts);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_MoveSameSizeCollisionPrompts:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Fairstream_MoveSameSizeCollisionPrompts timed out.");
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
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
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

        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Same-size collision move expected S_FALSE after Skip, got 0x{:08X}.",
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
#endif
#if defined(FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_B)

case SelfTestState::Step::Fairstream_GraphBandsFairUnderParallelCopy:
{
    FileOperationsPopupInternal::PopupLayoutDebugSnapshot graph{};
    if (! DebugBuildFileOperationsGraphFairnessHistorySnapshot(graph))
    {
        Fail(std::format(L"Graph fairness deterministic sampler failed (multi={}, single={}, distinct={}, min={:.3f}, max={:.3f}, accumCalls={}, pending={}, "
                         L"maxStreams={}, rowColorMatches={}, rowColorMismatches={}).",
                         graph.graphMultiHueBucketCount,
                         graph.graphSingleHueBucketCount,
                         graph.graphDistinctHueCount,
                         graph.graphMinHueShare,
                         graph.graphMaxHueShare,
                         graph.graphDebugAccumulateCalls,
                         graph.graphDebugLastPending,
                         graph.graphDebugMaxStreams,
                         graph.graphRowColorMatchCount,
                         graph.graphRowColorMismatchCount));
        return true;
    }

    if (graph.graphDebugMaxStreams != 4u)
    {
        Fail(std::format(L"Graph fairness deterministic sampler expected 4 concurrent streams, saw {}.", graph.graphDebugMaxStreams));
        return true;
    }

    if (graph.graphRowColorMatchCount != 4u || graph.graphRowColorMismatchCount != 0u)
    {
        Fail(std::format(L"Graph fairness deterministic sampler expected exact row/graph color parity for 4 streams, matched {} and mismatched {}.",
                         graph.graphRowColorMatchCount,
                         graph.graphRowColorMismatchCount));
        return true;
    }

    if (! graph.graphCurrentBandwidthLineVisible || graph.graphCurrentBandwidthBytesPerSecond <= 0.0)
    {
        Fail(L"Graph fairness deterministic sampler did not publish the current effective-bandwidth marker.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamProgressGraphColorParity",
                      L"shape=4-identified-streams-dark-rainbow",
                      graph.graphRowColorMatchCount,
                      graph.graphRowColorMismatchCount,
                      graph.graphDistinctHueCount,
                      S_OK);

    Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamGraphBandFairness",
                      L"shape=4-equal-streams-deterministic-history",
                      graph.graphMultiHueBucketCount,
                      static_cast<uint64_t>(graph.graphMinHueShare * 1000.0),
                      static_cast<uint64_t>(graph.graphMaxHueShare * 1000.0),
                      S_OK);
    NextStep(state, SelfTestState::Step::Fairstream_ManagedMergeMovePreservesChildren);
    return false;
}
case SelfTestState::Step::Fairstream_ManagedMergeMovePreservesChildren:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Fairstream_ManagedMergeMovePreservesChildren timed out.");
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

        // The prepared strategy is checked once the task has run.

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
        if (DebugGetPreparedTransferStrategyForSelfTest(state.taskA.value()) != FileOperations::OperationStrategy::Native)
        {
            Fail(L"An existing same-volume regular-directory merge is a Native plan that continues as a rename merge.");
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

        constexpr std::array<uint64_t, 3> kExpectedSizes{{256 * 1024, 64 * 1024, 128 * 1024}};
        for (size_t index = 0; index < destinationFiles.size(); ++index)
        {
            if (! FileSizeEquals(destinationFiles[index], kExpectedSizes[index]))
            {
                Fail(std::format(L"Managed merge-move test: destination child {} is missing or has the wrong size.",
                                 destinationFiles[index].wstring()));
                return true;
            }
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamManagedMergeMove",
                          L"shape=same-volume-managed-merge-3-files;source=removed;destination-sibling=preserved",
                          3,
                          1,
                          0,
                          S_OK);
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
    const std::filesystem::path sourceCollision = srcPayload / L"z-collision.bin";
    const std::filesystem::path destinationCollision = dstPayload / L"z-collision.bin";
    constexpr uint64_t kSourceCollisionBytes = 2u * 1024u;
    constexpr uint64_t kDestinationCollisionBytes = 5u * 1024u;

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcPayload / L"sub") || ! RecreateEmptyDirectory(dstPayload) || ! RecreateEmptyDirectory(linkTarget))
        {
            Fail(L"Junction-merge test failed to prepare directories.");
            return true;
        }
        if (! WriteTestFile(srcPayload / L"sub" / L"inner.bin", 4 * 1024) || ! WriteTestFile(linkTarget / L"sentinel.bin", 1024) ||
            ! WriteTestFile(sourceCollision, kSourceCollisionBytes) ||
            ! WriteTestFile(destinationCollision, kDestinationCollisionBytes))
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

    if (state.stepState >= 1 && state.stepState <= 3)
    {
        Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
        {
            if (prompt->bucket == Task::ConflictBucket::DestinationLink)
            {
                if (state.stepState == 2 ||
                    ! OrdinalString::EqualsNoCasePath(std::filesystem::path(prompt->destinationPath), dstPayload / L"sub") ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::ReplaceLink) ||
                    PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Junction-merge test did not expose exactly one typed Replace Link decision for the destination junction.");
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::ReplaceLink, false);
                state.stepState = state.stepState == 3 ? 4 : 2;
                return false;
            }
            if (prompt->bucket == Task::ConflictBucket::RegularFileExists)
            {
                if (state.stepState == 3 ||
                    ! OrdinalString::EqualsNoCasePath(std::filesystem::path(prompt->destinationPath), destinationCollision) ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
                {
                    Fail(L"Junction-merge test did not expose exactly one independent child-file collision after the link decision.");
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.stepState = state.stepState == 2 ? 4 : 3;
                return false;
            }

            Fail(std::format(L"Junction-merge test exposed unexpected conflict bucket {}.", static_cast<int>(prompt->bucket)));
            return true;
        }

        if (state.completedTasks.contains(state.taskA.value()))
        {
            Fail(L"Junction-merge test completed before both the link receipt and the colliding child received independent decisions.");
            return true;
        }
        return false;
    }

    if (state.stepState == 4)
    {
        const auto completed = state.completedTasks.find(state.taskA.value());
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT expectedHr = S_FALSE;
        if (completed->second.hr != expectedHr || completed->second.conflictPromptCount != 2u ||
            completed->second.conflictExpectedDestinationBoundCount < 1u ||
            completed->second.conflictExpectedDestinationReturnedCount != 1u ||
            completed->second.conflictExpectedDestinationUnavailableCount != 0u)
        {
            Fail(std::format(L"Junction-merge test expected two independent prompts, one exact replacement receipt, and 0x{:08X}; prompts={} bound={} returned={} unavailable={} hr=0x{:08X}.",
                             static_cast<unsigned long>(expectedHr),
                             completed->second.conflictPromptCount,
                             completed->second.conflictExpectedDestinationBoundCount,
                             completed->second.conflictExpectedDestinationReturnedCount,
                             completed->second.conflictExpectedDestinationUnavailableCount,
                             static_cast<unsigned long>(completed->second.hr)));
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

        if (! FileSizeEquals(destinationCollision, kDestinationCollisionBytes) || ! FileSizeEquals(sourceCollision, kSourceCollisionBytes))
        {
            Fail(L"Junction-merge test: Replace Link leaked into the colliding child or changed its source.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamJunctionMergeConsent",
                          L"typed-replace-link;child-collision-independent;target-untouched",
                          completed->second.conflictPromptCount,
                          1u,
                          0u,
                          completed->second.hr);
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

    // dstAlias is a junction pointing INSIDE the source tree. Admission must remain streaming;
    // the no-follow identity/ancestor guard rejects it when this item reaches execution.
    const std::filesystem::path root = state.tempRoot / L"fairstream-selfcopy";
    const std::filesystem::path srcTree = root / L"tree";
    const std::filesystem::path inside = srcTree / L"inside";
    const std::filesystem::path dstAlias = root / L"alias";

    if (state.stepState == 0)
    {
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
        state.taskA = StartFileOperationAndGetId(state.fileOps,
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
        if (! state.taskA.has_value())
        {
            Fail(L"Copy-into-self alias test: streaming admission rejected the task before just-in-time identity discovery.");
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
    if (SUCCEEDED(completed->second.hr))
    {
        Fail(L"Copy-into-self alias test: the just-in-time no-follow guard allowed an aliased destination inside the source.");
        return true;
    }
    std::error_code ec;
    if (std::filesystem::exists(inside / L"tree", ec) || ec)
    {
        Fail(L"Copy-into-self alias test: execution published a child through the rejected junction alias.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FairstreamCopyIntoSelfAlias",
                      L"shape=junction-alias-into-source;jit=1",
                      1,
                      0,
                      0,
                      completed->second.hr);
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
            R"json({"concurrencyMode":"auto","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));
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
        if (completed->second.sourceItemResults.size() != 1u || ! completed->second.sourceItemResults.front().has_value() ||
            completed->second.sourceItemResults.front()->completion != FileOperations::ItemCompletion::Failed ||
            completed->second.sourceItemResults.front()->sourceDisposition != FileOperations::SourceDisposition::Retained ||
            completed->second.sourceItemResults.front()->status != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            Fail(L"Parallel-delete continue test did not preserve known retained-source mutation truth for the partial root.");
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
case SelfTestState::Step::Cinderstar_BridgeMovePerFileConflictSkipsParallel:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    const bool parallelScenario = state.step == SelfTestState::Step::Cinderstar_BridgeMovePerFileConflictSkipsParallel;
    const std::wstring_view scenarioLabel = parallelScenario ? L"configured-parallel" : L"serial";
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(std::format(L"Cross-FS MOVE per-file conflict {} scenario timed out.", scenarioLabel));
        return true;
    }

    // Cross-filesystem MOVE (dummy -> local) of a DIRECTORY whose child collides with a
    // pre-existing local file. The host only sees the top-level directory item (no collision),
    // so the CHILD conflict is the bridge's job: it must raise a per-file conflict, honor Skip
    // without touching either colliding copy, move/delete the non-colliding sibling, and end PARTIAL.
    const std::wstring dummyDir = parallelScenario ? L"/csm-p" : L"/csm-s";
    const std::wstring dummyCollide = dummyDir + L"/collide.txt";
    const std::wstring dummyKeep = dummyDir + L"/keep.txt";
    const std::filesystem::path localDest = state.tempRoot / (parallelScenario ? L"csm-p-d" : L"csm-s-d");
    const std::filesystem::path destSubdir = localDest / (parallelScenario ? L"csm-p" : L"csm-s");
    const std::filesystem::path destCollide = destSubdir / L"collide.txt";
    const std::filesystem::path destKeep = destSubdir / L"keep.txt";
    constexpr size_t kSentinelBytes = 7u;
    constexpr std::string_view kCollidingSourceBytes = "dummy-collide-content";
    constexpr std::string_view kMovedSourceBytes = "dummy-keep-content";

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyDir) || ! DummyWriteTextFile(state.fsDummy.get(), dummyCollide, kCollidingSourceBytes) ||
            ! DummyWriteTextFile(state.fsDummy.get(), dummyKeep, kMovedSourceBytes))
        {
            Fail(std::format(L"Cross-FS MOVE per-file conflict {} scenario failed to seed the dummy source tree.", scenarioLabel));
            return true;
        }
        // Pre-create only the colliding child under the merge destination; keep.txt must NOT
        // exist yet (it is the proof a sibling still copies past the skipped child).
        if (! RecreateEmptyDirectory(destSubdir) || ! WriteTestFile(destCollide, kSentinelBytes))
        {
            Fail(std::format(L"Cross-FS MOVE per-file conflict {} scenario failed to seed local child '{}'.",
                             scenarioLabel,
                             destCollide.native()));
            return true;
        }

        // Exercise both scheduling implementations. The admitted request is CopyOnly, so neither
        // the serial nor parallel executor may remove a source entry.
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            parallelScenario ? R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4})json"
                             : R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
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
            Fail(std::format(L"Cross-FS MOVE per-file conflict {} scenario failed to start.", scenarioLabel));
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
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists)
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

        if (completed->second.conflictPromptCount != 1u)
        {
            Fail(std::format(L"Cross-FS MOVE per-file conflict {} scenario expected exactly one prompt, got {}.",
                             scenarioLabel,
                             completed->second.conflictPromptCount));
            return true;
        }
        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Bridge per-file conflict test expected S_FALSE after Skip, got 0x{:08X}.",
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

        wil::com_ptr<IFileSystemIO> dummyIo;
        wil::com_ptr<IFileSystemIO> localIo;
        if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo ||
            FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
        {
            Fail(L"Cross-FS MOVE per-file conflict test could not acquire source/destination IO.");
            return true;
        }

        std::string actual;
        if (! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyCollide), actual) || actual != kCollidingSourceBytes)
        {
            Fail(L"Cross-FS MOVE per-file conflict test did not retain the skipped source bytes.");
            return true;
        }
        actual.clear();
        if (! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyKeep), actual) || actual != kMovedSourceBytes)
        {
            Fail(L"Cross-FS CopyOnly per-file conflict test did not retain the successfully copied sibling source.");
            return true;
        }
        actual.clear();
        if (! ReadFileTextFsIo(localIo, destKeep, actual) || actual != kMovedSourceBytes)
        {
            Fail(L"Cross-FS MOVE per-file conflict test produced incorrect destination sibling bytes.");
            return true;
        }

        const unsigned int expectedConcurrency = parallelScenario ? 4u : 1u;
        if (completed->second.configuredMaxConcurrency != expectedConcurrency)
        {
            Fail(std::format(L"Cross-FS MOVE per-file conflict {} scenario used concurrency {}, expected {}.",
                             scenarioLabel,
                             completed->second.configuredMaxConcurrency,
                             expectedConcurrency));
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.CinderstarBridgeMovePerFileConflictSkip",
                          scenarioLabel,
                          completed->second.conflictPromptCount,
                          completed->second.configuredMaxConcurrency,
                          0,
                          completed->second.hr);
        NextStep(state,
                 parallelScenario ? SelfTestState::Step::Riptide_BridgeNestedDirVsFileSkipContinuesSiblings
                                  : SelfTestState::Step::Cinderstar_BridgeMovePerFileConflictSkipsParallel);
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
                if (prompt->bucket != Task::ConflictBucket::TypeMismatch ||
                    PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(std::format(L"Riptide bridge dir-vs-file test expected TypeMismatch without Overwrite, got bucket={} status=0x{:08X} destination='{}'.",
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

        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Riptide bridge dir-vs-file test expected S_FALSE after Skip, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (completed->second.sourceItemResults.size() != 1u || ! completed->second.sourceItemResults.front().has_value() ||
            completed->second.sourceItemResults.front()->publication != FileOperations::PublicationState::Published ||
            completed->second.sourceItemResults.front()->sourceDisposition != FileOperations::SourceDisposition::Retained ||
            completed->second.sourceItemResults.front()->completion != FileOperations::ItemCompletion::Skipped ||
            completed->second.sourceItemResults.front()->status != S_FALSE)
        {
            Fail(L"Riptide bridge dir-vs-file Skip did not preserve the top-level Published/Retained/Skipped axes.");
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
        NextStep(state, SelfTestState::Step::Riptide_BridgeDirectoryOverReadOnlyFileIsTypeMismatch);
        return false;
    }

    return false;
}
case SelfTestState::Step::Riptide_BridgeDirectoryOverReadOnlyFileIsTypeMismatch:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    const std::wstring dummyRoot = L"/riptide-bridge-readonly-dir-src";
    const std::wstring dummyChild = dummyRoot + L"/child.txt";
    const std::filesystem::path localDest = state.tempRoot / L"riptide-bridge-readonly-dir-dst";
    const std::filesystem::path readOnlyPath = localDest / L"riptide-bridge-readonly-dir-src";
    const std::filesystem::path keepBothPath = localDest / L"riptide-bridge-readonly-dir-src (2)";
    const std::filesystem::path copiedChild = keepBothPath / L"child.txt";
    constexpr std::string_view kChildBytes = "readonly-child";

    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
        Fail(L"Riptide_BridgeDirectoryOverReadOnlyFileIsTypeMismatch timed out.");
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
                if (prompt->bucket != Task::ConflictBucket::TypeMismatch ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::KeepBoth) ||
                    PromptHasAction(prompt.value(), Task::ConflictAction::ReplaceReadOnly) ||
                    PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
                    Fail(
                        std::format(L"Riptide bridge read-only directory test expected TypeMismatch/Keep Both without replacement actions, got bucket={} status=0x{:08X} destination='{}'.",
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

                task->SubmitConflictDecision(Task::ConflictAction::KeepBoth, false);
                state.markerTick = nowTick;
                state.stepState = 2;
                return false;
            }
        }

        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed != state.completedTasks.end())
        {
            static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
            Fail(std::format(L"Riptide bridge read-only directory test completed before TypeMismatch/Keep Both prompt: 0x{:08X}.",
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
                    Fail(std::format(L"Riptide bridge read-only directory test still had an active prompt after Keep Both: bucket={} status=0x{:08X} "
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
                    L"Riptide bridge read-only directory test raised a second prompt after Keep Both: bucket={} status=0x{:08X} destination='{}'.",
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
            Fail(std::format(L"Riptide bridge read-only directory test expected success after Keep Both, got 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        std::error_code ec;
        if (! std::filesystem::is_regular_file(readOnlyPath, ec) || ec)
        {
            static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
            Fail(L"Riptide bridge read-only directory test replaced the blocking file instead of preserving it.");
            return true;
        }
        if (! FileSizeEquals(copiedChild, kChildBytes.size()))
        {
            static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
            Fail(L"Riptide bridge read-only directory test did not copy the source child into the Keep Both sibling.");
            return true;
        }

        static_cast<void>(SetFileAttributesW(readOnlyPath.c_str(), FILE_ATTRIBUTE_NORMAL));
        Debug::Perf::Emit(L"FileOps.SelfTest.RiptideBridgeReadOnlyDirKeepBoth", L"", completed->second.conflictPromptCount, 0, 0, completed->second.hr);
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
                if (prompt->bucket != Task::ConflictBucket::TypeMismatch ||
                    PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(std::format(L"Riptide bridge create-directory race test expected TypeMismatch without Overwrite, got bucket={} status=0x{:08X} destination='{}'.",
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

        if (completed->second.hr != S_FALSE)
        {
            Fail(std::format(L"Riptide bridge create-directory race test expected S_FALSE after Skip, got 0x{:08X}.",
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
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));

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
        NextStep(state, SelfTestState::Step::Fairstream_DiscoveryAheadOverlapsTransfer);
        return false;
    }

    return false;
}
case SelfTestState::Step::Fairstream_DiscoveryAheadOverlapsTransfer:
{
    // Discovery and mutation share one traversal: bytes start moving before recursive discovery
    // closes. The deterministic proof is the
    // engine latching _firstMutationBeforeDiscoveryClosed when a transfer progress callback fires
    // while discovery remains open — impossible in the retired scan-then-execute model. The
    // operation must still complete correctly with authoritative totals at traversal closure.
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"Fairstream_DiscoveryAheadOverlapsTransfer timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"fairstream-earlyadmit-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"fairstream-earlyadmit-dst";
    // A wide tree makes producer/consumer overlap deterministic while modest files keep the run quick.
    constexpr unsigned int kDirs = 16u;
    constexpr unsigned int kFilesPerDir = 16u;
    constexpr size_t kFileBytes = 8u * 1024u;

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

        const std::string localConfig =
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json";
        if (! SetPluginConfiguration(state.infoLocal.get(), localConfig))
        {
            Fail(L"Discovery-overlap test failed to apply local concurrency configuration.");
            return true;
        }

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

    // The first mutation must overlap the still-open producer traversal.
    if (! completed->second.firstMutationBeforeDiscoveryClosed)
    {
        Fail(std::format(L"Discovery-ahead did not overlap transfer progress "
                         L"(discoveryUs={} completedBytes={} expectedBytes={} configuredConcurrency={}).",
                         completed->second.discoveryDurationUs,
                         completed->second.progressCompletedBytes,
                         static_cast<uint64_t>(kDirs) * kFilesPerDir * kFileBytes,
                         completed->second.configuredMaxConcurrency));
        return true;
    }

    // Traversal closure makes the discovered totals authoritative.
    if (! completed->second.discoveryClosed)
    {
        Fail(L"Discovery-overlap run finished before its traversal closed.");
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
    if (completed->second.progressCompletedBytes < expectedBytes || completed->second.discoveredTotalBytes != expectedBytes ||
        completed->second.discoveredFiles != static_cast<unsigned long>(kDirs * kFilesPerDir))
    {
        Fail(std::format(
            L"Discovery totals did not reconcile: completedBytes={} discoveredBytes={} discoveredFiles={} expectedBytes={} expectedFiles={}.",
            completed->second.progressCompletedBytes,
            completed->second.discoveredTotalBytes,
            completed->second.discoveredFiles,
            expectedBytes,
            kDirs * kFilesPerDir));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.DiscoveryAheadOverlap",
                      L"",
                      completed->second.discoveryDurationUs,
                      completed->second.discoveredTotalBytes,
                      expectedBytes,
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
    // A task card is presented only when it is still running at the 500-ms reveal deadline; a
    // finished task is never revealed before its completion summary. Throttle a 4 MiB copy to
    // about two seconds so the row exists before the post-finished pause is observed.
    constexpr size_t kFileBytes                    = 4u * 1024u * 1024u;
    constexpr uint64_t kRevealSpeedLimitBytesPerSecond = 2ull * 1024ull * 1024ull;

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
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFile},
                                                 dstRoot,
                                                 flags,
                                                 false,
                                                 kRevealSpeedLimitBytesPerSecond);
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
#endif
#if defined(FILEOPS_SELFTEST_INCLUDE_FAIRSTREAM_C)

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
    constexpr std::array<std::wstring_view, 5> kScenarioNames{{
        L"read-overreport", L"premature-eof", L"write-underconsume", L"write-overreport", L"null-source-reader"}};
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        SetFileOpsBridgeOverReportNextReadForSelfTest(0);
        SetFileOpsBridgePrematureEofNextReadForSelfTest(0);
        SetFileOpsBridgeUnderConsumeNextWriteForSelfTest(0);
        SetFileOpsBridgeOverReportNextWriteForSelfTest(0);
        SetFileOpsBridgeNullNextSourceReaderForSelfTest(0);
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
        SetFileOpsBridgeNullNextSourceReaderForSelfTest(0);
    };
    const auto takeAttempts = [scenario]() noexcept -> unsigned long
    {
        switch (scenario)
        {
            case 0: return TakeFileOpsBridgeOverReportNextReadAttemptsForSelfTest();
            case 1: return TakeFileOpsBridgePrematureEofNextReadAttemptsForSelfTest();
            case 2: return TakeFileOpsBridgeUnderConsumeNextWriteAttemptsForSelfTest();
            case 3: return TakeFileOpsBridgeOverReportNextWriteAttemptsForSelfTest();
            case 4: return TakeFileOpsBridgeNullNextSourceReaderAttemptsForSelfTest();
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
            case 4: SetFileOpsBridgeNullNextSourceReaderForSelfTest(1); break;
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
        Common::Settings::FileOperationsSettings& fileOperations = EnsureFileOperationsSettingsForSelfTest();
        fileOperations.crossFsBridgeBufferSizeKB                 = 2048u;
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":16,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));
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
    const HRESULT expectedCompletion = scenario == 0u ? S_OK : S_FALSE;
    if (completed->second.hr != expectedCompletion)
    {
        Fail(std::format(L"Causeway scheduling/resource {} expected 0x{:08X}, got 0x{:08X}.",
                         scenario == 0u ? L"COPY" : L"MOVE",
                         static_cast<unsigned long>(expectedCompletion),
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }
    if (! completed->second.discoveryClosed)
    {
        Fail(L"Causeway cross-filesystem operation did not close its single discovery traversal.");
        return true;
    }

    // Cross-provider Move is admitted as CopyOnly in P1. It performs the proven Copy traversal
    // exactly once and never starts a second source-cleanup traversal.
    const uint64_t expectedEnumerations = kDirectoryCount;
    if (completed->second.bridgeSourceDirectoryEnumerationCount != expectedEnumerations)
    {
        Fail(std::format(L"Causeway {} expected {} source directory listings, got {}.",
                         scenario == 0u ? L"COPY" : L"MOVE",
                         expectedEnumerations,
                         completed->second.bridgeSourceDirectoryEnumerationCount));
        return true;
    }

    if (scenario == 1u)
    {
        wil::com_ptr<IFileSystemIO> sourceIo;
        const HRESULT ioHr = state.fsDummy->QueryInterface(IID_PPV_ARGS(sourceIo.addressof()));
        unsigned long sourceAttributes = 0u;
        const HRESULT sourceHr = sourceIo ? sourceIo->GetAttributes(dummyRoot.c_str(), &sourceAttributes) : E_NOINTERFACE;
        if (FAILED(ioHr) || ! sourceIo || FAILED(sourceHr) || (sourceAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
        {
            Fail(std::format(L"Causeway CopyOnly MOVE did not retain its source directory (QI=0x{:08X}, source=0x{:08X}).",
                             static_cast<unsigned long>(ioHr),
                             static_cast<unsigned long>(sourceHr)));
            return true;
        }
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
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));
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
                                                                        : FileOpsBridgeReparsePolicyOverride::Preserve);
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
                Fail(L"Causeway file Preserve failure did not use the UnsupportedReparse conflict bucket.");
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
    const HRESULT expectedAggregate = scenario == 0u ? S_FALSE : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    if (injectionAttempts != 1u || completed->second.hr != expectedAggregate)
    {
        Fail(std::format(L"Causeway file-reparse {} expected one injection and aggregate 0x{:08X}; attempts={} hr=0x{:08X}.",
                         scenario == 0u ? L"Skip" : L"Preserve",
                         static_cast<unsigned long>(expectedAggregate),
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
            Fail(L"Causeway file Preserve scenario did not retain the ERROR_NOT_SUPPORTED diagnostic.");
            return true;
        }
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.CausewayBridgeFileReparsePolicy",
                      scenario == 0u ? L"Skip" : L"Preserve",
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
    NextStep(state, SelfTestState::Step::Beeline_CopySkipLinksKeepsPlaceholders);
    return false;
}
case SelfTestState::Step::Beeline_CopySkipLinksKeepsPlaceholders:
{
    // Beeline: Copy `Skip links` skips name-surrogate links only. A fully hydrated Cloud Files placeholder carries the
    // reparse attribute but binds as an ordinary file object, so the cross-provider bridge copies it
    // and skips the junction beside it.
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"preserve"})json"));
        Fail(L"Beeline_CopySkipLinksKeepsPlaceholders timed out.");
        return true;
    }

    const auto cleanupAlternate = [&]() noexcept
    {
        state.placeholderSyncRoot.reset();
        if (! state.fileOpsAlternateVolumeRoot.empty())
        {
            const auto root = std::move(state.fileOpsAlternateVolumeRoot);
            state.fileOpsAlternateVolumeRoot.clear();
            static_cast<void>(SelfTest::RemoveAll(root));
            PruneEmptyAlternateVolumeSandboxParents(root);
        }
    };
    const auto skipCase = [&](std::wstring_view reason) noexcept
    {
        Debug::Perf::Emit(L"FileOps.SelfTest.CopySkipLinksKeepsPlaceholders", std::format(L"skip={}", reason), 0u, 0u, 0u, S_FALSE);
        RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::skipped, reason);
        cleanupAlternate();
        NextStep(state, SelfTestState::Step::Causeway_BridgeFailureStatusAndPausedReader);
        return false;
    };
    if (state.stepState == 0u)
    {
        std::array<wchar_t, 32> fileSystemName{};
        if (! GetVolumeInformationW(state.tempRoot.root_path().c_str(), nullptr, 0u, nullptr, nullptr, nullptr,
                                   fileSystemName.data(), static_cast<DWORD>(fileSystemName.size())))
        {
            return skipCase(L"primary-filesystem-unavailable");
        }
        if (std::wstring_view(fileSystemName.data()) != L"NTFS")
        {
            std::wstring reason;
            const auto alternate = TryCreateAlternateWritableVolumeSelfTestRoot(state.tempRoot, reason);
            if (! alternate.has_value())
            {
                return skipCase(reason);
            }
            state.fileOpsAlternateVolumeRoot = alternate.value();
            if (! GetVolumeInformationW(alternate.value().root_path().c_str(), nullptr, 0u, nullptr, nullptr, nullptr,
                                       fileSystemName.data(), static_cast<DWORD>(fileSystemName.size())) ||
                std::wstring_view(fileSystemName.data()) != L"NTFS")
            {
                return skipCase(L"authorized-alternate-is-not-ntfs");
            }
        }
    }
    const std::filesystem::path srcRoot =
        (state.fileOpsAlternateVolumeRoot.empty() ? state.tempRoot : state.fileOpsAlternateVolumeRoot) / L"beeline-skip-links-src";
    const std::filesystem::path srcTree    = srcRoot / L"tree";
    const std::filesystem::path placeholder = srcTree / L"placeholder.bin";
    const std::filesystem::path junction   = srcTree / L"link";
    const std::filesystem::path linkTarget = srcRoot / L"target";
    const std::wstring dummyRoot           = L"/beeline-skip-links";
    const std::wstring dummyTree           = dummyRoot + L"/tree";

    if (state.stepState == 0u)
    {
        const FileSystemFlags cleanupFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        static_cast<void>(state.fsDummy->DeleteItem(dummyRoot.c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! RecreateEmptyDirectory(srcTree) || ! RecreateEmptyDirectory(linkTarget) || ! WriteFilledTestFile(placeholder, 64u * 1024u, 0x5A) ||
            ! TryCreateJunction(junction, linkTarget) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
        {
            Fail(L"Beeline Copy Skip could not stage its source tree.");
            return true;
        }

        // WOF compression can be transparent to reparse-attribute queries. Use a real, fully
        // hydrated Cloud Files placeholder instead, with no account, callbacks or network I/O.
        // Windows can disguise placeholders for legacy applications. Explicitly expose them
        // in this isolated selftest process (including worker threads) and the UI thread;
        // the WIL fixture owner restores both previous modes on every exit.
        auto& fixture = state.placeholderSyncRootStorage;
        const HMODULE ntdll = GetModuleHandleW(L"ntdll.dll");
        const FARPROC setProcess = ntdll ? GetProcAddress(ntdll, "RtlSetProcessPlaceholderCompatibilityMode") : nullptr;
        const FARPROC setThread = ntdll ? GetProcAddress(ntdll, "RtlSetThreadPlaceholderCompatibilityMode") : nullptr;
        if (! setProcess || ! setThread)
        {
            return skipCase(L"placeholder-compatibility-api-unavailable");
        }
#pragma warning(push)
#pragma warning(disable : 4191) // Exact documented NTAPI ABI; ntdll remains loaded for process lifetime.
        fixture.setProcessMode = reinterpret_cast<PlaceholderSyncRoot::SetCompatibilityModeFn>(setProcess);
        fixture.setThreadMode = reinterpret_cast<PlaceholderSyncRoot::SetCompatibilityModeFn>(setThread);
#pragma warning(pop)
        state.placeholderSyncRoot.reset(&fixture);
        constexpr CHAR exposePlaceholders = 2; // PHCM_EXPOSE_PLACEHOLDERS (ntifs.h).
        fixture.previousProcessMode = fixture.setProcessMode(exposePlaceholders);
        fixture.previousThreadMode = fixture.setThreadMode(exposePlaceholders);
        if (fixture.previousProcessMode < 0 || fixture.previousThreadMode < 0)
        {
            Fail(L"Placeholder fixture could not expose native reparse metadata.");
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.PlaceholderVisibility", L"mode=exposed;scope=fixture;restore=required", 0u,
                          static_cast<uint64_t>(fixture.previousProcessMode), static_cast<uint64_t>(fixture.previousThreadMode), S_OK);
        CF_SYNC_REGISTRATION registration{};
        registration.StructSize      = sizeof(registration);
        registration.ProviderName    = L"RedSalamander FileOps selftest";
        registration.ProviderVersion = L"1";
        registration.ProviderId      = {0x3e5b50c5, 0xe92f, 0x49d4, {0xa2, 0x1d, 0x8d, 0x5e, 0x5e, 0x34, 0xbd, 0x71}};
        CF_SYNC_POLICIES policies{};
        policies.StructSize         = sizeof(policies);
        policies.Hydration.Primary  = CF_HYDRATION_POLICY_ALWAYS_FULL;
        policies.Population.Primary = CF_POPULATION_POLICY_ALWAYS_FULL;
        state.placeholderSyncRootStorage.path      = srcRoot;
        state.placeholderSyncRootStorage.cleanupHr = S_OK;
        const HRESULT registerHr = CfRegisterSyncRoot(srcRoot.c_str(), &registration, &policies, CF_REGISTER_FLAG_NONE);
        if (FAILED(registerHr))
        {
            return skipCase(std::format(L"cloud-placeholder-registration-unavailable:0x{:08X}", static_cast<unsigned long>(registerHr)));
        }
        fixture.registered = true;
        {
            const wil::unique_hfile handle(CreateFileW(
                placeholder.c_str(), GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
            constexpr std::array<unsigned char, 4> identity{{1u, 2u, 3u, 4u}};
            const HRESULT convertHr = handle ? CfConvertToPlaceholder(handle.get(), identity.data(), static_cast<DWORD>(identity.size()),
                                                                      CF_CONVERT_FLAG_MARK_IN_SYNC, nullptr, nullptr)
                                             : HRESULT_FROM_WIN32(GetLastError());
            if (FAILED(convertHr))
            {
                Fail(std::format(L"Placeholder conversion failed after successful registration (hr=0x{:08X}).", static_cast<unsigned long>(convertHr)));
                return true;
            }
        }
        const DWORD placeholderAttributes = GetFileAttributesW(placeholder.c_str());
        const wil::unique_hfile tagHandle(CreateFileW(placeholder.c_str(), FILE_READ_ATTRIBUTES,
                                                     FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING,
                                                     FILE_FLAG_OPEN_REPARSE_POINT, nullptr));
        FILE_ATTRIBUTE_TAG_INFO tagInfo{};
        const bool tagRead = tagHandle && GetFileInformationByHandleEx(tagHandle.get(), FileAttributeTagInfo, &tagInfo, sizeof(tagInfo));
        const DWORD tagError = tagRead ? ERROR_SUCCESS : GetLastError();
        if (placeholderAttributes == INVALID_FILE_ATTRIBUTES || (placeholderAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0u || ! tagRead ||
            (tagInfo.ReparseTag & ~static_cast<DWORD>(IO_REPARSE_TAG_CLOUD_MASK)) != IO_REPARSE_TAG_CLOUD ||
            IsReparseTagNameSurrogate(tagInfo.ReparseTag))
        {
            Fail(std::format(L"Placeholder fixture must expose a Cloud Files reparse tag that is not a name surrogate "
                             L"(attributes=0x{:08X}, tagRead={}, tagAttributes=0x{:08X}, tag=0x{:08X}, error={}).",
                             placeholderAttributes, tagRead, tagInfo.FileAttributes, tagInfo.ReparseTag, tagError));
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.PlaceholderFixture", L"hydrated-cloud-file;filesystem=NTFS;network=none",
                          0u, tagInfo.ReparseTag, placeholderAttributes, S_OK);

        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"skip"})json"));
        SetFileOpsAutoConcurrencyOverrideForSelfTest(true, 4u, FILESYSTEM_STORAGE_SSD);
        HostSetAutoAcceptPrompts(false);
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcTree},
                                                 std::filesystem::path(dummyRoot),
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"preserve"})json"));
            Fail(L"Beeline Copy Skip could not start its bridge copy.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }

    using Task = FolderWindow::FileOperationState::Task;
    Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
    {
        task->SubmitConflictDecision(Task::ConflictAction::Cancel, false);
        Fail(L"The fully hydrated placeholder must copy without a hydration or conflict prompt.");
        return true;
    }
    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"preserve"})json"));
    SetFileOpsAutoConcurrencyOverrideForSelfTest(false, 1u, FILESYSTEM_STORAGE_UNKNOWN);
    HostSetAutoAcceptPrompts(true);

    wil::com_ptr<IFileSystemIO> dummyIo;
    unsigned long copiedAttributes = 0u;
    unsigned long linkAttributes   = 0u;
    const bool copiedFileExists    = SUCCEEDED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) && dummyIo &&
                                     SUCCEEDED(dummyIo->GetAttributes((dummyTree + L"/placeholder.bin").c_str(), &copiedAttributes));
    const bool linkCopied = dummyIo && SUCCEEDED(dummyIo->GetAttributes((dummyTree + L"/link").c_str(), &linkAttributes));
    std::string copiedBytes;
    const bool bytesMatch = copiedFileExists && ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyTree + L"/placeholder.bin"), copiedBytes) &&
                            copiedBytes.size() == 64u * 1024u && std::ranges::all_of(copiedBytes, [](char value) noexcept { return value == 0x5A; });
    if (completed->second.hr != S_FALSE || ! bytesMatch || linkCopied || completed->second.bridgeFileAdmissionCount != 1u ||
        completed->second.conflictPromptCount != 0u)
    {
        Fail(std::format(L"Copy Skip must copy the hydrated placeholder on a worker without hydration and skip only the junction "
                         L"(hr=0x{:08X}, copied={}, linkCopied={}, admissions={}, prompts={}, bytes={}).",
                         static_cast<unsigned long>(completed->second.hr),
                         bytesMatch ? 1 : 0,
                         linkCopied ? 1 : 0, completed->second.bridgeFileAdmissionCount, completed->second.conflictPromptCount, copiedBytes.size()));
        return true;
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.CopySkipLinksKeepsPlaceholders",
                      std::format(L"shape=hydrated-cloud-file+junction;policy=skip;sourceVolume={};filesystem=NTFS;bytes={}", srcRoot.root_path().native(), copiedBytes.size()),
                      (GetTickCount64() - state.phaseStartTick) * 1000u, 1u, completed->second.bridgeFileAdmissionCount, completed->second.hr);
    state.placeholderSyncRoot.reset();
    if (FAILED(state.placeholderSyncRootStorage.cleanupHr))
    {
        Fail(L"Placeholder fixture could not unregister its local sync root.");
        return true;
    }
    static_cast<void>(SelfTest::RemoveAll(srcRoot));
    cleanupAlternate();
    state.taskA.reset();
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
            Fail(L"Causeway paused-reader stop or peer-conflict cancellation did not honor its wakeup contract.");
            return true;
        }

        SetFileOpsBridgeFailNextFileCopiesForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest());
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));
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
    if (attempts != 1u || completed->second.hr != S_FALSE)
    {
        Fail(std::format(L"Causeway worker-failure propagation expected one failure followed by an intentional Skip; attempts={} hr=0x{:08X}.",
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
    NextStep(state, SelfTestState::Step::Floodgate_CrossFsMoveDestinationGetSizeFailurePreservesSource);
    return false;
}
case SelfTestState::Step::Floodgate_CrossFsMoveDestinationGetSizeFailurePreservesSource:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextDestinationGetSizeAttemptsForSelfTest());
        Fail(L"Floodgate_CrossFsMoveDestinationGetSizeFailurePreservesSource timed out.");
        return true;
    }

    constexpr std::wstring_view kDummyRoot = L"/floodgate-bridge-destination-getsize-src";
    const std::wstring dummyFile = std::wstring(kDummyRoot) + L"/payload.txt";
    const std::filesystem::path localDest = state.tempRoot / L"floodgate-bridge-destination-getsize-dst";
    const std::filesystem::path localDestFile = localDest / L"payload.txt";
    constexpr std::string_view kSourceContents = "source-and-promoted-destination-must-survive-destination-getsize-failure";

    if (state.stepState == 0)
    {
        SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextDestinationGetSizeAttemptsForSelfTest());
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        const FileSystemFlags cleanupFlags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        static_cast<void>(state.fsDummy->DeleteItem(std::wstring(kDummyRoot).c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), kDummyRoot) || ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, kSourceContents) ||
            ! RecreateEmptyDirectory(localDest))
        {
            Fail(L"Floodgate destination GetSize failure test failed to seed source/destination.");
            return true;
        }

        SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(1);
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
            SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(0);
            static_cast<void>(TakeFileOpsBridgeFailNextDestinationGetSizeAttemptsForSelfTest());
            Fail(L"Floodgate destination GetSize failure test failed to start the cross-FS move.");
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

    const unsigned long attempts = TakeFileOpsBridgeFailNextDestinationGetSizeAttemptsForSelfTest();
    SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(0);
    if (attempts != 0u)
    {
        Fail(std::format(L"Floodgate CopyOnly MOVE unexpectedly reached the retired destination-cleanup size hook {} time(s).", attempts));
        return true;
    }
    if (completed->second.bridgeCommitSizeProofCount != 0u || completed->second.bridgeCommitSizeProofFallbackCount != 0u ||
        completed->second.bridgeImmediateDestinationSizeProbeCount != 0u)
    {
        Fail(std::format(L"Floodgate CopyOnly MOVE expected exact published-object verification without a by-path reread, got proof={} fallback={} "
                         L"immediate={}.",
                         completed->second.bridgeCommitSizeProofCount,
                         completed->second.bridgeCommitSizeProofFallbackCount,
                         completed->second.bridgeImmediateDestinationSizeProbeCount));
        return true;
    }
    if (completed->second.hr != S_FALSE)
    {
        Fail(std::format(L"Floodgate CopyOnly MOVE expected truthful source-kept S_FALSE, got 0x{:08X}.",
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    wil::com_ptr<IFileSystemIO> localIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo ||
        FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
    {
        Fail(L"Floodgate destination GetSize failure test could not query source/destination IFileSystemIO.");
        return true;
    }

    std::string sourceText;
    std::string destinationText;
    if (! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyFile), sourceText) || sourceText != kSourceContents)
    {
        Fail(L"Floodgate destination GetSize failure test did not preserve the source bytes.");
        return true;
    }
    if (! ReadFileTextFsIo(localIo, localDestFile, destinationText) || destinationText != kSourceContents)
    {
        Fail(L"Floodgate destination GetSize failure test did not retain the valid promoted destination bytes.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateBridgeCopyOnlyCleanupHookUnreached",
                      L"strategy=copy-only;verification=bound-published-object",
                      attempts,
                      0u,
                      kSourceContents.size(),
                      completed->second.hr);
    NextStep(state, SelfTestState::Step::Cinderstar_LegacyWriterEventuallyConsistentMove);
    return false;
}
case SelfTestState::Step::Cinderstar_LegacyWriterEventuallyConsistentMove:
case SelfTestState::Step::Cinderstar_LegacyWriterPermanentMissMove:
case SelfTestState::Step::Cinderstar_LegacyWriterWrongSizeMove:
{
    const bool eventuallyConsistent = state.step == SelfTestState::Step::Cinderstar_LegacyWriterEventuallyConsistentMove;
    const bool permanentMiss = state.step == SelfTestState::Step::Cinderstar_LegacyWriterPermanentMissMove;
    const std::wstring_view scenario = eventuallyConsistent ? L"eventual" : (permanentMiss ? L"permanent" : L"wrong-size");
    const std::filesystem::path localSource = state.tempRoot / std::format(L"c20-{}-src", scenario);
    const std::filesystem::path localSourceFile = localSource / L"payload.txt";
    const std::wstring dummyDest = std::format(L"/c20-{}-dst", scenario);
    const std::wstring dummyDestFile = dummyDest + L"/payload.txt";
    constexpr std::string_view kSourceContents = "cinderstar-legacy-writer-publication-verification";
    const auto resetHooks = []() noexcept
    {
        SetFileOpsBridgeFailNextDestinationOpenForSelfTest(0, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
        static_cast<void>(TakeFileOpsBridgeFailNextDestinationOpenAttemptsForSelfTest());
        SetFileOpsBridgeReportWrongDestinationSizeForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeReportWrongDestinationSizeAttemptsForSelfTest());
    };

    if (HasTimedOut(state, GetTickCount64(), 120'000ull))
    {
        resetHooks();
        Fail(std::format(L"Cinderstar legacy-writer {} MOVE timed out.", scenario));
        return true;
    }

    if (state.stepState == 0u)
    {
        resetHooks();
        if (! ShouldRetryPublishedDestinationVerification(HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) ||
            ! ShouldRetryPublishedDestinationVerification(HRESULT_FROM_WIN32(ERROR_NETWORK_BUSY)) ||
            ShouldRetryPublishedDestinationVerification(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)) ||
            ShouldRetryPublishedDestinationVerification(HRESULT_FROM_WIN32(ERROR_INVALID_DATA)) ||
            ShouldRetryPublishedDestinationVerification(E_POINTER))
        {
            Fail(L"Cinderstar published-destination retry classifier broadened beyond missing/transient failures.");
            return true;
        }
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0","writerProof":false})json"));
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        const FileSystemFlags cleanupFlags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        wil::com_ptr<IFileSystemIO> localIo;
        const bool hasLocalIo = SUCCEEDED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) && localIo;
        static_cast<void>(state.fsDummy->DeleteItem(dummyDest.c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! hasLocalIo || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyDest) || ! RecreateEmptyDirectory(localSource) ||
            ! WriteFileTextFsIo(localIo, localSourceFile, kSourceContents))
        {
            Fail(std::format(L"Cinderstar legacy-writer {} MOVE failed to seed source/destination.", scenario));
            return true;
        }

        if (eventuallyConsistent)
        {
            SetFileOpsBridgeFailNextDestinationOpenForSelfTest(1u, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
        }
        else if (permanentMiss)
        {
            SetFileOpsBridgeFailNextDestinationOpenForSelfTest(3u, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
        }
        else
        {
            SetFileOpsBridgeReportWrongDestinationSizeForSelfTest(1u);
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSourceFile},
                                                 std::filesystem::path(dummyDest),
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            resetHooks();
            Fail(std::format(L"Cinderstar legacy-writer {} MOVE failed to start.", scenario));
            return true;
        }

        state.stepState = 1u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    const unsigned long openFailures = TakeFileOpsBridgeFailNextDestinationOpenAttemptsForSelfTest();
    const unsigned long wrongSizes = TakeFileOpsBridgeReportWrongDestinationSizeAttemptsForSelfTest();
    SetFileOpsBridgeFailNextDestinationOpenForSelfTest(0, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
    SetFileOpsBridgeReportWrongDestinationSizeForSelfTest(0);
    const uint64_t expectedProbeCount = eventuallyConsistent ? 2u : (permanentMiss ? 3u : 1u);
    const unsigned long expectedOpenFailures = eventuallyConsistent ? 1u : (permanentMiss ? 3u : 0u);
    const unsigned long expectedWrongSizes = permanentMiss || eventuallyConsistent ? 0u : 1u;
    const HRESULT expectedHr = eventuallyConsistent ? S_FALSE : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    if (openFailures != expectedOpenFailures || wrongSizes != expectedWrongSizes ||
        completed->second.bridgeImmediateDestinationSizeProbeCount != expectedProbeCount || completed->second.hr != expectedHr)
    {
        Fail(std::format(L"Cinderstar legacy-writer {} MOVE mismatch: hr=0x{:08X} opens={} wrongSizes={} probes={}.",
                         scenario,
                         static_cast<unsigned long>(completed->second.hr),
                         openFailures,
                         wrongSizes,
                         completed->second.bridgeImmediateDestinationSizeProbeCount));
        return true;
    }
    if (completed->second.bridgeCommitSizeProofCount != 0u || completed->second.bridgeCommitSizeProofFallbackCount != 0u)
    {
        Fail(std::format(L"Cinderstar legacy-writer {} Copy-only MOVE unexpectedly used Managed commit-size proof accounting.", scenario));
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    wil::com_ptr<IFileSystemIO> localIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo ||
        FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
    {
        Fail(L"Cinderstar legacy-writer MOVE could not query source/destination I/O.");
        return true;
    }

    std::string sourceText;
    std::string destinationText;
    const bool sourcePresent = ReadFileTextFsIo(localIo, localSourceFile, sourceText);
    if (! sourcePresent || sourceText != kSourceContents)
    {
        Fail(std::format(L"Cinderstar legacy-writer {} MOVE source retention/deletion was incorrect.", scenario));
        return true;
    }
    if (! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyDestFile), destinationText) || destinationText != kSourceContents)
    {
        Fail(std::format(L"Cinderstar legacy-writer {} MOVE destination bytes were incorrect.", scenario));
        return true;
    }

    const uint64_t durationUs = static_cast<uint64_t>(completed->second.completionTick - state.stepStartTick) * 1000u;
    Debug::Perf::Emit(L"FileOps.SelfTest.CinderstarLegacyWriterReverify",
                      std::format(L"{};strategy=copy-only", scenario),
                      durationUs,
                      completed->second.bridgeImmediateDestinationSizeProbeCount,
                      completed->second.bridgeImmediateDestinationSizeProbeUs,
                      completed->second.hr);
    NextStep(state,
             eventuallyConsistent ? SelfTestState::Step::Cinderstar_LegacyWriterPermanentMissMove
                                  : (permanentMiss ? SelfTestState::Step::Cinderstar_LegacyWriterWrongSizeMove
                                                   : SelfTestState::Step::Cinderstar_LegacyWriterConcurrentReplacementCopy));
    return false;
}
case SelfTestState::Step::Cinderstar_LegacyWriterConcurrentReplacementCopy:
{
    constexpr const wchar_t* kReplacementPathEnv = L"REDSALAMANDER_FILEOPS_BRIDGE_REPLACE_PUBLISHED_DESTINATION_PATH";
    constexpr const wchar_t* kReplacementPayloadEnv = L"REDSALAMANDER_FILEOPS_BRIDGE_REPLACE_PUBLISHED_DESTINATION_PAYLOAD";
    constexpr std::wstring_view kDummyRoot = L"/c20-replacement";
    const std::wstring dummyFile = std::wstring(kDummyRoot) + L"/payload.txt";
    const std::filesystem::path localDest = state.tempRoot / L"c20-replacement-dst";
    const std::filesystem::path localDestFile = localDest / L"payload.txt";
    constexpr std::string_view kSourceContents = "operation-owned-published-bytes";
    constexpr std::wstring_view kReplacementContents = L"replacement-owned-by-another-writer";
    constexpr std::string_view kReplacementContentsUtf8 = "replacement-owned-by-another-writer";
    const auto clearHooks = [&]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kReplacementPathEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kReplacementPayloadEnv, nullptr));
        static_cast<void>(TakeFileOpsBridgeReplacePublishedDestinationAttemptsForSelfTest());
    };

    if (HasTimedOut(state, GetTickCount64(), 120'000ull))
    {
        clearHooks();
        Fail(L"Cinderstar concurrent replacement COPY timed out.");
        return true;
    }

    if (state.stepState == 0u)
    {
        clearHooks();
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        const FileSystemFlags cleanupFlags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        static_cast<void>(state.fsDummy->DeleteItem(std::wstring(kDummyRoot).c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), kDummyRoot) || ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, kSourceContents) ||
            ! RecreateEmptyDirectory(localDest) || SetEnvironmentVariableW(kReplacementPathEnv, localDestFile.c_str()) == 0 ||
            SetEnvironmentVariableW(kReplacementPayloadEnv, std::wstring(kReplacementContents).c_str()) == 0)
        {
            clearHooks();
            Fail(L"Cinderstar concurrent replacement COPY failed to seed source/destination or hooks.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyFile)},
                                                 localDest,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            clearHooks();
            Fail(L"Cinderstar concurrent replacement COPY failed to start.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    static_cast<void>(SetEnvironmentVariableW(kReplacementPathEnv, nullptr));
    static_cast<void>(SetEnvironmentVariableW(kReplacementPayloadEnv, nullptr));
    const unsigned long replacements = TakeFileOpsBridgeReplacePublishedDestinationAttemptsForSelfTest();
    if (replacements != 1u || completed->second.hr != S_OK || completed->second.bridgeImmediateDestinationSizeProbeCount != 0u)
    {
        Fail(std::format(L"Cinderstar exact-publication concurrent replacement COPY mismatch: replacements={} hr=0x{:08X} pathnameProbes={}.",
                         replacements,
                         static_cast<unsigned long>(completed->second.hr),
                         completed->second.bridgeImmediateDestinationSizeProbeCount));
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    wil::com_ptr<IFileSystemIO> localIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo ||
        FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
    {
        Fail(L"Cinderstar concurrent replacement COPY could not query source/destination I/O.");
        return true;
    }
    std::string sourceText;
    std::string destinationText;
    if (! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyFile), sourceText) || sourceText != kSourceContents ||
        ! ReadFileTextFsIo(localIo, localDestFile, destinationText) || destinationText != kReplacementContentsUtf8)
    {
        Fail(L"Cinderstar concurrent replacement COPY did not preserve both the source and replacement bytes.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.CinderstarExactPublicationConcurrentReplacement",
                      L"copy;verification=retained-published-object",
                      static_cast<uint64_t>(completed->second.completionTick - state.stepStartTick) * 1000u,
                      completed->second.bridgeImmediateDestinationSizeProbeCount,
                      replacements,
                      completed->second.hr);
    NextStep(state, SelfTestState::Step::Cinderstar_LegacyWriterCancelDuringBackoff);
    return false;
}
case SelfTestState::Step::Cinderstar_LegacyWriterCancelDuringBackoff:
{
    const std::filesystem::path localSource = state.tempRoot / L"c20-cancel-src";
    const std::filesystem::path localSourceFile = localSource / L"payload.txt";
    constexpr std::wstring_view kDummyDest = L"/c20-cancel-dst";
    const std::wstring dummyDestFile = std::wstring(kDummyDest) + L"/payload.txt";
    constexpr std::string_view kSourceContents = "cancel-during-publication-reverify-backoff";
    const auto clearHooks = []() noexcept
    {
        ReleaseFileOpsBridgePublishedDestinationRetryPauseForSelfTest();
        SetFileOpsBridgePublishedDestinationRetryPauseForSelfTest(false);
        SetFileOpsBridgeFailNextDestinationOpenForSelfTest(0, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
    };

    if (HasTimedOut(state, GetTickCount64(), 120'000ull))
    {
        clearHooks();
        static_cast<void>(TakeFileOpsBridgeFailNextDestinationOpenAttemptsForSelfTest());
        Fail(L"Cinderstar cancel-during-backoff MOVE timed out.");
        return true;
    }

    if (state.stepState == 0u)
    {
        clearHooks();
        static_cast<void>(TakeFileOpsBridgeFailNextDestinationOpenAttemptsForSelfTest());
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0","writerProof":false})json"));
        const FileSystemFlags cleanupFlags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        wil::com_ptr<IFileSystemIO> localIo;
        const bool hasLocalIo = SUCCEEDED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) && localIo;
        static_cast<void>(state.fsDummy->DeleteItem(std::wstring(kDummyDest).c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! hasLocalIo || ! EnsureDummyFolderExists(state.fsDummy.get(), kDummyDest) || ! RecreateEmptyDirectory(localSource) ||
            ! WriteFileTextFsIo(localIo, localSourceFile, kSourceContents))
        {
            Fail(L"Cinderstar cancel-during-backoff MOVE failed to seed source/destination.");
            return true;
        }

        SetFileOpsBridgeFailNextDestinationOpenForSelfTest(3u, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
        SetFileOpsBridgePublishedDestinationRetryPauseForSelfTest(true);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSourceFile},
                                                 std::filesystem::path(kDummyDest),
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            clearHooks();
            Fail(L"Cinderstar cancel-during-backoff MOVE failed to start.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        if (! HasFileOpsBridgePublishedDestinationRetryPauseEnteredForSelfTest())
        {
            if (state.taskA.has_value() && state.completedTasks.contains(state.taskA.value()))
            {
                clearHooks();
                Fail(L"Cinderstar cancel-during-backoff MOVE completed before entering retry backoff.");
                return true;
            }
            return false;
        }

        auto* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (task == nullptr)
        {
            clearHooks();
            Fail(L"Cinderstar cancel-during-backoff MOVE lost its live task.");
            return true;
        }
        task->RequestCancel();
        ReleaseFileOpsBridgePublishedDestinationRetryPauseForSelfTest();
        state.stepState = 2u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    const unsigned long attempts = TakeFileOpsBridgeFailNextDestinationOpenAttemptsForSelfTest();
    clearHooks();
    const bool cancelled = completed->second.hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || completed->second.hr == E_ABORT;
    if (! cancelled || attempts != 1u || completed->second.bridgeImmediateDestinationSizeProbeCount != 1u)
    {
        Fail(std::format(L"Cinderstar cancel-during-backoff MOVE mismatch: hr=0x{:08X} attempts={} probes={}.",
                         static_cast<unsigned long>(completed->second.hr),
                         attempts,
                         completed->second.bridgeImmediateDestinationSizeProbeCount));
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    wil::com_ptr<IFileSystemIO> localIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo ||
        FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
    {
        Fail(L"Cinderstar cancel-during-backoff MOVE could not query source/destination I/O.");
        return true;
    }
    std::string sourceText;
    std::string destinationText;
    if (! ReadFileTextFsIo(localIo, localSourceFile, sourceText) || sourceText != kSourceContents ||
        ! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyDestFile), destinationText) || destinationText != kSourceContents)
    {
        Fail(L"Cinderstar cancel-during-backoff MOVE did not preserve both source and published destination bytes.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.CinderstarLegacyWriterCancelBackoff",
                      L"move;strategy=copy-only",
                      static_cast<uint64_t>(completed->second.completionTick - state.stepStartTick) * 1000u,
                      attempts,
                      completed->second.bridgeImmediateDestinationSizeProbeUs,
                      completed->second.hr);
    NextStep(state, SelfTestState::Step::Floodgate_CrossFsCopyOnlyUsesCopyVerification);
    return false;
}
case SelfTestState::Step::Floodgate_CrossFsCopyOnlyUsesCopyVerification:
{
    constexpr size_t kFileCount = 16u;
    constexpr size_t kFileBytes = 256u;
    constexpr unsigned long kMetadataLatencyMs = 25u;
    constexpr std::wstring_view kDummyDestinationRoot = L"/floodgate-commit-size-proof-dst";
    const std::filesystem::path localSourceRoot = state.tempRoot / L"floodgate-commit-size-proof-src";

    if (HasTimedOut(state, GetTickCount64(), 180'000ull))
    {
        Fail(L"Floodgate_CrossFsCopyOnlyUsesCopyVerification timed out.");
        return true;
    }

    if (state.stepState == 0u)
    {
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":25,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0","writerProof":false})json"));
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4})json"));

        const FileSystemFlags cleanupFlags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        static_cast<void>(state.fsDummy->DeleteItem(std::wstring(kDummyDestinationRoot).c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), kDummyDestinationRoot) || ! RecreateEmptyDirectory(localSourceRoot))
        {
            Fail(L"Floodgate commit-size proof test failed to reset source/destination roots.");
            return true;
        }

        for (size_t index = 0u; index < kFileCount; ++index)
        {
            const std::filesystem::path sourceFile = localSourceRoot / std::format(L"payload-{:02}.bin", index);
            if (! WriteFilledTestFile(sourceFile, kFileBytes, 0x5au))
            {
                Fail(std::format(L"Floodgate commit-size proof test failed to seed '{}'.", sourceFile.wstring()));
                return true;
            }
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSourceRoot},
                                                 std::filesystem::path(kDummyDestinationRoot),
                                                 flags,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            Fail(L"Floodgate commit-size proof test failed to start the local-to-Dummy MOVE.");
            return true;
        }

        state.stepState = 1u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    const CompletedTaskInfo& info = completed->second;
    if (info.hr != S_FALSE)
    {
        Fail(std::format(L"Floodgate CopyOnly MOVE expected S_FALSE, got 0x{:08X}.", static_cast<unsigned long>(info.hr)));
        return true;
    }

    const bool copyOnlyShape = info.bridgeCommitSizeProofCount == 0u && info.bridgeCommitSizeProofFallbackCount == 0u &&
                               info.bridgeImmediateDestinationSizeProbeCount == kFileCount;
    if (! copyOnlyShape)
    {
        Fail(std::format(L"Floodgate CopyOnly metric shape mismatch: proof={} fallback={} immediate={} files={}.",
                         info.bridgeCommitSizeProofCount,
                         info.bridgeCommitSizeProofFallbackCount,
                         info.bridgeImmediateDestinationSizeProbeCount,
                         kFileCount));
        return true;
    }

    std::error_code ec;
    if (! std::filesystem::exists(localSourceRoot, ec) || ec)
    {
        Fail(L"Floodgate CopyOnly MOVE removed the local source tree.");
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
    {
        Fail(L"Floodgate commit-size proof test could not query Dummy IFileSystemIO.");
        return true;
    }

    const std::string expected(kFileBytes, static_cast<char>(0x5a));
    for (size_t index = 0u; index < kFileCount; ++index)
    {
        const std::filesystem::path destinationFile = std::filesystem::path(kDummyDestinationRoot) /
                                                      localSourceRoot.filename() /
                                                      std::format(L"payload-{:02}.bin", index);
        std::string actual;
        if (! ReadFileTextFsIo(dummyIo, destinationFile, actual) || actual != expected)
        {
            Fail(std::format(L"Floodgate CopyOnly MOVE produced invalid bytes at '{}'.", destinationFile.wstring()));
            return true;
        }
    }

    const std::wstring detail = std::format(L"shape=copy-only files={} metadataLatencyMs={} immediateProbeUs={} proof={} fallback={}",
                                             kFileCount,
                                             kMetadataLatencyMs,
                                             info.bridgeImmediateDestinationSizeProbeUs,
                                             info.bridgeCommitSizeProofCount,
                                             info.bridgeCommitSizeProofFallbackCount);
    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateBridgeCommitSizeProof",
                      detail,
                      info.bridgeImmediateDestinationSizeProbeUs,
                      info.bridgeImmediateDestinationSizeProbeCount,
                      0u,
                      info.hr);
    NextStep(state, SelfTestState::Step::Floodgate_CrossFsMoveGetSizeFailuresParallelPreserveSourceTree);
    return false;
}
case SelfTestState::Step::Floodgate_CrossFsMoveGetSizeFailuresParallelPreserveSourceTree:
{
    const ULONGLONG nowTick = GetTickCount64();
    const auto resetHooks = []() noexcept
    {
        SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest());
        SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(0);
        static_cast<void>(TakeFileOpsBridgeFailNextDestinationGetSizeAttemptsForSelfTest());
    };
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        resetHooks();
        Fail(L"Floodgate_CrossFsMoveGetSizeFailuresParallelPreserveSourceTree timed out.");
        return true;
    }

    constexpr std::wstring_view kDummyRoot = L"/floodgate-bridge-getsize-parallel-src";
    const std::filesystem::path localDest = state.tempRoot / L"floodgate-bridge-getsize-parallel-dst";
    const std::filesystem::path localDestRoot = localDest / L"floodgate-bridge-getsize-parallel-src";
    constexpr std::array<std::wstring_view, 8> kFileNames{{
        L"payload-00.bin",
        L"payload-01.bin",
        L"payload-02.bin",
        L"payload-03.bin",
        L"payload-04.bin",
        L"payload-05.bin",
        L"payload-06.bin",
        L"payload-07.bin",
    }};
    constexpr std::string_view kSourceContents = "parallel-getsize-failure-payload";

    if (state.stepState == 0)
    {
        resetHooks();
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":10,"streamChunkLatencyMs":20,"virtualSpeedLimit":"0"})json"));
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));

        const FileSystemFlags cleanupFlags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        static_cast<void>(state.fsDummy->DeleteItem(std::wstring(kDummyRoot).c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), kDummyRoot) || ! RecreateEmptyDirectory(localDest))
        {
            Fail(L"Floodgate parallel GetSize failure test failed to reset source/destination roots.");
            return true;
        }
        for (const std::wstring_view fileName : kFileNames)
        {
            const std::wstring dummyFile = std::wstring(kDummyRoot) + L"/" + std::wstring(fileName);
            if (! DummyWriteTextFile(state.fsDummy.get(), dummyFile, kSourceContents))
            {
                Fail(std::format(L"Floodgate parallel GetSize failure test failed to seed '{}'.", dummyFile));
                return true;
            }
        }

        SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(1);
        SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(1);
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
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
            resetHooks();
            Fail(L"Floodgate parallel GetSize failure test failed to start the cross-FS move.");
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

    const unsigned long sourceAttempts = TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest();
    const unsigned long destinationAttempts = TakeFileOpsBridgeFailNextDestinationGetSizeAttemptsForSelfTest();
    SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0);
    SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(0);
    if (sourceAttempts != 1u || destinationAttempts != 0u)
    {
        Fail(std::format(L"Floodgate parallel CopyOnly test expected one source-size failure and no managed-Move destination cleanup probe, got "
                         L"source={} destination={}.",
                         sourceAttempts,
                         destinationAttempts));
        return true;
    }
    if (completed->second.bridgeCommitSizeProofFallbackCount != 0u)
    {
        Fail(std::format(L"Floodgate parallel CopyOnly test unexpectedly used a fallback publication proof: {}.",
                         completed->second.bridgeCommitSizeProofFallbackCount));
        return true;
    }
    if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        Fail(std::format(L"Floodgate parallel GetSize failure test expected ERROR_PARTIAL_COPY, got 0x{:08X}.",
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    wil::com_ptr<IFileSystemIO> localIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo ||
        FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
    {
        Fail(L"Floodgate parallel GetSize failure test could not query source/destination IFileSystemIO.");
        return true;
    }

    unsigned int sourcePresentCount = 0;
    unsigned int destinationPresentCount = 0;
    unsigned int sourceOnlyCount = 0;
    unsigned int bothPresentCount = 0;
    for (const std::wstring_view fileName : kFileNames)
    {
        const std::filesystem::path sourcePath = std::filesystem::path(kDummyRoot) / fileName;
        const std::filesystem::path destinationPath = localDestRoot / fileName;
        std::string sourceText;
        std::string destinationText;
        const bool sourcePresent = ReadFileTextFsIo(dummyIo, sourcePath, sourceText);
        const bool destinationPresent = ReadFileTextFsIo(localIo, destinationPath, destinationText);
        if ((sourcePresent && sourceText != kSourceContents) || (destinationPresent && destinationText != kSourceContents) ||
            (! sourcePresent && ! destinationPresent))
        {
            Fail(std::format(L"Floodgate parallel GetSize failure test found invalid source/destination state for '{}'.", fileName));
            return true;
        }

        sourcePresentCount += sourcePresent ? 1u : 0u;
        destinationPresentCount += destinationPresent ? 1u : 0u;
        sourceOnlyCount += sourcePresent && ! destinationPresent ? 1u : 0u;
        bothPresentCount += sourcePresent && destinationPresent ? 1u : 0u;
    }

    // CopyOnly never arms a source-cleanup pass. All sources remain, including siblings whose exact
    // published destinations were verified; the one source-size failure has no destination publication.
    if (sourcePresentCount != 8u || destinationPresentCount != 7u || sourceOnlyCount != 1u || bothPresentCount != 7u)
    {
        Fail(std::format(L"Floodgate parallel GetSize failure topology mismatch: sources={} destinations={} sourceOnly={} both={}.",
                         sourcePresentCount,
                         destinationPresentCount,
                         sourceOnlyCount,
                         bothPresentCount));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateBridgeCopyOnlyGetSizeFailuresParallel",
                      L"concurrency=4;source-retained;cleanup-probes=0",
                      sourceAttempts + destinationAttempts,
                      sourcePresentCount,
                      destinationPresentCount,
                      completed->second.hr);
    NextStep(state, SelfTestState::Step::Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup);
    return false;
}
case SelfTestState::Step::Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup timed out.");
        return true;
    }

    const std::wstring dummyRoot = L"/floodgate-crossfs-file-cleanup";
    const std::wstring dummyFile = dummyRoot + L"/payload.txt";
    const std::filesystem::path localDest = state.tempRoot / L"floodgate-crossfs-file-cleanup-dst";
    const std::filesystem::path localDestFile = localDest / L"payload.txt";
    constexpr std::string_view kSourceContents = "source-before-move";

    wil::com_ptr<IFileSystemIO> dummyIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
    {
        Fail(L"Floodgate bridge file MOVE cleanup corruption test could not query dummy IFileSystemIO.");
        return true;
    }

    wil::com_ptr<IFileSystemIO> localIo;
    if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
    {
        Fail(L"Floodgate bridge file MOVE cleanup corruption test could not query local IFileSystemIO.");
        return true;
    }

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0","writerProof":false})json"));
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json"));

        const FileSystemFlags dummyCleanupFlags =
            static_cast<FileSystemFlags>(static_cast<uint32_t>(FILESYSTEM_FLAG_RECURSIVE) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY));
        static_cast<void>(state.fsDummy->DeleteItem(dummyRoot.c_str(), dummyCleanupFlags, nullptr, nullptr, nullptr));

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, kSourceContents) ||
            ! RecreateEmptyDirectory(localDest))
        {
            Fail(L"Floodgate bridge file MOVE cleanup corruption test failed to seed source/destination.");
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

    if (completed->second.hr != S_FALSE)
    {
        Fail(std::format(L"Floodgate CopyOnly MOVE expected S_FALSE, got 0x{:08X}.",
                         static_cast<unsigned long>(completed->second.hr)));
        return true;
    }

    std::string text;
    if (! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyFile), text) || text != kSourceContents)
    {
        Fail(L"Floodgate CopyOnly MOVE did not preserve the source bytes.");
        return true;
    }
    if (! ReadFileTextFsIo(localIo, localDestFile, text) || text != kSourceContents)
    {
        Fail(L"Floodgate CopyOnly MOVE did not publish the source bytes unchanged.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateCrossFsFileMoveCleanupCorruption",
                      L"copy-only;source-retained;no-cleanup-authority",
                      0,
                      completed->second.progressCompletedBytes,
                      0,
                      completed->second.hr);
    NextStep(state, SelfTestState::Step::Floodgate_CrossFsDirectoryCopyOnlyRetainsSource);
    return false;
}
case SelfTestState::Step::Floodgate_CrossFsDirectoryCopyOnlyRetainsSource:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
        Fail(L"Floodgate_CrossFsDirectoryCopyOnlyRetainsSource timed out.");
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
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0","writerProof":false})json"));
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
            Fail(L"Floodgate bridge directory MOVE cleanup test failed to seed source/destination.");
            return true;
        }

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
            Fail(L"Floodgate bridge directory MOVE cleanup test failed to start the cross-FS move.");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    if (completed->second.hr != S_FALSE)
    {
        Fail(std::format(L"Floodgate directory CopyOnly MOVE expected S_FALSE, got 0x{:08X}.",
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
    if (! ReadFileTextFsIo(localIo, stableFile, text) || text != kStableBefore)
    {
        Fail(L"Floodgate directory CopyOnly MOVE did not retain the source file bytes.");
        return true;
    }
    std::error_code ec;
    if (! std::filesystem::exists(sourceTree, ec) || ec || ! std::filesystem::exists(sourceLevel, ec) || ec)
    {
        Fail(L"Floodgate directory CopyOnly MOVE removed part of the source tree.");
        return true;
    }

    if (! ReadFileTextFsIo(dummyIo, std::filesystem::path(dummyStable), text) || text != kStableBefore)
    {
        Fail(L"Floodgate directory CopyOnly MOVE did not copy the stable file to the destination.");
        return true;
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateCrossFsDirectoryMoveCleanup",
                      L"copy-only;source-retained;no-cleanup-authority",
                      0,
                      0,
                      0,
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
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownFired.data(), nullptr));
    auto clearStagedCopyEnv = wil::scope_exit([]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailPath.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailFired.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailPath.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailFired.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownPath.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownFired.data(), nullptr));
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

    if (! SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailPath.data(), targetPath.c_str()) ||
        ! SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownPath.data(), targetPath.c_str()))
    {
        Fail(L"Floodgate local copy overwrite test failed to arm the cleanup-unknown receipt fixture.");
        return true;
    }
    FileOpsRecursiveProgressRecorder cleanupUnknownRecorder{};
    const HRESULT cleanupUnknownHr = state.fsLocal->CopyItem(
        sourcePath.c_str(), targetPath.c_str(), overwriteFlags, nullptr, &cleanupUnknownRecorder, nullptr);
    const bool cleanupUnknownReceipt = cleanupUnknownRecorder.completedCount == 1u &&
        cleanupUnknownRecorder.lastCompletedStatus == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) &&
        cleanupUnknownRecorder.lastMutationResult.has_value() &&
        cleanupUnknownRecorder.lastMutationResult->sizeBytes >= sizeof(FileSystemItemMutationResult) &&
        cleanupUnknownRecorder.lastMutationResult->outcomeKnown != FALSE &&
        cleanupUnknownRecorder.lastMutationResult->mutationCommitted == FALSE &&
        cleanupUnknownRecorder.lastMutationResult->originalStillPresent != FALSE &&
        cleanupUnknownRecorder.lastMutationResult->ownedStageDisposition == FileSystemOwnedStageDisposition::Unknown;
    if (cleanupUnknownHr != HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) || ! cleanupUnknownReceipt ||
        GetEnvVarTrimmed(kSelfTestEnvAbortOwnedStageUnknownFired) != L"1")
    {
        Fail(std::format(
            L"Floodgate cleanup-unknown receipt must report known final non-publication plus Unknown stage (hr=0x{0:08X}, receipt={1}).",
            static_cast<unsigned long>(cleanupUnknownHr),
            cleanupUnknownReceipt));
        return true;
    }

    size_t retainedStageCount = 0u;
    bool retainedStageCleanupSucceeded = true;
    DWORD retainedStageAttributeError = ERROR_SUCCESS;
    DWORD retainedStageDeleteError = ERROR_SUCCESS;
    DWORD retainedStageProbeError = ERROR_SUCCESS;
    ec.clear();
    for (const auto& entry : std::filesystem::directory_iterator(targetPath.parent_path(), ec))
    {
        const std::wstring entryPath = entry.path().wstring();
        if (entryPath.starts_with(siblingTempPrefix.wstring()))
        {
            ++retainedStageCount;
            // The injected Unknown path deliberately bypasses provider cleanup. Antivirus or another
            // no-follow reader may briefly deny attribute/delete access or retain the test artifact
            // after DeleteFileW succeeds. Retry the exact observed sibling within a bounded window,
            // then prove it disappeared before reusing this prefix for the success path.
            bool stageGone = false;
            for (unsigned int attempt = 0u; attempt < 200u; ++attempt)
            {
                const DWORD attributes = ::GetFileAttributesW(entryPath.c_str());
                if (attributes == INVALID_FILE_ATTRIBUTES)
                {
                    const DWORD attributeError = ::GetLastError();
                    if (attributeError == ERROR_FILE_NOT_FOUND || attributeError == ERROR_PATH_NOT_FOUND)
                    {
                        stageGone = true;
                        break;
                    }
                    retainedStageProbeError = attributeError;
                    ::Sleep(10u);
                    continue;
                }

                if (::SetFileAttributesW(entryPath.c_str(), FILE_ATTRIBUTE_NORMAL) == 0)
                {
                    retainedStageAttributeError = ::GetLastError();
                    ::Sleep(10u);
                    continue;
                }
                if (::DeleteFileW(entryPath.c_str()) == 0)
                {
                    const DWORD deleteError = ::GetLastError();
                    if (deleteError == ERROR_FILE_NOT_FOUND || deleteError == ERROR_PATH_NOT_FOUND)
                    {
                        stageGone = true;
                        break;
                    }
                    retainedStageDeleteError = deleteError;
                }
                ::Sleep(10u);
            }
            retainedStageCleanupSucceeded = retainedStageCleanupSucceeded && stageGone;
        }
    }
    const bool finalReadable = ReadFileTextFsIo(localIo, targetPath, afterFailure);
    const bool finalPreserved = finalReadable && afterFailure == kOriginal;
    if (ec || retainedStageCount != 1u || ! retainedStageCleanupSucceeded || ! finalPreserved)
    {
        Fail(std::format(
            L"Floodgate cleanup-unknown fixture did not preserve the final destination or explicitly remove its one possible stage artifact "
            L"(enumerationError={0}, stageCount={1}, stageCleanup={2}, finalReadable={3}, finalPreserved={4}, "
            L"attributeError={5}, deleteError={6}, probeError={7}).",
            ec.value(),
            retainedStageCount,
            retainedStageCleanupSucceeded,
            finalReadable,
            finalPreserved,
            retainedStageAttributeError,
            retainedStageDeleteError,
            retainedStageProbeError));
        return true;
    }
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailFired.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownFired.data(), nullptr));

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

    std::string finalAttributesFailureBytes;
    if (! ReadFileTextFsIo(localIo, targetPath, finalAttributesFailureBytes) || finalAttributesFailureBytes != kReplacement)
    {
        Fail(L"Floodgate final-attributes failure must occur after publication, with replacement bytes already at the destination.");
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

    size_t remainingAddressableStageCount = 0u;
    DWORD remainingStageProbeError = ERROR_SUCCESS;
    ec.clear();
    for (const auto& entry : std::filesystem::directory_iterator(targetPath.parent_path(), ec))
    {
        if (ec)
        {
            break;
        }

        const std::wstring entryPath = entry.path().wstring();
        if (! entryPath.starts_with(siblingTempPrefix.wstring()))
        {
            continue;
        }

        if (::GetFileAttributesW(entryPath.c_str()) == INVALID_FILE_ATTRIBUTES)
        {
            const DWORD attributeError = ::GetLastError();
            if (attributeError == ERROR_FILE_NOT_FOUND || attributeError == ERROR_PATH_NOT_FOUND)
            {
                // A successful delete request can remain in a directory enumeration while another
                // process retains a delete-sharing handle. The name has no remaining path authority.
                continue;
            }
            remainingStageProbeError = attributeError;
        }
        ++remainingAddressableStageCount;
    }
    if (ec || remainingAddressableStageCount != 0u)
    {
        Fail(std::format(L"Floodgate local copy overwrite success path retained {0} addressable staged temp file(s) "
                         L"(enumerationError={1}, probeError={2}).",
                         remainingAddressableStageCount,
                         ec.value(),
                         remainingStageProbeError));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.FloodgateLocalCopyOverwriteIsStaged", L"", kOriginal.size(), kReplacement.size(), 0, S_OK);
    NextStep(state, SelfTestState::Step::Floodgate_LocalCopyNewNameConcurrentReplacementSurvives);
    return false;
}
case SelfTestState::Step::Floodgate_LocalCopyNewNameConcurrentReplacementSurvives:
{
    wil::com_ptr<IFileSystemIO> localIo;
    if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
    {
        Fail(L"R0a new-name Local Copy test could not query local IFileSystemIO.");
        return true;
    }

    const std::filesystem::path perfSource = state.tempRoot / L"r0a-direct-final-perf-source.bin";
    const std::filesystem::path perfTarget = state.tempRoot / L"r0a-direct-final-perf-target.bin";
    const std::filesystem::path perfSourceZone = std::filesystem::path(perfSource.wstring() + L":Zone.Identifier");
    const std::filesystem::path perfTargetZone = std::filesystem::path(perfTarget.wstring() + L":Zone.Identifier");
    const std::filesystem::path raceSource = state.tempRoot / L"r0a-direct-final-race-source.bin";
    const std::filesystem::path raceTarget = state.tempRoot / L"r0a-direct-final-race-target.bin";
    const std::filesystem::path raceMoved = state.tempRoot / L"r0a-direct-final-race-owned-partial.bin";
    constexpr size_t kPerfBytes = 64u * 1024u * 1024u;
    constexpr size_t kRaceBytes = 8u * 1024u * 1024u;
    constexpr std::string_view kZonePayload = "r0a-zone";
    constexpr std::string_view kForeignPayload = "r0a-foreign-replacement";

    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalRollbackSwapPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalRollbackSwapMovedPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalRollbackSwapFired.data(), nullptr));
    const auto clearR0aRaceEnv = wil::scope_exit([]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalRollbackSwapPath.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalRollbackSwapMovedPath.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalRollbackSwapFired.data(), nullptr));
    });

    std::error_code ec;
    for (const std::filesystem::path& path : {perfSource, perfTarget, raceSource, raceTarget, raceMoved})
    {
        std::filesystem::remove(path, ec);
        ec.clear();
    }
    if (! WriteTestFile(perfSource, kPerfBytes) || ! WriteFileTextFsIo(localIo, perfSourceZone, kZonePayload))
    {
        Fail(L"R0a new-name Local Copy test could not seed the performance/metadata source.");
        return true;
    }
    const DWORD sourceAttributes = GetFileAttributesW(perfSource.c_str());
    if (sourceAttributes == INVALID_FILE_ATTRIBUTES ||
        ! SetFileAttributesW(perfSource.c_str(), sourceAttributes | FILE_ATTRIBUTE_HIDDEN))
    {
        Fail(L"R0a new-name Local Copy test could not seed the source attributes.");
        return true;
    }

    FileOpsRecursiveProgressRecorder successRecorder{};
    const auto perfStart = std::chrono::steady_clock::now();
    const HRESULT perfHr = state.fsLocal->CopyItem(
        perfSource.c_str(), perfTarget.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &successRecorder, nullptr);
    const uint64_t perfUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(
        std::chrono::steady_clock::now() - perfStart).count());
    Debug::Perf::Emit(L"FileOps.SelfTest.R0aLocalDirectFinalCopyUs",
                      L"shape=new-name-local-regular-file",
                      perfUs,
                      kPerfBytes,
                      0u,
                      perfHr);
    std::string copiedZone;
    uintmax_t copiedSize = 0u;
    ec.clear();
    copiedSize = std::filesystem::file_size(perfTarget, ec);
    const DWORD copiedAttributes = GetFileAttributesW(perfTarget.c_str());
    if (FAILED(perfHr) || ec || copiedSize != kPerfBytes || copiedAttributes == INVALID_FILE_ATTRIBUTES ||
        (copiedAttributes & FILE_ATTRIBUTE_HIDDEN) == 0u || ! ReadFileTextFsIo(localIo, perfTargetZone, copiedZone) ||
        copiedZone != kZonePayload)
    {
        Fail(std::format(L"R0a new-name Local Copy success/metadata fixture failed (hr=0x{0:08X}, size={1}, ec={2}, attrs=0x{3:08X}).",
                         static_cast<unsigned long>(perfHr),
                         copiedSize,
                         ec.value(),
                         copiedAttributes));
        return true;
    }

    // Metadata the destination cannot hold is best effort, exactly as with CopyFileExW: the Copy
    // publishes complete content, and only an EFS source that cannot stay encrypted fails closed.
    {
        constexpr wchar_t kForcePresentEnv[] = L"REDSALAMANDER_FILEOPS_METADATA_FORCE_PRESENT_MASK";
        constexpr wchar_t kFailMaskEnv[]     = L"REDSALAMANDER_FILEOPS_METADATA_FAIL_MASK";
        const auto clearMetadataEnv          = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(SetEnvironmentVariableW(kForcePresentEnv, nullptr));
            static_cast<void>(SetEnvironmentVariableW(kFailMaskEnv, nullptr));
        });
        const std::filesystem::path lossSource = state.tempRoot / L"r0a-direct-final-loss-source.bin";
        const std::filesystem::path lossTarget = state.tempRoot / L"r0a-direct-final-loss-target.bin";
        const std::filesystem::path efsTarget  = state.tempRoot / L"r0a-direct-final-efs-target.bin";
        constexpr size_t kLossBytes            = 256u * 1024u;
        for (const std::filesystem::path& path : {lossSource, lossTarget, efsTarget})
        {
            std::filesystem::remove(path, ec);
            ec.clear();
        }
        if (! WriteTestFile(lossSource, kLossBytes) ||
            ! WriteFileTextFsIo(localIo, std::filesystem::path(lossSource.wstring() + L":Zone.Identifier"), kZonePayload))
        {
            Fail(L"R0a metadata-loss fixture could not seed its source.");
            return true;
        }
        const uint32_t forcedPresent = FILESYSTEM_METADATA_SPARSE;
        const uint32_t lossMask      = forcedPresent | FILESYSTEM_METADATA_MOTW | FILESYSTEM_METADATA_ALTERNATE_STREAMS;
        if (! SetEnvironmentVariableW(kForcePresentEnv, std::format(L"{}", forcedPresent).c_str()) ||
            ! SetEnvironmentVariableW(kFailMaskEnv, std::format(L"{}", lossMask).c_str()))
        {
            Fail(L"R0a metadata-loss fixture could not arm the loss injection.");
            return true;
        }
        FileOpsRecursiveProgressRecorder lossRecorder{};
        const HRESULT lossHr =
            state.fsLocal->CopyItem(lossSource.c_str(), lossTarget.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &lossRecorder, nullptr);
        ec.clear();
        const uintmax_t lossSize = std::filesystem::file_size(lossTarget, ec);
        const bool lossReceipt   = lossRecorder.completedCount == 1u && lossRecorder.lastMutationResult.has_value() &&
            lossRecorder.lastMutationResult->mutationCommitted != FALSE &&
            lossRecorder.lastMutationResult->ownedStageDisposition == FileSystemOwnedStageDisposition::Published;
        if (FAILED(lossHr) || ec || lossSize != kLossBytes || ! lossReceipt)
        {
            Fail(std::format(L"R0a direct-final Copy must publish complete content when the destination cannot hold "
                             L"sparse/stream metadata (hr=0x{0:08X}, size={1}, ec={2}).",
                             static_cast<unsigned long>(lossHr),
                             lossSize,
                             ec.value()));
            return true;
        }

        const uint32_t efsMask = FILESYSTEM_METADATA_EFS;
        if (! SetEnvironmentVariableW(kForcePresentEnv, std::format(L"{}", efsMask).c_str()) ||
            ! SetEnvironmentVariableW(kFailMaskEnv, std::format(L"{}", efsMask).c_str()))
        {
            Fail(L"R0a metadata-loss fixture could not arm the EFS injection.");
            return true;
        }
        FileOpsRecursiveProgressRecorder efsRecorder{};
        const HRESULT efsHr = state.fsLocal->CopyItem(lossSource.c_str(), efsTarget.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &efsRecorder, nullptr);
        ec.clear();
        const bool efsTargetExists = std::filesystem::exists(efsTarget, ec);
        if (efsHr != HRESULT_FROM_WIN32(ERROR_ENCRYPTION_FAILED) || efsTargetExists)
        {
            Fail(std::format(L"R0a direct-final Copy must fail closed before content when an EFS source cannot stay encrypted "
                             L"(hr=0x{0:08X}, exists={1}).",
                             static_cast<unsigned long>(efsHr),
                             efsTargetExists));
            return true;
        }

        // NTFS compression follows the destination folder, exactly as File Explorer: a plain source
        // copied into a compressed folder becomes compressed, and a compressed source copied into a
        // plain folder does not. Volumes without NTFS compression (ReFS, FAT) skip this check.
        static_cast<void>(SetEnvironmentVariableW(kForcePresentEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kFailMaskEnv, nullptr));
        const auto isCompressed = [](const std::filesystem::path& path) noexcept -> bool
        {
            const DWORD attributes = GetFileAttributesW(path.c_str());
            return attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_COMPRESSED) != 0u;
        };
        const std::filesystem::path compressedDir = state.tempRoot / L"r0a-compressed-dir";
        std::filesystem::remove_all(compressedDir, ec);
        ec.clear();
        std::filesystem::create_directories(compressedDir, ec);
        if (ec)
        {
            Fail(L"R0a compression-inheritance fixture could not create its folder.");
            return true;
        }
        bool volumeCompresses = false;
        {
            const wil::unique_hfile folder(CreateFileW(compressedDir.c_str(),
                                                       GENERIC_READ | GENERIC_WRITE,
                                                       FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                       nullptr,
                                                       OPEN_EXISTING,
                                                       FILE_FLAG_BACKUP_SEMANTICS,
                                                       nullptr));
            USHORT format    = COMPRESSION_FORMAT_DEFAULT;
            DWORD returned   = 0u;
            volumeCompresses = folder &&
                DeviceIoControl(folder.get(), FSCTL_SET_COMPRESSION, &format, sizeof(format), nullptr, 0u, &returned, nullptr) != FALSE &&
                isCompressed(compressedDir);
        }
        if (! volumeCompresses)
        {
            Debug::Perf::Emit(L"FileOps.SelfTest.R0aCompressionInheritance", L"skipped=volume-without-ntfs-compression", 0, 0, 0, S_OK);
        }
        else
        {
            const std::filesystem::path inheritTarget = compressedDir / L"inherit.bin";
            FileOpsRecursiveProgressRecorder inheritRecorder{};
            const HRESULT inheritHr =
                state.fsLocal->CopyItem(lossSource.c_str(), inheritTarget.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &inheritRecorder, nullptr);
            if (FAILED(inheritHr) || ! isCompressed(inheritTarget))
            {
                Fail(std::format(L"R0a Copy into a compressed folder must inherit the folder's compression like File Explorer "
                                 L"(hr=0x{0:08X}, compressed={1}).",
                                 static_cast<unsigned long>(inheritHr),
                                 isCompressed(inheritTarget)));
                return true;
            }
            // A file created inside the compressed folder is compressed by NTFS; copying it into the
            // plain temp root must not carry the source's compression along.
            const std::filesystem::path compressedSource = compressedDir / L"compressed-source.bin";
            if (! WriteTestFile(compressedSource, kLossBytes) || ! isCompressed(compressedSource))
            {
                Fail(L"R0a compression-inheritance fixture could not create a compressed source.");
                return true;
            }
            if (! isCompressed(state.tempRoot))
            {
                const std::filesystem::path plainTarget = state.tempRoot / L"r0a-plain-target.bin";
                std::filesystem::remove(plainTarget, ec);
                ec.clear();
                FileOpsRecursiveProgressRecorder plainRecorder{};
                const HRESULT plainHr =
                    state.fsLocal->CopyItem(compressedSource.c_str(), plainTarget.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &plainRecorder, nullptr);
                ec.clear();
                const uintmax_t plainSize = std::filesystem::file_size(plainTarget, ec);
                if (FAILED(plainHr) || ec || plainSize != kLossBytes || isCompressed(plainTarget))
                {
                    Fail(std::format(L"R0a Copy of a compressed source into a plain folder must not carry the source's compression "
                                     L"(hr=0x{0:08X}, size={1}, compressed={2}).",
                                     static_cast<unsigned long>(plainHr),
                                     plainSize,
                                     isCompressed(plainTarget)));
                    return true;
                }
            }
            Debug::Perf::Emit(L"FileOps.SelfTest.R0aCompressionInheritance", L"", 1, 0, 0, S_OK);
        }
    }

    if (! WriteTestFile(raceSource, kRaceBytes) ||
        ! SetEnvironmentVariableW(kSelfTestEnvDirectFinalRollbackSwapPath.data(), raceTarget.c_str()) ||
        ! SetEnvironmentVariableW(kSelfTestEnvDirectFinalRollbackSwapMovedPath.data(), raceMoved.c_str()))
    {
        Fail(L"R0a new-name Local Copy test could not arm the concurrent replacement fixture.");
        return true;
    }

    FileOpsRecursiveProgressRecorder raceRecorder{};
    raceRecorder.cancelAfterProgress = true;
    FileOpsRecorderOperationControl raceControl{raceRecorder};
    FileSystemOptions raceOptions{};
    raceOptions.sizeBytes        = sizeof(raceOptions);
    raceOptions.linkPolicy       = FILESYSTEM_LINK_PRESERVE;
    raceOptions.operationControl = &raceControl;
    const HRESULT raceHr = state.fsLocal->CopyItem(
        raceSource.c_str(), raceTarget.c_str(), FILESYSTEM_FLAG_NONE, &raceOptions, &raceRecorder, nullptr);
    std::string foreignBytes;
    const bool foreignSurvived = ReadFileTextFsIo(localIo, raceTarget, foreignBytes) && foreignBytes == kForeignPayload;
    ec.clear();
    const bool ownedPartialSurvived = std::filesystem::exists(raceMoved, ec) && ! ec;
    const bool raceFired = GetEnvVarTrimmed(kSelfTestEnvDirectFinalRollbackSwapFired) == L"1";
    if ((raceHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) && raceHr != E_ABORT) || ! raceFired || ! foreignSurvived || ownedPartialSurvived)
    {
        Fail(std::format(L"R0a exact abort must preserve a foreign replacement and remove only the retained partial "
                         L"(hr=0x{0:08X}, fired={1}, foreign={2}, partial={3}, ec={4}).",
                         static_cast<unsigned long>(raceHr),
                         raceFired,
                         foreignSurvived,
                         ownedPartialSurvived,
                         ec.value()));
        return true;
    }

    const bool successReceipt = successRecorder.completedCount == 1u && successRecorder.lastMutationResult.has_value() &&
        successRecorder.lastMutationResult->outcomeKnown != FALSE && successRecorder.lastMutationResult->mutationCommitted != FALSE &&
        successRecorder.lastMutationResult->originalStillPresent != FALSE &&
        successRecorder.lastMutationResult->ownedStageDisposition == FileSystemOwnedStageDisposition::Published;
    if (! successReceipt)
    {
        Fail(L"R0a successful direct-final Copy must report exact Published mutation truth.");
        return true;
    }

    const std::wstring tempPrefix = perfTarget.filename().wstring() + L".rs_copy_tmp_";
    for (const auto& entry : std::filesystem::directory_iterator(perfTarget.parent_path(), ec))
    {
        if (entry.path().filename().wstring().starts_with(tempPrefix))
        {
            Fail(L"R0a ordinary new-name Local Copy must not create a hidden sibling stage.");
            return true;
        }
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.R0aLocalDirectFinalRacePreserved", L"exact-abort", 1u, raceRecorder.progressCount, kRaceBytes, S_OK);
    NextStep(state, SelfTestState::Step::Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete);
    return false;
}
case SelfTestState::Step::Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete:
{
    const std::filesystem::path sourcePath = state.tempRoot / L"r0a-direct-final-retained-source.bin";
    const std::filesystem::path targetPath = state.tempRoot / L"r0a-direct-final-retained-target.bin";
    constexpr size_t kSourceBytes = 8u * 1024u * 1024u;

    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalAbortFailPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalAbortFailFired.data(), nullptr));
    const auto clearR0aAbortEnv = wil::scope_exit([]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalAbortFailPath.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDirectFinalAbortFailFired.data(), nullptr));
    });

    std::error_code ec;
    std::filesystem::remove(sourcePath, ec);
    ec.clear();
    std::filesystem::remove(targetPath, ec);
    if (! WriteTestFile(sourcePath, kSourceBytes) ||
        ! SetEnvironmentVariableW(kSelfTestEnvDirectFinalAbortFailPath.data(), targetPath.c_str()))
    {
        Fail(L"R0a retained-incomplete test could not seed or arm the direct-final abort fixture.");
        return true;
    }

    FileOpsRecursiveProgressRecorder recorder{};
    const HRESULT copyHr = state.fsLocal->CopyItem(
        sourcePath.c_str(), targetPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &recorder, nullptr);
    ec.clear();
    const bool retainedPartial = std::filesystem::exists(targetPath, ec) && ! ec && std::filesystem::file_size(targetPath, ec) > 0u && ! ec;
    const bool receiptMatches = recorder.completedCount == 1u &&
        recorder.lastCompletedStatus == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) && recorder.lastMutationResult.has_value() &&
        recorder.lastMutationResult->outcomeKnown != FALSE && recorder.lastMutationResult->mutationCommitted == FALSE &&
        recorder.lastMutationResult->originalStillPresent != FALSE &&
        recorder.lastMutationResult->ownedStageDisposition == FileSystemOwnedStageDisposition::RetainedIncomplete;
    const bool abortFired = GetEnvVarTrimmed(kSelfTestEnvDirectFinalAbortFailFired) == L"1";
    if (copyHr != HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) || ! abortFired || ! retainedPartial || ! receiptMatches)
    {
        Fail(std::format(L"R0a failed exact abort must retain partial content with RetainedIncomplete truth "
                         L"(hr=0x{0:08X}, fired={1}, retained={2}, receipt={3}, ec={4}).",
                         static_cast<unsigned long>(copyHr),
                         abortFired,
                         retainedPartial,
                         receiptMatches,
                         ec.value()));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.R0aLocalDirectFinalRetainedIncomplete", L"known-abort-failure", 1u, kSourceBytes, 0u, copyHr);
    NextStep(state, SelfTestState::Step::Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker);
    return false;
}
case SelfTestState::Step::Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker:
{
#if ! defined(_DEBUG)
    RecordCurrentPhase(state,
                       SelfTest::SelfTestCaseResult::Status::skipped,
                       L"FileSystem shared-scheduler shutdown selftests use a Debug-only plugin export; Debug coverage owns this proof.");
    NextStep(state, SelfTestState::Step::Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker);
    return false;
#else
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
#endif
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
    if (! RunFileOpsPerItemSchedulerFailurePolicySelfTestForSelfTest(*state.fileOps))
    {
        Fail(L"Riptide host scheduler failure-policy test observed a record-only worker failure cancel the owning task.");
        return true;
    }
    if (! RunFileOpsWorkerStartGateCancellationSelfTestForSelfTest(*state.fileOps))
    {
        Fail(L"Riptide worker start-gate test did not wake promptly for Cancel and stop requests.");
        return true;
    }
    if (! RunFileOpsBridgeDirectoryBufferValidationSelfTestForSelfTest())
    {
        Fail(L"Riptide bridge directory buffer validation selftest failed.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.RiptideHostSchedulerShutdownQuietPoint", L"", 1u, 0u, 0u, S_OK);
    NextStep(state, SelfTestState::Step::Floodgate_InlineF2WorkerQueueBypassAndInterlock);
    return false;
}
case SelfTestState::Step::Floodgate_InlineF2WorkerQueueBypassAndInterlock:
{
    using Task = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    const std::filesystem::path root = state.tempRoot / L"floodgate-inline-f2-interlock";
    const std::filesystem::path firstSource = root / L"first.txt";
    const std::filesystem::path firstDestination = root / L"first-renamed.txt";
    const std::filesystem::path secondSource = root / L"second.txt";
    const std::filesystem::path secondDestination = root / L"second-renamed.txt";
    const std::filesystem::path cancelOwnerSource = root / L"cancel-owner.txt";
    const std::filesystem::path cancelOwnerDestination = root / L"cancel-owner-renamed.txt";
    const std::filesystem::path cancelledWaiterSource = root / L"cancelled-waiter.txt";
    const std::filesystem::path cancelledWaiterDestination = root / L"cancelled-waiter-renamed.txt";
    const std::filesystem::path silentSource = root / L"silent-clean.txt";
    const std::filesystem::path silentDestination = root / L"silent-clean-renamed.txt";
    const std::filesystem::path conflictSource = root / L"conflict-source.txt";
    const std::filesystem::path conflictDestination = root / L"conflict-destination.txt";
    const wil::com_ptr<IFileSystem> paneFileSystem =
        state.folderWindow ? state.folderWindow->GetFileSystem(FolderWindow::Pane::Left) : nullptr;

    const auto restoreHarness = [&]() noexcept
    {
        ReleaseFileOpsInlineRenameBeforeMutationPauseForSelfTest();
        SetFileOpsInlineRenameBeforeMutationPauseForSelfTest(false);
        if (state.fileOps)
        {
            state.fileOps->SetQueueNewTasks(state.inlineF2QueueModeOriginal);
            state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(std::nullopt);
        }
    };
    const auto failAndRestore = [&](std::wstring message) noexcept
    {
        restoreHarness();
        Fail(std::move(message));
        return true;
    };
    const auto startRename = [&](const std::filesystem::path& source,
                                 std::wstring_view finalLeaf,
                                 std::optional<uint64_t>& taskId) noexcept -> HRESULT
    {
        uint64_t id = 0u;
        const HRESULT hr = state.fileOps->AdmitInlineRename(
            FolderWindow::Pane::Left, paneFileSystem, source, std::wstring(finalLeaf), &id);
        if (SUCCEEDED(hr) && id != 0u)
        {
            taskId = id;
        }
        return hr;
    };

    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        return failAndRestore(L"Floodgate_InlineF2WorkerQueueBypassAndInterlock timed out.");
    }
    if (! state.fileOps || ! paneFileSystem)
    {
        return failAndRestore(L"Floodgate inline-F2 test has no local File Operations endpoint.");
    }

    if (state.stepState == 0u)
    {
        if (! RecreateEmptyDirectory(root) || ! WriteFilledTestFile(firstSource, 32u, 0x11) ||
            ! WriteFilledTestFile(secondSource, 32u, 0x22) || ! WriteFilledTestFile(cancelOwnerSource, 32u, 0x33) ||
            ! WriteFilledTestFile(cancelledWaiterSource, 32u, 0x44) || ! WriteFilledTestFile(silentSource, 32u, 0x55) ||
            ! WriteFilledTestFile(conflictSource, 32u, 0x66) || ! WriteFilledTestFile(conflictDestination, 32u, 0x77))
        {
            return failAndRestore(L"Floodgate inline-F2 test failed to seed its local files.");
        }

        state.inlineF2QueueModeOriginal = state.fileOps->GetQueueNewTasks();
        state.inlineF2PresentedBaseline = state.fileOps->DebugTaskPresentedCount();
        state.inlineF2SilentCleanBaseline = state.fileOps->DebugInlineRenameSilentCleanCompletionCount();
        state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(1'000u);
        state.fileOps->SetQueueNewTasks(true);
        static_cast<void>(TakeFileOpsInlineRenameAdmissionThreadIdForSelfTest());
        static_cast<void>(TakeFileOpsInlineRenameExecutionThreadIdForSelfTest());
        static_cast<void>(TakeFileOpsInlineRenameExecutionAttemptsForSelfTest());
        SetFileOpsInlineRenameBeforeMutationPauseForSelfTest(true);

        const HRESULT startHr = startRename(firstSource, firstDestination.filename().native(), state.taskA);
        if (FAILED(startHr) || ! state.taskA.has_value())
        {
            return failAndRestore(std::format(L"Floodgate inline-F2 test could not admit the first rename: 0x{:08X}.",
                                              static_cast<unsigned long>(startHr)));
        }
        state.stepState = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        if (! HasFileOpsInlineRenameBeforeMutationPauseEnteredForSelfTest())
        {
            return false;
        }
        const DWORD admissionThreadId = TakeFileOpsInlineRenameAdmissionThreadIdForSelfTest();
        const DWORD executionThreadId = TakeFileOpsInlineRenameExecutionThreadIdForSelfTest();
        if (admissionThreadId == 0u || executionThreadId == 0u || admissionThreadId == executionThreadId)
        {
            return failAndRestore(std::format(L"Floodgate inline-F2 expected distinct admission/worker threads, got admission={} execution={}.",
                                              admissionThreadId,
                                              executionThreadId));
        }

        Task* const owner = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (! owner || owner->_mutationInterlockScopes.empty())
        {
            return failAndRestore(L"Floodgate inline rename reached its mutation pause without retained interlock authority.");
        }
        uint64_t retainedAuthorityNodes = 0u;
        for (const FileOperations::MutationInterlockScope& scope : owner->_mutationInterlockScopes)
        {
            if (! scope.root)
            {
                return failAndRestore(L"Floodgate inline rename retained an interlock scope without an exact bound authority chain.");
            }
            for (std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode> node = scope.root; node; node = node->parent)
            {
                ++retainedAuthorityNodes;
                if (! node->retained.authority.boundObject || retainedAuthorityNodes > 1024u)
                {
                    return failAndRestore(
                        L"Floodgate inline rename retained an invalid or unbounded interlock authority chain before mutation.");
                }
            }
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.InlineRenameRetainedAuthorityNodes",
                          L"paused-before-mutation",
                          retainedAuthorityNodes,
                          owner->_mutationInterlockScopes.size(),
                          1024u,
                          S_OK);

        if (! owner->IsPresentationHidden())
        {
            return failAndRestore(L"Floodgate inline rename was visible before its clean-running reveal deadline.");
        }
        state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(1'499u);
        static_cast<void>(state.fileOps->RefreshDeferredTaskPresentation());
        if (! owner->IsPresentationHidden())
        {
            return failAndRestore(L"Floodgate inline rename revealed before the fixed 500-ms deadline.");
        }
        state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(1'500u);
        static_cast<void>(state.fileOps->RefreshDeferredTaskPresentation());
        if (! owner->IsPresentationVisible() ||
            state.fileOps->DebugTaskPresentedCount() != state.inlineF2PresentedBaseline + 1u)
        {
            return failAndRestore(L"Floodgate inline rename did not reveal exactly at the fake-clock 500-ms deadline.");
        }

        const HRESULT startHr = startRename(secondSource, secondDestination.filename().native(), state.taskB);
        if (FAILED(startHr) || ! state.taskB.has_value())
        {
            return failAndRestore(std::format(L"Floodgate inline-F2 test could not admit the overlapping rename: 0x{:08X}.",
                                              static_cast<unsigned long>(startHr)));
        }
        state.stepState = 2u;
        return false;
    }

    if (state.stepState == 2u)
    {
        Task* const waiter = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
        if (! waiter)
        {
            return failAndRestore(L"Floodgate inline-F2 overlapping waiter disappeared before interlock validation.");
        }
        if (! waiter->IsWaitingInQueue())
        {
            return false;
        }
        if (waiter->IsWaitingForOthers())
        {
            return failAndRestore(L"Floodgate inline-F2 inherited global Queue waiting instead of only the overlapping-root interlock.");
        }
        if (! waiter->IsPresentationVisible())
        {
            return failAndRestore(L"Floodgate inline-F2 interlock waiter did not reveal immediately.");
        }
        const unsigned long attempts = TakeFileOpsInlineRenameExecutionAttemptsForSelfTest();
        if (attempts != 1u)
        {
            return failAndRestore(std::format(L"Floodgate inline-F2 expected only the interlock owner to reach execution; saw {} attempts.", attempts));
        }

        ReleaseFileOpsInlineRenameBeforeMutationPauseForSelfTest();
        SetFileOpsInlineRenameBeforeMutationPauseForSelfTest(false);
        state.stepState = 3u;
        return false;
    }

    if (state.stepState == 3u)
    {
        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            return failAndRestore(L"Floodgate inline-F2 success round lost its task ids.");
        }
        const auto firstCompleted = state.completedTasks.find(state.taskA.value());
        const auto secondCompleted = state.completedTasks.find(state.taskB.value());
        if (firstCompleted == state.completedTasks.end() || secondCompleted == state.completedTasks.end())
        {
            return false;
        }

        std::error_code ec;
        if (FAILED(firstCompleted->second.hr) || FAILED(secondCompleted->second.hr) ||
            ! std::filesystem::exists(firstDestination, ec) || ec || ! std::filesystem::exists(secondDestination, ec) || ec ||
            std::filesystem::exists(firstSource, ec) || ec || std::filesystem::exists(secondSource, ec) || ec)
        {
            return failAndRestore(std::format(L"Floodgate inline-F2 serialized success round failed (first=0x{:08X}, second=0x{:08X}).",
                                              static_cast<unsigned long>(firstCompleted->second.hr),
                                              static_cast<unsigned long>(secondCompleted->second.hr)));
        }
        if (secondCompleted->second.interlockWaitCount == 0u || secondCompleted->second.interlockWaitUs == 0u)
        {
            return failAndRestore(L"Floodgate inline-F2 overlapping rename completed without a recorded interlock wait.");
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.InlineF2InterlockWaitUs",
                          L"queue-bypass-overlapping-parent",
                          secondCompleted->second.interlockWaitUs,
                          secondCompleted->second.interlockWaitCount,
                          2u,
                          S_OK);

        state.taskA.reset();
        state.taskB.reset();
        static_cast<void>(TakeFileOpsInlineRenameAdmissionThreadIdForSelfTest());
        static_cast<void>(TakeFileOpsInlineRenameExecutionThreadIdForSelfTest());
        static_cast<void>(TakeFileOpsInlineRenameExecutionAttemptsForSelfTest());
        SetFileOpsInlineRenameBeforeMutationPauseForSelfTest(true);
        const HRESULT startHr = startRename(cancelOwnerSource, cancelOwnerDestination.filename().native(), state.taskA);
        if (FAILED(startHr) || ! state.taskA.has_value())
        {
            return failAndRestore(std::format(L"Floodgate inline-F2 cancellation round could not admit its owner: 0x{0:08X}.",
                                              static_cast<unsigned long>(startHr)));
        }
        state.stepState = 4u;
        return false;
    }

    if (state.stepState == 4u)
    {
        if (! HasFileOpsInlineRenameBeforeMutationPauseEnteredForSelfTest())
        {
            return false;
        }
        const HRESULT startHr = startRename(cancelledWaiterSource, cancelledWaiterDestination.filename().native(), state.taskB);
        if (FAILED(startHr) || ! state.taskB.has_value())
        {
            return failAndRestore(std::format(L"Floodgate inline-F2 cancellation round could not admit its waiter: 0x{0:08X}.",
                                              static_cast<unsigned long>(startHr)));
        }
        state.stepState = 5u;
        return false;
    }

    if (state.stepState == 5u)
    {
        Task* const waiter = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
        if (! waiter)
        {
            return failAndRestore(L"Floodgate inline-F2 cancellation waiter disappeared before it could be canceled.");
        }
        if (! waiter->IsWaitingInQueue())
        {
            return false;
        }
        if (waiter->IsWaitingForOthers())
        {
            return failAndRestore(L"Floodgate inline-F2 cancellation waiter inherited global Queue mode.");
        }
        const unsigned long attempts = TakeFileOpsInlineRenameExecutionAttemptsForSelfTest();
        if (attempts != 1u)
        {
            return failAndRestore(
                std::format(L"Floodgate inline-F2 cancellation round expected one execution attempt before cancel; saw {0}.", attempts));
        }

        waiter->RequestCancel();
        ReleaseFileOpsInlineRenameBeforeMutationPauseForSelfTest();
        SetFileOpsInlineRenameBeforeMutationPauseForSelfTest(false);
        state.stepState = 6u;
        return false;
    }

    if (state.stepState == 6u)
    {
        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            return failAndRestore(L"Floodgate inline-F2 cancellation round lost its task ids.");
        }
        const auto ownerCompleted = state.completedTasks.find(state.taskA.value());
        const auto waiterCompleted = state.completedTasks.find(state.taskB.value());
        if (ownerCompleted == state.completedTasks.end() || waiterCompleted == state.completedTasks.end())
        {
            return false;
        }

        std::error_code ec;
        const unsigned long postCancelAttempts = TakeFileOpsInlineRenameExecutionAttemptsForSelfTest();
        if (FAILED(ownerCompleted->second.hr) || waiterCompleted->second.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) ||
            ! std::filesystem::exists(cancelOwnerDestination, ec) || ec || std::filesystem::exists(cancelOwnerSource, ec) || ec ||
            ! std::filesystem::exists(cancelledWaiterSource, ec) || ec || std::filesystem::exists(cancelledWaiterDestination, ec) || ec ||
            postCancelAttempts != 0u || waiterCompleted->second.interlockWaitCount == 0u)
        {
            return failAndRestore(std::format(L"Floodgate inline-F2 cancellation round failed (owner=0x{0:08X}, waiter=0x{1:08X}, postCancelAttempts={2}).",
                                              static_cast<unsigned long>(ownerCompleted->second.hr),
                                              static_cast<unsigned long>(waiterCompleted->second.hr),
                                              postCancelAttempts));
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.InlineF2CancelledInterlockWaitUs",
                          L"cancelled-before-mutation",
                          waiterCompleted->second.interlockWaitUs,
                          waiterCompleted->second.interlockWaitCount,
                          postCancelAttempts,
                          waiterCompleted->second.hr);
        state.taskA.reset();
        state.taskB.reset();
        state.inlineF2SilentCleanBaseline = state.fileOps->DebugInlineRenameSilentCleanCompletionCount();
        state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(5'000u);
        const HRESULT startHr = startRename(silentSource, silentDestination.filename().native(), state.taskA);
        if (FAILED(startHr) || ! state.taskA.has_value())
        {
            return failAndRestore(std::format(L"Floodgate inline-F2 silent-clean round could not admit its rename: 0x{0:08X}.",
                                              static_cast<unsigned long>(startHr)));
        }
        state.stepState = 7u;
        return false;
    }

    if (state.stepState == 7u)
    {
        if (! state.taskA.has_value())
        {
            return failAndRestore(L"Floodgate inline-F2 silent-clean round lost its task id.");
        }
        const auto silentCompleted = state.completedTasks.find(state.taskA.value());
        if (silentCompleted == state.completedTasks.end())
        {
            return false;
        }

        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        const bool retainedCard = std::ranges::any_of(summaries, [&](const auto& summary) noexcept
        { return summary.taskId == state.taskA.value(); });
        std::error_code silentEc;
        if (FAILED(silentCompleted->second.hr) || retainedCard ||
            state.fileOps->DebugInlineRenameSilentCleanCompletionCount() != state.inlineF2SilentCleanBaseline + 1u ||
            ! std::filesystem::exists(silentDestination, silentEc) || silentEc ||
            std::filesystem::exists(silentSource, silentEc) || silentEc)
        {
            return failAndRestore(L"Floodgate inline-F2 clean completion did not stay silent or retain truthful rename output.");
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.InlineRenameSilentCleanCompletion",
                          L"fake-clock-before-deadline",
                          FileOperations::kInlineRenameCardRevealDelayMs,
                          retainedCard ? 1u : 0u,
                          state.fileOps->DebugInlineRenameSilentCleanCompletionCount(),
                          silentCompleted->second.hr);
        state.taskA.reset();
        state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(10'000u);
        const HRESULT conflictStartHr = startRename(conflictSource, conflictDestination.filename().native(), state.taskA);
        if (FAILED(conflictStartHr) || ! state.taskA.has_value())
        {
            return failAndRestore(std::format(L"Floodgate inline-F2 conflict round could not admit its rename: 0x{0:08X}.",
                                              static_cast<unsigned long>(conflictStartHr)));
        }
        state.stepState = 8u;
        return false;
    }

    if (state.stepState == 8u)
    {
        Task* const conflictTask = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (! conflictTask)
        {
            return failAndRestore(L"Floodgate inline-F2 conflict task disappeared before the prompt was observed.");
        }
        bool hasOverwriteAction = false;
        {
            std::scoped_lock lock(conflictTask->_conflictArbiter.mutex);
            if (! conflictTask->_conflictArbiter.prompt.active || conflictTask->_conflictArbiter.prompt.metadataLoading)
            {
                return false;
            }
            hasOverwriteAction = PromptHasAction(conflictTask->_conflictArbiter.prompt, Task::ConflictAction::Overwrite);
        }
        if (! hasOverwriteAction)
        {
            return failAndRestore(L"Floodgate inline-F2 exact-destination conflict did not publish a consumable Overwrite action.");
        }
        if (! conflictTask->IsPresentationVisible())
        {
            return failAndRestore(L"Floodgate inline-F2 conflict did not reveal before the 500-ms running deadline.");
        }
        conflictTask->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
        state.stepState = 9u;
        return false;
    }

    if (! state.taskA.has_value())
    {
        return failAndRestore(L"Floodgate inline-F2 conflict completion lost its task id.");
    }
    if (Task* const conflictTask = state.fileOps->FindTask(state.taskA.value()))
    {
        const auto repeatedPrompt = TryGetConflictPromptCopy(conflictTask);
        if (repeatedPrompt.has_value())
        {
            if (repeatedPrompt->metadataLoading)
            {
                return false;
            }
            return failAndRestore(std::format(
                L"Floodgate inline-F2 Overwrite receipt did not make the conditional rename terminal "
                L"(bucket={0}, status=0x{1:08X}, actions={2}).",
                static_cast<unsigned int>(repeatedPrompt->bucket),
                static_cast<unsigned long>(repeatedPrompt->status),
                repeatedPrompt->actionCount));
        }
    }
    const auto conflictCompleted = state.completedTasks.find(state.taskA.value());
    if (conflictCompleted == state.completedTasks.end())
    {
        return false;
    }
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> conflictSummaries;
    state.fileOps->CollectCompletedTasks(conflictSummaries);
    const bool conflictCardRetained = std::ranges::any_of(conflictSummaries, [&](const auto& summary) noexcept
    { return summary.taskId == state.taskA.value(); });
    std::error_code conflictEc;
    if (FAILED(conflictCompleted->second.hr) || ! conflictCardRetained ||
        ! std::filesystem::exists(conflictDestination, conflictEc) || conflictEc ||
        std::filesystem::exists(conflictSource, conflictEc) || conflictEc)
    {
        return failAndRestore(L"Floodgate inline-F2 exact-destination Overwrite did not retain its revealed card and committed result.");
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.InlineRenameConflictReveal",
                      L"exact-destination-overwrite",
                      FileOperations::kInlineRenameCardRevealDelayMs,
                      conflictCardRetained ? 1u : 0u,
                      1u,
                      conflictCompleted->second.hr);
    restoreHarness();
    state.taskA.reset();
    NextStep(state, SelfTestState::Step::C1_InlineRenameNameRefusedOnCard);
    return false;
}

case SelfTestState::Step::C1_InlineRenameNameRefusedOnCard:
{
    // C1 (b0): the inline F2 admission no longer queries the provider's child-name contract on
    // the UI thread. A leaf the Local provider cannot hold (trailing dot) is admitted and
    // published; Preparing fails the task with the provider's reason before any mutation, and
    // the source keeps its name. On the baseline the admission refused synchronously and no task
    // existed.
    const std::filesystem::path root   = state.tempRoot / L"c1-inline-rename-refused";
    const std::filesystem::path source = root / L"before.txt";
    const wil::com_ptr<IFileSystem> paneFileSystem =
        state.folderWindow ? state.folderWindow->GetFileSystem(FolderWindow::Pane::Left) : nullptr;
    if (HasTimedOut(state, GetTickCount64(), 60'000ull))
    {
        Fail(std::format(L"C1 inline rename refused timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        if (! state.fileOps || ! paneFileSystem)
        {
            Fail(L"C1 inline rename refused requires the left pane file system.");
            return true;
        }
        if (! RecreateEmptyDirectory(root) || ! WriteFilledTestFile(source, 1024u, 0x5A))
        {
            Fail(L"C1 inline rename refused could not stage its source.");
            return true;
        }
        uint64_t taskId  = 0u;
        const HRESULT hr = state.fileOps->AdmitInlineRename(FolderWindow::Pane::Left, paneFileSystem, source, std::wstring(L"after."), &taskId);
        if (FAILED(hr) || taskId == 0u)
        {
            Fail(std::format(L"C1 inline rename refused: admission must publish the task and leave the name to Preparing (hr=0x{:08X}, task={}).",
                             static_cast<unsigned long>(hr),
                             taskId));
            return true;
        }
        state.taskA     = taskId;
        state.stepState = 1u;
        return false;
    }
    const auto done = state.completedTasks.find(state.taskA.value());
    if (done == state.completedTasks.end())
    {
        return false;
    }
    std::error_code ec;
    if (done->second.hr != HRESULT_FROM_WIN32(ERROR_INVALID_NAME) || ! std::filesystem::exists(source, ec) ||
        std::filesystem::exists(root / L"after.", ec) || std::filesystem::exists(root / L"after", ec))
    {
        Fail(std::format(L"C1 inline rename refused: Preparing must fail the task with ERROR_INVALID_NAME and leave the source untouched (hr=0x{:08X}).",
                         static_cast<unsigned long>(done->second.hr)));
        return true;
    }
    static_cast<void>(SelfTest::RemoveAll(root));
    state.taskA.reset();
    NextStep(state, SelfTestState::Step::Phase5_DiscoverySingleTraversal);
    return false;
}

#endif
