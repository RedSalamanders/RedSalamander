#if defined(FILEOPS_SELFTEST_INCLUDE_COPY_MERGE)

case SelfTestState::Step::FileOps_CopyMergeIntoExistingFolder:
{
    using Task              = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"FileOps_CopyMergeIntoExistingFolder timed out.");
        return true;
    }

    const std::filesystem::path srcRoot     = state.tempRoot / L"clearflow-copy-merge-src";
    const std::filesystem::path dstRoot     = state.tempRoot / L"clearflow-copy-merge-dst";
    const std::filesystem::path srcFoo      = srcRoot / L"Foo";
    const std::filesystem::path dstFoo      = dstRoot / L"Foo";
    const std::filesystem::path srcConflict = srcFoo / L"a.bin";
    const std::filesystem::path dstConflict = dstFoo / L"a.bin";
    const std::filesystem::path srcNested   = srcFoo / L"nested" / L"c.bin";
    const std::filesystem::path dstNested   = dstFoo / L"nested" / L"c.bin";
    const std::filesystem::path dstKeep     = dstFoo / L"keep.bin";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Failed to reset copy-merge directories.");
            return true;
        }

        if (! WriteTestFile(srcConflict, 16 * 1024) || ! WriteTestFile(srcNested, 8 * 1024) || ! WriteTestFile(dstConflict, 1024) ||
            ! WriteTestFile(dstKeep, 2 * 1024))
        {
            Fail(L"Failed to seed copy-merge test tree.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
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
            Fail(L"Failed to start copy-merge task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        Task* task        = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }

        if (prompt->bucket != Task::ConflictBucket::RegularFileExists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
        {
            Fail(L"Copy-merge expected an Exists prompt with Overwrite for the colliding file.");
            return true;
        }

        const std::wstring expectedSource = NormalizePathForCompare(srcConflict.wstring());
        const std::wstring expectedDest   = NormalizePathForCompare(dstConflict.wstring());
        const std::wstring actualSource   = NormalizePathForCompare(prompt->sourcePath);
        const std::wstring actualDest     = NormalizePathForCompare(prompt->destinationPath);
        if (actualSource != expectedSource || actualDest != expectedDest)
        {
            Fail(std::format(
                L"Copy-merge prompted for the wrong item. expected='{}' -> '{}' actual='{}' -> '{}'.", expectedSource, expectedDest, actualSource, actualDest));
            return true;
        }

        const HWND popup = FindCurrentProcessPopupWindow();
        if (! popup)
        {
            return false;
        }

        FileOperationsPopupInternal::PopupLayoutDebugSnapshot conflictLayout{};
        conflictLayout.taskId = state.taskA.value();
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, conflictLayout) || ! conflictLayout.found)
        {
            return false;
        }
        if (! conflictLayout.conflictWaitingForDecisionVisible || ! conflictLayout.conflictPromptPrecedesActions ||
            ! conflictLayout.conflictPromptVisibleWithoutClipping || ! conflictLayout.conflictContextVisibleWithoutClipping ||
            conflictLayout.conflictDecisionDetailsLoadingVisible || conflictLayout.conflictDiscoveryIndicatorVisible ||
            conflictLayout.conflictTransferProgressVisible || ! conflictLayout.lastNoteVisibleWithoutClipping)
        {
            Fail(std::format(
                L"Actionable conflict card violated its decision-state contract (waiting={0}, promptBeforeActions={1}, promptFits={2}, contextFits={3}, loading={4}, discovery={5}, transfer={6}, lastNoteFits={7}).",
                conflictLayout.conflictWaitingForDecisionVisible,
                conflictLayout.conflictPromptPrecedesActions,
                conflictLayout.conflictPromptVisibleWithoutClipping,
                conflictLayout.conflictContextVisibleWithoutClipping,
                conflictLayout.conflictDecisionDetailsLoadingVisible,
                conflictLayout.conflictDiscoveryIndicatorVisible,
                conflictLayout.conflictTransferProgressVisible,
                conflictLayout.lastNoteVisibleWithoutClipping));
            return true;
        }

        task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
        state.markerTick = nowTick;
        state.stepState  = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                const std::wstring actualSource = NormalizePathForCompare(prompt->sourcePath);
                if ((nowTick - state.markerTick) < 500ull && actualSource == NormalizePathForCompare(srcConflict.wstring()))
                {
                    return false;
                }

                Fail(std::format(L"Copy-merge raised an unexpected second prompt for '{}'.", actualSource));
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
            Fail(std::format(L"Copy-merge task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        if (it->second.conflictPromptCount != 1u)
        {
            Fail(std::format(L"Copy-merge expected exactly one file conflict prompt, saw {}.", it->second.conflictPromptCount));
            return true;
        }

        if (! FilesEqualBytes(srcConflict, dstConflict) || ! FilesEqualBytes(srcNested, dstNested) || ! FileSizeEquals(dstKeep, 2 * 1024))
        {
            Fail(L"Copy-merge destination tree failed byte-for-byte integrity checks.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.CopyMergeIntoExistingFolder.PromptCount", L"", 0u, it->second.conflictPromptCount, 0u, S_OK);
        NextStep(state, SelfTestState::Step::FileOps_MoveMergeIntoExistingFolderSameVolume);
        return false;
    }

    return false;
}
#endif
#if defined(FILEOPS_SELFTEST_INCLUDE_MOVE_MERGE)

case SelfTestState::Step::FileOps_MoveMergeIntoExistingFolderSameVolume:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        AppendLog(std::format(L"Move-merge timeout at substep {} with race attempts {}.",
                              state.stepState,
                              TakeFileOpsNativeMoveCreateDirectoryRaceAttemptsForSelfTest()));
        const auto logTimedOutTask = [&](std::wstring_view label, const std::optional<uint64_t>& taskId) noexcept
        {
            const auto* task = state.fileOps != nullptr && taskId.has_value() ? state.fileOps->FindTask(taskId.value()) : nullptr;
            AppendLog(std::format(L"Move-merge timeout {}: present={} started={} entered={} waiting={} result=0x{:08X}",
                                  label,
                                  task != nullptr,
                                  task != nullptr && task->HasStarted(),
                                  task != nullptr && task->HasEnteredOperation(),
                                  task != nullptr && task->IsWaitingInQueue(),
                                  static_cast<unsigned long>(task != nullptr ? task->GetResult() : E_UNEXPECTED)));
        };
        logTimedOutTask(L"merge", state.taskB);
        logTimedOutTask(L"native", state.taskC);
        logTimedOutTask(L"active", state.taskA);
        SetFileOpsNativeMoveCreateDirectoryRaceForSelfTest(0u);
        Fail(L"FileOps_MoveMergeIntoExistingFolderSameVolume timed out.");
        return true;
    }

    const std::filesystem::path srcRoot   = state.tempRoot / L"clearflow-move-merge-src";
    const std::filesystem::path dstRoot   = state.tempRoot / L"clearflow-move-merge-dst";
    const std::filesystem::path srcFoo    = srcRoot / L"Foo";
    const std::filesystem::path dstFoo    = dstRoot / L"Foo";
    const std::filesystem::path srcFile   = srcFoo / L"new.bin";
    const std::filesystem::path dstFile   = dstFoo / L"new.bin";
    const std::filesystem::path srcNested = srcFoo / L"nested" / L"child.bin";
    const std::filesystem::path dstNested = dstFoo / L"nested" / L"child.bin";
    const std::filesystem::path dstKeep   = dstFoo / L"keep.bin";
    const std::filesystem::path srcConflict = srcFoo / L"conflict.bin";
    const std::filesystem::path dstConflict = dstFoo / L"conflict.bin";
    const std::filesystem::path nativeSrcRoot = state.tempRoot / L"clearflow-native-file-src";
    const std::filesystem::path nativeDstRoot = state.tempRoot / L"clearflow-native-file-dst";
    const std::filesystem::path nativeSource  = nativeSrcRoot / L"native.bin";
    const std::filesystem::path nativeTarget  = nativeDstRoot / nativeSource.filename();
    const std::filesystem::path raceSrcRoot   = state.tempRoot / L"clearflow-native-race-src";
    const std::filesystem::path raceDstRoot   = state.tempRoot / L"clearflow-native-race-dst";
    const std::filesystem::path raceSource    = raceSrcRoot / L"RacedFolder";
    const std::filesystem::path raceTarget    = raceDstRoot / raceSource.filename();
    const std::filesystem::path raceSourceFile = raceSource / L"source.bin";
    const std::filesystem::path raceTargetFile = raceTarget / raceSourceFile.filename();
    const std::filesystem::path mixedSrcRoot   = state.tempRoot / L"clearflow-mixed-strategy-src";
    const std::filesystem::path mixedDstRoot   = state.tempRoot / L"clearflow-mixed-strategy-dst";
    const std::filesystem::path mixedMergeSource = mixedSrcRoot / L"MergeFolder";
    const std::filesystem::path mixedNativeSource = mixedSrcRoot / L"native.bin";
    const std::filesystem::path mixedMergeTarget = mixedDstRoot / mixedMergeSource.filename();
    const std::filesystem::path mixedNativeTarget = mixedDstRoot / mixedNativeSource.filename();
    const std::filesystem::path atomicNativeSrcRoot = state.tempRoot / L"clearflow-atomic-native-src";
    const std::filesystem::path atomicNativeDstRoot = state.tempRoot / L"clearflow-atomic-native-dst";
    const std::filesystem::path atomicNativeSource = atomicNativeSrcRoot / L"Folder";
    const std::filesystem::path atomicNativeTarget = atomicNativeDstRoot / atomicNativeSource.filename();
    const std::filesystem::path atomicNativeSourceFile = atomicNativeSource / L"source.bin";
    const std::filesystem::path atomicNativeTargetFile = atomicNativeTarget / L"keep.bin";
    const std::filesystem::path caseOnlyRoot = state.tempRoot / L"clearflow-case-only-directory";
    const std::filesystem::path caseOnlySource = caseOnlyRoot / L"CaseOnlyFolder";
    const std::filesystem::path caseOnlyDestination = caseOnlyRoot / L"caseonlyfolder";
    const std::filesystem::path caseOnlyChild = caseOnlyDestination / L"child.bin";
    const std::filesystem::path caseEntropySource = caseOnlyRoot / L"CaseEntropyFolder";
    const std::filesystem::path caseEntropyDestination = caseOnlyRoot / L"caseentropyfolder";
    const std::filesystem::path caseCollisionSource = caseOnlyRoot / L"CaseCollisionFolder";
    const std::filesystem::path caseCollisionDestination = caseOnlyRoot / L"casecollisionfolder";
    const std::filesystem::path crossVolumeSource = state.tempRoot / L"clearflow-cross-volume-managed.bin";

    if (state.stepState == 0)
    {
        AppendLog(L"Move-merge qualification: seeding fixtures");
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) ||
            ! RecreateEmptyDirectory(nativeSrcRoot) || ! RecreateEmptyDirectory(nativeDstRoot) ||
            ! RecreateEmptyDirectory(raceSrcRoot) || ! RecreateEmptyDirectory(raceDstRoot) ||
            ! RecreateEmptyDirectory(mixedSrcRoot) || ! RecreateEmptyDirectory(mixedDstRoot) ||
            ! RecreateEmptyDirectory(atomicNativeSrcRoot) || ! RecreateEmptyDirectory(atomicNativeDstRoot) ||
            ! RecreateEmptyDirectory(caseOnlyRoot))
        {
            Fail(L"Failed to reset move-merge directories.");
            return true;
        }

        if (! WriteTestFile(srcFile, 4 * 1024) || ! WriteTestFile(srcNested, 6 * 1024) || ! WriteTestFile(dstKeep, 3 * 1024) ||
            ! WriteTestFile(srcConflict, 5 * 1024) || ! WriteTestFile(dstConflict, 1024) ||
            ! WriteTestFile(nativeSource, 2 * 1024) || ! WriteTestFile(raceSourceFile, 9 * 1024) ||
            ! WriteTestFile(mixedMergeSource / L"merge.bin", 5 * 1024) ||
            ! WriteTestFile(mixedNativeSource, 6 * 1024) ||
            ! WriteTestFile(mixedMergeTarget / L"keep.bin", 2 * 1024) ||
            ! WriteTestFile(atomicNativeSourceFile, 3 * 1024) || ! WriteTestFile(atomicNativeTargetFile, 2 * 1024) ||
            ! WriteTestFile(caseOnlySource / L"child.bin", 1536) ||
            ! WriteTestFile(caseEntropySource / L"child.bin", 1024) ||
            ! WriteTestFile(caseCollisionSource / L"child.bin", 2048))
        {
            Fail(L"Failed to seed move-merge test tree.");
            return true;
        }

        const HRESULT caseOnlyDirectoryHr = state.fsLocal->RenameItem(caseOnlySource.c_str(),
                                                                       caseOnlyDestination.c_str(),
                                                                       FILESYSTEM_FLAG_NONE,
                                                                       nullptr,
                                                                       nullptr,
                                                                       nullptr);
        bool sawOriginalCase = false;
        bool sawDestinationCase = false;
        std::error_code caseOnlyEc;
        for (std::filesystem::directory_iterator it(caseOnlyRoot, caseOnlyEc), end;
             ! caseOnlyEc && it != end;
             it.increment(caseOnlyEc))
        {
            const std::wstring leaf = it->path().filename().native();
            sawOriginalCase |= leaf == L"CaseOnlyFolder";
            sawDestinationCase |= leaf == L"caseonlyfolder";
        }
        if (FAILED(caseOnlyDirectoryHr) || caseOnlyEc || sawOriginalCase || ! sawDestinationCase ||
            ! FileSizeEquals(caseOnlyChild, 1536))
        {
            Fail(std::format(L"Local provider case-only directory Rename must preserve the existing temp-rename behavior (hr=0x{:08X}).",
                             static_cast<unsigned long>(caseOnlyDirectoryHr)));
            return true;
        }

        constexpr wchar_t kCaseEntropyFailureEnv[] = L"REDSALAMANDER_FILEOPS_CASE_RENAME_ENTROPY_FAIL_PATH";
        constexpr wchar_t kCaseCollisionPathEnv[] = L"REDSALAMANDER_FILEOPS_CASE_RENAME_TEMP_COLLISION_PATH";
        constexpr wchar_t kCaseCollisionFiredEnv[] = L"REDSALAMANDER_FILEOPS_CASE_RENAME_TEMP_COLLISION_FIRED";
        const auto clearCaseRenameHooks = wil::scope_exit([&]() noexcept
        {
            const std::wstring collisionTempPath = GetEnvVarTrimmed(kCaseCollisionFiredEnv);
            if (! collisionTempPath.empty())
            {
                const std::filesystem::path collisionPath(collisionTempPath);
                static_cast<void>(DeleteFileW((collisionPath / L"race.marker").c_str()));
                static_cast<void>(DeleteFileW(collisionPath.c_str()));
                static_cast<void>(RemoveDirectoryW(collisionPath.c_str()));
            }
            static_cast<void>(SetEnvironmentVariableW(kCaseEntropyFailureEnv, nullptr));
            static_cast<void>(SetEnvironmentVariableW(kCaseCollisionPathEnv, nullptr));
            static_cast<void>(SetEnvironmentVariableW(kCaseCollisionFiredEnv, nullptr));
        });
        const auto hasExactLeaf = [&](std::wstring_view expected) noexcept
        {
            std::error_code ec;
            for (std::filesystem::directory_iterator it(caseOnlyRoot, ec), end; ! ec && it != end; it.increment(ec))
            {
                if (it->path().filename().native() == expected)
                {
                    return true;
                }
            }
            return false;
        };

        if (SetEnvironmentVariableW(kCaseEntropyFailureEnv, caseEntropySource.c_str()) == FALSE)
        {
            Fail(L"Failed to arm case-only rename entropy-failure fixture.");
            return true;
        }
        const HRESULT entropyRenameHr = state.fsLocal->RenameItem(caseEntropySource.c_str(),
                                                                   caseEntropyDestination.c_str(),
                                                                   FILESYSTEM_FLAG_NONE,
                                                                   nullptr,
                                                                   nullptr,
                                                                   nullptr);
        if (entropyRenameHr != HRESULT_FROM_WIN32(ERROR_GEN_FAILURE) || ! hasExactLeaf(L"CaseEntropyFolder") ||
            hasExactLeaf(L"caseentropyfolder"))
        {
            Fail(std::format(L"Case-only rename entropy failure must occur before the first rename (hr=0x{0:08X}).",
                             static_cast<unsigned long>(entropyRenameHr)));
            return true;
        }

        if (SetEnvironmentVariableW(kCaseCollisionPathEnv, caseCollisionSource.c_str()) == FALSE)
        {
            Fail(L"Failed to arm case-only rename temporary-sibling collision fixture.");
            return true;
        }
        const HRESULT collisionRenameHr = state.fsLocal->RenameItem(caseCollisionSource.c_str(),
                                                                     caseCollisionDestination.c_str(),
                                                                     FILESYSTEM_FLAG_NONE,
                                                                     nullptr,
                                                                     nullptr,
                                                                     nullptr);
        const std::wstring collisionTempPath = GetEnvVarTrimmed(kCaseCollisionFiredEnv);
        const std::wstring collisionTempLeaf = std::filesystem::path(collisionTempPath).filename().native();
        bool leakedCaseTemp = false;
        std::error_code leakedTempEc;
        for (std::filesystem::directory_iterator it(caseOnlyRoot, leakedTempEc), end;
             ! leakedTempEc && it != end;
             it.increment(leakedTempEc))
        {
            const std::filesystem::path enumeratedPath = it->path();
            leakedCaseTemp |= enumeratedPath.filename().native().starts_with(L".rs_case_tmp_") &&
                              ! OrdinalString::EqualsNoCase(enumeratedPath.filename().native(), collisionTempLeaf);
        }
        if (FAILED(collisionRenameHr) || collisionTempPath.empty() ||
            ! FileSizeEquals(std::filesystem::path(collisionTempPath) / L"race.marker", 4u) ||
            ! hasExactLeaf(L"casecollisionfolder") || hasExactLeaf(L"CaseCollisionFolder") || leakedCaseTemp || leakedTempEc)
        {
            Fail(std::format(
                L"Case-only rename must recover from an exclusively claimed temporary-sibling collision without overwriting it "
                L"(hr=0x{0:08X}, temp='{1}', marker={2}, final={3}, original={4}, leaked={5}, enumError={6}).",
                static_cast<unsigned long>(collisionRenameHr),
                collisionTempPath,
                FileSizeEquals(std::filesystem::path(collisionTempPath) / L"race.marker", 4u) ? 1 : 0,
                hasExactLeaf(L"casecollisionfolder") ? 1 : 0,
                hasExactLeaf(L"CaseCollisionFolder") ? 1 : 0,
                leakedCaseTemp ? 1 : 0,
                leakedTempEc.value()));
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        FileSystemOptions nativeOptions{};
        nativeOptions.sizeBytes  = sizeof(nativeOptions);
        nativeOptions.moveMode   = FILESYSTEM_MOVE_NATIVE_ONLY;
        nativeOptions.linkPolicy = FILESYSTEM_LINK_PRESERVE;
        const HRESULT atomicNativeHr = state.fsLocal->MoveItem(
            atomicNativeSource.c_str(), atomicNativeTarget.c_str(), flags, &nativeOptions, nullptr, nullptr);
        if (atomicNativeHr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) ||
            ! FileSizeEquals(atomicNativeSourceFile, 3 * 1024) || ! FileSizeEquals(atomicNativeTargetFile, 2 * 1024) ||
            std::filesystem::exists(atomicNativeTarget / atomicNativeSourceFile.filename()))
        {
            Fail(std::format(L"Local Native directory Move must be one non-mutating provider operation when a regular destination directory exists (hr=0x{:08X}).",
                             static_cast<unsigned long>(atomicNativeHr)));
            return true;
        }

        AppendLog(L"Move-merge qualification: admitting Managed merge task");
        state.taskB                 = StartFileOperationAndGetId(state.fileOps,
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
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start same-volume move-merge task.");
            return true;
        }

        AppendLog(std::format(L"Move-merge qualification: Managed task admitted id={}", state.taskB.value()));
        AppendLog(L"Move-merge qualification: admitting destination-absent Native regular-file task");
        state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {nativeSource},
                                                 nativeDstRoot,
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskC.has_value())
        {
            Fail(L"Failed to start destination-absent same-volume native regular-file Move.");
            return true;
        }
        AppendLog(std::format(L"Move-merge qualification: Native task admitted id={}", state.taskC.value()));

        // C1: the shapes are Preparing facts now (admission no longer probes the provider); each
        // task's prepared strategy is checked once both have run.
        state.markerTick = nowTick;
        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        auto* task = state.fileOps && state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            if (state.taskB.has_value() && state.completedTasks.contains(state.taskB.value()))
            {
                Fail(L"Managed same-volume directory merge completed without exposing its colliding child through the central conflict surface.");
                return true;
            }
            if (task != nullptr && state.markerTick != 0u && (nowTick - state.markerTick) >= 1'000ull)
            {
                AppendLog(std::format(L"Move-merge qualification: waiting for conflict started={} entered={} waiting={} result=0x{:08X}",
                                      task->HasStarted(),
                                      task->HasEnteredOperation(),
                                      task->IsWaitingInQueue(),
                                      static_cast<unsigned long>(task->GetResult())));
                state.markerTick = 0u;
            }
            return false;
        }
        if (task == nullptr || prompt->bucket != FolderWindow::FileOperationState::Task::ConflictBucket::RegularFileExists ||
            ! PromptHasAction(prompt.value(), FolderWindow::FileOperationState::Task::ConflictAction::Overwrite) ||
            NormalizePathForCompare(prompt->sourcePath) != NormalizePathForCompare(srcConflict.native()) ||
            NormalizePathForCompare(prompt->destinationPath) != NormalizePathForCompare(dstConflict.native()))
        {
            Fail(std::format(L"Managed same-volume directory merge exposed the wrong conflict: '{}' -> '{}'.",
                             prompt->sourcePath,
                             prompt->destinationPath));
            return true;
        }
        task->SubmitConflictDecision(FolderWindow::FileOperationState::Task::ConflictAction::Overwrite, false);
        state.markerTick = nowTick;
        state.stepState  = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (auto* task = state.fileOps && state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if ((nowTick - state.markerTick) < 500ull && NormalizePathForCompare(prompt->sourcePath) == NormalizePathForCompare(srcConflict.native()))
                {
                    return false;
                }
                Fail(std::format(L"Managed same-volume directory merge raised an unexpected second conflict for '{}'.", prompt->sourcePath));
                return true;
            }
        }
        const auto it = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        const auto nativeIt = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end() || nativeIt == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(it->second.hr) || FAILED(nativeIt->second.hr))
        {
            Fail(std::format(L"Same-volume directory Move qualification tasks failed: merge=0x{:08X}, native=0x{:08X}.",
                             static_cast<unsigned long>(it->second.hr),
                             static_cast<unsigned long>(nativeIt->second.hr)));
            return true;
        }
        if (DebugGetPreparedTransferStrategyForSelfTest(state.taskB.value()) != FileOperations::OperationStrategy::Native ||
            DebugGetPreparedTransferStrategyForSelfTest(state.taskC.value()) != FileOperations::OperationStrategy::Native)
        {
            Fail(L"A same-endpoint Move is Native for a directory tree and for a regular file alike.");
            return true;
        }

        std::error_code ec;
        if (std::filesystem::exists(srcFoo, ec))
        {
            Fail(L"Same-volume move-merge left the source directory behind.");
            return true;
        }

        if (it->second.conflictPromptCount != 1u || nativeIt->second.conflictPromptCount != 0u || ! FileSizeEquals(dstFile, 4 * 1024) ||
            ! FileSizeEquals(dstNested, 6 * 1024) || ! FileSizeEquals(dstKeep, 3 * 1024) ||
            ! FileSizeEquals(dstConflict, 5 * 1024) ||
            std::filesystem::exists(nativeSource, ec) || ! FileSizeEquals(nativeTarget, 2 * 1024))
        {
            Fail(std::format(L"Same-volume directory Move integrity failure (mergePrompts={}, nativePrompts={}).",
                             it->second.conflictPromptCount,
                             nativeIt->second.conflictPromptCount));
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.MoveMergeIntoExistingFolderSameVolume.PromptCount",
                          L"native-rename-merge;native-regular-file",
                          0u,
                          it->second.conflictPromptCount + nativeIt->second.conflictPromptCount,
                          0u,
                          S_OK);

        SetFileOpsNativeMoveCreateDirectoryRaceForSelfTest(1u);
        AppendLog(L"Move-merge qualification: proving a raced destination directory continues as a rename merge");
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {raceSource},
                                                 raceDstRoot,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        AppendLog(std::format(L"Move-merge qualification: race admission returned task={}", state.taskA.value_or(0u)));
        const auto* raceTask = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const std::shared_ptr<const FileOperations::FileOperationPlanGroup> racePlans =
            raceTask != nullptr ? raceTask->LoadPlans() : nullptr;
        const auto* racePlan = racePlans && racePlans->size() == 1u
            ? std::get_if<FileOperations::TransferPlan>(&racePlans->front())
            : nullptr;
        if (racePlan == nullptr || racePlan->strategy != FileOperations::OperationStrategy::Native)
        {
            SetFileOpsNativeMoveCreateDirectoryRaceForSelfTest(0u);
            static_cast<void>(TakeFileOpsNativeMoveCreateDirectoryRaceAttemptsForSelfTest());
            Fail(L"A same-endpoint Local tree is admitted Native.");
            return true;
        }
        state.markerTick = nowTick;
        state.stepState  = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        const auto raceIt = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (raceIt == state.completedTasks.end())
        {
            return false;
        }
        const unsigned long raceAttempts = TakeFileOpsNativeMoveCreateDirectoryRaceAttemptsForSelfTest();
        SetFileOpsNativeMoveCreateDirectoryRaceForSelfTest(0u);
        std::error_code ec;
        if (raceAttempts != 1u || FAILED(raceIt->second.hr) || raceIt->second.conflictPromptCount != 0u ||
            std::filesystem::exists(raceSource, ec) || ! FileSizeEquals(raceTargetFile, 9 * 1024))
        {
            Fail(std::format(L"A raced destination directory must continue the Native item as a rename merge (nativeAttempts={}, hr=0x{:08X}, prompts={}).",
                             raceAttempts,
                             static_cast<unsigned long>(raceIt->second.hr),
                             raceIt->second.conflictPromptCount));
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.NativeDirectoryRaceContinuesAsRenameMerge",
                          L"known-noncommit;rename-merge;source-removed",
                          0u,
                          raceAttempts,
                          1u,
                          raceIt->second.hr);

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {mixedMergeSource, mixedNativeSource},
                                                 mixedDstRoot,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        AppendLog(std::format(L"Move-merge qualification: mixed admission returned task={}", state.taskA.value_or(0u)));
        // Both items share one endpoint and one Native plan; the existing-directory merge and the
        // destination-absent sibling differ only at execution.
        uint32_t mixedPlanCount = 0u;
        uint32_t mixedMask      = 0u;
        for (const ULONGLONG waitStart = GetTickCount64();
             state.taskA.has_value() && mixedPlanCount == 0u && GetTickCount64() - waitStart < 10'000ull;)
        {
            if (! DebugGetPreparedTransferPlanShapeForSelfTest(state.taskA.value(), &mixedPlanCount, &mixedMask))
            {
                Sleep(1);
            }
        }
        constexpr uint32_t kNativeBit = 1u << static_cast<uint32_t>(FileOperations::OperationStrategy::Native);
        if (mixedPlanCount != 1u || mixedMask != kNativeBit)
        {
            Fail(L"A Local selection with an existing-directory merge and a destination-absent sibling is one Native plan.");
            return true;
        }
        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        const auto mixedIt = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (mixedIt == state.completedTasks.end())
        {
            return false;
        }
        std::error_code ec;
        if (FAILED(mixedIt->second.hr) || mixedIt->second.conflictPromptCount != 0u ||
            std::filesystem::exists(mixedMergeSource, ec) || std::filesystem::exists(mixedNativeSource, ec) ||
            ! FileSizeEquals(mixedMergeTarget / L"merge.bin", 5 * 1024) ||
            ! FileSizeEquals(mixedMergeTarget / L"keep.bin", 2 * 1024) ||
            ! FileSizeEquals(mixedNativeTarget, 6 * 1024))
        {
            Fail(std::format(L"Mixed rename-merge and regular-file selection failed per-item execution (hr=0x{:08X}, prompts={}).",
                             static_cast<unsigned long>(mixedIt->second.hr),
                             mixedIt->second.conflictPromptCount));
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.MixedLocalMoveStrategies",
                          L"native-rename-merge;native-regular-file",
                          0u,
                          2u,
                          0u,
                          S_OK);

        std::wstring alternateVolumeSkipDetail;
        const std::optional<std::filesystem::path> alternateRoot =
            TryCreateAlternateWritableVolumeSelfTestRoot(state.tempRoot, alternateVolumeSkipDetail);
        if (! alternateRoot.has_value())
        {
            Debug::Perf::Emit(L"FileOps.SelfTest.CrossVolumeHostAdmission",
                              std::format(L"skip={}", alternateVolumeSkipDetail),
                              0u,
                              0u,
                              0u,
                              S_FALSE);
            NextStep(state, SelfTestState::Step::Beeline_RenameMergeSkipKeepsSourceFolder);
            return false;
        }

        state.fileOpsAlternateVolumeRoot = alternateRoot.value();
        if (! WriteTestFile(crossVolumeSource, 7 * 1024))
        {
            Fail(L"Failed to seed the real cross-volume host-admission source.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_MOVE,
                                                                 FolderWindow::Pane::Left,
                                                                 FolderWindow::Pane::Right,
                                                                 state.fsLocal,
                                                                 {crossVolumeSource},
                                                                 state.fileOpsAlternateVolumeRoot,
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to admit the real cross-volume Local Move through the host.");
            return true;
        }

        const auto* crossVolumeTask = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const std::shared_ptr<const FileOperations::FileOperationPlanGroup> crossVolumePlans =
            crossVolumeTask != nullptr ? crossVolumeTask->LoadPlans() : nullptr;
        const auto* crossVolumeTransfer = crossVolumePlans && crossVolumePlans->size() == 1u
            ? std::get_if<FileOperations::TransferPlan>(&crossVolumePlans->front())
            : nullptr;
        if (crossVolumeTransfer == nullptr || crossVolumeTransfer->strategy != FileOperations::OperationStrategy::Managed)
        {
            Fail(L"A real Local cross-volume Move must be admitted as Managed, not Native or CopyOnly.");
            return true;
        }

        state.stepState = 6;
        return false;
    }

    if (state.stepState == 6)
    {
        const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }

        const std::filesystem::path crossVolumeDestination = state.fileOpsAlternateVolumeRoot / crossVolumeSource.filename();
        std::error_code ec;
        if (FAILED(it->second.hr) || std::filesystem::exists(crossVolumeSource, ec) ||
            ! FileSizeEquals(crossVolumeDestination, 7 * 1024))
        {
            Fail(std::format(L"Real cross-volume host Managed Move failed integrity/source-disposition checks: 0x{:08X}.",
                             static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.CrossVolumeHostAdmission", L"strategy=managed;source=removed", 0u, 1u, 0u, S_OK);
        const std::filesystem::path alternateRoot = std::move(state.fileOpsAlternateVolumeRoot);
        std::filesystem::remove_all(alternateRoot, ec);
        PruneEmptyAlternateVolumeSandboxParents(alternateRoot);
        NextStep(state, SelfTestState::Step::Beeline_RenameMergeSkipKeepsSourceFolder);
        return false;
    }

    return false;
}
case SelfTestState::Step::Beeline_RenameMergeSkipKeepsSourceFolder:
{
    // Beeline: a rename merge answers each colliding child through the ordinary conflict surface.
    // Overwrite relocates through the provider's exact replace rename, Skip keeps that child, every
    // other child relocates by one rename, and the emptied-but-not-empty source folder stays as
    // `Moved; source folder kept`.
    using Task              = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(std::format(L"Beeline_RenameMergeSkipKeepsSourceFolder timed out at substep {}.", state.stepState));
        return true;
    }

    const std::filesystem::path srcRoot      = state.tempRoot / L"beeline-rename-merge-src";
    const std::filesystem::path dstRoot      = state.tempRoot / L"beeline-rename-merge-dst";
    const std::filesystem::path srcFoo       = srcRoot / L"Foo";
    const std::filesystem::path dstFoo       = dstRoot / L"Foo";
    const std::filesystem::path srcOverwrite = srcFoo / L"overwrite.bin";
    const std::filesystem::path dstOverwrite = dstFoo / L"overwrite.bin";
    const std::filesystem::path srcSkip      = srcFoo / L"skip.bin";
    const std::filesystem::path dstSkip      = dstFoo / L"skip.bin";
    const std::filesystem::path srcNested    = srcFoo / L"nested" / L"child.bin";
    const std::filesystem::path dstNested    = dstFoo / L"nested" / L"child.bin";

    if (state.stepState == 0u)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstFoo) || ! WriteTestFile(srcOverwrite, 5 * 1024) ||
            ! WriteTestFile(srcSkip, 3 * 1024) || ! WriteTestFile(srcNested, 2 * 1024) || ! WriteTestFile(dstOverwrite, 1 * 1024) ||
            ! WriteTestFile(dstSkip, 1 * 1024))
        {
            Fail(L"Beeline rename merge could not stage its trees.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFoo},
                                                 dstRoot,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Beeline rename merge could not start its move.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }

    // stepState bits: 2 = the Overwrite child was answered, 4 = the Skip child was answered.
    const bool overwriteAnswered = (state.stepState & 2u) != 0u;
    const bool skipAnswered      = (state.stepState & 4u) != 0u;
    Task* const task             = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    if (! overwriteAnswered || ! skipAnswered)
    {
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            if (state.taskA.has_value() && state.completedTasks.contains(state.taskA.value()))
            {
                Fail(L"Beeline rename merge completed without exposing both colliding children through the conflict surface.");
                return true;
            }
            return false;
        }
        const bool isOverwrite = NormalizePathForCompare(prompt->sourcePath) == NormalizePathForCompare(srcOverwrite.native());
        const bool isSkip      = NormalizePathForCompare(prompt->sourcePath) == NormalizePathForCompare(srcSkip.native());
        if (prompt->bucket != Task::ConflictBucket::RegularFileExists || (! isOverwrite && ! isSkip))
        {
            Fail(std::format(L"Beeline rename merge exposed an unexpected conflict: '{}' -> '{}'.", prompt->sourcePath, prompt->destinationPath));
            return true;
        }
        if (isOverwrite && ! overwriteAnswered)
        {
            task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
            state.stepState |= 2u;
        }
        else if (isSkip && ! skipAnswered)
        {
            task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
            state.stepState |= 4u;
        }
        return false;
    }

    const auto completed = state.completedTasks.find(state.taskA.value());
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    std::error_code ec;
    const bool sourceKept = completed->second.hr == S_FALSE || completed->second.hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    if (! sourceKept || completed->second.conflictPromptCount != 2u || ! FileSizeEquals(dstOverwrite, 5 * 1024) ||
        std::filesystem::exists(srcOverwrite, ec) || ! FileSizeEquals(dstNested, 2 * 1024) || std::filesystem::exists(srcFoo / L"nested", ec) ||
        ! FileSizeEquals(srcSkip, 3 * 1024) || ! FileSizeEquals(dstSkip, 1 * 1024) || ! std::filesystem::is_directory(srcFoo, ec))
    {
        Fail(std::format(L"Beeline rename merge with a skipped child must relocate the rest and keep the source folder (hr=0x{:08X}, prompts={}).",
                         static_cast<unsigned long>(completed->second.hr),
                         completed->second.conflictPromptCount));
        return true;
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.SameVolumeRenameMerge",
                      L"shape=overwrite+skip+subtree;source-folder-kept",
                      1u,
                      completed->second.conflictPromptCount,
                      0u,
                      completed->second.hr);
    state.taskA.reset();
    NextStep(state, SelfTestState::Step::FileOps_ReparseDirectoryMergeIntoExistingFolder);
    return false;
}
#endif
#if defined(FILEOPS_SELFTEST_INCLUDE_REPARSE_MERGE)

case SelfTestState::Step::FileOps_ReparseDirectoryMergeIntoExistingFolder:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        Fail(L"FileOps_ReparseDirectoryMergeIntoExistingFolder timed out.");
        return true;
    }

    const std::filesystem::path srcRoot         = state.tempRoot / L"clearflow-reparse-merge-src";
    const std::filesystem::path dstRoot         = state.tempRoot / L"clearflow-reparse-merge-dst";
    const std::filesystem::path targetRoot      = state.tempRoot / L"clearflow-reparse-merge-target";
    const std::filesystem::path targetFile      = targetRoot / L"payload.bin";
    const std::filesystem::path sourceLink      = srcRoot / L"linkToTarget";
    const std::filesystem::path destinationLink = dstRoot / L"linkToTarget";
    const std::filesystem::path nativeSourceRoot       = state.tempRoot / L"clearflow-native-destination-link-src";
    const std::filesystem::path nativeDestinationRoot  = state.tempRoot / L"clearflow-native-destination-link-dst";
    const std::filesystem::path nativeLinkTargetRoot   = state.tempRoot / L"clearflow-native-destination-link-target";
    const std::filesystem::path nativeSource           = nativeSourceRoot / L"Folder";
    const std::filesystem::path nativeSourceFile       = nativeSource / L"source-only.bin";
    const std::filesystem::path nativeDestinationLink = nativeDestinationRoot / nativeSource.filename();
    const std::filesystem::path nativeTargetSentinel  = nativeLinkTargetRoot / L"target-only.bin";

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"preserve"})json"));

        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(targetRoot) ||
            ! RecreateEmptyDirectory(nativeSourceRoot) || ! RecreateEmptyDirectory(nativeDestinationRoot) ||
            ! RecreateEmptyDirectory(nativeLinkTargetRoot) || ! WriteTestFile(targetFile, 512) ||
            ! WriteTestFile(nativeSourceFile, 384) || ! WriteTestFile(nativeTargetSentinel, 256))
        {
            Fail(L"Failed to reset reparse merge directories.");
            return true;
        }

        if (! TryCreateJunction(sourceLink, targetRoot))
        {
            Fail(L"Failed to create source junction for reparse merge test.");
            return true;
        }
        if (! TryCreateJunction(nativeDestinationLink, nativeLinkTargetRoot))
        {
            Fail(L"Failed to create destination junction for Native directory-Move qualification test.");
            return true;
        }

        FileSystemOptions nativeOptions{};
        nativeOptions.sizeBytes  = sizeof(nativeOptions);
        nativeOptions.moveMode   = FILESYSTEM_MOVE_NATIVE_ONLY;
        nativeOptions.linkPolicy = FILESYSTEM_LINK_PRESERVE;
        const HRESULT nativeProviderHr = state.fsLocal->MoveItem(nativeSource.c_str(),
                                                                  nativeDestinationLink.c_str(),
                                                                  static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                                  &nativeOptions,
                                                                  nullptr,
                                                                  nullptr);
        if (nativeProviderHr != HRESULT_FROM_WIN32(ERROR_REPARSE_POINT_ENCOUNTERED) || ! FileSizeEquals(nativeSourceFile, 384) ||
            ! FileSizeEquals(nativeTargetSentinel, 256) ||
            std::filesystem::exists(nativeLinkTargetRoot / nativeSourceFile.filename()))
        {
            Fail(std::format(L"Local Native Move must reject a destination junction without merging through it (hr=0x{:08X}).",
                             static_cast<unsigned long>(nativeProviderHr)));
            return true;
        }

        // Admission never probes the destination shape. The pre-consumption gate stops the task once
        // Preparing has published its strategy and before any mutation.
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {nativeSource},
                                                 nativeDestinationRoot,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 true,
                                                 nullptr,
                                                 {},
                                                 std::nullopt,
                                                 {},
                                                 []() noexcept -> HRESULT { return HRESULT_FROM_WIN32(ERROR_CANCELLED); });
        std::optional<FileOperations::OperationStrategy> destinationLinkStrategy;
        for (const ULONGLONG waitStart = GetTickCount64();
             state.taskA.has_value() && ! destinationLinkStrategy.has_value() && GetTickCount64() - waitStart < 10'000ull;)
        {
            destinationLinkStrategy = DebugGetPreparedTransferStrategyForSelfTest(state.taskA.value());
            if (! destinationLinkStrategy.has_value())
            {
                Sleep(1);
            }
        }
        if (destinationLinkStrategy != FileOperations::OperationStrategy::Native)
        {
            Fail(L"A same-endpoint directory Move is admitted Native; a destination junction is the provider's typed link conflict at execution.");
            return true;
        }
        state.taskA.reset();

        std::error_code ec;
        std::filesystem::create_directories(destinationLink, ec);
        if (ec)
        {
            Fail(L"Failed to create existing destination directory for reparse merge test.");
            return true;
        }

        const HRESULT noGrantCopyHr = state.fsLocal->CopyItem(
            sourceLink.c_str(), destinationLink.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), nullptr, nullptr, nullptr);
        if (noGrantCopyHr != HRESULT_FROM_WIN32(ERROR_DATATYPE_MISMATCH))
        {
            Fail(std::format(L"Link-on-directory Copy expected typed mismatch, got 0x{:08X}.",
                             static_cast<unsigned long>(noGrantCopyHr)));
            return true;
        }

        const DWORD noGrantAttributes = ::GetFileAttributesW(destinationLink.c_str());
        if (noGrantAttributes == INVALID_FILE_ATTRIBUTES || (noGrantAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0 ||
            (noGrantAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0)
        {
            Fail(L"Reparse merge copy without overwrite converted the existing real directory.");
            return true;
        }

        if (! FileSizeEquals(targetFile, 512))
        {
            Fail(L"Reparse merge copy without overwrite damaged the out-of-tree target sentinel.");
            return true;
        }

        const HRESULT overwriteCopyHr = state.fsLocal->CopyItem(
            sourceLink.c_str(),
            destinationLink.c_str(),
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE),
            nullptr,
            nullptr,
            nullptr);
        if (overwriteCopyHr != HRESULT_FROM_WIN32(ERROR_DATATYPE_MISMATCH))
        {
            Fail(std::format(L"An operation-wide Overwrite flag must not replace a directory with a link (hr=0x{:08X}).",
                             static_cast<unsigned long>(overwriteCopyHr)));
            return true;
        }

        const DWORD finalDestinationAttributes = ::GetFileAttributesW(destinationLink.c_str());
        if (finalDestinationAttributes == INVALID_FILE_ATTRIBUTES ||
            (finalDestinationAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u ||
            (finalDestinationAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
        {
            Fail(L"Type-mismatch handling changed the existing real destination directory.");
            return true;
        }

        if (! FileSizeEquals(targetFile, 512))
        {
            Fail(L"Type-mismatch handling damaged the out-of-tree link target sentinel.");
            return true;
        }

        NextStep(state, SelfTestState::Step::FileOps_ProviderCapabilityMatrix);
        return false;
    }

    return false;
}
#endif
#if defined(FILEOPS_SELFTEST_INCLUDE_PROVIDER_MATRIX)

case SelfTestState::Step::FileOps_ProviderCapabilityMatrix:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 30'000ull))
    {
        Fail(L"FileOps_ProviderCapabilityMatrix timed out.");
        return true;
    }

    if (state.stepState == 0)
    {
        const auto requireCapabilities = [&](IFileSystem* fs,
                                             std::wstring_view providerName,
                                             std::wstring_view providerId,
                                             ProviderCapabilitySnapshot& snapshot) noexcept -> bool
        {
            std::wstring reason;
            if (! TryReadTypedProviderCapabilities(fs, providerId, snapshot, reason))
            {
                Fail(std::format(L"{} capabilities failed: {}.", providerName, reason));
                return false;
            }
            return true;
        };

        ProviderCapabilitySnapshot localCaps{};
        ProviderCapabilitySnapshot dummyCaps{};
        ProviderCapabilitySnapshot sevenZipCaps{};
        if (! requireCapabilities(state.fsLocal.get(), L"Local FileSystem", kPluginIdLocal, localCaps) ||
            ! requireCapabilities(state.fsDummy.get(), L"FileSystemDummy", kPluginIdDummy, dummyCaps) ||
            ! requireCapabilities(state.fs7z.get(), L"FileSystem7z", kPluginId7z, sevenZipCaps))
        {
            return true;
        }

        const auto require = [&](bool condition, std::wstring_view message) noexcept -> bool
        {
            if (! condition)
            {
                Fail(message);
                return false;
            }
            return true;
        };

        constexpr uint64_t kR2RouteQueryCount = 4'096u;
        const auto r2RouteQueryStarted         = std::chrono::steady_clock::now();
        bool r2RouteQueryValid                 = true;
        uint64_t r2ResultBytes                 = 0u;
        uint64_t r2ArenaFallbackCount          = 0u;
        uint64_t r2RejectedCount               = 0u;
        for (uint64_t iteration = 0u; iteration < kR2RouteQueryCount; ++iteration)
        {
            const FileSystemRouteContract::QueryResult route =
                FileSystemRouteContract::Query(state.fsDummy.get(), L"/", FILESYSTEM_COPY, kPluginIdDummy);
            r2RouteQueryValid = route.state == FileSystemRouteContract::QueryState::Available && r2RouteQueryValid;
            r2ArenaFallbackCount += route.usedArenaFallback ? 1u : 0u;
            r2RejectedCount += route.state == FileSystemRouteContract::QueryState::Available ? 0u : 1u;
            r2ResultBytes += sizeof(FileSystemRouteContract::Snapshot);
        }
        const uint64_t r2RouteQueryElapsedUs = Debug::Perf::ElapsedUs(r2RouteQueryStarted);
        Debug::Perf::Emit(L"FileOps.RouteFacts.QueryUs",
                          L"dummy-offline-4096-route-queries",
                          r2RouteQueryElapsedUs,
                          kR2RouteQueryCount,
                          0u,
                          r2RouteQueryValid ? S_OK : E_FAIL);
        Debug::Perf::Emit(L"FileOps.RouteFacts.QueryCount",
                          L"dummy-offline-4096-route-queries",
                          0u,
                          kR2RouteQueryCount,
                          kR2RouteQueryCount,
                          r2RouteQueryValid ? S_OK : E_FAIL);
        Debug::Perf::Emit(L"FileOps.RouteFacts.JsonParseCount",
                          L"typed-route-candidate",
                          0u,
                          0u,
                          kR2RouteQueryCount,
                          r2RouteQueryValid ? S_OK : E_FAIL);
        Debug::Perf::Emit(L"FileOps.RouteFacts.TypedQueryCount",
                          L"typed-route-candidate",
                          0u,
                          kR2RouteQueryCount,
                          kR2RouteQueryCount,
                          r2RouteQueryValid ? S_OK : E_FAIL);
        Debug::Perf::Emit(L"FileOps.RouteFacts.ArenaFallbackCount",
                          L"typed-route-candidate",
                          0u,
                          r2ArenaFallbackCount,
                          0u,
                          r2ArenaFallbackCount == 0u ? S_OK : E_FAIL);
        Debug::Perf::Emit(L"FileOps.RouteFacts.RejectedCount",
                          L"typed-route-candidate",
                          0u,
                          r2RejectedCount,
                          0u,
                          r2RejectedCount == 0u ? S_OK : E_FAIL);
        Debug::Perf::Emit(L"FileOps.RouteFacts.ResultBytes",
                          L"typed-route-candidate",
                          0u,
                          r2ResultBytes,
                          sizeof(FileSystemRouteContract::Snapshot),
                          r2RouteQueryValid ? S_OK : E_FAIL);
        if (! require(r2RouteQueryValid && r2ArenaFallbackCount == 0u && r2RejectedCount == 0u &&
                          r2RouteQueryElapsedUs < 2'000'000u,
                      L"R2 typed capability queries must remain valid, allocation-free on normal facts, and below two seconds."))
        {
            return true;
        }

        constexpr IID kR2RouteCapabilitiesIid{
            0x1e924d87, 0x2e62, 0x4ab4, {0x9f, 0x37, 0xc5, 0x65, 0xd4, 0x65, 0xf2, 0x5e}};
        wil::com_ptr<IUnknown> typedRouteCapabilities;
        const HRESULT typedRouteCapabilitiesHr =
            state.fsDummy->QueryInterface(kR2RouteCapabilitiesIid, typedRouteCapabilities.put_void());
        if (! require(SUCCEEDED(typedRouteCapabilitiesHr) && typedRouteCapabilities,
                      L"Every shipped provider must expose the R2 typed route-capability IID; JSON cannot be executable authority."))
        {
            return true;
        }

        if (! require(localCaps.preserveFileLink && localCaps.preserveDirectoryLink && ! localCaps.retargetInTree &&
                          localCaps.exactLinkRemoval && ! dummyCaps.preserveFileLink && ! dummyCaps.preserveDirectoryLink &&
                          ! dummyCaps.retargetInTree && ! dummyCaps.exactLinkRemoval && ! sevenZipCaps.preserveFileLink &&
                          ! sevenZipCaps.preserveDirectoryLink && ! sevenZipCaps.retargetInTree && ! sevenZipCaps.exactLinkRemoval,
                      L"Literal link preservation (no in-tree retarget) must remain a truthful path-profile claim."))
        {
            return true;
        }

        if (! require(localCaps.createDirectoryOperation && dummyCaps.createDirectoryOperation && ! sevenZipCaps.createDirectoryOperation,
                      L"Create Directory capability claims must be truthful for Local, Dummy, and read-only 7z profiles."))
        {
            return true;
        }

        FileOperations::CreateDirectoryAdmission localDirectoryAdmission{};
        FileOperations::CreateDirectoryAdmission dummyDirectoryAdmission{};
        FileOperations::CreateDirectoryAdmission rejectedDirectoryAdmission{};
        const HRESULT localDirectoryAdmissionHr = state.fileOps->QualifyCreateDirectory(
            state.fsLocal, kPluginIdLocal, {}, state.tempRoot, L"qualified-create-directory", localDirectoryAdmission);
        const HRESULT dummyDirectoryAdmissionHr = state.fileOps->QualifyCreateDirectory(
            state.fsDummy, kPluginIdDummy, {}, std::filesystem::path(L"/"), L"qualified-create-directory", dummyDirectoryAdmission);
        const HRESULT readOnlyDirectoryAdmissionHr = state.fileOps->QualifyCreateDirectory(
            state.fs7z, kPluginId7z, {}, std::filesystem::path(L"/"), L"qualified-create-directory", rejectedDirectoryAdmission);
        const HRESULT invalidSeparatorAdmissionHr = state.fileOps->QualifyCreateDirectory(
            state.fsDummy, kPluginIdDummy, {}, std::filesystem::path(L"/"), L"nested/name", rejectedDirectoryAdmission);
        const HRESULT deviceNamespaceAdmissionHr = state.fileOps->QualifyCreateDirectory(
            state.fsLocal,
            kPluginIdLocal,
            {},
            std::filesystem::path(LR"(\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1)"),
            L"blocked",
            rejectedDirectoryAdmission);
        if (! require(localDirectoryAdmissionHr == S_OK && localDirectoryAdmission.allowLocalNativeFallback &&
                          ! localDirectoryAdmission.endpoint.rootId.empty() &&
                          localDirectoryAdmission.candidateProviderPath == state.tempRoot / L"qualified-create-directory",
                      L"Local Create Directory admission should qualify a concrete candidate and explicitly authorize only the Local native fallback."))
        {
            return true;
        }
        if (! require(dummyDirectoryAdmissionHr == S_OK && ! dummyDirectoryAdmission.allowLocalNativeFallback &&
                          dummyDirectoryAdmission.candidateProviderPath.native() == L"/qualified-create-directory",
                      L"Provider Create Directory admission should use the provider separator without enabling the Win32 fallback."))
        {
            return true;
        }
        if (! require(readOnlyDirectoryAdmissionHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) && FAILED(deviceNamespaceAdmissionHr),
                      L"Create Directory admission must reject a false capability claim and Local GLOBALROOT/device namespace envelopes before mutation."))
        {
            return true;
        }
        if (! require(invalidSeparatorAdmissionHr == HRESULT_FROM_WIN32(ERROR_INVALID_NAME),
                      L"Create Directory admission must reject a leaf containing the active provider separator before mutation."))
        {
            return true;
        }

        const FileOperations::MoveStrategyQualificationFacts managedFacts{
            .nativeMoveQualified = false,
            .copyPairQualified = true,
            .movePairQualified = true,
            .sourceBoundDelete = true,
            .sourceConditionalDelete = true,
            .destinationExclusiveStage = true,
            .destinationConditionalPublish = true,
            .sourceBindingAvailable = true,
            .destinationBindingAvailable = true,
        };
        const std::optional<FileOperations::OperationStrategy> managedSelection =
            FileOperations::SelectMoveStrategy(managedFacts);
        FileOperations::MoveStrategyQualificationFacts nativeFacts{};
        nativeFacts.nativeMoveQualified = true;
        FileOperations::MoveStrategyQualificationFacts copyOnlyFacts = managedFacts;
        copyOnlyFacts.movePairQualified = false;
        FileOperations::MoveStrategyQualificationFacts unboundFacts = managedFacts;
        unboundFacts.sourceBindingAvailable = false;
        FileOperations::MoveStrategyQualificationFacts unsupportedFacts = managedFacts;
        unsupportedFacts.copyPairQualified = false;
        unsupportedFacts.movePairQualified = false;
        if (! require(FileOperations::SelectMoveStrategy(nativeFacts) == FileOperations::OperationStrategy::Native &&
                          managedSelection == FileOperations::OperationStrategy::Managed &&
                          FileOperations::SelectMoveStrategy(copyOnlyFacts) == FileOperations::OperationStrategy::CopyOnly &&
                          FileOperations::SelectMoveStrategy(unboundFacts) == FileOperations::OperationStrategy::CopyOnly &&
                          ! FileOperations::SelectMoveStrategy(unsupportedFacts).has_value(),
                      L"Move strategy qualification must prefer Native, require every destructive proof for Managed, downgrade through Copy, and reject without Copy."))
        {
            return true;
        }

        const FileOperations::ManagedCleanupMutationFacts removedFacts{
            .status = S_OK,
            .outcomeKnown = true,
            .mutationCommitted = true,
            .originalStillPresent = false,
        };
        FileOperations::ManagedCleanupMutationFacts retainedFacts = removedFacts;
        retainedFacts.status = HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION);
        retainedFacts.mutationCommitted = false;
        retainedFacts.originalStillPresent = true;
        FileOperations::ManagedCleanupMutationFacts indeterminateFacts = retainedFacts;
        indeterminateFacts.outcomeKnown = false;
        FileOperations::ManagedCleanupMutationFacts contradictoryFacts = removedFacts;
        contradictoryFacts.originalStillPresent = true;
        if (! require(FileOperations::ClassifyManagedCleanupMutation(removedFacts) ==
                              FileOperations::ManagedCleanupAttemptDisposition::Removed &&
                          FileOperations::ClassifyManagedCleanupMutation(retainedFacts) ==
                              FileOperations::ManagedCleanupAttemptDisposition::Retained &&
                          FileOperations::ClassifyManagedCleanupMutation(indeterminateFacts) ==
                              FileOperations::ManagedCleanupAttemptDisposition::Indeterminate &&
                          FileOperations::ClassifyManagedCleanupMutation(contradictoryFacts) ==
                              FileOperations::ManagedCleanupAttemptDisposition::ProviderContractViolation,
                      L"Managed cleanup results must distinguish Removed, retryable Retained, unknown outcome, and provider-contract contradiction."))
        {
            return true;
        }

        const FileOperations::QualifiedEndpoint validationEndpoint{
            .pluginId   = L"selftest/provider",
            .instanceId = L"instance",
            .profileId  = L"profile",
            .rootId     = L"root",
            .pathIdentity = FileSystemPathIdentity{
                .pathTextStableIdentity = true,
                .componentComparison    = FileSystemPathComponentComparison::OrdinalCaseSensitive,
                .preferredSeparator     = L'/',
                .acceptedSeparators     = L"/",
                .casePreserving         = true,
                .caseOnlyRename         = FileSystemPathCaseOnlyRename::NotApplicable,
            },
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };
        FileOperations::TransferPlan validTransfer{};
        validTransfer.intent              = FileOperations::TransferIntent::Move;
        validTransfer.strategy            = FileOperations::OperationStrategy::CopyOnly;
        validTransfer.sourceEndpoint      = validationEndpoint;
        validTransfer.destinationEndpoint = validationEndpoint;
        validTransfer.selectedItems       = {{.providerPath = L"/source/a.txt"}, {.providerPath = L"/source/b.txt"}};
        validTransfer.destination.providerFolderPath = L"/destination";
        validTransfer.explicitMappings = {
            {.sourceIndex = 0u, .destinationProviderPath = L"/destination/a.txt"},
            {.sourceIndex = 1u, .destinationProviderPath = L"/destination/b.txt"},
        };
        validTransfer.moveClipboardSequence = FileOperations::ClipboardSequence{.windowsSequenceNumber = 42u};
        FileOperations::PlanRejectionBucket validationBucket = FileOperations::PlanRejectionBucket::None;
        if (! require(SUCCEEDED(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{validTransfer}, &validationBucket)) &&
                          validationBucket == FileOperations::PlanRejectionBucket::None,
                      L"A complete typed transfer plan should pass deterministic validation."))
        {
            return true;
        }

        FileOperations::TransferPlan explicitCopy = validTransfer;
        explicitCopy.intent                        = FileOperations::TransferIntent::Copy;
        explicitCopy.strategy                      = FileOperations::OperationStrategy::Copy;
        explicitCopy.moveClipboardSequence.reset();
        if (! require(SUCCEEDED(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{explicitCopy}, &validationBucket)),
                      L"Ordinary Copy must use its explicit Copy strategy instead of borrowing the Managed Move strategy."))
        {
            return true;
        }
        explicitCopy.strategy = FileOperations::OperationStrategy::Managed;
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{explicitCopy}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::InvalidStrategy,
                      L"A Copy plan mislabeled as Managed Move must fail typed validation."))
        {
            return true;
        }

        FileOperations::TransferPlan crossProfileDestination = validTransfer;
        crossProfileDestination.intent                       = FileOperations::TransferIntent::Copy;
        crossProfileDestination.strategy                     = FileOperations::OperationStrategy::Copy;
        crossProfileDestination.selectedItems                = {{.providerPath = L"/source/a.txt"}};
        crossProfileDestination.explicitMappings.clear();
        crossProfileDestination.moveClipboardSequence.reset();
        crossProfileDestination.destinationEndpoint.rootId = L"destination-root";
        crossProfileDestination.destinationEndpoint.pathIdentity = FileSystemPathIdentity{
            .pathTextStableIdentity = true,
            .componentComparison    = FileSystemPathComponentComparison::OrdinalCaseSensitive,
            .preferredSeparator     = L'\\',
            .acceptedSeparators     = L"\\",
            .casePreserving         = true,
            .caseOnlyRename         = FileSystemPathCaseOnlyRename::NotApplicable,
        };
        crossProfileDestination.destination.providerFolderPath = L"bucket\\destination";
        std::wstring resolvedProviderDestination;
        if (! require(FileOperations::TryResolveTransferDestinationProviderPath(
                          crossProfileDestination, 0u, resolvedProviderDestination) &&
                          resolvedProviderDestination == L"bucket\\destination\\a.txt",
                      L"Transfer destination resolution must derive the leaf from the source profile and join with the destination profile."))
        {
            return true;
        }
        crossProfileDestination.explicitMappings = {{.sourceIndex = 0u, .destinationProviderPath = L"bucket\\destination\\mapped.txt"}};
        if (! require(FileOperations::TryResolveTransferDestinationProviderPath(
                          crossProfileDestination, 0u, resolvedProviderDestination) &&
                          resolvedProviderDestination == L"bucket\\destination\\mapped.txt",
                      L"An explicit transfer mapping must remain the exact execution and completion destination."))
        {
            return true;
        }

        FileOperations::TransferPlan secondRootTransfer = validTransfer;
        secondRootTransfer.sourceEndpoint.rootId         = L"second-root";
        secondRootTransfer.selectedItems                 = {{.providerPath = L"/second-root/c.txt"}};
        secondRootTransfer.explicitMappings              = {{.sourceIndex = 0u, .destinationProviderPath = L"/destination/c.txt"}};
        const FileOperations::FileOperationPlanGroup mixedRootGroup{
            FileOperations::FileOperationPlan{validTransfer},
            FileOperations::FileOperationPlan{secondRootTransfer},
        };
        if (! require(mixedRootGroup.size() == 2u &&
                          std::get<FileOperations::TransferPlan>(mixedRootGroup[0]).sourceEndpoint.rootId !=
                              std::get<FileOperations::TransferPlan>(mixedRootGroup[1]).sourceEndpoint.rootId &&
                          SUCCEEDED(FileOperations::ValidatePlan(mixedRootGroup[0], &validationBucket)) &&
                          SUCCEEDED(FileOperations::ValidatePlan(mixedRootGroup[1], &validationBucket)),
                      L"A mixed-root admission must retain separately valid one-root child plans."))
        {
            return true;
        }

        FileOperations::TransferPlan duplicateMapping = validTransfer;
        duplicateMapping.explicitMappings[1].sourceIndex = 0u;
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{duplicateMapping}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::InvalidExplicitMappings,
                      L"Typed admission must reject duplicate or incomplete explicit mappings."))
        {
            return true;
        }

        FileOperations::TransferPlan escapingMapping = validTransfer;
        escapingMapping.explicitMappings[1].destinationProviderPath = L"/outside/b.txt";
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{escapingMapping}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::EscapingDestinationMapping,
                      L"Typed admission must reject an explicit destination mapping that escapes its qualified destination folder."))
        {
            return true;
        }

        FileOperations::QualifiedEndpoint ignoreCaseEndpoint = validationEndpoint;
        ignoreCaseEndpoint.pluginId                           = L"selftest/ignore-case-provider";
        ignoreCaseEndpoint.profileId                          = L"ignore-case-profile";
        ignoreCaseEndpoint.pathIdentity->componentComparison  = FileSystemPathComponentComparison::OrdinalIgnoreCase;
        ignoreCaseEndpoint.pathIdentity->acceptedSeparators   = L"\\/";

        FileOperations::TransferPlan ignoreCaseMapping = validTransfer;
        ignoreCaseMapping.intent                       = FileOperations::TransferIntent::Copy;
        ignoreCaseMapping.strategy                     = FileOperations::OperationStrategy::Copy;
        ignoreCaseMapping.sourceEndpoint               = ignoreCaseEndpoint;
        ignoreCaseMapping.destinationEndpoint          = ignoreCaseEndpoint;
        ignoreCaseMapping.selectedItems                = {{.providerPath = L"/source/a.txt"}};
        ignoreCaseMapping.destination.providerFolderPath = L"/Destination";
        ignoreCaseMapping.explicitMappings = {{.sourceIndex = 0u, .destinationProviderPath = L"/destination/a.txt"}};
        ignoreCaseMapping.moveClipboardSequence.reset();
        if (! require(SUCCEEDED(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{ignoreCaseMapping}, &validationBucket)),
                      L"Destination containment must use the admitted ignore-case path profile, not a host-local lexical comparison."))
        {
            return true;
        }

        FileOperations::TransferPlan ignoreCaseSameFolder = ignoreCaseMapping;
        ignoreCaseSameFolder.intent                        = FileOperations::TransferIntent::Move;
        ignoreCaseSameFolder.strategy                      = FileOperations::OperationStrategy::Native;
        ignoreCaseSameFolder.selectedItems                 = {{.providerPath = L"/Foo/a.txt"}};
        ignoreCaseSameFolder.destination.providerFolderPath = L"/foo";
        ignoreCaseSameFolder.explicitMappings.clear();
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{ignoreCaseSameFolder}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::SameFolderMove,
                      L"Same-folder Move admission must use the provider's ignore-case path profile."))
        {
            return true;
        }

        FileOperations::TransferPlan ignoreCaseEscape = ignoreCaseMapping;
        ignoreCaseEscape.explicitMappings.front().destinationProviderPath = L"/Destination-Else/a.txt";
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{ignoreCaseEscape}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::EscapingDestinationMapping,
                      L"Path-profile containment must still reject a sibling prefix with the same textual start."))
        {
            return true;
        }

        const FileOperations::QualifiedEndpoint localValidationEndpoint{
            .pluginId   = L"builtin/file-system",
            .instanceId = L"host/default",
            .profileId  = L"local-win32",
            .rootId     = L"local-volume",
            .pathIdentity = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };
        if (! require(! FileOperations::MutationInterlockAccessesConflict(FileOperations::MutationInterlockAccess::ReadSource,
                                                                          FileOperations::MutationInterlockAccess::ReadSource) &&
                          FileOperations::MutationInterlockAccessesConflict(FileOperations::MutationInterlockAccess::ReadSource,
                                                                            FileOperations::MutationInterlockAccess::WriteSource) &&
                          FileOperations::MutationInterlockAccessesConflict(FileOperations::MutationInterlockAccess::ReadSource,
                                                                            FileOperations::MutationInterlockAccess::PublishDestination) &&
                          FileOperations::MutationInterlockAccessesConflict(FileOperations::MutationInterlockAccess::WriteSource,
                                                                            FileOperations::MutationInterlockAccess::WriteSource) &&
                          FileOperations::MutationInterlockAccessesConflict(FileOperations::MutationInterlockAccess::PublishDestination,
                                                                            FileOperations::MutationInterlockAccess::PublishDestination),
                      L"Mutation interlock access roles must permit shared source reads and serialize every overlapping writer/publisher pair."))
        {
            return true;
        }
        FileOperations::TransferPlan sameFolderCopy{};
        sameFolderCopy.intent              = FileOperations::TransferIntent::Copy;
        sameFolderCopy.sourceEndpoint      = localValidationEndpoint;
        sameFolderCopy.destinationEndpoint = localValidationEndpoint;
        sameFolderCopy.selectedItems       = {{.providerPath = L"C:\\work\\report.txt"}};
        sameFolderCopy.destination.providerFolderPath = L"C:\\work";
        if (! require(SUCCEEDED(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{sameFolderCopy}, &validationBucket)),
                      L"Same-folder Copy must remain admissible so execution can create a Keep Both sibling."))
        {
            return true;
        }

        FileOperations::TransferPlan destinationInsideSource = sameFolderCopy;
        destinationInsideSource.selectedItems.front().providerPath = L"C:\\work\\tree";
        destinationInsideSource.destination.providerFolderPath     = L"C:\\work\\tree\\child";
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{destinationInsideSource}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::DestinationInsideSource,
                      L"Copy destination-inside-source must be rejected by the stable path profile before task publication."))
        {
            return true;
        }

        FileOperations::TransferPlan sameFolderMove = sameFolderCopy;
        sameFolderMove.intent   = FileOperations::TransferIntent::Move;
        sameFolderMove.strategy = FileOperations::OperationStrategy::Native;
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{sameFolderMove}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::SameFolderMove,
                      L"Same-folder Move must be rejected structurally while same-folder Copy remains admissible."))
        {
            return true;
        }

        FileOperations::TransferPlan deviceNamespace = sameFolderCopy;
        deviceNamespace.selectedItems.front().providerPath = LR"(\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\report.txt)";
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{deviceNamespace}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::UnsupportedDeviceNamespace,
                      L"Local device namespaces must fail the immutable plan envelope before confirmation or queue publication."))
        {
            return true;
        }

        FileOperations::TransferPlan invalidClipboard = validTransfer;
        invalidClipboard.intent   = FileOperations::TransferIntent::Copy;
        invalidClipboard.strategy = FileOperations::OperationStrategy::Copy;
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{invalidClipboard}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::InvalidClipboardSequence,
                      L"Only Move plans may carry a non-zero clipboard sequence."))
        {
            return true;
        }

        FileOperations::DeletePlan permanentDelete{};
        permanentDelete.endpoint      = validationEndpoint;
        permanentDelete.selectedItems = {{
            .providerPath = L"/source/a.txt",
            .ingressSnapshot = FileOperations::ProviderIdentitySnapshot{
                .objectId = {std::byte{0x01}},
                .pathProfileId = validationEndpoint.profileId,
            },
        }};
        permanentDelete.mode          = FileOperations::DeleteMode::Permanent;
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{permanentDelete}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::MissingDestructiveConsent,
                      L"A permanent Delete plan without a task-scoped consent receipt must fail validation."))
        {
            return true;
        }
        permanentDelete.initialConsent = FileOperations::DestructiveConsentReceipt{
            .kind      = FileOperations::ConsentKind::PermanentDelete,
            .taskNonce = 7u,
        };
        if (! require(SUCCEEDED(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{permanentDelete}, &validationBucket)),
                      L"A permanent Delete plan with a non-zero task-scoped consent receipt should pass validation."))
        {
            return true;
        }

        FileOperations::RenamePlan unspecifiedRename{};
        unspecifiedRename.endpoint = validationEndpoint;
        unspecifiedRename.finalMappings = {{.source = {.providerPath = L"/source/a.txt"}, .finalLeafName = L"renamed.txt"}};
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{unspecifiedRename}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::InvalidRename,
                      L"RenamePlan must reject an unspecified origin so no caller inherits the Inline Rename exception by default."))
        {
            return true;
        }

        FileOperations::RenamePlan invalidRename = unspecifiedRename;
        invalidRename.origin = FileOperations::RenameOrigin::InlineRename;
        invalidRename.finalMappings = {{.source = {.providerPath = L"/source/a.txt"}, .finalLeafName = L"nested/name.txt"}};
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{invalidRename}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::InvalidRename,
                      L"RenamePlan leaf names must not contain provider path separators."))
        {
            return true;
        }

        const std::filesystem::path bindingRoot = state.tempRoot / L"host-object-binding";
        const std::filesystem::path bindingPath = bindingRoot / L"source.bin";
        const std::filesystem::path bindingAlias = bindingRoot / L"source-alias.bin";
        if (! require(SelfTest::EnsureDirectory(bindingRoot) && SelfTest::WriteTextFile(bindingPath, "original") &&
                          CreateHardLinkW(bindingAlias.c_str(), bindingPath.c_str(), nullptr) != FALSE,
                      L"Host object-binding fixture should create a file and hard-link alias."))
        {
            return true;
        }

        constexpr FileSystemBindFlags hostBindingFlags =
            static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_CONTENT | FILESYSTEM_BIND_READ_METADATA);
        FileOperations::ObjectBindingResult boundAuthority =
            FileOperations::BindObjectAuthority(state.fsLocal.get(), bindingPath.native(), L"local-win32", hostBindingFlags);
        if (! require(boundAuthority.state == FileOperations::ObjectBindingState::Bound && SUCCEEDED(boundAuthority.status) &&
                          boundAuthority.authority.boundObject && ! boundAuthority.authority.identity.objectId.empty() &&
                          boundAuthority.authority.identity.revisionId.empty() &&
                          boundAuthority.authority.identity.pathProfileId == L"local-win32" &&
                          boundAuthority.authority.kind == FILESYSTEM_BOUND_REGULAR_FILE,
                      L"Host binding should copy a complete exact local identity snapshot."))
        {
            return true;
        }

        FileOperations::ObjectRevalidationResult sameAuthority = FileOperations::RevalidateObjectAuthority(
            state.fsLocal.get(), bindingAlias.native(), L"local-win32", hostBindingFlags, boundAuthority.authority);
        if (! require(sameAuthority.state == FileOperations::ObjectRevalidationState::Same && SUCCEEDED(sameAuthority.status),
                      L"Host revalidation should recognize a hard-link alias as the same object."))
        {
            return true;
        }

        FileOperations::RenamePlan duplicateObjectRename{};
        duplicateObjectRename.origin   = FileOperations::RenameOrigin::BatchRename;
        duplicateObjectRename.endpoint = validationEndpoint;
        duplicateObjectRename.endpoint.profileId = L"local-win32";
        duplicateObjectRename.finalMappings = {
            {
                .source = {
                    .providerPath = bindingPath.native(),
                    .ingressSnapshot = boundAuthority.authority.identity,
                },
                .finalLeafName = L"renamed-one.bin",
            },
            {
                .source = {
                    .providerPath = bindingAlias.native(),
                    .ingressSnapshot = sameAuthority.current.identity,
                },
                .finalLeafName = L"renamed-two.bin",
            },
        };
        if (! require(FileOperations::ValidatePlan(FileOperations::FileOperationPlan{duplicateObjectRename}, &validationBucket) == E_INVALIDARG &&
                          validationBucket == FileOperations::PlanRejectionBucket::InvalidRename,
                      L"Batch Rename validation must reject two paths bound to the same physical source object."))
        {
            return true;
        }

        if (! require(DeleteFileW(bindingPath.c_str()) != FALSE && SelfTest::WriteTextFile(bindingPath, "replacement"),
                      L"Host revalidation fixture should replace the original pathname while retaining its hard-link object."))
        {
            return true;
        }
        FileOperations::ObjectRevalidationResult changedAuthority = FileOperations::RevalidateObjectAuthority(
            state.fsLocal.get(), bindingPath.native(), L"local-win32", hostBindingFlags, boundAuthority.authority);
        if (! require(changedAuthority.state == FileOperations::ObjectRevalidationState::Changed && SUCCEEDED(changedAuthority.status),
                      L"Host revalidation should classify same-path replacement as changed."))
        {
            return true;
        }

        if (! require(DeleteFileW(bindingPath.c_str()) != FALSE, L"Host revalidation fixture should remove the replacement pathname."))
        {
            return true;
        }
        FileOperations::ObjectRevalidationResult missingAuthority = FileOperations::RevalidateObjectAuthority(
            state.fsLocal.get(), bindingPath.native(), L"local-win32", hostBindingFlags, boundAuthority.authority);
        if (! require(missingAuthority.state == FileOperations::ObjectRevalidationState::Missing && FAILED(missingAuthority.status),
                      L"Host revalidation should keep disappearance distinct from changed and indeterminate."))
        {
            return true;
        }

        FileOperations::ObjectBindingResult unsupportedAuthority =
            FileOperations::BindObjectAuthority(state.fsDummy.get(), L"/unsupported/object", L"dummy-local", hostBindingFlags);
        if (! require(unsupportedAuthority.state == FileOperations::ObjectBindingState::Unsupported && FAILED(unsupportedAuthority.status),
                      L"A provider without optional object binding should classify as unsupported, not missing or bound."))
        {
            return true;
        }

        const FileOperations::QualifiedEndpoint dummyValidationEndpoint{
            .pluginId   = L"builtin/file-system-dummy",
            .instanceId = L"host/default",
            .profileId  = L"dummy-local",
            .rootId     = L"dummy-root",
            .pathIdentity = FileSystemPathIdentity{
                .pathTextStableIdentity = true,
                .componentComparison    = FileSystemPathComponentComparison::OrdinalIgnoreCase,
                .preferredSeparator     = L'\\',
                .acceptedSeparators     = L"\\/",
                .casePreserving         = true,
                .caseOnlyRename         = FileSystemPathCaseOnlyRename::Supported,
            },
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };

        constexpr uint64_t kR0dPerfIterations = 4'096u;
        FileOperations::TransferPlan r0dPerfPlan{};
        r0dPerfPlan.intent              = FileOperations::TransferIntent::Move;
        r0dPerfPlan.strategy            = FileOperations::OperationStrategy::Native;
        r0dPerfPlan.sourceEndpoint      = dummyValidationEndpoint;
        r0dPerfPlan.destinationEndpoint = dummyValidationEndpoint;
        r0dPerfPlan.selectedItems       = {{.providerPath = L"/r0d-perf-source.bin"}};
        r0dPerfPlan.destination.providerFolderPath = L"/r0d-perf-destination";
        FolderWindow::FileOperationState::Task r0dPerfTask(*state.fileOps);
        r0dPerfTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(r0dPerfPlan)}}));
        r0dPerfTask._operation   = FILESYSTEM_MOVE;
        r0dPerfTask._sourcePaths = {std::filesystem::path(L"/r0d-perf-source.bin")};

        const auto r0dPerfStarted = std::chrono::steady_clock::now();
        bool r0dPerfValid          = true;
        uint64_t r0dAdmissionChecksum = 0u;
        for (uint64_t iteration = 0u; iteration < kR0dPerfIterations; ++iteration)
        {
            ProviderCapabilitySnapshot perfCapabilities{};
            std::wstring perfReason;
            r0dPerfValid = TryReadProviderCapabilities(state.fsDummy.get(), perfCapabilities, perfReason) && r0dPerfValid;
            r0dAdmissionChecksum += CanSameFileSystemOperation(state.fsDummy, L"/", FILESYSTEM_DELETE, kPluginIdDummy) ? 1u : 0u;
            r0dAdmissionChecksum += CanSameFileSystemOperation(state.fsDummy, L"/", FILESYSTEM_RENAME, kPluginIdDummy) ? 1u : 0u;

            r0dPerfTask.InitializeSourceItemResultBuilders();
            r0dPerfTask.MarkSourceItemsMutationPossible();
            r0dPerfTask._sourceItemResultBuilders.front().status = S_OK;
            const HRESULT classificationHr = r0dPerfTask.FinalizeTypedItemResults(S_OK);
            const FileOperations::FileOperationItemResult* classified =
                r0dPerfTask._sourceItemResultBuilders.size() == 1u && r0dPerfTask._sourceItemResultBuilders.front().terminal.has_value()
                ? std::addressof(r0dPerfTask._sourceItemResultBuilders.front().terminal.value())
                : nullptr;
            r0dPerfValid = classificationHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) && classified != nullptr &&
                classified->completion == FileOperations::ItemCompletion::Indeterminate && r0dPerfValid;
        }
        const uint64_t r0dPerfElapsedUs = Debug::Perf::ElapsedUs(r0dPerfStarted);
        constexpr uint64_t kR0dCapabilityQueryCount = kR0dPerfIterations * 3u;
        constexpr uint64_t kR0dAdmissionDecisionCount = kR0dPerfIterations * 2u;
        Debug::Perf::Emit(L"FileOps.SelfTest.R0d.CapabilityReceiptHonestyUs",
                          L"dummy-4096-capability-admission-receipt",
                          r0dPerfElapsedUs,
                          kR0dCapabilityQueryCount,
                          kR0dPerfIterations,
                          r0dPerfValid ? S_OK : E_FAIL);
        Debug::Perf::Emit(L"FileOps.SelfTest.R0d.CapabilityQueryCount",
                          L"dummy-4096-capability-admission-receipt",
                          kR0dCapabilityQueryCount,
                          kR0dPerfIterations,
                          0u,
                          r0dPerfValid ? S_OK : E_FAIL);
        Debug::Perf::Emit(L"FileOps.SelfTest.R0d.AdmissionDecisionCount",
                          L"dummy-4096-capability-admission-receipt",
                          kR0dAdmissionDecisionCount,
                          r0dAdmissionChecksum,
                          0u,
                          r0dPerfValid ? S_OK : E_FAIL);
        Debug::Perf::Emit(L"FileOps.SelfTest.R0d.ReceiptClassificationCount",
                          L"dummy-4096-capability-admission-receipt",
                          kR0dPerfIterations,
                          kR0dPerfIterations,
                          0u,
                          r0dPerfValid ? S_OK : E_FAIL);
        if (! require(r0dPerfValid && r0dPerfElapsedUs < 2'000'000u,
                      L"R0d capability/admission/receipt classification metric must remain valid and below two seconds."))
        {
            return true;
        }

        constexpr std::wstring_view r0dMoveSource      = L"/r0d-null-receipt-move/source.bin";
        constexpr std::wstring_view r0dMoveDestination = L"/r0d-null-receipt-move/destination";
        constexpr std::wstring_view r0dDeleteSource    = L"/r0d-null-receipt-delete/source.bin";
        static_cast<void>(state.fsDummy->DeleteItem(L"/r0d-null-receipt-move", FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr));
        static_cast<void>(state.fsDummy->DeleteItem(L"/r0d-null-receipt-delete", FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr));
        if (! require(EnsureDummyFolderExists(state.fsDummy.get(), L"/r0d-null-receipt-move") &&
                          EnsureDummyFolderExists(state.fsDummy.get(), r0dMoveDestination) &&
                          EnsureDummyFolderExists(state.fsDummy.get(), L"/r0d-null-receipt-delete") &&
                          DummyWriteTextFile(state.fsDummy.get(), r0dMoveSource, "move") &&
                          DummyWriteTextFile(state.fsDummy.get(), r0dDeleteSource, "delete"),
                      L"R0d null-receipt fixtures should create deterministic Dummy objects."))
        {
            return true;
        }

        FileOperations::TransferPlan r0dMovePlan{};
        r0dMovePlan.intent              = FileOperations::TransferIntent::Move;
        r0dMovePlan.strategy            = FileOperations::OperationStrategy::Native;
        r0dMovePlan.sourceEndpoint      = dummyValidationEndpoint;
        r0dMovePlan.destinationEndpoint = dummyValidationEndpoint;
        r0dMovePlan.selectedItems       = {{.providerPath = std::wstring(r0dMoveSource)}};
        r0dMovePlan.destination.providerFolderPath = std::wstring(r0dMoveDestination);
        FolderWindow::FileOperationState::Task r0dMoveTask(*state.fileOps);
        r0dMoveTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(r0dMovePlan)}}));
        r0dMoveTask._operation                = FILESYSTEM_MOVE;
        r0dMoveTask._executionMode            = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        r0dMoveTask._fileSystem               = state.fsDummy;
        r0dMoveTask._sourcePaths              = {std::filesystem::path(r0dMoveSource)};
        r0dMoveTask._destinationFolder        = std::filesystem::path(r0dMoveDestination);
        r0dMoveTask._sourcePathAttributesHint = {FILE_ATTRIBUTE_NORMAL};
        r0dMoveTask.InitializeSourceItemResultBuilders();
        r0dMoveTask.MarkSourceItemsMutationPossible();
        const HRESULT r0dMoveExecutionHr = r0dMoveTask.ExecuteOperation();
        const HRESULT r0dMoveFinalHr     = r0dMoveTask.FinalizeTypedItemResults(r0dMoveExecutionHr);

        FileOperations::DeletePlan r0dDeletePlan{};
        r0dDeletePlan.endpoint      = dummyValidationEndpoint;
        r0dDeletePlan.mode          = FileOperations::DeleteMode::Recycle;
        r0dDeletePlan.selectedItems = {{.providerPath = std::wstring(r0dDeleteSource)}};
        FolderWindow::FileOperationState::Task r0dDeleteTask(*state.fileOps);
        r0dDeleteTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(r0dDeletePlan)}}));
        r0dDeleteTask._operation                = FILESYSTEM_DELETE;
        r0dDeleteTask._executionMode            = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        r0dDeleteTask._fileSystem               = state.fsDummy;
        r0dDeleteTask._sourcePaths              = {std::filesystem::path(r0dDeleteSource)};
        r0dDeleteTask._flags                    = FILESYSTEM_FLAG_USE_RECYCLE_BIN;
        r0dDeleteTask._sourcePathAttributesHint = {FILE_ATTRIBUTE_NORMAL};
        r0dDeleteTask.InitializeSourceItemResultBuilders();
        r0dDeleteTask.MarkSourceItemsMutationPossible();
        const HRESULT r0dDeleteExecutionHr = r0dDeleteTask.ExecuteOperation();
        const HRESULT r0dDeleteFinalHr     = r0dDeleteTask.FinalizeTypedItemResults(r0dDeleteExecutionHr);

        const auto indeterminateResult = [](const FolderWindow::FileOperationState::Task& task,
                                            FileOperations::PublicationState expectedPublication) noexcept
        {
            return task._sourceItemResultBuilders.size() == 1u && task._sourceItemResultBuilders.front().terminal.has_value() &&
                task._sourceItemResultBuilders.front().terminal->publication == expectedPublication &&
                task._sourceItemResultBuilders.front().terminal->sourceDisposition == FileOperations::SourceDisposition::Unknown &&
                task._sourceItemResultBuilders.front().terminal->completion == FileOperations::ItemCompletion::Indeterminate &&
                task._sourceItemResultBuilders.front().terminal->status == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        };
        if (! require(r0dMoveExecutionHr == S_OK && r0dDeleteExecutionHr == S_OK &&
                          r0dMoveFinalHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) &&
                          r0dDeleteFinalHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) &&
                          indeterminateResult(r0dMoveTask, FileOperations::PublicationState::Unknown) &&
                          indeterminateResult(r0dDeleteTask, FileOperations::PublicationState::NotAttempted),
                      L"A successful destructive Native call with no mutation receipt must be Indeterminate immediately, never Completed."))
        {
            return true;
        }

        constexpr std::wstring_view dummyInlineRenameRoot   = L"/inline-rename-unbound";
        constexpr std::wstring_view dummyInlineRenameSource = L"/inline-rename-unbound/source.txt";
        constexpr std::wstring_view dummyInlineRenameTarget = L"/inline-rename-unbound/renamed.txt";
        static_cast<void>(state.fsDummy->DeleteItem(std::wstring(dummyInlineRenameRoot).c_str(),
                                                    FILESYSTEM_FLAG_RECURSIVE,
                                                    nullptr,
                                                    nullptr,
                                                    nullptr));
        if (! require(EnsureDummyFolderExists(state.fsDummy.get(), dummyInlineRenameRoot) &&
                          DummyWriteTextFile(state.fsDummy.get(), dummyInlineRenameSource, "unbound-rename"),
                      L"Unbound Inline Rename fixture should create its deterministic source."))
        {
            return true;
        }
        if (! require(state.fileOps != nullptr, L"Unbound Inline Rename fixture requires FileOperationState."))
        {
            return true;
        }
        FileOperations::RenamePlan unboundInlineRename{};
        unboundInlineRename.origin   = FileOperations::RenameOrigin::InlineRename;
        unboundInlineRename.endpoint = dummyValidationEndpoint;
        unboundInlineRename.finalMappings = {{
            .source        = {.providerPath = std::wstring(dummyInlineRenameSource)},
            .finalLeafName = L"renamed.txt",
        }};
        FolderWindow::FileOperationState::Task unboundInlineRenameTask(*state.fileOps);
        unboundInlineRenameTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(unboundInlineRename)}}));
        unboundInlineRenameTask._operation     = FILESYSTEM_RENAME;
        unboundInlineRenameTask._executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        unboundInlineRenameTask._fileSystem    = state.fsDummy;
        const HRESULT unboundInlineRenameHr = unboundInlineRenameTask.PrepareMutationInterlockScopes();
        wil::com_ptr<IFileSystemIO> unboundDummyIo;
        std::string retainedUnboundSource;
        std::string unexpectedUnboundTarget;
        if (! require(unboundInlineRenameHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) &&
                          SUCCEEDED(state.fsDummy->QueryInterface(IID_PPV_ARGS(unboundDummyIo.addressof()))) && unboundDummyIo &&
                          ReadFileTextFsIo(unboundDummyIo, dummyInlineRenameSource, retainedUnboundSource) &&
                          retainedUnboundSource == "unbound-rename" &&
                          ! ReadFileTextFsIo(unboundDummyIo, dummyInlineRenameTarget, unexpectedUnboundTarget),
                      L"A provider advertising base Rename without object binding must fail interlock preparation before mutation."))
        {
            return true;
        }

        FileOperations::TransferMutationGuard unsupportedBindingDescendant = FileOperations::PrepareTransferMutationGuard(state.fsDummy.get(),
                                                                                                                             state.fsDummy.get(),
                                                                                                                             dummyValidationEndpoint,
                                                                                                                             dummyValidationEndpoint,
                                                                                                                             FileOperations::TransferIntent::Move,
                                                                                                                             L"/Source",
                                                                                                                             L"\\source\\child");
        if (! require(unsupportedBindingDescendant.state == FileOperations::TransferSafetyState::DestinationInsideSource &&
                          FAILED(unsupportedBindingDescendant.status),
                      L"A stable provider path profile must stop a destination-inside-source transfer even when object binding is unsupported."))
        {
            return true;
        }

        const std::filesystem::path guardRoot               = state.tempRoot / L"transfer-mutation-guard";
        const std::filesystem::path guardSource             = guardRoot / L"source.bin";
        const std::filesystem::path guardAlias              = guardRoot / L"source-alias.bin";
        const std::filesystem::path guardDestinationRoot    = guardRoot / L"destination";
        const std::filesystem::path guardMissingDestination = guardDestinationRoot / L"copy.bin";
        if (! require(SelfTest::EnsureDirectory(guardDestinationRoot) && SelfTest::WriteTextFile(guardSource, "guard-source") &&
                          CreateHardLinkW(guardAlias.c_str(), guardSource.c_str(), nullptr) != FALSE,
                      L"Transfer mutation-guard fixture should create source, destination root, and hard-link alias."))
        {
            return true;
        }

        FileOperations::TransferMutationGuard directSameObject = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                                state.fsLocal.get(),
                                                                                                                localValidationEndpoint,
                                                                                                                localValidationEndpoint,
                                                                                                                FileOperations::TransferIntent::Copy,
                                                                                                                guardSource.native(),
                                                                                                                guardSource.native());
        if (! require(directSameObject.state == FileOperations::TransferSafetyState::SameObject && directSameObject.samePathText &&
                          FAILED(directSameObject.status),
                      L"A direct same-path Copy must be classified as the same object so execution can select Keep Both before mutation."))
        {
            return true;
        }

        FileOperations::TransferMutationGuard hardLinkSameObject = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                                  state.fsLocal.get(),
                                                                                                                  localValidationEndpoint,
                                                                                                                  localValidationEndpoint,
                                                                                                                  FileOperations::TransferIntent::Copy,
                                                                                                                  guardSource.native(),
                                                                                                                  guardAlias.native());
        if (! require(hardLinkSameObject.state == FileOperations::TransferSafetyState::SameObject && ! hardLinkSameObject.samePathText &&
                          FAILED(hardLinkSameObject.status),
                      L"A hard-link destination alias must stop as the same object and must not be converted into same-path Keep Both."))
        {
            return true;
        }

        const auto requireSameObjectAlias = [&](const std::filesystem::path& aliasPath,
                                                bool expectedSamePathText,
                                                std::wstring_view label) noexcept -> bool
        {
            FileOperations::TransferMutationGuard aliasGuard = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                               state.fsLocal.get(),
                                                                                                               localValidationEndpoint,
                                                                                                               localValidationEndpoint,
                                                                                                               FileOperations::TransferIntent::Copy,
                                                                                                               guardSource.native(),
                                                                                                               aliasPath.native());
            return require(aliasGuard.state == FileOperations::TransferSafetyState::SameObject &&
                               aliasGuard.samePathText == expectedSamePathText && FAILED(aliasGuard.status),
                           std::format(L"The {} alias must resolve to the exact source object before provider mutation.", label));
        };

        std::wstring caseAliasLeaf = guardSource.filename().native();
        std::ranges::transform(caseAliasLeaf, caseAliasLeaf.begin(), [](wchar_t ch) noexcept { return static_cast<wchar_t>(towupper(ch)); });
        const std::filesystem::path caseAlias = guardSource.parent_path() / caseAliasLeaf;
        if (! requireSameObjectAlias(caseAlias, true, L"case-only"))
        {
            return true;
        }

        const std::filesystem::path junctionAliasRoot = state.tempRoot / L"transfer-mutation-guard-junction-alias";
        if (! require(TryCreateJunction(junctionAliasRoot, guardRoot), L"Transfer mutation-guard fixture should create a junction alias root.") ||
            ! requireSameObjectAlias(junctionAliasRoot / guardSource.filename(), false, L"junction/mount"))
        {
            return true;
        }

        std::array<wchar_t, MAX_PATH> volumeRoot{};
        const std::optional<std::wstring> volumeGuidRoot = TryGetVolumeGuidPathForPath(guardSource);
        if (GetVolumePathNameW(guardSource.c_str(), volumeRoot.data(), static_cast<DWORD>(volumeRoot.size())) != FALSE &&
            volumeGuidRoot.has_value())
        {
            const std::wstring sourceText = guardSource.native();
            const std::wstring rootText(volumeRoot.data());
            if (sourceText.size() >= rootText.size() && OrdinalString::StartsWithNoCase(sourceText, rootText))
            {
                const std::filesystem::path volumeGuidAlias = volumeGuidRoot.value() + sourceText.substr(rootText.size());
                if (! requireSameObjectAlias(volumeGuidAlias, false, L"volume-GUID"))
                {
                    return true;
                }
            }
        }
        else
        {
            AppendLog(L"Provider identity matrix: volume-GUID alias unavailable; environment-gated case skipped.");
        }

        const std::filesystem::path longAliasSource = guardRoot / L"source object requiring an eight dot three alias.bin";
        if (! require(SelfTest::WriteTextFile(longAliasSource, "short-name-alias"), L"Short-name alias fixture should create the long-name source."))
        {
            return true;
        }
        const DWORD shortPathChars = GetShortPathNameW(longAliasSource.c_str(), nullptr, 0u);
        if (shortPathChars > 0u)
        {
            std::vector<wchar_t> shortPath(shortPathChars);
            const DWORD written = GetShortPathNameW(longAliasSource.c_str(), shortPath.data(), static_cast<DWORD>(shortPath.size()));
            if (written > 0u && written < shortPath.size() &&
                ! OrdinalString::EqualsNoCase(longAliasSource.native(), std::wstring_view(shortPath.data(), written)))
            {
                FileOperations::TransferMutationGuard shortNameGuard = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                                      state.fsLocal.get(),
                                                                                                                      localValidationEndpoint,
                                                                                                                      localValidationEndpoint,
                                                                                                                      FileOperations::TransferIntent::Copy,
                                                                                                                      longAliasSource.native(),
                                                                                                                      std::wstring_view(shortPath.data(), written));
                if (! require(shortNameGuard.state == FileOperations::TransferSafetyState::SameObject && ! shortNameGuard.samePathText &&
                                  FAILED(shortNameGuard.status),
                              L"An available 8.3 alias must resolve to the exact long-name source object."))
                {
                    return true;
                }
            }
            else
            {
                AppendLog(L"Provider identity matrix: filesystem returned no distinct 8.3 alias; environment-gated case skipped.");
            }
        }
        else
        {
            AppendLog(L"Provider identity matrix: 8.3 aliases are unavailable; environment-gated case skipped.");
        }

        wchar_t substLetter = L'\0';
        const DWORD logicalDrives = GetLogicalDrives();
        for (wchar_t candidate = L'Y'; candidate >= L'D'; --candidate)
        {
            const DWORD mask = 1u << static_cast<unsigned int>(candidate - L'A');
            if ((logicalDrives & mask) == 0u)
            {
                substLetter = candidate;
                break;
            }
        }
        if (substLetter != L'\0')
        {
            const std::wstring deviceName{substLetter, L':'};
            const std::wstring rawTarget = std::wstring(L"\\??\\") + guardRoot.native();
            constexpr DWORD defineFlags = DDD_RAW_TARGET_PATH | DDD_NO_BROADCAST_SYSTEM;
            if (DefineDosDeviceW(defineFlags, deviceName.c_str(), rawTarget.c_str()) != FALSE)
            {
                auto removeSubst = wil::scope_exit([&]() noexcept
                {
                    constexpr DWORD removeFlags =
                        DDD_REMOVE_DEFINITION | DDD_EXACT_MATCH_ON_REMOVE | DDD_RAW_TARGET_PATH | DDD_NO_BROADCAST_SYSTEM;
                    static_cast<void>(DefineDosDeviceW(removeFlags, deviceName.c_str(), rawTarget.c_str()));
                });
                const std::filesystem::path substAlias = deviceName + L"\\" + guardSource.filename().native();
                if (! requireSameObjectAlias(substAlias, false, L"SUBST"))
                {
                    return true;
                }
            }
            else
            {
                AppendLog(L"Provider identity matrix: SUBST namespace unavailable; environment-gated case skipped.");
            }
        }
        else
        {
            AppendLog(L"Provider identity matrix: no unused drive letter for SUBST alias; environment-gated case skipped.");
        }

        const std::filesystem::path uncAlias = LoopbackShareAlias(guardSource);
        if (! uncAlias.empty())
        {
            FileOperations::QualifiedEndpoint driveEndpoint{};
            FileOperations::QualifiedEndpoint uncEndpoint{};
            const bool driveQualified = FileOperations::TryQualifyEndpointForSelfTest(state.fsLocal,
                                                                                       guardSource.native(),
                                                                                       FILESYSTEM_COPY,
                                                                                       L"builtin/file-system",
                                                                                       L"host/default",
                                                                                       driveEndpoint);
            const bool uncQualified = FileOperations::TryQualifyEndpointForSelfTest(state.fsLocal,
                                                                                     uncAlias.native(),
                                                                                     FILESYSTEM_COPY,
                                                                                     L"builtin/file-system",
                                                                                     L"host/default",
                                                                                     uncEndpoint);
            if (! require(driveQualified && uncQualified &&
                              uncEndpoint.profileId == L"local-win32-smb" &&
                              uncEndpoint.cancellationRouteClass == FileOperations::CancellationRouteClass::Bounded &&
                              uncEndpoint.providerWatchdogTimeoutMs == 0u &&
                              CanSameFileSystemOperation(state.fsLocal,
                                                         uncAlias.native(),
                                                         FILESYSTEM_COPY,
                                                         L"builtin/file-system"),
                          L"Local UNC must stay a distinct route profile that R0f-SMB admits as bounded without opening the share."))
            {
                return true;
            }
            if (GetFileAttributesW(uncAlias.c_str()) != INVALID_FILE_ATTRIBUTES)
            {
                FileOperations::QualifiedEndpoint uncIdentityEndpoint = uncEndpoint;
                // Cancellation route classification is intentionally distinct for Local SMB,
                // while the lower-level Local object-binding contract still uses one Win32
                // identity profile. Isolate that existing identity proof from admission here.
                uncIdentityEndpoint.profileId = driveEndpoint.profileId;
                uncIdentityEndpoint.cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded;
                FileOperations::TransferMutationGuard uncAliasGuard = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                                      state.fsLocal.get(),
                                                                                                                      driveEndpoint,
                                                                                                                      uncIdentityEndpoint,
                                                                                                                      FileOperations::TransferIntent::Copy,
                                                                                                                      guardSource.native(),
                                                                                                                      uncAlias.native());
                if (! require(driveEndpoint.rootId != uncEndpoint.rootId &&
                                  uncAliasGuard.state == FileOperations::TransferSafetyState::SameObject &&
                                  ! uncAliasGuard.samePathText && FAILED(uncAliasGuard.status),
                              L"The Local Win32 identity profile must still detect a drive/UNC alias independently of cancellation admission."))
                {
                    return true;
                }
            }
            else
            {
                AppendLog(L"Provider identity matrix: localhost administrative-share UNC alias unavailable; environment-gated case skipped.");
            }
        }
        else
        {
            AppendLog(L"Provider identity matrix: drive/UNC alias requires a drive-letter sandbox; environment-gated case skipped.");
        }

        FileOperations::TransferMutationGuard readyGuard = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                          state.fsLocal.get(),
                                                                                                          localValidationEndpoint,
                                                                                                          localValidationEndpoint,
                                                                                                          FileOperations::TransferIntent::Copy,
                                                                                                          guardSource.native(),
                                                                                                          guardMissingDestination.native());
        if (! require(readyGuard.state == FileOperations::TransferSafetyState::Ready && readyGuard.destinationWasMissing &&
                          SUCCEEDED(readyGuard.status),
                      L"A distinct source and missing destination should retain exact authority and remain ready."))
        {
            return true;
        }

        HRESULT readyRevalidationStatus = E_UNEXPECTED;
        const FileOperations::TransferSafetyState readyRevalidation = FileOperations::RevalidateTransferMutationGuard(state.fsLocal.get(),
                                                                                                                        state.fsLocal.get(),
                                                                                                                        localValidationEndpoint,
                                                                                                                        localValidationEndpoint,
                                                                                                                        FileOperations::TransferIntent::Copy,
                                                                                                                        guardSource.native(),
                                                                                                                        guardMissingDestination.native(),
                                                                                                                        readyGuard,
                                                                                                                        readyRevalidationStatus);
        if (! require(readyRevalidation == FileOperations::TransferSafetyState::Ready && SUCCEEDED(readyRevalidationStatus),
                      L"Unchanged retained source and destination authority should revalidate at the provider boundary."))
        {
            return true;
        }

        FileOperations::TransferMutationGuard sameFolderMoveGuard = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                                   state.fsLocal.get(),
                                                                                                                   localValidationEndpoint,
                                                                                                                   localValidationEndpoint,
                                                                                                                   FileOperations::TransferIntent::Move,
                                                                                                                   guardSource.native(),
                                                                                                                   (guardRoot / L"renamed.bin").native());
        if (! require(sameFolderMoveGuard.state == FileOperations::TransferSafetyState::SameFolderMove && FAILED(sameFolderMoveGuard.status),
                      L"Exact parent identity must reject a same-folder Move even when the destination leaf differs."))
        {
            return true;
        }

        const std::filesystem::path guardTree      = guardRoot / L"tree";
        const std::filesystem::path guardTreeChild = guardTree / L"child";
        if (! require(SelfTest::EnsureDirectory(guardTreeChild), L"Transfer containment fixture should create an in-tree destination ancestor."))
        {
            return true;
        }
        FileOperations::TransferMutationGuard subtreeGuard = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                            state.fsLocal.get(),
                                                                                                            localValidationEndpoint,
                                                                                                            localValidationEndpoint,
                                                                                                            FileOperations::TransferIntent::Copy,
                                                                                                            guardTree.native(),
                                                                                                            (guardTreeChild / L"copy").native());
        if (! require(subtreeGuard.state == FileOperations::TransferSafetyState::DestinationInsideSource && FAILED(subtreeGuard.status),
                      L"No-follow ancestor identity must reject a directory destination discovered inside its source tree."))
        {
            return true;
        }

        const std::filesystem::path guardOutsideTarget = guardRoot / L"outside-target";
        const std::filesystem::path guardDestinationLink = guardDestinationRoot / L"linked-parent";
        if (! require(SelfTest::EnsureDirectory(guardOutsideTarget) && TryCreateJunction(guardDestinationLink, guardOutsideTarget),
                      L"Transfer containment fixture should create a destination junction ancestor."))
        {
            return true;
        }
        FileOperations::TransferMutationGuard linkAncestorGuard = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                                 state.fsLocal.get(),
                                                                                                                 localValidationEndpoint,
                                                                                                                 localValidationEndpoint,
                                                                                                                 FileOperations::TransferIntent::Copy,
                                                                                                                 guardTree.native(),
                                                                                                                 (guardDestinationLink / L"copy").native());
        if (! require(linkAncestorGuard.state == FileOperations::TransferSafetyState::AncestryLink && FAILED(linkAncestorGuard.status),
                      L"A no-follow destination junction ancestor must be an explicit conflict, never a traversed containment shortcut."))
        {
            return true;
        }

        FileOperations::TransferMutationGuard destinationSwapGuard = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                                    state.fsLocal.get(),
                                                                                                                    localValidationEndpoint,
                                                                                                                    localValidationEndpoint,
                                                                                                                    FileOperations::TransferIntent::Copy,
                                                                                                                    guardSource.native(),
                                                                                                                    guardMissingDestination.native());
        if (! require(destinationSwapGuard.state == FileOperations::TransferSafetyState::Ready &&
                          SelfTest::WriteTextFile(guardMissingDestination, "destination-race"),
                      L"Destination-swap fixture should prepare a missing destination and then publish a racing object."))
        {
            return true;
        }
        HRESULT destinationSwapStatus = E_UNEXPECTED;
        const FileOperations::TransferSafetyState destinationSwapState = FileOperations::RevalidateTransferMutationGuard(state.fsLocal.get(),
                                                                                                                          state.fsLocal.get(),
                                                                                                                          localValidationEndpoint,
                                                                                                                          localValidationEndpoint,
                                                                                                                          FileOperations::TransferIntent::Copy,
                                                                                                                          guardSource.native(),
                                                                                                                          guardMissingDestination.native(),
                                                                                                                          destinationSwapGuard,
                                                                                                                          destinationSwapStatus);
        if (! require(destinationSwapState == FileOperations::TransferSafetyState::DestinationChanged && FAILED(destinationSwapStatus),
                      L"A destination appearing after classification must fail closed at final mutation-boundary revalidation."))
        {
            return true;
        }

        const std::filesystem::path finalLinkSwap = guardDestinationRoot / L"final-link-swap";
        FileOperations::TransferMutationGuard finalLinkSwapGuard = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                                  state.fsLocal.get(),
                                                                                                                  localValidationEndpoint,
                                                                                                                  localValidationEndpoint,
                                                                                                                  FileOperations::TransferIntent::Copy,
                                                                                                                  guardSource.native(),
                                                                                                                  finalLinkSwap.native());
        if (! require(finalLinkSwapGuard.state == FileOperations::TransferSafetyState::Ready &&
                          TryCreateJunction(finalLinkSwap, guardOutsideTarget),
                      L"Final-link swap fixture should classify a missing destination before replacing it with a junction."))
        {
            return true;
        }
        HRESULT finalLinkSwapStatus = E_UNEXPECTED;
        const FileOperations::TransferSafetyState finalLinkSwapState = FileOperations::RevalidateTransferMutationGuard(state.fsLocal.get(),
                                                                                                                         state.fsLocal.get(),
                                                                                                                         localValidationEndpoint,
                                                                                                                         localValidationEndpoint,
                                                                                                                         FileOperations::TransferIntent::Copy,
                                                                                                                         guardSource.native(),
                                                                                                                         finalLinkSwap.native(),
                                                                                                                         finalLinkSwapGuard,
                                                                                                                         finalLinkSwapStatus);
        if (! require(finalLinkSwapState == FileOperations::TransferSafetyState::DestinationChanged && FAILED(finalLinkSwapStatus),
                      L"An exact destination link appearing after classification must be a destination change and fail closed before mutation."))
        {
            return true;
        }

        const std::filesystem::path ancestorSwapParent = guardDestinationRoot / L"ancestor-swap-parent";
        const std::filesystem::path ancestorSwapDestination = ancestorSwapParent / L"copy.bin";
        if (! require(SelfTest::EnsureDirectory(ancestorSwapParent), L"Ancestor-swap fixture should create its original destination parent."))
        {
            return true;
        }
        FileOperations::TransferMutationGuard ancestorSwapGuard = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                                  state.fsLocal.get(),
                                                                                                                  localValidationEndpoint,
                                                                                                                  localValidationEndpoint,
                                                                                                                  FileOperations::TransferIntent::Copy,
                                                                                                                  guardSource.native(),
                                                                                                                  ancestorSwapDestination.native());
        if (! require(ancestorSwapGuard.state == FileOperations::TransferSafetyState::Ready &&
                          RemoveDirectoryW(ancestorSwapParent.c_str()) != FALSE &&
                          TryCreateJunction(ancestorSwapParent, guardOutsideTarget),
                      L"Ancestor-swap fixture should replace a retained destination parent with a junction."))
        {
            return true;
        }
        HRESULT ancestorSwapStatus = E_UNEXPECTED;
        const FileOperations::TransferSafetyState ancestorSwapState = FileOperations::RevalidateTransferMutationGuard(state.fsLocal.get(),
                                                                                                                        state.fsLocal.get(),
                                                                                                                        localValidationEndpoint,
                                                                                                                        localValidationEndpoint,
                                                                                                                        FileOperations::TransferIntent::Copy,
                                                                                                                        guardSource.native(),
                                                                                                                        ancestorSwapDestination.native(),
                                                                                                                        ancestorSwapGuard,
                                                                                                                        ancestorSwapStatus);
        if (! require(ancestorSwapState != FileOperations::TransferSafetyState::Ready && FAILED(ancestorSwapStatus),
                      L"A retained destination ancestor replaced by a link must fail closed before provider mutation."))
        {
            return true;
        }

        const std::filesystem::path guardReplacementSource      = guardRoot / L"replacement-source.bin";
        const std::filesystem::path guardReplacementDestination = guardDestinationRoot / L"replacement-copy.bin";
        if (! require(SelfTest::WriteTextFile(guardReplacementSource, "original-source"),
                      L"Source-swap fixture should create the original source object."))
        {
            return true;
        }
        FileOperations::TransferMutationGuard sourceSwapGuard = FileOperations::PrepareTransferMutationGuard(state.fsLocal.get(),
                                                                                                               state.fsLocal.get(),
                                                                                                               localValidationEndpoint,
                                                                                                               localValidationEndpoint,
                                                                                                               FileOperations::TransferIntent::Copy,
                                                                                                               guardReplacementSource.native(),
                                                                                                               guardReplacementDestination.native());
        if (! require(sourceSwapGuard.state == FileOperations::TransferSafetyState::Ready &&
                          DeleteFileW(guardReplacementSource.c_str()) != FALSE &&
                          SelfTest::WriteTextFile(guardReplacementSource, "replacement-source"),
                      L"Source-swap fixture should replace the classified source pathname with a different object."))
        {
            return true;
        }
        HRESULT sourceSwapStatus = E_UNEXPECTED;
        const FileOperations::TransferSafetyState sourceSwapState = FileOperations::RevalidateTransferMutationGuard(state.fsLocal.get(),
                                                                                                                     state.fsLocal.get(),
                                                                                                                     localValidationEndpoint,
                                                                                                                     localValidationEndpoint,
                                                                                                                     FileOperations::TransferIntent::Copy,
                                                                                                                     guardReplacementSource.native(),
                                                                                                                     guardReplacementDestination.native(),
                                                                                                                     sourceSwapGuard,
                                                                                                                     sourceSwapStatus);
        if (! require(sourceSwapState == FileOperations::TransferSafetyState::SourceChanged && FAILED(sourceSwapStatus),
                      L"A source replacement after classification must fail closed before provider mutation."))
        {
            return true;
        }

        const FileOperations::ObjectBindingResult nullSuccess =
            FileOperations::ValidateSuccessfulBoundObjectForSelfTest(nullptr, L"local-win32");
        if (! require(nullSuccess.state == FileOperations::ObjectBindingState::ProviderContractViolation && FAILED(nullSuccess.status),
                      L"A successful provider call with a null bound object must be a contract violation."))
        {
            return true;
        }

        class MalformedBoundObject final : public IFileSystemBoundObject
        {
        public:
            HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** value) noexcept override
            {
                if (value == nullptr)
                    return E_POINTER;
                *value = nullptr;
                if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystemBoundObject))
                {
                    *value = static_cast<IFileSystemBoundObject*>(this);
                    AddRef();
                    return S_OK;
                }
                return E_NOINTERFACE;
            }
            ULONG STDMETHODCALLTYPE AddRef() noexcept override { return _refs.fetch_add(1u, std::memory_order_relaxed) + 1u; }
            ULONG STDMETHODCALLTYPE Release() noexcept override
            {
                const ULONG value = _refs.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
                if (value == 0u)
                    delete this;
                return value;
            }
            HRESULT STDMETHODCALLTYPE GetSnapshot(FileSystemBoundObjectSnapshot* snapshot) noexcept override
            {
                if (snapshot == nullptr)
                    return E_POINTER;
                snapshot->kind               = FILESYSTEM_BOUND_REGULAR_FILE;
                snapshot->objectId           = nullptr;
                snapshot->objectIdBytes      = 16u;
                snapshot->revisionId         = nullptr;
                snapshot->revisionIdBytes    = 0u;
                snapshot->committedSizeBytes = 0u;
                return S_OK;
            }
            HRESULT STDMETHODCALLTYPE IsSameObject(IFileSystemBoundObject*, BOOL*) noexcept override { return E_NOTIMPL; }
            HRESULT STDMETHODCALLTYPE OpenReader(const FileSystemOptions*, IFileReader**) noexcept override { return E_NOTIMPL; }
            HRESULT STDMETHODCALLTYPE GetBasicInformation(FileSystemBasicInformation*) noexcept override { return E_NOTIMPL; }
            HRESULT STDMETHODCALLTYPE SetBasicInformation(const FileSystemBasicInformation*) noexcept override { return E_NOTIMPL; }
            HRESULT STDMETHODCALLTYPE PublishAs(const wchar_t*,
                                                IFileSystemBoundObject*,
                                                FileSystemFlags,
                                                const FileSystemOptions*,
                                                FileSystemConditionalMutationResult*,
                                                IFileSystemBoundObject**) noexcept override
            {
                return E_NOTIMPL;
            }
            HRESULT STDMETHODCALLTYPE RenameIfUnchanged(const wchar_t*,
                                                        IFileSystemBoundObject*,
                                                        FileSystemFlags,
                                                        const FileSystemOptions*,
                                                        FileSystemConditionalMutationResult*,
                                                        IFileSystemBoundObject**) noexcept override
            {
                return E_NOTIMPL;
            }
            HRESULT STDMETHODCALLTYPE DeleteIfUnchanged(FileSystemFlags,
                                                        const FileSystemOptions*,
                                                        FileSystemConditionalMutationResult*) noexcept override
            {
                return E_NOTIMPL;
            }
            HRESULT STDMETHODCALLTYPE AbortOwnedObject(const FileSystemOptions*, FileSystemConditionalMutationResult*) noexcept override
            {
                return E_NOTIMPL;
            }

        private:
            ~MalformedBoundObject() = default;
            std::atomic_ulong _refs{1u};
        };

        wil::com_ptr<IFileSystemBoundObject> malformedBound;
        malformedBound.attach(new (std::nothrow) MalformedBoundObject());
        if (! require(static_cast<bool>(malformedBound), L"Malformed bound-object fixture should allocate."))
        {
            return true;
        }
        const FileOperations::ObjectBindingResult malformedSuccess =
            FileOperations::ValidateSuccessfulBoundObjectForSelfTest(malformedBound.get(), L"local-win32");
        if (! require(malformedSuccess.state == FileOperations::ObjectBindingState::ProviderContractViolation && FAILED(malformedSuccess.status),
                      L"A successful provider snapshot with a null identity payload must be a contract violation."))
        {
            return true;
        }

        class IndeterminateComparisonBoundObject final : public IFileSystemBoundObject
        {
        public:
            HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** value) noexcept override
            {
                if (value == nullptr)
                    return E_POINTER;
                *value = nullptr;
                if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystemBoundObject))
                {
                    *value = static_cast<IFileSystemBoundObject*>(this);
                    AddRef();
                    return S_OK;
                }
                return E_NOINTERFACE;
            }
            ULONG STDMETHODCALLTYPE AddRef() noexcept override { return _refs.fetch_add(1u, std::memory_order_relaxed) + 1u; }
            ULONG STDMETHODCALLTYPE Release() noexcept override
            {
                const ULONG value = _refs.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
                if (value == 0u)
                    delete this;
                return value;
            }
            HRESULT STDMETHODCALLTYPE GetSnapshot(FileSystemBoundObjectSnapshot* snapshot) noexcept override
            {
                if (snapshot == nullptr)
                    return E_POINTER;
                snapshot->kind               = FILESYSTEM_BOUND_REGULAR_FILE;
                snapshot->objectId           = _objectId.data();
                snapshot->objectIdBytes      = static_cast<uint32_t>(_objectId.size());
                snapshot->revisionId         = nullptr;
                snapshot->revisionIdBytes    = 0u;
                snapshot->committedSizeBytes = 0u;
                return S_OK;
            }
            HRESULT STDMETHODCALLTYPE IsSameObject(IFileSystemBoundObject*, BOOL* same) noexcept override
            {
                if (same == nullptr)
                    return E_POINTER;
                *same = FALSE;
                return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
            }
            HRESULT STDMETHODCALLTYPE OpenReader(const FileSystemOptions*, IFileReader**) noexcept override { return E_NOTIMPL; }
            HRESULT STDMETHODCALLTYPE GetBasicInformation(FileSystemBasicInformation*) noexcept override { return E_NOTIMPL; }
            HRESULT STDMETHODCALLTYPE SetBasicInformation(const FileSystemBasicInformation*) noexcept override { return E_NOTIMPL; }
            HRESULT STDMETHODCALLTYPE PublishAs(const wchar_t*,
                                                IFileSystemBoundObject*,
                                                FileSystemFlags,
                                                const FileSystemOptions*,
                                                FileSystemConditionalMutationResult*,
                                                IFileSystemBoundObject**) noexcept override
            {
                return E_NOTIMPL;
            }
            HRESULT STDMETHODCALLTYPE RenameIfUnchanged(const wchar_t*,
                                                        IFileSystemBoundObject*,
                                                        FileSystemFlags,
                                                        const FileSystemOptions*,
                                                        FileSystemConditionalMutationResult*,
                                                        IFileSystemBoundObject**) noexcept override
            {
                return E_NOTIMPL;
            }
            HRESULT STDMETHODCALLTYPE DeleteIfUnchanged(FileSystemFlags,
                                                        const FileSystemOptions*,
                                                        FileSystemConditionalMutationResult*) noexcept override
            {
                return E_NOTIMPL;
            }
            HRESULT STDMETHODCALLTYPE AbortOwnedObject(const FileSystemOptions*, FileSystemConditionalMutationResult*) noexcept override
            {
                return E_NOTIMPL;
            }

        private:
            ~IndeterminateComparisonBoundObject() = default;
            std::atomic_ulong _refs{1u};
            std::array<std::byte, 16u> _objectId{};
        };

        wil::com_ptr<IFileSystemBoundObject> indeterminateBound;
        indeterminateBound.attach(new (std::nothrow) IndeterminateComparisonBoundObject());
        if (! require(static_cast<bool>(indeterminateBound), L"Indeterminate bound-object fixture should allocate."))
        {
            return true;
        }
        FileOperations::ObjectBindingResult indeterminateAuthority =
            FileOperations::ValidateSuccessfulBoundObjectForSelfTest(indeterminateBound.get(), L"local-win32");
        if (! require(indeterminateAuthority.state == FileOperations::ObjectBindingState::Bound,
                      L"Indeterminate comparison fixture must first capture a valid immutable authority snapshot."))
        {
            return true;
        }
        FileOperations::TransferMutationGuard indeterminateGuard{};
        indeterminateGuard.state  = FileOperations::TransferSafetyState::Ready;
        indeterminateGuard.status = S_OK;
        indeterminateGuard.source = std::move(indeterminateAuthority.authority);
        indeterminateGuard.destinationWasMissing = true;
        HRESULT indeterminateStatus = S_OK;
        const FileOperations::TransferSafetyState indeterminateState = FileOperations::RevalidateTransferMutationGuard(state.fsLocal.get(),
                                                                                                                         state.fsLocal.get(),
                                                                                                                         localValidationEndpoint,
                                                                                                                         localValidationEndpoint,
                                                                                                                         FileOperations::TransferIntent::Copy,
                                                                                                                         guardSource.native(),
                                                                                                                         guardMissingDestination.native(),
                                                                                                                         indeterminateGuard,
                                                                                                                         indeterminateStatus);
        if (! require(indeterminateState == FileOperations::TransferSafetyState::Indeterminate && FAILED(indeterminateStatus),
                      L"An indeterminate provider comparison at the final mutation boundary must fail closed before I/O."))
        {
            return true;
        }

        const char* localCapabilitiesJson = nullptr;
        const HRESULT localCapabilitiesHr = state.fsLocal->GetPathCapabilities(L"/", FILESYSTEM_MOVE, &localCapabilitiesJson);
        if (! require(SUCCEEDED(localCapabilitiesHr) && localCapabilitiesJson != nullptr &&
                          AreFileSystemCapabilitiesV2ValidForSelfTest(localCapabilitiesJson),
                      L"Local FileSystem should return a host-parseable capability-v2 document.") ||
            ! require(! AreFileSystemCapabilitiesV2ValidForSelfTest("{}"), L"The host must reject an empty capability-v2 object."))
        {
            return true;
        }
        const std::string localCapabilitiesJsonSnapshot(localCapabilitiesJson);

        constexpr uint64_t kR0ePerfIterations = 4'096u;
        uint64_t r0eCapabilityParseCount      = 0u;
        uint64_t r0eAdmissionDecisionCount    = 0u;
        uint64_t r0eRejectedUncontainedCount  = 0u;
        const auto r0eStarted                 = std::chrono::steady_clock::now();
        for (uint64_t iteration = 0u; iteration < kR0ePerfIterations; ++iteration)
        {
            if (AreFileSystemCapabilitiesV2ValidForSelfTest(localCapabilitiesJsonSnapshot))
            {
                ++r0eCapabilityParseCount;
            }
            const bool admitted = CanSameFileSystemOperation(state.fsLocal, state.tempRoot.native(), FILESYSTEM_COPY, kPluginIdLocal);
            ++r0eAdmissionDecisionCount;
            if (! admitted)
            {
                ++r0eRejectedUncontainedCount;
            }
        }
        std::string missingMoveJson(localCapabilitiesJsonSnapshot);
        constexpr std::string_view kMoveMember = R"json("move":true,)json";
        const size_t moveMemberOffset           = missingMoveJson.find(kMoveMember);
        if (! require(moveMemberOffset != std::string::npos, L"Local capability fixture should contain the required operations.move member."))
        {
            return true;
        }
        missingMoveJson.erase(moveMemberOffset, kMoveMember.size());
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(missingMoveJson),
                      L"The host must reject capability-v2 JSON missing the required operations.move boolean."))
        {
            return true;
        }

        std::string missingAbortJson(localCapabilitiesJsonSnapshot);
        constexpr std::string_view kAbortMember = R"json("abort":false,)json";
        const size_t abortMemberOffset = missingAbortJson.find(kAbortMember);
        if (! require(abortMemberOffset != std::string::npos,
                      L"Local capability fixture should contain the required cancellation.abort member."))
        {
            return true;
        }
        missingAbortJson.erase(abortMemberOffset, kAbortMember.size());
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(missingAbortJson),
                      L"The host must reject capability-v2 JSON missing the required cancellation.abort boolean."))
        {
            return true;
        }

        std::string missingRouteClassJson(localCapabilitiesJsonSnapshot);
        constexpr std::string_view kBoundedRouteMember = R"json("routeClass":"bounded",)json";
        const size_t routeClassMemberOffset = missingRouteClassJson.find(kBoundedRouteMember);
        if (! require(routeClassMemberOffset != std::string::npos,
                      L"Local fixed-volume capability fixture should classify the exact operation route as bounded."))
        {
            return true;
        }
        missingRouteClassJson.erase(routeClassMemberOffset, kBoundedRouteMember.size());
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(missingRouteClassJson),
                      L"The host must reject capability-v2 JSON missing the required cancellation.routeClass."))
        {
            return true;
        }

        std::string invalidRouteClassJson(localCapabilitiesJsonSnapshot);
        const size_t invalidRouteClassOffset = invalidRouteClassJson.find(kBoundedRouteMember);
        invalidRouteClassJson.replace(invalidRouteClassOffset, kBoundedRouteMember.size(), R"json("routeClass":"hostDeadline",)json");
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(invalidRouteClassJson),
                      L"The host must reject a cancellation route class that invents host deadline containment."))
        {
            return true;
        }

        std::string invalidProviderWatchdogJson(localCapabilitiesJsonSnapshot);
        constexpr std::string_view kProviderWatchdogMember = R"json("providerWatchdogTimeoutMs":0)json";
        const size_t providerWatchdogMemberOffset = invalidProviderWatchdogJson.find(kProviderWatchdogMember);
        if (! require(providerWatchdogMemberOffset != std::string::npos,
                      L"A bounded route should carry an explicit zero provider-watchdog timeout."))
        {
            return true;
        }
        invalidProviderWatchdogJson.replace(
            providerWatchdogMemberOffset, kProviderWatchdogMember.size(), R"json("providerWatchdogTimeoutMs":-1)json");
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(invalidProviderWatchdogJson),
                      L"The host must reject a negative provider-owned watchdog timeout."))
        {
            return true;
        }

        std::string missingProviderWatchdogJson(localCapabilitiesJsonSnapshot);
        missingProviderWatchdogJson.erase(providerWatchdogMemberOffset, kProviderWatchdogMember.size());
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(missingProviderWatchdogJson),
                      L"The host must reject capability-v2 JSON missing the required provider-owned watchdog timeout."))
        {
            return true;
        }

        std::string boundedWithWatchdogJson(localCapabilitiesJsonSnapshot);
        boundedWithWatchdogJson.replace(
            providerWatchdogMemberOffset, kProviderWatchdogMember.size(), R"json("providerWatchdogTimeoutMs":30000)json");
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(boundedWithWatchdogJson),
                      L"A bounded route must not claim a provider-owned watchdog timeout."))
        {
            return true;
        }

        std::string watchdogWithoutTimeoutJson(localCapabilitiesJsonSnapshot);
        watchdogWithoutTimeoutJson.replace(
            routeClassMemberOffset, kBoundedRouteMember.size(), R"json("routeClass":"providerWatchdog",)json");
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(watchdogWithoutTimeoutJson),
                      L"A provider-watchdog route must carry a nonzero provider-owned timeout."))
        {
            return true;
        }

        std::string validProviderWatchdogJson(std::move(watchdogWithoutTimeoutJson));
        const size_t watchdogTimeoutOffset = validProviderWatchdogJson.find(kProviderWatchdogMember);
        if (! require(watchdogTimeoutOffset != std::string::npos,
                      L"Provider-watchdog fixture should retain the required timeout member."))
        {
            return true;
        }
        validProviderWatchdogJson.replace(
            watchdogTimeoutOffset, kProviderWatchdogMember.size(), R"json("providerWatchdogTimeoutMs":30000)json");
        if (! require(AreFileSystemCapabilitiesV2ValidForSelfTest(validProviderWatchdogJson),
                      L"A provider-watchdog route with a nonzero provider-owned timeout must remain valid."))
        {
            return true;
        }
        if (! require(std::string_view(localCapabilitiesJsonSnapshot).find("quietPointTimeoutMs") == std::string_view::npos,
                      L"Capability v2 must not advertise the retired unenforced host quiet-point timeout."))
        {
            return true;
        }

        std::string uncontainedCapabilitiesJson(localCapabilitiesJsonSnapshot);
        constexpr std::string_view kUncontainedRouteMember = R"json("routeClass":"uncontained",)json";
        uncontainedCapabilitiesJson.replace(
            uncontainedCapabilitiesJson.find(kBoundedRouteMember), kBoundedRouteMember.size(), kUncontainedRouteMember);

        auto typedFalseOwner = std::make_unique<UncontainedNeverReturningFileSystem>(
            uncontainedCapabilitiesJson, true, false);
        const FileSystemRouteContract::QueryResult typedFalse = FileSystemRouteContract::Query(
            static_cast<IFileSystem*>(typedFalseOwner.get()),
            L"/",
            FILESYSTEM_COPY,
            L"selftest/uncontained-never-return");
        auto malformedJsonOwner = std::make_unique<UncontainedNeverReturningFileSystem>("not-json", true, true);
        const FileSystemRouteContract::QueryResult typedTrueMalformedJson = FileSystemRouteContract::Query(
            static_cast<IFileSystem*>(malformedJsonOwner.get()),
            L"/",
            FILESYSTEM_COPY,
            L"selftest/uncontained-never-return");
        auto missingTypedOwner = std::make_unique<UncontainedNeverReturningFileSystem>(
            uncontainedCapabilitiesJson, false, true);
        const FileSystemRouteContract::QueryResult missingTyped = FileSystemRouteContract::Query(
            static_cast<IFileSystem*>(missingTypedOwner.get()),
            L"/",
            FILESYSTEM_COPY,
            L"selftest/uncontained-never-return");
        if (! require(typedFalse.state == FileSystemRouteContract::QueryState::Available &&
                          ! typedFalse.snapshot.copyOperation &&
                          typedTrueMalformedJson.state == FileSystemRouteContract::QueryState::Available &&
                          typedTrueMalformedJson.snapshot.copyOperation &&
                          missingTyped.state == FileSystemRouteContract::QueryState::ContractViolation &&
                          missingTyped.status == E_NOINTERFACE,
                      L"Typed route facts must override positive JSON, ignore malformed diagnostic JSON, and fail closed when the typed IID is absent."))
        {
            return true;
        }

        auto neverReturningOwner = std::make_unique<UncontainedNeverReturningFileSystem>(std::move(uncontainedCapabilitiesJson));
        UncontainedNeverReturningFileSystem* const neverReturningObserver = neverReturningOwner.get();
        wil::com_ptr<IFileSystem> neverReturningFileSystem;
        neverReturningFileSystem.attach(neverReturningOwner.release());
        std::vector<FolderWindow::FileOperationState::Task*> tasksBeforeUncontainedAdmission;
        std::vector<FolderWindow::FileOperationState::Task*> tasksAfterUncontainedAdmission;
        state.fileOps->CollectTasks(tasksBeforeUncontainedAdmission);
        uint64_t rejectedTaskId = 0u;
        const HRESULT uncontainedAdmissionHr = state.fileOps->AdmitOperation(
            FILESYSTEM_COPY,
            FolderWindow::Pane::Left,
            FolderWindow::Pane::Right,
            neverReturningFileSystem,
            {std::filesystem::path(L"/never-return/source.bin")},
            std::filesystem::path(L"/never-return/destination"),
            FILESYSTEM_FLAG_NONE,
            false,
            0u,
            FolderWindow::FileOperationState::ExecutionMode::PerItem,
            false,
            nullptr,
            &rejectedTaskId,
            {},
            {},
            L"selftest/uncontained-never-return",
            L"uncontained-never-return");
        state.fileOps->CollectTasks(tasksAfterUncontainedAdmission);
        const uint64_t providerOperationCallCount = neverReturningObserver->ProviderOperationCallCount();
        if (! require(uncontainedAdmissionHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) && rejectedTaskId == 0u &&
                          tasksAfterUncontainedAdmission.size() == tasksBeforeUncontainedAdmission.size() &&
                          providerOperationCallCount == 0u,
                      L"An uncontained never-returning route must be rejected before task publication, worker creation, or provider operation calls."))
        {
            return true;
        }
        ++r0eRejectedUncontainedCount;

        const uint64_t r0eElapsedUs = Debug::Perf::ElapsedUs(r0eStarted);
        Debug::Perf::Emit(L"FileOps.SelfTest.R0e.RouteContainmentUs",
                          L"4096-capability-admission-decisions",
                          r0eElapsedUs,
                          r0eCapabilityParseCount,
                          r0eAdmissionDecisionCount,
                          S_OK);
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0e.CapabilityParseCount", r0eCapabilityParseCount, S_OK);
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0e.AdmissionDecisionCount", r0eAdmissionDecisionCount, S_OK);
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0e.RejectedUncontainedCount", r0eRejectedUncontainedCount, S_OK);
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0e.ProviderCallCount", providerOperationCallCount, S_OK);
        // Release evidence retains the two-second acceptance gate. Debug carries
        // iterator/runtime checks and uses a wider diagnostic ceiling so focused
        // correctness runs do not masquerade as production perf qualification.
#if defined(_DEBUG)
        constexpr uint64_t kR0eSelfTestCeilingUs = 5'000'000u;
#else
        constexpr uint64_t kR0eSelfTestCeilingUs = 2'000'000u;
#endif
        if (! require(r0eCapabilityParseCount == kR0ePerfIterations &&
                          r0eAdmissionDecisionCount == kR0ePerfIterations &&
                          r0eRejectedUncontainedCount == 1u && providerOperationCallCount == 0u &&
                          r0eElapsedUs < kR0eSelfTestCeilingUs,
                      L"R0e capability/admission containment must remain complete, provider-call-free, and within the build-specific ceiling."))
        {
            return true;
        }

        std::string missingCreateDirectoryJson(localCapabilitiesJsonSnapshot);
        constexpr std::string_view kCreateDirectoryMember = R"json("createDirectory":true,)json";
        const size_t createDirectoryMemberOffset = missingCreateDirectoryJson.find(kCreateDirectoryMember);
        if (! require(createDirectoryMemberOffset != std::string::npos,
                      L"Local capability fixture should contain the required operations.createDirectory member."))
        {
            return true;
        }
        missingCreateDirectoryJson.erase(createDirectoryMemberOffset, kCreateDirectoryMember.size());
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(missingCreateDirectoryJson),
                      L"The host must reject capability-v2 JSON missing the required operations.createDirectory boolean."))
        {
            return true;
        }

        std::string emptyRootJson(localCapabilitiesJsonSnapshot);
        constexpr std::string_view kRootPrefix = R"json("rootId": ")json";
        const size_t rootValueBegin = emptyRootJson.find(kRootPrefix);
        const size_t rootValueEnd = rootValueBegin == std::string::npos
            ? std::string::npos
            : emptyRootJson.find('"', rootValueBegin + kRootPrefix.size());
        if (! require(rootValueBegin != std::string::npos && rootValueEnd != std::string::npos,
                      L"Local capability fixture should contain a replaceable rootId string."))
        {
            return true;
        }
        emptyRootJson.erase(rootValueBegin + kRootPrefix.size(), rootValueEnd - (rootValueBegin + kRootPrefix.size()));
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(emptyRootJson),
                      L"The host must reject capability-v2 JSON with an empty path-scoped rootId."))
        {
            return true;
        }

        std::string missingVerificationJson(localCapabilitiesJsonSnapshot);
        const size_t verificationMemberOffset = missingVerificationJson.find(R"json("verification")json");
        if (! require(verificationMemberOffset != std::string::npos,
                      L"Local capability fixture should contain the required verification object."))
        {
            return true;
        }
        const size_t verificationLineStart = missingVerificationJson.rfind('\n', verificationMemberOffset);
        const size_t verificationLineEnd = missingVerificationJson.find('\n', verificationMemberOffset);
        if (! require(verificationLineStart != std::string::npos && verificationLineEnd != std::string::npos,
                      L"Local capability verification fixture should occupy one removable JSON line."))
        {
            return true;
        }
        missingVerificationJson.erase(verificationLineStart + 1u, verificationLineEnd - verificationLineStart);
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(missingVerificationJson),
                      L"The host must reject capability-v2 JSON missing the required verification object."))
        {
            return true;
        }

        std::string unknownProofJson(localCapabilitiesJsonSnapshot);
        constexpr std::string_view kKnownProviderProof = "blake3-bound-object";
        const size_t providerProofOffset = unknownProofJson.find(kKnownProviderProof);
        if (! require(providerProofOffset != std::string::npos,
                      L"Local capability fixture should contain its exact BLAKE3 provider-proof claim."))
        {
            return true;
        }
        unknownProofJson.replace(providerProofOffset, kKnownProviderProof.size(), "unknown-proof");
        if (! require(! AreFileSystemCapabilitiesV2ValidForSelfTest(unknownProofJson),
                      L"The host must reject an unknown verification.providerProof algorithm."))
        {
            return true;
        }

        constexpr FileSystemFlags kExpectedAtomicFlags = static_cast<FileSystemFlags>(
            static_cast<uint32_t>(FILESYSTEM_FLAG_RECURSIVE) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE) |
            static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY));
        FlagSensitiveAtomicWriter flagSensitiveAtomicWriter(kExpectedAtomicFlags);
        FileSystemFlags effectiveFinalFlags   = FILESYSTEM_FLAG_NONE;
        FileSystemFlags fallbackSiblingFlags = FILESYSTEM_FLAG_NONE;
        const bool useAtomicWriter = ResolveFileOpsAtomicWriterRouteForSelfTest(&flagSensitiveAtomicWriter,
                                                                                L"/atomic-flags.bin",
                                                                                FILESYSTEM_FLAG_RECURSIVE,
                                                                                true,
                                                                                true,
                                                                                effectiveFinalFlags,
                                                                                fallbackSiblingFlags);
        if (! require(useAtomicWriter && flagSensitiveAtomicWriter.ProbeCount() == 1u &&
                          static_cast<uint32_t>(flagSensitiveAtomicWriter.ObservedFlags()) == static_cast<uint32_t>(kExpectedAtomicFlags) &&
                          static_cast<uint32_t>(effectiveFinalFlags) == static_cast<uint32_t>(kExpectedAtomicFlags) &&
                          fallbackSiblingFlags == FILESYSTEM_FLAG_RECURSIVE,
                      L"Atomic-writer probing and final-path creation should consume the same effective per-item overwrite/read-only grants."))
        {
            return true;
        }

        FlagSensitiveAtomicWriter decliningAtomicWriter(FILESYSTEM_FLAG_NONE);
        effectiveFinalFlags   = FILESYSTEM_FLAG_NONE;
        fallbackSiblingFlags = FILESYSTEM_FLAG_NONE;
        if (! require(! ResolveFileOpsAtomicWriterRouteForSelfTest(&decliningAtomicWriter,
                                                                    L"/atomic-decline.bin",
                                                                    FILESYSTEM_FLAG_RECURSIVE,
                                                                    true,
                                                                    true,
                                                                    effectiveFinalFlags,
                                                                    fallbackSiblingFlags) &&
                          decliningAtomicWriter.ProbeCount() == 1u &&
                          static_cast<uint32_t>(decliningAtomicWriter.ObservedFlags()) == static_cast<uint32_t>(kExpectedAtomicFlags) &&
                          fallbackSiblingFlags == FILESYSTEM_FLAG_RECURSIVE,
                      L"Only a declined atomic route should strip overwrite/read-only grants for the random sibling writer."))
        {
            return true;
        }

        const auto requirePathIdentity = [&](const ProviderCapabilitySnapshot& caps,
                                             std::wstring_view providerName,
                                             bool expectedStable,
                                             std::wstring_view expectedComparison,
                                             wchar_t expectedPreferredSeparator,
                                             std::wstring_view expectedAcceptedSeparators,
                                             std::wstring_view expectedCaseOnlyRename) noexcept -> bool
        {
            if (! caps.pathIdentityPresent)
            {
                Fail(std::format(L"{} should advertise pathIdentity.", providerName));
                return false;
            }
            if (caps.pathTextStableIdentity != expectedStable)
            {
                Fail(std::format(L"{} pathTextStableIdentity mismatch.", providerName));
                return false;
            }
            if (caps.componentComparison != expectedComparison)
            {
                Fail(std::format(L"{} componentComparison mismatch: expected {} actual {}.", providerName, expectedComparison, caps.componentComparison));
                return false;
            }
            if (caps.preferredSeparator != expectedPreferredSeparator)
            {
                Fail(std::format(L"{} preferredSeparator mismatch.", providerName));
                return false;
            }
            if (caps.acceptedSeparators != expectedAcceptedSeparators)
            {
                Fail(std::format(L"{} acceptedSeparators mismatch: expected {} actual {}.", providerName, expectedAcceptedSeparators, caps.acceptedSeparators));
                return false;
            }
            if (caps.caseOnlyRename != expectedCaseOnlyRename)
            {
                Fail(std::format(L"{} caseOnlyRename mismatch: expected {} actual {}.", providerName, expectedCaseOnlyRename, caps.caseOnlyRename));
                return false;
            }
            return true;
        };

        if (! require(localCaps.copyOperation && localCaps.moveOperation && localCaps.nativeMoveOperation && localCaps.renameOperation &&
                          localCaps.deleteOperation && localCaps.read && localCaps.write,
                      L"Local FileSystem should advertise its native-only same-volume Move route.") ||
            ! require(CanSameFileSystemOperation(state.fsLocal, state.tempRoot.native(), FILESYSTEM_MOVE, kPluginIdLocal),
                      L"Host admission should accept the local qualified native Move route.") ||
            ! require(localCaps.copyMoveMax >= 1u && localCaps.deleteMax >= 1u && localCaps.deleteRecycleMax >= 1u,
                      L"Local FileSystem should advertise positive concurrency limits.") ||
            ! require(localCaps.exportCopyWildcard && localCaps.exportMoveWildcard && localCaps.importCopyWildcard && localCaps.importMoveWildcard,
                      L"Local FileSystem should advertise wildcard cross-filesystem copy/move import/export.") ||
            ! requirePathIdentity(localCaps, L"Local FileSystem", true, L"ordinalIgnoreCase", L'\\', L"\\/", L"supported"))
        {
            return true;
        }


        if (! state.fileOps)
        {
            Fail(L"Provider capability matrix test lost FileOperationState.");
            return true;
        }

        const std::filesystem::path copyOnlySource = state.tempRoot / L"copy-only-move-source.bin";
        const std::filesystem::path copyOnlyDestinationRoot = state.tempRoot / L"copy-only-move-destination";
        const std::filesystem::path copyOnlyDestination = copyOnlyDestinationRoot / copyOnlySource.filename();
        std::error_code copyOnlyEc;
        std::filesystem::create_directories(copyOnlyDestinationRoot, copyOnlyEc);
        if (! require(! copyOnlyEc && SelfTest::WriteTextFile(copyOnlySource, "copy-only"),
                      L"Copy-only Move fixture should create its source and destination folder."))
        {
            return true;
        }

        FileOperations::TransferPlan copyOnlyPlan{};
        copyOnlyPlan.intent              = FileOperations::TransferIntent::Move;
        copyOnlyPlan.strategy            = FileOperations::OperationStrategy::CopyOnly;
        copyOnlyPlan.sourceEndpoint      = localValidationEndpoint;
        copyOnlyPlan.destinationEndpoint = localValidationEndpoint;
        copyOnlyPlan.selectedItems.push_back({.providerPath = copyOnlySource.native()});
        copyOnlyPlan.destination.providerFolderPath = copyOnlyDestinationRoot.native();

        FolderWindow::FileOperationState::Task copyOnlyTask(*state.fileOps);
        copyOnlyTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(copyOnlyPlan)}}));
        copyOnlyTask._operation       = FILESYSTEM_MOVE;
        copyOnlyTask._executionMode   = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        copyOnlyTask._fileSystem      = state.fsLocal;
        copyOnlyTask._sourcePaths     = {copyOnlySource};
        copyOnlyTask._destinationFolder = copyOnlyDestinationRoot;
        copyOnlyTask._sourcePathAttributesHint = {GetFileAttributesW(copyOnlySource.c_str())};
        AppendLog(L"Provider matrix: preparing Copy-only interlock roles.");
        const HRESULT copyOnlyInterlockHr = copyOnlyTask.PrepareMutationInterlockScopes();
        AppendLog(std::format(L"Provider matrix: Copy-only interlock preparation returned 0x{:08X}.",
                              static_cast<unsigned long>(copyOnlyInterlockHr)));
        const bool hasCopyOnlyReadScope = std::ranges::any_of(
            copyOnlyTask._mutationInterlockScopes,
            [&](const FileOperations::MutationInterlockScope& scope) noexcept
            {
                return scope.access == FileOperations::MutationInterlockAccess::ReadSource &&
                       EquivalentPath(localValidationEndpoint.pathIdentity.value(), scope.providerPath, copyOnlySource.native());
            });
        const auto copyOnlyPublishScopeIt = std::ranges::find_if(
            copyOnlyTask._mutationInterlockScopes,
            [&](const FileOperations::MutationInterlockScope& scope) noexcept
            {
                return scope.access == FileOperations::MutationInterlockAccess::PublishDestination &&
                       EquivalentPath(localValidationEndpoint.pathIdentity.value(), scope.providerPath, copyOnlyDestination.native()) &&
                       ! scope.anchors.empty();
            });
        const bool hasCopyOnlyPublishScope = copyOnlyPublishScopeIt != copyOnlyTask._mutationInterlockScopes.end();
        if (! require(SUCCEEDED(copyOnlyInterlockHr) && hasCopyOnlyReadScope && hasCopyOnlyPublishScope,
                      L"Copy-only Move must register its source as ReadSource and its final destination leaf as anchored PublishDestination."))
        {
            return true;
        }

        {
            const FileOperations::MutationInterlockScope& publishScope = *copyOnlyPublishScopeIt;
            FileOperations::MutationInterlockScope sameLeafScope       = publishScope;
            FileOperations::MutationInterlockScope distinctSiblingScope = publishScope;
            distinctSiblingScope.providerPath = (copyOnlyDestinationRoot / L"different-sibling.bin").native();
            bool rebuiltDistinctAnchors = true;
            for (FileOperations::MutationInterlockAnchor& anchor : distinctSiblingScope.anchors)
            {
                rebuiltDistinctAnchors = rebuiltDistinctAnchors && anchor.authority &&
                    TryGetFileSystemRelativePath(localValidationEndpoint.pathIdentity.value(),
                                                 anchor.authority->retained.providerPath,
                                                 distinctSiblingScope.providerPath,
                                                 anchor.relativePath);
                anchor.relativePathKey = TryMakePathKey(localValidationEndpoint.pathIdentity.value(), anchor.relativePath);
                rebuiltDistinctAnchors = rebuiltDistinctAnchors && anchor.relativePathKey.has_value();
            }

            FileOperations::MutationInterlockScope ancestorScope = publishScope;
            ancestorScope.providerPath = copyOnlyDestinationRoot.native();
            ancestorScope.anchors.resize(1u);
            ancestorScope.anchors.front().relativePath.clear();
            ancestorScope.anchors.front().relativePathKey =
                TryMakePathKey(localValidationEndpoint.pathIdentity.value(), L"");

            FileOperations::MutationInterlockScope conservativeScope = publishScope;
            conservativeScope.root.reset();
            conservativeScope.anchors.clear();
            conservativeScope.conservativeIdentityDomain = true;

            FileOperations::MutationInterlockScope sharedReadScope = publishScope;
            sharedReadScope.access = FileOperations::MutationInterlockAccess::ReadSource;
            sameLeafScope.access   = FileOperations::MutationInterlockAccess::ReadSource;

            if (! require(rebuiltDistinctAnchors &&
                              FileOperations::DebugMutationScopesOverlapForSelfTest(publishScope, publishScope) &&
                              ! FileOperations::DebugMutationScopesOverlapForSelfTest(publishScope, distinctSiblingScope) &&
                              FileOperations::DebugMutationScopesOverlapForSelfTest(publishScope, ancestorScope) &&
                              FileOperations::DebugMutationScopesOverlapForSelfTest(publishScope, conservativeScope) &&
                              ! FileOperations::DebugMutationScopesOverlapForSelfTest(sameLeafScope, sharedReadScope),
                          L"Anchored interlocks must serialize the same or ancestor target, admit distinct siblings, fail unsupported identity domains conservatively, and share reads."))
            {
                return true;
            }
        }

        {
            constexpr size_t kLargeSelectionCount = 256u;
            constexpr size_t kSharedAncestorAllowance = 32u;
            const std::filesystem::path retentionSourceRoot = state.tempRoot / L"interlock-large-selection-source";
            const std::filesystem::path retentionDestinationRoot = state.tempRoot / L"interlock-large-selection-destination";
            if (! require(RecreateEmptyDirectory(retentionSourceRoot) && RecreateEmptyDirectory(retentionDestinationRoot),
                          L"Large-selection interlock fixture should create its source and destination roots."))
            {
                return true;
            }

            FileOperations::TransferPlan retentionPlan{};
            retentionPlan.intent              = FileOperations::TransferIntent::Copy;
            retentionPlan.strategy            = FileOperations::OperationStrategy::Copy;
            retentionPlan.sourceEndpoint      = localValidationEndpoint;
            retentionPlan.destinationEndpoint = localValidationEndpoint;
            retentionPlan.destination.providerFolderPath = retentionDestinationRoot.native();
            std::vector<std::filesystem::path> retentionSources;
            retentionPlan.selectedItems.reserve(kLargeSelectionCount);
            retentionSources.reserve(kLargeSelectionCount);
            for (size_t index = 0u; index < kLargeSelectionCount; ++index)
            {
                const std::filesystem::path source = retentionSourceRoot / std::format(L"item-{:03}.bin", index);
                if (! SelfTest::WriteTextFile(source, "x"))
                {
                    Fail(L"Large-selection interlock fixture could not create every selected file.");
                    return true;
                }
                retentionPlan.selectedItems.emplace_back(FileOperations::QualifiedSourceItem{.providerPath = source.native()});
                retentionSources.emplace_back(source);
            }

            FolderWindow::FileOperationState::Task retentionTask(*state.fileOps);
            retentionTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
                FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(retentionPlan)}}));
            retentionTask._operation         = FILESYSTEM_COPY;
            retentionTask._executionMode     = FolderWindow::FileOperationState::ExecutionMode::PerItem;
            retentionTask._fileSystem        = state.fsLocal;
            retentionTask._sourcePaths       = std::move(retentionSources);
            retentionTask._destinationFolder = retentionDestinationRoot;

            const auto retentionStarted = std::chrono::steady_clock::now();
            const HRESULT retentionHr = retentionTask.PrepareMutationInterlockScopes();
            std::unordered_set<const FileOperations::MutationInterlockAuthorityNode*> uniqueAuthorityNodes;
            size_t maxAuthorityDepth = 0u;
            size_t readSourceScopes = 0u;
            size_t publishDestinationScopes = 0u;
            for (const FileOperations::MutationInterlockScope& scope : retentionTask._mutationInterlockScopes)
            {
                readSourceScopes += scope.access == FileOperations::MutationInterlockAccess::ReadSource ? 1u : 0u;
                publishDestinationScopes += scope.access == FileOperations::MutationInterlockAccess::PublishDestination ? 1u : 0u;
                size_t authorityDepth = 0u;
                for (std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode> node = scope.root;
                     node;
                     node = node->parent)
                {
                    uniqueAuthorityNodes.emplace(node.get());
                    ++authorityDepth;
                    if (authorityDepth > 1024u)
                    {
                        break;
                    }
                }
                maxAuthorityDepth = (std::max)(maxAuthorityDepth, authorityDepth);
            }
            const uint64_t retentionPrepareUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(
                                                                           std::chrono::steady_clock::now() - retentionStarted)
                                                                           .count());
            if (! require(SUCCEEDED(retentionHr) &&
                              retentionTask._mutationInterlockScopes.size() == kLargeSelectionCount * 2u &&
                              readSourceScopes == kLargeSelectionCount && publishDestinationScopes == kLargeSelectionCount &&
                              uniqueAuthorityNodes.size() <= kLargeSelectionCount + kSharedAncestorAllowance &&
                              maxAuthorityDepth <= 1024u,
                          L"Large-selection interlock retention must keep one source and final-leaf scope per item while interning shared ancestors."))
            {
                return true;
            }

            std::vector<FileOperations::MutationInterlockScope> retainedPublishScopes;
            std::vector<FileOperations::MutationInterlockScope> disjointPublishScopes;
            retainedPublishScopes.reserve(kLargeSelectionCount);
            disjointPublishScopes.reserve(kLargeSelectionCount);
            for (const FileOperations::MutationInterlockScope& scope : retentionTask._mutationInterlockScopes)
            {
                if (scope.access != FileOperations::MutationInterlockAccess::PublishDestination)
                {
                    continue;
                }

                retainedPublishScopes.emplace_back(scope);
                FileOperations::MutationInterlockScope siblingScope = scope;
                siblingScope.providerPath =
                    (retentionDestinationRoot / std::format(L"disjoint-{:03}.bin", disjointPublishScopes.size())).native();
                bool rebuiltAnchors = true;
                for (FileOperations::MutationInterlockAnchor& anchor : siblingScope.anchors)
                {
                    rebuiltAnchors = rebuiltAnchors && anchor.authority &&
                        TryGetFileSystemRelativePath(localValidationEndpoint.pathIdentity.value(),
                                                     anchor.authority->retained.providerPath,
                                                     siblingScope.providerPath,
                                                     anchor.relativePath);
                    anchor.relativePathKey = TryMakePathKey(localValidationEndpoint.pathIdentity.value(), anchor.relativePath);
                    rebuiltAnchors = rebuiltAnchors && anchor.relativePathKey.has_value();
                }
                if (! require(rebuiltAnchors,
                              L"Large-selection disjoint admission fixture must rebuild every retained anchor suffix."))
                {
                    return true;
                }
                disjointPublishScopes.emplace_back(std::move(siblingScope));
            }

            const auto comparisonStarted = std::chrono::steady_clock::now();
            uint64_t scopeComparisons = 0u;
            const bool disjointSetsOverlap = FileOperations::DebugMutationScopeSetsOverlapForSelfTest(
                retainedPublishScopes,
                disjointPublishScopes,
                scopeComparisons);
            const uint64_t comparisonUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(
                                                                     std::chrono::steady_clock::now() - comparisonStarted)
                                                                     .count());
            constexpr uint64_t kCartesianScopeComparisons = kLargeSelectionCount * kLargeSelectionCount;
            if (! require(! disjointSetsOverlap && retainedPublishScopes.size() == kLargeSelectionCount &&
                              disjointPublishScopes.size() == kLargeSelectionCount &&
                              scopeComparisons > 0u && scopeComparisons < kCartesianScopeComparisons,
                          L"Large-selection admission must prove 256x256 disjoint publication leaves without a false overlap or Cartesian scan."))
            {
                return true;
            }

            Debug::Perf::Emit(L"FileOps.Interlock.CheckUs",
                              L"selftest-local-256x256-disjoint-publish-leaves",
                              comparisonUs,
                              scopeComparisons,
                              1u,
                              S_OK);
            Debug::Perf::Emit(L"FileOps.Interlock.ScopeComparisons",
                              L"selftest-local-256x256-disjoint-publish-leaves",
                              0u,
                              scopeComparisons,
                              1u,
                              S_OK);
            Debug::Perf::Emit(L"FileOps.Interlock.ActiveCandidates",
                              L"selftest-local-256x256-disjoint-publish-leaves",
                              0u,
                              1u,
                              scopeComparisons,
                              S_OK);
            Debug::Perf::Emit(L"FileOps.SelfTest.InterlockLargeSelectionRetainedAuthorityNodes",
                              L"local-256-top-level-items",
                              retentionPrepareUs,
                              uniqueAuthorityNodes.size(),
                              kLargeSelectionCount,
                              S_OK);
            Debug::Perf::Emit(L"FileOps.SelfTest.InterlockLargeSelectionMaxAuthorityDepth",
                              L"local-256-top-level-items",
                              0u,
                              maxAuthorityDepth,
                              1024u,
                              S_OK);
        }

        {
            FolderWindow::FileOperationState::Task task(*state.fileOps);
            task._operation   = FILESYSTEM_COPY;
            task._sourcePaths = {state.tempRoot / L"typed-stage-cleanup-source.bin"};
            task.InitializeSourceItemResultBuilders();
            task.MarkSourceItemsMutationPossible();
            task._sourceItemResultBuilders.front().status = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
            task._sourceItemResultBuilders.front().mutation = FileSystemItemMutationResult{
                .sizeBytes             = sizeof(FileSystemItemMutationResult),
                .outcomeKnown          = TRUE,
                .mutationCommitted     = FALSE,
                .originalStillPresent  = TRUE,
                .ownedStageDisposition = FileSystemOwnedStageDisposition::Unknown,
            };

            const HRESULT typedHr = task.FinalizeTypedItemResults(HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE));
            const FileOperations::FileOperationItemResult* result = task._sourceItemResultBuilders.size() == 1u &&
                    task._sourceItemResultBuilders.front().terminal.has_value()
                ? std::addressof(task._sourceItemResultBuilders.front().terminal.value())
                : nullptr;
            const bool axesMatch = result != nullptr &&
                result->publication == FileOperations::PublicationState::NotPublished &&
                result->sourceDisposition == FileOperations::SourceDisposition::Retained &&
                result->completion == FileOperations::ItemCompletion::Indeterminate &&
                result->ownedStageDisposition == FileOperations::OwnedStageDisposition::Unknown;
            if (! require(typedHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) && axesMatch,
                          L"An unknown owned-stage cleanup must remain NotPublished + source Retained while the independent stage axis is Unknown."))
            {
                return true;
            }
        }

        {
            constexpr size_t kCleanupDebtResultCount = 4'096u;
            FolderWindow::FileOperationState::Task task(*state.fileOps);
            task._operation = FILESYSTEM_COPY;
            task._sourcePaths.resize(kCleanupDebtResultCount, state.tempRoot / L"cleanup-debt-source.bin");
            task.InitializeSourceItemResultBuilders();
            task.MarkSourceItemsMutationPossible();
            for (auto& builder : task._sourceItemResultBuilders)
            {
                builder.status = S_OK;
                builder.mutation = FileSystemItemMutationResult{
                    .sizeBytes             = sizeof(FileSystemItemMutationResult),
                    .outcomeKnown          = TRUE,
                    .mutationCommitted     = TRUE,
                    .originalStillPresent  = TRUE,
                    .ownedStageDisposition = FileSystemOwnedStageDisposition::Retained,
                };
            }

            const auto reductionStartedAt = std::chrono::steady_clock::now();
            const HRESULT reductionHr = task.FinalizeTypedItemResults(S_OK);
            const uint64_t reductionUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(
                                                                    std::chrono::steady_clock::now() - reductionStartedAt)
                                                                    .count());
            size_t completedCount = 0u;
            size_t retainedDebtCount = 0u;
            size_t indeterminateCount = 0u;
            for (const auto& builder : task._sourceItemResultBuilders)
            {
                if (! builder.terminal.has_value())
                {
                    continue;
                }
                const FileOperations::FileOperationItemResult& result = builder.terminal.value();
                completedCount += result.completion == FileOperations::ItemCompletion::Completed ? 1u : 0u;
                retainedDebtCount += result.ownedStageDisposition == FileOperations::OwnedStageDisposition::Retained ? 1u : 0u;
                indeterminateCount += result.completion == FileOperations::ItemCompletion::Indeterminate ? 1u : 0u;
            }
            const uint64_t retainedBytesPerResult = sizeof(task._sourceItemResultBuilders.front()) + sizeof(task._sourcePaths.front());
            Debug::Perf::Emit(L"FileOps.CleanupDebt.ReduceUs",
                              L"committed-retained-4096",
                              reductionUs,
                              kCleanupDebtResultCount,
                              completedCount,
                              reductionHr);
            Debug::Perf::EmitValue(L"FileOps.CleanupDebt.ResultCount", kCleanupDebtResultCount, reductionHr);
            Debug::Perf::EmitValue(L"FileOps.CleanupDebt.CompletedCount", completedCount, reductionHr);
            Debug::Perf::EmitValue(L"FileOps.CleanupDebt.RetainedDebtCount", retainedDebtCount, reductionHr);
            Debug::Perf::EmitValue(L"FileOps.CleanupDebt.IndeterminateCount", indeterminateCount, reductionHr);
            Debug::Perf::EmitValue(L"FileOps.CleanupDebt.RetainedBytesPerResult", retainedBytesPerResult, reductionHr);

            if (! require(reductionHr == S_OK && completedCount == kCleanupDebtResultCount &&
                              retainedDebtCount == kCleanupDebtResultCount && indeterminateCount == 0u,
                          L"Committed primary results with retained cleanup debt must remain Completed while preserving the independent Retained debt axis."))
            {
                return true;
            }
        }

        AppendLog(L"Provider matrix: executing Copy-only fixture.");
        const HRESULT copyOnlyHr = copyOnlyTask.ExecuteOperation();
        AppendLog(std::format(L"Provider matrix: Copy-only fixture returned 0x{:08X}.", static_cast<unsigned long>(copyOnlyHr)));
        if (! require(copyOnlyHr == S_FALSE && std::filesystem::exists(copyOnlySource, copyOnlyEc) &&
                          std::filesystem::exists(copyOnlyDestination, copyOnlyEc) &&
                          std::filesystem::file_size(copyOnlyDestination, copyOnlyEc) == 9u && ! copyOnlyEc,
                      L"Copy-only Move must execute Copy, retain the source, and report a non-clean source-kept result."))
        {
            return true;
        }

        constexpr std::wstring_view entropyDummyRoot = L"/stage-entropy-source";
        constexpr std::wstring_view entropyDummyFile = L"/stage-entropy-source/payload.bin";
        const std::filesystem::path entropyDestinationRoot = state.tempRoot / L"stage-entropy-destination";
        const std::filesystem::path entropyDestination = entropyDestinationRoot / L"payload.bin";
        static_cast<void>(state.fsDummy->DeleteItem(std::wstring(entropyDummyRoot).c_str(),
                                                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY),
                                                    nullptr,
                                                    nullptr,
                                                    nullptr));
        if (! require(EnsureDummyFolderExists(state.fsDummy.get(), entropyDummyRoot) &&
                          DummyWriteTextFile(state.fsDummy.get(), entropyDummyFile, "entropy-source") &&
                          RecreateEmptyDirectory(entropyDestinationRoot),
                      L"Stage entropy failure fixture should create its source and destination root."))
        {
            return true;
        }

        FileOperations::TransferPlan entropyPlan{};
        entropyPlan.intent              = FileOperations::TransferIntent::Copy;
        entropyPlan.strategy            = FileOperations::OperationStrategy::Copy;
        entropyPlan.sourceEndpoint      = dummyValidationEndpoint;
        entropyPlan.destinationEndpoint = localValidationEndpoint;
        entropyPlan.selectedItems.push_back({.providerPath = std::wstring(entropyDummyFile)});
        entropyPlan.destination.providerFolderPath = entropyDestinationRoot.native();

        FolderWindow::FileOperationState::Task entropyTask(*state.fileOps);
        entropyTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(entropyPlan)}}));
        entropyTask._operation                 = FILESYSTEM_COPY;
        entropyTask._executionMode             = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        entropyTask._fileSystem                = state.fsDummy;
        entropyTask._destinationFileSystem     = state.fsLocal;
        entropyTask._sourcePaths               = {std::filesystem::path(entropyDummyFile)};
        entropyTask._destinationFolder         = entropyDestinationRoot;
        entropyTask._sourcePathAttributesHint  = {FILE_ATTRIBUTE_NORMAL};
        entropyTask._flags = FILESYSTEM_FLAG_CONTINUE_ON_ERROR;

        SetFileOpsBridgeFailNextStageEntropyForSelfTest(1u);
        const auto resetStageEntropyHook = wil::scope_exit([]() noexcept
        {
            SetFileOpsBridgeFailNextStageEntropyForSelfTest(0u);
            static_cast<void>(TakeFileOpsBridgeFailNextStageEntropyAttemptsForSelfTest());
        });
        AppendLog(L"Provider matrix: executing stage-entropy failure fixture.");
        const HRESULT entropyHr = entropyTask.ExecuteOperation();
        AppendLog(std::format(L"Provider matrix: stage-entropy fixture returned 0x{:08X}.", static_cast<unsigned long>(entropyHr)));
        const unsigned long entropyAttempts = TakeFileOpsBridgeFailNextStageEntropyAttemptsForSelfTest();
        SetFileOpsBridgeFailNextStageEntropyForSelfTest(0u);
        std::string entropySourceText;
        wil::com_ptr<IFileSystemIO> entropyDummyIo;
        const bool entropySourceRetained = SUCCEEDED(state.fsDummy->QueryInterface(IID_PPV_ARGS(entropyDummyIo.addressof()))) && entropyDummyIo &&
                                           ReadFileTextFsIo(entropyDummyIo, entropyDummyFile, entropySourceText) &&
                                           entropySourceText == "entropy-source";
        std::error_code entropyEc;
        if (! require(entropyHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) && entropyAttempts == 1u &&
                          entropyTask._perf.bridgeStageCreateCount.load(std::memory_order_acquire) == 0u && entropySourceRetained &&
                          ! std::filesystem::exists(entropyDestination, entropyEc) && ! entropyEc,
                      L"CSPRNG failure must fail before exclusive writer creation, retain the source, and publish no destination."))
        {
            return true;
        }

        const std::filesystem::path bulkRejectedSource = state.tempRoot / L"bulk-transfer-rejected-source.bin";
        const std::filesystem::path bulkRejectedDestinationRoot = state.tempRoot / L"bulk-transfer-rejected-destination";
        const std::filesystem::path bulkRejectedDestination = bulkRejectedDestinationRoot / bulkRejectedSource.filename();
        std::error_code rejectedRouteEc;
        std::filesystem::create_directories(bulkRejectedDestinationRoot, rejectedRouteEc);
        if (! require(! rejectedRouteEc && SelfTest::WriteTextFile(bulkRejectedSource, "bulk-rejected"),
                      L"Bulk transfer rejection fixture should create its source and destination folder."))
        {
            return true;
        }
        FileOperations::TransferPlan bulkRejectedPlan{};
        bulkRejectedPlan.intent              = FileOperations::TransferIntent::Copy;
        bulkRejectedPlan.strategy            = FileOperations::OperationStrategy::Copy;
        bulkRejectedPlan.sourceEndpoint      = localValidationEndpoint;
        bulkRejectedPlan.destinationEndpoint = localValidationEndpoint;
        bulkRejectedPlan.selectedItems.push_back({.providerPath = bulkRejectedSource.native()});
        bulkRejectedPlan.destination.providerFolderPath = bulkRejectedDestinationRoot.native();
        FolderWindow::FileOperationState::Task bulkRejectedTask(*state.fileOps);
        bulkRejectedTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(bulkRejectedPlan)}}));
        bulkRejectedTask._operation         = FILESYSTEM_COPY;
        bulkRejectedTask._executionMode     = FolderWindow::FileOperationState::ExecutionMode::BulkItems;
        bulkRejectedTask._fileSystem        = state.fsLocal;
        bulkRejectedTask._sourcePaths       = {bulkRejectedSource};
        bulkRejectedTask._destinationFolder = bulkRejectedDestinationRoot;
        AppendLog(L"Provider matrix: executing bulk-transfer rejection fixture.");
        const HRESULT bulkRejectedHr = bulkRejectedTask.ExecuteOperation();
        AppendLog(std::format(L"Provider matrix: bulk-transfer fixture returned 0x{:08X}.", static_cast<unsigned long>(bulkRejectedHr)));
        if (! require(bulkRejectedHr == E_UNEXPECTED && std::filesystem::exists(bulkRejectedSource, rejectedRouteEc) &&
                          ! std::filesystem::exists(bulkRejectedDestination, rejectedRouteEc),
                      L"Bulk Copy/Move execution must fail before provider I/O because it has no per-item identity guard."))
        {
            return true;
        }

        // Beeline: the endpoint tuple decides the route. A Native plan stays a provider rename
        // whatever destination provider object the task carries for the same qualified endpoint.
        const std::filesystem::path nativeExplicitSource = state.tempRoot / L"native-explicit-destination-source.bin";
        const std::filesystem::path nativeExplicitDestinationRoot = state.tempRoot / L"native-explicit-destination";
        const std::filesystem::path nativeExplicitDestination = nativeExplicitDestinationRoot / nativeExplicitSource.filename();
        std::filesystem::create_directories(nativeExplicitDestinationRoot, rejectedRouteEc);
        if (! require(! rejectedRouteEc && SelfTest::WriteTextFile(nativeExplicitSource, "native-explicit-destination"),
                      L"Native explicit-destination fixture should create its source and destination folder."))
        {
            return true;
        }
        FileOperations::TransferPlan nativeExplicitPlan{};
        nativeExplicitPlan.intent              = FileOperations::TransferIntent::Move;
        nativeExplicitPlan.strategy            = FileOperations::OperationStrategy::Native;
        nativeExplicitPlan.sourceEndpoint      = localValidationEndpoint;
        nativeExplicitPlan.destinationEndpoint = localValidationEndpoint;
        nativeExplicitPlan.selectedItems.push_back({.providerPath = nativeExplicitSource.native()});
        nativeExplicitPlan.destination.providerFolderPath = nativeExplicitDestinationRoot.native();
        FolderWindow::FileOperationState::Task nativeExplicitTask(*state.fileOps);
        nativeExplicitTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(nativeExplicitPlan)}}));
        nativeExplicitTask._operation              = FILESYSTEM_MOVE;
        nativeExplicitTask._executionMode          = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        nativeExplicitTask._fileSystem             = state.fsLocal;
        nativeExplicitTask._destinationFileSystem  = state.fsLocal;
        nativeExplicitTask._sourcePaths            = {nativeExplicitSource};
        nativeExplicitTask._destinationFolder      = nativeExplicitDestinationRoot;
        nativeExplicitTask._sourcePathAttributesHint = {GetFileAttributesW(nativeExplicitSource.c_str())};
        AppendLog(L"Provider matrix: executing Native explicit-destination fixture.");
        const HRESULT nativeExplicitHr = nativeExplicitTask.ExecuteOperation();
        AppendLog(std::format(L"Provider matrix: Native explicit-destination fixture returned 0x{:08X}.",
                              static_cast<unsigned long>(nativeExplicitHr)));
        if (! require(nativeExplicitHr == S_OK && ! std::filesystem::exists(nativeExplicitSource, rejectedRouteEc) &&
                          std::filesystem::exists(nativeExplicitDestination, rejectedRouteEc),
                      L"A provider-native Move with an explicit destination provider object on the same endpoint must rename, never enter the bridge."))
        {
            return true;
        }

        const std::filesystem::path racedMergeSource      = state.tempRoot / L"native-race-rename-merge-source";
        const std::filesystem::path racedMergeChild       = racedMergeSource / L"child.bin";
        const std::filesystem::path racedMergeDestination = nativeExplicitDestinationRoot / racedMergeSource.filename();
        std::filesystem::create_directories(racedMergeSource, rejectedRouteEc);
        if (! require(! rejectedRouteEc && SelfTest::WriteTextFile(racedMergeChild, "relocated-by-rename"),
                      L"Native race rename-merge fixture should create its source tree."))
        {
            return true;
        }
        FileOperations::TransferPlan racedMergePlan{};
        racedMergePlan.intent              = FileOperations::TransferIntent::Move;
        racedMergePlan.strategy            = FileOperations::OperationStrategy::Native;
        racedMergePlan.sourceEndpoint      = localValidationEndpoint;
        racedMergePlan.destinationEndpoint = localValidationEndpoint;
        racedMergePlan.selectedItems.push_back({.providerPath = racedMergeSource.native()});
        racedMergePlan.destination.providerFolderPath = nativeExplicitDestinationRoot.native();
        FolderWindow::FileOperationState::Task racedMergeTask(*state.fileOps);
        racedMergeTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(racedMergePlan)}}));
        racedMergeTask._operation                = FILESYSTEM_MOVE;
        racedMergeTask._executionMode            = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        racedMergeTask._fileSystem               = state.fsLocal;
        racedMergeTask._sourcePaths              = {racedMergeSource};
        racedMergeTask._destinationFolder        = nativeExplicitDestinationRoot;
        racedMergeTask._flags                    = FILESYSTEM_FLAG_RECURSIVE;
        racedMergeTask._sourcePathAttributesHint = {GetFileAttributesW(racedMergeSource.c_str())};
        SetFileOpsNativeMoveCreateDirectoryRaceForSelfTest(1u);
        AppendLog(L"Provider matrix: executing the Native directory race that continues as a rename merge.");
        const HRESULT racedMergeHr = racedMergeTask.ExecuteOperation();
        AppendLog(std::format(L"Provider matrix: Native race rename-merge fixture returned 0x{:08X}.", static_cast<unsigned long>(racedMergeHr)));
        const unsigned long racedMergeAttempts = TakeFileOpsNativeMoveCreateDirectoryRaceAttemptsForSelfTest();
        SetFileOpsNativeMoveCreateDirectoryRaceForSelfTest(0u);
        if (! require(racedMergeHr == S_OK && racedMergeAttempts == 1u && ! std::filesystem::exists(racedMergeSource, rejectedRouteEc) &&
                          std::filesystem::exists(racedMergeDestination / racedMergeChild.filename(), rejectedRouteEc),
                      L"A raced destination directory continues the Native item as a rename merge: the child relocates and the emptied source is removed."))
        {
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.NativeDirectoryRaceContinuesAsRenameMerge",
                          L"known-noncommit;rename-merge;source-removed",
                          0u,
                          racedMergeAttempts,
                          1u,
                          racedMergeHr);

        const std::filesystem::path managedMoveSource = state.tempRoot / L"managed-move-source.bin";
        const std::filesystem::path managedMoveDestination = nativeExplicitDestinationRoot / managedMoveSource.filename();
        if (! require(SelfTest::WriteTextFile(managedMoveSource, "managed-move-exact"),
                      L"Managed Move fixture should create its source."))
        {
            return true;
        }
        FileOperations::TransferPlan managedMovePlan{};
        managedMovePlan.intent              = FileOperations::TransferIntent::Move;
        managedMovePlan.strategy            = FileOperations::OperationStrategy::Managed;
        managedMovePlan.sourceEndpoint      = localValidationEndpoint;
        managedMovePlan.destinationEndpoint = localValidationEndpoint;
        managedMovePlan.selectedItems.push_back({.providerPath = managedMoveSource.native()});
        managedMovePlan.destination.providerFolderPath = nativeExplicitDestinationRoot.native();
        FolderWindow::FileOperationState::Task managedMoveTask(*state.fileOps);
        managedMoveTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(managedMovePlan)}}));
        managedMoveTask._operation              = FILESYSTEM_MOVE;
        managedMoveTask._executionMode          = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        managedMoveTask._fileSystem             = state.fsLocal;
        managedMoveTask._destinationFileSystem  = state.fsLocal;
        managedMoveTask._sourcePaths            = {managedMoveSource};
        managedMoveTask._destinationFolder      = nativeExplicitDestinationRoot;
        managedMoveTask._sourcePathAttributesHint = {GetFileAttributesW(managedMoveSource.c_str())};
        AppendLog(L"Provider matrix: executing Managed Move fixture.");
        const HRESULT managedMoveHr = managedMoveTask.ExecuteOperation();
        AppendLog(std::format(L"Provider matrix: Managed Move fixture returned 0x{:08X}.",
                              static_cast<unsigned long>(managedMoveHr)));
        wil::com_ptr<IFileSystemIO> managedLocalIo;
        std::string managedDestinationText;
        if (! require(managedMoveHr == S_OK && ! std::filesystem::exists(managedMoveSource, rejectedRouteEc) &&
                          std::filesystem::exists(managedMoveDestination, rejectedRouteEc) &&
                          SUCCEEDED(state.fsLocal->QueryInterface(IID_PPV_ARGS(managedLocalIo.addressof()))) && managedLocalIo &&
                          ReadFileTextFsIo(managedLocalIo, managedMoveDestination.native(), managedDestinationText) &&
                          managedDestinationText == "managed-move-exact",
                      L"Managed Move must publish through retained destination authority and delete only its exact bound source."))
        {
            return true;
        }

        const std::filesystem::path managedBusySource = state.tempRoot / L"managed-move-busy-source.bin";
        const std::filesystem::path managedBusyDestination = nativeExplicitDestinationRoot / managedBusySource.filename();
        if (! require(SelfTest::WriteTextFile(managedBusySource, "managed-source-kept"),
                      L"Managed Move source-kept fixture should create its source."))
        {
            return true;
        }
        wil::unique_handle externalWriter(CreateFileW(managedBusySource.c_str(),
                                                       GENERIC_WRITE,
                                                       FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                       nullptr,
                                                       OPEN_EXISTING,
                                                       FILE_ATTRIBUTE_NORMAL,
                                                       nullptr));
        if (! require(static_cast<bool>(externalWriter),
                      L"Managed Move source-kept fixture should retain an external write handle."))
        {
            return true;
        }
        FileOperations::TransferPlan managedBusyPlan{};
        managedBusyPlan.intent              = FileOperations::TransferIntent::Move;
        managedBusyPlan.strategy            = FileOperations::OperationStrategy::Managed;
        managedBusyPlan.sourceEndpoint      = localValidationEndpoint;
        managedBusyPlan.destinationEndpoint = localValidationEndpoint;
        managedBusyPlan.selectedItems.push_back({.providerPath = managedBusySource.native()});
        managedBusyPlan.destination.providerFolderPath = nativeExplicitDestinationRoot.native();
        FolderWindow::FileOperationState::Task managedBusyTask(*state.fileOps);
        managedBusyTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(managedBusyPlan)}}));
        managedBusyTask._operation                = FILESYSTEM_MOVE;
        managedBusyTask._executionMode            = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        managedBusyTask._fileSystem               = state.fsLocal;
        managedBusyTask._destinationFileSystem    = state.fsLocal;
        managedBusyTask._sourcePaths              = {managedBusySource};
        managedBusyTask._destinationFolder        = nativeExplicitDestinationRoot;
        managedBusyTask._sourcePathAttributesHint = {GetFileAttributesW(managedBusySource.c_str())};
        AppendLog(L"Provider matrix: executing Managed source-kept fixture.");
        const HRESULT managedBusyHr = managedBusyTask.ExecuteOperation();
        AppendLog(std::format(L"Provider matrix: Managed source-kept fixture returned 0x{:08X}.",
                              static_cast<unsigned long>(managedBusyHr)));
        std::string managedBusyDestinationText;
        if (! require(managedBusyHr == S_FALSE && std::filesystem::exists(managedBusySource, rejectedRouteEc) &&
                          std::filesystem::exists(managedBusyDestination, rejectedRouteEc) &&
                          ReadFileTextFsIo(managedLocalIo, managedBusyDestination.native(), managedBusyDestinationText) &&
                          managedBusyDestinationText == "managed-source-kept",
                      L"Managed Move must downgrade to Copied/source-kept when a bound source cannot exclude concurrent writers."))
        {
            return true;
        }

        const std::filesystem::path managedRaceSource = state.tempRoot / L"managed-move-cleanup-race-source.bin";
        const std::filesystem::path managedRaceReplacement = state.tempRoot / L"managed-move-cleanup-race-replacement.bin";
        const std::filesystem::path managedRaceDestination = nativeExplicitDestinationRoot / managedRaceSource.filename();
        if (! require(SelfTest::WriteTextFile(managedRaceSource, "managed-cleanup-race"),
                      L"Managed Move cleanup-race fixture should create its source."))
        {
            return true;
        }
        FileOperations::TransferPlan managedRacePlan{};
        managedRacePlan.intent              = FileOperations::TransferIntent::Move;
        managedRacePlan.strategy            = FileOperations::OperationStrategy::Managed;
        managedRacePlan.sourceEndpoint      = localValidationEndpoint;
        managedRacePlan.destinationEndpoint = localValidationEndpoint;
        managedRacePlan.selectedItems.push_back({.providerPath = managedRaceSource.native()});
        managedRacePlan.destination.providerFolderPath = nativeExplicitDestinationRoot.native();
        FolderWindow::FileOperationState::Task managedRaceTask(*state.fileOps);
        managedRaceTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(managedRacePlan)}}));
        managedRaceTask._operation                = FILESYSTEM_MOVE;
        managedRaceTask._executionMode            = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        managedRaceTask._fileSystem               = state.fsLocal;
        managedRaceTask._destinationFileSystem    = state.fsLocal;
        managedRaceTask._sourcePaths              = {managedRaceSource};
        managedRaceTask._destinationFolder        = nativeExplicitDestinationRoot;
        managedRaceTask._sourcePathAttributesHint = {GetFileAttributesW(managedRaceSource.c_str())};
        AppendLog(L"Provider matrix: arming Managed cleanup-race pause.");
        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(true);
        HRESULT managedRaceHr = E_PENDING;
        std::jthread managedRaceWorker([&](std::stop_token) noexcept
        {
            AppendLog(L"Provider matrix: Managed cleanup-race worker entered ExecuteOperation.");
            managedRaceHr = managedRaceTask.ExecuteOperation();
            AppendLog(std::format(L"Provider matrix: Managed cleanup-race worker returned 0x{:08X}.",
                                  static_cast<unsigned long>(managedRaceHr)));
        });
        const ULONGLONG managedRaceWaitDeadline = GetTickCount64() + 4'000ull;
        while (! HasFileOpsBridgeMoveSourceCleanupPauseEnteredForSelfTest() && GetTickCount64() < managedRaceWaitDeadline)
        {
            Sleep(10);
        }
        const bool cleanupPauseEntered = HasFileOpsBridgeMoveSourceCleanupPauseEnteredForSelfTest();
        AppendLog(std::format(L"Provider matrix: Managed cleanup-race pause entered={0}.", cleanupPauseEntered));
        SetLastError(ERROR_SUCCESS);
        const BOOL replacementMove = cleanupPauseEntered
            ? MoveFileExW(managedRaceSource.c_str(), managedRaceReplacement.c_str(), MOVEFILE_REPLACE_EXISTING)
            : FALSE;
        const DWORD replacementMoveError = GetLastError();
        AppendLog(std::format(L"Provider matrix: Managed cleanup-race replacement move result={0}, error={1}.",
                              replacementMove,
                              replacementMoveError));
        ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest();
        SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(false);
        AppendLog(L"Provider matrix: Managed cleanup-race pause released; joining worker.");
        managedRaceWorker.join();
        AppendLog(L"Provider matrix: Managed cleanup-race worker joined.");
        std::string managedRaceDestinationText;
        if (! require(cleanupPauseEntered && replacementMove == FALSE &&
                          (replacementMoveError == ERROR_SHARING_VIOLATION || replacementMoveError == ERROR_ACCESS_DENIED) &&
                          managedRaceHr == S_OK && ! std::filesystem::exists(managedRaceSource, rejectedRouteEc) &&
                          ! std::filesystem::exists(managedRaceReplacement, rejectedRouteEc) &&
                          ReadFileTextFsIo(managedLocalIo, managedRaceDestination.native(), managedRaceDestinationText) &&
                          managedRaceDestinationText == "managed-cleanup-race",
                      L"Managed Move must retain exact source authority through cleanup so a pathname replacement cannot enter the publish/delete window."))
        {
            return true;
        }

        const auto makeManagedCleanupTask = [&](const std::filesystem::path& source)
        {
            FileOperations::TransferPlan plan{};
            plan.intent              = FileOperations::TransferIntent::Move;
            plan.strategy            = FileOperations::OperationStrategy::Managed;
            plan.sourceEndpoint      = localValidationEndpoint;
            plan.destinationEndpoint = localValidationEndpoint;
            plan.selectedItems.push_back({.providerPath = source.native()});
            plan.destination.providerFolderPath = nativeExplicitDestinationRoot.native();
            auto task = std::make_unique<FolderWindow::FileOperationState::Task>(*state.fileOps);
            task->StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
                FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(plan)}}));
            task->_operation                = FILESYSTEM_MOVE;
            task->_executionMode            = FolderWindow::FileOperationState::ExecutionMode::PerItem;
            task->_fileSystem               = state.fsLocal;
            task->_destinationFileSystem    = state.fsLocal;
            task->_sourcePaths              = {source};
            task->_destinationFolder        = nativeExplicitDestinationRoot;
            task->_sourcePathAttributesHint = {GetFileAttributesW(source.c_str())};
            task->InitializeSourceItemResultBuilders();
            task->MarkSourceItemsMutationPossible();
            return task;
        };

        using CleanupBucket = FolderWindow::FileOperationState::Task::ConflictBucket;
        struct CleanupFailureScenario final
        {
            HRESULT status = E_FAIL;
            CleanupBucket bucket = CleanupBucket::Unknown;
            std::wstring_view name;
        };
        const std::array cleanupFailureScenarios{
            CleanupFailureScenario{HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), CleanupBucket::SharingViolation, L"sharing"},
            CleanupFailureScenario{HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED), CleanupBucket::AccessDenied, L"access"},
            CleanupFailureScenario{HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE), CleanupBucket::PathTooLong, L"path"},
            CleanupFailureScenario{HRESULT_FROM_WIN32(ERROR_DISK_FULL), CleanupBucket::DiskFull, L"disk"},
        };
        unsigned long cleanupMatrixPromptCount = 0u;
        for (size_t scenarioIndex = 0u; scenarioIndex < cleanupFailureScenarios.size(); ++scenarioIndex)
        {
            const CleanupFailureScenario& scenario = cleanupFailureScenarios[scenarioIndex];
            for (const bool chooseRetry : {true, false})
            {
                const std::wstring decisionName = chooseRetry ? L"retry" : L"skip";
                const std::filesystem::path source = state.tempRoot /
                    std::format(L"managed-cleanup-{}-{}-{}.bin", scenario.name, decisionName, scenarioIndex);
                const std::filesystem::path destination = nativeExplicitDestinationRoot / source.filename();
                const std::string payload = std::format("managed-cleanup-{}-{}", scenarioIndex, chooseRetry ? "retry" : "skip");
                if (! require(SelfTest::WriteTextFile(source, payload),
                              std::format(L"Managed cleanup {} {} fixture should create its source.", scenario.name, decisionName)))
                {
                    return true;
                }

                AppendLog(std::format(L"Provider matrix: starting Managed cleanup {0} {1} fixture.", scenario.name, decisionName));
                auto cleanupTask = makeManagedCleanupTask(source);
                SetFileOpsManagedCleanupKnownNonCommitForSelfTest(scenario.status, 1u);
                HRESULT cleanupHr = E_PENDING;
                std::jthread cleanupWorker([&](std::stop_token) noexcept
                {
                    AppendLog(std::format(L"Provider matrix: Managed cleanup {0} {1} worker entered.", scenario.name, decisionName));
                    cleanupHr = cleanupTask->ExecuteOperation();
                    AppendLog(std::format(L"Provider matrix: Managed cleanup {0} {1} worker returned 0x{2:08X}.",
                                          scenario.name,
                                          decisionName,
                                          static_cast<unsigned long>(cleanupHr)));
                });

                std::optional<FolderWindow::FileOperationState::Task::ConflictPromptState> prompt;
                const ULONGLONG deadline = GetTickCount64() + 4'000ull;
                while (! prompt.has_value() && GetTickCount64() < deadline)
                {
                    prompt = TryGetConflictPromptCopy(cleanupTask.get());
                    if (! prompt.has_value())
                    {
                        Sleep(10);
                    }
                }

                const bool promptValid = prompt.has_value() && prompt->status == scenario.status && prompt->bucket == scenario.bucket &&
                    PromptHasAction(prompt.value(), FolderWindow::FileOperationState::Task::ConflictAction::Retry) &&
                    PromptHasAction(prompt.value(), FolderWindow::FileOperationState::Task::ConflictAction::Skip);
                AppendLog(std::format(L"Provider matrix: Managed cleanup {0} {1} prompt observed={2}, valid={3}.",
                                      scenario.name,
                                      decisionName,
                                      prompt.has_value(),
                                      promptValid));
                if (promptValid)
                {
                    cleanupTask->SubmitConflictDecision(chooseRetry ? FolderWindow::FileOperationState::Task::ConflictAction::Retry
                                                                    : FolderWindow::FileOperationState::Task::ConflictAction::Skip,
                                                        false);
                    ++cleanupMatrixPromptCount;
                }
                else
                {
                    cleanupTask->RequestCancel();
                }
                AppendLog(std::format(L"Provider matrix: Managed cleanup {0} {1} decision submitted; joining worker.",
                                      scenario.name,
                                      decisionName));
                cleanupWorker.join();
                AppendLog(std::format(L"Provider matrix: Managed cleanup {0} {1} worker joined.", scenario.name, decisionName));

                const unsigned long injectedAttempts = TakeFileOpsManagedCleanupKnownNonCommitAttemptsForSelfTest();
                SetFileOpsManagedCleanupKnownNonCommitForSelfTest(HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), 0u);
                AppendLog(std::format(L"Provider matrix: Managed cleanup {0} {1} injection count collected.", scenario.name, decisionName));
                std::string destinationText;
                std::error_code scenarioEc;
                const bool sourceExists = std::filesystem::exists(source, scenarioEc) && ! scenarioEc;
                AppendLog(std::format(L"Provider matrix: Managed cleanup {0} {1} source existence checked.", scenario.name, decisionName));
                scenarioEc.clear();
                const bool destinationMatches = ReadFileTextFsIo(managedLocalIo, destination.native(), destinationText) && destinationText == payload;
                AppendLog(std::format(L"Provider matrix: Managed cleanup {0} {1} destination checked.", scenario.name, decisionName));
                const bool terminalStateMatches = chooseRetry ? cleanupHr == S_OK && ! sourceExists
                                                              : cleanupHr == S_FALSE && sourceExists;
                if (! require(promptValid && injectedAttempts == 1u && terminalStateMatches && destinationMatches,
                              std::format(L"Managed cleanup {} {} must prompt in the source-cleanup bucket, preserve publication, and {} the source.",
                                          scenario.name,
                                          decisionName,
                                          chooseRetry ? L"remove" : L"retain")))
                {
                    return true;
                }
                AppendLog(std::format(L"Provider matrix: Managed cleanup {0} {1} fixture validated.", scenario.name, decisionName));
            }
        }
        AppendLog(L"Provider matrix: emitting Managed cleanup matrix perf evidence.");
        Debug::Perf::Emit(L"FileOps.SelfTest.ManagedCleanupKnownNonCommitMatrix",
                          L"sharing+access+path+disk;retry+skip;delete-only",
                          cleanupMatrixPromptCount,
                          cleanupFailureScenarios.size() * 2u,
                          0u,
                          S_OK);
        AppendLog(L"Provider matrix: Managed cleanup matrix perf evidence emitted.");

        const std::filesystem::path managedUnknownSource = state.tempRoot / L"managed-move-cleanup-unknown-source.bin";
        const std::filesystem::path managedUnknownDestination = nativeExplicitDestinationRoot / managedUnknownSource.filename();
        if (! require(SelfTest::WriteTextFile(managedUnknownSource, "managed-cleanup-unknown"),
                      L"Managed Move unknown-cleanup fixture should create its source."))
        {
            return true;
        }
        auto managedUnknownTask = makeManagedCleanupTask(managedUnknownSource);
        SetFileOpsManagedCleanupUnknownOutcomeForSelfTest(1u);
        AppendLog(L"Provider matrix: executing Managed unknown-cleanup fixture.");
        const HRESULT managedUnknownHr = managedUnknownTask->ExecuteOperation();
        const HRESULT managedUnknownTypedHr = managedUnknownTask->FinalizeTypedItemResults(managedUnknownHr);
        AppendLog(std::format(L"Provider matrix: Managed unknown-cleanup fixture returned 0x{:08X}.",
                              static_cast<unsigned long>(managedUnknownHr)));
        const unsigned long managedUnknownAttempts = TakeFileOpsManagedCleanupUnknownOutcomeAttemptsForSelfTest();
        SetFileOpsManagedCleanupUnknownOutcomeForSelfTest(0u);
        std::string managedUnknownDestinationText;
        const bool unknownAxes = managedUnknownTask->_sourceItemResultBuilders.size() == 1u &&
            managedUnknownTask->_sourceItemResultBuilders.front().terminal.has_value() &&
            managedUnknownTask->_sourceItemResultBuilders.front().terminal->publication == FileOperations::PublicationState::Published &&
            managedUnknownTask->_sourceItemResultBuilders.front().terminal->sourceDisposition == FileOperations::SourceDisposition::Unknown &&
            managedUnknownTask->_sourceItemResultBuilders.front().terminal->completion == FileOperations::ItemCompletion::Indeterminate;
        if (! require(managedUnknownHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) &&
                          managedUnknownTypedHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) && unknownAxes && managedUnknownAttempts == 1u &&
                          std::filesystem::exists(managedUnknownSource, rejectedRouteEc) &&
                          ReadFileTextFsIo(managedLocalIo, managedUnknownDestination.native(), managedUnknownDestinationText) &&
                          managedUnknownDestinationText == "managed-cleanup-unknown" &&
                          ! TryGetConflictPromptCopy(managedUnknownTask.get()).has_value(),
                      L"Managed Move must report an unknown cleanup outcome as indeterminate without recopied-item Retry or another Delete call."))
        {
            return true;
        }

        const std::filesystem::path managedTreeSource = state.tempRoot / L"managed-move-tree";
        const std::filesystem::path managedTreeChild = managedTreeSource / L"nested" / L"child.bin";
        const std::filesystem::path managedTreeDestination = nativeExplicitDestinationRoot / managedTreeSource.filename();
        std::filesystem::create_directories(managedTreeChild.parent_path(), rejectedRouteEc);
        if (! require(! rejectedRouteEc && SelfTest::WriteTextFile(managedTreeChild, "managed-tree"),
                      L"Managed Move tree fixture should create its nested source."))
        {
            return true;
        }
        FileOperations::TransferPlan managedTreePlan{};
        managedTreePlan.intent              = FileOperations::TransferIntent::Move;
        managedTreePlan.strategy            = FileOperations::OperationStrategy::Managed;
        managedTreePlan.sourceEndpoint      = localValidationEndpoint;
        managedTreePlan.destinationEndpoint = localValidationEndpoint;
        managedTreePlan.selectedItems.push_back({.providerPath = managedTreeSource.native()});
        managedTreePlan.destination.providerFolderPath = nativeExplicitDestinationRoot.native();
        FolderWindow::FileOperationState::Task managedTreeTask(*state.fileOps);
        managedTreeTask.StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(managedTreePlan)}}));
        managedTreeTask._operation                = FILESYSTEM_MOVE;
        managedTreeTask._executionMode            = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        managedTreeTask._fileSystem               = state.fsLocal;
        managedTreeTask._destinationFileSystem    = state.fsLocal;
        managedTreeTask._sourcePaths              = {managedTreeSource};
        managedTreeTask._destinationFolder        = nativeExplicitDestinationRoot;
        managedTreeTask._sourcePathAttributesHint = {GetFileAttributesW(managedTreeSource.c_str())};
        AppendLog(std::format(L"Provider matrix: executing Managed directory Move fixture with {0} bridge-budget bytes in use.",
                              GetFileOpsBridgeBufferBudgetInUseForSelfTest()));
        std::atomic<bool> managedTreeDone{false};
        HRESULT managedTreeHr = E_PENDING;
        std::jthread managedTreeWorker([&](std::stop_token) noexcept
        {
            managedTreeHr = managedTreeTask.ExecuteOperation();
            managedTreeDone.store(true, std::memory_order_release);
        });
        const ULONGLONG managedTreeDeadline = GetTickCount64() + 8'000ull;
        while (! managedTreeDone.load(std::memory_order_acquire) && GetTickCount64() < managedTreeDeadline)
        {
            Sleep(10);
        }
        if (! managedTreeDone.load(std::memory_order_acquire))
        {
            const std::optional<FolderWindow::FileOperationState::Task::ConflictPromptState> managedTreePrompt =
                TryGetConflictPromptCopy(&managedTreeTask);
            AppendLog(std::format(
                L"Provider matrix: Managed directory Move exceeded its bound with {0} bridge-budget bytes in use; prompt={1}, status=0x{2:08X}, bucket={3}; cancelling.",
                GetFileOpsBridgeBufferBudgetInUseForSelfTest(),
                managedTreePrompt.has_value(),
                managedTreePrompt.has_value() ? static_cast<unsigned long>(managedTreePrompt->status) : 0ul,
                managedTreePrompt.has_value() ? static_cast<unsigned int>(managedTreePrompt->bucket) : 0u));
            managedTreeTask.RequestCancel();
        }
        managedTreeWorker.join();
        AppendLog(std::format(L"Provider matrix: Managed directory Move fixture returned 0x{:08X}.",
                              static_cast<unsigned long>(managedTreeHr)));
        std::string managedTreeDestinationText;
        const std::filesystem::path managedTreeDestinationChild = managedTreeDestination / L"nested" / L"child.bin";
        if (! require(managedTreeHr == S_OK && ! std::filesystem::exists(managedTreeSource, rejectedRouteEc) &&
                          std::filesystem::exists(managedTreeDestinationChild, rejectedRouteEc) &&
                          ReadFileTextFsIo(managedLocalIo, managedTreeDestinationChild.native(), managedTreeDestinationText) &&
                          managedTreeDestinationText == "managed-tree",
                      L"Managed directory Move must delete exact child files and then remove exact directories non-recursively in post-order."))
        {
            return true;
        }

        if (! require(dummyCaps.copyOperation && dummyCaps.moveOperation && dummyCaps.nativeMoveOperation && ! dummyCaps.renameOperation &&
                          ! dummyCaps.deleteOperation && ! dummyCaps.recycleOperation && dummyCaps.read && dummyCaps.write,
                      L"FileSystemDummy should keep deterministic Copy/Move/read/write support but clear product Delete/Recycle/Rename claims.") ||
            ! require(dummyCaps.copyMoveMax == 4u && dummyCaps.deleteMax == 8u && dummyCaps.deleteRecycleMax == 2u,
                      L"FileSystemDummy should advertise the expected deterministic concurrency matrix.") ||
            ! require(dummyCaps.exportCopyWildcard && dummyCaps.exportMoveWildcard && dummyCaps.importCopyWildcard && dummyCaps.importMoveWildcard,
                      L"FileSystemDummy should advertise wildcard cross-filesystem copy/move import/export.") ||
            ! requirePathIdentity(dummyCaps, L"FileSystemDummy", true, L"ordinalIgnoreCase", L'\\', L"\\/", L"supported"))
        {
            return true;
        }
        if (! require(! CanSameFileSystemOperation(state.fsDummy, L"/", FILESYSTEM_DELETE, kPluginIdDummy) &&
                          ! CanSameFileSystemOperation(
                              state.fsDummy, L"/", FILESYSTEM_DELETE, kPluginIdDummy, FILESYSTEM_FLAG_USE_RECYCLE_BIN) &&
                          ! CanSameFileSystemOperation(state.fsDummy, L"/", FILESYSTEM_RENAME, kPluginIdDummy),
                      L"FileSystemDummy host admission should reject Permanent Delete, Recycle, and Rename before task creation."))
        {
            return true;
        }

        wil::com_ptr<IFileSystemIO> dummyNativeIo;
        if (! require(SUCCEEDED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyNativeIo.addressof()))) && dummyNativeIo,
                      L"Dummy NativeOnly Move fixture should expose file I/O."))
        {
            return true;
        }
        static_cast<void>(state.fsDummy->DeleteItem(L"/native-mode-source", FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr));
        static_cast<void>(state.fsDummy->DeleteItem(L"/native-mode-destination", FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr));
        if (! require(EnsureDummyFolderExists(state.fsDummy.get(), L"/native-mode-source") &&
                          EnsureDummyFolderExists(state.fsDummy.get(), L"/native-mode-destination") &&
                          DummyWriteTextFile(state.fsDummy.get(), L"/native-mode-source/single.txt", "single") &&
                          DummyWriteTextFile(state.fsDummy.get(), L"/native-mode-source/batch.txt", "batch"),
                      L"Dummy NativeOnly Move fixture should create deterministic source and destination objects."))
        {
            return true;
        }
        FileSystemOptions dummyNativeOptions{};
        dummyNativeOptions.sizeBytes = sizeof(FileSystemOptions);
        dummyNativeOptions.moveMode  = FILESYSTEM_MOVE_NATIVE_ONLY;
        const HRESULT dummyNativeSingleHr = state.fsDummy->MoveItem(L"/native-mode-source/single.txt",
                                                                    L"/native-mode-destination/single.txt",
                                                                    FILESYSTEM_FLAG_NONE,
                                                                    &dummyNativeOptions,
                                                                    nullptr,
                                                                    nullptr);
        const wchar_t* dummyBatchSources[] = {L"/native-mode-source/batch.txt"};
        const HRESULT dummyNativeBatchHr = state.fsDummy->MoveItems(dummyBatchSources,
                                                                    1u,
                                                                    L"/native-mode-destination",
                                                                    FILESYSTEM_FLAG_NONE,
                                                                    &dummyNativeOptions,
                                                                    nullptr,
                                                                    nullptr);
        std::string dummySingleText;
        std::string dummyBatchText;
        std::string dummyRetainedText;
        if (! require(dummyNativeSingleHr == S_OK && dummyNativeBatchHr == S_OK &&
                          ReadFileTextFsIo(dummyNativeIo, L"/native-mode-destination/single.txt", dummySingleText) && dummySingleText == "single" &&
                          ReadFileTextFsIo(dummyNativeIo, L"/native-mode-destination/batch.txt", dummyBatchText) && dummyBatchText == "batch" &&
                          ! ReadFileTextFsIo(dummyNativeIo, L"/native-mode-source/single.txt", dummyRetainedText) &&
                          ! ReadFileTextFsIo(dummyNativeIo, L"/native-mode-source/batch.txt", dummyRetainedText),
                      L"Dummy MoveItem/MoveItems NativeOnly routes should move nodes without a copy-delete fallback."))
        {
            return true;
        }
        FileSystemOptions dummyInvalidMoveOptions = dummyNativeOptions;
        dummyInvalidMoveOptions.moveMode = static_cast<FileSystemMoveMode>(2u);
        const HRESULT dummyInvalidSingleHr = state.fsDummy->MoveItem(L"/missing",
                                                                     L"/native-mode-destination/missing",
                                                                     FILESYSTEM_FLAG_NONE,
                                                                     &dummyInvalidMoveOptions,
                                                                     nullptr,
                                                                     nullptr);
        const HRESULT dummyInvalidBatchHr = state.fsDummy->MoveItems(dummyBatchSources,
                                                                     1u,
                                                                     L"/native-mode-destination",
                                                                     FILESYSTEM_FLAG_NONE,
                                                                     &dummyInvalidMoveOptions,
                                                                     nullptr,
                                                                     nullptr);
        if (! require(dummyInvalidSingleHr == E_INVALIDARG && dummyInvalidBatchHr == E_INVALIDARG,
                      L"Dummy MoveItem/MoveItems should reject an unknown moveMode before mutation."))
        {
            return true;
        }

        if (! require(! sevenZipCaps.copyOperation && ! sevenZipCaps.moveOperation && ! sevenZipCaps.nativeMoveOperation &&
                          ! sevenZipCaps.renameOperation && ! sevenZipCaps.deleteOperation && sevenZipCaps.properties && sevenZipCaps.read &&
                          ! sevenZipCaps.write,
                      L"FileSystem7z should advertise read/properties only for same-provider operations.") ||
            ! require(! CanSameFileSystemOperation(state.fs7z, L"/missing.txt", FILESYSTEM_RENAME, kPluginId7z),
                      L"Host admission should reject a rename:false provider before task publication.") ||
            ! require(sevenZipCaps.copyMoveMax == 1u && sevenZipCaps.deleteMax == 1u && sevenZipCaps.deleteRecycleMax == 1u,
                      L"FileSystem7z should advertise single-stream concurrency limits.") ||
            ! require(sevenZipCaps.exportCopyWildcard && ! sevenZipCaps.exportMoveWildcard && ! sevenZipCaps.importCopyWildcard &&
                          ! sevenZipCaps.importMoveWildcard,
                      L"FileSystem7z should advertise export-copy only for cross-filesystem transfers.") ||
            ! requirePathIdentity(sevenZipCaps, L"FileSystem7z", true, L"ordinalCaseSensitive", L'/', L"/", L"notApplicable"))
        {
            return true;
        }

        constexpr HRESULT kUnsupported   = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        const wchar_t* sevenZipSources[] = {L"/missing.txt"};
        const auto requireUnsupported    = [&](HRESULT hr, std::wstring_view label) noexcept -> bool
        {
            if (hr != kUnsupported)
            {
                Fail(std::format(L"FileSystem7z {} should return ERROR_NOT_SUPPORTED, got 0x{:08X}.", label, static_cast<unsigned long>(hr)));
                return false;
            }
            return true;
        };

        if (! requireUnsupported(state.fs7z->CopyItem(L"/missing.txt", L"/dest.txt", FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr), L"CopyItem") ||
            ! requireUnsupported(state.fs7z->CopyItems(sevenZipSources, 1, L"/dest", FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr), L"CopyItems") ||
            ! requireUnsupported(state.fs7z->MoveItem(L"/missing.txt", L"/dest.txt", FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr), L"MoveItem") ||
            ! requireUnsupported(state.fs7z->MoveItems(sevenZipSources, 1, L"/dest", FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr), L"MoveItems") ||
            ! requireUnsupported(state.fs7z->DeleteItem(L"/missing.txt", FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr), L"DeleteItem") ||
            ! requireUnsupported(state.fs7z->DeleteItems(sevenZipSources, 1, FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr), L"DeleteItems"))
        {
            return true;
        }

        std::vector<FolderWindow::FileOperationState::Task*> tasksBefore;
        state.fileOps->CollectTasks(tasksBefore);
        const HRESULT rejectedStart = state.fileOps->AdmitOperation(FILESYSTEM_COPY,
                                                                    FolderWindow::Pane::Left,
                                                                    std::nullopt,
                                                                    state.fs7z,
                                                                    {std::filesystem::path(L"/missing.txt")},
                                                                    std::filesystem::path(L"/dest"),
                                                                    FILESYSTEM_FLAG_NONE,
                                                                    false,
                                                                    0,
                                                                    FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                    false,
                                                                    nullptr,
                                                                    nullptr,
                                                                    {},
                                                                    {},
                                                                    std::wstring(kPluginId7z),
                                                                    L"7z");
        std::vector<FolderWindow::FileOperationState::Task*> tasksAfter;
        state.fileOps->CollectTasks(tasksAfter);
        if (rejectedStart != kUnsupported || tasksAfter.size() != tasksBefore.size())
        {
            Fail(std::format(L"Host typed admission should reject 7z copy before task creation (hr=0x{:08X} before={} after={}).",
                             static_cast<unsigned long>(rejectedStart),
                             tasksBefore.size(),
                             tasksAfter.size()));
            return true;
        }

        // 8A provider conformance: S3, Microsoft Drive and Curl/SFTP declare the merge/conflict
        // capabilities the host routes on. GetCapabilities needs no live connection, so this is
        // deterministic. The provider-SPECIFIC merge/conflict EXECUTION (2A S3 object-store merge,
        // 2B MS Drive directory-onto-directory move) requires a live/stub backend (out of in-process
        // scope) — covered by the external code review plus the shared host-side merge engine the
        // local/Dummy provider tests exercise. Providers that cannot load here are skipped, not failed.
        const auto loadProviderFs = [&](std::wstring_view pluginId) noexcept -> wil::com_ptr<IFileSystem>
        {
            wil::com_ptr<IFileSystem> fs;
            if (const FileSystemPluginManager::PluginEntry* entry = FindLoadedPluginEntry(pluginId))
            {
                fs = entry->fileSystem;
            }
            if (! fs)
            {
                static_cast<void>(FileSystemPluginManager::GetInstance().EnablePlugin(pluginId, g_settings));
                if (const FileSystemPluginManager::PluginEntry* entry = FindLoadedPluginEntry(pluginId))
                {
                    fs = entry->fileSystem;
                }
            }
            return fs;
        };

        const auto conformCloudProvider = [&](std::wstring_view pluginId,
                                              std::wstring_view providerName,
                                              const bool* expectSameProviderCopy,
                                              bool requireMoveMerge,
                                              bool expectNativeMove,
                                              bool expectDelete,
                                              bool expectRename,
                                              bool expectRecycle,
                                              bool expectMutationAdmission,
                                              bool requireRead,
                                              bool expectedStable,
                                              std::wstring_view expectedComparison,
                                              std::wstring_view expectedCaseOnlyRename) noexcept -> bool
        {
            wil::com_ptr<IFileSystem> fs = loadProviderFs(pluginId);
            if (! fs)
            {
                AppendLog(std::format(L"Provider conformance: {} not loadable in this environment; skipped.", providerName));
                return true;
            }

            ProviderCapabilitySnapshot caps{};
            std::wstring reason;
            if (! TryReadProviderCapabilities(fs.get(), caps, reason))
            {
                Fail(std::format(L"{} GetCapabilities (offline) failed to parse: {}.", providerName, reason));
                return false;
            }
            if ((requireRead && ! caps.read) || caps.copyMoveMax < 1u || caps.deleteMax < 1u || caps.deleteRecycleMax < 1u)
            {
                Fail(std::format(L"{} should advertise expected read support with positive copy/move, delete and recycle concurrency.", providerName));
                return false;
            }
            if (requireMoveMerge && ! caps.moveOperation)
            {
                Fail(std::format(L"{} should advertise Move so its provider-native directory merge remains available.", providerName));
                return false;
            }
            if (caps.nativeMoveOperation != expectNativeMove)
            {
                Fail(std::format(L"{} native Move capability mismatch: declared {} expected {}.",
                                 providerName,
                                 caps.nativeMoveOperation ? 1 : 0,
                                 expectNativeMove ? 1 : 0));
                return false;
            }
            if (expectSameProviderCopy != nullptr && caps.copyOperation != *expectSameProviderCopy)
            {
                Fail(std::format(L"{} same-provider copy capability mismatch: declared {} expected {} (the host routes copy/bridge on this).",
                                 providerName,
                                 caps.copyOperation ? 1 : 0,
                                 *expectSameProviderCopy ? 1 : 0));
                return false;
            }
            if (caps.deleteOperation != expectDelete || caps.renameOperation != expectRename || caps.recycleOperation != expectRecycle)
            {
                Fail(std::format(L"{} destructive capability mismatch: delete={} rename={} recycle={}, expected {}/{}/{}.",
                                 providerName,
                                 caps.deleteOperation ? 1 : 0,
                                 caps.renameOperation ? 1 : 0,
                                 caps.recycleOperation ? 1 : 0,
                                 expectDelete ? 1 : 0,
                                 expectRename ? 1 : 0,
                                 expectRecycle ? 1 : 0));
                return false;
            }
            const bool permanentDeleteAdmitted =
                CanSameFileSystemOperation(fs, L"/", FILESYSTEM_DELETE, pluginId, FILESYSTEM_FLAG_NONE);
            const bool recycleAdmitted =
                CanSameFileSystemOperation(fs, L"/", FILESYSTEM_DELETE, pluginId, FILESYSTEM_FLAG_USE_RECYCLE_BIN);
            const bool renameAdmitted = CanSameFileSystemOperation(fs, L"/", FILESYSTEM_RENAME, pluginId);
            const bool expectedPermanentDeleteAdmission = expectMutationAdmission && expectDelete;
            const bool expectedRecycleAdmission         = expectMutationAdmission && expectRecycle;
            const bool expectedRenameAdmission          = expectMutationAdmission && expectRename;
            if (permanentDeleteAdmitted != expectedPermanentDeleteAdmission || recycleAdmitted != expectedRecycleAdmission ||
                renameAdmitted != expectedRenameAdmission)
            {
                Fail(std::format(L"{} host admission mismatch: permanent={} recycle={} rename={}, expected {}/{}/{}.",
                                 providerName,
                                 permanentDeleteAdmitted ? 1 : 0,
                                 recycleAdmitted ? 1 : 0,
                                 renameAdmitted ? 1 : 0,
                                 expectedPermanentDeleteAdmission ? 1 : 0,
                                 expectedRecycleAdmission ? 1 : 0,
                                 expectedRenameAdmission ? 1 : 0));
                return false;
            }
            if (! requirePathIdentity(caps, providerName, expectedStable, expectedComparison, L'/', L"/", expectedCaseOnlyRename))
            {
                return false;
            }
            AppendLog(std::format(L"Provider conformance: {} declarations OK (copy={} move={} delete={} copyMoveMax={} deleteMax={}).",
                                  providerName,
                                  caps.copyOperation ? 1 : 0,
                                  caps.moveOperation ? 1 : 0,
                                  caps.deleteOperation ? 1 : 0,
                                  caps.copyMoveMax,
                                  caps.deleteMax));
            return true;
        };

        constexpr bool kExpectCopyYes = true;
        constexpr bool kExpectCopyNo  = false;
        // 2A: S3 advertises same-provider copy AND move so object-store folder merge is attempted.
        // R0f-S3: the route is providerWatchdog, so the host admits its mutations.
        if (! conformCloudProvider(kPluginIdS3,
                                   L"FileSystemS3 (S3)",
                                   &kExpectCopyYes,
                                   /*requireMoveMerge=*/true,
                                   /*expectNativeMove=*/true,
                                   /*expectDelete=*/true,
                                   /*expectRename=*/false,
                                   /*expectRecycle=*/false,
                                   /*expectMutationAdmission=*/true,
                                   /*requireRead=*/true,
                                   /*expectedStable=*/true,
                                   L"ordinalCaseSensitive",
                                   L"supported"))
        {
            return true;
        }
        if (! conformCloudProvider(kPluginIdS3Table,
                                   L"FileSystemS3 (S3 Table)",
                                   &kExpectCopyNo,
                                   /*requireMoveMerge=*/false,
                                   /*expectNativeMove=*/false,
                                   /*expectDelete=*/false,
                                   /*expectRename=*/false,
                                   /*expectRecycle=*/false,
                                   /*expectMutationAdmission=*/false,
                                   /*requireRead=*/true,
                                   /*expectedStable=*/true,
                                   L"ordinalCaseSensitive",
                                   L"notApplicable"))
        {
            return true;
        }
        // 2B: MS Drive advertises move (directory-onto-directory merge) but NOT same-provider copy —
        // the host bridges copy. This is the honest capability declaration 2B depends on.
        // R0f-Graph: the route is providerWatchdog, so the host admits its mutations; Rename is
        // advertised; Graph Delete stays Recycle (permanent Delete is honestly absent).
        if (! conformCloudProvider(kPluginIdOneDrivePersonal,
                                   L"FileSystemMicrosoftDrive (OneDrive)",
                                   &kExpectCopyNo,
                                   /*requireMoveMerge=*/true,
                                   /*expectNativeMove=*/true,
                                   /*expectDelete=*/false,
                                   /*expectRename=*/true,
                                   /*expectRecycle=*/true,
                                   /*expectMutationAdmission=*/true,
                                   /*requireRead=*/true,
                                   /*expectedStable=*/true,
                                   L"ordinalIgnoreCase",
                                   L"supported"))
        {
            return true;
        }
        // R0f-GDrive: Drive copies server-side, moves natively (parent change), deletes permanently
        // or trashes, renames; the route is providerWatchdog so the host admits every mutation.
        if (! conformCloudProvider(kPluginIdGoogleDrive,
                                   L"FileSystemGoogleDrive",
                                   &kExpectCopyYes,
                                   /*requireMoveMerge=*/true,
                                   /*expectNativeMove=*/true,
                                   /*expectDelete=*/true,
                                   /*expectRename=*/true,
                                   /*expectRecycle=*/true,
                                   /*expectMutationAdmission=*/true,
                                   /*requireRead=*/true,
                                   /*expectedStable=*/true,
                                   L"ordinalCaseSensitive",
                                   L"supported"))
        {
            return true;
        }
        const auto conformCurlConditionalMutationPolicy = [&](std::wstring_view pluginId, std::wstring_view providerName) noexcept -> bool
        {
            wil::com_ptr<IFileSystem> fs = loadProviderFs(pluginId);
            if (! fs)
            {
                Fail(std::format(L"{} must be loadable for the offline Curl bridge policy contract.", providerName));
                return false;
            }

            ProviderCapabilitySnapshot caps{};
            std::wstring reason;
            if (! TryReadProviderCapabilities(fs.get(), caps, reason))
            {
                Fail(std::format(L"{} GetCapabilities (offline) failed to parse: {}.", providerName, reason));
                return false;
            }
            // R0f-Curl: FTP/SFTP/SCP are full file-manager destinations (providerWatchdog route, same-provider
            // Copy/Move/Rename/Delete, Copy export and import). Move export/import stay denied until the
            // source can be deleted conditionally; Recycle never exists on these protocols.
            if (! caps.copyOperation || ! caps.moveOperation || ! caps.nativeMoveOperation ||
                ! caps.renameOperation || ! caps.deleteOperation || caps.recycleOperation || ! caps.read ||
                ! caps.write ||
                ! caps.exportCopyWildcard || caps.exportMoveWildcard || ! caps.importCopyWildcard || caps.importMoveWildcard ||
                ! requirePathIdentity(caps, providerName, true, L"ordinalCaseSensitive", L'/', L"/", L"notApplicable"))
            {
                Fail(std::format(L"{} must advertise same-provider Copy/Move/Rename/Delete plus Copy export and import, and deny Move export/import and Recycle (R0f-Curl).",
                                 providerName));
                return false;
            }

            if (! CanSameFileSystemOperation(fs, L"/", FILESYSTEM_COPY, pluginId) ||
                ! CanSameFileSystemOperation(fs, L"/", FILESYSTEM_MOVE, pluginId) ||
                ! CanSameFileSystemOperation(fs, L"/", FILESYSTEM_RENAME, pluginId) ||
                ! CanSameFileSystemOperation(fs, L"/", FILESYSTEM_DELETE, pluginId))
            {
                Fail(std::format(L"{} host planning must admit all four capability-v2 operations on the providerWatchdog route (R0f-Curl).", providerName));
                return false;
            }

            wil::com_ptr<IFileSystemAtomicWriter> atomicWriter;
            const HRESULT atomicHr = fs->QueryInterface(IID_PPV_ARGS(atomicWriter.addressof()));
            BOOL atomicSupported   = FALSE;
            if (FAILED(atomicHr) || ! atomicWriter ||
                FAILED(atomicWriter->SupportsAtomicWriterCommit(L"/offline-atomic-contract.bin", FILESYSTEM_FLAG_NONE, &atomicSupported)) ||
                atomicSupported != TRUE)
            {
                Fail(std::format(L"{} must advertise the atomic final writer (staged sibling + rename) so the host bridge can publish new names (R0f-Curl).",
                                 providerName));
                return false;
            }

            wil::com_ptr<IFileSystemIO> io;
            const HRESULT ioHr = fs->QueryInterface(IID_PPV_ARGS(io.addressof()));
            if (FAILED(ioHr) || ! io)
            {
                Fail(std::format(L"{} should expose IFileSystemIO for the direct writer-flag contract.", providerName));
                return false;
            }
            wil::com_ptr<IFileWriter> writer;
            const HRESULT contradictoryFlagsHr =
                io->CreateFileWriter(L"/offline-flag-contract.bin", FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY, writer.addressof());
            if (contradictoryFlagsHr != E_INVALIDARG || writer)
            {
                Fail(std::format(L"{} should reject ALLOW_REPLACE_READONLY without ALLOW_OVERWRITE as E_INVALIDARG before endpoint work.", providerName));
                return false;
            }

            const bool copyIn  = CanCrossFileSystemCopyMoveForSelfTest(state.fsLocal, state.tempRoot.native(), kPluginIdLocal, fs, L"/", pluginId, FILESYSTEM_COPY);
            const bool moveIn  = CanCrossFileSystemCopyMoveForSelfTest(state.fsLocal, state.tempRoot.native(), kPluginIdLocal, fs, L"/", pluginId, FILESYSTEM_MOVE);
            const bool copyOut = CanCrossFileSystemCopyMoveForSelfTest(fs, L"/", pluginId, state.fsLocal, state.tempRoot.native(), kPluginIdLocal, FILESYSTEM_COPY);
            const bool moveOut = CanCrossFileSystemCopyMoveForSelfTest(fs, L"/", pluginId, state.fsLocal, state.tempRoot.native(), kPluginIdLocal, FILESYSTEM_MOVE);
            // Cross-provider planning admits Move wherever it admits Copy; execution keeps the Move
            // source whenever the destination cannot prove publication (typed Move import/export
            // lists stay empty for these routes, see the caps checks above).
            if (! copyIn || ! moveIn || ! copyOut || ! moveOut)
            {
                Fail(std::format(
                    L"{} host planning must admit cross-provider Copy and Move in both directions on the providerWatchdog route (R0f-Curl): copyIn={} moveIn={} copyOut={} moveOut={}.",
                    providerName,
                    copyIn,
                    moveIn,
                    copyOut,
                    moveOut));
                return false;
            }

            AppendLog(std::format(L"Provider conformance: {} is a providerWatchdog full destination (Copy both ways, same-provider Copy/Move/Rename/Delete); Move export/import stay denied.",
                                  providerName));
            return true;
        };

        if (! conformCurlConditionalMutationPolicy(kPluginIdFtp, L"FileSystemCurl (FTP)") ||
            ! conformCurlConditionalMutationPolicy(kPluginIdSftp, L"FileSystemCurl (SFTP)") ||
            ! conformCurlConditionalMutationPolicy(kPluginIdScp, L"FileSystemCurl (SCP)"))
        {
            return true;
        }
        if (! conformCloudProvider(kPluginIdImap,
                                   L"FileSystemCurl (IMAP)",
                                   &kExpectCopyNo,
                                   /*requireMoveMerge=*/false,
                                   /*expectNativeMove=*/false,
                                   /*expectDelete=*/false,
                                   /*expectRename=*/false,
                                   /*expectRecycle=*/false,
                                   /*expectMutationAdmission=*/false,
                                   /*requireRead=*/true,
                                   /*expectedStable=*/true,
                                   L"ordinalCaseSensitive",
                                   L"notApplicable"))
        {
            return true;
        }

        wil::com_ptr<IFileSystemIO> dummyIo;
        const HRESULT hrDummyIo = state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()));
        if (FAILED(hrDummyIo) || ! dummyIo)
        {
            Fail(std::format(L"FileSystemDummy should expose IFileSystemIO for conformance reads (hr=0x{:08X}).", static_cast<unsigned long>(hrDummyIo)));
            return true;
        }

        const std::wstring root = std::format(L"/clearflow-provider-matrix-{}", GetTickCount64());
        const auto cleanup      = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(state.fsDummy->DeleteItem(
                root.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, nullptr, nullptr));
        });

        const std::filesystem::path rootPath(root);
        const std::filesystem::path copySrcFoo       = rootPath / L"copy-src" / L"Foo";
        const std::filesystem::path copySrcNested    = copySrcFoo / L"nested";
        const std::filesystem::path copyDstRoot      = rootPath / L"copy-dst";
        const std::filesystem::path copyDstFoo       = copyDstRoot / L"Foo";
        const std::filesystem::path copyDstKeep      = copyDstFoo / L"keep.txt";
        const std::filesystem::path copyDstNew       = copyDstFoo / L"new.txt";
        const std::filesystem::path copyDstNestedNew = copyDstFoo / L"nested" / L"child.txt";

        const auto ensureDummyDir = [&](const std::filesystem::path& path) noexcept -> bool
        { return EnsureDummyFolderExists(state.fsDummy.get(), path.generic_wstring()); };

        if (! ensureDummyDir(rootPath) || ! ensureDummyDir(rootPath / L"copy-src") || ! ensureDummyDir(copySrcFoo) || ! ensureDummyDir(copySrcNested) ||
            ! ensureDummyDir(copyDstRoot) || ! ensureDummyDir(copyDstFoo))
        {
            Fail(L"FileSystemDummy provider conformance failed to seed copy-merge directories.");
            return true;
        }

        if (! DummyWriteTextFile(state.fsDummy.get(), (copySrcFoo / L"new.txt").generic_wstring(), "new") ||
            ! DummyWriteTextFile(state.fsDummy.get(), (copySrcNested / L"child.txt").generic_wstring(), "child") ||
            ! DummyWriteTextFile(state.fsDummy.get(), copyDstKeep.generic_wstring(), "keep"))
        {
            Fail(L"FileSystemDummy provider conformance failed to seed copy-merge files.");
            return true;
        }

        FileOpsRecursiveProgressRecorder copyProgress{};
        const std::wstring copySrcFooText  = copySrcFoo.generic_wstring();
        const std::wstring copyDstRootText = copyDstRoot.generic_wstring();
        const wchar_t* copySources[]       = {copySrcFooText.c_str()};
        const HRESULT copyHr               = state.fsDummy->CopyItems(
            copySources, 1, copyDstRootText.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), nullptr, &copyProgress, nullptr);
        if (FAILED(copyHr))
        {
            Fail(std::format(L"FileSystemDummy directory copy merge failed: 0x{:08X}.", static_cast<unsigned long>(copyHr)));
            return true;
        }

        std::string text;
        if (! ReadFileTextFsIo(dummyIo, copyDstKeep, text) || text != "keep" || ! ReadFileTextFsIo(dummyIo, copyDstNew, text) || text != "new" ||
            ! ReadFileTextFsIo(dummyIo, copyDstNestedNew, text) || text != "child")
        {
            Fail(L"FileSystemDummy directory copy merge failed byte-for-byte destination checks.");
            return true;
        }

        if (copyProgress.progressCount == 0 || copyProgress.completedCount == 0 || copyProgress.streamCount == 0)
        {
            Fail(std::format(L"FileSystemDummy copy progress contract incomplete (progress={} completed={} streams={}).",
                             copyProgress.progressCount,
                             copyProgress.completedCount,
                             copyProgress.streamCount));
            return true;
        }

        const std::filesystem::path moveSrcFoo    = rootPath / L"move-src" / L"Foo";
        const std::filesystem::path moveDstRoot   = rootPath / L"move-dst";
        const std::filesystem::path moveDstFoo    = moveDstRoot / L"Foo";
        const std::filesystem::path moveDstKeep   = moveDstFoo / L"keep.txt";
        const std::filesystem::path moveDstNew    = moveDstFoo / L"moved.txt";
        const std::filesystem::path moveSrcNew    = moveSrcFoo / L"moved.txt";
        const std::filesystem::path moveSrcNested = moveSrcFoo / L"nested";
        const std::filesystem::path moveDstNested = moveDstFoo / L"nested" / L"moved-child.txt";
        if (! ensureDummyDir(rootPath / L"move-src") || ! ensureDummyDir(moveSrcFoo) || ! ensureDummyDir(moveSrcNested) || ! ensureDummyDir(moveDstRoot) ||
            ! ensureDummyDir(moveDstFoo))
        {
            Fail(L"FileSystemDummy provider conformance failed to seed move-merge directories.");
            return true;
        }

        if (! DummyWriteTextFile(state.fsDummy.get(), moveSrcNew.generic_wstring(), "moved") ||
            ! DummyWriteTextFile(state.fsDummy.get(), (moveSrcNested / L"moved-child.txt").generic_wstring(), "moved-child") ||
            ! DummyWriteTextFile(state.fsDummy.get(), moveDstKeep.generic_wstring(), "keep-move"))
        {
            Fail(L"FileSystemDummy provider conformance failed to seed move-merge files.");
            return true;
        }

        FileOpsRecursiveProgressRecorder moveProgress{};
        const std::wstring moveSrcFooText  = moveSrcFoo.generic_wstring();
        const std::wstring moveDstRootText = moveDstRoot.generic_wstring();
        const wchar_t* moveSources[]       = {moveSrcFooText.c_str()};
        const HRESULT moveHr               = state.fsDummy->MoveItems(
            moveSources, 1, moveDstRootText.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), nullptr, &moveProgress, nullptr);
        if (FAILED(moveHr))
        {
            Fail(std::format(L"FileSystemDummy directory move merge failed: 0x{:08X}.", static_cast<unsigned long>(moveHr)));
            return true;
        }

        unsigned long attributes = 0;
        if (PathExistsFsIo(dummyIo, moveSrcFoo, &attributes))
        {
            Fail(L"FileSystemDummy directory move merge left the source directory behind.");
            return true;
        }

        if (! ReadFileTextFsIo(dummyIo, moveDstKeep, text) || text != "keep-move" || ! ReadFileTextFsIo(dummyIo, moveDstNew, text) || text != "moved" ||
            ! ReadFileTextFsIo(dummyIo, moveDstNested, text) || text != "moved-child")
        {
            Fail(L"FileSystemDummy directory move merge failed byte-for-byte destination checks.");
            return true;
        }

        if (moveProgress.progressCount == 0 || moveProgress.completedCount == 0 || moveProgress.streamCount == 0)
        {
            Fail(std::format(L"FileSystemDummy move progress contract incomplete (progress={} completed={} streams={}).",
                             moveProgress.progressCount,
                             moveProgress.completedCount,
                             moveProgress.streamCount));
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.ProviderCapabilityMatrix.CopyProgressStreams", L"dummy", 0u, copyProgress.streamCount, 0u, S_OK);
        Debug::Perf::Emit(L"FileOps.SelfTest.ProviderCapabilityMatrix.MoveProgressStreams", L"dummy", 0u, moveProgress.streamCount, 0u, S_OK);
        NextStep(state, SelfTestState::Step::R4A02_SameDestinationWarns);
        return false;
    }

    return false;
}
#endif
#if defined(FILEOPS_SELFTEST_INCLUDE_PHASE5)

case SelfTestState::Step::R4A19_DiscoveryMeasurementFacts:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"R4A19_DiscoveryMeasurementFacts timed out.");
        return true;
    }

    constexpr unsigned int kDirectoryCount = 16u;
    constexpr unsigned int kFilesPerDirectory = 8u;
    constexpr size_t kFileBytes = 256u * 1024u;
    constexpr unsigned int kDeleteFileCount = 1'024u;
    constexpr unsigned int kDeleteDirectoryCount = 128u;
    constexpr unsigned int kDeleteFilesPerDirectory = kDeleteFileCount / kDeleteDirectoryCount;
    constexpr size_t kDeleteFileBytes = 1u * 1024u;
    constexpr uint64_t kExpectedBytes =
        static_cast<uint64_t>(kDirectoryCount) * kFilesPerDirectory * kFileBytes;
    constexpr uint64_t kExpectedDeleteBytes = static_cast<uint64_t>(kDeleteFileCount) * kDeleteFileBytes;
    const std::filesystem::path sourceRoot = state.tempRoot / L"r4-a19-measurement-src";
    const std::filesystem::path destinationRoot = state.tempRoot / L"r4-a19-measurement-dst";
    const std::filesystem::path copiedRoot = destinationRoot / sourceRoot.filename();
    const std::filesystem::path deleteRoot = state.tempRoot / L"r4-a19-small-delete";
    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(sourceRoot) || ! RecreateEmptyDirectory(destinationRoot))
        {
            Fail(L"Failed to reset the R4-A19 discovery measurement fixture.");
            return true;
        }

        for (unsigned int directoryIndex = 0u; directoryIndex < kDirectoryCount; ++directoryIndex)
        {
            const std::filesystem::path directory = sourceRoot / std::format(L"dir-{:02}", directoryIndex);
            if (! RecreateEmptyDirectory(directory))
            {
                Fail(L"Failed to seed an R4-A19 discovery measurement directory.");
                return true;
            }
            for (unsigned int fileIndex = 0u; fileIndex < kFilesPerDirectory; ++fileIndex)
            {
                if (! WriteTestFile(directory / std::format(L"file-{:02}.bin", fileIndex), kFileBytes))
                {
                    Fail(L"Failed to seed an R4-A19 discovery measurement file.");
                    return true;
                }
            }
        }

        const std::string config =
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json";
        if (! SetPluginConfiguration(state.infoLocal.get(), config))
        {
            Fail(L"Failed to configure Local concurrency for R4-A19 discovery measurement.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceRoot},
                                                 destinationRoot,
                                                 flags,
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start the R4-A19 discovery measurement copy.");
            return true;
        }

        state.r4A19ScenarioStartTick = GetTickCount64();
        state.stepState = 1;
        return false;
    }

    const auto completionIt = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completionIt == state.completedTasks.end())
    {
        return false;
    }

    const CompletedTaskInfo& completion = completionIt->second;
    const ULONGLONG durationMs = completion.completionTick >= state.r4A19ScenarioStartTick
                                    ? completion.completionTick - state.r4A19ScenarioStartTick
                                    : 0u;
    if (state.stepState == 1 &&
        (FAILED(completion.hr) || ! completion.discoveryClosed || ! completion.firstMutationBeforeDiscoveryClosed ||
         completion.discoveredTotalBytes != kExpectedBytes || completion.discoveredFiles != kDirectoryCount * kFilesPerDirectory ||
         completion.discoveryMaxQueueDepth > 256u || durationMs >= 30'000u ||
         CountFilesRecursive(copiedRoot) != kDirectoryCount * kFilesPerDirectory))
    {
        Fail(std::format(L"R4-A19 Local Copy control failed: hr=0x{:08X} closed={} mutationBeforeClose={} discoveredBytes={}/{} "
                         L"discoveredFiles={}/{} copiedFiles={} queueMax={} durationMs={}.",
                         static_cast<unsigned long>(completion.hr),
                         completion.discoveryClosed ? 1 : 0,
                         completion.firstMutationBeforeDiscoveryClosed ? 1 : 0,
                         completion.discoveredTotalBytes,
                         kExpectedBytes,
                         completion.discoveredFiles,
                         kDirectoryCount * kFilesPerDirectory,
                         CountFilesRecursive(copiedRoot),
                         completion.discoveryMaxQueueDepth,
                         durationMs));
        return true;
    }

    constexpr uint64_t kMissing = (std::numeric_limits<uint64_t>::max)();
    if (completion.discoveryFirstMutationUs == kMissing || completion.discoveryBytesCompletedWhileOpen == kMissing ||
        completion.discoveryMutationsCompletedWhileOpen == kMissing)
    {
        Fail(L"R4-A19 RED: completion evidence is missing first-mutation latency or open-traversal byte/mutation facts.");
        return true;
    }

    if (state.stepState == 1 &&
        (completion.discoveryFirstMutationUs == 0u || completion.discoveryFirstMutationUs > 1'000'000u ||
        completion.discoveryBytesCompletedWhileOpen == 0u || completion.discoveryMutationsCompletedWhileOpen == 0u)
       )
    {
        Fail(std::format(L"R4-A19 discovery facts are invalid: firstMutationUs={} bytesWhileOpen={} mutationsWhileOpen={}.",
                         completion.discoveryFirstMutationUs,
                         completion.discoveryBytesCompletedWhileOpen,
                         completion.discoveryMutationsCompletedWhileOpen));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.R4A19.DiscoveryMeasurementFacts",
                      state.stepState == 1 ? L"scenario=local-copy-128x256KiB" : L"scenario=local-delete-1024-small-files",
                      durationMs * 1000u,
                      completion.discoveryFirstMutationUs,
                      completion.discoveryMutationsCompletedWhileOpen,
                      completion.hr);

    if (state.stepState == 1)
    {
        if (! RecreateEmptyDirectory(deleteRoot))
        {
            Fail(L"Failed to reset the R4-A19 small-file Delete fixture.");
            return true;
        }
        for (unsigned int directoryIndex = 0u; directoryIndex < kDeleteDirectoryCount; ++directoryIndex)
        {
            const std::filesystem::path directory = deleteRoot / std::format(L"dir-{:03}", directoryIndex);
            if (! RecreateEmptyDirectory(directory))
            {
                Fail(L"Failed to seed an R4-A19 small-file Delete directory.");
                return true;
            }
            for (unsigned int fileIndex = 0u; fileIndex < kDeleteFilesPerDirectory; ++fileIndex)
            {
                if (! WriteTestFile(directory / std::format(L"delete-{:02}.bin", fileIndex), kDeleteFileBytes))
                {
                    Fail(L"Failed to seed the R4-A19 small-file Delete fixture.");
                    return true;
                }
            }
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {deleteRoot},
                                                 {},
                                                 flags,
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start the R4-A19 small-file Delete task.");
            return true;
        }
        state.r4A19ScenarioStartTick = GetTickCount64();
        state.stepState = 2;
        return false;
    }

    if (FAILED(completion.hr) || ! completion.discoveryClosed || ! completion.firstMutationBeforeDiscoveryClosed ||
        std::filesystem::exists(deleteRoot) || completion.discoveredTotalBytes != kExpectedDeleteBytes ||
        completion.discoveredFiles != kDeleteFileCount || completion.discoveryMaxQueueDepth > 256u ||
        completion.discoveryFirstMutationUs == 0u || completion.discoveryFirstMutationUs > 1'000'000u ||
        completion.discoveryBytesCompletedWhileOpen != kExpectedDeleteBytes ||
        completion.discoveryMutationsCompletedWhileOpen != 1u || completion.progressCallbackCount != 0u || durationMs >= 30'000u)
    {
        Fail(std::format(L"R4-A19 Local Delete control failed: hr=0x{:08X} closed={} mutationBeforeClose={} sourceExists={} "
                         L"discoveredBytes={}/{} discoveredFiles={}/{} firstMutationUs={} bytesWhileOpen={} mutationsWhileOpen={} "
                         L"progressCallbacks={} queueMax={} durationMs={}.",
                         static_cast<unsigned long>(completion.hr),
                         completion.discoveryClosed ? 1 : 0,
                         completion.firstMutationBeforeDiscoveryClosed ? 1 : 0,
                         std::filesystem::exists(deleteRoot) ? 1 : 0,
                         completion.discoveredTotalBytes,
                         kExpectedDeleteBytes,
                         completion.discoveredFiles,
                         kDeleteFileCount,
                         completion.discoveryFirstMutationUs,
                         completion.discoveryBytesCompletedWhileOpen,
                         completion.discoveryMutationsCompletedWhileOpen,
                         completion.progressCallbackCount,
                         completion.discoveryMaxQueueDepth,
                         durationMs));
        return true;
    }

    if (! state.localConfigOriginal.empty() && ! SetPluginConfiguration(state.infoLocal.get(), state.localConfigOriginal))
    {
        Fail(L"Failed to restore Local configuration after R4-A19 Local workloads.");
        return true;
    }
    NextStep(state, SelfTestState::Step::R4A19_DiscoveryProviderControls);
    return false;
}

case SelfTestState::Step::R4A19_DiscoveryProviderControls:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"R4A19_DiscoveryProviderControls timed out.");
        return true;
    }

    constexpr std::wstring_view kDummySourceRoot = L"/r4-a19-dummy-source";
    constexpr std::wstring_view kDummyDestinationRoot = L"/r4-a19-dummy-destination";
    constexpr unsigned int kDummyDirectoryCount = 8u;
    constexpr unsigned int kDummyFilesPerDirectory = 8u;
    constexpr size_t kDummyFileBytes = 8u * 1024u;
    constexpr uint64_t kExpectedDummyBytes =
        static_cast<uint64_t>(kDummyDirectoryCount) * kDummyFilesPerDirectory * kDummyFileBytes;
    constexpr std::wstring_view kMtpSourceRoot =
        L"/Fake Phone/Internal Storage/DCIM/Camera/r4-a19-source";
    constexpr std::wstring_view kMtpPhoto =
        L"/Fake Phone/Internal Storage/DCIM/Camera/r4-a19-source/file-00.bin";
    constexpr size_t kExpectedMtpFiles = 8u;
    constexpr size_t kMtpFileBytes = 64u * 1024u;
    constexpr uint64_t kExpectedMtpBytes = kExpectedMtpFiles * kMtpFileBytes;
    const std::filesystem::path mtpDestinationRoot = state.tempRoot / L"r4-a19-mtp-destination";
    const std::filesystem::path mtpCancelDestinationRoot = state.tempRoot / L"r4-a19-mtp-cancel-destination";
    const FileSystemFlags recursiveFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    const auto bridgeResourcesAreBounded = [](const CompletedTaskInfo& info) noexcept
    {
        return info.discoveryMaxQueueDepth <= 256u &&
               info.bridgeAdmissionMaxQueueDepth <= 256u &&
               info.bridgeTraversalMaxRetainedEntries <= 4'096u &&
               info.bridgeTraversalMaxQueuedPathBytes <= 16u * 1024u * 1024u &&
               info.bridgeTraversalMaxMetadataBytes <= 8u * 1024u * 1024u &&
               info.bridgeTraversalLimitHitCount == 0u;
    };

    if (state.stepState == 0u)
    {
        const std::string seedConfig =
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json";
        if (! SetPluginConfiguration(state.infoDummy.get(), seedConfig))
        {
            Fail(L"R4-A19 failed to reset Dummy before the delayed sequential control.");
            return true;
        }

        const FileSystemFlags cleanupFlags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        static_cast<void>(state.fsDummy->DeleteItem(std::wstring(kDummySourceRoot).c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        static_cast<void>(state.fsDummy->DeleteItem(std::wstring(kDummyDestinationRoot).c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), kDummySourceRoot) ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), kDummyDestinationRoot))
        {
            Fail(L"R4-A19 failed to create the Dummy control roots.");
            return true;
        }

        const std::string payload(kDummyFileBytes, static_cast<char>(0x5a));
        for (unsigned int directoryIndex = 0u; directoryIndex < kDummyDirectoryCount; ++directoryIndex)
        {
            const std::wstring directory =
                std::wstring(kDummySourceRoot) + std::format(L"/dir-{:02}", directoryIndex);
            if (! EnsureDummyFolderExists(state.fsDummy.get(), directory))
            {
                Fail(L"R4-A19 failed to create a Dummy control directory.");
                return true;
            }
            for (unsigned int fileIndex = 0u; fileIndex < kDummyFilesPerDirectory; ++fileIndex)
            {
                const std::wstring file = directory + std::format(L"/file-{:02}.bin", fileIndex);
                if (! DummyWriteTextFile(state.fsDummy.get(), file, payload))
                {
                    Fail(L"R4-A19 failed to seed a Dummy control file.");
                    return true;
                }
            }
        }

        const std::string delayedConfig =
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":5,"streamChunkLatencyMs":2,"virtualSpeedLimit":"1048576"})json";
        if (! SetPluginConfiguration(state.infoDummy.get(), delayedConfig))
        {
            Fail(L"R4-A19 failed to enable the bounded Dummy latency/speed control.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(kDummySourceRoot)},
                                                 std::filesystem::path(kDummyDestinationRoot),
                                                 recursiveFlags,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"R4-A19 failed to start the delayed Dummy Copy control.");
            return true;
        }
        state.r4A19ScenarioStartTick = nowTick;
        state.stepState = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        const CompletedTaskInfo& info = completed->second;
        const uint64_t durationUs =
            info.completionTick >= state.r4A19ScenarioStartTick
                ? static_cast<uint64_t>(info.completionTick - state.r4A19ScenarioStartTick) * 1000u
                : 0u;
        wil::com_ptr<IFileSystemIO> dummyIo;
        if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
        {
            Fail(L"R4-A19 delayed Dummy control could not query IFileSystemIO.");
            return true;
        }

        const std::string expected(kDummyFileBytes, static_cast<char>(0x5a));
        bool bytesMatch = true;
        for (unsigned int directoryIndex = 0u; directoryIndex < kDummyDirectoryCount && bytesMatch; ++directoryIndex)
        {
            for (unsigned int fileIndex = 0u; fileIndex < kDummyFilesPerDirectory; ++fileIndex)
            {
                const std::filesystem::path file =
                    std::filesystem::path(kDummyDestinationRoot) /
                    std::filesystem::path(kDummySourceRoot).filename() /
                    std::format(L"dir-{:02}", directoryIndex) /
                    std::format(L"file-{:02}.bin", fileIndex);
                std::string actual;
                if (! ReadFileTextFsIo(dummyIo, file, actual) || actual != expected)
                {
                    bytesMatch = false;
                    break;
                }
            }
        }

        if (FAILED(info.hr) || ! bytesMatch || ! info.discoveryClosed ||
            ! info.firstMutationBeforeDiscoveryClosed ||
            info.discoveryFirstMutationUs == 0u || info.discoveryFirstMutationUs > 1'000'000u ||
            info.discoveryBytesCompletedWhileOpen == 0u ||
            info.discoveryMutationsCompletedWhileOpen == 0u ||
            info.discoveryStarvationCount != 0u || info.discoveryMaxQueueDepth > 256u ||
            durationUs >= 30'000'000u)
        {
            Fail(std::format(L"R4-A19 Dummy control failed: hr=0x{:08X} bytesMatch={} closed={} beforeClose={} "
                             L"discoveredBytes={} expectedPayloadBytes={} files={} expectedPayloadFiles={} firstMutationUs={} bytesWhileOpen={} "
                             L"mutationsWhileOpen={} starvation={} queue={} durationUs={}.",
                             static_cast<unsigned long>(info.hr),
                             bytesMatch ? 1 : 0,
                             info.discoveryClosed ? 1 : 0,
                             info.firstMutationBeforeDiscoveryClosed ? 1 : 0,
                             info.discoveredTotalBytes,
                             kExpectedDummyBytes,
                             info.discoveredFiles,
                             kDummyDirectoryCount * kDummyFilesPerDirectory,
                             info.discoveryFirstMutationUs,
                             info.discoveryBytesCompletedWhileOpen,
                             info.discoveryMutationsCompletedWhileOpen,
                             info.discoveryStarvationCount,
                             info.discoveryMaxQueueDepth,
                             durationUs));
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.R4A19.ProviderControl",
                          L"provider=dummy;latencyMs=5;streamChunkLatencyMs=2;virtualSpeedLimit=1048576;mode=single-root-per-item",
                          durationUs,
                          info.discoveryFirstMutationUs,
                          info.discoveryMutationsCompletedWhileOpen,
                          info.hr);
        AppendLog(L"R4-A19 provider control: Dummy complete; loading fake MTP.");

        using CreateMtpForSelfTestFunc =
            HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const char*, void**) noexcept;
        wil::unique_hmodule module;
        FARPROC createAddress = nullptr;
        HRESULT hr = SelfTest::LoadMtpPluginSelfTestExport(
            "RedSalamanderMtpCreateForSelfTest", module, createAddress);
        if (FAILED(hr) || createAddress == nullptr)
        {
            Fail(std::format(L"R4-A19 fake-MTP factory is unavailable: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        AppendLog(L"R4-A19 provider control: fake-MTP export loaded.");
#pragma warning(push)
#pragma warning(disable : 4191) // The named self-test export fixes the typed factory ABI.
        const auto createMtp = reinterpret_cast<CreateMtpForSelfTestFunc>(createAddress);
#pragma warning(pop)

        FactoryOptions factoryOptions{};
        factoryOptions.debugLevel = DEBUG_LEVEL_NONE;
        wil::com_ptr<IFileSystem> mtp;
        hr = createMtp(__uuidof(IFileSystem),
                       &factoryOptions,
                       GetHostServices(),
                       R"json({"operationDelayMs":25})json",
                       mtp.put_void());
        if (FAILED(hr) || ! mtp)
        {
            Fail(std::format(L"R4-A19 failed to create the serialized fake-MTP instance: 0x{:08X}.",
                             static_cast<unsigned long>(hr)));
            return true;
        }
        AppendLog(L"R4-A19 provider control: fake-MTP object created.");

        wil::com_ptr<IInformations> mtpInfo;
        wil::com_ptr<IFileSystemInitialize> mtpInitialize;
        wil::com_ptr<IFileSystemIO> mtpIo;
        hr = mtp->QueryInterface(IID_PPV_ARGS(mtpInfo.addressof()));
        AppendLog(std::format(L"R4-A19 provider control: fake-MTP IInformations QI returned 0x{:08X}.", static_cast<unsigned long>(hr)));
        if (SUCCEEDED(hr))
        {
            hr = mtpInfo->SetConfiguration(
                R"json({"readOnly":false,"commandTimeoutMs":5000})json");
            AppendLog(std::format(L"R4-A19 provider control: fake-MTP configuration returned 0x{:08X}.", static_cast<unsigned long>(hr)));
        }
        if (SUCCEEDED(hr))
        {
            hr = mtp->QueryInterface(IID_PPV_ARGS(mtpInitialize.addressof()));
            AppendLog(std::format(L"R4-A19 provider control: fake-MTP initialize QI returned 0x{:08X}.", static_cast<unsigned long>(hr)));
        }
        if (SUCCEEDED(hr))
        {
            hr = mtpInitialize->Initialize(L"/", nullptr);
            AppendLog(std::format(L"R4-A19 provider control: fake-MTP Initialize returned 0x{:08X}.", static_cast<unsigned long>(hr)));
        }
        if (SUCCEEDED(hr))
        {
            hr = mtp->QueryInterface(IID_PPV_ARGS(mtpIo.addressof()));
            AppendLog(std::format(L"R4-A19 provider control: fake-MTP I/O QI returned 0x{:08X}.", static_cast<unsigned long>(hr)));
        }
        wil::com_ptr<IFileSystemDirectoryOperations> mtpDirectoryOperations;
        if (SUCCEEDED(hr))
        {
            hr = mtp->QueryInterface(IID_PPV_ARGS(mtpDirectoryOperations.addressof()));
            AppendLog(std::format(L"R4-A19 provider control: fake-MTP directory-operations QI returned 0x{:08X}.",
                                  static_cast<unsigned long>(hr)));
        }
        if (FAILED(hr) || ! mtpIo || ! mtpDirectoryOperations)
        {
            Fail(std::format(L"R4-A19 failed to initialize fake MTP: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }

        const std::wstring mtpSourceRoot(kMtpSourceRoot);
        hr = mtpDirectoryOperations->CreateDirectory(mtpSourceRoot.c_str());
        if (FAILED(hr))
        {
            Fail(std::format(L"R4-A19 failed to create the ordinary-name fake-MTP source directory: 0x{:08X}.",
                             static_cast<unsigned long>(hr)));
            return true;
        }
        const std::string mtpPayload(kMtpFileBytes, static_cast<char>(0x4d));
        for (size_t fileIndex = 0u; fileIndex < kExpectedMtpFiles; ++fileIndex)
        {
            const std::filesystem::path filePath =
                std::filesystem::path(kMtpSourceRoot) / std::format(L"file-{:02}.bin", fileIndex);
            if (! WriteFileTextFsIo(mtpIo, filePath, mtpPayload))
            {
                Fail(std::format(L"R4-A19 failed to seed ordinary fake-MTP file {:02}.", fileIndex));
                return true;
            }
        }

        state.r4A19MtpModule = std::move(module);
        state.r4A19Mtp = std::move(mtp);
        state.r4A19MtpIo = std::move(mtpIo);
        if (! RecreateEmptyDirectory(mtpDestinationRoot))
        {
            Fail(L"R4-A19 failed to reset the fake-MTP destination.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.r4A19Mtp,
                                                 {std::filesystem::path(kMtpSourceRoot)},
                                                 mtpDestinationRoot,
                                                 recursiveFlags,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        AppendLog(std::format(L"R4-A19 provider control: fake-MTP task admission returned taskId={}.", state.taskA.value_or(0u)));
        if (! state.taskA.has_value())
        {
            Fail(L"R4-A19 failed to start the fake-MTP Copy control.");
            return true;
        }
        state.r4A19ScenarioStartTick = GetTickCount64();
        state.stepState = 2u;
        return false;
    }

    if (state.stepState == 2u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        const CompletedTaskInfo& info = completed->second;
        const uint64_t durationUs =
            info.completionTick >= state.r4A19ScenarioStartTick
                ? static_cast<uint64_t>(info.completionTick - state.r4A19ScenarioStartTick) * 1000u
                : 0u;
        const std::filesystem::path copiedCamera = mtpDestinationRoot / L"r4-a19-source";
        wil::com_ptr<IFileSystemIO> localIo;
        const HRESULT localIoHr = state.fsLocal
            ? state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))
            : E_POINTER;
        const std::string expectedMtpPayload(kMtpFileBytes, static_cast<char>(0x4d));
        bool mtpBytesMatch = SUCCEEDED(localIoHr) && localIo;
        for (size_t fileIndex = 0u; fileIndex < kExpectedMtpFiles && mtpBytesMatch; ++fileIndex)
        {
            std::string actual;
            const std::filesystem::path filePath = copiedCamera / std::format(L"file-{:02}.bin", fileIndex);
            mtpBytesMatch = ReadFileTextFsIo(localIo, filePath, actual) && actual == expectedMtpPayload;
        }
        if (FAILED(info.hr) || CountFilesRecursive(copiedCamera) != kExpectedMtpFiles ||
            ! mtpBytesMatch || info.discoveredTotalBytes != kExpectedMtpBytes ||
            ! info.discoveryClosed || ! info.firstMutationBeforeDiscoveryClosed ||
            info.discoveredFiles != kExpectedMtpFiles ||
            info.discoveryFirstMutationUs == 0u || info.discoveryFirstMutationUs > 1'000'000u ||
            info.discoveryBytesCompletedWhileOpen == 0u ||
            info.discoveryMutationsCompletedWhileOpen == 0u ||
            info.discoveryStarvationCount != 0u || ! bridgeResourcesAreBounded(info) ||
            durationUs >= 30'000'000u)
        {
            Fail(std::format(L"R4-A19 fake-MTP control failed: hr=0x{:08X} copiedFiles={}/{} bytesMatch={} discoveredBytes={}/{} closed={} "
                             L"beforeClose={} discoveredFiles={} firstMutationUs={} bytesWhileOpen={} "
                             L"mutationsWhileOpen={} earlyStarts={} starvation={} readyQueue={} bridgeQueue={} "
                             L"retained={}/{}/{} limits={} durationUs={}.",
                             static_cast<unsigned long>(info.hr),
                             CountFilesRecursive(copiedCamera),
                             kExpectedMtpFiles,
                             mtpBytesMatch ? 1 : 0,
                             info.discoveredTotalBytes,
                             kExpectedMtpBytes,
                             info.discoveryClosed ? 1 : 0,
                             info.firstMutationBeforeDiscoveryClosed ? 1 : 0,
                             info.discoveredFiles,
                             info.discoveryFirstMutationUs,
                             info.discoveryBytesCompletedWhileOpen,
                             info.discoveryMutationsCompletedWhileOpen,
                             info.bridgeEarlyFileStartCount,
                             info.discoveryStarvationCount,
                             info.discoveryMaxQueueDepth,
                             info.bridgeAdmissionMaxQueueDepth,
                             info.bridgeTraversalMaxRetainedEntries,
                             info.bridgeTraversalMaxQueuedPathBytes,
                             info.bridgeTraversalMaxMetadataBytes,
                             info.bridgeTraversalLimitHitCount,
                             durationUs));
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.R4A19.ProviderControl",
                          L"provider=fake-mtp;operationDelayMs=25;session=serialized",
                          durationUs,
                          info.discoveryFirstMutationUs,
                          info.bridgeTraversalMaxRetainedEntries,
                          info.hr);
        if (! RecreateEmptyDirectory(mtpCancelDestinationRoot))
        {
            Fail(L"R4-A19 failed to reset the fake-MTP cancellation destination.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.r4A19Mtp,
                                                 {std::filesystem::path(kMtpSourceRoot)},
                                                 mtpCancelDestinationRoot,
                                                 recursiveFlags,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            Fail(L"R4-A19 failed to start the fake-MTP cancellation control.");
            return true;
        }
        state.stepState = 3u;
        return false;
    }

    if (state.stepState == 3u)
    {
        auto* task = state.fileOps && state.taskA.has_value()
            ? state.fileOps->FindTask(state.taskA.value())
            : nullptr;
        if (task != nullptr)
        {
            uint64_t completedBytes = 0u;
            {
                std::scoped_lock lock(task->_progressMutex);
                completedBytes = task->_progressCompletedBytes;
            }
            if (! task->_discoveryClosed.load(std::memory_order_acquire) && completedBytes > 0u)
            {
                task->RequestCancel();
                state.markerTick = nowTick;
                state.stepState = 4u;
                return false;
            }
        }

        if (state.taskA.has_value() && state.completedTasks.contains(state.taskA.value()))
        {
            Fail(L"R4-A19 fake-MTP cancellation control completed before discovery and mutation were simultaneously live.");
            return true;
        }
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }

    const CompletedTaskInfo& info = completed->second;
    const uint64_t cancelUs =
        info.completionTick >= state.markerTick
            ? static_cast<uint64_t>(info.completionTick - state.markerTick) * 1000u
            : 0u;
    const HRESULT canceledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    const HRESULT operationAbortedHr = HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED);
    const char* properties = nullptr;
    const HRESULT propertiesHr = state.r4A19MtpIo
        ? state.r4A19MtpIo->GetItemProperties(std::wstring(kMtpPhoto).c_str(), &properties)
        : E_POINTER;
    const std::optional<uint64_t> maxConcurrentBackendCalls =
        SUCCEEDED(propertiesHr) && properties != nullptr
            ? SelfTest::ExtractJsonUInt(properties, "maxConcurrentBackendCalls")
            : std::nullopt;
    if ((info.hr != canceledHr && info.hr != operationAbortedHr && info.hr != E_ABORT) || cancelUs > 500'000u ||
        ! info.discoveryClosed || ! maxConcurrentBackendCalls.has_value() ||
        maxConcurrentBackendCalls.value() != 1u || ! bridgeResourcesAreBounded(info))
    {
        Fail(std::format(L"R4-A19 fake-MTP cancel/serialization failed: hr=0x{:08X} cancelUs={} closed={} "
                         L"propsHr=0x{:08X} maxBackend={} discoveryQueue={} bridgeQueue={} retained={}/{}/{} limits={} "
                         L"checks(hr,cancel,closed,props,backend,bounded)={}{}{}{}{}{}.",
                         static_cast<unsigned long>(info.hr),
                         cancelUs,
                         info.discoveryClosed ? 1 : 0,
                         static_cast<unsigned long>(propertiesHr),
                         maxConcurrentBackendCalls.value_or((std::numeric_limits<uint64_t>::max)()),
                         info.discoveryMaxQueueDepth,
                         info.bridgeAdmissionMaxQueueDepth,
                         info.bridgeTraversalMaxRetainedEntries,
                         info.bridgeTraversalMaxQueuedPathBytes,
                         info.bridgeTraversalMaxMetadataBytes,
                         info.bridgeTraversalLimitHitCount,
                         (info.hr == canceledHr || info.hr == operationAbortedHr || info.hr == E_ABORT) ? 1 : 0,
                         cancelUs <= 500'000u ? 1 : 0,
                         info.discoveryClosed ? 1 : 0,
                         maxConcurrentBackendCalls.has_value() ? 1 : 0,
                         maxConcurrentBackendCalls.value_or(0u) == 1u ? 1 : 0,
                         bridgeResourcesAreBounded(info) ? 1 : 0));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.R4A19.ProviderCancel",
                      L"provider=fake-mtp;operationDelayMs=25;cancel-while-discovery-and-mutation-live",
                      cancelUs,
                      maxConcurrentBackendCalls.value(),
                      info.discoveryCallbackUs + info.discoveryLockWaitUs,
                      info.hr);
    NextStep(state, SelfTestState::Step::R4A19_DiscoveryIndependentVolumes);
    return false;
}

case SelfTestState::Step::R4A19_DiscoveryIndependentVolumes:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"R4A19_DiscoveryIndependentVolumes timed out.");
        return true;
    }

    constexpr unsigned int kFileCount = 4u;
    constexpr size_t kFileBytes = 4u * 1024u * 1024u;
    constexpr uint64_t kTaskBytes = static_cast<uint64_t>(kFileCount) * kFileBytes;
    constexpr uint64_t kSpeedLimitBytesPerSecond = 16u * 1024u * 1024u;
    const std::filesystem::path cBase = state.tempRoot / L"r4-a19-independent-c";
    const std::filesystem::path cSource = cBase / L"source-c";
    const std::filesystem::path cDestination = cBase / L"destination-c";

    if (state.stepState == 0u)
    {
        std::wstring alternateDetail;
        const std::optional<std::filesystem::path> alternate =
            TryCreateAlternateWritableVolumeSelfTestRoot(state.tempRoot, alternateDetail);
        if (! alternate.has_value())
        {
            Fail(std::format(L"R4-A19 independent-volume control is blocked: {}.", alternateDetail));
            return true;
        }
        state.fileOpsAlternateVolumeRoot = alternate.value();

        const std::filesystem::path dBase = state.fileOpsAlternateVolumeRoot / L"r4-a19-independent-d";
        const std::filesystem::path dSource = dBase / L"source-d";
        const std::filesystem::path dDestination = dBase / L"destination-d";
        if (! RecreateEmptyDirectory(cSource) || ! RecreateEmptyDirectory(cDestination) ||
            ! RecreateEmptyDirectory(dSource) || ! RecreateEmptyDirectory(dDestination))
        {
            Fail(L"R4-A19 failed to reset the independent-volume fixture.");
            return true;
        }

        for (unsigned int index = 0u; index < kFileCount; ++index)
        {
            if (! WriteFilledTestFile(cSource / std::format(L"payload-{:02}.bin", index), kFileBytes, 0x43) ||
                ! WriteFilledTestFile(dSource / std::format(L"payload-{:02}.bin", index), kFileBytes, 0x44))
            {
                Fail(L"R4-A19 failed to seed the independent-volume payloads.");
                return true;
            }
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {cSource},
                                                 cDestination,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 kSpeedLimitBytesPerSecond);
        if (! state.taskA.has_value())
        {
            Fail(L"R4-A19 failed to start the isolated primary-volume control.");
            return true;
        }
        state.r4A19ScenarioStartTick = nowTick;
        state.stepState = 1u;
        return false;
    }

    const std::filesystem::path dBase = state.fileOpsAlternateVolumeRoot / L"r4-a19-independent-d";
    const std::filesystem::path dSource = dBase / L"source-d";
    const std::filesystem::path dDestination = dBase / L"destination-d";
    const auto verifyTree = [&](const std::filesystem::path& source,
                                const std::filesystem::path& destinationRoot) noexcept
    {
        const std::filesystem::path copied = destinationRoot / source.filename();
        if (CountFilesRecursive(copied) != kFileCount)
        {
            return false;
        }
        for (unsigned int index = 0u; index < kFileCount; ++index)
        {
            const std::filesystem::path leaf = std::format(L"payload-{:02}.bin", index);
            if (! FilesEqualBytes(source / leaf, copied / leaf))
            {
                return false;
            }
        }
        return true;
    };

    if (state.stepState == 1u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        state.r4A19VolumeCIsolatedUs =
            completed->second.completionTick >= state.r4A19ScenarioStartTick
                ? static_cast<uint64_t>(completed->second.completionTick - state.r4A19ScenarioStartTick) * 1000u
                : 0u;
        if (FAILED(completed->second.hr) || ! verifyTree(cSource, cDestination) ||
            completed->second.discoveredTotalBytes != kTaskBytes ||
            state.r4A19VolumeCIsolatedUs >= 30'000'000u)
        {
            Fail(std::format(L"R4-A19 primary isolated control failed: hr=0x{:08X} bytes={}/{} durationUs={}.",
                             static_cast<unsigned long>(completed->second.hr),
                             completed->second.discoveredTotalBytes,
                             kTaskBytes,
                             state.r4A19VolumeCIsolatedUs));
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {dSource},
                                                 dDestination,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 kSpeedLimitBytesPerSecond);
        if (! state.taskA.has_value())
        {
            Fail(L"R4-A19 failed to start the isolated alternate-volume control.");
            return true;
        }
        state.r4A19ScenarioStartTick = nowTick;
        state.stepState = 2u;
        return false;
    }

    if (state.stepState == 2u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        state.r4A19VolumeDIsolatedUs =
            completed->second.completionTick >= state.r4A19ScenarioStartTick
                ? static_cast<uint64_t>(completed->second.completionTick - state.r4A19ScenarioStartTick) * 1000u
                : 0u;
        if (FAILED(completed->second.hr) || ! verifyTree(dSource, dDestination) ||
            completed->second.discoveredTotalBytes != kTaskBytes ||
            state.r4A19VolumeDIsolatedUs >= 30'000'000u)
        {
            Fail(std::format(L"R4-A19 alternate isolated control failed: hr=0x{:08X} bytes={}/{} durationUs={}.",
                             static_cast<unsigned long>(completed->second.hr),
                             completed->second.discoveredTotalBytes,
                             kTaskBytes,
                             state.r4A19VolumeDIsolatedUs));
            return true;
        }

        if (! RecreateEmptyDirectory(cDestination) || ! RecreateEmptyDirectory(dDestination))
        {
            Fail(L"R4-A19 failed to reset destinations before concurrent volume work.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {cSource},
                                                 cDestination,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 kSpeedLimitBytesPerSecond);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {dSource},
                                                 dDestination,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 kSpeedLimitBytesPerSecond);
        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            Fail(L"R4-A19 failed to start both concurrent independent-volume controls.");
            return true;
        }
        state.r4A19IndependentProgressOverlap = false;
        state.r4A19ScenarioStartTick = nowTick;
        state.stepState = 3u;
        return false;
    }

    auto* taskA = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    auto* taskB = state.fileOps && state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
    if (taskA != nullptr && taskB != nullptr)
    {
        uint64_t completedA = 0u;
        uint64_t completedB = 0u;
        {
            std::scoped_lock lock(taskA->_progressMutex);
            completedA = taskA->_progressCompletedBytes;
        }
        {
            std::scoped_lock lock(taskB->_progressMutex);
            completedB = taskB->_progressCompletedBytes;
        }
        state.r4A19IndependentProgressOverlap =
            state.r4A19IndependentProgressOverlap ||
            (taskA->HasStarted() && taskB->HasStarted() && completedA > 0u && completedB > 0u);
    }

    const auto completedA = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    const auto completedB = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
    if (completedA == state.completedTasks.end() || completedB == state.completedTasks.end())
    {
        return false;
    }

    const uint64_t concurrentUs =
        (std::max)(completedA->second.completionTick, completedB->second.completionTick) >= state.r4A19ScenarioStartTick
            ? static_cast<uint64_t>((std::max)(completedA->second.completionTick, completedB->second.completionTick) -
                                    state.r4A19ScenarioStartTick) * 1000u
            : 0u;
    const uint64_t slowerIsolatedUs = (std::max)(state.r4A19VolumeCIsolatedUs, state.r4A19VolumeDIsolatedUs);
    const uint64_t concurrentGateUs = (slowerIsolatedUs * 135u + 99u) / 100u;
    if (FAILED(completedA->second.hr) || FAILED(completedB->second.hr) ||
        ! state.r4A19IndependentProgressOverlap || ! verifyTree(cSource, cDestination) ||
        ! verifyTree(dSource, dDestination) || concurrentUs >= 30'000'000u ||
        concurrentUs > concurrentGateUs)
    {
        Fail(std::format(L"R4-A19 independent-volume control failed: hrA=0x{:08X} hrB=0x{:08X} overlap={} "
                         L"bytesA={}/{} bytesB={}/{} isolatedC={} isolatedD={} concurrent={} gate={}.",
                         static_cast<unsigned long>(completedA->second.hr),
                         static_cast<unsigned long>(completedB->second.hr),
                         state.r4A19IndependentProgressOverlap ? 1 : 0,
                         completedA->second.discoveredTotalBytes,
                         kTaskBytes,
                         completedB->second.discoveredTotalBytes,
                         kTaskBytes,
                         state.r4A19VolumeCIsolatedUs,
                         state.r4A19VolumeDIsolatedUs,
                         concurrentUs,
                         concurrentGateUs));
        return true;
    }

    const std::wstring volumeDetail = std::format(
        L"primary={};alternate={};bytesPerTask={};speedLimit={}",
        state.tempRoot.root_path().wstring(),
        state.fileOpsAlternateVolumeRoot.root_path().wstring(),
        kTaskBytes,
        kSpeedLimitBytesPerSecond);
    Debug::Perf::Emit(L"FileOps.SelfTest.R4A19.IndependentVolumes",
                      volumeDetail,
                      concurrentUs,
                      slowerIsolatedUs,
                      state.r4A19IndependentProgressOverlap ? 1u : 0u,
                      S_OK);
    Debug::Perf::EmitValue(L"FileOps.SelfTest.R4A19.IndependentVolumeCIsolatedUs",
                           state.r4A19VolumeCIsolatedUs,
                           S_OK);
    Debug::Perf::EmitValue(L"FileOps.SelfTest.R4A19.IndependentVolumeDIsolatedUs",
                           state.r4A19VolumeDIsolatedUs,
                           S_OK);

    const std::filesystem::path alternateRoot = std::move(state.fileOpsAlternateVolumeRoot);
    static_cast<void>(SelfTest::RemoveAll(alternateRoot));
    PruneEmptyAlternateVolumeSandboxParents(alternateRoot);
    NextStep(state, SelfTestState::Step::R0fSmb_BlockedSynchronousCallCancelReturns);
    return false;
}

case SelfTestState::Step::R0fSmb_BlockedSynchronousCallCancelReturns:
{
    // R0f-SMB containment witness: a bounded provider whose operation call wedges inside a
    // synchronous Win32 read (the state a dead SMB share leaves a worker in) must return with
    // ERROR_OPERATION_ABORTED after Cancel, and the task must end as canceled within the bound.
    static wil::com_ptr<IFileSystem> blockedFileSystem;
    static UncontainedNeverReturningFileSystem* blockedObserver = nullptr;
    static ULONGLONG cancelRequestedTick                       = 0u;
    const ULONGLONG nowTick                                    = GetTickCount64();
    if (HasTimedOut(state, nowTick, 30'000ull))
    {
        if (blockedObserver != nullptr)
        {
            blockedObserver->ReleaseBlockedRead();
        }
        Fail(L"R0fSmb_BlockedSynchronousCallCancelReturns timed out.");
        return true;
    }

    if (state.stepState == 0u)
    {
        const char* localCapabilitiesJson = nullptr;
        if (FAILED(state.fsLocal->GetPathCapabilities(L"/", FILESYSTEM_COPY, &localCapabilitiesJson)) || localCapabilitiesJson == nullptr)
        {
            Fail(L"R0f-SMB witness could not read the Local capability document.");
            return true;
        }
        auto owner      = std::make_unique<UncontainedNeverReturningFileSystem>(std::string(localCapabilitiesJson), true, true, true);
        blockedObserver = owner.get();
        blockedFileSystem.attach(owner.release());
        if (! blockedObserver->BlockedReadReady())
        {
            Fail(L"R0f-SMB witness could not create its anonymous pipe.");
            return true;
        }

        uint64_t taskId          = 0u;
        const HRESULT admissionHr = state.fileOps->AdmitOperation(FILESYSTEM_COPY,
                                                                  FolderWindow::Pane::Left,
                                                                  FolderWindow::Pane::Right,
                                                                  blockedFileSystem,
                                                                  {std::filesystem::path(L"/blocked-read/source.bin")},
                                                                  std::filesystem::path(L"/blocked-read/destination"),
                                                                  FILESYSTEM_FLAG_NONE,
                                                                  false,
                                                                  0u,
                                                                  FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                  false,
                                                                  nullptr,
                                                                  &taskId,
                                                                  {},
                                                                  {},
                                                                  L"selftest/bounded-blocked-read",
                                                                  L"bounded-blocked-read");
        if (FAILED(admissionHr) || taskId == 0u)
        {
            Fail(std::format(L"R0f-SMB witness: a bounded route must be admitted (hr=0x{:08X}).", static_cast<unsigned long>(admissionHr)));
            return true;
        }
        state.taskA         = taskId;
        cancelRequestedTick = 0u;
        state.stepState     = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        // Wait for the worker to enter the wedged read, then cancel and time the return.
        if (blockedObserver->ProviderOperationCallCount() == 0u)
        {
            return false;
        }
        Sleep(50u); // let the read reach the kernel
        auto* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (task == nullptr)
        {
            blockedObserver->ReleaseBlockedRead();
            Fail(L"R0f-SMB witness lost its admitted task before cancel.");
            return true;
        }
        cancelRequestedTick = GetTickCount64();
        task->RequestCancel();
        const bool returned      = blockedObserver->WaitForBlockedReadReturn(5'000u);
        const ULONGLONG returnMs = GetTickCount64() - cancelRequestedTick;
        const HRESULT readStatus = blockedObserver->BlockedReadStatus();
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fSmb.WedgedCallReturnMs", returnMs, readStatus);
        if (! returned || readStatus != HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED))
        {
            blockedObserver->ReleaseBlockedRead();
            Fail(std::format(L"R0f-SMB witness: the wedged synchronous call must return ERROR_OPERATION_ABORTED after Cancel (returned={}, status=0x{:08X}, ms={}).",
                             returned,
                             static_cast<unsigned long>(readStatus),
                             returnMs));
            return true;
        }
        state.stepState = 2u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    const ULONGLONG terminalMs = completed->second.completionTick >= cancelRequestedTick ? completed->second.completionTick - cancelRequestedTick : 0u;
    Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fSmb.CanceledTaskTerminalMs", terminalMs, completed->second.hr);
    if (completed->second.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) && completed->second.hr != E_ABORT)
    {
        std::wstring items;
        for (size_t index = 0u; index < completed->second.sourceItemResults.size(); ++index)
        {
            const auto& sealed = completed->second.sourceItemResults[index];
            const auto& raw    = index < completed->second.sourceItemStatuses.size() ? completed->second.sourceItemStatuses[index] : std::nullopt;
            items += std::format(L" item{}[raw=0x{:08X} sealed={} status=0x{:08X} completion={} publication={} source={}]",
                                 index,
                                 static_cast<unsigned long>(raw.value_or(S_OK)),
                                 sealed.has_value() ? 1 : 0,
                                 static_cast<unsigned long>(sealed.has_value() ? sealed->status : S_OK),
                                 sealed.has_value() ? static_cast<unsigned int>(sealed->completion) : 99u,
                                 sealed.has_value() ? static_cast<unsigned int>(sealed->publication) : 99u,
                                 sealed.has_value() ? static_cast<unsigned int>(sealed->sourceDisposition) : 99u);
        }
        Fail(std::format(L"R0f-SMB witness: the task must end as canceled, not 0x{:08X}.{}", static_cast<unsigned long>(completed->second.hr), items));
        return true;
    }
    blockedObserver = nullptr;
    blockedFileSystem.reset();
    NextStep(state, SelfTestState::Step::R0fSmb_LoopbackReadWriteCreateDelete);
    return false;
}

case SelfTestState::Step::R0fSmb_LoopbackReadWriteCreateDelete:
{
    // R0f-SMB functional matrix on the administrative-share alias of the sandbox
    // (\\localhost\<drive>$\...): Copy to the share, Move back from it, CreateDirectory and
    // Rename on it through the provider, Delete on it through a task, byte-exact verification.
    // Environment-gated: skipped with a log line when the loopback share is not reachable.
    static std::filesystem::path uncRoot;
    static ULONGLONG phaseStartTick = 0u;
    const ULONGLONG nowTick         = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"R0fSmb_LoopbackReadWriteCreateDelete timed out.");
        return true;
    }

    constexpr unsigned int kFileCount = 3u;
    constexpr size_t kFileBytes       = 256u * 1024u;
    const std::filesystem::path localBase   = state.tempRoot / L"r0f-smb-local";
    const std::filesystem::path localSource = localBase / L"source";
    const std::filesystem::path localMoved  = localBase / L"moved-back";
    const auto verifyTree = [&](const std::filesystem::path& copied) noexcept
    {
        if (CountFilesRecursive(copied) != kFileCount)
        {
            return false;
        }
        for (unsigned int index = 0u; index < kFileCount; ++index)
        {
            const std::filesystem::path leaf = std::format(L"payload-{:02}.bin", index);
            if (! FilesEqualBytes(localSource / leaf, copied / leaf))
            {
                return false;
            }
        }
        return true;
    };

    if (state.stepState == 0u)
    {
        const std::filesystem::path uncSandbox = LoopbackShareAlias(state.tempRoot);
        if (! LoopbackShareReachable(state.tempRoot))
        {
            AppendLog(std::format(L"R0f-SMB loopback: {} is not reachable; environment-gated case skipped.", uncSandbox.wstring()));
            NextStep(state, SelfTestState::Step::R0fCurl_FakeFtpReadWriteCreateDelete);
            return false;
        }
        uncRoot = uncSandbox / L"r0f-smb-share";

        FileOperations::QualifiedEndpoint uncEndpoint{};
        if (! FileOperations::TryQualifyEndpointForSelfTest(state.fsLocal, uncRoot.native(), FILESYSTEM_COPY, L"builtin/file-system", L"host/default", uncEndpoint) ||
            uncEndpoint.profileId != L"local-win32-smb" ||
            uncEndpoint.cancellationRouteClass != FileOperations::CancellationRouteClass::Bounded)
        {
            Fail(L"R0f-SMB loopback: the UNC alias must qualify as the bounded local-win32-smb route.");
            return true;
        }

        const std::filesystem::path copyDestination = uncRoot / L"copy-dst";
        if (! RecreateEmptyDirectory(localSource) || ! RecreateEmptyDirectory(localMoved) || ! RecreateEmptyDirectory(copyDestination))
        {
            Fail(L"R0f-SMB loopback: failed to reset the local and share fixtures.");
            return true;
        }
        for (unsigned int index = 0u; index < kFileCount; ++index)
        {
            if (! WriteFilledTestFile(localSource / std::format(L"payload-{:02}.bin", index), kFileBytes, static_cast<unsigned char>(0x51u + index)))
            {
                Fail(L"R0f-SMB loopback: failed to seed the local payloads.");
                return true;
            }
        }

        const std::filesystem::path created = uncRoot / L"created";
        wil::com_ptr<IFileSystemDirectoryOperations> shareDirectoryOps;
        HRESULT createHr = state.fsLocal->QueryInterface(__uuidof(IFileSystemDirectoryOperations), shareDirectoryOps.put_void());
        if (SUCCEEDED(createHr))
        {
            createHr = shareDirectoryOps->CreateDirectory(created.c_str());
        }
        if (FAILED(createHr) || ! std::filesystem::is_directory(created))
        {
            Fail(std::format(L"R0f-SMB loopback: provider CreateDirectory on the share failed (hr=0x{:08X}).", static_cast<unsigned long>(createHr)));
            return true;
        }
        const std::filesystem::path renameSource = uncRoot / L"rename-me.bin";
        const std::filesystem::path renamed      = uncRoot / L"renamed.bin";
        if (! WriteFilledTestFile(renameSource, 4096u, 0x5Au))
        {
            Fail(L"R0f-SMB loopback: failed to seed the rename payload on the share.");
            return true;
        }
        const HRESULT renameHr = state.fsLocal->RenameItem(renameSource.c_str(), renamed.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        if (FAILED(renameHr) || std::filesystem::exists(renameSource) || ! std::filesystem::is_regular_file(renamed))
        {
            Fail(std::format(L"R0f-SMB loopback: provider Rename on the share failed (hr=0x{:08X}).", static_cast<unsigned long>(renameHr)));
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSource},
                                                 copyDestination,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"R0f-SMB loopback: failed to start the Copy to the share.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        const std::filesystem::path copied = uncRoot / L"copy-dst" / L"source";
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fSmb.LoopbackCopyMs", nowTick - phaseStartTick, completed->second.hr);
        if (FAILED(completed->second.hr) || ! verifyTree(copied))
        {
            std::wstring itemDetail;
            for (const std::optional<FileOperations::FileOperationItemResult>& item : completed->second.sourceItemResults)
            {
                if (item.has_value())
                {
                    itemDetail += std::format(L" [item status=0x{:08X} completion={} publication={} destination={}]",
                                              static_cast<unsigned long>(item->status),
                                              static_cast<unsigned int>(item->completion),
                                              static_cast<unsigned int>(item->publication),
                                              item->finalDestinationPath);
                }
            }
            Fail(std::format(L"R0f-SMB loopback: Copy to the share failed or is not byte-exact (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             itemDetail));
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsLocal,
                                                 {copied},
                                                 localMoved,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"R0f-SMB loopback: failed to start the Move back from the share.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 2u;
        return false;
    }

    if (state.stepState == 2u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fSmb.LoopbackMoveBackMs", nowTick - phaseStartTick, completed->second.hr);
        if (FAILED(completed->second.hr) || std::filesystem::exists(uncRoot / L"copy-dst" / L"source") || ! verifyTree(localMoved / L"source"))
        {
            Fail(std::format(L"R0f-SMB loopback: Move back from the share failed, left the source, or is not byte-exact (hr=0x{:08X}).",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Right,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {uncRoot / L"created", uncRoot / L"renamed.bin", uncRoot / L"copy-dst"},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"R0f-SMB loopback: failed to start the Delete on the share.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 3u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fSmb.LoopbackDeleteMs", nowTick - phaseStartTick, completed->second.hr);
    if (FAILED(completed->second.hr) || std::filesystem::exists(uncRoot / L"created") || std::filesystem::exists(uncRoot / L"renamed.bin") ||
        std::filesystem::exists(uncRoot / L"copy-dst"))
    {
        Fail(std::format(L"R0f-SMB loopback: Delete on the share failed or left objects behind (hr=0x{:08X}).", static_cast<unsigned long>(completed->second.hr)));
        return true;
    }
    static_cast<void>(SelfTest::RemoveAll(uncRoot));
    static_cast<void>(SelfTest::RemoveAll(localBase));
    NextStep(state, SelfTestState::Step::R0fCurl_FakeFtpReadWriteCreateDelete);
    return false;
}

case SelfTestState::Step::R0fCurl_FakeFtpReadWriteCreateDelete:
case SelfTestState::Step::BR5_CurlHostDirectoryMove:
case SelfTestState::Step::BR5_CurlHostDirectoryMoveRefused:
case SelfTestState::Step::C0_CurlHostCommittedDeleteResponseLost:
case SelfTestState::Step::C0_CurlHostPartialTreeFailure:
{
    // R0f-Curl matrix through the host against the plugin's deterministic loopback FTP fixture:
    // cross-provider Copy local -> FTP, Copy FTP -> local (byte-exact), provider CreateDirectory and
    // Rename on FTP, Delete on FTP through a task. The route must qualify as providerWatchdog.
    static wil::unique_hmodule curlModule;
    static void* fakeFtp            = nullptr;
    static unsigned int fakeFtpPort = 0u;
    static wil::com_ptr<IFileSystem> ftp;
    static wil::com_ptr<IFileSystemIO> ftpIo;
    static std::wstring ftpRoot;
    static ULONGLONG phaseStartTick   = 0u;
    const bool refusedDirectoryMove = state.step == SelfTestState::Step::BR5_CurlHostDirectoryMoveRefused;
    const bool directoryMoveWitness = refusedDirectoryMove || state.step == SelfTestState::Step::BR5_CurlHostDirectoryMove;
    const bool partialDeleteWitness   = state.step == SelfTestState::Step::C0_CurlHostPartialTreeFailure;
    const bool committedDeleteWitness = partialDeleteWitness || state.step == SelfTestState::Step::C0_CurlHostCommittedDeleteResponseLost;
    using StopFakeFtpFn               = void(__stdcall*)(void*) noexcept;
    const auto stopFakeFtp            = []() noexcept
    {
        if (fakeFtp != nullptr && curlModule)
        {
            if (const FARPROC stopAddress = GetProcAddress(curlModule.get(), "RedSalamanderCurlStopFakeFtpForSelfTest"))
            {
#pragma warning(push)
#pragma warning(disable : 4191)
                reinterpret_cast<StopFakeFtpFn>(stopAddress)(fakeFtp);
#pragma warning(pop)
            }
        }
        fakeFtp = nullptr;
        ftpIo.reset();
        ftp.reset();
    };
    const auto taskDiagnostics = [&]() noexcept -> std::wstring
    {
        std::wstring text;
        if (! state.taskA.has_value())
        {
            return text;
        }
        std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> entries;
        state.fileOps->CollectDiagnostics(entries);
        unsigned int shown = 0u;
        for (const auto& entry : entries)
        {
            if (entry.taskId != state.taskA.value() || shown >= 6u)
            {
                continue;
            }
            text += std::format(L" [{} 0x{:08X} {}]", entry.category, static_cast<unsigned long>(entry.status), entry.message);
            ++shown;
        }
        return text;
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        const std::wstring diagnostics = taskDiagnostics();
        stopFakeFtp();
        Fail(std::format(L"R0fCurl_FakeFtpReadWriteCreateDelete timed out at step {}.{}", state.stepState, diagnostics));
        return true;
    }

    constexpr unsigned int kFileCount       = 3u;
    constexpr size_t kFileBytes             = 64u * 1024u;
    const std::filesystem::path localBase   = state.tempRoot / L"r0f-curl-local";
    const std::filesystem::path localSource = localBase / L"source";
    const std::filesystem::path localBack   = localBase / L"copied-back";
    const auto ftpPath                      = [&](std::wstring_view leaf) { return ftpRoot + std::wstring(leaf); };
    const auto ftpExists                    = [&](std::wstring_view leaf) noexcept
    {
        unsigned long attributes = 0u;
        return ftpIo && SUCCEEDED(ftpIo->GetAttributes(ftpPath(leaf).c_str(), &attributes));
    };
    // C10: FTP names no object identity, so the Permanent Delete card must say the deletion is by name.
    constexpr std::wstring_view kC10Leaf = L"c10-by-name.bin";
    if (state.stepState == 6u)
    {
        FolderWindow::FileOperationState::Task* const task = state.fileOps->FindTask(state.taskA.value());
        const auto prompt                                  = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            if (state.completedTasks.contains(state.taskA.value()))
            {
                HostSetAutoAcceptPrompts(true);
                stopFakeFtp();
                curlModule.reset();
                Fail(L"C10-Curl: the Permanent Delete ran without asking on the card.");
                return true;
            }
            return false;
        }
        const std::wstring byName = LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_PERMANENT_DELETE_BY_NAME);
        const bool saysByName     = prompt->bucket == FolderWindow::FileOperationState::Task::ConflictBucket::PermanentDeleteConfirmation && ! byName.empty() &&
                                prompt->consentDetail.find(byName) != std::wstring::npos;
        task->SubmitConflictDecision(FolderWindow::FileOperationState::Task::ConflictAction::Cancel, false);
        HostSetAutoAcceptPrompts(true);
        if (! saysByName)
        {
            stopFakeFtp();
            curlModule.reset();
            Fail(std::format(L"C10-Curl: the card must say that this location deletes by name (bucket={}, detail='{}').",
                             static_cast<unsigned int>(prompt->bucket),
                             prompt->consentDetail));
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 7u;
        return false;
    }
    const auto committedDelete = [&](BOOL arm, unsigned int& requests, unsigned int& commits, BOOL& stateCorrect) noexcept -> HRESULT
    {
        using WitnessFn = HRESULT(__stdcall*)(void*, BOOL, unsigned int*, unsigned int*, BOOL*) noexcept;
        const FARPROC address =
            curlModule ? GetProcAddress(curlModule.get(),
                                        partialDeleteWitness ? "RedSalamanderCurlPartialDeleteForSelfTest" : "RedSalamanderCurlCommittedDeleteForSelfTest")
                       : nullptr;
        if (address == nullptr)
        {
            return E_NOINTERFACE;
        }
#pragma warning(push)
#pragma warning(disable : 4191) // Test-only exact C ABI; DLL stays pinned through fixture stop.
        return reinterpret_cast<WitnessFn>(address)(fakeFtp, arm, &requests, &commits, &stateCorrect);
#pragma warning(pop)
    };

    const auto directoryMoveSnapshot = [&](BOOL arm, unsigned int& renameFromCount, unsigned int& renameToCount, BOOL& stateCorrect) noexcept -> HRESULT
    {
        using WitnessFn = HRESULT(__stdcall*)(void*, BOOL, BOOL, unsigned int*, unsigned int*, BOOL*) noexcept;
        const FARPROC address = curlModule ? GetProcAddress(curlModule.get(), "RedSalamanderCurlDirectoryMoveForSelfTest") : nullptr;
        if (! address)
        {
            return E_NOINTERFACE;
        }
#pragma warning(push)
#pragma warning(disable : 4191) // Test-only exact C ABI; the fixture's module stays pinned until stop.
        return reinterpret_cast<WitnessFn>(address)(fakeFtp, arm, refusedDirectoryMove ? TRUE : FALSE, &renameFromCount, &renameToCount, &stateCorrect);
#pragma warning(pop)
    };

    if (directoryMoveWitness && state.stepState == 11u)
    {
        using Task = FolderWindow::FileOperationState::Task;
        if (Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (TryGetConflictPromptCopy(task).has_value())
            {
                // No retry authority is available for the attempted Curl rename.
                task->SubmitConflictDecision(Task::ConflictAction::Cancel, false);
            }
        }
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        unsigned int renameFromCount = 0u;
        unsigned int renameToCount = 0u;
        BOOL stateCorrect = FALSE;
        const HRESULT witnessHr = directoryMoveSnapshot(FALSE, renameFromCount, renameToCount, stateCorrect);
        const auto& result = completed->second;
        const auto* item = result.sourceItemResults.size() == 1u && result.sourceItemResults.front().has_value()
                               ? &result.sourceItemResults.front().value() : nullptr;
        const bool truthful = item && (refusedDirectoryMove
            ? result.hr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) && item->completion == FileOperations::ItemCompletion::Indeterminate &&
              item->publication == FileOperations::PublicationState::Unknown && item->sourceDisposition == FileOperations::SourceDisposition::Unknown
            : result.hr == S_OK && item->completion == FileOperations::ItemCompletion::Completed &&
              item->publication == FileOperations::PublicationState::Published && item->sourceDisposition == FileOperations::SourceDisposition::Removed);
        const bool passed = SUCCEEDED(witnessHr) && renameFromCount == 1u && renameToCount == 1u && stateCorrect && truthful && result.conflictPromptCount == 0u;
        Debug::Perf::Emit(L"FileOps.SelfTest.CurlHostDirectoryMove",
                          refusedDirectoryMove ? L"RNTO=550;source-tree-intact;no-relay-or-cleanup;host=unknown" : L"RNTO=250;subtree-relocated;no-relay-or-cleanup",
                          (nowTick - phaseStartTick) * 1000u, renameFromCount, renameToCount, passed ? S_OK : E_FAIL);
        const std::wstring detail = passed ? std::wstring{} : std::format(
            L"hr=0x{:08X}, rnfr={}, rnto={}, state={}, truthful={}, prompts={}.{}",
            static_cast<unsigned long>(result.hr), renameFromCount, renameToCount, stateCorrect, truthful, result.conflictPromptCount, taskDiagnostics());
        HostSetAutoAcceptPrompts(true);
        stopFakeFtp();
        curlModule.reset();
        state.taskA.reset();
        if (! passed)
        {
            Fail(L"BR-5 host directory Move violated its rename/result contract: " + detail);
            return true;
        }
        NextStep(state, refusedDirectoryMove ? SelfTestState::Step::R0fS3_FakeS3ReadWriteCreateDelete : SelfTestState::Step::BR5_CurlHostDirectoryMoveRefused);
        return false;
    }

    if (committedDeleteWitness && state.stepState == 10u)
    {
        using Task = FolderWindow::FileOperationState::Task;
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (TryGetConflictPromptCopy(task).has_value())
            {
                task->SubmitConflictDecision(Task::ConflictAction::Cancel, false); // Drain an unexpected prompt before failing the test.
            }
        }
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        unsigned int requests    = 0u;
        unsigned int commits     = 0u;
        BOOL stateCorrect        = FALSE;
        const HRESULT snapshotHr = committedDelete(FALSE, requests, commits, stateCorrect);
        const auto& result       = completed->second;
        const bool truthful      = result.sourceItemResults.size() == 1u && result.sourceItemResults[0].has_value() &&
                                   result.sourceItemResults[0].value().completion == FileOperations::ItemCompletion::Indeterminate &&
                                   result.sourceItemResults[0].value().publication == FileOperations::PublicationState::NotAttempted &&
                                   result.sourceItemResults[0].value().sourceDisposition == FileOperations::SourceDisposition::Unknown;
        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        const auto summary    = std::ranges::find_if(summaries, [&](const auto& item) { return item.taskId == state.taskA.value(); });
        const bool humanTruth = summary != summaries.end() && summary->indeterminateItemCount == 1u &&
                                summary->resultSummary == LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN);
        const bool passed     = SUCCEEDED(snapshotHr) && requests == (partialDeleteWitness ? 2u : 1u) && commits == 1u && stateCorrect == TRUE && truthful &&
                                humanTruth && result.hr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) && result.conflictPromptCount == 0u;
        Debug::Perf::Emit(partialDeleteWitness ? L"FileOps.SelfTest.C0Curl.PartialTreeFailure" : L"FileOps.SelfTest.C0Curl.CommittedDeleteResponseLost",
                          L"native-delete-host-result",
                          (nowTick - phaseStartTick) * 1000u,
                          requests,
                          commits,
                          passed ? S_OK : E_FAIL);
        const std::wstring detail = passed ? std::wstring{}
                                           : std::format(L"hr=0x{:08X}, requests={}, commits={}, state={}, truthful={}, humanTruth={}, prompts={}.{}",
                                                         static_cast<unsigned long>(result.hr),
                                                         requests,
                                                         commits,
                                                         stateCorrect,
                                                         truthful,
                                                         humanTruth,
                                                         result.conflictPromptCount,
                                                         taskDiagnostics());
        stopFakeFtp();
        curlModule.reset();
        if (! passed)
        {
            Fail(L"C0-Curl: committed Delete must finish uncertain without replay or conflict actions: " + detail);
            return true;
        }
        NextStep(state, partialDeleteWitness ? SelfTestState::Step::C0_CurlNativeDeleteLateListing : SelfTestState::Step::C0_CurlPartialTreeFailure);
        return false;
    }

    if (state.stepState == 0u)
    {
        using StartFakeFtpFn = HRESULT(__stdcall*)(unsigned int*, void**) noexcept;
        using CreateFn       = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);
        FARPROC startAddress = nullptr;
        HRESULT hr           = SelfTest::LoadPluginSelfTestExport(kPluginIdFtp, "RedSalamanderCurlStartFakeFtpForSelfTest", curlModule, startAddress);
        if (FAILED(hr) || startAddress == nullptr)
        {
            Fail(std::format(L"R0f-Curl: the fake FTP fixture export is unavailable: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        const FARPROC createAddress = GetProcAddress(curlModule.get(), "RedSalamanderCreate");
        if (createAddress == nullptr)
        {
            Fail(L"R0f-Curl: the Curl plugin factory export is unavailable.");
            return true;
        }
#pragma warning(push)
#pragma warning(disable : 4191)
        hr = reinterpret_cast<StartFakeFtpFn>(startAddress)(&fakeFtpPort, &fakeFtp);
#pragma warning(pop)
        if (FAILED(hr) || fakeFtp == nullptr || fakeFtpPort == 0u)
        {
            Fail(std::format(L"R0f-Curl: the fake FTP fixture did not start: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        FactoryOptions factoryOptions{};
        factoryOptions.debugLevel = DEBUG_LEVEL_NONE;
#pragma warning(push)
#pragma warning(disable : 4191)
        hr = reinterpret_cast<CreateFn>(createAddress)(__uuidof(IFileSystem), &factoryOptions, GetHostServices(), kPluginIdFtp.data(), ftp.put_void());
#pragma warning(pop)
        if (FAILED(hr) || ! ftp)
        {
            stopFakeFtp();
            Fail(std::format(L"R0f-Curl: failed to create the FTP provider instance: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        wil::com_ptr<IInformations> information;
        const std::string fixtureConfiguration =
            std::format(R"({{"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":2,"deleteMaxConcurrency":{},"ftpUseEpsv":true}})",
                        partialDeleteWitness ? 1u : 2u);
        if (FAILED(ftp->QueryInterface(__uuidof(IInformations), information.put_void())) || ! information ||
            FAILED(information->SetConfiguration(fixtureConfiguration.c_str())) || FAILED(ftp->QueryInterface(__uuidof(IFileSystemIO), ftpIo.put_void())) ||
            ! ftpIo)
        {
            stopFakeFtp();
            Fail(L"R0f-Curl: the FTP provider did not accept the fixture configuration or lacks IFileSystemIO.");
            return true;
        }
        ftpRoot = std::format(L"//anonymous@127.0.0.1:{}/", fakeFtpPort);

        FileOperations::QualifiedEndpoint ftpEndpoint{};
        if (! FileOperations::TryQualifyEndpointForSelfTest(ftp, ftpRoot, FILESYSTEM_COPY, kPluginIdFtp, L"host/default", ftpEndpoint) ||
            ftpEndpoint.cancellationRouteClass != FileOperations::CancellationRouteClass::ProviderWatchdog || ftpEndpoint.providerWatchdogTimeoutMs == 0u)
        {
            stopFakeFtp();
            Fail(L"R0f-Curl: the FTP route must qualify as providerWatchdog with a nonzero provider-owned bound.");
            return true;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fCurl.ProviderWatchdogTimeoutMs", ftpEndpoint.providerWatchdogTimeoutMs, S_OK);

        if (directoryMoveWitness)
        {
            unsigned int renameFromCount = 0u;
            unsigned int renameToCount = 0u;
            BOOL stateCorrect = FALSE;
            if (FAILED(directoryMoveSnapshot(TRUE, renameFromCount, renameToCount, stateCorrect)))
            {
                stopFakeFtp();
                curlModule.reset();
                Fail(L"BR-5 could not seed the host directory Move fixture.");
                return true;
            }
            HostSetAutoAcceptPrompts(false);
            state.taskA = StartFileOperationAndGetId(state.fileOps, FILESYSTEM_MOVE,
                                                     FolderWindow::Pane::Left, FolderWindow::Pane::Right,
                                                     ftp, {std::filesystem::path(ftpPath(L"br5/source"))},
                                                     std::filesystem::path(ftpPath(L"br5/destination")),
                                                     FILESYSTEM_FLAG_RECURSIVE, false, 0u,
                                                     FolderWindow::FileOperationState::ExecutionMode::PerItem, false, ftp);
            if (! state.taskA.has_value())
            {
                HostSetAutoAcceptPrompts(true);
                stopFakeFtp();
                curlModule.reset();
                Fail(L"BR-5 could not admit the same-connection host directory Move.");
                return true;
            }
            phaseStartTick = nowTick;
            state.stepState = 11u;
            return false;
        }

        if (committedDeleteWitness)
        {
            unsigned int requests = 0u;
            unsigned int commits  = 0u;
            BOOL stateCorrect     = FALSE;
            if (FAILED(committedDelete(TRUE, requests, commits, stateCorrect)))
            {
                stopFakeFtp();
                Fail(L"C0-Curl: failed to arm the committed Delete fixture.");
                return true;
            }
            state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                     FILESYSTEM_DELETE,
                                                     FolderWindow::Pane::Left,
                                                     std::nullopt,
                                                     ftp,
                                                     {std::filesystem::path(ftpPath(partialDeleteWitness ? L"selected" : L"c0/source.bin"))},
                                                     {},
                                                     partialDeleteWitness ? FILESYSTEM_FLAG_RECURSIVE : FILESYSTEM_FLAG_NONE,
                                                     false);
            if (! state.taskA.has_value())
            {
                stopFakeFtp();
                Fail(L"C0-Curl: failed to admit the committed Delete witness.");
                return true;
            }
            phaseStartTick  = nowTick;
            state.stepState = 10u;
            return false;
        }

        if (! RecreateEmptyDirectory(localSource) || ! RecreateEmptyDirectory(localBack))
        {
            stopFakeFtp();
            Fail(L"R0f-Curl: failed to reset the local fixtures.");
            return true;
        }
        for (unsigned int index = 0u; index < kFileCount; ++index)
        {
            if (! WriteFilledTestFile(localSource / std::format(L"payload-{:02}.bin", index), kFileBytes, static_cast<unsigned char>(0x61u + index)))
            {
                stopFakeFtp();
                Fail(L"R0f-Curl: failed to seed the local payloads.");
                return true;
            }
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSource},
                                                 std::filesystem::path(ftpRoot),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 ftp);
        if (! state.taskA.has_value())
        {
            stopFakeFtp();
            Fail(L"R0f-Curl: failed to start the Copy to the fake FTP server.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fCurl.CopyToFtpMs", nowTick - phaseStartTick, completed->second.hr);
        if (FAILED(completed->second.hr) || ! ftpExists(L"source/payload-00.bin") || ! ftpExists(L"source/payload-02.bin"))
        {
            const std::wstring diagnostics = taskDiagnostics();
            stopFakeFtp();
            Fail(std::format(L"R0f-Curl: Copy to the fake FTP server failed or left files missing (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             diagnostics));
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 ftp,
                                                 {std::filesystem::path(ftpPath(L"source"))},
                                                 localBack,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            stopFakeFtp();
            Fail(L"R0f-Curl: failed to start the Copy back from the fake FTP server.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 2u;
        return false;
    }

    if (state.stepState == 2u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fCurl.CopyFromFtpMs", nowTick - phaseStartTick, completed->second.hr);
        const std::filesystem::path copied = localBack / L"source";
        bool byteExact                     = SUCCEEDED(completed->second.hr) && CountFilesRecursive(copied) == kFileCount;
        for (unsigned int index = 0u; byteExact && index < kFileCount; ++index)
        {
            const std::filesystem::path leaf = std::format(L"payload-{:02}.bin", index);
            byteExact                        = FilesEqualBytes(localSource / leaf, copied / leaf);
        }
        if (! byteExact)
        {
            stopFakeFtp();
            Fail(std::format(L"R0f-Curl: Copy back from the fake FTP server failed or is not byte-exact (hr=0x{:08X}).", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> ftpDirectoryOps;
        HRESULT createHr = ftp->QueryInterface(__uuidof(IFileSystemDirectoryOperations), ftpDirectoryOps.put_void());
        if (SUCCEEDED(createHr))
        {
            createHr = ftpDirectoryOps->CreateDirectory(ftpPath(L"created").c_str());
        }
        if (FAILED(createHr) || ! ftpExists(L"created"))
        {
            stopFakeFtp();
            Fail(std::format(L"R0f-Curl: provider CreateDirectory on the fake FTP server failed (hr=0x{:08X}).", static_cast<unsigned long>(createHr)));
            return true;
        }
        const HRESULT renameHr = ftp->RenameItem(ftpPath(L"source/payload-01.bin").c_str(), ftpPath(L"source/renamed.bin").c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        if (FAILED(renameHr) || ftpExists(L"source/payload-01.bin") || ! ftpExists(L"source/renamed.bin"))
        {
            stopFakeFtp();
            Fail(std::format(L"R0f-Curl: provider Rename on the fake FTP server failed (hr=0x{:08X}).", static_cast<unsigned long>(renameHr)));
            return true;
        }

        // R3-1: replace through the provider's atomic-final writer (this route binds no destination
        // objects): the source payload changes, the Copy collides, the Exists prompt offers
        // Overwrite, and the replaced object is copied back byte-exact.
        if (! WriteFilledTestFile(localSource / L"payload-00.bin", kFileBytes, 0x7Au))
        {
            stopFakeFtp();
            Fail(L"R0f-Curl: failed to change the local payload for the replace step.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSource},
                                                 std::filesystem::path(ftpRoot),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 ftp);
        if (! state.taskA.has_value())
        {
            stopFakeFtp();
            Fail(L"R0f-Curl: failed to start the replacing Copy to the fake FTP endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 3u;
        return false;
    }

    if (state.stepState == 3u)
    {
        using Task = FolderWindow::FileOperationState::Task;
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    stopFakeFtp();
                    Fail(L"R0f-Curl: the Exists prompt on the identity-less FTP route must offer Overwrite.");
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, true);
                return false;
            }
        }
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fCurl.ReplaceOnFTPMs", nowTick - phaseStartTick, completed->second.hr);
        if (FAILED(completed->second.hr) || completed->second.conflictPromptCount == 0u)
        {
            const std::wstring diagnostics = taskDiagnostics();
            stopFakeFtp();
            Fail(std::format(L"R0f-Curl: replacing through the atomic-final writer failed (hr=0x{:08X}, prompts={}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             completed->second.conflictPromptCount,
                             diagnostics));
            return true;
        }
        const std::filesystem::path localBack2 = localBase / L"copied-back-2";
        if (! RecreateEmptyDirectory(localBack2))
        {
            stopFakeFtp();
            Fail(L"R0f-Curl: failed to reset the second copy-back folder.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 ftp,
                                                 {std::filesystem::path(ftpPath(L"source"))},
                                                 localBack2,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            stopFakeFtp();
            Fail(L"R0f-Curl: failed to start the second Copy back from the fake FTP endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 4u;
        return false;
    }

    if (state.stepState == 4u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        const std::filesystem::path copied2 = localBase / L"copied-back-2" / L"source";
        if (FAILED(completed->second.hr) || ! FilesEqualBytes(localSource / L"payload-00.bin", copied2 / L"payload-00.bin") ||
            ! FilesEqualBytes(localSource / L"payload-02.bin", copied2 / L"payload-02.bin"))
        {
            const std::wstring diagnostics = taskDiagnostics();
            stopFakeFtp();
            Fail(std::format(L"R0f-Curl: the replaced object did not come back with the new bytes (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             diagnostics));
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Right,
                                                 std::nullopt,
                                                 ftp,
                                                 {std::filesystem::path(ftpPath(L"source")), std::filesystem::path(ftpPath(L"created"))},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false);
        if (! state.taskA.has_value())
        {
            stopFakeFtp();
            Fail(L"R0f-Curl: failed to start the Delete on the fake FTP server.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 5u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    if (state.stepState == 7u)
    {
        const bool victimRemains = ftpExists(kC10Leaf);
        const bool canceled      = completed->second.hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && victimRemains;
        stopFakeFtp();
        curlModule.reset();
        if (! canceled)
        {
            Fail(std::format(L"C10-Curl: Cancel on the by-name card must delete nothing (hr=0x{:08X}, remains={}).",
                             static_cast<unsigned long>(completed->second.hr),
                             victimRemains));
            return true;
        }
        static_cast<void>(SelfTest::RemoveAll(localBase));
        NextStep(state, SelfTestState::Step::BR5_CurlHostDirectoryMove);
        return false;
    }
    Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fCurl.DeleteOnFtpMs", nowTick - phaseStartTick, completed->second.hr);
    const bool sourceRemains  = ftpExists(L"source");
    const bool createdRemains = ftpExists(L"created");
    const bool renamedRemains = ftpExists(L"source/renamed.bin");
    const bool deleted        = SUCCEEDED(completed->second.hr) && ! sourceRemains && ! createdRemains && ! renamedRemains;
    std::wstring deleteDetail;
    if (! deleted)
    {
        for (const std::optional<FileOperations::FileOperationItemResult>& item : completed->second.sourceItemResults)
        {
            if (item.has_value())
            {
                deleteDetail += std::format(L" [item status=0x{:08X} completion={} publication={} source={} path={}]",
                                            static_cast<unsigned long>(item->status),
                                            static_cast<unsigned int>(item->completion),
                                            static_cast<unsigned int>(item->publication),
                                            static_cast<unsigned int>(item->sourceDisposition),
                                            item->finalSourcePath);
            }
        }
        deleteDetail += taskDiagnostics();
        deleteDetail += std::format(L" remains: source={} created={} renamed={}", sourceRemains, createdRemains, renamedRemains);
    }
    if (! deleted)
    {
        stopFakeFtp();
        curlModule.reset();
        Fail(std::format(L"R0f-Curl: Delete on the fake FTP server failed or left objects behind (hr=0x{:08X}).{}",
                         static_cast<unsigned long>(completed->second.hr),
                         deleteDetail));
        return true;
    }
    // C10 phase 5: a Permanent Delete on a route without identity asks on the card and says so.
    {
        wil::com_ptr<IFileWriter> writer;
        const std::wstring victim = ftpPath(kC10Leaf);
        HRESULT hr                = ftpIo ? ftpIo->CreateFileWriter(victim.c_str(), FILESYSTEM_FLAG_NONE, writer.put()) : E_POINTER;
        if (SUCCEEDED(hr) && writer)
        {
            constexpr std::string_view kPayload = "object deleted by name";
            unsigned long written               = 0u;
            hr                                  = writer->Write(kPayload.data(), static_cast<unsigned long>(kPayload.size()), &written);
            if (SUCCEEDED(hr))
            {
                hr = writer->Commit();
            }
        }
        if (FAILED(hr) || ! ftpExists(kC10Leaf))
        {
            stopFakeFtp();
            curlModule.reset();
            Fail(std::format(L"C10-Curl: failed to create the by-name victim on the fake FTP server (hr=0x{:08X}).", static_cast<unsigned long>(hr)));
            return true;
        }
        HostSetAutoAcceptPrompts(false);
        HostClearTestPromptResultOverride();
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Right,
                                                 std::nullopt,
                                                 ftp,
                                                 {std::filesystem::path(victim)},
                                                 {},
                                                 FILESYSTEM_FLAG_NONE,
                                                 false);
        if (! state.taskA.has_value())
        {
            HostSetAutoAcceptPrompts(true);
            stopFakeFtp();
            curlModule.reset();
            Fail(L"C10-Curl: failed to start the by-name Permanent Delete on the fake FTP server.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 6u;
        return false;
    }
}

case SelfTestState::Step::R0fS3_FakeS3ReadWriteCreateDelete:
{
    // R0f-S3 matrix through the host against the plugin's deterministic loopback S3 fixture
    // (path-style, anonymous): cross-provider Copy local -> S3, Copy S3 -> local (byte-exact),
    // provider CreateDirectory and Rename on S3, Delete on S3 through a task. The route must
    // qualify as providerWatchdog with a nonzero provider-owned bound.
    static wil::unique_hmodule s3Module;
    static void* fakeS3             = nullptr;
    static unsigned int fakeS3Port  = 0u;
    static wil::com_ptr<IFileSystem> s3;
    static wil::com_ptr<IFileSystemIO> s3Io;
    static std::wstring s3Root;
    static ULONGLONG phaseStartTick = 0u;
    using StopFakeS3Fn              = void(__stdcall*)(void*) noexcept;
    const auto stopFakeS3 = []() noexcept
    {
        s3Io.reset();
        s3.reset();
        if (fakeS3 != nullptr && s3Module)
        {
            if (const FARPROC stopAddress = GetProcAddress(s3Module.get(), "RedSalamanderS3StopFakeS3ForSelfTest"))
            {
#pragma warning(push)
#pragma warning(disable : 4191)
                reinterpret_cast<StopFakeS3Fn>(stopAddress)(fakeS3);
#pragma warning(pop)
            }
        }
        fakeS3 = nullptr;
    };
    using FakeS3LogFn     = HRESULT(__stdcall*)(void*, wchar_t*, unsigned int) noexcept;
    const auto fixtureLog = []() noexcept -> std::wstring
    {
        std::wstring text;
        if (fakeS3 == nullptr || ! s3Module)
        {
            return text;
        }
        const FARPROC logAddress = GetProcAddress(s3Module.get(), "RedSalamanderS3FakeS3RequestLogForSelfTest");
        if (logAddress == nullptr)
        {
            return text;
        }
        wchar_t buffer[4096];
#pragma warning(push)
#pragma warning(disable : 4191)
        const HRESULT hr = reinterpret_cast<FakeS3LogFn>(logAddress)(fakeS3, buffer, static_cast<unsigned int>(std::size(buffer)));
#pragma warning(pop)
        if (SUCCEEDED(hr))
        {
            text = L" fixture:";
            text += buffer;
        }
        return text;
    };
    const auto taskDiagnostics = [&]() noexcept -> std::wstring
    {
        std::wstring text;
        if (! state.taskA.has_value())
        {
            return text;
        }
        std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> entries;
        state.fileOps->CollectDiagnostics(entries);
        unsigned int shown = 0u;
        for (const auto& entry : entries)
        {
            if (entry.taskId != state.taskA.value() || shown >= 6u)
            {
                continue;
            }
            text += std::format(L" [{} 0x{:08X} {}]", entry.category, static_cast<unsigned long>(entry.status), entry.message);
            ++shown;
        }
        return text;
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
        stopFakeS3();
        Fail(std::format(L"R0fS3_FakeS3ReadWriteCreateDelete timed out at step {}.{}", state.stepState, diagnostics));
        return true;
    }

    constexpr unsigned int kFileCount = 3u;
    constexpr size_t kFileBytes       = 64u * 1024u;
    const std::filesystem::path localBase   = state.tempRoot / L"r0f-s3-local";
    const std::filesystem::path localSource = localBase / L"source";
    const std::filesystem::path localBack   = localBase / L"copied-back";
    const auto s3Path = [&](std::wstring_view leaf) { return s3Root + std::wstring(leaf); };
    const auto s3Exists = [&](std::wstring_view leaf) noexcept
    {
        unsigned long attributes = 0u;
        return s3Io && SUCCEEDED(s3Io->GetAttributes(s3Path(leaf).c_str(), &attributes));
    };
    // C10: the object a Permanent Delete pins, written and replaced through the plugin as another
    // client would.
    constexpr std::wstring_view kC10Leaf = L"c10-pinned.bin";
    const auto c10Write                  = [&](std::wstring_view leaf, std::string_view payload) noexcept -> HRESULT
    {
        wil::com_ptr<IFileWriter> writer;
        const std::wstring path = s3Path(leaf);
        HRESULT hr              = s3Io ? s3Io->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_NONE, writer.put()) : E_POINTER;
        if (FAILED(hr) || ! writer)
        {
            return FAILED(hr) ? hr : E_UNEXPECTED;
        }
        unsigned long written = 0u;
        hr                    = writer->Write(payload.data(), static_cast<unsigned long>(payload.size()), &written);
        if (FAILED(hr))
        {
            return hr;
        }
        return writer->Commit();
    };
    const auto c10Swap = [&](std::string_view replacement) noexcept -> HRESULT
    {
        // A concurrent client replaces the object: the name loses its object and gains another
        // (different bytes, so the fake's content ETag changes as a real replacement would).
        const std::wstring path = s3Path(kC10Leaf);
        FileSystemOptions options{};
        options.sizeBytes = sizeof(options);
        const HRESULT hr  = s3->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, &options, nullptr, nullptr);
        if (FAILED(hr))
        {
            return hr;
        }
        return c10Write(kC10Leaf, replacement);
    };
    const auto c10ProviderContract = [&]() noexcept -> HRESULT
    {
        HRESULT hr = c10Write(kC10Leaf, "object the user confirmed for deletion");
        if (FAILED(hr))
        {
            return hr;
        }
        wil::com_ptr<IFileSystemIdentityDelete> identityDelete;
        if (FAILED(s3->QueryInterface(__uuidof(IFileSystemIdentityDelete), identityDelete.put_void())) || ! identityDelete)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        const std::wstring path = s3Path(kC10Leaf);
        FileSystemOptions options{};
        options.sizeBytes = sizeof(options);
        FileSystemDeleteIdentity pinned{};
        pinned.sizeBytes = sizeof(pinned);
        hr               = identityDelete->ResolveDeleteIdentity(path.c_str(), &options, &pinned);
        if (FAILED(hr))
        {
            return hr;
        }
        if (pinned.identity[0] == L'\0' || pinned.isDirectory != FALSE)
        {
            return E_UNEXPECTED;
        }
        hr = c10Swap("replacement object that the confirmed delete must not remove");
        if (FAILED(hr))
        {
            return hr;
        }
        hr = identityDelete->DeleteIfIdentity(path.c_str(), &pinned, FILESYSTEM_FLAG_NONE, &options, nullptr, nullptr);
        if (hr != HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH))
        {
            return SUCCEEDED(hr) ? E_UNEXPECTED : hr;
        }
        return s3Exists(kC10Leaf) ? S_OK : HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    };

    // C10 phase 6: the pinned Permanent Delete is held before its post-confirmation re-check while
    // the object is replaced. The hold precedes completion, so this runs before the completion lookup.
    if (state.stepState == 6u)
    {
        if (! HasFileOpsPermanentDeleteBeforeRecheckPauseEnteredForSelfTest())
        {
            if (state.taskA.has_value() && state.completedTasks.contains(state.taskA.value()))
            {
                ReleaseFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest();
                SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(false);
                stopFakeS3();
                s3Module.reset();
                Fail(L"C10-S3: the pinned Permanent Delete completed before its pre-recheck hold.");
                return true;
            }
            return false;
        }
        const HRESULT swapHr = c10Swap("second replacement object that the pinned Permanent Delete must not remove");
        ReleaseFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest();
        SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(false);
        if (FAILED(swapHr))
        {
            stopFakeS3();
            s3Module.reset();
            Fail(std::format(L"C10-S3: replacing the pinned object during the hold failed (hr=0x{:08X}).", static_cast<unsigned long>(swapHr)));
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 7u;
        return false;
    }

    if (state.stepState == 0u)
    {
        using StartFakeS3Fn = HRESULT(__stdcall*)(unsigned int*, void**) noexcept;
        using CreateFn      = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);
        FARPROC startAddress = nullptr;
        HRESULT hr = SelfTest::LoadPluginSelfTestExport(kPluginIdS3, "RedSalamanderS3StartFakeS3ForSelfTest", s3Module, startAddress);
        if (FAILED(hr) || startAddress == nullptr)
        {
            Fail(std::format(L"R0f-S3: the fake S3 fixture export is unavailable: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        const FARPROC createAddress = GetProcAddress(s3Module.get(), "RedSalamanderCreate");
        if (createAddress == nullptr)
        {
            Fail(L"R0f-S3: the S3 plugin factory export is unavailable.");
            return true;
        }
#pragma warning(push)
#pragma warning(disable : 4191)
        hr = reinterpret_cast<StartFakeS3Fn>(startAddress)(&fakeS3Port, &fakeS3);
#pragma warning(pop)
        if (FAILED(hr) || fakeS3 == nullptr || fakeS3Port == 0u)
        {
            Fail(std::format(L"R0f-S3: the fake S3 fixture did not start: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        FactoryOptions factoryOptions{};
        factoryOptions.debugLevel = DEBUG_LEVEL_NONE;
#pragma warning(push)
#pragma warning(disable : 4191)
        hr = reinterpret_cast<CreateFn>(createAddress)(__uuidof(IFileSystem), &factoryOptions, GetHostServices(), kPluginIdS3.data(), s3.put_void());
#pragma warning(pop)
        if (FAILED(hr) || ! s3)
        {
            stopFakeS3();
            Fail(std::format(L"R0f-S3: failed to create the S3 provider instance: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        const std::string configuration = std::format(
            R"({{"defaultRegion":"us-east-1","defaultEndpointOverride":"http://127.0.0.1:{}","useHttps":false,"verifyTls":false,"useVirtualAddressing":false,"anonymous":true,"connectTimeoutMs":2000,"requestTimeoutMs":8000}})",
            fakeS3Port);
        wil::com_ptr<IInformations> information;
        if (FAILED(s3->QueryInterface(__uuidof(IInformations), information.put_void())) || ! information ||
            FAILED(information->SetConfiguration(configuration.c_str())) || FAILED(s3->QueryInterface(__uuidof(IFileSystemIO), s3Io.put_void())) || ! s3Io)
        {
            stopFakeS3();
            Fail(L"R0f-S3: the S3 provider did not accept the fixture configuration or lacks IFileSystemIO.");
            return true;
        }
        s3Root = L"/r0f-bucket/";

        FileOperations::QualifiedEndpoint s3Endpoint{};
        if (! FileOperations::TryQualifyEndpointForSelfTest(s3, s3Root, FILESYSTEM_COPY, kPluginIdS3, L"host/default", s3Endpoint) ||
            s3Endpoint.cancellationRouteClass != FileOperations::CancellationRouteClass::ProviderWatchdog || s3Endpoint.providerWatchdogTimeoutMs == 0u)
        {
            stopFakeS3();
            Fail(L"R0f-S3: the S3 route must qualify as providerWatchdog with a nonzero provider-owned bound.");
            return true;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fS3.ProviderWatchdogTimeoutMs", s3Endpoint.providerWatchdogTimeoutMs, S_OK);

        if (! RecreateEmptyDirectory(localSource) || ! RecreateEmptyDirectory(localBack))
        {
            stopFakeS3();
            Fail(L"R0f-S3: failed to reset the local fixtures.");
            return true;
        }
        for (unsigned int index = 0u; index < kFileCount; ++index)
        {
            if (! WriteFilledTestFile(localSource / std::format(L"payload-{:02}.bin", index), kFileBytes, static_cast<unsigned char>(0x61u + index)))
            {
                stopFakeS3();
                Fail(L"R0f-S3: failed to seed the local payloads.");
                return true;
            }
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSource},
                                                 std::filesystem::path(s3Root),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 s3);
        if (! state.taskA.has_value())
        {
            stopFakeS3();
            Fail(L"R0f-S3: failed to start the Copy to the fake S3 endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fS3.CopyToS3Ms", nowTick - phaseStartTick, completed->second.hr);
        if (FAILED(completed->second.hr) || ! s3Exists(L"source/payload-00.bin") || ! s3Exists(L"source/payload-02.bin"))
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeS3();
            Fail(std::format(L"R0f-S3: Copy to the fake S3 endpoint failed or left objects missing (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             diagnostics));
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 s3,
                                                 {std::filesystem::path(s3Path(L"source"))},
                                                 localBack,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            stopFakeS3();
            Fail(L"R0f-S3: failed to start the Copy back from the fake S3 endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 2u;
        return false;
    }

    if (state.stepState == 2u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fS3.CopyFromS3Ms", nowTick - phaseStartTick, completed->second.hr);
        const std::filesystem::path copied = localBack / L"source";
        bool byteExact                     = SUCCEEDED(completed->second.hr) && CountFilesRecursive(copied) == kFileCount;
        for (unsigned int index = 0u; byteExact && index < kFileCount; ++index)
        {
            const std::filesystem::path leaf = std::format(L"payload-{:02}.bin", index);
            byteExact                        = FilesEqualBytes(localSource / leaf, copied / leaf);
        }
        if (! byteExact)
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeS3();
            Fail(std::format(L"R0f-S3: Copy back from the fake S3 endpoint failed or is not byte-exact (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             diagnostics));
            return true;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> s3DirectoryOps;
        HRESULT createHr = s3->QueryInterface(__uuidof(IFileSystemDirectoryOperations), s3DirectoryOps.put_void());
        if (SUCCEEDED(createHr))
        {
            createHr = s3DirectoryOps->CreateDirectory(s3Path(L"created").c_str());
        }
        if (FAILED(createHr) || ! s3Exists(L"created"))
        {
            const std::wstring diagnostics = fixtureLog();
            stopFakeS3();
            Fail(std::format(L"R0f-S3: provider CreateDirectory on the fake S3 endpoint failed (hr=0x{:08X}).{}", static_cast<unsigned long>(createHr), diagnostics));
            return true;
        }
        const HRESULT renameHr =
            s3->RenameItem(s3Path(L"source/payload-01.bin").c_str(), s3Path(L"source/renamed.bin").c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        if (FAILED(renameHr) || s3Exists(L"source/payload-01.bin") || ! s3Exists(L"source/renamed.bin"))
        {
            const std::wstring diagnostics = fixtureLog();
            stopFakeS3();
            Fail(std::format(L"R0f-S3: provider Rename on the fake S3 endpoint failed (hr=0x{:08X}).{}", static_cast<unsigned long>(renameHr), diagnostics));
            return true;
        }

        // R3-1: replace through the provider's atomic-final writer (this route binds no destination
        // objects): the source payload changes, the Copy collides, the Exists prompt offers
        // Overwrite, and the replaced object is copied back byte-exact.
        if (! WriteFilledTestFile(localSource / L"payload-00.bin", kFileBytes, 0x7Au))
        {
            stopFakeS3();
            Fail(L"R0f-S3: failed to change the local payload for the replace step.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSource},
                                                 std::filesystem::path(s3Root),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 s3);
        if (! state.taskA.has_value())
        {
            stopFakeS3();
            Fail(L"R0f-S3: failed to start the replacing Copy to the fake S3 endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 3u;
        return false;
    }

    if (state.stepState == 3u)
    {
        using Task = FolderWindow::FileOperationState::Task;
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    stopFakeS3();
                    Fail(L"R0f-S3: the Exists prompt on the identity-less S3 route must offer Overwrite.");
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, true);
                return false;
            }
        }
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fS3.ReplaceOnS3Ms", nowTick - phaseStartTick, completed->second.hr);
        if (FAILED(completed->second.hr) || completed->second.conflictPromptCount == 0u)
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeS3();
            Fail(std::format(L"R0f-S3: replacing through the atomic-final writer failed (hr=0x{:08X}, prompts={}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             completed->second.conflictPromptCount,
                             diagnostics));
            return true;
        }
        const std::filesystem::path localBack2 = localBase / L"copied-back-2";
        if (! RecreateEmptyDirectory(localBack2))
        {
            stopFakeS3();
            Fail(L"R0f-S3: failed to reset the second copy-back folder.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 s3,
                                                 {std::filesystem::path(s3Path(L"source"))},
                                                 localBack2,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            stopFakeS3();
            Fail(L"R0f-S3: failed to start the second Copy back from the fake S3 endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 4u;
        return false;
    }

    if (state.stepState == 4u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        const std::filesystem::path copied2 = localBase / L"copied-back-2" / L"source";
        if (FAILED(completed->second.hr) || ! FilesEqualBytes(localSource / L"payload-00.bin", copied2 / L"payload-00.bin") ||
            ! FilesEqualBytes(localSource / L"payload-02.bin", copied2 / L"payload-02.bin"))
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeS3();
            Fail(std::format(L"R0f-S3: the replaced object did not come back with the new bytes (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             diagnostics));
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Right,
                                                 std::nullopt,
                                                 s3,
                                                 {std::filesystem::path(s3Path(L"source")), std::filesystem::path(s3Path(L"created"))},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false);
        if (! state.taskA.has_value())
        {
            stopFakeS3();
            Fail(L"R0f-S3: failed to start the Delete on the fake S3 endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 5u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    if (state.stepState == 7u)
    {
        Debug::Perf::EmitValue(L"FileOps.SelfTest.C10.S3PinnedDeleteRefusedMs", nowTick - phaseStartTick, completed->second.hr);
        const bool victimRemains = s3Exists(kC10Leaf);
        const bool refused       = completed->second.hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) && victimRemains;
        std::wstring refuseDetail;
        if (! refused)
        {
            refuseDetail += taskDiagnostics();
            refuseDetail += fixtureLog();
        }
        stopFakeS3();
        s3Module.reset();
        if (! refused)
        {
            Fail(std::format(L"C10-S3: a Permanent Delete whose object was replaced after the confirmation must refuse with "
                             L"ERROR_REVISION_MISMATCH and keep the new object (hr=0x{:08X}, remains={}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             victimRemains,
                             refuseDetail));
            return true;
        }
        static_cast<void>(SelfTest::RemoveAll(localBase));
        NextStep(state, SelfTestState::Step::R0fGraph_FakeGraphReadWriteCreateRenameRecycle);
        return false;
    }
    Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fS3.DeleteOnS3Ms", nowTick - phaseStartTick, completed->second.hr);
    const bool sourceRemains  = s3Exists(L"source");
    const bool createdRemains = s3Exists(L"created");
    const bool renamedRemains = s3Exists(L"source/renamed.bin");
    const bool deleted        = SUCCEEDED(completed->second.hr) && ! sourceRemains && ! createdRemains && ! renamedRemains;
    std::wstring deleteDetail;
    if (! deleted)
    {
        for (const std::optional<FileOperations::FileOperationItemResult>& item : completed->second.sourceItemResults)
        {
            if (item.has_value())
            {
                deleteDetail += std::format(L" [item status=0x{:08X} completion={} publication={} source={} path={}]",
                                            static_cast<unsigned long>(item->status),
                                            static_cast<unsigned int>(item->completion),
                                            static_cast<unsigned int>(item->publication),
                                            static_cast<unsigned int>(item->sourceDisposition),
                                            item->finalSourcePath);
            }
        }
        deleteDetail += taskDiagnostics();
        deleteDetail += fixtureLog();
        deleteDetail += std::format(L" remains: source={} created={} renamed={}", sourceRemains, createdRemains, renamedRemains);
    }
    if (! deleted)
    {
        stopFakeS3();
        s3Module.reset();
        Fail(std::format(L"R0f-S3: Delete on the fake S3 endpoint failed or left objects behind (hr=0x{:08X}).{}",
                         static_cast<unsigned long>(completed->second.hr),
                         deleteDetail));
        return true;
    }
    // C10 phase 5: the provider contract on its own, then the host's pinned Permanent Delete held
    // before its post-confirmation re-check.
    {
        const HRESULT contractHr = c10ProviderContract();
        if (FAILED(contractHr))
        {
            const std::wstring detail = fixtureLog();
            stopFakeS3();
            s3Module.reset();
            Fail(std::format(L"C10-S3: the identity-delete contract must pin the object revision and refuse a replaced object (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(contractHr),
                             detail));
            return true;
        }
        SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(true);
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Right,
                                                 std::nullopt,
                                                 s3,
                                                 {std::filesystem::path(s3Path(kC10Leaf))},
                                                 {},
                                                 FILESYSTEM_FLAG_NONE,
                                                 false);
        HostClearTestPromptResultOverride();
        if (! state.taskA.has_value())
        {
            ReleaseFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest();
            SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(false);
            stopFakeS3();
            s3Module.reset();
            Fail(L"C10-S3: failed to start the pinned Permanent Delete on the fake S3 endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 6u;
        return false;
    }
}

case SelfTestState::Step::R0fGraph_FakeGraphReadWriteCreateRenameRecycle:
{
    // R0f-Graph matrix through the host against the plugin's deterministic loopback Graph fixture:
    // cross-provider Copy local -> Graph (upload), Copy Graph -> local (ranged download, byte-exact),
    // provider CreateDirectory and Rename on Graph, Recycle on Graph through a task (Graph Delete is
    // Recycle by stable item ID). The route must qualify as providerWatchdog with a nonzero bound.
    static wil::unique_hmodule graphModule;
    static void* fakeGraph            = nullptr;
    static unsigned int fakeGraphPort = 0u;
    static wil::com_ptr<IFileSystem> graph;
    static wil::com_ptr<IFileSystemIO> graphIo;
    static std::wstring graphRoot;
    static ULONGLONG phaseStartTick = 0u;
    using StopFakeGraphFn           = void(__stdcall*)(void*) noexcept;
    using FakeGraphLogFn            = HRESULT(__stdcall*)(void*, wchar_t*, unsigned int) noexcept;
    const auto stopFakeGraph = []() noexcept
    {
        graphIo.reset();
        graph.reset();
        if (fakeGraph != nullptr && graphModule)
        {
            if (const FARPROC stopAddress = GetProcAddress(graphModule.get(), "RedSalamanderMicrosoftDriveStopFakeGraphForSelfTest"))
            {
#pragma warning(push)
#pragma warning(disable : 4191)
                reinterpret_cast<StopFakeGraphFn>(stopAddress)(fakeGraph);
#pragma warning(pop)
            }
        }
        fakeGraph = nullptr;
    };
    const auto fixtureLog = []() noexcept -> std::wstring
    {
        std::wstring text;
        if (fakeGraph == nullptr || ! graphModule)
        {
            return text;
        }
        const FARPROC logAddress = GetProcAddress(graphModule.get(), "RedSalamanderMicrosoftDriveFakeGraphRequestLogForSelfTest");
        if (logAddress == nullptr)
        {
            return text;
        }
        wchar_t buffer[4096];
#pragma warning(push)
#pragma warning(disable : 4191)
        const HRESULT hr = reinterpret_cast<FakeGraphLogFn>(logAddress)(fakeGraph, buffer, static_cast<unsigned int>(std::size(buffer)));
#pragma warning(pop)
        if (SUCCEEDED(hr))
        {
            text = L" fixture:";
            text += buffer;
        }
        return text;
    };
    const auto taskDiagnostics = [&]() noexcept -> std::wstring
    {
        std::wstring text;
        if (! state.taskA.has_value())
        {
            return text;
        }
        std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> entries;
        state.fileOps->CollectDiagnostics(entries);
        unsigned int shown = 0u;
        for (const auto& entry : entries)
        {
            if (entry.taskId != state.taskA.value() || shown >= 6u)
            {
                continue;
            }
            text += std::format(L" [{} 0x{:08X} {}]", entry.category, static_cast<unsigned long>(entry.status), entry.message);
            ++shown;
        }
        return text;
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
        stopFakeGraph();
        Fail(std::format(L"R0fGraph_FakeGraphReadWriteCreateRenameRecycle timed out at step {}.{}", state.stepState, diagnostics));
        return true;
    }

    constexpr unsigned int kFileCount = 3u;
    constexpr size_t kFileBytes       = 64u * 1024u;
    const std::filesystem::path localBase   = state.tempRoot / L"r0f-graph-local";
    const std::filesystem::path localSource = localBase / L"source";
    const std::filesystem::path localBack   = localBase / L"copied-back";
    const auto graphPath = [&](std::wstring_view leaf) { return graphRoot + std::wstring(leaf); };
    const auto graphExists = [&](std::wstring_view leaf) noexcept
    {
        unsigned long attributes = 0u;
        return graphIo && SUCCEEDED(graphIo->GetAttributes(graphPath(leaf).c_str(), &attributes));
    };

    if (state.stepState == 0u)
    {
        using StartFakeGraphFn = HRESULT(__stdcall*)(unsigned int*, void**) noexcept;
        using CreateFn         = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);
        FARPROC startAddress = nullptr;
        HRESULT hr = SelfTest::LoadPluginSelfTestExport(kPluginIdOneDrivePersonal, "RedSalamanderMicrosoftDriveStartFakeGraphForSelfTest", graphModule, startAddress);
        if (FAILED(hr) || startAddress == nullptr)
        {
            Fail(std::format(L"R0f-Graph: the fake Graph fixture export is unavailable: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        const FARPROC createAddress = GetProcAddress(graphModule.get(), "RedSalamanderCreate");
        if (createAddress == nullptr)
        {
            Fail(L"R0f-Graph: the Microsoft Drive plugin factory export is unavailable.");
            return true;
        }
#pragma warning(push)
#pragma warning(disable : 4191)
        hr = reinterpret_cast<StartFakeGraphFn>(startAddress)(&fakeGraphPort, &fakeGraph);
#pragma warning(pop)
        if (FAILED(hr) || fakeGraph == nullptr || fakeGraphPort == 0u)
        {
            Fail(std::format(L"R0f-Graph: the fake Graph fixture did not start: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        FactoryOptions factoryOptions{};
        factoryOptions.debugLevel = DEBUG_LEVEL_NONE;
#pragma warning(push)
#pragma warning(disable : 4191)
        hr = reinterpret_cast<CreateFn>(createAddress)(__uuidof(IFileSystem), &factoryOptions, GetHostServices(), kPluginIdOneDrivePersonal.data(), graph.put_void());
#pragma warning(pop)
        if (FAILED(hr) || ! graph)
        {
            stopFakeGraph();
            Fail(std::format(L"R0f-Graph: failed to create the Microsoft Drive provider instance: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        wil::com_ptr<IInformations> information;
        if (FAILED(graph->QueryInterface(__uuidof(IInformations), information.put_void())) || ! information ||
            FAILED(information->SetConfiguration(R"({"connectTimeoutMs":2000,"requestTimeoutMs":8000})")) ||
            FAILED(graph->QueryInterface(__uuidof(IFileSystemIO), graphIo.put_void())) || ! graphIo)
        {
            stopFakeGraph();
            Fail(L"R0f-Graph: the Microsoft Drive provider did not accept the fixture configuration or lacks IFileSystemIO.");
            return true;
        }
        graphRoot = L"/@conn:microsoft-drive-selftest/";

        FileOperations::QualifiedEndpoint graphEndpoint{};
        if (! FileOperations::TryQualifyEndpointForSelfTest(graph, graphRoot, FILESYSTEM_COPY, kPluginIdOneDrivePersonal, L"host/default", graphEndpoint) ||
            graphEndpoint.cancellationRouteClass != FileOperations::CancellationRouteClass::ProviderWatchdog || graphEndpoint.providerWatchdogTimeoutMs == 0u)
        {
            stopFakeGraph();
            Fail(L"R0f-Graph: the Graph route must qualify as providerWatchdog with a nonzero provider-owned bound.");
            return true;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fGraph.ProviderWatchdogTimeoutMs", graphEndpoint.providerWatchdogTimeoutMs, S_OK);

        if (! RecreateEmptyDirectory(localSource) || ! RecreateEmptyDirectory(localBack))
        {
            stopFakeGraph();
            Fail(L"R0f-Graph: failed to reset the local fixtures.");
            return true;
        }
        for (unsigned int index = 0u; index < kFileCount; ++index)
        {
            if (! WriteFilledTestFile(localSource / std::format(L"payload-{:02}.bin", index), kFileBytes, static_cast<unsigned char>(0x61u + index)))
            {
                stopFakeGraph();
                Fail(L"R0f-Graph: failed to seed the local payloads.");
                return true;
            }
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSource},
                                                 std::filesystem::path(graphRoot),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 graph);
        if (! state.taskA.has_value())
        {
            stopFakeGraph();
            Fail(L"R0f-Graph: failed to start the Copy to the fake Graph endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fGraph.CopyToGraphMs", nowTick - phaseStartTick, completed->second.hr);
        if (FAILED(completed->second.hr) || ! graphExists(L"source/payload-00.bin") || ! graphExists(L"source/payload-02.bin"))
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeGraph();
            Fail(std::format(L"R0f-Graph: Copy to the fake Graph endpoint failed or left items missing (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             diagnostics));
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 graph,
                                                 {std::filesystem::path(graphPath(L"source"))},
                                                 localBack,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            stopFakeGraph();
            Fail(L"R0f-Graph: failed to start the Copy back from the fake Graph endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 2u;
        return false;
    }

    if (state.stepState == 2u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fGraph.CopyFromGraphMs", nowTick - phaseStartTick, completed->second.hr);
        const std::filesystem::path copied = localBack / L"source";
        bool byteExact                     = SUCCEEDED(completed->second.hr) && CountFilesRecursive(copied) == kFileCount;
        for (unsigned int index = 0u; byteExact && index < kFileCount; ++index)
        {
            const std::filesystem::path leaf = std::format(L"payload-{:02}.bin", index);
            byteExact                        = FilesEqualBytes(localSource / leaf, copied / leaf);
        }
        if (! byteExact)
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeGraph();
            Fail(std::format(L"R0f-Graph: Copy back from the fake Graph endpoint failed or is not byte-exact (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             diagnostics));
            return true;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> graphDirectoryOps;
        HRESULT createHr = graph->QueryInterface(__uuidof(IFileSystemDirectoryOperations), graphDirectoryOps.put_void());
        if (SUCCEEDED(createHr))
        {
            createHr = graphDirectoryOps->CreateDirectory(graphPath(L"created").c_str());
        }
        if (FAILED(createHr) || ! graphExists(L"created"))
        {
            const std::wstring diagnostics = fixtureLog();
            stopFakeGraph();
            Fail(std::format(L"R0f-Graph: provider CreateDirectory on the fake Graph endpoint failed (hr=0x{:08X}).{}", static_cast<unsigned long>(createHr), diagnostics));
            return true;
        }
        const HRESULT renameHr =
            graph->RenameItem(graphPath(L"source/payload-01.bin").c_str(), graphPath(L"source/renamed.bin").c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        if (FAILED(renameHr) || graphExists(L"source/payload-01.bin") || ! graphExists(L"source/renamed.bin"))
        {
            const std::wstring diagnostics = fixtureLog();
            stopFakeGraph();
            Fail(std::format(L"R0f-Graph: provider Rename on the fake Graph endpoint failed (hr=0x{:08X}).{}", static_cast<unsigned long>(renameHr), diagnostics));
            return true;
        }

        // R3-1: replace through the provider's atomic-final writer (this route binds no destination
        // objects): the source payload changes, the Copy collides, the Exists prompt offers
        // Overwrite, and the replaced object is copied back byte-exact.
        if (! WriteFilledTestFile(localSource / L"payload-00.bin", kFileBytes, 0x7Au))
        {
            stopFakeGraph();
            Fail(L"R0f-Graph: failed to change the local payload for the replace step.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSource},
                                                 std::filesystem::path(graphRoot),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 graph);
        if (! state.taskA.has_value())
        {
            stopFakeGraph();
            Fail(L"R0f-Graph: failed to start the replacing Copy to the fake Graph endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 3u;
        return false;
    }

    if (state.stepState == 3u)
    {
        using Task = FolderWindow::FileOperationState::Task;
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    stopFakeGraph();
                    Fail(L"R0f-Graph: the Exists prompt on the identity-less Graph route must offer Overwrite.");
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, true);
                return false;
            }
        }
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fGraph.ReplaceOnGraphMs", nowTick - phaseStartTick, completed->second.hr);
        if (FAILED(completed->second.hr) || completed->second.conflictPromptCount == 0u)
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeGraph();
            Fail(std::format(L"R0f-Graph: replacing through the atomic-final writer failed (hr=0x{:08X}, prompts={}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             completed->second.conflictPromptCount,
                             diagnostics));
            return true;
        }
        const std::filesystem::path localBack2 = localBase / L"copied-back-2";
        if (! RecreateEmptyDirectory(localBack2))
        {
            stopFakeGraph();
            Fail(L"R0f-Graph: failed to reset the second copy-back folder.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 graph,
                                                 {std::filesystem::path(graphPath(L"source"))},
                                                 localBack2,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            stopFakeGraph();
            Fail(L"R0f-Graph: failed to start the second Copy back from the fake Graph endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 4u;
        return false;
    }

    if (state.stepState == 4u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        const std::filesystem::path copied2 = localBase / L"copied-back-2" / L"source";
        if (FAILED(completed->second.hr) || ! FilesEqualBytes(localSource / L"payload-00.bin", copied2 / L"payload-00.bin") ||
            ! FilesEqualBytes(localSource / L"payload-02.bin", copied2 / L"payload-02.bin"))
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeGraph();
            Fail(std::format(L"R0f-Graph: the replaced object did not come back with the new bytes (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             diagnostics));
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Right,
                                                 std::nullopt,
                                                 graph,
                                                 {std::filesystem::path(graphPath(L"source")), std::filesystem::path(graphPath(L"created"))},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN),
                                                 false);
        if (! state.taskA.has_value())
        {
            stopFakeGraph();
            Fail(L"R0f-Graph: failed to start the Recycle on the fake Graph endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 5u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fGraph.RecycleOnGraphMs", nowTick - phaseStartTick, completed->second.hr);
    const bool sourceRemains  = graphExists(L"source");
    const bool createdRemains = graphExists(L"created");
    const bool renamedRemains = graphExists(L"source/renamed.bin");
    const bool recycled       = SUCCEEDED(completed->second.hr) && ! sourceRemains && ! createdRemains && ! renamedRemains;
    std::wstring recycleDetail;
    if (! recycled)
    {
        for (const std::optional<FileOperations::FileOperationItemResult>& item : completed->second.sourceItemResults)
        {
            if (item.has_value())
            {
                recycleDetail += std::format(L" [item status=0x{:08X} completion={} publication={} source={} path={}]",
                                             static_cast<unsigned long>(item->status),
                                             static_cast<unsigned int>(item->completion),
                                             static_cast<unsigned int>(item->publication),
                                             static_cast<unsigned int>(item->sourceDisposition),
                                             item->finalSourcePath);
            }
        }
        recycleDetail += taskDiagnostics();
        recycleDetail += fixtureLog();
        recycleDetail += std::format(L" remains: source={} created={} renamed={}", sourceRemains, createdRemains, renamedRemains);
    }
    stopFakeGraph();
    graphModule.reset();
    if (! recycled)
    {
        Fail(std::format(L"R0f-Graph: Recycle on the fake Graph endpoint failed or left items behind (hr=0x{:08X}).{}",
                         static_cast<unsigned long>(completed->second.hr),
                         recycleDetail));
        return true;
    }
    static_cast<void>(SelfTest::RemoveAll(localBase));
    NextStep(state, SelfTestState::Step::R0fGDrive_FakeDriveReadWriteCreateMoveDelete);
    return false;
}

case SelfTestState::Step::R0fGDrive_FakeDriveReadWriteCreateMoveDelete:
case SelfTestState::Step::C0_GDriveCommittedCopyResponseLost:
case SelfTestState::Step::C0_GDriveCommittedDeleteResponseLost:
{
    // R0f-GDrive matrix through the host against the plugin's deterministic loopback Drive fixture:
    // cross-provider Copy local -> Drive (resumable upload), Copy Drive -> local (ranged download,
    // byte-exact), provider CreateDirectory, Rename and native Move on Drive, then a permanent
    // Delete on Drive through a task. The route must qualify as providerWatchdog with a nonzero bound.
    static wil::unique_hmodule driveModule;
    static void* fakeDrive            = nullptr;
    static unsigned int fakeDrivePort = 0u;
    static wil::com_ptr<IFileSystem> drive;
    static wil::com_ptr<IFileSystemIO> driveIo;
    static std::wstring driveRoot;
    static ULONGLONG phaseStartTick            = 0u;
    const bool committedCopyWitness            = state.step == SelfTestState::Step::C0_GDriveCommittedCopyResponseLost;
    const bool committedDeleteWitness          = state.step == SelfTestState::Step::C0_GDriveCommittedDeleteResponseLost;
    const FileSystemOperation witnessOperation = committedDeleteWitness ? FILESYSTEM_DELETE : FILESYSTEM_COPY;
    using MutationCommitFailureFn              = HRESULT(__stdcall*)(void*, FileSystemOperation, BOOL, unsigned int*, unsigned int*) noexcept;
    const auto mutationCommitFailure           = [witnessOperation](BOOL arm, unsigned int& requests, unsigned int& commits) noexcept -> HRESULT
    {
        const FARPROC address = driveModule ? GetProcAddress(driveModule.get(), "RedSalamanderGoogleDriveMutationCommitFailureForSelfTest") : nullptr;
        if (address == nullptr)
        {
            return E_NOINTERFACE;
        }
#pragma warning(push)
#pragma warning(disable : 4191)
        return reinterpret_cast<MutationCommitFailureFn>(address)(fakeDrive, witnessOperation, arm, &requests, &commits);
#pragma warning(pop)
    };
    using StopFakeDriveFn    = void(__stdcall*)(void*) noexcept;
    using FakeDriveLogFn     = HRESULT(__stdcall*)(void*, wchar_t*, unsigned int) noexcept;
    const auto stopFakeDrive = []() noexcept
    {
        driveIo.reset();
        drive.reset();
        if (fakeDrive != nullptr && driveModule)
        {
            if (const FARPROC stopAddress = GetProcAddress(driveModule.get(), "RedSalamanderGoogleDriveStopFakeDriveForSelfTest"))
            {
#pragma warning(push)
#pragma warning(disable : 4191)
                reinterpret_cast<StopFakeDriveFn>(stopAddress)(fakeDrive);
#pragma warning(pop)
            }
        }
        fakeDrive = nullptr;
    };
    const auto fixtureLog = []() noexcept -> std::wstring
    {
        std::wstring text;
        if (fakeDrive == nullptr || ! driveModule)
        {
            return text;
        }
        const FARPROC logAddress = GetProcAddress(driveModule.get(), "RedSalamanderGoogleDriveFakeDriveRequestLogForSelfTest");
        if (logAddress == nullptr)
        {
            return text;
        }
        wchar_t buffer[4096];
#pragma warning(push)
#pragma warning(disable : 4191)
        const HRESULT hr = reinterpret_cast<FakeDriveLogFn>(logAddress)(fakeDrive, buffer, static_cast<unsigned int>(std::size(buffer)));
#pragma warning(pop)
        if (SUCCEEDED(hr))
        {
            text = L" fixture:";
            text += buffer;
        }
        return text;
    };
    const auto taskDiagnostics = [&]() noexcept -> std::wstring
    {
        std::wstring text;
        if (! state.taskA.has_value())
        {
            return text;
        }
        std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> entries;
        state.fileOps->CollectDiagnostics(entries);
        unsigned int shown = 0u;
        for (const auto& entry : entries)
        {
            if (entry.taskId != state.taskA.value() || shown >= 6u)
            {
                continue;
            }
            text += std::format(L" [{} 0x{:08X} {}]", entry.category, static_cast<unsigned long>(entry.status), entry.message);
            ++shown;
        }
        return text;
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
        stopFakeDrive();
        Fail(std::format(L"R0fGDrive_FakeDriveReadWriteCreateMoveDelete timed out at step {}.{}", state.stepState, diagnostics));
        return true;
    }

    constexpr unsigned int kFileCount = 3u;
    constexpr size_t kFileBytes       = 64u * 1024u;
    const std::filesystem::path localBase   = state.tempRoot / L"r0f-gdrive-local";
    const std::filesystem::path localSource = localBase / L"source";
    const std::filesystem::path localBack   = localBase / L"copied-back";
    const auto drivePath = [&](std::wstring_view leaf) { return driveRoot + std::wstring(leaf); };
    const auto driveExists = [&](std::wstring_view leaf) noexcept
    {
        unsigned long attributes = 0u;
        return driveIo && SUCCEEDED(driveIo->GetAttributes(drivePath(leaf).c_str(), &attributes));
    };
    // C10: the identity-delete contract on Drive pins the file id; a name that gained another file id
    // is refused.
    const auto c10DriveContract = [&]() noexcept -> HRESULT
    {
        constexpr std::wstring_view kLeaf = L"c10-pinned.bin";
        const auto write                  = [&](std::string_view payload) noexcept -> HRESULT
        {
            wil::com_ptr<IFileWriter> writer;
            const std::wstring path = drivePath(kLeaf);
            HRESULT hr              = driveIo ? driveIo->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_NONE, writer.put()) : E_POINTER;
            if (FAILED(hr) || ! writer)
            {
                return FAILED(hr) ? hr : E_UNEXPECTED;
            }
            unsigned long written = 0u;
            hr                    = writer->Write(payload.data(), static_cast<unsigned long>(payload.size()), &written);
            return FAILED(hr) ? hr : writer->Commit();
        };
        HRESULT hr = write("object the user confirmed for deletion");
        if (FAILED(hr))
        {
            return hr;
        }
        wil::com_ptr<IFileSystemIdentityDelete> identityDelete;
        if (FAILED(drive->QueryInterface(__uuidof(IFileSystemIdentityDelete), identityDelete.put_void())) || ! identityDelete)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        const std::wstring path = drivePath(kLeaf);
        FileSystemOptions options{};
        options.sizeBytes = sizeof(options);
        FileSystemDeleteIdentity pinned{};
        pinned.sizeBytes = sizeof(pinned);
        hr               = identityDelete->ResolveDeleteIdentity(path.c_str(), &options, &pinned);
        if (FAILED(hr))
        {
            return hr;
        }
        if (pinned.identity[0] == L'\0' || pinned.isDirectory != FALSE)
        {
            return E_UNEXPECTED;
        }
        hr = drive->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, &options, nullptr, nullptr);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = write("replacement file that the confirmed delete must not remove");
        if (FAILED(hr))
        {
            return hr;
        }
        hr = identityDelete->DeleteIfIdentity(path.c_str(), &pinned, FILESYSTEM_FLAG_NONE, &options, nullptr, nullptr);
        if (hr != HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH))
        {
            return SUCCEEDED(hr) ? E_UNEXPECTED : hr;
        }
        return driveExists(kLeaf) ? S_OK : HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    };

    if (state.stepState == 0u)
    {
        using StartFakeDriveFn = HRESULT(__stdcall*)(unsigned int*, void**) noexcept;
        using CreateFn         = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);
        FARPROC startAddress = nullptr;
        HRESULT hr = SelfTest::LoadPluginSelfTestExport(kPluginIdGoogleDrive, "RedSalamanderGoogleDriveStartFakeDriveForSelfTest", driveModule, startAddress);
        if (FAILED(hr) || startAddress == nullptr)
        {
            Fail(std::format(L"R0f-GDrive: the fake Drive fixture export is unavailable: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        const FARPROC createAddress = GetProcAddress(driveModule.get(), "RedSalamanderCreate");
        if (createAddress == nullptr)
        {
            Fail(L"R0f-GDrive: the Google Drive plugin factory export is unavailable.");
            return true;
        }
#pragma warning(push)
#pragma warning(disable : 4191)
        hr = reinterpret_cast<StartFakeDriveFn>(startAddress)(&fakeDrivePort, &fakeDrive);
#pragma warning(pop)
        if (FAILED(hr) || fakeDrive == nullptr || fakeDrivePort == 0u)
        {
            Fail(std::format(L"R0f-GDrive: the fake Drive fixture did not start: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        FactoryOptions factoryOptions{};
        factoryOptions.debugLevel = DEBUG_LEVEL_NONE;
#pragma warning(push)
#pragma warning(disable : 4191)
        hr = reinterpret_cast<CreateFn>(createAddress)(__uuidof(IFileSystem), &factoryOptions, GetHostServices(), kPluginIdGoogleDrive.data(), drive.put_void());
#pragma warning(pop)
        if (FAILED(hr) || ! drive)
        {
            stopFakeDrive();
            Fail(std::format(L"R0f-GDrive: failed to create the Google Drive provider instance: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        wil::com_ptr<IInformations> information;
        if (FAILED(drive->QueryInterface(__uuidof(IInformations), information.put_void())) || ! information ||
            FAILED(information->SetConfiguration(R"({"connectTimeoutMs":2000,"requestTimeoutMs":8000})")) ||
            FAILED(drive->QueryInterface(__uuidof(IFileSystemIO), driveIo.put_void())) || ! driveIo)
        {
            stopFakeDrive();
            Fail(L"R0f-GDrive: the Google Drive provider did not accept the fixture configuration or lacks IFileSystemIO.");
            return true;
        }
        driveRoot = L"/@conn:google-drive-selftest/";

        FileOperations::QualifiedEndpoint driveEndpoint{};
        if (! FileOperations::TryQualifyEndpointForSelfTest(drive, driveRoot, FILESYSTEM_COPY, kPluginIdGoogleDrive, L"host/default", driveEndpoint) ||
            driveEndpoint.cancellationRouteClass != FileOperations::CancellationRouteClass::ProviderWatchdog || driveEndpoint.providerWatchdogTimeoutMs == 0u)
        {
            stopFakeDrive();
            Fail(L"R0f-GDrive: the Drive route must qualify as providerWatchdog with a nonzero provider-owned bound.");
            return true;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fGDrive.ProviderWatchdogTimeoutMs", driveEndpoint.providerWatchdogTimeoutMs, S_OK);

        if (! RecreateEmptyDirectory(localSource) || ! RecreateEmptyDirectory(localBack))
        {
            stopFakeDrive();
            Fail(L"R0f-GDrive: failed to reset the local fixtures.");
            return true;
        }
        for (unsigned int index = 0u; index < kFileCount; ++index)
        {
            if (! WriteFilledTestFile(localSource / std::format(L"payload-{:02}.bin", index), kFileBytes, static_cast<unsigned char>(0x61u + index)))
            {
                stopFakeDrive();
                Fail(L"R0f-GDrive: failed to seed the local payloads.");
                return true;
            }
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSource},
                                                 std::filesystem::path(driveRoot),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 drive);
        if (! state.taskA.has_value())
        {
            stopFakeDrive();
            Fail(L"R0f-GDrive: failed to start the Copy to the fake Drive endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fGDrive.CopyToDriveMs", nowTick - phaseStartTick, completed->second.hr);
        if (FAILED(completed->second.hr) || ! driveExists(L"source/payload-00.bin") || ! driveExists(L"source/payload-02.bin"))
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeDrive();
            Fail(std::format(L"R0f-GDrive: Copy to the fake Drive endpoint failed or left items missing (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             diagnostics));
            return true;
        }
        if (committedCopyWitness || committedDeleteWitness)
        {
            unsigned int requests = 0u;
            unsigned int commits  = 0u;
            if (FAILED(mutationCommitFailure(TRUE, requests, commits)))
            {
                stopFakeDrive();
                Fail(L"C0-GDrive: failed to arm the commit-before-error fixture.");
                return true;
            }
            state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                     witnessOperation,
                                                     FolderWindow::Pane::Left,
                                                     FolderWindow::Pane::Right,
                                                     drive,
                                                     {std::filesystem::path(drivePath(L"source/payload-00.bin"))},
                                                     committedDeleteWitness ? std::filesystem::path{} : std::filesystem::path(driveRoot),
                                                     FILESYSTEM_FLAG_NONE,
                                                     false);
            if (! state.taskA.has_value())
            {
                stopFakeDrive();
                Fail(L"C0-GDrive: failed to start the committed-mutation witness.");
                return true;
            }
            phaseStartTick  = nowTick;
            state.stepState = 10u;
            return false;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 drive,
                                                 {std::filesystem::path(drivePath(L"source"))},
                                                 localBack,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            stopFakeDrive();
            Fail(L"R0f-GDrive: failed to start the Copy back from the fake Drive endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 2u;
        return false;
    }

    if (state.stepState == 2u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fGDrive.CopyFromDriveMs", nowTick - phaseStartTick, completed->second.hr);
        const std::filesystem::path copied = localBack / L"source";
        bool byteExact                     = SUCCEEDED(completed->second.hr) && CountFilesRecursive(copied) == kFileCount;
        for (unsigned int index = 0u; byteExact && index < kFileCount; ++index)
        {
            const std::filesystem::path leaf = std::format(L"payload-{:02}.bin", index);
            byteExact                        = FilesEqualBytes(localSource / leaf, copied / leaf);
        }
        if (! byteExact)
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeDrive();
            Fail(std::format(L"R0f-GDrive: Copy back from the fake Drive endpoint failed or is not byte-exact (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             diagnostics));
            return true;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> driveDirectoryOps;
        HRESULT createHr = drive->QueryInterface(__uuidof(IFileSystemDirectoryOperations), driveDirectoryOps.put_void());
        if (SUCCEEDED(createHr))
        {
            createHr = driveDirectoryOps->CreateDirectory(drivePath(L"created").c_str());
        }
        if (FAILED(createHr) || ! driveExists(L"created"))
        {
            const std::wstring diagnostics = fixtureLog();
            stopFakeDrive();
            Fail(std::format(L"R0f-GDrive: provider CreateDirectory on the fake Drive endpoint failed (hr=0x{:08X}).{}", static_cast<unsigned long>(createHr), diagnostics));
            return true;
        }
        const HRESULT renameHr =
            drive->RenameItem(drivePath(L"source/payload-01.bin").c_str(), drivePath(L"source/renamed.bin").c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        if (FAILED(renameHr) || driveExists(L"source/payload-01.bin") || ! driveExists(L"source/renamed.bin"))
        {
            const std::wstring diagnostics = fixtureLog();
            stopFakeDrive();
            Fail(std::format(L"R0f-GDrive: provider Rename on the fake Drive endpoint failed (hr=0x{:08X}).{}", static_cast<unsigned long>(renameHr), diagnostics));
            return true;
        }
        // Native Move: one parent change re-homes the object without copying bytes.
        const HRESULT moveHr =
            drive->MoveItem(drivePath(L"source/renamed.bin").c_str(), drivePath(L"created/moved.bin").c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        if (FAILED(moveHr) || driveExists(L"source/renamed.bin") || ! driveExists(L"created/moved.bin"))
        {
            const std::wstring diagnostics = fixtureLog();
            stopFakeDrive();
            Fail(std::format(L"R0f-GDrive: provider native Move on the fake Drive endpoint failed (hr=0x{:08X}).{}", static_cast<unsigned long>(moveHr), diagnostics));
            return true;
        }

        // R3-1: replace through the provider's atomic-final writer (this route binds no destination
        // objects): the source payload changes, the Copy collides, the Exists prompt offers
        // Overwrite, and the replaced object is copied back byte-exact.
        if (! WriteFilledTestFile(localSource / L"payload-00.bin", kFileBytes, 0x7Au))
        {
            stopFakeDrive();
            Fail(L"R0f-GDrive: failed to change the local payload for the replace step.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {localSource},
                                                 std::filesystem::path(driveRoot),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 drive);
        if (! state.taskA.has_value())
        {
            stopFakeDrive();
            Fail(L"R0f-GDrive: failed to start the replacing Copy to the fake Drive endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 3u;
        return false;
    }

    if (state.stepState == 3u)
    {
        using Task = FolderWindow::FileOperationState::Task;
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    stopFakeDrive();
                    Fail(L"R0f-GDrive: the Exists prompt on the identity-less Drive route must offer Overwrite.");
                    return true;
                }
                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, true);
                return false;
            }
        }
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fGDrive.ReplaceOnDriveMs", nowTick - phaseStartTick, completed->second.hr);
        if (FAILED(completed->second.hr) || completed->second.conflictPromptCount == 0u)
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeDrive();
            Fail(std::format(L"R0f-GDrive: replacing through the atomic-final writer failed (hr=0x{:08X}, prompts={}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             completed->second.conflictPromptCount,
                             diagnostics));
            return true;
        }
        const std::filesystem::path localBack2 = localBase / L"copied-back-2";
        if (! RecreateEmptyDirectory(localBack2))
        {
            stopFakeDrive();
            Fail(L"R0f-GDrive: failed to reset the second copy-back folder.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 drive,
                                                 {std::filesystem::path(drivePath(L"source"))},
                                                 localBack2,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            stopFakeDrive();
            Fail(L"R0f-GDrive: failed to start the second Copy back from the fake Drive endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 4u;
        return false;
    }

    if (state.stepState == 4u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        const std::filesystem::path copied2 = localBase / L"copied-back-2" / L"source";
        if (FAILED(completed->second.hr) || ! FilesEqualBytes(localSource / L"payload-00.bin", copied2 / L"payload-00.bin") ||
            ! FilesEqualBytes(localSource / L"payload-02.bin", copied2 / L"payload-02.bin"))
        {
            const std::wstring diagnostics = taskDiagnostics() + fixtureLog();
            stopFakeDrive();
            Fail(std::format(L"R0f-GDrive: the replaced object did not come back with the new bytes (hr=0x{:08X}).{}",
                             static_cast<unsigned long>(completed->second.hr),
                             diagnostics));
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Right,
                                                 std::nullopt,
                                                 drive,
                                                 {std::filesystem::path(drivePath(L"source")), std::filesystem::path(drivePath(L"created"))},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false);
        if (! state.taskA.has_value())
        {
            stopFakeDrive();
            Fail(L"R0f-GDrive: failed to start the Delete on the fake Drive endpoint.");
            return true;
        }
        phaseStartTick  = nowTick;
        state.stepState = 5u;
        return false;
    }

    if (state.stepState == 10u)
    {
        using Task = FolderWindow::FileOperationState::Task;
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (TryGetConflictPromptCopy(task).has_value())
            {
                // Drain the RED fixture safely before reporting the unexpected replay/skip surface.
                task->SubmitConflictDecision(Task::ConflictAction::Cancel, false);
            }
        }
        const auto failedCopy = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (failedCopy == state.completedTasks.end())
        {
            return false;
        }
        unsigned int requests     = 0u;
        unsigned int commits      = 0u;
        const HRESULT snapshotHr  = mutationCommitFailure(FALSE, requests, commits);
        const bool backendChanged =
            committedDeleteWitness ? ! driveExists(L"source/payload-00.bin") : driveExists(L"payload-00.bin") && driveExists(L"source/payload-00.bin");
        const auto& result  = failedCopy->second;
        const bool truthful = result.sourceItemResults.size() == 1u && result.sourceItemResults[0].has_value() &&
                              result.sourceItemResults[0].value().completion == FileOperations::ItemCompletion::Indeterminate &&
                              result.sourceItemResults[0].value().publication ==
                                  (committedDeleteWitness ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::Unknown) &&
                              result.sourceItemResults[0].value().sourceDisposition ==
                                  (committedDeleteWitness ? FileOperations::SourceDisposition::Unknown : FileOperations::SourceDisposition::Retained);
        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        const auto summary    = std::ranges::find_if(summaries, [&](const auto& item) { return item.taskId == state.taskA.value(); });
        const bool humanTruth = summary != summaries.end() && summary->indeterminateItemCount == 1u &&
                                summary->resultSummary == LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN);
        Debug::Perf::Emit(committedDeleteWitness ? L"FileOps.SelfTest.C0GDrive.CommittedDeleteResponseLost"
                                                 : L"FileOps.SelfTest.C0GDrive.CommittedCopyResponseLost",
                          committedDeleteWitness ? L"native-permanent-delete" : L"native-copy",
                          (nowTick - phaseStartTick) * 1000u,
                          requests,
                          commits,
                          result.hr);
        const bool passed         = SUCCEEDED(snapshotHr) && requests == 1u && commits == 1u && backendChanged && truthful && humanTruth &&
                                    result.hr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) && result.conflictPromptCount == 0u;
        const std::wstring detail = passed ? std::wstring{}
                                           : std::format(L"hr=0x{:08X}, requests={}, commits={}, backendChanged={}, truthful={}, humanTruth={}, prompts={}.{}",
                                                         static_cast<unsigned long>(result.hr),
                                                         requests,
                                                         commits,
                                                         backendChanged,
                                                         truthful,
                                                         humanTruth,
                                                         result.conflictPromptCount,
                                                         taskDiagnostics());
        stopFakeDrive();
        driveModule.reset();
        if (! passed)
        {
            Fail(L"C0-GDrive: committed mutation must finish uncertain without replay or conflict actions: " + detail);
            return true;
        }
        static_cast<void>(SelfTest::RemoveAll(localBase));
        NextStep(state, committedDeleteWitness ? SelfTestState::Step::C0_NativeCopyFailurePreservesKnownAxes
                                              : SelfTestState::Step::C0_GDriveCommittedDeleteResponseLost);
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    Debug::Perf::EmitValue(L"FileOps.SelfTest.R0fGDrive.DeleteOnDriveMs", nowTick - phaseStartTick, completed->second.hr);
    const bool sourceRemains  = driveExists(L"source");
    const bool createdRemains = driveExists(L"created");
    const bool movedRemains   = driveExists(L"created/moved.bin");
    const bool deleted        = SUCCEEDED(completed->second.hr) && ! sourceRemains && ! createdRemains && ! movedRemains;
    std::wstring deleteDetail;
    if (! deleted)
    {
        for (const std::optional<FileOperations::FileOperationItemResult>& item : completed->second.sourceItemResults)
        {
            if (item.has_value())
            {
                deleteDetail += std::format(L" [item status=0x{:08X} completion={} publication={} source={} path={}]",
                                            static_cast<unsigned long>(item->status),
                                            static_cast<unsigned int>(item->completion),
                                            static_cast<unsigned int>(item->publication),
                                            static_cast<unsigned int>(item->sourceDisposition),
                                            item->finalSourcePath);
            }
        }
        deleteDetail += taskDiagnostics();
        deleteDetail += fixtureLog();
        deleteDetail += std::format(L" remains: source={} created={} moved={}", sourceRemains, createdRemains, movedRemains);
    }
    // C10: the provider contract, while the fixture is still up.
    const HRESULT c10Hr = deleted ? c10DriveContract() : S_OK;
    if (FAILED(c10Hr))
    {
        deleteDetail += fixtureLog();
    }
    stopFakeDrive();
    driveModule.reset();
    if (! deleted)
    {
        Fail(std::format(L"R0f-GDrive: Delete on the fake Drive endpoint failed or left items behind (hr=0x{:08X}).{}",
                         static_cast<unsigned long>(completed->second.hr),
                         deleteDetail));
        return true;
    }
    if (FAILED(c10Hr))
    {
        Fail(std::format(L"C10-GDrive: the identity-delete contract must pin the file id and refuse a name that gained another file (hr=0x{:08X}).{}",
                         static_cast<unsigned long>(c10Hr),
                         deleteDetail));
        return true;
    }
    static_cast<void>(SelfTest::RemoveAll(localBase));
    NextStep(state, SelfTestState::Step::C0_GDriveCommittedCopyResponseLost);
    return false;
}

case SelfTestState::Step::C0_NativeCopyFailurePreservesKnownAxes:
{
    // Reuse real Local publication fault hooks through the host. Unknown stage cleanup
    // must retain known non-publication; a post-publication attributes error stays published.
    using Task              = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        Fail(L"C0 native Copy receipt test timed out.");
        return true;
    }
    const bool committed                    = state.stepState >= 3u;
    const unsigned int phase                = state.stepState % 3u;
    const std::filesystem::path root        = state.tempRoot / L"c0-native-receipt";
    const std::filesystem::path source      = root / L"source" / L"payload.bin";
    const std::filesystem::path destination = root / L"destination" / L"payload.bin";
    constexpr std::string_view original     = "original destination retained";
    constexpr std::string_view replacement  = "new copied contents";
    wil::com_ptr<IFileSystemIO> io;
    if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(io.addressof()))) || ! io)
    {
        Fail(L"C0 native Copy receipt test requires Local I/O.");
        return true;
    }
    if (phase == 0u)
    {
        if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(source.parent_path()) || ! RecreateEmptyDirectory(destination.parent_path()) ||
            ! WriteFileTextFsIo(io, source, replacement) || ! WriteFileTextFsIo(io, destination, original))
        {
            Fail(L"C0 native Copy receipt test failed to seed its owned fixture.");
            return true;
        }
        const bool armed = committed ? SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailPath.data(), destination.c_str()) != FALSE
                                     : SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailPath.data(), destination.c_str()) != FALSE &&
                                           SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownPath.data(), destination.c_str()) != FALSE;
        if (! armed)
        {
            Fail(L"C0 native Copy receipt test failed to arm the existing Local fault hooks.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {source},
                                                 destination.parent_path(),
                                                 FILESYSTEM_FLAG_NONE,
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"C0 native Copy receipt test failed to start its task.");
            return true;
        }
        ++state.stepState;
        return false;
    }
    Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
    {
        if (phase == 1u && prompt->bucket == Task::ConflictBucket::RegularFileExists && PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
        {
            task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
            ++state.stepState;
            return false;
        }
        // The first published prompt can remain visible until the worker consumes our
        // decision. Do not replace that queued Overwrite with Cancel during its drain.
        if (phase != 2u || prompt->bucket != Task::ConflictBucket::RegularFileExists)
        {
            // Drain an unexpected post-failure decision, then fail on its count/truth.
            task->SubmitConflictDecision(Task::ConflictAction::Cancel, false);
        }
    }
    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    const auto& result           = completed->second;
    const HRESULT expectedStatus = HRESULT_FROM_WIN32(committed ? ERROR_ACCESS_DENIED : ERROR_IO_INCOMPLETE);
    const bool hookFired         = GetEnvVarTrimmed(committed ? kSelfTestEnvFinalAttributesFailFired : kSelfTestEnvAbortOwnedStageUnknownFired) == L"1";
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvFinalAttributesFailFired.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvStagedCopyPromoteFailFired.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownPath.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvAbortOwnedStageUnknownFired.data(), nullptr));
    std::string sourceBytes;
    std::string destinationBytes;
    const bool bytesCorrect = ReadFileTextFsIo(io, source, sourceBytes) && sourceBytes == replacement && ReadFileTextFsIo(io, destination, destinationBytes) &&
                              destinationBytes == (committed ? replacement : original);
    const bool truthful =
        result.sourceItemResults.size() == 1u && result.sourceItemResults[0].has_value() &&
        result.sourceItemResults[0]->publication ==
            (committed ? FileOperations::PublicationState::Published : FileOperations::PublicationState::NotPublished) &&
        result.sourceItemResults[0]->sourceDisposition == FileOperations::SourceDisposition::Retained &&
        result.sourceItemResults[0]->completion == (committed ? FileOperations::ItemCompletion::Failed : FileOperations::ItemCompletion::Indeterminate) &&
        result.sourceItemResults[0]->ownedStageDisposition ==
            (committed ? FileOperations::OwnedStageDisposition::Published : FileOperations::OwnedStageDisposition::Unknown);
    const bool passed = hookFired && bytesCorrect && truthful && result.hr == expectedStatus && result.conflictPromptCount == 1u;
    Debug::Perf::EmitValue(committed ? L"FileOps.SelfTest.C0Native.PublishedAttributesFailurePromptCount"
                                     : L"FileOps.SelfTest.C0Native.UnknownCleanupPromptCount",
                           result.conflictPromptCount,
                           passed ? S_OK : E_FAIL);
    if (! passed)
    {
        Fail(std::format(L"C0 native Copy lost independent receipt truth (committed={}, hook={}, bytes={}, truth={}, prompts={}, hr=0x{:08X}).",
                         committed,
                         hookFired,
                         bytesCorrect,
                         truthful,
                         result.conflictPromptCount,
                         static_cast<unsigned long>(result.hr)));
        return true;
    }
    if (! committed)
    {
        state.taskA.reset();
        state.stepState = 3u;
        return false;
    }
    static_cast<void>(SelfTest::RemoveAll(root));
    NextStep(state, SelfTestState::Step::C0_MutationReceiptPrefixBoundary);
    return false;
}

case SelfTestState::Step::C0_MutationReceiptPrefixBoundary:
{
    using Task           = FolderWindow::FileOperationState::Task;
    using Classification = FileSystemRouteContract::MutationClassification;
    const auto started   = std::chrono::steady_clock::now();
    SYSTEM_INFO systemInfo{};
    GetSystemInfo(&systemInfo);
    const size_t pageBytes = systemInfo.dwPageSize;
    wil::unique_virtualalloc_ptr<std::byte> allocation(static_cast<std::byte*>(VirtualAlloc(nullptr, pageBytes * 2u, MEM_RESERVE, PAGE_NOACCESS)));
    if (! allocation || VirtualAlloc(allocation.get(), pageBytes, MEM_COMMIT, PAGE_READWRITE) == nullptr)
    {
        Fail(L"C0 receipt boundary could not allocate the owned readable page and inaccessible suffix.");
        return true;
    }

    // Only the size field is accessible. Rejection must short-circuit before any
    // Boolean or stage field is read, including partially advertised extensions.
    std::byte* const headerBytes = allocation.get() + pageBytes - sizeof(uint32_t);
    for (const uint32_t badSize : {0u, 4u, 15u, 17u, 19u})
    {
        std::memcpy(headerBytes, &badSize, sizeof(badSize));
        const auto* header = reinterpret_cast<const FileSystemItemMutationResult*>(headerBytes);
        if (FileSystemRouteContract::SnapshotItemMutationResult(header).has_value() ||
            FileSystemRouteContract::ClassifyFailedMutation(true, E_FAIL, header) != Classification::ContractViolation)
        {
            Fail(L"C0 receipt boundary accepted an unsupported size.");
            return true;
        }
    }

    const FileSystemItemMutationResult v1{FILESYSTEM_ITEM_MUTATION_RESULT_V1_SIZE, TRUE, FALSE, TRUE};
    std::byte* const v1Bytes = allocation.get() + pageBytes - FILESYSTEM_ITEM_MUTATION_RESULT_V1_SIZE;
    std::memcpy(v1Bytes, &v1, FILESYSTEM_ITEM_MUTATION_RESULT_V1_SIZE);
    const auto* v1Receipt = reinterpret_cast<const FileSystemItemMutationResult*>(v1Bytes);
    Task task(*state.fileOps);
    task._operation     = FILESYSTEM_COPY;
    task._executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
    task._sourcePaths   = {state.tempRoot / L"receipt-prefix-no-io.bin"};
    task.InitializeSourceItemResultBuilders();
    task.MarkSourceItemsMutationPossible();
    // This synchronous fixture is the callback owner; no provider worker runs.
    task._dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
    const auto callbackScope = wil::scope_exit([&] noexcept { task._dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
    const HRESULT v1Hr       = task.FileSystemItemCompleted(FILESYSTEM_COPY, 0u, nullptr, nullptr, E_ACCESSDENIED, v1Receipt, nullptr, nullptr);
    auto& builder            = task._sourceItemResultBuilders.front();
    if (v1Hr != S_OK || ! builder.mutation.has_value() || builder.mutation.value().sizeBytes != FILESYSTEM_ITEM_MUTATION_RESULT_V1_SIZE ||
        builder.mutation.value().ownedStageDisposition != FileSystemOwnedStageDisposition::NotApplicable ||
        FileSystemRouteContract::ClassifyFailedMutation(true, E_ACCESSDENIED, &builder.mutation.value()) != Classification::RetryableNoCommit)
    {
        Fail(L"C0 receipt boundary did not safely snapshot the exact V1 provider buffer.");
        return true;
    }

    FileSystemItemMutationResult full{sizeof(FileSystemItemMutationResult), TRUE, TRUE, TRUE, FileSystemOwnedStageDisposition::Published};
    const HRESULT fullHr = task.FileSystemItemCompleted(FILESYSTEM_COPY, 0u, nullptr, nullptr, S_OK, &full, nullptr, nullptr);
    const bool fullCopied =
        fullHr == S_OK && builder.mutation.has_value() && builder.mutation.value().ownedStageDisposition == FileSystemOwnedStageDisposition::Published;
    // A missing receipt must clear the observation, not retain earlier proof.
    const HRESULT nullHr = task.FileSystemItemCompleted(FILESYSTEM_COPY, 0u, nullptr, nullptr, E_FAIL, nullptr, nullptr, nullptr);
    if (nullHr != S_OK || builder.mutation.has_value() || FileSystemRouteContract::SnapshotItemMutationResult(nullptr).has_value())
    {
        Fail(L"C0 missing completion inherited an earlier receipt.");
        return true;
    }
    static_cast<void>(task.FileSystemItemCompleted(FILESYSTEM_COPY, 0u, nullptr, nullptr, S_OK, &full, nullptr, nullptr));
    full.outcomeKnown        = 2;
    const HRESULT invalidHr  = task.FileSystemItemCompleted(FILESYSTEM_COPY, 0u, nullptr, nullptr, S_OK, &full, nullptr, nullptr);
    const HRESULT terminalHr = task.FinalizeTypedItemResults(S_OK);
    if (! fullCopied || invalidHr != E_INVALIDARG || terminalHr != HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) || ! builder.terminal.has_value() ||
        builder.terminal.value().publication != FileOperations::PublicationState::Unknown ||
        builder.terminal.value().completion != FileOperations::ItemCompletion::Indeterminate)
    {
        Fail(L"C0 malformed completion inherited prior publication or became successful when the provider ignored E_INVALIDARG.");
        return true;
    }

    full.outcomeKnown          = TRUE;
    full.sizeBytes             = static_cast<uint32_t>(sizeof(full)) + 4u;
    const auto future          = FileSystemRouteContract::SnapshotItemMutationResult(&full);
    full.sizeBytes             = sizeof(full);
    full.ownedStageDisposition = static_cast<FileSystemOwnedStageDisposition>(999u);
    if (! future.has_value() || future.value().sizeBytes != sizeof(full) ||
        future.value().ownedStageDisposition != FileSystemOwnedStageDisposition::Published ||
        FileSystemRouteContract::SnapshotItemMutationResult(&full).has_value())
    {
        Fail(L"C0 receipt boundary mishandled an unknown suffix or invalid stage value.");
        return true;
    }
    const uint64_t durationUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());
    Debug::Perf::Emit(L"FileOps.SelfTest.C0MutationReceipt.PrefixBoundary", L"v1-guarded/full/malformed/null/future", durationUs, 5u, 5u, S_OK);
    NextStep(state, SelfTestState::Step::C0_LocalKnownNoCommitReceipts);
    return false;
}

case SelfTestState::Step::C0_LocalKnownNoCommitReceipts:
{
    // The host classifies every failed provider call through its receipt. The Local provider
    // must therefore prove "nothing was committed" for failures that never reached a stage
    // (a Cancel at the conflict prompt, a refused native rename) and "the destination changed"
    // for a canceled merge whose earlier children were published. None of these is uncertain.
    using Task              = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 90'000ull))
    {
        Fail(L"C0 Local known no-commit receipt test timed out.");
        return true;
    }
    enum class Scenario : unsigned int
    {
        FileCancelAtPrompt,
        MergeCancelAfterChildPublished,
        LockedNativeMove,
    };
    const auto scenario                     = static_cast<Scenario>(state.stepState / 2u);
    const bool starting                     = (state.stepState % 2u) == 0u;
    const std::filesystem::path root        = state.tempRoot / L"c0-local-receipts";
    const std::filesystem::path sourceDir   = root / L"source";
    const std::filesystem::path destination = root / L"destination";
    constexpr std::string_view sourceText   = "copied payload";
    constexpr std::string_view occupantText = "existing destination occupant";
    wil::com_ptr<IFileSystemIO> io;
    if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(io.addressof()))) || ! io)
    {
        Fail(L"C0 Local receipt test requires Local I/O.");
        return true;
    }
    const auto scenarioName = [scenario]() noexcept -> std::wstring_view
    {
        switch (scenario)
        {
            case Scenario::FileCancelAtPrompt: return L"file-cancel-at-prompt";
            case Scenario::MergeCancelAfterChildPublished: return L"merge-cancel-after-child";
            case Scenario::LockedNativeMove:
            default: return L"locked-native-move";
        }
    };
    if (starting)
    {
        state.c0LocalLockedHandle.reset();
        bool seeded = RecreateEmptyDirectory(root) && RecreateEmptyDirectory(sourceDir) && RecreateEmptyDirectory(destination);
        std::vector<std::filesystem::path> sources;
        FileSystemOperation operation = FILESYSTEM_COPY;
        switch (scenario)
        {
            case Scenario::FileCancelAtPrompt:
                seeded = seeded && WriteFileTextFsIo(io, sourceDir / L"a.bin", sourceText) && WriteFileTextFsIo(io, destination / L"a.bin", occupantText);
                sources.push_back(sourceDir / L"a.bin");
                break;
            case Scenario::MergeCancelAfterChildPublished:
                // Both Local walkers publish a child directory on the traversal thread, in
                // enumeration order, before any later file is even queued. "a-sub" therefore
                // exists in the destination before "z.bin" raises its conflict, serial or parallel.
                seeded = seeded && RecreateEmptyDirectory(sourceDir / L"tree") && RecreateEmptyDirectory(sourceDir / L"tree" / L"a-sub") &&
                         RecreateEmptyDirectory(destination / L"tree") && WriteFileTextFsIo(io, sourceDir / L"tree" / L"z.bin", sourceText) &&
                         WriteFileTextFsIo(io, destination / L"tree" / L"z.bin", occupantText);
                sources.push_back(sourceDir / L"tree");
                break;
            case Scenario::LockedNativeMove:
            default:
                seeded = seeded && WriteFileTextFsIo(io, sourceDir / L"m.bin", sourceText);
                if (seeded)
                {
                    // No FILE_SHARE_DELETE: the same-volume rename is refused before it can commit.
                    state.c0LocalLockedHandle.reset(
                        CreateFileW((sourceDir / L"m.bin").c_str(), GENERIC_READ, 0u, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
                    seeded = static_cast<bool>(state.c0LocalLockedHandle);
                }
                sources.push_back(sourceDir / L"m.bin");
                operation = FILESYSTEM_MOVE;
                break;
        }
        if (! seeded)
        {
            Fail(std::format(L"C0 Local receipt test failed to seed '{}'.", scenarioName()));
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 operation,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 sources,
                                                 destination,
                                                 scenario == Scenario::MergeCancelAfterChildPublished ? FILESYSTEM_FLAG_RECURSIVE : FILESYSTEM_FLAG_NONE,
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(std::format(L"C0 Local receipt test failed to start '{}'.", scenarioName()));
            return true;
        }
        ++state.stepState;
        return false;
    }

    Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
    {
        if (scenario == Scenario::LockedNativeMove)
        {
            // A refused native rename is a proved no-commit: the host must still offer Retry.
            if (! PromptHasAction(prompt.value(), Task::ConflictAction::Retry))
            {
                Fail(L"C0 Local receipt: the refused native Move prompt does not offer Retry.");
                return true;
            }
            task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
        }
        else
        {
            if (prompt->bucket != Task::ConflictBucket::RegularFileExists)
            {
                Fail(std::format(L"C0 Local receipt '{}' expected a RegularFileExists prompt.", scenarioName()));
                return true;
            }
            task->SubmitConflictDecision(Task::ConflictAction::Cancel, false);
        }
        return false;
    }
    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    state.c0LocalLockedHandle.reset();
    const auto& result = completed->second;
    std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> diagnostics;
    state.fileOps->CollectDiagnostics(diagnostics);
    const bool indeterminateLogged = std::ranges::any_of(diagnostics,
                                                         [&](const auto& entry) noexcept
    { return entry.taskId == state.taskA.value() && entry.category == L"item.mutation.indeterminate"; });
    std::string bytes;
    bool diskCorrect = false;
    bool truthful    = false;
    const bool haveItem = result.sourceItemResults.size() == 1u && result.sourceItemResults[0].has_value();
    const FileOperations::FileOperationItemResult item = haveItem ? result.sourceItemResults[0].value() : FileOperations::FileOperationItemResult{};
    switch (scenario)
    {
        case Scenario::FileCancelAtPrompt:
            diskCorrect = ReadFileTextFsIo(io, destination / L"a.bin", bytes) && bytes == occupantText && ReadFileTextFsIo(io, sourceDir / L"a.bin", bytes) &&
                          bytes == sourceText;
            truthful = haveItem && item.completion == FileOperations::ItemCompletion::Canceled &&
                       item.publication == FileOperations::PublicationState::NotPublished &&
                       item.sourceDisposition == FileOperations::SourceDisposition::Retained &&
                       item.ownedStageDisposition == FileOperations::OwnedStageDisposition::NotCreated && result.conflictPromptCount == 1u;
            break;
        case Scenario::MergeCancelAfterChildPublished:
        {
            std::error_code ec;
            diskCorrect = std::filesystem::is_directory(destination / L"tree" / L"a-sub", ec) && ! ec &&
                          ReadFileTextFsIo(io, destination / L"tree" / L"z.bin", bytes) && bytes == occupantText;
            truthful = haveItem && item.completion == FileOperations::ItemCompletion::Canceled &&
                       item.publication == FileOperations::PublicationState::Published &&
                       item.sourceDisposition == FileOperations::SourceDisposition::Retained && result.conflictPromptCount == 1u;
            break;
        }
        case Scenario::LockedNativeMove:
        default:
        {
            std::error_code ec;
            diskCorrect = ReadFileTextFsIo(io, sourceDir / L"m.bin", bytes) && bytes == sourceText && ! std::filesystem::exists(destination / L"m.bin", ec) && ! ec;
            truthful    = haveItem && item.completion == FileOperations::ItemCompletion::Skipped &&
                       item.publication == FileOperations::PublicationState::NotPublished &&
                       item.sourceDisposition == FileOperations::SourceDisposition::Retained && result.conflictPromptCount == 1u;
            break;
        }
    }
    const bool passed = diskCorrect && truthful && ! indeterminateLogged;
    Debug::Perf::Emit(L"FileOps.SelfTest.C0LocalReceipt",
                      scenarioName().data(),
                      0u,
                      result.conflictPromptCount,
                      haveItem ? static_cast<uint64_t>(item.completion) : 99u,
                      passed ? S_OK : E_FAIL);
    if (! passed)
    {
        Fail(std::format(L"C0 Local receipt '{}' lost known truth (disk={}, truth={}, indeterminateLogged={}, completion={}, publication={}, stage={}, "
                         L"prompts={}, hr=0x{:08X}).",
                         scenarioName(),
                         diskCorrect,
                         truthful,
                         indeterminateLogged,
                         haveItem ? static_cast<unsigned int>(item.completion) : 99u,
                         haveItem ? static_cast<unsigned int>(item.publication) : 99u,
                         haveItem ? static_cast<unsigned int>(item.ownedStageDisposition) : 99u,
                         result.conflictPromptCount,
                         static_cast<unsigned long>(result.hr)));
        return true;
    }
    state.taskA.reset();
    if (scenario != Scenario::LockedNativeMove)
    {
        ++state.stepState;
        return false;
    }
    static_cast<void>(SelfTest::RemoveAll(root));
    NextStep(state, SelfTestState::Step::C0_CurlNativeDeleteGuards);
    return false;
}

case SelfTestState::Step::C0_CurlNativeDeleteGuards:
case SelfTestState::Step::C0_CurlCommittedMutationResponseLost:
case SelfTestState::Step::C0_CurlPartialTreeFailure:
case SelfTestState::Step::C0_CurlNativeDeleteLateListing:
case SelfTestState::Step::C0_CurlDeleteTraversalBounds:
case SelfTestState::Step::C0_CurlDirectorySizeTruth:
case SelfTestState::Step::C0_CurlMovePreflightTruth:
case SelfTestState::Step::C0_CurlCopyTraversalTruth:
case SelfTestState::Step::C0_CurlEntryLookupTruth:
case SelfTestState::Step::C0_CurlImapListingTruth:
case SelfTestState::Step::C0_CurlImapTransportTruth:
{
    const bool imapTransport   = state.step == SelfTestState::Step::C0_CurlImapTransportTruth;
    const bool imapListing     = state.step == SelfTestState::Step::C0_CurlImapListingTruth;
    const bool entryLookup     = state.step == SelfTestState::Step::C0_CurlEntryLookupTruth;
    const bool copyTraversal   = state.step == SelfTestState::Step::C0_CurlCopyTraversalTruth;
    const bool replyLoss       = state.step == SelfTestState::Step::C0_CurlCommittedMutationResponseLost;
    const bool partialTree     = state.step == SelfTestState::Step::C0_CurlPartialTreeFailure;
    const bool lateListing     = state.step == SelfTestState::Step::C0_CurlNativeDeleteLateListing;
    const bool traversalBounds = state.step == SelfTestState::Step::C0_CurlDeleteTraversalBounds;
    const bool directorySize   = state.step == SelfTestState::Step::C0_CurlDirectorySizeTruth;
    const bool movePreflight   = state.step == SelfTestState::Step::C0_CurlMovePreflightTruth;
    if (state.stepState == 0u)
    {
        FARPROC address      = nullptr;
        const HRESULT loadHr = SelfTest::LoadPluginSelfTestExport(kPluginIdFtp,
                                                                  imapTransport     ? "RedSalamanderCurlImapTransportTruthForSelfTest"
                                                                  : imapListing     ? "RedSalamanderCurlImapListingTruthForSelfTest"
                                                                  : entryLookup     ? "RedSalamanderCurlEntryLookupTruthForSelfTest"
                                                                  : copyTraversal   ? "RedSalamanderCurlCopyTraversalTruthForSelfTest"
                                                                  : movePreflight   ? "RedSalamanderCurlMovePreflightTruthForSelfTest"
                                                                  : directorySize   ? "RedSalamanderCurlDirectorySizeTruthForSelfTest"
                                                                  : traversalBounds ? "RedSalamanderCurlDeleteTraversalBoundsForSelfTest"
                                                                  : lateListing     ? "RedSalamanderCurlNativeDeleteLateListingForSelfTest"
                                                                  : partialTree     ? "RedSalamanderCurlPartialTreeFailureForSelfTest"
                                                                  : replyLoss       ? "RedSalamanderCurlLostMutationReplyForSelfTest"
                                                                                    : "RedSalamanderCurlNativeDeleteGuardsForSelfTest",
                                                                  state.c0CurlModule,
                                                                  address);
        if (FAILED(loadHr) || address == nullptr)
        {
            state.c0CurlModule.reset();
            Fail(L"C0 Curl native safety test export is unavailable.");
            return true;
        }
        using RunFn = HRESULT(__stdcall*)(unsigned int*, unsigned int*) noexcept;
#pragma warning(push)
#pragma warning(disable : 4191) // Test-only, exact exported C ABI signature.
        const auto run = reinterpret_cast<RunFn>(address);
#pragma warning(pop)
        state.c0CurlPassed = 0u;
        state.c0CurlFailed = 0u;
        state.c0CurlResult.store(E_PENDING, std::memory_order_release);
        state.c0CurlWorker =
            std::jthread([&state, run]() noexcept { state.c0CurlResult.store(run(&state.c0CurlPassed, &state.c0CurlFailed), std::memory_order_release); });
        ++state.stepState;
        return false;
    }
    const HRESULT hr = state.c0CurlResult.load(std::memory_order_acquire);
    if (hr == E_PENDING)
    {
        // The copy-traversal export alone runs its Wide matrix (4,096 files, eight variants) in about
        // four minutes on loopback; the budget covers the whole export, not one scenario.
        if (HasTimedOut(state, GetTickCount64(), 480'000ull))
        {
            Fail(L"C0 Curl native safety worker timed out.");
            return true;
        }
        return false;
    }
    state.c0CurlWorker.join();
    state.c0CurlModule.reset();
    const unsigned int expectedChecks = imapTransport     ? 15u
                                        : imapListing     ? 369u
                                        : entryLookup     ? 298u
                                        : copyTraversal   ? 142u
                                        : movePreflight   ? 256u
                                        : directorySize   ? 85u
                                        : traversalBounds ? 127u
                                        : lateListing     ? 49u
                                        : partialTree     ? 37u
                                        : replyLoss       ? 33u
                                                          : 29u;
    AppendLog(std::format(L"C0 Curl {}: {} passed, {} failed, hr=0x{:08X}; expected {} checks.",
                          imapTransport     ? L"IMAP transport truth"
                          : imapListing     ? L"IMAP listing truth"
                          : entryLookup     ? L"entry lookup truth"
                          : copyTraversal   ? L"Copy traversal truth"
                          : movePreflight   ? L"Move preflight truth"
                          : directorySize   ? L"directory-size truth"
                          : traversalBounds ? L"native Delete traversal bounds"
                          : lateListing     ? L"native Delete late listing"
                          : partialTree     ? L"partial-tree failure"
                          : replyLoss       ? L"lost mutation reply"
                                            : L"native Delete guards",
                          state.c0CurlPassed,
                          state.c0CurlFailed,
                          static_cast<unsigned long>(hr),
                          expectedChecks));
    if (FAILED(hr) || state.c0CurlFailed != 0u || state.c0CurlPassed != expectedChecks)
    {
        Fail(std::format(L"C0 Curl native safety test failed: {} of {} checks passed, {} failed, hr=0x{:08X}. See FileOps.Curl metrics for the exact route.",
                         state.c0CurlPassed,
                         expectedChecks,
                         state.c0CurlFailed,
                         static_cast<unsigned long>(hr)));
        return true;
    }
    NextStep(state,
             imapTransport     ? SelfTestState::Step::R3_1_IdentityLessReplaceDummy
             : imapListing     ? SelfTestState::Step::C0_CurlImapTransportTruth
             : entryLookup     ? SelfTestState::Step::C0_CurlImapListingTruth
             : copyTraversal   ? SelfTestState::Step::C0_CurlEntryLookupTruth
             : movePreflight   ? SelfTestState::Step::C0_CurlCopyTraversalTruth
             : directorySize   ? SelfTestState::Step::C0_CurlMovePreflightTruth
             : traversalBounds ? SelfTestState::Step::C0_CurlDirectorySizeTruth
             : lateListing     ? SelfTestState::Step::C0_CurlDeleteTraversalBounds
             : partialTree     ? SelfTestState::Step::C0_CurlHostPartialTreeFailure
             : replyLoss       ? SelfTestState::Step::C0_CurlHostCommittedDeleteResponseLost
                               : SelfTestState::Step::C0_CurlCommittedMutationResponseLost);
    return false;
}

case SelfTestState::Step::R3_1_IdentityLessReplaceDummy:
{
    // R3-1: a destination without object binding (Dummy) replaces one occupant conditionally through
    // its atomic-final writer. (1) The Exists prompt offers Overwrite and the replacement lands with
    // the source bytes. (2) When the occupant changes while the decision is pending, nothing is
    // replaced: the item fails with ERROR_REVISION_MISMATCH and the raced occupant survives.
    using Task                          = FolderWindow::FileOperationState::Task;
    constexpr size_t kFileBytes         = 16u * 1024u;
    const std::filesystem::path srcRoot = state.tempRoot / L"r3-1-replace-src";
    const std::filesystem::path srcFile = srcRoot / L"replace.bin";
    const std::wstring dummyRoot        = L"/r3-1-identity-less-replace";
    const std::wstring dummyFile        = dummyRoot + L"/replace.bin";
    const std::string occupantBytes(kFileBytes, static_cast<char>(0xA7));
    const std::string sourceBytes(kFileBytes, static_cast<char>(0x31));
    const std::string racedBytes(kFileBytes / 2u, static_cast<char>(0x5C));
    const auto dummyBytesEqual = [&](std::string_view expected) noexcept -> bool
    {
        wil::com_ptr<IFileSystemIO> io;
        wil::com_ptr<IFileReader> reader;
        if (! state.fsDummy || FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(io.addressof()))) || ! io ||
            FAILED(io->CreateFileReader(dummyFile.c_str(), reader.addressof())) || ! reader)
        {
            return false;
        }
        uint64_t sizeBytes = 0u;
        if (FAILED(reader->GetSize(&sizeBytes)) || sizeBytes != expected.size())
        {
            return false;
        }
        std::vector<char> actual(expected.size());
        unsigned long bytesRead = 0u;
        return expected.empty() || (SUCCEEDED(reader->Read(actual.data(), static_cast<unsigned long>(actual.size()), &bytesRead)) &&
                                    bytesRead == actual.size() && std::string_view(actual.data(), actual.size()) == expected);
    };
    const auto startReplaceCopy = [&]() noexcept
    {
        return StartFileOperationAndGetId(state.fileOps,
                                          FILESYSTEM_COPY,
                                          FolderWindow::Pane::Left,
                                          FolderWindow::Pane::Right,
                                          state.fsLocal,
                                          {srcFile},
                                          std::filesystem::path(dummyRoot),
                                          FILESYSTEM_FLAG_NONE,
                                          false,
                                          0,
                                          FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                          false,
                                          state.fsDummy);
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        Fail(std::format(L"R3_1_IdentityLessReplaceDummy timed out at step {}.", state.stepState));
        return true;
    }

    if (state.stepState == 0u)
    {
        if (! state.fsDummy)
        {
            Fail(L"R3-1: the Dummy provider is required for the identity-less replace proof.");
            return true;
        }
        if (! RecreateEmptyDirectory(srcRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
        {
            Fail(L"R3-1: failed to reset the local source and the Dummy destination folder.");
            return true;
        }
        if (! WriteFilledTestFile(srcFile, kFileBytes, 0x31u) || ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, occupantBytes, true))
        {
            Fail(L"R3-1: failed to seed the source file and the Dummy occupant.");
            return true;
        }
        state.taskA = startReplaceCopy();
        if (! state.taskA.has_value())
        {
            Fail(L"R3-1: failed to start the Copy onto the Dummy occupant.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket != Task::ConflictBucket::RegularFileExists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    std::wstring actions;
                    for (size_t index = 0; index < prompt->actionCount; ++index)
                    {
                        actions += std::format(L"{} ", static_cast<int>(prompt->actions[index]));
                    }
                    std::wstring diagnostics;
                    std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> entries;
                    state.fileOps->CollectDiagnostics(entries);
                    for (const auto& entry : entries)
                    {
                        if (entry.taskId == state.taskA.value())
                        {
                            diagnostics += std::format(L" [{} 0x{:08X} {}]", entry.category, static_cast<unsigned long>(entry.status), entry.message);
                        }
                    }
                    Fail(std::format(L"R3-1: the Exists prompt on the identity-less Dummy route must offer Overwrite (bucket={} status=0x{:08X} actions=[{}] grants={}).{}",
                                     static_cast<int>(prompt->bucket),
                                     static_cast<unsigned long>(prompt->status),
                                     actions,
                                     task->_conflictAtomicReplaceGrantCount.load(std::memory_order_acquire),
                                     diagnostics));
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
        if (FAILED(completed->second.hr) || completed->second.conflictPromptCount != 1u || ! dummyBytesEqual(sourceBytes))
        {
            Fail(std::format(L"R3-1: replacing the Dummy occupant through the atomic-final writer failed (hr=0x{:08X}, prompts={}).",
                             static_cast<unsigned long>(completed->second.hr),
                             completed->second.conflictPromptCount));
            return true;
        }

        // (2) The occupant changes while the decision is pending.
        if (! DummyWriteTextFile(state.fsDummy.get(), dummyFile, occupantBytes, true))
        {
            Fail(L"R3-1: failed to reseed the Dummy occupant for the race proof.");
            return true;
        }
        state.taskA = startReplaceCopy();
        if (! state.taskA.has_value())
        {
            Fail(L"R3-1: failed to start the raced Copy onto the Dummy occupant.");
            return true;
        }
        state.stepState = 2u;
        return false;
    }

    if (state.stepState == 2u)
    {
        if (Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                if (prompt->bucket == Task::ConflictBucket::RegularFileExists)
                {
                    if (! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                    {
                        Fail(L"R3-1: the raced Exists prompt must still offer Overwrite.");
                        return true;
                    }
                    // Replace the occupant behind the prompt, then answer for the object that is gone.
                    // The bridge refuses, re-raises the collision on the current occupant (twice), and
                    // finally reports the refusal through the ordinary retryable conflict.
                    if (! DummyWriteTextFile(state.fsDummy.get(), dummyFile, racedBytes, true))
                    {
                        Fail(L"R3-1: failed to race the Dummy occupant behind the prompt.");
                        return true;
                    }
                    task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
                    return false;
                }
                if (prompt->status != HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) || ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
                {
                    Fail(std::format(L"R3-1: the refused replacement must surface as a retryable conflict carrying ERROR_REVISION_MISMATCH (bucket={} status=0x{:08X}).",
                                     static_cast<int>(prompt->bucket),
                                     static_cast<unsigned long>(prompt->status)));
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
        // Three Exists decisions (the first plus two re-raised collisions) and one retryable conflict.
        if (completed->second.conflictPromptCount != 4u)
        {
            Fail(std::format(L"R3-1: expected 3 Exists prompts plus 1 retryable conflict for the raced occupant, saw {} prompts (task hr=0x{:08X}).",
                             completed->second.conflictPromptCount,
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        if (! dummyBytesEqual(racedBytes))
        {
            Fail(L"R3-1: the raced occupant must survive untouched when the replacement is refused.");
            return true;
        }
        static_cast<void>(SelfTest::RemoveAll(srcRoot));
        NextStep(state, SelfTestState::Step::R3_2_WriterProofDummy);
        return false;
    }
    return false;
}

case SelfTestState::Step::R3_2_WriterProofDummy:
{
    // R3-2: a destination without object binding (Dummy) proves the content it published through
    // its writer's own SHA-256. (1) Verify On: the copy is Verified from the writer proof alone,
    // with no host readback. (2) Managed Move: the cleanup record survives publication and the
    // bound Local source is removed only after the proof matched.
    using Task                           = FolderWindow::FileOperationState::Task;
    constexpr size_t kFileBytes          = 64u * 1024u;
    const std::filesystem::path srcRoot  = state.tempRoot / L"r3-2-proof-src";
    const std::filesystem::path copyFile = srcRoot / L"verify.bin";
    const std::filesystem::path moveFile = srcRoot / L"move.bin";
    const std::wstring dummyRoot         = L"/r3-2-writer-proof";
    const std::string copyBytes(kFileBytes, static_cast<char>(0x6B));
    const std::string moveBytes(kFileBytes, static_cast<char>(0x2E));
    const auto restoreVerifySetting = [&]() noexcept
    {
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy =
            state.fileOperationsOriginal.has_value() ? state.fileOperationsOriginal->verifyAfterCopy : false;
    };
    const auto dummyBytesEqual = [&](const std::wstring& dummyFile, std::string_view expected) noexcept -> bool
    {
        wil::com_ptr<IFileSystemIO> io;
        wil::com_ptr<IFileReader> reader;
        if (! state.fsDummy || FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(io.addressof()))) || ! io ||
            FAILED(io->CreateFileReader(dummyFile.c_str(), reader.addressof())) || ! reader)
        {
            return false;
        }
        uint64_t sizeBytes = 0u;
        if (FAILED(reader->GetSize(&sizeBytes)) || sizeBytes != expected.size())
        {
            return false;
        }
        std::vector<char> actual(expected.size());
        unsigned long bytesRead = 0u;
        return expected.empty() || (SUCCEEDED(reader->Read(actual.data(), static_cast<unsigned long>(actual.size()), &bytesRead)) &&
                                    bytesRead == actual.size() && std::string_view(actual.data(), actual.size()) == expected);
    };
    const auto startTransfer = [&](FileSystemOperation operation, const std::filesystem::path& sourceFile) noexcept
    {
        return StartFileOperationAndGetId(state.fileOps,
                                          operation,
                                          FolderWindow::Pane::Left,
                                          FolderWindow::Pane::Right,
                                          state.fsLocal,
                                          {sourceFile},
                                          std::filesystem::path(dummyRoot),
                                          FILESYSTEM_FLAG_NONE,
                                          false,
                                          0,
                                          FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                          false,
                                          state.fsDummy);
    };
    const auto describe = [&](const CompletedTaskInfo& completed) -> std::wstring
    {
        const auto* const result = completed.sourceItemResults.empty() || ! completed.sourceItemResults.front().has_value()
                                       ? nullptr
                                       : &completed.sourceItemResults.front().value();
        return std::format(L"task hr=0x{:08X} verification={} disposition={} providerProofs={} hostReadbacks={} prompts={}",
                           static_cast<unsigned long>(completed.hr),
                           result != nullptr ? static_cast<unsigned>(result->verification) : 255u,
                           result != nullptr ? static_cast<unsigned>(result->sourceDisposition) : 255u,
                           completed.verificationProviderProofCount,
                           completed.verificationHostReadbackCount,
                           completed.conflictPromptCount);
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        restoreVerifySetting();
        Fail(std::format(L"R3_2_WriterProofDummy timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        if (! Common::Crypto::ContentDigestSelfCheck())
        {
            // Known-answer vectors (SHA-256 "abc", CRC-64/NVME "123456789"): a wrong digest here would
            // let the fixtures and the host agree on a value a real service never returns.
            Fail(L"R3-2: the content-digest library failed its known-answer self-check.");
            return true;
        }
        if (! state.fsDummy)
        {
            Fail(L"R3-2: the Dummy file system is required for the writer-proof case.");
            return true;
        }
        std::error_code ec;
        std::filesystem::create_directories(srcRoot, ec);
        if (! WriteFilledTestFile(copyFile, kFileBytes, 0x6Bu) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
        {
            Fail(L"R3-2: could not stage the source file or the Dummy destination folder.");
            return true;
        }
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = 1u;
        state.taskA                                               = startTransfer(FILESYSTEM_COPY, copyFile);
        if (! state.taskA.has_value())
        {
            restoreVerifySetting();
            Fail(L"R3-2: Verify On copy into the Dummy destination did not start.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        restoreVerifySetting();
        const CompletedTaskInfo& info = completed->second;
        const auto* const result      = info.sourceItemResults.empty() || ! info.sourceItemResults.front().has_value() ? nullptr : &info.sourceItemResults.front().value();
        if (info.hr != S_OK || result == nullptr || result->verification != FileOperations::VerificationState::Verified ||
            info.verificationProviderProofCount != 1u || info.verificationHostReadbackCount != 0u || info.conflictPromptCount != 0u)
        {
            Fail(std::format(L"R3-2: Verify On copy into the Dummy destination must be Verified from the writer proof alone ({}).", describe(info)));
            return true;
        }
        if (! dummyBytesEqual(dummyRoot + L"/verify.bin", copyBytes))
        {
            Fail(L"R3-2: the verified copy did not land with the source bytes.");
            return true;
        }
        if (! WriteFilledTestFile(moveFile, kFileBytes, 0x2Eu))
        {
            Fail(L"R3-2: could not stage the Move source file.");
            return true;
        }
        state.taskB = startTransfer(FILESYSTEM_MOVE, moveFile);
        if (! state.taskB.has_value())
        {
            Fail(L"R3-2: Managed Move into the Dummy destination did not start.");
            return true;
        }
        state.stepState = 2u;
        return false;
    }
    if (state.stepState == 2u)
    {
        const auto completed = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        const CompletedTaskInfo& info = completed->second;
        const auto* const result      = info.sourceItemResults.empty() || ! info.sourceItemResults.front().has_value() ? nullptr : &info.sourceItemResults.front().value();
        std::error_code ec;
        const bool sourceGone = ! std::filesystem::exists(moveFile, ec);
        if (info.hr != S_OK || result == nullptr || result->sourceDisposition != FileOperations::SourceDisposition::Removed || ! sourceGone ||
            info.conflictPromptCount != 0u)
        {
            std::wstring diagnostics;
            std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> entries;
            state.fileOps->CollectDiagnostics(entries);
            for (const auto& entry : entries)
            {
                if (entry.taskId == state.taskB.value())
                {
                    diagnostics += std::format(L" [{} 0x{:08X} {}]", entry.category, static_cast<unsigned long>(entry.status), entry.message);
                }
            }
            Fail(std::format(L"R3-2: Managed Move into the Dummy destination must remove the source after the writer proof matched (sourceGone={} {} "
                             L"item status=0x{:08X} completion={} publication={} strategy={}).{}",
                             sourceGone,
                             describe(info),
                             result != nullptr ? static_cast<unsigned long>(result->status) : 0xFFFFFFFFul,
                             result != nullptr ? static_cast<unsigned>(result->completion) : 255u,
                             result != nullptr ? static_cast<unsigned>(result->publication) : 255u,
                             result != nullptr ? static_cast<unsigned>(result->strategy) : 255u,
                             diagnostics));
            return true;
        }
        if (! dummyBytesEqual(dummyRoot + L"/move.bin", moveBytes))
        {
            Fail(L"R3-2: the moved content did not land with the source bytes.");
            return true;
        }
        static_cast<void>(SelfTest::RemoveAll(srcRoot));
        NextStep(state, SelfTestState::Step::R3_3_PublicationFaultMatrix);
        return false;
    }
    return false;
}

case SelfTestState::Step::R3_3_PublicationFaultMatrix:
{
    // R3-3: every bridge fault hook, on a bound route (Local -> Local), an identity-less route with
    // a writer proof (Local -> Dummy) and an identity-less route without proof (Local -> Dummy,
    // "writerProof":false), for Copy and Move. Each row must respect the publication invariants of
    // the transaction record: an object exists under the final name only when the item reports it
    // published; a Move source is removed only when the item was published and its verification
    // did not fail; a retained source still exists; no owned stage is left beside the final name.
    using Task = FolderWindow::FileOperationState::Task;
    struct FaultRow final
    {
        const wchar_t* name;
        void (*arm)();
        void (*disarm)();
        bool verify;
    };
    static constexpr std::array<FaultRow, 17> kRows{{
        {L"failNextFileCopy", [] { SetFileOpsBridgeFailNextFileCopiesForSelfTest(1u); }, [] { SetFileOpsBridgeFailNextFileCopiesForSelfTest(0u); }, false},
        {L"failSourceGetSize", [] { SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(1u); }, [] { SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(0u); }, false},
        {L"failDestinationGetSize", [] { SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(1u); }, [] { SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(0u); }, false},
        {L"failDestinationOpen", [] { SetFileOpsBridgeFailNextDestinationOpenForSelfTest(1u, E_ACCESSDENIED); }, [] { SetFileOpsBridgeFailNextDestinationOpenForSelfTest(0u, S_OK); }, false},
        {L"nullSourceReader", [] { SetFileOpsBridgeNullNextSourceReaderForSelfTest(1u); }, [] { SetFileOpsBridgeNullNextSourceReaderForSelfTest(0u); }, false},
        {L"wrongDestinationSize", [] { SetFileOpsBridgeReportWrongDestinationSizeForSelfTest(1u); }, [] { SetFileOpsBridgeReportWrongDestinationSizeForSelfTest(0u); }, false},
        {L"failStageEntropy", [] { SetFileOpsBridgeFailNextStageEntropyForSelfTest(1u); }, [] { SetFileOpsBridgeFailNextStageEntropyForSelfTest(0u); }, false},
        {L"overReportRead", [] { SetFileOpsBridgeOverReportNextReadForSelfTest(1u); }, [] { SetFileOpsBridgeOverReportNextReadForSelfTest(0u); }, false},
        {L"prematureEofRead", [] { SetFileOpsBridgePrematureEofNextReadForSelfTest(1u); }, [] { SetFileOpsBridgePrematureEofNextReadForSelfTest(0u); }, false},
        {L"underConsumeWrite", [] { SetFileOpsBridgeUnderConsumeNextWriteForSelfTest(1u); }, [] { SetFileOpsBridgeUnderConsumeNextWriteForSelfTest(0u); static_cast<void>(TakeFileOpsBridgeUnderConsumeNextWriteAttemptsForSelfTest()); }, false},
        {L"overReportWrite", [] { SetFileOpsBridgeOverReportNextWriteForSelfTest(1u); }, [] { SetFileOpsBridgeOverReportNextWriteForSelfTest(0u); }, false},
        {L"injectFileReparse", [] { SetFileOpsBridgeInjectFileReparseForSelfTest(1u); }, [] { SetFileOpsBridgeInjectFileReparseForSelfTest(0u); }, false},
        {L"verificationUnavailable", [] { SetFileOpsVerificationForceUnavailableForSelfTest(1u); }, [] { SetFileOpsVerificationForceUnavailableForSelfTest(0u); }, true},
        {L"verificationMismatch", [] { SetFileOpsVerificationForceMismatchForSelfTest(1u); }, [] { SetFileOpsVerificationForceMismatchForSelfTest(0u); }, true},
        {L"verificationHostReadback", [] { SetFileOpsVerificationForceHostReadbackForSelfTest(1u); }, [] { SetFileOpsVerificationForceHostReadbackForSelfTest(0u); }, true},
        {L"managedCleanupKnownNonCommit", [] { SetFileOpsManagedCleanupKnownNonCommitForSelfTest(E_ACCESSDENIED, 1u); }, [] { SetFileOpsManagedCleanupKnownNonCommitForSelfTest(S_OK, 0u); }, false},
        {L"managedCleanupUnknownOutcome", [] { SetFileOpsManagedCleanupUnknownOutcomeForSelfTest(1u); }, [] { SetFileOpsManagedCleanupUnknownOutcomeForSelfTest(0u); }, false},
    }};
    constexpr size_t kRouteCount            = 3u; // 0 Local->Local, 1 Local->Dummy (proof), 2 Local->Dummy (no proof)
    constexpr size_t kRowCount              = kRows.size() * kRouteCount * 2u;
    constexpr size_t kFileBytes             = 64u * 1024u;
    const std::filesystem::path srcRoot     = state.tempRoot / L"r3-3-matrix-src";
    const std::filesystem::path localDst    = state.tempRoot / L"r3-3-matrix-dst";
    const std::wstring dummyDst             = L"/r3-3-matrix-dst";
    const auto restoreVerifySetting         = [&]() noexcept
    {
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy =
            state.fileOperationsOriginal.has_value() ? state.fileOperationsOriginal->verifyAfterCopy : false;
    };
    // stepState encodes the row (0..kRowCount-1) * 2 + phase (0 = start, 1 = waiting).
    const size_t row      = state.stepState / 2u;
    const size_t rowIndex = row % kRows.size();
    const size_t route    = (row / kRows.size()) % kRouteCount;
    const bool isMove     = (row / (kRows.size() * kRouteCount)) != 0u;
    const std::wstring fileName = std::format(L"matrix-{}.bin", row);
    const std::filesystem::path sourceFile = srcRoot / fileName;
    const auto dummyObjectExists = [&](const std::wstring& dummyPath) noexcept -> bool
    {
        wil::com_ptr<IFileSystemIO> io;
        wil::com_ptr<IFileReader> reader;
        return state.fsDummy && SUCCEEDED(state.fsDummy->QueryInterface(IID_PPV_ARGS(io.addressof()))) && io &&
               SUCCEEDED(io->CreateFileReader(dummyPath.c_str(), reader.addressof())) && reader;
    };
    const auto describeRow = [&]() { return std::format(L"row {} ({}, route {}, {})", row, kRows[rowIndex].name, route, isMove ? L"move" : L"copy"); };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 600'000ull))
    {
        kRows[rowIndex].disarm();
        restoreVerifySetting();
        Fail(std::format(L"R3_3_PublicationFaultMatrix timed out at {}.", describeRow()));
        return true;
    }
    if (row >= kRowCount)
    {
        restoreVerifySetting();
        static_cast<void>(SelfTest::RemoveAll(srcRoot));
        static_cast<void>(SelfTest::RemoveAll(localDst));
        NextStep(state, SelfTestState::Step::R3_4_ReplaceWithoutOccupantTokenRefused);
        return false;
    }
    if ((state.stepState % 2u) == 0u)
    {
        if (! state.fsDummy)
        {
            Fail(L"R3-3: the Dummy file system is required for the fault matrix.");
            return true;
        }
        if (rowIndex == 0u && ! isMove)
        {
            // First row of a route: reset the Dummy proof knob for that route.
            static_cast<void>(SetPluginConfiguration(
                state.infoDummy.get(),
                route == 2u
                    ? R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0","writerProof":false})json"
                    : R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        }
        std::error_code ec;
        std::filesystem::create_directories(srcRoot, ec);
        if (route == 0u)
        {
            static_cast<void>(SelfTest::RemoveAll(localDst));
            std::filesystem::create_directories(localDst, ec);
        }
        if (! WriteFilledTestFile(sourceFile, kFileBytes, static_cast<unsigned>(0x40u + (row % 64u))) ||
            (route != 0u && ! EnsureDummyFolderExists(state.fsDummy.get(), dummyDst)))
        {
            Fail(std::format(L"R3-3: could not stage {}.", describeRow()));
            return true;
        }
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = kRows[rowIndex].verify ? 1u : 0u;
        kRows[rowIndex].arm();
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 isMove ? FILESYSTEM_MOVE : FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceFile},
                                                 route == 0u ? localDst : std::filesystem::path(dummyDst),
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 route == 0u ? state.fsLocal : state.fsDummy);
        if (! state.taskA.has_value())
        {
            kRows[rowIndex].disarm();
            restoreVerifySetting();
            Fail(std::format(L"R3-3: {} did not start.", describeRow()));
            return true;
        }
        state.stepState += 1u;
        return false;
    }

    if (Task* const task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
    {
        if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
        {
            // A fault row's failure surfaces the ordinary retryable conflict; every row answers it the
            // safe way (Skip when offered, else Cancel) so the invariants judge what was left behind.
            const Task::ConflictAction answer = PromptHasAction(prompt.value(), Task::ConflictAction::Skip)     ? Task::ConflictAction::Skip
                                                : PromptHasAction(prompt.value(), Task::ConflictAction::Cancel) ? Task::ConflictAction::Cancel
                                                                                                                : Task::ConflictAction::Proceed;
            task->SubmitConflictDecision(answer, false);
            return false;
        }
    }
    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    kRows[rowIndex].disarm();
    restoreVerifySetting();
    const CompletedTaskInfo& info = completed->second;
    const auto* const result      = info.sourceItemResults.empty() || ! info.sourceItemResults.front().has_value() ? nullptr : &info.sourceItemResults.front().value();
    std::error_code ec;
    const bool sourceExists      = std::filesystem::exists(sourceFile, ec);
    const bool destinationExists = route == 0u ? std::filesystem::exists(localDst / fileName, ec) : dummyObjectExists(dummyDst + L"/" + fileName);
    size_t localDestinationEntries = 0u;
    if (route == 0u)
    {
        for (const auto& entry : std::filesystem::directory_iterator(localDst, ec))
        {
            static_cast<void>(entry);
            ++localDestinationEntries;
        }
    }
    std::wstring violation;
    if (result == nullptr)
    {
        violation = L"no item result";
    }
    else
    {
        const bool published = result->publication == FileOperations::PublicationState::Published;
        const bool unknown   = result->publication == FileOperations::PublicationState::Unknown;
        if (FAILED(result->status) && result->completion == FileOperations::ItemCompletion::Completed &&
            ! (published && result->verification == FileOperations::VerificationState::Unavailable))
        {
            // Published content whose requested verification was unavailable completes with
            // ERROR_PARTIAL_COPY by contract; any other failed status must not read as Completed.
            violation = L"failed status reported as Completed";
        }
        else if (published && ! destinationExists)
        {
            violation = L"published but nothing under the final name";
        }
        else if (! published && ! unknown && destinationExists)
        {
            violation = L"not published but an object exists under the final name";
        }
        else if (route == 0u && localDestinationEntries != (destinationExists ? 1u : 0u))
        {
            violation = std::format(L"{} entries left in the destination folder", localDestinationEntries);
        }
        else if (! isMove && (result->sourceDisposition != FileOperations::SourceDisposition::Retained || ! sourceExists))
        {
            violation = L"a Copy must keep its source";
        }
        else if (isMove && result->sourceDisposition == FileOperations::SourceDisposition::Removed &&
                 (sourceExists || ! published || result->verification == FileOperations::VerificationState::Failed))
        {
            violation = L"Move source removed without a published, unfailed destination";
        }
        else if (isMove && result->sourceDisposition == FileOperations::SourceDisposition::Retained && ! sourceExists)
        {
            violation = L"Move source reported retained but it is gone";
        }
        else if (! kRows[rowIndex].verify && result->verification != FileOperations::VerificationState::NotRequested &&
                 result->verification != FileOperations::VerificationState::NotApplicable)
        {
            violation = L"verification reported without being requested";
        }
    }
    if (! violation.empty())
    {
        Fail(std::format(L"R3-3 fault matrix violation at {}: {} (task hr=0x{:08X} status=0x{:08X} completion={} publication={} verification={} disposition={} sourceExists={} destinationExists={}).",
                         describeRow(),
                         violation,
                         static_cast<unsigned long>(info.hr),
                         result != nullptr ? static_cast<unsigned long>(result->status) : 0xFFFFFFFFul,
                         result != nullptr ? static_cast<unsigned>(result->completion) : 255u,
                         result != nullptr ? static_cast<unsigned>(result->publication) : 255u,
                         result != nullptr ? static_cast<unsigned>(result->verification) : 255u,
                         result != nullptr ? static_cast<unsigned>(result->sourceDisposition) : 255u,
                         sourceExists,
                         destinationExists));
        return true;
    }
    state.stepState += 1u;
    return false;
}

case SelfTestState::Step::R3_4_ReplaceWithoutOccupantTokenRefused:
{
    // R0-RC3: the user grants Overwrite on a Dummy occupant whose basic information the host could
    // not read (no token). The replacement must be refused before the first byte: the occupant
    // survives with its bytes, nothing is published, and the item ends without a replacement.
    using Task                          = FolderWindow::FileOperationState::Task;
    constexpr size_t kFileBytes         = 16u * 1024u;
    const std::filesystem::path srcRoot = state.tempRoot / L"r3-4-no-token-src";
    const std::filesystem::path srcFile = srcRoot / L"replace.bin";
    const std::wstring dummyRoot        = L"/r3-4-no-token";
    const std::wstring dummyFile        = dummyRoot + L"/replace.bin";
    const std::string occupantBytes(kFileBytes, static_cast<char>(0xB4));
    const auto dummyBytesEqual = [&](std::string_view expected) noexcept -> bool
    {
        wil::com_ptr<IFileSystemIO> io;
        wil::com_ptr<IFileReader> reader;
        if (! state.fsDummy || FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(io.addressof()))) || ! io ||
            FAILED(io->CreateFileReader(dummyFile.c_str(), reader.addressof())) || ! reader)
        {
            return false;
        }
        uint64_t sizeBytes = 0u;
        if (FAILED(reader->GetSize(&sizeBytes)) || sizeBytes != expected.size())
        {
            return false;
        }
        std::vector<char> actual(expected.size());
        unsigned long bytesRead = 0u;
        return SUCCEEDED(reader->Read(actual.data(), static_cast<unsigned long>(actual.size()), &bytesRead)) && bytesRead == actual.size() &&
               std::string_view(actual.data(), actual.size()) == expected;
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        SetFileOpsBridgeFailNextDestinationBasicInfoForSelfTest(0u);
        Fail(std::format(L"R3_4_ReplaceWithoutOccupantTokenRefused timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        if (! state.fsDummy)
        {
            Fail(L"R3-4: the Dummy file system is required for the no-token replace case.");
            return true;
        }
        std::error_code ec;
        std::filesystem::create_directories(srcRoot, ec);
        if (! WriteFilledTestFile(srcFile, kFileBytes, 0x31u) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) ||
            ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, occupantBytes, true))
        {
            Fail(L"R3-4: could not stage the source file and the Dummy occupant.");
            return true;
        }
        SetFileOpsBridgeFailNextDestinationBasicInfoForSelfTest(1u);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFile},
                                                 std::filesystem::path(dummyRoot),
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            SetFileOpsBridgeFailNextDestinationBasicInfoForSelfTest(0u);
            Fail(L"R3-4: the Copy onto the Dummy occupant did not start.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        Task* task        = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }
        if (prompt->bucket != Task::ConflictBucket::RegularFileExists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
        {
            SetFileOpsBridgeFailNextDestinationBasicInfoForSelfTest(0u);
            Fail(std::format(L"R3-4: expected the Exists prompt with Overwrite (bucket={}).", static_cast<int>(prompt->bucket)));
            return true;
        }
        task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
        state.stepState = 2u;
        return false;
    }
    if (state.stepState == 2u)
    {
        // The refused replacement surfaces as the ordinary retryable conflict (answered Skip) or
        // ends the item; either way nothing may have been written.
        Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value() && PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
        {
            task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
            return false;
        }
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        SetFileOpsBridgeFailNextDestinationBasicInfoForSelfTest(0u);
        const CompletedTaskInfo& info = completed->second;
        const auto* const result      = info.sourceItemResults.empty() || ! info.sourceItemResults.front().has_value() ? nullptr : &info.sourceItemResults.front().value();
        if (result == nullptr || result->publication == FileOperations::PublicationState::Published || info.hr == S_OK)
        {
            Fail(std::format(L"R3-4: a replacement without an occupant token must not publish (task hr=0x{:08X} publication={}).",
                             static_cast<unsigned long>(info.hr),
                             result != nullptr ? static_cast<unsigned>(result->publication) : 255u));
            return true;
        }
        if (! dummyBytesEqual(occupantBytes))
        {
            Fail(L"R3-4: the occupant must survive untouched when the replacement cannot be conditioned on it.");
            return true;
        }
        static_cast<void>(SelfTest::RemoveAll(srcRoot));
        NextStep(state, SelfTestState::Step::RC3_9_TransferDestinationNameRefused);
        return false;
    }
    return false;
}

case SelfTestState::Step::RC3_9_TransferDestinationNameRefused:
{
    // R0-RC3 (9): a Copy whose destination leaf breaks the destination provider's child-name
    // contract is refused at admission, before any path is joined and before any byte moves. The
    // Dummy source accepts a trailing dot in a leaf; the local destination does not (Win32 would
    // silently strip it, publishing a different name).
    const std::wstring dummyRoot                = L"/rc3-9-name-contract";
    const std::wstring dummyFile                = dummyRoot + L"/trailing-dot.";
    const std::filesystem::path destinationRoot = state.tempRoot / L"rc3-9-name-contract-dst";
    if (! state.fsDummy || ! state.fsLocal || ! state.fileOps)
    {
        Fail(L"RC3-9: the Dummy and local file systems are required for the name-contract case.");
        return true;
    }
    if (HasTimedOut(state, GetTickCount64(), 60'000ull))
    {
        Fail(std::format(L"RC3-9 timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        std::error_code ec;
        std::filesystem::create_directories(destinationRoot, ec);
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, "trailing dot", true))
        {
            Fail(L"RC3-9: could not stage the Dummy source with a trailing-dot leaf.");
            return true;
        }
        // C1 (R1d-OR1 a3): the child-name contract is a provider call, so it runs in Preparing on the
        // task thread. Admission only proves the leaf has a shape and publishes the task; the task
        // then fails on its card with the provider's reason before any mutation.
        uint64_t taskId       = 0;
        const HRESULT admitHr = state.fileOps->AdmitOperation(FILESYSTEM_COPY,
                                                              FolderWindow::Pane::Left,
                                                              FolderWindow::Pane::Right,
                                                              state.fsDummy,
                                                              {std::filesystem::path(dummyFile)},
                                                              destinationRoot,
                                                              FILESYSTEM_FLAG_NONE,
                                                              false,
                                                              0,
                                                              FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                              false,
                                                              state.fsLocal,
                                                              &taskId,
                                                              {},
                                                              {},
                                                              {},
                                                              {},
                                                              std::nullopt,
                                                              FileOperations::DeleteOrigin::PaneCommand,
                                                              std::nullopt,
                                                              std::nullopt,
                                                              {},
                                                              {});
        if (FAILED(admitHr) || taskId == 0u)
        {
            Fail(std::format(L"RC3-9: admission must publish the task and leave the destination name to Preparing (hr=0x{:08X}, task={}).",
                             static_cast<unsigned long>(admitHr),
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
    std::error_code existsEc;
    if (done->second.hr != HRESULT_FROM_WIN32(ERROR_INVALID_NAME) || std::filesystem::exists(destinationRoot / L"trailing-dot.", existsEc) ||
        std::filesystem::exists(destinationRoot / L"trailing-dot", existsEc))
    {
        Fail(std::format(L"RC3-9: Preparing must fail the task with ERROR_INVALID_NAME and leave the destination untouched (hr=0x{:08X}).",
                         static_cast<unsigned long>(done->second.hr)));
        return true;
    }
    static_cast<void>(SelfTest::RemoveAll(destinationRoot));
    NextStep(state, SelfTestState::Step::R4T_DeepTreeCopyCompletesIteratively);
    return false;
}

case SelfTestState::Step::R4T_DeepTreeCopyCompletesIteratively:
{
    // R4-T1: the bridge walks with an explicit frame stack, so a tree far deeper than the former
    // 128-edge recursion ceiling completes. A 300-level Dummy chain (one file per level) is copied
    // (parallel producer) and the copy is copied again with verification requested (sequential
    // walker); both complete with every item published and no traversal-limit diagnostic. Dummy
    // paths have no length limit, so the walkers, not the provider, are what this exercises.
    constexpr uint32_t kLevels                  = 300u;
    constexpr std::wstring_view kRoot           = L"/r4t-deep";
    const std::wstring sourceRoot               = std::wstring(kRoot) + L"/src";
    const std::wstring copyRoot                 = std::wstring(kRoot) + L"/copy";
    const std::wstring moveRoot                 = std::wstring(kRoot) + L"/moved";
    const auto chainPath = [&](const std::wstring& root, uint32_t levels) noexcept
    {
        std::wstring path = root;
        for (uint32_t level = 1u; level <= levels; ++level)
        {
            path += L"/l" + std::to_wstring(level);
        }
        return path;
    };
    const auto restoreVerification = [&]() noexcept
    {
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy =
            state.fileOperationsOriginal.has_value() ? state.fileOperationsOriginal->verifyAfterCopy : false;
    };
    const auto dummyExists = [&](const std::wstring& path) noexcept -> bool
    {
        wil::com_ptr<IFileSystemIO> io;
        unsigned long attributes = 0;
        return state.fsDummy && SUCCEEDED(state.fsDummy->QueryInterface(IID_PPV_ARGS(io.addressof()))) && io &&
               SUCCEEDED(io->GetAttributes(path.c_str(), &attributes));
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        restoreVerification();
        Fail(std::format(L"R4T_DeepTreeCopyCompletesIteratively timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        if (! state.fsDummy || ! state.fileOps)
        {
            Fail(L"R4T: the Dummy file system is required for the deep-tree case.");
            return true;
        }
        if (! EnsureDummyFolderExists(state.fsDummy.get(), std::wstring(kRoot)) || ! EnsureDummyFolderExists(state.fsDummy.get(), sourceRoot) ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), copyRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), moveRoot))
        {
            Fail(L"R4T: could not create the deep-tree roots.");
            return true;
        }
        std::wstring current = sourceRoot;
        for (uint32_t level = 1u; level <= kLevels; ++level)
        {
            current += L"/l" + std::to_wstring(level);
            if (! EnsureDummyFolderExists(state.fsDummy.get(), current) || ! DummyWriteTextFile(state.fsDummy.get(), current + L"/f.txt", "r4t", true))
            {
                Fail(std::format(L"R4T: could not stage level {} of the deep chain.", level));
                return true;
            }
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(sourceRoot)},
                                                 std::filesystem::path(copyRoot),
                                                 FILESYSTEM_FLAG_RECURSIVE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            Fail(L"R4T: the deep-tree Copy did not start.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        const std::wstring deepestCopied = chainPath(copyRoot + L"/src", kLevels) + L"/f.txt";
        if (FAILED(completed->second.hr) || ! dummyExists(deepestCopied))
        {
            Fail(std::format(L"R4T: the 300-level Copy must complete through the iterative walker (hr=0x{:08X}, deepest present={}).",
                             static_cast<unsigned long>(completed->second.hr),
                             dummyExists(deepestCopied)));
            return true;
        }
        // A Dummy-source directory Move is refused as indeterminate by design (no bound source
        // identity), so the sequential walker is reached through verification: a verified Copy
        // runs with a within-folder budget of one.
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = true;
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(copyRoot + L"/src")},
                                                 std::filesystem::path(moveRoot),
                                                 FILESYSTEM_FLAG_RECURSIVE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            restoreVerification();
            Fail(L"R4T: the verified deep-tree Copy did not start.");
            return true;
        }
        state.stepState = 2u;
        return false;
    }
    if (state.stepState == 2u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        restoreVerification();
        const std::wstring deepestMoved = chainPath(moveRoot + L"/src", kLevels) + L"/f.txt";
        if (FAILED(completed->second.hr) || ! dummyExists(deepestMoved))
        {
            Fail(std::format(L"R4T: the verified 300-level Copy must complete through the sequential walker (hr=0x{:08X}, deepest present={}).",
                             static_cast<unsigned long>(completed->second.hr),
                             dummyExists(deepestMoved)));
            return true;
        }
        NextStep(state, SelfTestState::Step::R4T2_DeepTreeLocalDeleteCompletes);
        return false;
    }
    return false;
}

case SelfTestState::Step::R4T2_DeepTreeLocalDeleteCompletes:
{
    // R4-T2: the Local provider's permanent recursive Delete keeps its frames on an explicit stack,
    // so a directory chain far deeper than the former 128-level walk-depth ceiling is removed. The
    // chain is created through the Local plugin (extended-length paths), lies beyond MAX_PATH, and
    // is deleted permanently through the host File Operations.
    constexpr uint32_t kLevels             = 300u;
    const std::filesystem::path root       = state.tempRoot / L"r4t2-deep";
    const ULONGLONG nowTick                = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        HostClearTestPromptResultOverride();
        Fail(std::format(L"R4T2_DeepTreeLocalDeleteCompletes timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        if (! state.fsLocal || ! state.fileOps)
        {
            Fail(L"R4T2: the local file system is required for the deep-delete case.");
            return true;
        }
        std::error_code ec;
        std::filesystem::remove_all(root, ec);
        std::wstring current = root.native();
        if (! EnsureDummyFolderExists(state.fsLocal.get(), current))
        {
            Fail(L"R4T2: could not create the deep-delete root.");
            return true;
        }
        for (uint32_t level = 1u; level <= kLevels; ++level)
        {
            current += L"\\d" + std::to_wstring(level);
            if (! EnsureDummyFolderExists(state.fsLocal.get(), current) || ! WriteTestFile(std::filesystem::path(current) / L"f.txt", 16u))
            {
                Fail(std::format(L"R4T2: could not create level {} of the deep chain (path length {}).", level, current.size()));
                return true;
            }
        }
        if (current.size() <= MAX_PATH)
        {
            Fail(std::format(L"R4T2: the deep chain must exceed MAX_PATH; the deepest path has {} characters.", current.size()));
            return true;
        }
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {root},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        HostClearTestPromptResultOverride();
        if (! state.taskA.has_value())
        {
            Fail(L"R4T2: the deep-tree permanent Delete did not start.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        std::error_code ec;
        const bool rootGone = ! std::filesystem::exists(root, ec);
        if (FAILED(completed->second.hr) || ! rootGone)
        {
            Fail(std::format(L"R4T2: the 300-level permanent Delete must complete through the iterative walker (hr=0x{:08X}, root gone={}).",
                             static_cast<unsigned long>(completed->second.hr),
                             rootGone));
            return true;
        }
        NextStep(state, SelfTestState::Step::R4T3_WideDirectoryCopyCompletes);
        return false;
    }
    return false;
}

case SelfTestState::Step::R4T3_WideDirectoryCopyCompletes:
{
    // R4-T3: one directory's retained listing is reported, not a ceiling. A Dummy directory with
    // 20,000 children of 88-character names is copied through the bridge (parallel producer). Under
    // the former 8 MiB aggregate metadata budget (listing buffer plus two owned-name sets per open
    // frame) this single directory exceeded the ceiling and the Copy stopped with
    // bridge.traversal.metadataBudget; now it completes with every child published.
    constexpr uint32_t kChildren      = 20'000u;
    constexpr std::wstring_view kRoot = L"/r4t3-wide";
    const std::wstring sourceRoot     = std::wstring(kRoot) + L"/src";
    const std::wstring copyRoot       = std::wstring(kRoot) + L"/copy";
    const std::wstring pad(80u, L'x');
    const auto childName = [&](uint32_t index) { return std::format(L"f{:05}-{}", index, pad); };
    // The harness seeds the Dummy provider with 5 ms of simulated latency per access; 20,000 bridge
    // copies would spend the whole case budget sleeping. This case runs at zero latency with the
    // same structure keys (no reseed) and restores the cached seed configuration afterwards.
    const auto restoreDummyConfig = [&]() noexcept
    {
        if (! state.dummySeedConfig.empty())
        {
            static_cast<void>(SetPluginConfiguration(state.infoDummy.get(), state.dummySeedConfig));
        }
    };
    const auto dummyExists = [&](const std::wstring& path) noexcept -> bool
    {
        wil::com_ptr<IFileSystemIO> io;
        unsigned long attributes = 0;
        return state.fsDummy && SUCCEEDED(state.fsDummy->QueryInterface(IID_PPV_ARGS(io.addressof()))) && io &&
               SUCCEEDED(io->GetAttributes(path.c_str(), &attributes));
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 600'000ull))
    {
        restoreDummyConfig();
        Fail(std::format(L"R4T3_WideDirectoryCopyCompletes timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        if (! state.fsDummy || ! state.fileOps)
        {
            Fail(L"R4T3: the Dummy file system is required for the wide-directory case.");
            return true;
        }
        std::string fastConfig = state.dummySeedConfig;
        if (const size_t latencyPos = fastConfig.find("\"latencyMs\":5"); latencyPos != std::string::npos)
        {
            fastConfig.replace(latencyPos, std::string_view("\"latencyMs\":5").size(), "\"latencyMs\":0");
        }
        if (fastConfig.empty() || ! SetPluginConfiguration(state.infoDummy.get(), fastConfig))
        {
            Fail(L"R4T3: could not apply the zero-latency Dummy configuration.");
            return true;
        }
        if (! EnsureDummyFolderExists(state.fsDummy.get(), std::wstring(kRoot)) || ! EnsureDummyFolderExists(state.fsDummy.get(), sourceRoot) ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), copyRoot))
        {
            restoreDummyConfig();
            Fail(L"R4T3: could not create the wide-directory roots.");
            return true;
        }
        for (uint32_t index = 0u; index < kChildren; ++index)
        {
            if (! DummyWriteTextFile(state.fsDummy.get(), sourceRoot + L"/" + childName(index), "r4t3", true))
            {
                restoreDummyConfig();
                Fail(std::format(L"R4T3: could not stage child {} of the wide directory.", index));
                return true;
            }
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(sourceRoot)},
                                                 std::filesystem::path(copyRoot),
                                                 FILESYSTEM_FLAG_RECURSIVE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            restoreDummyConfig();
            Fail(L"R4T3: the wide-directory Copy did not start.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        restoreDummyConfig();
        const std::wstring firstCopied = copyRoot + L"/src/" + childName(0u);
        const std::wstring lastCopied  = copyRoot + L"/src/" + childName(kChildren - 1u);
        if (FAILED(completed->second.hr) || completed->second.bridgeTraversalLimitHitCount != 0u || ! dummyExists(firstCopied) ||
            ! dummyExists(lastCopied))
        {
            Fail(std::format(L"R4T3: the 20,000-child directory Copy must complete without a traversal ceiling (hr=0x{:08X}, limit hits={}, "
                             L"retained metadata high-water={} bytes, first present={}, last present={}).",
                             static_cast<unsigned long>(completed->second.hr),
                             completed->second.bridgeTraversalLimitHitCount,
                             completed->second.bridgeTraversalMaxMetadataBytes,
                             dummyExists(firstCopied),
                             dummyExists(lastCopied)));
            return true;
        }
        NextStep(state, SelfTestState::Step::R4A02_DeleteOverCopySourceWarnsAndQueues);
        return false;
    }
    return false;
}

case SelfTestState::Step::R4A02_DeleteOverCopySourceWarnsAndQueues:
case SelfTestState::Step::R4A02_DontStartCancelsBeforeMutation:
{
    using Task = FolderWindow::FileOperationState::Task;
    // R4-A02-1: a Delete injected over the source of a running Copy overlaps it (WriteSource over
    // ReadSource). During Preparing the Delete raises the same-host overlap warning on its card with
    // Queue after these tasks (default) and Don't start. Queue: the Delete waits behind the Copy and
    // runs once it terminates (source gone, Copy complete). Don't start: the Delete ends as a Preparing
    // cancel, nothing is removed, the Copy completes. The Copy is slow because it publishes into the
    // seeded Dummy provider (5 ms per access).
    const bool dontStart                = state.step == SelfTestState::Step::R4A02_DontStartCancelsBeforeMutation;
    const wchar_t* const label          = dontStart ? L"R4A02 don't start" : L"R4A02 queue after";
    const std::filesystem::path srcDir  = state.tempRoot / (dontStart ? L"r4a02-dontstart-src" : L"r4a02-queue-src");
    constexpr std::wstring_view kDummyRoot = L"/r4a02";
    const std::wstring dummyDest        = std::wstring(kDummyRoot) + (dontStart ? L"/dontstart" : L"/queue");
    const ULONGLONG nowTick             = GetTickCount64();
    if (HasTimedOut(state, nowTick, 300'000ull))
    {
        HostClearTestPromptResultOverride();
        Fail(std::format(L"{} timed out at step {}.", label, state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        if (! state.fsLocal || ! state.fsDummy || ! state.fileOps)
        {
            Fail(std::format(L"{}: Local and Dummy file systems are required.", label));
            return true;
        }
        std::error_code ec;
        std::filesystem::remove_all(srcDir, ec);
        std::filesystem::create_directories(srcDir, ec);
        for (uint32_t index = 0u; index < 200u; ++index)
        {
            if (! WriteTestFile(srcDir / std::format(L"f{:03}.bin", index), 4u * 1024u))
            {
                Fail(std::format(L"{}: could not stage source file {}.", label, index));
                return true;
            }
        }
        if (! EnsureDummyFolderExists(state.fsDummy.get(), std::wstring(kDummyRoot)) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyDest))
        {
            Fail(std::format(L"{}: could not create the Dummy destination.", label));
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcDir},
                                                 std::filesystem::path(dummyDest),
                                                 FILESYSTEM_FLAG_RECURSIVE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            Fail(std::format(L"{}: the Copy did not start.", label));
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        // Inject the Delete once the Copy is actually running (its scopes are active).
        Task* copyTask = state.fileOps->FindTask(state.taskA.value());
        if (copyTask == nullptr || ! copyTask->HasEnteredOperation())
        {
            if (state.completedTasks.contains(state.taskA.value()))
            {
                Fail(std::format(L"{}: the Copy completed before the Delete could be injected.", label));
                return true;
            }
            return false;
        }
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {srcDir},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        HostClearTestPromptResultOverride();
        if (! state.taskB.has_value())
        {
            Fail(std::format(L"{}: the Delete did not start.", label));
            return true;
        }
        state.stepState = 2u;
        return false;
    }
    if (state.stepState == 2u)
    {
        Task* deleteTask  = state.fileOps->FindTask(state.taskB.value());
        const auto prompt = TryGetConflictPromptCopy(deleteTask);
        if (! prompt.has_value())
        {
            if (state.completedTasks.contains(state.taskB.value()))
            {
                Fail(std::format(L"{}: the Delete ran or waited without the same-host overlap warning.", label));
                return true;
            }
            return false;
        }
        if (prompt->bucket != Task::ConflictBucket::SameHostOverlap || ! PromptHasAction(prompt.value(), Task::ConflictAction::Proceed) ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::RunConcurrently) ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::Cancel) || prompt->defaultAction != Task::ConflictAction::Proceed ||
            prompt->escapeAction != Task::ConflictAction::Cancel || prompt->overlapTaskId != state.taskA.value())
        {
            Fail(std::format(L"{}: the warning must name the overlapping Copy (task {}) with Queue (default), Run, and Don't start "
                             L"(bucket={}, named task={}).",
                             label,
                             state.taskA.value(),
                             static_cast<unsigned int>(prompt->bucket),
                             prompt->overlapTaskId));
            return true;
        }
        deleteTask->SubmitConflictDecision(dontStart ? Task::ConflictAction::Cancel : Task::ConflictAction::Proceed, false);
        state.stepState = 3u;
        return false;
    }
    if (state.stepState == 3u)
    {
        const auto copyDone   = state.completedTasks.find(state.taskA.value());
        const auto deleteDone = state.completedTasks.find(state.taskB.value());
        if (copyDone == state.completedTasks.end() || deleteDone == state.completedTasks.end())
        {
            if (! dontStart && deleteDone == state.completedTasks.end() && copyDone == state.completedTasks.end())
            {
                Task* deleteTask = state.fileOps->FindTask(state.taskB.value());
                if (deleteTask != nullptr && deleteTask->HasEnteredOperation() && ! deleteTask->IsOverlapQueued())
                {
                    Fail(std::format(L"{}: a queued Delete must carry the overlap edge until the Copy terminates.", label));
                    return true;
                }
            }
            return false;
        }
        std::error_code ec;
        const bool sourceExists = std::filesystem::exists(srcDir, ec);
        if (dontStart)
        {
            if (deleteDone->second.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || ! sourceExists || FAILED(copyDone->second.hr))
            {
                Fail(std::format(L"{}: Don't start must end the Delete as a cancel with the source intact (delete hr=0x{:08X}, source exists={}, copy hr=0x{:08X}).",
                                 label,
                                 static_cast<unsigned long>(deleteDone->second.hr),
                                 sourceExists,
                                 static_cast<unsigned long>(copyDone->second.hr)));
                return true;
            }
            NextStep(state, SelfTestState::Step::R4A02_ReadReadDoesNotWarn);
            return false;
        }
        if (FAILED(deleteDone->second.hr) || sourceExists || FAILED(copyDone->second.hr) || deleteDone->second.interlockWaitCount == 0u)
        {
            Fail(std::format(L"{}: Queue after must run the Delete after the Copy (delete hr=0x{:08X}, source exists={}, copy hr=0x{:08X}, waits={}).",
                             label,
                             static_cast<unsigned long>(deleteDone->second.hr),
                             sourceExists,
                             static_cast<unsigned long>(copyDone->second.hr),
                             deleteDone->second.interlockWaitCount));
            return true;
        }
        NextStep(state, SelfTestState::Step::R4A02_DontStartCancelsBeforeMutation);
        return false;
    }
    return false;
}

case SelfTestState::Step::R4A02_ReadReadDoesNotWarn:
{
    using Task = FolderWindow::FileOperationState::Task;
    // R4-A02-1: two Copies reading one source into disjoint destinations are read/read: no warning,
    // both run and complete.
    const std::filesystem::path srcDir = state.tempRoot / L"r4a02-readread-src";
    const std::filesystem::path dstA   = state.tempRoot / L"r4a02-readread-a";
    const std::filesystem::path dstB   = state.tempRoot / L"r4a02-readread-b";
    const ULONGLONG nowTick            = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(std::format(L"R4A02 read/read timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        std::error_code ec;
        for (const auto& dir : {srcDir, dstA, dstB})
        {
            std::filesystem::remove_all(dir, ec);
            std::filesystem::create_directories(dir, ec);
        }
        for (uint32_t index = 0u; index < 64u; ++index)
        {
            if (! WriteTestFile(srcDir / std::format(L"f{:03}.bin", index), 16u * 1024u))
            {
                Fail(std::format(L"R4A02 read/read: could not stage source file {}.", index));
                return true;
            }
        }
        state.taskA = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {srcDir}, dstA, FILESYSTEM_FLAG_RECURSIVE, false);
        state.taskB = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {srcDir}, dstB, FILESYSTEM_FLAG_RECURSIVE, false);
        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            Fail(L"R4A02 read/read: one of the Copies did not start.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        Task* second = state.fileOps->FindTask(state.taskB.value());
        if (second != nullptr && TryGetConflictPromptCopy(second).has_value())
        {
            Fail(L"R4A02 read/read: two Copies reading one source must not raise the overlap warning.");
            return true;
        }
        const auto doneA = state.completedTasks.find(state.taskA.value());
        const auto doneB = state.completedTasks.find(state.taskB.value());
        if (doneA == state.completedTasks.end() || doneB == state.completedTasks.end())
        {
            return false;
        }
        if (FAILED(doneA->second.hr) || FAILED(doneB->second.hr))
        {
            Fail(std::format(L"R4A02 read/read: both Copies must complete (hr A=0x{:08X}, hr B=0x{:08X}).",
                             static_cast<unsigned long>(doneA->second.hr),
                             static_cast<unsigned long>(doneB->second.hr)));
            return true;
        }
        NextStep(state, SelfTestState::Step::Phase5_DiscoverySingleTraversal);
        return false;
    }
    return false;
}

case SelfTestState::Step::R4A02_SameDestinationWarns:
{
    using Task = FolderWindow::FileOperationState::Task;
    // Two tasks are injected back-to-back with the same publish leaf. Their immutable scope vectors
    // are published before either advisory scan, so at least the later scan must name the other task
    // even when both workers finish Preparing before either enters operation.
    const std::filesystem::path rootA       = state.tempRoot / L"r4a02-same-destination-a";
    const std::filesystem::path rootB       = state.tempRoot / L"r4a02-same-destination-b";
    const std::filesystem::path sourceA     = rootA / L"shared";
    const std::filesystem::path sourceB     = rootB / L"shared";
    const std::filesystem::path destination = state.tempRoot / L"r4a02-same-destination-out";
    const ULONGLONG nowTick                 = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(std::format(L"R4A02 same destination timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        std::error_code ec;
        for (const auto& path : {rootA, rootB, destination})
        {
            std::filesystem::remove_all(path, ec);
        }
        std::filesystem::create_directories(sourceA, ec);
        std::filesystem::create_directories(sourceB, ec);
        std::filesystem::create_directories(destination, ec);
        for (uint32_t index = 0u; index < 96u; ++index)
        {
            if (! WriteTestFile(sourceA / std::format(L"a{:03}.bin", index), 64u * 1024u) ||
                ! WriteTestFile(sourceB / std::format(L"b{:03}.bin", index), 64u * 1024u))
            {
                Fail(std::format(L"R4A02 same destination could not stage source file {}.", index));
                return true;
            }
        }
        state.taskA = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {sourceA}, destination, FILESYSTEM_FLAG_RECURSIVE, false);
        state.taskB = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {sourceB}, destination, FILESYSTEM_FLAG_RECURSIVE, false);
        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            Fail(L"R4A02 same destination could not start both Copies.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        Task* const taskA = state.fileOps->FindTask(state.taskA.value());
        Task* const taskB = state.fileOps->FindTask(state.taskB.value());
        const auto promptA = TryGetConflictPromptCopy(taskA);
        const auto promptB = TryGetConflictPromptCopy(taskB);
        if (! promptA.has_value() && ! promptB.has_value())
        {
            if (state.completedTasks.contains(state.taskA.value()) || state.completedTasks.contains(state.taskB.value()))
            {
                Fail(L"R4A02 same destination completed without an overlap warning.");
                return true;
            }
            return false;
        }
        const auto validPrompt = [&](const Task::ConflictPromptState& prompt, uint64_t otherTaskId) noexcept
        {
            return prompt.bucket == Task::ConflictBucket::SameHostOverlap && PromptHasAction(prompt, Task::ConflictAction::Proceed) &&
                   PromptHasAction(prompt, Task::ConflictAction::RunConcurrently) && PromptHasAction(prompt, Task::ConflictAction::Cancel) &&
                   prompt.defaultAction == Task::ConflictAction::Proceed && prompt.escapeAction == Task::ConflictAction::Cancel &&
                   prompt.overlapTaskId == otherTaskId;
        };
        if ((promptA.has_value() && ! validPrompt(promptA.value(), state.taskB.value())) ||
            (promptB.has_value() && ! validPrompt(promptB.value(), state.taskA.value())))
        {
            Fail(L"R4A02 same destination warning did not name the concurrently prepared Copy and both scheduling actions.");
            return true;
        }
        if (promptA.has_value())
        {
            taskA->SubmitConflictDecision(Task::ConflictAction::Cancel, false);
            if (promptB.has_value())
            {
                taskB->SubmitConflictDecision(Task::ConflictAction::Proceed, false);
            }
            state.stepState = 2u; // A canceled, B completes.
        }
        else
        {
            taskB->SubmitConflictDecision(Task::ConflictAction::Cancel, false);
            state.stepState = 3u; // B canceled, A completes.
        }
        return false;
    }
    if (state.stepState == 2u || state.stepState == 3u)
    {
        const auto completedA = state.completedTasks.find(state.taskA.value());
        const auto completedB = state.completedTasks.find(state.taskB.value());
        if (completedA == state.completedTasks.end() || completedB == state.completedTasks.end())
        {
            return false;
        }
        const HRESULT canceledHr = state.stepState == 2u ? completedA->second.hr : completedB->second.hr;
        const HRESULT successHr  = state.stepState == 2u ? completedB->second.hr : completedA->second.hr;
        if (canceledHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || FAILED(successHr))
        {
            Fail(std::format(L"R4A02 same destination expected one pre-mutation cancel and one success (cancel=0x{:08X}, success=0x{:08X}).",
                             static_cast<unsigned long>(canceledHr),
                             static_cast<unsigned long>(successHr)));
            return true;
        }
        NextStep(state, SelfTestState::Step::R4A02_RunConcurrentSameDestinationStaysSafe);
        return false;
    }
    return false;
}

case SelfTestState::Step::R4A02_RunConcurrentSameDestinationStaysSafe:
{
    using Task = FolderWindow::FileOperationState::Task;
    // R4-A02-2: start one throttled Local Copy, then explicitly Run a second Copy whose publish
    // scope is the same destination subtree. The second task bypasses only the disclosed active
    // relation, both tasks are admitted at once, and disjoint members from both immutable result
    // sets reach the destination without disabling ordinary per-item policy.
    const std::filesystem::path rootA   = state.tempRoot / L"r4a02-run-a";
    const std::filesystem::path rootB   = state.tempRoot / L"r4a02-run-b";
    const std::filesystem::path rootC   = state.tempRoot / L"r4a02-run-c";
    const std::filesystem::path sourceA = rootA / L"shared";
    const std::filesystem::path sourceB = rootB / L"shared";
    const std::filesystem::path sourceC = rootC / L"shared";
    const std::filesystem::path destination = state.tempRoot / L"r4a02-run-out";
    constexpr uint64_t kConcurrentWitnessSpeedLimitBytesPerSecond = 128ull * 1024ull;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 300'000ull))
    {
        Fail(std::format(L"R4A02 Run concurrent timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        if (! state.fsLocal || ! state.fileOps)
        {
            Fail(L"R4A02 Run concurrent requires the Local file system.");
            return true;
        }
        std::error_code ec;
        std::filesystem::remove_all(rootA, ec);
        std::filesystem::remove_all(rootB, ec);
        std::filesystem::remove_all(rootC, ec);
        std::filesystem::remove_all(destination, ec);
        std::filesystem::create_directories(sourceA, ec);
        std::filesystem::create_directories(sourceB, ec);
        std::filesystem::create_directories(sourceC, ec);
        std::filesystem::create_directories(destination, ec);
        for (uint32_t index = 0u; index < 400u; ++index)
        {
            if (! WriteTestFile(sourceA / std::format(L"a{:03}.bin", index), 4u * 1024u) ||
                ! WriteTestFile(sourceB / std::format(L"b{:03}.bin", index), 4u * 1024u))
            {
                Fail(std::format(L"R4A02 Run concurrent could not stage source file {}.", index));
                return true;
            }
        }
        if (! WriteTestFile(sourceC / L"c000.bin", 4u * 1024u))
        {
            Fail(L"R4A02 Run concurrent could not stage the undisclosed third task.");
            return true;
        }
        state.r4a02ConcurrentAdmissionObserved = false;
        state.taskC.reset();
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceA},
                                                 destination,
                                                 FILESYSTEM_FLAG_RECURSIVE,
                                                 false,
                                                 kConcurrentWitnessSpeedLimitBytesPerSecond,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"R4A02 Run concurrent could not start the first Copy.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        Task* const first = state.fileOps->FindTask(state.taskA.value());
        if (first == nullptr || ! first->HasEnteredOperation())
        {
            if (state.completedTasks.contains(state.taskA.value()))
            {
                Fail(L"R4A02 Run concurrent first Copy completed before the relation could be injected.");
                return true;
            }
            return false;
        }
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceB},
                                                 destination,
                                                 FILESYSTEM_FLAG_RECURSIVE,
                                                 false,
                                                 kConcurrentWitnessSpeedLimitBytesPerSecond,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        if (! state.taskB.has_value())
        {
            Fail(L"R4A02 Run concurrent could not start the second Copy.");
            return true;
        }
        state.stepState = 2u;
        return false;
    }
    if (state.stepState == 2u)
    {
        Task* const second = state.fileOps->FindTask(state.taskB.value());
        const auto prompt = TryGetConflictPromptCopy(second);
        if (! prompt.has_value())
        {
            if (state.completedTasks.contains(state.taskB.value()))
            {
                Fail(L"R4A02 Run concurrent second Copy completed without the overlap warning.");
                return true;
            }
            return false;
        }
        if (prompt->bucket != Task::ConflictBucket::SameHostOverlap || prompt->overlapTaskId != state.taskA.value() ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::Proceed) ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::RunConcurrently) ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::Cancel) || prompt->defaultAction != Task::ConflictAction::Proceed ||
            prompt->escapeAction != Task::ConflictAction::Cancel)
        {
            Fail(L"R4A02 Run concurrent warning must offer Queue (default), Run, and Don't start for the named active task.");
            return true;
        }
        second->SubmitConflictDecision(Task::ConflictAction::RunConcurrently, false);
        state.stepState = 3u;
        return false;
    }

    Task* const first  = state.fileOps->FindTask(state.taskA.value());
    Task* const second = state.fileOps->FindTask(state.taskB.value());
    if (first != nullptr && second != nullptr && first->HasEnteredOperation() && second->HasEnteredOperation() &&
        ! state.completedTasks.contains(state.taskA.value()) && ! state.completedTasks.contains(state.taskB.value()))
    {
        state.r4a02ConcurrentAdmissionObserved = true;
    }
    if (state.stepState == 3u && state.r4a02ConcurrentAdmissionObserved)
    {
        state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceC},
                                                 destination,
                                                 FILESYSTEM_FLAG_RECURSIVE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        if (! state.taskC.has_value())
        {
            Fail(L"R4A02 Run concurrent could not start the undisclosed third task.");
            return true;
        }
        state.stepState = 4u;
        return false;
    }
    if (state.stepState == 4u)
    {
        Task* const third = state.fileOps->FindTask(state.taskC.value());
        const auto prompt = TryGetConflictPromptCopy(third);
        if (! prompt.has_value())
        {
            if (state.completedTasks.contains(state.taskC.value()))
            {
                Fail(L"R4A02 Run concurrent third task bypassed the undisclosed overlap warning.");
                return true;
            }
            return false;
        }
        const bool namesDisclosedTask = prompt->overlapTaskId == state.taskA.value() || prompt->overlapTaskId == state.taskB.value();
        if (prompt->bucket != Task::ConflictBucket::SameHostOverlap || ! namesDisclosedTask ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::RunConcurrently))
        {
            Fail(L"R4A02 Run concurrent third task did not receive an independent warning naming an active task.");
            return true;
        }
        third->SubmitConflictDecision(Task::ConflictAction::Cancel, false);
        state.stepState = 5u;
        return false;
    }
    const auto completedA = state.completedTasks.find(state.taskA.value());
    const auto completedB = state.completedTasks.find(state.taskB.value());
    const auto completedC = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
    if (completedA == state.completedTasks.end() || completedB == state.completedTasks.end() || completedC == state.completedTasks.end())
    {
        return false;
    }
    std::error_code outputEc;
    const bool outputsPresent = std::filesystem::exists(destination / L"shared" / L"a000.bin", outputEc) &&
                                std::filesystem::exists(destination / L"shared" / L"a399.bin", outputEc) &&
                                std::filesystem::exists(destination / L"shared" / L"b000.bin", outputEc) &&
                                std::filesystem::exists(destination / L"shared" / L"b399.bin", outputEc);
    if (! state.r4a02ConcurrentAdmissionObserved || FAILED(completedA->second.hr) || FAILED(completedB->second.hr) ||
        completedC->second.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) ||
        completedB->second.interlockWaitCount != 0u || ! outputsPresent)
    {
        Fail(std::format(L"R4A02 Run concurrent must admit both disclosed tasks and retain both result sets "
                         L"(observed={}, A=0x{:08X}, B=0x{:08X}, C=0x{:08X}, B waits={}, outputs={}).",
                         state.r4a02ConcurrentAdmissionObserved,
                         static_cast<unsigned long>(completedA->second.hr),
                         static_cast<unsigned long>(completedB->second.hr),
                         static_cast<unsigned long>(completedC->second.hr),
                         completedB->second.interlockWaitCount,
                         outputsPresent));
        return true;
    }
    NextStep(state, SelfTestState::Step::R4A02_LiveOutputGuardChoices);
    return false;
}

case SelfTestState::Step::R4A02_LiveOutputGuardChoices:
{
    using Task = FolderWindow::FileOperationState::Task;
    // R4-A02-2: pause a Local Copy immediately after it publishes one exact destination item, then
    // start a permanent Delete of that item. The Delete chooses Run at admission and the test drops
    // only that relation at the pre-mutation pause, modelling a newly uncovered destructive
    // consequence without relying on transfer timing. The item gate must protect the live publisher
    // for Skip, Queue, and explicit destructive consent. The Queue row forces the fixed publication
    // index into overflow and proves the conservative active-scope fallback.
    const uint32_t scenario = state.r4a02LiveOutputScenario;
    const wchar_t* const scenarioName = scenario == 0u ? L"Skip" : scenario == 1u ? L"Queue" : L"Delete anyway";
    const std::filesystem::path sourceParent = state.tempRoot / std::format(L"r4a02-live-source-{}", scenario);
    const std::filesystem::path sourceItem   = sourceParent / L"a-published.bin";
    const std::filesystem::path destination  = state.tempRoot / std::format(L"r4a02-live-output-{}", scenario);
    const std::filesystem::path publishedItem = destination / L"a-published.bin";
    const auto releasePauses = []() noexcept {
        ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest();
        SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(false);
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
    };
    const ULONGLONG nowTick                   = GetTickCount64();
    if (HasTimedOut(state, nowTick, 300'000ull))
    {
        releasePauses();
        Fail(std::format(L"R4A02 live-output {} scenario timed out at step {}.", scenarioName, state.stepState));
        return true;
    }

    if (state.stepState == 0u)
    {
        if (scenario >= 3u || ! state.fsLocal || ! state.fileOps)
        {
            releasePauses();
            Fail(L"R4A02 live-output guard requires three scenarios and the Local file system.");
            return true;
        }
        std::error_code ec;
        std::filesystem::remove_all(sourceParent, ec);
        std::filesystem::remove_all(destination, ec);
        std::filesystem::create_directories(sourceParent, ec);
        std::filesystem::create_directories(destination, ec);
        if (ec || ! WriteTestFile(sourceItem, 4u * 1024u))
        {
            releasePauses();
            Fail(std::format(L"R4A02 live-output {} scenario could not stage its publisher.", scenarioName));
            return true;
        }
        state.r4a02LiveOutputQueueObserved = false;
        SetFileOpsLiveOutputPublishedPauseForSelfTest(true);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceItem},
                                                 destination,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        if (! state.taskA.has_value())
        {
            releasePauses();
            Fail(std::format(L"R4A02 live-output {} scenario could not start its publisher.", scenarioName));
            return true;
        }
        state.stepState = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        Task* const publisher = state.fileOps->FindTask(state.taskA.value());
        std::error_code ec;
        if (! HasFileOpsLiveOutputPublishedPauseEnteredForSelfTest() || ! std::filesystem::exists(publishedItem, ec))
        {
            if (state.completedTasks.contains(state.taskA.value()))
            {
                releasePauses();
                Fail(std::format(L"R4A02 live-output {} publisher completed before its live-output window was observed.", scenarioName));
                return true;
            }
            return false;
        }
        if (publisher == nullptr || ! publisher->HasEnteredOperation() || state.completedTasks.contains(state.taskA.value()))
        {
            releasePauses();
            Fail(std::format(L"R4A02 live-output {} did not observe the published item while its task remained live.", scenarioName));
            return true;
        }
        SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(true);
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {publishedItem},
                                                 {},
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        HostClearTestPromptResultOverride();
        if (! state.taskB.has_value())
        {
            releasePauses();
            Fail(std::format(L"R4A02 live-output {} could not start the invalidating Delete.", scenarioName));
            return true;
        }
        state.stepState = 2u;
        return false;
    }

    if (state.stepState == 2u)
    {
        Task* const invalidator = state.fileOps->FindTask(state.taskB.value());
        const auto prompt       = TryGetConflictPromptCopy(invalidator);
        if (! prompt.has_value())
        {
            if (state.completedTasks.contains(state.taskB.value()))
            {
                releasePauses();
                Fail(std::format(L"R4A02 live-output {} Delete completed without its admission warning.", scenarioName));
                return true;
            }
            return false;
        }
        if (prompt->bucket != Task::ConflictBucket::SameHostOverlap || prompt->overlapTaskId != state.taskA.value() ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::RunConcurrently))
        {
            releasePauses();
            Fail(std::format(L"R4A02 live-output {} expected a Run-capable warning naming the publisher.", scenarioName));
            return true;
        }
        invalidator->SubmitConflictDecision(Task::ConflictAction::RunConcurrently, false);
        state.stepState = 3u;
        return false;
    }

    if (state.stepState == 3u)
    {
        if (! HasFileOpsPermanentDeleteBeforeLiveOutputGuardPauseEnteredForSelfTest())
        {
            if (state.completedTasks.contains(state.taskB.value()))
            {
                releasePauses();
                Fail(std::format(L"R4A02 live-output {} Delete bypassed the mutation pause.", scenarioName));
                return true;
            }
            return false;
        }
        Task* const invalidator = state.fileOps->FindTask(state.taskB.value());
        if (invalidator == nullptr)
        {
            releasePauses();
            Fail(std::format(L"R4A02 live-output {} invalidator disappeared at the mutation pause.", scenarioName));
            return true;
        }
        invalidator->DebugClearConcurrentOverlapRelationsForSelfTest();
        if (scenario == 1u)
        {
            state.fileOps->DebugForceLiveOutputIndexOverflowForSelfTest();
        }
        ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest();
        SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(false);
        state.stepState = 4u;
        return false;
    }

    if (state.stepState == 4u)
    {
        Task* const invalidator = state.fileOps->FindTask(state.taskB.value());
        const auto prompt       = TryGetConflictPromptCopy(invalidator);
        if (! prompt.has_value())
        {
            if (state.completedTasks.contains(state.taskB.value()))
            {
                releasePauses();
                Fail(std::format(L"R4A02 live-output {} Delete completed without the live-output item gate.", scenarioName));
                return true;
            }
            return false;
        }
        const bool offersInvalidation = PromptHasAction(prompt.value(), Task::ConflictAction::InvalidateLiveOutput);
        if (prompt->bucket != Task::ConflictBucket::SameHostLiveOutput || prompt->overlapTaskId != state.taskA.value() ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::QueueUntilOtherTaskFinishes) ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip) || ! PromptHasAction(prompt.value(), Task::ConflictAction::Cancel) ||
            prompt->defaultAction != Task::ConflictAction::QueueUntilOtherTaskFinishes || prompt->escapeAction != Task::ConflictAction::Cancel ||
            offersInvalidation != (scenario != 1u))
        {
            releasePauses();
            Fail(std::format(L"R4A02 live-output {} gate had the wrong publisher/actions/defaults (bucket={}, publisher={}, destructive={}).",
                             scenarioName,
                             static_cast<unsigned int>(prompt->bucket),
                             prompt->overlapTaskId,
                             offersInvalidation));
            return true;
        }

        const Task::ConflictAction answer = scenario == 0u   ? Task::ConflictAction::Skip
                                            : scenario == 1u ? Task::ConflictAction::QueueUntilOtherTaskFinishes
                                                             : Task::ConflictAction::InvalidateLiveOutput;
        invalidator->SubmitConflictDecision(answer, false);
        Task* const publisher = state.fileOps->FindTask(state.taskA.value());
        if (scenario == 1u)
        {
            state.r4a02LiveOutputQueueObserved = publisher != nullptr && ! state.completedTasks.contains(state.taskA.value()) &&
                                                 ! state.completedTasks.contains(state.taskB.value());
            ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
            SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
        }
        state.stepState = 5u;
        return false;
    }

    const auto publisherDone = state.completedTasks.find(state.taskA.value());
    const auto invalidatorDone = state.completedTasks.find(state.taskB.value());
    std::error_code existsEc;
    const bool itemExists = std::filesystem::exists(publishedItem, existsEc);
    if (scenario != 1u && invalidatorDone != state.completedTasks.end() && publisherDone == state.completedTasks.end())
    {
        state.r4a02LiveOutputQueueObserved = itemExists == (scenario == 0u);
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
    }
    if (publisherDone == state.completedTasks.end() || invalidatorDone == state.completedTasks.end())
    {
        return false;
    }

    const bool expectedItemExists = scenario == 0u;
    const HRESULT expectedInvalidatorHr = scenario == 0u ? S_FALSE : S_OK;
    if (FAILED(publisherDone->second.hr) || invalidatorDone->second.hr != expectedInvalidatorHr || itemExists != expectedItemExists ||
        ! state.r4a02LiveOutputQueueObserved)
    {
        releasePauses();
        Fail(std::format(L"R4A02 live-output {} consequence was wrong (publisher=0x{:08X}, invalidator=0x{:08X}, item exists={}, live ordering={}).",
                         scenarioName,
                         static_cast<unsigned long>(publisherDone->second.hr),
                         static_cast<unsigned long>(invalidatorDone->second.hr),
                         itemExists,
                         state.r4a02LiveOutputQueueObserved));
        return true;
    }

    releasePauses();
    ++state.r4a02LiveOutputScenario;
    state.taskA.reset();
    state.taskB.reset();
    if (state.r4a02LiveOutputScenario < 3u)
    {
        state.stepState = 0u;
        return false;
    }
    NextStep(state, SelfTestState::Step::R4A02_DisjointWritesDoNotWarn);
    return false;
}

case SelfTestState::Step::R4A02_DisjointWritesDoNotWarn:
{
    using Task = FolderWindow::FileOperationState::Task;
    const std::filesystem::path sourceA = state.tempRoot / L"r4a02-disjoint-source-a";
    const std::filesystem::path sourceB = state.tempRoot / L"r4a02-disjoint-source-b";
    const std::filesystem::path targetA = state.tempRoot / L"r4a02-disjoint-target-a";
    const std::filesystem::path targetB = state.tempRoot / L"r4a02-disjoint-target-b";
    const ULONGLONG nowTick             = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(std::format(L"R4A02 disjoint writes timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        std::error_code ec;
        for (const auto& path : {sourceA, sourceB, targetA, targetB})
        {
            std::filesystem::remove_all(path, ec);
            std::filesystem::create_directories(path, ec);
        }
        if (! WriteTestFile(sourceA / L"a.bin", 512u * 1024u) || ! WriteTestFile(sourceB / L"b.bin", 512u * 1024u))
        {
            Fail(L"R4A02 disjoint writes could not stage the sources.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {sourceA}, targetA, FILESYSTEM_FLAG_RECURSIVE, false);
        state.taskB = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {sourceB}, targetB, FILESYSTEM_FLAG_RECURSIVE, false);
        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            Fail(L"R4A02 disjoint writes could not start both Copies.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    Task* const taskA = state.fileOps->FindTask(state.taskA.value());
    Task* const taskB = state.fileOps->FindTask(state.taskB.value());
    if (TryGetConflictPromptCopy(taskA).has_value() || TryGetConflictPromptCopy(taskB).has_value())
    {
        Fail(L"R4A02 disjoint source and destination roots must not raise an overlap warning.");
        return true;
    }
    const auto completedA = state.completedTasks.find(state.taskA.value());
    const auto completedB = state.completedTasks.find(state.taskB.value());
    if (completedA == state.completedTasks.end() || completedB == state.completedTasks.end())
    {
        return false;
    }
    if (FAILED(completedA->second.hr) || FAILED(completedB->second.hr))
    {
        Fail(std::format(L"R4A02 disjoint Copies must both complete (A=0x{:08X}, B=0x{:08X}).",
                         static_cast<unsigned long>(completedA->second.hr),
                         static_cast<unsigned long>(completedB->second.hr)));
        return true;
    }
    NextStep(state, SelfTestState::Step::R4A02_RenamePublishKeepsIndexPrecise);
    return false;
}

case SelfTestState::Step::R4A02_RenamePublishKeepsIndexPrecise:
{
    // Step 3 (2026-09-06 review): a rename publishes its final name under the parent WriteSource
    // scope it holds. That publication must be indexed like any other live output, not degrade the
    // whole host to the overflow fallback (which also withholds Invalidate for every other task) for
    // as long as the rename task lives. Observed at the publish pause point while the task is live.
    using Task = FolderWindow::FileOperationState::Task;
    const std::filesystem::path root   = state.tempRoot / L"r4a02-rename-index";
    const std::filesystem::path source = root / L"before.bin";
    const std::filesystem::path target = root / L"after.bin";
    const auto releasePauses           = []() noexcept
    {
        ReleaseFileOpsLiveOutputPublishedPauseForSelfTest();
        SetFileOpsLiveOutputPublishedPauseForSelfTest(false);
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        releasePauses();
        Fail(std::format(L"R4A02 rename index timed out at step {}.", state.stepState));
        return true;
    }
    if (state.stepState == 0u)
    {
        const wil::com_ptr<IFileSystem> paneFileSystem =
            state.folderWindow ? state.folderWindow->GetFileSystem(FolderWindow::Pane::Left) : nullptr;
        if (! state.fileOps || ! paneFileSystem)
        {
            Fail(L"R4A02 rename index requires the left pane file system.");
            return true;
        }
        std::error_code ec;
        std::filesystem::remove_all(root, ec);
        std::filesystem::create_directories(root, ec);
        if (ec || ! WriteTestFile(source, 2u * 1024u))
        {
            Fail(L"R4A02 rename index could not stage its source.");
            return true;
        }
        if (state.fileOps->DebugLiveOutputIndexOverflowedForSelfTest())
        {
            Fail(L"R4A02 rename index started with the live-output index already in overflow.");
            return true;
        }
        SetFileOpsLiveOutputPublishedPauseForSelfTest(true);
        uint64_t taskId  = 0u;
        const HRESULT hr = state.fileOps->AdmitInlineRename(FolderWindow::Pane::Left, paneFileSystem, source, std::wstring(L"after.bin"), &taskId);
        if (FAILED(hr) || taskId == 0u)
        {
            releasePauses();
            Fail(std::format(L"R4A02 rename index could not admit the inline rename (hr=0x{:08X}).", static_cast<unsigned long>(hr)));
            return true;
        }
        state.taskA     = taskId;
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        if (! HasFileOpsLiveOutputPublishedPauseEnteredForSelfTest())
        {
            if (state.completedTasks.contains(state.taskA.value()))
            {
                releasePauses();
                Fail(L"R4A02 rename index: the rename completed before its publication was observed.");
                return true;
            }
            return false;
        }
        Task* const rename    = state.fileOps->FindTask(state.taskA.value());
        const bool overflowed = state.fileOps->DebugLiveOutputIndexOverflowedForSelfTest();
        releasePauses();
        if (rename == nullptr || overflowed)
        {
            Fail(std::format(L"R4A02 rename index: a live rename publication must be indexed under its parent scope, not overflow the host index (task present={}, overflow={}).",
                             rename != nullptr ? 1 : 0,
                             overflowed ? 1 : 0));
            return true;
        }
        state.stepState = 2u;
        return false;
    }
    if (state.stepState == 2u)
    {
        const auto done = state.completedTasks.find(state.taskA.value());
        if (done == state.completedTasks.end())
        {
            return false;
        }
        std::error_code ec;
        if (FAILED(done->second.hr) || ! std::filesystem::exists(target, ec) || std::filesystem::exists(source, ec))
        {
            Fail(std::format(L"R4A02 rename index: the rename must complete and publish its final name (hr=0x{:08X}).",
                             static_cast<unsigned long>(done->second.hr)));
            return true;
        }
        if (state.fileOps->DebugLiveOutputIndexOverflowedForSelfTest())
        {
            Fail(L"R4A02 rename index: the host index must leave overflow once the rename task terminates.");
            return true;
        }
        static_cast<void>(SelfTest::RemoveAll(root));
        NextStep(state, SelfTestState::Step::R4A02_LatePublisherStillWarns);
        return false;
    }
    return false;
}

case SelfTestState::Step::R4A02_LatePublisherStillWarns:
{
    // Step 4 (2026-09-06 review): the older task publishes its scopes after a newer peer already
    // prepared and decided its advisory. The older task must still name that peer, and queue after
    // it, instead of falling back to the silent interlock wait. T0 is a slow Dummy copy that keeps
    // queue mode busy so the newer Copy stays prepared; T1 is a Permanent Delete of folder X held
    // before its bind (before publication); T2 is a Copy into X admitted while T1 is held.
    using Task = FolderWindow::FileOperationState::Task;
    static ULONGLONG releaseTick = 0u;
    static bool slowCanceled     = false;
    const std::filesystem::path root       = state.tempRoot / L"r4a02-late-publisher";
    const std::filesystem::path folderX    = root / L"x";
    const std::filesystem::path sourceFile = root / L"payload.bin";
    const auto releaseAll                  = [&]() noexcept
    {
        ReleaseFileOpsPermanentDeleteBeforeBindPauseForSelfTest();
        SetFileOpsPermanentDeleteBeforeBindPauseForSelfTest(false);
        HostClearTestPromptResultOverride();
        if (state.fileOps)
        {
            state.fileOps->ApplyQueueMode(state.r4a02QueueModeOriginal);
        }
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 300'000ull))
    {
        const auto describe = [&](const wchar_t* name, const std::optional<uint64_t>& id) noexcept
        {
            if (! id.has_value())
            {
                return std::format(L" {}=none", name);
            }
            if (const auto done = state.completedTasks.find(id.value()); done != state.completedTasks.end())
            {
                return std::format(L" {}=done(0x{:08X})", name, static_cast<unsigned long>(done->second.hr));
            }
            Task* const task = state.fileOps ? state.fileOps->FindTask(id.value()) : nullptr;
            if (task == nullptr)
            {
                return std::format(L" {}=missing", name);
            }
            const auto prompt = TryGetConflictPromptCopy(task);
            return std::format(L" {}=live(entered={} queue={} waitingForOthers={} overlapQueued={} prompt={})",
                               name,
                               task->HasEnteredOperation() ? 1 : 0,
                               task->IsWaitingInQueue() ? 1 : 0,
                               task->IsWaitingForOthers() ? 1 : 0,
                               task->IsOverlapQueued() ? 1 : 0,
                               prompt.has_value() ? static_cast<unsigned int>(prompt->bucket) : 999u);
        };
        const std::wstring detail = describe(L"T0", state.taskC) + describe(L"T1", state.taskA) + describe(L"T2", state.taskB);
        releaseAll();
        Fail(std::format(L"R4A02 late publisher timed out at step {}.{}", state.stepState, detail));
        return true;
    }
    if (state.stepState == 0u)
    {
        if (! state.fileOps || ! state.fsLocal || ! state.fsDummy || state.dummyPaths.empty())
        {
            Fail(L"R4A02 late publisher requires the Local and Dummy file systems.");
            return true;
        }
        std::error_code ec;
        std::filesystem::remove_all(root, ec);
        std::filesystem::create_directories(folderX, ec);
        if (ec || ! WriteTestFile(folderX / L"existing.bin", 1024u) || ! WriteTestFile(sourceFile, 4u * 1024u) ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), L"/r4a02-late"))
        {
            Fail(L"R4A02 late publisher could not stage its folders.");
            return true;
        }
        releaseTick  = 0u;
        slowCanceled = false;
        state.r4a02QueueModeOriginal = state.fileOps->GetQueueNewTasks();
        state.fileOps->ApplyQueueMode(true);
        state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[0])},
                                                 std::filesystem::path(L"/r4a02-late"),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                 false,
                                                 8ull * 1024ull);
        if (! state.taskC.has_value())
        {
            releaseAll();
            Fail(L"R4A02 late publisher could not start the slow Dummy copy (T0).");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        Task* const slow = state.fileOps->FindTask(state.taskC.value());
        if (slow == nullptr || ! slow->HasEnteredOperation())
        {
            if (state.completedTasks.contains(state.taskC.value()))
            {
                releaseAll();
                Fail(L"R4A02 late publisher: the slow Dummy copy completed before the peers were admitted.");
                return true;
            }
            return false;
        }
        // T1: Permanent Delete of X, held inside its interlock preparation before publication.
        SetFileOpsPermanentDeleteBeforeBindPauseForSelfTest(true);
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {folderX},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 true,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        HostClearTestPromptResultOverride();
        if (! state.taskA.has_value())
        {
            releaseAll();
            Fail(L"R4A02 late publisher could not start the held Permanent Delete (T1).");
            return true;
        }
        state.stepState = 2u;
        return false;
    }
    if (state.stepState == 2u)
    {
        if (! HasFileOpsPermanentDeleteBeforeBindPauseEnteredForSelfTest())
        {
            if (state.completedTasks.contains(state.taskA.value()))
            {
                releaseAll();
                Fail(L"R4A02 late publisher: the Delete completed before it reached its pre-bind hold.");
                return true;
            }
            return false;
        }
        // T2: a Copy into X, admitted while T1 is held (T1 has published nothing yet).
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceFile},
                                                 folderX,
                                                 FILESYSTEM_FLAG_NONE,
                                                 true,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        if (! state.taskB.has_value())
        {
            releaseAll();
            Fail(L"R4A02 late publisher could not start the Copy into X (T2).");
            return true;
        }
        state.stepState = 3u;
        return false;
    }
    if (state.stepState == 3u)
    {
        // T2 prepares, decides its advisory (nothing to warn about) and waits in queue mode.
        Task* const copy = state.fileOps->FindTask(state.taskB.value());
        if (copy == nullptr || ! copy->IsWaitingInQueue())
        {
            if (state.completedTasks.contains(state.taskB.value()))
            {
                releaseAll();
                Fail(L"R4A02 late publisher: the Copy completed instead of waiting in queue mode.");
                return true;
            }
            return false;
        }
        if (TryGetConflictPromptCopy(copy).has_value())
        {
            releaseAll();
            Fail(L"R4A02 late publisher: the Copy must warn nothing (the Delete had published nothing when it prepared).");
            return true;
        }
        ReleaseFileOpsPermanentDeleteBeforeBindPauseForSelfTest();
        SetFileOpsPermanentDeleteBeforeBindPauseForSelfTest(false);
        releaseTick     = GetTickCount64();
        state.stepState = 4u;
        return false;
    }
    if (state.stepState == 4u)
    {
        Task* const remove = state.fileOps->FindTask(state.taskA.value());
        const auto prompt  = TryGetConflictPromptCopy(remove);
        if (! prompt.has_value())
        {
            // The warning appears within the Delete's own Preparing; a Delete that instead ran, or
            // sits in the queue with nothing to say ten seconds after its release, missed the peer.
            const bool silentWait = remove != nullptr && remove->IsWaitingInQueue() && GetTickCount64() - releaseTick > 10'000ull;
            if (state.completedTasks.contains(state.taskA.value()) || (remove != nullptr && remove->HasEnteredOperation()) || silentWait)
            {
                releaseAll();
                Fail(L"R4A02 late publisher: the Delete published after the Copy prepared and must still warn, naming the Copy; it ran or waited silently.");
                return true;
            }
            return false;
        }
        if (prompt->bucket != Task::ConflictBucket::SameHostOverlap || prompt->overlapTaskId != state.taskB.value() ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::Proceed) || ! PromptHasAction(prompt.value(), Task::ConflictAction::Cancel))
        {
            releaseAll();
            Fail(std::format(L"R4A02 late publisher: the Delete's warning must name the Copy (task {}) with Queue after and Don't start (bucket={}, named task={}).",
                             state.taskB.value(),
                             static_cast<unsigned int>(prompt->bucket),
                             prompt->overlapTaskId));
            return true;
        }
        remove->SubmitConflictDecision(Task::ConflictAction::Proceed, false);
        state.stepState = 5u;
        return false;
    }
    if (state.stepState == 5u)
    {
        if (! slowCanceled)
        {
            // The slow copy only had to keep queue mode busy; drain the queue now.
            if (Task* const slow = state.fileOps->FindTask(state.taskC.value()); slow != nullptr)
            {
                slow->RequestCancel();
            }
            slowCanceled = true;
        }
        const auto copyDone   = state.completedTasks.find(state.taskB.value());
        const auto removeDone = state.completedTasks.find(state.taskA.value());
        const auto slowDone   = state.completedTasks.find(state.taskC.value());
        if (copyDone == state.completedTasks.end() || removeDone == state.completedTasks.end() || slowDone == state.completedTasks.end())
        {
            return false;
        }
        releaseAll();
        std::error_code ec;
        if (FAILED(copyDone->second.hr) || FAILED(removeDone->second.hr) || std::filesystem::exists(folderX, ec) ||
            removeDone->second.completionTick < copyDone->second.completionTick)
        {
            Fail(std::format(L"R4A02 late publisher: the Copy must complete first (hr=0x{:08X}) and the queued Delete after it (hr=0x{:08X}, folder exists={}).",
                             static_cast<unsigned long>(copyDone->second.hr),
                             static_cast<unsigned long>(removeDone->second.hr),
                             std::filesystem::exists(folderX, ec) ? 1 : 0));
            return true;
        }
        static_cast<void>(SelfTest::RemoveAll(root));
        NextStep(state, SelfTestState::Step::Phase5_DiscoverySingleTraversal);
        return false;
    }
    return false;
}

case SelfTestState::Step::Phase5_DiscoverySingleTraversal:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        Fail(L"Phase5_DiscoverySingleTraversal timed out.");
        return true;
    }

    constexpr unsigned int kDirectoryCount = 8u;
    constexpr unsigned int kFilesPerDirectory = 8u;
    constexpr size_t kFileBytes = 1024u;
    constexpr uint64_t kExpectedBytes =
        static_cast<uint64_t>(kDirectoryCount) * kFilesPerDirectory * kFileBytes;
    constexpr unsigned long kExpectedFiles = kDirectoryCount * kFilesPerDirectory;
    constexpr unsigned long kExpectedDirectories = kDirectoryCount + 1u;
    const std::filesystem::path sourceRoot = state.tempRoot / L"discovery-single-traversal-src";
    const std::filesystem::path destinationRoot = state.tempRoot / L"discovery-single-traversal-dst";
    const FileSystemFlags flags =
        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(sourceRoot) || ! RecreateEmptyDirectory(destinationRoot))
        {
            Fail(L"Failed to reset the single-traversal discovery fixture.");
            return true;
        }

        for (unsigned int directoryIndex = 0u; directoryIndex < kDirectoryCount; ++directoryIndex)
        {
            const std::filesystem::path directory = sourceRoot / std::format(L"dir-{:02}", directoryIndex);
            std::error_code ec;
            std::filesystem::create_directories(directory, ec);
            if (ec)
            {
                Fail(L"Failed to create the single-traversal discovery directory.");
                return true;
            }

            for (unsigned int fileIndex = 0u; fileIndex < kFilesPerDirectory; ++fileIndex)
            {
                if (! WriteTestFile(directory / std::format(L"file-{:02}.bin", fileIndex), kFileBytes))
                {
                    Fail(L"Failed to seed the single-traversal discovery file.");
                    return true;
                }
            }
        }

        const std::string config =
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":1,"deleteRecycleBinMaxConcurrency":1,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json";
        if (! SetPluginConfiguration(state.infoLocal.get(), config))
        {
            Fail(L"Failed to configure Local concurrency for the single-traversal discovery test.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceRoot},
                                                 destinationRoot,
                                                 flags,
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start the single-traversal discovery copy.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    const auto completionIt = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completionIt == state.completedTasks.end())
    {
        return false;
    }

    if (! state.localConfigOriginal.empty() && ! SetPluginConfiguration(state.infoLocal.get(), state.localConfigOriginal))
    {
        Fail(L"Failed to restore Local configuration after the single-traversal discovery test.");
        return true;
    }

    const CompletedTaskInfo& completion = completionIt->second;
    if (FAILED(completion.hr) || ! completion.discoveryClosed ||
        completion.discoveredTotalBytes != kExpectedBytes ||
        completion.discoveredFiles != kExpectedFiles ||
        completion.discoveredDirectories != kExpectedDirectories ||
        completion.discoveryMaxQueueDepth > 256u)
    {
        Fail(std::format(
            L"Single traversal mismatch: hr=0x{:08X} closed={} bytes={}/{} files={}/{} dirs={}/{} queueMax={}.",
            static_cast<unsigned long>(completion.hr),
            completion.discoveryClosed ? 1 : 0,
            completion.discoveredTotalBytes,
            kExpectedBytes,
            completion.discoveredFiles,
            kExpectedFiles,
            completion.discoveredDirectories,
            kExpectedDirectories,
            completion.discoveryMaxQueueDepth));
        return true;
    }

    const std::filesystem::path copiedRoot = destinationRoot / sourceRoot.filename();
    for (unsigned int directoryIndex = 0u; directoryIndex < kDirectoryCount; ++directoryIndex)
    {
        for (unsigned int fileIndex = 0u; fileIndex < kFilesPerDirectory; ++fileIndex)
        {
            const std::filesystem::path copied =
                copiedRoot / std::format(L"dir-{:02}", directoryIndex) / std::format(L"file-{:02}.bin", fileIndex);
            if (! FileSizeEquals(copied, kFileBytes))
            {
                Fail(std::format(L"Single traversal failed to publish {} intact.", copied.wstring()));
                return true;
            }
        }
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.DiscoverySingleTraversal",
                      L"local-wide-tree",
                      completion.discoveryDurationUs,
                      completion.discoveredTotalBytes,
                      completion.discoveryMaxQueueDepth,
                      completion.hr);
    NextStep(state, SelfTestState::Step::Phase5_DiscoveryCancelReleasesSlot);
    return false;
}
case SelfTestState::Step::Phase5_DiscoveryCancelReleasesSlot:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        Fail(L"Phase5_DiscoveryCancelReleasesSlot timed out.");
        return true;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                               FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    if (state.stepState == 0)
    {
        state.fileOps->ApplyQueueMode(true);
        state.taskC.reset();

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[0])},
                                                 std::filesystem::path(L"/dest-a"),
                                                 flags,
                                                 false,
                                                 8ull * 1024ull);

        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start active dummy copy task for discovery cancel test.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
        if (! taskA || ! taskA->HasEnteredOperation())
        {
            if (HasTimedOut(state, nowTick, 10'000ull))
            {
                Fail(std::format(L"Active copy task did not enter the operation before queue-reorder setup; present={}.", taskA ? 1 : 0));
                return true;
            }
            return false;
        }

        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[1])},
                                                 std::filesystem::path(L"/dest-b"),
                                                 flags,
                                                 true);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start queued dummy copy task for discovery cancel test.");
            return true;
        }

        state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths.front())},
                                                 std::filesystem::path(L"/dest-c"),
                                                 flags,
                                                 true);
        if (! state.taskC.has_value())
        {
            Fail(L"Failed to start the second queued dummy copy task for queue-reorder validation.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());
        FolderWindow::FileOperationState::Task* taskC = state.fileOps->FindTask(state.taskC.value());
        if (! taskB || ! taskC || ! taskB->IsWaitingForOthers() || ! taskC->IsWaitingForOthers())
        {
            if (HasTimedOut(state, nowTick, 10'000ull))
            {
                Fail(std::format(L"Queued copy tasks did not both reach the reorderable wait state; "
                                 L"B(present={} waitingForOthers={} waitingInQueue={} entered={}) "
                                 L"C(present={} waitingForOthers={} waitingInQueue={} entered={}).",
                                 taskB ? 1 : 0,
                                 taskB && taskB->IsWaitingForOthers() ? 1 : 0,
                                 taskB && taskB->IsWaitingInQueue() ? 1 : 0,
                                 taskB && taskB->HasEnteredOperation() ? 1 : 0,
                                 taskC ? 1 : 0,
                                 taskC && taskC->IsWaitingForOthers() ? 1 : 0,
                                 taskC && taskC->IsWaitingInQueue() ? 1 : 0,
                                 taskC && taskC->HasEnteredOperation() ? 1 : 0));
                return true;
            }
            return false;
        }

        const HWND popup = state.fileOps ? state.fileOps->GetPopupHwndForSelfTest() : nullptr;
        if (! popup || IsWindow(popup) == FALSE)
        {
            return false;
        }

        FileOperationsPopupInternal::PopupLayoutDebugSnapshot queuedLayout{};
        queuedLayout.taskId = state.taskB.value();
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, queuedLayout))
        {
            if (! HasTimedOut(state, GetTickCount64(), 5'000ull))
            {
                InvalidateRect(popup, nullptr, FALSE);
                return false;
            }
            Fail(L"Failed to capture queued file-operation layout for Start now validation.");
            return true;
        }

        FileOperationsPopupInternal::PopupLayoutDebugSnapshot queuedLayoutC{};
        queuedLayoutC.taskId = state.taskC.value();
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, queuedLayoutC))
        {
            if (! HasTimedOut(state, GetTickCount64(), 5'000ull))
            {
                InvalidateRect(popup, nullptr, FALSE);
                return false;
            }
            Fail(L"Failed to capture the second queued file-operation layout for reorder validation.");
            return true;
        }

        FileOperationsPopupInternal::TaskSnapshot queuedTaskSnapshot{};
        static_cast<void>(DebugGetFileOperationsPopupTaskSnapshot(popup, state.taskB.value(), queuedTaskSnapshot));
        if (! queuedLayout.taskStartNowVisible || ! queuedLayout.taskCancelVisible || ! queuedLayout.taskSpeedLimitVisible ||
            ! queuedLayout.taskSpeedLimitUsesSelectorChrome || ! queuedLayout.taskSpeedLimitShowsCurrentValue ||
            queuedLayout.taskQueueMoveUpVisible || ! queuedLayout.taskQueueMoveDownVisible ||
            ! queuedLayoutC.taskQueueMoveUpVisible || queuedLayoutC.taskQueueMoveDownVisible)
        {
            if (! HasTimedOut(state, GetTickCount64(), 5'000ull))
            {
                return false;
            }

            Fail(std::format(L"Two queued copy tasks should expose stable reorder edges plus Start now, Speed Limit, and Cancel controls; "
                             L"first(up={} down={}) second(up={} down={}) startNow={} speedLimit={} speedLimitSelector={} speedLimitValue={} cancel={} "
                             L"snapshot(started={} waitingForOthers={} waitingInQueue={} discovery={} status={}).",
                             queuedLayout.taskQueueMoveUpVisible ? 1 : 0,
                             queuedLayout.taskQueueMoveDownVisible ? 1 : 0,
                             queuedLayoutC.taskQueueMoveUpVisible ? 1 : 0,
                             queuedLayoutC.taskQueueMoveDownVisible ? 1 : 0,
                             queuedLayout.taskStartNowVisible ? 1 : 0,
                             queuedLayout.taskSpeedLimitVisible ? 1 : 0,
                             queuedLayout.taskSpeedLimitUsesSelectorChrome ? 1 : 0,
                             queuedLayout.taskSpeedLimitShowsCurrentValue ? 1 : 0,
                             queuedLayout.taskCancelVisible ? 1 : 0,
                             queuedTaskSnapshot.started ? 1 : 0,
                             queuedTaskSnapshot.waitingForOthers ? 1 : 0,
                             queuedTaskSnapshot.waitingInQueue ? 1 : 0,
                             queuedTaskSnapshot.discoveryAheadActive ? 1 : 0,
                             static_cast<unsigned int>(queuedTaskSnapshot.statusKind)));
            return true;
        }


        FileOperationsPopupInternal::TaskSnapshot queuedTaskSnapshotC{};
        static_cast<void>(DebugGetFileOperationsPopupTaskSnapshot(popup, state.taskC.value(), queuedTaskSnapshotC));
        const uint64_t firstKeyBefore  = queuedTaskSnapshot.queueOrderKey;
        const uint64_t secondKeyBefore = queuedTaskSnapshotC.queueOrderKey;
        FileOperationsPopupInternal::PopupSelfTestInvoke moveSecondUp{};
        moveSecondUp.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskQueueMoveUp;
        moveSecondUp.taskId = state.taskC.value();
        if (! DebugInvokeFileOperationsPopup(popup, moveSecondUp))
        {
            Fail(L"Failed to invoke Move up for the second queued task through the File Operations popup.");
            return true;
        }
        FileOperationsPopupInternal::TaskSnapshot reorderedTaskB{};
        FileOperationsPopupInternal::TaskSnapshot reorderedTaskC{};
        if (! DebugGetFileOperationsPopupTaskSnapshot(popup, state.taskB.value(), reorderedTaskB) ||
            ! DebugGetFileOperationsPopupTaskSnapshot(popup, state.taskC.value(), reorderedTaskC) ||
            reorderedTaskB.queueOrderKey != secondKeyBefore || reorderedTaskC.queueOrderKey != firstKeyBefore)
        {
            Fail(std::format(L"Queue reorder should swap engine order keys; before B={} C={}, after B={} C={}.",
                             firstKeyBefore,
                             secondKeyBefore,
                             reorderedTaskB.queueOrderKey,
                             reorderedTaskC.queueOrderKey));
            return true;
        }

        FileOperationsPopupInternal::PopupSelfTestInvoke startNow{};
        startNow.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskStartNow;
        startNow.taskId = state.taskB.value();
        if (! DebugInvokeFileOperationsPopup(popup, startNow))
        {
            Fail(L"Failed to invoke queued task Start now through the File Operations popup.");
            return true;
        }

        if (taskB->IsWaitingInQueue() || taskB->IsWaitingForOthers())
        {
            Fail(L"Start now did not release the queued task wait state.");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
        if (taskA)
        {
            const HWND popup = state.fileOps ? state.fileOps->GetPopupHwndForSelfTest() : nullptr;
            if (! popup || IsWindow(popup) == FALSE)
            {
                Fail(L"Discovery-ahead copy task did not expose the File Operations popup for layout validation.");
                return true;
            }

            FileOperationsPopupInternal::PopupLayoutDebugSnapshot layout{};
            layout.taskId = state.taskA.value();
            if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, layout))
            {
                if (! HasTimedOut(state, GetTickCount64(), 5'000ull))
                {
                    InvalidateRect(popup, nullptr, FALSE);
                    return false;
                }
                Fail(L"Failed to capture File Operations popup layout while discovery-ahead copy was in progress.");
                return true;
            }

            // Single-pass discovery runs concurrently with transfer. Before the first mutation the
            // task may be Discovering; afterward the truthful primary status is Running while the
            // secondary discovery state remains visible until traversal closure.
            if (layout.taskStatusKind != FileOperationsPopupInternal::TaskSnapshot::StatusKind::Discovering &&
                layout.taskStatusKind != FileOperationsPopupInternal::TaskSnapshot::StatusKind::Running)
            {
                Fail(L"Discovery-ahead copy task should report Discovering or Running.");
                return true;
            }

            if (! layout.taskToggleCollapseVisible || ! layout.taskCancelVisible || ! layout.taskSpeedLimitVisible ||
                ! layout.taskSpeedLimitUsesSelectorChrome || ! layout.taskSpeedLimitShowsCurrentValue ||
                (layout.taskDiscoveryAheadActive && ! layout.taskSkipVisible) ||
                (! layout.taskDiscoveryAheadActive && layout.taskSkipVisible))
            {
                Fail(std::format(
                    L"Copy task should expose collapse, Speed Limit, and Cancel throughout, with Skip exactly while discovery-ahead is active; "
                    L"collapse={} discovery={} skip={} speedLimit={} speedLimitSelector={} speedLimitValue={} cancel={}.",
                    layout.taskToggleCollapseVisible ? 1 : 0,
                    layout.taskDiscoveryAheadActive ? 1 : 0,
                    layout.taskSkipVisible ? 1 : 0,
                    layout.taskSpeedLimitVisible ? 1 : 0,
                    layout.taskSpeedLimitUsesSelectorChrome ? 1 : 0,
                    layout.taskSpeedLimitShowsCurrentValue ? 1 : 0,
                    layout.taskCancelVisible ? 1 : 0));
                return true;
            }

            if (layout.taskUnderGraphProgressBarCount != 1u || layout.taskDuplicateUnderGraphItemBarVisible)
            {
                Fail(std::format(L"Discovery-ahead copy task should render one under-graph progress bar and no duplicate item bar; bars={} duplicate={}.",
                                 layout.taskUnderGraphProgressBarCount,
                                 layout.taskDuplicateUnderGraphItemBarVisible ? 1 : 0));
                return true;
            }

            if (layout.footerQueueModeSegmentedVisible || ! layout.footerQueueModeSelectorVisible || ! layout.footerQueueModeUsesSelectorChrome ||
                ! layout.footerQueueModeSelectorHitTargetActive || ! layout.footerOptionsVisible || ! layout.footerOptionsHitTargetActive ||
                ! layout.footerAggregateProgressVisible || ! layout.footerDetailsToggleVisible ||
                ! layout.footerDetailsToggleUsesDisclosureChrome || layout.footerAutoDismissVisible ||
                layout.footerDensityToggleVisible || ! layout.footerDetailsToggleRightAligned || layout.reducedMotionEnabled ||
                ! layout.autoResizeAnimationEnabled || layout.footerQueueModeAnimationEnabled || ! layout.taskCardAnimationEnabled ||
                ! layout.hostedProgressAnimationEnabled || ! layout.usesDxUiHost || layout.hostedActionControlCount == 0u ||
                layout.hostedProgressControlCount == 0u || layout.hostedGraphControlCount == 0u)
            {
                Fail(std::format(L"File Operations popup footer should expose one mode selector, Options, aggregate progress, animated resize, and a "
                                  L"right-aligned details toggle through hosted controls during active work. segmented={} selector={} selectorChrome={} selectorHit={} "
                                  L"options={} optionsHit={} aggregate={} details={} detailsDisclosure={} autoDismiss={} density={} rightAligned={} reducedMotion={} "
                                  L"animatedResize={} queueAnimation={} cardAnimation={} progressAnimation={} hosted={} actions={} progress={} graphs={}.",
                                  layout.footerQueueModeSegmentedVisible ? 1 : 0,
                                  layout.footerQueueModeSelectorVisible ? 1 : 0,
                                  layout.footerQueueModeUsesSelectorChrome ? 1 : 0,
                                  layout.footerQueueModeSelectorHitTargetActive ? 1 : 0,
                                  layout.footerOptionsVisible ? 1 : 0,
                                  layout.footerOptionsHitTargetActive ? 1 : 0,
                                  layout.footerAggregateProgressVisible ? 1 : 0,
                                  layout.footerDetailsToggleVisible ? 1 : 0,
                                  layout.footerDetailsToggleUsesDisclosureChrome ? 1 : 0,
                                  layout.footerAutoDismissVisible ? 1 : 0,
                                  layout.footerDensityToggleVisible ? 1 : 0,
                                  layout.footerDetailsToggleRightAligned ? 1 : 0,
                                  layout.reducedMotionEnabled ? 1 : 0,
                                  layout.autoResizeAnimationEnabled ? 1 : 0,
                                  layout.footerQueueModeAnimationEnabled ? 1 : 0,
                                  layout.taskCardAnimationEnabled ? 1 : 0,
                                  layout.hostedProgressAnimationEnabled ? 1 : 0,
                                  layout.usesDxUiHost ? 1 : 0,
                                  layout.hostedActionControlCount,
                                  layout.hostedProgressControlCount,
                                  layout.hostedGraphControlCount));
                return true;
            }

            if (layout.taskbarProgressState == static_cast<uint32_t>(TBPF_NOPROGRESS))
            {
                Fail(L"Active discovery-ahead copy task should expose a non-empty global taskbar progress model.");
                return true;
            }

            if (! layout.taskStatusStripeVisible || ! layout.taskStatusChipVisible ||
                layout.taskStatusVisualTone != static_cast<uint32_t>(FileOperationsPopupInternal::PopupStatusVisualTone::Accent))
            {
                Fail(std::format(L"Active discovery-ahead copy task should expose an accent status stripe and chip; stripe={} chip={} tone={}.",
                                 layout.taskStatusStripeVisible ? 1 : 0,
                                 layout.taskStatusChipVisible ? 1 : 0,
                                 layout.taskStatusVisualTone));
                return true;
            }

            if (layout.taskDiscoveryAheadActive && ! layout.taskDiscoveryIndicatorVisible)
            {
                Fail(L"Discovery-ahead should expose its dedicated card activity indicator instead of taking over the throughput graph.");
                return true;
            }

            if (layout.graphStatusAnimationEnabled)
            {
                Fail(L"Discovery-ahead should not paint a status animation over the throughput graph.");
                return true;
            }

            if (taskA->_lastProgressCallbackTick != 0 &&
                (! layout.taskTransferProgressVisible ||
                 (! layout.taskTransferProgressDeterminate && ! layout.taskTransferProgressMarqueeVisible)))
            {
                Fail(std::format(L"A discovery-ahead task with transfer callbacks should expose determinate/provisional transfer progress or a marquee; "
                                 L"visible={} determinate={} provisional={} marquee={}.",
                                 layout.taskTransferProgressVisible ? 1 : 0,
                                 layout.taskTransferProgressDeterminate ? 1 : 0,
                                 layout.taskTransferProgressProvisional ? 1 : 0,
                                 layout.taskTransferProgressMarqueeVisible ? 1 : 0));
                return true;
            }

            if (! state.folderWindow)
            {
                Fail(L"Discovery reduced-motion validation requires the folder window theme source.");
                return true;
            }

            const AppTheme motionRestoreTheme = state.folderWindow->GetTheme();
            const auto restoreMotionTheme     = wil::scope_exit([&]() noexcept { state.folderWindow->ApplyTheme(motionRestoreTheme); });
            AppTheme reducedMotionTheme       = motionRestoreTheme;
            reducedMotionTheme.reducedMotionOverride = true;
            state.folderWindow->ApplyTheme(reducedMotionTheme);

            FileOperationsPopupInternal::PopupLayoutDebugSnapshot reducedMotionLayout{};
            reducedMotionLayout.taskId = state.taskA.value();
            if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, reducedMotionLayout))
            {
                Fail(L"Failed to capture reduced-motion File Operations popup layout while discovery-ahead copy was in progress.");
                return true;
            }

            if (! reducedMotionLayout.reducedMotionEnabled || reducedMotionLayout.autoResizeAnimationEnabled ||
                reducedMotionLayout.footerQueueModeAnimationEnabled || reducedMotionLayout.graphStatusAnimationEnabled ||
                reducedMotionLayout.taskCardAnimationEnabled || reducedMotionLayout.hostedProgressAnimationEnabled)
            {
                Fail(std::format(L"Reduced-motion discovery layout should disable popup animations. reducedMotion={} animatedResize={} queueAnimation={} "
                                 L"graphStatusAnimation={} cardAnimation={} progressAnimation={}.",
                                 reducedMotionLayout.reducedMotionEnabled ? 1 : 0,
                                 reducedMotionLayout.autoResizeAnimationEnabled ? 1 : 0,
                                 reducedMotionLayout.footerQueueModeAnimationEnabled ? 1 : 0,
                                 reducedMotionLayout.graphStatusAnimationEnabled ? 1 : 0,
                                 reducedMotionLayout.taskCardAnimationEnabled ? 1 : 0,
                                 reducedMotionLayout.hostedProgressAnimationEnabled ? 1 : 0));
                return true;
            }

            taskA->RequestCancel();
            state.stepState = 4;
        }
        else if (HasTimedOut(state, nowTick, 10'000ull))
        {
            Fail(std::format(L"Active copy task disappeared before popup validation; completionSeen={}.",
                             state.completedTasks.contains(state.taskA.value()) ? 1 : 0));
            return true;
        }
        return false;
    }

    if (state.stepState == 4)
    {
        const auto itA = state.completedTasks.find(state.taskA.value());
        if (itA == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT hrA = itA->second.hr;
        if (hrA != HRESULT_FROM_WIN32(ERROR_CANCELLED) && hrA != E_ABORT)
        {
            std::wstring items;
            for (size_t index = 0u; index < itA->second.sourceItemResults.size(); ++index)
            {
                const auto& sealed = itA->second.sourceItemResults[index];
                const auto& raw    = index < itA->second.sourceItemStatuses.size() ? itA->second.sourceItemStatuses[index] : std::nullopt;
                items += std::format(L" item{}[raw=0x{:08X} sealed={} status=0x{:08X} completion={} publication={} source={}]",
                                     index,
                                     static_cast<unsigned long>(raw.value_or(S_OK)),
                                     sealed.has_value() ? 1 : 0,
                                     static_cast<unsigned long>(sealed.has_value() ? sealed->status : S_OK),
                                     sealed.has_value() ? static_cast<unsigned int>(sealed->completion) : 99u,
                                     sealed.has_value() ? static_cast<unsigned int>(sealed->publication) : 99u,
                                     sealed.has_value() ? static_cast<unsigned int>(sealed->sourceDisposition) : 99u);
            }
            Fail(std::format(L"Unexpected hr for cancelled discovery task: 0x{:08X}{}", static_cast<unsigned long>(hrA), items));
            return true;
        }

        if (state.completedTasks.find(state.taskB.value()) != state.completedTasks.end())
        {
            state.stepState = 5;
            return false;
        }

        FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());
        if (! taskB || ! taskB->HasEnteredOperation())
        {
            return false;
        }

        taskB->RequestCancel();
        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        if (state.completedTasks.find(state.taskB.value()) == state.completedTasks.end())
        {
            return false;
        }

        FolderWindow::FileOperationState::Task* taskC = state.fileOps->FindTask(state.taskC.value());
        if (taskC && ! taskC->HasEnteredOperation())
        {
            return false;
        }
        if (taskC)
        {
            taskC->RequestCancel();
        }
        state.stepState = 6;
        return false;
    }

    if (state.stepState == 6)
    {
        if (state.completedTasks.find(state.taskC.value()) == state.completedTasks.end())
        {
            return false;
        }

        NextStep(state, SelfTestState::Step::Phase5_DiscoveryCancelLatencyLocal);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase5_DiscoveryCancelLatencyLocal:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        const std::wstring status = state.taskA.has_value() ? std::format(L"task={} completionSeen={} stepState={}",
                                                                          state.taskA.value(),
                                                                          state.completedTasks.contains(state.taskA.value()) ? 1 : 0,
                                                                          state.stepState)
                                                            : L"task=(none)";
        Fail(std::format(L"Phase5_DiscoveryCancelLatencyLocal timed out. {}", status));
        return true;
    }

    constexpr ULONGLONG kCancelLatencyThresholdMs = 500ull;
    constexpr unsigned int kFileCount             = 512u;
    // Keep enough throttled payload outstanding that observing discovery-ahead cannot race a
    // complete transfer on a fast machine before the UI thread submits cancellation.
    constexpr size_t kFileBytes                   = 64u * 1024u;
    constexpr uint64_t kSpeedLimitBytesPerSecond  = 64u * 1024u;
    const std::filesystem::path sourceRoot         = state.tempRoot / L"discovery-cancel-latency-src";
    const std::filesystem::path destinationRoot    = state.tempRoot / L"discovery-cancel-latency-dst";
    const FileSystemFlags flags                    = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    if (state.stepState == 0)
    {
        state.fileOps->ApplyQueueMode(false);
        state.taskA.reset();
        state.taskB.reset();
        state.markerTick = 0;

        if (! RecreateEmptyDirectory(sourceRoot) || ! RecreateEmptyDirectory(destinationRoot))
        {
            Fail(L"Failed to reset discovery cancel-latency roots.");
            return true;
        }
        for (unsigned int index = 0u; index < kFileCount; ++index)
        {
            if (! WriteTestFile(sourceRoot / std::format(L"payload-{:04}.bin", index), kFileBytes))
            {
                Fail(L"Failed to seed the discovery cancel-latency tree.");
                return true;
            }
        }

        const std::string config =
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1,"deleteMaxConcurrency":1,"deleteRecycleBinMaxConcurrency":1,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json";
        if (! SetPluginConfiguration(state.infoLocal.get(), config))
        {
            Fail(L"Failed to apply local plugin config for discovery cancel-latency test.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceRoot},
                                                 destinationRoot,
                                                 flags,
                                                 false,
                                                 kSpeedLimitBytesPerSecond);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start local copy task for discovery cancel-latency test.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    if (state.stepState == 1)
    {
        if (! task)
        {
            return false;
        }

        if (! task->_discoveryAheadActive.load(std::memory_order_acquire))
        {
            return false;
        }

        task->RequestCancel();
        state.markerTick = nowTick;
        state.stepState  = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto completionIt = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completionIt == state.completedTasks.end())
        {
            return false;
        }

        const CompletedTaskInfo& completion = completionIt->second;
        if (completion.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) && completion.hr != E_ABORT)
        {
            Fail(std::format(L"Unexpected hr for discovery cancel-latency task: 0x{:08X}", static_cast<unsigned long>(completion.hr)));
            return true;
        }

        const ULONGLONG cancelLatencyMs =
            (state.markerTick != 0 && completion.completionTick >= state.markerTick) ? (completion.completionTick - state.markerTick) : 0ull;
        AppendLog(std::format(L"Phase5_DiscoveryCancelLatencyLocal latency={}ms threshold={}ms files={}",
                              cancelLatencyMs,
                              kCancelLatencyThresholdMs,
                              kFileCount));
        Debug::Perf::Emit(
            L"FileOps.SelfTest.DiscoveryCancelLatency",
            std::format(L"files={} fileBytes={} thresholdMs={} path={}", kFileCount, kFileBytes, kCancelLatencyThresholdMs, sourceRoot.wstring()),
            cancelLatencyMs * 1000ull,
            kFileCount,
            kCancelLatencyThresholdMs,
            completion.hr);

        if (state.markerTick == 0)
        {
            Fail(L"Discovery cancel-latency marker was not recorded.");
            return true;
        }

        if (cancelLatencyMs > kCancelLatencyThresholdMs)
        {
            Fail(std::format(L"Discovery cancel latency too high: {}ms (threshold {}ms).", cancelLatencyMs, kCancelLatencyThresholdMs));
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase5_DiscoverySkipContinues);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase5_DiscoverySkipContinues:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        Fail(L"Phase5_DiscoverySkipContinues timed out.");
        return true;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                               FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    if (state.stepState == 0)
    {
        state.fileOps->ApplyQueueMode(true);

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[0])},
                                                 std::filesystem::path(L"/dest-skip-a"),
                                                 flags,
                                                 false,
                                                 8ull * 1024ull);

        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start the active dummy copy task for discovery-Skip testing.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
        if (! taskA || ! taskA->HasEnteredOperation())
        {
            return false;
        }

        if (! taskA->_discoverySkipped.load(std::memory_order_acquire))
        {
            taskA->SkipDiscovery();
        }

        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[1])},
                                                 std::filesystem::path(L"/dest-skip-b"),
                                                 flags,
                                                 true,
                                                 8ull * 1024ull);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start the queued dummy copy task for discovery-Skip testing.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
    FolderWindow::FileOperationState::Task* taskB = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;

    if (state.stepState == 2)
    {
        if (state.completedTasks.find(state.taskA.value()) != state.completedTasks.end() ||
            state.completedTasks.find(state.taskB.value()) != state.completedTasks.end())
        {
            Fail(L"Discovery-Skip tasks completed before validation could run.");
            return true;
        }

        if (taskA && ! taskA->_discoverySkipped.load(std::memory_order_acquire))
        {
            taskA->SkipDiscovery();
        }

        if (! taskA || ! taskB)
        {
            return false;
        }

        if (taskA->_discoveryClosed.load(std::memory_order_acquire))
        {
            Fail(L"Traversal closed before the post-Skip just-in-time state could be observed.");
            return true;
        }

        if (! taskA->_discoverySkipped.load(std::memory_order_acquire))
        {
            return false;
        }

        if (! taskA->HasStarted())
        {
            return false;
        }

        if (taskB->HasEnteredOperation())
        {
            Fail(L"Skipping discovery-ahead must not bypass the global Queue admission barrier.");
            return true;
        }
        if (! taskB->IsWaitingInQueue() && ! taskB->IsWaitingForOthers())
        {
            return false;
        }

        taskA->RequestCancel();
        taskB->RequestCancel();
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (state.completedTasks.find(state.taskA.value()) == state.completedTasks.end())
        {
            return false;
        }
        if (state.completedTasks.find(state.taskB.value()) == state.completedTasks.end())
        {
            return false;
        }

        NextStep(state, SelfTestState::Step::BR3_CancelQueuePausedTransfer);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase5_CancelQueuedTask:
{
    const auto isCancelHr = [](HRESULT hr) noexcept -> bool
    {
        return hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT;
    };
    const auto summarizeTask = [&](std::optional<std::uint64_t> idOpt) -> std::wstring
    {
        if (! idOpt.has_value() || ! state.fileOps)
        {
            return L"(missing)";
        }

        const std::uint64_t id                       = idOpt.value();
        FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(id);
        const auto completionIt                      = state.completedTasks.find(id);
        const bool completed                         = completionIt != state.completedTasks.end();
        const std::wstring completedText             = completed
                                                           ? std::format(L" completed=1 hr=0x{:08X} started={} items={}/{} bytes={}",
                                                                         static_cast<unsigned long>(completionIt->second.hr),
                                                                         completionIt->second.started,
                                                                         completionIt->second.progressCompletedItems,
                                                                         completionIt->second.progressTotalItems,
                                                                         completionIt->second.progressCompletedBytes)
                                                           : L" completed=0";
        if (! task)
        {
            return std::format(L"id={} live=0{}", id, completedText);
        }

        unsigned long totalItems     = 0;
        unsigned long completedItems = 0;
        uint64_t completedBytes      = 0;
        {
            std::scoped_lock lock(task->_progressMutex);
            totalItems     = task->_progressTotalItems;
            completedItems = task->_progressCompletedItems;
            completedBytes = task->_progressCompletedBytes;
        }

        return std::format(L"id={} live=1 waiting={} entered={} started={} qpause={} discoveryAhead={} discoveryClosed={} discoverySkipped={} items={}/{} bytes={}{}",
                           id,
                           task->IsWaitingInQueue(),
                           task->HasEnteredOperation(),
                           task->HasStarted(),
                           task->IsQueuePaused(),
                           task->_discoveryAheadActive.load(std::memory_order_acquire),
                           task->_discoveryClosed.load(std::memory_order_acquire),
                           task->_discoverySkipped.load(std::memory_order_acquire),
                           completedItems,
                           totalItems,
                           completedBytes,
                           completedText);
    };

    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        Fail(std::format(L"Phase5_CancelQueuedTask timed out. stepState={} A: {} B: {} C: {}",
                         state.stepState,
                         summarizeTask(state.taskA),
                         summarizeTask(state.taskB),
                         summarizeTask(state.taskC)));
        return true;
    }

    constexpr uint64_t kQueueHoldSpeedLimitBytesPerSecond = 8ull * 1024ull;
    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                               FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    if (state.stepState == 0)
    {
        state.fileOps->ApplyQueueMode(true);
        state.taskA.reset();
        state.taskB.reset();
        state.taskC.reset();

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[0])},
                                                 std::filesystem::path(L"/dest-queued-a"),
                                                 flags,
                                                 false,
                                                 kQueueHoldSpeedLimitBytesPerSecond);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start the active dummy copy task for queued-cancel test.");
            return true;
        }

        if (auto* taskA = state.fileOps->FindTask(state.taskA.value()))
        {
            taskA->SetDesiredSpeedLimit(kQueueHoldSpeedLimitBytesPerSecond);
            taskA->SkipDiscovery();
        }

        state.stepState = 1;
        return false;
    }

    if (! state.taskA.has_value())
    {
        Fail(L"Phase5_CancelQueuedTask lost its active queue-holder task id.");
        return true;
    }

    FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());

    if (state.stepState == 1)
    {
        if (state.completedTasks.find(state.taskA.value()) != state.completedTasks.end())
        {
            Fail(std::format(L"Active queue holder completed before the queued task could be created. A: {} B: {}",
                             summarizeTask(state.taskA),
                             summarizeTask(state.taskB)));
            return true;
        }
        if (taskA && ! taskA->_discoverySkipped.load(std::memory_order_acquire))
        {
            taskA->SkipDiscovery();
        }
        if (! taskA)
        {
            return false;
        }
        taskA->SetDesiredSpeedLimit(kQueueHoldSpeedLimitBytesPerSecond);
        if (! taskA->HasEnteredOperation())
        {
            return false;
        }

        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[1])},
                                                 std::filesystem::path(L"/dest-queued-b"),
                                                 flags,
                                                 true,
                                                 kQueueHoldSpeedLimitBytesPerSecond);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start the queued dummy copy task for queued-cancel test.");
            return true;
        }

        if (auto* taskB = state.fileOps->FindTask(state.taskB.value()))
        {
            taskB->SetDesiredSpeedLimit(kQueueHoldSpeedLimitBytesPerSecond);
        }

        state.stepState = 2;
        return false;
    }

    if (! state.taskB.has_value())
    {
        Fail(L"Phase5_CancelQueuedTask lost its queued task id.");
        return true;
    }

    FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());

    if (state.stepState == 2)
    {
        if (state.completedTasks.find(state.taskA.value()) != state.completedTasks.end())
        {
            Fail(std::format(L"Active queue holder completed before queued cancellation could observe the queue. A: {} B: {}",
                             summarizeTask(state.taskA),
                             summarizeTask(state.taskB)));
            return true;
        }
        if (state.completedTasks.find(state.taskB.value()) != state.completedTasks.end())
        {
            Fail(std::format(L"Queued task completed before cancellation could observe the queue. A: {} B: {}",
                             summarizeTask(state.taskA),
                             summarizeTask(state.taskB)));
            return true;
        }
        if (! taskB)
        {
            return false;
        }
        taskB->SetDesiredSpeedLimit(kQueueHoldSpeedLimitBytesPerSecond);
        if (taskB->HasEnteredOperation())
        {
            Fail(std::format(L"Queued task entered operation before queued cancellation. A: {} B: {}",
                             summarizeTask(state.taskA),
                             summarizeTask(state.taskB)));
            return true;
        }
        if (taskB->IsWaitingInQueue())
        {
            taskB->RequestCancel();
            state.stepState = 3;
        }
        return false;
    }

    if (state.stepState == 3)
    {
        const auto queuedCompletionIt = state.completedTasks.find(state.taskB.value());
        if (queuedCompletionIt == state.completedTasks.end())
        {
            return false;
        }
        if (! isCancelHr(queuedCompletionIt->second.hr))
        {
            Fail(std::format(L"Queued task cancellation expected cancel hr, got 0x{:08X}. A: {} B: {}",
                             static_cast<unsigned long>(queuedCompletionIt->second.hr),
                             summarizeTask(state.taskA),
                             summarizeTask(state.taskB)));
            return true;
        }

        state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[2 % state.dummyPaths.size()])},
                                                 std::filesystem::path(L"/dest-queued-c"),
                                                 flags,
                                                 true,
                                                 kQueueHoldSpeedLimitBytesPerSecond);
        if (! state.taskC.has_value())
        {
            Fail(L"Failed to start follow-up task after cancelling queued task.");
            return true;
        }
        if (auto* taskC = state.fileOps->FindTask(state.taskC.value()))
        {
            taskC->SetDesiredSpeedLimit(kQueueHoldSpeedLimitBytesPerSecond);
        }

        taskA = state.fileOps->FindTask(state.taskA.value());
        if (taskA)
        {
            taskA->RequestCancel();
        }
        else if (state.completedTasks.find(state.taskA.value()) == state.completedTasks.end())
        {
            Fail(std::format(L"Active queue holder disappeared before cancellation. A: {} B: {} C: {}",
                             summarizeTask(state.taskA),
                             summarizeTask(state.taskB),
                             summarizeTask(state.taskC)));
            return true;
        }

        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        const auto earlyCompletionIt = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
        if (earlyCompletionIt != state.completedTasks.end())
        {
            if (isCancelHr(earlyCompletionIt->second.hr))
            {
                state.stepState = 5;
                return false;
            }

            Fail(std::format(L"Follow-up queued-cancel task completed before the test could cancel it. C: {}", summarizeTask(state.taskC)));
            return true;
        }

        FolderWindow::FileOperationState::Task* taskC = state.fileOps->FindTask(state.taskC.value());
        if (! taskC)
        {
            return false;
        }

        if (! taskC->HasEnteredOperation())
        {
            return false;
        }

        taskC->RequestCancel();
        state.stepState = 5;
        return false;
    }

    const auto taskCCompletionIt = state.completedTasks.find(state.taskC.value());
    if (taskCCompletionIt == state.completedTasks.end())
    {
        return false;
    }
    if (! isCancelHr(taskCCompletionIt->second.hr))
    {
        Fail(std::format(L"Follow-up queued-cancel task expected cancel hr, got 0x{:08X}. A: {} B: {} C: {}",
                         static_cast<unsigned long>(taskCCompletionIt->second.hr),
                         summarizeTask(state.taskA),
                         summarizeTask(state.taskB),
                         summarizeTask(state.taskC)));
        return true;
    }

    const auto taskACompletionIt = state.completedTasks.find(state.taskA.value());
    if (taskACompletionIt == state.completedTasks.end())
    {
        return false;
    }
    if (! isCancelHr(taskACompletionIt->second.hr))
    {
        Fail(std::format(L"Active queue holder expected cancel hr, got 0x{:08X}. A: {} B: {} C: {}",
                         static_cast<unsigned long>(taskACompletionIt->second.hr),
                         summarizeTask(state.taskA),
                         summarizeTask(state.taskB),
                         summarizeTask(state.taskC)));
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase5_SwitchParallelToWaitDuringDiscovery);
    return false;
}
case SelfTestState::Step::BR3_CancelQueuePausedTransfer:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 30'000ull))
    {
        // Release the queue before failure cleanup, including on the unfixed baseline.
        state.fileOps->ApplyQueueMode(false);
        Fail(L"BR3 queue-paused cancellation fixture did not reach its bounded quiet point.");
        return true;
    }

    if (state.stepState == 0)
    {
        const std::filesystem::path root = state.tempRoot / L"br3-queue-paused-cancel";
        const std::filesystem::path destinationA = root / L"destination-a";
        const std::filesystem::path destinationB = root / L"destination-b";
        const std::filesystem::path sourceA = root / L"source-a.bin";
        const std::filesystem::path sourceB = root / L"source-b.bin";
        if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(destinationA) || ! RecreateEmptyDirectory(destinationB) ||
            ! WriteTestFile(sourceA, 16u * 1024u * 1024u) || ! WriteTestFile(sourceB, 16u * 1024u * 1024u))
        {
            Fail(L"BR3 queue-paused cancellation could not create its owned copy fixture.");
            return true;
        }

        state.fileOps->ApplyQueueMode(false);
        state.queuePausedTask.reset();
        constexpr uint64_t kSpeedLimit = 512u * 1024u;
        state.taskA = StartFileOperationAndGetId(state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right,
                                                state.fsLocal, {sourceA}, destinationA, FILESYSTEM_FLAG_NONE, false, kSpeedLimit);
        state.taskB = StartFileOperationAndGetId(state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right,
                                                state.fsLocal, {sourceB}, destinationB, FILESYSTEM_FLAG_NONE, false, kSpeedLimit);
        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            Fail(L"BR3 queue-paused cancellation could not admit both copy tasks.");
            return true;
        }
        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        auto* taskA = state.fileOps->FindTask(state.taskA.value());
        auto* taskB = state.fileOps->FindTask(state.taskB.value());
        if (! taskA || ! taskB || ! taskA->HasEnteredOperation() || ! taskB->HasEnteredOperation() ||
            taskA->_progressCallbackCount.load(std::memory_order_acquire) == 0u || taskB->_progressCallbackCount.load(std::memory_order_acquire) == 0u)
        {
            return false;
        }
        state.fileOps->ApplyQueueMode(true);
        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        auto* taskA = state.fileOps->FindTask(state.taskA.value());
        auto* taskB = state.fileOps->FindTask(state.taskB.value());
        if (! taskA || ! taskB || taskA->IsQueuePaused() == taskB->IsQueuePaused())
        {
            return false;
        }
        const bool aQueued = taskA->IsQueuePaused();
        auto* queuedTask = aQueued ? taskA : taskB;
        auto* holderTask = aQueued ? taskB : taskA;
        holderTask->SetPaused(true);
        if (queuedTask->_dbgPauseWaiterCount.load(std::memory_order_acquire) == 0u ||
            holderTask->_dbgPauseWaiterCount.load(std::memory_order_acquire) == 0u)
        {
            return false;
        }
        // Both workers are actually parked. The first remains active and paused throughout
        // the assertion, so finishing it or changing Queue mode cannot mask failed cancellation.
        state.queuePausedTask = aQueued ? state.taskA : state.taskB;
        state.markerTick = nowTick;
        queuedTask->RequestCancel();
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        const auto completed = state.completedTasks.find(state.queuePausedTask.value());
        const bool timedOut = nowTick >= state.markerTick && nowTick - state.markerTick > 3'000ull;
        if (completed == state.completedTasks.end() && ! timedOut)
        {
            return false;
        }
        const uint64_t holderId = state.queuePausedTask == state.taskA ? state.taskB.value() : state.taskA.value();
        auto* holderTask = state.fileOps->FindTask(holderId);
        const bool cancelled = completed != state.completedTasks.end() &&
            (completed->second.hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || completed->second.hr == E_ABORT);
        const bool passed = cancelled && holderTask && holderTask->IsPaused() && holderTask->HasEnteredOperation() && ! timedOut;
        Debug::Perf::Emit(L"FileOps.SelfTest.QueuePausedCancelLatency", L"local-two-active-transfers;holder-remains-paused",
                          (nowTick - state.markerTick) * 1000ull, passed ? 1u : 0u, 3'000'000u,
                          passed ? S_OK : HRESULT_FROM_WIN32(ERROR_TIMEOUT));
        // Recover the pre-fix loop before reporting a failed assertion, so the RED run itself
        // cannot leave a spinning task behind or hang the following family at shutdown.
        state.fileOps->ApplyQueueMode(false);
        for (const uint64_t id : {state.taskA.value(), state.taskB.value()})
        {
            if (auto* task = state.fileOps->FindTask(id))
            {
                task->RequestCancel();
            }
        }
        state.stepState = passed ? 4 : 5;
        return false;
    }

    if (state.completedTasks.find(state.taskA.value()) == state.completedTasks.end() ||
        state.completedTasks.find(state.taskB.value()) == state.completedTasks.end())
    {
        return false;
    }
    if (state.stepState == 5)
    {
        Fail(L"Cancel did not finish the queue-paused transfer within 3 seconds while its predecessor remained paused.");
        return true;
    }
    NextStep(state, SelfTestState::Step::Phase5_CancelQueuedTask);
    return false;
}
case SelfTestState::Step::Phase5_SwitchParallelToWaitDuringDiscovery:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        const auto summarizeTask = [&](std::optional<std::uint64_t> idOpt) -> std::wstring
        {
            if (! idOpt.has_value() || ! state.fileOps)
            {
                return L"(missing)";
            }

            const std::uint64_t id                       = idOpt.value();
            FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(id);
            if (! task)
            {
                return std::format(L"id={} (missing)", id);
            }

            unsigned long totalItems     = 0;
            unsigned long completedItems = 0;
            {
                std::scoped_lock lock(task->_progressMutex);
                totalItems     = task->_progressTotalItems;
                completedItems = task->_progressCompletedItems;
            }

            return std::format(L"id={} entered={} started={} qpause={} discoveryAhead={} discoveryClosed={} discoverySkipped={} items={}/{}",
                               id,
                               task->HasEnteredOperation(),
                               task->HasStarted(),
                               task->IsQueuePaused(),
                               task->_discoveryAheadActive.load(std::memory_order_acquire),
                               task->_discoveryClosed.load(std::memory_order_acquire),
                               task->_discoverySkipped.load(std::memory_order_acquire),
                               completedItems,
                               totalItems);
        };

        Fail(std::format(L"Phase5_SwitchParallelToWaitDuringDiscovery timed out. A: {} B: {}", summarizeTask(state.taskA), summarizeTask(state.taskB)));
        return true;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    const std::filesystem::path sourceA = state.tempRoot / L"discovery-a";
    const std::filesystem::path sourceB = state.tempRoot / L"discovery-b";
    const std::filesystem::path destinationA = state.tempRoot / L"discovery-switch-dst-a";
    const std::filesystem::path destinationB = state.tempRoot / L"discovery-switch-dst-b";
    constexpr uint64_t kSpeedLimitBytesPerSecond = 64u;

    if (state.stepState == 0)
    {
        state.queuePausedTask.reset();
        state.fileOps->ApplyQueueMode(false);

        // Bounded producer queues plus a low transfer rate keep both single-pass traversals open
        // long enough to exercise the Parallel -> Queue transition deterministically.
        const std::string config =
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1,"deleteMaxConcurrency":1,"deleteRecycleBinMaxConcurrency":1,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json";
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), config));

        if (! RecreateEmptyDirectory(destinationA) || ! RecreateEmptyDirectory(destinationB))
        {
            Fail(L"Failed to reset destinations for the Parallel-to-Queue discovery test.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceA},
                                                 destinationA,
                                                 flags,
                                                 false,
                                                 kSpeedLimitBytesPerSecond);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {sourceB},
                                                 destinationB,
                                                 flags,
                                                 false,
                                                 kSpeedLimitBytesPerSecond);
        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            Fail(L"Failed to start local copy tasks for the Parallel-to-Queue discovery test.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
    FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());
    if (! taskA || ! taskB)
    {
        return false;
    }

    if (state.stepState == 1)
    {
        if (taskA->HasEnteredOperation() && taskB->HasEnteredOperation())
        {
            state.fileOps->ApplyQueueMode(true);
            state.stepState = 2;
        }
        return false;
    }

    if (state.stepState == 2)
    {
        const bool aPaused = taskA->IsQueuePaused();
        const bool bPaused = taskB->IsQueuePaused();
        if (aPaused == bPaused)
        {
            return false;
        }

        state.queuePausedTask                              = aPaused ? state.taskA : state.taskB;
        FolderWindow::FileOperationState::Task* pausedTask = aPaused ? taskA : taskB;

        if (! pausedTask->_discoveryAheadActive.load(std::memory_order_acquire))
        {
            return false;
        }

        pausedTask->SkipDiscovery();
        {
            std::scoped_lock lock(pausedTask->_progressMutex);
            state.pauseResumeBaselineItems = pausedTask->_progressCompletedItems;
            state.pauseResumeBaselineBytes = pausedTask->_progressCompletedBytes;
        }
        state.markerTick = nowTick;
        state.stepState  = 3;
        return false;
    }

    if (! state.queuePausedTask.has_value())
    {
        return false;
    }

    FolderWindow::FileOperationState::Task* pausedTask = state.fileOps->FindTask(state.queuePausedTask.value());
    if (! pausedTask)
    {
        return false;
    }

    const bool discoveryStill = pausedTask->_discoveryAheadActive.load(std::memory_order_acquire);
    if (discoveryStill)
    {
        return false;
    }

    if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) < 500ull)
    {
        return false;
    }

    {
        std::scoped_lock lock(pausedTask->_progressMutex);
        if (pausedTask->_progressCompletedItems != state.pauseResumeBaselineItems ||
            pausedTask->_progressCompletedBytes != state.pauseResumeBaselineBytes)
        {
            // Queue mode is observed at bounded provider progress checkpoints. A callback already
            // in flight when the mode changes may publish its final chunk before blocking. Move the
            // baseline forward and require a complete quiet window; continuous mutation will still
            // hit the enclosing test deadline instead of being mistaken for an immediate pause.
            state.pauseResumeBaselineItems = pausedTask->_progressCompletedItems;
            state.pauseResumeBaselineBytes = pausedTask->_progressCompletedBytes;
            state.markerTick               = nowTick;
            return false;
        }
    }

    const ULONGLONG skipRequestedTick = pausedTask->_discoverySkipRequestedTick.load(std::memory_order_acquire);
    const ULONGLONG reservationReleasedTick = pausedTask->_discoveryReservationReleasedTick.load(std::memory_order_acquire);
    if (skipRequestedTick == 0u || reservationReleasedTick < skipRequestedTick)
    {
        Fail(L"Skip discovery did not atomically release the host discovery-ahead reservation.");
        return true;
    }
    const ULONGLONG releaseLatencyMs = reservationReleasedTick - skipRequestedTick;
    if (releaseLatencyMs > 50ull)
    {
        Fail(std::format(L"Skip discovery release latency was {}ms (limit 50ms).", releaseLatencyMs));
        return true;
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.DiscoverySkipRelease",
                      L"",
                      releaseLatencyMs * 1000ull,
                      state.pauseResumeBaselineItems,
                      state.pauseResumeBaselineBytes,
                      S_OK);

    NextStep(state, SelfTestState::Step::Phase5_SwitchWaitToParallelResume);
    return false;
}
case SelfTestState::Step::Phase5_SwitchWaitToParallelResume:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        const auto summarize = [&](std::optional<std::uint64_t> idOpt) -> std::wstring
        {
            if (! idOpt.has_value() || ! state.fileOps)
            {
                return L"(missing)";
            }

            const std::uint64_t id = idOpt.value();
            if (FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(id))
            {
                return std::format(L"id={} started={} qpause={} discovery={} done={} skipped={}",
                                   id,
                                   task->HasStarted(),
                                   task->IsQueuePaused(),
                                   task->_discoveryAheadActive.load(std::memory_order_acquire),
                                   task->_discoveryClosed.load(std::memory_order_acquire),
                                   task->_discoverySkipped.load(std::memory_order_acquire));
            }

            const auto it = state.completedTasks.find(id);
            if (it != state.completedTasks.end())
            {
                return std::format(L"id={} (completed hr=0x{:08X})", id, static_cast<unsigned long>(it->second.hr));
            }

            return std::format(L"id={} (missing)", id);
        };

        Fail(std::format(L"Phase5_SwitchWaitToParallelResume timed out. A: {} B: {} paused: {}",
                         summarize(state.taskA),
                         summarize(state.taskB),
                         summarize(state.queuePausedTask)));
        return true;
    }

    if (state.stepState == 0)
    {
        state.fileOps->ApplyQueueMode(false);
        state.stepState = 1;
        return false;
    }

    if (! state.queuePausedTask.has_value())
    {
        Fail(L"Phase5_SwitchWaitToParallelResume missing paused task id.");
        return true;
    }

    const std::uint64_t pausedId                       = state.queuePausedTask.value();
    FolderWindow::FileOperationState::Task* pausedTask = state.fileOps->FindTask(pausedId);
    if (state.stepState == 1)
    {
        if (! pausedTask)
        {
            const auto it = state.completedTasks.find(pausedId);
            if (it != state.completedTasks.end())
            {
                if (it->second.progressCompletedItems <= state.pauseResumeBaselineItems &&
                    it->second.progressCompletedBytes <= state.pauseResumeBaselineBytes)
                {
                    Fail(std::format(L"Queue-paused task completed without publishing progress after resume (hr=0x{:08X}).",
                                     static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                if (state.taskA.has_value() && state.taskA.value() != pausedId)
                {
                    if (FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(state.taskA.value()))
                    {
                        task->RequestCancel();
                    }
                }
                if (state.taskB.has_value() && state.taskB.value() != pausedId)
                {
                    if (FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(state.taskB.value()))
                    {
                        task->RequestCancel();
                    }
                }
                state.stepState = 2;
                return false;
            }
            return false;
        }

        if (pausedTask->IsQueuePaused())
        {
            return false;
        }

        if (! pausedTask->HasStarted())
        {
            return false;
        }

        unsigned long completedItems = 0;
        uint64_t completedBytes      = 0;
        {
            std::scoped_lock lock(pausedTask->_progressMutex);
            completedItems = pausedTask->_progressCompletedItems;
            completedBytes = pausedTask->_progressCompletedBytes;
        }
        if (completedItems <= state.pauseResumeBaselineItems && completedBytes <= state.pauseResumeBaselineBytes)
        {
            return false;
        }

        // Cancel any remaining tasks so the next phases start with a clean slate.
        pausedTask->RequestCancel();

        if (state.taskA.has_value() && state.taskA.value() != pausedId)
        {
            if (FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(state.taskA.value()))
            {
                task->RequestCancel();
            }
        }
        if (state.taskB.has_value() && state.taskB.value() != pausedId)
        {
            if (FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(state.taskB.value()))
            {
                task->RequestCancel();
            }
        }

        state.stepState = 2;
        return false;
    }

    const auto ensureCompleted = [&](std::optional<std::uint64_t> idOpt) noexcept -> bool
    {
        if (! idOpt.has_value())
        {
            return true;
        }

        return state.completedTasks.find(idOpt.value()) != state.completedTasks.end();
    };

    if (! ensureCompleted(state.queuePausedTask))
    {
        return false;
    }

    if (! ensureCompleted(state.taskA))
    {
        return false;
    }

    if (! ensureCompleted(state.taskB))
    {
        return false;
    }

    NextStep(state, SelfTestState::Step::Phase6_PopupRateSmoothing);
    return false;
}
#endif
#if defined(FILEOPS_SELFTEST_INCLUDE_PHASE6)

case SelfTestState::Step::Phase6_PopupRateSmoothing:
{
    using Kind = FileOperationsPopupInternal::PopupHitTest::Kind;
    if (DebugResolveFileOperationsHostedButtonVariant(Kind::FooterQueueMode) != RedSalamander::DxUi::ButtonVariant::Selector ||
        DebugResolveFileOperationsHostedButtonVariant(Kind::TaskSpeedLimit) != RedSalamander::DxUi::ButtonVariant::Selector ||
        DebugResolveFileOperationsHostedButtonVariant(Kind::TaskDestination) != RedSalamander::DxUi::ButtonVariant::DropDown ||
        DebugResolveFileOperationsHostedButtonVariant(Kind::FooterOptions) != RedSalamander::DxUi::ButtonVariant::IconOnly ||
        DebugResolveFileOperationsHostedButtonVariant(Kind::CompletedGroupToggle) != RedSalamander::DxUi::ButtonVariant::Disclosure ||
        DebugResolveFileOperationsHostedButtonVariant(Kind::TaskToggleCollapse) != RedSalamander::DxUi::ButtonVariant::Disclosure ||
        DebugResolveFileOperationsHostedButtonVariant(Kind::TaskPause) != RedSalamander::DxUi::ButtonVariant::Standard)
    {
        Fail(L"File Operations hosted controls did not preserve selector, menu, icon-only, disclosure, and action button semantics.");
        return true;
    }

    constexpr uint64_t kLimitedBytesPerSecond = 8ull * 1024ull * 1024ull;
    const std::wstring unlimitedText          = DebugFormatFileOperationsSpeedLimitSelectorText(0);
    const std::wstring expectedUnlimitedText  = LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_BUTTON_UNLIMITED);
    const std::wstring limitedText             = DebugFormatFileOperationsSpeedLimitSelectorText(kLimitedBytesPerSecond);
    const std::wstring expectedLimitedText =
        FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_LIMIT_BUTTON_BYTES, FormatBytesCompact(kLimitedBytesPerSecond));
    if (unlimitedText != expectedUnlimitedText || limitedText != expectedLimitedText || limitedText == unlimitedText)
    {
        Fail(std::format(L"Speed Limit selector did not expose its localized current value: unlimited='{}' limited='{}'.", unlimitedText, limitedText));
        return true;
    }

    constexpr double kBaseRate = 100.0 * 1024.0 * 1024.0;

    const double firstRate = DebugSmoothRateForDisplay(0.0, kBaseRate, 100ull);
    if (firstRate != kBaseRate)
    {
        Fail(std::format(L"Initial smoothed rate should adopt first sample, got {}.", firstRate));
        return true;
    }

    const double spikedRate = DebugSmoothRateForDisplay(firstRate, kBaseRate * 4.0, 100ull);
    if (spikedRate <= firstRate || spikedRate >= kBaseRate * 1.5)
    {
        Fail(std::format(L"Short spike should be damped: base={} spiked={}.", firstRate, spikedRate));
        return true;
    }

    const double heldRate = DebugDecayRateForCallbackSilence(firstRate, 250ull);
    if (heldRate != firstRate)
    {
        Fail(std::format(L"Short callback silence should hold displayed rate, got {}.", heldRate));
        return true;
    }

    const double decayedRate = DebugDecayRateForCallbackSilence(firstRate, 2500ull);
    if (decayedRate <= 0.0 || decayedRate >= firstRate * 0.5)
    {
        Fail(std::format(L"Long callback silence should visibly decay rate: base={} decayed={}.", firstRate, decayedRate));
        return true;
    }

    const double etaAfterShortSample = DebugSmoothEtaSecondsForDisplay(120.0, 30.0, 100ull);
    if (etaAfterShortSample <= 100.0 || etaAfterShortSample >= 120.0)
    {
        Fail(std::format(L"Short ETA sample should move gradually: eta={}.", etaAfterShortSample));
        return true;
    }

    const double etaAfterLongSample = DebugSmoothEtaSecondsForDisplay(120.0, 30.0, 2000ull);
    if (etaAfterLongSample <= 30.0 || etaAfterLongSample >= 90.0)
    {
        Fail(std::format(L"Long ETA sample should converge faster without snapping: eta={}.", etaAfterLongSample));
        return true;
    }

    const float graphStartY = DebugEaseFileOperationsGraphLatestPointYForDisplay(80.0f, 20.0f, 0ull);
    if (std::abs(graphStartY - 80.0f) > 0.01f)
    {
        Fail(std::format(L"Graph line easing should start at the previous point, got {}.", graphStartY));
        return true;
    }

    const float graphMidY = DebugEaseFileOperationsGraphLatestPointYForDisplay(80.0f, 20.0f, 130ull);
    if (graphMidY <= 20.0f || graphMidY >= 80.0f)
    {
        Fail(std::format(L"Graph line easing should move between previous and target points, got {}.", graphMidY));
        return true;
    }

    const float graphEndY = DebugEaseFileOperationsGraphLatestPointYForDisplay(80.0f, 20.0f, 260ull);
    if (std::abs(graphEndY - 20.0f) > 0.01f)
    {
        Fail(std::format(L"Graph line easing should finish at the target point, got {}.", graphEndY));
        return true;
    }

    const float resizeStart = DebugEaseFileOperationsAutoResizeFraction(0ull, 240ull);
    const float resizeMid   = DebugEaseFileOperationsAutoResizeFraction(120ull, 240ull);
    const float resizeEnd   = DebugEaseFileOperationsAutoResizeFraction(240ull, 240ull);
    if (std::abs(resizeStart) > 0.01f || resizeMid <= 0.0f || resizeMid >= 1.0f || std::abs(resizeEnd - 1.0f) > 0.01f)
    {
        Fail(std::format(L"Auto-resize easing should progress from 0 to 1 without snapping; start={} mid={} end={}.", resizeStart, resizeMid, resizeEnd));
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase6_PopupSmokeResizeAndPause);
    return false;
}
case SelfTestState::Step::Phase6_PopupSmokeResizeAndPause:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        const HWND popup     = FindCurrentProcessPopupWindow();
        const bool hasTask   = state.taskA.has_value() && state.fileOps && state.fileOps->FindTask(state.taskA.value()) != nullptr;
        const bool completed = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();
        Fail(std::format(L"Phase6_PopupSmokeResizeAndPause timed out. stepState={} popup={} taskExists={} completed={}",
                         state.stepState,
                         popup != nullptr,
                         hasTask,
                         completed));
        return true;
    }

    const std::filesystem::path srcDir  = state.tempRoot / L"phase6-src";
    const std::filesystem::path dstDir  = state.tempRoot / L"phase6-dst";

    if (state.stepState == 0)
    {
        state.fileOps->ApplyQueueMode(false);
        state.pauseResumeBaselineItems = 0;
        state.pauseResumeBaselineBytes = 0;
        if (! RecreateEmptyDirectory(srcDir))
        {
            Fail(L"Failed to reset phase6-src directory.");
            return true;
        }
        if (! RecreateEmptyDirectory(dstDir))
        {
            Fail(L"Failed to reset phase6-dst directory.");
            return true;
        }

        constexpr size_t kSourceFileCount   = 16u;
        constexpr uint64_t kSourceFileBytes = 2ull * 1024ull * 1024ull;
        std::vector<std::filesystem::path> sources;
        sources.reserve(kSourceFileCount);
        for (size_t index = 0; index < kSourceFileCount; ++index)
        {
            const std::filesystem::path source = srcDir / std::format(L"part-{:02}.bin", index);
            if (! WriteTestFile(source, kSourceFileBytes))
            {
                Fail(std::format(L"Failed to write popup smoke source file {}.", index));
                return true;
            }
            sources.push_back(source);
        }

        const FileSystemFlags flags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 std::move(sources),
                                                 dstDir,
                                                 flags,
                                                 false,
                                                 1ull * 1024ull * 1024ull);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start local copy task for popup smoke test.");
            return true;
        }

        AppendLog(std::format(L"Phase6_PopupSmokeResizeAndPause started task={}", state.taskA.value()));
        state.stepState = 1;
        return false;
    }

    const HWND popup = FindCurrentProcessPopupWindow();
    if (popup && ! state.popupOriginalRectValid)
    {
        state.popupOriginalRectValid = GetWindowRect(popup, &state.popupOriginalRect) != FALSE;
    }

    if (state.stepState == 1)
    {
        if (state.taskA.has_value())
        {
            const auto it = state.completedTasks.find(state.taskA.value());
            if (it != state.completedTasks.end())
            {
                Fail(std::format(L"Copy task completed before popup could be validated (hr=0x{:08X}).", static_cast<unsigned long>(it->second.hr)));
                return true;
            }
        }

        auto* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (! task)
        {
            if (state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end())
            {
                const HRESULT hr = state.completedTasks.find(state.taskA.value())->second.hr;
                Fail(std::format(L"Copy task completed before popup/pause validation finished (hr=0x{:08X}).", static_cast<unsigned long>(hr)));
                return true;
            }
            return false;
        }

        if (! popup)
        {
            return false;
        }

        FileOperationsPopupInternal::PopupLayoutDebugSnapshot layout{};
        layout.taskId = state.taskA.value();
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, layout))
        {
            return false;
        }
        if (! layout.graphCurrentBandwidthLineVisible || ! layout.graphCurrentBandwidthLabelVisible || ! layout.graphEtaLabelVisible)
        {
            return false;
        }
        if (! layout.graphEtaLabelRightAligned || layout.taskInlineSpeedRowVisible || layout.taskInlineEtaRowVisible ||
            layout.taskExpandedBaseHeightDip != 244.0f)
        {
            Fail(std::format(L"Active transfer card did not use the compact graph-label layout (etaRight={0}, inlineSpeed={1}, inlineEta={2}, baseHeightDip={3}).",
                             layout.graphEtaLabelRightAligned,
                             layout.taskInlineSpeedRowVisible,
                             layout.taskInlineEtaRowVisible,
                             layout.taskExpandedBaseHeightDip));
            return true;
        }
        if (! layout.footerPauseResumeAllVisible || ! layout.footerPauseResumeAllPauses)
        {
            Fail(std::format(L"Popup footer should expose Pause all for an active copy (visible={}, pauses={}).",
                             layout.footerPauseResumeAllVisible,
                             layout.footerPauseResumeAllPauses));
            return true;
        }

        FileOperationsPopupInternal::PopupSelfTestInvoke pauseAll{};
        pauseAll.kind = FileOperationsPopupInternal::PopupHitTest::Kind::FooterPauseResumeAll;
        pauseAll.data = 1u;
        if (! DebugInvokeFileOperationsPopup(popup, pauseAll))
        {
            Fail(L"Failed to invoke file-operations footer Pause all.");
            return true;
        }
        AppendLog(L"Phase6_PopupSmokeResizeAndPause invoked Pause all");
        state.markerTick = nowTick;
        state.stepState  = 2;
        return false;
    }

    auto* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    if (! task)
    {
        if (state.taskA.has_value())
        {
            const auto it = state.completedTasks.find(state.taskA.value());
            if (it != state.completedTasks.end() && state.stepState == 4 && SUCCEEDED(it->second.hr))
            {
                // The copy may finish immediately after Resume. Reaching step 4
                // proves that pause, resized layout, and resume were all observed.
                if (popup && state.popupOriginalRectValid)
                {
                    const int width  = state.popupOriginalRect.right - state.popupOriginalRect.left;
                    const int height = state.popupOriginalRect.bottom - state.popupOriginalRect.top;
                    SetWindowPos(popup, nullptr, state.popupOriginalRect.left, state.popupOriginalRect.top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
                }
                state.stepState = 6;
            }
            else if (it != state.completedTasks.end() && state.stepState < 6)
            {
                Fail(std::format(L"Copy task completed before pause/resize validation finished (hr=0x{:08X}, stepState={}, popup={}, originalRect={}).",
                                 static_cast<unsigned long>(it->second.hr),
                                 state.stepState,
                                 popup != nullptr,
                                 state.popupOriginalRectValid));
                return true;
            }
        }
        if (state.stepState < 6)
        {
            return false;
        }
    }

    if (state.stepState == 2)
    {
        if (! popup || ! state.popupOriginalRectValid)
        {
            return false;
        }

        if (nowTick >= state.markerTick && (nowTick - state.markerTick) < 500ull)
        {
            return false;
        }

        if (! task->IsPaused())
        {
            return false;
        }

        AppendLog(L"Phase6_PopupSmokeResizeAndPause observed paused task");
        const int height = state.popupOriginalRect.bottom - state.popupOriginalRect.top;
        SetWindowPos(popup, nullptr, state.popupOriginalRect.left, state.popupOriginalRect.top, 420, height, SWP_NOZORDER | SWP_NOACTIVATE);

        AppendLog(L"Phase6_PopupSmokeResizeAndPause resized popup");
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (! popup)
        {
            return false;
        }

        FileOperationsPopupInternal::PopupLayoutDebugSnapshot layout{};
        layout.taskId = state.taskA.value();
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, layout))
        {
            return false;
        }
        if (! layout.footerPauseResumeAllVisible || layout.footerPauseResumeAllPauses)
        {
            Fail(std::format(L"Popup footer should expose Resume all after bulk pause (visible={}, pauses={}).",
                             layout.footerPauseResumeAllVisible,
                             layout.footerPauseResumeAllPauses));
            return true;
        }

        {
            std::scoped_lock lock(task->_progressMutex);
            state.pauseResumeBaselineItems = task->_progressCompletedItems;
            state.pauseResumeBaselineBytes = task->_progressCompletedBytes;
        }

        FileOperationsPopupInternal::PopupSelfTestInvoke resumeAll{};
        resumeAll.kind = FileOperationsPopupInternal::PopupHitTest::Kind::FooterPauseResumeAll;
        resumeAll.data = 2u;
        if (! DebugInvokeFileOperationsPopup(popup, resumeAll))
        {
            Fail(L"Failed to invoke file-operations footer Resume all.");
            return true;
        }
        AppendLog(std::format(L"Phase6_PopupSmokeResizeAndPause invoked Resume all baselineItems={} baselineBytes={}",
                              state.pauseResumeBaselineItems,
                              state.pauseResumeBaselineBytes));
        state.markerTick = nowTick;
        state.stepState  = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        if (nowTick >= state.markerTick && (nowTick - state.markerTick) < 500ull)
        {
            return false;
        }

        if (task->IsPaused())
        {
            return false;
        }

        unsigned long completedItems = 0;
        uint64_t completedBytes      = 0;
        {
            std::scoped_lock lock(task->_progressMutex);
            completedItems = task->_progressCompletedItems;
            completedBytes = task->_progressCompletedBytes;
        }
        if (completedItems <= state.pauseResumeBaselineItems && completedBytes <= state.pauseResumeBaselineBytes)
        {
            return false;
        }

        AppendLog(std::format(L"Phase6_PopupSmokeResizeAndPause observed resumed progress items={} bytes={}", completedItems, completedBytes));
        if (popup && state.popupOriginalRectValid)
        {
            const int width  = state.popupOriginalRect.right - state.popupOriginalRect.left;
            const int height = state.popupOriginalRect.bottom - state.popupOriginalRect.top;
            SetWindowPos(popup, nullptr, state.popupOriginalRect.left, state.popupOriginalRect.top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
        }

        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        AppendLog(L"Phase6_PopupSmokeResizeAndPause requesting cancel");
        task->RequestCancel();
        state.stepState = 6;
        return false;
    }

    if (state.completedTasks.find(state.taskA.value()) == state.completedTasks.end())
    {
        return false;
    }

    NextStep(state, SelfTestState::Step::Phase6_DeleteBytesMeaningful);
    return false;
}
case SelfTestState::Step::Phase6_DeleteBytesMeaningful:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase6_DeleteBytesMeaningful timed out.");
        return true;
    }

    const std::filesystem::path deleteTree = state.tempRoot / L"delete-tree";
    if (state.stepState == 0)
    {
        FileOperationsPopupInternal::TaskSnapshot completedDelete{};
        completedDelete.operation      = FILESYSTEM_DELETE;
        completedDelete.finished       = true;
        completedDelete.resultHr       = S_OK;
        completedDelete.totalItems     = 4;
        completedDelete.completedItems = 2;
        completedDelete.totalBytes     = 100;
        completedDelete.completedBytes = 50;

        const float completedFraction = DebugComputeFileOperationsTaskCompleteFraction(completedDelete);
        if (completedFraction != 1.0f)
        {
            Fail(std::format(L"Completed delete popup progress should display 100%; got {:.3f}.", completedFraction));
            return true;
        }

        if (! std::filesystem::exists(deleteTree))
        {
            Fail(L"Delete-tree folder missing before delete-bytes test.");
            return true;
        }

        const FileSystemFlags flags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

        state.taskA =
            StartFileOperationAndGetId(state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {deleteTree}, {}, flags, false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start delete task for delete-bytes validation.");
            return true;
        }

        state.markerTick = 0;
        state.stepState  = 1;
        return false;
    }

    const uint64_t deleteTaskId = state.taskA.value();

    // Keep observing progress while the task exists (it is removed immediately on completion).
    if (auto* task = state.fileOps->FindTask(deleteTaskId))
    {
        const bool discoveryDone = task->_discoveryClosed.load(std::memory_order_acquire);
        const uint64_t total   = task->_discoveredTotalBytes.load(std::memory_order_acquire);
        if (discoveryDone && total > 0)
        {
            state.markerTick |= 1ull;
        }

        uint64_t completedBytes = 0;
        {
            std::scoped_lock lock(task->_progressMutex);
            completedBytes = task->_progressCompletedBytes;
        }

        if (task->HasStarted() && completedBytes > 0)
        {
            state.markerTick |= 2ull;
        }
    }

    const auto completionIt = state.completedTasks.find(deleteTaskId);
    if (completionIt == state.completedTasks.end())
    {
        return false;
    }

    const CompletedTaskInfo& completion = completionIt->second;
    if (completion.discoveryClosed && completion.discoveredTotalBytes > 0)
    {
        state.markerTick |= 1ull;
    }
    if (completion.started && completion.progressCompletedBytes > 0)
    {
        state.markerTick |= 2ull;
    }

    if (std::filesystem::exists(deleteTree))
    {
        Fail(L"Delete-tree folder still exists after delete task completed.");
        return true;
    }

    if ((state.markerTick & 1ull) == 0)
    {
        Fail(L"Delete-bytes validation failed: did not observe non-zero discovered total bytes at traversal closure.");
        return true;
    }

    if ((state.markerTick & 2ull) == 0)
    {
        Fail(L"Delete-bytes validation failed: did not observe delete completedBytes > 0 (check delete progress reporting).");
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase6_LocalBandwidthThrottle);
    return false;
}
case SelfTestState::Step::Phase6_LocalBandwidthThrottle:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"Phase6_LocalBandwidthThrottle timed out.");
        return true;
    }

    constexpr uint64_t kDurationFileBytes          = 16ull * 1024ull * 1024ull;
    constexpr uint64_t kDurationSpeedLimitBytes    = 4ull * 1024ull * 1024ull;
    constexpr uint64_t kCancelSpeedLimitBytes      = 1ull * 1024ull * 1024ull;
    constexpr ULONGLONG kCancelLatencyThresholdMs  = 250ull;
    constexpr ULONGLONG kThroughputWindowMs        = 1000ull;
    const std::filesystem::path srcDir             = state.tempRoot / L"phase6-bandwidth-src";
    const std::filesystem::path durationDstDir     = state.tempRoot / L"phase6-bandwidth-dst";
    const std::filesystem::path cancelDstDir       = state.tempRoot / L"phase6-bandwidth-cancel";
    const std::filesystem::path srcFile            = srcDir / L"payload.bin";
    const std::filesystem::path durationCopiedFile = durationDstDir / L"payload.bin";
    const FileSystemFlags flags =
        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    const auto appendBandwidthSample = [&](ULONGLONG tick, uint64_t completedBytes) noexcept
    {
        if (! state.localBandwidthSamples.empty())
        {
            const auto& last = state.localBandwidthSamples.back();
            if (last.first == tick && last.second == completedBytes)
            {
                return;
            }

            if (completedBytes >= last.second)
            {
                state.localBandwidthMaxSampleDeltaBytes = (std::max)(state.localBandwidthMaxSampleDeltaBytes, completedBytes - last.second);
            }
        }

        state.localBandwidthSamples.emplace_back(tick, completedBytes);
    };

    if (state.stepState == 0)
    {
        AppendLog(L"Phase6_LocalBandwidthThrottle setup begin");
        state.taskA.reset();
        state.taskB.reset();
        state.localBandwidthRunStartTick    = 0;
        state.localBandwidthCancelStartTick = 0;
        state.localBandwidthDurationUs      = 0;
        state.localBandwidthDurationLeadUs  = 0;
        state.localBandwidthCancelLatencyUs = 0;
        state.localBandwidthMaxWindowBytes  = 0;
        state.localBandwidthSamples.clear();

        if (! RecreateEmptyDirectory(srcDir))
        {
            Fail(L"Failed to reset phase6-bandwidth source directory.");
            return true;
        }
        AppendLog(L"Phase6_LocalBandwidthThrottle source directory reset");
        if (! RecreateEmptyDirectory(durationDstDir))
        {
            Fail(L"Failed to reset phase6-bandwidth duration destination directory.");
            return true;
        }
        AppendLog(L"Phase6_LocalBandwidthThrottle duration destination reset");
        if (! RecreateEmptyDirectory(cancelDstDir))
        {
            Fail(L"Failed to reset phase6-bandwidth cancel destination directory.");
            return true;
        }
        AppendLog(L"Phase6_LocalBandwidthThrottle cancel destination reset");

        AppendLog(L"Phase6_LocalBandwidthThrottle source seed begin");
        if (! WriteTestFile(srcFile, kDurationFileBytes))
        {
            Fail(L"Failed to seed source file for local bandwidth throttle validation.");
            return true;
        }
        AppendLog(L"Phase6_LocalBandwidthThrottle source seed complete");

        AppendLog(L"Phase6_LocalBandwidthThrottle duration task start begin");
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFile},
                                                 durationDstDir,
                                                 flags,
                                                 false,
                                                 kDurationSpeedLimitBytes);
        AppendLog(L"Phase6_LocalBandwidthThrottle duration task start returned");
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start local bandwidth throttle duration task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        auto* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (task && task->HasStarted() && state.localBandwidthRunStartTick == 0)
        {
            state.localBandwidthRunStartTick = nowTick;
        }
        if (task && state.localBandwidthRunStartTick != 0)
        {
            uint64_t completedBytes = 0;
            {
                std::scoped_lock lock(task->_progressMutex);
                completedBytes = task->_progressCompletedBytes;
            }
            appendBandwidthSample(nowTick, completedBytes);
        }

        const auto completionIt = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completionIt == state.completedTasks.end())
        {
            return false;
        }

        const CompletedTaskInfo& completion = completionIt->second;
        if (FAILED(completion.hr))
        {
            Fail(std::format(L"Local bandwidth throttle duration task failed: 0x{:08X}.", static_cast<unsigned long>(completion.hr)));
            return true;
        }

        std::error_code durationEc;
        const bool durationExists = std::filesystem::exists(durationCopiedFile, durationEc);
        const uint64_t durationFileSize =
            durationExists && ! durationEc ? static_cast<uint64_t>(std::filesystem::file_size(durationCopiedFile, durationEc)) : 0ull;
        if (! durationExists || durationEc || durationFileSize != kDurationFileBytes)
        {
            Fail(L"Local bandwidth throttle duration task produced the wrong output size.");
            return true;
        }

        if (state.localBandwidthRunStartTick == 0 || completion.completionTick < state.localBandwidthRunStartTick)
        {
            Fail(L"Local bandwidth throttle duration task did not record a valid runtime window.");
            return true;
        }

        state.localBandwidthDurationUs     = static_cast<uint64_t>(completion.completionTick - state.localBandwidthRunStartTick) * 1000ull;
        const uint64_t idealDurationUs     = (kDurationFileBytes * 1000000ull) / kDurationSpeedLimitBytes;
        state.localBandwidthDurationLeadUs = (idealDurationUs > state.localBandwidthDurationUs) ? (idealDurationUs - state.localBandwidthDurationUs) : 0ull;
        appendBandwidthSample(completion.completionTick, completion.progressCompletedBytes);

        for (size_t begin = 0, end = 0; begin < state.localBandwidthSamples.size(); ++begin)
        {
            const ULONGLONG windowStartTick = state.localBandwidthSamples[begin].first;
            while (end + 1 < state.localBandwidthSamples.size() && state.localBandwidthSamples[end + 1].first <= (windowStartTick + kThroughputWindowMs))
            {
                ++end;
            }

            const uint64_t windowStartBytes = state.localBandwidthSamples[begin].second;
            const uint64_t windowEndBytes   = state.localBandwidthSamples[end].second;
            if (windowEndBytes >= windowStartBytes)
            {
                state.localBandwidthMaxWindowBytes = (std::max)(state.localBandwidthMaxWindowBytes, windowEndBytes - windowStartBytes);
            }
        }

        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcFile},
                                                 cancelDstDir,
                                                 flags,
                                                 false,
                                                 kCancelSpeedLimitBytes);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start local bandwidth throttle cancel task.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        auto* task = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
        if (! task)
        {
            return false;
        }

        if (! task->HasStarted())
        {
            return false;
        }

        task->RequestCancel();
        state.localBandwidthCancelStartTick = nowTick;
        state.stepState                     = 3;
        return false;
    }

    const auto completionIt = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
    if (completionIt == state.completedTasks.end())
    {
        return false;
    }

    const CompletedTaskInfo& completion = completionIt->second;
    if (completion.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) && completion.hr != E_ABORT)
    {
        Fail(std::format(L"Local bandwidth throttle cancel task expected cancel hr, got 0x{:08X}.", static_cast<unsigned long>(completion.hr)));
        return true;
    }

    if (state.localBandwidthCancelStartTick == 0 || completion.completionTick < state.localBandwidthCancelStartTick)
    {
        Fail(L"Local bandwidth throttle cancel task did not record a valid cancel marker.");
        return true;
    }

    state.localBandwidthCancelLatencyUs = static_cast<uint64_t>(completion.completionTick - state.localBandwidthCancelStartTick) * 1000ull;

    const uint64_t idealDurationUs = (kDurationFileBytes * 1000000ull) / kDurationSpeedLimitBytes;
    const std::wstring detail      = std::format(L"fileBytes={} durationLimit={} cancelLimit={} source={} durationDestination={} cancelDestination={}",
                                                 kDurationFileBytes,
                                                 kDurationSpeedLimitBytes,
                                                 kCancelSpeedLimitBytes,
                                                 srcFile.wstring(),
                                                 durationDstDir.wstring(),
                                                 cancelDstDir.wstring());

    Debug::Perf::Emit(
        L"FileOps.SelfTest.LocalBandwidthThrottleDuration", detail, state.localBandwidthDurationUs, kDurationFileBytes, kDurationSpeedLimitBytes, S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.LocalBandwidthThrottleDurationLead",
                      detail,
                      state.localBandwidthDurationLeadUs,
                      idealDurationUs,
                      state.localBandwidthDurationUs,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.LocalBandwidthThrottleMaxWindowBytes",
                      detail,
                      state.localBandwidthMaxWindowBytes,
                      kThroughputWindowMs,
                      kDurationSpeedLimitBytes,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.LocalBandwidthThrottleMaxSampleDeltaBytes",
                      detail,
                      state.localBandwidthMaxSampleDeltaBytes,
                      state.localBandwidthSamples.size(),
                      kDurationSpeedLimitBytes,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.LocalBandwidthThrottleCancelLatency",
                      detail,
                      state.localBandwidthCancelLatencyUs,
                      kCancelLatencyThresholdMs,
                      kCancelSpeedLimitBytes,
                      completion.hr);

    AppendLog(std::format(L"Phase6_LocalBandwidthThrottle duration={}us ideal={}us lead={}us maxWindowBytes={} maxSampleDeltaBytes={} cancelLatency={}us",
                          state.localBandwidthDurationUs,
                          idealDurationUs,
                          state.localBandwidthDurationLeadUs,
                          state.localBandwidthMaxWindowBytes,
                          state.localBandwidthMaxSampleDeltaBytes,
                          state.localBandwidthCancelLatencyUs));

    if (state.localBandwidthDurationUs == 0)
    {
        Fail(L"Local bandwidth throttle validation captured zero-valued timing.");
        return true;
    }

    const uint64_t minExpectedDurationUs = (idealDurationUs * 70ull) / 100ull;
    if (state.localBandwidthDurationUs < minExpectedDurationUs)
    {
        Fail(std::format(L"Local bandwidth throttle duration completed too quickly: observed={}us expectedAtLeast={}us.",
                         state.localBandwidthDurationUs,
                         minExpectedDurationUs));
        return true;
    }

    constexpr uint64_t kMaxDurationLeadUs = 250'000ull;
    if (state.localBandwidthDurationLeadUs > kMaxDurationLeadUs)
    {
        Fail(std::format(L"Local bandwidth throttle overshot the configured rate too much: lead={}us maxAllowed={}us.",
                         state.localBandwidthDurationLeadUs,
                         kMaxDurationLeadUs));
        return true;
    }

    const uint64_t maxAllowedWindowBytes        = (kDurationSpeedLimitBytes * 110ull) / 100ull;
    const uint64_t maxAllowedSampledWindowBytes = state.localBandwidthMaxSampleDeltaBytes > std::numeric_limits<uint64_t>::max() - maxAllowedWindowBytes
                                                      ? std::numeric_limits<uint64_t>::max()
                                                      : maxAllowedWindowBytes + state.localBandwidthMaxSampleDeltaBytes;
    if (state.localBandwidthMaxWindowBytes > maxAllowedSampledWindowBytes)
    {
        Fail(std::format(L"Local bandwidth throttle exceeded the 1-second window budget beyond one sampled progress quantum: observed={}B "
                         L"maxAllowed={}B sampleQuantum={}B.",
                         state.localBandwidthMaxWindowBytes,
                         maxAllowedWindowBytes,
                         state.localBandwidthMaxSampleDeltaBytes));
        return true;
    }

    if (state.localBandwidthCancelLatencyUs > (kCancelLatencyThresholdMs * 1000ull))
    {
        Fail(std::format(L"Local bandwidth throttle cancel latency too high: observed={}us threshold={}us.",
                         state.localBandwidthCancelLatencyUs,
                         kCancelLatencyThresholdMs * 1000ull));
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase6_ParallelBandwidthThrottleFairness);
    return false;
}
case SelfTestState::Step::Phase6_ParallelBandwidthThrottleFairness:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase6_ParallelBandwidthThrottleFairness timed out.");
        return true;
    }

    constexpr std::wstring_view kSharedMode        = L"shared";
    constexpr std::wstring_view kPerWorkerMode     = L"perworker";
    constexpr unsigned int kCopyConcurrency        = 4u;
    constexpr int kFileCount                       = 4;
    constexpr uint64_t kFileBytes                  = 8ull * 1024ull * 1024ull;
    constexpr uint64_t kSpeedLimitBytesPerSecond      = 4ull * 1024ull * 1024ull;
    constexpr uint64_t kAllowedDurationSlowdownUs     = 250'000ull;
    constexpr uint64_t kAllowedSkewRegressionBytes    = 1ull * 1024ull * 1024ull;
    // A host snapshot cannot resolve skew below one callback publication. Keep the floor tied to
    // the measured callback quantum; larger scheduler skew remains a failure.
    constexpr uint64_t kObservableSkewCallbackQuanta = 1ull;
    constexpr uint64_t kMinSamplesPerRun              = 4ull;
    const std::filesystem::path srcDir                = state.tempRoot / L"phase6-parallel-bandwidth-src";
    const std::filesystem::path sharedDstDir          = state.tempRoot / L"phase6-parallel-bandwidth-shared-dst";
    const std::filesystem::path perWorkerDstDir       = state.tempRoot / L"phase6-parallel-bandwidth-perworker-dst";

    const auto ensureWorkerModeBackup = [&]() noexcept
    {
        if (state.bandwidthThrottleWorkerModeEnvBackedUp)
        {
            return;
        }

        SetLastError(ERROR_SUCCESS);
        const DWORD required                            = GetEnvironmentVariableW(kSelfTestEnvBandwidthThrottleWorkerMode.data(), nullptr, 0);
        const DWORD error                               = GetLastError();
        state.bandwidthThrottleWorkerModeEnvHadOriginal = ! (required == 0 && error == ERROR_ENVVAR_NOT_FOUND);
        state.bandwidthThrottleWorkerModeEnvOriginal    = GetEnvVarTrimmed(kSelfTestEnvBandwidthThrottleWorkerMode);
        state.bandwidthThrottleWorkerModeEnvBackedUp    = true;
    };

    const auto setWorkerMode = [&](std::wstring_view mode, std::wstring_view label) noexcept -> bool
    {
        std::wstring raw(mode);
        if (! SetEnvironmentVariableW(kSelfTestEnvBandwidthThrottleWorkerMode.data(), raw.c_str()))
        {
            Fail(std::format(L"Failed to set bandwidth worker mode to {} for {}.", mode, label));
            return false;
        }

        return true;
    };

    const auto restoreOriginalWorkerMode = [&]() noexcept -> bool
    {
        if (! state.bandwidthThrottleWorkerModeEnvBackedUp)
        {
            return true;
        }

        const BOOL restored =
            state.bandwidthThrottleWorkerModeEnvHadOriginal
                ? SetEnvironmentVariableW(kSelfTestEnvBandwidthThrottleWorkerMode.data(), state.bandwidthThrottleWorkerModeEnvOriginal.c_str())
                : SetEnvironmentVariableW(kSelfTestEnvBandwidthThrottleWorkerMode.data(), nullptr);
        if (! restored)
        {
            Fail(L"Failed to restore the original bandwidth worker mode override.");
            return false;
        }

        state.bandwidthThrottleWorkerModeEnvBackedUp    = false;
        state.bandwidthThrottleWorkerModeEnvHadOriginal = false;
        state.bandwidthThrottleWorkerModeEnvOriginal.clear();
        return true;
    };

    const auto applyCopyConfig = [&]() noexcept -> bool
    {
        const std::string config = std::format(
            R"json({{"concurrencyMode":"manual","copyMoveMaxConcurrency":{},"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"recycleBinBatchSize":500,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4}})json",
            kCopyConcurrency);
        if (! SetPluginConfiguration(state.infoLocal.get(), config))
        {
            Fail(L"Failed to apply local plugin config for parallel bandwidth fairness.");
            return false;
        }

        state.localConfigDirty = true;
        return true;
    };

    const auto seedSourceDir = [&](const std::filesystem::path& root) noexcept -> bool
    {
        if (! RecreateEmptyDirectory(root))
        {
            Fail(L"Failed to reset the parallel bandwidth fairness source directory.");
            return false;
        }

        for (int i = 0; i < kFileCount; ++i)
        {
            const std::filesystem::path file = root / std::format(L"payload_{:02}.bin", i);
            if (! WriteTestFile(file, kFileBytes))
            {
                Fail(std::format(L"Failed to write fairness source file {}.", file.native()));
                return false;
            }
        }

        return true;
    };

    const auto startCopy = [&](const std::filesystem::path& dstDir, std::optional<std::uint64_t>& taskSlot, std::wstring_view label) noexcept -> bool
    {
        if (! RecreateEmptyDirectory(dstDir))
        {
            Fail(std::format(L"Failed to reset {} destination.", label));
            return false;
        }

        std::vector<std::filesystem::path> sources = CollectFiles(srcDir, 64u);
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
                                                                 dstDir,
                                                                 flags,
                                                                 false,
                                                                 kSpeedLimitBytesPerSecond);
        if (! taskSlot.has_value())
        {
            Fail(std::format(L"Failed to start {}.", label));
            return false;
        }

        return true;
    };

    const auto sampleTask = [&](FolderWindow::FileOperationState::Task& task,
                                uint64_t& maxSkewBytes,
                                uint64_t& maxCallbackDeltaBytes,
                                size_t& maxActiveStreams,
                                uint64_t& sampleCount) noexcept
    {
        {
            std::scoped_lock lock(task._inFlightFilesMutex);

            size_t activeCount   = 0;
            uint64_t minBytes    = 0;
            uint64_t maxBytes    = 0;
            bool haveActiveBytes = false;
            for (size_t i = 0; i < task._inFlightFileCount; ++i)
            {
                const auto& entry = task._inFlightFiles[i];
                if (entry.totalBytes == 0 || entry.completedBytes >= entry.totalBytes)
                {
                    continue;
                }

                ++activeCount;
                if (! haveActiveBytes)
                {
                    minBytes        = entry.completedBytes;
                    maxBytes        = entry.completedBytes;
                    haveActiveBytes = true;
                }
                else
                {
                    minBytes = (std::min)(minBytes, entry.completedBytes);
                    maxBytes = (std::max)(maxBytes, entry.completedBytes);
                }
            }

            if (activeCount > 0)
            {
                ++sampleCount;
                maxActiveStreams = (std::max)(maxActiveStreams, activeCount);
            }

            if (activeCount > 1 && haveActiveBytes && maxBytes >= minBytes)
            {
                maxSkewBytes = (std::max)(maxSkewBytes, maxBytes - minBytes);
            }
        }

        {
            std::scoped_lock lock(task._progressStreamPerfMutex);
            for (size_t i = 0; i < task._progressStreamPerfCount; ++i)
            {
                maxCallbackDeltaBytes = (std::max)(maxCallbackDeltaBytes, task._progressStreamPerf[i].maxCallbackDeltaBytes);
            }
        }
    };

    const auto finalizeCopy = [&](const std::optional<std::uint64_t>& taskSlot,
                                  const std::filesystem::path& dstDir,
                                  ULONGLONG runStartTick,
                                  uint64_t& durationOut,
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
            Fail(std::format(L"{} failed with hr=0x{:08X}.", label, static_cast<unsigned long>(completion.hr)));
            return -1;
        }

        if (runStartTick == 0 || completion.completionTick < runStartTick)
        {
            Fail(std::format(L"{} did not record a valid runtime window.", label));
            return -1;
        }

        std::vector<std::filesystem::path> outputs = CollectFiles(dstDir, 64u);
        if (outputs.size() != static_cast<size_t>(kFileCount))
        {
            Fail(std::format(L"{} expected {} copied files, got {}.", label, kFileCount, outputs.size()));
            return -1;
        }

        for (const std::filesystem::path& output : outputs)
        {
            std::error_code ec;
            const uint64_t fileBytes = std::filesystem::exists(output, ec) && ! ec ? static_cast<uint64_t>(std::filesystem::file_size(output, ec)) : 0ull;
            if (ec || fileBytes != kFileBytes)
            {
                Fail(std::format(L"{} produced an unexpected output size for {}.", label, output.native()));
                return -1;
            }
        }

        durationOut = static_cast<uint64_t>(completion.completionTick - runStartTick) * 1000ull;
        return 1;
    };

    if (state.stepState == 0)
    {
        state.taskA.reset();
        state.taskB.reset();
        state.parallelBandwidthRunStartTick          = 0;
        state.parallelBandwidthBaselineUs            = 0;
        state.parallelBandwidthCandidateUs           = 0;
        state.parallelBandwidthBaselineMaxSkewBytes  = 0;
        state.parallelBandwidthCandidateMaxSkewBytes = 0;
        state.parallelBandwidthBaselineMaxCallbackDeltaBytes  = 0;
        state.parallelBandwidthCandidateMaxCallbackDeltaBytes = 0;
        state.parallelBandwidthBaselineMaxActive     = 0;
        state.parallelBandwidthCandidateMaxActive    = 0;
        state.parallelBandwidthBaselineSamples       = 0;
        state.parallelBandwidthCandidateSamples      = 0;

        ensureWorkerModeBackup();
        if (! applyCopyConfig() || ! seedSourceDir(srcDir) || ! setWorkerMode(kSharedMode, L"shared-only baseline") ||
            ! startCopy(sharedDstDir, state.taskA, L"shared-only parallel bandwidth baseline"))
        {
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        auto* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (task && task->HasStarted() && state.parallelBandwidthRunStartTick == 0)
        {
            state.parallelBandwidthRunStartTick = nowTick;
        }
        if (task && state.parallelBandwidthRunStartTick != 0)
        {
            sampleTask(*task,
                       state.parallelBandwidthBaselineMaxSkewBytes,
                       state.parallelBandwidthBaselineMaxCallbackDeltaBytes,
                       state.parallelBandwidthBaselineMaxActive,
                       state.parallelBandwidthBaselineSamples);
        }

        const int finalize = finalizeCopy(
            state.taskA, sharedDstDir, state.parallelBandwidthRunStartTick, state.parallelBandwidthBaselineUs, L"shared-only parallel bandwidth baseline");
        if (finalize < 0)
        {
            return true;
        }
        if (finalize == 0)
        {
            return false;
        }

        state.parallelBandwidthRunStartTick = 0;
        if (! setWorkerMode(kPerWorkerMode, L"per-worker candidate") || ! startCopy(perWorkerDstDir, state.taskB, L"per-worker parallel bandwidth candidate"))
        {
            return true;
        }

        state.stepState = 2;
        return false;
    }

    auto* task = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
    if (task && task->HasStarted() && state.parallelBandwidthRunStartTick == 0)
    {
        state.parallelBandwidthRunStartTick = nowTick;
    }
    if (task && state.parallelBandwidthRunStartTick != 0)
    {
        sampleTask(*task,
                   state.parallelBandwidthCandidateMaxSkewBytes,
                   state.parallelBandwidthCandidateMaxCallbackDeltaBytes,
                   state.parallelBandwidthCandidateMaxActive,
                   state.parallelBandwidthCandidateSamples);
    }

    const int finalize = finalizeCopy(
        state.taskB, perWorkerDstDir, state.parallelBandwidthRunStartTick, state.parallelBandwidthCandidateUs, L"per-worker parallel bandwidth candidate");
    if (finalize < 0)
    {
        return true;
    }
    if (finalize == 0)
    {
        return false;
    }

    if (! restoreOriginalWorkerMode())
    {
        return true;
    }

    if (! state.localConfigOriginal.empty() && ! SetPluginConfiguration(state.infoLocal.get(), state.localConfigOriginal))
    {
        Fail(L"Failed to restore local plugin config after parallel bandwidth fairness validation.");
        return true;
    }
    state.localConfigDirty = false;

    const uint64_t skewImprovementBytes = (state.parallelBandwidthBaselineMaxSkewBytes > state.parallelBandwidthCandidateMaxSkewBytes)
                                              ? (state.parallelBandwidthBaselineMaxSkewBytes - state.parallelBandwidthCandidateMaxSkewBytes)
                                              : 0ull;
    const std::wstring detail           = std::format(
        L"fileCount={} fileBytes={} speedLimitBps={} copyConcurrency={} sharedMaxActive={} perWorkerMaxActive={} sharedSamples={} perWorkerSamples={} "
        L"sharedMaxCallbackDeltaBytes={} perWorkerMaxCallbackDeltaBytes={}",
        kFileCount,
        kFileBytes,
        kSpeedLimitBytesPerSecond,
        kCopyConcurrency,
        state.parallelBandwidthBaselineMaxActive,
        state.parallelBandwidthCandidateMaxActive,
        state.parallelBandwidthBaselineSamples,
        state.parallelBandwidthCandidateSamples,
        state.parallelBandwidthBaselineMaxCallbackDeltaBytes,
        state.parallelBandwidthCandidateMaxCallbackDeltaBytes);
    Debug::Perf::Emit(L"FileOps.SelfTest.ParallelBandwidthSharedOnlyDuration",
                      detail,
                      state.parallelBandwidthBaselineUs,
                      state.parallelBandwidthBaselineMaxActive,
                      state.parallelBandwidthBaselineMaxSkewBytes,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.ParallelBandwidthPerWorkerDuration",
                      detail,
                      state.parallelBandwidthCandidateUs,
                      state.parallelBandwidthCandidateMaxActive,
                      state.parallelBandwidthCandidateMaxSkewBytes,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.ParallelBandwidthSkewImprovement",
                      detail,
                      skewImprovementBytes,
                      state.parallelBandwidthBaselineMaxSkewBytes,
                      state.parallelBandwidthCandidateMaxSkewBytes,
                      S_OK);

    AppendLog(
        std::format(L"Phase6_ParallelBandwidthThrottleFairness shared={}us perWorker={}us sharedSkew={}B perWorkerSkew={}B sharedActive={} "
                    L"perWorkerActive={} sharedMaxCallbackDelta={}B perWorkerMaxCallbackDelta={}B",
                    state.parallelBandwidthBaselineUs,
                    state.parallelBandwidthCandidateUs,
                    state.parallelBandwidthBaselineMaxSkewBytes,
                    state.parallelBandwidthCandidateMaxSkewBytes,
                    state.parallelBandwidthBaselineMaxActive,
                    state.parallelBandwidthCandidateMaxActive,
                    state.parallelBandwidthBaselineMaxCallbackDeltaBytes,
                    state.parallelBandwidthCandidateMaxCallbackDeltaBytes));

    if (state.parallelBandwidthBaselineUs == 0 || state.parallelBandwidthCandidateUs == 0)
    {
        Fail(L"Parallel bandwidth fairness validation captured zero-valued timing.");
        return true;
    }

    if (state.parallelBandwidthBaselineMaxActive < 2u || state.parallelBandwidthCandidateMaxActive < 2u)
    {
        Fail(std::format(L"Parallel bandwidth fairness did not observe enough concurrent streams: shared={} perWorker={}.",
                         state.parallelBandwidthBaselineMaxActive,
                         state.parallelBandwidthCandidateMaxActive));
        return true;
    }

    if (state.parallelBandwidthBaselineSamples < kMinSamplesPerRun || state.parallelBandwidthCandidateSamples < kMinSamplesPerRun)
    {
        Fail(std::format(L"Parallel bandwidth fairness captured too few samples: shared={} perWorker={} required={}.",
                         state.parallelBandwidthBaselineSamples,
                         state.parallelBandwidthCandidateSamples,
                         kMinSamplesPerRun));
        return true;
    }

    const uint64_t maxAllowedCandidateDurationUs = state.parallelBandwidthBaselineUs + (state.parallelBandwidthBaselineUs / 5ull) + kAllowedDurationSlowdownUs;
    if (state.parallelBandwidthCandidateUs > maxAllowedCandidateDurationUs)
    {
        Fail(std::format(L"Per-worker bandwidth fairness slowed down too much: candidate={}us maxAllowed={}us baseline={}us.",
                         state.parallelBandwidthCandidateUs,
                         maxAllowedCandidateDurationUs,
                         state.parallelBandwidthBaselineUs));
        return true;
    }

    const uint64_t maxCallbackDeltaBytes =
        (std::max)(state.parallelBandwidthBaselineMaxCallbackDeltaBytes, state.parallelBandwidthCandidateMaxCallbackDeltaBytes);
    const uint64_t observableSkewFloorBytes =
        maxCallbackDeltaBytes > (std::numeric_limits<uint64_t>::max)() / kObservableSkewCallbackQuanta
            ? (std::numeric_limits<uint64_t>::max)()
            : maxCallbackDeltaBytes * kObservableSkewCallbackQuanta;
    const uint64_t comparativeSkewLimitBytes =
        state.parallelBandwidthBaselineMaxSkewBytes > (std::numeric_limits<uint64_t>::max)() - kAllowedSkewRegressionBytes
            ? (std::numeric_limits<uint64_t>::max)()
            : state.parallelBandwidthBaselineMaxSkewBytes + kAllowedSkewRegressionBytes;
    const uint64_t maxAllowedCandidateSkewBytes = (std::max)(observableSkewFloorBytes, comparativeSkewLimitBytes);
    if (state.parallelBandwidthCandidateMaxSkewBytes > maxAllowedCandidateSkewBytes)
    {
        Fail(std::format(L"Per-worker bandwidth fairness regressed skew too much: candidate={}B maxAllowed={}B baseline={}B observableFloor={}B "
                         L"maxCallbackDelta={}B.",
                         state.parallelBandwidthCandidateMaxSkewBytes,
                         maxAllowedCandidateSkewBytes,
                         state.parallelBandwidthBaselineMaxSkewBytes,
                         observableSkewFloorBytes,
                         maxCallbackDeltaBytes));
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase7_WatcherChurn);
    return false;
}

#endif
