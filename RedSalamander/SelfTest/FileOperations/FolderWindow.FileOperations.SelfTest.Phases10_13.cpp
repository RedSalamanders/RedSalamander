#if defined(FILEOPS_SELFTEST_INCLUDE_PHASE10)

case SelfTestState::Step::FileOps_CrossVolumeMovePartialFailureStatus:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"FileOps_CrossVolumeMovePartialFailureStatus timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"phase4-copy-only-move-source";
    const std::filesystem::path srcFile = srcRoot / L"retained.bin";
    const std::wstring dummyRoot        = L"/phase4-copy-only-move-destination";
    const std::wstring dummyFile        = dummyRoot + L"/retained.bin";

    if (state.stepState == 0)
    {
        // R3-2: this case covers Copy-only Move into a destination without content proof.
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0","writerProof":false})json"));
        if (! RecreateEmptyDirectory(srcRoot) || ! WriteTestFile(srcFile, 256ull * 1024ull) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
        {
            Fail(L"Copy-only Move status test failed to prepare source/destination folders.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
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
            Fail(L"Copy-only Move status test failed to start the move task.");
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
            Fail(std::format(L"Copy-only Move status expected S_FALSE, got 0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        std::error_code ec;
        if (! std::filesystem::exists(srcFile, ec) || ec)
        {
            Fail(L"Copy-only Move status test did not preserve the source file.");
            return true;
        }

        wil::com_ptr<IFileSystemIO> dummyIo;
        unsigned long destinationAttributes = 0u;
        if (! state.fsDummy || FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo ||
            FAILED(dummyIo->GetAttributes(dummyFile.c_str(), &destinationAttributes)) || (destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
        {
            Fail(L"Copy-only Move status test did not publish the destination copy.");
            return true;
        }

        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        const auto summaryIt = std::find_if(
            summaries.begin(), summaries.end(), [&](const auto& summary) noexcept { return state.taskA.has_value() && summary.taskId == state.taskA.value(); });
        if (summaryIt == summaries.end())
        {
            Fail(L"Partial move status test could not find the completed task summary.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.CrossVolumeMovePartialFailureStatus",
                          L"shape=admitted-copy-only-source-kept",
                          completed->second.progressCompletedBytes,
                          summaryIt->issueDiagnostics.size(),
                          1,
                          completed->second.hr);
        NextStep(state, SelfTestState::Step::FileOps_ReparseLiteralCopyKeepsTargets);
        return false;
    }

    return false;
}

case SelfTestState::Step::FileOps_ReparseLiteralCopyKeepsTargets:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        Fail(L"FileOps_ReparseLiteralCopyKeepsTargets timed out.");
        return true;
    }

    static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"preserve"})json"));

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

    // C3: literal Preserve keeps the stored target text, so the copied in-tree junction still names
    // the source tree; nothing is rewritten under the destination.
    const std::wstring expectedInsideTarget = NormalizePathForCompare(std::filesystem::absolute(insideTarget).wstring());
    if (copiedInsideTarget.value() != expectedInsideTarget)
    {
        Fail(std::format(L"Copied in-tree reparse target must keep its stored source target text. expected='{}' actual='{}'.",
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
        Fail(std::format(
            L"Copied out-of-tree reparse target should not be retargeted. expected='{}' actual='{}'.", expectedOutsideTarget, copiedOutsideTarget.value()));
        return true;
    }

    Debug::Perf::Emit(
        L"FileOps.SelfTest.ReparseLiteralCopyKeepsTargets", L"shape=in-tree-literal-plus-out-of-tree-literal", 2, CountFilesRecursive(copiedRoot), 1, S_OK);

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

    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapFired.data(), nullptr));
    if (! SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapPath.data(), victimDir.c_str()) ||
        ! SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapTarget.data(), targetRoot.c_str()))
    {
        Fail(L"Delete TOCTOU guard test failed to arm the debug swap hook.");
        return true;
    }
    auto clearDeleteToctouSwapEnv = wil::scope_exit([]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapPath.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapTarget.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvDeleteToctouSwapFired.data(), nullptr));
    });

    const HRESULT deleteHr = state.fsLocal->DeleteItem(deleteRoot.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), nullptr, nullptr, nullptr);
    if (FAILED(deleteHr))
    {
        Fail(std::format(L"Delete TOCTOU guard test delete failed: 0x{:08X}.", static_cast<unsigned long>(deleteHr)));
        return true;
    }

    if (GetEnvVarTrimmed(kSelfTestEnvDeleteToctouSwapFired) != L"1")
    {
        Fail(L"Delete TOCTOU guard test did not inject the debug dir-to-junction swap.");
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

    Debug::Perf::Emit(
        L"FileOps.SelfTest.DeleteToctouSwapGuard", L"shape=dir-to-reparse-swap-during-parallel-flatten", 1, CountFilesRecursive(targetRoot), 1, S_OK);

    NextStep(state, SelfTestState::Step::FileOps_ResolvedItemsExactDestinations);
    return false;
}

case SelfTestState::Step::FileOps_ResolvedItemsExactDestinations:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"FileOps_ResolvedItemsExactDestinations timed out.");
        return true;
    }

    constexpr size_t kChangedBytes  = 8 * 1024;
    constexpr size_t kKeepBytes     = 4 * 1024;
    constexpr size_t kWholeBytes    = 6 * 1024;
    constexpr size_t kSentinelBytes = 2 * 1024;

    const std::filesystem::path srcRoot        = state.tempRoot / L"phase4-resolved-sync-src";
    const std::filesystem::path dstRoot        = state.tempRoot / L"phase4-resolved-sync-dst";
    const std::filesystem::path refRoot        = state.tempRoot / L"phase4-resolved-sync-ref";
    const std::filesystem::path srcSub         = srcRoot / L"sub";
    const std::filesystem::path dstSub         = dstRoot / L"sub";
    const std::filesystem::path srcChanged     = srcSub / L"changed.bin";
    const std::filesystem::path dstChanged     = dstSub / L"changed.bin";
    const std::filesystem::path srcKeep        = srcSub / L"identical.bin";
    const std::filesystem::path dstKeep        = dstSub / L"identical.bin";
    const std::filesystem::path srcWhole       = srcRoot / L"left-only-dir";
    const std::filesystem::path dstWhole       = dstRoot / L"left-only-dir";
    const std::filesystem::path srcWholeNested = srcWhole / L"nested";
    const std::filesystem::path dstWholeNested = dstWhole / L"nested";
    const std::filesystem::path srcWholeFile   = srcWholeNested / L"payload.bin";
    const std::filesystem::path dstWholeFile   = dstWholeNested / L"payload.bin";
    const std::filesystem::path changedRef     = refRoot / L"changed-ref.bin";
    const std::filesystem::path keepRef        = refRoot / L"keep-ref.bin";
    const std::filesystem::path wholeRef       = refRoot / L"whole-ref.bin";
    const std::filesystem::path sentinel       = dstRoot / L"destination-sentinel.bin";
    const std::filesystem::path sentinelRef    = refRoot / L"destination-sentinel-ref.bin";

    const auto existsNoError = [](const std::filesystem::path& path) noexcept
    {
        std::error_code ec;
        const bool exists = std::filesystem::exists(path, ec);
        return exists && ! ec;
    };

    if (state.stepState == 0)
    {
        state.taskA.reset();
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(refRoot))
        {
            Fail(L"Resolved-items exact destination test failed to reset directories.");
            return true;
        }

        if (! WriteFilledTestFile(srcChanged, kChangedBytes, 0x41) || ! WriteFilledTestFile(changedRef, kChangedBytes, 0x41) ||
            ! WriteFilledTestFile(srcKeep, kKeepBytes, 0x52) || ! WriteFilledTestFile(keepRef, kKeepBytes, 0x52) ||
            ! WriteFilledTestFile(srcWholeFile, kWholeBytes, 0x63) || ! WriteFilledTestFile(wholeRef, kWholeBytes, 0x63) ||
            ! WriteFilledTestFile(sentinel, kSentinelBytes, 0x74) || ! WriteFilledTestFile(sentinelRef, kSentinelBytes, 0x74))
        {
            Fail(L"Resolved-items exact destination test failed to seed files.");
            return true;
        }

        std::vector<FolderWindow::ResolvedFileOperationItem> resolvedItems{
            FolderWindow::ResolvedFileOperationItem{srcChanged, dstChanged, FILESYSTEM_FLAG_NONE, FolderWindow::ResolvedFileOperationItemKind::Normal},
            FolderWindow::ResolvedFileOperationItem{srcSub, dstSub, FILESYSTEM_FLAG_NONE, FolderWindow::ResolvedFileOperationItemKind::DirectoryShell},
            FolderWindow::ResolvedFileOperationItem{
                srcWhole, dstWhole, static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), FolderWindow::ResolvedFileOperationItemKind::Normal},
        };

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {srcChanged, srcSub, srcWhole},
                                                 dstRoot,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 nullptr,
                                                 std::move(resolvedItems));
        if (! state.taskA.has_value())
        {
            Fail(L"Resolved-items exact destination test failed to start move task.");
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

    if (FAILED(completed->second.hr))
    {
        std::wstring itemStatuses;
        for (size_t index = 0; index < completed->second.sourceItemStatuses.size(); ++index)
        {
            if (! itemStatuses.empty())
            {
                itemStatuses.append(L", ");
            }
            itemStatuses.append(std::format(L"{}={}",
                                            index,
                                            completed->second.sourceItemStatuses[index].has_value()
                                                ? std::format(L"0x{:08X}", static_cast<unsigned long>(completed->second.sourceItemStatuses[index].value()))
                                                : L"unset"));
        }
        Fail(std::format(
            L"Resolved-items exact destination move failed: 0x{:08X}; items [{}].", static_cast<unsigned long>(completed->second.hr), itemStatuses));
        return true;
    }

    if (existsNoError(srcChanged))
    {
        Fail(L"Resolved-items exact destination move left the changed file at the source.");
        return true;
    }
    if (! FileSizeEquals(dstChanged, kChangedBytes) || ! FilesEqualBytes(dstChanged, changedRef))
    {
        Fail(L"Resolved-items exact destination move did not place the changed file at its exact destination.");
        return true;
    }
    if (! FileSizeEquals(srcKeep, kKeepBytes) || ! FilesEqualBytes(srcKeep, keepRef))
    {
        Fail(L"Resolved-items exact destination move changed the unlisted identical source file.");
        return true;
    }
    if (existsNoError(dstKeep))
    {
        Fail(L"Resolved-items exact destination move copied an unlisted identical file into the destination shell.");
        return true;
    }
    if (existsNoError(srcWhole))
    {
        Fail(L"Resolved-items exact destination move left the source-only recursive subtree at the source.");
        return true;
    }
    if (! FileSizeEquals(dstWholeFile, kWholeBytes) || ! FilesEqualBytes(dstWholeFile, wholeRef))
    {
        Fail(L"Resolved-items exact destination move did not place the recursive subtree at its exact destination.");
        return true;
    }
    if (! FileSizeEquals(sentinel, kSentinelBytes) || ! FilesEqualBytes(sentinel, sentinelRef))
    {
        Fail(L"Resolved-items exact destination move changed an unrelated destination sentinel.");
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.ResolvedItemsExactDestinations",
                      L"shape=directory-shell-file-and-recursive-subtree",
                      3,
                      CountFilesRecursive(srcRoot),
                      CountFilesRecursive(dstRoot),
                      completed->second.hr);

    NextStep(state, SelfTestState::Step::Fairstream_ManagedMovePreservesUncopiedNewFiles);
    return false;
}

case SelfTestState::Step::Phase10_PermanentDelete:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        ReleaseFileOpsPermanentDeleteBeforeBindPauseForSelfTest();
        SetFileOpsPermanentDeleteBeforeBindPauseForSelfTest(false);
        ReleaseFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest();
        SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(false);
        Fail(L"Phase10_PermanentDelete timed out.");
        return true;
    }

    const std::filesystem::path delDir                    = state.tempRoot / L"perm-delete";
    const std::filesystem::path delFile                   = delDir / L"perm.bin";
    const std::filesystem::path cancelFile                = delDir / L"cancelled-perm.bin";
    const std::filesystem::path copyMovePromptFile        = delDir / L"copy-move-confirmation.bin";
    const std::filesystem::path copyMovePromptDestination = delDir / L"copy-move-confirmation-destination";
    const std::filesystem::path replacementSource         = delDir / L"replacement-race.bin";
    const std::filesystem::path replacementOriginal       = delDir / L"replacement-race-original.bin";
    const std::filesystem::path recursiveDeleteRoot       = delDir / L"recursive-delete";
    const std::filesystem::path recursiveDeleteChild      = recursiveDeleteRoot / L"nested" / L"child.bin";

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

        if (! WriteTestFile(copyMovePromptFile, 1024) || ! RecreateEmptyDirectory(copyMovePromptDestination))
        {
            Fail(L"Failed to prepare Copy/Move confirmation-presentation fixtures.");
            return true;
        }

        const auto verifyCopyMovePromptPresentation = [&](FileSystemOperation operation,
                                                          HostPromptPresentation expectedPresentation,
                                                          UINT expectedTitleId,
                                                          HostFileOperationVerificationAvailability expectedVerificationAvailability) noexcept -> bool
        {
            HostResetTestPromptRequestCount();
            const std::optional<std::uint64_t> task = StartFileOperationAndGetId(state.fileOps,
                                                                                 operation,
                                                                                 FolderWindow::Pane::Left,
                                                                                 FolderWindow::Pane::Right,
                                                                                 state.fsLocal,
                                                                                 {copyMovePromptFile},
                                                                                 copyMovePromptDestination,
                                                                                 FILESYSTEM_FLAG_NONE,
                                                                                 false,
                                                                                 0,
                                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                                 true);
            if (task.has_value())
            {
                Fail(L"Copy/Move confirmation-presentation probe started even though the prompt was cancelled.");
                return false;
            }

            HostPromptDebugSnapshot prompt{};
            if (HostGetTestPromptRequestCount() != 1u || ! HostGetTestLastPromptDebugSnapshot(prompt))
            {
                Fail(std::format(L"Copy/Move confirmation-presentation probe expected one captured prompt, got {}.", HostGetTestPromptRequestCount()));
                return false;
            }

            const std::wstring expectedTitle = LoadStringResource(nullptr, expectedTitleId);
            if (prompt.presentation != expectedPresentation || prompt.title != expectedTitle || prompt.buttons != HOST_PROMPT_BUTTONS_OK_CANCEL ||
                ! prompt.hasFileOperationOptions || prompt.fileOperationOptions.sizeBytes != sizeof(HostFileOperationPromptOptions) ||
                prompt.fileOperationOptions.linkPolicy != HOST_FILE_OPERATION_LINK_PRESERVE ||
                prompt.fileOperationOptions.executionMode != HOST_FILE_OPERATION_EXECUTION_PARALLEL ||
                prompt.fileOperationOptions.verificationAvailability != expectedVerificationAvailability)
            {
                Fail(std::format(L"Copy/Move confirmation snapshot mismatch: op={} presentation={} expected={} title='{}' expectedTitle='{}' "
                                 L"options={} link={} mode={} verification={} expectedVerification={}.",
                                 static_cast<unsigned int>(operation),
                                 static_cast<unsigned int>(prompt.presentation),
                                 static_cast<unsigned int>(expectedPresentation),
                                 prompt.title,
                                 expectedTitle,
                                 prompt.hasFileOperationOptions ? 1 : 0,
                                 static_cast<unsigned int>(prompt.fileOperationOptions.linkPolicy),
                                 static_cast<unsigned int>(prompt.fileOperationOptions.executionMode),
                                 static_cast<unsigned int>(prompt.fileOperationOptions.verificationAvailability),
                                 static_cast<unsigned int>(expectedVerificationAvailability)));
                return false;
            }

            return true;
        };

        {
            HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
            const auto clearPromptOverride = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });
            if (! verifyCopyMovePromptPresentation(
                    FILESYSTEM_COPY, HOST_PROMPT_PRESENTATION_COPY, IDS_FILEOP_OPERATION_COPY, HOST_FILE_OPERATION_VERIFICATION_SUPPORTED) ||
                ! verifyCopyMovePromptPresentation(
                    FILESYSTEM_MOVE, HOST_PROMPT_PRESENTATION_MOVE, IDS_FILEOP_OPERATION_MOVE, HOST_FILE_OPERATION_VERIFICATION_NOT_APPLICABLE))
            {
                return true;
            }

            SetFileOpsVerificationCapabilityBudgetExceededForSelfTest(true);
            const auto resetCapabilityBudgetHook = wil::scope_exit([]() noexcept { SetFileOpsVerificationCapabilityBudgetExceededForSelfTest(false); });
            if (! verifyCopyMovePromptPresentation(
                    FILESYSTEM_COPY, HOST_PROMPT_PRESENTATION_COPY, IDS_FILEOP_OPERATION_COPY, HOST_FILE_OPERATION_VERIFICATION_CHECK_DURING_OPERATION))
            {
                return true;
            }
        }

        HostFileOperationPromptOptions acceptedOptions{};
        acceptedOptions.sizeBytes       = sizeof(acceptedOptions);
        acceptedOptions.linkPolicy      = HOST_FILE_OPERATION_LINK_SKIP;
        acceptedOptions.verifyAfterCopy = 1u;
        // A capability response that exceeded the confirmation budget keeps the user's Verify
        // intent intact; execution, not the prompt, resolves exact proof availability.
        acceptedOptions.verificationAvailability     = HOST_FILE_OPERATION_VERIFICATION_CHECK_DURING_OPERATION;
        acceptedOptions.executionMode                = HOST_FILE_OPERATION_EXECUTION_PARALLEL;
        acceptedOptions.bandwidthLimitBytesPerSecond = 50ull << 20u;
        std::optional<std::uint64_t> capturedOptionsTask;
        {
            HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
            HostSetTestPromptFileOperationOptionsOverride(acceptedOptions);
            const auto clearPromptOverrides = wil::scope_exit([]() noexcept
            {
                HostClearTestPromptFileOperationOptionsOverride();
                HostClearTestPromptResultOverride();
            });
            capturedOptionsTask             = StartFileOperationAndGetId(state.fileOps,
                                                                         FILESYSTEM_COPY,
                                                                         FolderWindow::Pane::Left,
                                                                         FolderWindow::Pane::Right,
                                                                         state.fsLocal,
                                                                         {copyMovePromptFile},
                                                                         copyMovePromptDestination,
                                                                         FILESYSTEM_FLAG_NONE,
                                                                         true,
                                                                         0,
                                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                         true);
        }
        auto* capturedTask = capturedOptionsTask.has_value() ? state.fileOps->FindTask(capturedOptionsTask.value()) : nullptr;
        const std::shared_ptr<const FileOperations::FileOperationPlanGroup> capturedPlans = capturedTask != nullptr ? capturedTask->LoadPlans() : nullptr;
        if (! capturedPlans || capturedPlans->empty())
        {
            Fail(L"Accepted Copy confirmation did not publish an immutable plan group.");
            return true;
        }
        for (const FileOperations::FileOperationPlan& plan : *capturedPlans)
        {
            const FileOperations::OperationOptions captured = std::visit([](const auto& typedPlan) noexcept { return typedPlan.options; }, plan);
            if (captured.linkPolicy != FileOperations::LinkPolicy::Skip || ! captured.verifyAfterCopy ||
                captured.executionMode != FileOperations::ExecutionMode::Parallel ||
                captured.bandwidthLimitBytesPerSecond != std::optional<uint64_t>(50ull << 20u))
            {
                Fail(L"Accepted Copy confirmation values were not copied consistently into every immutable child plan.");
                return true;
            }
        }
        if (capturedTask->_waitForOthers.load(std::memory_order_acquire) ||
            capturedTask->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire) != (50ull << 20u))
        {
            Fail(L"Accepted Copy confirmation start mode or bandwidth was not applied to task scheduling.");
            return true;
        }
        capturedTask->RequestCancel();

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        // C1: the permanent-delete confirmation is a card consent after Preparing has pinned the
        // roots. No modal prompt runs; the host prompt override captured at admission answers the
        // card, and a cancelled confirmation ends the published task as canceled with the source
        // intact.
        HostResetTestPromptRequestCount();
        {
            HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
            const auto clearPromptOverride = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });
            state.taskB                    = StartFileOperationAndGetId(state.fileOps,
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
        }
        if (! state.taskB.has_value())
        {
            Fail(L"Permanent delete without Recycle Bin must publish its task and confirm on the card (C1).");
            return true;
        }
        state.stepState = 10;
        return false;
    }

    if (state.stepState == 10)
    {
        const auto cancelled = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (cancelled == state.completedTasks.end())
        {
            return false;
        }
        std::error_code cancelExistsEc;
        if (cancelled->second.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || ! std::filesystem::exists(cancelFile, cancelExistsEc) || cancelExistsEc ||
            HostGetTestPromptRequestCount() != 0u)
        {
            Fail(std::format(L"A cancelled permanent-delete confirmation must end the task as canceled with the source intact and no modal prompt "
                             L"(hr=0x{:08X}, prompts={}).",
                             static_cast<unsigned long>(cancelled->second.hr),
                             HostGetTestPromptRequestCount()));
            return true;
        }
        state.taskB.reset();
        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

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

        if (HostGetTestPromptRequestCount() != 0u)
        {
            Fail(std::format(L"A confirmed permanent delete must not run a modal prompt; the confirmation is on the card (got {}).",
                             HostGetTestPromptRequestCount()));
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

    if (state.stepState == 2)
    {
        const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Permanent delete task failed: 0x{0:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        std::error_code ec;
        if (std::filesystem::exists(delFile, ec))
        {
            Fail(L"Permanent delete task did not remove the source file.");
            return true;
        }

        if (! WriteTestFile(replacementSource, 2048u))
        {
            Fail(L"Failed to create permanent-delete replacement-race source.");
            return true;
        }
        // C1: Preparing pins the root, the card confirms, and the post-confirmation re-check
        // refuses a pathname that no longer names the pinned object. The swap lands in that
        // window.
        SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(true);
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {replacementSource},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        HostClearTestPromptResultOverride();
        if (! state.taskB.has_value())
        {
            ReleaseFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest();
            SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(false);
            Fail(L"Failed to start permanent-delete replacement-race task.");
            return true;
        }
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (! HasFileOpsPermanentDeleteBeforeRecheckPauseEnteredForSelfTest())
        {
            return false;
        }
        // C1: the confirmed root is pinned with delete sharing denied, so its pathname cannot be
        // replaced between the card's answer and the mutation; the swap must be refused.
        const BOOL swapped    = MoveFileExW(replacementSource.c_str(), replacementOriginal.c_str(), MOVEFILE_WRITE_THROUGH);
        const DWORD swapError = GetLastError();
        ReleaseFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest();
        SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(false);
        if (swapped != FALSE || swapError != ERROR_SHARING_VIOLATION)
        {
            Fail(std::format(L"The pinned permanent-delete root must refuse a pathname swap while its confirmation is re-checked (swapped={}, error={}).",
                             swapped != FALSE,
                             swapError));
            return true;
        }
        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        const auto completed = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        std::error_code replacementEc;
        if (FAILED(completed->second.hr) || std::filesystem::exists(replacementSource, replacementEc) ||
            std::filesystem::exists(replacementOriginal, replacementEc))
        {
            Fail(std::format(L"Permanent delete of the pinned root must complete and leave no replacement behind (hr=0x{0:08X}).",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        if (! WriteTestFile(recursiveDeleteChild, 3072u))
        {
            Fail(L"Failed to create recursive permanent-delete fixture.");
            return true;
        }
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
        state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {recursiveDeleteRoot},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        HostClearTestPromptResultOverride();
        if (! state.taskC.has_value())
        {
            Fail(L"Failed to start recursive exact-authority permanent-delete task.");
            return true;
        }
        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        const auto completed = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        std::error_code recursiveEc;
        if (FAILED(completed->second.hr) || std::filesystem::exists(recursiveDeleteRoot, recursiveEc) || recursiveEc)
        {
            Fail(std::format(L"Recursive permanent delete did not remove the exact bound root (hr=0x{0:08X}).",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        state.stepState = 6;
    }

    std::error_code ec;

    // Validate the legacy directory-size API file-root contract on local filesystem (S_OK + fileCount=1).
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

    // A non-root enumeration failure is partial evidence, not a reason to discard later sibling totals.
    constexpr std::wstring_view kDirectorySizeFailChildPathEnv  = L"REDSALAMANDER_DIRECTORY_SIZE_FAIL_CHILD_PATH";
    constexpr std::wstring_view kDirectorySizeFailChildFiredEnv = L"REDSALAMANDER_DIRECTORY_SIZE_FAIL_CHILD_FIRED";
    const std::filesystem::path partialSizeRoot                 = state.tempRoot / L"directory-size-partial";
    const std::filesystem::path deniedChild                     = partialSizeRoot / L"denied-child";
    const std::filesystem::path readableSibling                 = partialSizeRoot / L"readable-sibling";
    auto clearDirectorySizeHook                                 = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(::SetEnvironmentVariableW(kDirectorySizeFailChildPathEnv.data(), nullptr));
        static_cast<void>(::SetEnvironmentVariableW(kDirectorySizeFailChildFiredEnv.data(), nullptr));
    });
    std::filesystem::remove_all(partialSizeRoot, ec);
    ec.clear();
    std::filesystem::create_directories(deniedChild, ec);
    std::filesystem::create_directories(readableSibling, ec);
    if (ec || ! WriteTestFile(deniedChild / L"ignored.bin", 111u) || ! WriteTestFile(readableSibling / L"counted.bin", 222u) ||
        ! ::SetEnvironmentVariableW(kDirectorySizeFailChildPathEnv.data(), deniedChild.c_str()))
    {
        Fail(L"Failed to prepare the partial directory-size child-enumeration test.");
        return true;
    }

    FileSystemDirectorySizeResult partialSizeResult{};
    partialSizeResult.sizeBytes = sizeof(FileSystemDirectorySizeResult);
    const HRESULT partialSizeHr = localDirOps->GetDirectorySize(partialSizeRoot.c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, &partialSizeResult);
    if (partialSizeHr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) || partialSizeResult.status != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) ||
        partialSizeResult.totalBytes != 222u || partialSizeResult.fileCount != 1u || GetEnvVarTrimmed(kDirectorySizeFailChildFiredEnv) != L"1")
    {
        Fail(std::format(L"Directory-size partial traversal mismatch: hr=0x{:08X} status=0x{:08X} bytes={} files={} fired={}.",
                         static_cast<unsigned long>(partialSizeHr),
                         static_cast<unsigned long>(partialSizeResult.status),
                         partialSizeResult.totalBytes,
                         partialSizeResult.fileCount,
                         GetEnvVarTrimmed(kDirectorySizeFailChildFiredEnv)));
        return true;
    }

    // Validate the legacy directory-size API file-root contract on dummy filesystem (S_OK + fileCount=1).
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

    NextStep(state, SelfTestState::Step::Phase10_DeferredConsentAndRecycleEscalation);
    return false;
}
case SelfTestState::Step::Phase10_DeferredConsentAndRecycleEscalation:
{
    constexpr wchar_t kRecycleFailurePathEnvVar[] = L"REDSALAMANDER_FILEOPS_RECYCLE_FAIL_PATH";
    const ULONGLONG nowTick                       = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        static_cast<void>(SetEnvironmentVariableW(kRecycleFailurePathEnvVar, nullptr));
        ReleaseFileOpsRecycleEscalationBeforeBindPauseForSelfTest();
        SetFileOpsRecycleEscalationBeforeBindPauseForSelfTest(false);
        state.lockedFileHandle.reset();
        Fail(L"Phase10_DeferredConsentAndRecycleEscalation timed out.");
        return true;
    }

    using Task                                   = FolderWindow::FileOperationState::Task;
    using ConflictAction                         = Task::ConflictAction;
    using ConflictBucket                         = Task::ConflictBucket;
    const std::filesystem::path escalationSource = state.tempRoot / L"recycle-escalation.bin";
    const std::filesystem::path swapSource       = state.tempRoot / L"recycle-swap.bin";
    const std::filesystem::path swapOriginal     = state.tempRoot / L"recycle-swap-original.bin";

    if (state.stepState == 0)
    {
        if (! DebugFileOpsUnboundCollisionWithholdsDestructiveActionsForSelfTest())
        {
            Fail(L"A provider-returned collision exposed a destructive action without exact destination authority.");
            return true;
        }
        if (! DebugFileOpsConflictActionPolicyCoverageForSelfTest())
        {
            Fail(L"The engine-owned conflict action policy did not preserve ordered actions, placement, default/Escape, or scope eligibility.");
            return true;
        }

        struct ConsentScenario final
        {
            FileOperations::DeferredConsentRisk risk;
            ConflictBucket bucket;
            ConflictAction decision;
        };
        constexpr std::array scenarios{
            ConsentScenario{FileOperations::DeferredConsentRisk::InsufficientSpace, ConflictBucket::InsufficientSpace, ConflictAction::Proceed},
            ConsentScenario{FileOperations::DeferredConsentRisk::SpaceUnknown, ConflictBucket::SpaceUnknown, ConflictAction::Proceed},
            ConsentScenario{FileOperations::DeferredConsentRisk::EfsPlaintext, ConflictBucket::EfsPlaintext, ConflictAction::RetainSource},
            ConsentScenario{FileOperations::DeferredConsentRisk::SparseInflation, ConflictBucket::SparseInflation, ConflictAction::Proceed},
            ConsentScenario{FileOperations::DeferredConsentRisk::PlaceholderHydration, ConflictBucket::PlaceholderHydration, ConflictAction::Proceed},
            ConsentScenario{FileOperations::DeferredConsentRisk::MetadataLoss, ConflictBucket::MetadataLoss, ConflictAction::RetainSource},
            ConsentScenario{FileOperations::DeferredConsentRisk::RecycleEscalation, ConflictBucket::RecycleFailed, ConflictAction::PermanentDelete},
        };

        Task consentTask(*state.fileOps);
        consentTask._taskId = 0x4'200u;
        for (const ConsentScenario& scenario : scenarios)
        {
            HRESULT workerHr              = E_PENDING;
            ConflictAction selectedAction = ConflictAction::None;
            const bool applyEligible      = scenario.risk != FileOperations::DeferredConsentRisk::RecycleEscalation;
            std::jthread worker([&](std::stop_token) noexcept
            { workerHr = consentTask.DebugRequestDeferredConsentForSelfTest(scenario.risk, applyEligible, true, 7u, true, 12'345u, &selectedAction); });

            std::optional<Task::ConflictPromptState> prompt;
            const ULONGLONG deadline = GetTickCount64() + 4'000ull;
            while (! prompt.has_value() && GetTickCount64() < deadline)
            {
                prompt = TryGetConflictPromptCopy(&consentTask);
                if (! prompt.has_value())
                {
                    Sleep(10);
                }
            }
            if (! prompt.has_value() || ! prompt->deferredConsent || prompt->bucket != scenario.bucket || prompt->factItemCountKnown == false ||
                prompt->factItemCount != 7u || prompt->factBytesKnown == false || prompt->factBytes != 12'345u || prompt->applyToAllEligible != applyEligible ||
                prompt->actionCount == 0u ||
                (scenario.risk == FileOperations::DeferredConsentRisk::RecycleEscalation
                     ? prompt->actions[0] != ConflictAction::Cancel
                     : prompt->actions[prompt->actionCount - 1u] != ConflictAction::Cancel) ||
                ! PromptHasAction(prompt.value(), scenario.decision))
            {
                consentTask.RequestCancel();
                worker.join();
                Fail(L"Deferred-consent risk matrix did not expose the typed facts, action, scope, and safe Cancel default.");
                return true;
            }
            if (scenario.risk == FileOperations::DeferredConsentRisk::RecycleEscalation &&
                (PromptHasAction(prompt.value(), ConflictAction::SkipAll) || prompt->applyToAllEligible))
            {
                consentTask.RequestCancel();
                worker.join();
                Fail(L"Recycle escalation exposed an Apply-to-all action or scope.");
                return true;
            }

            consentTask.SubmitConflictDecision(scenario.decision, true);
            worker.join();
            if (FAILED(workerHr) || selectedAction != scenario.decision)
            {
                Fail(std::format(L"Deferred-consent decision failed: hr=0x{:08X} action={}.",
                                 static_cast<unsigned long>(workerHr),
                                 static_cast<unsigned int>(selectedAction)));
                return true;
            }
        }

        size_t receiptCount = 0u;
        {
            std::scoped_lock lock(consentTask._conflictArbiter.mutex);
            receiptCount = static_cast<size_t>(std::ranges::count_if(consentTask._conflictArbiter.consentReceipts,
                                                                     [](const std::optional<FileOperations::DeferredConsentReceipt>& receipt) noexcept
            { return receipt.has_value(); }));
        }
        if (receiptCount != scenarios.size() || consentTask._perf.consentPromptCount != scenarios.size() || consentTask._perf.conflictPromptCount != 0u)
        {
            Fail(std::format(L"Deferred-consent receipts/metrics mismatch: receipts={} consentPrompts={} conflictPrompts={}.",
                             receiptCount,
                             consentTask._perf.consentPromptCount,
                             consentTask._perf.conflictPromptCount));
            return true;
        }
        AppendLog(L"Phase10 deferred-consent matrix complete");

        Task invalidActionTask(*state.fileOps);
        HRESULT invalidActionHr        = E_PENDING;
        ConflictAction invalidSelected = ConflictAction::None;
        std::jthread invalidActionWorker([&](std::stop_token) noexcept
        {
            invalidActionHr = invalidActionTask.DebugRequestDeferredConsentForSelfTest(
                FileOperations::DeferredConsentRisk::InsufficientSpace, true, false, 0u, false, 0u, &invalidSelected);
        });
        const ULONGLONG invalidDeadline = GetTickCount64() + 4'000ull;
        while (! TryGetConflictPromptCopy(&invalidActionTask).has_value() && GetTickCount64() < invalidDeadline)
        {
            Sleep(10);
        }
        invalidActionTask.SubmitConflictDecision(ConflictAction::Overwrite, true);
        invalidActionWorker.join();
        if (invalidActionHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || invalidSelected != ConflictAction::Cancel)
        {
            Fail(L"An off-layout deferred-consent action did not fail closed to Cancel.");
            return true;
        }
        AppendLog(L"Phase10 deferred-consent off-layout action rejected");

        if (! WriteTestFile(escalationSource, 1024u))
        {
            Fail(L"Failed to create the exact Recycle-escalation fixture.");
            return true;
        }
        if (SetEnvironmentVariableW(kRecycleFailurePathEnvVar, escalationSource.c_str()) == FALSE)
        {
            Fail(L"Failed to arm the deterministic Recycle-failure fixture.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {escalationSource},
                                                 {},
                                                 FILESYSTEM_FLAG_USE_RECYCLE_BIN,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        if (! state.taskA.has_value())
        {
            static_cast<void>(SetEnvironmentVariableW(kRecycleFailurePathEnvVar, nullptr));
            Fail(L"Failed to start the exact Recycle-escalation task.");
            return true;
        }
        AppendLog(std::format(L"Phase10 Recycle escalation task started id={}", state.taskA.value()));
        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        auto* task                                            = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const std::optional<Task::ConflictPromptState> prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }
        if (! prompt->deferredConsent || prompt->bucket != ConflictBucket::RecycleFailed || prompt->applyToAllEligible ||
            ! PromptHasAction(prompt.value(), ConflictAction::PermanentDelete) || PromptHasAction(prompt.value(), ConflictAction::SkipAll) ||
            ! prompt->sourceIdentity.has_value())
        {
            Fail(L"Exact Recycle escalation did not expose the one-item identity-bound consent prompt.");
            return true;
        }
        task->SubmitConflictDecision(ConflictAction::PermanentDelete, true);
        AppendLog(L"Phase10 Recycle escalation permanent-delete consent submitted");
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
        std::error_code ec;
        if (FAILED(completed->second.hr) || std::filesystem::exists(escalationSource, ec))
        {
            Fail(std::format(L"Exact Recycle escalation failed or retained the source: hr=0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }
        AppendLog(L"Phase10 exact Recycle escalation completed");

        if (! WriteTestFile(swapSource, 1024u))
        {
            Fail(L"Failed to create the Recycle identity-swap fixture.");
            return true;
        }
        std::filesystem::remove(swapOriginal, ec);
        if (SetEnvironmentVariableW(kRecycleFailurePathEnvVar, swapSource.c_str()) == FALSE)
        {
            Fail(L"Failed to arm the Recycle identity-swap failure fixture.");
            return true;
        }
        SetFileOpsRecycleEscalationBeforeBindPauseForSelfTest(true);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {swapSource},
                                                 {},
                                                 FILESYSTEM_FLAG_USE_RECYCLE_BIN,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        if (! state.taskB.has_value())
        {
            static_cast<void>(SetEnvironmentVariableW(kRecycleFailurePathEnvVar, nullptr));
            ReleaseFileOpsRecycleEscalationBeforeBindPauseForSelfTest();
            SetFileOpsRecycleEscalationBeforeBindPauseForSelfTest(false);
            Fail(L"Failed to start the Recycle identity-swap task.");
            return true;
        }
        AppendLog(std::format(L"Phase10 Recycle identity-swap task started id={}", state.taskB.value()));
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (! HasFileOpsRecycleEscalationBeforeBindPauseEnteredForSelfTest())
        {
            return false;
        }
        if (MoveFileExW(swapSource.c_str(), swapOriginal.c_str(), MOVEFILE_REPLACE_EXISTING) == FALSE || ! WriteTestFile(swapSource, 2048u))
        {
            ReleaseFileOpsRecycleEscalationBeforeBindPauseForSelfTest();
            SetFileOpsRecycleEscalationBeforeBindPauseForSelfTest(false);
            Fail(std::format(L"Failed to replace the Recycle source at the deterministic identity boundary (err={}).", GetLastError()));
            return true;
        }
        ReleaseFileOpsRecycleEscalationBeforeBindPauseForSelfTest();
        SetFileOpsRecycleEscalationBeforeBindPauseForSelfTest(false);
        AppendLog(L"Phase10 Recycle identity replacement injected before escalation bind");
        state.stepState = 4;
        return false;
    }

    auto* swapTask = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
    if (TryGetConflictPromptCopy(swapTask).has_value())
    {
        swapTask->RequestCancel();
        Fail(L"Recycle identity replacement reached a permanent-delete prompt instead of failing before consent.");
        return true;
    }
    const auto completed = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    std::error_code ec;
    if (SUCCEEDED(completed->second.hr) || ! std::filesystem::exists(swapSource, ec) || ! std::filesystem::exists(swapOriginal, ec))
    {
        Fail(std::format(L"Recycle identity-swap fail-closed result mismatch: hr=0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
        return true;
    }
    AppendLog(L"Phase10 Recycle identity-swap failed closed");

    NextStep(state, SelfTestState::Step::Phase10_MetadataPreservationAndSourceRetention);
    return false;
}
case SelfTestState::Step::Phase10_MetadataPreservationAndSourceRetention:
{
    constexpr wchar_t kForcePresentEnv[] = L"REDSALAMANDER_FILEOPS_METADATA_FORCE_PRESENT_MASK";
    constexpr wchar_t kFailMaskEnv[]     = L"REDSALAMANDER_FILEOPS_METADATA_FAIL_MASK";
    const ULONGLONG nowTick              = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        static_cast<void>(SetEnvironmentVariableW(kForcePresentEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kFailMaskEnv, nullptr));
        Fail(L"Phase10_MetadataPreservationAndSourceRetention timed out.");
        return true;
    }

    using Task             = FolderWindow::FileOperationState::Task;
    using ConflictAction   = Task::ConflictAction;
    using ConflictBucket   = Task::ConflictBucket;
    const auto streamPath  = [](const std::filesystem::path& path, std::wstring_view name) { return std::format(L"{}:{}", path.wstring(), name); };
    const auto writeStream = [&](const std::filesystem::path& path, std::wstring_view name, std::string_view payload)
    {
        const std::wstring fullPath = streamPath(path, name);
        wil::unique_handle stream(CreateFileW(
            fullPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! stream)
        {
            return false;
        }
        DWORD written = 0u;
        return WriteFile(stream.get(), payload.data(), static_cast<DWORD>(payload.size()), &written, nullptr) != FALSE &&
               written == static_cast<DWORD>(payload.size());
    };
    const auto readStream = [&](const std::filesystem::path& path, std::wstring_view name) -> std::optional<std::string>
    {
        const std::wstring fullPath = streamPath(path, name);
        wil::unique_handle stream(CreateFileW(
            fullPath.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! stream)
        {
            return std::nullopt;
        }
        LARGE_INTEGER size{};
        if (GetFileSizeEx(stream.get(), &size) == FALSE || size.QuadPart < 0 || size.QuadPart > 64ll * 1024ll)
        {
            return std::nullopt;
        }
        std::string payload(static_cast<size_t>(size.QuadPart), '\0');
        DWORD read = 0u;
        if (! payload.empty() &&
            (ReadFile(stream.get(), payload.data(), static_cast<DWORD>(payload.size()), &read, nullptr) == FALSE || read != static_cast<DWORD>(payload.size())))
        {
            return std::nullopt;
        }
        return payload;
    };
    // Beeline: the Managed route exists only across endpoints. On one volume the loopback
    // administrative-share alias of the sandbox (local-win32-smb) is the second qualified endpoint;
    // every assertion still reads the published files through their drive-letter paths.
    const auto startManagedDirectoryMove = [&](const std::filesystem::path& sourceTree,
                                               const std::filesystem::path& destinationParent) -> std::optional<uint64_t>
    {
        return StartFileOperationAndGetId(state.fileOps,
                                          FILESYSTEM_MOVE,
                                          FolderWindow::Pane::Left,
                                          FolderWindow::Pane::Right,
                                          state.fsLocal,
                                          {sourceTree},
                                          LoopbackShareAlias(destinationParent),
                                          FILESYSTEM_FLAG_RECURSIVE,
                                          false,
                                          0u,
                                          FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                          false,
                                          state.fsLocal);
    };

    const std::filesystem::path preserveSourceTree          = state.tempRoot / L"p4-metadata-preserve-source" / L"tree";
    const std::filesystem::path preserveDestinationParent   = state.tempRoot / L"p4-metadata-preserve-destination";
    const std::filesystem::path preserveDestinationTree     = preserveDestinationParent / preserveSourceTree.filename();
    const std::filesystem::path preserveSourceFile          = preserveSourceTree / L"payload.bin";
    const std::filesystem::path preserveDestinationFile     = preserveDestinationTree / preserveSourceFile.filename();
    const std::filesystem::path cancelLossSourceTree        = state.tempRoot / L"p4-metadata-cancel-source" / L"tree";
    const std::filesystem::path cancelLossDestinationParent = state.tempRoot / L"p4-metadata-cancel-destination";
    const std::filesystem::path cancelLossSourceFile        = cancelLossSourceTree / L"cancel.bin";
    const std::filesystem::path cancelLossDestinationFile   = cancelLossDestinationParent / cancelLossSourceTree.filename() / cancelLossSourceFile.filename();
    constexpr std::string_view kMotw                        = "[ZoneTransfer]\r\nZoneId=3\r\n";
    constexpr std::string_view kAlternate                   = "red-salamander-metadata";

    if (state.stepState == 0)
    {
        if (! LoopbackShareReachable(state.tempRoot))
        {
            AppendLog(L"P4.3: the loopback administrative share is not reachable; the Managed-route metadata cases are environment-gated and skipped.");
            NextStep(state, SelfTestState::Step::Phase10_TypedResultsAndConsumers);
            return false;
        }
        static_cast<void>(SetEnvironmentVariableW(kForcePresentEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kFailMaskEnv, nullptr));
        state.phase10MetadataSparseSeeded = false;
        if (! RecreateEmptyDirectory(preserveSourceTree) || ! RecreateEmptyDirectory(preserveDestinationTree) ||
            ! WriteTestFile(preserveSourceFile, 512u * 1024u) || ! WriteTestFile(preserveDestinationTree / L"existing.bin", 64u) ||
            ! writeStream(preserveSourceFile, L"Zone.Identifier", kMotw) || ! writeStream(preserveSourceFile, L"note", kAlternate))
        {
            Fail(L"P4.3 metadata-preservation fixture could not be created on the local filesystem.");
            return true;
        }

        {
            // The write handle closes before the Move starts; the fixture must not hold its own
            // source open while the bridge inspects and relocates it.
            const wil::unique_handle sparseFile(CreateFileW(preserveSourceFile.c_str(),
                                                            GENERIC_WRITE,
                                                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                            nullptr,
                                                            OPEN_EXISTING,
                                                            FILE_ATTRIBUTE_NORMAL,
                                                            nullptr));
            if (sparseFile)
            {
                FILE_SET_SPARSE_BUFFER sparse{};
                sparse.SetSparse = TRUE;
                DWORD returned   = 0u;
                state.phase10MetadataSparseSeeded =
                    DeviceIoControl(sparseFile.get(), FSCTL_SET_SPARSE, &sparse, sizeof(sparse), nullptr, 0u, &returned, nullptr) != FALSE;
            }
        }

        wil::com_ptr<IFileSystemIO> localIo;
        FileSystemBasicInformation sourceBasic{};
        sourceBasic.sizeBytes = sizeof(sourceBasic);
        if (! state.fsLocal || FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo ||
            FAILED(localIo->GetFileBasicInformation(preserveSourceFile.c_str(), &sourceBasic)))
        {
            Fail(L"P4.3 metadata-preservation fixture could not read exact source basic information.");
            return true;
        }
        sourceBasic.creationTime -= 36ll * 60ll * 60ll * 10'000'000ll;
        sourceBasic.lastWriteTime -= 36ll * 60ll * 60ll * 10'000'000ll;
        sourceBasic.attributes |= FILE_ATTRIBUTE_HIDDEN;
        if (FAILED(localIo->SetFileBasicInformation(preserveSourceFile.c_str(), &sourceBasic)))
        {
            Fail(L"P4.3 metadata-preservation fixture could not seed source timestamps.");
            return true;
        }
        FileSystemBasicInformation seededBasic{};
        seededBasic.sizeBytes = sizeof(seededBasic);
        if (FAILED(localIo->GetFileBasicInformation(preserveSourceFile.c_str(), &seededBasic)))
        {
            Fail(L"P4.3 metadata-preservation fixture could not verify seeded source metadata.");
            return true;
        }
        state.phase10MetadataExpectedWriteTime = seededBasic.lastWriteTime;

        state.taskA = startManagedDirectoryMove(preserveSourceTree, preserveDestinationParent);
        if (! state.taskA.has_value())
        {
            Fail(L"P4.3 metadata-preservation Managed Move was not admitted.");
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
        std::error_code ec;
        const auto copiedMotw             = readStream(preserveDestinationFile, L"Zone.Identifier");
        const auto copiedAlternate        = readStream(preserveDestinationFile, L"note");
        const DWORD destinationAttributes = GetFileAttributesW(preserveDestinationFile.c_str());
        wil::com_ptr<IFileSystemIO> localIo;
        FileSystemBasicInformation destinationBasic{};
        destinationBasic.sizeBytes       = sizeof(destinationBasic);
        const HRESULT destinationBasicHr = state.fsLocal && SUCCEEDED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) && localIo
                                               ? localIo->GetFileBasicInformation(preserveDestinationFile.c_str(), &destinationBasic)
                                               : E_NOINTERFACE;
        if (FAILED(completed->second.hr) || std::filesystem::exists(preserveSourceTree, ec) || ec || ! std::filesystem::exists(preserveDestinationFile, ec) ||
            ec || ! copiedMotw.has_value() || copiedMotw.value() != kMotw || ! copiedAlternate.has_value() || copiedAlternate.value() != kAlternate ||
            FAILED(destinationBasicHr) || destinationBasic.lastWriteTime != state.phase10MetadataExpectedWriteTime ||
            (destinationBasic.attributes & FILE_ATTRIBUTE_HIDDEN) == 0u ||
            (state.phase10MetadataSparseSeeded &&
             (destinationAttributes == INVALID_FILE_ATTRIBUTES || (destinationAttributes & FILE_ATTRIBUTE_SPARSE_FILE) == 0u)))
        {
            Fail(std::format(L"P4.3 exact metadata preservation failed: hr=0x{:08X} srcExists={} dstAttrs=0x{:08X} "
                             L"basicHr=0x{:08X} hidden={} expectedWrite={} actualWrite={} motw={} ads={} sparseSeeded={}.",
                             static_cast<unsigned long>(completed->second.hr),
                             std::filesystem::exists(preserveSourceTree, ec) ? 1 : 0,
                             destinationAttributes,
                             static_cast<unsigned long>(destinationBasicHr),
                             (destinationBasic.attributes & FILE_ATTRIBUTE_HIDDEN) != 0u ? 1 : 0,
                             state.phase10MetadataExpectedWriteTime,
                             destinationBasic.lastWriteTime,
                             copiedMotw.has_value() ? 1 : 0,
                             copiedAlternate.has_value() ? 1 : 0,
                             state.phase10MetadataSparseSeeded ? 1 : 0));
            return true;
        }

        const std::filesystem::path riskSourceTree        = state.tempRoot / L"p4-metadata-risk-source" / L"tree";
        const std::filesystem::path riskDestinationParent = state.tempRoot / L"p4-metadata-risk-destination";
        if (! RecreateEmptyDirectory(riskSourceTree) || ! RecreateEmptyDirectory(riskDestinationParent / riskSourceTree.filename()) ||
            ! WriteTestFile(riskSourceTree / L"risk.bin", 64u * 1024u))
        {
            Fail(L"P4.3 deferred metadata-risk fixture could not be created.");
            return true;
        }
        const uint32_t forceMask     = FILESYSTEM_METADATA_PLACEHOLDER | FILESYSTEM_METADATA_EFS | FILESYSTEM_METADATA_SPARSE;
        const uint32_t failMask      = FILESYSTEM_METADATA_EFS | FILESYSTEM_METADATA_SPARSE;
        const std::wstring forceText = std::format(L"0x{:X}", forceMask);
        const std::wstring failText  = std::format(L"0x{:X}", failMask);
        if (SetEnvironmentVariableW(kForcePresentEnv, forceText.c_str()) == FALSE || SetEnvironmentVariableW(kFailMaskEnv, failText.c_str()) == FALSE)
        {
            Fail(L"P4.3 could not arm deterministic metadata-risk injection.");
            return true;
        }
        state.taskB = startManagedDirectoryMove(riskSourceTree, riskDestinationParent);
        if (! state.taskB.has_value())
        {
            Fail(L"P4.3 metadata-risk Managed Move was not admitted.");
            return true;
        }
        state.stepState = 2;
        return false;
    }

    if (state.stepState >= 2 && state.stepState <= 4)
    {
        Task* task        = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }
        constexpr std::array<ConflictBucket, 3> expectedBuckets{{
            ConflictBucket::PlaceholderHydration,
            ConflictBucket::EfsPlaintext,
            ConflictBucket::SparseInflation,
        }};
        constexpr std::array<ConflictAction, 3> decisions{{
            ConflictAction::Proceed,
            ConflictAction::RetainSource,
            ConflictAction::Proceed,
        }};
        const size_t index = static_cast<size_t>(state.stepState - 2u);
        // SubmitConflictDecision is asynchronous. The just-submitted prompt can remain visible for
        // one UI pump while the worker consumes it; wait for the next typed gate instead of treating
        // that transient copy as an ordering failure.
        if (index > 0u && prompt->bucket == expectedBuckets[index - 1u])
        {
            return false;
        }
        if (! prompt->deferredConsent || prompt->bucket != expectedBuckets[index] || ! PromptHasAction(prompt.value(), decisions[index]))
        {
            Fail(L"P4.3 transfer-side placeholder/EFS/sparse gate did not expose the expected typed decision.");
            return true;
        }
        task->SubmitConflictDecision(decisions[index], false);
        ++state.stepState;
        return false;
    }

    if (state.stepState == 5)
    {
        const auto completed = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        static_cast<void>(SetEnvironmentVariableW(kForcePresentEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kFailMaskEnv, nullptr));
        const std::filesystem::path riskSourceTree      = state.tempRoot / L"p4-metadata-risk-source" / L"tree";
        const std::filesystem::path riskDestinationFile = state.tempRoot / L"p4-metadata-risk-destination" / L"tree" / L"risk.bin";
        std::error_code ec;
        if (completed->second.hr != S_FALSE || ! std::filesystem::exists(riskSourceTree, ec) || ec || ! std::filesystem::exists(riskDestinationFile, ec) || ec)
        {
            Fail(std::format(L"P4.3 retained-source metadata-risk outcome mismatch: hr=0x{:08X}.", static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        const std::filesystem::path lossSourceTree        = state.tempRoot / L"p4-metadata-loss-source" / L"tree";
        const std::filesystem::path lossDestinationParent = state.tempRoot / L"p4-metadata-loss-destination";
        const std::filesystem::path lossSourceFile        = lossSourceTree / L"loss.bin";
        if (! RecreateEmptyDirectory(lossSourceTree) || ! RecreateEmptyDirectory(lossDestinationParent / lossSourceTree.filename()) ||
            ! WriteTestFile(lossSourceFile, 64u * 1024u) || ! writeStream(lossSourceFile, L"Zone.Identifier", kMotw) ||
            ! writeStream(lossSourceFile, L"note", kAlternate))
        {
            Fail(L"P4.3 security-significant metadata-loss fixture could not be created.");
            return true;
        }
        const std::wstring failText =
            std::format(L"0x{:X}", FILESYSTEM_METADATA_MOTW | FILESYSTEM_METADATA_ALTERNATE_STREAMS | FILESYSTEM_METADATA_EXTENDED_ATTRIBUTES);
        if (SetEnvironmentVariableW(kFailMaskEnv, failText.c_str()) == FALSE)
        {
            Fail(L"P4.3 could not arm deterministic security-significant metadata loss.");
            return true;
        }
        state.taskC = startManagedDirectoryMove(lossSourceTree, lossDestinationParent);
        if (! state.taskC.has_value())
        {
            Fail(L"P4.3 metadata-loss Managed Move was not admitted.");
            return true;
        }
        state.stepState = 6;
        return false;
    }

    if (state.stepState == 6)
    {
        Task* task        = state.taskC.has_value() ? state.fileOps->FindTask(state.taskC.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }
        if (! prompt->deferredConsent || prompt->bucket != ConflictBucket::MetadataLoss || ! PromptHasAction(prompt.value(), ConflictAction::RetainSource) ||
            ! prompt->sourceIdentity.has_value())
        {
            Fail(L"P4.3 security-significant metadata loss did not expose identity-bound Keep Source consent.");
            return true;
        }
        task->SubmitConflictDecision(ConflictAction::RetainSource, false);
        state.stepState = 7;
        return false;
    }

    if (state.stepState == 7)
    {
        const auto completed = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        static_cast<void>(SetEnvironmentVariableW(kForcePresentEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kFailMaskEnv, nullptr));
        const std::filesystem::path lossSourceTree      = state.tempRoot / L"p4-metadata-loss-source" / L"tree";
        const std::filesystem::path lossDestinationFile = state.tempRoot / L"p4-metadata-loss-destination" / L"tree" / L"loss.bin";
        std::error_code ec;
        if (completed->second.hr != S_FALSE || ! std::filesystem::exists(lossSourceTree, ec) || ec || ! std::filesystem::exists(lossDestinationFile, ec) ||
            ec || readStream(lossDestinationFile, L"Zone.Identifier").has_value() || readStream(lossDestinationFile, L"note").has_value())
        {
            Fail(std::format(L"P4.3 security-significant loss did not publish-and-retain exactly: hr=0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        const auto summary =
            std::ranges::find_if(summaries, [&](const auto& value) noexcept { return state.taskC.has_value() && value.taskId == state.taskC.value(); });
        const bool exactLossReported = summary != summaries.end() && std::ranges::any_of(summary->issueDiagnostics, [](const auto& issue) noexcept {
            return issue.category == L"bridge.metadata.motw.lost" || issue.category == L"bridge.metadata.alternateStreams.lost";
        });
        if (! exactLossReported)
        {
            Fail(L"P4.3 security-significant loss did not report the exact metadata class.");
            return true;
        }
        if (! RecreateEmptyDirectory(cancelLossSourceTree) || ! RecreateEmptyDirectory(cancelLossDestinationParent / cancelLossSourceTree.filename()) ||
            ! WriteTestFile(cancelLossSourceFile, 64u * 1024u) || ! writeStream(cancelLossSourceFile, L"Zone.Identifier", kMotw))
        {
            Fail(L"P4.3 metadata-loss Cancel fixture could not be created.");
            return true;
        }
        const std::wstring cancelFailText = std::format(L"0x{:X}", static_cast<uint32_t>(FILESYSTEM_METADATA_MOTW));
        if (SetEnvironmentVariableW(kFailMaskEnv, cancelFailText.c_str()) == FALSE)
        {
            Fail(L"P4.3 could not arm deterministic metadata loss for Cancel.");
            return true;
        }
        state.taskA = startManagedDirectoryMove(cancelLossSourceTree, cancelLossDestinationParent);
        if (! state.taskA.has_value())
        {
            Fail(L"P4.3 metadata-loss Cancel Managed Move was not admitted.");
            return true;
        }
        state.stepState = 8;
        return false;
    }

    if (state.stepState == 8)
    {
        Task* task        = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }
        if (! prompt->deferredConsent || prompt->bucket != ConflictBucket::MetadataLoss || ! PromptHasAction(prompt.value(), ConflictAction::Cancel))
        {
            Fail(L"P4.3 metadata loss did not expose Cancel after publication.");
            return true;
        }
        task->SubmitConflictDecision(ConflictAction::Cancel, false);
        state.stepState = 9;
        return false;
    }

    if (state.stepState == 9)
    {
        const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completed == state.completedTasks.end())
        {
            return false;
        }
        static_cast<void>(SetEnvironmentVariableW(kForcePresentEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kFailMaskEnv, nullptr));
        const auto* const result = completed->second.sourceItemResults.size() == 1u && completed->second.sourceItemResults.front().has_value()
                                       ? std::addressof(completed->second.sourceItemResults.front().value())
                                       : nullptr;
        std::error_code ec;
        if (result == nullptr || result->publication != FileOperations::PublicationState::NotPublished ||
            result->sourceDisposition != FileOperations::SourceDisposition::Retained || result->completion != FileOperations::ItemCompletion::Canceled ||
            ! std::filesystem::exists(cancelLossSourceFile, ec) || ec || std::filesystem::exists(cancelLossDestinationFile, ec) || ec)
        {
            Fail(L"P4.3 metadata-loss Cancel did not report NotPublished + Retained + Canceled before destination publication.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.MetadataMatrix",
                          L"ads-motw-sparse=preserved;placeholder-efs-sparse=consented;loss=source-retained;cancel=not-published-retained",
                          0u,
                          4u,
                          0u,
                          S_OK);
        NextStep(state, SelfTestState::Step::Phase10_TypedResultsAndConsumers);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase10_TypedResultsAndConsumers:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Phase10_TypedResultsAndConsumers timed out.");
        return true;
    }
    const std::filesystem::path moveSource            = state.tempRoot / L"p4-results-move-source" / L"tree";
    const std::filesystem::path moveDestinationParent = state.tempRoot / L"p4-results-move-destination";
    const std::filesystem::path copySource            = state.tempRoot / L"p4-results-copy-source.bin";
    const std::filesystem::path copyDestinationParent = state.tempRoot / L"p4-results-copy-destination";

    if (state.stepState == 0)
    {
        state.taskA.reset();
        state.taskC.reset();
        if (! RecreateEmptyDirectory(moveSource) || ! RecreateEmptyDirectory(moveDestinationParent / moveSource.filename()) ||
            ! WriteTestFile(moveSource / L"payload.bin", 32u * 1024u) ||
            ! WriteTestFile(moveDestinationParent / moveSource.filename() / L"existing.bin", 64u) || ! RecreateEmptyDirectory(copyDestinationParent) ||
            ! WriteTestFile(copySource, 16u * 1024u))
        {
            Fail(L"P4.4 typed-result fixtures could not be created.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {moveSource},
                                                 moveDestinationParent,
                                                 FILESYSTEM_FLAG_RECURSIVE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskA.has_value())
        {
            Fail(L"P4.4 Move fixture was not admitted.");
            return true;
        }
        state.stepState = 1;
        return false;
    }
    if (state.stepState == 1)
    {
        if (! state.taskA.has_value() || ! state.completedTasks.contains(state.taskA.value()))
        {
            return false;
        }
        state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {copySource},
                                                 copyDestinationParent,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskC.has_value())
        {
            Fail(L"P4.4 Copy fixture was not admitted.");
            return true;
        }
        state.stepState = 2;
        return false;
    }

    const auto completedA = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    const auto completedC = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
    if (completedA == state.completedTasks.end() || completedC == state.completedTasks.end())
    {
        return false;
    }
    const auto exactResult = [](const CompletedTaskInfo& info) noexcept -> const FileOperations::FileOperationItemResult*
    {
        return info.sourceItemResults.size() == 1u && info.sourceItemResults.front().has_value() ? std::addressof(info.sourceItemResults.front().value())
                                                                                                 : nullptr;
    };
    const FileOperations::FileOperationItemResult* const moved  = exactResult(completedA->second);
    const FileOperations::FileOperationItemResult* const copied = exactResult(completedC->second);
    if (moved == nullptr || moved->publication != FileOperations::PublicationState::Published ||
        moved->sourceDisposition != FileOperations::SourceDisposition::Removed || moved->completion != FileOperations::ItemCompletion::Completed ||
        moved->status != S_OK)
    {
        Fail(L"P4.4 successful same-endpoint Move did not emit Published + Removed + Completed.");
        return true;
    }
    if (copied == nullptr || copied->publication != FileOperations::PublicationState::Published ||
        copied->sourceDisposition != FileOperations::SourceDisposition::Retained || copied->completion != FileOperations::ItemCompletion::Completed ||
        copied->status != S_OK)
    {
        Fail(L"P4.4 Copy did not emit Published + Retained + Completed.");
        return true;
    }

    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
    state.fileOps->CollectCompletedTasks(summaries);
    const auto movedSummary =
        std::ranges::find_if(summaries, [&](const auto& summary) noexcept { return state.taskA.has_value() && summary.taskId == state.taskA.value(); });
    const auto copiedSummary =
        std::ranges::find_if(summaries, [&](const auto& summary) noexcept { return state.taskC.has_value() && summary.taskId == state.taskC.value(); });
    if (movedSummary == summaries.end() || movedSummary->publishedItemCount != 1u || movedSummary->removedSourceCount != 1u ||
        movedSummary->retainedSourceCount != 0u || movedSummary->resultSummary != LoadStringResource(nullptr, IDS_FILEOPS_RESULT_MOVED))
    {
        Fail(L"P4.4 completed-task summary did not map the exact removed source to Moved.");
        return true;
    }
    if (copiedSummary == summaries.end() || copiedSummary->publishedItemCount != 1u || copiedSummary->retainedSourceCount != 1u ||
        copiedSummary->removedSourceCount != 0u || copiedSummary->resultSummary != LoadStringResource(nullptr, IDS_FILEOPS_RESULT_COPIED))
    {
        Fail(L"P4.4 completed-task summary did not map the retained Copy source to Copied.");
        return true;
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.TypedResults", L"published+removed;published+retained;consumer-removal=exact", 0u, 2u, 0u, S_OK);
    NextStep(state, SelfTestState::Step::Phase10_ClipboardAdmissionAndRetainedActions);
    return false;
}
case SelfTestState::Step::Phase10_ClipboardAdmissionAndRetainedActions:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Phase10_ClipboardAdmissionAndRetainedActions timed out.");
        return true;
    }

    const std::filesystem::path sourceRoot            = state.tempRoot / L"p4-clipboard-source";
    const std::filesystem::path destinationRoot       = state.tempRoot / L"p4-clipboard-destination";
    const std::filesystem::path acceptedSource        = sourceRoot / L"accepted.bin";
    const std::filesystem::path failedClearSource     = sourceRoot / L"clear-failed.bin";
    const std::filesystem::path failedReadinessSource = sourceRoot / L"readiness-failed.bin";
    const std::filesystem::path duplicateSource       = sourceRoot / L"duplicate.bin";
    constexpr uint32_t kAcceptedSequence              = 0x41A50001u;
    constexpr uint32_t kFailedClearSequence           = 0x41A50002u;
    constexpr uint32_t kFailedReadinessSequence       = 0x41A50003u;

    if (state.stepState == 0)
    {
        state.taskA.reset();
        state.taskB.reset();
        state.taskC.reset();
        state.phase10ClipboardPublishedNonRunning.store(false, std::memory_order_release);
        state.phase10ClipboardReadinessBarrierCalled.store(false, std::memory_order_release);
        state.phase10ClipboardSlowGateRelease.store(false, std::memory_order_release);
        state.phase10ClipboardReturnedBeforeReadiness = false;
        if (! RecreateEmptyDirectory(sourceRoot) || ! RecreateEmptyDirectory(destinationRoot) || ! WriteTestFile(acceptedSource, 4096u) ||
            ! WriteTestFile(failedClearSource, 4096u) || ! WriteTestFile(failedReadinessSource, 4096u) || ! WriteTestFile(duplicateSource, 4096u))
        {
            Fail(L"P4.5 clipboard-admission fixtures could not be created.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {acceptedSource},
                                                 destinationRoot,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal,
                                                 {},
                                                 kAcceptedSequence,
                                                 [&]() noexcept
        {
            std::vector<FolderWindow::FileOperationState::Task*> tasks;
            state.fileOps->CollectTasks(tasks);
            state.phase10ClipboardPublishedNonRunning.store(
                std::ranges::any_of(tasks, [](const auto* task) noexcept { return task != nullptr && task->_clipboardMoveAdmission && ! task->HasStarted(); }),
                std::memory_order_release);
            return S_OK;
        });
        if (! state.taskA.has_value())
        {
            Fail(L"P4.5 did not publish an accepted clipboard Move task.");
            return true;
        }

        bool duplicateBarrierCalled                 = false;
        const std::optional<uint64_t> duplicateTask = StartFileOperationAndGetId(state.fileOps,
                                                                                 FILESYSTEM_MOVE,
                                                                                 FolderWindow::Pane::Left,
                                                                                 FolderWindow::Pane::Right,
                                                                                 state.fsLocal,
                                                                                 {duplicateSource},
                                                                                 destinationRoot,
                                                                                 FILESYSTEM_FLAG_NONE,
                                                                                 false,
                                                                                 0u,
                                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                                 false,
                                                                                 state.fsLocal,
                                                                                 {},
                                                                                 kAcceptedSequence,
                                                                                 [&]() noexcept
        {
            duplicateBarrierCalled = true;
            return S_OK;
        });
        if (duplicateTask.has_value() || duplicateBarrierCalled)
        {
            Fail(L"P4.5 admitted the same clipboard sequence twice.");
            return true;
        }

        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {failedClearSource},
                                                 destinationRoot,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal,
                                                 {},
                                                 kFailedClearSequence,
                                                 []() noexcept { return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED); });
        if (! state.taskB.has_value())
        {
            Fail(L"P4.5 discarded an accepted Move when clipboard clearing failed.");
            return true;
        }

        state.taskC                                           = StartFileOperationAndGetId(state.fileOps,
                                                                                           FILESYSTEM_MOVE,
                                                                                           FolderWindow::Pane::Left,
                                                                                           FolderWindow::Pane::Right,
                                                                                           state.fsLocal,
                                                                                           {failedReadinessSource},
                                                                                           destinationRoot,
                                                                                           FILESYSTEM_FLAG_NONE,
                                                                                           false,
                                                                                           0u,
                                                                                           FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                                           false,
                                                                                           state.fsLocal,
                                                                                           {},
                                                                                           kFailedReadinessSequence,
                                                                                           [&]() noexcept
        {
            state.phase10ClipboardReadinessBarrierCalled.store(true, std::memory_order_release);
            return S_OK;
        },
                                                 [&]() noexcept
        {
            state.phase10ClipboardSlowGateRelease.wait(false, std::memory_order_acquire);
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        });
        FolderWindow::FileOperationState::Task* readinessTask = state.taskC.has_value() ? state.fileOps->FindTask(state.taskC.value()) : nullptr;
        state.phase10ClipboardReturnedBeforeReadiness =
            readinessTask != nullptr && ! readinessTask->_selectedRootReadinessComplete.load(std::memory_order_acquire);
        state.phase10ClipboardSlowGateRelease.store(true, std::memory_order_release);
        state.phase10ClipboardSlowGateRelease.notify_all();
        if (! state.taskC.has_value() || ! state.phase10ClipboardReturnedBeforeReadiness)
        {
            Fail(L"R1a clipboard admission did not return to the UI before slow worker-owned readiness completed.");
            return true;
        }
        state.stepState = 1;
        return false;
    }

    if (! state.taskA.has_value() || ! state.taskB.has_value() || ! state.taskC.has_value() || ! state.completedTasks.contains(state.taskA.value()) ||
        ! state.completedTasks.contains(state.taskB.value()) || ! state.completedTasks.contains(state.taskC.value()))
    {
        return false;
    }

    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
    state.fileOps->CollectCompletedTasks(summaries);
    const auto accepted        = std::ranges::find_if(summaries, [&](const auto& summary) noexcept { return summary.taskId == state.taskA.value(); });
    const auto clearFailed     = std::ranges::find_if(summaries, [&](const auto& summary) noexcept { return summary.taskId == state.taskB.value(); });
    const auto readinessFailed = std::ranges::find_if(summaries, [&](const auto& summary) noexcept { return summary.taskId == state.taskC.value(); });

    constexpr size_t kR1aBuilderCount        = 4'096u;
    constexpr uint64_t kR1aGateDecisionCount = 1'024u;
    const auto r1aPerfStarted                = std::chrono::steady_clock::now();
    FolderWindow::FileOperationState::Task r1aTerminalTask(*state.fileOps);
    r1aTerminalTask._operation = FILESYSTEM_MOVE;
    r1aTerminalTask._sourcePaths.assign(kR1aBuilderCount, std::filesystem::path(L"r1a-selected-root"));
    r1aTerminalTask.InitializeSourceItemResultBuilders();
    static_cast<void>(r1aTerminalTask.FinalizeTypedItemResults(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)));
    const uint64_t r1aTerminalStoreCount = static_cast<uint64_t>(
        std::ranges::count_if(r1aTerminalTask._sourceItemResultBuilders, [](const auto& builder) noexcept { return builder.terminal.has_value(); }));
    const uint64_t r1aPendingResultCount =
        static_cast<uint64_t>(std::ranges::count_if(r1aTerminalTask._sourceItemResultBuilders, [](const auto& builder) noexcept {
        return builder.terminal.has_value() && builder.terminal->status == E_PENDING;
    }));

    FolderWindow::FileOperationState::Task r1aCanceledTask(*state.fileOps);
    r1aCanceledTask._operation   = FILESYSTEM_MOVE;
    r1aCanceledTask._sourcePaths = {std::filesystem::path(L"r1a-canceled-root-a"), std::filesystem::path(L"r1a-canceled-root-b")};
    r1aCanceledTask.InitializeSourceItemResultBuilders();
    static_cast<void>(r1aCanceledTask.FinalizeTypedItemResults(HRESULT_FROM_WIN32(ERROR_CANCELLED)));
    const bool canceledPreparingDense = std::ranges::all_of(r1aCanceledTask._sourceItemResultBuilders,
                                                            [](const auto& builder) noexcept
    {
        return builder.terminal.has_value() && builder.terminal->status == HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
               builder.terminal->publication == FileOperations::PublicationState::NotAttempted &&
               builder.terminal->sourceDisposition == FileOperations::SourceDisposition::Retained &&
               builder.terminal->completion == FileOperations::ItemCompletion::Canceled;
    });

    FolderWindow::FileOperationState::Task r1aDuplicateTask(*state.fileOps);
    r1aDuplicateTask._operation   = FILESYSTEM_COPY;
    r1aDuplicateTask._sourcePaths = {std::filesystem::path(L"r1a-duplicate-root")};
    r1aDuplicateTask.InitializeSourceItemResultBuilders();
    FileOperations::FileOperationItemResult firstResult{};
    firstResult.sourceIndex                                 = 0u;
    firstResult.publication                                 = FileOperations::PublicationState::Published;
    firstResult.sourceDisposition                           = FileOperations::SourceDisposition::Retained;
    firstResult.completion                                  = FileOperations::ItemCompletion::Completed;
    firstResult.status                                      = S_OK;
    const bool firstStoreAccepted                           = r1aDuplicateTask.StoreTypedItemResult(firstResult);
    FileOperations::FileOperationItemResult duplicateResult = firstResult;
    duplicateResult.publication                             = FileOperations::PublicationState::NotPublished;
    duplicateResult.completion                              = FileOperations::ItemCompletion::Failed;
    duplicateResult.status                                  = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    const bool duplicateStoreAccepted                       = r1aDuplicateTask.StoreTypedItemResult(std::move(duplicateResult));
    const bool duplicateStorePreservedFirst =
        firstStoreAccepted && ! duplicateStoreAccepted && r1aDuplicateTask._sourceItemResultBuilders.size() == 1u &&
        r1aDuplicateTask._sourceItemResultBuilders.front().terminal.has_value() &&
        r1aDuplicateTask._sourceItemResultBuilders.front().terminal->status == S_OK &&
        r1aDuplicateTask._sourceItemResultBuilders.front().terminal->publication == FileOperations::PublicationState::Published;
    const uint64_t duplicateStoreRejectCount = r1aDuplicateTask._duplicateTerminalStoreRejectCount.load(std::memory_order_acquire);

    uint64_t clipboardGateMutationPermitCount = 0u;
    for (uint64_t decision = 0u; decision < kR1aGateDecisionCount; ++decision)
    {
        const HRESULT consumeHr = (decision % 2u) == 0u ? S_OK : HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        clipboardGateMutationPermitCount += FolderWindow::FileOperationState::Task::ClipboardMutationGateAllows(S_OK, consumeHr, consumeHr == S_OK) ? 1u : 0u;
    }
    const uint64_t r1aTerminalFunnelUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - r1aPerfStarted).count());
    Debug::Perf::Emit(L"FileOps.SelfTest.R1a.TerminalFunnelUs", L"candidate", r1aTerminalFunnelUs, kR1aBuilderCount, r1aTerminalStoreCount, S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.R1a.BuilderCount", L"selected-roots", 0u, kR1aBuilderCount, kR1aBuilderCount, S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.R1a.TerminalStoreCount", L"dense", 0u, r1aTerminalStoreCount, kR1aBuilderCount, S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.R1a.DuplicateStoreRejectCount", L"first-truth-wins", 0u, duplicateStoreRejectCount, 1u, S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.R1a.ClipboardGateDecisionCount", L"no-provider", 0u, kR1aGateDecisionCount, clipboardGateMutationPermitCount, S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.R1a.ProviderCallCount", L"no-provider", 0u, 0u, 0u, S_OK);

    if (accepted == summaries.end() || ! accepted->clipboardMoveAdmission || ! accepted->clipboardMoveConsumed ||
        accepted->clipboardMoveConsumptionStatus != S_OK || ! state.phase10ClipboardPublishedNonRunning.load(std::memory_order_acquire))
    {
        Fail(L"P4.5 accepted clipboard Move did not retain the successful consumption receipt.");
        return true;
    }
    if (clearFailed == summaries.end() || ! clearFailed->clipboardMoveAdmission || clearFailed->clipboardMoveConsumed ||
        clearFailed->clipboardMoveConsumptionStatus != HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) || SUCCEEDED(clearFailed->resultHr) ||
        clearFailed->retainedSourceCount != 1u || clearFailed->unknownSourceCount != 0u || ! std::filesystem::exists(failedClearSource) ||
        std::filesystem::exists(destinationRoot / failedClearSource.filename()))
    {
        Fail(L"R1a clipboard-consumption failure did not terminate with dense known-no-mutation truth.");
        return true;
    }
    if (readinessFailed == summaries.end() || ! readinessFailed->clipboardMoveAdmission || readinessFailed->clipboardMoveConsumed ||
        readinessFailed->clipboardMoveConsumptionStatus != HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) || SUCCEEDED(readinessFailed->resultHr) ||
        readinessFailed->retainedSourceCount != 1u || readinessFailed->unknownSourceCount != 0u || ! std::filesystem::exists(failedReadinessSource) ||
        std::filesystem::exists(destinationRoot / failedReadinessSource.filename()) ||
        state.phase10ClipboardReadinessBarrierCalled.load(std::memory_order_acquire))
    {
        Fail(L"R1a selected-root readiness failure did not terminate before clipboard consumption with dense known-no-mutation truth.");
        return true;
    }
    if (r1aPendingResultCount != 0u || r1aTerminalStoreCount != kR1aBuilderCount || ! duplicateStorePreservedFirst ||
        clipboardGateMutationPermitCount != kR1aGateDecisionCount / 2u || ! canceledPreparingDense)
    {
        Fail(L"R1a terminal funnel still permits E_PENDING synthesis, terminal overwrite, or mutation after a failed clipboard gate.");
        return true;
    }

    constexpr FileSystemBindFlags retainedBindFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    FileOperations::ObjectBindingResult retainedBinding =
        FileOperations::BindObjectAuthority(state.fsLocal.get(), duplicateSource.native(), L"local-win32", retainedBindFlags);
    if (retainedBinding.state != FileOperations::ObjectBindingState::Bound)
    {
        Fail(L"P4.5 could not bind the retained-source identity fixture.");
        return true;
    }
    constexpr uint64_t kRetainedActionTaskId = 0xF4500001ull;
    FolderWindow::FileOperationState::CompletedTaskSummary retainedSummary{};
    retainedSummary.taskId                 = kRetainedActionTaskId;
    retainedSummary.operation              = FILESYSTEM_MOVE;
    retainedSummary.clipboardMoveAdmission = true;
    retainedSummary.retainedSourceCount    = 1u;
    retainedSummary.retainedSourcePaths.push_back(duplicateSource);
    retainedSummary.exactRetainedSourceItems.push_back({duplicateSource, retainedBinding.authority.identity});
    retainedBinding.authority.boundObject.reset();
    state.fileOps->DebugAppendCompletedTaskForSelfTest(std::move(retainedSummary));
    if (! state.fileOps->DebugValidateRetainedSourceActionForSelfTest(kRetainedActionTaskId, state.fsLocal.get()))
    {
        Fail(L"P4.5 rejected an unchanged exact retained-source action set.");
        return true;
    }

    std::error_code replaceError;
    if (! std::filesystem::remove(duplicateSource, replaceError) || replaceError || ! WriteTestFile(duplicateSource, 8192u) ||
        state.fileOps->DebugValidateRetainedSourceActionForSelfTest(kRetainedActionTaskId, state.fsLocal.get()))
    {
        Fail(L"P4.5 did not fail closed after a retained source was replaced at the same pathname.");
        return true;
    }
    state.fileOps->DismissCompletedTask(kRetainedActionTaskId);

    constexpr uint64_t kUnknownSourceTaskId = 0xF4500002ull;
    FolderWindow::FileOperationState::CompletedTaskSummary unknownSummary{};
    unknownSummary.taskId                 = kUnknownSourceTaskId;
    unknownSummary.operation              = FILESYSTEM_MOVE;
    unknownSummary.clipboardMoveAdmission = true;
    unknownSummary.clipboardMoveConsumed  = true;
    unknownSummary.unknownSourceCount     = 1u;
    unknownSummary.unknownSourcePaths.push_back(duplicateSource);
    state.fileOps->DebugAppendCompletedTaskForSelfTest(std::move(unknownSummary));
    if (state.fileOps->DebugValidateRetainedSourceActionForSelfTest(kUnknownSourceTaskId, state.fsLocal.get()))
    {
        Fail(L"P4.5 exposed exact retained-source actions for an Unknown source disposition.");
        return true;
    }
    state.fileOps->DismissCompletedTask(kUnknownSourceTaskId);

    Debug::Perf::Emit(L"FileOps.SelfTest.ClipboardAdmission",
                      L"published-before-consume;duplicate-rejected;clear-failure-terminal;retained-identity-rechecked;unknown-actions-withheld",
                      0u,
                      5u,
                      0u,
                      S_OK);
    NextStep(state, SelfTestState::Step::Phase10_ContentVerification);
    return false;
}
case SelfTestState::Step::Phase10_ContentVerification:
{
    const ULONGLONG nowTick               = GetTickCount64();
    const auto restoreVerificationSetting = [&]() noexcept
    {
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy =
            state.fileOperationsOriginal.has_value() ? state.fileOperationsOriginal->verifyAfterCopy : false;
    };
    const auto resetVerificationHooks = [&]() noexcept
    {
        SetFileOpsVerificationForceHostReadbackForSelfTest(0u);
        SetFileOpsVerificationForceUnavailableForSelfTest(0u);
        SetFileOpsVerificationForceMismatchForSelfTest(0u);
        ReleaseFileOpsVerificationReadbackPauseForSelfTest();
    };
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        resetVerificationHooks();
        restoreVerificationSetting();
        Fail(std::format(L"Phase10_ContentVerification timed out at substep {}.", state.stepState));
        return true;
    }

    const std::filesystem::path root                       = state.tempRoot / L"p4-content-verification";
    const std::filesystem::path offSource                  = root / L"off-source.bin";
    const std::filesystem::path offDestination             = root / L"off-destination";
    const std::filesystem::path proofSource                = root / L"proof-source.bin";
    const std::filesystem::path proofDestination           = root / L"proof-destination";
    const std::filesystem::path zeroSource                 = root / L"zero-source.bin";
    const std::filesystem::path zeroDestination            = root / L"zero-destination";
    const std::filesystem::path readbackSource             = root / L"readback-source.bin";
    const std::filesystem::path readbackDestination        = root / L"readback-destination";
    const std::filesystem::path mismatchSource             = root / L"mismatch-source" / L"tree";
    const std::filesystem::path mismatchDestination        = root / L"mismatch-destination";
    const std::filesystem::path unavailableSource          = root / L"unavailable-source" / L"tree";
    const std::filesystem::path unavailableDestination     = root / L"unavailable-destination";
    const std::filesystem::path unavailableCopySource      = root / L"unavailable-copy-source.bin";
    const std::filesystem::path unavailableCopyDestination = root / L"unavailable-copy-destination";
    const std::filesystem::path cancelSource               = root / L"cancel-source" / L"tree";
    const std::filesystem::path cancelDestination          = root / L"cancel-destination";
    const std::filesystem::path serializedSourceA          = root / L"serialized-a.bin";
    const std::filesystem::path serializedSourceB          = root / L"serialized-b.bin";
    const std::filesystem::path serializedDestination      = root / L"serialized-destination";
    const std::filesystem::path directorySource            = root / L"directory-source" / L"tree";
    const std::filesystem::path directoryDestination       = root / L"directory-destination";
    const std::filesystem::path nativeSource               = root / L"native-source.bin";
    const std::filesystem::path nativeDestination          = root / L"native-destination";

    const auto completedInfo = [&](const std::optional<uint64_t>& taskId) noexcept -> const CompletedTaskInfo*
    {
        if (! taskId.has_value())
        {
            return nullptr;
        }
        const auto found = state.completedTasks.find(taskId.value());
        return found != state.completedTasks.end() ? std::addressof(found->second) : nullptr;
    };
    const auto exactResult = [](const CompletedTaskInfo& info) noexcept -> const FileOperations::FileOperationItemResult*
    {
        return info.sourceItemResults.size() == 1u && info.sourceItemResults.front().has_value() ? std::addressof(info.sourceItemResults.front().value())
                                                                                                 : nullptr;
    };
    const auto startLocalTransfer =
        [&](FileSystemOperation operation, const std::filesystem::path& source, const std::filesystem::path& destination, FileSystemFlags flags) noexcept
    {
        return StartFileOperationAndGetId(state.fileOps,
                                          operation,
                                          FolderWindow::Pane::Left,
                                          FolderWindow::Pane::Right,
                                          state.fsLocal,
                                          {source},
                                          destination,
                                          flags,
                                          false,
                                          0u,
                                          FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                          false,
                                          nullptr);
    };
    // Beeline: a same-endpoint Move is a rename with nothing to verify. The Managed Move claims
    // below (verify before the exact source delete) run across endpoints, into the loopback share.
    const auto startBridgeMove = [&](const std::filesystem::path& source, const std::filesystem::path& destination) noexcept
    { return startLocalTransfer(FILESYSTEM_MOVE, source, LoopbackShareAlias(destination), FILESYSTEM_FLAG_RECURSIVE); };
    // Two regular files verified one at a time: the serialized-readback claim is endpoint-neutral.
    const auto startSerializedVerification = [&]() noexcept -> bool
    {
        SetFileOpsVerificationForceHostReadbackForSelfTest(2u);
        SetFileOpsVerificationReadbackPauseForSelfTest(true);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {serializedSourceA, serializedSourceB},
                                                 serializedDestination,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 nullptr);
        if (! state.taskA.has_value())
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 serialized transfer/verification fixture was not admitted.");
            return true;
        }
        state.stepState = 9u;
        return false;
    };

    if (state.stepState == 0u)
    {
        constexpr Common::Crypto::Blake3Digest kExpectedEmptyDigest{{
            std::byte{0xaf}, std::byte{0x13}, std::byte{0x49}, std::byte{0xb9}, std::byte{0xf5}, std::byte{0xf9}, std::byte{0xa1}, std::byte{0xa6},
            std::byte{0xa0}, std::byte{0x40}, std::byte{0x4d}, std::byte{0xea}, std::byte{0x36}, std::byte{0xdc}, std::byte{0xc9}, std::byte{0x49},
            std::byte{0x9b}, std::byte{0xcb}, std::byte{0x25}, std::byte{0xc9}, std::byte{0xad}, std::byte{0xc1}, std::byte{0x12}, std::byte{0xb7},
            std::byte{0xcc}, std::byte{0x9a}, std::byte{0x93}, std::byte{0xca}, std::byte{0xe4}, std::byte{0x1f}, std::byte{0x32}, std::byte{0x62},
        }};
        Common::Crypto::Blake3Hasher emptyHasher;
        if (emptyHasher.Finalize() != kExpectedEmptyDigest)
        {
            restoreVerificationSetting();
            Fail(L"P4.6 canonical BLAKE3 helper failed the empty-input known-answer test.");
            return true;
        }

        state.taskA.reset();
        state.taskB.reset();
        state.taskC.reset();
        resetVerificationHooks();
        if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(offDestination) || ! RecreateEmptyDirectory(proofDestination) ||
            ! RecreateEmptyDirectory(zeroDestination) || ! RecreateEmptyDirectory(readbackDestination) || ! RecreateEmptyDirectory(mismatchSource) ||
            ! RecreateEmptyDirectory(mismatchDestination / mismatchSource.filename()) || ! RecreateEmptyDirectory(unavailableSource) ||
            ! RecreateEmptyDirectory(unavailableDestination / unavailableSource.filename()) || ! RecreateEmptyDirectory(unavailableCopyDestination) ||
            ! RecreateEmptyDirectory(cancelSource) || ! RecreateEmptyDirectory(cancelDestination / cancelSource.filename()) ||
            ! RecreateEmptyDirectory(serializedDestination) || ! RecreateEmptyDirectory(directorySource) || ! RecreateEmptyDirectory(directoryDestination) ||
            ! RecreateEmptyDirectory(nativeDestination) || ! WriteTestFile(offSource, 64u * 1024u) || ! WriteTestFile(proofSource, 512u * 1024u) ||
            ! WriteTestFile(zeroSource, 0u) || ! WriteTestFile(readbackSource, 2u * 1024u * 1024u) ||
            ! WriteTestFile(mismatchSource / L"payload.bin", 256u * 1024u) || ! WriteTestFile(unavailableSource / L"payload.bin", 256u * 1024u) ||
            ! WriteTestFile(unavailableCopySource, 128u * 1024u) || ! WriteTestFile(cancelSource / L"payload.bin", 32u * 1024u * 1024u) ||
            ! WriteTestFile(serializedSourceA, 8u * 1024u * 1024u) || ! WriteTestFile(serializedSourceB, 8u * 1024u * 1024u) ||
            ! WriteTestFile(directorySource / L"payload.bin", 256u * 1024u) || ! WriteTestFile(nativeSource, 64u * 1024u))
        {
            restoreVerificationSetting();
            Fail(L"P4.6 content-verification fixtures could not be created.");
            return true;
        }

        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
        state.taskA                                               = startLocalTransfer(FILESYSTEM_COPY, offSource, offDestination, FILESYSTEM_FLAG_NONE);
        if (! state.taskA.has_value())
        {
            restoreVerificationSetting();
            Fail(L"P4.6 verification-off Copy was not admitted.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }

    if (state.stepState == 1u)
    {
        const CompletedTaskInfo* const completed = completedInfo(state.taskA);
        if (completed == nullptr)
        {
            return false;
        }
        const auto* const result = exactResult(*completed);
        if (completed->hr != S_OK || result == nullptr || result->verification != FileOperations::VerificationState::NotRequested ||
            completed->verificationUs != 0u || completed->verificationReadBytes != 0u || completed->verificationReadCalls != 0u ||
            completed->verificationProviderProofCount != 0u || completed->verificationHostReadbackCount != 0u)
        {
            restoreVerificationSetting();
            Fail(L"P4.6 verification-off Copy performed verification work or reported the wrong state.");
            return true;
        }

        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = true;
        state.taskA                                               = startLocalTransfer(FILESYSTEM_COPY, proofSource, proofDestination, FILESYSTEM_FLAG_NONE);
        if (! state.taskA.has_value())
        {
            restoreVerificationSetting();
            Fail(L"P4.6 provider-proof Copy was not admitted.");
            return true;
        }
        state.stepState = 2u;
        return false;
    }

    if (state.stepState == 2u)
    {
        const CompletedTaskInfo* const completed = completedInfo(state.taskA);
        if (completed == nullptr)
        {
            return false;
        }
        const auto* const result = exactResult(*completed);
        if (completed->hr != S_OK || result == nullptr || result->verification != FileOperations::VerificationState::Verified ||
            completed->verificationProviderProofCount != 1u || completed->verificationHostReadbackCount != 0u || completed->verificationReadCalls != 0u ||
            completed->verificationCompletedBytes != 512u * 1024u)
        {
            restoreVerificationSetting();
            Fail(L"P4.6 Local provider-proof Copy did not verify the exact published object without a host reread.");
            return true;
        }

        state.taskA = startLocalTransfer(FILESYSTEM_COPY, zeroSource, zeroDestination, FILESYSTEM_FLAG_NONE);
        if (! state.taskA.has_value())
        {
            restoreVerificationSetting();
            Fail(L"P4.6 zero-byte verification Copy was not admitted.");
            return true;
        }
        state.stepState = 3u;
        return false;
    }

    if (state.stepState == 3u)
    {
        const CompletedTaskInfo* const completed = completedInfo(state.taskA);
        if (completed == nullptr)
        {
            return false;
        }
        const auto* const result = exactResult(*completed);
        if (completed->hr != S_OK || result == nullptr || result->verification != FileOperations::VerificationState::Verified ||
            completed->verificationProviderProofCount != 1u || completed->verificationReadBytes != 0u)
        {
            restoreVerificationSetting();
            Fail(L"P4.6 zero-byte published object did not complete the selected proof contract.");
            return true;
        }

        SetFileOpsVerificationForceHostReadbackForSelfTest(1u);
        state.taskA = startLocalTransfer(FILESYSTEM_COPY, readbackSource, readbackDestination, FILESYSTEM_FLAG_NONE);
        if (! state.taskA.has_value())
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 host-readback verification Copy was not admitted.");
            return true;
        }
        state.stepState = 4u;
        return false;
    }

    if (state.stepState == 4u)
    {
        const CompletedTaskInfo* const completed = completedInfo(state.taskA);
        if (completed == nullptr)
        {
            return false;
        }
        const auto* const result = exactResult(*completed);
        if (completed->hr != S_OK || result == nullptr || result->verification != FileOperations::VerificationState::Verified ||
            TakeFileOpsVerificationForceHostReadbackAttemptsForSelfTest() != 1u || completed->verificationProviderProofCount != 0u ||
            completed->verificationHostReadbackCount != 1u || completed->verificationReadCalls == 0u || completed->verificationReadBytes != 2u * 1024u * 1024u)
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 bound-object host readback did not produce a verified BLAKE3 result.");
            return true;
        }

        if (! LoopbackShareReachable(state.tempRoot))
        {
            AppendLog(L"P4.6: the loopback administrative share is not reachable; the Managed Move verification sub-steps are environment-gated and skipped.");
            SetFileOpsVerificationReadbackPauseForSelfTest(false);
            return startSerializedVerification();
        }
        state.taskB = state.taskA;
        SetFileOpsVerificationForceMismatchForSelfTest(1u);
        state.taskA = startBridgeMove(mismatchSource, mismatchDestination);
        if (! state.taskA.has_value())
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 mismatch Managed Move was not admitted.");
            return true;
        }
        state.stepState = 5u;
        return false;
    }

    if (state.stepState == 5u)
    {
        const CompletedTaskInfo* const completed = completedInfo(state.taskA);
        if (completed == nullptr)
        {
            return false;
        }
        const auto* const result = exactResult(*completed);
        std::error_code existsError;
        const unsigned int mismatchAttempts = TakeFileOpsVerificationForceMismatchAttemptsForSelfTest();
        const bool sourceStillExists        = std::filesystem::exists(mismatchSource / L"payload.bin", existsError) && ! existsError;
        if (result == nullptr || result->verification != FileOperations::VerificationState::NotApplicable ||
            result->publication != FileOperations::PublicationState::Published || result->sourceDisposition != FileOperations::SourceDisposition::Retained ||
            result->completion != FileOperations::ItemCompletion::Failed || mismatchAttempts != 1u || ! sourceStillExists)
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(std::format(L"P4.6 verification mismatch did not retain the exact Managed-Move source and report Published + Failed "
                             L"while keeping the selected-directory verification axis NotApplicable "
                             L"(result={}, verification={}, publication={}, source={}, completion={}, attempts={}, exists={}, hr=0x{:08X}).",
                             result != nullptr,
                             result != nullptr ? static_cast<unsigned int>(result->verification) : 0xffu,
                             result != nullptr ? static_cast<unsigned int>(result->publication) : 0xffu,
                             result != nullptr ? static_cast<unsigned int>(result->sourceDisposition) : 0xffu,
                             result != nullptr ? static_cast<unsigned int>(result->completion) : 0xffu,
                             mismatchAttempts,
                             sourceStillExists,
                             static_cast<unsigned long>(completed->hr)));
            return true;
        }

        SetFileOpsVerificationForceUnavailableForSelfTest(2u);
        state.taskA = startBridgeMove(unavailableSource, unavailableDestination);
        state.taskC = startLocalTransfer(FILESYSTEM_COPY, unavailableCopySource, unavailableCopyDestination, FILESYSTEM_FLAG_NONE);
        if (! state.taskA.has_value() || ! state.taskC.has_value())
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 unavailable-proof Copy/Managed Move pair was not admitted.");
            return true;
        }
        state.stepState = 6u;
        return false;
    }

    if (state.stepState == 6u)
    {
        const CompletedTaskInfo* const completed     = completedInfo(state.taskA);
        const CompletedTaskInfo* const copyCompleted = completedInfo(state.taskC);
        if (completed == nullptr || copyCompleted == nullptr)
        {
            return false;
        }
        const auto* const result     = exactResult(*completed);
        const auto* const copyResult = exactResult(*copyCompleted);
        std::error_code existsError;
        if (completed->hr != S_FALSE || result == nullptr || result->verification != FileOperations::VerificationState::NotApplicable ||
            result->publication != FileOperations::PublicationState::Published || result->sourceDisposition != FileOperations::SourceDisposition::Retained ||
            copyCompleted->hr != S_FALSE || copyResult == nullptr || copyResult->verification != FileOperations::VerificationState::Unavailable ||
            copyResult->publication != FileOperations::PublicationState::Published ||
            copyResult->sourceDisposition != FileOperations::SourceDisposition::Retained ||
            copyResult->completion != FileOperations::ItemCompletion::Completed || TakeFileOpsVerificationForceUnavailableAttemptsForSelfTest() != 2u ||
            ! std::filesystem::exists(unavailableSource / L"payload.bin", existsError) || existsError)
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 unavailable verification was not typed truthfully as source-retained S_FALSE while keeping the "
                 L"selected-directory verification axis NotApplicable.");
            return true;
        }

        SetFileOpsVerificationForceHostReadbackForSelfTest(1u);
        SetFileOpsVerificationReadbackPauseForSelfTest(true);
        state.taskA = startBridgeMove(cancelSource, cancelDestination);
        if (! state.taskA.has_value())
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 cancel-during-readback Managed Move was not admitted.");
            return true;
        }
        state.stepState = 7u;
        return false;
    }

    if (state.stepState == 7u)
    {
        if (! HasFileOpsVerificationReadbackPauseEnteredForSelfTest())
        {
            return false;
        }
        auto* const task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (task == nullptr || ! task->_verificationActive.load(std::memory_order_acquire) ||
            task->_verificationCompletedBytes.load(std::memory_order_acquire) == 0u || task->_verificationReadCount.load(std::memory_order_acquire) == 0u)
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 readback pause was not observable as active verification with real progress.");
            return true;
        }
        task->RequestCancel();
        ReleaseFileOpsVerificationReadbackPauseForSelfTest();
        state.stepState = 8u;
        return false;
    }

    if (state.stepState == 8u)
    {
        const CompletedTaskInfo* const completed = completedInfo(state.taskA);
        if (completed == nullptr)
        {
            return false;
        }
        const auto* const result = exactResult(*completed);
        std::error_code existsError;
        if (result == nullptr || result->verification != FileOperations::VerificationState::NotApplicable ||
            result->publication != FileOperations::PublicationState::Published || result->sourceDisposition != FileOperations::SourceDisposition::Retained ||
            result->completion != FileOperations::ItemCompletion::Canceled || completed->verificationReadCalls == 0u ||
            completed->verificationReadBytes == 0u || ! std::filesystem::exists(cancelSource / L"payload.bin", existsError) || existsError)
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(std::format(L"P4.6 cancel during readback mismatch: result={} verification={} publication={} source={} completion={} "
                             L"reads={} bytes={} sourceExists={} existsError={} taskHr=0x{:08X}.",
                             result != nullptr,
                             result != nullptr ? static_cast<unsigned int>(result->verification) : 0u,
                             result != nullptr ? static_cast<unsigned int>(result->publication) : 0u,
                             result != nullptr ? static_cast<unsigned int>(result->sourceDisposition) : 0u,
                             result != nullptr ? static_cast<unsigned int>(result->completion) : 0u,
                             completed->verificationReadCalls,
                             completed->verificationReadBytes,
                             std::filesystem::exists(cancelSource / L"payload.bin", existsError),
                             existsError.value(),
                             static_cast<unsigned long>(completed->hr)));
            return true;
        }
        static_cast<void>(TakeFileOpsVerificationForceHostReadbackAttemptsForSelfTest());
        SetFileOpsVerificationReadbackPauseForSelfTest(false);
        return startSerializedVerification();
    }

    if (state.stepState == 9u)
    {
        if (! HasFileOpsVerificationReadbackPauseEnteredForSelfTest())
        {
            return false;
        }
        auto* const task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        std::error_code existsErrorA;
        std::error_code existsErrorB;
        const bool destinationAExists = std::filesystem::exists(serializedDestination / serializedSourceA.filename(), existsErrorA);
        const bool destinationBExists = std::filesystem::exists(serializedDestination / serializedSourceB.filename(), existsErrorB);
        if (task == nullptr || task->_perItemMaxConcurrencyBudget != 1u || task->_perItemMaxConcurrency != 1u || existsErrorA || existsErrorB ||
            destinationAExists == destinationBExists)
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 allowed another transfer to publish while the current item was being verified.");
            return true;
        }
        ReleaseFileOpsVerificationReadbackPauseForSelfTest();
        state.stepState = 10u;
        return false;
    }

    if (state.stepState == 10u)
    {
        const CompletedTaskInfo* const serializedCompleted = completedInfo(state.taskA);
        if (serializedCompleted == nullptr)
        {
            return false;
        }
        const bool allVerified = serializedCompleted->sourceItemResults.size() == 2u && std::ranges::all_of(serializedCompleted->sourceItemResults,
                                                                                                            [](const auto& item) noexcept
        { return item.has_value() && item->verification == FileOperations::VerificationState::Verified; });
        if (serializedCompleted->hr != S_OK || ! allVerified || serializedCompleted->verificationHostReadbackCount != 2u ||
            TakeFileOpsVerificationForceHostReadbackAttemptsForSelfTest() != 2u)
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 serialized verification task did not verify both published items exactly once.");
            return true;
        }
        SetFileOpsVerificationReadbackPauseForSelfTest(false);

        state.taskA = startLocalTransfer(FILESYSTEM_COPY, directorySource, directoryDestination, FILESYSTEM_FLAG_RECURSIVE);
        if (! state.taskA.has_value())
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 directory verification-axis fixture was not admitted.");
            return true;
        }
        state.stepState = 11u;
        return false;
    }

    if (state.stepState == 11u)
    {
        const CompletedTaskInfo* const directoryCompleted = completedInfo(state.taskA);
        if (directoryCompleted == nullptr)
        {
            return false;
        }
        const auto* const directoryResult = exactResult(*directoryCompleted);
        if (directoryCompleted->hr != S_OK || directoryResult == nullptr || directoryResult->verification != FileOperations::VerificationState::NotApplicable ||
            directoryCompleted->verificationProviderProofCount + directoryCompleted->verificationHostReadbackCount == 0u)
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(std::format(L"P4.6 selected directory verification mismatch: taskHr=0x{:08X} result={} verification={} "
                             L"providerProofs={} hostReadbacks={}.",
                             static_cast<unsigned long>(directoryCompleted->hr),
                             directoryResult != nullptr,
                             directoryResult != nullptr ? static_cast<unsigned int>(directoryResult->verification) : 0u,
                             directoryCompleted->verificationProviderProofCount,
                             directoryCompleted->verificationHostReadbackCount));
            return true;
        }

        state.taskA = startLocalTransfer(FILESYSTEM_MOVE, nativeSource, nativeDestination, FILESYSTEM_FLAG_NONE);
        if (! state.taskA.has_value())
        {
            resetVerificationHooks();
            restoreVerificationSetting();
            Fail(L"P4.6 Native Move NotApplicable fixture was not admitted.");
            return true;
        }
        state.stepState = 12u;
        return false;
    }

    const CompletedTaskInfo* const completed = completedInfo(state.taskA);
    if (completed == nullptr)
    {
        return false;
    }
    const auto* const result                         = exactResult(*completed);
    const CompletedTaskInfo* const readbackCompleted = completedInfo(state.taskB);
    const bool presentationValid                     = DebugValidateFileOperationsVerificationPresentation();
    if (completed->hr != S_OK || result == nullptr || result->strategy != FileOperations::OperationStrategy::Native ||
        result->verification != FileOperations::VerificationState::NotApplicable || completed->verificationReadBytes != 0u ||
        completed->verificationReadCalls != 0u || completed->verificationProviderProofCount != 0u || completed->verificationHostReadbackCount != 0u ||
        readbackCompleted == nullptr || readbackCompleted->verificationReadBytes != 2u * 1024u * 1024u || readbackCompleted->verificationReadCalls == 0u ||
        ! presentationValid)
    {
        resetVerificationHooks();
        restoreVerificationSetting();
        Fail(std::format(L"P4.6 Native/UI mismatch: result={} strategy={} verification={} reads={} bytes={} proofs={} readbacks={} "
                         L"readbackResult={} readbackReads={} readbackBytes={} presentation={} taskHr=0x{:08X}.",
                         result != nullptr,
                         result != nullptr ? static_cast<unsigned int>(result->strategy) : 0u,
                         result != nullptr ? static_cast<unsigned int>(result->verification) : 0u,
                         completed->verificationReadCalls,
                         completed->verificationReadBytes,
                         completed->verificationProviderProofCount,
                         completed->verificationHostReadbackCount,
                         readbackCompleted != nullptr,
                         readbackCompleted != nullptr ? readbackCompleted->verificationReadCalls : 0u,
                         readbackCompleted != nullptr ? readbackCompleted->verificationReadBytes : 0u,
                         presentationValid,
                         static_cast<unsigned long>(completed->hr)));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.ContentVerification",
                      L"off=zero-work;provider-proof=blake3;host-readback=blake3;mismatch-source-kept;unavailable=typed-sfalse;cancel-source-kept;native=not-"
                      L"applicable;ui=two-segment",
                      readbackCompleted->verificationUs,
                      readbackCompleted->verificationReadBytes,
                      readbackCompleted->verificationReadCalls,
                      S_OK);
    resetVerificationHooks();
    restoreVerificationSetting();
    NextStep(state, SelfTestState::Step::Phase10_ArtifactTouchGuard);
    return false;
}
case SelfTestState::Step::Phase10_ArtifactTouchGuard:
{
    const std::filesystem::path root     = state.tempRoot / L"p5-artifact-touch";
    const std::filesystem::path ordinary = root / L"ordinary.bin";
    const std::filesystem::path possible = root / L"payload.rs_ren_0123456789abcdef0123456789abcdef";
    if (! RecreateEmptyDirectory(root) || ! WriteTestFile(ordinary, 64u) || ! WriteTestFile(possible, 64u))
    {
        Fail(L"P5.4 external artifact-touch fixture setup failed.");
        return true;
    }

    HostResetTestPromptRequestCount();
    const std::array ordinaryPaths{ordinary};
    const HRESULT ordinaryHr = state.fileOps->ConfirmExternalArtifactTouch(state.fsLocal, kPluginIdLocal, {}, ordinaryPaths);
    if (ordinaryHr != S_OK || HostGetTestPromptRequestCount() != 0u)
    {
        Fail(std::format(L"P5.4 ordinary external touch was not prompt-free: hr=0x{:08X} prompts={}.",
                         static_cast<unsigned long>(ordinaryHr),
                         HostGetTestPromptRequestCount()));
        return true;
    }

    const std::array possiblePaths{possible};
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    const HRESULT cancelHr = state.fileOps->ConfirmExternalArtifactTouch(state.fsLocal, kPluginIdLocal, {}, possiblePaths);
    HostClearTestPromptResultOverride();
    HostPromptDebugSnapshot cancelPrompt{};
    if (cancelHr != S_FALSE || HostGetTestPromptRequestCount() != 1u || ! HostGetTestLastPromptDebugSnapshot(cancelPrompt) ||
        cancelPrompt.presentation != HOST_PROMPT_PRESENTATION_ARTIFACT_TOUCH || cancelPrompt.buttons != HOST_PROMPT_BUTTONS_OK_CANCEL ||
        cancelPrompt.defaultResult != HOST_PROMPT_RESULT_CANCEL)
    {
        Fail(std::format(L"P5.4 external artifact Cancel contract mismatch: hr=0x{:08X} prompts={} presentation={} buttons={} default={}.",
                         static_cast<unsigned long>(cancelHr),
                         HostGetTestPromptRequestCount(),
                         static_cast<unsigned int>(cancelPrompt.presentation),
                         static_cast<unsigned int>(cancelPrompt.buttons),
                         static_cast<unsigned int>(cancelPrompt.defaultResult)));
        return true;
    }

    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    FileOperationArtifacts::TouchGuardReceipt workerReceipt;
    const HRESULT proceedHr = state.fileOps->ConfirmExternalArtifactTouch(state.fsLocal, kPluginIdLocal, {}, possiblePaths, &workerReceipt);
    HostClearTestPromptResultOverride();
    const HRESULT workerRevalidateHr =
        FileOperationArtifacts::RevalidateProviderTouchGuard(state.fsLocal.get(), kPluginIdLocal, {}, workerReceipt, possiblePaths);
    if (proceedHr != S_OK || workerReceipt.items.size() != 1u || workerRevalidateHr != S_OK || HostGetTestPromptRequestCount() != 1u ||
        ! std::filesystem::exists(possible))
    {
        Fail(std::format(L"P5.4 accepted external artifact touch did not revalidate cleanly: hr=0x{:08X} receipt={} workerHr=0x{:08X} prompts={} exists={}.",
                         static_cast<unsigned long>(proceedHr),
                         workerReceipt.items.size(),
                         static_cast<unsigned long>(workerRevalidateHr),
                         HostGetTestPromptRequestCount(),
                         std::filesystem::exists(possible)));
        return true;
    }

    std::error_code replaceError;
    if (! std::filesystem::remove(possible, replaceError) || replaceError || ! WriteTestFile(possible, 64u))
    {
        Fail(L"P5.4 external artifact replacement-race fixture could not replace the accepted object.");
        return true;
    }
    const HRESULT replacedHr = FileOperationArtifacts::RevalidateProviderTouchGuard(state.fsLocal.get(), kPluginIdLocal, {}, workerReceipt, possiblePaths);
    if (replacedHr != HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH))
    {
        Fail(std::format(L"P5.4 worker receipt accepted a replacement object: hr=0x{:08X}.", static_cast<unsigned long>(replacedHr)));
        return true;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.ArtifactTouchGuard",
                      L"ordinary=prompt-free;possible=cancel-default;accepted=worker-revalidated;replacement=blocked",
                      0u,
                      3u,
                      HostGetTestPromptRequestCount(),
                      S_OK);
    NextStep(state, SelfTestState::Step::Phase11_CrossFileSystemBridge);
    return false;
}

case SelfTestState::Step::R1d_PreparingLifecycle:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        state.r1dPreparingGateRelease.store(true, std::memory_order_release);
        state.r1dPreparingGateRelease.notify_all();
        state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(std::nullopt);
        Fail(L"R1d_PreparingLifecycle timed out.");
        return true;
    }

    const std::filesystem::path sourceRoot         = state.tempRoot / L"r1d-preparing-source";
    const std::filesystem::path destinationRoot    = state.tempRoot / L"r1d-preparing-destination";
    const std::filesystem::path copySource         = sourceRoot / L"copy.bin";
    const std::filesystem::path fastCopySource     = sourceRoot / L"fast-copy.bin";
    const std::filesystem::path moveSource         = sourceRoot / L"move.bin";
    const std::filesystem::path failedGateSource   = sourceRoot / L"failed-gate.bin";
    const std::filesystem::path canceledGateSource = sourceRoot / L"canceled-gate.bin";
    constexpr uint32_t kMoveClipboardSequence      = 0x41D10001u;

    if (state.stepState == 0)
    {
        state.taskA.reset();
        state.taskB.reset();
        state.taskC.reset();
        state.r1dPreparingGateEntered.store(false, std::memory_order_release);
        state.r1dPreparingGateRelease.store(false, std::memory_order_release);
        state.r1dAdmissionReturnedBeforeReadiness   = false;
        state.r1dMoveBreadcrumbPublishedBeforeReady = false;
        if (! RecreateEmptyDirectory(sourceRoot) || ! RecreateEmptyDirectory(destinationRoot) || ! WriteTestFile(copySource, 4096u) ||
            ! WriteTestFile(fastCopySource, 4096u) || ! WriteTestFile(moveSource, 4096u) || ! WriteTestFile(failedGateSource, 4096u) ||
            ! WriteTestFile(canceledGateSource, 4096u))
        {
            Fail(L"R1d Preparing fixtures could not be created.");
            return true;
        }

        constexpr size_t kSelectedRootCount = 4'096u;
        FileOperations::QualifiedEndpoint sourceEndpoint{};
        FileOperations::QualifiedEndpoint destinationEndpoint{};
        if (! FileOperations::TryQualifyEndpointForSelfTest(
                state.fsLocal, sourceRoot.native(), FILESYSTEM_COPY, kPluginIdLocal, L"host/default", sourceEndpoint) ||
            ! FileOperations::TryQualifyEndpointForSelfTest(
                state.fsLocal, destinationRoot.native(), FILESYSTEM_COPY, kPluginIdLocal, L"host/default", destinationEndpoint))
        {
            Fail(L"R1d Preparing performance endpoints could not be qualified.");
            return true;
        }

        FileOperations::TransferPlan perfPlan{};
        perfPlan.intent                         = FileOperations::TransferIntent::Copy;
        perfPlan.strategy                       = FileOperations::OperationStrategy::Copy;
        perfPlan.sourceEndpoint                 = std::move(sourceEndpoint);
        perfPlan.destinationEndpoint            = std::move(destinationEndpoint);
        perfPlan.destination.providerFolderPath = destinationRoot.native();
        perfPlan.selectedItems.reserve(kSelectedRootCount);
        for (size_t index = 0u; index < kSelectedRootCount; ++index)
        {
            perfPlan.selectedItems.emplace_back(FileOperations::QualifiedSourceItem{
                .providerPath = (sourceRoot / std::format(L"selected-{:04}.bin", index)).native(),
            });
        }
        auto perfPlans = std::make_shared<FileOperations::FileOperationPlanGroup>();
        perfPlans->emplace_back(std::move(perfPlan));

        FolderWindow::FileOperationState::Task perfTask(*state.fileOps);
        perfTask._operation             = FILESYSTEM_COPY;
        perfTask._fileSystem            = state.fsLocal;
        perfTask._destinationFileSystem = state.fsLocal;
        perfTask.StorePlans(perfPlans);
        perfTask._mutationInterlockScopes.reserve(kSelectedRootCount * 2u);
        const auto* transferPlan = std::get_if<FileOperations::TransferPlan>(&perfPlans->front());
        if (transferPlan == nullptr)
        {
            Fail(L"R1d Preparing performance plan lost its typed Transfer shape.");
            return true;
        }
        for (const FileOperations::QualifiedSourceItem& item : transferPlan->selectedItems)
        {
            perfTask._mutationInterlockScopes.emplace_back(FileOperations::MutationInterlockScope{
                .endpoint     = transferPlan->sourceEndpoint,
                .providerPath = item.providerPath,
                .access       = FileOperations::MutationInterlockAccess::ReadSource,
            });
            std::wstring destinationPath;
            const size_t sourceIndex = static_cast<size_t>(&item - transferPlan->selectedItems.data());
            if (! FileOperations::TryResolveTransferDestinationProviderPath(*transferPlan, sourceIndex, destinationPath))
            {
                Fail(L"R1d Preparing performance destination mapping could not be resolved.");
                return true;
            }
            perfTask._mutationInterlockScopes.emplace_back(FileOperations::MutationInterlockScope{
                .endpoint     = transferPlan->destinationEndpoint,
                .providerPath = std::move(destinationPath),
                .access       = FileOperations::MutationInterlockAccess::PublishDestination,
            });
        }

        const auto readinessStarted                   = std::chrono::steady_clock::now();
        FileOperations::PlanRejectionBucket rejection = FileOperations::PlanRejectionBucket::None;
        HRESULT readinessHr                           = FileOperations::ValidatePlan(perfPlans->front(), &rejection);
        if (SUCCEEDED(readinessHr))
        {
            const uint64_t validationUs =
                static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - readinessStarted).count());
            readinessHr = perfTask.BuildPreparationSnapshot(validationUs);
        }
        if (FAILED(readinessHr))
        {
            Fail(std::format(L"R1d Preparing candidate snapshot failed: hr=0x{:08X}.", static_cast<unsigned long>(readinessHr)));
            return true;
        }
        const std::shared_ptr<const FileOperations::PreparationSnapshot> preparation = perfTask.LoadPreparationSnapshot();
        if (! preparation || preparation->selectedRootCount != kSelectedRootCount || preparation->scopes.size() != kSelectedRootCount * 2u ||
            preparation->endpoints.size() != 1u || preparation->strategies.size() != 1u || preparation->copyOnlyCount != 0u ||
            preparation->buildSnapshotUs >= 2'000'000u || preparation->selectedRootReadinessUs >= 2'000'000u ||
            preparation->retainedBytes >= 64u * 1024u * 1024u)
        {
            Fail(L"R1d Preparing candidate snapshot violated its 4096-root fact, time, or retention budget.");
            return true;
        }

        state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(10'000u);
        state.r1dPresentationBaseline = state.fileOps->DebugTaskPresentedCount();
        HostResetTestPromptRequestCount();
        state.taskA                                          = StartFileOperationAndGetId(state.fileOps,
                                                                                          FILESYSTEM_COPY,
                                                                                          FolderWindow::Pane::Left,
                                                                                          FolderWindow::Pane::Right,
                                                                                          state.fsLocal,
                                                                                          {copySource},
                                                                                          destinationRoot,
                                                                                          FILESYSTEM_FLAG_NONE,
                                                                                          false,
                                                                                          0u,
                                                                                          FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                                          false,
                                                                                          state.fsLocal,
                                                                                          {},
                                                                                          std::nullopt,
                                                                                          {},
                                                                                          [&]() noexcept
        {
            state.r1dPreparingGateEntered.store(true, std::memory_order_release);
            state.r1dPreparingGateEntered.notify_all();
            state.r1dPreparingGateRelease.wait(false, std::memory_order_acquire);
            return S_OK;
        });
        FolderWindow::FileOperationState::Task* admittedTask = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (! state.taskA.has_value() || admittedTask == nullptr || ! admittedTask->IsPresentationHidden() || HostGetTestPromptRequestCount() != 0u)
        {
            Fail(L"R1d routine Copy did not enter hidden common Preparing without a generic confirmation prompt.");
            return true;
        }

        FolderWindow::FileOperationState::Task* task = admittedTask;
        state.r1dAdmissionReturnedBeforeReadiness    = task != nullptr && ! task->_selectedRootReadinessComplete.load(std::memory_order_acquire);
        state.stepState                              = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (! state.r1dPreparingGateEntered.load(std::memory_order_acquire))
        {
            return false;
        }
        const FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const std::shared_ptr<const FileOperations::PreparationSnapshot> preparation = task == nullptr ? nullptr : task->LoadPreparationSnapshot();
        state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(10'499u);
        static_cast<void>(state.fileOps->RefreshDeferredTaskPresentation());
        const bool hiddenBeforeDeadline = task != nullptr && task->IsPresentationHidden();
        state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(10'500u);
        static_cast<void>(state.fileOps->RefreshDeferredTaskPresentation());
        const bool revealedAtDeadline =
            task != nullptr && task->IsPresentationVisible() && state.fileOps->DebugTaskPresentedCount() == state.r1dPresentationBaseline + 1u;
        if (! state.r1dAdmissionReturnedBeforeReadiness || task == nullptr ||
            task->GetLifecyclePhase() != FileOperations::TaskLifecyclePhase::AwaitingAcceptance || ! preparation || preparation->selectedRootCount != 1u ||
            preparation->scopes.size() != 2u || preparation->endpoints.size() != 1u || preparation->strategies.size() != 1u || preparation->status != S_OK ||
            ! hiddenBeforeDeadline || ! revealedAtDeadline || task->_selectedRootReadinessComplete.load(std::memory_order_acquire) ||
            std::filesystem::exists(destinationRoot / copySource.filename()))
        {
            state.r1dPreparingGateRelease.store(true, std::memory_order_release);
            state.r1dPreparingGateRelease.notify_all();
            Fail(L"R1d ordinary Copy crossed readiness or mutated while the common Preparing gate was unresolved.");
            return true;
        }
        state.r1dPreparingGateRelease.store(true, std::memory_order_release);
        state.r1dPreparingGateRelease.notify_all();
        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (! state.taskA.has_value() || ! state.completedTasks.contains(state.taskA.value()))
        {
            return false;
        }
        if (! std::filesystem::exists(destinationRoot / copySource.filename()))
        {
            Fail(L"R1d ordinary Copy did not execute after the common Preparing gate accepted it.");
            return true;
        }

        state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(20'000u);
        state.r1dPresentationBaseline = state.fileOps->DebugTaskPresentedCount();
        state.taskA                   = StartFileOperationAndGetId(state.fileOps,
                                                                   FILESYSTEM_COPY,
                                                                   FolderWindow::Pane::Left,
                                                                   FolderWindow::Pane::Right,
                                                                   state.fsLocal,
                                                                   {fastCopySource},
                                                                   destinationRoot,
                                                                   FILESYSTEM_FLAG_NONE,
                                                                   false,
                                                                   0u,
                                                                   FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                   false,
                                                                   state.fsLocal);
        if (! state.taskA.has_value())
        {
            Fail(L"R1d fast routine Copy could not be admitted for no-flash coverage.");
            return true;
        }
        state.stepState = 20;
        return false;
    }

    if (state.stepState == 20)
    {
        if (! state.taskA.has_value() || ! state.completedTasks.contains(state.taskA.value()))
        {
            return false;
        }
        if (! std::filesystem::exists(destinationRoot / fastCopySource.filename()) || state.fileOps->DebugTaskPresentedCount() != state.r1dPresentationBaseline)
        {
            Fail(L"R1d fast routine Copy did not complete cleanly without revealing a task surface.");
            return true;
        }

        state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(std::nullopt);
        state.r1dPreparingGateEntered.store(false, std::memory_order_release);
        state.r1dPreparingGateRelease.store(false, std::memory_order_release);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {moveSource},
                                                 destinationRoot,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 {},
                                                 {},
                                                 kMoveClipboardSequence,
                                                 []() noexcept { return S_OK; },
                                                 [&]() noexcept
        {
            state.r1dPreparingGateEntered.store(true, std::memory_order_release);
            state.r1dPreparingGateEntered.notify_all();
            state.r1dPreparingGateRelease.wait(false, std::memory_order_acquire);
            return S_OK;
        });
        if (! state.taskB.has_value())
        {
            Fail(L"R1d clipboard Move could not be admitted for breadcrumb ordering coverage.");
            return true;
        }
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (! state.r1dPreparingGateEntered.load(std::memory_order_acquire))
        {
            return false;
        }
        const FolderWindow::FileOperationState::Task* task = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
        state.r1dMoveBreadcrumbPublishedBeforeReady        = task != nullptr && task->_moveBreadcrumb.has_value();
        const bool mutationBeforeReady = ! std::filesystem::exists(moveSource) || std::filesystem::exists(destinationRoot / moveSource.filename());
        Debug::Perf::Emit(L"FileOps.Preparing.MutationBeforeReadyCount", L"behavioral", 0u, mutationBeforeReady ? 1u : 0u, 0u, S_OK);
        Debug::Perf::Emit(L"FileOps.Preparing.ClipboardBeforeReadyCount", L"behavioral", 0u, 0u, 0u, S_OK);
        Debug::Perf::Emit(L"FileOps.Preparing.BreadcrumbBeforeReadyCount", L"behavioral", 0u, state.r1dMoveBreadcrumbPublishedBeforeReady ? 1u : 0u, 0u, S_OK);
        state.r1dPreparingGateRelease.store(true, std::memory_order_release);
        state.r1dPreparingGateRelease.notify_all();
        if (task == nullptr || mutationBeforeReady || state.r1dMoveBreadcrumbPublishedBeforeReady)
        {
            Fail(L"R1d clipboard Move published mutation or its breadcrumb before Preparing and acceptance completed.");
            return true;
        }
        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        if (! state.taskB.has_value() || ! state.completedTasks.contains(state.taskB.value()))
        {
            return false;
        }
        if (std::filesystem::exists(moveSource) || ! std::filesystem::exists(destinationRoot / moveSource.filename()))
        {
            Fail(L"R1d clipboard Move did not execute after Preparing and clipboard acceptance completed.");
            return true;
        }

        state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {failedGateSource},
                                                 destinationRoot,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal,
                                                 {},
                                                 std::nullopt,
                                                 {},
                                                 []() noexcept { return E_ACCESSDENIED; });
        if (! state.taskC.has_value())
        {
            Fail(L"R1d failed-gate Copy could not be admitted.");
            return true;
        }
        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        if (! state.taskC.has_value() || ! state.completedTasks.contains(state.taskC.value()))
        {
            return false;
        }
        const CompletedTaskInfo& completed                    = state.completedTasks.at(state.taskC.value());
        const FileOperations::FileOperationItemResult* result = completed.sourceItemResults.size() == 1u && completed.sourceItemResults.front().has_value()
                                                                    ? std::addressof(completed.sourceItemResults.front().value())
                                                                    : nullptr;
        if (completed.hr != E_ACCESSDENIED || result == nullptr || result->publication != FileOperations::PublicationState::NotAttempted ||
            result->sourceDisposition != FileOperations::SourceDisposition::Retained || result->completion != FileOperations::ItemCompletion::Failed ||
            result->status != E_ACCESSDENIED || ! std::filesystem::exists(failedGateSource) ||
            std::filesystem::exists(destinationRoot / failedGateSource.filename()))
        {
            Fail(L"R1d failed decision gate did not reduce to one dense source-retained, no-mutation result.");
            return true;
        }

        state.r1dPreparingGateEntered.store(false, std::memory_order_release);
        state.r1dPreparingGateRelease.store(false, std::memory_order_release);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {canceledGateSource},
                                                 destinationRoot,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal,
                                                 {},
                                                 std::nullopt,
                                                 {},
                                                 [&]() noexcept
        {
            state.r1dPreparingGateEntered.store(true, std::memory_order_release);
            state.r1dPreparingGateEntered.notify_all();
            state.r1dPreparingGateRelease.wait(false, std::memory_order_acquire);
            return S_OK;
        });
        if (! state.taskA.has_value())
        {
            Fail(L"R1d cancel-during-gate Copy could not be admitted.");
            return true;
        }
        state.stepState = 6;
        return false;
    }

    if (state.stepState == 6)
    {
        if (! state.r1dPreparingGateEntered.load(std::memory_order_acquire))
        {
            return false;
        }
        FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (task == nullptr || task->GetLifecyclePhase() != FileOperations::TaskLifecyclePhase::AwaitingAcceptance ||
            task->_selectedRootReadinessComplete.load(std::memory_order_acquire))
        {
            state.r1dPreparingGateRelease.store(true, std::memory_order_release);
            state.r1dPreparingGateRelease.notify_all();
            Fail(L"R1d cancel-during-gate task did not remain in AwaitingAcceptance before cancellation.");
            return true;
        }
        task->RequestCancel();
        state.r1dPreparingGateRelease.store(true, std::memory_order_release);
        state.r1dPreparingGateRelease.notify_all();
        state.stepState = 7;
        return false;
    }

    if (! state.taskA.has_value() || ! state.completedTasks.contains(state.taskA.value()))
    {
        return false;
    }
    const CompletedTaskInfo& canceled                             = state.completedTasks.at(state.taskA.value());
    const FileOperations::FileOperationItemResult* canceledResult = canceled.sourceItemResults.size() == 1u && canceled.sourceItemResults.front().has_value()
                                                                        ? std::addressof(canceled.sourceItemResults.front().value())
                                                                        : nullptr;
    if (canceled.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || canceledResult == nullptr ||
        canceledResult->publication != FileOperations::PublicationState::NotAttempted ||
        canceledResult->sourceDisposition != FileOperations::SourceDisposition::Retained ||
        canceledResult->completion != FileOperations::ItemCompletion::Canceled || canceledResult->status != HRESULT_FROM_WIN32(ERROR_CANCELLED) ||
        ! std::filesystem::exists(canceledGateSource) || std::filesystem::exists(destinationRoot / canceledGateSource.filename()))
    {
        Fail(L"R1d canceled decision gate did not reduce to one dense source-retained, no-mutation result.");
        return true;
    }

    state.fileOps->DebugSetTaskPresentationNowTickForSelfTest(std::nullopt);
    NextStep(state, SelfTestState::Step::C1_ExitCloseDeferredUntilTasksQuiet);
    return false;
}

#endif
#if defined(FILEOPS_SELFTEST_INCLUDE_PHASE11)

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
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));

        // R3-2: this case covers Copy-only Move into a destination without content proof.
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(),
            R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0","writerProof":false})json"));

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
        HostResetTestPromptRequestCount();
        state.taskC = StartFileOperationAndGetId(state.fileOps,
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
        if (HostGetTestPromptRequestCount() != 1u)
        {
            Fail(std::format(L"Known Copy-only Move must receive exactly one acceptance before mutation; got {} prompts.", HostGetTestPromptRequestCount()));
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

        if (it->second.hr != S_FALSE)
        {
            Fail(std::format(L"Cross-filesystem CopyOnly move (local -> dummy) expected S_FALSE, got 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        std::error_code ec;
        if (! std::filesystem::exists(moveFile, ec))
        {
            Fail(L"Cross-filesystem CopyOnly move did not retain the source file.");
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

        if (prompt->bucket != Task::ConflictBucket::RegularFileExists)
        {
            Fail(L"Cross-filesystem overwrite test did not produce an Exists prompt.");
            return true;
        }

        // R3-1: a destination without object binding offers conditional Overwrite through its
        // atomic-final writer, alongside the non-destructive top-level Keep Both.
        if (! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite) || ! PromptHasAction(prompt.value(), Task::ConflictAction::KeepBoth))
        {
            Fail(L"An unbound destination must offer conditional Overwrite (R3-1) and non-destructive top-level Keep Both.");
            return true;
        }

        task->SubmitConflictDecision(Task::ConflictAction::KeepBoth, false);
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
            Fail(std::format(L"Cross-filesystem Keep Both test copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
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

        const std::wstring dummyOriginalPath = std::wstring(dummyCopyRoot) + L"/bridge-src/a.bin";
        std::wstring dummyKeepBothPath;
        if (FAILED(Common::Paths::BuildUniqueSiblingPathCandidate(dummyOriginalPath, false, 2u, dummyKeepBothPath)))
        {
            Fail(L"Cross-filesystem Keep Both test could not build the expected sibling path.");
            return true;
        }

        wil::com_ptr<IFileReader> reader;
        const HRESULT hrReader = dummyIo->CreateFileReader(dummyKeepBothPath.c_str(), reader.addressof());
        if (FAILED(hrReader) || ! reader)
        {
            Fail(L"Cross-filesystem Keep Both test: failed to open the new sibling in the dummy filesystem.");
            return true;
        }

        uint64_t sizeBytes = 0;
        if (FAILED(reader->GetSize(&sizeBytes)) || sizeBytes != 512ull)
        {
            Fail(L"Cross-filesystem Keep Both test: sibling file size mismatch.");
            return true;
        }
        reader.reset();
        if (FAILED(dummyIo->CreateFileReader(dummyOriginalPath.c_str(), reader.addressof())) || ! reader || FAILED(reader->GetSize(&sizeBytes)) ||
            sizeBytes != 128ull)
        {
            Fail(L"Cross-filesystem Keep Both test changed the unbound original destination.");
            return true;
        }

        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        state.fileOps->CollectCompletedTasks(summaries);
        const auto summary = std::find_if(
            summaries.begin(), summaries.end(), [&](const auto& value) noexcept { return state.taskA.has_value() && value.taskId == state.taskA.value(); });
        const bool reportedUnboundMetadataLoss = summary != summaries.end() && std::ranges::any_of(summary->issueDiagnostics, [](const auto& issue) noexcept {
            return issue.category == L"bridge.metadata.authorityUnavailable";
        });
        if (! reportedUnboundMetadataLoss)
        {
            Fail(L"Cross-filesystem metadata test did not report that exact destination authority is unavailable.");
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
                std::scoped_lock lock(task->_progressMutex, task->_perItemInFlightCallsMutex);
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
                std::scoped_lock lock(task->_inFlightFilesMutex, task->_perItemInFlightCallsMutex);
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
        SetFileOpsBridgeTraversalDepthLimitForSelfTest(0);
        Fail(L"Phase11_BridgeMultiFolderParallelCopyInFlightLines timed out.");
        return true;
    }

    const std::filesystem::path srcDir         = state.tempRoot / L"bridge-multifolder-src";
    const std::wstring dummyRoot               = L"/bridge-multifolder";
    const std::wstring traversalLimitDummyRoot = L"/bridge-traversal-limit";

    constexpr int kFolderCount    = 2;
    constexpr int kFilesPerFolder = 24;
    // File count drives queue saturation; a lower byte volume keeps this concurrency contract
    // deterministic without making fixture creation dominate the full FileOps suite.
    constexpr size_t kFileBytes             = 512ull * 1024ull;
    constexpr uint64_t kSpeedLimitBps       = 256ull * 1024ull;
    constexpr unsigned int kProducerDelayMs = 20u;

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));

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

        wil::com_ptr<IFileSystemIO> localIo;
        if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
        {
            Fail(L"Local filesystem does not support IFileSystemIO for bridge directory metadata setup.");
            return true;
        }
        FileSystemBasicInformation directoryBasic{};
        directoryBasic.sizeBytes                         = sizeof(FileSystemBasicInformation);
        const std::filesystem::path metadataSourceFolder = srcDir / L"mf_0";
        if (FAILED(localIo->GetFileBasicInformation(metadataSourceFolder.c_str(), &directoryBasic)))
        {
            Fail(L"Bridge directory metadata setup could not read the source folder metadata.");
            return true;
        }
        constexpr int64_t kBridgeDirectoryMetadataOffset = 2ll * 24ll * 60ll * 60ll * 10'000'000ll;
        directoryBasic.creationTime -= kBridgeDirectoryMetadataOffset;
        directoryBasic.lastWriteTime -= kBridgeDirectoryMetadataOffset;
        if (FAILED(localIo->SetFileBasicInformation(metadataSourceFolder.c_str(), &directoryBasic)))
        {
            Fail(L"Bridge directory metadata setup could not seed source timestamps.");
            return true;
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

        // The fixture writes 96 MiB and can be disproportionately slow under filesystem filters.
        // Measure the task timeout from admission so setup latency cannot hide a successful bridge run.
        state.stepStartTick = GetTickCount64();
        state.stepState     = 1;
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
                std::scoped_lock lock(task->_progressMutex, task->_inFlightFilesMutex, task->_perItemInFlightCallsMutex);
                inFlightCount     = task->_inFlightFileCount;
                inFlightItemCalls = task->_perItemInFlightCallCount;
                budget            = task->_perItemMaxConcurrencyBudget;
            }

            if (budget <= 1u)
            {
                SetFileOpsBridgeProducerDelayForSelfTest(0);
                Fail(L"Bridge multi-folder test expected bridge concurrency budget >1, but task budget is 1.");
                return true;
            }

            if (inFlightItemCalls < 2u)
            {
                if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                {
                    SetFileOpsBridgeProducerDelayForSelfTest(0);
                    Fail(std::format(L"Bridge multi-folder test expected >=2 top-level in-flight calls but observed {} (liveFiles={}, budget={}).",
                                     inFlightItemCalls,
                                     inFlightCount,
                                     budget));
                    return true;
                }
                return false;
            }

            if (inFlightCount == 0)
            {
                if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                {
                    SetFileOpsBridgeProducerDelayForSelfTest(0);
                    Fail(std::format(L"Bridge multi-folder test expected live child file progress while {} top-level calls were active (budget={}).",
                                     inFlightItemCalls,
                                     budget));
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

    if (state.stepState == 4)
    {
        SetFileOpsBridgeTraversalDepthLimitForSelfTest(0);
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));

        const HRESULT expectedLimitHr = HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW);
        if (it->second.hr != expectedLimitHr || it->second.progressCompletedItems != 1u || it->second.bridgeTraversalLimitHitCount == 0u ||
            it->second.bridgeTraversalMaxDepth != 2u)
        {
            Fail(std::format(L"Bridge traversal-limit result was not fail-closed with prior success retained: task=0x{:08X}, completed={}, "
                             L"depth={}, hits={}.",
                             static_cast<unsigned long>(it->second.hr),
                             it->second.progressCompletedItems,
                             it->second.bridgeTraversalMaxDepth,
                             it->second.bridgeTraversalLimitHitCount));
            return true;
        }

        wil::com_ptr<IFileSystemIO> dummyIo;
        if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
        {
            Fail(L"Dummy filesystem does not support IFileSystemIO for traversal-limit validation.");
            return true;
        }

        unsigned long attrs                 = 0;
        const std::wstring priorDestination = traversalLimitDummyRoot + L"/00-prior.bin";
        if (FAILED(dummyIo->GetAttributes(priorDestination.c_str(), &attrs)))
        {
            Fail(L"Bridge traversal-limit test did not retain the item completed before the bound was reached.");
            return true;
        }
        const std::wstring beyondLimitDestination = traversalLimitDummyRoot + L"/10-limit-tree/level-0/level-1/level-2/never.bin";
        if (SUCCEEDED(dummyIo->GetAttributes(beyondLimitDestination.c_str(), &attrs)))
        {
            Fail(L"Bridge traversal-limit test published an item below the rejected depth.");
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.BridgeTraversalLimitPartialPriorSuccess",
                          L"overrideDepth=1 expectedObservedDepth=2",
                          0u,
                          it->second.bridgeTraversalMaxDepth,
                          it->second.bridgeTraversalLimitHitCount,
                          it->second.hr);
        NextStep(state, SelfTestState::Step::Phase11_BridgePipelineDummyToDummyPerf);
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
    Debug::Perf::Emit(L"FileOps.SelfTest.BridgeTraversalRetentionHighWater",
                      std::format(L"depth={} retainedEntries={} queuedPathBytes={} metadataBytes={} limitHits={}",
                                  it->second.bridgeTraversalMaxDepth,
                                  it->second.bridgeTraversalMaxRetainedEntries,
                                  it->second.bridgeTraversalMaxQueuedPathBytes,
                                  it->second.bridgeTraversalMaxMetadataBytes,
                                  it->second.bridgeTraversalLimitHitCount),
                      it->second.bridgeTraversalMaxMetadataBytes,
                      it->second.bridgeTraversalMaxQueuedPathBytes,
                      it->second.bridgeTraversalMaxRetainedEntries,
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
    constexpr uint64_t kExpectedMaxBridgeAdmissionQueueDepth = 256u;
    if (it->second.bridgeAdmissionMaxQueueDepth > kExpectedMaxBridgeAdmissionQueueDepth)
    {
        Fail(std::format(L"Bridge discovery-ahead queue should stay bounded at {} entries; observed {}.",
                         kExpectedMaxBridgeAdmissionQueueDepth,
                         it->second.bridgeAdmissionMaxQueueDepth));
        return true;
    }
    constexpr uint64_t kTraversalHardEntryLimit    = 4'096u;
    constexpr uint64_t kTraversalHardPathBytes     = 16ull * 1024ull * 1024ull;
    constexpr uint64_t kTraversalHardMetadataBytes = 8ull * 1024ull * 1024ull;
    if (it->second.bridgeTraversalMaxRetainedEntries == 0u || it->second.bridgeTraversalMaxRetainedEntries > kTraversalHardEntryLimit ||
        it->second.bridgeTraversalMaxQueuedPathBytes == 0u || it->second.bridgeTraversalMaxQueuedPathBytes > kTraversalHardPathBytes ||
        it->second.bridgeTraversalMaxMetadataBytes == 0u || it->second.bridgeTraversalMaxMetadataBytes > kTraversalHardMetadataBytes ||
        it->second.bridgeTraversalLimitHitCount != 0u)
    {
        Fail(std::format(L"Bridge traversal retention was not bounded as specified: entries={}/{}, pathBytes={}/{}, metadataBytes={}/{}, limitHits={}.",
                         it->second.bridgeTraversalMaxRetainedEntries,
                         kTraversalHardEntryLimit,
                         it->second.bridgeTraversalMaxQueuedPathBytes,
                         kTraversalHardPathBytes,
                         it->second.bridgeTraversalMaxMetadataBytes,
                         kTraversalHardMetadataBytes,
                         it->second.bridgeTraversalLimitHitCount));
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

    const std::filesystem::path priorSource = srcDir / L"00-prior.bin";
    std::filesystem::path limitTree         = srcDir / L"10-limit-tree";
    std::filesystem::path limitDeep         = limitTree;
    for (unsigned int depth = 0; depth < 3u; ++depth)
    {
        limitDeep /= std::format(L"level-{}", depth);
    }
    if (! WriteTestFile(priorSource, 64u) || ! RecreateEmptyDirectory(limitDeep) || ! WriteTestFile(limitDeep / L"never.bin", 64u) ||
        ! EnsureDummyFolderExists(state.fsDummy.get(), traversalLimitDummyRoot))
    {
        Fail(L"Failed to prepare bridge traversal-limit partial-success fixture.");
        return true;
    }

    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> metadataSummaries;
    state.fileOps->CollectCompletedTasks(metadataSummaries);
    const auto metadataSummary                  = std::find_if(metadataSummaries.begin(), metadataSummaries.end(), [&](const auto& value) noexcept {
        return state.taskA.has_value() && value.taskId == state.taskA.value();
    });
    const bool reportedUnboundDirectoryMetadata = metadataSummary != metadataSummaries.end() && std::ranges::any_of(metadataSummary->issueDiagnostics,
                                                                                                                    [](const auto& issue) noexcept
    { return issue.category == L"bridge.directoryMetadata.authorityUnavailable"; });
    if (! reportedUnboundDirectoryMetadata)
    {
        Fail(L"Bridge did not report that exact destination-directory metadata authority is unavailable.");
        return true;
    }

    static_cast<void>(SetPluginConfiguration(
        state.infoLocal.get(),
        R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4})json"));
    SetFileOpsBridgeTraversalDepthLimitForSelfTest(1u);
    std::vector<std::filesystem::path> limitSources{priorSource, limitTree};
    const FileSystemFlags limitFlags =
        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
    state.taskA = StartFileOperationAndGetId(state.fileOps,
                                             FILESYSTEM_COPY,
                                             FolderWindow::Pane::Left,
                                             FolderWindow::Pane::Right,
                                             state.fsLocal,
                                             std::move(limitSources),
                                             std::filesystem::path(traversalLimitDummyRoot),
                                             limitFlags,
                                             false,
                                             0u,
                                             FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                             false,
                                             state.fsDummy);
    if (! state.taskA.has_value())
    {
        SetFileOpsBridgeTraversalDepthLimitForSelfTest(0);
        Fail(L"Failed to start bridge traversal-limit partial-success test.");
        return true;
    }
    state.stepState = 4;
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
    const std::wstring dummyBaselineDestinationRoot  = L"/bridge-pipeline-dst-baseline";
    const std::wstring dummyCandidateDestinationRoot = L"/bridge-pipeline-dst-candidate";

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

        const FileSystemFlags cleanupFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        static_cast<void>(state.fsDummy->DeleteItem(dummyBaselineDestinationRoot.c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        static_cast<void>(state.fsDummy->DeleteItem(dummyCandidateDestinationRoot.c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummySourceRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummySourceFolder) ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), dummyBaselineDestinationRoot) ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), dummyCandidateDestinationRoot))
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
                                                 std::filesystem::path(dummyBaselineDestinationRoot),
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
                                                 std::filesystem::path(dummyCandidateDestinationRoot),
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
                                            dummyCandidateDestinationRoot);
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

    const std::wstring dummyProbe = std::format(L"{}/dataset/bp_{:02}.bin", dummyCandidateDestinationRoot, 0);
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

    // The no-override baseline resolves to the local storage profile's preferred concurrency
    // (the engine probes the real medium since Fairstream 5D); derive it from the same probe.
    unsigned int kBaselineConcurrency = 4u;
    {
        FileSystemStorageCharacteristics storageCharacteristics{};
        storageCharacteristics.sizeBytes    = sizeof(storageCharacteristics);
        const std::wstring storageProbePath = state.tempRoot.wstring();
        if (state.fsLocal && SUCCEEDED(state.fsLocal->GetStorageCharacteristics(storageProbePath.c_str(), &storageCharacteristics)) &&
            storageCharacteristics.preferredCopyMoveConcurrency != 0u)
        {
            kBaselineConcurrency = storageCharacteristics.preferredCopyMoveConcurrency;
        }
    }
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
    unsigned int kCandidatePopupResolvedConcurrency = kBaselineConcurrency;
    {
        FileSystemStorageCharacteristics storageCharacteristics{};
        storageCharacteristics.sizeBytes = sizeof(storageCharacteristics);
        if (state.fsDummy && SUCCEEDED(state.fsDummy->GetStorageCharacteristics(resolvedCandidateFolder.c_str(), &storageCharacteristics)) &&
            storageCharacteristics.preferredCopyMoveConcurrency != 0u)
        {
            kCandidatePopupResolvedConcurrency = storageCharacteristics.preferredCopyMoveConcurrency;
        }
    }

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

        const unsigned int effectiveBudget = task->_effectiveConcurrencyBudget.load(std::memory_order_acquire);
        if (effectiveBudget == 0u)
        {
            return;
        }

        if (runStartTick == 0)
        {
            runStartTick = nowTick;
        }

        configured = (std::max)(configured, effectiveBudget);
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

        if (popupSnapshot.effectiveConcurrencyBudget == 0u)
        {
            return 0; // resolved values not published yet; keep polling
        }

        if (! popupSnapshot.autoConcurrencyUsed || popupSnapshot.autoTunedConcurrency != kCandidatePopupResolvedConcurrency ||
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

        const std::string localConfig = std::format(
            R"json({{"concurrencyMode":"manual","copyMoveMaxConcurrency":{},"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"preserve","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4}})json",
            kBaselineConcurrency);
        if (! SetPluginConfiguration(state.infoLocal.get(), localConfig))
        {
            restoreOverridePerfState();
            Fail(L"Failed to reset local plugin concurrency before @conn precedence validation.");
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

        state.markerTick    = nowTick;
        state.stepStartTick = nowTick;
        state.stepState     = 1;
        return false;
    }

    FolderWindow::FileOperationState::Task* taskA = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    FolderWindow::FileOperationState::Task* taskB = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
    const bool completedA                         = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();
    const bool completedB                         = state.taskB.has_value() && state.completedTasks.find(state.taskB.value()) != state.completedTasks.end();

    if (state.stepState == 1)
    {
        const auto hasStartedOrCompleted = [&](FolderWindow::FileOperationState::Task* task, const std::optional<std::uint64_t>& taskId) noexcept
        {
            if (task && task->HasStarted())
            {
                return true;
            }

            if (! taskId.has_value())
            {
                return false;
            }

            const auto completed = state.completedTasks.find(taskId.value());
            return completed != state.completedTasks.end() && completed->second.started;
        };

        // Fail-closed path-only interlocks can serialize these tasks. In that mode A may be
        // removed from the live task list before B starts, so completion is valid start proof.
        if (! hasStartedOrCompleted(taskA, state.taskA) || ! hasStartedOrCompleted(taskB, state.taskB))
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
                std::scoped_lock lock(taskA->_perItemInFlightCallsMutex, taskB->_perItemInFlightCallsMutex);
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
        if (state.taskB.has_value())
        {
            if (FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(state.taskB.value()))
            {
                task->RequestCancel();
            }
            Fail(L"@conn deleteMaxConcurrency incorrectly broadened an unbound provider's delete authority.");
            return true;
        }

        // Provider-direct deletion is test teardown only. The assertion above is that a connection
        // performance override may constrain an authorized operation, but can never authorize one.
        const FileSystemFlags cleanupFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        const HRESULT cleanupHr            = state.fsDummy->DeleteItem(L"/conn-selftest/copy", cleanupFlags, nullptr, nullptr, nullptr);
        if (FAILED(cleanupHr))
        {
            Fail(std::format(L"@conn override teardown failed: 0x{:08X}.", static_cast<unsigned long>(cleanupHr)));
            return true;
        }

        RemoveConnectionProfileByName(state.connOverrideProfileName);
        state.connOverrideProfileName.clear();

        NextStep(state, SelfTestState::Step::Phase12_ReparsePointPolicy);
        return false;
    }

    return false;
}

#endif
#if defined(FILEOPS_SELFTEST_INCLUDE_PHASE12)

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

    const std::filesystem::path srcDir                  = state.tempRoot / L"reparse-src";
    const std::filesystem::path dstDir                  = state.tempRoot / L"reparse-dst";
    const std::filesystem::path moveSrc                 = state.tempRoot / L"reparse-move-src";
    const std::filesystem::path moveDst                 = state.tempRoot / L"reparse-move-dst";
    const std::filesystem::path delDir                  = state.tempRoot / L"reparse-delete";
    const std::filesystem::path targetDir               = state.tempRoot / L"reparse-target";
    const std::filesystem::path targetFile              = targetDir / L"keep.bin";
    const std::filesystem::path moveInsideTarget        = moveSrc / L"inside";
    const std::filesystem::path moveInsideLink          = moveSrc / L"toInside";
    const std::filesystem::path bridgeMoveRootReparse   = state.tempRoot / L"bridge-move-root-link";
    const std::filesystem::path bridgeFollowRootReparse = state.tempRoot / L"bridge-follow-root-link";
    const std::filesystem::path bridgeCopyRootReparse   = state.tempRoot / L"bridge-copy-root-link";

    const std::wstring dummyBridgeMoveRoot       = L"/bridge-reparse-move";
    const std::wstring dummyBridgeFollowMoveRoot = L"/bridge-reparse-follow-move";
    const std::wstring dummyBridgeCopyRoot       = L"/bridge-reparse-copy";

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"preserve"})json"));

        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir) || ! RecreateEmptyDirectory(moveSrc) || ! RecreateEmptyDirectory(moveDst) ||
            ! RecreateEmptyDirectory(delDir) || ! RecreateEmptyDirectory(targetDir))
        {
            Fail(L"Failed to reset reparse test directories.");
            return true;
        }

        std::error_code ec;
        static_cast<void>(std::filesystem::remove_all(bridgeMoveRootReparse, ec));
        ec.clear();
        static_cast<void>(std::filesystem::remove_all(bridgeFollowRootReparse, ec));
        ec.clear();
        static_cast<void>(std::filesystem::remove_all(bridgeCopyRootReparse, ec));

        if (! WriteTestFile(srcDir / L"seed.bin", 128) || ! WriteTestFile(moveSrc / L"moved.bin", 96) ||
            ! WriteTestFile(moveInsideTarget / L"inside.bin", 80) || ! WriteTestFile(targetFile, 256))
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
        if (! TryCreateJunction(moveInsideLink, moveInsideTarget))
        {
            Fail(L"Failed to create the absolute in-tree junction for semantic Native qualification.");
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

        // C3: literal Preserve keeps the stored target text, so the copied loop still names the source tree.
        const std::wstring expectedLoopTarget = NormalizePathForCompare(std::filesystem::absolute(srcDir).wstring());
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
        // The prepared strategy is checked once the task has run.

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
        if (DebugGetPreparedTransferStrategyForSelfTest(state.taskB.value()) != FileOperations::OperationStrategy::Native)
        {
            Fail(L"A destination-absent Local tree with junctions inside is one Native rename.");
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

        const std::filesystem::path movedInsideTarget = moveDst / moveSrc.filename() / moveInsideTarget.filename();
        const std::filesystem::path movedInsideLink   = moveDst / moveSrc.filename() / moveInsideLink.filename();
        const auto movedInsideLinkTarget              = TryGetDirectoryReparseTargetAbsolute(movedInsideLink);
        const std::wstring expectedInsideMoveTarget   = NormalizePathForCompare(std::filesystem::absolute(moveInsideTarget).wstring());
        if (! movedInsideLinkTarget.has_value() || movedInsideLinkTarget.value() != expectedInsideMoveTarget ||
            GetFileAttributesW(movedInsideTarget.c_str()) == INVALID_FILE_ATTRIBUTES)
        {
            Fail(L"Destination-absent tree Move did not keep its absolute in-tree junction's stored target text.");
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.SameVolumeTreeMoveIsRename", L"destination-absent;absolute-in-tree-junction;native", 0u, 1u, 0u, S_OK);

        const std::filesystem::path optionPreserveSource = state.tempRoot / L"reparse-option-preserve-source";
        const std::filesystem::path optionPreserveDest   = state.tempRoot / L"reparse-option-preserve-destination";
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"skip"})json"));
        if (! TryCreateJunction(optionPreserveSource, targetDir))
        {
            Fail(L"Per-task link-policy test could not create the Preserve source junction.");
            return true;
        }
        FileSystemOptions preserveOptions{};
        preserveOptions.sizeBytes      = sizeof(FileSystemOptions);
        preserveOptions.linkPolicy     = FILESYSTEM_LINK_PRESERVE;
        const HRESULT optionPreserveHr = state.fsLocal->CopyItem(optionPreserveSource.c_str(),
                                                                 optionPreserveDest.c_str(),
                                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                                 &preserveOptions,
                                                                 nullptr,
                                                                 nullptr);
        if (FAILED(optionPreserveHr) || ! TryGetReparseTag(optionPreserveDest).has_value())
        {
            Fail(std::format(L"Per-task Preserve did not override the provider's live Skip default (hr=0x{:08X}).",
                             static_cast<unsigned long>(optionPreserveHr)));
            return true;
        }

        const std::filesystem::path optionSkipSource = state.tempRoot / L"reparse-option-skip-source";
        const std::filesystem::path optionSkipDest   = state.tempRoot / L"reparse-option-skip-destination";
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"preserve"})json"));
        if (! TryCreateJunction(optionSkipSource, targetDir))
        {
            Fail(L"Per-task link-policy test could not create the Skip source junction.");
            return true;
        }
        FileSystemOptions skipOptions{};
        skipOptions.sizeBytes      = sizeof(FileSystemOptions);
        skipOptions.linkPolicy     = FILESYSTEM_LINK_SKIP;
        const HRESULT optionSkipHr = state.fsLocal->CopyItem(
            optionSkipSource.c_str(), optionSkipDest.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), &skipOptions, nullptr, nullptr);
        if (optionSkipHr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) || GetFileAttributesW(optionSkipDest.c_str()) != INVALID_FILE_ATTRIBUTES ||
            GetFileAttributesW(optionSkipSource.c_str()) == INVALID_FILE_ATTRIBUTES)
        {
            Fail(std::format(L"Per-task Skip did not override the provider's live Preserve default (hr=0x{:08X}).", static_cast<unsigned long>(optionSkipHr)));
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

        const std::filesystem::path metadataMergeSource      = state.tempRoot / L"metadata-merge-source";
        const std::filesystem::path metadataMergeParent      = state.tempRoot / L"metadata-merge-parent";
        const std::filesystem::path metadataMergeDestination = metadataMergeParent / metadataMergeSource.filename();
        if (! RecreateEmptyDirectory(metadataMergeSource) || ! RecreateEmptyDirectory(metadataMergeDestination) ||
            ! WriteTestFile(metadataMergeSource / L"new-child.bin", 64u) || ! WriteTestFile(metadataMergeDestination / L"existing-child.bin", 64u))
        {
            Fail(L"Failed to prepare existing-directory metadata merge test.");
            return true;
        }

        FileSystemBasicInformation mergeSourceBasic{};
        mergeSourceBasic.sizeBytes = sizeof(FileSystemBasicInformation);
        FileSystemBasicInformation mergeDestinationBefore{};
        mergeDestinationBefore.sizeBytes = sizeof(FileSystemBasicInformation);
        if (FAILED(localIo->GetFileBasicInformation(metadataMergeSource.c_str(), &mergeSourceBasic)) ||
            FAILED(localIo->GetFileBasicInformation(metadataMergeDestination.c_str(), &mergeDestinationBefore)))
        {
            Fail(L"Existing-directory metadata merge test could not read initial metadata.");
            return true;
        }
        mergeSourceBasic.creationTime -= 7ll * kMetadataTickOffset;
        mergeSourceBasic.lastWriteTime -= 7ll * kMetadataTickOffset;
        if (FAILED(localIo->SetFileBasicInformation(metadataMergeSource.c_str(), &mergeSourceBasic)))
        {
            Fail(L"Existing-directory metadata merge test could not seed distinct source metadata.");
            return true;
        }

        const HRESULT metadataMergeHr = state.fsLocal->CopyItem(
            metadataMergeSource.c_str(), metadataMergeDestination.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), nullptr, nullptr, nullptr);
        if (FAILED(metadataMergeHr))
        {
            Fail(std::format(L"Existing-directory metadata merge failed: 0x{:08X}.", static_cast<unsigned long>(metadataMergeHr)));
            return true;
        }

        FileSystemBasicInformation mergeDestinationAfter{};
        mergeDestinationAfter.sizeBytes = sizeof(FileSystemBasicInformation);
        if (FAILED(localIo->GetFileBasicInformation(metadataMergeDestination.c_str(), &mergeDestinationAfter)) ||
            mergeDestinationAfter.creationTime != mergeDestinationBefore.creationTime || mergeDestinationAfter.creationTime == mergeSourceBasic.creationTime)
        {
            Fail(L"Existing-directory metadata merge replaced the destination directory's identity metadata with the source directory metadata.");
            return true;
        }

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
        const auto* bridgeMoveTask = state.fileOps ? state.fileOps->FindTask(state.taskC.value()) : nullptr;
        const std::shared_ptr<const FileOperations::FileOperationPlanGroup> bridgeMovePlans = bridgeMoveTask != nullptr ? bridgeMoveTask->LoadPlans() : nullptr;
        const auto* bridgeMovePlan =
            bridgeMovePlans && bridgeMovePlans->size() == 1u ? std::get_if<FileOperations::TransferPlan>(&bridgeMovePlans->front()) : nullptr;
        // R3-2: the Dummy destination proves published content through its writer, so the pair is
        // admitted Managed; a link carries no writer proof and still retains its source per item.
        if (bridgeMovePlan == nullptr || bridgeMovePlan->strategy != FileOperations::OperationStrategy::Managed)
        {
            Fail(L"A Local link Move to a writer-proof destination must be admitted Managed and retain its source per item.");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    // A Move never applies the Skip default. The Dummy destination cannot preserve a link, so the
    // bridge raises the typed unsupported-reparse prompt and Skip retains the source link.
    const auto answerUnsupportedLinkPrompt = [&](const std::optional<std::uint64_t>& taskId, uint32_t answeredState) noexcept -> std::optional<bool>
    {
        using Task             = FolderWindow::FileOperationState::Task;
        Task* const bridgeTask = state.fileOps && taskId.has_value() ? state.fileOps->FindTask(taskId.value()) : nullptr;
        const auto prompt      = TryGetConflictPromptCopy(bridgeTask);
        if (! prompt.has_value())
        {
            return std::nullopt;
        }
        if (prompt->bucket != Task::ConflictBucket::UnsupportedReparse || ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
        {
            return false;
        }
        bridgeTask->SubmitConflictDecision(Task::ConflictAction::Skip, false);
        state.stepState = answeredState;
        return true;
    };

    if (state.stepState == 3 || state.stepState == 8)
    {
        const auto itBridgeMove = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
        if (itBridgeMove == state.completedTasks.end())
        {
            if (state.stepState == 3 && answerUnsupportedLinkPrompt(state.taskC, 8u) == false)
            {
                Fail(L"Bridge Move of a link into a provider without link support must raise the UnsupportedReparse prompt with Skip.");
                return true;
            }
            return false;
        }

        constexpr HRESULT expectedSkip = S_FALSE;
        if (itBridgeMove->second.hr != expectedSkip)
        {
            std::wstring diagnostics;
            std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> entries;
            state.fileOps->CollectDiagnostics(entries);
            for (const auto& entry : entries)
            {
                if (entry.taskId == state.taskC.value())
                {
                    diagnostics += std::format(L" [{} 0x{:08X} {}]", entry.category, static_cast<unsigned long>(entry.status), entry.message);
                }
            }
            const auto& results    = itBridgeMove->second.sourceItemResults;
            const auto* const item = results.empty() || ! results.front().has_value() ? nullptr : &results.front().value();
            Fail(std::format(L"Bridge move reparse expected intentional-Skip S_FALSE (0x{:08X}) but got 0x{:08X} (item status=0x{:08X} completion={} "
                             L"publication={} disposition={} strategy={}).{}",
                             static_cast<unsigned long>(expectedSkip),
                             static_cast<unsigned long>(itBridgeMove->second.hr),
                             item != nullptr ? static_cast<unsigned long>(item->status) : 0xFFFFFFFFul,
                             item != nullptr ? static_cast<unsigned>(item->completion) : 255u,
                             item != nullptr ? static_cast<unsigned>(item->publication) : 255u,
                             item != nullptr ? static_cast<unsigned>(item->sourceDisposition) : 255u,
                             item != nullptr ? static_cast<unsigned>(item->strategy) : 255u,
                             diagnostics));
            return true;
        }

        std::error_code ec;
        if (! std::filesystem::exists(bridgeMoveRootReparse, ec))
        {
            Fail(L"Bridge move reparse skipped item but source link was removed.");
            return true;
        }

        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"followTargets"})json"));

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyBridgeFollowMoveRoot))
        {
            Fail(L"Failed to prepare dummy root for bridge follow-target move reparse test.");
            return true;
        }

        if (! TryCreateJunction(bridgeFollowRootReparse, targetDir))
        {
            Fail(L"Failed to create bridge follow-target move reparse source.");
            return true;
        }

        const FileSystemFlags bridgeFollowMoveFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskB                                 = StartFileOperationAndGetId(state.fileOps,
                                                                                 FILESYSTEM_MOVE,
                                                                                 FolderWindow::Pane::Left,
                                                                                 FolderWindow::Pane::Right,
                                                                                 state.fsLocal,
                                                                                 {bridgeFollowRootReparse},
                                                                                 std::filesystem::path(dummyBridgeFollowMoveRoot),
                                                                                 bridgeFollowMoveFlags,
                                                                                 false,
                                                                                 0,
                                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                                 false,
                                                                                 state.fsDummy);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start bridge follow-target move reparse task.");
            return true;
        }

        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4 || state.stepState == 9)
    {
        const auto itBridgeFollowMove = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (itBridgeFollowMove == state.completedTasks.end())
        {
            if (state.stepState == 4 && answerUnsupportedLinkPrompt(state.taskB, 9u) == false)
            {
                Fail(L"Legacy FollowTargets migration: the Move must raise the UnsupportedReparse prompt with Skip.");
                return true;
            }
            return false;
        }

        constexpr HRESULT expectedSkip = S_FALSE;
        if (itBridgeFollowMove->second.hr != expectedSkip)
        {
            Fail(std::format(L"Legacy FollowTargets migration expected intentional-Skip S_FALSE (0x{:08X}) but got 0x{:08X}.",
                             static_cast<unsigned long>(expectedSkip),
                             static_cast<unsigned long>(itBridgeFollowMove->second.hr)));
            return true;
        }

        std::error_code ec;
        if (GetFileAttributesW(bridgeFollowRootReparse.c_str()) == INVALID_FILE_ATTRIBUTES)
        {
            Fail(L"Legacy FollowTargets migration did not retain the skipped source link.");
            return true;
        }

        ec.clear();
        if (! std::filesystem::exists(targetFile, ec) || ec)
        {
            Fail(L"Legacy FollowTargets migration touched the skipped junction target file.");
            return true;
        }

        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"preserve"})json"));

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

        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
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
        state.stepState = 6;
        return false;
    }

    if (state.stepState == 6)
    {
        const auto itBridgeCopy = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (itBridgeCopy == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT expectedSkip = S_FALSE;
        if (itBridgeCopy->second.hr != expectedSkip)
        {
            Fail(std::format(L"Bridge copy unsupported reparse expected intentional-Skip S_FALSE (0x{:08X}) but got 0x{:08X}.",
                             static_cast<unsigned long>(expectedSkip),
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

        state.stepState = 7;
        return false;
    }

    if (state.stepState == 7)
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

#endif
#if defined(FILEOPS_SELFTEST_INCLUDE_PHASE13)

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

        const uint64_t warningTaskId = 0xC200000000000000ull | (static_cast<uint64_t>(GetTickCount64()) & 0x0000FFFFFFFFFFFFull);
        FolderWindow::FileOperationState::CompletedTaskSummary warningSummary{};
        warningSummary.taskId                = warningTaskId;
        warningSummary.operation             = FILESYSTEM_COPY;
        warningSummary.resultHr              = S_OK;
        warningSummary.warningCount          = 1;
        warningSummary.completedTick         = GetTickCount64();
        warningSummary.lastDiagnosticMessage = L"Selftest warning summary should require attention.";
        state.fileOps->DebugAppendCompletedTaskForSelfTest(std::move(warningSummary));

        // Enabling auto-dismiss should immediately remove already-completed success tasks.
        state.fileOps->SetAutoDismissSuccess(true);
        summaries.clear();
        state.fileOps->CollectCompletedTasks(summaries);
        bool warningRetained = false;
        for (const auto& summary : summaries)
        {
            if (summary.taskId == state.taskA.value())
            {
                Fail(L"Phase13_PostMortemDiagnostics enabling auto-dismiss did not remove the existing success task.");
                return true;
            }
            if (summary.taskId == warningTaskId)
            {
                warningRetained = true;
            }
        }

        if (! warningRetained)
        {
            Fail(L"Phase13_PostMortemDiagnostics auto-dismiss removed an S_OK task that had warnings.");
            return true;
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

        FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());
        if (! taskB)
        {
            Fail(L"Phase13_PostMortemDiagnostics could not find the admitted cancellation task.");
            return true;
        }

        // Cancel at the admission boundary. Waiting for a started/running observation is
        // racy for memory-backed providers whose atomic publication can finish in one UI tick.
        taskB->RequestCancel();
        state.stepState = 5;
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

case SelfTestState::Step::C1_ExitCloseDeferredUntilTasksQuiet:
{
    // C1: after the exit confirmation, cancellation drains off the UI thread. The window stays
    // alive and keeps stepping while a canceled worker is still finishing, and the close resumes
    // only once the last task has been reaped. The post-finished pause point holds the worker
    // after its cancel; a message-only STATIC window stands in for the main window, since the
    // default WM_CLOSE handling destroys it. On the baseline the exit path canceled and closed at
    // once.
    const std::filesystem::path root        = state.tempRoot / L"c1-exit-close-deferred";
    const std::filesystem::path source      = root / L"source.bin";
    const std::filesystem::path destination = root / L"destination";
    const auto cleanup                      = [&]() noexcept
    {
        SetFileOpsPostFinishedCompletionPauseForSelfTest(false);
        if (state.c1DeferredCloseTarget != nullptr && IsWindow(state.c1DeferredCloseTarget) != FALSE)
        {
            DestroyWindow(state.c1DeferredCloseTarget);
        }
        state.c1DeferredCloseTarget = nullptr;
        state.taskA.reset();
        static_cast<void>(SelfTest::RemoveAll(root));
    };
    if (HasTimedOut(state, GetTickCount64(), 60'000ull))
    {
        Fail(std::format(L"C1 exit close deferred timed out at step {}.", state.stepState));
        cleanup();
        return true;
    }
    if (state.stepState == 0u)
    {
        const wil::com_ptr<IFileSystem> paneFileSystem = state.folderWindow ? state.folderWindow->GetFileSystem(FolderWindow::Pane::Left) : nullptr;
        if (! state.fileOps || ! state.folderWindow || ! paneFileSystem)
        {
            Fail(L"C1 exit close deferred requires the folder window and its left pane file system.");
            return true;
        }
        if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(destination) || ! WriteFilledTestFile(source, 256u * 1024u, 0x3C))
        {
            Fail(L"C1 exit close deferred could not stage its source.");
            cleanup();
            return true;
        }
        state.c1DeferredCloseTarget = CreateWindowExW(0, L"STATIC", L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, nullptr, nullptr);
        if (state.c1DeferredCloseTarget == nullptr)
        {
            Fail(L"C1 exit close deferred could not create its message-only close target.");
            cleanup();
            return true;
        }
        SetFileOpsPostFinishedCompletionPauseForSelfTest(true);
        uint64_t taskId           = 0u;
        const HRESULT admissionHr = state.fileOps->AdmitOperation(FILESYSTEM_COPY,
                                                                  FolderWindow::Pane::Left,
                                                                  FolderWindow::Pane::Right,
                                                                  paneFileSystem,
                                                                  {source},
                                                                  destination,
                                                                  FILESYSTEM_FLAG_NONE,
                                                                  false,
                                                                  0u,
                                                                  FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                  false,
                                                                  nullptr,
                                                                  &taskId);
        if (FAILED(admissionHr) || taskId == 0u)
        {
            Fail(std::format(L"C1 exit close deferred could not admit its copy (hr=0x{:08X}).", static_cast<unsigned long>(admissionHr)));
            cleanup();
            return true;
        }
        state.taskA = taskId;
        if (state.folderWindow->CancelAllFileOperationsThenClose(state.c1DeferredCloseTarget))
        {
            Fail(L"C1 exit close deferred: with a task still owned, the exit path must cancel and defer the close, not close at once.");
            cleanup();
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        if (! HasFileOpsPostFinishedCompletionPauseEnteredForSelfTest())
        {
            return false;
        }
        state.c1HeartbeatStartTick = GetTickCount64();
        state.c1HeartbeatTicks     = 0u;
        state.stepState            = 2u;
        return false;
    }
    if (state.stepState == 2u)
    {
        // The worker is held after its cancel: the UI thread keeps stepping, the task is still
        // owned, and the close target is untouched.
        if (IsWindow(state.c1DeferredCloseTarget) == FALSE || ! state.fileOps->HasActiveOperations())
        {
            Fail(L"C1 exit close deferred: the close resumed while a canceled task was still owned.");
            cleanup();
            return true;
        }
        ++state.c1HeartbeatTicks;
        if (state.c1HeartbeatTicks < 10u || GetTickCount64() - state.c1HeartbeatStartTick < 300ull)
        {
            return false;
        }
        ReleaseFileOpsPostFinishedCompletionPauseForSelfTest();
        state.stepState = 3u;
        return false;
    }
    if (IsWindow(state.c1DeferredCloseTarget) != FALSE)
    {
        return false;
    }
    const auto done = state.completedTasks.find(state.taskA.value());
    if (done == state.completedTasks.end() || state.fileOps->HasActiveOperations() || IsWindow(state.folderWindow->GetHwnd()) == FALSE)
    {
        Fail(L"C1 exit close deferred: the close must resume only after the last task is reaped, with the folder window still alive.");
        cleanup();
        return true;
    }
    const HRESULT drainedHr = done->second.hr;
    if (drainedHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) && drainedHr != HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED) &&
        drainedHr != HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE))
    {
        Fail(std::format(L"C1 exit close deferred: the drained task must end as a cancellation (hr=0x{:08X}).", static_cast<unsigned long>(drainedHr)));
        cleanup();
        return true;
    }
    cleanup();
    NextStep(state, SelfTestState::Step::Beeline_SameVolumeTreeMoveIsRename);
    return false;
}

case SelfTestState::Step::Beeline_SameVolumeTreeMoveIsRename:
{
    // Beeline: a same-volume Move of a regular file and a directory tree is one Native rename per
    // selected item. The identity of a nested child survives the move, which no copy could reproduce.
    const std::filesystem::path root         = state.tempRoot / L"beeline-tree-move";
    const std::filesystem::path source       = root / L"source";
    const std::filesystem::path destination  = root / L"destination";
    const std::filesystem::path innerSource  = source / L"tree" / L"inner.bin";
    const std::filesystem::path innerMoved   = destination / L"tree" / L"inner.bin";
    const std::filesystem::path bulkSource   = root / L"bulk";
    const std::filesystem::path bulkCopyRoot = root / L"bulk-copy";
    const std::filesystem::path bulkCopied   = bulkCopyRoot / L"bulk";
    const std::filesystem::path bulkMoveRoot = root / L"bulk-moved";
    const std::filesystem::path bulkMoved    = bulkMoveRoot / L"bulk";
    constexpr unsigned int kBulkFiles        = 2'000u;
    constexpr size_t kBulkFileBytes          = 8u * 1024u;
    const auto fileIdentity                  = [](const std::filesystem::path& path) noexcept -> std::optional<std::pair<DWORD, ULONGLONG>>
    {
        const wil::unique_hfile handle(CreateFileW(path.c_str(),
                                                   FILE_READ_ATTRIBUTES,
                                                   FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                   nullptr,
                                                   OPEN_EXISTING,
                                                   FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                                                   nullptr));
        BY_HANDLE_FILE_INFORMATION information{};
        if (! handle || GetFileInformationByHandle(handle.get(), &information) == FALSE)
        {
            return std::nullopt;
        }
        return std::pair<DWORD, ULONGLONG>{information.dwVolumeSerialNumber,
                                           (static_cast<ULONGLONG>(information.nFileIndexHigh) << 32u) | information.nFileIndexLow};
    };
    if (HasTimedOut(state, GetTickCount64(), 180'000ull))
    {
        Fail(std::format(L"Beeline tree move timed out at step {}.", state.stepState));
        state.taskA.reset();
        return true;
    }
    if (state.stepState == 0u)
    {
        if (! state.fileOps || ! state.fsLocal)
        {
            Fail(L"Beeline tree move requires the local file system.");
            return true;
        }
        if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(source / L"tree") || ! RecreateEmptyDirectory(destination) ||
            ! WriteFilledTestFile(source / L"file.bin", 64u * 1024u, 0x11) || ! WriteFilledTestFile(innerSource, 4u * 1024u, 0x22))
        {
            Fail(L"Beeline tree move could not stage its tree.");
            return true;
        }
        const auto identityBefore = fileIdentity(innerSource);
        if (! identityBefore.has_value())
        {
            Fail(L"Beeline tree move could not read the nested child's identity.");
            return true;
        }
        state.beelineVolumeSerial = identityBefore->first;
        state.beelineFileIndex    = identityBefore->second;
        state.taskA               = StartFileOperationAndGetId(state.fileOps,
                                                               FILESYSTEM_MOVE,
                                                               FolderWindow::Pane::Left,
                                                               FolderWindow::Pane::Right,
                                                               state.fsLocal,
                                                               {source / L"file.bin", source / L"tree"},
                                                               destination,
                                                               static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                               false,
                                                               0,
                                                               FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                               false,
                                                               nullptr);
        if (! state.taskA.has_value())
        {
            Fail(L"Beeline tree move could not start its move.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    const auto done = state.completedTasks.find(state.taskA.value());
    if (done == state.completedTasks.end())
    {
        return false;
    }
    std::error_code ec;
    if (state.stepState == 1u)
    {
        const auto identityAfter = fileIdentity(innerMoved);
        if (FAILED(done->second.hr) || ! std::filesystem::exists(destination / L"file.bin", ec) || std::filesystem::exists(source / L"file.bin", ec) ||
            std::filesystem::exists(source / L"tree", ec) || ! identityAfter.has_value() || identityAfter->first != state.beelineVolumeSerial ||
            identityAfter->second != state.beelineFileIndex ||
            DebugGetPreparedTransferStrategyForSelfTest(state.taskA.value()) != FileOperations::OperationStrategy::Native)
        {
            Fail(std::format(L"Beeline tree move: the file and the tree must relocate by Native rename with identity preserved (hr=0x{:08X}).",
                             static_cast<unsigned long>(done->second.hr)));
            state.taskA.reset();
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.SameVolumeTreeMoveIsRename", L"shape=file+tree;identity-preserved", 0u, 2u, 0u, done->second.hr);

        // Throughput: a same-volume Copy of an identical tree is the byte cost the former Managed
        // Move paid before its exact delete; the rename Move of that copy pays none of it.
        if (! RecreateEmptyDirectory(bulkSource) || ! RecreateEmptyDirectory(bulkCopyRoot) || ! RecreateEmptyDirectory(bulkMoveRoot))
        {
            Fail(L"Beeline tree move could not stage its throughput tree.");
            return true;
        }
        for (unsigned int index = 0u; index < kBulkFiles; ++index)
        {
            if (! WriteFilledTestFile(bulkSource / std::format(L"file{:04}.bin", index), kBulkFileBytes, static_cast<unsigned char>(index)))
            {
                Fail(L"Beeline tree move could not stage its throughput files.");
                return true;
            }
        }
        state.markerTick = GetTickCount64();
        state.taskA      = StartFileOperationAndGetId(state.fileOps,
                                                      FILESYSTEM_COPY,
                                                      FolderWindow::Pane::Left,
                                                      FolderWindow::Pane::Right,
                                                      state.fsLocal,
                                                      {bulkSource},
                                                      bulkCopyRoot,
                                                      static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                      false,
                                                      0,
                                                      FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Beeline tree move could not start its baseline copy.");
            return true;
        }
        state.stepState = 2u;
        return false;
    }
    if (state.stepState == 2u)
    {
        if (FAILED(done->second.hr) || ! FileSizeEquals(bulkCopied / std::format(L"file{:04}.bin", kBulkFiles - 1u), kBulkFileBytes))
        {
            Fail(
                std::format(L"Beeline tree move: the baseline copy of the throughput tree failed (hr=0x{:08X}).", static_cast<unsigned long>(done->second.hr)));
            state.taskA.reset();
            return true;
        }
        state.beelineCopyMs = done->second.completionTick - state.markerTick;
        state.markerTick    = GetTickCount64();
        state.taskA         = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_MOVE,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         {bulkCopied},
                                                         bulkMoveRoot,
                                                         static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Beeline tree move could not start its throughput move.");
            return true;
        }
        state.stepState = 3u;
        return false;
    }
    const uint64_t moveMs = done->second.completionTick - state.markerTick;
    if (FAILED(done->second.hr) || std::filesystem::exists(bulkCopied, ec) ||
        ! FileSizeEquals(bulkMoved / std::format(L"file{:04}.bin", kBulkFiles - 1u), kBulkFileBytes) ||
        DebugGetPreparedTransferStrategyForSelfTest(state.taskA.value()) != FileOperations::OperationStrategy::Native || moveMs > state.beelineCopyMs)
    {
        Fail(std::format(
            L"Beeline tree move: the rename Move of the throughput tree must be Native and no slower than its copy (hr=0x{:08X}, copyMs={}, moveMs={}).",
            static_cast<unsigned long>(done->second.hr),
            state.beelineCopyMs,
            moveMs));
        state.taskA.reset();
        return true;
    }
    Debug::Perf::Emit(
        L"FileOps.SelfTest.SameVolumeTreeMoveThroughput",
        std::format(L"files={};bytes={};copyMs={};moveMs={}", kBulkFiles, static_cast<uint64_t>(kBulkFiles) * kBulkFileBytes, state.beelineCopyMs, moveMs),
        moveMs,
        state.beelineCopyMs,
        static_cast<uint64_t>(kBulkFiles),
        done->second.hr);
    static_cast<void>(SelfTest::RemoveAll(root));
    state.taskA.reset();
    NextStep(state, SelfTestState::Step::Beeline_LoopbackShareMoveIsRename);
    return false;
}
case SelfTestState::Step::Beeline_LoopbackShareMoveIsRename:
{
    // Beeline: the SMB profile takes the same route. A Move whose source and destination both live
    // on the loopback administrative-share alias of the sandbox (local-win32-smb) is one Native
    // rename per selected item, never a bridge copy. Environment-gated on the alias being reachable.
    if (HasTimedOut(state, GetTickCount64(), 120'000ull))
    {
        Fail(L"Beeline loopback share move timed out.");
        state.taskA.reset();
        return true;
    }
    const std::filesystem::path root        = state.tempRoot / L"beeline-share-move";
    const std::filesystem::path source      = root / L"source";
    const std::filesystem::path destination = root / L"destination";
    if (state.stepState == 0u)
    {
        if (! LoopbackShareReachable(state.tempRoot))
        {
            AppendLog(L"Beeline loopback share move: the administrative-share alias is not reachable; environment-gated case skipped.");
            Debug::Perf::Emit(L"FileOps.SelfTest.LoopbackShareMoveIsRename", L"skip=share-unreachable", 0u, 0u, 0u, S_FALSE);
            NextStep(state, SelfTestState::Step::C1_PermanentDeleteConfirmsOnCard);
            return false;
        }
        if (! RecreateEmptyDirectory(source / L"tree") || ! RecreateEmptyDirectory(destination) ||
            ! WriteFilledTestFile(source / L"file.bin", 64u * 1024u, 0x33) || ! WriteFilledTestFile(source / L"tree" / L"inner.bin", 4u * 1024u, 0x44))
        {
            Fail(L"Beeline loopback share move could not stage its tree.");
            return true;
        }
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 {LoopbackShareAlias(source / L"file.bin"), LoopbackShareAlias(source / L"tree")},
                                                 LoopbackShareAlias(destination),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 nullptr);
        if (! state.taskA.has_value())
        {
            Fail(L"Beeline loopback share move could not start its move.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    const auto done = state.completedTasks.find(state.taskA.value());
    if (done == state.completedTasks.end())
    {
        return false;
    }
    std::error_code ec;
    if (FAILED(done->second.hr) || ! std::filesystem::exists(destination / L"file.bin", ec) ||
        ! std::filesystem::exists(destination / L"tree" / L"inner.bin", ec) || std::filesystem::exists(source / L"file.bin", ec) ||
        std::filesystem::exists(source / L"tree", ec) ||
        DebugGetPreparedTransferStrategyForSelfTest(state.taskA.value()) != FileOperations::OperationStrategy::Native)
    {
        Fail(std::format(L"Beeline loopback share move: the file and the tree must relocate by Native rename on the SMB profile (hr=0x{:08X}).",
                         static_cast<unsigned long>(done->second.hr)));
        state.taskA.reset();
        return true;
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.LoopbackShareMoveIsRename", L"shape=file+tree;profile=local-win32-smb", 0u, 2u, 0u, done->second.hr);
    static_cast<void>(SelfTest::RemoveAll(root));
    state.taskA.reset();
    NextStep(state, SelfTestState::Step::C1_PermanentDeleteConfirmsOnCard);
    return false;
}

case SelfTestState::Step::C1_PermanentDeleteConfirmsOnCard:
{
    // C1: the permanent-delete confirmation is a card consent after Preparing has pinned the
    // selected roots; no modal runs at ingress. With the host prompt override set to Cancel the
    // published task ends canceled with the file intact (on the baseline the modal refused and no
    // task existed). Without an override the card shows the consent with Cancel and Permanently
    // delete, and answering Permanently delete removes the file.
    using Task                         = FolderWindow::FileOperationState::Task;
    const std::filesystem::path root   = state.tempRoot / L"c1-permanent-delete-card";
    const std::filesystem::path victim = root / L"victim.bin";
    if (HasTimedOut(state, GetTickCount64(), 60'000ull))
    {
        HostClearTestPromptResultOverride();
        Fail(std::format(L"C1 permanent delete card timed out at step {}.", state.stepState));
        state.taskA.reset();
        return true;
    }
    if (state.stepState == 0u)
    {
        if (! state.fileOps || ! state.fsLocal)
        {
            Fail(L"C1 permanent delete card requires the local file system.");
            return true;
        }
        if (! RecreateEmptyDirectory(root) || ! WriteFilledTestFile(victim, 8u * 1024u, 0x7E))
        {
            Fail(L"C1 permanent delete card could not stage its file.");
            return true;
        }
        HostResetTestPromptRequestCount();
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {victim},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        HostClearTestPromptResultOverride();
        if (! state.taskA.has_value())
        {
            Fail(L"C1 permanent delete card: admission must publish the task and leave the confirmation to the card, not refuse at a modal.");
            return true;
        }
        state.stepState = 1u;
        return false;
    }
    if (state.stepState == 1u)
    {
        const auto done = state.completedTasks.find(state.taskA.value());
        if (done == state.completedTasks.end())
        {
            return false;
        }
        std::error_code ec;
        if (done->second.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || ! std::filesystem::exists(victim, ec) || HostGetTestPromptRequestCount() != 0u)
        {
            Fail(std::format(L"C1 permanent delete card: a cancelled confirmation must end the task as canceled with the file intact and no modal "
                             L"(hr=0x{:08X}, prompts={}).",
                             static_cast<unsigned long>(done->second.hr),
                             HostGetTestPromptRequestCount()));
            state.taskA.reset();
            return true;
        }
        // The run auto-accepts host prompts; this admission must reach the card unanswered, so the
        // default is off only for the moment admission captures it.
        HostSetAutoAcceptPrompts(false);
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {victim},
                                                 {},
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false);
        HostSetAutoAcceptPrompts(true);
        if (! state.taskA.has_value())
        {
            Fail(L"C1 permanent delete card: the second admission must publish its task.");
            return true;
        }
        state.stepState = 2u;
        return false;
    }
    if (state.stepState == 2u)
    {
        Task* const task  = state.fileOps->FindTask(state.taskA.value());
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            if (state.completedTasks.contains(state.taskA.value()))
            {
                Fail(L"C1 permanent delete card: the task ran without asking on the card.");
                state.taskA.reset();
                return true;
            }
            return false;
        }
        if (prompt->bucket != Task::ConflictBucket::PermanentDeleteConfirmation || ! PromptHasAction(prompt.value(), Task::ConflictAction::PermanentDelete) ||
            ! PromptHasAction(prompt.value(), Task::ConflictAction::Cancel) || prompt->consentDetail.empty())
        {
            Fail(
                std::format(L"C1 permanent delete card: the card must ask with Permanently delete and Cancel and say what is deleted (bucket={}, detail='{}').",
                            static_cast<unsigned int>(prompt->bucket),
                            prompt->consentDetail));
            task->SubmitConflictDecision(Task::ConflictAction::Cancel, false);
            state.taskA.reset();
            return true;
        }
        task->SubmitConflictDecision(Task::ConflictAction::PermanentDelete, false);
        state.stepState = 3u;
        return false;
    }
    const auto done = state.completedTasks.find(state.taskA.value());
    if (done == state.completedTasks.end())
    {
        return false;
    }
    std::error_code ec;
    if (FAILED(done->second.hr) || std::filesystem::exists(victim, ec) || HostGetTestPromptRequestCount() != 0u)
    {
        Fail(std::format(L"C1 permanent delete card: Permanently delete must remove the file with no modal (hr=0x{:08X}, prompts={}).",
                         static_cast<unsigned long>(done->second.hr),
                         HostGetTestPromptRequestCount()));
        state.taskA.reset();
        return true;
    }
    static_cast<void>(SelfTest::RemoveAll(root));
    state.taskA.reset();
    NextStep(state, SelfTestState::Step::Phase11_CrossFileSystemBridge);
    return false;
}

#endif
