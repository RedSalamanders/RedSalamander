case SelfTestState::Step::FileOps_CopyMergeIntoExistingFolder:
{
    using Task              = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"FileOps_CopyMergeIntoExistingFolder timed out.");
        return true;
    }

    const std::filesystem::path srcRoot      = state.tempRoot / L"clearflow-copy-merge-src";
    const std::filesystem::path dstRoot      = state.tempRoot / L"clearflow-copy-merge-dst";
    const std::filesystem::path srcFoo       = srcRoot / L"Foo";
    const std::filesystem::path dstFoo       = dstRoot / L"Foo";
    const std::filesystem::path srcConflict  = srcFoo / L"a.bin";
    const std::filesystem::path dstConflict  = dstFoo / L"a.bin";
    const std::filesystem::path srcNested    = srcFoo / L"nested" / L"c.bin";
    const std::filesystem::path dstNested    = dstFoo / L"nested" / L"c.bin";
    const std::filesystem::path dstKeep      = dstFoo / L"keep.bin";

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

        if (prompt->bucket != Task::ConflictBucket::Exists || ! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
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
            Fail(std::format(L"Copy-merge prompted for the wrong item. expected='{}' -> '{}' actual='{}' -> '{}'.",
                             expectedSource,
                             expectedDest,
                             actualSource,
                             actualDest));
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
case SelfTestState::Step::FileOps_MoveMergeIntoExistingFolderSameVolume:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
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

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot))
        {
            Fail(L"Failed to reset move-merge directories.");
            return true;
        }

        if (! WriteTestFile(srcFile, 4 * 1024) || ! WriteTestFile(srcNested, 6 * 1024) || ! WriteTestFile(dstKeep, 3 * 1024))
        {
            Fail(L"Failed to seed move-merge test tree.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
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

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (auto* task = state.fileOps && state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr)
        {
            if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
            {
                Fail(std::format(L"Same-volume move-merge unexpectedly prompted for '{}' -> '{}'.", prompt->sourcePath, prompt->destinationPath));
                return true;
            }
        }

        const auto it = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Same-volume move-merge task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        std::error_code ec;
        if (std::filesystem::exists(srcFoo, ec))
        {
            Fail(L"Same-volume move-merge left the source directory behind.");
            return true;
        }

        if (it->second.conflictPromptCount != 0u || ! FileSizeEquals(dstFile, 4 * 1024) || ! FileSizeEquals(dstNested, 6 * 1024) ||
            ! FileSizeEquals(dstKeep, 3 * 1024))
        {
            Fail(std::format(L"Same-volume move-merge integrity failure (promptCount={}).", it->second.conflictPromptCount));
            return true;
        }

        Debug::Perf::Emit(L"FileOps.SelfTest.MoveMergeIntoExistingFolderSameVolume.PromptCount", L"", 0u, it->second.conflictPromptCount, 0u, S_OK);
        NextStep(state, SelfTestState::Step::FileOps_ReparseDirectoryMergeIntoExistingFolder);
        return false;
    }

    return false;
}
case SelfTestState::Step::FileOps_ReparseDirectoryMergeIntoExistingFolder:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        Fail(L"FileOps_ReparseDirectoryMergeIntoExistingFolder timed out.");
        return true;
    }

    const std::filesystem::path srcRoot       = state.tempRoot / L"clearflow-reparse-merge-src";
    const std::filesystem::path dstRoot       = state.tempRoot / L"clearflow-reparse-merge-dst";
    const std::filesystem::path targetRoot    = state.tempRoot / L"clearflow-reparse-merge-target";
    const std::filesystem::path targetFile    = targetRoot / L"payload.bin";
    const std::filesystem::path sourceLink    = srcRoot / L"linkToTarget";
    const std::filesystem::path destinationLink = dstRoot / L"linkToTarget";

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"copyReparse"})json"));

        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(targetRoot) || ! WriteTestFile(targetFile, 512))
        {
            Fail(L"Failed to reset reparse merge directories.");
            return true;
        }

        if (! TryCreateJunction(sourceLink, targetRoot))
        {
            Fail(L"Failed to create source junction for reparse merge test.");
            return true;
        }

        std::error_code ec;
        std::filesystem::create_directories(destinationLink, ec);
        if (ec)
        {
            Fail(L"Failed to create existing destination directory for reparse merge test.");
            return true;
        }

        const HRESULT noGrantCopyHr = state.fsLocal->CopyItem(
            sourceLink.c_str(), destinationLink.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE), nullptr, nullptr, nullptr);
        if (noGrantCopyHr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
        {
            Fail(std::format(L"Reparse merge copy without overwrite expected ERROR_ALREADY_EXISTS, got 0x{:08X}.",
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

        const HRESULT overwriteCopyHr = state.fsLocal->CopyItem(sourceLink.c_str(),
                                                               destinationLink.c_str(),
                                                               static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                               nullptr,
                                                               nullptr,
                                                               nullptr);
        if (FAILED(overwriteCopyHr))
        {
            Fail(std::format(L"Reparse merge copy with overwrite failed: 0x{:08X}.", static_cast<unsigned long>(overwriteCopyHr)));
            return true;
        }

        const auto tag = TryGetReparseTag(destinationLink);
        if (! tag.has_value() || tag.value() != IO_REPARSE_TAG_MOUNT_POINT)
        {
            Fail(L"Reparse merge destination did not become a directory junction.");
            return true;
        }

        const auto copiedTarget = TryGetDirectoryReparseTargetAbsolute(destinationLink);
        if (! copiedTarget.has_value())
        {
            Fail(L"Reparse merge destination target could not be read.");
            return true;
        }

        const std::wstring expectedTarget = NormalizePathForCompare(std::filesystem::absolute(targetRoot).wstring());
        if (copiedTarget.value() != expectedTarget)
        {
            Fail(std::format(L"Reparse merge target mismatch. expected='{}' actual='{}'.", expectedTarget, copiedTarget.value()));
            return true;
        }

        NextStep(state, SelfTestState::Step::FileOps_ProviderCapabilityMatrix);
        return false;
    }

    return false;
}
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
        const auto requireCapabilities = [&](IFileSystem* fs, std::wstring_view providerName, ProviderCapabilitySnapshot& snapshot) noexcept -> bool
        {
            std::wstring reason;
            if (! TryReadProviderCapabilities(fs, snapshot, reason))
            {
                Fail(std::format(L"{} capabilities failed: {}.", providerName, reason));
                return false;
            }
            return true;
        };

        ProviderCapabilitySnapshot localCaps{};
        ProviderCapabilitySnapshot dummyCaps{};
        ProviderCapabilitySnapshot sevenZipCaps{};
        if (! requireCapabilities(state.fsLocal.get(), L"Local FileSystem", localCaps) ||
            ! requireCapabilities(state.fsDummy.get(), L"FileSystemDummy", dummyCaps) ||
            ! requireCapabilities(state.fs7z.get(), L"FileSystem7z", sevenZipCaps))
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
                Fail(std::format(L"{} componentComparison mismatch: expected {} actual {}.",
                                 providerName,
                                 expectedComparison,
                                 caps.componentComparison));
                return false;
            }
            if (caps.preferredSeparator != expectedPreferredSeparator)
            {
                Fail(std::format(L"{} preferredSeparator mismatch.", providerName));
                return false;
            }
            if (caps.acceptedSeparators != expectedAcceptedSeparators)
            {
                Fail(std::format(L"{} acceptedSeparators mismatch: expected {} actual {}.",
                                 providerName,
                                 expectedAcceptedSeparators,
                                 caps.acceptedSeparators));
                return false;
            }
            if (caps.caseOnlyRename != expectedCaseOnlyRename)
            {
                Fail(std::format(L"{} caseOnlyRename mismatch: expected {} actual {}.",
                                 providerName,
                                 expectedCaseOnlyRename,
                                 caps.caseOnlyRename));
                return false;
            }
            return true;
        };

        if (! require(localCaps.copyOperation && localCaps.moveOperation && localCaps.deleteOperation && localCaps.read && localCaps.write,
                      L"Local FileSystem should advertise copy/move/delete/read/write support.") ||
            ! require(localCaps.copyMoveMax >= 1u && localCaps.deleteMax >= 1u && localCaps.deleteRecycleMax >= 1u,
                      L"Local FileSystem should advertise positive concurrency limits.") ||
            ! require(localCaps.exportCopyWildcard && localCaps.exportMoveWildcard && localCaps.importCopyWildcard && localCaps.importMoveWildcard,
                      L"Local FileSystem should advertise wildcard cross-filesystem copy/move import/export.") ||
            ! requirePathIdentity(localCaps, L"Local FileSystem", true, L"ordinalIgnoreCase", L'\\', L"\\/", L"supported"))
        {
            return true;
        }

        if (! require(dummyCaps.copyOperation && dummyCaps.moveOperation && dummyCaps.deleteOperation && dummyCaps.read && dummyCaps.write,
                      L"FileSystemDummy should advertise copy/move/delete/read/write support.") ||
            ! require(dummyCaps.copyMoveMax == 4u && dummyCaps.deleteMax == 8u && dummyCaps.deleteRecycleMax == 2u,
                      L"FileSystemDummy should advertise the expected deterministic concurrency matrix.") ||
            ! require(dummyCaps.exportCopyWildcard && dummyCaps.exportMoveWildcard && dummyCaps.importCopyWildcard && dummyCaps.importMoveWildcard,
                      L"FileSystemDummy should advertise wildcard cross-filesystem copy/move import/export.") ||
            ! requirePathIdentity(dummyCaps, L"FileSystemDummy", true, L"ordinalIgnoreCase", L'\\', L"\\/", L"supported"))
        {
            return true;
        }

        if (! require(! sevenZipCaps.copyOperation && ! sevenZipCaps.moveOperation && ! sevenZipCaps.deleteOperation && sevenZipCaps.properties &&
                          sevenZipCaps.read && ! sevenZipCaps.write,
                      L"FileSystem7z should advertise read/properties only for same-provider operations.") ||
            ! require(sevenZipCaps.copyMoveMax == 1u && sevenZipCaps.deleteMax == 1u && sevenZipCaps.deleteRecycleMax == 1u,
                      L"FileSystem7z should advertise single-stream concurrency limits.") ||
            ! require(sevenZipCaps.exportCopyWildcard && ! sevenZipCaps.exportMoveWildcard && ! sevenZipCaps.importCopyWildcard &&
                          ! sevenZipCaps.importMoveWildcard,
                      L"FileSystem7z should advertise export-copy only for cross-filesystem transfers.") ||
            ! requirePathIdentity(sevenZipCaps, L"FileSystem7z", true, L"ordinalCaseSensitive", L'/', L"/", L"notApplicable"))
        {
            return true;
        }

        constexpr HRESULT kUnsupported = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        const wchar_t* sevenZipSources[] = {L"/missing.txt"};
        const auto requireUnsupported = [&](HRESULT hr, std::wstring_view label) noexcept -> bool
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

        if (! state.fileOps)
        {
            Fail(L"Provider capability matrix test lost FileOperationState.");
            return true;
        }

        std::vector<FolderWindow::FileOperationState::Task*> tasksBefore;
        state.fileOps->CollectTasks(tasksBefore);
        const HRESULT rejectedStart = state.fileOps->StartOperation(FILESYSTEM_COPY,
                                                                    FolderWindow::Pane::Left,
                                                                    std::nullopt,
                                                                    state.fs7z,
                                                                    {std::filesystem::path(L"/missing.txt")},
                                                                    std::filesystem::path(L"/dest"),
                                                                    FILESYSTEM_FLAG_NONE,
                                                                    false,
                                                                    0,
                                                                    FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                    false);
        std::vector<FolderWindow::FileOperationState::Task*> tasksAfter;
        state.fileOps->CollectTasks(tasksAfter);
        if (rejectedStart != kUnsupported || tasksAfter.size() != tasksBefore.size())
        {
            Fail(std::format(L"Host StartOperation should reject 7z copy before task creation (hr=0x{:08X} before={} after={}).",
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
            if (requireMoveMerge && (! caps.moveOperation || ! caps.deleteOperation))
            {
                Fail(std::format(L"{} should advertise move + delete so directory merge is attempted, not rejected.", providerName));
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
        if (! conformCloudProvider(kPluginIdS3,
                                   L"FileSystemS3 (S3)",
                                   &kExpectCopyYes,
                                   /*requireMoveMerge=*/true,
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
                                   /*requireRead=*/true,
                                   /*expectedStable=*/true,
                                   L"ordinalCaseSensitive",
                                   L"notApplicable"))
        {
            return true;
        }
        // 2B: MS Drive advertises move (directory-onto-directory merge) but NOT same-provider copy —
        // the host bridges copy. This is the honest capability declaration 2B depends on.
        if (! conformCloudProvider(kPluginIdOneDrivePersonal,
                                   L"FileSystemMicrosoftDrive (OneDrive)",
                                   &kExpectCopyNo,
                                   /*requireMoveMerge=*/true,
                                   /*requireRead=*/true,
                                   /*expectedStable=*/true,
                                   L"ordinalIgnoreCase",
                                   L"supported"))
        {
            return true;
        }
        if (! conformCloudProvider(kPluginIdGoogleDrive,
                                   L"FileSystemGoogleDrive",
                                   &kExpectCopyNo,
                                   /*requireMoveMerge=*/false,
                                   /*requireRead=*/false,
                                   /*expectedStable=*/false,
                                   L"ordinalCaseSensitive",
                                   L"notApplicable"))
        {
            return true;
        }
        // Curl/SFTP: general matrix coverage (loads + parseable read/concurrency declarations); not a 2A/2B provider.
        if (! conformCloudProvider(kPluginIdSftp,
                                   L"FileSystemCurl (SFTP)",
                                   nullptr,
                                   /*requireMoveMerge=*/false,
                                   /*requireRead=*/true,
                                   /*expectedStable=*/true,
                                   L"ordinalCaseSensitive",
                                   L"supported"))
        {
            return true;
        }
        if (! conformCloudProvider(kPluginIdImap,
                                   L"FileSystemCurl (IMAP)",
                                   &kExpectCopyNo,
                                   /*requireMoveMerge=*/false,
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
        const auto cleanup     = wil::scope_exit([&]() noexcept
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
        {
            return EnsureDummyFolderExists(state.fsDummy.get(), path.generic_wstring());
        };

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
        const std::wstring copySrcFooText = copySrcFoo.generic_wstring();
        const std::wstring copyDstRootText = copyDstRoot.generic_wstring();
        const wchar_t* copySources[]      = {copySrcFooText.c_str()};
        const HRESULT copyHr              = state.fsDummy->CopyItems(copySources,
                                                        1,
                                                        copyDstRootText.c_str(),
                                                        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                        nullptr,
                                                        &copyProgress,
                                                        nullptr);
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
        const std::wstring moveSrcFooText = moveSrcFoo.generic_wstring();
        const std::wstring moveDstRootText = moveDstRoot.generic_wstring();
        const wchar_t* moveSources[]      = {moveSrcFooText.c_str()};
        const HRESULT moveHr              = state.fsDummy->MoveItems(moveSources,
                                                        1,
                                                        moveDstRootText.c_str(),
                                                        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE),
                                                        nullptr,
                                                        &moveProgress,
                                                        nullptr);
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
        NextStep(state, SelfTestState::Step::Phase5_PreCalcSettingsApplied);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase5_PreCalcSettingsApplied:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        Fail(L"Phase5_PreCalcSettingsApplied timed out.");
        return true;
    }

    const std::filesystem::path preCalcSettingsDisabled = state.tempRoot / L"precalc-settings-disabled";
    const std::filesystem::path preCalcSettingsWorkerA  = state.tempRoot / L"precalc-settings-worker-a";
    const std::filesystem::path preCalcSettingsWorkerB  = state.tempRoot / L"precalc-settings-worker-b";
    const std::filesystem::path preCalcFanOutBudget1    = state.tempRoot / L"precalc-fanout-budget1";
    const std::filesystem::path preCalcFanOutBudget4    = state.tempRoot / L"precalc-fanout-budget4";
    const auto buildPerfSources                         = [&](std::wstring_view prefix) noexcept
    {
        std::vector<std::filesystem::path> paths;
        paths.reserve(8);
        for (int index = 0; index < 8; ++index)
        {
            paths.push_back(state.tempRoot / std::format(L"{}-{}", prefix, index));
        }
        return paths;
    };
    const auto allPathsDeleted = [](const std::vector<std::filesystem::path>& paths) noexcept
    {
        std::error_code ec;
        for (const auto& path : paths)
        {
            ec.clear();
            if (std::filesystem::exists(path, ec))
            {
                return false;
            }
        }
        return true;
    };
    const auto pathDeleted = [](const std::filesystem::path& path) noexcept
    {
        std::error_code ec;
        return ! std::filesystem::exists(path, ec);
    };
    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                               FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    if (state.stepState == 0)
    {
        if (! g_settings.fileOperations.has_value())
        {
            g_settings.fileOperations.emplace();
        }

        g_settings.fileOperations->preCalcEnabled    = false;
        g_settings.fileOperations->preCalcMaxWorkers = 4u;
        state.taskA                                  = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {preCalcSettingsDisabled}, {}, flags, false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start delete task with pre-calc disabled from settings.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        const auto completionIt = state.completedTasks.find(state.taskA.value());
        if (completionIt == state.completedTasks.end())
        {
            return false;
        }

        const CompletedTaskInfo& completion = completionIt->second;
        if (FAILED(completion.hr))
        {
            Fail(std::format(L"Unexpected hr for disabled pre-calc delete: 0x{:08X}", static_cast<unsigned long>(completion.hr)));
            return true;
        }
        if (completion.preCalcCompleted || completion.preCalcTotalBytes != 0 || completion.preCalcWorkerCountUsed != 0u)
        {
            Fail(std::format(L"Disabled pre-calc still ran (completed={} bytes={} workers={}).",
                             completion.preCalcCompleted ? 1 : 0,
                             completion.preCalcTotalBytes,
                             completion.preCalcWorkerCountUsed));
            return true;
        }

        std::error_code existsEc;
        if (std::filesystem::exists(preCalcSettingsDisabled, existsEc))
        {
            Fail(L"Disabled pre-calc delete did not remove the source tree.");
            return true;
        }

        g_settings.fileOperations->preCalcEnabled    = true;
        g_settings.fileOperations->preCalcMaxWorkers = 1u;
        state.taskB                                  = StartFileOperationAndGetId(state.fileOps,
                                                                                  FILESYSTEM_DELETE,
                                                                                  FolderWindow::Pane::Left,
                                                                                  std::nullopt,
                                                                                  state.fsLocal,
                                                                                  {preCalcSettingsWorkerA, preCalcSettingsWorkerB},
                                                                                  {},
                                                                                  flags,
                                                                                  false);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start delete task with pre-calc worker limit from settings.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto completionIt = state.completedTasks.find(state.taskB.value());
        if (completionIt == state.completedTasks.end())
        {
            return false;
        }

        const CompletedTaskInfo& completion = completionIt->second;
        if (FAILED(completion.hr))
        {
            Fail(std::format(L"Unexpected hr for worker-limited pre-calc delete: 0x{:08X}", static_cast<unsigned long>(completion.hr)));
            return true;
        }
        if (! completion.preCalcCompleted || completion.preCalcTotalBytes == 0)
        {
            Fail(std::format(L"Expected pre-calc to run for worker-limited delete (completed={} bytes={}).",
                             completion.preCalcCompleted ? 1 : 0,
                             completion.preCalcTotalBytes));
            return true;
        }
        if (completion.preCalcWorkerCountUsed != 1u)
        {
            Fail(std::format(L"Expected pre-calc worker limit of 1 to apply, but task used {} workers.", completion.preCalcWorkerCountUsed));
            return true;
        }

        std::error_code workerExistsEc;
        if (std::filesystem::exists(preCalcSettingsWorkerA, workerExistsEc) || std::filesystem::exists(preCalcSettingsWorkerB, workerExistsEc))
        {
            Fail(L"Worker-limited pre-calc delete did not remove both source trees.");
            return true;
        }

        g_settings.fileOperations->preCalcEnabled    = true;
        g_settings.fileOperations->preCalcMaxWorkers = 4u;
        state.taskC                                  = StartFileOperationAndGetId(state.fileOps,
                                                                                  FILESYSTEM_DELETE,
                                                                                  FolderWindow::Pane::Left,
                                                                                  std::nullopt,
                                                                                  state.fsLocal,
                                                                                  buildPerfSources(L"precalc-settings-perf4"),
                                                                                  {},
                                                                                  flags,
                                                                                  false);
        if (! state.taskC.has_value())
        {
            Fail(L"Failed to start pre-calc perf delete task for worker budget 4.");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        const auto completionIt = state.completedTasks.find(state.taskC.value());
        if (completionIt == state.completedTasks.end())
        {
            return false;
        }

        const CompletedTaskInfo& completion = completionIt->second;
        if (FAILED(completion.hr))
        {
            Fail(std::format(L"Unexpected hr for worker-budget-4 pre-calc delete: 0x{:08X}", static_cast<unsigned long>(completion.hr)));
            return true;
        }
        if (! completion.preCalcCompleted || completion.preCalcWorkerCountUsed != 4u)
        {
            Fail(std::format(L"Expected worker-budget-4 task to use 4 workers (completed={} workers={}).",
                             completion.preCalcCompleted ? 1 : 0,
                             completion.preCalcWorkerCountUsed));
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.ClearflowPreCalcMultiRootWorkers",
                          L"shape=8-root-delete workerBudget=4",
                          completion.preCalcDurationUs,
                          completion.preCalcWorkerCountUsed,
                          completion.preCalcTotalBytes,
                          completion.hr);
        if (! allPathsDeleted(buildPerfSources(L"precalc-settings-perf4")))
        {
            Fail(L"Worker-budget-4 pre-calc delete did not remove all perf source trees.");
            return true;
        }

        g_settings.fileOperations->preCalcEnabled    = true;
        g_settings.fileOperations->preCalcMaxWorkers = 8u;
        state.taskA                                  = StartFileOperationAndGetId(state.fileOps,
                                                                                  FILESYSTEM_DELETE,
                                                                                  FolderWindow::Pane::Left,
                                                                                  std::nullopt,
                                                                                  state.fsLocal,
                                                                                  buildPerfSources(L"precalc-settings-perf8"),
                                                                                  {},
                                                                                  flags,
                                                                                  false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start pre-calc perf delete task for worker budget 8.");
            return true;
        }

        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        const auto completionIt = state.completedTasks.find(state.taskA.value());
        if (completionIt == state.completedTasks.end())
        {
            return false;
        }

        const CompletedTaskInfo& completion = completionIt->second;
        if (FAILED(completion.hr))
        {
            Fail(std::format(L"Unexpected hr for worker-budget-8 pre-calc delete: 0x{:08X}", static_cast<unsigned long>(completion.hr)));
            return true;
        }
        if (! completion.preCalcCompleted || completion.preCalcWorkerCountUsed != 8u)
        {
            Fail(std::format(L"Expected worker-budget-8 task to use 8 workers (completed={} workers={}).",
                             completion.preCalcCompleted ? 1 : 0,
                             completion.preCalcWorkerCountUsed));
            return true;
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.ClearflowPreCalcMultiRootWorkers",
                          L"shape=8-root-delete workerBudget=8",
                          completion.preCalcDurationUs,
                          completion.preCalcWorkerCountUsed,
                          completion.preCalcTotalBytes,
                          completion.hr);
        if (! allPathsDeleted(buildPerfSources(L"precalc-settings-perf8")))
        {
            Fail(L"Worker-budget-8 pre-calc delete did not remove all perf source trees.");
            return true;
        }

        const std::string config =
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":1,"deleteRecycleBinMaxConcurrency":1,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"directorySizeDelayMs":1})json";
        if (! SetPluginConfiguration(state.infoLocal.get(), config))
        {
            Fail(L"Failed to apply local plugin config for pre-calc fan-out test.");
            return true;
        }

        g_settings.fileOperations->preCalcEnabled    = true;
        g_settings.fileOperations->preCalcMaxWorkers = 1u;
        state.taskB                                  = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {preCalcFanOutBudget1}, {}, flags, false);
        if (! state.taskB.has_value())
        {
            Fail(L"Failed to start single-root pre-calc fan-out delete task for worker budget 1.");
            return true;
        }

        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        const auto completionIt = state.completedTasks.find(state.taskB.value());
        if (completionIt == state.completedTasks.end())
        {
            return false;
        }

        const CompletedTaskInfo& completion = completionIt->second;
        if (FAILED(completion.hr))
        {
            Fail(std::format(L"Unexpected hr for single-root worker-budget-1 pre-calc delete: 0x{:08X}", static_cast<unsigned long>(completion.hr)));
            return true;
        }
        if (! completion.preCalcCompleted || completion.preCalcTotalBytes == 0 || completion.preCalcWorkerCountUsed != 1u)
        {
            Fail(std::format(L"Expected single-root worker-budget-1 task to use 1 worker (completed={} bytes={} workers={}).",
                             completion.preCalcCompleted ? 1 : 0,
                             completion.preCalcTotalBytes,
                             completion.preCalcWorkerCountUsed));
            return true;
        }
        if (! pathDeleted(preCalcFanOutBudget1))
        {
            Fail(L"Single-root worker-budget-1 pre-calc delete did not remove the source tree.");
            return true;
        }

        g_settings.fileOperations->preCalcEnabled    = true;
        g_settings.fileOperations->preCalcMaxWorkers = 4u;
        state.taskC                                  = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {preCalcFanOutBudget4}, {}, flags, false);
        if (! state.taskC.has_value())
        {
            Fail(L"Failed to start single-root pre-calc fan-out delete task for worker budget 4.");
            return true;
        }

        state.stepState = 6;
        return false;
    }

    const auto completionIt = state.completedTasks.find(state.taskC.value());
    if (completionIt == state.completedTasks.end())
    {
        return false;
    }

    const CompletedTaskInfo& completion = completionIt->second;
    if (FAILED(completion.hr))
    {
        Fail(std::format(L"Unexpected hr for single-root worker-budget-4 pre-calc delete: 0x{:08X}", static_cast<unsigned long>(completion.hr)));
        return true;
    }
    if (! completion.preCalcCompleted || completion.preCalcTotalBytes == 0 || completion.preCalcWorkerCountUsed != 4u)
    {
        Fail(std::format(L"Expected single-root worker-budget-4 task to use 4 workers (completed={} bytes={} workers={}).",
                         completion.preCalcCompleted ? 1 : 0,
                         completion.preCalcTotalBytes,
                         completion.preCalcWorkerCountUsed));
        return true;
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.ClearflowPreCalcSingleRootFanOutWorkers",
                      L"shape=single-root-wide-fanout workerBudget=4",
                      completion.preCalcDurationUs,
                      completion.preCalcWorkerCountUsed,
                      completion.preCalcTotalBytes,
                      completion.hr);
    if (! pathDeleted(preCalcFanOutBudget4))
    {
        Fail(L"Single-root worker-budget-4 pre-calc delete did not remove the source tree.");
        return true;
    }
    if (! state.localConfigOriginal.empty() && ! SetPluginConfiguration(state.infoLocal.get(), state.localConfigOriginal))
    {
        Fail(L"Failed to restore local plugin config after pre-calc fan-out test.");
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase5_PreCalcCancelReleasesSlot);
    return false;
}
case SelfTestState::Step::Phase5_PreCalcCancelReleasesSlot:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        Fail(L"Phase5_PreCalcCancelReleasesSlot timed out.");
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
                                                 std::filesystem::path(L"/dest-a"),
                                                 flags,
                                                 false);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[1])},
                                                 std::filesystem::path(L"/dest-b"),
                                                 flags,
                                                 true);

        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            Fail(L"Failed to start dummy copy tasks for pre-calc cancel test.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
        if (taskA && taskA->_preCalcInProgress.load(std::memory_order_acquire))
        {
            const HWND popup = state.fileOps ? state.fileOps->GetPopupHwndForSelfTest() : nullptr;
            if (! popup || IsWindow(popup) == FALSE)
            {
                Fail(L"Pre-calc copy task did not expose the File Operations popup for layout validation.");
                return true;
            }

            FileOperationsPopupInternal::PopupLayoutDebugSnapshot layout{};
            layout.taskId = state.taskA.value();
            if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, layout))
            {
                Fail(L"Failed to capture File Operations popup layout while copy pre-calc was in progress.");
                return true;
            }

            // 5F early admission: a COPY runs pre-calc CONCURRENTLY with the transfer, so once
            // execution has started the truthful status is Running (transferring, ETA estimating)
            // rather than the blocking Calculating. Accept either: Calculating only in the brief
            // window before the transfer's operationStartTick is set, Running thereafter.
            if (layout.taskStatusKind != FileOperationsPopupInternal::TaskSnapshot::StatusKind::Calculating &&
                layout.taskStatusKind != FileOperationsPopupInternal::TaskSnapshot::StatusKind::Running)
            {
                Fail(L"Pre-calc copy task layout should report the Calculating or (early-admission) Running status.");
                return true;
            }

            if (! layout.taskToggleCollapseVisible || ! layout.taskSkipVisible || ! layout.taskCancelVisible || ! layout.taskSpeedLimitVisible)
            {
                Fail(std::format(L"Pre-calc copy task should expose collapse, Skip, Speed Limit, and Cancel controls; collapse={} skip={} speedLimit={} cancel={}.",
                                 layout.taskToggleCollapseVisible ? 1 : 0,
                                 layout.taskSkipVisible ? 1 : 0,
                                 layout.taskSpeedLimitVisible ? 1 : 0,
                                 layout.taskCancelVisible ? 1 : 0));
                return true;
            }

            taskA->RequestCancel();
            state.stepState = 2;
        }
        return false;
    }

    if (state.stepState == 2)
    {
        const auto itA = state.completedTasks.find(state.taskA.value());
        if (itA == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT hrA = itA->second.hr;
        if (hrA != HRESULT_FROM_WIN32(ERROR_CANCELLED) && hrA != E_ABORT)
        {
            Fail(std::format(L"Unexpected hr for cancelled pre-calc task: 0x{:08X}", static_cast<unsigned long>(hrA)));
            return true;
        }

        if (state.completedTasks.find(state.taskB.value()) != state.completedTasks.end())
        {
            state.stepState = 3;
            return false;
        }

        FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());
        if (! taskB || ! taskB->HasEnteredOperation())
        {
            return false;
        }

        taskB->RequestCancel();
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (state.completedTasks.find(state.taskB.value()) == state.completedTasks.end())
        {
            return false;
        }

        NextStep(state, SelfTestState::Step::Phase5_PreCalcCancelLatencyLocal);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase5_PreCalcCancelLatencyLocal:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        const std::wstring status = state.taskA.has_value() ? std::format(L"task={} completionSeen={} stepState={}",
                                                                          state.taskA.value(),
                                                                          state.completedTasks.contains(state.taskA.value()) ? 1 : 0,
                                                                          state.stepState)
                                                            : L"task=(none)";
        Fail(std::format(L"Phase5_PreCalcCancelLatencyLocal timed out. {}", status));
        return true;
    }

    constexpr ULONGLONG kCancelLatencyThresholdMs = 150ull;
    constexpr unsigned int kDirectorySizeDelayMs  = 50u;
    const FileSystemFlags flags                   = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    if (state.stepState == 0)
    {
        state.fileOps->ApplyQueueMode(false);
        state.taskA.reset();
        state.taskB.reset();
        state.markerTick = 0;

        const std::string config = std::format(
            R"json({{"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":1,"deleteRecycleBinMaxConcurrency":1,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"directorySizeDelayMs":{}}})json",
            kDirectorySizeDelayMs);
        if (! SetPluginConfiguration(state.infoLocal.get(), config))
        {
            Fail(L"Failed to apply local plugin config for pre-calc cancel latency test.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {state.tempRoot / L"precalc-cancel-latency"},
                                                 {},
                                                 flags,
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start local delete task for pre-calc cancel latency test.");
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

        if (! task->_preCalcInProgress.load(std::memory_order_acquire))
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
            Fail(std::format(L"Unexpected hr for pre-calc cancel latency task: 0x{:08X}", static_cast<unsigned long>(completion.hr)));
            return true;
        }

        const ULONGLONG cancelLatencyMs =
            (state.markerTick != 0 && completion.completionTick >= state.markerTick) ? (completion.completionTick - state.markerTick) : 0ull;
        AppendLog(std::format(
            L"Phase5_PreCalcCancelLatencyLocal latency={}ms threshold={}ms delay={}ms", cancelLatencyMs, kCancelLatencyThresholdMs, kDirectorySizeDelayMs));
        Debug::Perf::Emit(
            L"FileOps.SelfTest.PreCalcCancelLatency",
            std::format(
                L"delayMs={} thresholdMs={} path={}", kDirectorySizeDelayMs, kCancelLatencyThresholdMs, (state.tempRoot / L"precalc-cancel-latency").wstring()),
            cancelLatencyMs * 1000ull,
            kDirectorySizeDelayMs,
            kCancelLatencyThresholdMs,
            completion.hr);

        if (state.markerTick == 0)
        {
            Fail(L"Pre-calc cancel latency marker was not recorded.");
            return true;
        }

        if (cancelLatencyMs > kCancelLatencyThresholdMs)
        {
            Fail(std::format(L"Pre-calc cancel latency too high: {}ms (threshold {}ms).", cancelLatencyMs, kCancelLatencyThresholdMs));
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase5_PreCalcSkipContinues);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase5_PreCalcSkipContinues:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        Fail(L"Phase5_PreCalcSkipContinues timed out.");
        return true;
    }

    if (state.stepState == 0)
    {
        state.fileOps->ApplyQueueMode(true);

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                                   FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[0])},
                                                 std::filesystem::path(L"/dest-skip-a"),
                                                 flags,
                                                 false);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[1])},
                                                 std::filesystem::path(L"/dest-skip-b"),
                                                 flags,
                                                 true);

        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            Fail(L"Failed to start dummy copy tasks for pre-calc skip test.");
            return true;
        }

        if (auto* taskA = state.fileOps->FindTask(state.taskA.value()))
        {
            taskA->SetDesiredSpeedLimit(8ull * 1024ull);
            taskA->SkipPreCalculation();
        }
        if (auto* taskB = state.fileOps->FindTask(state.taskB.value()))
        {
            taskB->SetDesiredSpeedLimit(8ull * 1024ull);
        }

        state.stepState = 1;
        return false;
    }

    FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
    FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());

    if (state.stepState == 1)
    {
        if (state.completedTasks.find(state.taskA.value()) != state.completedTasks.end() ||
            state.completedTasks.find(state.taskB.value()) != state.completedTasks.end())
        {
            Fail(L"Pre-calc skip tasks completed before validation could run.");
            return true;
        }

        if (taskA && ! taskA->_preCalcSkipped.load(std::memory_order_acquire))
        {
            taskA->SkipPreCalculation();
        }

        if (! taskA || ! taskB)
        {
            return false;
        }

        if (taskA->_preCalcCompleted.load(std::memory_order_acquire))
        {
            Fail(L"Pre-calc completed despite Skip being requested.");
            return true;
        }

        if (! taskA->_preCalcSkipped.load(std::memory_order_acquire))
        {
            return false;
        }

        if (! taskA->HasStarted())
        {
            return false;
        }

        if (! taskB->IsWaitingInQueue())
        {
            Fail(L"Skipping pre-calc released the queue slot unexpectedly.");
            return true;
        }

        taskA->RequestCancel();
        taskB->RequestCancel();
        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (state.completedTasks.find(state.taskA.value()) == state.completedTasks.end())
        {
            return false;
        }
        if (state.completedTasks.find(state.taskB.value()) == state.completedTasks.end())
        {
            return false;
        }

        NextStep(state, SelfTestState::Step::Phase5_CancelQueuedTask);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase5_CancelQueuedTask:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        Fail(L"Phase5_CancelQueuedTask timed out.");
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
                                                 std::filesystem::path(L"/dest-queued-a"),
                                                 flags,
                                                 false);
        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[1])},
                                                 std::filesystem::path(L"/dest-queued-b"),
                                                 flags,
                                                 true);
        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            Fail(L"Failed to start dummy copy tasks for queued-cancel test.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
    FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());

    if (state.stepState == 1)
    {
        if (taskB && taskB->IsWaitingInQueue())
        {
            taskB->RequestCancel();
            state.stepState = 2;
        }
        return false;
    }

    if (state.stepState == 2)
    {
        if (state.completedTasks.find(state.taskB.value()) == state.completedTasks.end())
        {
            return false;
        }

        state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(state.dummyPaths[2 % state.dummyPaths.size()])},
                                                 std::filesystem::path(L"/dest-queued-c"),
                                                 flags,
                                                 true);
        if (! state.taskC.has_value())
        {
            Fail(L"Failed to start follow-up task after cancelling queued task.");
            return true;
        }

        if (taskA)
        {
            taskA->RequestCancel();
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
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
        state.stepState = 4;
        return false;
    }

    if (state.completedTasks.find(state.taskC.value()) == state.completedTasks.end())
    {
        return false;
    }

    if (state.completedTasks.find(state.taskA.value()) == state.completedTasks.end())
    {
        return false;
    }

    NextStep(state, SelfTestState::Step::Phase5_SwitchParallelToWaitDuringPreCalc);
    return false;
}
case SelfTestState::Step::Phase5_SwitchParallelToWaitDuringPreCalc:
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

            return std::format(L"id={} entered={} started={} qpause={} preCalc={} preDone={} preSkipped={} items={}/{}",
                               id,
                               task->HasEnteredOperation(),
                               task->HasStarted(),
                               task->IsQueuePaused(),
                               task->_preCalcInProgress.load(std::memory_order_acquire),
                               task->_preCalcCompleted.load(std::memory_order_acquire),
                               task->_preCalcSkipped.load(std::memory_order_acquire),
                               completedItems,
                               totalItems);
        };

        Fail(std::format(L"Phase5_SwitchParallelToWaitDuringPreCalc timed out. A: {} B: {}", summarizeTask(state.taskA), summarizeTask(state.taskB)));
        return true;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    if (state.stepState == 0)
    {
        state.queuePausedTask.reset();
        state.fileOps->ApplyQueueMode(false);

        // Make deletion slow/predictable so we can reliably observe pre-calc and queue-pause behavior.
        const std::string config =
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":1,"deleteRecycleBinMaxConcurrency":1,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"directorySizeDelayMs":1})json";
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), config));

        state.taskA = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {state.tempRoot / L"precalc-a"}, {}, flags, false);
        state.taskB = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {state.tempRoot / L"precalc-b"}, {}, flags, false);
        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            Fail(L"Failed to start local delete tasks for Parallel->Wait switch test.");
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

        if (! pausedTask->_preCalcInProgress.load(std::memory_order_acquire))
        {
            return false;
        }

        pausedTask->SkipPreCalculation();
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

    const bool preCalcStill = pausedTask->_preCalcInProgress.load(std::memory_order_acquire);
    if (pausedTask->HasStarted() && pausedTask->IsQueuePaused())
    {
        Fail(L"Queue-paused task started operation unexpectedly.");
        return true;
    }

    if (preCalcStill)
    {
        return false;
    }

    if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) < 500ull)
    {
        return false;
    }

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
                return std::format(L"id={} started={} qpause={} preCalc={} done={} skipped={}",
                                   id,
                                   task->HasStarted(),
                                   task->IsQueuePaused(),
                                   task->_preCalcInProgress.load(std::memory_order_acquire),
                                   task->_preCalcCompleted.load(std::memory_order_acquire),
                                   task->_preCalcSkipped.load(std::memory_order_acquire));
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
                Fail(std::format(L"Paused task completed before it resumed (hr=0x{:08X})", static_cast<unsigned long>(it->second.hr)));
                return true;
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
case SelfTestState::Step::Phase6_PopupRateSmoothing:
{
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

    NextStep(state, SelfTestState::Step::Phase6_PopupSmokeResizeAndPause);
    return false;
}
case SelfTestState::Step::Phase6_PopupSmokeResizeAndPause:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        const HWND popup     = FindWindowW(kPopupClassName.data(), nullptr);
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
    const std::filesystem::path srcFile = srcDir / L"big.bin";

    if (state.stepState == 0)
    {
        state.fileOps->ApplyQueueMode(false);
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

        if (! WriteTestFile(srcFile, 32ull * 1024ull * 1024ull))
        {
            Fail(L"Failed to write large source file for popup smoke test.");
            return true;
        }

        std::vector<std::filesystem::path> sources{srcFile};

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

        state.stepState = 1;
        return false;
    }

    const HWND popup = FindWindowW(kPopupClassName.data(), nullptr);
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

        task->TogglePause();
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
            if (it != state.completedTasks.end() && state.stepState < 6)
            {
                Fail(std::format(L"Copy task completed before pause/resize validation finished (hr=0x{:08X}).", static_cast<unsigned long>(it->second.hr)));
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

        const int height = state.popupOriginalRect.bottom - state.popupOriginalRect.top;
        SetWindowPos(popup, nullptr, state.popupOriginalRect.left, state.popupOriginalRect.top, 420, height, SWP_NOZORDER | SWP_NOACTIVATE);

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        task->TogglePause();
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
        const bool preCalcDone = task->_preCalcCompleted.load(std::memory_order_acquire);
        const uint64_t total   = task->_preCalcTotalBytes.load(std::memory_order_acquire);
        if (preCalcDone && total > 0)
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
    if (completion.preCalcCompleted && completion.preCalcTotalBytes > 0)
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
        Fail(L"Delete-bytes validation failed: did not observe a non-zero pre-calc total bytes.");
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
        if (! RecreateEmptyDirectory(durationDstDir))
        {
            Fail(L"Failed to reset phase6-bandwidth duration destination directory.");
            return true;
        }
        if (! RecreateEmptyDirectory(cancelDstDir))
        {
            Fail(L"Failed to reset phase6-bandwidth cancel destination directory.");
            return true;
        }

        if (! WriteTestFile(srcFile, kDurationFileBytes))
        {
            Fail(L"Failed to seed source file for local bandwidth throttle validation.");
            return true;
        }

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
    constexpr uint64_t kSpeedLimitBytesPerSecond   = 4ull * 1024ull * 1024ull;
    constexpr uint64_t kAllowedDurationSlowdownUs  = 250'000ull;
    constexpr uint64_t kAllowedSkewRegressionBytes = 1ull * 1024ull * 1024ull;
    constexpr uint64_t kMinSamplesPerRun           = 4ull;
    const std::filesystem::path srcDir             = state.tempRoot / L"phase6-parallel-bandwidth-src";
    const std::filesystem::path sharedDstDir       = state.tempRoot / L"phase6-parallel-bandwidth-shared-dst";
    const std::filesystem::path perWorkerDstDir    = state.tempRoot / L"phase6-parallel-bandwidth-perworker-dst";

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
            R"json({{"concurrencyMode":"manual","copyMoveMaxConcurrency":{},"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"recycleBinBatchSize":500,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"copyReparse","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4}})json",
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

    const auto sampleTask = [&](FolderWindow::FileOperationState::Task& task, uint64_t& maxSkewBytes, size_t& maxActiveStreams, uint64_t& sampleCount) noexcept
    {
        std::scoped_lock lock(task._progressMutex);

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
            sampleTask(*task, state.parallelBandwidthBaselineMaxSkewBytes, state.parallelBandwidthBaselineMaxActive, state.parallelBandwidthBaselineSamples);
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
        sampleTask(*task, state.parallelBandwidthCandidateMaxSkewBytes, state.parallelBandwidthCandidateMaxActive, state.parallelBandwidthCandidateSamples);
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
        L"fileCount={} fileBytes={} speedLimitBps={} copyConcurrency={} sharedMaxActive={} perWorkerMaxActive={} sharedSamples={} perWorkerSamples={}",
        kFileCount,
        kFileBytes,
        kSpeedLimitBytesPerSecond,
        kCopyConcurrency,
        state.parallelBandwidthBaselineMaxActive,
        state.parallelBandwidthCandidateMaxActive,
        state.parallelBandwidthBaselineSamples,
        state.parallelBandwidthCandidateSamples);
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
        std::format(L"Phase6_ParallelBandwidthThrottleFairness shared={}us perWorker={}us sharedSkew={}B perWorkerSkew={}B sharedActive={} perWorkerActive={}",
                    state.parallelBandwidthBaselineUs,
                    state.parallelBandwidthCandidateUs,
                    state.parallelBandwidthBaselineMaxSkewBytes,
                    state.parallelBandwidthCandidateMaxSkewBytes,
                    state.parallelBandwidthBaselineMaxActive,
                    state.parallelBandwidthCandidateMaxActive));

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

    const uint64_t maxAllowedCandidateSkewBytes = state.parallelBandwidthBaselineMaxSkewBytes + kAllowedSkewRegressionBytes;
    if (state.parallelBandwidthCandidateMaxSkewBytes > maxAllowedCandidateSkewBytes)
    {
        Fail(std::format(L"Per-worker bandwidth fairness regressed skew too much: candidate={}B maxAllowed={}B baseline={}B.",
                         state.parallelBandwidthCandidateMaxSkewBytes,
                         maxAllowedCandidateSkewBytes,
                         state.parallelBandwidthBaselineMaxSkewBytes));
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase7_WatcherChurn);
    return false;
}
