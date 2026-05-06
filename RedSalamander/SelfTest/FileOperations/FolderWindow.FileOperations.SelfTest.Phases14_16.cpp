case SelfTestState::Step::Phase14_PopupHostLifetimeGuard:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 30'000ull))
    {
        const HWND popup = FindWindowW(kPopupClassName.data(), nullptr);
        Fail(std::format(L"Phase14_PopupHostLifetimeGuard timed out. stepState={} popup={} shutdownDone={}",
                         state.stepState,
                         popup != nullptr,
                         state.phase14ShutdownDone.load(std::memory_order_acquire)));
        return true;
    }

    if (state.stepState == 0)
    {
        state.phase14ShutdownDone.store(false, std::memory_order_release);

        FolderWindow::InformationalTaskUpdate update{};
        update.kind  = FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories;
        update.title = L"FileOpsSelfTest: Phase 14";

        update.leftRoot  = L"/";
        update.rightRoot = L"/";

        update.scanActive          = true;
        update.scanCurrentRelative = L"phase14";
        update.scanFolderCount     = 1;
        update.scanEntryCount      = 1;

        update.finished = false;
        update.resultHr = S_OK;

        const uint64_t infoTaskId = state.fileOps ? state.fileOps->CreateOrUpdateInformationalTask(update) : 0;
        if (infoTaskId == 0)
        {
            Fail(L"Phase14_PopupHostLifetimeGuard failed to create an informational task.");
            return true;
        }

        state.phase14InfoTask = infoTaskId;
        state.stepState       = 1;
        return false;
    }

    const HWND popup = FindWindowW(kPopupClassName.data(), nullptr);

    if (state.stepState == 1)
    {
        if (! popup)
        {
            return false;
        }

        ShowWindow(popup, SW_HIDE);
        SetWindowPos(popup, nullptr, -32000, -32000, 0, 0, SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOZORDER);

        FileOperationsPopupInternal::PopupSelfTestInvoke destroyOnNextShow{};
        destroyOnNextShow.data = FileOperationsPopupInternal::kPopupSelfTestDestroyOnNextShowData;
        if (! InvokePopupSelfTest(popup, destroyOnNextShow))
        {
            Fail(L"Phase14_PopupHostLifetimeGuard failed to arm popup stale-HWND selftest.");
            return true;
        }

        const std::filesystem::path reentryRoot = state.tempRoot / L"phase14-popup-reentry";
        const std::filesystem::path sourceDir   = reentryRoot / L"src";
        const std::filesystem::path destDir     = reentryRoot / L"dst";
        const std::filesystem::path sourceFile  = sourceDir / L"source.bin";
        if (! RecreateEmptyDirectory(sourceDir) || ! RecreateEmptyDirectory(destDir))
        {
            Fail(L"Phase14_PopupHostLifetimeGuard failed to reset popup reentry folders.");
            return true;
        }

        if (! WriteTestFile(sourceFile, 2048))
        {
            Fail(L"Phase14_PopupHostLifetimeGuard failed to write popup reentry source file.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsLocal,
                                                 {sourceFile},
                                                 destDir,
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Phase14_PopupHostLifetimeGuard failed to start popup reentry copy task.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (! state.taskA.has_value())
        {
            Fail(L"Phase14_PopupHostLifetimeGuard lost popup reentry task id.");
            return true;
        }

        const auto completed = state.completedTasks.find(state.taskA.value());
        if (completed == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(completed->second.hr))
        {
            Fail(std::format(L"Phase14_PopupHostLifetimeGuard popup reentry copy failed: 0x{:08X}.",
                             static_cast<unsigned long>(completed->second.hr)));
            return true;
        }

        const std::filesystem::path copiedFile = state.tempRoot / L"phase14-popup-reentry" / L"dst" / L"source.bin";
        std::error_code existsEc;
        if (! std::filesystem::exists(copiedFile, existsEc) || existsEc)
        {
            Fail(L"Phase14_PopupHostLifetimeGuard popup reentry copy did not produce the destination file.");
            return true;
        }

        if (! FindWindowW(kPopupClassName.data(), nullptr))
        {
            Fail(L"Phase14_PopupHostLifetimeGuard did not recreate the popup after it was destroyed during ShowWindow.");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        auto work     = std::make_unique<Phase14ShutdownWork>();
        work->fileOps = state.fileOps;
        work->done    = &state.phase14ShutdownDone;

        if (! TrySubmitThreadpoolCallback(Phase14ShutdownCallback, work.get(), nullptr))
        {
            Fail(L"Phase14_PopupHostLifetimeGuard could not submit shutdown callback.");
            return true;
        }

        static_cast<void>(work.release());
        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        if (! state.phase14ShutdownDone.load(std::memory_order_acquire))
        {
            return false;
        }

        const HWND popupAfterShutdown = FindWindowW(kPopupClassName.data(), nullptr);
        if (! popupAfterShutdown)
        {
            // Popup already self-closed after host lifetime ended; that's acceptable as long as we didn't crash.
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
            NextStep(state, SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke);
            return false;
        }

        FileOperationsPopupInternal::PopupSelfTestInvoke dismiss{};
        dismiss.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskDismiss;
        dismiss.taskId = state.phase14InfoTask.value_or(0);
        static_cast<void>(InvokePopupSelfTest(popupAfterShutdown, dismiss));

        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        if (FindWindowW(kPopupClassName.data(), nullptr))
        {
            return false;
        }

        RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
        NextStep(state, SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke:
{
    if (! state.fs7z)
    {
        RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::skipped, L"7z plugin not loaded.");
        NextStep(state, SelfTestState::Step::Phase15_FileSystem7zMountPathImpact);
        return false;
    }

    RecordCurrentPhase(state,
                       SelfTest::SelfTestCaseResult::Status::skipped,
                       L"7z read/seek smoke skipped in the file-operations suite; 7z coverage remains in compare selftests.");
    NextStep(state, SelfTestState::Step::Phase15_FileSystem7zMountPathImpact);
    return false;
}
case SelfTestState::Step::Phase15_FileSystem7zMountPathImpact:
{
    if (! state.fs7z)
    {
        RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::skipped, L"7z plugin not loaded.");
        NextStep(state, SelfTestState::Step::Phase16_RemoteWatchContractExposure);
        return false;
    }

    RecordCurrentPhase(state,
                       SelfTest::SelfTestCaseResult::Status::skipped,
                       L"7z mount-path impact skipped in the file-operations suite; 7z coverage remains in compare selftests.");
    NextStep(state, SelfTestState::Step::Phase16_RemoteWatchContractExposure);
    return false;
}
case SelfTestState::Step::Phase16_RemoteWatchContractExposure:
{
    const std::array<std::pair<std::wstring_view, std::wstring_view>, 5> plugins = {
        std::pair{kPluginIdFtp, L"FTP"},
        std::pair{kPluginIdSftp, L"SFTP"},
        std::pair{kPluginIdScp, L"SCP"},
        std::pair{kPluginIdImap, L"IMAP"},
        std::pair{kPluginIdS3, L"S3"},
    };

    FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
    for (const auto& [pluginId, label] : plugins)
    {
        const HRESULT enableHr = pluginManager.EnablePlugin(pluginId, g_settings);
        if (FAILED(enableHr))
        {
            Fail(std::format(L"Phase16_RemoteWatchContractExposure failed to enable {} plugin (hr=0x{:08X}).", label, static_cast<unsigned long>(enableHr)));
            return true;
        }

        const FileSystemPluginManager::PluginEntry* entry = FindLoadedPluginEntry(pluginId);
        if (! entry || ! entry->fileSystem)
        {
            Fail(std::format(L"Phase16_RemoteWatchContractExposure could not locate loaded {} plugin entry.", label));
            return true;
        }

        wil::com_ptr<IFileSystemDirectoryWatch> watch;
        const HRESULT hrWatch = entry->fileSystem->QueryInterface(__uuidof(IFileSystemDirectoryWatch), watch.put_void());
        if (FAILED(hrWatch) || ! watch)
        {
            AppendLog(std::format(L"Phase16_RemoteWatchContractExposure {} does not expose IFileSystemDirectoryWatch (hr=0x{:08X}); continuing.",
                                  label,
                                  static_cast<unsigned long>(hrWatch)));
            continue;
        }

        WatchCallback callback{};
        const HRESULT registerHr = watch->WatchDirectory(L"/", &callback, nullptr);
        if (FAILED(registerHr))
        {
            Fail(std::format(
                L"Phase16_RemoteWatchContractExposure {} WatchDirectory('/') failed (hr=0x{:08X}).", label, static_cast<unsigned long>(registerHr)));
            return true;
        }

        const HRESULT duplicateHr = watch->WatchDirectory(L"/", &callback, nullptr);
        if (duplicateHr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
        {
            Fail(std::format(L"Phase16_RemoteWatchContractExposure {} duplicate WatchDirectory expected ERROR_ALREADY_EXISTS, got 0x{:08X}.",
                             label,
                             static_cast<unsigned long>(duplicateHr)));
            return true;
        }

        const HRESULT unwatchHr = watch->UnwatchDirectory(L"/");
        if (FAILED(unwatchHr))
        {
            Fail(std::format(
                L"Phase16_RemoteWatchContractExposure {} UnwatchDirectory('/') failed (hr=0x{:08X}).", label, static_cast<unsigned long>(unwatchHr)));
            return true;
        }

        const HRESULT missingHr = watch->UnwatchDirectory(L"/");
        if (missingHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            Fail(std::format(L"Phase16_RemoteWatchContractExposure {} second UnwatchDirectory expected ERROR_FILE_NOT_FOUND, got 0x{:08X}.",
                             label,
                             static_cast<unsigned long>(missingHr)));
            return true;
        }
    }

    RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
    NextStep(state, SelfTestState::Step::Phase16_RemoteFtpSecret);
    return false;
}
case SelfTestState::Step::Phase16_RemoteFtpSecret:
{
    const PhaseCheckResult outcome = CheckRemoteConnectionSecret(L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kPluginIdFtp);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteFtpSandbox);
    return false;
}
case SelfTestState::Step::Phase16_RemoteFtpSandbox:
{
    const PhaseCheckResult outcome = CheckRemoteConnectionSandbox(L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kPluginIdFtp);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteSftpSecret);
    return false;
}
case SelfTestState::Step::Phase16_RemoteSftpSecret:
{
    const PhaseCheckResult outcome = CheckRemoteConnectionSecret(L"SFTP", kSelfTestEnvConnSftp, kSelfTestDefaultConnSftp, kPluginIdSftp);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteSftpSandbox);
    return false;
}
case SelfTestState::Step::Phase16_RemoteSftpSandbox:
{
    const PhaseCheckResult outcome = CheckRemoteConnectionSandbox(L"SFTP", kSelfTestEnvConnSftp, kSelfTestDefaultConnSftp, kPluginIdSftp);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteScpSecret);
    return false;
}
case SelfTestState::Step::Phase16_RemoteScpSecret:
{
    const PhaseCheckResult outcome = CheckRemoteConnectionSecret(L"SCP", kSelfTestEnvConnScp, kSelfTestDefaultConnScp, kPluginIdScp);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteScpSandbox);
    return false;
}
case SelfTestState::Step::Phase16_RemoteScpSandbox:
{
    const PhaseCheckResult outcome = CheckRemoteConnectionSandbox(L"SCP", kSelfTestEnvConnScp, kSelfTestDefaultConnScp, kPluginIdScp);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteImapSecret);
    return false;
}
case SelfTestState::Step::Phase16_RemoteImapSecret:
{
    const PhaseCheckResult outcome = CheckRemoteConnectionSecret(L"IMAP", kSelfTestEnvConnImap, kSelfTestDefaultConnImap, kPluginIdImap);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteImapSandbox);
    return false;
}
case SelfTestState::Step::Phase16_RemoteImapSandbox:
{
    const PhaseCheckResult outcome = CheckRemoteConnectionSandbox(L"IMAP", kSelfTestEnvConnImap, kSelfTestDefaultConnImap, kPluginIdImap);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteS3Secret);
    return false;
}
case SelfTestState::Step::Phase16_RemoteS3Secret:
{
    const PhaseCheckResult outcome = CheckRemoteConnectionSecret(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kPluginIdS3);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteS3Sandbox);
    return false;
}
case SelfTestState::Step::Phase16_RemoteS3Sandbox:
{
    const PhaseCheckResult outcome = CheckRemoteConnectionSandbox(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kPluginIdS3);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteS3FileOps);
    return false;
}
case SelfTestState::Step::Phase16_RemoteS3FileOps:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 420'000ull))
    {
        Fail(std::format(L"Phase16_RemoteS3FileOps timed out (stepState={} taskA={} taskB={} taskC={} caseRoot='{}').",
                         state.stepState,
                         state.taskA.value_or(0ull),
                         state.taskB.value_or(0ull),
                         state.taskC.value_or(0ull),
                         state.remoteS3CaseRootConn));
        return true;
    }

    auto getRemoteIo = [&]() noexcept -> wil::com_ptr<IFileSystemIO>
    {
        wil::com_ptr<IFileSystemIO> io;
        if (state.fsRemoteS3)
        {
            static_cast<void>(state.fsRemoteS3->QueryInterface(IID_PPV_ARGS(io.put())));
        }
        return io;
    };

    auto getLocalIo = [&]() noexcept -> wil::com_ptr<IFileSystemIO>
    {
        wil::com_ptr<IFileSystemIO> io;
        if (state.fsLocal)
        {
            static_cast<void>(state.fsLocal->QueryInterface(IID_PPV_ARGS(io.put())));
        }
        return io;
    };

    constexpr uint64_t kRemoteS3LargeFileBytes = 65ull * 1024ull * 1024ull + 32'123ull;

    if (state.stepState == 0)
    {
        const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kPluginIdS3);
        if (secretOutcome.status == SelfTest::SelfTestCaseResult::Status::failed)
        {
            Fail(secretOutcome.reason);
            return true;
        }

        if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
        {
            RecordCurrentPhase(state, secretOutcome.status, secretOutcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteOneDrivePersonalSecret);
            return false;
        }

        const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kPluginIdS3);
        if (sandboxOutcome.status == SelfTest::SelfTestCaseResult::Status::failed)
        {
            Fail(sandboxOutcome.reason);
            return true;
        }

        if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
        {
            RecordCurrentPhase(state, sandboxOutcome.status, sandboxOutcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteOneDrivePersonalSecret);
            return false;
        }

        CleanupRemoteS3Case(state);
        state.taskA.reset();
        state.taskB.reset();
        state.taskC.reset();

        const ResolvedRemoteProfile resolvedProfile        = ResolveRemoteConnectionProfile(kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kPluginIdS3);
        const std::wstring profileName                     = resolvedProfile.profileName;
        const Common::Settings::ConnectionProfile* profile = resolvedProfile.profile;
        if (! profile)
        {
            Fail(L"Phase16_RemoteS3FileOps could not resolve the connection profile after validation.");
            return true;
        }

        FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
        const HRESULT enableHr                 = pluginManager.EnablePlugin(kPluginIdS3, g_settings);
        if (FAILED(enableHr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps failed to enable plugin (hr=0x{:08X}).", static_cast<unsigned long>(enableHr)));
            return true;
        }

        const FileSystemPluginManager::PluginEntry* entry = FindLoadedPluginEntry(kPluginIdS3);
        if (! entry || ! entry->fileSystem)
        {
            Fail(L"Phase16_RemoteS3FileOps could not locate the loaded plugin entry.");
            return true;
        }

        wil::com_ptr<IFileSystemIO> remoteIo;
        const HRESULT remoteIoHr = entry->fileSystem->QueryInterface(IID_PPV_ARGS(remoteIo.put()));
        if (FAILED(remoteIoHr) || ! remoteIo)
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps missing IFileSystemIO (hr=0x{:08X}).", static_cast<unsigned long>(remoteIoHr)));
            return true;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> remoteDirOps;
        const HRESULT remoteDirHr = entry->fileSystem->QueryInterface(IID_PPV_ARGS(remoteDirOps.put()));
        if (FAILED(remoteDirHr) || ! remoteDirOps)
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps missing IFileSystemDirectoryOperations (hr=0x{:08X}).", static_cast<unsigned long>(remoteDirHr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> localIo = getLocalIo();
        if (! localIo)
        {
            Fail(L"Phase16_RemoteS3FileOps missing local IFileSystemIO.");
            return true;
        }

        const std::wstring guid = NewGuidString();
        if (guid.empty())
        {
            Fail(L"Phase16_RemoteS3FileOps failed to generate a GUID for the sandbox.");
            return true;
        }

        const std::wstring caseRootPlugin = JoinPluginPathForSelfTest(profile->initialPath, std::format(L"fileops-selftest-s3-{}", guid));
        const std::wstring caseRootConn   = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
        if (caseRootPlugin.empty() || caseRootConn.empty())
        {
            Fail(L"Phase16_RemoteS3FileOps failed to build sandbox paths.");
            return true;
        }

        const std::filesystem::path caseRoot(caseRootConn);
        const std::filesystem::path uploadDir         = caseRoot / L"upload";
        const std::filesystem::path moveDir           = caseRoot / L"move";
        const std::filesystem::path seedFile          = caseRoot / L"seed-remote.txt";
        const std::filesystem::path promoteProbeTemp  = caseRoot / L"bridge-promote-probe.tmp";
        const std::filesystem::path promoteProbeFinal = caseRoot / L"bridge-promote-probe.txt";
        const std::filesystem::path renamedFile       = moveDir / L"renamed-upload-source.txt";
        const std::filesystem::path largeRemoteSource = caseRoot / L"large-source.bin";

        const std::filesystem::path workspace          = state.tempRoot / L"s3-fileops";
        const std::filesystem::path downloadDir        = workspace / L"download";
        const std::filesystem::path uploadSource       = workspace / L"upload-source.txt";
        const std::filesystem::path dirUploadSource    = workspace / L"folder-upload";
        const std::filesystem::path dirUploadFile      = dirUploadSource / L"nested" / L"dir-payload.txt";
        const std::filesystem::path dirMoveDownloadDir = workspace / L"dir-move-download";
        if (! RecreateEmptyDirectory(workspace) || ! RecreateEmptyDirectory(downloadDir) || ! RecreateEmptyDirectory(dirMoveDownloadDir))
        {
            Fail(L"Phase16_RemoteS3FileOps failed to prepare the local workspace.");
            return true;
        }

        std::error_code dirUploadEc;
        if (! std::filesystem::create_directories(dirUploadFile.parent_path(), dirUploadEc) && dirUploadEc)
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps failed to prepare the local upload directory tree (ec={}).", dirUploadEc.value()));
            return true;
        }

        const std::string payload = std::format("s3-fileops-{}-roundtrip", static_cast<unsigned long long>(GetTickCount64()));
        if (! EnsureDirectoryExistsFsOps(remoteDirOps, caseRoot) || ! EnsureDirectoryExistsFsOps(remoteDirOps, uploadDir) ||
            ! EnsureDirectoryExistsFsOps(remoteDirOps, moveDir))
        {
            Fail(L"Phase16_RemoteS3FileOps failed to prepare the remote sandbox.");
            return true;
        }

        std::wstring altCaseRootConn;
        const ResolvedRemoteProfile altResolvedProfile = ResolveRemoteConnectionProfile(kSelfTestEnvConnS3Alt, kSelfTestDefaultConnS3Alt, kPluginIdS3);
        if (altResolvedProfile.profile && ! EqualsIgnoreCase(altResolvedProfile.profileName, profileName))
        {
            const PhaseCheckResult altSecretOutcome  = CheckRemoteConnectionSecret(L"S3 Alt", kSelfTestEnvConnS3Alt, kSelfTestDefaultConnS3Alt, kPluginIdS3);
            const PhaseCheckResult altSandboxOutcome = CheckRemoteConnectionSandbox(L"S3 Alt", kSelfTestEnvConnS3Alt, kSelfTestDefaultConnS3Alt, kPluginIdS3);
            if (altSecretOutcome.status == SelfTest::SelfTestCaseResult::Status::passed &&
                altSandboxOutcome.status == SelfTest::SelfTestCaseResult::Status::passed)
            {
                const std::wstring altCaseRootPlugin =
                    JoinPluginPathForSelfTest(altResolvedProfile.profile->initialPath, std::format(L"fileops-selftest-s3-{}", guid));
                altCaseRootConn = MakeConnectionPathForSelfTest(altResolvedProfile.profileName, altCaseRootPlugin);
                if (! altCaseRootConn.empty() && ! EnsureDirectoryExistsFsOps(remoteDirOps, std::filesystem::path(altCaseRootConn)))
                {
                    Fail(L"Phase16_RemoteS3FileOps failed to prepare the optional alternate S3 sandbox.");
                    return true;
                }
            }
        }

        if (! WriteFileTextFsIo(localIo, uploadSource, payload))
        {
            Fail(L"Phase16_RemoteS3FileOps failed to create the local upload file.");
            return true;
        }

        if (! WriteFileTextFsIo(localIo, dirUploadFile, payload + "-dir"))
        {
            Fail(L"Phase16_RemoteS3FileOps failed to create the local upload directory payload.");
            return true;
        }

        if (! WriteFileTextFsIo(remoteIo, seedFile, payload))
        {
            Fail(L"Phase16_RemoteS3FileOps failed to seed the remote probe file.");
            return true;
        }

        std::string seedReadBack;
        if (! ReadFileTextFsIo(remoteIo, seedFile, seedReadBack) || seedReadBack != payload)
        {
            Fail(L"Phase16_RemoteS3FileOps remote CreateFileReader readback failed.");
            return true;
        }

        const std::string promotePayload = payload + "-promote";
        if (! WriteFileTextFsIo(remoteIo, promoteProbeTemp, promotePayload))
        {
            Fail(L"Phase16_RemoteS3FileOps failed to seed the remote temp-promotion probe.");
            return true;
        }

        const std::wstring promoteProbeTempText  = ToPluginPathText(promoteProbeTemp);
        const std::wstring promoteProbeFinalText = ToPluginPathText(promoteProbeFinal);
        const HRESULT promoteProbeHr             = entry->fileSystem->MoveItem(promoteProbeTempText.c_str(),
                                                                               promoteProbeFinalText.c_str(),
                                                                               static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                                               nullptr,
                                                                               nullptr,
                                                                               nullptr);
        if (FAILED(promoteProbeHr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps direct temp promotion move failed: 0x{:08X}.", static_cast<unsigned long>(promoteProbeHr)));
            return true;
        }

        if (PathExistsFsIo(remoteIo, promoteProbeTemp) || ! PathExistsFsIo(remoteIo, promoteProbeFinal))
        {
            Fail(L"Phase16_RemoteS3FileOps direct temp promotion left unexpected probe paths.");
            return true;
        }

        std::string promoteReadBack;
        if (! ReadFileTextFsIo(remoteIo, promoteProbeFinal, promoteReadBack) || promoteReadBack != promotePayload)
        {
            Fail(L"Phase16_RemoteS3FileOps direct temp promotion verification failed.");
            return true;
        }

        if (! WritePatternFileFsIo(remoteIo, largeRemoteSource, kRemoteS3LargeFileBytes))
        {
            Fail(L"Phase16_RemoteS3FileOps failed to create the remote large-file upload.");
            return true;
        }

        uint64_t largeRemoteSize = 0;
        if (! GetFileSizeFsIo(remoteIo, largeRemoteSource, largeRemoteSize) || largeRemoteSize != kRemoteS3LargeFileBytes)
        {
            Fail(L"Phase16_RemoteS3FileOps remote large-file upload verification failed.");
            return true;
        }

        state.fsRemoteS3                 = entry->fileSystem;
        state.remoteS3ProfileName        = profileName;
        state.remoteS3CaseRootConn       = caseRootConn;
        state.remoteS3AltCaseRootConn    = altCaseRootConn;
        state.remoteS3UploadDirConn      = ToPluginPathText(uploadDir);
        state.remoteS3MoveDirConn        = ToPluginPathText(moveDir);
        state.remoteS3SeedFileConn       = ToPluginPathText(seedFile);
        state.remoteS3UploadedFileConn   = ToPluginPathText(uploadDir / uploadSource.filename());
        state.remoteS3MovedFileConn      = ToPluginPathText(moveDir / uploadSource.filename());
        state.remoteS3RenamedFileConn    = ToPluginPathText(renamedFile);
        state.remoteS3UploadedDirConn    = ToPluginPathText(caseRoot / dirUploadSource.filename());
        state.remoteS3DisplayPath        = MakePluginDisplayPath(entry->shortId, state.remoteS3CaseRootConn);
        state.remoteS3Workspace          = workspace;
        state.remoteS3UploadSource       = uploadSource;
        state.remoteS3DownloadDir        = downloadDir;
        state.remoteS3DownloadedFile     = downloadDir / uploadSource.filename();
        state.remoteS3DirUploadSource    = dirUploadSource;
        state.remoteS3DirMoveDownloadDir = dirMoveDownloadDir;
        state.remoteS3DirMovedFile       = dirMoveDownloadDir / dirUploadSource.filename() / L"nested" / L"dir-payload.txt";
        state.remoteS3Payload            = payload;

        state.folderWindow->SetFolderPath(FolderWindow::Pane::Left, state.remoteS3DisplayPath);
        state.folderWindow->SetFolderPath(FolderWindow::Pane::Right, state.remoteS3Workspace);
        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (! CurrentPanePathEquals(state.folderWindow, FolderWindow::Pane::Left, state.remoteS3DisplayPath) ||
            ! CurrentPanePathEquals(state.folderWindow, FolderWindow::Pane::Right, state.remoteS3Workspace))
        {
            return false;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsLocal,
                                                 {state.remoteS3UploadSource},
                                                 std::filesystem::path(state.remoteS3UploadDirConn),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsRemoteS3);
        if (! state.taskA.has_value())
        {
            Fail(L"Phase16_RemoteS3FileOps failed to start local->remote copy.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto itCopy = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (itCopy == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itCopy->second.hr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps local->remote copy failed: 0x{:08X}.", static_cast<unsigned long>(itCopy->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> remoteIo = getRemoteIo();
        if (! remoteIo)
        {
            Fail(L"Phase16_RemoteS3FileOps lost remote IFileSystemIO after upload.");
            return true;
        }

        const std::filesystem::path uploadedFile(state.remoteS3UploadedFileConn);
        if (! PathExistsFsIo(remoteIo, uploadedFile))
        {
            Fail(L"Phase16_RemoteS3FileOps upload did not materialize the remote file.");
            return true;
        }

        std::string uploadReadBack;
        if (! ReadFileTextFsIo(remoteIo, uploadedFile, uploadReadBack) || uploadReadBack != state.remoteS3Payload)
        {
            Fail(L"Phase16_RemoteS3FileOps remote upload verification failed.");
            return true;
        }

        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsRemoteS3,
                                                 {uploadedFile},
                                                 state.remoteS3DownloadDir,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskB.has_value())
        {
            Fail(L"Phase16_RemoteS3FileOps failed to start remote->local copy.");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        const auto itCopyBack = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (itCopyBack == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itCopyBack->second.hr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps remote->local copy failed: 0x{:08X}.", static_cast<unsigned long>(itCopyBack->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> localIo = getLocalIo();
        if (! localIo)
        {
            Fail(L"Phase16_RemoteS3FileOps lost local IFileSystemIO after download.");
            return true;
        }

        std::string downloadedText;
        if (! ReadFileTextFsIo(localIo, state.remoteS3DownloadedFile, downloadedText) || downloadedText != state.remoteS3Payload)
        {
            Fail(L"Phase16_RemoteS3FileOps downloaded file contents did not match the upload.");
            return true;
        }

        state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Left,
                                                 state.fsRemoteS3,
                                                 {std::filesystem::path(state.remoteS3UploadedFileConn)},
                                                 std::filesystem::path(state.remoteS3MoveDirConn),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsRemoteS3);
        if (! state.taskC.has_value())
        {
            Fail(L"Phase16_RemoteS3FileOps failed to start same-bucket move.");
            return true;
        }

        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        const auto itMove = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
        if (itMove == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itMove->second.hr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps same-bucket move failed: 0x{:08X}.", static_cast<unsigned long>(itMove->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> remoteIo = getRemoteIo();
        if (! remoteIo)
        {
            Fail(L"Phase16_RemoteS3FileOps lost remote IFileSystemIO after move.");
            return true;
        }

        if (PathExistsFsIo(remoteIo, std::filesystem::path(state.remoteS3UploadedFileConn)))
        {
            Fail(L"Phase16_RemoteS3FileOps move left the source file behind.");
            return true;
        }

        if (! PathExistsFsIo(remoteIo, std::filesystem::path(state.remoteS3MovedFileConn)))
        {
            Fail(L"Phase16_RemoteS3FileOps move did not materialize the destination file.");
            return true;
        }

        constexpr wchar_t kRenamedLeaf[]      = L"renamed-upload-source.txt";
        const FileSystemRenamePair renameItem = {
            .sizeBytes = sizeof(FileSystemRenamePair), .sourcePath = state.remoteS3MovedFileConn.c_str(), .newName = kRenamedLeaf};
        const HRESULT renameHr =
            state.fsRemoteS3->RenameItems(&renameItem, 1, static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE), nullptr, nullptr, nullptr);
        if (FAILED(renameHr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps RenameItems failed: 0x{:08X}.", static_cast<unsigned long>(renameHr)));
            return true;
        }

        if (PathExistsFsIo(remoteIo, std::filesystem::path(state.remoteS3MovedFileConn)) ||
            ! PathExistsFsIo(remoteIo, std::filesystem::path(state.remoteS3RenamedFileConn)))
        {
            Fail(L"Phase16_RemoteS3FileOps rename did not leave the expected remote paths.");
            return true;
        }

        std::string renamedReadBack;
        if (! ReadFileTextFsIo(remoteIo, std::filesystem::path(state.remoteS3RenamedFileConn), renamedReadBack) || renamedReadBack != state.remoteS3Payload)
        {
            Fail(L"Phase16_RemoteS3FileOps renamed file verification failed.");
            return true;
        }

        const std::filesystem::path caseRoot(state.remoteS3CaseRootConn);
        const std::filesystem::path remoteCopiedFile = caseRoot / L"copied-remote.txt";
        const std::wstring remoteCopiedFileText      = ToPluginPathText(remoteCopiedFile);
        const HRESULT remoteCopyHr                   = state.fsRemoteS3->CopyItem(state.remoteS3RenamedFileConn.c_str(),
                                                                                  remoteCopiedFileText.c_str(),
                                                                                  static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                                                  nullptr,
                                                                                  nullptr,
                                                                                  nullptr);
        if (FAILED(remoteCopyHr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps same-plugin copy failed: 0x{:08X}.", static_cast<unsigned long>(remoteCopyHr)));
            return true;
        }

        std::string copiedReadBack;
        if (! ReadFileTextFsIo(remoteIo, remoteCopiedFile, copiedReadBack) || copiedReadBack != state.remoteS3Payload)
        {
            Fail(L"Phase16_RemoteS3FileOps same-plugin copy verification failed.");
            return true;
        }

        const HRESULT noOverwriteHr =
            state.fsRemoteS3->CopyItem(state.remoteS3RenamedFileConn.c_str(), remoteCopiedFileText.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        if (noOverwriteHr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && noOverwriteHr != HRESULT_FROM_WIN32(ERROR_FILE_EXISTS))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps no-overwrite copy returned 0x{:08X}.", static_cast<unsigned long>(noOverwriteHr)));
            return true;
        }

        std::string copiedReadBackAfterConflict;
        if (! ReadFileTextFsIo(remoteIo, remoteCopiedFile, copiedReadBackAfterConflict) || copiedReadBackAfterConflict != state.remoteS3Payload)
        {
            Fail(L"Phase16_RemoteS3FileOps no-overwrite copy mutated the destination.");
            return true;
        }

        const std::filesystem::path largeRemoteSource = caseRoot / L"large-source.bin";
        const std::filesystem::path largeRemoteCopy   = std::filesystem::path(state.remoteS3MoveDirConn) / L"large-copy.bin";
        const std::wstring largeRemoteSourceText      = ToPluginPathText(largeRemoteSource);
        const std::wstring largeRemoteCopyText        = ToPluginPathText(largeRemoteCopy);
        const HRESULT largeCopyHr                     = state.fsRemoteS3->CopyItem(largeRemoteSourceText.c_str(),
                                                                                   largeRemoteCopyText.c_str(),
                                                                                   static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                                                   nullptr,
                                                                                   nullptr,
                                                                                   nullptr);
        if (FAILED(largeCopyHr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps large same-plugin copy failed: 0x{:08X}.", static_cast<unsigned long>(largeCopyHr)));
            return true;
        }

        uint64_t largeCopiedSize = 0;
        if (! GetFileSizeFsIo(remoteIo, largeRemoteCopy, largeCopiedSize) || largeCopiedSize != kRemoteS3LargeFileBytes)
        {
            Fail(L"Phase16_RemoteS3FileOps large same-plugin copy verification failed.");
            return true;
        }

        if (! state.remoteS3AltCaseRootConn.empty())
        {
            const std::filesystem::path altCopiedFile = std::filesystem::path(state.remoteS3AltCaseRootConn) / L"alt-copied-remote.txt";
            const std::wstring altCopiedFileText      = ToPluginPathText(altCopiedFile);
            const HRESULT altCopyHr                   = state.fsRemoteS3->CopyItem(state.remoteS3RenamedFileConn.c_str(),
                                                                                   altCopiedFileText.c_str(),
                                                                                   static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                                                   nullptr,
                                                                                   nullptr,
                                                                                   nullptr);
            if (FAILED(altCopyHr))
            {
                Fail(std::format(L"Phase16_RemoteS3FileOps alternate-connection copy failed: 0x{:08X}.", static_cast<unsigned long>(altCopyHr)));
                return true;
            }

            std::string altCopiedReadBack;
            if (! ReadFileTextFsIo(remoteIo, altCopiedFile, altCopiedReadBack) || altCopiedReadBack != state.remoteS3Payload)
            {
                Fail(L"Phase16_RemoteS3FileOps alternate-connection copy verification failed.");
                return true;
            }
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsLocal,
                                                 {state.remoteS3DirUploadSource},
                                                 std::filesystem::path(state.remoteS3CaseRootConn),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsRemoteS3);
        if (! state.taskA.has_value())
        {
            Fail(L"Phase16_RemoteS3FileOps failed to start local->remote directory copy.");
            return true;
        }

        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        const auto itDirCopy = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (itDirCopy == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itDirCopy->second.hr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps local->remote directory copy failed: 0x{:08X}.", static_cast<unsigned long>(itDirCopy->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> remoteIo = getRemoteIo();
        if (! remoteIo)
        {
            Fail(L"Phase16_RemoteS3FileOps lost remote IFileSystemIO after directory copy.");
            return true;
        }

        const std::filesystem::path remoteUploadedDirFile = std::filesystem::path(state.remoteS3UploadedDirConn) / L"nested" / L"dir-payload.txt";
        std::string remoteDirText;
        if (! ReadFileTextFsIo(remoteIo, remoteUploadedDirFile, remoteDirText) || remoteDirText != state.remoteS3Payload + "-dir")
        {
            Fail(L"Phase16_RemoteS3FileOps remote directory copy verification failed.");
            return true;
        }

        const std::filesystem::path caseRoot(state.remoteS3CaseRootConn);
        const std::filesystem::path copiedDir        = caseRoot / L"copied-folder";
        const std::filesystem::path movedCopiedDir   = caseRoot / L"moved-folder";
        const std::filesystem::path renamedCopiedDir = caseRoot / L"renamed-folder";
        const std::wstring copiedDirText             = ToPluginPathText(copiedDir);
        const std::wstring movedCopiedDirText        = ToPluginPathText(movedCopiedDir);
        const std::wstring renamedCopiedDirText      = ToPluginPathText(renamedCopiedDir);

        const HRESULT dirCopyHr = state.fsRemoteS3->CopyItem(state.remoteS3UploadedDirConn.c_str(),
                                                             copiedDirText.c_str(),
                                                             static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                             nullptr,
                                                             nullptr,
                                                             nullptr);
        if (FAILED(dirCopyHr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps same-plugin directory copy failed: 0x{:08X}.", static_cast<unsigned long>(dirCopyHr)));
            return true;
        }

        const std::filesystem::path copiedDirFile = copiedDir / L"nested" / L"dir-payload.txt";
        std::string copiedDirTextPayload;
        if (! ReadFileTextFsIo(remoteIo, copiedDirFile, copiedDirTextPayload) || copiedDirTextPayload != state.remoteS3Payload + "-dir")
        {
            Fail(L"Phase16_RemoteS3FileOps same-plugin directory copy verification failed.");
            return true;
        }

        const HRESULT dirMoveHr = state.fsRemoteS3->MoveItem(
            copiedDirText.c_str(), movedCopiedDirText.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE), nullptr, nullptr, nullptr);
        if (FAILED(dirMoveHr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps same-plugin directory move failed: 0x{:08X}.", static_cast<unsigned long>(dirMoveHr)));
            return true;
        }

        if (PathExistsFsIo(remoteIo, copiedDir))
        {
            Fail(L"Phase16_RemoteS3FileOps same-plugin directory move left the source directory behind.");
            return true;
        }

        const std::filesystem::path movedDirFile = movedCopiedDir / L"nested" / L"dir-payload.txt";
        std::string movedDirTextPayload;
        if (! ReadFileTextFsIo(remoteIo, movedDirFile, movedDirTextPayload) || movedDirTextPayload != state.remoteS3Payload + "-dir")
        {
            Fail(L"Phase16_RemoteS3FileOps same-plugin directory move verification failed.");
            return true;
        }

        constexpr wchar_t kRenamedDirLeaf[]      = L"renamed-folder";
        const FileSystemRenamePair dirRenameItem = {
            .sizeBytes = sizeof(FileSystemRenamePair), .sourcePath = movedCopiedDirText.c_str(), .newName = kRenamedDirLeaf};
        const HRESULT dirRenameHr =
            state.fsRemoteS3->RenameItems(&dirRenameItem, 1, static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE), nullptr, nullptr, nullptr);
        if (FAILED(dirRenameHr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps same-plugin directory rename failed: 0x{:08X}.", static_cast<unsigned long>(dirRenameHr)));
            return true;
        }

        if (PathExistsFsIo(remoteIo, movedCopiedDir) || ! PathExistsFsIo(remoteIo, renamedCopiedDir))
        {
            Fail(L"Phase16_RemoteS3FileOps same-plugin directory rename left unexpected paths.");
            return true;
        }

        const std::filesystem::path renamedDirFile = renamedCopiedDir / L"nested" / L"dir-payload.txt";
        std::string renamedDirTextPayload;
        if (! ReadFileTextFsIo(remoteIo, renamedDirFile, renamedDirTextPayload) || renamedDirTextPayload != state.remoteS3Payload + "-dir")
        {
            Fail(L"Phase16_RemoteS3FileOps same-plugin directory rename verification failed.");
            return true;
        }

        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsRemoteS3,
                                                 {std::filesystem::path(state.remoteS3UploadedDirConn)},
                                                 state.remoteS3DirMoveDownloadDir,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskB.has_value())
        {
            Fail(L"Phase16_RemoteS3FileOps failed to start remote->local directory move.");
            return true;
        }

        state.stepState = 6;
        return false;
    }

    if (state.stepState == 6)
    {
        const auto itDirMove = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (itDirMove == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itDirMove->second.hr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps remote->local directory move failed: 0x{:08X}.", static_cast<unsigned long>(itDirMove->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> remoteIo = getRemoteIo();
        wil::com_ptr<IFileSystemIO> localIo  = getLocalIo();
        if (! remoteIo || ! localIo)
        {
            Fail(L"Phase16_RemoteS3FileOps lost IFileSystemIO after directory move.");
            return true;
        }

        if (PathExistsFsIo(remoteIo, std::filesystem::path(state.remoteS3UploadedDirConn)))
        {
            Fail(L"Phase16_RemoteS3FileOps remote directory move left the source directory behind.");
            return true;
        }

        std::string localDirText;
        if (! ReadFileTextFsIo(localIo, state.remoteS3DirMovedFile, localDirText) || localDirText != state.remoteS3Payload + "-dir")
        {
            Fail(L"Phase16_RemoteS3FileOps local directory move verification failed.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsRemoteS3,
                                                 {std::filesystem::path(state.remoteS3RenamedFileConn)},
                                                 {},
                                                 FILESYSTEM_FLAG_NONE,
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"Phase16_RemoteS3FileOps failed to start remote delete.");
            return true;
        }

        state.stepState = 7;
        return false;
    }

    if (state.stepState == 7)
    {
        const auto itDelete = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (itDelete == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itDelete->second.hr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps remote delete failed: 0x{:08X}.", static_cast<unsigned long>(itDelete->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> remoteIo = getRemoteIo();
        if (! remoteIo)
        {
            Fail(L"Phase16_RemoteS3FileOps lost remote IFileSystemIO after delete.");
            return true;
        }

        if (PathExistsFsIo(remoteIo, std::filesystem::path(state.remoteS3RenamedFileConn)))
        {
            Fail(L"Phase16_RemoteS3FileOps delete left the file in place.");
            return true;
        }

        const HRESULT cleanupHr = state.fsRemoteS3->DeleteItem(state.remoteS3CaseRootConn.c_str(),
                                                               static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                               nullptr,
                                                               nullptr,
                                                               nullptr);
        if (FAILED(cleanupHr))
        {
            Fail(std::format(L"Phase16_RemoteS3FileOps failed to clean the remote sandbox: 0x{:08X}.", static_cast<unsigned long>(cleanupHr)));
            return true;
        }

        ResetRemoteS3State(state);
        RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
        NextStep(state, SelfTestState::Step::Phase16_RemoteOneDrivePersonalSecret);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase16_RemoteOneDrivePersonalSecret:
{
    const PhaseCheckResult outcome =
        CheckRemoteConnectionSecret(L"OneDrive Personal", kSelfTestEnvConnOneDrivePersonal, kSelfTestDefaultConnOneDrivePersonal, kPluginIdOneDrivePersonal);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteOneDrivePersonalSandbox);
    return false;
}
case SelfTestState::Step::Phase16_RemoteOneDrivePersonalSandbox:
{
    const PhaseCheckResult outcome =
        CheckRemoteConnectionSandbox(L"OneDrive Personal", kSelfTestEnvConnOneDrivePersonal, kSelfTestDefaultConnOneDrivePersonal, kPluginIdOneDrivePersonal);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteOneDrivePersonalFileOps);
    return false;
}
case SelfTestState::Step::Phase16_RemoteOneDrivePersonalFileOps:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(std::format(L"Phase16_RemoteOneDrivePersonalFileOps timed out (stepState={} taskA={} taskB={} taskC={} caseRoot='{}').",
                         state.stepState,
                         state.taskA.value_or(0ull),
                         state.taskB.value_or(0ull),
                         state.taskC.value_or(0ull),
                         state.remoteOneDrivePersonalCaseRootConn));
        return true;
    }

    auto getRemoteIo = [&]() noexcept -> wil::com_ptr<IFileSystemIO>
    {
        wil::com_ptr<IFileSystemIO> io;
        if (state.fsRemoteOneDrivePersonal)
        {
            static_cast<void>(state.fsRemoteOneDrivePersonal->QueryInterface(IID_PPV_ARGS(io.put())));
        }
        return io;
    };

    auto getLocalIo = [&]() noexcept -> wil::com_ptr<IFileSystemIO>
    {
        wil::com_ptr<IFileSystemIO> io;
        if (state.fsLocal)
        {
            static_cast<void>(state.fsLocal->QueryInterface(IID_PPV_ARGS(io.put())));
        }
        return io;
    };

    if (state.stepState == 0)
    {
        const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(
            L"OneDrive Personal", kSelfTestEnvConnOneDrivePersonal, kSelfTestDefaultConnOneDrivePersonal, kPluginIdOneDrivePersonal);
        if (secretOutcome.status == SelfTest::SelfTestCaseResult::Status::failed)
        {
            Fail(secretOutcome.reason);
            return true;
        }

        if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
        {
            RecordCurrentPhase(state, secretOutcome.status, secretOutcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteOneDriveBusinessSecret);
            return false;
        }

        const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(
            L"OneDrive Personal", kSelfTestEnvConnOneDrivePersonal, kSelfTestDefaultConnOneDrivePersonal, kPluginIdOneDrivePersonal);
        if (sandboxOutcome.status == SelfTest::SelfTestCaseResult::Status::failed)
        {
            Fail(sandboxOutcome.reason);
            return true;
        }

        if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
        {
            RecordCurrentPhase(state, sandboxOutcome.status, sandboxOutcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteOneDriveBusinessSecret);
            return false;
        }

        CleanupRemoteOneDrivePersonalCase(state);
        state.taskA.reset();
        state.taskB.reset();
        state.taskC.reset();

        const ResolvedRemoteProfile resolvedProfile =
            ResolveRemoteConnectionProfile(kSelfTestEnvConnOneDrivePersonal, kSelfTestDefaultConnOneDrivePersonal, kPluginIdOneDrivePersonal);
        const std::wstring profileName                     = resolvedProfile.profileName;
        const Common::Settings::ConnectionProfile* profile = resolvedProfile.profile;
        if (! profile)
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps could not resolve the connection profile after validation.");
            return true;
        }

        FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
        const HRESULT enableHr                 = pluginManager.EnablePlugin(kPluginIdOneDrivePersonal, g_settings);
        if (FAILED(enableHr))
        {
            Fail(std::format(L"Phase16_RemoteOneDrivePersonalFileOps failed to enable plugin (hr=0x{:08X}).", static_cast<unsigned long>(enableHr)));
            return true;
        }

        const FileSystemPluginManager::PluginEntry* entry = FindLoadedPluginEntry(kPluginIdOneDrivePersonal);
        if (! entry || ! entry->fileSystem)
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps could not locate the loaded plugin entry.");
            return true;
        }

        wil::com_ptr<IFileSystemIO> remoteIo;
        const HRESULT remoteIoHr = entry->fileSystem->QueryInterface(IID_PPV_ARGS(remoteIo.put()));
        if (FAILED(remoteIoHr) || ! remoteIo)
        {
            Fail(std::format(L"Phase16_RemoteOneDrivePersonalFileOps missing IFileSystemIO (hr=0x{:08X}).", static_cast<unsigned long>(remoteIoHr)));
            return true;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> remoteDirOps;
        const HRESULT remoteDirHr = entry->fileSystem->QueryInterface(IID_PPV_ARGS(remoteDirOps.put()));
        if (FAILED(remoteDirHr) || ! remoteDirOps)
        {
            Fail(std::format(L"Phase16_RemoteOneDrivePersonalFileOps missing IFileSystemDirectoryOperations (hr=0x{:08X}).",
                             static_cast<unsigned long>(remoteDirHr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> localIo = getLocalIo();
        if (! localIo)
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps missing local IFileSystemIO.");
            return true;
        }

        const std::wstring guid = NewGuidString();
        if (guid.empty())
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps failed to generate a GUID for the sandbox.");
            return true;
        }

        const std::wstring caseRootPlugin = JoinPluginPathForSelfTest(profile->initialPath, std::format(L"fileops-selftest-onedrivep-{}", guid));
        const std::wstring caseRootConn   = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
        if (caseRootPlugin.empty() || caseRootConn.empty())
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps failed to build sandbox paths.");
            return true;
        }

        const std::filesystem::path caseRoot(caseRootConn);
        const std::filesystem::path uploadDir = caseRoot / L"upload";
        const std::filesystem::path moveDir   = caseRoot / L"move";
        const std::filesystem::path seedFile  = caseRoot / L"seed-remote.txt";

        const std::filesystem::path workspace    = state.tempRoot / L"onedrive-personal-fileops";
        const std::filesystem::path downloadDir  = workspace / L"download";
        const std::filesystem::path uploadSource = workspace / L"upload-source.txt";
        if (! RecreateEmptyDirectory(workspace) || ! RecreateEmptyDirectory(downloadDir))
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps failed to prepare the local workspace.");
            return true;
        }

        const std::string payload = std::format("onedrive-personal-fileops-{}-roundtrip", static_cast<unsigned long long>(GetTickCount64()));
        if (! EnsureDirectoryExistsFsOps(remoteDirOps, caseRoot) || ! EnsureDirectoryExistsFsOps(remoteDirOps, uploadDir) ||
            ! EnsureDirectoryExistsFsOps(remoteDirOps, moveDir))
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps failed to prepare the remote sandbox.");
            return true;
        }

        if (! WriteFileTextFsIo(localIo, uploadSource, payload))
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps failed to create the local upload file.");
            return true;
        }

        if (! WriteFileTextFsIo(remoteIo, seedFile, payload))
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps failed to seed the remote probe file.");
            return true;
        }

        std::string seedReadBack;
        if (! ReadFileTextFsIo(remoteIo, seedFile, seedReadBack) || seedReadBack != payload)
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps remote CreateFileReader readback failed.");
            return true;
        }

        state.fsRemoteOneDrivePersonal               = entry->fileSystem;
        state.remoteOneDrivePersonalProfileName      = profileName;
        state.remoteOneDrivePersonalCaseRootConn     = caseRootConn;
        state.remoteOneDrivePersonalUploadDirConn    = ToPluginPathText(uploadDir);
        state.remoteOneDrivePersonalMoveDirConn      = ToPluginPathText(moveDir);
        state.remoteOneDrivePersonalSeedFileConn     = ToPluginPathText(seedFile);
        state.remoteOneDrivePersonalUploadedFileConn = ToPluginPathText(uploadDir / uploadSource.filename());
        state.remoteOneDrivePersonalMovedFileConn    = ToPluginPathText(moveDir / uploadSource.filename());
        state.remoteOneDrivePersonalDisplayPath      = MakePluginDisplayPath(entry->shortId, state.remoteOneDrivePersonalCaseRootConn);
        state.remoteOneDrivePersonalWorkspace        = workspace;
        state.remoteOneDrivePersonalUploadSource     = uploadSource;
        state.remoteOneDrivePersonalDownloadDir      = downloadDir;
        state.remoteOneDrivePersonalDownloadedFile   = downloadDir / uploadSource.filename();
        state.remoteOneDrivePersonalPayload          = payload;

        state.folderWindow->SetFolderPath(FolderWindow::Pane::Left, state.remoteOneDrivePersonalDisplayPath);
        state.folderWindow->SetFolderPath(FolderWindow::Pane::Right, state.remoteOneDrivePersonalWorkspace);
        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (! CurrentPanePathEquals(state.folderWindow, FolderWindow::Pane::Left, state.remoteOneDrivePersonalDisplayPath) ||
            ! CurrentPanePathEquals(state.folderWindow, FolderWindow::Pane::Right, state.remoteOneDrivePersonalWorkspace))
        {
            return false;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsLocal,
                                                 {state.remoteOneDrivePersonalUploadSource},
                                                 std::filesystem::path(state.remoteOneDrivePersonalUploadDirConn),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsRemoteOneDrivePersonal);
        if (! state.taskA.has_value())
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps failed to start local->remote copy.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto itCopy = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (itCopy == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itCopy->second.hr))
        {
            Fail(std::format(L"Phase16_RemoteOneDrivePersonalFileOps local->remote copy failed: 0x{:08X}.", static_cast<unsigned long>(itCopy->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> remoteIo = getRemoteIo();
        if (! remoteIo)
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps lost remote IFileSystemIO after upload.");
            return true;
        }

        const std::filesystem::path uploadedFile(state.remoteOneDrivePersonalUploadedFileConn);
        if (! PathExistsFsIo(remoteIo, uploadedFile))
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps upload did not materialize the remote file.");
            return true;
        }

        std::string uploadReadBack;
        if (! ReadFileTextFsIo(remoteIo, uploadedFile, uploadReadBack) || uploadReadBack != state.remoteOneDrivePersonalPayload)
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps remote upload verification failed.");
            return true;
        }

        state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsRemoteOneDrivePersonal,
                                                 {uploadedFile},
                                                 state.remoteOneDrivePersonalDownloadDir,
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsLocal);
        if (! state.taskB.has_value())
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps failed to start remote->local copy.");
            return true;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        const auto itCopyBack = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (itCopyBack == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itCopyBack->second.hr))
        {
            Fail(std::format(L"Phase16_RemoteOneDrivePersonalFileOps remote->local copy failed: 0x{:08X}.", static_cast<unsigned long>(itCopyBack->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> localIo = getLocalIo();
        if (! localIo)
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps lost local IFileSystemIO after download.");
            return true;
        }

        std::string downloadedText;
        if (! ReadFileTextFsIo(localIo, state.remoteOneDrivePersonalDownloadedFile, downloadedText) || downloadedText != state.remoteOneDrivePersonalPayload)
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps downloaded file contents did not match the upload.");
            return true;
        }

        state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Left,
                                                 state.fsRemoteOneDrivePersonal,
                                                 {std::filesystem::path(state.remoteOneDrivePersonalUploadedFileConn)},
                                                 std::filesystem::path(state.remoteOneDrivePersonalMoveDirConn),
                                                 static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsRemoteOneDrivePersonal);
        if (! state.taskC.has_value())
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps failed to start same-drive move.");
            return true;
        }

        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        const auto itMove = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
        if (itMove == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itMove->second.hr))
        {
            Fail(std::format(L"Phase16_RemoteOneDrivePersonalFileOps same-drive move failed: 0x{:08X}.", static_cast<unsigned long>(itMove->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> remoteIo = getRemoteIo();
        if (! remoteIo)
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps lost remote IFileSystemIO after move.");
            return true;
        }

        if (PathExistsFsIo(remoteIo, std::filesystem::path(state.remoteOneDrivePersonalUploadedFileConn)))
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps move left the source file behind.");
            return true;
        }

        if (! PathExistsFsIo(remoteIo, std::filesystem::path(state.remoteOneDrivePersonalMovedFileConn)))
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps move did not materialize the destination file.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsRemoteOneDrivePersonal,
                                                 {std::filesystem::path(state.remoteOneDrivePersonalMovedFileConn)},
                                                 {},
                                                 FILESYSTEM_FLAG_NONE,
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps failed to start remote delete.");
            return true;
        }

        state.stepState = 5;
        return false;
    }

    if (state.stepState == 5)
    {
        const auto itDelete = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (itDelete == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(itDelete->second.hr))
        {
            Fail(std::format(L"Phase16_RemoteOneDrivePersonalFileOps remote delete failed: 0x{:08X}.", static_cast<unsigned long>(itDelete->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> remoteIo = getRemoteIo();
        if (! remoteIo)
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps lost remote IFileSystemIO after delete.");
            return true;
        }

        if (PathExistsFsIo(remoteIo, std::filesystem::path(state.remoteOneDrivePersonalMovedFileConn)))
        {
            Fail(L"Phase16_RemoteOneDrivePersonalFileOps delete left the file in place.");
            return true;
        }

        const HRESULT cleanupHr =
            state.fsRemoteOneDrivePersonal->DeleteItem(state.remoteOneDrivePersonalCaseRootConn.c_str(),
                                                       static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                       nullptr,
                                                       nullptr,
                                                       nullptr);
        if (FAILED(cleanupHr))
        {
            Fail(std::format(L"Phase16_RemoteOneDrivePersonalFileOps failed to clean the remote sandbox: 0x{:08X}.", static_cast<unsigned long>(cleanupHr)));
            return true;
        }

        ResetRemoteOneDrivePersonalState(state);
        RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
        NextStep(state, SelfTestState::Step::Phase16_RemoteOneDriveBusinessSecret);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase16_RemoteOneDriveBusinessSecret:
{
    const PhaseCheckResult outcome =
        CheckRemoteConnectionSecret(L"OneDrive Business", kSelfTestEnvConnOneDriveBusiness, kSelfTestDefaultConnOneDriveBusiness, kPluginIdOneDriveBusiness);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteOneDriveBusinessSandbox);
    return false;
}
case SelfTestState::Step::Phase16_RemoteOneDriveBusinessSandbox:
{
    const PhaseCheckResult outcome =
        CheckRemoteConnectionSandbox(L"OneDrive Business", kSelfTestEnvConnOneDriveBusiness, kSelfTestDefaultConnOneDriveBusiness, kPluginIdOneDriveBusiness);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteSharePointSecret);
    return false;
}
case SelfTestState::Step::Phase16_RemoteSharePointSecret:
{
    const PhaseCheckResult outcome =
        CheckRemoteConnectionSecret(L"SharePoint", kSelfTestEnvConnSharePoint, kSelfTestDefaultConnSharePoint, kPluginIdSharePoint);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Phase16_RemoteSharePointSandbox);
    return false;
}
case SelfTestState::Step::Phase16_RemoteSharePointSandbox:
{
    const PhaseCheckResult outcome =
        CheckRemoteConnectionSandbox(L"SharePoint", kSelfTestEnvConnSharePoint, kSelfTestDefaultConnSharePoint, kPluginIdSharePoint);
    if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        Fail(outcome.reason);
        return true;
    }

    RecordCurrentPhase(state, outcome.status, outcome.reason);
    NextStep(state, SelfTestState::Step::Cleanup_RestorePluginConfig);
    return false;
}
case SelfTestState::Step::Cleanup_RestorePluginConfig:
{
    PerformCleanup(state);
    RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);

    state.step = SelfTestState::Step::Done;
    state.running.store(false, std::memory_order_release);
    state.done.store(true, std::memory_order_release);
    Debug::Info(L"FileOpsSelfTest: {}", state.failed.load(std::memory_order_acquire) ? L"FAIL" : L"PASS");
    return true;
}
case SelfTestState::Step::Done: return true;
case SelfTestState::Step::Failed:
{
    NextStep(state, SelfTestState::Step::Cleanup_RestorePluginConfig);
    return false;
}
case SelfTestState::Step::Idle:
default: break;
    }
