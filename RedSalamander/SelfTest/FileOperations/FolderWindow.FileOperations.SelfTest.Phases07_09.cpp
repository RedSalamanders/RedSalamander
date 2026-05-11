case SelfTestState::Step::Phase7_WatcherChurn:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 60'000ull))
    {
        Fail(L"Phase7_WatcherChurn timed out.");
        return true;
    }

    if (state.stepState == 0)
    {
        state.watchDir = state.tempRoot / L"watch";
        if (! RecreateEmptyDirectory(state.watchDir))
        {
            Fail(L"Failed to reset watch directory.");
            return true;
        }

        state.directoryWatch.reset();
        const HRESULT hrQI = state.fsLocal->QueryInterface(__uuidof(IFileSystemDirectoryWatch), state.directoryWatch.put_void());
        if (FAILED(hrQI) || ! state.directoryWatch)
        {
            Fail(L"Local file system plugin does not expose IFileSystemDirectoryWatch.");
            return true;
        }

        state.directoryWatchCallback = std::make_unique<WatchCallback>();
        const HRESULT hrWatch        = state.directoryWatch->WatchDirectory(state.watchDir.c_str(), state.directoryWatchCallback.get(), nullptr);
        if (FAILED(hrWatch))
        {
            Fail(std::format(L"WatchDirectory failed: 0x{:08X}", static_cast<unsigned long>(hrWatch)));
            return true;
        }

        // Churn: create/rename/delete a bunch of files quickly.
        for (int i = 0; i < 200; ++i)
        {
            const std::filesystem::path p1 = state.watchDir / std::format(L"churn_{:04}.tmp", i);
            const std::filesystem::path p2 = state.watchDir / std::format(L"churn_{:04}.renamed", i);

            static_cast<void>(WriteTestFile(p1, 32));
            static_cast<void>(MoveFileExW(p1.c_str(), p2.c_str(), MOVEFILE_REPLACE_EXISTING));
            static_cast<void>(DeleteFileW(p2.c_str()));
        }

        state.markerTick = nowTick;
        state.stepState  = 1;
        return false;
    }

    auto* cb                     = static_cast<WatchCallback*>(state.directoryWatchCallback.get());
    const uint64_t callbackCount = cb ? cb->callbackCount.load(std::memory_order_relaxed) : 0ull;

    if (state.stepState == 1)
    {
        if (nowTick >= state.markerTick && (nowTick - state.markerTick) < 1000ull)
        {
            return false;
        }

        static_cast<void>(state.directoryWatch->UnwatchDirectory(state.watchDir.c_str()));
        state.watchCounter = cb ? cb->callbackCount.load(std::memory_order_relaxed) : 0ull;

        // After UnwatchDirectory returns, no further callbacks are allowed for that registration.
        for (int i = 0; i < 100; ++i)
        {
            const std::filesystem::path p1 = state.watchDir / std::format(L"postunwatch_{:04}.tmp", i);
            const std::filesystem::path p2 = state.watchDir / std::format(L"postunwatch_{:04}.renamed", i);

            static_cast<void>(WriteTestFile(p1, 32));
            static_cast<void>(MoveFileExW(p1.c_str(), p2.c_str(), MOVEFILE_REPLACE_EXISTING));
            static_cast<void>(DeleteFileW(p2.c_str()));
        }

        state.markerTick = nowTick;
        state.stepState  = 2;
        return false;
    }

    if (nowTick >= state.markerTick && (nowTick - state.markerTick) < 500ull)
    {
        return false;
    }

    if (callbackCount == 0)
    {
        Fail(L"Watcher churn did not produce any callbacks.");
        return true;
    }

    const uint64_t badSizeCount = cb ? cb->badNotificationSizeCount.load(std::memory_order_relaxed) : 0ull;
    if (badSizeCount != 0)
    {
        Fail(std::format(L"Watcher churn produced notifications with invalid sizeBytes (count={}).", badSizeCount));
        return true;
    }

    if (callbackCount != state.watchCounter)
    {
        Fail(std::format(L"Watcher received callbacks after UnwatchDirectory returned (before={} after={}).", state.watchCounter, callbackCount));
        return true;
    }

    wil::com_ptr<IFileSystemDirectoryWatch> dummyWatch;
    const HRESULT hrDummyWatch = state.fsDummy ? state.fsDummy->QueryInterface(__uuidof(IFileSystemDirectoryWatch), dummyWatch.put_void()) : E_NOINTERFACE;
    if (FAILED(hrDummyWatch) || ! dummyWatch)
    {
        Fail(std::format(L"Dummy filesystem does not expose IFileSystemDirectoryWatch (hr=0x{:08X}).", static_cast<unsigned long>(hrDummyWatch)));
        return true;
    }

    wil::com_ptr<IFileSystemIO> dummyIo;
    const HRESULT hrDummyIo = state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()));
    if (FAILED(hrDummyIo) || ! dummyIo)
    {
        Fail(std::format(L"Dummy filesystem does not expose IFileSystemIO for watch validation (hr=0x{:08X}).", static_cast<unsigned long>(hrDummyIo)));
        return true;
    }

    const std::wstring dummyWatchDir = std::format(L"/watch-selftest-{}", GetTickCount64());
    const std::wstring fileA         = std::format(L"{}/original.txt", dummyWatchDir);
    const std::wstring fileB         = std::format(L"{}/renamed.txt", dummyWatchDir);
    const std::wstring fileC         = std::format(L"{}/final.txt", dummyWatchDir);
    const auto cleanupDummyWatch     = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(dummyWatch->UnwatchDirectory(dummyWatchDir.c_str()));
        static_cast<void>(state.fsDummy->DeleteItem(
            dummyWatchDir.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, nullptr, nullptr));
    });

    if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyWatchDir))
    {
        Fail(L"Dummy watcher regression: failed to create watch directory.");
        return true;
    }

    {
        wil::com_ptr<IFileWriter> writer;
        const HRESULT hrWriter = dummyIo->CreateFileWriter(fileA.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
        if (FAILED(hrWriter) || ! writer)
        {
            Fail(std::format(L"Dummy watcher regression: CreateFileWriter failed (hr=0x{:08X}).", static_cast<unsigned long>(hrWriter)));
            return true;
        }

        static constexpr char kDummyText[] = "watch";
        unsigned long written              = 0;
        const HRESULT hrWrite              = writer->Write(kDummyText, static_cast<unsigned long>(sizeof(kDummyText) - 1u), &written);
        if (FAILED(hrWrite) || written != (sizeof(kDummyText) - 1u) || FAILED(writer->Commit()))
        {
            Fail(std::format(L"Dummy watcher regression: failed to seed file (write hr=0x{:08X} written={}).", static_cast<unsigned long>(hrWrite), written));
            return true;
        }
    }

    DummyReentrantWatchCallback callbackA{};
    callbackA.watch       = dummyWatch.get();
    callbackA.watchedPath = dummyWatchDir;

    const HRESULT hrWatchA         = dummyWatch->WatchDirectory(dummyWatchDir.c_str(), &callbackA, nullptr);
    const HRESULT hrDuplicateWatch = dummyWatch->WatchDirectory(dummyWatchDir.c_str(), &callbackA, nullptr);
    if (FAILED(hrWatchA))
    {
        Fail(std::format(L"Dummy watcher regression: WatchDirectory failed (hr=0x{:08X}).", static_cast<unsigned long>(hrWatchA)));
        return true;
    }
    if (hrDuplicateWatch != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
    {
        Fail(std::format(L"Dummy watcher regression: duplicate WatchDirectory should return ERROR_ALREADY_EXISTS, got 0x{:08X}.",
                         static_cast<unsigned long>(hrDuplicateWatch)));
        return true;
    }

    const HRESULT hrRename1 = state.fsDummy->RenameItem(fileA.c_str(), fileB.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    if (FAILED(hrRename1))
    {
        Fail(std::format(L"Dummy watcher regression: first RenameItem failed (hr=0x{:08X}).", static_cast<unsigned long>(hrRename1)));
        return true;
    }

    if (callbackA.callbackCount.load(std::memory_order_acquire) != 1u)
    {
        Fail(std::format(L"Dummy watcher regression: reentrant watcher expected exactly one callback, got {}.",
                         callbackA.callbackCount.load(std::memory_order_acquire)));
        return true;
    }
    if (FAILED(callbackA.unwatchHr.load(std::memory_order_acquire)))
    {
        Fail(std::format(L"Dummy watcher regression: reentrant UnwatchDirectory failed (hr=0x{:08X}).",
                         static_cast<unsigned long>(callbackA.unwatchHr.load(std::memory_order_acquire))));
        return true;
    }
    if (callbackA.renamedOldCount.load(std::memory_order_acquire) == 0u || callbackA.renamedNewCount.load(std::memory_order_acquire) == 0u)
    {
        Fail(L"Dummy watcher regression: reentrant watcher did not receive rename old/new notifications.");
        return true;
    }

    const uint64_t callbackACountBefore = callbackA.callbackCount.load(std::memory_order_acquire);

    const HRESULT hrRename2 = state.fsDummy->RenameItem(fileB.c_str(), fileC.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    if (FAILED(hrRename2))
    {
        Fail(std::format(L"Dummy watcher regression: second RenameItem failed (hr=0x{:08X}).", static_cast<unsigned long>(hrRename2)));
        return true;
    }

    if (callbackA.callbackCount.load(std::memory_order_acquire) != callbackACountBefore)
    {
        Fail(L"Dummy watcher regression: reentrant watcher received callbacks after UnwatchDirectory returned.");
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase7_CacheBorrowNoWatchInvalidation);
    return false;
}
case SelfTestState::Step::Phase7_CacheBorrowNoWatchInvalidation:
{
    const ULONGLONG nowTick                  = GetTickCount64();
    const std::filesystem::path phaseRoot    = state.tempRoot / L"phase7-cache-nowatch";
    const std::filesystem::path sourceParent = phaseRoot / L"source";
    const std::filesystem::path sourceChild  = sourceParent / L"child";
    const std::filesystem::path sourceFile   = sourceChild / L"payload.txt";
    const std::filesystem::path movedParent  = phaseRoot / L"renamed";
    const std::filesystem::path movedChild   = movedParent / L"child";
    DirectoryInfoCache& cache                = DirectoryInfoCache::GetInstance();

    if (HasTimedOut(state, nowTick, 10'000ull))
    {
        Fail(L"Phase7_CacheBorrowNoWatchInvalidation timed out.");
        return true;
    }

    if (state.stepState == 0)
    {
        cache.RegisterProvider(state.fsLocal.get(), kPluginIdLocal, L"file", {});

        if (! RecreateEmptyDirectory(phaseRoot))
        {
            Fail(L"Phase7_CacheBorrowNoWatchInvalidation failed to recreate root.");
            return true;
        }

        std::error_code ec;
        std::filesystem::create_directories(sourceChild, ec);
        if (ec || ! WriteTestFile(sourceFile, 64))
        {
            Fail(std::format(L"Phase7_CacheBorrowNoWatchInvalidation failed to seed source tree (ec={}).", ec.value()));
            return true;
        }

        auto initialBorrow = cache.BorrowDirectoryInfo(state.fsLocal.get(), sourceChild, DirectoryInfoCache::BorrowMode::AllowEnumerate);
        if (FAILED(initialBorrow.Status()) || ! initialBorrow.Get())
        {
            Fail(std::format(L"Phase7_CacheBorrowNoWatchInvalidation failed to enumerate cached folder (hr=0x{:08X}).",
                             static_cast<unsigned long>(initialBorrow.Status())));
            return true;
        }

        if (cache.IsFolderWatched(state.fsLocal.get(), sourceChild))
        {
            Fail(L"Phase7_CacheBorrowNoWatchInvalidation unexpectedly started a watcher for a cache-only folder.");
            return true;
        }

        if (! MoveFileExW(sourceParent.c_str(), movedParent.c_str(), MOVEFILE_REPLACE_EXISTING))
        {
            Fail(std::format(L"Phase7_CacheBorrowNoWatchInvalidation failed to rename ancestor directory (lastError={}).", GetLastError()));
            return true;
        }

        cache.NotifyPathMoved(state.fsLocal.get(), sourceParent, movedParent);

        auto staleBorrow = cache.BorrowDirectoryInfo(state.fsLocal.get(), sourceChild, DirectoryInfoCache::BorrowMode::AllowEnumerate);
        if (SUCCEEDED(staleBorrow.Status()) || staleBorrow.Get())
        {
            Fail(std::format(L"Phase7_CacheBorrowNoWatchInvalidation returned a clean snapshot for the old path after routed move (hr=0x{:08X}).",
                             static_cast<unsigned long>(staleBorrow.Status())));
            return true;
        }

        auto movedBorrow = cache.BorrowDirectoryInfo(state.fsLocal.get(), movedChild, DirectoryInfoCache::BorrowMode::AllowEnumerate);
        if (FAILED(movedBorrow.Status()) || ! movedBorrow.Get())
        {
            Fail(std::format(L"Phase7_CacheBorrowNoWatchInvalidation failed to enumerate moved folder (hr=0x{:08X}).",
                             static_cast<unsigned long>(movedBorrow.Status())));
            return true;
        }

        if (cache.IsFolderWatched(state.fsLocal.get(), movedChild))
        {
            Fail(L"Phase7_CacheBorrowNoWatchInvalidation unexpectedly watched the moved folder after a direct borrow.");
            return true;
        }

        std::filesystem::remove_all(movedParent, ec);
        if (ec)
        {
            Fail(std::format(L"Phase7_CacheBorrowNoWatchInvalidation failed to delete moved ancestor (ec={}).", ec.value()));
            return true;
        }

        cache.NotifyPathDeleted(state.fsLocal.get(), movedParent);

        auto deletedBorrow = cache.BorrowDirectoryInfo(state.fsLocal.get(), movedChild, DirectoryInfoCache::BorrowMode::AllowEnumerate);
        if (SUCCEEDED(deletedBorrow.Status()) || deletedBorrow.Get())
        {
            Fail(std::format(L"Phase7_CacheBorrowNoWatchInvalidation returned a clean snapshot after routed delete (hr=0x{:08X}).",
                             static_cast<unsigned long>(deletedBorrow.Status())));
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase7_CrossPaneVisibleRefreshLocal);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase7_CrossPaneVisibleRefreshLocal:
{
    const ULONGLONG nowTick                  = GetTickCount64();
    const std::filesystem::path phaseRoot    = state.tempRoot / L"phase7-visible-local";
    const std::filesystem::path steadyFile   = phaseRoot / L"steady.txt";
    const std::filesystem::path renameSource = phaseRoot / L"rename-me.txt";
    const std::filesystem::path renameTarget = phaseRoot / L"renamed.txt";
    const std::filesystem::path newFile      = phaseRoot / L"new.txt";

    if (HasTimedOut(state, nowTick, 45'000ull))
    {
        const auto leftPath  = state.folderWindow ? state.folderWindow->GetCurrentPath(FolderWindow::Pane::Left) : std::nullopt;
        const auto rightPath = state.folderWindow ? state.folderWindow->GetCurrentPath(FolderWindow::Pane::Right) : std::nullopt;
        std::error_code sourceEc;
        std::error_code targetEc;
        const bool sourceExists = std::filesystem::exists(renameSource, sourceEc);
        const bool targetExists = std::filesystem::exists(renameTarget, targetEc);
        Fail(std::format(
            L"Phase7_CrossPaneVisibleRefreshLocal timed out. stepState={} leftPath='{}' rightPath='{}' leftSteady={} rightSteady={} leftRenameSrc={} "
            L"rightRenameSrc={} leftRenameDst={} rightRenameDst={} leftNew={} rightNew={} sourceExists={} targetExists={} sourceEc={} targetEc={}",
            state.stepState,
            leftPath.has_value() ? leftPath->wstring() : L"<null>",
            rightPath.has_value() ? rightPath->wstring() : L"<null>",
            state.folderWindow && state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"steady.txt"),
            state.folderWindow && state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"steady.txt"),
            state.folderWindow && state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-me.txt"),
            state.folderWindow && state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"rename-me.txt"),
            state.folderWindow && state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"renamed.txt"),
            state.folderWindow && state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"renamed.txt"),
            state.folderWindow && state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"new.txt"),
            state.folderWindow && state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"new.txt"),
            sourceExists,
            targetExists,
            sourceEc.value(),
            targetEc.value()));
        return true;
    }

    if (state.stepState == 0)
    {
        const HRESULT leftHr  = state.folderWindow->SetFileSystemPluginForPane(FolderWindow::Pane::Left, kPluginIdLocal);
        const HRESULT rightHr = state.folderWindow->SetFileSystemPluginForPane(FolderWindow::Pane::Right, kPluginIdLocal);
        if (FAILED(leftHr) || FAILED(rightHr))
        {
            Fail(std::format(L"Phase7_CrossPaneVisibleRefreshLocal failed to switch panes to local plugin (left=0x{:08X} right=0x{:08X}).",
                             static_cast<unsigned long>(leftHr),
                             static_cast<unsigned long>(rightHr)));
            return true;
        }

        if (! RecreateEmptyDirectory(phaseRoot) || ! WriteTestFile(steadyFile, 64) || ! WriteTestFile(renameSource, 64))
        {
            Fail(L"Phase7_CrossPaneVisibleRefreshLocal failed to seed test directory.");
            return true;
        }

        state.folderWindow->SetFolderPath(FolderWindow::Pane::Left, phaseRoot);
        state.folderWindow->SetFolderPath(FolderWindow::Pane::Right, phaseRoot);
        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (! CurrentPanePathEquals(state.folderWindow, FolderWindow::Pane::Left, phaseRoot) ||
            ! CurrentPanePathEquals(state.folderWindow, FolderWindow::Pane::Right, phaseRoot) ||
            ! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"steady.txt") ||
            ! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"steady.txt") ||
            ! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-me.txt") ||
            ! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"rename-me.txt"))
        {
            return false;
        }

        if (! state.folderWindow->DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"steady.txt"))
        {
            return false;
        }

        if (! WriteTestFile(newFile, 37))
        {
            Fail(L"Phase7_CrossPaneVisibleRefreshLocal failed to create watched file.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"new.txt") ||
            ! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"new.txt"))
        {
            return false;
        }

        if (! MoveFileExW(renameSource.c_str(), renameTarget.c_str(), MOVEFILE_REPLACE_EXISTING))
        {
            Fail(std::format(L"Phase7_CrossPaneVisibleRefreshLocal failed to rename watched file (lastError={}).", GetLastError()));
            return true;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        if (! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"renamed.txt") ||
            ! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"renamed.txt") ||
            state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-me.txt") ||
            state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"rename-me.txt"))
        {
            return false;
        }

        if (! DeleteFileW(newFile.c_str()))
        {
            Fail(std::format(L"Phase7_CrossPaneVisibleRefreshLocal failed to delete watched file (lastError={}).", GetLastError()));
            return true;
        }

        state.stepState = 4;
        return false;
    }

    if (state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"new.txt") ||
        state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"new.txt"))
    {
        return false;
    }

    NextStep(state, SelfTestState::Step::Phase7_CrossPaneVisibleRefreshDummy);
    return false;
}
case SelfTestState::Step::Phase7_CrossPaneVisibleRefreshDummy:
{
    const ULONGLONG nowTick                                = GetTickCount64();
    const std::wstring dummyFolder                         = std::format(L"/ui-sync-{}", state.runStartTick != 0 ? state.runStartTick : nowTick);
    const std::wstring dummyFile                           = std::format(L"{}/visible.txt", dummyFolder);
    const FileSystemPluginManager::PluginEntry* dummyEntry = FindLoadedPluginEntry(kPluginIdDummy);
    const std::wstring_view dummyShortId = dummyEntry && ! dummyEntry->shortId.empty() ? std::wstring_view(dummyEntry->shortId) : std::wstring_view(L"dummy");
    const std::filesystem::path displayPath = MakePluginDisplayPath(dummyShortId, dummyFolder);

    if (HasTimedOut(state, nowTick, 45'000ull))
    {
        const auto leftPath       = state.folderWindow ? state.folderWindow->GetCurrentPath(FolderWindow::Pane::Left) : std::nullopt;
        const auto rightPath      = state.folderWindow ? state.folderWindow->GetCurrentPath(FolderWindow::Pane::Right) : std::nullopt;
        const auto completionIt   = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        const bool taskCompleted  = completionIt != state.completedTasks.end();
        const HRESULT completedHr = taskCompleted ? completionIt->second.hr : E_PENDING;
        Fail(std::format(L"Phase7_CrossPaneVisibleRefreshDummy timed out. stepState={} leftPath='{}' rightPath='{}' leftVisible={} rightVisible={} "
                         L"taskStarted={} taskCompleted={} taskHr=0x{:08X}",
                         state.stepState,
                         leftPath.has_value() ? leftPath->wstring() : L"<null>",
                         rightPath.has_value() ? rightPath->wstring() : L"<null>",
                         state.folderWindow && state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"visible.txt"),
                         state.folderWindow && state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"visible.txt"),
                         state.taskA.has_value(),
                         taskCompleted,
                         static_cast<unsigned long>(completedHr)));
        return true;
    }

    if (state.stepState == 0)
    {
        const HRESULT leftHr  = state.folderWindow->SetFileSystemPluginForPane(FolderWindow::Pane::Left, kPluginIdDummy);
        const HRESULT rightHr = state.folderWindow->SetFileSystemPluginForPane(FolderWindow::Pane::Right, kPluginIdDummy);
        if (FAILED(leftHr) || FAILED(rightHr))
        {
            Fail(std::format(L"Phase7_CrossPaneVisibleRefreshDummy failed to switch panes to dummy plugin (left=0x{:08X} right=0x{:08X}).",
                             static_cast<unsigned long>(leftHr),
                             static_cast<unsigned long>(rightHr)));
            return true;
        }

        static_cast<void>(state.fsDummy->DeleteItem(
            dummyFolder.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, nullptr, nullptr));
        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyFolder) || ! DummyWriteTextFile(state.fsDummy.get(), dummyFile, "visible"))
        {
            Fail(L"Phase7_CrossPaneVisibleRefreshDummy failed to seed dummy folder.");
            return true;
        }

        state.taskA.reset();
        state.folderWindow->SetFolderPath(FolderWindow::Pane::Left, displayPath);
        state.folderWindow->SetFolderPath(FolderWindow::Pane::Right, displayPath);
        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (! CurrentPanePathEquals(state.folderWindow, FolderWindow::Pane::Left, displayPath) ||
            ! CurrentPanePathEquals(state.folderWindow, FolderWindow::Pane::Right, displayPath) ||
            ! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"visible.txt") ||
            ! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"visible.txt"))
        {
            return false;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_DELETE,
                                                 FolderWindow::Pane::Left,
                                                 std::nullopt,
                                                 state.fsDummy,
                                                 {std::filesystem::path(dummyFile)},
                                                 {},
                                                 FILESYSTEM_FLAG_NONE,
                                                 false);
        if (! state.taskA.has_value())
        {
            Fail(L"Phase7_CrossPaneVisibleRefreshDummy failed to start dummy delete task.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (it == state.completedTasks.end())
    {
        return false;
    }

    if (FAILED(it->second.hr))
    {
        Fail(std::format(L"Phase7_CrossPaneVisibleRefreshDummy delete task failed. hr=0x{:08X}", static_cast<unsigned long>(it->second.hr)));
        return true;
    }

    if (state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"visible.txt") ||
        state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"visible.txt"))
    {
        return false;
    }

    static_cast<void>(state.fsDummy->DeleteItem(
        dummyFolder.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, nullptr, nullptr));
    NextStep(state, SelfTestState::Step::Phase7_CrossPaneRelocateLocal);
    return false;
}
case SelfTestState::Step::Phase7_CrossPaneRelocateLocal:
{
    const ULONGLONG nowTick                    = GetTickCount64();
    const std::filesystem::path phaseRoot      = state.tempRoot / L"phase7-relocate";
    const std::filesystem::path doomedFolder   = phaseRoot / L"victim";
    const std::filesystem::path currentFolder  = doomedFolder / L"child" / L"grand";
    const std::filesystem::path residentFile   = currentFolder / L"inside.txt";
    const std::filesystem::path unaffectedFile = phaseRoot / L"anchor.txt";

    if (HasTimedOut(state, nowTick, 45'000ull))
    {
        Fail(L"Phase7_CrossPaneRelocateLocal timed out.");
        return true;
    }

    if (state.stepState == 0)
    {
        const HRESULT leftHr  = state.folderWindow->SetFileSystemPluginForPane(FolderWindow::Pane::Left, kPluginIdLocal);
        const HRESULT rightHr = state.folderWindow->SetFileSystemPluginForPane(FolderWindow::Pane::Right, kPluginIdLocal);
        if (FAILED(leftHr) || FAILED(rightHr))
        {
            Fail(std::format(L"Phase7_CrossPaneRelocateLocal failed to switch panes to local plugin (left=0x{:08X} right=0x{:08X}).",
                             static_cast<unsigned long>(leftHr),
                             static_cast<unsigned long>(rightHr)));
            return true;
        }

        if (! RecreateEmptyDirectory(phaseRoot))
        {
            Fail(L"Phase7_CrossPaneRelocateLocal failed to recreate root.");
            return true;
        }

        std::error_code ec;
        std::filesystem::create_directories(currentFolder, ec);
        if (ec || ! WriteTestFile(residentFile, 32) || ! WriteTestFile(unaffectedFile, 32))
        {
            Fail(L"Phase7_CrossPaneRelocateLocal failed to seed subtree.");
            return true;
        }

        state.taskA.reset();
        state.folderWindow->SetFolderPath(FolderWindow::Pane::Left, currentFolder);
        state.folderWindow->SetFolderPath(FolderWindow::Pane::Right, phaseRoot);
        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (! CurrentPanePathEquals(state.folderWindow, FolderWindow::Pane::Left, currentFolder) ||
            ! CurrentPanePathEquals(state.folderWindow, FolderWindow::Pane::Right, phaseRoot) ||
            ! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Left, L"inside.txt") ||
            ! state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"victim"))
        {
            return false;
        }

        const FileSystemFlags deleteFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        state.taskA                       = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Right, std::nullopt, state.fsLocal, {doomedFolder}, {}, deleteFlags, false);
        if (! state.taskA.has_value())
        {
            Fail(L"Phase7_CrossPaneRelocateLocal failed to start subtree delete task.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (it == state.completedTasks.end())
    {
        return false;
    }

    if (FAILED(it->second.hr))
    {
        Fail(std::format(L"Phase7_CrossPaneRelocateLocal delete task failed. hr=0x{:08X}", static_cast<unsigned long>(it->second.hr)));
        return true;
    }

    if (! CurrentPanePathEquals(state.folderWindow, FolderWindow::Pane::Left, phaseRoot) ||
        state.folderWindow->DebugHasItemDisplayName(FolderWindow::Pane::Right, L"victim"))
    {
        return false;
    }

    NextStep(state, SelfTestState::Step::Phase7_LargeDirectoryEnumeration);
    return false;
}
case SelfTestState::Step::Phase7_LargeDirectoryEnumeration:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"Phase7_LargeDirectoryEnumeration timed out.");
        return true;
    }

    const std::filesystem::path enumDir = state.tempRoot / L"enum";
    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(enumDir))
        {
            Fail(L"Failed to reset enum directory.");
            return true;
        }

        // Force the enumeration code down the grow/trim paths by lowering caps.
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":1,"enumerationHardMaxBufferMiB":8})json"));

        // Create a lot of long-named files (but stay under MAX_PATH).
        constexpr int kFileCount = 4000;
        constexpr int kPadChars  = 120;
        const std::wstring pad(kPadChars, L'x');
        for (int i = 0; i < kFileCount; ++i)
        {
            const std::filesystem::path file = enumDir / std::format(L"e_{:04}_{}.txt", i, pad);
            if (! WriteTestFile(file, 1))
            {
                Fail(L"Failed to create enum stress file.");
                return true;
            }
        }

        const std::filesystem::path unicodeA = enumDir / L"\u3053\u3093\u306b\u3061\u306f.txt"; // ã“ã‚“ã«ã¡ã¯.txt
        const std::filesystem::path unicodeB = enumDir / L"emoji_\U0001F600.txt";               // emoji_\U0001F600.txt
        if (! WriteTestFile(unicodeA, 1) || ! WriteTestFile(unicodeB, 1))
        {
            Fail(L"Failed to create enum Unicode stress files.");
            return true;
        }

        wil::com_ptr<IFilesInformation> files;
        const HRESULT hr = state.fsLocal->ReadDirectoryInfo(enumDir.c_str(), files.put());
        if (FAILED(hr))
        {
            Fail(std::format(L"ReadDirectoryInfo(enum) failed: 0x{:08X}", static_cast<unsigned long>(hr)));
            return true;
        }

        FileInfo* head = nullptr;
        if (FAILED(files->GetBuffer(&head)) || ! head)
        {
            Fail(L"ReadDirectoryInfo(enum) returned an empty buffer.");
            return true;
        }

        const std::wstring expectedA = unicodeA.filename().wstring();
        const std::wstring expectedB = unicodeB.filename().wstring();

        bool foundA = false;
        bool foundB = false;
        for (FileInfo* entry = head; entry;)
        {
            if (entry->FileNameSize >= sizeof(wchar_t))
            {
                const size_t charCount = entry->FileNameSize / sizeof(wchar_t);
                const std::wstring_view name(entry->FileName, charCount);
                if (name == expectedA)
                {
                    foundA = true;
                }
                else if (name == expectedB)
                {
                    foundB = true;
                }
            }

            if (foundA && foundB)
            {
                break;
            }

            if (entry->NextEntryOffset == 0)
            {
                break;
            }
            entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
        }

        if (! foundA || ! foundB)
        {
            std::wstring missing;
            if (! foundA)
            {
                missing.append(L" ").append(expectedA);
            }
            if (! foundB)
            {
                missing.append(L" ").append(expectedB);
            }
            Fail(std::format(L"ReadDirectoryInfo(enum) missing Unicode entries:{}", missing));
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase7_ParallelCopyMoveKnobs);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase7_ParallelCopyMoveKnobs:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 300'000ull))
    {
        Fail(L"Phase7_ParallelCopyMoveKnobs timed out.");
        return true;
    }

    const std::array<unsigned, 3> concurrencies{1u, 4u, 8u};
    const std::filesystem::path srcDir   = state.tempRoot / L"copy-src";
    const std::filesystem::path dstDir   = state.tempRoot / L"copy-dst";
    constexpr uint64_t kPromptSpeedLimit = 64ull * 1024ull;

    size_t expectedCount = CountFiles(srcDir);
    if (expectedCount == 0)
    {
        if (! SeedCopyKnobSourceFiles(srcDir))
        {
            Fail(L"Failed to reseed copy-src for knob test.");
            return true;
        }

        expectedCount = CountFiles(srcDir);
    }

    if (expectedCount == 0)
    {
        Fail(L"No files found in copy-src for knob test.");
        return true;
    }

    if (state.stepState == 0)
    {
        state.copyKnobIndex      = 0;
        state.copyKnobRetryCount = 0;
        state.stepState          = 1;
    }

    if (state.stepState == 1)
    {
        if (state.copyKnobIndex >= concurrencies.size())
        {
            NextStep(state, SelfTestState::Step::Phase7_CopyMoveConcurrency16Perf);
            return false;
        }

        const unsigned conc                       = concurrencies[state.copyKnobIndex];
        state.copySpeedLimitCleared               = false;
        state.copyPromptValidated                 = false;
        state.copyKnobObservedPerCallShare        = false;
        state.copyKnobObservedActiveCalls         = 0;
        state.copyKnobObservedDesiredSpeedLimit   = 0;
        state.copyKnobObservedAppliedSpeedLimit   = 0;
        state.copyKnobObservedEffectiveSpeedLimit = 0;
        state.copyTaskStartTick                   = nowTick;
        state.copyKnobRetryCount                  = state.copyKnobRetryCount > 1 ? 1 : state.copyKnobRetryCount;

        const std::string config = std::format(
            R"json({{"concurrencyMode":"manual","copyMoveMaxConcurrency":{},"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048}})json",
            conc);
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), config));

        if (! RecreateEmptyDirectory(dstDir))
        {
            Fail(L"Failed to reset copy-dst directory for knob test.");
            return true;
        }

        std::vector<std::filesystem::path> sources = CollectFiles(srcDir, 512u);
        const FileSystemFlags flags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

        state.taskA = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, std::move(sources), dstDir, flags, false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start copy task for knob test.");
            return true;
        }

        if (auto* task = state.fileOps->FindTask(state.taskA.value()))
        {
            const uint64_t speedLimit = state.copyKnobRetryCount == 0   ? (512ull * 1024ull)
                                        : state.copyKnobRetryCount == 1 ? (128ull * 1024ull)
                                                                        : (64ull * 1024ull);
            task->SetDesiredSpeedLimit(speedLimit);
        }

        state.stepState = 2;
        return false;
    }

    const unsigned conc                          = concurrencies[state.copyKnobIndex];
    FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    const bool completed                         = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();

    if (task)
    {
        if (conc == 8u && ! state.copySpeedLimitCleared && state.copyTaskStartTick > 0 && nowTick >= state.copyTaskStartTick &&
            (nowTick - state.copyTaskStartTick) > 1000ull)
        {
            task->SetDesiredSpeedLimit(0);
            state.copySpeedLimitCleared = true;
        }

        if (state.stepState == 2)
        {
            if (! state.copyPromptValidated && state.copyKnobIndex == 0u && state.copyKnobRetryCount == 0u)
            {
                const HWND popup = state.fileOps ? state.fileOps->GetPopupHwndForSelfTest() : nullptr;
                if (! popup)
                {
                    return false;
                }

                FileOperationsPopupInternal::PopupSelfTestInvoke invoke{};
                invoke.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskSpeedLimit;
                invoke.taskId = state.taskA.value();
                invoke.data   = 1u;
                if (! InvokePopupSelfTest(popup, invoke))
                {
                    Fail(L"Failed to invoke custom speed-limit prompt via popup self-test message.");
                    return true;
                }

                state.stepState = 3;
                return false;
            }

            if (! task->HasStarted())
            {
                return false;
            }

            state.markerTick = nowTick;
            state.stepState  = 6;
            return false;
        }

        if (state.stepState == 3)
        {
            const HWND prompt = GetFileOperationsSpeedLimitPromptHandle();
            if (! prompt)
            {
                return false;
            }

            FileOperationsSpeedLimitPromptDebugSnapshot snapshot{};
            if (! DebugGetFileOperationsSpeedLimitPromptSnapshot(snapshot))
            {
                Fail(L"Failed to capture custom speed-limit prompt snapshot.");
                return true;
            }

            if (! snapshot.usesDxUiHost)
            {
                Fail(L"Custom speed-limit prompt is not using the shared DxUi host.");
                return true;
            }

            if (snapshot.visibleChildWindowCount > 1u)
            {
                Fail(std::format(L"Custom speed-limit prompt should not expose more than one visible text-input bridge child window, got {} visible child window(s).",
                                 snapshot.visibleChildWindowCount));
                return true;
            }

            if (! DebugSetFileOperationsSpeedLimitPromptText(L"invalid-limit"))
            {
                Fail(L"Failed to write invalid text into the custom speed-limit prompt.");
                return true;
            }

            if (! DebugConfirmFileOperationsSpeedLimitPrompt())
            {
                Fail(L"Failed to submit invalid text in the custom speed-limit prompt.");
                return true;
            }

            state.stepState = 4;
            return false;
        }

        if (state.stepState == 4)
        {
            const HWND prompt = GetFileOperationsSpeedLimitPromptHandle();
            if (! prompt)
            {
                Fail(L"Custom speed-limit prompt closed unexpectedly after invalid submission.");
                return true;
            }

            FileOperationsSpeedLimitPromptDebugSnapshot snapshot{};
            if (! DebugGetFileOperationsSpeedLimitPromptSnapshot(snapshot))
            {
                Fail(L"Failed to capture validation snapshot from the custom speed-limit prompt.");
                return true;
            }

            if (snapshot.validationText.empty())
            {
                Fail(L"Custom speed-limit prompt did not surface validation after invalid input.");
                return true;
            }

            if (! DebugSetFileOperationsSpeedLimitPromptText(L"64KB"))
            {
                Fail(L"Failed to write valid text into the custom speed-limit prompt.");
                return true;
            }

            if (! DebugConfirmFileOperationsSpeedLimitPrompt())
            {
                Fail(L"Failed to confirm valid text in the custom speed-limit prompt.");
                return true;
            }

            state.stepState = 5;
            return false;
        }

        if (state.stepState == 5)
        {
            if (GetFileOperationsSpeedLimitPromptHandle() != nullptr)
            {
                return false;
            }

            const uint64_t promptLimit = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
            if (promptLimit != kPromptSpeedLimit)
            {
                return false;
            }

            const uint64_t speedLimit = state.copyKnobRetryCount == 0   ? (512ull * 1024ull)
                                        : state.copyKnobRetryCount == 1 ? (128ull * 1024ull)
                                                                        : kPromptSpeedLimit;
            task->SetDesiredSpeedLimit(speedLimit);
            state.copyPromptValidated = true;
            state.stepState           = 2;
            return false;
        }

        if (state.stepState == 6)
        {
            size_t inFlightCount     = 0;
            unsigned int activeCalls = 0;
            {
                std::scoped_lock lock(task->_progressMutex);
                inFlightCount = task->_inFlightFileCount;
                activeCalls   = static_cast<unsigned int>(task->_perItemInFlightCallCount);
            }

            if (conc == 1u)
            {
                if (inFlightCount > 1u)
                {
                    Fail(L"copyMoveMaxConcurrency=1 still produced >1 in-flight entries.");
                    return true;
                }
            }
            else
            {
                if (inFlightCount <= 1u)
                {
                    if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                    {
                        Fail(L"Expected >1 in-flight entries but did not observe them.");
                        return true;
                    }
                    return false;
                }

                const uint64_t desiredLimit   = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                const uint64_t appliedLimit   = task->_appliedSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                const uint64_t effectiveLimit = task->_effectiveSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
                if (! state.copySpeedLimitCleared && desiredLimit > 0 && activeCalls > 1u)
                {
                    const uint64_t expectedPerCall = std::max<uint64_t>(uint64_t{1}, desiredLimit / static_cast<uint64_t>(activeCalls));
                    if (appliedLimit == expectedPerCall && effectiveLimit == desiredLimit)
                    {
                        state.copyKnobObservedPerCallShare        = true;
                        state.copyKnobObservedActiveCalls         = activeCalls;
                        state.copyKnobObservedDesiredSpeedLimit   = desiredLimit;
                        state.copyKnobObservedAppliedSpeedLimit   = appliedLimit;
                        state.copyKnobObservedEffectiveSpeedLimit = effectiveLimit;
                    }
                }

                if (! state.copySpeedLimitCleared && desiredLimit > 0 && activeCalls > 1u && ! state.copyKnobObservedPerCallShare)
                {
                    return false;
                }
            }

            state.stepState = 7;
            return false;
        }
    }
    else if (! completed)
    {
        return false;
    }
    else if (state.stepState < 7)
    {
        if (conc == 1u)
        {
            state.stepState = 7;
            return false;
        }

        if (state.copyKnobRetryCount < 2)
        {
            if (! RecreateEmptyDirectory(dstDir))
            {
                Fail(L"Failed to reset copy-dst directory for knob retry.");
                return true;
            }

            if (state.taskA.has_value())
            {
                state.completedTasks.erase(state.taskA.value());
            }
            state.taskA.reset();
            state.markerTick                          = 0;
            state.copyTaskStartTick                   = 0;
            state.copySpeedLimitCleared               = false;
            state.copyPromptValidated                 = false;
            state.copyKnobObservedPerCallShare        = false;
            state.copyKnobObservedActiveCalls         = 0;
            state.copyKnobObservedDesiredSpeedLimit   = 0;
            state.copyKnobObservedAppliedSpeedLimit   = 0;
            state.copyKnobObservedEffectiveSpeedLimit = 0;
            ++state.copyKnobRetryCount;
            state.stepState = 1;
            return false;
        }

        AppendLog(L"Phase7_ParallelCopyMoveKnobs: fast completion prevented in-flight validation after retries; accepting result.");
        state.stepState = 7;
        return false;
    }

    if (! completed)
    {
        return false;
    }

    // Task completion can beat a final directory enumeration refresh on disk by a short window.
    if (! WaitForFileCount(dstDir, expectedCount, 5'000ull))
    {
        const size_t dstCount = CountFiles(dstDir);
        Fail(std::format(L"Copy output mismatch: expected {} files, got {}.", expectedCount, dstCount));
        return true;
    }

    ++state.copyKnobIndex;
    state.copyKnobRetryCount = 0;
    state.stepState          = 1;
    return false;
}
case SelfTestState::Step::Phase7_CopyMoveConcurrency16Perf:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 300'000ull))
    {
        Fail(L"Phase7_CopyMoveConcurrency16Perf timed out.");
        return true;
    }

    constexpr unsigned int kBaselineConcurrency  = 8u;
    constexpr unsigned int kCandidateConcurrency = 16u;
    constexpr int kFileCount                     = 24;
    constexpr size_t kFileBytes                  = 8ull * 1024ull * 1024ull;

    const std::filesystem::path baselineSrcDir  = state.tempRoot / L"copy-concurrency-8-src";
    const std::filesystem::path candidateSrcDir = state.tempRoot / L"copy-concurrency-16-src";
    const std::filesystem::path baselineDstDir  = state.tempRoot / L"copy-concurrency-8-dst";
    const std::filesystem::path candidateDstDir = state.tempRoot / L"copy-concurrency-16-dst";

    const auto seedSourceDir = [&](const std::filesystem::path& root, std::wstring_view label) noexcept -> bool
    {
        if (! RecreateEmptyDirectory(root))
        {
            Fail(std::format(L"Failed to reset {}.", label));
            return false;
        }

        for (int i = 0; i < kFileCount; ++i)
        {
            const std::filesystem::path file = root / std::format(L"payload_{:02}.bin", i);
            if (! WriteTestFile(file, kFileBytes))
            {
                Fail(std::format(L"Failed to write {} file {}.", label, file.native()));
                return false;
            }
        }

        return true;
    };

    const auto applyCopyMoveConfig = [&](unsigned int concurrency, std::wstring_view label) noexcept -> bool
    {
        const std::string config = std::format(
            R"json({{"concurrencyMode":"manual","copyMoveMaxConcurrency":{},"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"recycleBinBatchSize":500,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"copyReparse","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4}})json",
            concurrency);
        if (! SetPluginConfiguration(state.infoLocal.get(), config))
        {
            Fail(std::format(L"Failed to apply local plugin config for {}.", label));
            return false;
        }
        return true;
    };

    const auto startCopy = [&](const std::filesystem::path& srcDir,
                               const std::filesystem::path& dstDir,
                               std::optional<std::uint64_t>& taskSlot,
                               std::wstring_view label) noexcept -> bool
    {
        if (! RecreateEmptyDirectory(dstDir))
        {
            Fail(std::format(L"Failed to reset {} destination.", label));
            return false;
        }

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
                                                                 dstDir,
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! taskSlot.has_value())
        {
            Fail(std::format(L"Failed to start {}.", label));
            return false;
        }

        return true;
    };

    const auto observeTask = [&](FolderWindow::FileOperationState::Task* task, size_t& maxActive, unsigned int& configured) noexcept
    {
        if (! task)
        {
            return;
        }

        std::scoped_lock lock(task->_progressMutex);
        maxActive  = (std::max)(maxActive, task->_perItemInFlightCallCount);
        configured = (std::max)(configured, task->_perItemMaxConcurrency);
    };

    const auto finalizeCopy = [&](const std::optional<std::uint64_t>& taskSlot,
                                  const std::filesystem::path& dstDir,
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

        if (! WaitForFileCount(dstDir, static_cast<size_t>(kFileCount), 15'000ull))
        {
            const size_t dstCount = CountFiles(dstDir);
            Fail(std::format(L"{} output mismatch: expected {} files, got {}.", label, kFileCount, dstCount));
            return -1;
        }

        durationOut = (state.copyMoveConcurrencyPerfRunStartTick != 0 && completion.completionTick >= state.copyMoveConcurrencyPerfRunStartTick)
                          ? static_cast<uint64_t>(completion.completionTick - state.copyMoveConcurrencyPerfRunStartTick) * 1000ull
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

    if (state.stepState == 0)
    {
        state.taskA.reset();
        state.taskB.reset();
        state.copyMoveConcurrencyPerfRunStartTick        = 0;
        state.copyMoveConcurrencyPerfBaselineUs          = 0;
        state.copyMoveConcurrencyPerfCandidateUs         = 0;
        state.copyMoveConcurrencyPerfBaselineConfigured  = 0;
        state.copyMoveConcurrencyPerfCandidateConfigured = 0;
        state.copyMoveConcurrencyPerfBaselineMaxActive   = 0;
        state.copyMoveConcurrencyPerfCandidateMaxActive  = 0;

        if (! seedSourceDir(baselineSrcDir, L"copy-concurrency-8 source") || ! seedSourceDir(candidateSrcDir, L"copy-concurrency-16 source"))
        {
            return true;
        }

        if (! applyCopyMoveConfig(kBaselineConcurrency, L"copy-concurrency baseline") ||
            ! startCopy(baselineSrcDir, baselineDstDir, state.taskA, L"copy-concurrency baseline"))
        {
            return true;
        }

        state.copyMoveConcurrencyPerfRunStartTick = nowTick;
        state.stepState                           = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        observeTask(state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr,
                    state.copyMoveConcurrencyPerfBaselineMaxActive,
                    state.copyMoveConcurrencyPerfBaselineConfigured);

        const int baselineStatus = finalizeCopy(state.taskA,
                                                baselineDstDir,
                                                state.copyMoveConcurrencyPerfBaselineUs,
                                                kBaselineConcurrency,
                                                state.copyMoveConcurrencyPerfBaselineConfigured,
                                                L"copy-concurrency baseline");
        if (baselineStatus < 0)
        {
            return true;
        }
        if (baselineStatus == 0)
        {
            return false;
        }

        if (! applyCopyMoveConfig(kCandidateConcurrency, L"copy-concurrency candidate") ||
            ! startCopy(candidateSrcDir, candidateDstDir, state.taskB, L"copy-concurrency candidate"))
        {
            return true;
        }

        state.copyMoveConcurrencyPerfRunStartTick = nowTick;
        state.stepState                           = 2;
        return false;
    }

    observeTask(state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr,
                state.copyMoveConcurrencyPerfCandidateMaxActive,
                state.copyMoveConcurrencyPerfCandidateConfigured);

    const int candidateStatus = finalizeCopy(state.taskB,
                                             candidateDstDir,
                                             state.copyMoveConcurrencyPerfCandidateUs,
                                             kCandidateConcurrency,
                                             state.copyMoveConcurrencyPerfCandidateConfigured,
                                             L"copy-concurrency candidate");
    if (candidateStatus < 0)
    {
        return true;
    }
    if (candidateStatus == 0)
    {
        return false;
    }

    const std::wstring detail =
        std::format(L"fileCount={} fileBytes={} baselineConcurrency={} candidateConcurrency={} baselineMaxActive={} candidateMaxActive={}",
                    kFileCount,
                    kFileBytes,
                    kBaselineConcurrency,
                    kCandidateConcurrency,
                    state.copyMoveConcurrencyPerfBaselineMaxActive,
                    state.copyMoveConcurrencyPerfCandidateMaxActive);
    Debug::Perf::Emit(L"FileOps.SelfTest.CopyMoveConcurrency8",
                      detail,
                      state.copyMoveConcurrencyPerfBaselineUs,
                      state.copyMoveConcurrencyPerfBaselineMaxActive,
                      kBaselineConcurrency,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.CopyMoveConcurrency16",
                      detail,
                      state.copyMoveConcurrencyPerfCandidateUs,
                      state.copyMoveConcurrencyPerfCandidateMaxActive,
                      kCandidateConcurrency,
                      S_OK);
    const uint64_t improvementUs = (state.copyMoveConcurrencyPerfBaselineUs > state.copyMoveConcurrencyPerfCandidateUs)
                                       ? (state.copyMoveConcurrencyPerfBaselineUs - state.copyMoveConcurrencyPerfCandidateUs)
                                       : 0ull;
    Debug::Perf::Emit(L"FileOps.SelfTest.CopyMoveConcurrencyImprovement",
                      detail,
                      improvementUs,
                      state.copyMoveConcurrencyPerfBaselineUs,
                      state.copyMoveConcurrencyPerfCandidateUs,
                      S_OK);

    AppendLog(std::format(L"Phase7_CopyMoveConcurrency16Perf baseline={}us candidate={}us active8={} active16={}",
                          state.copyMoveConcurrencyPerfBaselineUs,
                          state.copyMoveConcurrencyPerfCandidateUs,
                          state.copyMoveConcurrencyPerfBaselineMaxActive,
                          state.copyMoveConcurrencyPerfCandidateMaxActive));

    NextStep(state, SelfTestState::Step::Phase7_AutoConcurrencyHints);
    return false;
}
case SelfTestState::Step::Phase7_AutoConcurrencyHints:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 300'000ull))
    {
        const bool manualCompleted = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();
        const bool autoCompleted   = state.taskB.has_value() && state.completedTasks.find(state.taskB.value()) != state.completedTasks.end();
        const bool deleteCompleted = state.taskC.has_value() && state.completedTasks.find(state.taskC.value()) != state.completedTasks.end();
        Fail(std::format(L"Phase7_AutoConcurrencyHints timed out. stepState={} manualCompleted={} autoCompleted={} deleteCompleted={} "
                         L"popupAutoObserved={} completedPopupAutoObserved={} popupDeleteObserved={} completedPopupDeleteObserved={} "
                         L"manualConfigured={} autoConfigured={} deleteConfigured={}.",
                         state.stepState,
                         manualCompleted,
                         autoCompleted,
                         deleteCompleted,
                         state.autoConcurrencyAutoPopupObserved,
                         state.autoConcurrencyAutoCompletedPopupObserved,
                         state.autoDeletePopupObserved,
                         state.autoDeleteCompletedPopupObserved,
                         state.autoConcurrencyManualConfigured,
                         state.autoConcurrencyAutoConfigured,
                         state.autoDeleteConfigured));
        return true;
    }

    constexpr unsigned int kManualCopyConcurrency = 1u;
    constexpr unsigned int kAutoCopyConcurrency   = 4u;
    constexpr unsigned int kAutoDeleteConcurrency = 8u;
    constexpr int kCopyFileCount                  = 16;
    constexpr size_t kCopyFileBytes               = 4ull * 1024ull * 1024ull;
    constexpr int kDeleteFileCount                = 256;
    constexpr size_t kDeleteFileBytes             = 256ull * 1024ull;

    const std::filesystem::path manualSrcDir  = state.tempRoot / L"auto-concurrency-manual-src";
    const std::filesystem::path manualDstDir  = state.tempRoot / L"auto-concurrency-manual-dst";
    const std::filesystem::path autoSrcDir    = state.tempRoot / L"auto-concurrency-auto-src";
    const std::filesystem::path autoDstDir    = state.tempRoot / L"auto-concurrency-auto-dst";
    const std::filesystem::path autoDeleteDir = state.tempRoot / L"auto-concurrency-delete";

    const auto waitForLocalConfig = [&](std::string_view expectedMode,
                                        unsigned int expectedCopyConcurrency,
                                        unsigned int expectedDeleteConcurrency,
                                        std::wstring_view label) noexcept -> bool
    {
        if (! state.fsLocal && ! state.infoLocal)
        {
            Fail(std::format(L"{} could not resolve local plugin configuration interface.", label));
            return false;
        }

        wil::com_ptr<IInformations> liveInfo;
        if (state.fsLocal)
        {
            static_cast<void>(state.fsLocal->QueryInterface(__uuidof(IInformations), liveInfo.put_void()));
        }

        IInformations* const info = liveInfo ? liveInfo.get() : state.infoLocal.get();
        if (! info)
        {
            Fail(std::format(L"{} could not query local plugin configuration.", label));
            return false;
        }

        const std::string modeNeedle   = std::format("\"concurrencyMode\":\"{}\"", expectedMode);
        const std::string copyNeedle   = std::format("\"copyMoveMaxConcurrency\":{}", expectedCopyConcurrency);
        const std::string deleteNeedle = std::format("\"deleteMaxConcurrency\":{}", expectedDeleteConcurrency);
        std::string lastConfig;
        const ULONGLONG startTick = GetTickCount64();
        for (;;)
        {
            const char* configUtf8 = nullptr;
            if (SUCCEEDED(info->GetConfiguration(&configUtf8)) && configUtf8)
            {
                lastConfig.assign(configUtf8);
                if (lastConfig.find(modeNeedle) != std::string::npos && lastConfig.find(copyNeedle) != std::string::npos &&
                    lastConfig.find(deleteNeedle) != std::string::npos)
                {
                    return true;
                }
            }

            const ULONGLONG nowTick = GetTickCount64();
            if ((nowTick - startTick) >= 2000ull)
            {
                Fail(std::format(L"{} live local plugin config did not converge within timeout. expectedMode='{}' copy={} delete={} config='{}'.",
                                 label,
                                 std::wstring(expectedMode.begin(), expectedMode.end()),
                                 expectedCopyConcurrency,
                                 expectedDeleteConcurrency,
                                 std::wstring(lastConfig.begin(), lastConfig.end())));
                return false;
            }

            ::Sleep(10);
        }
    };

    const auto seedSourceDir = [&](const std::filesystem::path& root, std::wstring_view label) noexcept -> bool
    {
        if (! RecreateEmptyDirectory(root))
        {
            Fail(std::format(L"Failed to reset {}.", label));
            return false;
        }

        for (int i = 0; i < kCopyFileCount; ++i)
        {
            const std::filesystem::path file = root / std::format(L"payload_{:02}.bin", i);
            if (! WriteTestFile(file, kCopyFileBytes))
            {
                Fail(std::format(L"Failed to write {} file {}.", label, file.native()));
                return false;
            }
        }

        return true;
    };

    const auto applyLocalConfig =
        [&](std::string_view mode, unsigned int copyConcurrency, unsigned int deleteConcurrency, std::wstring_view label) noexcept -> bool
    {
        const std::string config = std::format(
            R"json({{"concurrencyMode":"{}","copyMoveMaxConcurrency":{},"deleteMaxConcurrency":{},"deleteRecycleBinMaxConcurrency":1,"recycleBinBatchSize":500,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"copyReparse","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4}})json",
            mode,
            copyConcurrency,
            deleteConcurrency);
        if (! SetPluginConfiguration(state.infoLocal.get(), config))
        {
            Fail(std::format(L"Failed to apply local plugin config for {}.", label));
            return false;
        }

        state.localConfigDirty = true;
        return waitForLocalConfig(mode, copyConcurrency, deleteConcurrency, label);
    };

    const auto startCopy = [&](const std::filesystem::path& srcDir,
                               const std::filesystem::path& dstDir,
                               std::optional<std::uint64_t>& taskSlot,
                               std::wstring_view label) noexcept -> bool
    {
        if (! RecreateEmptyDirectory(dstDir))
        {
            Fail(std::format(L"Failed to reset {} destination.", label));
            return false;
        }

        std::vector<std::filesystem::path> sources = CollectFiles(srcDir, 512u);
        if (sources.size() != static_cast<size_t>(kCopyFileCount))
        {
            Fail(std::format(L"{} expected {} source files, got {}.", label, kCopyFileCount, sources.size()));
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
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! taskSlot.has_value())
        {
            Fail(std::format(L"Failed to start {}.", label));
            return false;
        }

        return true;
    };

    const auto observeConfigured = [&](FolderWindow::FileOperationState::Task* task, unsigned int& configured) noexcept
    {
        if (! task)
        {
            return;
        }

        std::scoped_lock lock(task->_progressMutex);
        configured = (std::max)(configured, task->_perItemMaxConcurrency);
    };

    const auto finalizeCopy = [&](const std::optional<std::uint64_t>& taskSlot,
                                  const std::filesystem::path& dstDir,
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

        if (! WaitForFileCount(dstDir, static_cast<size_t>(kCopyFileCount), 15'000ull))
        {
            const size_t dstCount = CountFiles(dstDir);
            Fail(std::format(L"{} output mismatch: expected {} files, got {}.", label, kCopyFileCount, dstCount));
            return -1;
        }

        durationOut = (state.autoConcurrencyRunStartTick != 0 && completion.completionTick >= state.autoConcurrencyRunStartTick)
                          ? static_cast<uint64_t>(completion.completionTick - state.autoConcurrencyRunStartTick) * 1000ull
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

    const auto startDelete = [&](std::optional<std::uint64_t>& taskSlot, std::wstring_view label) noexcept -> bool
    {
        std::vector<std::filesystem::path> deletePaths;
        if (! CreateSiblingFiles(autoDeleteDir, kDeleteFileCount, kDeleteFileBytes, deletePaths))
        {
            Fail(std::format(L"Failed to create sibling files for {}.", label));
            return false;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        taskSlot                    = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_DELETE,
                                                                 FolderWindow::Pane::Left,
                                                                 std::nullopt,
                                                                 state.fsLocal,
                                                                 std::move(deletePaths),
                                                                 {},
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! taskSlot.has_value())
        {
            Fail(std::format(L"Failed to start {}.", label));
            return false;
        }

        return true;
    };

    const auto finalizeDelete = [&](const std::optional<std::uint64_t>& taskSlot,
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

        if (observedConfigured != expectedConfigured)
        {
            Fail(std::format(L"{} expected configured concurrency {} but observed {}.", label, expectedConfigured, observedConfigured));
            return -1;
        }

        const std::vector<std::filesystem::path> remaining = CollectFiles(autoDeleteDir, 128u);
        if (! remaining.empty())
        {
            Fail(std::format(L"{} left {} undeleted files.", label, remaining.size()));
            return -1;
        }

        return 1;
    };

    const auto verifyAutoConcurrencyDiagnostics =
        [&](uint64_t taskId, unsigned int expectedAutoConcurrency, unsigned int expectedEffectiveConcurrency, std::wstring_view label) noexcept -> int
    {
        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(L"RedSalamander");
        if (settingsPath.empty())
        {
            Fail(std::format(L"{} could not resolve settings path for diagnostics lookup.", label));
            return -1;
        }

        const std::filesystem::path settingsDir = settingsPath.parent_path();
        const std::filesystem::path logsDir     = settingsDir.parent_path().empty() ? (settingsDir / L"Logs") : (settingsDir.parent_path() / L"Logs");

        std::filesystem::path diagnosticsPath;
        std::error_code ec;
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

            if (diagnosticsPath.empty() || de.path().filename().wstring() > diagnosticsPath.filename().wstring())
            {
                diagnosticsPath = de.path();
            }
        }

        if (diagnosticsPath.empty())
        {
            return 0;
        }

        std::wstring diagnosticsText;
        if (! ReadUtf16TextFile(diagnosticsPath, diagnosticsText))
        {
            return 0;
        }

        std::wstring line;
        if (! TryFindDiagnosticLine(diagnosticsText, taskId, L"task.autoConcurrency", line))
        {
            return 0;
        }

        const std::wstring autoNeedle      = std::format(L"\"autoTunedConcurrency\":{}", expectedAutoConcurrency);
        const std::wstring effectiveNeedle = std::format(L"\"effectiveConcurrencyBudget\":{}", expectedEffectiveConcurrency);
        if (line.find(L"\"concurrencyMode\":\"auto\"") == std::wstring::npos || line.find(L"\"storageType\":\"") == std::wstring::npos ||
            line.find(autoNeedle) == std::wstring::npos || line.find(effectiveNeedle) == std::wstring::npos)
        {
            Fail(std::format(L"{} diagnostics line missing expected auto-concurrency fields: {}.", label, line));
            return -1;
        }

        return 1;
    };

    const auto verifyAutoConcurrencyPopup = [&](const std::optional<std::uint64_t>& taskSlot,
                                                unsigned int expectedAutoConcurrency,
                                                unsigned int expectedEffectiveConcurrency,
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

        if (! popupSnapshot.autoConcurrencyUsed || popupSnapshot.autoTunedConcurrency != expectedAutoConcurrency ||
            popupSnapshot.effectiveConcurrencyBudget != expectedEffectiveConcurrency)
        {
            const auto* liveTask                = state.fileOps ? state.fileOps->FindTask(taskSlot.value()) : nullptr;
            const bool liveAutoUsed             = liveTask ? liveTask->_autoConcurrencyUsed.load(std::memory_order_acquire) : false;
            const unsigned int liveResolved     = liveTask ? liveTask->_autoTunedConcurrency.load(std::memory_order_acquire) : 0u;
            const unsigned int liveApplied      = liveTask ? liveTask->_effectiveConcurrencyBudget.load(std::memory_order_acquire) : 0u;
            const unsigned int liveConfigured   = liveTask ? liveTask->_perItemMaxConcurrency : 0u;
            const unsigned int liveFlags        = liveTask ? static_cast<unsigned int>(liveTask->_flags) : 0u;
            HRESULT storageHr                   = E_FAIL;
            unsigned int storagePreferredCopy   = 0u;
            unsigned int storagePreferredDelete = 0u;
            if (liveTask && state.fsLocal && ! liveTask->_sourcePaths.empty())
            {
                FileSystemStorageCharacteristics characteristics{};
                characteristics.sizeBytes = sizeof(FileSystemStorageCharacteristics);
                storageHr                 = state.fsLocal->GetStorageCharacteristics(liveTask->_sourcePaths.front().c_str(), &characteristics);
                if (SUCCEEDED(storageHr))
                {
                    storagePreferredCopy   = characteristics.preferredCopyMoveConcurrency;
                    storagePreferredDelete = characteristics.preferredDeleteConcurrency;
                }
            }
            if (liveTask && liveApplied == 0u)
            {
                return 0;
            }
            Fail(std::format(L"{} popup diagnostics mismatch: popupFinished={} autoUsed={} resolved={} applied={} livePresent={} liveAutoUsed={} "
                             L"liveResolved={} liveApplied={} liveConfigured={} liveFlags=0x{:08X} storageHr=0x{:08X} storageCopy={} storageDelete={}.",
                             label,
                             popupSnapshot.finished,
                             popupSnapshot.autoConcurrencyUsed,
                             popupSnapshot.autoTunedConcurrency,
                             popupSnapshot.effectiveConcurrencyBudget,
                             liveTask != nullptr,
                             liveAutoUsed,
                             liveResolved,
                             liveApplied,
                             liveConfigured,
                             liveFlags,
                             static_cast<unsigned long>(storageHr),
                             storagePreferredCopy,
                             storagePreferredDelete));
            return -1;
        }

        observed    = true;
        resolvedOut = popupSnapshot.autoTunedConcurrency;
        appliedOut  = popupSnapshot.effectiveConcurrencyBudget;
        return 1;
    };

    const auto verifyCompletedAutoConcurrencyPopup = [&](const std::optional<std::uint64_t>& taskSlot,
                                                         unsigned int expectedAutoConcurrency,
                                                         unsigned int expectedEffectiveConcurrency,
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
            Fail(std::format(L"{} was never started for completed popup diagnostics.", label));
            return -1;
        }

        if (state.completedTasks.find(taskSlot.value()) == state.completedTasks.end())
        {
            return 0;
        }

        FolderWindow::FileOperationState::CompletedTaskSummary completedSummary{};
        bool haveCompletedSummary = false;
        if (state.fileOps)
        {
            std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
            state.fileOps->CollectCompletedTasks(summaries);
            for (const auto& candidate : summaries)
            {
                if (candidate.taskId != taskSlot.value())
                {
                    continue;
                }

                completedSummary     = candidate;
                haveCompletedSummary = true;
                break;
            }
        }

        FileOperationsPopupInternal::TaskSnapshot popupSnapshot{};
        const bool havePopupSnapshot = TryGetPopupTaskSnapshot(state.fileOps, taskSlot.value(), popupSnapshot);
        if (havePopupSnapshot && ! popupSnapshot.finished)
        {
            return 0;
        }

        if (havePopupSnapshot)
        {
            if (! popupSnapshot.autoConcurrencyUsed || popupSnapshot.autoTunedConcurrency != expectedAutoConcurrency ||
                popupSnapshot.effectiveConcurrencyBudget != expectedEffectiveConcurrency)
            {
                Fail(std::format(L"{} completed popup diagnostics mismatch: autoUsed={} resolved={} applied={} finished={}.",
                                 label,
                                 popupSnapshot.autoConcurrencyUsed,
                                 popupSnapshot.autoTunedConcurrency,
                                 popupSnapshot.effectiveConcurrencyBudget,
                                 popupSnapshot.finished));
                return -1;
            }

            observed    = true;
            resolvedOut = popupSnapshot.autoTunedConcurrency;
            appliedOut  = popupSnapshot.effectiveConcurrencyBudget;
            return 1;
        }

        if (! haveCompletedSummary)
        {
            return 0;
        }

        if (! completedSummary.autoConcurrencyUsed || completedSummary.autoTunedConcurrency != expectedAutoConcurrency ||
            completedSummary.effectiveConcurrencyBudget != expectedEffectiveConcurrency)
        {
            Fail(std::format(L"{} completed task summary mismatch: autoUsed={} resolved={} applied={}.",
                             label,
                             completedSummary.autoConcurrencyUsed,
                             completedSummary.autoTunedConcurrency,
                             completedSummary.effectiveConcurrencyBudget));
            return -1;
        }

        observed    = true;
        resolvedOut = completedSummary.autoTunedConcurrency;
        appliedOut  = completedSummary.effectiveConcurrencyBudget;
        return 1;
    };

    if (state.stepState == 0)
    {
        state.taskA.reset();
        state.taskB.reset();
        state.taskC.reset();
        state.fileOps->SetAutoDismissSuccess(false);
        state.autoConcurrencyRunStartTick               = 0;
        state.autoConcurrencyManualUs                   = 0;
        state.autoConcurrencyAutoUs                     = 0;
        state.autoConcurrencyManualConfigured           = 0;
        state.autoConcurrencyAutoConfigured             = 0;
        state.autoDeleteConfigured                      = 0;
        state.autoConcurrencyAutoPopupObserved          = false;
        state.autoConcurrencyAutoPopupResolved          = 0;
        state.autoConcurrencyAutoPopupApplied           = 0;
        state.autoConcurrencyAutoCompletedPopupObserved = false;
        state.autoConcurrencyAutoCompletedPopupResolved = 0;
        state.autoConcurrencyAutoCompletedPopupApplied  = 0;
        state.autoDeletePopupObserved                   = false;
        state.autoDeletePopupResolved                   = 0;
        state.autoDeletePopupApplied                    = 0;
        state.autoDeleteCompletedPopupObserved          = false;
        state.autoDeleteCompletedPopupResolved          = 0;
        state.autoDeleteCompletedPopupApplied           = 0;

        if (! seedSourceDir(manualSrcDir, L"auto-concurrency manual source") || ! seedSourceDir(autoSrcDir, L"auto-concurrency auto source"))
        {
            return true;
        }

        if (! applyLocalConfig("manual", kManualCopyConcurrency, 1u, L"auto-concurrency manual copy") ||
            ! startCopy(manualSrcDir, manualDstDir, state.taskA, L"auto-concurrency manual copy"))
        {
            return true;
        }

        state.autoConcurrencyRunStartTick = nowTick;
        state.stepState                   = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        observeConfigured(state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr, state.autoConcurrencyManualConfigured);

        const int manualStatus = finalizeCopy(state.taskA,
                                              manualDstDir,
                                              state.autoConcurrencyManualUs,
                                              kManualCopyConcurrency,
                                              state.autoConcurrencyManualConfigured,
                                              L"auto-concurrency manual copy");
        if (manualStatus < 0)
        {
            return true;
        }
        if (manualStatus == 0)
        {
            return false;
        }

        if (! applyLocalConfig("auto", 1u, 1u, L"auto-concurrency auto copy") ||
            ! startCopy(autoSrcDir, autoDstDir, state.taskB, L"auto-concurrency auto copy"))
        {
            return true;
        }

        state.autoConcurrencyRunStartTick = nowTick;
        state.stepState                   = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        observeConfigured(state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr, state.autoConcurrencyAutoConfigured);
        const int autoPopupStatus = verifyAutoConcurrencyPopup(state.taskB,
                                                               kAutoCopyConcurrency,
                                                               kAutoCopyConcurrency,
                                                               state.autoConcurrencyAutoPopupObserved,
                                                               state.autoConcurrencyAutoPopupResolved,
                                                               state.autoConcurrencyAutoPopupApplied,
                                                               L"auto-concurrency auto copy");
        if (autoPopupStatus < 0)
        {
            return true;
        }

        const int autoStatus = finalizeCopy(
            state.taskB, autoDstDir, state.autoConcurrencyAutoUs, kAutoCopyConcurrency, state.autoConcurrencyAutoConfigured, L"auto-concurrency auto copy");
        if (autoStatus < 0)
        {
            return true;
        }
        if (autoStatus == 0)
        {
            return false;
        }
        if (autoPopupStatus == 0)
        {
            Fail(L"auto-concurrency auto copy completed before popup diagnostics could be observed.");
            return true;
        }

        const int autoCompletedPopupStatus = verifyCompletedAutoConcurrencyPopup(state.taskB,
                                                                                 kAutoCopyConcurrency,
                                                                                 kAutoCopyConcurrency,
                                                                                 state.autoConcurrencyAutoCompletedPopupObserved,
                                                                                 state.autoConcurrencyAutoCompletedPopupResolved,
                                                                                 state.autoConcurrencyAutoCompletedPopupApplied,
                                                                                 L"auto-concurrency auto copy");
        if (autoCompletedPopupStatus < 0)
        {
            return true;
        }
        if (autoCompletedPopupStatus == 0)
        {
            return false;
        }

        if (! applyLocalConfig("auto", 1u, 1u, L"auto-concurrency auto delete") || ! startDelete(state.taskC, L"auto-concurrency auto delete"))
        {
            return true;
        }

        state.stepState = 3;
        return false;
    }

    observeConfigured(state.taskC.has_value() ? state.fileOps->FindTask(state.taskC.value()) : nullptr, state.autoDeleteConfigured);
    const int deletePopupStatus = verifyAutoConcurrencyPopup(state.taskC,
                                                             kAutoDeleteConcurrency,
                                                             kAutoDeleteConcurrency,
                                                             state.autoDeletePopupObserved,
                                                             state.autoDeletePopupResolved,
                                                             state.autoDeletePopupApplied,
                                                             L"auto-concurrency auto delete");
    if (deletePopupStatus < 0)
    {
        return true;
    }

    const int deleteStatus = finalizeDelete(state.taskC, kAutoDeleteConcurrency, state.autoDeleteConfigured, L"auto-concurrency auto delete");
    if (deleteStatus < 0)
    {
        return true;
    }
    if (deleteStatus == 0)
    {
        return false;
    }
    if (deletePopupStatus == 0)
    {
        Fail(L"auto-concurrency auto delete completed before popup diagnostics could be observed.");
        return true;
    }

    const int deleteCompletedPopupStatus = verifyCompletedAutoConcurrencyPopup(state.taskC,
                                                                               kAutoDeleteConcurrency,
                                                                               kAutoDeleteConcurrency,
                                                                               state.autoDeleteCompletedPopupObserved,
                                                                               state.autoDeleteCompletedPopupResolved,
                                                                               state.autoDeleteCompletedPopupApplied,
                                                                               L"auto-concurrency auto delete");
    if (deleteCompletedPopupStatus < 0)
    {
        return true;
    }
    if (deleteCompletedPopupStatus == 0)
    {
        return false;
    }

    const int autoCopyDiagnostics =
        verifyAutoConcurrencyDiagnostics(state.taskB.value(), kAutoCopyConcurrency, kAutoCopyConcurrency, L"auto-concurrency auto copy");
    if (autoCopyDiagnostics < 0)
    {
        return true;
    }
    if (autoCopyDiagnostics == 0)
    {
        return false;
    }

    const int autoDeleteDiagnostics =
        verifyAutoConcurrencyDiagnostics(state.taskC.value(), kAutoDeleteConcurrency, kAutoDeleteConcurrency, L"auto-concurrency auto delete");
    if (autoDeleteDiagnostics < 0)
    {
        return true;
    }
    if (autoDeleteDiagnostics == 0)
    {
        return false;
    }

    const std::wstring detail =
        std::format(L"fileCount={} fileBytes={} manualConfigured={} autoConfigured={} deleteConfigured={} popupAuto={}/{} popupDelete={}/{} "
                    L"completedPopupAuto={}/{} completedPopupDelete={}/{}",
                    kCopyFileCount,
                    kCopyFileBytes,
                    state.autoConcurrencyManualConfigured,
                    state.autoConcurrencyAutoConfigured,
                    state.autoDeleteConfigured,
                    state.autoConcurrencyAutoPopupResolved,
                    state.autoConcurrencyAutoPopupApplied,
                    state.autoDeletePopupResolved,
                    state.autoDeletePopupApplied,
                    state.autoConcurrencyAutoCompletedPopupResolved,
                    state.autoConcurrencyAutoCompletedPopupApplied,
                    state.autoDeleteCompletedPopupResolved,
                    state.autoDeleteCompletedPopupApplied);
    Debug::Perf::Emit(
        L"FileOps.SelfTest.AutoConcurrencyManual", detail, state.autoConcurrencyManualUs, state.autoConcurrencyManualConfigured, kManualCopyConcurrency, S_OK);
    Debug::Perf::Emit(
        L"FileOps.SelfTest.AutoConcurrencyAuto", detail, state.autoConcurrencyAutoUs, state.autoConcurrencyAutoConfigured, kAutoCopyConcurrency, S_OK);
    const uint64_t improvementUs =
        (state.autoConcurrencyManualUs > state.autoConcurrencyAutoUs) ? (state.autoConcurrencyManualUs - state.autoConcurrencyAutoUs) : 0ull;
    Debug::Perf::Emit(L"FileOps.SelfTest.AutoConcurrencyImprovement", detail, improvementUs, state.autoConcurrencyManualUs, state.autoConcurrencyAutoUs, S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.AutoDeleteConfigured", detail, state.autoDeleteConfigured, kAutoDeleteConcurrency, state.autoDeleteConfigured, S_OK);

    AppendLog(std::format(L"Phase7_AutoConcurrencyHints manual={}us auto={}us configuredManual={} configuredAuto={} configuredDelete={}",
                          state.autoConcurrencyManualUs,
                          state.autoConcurrencyAutoUs,
                          state.autoConcurrencyManualConfigured,
                          state.autoConcurrencyAutoConfigured,
                          state.autoDeleteConfigured));

    NextStep(state, SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines);
    return false;
}
case SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase7_PerItemDirectoryCopyInFlightLines timed out.");
        return true;
    }

    const std::filesystem::path srcRoot = state.tempRoot / L"peritem-dir-src";
    const std::filesystem::path dstRoot = state.tempRoot / L"peritem-dir-dst";
    const std::filesystem::path srcDir  = srcRoot / L"payload";

    constexpr int kFileCount       = 4;
    constexpr size_t kFileBytes    = 4ull * 1024ull * 1024ull;
    constexpr uint64_t kSpeedLimit = 2ull * 1024ull * 1024ull;

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json"));

        state.taskA.reset();
        state.markerTick                          = 0;
        state.copyKnobObservedPerCallShare        = false;
        state.copyKnobObservedActiveCalls         = 0;
        state.copyKnobObservedDesiredSpeedLimit   = 0;
        state.copyKnobObservedAppliedSpeedLimit   = 0;
        state.copyKnobObservedEffectiveSpeedLimit = 0;
        state.copyTaskStartTick                   = nowTick;

        if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(srcDir))
        {
            Fail(L"Failed to reset per-item directory copy directories.");
            return true;
        }

        for (int i = 0; i < kFileCount; ++i)
        {
            const std::filesystem::path file = srcDir / std::format(L"p_{:02}.bin", i);
            if (! WriteTestFile(file, kFileBytes))
            {
                Fail(L"Failed to write per-item directory copy test file.");
                return true;
            }
        }

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
                                                 kSpeedLimit,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start per-item directory copy task.");
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
            size_t inFlightCount = 0;
            {
                std::scoped_lock lock(task->_progressMutex);
                inFlightCount = task->_inFlightFileCount;
            }

            if (inFlightCount <= 1u)
            {
                if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                {
                    Fail(L"Per-item directory copy: expected >1 in-flight entries but did not observe them.");
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
        Fail(std::format(L"Per-item directory copy task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
        return true;
    }

    const std::filesystem::path dstCopiedDir = dstRoot / srcDir.filename();
    const size_t dstCount                    = CountFiles(dstCopiedDir);
    if (dstCount != static_cast<size_t>(kFileCount))
    {
        Fail(std::format(L"Per-item directory copy output mismatch: expected {} files, got {}.", kFileCount, dstCount));
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase7_CopyItemsSingleFolderRecursiveParallelism);
    return false;
}
case SelfTestState::Step::Phase7_CopyItemsSingleFolderRecursiveParallelism:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Phase7_CopyItemsSingleFolderRecursiveParallelism timed out.");
        return true;
    }

    struct RecordingCallback final : IFileSystemCallback
    {
        std::array<uint64_t, 16> streamIds{};
        size_t streamCount      = 0;
        uint64_t progressCount  = 0;
        uint64_t completedCount = 0;
        uint64_t completedBytes = 0;

        void RecordStream(uint64_t streamId) noexcept
        {
            for (size_t i = 0; i < streamCount; ++i)
            {
                if (streamIds[i] == streamId)
                {
                    return;
                }
            }

            if (streamCount < streamIds.size())
            {
                streamIds[streamCount++] = streamId;
            }
        }

        HRESULT STDMETHODCALLTYPE FileSystemProgress([[maybe_unused]] FileSystemOperation operationType,
                                                     [[maybe_unused]] unsigned long totalItems,
                                                     [[maybe_unused]] unsigned long completedItems,
                                                     [[maybe_unused]] uint64_t totalBytes,
                                                     uint64_t completedBytesValue,
                                                     const wchar_t* currentSourcePath,
                                                     [[maybe_unused]] const wchar_t* currentDestinationPath,
                                                     uint64_t currentItemTotalBytes,
                                                     [[maybe_unused]] uint64_t currentItemCompletedBytes,
                                                     [[maybe_unused]] FileSystemOptions* options,
                                                     uint64_t progressStreamId,
                                                     [[maybe_unused]] void* cookie) noexcept override
        {
            ++progressCount;
            completedBytes = (std::max)(completedBytes, completedBytesValue);
            if (currentSourcePath && currentItemTotalBytes > 0)
            {
                RecordStream(progressStreamId);
            }
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemItemCompleted([[maybe_unused]] FileSystemOperation operationType,
                                                          [[maybe_unused]] unsigned long itemIndex,
                                                          [[maybe_unused]] const wchar_t* sourcePath,
                                                          [[maybe_unused]] const wchar_t* destinationPath,
                                                          [[maybe_unused]] HRESULT status,
                                                          [[maybe_unused]] FileSystemOptions* options,
                                                          [[maybe_unused]] void* cookie) noexcept override
        {
            ++completedCount;
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* cancel, [[maybe_unused]] void* cookie) noexcept override
        {
            if (! cancel)
            {
                return E_POINTER;
            }
            *cancel = FALSE;
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemIssue([[maybe_unused]] FileSystemOperation operationType,
                                                  [[maybe_unused]] const wchar_t* sourcePath,
                                                  [[maybe_unused]] const wchar_t* destinationPath,
                                                  [[maybe_unused]] HRESULT status,
                                                  FileSystemIssueAction* action,
                                                  [[maybe_unused]] FileSystemOptions* options,
                                                  [[maybe_unused]] void* cookie) noexcept override
        {
            if (! action)
            {
                return E_POINTER;
            }
            *action = FileSystemIssueAction::Cancel;
            return S_OK;
        }
    };

    const std::filesystem::path srcRoot   = state.tempRoot / L"copyitems-single-src";
    const std::filesystem::path dstRoot   = state.tempRoot / L"copyitems-single-dst";
    const std::filesystem::path srcDir    = srcRoot / L"payload";
    const std::filesystem::path nestedDir = srcDir / L"single-child";

    constexpr int kFileCount    = 4;
    constexpr size_t kFileBytes = 512ull * 1024ull;

    static_cast<void>(SetPluginConfiguration(
        state.infoLocal.get(),
        R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json"));

    if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(nestedDir))
    {
        Fail(L"CopyItems single-folder recursive parallelism failed to reset directories.");
        return true;
    }

    for (int i = 0; i < kFileCount; ++i)
    {
        if (! WriteTestFile(nestedDir / std::format(L"nested_{:02}.bin", i), kFileBytes))
        {
            Fail(L"CopyItems single-folder recursive parallelism failed to seed source files.");
            return true;
        }
    }

    RecordingCallback callback{};
    FileSystemOptions options{};
    options.sizeBytes = sizeof(FileSystemOptions);

    const std::array<const wchar_t*, 1> sources{srcDir.c_str()};
    const FileSystemFlags flags =
        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

    const ULONGLONG startedTick = GetTickCount64();
    const HRESULT hr =
        state.fsLocal->CopyItems(sources.data(), static_cast<unsigned long>(sources.size()), dstRoot.c_str(), flags, &options, &callback, nullptr);
    const uint64_t durationUs = (GetTickCount64() >= startedTick) ? (GetTickCount64() - startedTick) * 1000ull : 0ull;

    Debug::Perf::Emit(L"FileOps.SelfTest.CopyItemsSingleFolderRecursiveParallelismUs",
                      L"shape=single-selected-folder-with-one-nested-child",
                      durationUs,
                      static_cast<uint64_t>(callback.streamCount),
                      static_cast<uint64_t>(kFileCount),
                      hr);
    Debug::Perf::Emit(L"FileOps.SelfTest.CopyItemsSingleFolderProgressStreams",
                      L"shape=single-selected-folder-with-one-nested-child",
                      static_cast<uint64_t>(callback.streamCount),
                      callback.progressCount,
                      callback.completedCount,
                      hr);

    if (FAILED(hr))
    {
        Fail(std::format(L"CopyItems single-folder recursive parallelism copy failed: 0x{:08X}.", static_cast<unsigned long>(hr)));
        return true;
    }

    const std::filesystem::path dstNestedDir = dstRoot / srcDir.filename() / nestedDir.filename();
    const size_t dstCount                    = CountFiles(dstNestedDir);
    if (dstCount != static_cast<size_t>(kFileCount))
    {
        Fail(std::format(L"CopyItems single-folder recursive parallelism output mismatch: expected {} files, got {}.", kFileCount, dstCount));
        return true;
    }

    if (callback.completedCount != 1u)
    {
        Fail(std::format(L"CopyItems single-folder recursive parallelism expected one item-completed callback, got {}.", callback.completedCount));
        return true;
    }

    if (callback.streamCount <= 1u)
    {
        Fail(std::format(L"CopyItems(count==1) nested folder copy expected multiple progress streams, got {}.", callback.streamCount));
        return true;
    }

    const std::filesystem::path directDstRoot = state.tempRoot / L"copyitem-single-dst";
    const std::filesystem::path directDstDir  = directDstRoot / srcDir.filename();
    if (! RecreateEmptyDirectory(directDstRoot))
    {
        Fail(L"CopyItem single-folder recursive parallelism failed to reset destination directory.");
        return true;
    }

    RecordingCallback directCallback{};
    const ULONGLONG directStartedTick = GetTickCount64();
    const HRESULT directHr            = state.fsLocal->CopyItem(srcDir.c_str(), directDstDir.c_str(), flags, &options, &directCallback, nullptr);
    const uint64_t directDurationUs   = (GetTickCount64() >= directStartedTick) ? (GetTickCount64() - directStartedTick) * 1000ull : 0ull;

    Debug::Perf::Emit(L"FileOps.SelfTest.CopyItemSingleFolderRecursiveParallelismUs",
                      L"shape=single-selected-folder-with-one-nested-child",
                      directDurationUs,
                      static_cast<uint64_t>(directCallback.streamCount),
                      static_cast<uint64_t>(kFileCount),
                      directHr);
    Debug::Perf::Emit(L"FileOps.SelfTest.CopyItemSingleFolderProgressStreams",
                      L"shape=single-selected-folder-with-one-nested-child",
                      static_cast<uint64_t>(directCallback.streamCount),
                      directCallback.progressCount,
                      directCallback.completedCount,
                      directHr);

    if (FAILED(directHr))
    {
        Fail(std::format(L"CopyItem single-folder recursive parallelism copy failed: 0x{:08X}.", static_cast<unsigned long>(directHr)));
        return true;
    }

    const std::filesystem::path directDstNestedDir = directDstDir / nestedDir.filename();
    const size_t directDstCount                    = CountFiles(directDstNestedDir);
    if (directDstCount != static_cast<size_t>(kFileCount))
    {
        Fail(std::format(L"CopyItem single-folder recursive parallelism output mismatch: expected {} files, got {}.", kFileCount, directDstCount));
        return true;
    }

    if (directCallback.completedCount != 1u)
    {
        Fail(std::format(L"CopyItem single-folder recursive parallelism expected one item-completed callback, got {}.", directCallback.completedCount));
        return true;
    }

    if (directCallback.streamCount <= 1u)
    {
        Fail(std::format(L"CopyItem nested folder copy expected multiple progress streams, got {}.", directCallback.streamCount));
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase7_CopyItemsMultiRootUnevenRecursiveParallelism);
    return false;
}
case SelfTestState::Step::Phase7_CopyItemsMultiRootUnevenRecursiveParallelism:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Phase7_CopyItemsMultiRootUnevenRecursiveParallelism timed out.");
        return true;
    }

    struct RecordingCallback final : IFileSystemCallback
    {
        std::wstring dominantNeedle;
        std::array<uint64_t, 16> dominantStreamIds{};
        size_t dominantStreamCount = 0;
        uint64_t progressCount     = 0;
        uint64_t dominantProgress  = 0;
        uint64_t completedCount    = 0;

        void RecordDominantStream(uint64_t streamId) noexcept
        {
            for (size_t i = 0; i < dominantStreamCount; ++i)
            {
                if (dominantStreamIds[i] == streamId)
                {
                    return;
                }
            }

            if (dominantStreamCount < dominantStreamIds.size())
            {
                dominantStreamIds[dominantStreamCount++] = streamId;
            }
        }

        HRESULT STDMETHODCALLTYPE FileSystemProgress([[maybe_unused]] FileSystemOperation operationType,
                                                     [[maybe_unused]] unsigned long totalItems,
                                                     [[maybe_unused]] unsigned long completedItems,
                                                     [[maybe_unused]] uint64_t totalBytes,
                                                     [[maybe_unused]] uint64_t completedBytes,
                                                     const wchar_t* currentSourcePath,
                                                     [[maybe_unused]] const wchar_t* currentDestinationPath,
                                                     uint64_t currentItemTotalBytes,
                                                     [[maybe_unused]] uint64_t currentItemCompletedBytes,
                                                     [[maybe_unused]] FileSystemOptions* options,
                                                     uint64_t progressStreamId,
                                                     [[maybe_unused]] void* cookie) noexcept override
        {
            ++progressCount;
            if (currentSourcePath && currentItemTotalBytes > 0 && ! dominantNeedle.empty())
            {
                const std::wstring_view sourceView{currentSourcePath};
                if (sourceView.find(std::wstring_view{dominantNeedle}) != std::wstring_view::npos)
                {
                    ++dominantProgress;
                    RecordDominantStream(progressStreamId);
                }
            }
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemItemCompleted([[maybe_unused]] FileSystemOperation operationType,
                                                          [[maybe_unused]] unsigned long itemIndex,
                                                          [[maybe_unused]] const wchar_t* sourcePath,
                                                          [[maybe_unused]] const wchar_t* destinationPath,
                                                          [[maybe_unused]] HRESULT status,
                                                          [[maybe_unused]] FileSystemOptions* options,
                                                          [[maybe_unused]] void* cookie) noexcept override
        {
            ++completedCount;
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* cancel, [[maybe_unused]] void* cookie) noexcept override
        {
            if (! cancel)
            {
                return E_POINTER;
            }
            *cancel = FALSE;
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemIssue([[maybe_unused]] FileSystemOperation operationType,
                                                  [[maybe_unused]] const wchar_t* sourcePath,
                                                  [[maybe_unused]] const wchar_t* destinationPath,
                                                  [[maybe_unused]] HRESULT status,
                                                  FileSystemIssueAction* action,
                                                  [[maybe_unused]] FileSystemOptions* options,
                                                  [[maybe_unused]] void* cookie) noexcept override
        {
            if (! action)
            {
                return E_POINTER;
            }
            *action = FileSystemIssueAction::Cancel;
            return S_OK;
        }
    };

    const std::filesystem::path srcRoot      = state.tempRoot / L"copyitems-multiroot-src";
    const std::filesystem::path dstRoot      = state.tempRoot / L"copyitems-multiroot-dst";
    const std::filesystem::path lightRootA   = srcRoot / L"light-a";
    const std::filesystem::path dominantRoot = srcRoot / L"dominant-b";
    const std::filesystem::path lightRootC   = srcRoot / L"light-c";
    const std::filesystem::path dominantDir  = dominantRoot / L"dominant-child";

    constexpr int kDominantFileCount = 6;
    constexpr size_t kDominantBytes  = 512ull * 1024ull;

    static_cast<void>(SetPluginConfiguration(
        state.infoLocal.get(),
        R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json"));

    if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(lightRootA) ||
        ! RecreateEmptyDirectory(dominantDir) || ! RecreateEmptyDirectory(lightRootC))
    {
        Fail(L"CopyItems multi-root uneven recursive parallelism failed to reset directories.");
        return true;
    }

    if (! WriteTestFile(lightRootA / L"light-a.bin", 4096) || ! WriteTestFile(lightRootC / L"light-c.bin", 4096))
    {
        Fail(L"CopyItems multi-root uneven recursive parallelism failed to seed light roots.");
        return true;
    }

    for (int i = 0; i < kDominantFileCount; ++i)
    {
        if (! WriteTestFile(dominantDir / std::format(L"dominant_{:02}.bin", i), kDominantBytes))
        {
            Fail(L"CopyItems multi-root uneven recursive parallelism failed to seed dominant files.");
            return true;
        }
    }

    RecordingCallback callback{};
    callback.dominantNeedle = dominantDir.wstring();

    FileSystemOptions options{};
    options.sizeBytes = sizeof(FileSystemOptions);

    const std::array<const wchar_t*, 3> sources{lightRootA.c_str(), dominantRoot.c_str(), lightRootC.c_str()};
    const FileSystemFlags flags =
        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

    const ULONGLONG startedTick = GetTickCount64();
    const HRESULT hr =
        state.fsLocal->CopyItems(sources.data(), static_cast<unsigned long>(sources.size()), dstRoot.c_str(), flags, &options, &callback, nullptr);
    const uint64_t durationUs = (GetTickCount64() >= startedTick) ? (GetTickCount64() - startedTick) * 1000ull : 0ull;

    Debug::Perf::Emit(L"FileOps.SelfTest.CopyItemsMultiRootUnevenRecursiveParallelismUs",
                      L"shape=three-selected-folders-one-dominant-nested-subtree",
                      durationUs,
                      static_cast<uint64_t>(callback.dominantStreamCount),
                      static_cast<uint64_t>(kDominantFileCount),
                      hr);
    Debug::Perf::Emit(L"FileOps.SelfTest.CopyItemsMultiRootDominantProgressStreams",
                      L"shape=three-selected-folders-one-dominant-nested-subtree",
                      static_cast<uint64_t>(callback.dominantStreamCount),
                      callback.dominantProgress,
                      callback.completedCount,
                      hr);

    if (FAILED(hr))
    {
        Fail(std::format(L"CopyItems multi-root uneven recursive parallelism copy failed: 0x{:08X}.", static_cast<unsigned long>(hr)));
        return true;
    }

    const std::filesystem::path dstDominantDir = dstRoot / dominantRoot.filename() / dominantDir.filename();
    const size_t dstDominantCount              = CountFiles(dstDominantDir);
    if (dstDominantCount != static_cast<size_t>(kDominantFileCount))
    {
        Fail(std::format(
            L"CopyItems multi-root uneven recursive parallelism output mismatch: expected {} dominant files, got {}.", kDominantFileCount, dstDominantCount));
        return true;
    }

    if (callback.completedCount != sources.size())
    {
        Fail(std::format(
            L"CopyItems multi-root uneven recursive parallelism expected {} item-completed callbacks, got {}.", sources.size(), callback.completedCount));
        return true;
    }

    if (callback.dominantStreamCount <= 1u)
    {
        Fail(std::format(L"CopyItems multi-root uneven subtree expected multiple dominant progress streams, got {}.", callback.dominantStreamCount));
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase7_CopyRecursiveParallelismMatrix);
    return false;
}
case SelfTestState::Step::Phase7_CopyRecursiveParallelismMatrix:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase7_CopyRecursiveParallelismMatrix timed out.");
        return true;
    }

    static_cast<void>(SetPluginConfiguration(
        state.infoLocal.get(),
        R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"copyReparse"})json"));

    FileSystemOptions options{};
    options.sizeBytes = sizeof(FileSystemOptions);

    const FileSystemFlags copyFlags =
        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
    const auto elapsedUs = [](ULONGLONG startTick) noexcept -> uint64_t
    {
        const ULONGLONG endTick = GetTickCount64();
        return endTick >= startTick ? (endTick - startTick) * 1000ull : 0ull;
    };
    const auto emitShape =
        [&](const wchar_t* name, const wchar_t* shape, uint64_t durationUs, const FileOpsRecursiveProgressRecorder& callback, HRESULT hr) noexcept
    {
        std::scoped_lock lock(callback.mutex);
        Debug::Perf::Emit(name, shape, durationUs, static_cast<uint64_t>(callback.streamCount), static_cast<uint64_t>(callback.matchingStreamCount), hr);
    };

    constexpr int kLargeFileCount = 6;
    constexpr size_t kLargeBytes  = 512ull * 1024ull;

    const std::filesystem::path shallowSrcRoot = state.tempRoot / L"copy-matrix-shallow-src";
    const std::filesystem::path shallowDstRoot = state.tempRoot / L"copy-matrix-shallow-dst";
    const std::filesystem::path shallowPayload = shallowSrcRoot / L"payload";
    if (! RecreateEmptyDirectory(shallowPayload) || ! RecreateEmptyDirectory(shallowDstRoot))
    {
        Fail(L"Copy recursive matrix failed to reset shallow directories.");
        return true;
    }
    for (int i = 0; i < kLargeFileCount; ++i)
    {
        if (! WriteTestFile(shallowPayload / std::format(L"shallow_{:02}.bin", i), kLargeBytes))
        {
            Fail(L"Copy recursive matrix failed to seed shallow files.");
            return true;
        }
    }

    FileOpsRecursiveProgressRecorder shallowCallback{};
    const std::array<const wchar_t*, 1> shallowSources{shallowPayload.c_str()};
    ULONGLONG startedTick = GetTickCount64();
    HRESULT hr            = state.fsLocal->CopyItems(
        shallowSources.data(), static_cast<unsigned long>(shallowSources.size()), shallowDstRoot.c_str(), copyFlags, &options, &shallowCallback, nullptr);
    emitShape(L"FileOps.SelfTest.CopyRecursiveMatrixShallowFiles", L"shape=single-root-many-shallow-files", elapsedUs(startedTick), shallowCallback, hr);
    if (FAILED(hr))
    {
        Fail(std::format(L"Copy recursive matrix shallow copy failed: 0x{:08X}.", static_cast<unsigned long>(hr)));
        return true;
    }
    if (CountFiles(shallowDstRoot / shallowPayload.filename()) != static_cast<size_t>(kLargeFileCount))
    {
        Fail(L"Copy recursive matrix shallow copy output mismatch.");
        return true;
    }
    if (shallowCallback.streamCount <= 1u)
    {
        Fail(std::format(L"Copy recursive matrix shallow copy expected multiple streams, got {}.", shallowCallback.streamCount));
        return true;
    }

    const std::filesystem::path dominantSrcRoot = state.tempRoot / L"copy-matrix-dominant-src";
    const std::filesystem::path dominantDstRoot = state.tempRoot / L"copy-matrix-dominant-dst";
    const std::filesystem::path dominantPayload = dominantSrcRoot / L"payload";
    const std::filesystem::path dominantA       = dominantPayload / L"child-a";
    const std::filesystem::path dominantB       = dominantPayload / L"child-b" / L"dominant";
    const std::filesystem::path dominantC       = dominantPayload / L"child-c";
    if (! RecreateEmptyDirectory(dominantA) || ! RecreateEmptyDirectory(dominantB) || ! RecreateEmptyDirectory(dominantC) ||
        ! RecreateEmptyDirectory(dominantDstRoot))
    {
        Fail(L"Copy recursive matrix failed to reset dominant directories.");
        return true;
    }
    if (! WriteTestFile(dominantA / L"a.bin", 4096) || ! WriteTestFile(dominantC / L"c.bin", 4096))
    {
        Fail(L"Copy recursive matrix failed to seed light dominant siblings.");
        return true;
    }
    for (int i = 0; i < kLargeFileCount; ++i)
    {
        if (! WriteTestFile(dominantB / std::format(L"dominant_{:02}.bin", i), kLargeBytes))
        {
            Fail(L"Copy recursive matrix failed to seed dominant child files.");
            return true;
        }
    }

    FileOpsRecursiveProgressRecorder dominantCallback{};
    dominantCallback.matchingNeedle = dominantB.wstring();
    startedTick                     = GetTickCount64();
    hr                              = state.fsLocal->CopyItem(
        dominantPayload.c_str(), (dominantDstRoot / dominantPayload.filename()).c_str(), copyFlags, &options, &dominantCallback, nullptr);
    emitShape(L"FileOps.SelfTest.CopyRecursiveMatrixDominantSubtree",
              L"shape=single-root-several-child-dirs-one-dominant",
              elapsedUs(startedTick),
              dominantCallback,
              hr);
    if (FAILED(hr))
    {
        Fail(std::format(L"Copy recursive matrix dominant subtree copy failed: 0x{:08X}.", static_cast<unsigned long>(hr)));
        return true;
    }
    if (CountFilesRecursive(dominantDstRoot / dominantPayload.filename()) != static_cast<size_t>(kLargeFileCount + 2))
    {
        Fail(L"Copy recursive matrix dominant subtree output mismatch.");
        return true;
    }
    if (dominantCallback.matchingStreamCount <= 1u)
    {
        Fail(std::format(L"Copy recursive matrix dominant subtree expected multiple matching streams, got {}.", dominantCallback.matchingStreamCount));
        return true;
    }

    const std::filesystem::path reparseSrcRoot = state.tempRoot / L"copy-matrix-reparse-src";
    const std::filesystem::path reparseDstRoot = state.tempRoot / L"copy-matrix-reparse-dst";
    const std::filesystem::path reparsePayload = reparseSrcRoot / L"payload";
    const std::filesystem::path reparseTarget  = state.tempRoot / L"copy-matrix-reparse-target";
    const std::filesystem::path reparseLink    = reparsePayload / L"target-link";
    if (! RecreateEmptyDirectory(reparsePayload) || ! RecreateEmptyDirectory(reparseDstRoot) || ! RecreateEmptyDirectory(reparseTarget))
    {
        Fail(L"Copy recursive matrix failed to reset reparse directories.");
        return true;
    }
    if (! WriteTestFile(reparsePayload / L"before-link.bin", 4096) || ! WriteTestFile(reparseTarget / L"target.bin", 4096))
    {
        Fail(L"Copy recursive matrix failed to seed reparse files.");
        return true;
    }
    if (! TryCreateJunction(reparseLink, reparseTarget))
    {
        Fail(L"Copy recursive matrix failed to create a reparse work item.");
        return true;
    }

    FileOpsRecursiveProgressRecorder reparseCallback{};
    startedTick = GetTickCount64();
    hr = state.fsLocal->CopyItem(reparsePayload.c_str(), (reparseDstRoot / reparsePayload.filename()).c_str(), copyFlags, &options, &reparseCallback, nullptr);
    emitShape(L"FileOps.SelfTest.CopyRecursiveMatrixReparsePoint", L"shape=single-root-child-reparse-copyreparse", elapsedUs(startedTick), reparseCallback, hr);
    if (FAILED(hr))
    {
        Fail(std::format(L"Copy recursive matrix reparse copy failed: 0x{:08X}.", static_cast<unsigned long>(hr)));
        return true;
    }
    const std::filesystem::path copiedReparseLink = reparseDstRoot / reparsePayload.filename() / reparseLink.filename();
    const std::optional<DWORD> copiedReparseTag   = TryGetReparseTag(copiedReparseLink);
    if (! copiedReparseTag.has_value() || (copiedReparseTag.value() != IO_REPARSE_TAG_MOUNT_POINT && copiedReparseTag.value() != IO_REPARSE_TAG_SYMLINK))
    {
        Fail(L"Copy recursive matrix did not preserve the child reparse point with copyReparse policy.");
        return true;
    }

    const std::filesystem::path mixedSrcRoot = state.tempRoot / L"copy-matrix-mixed-src";
    const std::filesystem::path mixedDstRoot = state.tempRoot / L"copy-matrix-mixed-dst";
    const std::filesystem::path mixedFolderA = mixedSrcRoot / L"folder-a";
    const std::filesystem::path mixedHeavy   = mixedFolderA / L"nested-heavy";
    const std::filesystem::path mixedFolderB = mixedSrcRoot / L"folder-b";
    const std::filesystem::path mixedFileA   = mixedSrcRoot / L"root-a.bin";
    if (! RecreateEmptyDirectory(mixedHeavy) || ! RecreateEmptyDirectory(mixedFolderB) || ! RecreateEmptyDirectory(mixedDstRoot))
    {
        Fail(L"Copy recursive matrix failed to reset mixed source directories.");
        return true;
    }
    if (! WriteTestFile(mixedFolderB / L"small.bin", 4096) || ! WriteTestFile(mixedFileA, 4096))
    {
        Fail(L"Copy recursive matrix failed to seed mixed roots.");
        return true;
    }
    for (int i = 0; i < 4; ++i)
    {
        if (! WriteTestFile(mixedHeavy / std::format(L"mixed_{:02}.bin", i), kLargeBytes))
        {
            Fail(L"Copy recursive matrix failed to seed mixed heavy files.");
            return true;
        }
    }

    FileOpsRecursiveProgressRecorder mixedCallback{};
    mixedCallback.matchingNeedle = mixedHeavy.wstring();
    const std::array<const wchar_t*, 3> mixedSources{mixedFolderA.c_str(), mixedFileA.c_str(), mixedFolderB.c_str()};
    startedTick = GetTickCount64();
    hr          = state.fsLocal->CopyItems(
        mixedSources.data(), static_cast<unsigned long>(mixedSources.size()), mixedDstRoot.c_str(), copyFlags, &options, &mixedCallback, nullptr);
    emitShape(L"FileOps.SelfTest.CopyRecursiveMatrixMixedRoots", L"shape=batch-folders-and-files", elapsedUs(startedTick), mixedCallback, hr);
    if (FAILED(hr))
    {
        Fail(std::format(L"Copy recursive matrix mixed roots copy failed: 0x{:08X}.", static_cast<unsigned long>(hr)));
        return true;
    }
    if (CountFilesRecursive(mixedDstRoot) != 6u)
    {
        Fail(L"Copy recursive matrix mixed roots output mismatch.");
        return true;
    }
    if (mixedCallback.completedCount != mixedSources.size())
    {
        Fail(std::format(L"Copy recursive matrix mixed roots expected {} completions, got {}.", mixedSources.size(), mixedCallback.completedCount));
        return true;
    }
    if (mixedCallback.matchingStreamCount <= 1u)
    {
        Fail(std::format(L"Copy recursive matrix mixed roots expected multiple heavy-folder streams, got {}.", mixedCallback.matchingStreamCount));
        return true;
    }

    const std::filesystem::path nestedOneSrcRoot = state.tempRoot / L"copy-matrix-nested-one-src";
    const std::filesystem::path nestedOneDstRoot = state.tempRoot / L"copy-matrix-nested-one-dst";
    std::array<std::filesystem::path, 4> nestedOneSources{};
    if (! RecreateEmptyDirectory(nestedOneSrcRoot) || ! RecreateEmptyDirectory(nestedOneDstRoot))
    {
        Fail(L"Copy recursive matrix failed to reset nested-concurrency-one directories.");
        return true;
    }
    for (size_t rootIndex = 0; rootIndex < nestedOneSources.size(); ++rootIndex)
    {
        const std::filesystem::path root  = nestedOneSrcRoot / std::format(L"root-{}", rootIndex);
        const std::filesystem::path heavy = root / L"nested-heavy";
        nestedOneSources[rootIndex]       = root;
        if (! RecreateEmptyDirectory(heavy))
        {
            Fail(L"Copy recursive matrix failed to create nested-concurrency-one source.");
            return true;
        }
        for (int fileIndex = 0; fileIndex < 3; ++fileIndex)
        {
            if (! WriteTestFile(heavy / std::format(L"nested_one_{}_{:02}.bin", rootIndex, fileIndex), kLargeBytes))
            {
                Fail(L"Copy recursive matrix failed to seed nested-concurrency-one files.");
                return true;
            }
        }
    }

    std::array<const wchar_t*, 4> nestedOneSourcePtrs{};
    for (size_t index = 0; index < nestedOneSources.size(); ++index)
    {
        nestedOneSourcePtrs[index] = nestedOneSources[index].c_str();
    }
    FileOpsRecursiveProgressRecorder nestedOneCallback{};
    nestedOneCallback.matchingNeedle = (nestedOneSources[0] / L"nested-heavy").wstring();
    startedTick                      = GetTickCount64();
    hr                               = state.fsLocal->CopyItems(nestedOneSourcePtrs.data(),
                                                                static_cast<unsigned long>(nestedOneSourcePtrs.size()),
                                                                nestedOneDstRoot.c_str(),
                                                                copyFlags,
                                                                &options,
                                                                &nestedOneCallback,
                                                                nullptr);
    emitShape(
        L"FileOps.SelfTest.CopyRecursiveMatrixNestedConcurrencyOne", L"shape=four-roots-spare-budget-zero", elapsedUs(startedTick), nestedOneCallback, hr);
    if (FAILED(hr))
    {
        Fail(std::format(L"Copy recursive matrix nested-concurrency-one copy failed: 0x{:08X}.", static_cast<unsigned long>(hr)));
        return true;
    }
    if (CountFilesRecursive(nestedOneDstRoot) != 12u)
    {
        Fail(L"Copy recursive matrix nested-concurrency-one output mismatch.");
        return true;
    }
    if (nestedOneCallback.completedCount != nestedOneSourcePtrs.size())
    {
        Fail(std::format(
            L"Copy recursive matrix nested-concurrency-one expected {} completions, got {}.", nestedOneSourcePtrs.size(), nestedOneCallback.completedCount));
        return true;
    }
    if (nestedOneCallback.matchingStreamCount != 1u)
    {
        Fail(std::format(L"Copy recursive matrix nested-concurrency-one expected one matching stream for the spare-budget-zero subtree, got {}.",
                         nestedOneCallback.matchingStreamCount));
        return true;
    }

    const std::filesystem::path moveSrcRoot = state.tempRoot / L"copy-matrix-move-src";
    const std::filesystem::path moveDstRoot = state.tempRoot / L"copy-matrix-move-dst";
    const std::filesystem::path movePayload = moveSrcRoot / L"payload";
    const std::filesystem::path moveHeavy   = movePayload / L"forced-fallback-heavy";
    if (! RecreateEmptyDirectory(moveHeavy) || ! RecreateEmptyDirectory(moveDstRoot))
    {
        Fail(L"Copy recursive matrix failed to reset forced move fallback directories.");
        return true;
    }
    for (int i = 0; i < kLargeFileCount; ++i)
    {
        if (! WriteTestFile(moveHeavy / std::format(L"move_{:02}.bin", i), kLargeBytes))
        {
            Fail(L"Copy recursive matrix failed to seed forced move fallback files.");
            return true;
        }
    }

    DWORD previousRequired = GetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(), nullptr, 0u);
    std::wstring previousValue;
    bool hadPreviousValue = false;
    if (previousRequired > 0)
    {
        previousValue.resize(previousRequired, L'\0');
        const DWORD previousWritten = GetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(), previousValue.data(), previousRequired);
        if (previousWritten > 0 && previousWritten < previousRequired)
        {
            previousValue.resize(previousWritten);
            hadPreviousValue = true;
        }
    }
    auto restoreForcedMoveFallback = wil::scope_exit([&]() noexcept
    { static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(), hadPreviousValue ? previousValue.c_str() : nullptr)); });
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(), L"1"));

    FileOpsRecursiveProgressRecorder moveCallback{};
    moveCallback.matchingNeedle = moveHeavy.wstring();
    startedTick                 = GetTickCount64();
    hr = state.fsLocal->MoveItem(movePayload.c_str(), (moveDstRoot / movePayload.filename()).c_str(), copyFlags, &options, &moveCallback, nullptr);
    emitShape(
        L"FileOps.SelfTest.CopyRecursiveMatrixForcedMoveFallback", L"shape=debug-forced-cross-volume-move-fallback", elapsedUs(startedTick), moveCallback, hr);
    if (FAILED(hr))
    {
        Fail(std::format(L"Copy recursive matrix forced move fallback failed: 0x{:08X}.", static_cast<unsigned long>(hr)));
        return true;
    }
    std::error_code ec;
    if (std::filesystem::exists(movePayload, ec))
    {
        Fail(L"Copy recursive matrix forced move fallback left source directory behind.");
        return true;
    }
    if (CountFilesRecursive(moveDstRoot / movePayload.filename()) != static_cast<size_t>(kLargeFileCount))
    {
        Fail(L"Copy recursive matrix forced move fallback output mismatch.");
        return true;
    }
    if (moveCallback.matchingStreamCount <= 1u)
    {
        Fail(std::format(L"Copy recursive matrix forced move fallback expected multiple matching streams, got {}.", moveCallback.matchingStreamCount));
        return true;
    }
    static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvForceMoveCopyFallback.data(), hadPreviousValue ? previousValue.c_str() : nullptr));

    std::wstring alternateVolumeSkipDetail;
    std::optional<std::filesystem::path> alternateVolumeRoot = TryCreateAlternateWritableVolumeSelfTestRoot(state.tempRoot, alternateVolumeSkipDetail);
    if (! alternateVolumeRoot.has_value())
    {
        Debug::Perf::Emit(L"FileOps.SelfTest.CopyRecursiveMatrixRealCrossVolumeMove", std::format(L"skip={}", alternateVolumeSkipDetail), 0, 0, 0, S_FALSE);
    }
    else
    {
        auto cleanupAlternateVolumeRoot = wil::scope_exit([&]() noexcept
        {
            std::error_code cleanupEc;
            std::filesystem::remove_all(alternateVolumeRoot.value(), cleanupEc);
        });

        const std::filesystem::path realMoveSrcRoot = state.tempRoot / L"copy-matrix-real-cross-volume-src";
        const std::filesystem::path realMovePayload = realMoveSrcRoot / L"payload";
        const std::filesystem::path realMoveHeavy   = realMovePayload / L"real-cross-volume-heavy";
        if (! RecreateEmptyDirectory(realMoveHeavy))
        {
            Fail(L"Copy recursive matrix failed to reset real cross-volume move source.");
            return true;
        }
        for (int i = 0; i < 3; ++i)
        {
            if (! WriteTestFile(realMoveHeavy / std::format(L"real_move_{:02}.bin", i), kLargeBytes))
            {
                Fail(L"Copy recursive matrix failed to seed real cross-volume move files.");
                return true;
            }
        }

        FileOpsRecursiveProgressRecorder realMoveCallback{};
        realMoveCallback.matchingNeedle = realMoveHeavy.wstring();
        startedTick                     = GetTickCount64();
        hr                              = state.fsLocal->MoveItem(
            realMovePayload.c_str(), (alternateVolumeRoot.value() / realMovePayload.filename()).c_str(), copyFlags, &options, &realMoveCallback, nullptr);
        emitShape(
            L"FileOps.SelfTest.CopyRecursiveMatrixRealCrossVolumeMove", L"shape=real-cross-volume-move-fallback", elapsedUs(startedTick), realMoveCallback, hr);
        if (FAILED(hr))
        {
            Fail(std::format(L"Copy recursive matrix real cross-volume move failed: 0x{:08X}.", static_cast<unsigned long>(hr)));
            return true;
        }
        if (std::filesystem::exists(realMovePayload, ec))
        {
            Fail(L"Copy recursive matrix real cross-volume move left source directory behind.");
            return true;
        }
        if (CountFilesRecursive(alternateVolumeRoot.value() / realMovePayload.filename()) != 3u)
        {
            Fail(L"Copy recursive matrix real cross-volume move output mismatch.");
            return true;
        }
    }

    const std::filesystem::path errorSrcRoot = state.tempRoot / L"copy-matrix-error-src";
    const std::filesystem::path errorDstRoot = state.tempRoot / L"copy-matrix-error-dst";
    const std::filesystem::path errorPayload = errorSrcRoot / L"payload";
    const std::filesystem::path lockedFile   = errorPayload / L"locked.bin";
    if (! RecreateEmptyDirectory(errorPayload) || ! RecreateEmptyDirectory(errorDstRoot))
    {
        Fail(L"Copy recursive matrix failed to reset continue-on-error directories.");
        return true;
    }
    if (! WriteTestFile(errorPayload / L"ok.bin", 4096) || ! WriteTestFile(lockedFile, 4096))
    {
        Fail(L"Copy recursive matrix failed to seed continue-on-error files.");
        return true;
    }

    wil::unique_handle lockedHandle(CreateFileW(lockedFile.c_str(), GENERIC_READ, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! lockedHandle)
    {
        Fail(L"Copy recursive matrix failed to open locked source file.");
        return true;
    }

    FileOpsRecursiveProgressRecorder errorCallback{};
    const FileSystemFlags continueFlags = static_cast<FileSystemFlags>(copyFlags | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    startedTick                         = GetTickCount64();
    hr = state.fsLocal->CopyItem(errorPayload.c_str(), (errorDstRoot / errorPayload.filename()).c_str(), continueFlags, &options, &errorCallback, nullptr);
    emitShape(L"FileOps.SelfTest.CopyRecursiveMatrixContinueOnError",
              L"shape=recursive-copy-continue-on-error-locked-file",
              elapsedUs(startedTick),
              errorCallback,
              hr);
    if (hr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        Fail(std::format(L"Copy recursive matrix continue-on-error expected ERROR_PARTIAL_COPY, got 0x{:08X}.", static_cast<unsigned long>(hr)));
        return true;
    }
    if (! std::filesystem::exists(errorDstRoot / errorPayload.filename() / L"ok.bin", ec))
    {
        Fail(L"Copy recursive matrix continue-on-error did not copy the unlocked file.");
        return true;
    }

    const std::filesystem::path cancelSrcRoot = state.tempRoot / L"copy-matrix-cancel-src";
    const std::filesystem::path cancelDstRoot = state.tempRoot / L"copy-matrix-cancel-dst";
    const std::filesystem::path cancelPayload = cancelSrcRoot / L"payload";
    if (! RecreateEmptyDirectory(cancelPayload) || ! RecreateEmptyDirectory(cancelDstRoot))
    {
        Fail(L"Copy recursive matrix failed to reset cancellation directories.");
        return true;
    }
    for (int i = 0; i < 16; ++i)
    {
        if (! WriteTestFile(cancelPayload / std::format(L"cancel_{:02}.bin", i), kLargeBytes))
        {
            Fail(L"Copy recursive matrix failed to seed cancellation files.");
            return true;
        }
    }

    FileOpsRecursiveProgressRecorder cancelCallback{};
    cancelCallback.cancelAfterProgress = true;
    startedTick                        = GetTickCount64();
    hr = state.fsLocal->CopyItem(cancelPayload.c_str(), (cancelDstRoot / cancelPayload.filename()).c_str(), copyFlags, &options, &cancelCallback, nullptr);
    emitShape(L"FileOps.SelfTest.CopyRecursiveMatrixCancel", L"shape=recursive-copy-cancel-while-workers-active", elapsedUs(startedTick), cancelCallback, hr);
    if (hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) && hr != E_ABORT)
    {
        Fail(std::format(L"Copy recursive matrix cancellation expected cancellation HRESULT, got 0x{:08X}.", static_cast<unsigned long>(hr)));
        return true;
    }
    if (cancelCallback.progressCount == 0)
    {
        Fail(L"Copy recursive matrix cancellation did not observe any progress before cancellation.");
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase7_SharedPerItemScheduler);
    return false;
}
case SelfTestState::Step::Phase7_SharedPerItemScheduler:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase7_SharedPerItemScheduler timed out.");
        return true;
    }

    const std::filesystem::path srcDir = state.tempRoot / L"shared-sched-src";
    const std::filesystem::path dstA   = state.tempRoot / L"shared-sched-dst-a";
    const std::filesystem::path dstB   = state.tempRoot / L"shared-sched-dst-b";

    constexpr int kFileCount         = 12;
    constexpr size_t kFileBytes      = 2ull * 1024ull * 1024ull;
    constexpr uint64_t kInitialLimit = 1ull * 1024ull * 1024ull;

    const std::filesystem::path smallFolder = srcDir / L"a_folder";
    const std::filesystem::path slowFolder  = srcDir / L"z_slow_dir";

    const auto maybeLogProgress = [&]() noexcept
    {
        if (state.lastProgressLogTick != 0 && nowTick >= state.lastProgressLogTick && (nowTick - state.lastProgressLogTick) < 1000ull)
        {
            return;
        }
        state.lastProgressLogTick = nowTick;

        size_t selectionCount = 0;
        if (state.folderWindow)
        {
            if (FolderView* fv = TryGetFolderView(state.folderWindow, FolderWindow::Pane::Left))
            {
                selectionCount = fv->GetSelectedOrFocusedPathAttributes().size();
            }
        }

        const auto describeTask = [&](wchar_t name, const std::optional<std::uint64_t>& taskId) noexcept -> std::wstring
        {
            if (! taskId.has_value())
            {
                return std::format(L"{}:none", name);
            }

            const std::uint64_t id = taskId.value();

            using Task = FolderWindow::FileOperationState::Task;
            Task* task = state.fileOps ? state.fileOps->FindTask(id) : nullptr;
            if (task)
            {
                unsigned int maxConc           = 0;
                size_t inFlight                = 0;
                unsigned long completedItems   = 0;
                unsigned long completedFiles   = 0;
                unsigned long completedFolders = 0;
                {
                    std::scoped_lock lock(task->_progressMutex);
                    maxConc          = task->_perItemMaxConcurrency;
                    inFlight         = task->_perItemInFlightCallCount;
                    completedItems   = task->_progressCompletedItems;
                    completedFiles   = task->_completedTopLevelFiles;
                    completedFolders = task->_completedTopLevelFolders;
                }

                const bool started        = task->HasStarted();
                const bool entered        = task->HasEnteredOperation();
                const bool waiting        = task->IsWaitingInQueue();
                const bool queuePaused    = task->IsQueuePaused();
                const bool paused         = task->IsPaused();
                const bool preCalcInProg  = task->_preCalcInProgress.load(std::memory_order_acquire);
                const bool preCalcSkipped = task->_preCalcSkipped.load(std::memory_order_acquire);
                const bool preCalcDone    = task->_preCalcCompleted.load(std::memory_order_acquire);

                return std::format(L"{}:{} started={} entered={} waiting={} qPaused={} paused={} preCalc(inProg={} skipped={} done={}) maxConc={} "
                                   L"preCalcWorkers={} inFlight={} completedItems={} files={} folders={}",
                                   name,
                                   id,
                                   started ? 1 : 0,
                                   entered ? 1 : 0,
                                   waiting ? 1 : 0,
                                   queuePaused ? 1 : 0,
                                   paused ? 1 : 0,
                                   preCalcInProg ? 1 : 0,
                                   preCalcSkipped ? 1 : 0,
                                   preCalcDone ? 1 : 0,
                                   maxConc,
                                   task->_preCalcWorkerCountUsed.load(std::memory_order_acquire),
                                   inFlight,
                                   completedItems,
                                   completedFiles,
                                   completedFolders);
            }

            const auto it = state.completedTasks.find(id);
            if (it != state.completedTasks.end())
            {
                const CompletedTaskInfo& info = it->second;
                return std::format(L"{}:{} completed hr=0x{:08X} started={} preCalcSkipped={} preCalcWorkers={} items={} files={} folders={}",
                                   name,
                                   id,
                                   static_cast<unsigned long>(info.hr),
                                   info.started ? 1 : 0,
                                   info.preCalcSkipped ? 1 : 0,
                                   info.preCalcWorkerCountUsed,
                                   info.progressCompletedItems,
                                   info.completedFiles,
                                   info.completedFolders);
            }

            return std::format(L"{}:{} missing", name, id);
        };

        AppendLog(std::format(L"Phase7_SharedPerItemScheduler dbg stepState={} selection={} {} {}",
                              state.stepState,
                              selectionCount,
                              describeTask(L'A', state.taskA),
                              describeTask(L'B', state.taskB)));
    };
    maybeLogProgress();

    if (state.stepState == 0)
    {
        state.fileOps->ApplyQueueMode(false);
        state.taskA.reset();
        state.taskB.reset();
        state.markerTick          = 0;
        state.baselineThreadCount = 0;
        state.lastProgressLogTick = 0;

        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":8,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"directorySizeDelayMs":1})json"));

        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstA) || ! RecreateEmptyDirectory(dstB))
        {
            Fail(L"Failed to reset shared scheduler directories.");
            return true;
        }

        if (! CreateDeleteTree(slowFolder, 6, 50, 1))
        {
            Fail(L"Failed to create slow directory tree for shared scheduler test.");
            return true;
        }

        if (! RecreateEmptyDirectory(smallFolder))
        {
            Fail(L"Failed to create small folder for shared scheduler test.");
            return true;
        }

        if (! WriteTestFile(smallFolder / L"inside.bin", 1024))
        {
            Fail(L"Failed to write small folder test file.");
            return true;
        }

        for (int i = 0; i < kFileCount; ++i)
        {
            const std::filesystem::path file = srcDir / std::format(L"f_{:02}.bin", i);
            if (! WriteTestFile(file, kFileBytes))
            {
                Fail(L"Failed to write shared scheduler test file.");
                return true;
            }
        }

        if (! state.folderWindow)
        {
            Fail(L"Missing FolderWindow for shared scheduler test.");
            return true;
        }

        FolderView* folderView = TryGetFolderView(state.folderWindow, FolderWindow::Pane::Left);
        if (! folderView)
        {
            Fail(L"Failed to locate left FolderView for shared scheduler test.");
            return true;
        }

        folderView->SetFolderPath(srcDir);

        state.stepState = 1;
        return false;
    }

    if (! state.folderWindow)
    {
        return false;
    }

    FolderView* folderView = TryGetFolderView(state.folderWindow, FolderWindow::Pane::Left);
    if (! folderView)
    {
        return false;
    }
    const auto applySelection = [&]() noexcept
    {
        folderView->SetSelectionByDisplayNamePredicate([&](std::wstring_view displayName) noexcept -> bool
        {
            if (displayName == L"a_folder" || displayName == L"z_slow_dir")
            {
                return true;
            }
            if (displayName.size() >= 6 && displayName.starts_with(L"f_") && displayName.ends_with(L".bin"))
            {
                return true;
            }
            return false;
        });
    };
    const size_t expectedSelectionCount = static_cast<size_t>(kFileCount + 2);

    if (state.stepState == 1)
    {
        applySelection();

        const std::vector<FolderView::PathAttributes> selected = folderView->GetSelectedOrFocusedPathAttributes();
        if (selected.size() != expectedSelectionCount)
        {
            return false;
        }

        std::vector<std::filesystem::path> sourcePaths;
        sourcePaths.reserve(selected.size());
        for (const auto& item : selected)
        {
            sourcePaths.push_back(item.path);
        }

        state.baselineThreadCount = 0;

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                                   FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsLocal,
                                                 std::move(sourcePaths),
                                                 dstA,
                                                 flags,
                                                 false,
                                                 kInitialLimit,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start shared scheduler copy task A.");
            return true;
        }

        state.markerTick = 0;
        state.stepState  = 2;
        return false;
    }

    if (state.stepState == 4)
    {
        if (! state.taskA.has_value() || ! state.taskB.has_value())
        {
            return false;
        }

        const auto itA = state.completedTasks.find(state.taskA.value());
        const auto itB = state.completedTasks.find(state.taskB.value());
        if (itA == state.completedTasks.end() || itB == state.completedTasks.end())
        {
            return false;
        }

        if (! itA->second.preCalcSkipped || ! itB->second.preCalcSkipped)
        {
            Fail(L"Expected shared scheduler tasks to have pre-calc skipped.");
            return true;
        }

        const auto isCancelHr = [](HRESULT hr) noexcept -> bool { return hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT; };

        if (! isCancelHr(itA->second.hr) || ! isCancelHr(itB->second.hr))
        {
            Fail(std::format(L"Expected shared scheduler tasks to be cancelled. A=0x{:08X} B=0x{:08X}",
                             static_cast<unsigned long>(itA->second.hr),
                             static_cast<unsigned long>(itB->second.hr)));
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase7_ParallelDeleteKnobs);
        return false;
    }

    if (! state.taskA.has_value())
    {
        return false;
    }

    using Task  = FolderWindow::FileOperationState::Task;
    Task* taskA = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
    Task* taskB = state.taskB.has_value() && state.fileOps ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
    if (! taskA)
    {
        return false;
    }

    if (taskA->_preCalcInProgress.load(std::memory_order_acquire) && ! taskA->_preCalcSkipped.load(std::memory_order_acquire))
    {
        taskA->SkipPreCalculation();
    }
    if (taskB && taskB->_preCalcInProgress.load(std::memory_order_acquire) && ! taskB->_preCalcSkipped.load(std::memory_order_acquire))
    {
        taskB->SkipPreCalculation();
    }

    if (state.stepState == 2)
    {
        if (state.markerTick == 0 && taskA->HasStarted())
        {
            state.markerTick = nowTick;
        }

        unsigned int maxConcA           = 0;
        size_t inFlightA                = 0;
        unsigned long completedItemsA   = 0;
        unsigned long completedFilesA   = 0;
        unsigned long completedFoldersA = 0;
        {
            std::scoped_lock lock(taskA->_progressMutex);
            maxConcA          = taskA->_perItemMaxConcurrency;
            inFlightA         = taskA->_perItemInFlightCallCount;
            completedItemsA   = taskA->_progressCompletedItems;
            completedFilesA   = taskA->_completedTopLevelFiles;
            completedFoldersA = taskA->_completedTopLevelFolders;
        }

        if (maxConcA <= 1u)
        {
            return false;
        }

        if (inFlightA <= 1u)
        {
            if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
            {
                Fail(L"Expected >1 in-flight per-item calls for task A but did not observe them.");
                return true;
            }
            return false;
        }

        if (! state.taskB.has_value())
        {
            state.baselineThreadCount = GetProcessThreadCount();
            if (state.baselineThreadCount == 0)
            {
                Fail(L"Failed to snapshot process thread count after starting task A.");
                return true;
            }

            applySelection();
            const std::vector<FolderView::PathAttributes> selected = folderView->GetSelectedOrFocusedPathAttributes();
            if (selected.size() != expectedSelectionCount)
            {
                return false;
            }

            std::vector<std::filesystem::path> sourcePaths;
            sourcePaths.reserve(selected.size());
            for (const auto& item : selected)
            {
                sourcePaths.push_back(item.path);
            }

            const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                                       FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

            state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                     FILESYSTEM_COPY,
                                                     FolderWindow::Pane::Left,
                                                     FolderWindow::Pane::Right,
                                                     state.fsLocal,
                                                     std::move(sourcePaths),
                                                     dstB,
                                                     flags,
                                                     false,
                                                     kInitialLimit,
                                                     FolderWindow::FileOperationState::ExecutionMode::PerItem);
            if (! state.taskB.has_value())
            {
                Fail(L"Failed to start shared scheduler copy task B.");
                return true;
            }

            state.markerTick = 0;
            state.stepState  = 3;
            return false;
        }
    }

    if (state.stepState == 3)
    {
        if (! state.taskB.has_value())
        {
            return false;
        }

        taskB = state.fileOps ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
        if (! taskB)
        {
            return false;
        }

        if (taskB->_preCalcInProgress.load(std::memory_order_acquire) && ! taskB->_preCalcSkipped.load(std::memory_order_acquire))
        {
            taskB->SkipPreCalculation();
        }

        if (state.markerTick == 0 && taskA->HasStarted() && taskB->HasStarted())
        {
            state.markerTick = nowTick;
        }

        unsigned int maxConcA           = 0;
        unsigned int maxConcB           = 0;
        size_t inFlightA                = 0;
        size_t inFlightB                = 0;
        unsigned long completedItemsA   = 0;
        unsigned long completedItemsB   = 0;
        unsigned long completedFilesA   = 0;
        unsigned long completedFilesB   = 0;
        unsigned long completedFoldersA = 0;
        unsigned long completedFoldersB = 0;
        {
            std::scoped_lock lock(taskA->_progressMutex);
            maxConcA          = taskA->_perItemMaxConcurrency;
            inFlightA         = taskA->_perItemInFlightCallCount;
            completedItemsA   = taskA->_progressCompletedItems;
            completedFilesA   = taskA->_completedTopLevelFiles;
            completedFoldersA = taskA->_completedTopLevelFolders;
        }
        {
            std::scoped_lock lock(taskB->_progressMutex);
            maxConcB          = taskB->_perItemMaxConcurrency;
            inFlightB         = taskB->_perItemInFlightCallCount;
            completedItemsB   = taskB->_progressCompletedItems;
            completedFilesB   = taskB->_completedTopLevelFiles;
            completedFoldersB = taskB->_completedTopLevelFolders;
        }

        if (maxConcA <= 1u || maxConcB <= 1u)
        {
            return false;
        }

        if (inFlightA == 0u || inFlightB == 0u)
        {
            if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
            {
                Fail(L"Expected both tasks to have in-flight per-item calls but did not observe them.");
                return true;
            }
            return false;
        }

        if (state.baselineThreadCount != 0)
        {
            const size_t threadsNow = GetProcessThreadCount();
            if (threadsNow == 0)
            {
                Fail(L"Failed to read process thread count during shared scheduler test.");
                return true;
            }

            const size_t delta                       = (threadsNow >= state.baselineThreadCount) ? (threadsNow - state.baselineThreadCount) : 0;
            constexpr size_t kMaxExpectedThreadDelta = FolderWindow::FileOperationState::Task::kMaxInFlightFiles;
            if (delta > kMaxExpectedThreadDelta)
            {
                Fail(std::format(L"Shared scheduler thread delta too high after starting task B: baseline={} now={} delta={}.",
                                 state.baselineThreadCount,
                                 threadsNow,
                                 delta));
                return true;
            }

            state.baselineThreadCount = 0;
        }

        const bool skippedA = taskA->_preCalcSkipped.load(std::memory_order_acquire);
        const bool skippedB = taskB->_preCalcSkipped.load(std::memory_order_acquire);
        if (! skippedA || ! skippedB)
        {
            return false;
        }

        if (completedItemsA == 0 || completedItemsB == 0)
        {
            return false;
        }

        const uint64_t totalA = static_cast<uint64_t>(completedFilesA) + static_cast<uint64_t>(completedFoldersA);
        const uint64_t totalB = static_cast<uint64_t>(completedFilesB) + static_cast<uint64_t>(completedFoldersB);
        if (totalA != completedItemsA || totalB != completedItemsB)
        {
            Fail(std::format(L"Skipped pre-calc counts mismatch: A items={} files={} folders={} / B items={} files={} folders={}",
                             completedItemsA,
                             completedFilesA,
                             completedFoldersA,
                             completedItemsB,
                             completedFilesB,
                             completedFoldersB));
            return true;
        }

        taskA->RequestCancel();
        taskB->RequestCancel();
        state.stepState = 4;
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase7_ParallelDeleteKnobs:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase7_ParallelDeleteKnobs timed out.");
        return true;
    }

    const std::array<unsigned, 2> concurrencies{1u, 8u};
    const std::filesystem::path delRoot = state.tempRoot / L"delete-knob-tree";

    if (state.stepState == 0)
    {
        state.deleteKnobIndex = 0;
        state.stepState       = 1;
    }

    if (state.stepState == 1)
    {
        if (state.deleteKnobIndex >= concurrencies.size())
        {
            NextStep(state, SelfTestState::Step::Phase7_RecycleBinBatchDelete);
            return false;
        }

        const unsigned conc      = concurrencies[state.deleteKnobIndex];
        const std::string config = std::format(
            R"json({{"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":{},"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048}})json",
            conc);
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), config));

        if (! CreateDeleteTree(delRoot, 6, 30, 16 * 1024))
        {
            Fail(L"Failed to create delete-knob-tree.");
            return true;
        }

        const FileSystemFlags flags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        state.taskA =
            StartFileOperationAndGetId(state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {delRoot}, {}, flags, false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start delete task for knob test.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        if (state.completedTasks.find(state.taskA.value()) == state.completedTasks.end())
        {
            return false;
        }

        if (std::filesystem::exists(delRoot))
        {
            Fail(L"delete-knob-tree still exists after delete task completed.");
            return true;
        }

        ++state.deleteKnobIndex;
        state.stepState = 1;
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase7_RecycleBinBatchDelete:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase7_RecycleBinBatchDelete timed out.");
        return true;
    }

    const std::filesystem::path batchRoot      = state.tempRoot / L"recyclebin-batch";
    constexpr int kBatchFileCount              = 384;
    constexpr size_t kBatchFileBytes           = 4096;
    constexpr unsigned int kBaselineBatchSize  = 1u;
    constexpr unsigned int kCandidateBatchSize = 500u;

    const auto applyRecycleBinBatchConfig = [&](unsigned int batchSize, std::wstring_view label) noexcept -> bool
    {
        const std::string config = std::format(
            R"json({{"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":1,"recycleBinBatchSize":{},"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"copyReparse","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4}})json",
            batchSize);
        if (! SetPluginConfiguration(state.infoLocal.get(), config))
        {
            Fail(std::format(L"Failed to apply local plugin config for {}.", label));
            return false;
        }

        std::string currentConfig;
        if (! BackupPluginConfiguration(state.infoLocal.get(), currentConfig))
        {
            Fail(std::format(L"Failed to read local plugin config back for {}.", label));
            return false;
        }

        const std::string expectedToken = std::format("\"recycleBinBatchSize\":{}", batchSize);
        if (currentConfig.find(expectedToken) == std::string::npos)
        {
            Fail(std::format(L"Local plugin config did not persist recycleBinBatchSize={} for {}.", batchSize, label));
            return false;
        }

        return true;
    };

    const auto startRecycleBinDelete = [&](std::optional<std::uint64_t>& taskSlot, std::wstring_view label) noexcept -> bool
    {
        std::vector<std::filesystem::path> sourcePaths;
        if (! CreateSiblingFiles(batchRoot, kBatchFileCount, kBatchFileBytes, sourcePaths))
        {
            Fail(std::format(L"Failed to create sibling files for {}.", label));
            return false;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_USE_RECYCLE_BIN | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        taskSlot                    = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, std::move(sourcePaths), {}, flags, false);
        if (! taskSlot.has_value())
        {
            Fail(std::format(L"Failed to start recycle-bin delete task for {}.", label));
            return false;
        }

        auto* task = state.fileOps->FindTask(taskSlot.value());
        if (! task)
        {
            Fail(std::format(L"Recycle-bin delete task for {} was not discoverable immediately after start.", label));
            return false;
        }

        if ((task->_flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) == 0)
        {
            Fail(std::format(L"Recycle-bin delete task for {} did not carry the recycle-bin flag.", label));
            return false;
        }

        return true;
    };

    const auto finalizeRecycleBinDelete = [&](const std::optional<std::uint64_t>& taskSlot, uint64_t& durationOut, std::wstring_view label) noexcept -> int
    {
        if (! taskSlot.has_value())
        {
            Fail(std::format(L"Recycle-bin delete task for {} was never started.", label));
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
            Fail(std::format(L"Recycle-bin delete task for {} failed: hr=0x{:08X}.", label, static_cast<unsigned long>(completion.hr)));
            return -1;
        }

        if (! completion.started)
        {
            Fail(std::format(L"Recycle-bin delete task for {} never reported started.", label));
            return -1;
        }

        std::error_code ec;
        if (! std::filesystem::exists(batchRoot, ec))
        {
            Fail(std::format(L"Recycle-bin batch root folder disappeared during {}.", label));
            return -1;
        }

        ec.clear();
        if (! std::filesystem::is_empty(batchRoot, ec))
        {
            Fail(std::format(L"Recycle-bin batch root folder still contains entries after {} completed.", label));
            return -1;
        }

        if (completion.progressCompletedItems < static_cast<unsigned long>(kBatchFileCount))
        {
            Fail(std::format(
                L"Recycle-bin delete task for {} completed only {} items; expected at least {}.", label, completion.progressCompletedItems, kBatchFileCount));
            return -1;
        }

        durationOut = (state.recycleBinBatchRunStartTick != 0 && completion.completionTick >= state.recycleBinBatchRunStartTick)
                          ? static_cast<uint64_t>(completion.completionTick - state.recycleBinBatchRunStartTick) * 1000ull
                          : 0ull;
        if (durationOut == 0)
        {
            Fail(std::format(L"Recycle-bin delete task for {} captured zero duration.", label));
            return -1;
        }

        return 1;
    };

    if (state.stepState == 0)
    {
        state.taskA.reset();
        state.taskB.reset();
        state.recycleBinBatchRunStartTick = 0;
        state.recycleBinBatchBaselineUs   = 0;
        state.recycleBinBatchCandidateUs  = 0;

        if (! applyRecycleBinBatchConfig(kBaselineBatchSize, L"recycle-bin baseline"))
        {
            return true;
        }

        if (! startRecycleBinDelete(state.taskA, L"recycle-bin baseline"))
        {
            return true;
        }

        state.recycleBinBatchRunStartTick = nowTick;
        state.stepState                   = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        const int baselineStatus = finalizeRecycleBinDelete(state.taskA, state.recycleBinBatchBaselineUs, L"recycle-bin baseline");
        if (baselineStatus < 0)
        {
            return true;
        }
        if (baselineStatus == 0)
        {
            return false;
        }

        if (! applyRecycleBinBatchConfig(kCandidateBatchSize, L"recycle-bin candidate"))
        {
            return true;
        }

        if (! startRecycleBinDelete(state.taskB, L"recycle-bin candidate"))
        {
            return true;
        }

        state.recycleBinBatchRunStartTick = nowTick;
        state.stepState                   = 2;
        return false;
    }

    const int candidateStatus = finalizeRecycleBinDelete(state.taskB, state.recycleBinBatchCandidateUs, L"recycle-bin candidate");
    if (candidateStatus < 0)
    {
        return true;
    }
    if (candidateStatus == 0)
    {
        return false;
    }

    const std::wstring detail = std::format(L"fileCount={} fileBytes={} baselineBatchSize={} candidateBatchSize={} root={}",
                                            kBatchFileCount,
                                            kBatchFileBytes,
                                            kBaselineBatchSize,
                                            kCandidateBatchSize,
                                            batchRoot.wstring());
    Debug::Perf::Emit(L"FileOps.SelfTest.RecycleBinBatchBaseline", detail, state.recycleBinBatchBaselineUs, kBatchFileCount, kBaselineBatchSize, S_OK);
    Debug::Perf::Emit(L"FileOps.SelfTest.RecycleBinBatchCandidate", detail, state.recycleBinBatchCandidateUs, kBatchFileCount, kCandidateBatchSize, S_OK);

    const uint64_t improvementUs =
        (state.recycleBinBatchBaselineUs > state.recycleBinBatchCandidateUs) ? (state.recycleBinBatchBaselineUs - state.recycleBinBatchCandidateUs) : 0ull;
    Debug::Perf::Emit(
        L"FileOps.SelfTest.RecycleBinBatchImprovement", detail, improvementUs, state.recycleBinBatchBaselineUs, state.recycleBinBatchCandidateUs, S_OK);

    AppendLog(std::format(L"Phase7_RecycleBinBatchDelete baseline={}us candidate={}us batchSize1={} batchSize2={}",
                          state.recycleBinBatchBaselineUs,
                          state.recycleBinBatchCandidateUs,
                          kBaselineBatchSize,
                          kCandidateBatchSize));

    if (state.recycleBinBatchCandidateUs >= state.recycleBinBatchBaselineUs)
    {
        Fail(std::format(L"Expected recycle-bin batching to beat the batchSize=1 baseline, but baseline={}us and candidate={}us.",
                         state.recycleBinBatchBaselineUs,
                         state.recycleBinBatchCandidateUs));
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase7_RecycleBinBatchDeleteMultiBatch);
    return false;
}
case SelfTestState::Step::Phase7_RecycleBinBatchDeleteMultiBatch:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase7_RecycleBinBatchDeleteMultiBatch timed out.");
        return true;
    }

    const std::filesystem::path batchRoot       = state.tempRoot / L"recyclebin-batch-multibatch";
    constexpr int kBatchFileCount               = 768;
    constexpr size_t kBatchFileBytes            = 4096;
    constexpr unsigned int kConfiguredBatchSize = 256u;

    const auto applyRecycleBinBatchConfig = [&](unsigned int batchSize) noexcept -> bool
    {
        const std::string config = std::format(
            R"json({{"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":1,"recycleBinBatchSize":{},"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"reparsePointPolicy":"copyReparse","searchBackendPreference":"auto","searchMaxDirectoryWalkers":4}})json",
            batchSize);
        if (! SetPluginConfiguration(state.infoLocal.get(), config))
        {
            Fail(L"Failed to apply local plugin config for recycle-bin multi-batch validation.");
            return false;
        }

        std::string currentConfig;
        if (! BackupPluginConfiguration(state.infoLocal.get(), currentConfig))
        {
            Fail(L"Failed to read local plugin config back for recycle-bin multi-batch validation.");
            return false;
        }

        const std::string expectedToken = std::format("\"recycleBinBatchSize\":{}", batchSize);
        if (currentConfig.find(expectedToken) == std::string::npos)
        {
            Fail(std::format(L"Local plugin config did not persist recycleBinBatchSize={} for recycle-bin multi-batch validation.", batchSize));
            return false;
        }

        return true;
    };

    if (state.stepState == 0)
    {
        if (! applyRecycleBinBatchConfig(kConfiguredBatchSize))
        {
            return true;
        }

        std::vector<std::filesystem::path> sourcePaths;
        if (! CreateSiblingFiles(batchRoot, kBatchFileCount, kBatchFileBytes, sourcePaths))
        {
            Fail(L"Failed to create sibling files for recycle-bin multi-batch delete validation.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_USE_RECYCLE_BIN | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
        state.taskA                 = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, std::move(sourcePaths), {}, flags, false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start recycle-bin multi-batch delete task.");
            return true;
        }

        auto* task = state.fileOps->FindTask(state.taskA.value());
        if (! task)
        {
            Fail(L"Recycle-bin multi-batch delete task was not discoverable immediately after start.");
            return true;
        }

        if ((task->_flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) == 0)
        {
            Fail(L"Recycle-bin multi-batch delete task did not carry the recycle-bin flag.");
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

    const CompletedTaskInfo& completion = completionIt->second;
    if (FAILED(completion.hr))
    {
        Fail(std::format(L"Recycle-bin multi-batch delete task failed: hr=0x{:08X}.", static_cast<unsigned long>(completion.hr)));
        return true;
    }

    if (! completion.started)
    {
        Fail(L"Recycle-bin multi-batch delete task never reported started.");
        return true;
    }

    std::error_code ec;
    if (! std::filesystem::exists(batchRoot, ec))
    {
        Fail(L"Recycle-bin multi-batch root folder disappeared; the test expects only the sibling files to be recycled.");
        return true;
    }

    ec.clear();
    if (! std::filesystem::is_empty(batchRoot, ec))
    {
        Fail(L"Recycle-bin multi-batch root folder still contains entries after the task completed.");
        return true;
    }

    if (completion.progressCompletedItems < static_cast<unsigned long>(kBatchFileCount))
    {
        Fail(std::format(
            L"Recycle-bin multi-batch delete task completed only {} items; expected at least {}.", completion.progressCompletedItems, kBatchFileCount));
        return true;
    }

    const std::wstring detail =
        std::format(L"fileCount={} fileBytes={} configuredBatchSize={} root={}", kBatchFileCount, kBatchFileBytes, kConfiguredBatchSize, batchRoot.wstring());
    Debug::Perf::Emit(
        L"FileOps.SelfTest.RecycleBinBatchConfiguredSize", detail, kConfiguredBatchSize, kBatchFileCount, completion.progressCompletedItems, S_OK);

    if (! state.localConfigOriginal.empty() && ! SetPluginConfiguration(state.infoLocal.get(), state.localConfigOriginal))
    {
        Fail(L"Failed to restore local plugin config after recycle-bin multi-batch validation.");
        return true;
    }
    state.localConfigDirty = false;

    NextStep(state, SelfTestState::Step::Phase8_DefaultBandwidthLimitFromSettings);
    return false;
}
case SelfTestState::Step::Phase8_DefaultBandwidthLimitFromSettings:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        if (g_settings.fileOperations.has_value())
        {
            g_settings.fileOperations->defaultBandwidthLimitBytesPerSecond = 0;
        }
        if (state.infoDummy && ! state.defaultSpeedLimitDummyConfigSnapshot.empty())
        {
            static_cast<void>(SetPluginConfiguration(state.infoDummy.get(), state.defaultSpeedLimitDummyConfigSnapshot));
            state.defaultSpeedLimitDummyConfigSnapshot.clear();
        }
        Fail(L"Phase8_DefaultBandwidthLimitFromSettings timed out.");
        return true;
    }

    constexpr uint64_t kFileBytes                             = 4ull * 1024ull * 1024ull;
    constexpr uint64_t kDefaultBandwidthLimitBytesPerSecond   = 1ull * 1024ull * 1024ull;
    const std::wstring dummySourceRoot                        = L"/default-bandwidth-src";
    const std::wstring dummyBaselineDestinationRoot           = L"/default-bandwidth-dst-baseline";
    const std::wstring dummyCandidateDestinationRoot          = L"/default-bandwidth-dst-candidate";
    const std::filesystem::path dummySourceFile               = std::filesystem::path(dummySourceRoot) / L"payload.bin";
    const std::filesystem::path dummyBaselineDestinationFile  = std::filesystem::path(dummyBaselineDestinationRoot) / L"payload.bin";
    const std::filesystem::path dummyCandidateDestinationFile = std::filesystem::path(dummyCandidateDestinationRoot) / L"payload.bin";

    const auto restoreDefaultBandwidthState = [&]() noexcept
    {
        if (g_settings.fileOperations.has_value())
        {
            g_settings.fileOperations->defaultBandwidthLimitBytesPerSecond = 0;
        }
        if (state.infoDummy && ! state.defaultSpeedLimitDummyConfigSnapshot.empty())
        {
            static_cast<void>(SetPluginConfiguration(state.infoDummy.get(), state.defaultSpeedLimitDummyConfigSnapshot));
            state.defaultSpeedLimitDummyConfigSnapshot.clear();
        }
    };

    wil::com_ptr<IFileSystemIO> dummyIo;
    if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
    {
        restoreDefaultBandwidthState();
        Fail(L"Dummy filesystem does not support IFileSystemIO for default bandwidth validation.");
        return true;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

    if (state.stepState == 0)
    {
        state.taskA.reset();
        state.taskB.reset();
        state.defaultSpeedLimitBaselineUs   = 0;
        state.defaultSpeedLimitCandidateUs  = 0;
        state.defaultSpeedLimitRunStartTick = 0;
        state.defaultSpeedLimitDummyConfigSnapshot.clear();

        if (! g_settings.fileOperations.has_value())
        {
            g_settings.fileOperations.emplace();
        }
        g_settings.fileOperations->defaultBandwidthLimitBytesPerSecond = 0;

        if (! BackupPluginConfiguration(state.infoDummy.get(), state.defaultSpeedLimitDummyConfigSnapshot))
        {
            restoreDefaultBandwidthState();
            Fail(L"Failed to snapshot dummy configuration for default bandwidth validation.");
            return true;
        }

        constexpr std::string_view kDummyConfig =
            R"json({"maxChildrenPerDirectory":42,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json";
        if (! SetPluginConfiguration(state.infoDummy.get(), std::string(kDummyConfig)))
        {
            restoreDefaultBandwidthState();
            Fail(L"Failed to apply deterministic dummy configuration for default bandwidth validation.");
            return true;
        }

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummySourceRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyBaselineDestinationRoot) ||
            ! EnsureDummyFolderExists(state.fsDummy.get(), dummyCandidateDestinationRoot))
        {
            restoreDefaultBandwidthState();
            Fail(L"Failed to create dummy folders for default bandwidth validation.");
            return true;
        }

        if (! WritePatternFileFsIo(dummyIo, dummySourceFile, kFileBytes))
        {
            restoreDefaultBandwidthState();
            Fail(L"Failed to seed dummy source file for default bandwidth validation.");
            return true;
        }

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Right,
                                                 FolderWindow::Pane::Left,
                                                 state.fsDummy,
                                                 {dummySourceFile},
                                                 std::filesystem::path(dummyBaselineDestinationRoot),
                                                 flags,
                                                 false,
                                                 0,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 state.fsDummy);
        if (! state.taskA.has_value())
        {
            restoreDefaultBandwidthState();
            Fail(L"Failed to start baseline default-bandwidth copy.");
            return true;
        }

        state.defaultSpeedLimitRunStartTick = nowTick;
        state.stepState                     = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        if (auto* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr)
        {
            const uint64_t currentLimit = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
            if (currentLimit != 0)
            {
                restoreDefaultBandwidthState();
                Fail(std::format(L"Baseline task should inherit unlimited speed, observed {} B/s.", currentLimit));
                return true;
            }
        }

        const auto completionIt = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (completionIt == state.completedTasks.end())
        {
            return false;
        }

        const CompletedTaskInfo& completion = completionIt->second;
        if (FAILED(completion.hr))
        {
            restoreDefaultBandwidthState();
            Fail(std::format(L"Baseline default-bandwidth copy failed: 0x{:08X}.", static_cast<unsigned long>(completion.hr)));
            return true;
        }

        uint64_t copiedSize = 0;
        if (! GetFileSizeFsIo(dummyIo, dummyBaselineDestinationFile, copiedSize) || copiedSize != kFileBytes)
        {
            restoreDefaultBandwidthState();
            Fail(L"Baseline default-bandwidth copy produced the wrong output size.");
            return true;
        }

        state.defaultSpeedLimitBaselineUs = (state.defaultSpeedLimitRunStartTick != 0 && completion.completionTick >= state.defaultSpeedLimitRunStartTick)
                                                ? static_cast<uint64_t>(completion.completionTick - state.defaultSpeedLimitRunStartTick) * 1000ull
                                                : 0ull;

        g_settings.fileOperations->defaultBandwidthLimitBytesPerSecond = kDefaultBandwidthLimitBytesPerSecond;
        state.taskB                                                    = StartFileOperationAndGetId(state.fileOps,
                                                                                                    FILESYSTEM_COPY,
                                                                                                    FolderWindow::Pane::Right,
                                                                                                    FolderWindow::Pane::Left,
                                                                                                    state.fsDummy,
                                                                                                    {dummySourceFile},
                                                                                                    std::filesystem::path(dummyCandidateDestinationRoot),
                                                                                                    flags,
                                                                                                    false,
                                                                                                    0,
                                                                                                    FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                                                    false,
                                                                                                    state.fsDummy);
        if (! state.taskB.has_value())
        {
            restoreDefaultBandwidthState();
            Fail(L"Failed to start candidate default-bandwidth copy.");
            return true;
        }

        state.defaultSpeedLimitRunStartTick = nowTick;
        state.stepState                     = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        auto* task = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
        if (task)
        {
            const uint64_t currentLimit = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
            if (currentLimit != kDefaultBandwidthLimitBytesPerSecond)
            {
                restoreDefaultBandwidthState();
                Fail(std::format(L"Candidate task should inherit {} B/s, observed {} B/s.", kDefaultBandwidthLimitBytesPerSecond, currentLimit));
                return true;
            }
        }

        const auto completionIt = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
        if (completionIt == state.completedTasks.end())
        {
            return false;
        }

        const CompletedTaskInfo& completion = completionIt->second;
        if (FAILED(completion.hr))
        {
            restoreDefaultBandwidthState();
            Fail(std::format(L"Candidate default-bandwidth copy failed: 0x{:08X}.", static_cast<unsigned long>(completion.hr)));
            return true;
        }

        uint64_t copiedSize = 0;
        if (! GetFileSizeFsIo(dummyIo, dummyCandidateDestinationFile, copiedSize) || copiedSize != kFileBytes)
        {
            restoreDefaultBandwidthState();
            Fail(L"Candidate default-bandwidth copy produced the wrong output size.");
            return true;
        }

        state.defaultSpeedLimitCandidateUs = (state.defaultSpeedLimitRunStartTick != 0 && completion.completionTick >= state.defaultSpeedLimitRunStartTick)
                                                 ? static_cast<uint64_t>(completion.completionTick - state.defaultSpeedLimitRunStartTick) * 1000ull
                                                 : 0ull;

        const std::wstring detail = std::format(L"fileBytes={} defaultLimit={} source={} baselineDestination={} candidateDestination={}",
                                                kFileBytes,
                                                kDefaultBandwidthLimitBytesPerSecond,
                                                dummySourceFile.native(),
                                                dummyBaselineDestinationRoot,
                                                dummyCandidateDestinationRoot);
        Debug::Perf::Emit(L"FileOps.SelfTest.DefaultBandwidthLimitBaseline", detail, state.defaultSpeedLimitBaselineUs, kFileBytes, 0, S_OK);
        Debug::Perf::Emit(L"FileOps.SelfTest.DefaultBandwidthLimitCandidate",
                          detail,
                          state.defaultSpeedLimitCandidateUs,
                          kFileBytes,
                          kDefaultBandwidthLimitBytesPerSecond,
                          S_OK);
        const uint64_t slowdownUs = (state.defaultSpeedLimitCandidateUs > state.defaultSpeedLimitBaselineUs)
                                        ? (state.defaultSpeedLimitCandidateUs - state.defaultSpeedLimitBaselineUs)
                                        : 0ull;
        Debug::Perf::Emit(
            L"FileOps.SelfTest.DefaultBandwidthLimitSlowdown", detail, slowdownUs, state.defaultSpeedLimitBaselineUs, state.defaultSpeedLimitCandidateUs, S_OK);

        AppendLog(std::format(L"Phase8_DefaultBandwidthLimitFromSettings baseline={}us candidate={}us defaultLimit={}B/s",
                              state.defaultSpeedLimitBaselineUs,
                              state.defaultSpeedLimitCandidateUs,
                              kDefaultBandwidthLimitBytesPerSecond));

        if (state.defaultSpeedLimitBaselineUs == 0 || state.defaultSpeedLimitCandidateUs == 0)
        {
            restoreDefaultBandwidthState();
            Fail(L"Default bandwidth validation captured zero duration.");
            return true;
        }

        const uint64_t expectedMinCandidateUs = (((kFileBytes * 1000000ull) / kDefaultBandwidthLimitBytesPerSecond) * 70ull) / 100ull;
        if (state.defaultSpeedLimitCandidateUs < expectedMinCandidateUs)
        {
            restoreDefaultBandwidthState();
            Fail(std::format(L"Default bandwidth limit did not slow the candidate enough: candidate={}us expectedAtLeast={}us.",
                             state.defaultSpeedLimitCandidateUs,
                             expectedMinCandidateUs));
            return true;
        }

        if (state.defaultSpeedLimitCandidateUs <= (state.defaultSpeedLimitBaselineUs + 1'500'000ull))
        {
            restoreDefaultBandwidthState();
            Fail(std::format(L"Default bandwidth limit slowdown too small: baseline={}us candidate={}us.",
                             state.defaultSpeedLimitBaselineUs,
                             state.defaultSpeedLimitCandidateUs));
            return true;
        }

        restoreDefaultBandwidthState();
        NextStep(state, SelfTestState::Step::Phase8_TightDefaults_NoOverwrite);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase8_TightDefaults_NoOverwrite:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        Fail(L"Phase8_TightDefaults_NoOverwrite timed out.");
        return true;
    }

    const std::filesystem::path srcDir  = state.tempRoot / L"defaults-src";
    const std::filesystem::path dstDir  = state.tempRoot / L"defaults-dst";
    const std::filesystem::path srcFile = srcDir / L"conflict.bin";
    const std::filesystem::path dstFile = dstDir / L"conflict.bin";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
        {
            Fail(L"Failed to reset defaults-src/defaults-dst directories.");
            return true;
        }

        if (! WriteTestFile(srcFile, 4096) || ! WriteTestFile(dstFile, 8192))
        {
            Fail(L"Failed to write conflict test files.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                 = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {srcFile}, dstDir, flags, false);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start no-overwrite copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        const auto it = state.completedTasks.find(state.taskA.value());
        if (it == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT expectedHr = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        if (it->second.hr != expectedHr)
        {
            Fail(std::format(L"Expected no-overwrite copy to fail with 0x{:08X}, got 0x{:08X}.",
                             static_cast<unsigned long>(expectedHr),
                             static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        std::error_code ec;
        const auto size = std::filesystem::file_size(dstFile, ec);
        if (ec || size != 8192u)
        {
            Fail(L"Destination file size changed despite no-overwrite copy failure.");
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase8_InvalidDestinationRejected);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase8_InvalidDestinationRejected:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        Fail(L"Phase8_InvalidDestinationRejected timed out.");
        return true;
    }

    const std::filesystem::path srcDir   = state.tempRoot / L"invalid-dest-src";
    const std::filesystem::path childDir = srcDir / L"child";
    const std::filesystem::path srcFile  = srcDir / L"ok.bin";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(childDir))
        {
            Fail(L"Failed to reset invalid-dest-src/child directories.");
            return true;
        }

        if (! WriteTestFile(srcFile, 4096))
        {
            Fail(L"Failed to write invalid destination test file.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        const auto taskId           = StartFileOperationAndGetId(
            state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {srcDir}, childDir, flags, false);
        if (taskId.has_value())
        {
            Fail(L"Expected invalid destination copy to be rejected, but a task was created.");
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase8_InvalidSizeBytesRejected);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase8_InvalidSizeBytesRejected:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick))
    {
        Fail(L"Phase8_InvalidSizeBytesRejected timed out.");
        return true;
    }

    if (! state.fsDummy)
    {
        Fail(L"Phase8_InvalidSizeBytesRejected: dummy filesystem plugin was not loaded.");
        return true;
    }

    FileSystemOptions badOptions{};
    badOptions.sizeBytes                    = sizeof(FileSystemOptions) - 1u;
    badOptions.bandwidthLimitBytesPerSecond = 0;

    const HRESULT copyHr = state.fsDummy->CopyItem(L"/src.bin", L"/dst.bin", FILESYSTEM_FLAG_NONE, &badOptions, nullptr, nullptr);
    if (copyHr != E_INVALIDARG)
    {
        Fail(std::format(L"Expected CopyItem to reject invalid FileSystemOptions::sizeBytes with E_INVALIDARG, got 0x{:08X}.",
                         static_cast<unsigned long>(copyHr)));
        return true;
    }

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    const HRESULT qiDirHr = state.fsDummy->QueryInterface(IID_PPV_ARGS(dirOps.addressof()));
    if (FAILED(qiDirHr) || ! dirOps)
    {
        Fail(std::format(L"Phase8_InvalidSizeBytesRejected: dummy filesystem missing IFileSystemDirectoryOperations (hr=0x{:08X}).",
                         static_cast<unsigned long>(qiDirHr)));
        return true;
    }

    FileSystemDirectorySizeResult badSize{};
    badSize.sizeBytes = sizeof(FileSystemDirectorySizeResult) - 1u;

    const HRESULT sizeHr = dirOps->GetDirectorySize(L"/", FILESYSTEM_FLAG_NONE, nullptr, nullptr, &badSize);
    if (sizeHr != E_INVALIDARG)
    {
        Fail(std::format(L"Expected GetDirectorySize to reject invalid FileSystemDirectorySizeResult::sizeBytes with E_INVALIDARG, got 0x{:08X}.",
                         static_cast<unsigned long>(sizeHr)));
        return true;
    }

    wil::com_ptr<IFileSystemIO> io;
    const HRESULT qiIoHr = state.fsDummy->QueryInterface(IID_PPV_ARGS(io.addressof()));
    if (FAILED(qiIoHr) || ! io)
    {
        Fail(std::format(L"Phase8_InvalidSizeBytesRejected: dummy filesystem missing IFileSystemIO (hr=0x{:08X}).", static_cast<unsigned long>(qiIoHr)));
        return true;
    }

    FileSystemBasicInformation badBasic{};
    badBasic.sizeBytes = sizeof(FileSystemBasicInformation) - 1u;

    const HRESULT basicHr = io->GetFileBasicInformation(L"/src.bin", &badBasic);
    if (basicHr != E_INVALIDARG)
    {
        Fail(std::format(L"Expected GetFileBasicInformation to reject invalid FileSystemBasicInformation::sizeBytes with E_INVALIDARG, got 0x{:08X}.",
                         static_cast<unsigned long>(basicHr)));
        return true;
    }

    FileSystemRenamePair badRename{};
    badRename.sizeBytes  = sizeof(FileSystemRenamePair) - 1u;
    badRename.sourcePath = L"/src.bin";
    badRename.newName    = L"renamed.bin";

    const HRESULT renameHr = state.fsDummy->RenameItems(&badRename, 1, FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    if (renameHr != E_INVALIDARG)
    {
        Fail(std::format(L"Expected RenameItems to reject invalid FileSystemRenamePair::sizeBytes with E_INVALIDARG, got 0x{:08X}.",
                         static_cast<unsigned long>(renameHr)));
        return true;
    }

    NextStep(state, SelfTestState::Step::Phase8_PerItemOrchestration);
    return false;
}
case SelfTestState::Step::Phase8_PerItemOrchestration:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 240'000ull))
    {
        Fail(L"Phase8_PerItemOrchestration timed out.");
        return true;
    }

    const std::filesystem::path srcDir = state.tempRoot / L"peritem-src";
    const std::filesystem::path dstDir = state.tempRoot / L"peritem-dst";
    const std::filesystem::path fileA  = srcDir / L"big_a.bin";
    const std::filesystem::path fileB  = srcDir / L"big_b.bin";

    constexpr size_t kFileBytes = 8u * 1024u * 1024u;

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
        {
            Fail(L"Failed to reset peritem-src/peritem-dst directories.");
            return true;
        }

        if (! WriteTestFile(fileA, kFileBytes) || ! WriteTestFile(fileB, kFileBytes))
        {
            Fail(L"Failed to write per-item source files.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_COPY,
                                                                 FolderWindow::Pane::Left,
                                                                 FolderWindow::Pane::Right,
                                                                 state.fsLocal,
                                                                 {fileA, fileB},
                                                                 dstDir,
                                                                 flags,
                                                                 false,
                                                                 1ull * 1024ull * 1024ull,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start per-item copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(state.taskA.value());
        if (! task || ! task->HasStarted())
        {
            return false;
        }

        unsigned long totalItems = 0;
        uint64_t callbackCount   = 0;
        {
            std::scoped_lock lock(task->_progressMutex);
            totalItems    = task->_progressTotalItems;
            callbackCount = task->_progressCallbackCount;
        }

        if (callbackCount == 0)
        {
            return false;
        }

        if (totalItems != 2u)
        {
            Fail(std::format(L"Per-item progress totalItems expected 2, got {}.", totalItems));
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto it = state.completedTasks.find(state.taskA.value());
        if (it == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Per-item copy task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        const size_t dstCount = CountFiles(dstDir);
        if (dstCount != 2u)
        {
            Fail(std::format(L"Per-item copy output mismatch: expected 2 files, got {}.", dstCount));
            return true;
        }

        std::error_code ec;
        const auto sizeA = std::filesystem::file_size(dstDir / fileA.filename(), ec);
        if (ec || sizeA != kFileBytes)
        {
            Fail(L"Per-item destination file A has incorrect size.");
            return true;
        }
        ec.clear();
        const auto sizeB = std::filesystem::file_size(dstDir / fileB.filename(), ec);
        if (ec || sizeB != kFileBytes)
        {
            Fail(L"Per-item destination file B has incorrect size.");
            return true;
        }

        const uint64_t expectedTotalBytes = static_cast<uint64_t>(kFileBytes) * 2ull;
        if (it->second.preCalcTotalBytes != expectedTotalBytes || it->second.progressCompletedBytes != expectedTotalBytes)
        {
            Fail(std::format(L"Per-item byte aggregation mismatch: preCalc={} progress={} expected={}.",
                             it->second.preCalcTotalBytes,
                             it->second.progressCompletedBytes,
                             expectedTotalBytes));
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly:
{
    using Task              = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Phase9_ConflictPrompt_OverwriteReplaceReadonly timed out.");
        return true;
    }

    const std::filesystem::path srcDir  = state.tempRoot / L"conflict-src";
    const std::filesystem::path dstDir  = state.tempRoot / L"conflict-dst";
    const std::filesystem::path srcFile = srcDir / L"conflict.bin";
    const std::filesystem::path dstFile = dstDir / L"conflict.bin";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
        {
            Fail(L"Failed to reset conflict-src/conflict-dst directories.");
            return true;
        }

        if (! WriteTestFile(srcFile, 16 * 1024) || ! WriteTestFile(dstFile, 4 * 1024))
        {
            Fail(L"Failed to write conflict overwrite/read-only test files.");
            return true;
        }

        const DWORD attrs = GetFileAttributesW(dstFile.c_str());
        if (attrs == INVALID_FILE_ATTRIBUTES || ! SetFileAttributesW(dstFile.c_str(), attrs | FILE_ATTRIBUTE_READONLY))
        {
            Fail(L"Failed to set destination file to read-only.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_COPY,
                                                                 FolderWindow::Pane::Left,
                                                                 FolderWindow::Pane::Right,
                                                                 state.fsLocal,
                                                                 {srcFile},
                                                                 dstDir,
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start overwrite/readonly conflict copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }

        if (prompt->bucket != Task::ConflictBucket::Exists)
        {
            Fail(L"Expected Exists conflict bucket for overwrite prompt.");
            return true;
        }

        if (! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
        {
            Fail(L"Overwrite action not offered for Exists conflict.");
            return true;
        }

        task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }

        if (prompt->bucket == Task::ConflictBucket::Exists)
        {
            // Still draining the first prompt after submitting the overwrite decision.
            return false;
        }

        if (prompt->bucket != Task::ConflictBucket::ReadOnly)
        {
            Fail(L"Expected ReadOnly conflict bucket after overwrite on read-only destination.");
            return true;
        }

        if (! PromptHasAction(prompt.value(), Task::ConflictAction::ReplaceReadOnly))
        {
            Fail(L"ReplaceReadOnly action not offered for ReadOnly conflict.");
            return true;
        }

        task->SubmitConflictDecision(Task::ConflictAction::ReplaceReadOnly, false);
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        const auto it = state.completedTasks.find(state.taskA.value());
        if (it == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Conflict copy task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        std::error_code ec;
        const auto size = std::filesystem::file_size(dstFile, ec);
        if (ec || size != 16u * 1024u)
        {
            Fail(L"Destination file size mismatch after overwrite/readonly resolution.");
            return true;
        }

        const DWORD attrs = GetFileAttributesW(dstFile.c_str());
        if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_READONLY) != 0)
        {
            Fail(L"Destination file is still read-only after ReplaceReadOnly resolution.");
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache:
{
    using Task                                   = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick                      = GetTickCount64();
    constexpr ULONGLONG kPromptHoldMs            = 250ull;
    constexpr uint64_t kMinimumConvergenceWaitUs = 100'000ull;
    constexpr std::array<std::pair<std::wstring_view, size_t>, 4> kFiles{
        std::pair{std::wstring_view{L"a.bin"}, static_cast<size_t>(8ull * 1024ull)},
        std::pair{std::wstring_view{L"b.bin"}, static_cast<size_t>(16ull * 1024ull)},
        std::pair{std::wstring_view{L"c.bin"}, static_cast<size_t>(24ull * 1024ull)},
        std::pair{std::wstring_view{L"d.bin"}, static_cast<size_t>(32ull * 1024ull)},
    };
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        const HWND popup        = state.fileOps ? state.fileOps->GetPopupHwndForSelfTest() : nullptr;
        Task* task              = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const bool promptActive = TryGetConflictPromptCopy(task).has_value();
        Fail(std::format(L"Phase9_ConflictPrompt_ApplyToAllUiCache timed out. stepState={} popup={} taskExists={} promptActive={}",
                         state.stepState,
                         popup != nullptr,
                         task != nullptr,
                         promptActive));
        return true;
    }

    const std::filesystem::path srcDir = state.tempRoot / L"applyall-src";
    const std::filesystem::path dstDir = state.tempRoot / L"applyall-dst";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
        {
            Fail(L"Failed to reset applyall-src/applyall-dst directories.");
            return true;
        }

        std::vector<std::filesystem::path> sourceFiles;
        sourceFiles.reserve(kFiles.size());
        for (const auto& [fileName, sourceBytes] : kFiles)
        {
            const std::filesystem::path srcPath = srcDir / fileName;
            const std::filesystem::path dstPath = dstDir / fileName;
            if (! WriteTestFile(srcPath, sourceBytes) || ! WriteTestFile(dstPath, 1024))
            {
                Fail(L"Failed to write apply-to-all cache test files.");
                return true;
            }

            sourceFiles.push_back(srcPath);
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_COPY,
                                                                 FolderWindow::Pane::Left,
                                                                 FolderWindow::Pane::Right,
                                                                 state.fsLocal,
                                                                 sourceFiles,
                                                                 dstDir,
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start apply-to-all cache copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (! task)
        {
            return false;
        }

        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }

        if (prompt->bucket != Task::ConflictBucket::Exists)
        {
            Fail(L"Expected Exists conflict bucket for Apply-to-all cache prompt.");
            return true;
        }

        if (! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
        {
            Fail(L"Expected Overwrite action for Apply-to-all cache prompt.");
            return true;
        }

        if (state.markerTick == 0)
        {
            state.markerTick = nowTick;
            return false;
        }

        if ((nowTick - state.markerTick) < kPromptHoldMs)
        {
            return false;
        }

        const HWND popup = state.fileOps ? state.fileOps->GetPopupHwndForSelfTest() : nullptr;
        if (! popup)
        {
            return false;
        }

        FileOperationsPopupInternal::PopupSelfTestInvoke toggle{};
        toggle.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskConflictToggleApplyToAll;
        toggle.taskId = state.taskA.value();
        if (! InvokePopupSelfTest(popup, toggle))
        {
            Fail(L"Failed to invoke apply-to-all toggle via popup self-test message.");
            return true;
        }

        FileOperationsPopupInternal::PopupSelfTestInvoke click{};
        click.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskConflictAction;
        click.taskId = state.taskA.value();
        click.data   = static_cast<uint32_t>(Task::ConflictAction::Overwrite);
        if (! InvokePopupSelfTest(popup, click))
        {
            Fail(L"Failed to invoke overwrite via popup self-test message.");
            return true;
        }

        state.markerTick = 0;
        state.stepState  = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (! task)
        {
            state.stepState = 3;
            return false;
        }

        // Wait for the first prompt to clear before treating any later prompt as a "second prompt".
        if (TryGetConflictPromptCopy(task).has_value())
        {
            return false;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (TryGetConflictPromptCopy(task).has_value())
        {
            // Apply-to-all should have cached the resolution and avoided a second prompt for the same bucket.
            if (task)
            {
                task->SubmitConflictDecision(Task::ConflictAction::Cancel, false);
            }
            Fail(L"Unexpected second conflict prompt after Apply-to-all overwrite.");
            return true;
        }

        const auto it = state.completedTasks.find(state.taskA.value());
        if (it == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Apply-to-all cache task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        const CompletedTaskInfo& completion = it->second;
        if (completion.conflictPromptCount != 1u)
        {
            Fail(std::format(L"Apply-to-all: expected one prompt, saw {}.", completion.conflictPromptCount));
            return true;
        }

        if (completion.conflictConvergenceWaitUs < kMinimumConvergenceWaitUs)
        {
            Fail(std::format(
                L"Apply-to-all: conflict convergence wait too small: {}us threshold={}us.", completion.conflictConvergenceWaitUs, kMinimumConvergenceWaitUs));
            return true;
        }

        const std::wstring perfDetail = std::format(L"fileCount={} promptHoldMs={} promptCount={} conflictWaitUs={}",
                                                    kFiles.size(),
                                                    kPromptHoldMs,
                                                    completion.conflictPromptCount,
                                                    completion.conflictWaitUs);
        Debug::Perf::Emit(L"FileOps.SelfTest.ConflictApplyToAllConvergenceWait",
                          perfDetail,
                          completion.conflictConvergenceWaitUs,
                          completion.conflictPromptCount,
                          static_cast<uint64_t>(kFiles.size()),
                          S_OK);

        std::error_code ec;
        for (const auto& [fileName, sourceBytes] : kFiles)
        {
            const auto size = std::filesystem::file_size(dstDir / fileName, ec);
            if (ec || size != sourceBytes)
            {
                Fail(std::format(L"Apply-to-all: destination file {} has incorrect size after overwrite.", fileName));
                return true;
            }
            ec.clear();
        }

        NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_OverwriteAutoCap);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase9_ConflictPrompt_OverwriteAutoCap:
{
    using Task              = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Phase9_ConflictPrompt_OverwriteAutoCap timed out.");
        return true;
    }

    const std::filesystem::path srcDir  = state.tempRoot / L"overwritecap-src";
    const std::filesystem::path srcFile = srcDir / L"stuck.bin";
    const std::wstring dummyRoot        = L"/overwritecap";
    const std::wstring dummyConflictDir = L"/overwritecap/stuck.bin";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcDir))
        {
            Fail(L"Failed to reset overwritecap-src directory.");
            return true;
        }

        if (! WriteTestFile(srcFile, 4096))
        {
            Fail(L"Failed to write overwritecap source file.");
            return true;
        }

        if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyConflictDir))
        {
            Fail(L"Failed to prepare dummy destination conflict folder for overwrite-cap test.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_COPY,
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
            Fail(L"Failed to start overwrite-cap copy task.");
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

        if (prompt->bucket != Task::ConflictBucket::Exists)
        {
            Fail(L"Expected Exists conflict bucket for overwrite-cap prompt.");
            return true;
        }

        if (! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
        {
            Fail(L"Expected Overwrite action for overwrite-cap prompt.");
            return true;
        }

        const HWND popup = state.fileOps ? state.fileOps->GetPopupHwndForSelfTest() : nullptr;
        if (! popup)
        {
            return false;
        }

        FileOperationsPopupInternal::PopupSelfTestInvoke toggle{};
        toggle.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskConflictToggleApplyToAll;
        toggle.taskId = state.taskA.value();
        if (! InvokePopupSelfTest(popup, toggle))
        {
            Fail(L"Failed to toggle apply-to-all for overwrite-cap test.");
            return true;
        }

        FileOperationsPopupInternal::PopupSelfTestInvoke click{};
        click.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskConflictAction;
        click.taskId = state.taskA.value();
        click.data   = static_cast<uint32_t>(Task::ConflictAction::Overwrite);
        if (! InvokePopupSelfTest(popup, click))
        {
            Fail(L"Failed to invoke overwrite for overwrite-cap test.");
            return true;
        }

        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        if (! task)
        {
            Fail(L"Overwrite-cap task disappeared before second prompt.");
            return true;
        }

        const auto prompt = TryGetConflictPromptCopy(task);
        if (prompt.has_value() && prompt->applyToAllChecked)
        {
            return false;
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        Task* task        = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }

        if (prompt->bucket != Task::ConflictBucket::Exists)
        {
            Fail(L"Expected second Exists prompt after capped cached overwrite attempt.");
            return true;
        }

        if (! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
        {
            Fail(L"Expected Skip action on second overwrite-cap prompt.");
            return true;
        }

        task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
        state.stepState = 4;
        return false;
    }

    if (state.stepState == 4)
    {
        const auto it = state.completedTasks.find(state.taskA.value());
        if (it == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT expectedHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (it->second.hr != expectedHr)
        {
            Fail(std::format(L"Expected overwrite-cap copy task to return 0x{:08X}, got 0x{:08X}.",
                             static_cast<unsigned long>(expectedHr),
                             static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        wil::com_ptr<IFileSystemIO> dummyIo;
        if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
        {
            Fail(L"Dummy filesystem does not support IFileSystemIO for overwrite-cap validation.");
            return true;
        }

        unsigned long attrs = 0;
        if (FAILED(dummyIo->GetAttributes(dummyConflictDir.c_str(), &attrs)) || (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
        {
            Fail(L"Overwrite-cap: destination conflict directory was unexpectedly replaced.");
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_SkipAll);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase9_ConflictPrompt_SkipAll:
{
    using Task              = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Phase9_ConflictPrompt_SkipAll timed out.");
        return true;
    }

    const std::filesystem::path srcDir = state.tempRoot / L"skipall-src";
    const std::filesystem::path dstDir = state.tempRoot / L"skipall-dst";
    const std::filesystem::path srcA   = srcDir / L"a.bin";
    const std::filesystem::path srcB   = srcDir / L"b.bin";
    const std::filesystem::path dstA   = dstDir / L"a.bin";
    const std::filesystem::path dstB   = dstDir / L"b.bin";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
        {
            Fail(L"Failed to reset skipall-src/skipall-dst directories.");
            return true;
        }

        if (! WriteTestFile(srcA, 1024) || ! WriteTestFile(srcB, 2048) || ! WriteTestFile(dstA, 4096) || ! WriteTestFile(dstB, 4096))
        {
            Fail(L"Failed to write skip-all conflict test files.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_COPY,
                                                                 FolderWindow::Pane::Left,
                                                                 FolderWindow::Pane::Right,
                                                                 state.fsLocal,
                                                                 {srcA, srcB},
                                                                 dstDir,
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start skip-all copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }

        if (prompt->bucket != Task::ConflictBucket::Exists)
        {
            Fail(L"Expected Exists conflict bucket for SkipAll prompt.");
            return true;
        }

        if (! PromptHasAction(prompt.value(), Task::ConflictAction::SkipAll))
        {
            Fail(L"SkipAll action not offered for Exists conflict.");
            return true;
        }

        task->SubmitConflictDecision(Task::ConflictAction::SkipAll, false);
        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto it = state.completedTasks.find(state.taskA.value());
        if (it == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT expectedHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (it->second.hr != expectedHr)
        {
            Fail(std::format(L"Expected SkipAll copy task to return 0x{:08X}, got 0x{:08X}.",
                             static_cast<unsigned long>(expectedHr),
                             static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        std::error_code ec;
        const auto sizeA = std::filesystem::file_size(dstA, ec);
        if (ec || sizeA != 4096u)
        {
            Fail(L"SkipAll: destination file A size changed unexpectedly.");
            return true;
        }
        ec.clear();
        const auto sizeB = std::filesystem::file_size(dstB, ec);
        if (ec || sizeB != 4096u)
        {
            Fail(L"SkipAll: destination file B size changed unexpectedly.");
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_RetryCap);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase9_ConflictPrompt_RetryCap:
{
    using Task              = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Phase9_ConflictPrompt_RetryCap timed out.");
        return true;
    }

    const std::filesystem::path dir  = state.tempRoot / L"retrycap";
    const std::filesystem::path file = dir / L"locked.bin";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(dir))
        {
            Fail(L"Failed to reset retrycap directory.");
            return true;
        }

        if (! WriteTestFile(file, 16))
        {
            Fail(L"Failed to write retry-cap test file.");
            return true;
        }

        const DWORD initialAttributes = GetFileAttributesW(file.c_str());
        if (initialAttributes == INVALID_FILE_ATTRIBUTES || ! SetFileAttributesW(file.c_str(), initialAttributes | FILE_ATTRIBUTE_READONLY))
        {
            Fail(L"Failed to mark retry-cap test file read-only.");
            return true;
        }

        state.lockedFileHandle.reset(CreateFileW(file.c_str(), GENERIC_READ, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! state.lockedFileHandle)
        {
            Fail(L"Failed to open exclusive handle for retry-cap test file.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_DELETE,
                                                                 FolderWindow::Pane::Left,
                                                                 std::nullopt,
                                                                 state.fsLocal,
                                                                 {file},
                                                                 {},
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start retry-cap delete task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }

        if (prompt->bucket != Task::ConflictBucket::SharingViolation)
        {
            Fail(L"Expected SharingViolation conflict bucket for retry-cap prompt.");
            return true;
        }

        if (! PromptHasAction(prompt.value(), Task::ConflictAction::Retry) || prompt->retryFailed)
        {
            Fail(L"Expected Retry action to be offered for first SharingViolation prompt.");
            return true;
        }

        std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> diagnostics;
        state.fileOps->CollectDiagnostics(diagnostics);
        bool foundLivePromptDiagnostic = false;
        for (const auto& entry : diagnostics)
        {
            if (entry.taskId == state.taskA.value() && entry.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Warning &&
                entry.category == L"item.conflict.prompt")
            {
                foundLivePromptDiagnostic = true;
                break;
            }
        }
        if (! foundLivePromptDiagnostic)
        {
            Fail(L"Expected live conflict diagnostic entry for retry-cap prompt.");
            return true;
        }

        task->SubmitConflictDecision(Task::ConflictAction::Retry, false);
        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }

        if (prompt->bucket != Task::ConflictBucket::SharingViolation)
        {
            Fail(L"Expected SharingViolation conflict bucket for retry-cap prompt.");
            return true;
        }

        if (! prompt->retryFailed)
        {
            // Still draining the first prompt after submitting the retry decision.
            return false;
        }

        if (PromptHasAction(prompt.value(), Task::ConflictAction::Retry))
        {
            Fail(L"Expected second SharingViolation prompt to not offer Retry.");
            return true;
        }

        task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
        state.lockedFileHandle.reset();
        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        const auto it = state.completedTasks.find(state.taskA.value());
        if (it == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT expectedHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (it->second.hr != expectedHr)
        {
            Fail(std::format(L"Expected RetryCap delete task to return 0x{:08X}, got 0x{:08X}.",
                             static_cast<unsigned long>(expectedHr),
                             static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        std::error_code ec;
        if (! std::filesystem::exists(file, ec) || ec)
        {
            Fail(L"RetryCap: expected skipped file to still exist.");
            return true;
        }

        const DWORD finalAttributes = GetFileAttributesW(file.c_str());
        if (finalAttributes == INVALID_FILE_ATTRIBUTES || (finalAttributes & FILE_ATTRIBUTE_READONLY) == 0)
        {
            Fail(L"RetryCap: skipped file did not preserve its read-only attribute.");
            return true;
        }

        if (! SetFileAttributesW(file.c_str(), finalAttributes & ~FILE_ATTRIBUTE_READONLY))
        {
            Fail(L"RetryCap: failed to clear read-only attribute for cleanup.");
            return true;
        }

        ec.clear();
        static_cast<void>(std::filesystem::remove(file, ec));
        if (ec)
        {
            Fail(L"RetryCap: failed to remove skipped file after closing handle.");
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy:
{
    using Task              = FolderWindow::FileOperationState::Task;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(L"Phase9_ConflictPrompt_SkipContinuesDirectoryCopy timed out.");
        return true;
    }

    const std::filesystem::path srcDir     = state.tempRoot / L"skipdir-src";
    const std::filesystem::path dstDir     = state.tempRoot / L"skipdir-dst";
    const std::filesystem::path okFile     = srcDir / L"ok.bin";
    const std::filesystem::path lockedFile = srcDir / L"locked.bin";

    if (state.stepState == 0)
    {
        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
        {
            Fail(L"Failed to reset skipdir-src/skipdir-dst directories.");
            return true;
        }

        if (! WriteTestFile(okFile, 4096) || ! WriteTestFile(lockedFile, 4096))
        {
            Fail(L"Failed to write skip-continues directory test files.");
            return true;
        }

        state.lockedFileHandle.reset(CreateFileW(lockedFile.c_str(), GENERIC_READ, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! state.lockedFileHandle)
        {
            Fail(L"Failed to open exclusive handle for skip-continues directory test file.");
            return true;
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_COPY,
                                                                 FolderWindow::Pane::Left,
                                                                 FolderWindow::Pane::Right,
                                                                 state.fsLocal,
                                                                 {srcDir},
                                                                 dstDir,
                                                                 flags,
                                                                 false,
                                                                 0,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start skip-continues directory copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    if (state.stepState == 1)
    {
        Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
        const auto prompt = TryGetConflictPromptCopy(task);
        if (! prompt.has_value())
        {
            return false;
        }

        if (prompt->bucket != Task::ConflictBucket::SharingViolation)
        {
            Fail(L"Expected SharingViolation conflict bucket for skip-continues directory copy prompt.");
            return true;
        }

        if (! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
        {
            Fail(L"Skip action not offered for skip-continues directory copy prompt.");
            return true;
        }

        task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
        state.stepState = 2;
        return false;
    }

    if (state.stepState == 2)
    {
        const auto it = state.completedTasks.find(state.taskA.value());
        if (it == state.completedTasks.end())
        {
            return false;
        }

        const HRESULT expectedHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (it->second.hr != expectedHr)
        {
            Fail(std::format(L"Expected skip-continues directory copy to return 0x{:08X}, got 0x{:08X}.",
                             static_cast<unsigned long>(expectedHr),
                             static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        state.lockedFileHandle.reset();

        const std::filesystem::path dstCopiedDir = dstDir / srcDir.filename();

        std::error_code ec;
        const auto okSize = std::filesystem::file_size(dstCopiedDir / okFile.filename(), ec);
        if (ec || okSize != 4096u)
        {
            Fail(L"Skip-continues directory copy did not copy the expected ok.bin file.");
            return true;
        }

        ec.clear();
        const bool lockedExists = std::filesystem::exists(dstCopiedDir / lockedFile.filename(), ec);
        if (ec)
        {
            Fail(L"Skip-continues directory copy destination exists check failed.");
            return true;
        }
        if (lockedExists)
        {
            Fail(L"Skip-continues directory copy unexpectedly created locked.bin at destination.");
            return true;
        }

        NextStep(state, SelfTestState::Step::Phase9_PerItemConcurrency);
        return false;
    }

    return false;
}
case SelfTestState::Step::Phase9_PerItemConcurrency:
{
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        Fail(L"Phase9_PerItemConcurrency timed out.");
        return true;
    }

    const std::filesystem::path srcDir = state.tempRoot / L"peritem-conc-src";
    const std::filesystem::path dstDir = state.tempRoot / L"peritem-conc-dst";

    constexpr size_t kFileBytes    = 2ull * 1024ull * 1024ull;
    constexpr uint64_t kSpeedLimit = 1ull * 1024ull * 1024ull;
    constexpr int kFileCount       = 4;

    if (state.stepState == 0)
    {
        static_cast<void>(SetPluginConfiguration(
            state.infoLocal.get(),
            R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json"));

        if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
        {
            Fail(L"Failed to reset peritem-conc-src/peritem-conc-dst directories.");
            return true;
        }

        std::vector<std::filesystem::path> sources;
        sources.reserve(static_cast<size_t>(kFileCount));
        for (int i = 0; i < kFileCount; ++i)
        {
            const std::filesystem::path file = srcDir / std::format(L"c_{:02}.bin", i);
            if (! WriteTestFile(file, kFileBytes))
            {
                Fail(L"Failed to write per-item concurrency test file.");
                return true;
            }
            sources.push_back(file);
        }

        const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                                 FILESYSTEM_COPY,
                                                                 FolderWindow::Pane::Left,
                                                                 FolderWindow::Pane::Right,
                                                                 state.fsLocal,
                                                                 std::move(sources),
                                                                 dstDir,
                                                                 flags,
                                                                 false,
                                                                 kSpeedLimit,
                                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! state.taskA.has_value())
        {
            Fail(L"Failed to start per-item concurrency copy task.");
            return true;
        }

        state.stepState = 1;
        return false;
    }

    FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
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
            unsigned int maxConc = 0;
            size_t inFlight      = 0;
            {
                std::scoped_lock lock(task->_progressMutex);
                maxConc  = task->_perItemMaxConcurrency;
                inFlight = task->_perItemInFlightCallCount;
            }

            if (maxConc <= 1u)
            {
                Fail(L"Per-item concurrency expected >1, but task max concurrency is 1.");
                return true;
            }

            if (inFlight <= 1u)
            {
                if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                {
                    Fail(L"Expected >1 in-flight per-item calls but did not observe them.");
                    return true;
                }
                return false;
            }
        }

        state.stepState = 3;
        return false;
    }

    if (state.stepState == 3)
    {
        const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
        if (it == state.completedTasks.end())
        {
            return false;
        }

        if (FAILED(it->second.hr))
        {
            Fail(std::format(L"Per-item concurrency copy task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
            return true;
        }

        const size_t dstCount = CountFiles(dstDir);
        if (dstCount != static_cast<size_t>(kFileCount))
        {
            Fail(std::format(L"Per-item concurrency output mismatch: expected {} files, got {}.", kFileCount, dstCount));
            return true;
        }

        std::error_code ec;
        for (int i = 0; i < kFileCount; ++i)
        {
            const auto file = dstDir / std::format(L"c_{:02}.bin", i);
            const auto size = std::filesystem::file_size(file, ec);
            if (ec || size != kFileBytes)
            {
                Fail(L"Per-item concurrency: destination file has incorrect size.");
                return true;
            }
            ec.clear();
        }

        NextStep(state, SelfTestState::Step::Phase10_PermanentDelete);
        return false;
    }

    return false;
}
