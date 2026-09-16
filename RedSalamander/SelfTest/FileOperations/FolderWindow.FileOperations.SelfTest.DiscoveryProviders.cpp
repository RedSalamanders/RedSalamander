#if defined(FILEOPS_SELFTEST_INCLUDE_DISCOVERY_PROVIDERS)

// Provider discovery parity on direct routes. Each case below drives one provider's mutation entry
// points against its fake backend with no host in the loop, through a recording
// IFileSystemOperationControl, and judges the stream itself: it never decreases, exactly one report
// closes it, that closure is the last event, and it closed on the exact expected totals. The host
// hides every one of those defects, because it clamps a regressing cumulative value to a zero delta
// and counts late discovery without rejecting it. Each case then runs one same-endpoint leaf Copy
// through the host, the route that bypasses the bridge, and requires the task to have seen a closed
// exact record before the transfer moved a byte.
//
// Discovery describes what a call had to discover, not what it moved: a leaf is one file of its
// exact size; a Move that relocates a directory by renaming it is one directory and none of its
// descendants; a Copy or Delete that walks a directory reports the directory and everything under it.

case SelfTestState::Step::DiscoveryScope_DummyDirectApi:
{
    constexpr std::wstring_view label = L"DiscoveryScope_DummyDirectApi";
    static DiscoveryFixtureSelection selection{};
    static std::vector<std::wstring> defects;
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 120'000ull))
    {
        Fail(std::format(L"{} timed out.", label));
        return true;
    }
    if (! state.fsDummy || ! state.infoDummy)
    {
        Fail(std::format(L"{} requires the Dummy provider.", label));
        return true;
    }

    using Selection             = DiscoveryFixtureSelection;
    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
    IFileSystem* const dummy    = state.fsDummy.get();

    if (state.stepState == 0u)
    {
        // No lazily generated children: the tree holds exactly what the fixture wrote, so every
        // expected total below is exact rather than a lower bound.
        static_cast<void>(SetPluginConfiguration(
            state.infoDummy.get(), R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json"));
        if (! SeedDiscoveryFixtureSelection(dummy, L"/discovery-scope-direct-api", selection))
        {
            Fail(std::format(L"{}: could not seed the Dummy fixture.", label));
            return true;
        }
        defects.clear();
        const auto appendDefect = [&](std::wstring defect)
        {
            if (! defect.empty())
            {
                defects.push_back(std::move(defect));
            }
        };

        const std::wstring copiedLeaf = selection.Under(selection.copyDestination, L"leaf-a.bin");
        const std::wstring movedTree  = selection.Under(selection.moveDestination, L"tree");
        const std::array<const wchar_t*, 4> bulkRoots{selection.leafA.c_str(), selection.leafB.c_str(), selection.leafC.c_str(), selection.tree.c_str()};
        const std::array<const wchar_t*, 3> leaves{selection.leafA.c_str(), selection.leafB.c_str(), selection.leafC.c_str()};

        // Every call runs even when an earlier one is already wrong, so one run names every defect.
        appendDefect(RunDirectApiDiscoveryCall(L"CopyItem(leaf)", {Selection::kLeafBytesA, 1u, 0u}, [&](FileSystemOptions& options) noexcept {
            return dummy->CopyItem(selection.leafA.c_str(), copiedLeaf.c_str(), flags, &options, nullptr, nullptr);
        }));
        // The bulk walk visits one node per work item, descendants included.
        appendDefect(RunDirectApiDiscoveryCall(
            L"CopyItems(leaves+tree)", {Selection::kLeafBytes + Selection::kTreeBytes, 3u + Selection::kTreeChildCount, 1u}, [&](FileSystemOptions& options) noexcept {
                return dummy->CopyItems(
                    bulkRoots.data(), static_cast<unsigned long>(bulkRoots.size()), selection.bulkCopyDestination.c_str(), flags, &options, nullptr, nullptr);
            }));
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItems(leaves)", {Selection::kLeafBytes, 3u, 0u}, [&](FileSystemOptions& options) noexcept {
            return dummy->MoveItems(leaves.data(), static_cast<unsigned long>(leaves.size()), selection.moveDestination.c_str(), flags, &options, nullptr, nullptr);
        }));
        // A single-item Move relocates the directory node itself.
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItem(tree)", {0u, 0u, 1u}, [&](FileSystemOptions& options) noexcept {
            return dummy->MoveItem(selection.tree.c_str(), movedTree.c_str(), flags, &options, nullptr, nullptr);
        }));
        appendDefect(RunDirectApiDiscoveryCall(L"DeleteItem(tree)", {Selection::kTreeBytes, Selection::kTreeChildCount, 1u}, [&](FileSystemOptions& options) noexcept {
            return dummy->DeleteItem(movedTree.c_str(), flags, &options, nullptr, nullptr);
        }));
        const std::array<std::wstring, 3> movedLeaves{selection.Under(selection.moveDestination, L"leaf-a.bin"),
                                                      selection.Under(selection.moveDestination, L"leaf-b.bin"),
                                                      selection.Under(selection.moveDestination, L"leaf-c.bin")};
        const std::array<const wchar_t*, 3> movedLeafPaths{movedLeaves[0].c_str(), movedLeaves[1].c_str(), movedLeaves[2].c_str()};
        appendDefect(RunDirectApiDiscoveryCall(L"DeleteItems(leaves)", {Selection::kLeafBytes, 3u, 0u}, [&](FileSystemOptions& options) noexcept {
            return dummy->DeleteItems(movedLeafPaths.data(), static_cast<unsigned long>(movedLeafPaths.size()), flags, &options, nullptr, nullptr);
        }));

        // The admitted direct route through the host: a same-endpoint leaf Copy. No destination file
        // system is named, which is what makes the host call the provider directly instead of bridging.
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 state.fsDummy,
                                                 {std::filesystem::path(copiedLeaf)},
                                                 std::filesystem::path(selection.moveDestination),
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 nullptr);
        if (! state.taskA.has_value())
        {
            Fail(std::format(L"{}: could not admit the same-endpoint host leaf Copy.", label));
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
    if (const std::wstring defect = DescribeHostLeafDiscoveryDefect(label, completed->second, Selection::kLeafBytesA); ! defect.empty())
    {
        defects.push_back(defect);
    }
    if (! defects.empty())
    {
        Fail(std::format(L"{}: {}", label, JoinDiscoveryDefects(defects)));
        return true;
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.DiscoveryScope.DummyDirectApi",
                      L"calls=CopyItem,CopyItems,MoveItems,MoveItem,DeleteItem,DeleteItems,host-leaf-copy",
                      completed->second.discoveryDurationUs,
                      Selection::kLeafBytes + Selection::kTreeBytes,
                      completed->second.discoveryMutationsCompletedWhileOpen,
                      S_OK);
    NextStep(state, SelfTestState::Step::DiscoveryScope_CurlDirectApi);
    return false;
}

case SelfTestState::Step::DiscoveryScope_CurlDirectApi:
{
    constexpr std::wstring_view label = L"DiscoveryScope_CurlDirectApi";
    static std::optional<FakeBackendProvider> ftp;
    static DiscoveryFixtureSelection selection{};
    static std::vector<std::wstring> defects;
    const auto finish = [&](std::wstring failure) noexcept -> bool
    {
        ftp.reset();
        if (failure.empty())
        {
            return false;
        }
        Fail(failure);
        return true;
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        return finish(std::format(L"{} timed out at step {}.", label, state.stepState));
    }

    using Selection             = DiscoveryFixtureSelection;
    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    if (state.stepState == 0u)
    {
        ftp.emplace();
        std::wstring detail;
        const HRESULT startHr = StartFakeBackendProvider(
            kPluginIdFtp,
            "RedSalamanderCurlStartFakeFtpForSelfTest",
            "RedSalamanderCurlStopFakeFtpForSelfTest",
            [](unsigned int) {
                return std::string(R"({"connectTimeoutMs":2000,"operationTimeoutMs":8000,"copyMoveMaxConcurrency":2,"deleteMaxConcurrency":2,"ftpUseEpsv":true})");
            },
            ftp.value(),
            detail);
        if (FAILED(startHr))
        {
            return finish(std::format(L"{}: {} (hr=0x{:08X}).", label, detail, static_cast<unsigned long>(startHr)));
        }
        IFileSystem* const provider = ftp->fileSystem.get();
        if (! SeedDiscoveryFixtureSelection(provider, std::format(L"//anonymous@127.0.0.1:{}/discovery-scope", ftp->port), selection))
        {
            return finish(std::format(L"{}: could not seed the fake FTP fixture.", label));
        }
        defects.clear();
        const auto appendDefect = [&](std::wstring defect)
        {
            if (! defect.empty())
            {
                defects.push_back(std::move(defect));
            }
        };

        const std::wstring copiedLeaf = selection.Under(selection.copyDestination, L"leaf-a.bin");
        const std::wstring movedTree  = selection.Under(selection.moveDestination, L"tree");
        const std::array<const wchar_t*, 4> bulkRoots{selection.leafA.c_str(), selection.leafB.c_str(), selection.leafC.c_str(), selection.tree.c_str()};
        const std::array<const wchar_t*, 3> leaves{selection.leafA.c_str(), selection.leafB.c_str(), selection.leafC.c_str()};

        // A known leaf: its exact size is in hand before the first byte moves.
        appendDefect(RunDirectApiDiscoveryCall(L"CopyItem(leaf)", {Selection::kLeafBytesA, 1u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->CopyItem(selection.leafA.c_str(), copiedLeaf.c_str(), flags, &options, nullptr, nullptr);
        }));
        // Each selected root on the sequential pre-pass, then every descendant on the tree walk.
        appendDefect(RunDirectApiDiscoveryCall(
            L"CopyItems(leaves+tree)", {Selection::kLeafBytes + Selection::kTreeBytes, 3u + Selection::kTreeChildCount, 1u}, [&](FileSystemOptions& options) noexcept {
                return provider->CopyItems(
                    bulkRoots.data(), static_cast<unsigned long>(bulkRoots.size()), selection.bulkCopyDestination.c_str(), flags, &options, nullptr, nullptr);
            }));
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItems(leaves)", {Selection::kLeafBytes, 3u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->MoveItems(leaves.data(), static_cast<unsigned long>(leaves.size()), selection.moveDestination.c_str(), flags, &options, nullptr, nullptr);
        }));
        // A server-side rename relocates the directory and visits none of its descendants.
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItem(tree)", {0u, 0u, 1u}, [&](FileSystemOptions& options) noexcept {
            return provider->MoveItem(selection.tree.c_str(), movedTree.c_str(), flags, &options, nullptr, nullptr);
        }));
        // A recursive delete walks what it removes.
        appendDefect(RunDirectApiDiscoveryCall(L"DeleteItem(tree)", {Selection::kTreeBytes, Selection::kTreeChildCount, 1u}, [&](FileSystemOptions& options) noexcept {
            return provider->DeleteItem(movedTree.c_str(), flags, &options, nullptr, nullptr);
        }));
        const std::array<std::wstring, 3> movedLeaves{selection.Under(selection.moveDestination, L"leaf-a.bin"),
                                                      selection.Under(selection.moveDestination, L"leaf-b.bin"),
                                                      selection.Under(selection.moveDestination, L"leaf-c.bin")};
        const std::array<const wchar_t*, 3> movedLeafPaths{movedLeaves[0].c_str(), movedLeaves[1].c_str(), movedLeaves[2].c_str()};
        appendDefect(RunDirectApiDiscoveryCall(L"DeleteItems(leaves)", {Selection::kLeafBytes, 3u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->DeleteItems(movedLeafPaths.data(), static_cast<unsigned long>(movedLeafPaths.size()), flags, &options, nullptr, nullptr);
        }));

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 ftp->fileSystem,
                                                 {std::filesystem::path(copiedLeaf)},
                                                 std::filesystem::path(selection.moveDestination),
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 nullptr);
        if (! state.taskA.has_value())
        {
            return finish(std::format(L"{}: could not admit the same-endpoint host leaf Copy.", label));
        }
        state.stepState = 1u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    if (const std::wstring defect = DescribeHostLeafDiscoveryDefect(label, completed->second, Selection::kLeafBytesA); ! defect.empty())
    {
        defects.push_back(defect);
    }
    if (! defects.empty())
    {
        return finish(std::format(L"{}: {}", label, JoinDiscoveryDefects(defects)));
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.DiscoveryScope.CurlDirectApi",
                      L"calls=CopyItem,CopyItems,MoveItems,MoveItem,DeleteItem,DeleteItems,host-leaf-copy",
                      completed->second.discoveryDurationUs,
                      Selection::kLeafBytes + Selection::kTreeBytes,
                      completed->second.discoveryMutationsCompletedWhileOpen,
                      S_OK);
    static_cast<void>(finish({}));
    NextStep(state, SelfTestState::Step::DiscoveryScope_S3DirectApi);
    return false;
}

case SelfTestState::Step::DiscoveryScope_S3DirectApi:
{
    constexpr std::wstring_view label = L"DiscoveryScope_S3DirectApi";
    static std::optional<FakeBackendProvider> s3;
    static DiscoveryFixtureSelection selection{};
    static std::vector<std::wstring> defects;
    const auto finish = [&](std::wstring failure) noexcept -> bool
    {
        s3.reset();
        if (failure.empty())
        {
            return false;
        }
        Fail(failure);
        return true;
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        return finish(std::format(L"{} timed out at step {}.", label, state.stepState));
    }

    using Selection             = DiscoveryFixtureSelection;
    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    if (state.stepState == 0u)
    {
        s3.emplace();
        std::wstring detail;
        const HRESULT startHr = StartFakeBackendProvider(
            kPluginIdS3,
            "RedSalamanderS3StartFakeS3ForSelfTest",
            "RedSalamanderS3StopFakeS3ForSelfTest",
            [](unsigned int port) {
                return std::format(
                    R"({{"defaultRegion":"us-east-1","defaultEndpointOverride":"http://127.0.0.1:{}","useHttps":false,"verifyTls":false,"useVirtualAddressing":false,"anonymous":true,"connectTimeoutMs":2000,"requestTimeoutMs":8000}})",
                    port);
            },
            s3.value(),
            detail,
            "RedSalamanderS3FakeS3RequestLogForSelfTest");
        if (FAILED(startHr))
        {
            return finish(std::format(L"{}: {} (hr=0x{:08X}).", label, detail, static_cast<unsigned long>(startHr)));
        }
        IFileSystem* const provider = s3->fileSystem.get();
        // The tree is an implicit prefix: the provider's transfer of an explicit folder marker
        // object inside a prefix is a separate known defect, and discovery is what is under test.
        if (! SeedDiscoveryFixtureSelection(provider, L"/r0f-bucket/discovery-scope", selection, false))
        {
            return finish(std::format(L"{}: could not seed the fake S3 fixture.", label));
        }
        defects.clear();
        const auto appendDefect = [&](std::wstring defect)
        {
            if (! defect.empty())
            {
                defects.push_back(std::move(defect));
            }
        };

        const std::wstring copiedLeaf = selection.Under(selection.copyDestination, L"leaf-a.bin");
        const std::wstring movedTree  = selection.Under(selection.moveDestination, L"tree");
        const std::array<const wchar_t*, 4> bulkRoots{selection.leafA.c_str(), selection.leafB.c_str(), selection.leafC.c_str(), selection.tree.c_str()};
        const std::array<const wchar_t*, 3> leaves{selection.leafA.c_str(), selection.leafB.c_str(), selection.leafC.c_str()};

        // S3 builds a complete transfer plan before it touches an object, so every route below
        // reports one exact closed total up front. A prefix is one directory; its marker object,
        // when the bucket keeps one, is that directory rather than a file.
        appendDefect(RunDirectApiDiscoveryCall(L"CopyItem(leaf)", {Selection::kLeafBytesA, 1u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->CopyItem(selection.leafA.c_str(), copiedLeaf.c_str(), flags, &options, nullptr, nullptr);
        }));
        appendDefect(RunDirectApiDiscoveryCall(
            L"CopyItems(leaves+tree)", {Selection::kLeafBytes + Selection::kTreeBytes, 3u + Selection::kTreeChildCount, 1u}, [&](FileSystemOptions& options) noexcept {
                return provider->CopyItems(
                    bulkRoots.data(), static_cast<unsigned long>(bulkRoots.size()), selection.bulkCopyDestination.c_str(), flags, &options, nullptr, nullptr);
            }));
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItems(leaves)", {Selection::kLeafBytes, 3u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->MoveItems(leaves.data(), static_cast<unsigned long>(leaves.size()), selection.moveDestination.c_str(), flags, &options, nullptr, nullptr);
        }));
        // S3 has no rename: a prefix Move copies and deletes every object, so it discovers them all.
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItem(tree)", {Selection::kTreeBytes, Selection::kTreeChildCount, 1u}, [&](FileSystemOptions& options) noexcept {
            return provider->MoveItem(selection.tree.c_str(), movedTree.c_str(), flags, &options, nullptr, nullptr);
        }));
        appendDefect(RunDirectApiDiscoveryCall(L"DeleteItem(tree)", {Selection::kTreeBytes, Selection::kTreeChildCount, 1u}, [&](FileSystemOptions& options) noexcept {
            return provider->DeleteItem(movedTree.c_str(), flags, &options, nullptr, nullptr);
        }));
        const std::array<std::wstring, 3> movedLeaves{selection.Under(selection.moveDestination, L"leaf-a.bin"),
                                                      selection.Under(selection.moveDestination, L"leaf-b.bin"),
                                                      selection.Under(selection.moveDestination, L"leaf-c.bin")};
        const std::array<const wchar_t*, 3> movedLeafPaths{movedLeaves[0].c_str(), movedLeaves[1].c_str(), movedLeaves[2].c_str()};
        appendDefect(RunDirectApiDiscoveryCall(L"DeleteItems(leaves)", {Selection::kLeafBytes, 3u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->DeleteItems(movedLeafPaths.data(), static_cast<unsigned long>(movedLeafPaths.size()), flags, &options, nullptr, nullptr);
        }));

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 s3->fileSystem,
                                                 {std::filesystem::path(copiedLeaf)},
                                                 std::filesystem::path(selection.moveDestination),
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 nullptr);
        if (! state.taskA.has_value())
        {
            return finish(std::format(L"{}: could not admit the same-endpoint host leaf Copy.", label));
        }
        state.stepState = 1u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    if (const std::wstring defect = DescribeHostLeafDiscoveryDefect(label, completed->second, Selection::kLeafBytesA); ! defect.empty())
    {
        defects.push_back(defect);
    }
    if (! defects.empty())
    {
        return finish(std::format(L"{}: {}", label, JoinDiscoveryDefects(defects, s3->RequestLog())));
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.DiscoveryScope.S3DirectApi",
                      L"calls=CopyItem,CopyItems,MoveItems,MoveItem,DeleteItem,DeleteItems,host-leaf-copy",
                      completed->second.discoveryDurationUs,
                      Selection::kLeafBytes + Selection::kTreeBytes,
                      completed->second.discoveryMutationsCompletedWhileOpen,
                      S_OK);
    static_cast<void>(finish({}));
    NextStep(state, SelfTestState::Step::DiscoveryScope_GDriveDirectApi);
    return false;
}

case SelfTestState::Step::DiscoveryScope_GDriveDirectApi:
{
    constexpr std::wstring_view label = L"DiscoveryScope_GDriveDirectApi";
    static std::optional<FakeBackendProvider> drive;
    static DiscoveryFixtureSelection selection{};
    static std::vector<std::wstring> defects;
    const auto finish = [&](std::wstring failure) noexcept -> bool
    {
        drive.reset();
        if (failure.empty())
        {
            return false;
        }
        Fail(failure);
        return true;
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        return finish(std::format(L"{} timed out at step {}.", label, state.stepState));
    }

    using Selection             = DiscoveryFixtureSelection;
    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    if (state.stepState == 0u)
    {
        drive.emplace();
        std::wstring detail;
        const HRESULT startHr = StartFakeBackendProvider(
            kPluginIdGoogleDrive,
            "RedSalamanderGoogleDriveStartFakeDriveForSelfTest",
            "RedSalamanderGoogleDriveStopFakeDriveForSelfTest",
            [](unsigned int) { return std::string(R"({"connectTimeoutMs":2000,"requestTimeoutMs":8000})"); },
            drive.value(),
            detail,
            "RedSalamanderGoogleDriveFakeDriveRequestLogForSelfTest");
        if (FAILED(startHr))
        {
            return finish(std::format(L"{}: {} (hr=0x{:08X}).", label, detail, static_cast<unsigned long>(startHr)));
        }
        IFileSystem* const provider = drive->fileSystem.get();
        if (! SeedDiscoveryFixtureSelection(provider, L"/@conn:google-drive-selftest/discovery-scope", selection))
        {
            return finish(std::format(L"{}: could not seed the fake Drive fixture.", label));
        }
        defects.clear();
        const auto appendDefect = [&](std::wstring defect)
        {
            if (! defect.empty())
            {
                defects.push_back(std::move(defect));
            }
        };

        const std::wstring copiedLeaf = selection.Under(selection.copyDestination, L"leaf-a.bin");
        const std::wstring movedTree  = selection.Under(selection.moveDestination, L"tree");
        const std::array<const wchar_t*, 4> bulkRoots{selection.leafA.c_str(), selection.leafB.c_str(), selection.leafC.c_str(), selection.tree.c_str()};
        const std::array<const wchar_t*, 3> leaves{selection.leafA.c_str(), selection.leafB.c_str(), selection.leafC.c_str()};

        // A server-side file copy knows the leaf's size before it starts; the tree copy reports the
        // folder and then each child as it is popped from its frame.
        appendDefect(RunDirectApiDiscoveryCall(L"CopyItem(leaf)", {Selection::kLeafBytesA, 1u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->CopyItem(selection.leafA.c_str(), copiedLeaf.c_str(), flags, &options, nullptr, nullptr);
        }));
        appendDefect(RunDirectApiDiscoveryCall(
            L"CopyItems(leaves+tree)", {Selection::kLeafBytes + Selection::kTreeBytes, 3u + Selection::kTreeChildCount, 1u}, [&](FileSystemOptions& options) noexcept {
                return provider->CopyItems(
                    bulkRoots.data(), static_cast<unsigned long>(bulkRoots.size()), selection.bulkCopyDestination.c_str(), flags, &options, nullptr, nullptr);
            }));
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItems(leaves)", {Selection::kLeafBytes, 3u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->MoveItems(leaves.data(), static_cast<unsigned long>(leaves.size()), selection.moveDestination.c_str(), flags, &options, nullptr, nullptr);
        }));
        // The native folder move re-homes the whole subtree with one parent change.
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItem(tree)", {0u, 0u, 1u}, [&](FileSystemOptions& options) noexcept {
            return provider->MoveItem(selection.tree.c_str(), movedTree.c_str(), flags, &options, nullptr, nullptr);
        }));
        // Drive removes a subtree as one object, so a folder delete discovers the folder alone.
        appendDefect(RunDirectApiDiscoveryCall(L"DeleteItem(tree)", {0u, 0u, 1u}, [&](FileSystemOptions& options) noexcept {
            return provider->DeleteItem(movedTree.c_str(), flags, &options, nullptr, nullptr);
        }));
        const std::array<std::wstring, 3> movedLeaves{selection.Under(selection.moveDestination, L"leaf-a.bin"),
                                                      selection.Under(selection.moveDestination, L"leaf-b.bin"),
                                                      selection.Under(selection.moveDestination, L"leaf-c.bin")};
        const std::array<const wchar_t*, 3> movedLeafPaths{movedLeaves[0].c_str(), movedLeaves[1].c_str(), movedLeaves[2].c_str()};
        appendDefect(RunDirectApiDiscoveryCall(L"DeleteItems(leaves)", {Selection::kLeafBytes, 3u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->DeleteItems(movedLeafPaths.data(), static_cast<unsigned long>(movedLeafPaths.size()), flags, &options, nullptr, nullptr);
        }));

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 drive->fileSystem,
                                                 {std::filesystem::path(copiedLeaf)},
                                                 std::filesystem::path(selection.moveDestination),
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 nullptr);
        if (! state.taskA.has_value())
        {
            return finish(std::format(L"{}: could not admit the same-endpoint host leaf Copy.", label));
        }
        state.stepState = 1u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    if (const std::wstring defect = DescribeHostLeafDiscoveryDefect(label, completed->second, Selection::kLeafBytesA); ! defect.empty())
    {
        defects.push_back(defect);
    }
    if (! defects.empty())
    {
        return finish(std::format(L"{}: {}", label, JoinDiscoveryDefects(defects, drive->RequestLog())));
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.DiscoveryScope.GDriveDirectApi",
                      L"calls=CopyItem,CopyItems,MoveItems,MoveItem,DeleteItem,DeleteItems,host-leaf-copy",
                      completed->second.discoveryDurationUs,
                      Selection::kLeafBytes + Selection::kTreeBytes,
                      completed->second.discoveryMutationsCompletedWhileOpen,
                      S_OK);
    static_cast<void>(finish({}));
    NextStep(state, SelfTestState::Step::DiscoveryScope_GraphDirectApi);
    return false;
}

case SelfTestState::Step::DiscoveryScope_GraphDirectApi:
{
    // Microsoft Drive owns exactly one direct route, Native Move; it advertises no same-provider
    // Copy and CopyItem is ERROR_NOT_SUPPORTED. A Move is one metadata change, but the metadata GET
    // it already pays for describes the item, so a leaf reports its exact size and a folder is one
    // directory. The host witness is therefore the Move itself rather than a leaf Copy.
    constexpr std::wstring_view label = L"DiscoveryScope_GraphDirectApi";
    static std::optional<FakeBackendProvider> graph;
    static DiscoveryFixtureSelection selection{};
    static std::vector<std::wstring> defects;
    const auto finish = [&](std::wstring failure) noexcept -> bool
    {
        graph.reset();
        if (failure.empty())
        {
            return false;
        }
        Fail(failure);
        return true;
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        return finish(std::format(L"{} timed out at step {}.", label, state.stepState));
    }

    using Selection = DiscoveryFixtureSelection;

    if (state.stepState == 0u)
    {
        graph.emplace();
        std::wstring detail;
        const HRESULT startHr = StartFakeBackendProvider(
            kPluginIdOneDrivePersonal,
            "RedSalamanderMicrosoftDriveStartFakeGraphForSelfTest",
            "RedSalamanderMicrosoftDriveStopFakeGraphForSelfTest",
            [](unsigned int) { return std::string(R"({"connectTimeoutMs":2000,"requestTimeoutMs":8000})"); },
            graph.value(),
            detail,
            "RedSalamanderMicrosoftDriveFakeGraphRequestLogForSelfTest");
        if (FAILED(startHr))
        {
            return finish(std::format(L"{}: {} (hr=0x{:08X}).", label, detail, static_cast<unsigned long>(startHr)));
        }
        IFileSystem* const provider = graph->fileSystem.get();
        if (! SeedDiscoveryFixtureSelection(provider, L"/@conn:microsoft-drive-selftest/discovery-scope", selection))
        {
            return finish(std::format(L"{}: could not seed the fake Graph fixture.", label));
        }
        defects.clear();
        const auto appendDefect = [&](std::wstring defect)
        {
            if (! defect.empty())
            {
                defects.push_back(std::move(defect));
            }
        };

        const std::wstring movedLeafA = selection.Under(selection.moveDestination, L"leaf-a.bin");
        const std::wstring movedTree  = selection.Under(selection.moveDestination, L"tree");
        const std::array<const wchar_t*, 2> leaves{selection.leafB.c_str(), selection.leafC.c_str()};
        const FileSystemFlags moveFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

        appendDefect(RunDirectApiDiscoveryCall(L"MoveItem(leaf)", {Selection::kLeafBytesA, 1u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->MoveItem(selection.leafA.c_str(), movedLeafA.c_str(), moveFlags, &options, nullptr, nullptr);
        }));
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItems(leaves)", {Selection::kLeafBytesB + Selection::kLeafBytesC, 2u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->MoveItems(leaves.data(), static_cast<unsigned long>(leaves.size()), selection.moveDestination.c_str(), moveFlags, &options, nullptr, nullptr);
        }));
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItem(tree)", {0u, 0u, 1u}, [&](FileSystemOptions& options) noexcept {
            return provider->MoveItem(selection.tree.c_str(), movedTree.c_str(), moveFlags, &options, nullptr, nullptr);
        }));
        // Delete is recycle-only on this provider and removes the item as one object.
        const FileSystemFlags recycleFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN);
        appendDefect(RunDirectApiDiscoveryCall(L"DeleteItem(tree)", {0u, 0u, 1u}, [&](FileSystemOptions& options) noexcept {
            return provider->DeleteItem(movedTree.c_str(), recycleFlags, &options, nullptr, nullptr);
        }));
        const std::array<std::wstring, 2> movedLeaves{selection.Under(selection.moveDestination, L"leaf-b.bin"),
                                                      selection.Under(selection.moveDestination, L"leaf-c.bin")};
        const std::array<const wchar_t*, 2> movedLeafPaths{movedLeaves[0].c_str(), movedLeaves[1].c_str()};
        appendDefect(RunDirectApiDiscoveryCall(L"DeleteItems(leaves)", {Selection::kLeafBytesB + Selection::kLeafBytesC, 2u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->DeleteItems(movedLeafPaths.data(), static_cast<unsigned long>(movedLeafPaths.size()), recycleFlags, &options, nullptr, nullptr);
        }));

        // The admitted direct route through the host: a same-endpoint Native Move of one leaf.
        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_MOVE,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 graph->fileSystem,
                                                 {std::filesystem::path(movedLeafA)},
                                                 std::filesystem::path(selection.copyDestination),
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 nullptr);
        if (! state.taskA.has_value())
        {
            return finish(std::format(L"{}: could not admit the same-endpoint host leaf Move.", label));
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
    if (FAILED(info.hr) || info.bridgeFileAdmissionCount != 0ull || ! info.discoveryClosed || info.discoveredFiles != 1ul ||
        info.discoveredTotalBytes != Selection::kLeafBytesA || info.discoveryGrowthAfterCloseCount != 0ull)
    {
        defects.push_back(std::format(L"the host leaf Move did not see one closed exact record on the direct route (hr=0x{:08X} bridgeAdmissions={} "
                                      L"closed={} files={} bytes={}/{} growthAfterClose={}).",
                                      static_cast<unsigned long>(info.hr),
                                      info.bridgeFileAdmissionCount,
                                      info.discoveryClosed ? 1 : 0,
                                      info.discoveredFiles,
                                      info.discoveredTotalBytes,
                                      Selection::kLeafBytesA,
                                      info.discoveryGrowthAfterCloseCount));
    }
    if (! defects.empty())
    {
        return finish(std::format(L"{}: {}", label, JoinDiscoveryDefects(defects, graph->RequestLog())));
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.DiscoveryScope.GraphDirectApi",
                      L"calls=MoveItem,MoveItems,DeleteItem,DeleteItems,host-leaf-move",
                      info.discoveryDurationUs,
                      Selection::kLeafBytes,
                      info.discoveryMutationsCompletedWhileOpen,
                      S_OK);
    static_cast<void>(finish({}));
    NextStep(state, SelfTestState::Step::DiscoveryScope_MtpDirectApi);
    return false;
}

case SelfTestState::Step::DiscoveryScope_MtpDirectApi:
{
    // MTP advertises Copy and Move but no native Move, so its direct routes are the same-device
    // CopyItem and the CopyOnly Move the host routes through it. The backend answers a source's
    // kind and size from the path cache the listing already filled, so a leaf reports its exact
    // size, and a directory is handed to the device as one object and discovers the directory.
    constexpr std::wstring_view label = L"DiscoveryScope_MtpDirectApi";
    static wil::unique_hmodule mtpModule;
    static wil::com_ptr<IFileSystem> mtp;
    static DiscoveryFixtureSelection selection{};
    static std::vector<std::wstring> defects;
    const auto finish = [&](std::wstring failure) noexcept -> bool
    {
        mtp.reset();
        mtpModule.reset();
        if (failure.empty())
        {
            return false;
        }
        Fail(failure);
        return true;
    };
    const ULONGLONG nowTick = GetTickCount64();
    if (HasTimedOut(state, nowTick, 180'000ull))
    {
        return finish(std::format(L"{} timed out at step {}.", label, state.stepState));
    }

    using Selection             = DiscoveryFixtureSelection;
    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    if (state.stepState == 0u)
    {
        using CreateMtpForSelfTestFn = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const char*, void**) noexcept;
        FARPROC createAddress        = nullptr;
        HRESULT hr                   = SelfTest::LoadMtpPluginSelfTestExport("RedSalamanderMtpCreateForSelfTest", mtpModule, createAddress);
        if (FAILED(hr) || createAddress == nullptr)
        {
            return finish(std::format(L"{}: the fake-MTP factory is unavailable (hr=0x{:08X}).", label, static_cast<unsigned long>(hr)));
        }
        FactoryOptions factoryOptions{};
        factoryOptions.debugLevel = DEBUG_LEVEL_NONE;
#pragma warning(push)
#pragma warning(disable : 4191) // The named self-test export fixes the typed factory ABI.
        hr = reinterpret_cast<CreateMtpForSelfTestFn>(createAddress)(__uuidof(IFileSystem), &factoryOptions, GetHostServices(), R"json({"operationDelayMs":0})json", mtp.put_void());
#pragma warning(pop)
        if (FAILED(hr) || ! mtp)
        {
            return finish(std::format(L"{}: the fake-MTP instance was not created (hr=0x{:08X}).", label, static_cast<unsigned long>(hr)));
        }
        wil::com_ptr<IInformations> information;
        wil::com_ptr<IFileSystemInitialize> initialize;
        if (FAILED(mtp->QueryInterface(IID_PPV_ARGS(information.addressof()))) || ! information ||
            FAILED(information->SetConfiguration(R"json({"readOnly":false,"commandTimeoutMs":5000})json")) ||
            FAILED(mtp->QueryInterface(IID_PPV_ARGS(initialize.addressof()))) || ! initialize || FAILED(initialize->Initialize(L"/", nullptr)))
        {
            return finish(std::format(L"{}: the fake-MTP instance did not accept the fixture configuration.", label));
        }
        IFileSystem* const provider = mtp.get();
        if (! SeedDiscoveryFixtureSelection(provider, L"/Fake Phone/Internal Storage/DCIM/Camera/discovery-scope", selection))
        {
            return finish(std::format(L"{}: could not seed the fake-MTP fixture.", label));
        }
        defects.clear();
        const auto appendDefect = [&](std::wstring defect)
        {
            if (! defect.empty())
            {
                defects.push_back(std::move(defect));
            }
        };

        const std::wstring copiedLeaf = selection.Under(selection.copyDestination, L"leaf-a.bin");
        const std::wstring movedTree  = selection.Under(selection.moveDestination, L"tree");
        const std::wstring movedLeafB = selection.Under(selection.moveDestination, L"leaf-b.bin");
        const std::array<const wchar_t*, 2> leaves{selection.leafB.c_str(), selection.leafC.c_str()};

        appendDefect(RunDirectApiDiscoveryCall(L"CopyItem(leaf)", {Selection::kLeafBytesA, 1u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->CopyItem(selection.leafA.c_str(), copiedLeaf.c_str(), flags, &options, nullptr, nullptr);
        }));
        // The device relocates a directory as one object; nothing beneath it is visited here.
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItem(tree)", {0u, 0u, 1u}, [&](FileSystemOptions& options) noexcept {
            return provider->MoveItem(selection.tree.c_str(), movedTree.c_str(), flags, &options, nullptr, nullptr);
        }));
        appendDefect(RunDirectApiDiscoveryCall(L"MoveItems(leaves)", {Selection::kLeafBytesB + Selection::kLeafBytesC, 2u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->MoveItems(leaves.data(), static_cast<unsigned long>(leaves.size()), selection.moveDestination.c_str(), flags, &options, nullptr, nullptr);
        }));
        appendDefect(RunDirectApiDiscoveryCall(L"DeleteItem(leaf)", {Selection::kLeafBytesB, 1u, 0u}, [&](FileSystemOptions& options) noexcept {
            return provider->DeleteItem(movedLeafB.c_str(), flags, &options, nullptr, nullptr);
        }));

        state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                 FILESYSTEM_COPY,
                                                 FolderWindow::Pane::Left,
                                                 FolderWindow::Pane::Right,
                                                 mtp,
                                                 {std::filesystem::path(copiedLeaf)},
                                                 std::filesystem::path(selection.moveDestination),
                                                 FILESYSTEM_FLAG_NONE,
                                                 false,
                                                 0u,
                                                 FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                 false,
                                                 nullptr);
        if (! state.taskA.has_value())
        {
            return finish(std::format(L"{}: could not admit the same-endpoint host leaf Copy.", label));
        }
        state.stepState = 1u;
        return false;
    }

    const auto completed = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
    if (completed == state.completedTasks.end())
    {
        return false;
    }
    if (const std::wstring defect = DescribeHostLeafDiscoveryDefect(label, completed->second, Selection::kLeafBytesA); ! defect.empty())
    {
        defects.push_back(defect);
    }
    if (! defects.empty())
    {
        return finish(std::format(L"{}: {}", label, JoinDiscoveryDefects(defects)));
    }
    Debug::Perf::Emit(L"FileOps.SelfTest.DiscoveryScope.MtpDirectApi",
                      L"calls=CopyItem,MoveItem(tree),MoveItems,DeleteItem,host-leaf-copy",
                      completed->second.discoveryDurationUs,
                      Selection::kLeafBytes,
                      completed->second.discoveryMutationsCompletedWhileOpen,
                      S_OK);
    static_cast<void>(finish({}));
    NextStep(state, SelfTestState::Step::R0fSmb_BlockedSynchronousCallCancelReturns);
    return false;
}

#endif
