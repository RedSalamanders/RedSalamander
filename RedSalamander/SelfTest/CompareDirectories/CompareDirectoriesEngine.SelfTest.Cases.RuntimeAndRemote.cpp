SelfTest::RunCase(options,
                  suite,
                  L"showIdentical",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: keepIdenticalItems retains identical files; showIdenticalItems toggles view without invalidating decisions.
    if (const auto foldersOpt = CreateCaseFolders(root, L"identical"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"same.txt", "SAME"), L"Failed to create same.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"same.txt", "SAME"), L"Failed to create same.txt (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.keepIdenticalItems = true;
        auto session                = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        const auto fsLeft           = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
        const auto fsRight          = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

        const uint64_t versionBefore = session->GetVersion();
        const auto decisionBefore    = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decisionBefore), L"Decision missing (before showIdentical).");
        if (decisionBefore)
        {
            const auto* item = FindItem(*decisionBefore, L"same.txt");
            state.Require(item != nullptr, L"same.txt missing from cached decision when keepIdenticalItems is enabled.");
            if (item)
            {
                state.Require(! item->isDifferent, L"same.txt expected identical (before showIdentical).");
                state.Require(item->differenceMask == 0u, L"same.txt expected differenceMask=0 (before showIdentical).");
            }
        }

        state.Require(! ContainsName(EnumerateDirectoryNames(fsLeft, folders.left, state), L"same.txt"),
                      L"same.txt expected excluded from left enumeration (before showIdentical).");
        state.Require(! ContainsName(EnumerateDirectoryNames(fsRight, folders.right, state), L"same.txt"),
                      L"same.txt expected excluded from right enumeration (before showIdentical).");

        settings.showIdenticalItems = true;
        session->SetSettings(settings);

        const uint64_t versionAfter = session->GetVersion();
        state.Require(versionAfter == versionBefore, L"SetSettings(showIdenticalItems) must not invalidate decisions.");

        const auto decisionAfter = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(decisionAfter == decisionBefore, L"Decision should remain cached across showIdenticalItems toggle.");
        state.Require(static_cast<bool>(decisionAfter), L"Decision missing (after showIdentical).");
        if (decisionAfter)
        {
            const auto* item = FindItem(*decisionAfter, L"same.txt");
            state.Require(item != nullptr, L"same.txt missing from decision (after showIdentical).");
            if (item)
            {
                state.Require(! item->isDifferent, L"same.txt expected identical (after showIdentical).");
                state.Require(item->differenceMask == 0u, L"same.txt expected differenceMask=0 (after showIdentical).");
            }
        }

        state.Require(ContainsName(EnumerateDirectoryNames(fsLeft, folders.left, state), L"same.txt"),
                      L"same.txt expected included in left enumeration (after showIdentical).");
        state.Require(ContainsName(EnumerateDirectoryNames(fsRight, folders.right, state), L"same.txt"),
                      L"same.txt expected included in right enumeration (after showIdentical).");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: identical.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"content_pending_elided",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: When keepIdenticalItems is off, file-level ContentPending placeholders are elided (tracked per-folder)
    // so content compare does not explode memory on very large folders.
    if (const auto foldersOpt = CreateCaseFolders(root, L"content_pending_elided"))
    {
        const auto& folders = foldersOpt.value();

        constexpr int kFileCount = 200;

        std::vector<std::byte> payload(16 * 1024);
        for (size_t i = 0; i < payload.size(); ++i)
        {
            payload[i] = static_cast<std::byte>(i & 0xFF);
        }

        const std::span<const std::byte> payloadSpan(payload.data(), payload.size());
        for (int i = 0; i < kFileCount; ++i)
        {
            const std::wstring name = std::format(L"f_{:04}.bin", i);
            state.Require(SelfTest::WriteBinaryFile(folders.left / name, payloadSpan), L"Failed to write test file (left).");
            state.Require(SelfTest::WriteBinaryFile(folders.right / name, payloadSpan), L"Failed to write test file (right).");
        }

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSize        = false;
        settings.compareDateTime    = false;
        settings.compareAttributes  = false;
        settings.compareContent     = true;
        settings.keepIdenticalItems = false;

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

        std::mutex mutex;
        std::condition_variable cv;
        bool contentDone      = false;
        uint64_t lastPending  = 0;
        uint64_t lastTotal    = 0;
        uint64_t lastComplete = 0;

        session->SetContentProgressCallback([&](uint32_t /*workerIndex*/,
                                                const std::filesystem::path& /*relativeFolder*/,
                                                std::wstring_view /*entryName*/,
                                                uint64_t /*fileTotalBytes*/,
                                                uint64_t /*fileCompletedBytes*/,
                                                uint64_t /*overallTotalBytes*/,
                                                uint64_t /*overallCompletedBytes*/,
                                                uint64_t pendingContentCompares,
                                                uint64_t totalContentCompares,
                                                uint64_t completedContentCompares) noexcept
        {
            std::lock_guard guard(mutex);
            lastPending  = pendingContentCompares;
            lastTotal    = totalContentCompares;
            lastComplete = completedContentCompares;
            if (pendingContentCompares == 0u && totalContentCompares != 0u && completedContentCompares == totalContentCompares)
            {
                contentDone = true;
            }
            cv.notify_all();
        });

        const auto decisionInitial = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decisionInitial), L"content_pending_elided: decision missing.");
        if (decisionInitial)
        {
            state.Require(decisionInitial->items.empty(), L"content_pending_elided: expected no per-file ContentPending items in differences-only mode.");
            state.Require(decisionInitial->pendingContentCompareCount == static_cast<uint32_t>(kFileCount),
                          L"content_pending_elided: expected pendingContentCompareCount to match the file count.");
            state.Require(decisionInitial->anyPending, L"content_pending_elided: expected anyPending=true while compares are queued.");
        }

        const auto deadline =
            std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(20'000))};
        {
            std::unique_lock lock(mutex);
            cv.wait_until(lock, deadline, [&] { return contentDone; });
        }

        state.Require(
            contentDone,
            std::format(L"content_pending_elided: content compares did not complete. pending={} total={} completed={}", lastPending, lastTotal, lastComplete));
        state.Require(lastTotal == static_cast<uint64_t>(kFileCount), L"content_pending_elided: unexpected totalContentCompares.");
        state.Require(lastComplete == static_cast<uint64_t>(kFileCount), L"content_pending_elided: unexpected completedContentCompares.");

        const auto decisionFinal = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decisionFinal), L"content_pending_elided: final decision missing.");
        if (decisionFinal)
        {
            state.Require(decisionFinal->items.empty(), L"content_pending_elided: expected no surfaced items after equal content compares.");
            state.Require(decisionFinal->pendingContentCompareCount == 0u, L"content_pending_elided: expected pendingContentCompareCount=0 after completion.");
            state.Require(! decisionFinal->anyPending, L"content_pending_elided: expected anyPending=false after completion.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: content_pending_elided.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"setCompareEnabled",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: SetCompareEnabled(false) stops producing decisions; re-enabling resumes.
    if (const auto foldersOpt = CreateCaseFolders(root, L"setCompareEnabled"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"b.txt", "B"), L"Failed to create b.txt (right).");

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

        state.Require(session->IsCompareEnabled(), L"IsCompareEnabled should be true by default.");

        // When compare is disabled, ReadDirectoryInfo falls back to the base filesystem
        // and shows all files (no comparison filtering applied).
        session->SetCompareEnabled(false);
        state.Require(! session->IsCompareEnabled(), L"IsCompareEnabled should be false after SetCompareEnabled(false).");

        {
            const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
            const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

            const auto leftNames  = EnumerateDirectoryNames(fsLeft, folders.left, state);
            const auto rightNames = EnumerateDirectoryNames(fsRight, folders.right, state);

            // Disabled compare: both sides should see their own files unfiltered.
            state.Require(ContainsName(leftNames, L"a.txt"), L"setCompareEnabled: a.txt should be visible in left when compare is disabled.");
            state.Require(ContainsName(rightNames, L"b.txt"), L"setCompareEnabled: b.txt should be visible in right when compare is disabled.");
            // a.txt only exists on the left, b.txt only exists on the right â€” in enabled mode
            // they would be filtered to their own pane; disabled should expose them as-is.
            state.Require(! ContainsName(leftNames, L"b.txt"), L"setCompareEnabled: b.txt should not appear in the left pane.");
            state.Require(! ContainsName(rightNames, L"a.txt"), L"setCompareEnabled: a.txt should not appear in the right pane.");
        }

        session->SetCompareEnabled(true);
        state.Require(session->IsCompareEnabled(), L"IsCompareEnabled should be true after re-enabling.");

        // After re-enabling, decisions should be obtainable and filtering should be back.
        auto decision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decision), L"GetOrComputeDecision should succeed after re-enabling compare.");

        {
            const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
            const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

            const auto leftNames  = EnumerateDirectoryNames(fsLeft, folders.left, state);
            const auto rightNames = EnumerateDirectoryNames(fsRight, folders.right, state);

            // Re-enabled compare: only pane-relevant different items are shown.
            state.Require(ContainsName(leftNames, L"a.txt"), L"setCompareEnabled: a.txt should be shown in left pane after re-enable (only in left).");
            state.Require(! ContainsName(leftNames, L"b.txt"), L"setCompareEnabled: b.txt should not appear in left pane after re-enable.");
            state.Require(ContainsName(rightNames, L"b.txt"), L"setCompareEnabled: b.txt should be shown in right pane after re-enable (only in right).");
            state.Require(! ContainsName(rightNames, L"a.txt"), L"setCompareEnabled: a.txt should not appear in right pane after re-enable.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: setCompareEnabled.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"invalidateForPath",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: InvalidateForAbsolutePath invalidates only the targeted subtree.
    if (const auto foldersOpt = CreateCaseFolders(root, L"invalidateForPath"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"sub1"), L"Failed to create sub1 (left).");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"sub1"), L"Failed to create sub1 (right).");
        state.Require(SelfTest::WriteTextFile(folders.left / L"sub1" / L"f.txt", "X"), L"Failed to create sub1/f.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"sub1" / L"f.txt", "X"), L"Failed to create sub1/f.txt (right).");
        state.Require(SelfTest::EnsureDirectory(folders.left / L"sub2"), L"Failed to create sub2 (left).");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"sub2"), L"Failed to create sub2 (right).");
        state.Require(SelfTest::WriteTextFile(folders.left / L"sub2" / L"g.txt", "Y"), L"Failed to create sub2/g.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"sub2" / L"g.txt", "Y"), L"Failed to create sub2/g.txt (right).");

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

        // Warm up both subtrees.
        const auto decisionSub1Before = session->GetOrComputeDecision(std::filesystem::path(L"sub1"));
        const auto decisionSub2Before = session->GetOrComputeDecision(std::filesystem::path(L"sub2"));
        state.Require(static_cast<bool>(decisionSub1Before), L"sub1 decision missing before invalidate.");
        state.Require(static_cast<bool>(decisionSub2Before), L"sub2 decision missing before invalidate.");

        // Invalidate only sub1's absolute path.
        session->InvalidateForAbsolutePath(folders.left / L"sub1", /*includeSubtree=*/true);

        const auto decisionSub1After = session->GetOrComputeDecision(std::filesystem::path(L"sub1"));
        const auto decisionSub2After = session->GetOrComputeDecision(std::filesystem::path(L"sub2"));

        state.Require(static_cast<bool>(decisionSub1After), L"sub1 decision missing after invalidate.");
        state.Require(static_cast<bool>(decisionSub2After), L"sub2 decision missing after invalidate.");

        // Sub1 must be a different (newly computed) decision object.
        state.Require(decisionSub1After != decisionSub1Before, L"sub1 decision should be new after InvalidateForAbsolutePath.");
        // Sub2 must be the same cached object â€” it was not invalidated.
        state.Require(decisionSub2After == decisionSub2Before, L"sub2 decision should remain cached (not invalidated).");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: invalidateForPath.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"decisionUpdatedCallback",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: SetDecisionUpdatedCallback fires after Invalidate().
    if (const auto foldersOpt = CreateCaseFolders(root, L"decisionUpdatedCallback"))
    {
        const auto& folders = foldersOpt.value();
        // Use compareContent=true with same-size but byte-different files so a content-compare
        // job is enqueued and dispatched to a worker thread.  The callback fires on that worker
        // thread when the compare job completes (size-different files are short-circuited without
        // an async job and would never fire the callback).
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "AAAA"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "BBBB"), L"Failed to create a.txt (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent = true;

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

        std::atomic<int> callbackCount{0};
        session->SetDecisionUpdatedCallback([&]() noexcept { callbackCount.fetch_add(1, std::memory_order_relaxed); });

        // Trigger a scan so content-compare workers are started.
        static_cast<void>(session->GetOrComputeDecision(std::filesystem::path{}));

        // Wait up to 10 s for the callback to fire at least once, polling GetOrComputeDecision
        // to keep the scan driving (consistent with the WaitForContentCompare pattern).
        const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(SelfTest::ScaleTimeout(10'000));
        while (callbackCount.load(std::memory_order_relaxed) == 0 && std::chrono::steady_clock::now() < deadline)
        {
            static_cast<void>(session->GetOrComputeDecision(std::filesystem::path{}));
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
        }

        state.Require(callbackCount.load(std::memory_order_relaxed) > 0, L"DecisionUpdatedCallback must fire at least once after content compare completes.");

        // Unregister before session is destroyed to avoid dangling reference.
        session->SetDecisionUpdatedCallback(nullptr);
    }
    else
    {
        state.Require(false, L"Failed to create case folders: decisionUpdatedCallback.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"uiVersion",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: GetUiVersion increments on Invalidate() and after FlushPendingContentCompareUpdates().
    if (const auto foldersOpt = CreateCaseFolders(root, L"uiVersion"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "A"), L"Failed to create a.txt (right).");

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

        const uint64_t uiV0 = session->GetUiVersion();
        const uint64_t ver0 = session->GetVersion();

        session->Invalidate();
        const uint64_t uiV1 = session->GetUiVersion();
        const uint64_t ver1 = session->GetVersion();

        state.Require(uiV1 != uiV0, L"GetUiVersion should change after Invalidate().");
        state.Require(ver1 != ver0, L"GetVersion should change after Invalidate().");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: uiVersion.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"accessors",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Accessor getters return correct values after construction.
    if (const auto foldersOpt = CreateCaseFolders(root, L"accessors"))
    {
        const auto& folders = foldersOpt.value();
        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSize = true;

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

        state.Require(session->GetRoot(ComparePane::Left) == folders.left, L"GetRoot(Left) should match the left root passed to constructor.");
        state.Require(session->GetRoot(ComparePane::Right) == folders.right, L"GetRoot(Right) should match the right root passed to constructor.");
        state.Require(session->GetSettings().compareSize == settings.compareSize, L"GetSettings().compareSize should match the value passed to constructor.");

        // TryMakeRelative / ResolveAbsolute round-trip.
        const std::filesystem::path sub(L"subdir");
        const std::filesystem::path absLeft = folders.left / sub;
        const auto relOpt                   = session->TryMakeRelative(ComparePane::Left, absLeft);
        state.Require(relOpt.has_value(), L"TryMakeRelative should succeed for a path under the left root.");
        if (relOpt.has_value())
        {
            state.Require(relOpt.value() == sub, L"TryMakeRelative should return the expected relative path.");
            const std::filesystem::path resolved = session->ResolveAbsolute(ComparePane::Left, relOpt.value());
            state.Require(resolved == absLeft, L"ResolveAbsolute round-trip should match the original absolute path.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: accessors.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"plugin_path_math",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: TryMakeRelative / ResolveAbsolute behave correctly for plugin-style (forward-slash) paths.
    const std::filesystem::path pluginRoot(L"/a");
    auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, pluginRoot, pluginRoot, Common::Settings::CompareDirectoriesSettings{});

    const auto relRoot = session->TryMakeRelative(ComparePane::Left, std::filesystem::path(L"/a"));
    state.Require(relRoot.has_value(), L"plugin_path_math: TryMakeRelative should succeed for the root itself.");
    if (relRoot.has_value())
    {
        state.Require(relRoot.value().empty(), L"plugin_path_math: TryMakeRelative(root, root) should return empty relative path.");
    }

    const auto relOpt = session->TryMakeRelative(ComparePane::Left, std::filesystem::path(L"/a/b"));
    state.Require(relOpt.has_value(), L"plugin_path_math: TryMakeRelative should succeed for /a/b.");
    if (relOpt.has_value())
    {
        state.Require(relOpt.value().generic_wstring() == L"b", L"plugin_path_math: TryMakeRelative(/a, /a/b) should return 'b'.");

        const std::filesystem::path resolved = session->ResolveAbsolute(ComparePane::Left, relOpt.value());
        state.Require(resolved.generic_wstring() == L"/a/b", L"plugin_path_math: ResolveAbsolute(/a, b) should return /a/b.");
    }

    const std::filesystem::path resolvedRoot = session->ResolveAbsolute(ComparePane::Left, std::filesystem::path{});
    state.Require(resolvedRoot.generic_wstring() == L"/a", L"plugin_path_math: ResolveAbsolute(/a, '') should return /a.");

    const auto outside = session->TryMakeRelative(ComparePane::Left, std::filesystem::path(L"/x"));
    state.Require(! outside.has_value(), L"plugin_path_math: TryMakeRelative should return nullopt for an out-of-scope plugin path.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"connection_display_url",
                  [&](SelfTest::CaseState& state) noexcept
{
    Common::Settings::ConnectionProfile ftpAnonymous{};
    ftpAnonymous.pluginId     = L"builtin/file-system-ftp";
    ftpAnonymous.authMode     = Common::Settings::ConnectionAuthMode::Anonymous;
    ftpAnonymous.host         = L"ftp.example.test";
    ftpAnonymous.port         = 21;
    const std::wstring ftpUrl = ConnectionProfileUtils::BuildConnectionDisplayUrl(ftpAnonymous);
    state.Require(ftpUrl == L"ftp://ftp.example.test:21", L"FTP anonymous display URL should hide the anonymous user name.");

    Common::Settings::ConnectionProfile sftp{};
    sftp.pluginId              = L"builtin/file-system-sftp";
    sftp.authMode              = Common::Settings::ConnectionAuthMode::Password;
    sftp.userName              = L"alice";
    sftp.host                  = L"files.example.test";
    sftp.port                  = 2222;
    const std::wstring sftpUrl = ConnectionProfileUtils::BuildConnectionDisplayUrl(sftp);
    state.Require(sftpUrl == L"sftp://alice@files.example.test:2222", L"SFTP display URL should include user and port.");

    Common::Settings::ConnectionProfile oneDrivePersonal{};
    oneDrivePersonal.pluginId      = L"builtin/file-system-onedrive-personal";
    oneDrivePersonal.authMode      = Common::Settings::ConnectionAuthMode::OAuth2Pkce;
    const std::wstring oneDriveUrl = ConnectionProfileUtils::BuildConnectionDisplayUrl(oneDrivePersonal);
    state.Require(oneDriveUrl == L"onedrive://", L"OneDrive Personal display URL should still render a hostless scheme.");

    Common::Settings::ConnectionProfile googleDrive{};
    googleDrive.pluginId              = L"builtin/file-system-gdrive";
    googleDrive.authMode              = Common::Settings::ConnectionAuthMode::OAuth2Pkce;
    const std::wstring googleDriveUrl = ConnectionProfileUtils::BuildConnectionDisplayUrl(googleDrive);
    state.Require(googleDriveUrl == L"gdrive://", L"Google Drive display URL should still render a hostless scheme.");

    Common::Settings::ConnectionProfile s3{};
    s3.pluginId              = L"builtin/file-system-s3";
    s3.authMode              = Common::Settings::ConnectionAuthMode::Password;
    s3.host                  = L"us-east-1";
    const std::wstring s3Url = ConnectionProfileUtils::BuildConnectionDisplayUrl(s3);
    state.Require(s3Url == L"s3://us-east-1", L"S3 display URL should use the shared formatter.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"try_make_relative_outside_root",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: TryMakeRelative returns nullopt for an absolute path not under the root.
    if (const auto foldersOpt = CreateCaseFolders(root, L"try_make_relative_outside_root"))
    {
        const auto& folders = foldersOpt.value();
        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

        const std::filesystem::path outsideLeft = folders.left.parent_path();
        const auto relLeft                      = session->TryMakeRelative(ComparePane::Left, outsideLeft);
        state.Require(! relLeft.has_value(), L"TryMakeRelative should return nullopt for a path outside the left root.");

        const std::filesystem::path outsideRight = folders.right.parent_path();
        const auto relRight                      = session->TryMakeRelative(ComparePane::Right, outsideRight);
        state.Require(! relRight.has_value(), L"TryMakeRelative should return nullopt for a path outside the right root.");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: try_make_relative_outside_root.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"baseInterfaces",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Base interface accessors return non-null objects after construction.
    if (const auto foldersOpt = CreateCaseFolders(root, L"baseInterfaces"))
    {
        const auto& folders = foldersOpt.value();
        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

        state.Require(static_cast<bool>(session->GetFileSystem(ComparePane::Left)), L"GetFileSystem(Left) should return non-null.");
        state.Require(static_cast<bool>(session->GetFileSystem(ComparePane::Right)), L"GetFileSystem(Right) should return non-null.");
        state.Require(static_cast<bool>(session->GetInformations(ComparePane::Left)), L"GetInformations(Left) should return non-null.");
        state.Require(static_cast<bool>(session->GetInformations(ComparePane::Right)), L"GetInformations(Right) should return non-null.");
        state.Require(static_cast<bool>(session->GetFileSystemIO(ComparePane::Left)), L"GetFileSystemIO(Left) should return non-null.");
        state.Require(static_cast<bool>(session->GetFileSystemIO(ComparePane::Right)), L"GetFileSystemIO(Right) should return non-null.");
        state.Require(session->IsContentCompareSupported(), L"IsContentCompareSupported() should return true.");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: baseInterfaces.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"contentCacheHit",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Repeated GetOrComputeDecision without invalidation returns the same cached object.
    if (const auto foldersOpt = CreateCaseFolders(root, L"contentCacheHit"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "CacheA"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "CacheA"), L"Failed to create a.txt (right).");

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

        const auto decision1 = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decision1), L"First call should return a valid decision.");
        const auto decision2 = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decision2), L"Second call should return a valid decision.");

        // Without any intervening Invalidate(), both calls must return the identical cached shared_ptr.
        state.Require(decision1 == decision2, L"Repeated GetOrComputeDecision without invalidation must return the same cached decision.");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: contentCacheHit.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"root_decision_empty_directories",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Empty directory roots produce an empty decision.
    if (const auto foldersOpt = CreateCaseFolders(root, L"root_decision_empty_directories"))
    {
        const auto& folders = foldersOpt.value();

        auto decision = ComputeRootDecision(baseFs, folders, Common::Settings::CompareDirectoriesSettings{}, state);
        if (decision)
        {
            state.Require(decision->items.empty(), L"Empty roots expected decision.items empty.");
            state.Require(! decision->anyDifferent, L"Empty roots expected anyDifferent=false.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: root_decision_empty_directories.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"zeroByteContent",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: compareContent=true on two zero-byte files reports them as identical.
    if (const auto foldersOpt = CreateCaseFolders(root, L"zeroByteContent"))
    {
        const auto& folders = foldersOpt.value();
        // Create empty files on both sides.
        state.Require(SelfTest::WriteBinaryFile(folders.left / L"empty.txt", {}), L"Failed to create empty.txt (left).");
        state.Require(SelfTest::WriteBinaryFile(folders.right / L"empty.txt", {}), L"Failed to create empty.txt (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent     = true;
        settings.keepIdenticalItems = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"empty.txt");
            state.Require(item != nullptr, L"empty.txt should appear in the decision.");
            if (item)
            {
                state.Require(! item->isDifferent, L"Zero-byte files on both sides must be identical.");
                state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"Zero-byte files must not have the Content diff bit set.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: zeroByteContent.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"setSettingsInvalidates",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: SetSettings with a comparison-changing setting increments GetVersion(); a view-only toggle does not.
    if (const auto foldersOpt = CreateCaseFolders(root, L"setSettingsInvalidates"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "V"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "V"), L"Failed to create a.txt (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent = false;
        auto session            = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

        const uint64_t v0 = session->GetVersion();

        // keepIdenticalItems affects cached decision shape (pruning), so it must invalidate.
        settings.keepIdenticalItems = true;
        session->SetSettings(settings);
        const uint64_t v1 = session->GetVersion();
        state.Require(v1 != v0, L"SetSettings with keepIdenticalItems toggled must increment GetVersion().");

        // showIdenticalItems is view-only and must NOT invalidate when keepIdenticalItems is unchanged.
        settings.showIdenticalItems = true;
        session->SetSettings(settings);
        const uint64_t v2 = session->GetVersion();
        state.Require(v2 == v1, L"SetSettings with showIdenticalItems toggled must not increment GetVersion().");

        // Changing compareContent must invalidate the cache (version bump).
        settings.compareContent = true;
        session->SetSettings(settings);
        const uint64_t v3 = session->GetVersion();
        state.Require(v3 != v2, L"SetSettings with compareContent toggled must increment GetVersion().");

        // Setting the same value again must NOT bump the version.
        session->SetSettings(settings);
        const uint64_t v4 = session->GetVersion();
        state.Require(v4 == v3, L"SetSettings with identical settings must not increment GetVersion().");

        // Changing compareSize must also invalidate.
        settings.compareSize = ! settings.compareSize;
        session->SetSettings(settings);
        const uint64_t v5 = session->GetVersion();
        state.Require(v5 != v4, L"SetSettings with compareSize toggled must increment GetVersion().");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: setSettingsInvalidates.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"dircache_not_polluted_by_compare_scan",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Compare scans must not populate DirectoryInfoCache (memory blow-up regression guard).
    if (const auto foldersOpt = CreateCaseFolders(root, L"dircache_not_polluted_by_compare_scan"))
    {
        const auto& folders = foldersOpt.value();

        constexpr size_t kDirCount  = 64;
        constexpr size_t kFileCount = 10;

        for (size_t i = 0; i < kDirCount; ++i)
        {
            const std::wstring dirName = std::format(L"d{:03}", i);

            const std::filesystem::path leftSub  = folders.left / dirName / L"sub";
            const std::filesystem::path rightSub = folders.right / dirName / L"sub";

            state.Require(SelfTest::EnsureDirectory(leftSub), std::format(L"Failed to create {} (left).", leftSub.wstring()));
            state.Require(SelfTest::EnsureDirectory(rightSub), std::format(L"Failed to create {} (right).", rightSub.wstring()));

            for (size_t j = 0; j < kFileCount; ++j)
            {
                const std::wstring fileName = std::format(L"f{:03}.txt", j);
                state.Require(SelfTest::WriteTextFile(leftSub / fileName, "X"), std::format(L"Failed to create {} (left).", fileName));
                state.Require(SelfTest::WriteTextFile(rightSub / fileName, "X"), std::format(L"Failed to create {} (right).", fileName));
            }
        }

        DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
        cache.ClearForFileSystem(baseFs.get());

        const DirectoryInfoCache::Stats before = cache.GetStats();

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSubdirectories = true;

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(30'000)}),
                      L"DirectoryInfoCache regression: scan did not become idle within timeout.");

        const DirectoryInfoCache::Stats after = cache.GetStats();

        const uint64_t deltaEnumerations = after.enumerations - before.enumerations;
        state.Require(deltaEnumerations == 0u, std::format(L"DirectoryInfoCache regression: expected enumerations delta=0, got {}.", deltaEnumerations));
    }
    else
    {
        state.Require(false, L"Failed to create case folders: dircache_not_polluted_by_compare_scan.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"content_queue_bounded_hi_lo",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Content compare job queues are bounded (hi/lo) under load; scan workers backpressure instead of OOM.
    if (const auto foldersOpt = CreateCaseFolders(root, L"content_queue_bounded_hi_lo"))
    {
        const auto& folders = foldersOpt.value();

        constexpr size_t kRootFiles = 1000;
        constexpr size_t kHotFiles  = 200;

        state.Require(SelfTest::EnsureDirectory(folders.left / L"hot"), L"Failed to create hot (left).");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"hot"), L"Failed to create hot (right).");

        for (size_t i = 0; i < kRootFiles; ++i)
        {
            const std::wstring name = std::format(L"r{:04}.bin", i);
            state.Require(SelfTest::WriteTextFile(folders.left / name, "AAAA"), std::format(L"Failed to write {} (left).", name));
            state.Require(SelfTest::WriteTextFile(folders.right / name, "BBBB"), std::format(L"Failed to write {} (right).", name));
        }

        for (size_t i = 0; i < kHotFiles; ++i)
        {
            const std::wstring name = std::format(L"h{:04}.bin", i);
            state.Require(SelfTest::WriteTextFile(folders.left / L"hot" / name, "AAAA"), std::format(L"Failed to write hot\\{} (left).", name));
            state.Require(SelfTest::WriteTextFile(folders.right / L"hot" / name, "BBBB"), std::format(L"Failed to write hot\\{} (right).", name));
        }

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent     = true;
        settings.compareSize        = false;
        settings.compareDateTime    = false;
        settings.compareAttributes  = false;
        settings.keepIdenticalItems = true;

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

        std::mutex mutex;
        std::condition_variable cv;
        bool started = false;
        bool done    = false;

        session->SetScanProgressCallback(
            [&](const std::filesystem::path&, std::wstring_view, uint64_t, uint64_t, uint32_t activeScans, uint64_t, uint64_t) noexcept
        {
            std::lock_guard lock(mutex);
            if (activeScans != 0u)
            {
                started = true;
            }
            if (started && activeScans == 0u)
            {
                done = true;
                cv.notify_all();
            }
        });

        session->RequestScanForFolder(std::filesystem::path(L"hot"));
        session->StartScan();

        {
            std::unique_lock lock(mutex);
            static_cast<void>(cv.wait_for(lock, std::chrono::milliseconds{SelfTest::ScaleTimeout(60'000)}, [&] { return done; }));
        }

        session->SetScanProgressCallback({});
        state.Require(done, L"content_queue_bounded_hi_lo: scan did not become idle within timeout.");

        const CompareDirectoriesPerfStats stats = session->GetPerfStats();

        constexpr size_t kMaxHighJobs = 128;
        constexpr size_t kMaxLowJobs  = 896;
        constexpr size_t kMaxTotal    = kMaxHighJobs + kMaxLowJobs;

        state.Require(stats.contentQueueHighHighWater <= kMaxHighJobs,
                      std::format(L"High content queue exceeded cap: {} > {}.", stats.contentQueueHighHighWater, kMaxHighJobs));
        state.Require(stats.contentQueueLowHighWater <= kMaxLowJobs,
                      std::format(L"Low content queue exceeded cap: {} > {}.", stats.contentQueueLowHighWater, kMaxLowJobs));
        state.Require(stats.contentQueueHighWater <= kMaxTotal,
                      std::format(L"Total content queue exceeded cap: {} > {}.", stats.contentQueueHighWater, kMaxTotal));
    }
    else
    {
        state.Require(false, L"Failed to create case folders: content_queue_bounded_hi_lo.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"decision_cache_eviction_budget_pins_visible",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Decision cache eviction respects pinned visible folders (prevents UI thrash).
    if (const auto foldersOpt = CreateCaseFolders(root, L"decision_cache_eviction_budget_pins_visible"))
    {
        const auto& folders = foldersOpt.value();

        state.Require(SelfTest::EnsureDirectory(folders.left / L"keep"), L"Failed to create keep (left).");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"keep"), L"Failed to create keep (right).");
        state.Require(SelfTest::WriteTextFile(folders.left / L"keep" / L"keep.txt", "K"), L"Failed to write keep.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"keep" / L"keep.txt", "K"), L"Failed to write keep.txt (right).");

        state.Require(SelfTest::EnsureDirectory(folders.left / L"spill"), L"Failed to create spill (left).");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"spill"), L"Failed to create spill (right).");

        constexpr size_t kSpillDirs   = 120;
        constexpr size_t kFilesPerDir = 10;

        for (size_t i = 0; i < kSpillDirs; ++i)
        {
            const std::wstring dirName           = std::format(L"d{:04}", i);
            const std::filesystem::path leftDir  = folders.left / L"spill" / dirName;
            const std::filesystem::path rightDir = folders.right / L"spill" / dirName;

            state.Require(SelfTest::EnsureDirectory(leftDir), std::format(L"Failed to create {} (left).", leftDir.wstring()));
            state.Require(SelfTest::EnsureDirectory(rightDir), std::format(L"Failed to create {} (right).", rightDir.wstring()));

            for (size_t j = 0; j < kFilesPerDir; ++j)
            {
                const std::wstring fileName = std::format(L"f{:03}.txt", j);
                state.Require(SelfTest::WriteTextFile(leftDir / fileName, "X"), std::format(L"Failed to write spill file {} (left).", fileName));
                state.Require(SelfTest::WriteTextFile(rightDir / fileName, "X"), std::format(L"Failed to write spill file {} (right).", fileName));
            }
        }

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.keepIdenticalItems = true;

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        session->SetDecisionCacheBudgetBytesForSelfTest(64u * 1024u);
        session->SetPinnedFolders(std::filesystem::path(L"keep"), std::filesystem::path(L"keep"));

        std::mutex mutex;
        std::condition_variable cv;
        bool started = false;
        bool done    = false;

        session->SetScanProgressCallback(
            [&](const std::filesystem::path&, std::wstring_view, uint64_t, uint64_t, uint32_t activeScans, uint64_t, uint64_t) noexcept
        {
            std::lock_guard lock(mutex);
            if (activeScans != 0u)
            {
                started = true;
            }
            if (started && activeScans == 0u)
            {
                done = true;
                cv.notify_all();
            }
        });

        session->RequestScanForFolder(std::filesystem::path(L"keep"));
        for (size_t i = 0; i < kSpillDirs; ++i)
        {
            const std::wstring dirName = std::format(L"d{:04}", i);
            session->RequestScanForFolder(std::filesystem::path(L"spill") / dirName);
        }
        session->StartScan();

        {
            std::unique_lock lock(mutex);
            static_cast<void>(cv.wait_for(lock, std::chrono::milliseconds{SelfTest::ScaleTimeout(60'000)}, [&] { return done; }));
        }

        session->SetScanProgressCallback({});
        state.Require(done, L"decision_cache_eviction_budget_pins_visible: scan did not become idle within timeout.");

        const CompareDirectoriesPerfStats stats = session->GetPerfStats();
        state.Require(stats.decisionCacheEntriesHighWater > stats.decisionCacheEntries,
                      L"decision cache eviction expected (high-water should exceed current entries).");

        const auto keepDecision = session->TryGetCachedDecision(std::filesystem::path(L"keep"));
        state.Require(static_cast<bool>(keepDecision), L"Pinned folder decision (keep) must remain cached.");

        size_t evicted = 0;
        for (size_t i = 0; i < 16; ++i)
        {
            const std::wstring dirName = std::format(L"d{:04}", i);
            if (! session->TryGetCachedDecision(std::filesystem::path(L"spill") / dirName))
            {
                ++evicted;
            }
        }
        state.Require(evicted != 0u, L"Expected at least one spill folder decision to be evicted under small budget.");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: decision_cache_eviction_budget_pins_visible.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"cancel_completes_bounded",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Disabling background work cancels scan/content work promptly (exit/cancel regression guard).
    if (const auto foldersOpt = CreateCaseFolders(root, L"cancel_completes_bounded"))
    {
        const auto& folders = foldersOpt.value();

        constexpr size_t kBytes = 16u * 1024u * 1024u;
        state.Require(WriteFileFill(folders.left / L"big.bin", 'A', kBytes), L"Failed to create big.bin (left).");
        state.Require(WriteFileFill(folders.right / L"big.bin", 'B', kBytes), L"Failed to create big.bin (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent    = true;
        settings.compareSize       = false;
        settings.compareDateTime   = false;
        settings.compareAttributes = false;

        wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left, 4096u, 1u);
        state.Require(static_cast<bool>(wrapped), L"Failed to create short-read file system wrapper (cancel).");

        const wil::com_ptr<IFileSystem> compareFs = wrapped ? wrapped : baseFs;
        auto session                              = std::make_shared<CompareDirectoriesSession>(compareFs, compareFs, folders.left, folders.right, settings);

        static_cast<void>(session->GetOrComputeDecision(std::filesystem::path{}));

        const auto startedDeadline =
            std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(5'000))};
        bool sawPending = false;
        while (std::chrono::steady_clock::now() < startedDeadline)
        {
            const CompareDirectoriesPerfStats stats = session->GetPerfStats();
            if (stats.contentPendingCompares != 0u || stats.contentInFlightSize != 0u || stats.contentQueueSize != 0u)
            {
                sawPending = true;
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
        }

        state.Require(sawPending, L"cancel_completes_bounded: expected pending content compare work.");

        session->SetBackgroundWorkEnabled(false);

        const auto doneDeadline =
            std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(10'000))};
        bool canceled = false;
        while (std::chrono::steady_clock::now() < doneDeadline)
        {
            const CompareDirectoriesPerfStats stats = session->GetPerfStats();
            if (stats.scanActiveScans == 0u && stats.contentPendingCompares == 0u && stats.contentQueueSize == 0u && stats.contentInFlightSize == 0u)
            {
                canceled = true;
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
        }

        state.Require(canceled, L"cancel_completes_bounded: background work did not cancel/drain within timeout.");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: cancel_completes_bounded.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"invalid_directory_entry_buffer",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Corrupt FileInfo buffers are rejected with ERROR_INVALID_DATA (no crash / OOB walk).
    if (const auto foldersOpt = CreateCaseFolders(root, L"invalid_directory_entry_buffer"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "L"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "R"), L"Failed to create a.txt (right).");

        ReadDirectoryTestBehavior behavior{};
        behavior.targetPath      = folders.left;
        behavior.returnMalformed = true;

        wil::com_ptr<IFileSystem> malformedLeft = CreateReadDirectoryBehaviorFileSystem(baseFs, behavior);
        state.Require(static_cast<bool>(malformedLeft), L"Failed to create malformed ReadDirectoryInfo wrapper.");

        auto session = std::make_shared<CompareDirectoriesSession>(
            malformedLeft ? malformedLeft : baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

        const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decision), L"invalid_directory_entry_buffer: decision is null.");
        if (decision)
        {
            state.Require(FAILED(decision->hr), L"invalid_directory_entry_buffer: expected failure HRESULT.");
            state.Require(decision->hr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
                          std::format(L"invalid_directory_entry_buffer: expected ERROR_INVALID_DATA, got 0x{:08X}.", static_cast<unsigned int>(decision->hr)));
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: invalid_directory_entry_buffer.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"scan_inflight_stamp_guards_restart",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Stale scan worker completion must not clear active tracking for a restarted run.
    if (const auto foldersOpt = CreateCaseFolders(root, L"scan_inflight_stamp_guards_restart"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "L"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "R"), L"Failed to create a.txt (right).");

        ReadDirectoryTestBehavior behavior{};
        behavior.targetPath = folders.left;
        behavior.delayMs    = static_cast<DWORD>(SelfTest::ScaleTimeout(1200));

        wil::com_ptr<IFileSystem> delayedLeft = CreateReadDirectoryBehaviorFileSystem(baseFs, behavior);
        state.Require(static_cast<bool>(delayedLeft), L"Failed to create delayed ReadDirectoryInfo wrapper.");

        auto session = std::make_shared<CompareDirectoriesSession>(
            delayedLeft ? delayedLeft : baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});

        session->StartScan();

        const auto startDeadline =
            std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(5'000))};
        bool sawActive = false;
        while (std::chrono::steady_clock::now() < startDeadline)
        {
            if (session->GetPerfStats().scanActiveScans != 0u)
            {
                sawActive = true;
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
        }
        state.Require(sawActive, L"scan_inflight_stamp_guards_restart: initial scan did not become active.");

        std::this_thread::sleep_for(std::chrono::milliseconds(SelfTest::ScaleTimeout(800)));

        session->SetBackgroundWorkEnabled(false);
        session->InvalidateForAbsolutePath(folders.left, true);
        session->SetBackgroundWorkEnabled(true);
        session->StartScan();

        // This point is after stale run completion should have happened, but before the restarted run can finish.
        std::this_thread::sleep_for(std::chrono::milliseconds(SelfTest::ScaleTimeout(1000)));
        const CompareDirectoriesPerfStats midStats = session->GetPerfStats();
        state.Require(midStats.scanActiveScans != 0u, L"scan_inflight_stamp_guards_restart: restarted scan lost active tracking before completion.");

        const auto doneDeadline =
            std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(15'000))};
        bool idle = false;
        while (std::chrono::steady_clock::now() < doneDeadline)
        {
            const CompareDirectoriesPerfStats stats = session->GetPerfStats();
            if (stats.scanActiveScans == 0u && stats.scanQueueSize == 0u && stats.scanInFlightKeys == 0u)
            {
                idle = true;
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
        }
        state.Require(idle, L"scan_inflight_stamp_guards_restart: restarted scan did not become idle within timeout.");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: scan_inflight_stamp_guards_restart.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"content_inflight_stamp_guards_restart",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Stale content worker completion must not consume a restarted run's in-flight slot.
    if (const auto foldersOpt = CreateCaseFolders(root, L"content_inflight_stamp_guards_restart"))
    {
        const auto& folders     = foldersOpt.value();
        constexpr size_t kBytes = 8u * 1024u * 1024u;
        state.Require(WriteFileFill(folders.left / L"a.bin", 'A', kBytes), L"Failed to create a.bin (left).");
        state.Require(WriteFileFill(folders.right / L"a.bin", 'B', kBytes), L"Failed to create a.bin (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent     = true;
        settings.compareSize        = false;
        settings.compareDateTime    = false;
        settings.compareAttributes  = false;
        settings.keepIdenticalItems = true;

        wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left, 4096u, static_cast<DWORD>(SelfTest::ScaleTimeout(2)));
        state.Require(static_cast<bool>(wrapped), L"Failed to create short-read file system wrapper (content restart stamp).");

        const wil::com_ptr<IFileSystem> compareFs = wrapped ? wrapped : baseFs;
        auto session                              = std::make_shared<CompareDirectoriesSession>(compareFs, compareFs, folders.left, folders.right, settings);

        static_cast<void>(session->GetOrComputeDecision(std::filesystem::path{}));

        const auto startDeadline =
            std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(10'000))};
        bool sawInFlight = false;
        while (std::chrono::steady_clock::now() < startDeadline)
        {
            const CompareDirectoriesPerfStats stats = session->GetPerfStats();
            if (stats.contentPendingCompares != 0u && stats.contentInFlightSize != 0u)
            {
                sawInFlight = true;
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
        }
        state.Require(sawInFlight, L"content_inflight_stamp_guards_restart: initial content compare did not become active.");

        session->SetBackgroundWorkEnabled(false);
        session->InvalidateForAbsolutePath(folders.left, true);
        session->SetBackgroundWorkEnabled(true);
        session->StartScan();

        const auto restartedDecision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(restartedDecision), L"content_inflight_stamp_guards_restart: restarted decision is null.");
        if (restartedDecision)
        {
            const auto* restartedItem = FindItem(*restartedDecision, L"a.bin");
            state.Require(restartedItem != nullptr, L"content_inflight_stamp_guards_restart: a.bin missing from restarted decision.");
            if (restartedItem)
            {
                state.Require(HasFlag(restartedItem->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                              L"content_inflight_stamp_guards_restart: restarted decision did not re-enter ContentPending.");
            }
        }

        const auto restartedDeadline =
            std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(10'000))};
        bool restartedObserved = false;
        while (std::chrono::steady_clock::now() < restartedDeadline)
        {
            const CompareDirectoriesPerfStats stats = session->GetPerfStats();
            if (stats.contentPendingCompares != 0u || stats.contentInFlightSize != 0u)
            {
                restartedObserved = true;
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
        }
        state.Require(restartedObserved, L"content_inflight_stamp_guards_restart: restarted content compare did not become pending or in-flight.");

        const auto doneDeadline =
            std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(30'000))};
        bool done = false;
        while (std::chrono::steady_clock::now() < doneDeadline)
        {
            const CompareDirectoriesPerfStats stats = session->GetPerfStats();
            if (stats.scanActiveScans == 0u && stats.contentPendingCompares == 0u && stats.contentQueueSize == 0u && stats.contentInFlightSize == 0u)
            {
                done = true;
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
        }
        state.Require(done, L"content_inflight_stamp_guards_restart: compare did not complete within timeout.");

        session->FlushPendingContentCompareUpdates();
        const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decision), L"content_inflight_stamp_guards_restart: final decision is null.");
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.bin");
            state.Require(item != nullptr, L"content_inflight_stamp_guards_restart: a.bin missing from final decision.");
            if (item)
            {
                state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                              L"content_inflight_stamp_guards_restart: ContentPending should be cleared after completion.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content),
                              L"content_inflight_stamp_guards_restart: expected Content difference after completion.");
                state.Require(item->isDifferent, L"content_inflight_stamp_guards_restart: a.bin should be marked different.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: content_inflight_stamp_guards_restart.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"directory_size_local_callback_contract",
                  [&](SelfTest::CaseState& state) noexcept
{
    if (const auto foldersOpt = CreateCaseFolders(root, L"directory_size_local_callback_contract"))
    {
        const auto& folders                  = foldersOpt.value();
        const std::filesystem::path filePath = folders.left / L"probe.bin";
        state.Require(WriteFileFill(filePath, 'L', 13u), L"Directory size local callback: failed to create probe file.");

        wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
        state.Require(CreateFileSystemDirectoryOperations(baseFs, dirOps), L"Directory size local callback: missing IFileSystemDirectoryOperations.");
        if (dirOps)
        {
            static_cast<void>(ValidateDirectorySizeCallbackContract(state, dirOps.get(), filePath.wstring()));
        }
    }
    else
    {
        state.Require(false, L"Directory size local callback: failed to create case folders.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"directory_size_dummy_callback_contract",
                  [&](SelfTest::CaseState& state) noexcept
{
    state.Require(dummyFs && dummyIo && dummyOps, L"Directory size dummy callback: dummy filesystem setup is unavailable.");
    if (! dummyFs || ! dummyIo || ! dummyOps)
    {
        return false;
    }

    const std::filesystem::path dirPath(std::format(L"/directory-size-dummy-{}", guid));
    const std::filesystem::path filePath = dirPath / L"probe.txt";
    const std::wstring fileText          = ToPluginPathText(filePath);
    const auto cleanup                   = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(dummyFs->DeleteItem(
            dirPath.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, nullptr, nullptr));
    });

    state.Require(EnsureDirectoryExistsFsOps(dummyOps, dirPath), L"Directory size dummy callback: failed to create dummy sandbox directory.");
    state.Require(WriteFileTextFsIo(dummyIo, filePath, "dummy-directory-size"), L"Directory size dummy callback: failed to write probe file.");
    if (! state.failure.empty())
    {
        return false;
    }

    static_cast<void>(ValidateDirectorySizeCallbackContract(state, dummyOps.get(), fileText));
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"directory_size_7z_callback_contract",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance archiveCreated{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltin7zFileSystemId, {}, archiveCreated);
    state.Require(SUCCEEDED(createHr) && archiveCreated.fileSystem,
                  std::format(L"Directory size 7z callback: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! archiveCreated.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemInitialize> init;
    const HRESULT initQiHr = archiveCreated.fileSystem->QueryInterface(IID_PPV_ARGS(init.put()));
    state.Require(SUCCEEDED(initQiHr) && init, L"Directory size 7z callback: missing IFileSystemInitialize.");
    if (FAILED(initQiHr) || ! init)
    {
        return false;
    }

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    const HRESULT dirQiHr = archiveCreated.fileSystem->QueryInterface(IID_PPV_ARGS(dirOps.put()));
    state.Require(SUCCEEDED(dirQiHr) && dirOps, L"Directory size 7z callback: missing IFileSystemDirectoryOperations.");
    if (FAILED(dirQiHr) || ! dirOps)
    {
        return false;
    }

    const std::filesystem::path archivePath = GetWorkspaceRootFromSourcePath() / L"Plugins" / L"FileSystem7z" / L"Tests" / L"Tests.zip";
    state.Require(SelfTest::PathExists(archivePath), std::format(L"Directory size 7z callback: fixture archive missing: {}", archivePath.wstring()));
    if (! SelfTest::PathExists(archivePath))
    {
        return false;
    }

    const HRESULT initHr = init->Initialize(archivePath.c_str(), nullptr);
    state.Require(SUCCEEDED(initHr), std::format(L"Directory size 7z callback: Initialize failed. hr=0x{:08X}", static_cast<unsigned long>(initHr)));
    if (FAILED(initHr))
    {
        return false;
    }

    const auto filePathOpt = FindFirstRegularEntryPath(archiveCreated.fileSystem, L"/");
    state.Require(filePathOpt.has_value(), L"Directory size 7z callback: no regular file found in fixture archive.");
    if (! filePathOpt.has_value())
    {
        return false;
    }

    static_cast<void>(ValidateDirectorySizeCallbackContract(state, dirOps.get(), filePathOpt.value()));
    return state.failure.empty();
});

const auto runRemoteFileCompare = [&](std::wstring_view caseName,
                                      std::wstring_view protocolLabel,
                                      std::wstring_view envVarName,
                                      std::wstring_view defaultProfileName,
                                      std::wstring_view pluginId) noexcept
{
    if (! SelfTest::CaseFilterMatches(options.caseFilter, caseName))
    {
        return;
    }

    if (options.failFast && suite.failed != 0)
    {
        AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
        return;
    }

    const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(protocolLabel, envVarName, defaultProfileName, pluginId);
    if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
        return;
    }

    const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(protocolLabel, envVarName, defaultProfileName, pluginId);
    if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
        return;
    }

    SelfTest::RunCase(options,
                      suite,
                      caseName,
                      [&](SelfTest::CaseState& state) noexcept
    {
        const ResolvedRemoteProfile resolvedProfile        = ResolveRemoteConnectionProfile(envVarName, defaultProfileName, pluginId);
        const std::wstring& profileName                    = resolvedProfile.profileName;
        const Common::Settings::ConnectionProfile* profile = resolvedProfile.profile;
        state.Require(profile != nullptr, L"Remote compare: profile missing after preconditions passed.");
        if (! profile)
        {
            return false;
        }

        const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
        state.Require(! initialPath.empty() && initialPath[0] == L'/', L"Remote compare: initialPath must be an absolute plugin path.");
        if (initialPath.empty() || initialPath[0] != L'/')
        {
            return false;
        }

        const std::filesystem::path remoteRoot = std::filesystem::path(std::format(L"/@conn:{}{}", profileName, initialPath));

        const std::filesystem::path localRoot = root / std::wstring(caseName) / L"left";
        state.Require(SelfTest::EnsureDirectory(localRoot), L"Remote compare: failed to create local root folder.");

        const std::wstring uniqueName = std::format(L"only_left_{}.txt", guid);
        state.Require(SelfTest::WriteTextFile(localRoot / uniqueName, "L"), L"Remote compare: failed to write local test file.");

        CreatedFileSystemInstance remoteCreated{};
        const HRESULT createHr = TryCreateFileSystemInstance(pluginId, {}, remoteCreated);
        state.Require(SUCCEEDED(createHr),
                      std::format(L"Remote compare: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! remoteCreated.fileSystem)
        {
            return false;
        }

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent = true;

        auto session  = std::make_shared<CompareDirectoriesSession>(baseFs, remoteCreated.fileSystem, localRoot, remoteRoot, settings);
        auto decision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decision), L"Remote compare: decision is null.");
        if (! decision)
        {
            return false;
        }

        state.Require(SUCCEEDED(decision->hr), std::format(L"Remote compare: decision hr failed. hr=0x{:08X}", static_cast<unsigned long>(decision->hr)));
        state.Require(! decision->rightFolderMissing, L"Remote compare: remote root reported missing.");
        if (FAILED(decision->hr) || decision->rightFolderMissing)
        {
            return false;
        }

        const auto relLeftRoot = session->TryMakeRelative(ComparePane::Left, localRoot);
        state.Require(relLeftRoot.has_value() && relLeftRoot.value().empty(), L"Remote compare: TryMakeRelative(Left, root) should return empty.");

        const auto relRightRoot = session->TryMakeRelative(ComparePane::Right, remoteRoot);
        state.Require(relRightRoot.has_value() && relRightRoot.value().empty(), L"Remote compare: TryMakeRelative(Right, root) should return empty.");

        const auto* item = FindItem(*decision, uniqueName);
        state.Require(item != nullptr, L"Remote compare: unique local file missing from decision.");
        if (item)
        {
            state.Require(item->isDifferent, L"Remote compare: unique local file expected different.");
            state.Require(item->selectLeft && ! item->selectRight, L"Remote compare: unique local file expected selectLeft only.");
            state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft),
                          L"Remote compare: unique local file expected differenceMask=OnlyInLeft.");
        }

        return state.failure.empty();
    });
};

const auto runRemoteDirectorySizeCallbackContract = [&](std::wstring_view caseName,
                                                        std::wstring_view protocolLabel,
                                                        std::wstring_view envVarName,
                                                        std::wstring_view defaultProfileName,
                                                        std::wstring_view pluginId) noexcept
{
    if (! SelfTest::CaseFilterMatches(options.caseFilter, caseName))
    {
        return;
    }

    if (options.failFast && suite.failed != 0)
    {
        AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
        return;
    }

    const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(protocolLabel, envVarName, defaultProfileName, pluginId);
    if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
        return;
    }

    const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(protocolLabel, envVarName, defaultProfileName, pluginId);
    if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
        return;
    }

    SelfTest::RunCase(options,
                      suite,
                      caseName,
                      [&](SelfTest::CaseState& state) noexcept
    {
        const ResolvedRemoteProfile resolvedProfile        = ResolveRemoteConnectionProfile(envVarName, defaultProfileName, pluginId);
        const std::wstring& profileName                    = resolvedProfile.profileName;
        const Common::Settings::ConnectionProfile* profile = resolvedProfile.profile;
        state.Require(profile != nullptr, std::format(L"Remote {} directory size: profile missing after preconditions passed.", protocolLabel));
        if (! profile)
        {
            return false;
        }

        const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
        state.Require(! initialPath.empty() && initialPath[0] == L'/',
                      std::format(L"Remote {} directory size: initialPath must be an absolute plugin path.", protocolLabel));
        if (initialPath.empty() || initialPath[0] != L'/')
        {
            return false;
        }

        CreatedFileSystemInstance remoteCreated{};
        const HRESULT createHr = TryCreateFileSystemInstance(pluginId, {}, remoteCreated);
        state.Require(
            SUCCEEDED(createHr) && remoteCreated.fileSystem,
            std::format(L"Remote {} directory size: failed to create filesystem instance. hr=0x{:08X}", protocolLabel, static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! remoteCreated.fileSystem)
        {
            return false;
        }

        wil::com_ptr<IFileSystemIO> io;
        const HRESULT ioHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(io.put()));
        state.Require(SUCCEEDED(ioHr) && io, std::format(L"Remote {} directory size: filesystem instance missing IFileSystemIO.", protocolLabel));
        if (FAILED(ioHr) || ! io)
        {
            return false;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
        const HRESULT dirHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(dirOps.put()));
        state.Require(SUCCEEDED(dirHr) && dirOps,
                      std::format(L"Remote {} directory size: filesystem instance missing IFileSystemDirectoryOperations.", protocolLabel));
        if (FAILED(dirHr) || ! dirOps)
        {
            return false;
        }

        const std::wstring caseRootPlugin = JoinPluginPathForSelfTest(initialPath, std::format(L"compare-selftest-dirsize-{}-{}", protocolLabel, guid));
        state.Require(! caseRootPlugin.empty(), std::format(L"Remote {} directory size: failed to build sandbox path.", protocolLabel));
        if (caseRootPlugin.empty())
        {
            return false;
        }

        const std::wstring caseRootText = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
        state.Require(! caseRootText.empty(), std::format(L"Remote {} directory size: failed to build connection path.", protocolLabel));
        if (caseRootText.empty())
        {
            return false;
        }

        const std::filesystem::path caseRoot(caseRootText);

        state.Require(EnsureDirectoryExistsFsOps(dirOps, caseRoot), std::format(L"Remote {} directory size: failed to prepare sandbox prefix.", protocolLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        const std::filesystem::path objectPath = caseRoot / L"probe.txt";
        const std::wstring objectText          = ToPluginPathText(objectPath);
        const auto cleanup                     = wil::scope_exit([&]() noexcept
        {
            if (pluginId == kBuiltinS3FileSystemId)
            {
                static_cast<void>(remoteCreated.fileSystem->DeleteItem(objectText.c_str(), FILESYSTEM_FLAG_CONTINUE_ON_ERROR, nullptr, nullptr, nullptr));
                return;
            }

            static_cast<void>(remoteCreated.fileSystem->DeleteItem(
                caseRootText.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, nullptr, nullptr));
        });

        state.Require(WriteFileTextFsIo(io, objectPath, "directory-size-remote"),
                      std::format(L"Remote {} directory size: failed to write probe object.", protocolLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        static_cast<void>(ValidateDirectorySizeCallbackContract(state, dirOps.get(), objectText));
        return state.failure.empty();
    });
};

const auto runRemoteS3Pagination = [&](std::wstring_view caseName) noexcept
{
    if (! SelfTest::CaseFilterMatches(options.caseFilter, caseName))
    {
        return;
    }

    if (options.failFast && suite.failed != 0)
    {
        AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
        return;
    }

    const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
    if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
        return;
    }

    const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
    if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
        return;
    }

    SelfTest::RunCase(options,
                      suite,
                      caseName,
                      [&](SelfTest::CaseState& state) noexcept
    {
        const ResolvedRemoteProfile resolvedProfile        = ResolveRemoteConnectionProfile(kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
        const std::wstring& profileName                    = resolvedProfile.profileName;
        const Common::Settings::ConnectionProfile* profile = resolvedProfile.profile;
        state.Require(profile != nullptr, L"Remote S3 pagination: profile missing after preconditions passed.");
        if (! profile)
        {
            return false;
        }

        const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
        state.Require(! initialPath.empty() && initialPath[0] == L'/', L"Remote S3 pagination: initialPath must be an absolute plugin path.");
        if (initialPath.empty() || initialPath[0] != L'/')
        {
            return false;
        }

        CreatedFileSystemInstance remoteCreated{};
        const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinS3FileSystemId, {}, remoteCreated);
        state.Require(SUCCEEDED(createHr),
                      std::format(L"Remote S3 pagination: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! remoteCreated.fileSystem)
        {
            return false;
        }

        wil::com_ptr<IFileSystemIO> io;
        const HRESULT ioHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(io.put()));
        state.Require(SUCCEEDED(ioHr) && io, L"Remote S3 pagination: filesystem instance missing IFileSystemIO.");
        if (FAILED(ioHr) || ! io)
        {
            return false;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
        const HRESULT hrQI = remoteCreated.fileSystem->QueryInterface(__uuidof(IFileSystemDirectoryOperations), dirOps.put_void());
        state.Require(SUCCEEDED(hrQI) && dirOps, L"Remote S3 pagination: filesystem instance missing IFileSystemDirectoryOperations.");
        if (FAILED(hrQI) || ! dirOps)
        {
            return false;
        }

        const std::wstring caseRootPlugin = JoinPluginPathForSelfTest(initialPath, std::format(L"compare-selftest-s3-pagination-{}", guid));
        state.Require(! caseRootPlugin.empty(), L"Remote S3 pagination: failed to build sandbox path.");
        if (caseRootPlugin.empty())
        {
            return false;
        }

        const std::wstring caseRootText = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
        state.Require(! caseRootText.empty(), L"Remote S3 pagination: failed to build connection path.");
        if (caseRootText.empty())
        {
            return false;
        }

        const std::filesystem::path caseRoot(caseRootText);

        state.Require(EnsureDirectoryExistsFsOps(dirOps, caseRoot), L"Remote S3 pagination: failed to prepare sandbox prefix.");
        if (! state.failure.empty())
        {
            return false;
        }

        constexpr unsigned long kObjectCount = 1005u;
        constexpr std::string_view kPayload  = "p";
        std::vector<std::wstring> cleanupPaths;
        cleanupPaths.reserve(kObjectCount);
        std::vector<const wchar_t*> cleanupPathPointers;
        cleanupPathPointers.reserve(kObjectCount);
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (cleanupPathPointers.empty())
            {
                return;
            }

            static_cast<void>(remoteCreated.fileSystem->DeleteItems(cleanupPathPointers.data(),
                                                                    static_cast<unsigned long>(cleanupPathPointers.size()),
                                                                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                                    nullptr,
                                                                    nullptr,
                                                                    nullptr));
        });

        for (unsigned long index = 0; index < kObjectCount; ++index)
        {
            const std::filesystem::path objectPath = caseRoot / std::format(L"object_{:04}.txt", index);
            cleanupPaths.push_back(ToPluginPathText(objectPath));
            cleanupPathPointers.push_back(cleanupPaths.back().c_str());
            state.Require(WriteFileTextFsIo(io, objectPath, kPayload), std::format(L"Remote S3 pagination: failed to write object {}", objectPath.wstring()));
            if (! state.failure.empty())
            {
                return false;
            }
        }

        FileSystemDirectorySizeResult sizeResult{};
        sizeResult.sizeBytes = sizeof(FileSystemDirectorySizeResult);
        const HRESULT sizeHr = dirOps->GetDirectorySize(caseRootText.c_str(), FileSystemFlags{}, nullptr, nullptr, &sizeResult);
        state.Require(SUCCEEDED(sizeHr) && SUCCEEDED(sizeResult.status),
                      std::format(L"Remote S3 pagination: GetDirectorySize failed. hr=0x{:08X} status=0x{:08X}",
                                  static_cast<unsigned long>(sizeHr),
                                  static_cast<unsigned long>(sizeResult.status)));
        if (FAILED(sizeHr) || FAILED(sizeResult.status))
        {
            return false;
        }

        state.Require(sizeResult.fileCount == kObjectCount,
                      std::format(L"Remote S3 pagination: fileCount mismatch. expected={} got={}", kObjectCount, sizeResult.fileCount));
        state.Require(sizeResult.directoryCount == 0u, std::format(L"Remote S3 pagination: expected directoryCount=0, got {}.", sizeResult.directoryCount));
        state.Require(sizeResult.totalBytes == static_cast<uint64_t>(kObjectCount * kPayload.size()),
                      std::format(L"Remote S3 pagination: totalBytes mismatch. expected={} got={}.",
                                  static_cast<uint64_t>(kObjectCount * kPayload.size()),
                                  sizeResult.totalBytes));
        if (! state.failure.empty())
        {
            return false;
        }

        wil::com_ptr<IFilesInformation> listing;
        const HRESULT listHr = remoteCreated.fileSystem->ReadDirectoryInfo(caseRootText.c_str(), listing.put());
        state.Require(SUCCEEDED(listHr) && listing,
                      std::format(L"Remote S3 pagination: ReadDirectoryInfo failed. hr=0x{:08X}", static_cast<unsigned long>(listHr)));
        if (FAILED(listHr) || ! listing)
        {
            return false;
        }

        unsigned long got     = 0;
        const HRESULT countHr = listing->GetCount(&got);
        state.Require(SUCCEEDED(countHr), std::format(L"Remote S3 pagination: GetCount failed. hr=0x{:08X}", static_cast<unsigned long>(countHr)));
        if (FAILED(countHr))
        {
            return false;
        }

        state.Require(got == kObjectCount, std::format(L"Remote S3 pagination: listing count mismatch. expected={} got={}", kObjectCount, got));
        return state.failure.empty();
    });
};

const auto runRemoteFtpPartialContinue = [&](std::wstring_view caseName) noexcept
{
    if (! SelfTest::CaseFilterMatches(options.caseFilter, caseName))
    {
        return;
    }

    if (options.failFast && suite.failed != 0)
    {
        AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
        return;
    }

    const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kBuiltinFtpFileSystemId);
    if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
        return;
    }

    const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kBuiltinFtpFileSystemId);
    if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
        return;
    }

    SelfTest::RunCase(options,
                      suite,
                      caseName,
                      [&](SelfTest::CaseState& state) noexcept
    {
        const ResolvedRemoteProfile resolvedProfile = ResolveRemoteConnectionProfile(kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kBuiltinFtpFileSystemId);
        const std::wstring& profileName             = resolvedProfile.profileName;
        const Common::Settings::ConnectionProfile* profile = resolvedProfile.profile;
        state.Require(profile != nullptr, L"Remote FTP partial: profile missing after preconditions passed.");
        if (! profile)
        {
            return false;
        }

        const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
        state.Require(! initialPath.empty() && initialPath[0] == L'/', L"Remote FTP partial: initialPath must be an absolute plugin path.");
        if (initialPath.empty() || initialPath[0] != L'/')
        {
            return false;
        }

        CreatedFileSystemInstance remoteCreated{};
        const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinFtpFileSystemId, {}, remoteCreated);
        state.Require(SUCCEEDED(createHr),
                      std::format(L"Remote FTP partial: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! remoteCreated.fileSystem)
        {
            return false;
        }

        wil::com_ptr<IFileSystemIO> io;
        const HRESULT ioHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(io.put()));
        state.Require(SUCCEEDED(ioHr) && io, L"Remote FTP partial: filesystem instance missing IFileSystemIO.");
        if (FAILED(ioHr) || ! io)
        {
            return false;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
        const HRESULT dirHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(dirOps.put()));
        state.Require(SUCCEEDED(dirHr) && dirOps, L"Remote FTP partial: filesystem instance missing IFileSystemDirectoryOperations.");
        if (FAILED(dirHr) || ! dirOps)
        {
            return false;
        }

        const std::wstring caseRootPlugin = JoinPluginPathForSelfTest(initialPath, std::format(L"compare-selftest-ftp-partial-{}", guid));
        state.Require(! caseRootPlugin.empty(), L"Remote FTP partial: failed to build sandbox path.");
        if (caseRootPlugin.empty())
        {
            return false;
        }

        const std::wstring caseRootText = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
        state.Require(! caseRootText.empty(), L"Remote FTP partial: failed to build connection path.");
        if (caseRootText.empty())
        {
            return false;
        }

        const std::filesystem::path caseRoot(caseRootText);
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            const FileSystemFlags cleanupFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
            const std::wstring cleanupPath     = ToPluginPathText(caseRoot);
            static_cast<void>(remoteCreated.fileSystem->DeleteItem(cleanupPath.c_str(), cleanupFlags, nullptr, nullptr, nullptr));
        });

        const std::filesystem::path srcDir     = caseRoot / L"src";
        const std::filesystem::path copyDst    = caseRoot / L"copy-dst";
        const std::filesystem::path moveDst    = caseRoot / L"move-dst";
        const std::filesystem::path copyGood   = srcDir / L"copy-good.txt";
        const std::filesystem::path copyMiss   = srcDir / L"copy-missing.txt";
        const std::filesystem::path moveGood   = srcDir / L"move-good.txt";
        const std::filesystem::path moveMiss   = srcDir / L"move-missing.txt";
        const std::filesystem::path renameGood = srcDir / L"rename-good.txt";
        const std::filesystem::path renameMiss = srcDir / L"rename-missing.txt";

        state.Require(EnsureDirectoryExistsFsOps(dirOps, srcDir), L"Remote FTP partial: failed to create source directory.");
        state.Require(EnsureDirectoryExistsFsOps(dirOps, copyDst), L"Remote FTP partial: failed to create copy destination.");
        state.Require(EnsureDirectoryExistsFsOps(dirOps, moveDst), L"Remote FTP partial: failed to create move destination.");
        state.Require(WriteFileTextFsIo(io, copyGood, "copy-good"), L"Remote FTP partial: failed to write copy-good.txt.");
        state.Require(WriteFileTextFsIo(io, moveGood, "move-good"), L"Remote FTP partial: failed to write move-good.txt.");
        state.Require(WriteFileTextFsIo(io, renameGood, "rename-good"), L"Remote FTP partial: failed to write rename-good.txt.");
        if (! state.failure.empty())
        {
            return false;
        }

        const std::wstring copyGoodPluginPath = ToPluginPathText(copyGood);
        const char* copyGoodPropertiesJson    = nullptr;
        const HRESULT copyGoodPropertiesHr    = io->GetItemProperties(copyGoodPluginPath.c_str(), &copyGoodPropertiesJson);
        state.Require(SUCCEEDED(copyGoodPropertiesHr) && copyGoodPropertiesJson != nullptr && copyGoodPropertiesJson[0] != '\0',
                      std::format(L"Remote FTP partial: GetItemProperties failed for timestamp validation. hr=0x{:08X}",
                                  static_cast<unsigned long>(copyGoodPropertiesHr)));
        if (! state.failure.empty())
        {
            return false;
        }

        const std::string_view copyGoodProperties(copyGoodPropertiesJson);
        const auto findPropertyValue = [&](std::string_view key) noexcept -> std::optional<std::string_view>
        {
            const std::string needle = std::format("\"key\":\"{}\"", key);
            const size_t keyPos      = copyGoodProperties.find(needle);
            if (keyPos == std::string_view::npos)
            {
                return std::nullopt;
            }

            const size_t valuePos = copyGoodProperties.find("\"value\":\"", keyPos + needle.size());
            if (valuePos == std::string_view::npos)
            {
                return std::nullopt;
            }

            const size_t begin = valuePos + std::string_view("\"value\":\"").size();
            const size_t end   = copyGoodProperties.find('"', begin);
            if (end == std::string_view::npos || end < begin)
            {
                return std::nullopt;
            }

            return copyGoodProperties.substr(begin, end - begin);
        };

        const std::optional<std::string_view> ftpModifiedTime = findPropertyValue("lastWriteTime");
        state.Require(ftpModifiedTime.has_value() && ftpModifiedTime.value() != "0",
                      L"Remote FTP partial: item properties should expose a non-zero modified timestamp.");
        state.Require(! findPropertyValue("creationTime").has_value(), L"Remote FTP partial: item properties should not expose an unavailable zero creation time.");
        state.Require(! findPropertyValue("lastAccessTime").has_value(), L"Remote FTP partial: item properties should not expose an unavailable zero access time.");
        state.Require(! findPropertyValue("changeTime").has_value(), L"Remote FTP partial: item properties should not expose an unavailable zero change time.");
        if (! state.failure.empty())
        {
            return false;
        }

        const std::wstring copyGoodText = ToPluginPathText(copyGood);
        const std::wstring copyMissText = ToPluginPathText(copyMiss);
        const std::wstring copyDstText  = ToPluginPathText(copyDst);
        const wchar_t* copySources[]    = {copyGoodText.c_str(), copyMissText.c_str()};
        RecordingFileSystemCallback copyCallback(2);
        const HRESULT copyHr = remoteCreated.fileSystem->CopyItems(
            copySources, 2, copyDstText.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, &copyCallback, nullptr);
        const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        state.Require(copyHr == partialHr,
                      std::format(L"Remote FTP partial: CopyItems expected 0x{:08X}, got 0x{:08X}.",
                                  static_cast<unsigned long>(partialHr),
                                  static_cast<unsigned long>(copyHr)));
        state.Require(! copyCallback.SawUnexpectedIssue(), L"Remote FTP partial: CopyItems triggered an unexpected issue callback.");
        state.Require(copyCallback.CompletedCount() == 2u,
                      std::format(L"Remote FTP partial: CopyItems expected 2 completion callbacks, got {}.", copyCallback.CompletedCount()));
        RecordedFileSystemItem copyItem0{};
        RecordedFileSystemItem copyItem1{};
        state.Require(copyCallback.TryGetItem(0, copyItem0), L"Remote FTP partial: CopyItems missing callback for item 0.");
        state.Require(copyCallback.TryGetItem(1, copyItem1), L"Remote FTP partial: CopyItems missing callback for item 1.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(SUCCEEDED(copyItem0.status),
                      std::format(L"Remote FTP partial: CopyItems item 0 expected success, got 0x{:08X}.", static_cast<unsigned long>(copyItem0.status)));
        state.Require(
            copyItem1.status == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
            std::format(L"Remote FTP partial: CopyItems item 1 expected ERROR_FILE_NOT_FOUND, got 0x{:08X}.", static_cast<unsigned long>(copyItem1.status)));
        state.Require(PathExistsFsIo(io, copyGood), L"Remote FTP partial: CopyItems removed the successful source unexpectedly.");
        state.Require(PathExistsFsIo(io, copyDst / L"copy-good.txt"), L"Remote FTP partial: CopyItems did not materialize the successful destination file.");
        if (! state.failure.empty())
        {
            return false;
        }

        const std::wstring moveGoodText = ToPluginPathText(moveGood);
        const std::wstring moveMissText = ToPluginPathText(moveMiss);
        const std::wstring moveDstText  = ToPluginPathText(moveDst);
        const wchar_t* moveSources[]    = {moveGoodText.c_str(), moveMissText.c_str()};
        RecordingFileSystemCallback moveCallback(2);
        const HRESULT moveHr = remoteCreated.fileSystem->MoveItems(
            moveSources, 2, moveDstText.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, &moveCallback, nullptr);
        state.Require(moveHr == partialHr,
                      std::format(L"Remote FTP partial: MoveItems expected 0x{:08X}, got 0x{:08X}.",
                                  static_cast<unsigned long>(partialHr),
                                  static_cast<unsigned long>(moveHr)));
        state.Require(! moveCallback.SawUnexpectedIssue(), L"Remote FTP partial: MoveItems triggered an unexpected issue callback.");
        state.Require(moveCallback.CompletedCount() == 2u,
                      std::format(L"Remote FTP partial: MoveItems expected 2 completion callbacks, got {}.", moveCallback.CompletedCount()));
        RecordedFileSystemItem moveItem0{};
        RecordedFileSystemItem moveItem1{};
        state.Require(moveCallback.TryGetItem(0, moveItem0), L"Remote FTP partial: MoveItems missing callback for item 0.");
        state.Require(moveCallback.TryGetItem(1, moveItem1), L"Remote FTP partial: MoveItems missing callback for item 1.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(SUCCEEDED(moveItem0.status),
                      std::format(L"Remote FTP partial: MoveItems item 0 expected success, got 0x{:08X}.", static_cast<unsigned long>(moveItem0.status)));
        state.Require(
            moveItem1.status == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
            std::format(L"Remote FTP partial: MoveItems item 1 expected ERROR_FILE_NOT_FOUND, got 0x{:08X}.", static_cast<unsigned long>(moveItem1.status)));
        state.Require(! PathExistsFsIo(io, moveGood), L"Remote FTP partial: MoveItems left the successful source in place.");
        state.Require(PathExistsFsIo(io, moveDst / L"move-good.txt"), L"Remote FTP partial: MoveItems did not create the successful destination file.");
        if (! state.failure.empty())
        {
            return false;
        }

        const std::wstring renameGoodText = ToPluginPathText(renameGood);
        const std::wstring renameMissText = ToPluginPathText(renameMiss);
        FileSystemRenamePair renamePairs[2]{};
        renamePairs[0].sizeBytes  = sizeof(FileSystemRenamePair);
        renamePairs[0].sourcePath = renameGoodText.c_str();
        renamePairs[0].newName    = L"rename-good-dst.txt";
        renamePairs[1].sizeBytes  = sizeof(FileSystemRenamePair);
        renamePairs[1].sourcePath = renameMissText.c_str();
        renamePairs[1].newName    = L"rename-missing-dst.txt";

        RecordingFileSystemCallback renameCallback(2);
        const HRESULT renameHr = remoteCreated.fileSystem->RenameItems(
            renamePairs, 2, static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, &renameCallback, nullptr);
        state.Require(renameHr == partialHr,
                      std::format(L"Remote FTP partial: RenameItems expected 0x{:08X}, got 0x{:08X}.",
                                  static_cast<unsigned long>(partialHr),
                                  static_cast<unsigned long>(renameHr)));
        state.Require(! renameCallback.SawUnexpectedIssue(), L"Remote FTP partial: RenameItems triggered an unexpected issue callback.");
        state.Require(renameCallback.CompletedCount() == 2u,
                      std::format(L"Remote FTP partial: RenameItems expected 2 completion callbacks, got {}.", renameCallback.CompletedCount()));
        RecordedFileSystemItem renameItem0{};
        RecordedFileSystemItem renameItem1{};
        state.Require(renameCallback.TryGetItem(0, renameItem0), L"Remote FTP partial: RenameItems missing callback for item 0.");
        state.Require(renameCallback.TryGetItem(1, renameItem1), L"Remote FTP partial: RenameItems missing callback for item 1.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(SUCCEEDED(renameItem0.status),
                      std::format(L"Remote FTP partial: RenameItems item 0 expected success, got 0x{:08X}.", static_cast<unsigned long>(renameItem0.status)));
        state.Require(renameItem1.status == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                      std::format(L"Remote FTP partial: RenameItems item 1 expected ERROR_FILE_NOT_FOUND, got 0x{:08X}.",
                                  static_cast<unsigned long>(renameItem1.status)));
        state.Require(! PathExistsFsIo(io, renameGood), L"Remote FTP partial: RenameItems left the successful source in place.");
        state.Require(PathExistsFsIo(io, srcDir / L"rename-good-dst.txt"), L"Remote FTP partial: RenameItems did not create the successful destination name.");
        return state.failure.empty();
    });
};

const auto runRemoteS3MetadataSmoke = [&](std::wstring_view caseName) noexcept
{
    if (! SelfTest::CaseFilterMatches(options.caseFilter, caseName))
    {
        return;
    }

    if (options.failFast && suite.failed != 0)
    {
        AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
        return;
    }

    const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
    if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
        return;
    }

    const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
    if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
        return;
    }

    SelfTest::RunCase(options,
                      suite,
                      caseName,
                      [&](SelfTest::CaseState& state) noexcept
    {
        const ResolvedRemoteProfile resolvedProfile        = ResolveRemoteConnectionProfile(kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
        const std::wstring& profileName                    = resolvedProfile.profileName;
        const Common::Settings::ConnectionProfile* profile = resolvedProfile.profile;
        state.Require(profile != nullptr, L"Remote S3 metadata: profile missing after preconditions passed.");
        if (! profile)
        {
            return false;
        }

        const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
        state.Require(! initialPath.empty() && initialPath[0] == L'/', L"Remote S3 metadata: initialPath must be an absolute plugin path.");
        if (initialPath.empty() || initialPath[0] != L'/')
        {
            return false;
        }

        CreatedFileSystemInstance remoteCreated{};
        const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinS3FileSystemId, {}, remoteCreated);
        state.Require(SUCCEEDED(createHr),
                      std::format(L"Remote S3 metadata: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! remoteCreated.fileSystem)
        {
            return false;
        }

        wil::com_ptr<IFileSystemIO> io;
        const HRESULT ioHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(io.put()));
        state.Require(SUCCEEDED(ioHr) && io, L"Remote S3 metadata: filesystem instance missing IFileSystemIO.");
        if (FAILED(ioHr) || ! io)
        {
            return false;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
        const HRESULT dirHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(dirOps.put()));
        state.Require(SUCCEEDED(dirHr) && dirOps, L"Remote S3 metadata: filesystem instance missing IFileSystemDirectoryOperations.");
        if (FAILED(dirHr) || ! dirOps)
        {
            return false;
        }

        const std::wstring caseRootPlugin = JoinPluginPathForSelfTest(initialPath, std::format(L"compare-selftest-s3-metadata-{}", guid));
        state.Require(! caseRootPlugin.empty(), L"Remote S3 metadata: failed to build sandbox path.");
        if (caseRootPlugin.empty())
        {
            return false;
        }

        const std::wstring caseRootText = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
        state.Require(! caseRootText.empty(), L"Remote S3 metadata: failed to build connection path.");
        if (caseRootText.empty())
        {
            return false;
        }

        const std::filesystem::path caseRoot(caseRootText);
        state.Require(EnsureDirectoryExistsFsOps(dirOps, caseRoot), L"Remote S3 metadata: failed to prepare sandbox prefix.");
        if (! state.failure.empty())
        {
            return false;
        }

        const std::filesystem::path objectPath = caseRoot / L"metadata.txt";
        const std::wstring objectText          = ToPluginPathText(objectPath);
        const auto cleanup                     = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(remoteCreated.fileSystem->DeleteItem(
                objectText.c_str(), static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, nullptr, nullptr));
        });

        constexpr std::string_view kPayload = "s3-metadata-smoke";
        state.Require(WriteFileTextFsIo(io, objectPath, kPayload), L"Remote S3 metadata: failed to write test object.");
        if (! state.failure.empty())
        {
            return false;
        }

        FileSystemBasicInformation info{};
        info.sizeBytes       = sizeof(FileSystemBasicInformation);
        const HRESULT infoHr = io->GetFileBasicInformation(objectText.c_str(), &info);
        state.Require(SUCCEEDED(infoHr), std::format(L"Remote S3 metadata: GetFileBasicInformation failed. hr=0x{:08X}", static_cast<unsigned long>(infoHr)));
        if (FAILED(infoHr))
        {
            return false;
        }

        state.Require((info.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0, L"Remote S3 metadata: object reported as a directory.");
        state.Require(info.lastWriteTime != 0, L"Remote S3 metadata: object lastWriteTime was not populated.");
        if (! state.failure.empty())
        {
            return false;
        }

        wil::com_ptr<IFileReader> reader;
        const HRESULT readerHr = io->CreateFileReader(objectText.c_str(), reader.put());
        state.Require(SUCCEEDED(readerHr) && reader,
                      std::format(L"Remote S3 metadata: CreateFileReader failed. hr=0x{:08X}", static_cast<unsigned long>(readerHr)));
        if (FAILED(readerHr) || ! reader)
        {
            return false;
        }

        uint64_t sizeBytes   = 0;
        const HRESULT sizeHr = reader->GetSize(&sizeBytes);
        state.Require(SUCCEEDED(sizeHr), std::format(L"Remote S3 metadata: reader->GetSize failed. hr=0x{:08X}", static_cast<unsigned long>(sizeHr)));
        state.Require(sizeBytes == static_cast<uint64_t>(kPayload.size()),
                      std::format(L"Remote S3 metadata: expected size {} but got {}.", kPayload.size(), sizeBytes));
        if (! state.failure.empty())
        {
            return false;
        }

        std::string readBack;
        readBack.resize(kPayload.size());
        size_t totalRead = 0;
        while (totalRead < readBack.size())
        {
            unsigned long chunkRead = 0;
            const unsigned long requestBytes =
                static_cast<unsigned long>((std::min)(readBack.size() - totalRead, static_cast<size_t>((std::numeric_limits<unsigned long>::max)())));
            const HRESULT readHr = reader->Read(readBack.data() + totalRead, requestBytes, &chunkRead);
            state.Require(SUCCEEDED(readHr),
                          std::format(L"Remote S3 metadata: reader->Read failed after {} bytes. hr=0x{:08X}", totalRead, static_cast<unsigned long>(readHr)));
            if (FAILED(readHr))
            {
                return false;
            }
            if (chunkRead == 0)
            {
                break;
            }
            totalRead += chunkRead;
        }

        state.Require(totalRead == kPayload.size(), std::format(L"Remote S3 metadata: expected to read {} bytes but got {}.", kPayload.size(), totalRead));
        state.Require(std::string_view(readBack.data(), totalRead) == kPayload, L"Remote S3 metadata: reader contents did not match the uploaded payload.");
        return state.failure.empty();
    });
};

const auto runRemoteS3DeleteMissing = [&](std::wstring_view caseName) noexcept
{
    if (! SelfTest::CaseFilterMatches(options.caseFilter, caseName))
    {
        return;
    }

    if (options.failFast && suite.failed != 0)
    {
        AppendCaseResult(suite, caseName, SelfTest::SelfTestCaseResult::Status::skipped, L"not executed (fail-fast)");
        return;
    }

    const PhaseCheckResult secretOutcome = CheckRemoteConnectionSecret(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
    if (secretOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, secretOutcome.status, secretOutcome.reason);
        return;
    }

    const PhaseCheckResult sandboxOutcome = CheckRemoteConnectionSandbox(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
    if (sandboxOutcome.status != SelfTest::SelfTestCaseResult::Status::passed)
    {
        AppendCaseResult(suite, caseName, sandboxOutcome.status, sandboxOutcome.reason);
        return;
    }

    SelfTest::RunCase(options,
                      suite,
                      caseName,
                      [&](SelfTest::CaseState& state) noexcept
    {
        const ResolvedRemoteProfile resolvedProfile        = ResolveRemoteConnectionProfile(kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
        const std::wstring& profileName                    = resolvedProfile.profileName;
        const Common::Settings::ConnectionProfile* profile = resolvedProfile.profile;
        state.Require(profile != nullptr, L"Remote S3 delete: profile missing after preconditions passed.");
        if (! profile)
        {
            return false;
        }

        const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
        state.Require(! initialPath.empty() && initialPath[0] == L'/', L"Remote S3 delete: initialPath must be an absolute plugin path.");
        if (initialPath.empty() || initialPath[0] != L'/')
        {
            return false;
        }

        CreatedFileSystemInstance remoteCreated{};
        const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinS3FileSystemId, {}, remoteCreated);
        state.Require(SUCCEEDED(createHr),
                      std::format(L"Remote S3 delete: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! remoteCreated.fileSystem)
        {
            return false;
        }

        wil::com_ptr<IFileSystemIO> io;
        const HRESULT ioHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(io.put()));
        state.Require(SUCCEEDED(ioHr) && io, L"Remote S3 delete: filesystem instance missing IFileSystemIO.");
        if (FAILED(ioHr) || ! io)
        {
            return false;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
        const HRESULT dirHr = remoteCreated.fileSystem->QueryInterface(IID_PPV_ARGS(dirOps.put()));
        state.Require(SUCCEEDED(dirHr) && dirOps, L"Remote S3 delete: filesystem instance missing IFileSystemDirectoryOperations.");
        if (FAILED(dirHr) || ! dirOps)
        {
            return false;
        }

        const std::wstring caseRootPlugin = JoinPluginPathForSelfTest(initialPath, std::format(L"compare-selftest-s3-delete-{}", guid));
        state.Require(! caseRootPlugin.empty(), L"Remote S3 delete: failed to build sandbox path.");
        if (caseRootPlugin.empty())
        {
            return false;
        }

        const std::wstring caseRootText = MakeConnectionPathForSelfTest(profileName, caseRootPlugin);
        state.Require(! caseRootText.empty(), L"Remote S3 delete: failed to build connection path.");
        if (caseRootText.empty())
        {
            return false;
        }

        const std::filesystem::path caseRoot(caseRootText);
        state.Require(EnsureDirectoryExistsFsOps(dirOps, caseRoot), L"Remote S3 delete: failed to prepare sandbox prefix.");
        if (! state.failure.empty())
        {
            return false;
        }

        const std::filesystem::path singlePath    = caseRoot / L"single-delete.txt";
        const std::filesystem::path batchGoodPath = caseRoot / L"batch-existing.txt";
        const std::filesystem::path batchMissPath = caseRoot / L"batch-missing.txt";
        const std::wstring singleText             = ToPluginPathText(singlePath);
        const std::wstring batchGoodText          = ToPluginPathText(batchGoodPath);
        const std::wstring batchMissText          = ToPluginPathText(batchMissPath);
        const auto cleanup                        = wil::scope_exit([&]() noexcept
        {
            const wchar_t* cleanupPaths[] = {singleText.c_str(), batchGoodText.c_str()};
            static_cast<void>(remoteCreated.fileSystem->DeleteItems(
                cleanupPaths, 2, static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, nullptr, nullptr));
        });

        state.Require(WriteFileTextFsIo(io, singlePath, "single-delete"), L"Remote S3 delete: failed to write single-delete.txt.");
        state.Require(WriteFileTextFsIo(io, batchGoodPath, "batch-delete"), L"Remote S3 delete: failed to write batch-existing.txt.");
        if (! state.failure.empty())
        {
            return false;
        }

        const HRESULT deleteHr = remoteCreated.fileSystem->DeleteItem(singleText.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        state.Require(SUCCEEDED(deleteHr), std::format(L"Remote S3 delete: first DeleteItem failed. hr=0x{:08X}", static_cast<unsigned long>(deleteHr)));
        state.Require(! PathExistsFsIo(io, singlePath), L"Remote S3 delete: first DeleteItem did not remove the object.");
        if (! state.failure.empty())
        {
            return false;
        }

        const HRESULT deleteMissingHr = remoteCreated.fileSystem->DeleteItem(singleText.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        const HRESULT notFoundHr      = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        state.Require(deleteMissingHr == notFoundHr,
                      std::format(L"Remote S3 delete: second DeleteItem expected 0x{:08X}, got 0x{:08X}.",
                                  static_cast<unsigned long>(notFoundHr),
                                  static_cast<unsigned long>(deleteMissingHr)));
        if (! state.failure.empty())
        {
            return false;
        }

        RecordingFileSystemCallback deleteCallback(2);
        const wchar_t* deletePaths[] = {batchGoodText.c_str(), batchMissText.c_str()};
        const HRESULT batchDeleteHr  = remoteCreated.fileSystem->DeleteItems(
            deletePaths, 2, static_cast<FileSystemFlags>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR), nullptr, &deleteCallback, nullptr);
        const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        state.Require(batchDeleteHr == partialHr,
                      std::format(L"Remote S3 delete: DeleteItems expected 0x{:08X}, got 0x{:08X}.",
                                  static_cast<unsigned long>(partialHr),
                                  static_cast<unsigned long>(batchDeleteHr)));
        state.Require(! deleteCallback.SawUnexpectedIssue(), L"Remote S3 delete: DeleteItems triggered an unexpected issue callback.");
        state.Require(deleteCallback.CompletedCount() == 2u,
                      std::format(L"Remote S3 delete: DeleteItems expected 2 completion callbacks, got {}.", deleteCallback.CompletedCount()));
        RecordedFileSystemItem deleteItem0{};
        RecordedFileSystemItem deleteItem1{};
        state.Require(deleteCallback.TryGetItem(0, deleteItem0), L"Remote S3 delete: DeleteItems missing callback for item 0.");
        state.Require(deleteCallback.TryGetItem(1, deleteItem1), L"Remote S3 delete: DeleteItems missing callback for item 1.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(SUCCEEDED(deleteItem0.status),
                      std::format(L"Remote S3 delete: DeleteItems item 0 expected success, got 0x{:08X}.", static_cast<unsigned long>(deleteItem0.status)));
        state.Require(deleteItem1.status == notFoundHr,
                      std::format(L"Remote S3 delete: DeleteItems item 1 expected 0x{:08X}, got 0x{:08X}.",
                                  static_cast<unsigned long>(notFoundHr),
                                  static_cast<unsigned long>(deleteItem1.status)));
        state.Require(! PathExistsFsIo(io, batchGoodPath), L"Remote S3 delete: DeleteItems did not remove the existing object.");
        return state.failure.empty();
    });
};

// Optional remote smoke: runs only when Connection Manager profiles + secrets exist.
runRemoteFileCompare(L"remote_file_s3", L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
runRemoteDirectorySizeCallbackContract(
    L"remote_s3_directory_size_callback_contract", L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kBuiltinS3FileSystemId);
runRemoteS3Pagination(L"remote_s3_pagination");
runRemoteFileCompare(L"remote_file_onedrive_personal",
                     L"OneDrive Personal",
                     kSelfTestEnvConnOneDrivePersonal,
                     kSelfTestDefaultConnOneDrivePersonal,
                     kBuiltinOneDrivePersonalFileSystemId);
runRemoteDirectorySizeCallbackContract(L"remote_onedrive_personal_directory_size_callback_contract",
                                       L"OneDrive Personal",
                                       kSelfTestEnvConnOneDrivePersonal,
                                       kSelfTestDefaultConnOneDrivePersonal,
                                       kBuiltinOneDrivePersonalFileSystemId);
runRemoteFileCompare(L"remote_file_onedrive_business",
                     L"OneDrive Business",
                     kSelfTestEnvConnOneDriveBusiness,
                     kSelfTestDefaultConnOneDriveBusiness,
                     kBuiltinOneDriveBusinessFileSystemId);
runRemoteDirectorySizeCallbackContract(L"remote_onedrive_business_directory_size_callback_contract",
                                       L"OneDrive Business",
                                       kSelfTestEnvConnOneDriveBusiness,
                                       kSelfTestDefaultConnOneDriveBusiness,
                                       kBuiltinOneDriveBusinessFileSystemId);
runRemoteFileCompare(L"remote_file_sharepoint", L"SharePoint", kSelfTestEnvConnSharePoint, kSelfTestDefaultConnSharePoint, kBuiltinSharePointFileSystemId);
runRemoteDirectorySizeCallbackContract(L"remote_sharepoint_directory_size_callback_contract",
                                       L"SharePoint",
                                       kSelfTestEnvConnSharePoint,
                                       kSelfTestDefaultConnSharePoint,
                                       kBuiltinSharePointFileSystemId);
runRemoteFileCompare(L"remote_file_ftp", L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kBuiltinFtpFileSystemId);
runRemoteDirectorySizeCallbackContract(
    L"remote_ftp_directory_size_callback_contract", L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kBuiltinFtpFileSystemId);
runRemoteFtpPartialContinue(L"remote_ftp_continue_on_error_partial");
runRemoteS3MetadataSmoke(L"remote_s3_metadata_smoke");
runRemoteS3DeleteMissing(L"remote_s3_delete_missing");
}
