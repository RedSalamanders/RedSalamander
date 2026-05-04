// Commands.SelfTest.Search.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// Search test family: 55 test functions.

struct LocalSearchProbeCallback final : IFileSystemSearchCallback
{
    LocalSearchProbeCallback()                                           = default;
    LocalSearchProbeCallback(const LocalSearchProbeCallback&)            = delete;
    LocalSearchProbeCallback& operator=(const LocalSearchProbeCallback&) = delete;
    LocalSearchProbeCallback(LocalSearchProbeCallback&&)                 = delete;
    LocalSearchProbeCallback& operator=(LocalSearchProbeCallback&&)      = delete;

    std::atomic<uint32_t> matchCallbacks{0};
    std::atomic<uint32_t> progressCallbacks{0};
    std::atomic<uint32_t> completedProgressCallbacks{0};
    std::atomic<uint32_t> invalidProgressPayloads{0};
    std::atomic<uint32_t> warningFlags{FILESYSTEM_SEARCH_WARNING_NONE};
    std::atomic<uint32_t> shouldCancelCallbacks{0};
    std::atomic<uint32_t> concurrentCallbacks{0};
    std::atomic<uint32_t> cancelAfterMatchCallbacks{0};
    std::atomic<HRESULT> lastCompletedStatus{S_OK};
    std::atomic<bool> inCallback{false};

    HRESULT STDMETHODCALLTYPE FileSystemSearchMatch(const ::FileSystemSearchMatch* match, void* /*cookie*/) noexcept override
    {
        const bool alreadyInCallback = inCallback.exchange(true, std::memory_order_acq_rel);
        const auto callbackExit      = wil::scope_exit([&] { inCallback.store(false, std::memory_order_release); });
        if (alreadyInCallback)
        {
            concurrentCallbacks.fetch_add(1u, std::memory_order_acq_rel);
        }

        if (! match || match->sizeBytes != sizeof(::FileSystemSearchMatch))
        {
            return E_INVALIDARG;
        }

        matchCallbacks.fetch_add(1u, std::memory_order_acq_rel);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchProgress(const ::FileSystemSearchProgress* progress, void* /*cookie*/) noexcept override
    {
        const bool alreadyInCallback = inCallback.exchange(true, std::memory_order_acq_rel);
        const auto callbackExit      = wil::scope_exit([&] { inCallback.store(false, std::memory_order_release); });
        if (alreadyInCallback)
        {
            concurrentCallbacks.fetch_add(1u, std::memory_order_acq_rel);
        }

        if (! progress || progress->sizeBytes != sizeof(::FileSystemSearchProgress))
        {
            invalidProgressPayloads.fetch_add(1u, std::memory_order_acq_rel);
            return S_OK;
        }

        progressCallbacks.fetch_add(1u, std::memory_order_acq_rel);
        warningFlags.fetch_or(progress->warningFlags, std::memory_order_acq_rel);
        if (progress->phase == FILESYSTEM_SEARCH_PHASE_COMPLETED)
        {
            completedProgressCallbacks.fetch_add(1u, std::memory_order_acq_rel);
            lastCompletedStatus.store(progress->statusHint, std::memory_order_release);
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchShouldCancel(BOOL* pCancel, void* /*cookie*/) noexcept override
    {
        if (! pCancel)
        {
            return E_POINTER;
        }

        shouldCancelCallbacks.fetch_add(1u, std::memory_order_acq_rel);
        const uint32_t cancelAfter = cancelAfterMatchCallbacks.load(std::memory_order_acquire);
        *pCancel                   = (cancelAfter != 0u && matchCallbacks.load(std::memory_order_acquire) >= cancelAfter) ? TRUE : FALSE;
        return S_OK;
    }
};

struct BlockingDirectoryWatchCallback final : IFileSystemDirectoryWatchCallback
{
    BlockingDirectoryWatchCallback()                                                 = default;
    BlockingDirectoryWatchCallback(const BlockingDirectoryWatchCallback&)            = delete;
    BlockingDirectoryWatchCallback& operator=(const BlockingDirectoryWatchCallback&) = delete;
    BlockingDirectoryWatchCallback(BlockingDirectoryWatchCallback&&)                 = delete;
    BlockingDirectoryWatchCallback& operator=(BlockingDirectoryWatchCallback&&)      = delete;

    std::atomic<uint32_t> callbacks{0};
    std::atomic<uint32_t> invalidPayloads{0};
    std::atomic<bool> blockFirstCallback{false};
    std::atomic<bool> firstCallbackEntered{false};
    std::atomic<bool> firstCallbackExited{false};
    std::atomic<bool> allowFirstCallbackReturn{false};

    HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const ::FileSystemDirectoryChangeNotification* notification, void* /*cookie*/) noexcept override
    {
        if (! notification || notification->sizeBytes != sizeof(::FileSystemDirectoryChangeNotification))
        {
            invalidPayloads.fetch_add(1u, std::memory_order_acq_rel);
            return S_OK;
        }

        const uint32_t callbackIndex = callbacks.fetch_add(1u, std::memory_order_acq_rel) + 1u;
        if (callbackIndex == 1u && blockFirstCallback.load(std::memory_order_acquire))
        {
            firstCallbackEntered.store(true, std::memory_order_release);
            while (! allowFirstCallbackReturn.load(std::memory_order_acquire))
            {
                ::Sleep(1);
            }
            firstCallbackExited.store(true, std::memory_order_release);
        }

        return S_OK;
    }
};

[[nodiscard]] bool WaitForFlag(const std::atomic<bool>& flag, DWORD timeoutMs) noexcept
{
    const ULONGLONG deadline = ::GetTickCount64() + timeoutMs;
    while (::GetTickCount64() < deadline)
    {
        if (flag.load(std::memory_order_acquire))
        {
            return true;
        }

        ::Sleep(10);
    }

    return flag.load(std::memory_order_acquire);
}

[[nodiscard]] bool GetConfiguredLocalFileSystemForSelfTest(CaseState& state, std::string_view configurationJson, CreatedFileSystemInstance& created) noexcept
{
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    if (! configurationJson.empty())
    {
        wil::com_ptr<IInformations> informations;
        state.Require(CreateInformations(created.fileSystem, informations), L"Isolated local file system instance missing IInformations.");
        if (! informations)
        {
            return false;
        }

        const HRESULT configHr = informations->SetConfiguration(configurationJson.data());
        state.Require(SUCCEEDED(configHr),
                      std::format(L"Failed to configure isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(configHr)));
        if (FAILED(configHr))
        {
            return false;
        }
    }

    return state.failure.empty();
}

[[nodiscard]] wil::com_ptr<IFileSystemSearch> QueryLocalFileSystemSearchForSelfTest(CaseState& state, IFileSystem* fs) noexcept
{
    if (! fs)
    {
        state.Require(false, L"Local file-system plugin is not loaded for search selftest.");
        return {};
    }

    wil::com_ptr<IFileSystemSearch> search;
    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemSearch), search.put_void());
    state.Require(SUCCEEDED(hr) && search != nullptr,
                  std::format(L"Local file-system plugin does not expose IFileSystemSearch: 0x{:08X}.", static_cast<unsigned long>(hr)));
    return search;
}

[[nodiscard]] wil::com_ptr<IFileSystemSearch> GetLocalFileSystemSearchForSelfTest(CaseState& state) noexcept
{
    constexpr std::wstring_view kLocalPluginId = L"builtin/file-system";

    wil::com_ptr<IFileSystem> fs = SelfTest::GetFileSystem(kLocalPluginId);
    if (! fs)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(kLocalPluginId, g_settings);
        state.Require(SUCCEEDED(enableHr),
                      std::format(L"Failed to enable local file-system plugin for search selftest: 0x{:08X}.", static_cast<unsigned long>(enableHr)));
        fs = SelfTest::GetFileSystem(kLocalPluginId);
    }

    state.Require(fs != nullptr, L"Local file-system plugin is not loaded for search selftest.");
    if (! fs)
    {
        return {};
    }

    return QueryLocalFileSystemSearchForSelfTest(state, fs.get());
}

[[nodiscard]] bool TestLocalPluginInvalidRegexReportsSingleCompletion(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path caseRoot = suiteRoot / L"work" / (L"search_invalid_regex_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(caseRoot, ec);
    const auto cleanup = wil::scope_exit([&]
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(caseRoot, cleanupEc);
    });

    state.Require(SelfTest::EnsureDirectory(caseRoot), L"Failed to create invalid-regex search root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"candidate.txt", "needle"), L"Failed to create invalid-regex candidate file.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystemSearch> search = GetLocalFileSystemSearchForSelfTest(state);
    if (! search || ! state.failure.empty())
    {
        return false;
    }

    const auto requireInvalidRegexReportedOnce = [&](std::wstring_view label,
                                                     FileSystemSearchNameMode nameMode,
                                                     const wchar_t* namePattern,
                                                     FileSystemSearchContentMode contentMode,
                                                     const wchar_t* contentPattern) noexcept
    {
        LocalSearchProbeCallback callback{};
        FileSystemSearchQuery query{};
        query.sizeBytes              = sizeof(FileSystemSearchQuery);
        query.rootPath               = caseRoot.c_str();
        query.namePattern            = namePattern;
        query.contentPattern         = contentPattern;
        query.flags                  = FILESYSTEM_SEARCH_INCLUDE_FILES;
        query.nameMode               = nameMode;
        query.contentMode            = contentMode;
        query.maxContentBytesPerFile = 0;
        query.maxSnippetCharacters   = 0;

        const HRESULT hr = search->Search(&query, &callback, nullptr);
        state.Require(hr == E_INVALIDARG, std::format(L"{} invalid regex should return E_INVALIDARG, got 0x{:08X}.", label, static_cast<unsigned long>(hr)));
        state.Require(callback.matchCallbacks.load(std::memory_order_acquire) == 0u,
                      std::format(L"{} invalid regex should not emit matches, got {}.", label, callback.matchCallbacks.load(std::memory_order_acquire)));
        state.Require(callback.invalidProgressPayloads.load(std::memory_order_acquire) == 0u,
                      std::format(L"{} invalid regex emitted malformed progress payloads (count={}).",
                                  label,
                                  callback.invalidProgressPayloads.load(std::memory_order_acquire)));
        state.Require(callback.progressCallbacks.load(std::memory_order_acquire) == 1u,
                      std::format(L"{} invalid regex should emit exactly one progress callback, got {}.",
                                  label,
                                  callback.progressCallbacks.load(std::memory_order_acquire)));
        state.Require(callback.completedProgressCallbacks.load(std::memory_order_acquire) == 1u,
                      std::format(L"{} invalid regex should emit exactly one completed progress callback, got {}.",
                                  label,
                                  callback.completedProgressCallbacks.load(std::memory_order_acquire)));
        state.Require((callback.warningFlags.load(std::memory_order_acquire) & FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED) != 0u,
                      std::format(L"{} invalid regex did not include REGEX_REJECTED warning flags (0x{:08X}).",
                                  label,
                                  callback.warningFlags.load(std::memory_order_acquire)));
        state.Require(callback.lastCompletedStatus.load(std::memory_order_acquire) == E_INVALIDARG,
                      std::format(L"{} invalid regex completed status should be E_INVALIDARG, got 0x{:08X}.",
                                  label,
                                  static_cast<unsigned long>(callback.lastCompletedStatus.load(std::memory_order_acquire))));
    };

    requireInvalidRegexReportedOnce(L"name", FILESYSTEM_SEARCH_NAME_REGEX, L"[", FILESYSTEM_SEARCH_CONTENT_DISABLED, nullptr);
    requireInvalidRegexReportedOnce(L"content", FILESYSTEM_SEARCH_NAME_WILDCARD, L"*", FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX, L"[");
    requireInvalidRegexReportedOnce(L"unsafe name", FILESYSTEM_SEARCH_NAME_REGEX, L"(a+)+", FILESYSTEM_SEARCH_CONTENT_DISABLED, nullptr);

    return state.failure.empty();
}

[[nodiscard]] bool TestLocalPluginParallelSearchCancellationAndFanIn(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path caseRoot = suiteRoot / L"work" / (L"search_parallel_cancel_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(caseRoot, ec);
    const auto cleanup = wil::scope_exit([&]
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(caseRoot, cleanupEc);
    });

    state.Require(SelfTest::EnsureDirectory(caseRoot), L"Failed to create parallel-search root.");
    if (! state.failure.empty())
    {
        return false;
    }

    uint32_t expectedMatches = 0u;
    for (uint32_t index = 0u; index < 160u; ++index)
    {
        const std::filesystem::path path = caseRoot / std::format(L"match_flat_{:03}.txt", index);
        state.Require(SelfTest::WriteTextFile(path, "flat"), std::format(L"Failed to create {}.", path.wstring()));
        ++expectedMatches;
    }

    const std::filesystem::path broadRoot = caseRoot / L"broad";
    state.Require(SelfTest::EnsureDirectory(broadRoot), L"Failed to create broad search root.");
    for (uint32_t dirIndex = 0u; dirIndex < 12u; ++dirIndex)
    {
        const std::filesystem::path dir = broadRoot / std::format(L"dir_{:02}", dirIndex);
        state.Require(SelfTest::EnsureDirectory(dir), std::format(L"Failed to create {}.", dir.wstring()));
        for (uint32_t fileIndex = 0u; fileIndex < 4u; ++fileIndex)
        {
            const std::filesystem::path path = dir / std::format(L"match_broad_{:02}_{:02}.txt", dirIndex, fileIndex);
            state.Require(SelfTest::WriteTextFile(path, "broad"), std::format(L"Failed to create {}.", path.wstring()));
            ++expectedMatches;
        }
    }

    std::filesystem::path deepDir = caseRoot / L"deep";
    state.Require(SelfTest::EnsureDirectory(deepDir), L"Failed to create deep search root.");
    for (uint32_t level = 0u; level < 8u; ++level)
    {
        deepDir /= std::format(L"d{:02}", level);
        state.Require(SelfTest::EnsureDirectory(deepDir), std::format(L"Failed to create {}.", deepDir.wstring()));
        const std::filesystem::path path = deepDir / std::format(L"match_deep_{:02}.txt", level);
        state.Require(SelfTest::WriteTextFile(path, "deep"), std::format(L"Failed to create {}.", path.wstring()));
        ++expectedMatches;
    }

    if (! state.failure.empty())
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    constexpr std::string_view kScanFourWalkerConfig = R"json({"searchBackendPreference":"scan","searchMaxDirectoryWalkers":4})json";
    if (! GetConfiguredLocalFileSystemForSelfTest(state, kScanFourWalkerConfig, created))
    {
        return false;
    }

    wil::com_ptr<IFileSystemSearch> search = QueryLocalFileSystemSearchForSelfTest(state, created.fileSystem.get());
    if (! search || ! state.failure.empty())
    {
        return false;
    }

    FileSystemSearchQuery query{};
    query.sizeBytes              = sizeof(FileSystemSearchQuery);
    query.rootPath               = caseRoot.c_str();
    query.namePattern            = L"match_*.txt";
    query.flags                  = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_INCLUDE_FILES | FILESYSTEM_SEARCH_RECURSIVE);
    query.nameMode               = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode            = FILESYSTEM_SEARCH_CONTENT_DISABLED;
    query.maxContentBytesPerFile = 0;
    query.maxSnippetCharacters   = 0;

    LocalSearchProbeCallback fullCallback{};
    const HRESULT fullHr = search->Search(&query, &fullCallback, nullptr);
    state.Require(SUCCEEDED(fullHr), std::format(L"Full parallel scan should succeed, got 0x{:08X}.", static_cast<unsigned long>(fullHr)));
    state.Require(fullCallback.matchCallbacks.load(std::memory_order_acquire) == expectedMatches,
                  std::format(L"Full parallel scan should find {} matches across flat, broad, and deep trees; got {}.",
                              expectedMatches,
                              fullCallback.matchCallbacks.load(std::memory_order_acquire)));
    state.Require(
        fullCallback.concurrentCallbacks.load(std::memory_order_acquire) == 0u,
        std::format(L"Search callbacks must be serialized; saw {} concurrent entries.", fullCallback.concurrentCallbacks.load(std::memory_order_acquire)));
    state.Require(fullCallback.completedProgressCallbacks.load(std::memory_order_acquire) == 1u,
                  std::format(L"Full parallel scan should emit one completed progress callback, got {}.",
                              fullCallback.completedProgressCallbacks.load(std::memory_order_acquire)));
    state.Require(fullCallback.lastCompletedStatus.load(std::memory_order_acquire) == S_OK,
                  std::format(L"Full parallel scan completion status should be S_OK, got 0x{:08X}.",
                              static_cast<unsigned long>(fullCallback.lastCompletedStatus.load(std::memory_order_acquire))));

    LocalSearchProbeCallback cancelCallback{};
    cancelCallback.cancelAfterMatchCallbacks.store(1u, std::memory_order_release);
    const HRESULT cancelHr = search->Search(&query, &cancelCallback, nullptr);
    state.Require(cancelHr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Parallel scan should honor cancellation after first match, got 0x{:08X}.", static_cast<unsigned long>(cancelHr)));
    state.Require(cancelCallback.matchCallbacks.load(std::memory_order_acquire) <= 4u,
                  std::format(L"Parallel scan emitted too many matches after cancellation became visible: {}.",
                              cancelCallback.matchCallbacks.load(std::memory_order_acquire)));
    state.Require(cancelCallback.shouldCancelCallbacks.load(std::memory_order_acquire) > 0u,
                  L"Parallel scan did not poll FileSystemSearchShouldCancel during the match fan-in path.");
    state.Require(cancelCallback.concurrentCallbacks.load(std::memory_order_acquire) == 0u,
                  std::format(L"Cancelled parallel scan callbacks must be serialized; saw {} concurrent entries.",
                              cancelCallback.concurrentCallbacks.load(std::memory_order_acquire)));
    state.Require(cancelCallback.completedProgressCallbacks.load(std::memory_order_acquire) == 1u,
                  std::format(L"Cancelled parallel scan should emit one completed progress callback, got {}.",
                              cancelCallback.completedProgressCallbacks.load(std::memory_order_acquire)));
    state.Require(cancelCallback.lastCompletedStatus.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Cancelled parallel scan completion status should be ERROR_CANCELLED, got 0x{:08X}.",
                              static_cast<unsigned long>(cancelCallback.lastCompletedStatus.load(std::memory_order_acquire))));

    return state.failure.empty();
}

[[nodiscard]] bool TestLocalPluginWatchUnwatchDrainsInflightCallback(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path caseRoot = suiteRoot / L"work" / (L"watch_unwatch_drain_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(caseRoot, ec);
    const auto cleanup = wil::scope_exit([&]
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(caseRoot, cleanupEc);
    });

    state.Require(SelfTest::EnsureDirectory(caseRoot), L"Failed to create watch-drain root.");
    if (! state.failure.empty())
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    if (! GetConfiguredLocalFileSystemForSelfTest(state, {}, created))
    {
        return false;
    }

    wil::com_ptr<IFileSystemDirectoryWatch> watch;
    const HRESULT watchQiHr = created.fileSystem->QueryInterface(__uuidof(IFileSystemDirectoryWatch), watch.put_void());
    state.Require(SUCCEEDED(watchQiHr) && watch != nullptr,
                  std::format(L"Local file-system plugin does not expose IFileSystemDirectoryWatch: 0x{:08X}.", static_cast<unsigned long>(watchQiHr)));
    if (! watch || ! state.failure.empty())
    {
        return false;
    }

    BlockingDirectoryWatchCallback callback{};
    callback.blockFirstCallback.store(true, std::memory_order_release);

    const HRESULT watchHr = watch->WatchDirectory(caseRoot.c_str(), &callback, nullptr);
    state.Require(SUCCEEDED(watchHr), std::format(L"WatchDirectory failed: 0x{:08X}.", static_cast<unsigned long>(watchHr)));
    if (FAILED(watchHr))
    {
        return false;
    }

    state.Require(SelfTest::WriteTextFile(caseRoot / L"first.txt", "first"), L"Failed to create first watched file.");
    state.Require(WaitForFlag(callback.firstCallbackEntered, 5000u), L"Directory watch callback did not start after creating a file.");
    if (! state.failure.empty())
    {
        callback.allowFirstCallbackReturn.store(true, std::memory_order_release);
        static_cast<void>(watch->UnwatchDirectory(caseRoot.c_str()));
        return false;
    }

    std::jthread releaseThread([&callback]
    {
        ::Sleep(150);
        callback.allowFirstCallbackReturn.store(true, std::memory_order_release);
    });

    const ULONGLONG unwatchStartTick = ::GetTickCount64();
    const HRESULT unwatchHr          = watch->UnwatchDirectory(caseRoot.c_str());
    const ULONGLONG unwatchElapsed   = ::GetTickCount64() - unwatchStartTick;
    releaseThread.join();

    state.Require(SUCCEEDED(unwatchHr), std::format(L"UnwatchDirectory failed: 0x{:08X}.", static_cast<unsigned long>(unwatchHr)));
    state.Require(callback.firstCallbackExited.load(std::memory_order_acquire), L"UnwatchDirectory returned before the in-flight callback could exit.");
    state.Require(unwatchElapsed >= 100u, std::format(L"UnwatchDirectory did not appear to drain the blocked callback; elapsed={}ms.", unwatchElapsed));
    state.Require(callback.invalidPayloads.load(std::memory_order_acquire) == 0u,
                  std::format(L"Directory watch emitted malformed payloads (count={}).", callback.invalidPayloads.load(std::memory_order_acquire)));

    const uint32_t callbacksAfterUnwatch = callback.callbacks.load(std::memory_order_acquire);
    state.Require(SelfTest::WriteTextFile(caseRoot / L"after_unwatch.txt", "after"), L"Failed to create post-unwatch file.");
    ::Sleep(250);
    state.Require(callback.callbacks.load(std::memory_order_acquire) == callbacksAfterUnwatch,
                  std::format(L"Directory watch emitted callbacks after UnwatchDirectory returned: before={}, after={}.",
                              callbacksAfterUnwatch,
                              callback.callbacks.load(std::memory_order_acquire)));

    return state.failure.empty();
}

[[nodiscard]] bool TestLocalSearchIndexEnumerateStopsAfterFirstCandidate(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path caseRoot     = suiteRoot / L"work" / (L"search_index_stream_" + NewGuidText());
    const std::filesystem::path dataRoot     = caseRoot / L"data";
    const std::filesystem::path nestedRoot   = dataRoot / L"nested";
    const std::filesystem::path snapshotRoot = caseRoot / L"snapshots";

    std::error_code ec;
    std::filesystem::remove_all(caseRoot, ec);
    const auto cleanup = wil::scope_exit([&]
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(caseRoot, cleanupEc);
    });

    state.Require(SelfTest::EnsureDirectory(nestedRoot), L"Failed to create indexed-search nested folder.");
    state.Require(SelfTest::EnsureDirectory(snapshotRoot), L"Failed to create indexed-search snapshot folder.");
    for (int index = 0; index < 12; ++index)
    {
        const std::filesystem::path folder = index < 6 ? dataRoot : nestedRoot;
        const std::filesystem::path file   = folder / std::format(L"match_{:02}.txt", index);
        state.Require(SelfTest::WriteTextFile(file, "streamed candidate"), std::format(L"Failed to create {}.", file.filename().native()));
    }
    state.Require(SelfTest::WriteTextFile(dataRoot / L"ignore.bin", "ignored"), L"Failed to create non-matching indexed-search file.");
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::Repository repository({.snapshotRootDirectory = snapshotRoot.native()});
    LocalSearchIndexCore::SupportInfo support{};
    HRESULT hr = repository.ProbePath(dataRoot.native(), support);
    state.Require(SUCCEEDED(hr), std::format(L"ProbePath failed for indexed-search stream test: 0x{:08X}.", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(support.indexable, L"Indexed-search stream test requires an indexable local volume.");
    if (! support.indexable)
    {
        return false;
    }

    LocalSearchIndexCore::QueryPlan plan{};
    plan.rootPath           = dataRoot.native();
    plan.namePattern        = L"match_*.txt";
    plan.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    plan.recursive          = true;
    plan.includeFiles       = true;
    plan.includeDirectories = false;

    EnumerateStopAfterFirstState callbackState{};
    LocalSearchIndexCore::QueryStats stats{};
    hr = repository.Enumerate(plan, nullptr, nullptr, &StopAfterFirstIndexedCandidate, &callbackState, &stats);
    state.Require(hr == S_OK, std::format(L"Enumerate should stop cleanly after the first candidate, got 0x{:08X}.", static_cast<unsigned long>(hr)));
    state.Require(callbackState.seenCandidates == 1u,
                  std::format(L"Expected Enumerate callback to see exactly one candidate, got {}.", callbackState.seenCandidates));
    state.Require(stats.candidateCount == 1u, std::format(L"Expected indexed candidate count to stop at 1, got {}.", stats.candidateCount));
    state.Require(stats.fileCount >= 12u, std::format(L"Expected indexed file count to include the prepared files, got {}.", stats.fileCount));
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogSearchOps(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    ShortcutManager shortcuts;
    shortcuts.Load(ShortcutDefaults::CreateDefaultShortcuts());
    const auto findChordOpt = shortcuts.TryGetShortcutForCommand(L"cmd/pane/find");
    state.Require(findChordOpt.has_value(), L"Find default shortcut missing.");
    if (findChordOpt.has_value())
    {
        state.Require(findChordOpt->vk == VK_F7, L"Find default shortcut expected F7.");
        state.Require(findChordOpt->modifiers == ShortcutManager::kModAlt, L"Find default shortcut expected Alt+F7.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePaths                                = wil::scope_exit([&]
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }

        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog";
    const std::filesystem::path sub  = root / L"sub";
    const std::filesystem::path bulk = root / L"bulk";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create find dialog subdirectory.");
    state.Require(SelfTest::EnsureDirectory(bulk), L"Failed to create find dialog bulk directory.");
    state.Require(SelfTest::WriteTextFile(root / L"a.jsonl", "{\"message\":\"needle alpha\"}"), L"Failed to create a.jsonl.");
    state.Require(SelfTest::WriteTextFile(root / L"b.txt", "plain text"), L"Failed to create b.txt.");
    state.Require(SelfTest::WriteTextFile(sub / L"c.jsonl", "{\"message\":\"other\"}"), L"Failed to create c.jsonl.");
    state.Require(SelfTest::WriteTextFile(sub / L"d.txt", "content needle beta"), L"Failed to create d.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for find dialog test.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for find dialog test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for find dialog test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.jsonl", L"b.txt", L"sub", L"bulk"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for find dialog test.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (const HWND existing = GetFindFilesWindowHandle(); existing && IsWindow(existing))
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Find window did not close before find dialog test.");
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open from cmd/pane/find.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, true), L"Failed to configure Find window options.");

    const auto waitMs = [](std::chrono::milliseconds value) noexcept -> uint32_t { return static_cast<uint32_t>(value.count()); };

    auto requireIdle = [&](std::wstring_view label) noexcept
    { state.Require(DebugWaitForFindFilesWindowIdle(waitMs(SelfTest::Scale(10000ms))), std::format(L"Find window did not become idle after {}.", label)); };

    struct RenderSnapshot
    {
        uint64_t dxRenderCount  = 0u;
        uint64_t gridPaintCount = 0u;
    };

    struct RenderSample
    {
        std::wstring label;
        std::wstring action;
        uint64_t waitUs         = 0u;
        uint64_t dxRenderDelta  = 0u;
        uint64_t gridPaintDelta = 0u;
    };

    std::vector<RenderSample> renderSamples;

    auto captureRenderCounts = [&]() noexcept -> RenderSnapshot
    {
        FindFilesDebugSnapshot snapshot{};
        if (! DebugGetFindFilesWindowSnapshot(snapshot))
        {
            return {};
        }
        return {.dxRenderCount = snapshot.dxRenderCount, .gridPaintCount = snapshot.resultGridPaintCount};
    };

    auto requireRenderedAfter = [&](RenderSnapshot before, std::wstring_view label) noexcept
    {
        std::wstring renderAction = L"idle";
        FindFilesDebugSnapshot current{};
        if (DebugGetFindFilesWindowSnapshot(current))
        {
            static_cast<void>(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid));
            if (current.resultListHasVerticalScrollbar && current.visibleResultRowCount < current.resultCount)
            {
                if (DebugScrollFindFilesWindowResultsByWheelDetents(-1))
                {
                    renderAction = L"scroll";
                }
            }
            else if (! current.fullPaths.empty() && DebugSelectFindFilesWindowResult(current.fullPaths.front()))
            {
                renderAction = L"select";
            }
        }

        const auto waitStartedAt = std::chrono::steady_clock::now();
        FindFilesDebugSnapshot rendered{};
        static_cast<void>(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& snapshot) noexcept {
            return snapshot.dxRenderCount > before.dxRenderCount || snapshot.resultGridPaintCount > before.gridPaintCount;
        }, SelfTest::Scale(3000ms), &rendered));
        const uint64_t waitUs =
            static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - waitStartedAt).count());
        renderSamples.push_back(RenderSample{
            .label          = std::wstring(label),
            .action         = std::move(renderAction),
            .waitUs         = waitUs,
            .dxRenderDelta  = rendered.dxRenderCount >= before.dxRenderCount ? rendered.dxRenderCount - before.dxRenderCount : 0u,
            .gridPaintDelta = rendered.resultGridPaintCount >= before.gridPaintCount ? rendered.resultGridPaintCount - before.gridPaintCount : 0u,
        });
    };

    auto requireSnapshotCount = [&](size_t expectedCount, std::wstring_view label, std::initializer_list<std::filesystem::path> expectedPaths) noexcept
    {
        FindFilesDebugSnapshot snapshot{};
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture Find snapshot after {}.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(snapshot.resultCount == expectedCount,
                      std::format(L"Unexpected Find result count after {}. Expected {}, got {}.", label, expectedCount, snapshot.resultCount));
        for (const auto& path : expectedPaths)
        {
            const std::wstring expected = path.native();
            const bool found            = std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), expected) != snapshot.fullPaths.end();
            state.Require(found, std::format(L"Expected Find results after {} to contain '{}'.", label, expected));
        }
    };

    auto requireSortedResultNames = [&](size_t expectedCount, std::wstring_view label) noexcept
    {
        FindFilesDebugSnapshot snapshot{};
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture sorted Find snapshot after {}.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(snapshot.resultCount == expectedCount,
                      std::format(L"Unexpected sorted Find result count after {}. Expected {}, got {}.", label, expectedCount, snapshot.resultCount));
        if (! state.failure.empty())
        {
            return;
        }

        for (size_t index = 1; index < snapshot.fullPaths.size(); ++index)
        {
            const std::wstring previousName = std::filesystem::path(snapshot.fullPaths[index - 1]).filename().native();
            const std::wstring currentName  = std::filesystem::path(snapshot.fullPaths[index]).filename().native();
            state.Require(! OrdinalString::LessNoCase(currentName, previousName),
                          std::format(L"Find results should remain sorted by name after {}. '{}' appeared after '{}'.", label, currentName, previousName));
            if (! state.failure.empty())
            {
                return;
            }
        }
    };

    const RenderSnapshot findRenderBefore = captureRenderCounts();
    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for wildcard search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find operation.");
    requireIdle(L"Find");
    requireRenderedAfter(findRenderBefore, L"Find");
    requireSnapshotCount(2u, L"Find", {root / L"a.jsonl", sub / L"c.jsonl"});

    const RenderSnapshot appendRenderBefore = captureRenderCounts();
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for append search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Append), L"Failed to start Append operation.");
    requireIdle(L"Append");
    requireRenderedAfter(appendRenderBefore, L"Append");
    requireSnapshotCount(4u, L"Append", {root / L"a.jsonl", root / L"b.txt", sub / L"c.jsonl", sub / L"d.txt"});

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"a*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for intersect search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Intersect), L"Failed to start Intersect operation.");
    requireIdle(L"Intersect");
    requireSnapshotCount(1u, L"Intersect", {root / L"a.jsonl"});

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Subtract), L"Failed to start Subtract operation.");
    requireIdle(L"Subtract");
    requireSnapshotCount(0u, L"Subtract", {});

    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"", L"needle", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure Find window for content search.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, true), L"Failed to restore Find options for content search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start content Find operation.");
    requireIdle(L"content Find");
    requireSnapshotCount(2u, L"content Find", {root / L"a.jsonl", sub / L"d.txt"});

    std::string largeBody(2u * 1024u * 1024u, 'x');
    largeBody.back() = '\n';
    for (int i = 0; i < 48; ++i)
    {
        const std::filesystem::path bulkFile = bulk / std::format(L"bulk_{:03}.txt", i);
        state.Require(SelfTest::WriteTextFile(bulkFile, largeBody), std::format(L"Failed to create {}.", bulkFile.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const RenderSnapshot sortedRenderBefore = captureRenderCounts();
    state.Require(DebugSetFindFilesWindowResultSort(0u, false), L"Failed to enable ascending Name sort for Find results.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for sorted search.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for sorted search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start sorted Find operation.");
    requireIdle(L"sorted Find");
    requireRenderedAfter(sortedRenderBefore, L"sorted Find");
    requireSortedResultNames(50u, L"sorted Find");

    state.Require(
        DebugConfigureFindFilesWindow(
            root.native(), L"*.txt", L"ZZZ_NOT_PRESENT_123456789", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::TextRegex),
        L"Failed to configure Find window for cancellation search.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for cancellation search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start cancellation search.");
    state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel Find search.");
    requireIdle(L"cancelled Find");

    FindFilesDebugSnapshot cancelled{};
    state.Require(DebugGetFindFilesWindowSnapshot(cancelled), L"Failed to capture cancelled Find snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(cancelled.searchActive == false, L"Find window remained active after cancellation.");
    state.Require(cancelled.lastStatusHint == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Expected cancelled Find HRESULT 0x{:08X}, got 0x{:08X}.",
                              static_cast<unsigned long>(HRESULT_FROM_WIN32(ERROR_CANCELLED)),
                              static_cast<unsigned long>(cancelled.lastStatusHint)));

    for (const auto& sample : renderSamples)
    {
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                                   std::format(L"find_dialog_search_ops render sample label='{}' action='{}' waitUs={} dxRenderDelta={} gridPaintDelta={}",
                                               sample.label,
                                               sample.action,
                                               sample.waitUs,
                                               sample.dxRenderDelta,
                                               sample.gridPaintDelta));
    }
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, std::format(L"find_dialog_search_ops render sampleCount={}", renderSamples.size()));

    PostMessageW(findWindow, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)), L"Find window did not close after operations test.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogTypedLocalRootOverridesStaleContext(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            static_cast<void>(SendMessageW(findWindow, WM_CLOSE, 0, 0));
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(5000ms)));
        }
    };
    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_local_override";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create override test subdirectory.");
    state.Require(SelfTest::WriteTextFile(root / L"first.jsonl", "{\"message\":\"alpha\"}"), L"Failed to create first.jsonl.");
    state.Require(SelfTest::WriteTextFile(sub / L"second.jsonl", "{\"message\":\"beta\"}"), L"Failed to create second.jsonl.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesPaneContext staleContext{};
    staleContext.pluginId        = L"builtin/file-system-gdrive";
    staleContext.pluginShortId   = L"gdrive";
    staleContext.instanceContext = L"@stale";
    staleContext.rootPluginPath  = std::filesystem::path(L"gdrive://@stale");

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-local-override");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(staleContext)), L"Failed to open Find window with stale context.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for stale-context override test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for stale-context override test.");
    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for stale-context override test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for stale-context override test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for stale-context override test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture stale-context override snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.lastStatusHint == S_OK,
                  std::format(L"Expected stale-context override search to succeed, got 0x{:08X}.", static_cast<unsigned long>(snapshot.lastStatusHint)));
    state.Require(snapshot.resultCount == 2u, std::format(L"Expected 2 results for stale-context override search, got {}.", snapshot.resultCount));

    for (const std::filesystem::path& expectedPath : {root / L"first.jsonl", sub / L"second.jsonl"})
    {
        const std::wstring expected = expectedPath.native();
        const bool found            = std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), expected) != snapshot.fullPaths.end();
        state.Require(found, std::format(L"Expected stale-context override results to contain '{}'.", expected));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogLargeLocalSearchUsesIncrementalUpdates(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            static_cast<void>(SendMessageW(findWindow, WM_CLOSE, 0, 0));
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(5000ms)));
        }
    };
    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_large_incremental";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create incremental-update test subdirectory.");
    constexpr size_t kRootMatchCount = 80u;
    constexpr size_t kSubMatchCount  = 80u;
    for (size_t index = 0; index < kRootMatchCount; ++index)
    {
        const std::filesystem::path filePath = root / std::format(L"root_{:03}.jsonl", index);
        state.Require(SelfTest::WriteTextFile(filePath, "{\"message\":\"root\"}\n"), std::format(L"Failed to create '{}'.", filePath.filename().native()));
    }
    for (size_t index = 0; index < kSubMatchCount; ++index)
    {
        const std::filesystem::path filePath = sub / std::format(L"sub_{:03}.jsonl", index);
        state.Require(SelfTest::WriteTextFile(filePath, "{\"message\":\"sub\"}\n"), std::format(L"Failed to create '{}'.", filePath.filename().native()));
    }
    state.Require(SelfTest::WriteTextFile(root / L"ignore.txt", L"ignore"), L"Failed to create ignore.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesPaneContext staleContext{};
    staleContext.pluginId        = L"builtin/file-system-onedrive-business";
    staleContext.pluginShortId   = L"onedrive";
    staleContext.instanceContext = L"@stale";
    staleContext.rootPluginPath  = std::filesystem::path(L"onedrive://@stale");

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-large-incremental");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(staleContext)), L"Failed to open Find window for incremental-update test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for incremental-update test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window in incremental-update test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for incremental-update test.");
    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for incremental-update test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for incremental-update test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())),
                  L"Find window did not become idle for incremental-update test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture incremental-update snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kExpectedMatches = kRootMatchCount + kSubMatchCount;
    state.Require(snapshot.lastStatusHint == S_OK,
                  std::format(L"Expected incremental-update search to succeed, got 0x{:08X}.", static_cast<unsigned long>(snapshot.lastStatusHint)));
    state.Require(snapshot.resultCount == kExpectedMatches,
                  std::format(L"Expected {} incremental-update results, got {}.", kExpectedMatches, snapshot.resultCount));
    state.Require(snapshot.resultListFullRebuildCount == 0u,
                  std::format(L"Append-only Find search should not rebuild the whole results list, got {} rebuild(s).", snapshot.resultListFullRebuildCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRunningStatusShowsPhaseAndPath(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_running_status";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create running-status test subdirectory.");
    std::string largeBody(2u * 1024u * 1024u, 'x');
    largeBody.back() = '\n';
    for (int index = 0; index < 32; ++index)
    {
        const std::filesystem::path filePath =
            (index % 2 == 0) ? (root / std::format(L"root_{:03}.json", index)) : (sub / std::format(L"sub_{:03}.json", index));
        state.Require(SelfTest::WriteTextFile(filePath, largeBody), std::format(L"Failed to create '{}'.", filePath.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesPaneContext staleContext{};
    staleContext.pluginId        = L"builtin/file-system-onedrive-business";
    staleContext.pluginShortId   = L"onedrive";
    staleContext.instanceContext = L"@stale";
    staleContext.rootPluginPath  = std::filesystem::path(L"onedrive://@stale");

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-running-status");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(staleContext)), L"Failed to open Find window for running-status test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for running-status test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window in running-status test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to configure Find options for running-status test.");
    state.Require(DebugConfigureFindFilesWindow(root.native(),
                                                L"*.json",
                                                L"ZZZ_NOT_PRESENT_123456789",
                                                Common::Settings::SearchNameMode::Wildcard,
                                                Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure Find window for running-status test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for running-status test.");

    FindFilesDebugSnapshot active{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.searchActive && snapshot.statusText.contains(L"Searching with ") && snapshot.statusText.contains(L"last progress ") &&
               snapshot.statusText.find(root.native()) != std::wstring::npos &&
               (snapshot.statusText.contains(L"initializing") || snapshot.statusText.contains(L"enumerating") ||
                snapshot.statusText.contains(L"content scan") || snapshot.statusText.contains(L"index lookup"));
    },
                      SelfTest::Scale(10000ms),
                      &active),
                  std::format(L"Running Find status did not expose phase/path details. Last status='{}'.", active.statusText));
    if (! state.failure.empty())
    {
        static_cast<void>(DebugCancelFindFilesWindowSearch());
        static_cast<void>(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())));
        return false;
    }

    state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel Find search for running-status test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle after running-status cancellation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogServiceStatusShowsBackendDiagnostics(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_service_status";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create service-status test subdirectory.");
    std::string largeBody(2u * 1024u * 1024u, 'x');
    largeBody.back() = '\n';
    for (int index = 0; index < 48; ++index)
    {
        const std::filesystem::path filePath =
            (index % 2 == 0) ? (root / std::format(L"root_{:03}.json", index)) : (sub / std::format(L"sub_{:03}.json", index));
        state.Require(SelfTest::WriteTextFile(filePath, largeBody), std::format(L"Failed to create '{}'.", filePath.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0,
                  L"Failed to override the search service pipe for Find backend-status test.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    const std::filesystem::path sqlitePath = root / L"find-service-status.sqlite3";
    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring serviceArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 32u, serviceArgs, true, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForBrokerReady = [&](SearchServiceBroker::ServiceStatus& outStatus) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(10000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outStatus = {};
            if (SUCCEEDED(SearchServiceBroker::GetStatus(outStatus)) && outStatus.pipeName == pipeName)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outStatus = {};
        return SUCCEEDED(SearchServiceBroker::GetStatus(outStatus)) && outStatus.pipeName == pipeName;
    };

    SearchServiceBroker::ServiceStatus initialStatus{};
    state.Require(waitForBrokerReady(initialStatus), std::format(L"Search service broker did not become ready for the isolated pipe '{}'.", pipeName));
    if (! state.failure.empty())
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    if (! info)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"service\"}");
    state.Require(SUCCEEDED(setHr),
                  std::format(L"Failed to configure service backend for Find backend-status test. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    FindFilesPaneContext context{};
    context.fileSystem     = created.fileSystem;
    context.pluginId       = std::wstring(kBuiltinLocalFileSystemId);
    context.pluginShortId  = L"file";
    context.rootPluginPath = root;

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-service-status");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for service backend-status test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for service backend-status test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window in service backend-status test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to configure Find options for service backend-status test.");
    state.Require(DebugConfigureFindFilesWindow(root.native(),
                                                L"*.json",
                                                L"ZZZ_NOT_PRESENT_123456789",
                                                Common::Settings::SearchNameMode::Wildcard,
                                                Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure Find window for service backend-status test.");

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for service backend-status test.");

    const auto leftInitialPlaceholder = [](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.backend != FILESYSTEM_SEARCH_BACKEND_UNKNOWN || snapshot.phase != FILESYSTEM_SEARCH_PHASE_INITIALIZING ||
               snapshot.warningFlags != FILESYSTEM_SEARCH_WARNING_NONE || snapshot.statusText != L"Ready" || snapshot.hasServiceStatus ||
               ! snapshot.backendStatusText.empty();
    };

    FindFilesDebugSnapshot firstProgress{};
    state.Require(WaitForFindSnapshot(leftInitialPlaceholder, SelfTest::Scale(5000ms), &firstProgress),
                  std::format(L"Service-backed Find search did not transition out of the initial placeholder status. status='{}' backend='{}' active={} "
                              L"phase={} warnings=0x{:08X}.",
                              firstProgress.statusText,
                              firstProgress.backendStatusText,
                              firstProgress.searchActive ? 1 : 0,
                              static_cast<unsigned>(firstProgress.phase),
                              static_cast<unsigned long>(firstProgress.warningFlags)));
    if (! state.failure.empty())
    {
        static_cast<void>(DebugCancelFindFilesWindowSearch());
        static_cast<void>(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())));
        return false;
    }
    const auto containsAny = [](std::wstring_view text, std::initializer_list<std::wstring_view> tokens) noexcept
    {
        for (const std::wstring_view token : tokens)
        {
            if (! token.empty() && text.find(token) != std::wstring_view::npos)
            {
                return true;
            }
        }
        return false;
    };

    const auto waitForServiceDiagnostics = [&](FindFilesDebugSnapshot& outActive) noexcept
    {
        return WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& snapshot) noexcept
        {
            return snapshot.backend == FILESYSTEM_SEARCH_BACKEND_SERVICE && snapshot.hasServiceStatus &&
                   (snapshot.warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u && snapshot.backendStatusText.contains(L"db ") &&
                   snapshot.backendStatusText.contains(L" | search ") && snapshot.backendStatusText.contains(L"live-scan") &&
                   containsAny(snapshot.backendStatusText, {L"store-missing", L"store-invalid", L"warmup-running", L"store-stale"});
        },
            SelfTest::Scale(15000ms),
            &outActive);
    };

    FindFilesDebugSnapshot active{};
    bool reachedDiagnostics = waitForServiceDiagnostics(active);
    if (! reachedDiagnostics)
    {
        state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel the initial Find search after missing service backend diagnostics.");
        state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())),
                      L"Find window did not become idle after cancelling the initial service backend-status attempt.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to restart Find search for the service backend-status retry.");
        if (! state.failure.empty())
        {
            return false;
        }

        firstProgress = {};
        state.Require(WaitForFindSnapshot(leftInitialPlaceholder, SelfTest::Scale(5000ms), &firstProgress),
                      std::format(L"Service-backed Find retry did not transition out of the initial placeholder status. status='{}' backend='{}' active={} "
                                  L"phase={} warnings=0x{:08X}.",
                                  firstProgress.statusText,
                                  firstProgress.backendStatusText,
                                  firstProgress.searchActive ? 1 : 0,
                                  static_cast<unsigned>(firstProgress.phase),
                                  static_cast<unsigned long>(firstProgress.warningFlags)));
        if (! state.failure.empty())
        {
            static_cast<void>(DebugCancelFindFilesWindowSearch());
            static_cast<void>(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())));
            return false;
        }

        reachedDiagnostics = waitForServiceDiagnostics(active);
    }

    state.Require(reachedDiagnostics,
                  std::format(L"Service-backed Find status did not expose database diagnostics. Last status='{}' backend='{}'.",
                              active.statusText,
                              active.backendStatusText));
    if (! state.failure.empty())
    {
        static_cast<void>(DebugCancelFindFilesWindowSearch());
        static_cast<void>(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())));
        return false;
    }

    state.Require((active.warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) == 0u,
                  std::format(L"Healthy service degraded path should not report service unavailable. warnings=0x{:08X}.",
                              static_cast<unsigned long>(active.warningFlags)));
    state.Require(active.statusText.find(active.backendStatusText) != std::wstring::npos,
                  std::format(L"Running Find status should include the backend diagnostics text. status='{}' backend='{}'.",
                              active.statusText,
                              active.backendStatusText));

    state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel Find search for service backend-status test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())),
                  L"Find window did not become idle after service backend-status cancellation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogServiceUnavailableWarningIsDistinct(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_service_unavailable";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create service-unavailable test root.");
    state.Require(SelfTest::WriteTextFile(root / L"match.jsonl", "{\"message\":\"seed\"}"), L"Failed to create match.jsonl for service-unavailable test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring unavailablePipe      = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, unavailablePipe.c_str()) != 0,
                  L"Failed to override the search service pipe for the unavailable-service Find test.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    if (! info)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"service\"}");
    state.Require(SUCCEEDED(setHr),
                  std::format(L"Failed to configure service backend for unavailable-service Find test. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    FindFilesPaneContext context{};
    context.fileSystem     = created.fileSystem;
    context.pluginId       = std::wstring(kBuiltinLocalFileSystemId);
    context.pluginShortId  = L"file";
    context.rootPluginPath = root;

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-service-unavailable");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for unavailable-service warning test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for unavailable-service warning test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window in unavailable-service warning test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to configure Find options for unavailable-service warning test.");
    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for unavailable-service warning test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for unavailable-service warning test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())),
                  L"Find window did not become idle for unavailable-service warning test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture unavailable-service warning snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring serviceUnavailableText = LoadStringResource(nullptr, IDS_FIND_WARNING_SERVICE_UNAVAILABLE);
    state.Require(
        snapshot.lastStatusHint == S_OK,
        std::format(L"Unavailable-service Find search should still succeed via fallback. hr=0x{:08X}.", static_cast<unsigned long>(snapshot.lastStatusHint)));
    state.Require(snapshot.resultCount == 1u, std::format(L"Unavailable-service Find search should still return one match, got {}.", snapshot.resultCount));
    state.Require((snapshot.warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) != 0u,
                  std::format(L"Unavailable-service Find search should report the service-unavailable warning. warnings=0x{:08X}.",
                              static_cast<unsigned long>(snapshot.warningFlags)));
    state.Require(snapshot.statusText.contains(serviceUnavailableText),
                  std::format(L"Unavailable-service Find status should surface '{}' in the UI, got '{}'.", serviceUnavailableText, snapshot.statusText));
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogFailureShowsReadableStatus(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path paneRoot    = suiteRoot / L"work" / L"find_dialog_failure_status";
    const std::filesystem::path missingRoot = paneRoot / L"missing";

    std::error_code ec;
    std::filesystem::remove_all(paneRoot, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(paneRoot), L"Failed to create failure-status pane root.");
    state.Require(SelfTest::WriteTextFile(paneRoot / L"seed.jsonl", "{\"message\":\"seed\"}"), L"Failed to create seed.jsonl.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for failure-status test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, paneRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, paneRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for failure-status test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"seed.jsonl"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for failure-status test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in failure-status test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for failure-status test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window in failure-status test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to configure Find options for failure-status test.");
    state.Require(DebugConfigureFindFilesWindow(
                      missingRoot.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for failure-status test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for failure-status test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for failure-status test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture failure-status snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FAILED(snapshot.lastStatusHint), L"Expected failure-status test search to fail.");
    const std::wstring expectedHr = std::format(L"0x{:08X}", static_cast<unsigned long>(snapshot.lastStatusHint));
    state.Require(snapshot.statusText.starts_with(L"Search failed with "),
                  std::format(L"Expected readable failure status with backend context, got '{}'.", snapshot.statusText));
    state.Require(snapshot.statusText.find(L": ") != std::wstring::npos,
                  std::format(L"Readable failure status should include a backend/value separator, got '{}'.", snapshot.statusText));
    state.Require(snapshot.statusText.contains(expectedHr),
                  std::format(L"Readable failure status should preserve the HRESULT code '{}', got '{}'.", expectedHr, snapshot.statusText));
    state.Require(snapshot.statusText != std::format(L"Search failed ({})", expectedHr), L"Failure status regressed to the old raw-HRESULT format.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogUsesDxUiHostWithNoVisibleChildControls(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const auto waitForFindWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms)); };

    const auto validateFindHostSurface = [&](std::wstring_view context) noexcept
    {
        FocusFolderViewPane(FolderWindow::Pane::Left);
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/find during {}.", context));

        const HWND findWindow = waitForFindWindow();
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, std::format(L"Find window did not open during {}.", context));
        if (! findWindow || IsWindow(findWindow) == FALSE)
        {
            return false;
        }

        FindFilesDebugSnapshot snapshot{};
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture Find snapshot during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(snapshot.usesDxUiHost, std::format(L"Find window is not attached to the shared DxUi host during {}.", context));
        state.Require(snapshot.hasStatusStrip, std::format(L"Find window should expose a retained DxUi StatusStrip during {}.", context));
        state.Require(snapshot.statusStripVisible, std::format(L"Find window StatusStrip should remain visible during {}.", context));
        state.Require(
            snapshot.statusStripHeightDip >= 20.0f,
            std::format(L"Find window StatusStrip should keep a usable height during {}; observed {:.2f} DIP.", context, snapshot.statusStripHeightDip));
        state.Require(WindowExposesUiaProvider(findWindow), std::format(L"Find window should answer WM_GETOBJECT during {}.", context));
        state.Require(snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Find window should not expose visible child-control fallback during {}; got {} visible child window(s).",
                                  context,
                                  snapshot.visibleChildWindowCount));
        state.Require(snapshot.dxResizeFailureCount == 0u,
                      std::format(L"Find window should not report DX resize failures during {}; saw {}.", context, snapshot.dxResizeFailureCount));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(findWindow);
        state.Require(uiaPatternStats.has_value(), std::format(L"Failed to collect live UI Automation pattern statistics for Find window during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Find window should expose visible UI Automation descendants during {}.", context));
            state.Require(uiaPatternStats->editControlCount + uiaPatternStats->comboBoxControlCount > 0u,
                          std::format(L"Find window should expose visible UI Automation edit or combo descendants during {}.", context));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation ValuePattern support during {}.", context));
            state.Require(uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Find window should expose a visible UI Automation command button during {}.", context));
            state.Require(uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation InvokePattern support during {}.", context));
        }

        auto valueState = CollectVisibleDescendantValuePatternState(findWindow, UIA_EditControlTypeId);
        if (! valueState.has_value())
        {
            valueState = CollectVisibleDescendantValuePatternState(findWindow, UIA_ComboBoxControlTypeId);
        }
        state.Require(valueState.has_value(), std::format(L"Find window should expose a visible editable DX field during {}.", context));
        if (valueState.has_value())
        {
            state.Require(! valueState->isReadOnly, std::format(L"Find window editable DX field should not be read-only during {}.", context));
            state.Require(! valueState->name.empty(), std::format(L"Find window editable DX field should expose a stable accessible name during {}.", context));
        }

        const auto buttonState = CollectVisibleDescendantNamedElementState(findWindow, UIA_ButtonControlTypeId);
        state.Require(buttonState.has_value(), std::format(L"Find window should expose a visible DX command button during {}.", context));
        if (buttonState.has_value())
        {
            state.Require(! buttonState->name.empty(),
                          std::format(L"Find window visible DX command button should expose a stable accessible name during {}.", context));
        }

        PostMessageW(findWindow, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)), std::format(L"Find window did not close cleanly during {}.", context));
        return state.failure.empty();
    };

    if (! validateFindHostSurface(L"the initial Find DX host baseline probe"))
    {
        return false;
    }

    if (! validateFindHostSurface(L"the reopened Find DX host baseline probe"))
    {
        return false;
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogExposesLiveUiaSelectionAndInputs(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_live_uia";
    const std::filesystem::path subA = root / L"sub_a";
    const std::filesystem::path subB = root / L"sub_b";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(subA), L"Failed to create first Find UIA test subdirectory.");
    state.Require(SelfTest::EnsureDirectory(subB), L"Failed to create second Find UIA test subdirectory.");
    state.Require(SelfTest::WriteTextFile(subA / L"inside.txt", "payload-a"), L"Failed to create first Find UIA test file.");
    state.Require(SelfTest::WriteTextFile(subB / L"inside.txt", "payload-b"), L"Failed to create second Find UIA test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Find UIA test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Find UIA test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub_a", L"sub_b"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for Find UIA test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring subAPath = subA.native();
    const std::wstring subBPath = subB.native();

    const auto runLiveUiaCycle = [&](std::wstring_view phaseLabel) noexcept
    {
        FocusFolderViewPane(FolderWindow::Pane::Left);
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/find during {}.", phaseLabel));

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, std::format(L"Find window did not open during {}.", phaseLabel));
        if (! findWindow || IsWindow(findWindow) == FALSE)
        {
            return false;
        }

        state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), std::format(L"Failed to configure Find options during {}.", phaseLabel));
        state.Require(DebugConfigureFindFilesWindow(
                          root.native(), L"sub*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                      std::format(L"Failed to configure Find window during {}.", phaseLabel));
        state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), std::format(L"Failed to start Find search during {}.", phaseLabel));
        state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                      std::format(L"Find window did not become idle during {}.", phaseLabel));

        FindFilesDebugSnapshot snapshot{};
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture Find snapshot during {}.", phaseLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(snapshot.usesDxUiHost, std::format(L"Find window should stay attached to the shared DxUi host during {}.", phaseLabel));
        state.Require(
            snapshot.visibleChildWindowCount <= 1u,
            std::format(L"Find window should not expose visible child-control fallback during {}; saw {}.", phaseLabel, snapshot.visibleChildWindowCount));
        state.Require(snapshot.resultCount == 2u,
                      std::format(L"Find UIA test should expose exactly two visible results during {}; saw {}.", phaseLabel, snapshot.resultCount));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(findWindow);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for Find window during {}.", phaseLabel));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Find window should expose visible UI Automation descendants during {}.", phaseLabel));
            state.Require(uiaPatternStats->editControlCount + uiaPatternStats->comboBoxControlCount > 0u,
                          std::format(L"Find window should expose visible UI Automation edit or combo descendants during {}.", phaseLabel));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation ValuePattern support during {}.", phaseLabel));
            state.Require(uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation TogglePattern support during {}.", phaseLabel));
        }
        if (! state.failure.empty())
        {
            return false;
        }

        const auto requireFindSelection = [&](const std::wstring& fullPath, const std::wstring_view expectedLeaf, const std::wstring_view label) noexcept
        {
            state.Require(DebugSelectFindFilesWindowResult(fullPath),
                          std::format(L"Failed to select '{}' in Find results during {} ({})", fullPath, phaseLabel, label));
            if (! state.failure.empty())
            {
                return false;
            }

            FindFilesDebugSnapshot selectedSnapshot{};
            state.Require(DebugGetFindFilesWindowSnapshot(selectedSnapshot),
                          std::format(L"Failed to capture Find snapshot after {} during {}.", label, phaseLabel));
            if (! state.failure.empty())
            {
                return false;
            }

            state.Require(
                selectedSnapshot.selectedResultCount == 1u,
                std::format(
                    L"Find should expose exactly one selected result after {} during {}; saw {}.", label, phaseLabel, selectedSnapshot.selectedResultCount));

            const auto selectionState = CollectVisibleDescendantSelectionPatternState(findWindow, UIA_DataGridControlTypeId);
            state.Require(
                selectionState.has_value(),
                std::format(L"Failed to collect live UI Automation selection state for the Find results grid after {} during {}.", label, phaseLabel));
            if (! selectionState.has_value())
            {
                return false;
            }

            state.Require(selectionState->rootControlType == UIA_DataGridControlTypeId, L"Find results host should expose a UI Automation DataGrid control.");
            state.Require(selectionState->hasSelectionPattern, L"Find results grid should expose SelectionPattern.");
            state.Require(selectionState->selectionCount == 1u,
                          std::format(L"Find results grid should expose exactly one selected UIA row after {} during {}; saw {}.",
                                      label,
                                      phaseLabel,
                                      selectionState->selectionCount));
            state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId, L"Find selected UIA row should expose the DataItem control type.");
            state.Require(selectionState->selectedHasSelectionItemPattern, L"Find selected UIA row should expose SelectionItemPattern.");
            state.Require(! selectionState->selectedName.empty(), L"Find selected UIA row should expose a non-empty accessible name.");
            state.Require(selectionState->selectedName.find(expectedLeaf) != std::wstring::npos,
                          std::format(L"Find selected UIA row name '{}' should include the selected directory name '{}' after {} during {}.",
                                      selectionState->selectedName,
                                      expectedLeaf,
                                      label,
                                      phaseLabel));
            return state.failure.empty();
        };

        state.Require(requireFindSelection(subAPath, subA.filename().native(), L"selecting the first Find result"),
                      std::format(L"Find should keep the UIA results-grid selection synchronized for the first result during {}.", phaseLabel));
        state.Require(requireFindSelection(subBPath, subB.filename().native(), L"switching to the second Find result"),
                      std::format(L"Find should keep the UIA results-grid selection synchronized for the second result during {}.", phaseLabel));
        state.Require(requireFindSelection(subAPath, subA.filename().native(), L"switching back to the first Find result"),
                      std::format(L"Find should keep the UIA results-grid selection synchronized when switching back during {}.", phaseLabel));
        return state.failure.empty();
    };

    state.Require(runLiveUiaCycle(L"initial live UIA validation"), L"Find window initial live UIA validation cycle failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(GetFindFilesWindowHandle() == nullptr || IsWindow(GetFindFilesWindowHandle()) == FALSE,
                  L"Find window should fully close before the reopened live UIA validation cycle.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runLiveUiaCycle(L"reopened live UIA validation"), L"Find window reopened live UIA validation cycle failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogLongRunScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for Find scrolling validation.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_long_run_scrolling";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find long-run scrolling test root.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kResultCount = 48u;
    for (size_t index = 0; index < kResultCount; ++index)
    {
        const std::filesystem::path subdir = root / std::format(L"dir_{:02}", index);
        state.Require(SelfTest::EnsureDirectory(subdir), std::format(L"Failed to create Find long-run scrolling subdirectory '{}'.", subdir.native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanupFindWindow = wil::scope_exit([&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow) != FALSE)
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Find long-run scrolling test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Find long-run scrolling test.");
    FocusFolderViewPane(FolderWindow::Pane::Left);

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in Find long-run scrolling test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for long-run scrolling validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(false, false, true, false, false), L"Failed to configure Find options for long-run scrolling validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"dir_*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for long-run scrolling validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for long-run scrolling validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept { return ! value.searchActive && value.resultCount >= kResultCount; },
                                      SelfTest::Scale(15000ms),
                                      &snapshot),
                  L"Find window did not finish loading enough results for long-run scrolling validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the Find results grid for long-run scrolling validation.");
    state.Require(snapshot.usesDxUiHost, L"Find window should stay attached to the shared DxUi host during long-run scrolling validation.");
    state.Require(snapshot.visibleChildWindowCount <= 1u,
                  std::format(L"Find window should not expose visible child-control fallback during long-run scrolling validation; saw {}.",
                              snapshot.visibleChildWindowCount));
    state.Require(snapshot.visibleResultRowCount > 0u, L"Find results grid should expose visible rows before long-run scrolling validation.");
    state.Require(snapshot.visibleResultColumnCount > 0u, L"Find results grid should expose visible columns before long-run scrolling validation.");
    state.Require(snapshot.visibleResultRowCount < snapshot.resultCount,
                  std::format(L"Find results grid should stay virtualized during long-run scrolling validation; visible rows={} total results={}.",
                              snapshot.visibleResultRowCount,
                              snapshot.resultCount));
    state.Require(snapshot.resultListHasVerticalScrollbar, L"Find results grid should expose a vertical scrollbar during long-run scrolling validation.");
    state.Require(snapshot.dxResizeFailureCount == 0u,
                  std::format(L"Find results grid should start with zero DX resize failures; saw {}.", snapshot.dxResizeFailureCount));

    const size_t initialVisibleRows        = snapshot.visibleResultRowCount;
    const size_t initialVisibleColumns     = snapshot.visibleResultColumnCount;
    const uint32_t initialFullRebuildCount = snapshot.resultListFullRebuildCount;
    const uint64_t initialResizeCount      = snapshot.dxResizeCount;
    uint64_t previousRenderCount           = snapshot.dxRenderCount;

    for (size_t chunk = 0; chunk < 8u; ++chunk)
    {
        state.Require(DebugScrollFindFilesWindowResultsByWheelDetents(-12), std::format(L"Find results grid did not accept long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept { return value.dxRenderCount > previousRenderCount; },
                                          SelfTest::Scale(3000ms),
                                          &snapshot),
                      std::format(L"Find results grid did not repaint after long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }

        previousRenderCount = snapshot.dxRenderCount;
        state.Require(snapshot.resultCount >= kResultCount,
                      std::format(L"Find results grid lost results during long-run scroll chunk {}; saw {}.", chunk, snapshot.resultCount));
        state.Require(snapshot.visibleResultRowCount > 0u && snapshot.visibleResultRowCount <= initialVisibleRows + 1u,
                      std::format(L"Find results grid visible row work became unbounded during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.visibleResultRowCount,
                                  initialVisibleRows));
        state.Require(snapshot.visibleResultColumnCount == initialVisibleColumns,
                      std::format(L"Find results grid visible column work changed unexpectedly during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.visibleResultColumnCount,
                                  initialVisibleColumns));
        state.Require(snapshot.visibleResultCellCount <= snapshot.visibleResultRowCount * snapshot.visibleResultColumnCount,
                      std::format(L"Find results grid visible cell work became inconsistent during chunk {}; saw {} cells for {} rows and {} columns.",
                                  chunk,
                                  snapshot.visibleResultCellCount,
                                  snapshot.visibleResultRowCount,
                                  snapshot.visibleResultColumnCount));
        state.Require(snapshot.resultListHasVerticalScrollbar,
                      std::format(L"Find results grid lost its vertical scrollbar during long-run scroll chunk {}.", chunk));
        state.Require(snapshot.resultListFullRebuildCount == initialFullRebuildCount,
                      std::format(L"Find results grid unexpectedly rebuilt its full list during chunk {}; rebuild count moved from {} to {}.",
                                  chunk,
                                  initialFullRebuildCount,
                                  snapshot.resultListFullRebuildCount));
        state.Require(snapshot.dxResizeCount == initialResizeCount,
                      std::format(L"Find results grid churned DX host resizes during chunk {}; resize count moved from {} to {}.",
                                  chunk,
                                  initialResizeCount,
                                  snapshot.dxResizeCount));
        state.Require(snapshot.dxResizeFailureCount == 0u,
                      std::format(L"Find results grid hit DX resize failures during chunk {}; saw {}.", chunk, snapshot.dxResizeFailureCount));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow) != FALSE)
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_long_run_open_close";
    const std::filesystem::path subA = root / L"sub_a";
    const std::filesystem::path subB = root / L"sub_b";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(subA), L"Failed to create first Find churn subdirectory.");
    state.Require(SelfTest::EnsureDirectory(subB), L"Failed to create second Find churn subdirectory.");
    state.Require(SelfTest::WriteTextFile(subA / L"inside.txt", "payload-a"), L"Failed to create first Find churn file.");
    state.Require(SelfTest::WriteTextFile(subB / L"inside.txt", "payload-b"), L"Failed to create second Find churn file.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Find churn test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Find churn test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub_a", L"sub_b"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for Find churn test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring subAPath = subA.native();
    const std::wstring subBPath = subB.native();

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        closeFindWindow();
        FocusFolderViewPane(FolderWindow::Pane::Left);
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/find during cycle {}.", cycle));

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, std::format(L"Find window did not open during cycle {}.", cycle));
        if (! findWindow || IsWindow(findWindow) == FALSE)
        {
            return false;
        }

        state.Require(! IsOwnedBy(findWindow, mainWindow), std::format(L"Find window should be an independent top-level window during cycle {}.", cycle));
        state.Require(WindowExposesUiaProvider(findWindow), std::format(L"Find window should answer WM_GETOBJECT during cycle {}.", cycle));
        state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), std::format(L"Failed to configure Find options during cycle {}.", cycle));
        state.Require(DebugConfigureFindFilesWindow(
                          root.native(), L"sub*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                      std::format(L"Failed to configure Find window during cycle {}.", cycle));
        state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), std::format(L"Failed to start Find search during cycle {}.", cycle));
        state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                      std::format(L"Find window did not become idle during cycle {}.", cycle));

        const std::wstring& selectedPath     = ((cycle % 2u) == 0u) ? subAPath : subBPath;
        const std::wstring_view expectedLeaf = ((cycle % 2u) == 0u) ? L"sub_a" : L"sub_b";
        state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' during cycle {}.", selectedPath, cycle));

        FindFilesDebugSnapshot snapshot{};
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture Find snapshot during cycle {}.", cycle));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(snapshot.usesDxUiHost, std::format(L"Find window should stay attached to the shared DxUi host during cycle {}.", cycle));
        state.Require(
            snapshot.visibleChildWindowCount <= 1u,
            std::format(L"Find window should not expose visible child-control fallback during cycle {}; saw {}.", cycle, snapshot.visibleChildWindowCount));
        state.Require(snapshot.resultCount == 2u,
                      std::format(L"Find window should keep exactly two visible results during cycle {}; saw {}.", cycle, snapshot.resultCount));
        state.Require(snapshot.selectedResultCount == 1u,
                      std::format(L"Find window should keep exactly one selected result during cycle {}; saw {}.", cycle, snapshot.selectedResultCount));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(findWindow);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for Find window during cycle {}.", cycle));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Find window should expose visible UI Automation descendants during cycle {}.", cycle));
            state.Require(uiaPatternStats->editControlCount + uiaPatternStats->comboBoxControlCount > 0u,
                          std::format(L"Find window should expose visible editable form descendants during cycle {}.", cycle));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation ValuePattern support during cycle {}.", cycle));
            state.Require(uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation TogglePattern support during cycle {}.", cycle));
        }

        const auto selectionState = CollectVisibleDescendantSelectionPatternState(findWindow, UIA_DataGridControlTypeId);
        state.Require(selectionState.has_value(),
                      std::format(L"Failed to collect live UI Automation selection state for the Find results grid during cycle {}.", cycle));
        if (selectionState.has_value())
        {
            state.Require(selectionState->hasSelectionPattern, std::format(L"Find results grid should expose SelectionPattern during cycle {}.", cycle));
            state.Require(
                selectionState->selectionCount == 1u,
                std::format(L"Find results grid should keep exactly one selected row during cycle {}; saw {}.", cycle, selectionState->selectionCount));
            state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                          std::format(L"Find selected UIA row should remain a DataItem during cycle {}.", cycle));
            state.Require(selectionState->selectedHasSelectionItemPattern,
                          std::format(L"Find selected UIA row should keep SelectionItemPattern during cycle {}.", cycle));
            state.Require(
                selectionState->selectedName.find(expectedLeaf) != std::wstring::npos,
                std::format(L"Find selected UIA row name should track '{}' during cycle {}; saw '{}'.", expectedLeaf, cycle, selectionState->selectedName));
        }

        auto valueState = CollectVisibleDescendantValuePatternState(findWindow, UIA_EditControlTypeId);
        if (! valueState.has_value())
        {
            valueState = CollectVisibleDescendantValuePatternState(findWindow, UIA_ComboBoxControlTypeId);
        }
        state.Require(valueState.has_value(), std::format(L"Failed to collect UI Automation ValuePattern state for the Find form during cycle {}.", cycle));
        if (valueState.has_value())
        {
            state.Require(! valueState->isReadOnly, std::format(L"Find visible DX edit surface should remain editable during cycle {}.", cycle));
            state.Require(! valueState->name.empty(),
                          std::format(L"Find visible DX edit surface should expose a stable accessible name during cycle {}.", cycle));
        }

        const auto buttonState = CollectVisibleDescendantNamedElementState(findWindow, UIA_ButtonControlTypeId);
        state.Require(buttonState.has_value(), std::format(L"Failed to collect a visible DX command button state for the Find window during cycle {}.", cycle));
        if (buttonState.has_value())
        {
            state.Require(! buttonState->name.empty(),
                          std::format(L"Find visible DX command button should expose a stable accessible name during cycle {}.", cycle));
        }

        PostMessageW(findWindow, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)), std::format(L"Find window did not close cleanly during cycle {}.", cycle));
        state.Require(GetFindFilesWindowHandle() == nullptr || IsWindow(GetFindFilesWindowHandle()) == FALSE,
                      std::format(L"Find window should not linger after close during cycle {}.", cycle));
    }

    state.Require(GetFindFilesWindowHandle() == nullptr || IsWindow(GetFindFilesWindowHandle()) == FALSE,
                  L"Find window should not remain open after repeated churn.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogDirectoryActivationNavigatesIntoSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_directory_activation";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create activation test subdirectory.");
    state.Require(SelfTest::WriteTextFile(sub / L"inside.txt", "payload"), L"Failed to create activation test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Find activation test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Find activation test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for Find activation test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in directory activation test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for directory activation test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), L"Failed to configure Find options for directory activation test.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for directory activation test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for directory activation test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for directory activation test.");

    const std::wstring subPath = sub.native();
    state.Require(DebugSelectFindFilesWindowResult(subPath), std::format(L"Failed to select '{}' in Find results.", subPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot for directory activation test.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.selectedResultCount == 1u, std::format(L"Expected one selected Find result, got {}.", snapshot.selectedResultCount));
    state.Require(snapshot.openButtonEnabled, L"Open button should be enabled after selecting a Find result.");
    state.Require(snapshot.parentButtonEnabled, L"Parent button should be enabled after selecting a Find result.");

    state.Require(DebugActivateSelectedFindFilesWindowResult(), L"Failed to activate selected Find result.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sub, SelfTest::Scale(5000ms)),
                  L"Activating directory Find result did not navigate into the selected directory.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogOpenParentKeepsDirectoryFocusedInParent(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_open_parent";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create open-parent test subdirectory.");
    state.Require(SelfTest::WriteTextFile(sub / L"inside.txt", "payload"), L"Failed to create open-parent test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Find open-parent test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Find open-parent test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for Find open-parent test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in open-parent test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for open-parent test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), L"Failed to configure Find options for open-parent test.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for open-parent test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for open-parent test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for open-parent test.");

    const std::wstring subPath = sub.native();
    state.Require(DebugSelectFindFilesWindowResult(subPath), std::format(L"Failed to select '{}' in Find results for open-parent test.", subPath));
    state.Require(DebugOpenSelectedFindFilesWindowResultParent(), L"Failed to invoke Open Parent on the selected Find result.");

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(5000ms)),
                  L"Open Parent did not navigate to the selected result parent folder.");

    const auto focusDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < focusDeadline)
    {
        MSG msg{};
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        if (g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"sub")
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    state.Require(false,
                  std::format(L"Open Parent should focus 'sub' in '{}', got focus on '{}'.",
                              root.native(),
                              std::wstring(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left))));
    return false;
}

[[nodiscard]] bool TestFindDialogEnterFromCheckboxInvokesDefaultSearch(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_default_enter";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create default-enter test directory.");
    state.Require(SelfTest::WriteTextFile(root / L"match.txt", "payload"), L"Failed to create default-enter test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in default-enter test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for default-enter test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"match.txt", L"", Common::Settings::SearchNameMode::Literal, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for default-enter test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for default-enter test.");
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::RecursiveCheck), L"Failed to focus recursive checkbox before default-enter test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture pre-search snapshot for default-enter test.");
    state.Require(snapshot.focusTarget == FindFilesDebugFocusTarget::RecursiveCheck,
                  L"Recursive checkbox did not keep focus before Enter default-button routing.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(findWindow, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(findWindow, WM_KEYUP, VK_RETURN, 0);

    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for default-enter test.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.resultCount == 1u; }, SelfTest::Scale(3000ms), &snapshot),
                  L"Pressing Enter from a non-editor Find control did not trigger the default Find action.");
    state.Require(snapshot.statusText != LoadStringResource(nullptr, IDS_FIND_STATUS_READY), L"Default Enter should change the Find status away from Ready.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogPointerClickTogglesRecursiveCheckbox(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
            static_cast<void>(WaitForFindWindowUnavailable(SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"),
                  L"Shortcut dispatch failed for cmd/pane/find in recursive-checkbox pointer test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for recursive-checkbox pointer test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(false, true, false, true, false), L"Failed to configure Find options for recursive-checkbox pointer test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && ! value.searchActive && value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u &&
               ! value.recursiveChecked;
    },
                      SelfTest::Scale(10000ms),
                      &snapshot),
                  L"Find window did not reach the expected baseline state before recursive-checkbox pointer test.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT recursiveRect{};
    state.Require(DebugGetFindFilesWindowTargetClientRect(FindFilesDebugFocusTarget::RecursiveCheck, recursiveRect),
                  L"Failed to capture recursive-checkbox client rect for pointer test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int clickX                     = recursiveRect.left + ((recursiveRect.right - recursiveRect.left) / 2);
    const int clickY                     = recursiveRect.top + ((recursiveRect.bottom - recursiveRect.top) / 2);
    const LPARAM clickPoint              = MAKELPARAM(clickX, clickY);
    const uint64_t baselineRenderCount   = snapshot.dxRenderCount;
    const uint64_t baselineResizeCount   = snapshot.dxResizeCount;
    const size_t baselineVisibleChildren = snapshot.visibleChildWindowCount;

    const auto clickRecursive = [&]() noexcept
    {
        SendMessageW(findWindow, WM_MOUSEMOVE, 0, clickPoint);
        SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        SendMessageW(findWindow, WM_LBUTTONUP, 0, clickPoint);
        PumpPendingMessages();
    };

    const auto requireRecursiveState = [&](const bool expectedChecked, std::wstring_view label) noexcept
    {
        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && ! value.searchActive && value.visibleChildWindowCount == baselineVisibleChildren && value.dxResizeFailureCount == 0u &&
                   value.dxResizeCount == baselineResizeCount && value.focusTarget == FindFilesDebugFocusTarget::RecursiveCheck &&
                   value.recursiveChecked == expectedChecked && value.dxRenderCount >= baselineRenderCount;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      std::format(L"Recursive checkbox did not reach the expected state after {}.", label));
    };

    clickRecursive();
    requireRecursiveState(true, L"first pointer click");
    if (! state.failure.empty())
    {
        return false;
    }

    clickRecursive();
    requireRecursiveState(false, L"second pointer click");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogEscapeClosesPopupBeforeCancel(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in popup-escape test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for popup-escape test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameModeCombo), L"Failed to focus name-mode combo before popup-escape test.");

    SendMessageW(findWindow, WM_SYSKEYDOWN, VK_DOWN, 0);
    SendMessageW(findWindow, WM_SYSKEYUP, VK_DOWN, 0);

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return value.nameModePopupOpen && value.focusTarget == FindFilesDebugFocusTarget::NameModeCombo; },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Alt+Down did not open the focused Find mode combo popup.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(findWindow, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(findWindow, WM_KEYUP, VK_ESCAPE, 0);

    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return ! value.nameModePopupOpen && value.focusTarget == FindFilesDebugFocusTarget::NameModeCombo && ! value.searchActive; },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Esc should close the focused Find combo popup before any cancel routing.");
    state.Require(IsWindow(findWindow) != FALSE, L"Esc should not close the Find window while dismissing a combo popup.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogEscapeFromDxControlClosesCancel(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
            static_cast<void>(WaitForFindWindowUnavailable(SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in cancel-escape test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for cancel-escape test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::RecursiveCheck), L"Failed to focus recursive checkbox before cancel-escape test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::RecursiveCheck && ! value.searchActive && value.usesDxUiHost &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(10000ms),
                      &snapshot),
                  L"Find window did not reach the expected focused DX idle state before cancel-escape test.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(findWindow, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(findWindow, WM_KEYUP, VK_ESCAPE, 0);

    const bool closedAfterEscape = WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms));
    if (! closedAfterEscape && IsWindow(findWindow) != FALSE)
    {
        PostMessageW(findWindow, WM_CLOSE, 0, 0);
        static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
    }

    state.Require(closedAfterEscape, L"Esc from a focused DX Find control should close the dialog through cancel routing.");
    state.Require(GetFindFilesWindowHandle() == nullptr || IsWindow(GetFindFilesWindowHandle()) == FALSE,
                  L"Find window should not remain open after Esc cancel routing.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogAccessKeysFocusExpectedFields(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in access-key test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for access-key test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    const auto expectFocusAfterMnemonic = [&](wchar_t mnemonic, FindFilesDebugFocusTarget expected, std::wstring_view label) noexcept
    {
        FindFilesDebugSnapshot snapshot{};
        SendMessageW(findWindow, WM_SYSCHAR, mnemonic, 0);
        state.Require(WaitForFindSnapshot([expected](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == expected; },
                                          SelfTest::Scale(2000ms),
                                          &snapshot),
                      std::format(L"Mnemonic '{}' did not focus the expected Find field '{}'.", mnemonic, label));
    };

    expectFocusAfterMnemonic(L'l', FindFilesDebugFocusTarget::RootCombo, L"Look in");
    expectFocusAfterMnemonic(L'n', FindFilesDebugFocusTarget::NameCombo, L"Named");
    expectFocusAfterMnemonic(L'c', FindFilesDebugFocusTarget::ContentCombo, L"Containing");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogGridFocusedEnterActivatesSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_grid_enter";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create grid-enter test subdirectory.");
    state.Require(SelfTest::WriteTextFile(sub / L"inside.txt", "payload"), L"Failed to create grid-enter test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for grid-enter test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for grid-enter test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for grid-enter test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in grid-enter test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for grid-enter test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), L"Failed to configure Find options for grid-enter test.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for grid-enter test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for grid-enter test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for grid-enter test.");

    const std::wstring subPath = sub.native();
    state.Require(DebugSelectFindFilesWindowResult(subPath), std::format(L"Failed to select '{}' in Find results for grid-enter test.", subPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot for grid-enter test.");
    state.Require(snapshot.focusTarget == FindFilesDebugFocusTarget::ResultsGrid,
                  L"Selecting a Find result for grid-enter test should focus the results grid.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(findWindow, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(findWindow, WM_KEYUP, VK_RETURN, 0);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sub, SelfTest::Scale(5000ms)),
                  L"Pressing Enter on a focused Find result did not activate the selected row.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogGridDoubleClickActivatesSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_grid_double_click";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create find_grid_double_click test folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for grid double-click test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for grid double-click test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for grid double-click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in grid double-click test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for grid double-click test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), L"Failed to configure Find options for grid double-click test.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for grid double-click test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for grid double-click test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for grid double-click test.");

    const std::wstring subPath = sub.native();
    state.Require(DebugSelectFindFilesWindowResult(subPath), std::format(L"Failed to select '{}' in Find results for grid double-click test.", subPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot for grid double-click test.");
    state.Require(snapshot.focusTarget == FindFilesDebugFocusTarget::ResultsGrid,
                  L"Selecting a Find result for grid double-click test should focus the results grid.");
    state.Require(snapshot.selectedResultRowRect.right > snapshot.selectedResultRowRect.left &&
                      snapshot.selectedResultRowRect.bottom > snapshot.selectedResultRowRect.top,
                  L"Find results grid did not expose selected-row geometry for grid double-click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LPARAM doubleClickPoint = DipPointToEffectiveMouseInputLParam(findWindow,
                                                                        (snapshot.selectedResultRowRect.left + snapshot.selectedResultRowRect.right) * 0.5f,
                                                                        (snapshot.selectedResultRowRect.top + snapshot.selectedResultRowRect.bottom) * 0.5f);
    SendMouseDoubleClickToWindow(findWindow, doubleClickPoint);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sub, SelfTest::Scale(5000ms)),
                  L"Double-clicking a focused Find result did not activate the selected row.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogTabTraversalMatchesExpectedOrder(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path findRoot = suiteRoot / L"work" / L"find_tab_traversal";
    const std::filesystem::path subDir   = findRoot / L"sub";
    const std::filesystem::path hitFile  = subDir / L"needle.txt";

    std::error_code ec;
    std::filesystem::remove_all(findRoot, ec);
    state.Require(SelfTest::EnsureDirectory(subDir), L"Failed to create find_tab_traversal folder.");
    state.Require(SelfTest::WriteTextFile(hitFile, "needle\n"), L"Failed to seed find_tab_traversal result file.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in tab-traversal test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for tab-traversal test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugConfigureFindFilesWindow(
                      findRoot.native(), L"*.txt", L"needle", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure Find window for tab-traversal test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, true), L"Failed to configure Find options for tab-traversal test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for tab-traversal test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for tab-traversal test.");
    state.Require(DebugSelectFindFilesWindowResult(hitFile.native()),
                  std::format(L"Failed to select '{}' before Find tab-traversal validation.", hitFile.native()));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return ! value.searchActive && value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u &&
               value.resultCount == 1u && value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled &&
               value.appendButtonEnabled && value.intersectButtonEnabled && value.subtractButtonEnabled && value.visibleResultRowCount > 0u &&
               value.visibleResultColumnCount > 0u && value.visibleResultCellCount > 0u;
    },
                      SelfTest::Scale(5000ms),
                      &snapshot),
                  L"Find window did not settle to the expected DX results state before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::RootCombo), L"Failed to focus root combo before tab-traversal test.");
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::RootCombo && ! value.searchActive && value.usesDxUiHost && value.visibleChildWindowCount <= 1u &&
               value.dxResizeFailureCount == 0u && value.resultCount == 1u && value.selectedResultCount == 1u;
    },
                      SelfTest::Scale(2000ms),
                      &snapshot),
                  L"Root combo did not take focus before Find tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineResizeCount      = snapshot.dxResizeCount;
    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;

    const auto sendTab = [&](bool reverse, FindFilesDebugFocusTarget expected, std::wstring_view label) noexcept
    {
        if (reverse)
        {
            SendMessageW(findWindow, WM_KEYDOWN, VK_SHIFT, 0);
        }
        SendMessageW(findWindow, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(findWindow, WM_KEYUP, VK_TAB, 0);
        if (reverse)
        {
            SendMessageW(findWindow, WM_KEYUP, VK_SHIFT, 0);
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.focusTarget == expected && ! value.searchActive && value.usesDxUiHost && value.visibleChildWindowCount <= 1u &&
                   value.dxResizeFailureCount == 0u && value.dxResizeCount == baselineResizeCount && value.resultCount == 1u &&
                   value.selectedResultCount == 1u && value.visibleResultRowCount == baselineVisibleRowCount &&
                   value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount;
        },
                          SelfTest::Scale(2000ms),
                          &snapshot),
                      std::format(L"{} focus target not reached during Find tab traversal.", label));
    };

    sendTab(false, FindFilesDebugFocusTarget::NameCombo, L"Name combo");
    sendTab(false, FindFilesDebugFocusTarget::NameModeCombo, L"Name mode combo");
    sendTab(false, FindFilesDebugFocusTarget::ContentCombo, L"Content combo");
    sendTab(false, FindFilesDebugFocusTarget::ContentModeCombo, L"Content mode combo");
    sendTab(false, FindFilesDebugFocusTarget::RecursiveCheck, L"Recursive checkbox");
    sendTab(false, FindFilesDebugFocusTarget::IncludeFilesCheck, L"Include files checkbox");
    sendTab(false, FindFilesDebugFocusTarget::IncludeDirectoriesCheck, L"Include directories checkbox");
    sendTab(false, FindFilesDebugFocusTarget::FollowSymlinksCheck, L"Follow symlinks checkbox");
    sendTab(false, FindFilesDebugFocusTarget::MatchCaseNameCheck, L"Match-case name checkbox");
    sendTab(false, FindFilesDebugFocusTarget::MatchCaseContentCheck, L"Match-case content checkbox");
    sendTab(false, FindFilesDebugFocusTarget::PreferIndexCheck, L"Prefer index checkbox");
    sendTab(false, FindFilesDebugFocusTarget::WantSnippetsCheck, L"Include snippets checkbox");
    sendTab(false, FindFilesDebugFocusTarget::FindButton, L"Find button");
    sendTab(false, FindFilesDebugFocusTarget::AppendButton, L"Append button");
    sendTab(false, FindFilesDebugFocusTarget::IntersectButton, L"Intersect button");
    sendTab(false, FindFilesDebugFocusTarget::SubtractButton, L"Subtract button");
    sendTab(false, FindFilesDebugFocusTarget::OpenButton, L"Open button");
    sendTab(false, FindFilesDebugFocusTarget::ParentButton, L"Open Parent button");
    sendTab(false, FindFilesDebugFocusTarget::ResultsGrid, L"results grid");
    sendTab(false, FindFilesDebugFocusTarget::RootCombo, L"wrapped root combo");

    sendTab(true, FindFilesDebugFocusTarget::ResultsGrid, L"reverse wrapped results grid");
    sendTab(true, FindFilesDebugFocusTarget::ParentButton, L"reverse Open Parent button");
    sendTab(true, FindFilesDebugFocusTarget::OpenButton, L"reverse Open button");
    sendTab(true, FindFilesDebugFocusTarget::SubtractButton, L"reverse Subtract button");
    sendTab(true, FindFilesDebugFocusTarget::IntersectButton, L"reverse Intersect button");
    sendTab(true, FindFilesDebugFocusTarget::AppendButton, L"reverse Append button");
    sendTab(true, FindFilesDebugFocusTarget::FindButton, L"reverse Find button");
    sendTab(true, FindFilesDebugFocusTarget::WantSnippetsCheck, L"reverse Include snippets checkbox");
    sendTab(true, FindFilesDebugFocusTarget::PreferIndexCheck, L"reverse Prefer index checkbox");
    sendTab(true, FindFilesDebugFocusTarget::MatchCaseContentCheck, L"reverse Match-case content checkbox");
    sendTab(true, FindFilesDebugFocusTarget::MatchCaseNameCheck, L"reverse Match-case name checkbox");
    sendTab(true, FindFilesDebugFocusTarget::FollowSymlinksCheck, L"reverse Follow symlinks checkbox");
    sendTab(true, FindFilesDebugFocusTarget::IncludeDirectoriesCheck, L"reverse Include directories checkbox");
    sendTab(true, FindFilesDebugFocusTarget::IncludeFilesCheck, L"reverse Include files checkbox");
    sendTab(true, FindFilesDebugFocusTarget::RecursiveCheck, L"reverse Recursive checkbox");
    sendTab(true, FindFilesDebugFocusTarget::ContentModeCombo, L"reverse Content mode combo");
    sendTab(true, FindFilesDebugFocusTarget::ContentCombo, L"reverse Content combo");
    sendTab(true, FindFilesDebugFocusTarget::NameModeCombo, L"reverse Name mode combo");
    sendTab(true, FindFilesDebugFocusTarget::NameCombo, L"reverse Name combo");
    sendTab(true, FindFilesDebugFocusTarget::RootCombo, L"reverse Root combo");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogModeTypeaheadUpdatesSelectionAndDependencies(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in mode-typeahead test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for mode-typeahead test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugConfigureFindFilesWindow(L"C:\\", L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for mode-typeahead test.");

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameModeCombo), L"Failed to focus name-mode combo before typeahead test.");
    SendMessageW(findWindow, WM_CHAR, L'r', 0);

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::NameModeCombo && value.nameModeSelectedIndex.has_value() &&
               value.nameModeSelectedIndex.value() == 2u;
    },
                      SelfTest::Scale(2000ms),
                      &snapshot),
                  L"Typing on the focused name-mode combo did not select Regex via DX combo typeahead.");

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ContentModeCombo),
                  L"Failed to focus content-mode combo before typeahead dependency test.");
    SendMessageW(findWindow, WM_CHAR, L'r', 0);
    const bool contentModeReady = WaitForFindSnapshot(
        [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::ContentModeCombo && value.contentModeSelectedIndex.has_value() &&
               value.contentModeSelectedIndex.value() == 2u && value.matchCaseContentEnabled && value.wantSnippetsEnabled;
    },
        SelfTest::Scale(2000ms),
        &snapshot);
    state.Require(contentModeReady,
                  std::format(L"Typing on the focused content-mode combo did not select Regex and enable content-specific options. "
                              L"(focus={}, selected={}, matchCaseEnabled={}, wantSnippetsEnabled={}, popupOpen={}, contentText='{}')",
                              static_cast<uint32_t>(snapshot.focusTarget),
                              snapshot.contentModeSelectedIndex ? std::format(L"{}", snapshot.contentModeSelectedIndex.value()) : L"<none>",
                              snapshot.matchCaseContentEnabled ? L"true" : L"false",
                              snapshot.wantSnippetsEnabled ? L"true" : L"false",
                              snapshot.contentModePopupOpen ? L"true" : L"false",
                              snapshot.contentPatternText));

    std::this_thread::sleep_for(SelfTest::Scale(1200ms));
    PumpPendingMessages();
    SendMessageW(findWindow, WM_CHAR, L'd', 0);
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::ContentModeCombo && value.contentModeSelectedIndex.has_value() &&
               value.contentModeSelectedIndex.value() == 0u && ! value.matchCaseContentEnabled && ! value.wantSnippetsEnabled;
    },
                      SelfTest::Scale(2000ms),
                      &snapshot),
                  L"Selecting Disabled via DX combo typeahead should disable the content-specific option row.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogCommandEnablementMatchesIdleRunningAndSelectionStates(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_command_enablement";
    const std::filesystem::path sub  = root / L"sub";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create command-enablement test subdirectory.");
    state.Require(SelfTest::WriteTextFile(sub / L"inside.txt", "payload"), L"Failed to create command-enablement result file.");

    std::string largeBody(2u * 1024u * 1024u, 'x');
    largeBody.back() = '\n';
    for (int index = 0; index < 24; ++index)
    {
        const std::filesystem::path filePath =
            (index % 2 == 0) ? (root / std::format(L"root_{:03}.json", index)) : (sub / std::format(L"sub_{:03}.json", index));
        state.Require(SelfTest::WriteTextFile(filePath, largeBody),
                      std::format(L"Failed to create '{}' for command-enablement test.", filePath.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in command-enablement test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for command-enablement test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, true, true, false, true), L"Failed to configure Find options for command-enablement test.");
    state.Require(DebugConfigureFindFilesWindow(root.native(),
                                                L"*.json",
                                                L"ZZZ_NOT_PRESENT_123456789",
                                                Common::Settings::SearchNameMode::Wildcard,
                                                Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure Find window for command-enablement test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture idle Find snapshot for command-enablement test.");
    state.Require(snapshot.findButtonEnabled && snapshot.appendButtonEnabled && snapshot.intersectButtonEnabled && snapshot.subtractButtonEnabled,
                  L"Search command buttons should be enabled while the Find window is idle.");
    state.Require(! snapshot.cancelButtonEnabled && ! snapshot.openButtonEnabled && ! snapshot.parentButtonEnabled,
                  L"Cancel/Open/Parent buttons should be disabled while idle with no selection.");
    state.Require(snapshot.rootComboEnabled && snapshot.nameComboEnabled && snapshot.nameModeComboEnabled && snapshot.contentComboEnabled &&
                      snapshot.contentModeComboEnabled,
                  L"Form combos should remain enabled while the Find window is idle.");
    state.Require(snapshot.matchCaseContentEnabled && snapshot.wantSnippetsEnabled,
                  L"Content-mode-dependent controls should be enabled when content mode is active.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for command-enablement test.");
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.searchActive && ! value.findButtonEnabled && ! value.appendButtonEnabled && ! value.intersectButtonEnabled &&
               ! value.subtractButtonEnabled && value.cancelButtonEnabled && ! value.openButtonEnabled && ! value.parentButtonEnabled &&
               ! value.rootComboEnabled && ! value.nameComboEnabled && ! value.nameModeComboEnabled && ! value.contentComboEnabled &&
               ! value.contentModeComboEnabled;
    },
                      SelfTest::Scale(10000ms),
                      &snapshot),
                  L"Active Find search did not expose the expected disabled/enabled command state.");
    if (! state.failure.empty())
    {
        static_cast<void>(DebugCancelFindFilesWindowSearch());
        static_cast<void>(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())));
        return false;
    }

    state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel Find search for command-enablement test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle after cancellation in command-enablement test.");

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to reconfigure Find window for selection-state command-enablement test.");
    state.Require(DebugSetFindFilesWindowOptions(true, false, true, false, false),
                  L"Failed to configure selection-state Find options for command-enablement test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start selection-state Find search for command-enablement test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Selection-state Find search did not become idle in command-enablement test.");
    state.Require(DebugSelectFindFilesWindowResult(sub.native()), std::format(L"Failed to select '{}' for command-enablement test.", sub.native()));
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return ! value.searchActive && value.findButtonEnabled && value.appendButtonEnabled && value.intersectButtonEnabled && value.subtractButtonEnabled &&
               ! value.cancelButtonEnabled && value.openButtonEnabled && value.parentButtonEnabled && value.selectedResultCount == 1u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Selecting a Find result did not restore idle commands and enable Open/Parent.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogActionButtonsActivateExpectedCommands(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_action_buttons";
    const std::filesystem::path sub  = root / L"sub";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create action-buttons test directory.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.jsonl", "payload"), L"Failed to create alpha.jsonl.");
    state.Require(SelfTest::WriteTextFile(root / L"apple.txt", "payload"), L"Failed to create apple.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "payload"), L"Failed to create beta.txt.");
    state.Require(SelfTest::WriteTextFile(sub / L"inside.txt", "payload"), L"Failed to create nested test file.");

    std::string largeBody(2u * 1024u * 1024u, 'x');
    largeBody.back() = '\n';
    for (int index = 0; index < 24; ++index)
    {
        const std::filesystem::path filePath =
            (index % 2 == 0) ? (root / std::format(L"cancel_{:03}.json", index)) : (sub / std::format(L"cancel_{:03}.json", index));
        state.Require(SelfTest::WriteTextFile(filePath, largeBody),
                      std::format(L"Failed to create '{}' for action-buttons cancellation coverage.", filePath.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for action-buttons test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for action-buttons test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.jsonl", L"apple.txt", L"beta.txt", L"sub"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for action-buttons test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in action-buttons test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for action-buttons test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::filesystem::path& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path.native()) != snapshot.fullPaths.end(); };

    const std::wstring findButtonText      = LoadStringResource(nullptr, IDS_FIND_ACTION_FIND);
    const std::wstring appendButtonText    = LoadStringResource(nullptr, IDS_FIND_ACTION_APPEND);
    const std::wstring intersectButtonText = LoadStringResource(nullptr, IDS_FIND_ACTION_INTERSECT);
    const std::wstring subtractButtonText  = LoadStringResource(nullptr, IDS_FIND_ACTION_SUBTRACT);
    const std::wstring cancelButtonText    = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    const std::wstring openButtonText      = LoadStringResource(nullptr, IDS_FIND_ACTION_OPEN);
    const std::wstring parentButtonText    = LoadStringResource(nullptr, IDS_FIND_ACTION_PARENT);

    state.Require(! findButtonText.empty(), L"Failed to resolve Find button caption for action-buttons test.");
    state.Require(! appendButtonText.empty(), L"Failed to resolve Append button caption for action-buttons test.");
    state.Require(! intersectButtonText.empty(), L"Failed to resolve Intersect button caption for action-buttons test.");
    state.Require(! subtractButtonText.empty(), L"Failed to resolve Subtract button caption for action-buttons test.");
    state.Require(! cancelButtonText.empty(), L"Failed to resolve Cancel button caption for action-buttons test.");
    state.Require(! openButtonText.empty(), L"Failed to resolve Open button caption for action-buttons test.");
    state.Require(! parentButtonText.empty(), L"Failed to resolve Open Parent button caption for action-buttons test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto invokeButton = [&](std::wstring_view buttonText, std::wstring_view label) noexcept
    {
        state.Require(InvokeVisibleDescendantByName(findWindow, UIA_ButtonControlTypeId, buttonText),
                      std::format(L"{} did not expose live UIA InvokePattern interaction.", label));
        if (! state.failure.empty())
        {
            return;
        }
    };

    FindFilesDebugSnapshot snapshot{};

    state.Require(DebugSetFindFilesWindowOptions(false, true, false, true, false), L"Failed to configure root-level Find options for action-buttons test.");
    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure initial Append search for action-buttons test.");
    invokeButton(appendButtonText, L"Append button before initial append search");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Initial Append search did not become idle in action-buttons test.");
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.resultCount == 1u && containsPath(value, root / L"alpha.jsonl"); },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Append button live UIA invoke did not add the expected initial result.");

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure second Append search for action-buttons test.");
    invokeButton(appendButtonText, L"Append button before second append search");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Second Append search did not become idle in action-buttons test.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 3u && containsPath(value, root / L"alpha.jsonl") && containsPath(value, root / L"apple.txt") &&
               containsPath(value, root / L"beta.txt");
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Append button live UIA invoke did not merge the expected TXT results.");

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"a*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Intersect search for action-buttons test.");
    invokeButton(intersectButtonText, L"Intersect button before intersect search");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Intersect search did not become idle in action-buttons test.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && containsPath(value, root / L"alpha.jsonl") && containsPath(value, root / L"apple.txt") &&
               ! containsPath(value, root / L"beta.txt");
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Intersect button live UIA invoke did not retain only the expected results.");

    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Subtract search for action-buttons test.");
    invokeButton(subtractButtonText, L"Subtract button before subtract search");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Subtract search did not become idle in action-buttons test.");
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.resultCount == 1u && containsPath(value, root / L"apple.txt") && ! containsPath(value, root / L"alpha.jsonl"); },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Subtract button live UIA invoke did not remove the expected JSONL result.");

    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to configure cancellation options for action-buttons test.");
    state.Require(DebugConfigureFindFilesWindow(root.native(),
                                                L"*.json",
                                                L"ZZZ_NOT_PRESENT_123456789",
                                                Common::Settings::SearchNameMode::Wildcard,
                                                Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure cancellation search for action-buttons test.");
    invokeButton(findButtonText, L"Find button before cancellation search");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return value.searchActive && value.cancelButtonEnabled && ! value.findButtonEnabled; },
                                      SelfTest::Scale(10000ms),
                                      &snapshot),
                  L"Find button live UIA invoke did not start the cancellable search.");
    invokeButton(cancelButtonText, L"Cancel button during active search");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Cancel button live UIA invoke did not return Find to idle.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return ! value.searchActive && ! value.cancelButtonEnabled && value.findButtonEnabled; },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Cancel button live UIA invoke did not restore the idle Find command state.");

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), L"Failed to configure directory-result options for action-buttons test.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure directory-result search for action-buttons test.");
    invokeButton(findButtonText, L"Find button before directory-result search");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Directory-result search did not become idle in action-buttons test.");
    state.Require(DebugSelectFindFilesWindowResult(sub.native()), std::format(L"Failed to select '{}' for action-buttons test.", sub.native()));
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && containsPath(value, sub); },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Selecting the directory result did not enable Open/Parent for action-buttons test.");

    invokeButton(openButtonText, L"Open button before directory activation");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sub, SelfTest::Scale(5000ms)),
                  L"Open button live UIA invoke did not navigate into the selected directory.");

    state.Require(DebugSelectFindFilesWindowResult(sub.native()),
                  std::format(L"Failed to reselect '{}' before Parent action in action-buttons test.", sub.native()));
    invokeButton(parentButtonText, L"Parent button before parent navigation");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(5000ms)),
                  L"Parent button live UIA invoke did not return to the selected directory parent.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.jsonl", L"apple.txt", L"beta.txt", L"sub"}, SelfTest::Scale(3000ms)),
                  L"Parent button live UIA invoke did not restore the parent directory contents.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresPersistedGridLayout(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_persisted_grid_layout";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create persisted-grid-layout test root.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode          = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode       = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout = {
        Common::Settings::GridColumnLayoutEntry{.columnId = L"path", .displayIndex = 0u, .widthDip = 420.0f},
        Common::Settings::GridColumnLayoutEntry{.columnId = L"modified", .displayIndex = 1u, .widthDip = 180.0f},
        Common::Settings::GridColumnLayoutEntry{.columnId = L"name", .displayIndex = 2u, .widthDip = 260.0f},
    };
    g_settings.search = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-persisted-grid-layout");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for persisted-grid-layout test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for persisted-grid-layout test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept { return value.resultColumnIds.size() >= 5u; }, SelfTest::Scale(2000ms), &snapshot),
                  L"Find snapshot did not expose grid columns for persisted-grid-layout test.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.resultColumnIds.size() == snapshot.resultColumnWidthsDip.size(), L"Grid layout snapshot shape mismatch.");
    state.Require(snapshot.resultColumnIds.size() >= 5u, L"Expected base Find grid columns in persisted-grid-layout snapshot.");
    if (snapshot.resultColumnIds.size() >= 5u && snapshot.resultColumnWidthsDip.size() >= 5u)
    {
        state.Require(snapshot.resultColumnIds[0] == L"path", L"Persisted Find grid layout did not move Path to display index 0.");
        state.Require(snapshot.resultColumnIds[1] == L"modified", L"Persisted Find grid layout did not move Modified to display index 1.");
        state.Require(snapshot.resultColumnIds[2] == L"name", L"Persisted Find grid layout did not move Name to display index 2.");
        state.Require(std::fabs(snapshot.resultColumnWidthsDip[0] - 420.0f) <= 0.1f, L"Persisted Find grid layout did not restore the Path column width.");
        state.Require(std::fabs(snapshot.resultColumnWidthsDip[1] - 180.0f) <= 0.1f, L"Persisted Find grid layout did not restore the Modified column width.");
        state.Require(std::fabs(snapshot.resultColumnWidthsDip[2] - 260.0f) <= 0.1f, L"Persisted Find grid layout did not restore the Name column width.");
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Find grid layout close path should keep search settings present.");
    if (g_settings.search.has_value())
    {
        const auto& savedLayout = g_settings.search->resultsGridLayout;
        state.Require(savedLayout.size() >= 5u, L"Closing the Find window should persist the active grid layout.");
        if (savedLayout.size() >= 3u)
        {
            state.Require(savedLayout[0].columnId == L"path", L"Persisted Find layout should keep Path first after close.");
            state.Require(savedLayout[1].columnId == L"modified", L"Persisted Find layout should keep Modified second after close.");
            state.Require(savedLayout[2].columnId == L"name", L"Persisted Find layout should keep Name third after close.");
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogHeaderDragReordersColumnsWithoutSort(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto restoreSearch                                                   = wil::scope_exit([&] noexcept { g_settings.search = previousSearch; });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_header_drag_reorder";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find header-reorder test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find header-reorder subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find header reorder.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find header reorder.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-header-reorder");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for header reorder validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for header reorder validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for header reorder validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for header reorder validation.");
    if (! DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find))
    {
        FindFilesDebugSnapshot startFailureSnapshot{};
        static_cast<void>(DebugGetFindFilesWindowSnapshot(startFailureSnapshot));
        state.Require(false, std::format(L"Failed to start Find search for header reorder validation. {}", DescribeFindSnapshotBrief(startFailureSnapshot)));
    }
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for header reorder validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for header reorder validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot before header reorder validation.");
    state.Require(snapshot.usesDxUiHost && snapshot.resultCount == 2u && snapshot.selectedResultCount == 1u && snapshot.visibleChildWindowCount <= 1u &&
                      snapshot.resultColumnIds.size() >= 2u && snapshot.firstResultHeaderRect.right > snapshot.firstResultHeaderRect.left &&
                      snapshot.secondResultHeaderRect.right > snapshot.secondResultHeaderRect.left && snapshot.resultColumnIds[0] == L"name" &&
                      snapshot.resultColumnIds[1] == L"path" && snapshot.dxResizeFailureCount == 0u && containsPath(snapshot, selectedPath),
                  std::format(L"Find window did not expose the expected baseline results-grid state before header reorder validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;
    state.Require(DebugReorderFindFilesWindowVisibleResultColumn(1u, 0u),
                  L"Failed to reorder the visible Path column ahead of Name through the Find results grid debug seam.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(10000ms),
                      &snapshot),
                  L"Find header drag should reorder visible columns without losing selection or bounded visible work.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogCopyFollowsReorderedColumns(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto restoreSearch                                                   = wil::scope_exit([&] noexcept { g_settings.search = previousSearch; });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_copy_reordered_columns";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find reordered-copy test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find reordered-copy subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find reordered-copy.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find reordered-copy.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-copy-reordered-columns");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reordered-copy validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reordered-copy validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reordered-copy validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reordered-copy validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reordered-copy validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reordered-copy validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath         = (root / L"sub" / L"beta.txt").native();
    const std::wstring expectedRelativePath = L"sub";
    const std::wstring expectedName         = L"beta.txt";
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for reordered-copy validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the expected baseline results-grid state before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const D2D1_RECT_F headerRect = snapshot.secondResultHeaderRect;
    const LPARAM dragStart =
        DipPointToEffectiveMouseInputLParam(findWindow, (headerRect.left + headerRect.right) * 0.5f, (headerRect.top + headerRect.bottom) * 0.5f);
    const LPARAM dragTarget =
        DipPointToEffectiveMouseInputLParam(findWindow, snapshot.firstResultHeaderRect.left + 12.0f, (headerRect.top + headerRect.bottom) * 0.5f);

    SendMouseDragToWindow(findWindow, dragStart, dragTarget);

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not settle on the reordered visible column order before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the Find results grid before reordered-copy validation.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid; },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Find results grid did not take focus before reordered-copy validation.");

    ClearClipboardContents(findWindow);
    SendMessageW(findWindow, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(findWindow, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(findWindow, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(findWindow, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(findWindow);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(), L"Find Ctrl+C should copy the reordered visible row content to the clipboard.");
    state.Require(copiedSelection.rfind((expectedRelativePath + L"\t"), 0u) == 0u,
                  L"Find clipboard copy should start with the visible Path column after header reorder.");
    state.Require(copiedSelection.find(expectedName) != std::wstring::npos,
                  L"Find clipboard copy should still include the selected result name after header reorder.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogReorderedColumnsSurviveSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto restoreSearch                                                   = wil::scope_exit([&] noexcept { g_settings.search = previousSearch; });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_reordered_columns_survive_sort_cycles";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find reorder/sort test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find reorder/sort subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find reorder/sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find reorder/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-reordered-sort-cycles");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reorder/sort validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reorder/sort validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reorder/sort validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reorder/sort validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reorder/sort validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reorder/sort validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    const std::wstring otherPath    = (root / L"alpha.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for reorder/sort validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath) && containsPath(value, otherPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before reorder/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;

    state.Require(DebugReorderFindFilesWindowVisibleResultColumn(1u, 0u),
                  L"Failed to reorder the visible Path column ahead of Name through the deterministic debug layout path before sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not settle on the reordered visible column order before sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX = static_cast<LONG>(std::lround((snapshot.firstResultHeaderRect.left + snapshot.firstResultHeaderRect.right) * 0.5f));
    const LONG clickY = static_cast<LONG>(std::lround((snapshot.firstResultHeaderRect.top + snapshot.firstResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount &&
               value.dxResizeFailureCount == 0u && value.fullPaths.size() >= 2u && value.fullPaths[0] == selectedPath && value.fullPaths[1] == otherPath &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find reordered-column sort cycles should keep the visible display order while sorting the visible Path column.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogHeaderResizeChangesVisibleWidth(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto restoreSearch                                                   = wil::scope_exit([&] noexcept { g_settings.search = previousSearch; });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_header_resize";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find header-resize test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find header-resize subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find header resize.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find header resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-header-resize");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for header resize validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for header resize validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for header resize validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for header resize validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for header resize validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for header resize validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for header resize validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the expected baseline results-grid state before header resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;
    const float baselineFirstHeaderWidth    = std::max(0.0f, snapshot.firstResultHeaderRect.right - snapshot.firstResultHeaderRect.left);
    const float baselineFirstColumnWidthDip = snapshot.resultColumnWidthsDip[0];
    const float baselineSecondHeaderLeft    = snapshot.secondResultHeaderRect.left;

    const auto waitForHeaderResize = [&]() noexcept
    {
        return WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
        {
            const float newFirstHeaderWidth = std::max(0.0f, value.firstResultHeaderRect.right - value.firstResultHeaderRect.left);
            return value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
                   value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
                   newFirstHeaderWidth >= baselineFirstHeaderWidth + 20.0f && value.resultColumnWidthsDip[0] >= baselineFirstColumnWidthDip + 20.0f &&
                   value.secondResultHeaderRect.left > baselineSecondHeaderLeft + 10.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
                   value.visibleResultColumnCount > 0u && value.visibleResultColumnCount <= baselineVisibleColumnCount &&
                   value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount && value.visibleResultCellCount > 0u &&
                   value.visibleResultCellCount <= baselineVisibleCellCount &&
                   value.visibleResultCellCount <= value.visibleResultRowCount * value.visibleResultColumnCount && value.dxResizeFailureCount == 0u &&
                   containsPath(value, selectedPath);
        },
            SelfTest::Scale(3000ms),
            &snapshot);
    };

    state.Require(SendFindFirstVisibleHeaderResizeDrag(findWindow, snapshot, 48.0f),
                  L"Failed to locate the live Find header-resize grip before header resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! waitForHeaderResize())
    {
        state.Require(SendFindFirstVisibleHeaderResizeDragToHost(findWindow, snapshot, 48.0f),
                      L"Failed to retry the live Find header-resize drag against the host window.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const float currentFirstHeaderWidth = std::max(0.0f, snapshot.firstResultHeaderRect.right - snapshot.firstResultHeaderRect.left);
    const float currentSecondHeaderLeft = snapshot.secondResultHeaderRect.left;
    state.Require(waitForHeaderResize(),
                  std::format(L"Find header resize should widen the visible Name column without reordering columns or losing the selected row. "
                              L"baselineFirstHeaderW={:.1f} currentFirstHeaderW={:.1f} baselineFirstColumnW={:.1f} currentFirstColumnW={:.1f} "
                              L"baselineSecondHeaderLeft={:.1f} currentSecondHeaderLeft={:.1f} baselineVisibleRows={} currentVisibleRows={} "
                              L"baselineVisibleCols={} currentVisibleCols={} baselineVisibleCells={} currentVisibleCells={}. {}",
                              baselineFirstHeaderWidth,
                              currentFirstHeaderWidth,
                              baselineFirstColumnWidthDip,
                              snapshot.resultColumnWidthsDip.size() >= 1u ? snapshot.resultColumnWidthsDip[0] : 0.0f,
                              baselineSecondHeaderLeft,
                              currentSecondHeaderLeft,
                              baselineVisibleRowCount,
                              snapshot.visibleResultRowCount,
                              baselineVisibleColumnCount,
                              snapshot.visibleResultColumnCount,
                              baselineVisibleCellCount,
                              snapshot.visibleResultCellCount,
                              DescribeFindSnapshotBrief(snapshot)));

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogReorderedColumnsSurviveSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_reordered_columns_survive_search_rerun";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find reorder/rerun test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find reorder/rerun subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find reorder/rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find reorder/rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-reordered-search-rerun");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reorder/rerun validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reorder/rerun validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reorder/rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reorder/rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reorder/rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reorder/rerun validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the baseline results-grid state before reorder/rerun validation. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;

    state.Require(DebugReorderFindFilesWindowVisibleResultColumn(1u, 0u),
                  L"Failed to reorder the visible Path column ahead of Name through the deterministic debug layout path before search rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find header drag did not settle on the reordered visible column order before search rerun validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameCombo),
                  L"Failed to focus the Find name combo before reordered-layout rerun validation.");
    state.Require(DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget::NameCombo, L"beta*"),
                  L"Failed to narrow the Find name combo text for reordered-layout rerun validation.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::NameCombo && value.namePatternText == L"beta*" && value.rootText == root.native() &&
               value.contentPatternText.empty() && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the narrowed beta* query before reordered-layout rerun validation. {}", DescribeFindSnapshotBrief(snapshot)));
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to rerun Find search for reordered-layout rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Narrowed Find search did not become idle for reordered-layout rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.visibleResultColumnCount == baselineVisibleColumnCount && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find reordered visible column order should survive a narrowed search rerun. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameCombo),
                  L"Failed to refocus the Find name combo before restoring the reordered-layout baseline search.");
    state.Require(DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget::NameCombo, L"*.txt"),
                  L"Failed to restore the Find baseline name combo text after reordered-layout rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after reordered-layout rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after reordered-layout rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find reordered visible column order should restore after returning to the baseline search rerun.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogResizedColumnsSurviveSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_resized_columns_survive_search_rerun";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find resize/rerun test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find resize/rerun subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find resize/rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find resize/rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-resized-search-rerun");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for resize/rerun validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for resize/rerun validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for resize/rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for resize/rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for resize/rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for resize/rerun validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the baseline results-grid state before resize/rerun validation. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;
    const float baselineFirstColumnWidthDip = snapshot.resultColumnWidthsDip[0];
    const auto waitForHeaderResize          = [&]() noexcept
    {
        return WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
                   value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
                   value.resultColumnWidthsDip[0] >= baselineFirstColumnWidthDip + 20.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
                   value.visibleResultColumnCount > 0u && value.visibleResultColumnCount <= baselineVisibleColumnCount &&
                   value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount && value.visibleResultCellCount > 0u &&
                   value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
        },
            SelfTest::Scale(3000ms),
            &snapshot);
    };

    state.Require(SendFindFirstVisibleHeaderResizeDrag(findWindow, snapshot, 48.0f),
                  L"Failed to locate the live Find header-resize grip before resize/rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! waitForHeaderResize())
    {
        state.Require(SendFindFirstVisibleHeaderResizeDragToHost(findWindow, snapshot, 48.0f),
                      L"Failed to retry the live Find header-resize drag against the host window before resize/rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(
        waitForHeaderResize(),
        std::format(L"Find deterministic resize/rerun validation did not settle on the widened Name column. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const float widenedWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameCombo),
                  L"Failed to focus the Find name combo before resized-layout rerun validation.");
    state.Require(DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget::NameCombo, L"beta*"),
                  L"Failed to narrow the Find name combo text for resized-layout rerun validation.");
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot) && snapshot.namePatternText == L"beta*",
                  std::format(L"Find window did not accept the narrowed beta* query immediately before resized-layout rerun validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::NameCombo && value.namePatternText == L"beta*" && value.rootText == root.native() &&
               value.contentPatternText.empty() && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnWidthsDip[0] >= widenedWidthDip - 1.0f &&
               value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the narrowed beta* query before resized-layout rerun validation. {}", DescribeFindSnapshotBrief(snapshot)));
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to rerun Find search for resized-layout rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Narrowed Find search did not become idle for resized-layout rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
               value.resultColumnWidthsDip[0] >= widenedWidthDip - 1.0f && value.visibleResultColumnCount > 0u &&
               value.visibleResultColumnCount <= baselineVisibleColumnCount && value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find widened Name column should survive a narrowed search rerun. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameCombo),
                  L"Failed to refocus the Find name combo before restoring the resized-layout baseline search.");
    state.Require(DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget::NameCombo, L"*.txt"),
                  L"Failed to restore the Find baseline name combo text after resized-layout rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after resized-layout rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after resized-layout rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
               value.resultColumnWidthsDip[0] >= widenedWidthDip - 1.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount > 0u && value.visibleResultColumnCount <= baselineVisibleColumnCount &&
               value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount && value.visibleResultCellCount > 0u &&
               value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find widened Name column should restore after returning to the baseline search rerun.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogReorderedResizedColumnsSurviveSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_reordered_resized_columns_survive_search_rerun";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find reorder+resize/rerun test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find reorder+resize/rerun subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find reorder+resize/rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for Find reorder+resize/rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-reordered-resized-search-rerun");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reorder+resize/rerun validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reorder+resize/rerun validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reorder+resize/rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reorder+resize/rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reorder+resize/rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reorder+resize/rerun validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before reorder+resize/rerun validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;
    state.Require(DebugReorderFindFilesWindowVisibleResultColumn(1u, 0u),
                  L"Failed to reorder the visible Path column ahead of Name before combined-view-state persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not settle on the reordered visible column order before reorder+resize/rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        std::format(L"Find reordered first visible column did not widen before reorder+resize/rerun validation. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const float widenedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to rerun Find search for reorder+resize/rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reorder+resize/rerun validation.");

    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip[0] >= widenedFirstVisibleWidthDip - 1.0f && value.visibleResultColumnCount > 0u &&
               value.visibleResultColumnCount <= baselineVisibleColumnCount && value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
            SelfTest::Scale(10000ms),
            &snapshot),
        std::format(L"Find reordered visible layout and widened first column should survive a search rerun together. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameCombo),
                  L"Failed to refocus the Find name combo before restoring the reorder+resize baseline search.");
    state.Require(DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget::NameCombo, L"*.txt"),
                  L"Failed to restore the Find baseline name combo text after reorder+resize/rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after reorder+resize/rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after reorder+resize/rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip[0] >= widenedFirstVisibleWidthDip - 1.0f && value.visibleResultColumnCount > 0u &&
               value.visibleResultColumnCount <= baselineVisibleColumnCount && value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount &&
               value.visibleResultCellCount > 0u && value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find reordered visible layout and widened first column should restore after returning to the baseline search rerun.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogReorderedResizedColumnsSurviveSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_reordered_resized_columns_survive_sort_cycles";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find reorder+resize/sort test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find reorder+resize/sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find reorder+resize/sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for Find reorder+resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-reordered-resized-sort-cycles");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reorder+resize/sort validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reorder+resize/sort validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reorder+resize/sort validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reorder+resize/sort validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reorder+resize/sort validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reorder+resize/sort validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for reorder+resize/sort validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 3u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before reorder+resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;

    state.Require(
        DebugReorderFindFilesWindowVisibleResultColumn(1u, 0u),
        L"Failed to reorder the visible Path column ahead of Name through the deterministic debug layout path before reorder+resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 3u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not settle on the reordered visible column order before reorder+resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic reorder+resize/sort validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float widenedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];
    const LONG sortClickX = static_cast<LONG>(std::lround((snapshot.secondResultHeaderRect.left + snapshot.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((snapshot.secondResultHeaderRect.top + snapshot.secondResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    const std::wstring firstExpected  = (root / L"gamma.txt").native();
    const std::wstring secondExpected = (root / L"beta.txt").native();
    const std::wstring thirdExpected  = (root / L"alpha.txt").native();
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 3u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip[0] >= widenedFirstVisibleWidthDip - 1.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount > 0u && value.visibleResultColumnCount <= baselineVisibleColumnCount &&
               value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount && value.visibleResultCellCount > 0u &&
               value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u && value.fullPaths.size() >= 3u &&
               value.fullPaths[0] == firstExpected && value.fullPaths[1] == secondExpected && value.fullPaths[2] == thirdExpected &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find sort cycles should keep the reordered visible layout and widened first visible column intact together. {}",
                              DescribeFindSnapshotBrief(snapshot)));

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogResizedColumnsSurviveSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_resized_columns_survive_sort_cycles";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find resize/sort test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find resize/sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find resize/sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for Find resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-resized-sort-cycles");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for resize/sort validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for resize/sort validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for resize/sort validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for resize/sort validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for resize/sort validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for resize/sort validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for resize/sort validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 3u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the baseline results-grid state before resize/sort validation. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;
    const float baselineFirstColumnWidthDip = snapshot.resultColumnWidthsDip[0];
    const auto waitForResizedNameColumn     = [&]() noexcept
    {
        return WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" &&
                   value.resultColumnIds[1] == L"path" && value.resultColumnWidthsDip[0] >= baselineFirstColumnWidthDip + 20.0f &&
                   value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount > 0u &&
                   value.visibleResultColumnCount <= baselineVisibleColumnCount && value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount &&
                   value.visibleResultCellCount > 0u && value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u &&
                   containsPath(value, selectedPath);
        },
            SelfTest::Scale(3000ms),
            &snapshot);
    };

    state.Require(SendFindFirstVisibleHeaderResizeDrag(findWindow, snapshot, 48.0f),
                  L"Failed to locate the live Find header-resize grip before reorder+resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! waitForResizedNameColumn())
    {
        state.Require(SendFindFirstVisibleHeaderResizeDragToHost(findWindow, snapshot, 48.0f),
                      L"Failed to retry the live Find header-resize drag against the host window before reorder+resize/sort validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(
        waitForResizedNameColumn(),
        std::format(L"Find header resize did not settle on the widened Name column before resize/sort validation. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const float widenedWidthDip = snapshot.resultColumnWidthsDip[0];
    const LONG clickX           = static_cast<LONG>(std::lround((snapshot.firstResultHeaderRect.left + snapshot.firstResultHeaderRect.right) * 0.5f));
    const LONG clickY           = static_cast<LONG>(std::lround((snapshot.firstResultHeaderRect.top + snapshot.firstResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));

    const std::wstring firstExpected  = (root / L"gamma.txt").native();
    const std::wstring secondExpected = (root / L"beta.txt").native();
    const std::wstring thirdExpected  = (root / L"alpha.txt").native();
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 3u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
               value.resultColumnWidthsDip[0] >= widenedWidthDip - 1.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount > 0u && value.visibleResultColumnCount <= baselineVisibleColumnCount &&
               value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount && value.visibleResultCellCount > 0u &&
               value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u && value.fullPaths.size() >= 3u &&
               value.fullPaths[0] == firstExpected && value.fullPaths[1] == secondExpected && value.fullPaths[2] == thirdExpected &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find sort cycles should keep the widened Name column intact while sorting the visible Name header.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogHeaderClickSortsResults(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_header_click_sort";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find header-click-sort test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find header-click sort.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find header-click sort.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for Find header-click sort.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    std::optional<std::filesystem::path> leftBefore;
    const auto restorePath = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    HWND findWindow = nullptr;
    state.Require(OpenFindWindowFromLocalPaneRoot(mainWindow, root, {L"alpha.txt", L"beta.txt", L"gamma.txt"}, findWindow, leftBefore),
                  L"Find window did not open from a deterministic local pane root for header-click sort validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for header-click sort validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for header-click sort validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for header-click sort validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for header-click sort validation.");

    const std::wstring selectedPath = (root / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for header-click sort validation.", selectedPath));

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot before header-click sort validation.");
    state.Require(snapshot.usesDxUiHost && snapshot.resultCount == 3u && snapshot.selectedResultCount == 1u && snapshot.visibleChildWindowCount <= 1u &&
                      snapshot.resultColumnIds.size() >= 2u && snapshot.firstResultHeaderRect.right > snapshot.firstResultHeaderRect.left &&
                      snapshot.resultColumnIds[0] == L"name" && snapshot.resultColumnIds[1] == L"path" && snapshot.dxResizeFailureCount == 0u,
                  std::format(L"Find window did not expose the baseline results-grid state before header-click sort validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX = static_cast<LONG>(std::lround((snapshot.firstResultHeaderRect.left + snapshot.firstResultHeaderRect.right) * 0.5f));
    const LONG clickY = static_cast<LONG>(std::lround((snapshot.firstResultHeaderRect.top + snapshot.firstResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
               value.dxResizeFailureCount == 0u && value.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), value.fullPaths.begin());
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header clicks did not sort the visible results into descending Name order while preserving selection and visible layout.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresResizedGridLayout(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restores_resized_grid_layout";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create resized-layout test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create resized-layout subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for resized-layout validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for resized-layout validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    std::optional<std::filesystem::path> leftBefore;
    const auto restorePath = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for resized-layout validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for resized-layout validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-resized-layout-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for resized-layout validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for resized-layout validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for resized-layout validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for resized-layout validation.");

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for resized-layout validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before resized-layout persistence validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstHeaderWidth    = std::max(0.0f, snapshot.firstResultHeaderRect.right - snapshot.firstResultHeaderRect.left);
    const float baselineFirstColumnWidthDip = snapshot.resultColumnWidthsDip[0];
    state.Require(SendFindFirstVisibleHeaderResizeDrag(findWindow, snapshot, 48.0f),
                  L"Failed to locate the live Find header-resize grip before resized-layout persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        const float newFirstHeaderWidth = std::max(0.0f, value.firstResultHeaderRect.right - value.firstResultHeaderRect.left);
        return value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && newFirstHeaderWidth >= baselineFirstHeaderWidth + 20.0f &&
               value.resultColumnWidthsDip[0] >= baselineFirstColumnWidthDip + 20.0f && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header resize did not widen the visible Name column before persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedWidthDip = snapshot.resultColumnWidthsDip[0];

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after a live header resize.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    const auto& savedLayout = g_settings.search->resultsGridLayout;
    state.Require(savedLayout.size() >= 2u, L"Closing the Find window should persist the resized grid layout.");
    if (savedLayout.size() >= 2u)
    {
        state.Require(savedLayout[0].columnId == L"name", L"Persisted resized Find layout should keep Name first after close.");
        state.Require(savedLayout[1].columnId == L"path", L"Persisted resized Find layout should keep Path second after close.");
        state.Require(std::fabs(savedLayout[0].widthDip - resizedWidthDip) <= 1.0f,
                      L"Persisted resized Find layout should keep the widened Name column width after close.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-resized-layout-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && std::fabs(value.resultColumnWidthsDip[0] - resizedWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the persisted resized Name column width.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresReorderedGridLayout(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restores_reordered_grid_layout";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create reordered-layout test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create reordered-layout subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for reordered-layout validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for reordered-layout validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reordered-layout validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reordered-layout validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-reordered-layout-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reordered-layout validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reordered-layout validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reordered-layout validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reordered-layout validation.");

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for reordered-layout validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before reordered-layout persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SendFindSecondVisibleHeaderAheadOfFirst(findWindow, snapshot),
                  L"Failed to drag the live Find Path header ahead of Name before persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not reorder the live results-grid columns before persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after a live header reorder.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    const auto& savedLayout = g_settings.search->resultsGridLayout;
    state.Require(savedLayout.size() >= 2u, L"Closing the Find window should persist the reordered grid layout.");
    if (savedLayout.size() >= 2u)
    {
        state.Require(savedLayout[0].columnId == L"path", L"Persisted reordered Find layout should keep Path first after close.");
        state.Require(savedLayout[1].columnId == L"name", L"Persisted reordered Find layout should keep Name second after close.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-reordered-layout-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the persisted reordered column layout.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredReorderedLayoutCopyFollowsVisibleColumns(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_reordered_copy";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored-reordered-copy test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored-reordered-copy subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for restored-reordered-copy validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for restored-reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored-reordered-copy validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for restored-reordered-copy validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-reordered-copy-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored-reordered-copy validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for restored-reordered-copy validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for restored-reordered-copy validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored-reordered-copy validation.");

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for restored-reordered-copy validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored-reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SendFindSecondVisibleHeaderAheadOfFirst(findWindow, snapshot),
                  L"Failed to drag the live Find Path header ahead of Name before restored-reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not reorder the live results-grid columns before restored-reordered-copy persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after reordered-copy setup.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    state.Require(g_settings.search->resultsGridLayout.size() >= 2u,
                  L"Closing the Find window should persist the reordered layout before restored-copy validation.");
    if (g_settings.search->resultsGridLayout.size() >= 2u)
    {
        state.Require(g_settings.search->resultsGridLayout[0].columnId == L"path",
                      L"Persisted reordered Find layout should keep Path first before restored-copy validation.");
        state.Require(g_settings.search->resultsGridLayout[1].columnId == L"name",
                      L"Persisted reordered Find layout should keep Name second before restored-copy validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-reordered-copy-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the persisted reordered column layout before restored-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to reconfigure Find window after restored reordered-layout reopen.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to restore Find options after restored reordered-layout reopen.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored reordered-layout copy validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Restarted Find search did not become idle for restored reordered-layout copy validation.");
    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored reordered-layout reopen.", selectedPath));

    const std::wstring expectedRelativePath = L"sub";
    const std::wstring expectedName         = L"beta.txt";
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected reordered results-grid state before restored-copy validation.");
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the reopened Find results grid before restored reordered-copy validation.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid; },
                                      SelfTest::Scale(3000ms),
                                      &reopened),
                  L"Reopened Find results grid did not take focus before restored reordered-copy validation.");

    ClearClipboardContents(findWindow);
    SendMessageW(findWindow, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(findWindow, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(findWindow, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(findWindow, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(findWindow);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(), L"Find Ctrl+C should copy the restored reordered visible row content to the clipboard after reopen.");
    state.Require(copiedSelection.rfind((expectedRelativePath + L"\t"), 0u) == 0u,
                  L"Reopened Find clipboard copy should start with the restored visible Path column after header reorder persistence.");
    state.Require(copiedSelection.find(expectedName) != std::wstring::npos,
                  L"Reopened Find clipboard copy should still include the selected result name after restored header reorder.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresPersistedSortOrder(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restores_persisted_sort_order";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create persisted-sort test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for persisted-sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for persisted-sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for persisted-sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode          = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode       = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId      = L"name";
    search.sortDescending    = true;
    search.resultsGridLayout = {
        Common::Settings::GridColumnLayoutEntry{.columnId = L"path", .displayIndex = 0u, .widthDip = 320.0f},
        Common::Settings::GridColumnLayoutEntry{.columnId = L"name", .displayIndex = 1u, .widthDip = 240.0f},
    };
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for persisted-sort validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for persisted-sort validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-sort-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowResultSort(0u, true), L"Failed to enable descending Name sort for persisted-sort validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for persisted-sort validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for persisted-sort validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value); },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  std::format(L"Find window did not expose the expected descending Name sort order before persistence validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after a live sort change.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    state.Require(g_settings.search->sortColumnId == L"name", L"Closing the Find window should persist the logical Name sort column id.");
    state.Require(g_settings.search->sortDescending, L"Closing the Find window should persist descending sort direction.");
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-sort-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to restart Find search for persisted-sort restore validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for persisted-sort restore validation.");

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value); },
                                      SelfTest::Scale(3000ms),
                                      &reopened),
                  L"Reopened Find window did not restore the persisted descending Name sort order.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresReorderedSortedGridLayout(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restores_reordered_sorted_grid_layout";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create reordered-sorted-layout test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for reordered-sorted-layout validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for reordered-sorted-layout validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for reordered-sorted-layout validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for reordered-sorted-layout validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reordered-sorted-layout validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-reordered-sorted-layout-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reordered-sorted-layout validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reordered-sorted-layout validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before reordered-sorted-layout validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  std::format(L"Find deterministic header reorder did not move the visible Path column ahead of Name before combined restore validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowResultSort(0u, true), L"Failed to enable descending logical Name sort for reordered-sorted-layout validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort did not stay correct after visible header reorder.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    state.Require(g_settings.search->resultsGridLayout.size() >= 2u, L"Closing the Find window should persist the reordered visible layout.");
    if (g_settings.search->resultsGridLayout.size() >= 2u)
    {
        state.Require(g_settings.search->resultsGridLayout[0].columnId == L"path", L"Persisted combined Find layout should keep Path first after close.");
        state.Require(g_settings.search->resultsGridLayout[1].columnId == L"name", L"Persisted combined Find layout should keep Name second after close.");
    }
    state.Require(g_settings.search->sortColumnId == L"name", L"Persisted combined Find sort should remain bound to the logical Name column id.");
    state.Require(g_settings.search->sortDescending, L"Persisted combined Find sort should remain descending after close.");
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-reordered-sorted-layout-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for reordered-sorted-layout restore validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for reordered-sorted-layout restore validation.");

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered layout plus logical Name sort.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresCombinedViewState(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restores_combined_view_state";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create combined-view-state test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for combined-view-state validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for combined-view-state validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for combined-view-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    std::optional<std::filesystem::path> leftBefore;
    const auto restorePath = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* /*themeTag*/) noexcept -> HWND
    {
        HWND findWindow = nullptr;
        state.Require(OpenFindWindowFromLocalPaneRoot(mainWindow, root, {L"alpha.txt", L"beta.txt", L"gamma.txt"}, findWindow, leftBefore),
                      L"Find window did not open from a deterministic local pane root for combined-view-state validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-combined-view-state-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot initializationSnapshot{};
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return value.usesDxUiHost && value.findButtonEnabled && value.rootComboEnabled && value.nameComboEnabled && value.nameModeComboEnabled; },
                                      SelfTest::Scale(3000ms),
                                      &initializationSnapshot),
                  std::format(L"Find window did not finish initializing its DX shell before combined-view-state validation. {}",
                              DescribeFindSnapshotBrief(initializationSnapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for combined-view-state validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for combined-view-state validation.");
    FindFilesDebugSnapshot configuredSnapshot{};
    state.Require(
        WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return value.usesDxUiHost && value.findButtonEnabled && value.rootComboEnabled && value.nameComboEnabled && value.nameModeComboEnabled; },
                            SelfTest::Scale(1000ms),
                            &configuredSnapshot),
        std::format(L"Find window lost its live DX shell after combined-view-state configuration. {}", DescribeFindSnapshotBrief(configuredSnapshot)));
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for combined-view-state validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for combined-view-state validation.");
    FindFilesDebugSnapshot postSearchSnapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(postSearchSnapshot), L"Failed to capture Find snapshot immediately after combined-view-state search.");
    state.Require(postSearchSnapshot.usesDxUiHost,
                  std::format(L"Find window lost its DX host immediately after combined-view-state search. {}", DescribeFindSnapshotBrief(postSearchSnapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPath = (root / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to select '{}' before combined-view-state baseline validation.", selectedPath));
    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot before combined-view-state validation.");
    state.Require(snapshot.usesDxUiHost && snapshot.visibleChildWindowCount <= 1u && snapshot.resultColumnIds.size() >= 2u &&
                      snapshot.resultColumnWidthsDip.size() >= 2u && snapshot.firstResultHeaderRect.right > snapshot.firstResultHeaderRect.left &&
                      snapshot.secondResultHeaderRect.right > snapshot.secondResultHeaderRect.left && snapshot.resultColumnIds[0] == L"name" &&
                      snapshot.resultColumnIds[1] == L"path" && snapshot.selectedResultCount == 1u &&
                      std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), selectedPath) != snapshot.fullPaths.end() &&
                      snapshot.dxResizeFailureCount == 0u,
                  std::format(L"Find window did not expose the baseline results-grid state before combined-view-state validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Find combined-view-state validation did not reorder the visible Path column ahead of Name through the deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find combined-view-state validation did not widen the visible reordered Path column through the deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true), L"Failed to enable descending logical Name sort for combined-view-state validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not stay correct before combined-view-state persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    state.Require(g_settings.search->resultsGridLayout.size() >= 2u, L"Closing the Find window should persist the combined visible grid layout.");
    if (g_settings.search->resultsGridLayout.size() >= 2u)
    {
        state.Require(g_settings.search->resultsGridLayout[0].columnId == L"path", L"Persisted combined Find layout should keep Path first after close.");
        state.Require(g_settings.search->resultsGridLayout[1].columnId == L"name", L"Persisted combined Find layout should keep Name second after close.");
        state.Require(std::fabs(g_settings.search->resultsGridLayout[0].widthDip - resizedFirstVisibleWidthDip) <= 1.0f,
                      L"Persisted combined Find layout should keep the widened first visible column width after close.");
    }
    state.Require(g_settings.search->sortColumnId == L"name", L"Persisted combined Find sort should remain bound to the logical Name column id.");
    state.Require(g_settings.search->sortDescending, L"Persisted combined Find sort should remain descending after close.");
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-combined-view-state-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for combined-view-state validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to restore Find options for combined-view-state validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to restart Find search for combined-view-state restore validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for combined-view-state restore validation.");

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered layout, resized width, and logical Name sort.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumns(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_copy";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored combined-view-state copy test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored combined-view-state copy subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"),
                  L"Failed to create alpha.txt for restored combined-view-state copy validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for restored combined-view-state copy validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"),
                  L"Failed to create gamma.txt for restored combined-view-state copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"sub" / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state copy validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for restored combined-view-state copy validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state copy validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to configure Find options for restored combined-view-state copy validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state copy validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state copy validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored combined-view-state copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Find restored combined-view-state copy validation did not reorder the visible Path column ahead of Name through the deterministic debug "
                  L"layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        L"Find restored combined-view-state copy validation did not widen the visible reordered Path column through the deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state copy validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state copy persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    state.Require(g_settings.search->resultsGridLayout.size() >= 2u,
                  L"Closing the Find window should persist the combined visible grid layout before restored combined copy validation.");
    if (g_settings.search->resultsGridLayout.size() >= 2u)
    {
        state.Require(g_settings.search->resultsGridLayout[0].columnId == L"path",
                      L"Persisted combined Find layout should keep Path first before restored combined copy validation.");
        state.Require(g_settings.search->resultsGridLayout[1].columnId == L"name",
                      L"Persisted combined Find layout should keep Name second before restored combined copy validation.");
        state.Require(std::fabs(g_settings.search->resultsGridLayout[0].widthDip - resizedFirstVisibleWidthDip) <= 1.0f,
                      L"Persisted combined Find layout should keep the widened first visible column width before restored combined copy validation.");
    }
    state.Require(g_settings.search->sortColumnId == L"name",
                  L"Persisted combined Find sort should remain bound to the logical Name column id before restored combined copy validation.");
    state.Require(g_settings.search->sortDescending, L"Persisted combined Find sort should remain descending before restored combined copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for restored combined-view-state copy validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to restore Find options for restored combined-view-state copy validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state copy validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state copy validation.");

    const std::wstring selectedPath         = (root / L"sub" / L"beta.txt").native();
    const std::wstring expectedRelativePath = L"sub";
    const std::wstring expectedName         = L"beta.txt";

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before restored combined copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before restored combined copy validation.");
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the reopened Find results grid before restored combined copy validation.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid; },
                                      SelfTest::Scale(3000ms),
                                      &reopened),
                  L"Reopened Find results grid did not take focus before restored combined copy validation.");

    ClearClipboardContents(findWindow);
    SendMessageW(findWindow, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(findWindow, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(findWindow, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(findWindow, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(findWindow);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(), L"Find Ctrl+C should copy the restored combined visible row content to the clipboard after reopen.");
    state.Require(copiedSelection.rfind((expectedRelativePath + L"\t"), 0u) == 0u,
                  L"Reopened Find clipboard copy should start with the restored visible Path column after combined view-state persistence.");
    state.Require(copiedSelection.find(expectedName) != std::wstring::npos,
                  L"Reopened Find clipboard copy should still include the selected result name after restored combined view-state persistence.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateSurvivesSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_search_rerun";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored combined-view-state rerun test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored combined-view-state rerun subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"),
                  L"Failed to create alpha.txt for restored combined-view-state rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for restored combined-view-state rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"),
                  L"Failed to create gamma.txt for restored combined-view-state rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"sub" / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state rerun validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for restored combined-view-state rerun validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-rerun-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to configure Find options for restored combined-view-state rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state rerun validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored combined-view-state rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Find restored combined-view-state rerun validation did not reorder the visible Path column ahead of Name through the deterministic debug "
                  L"layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        L"Find restored combined-view-state rerun validation did not widen the visible reordered Path column through the deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state rerun validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state rerun persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-rerun-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for restored combined-view-state rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to restore Find options for restored combined-view-state rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state rerun validation.");

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before rerun validation.");

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state reopen.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state reopen.");

    const auto restoredCombinedRerunStateOk = [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.submittedNamePatternText == L"*.txt" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    };
    state.Require(DebugGetFindFilesWindowSnapshot(reopened), L"Failed to capture Find snapshot after restored combined-view-state baseline rerun.");
    state.Require(
        restoredCombinedRerunStateOk(reopened),
        std::format(L"Reopened combined Find view-state should survive a baseline search rerun. expectedWidth={:.1f} observedWidth={:.1f} "
                    L"widthDelta={:.1f} submittedNameOk={} submittedNameLen={} containsSelected={} expectedOrder={} selectedPath='{}' Last snapshot: {}",
                    resizedFirstVisibleWidthDip,
                    reopened.resultColumnWidthsDip.empty() ? 0.0f : reopened.resultColumnWidthsDip[0],
                    reopened.resultColumnWidthsDip.empty() ? resizedFirstVisibleWidthDip
                                                           : std::fabs(reopened.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip),
                    reopened.submittedNamePatternText == L"*.txt" ? 1 : 0,
                    reopened.submittedNamePatternText.size(),
                    containsPath(reopened, selectedPath) ? 1 : 0,
                    matchesExpectedOrder(reopened) ? 1 : 0,
                    selectedPath,
                    DescribeFindSnapshotBrief(reopened)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore baseline Find search after restored combined-view-state rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should restore after returning to the baseline search rerun.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateSurvivesSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_sort_cycles";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored combined-view-state sort-cycles test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored combined-view-state sort-cycles subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"),
                  L"Failed to create alpha.txt for restored combined-view-state sort-cycles validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for restored combined-view-state sort-cycles validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"),
                  L"Failed to create gamma.txt for restored combined-view-state sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"sub" / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedDescendingOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state sort-cycles validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state sort-cycles validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-sort-cycles-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state sort-cycles validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to configure Find options for restored combined-view-state sort-cycles validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state sort-cycles validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state sort-cycles validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored combined-view-state sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Failed to apply the deterministic Find header reorder before restored combined-view-state sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state sort-cycles validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state sort-cycles validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedDescendingOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state sort-cycles persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-sort-cycles-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for restored combined-view-state sort-cycles validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to restore Find options for restored combined-view-state sort-cycles validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state sort-cycles validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state sort-cycles validation.");

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedDescendingOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    const LONG sortClickX = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.left + reopened.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.top + reopened.secondResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f &&
               value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && matchesExpectedDescendingOrder(value) &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should survive repeated sort cycles.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumnsAfterSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_copy_after_sort_cycles";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored combined-view-state copy-after-sort-cycles test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored combined-view-state copy-after-sort-cycles subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"),
                  L"Failed to create alpha.txt for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"),
                  L"Failed to create gamma.txt for restored combined-view-state copy-after-sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"sub" / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state copy-after-sort-cycles validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state copy-after-sort-cycles validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-after-sort-cycles-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to configure Find options for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state copy-after-sort-cycles validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored combined-view-state copy-after-sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Find restored combined-view-state copy-after-sort-cycles validation did not reorder the visible Path column ahead of Name through the "
                  L"deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find restored combined-view-state copy-after-sort-cycles validation did not widen the visible reordered Path column through the "
                  L"deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state copy-after-sort-cycles persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-after-sort-cycles-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to restore Find options for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state copy-after-sort-cycles validation.");

    const std::wstring selectedPath         = (root / L"sub" / L"beta.txt").native();
    const std::wstring expectedRelativePath = L"sub";
    const std::wstring expectedName         = L"beta.txt";

    FindFilesDebugSnapshot reopened{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not restore the combined reordered, resized, and sorted state before restored combined copy-after-sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not expose the expected selected combined results-grid state before restored combined copy-after-sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    const LONG sortClickX = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.left + reopened.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.top + reopened.secondResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f &&
               value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should survive repeated sort cycles before restored combined copy-after-sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the reopened Find results grid before restored combined copy-after-sort-cycles validation.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid; },
                                      SelfTest::Scale(3000ms),
                                      &reopened),
                  L"Reopened Find results grid did not take focus before restored combined copy-after-sort-cycles validation.");

    ClearClipboardContents(findWindow);
    SendMessageW(findWindow, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(findWindow, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(findWindow, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(findWindow, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(findWindow);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(),
                  L"Find Ctrl+C should copy the restored combined visible row content to the clipboard after reopen and sort cycles.");
    state.Require(copiedSelection.rfind((expectedRelativePath + L"\t"), 0u) == 0u,
                  L"Reopened Find clipboard copy should still start with the restored visible Path column after combined view-state sort churn.");
    state.Require(copiedSelection.find(expectedName) != std::wstring::npos,
                  L"Reopened Find clipboard copy should still include the selected result name after combined view-state sort churn.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumnsAfterSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_copy_after_search_rerun";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored combined-view-state copy-after-search-rerun test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored combined-view-state copy-after-search-rerun subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"),
                  L"Failed to create alpha.txt for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"),
                  L"Failed to create gamma.txt for restored combined-view-state copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"sub" / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state copy-after-search-rerun validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state copy-after-search-rerun validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-after-search-rerun-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to configure Find options for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state copy-after-search-rerun validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored combined-view-state copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Failed to apply the deterministic Find header reorder before restored combined-view-state copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state "
                  L"copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state copy-after-search-rerun validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state copy-after-search-rerun persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-after-search-rerun-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to restore Find options for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state copy-after-search-rerun validation.");

    const std::wstring selectedPath         = (root / L"sub" / L"beta.txt").native();
    const std::wstring expectedRelativePath = L"sub";
    const std::wstring expectedName         = L"beta.txt";

    FindFilesDebugSnapshot reopened{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not restore the combined reordered, resized, and sorted state before restored combined copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not expose the expected selected combined results-grid state before restored combined copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state reopen.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state reopen.");

    const auto restoredCombinedCopyRerunStateOk = [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    };
    state.Require(DebugGetFindFilesWindowSnapshot(reopened),
                  L"Failed to capture Find snapshot after restored combined copy-after-search-rerun baseline rerun.");
    state.Require(restoredCombinedCopyRerunStateOk(reopened),
                  std::format(L"Reopened combined Find view-state should survive a baseline search rerun before restored combined copy-after-search-rerun "
                              L"validation. expectedWidth={:.1f} observedWidth={:.1f} widthDelta={:.1f} submittedNameOk={} submittedNameLen={} "
                              L"containsSelected={} expectedOrder={} "
                              L"selectedPath='{}' Last snapshot: {}",
                              resizedFirstVisibleWidthDip,
                              reopened.resultColumnWidthsDip.empty() ? 0.0f : reopened.resultColumnWidthsDip[0],
                              reopened.resultColumnWidthsDip.empty() ? resizedFirstVisibleWidthDip
                                                                     : std::fabs(reopened.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip),
                              reopened.submittedNamePatternText == L"*.txt" ? 1 : 0,
                              reopened.submittedNamePatternText.size(),
                              containsPath(reopened, selectedPath) ? 1 : 0,
                              matchesExpectedOrder(reopened) ? 1 : 0,
                              selectedPath,
                              DescribeFindSnapshotBrief(reopened)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore baseline Find search after restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state copy-after-search-rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should restore after returning to the baseline search rerun before restored combined "
                  L"copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the reopened Find results grid before restored combined copy-after-search-rerun validation.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid; },
                                      SelfTest::Scale(3000ms),
                                      &reopened),
                  L"Reopened Find results grid did not take focus before restored combined copy-after-search-rerun validation.");

    ClearClipboardContents(findWindow);
    SendMessageW(findWindow, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(findWindow, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(findWindow, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(findWindow, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(findWindow);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(),
                  L"Find Ctrl+C should copy the restored combined visible row content to the clipboard after reopen and search rerun.");
    state.Require(copiedSelection.rfind((expectedRelativePath + L"\t"), 0u) == 0u,
                  L"Reopened Find clipboard copy should still start with the restored visible Path column after combined view-state search rerun.");
    state.Require(copiedSelection.find(expectedName) != std::wstring::npos,
                  L"Reopened Find clipboard copy should still include the selected result name after combined view-state search rerun.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateGridEnterActivatesSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_enter_activation";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir), L"Failed to create alpha directory for restored combined-view-state Enter activation validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir), L"Failed to create beta directory for restored combined-view-state Enter activation validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir), L"Failed to create gamma directory for restored combined-view-state Enter activation validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state Enter activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state Enter activation validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state Enter activation validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state Enter activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state Enter activation validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state Enter activation validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-enter-activation-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state Enter activation validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state Enter activation validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state Enter activation validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state Enter activation validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state Enter activation validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                      L"Failed to apply the deterministic Find header reorder before restored combined-view-state Enter activation validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state Enter "
                      L"activation validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state Enter activation validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state Enter activation validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state Enter activation persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-enter-activation-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for Enter activation validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for Enter activation validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state Enter activation validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state Enter activation validation.");

    const std::wstring selectedPath = betaDir.native();

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before Enter activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before Enter activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(findWindow, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(findWindow, WM_KEYUP, VK_RETURN, 0);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Pressing Enter on the restored combined Find result did not activate the selected directory.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateGridDoubleClickActivatesSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_double_click_activation";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir),
                  L"Failed to create alpha directory for restored combined-view-state double-click activation validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir), L"Failed to create beta directory for restored combined-view-state double-click activation validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir),
                  L"Failed to create gamma directory for restored combined-view-state double-click activation validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state double-click activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state double-click activation validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state double-click activation validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state double-click activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state double-click activation validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state double-click activation validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-double-click-activation-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state double-click activation validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state double-click activation validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state double-click activation validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state double-click activation validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state double-click activation validation. {}",
                    DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                      L"Failed to apply the deterministic Find header reorder before restored combined-view-state double-click activation validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state double-click "
                      L"activation validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state double-click activation validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state double-click activation validation.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state double-click activation persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-double-click-activation-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for double-click activation validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for double-click activation validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state double-click activation validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state double-click activation validation.");

    const std::wstring selectedPath = betaDir.native();

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before double-click activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               value.selectedResultRowRect.right > value.selectedResultRowRect.left && value.selectedResultRowRect.bottom > value.selectedResultRowRect.top &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before double-click activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LPARAM doubleClickPoint = DipPointToEffectiveMouseInputLParam(findWindow,
                                                                        (reopened.selectedResultRowRect.left + reopened.selectedResultRowRect.right) * 0.5f,
                                                                        (reopened.selectedResultRowRect.top + reopened.selectedResultRowRect.bottom) * 0.5f);
    SendMouseDoubleClickToWindow(findWindow, doubleClickPoint);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Double-clicking the restored combined Find result did not activate the selected directory.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_action_buttons";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir), L"Failed to create alpha directory for restored combined-view-state action-button validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir), L"Failed to create beta directory for restored combined-view-state action-button validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir), L"Failed to create gamma directory for restored combined-view-state action-button validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state action-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state action-button validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state action-button validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state action-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state action-button validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state action-button validation.");
        FindFilesDebugSnapshot readySnapshot{};
        state.Require(WaitForFindWindowReady(SelfTest::Scale(3000ms), &readySnapshot),
                      std::format(L"Find window did not expose its live DX shell before restored combined-view-state action-button validation. {}",
                                  DescribeFindSnapshotBrief(readySnapshot)));
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state action-button validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state action-button validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state action-button validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state action-button validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state action-button validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                      L"Failed to apply the deterministic Find header reorder before restored combined-view-state action-button validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state action-button "
                      L"validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state action-button validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state action-button validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state action-button persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for action-button validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for action-button validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state action-button validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state action-button validation.");

    const std::wstring selectedPath     = betaDir.native();
    const std::wstring openButtonText   = LoadStringResource(nullptr, IDS_FIND_ACTION_OPEN);
    const std::wstring parentButtonText = LoadStringResource(nullptr, IDS_FIND_ACTION_PARENT);
    state.Require(! openButtonText.empty(), L"Failed to resolve Open button caption for restored combined-view-state action-button validation.");
    state.Require(! parentButtonText.empty(), L"Failed to resolve Open Parent button caption for restored combined-view-state action-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before action-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before action-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(findWindow, UIA_ButtonControlTypeId, openButtonText),
                  L"Open button did not expose live UIA InvokePattern interaction after restored combined state was reapplied.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Open button did not activate the selected directory after restored combined view state was reapplied.");

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' before restored combined-view-state Parent action validation.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-enable Open Parent on the selected combined-state result.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(findWindow, UIA_ButtonControlTypeId, parentButtonText),
                  L"Open Parent button did not expose live UIA InvokePattern interaction after restored combined state was reapplied.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(5000ms)),
                  L"Open Parent button did not navigate back to the selected directory parent after restored combined view state was reapplied.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Open Parent button did not restore the parent directory contents after restored combined view state was reapplied.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelectionAfterSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_action_buttons_after_sort_cycles";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir),
                  L"Failed to create alpha directory for restored combined-view-state action-button sort-cycle validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir), L"Failed to create beta directory for restored combined-view-state action-button sort-cycle validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir),
                  L"Failed to create gamma directory for restored combined-view-state action-button sort-cycle validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state action-button sort-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state action-button sort-cycle validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state action-button sort-cycle validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state action-button sort-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state action-button sort-cycle validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state action-button sort-cycle validation.");
        FindFilesDebugSnapshot readySnapshot{};
        state.Require(WaitForFindWindowReady(SelfTest::Scale(3000ms), &readySnapshot),
                      std::format(L"Find window did not expose its live DX shell before restored combined-view-state action-button sort-cycle validation. {}",
                                  DescribeFindSnapshotBrief(readySnapshot)));
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-after-sort-cycles-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state action-button sort-cycle validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state action-button sort-cycle validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state action-button sort-cycle validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state action-button sort-cycle validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state action-button sort-cycle validation. {}",
                    DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                      L"Failed to apply the deterministic Find header reorder before restored combined-view-state action-button sort-cycle validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state action-button "
                      L"sort-cycle validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state action-button sort-cycle validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state action-button sort-cycle validation.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state action-button sort-cycle persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-after-sort-cycles-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for action-button sort-cycle validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for action-button sort-cycle validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state action-button sort-cycle validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state action-button sort-cycle validation.");

    const std::wstring selectedPath     = betaDir.native();
    const std::wstring openButtonText   = LoadStringResource(nullptr, IDS_FIND_ACTION_OPEN);
    const std::wstring parentButtonText = LoadStringResource(nullptr, IDS_FIND_ACTION_PARENT);
    state.Require(! openButtonText.empty(), L"Failed to resolve Open button caption for restored combined-view-state action-button sort-cycle validation.");
    state.Require(! parentButtonText.empty(),
                  L"Failed to resolve Open Parent button caption for restored combined-view-state action-button sort-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before action-button sort-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before action-button sort-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    const LONG sortClickX = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.left + reopened.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.top + reopened.secondResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled &&
               value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should preserve Open and Open Parent enablement across repeated sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(findWindow, UIA_ButtonControlTypeId, openButtonText),
                  L"Open button did not expose live UIA InvokePattern interaction after restored combined state and reopened sort cycles.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Open button did not activate the selected directory after restored combined view state and reopened sort cycles.");

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' before restored combined-view-state Parent action sort-cycle validation.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-enable Open Parent on the selected combined-state result after sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(findWindow, UIA_ButtonControlTypeId, parentButtonText),
                  L"Open Parent button did not expose live UIA InvokePattern interaction after restored combined state and reopened sort cycles.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(5000ms)),
                  L"Open Parent button did not navigate back to the selected directory parent after restored combined view state and reopened sort cycles.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Open Parent button did not restore the parent directory contents after restored combined view state and reopened sort cycles.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelectionAfterSortCyclesAndSearchRerun(HWND mainWindow,
                                                                                                                      CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.failure = L"Main window unavailable for restored combined-view-state action-button sort-cycle search-rerun validation.";
        return false;
    }

    const auto closeFindWindow = []() noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_action_buttons_after_sort_cycles_and_search_rerun";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir),
                  L"Failed to create alpha directory for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir),
                  L"Failed to create beta directory for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir),
                  L"Failed to create gamma directory for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state action-button sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state action-button sort-cycle search-rerun validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state action-button sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state action-button sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state action-button sort-cycle search-rerun validation.");
        FindFilesDebugSnapshot readySnapshot{};
        state.Require(
            WaitForFindWindowReady(SelfTest::Scale(3000ms), &readySnapshot),
            std::format(
                L"Find window did not expose its live DX shell before restored combined-view-state action-button sort-cycle search-rerun validation. {}",
                DescribeFindSnapshotBrief(readySnapshot)));
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-after-sort-cycles-and-search-rerun-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state action-button sort-cycle search-rerun validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state action-button sort-cycle "
                              L"search-rerun validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(
            ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
            L"Failed to apply the deterministic Find header reorder before restored combined-view-state action-button sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state action-button "
                      L"sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        L"Find deterministic restored combined-view-state action-button sort-cycle search-rerun validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state action-button sort-cycle "
                  L"search-rerun persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-after-sort-cycles-and-search-rerun-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for action-button sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for action-button sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state action-button sort-cycle search-rerun validation.");

    const std::wstring selectedPath     = betaDir.native();
    const std::wstring openButtonText   = LoadStringResource(nullptr, IDS_FIND_ACTION_OPEN);
    const std::wstring parentButtonText = LoadStringResource(nullptr, IDS_FIND_ACTION_PARENT);
    state.Require(! openButtonText.empty(),
                  L"Failed to resolve Open button caption for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(! parentButtonText.empty(),
                  L"Failed to resolve Open Parent button caption for restored combined-view-state action-button sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not restore the combined reordered, resized, and sorted state before action-button sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not expose the expected selected combined results-grid state before action-button sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    const LONG sortClickX = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.left + reopened.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.top + reopened.secondResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled &&
               value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should preserve Open and Open Parent enablement across repeated sort cycles before the search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"beta*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to narrow Find search after restored combined-view-state sort cycles.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun narrowed Find search after restored combined-view-state sort cycles.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Narrowed Find search did not become idle after restored combined-view-state sort cycles.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               value.fullPaths.size() == 1u && value.fullPaths[0] == selectedPath;
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should survive a narrowed search rerun after sort cycles while preserving the visible layout.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore baseline Find search after restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state action-button sort-cycle search-rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               matchesExpectedOrder(value) && std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should restore after returning to the baseline search rerun after sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state sort-cycle search rerun.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-expose the expected selected combined results-grid state after sort cycles and search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(findWindow, UIA_ButtonControlTypeId, openButtonText),
                  L"Open button did not expose live UIA InvokePattern interaction after restored combined state, sort cycles, and search rerun.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Open button did not activate the selected directory after restored combined view state, sort cycles, and search rerun.");

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' before restored combined-view-state Parent action sort-cycle search-rerun validation.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-enable Open Parent on the selected combined-state result after sort cycles and search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(findWindow, UIA_ButtonControlTypeId, parentButtonText),
                  L"Open Parent button did not expose live UIA InvokePattern interaction after restored combined state, sort cycles, and search rerun.");
    state.Require(
        WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(5000ms)),
        L"Open Parent button did not navigate back to the selected directory parent after restored combined view state, sort cycles, and search rerun.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Open Parent button did not restore the parent directory contents after restored combined view state, sort cycles, and search rerun.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateGridEnterActivatesSelectionAfterSortCyclesAndSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.failure = L"Main window unavailable for restored combined-view-state Enter sort-cycle search-rerun validation.";
        return false;
    }

    const auto closeFindWindow = []() noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_enter_after_sort_cycles_and_search_rerun";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir),
                  L"Failed to create alpha directory for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir),
                  L"Failed to create beta directory for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir),
                  L"Failed to create gamma directory for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state Enter sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state Enter sort-cycle search-rerun validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state Enter sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state Enter sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state Enter sort-cycle search-rerun validation.");
        FindFilesDebugSnapshot readySnapshot{};
        state.Require(
            WaitForFindWindowReady(SelfTest::Scale(3000ms), &readySnapshot),
            std::format(L"Find window did not expose its live DX shell before restored combined-view-state Enter sort-cycle search-rerun validation. {}",
                        DescribeFindSnapshotBrief(readySnapshot)));
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-enter-after-sort-cycles-and-search-rerun-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state Enter sort-cycle search-rerun validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(
            L"Find window did not expose the baseline results-grid state before restored combined-view-state Enter sort-cycle search-rerun validation. {}",
            DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                      L"Failed to apply the deterministic Find header reorder before restored combined-view-state Enter sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state Enter "
                      L"sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        L"Find deterministic restored combined-view-state Enter sort-cycle search-rerun validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state Enter sort-cycle search-rerun persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-enter-after-sort-cycles-and-search-rerun-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for Enter sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for Enter sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state Enter sort-cycle search-rerun validation.");

    const std::wstring selectedPath = betaDir.native();
    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before Enter sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before Enter sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    const LONG sortClickX = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.left + reopened.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.top + reopened.secondResultHeaderRect.bottom) * 0.5f));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f &&
               value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should preserve selected Enter activation state across repeated sort cycles before the search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"beta*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to narrow Find search after restored combined-view-state Enter sort cycles.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun narrowed Find search after restored combined-view-state Enter sort cycles.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Narrowed Find search did not become idle after restored combined-view-state Enter sort cycles.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               value.fullPaths.size() == 1u && value.fullPaths[0] == selectedPath;
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should survive a narrowed search rerun after Enter sort cycles while preserving the visible layout.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore baseline Find search after restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               matchesExpectedOrder(value) && std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should restore after returning to the baseline search rerun after Enter sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state Enter sort-cycle search rerun.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-expose the expected selected combined results-grid state after Enter sort cycles and search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND enterTarget = GetFocus();
    if (! enterTarget || (enterTarget != findWindow && IsChild(findWindow, enterTarget) == FALSE))
    {
        enterTarget = findWindow;
    }

    SendMessageW(enterTarget, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(enterTarget, WM_KEYUP, VK_RETURN, 0);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Pressing Enter on the restored combined Find result did not activate the selected directory after sort cycles and search rerun.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateGridDoubleClickActivatesSelectionAfterSortCyclesAndSearchRerun(HWND mainWindow,
                                                                                                                         CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.failure = L"Main window unavailable for restored combined-view-state double-click sort-cycle search-rerun validation.";
        return false;
    }

    const auto closeFindWindow = []() noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });
    const auto traceStep                                                       = [](std::wstring_view step) noexcept
    {
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                                   std::format(L"find_dialog_restored_combined_view_state_grid_doubleClick_after_sort_cycles_and_search_rerun {}", step));
    };

    closeFindWindow();
    traceStep(L"begin");

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_double_click_after_sort_cycles_and_search_rerun";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir),
                  L"Failed to create alpha directory for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir),
                  L"Failed to create beta directory for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir),
                  L"Failed to create gamma directory for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state double-click sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state double-click sort-cycle search-rerun validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state double-click sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state double-click sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state double-click sort-cycle search-rerun validation.");
        FindFilesDebugSnapshot readySnapshot{};
        state.Require(
            WaitForFindWindowReady(SelfTest::Scale(3000ms), &readySnapshot),
            std::format(L"Find window did not expose its live DX shell before restored combined-view-state double-click sort-cycle search-rerun validation. {}",
                        DescribeFindSnapshotBrief(readySnapshot)));
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-double-click-after-sort-cycles-and-search-rerun-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state double-click sort-cycle search-rerun validation.");
    traceStep(L"initial_search_idle");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state double-click sort-cycle "
                              L"search-rerun validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(
            ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
            L"Failed to apply the deterministic Find header reorder before restored combined-view-state double-click sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state double-click "
                      L"sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        L"Find deterministic restored combined-view-state double-click sort-cycle search-rerun validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state double-click sort-cycle search-rerun "
                  L"persistence.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"saved_layout_sorted");

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-double-click-after-sort-cycles-and-search-rerun-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for double-click sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for double-click sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state double-click sort-cycle search-rerun validation.");
    traceStep(L"reopened_search_idle");

    const std::wstring selectedPath = betaDir.native();
    FindFilesDebugSnapshot reopened{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not restore the combined reordered, resized, and sorted state before double-click sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"reopened_selected");

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               value.selectedResultRowRect.right > value.selectedResultRowRect.left && value.selectedResultRowRect.bottom > value.selectedResultRowRect.top &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not expose the expected selected combined results-grid state before double-click sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"sort_cycles_settled");

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    const LONG sortClickX = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.left + reopened.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.top + reopened.secondResultHeaderRect.bottom) * 0.5f));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f &&
               value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened combined Find view-state should preserve selected double-click activation state across repeated sort cycles before the search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"beta*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to narrow Find search after restored combined-view-state double-click sort cycles.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun narrowed Find search after restored combined-view-state double-click sort cycles.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Narrowed Find search did not become idle after restored combined-view-state double-click sort cycles.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               value.fullPaths.size() == 1u && value.fullPaths[0] == selectedPath;
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened combined Find view-state should survive a narrowed search rerun after double-click sort cycles while preserving the visible layout.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"narrowed_rerun_settled");

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore baseline Find search after restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               matchesExpectedOrder(value) && std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should restore after returning to the baseline search rerun after double-click sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"baseline_rerun_settled");

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state double-click sort-cycle search rerun.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               value.selectedResultRowRect.right > value.selectedResultRowRect.left && value.selectedResultRowRect.bottom > value.selectedResultRowRect.top &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-expose the expected selected combined results-grid state after double-click sort cycles and search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"final_selection_ready");

    const LPARAM doubleClickPoint = DipPointToEffectiveMouseInputLParam(findWindow,
                                                                        (reopened.selectedResultRowRect.left + reopened.selectedResultRowRect.right) * 0.5f,
                                                                        (reopened.selectedResultRowRect.top + reopened.selectedResultRowRect.bottom) * 0.5f);
    traceStep(L"dispatch_double_click");
    SendMouseDoubleClickToWindow(findWindow, doubleClickPoint);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Double-clicking the restored combined Find result did not activate the selected directory after sort cycles and search rerun.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogThemeCycleKeepsGridLegible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_theme_cycle";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create theme-cycle Find root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create theme-cycle Find subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for theme-cycle Find test.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for theme-cycle Find test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesPaneContext context{};
    context.rootPluginPath      = root;
    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-theme-cycle-initial");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, initialTheme, std::move(context)), L"Failed to open Find window for theme-cycle validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for theme-cycle validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for theme-cycle validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for theme-cycle validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for theme-cycle validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept { return value.usesDxUiHost && value.resultCount == 2u; },
                                      SelfTest::Scale(10000ms),
                                      &snapshot),
                  L"Find window did not produce the expected results for theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for theme-cycle validation.", selectedPath));
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.selectedResultCount == 1u && value.selectedResultRowFillArgb != 0u && value.selectedResultRowTextArgb != 0u; },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Find window did not expose a selected row for theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring expectedLeaf = std::filesystem::path(selectedPath).filename().native();
    std::wstring baselineSelectedName;
    const auto requireSelectedUiaRowState = [&](std::wstring_view label) noexcept
    {
        const auto selectionState = CollectVisibleDescendantSelectionPatternState(findWindow, UIA_DataGridControlTypeId);
        state.Require(selectionState.has_value(), std::format(L"Failed to collect Find results-grid UIA selection state after {}.", label));
        if (! selectionState.has_value())
        {
            return false;
        }

        state.Require(selectionState->rootControlType == UIA_DataGridControlTypeId,
                      std::format(L"Find results-grid root should stay a DataGrid after {}.", label));
        state.Require(selectionState->hasSelectionPattern, std::format(L"Find results-grid should keep SelectionPattern after {}.", label));
        state.Require(selectionState->selectionCount == 1u,
                      std::format(L"Find results-grid should keep exactly one selected UIA row after {}; saw {}.", label, selectionState->selectionCount));
        state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                      std::format(L"Find selected UIA row should stay a DataItem after {}.", label));
        state.Require(selectionState->selectedHasSelectionItemPattern, std::format(L"Find selected UIA row should keep SelectionItemPattern after {}.", label));
        state.Require(! selectionState->selectedName.empty(), std::format(L"Find selected UIA row should keep a non-empty accessible name after {}.", label));
        state.Require(selectionState->selectedName.find(expectedLeaf) != std::wstring::npos,
                      std::format(L"Find selected UIA row name '{}' should still include '{}' after {}.", selectionState->selectedName, expectedLeaf, label));
        if (! state.failure.empty())
        {
            return false;
        }

        if (baselineSelectedName.empty())
        {
            baselineSelectedName = selectionState->selectedName;
        }

        state.Require(selectionState->selectedName == baselineSelectedName,
                      std::format(L"Find selected UIA row accessible name should stay stable after {}; expected '{}', saw '{}'.",
                                  label,
                                  baselineSelectedName,
                                  selectionState->selectedName));
        return state.failure.empty();
    };

    state.Require(requireSelectedUiaRowState(L"the baseline theme-cycle selection capture"),
                  L"Find window baseline UIA selection state was not stable before theme churn.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto unpackColor = [](uint32_t argb) noexcept
    {
        return D2D1::ColorF(static_cast<float>((argb >> 16u) & 0xFFu) / 255.0f,
                            static_cast<float>((argb >> 8u) & 0xFFu) / 255.0f,
                            static_cast<float>(argb & 0xFFu) / 255.0f,
                            static_cast<float>((argb >> 24u) & 0xFFu) / 255.0f);
    };
    const auto luminance = [&](uint32_t argb) noexcept
    {
        const D2D1_COLOR_F color = unpackColor(argb);
        const auto linearize = [](float channel) noexcept { return (channel <= 0.03928f) ? (channel / 12.92f) : std::pow((channel + 0.055f) / 1.055f, 2.4f); };

        return (0.2126f * linearize(color.r)) + (0.7152f * linearize(color.g)) + (0.0722f * linearize(color.b));
    };
    const auto contrastRatio = [&](uint32_t a, uint32_t b) noexcept
    {
        const float lumA    = luminance(a);
        const float lumB    = luminance(b);
        const float lighter = (std::max)(lumA, lumB);
        const float darker  = (std::min)(lumA, lumB);
        return (lighter + 0.05f) / (darker + 0.05f);
    };

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        const uint64_t previousRenderCount = snapshot.dxRenderCount;
        UpdateFindFilesWindowsTheme(theme);
        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultCount == 2u && value.selectedResultCount == 1u &&
                   value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast && value.themeRainbow == theme.menu.rainbowMode &&
                   value.selectedResultRowFillArgb != 0u && value.selectedResultRowTextArgb != 0u && value.dxRenderCount >= previousRenderCount;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      std::format(L"Find window did not settle after switching to {} theme.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(IsWindow(findWindow) != FALSE, std::format(L"Find window did not survive the {} theme update.", label));
        state.Require(snapshot.selectedResultRowUsesRainbow == expectRainbow,
                      std::format(L"Find selected-row rainbow state mismatch after {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast, std::format(L"Find high-contrast state mismatch after {} theme update.", label));
        state.Require(snapshot.selectedResultRowFillArgb != snapshot.selectedResultRowTextArgb,
                      std::format(L"Find selected-row colors collapsed to the same value after {} theme update.", label));

        const float minimumContrast = expectHighContrast ? 4.5f : 3.0f;
        state.Require(contrastRatio(snapshot.selectedResultRowFillArgb, snapshot.selectedResultRowTextArgb) >= minimumContrast,
                      std::format(L"Find selected-row text contrast dropped below {:.1f}:1 after {} theme update.", minimumContrast, label));
        state.Require(requireSelectedUiaRowState(std::format(L"the {} theme update", label)),
                      std::format(L"Find selected UIA row state did not remain stable after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"find-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"find-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"find-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"find-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogCompactModeShrinksResultsGridMetrics(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_compact_mode";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create compact-mode Find root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create compact-mode Find subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha compact test\n"), L"Failed to create alpha.txt for compact-mode Find test.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta compact test\n"), L"Failed to create beta.txt for compact-mode Find test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesPaneContext context{};
    context.rootPluginPath    = root;
    AppTheme standardTheme    = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-compact-standard");
    standardTheme.compactMode = false;

    state.Require(ShowFindFilesWindow(mainWindow, g_settings, standardTheme, std::move(context)), L"Failed to open Find window for compact-mode validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for compact-mode validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for compact-mode validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for compact-mode validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for compact-mode validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.firstResultHeaderRect.bottom > value.firstResultHeaderRect.top && value.statusStripVisible && value.statusStripHeightDip > 0.0f &&
               ! value.themeCompactMode;
    },
                      SelfTest::Scale(10000ms),
                      &snapshot),
                  L"Find window did not reach the standard-density baseline state for compact-mode validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for compact-mode validation.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.selectedResultRowRect.right > value.selectedResultRowRect.left &&
               value.selectedResultRowRect.bottom > value.selectedResultRowRect.top && value.statusStripHeightDip > 0.0f && ! value.themeCompactMode;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose a selected row for compact-mode validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineHeaderHeight = std::max(0.0f, snapshot.firstResultHeaderRect.bottom - snapshot.firstResultHeaderRect.top);
    const float baselineRowHeight    = std::max(0.0f, snapshot.selectedResultRowRect.bottom - snapshot.selectedResultRowRect.top);
    const float baselineStatusHeight = snapshot.statusStripHeightDip;

    AppTheme compactTheme              = standardTheme;
    compactTheme.compactMode           = true;
    const uint64_t baselineRenderCount = snapshot.dxRenderCount;
    UpdateFindFilesWindowsTheme(compactTheme);

    FindFilesDebugSnapshot compactSnapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.themeCompactMode &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.firstResultHeaderRect.bottom > value.firstResultHeaderRect.top &&
               value.selectedResultRowRect.right > value.selectedResultRowRect.left && value.selectedResultRowRect.bottom > value.selectedResultRowRect.top &&
               value.statusStripHeightDip > 0.0f && value.dxRenderCount >= baselineRenderCount;
    },
                      SelfTest::Scale(3000ms),
                      &compactSnapshot),
                  L"Find window did not settle into compact density.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float compactHeaderHeight = std::max(0.0f, compactSnapshot.firstResultHeaderRect.bottom - compactSnapshot.firstResultHeaderRect.top);
    const float compactRowHeight    = std::max(0.0f, compactSnapshot.selectedResultRowRect.bottom - compactSnapshot.selectedResultRowRect.top);

    state.Require(compactHeaderHeight + 0.5f < baselineHeaderHeight,
                  std::format(L"Find compact header height should shrink below the standard height; baseline {:.2f} dip vs compact {:.2f} dip.",
                              baselineHeaderHeight,
                              compactHeaderHeight));
    state.Require(compactRowHeight + 0.5f < baselineRowHeight,
                  std::format(L"Find compact row height should shrink below the standard height; baseline {:.2f} dip vs compact {:.2f} dip.",
                              baselineRowHeight,
                              compactRowHeight));
    state.Require(compactSnapshot.statusStripHeightDip + 0.5f < baselineStatusHeight,
                  std::format(L"Find compact status strip height should shrink below the standard height; baseline {:.2f} dip vs compact {:.2f} dip.",
                              baselineStatusHeight,
                              compactSnapshot.statusStripHeightDip));
    if (! state.failure.empty())
    {
        return false;
    }

    UpdateFindFilesWindowsTheme(standardTheme);
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && ! value.themeCompactMode &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.selectedResultRowRect.right > value.selectedResultRowRect.left &&
               value.statusStripHeightDip > 0.0f;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not return to standard density after compact-mode validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float restoredHeaderHeight = std::max(0.0f, snapshot.firstResultHeaderRect.bottom - snapshot.firstResultHeaderRect.top);
    const float restoredRowHeight    = std::max(0.0f, snapshot.selectedResultRowRect.bottom - snapshot.selectedResultRowRect.top);
    state.Require(restoredHeaderHeight + 0.5f >= baselineHeaderHeight,
                  std::format(L"Find header height did not restore after leaving compact mode; baseline {:.2f} dip vs restored {:.2f} dip.",
                              baselineHeaderHeight,
                              restoredHeaderHeight));
    state.Require(restoredRowHeight + 0.5f >= baselineRowHeight,
                  std::format(L"Find row height did not restore after leaving compact mode; baseline {:.2f} dip vs restored {:.2f} dip.",
                              baselineRowHeight,
                              restoredRowHeight));
    state.Require(snapshot.statusStripHeightDip + 0.5f >= baselineStatusHeight,
                  std::format(L"Find status strip height did not restore after leaving compact mode; baseline {:.2f} dip vs restored {:.2f} dip.",
                              baselineStatusHeight,
                              snapshot.statusStripHeightDip));
    return state.failure.empty();
}

[[nodiscard]] bool QuickSearchSnapshotHasMatch(const FolderView::IncrementalSearchDebugSnapshot& snapshot,
                                               std::wstring_view displayName,
                                               UINT32 startPosition,
                                               UINT32 length,
                                               bool startsWith) noexcept
{
    return std::any_of(snapshot.matches.begin(),
                       snapshot.matches.end(),
                       [&](const FolderView::IncrementalSearchDebugMatch& match) noexcept
    {
        return match.displayName == displayName && match.range.startPosition == startPosition && match.range.length == length &&
               match.startsWith == startsWith;
    });
}

[[nodiscard]] bool TestPaneQuickSearchIntegratedNavigation(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"quick_search_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create quick-search test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for quick-search test.");
    state.Require(SelfTest::WriteTextFile(root / L"alpine.log", "alpine"), L"Failed to create alpine.log for quick-search test.");
    state.Require(SelfTest::WriteTextFile(root / L"beta-alpha.txt", "beta"), L"Failed to create beta-alpha.txt for quick-search test.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma"), L"Failed to create gamma.txt for quick-search test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for quick-search test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for quick-search test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"alpine.log", L"beta-alpha.txt", L"gamma.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for quick-search test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"gamma.txt"),
                  L"Failed to focus gamma.txt before quick-search test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Left folder view handle unavailable for quick-search test.");
    if (! folderView)
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_QUICK_SEARCH, 0), 0);
    PumpPendingMessages();

    FolderView::IncrementalSearchDebugSnapshot snapshot{};
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after command activation.");
    state.Require(snapshot.active, L"Quick Search command should activate integrated pane search.");
    state.Require(snapshot.query.empty(), L"Quick Search command should start with an empty query.");

    SendMessageW(folderView, WM_CHAR, static_cast<WPARAM>(L'a'), 0);
    SendMessageW(folderView, WM_CHAR, static_cast<WPARAM>(L'l'), 0);
    PumpPendingMessages();

    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after typing.");
    state.Require(snapshot.active, L"Quick Search should remain active after typing.");
    state.Require(snapshot.query == L"al", std::format(L"Quick Search query should be 'al'; got '{}'.", snapshot.query));
    state.Require(snapshot.focusedDisplayName == L"alpha.txt",
                  std::format(L"Quick Search should select the first starts-with match; got '{}'.", snapshot.focusedDisplayName));
    state.Require(snapshot.matches.size() == 3u, std::format(L"Quick Search should highlight three containing matches; got {}.", snapshot.matches.size()));
    state.Require(QuickSearchSnapshotHasMatch(snapshot, L"alpha.txt", 0u, 2u, true), L"Quick Search should highlight prefix match alpha.txt.");
    state.Require(QuickSearchSnapshotHasMatch(snapshot, L"alpine.log", 0u, 2u, true), L"Quick Search should highlight prefix match alpine.log.");
    state.Require(QuickSearchSnapshotHasMatch(snapshot, L"beta-alpha.txt", 5u, 2u, false), L"Quick Search should highlight contained match beta-alpha.txt.");

    SendMessageW(folderView, WM_KEYDOWN, VK_DOWN, 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after first match navigation.");
    state.Require(snapshot.focusedDisplayName == L"alpine.log",
                  std::format(L"Quick Search Down should move to the next match; got '{}'.", snapshot.focusedDisplayName));

    SendMessageW(folderView, WM_KEYDOWN, VK_DOWN, 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after second match navigation.");
    state.Require(snapshot.focusedDisplayName == L"beta-alpha.txt",
                  std::format(L"Quick Search Down should include contained matches; got '{}'.", snapshot.focusedDisplayName));

    SendMessageW(folderView, WM_KEYDOWN, VK_RETURN, 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after Enter.");
    state.Require(! snapshot.active, L"Quick Search Enter should accept the current item and exit search mode.");
    state.Require(snapshot.query.empty(), L"Quick Search Enter should clear the query.");
    state.Require(snapshot.focusedDisplayName == L"beta-alpha.txt", L"Quick Search Enter should keep focus on the accepted item.");
    state.Require(g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left).value_or(std::filesystem::path{}) == root,
                  L"Quick Search Enter should not navigate away from the folder.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_QUICK_SEARCH, 0), 0);
    SendMessageW(folderView, WM_CHAR, static_cast<WPARAM>(L'z'), 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search no-match snapshot should be available.");
    state.Require(snapshot.active, L"Quick Search no-match state should remain active.");
    state.Require(snapshot.query == L"z", L"Quick Search no-match query should remain visible.");
    state.Require(snapshot.matches.empty(), std::format(L"Quick Search no-match state should expose zero matches; got {}.", snapshot.matches.size()));

    SendMessageW(folderView, WM_KEYDOWN, VK_ESCAPE, 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after Escape.");
    state.Require(! snapshot.active, L"Quick Search Escape should exit search mode.");
    state.Require(snapshot.query.empty(), L"Quick Search Escape should clear the query.");

    return state.failure.empty();
}

} // namespace (tests)

void RunSearchCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(options, suite, L"search_local_plugin_invalid_regex_reports_single_completion", [](CaseState& state) noexcept {
        return TestLocalPluginInvalidRegexReportsSingleCompletion(state);
    });
    SelfTest::RunCase(options, suite, L"search_local_plugin_parallel_cancel_fanin", [](CaseState& state) noexcept {
        return TestLocalPluginParallelSearchCancellationAndFanIn(state);
    });
    SelfTest::RunCase(options, suite, L"filesystem_local_watch_unwatch_drains_inflight_callback", [](CaseState& state) noexcept {
        return TestLocalPluginWatchUnwatchDrainsInflightCallback(state);
    });
    SelfTest::RunCase(options, suite, L"search_local_index_stream_stop_after_first", [](CaseState& state) noexcept {
        return TestLocalSearchIndexEnumerateStopsAfterFirstCandidate(state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_find_dialog_search_ops", [=](CaseState& state) noexcept { return TestFindDialogSearchOps(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_quickSearch_integrated_navigation", [=](CaseState& state) noexcept { return TestPaneQuickSearchIntegratedNavigation(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_local_root_overrides_stale_context", [=](CaseState& state) noexcept {
        return TestFindDialogTypedLocalRootOverridesStaleContext(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_large_local_search_uses_incremental_updates", [=](CaseState& state) noexcept {
        return TestFindDialogLargeLocalSearchUsesIncrementalUpdates(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_running_status_shows_phase_and_path", [=](CaseState& state) noexcept {
        return TestFindDialogRunningStatusShowsPhaseAndPath(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_service_status_shows_backend_diagnostics", [=](CaseState& state) noexcept {
        return TestFindDialogServiceStatusShowsBackendDiagnostics(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_service_unavailable_warning_is_distinct", [=](CaseState& state) noexcept {
        return TestFindDialogServiceUnavailableWarningIsDistinct(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_failure_status_is_readable", [=](CaseState& state) noexcept {
        return TestFindDialogFailureShowsReadableStatus(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_uses_dxui_host_without_visible_child_controls", [=](CaseState& state) noexcept {
        return TestFindDialogUsesDxUiHostWithNoVisibleChildControls(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_exposes_live_uia_selection_and_inputs", [=](CaseState& state) noexcept {
        return TestFindDialogExposesLiveUiaSelectionAndInputs(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_long_run_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestFindDialogLongRunScrollingStaysBounded(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestFindDialogLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_directory_activation_navigates_into_selection", [=](CaseState& state) noexcept {
        return TestFindDialogDirectoryActivationNavigatesIntoSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_open_parent_keeps_directory_focused_in_parent", [=](CaseState& state) noexcept {
        return TestFindDialogOpenParentKeepsDirectoryFocusedInParent(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_enter_from_checkbox_invokes_default_search", [=](CaseState& state) noexcept {
        return TestFindDialogEnterFromCheckboxInvokesDefaultSearch(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_pointer_click_toggles_recursive_checkbox", [=](CaseState& state) noexcept {
        return TestFindDialogPointerClickTogglesRecursiveCheckbox(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_escape_closes_popup_before_cancel", [=](CaseState& state) noexcept {
        return TestFindDialogEscapeClosesPopupBeforeCancel(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_escape_from_dx_control_closes_cancel", [=](CaseState& state) noexcept {
        return TestFindDialogEscapeFromDxControlClosesCancel(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_access_keys_focus_expected_fields", [=](CaseState& state) noexcept {
        return TestFindDialogAccessKeysFocusExpectedFields(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_grid_enter_activates_selection", [=](CaseState& state) noexcept {
        return TestFindDialogGridFocusedEnterActivatesSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_grid_doubleClick_activates_selection", [=](CaseState& state) noexcept {
        return TestFindDialogGridDoubleClickActivatesSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_tab_traversal_matches_expected_order", [=](CaseState& state) noexcept {
        return TestFindDialogTabTraversalMatchesExpectedOrder(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_mode_typeahead_updates_selection_and_dependencies", [=](CaseState& state) noexcept {
        return TestFindDialogModeTypeaheadUpdatesSelectionAndDependencies(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_command_enablement_matches_idle_running_and_selection_states", [=](CaseState& state) noexcept {
        return TestFindDialogCommandEnablementMatchesIdleRunningAndSelectionStates(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_action_buttons_activate_expected_commands", [=](CaseState& state) noexcept {
        return TestFindDialogActionButtonsActivateExpectedCommands(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_persisted_grid_layout", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresPersistedGridLayout(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_header_drag_reorders_columns_without_sort", [=](CaseState& state) noexcept {
        return TestFindDialogHeaderDragReordersColumnsWithoutSort(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_copy_follows_reordered_columns", [=](CaseState& state) noexcept {
        return TestFindDialogCopyFollowsReorderedColumns(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_reordered_columns_survive_sort_cycles", [=](CaseState& state) noexcept {
        return TestFindDialogReorderedColumnsSurviveSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_header_resize_changes_visible_width", [=](CaseState& state) noexcept {
        return TestFindDialogHeaderResizeChangesVisibleWidth(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_reordered_columns_survive_search_rerun", [=](CaseState& state) noexcept {
        return TestFindDialogReorderedColumnsSurviveSearchRerun(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_resized_columns_survive_search_rerun", [=](CaseState& state) noexcept {
        return TestFindDialogResizedColumnsSurviveSearchRerun(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_reordered_resized_columns_survive_search_rerun", [=](CaseState& state) noexcept {
        return TestFindDialogReorderedResizedColumnsSurviveSearchRerun(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_reordered_resized_columns_survive_sort_cycles", [=](CaseState& state) noexcept {
        return TestFindDialogReorderedResizedColumnsSurviveSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_resized_columns_survive_sort_cycles", [=](CaseState& state) noexcept {
        return TestFindDialogResizedColumnsSurviveSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_header_click_sorts_results", [=](CaseState& state) noexcept {
        return TestFindDialogHeaderClickSortsResults(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_resized_grid_layout", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresResizedGridLayout(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_reordered_grid_layout", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresReorderedGridLayout(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_reordered_layout_copy_follows_visible_columns", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredReorderedLayoutCopyFollowsVisibleColumns(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_persisted_sort_order", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresPersistedSortOrder(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_reordered_sorted_grid_layout", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresReorderedSortedGridLayout(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_combined_view_state", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresCombinedViewState(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumns(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_survives_search_rerun", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateSurvivesSearchRerun(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_survives_sort_cycles", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateSurvivesSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns_after_sort_cycles", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumnsAfterSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns_after_search_rerun", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumnsAfterSearchRerun(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_grid_enter_activates_selection", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateGridEnterActivatesSelection(mainWindow, state);
    });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_pane_find_dialog_restored_combined_view_state_grid_doubleClick_activates_selection",
                      [=](CaseState& state) noexcept { return TestFindDialogRestoredCombinedViewStateGridDoubleClickActivatesSelection(mainWindow, state); });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_pane_find_dialog_restored_combined_view_state_grid_enter_activates_selection_after_sort_cycles_and_search_rerun",
                      [=](CaseState& state) noexcept
    { return TestFindDialogRestoredCombinedViewStateGridEnterActivatesSelectionAfterSortCyclesAndSearchRerun(mainWindow, state); });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_pane_find_dialog_restored_combined_view_state_grid_doubleClick_activates_selection_after_sort_cycles_and_search_rerun",
                      [=](CaseState& state) noexcept
    { return TestFindDialogRestoredCombinedViewStateGridDoubleClickActivatesSelectionAfterSortCyclesAndSearchRerun(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_action_buttons_activate_selection", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelection(mainWindow, state);
    });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_pane_find_dialog_restored_combined_view_state_action_buttons_activate_selection_after_sort_cycles",
                      [=](CaseState& state) noexcept
    { return TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelectionAfterSortCycles(mainWindow, state); });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_pane_find_dialog_restored_combined_view_state_action_buttons_activate_selection_after_sort_cycles_and_search_rerun",
                      [=](CaseState& state) noexcept
    { return TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelectionAfterSortCyclesAndSearchRerun(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_theme_cycle_keeps_grid_legible", [=](CaseState& state) noexcept {
        return TestFindDialogThemeCycleKeepsGridLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_compact_mode_shrinks_results_grid_metrics", [=](CaseState& state) noexcept {
        return TestFindDialogCompactModeShrinksResultsGridMetrics(mainWindow, state);
    });
}

namespace
{
