SelfTest::RunCase(options,
                  suite,
                  L"local_search_qi_and_capabilities",
                  [&](SelfTest::CaseState& state) noexcept
{
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(baseFs, search), L"Local file system plugin missing IFileSystemSearch.");

    const char* capabilities = nullptr;
    const HRESULT hr         = baseFs->GetCapabilities(&capabilities);
    state.Require(SUCCEEDED(hr) && capabilities != nullptr,
                  std::format(L"Local file system GetCapabilities failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (capabilities != nullptr)
    {
        const std::string_view capabilitiesView(capabilities);
        state.Require(capabilitiesView.find("\"search\"") != std::string_view::npos, L"Capabilities JSON missing search section.");
        state.Require(capabilitiesView.find("\"preferredBackend\":\"service\"") != std::string_view::npos,
                      L"Capabilities JSON missing preferredBackend=service.");
        state.Require(capabilitiesView.find("\"indexed\":true") != std::string_view::npos, L"Capabilities JSON missing indexed=true.");
        state.Require(capabilitiesView.find("\"serviceBacked\":true") != std::string_view::npos, L"Capabilities JSON missing serviceBacked=true.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_callback_contract",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_callback_contract", caseRoot), L"Failed to prepare search_callback_contract root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"match.txt", "contract"), L"Failed to create match.txt.");

    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(baseFs, search), L"Local file system plugin missing IFileSystemSearch.");

    std::wstring rootText    = caseRoot.wstring();
    std::wstring namePattern = L"*.txt";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    RecordingSearchCallback successCallback;
    HRESULT hr = search->Search(&query, &successCallback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Search success case failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(successCallback.ProgressCalls() >= 1u, L"Search success case expected at least one progress callback.");
    state.Require(successCallback.CancelCalls() >= 1u, L"Search success case expected at least one cancel callback.");
    state.Require(successCallback.NullPathProgressCalls() >= 1u, L"Search success case expected a final null-path progress callback.");

    RecordingSearchCallback failProgressCallback(RecordingSearchCallback::Mode::FailProgress);
    hr = search->Search(&query, &failProgressCallback, nullptr);
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_RETRY),
                  std::format(L"Search fail-progress case expected ERROR_RETRY. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    RecordingSearchCallback failCancelCallback(RecordingSearchCallback::Mode::FailCancel);
    hr = search->Search(&query, &failCancelCallback, nullptr);
    state.Require(hr == E_ACCESSDENIED, std::format(L"Search fail-cancel case expected E_ACCESSDENIED. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    RecordingSearchCallback cancelCallback(RecordingSearchCallback::Mode::Cancel);
    hr = search->Search(&query, &cancelCallback, nullptr);
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Search cancel case expected ERROR_CANCELLED. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    RecordingSearchCallback abortProgressCallback(RecordingSearchCallback::Mode::AbortProgress);
    hr = search->Search(&query, &abortProgressCallback, nullptr);
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Search abort-progress case expected ERROR_CANCELLED. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    RecordingSearchCallback abortMatchCallback(RecordingSearchCallback::Mode::AbortMatch);
    hr = search->Search(&query, &abortMatchCallback, nullptr);
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Search abort-match case expected ERROR_CANCELLED. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_backend_preferences_roundtrip",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
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

    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! search)
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring unavailablePipe      = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, unavailablePipe.c_str()) != 0,
                  L"Failed to override the search service pipe for the unavailable-service preference test.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_backend_preferences_roundtrip", caseRoot),
                  L"Failed to prepare search_backend_preferences_roundtrip root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"pref.txt", "preference"), L"Failed to create pref.txt.");

    const char* schema = nullptr;
    HRESULT hr         = info->GetConfigurationSchema(&schema);
    state.Require(SUCCEEDED(hr) && schema != nullptr,
                  std::format(L"Local file system GetConfigurationSchema failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (schema != nullptr)
    {
        const std::string_view schemaView(schema);
        state.Require(schemaView.find("\"searchBackendPreference\"") != std::string_view::npos, L"Configuration schema missing searchBackendPreference.");
        state.Require(schemaView.find("\"searchMaxDirectoryWalkers\"") != std::string_view::npos, L"Configuration schema missing searchMaxDirectoryWalkers.");
        state.Require(schemaView.find("\"auto\"") != std::string_view::npos, L"Configuration schema missing auto search backend value.");
        state.Require(schemaView.find("\"service\"") != std::string_view::npos, L"Configuration schema missing service search backend value.");
        state.Require(schemaView.find("\"local-index\"") != std::string_view::npos, L"Configuration schema missing local-index search backend value.");
        state.Require(schemaView.find("\"scan\"") != std::string_view::npos, L"Configuration schema missing scan search backend value.");
    }

    const auto verifyCapabilities = [&](std::string_view configuredPreference) noexcept
    {
        const char* capabilities     = nullptr;
        const HRESULT capabilitiesHr = created.fileSystem->GetCapabilities(&capabilities);
        state.Require(SUCCEEDED(capabilitiesHr) && capabilities != nullptr,
                      std::format(L"Local file system GetCapabilities failed for {}. hr=0x{:08X}",
                                  std::wstring(configuredPreference.begin(), configuredPreference.end()),
                                  static_cast<unsigned long>(capabilitiesHr)));
        if (capabilities == nullptr)
        {
            return;
        }

        const std::string_view capabilitiesView(capabilities);
        state.Require(capabilitiesView.find("\"preferredBackend\":\"service\"") != std::string_view::npos,
                      L"Capabilities JSON should advertise preferredBackend=service once the Windows service backend exists.");
        state.Require(capabilitiesView.find("\"indexed\":true") != std::string_view::npos,
                      L"Capabilities JSON should advertise indexed=true once the local index backend exists.");
        state.Require(capabilitiesView.find("\"serviceBacked\":true") != std::string_view::npos,
                      L"Capabilities JSON should advertise serviceBacked=true once the broker exists.");
    };

    const auto runPreferenceCase = [&](std::string_view configuredPreference,
                                       BOOL expectedDirty,
                                       FileSystemSearchBackend expectedBackend,
                                       bool expectDegraded,
                                       bool expectServiceUnavailable,
                                       bool preferIndexHint) noexcept
    {
        const std::string configurationJson = std::format("{{\"searchBackendPreference\":\"{}\"}}", configuredPreference);
        const HRESULT setHr                 = info->SetConfiguration(configurationJson.c_str());
        state.Require(SUCCEEDED(setHr),
                      std::format(L"SetConfiguration failed for {}. hr=0x{:08X}",
                                  std::wstring(configuredPreference.begin(), configuredPreference.end()),
                                  static_cast<unsigned long>(setHr)));
        if (FAILED(setHr))
        {
            return;
        }

        BOOL somethingToSave = FALSE;
        const HRESULT saveHr = info->SomethingToSave(&somethingToSave);
        state.Require(SUCCEEDED(saveHr),
                      std::format(L"SomethingToSave failed for {}. hr=0x{:08X}",
                                  std::wstring(configuredPreference.begin(), configuredPreference.end()),
                                  static_cast<unsigned long>(saveHr)));
        if (SUCCEEDED(saveHr))
        {
            state.Require(somethingToSave == expectedDirty,
                          std::format(L"Unexpected SomethingToSave result for {}.", std::wstring(configuredPreference.begin(), configuredPreference.end())));
        }

        const char* configuration = nullptr;
        const HRESULT getHr       = info->GetConfiguration(&configuration);
        state.Require(SUCCEEDED(getHr) && configuration != nullptr,
                      std::format(L"GetConfiguration failed for {}. hr=0x{:08X}",
                                  std::wstring(configuredPreference.begin(), configuredPreference.end()),
                                  static_cast<unsigned long>(getHr)));
        if (configuration != nullptr)
        {
            const std::string expectedField = std::format("\"searchBackendPreference\":\"{}\"", configuredPreference);
            state.Require(std::string_view(configuration).find(expectedField) != std::string_view::npos,
                          std::format(L"Configuration JSON missing searchBackendPreference={} field.",
                                      std::wstring(configuredPreference.begin(), configuredPreference.end())));
        }

        verifyCapabilities(configuredPreference);

        std::wstring rootText    = caseRoot.wstring();
        std::wstring namePattern = L"*.txt";

        FileSystemSearchQuery query{};
        query.sizeBytes   = sizeof(FileSystemSearchQuery);
        query.rootPath    = rootText.c_str();
        query.namePattern = namePattern.c_str();
        query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES |
                                                               (preferIndexHint ? FILESYSTEM_SEARCH_PREFER_INDEX : 0u));
        query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
        query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

        RecordingSearchCallback callback;
        const HRESULT searchHr = search->Search(&query, &callback, nullptr);
        state.Require(SUCCEEDED(searchHr),
                      std::format(L"Search failed for preference {}. hr=0x{:08X}",
                                  std::wstring(configuredPreference.begin(), configuredPreference.end()),
                                  static_cast<unsigned long>(searchHr)));
        if (FAILED(searchHr))
        {
            return;
        }

        const auto matches = callback.Matches();
        state.Require(matches.size() == 1u,
                      std::format(L"Search for preference {} expected 1 match, got {}.",
                                  std::wstring(configuredPreference.begin(), configuredPreference.end()),
                                  matches.size()));
        state.Require(FindRecordedSearchMatch(matches, L"pref.txt") != nullptr,
                      std::format(L"Search for preference {} should match pref.txt.", std::wstring(configuredPreference.begin(), configuredPreference.end())));

        const auto progressSnapshots            = callback.ProgressSnapshots();
        const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
        state.Require(
            completed != nullptr,
            std::format(L"Search for preference {} missing completed progress.", std::wstring(configuredPreference.begin(), configuredPreference.end())));
        if (completed != nullptr)
        {
            state.Require(completed->backend == expectedBackend,
                          std::format(L"Search for preference {} reported an unexpected backend.",
                                      std::wstring(configuredPreference.begin(), configuredPreference.end())));
            const bool degradedNoIndex = (completed->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u;
            state.Require(degradedNoIndex == expectDegraded,
                          std::format(L"Search for preference {} returned an unexpected degraded-no-index warning state.",
                                      std::wstring(configuredPreference.begin(), configuredPreference.end())));
            const bool serviceUnavailable = (completed->warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) != 0u;
            state.Require(serviceUnavailable == expectServiceUnavailable,
                          std::format(L"Search for preference {} returned an unexpected service-unavailable warning state.",
                                      std::wstring(configuredPreference.begin(), configuredPreference.end())));
        }
    };

    runPreferenceCase("auto", FALSE, FILESYSTEM_SEARCH_BACKEND_INDEX, false, true, false);
    runPreferenceCase("service", TRUE, FILESYSTEM_SEARCH_BACKEND_INDEX, false, true, false);
    runPreferenceCase("local-index", TRUE, FILESYSTEM_SEARCH_BACKEND_INDEX, false, false, false);
    runPreferenceCase("scan", TRUE, FILESYSTEM_SEARCH_BACKEND_SCAN, false, false, false);
    runPreferenceCase("auto", FALSE, FILESYSTEM_SEARCH_BACKEND_INDEX, false, true, true);
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_scan_follow_symlink_loop_guard",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! info || ! search)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"scan\"}");
    state.Require(SUCCEEDED(setHr), std::format(L"Failed to configure scan backend for symlink-loop guard. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_scan_follow_symlink_loop_guard", caseRoot),
                  L"Failed to prepare search_scan_follow_symlink_loop_guard root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"needle.txt", "loop guard"), L"Failed to create needle.txt.");

    const std::filesystem::path loopPath = caseRoot / L"loop";
    const bool linkCreated               = TryCreateDirectorySymlink(loopPath, caseRoot);
    if (! linkCreated)
    {
        const DWORD err = ::GetLastError();
        if (err == ERROR_PRIVILEGE_NOT_HELD || err == ERROR_ACCESS_DENIED || err == ERROR_INVALID_PARAMETER)
        {
            Debug::Warning(L"CompareSelfTest: skipping search symlink-loop guard test (CreateSymbolicLinkW failed: {}).", err);
            return true;
        }

        state.Require(false, std::format(L"CreateSymbolicLinkW failed unexpectedly for search loop guard: {}.", err));
        return false;
    }

    std::wstring rootText    = caseRoot.wstring();
    std::wstring namePattern = L"needle.txt";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES | FILESYSTEM_SEARCH_FOLLOW_SYMLINKS);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    RecordingSearchCallback callback(RecordingSearchCallback::Mode::Success, 64u);
    const HRESULT hr = search->Search(&query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Follow-symlink scan search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const auto matches = callback.Matches();
    state.Require(matches.size() == 1u, std::format(L"Follow-symlink scan search expected 1 match, got {}.", matches.size()));
    state.Require(FindRecordedSearchMatch(matches, L"needle.txt") != nullptr, L"Follow-symlink scan search missing needle.txt.");

    const auto progressSnapshots            = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Follow-symlink scan search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SCAN, L"Follow-symlink scan search should report scan backend.");
        state.Require(completed->scannedDirectories < 8u,
                      std::format(L"Follow-symlink scan search visited too many directories: {}.", completed->scannedDirectories));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_fallback_follow_symlink_loop_guard",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_fallback_follow_symlink_loop_guard", caseRoot),
                  L"Failed to prepare search_fallback_follow_symlink_loop_guard root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"needle.txt", "fallback loop guard"), L"Failed to create fallback needle.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path loopPath = caseRoot / L"loop";
    const bool linkCreated               = TryCreateDirectorySymlink(loopPath, caseRoot);
    if (! linkCreated)
    {
        const DWORD err = ::GetLastError();
        if (err == ERROR_PRIVILEGE_NOT_HELD || err == ERROR_ACCESS_DENIED || err == ERROR_INVALID_PARAMETER)
        {
            Debug::Warning(L"CompareSelfTest: skipping fallback search symlink-loop guard test (CreateSymbolicLinkW failed: {}).", err);
            return true;
        }

        state.Require(false, std::format(L"CreateSymbolicLinkW failed unexpectedly for fallback search loop guard: {}.", err));
        return false;
    }

    std::wstring rootText    = caseRoot.wstring();
    std::wstring namePattern = L"needle.txt";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES | FILESYSTEM_SEARCH_FOLLOW_SYMLINKS);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    RecordingSearchCallback callback(RecordingSearchCallback::Mode::Success, 64u);
    const HRESULT hr = SearchFallbackEngine::Execute(baseFs.get(), &query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Fallback follow-symlink search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const auto matches = callback.Matches();
    state.Require(matches.size() == 1u, std::format(L"Fallback follow-symlink search expected 1 match, got {}.", matches.size()));
    state.Require(FindRecordedSearchMatch(matches, L"needle.txt") != nullptr, L"Fallback follow-symlink search missing needle.txt.");

    const auto progressSnapshots            = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Fallback follow-symlink search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SCAN, L"Fallback follow-symlink search should report scan backend.");
        state.Require(completed->scannedDirectories < 8u,
                      std::format(L"Fallback follow-symlink search visited too many directories: {}.", completed->scannedDirectories));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_fallback_single_pass_large_directory",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_fallback_single_pass_large_directory", caseRoot),
                  L"Failed to prepare search_fallback_single_pass_large_directory root.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kEntryCount = 4096u;

    ReadDirectoryTestBehavior behavior{};
    behavior.targetPath = caseRoot;
    behavior.rawEntries.reserve(kEntryCount);
    for (size_t index = 0u; index < kEntryCount; ++index)
    {
        behavior.rawEntries.push_back({std::format(L"hit_{:04}.txt", index), FILE_ATTRIBUTE_ARCHIVE, false});
    }

    const wil::com_ptr<IFileSystem> wrappedFs = CreateReadDirectoryBehaviorFileSystem(baseFs, behavior);
    state.Require(static_cast<bool>(wrappedFs), L"Failed to create synthetic large-directory ReadDirectoryInfo wrapper.");
    if (! wrappedFs)
    {
        return false;
    }

    const std::wstring rootText    = caseRoot.wstring();
    const std::wstring namePattern = L"hit_*.txt";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = FILESYSTEM_SEARCH_INCLUDE_FILES;
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    RecordingSearchCallback callback;
    const auto started       = std::chrono::steady_clock::now();
    const HRESULT hr         = SearchFallbackEngine::Execute(wrappedFs.get(), &query, &callback, nullptr);
    const uint64_t elapsedUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());

    const auto matches = callback.Matches();
    Debug::Perf::Emit(L"compare.selftest.local_search_fallback_buffer_enumeration_us",
                      std::format(L"entries={}", kEntryCount),
                      elapsedUs,
                      kEntryCount,
                      static_cast<uint64_t>(matches.size()),
                      hr);
    AppendCompareSelfTestTraceLine(std::format(L"Case: local_search_fallback_single_pass_large_directory elapsedUs={} entries={} matches={} hr=0x{:08X}",
                                               elapsedUs,
                                               kEntryCount,
                                               matches.size(),
                                               static_cast<unsigned long>(hr)));

    state.Require(SUCCEEDED(hr), std::format(L"Fallback large-directory search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(matches.size() == kEntryCount, std::format(L"Fallback large-directory search expected {} matches, got {}.", kEntryCount, matches.size()));
    state.Require(FindRecordedSearchMatch(matches, L"hit_0000.txt") != nullptr, L"Fallback large-directory search missing first entry.");
    state.Require(FindRecordedSearchMatch(matches, L"hit_4095.txt") != nullptr, L"Fallback large-directory search missing last entry.");

    const auto progressSnapshots            = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Fallback large-directory search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SCAN, L"Fallback large-directory search should report scan backend.");
        state.Require(completed->scannedDirectories == 1u,
                      std::format(L"Fallback large-directory search scannedDirectories mismatch. got={}", completed->scannedDirectories));
        state.Require(completed->scannedFiles == kEntryCount,
                      std::format(L"Fallback large-directory search scannedFiles mismatch. got={}", completed->scannedFiles));
        state.Require(completed->matchedEntries == kEntryCount,
                      std::format(L"Fallback large-directory search matchedEntries mismatch. got={}", completed->matchedEntries));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_fallback_malformed_entry_skips_with_warning",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_fallback_malformed_entry_skips_with_warning", caseRoot),
                  L"Failed to prepare search_fallback_malformed_entry_skips_with_warning root.");
    if (! state.failure.empty())
    {
        return false;
    }

    ReadDirectoryTestBehavior behavior{};
    behavior.targetPath = caseRoot;
    behavior.rawEntries = {
        {L"good-before.txt", FILE_ATTRIBUTE_ARCHIVE, false},
        {L"bad-entry.txt", FILE_ATTRIBUTE_ARCHIVE, true},
        {L"good-after.txt", FILE_ATTRIBUTE_ARCHIVE, false},
    };

    const wil::com_ptr<IFileSystem> wrappedFs = CreateReadDirectoryBehaviorFileSystem(baseFs, behavior);
    state.Require(static_cast<bool>(wrappedFs), L"Failed to create malformed-entry ReadDirectoryInfo wrapper.");
    if (! wrappedFs)
    {
        return false;
    }

    const std::wstring rootText    = caseRoot.wstring();
    const std::wstring namePattern = L"*.txt";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = FILESYSTEM_SEARCH_INCLUDE_FILES;
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    RecordingSearchCallback callback;
    const HRESULT hr = SearchFallbackEngine::Execute(wrappedFs.get(), &query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Fallback malformed-entry search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const auto matches = callback.Matches();
    state.Require(matches.size() == 2u, std::format(L"Fallback malformed-entry search expected 2 matches, got {}.", matches.size()));
    state.Require(FindRecordedSearchMatch(matches, L"good-before.txt") != nullptr, L"Fallback malformed-entry search missing good-before.txt.");
    state.Require(FindRecordedSearchMatch(matches, L"good-after.txt") != nullptr, L"Fallback malformed-entry search missing good-after.txt.");
    state.Require(FindRecordedSearchMatch(matches, L"bad-entry.txt") == nullptr, L"Fallback malformed-entry search should skip the bad entry.");

    const auto progressSnapshots            = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Fallback malformed-entry search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SCAN, L"Fallback malformed-entry search should report scan backend.");
        state.Require((completed->warningFlags & FILESYSTEM_SEARCH_WARNING_OVERFLOW) != 0u,
                      L"Fallback malformed-entry search should report an overflow/corrupt-entry warning.");
        state.Require(completed->scannedFiles == 2u, std::format(L"Fallback malformed-entry search scannedFiles mismatch. got={}", completed->scannedFiles));
        state.Require(completed->matchedEntries == 2u,
                      std::format(L"Fallback malformed-entry search matchedEntries mismatch. got={}", completed->matchedEntries));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_scan_wide_tree_parallel_walk_name_only",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
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

    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! search)
    {
        return false;
    }

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_scan_wide_tree_parallel_walk_name_only", caseRoot),
                  L"Failed to prepare search_scan_wide_tree_parallel_walk_name_only root.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kTopDirectoryCount   = 256u;
    constexpr size_t kChildDirectoryCount = 20u;
    constexpr size_t kMatchStride         = 5u;

    for (size_t topIndex = 0; topIndex < kTopDirectoryCount; ++topIndex)
    {
        const std::filesystem::path topDirectory = caseRoot / std::format(L"top_{:03}", topIndex);
        state.Require(SelfTest::EnsureDirectory(topDirectory), std::format(L"Failed to create '{}'.", topDirectory.filename().native()));
        if (! state.failure.empty())
        {
            return false;
        }

        for (size_t childIndex = 0; childIndex < kChildDirectoryCount; ++childIndex)
        {
            const std::filesystem::path childDirectory = topDirectory / std::format(L"child_{:03}", childIndex);
            state.Require(SelfTest::EnsureDirectory(childDirectory), std::format(L"Failed to create '{}'.", childDirectory.filename().native()));
            if (! state.failure.empty())
            {
                return false;
            }

            const bool isMatch                   = ((childIndex + 1u) % kMatchStride) == 0u;
            const std::filesystem::path filePath = childDirectory / (isMatch ? std::format(L"hit_{:03}_{:03}.cpp", topIndex, childIndex)
                                                                             : std::format(L"ignore_{:03}_{:03}.bin", topIndex, childIndex));
            state.Require(SelfTest::WriteTextFile(filePath, isMatch ? "match" : "ignore"),
                          std::format(L"Failed to create '{}'.", filePath.filename().native()));
            if (! state.failure.empty())
            {
                return false;
            }
        }
    }

    uint64_t expectedDirectories = 1u;
    uint64_t expectedFiles       = 0u;
    uint64_t expectedMatches     = 0u;
    std::error_code walkError;
    for (std::filesystem::recursive_directory_iterator it(caseRoot, walkError), end; ! walkError && it != end; it.increment(walkError))
    {
        const std::filesystem::path currentPath = it->path();
        if (it->is_directory(walkError))
        {
            ++expectedDirectories;
            continue;
        }

        if (! it->is_regular_file(walkError))
        {
            continue;
        }

        ++expectedFiles;
        const std::wstring filename = currentPath.filename().native();
        if (filename.rfind(L"hit_", 0u) == 0u && currentPath.extension() == L".cpp")
        {
            ++expectedMatches;
        }
    }
    state.Require(! walkError, std::format(L"Failed to enumerate the prepared wide-tree case root. ec={}.", walkError.value()));
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring rootText    = caseRoot.wstring();
    std::wstring namePattern = L"hit_*.cpp";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    struct WideTreeRunResult final
    {
        uint64_t elapsedUs          = 0u;
        uint64_t scannedDirectories = 0u;
        uint64_t scannedFiles       = 0u;
        uint64_t matchedEntries     = 0u;
        uint64_t matchCount         = 0u;
    };

    const auto runWideTreeSearch = [&](unsigned int configuredWorkers, WideTreeRunResult& outResult) noexcept
    {
        const std::string configurationJson = std::format("{{\"searchBackendPreference\":\"scan\",\"searchMaxDirectoryWalkers\":{}}}", configuredWorkers);
        const HRESULT setHr                 = info->SetConfiguration(configurationJson.c_str());
        state.Require(SUCCEEDED(setHr),
                      std::format(L"Failed to configure scan backend + walkers={} for wide-tree search. hr=0x{:08X}",
                                  configuredWorkers,
                                  static_cast<unsigned long>(setHr)));
        if (FAILED(setHr))
        {
            return;
        }

        const char* configuration = nullptr;
        const HRESULT getHr       = info->GetConfiguration(&configuration);
        state.Require(
            SUCCEEDED(getHr) && configuration != nullptr,
            std::format(L"Wide-tree search configuration roundtrip failed for walkers={}. hr=0x{:08X}", configuredWorkers, static_cast<unsigned long>(getHr)));
        if (configuration != nullptr)
        {
            const std::string expectedField = std::format("\"searchMaxDirectoryWalkers\":{}", configuredWorkers);
            state.Require(std::string_view(configuration).find(expectedField) != std::string_view::npos,
                          std::format(L"Configuration JSON missing searchMaxDirectoryWalkers={} field.", configuredWorkers));
        }

        RecordingSearchCallback callback;
        const auto started = std::chrono::steady_clock::now();
        const HRESULT hr   = search->Search(&query, &callback, nullptr);
        const uint64_t elapsedUs =
            static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());

        const auto matches                      = callback.Matches();
        const auto progressSnapshots            = callback.ProgressSnapshots();
        const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
        const uint64_t scannedDirectories       = completed != nullptr ? completed->scannedDirectories : 0u;
        const uint64_t scannedFiles             = completed != nullptr ? completed->scannedFiles : 0u;
        const uint64_t matchedEntries           = completed != nullptr ? completed->matchedEntries : 0u;
        const std::wstring perfDetail           = std::format(L"{}|workers={}", caseRoot.native(), configuredWorkers);
        Debug::Perf::Emit(
            L"compare.selftest.local_search_scan_wide_tree_workers_us", perfDetail, elapsedUs, scannedDirectories, static_cast<uint64_t>(matches.size()), hr);
        if (configuredWorkers == 4u)
        {
            Debug::Perf::Emit(L"compare.selftest.local_search_scan_wide_tree_us",
                              caseRoot.native(),
                              elapsedUs,
                              scannedDirectories,
                              static_cast<uint64_t>(matches.size()),
                              hr);
        }
        AppendCompareSelfTestTraceLine(
            std::format(L"Case: local_search_scan_wide_tree_parallel_walk_name_only walkers={} elapsedUs={} scannedDirectories={} scannedFiles={} "
                        L"matches={} expectedDirectories={} expectedFiles={} expectedMatches={}",
                        configuredWorkers,
                        elapsedUs,
                        scannedDirectories,
                        scannedFiles,
                        matches.size(),
                        expectedDirectories,
                        expectedFiles,
                        expectedMatches));

        state.Require(SUCCEEDED(hr),
                      std::format(L"Wide-tree scan search failed for walkers={}. hr=0x{:08X}", configuredWorkers, static_cast<unsigned long>(hr)));
        state.Require(completed != nullptr, std::format(L"Wide-tree scan search missing completed progress for walkers={}.", configuredWorkers));
        if (completed != nullptr)
        {
            state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SCAN,
                          std::format(L"Wide-tree scan search should report the scan backend for walkers={}.", configuredWorkers));
            state.Require(
                scannedDirectories == expectedDirectories,
                std::format(
                    L"Wide-tree scan search expected {} directories, got {} for walkers={}.", expectedDirectories, scannedDirectories, configuredWorkers));
            state.Require(scannedFiles == expectedFiles,
                          std::format(L"Wide-tree scan search expected {} files, got {} for walkers={}.", expectedFiles, scannedFiles, configuredWorkers));
            state.Require(
                matchedEntries == expectedMatches,
                std::format(L"Wide-tree scan search expected {} matchedEntries, got {} for walkers={}.", expectedMatches, matchedEntries, configuredWorkers));
        }

        state.Require(matches.size() == expectedMatches,
                      std::format(L"Wide-tree scan search expected {} matches, got {} for walkers={}.", expectedMatches, matches.size(), configuredWorkers));
        for (const auto& match : matches)
        {
            state.Require(match.displayName.rfind(L"hit_", 0u) == 0u && std::filesystem::path(match.displayName).extension() == L".cpp",
                          std::format(L"Wide-tree scan search returned an unexpected match '{}' for walkers={}.", match.displayName, configuredWorkers));
            if (! state.failure.empty())
            {
                return;
            }
        }

        outResult.elapsedUs          = elapsedUs;
        outResult.scannedDirectories = scannedDirectories;
        outResult.scannedFiles       = scannedFiles;
        outResult.matchedEntries     = matchedEntries;
        outResult.matchCount         = static_cast<uint64_t>(matches.size());
    };

    WideTreeRunResult workers1{};
    runWideTreeSearch(1u, workers1);
    if (! state.failure.empty())
    {
        return false;
    }

    WideTreeRunResult workers4{};
    runWideTreeSearch(4u, workers4);
    if (! state.failure.empty())
    {
        return false;
    }

    WideTreeRunResult workers8{};
    runWideTreeSearch(8u, workers8);
    if (! state.failure.empty())
    {
        return false;
    }

    AppendCompareSelfTestTraceLine(std::format(L"Case: local_search_scan_wide_tree_parallel_walk_name_only summary workers1={}us workers4={}us workers8={}us",
                                               workers1.elapsedUs,
                                               workers4.elapsedUs,
                                               workers8.elapsedUs));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_scan_deferred_path_materialization",
                  [&](SelfTest::CaseState& state) noexcept
{
    constexpr uint32_t kFileCount = 768u;

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! info || ! search)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"scan\",\"searchMaxDirectoryWalkers\":4}");
    state.Require(SUCCEEDED(setHr),
                  std::format(L"Failed to configure scan backend for deferred path materialization. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_deferred_path_materialization", caseRoot),
                  L"Failed to prepare search_deferred_path_materialization root.");
    for (uint32_t index = 0u; index < kFileCount; ++index)
    {
        const std::filesystem::path filePath = caseRoot / std::format(L"miss_{:04}.dat", index);
        state.Require(SelfTest::WriteTextFile(filePath, "miss"), std::format(L"Failed to create {}.", filePath.filename().wstring()));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const std::filesystem::path perfPath = SelfTest::GetPerfArtifactPath(L"perf_metrics.jsonl");
    std::error_code perfEc;
    const std::uintmax_t perfStartOffset = std::filesystem::exists(perfPath, perfEc) ? std::filesystem::file_size(perfPath, perfEc) : 0u;

    std::wstring rootText    = caseRoot.wstring();
    std::wstring namePattern = L"hit_*.txt";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    RecordingSearchCallback callback;
    const auto started       = std::chrono::steady_clock::now();
    const HRESULT hr         = search->Search(&query, &callback, nullptr);
    const uint64_t elapsedUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());

    const auto progressSnapshots            = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(SUCCEEDED(hr), std::format(L"Deferred materialization scan failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(callback.Matches().empty(), L"Deferred materialization scan should return no matches.");
    state.Require(completed != nullptr, L"Deferred materialization scan missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SCAN, L"Deferred materialization scan should use the scan backend.");
        state.Require(completed->scannedFiles == kFileCount,
                      std::format(L"Deferred materialization scan expected {} scanned files, got {}.", kFileCount, completed->scannedFiles));
        state.Require(completed->matchedEntries == 0u, L"Deferred materialization scan should report zero matched entries.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const auto parseUnsignedField = [](std::string_view line, std::string_view field) noexcept -> std::optional<uint64_t>
    {
        const size_t fieldPos = line.find(field);
        if (fieldPos == std::string_view::npos)
        {
            return std::nullopt;
        }

        size_t cursor  = fieldPos + field.size();
        uint64_t value = 0u;
        bool anyDigit  = false;
        while (cursor < line.size() && line[cursor] >= '0' && line[cursor] <= '9')
        {
            anyDigit             = true;
            const uint64_t digit = static_cast<uint64_t>(line[cursor] - '0');
            if (value > ((std::numeric_limits<uint64_t>::max)() - digit) / 10u)
            {
                return std::nullopt;
            }

            value = (value * 10u) + digit;
            ++cursor;
        }

        return anyDigit ? std::optional<uint64_t>(value) : std::nullopt;
    };

    std::optional<uint64_t> materializedPathPairs;
    std::optional<uint64_t> scannedItems;
    std::ifstream perfFile(perfPath, std::ios::binary);
    if (perfFile)
    {
        if (perfStartOffset != 0u)
        {
            perfFile.seekg(static_cast<std::streamoff>(perfStartOffset), std::ios::beg);
        }

        std::string line;
        while (std::getline(perfFile, line))
        {
            if (line.find("\"metric\":\"FileSystem.Search.EntryPathMaterialized\"") == std::string::npos)
            {
                continue;
            }

            materializedPathPairs = parseUnsignedField(line, "\"value0\":");
            scannedItems          = parseUnsignedField(line, "\"value1\":");
        }
    }

    state.Require(materializedPathPairs.has_value(), L"Deferred materialization scan missing FileSystem.Search.EntryPathMaterialized perf metric.");
    if (materializedPathPairs.has_value())
    {
        state.Require(materializedPathPairs.value() == 0u,
                      std::format(L"Deferred materialization should not materialize nonmatching file paths. got={}.", materializedPathPairs.value()));
    }
    if (scannedItems.has_value())
    {
        state.Require(scannedItems.value() >= kFileCount,
                      std::format(L"Deferred materialization metric should include scanned items. value1={} fileCount={}.", scannedItems.value(), kFileCount));
    }

    Debug::Perf::Emit(L"compare.selftest.local_search_scan_deferred_path_materialization_us",
                      caseRoot.native(),
                      elapsedUs,
                      materializedPathPairs.value_or((std::numeric_limits<uint64_t>::max)()),
                      kFileCount,
                      hr);

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_name_wildcard_recursive",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_name_wildcard_recursive", caseRoot), L"Failed to prepare search_name_wildcard_recursive root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"sub"), L"Failed to create sub directory.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"a.txt", "A"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"b.log", "B"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"sub" / L"c.txt", "C"), L"Failed to create sub\\c.txt.");

    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(baseFs, search), L"Local file system plugin missing IFileSystemSearch.");

    std::wstring rootText    = caseRoot.wstring();
    std::wstring namePattern = L"*.txt";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    RecordingSearchCallback callback;
    const HRESULT hr = search->Search(&query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Wildcard search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    const auto matches = callback.Matches();
    state.Require(matches.size() == 2u, std::format(L"Wildcard search expected 2 matches, got {}.", matches.size()));

    const RecordedSearchMatch* aMatch = FindRecordedSearchMatch(matches, L"a.txt");
    state.Require(aMatch != nullptr, L"Wildcard search missing a.txt.");
    if (aMatch != nullptr)
    {
        state.Require((aMatch->matchedBy & FILESYSTEM_SEARCH_MATCH_SOURCE_NAME) != 0u, L"a.txt should be reported as a name match.");
    }

    const RecordedSearchMatch* cMatch = FindRecordedSearchMatch(matches, L"c.txt");
    state.Require(cMatch != nullptr, L"Wildcard search missing c.txt.");
    if (cMatch != nullptr)
    {
        state.Require(cMatch->relativePath == L"sub\\c.txt", L"c.txt should preserve its relative path.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_name_windows_filesystem_case_parity",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_name_windows_filesystem_case_parity", caseRoot),
                  L"Failed to prepare search_name_windows_filesystem_case_parity root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring accentName     = std::wstring(L"\x00C9") + L"cole.txt";
    const std::wstring accentQuery    = std::wstring(L"\x00E9") + L"COLE.txt";
    const std::wstring composedName   = std::wstring(L"\x00E9") + L".txt";
    const std::wstring decomposedName = std::wstring(L"e") + std::wstring(L"\x0301") + L".txt";

    state.Require(SelfTest::WriteTextFile(caseRoot / accentName, "accent"), L"Failed to create accented test file.");
    state.Require(SelfTest::WriteTextFile(caseRoot / composedName, "composed"), L"Failed to create composed Unicode test file.");
    state.Require(SelfTest::WriteTextFile(caseRoot / decomposedName, "decomposed"), L"Failed to create decomposed Unicode test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(OrdinalString::EqualsNoCase(accentName, accentQuery),
                  L"Windows ordinal ignore-case comparison should treat accented case variants as equal.");
    state.Require(! OrdinalString::EqualsNoCase(composedName, decomposedName),
                  L"Windows ordinal ignore-case comparison should keep composed and decomposed names distinct.");
    state.Require(OrdinalString::FoldCaseInvariant(accentName) == OrdinalString::FoldCaseInvariant(accentQuery),
                  L"Invariant case-fold keys should match for accented case variants.");
    state.Require(OrdinalString::FoldCaseInvariant(composedName) != OrdinalString::FoldCaseInvariant(decomposedName),
                  L"Invariant case-fold keys should remain distinct for composed and decomposed names.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(baseFs, search), L"Local file system plugin missing IFileSystemSearch.");
    if (! search)
    {
        return false;
    }

    const std::wstring rootText = caseRoot.wstring();
    {
        FileSystemSearchQuery query{};
        query.sizeBytes   = sizeof(FileSystemSearchQuery);
        query.rootPath    = rootText.c_str();
        query.namePattern = accentQuery.c_str();
        query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
        query.nameMode    = FILESYSTEM_SEARCH_NAME_LITERAL;
        query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

        RecordingSearchCallback callback;
        const HRESULT pluginHr = search->Search(&query, &callback, nullptr);
        state.Require(SUCCEEDED(pluginHr), std::format(L"Accented plugin search failed. hr=0x{:08X}", static_cast<unsigned long>(pluginHr)));
        if (FAILED(pluginHr))
        {
            return false;
        }
        state.Require(callback.Matches().size() == 1u && callback.Matches().front().displayName == accentName,
                      L"Accented plugin search should match the original file name.");
    }

    {
        FileSystemSearchQuery query{};
        query.sizeBytes   = sizeof(FileSystemSearchQuery);
        query.rootPath    = rootText.c_str();
        query.namePattern = composedName.c_str();
        query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
        query.nameMode    = FILESYSTEM_SEARCH_NAME_LITERAL;
        query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

        RecordingSearchCallback callback;
        const HRESULT pluginHr = search->Search(&query, &callback, nullptr);
        state.Require(SUCCEEDED(pluginHr), std::format(L"Composed plugin search failed. hr=0x{:08X}", static_cast<unsigned long>(pluginHr)));
        if (FAILED(pluginHr))
        {
            return false;
        }
        state.Require(callback.Matches().size() == 1u && callback.Matches().front().displayName == composedName,
                      L"Composed plugin search should return only the composed file name.");
    }

    {
        FileSystemSearchQuery query{};
        query.sizeBytes   = sizeof(FileSystemSearchQuery);
        query.rootPath    = rootText.c_str();
        query.namePattern = decomposedName.c_str();
        query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
        query.nameMode    = FILESYSTEM_SEARCH_NAME_LITERAL;
        query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

        RecordingSearchCallback callback;
        const HRESULT pluginHr = search->Search(&query, &callback, nullptr);
        state.Require(SUCCEEDED(pluginHr), std::format(L"Decomposed plugin search failed. hr=0x{:08X}", static_cast<unsigned long>(pluginHr)));
        if (FAILED(pluginHr))
        {
            return false;
        }
        state.Require(callback.Matches().size() == 1u && callback.Matches().front().displayName == decomposedName,
                      L"Decomposed plugin search should return only the decomposed file name.");
    }

    LocalSearchIndexCore::Repository repository({
        .snapshotRootDirectory = (caseRoot / L"store").wstring(),
    });

    LocalSearchIndexCore::QueryStats accentStats{};
    std::vector<LocalSearchIndexCore::Candidate> accentCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), accentQuery, FILESYSTEM_SEARCH_NAME_LITERAL, accentStats, accentCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Accented case-parity query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(CollectIndexedCandidateNames(accentCandidates) == std::vector<std::wstring>{accentName},
                  L"Accented case-parity query should match the original file name.");

    LocalSearchIndexCore::QueryStats composedStats{};
    std::vector<LocalSearchIndexCore::Candidate> composedCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), composedName, FILESYSTEM_SEARCH_NAME_LITERAL, composedStats, composedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Composed Unicode query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(CollectIndexedCandidateNames(composedCandidates) == std::vector<std::wstring>{composedName},
                  L"Composed Unicode query should return only the composed file name.");

    LocalSearchIndexCore::QueryStats decomposedStats{};
    std::vector<LocalSearchIndexCore::Candidate> decomposedCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), decomposedName, FILESYSTEM_SEARCH_NAME_LITERAL, decomposedStats, decomposedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Decomposed Unicode query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(CollectIndexedCandidateNames(decomposedCandidates) == std::vector<std::wstring>{decomposedName},
                  L"Decomposed Unicode query should return only the decomposed file name.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_content_literal",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_content_literal", caseRoot), L"Failed to prepare search_content_literal root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"plain.txt", "alpha beta gamma"), L"Failed to create plain.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"other.txt", "delta"), L"Failed to create other.txt.");

    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(baseFs, search), L"Local file system plugin missing IFileSystemSearch.");

    std::wstring rootText       = caseRoot.wstring();
    std::wstring contentPattern = L"beta";

    FileSystemSearchQuery query{};
    query.sizeBytes      = sizeof(FileSystemSearchQuery);
    query.rootPath       = rootText.c_str();
    query.contentPattern = contentPattern.c_str();
    query.flags          = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES | FILESYSTEM_SEARCH_WANT_SNIPPETS);
    query.nameMode       = FILESYSTEM_SEARCH_NAME_DISABLED;
    query.contentMode    = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;
    query.maxSnippetCharacters = 32u;

    RecordingSearchCallback callback;
    const HRESULT hr = search->Search(&query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Content search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    const auto matches = callback.Matches();
    state.Require(matches.size() == 1u, std::format(L"Content search expected 1 match, got {}.", matches.size()));
    const RecordedSearchMatch* plainMatch = FindRecordedSearchMatch(matches, L"plain.txt");
    state.Require(plainMatch != nullptr, L"Content search missing plain.txt.");
    if (plainMatch != nullptr)
    {
        state.Require((plainMatch->matchedBy & FILESYSTEM_SEARCH_MATCH_SOURCE_CONTENT) != 0u, L"plain.txt should be reported as a content match.");
        state.Require(plainMatch->previewText.find(L"beta") != std::wstring::npos, L"plain.txt snippet should include the matched token.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_name_and_content_and_semantics",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_name_and_content", caseRoot), L"Failed to prepare search_name_and_content root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"foo.txt", "hello"), L"Failed to create foo.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"foo.log", "hello"), L"Failed to create foo.log.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"bar.txt", "nope"), L"Failed to create bar.txt.");

    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(baseFs, search), L"Local file system plugin missing IFileSystemSearch.");

    std::wstring rootText       = caseRoot.wstring();
    std::wstring namePattern    = L"*.txt";
    std::wstring contentPattern = L"hello";

    FileSystemSearchQuery query{};
    query.sizeBytes      = sizeof(FileSystemSearchQuery);
    query.rootPath       = rootText.c_str();
    query.namePattern    = namePattern.c_str();
    query.contentPattern = contentPattern.c_str();
    query.flags          = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode       = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode    = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;

    RecordingSearchCallback callback;
    const HRESULT hr = search->Search(&query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Name+content search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    const auto matches = callback.Matches();
    state.Require(matches.size() == 1u, std::format(L"Name+content search expected 1 match, got {}.", matches.size()));
    const RecordedSearchMatch* fooMatch = FindRecordedSearchMatch(matches, L"foo.txt");
    state.Require(fooMatch != nullptr, L"Name+content search should only match foo.txt.");
    if (fooMatch != nullptr)
    {
        state.Require((fooMatch->matchedBy & FILESYSTEM_SEARCH_MATCH_SOURCE_NAME) != 0u, L"foo.txt should include the name match source bit.");
        state.Require((fooMatch->matchedBy & FILESYSTEM_SEARCH_MATCH_SOURCE_CONTENT) != 0u, L"foo.txt should include the content match source bit.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_invalid_query_rejected",
                  [&](SelfTest::CaseState& state) noexcept
{
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(baseFs, search), L"Local file system plugin missing IFileSystemSearch.");

    std::wstring rootText       = root.wstring();
    std::wstring contentPattern = L"anything";
    RecordingSearchCallback callback;

    FileSystemSearchQuery disabledQuery{};
    disabledQuery.sizeBytes   = sizeof(FileSystemSearchQuery);
    disabledQuery.rootPath    = rootText.c_str();
    disabledQuery.flags       = FILESYSTEM_SEARCH_INCLUDE_FILES;
    disabledQuery.nameMode    = FILESYSTEM_SEARCH_NAME_DISABLED;
    disabledQuery.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    HRESULT hr = search->Search(&disabledQuery, &callback, nullptr);
    state.Require(hr == E_INVALIDARG, std::format(L"Disabled search query should be rejected with E_INVALIDARG. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    FileSystemSearchQuery invalidSizeQuery{};
    invalidSizeQuery.sizeBytes   = sizeof(FileSystemSearchQuery) - sizeof(uint32_t);
    invalidSizeQuery.rootPath    = rootText.c_str();
    invalidSizeQuery.namePattern = L"*";
    invalidSizeQuery.flags       = FILESYSTEM_SEARCH_INCLUDE_FILES;
    invalidSizeQuery.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    invalidSizeQuery.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    hr = search->Search(&invalidSizeQuery, &callback, nullptr);
    state.Require(hr == E_INVALIDARG,
                  std::format(L"Invalid sizeBytes search query should be rejected with E_INVALIDARG. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    FileSystemSearchQuery invalidContentQuery{};
    invalidContentQuery.sizeBytes      = sizeof(FileSystemSearchQuery);
    invalidContentQuery.rootPath       = rootText.c_str();
    invalidContentQuery.contentPattern = contentPattern.c_str();
    invalidContentQuery.flags          = FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES;
    invalidContentQuery.nameMode       = FILESYSTEM_SEARCH_NAME_DISABLED;
    invalidContentQuery.contentMode    = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;

    hr = search->Search(&invalidContentQuery, &callback, nullptr);
    state.Require(hr == E_INVALIDARG,
                  std::format(L"Content-only directory query should be rejected with E_INVALIDARG. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_snapshot_reload",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_snapshot_reload", caseRoot), L"Failed to prepare index_core_snapshot_reload root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"sub"), L"Failed to create sub directory.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"sub" / L"beta.txt", "beta"), L"Failed to create sub\\beta.txt.");

    LocalSearchIndexCore::Repository repository;
    LocalSearchIndexCore::SupportInfo support{};
    const HRESULT probeHr = repository.ProbePath(caseRoot.wstring(), support);
    state.Require(SUCCEEDED(probeHr), std::format(L"ProbePath failed. hr=0x{:08X}", static_cast<unsigned long>(probeHr)));
    if (FAILED(probeHr))
    {
        return false;
    }

    state.Require(support.indexable, L"Test root should be indexable on the local development volume.");
    state.Require(support.fileSystemKind == LocalSearchIndexCore::FileSystemKind::Ntfs ||
                      support.fileSystemKind == LocalSearchIndexCore::FileSystemKind::Refs,
                  L"Test root should resolve to a supported indexed filesystem (NTFS or ReFS).");

    LocalSearchIndexCore::SupportInfo fakeUnc{};
    const HRESULT fakeUncHr = repository.ProbePath(L"\\\\server\\share\\folder", fakeUnc);
    state.Require(SUCCEEDED(fakeUncHr), std::format(L"ProbePath(UNC) failed. hr=0x{:08X}", static_cast<unsigned long>(fakeUncHr)));
    if (SUCCEEDED(fakeUncHr))
    {
        state.Require(! fakeUnc.indexable, L"UNC paths should not be treated as indexable local roots.");
    }

    LocalSearchIndexCore::QueryStats coldStats{};
    std::vector<LocalSearchIndexCore::Candidate> coldCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", coldStats, coldCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Cold indexed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto coldNames = CollectIndexedCandidateNames(coldCandidates);
    state.Require(coldNames == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"}, L"Cold indexed query returned unexpected names.");
    state.Require(coldStats.usedNtfsEnumeration || coldStats.usedTraversalSeed,
                  L"Cold indexed query should seed from either NTFS USN/MFT enumeration or the user-mode traversal fallback.");
    state.Require(coldStats.snapshotSaved, L"Cold indexed query should persist a snapshot.");
    state.Require(coldStats.snapshotFileBytes != 0u, L"Cold indexed query should report a non-zero snapshot size.");
    state.Require(coldStats.estimatedMemoryBytes != 0u, L"Cold indexed query should report a non-zero memory estimate.");
    state.Require(! coldStats.snapshotPath.empty(), L"Cold indexed query should report the snapshot path.");

    const HRESULT dropHr = repository.DropCachedVolumeForTests(caseRoot.wstring());
    state.Require(SUCCEEDED(dropHr), std::format(L"DropCachedVolumeForTests failed. hr=0x{:08X}", static_cast<unsigned long>(dropHr)));

    LocalSearchIndexCore::QueryStats warmStats{};
    std::vector<LocalSearchIndexCore::Candidate> warmCandidates;
    const auto warmStart = std::chrono::steady_clock::now();
    hr                   = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", warmStats, warmCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Warm indexed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    const uint64_t warmElapsedUs = Debug::Perf::ElapsedUs(warmStart);
    Debug::Perf::EmitDurationUs(
        L"compare.selftest.local_index.snapshot_reload_us", warmElapsedUs, static_cast<uint64_t>(warmCandidates.size()), warmStats.snapshotFileBytes, hr);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(warmStats.snapshotLoaded, L"Warm indexed query should reload from snapshot.");
    if (warmStats.journalAvailable)
    {
        state.Require(! warmStats.usedNtfsEnumeration && ! warmStats.usedTraversalSeed,
                      L"Warm indexed query should not rebuild the cold seed when journal access is available.");
    }
    state.Require(warmStats.snapshotFileBytes != 0u, L"Warm indexed query should report a non-zero snapshot size.");
    state.Require(warmStats.estimatedMemoryBytes != 0u, L"Warm indexed query should report a non-zero memory estimate.");
    state.Require(! warmStats.snapshotPath.empty(), L"Warm indexed query should report the snapshot path.");
    state.Require(CollectIndexedCandidateNames(warmCandidates) == coldNames, L"Warm indexed query should match the cold candidate set.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_query_honors_cancel_check",
                  [&](SelfTest::CaseState& state) noexcept
{
    struct CancelAfterIndexedCandidates final
    {
        CancelAfterIndexedCandidates()                                               = default;
        CancelAfterIndexedCandidates(const CancelAfterIndexedCandidates&)            = delete;
        CancelAfterIndexedCandidates(CancelAfterIndexedCandidates&&)                 = delete;
        CancelAfterIndexedCandidates& operator=(const CancelAfterIndexedCandidates&) = delete;
        CancelAfterIndexedCandidates& operator=(CancelAfterIndexedCandidates&&)      = delete;

        const std::vector<LocalSearchIndexCore::Candidate>* candidates = nullptr;
        size_t cancelAfterCandidates                                   = 0u;
        std::atomic<uint32_t> checks{0u};
    };

    const auto cancelAfterCandidatesThunk = [](void* cookie) noexcept -> HRESULT
    {
        auto* state = static_cast<CancelAfterIndexedCandidates*>(cookie);
        if (state == nullptr || state->candidates == nullptr)
        {
            return E_POINTER;
        }

        state->checks.fetch_add(1u, std::memory_order_acq_rel);
        return state->candidates->size() >= state->cancelAfterCandidates ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
    };

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_query_cancel_check", caseRoot), L"Failed to prepare index_core_query_cancel_check root.");

    constexpr uint32_t kFileCount = 768u;
    for (uint32_t index = 0u; index < kFileCount; ++index)
    {
        const std::filesystem::path filePath = caseRoot / std::format(L"indexed_cancel_{:04}.txt", index);
        state.Require(SelfTest::WriteTextFile(filePath, "cancel"), std::format(L"Failed to create {}.", filePath.wstring()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::Repository repository;
    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"indexed_cancel_*.txt", seedStats, seedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Seed indexed cancel query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(seedCandidates.size() == kFileCount,
                  std::format(L"Seed indexed cancel query expected {} candidates, got {}.", kFileCount, seedCandidates.size()));
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::QueryStats cancelStats{};
    std::vector<LocalSearchIndexCore::Candidate> cancelCandidates;
    CancelAfterIndexedCandidates cancelState{};
    cancelState.candidates            = &cancelCandidates;
    cancelState.cancelAfterCandidates = 64u;
    hr                                = RunIndexedNameQuery(repository,
                                                            caseRoot.wstring(),
                                                            L"indexed_cancel_*.txt",
                                                            FILESYSTEM_SEARCH_NAME_WILDCARD,
                                                            cancelStats,
                                                            cancelCandidates,
                                                            cancelAfterCandidatesThunk,
                                                            &cancelState);
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Indexed query should honor CancelCheckFn with ERROR_CANCELLED. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(cancelState.checks.load(std::memory_order_acquire) >= 1u, L"Indexed query did not invoke the cancel check.");
    state.Require(
        cancelCandidates.size() >= cancelState.cancelAfterCandidates && cancelCandidates.size() < seedCandidates.size(),
        std::format(L"Indexed query cancellation should stop before the full result set. got={} full={}.", cancelCandidates.size(), seedCandidates.size()));

    return state.failure.empty();
});

constexpr std::wstring_view kRefsCaseName = L"local_index_core_refs_probe_and_query_if_available";
SelfTest::RunCase(options,
                  suite,
                  kRefsCaseName,
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto refsRoot = FindFirstFixedVolumeByFileSystem(L"ReFS");
    if (! refsRoot.has_value())
    {
        return state.Skip(L"No fixed ReFS volume detected on this machine.");
    }

    const std::filesystem::path refsVolumeRoot = refsRoot.value();
    const std::filesystem::path refsSuiteRoot  = refsVolumeRoot / L"RedSalamanderSelfTest";
    state.Require(SelfTest::EnsureDirectory(refsSuiteRoot), std::format(L"Failed to create ReFS self-test root under {}.", refsSuiteRoot.wstring()));
    if (! state.failure.empty())
    {
        return false;
    }

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(refsSuiteRoot, L"index_core_refs_probe_and_query", caseRoot),
                  L"Failed to prepare index_core_refs_probe_and_query root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"sub"), L"Failed to create ReFS sub directory.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create ReFS alpha.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"sub" / L"beta.txt", "beta"), L"Failed to create ReFS sub\\beta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::Repository repository;
    LocalSearchIndexCore::SupportInfo support{};
    const HRESULT probeHr = repository.ProbePath(caseRoot.wstring(), support);
    state.Require(SUCCEEDED(probeHr), std::format(L"ReFS ProbePath failed. hr=0x{:08X}", static_cast<unsigned long>(probeHr)));
    if (FAILED(probeHr))
    {
        return false;
    }

    state.Require(support.indexable, L"ReFS test root should be indexable.");
    state.Require(support.fileSystemKind == LocalSearchIndexCore::FileSystemKind::Refs, L"ReFS test root should resolve to ReFS.");

    LocalSearchIndexCore::QueryStats stats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    const HRESULT queryHr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", stats, candidates);
    state.Require(SUCCEEDED(queryHr), std::format(L"ReFS indexed query failed. hr=0x{:08X}", static_cast<unsigned long>(queryHr)));
    if (FAILED(queryHr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"ReFS indexed query returned an unexpected candidate set.");
    state.Require(stats.fileSystemKind == LocalSearchIndexCore::FileSystemKind::Refs, L"ReFS indexed query should report FileSystemKind::Refs.");
    state.Require(! stats.usedNtfsEnumeration, L"ReFS indexed query should not report NTFS enumeration.");
    state.Require(stats.snapshotSaved || stats.snapshotLoaded, L"ReFS indexed query should persist or reload a snapshot.");
    state.Require(stats.snapshotFileBytes != 0u, L"ReFS indexed query should report a non-zero snapshot size.");
    state.Require(stats.estimatedMemoryBytes != 0u, L"ReFS indexed query should report a non-zero memory estimate.");

    std::error_code cleanupEc;
    std::filesystem::remove_all(caseRoot, cleanupEc);
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_journal_replay_rename_delete_create",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_journal_replay", caseRoot), L"Failed to prepare index_core_journal_replay root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"sub"), L"Failed to create sub directory.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"sub" / L"gamma.txt", "gamma"), L"Failed to create sub\\gamma.txt.");

    LocalSearchIndexCore::Repository repository;
    LocalSearchIndexCore::QueryStats initialStats{};
    std::vector<LocalSearchIndexCore::Candidate> initialCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", initialStats, initialCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Initial indexed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (! state.failure.empty())
    {
        return false;
    }

    std::error_code ec;
    std::filesystem::rename(caseRoot / L"alpha.txt", caseRoot / L"renamed.txt", ec);
    state.Require(! ec, L"Failed to rename alpha.txt.");
    std::filesystem::remove(caseRoot / L"sub" / L"gamma.txt", ec);
    state.Require(! ec, L"Failed to delete sub\\gamma.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"delta.txt", "delta"), L"Failed to create delta.txt.");

    std::this_thread::sleep_for(std::chrono::milliseconds(150));

    LocalSearchIndexCore::QueryStats replayStats{};
    std::vector<LocalSearchIndexCore::Candidate> replayCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", replayStats, replayCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Replay indexed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(replayStats.journalReplayApplied || replayStats.usedTraversalSeed,
                  L"Replay indexed query should apply journal changes or rebuild from traversal when journal access is unavailable.");
    state.Require(CollectIndexedCandidateNames(replayCandidates) == std::vector<std::wstring>{L"delta.txt", L"renamed.txt"},
                  L"Replay indexed query returned an unexpected candidate set.");

    return state.failure.empty();
});

const auto expectedJournalDirectoryMoveNames = []() noexcept -> std::vector<std::wstring>
{
    return {
        L"inside-child.txt",
        L"moved-in-child.txt",
        L"moved-in-nested.txt",
        L"stable.txt",
        L"unzipped-child.txt",
        L"unzipped-nested.txt",
    };
};

const auto hasCandidateFullPath = [](const std::vector<LocalSearchIndexCore::Candidate>& candidates, const std::filesystem::path& expectedPath) noexcept -> bool
{
    const std::wstring expected = expectedPath.lexically_normal().wstring();
    return std::ranges::any_of(candidates, [&](const LocalSearchIndexCore::Candidate& candidate) noexcept {
        return EqualsIgnoreCase(std::filesystem::path(candidate.fullPath).lexically_normal().wstring(), expected);
    });
};

const auto hasSqliteEntryFullPath = [](const SqliteIndexStore::ReplaceVolumeRequest& volume, const std::filesystem::path& expectedPath) noexcept -> bool
{
    const std::wstring expected = expectedPath.lexically_normal().wstring();
    return std::ranges::any_of(volume.entries, [&](const SqliteIndexStore::ImportedEntry& entry) noexcept {
        return EqualsIgnoreCase(std::filesystem::path(entry.fullPath).lexically_normal().wstring(), expected);
    });
};

const auto collectSqliteStoredFileNames = [](const SqliteIndexStore::ReplaceVolumeRequest& volume) noexcept -> std::vector<std::wstring>
{
    std::vector<std::wstring> names;
    for (const SqliteIndexStore::ImportedEntry& entry : volume.entries)
    {
        if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u && EqualsIgnoreCase(std::filesystem::path(entry.name).extension().wstring(), L".txt"))
        {
            names.push_back(entry.name);
        }
    }

    std::sort(names.begin(), names.end());
    return names;
};

const auto prepareJournalDirectoryMoveFixture = [&](SelfTest::CaseState& state,
                                                    const std::filesystem::path& caseRoot,
                                                    std::wstring_view outsideCaseName,
                                                    std::filesystem::path& outsideRoot) noexcept -> bool
{
    state.Require(PrepareSearchCaseRoot(root, outsideCaseName, outsideRoot), L"Failed to prepare outside journal directory-move root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"inside-dir"), L"Failed to create inside-dir.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"inside-dir" / L"inside-child.txt", "inside"), L"Failed to create inside child.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"move-out-dir"), L"Failed to create move-out-dir.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"move-out-dir" / L"moved-out-child.txt", "out"), L"Failed to create moved-out child.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"stable.txt", "stable"), L"Failed to create stable.txt.");
    return state.failure.empty();
};

const auto applyJournalDirectoryMoveMutations =
    [&](SelfTest::CaseState& state, const std::filesystem::path& caseRoot, const std::filesystem::path& outsideRoot) noexcept -> bool
{
    const std::filesystem::path moveInSource = outsideRoot / L"move-in-source";
    state.Require(SelfTest::EnsureDirectory(moveInSource / L"nested"), L"Failed to create move-in-source nested directory.");
    state.Require(SelfTest::WriteTextFile(moveInSource / L"moved-in-child.txt", "move-in"), L"Failed to create moved-in child.");
    state.Require(SelfTest::WriteTextFile(moveInSource / L"nested" / L"moved-in-nested.txt", "move-in nested"), L"Failed to create moved-in nested child.");

    const std::filesystem::path unzipSource = outsideRoot / L"unzipped-source";
    state.Require(SelfTest::EnsureDirectory(unzipSource / L"nested"), L"Failed to create unzipped-source nested directory.");
    state.Require(SelfTest::WriteTextFile(unzipSource / L"unzipped-child.txt", "unzipped"), L"Failed to create unzipped child.");
    state.Require(SelfTest::WriteTextFile(unzipSource / L"nested" / L"unzipped-nested.txt", "unzipped nested"), L"Failed to create unzipped nested child.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::error_code ec;
    std::filesystem::rename(moveInSource, caseRoot / L"moved-in-dir", ec);
    state.Require(! ec, L"Failed to move populated directory into indexed root.");
    ec.clear();
    std::filesystem::rename(unzipSource, caseRoot / L"unzipped-dir", ec);
    state.Require(! ec, L"Failed to move unzipped directory tree into indexed root.");
    ec.clear();
    std::filesystem::rename(caseRoot / L"move-out-dir", outsideRoot / L"moved-out-dir", ec);
    state.Require(! ec, L"Failed to move tracked directory out of indexed root.");
    ec.clear();
    std::filesystem::rename(caseRoot / L"inside-dir", caseRoot / L"renamed-inside-dir", ec);
    state.Require(! ec, L"Failed to rename tracked directory inside indexed root.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::this_thread::sleep_for(std::chrono::milliseconds(250));
    return true;
};

const auto buildJournalDirectoryMoveSyntheticRecords =
    [](const std::filesystem::path& caseRoot,
       const std::filesystem::path& outsideRoot) noexcept -> std::vector<LocalSearchIndexCore::SyntheticJournalRecordForTests>
{
    return {
        {
            .idPath         = (caseRoot / L"moved-in-dir").wstring(),
            .parentPath     = outsideRoot.wstring(),
            .name           = L"move-in-source",
            .fileAttributes = FILE_ATTRIBUTE_DIRECTORY,
            .reason         = USN_REASON_RENAME_OLD_NAME,
        },
        {
            .idPath         = (caseRoot / L"moved-in-dir").wstring(),
            .parentPath     = caseRoot.wstring(),
            .name           = L"moved-in-dir",
            .fileAttributes = FILE_ATTRIBUTE_DIRECTORY,
            .reason         = USN_REASON_RENAME_NEW_NAME,
        },
        {
            .idPath         = (caseRoot / L"unzipped-dir").wstring(),
            .parentPath     = outsideRoot.wstring(),
            .name           = L"unzipped-source",
            .fileAttributes = FILE_ATTRIBUTE_DIRECTORY,
            .reason         = USN_REASON_RENAME_OLD_NAME,
        },
        {
            .idPath         = (caseRoot / L"unzipped-dir").wstring(),
            .parentPath     = caseRoot.wstring(),
            .name           = L"unzipped-dir",
            .fileAttributes = FILE_ATTRIBUTE_DIRECTORY,
            .reason         = USN_REASON_RENAME_NEW_NAME,
        },
        {
            .idPath         = (outsideRoot / L"moved-out-dir").wstring(),
            .parentPath     = caseRoot.wstring(),
            .name           = L"move-out-dir",
            .fileAttributes = FILE_ATTRIBUTE_DIRECTORY,
            .reason         = USN_REASON_RENAME_OLD_NAME,
        },
        {
            .idPath         = (outsideRoot / L"moved-out-dir").wstring(),
            .parentPath     = outsideRoot.wstring(),
            .name           = L"moved-out-dir",
            .fileAttributes = FILE_ATTRIBUTE_DIRECTORY,
            .reason         = USN_REASON_RENAME_NEW_NAME,
        },
        {
            .idPath         = (caseRoot / L"renamed-inside-dir").wstring(),
            .parentPath     = caseRoot.wstring(),
            .name           = L"inside-dir",
            .fileAttributes = FILE_ATTRIBUTE_DIRECTORY,
            .reason         = USN_REASON_RENAME_OLD_NAME,
        },
        {
            .idPath         = (caseRoot / L"renamed-inside-dir").wstring(),
            .parentPath     = caseRoot.wstring(),
            .name           = L"renamed-inside-dir",
            .fileAttributes = FILE_ATTRIBUTE_DIRECTORY,
            .reason         = USN_REASON_RENAME_NEW_NAME,
        },
    };
};

const auto runPersistedIndexedNameQueryForTests = [](LocalSearchIndexCore::Repository& repository,
                                                     std::wstring_view rootPath,
                                                     LocalSearchIndexCore::QueryStats& outStats,
                                                     std::vector<LocalSearchIndexCore::Candidate>& outCandidates) noexcept -> HRESULT
{
    LocalSearchIndexCore::QueryPlan plan{};
    plan.rootPath           = std::wstring(rootPath);
    plan.namePattern        = L"*.txt";
    plan.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    plan.recursive          = true;
    plan.includeFiles       = true;
    plan.includeDirectories = false;
    return repository.QueryPersistedVolumeForTests(plan, outCandidates, &outStats);
};

const auto requireJournalDirectoryMoveResults = [&](SelfTest::CaseState& state,
                                                    const std::filesystem::path& caseRoot,
                                                    const LocalSearchIndexCore::QueryStats& replayStats,
                                                    const std::vector<LocalSearchIndexCore::Candidate>& replayCandidates,
                                                    std::wstring_view label) noexcept -> bool
{
    state.Require(replayStats.journalReplayApplied, std::format(L"{} should apply journal replay.", label));
    state.Require(! replayStats.usedNtfsEnumeration, std::format(L"{} should not rebuild with NTFS enumeration during replay.", label));
    state.Require(CollectIndexedCandidateNames(replayCandidates) == expectedJournalDirectoryMoveNames(),
                  std::format(L"{} returned an unexpected candidate set after directory moves.", label));
    state.Require(hasCandidateFullPath(replayCandidates, caseRoot / L"renamed-inside-dir" / L"inside-child.txt"),
                  std::format(L"{} should keep the in-tree renamed child at the new path.", label));
    state.Require(! hasCandidateFullPath(replayCandidates, caseRoot / L"inside-dir" / L"inside-child.txt"),
                  std::format(L"{} should not keep the old in-tree renamed path.", label));
    state.Require(hasCandidateFullPath(replayCandidates, caseRoot / L"moved-in-dir" / L"nested" / L"moved-in-nested.txt"),
                  std::format(L"{} should hydrate children of a populated moved-in directory.", label));
    state.Require(hasCandidateFullPath(replayCandidates, caseRoot / L"unzipped-dir" / L"nested" / L"unzipped-nested.txt"),
                  std::format(L"{} should hydrate children of an unzipped tree moved into the root.", label));
    state.Require(! hasCandidateFullPath(replayCandidates, caseRoot / L"move-out-dir" / L"moved-out-child.txt"),
                  std::format(L"{} should remove the moved-out subtree from the index.", label));
    return state.failure.empty();
};

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_journal_replay_directory_moves_snapshot",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_journal_directory_moves_snapshot", caseRoot),
                  L"Failed to prepare index_core_journal_directory_moves_snapshot root.");
    std::filesystem::path outsideRoot;
    if (! prepareJournalDirectoryMoveFixture(state, caseRoot, L"index_core_journal_directory_moves_snapshot_outside", outsideRoot))
    {
        return false;
    }

    LocalSearchIndexCore::Repository repository;
    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", seedStats, seedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Directory-move snapshot seed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(CollectIndexedCandidateNames(seedCandidates) == std::vector<std::wstring>{L"inside-child.txt", L"moved-out-child.txt", L"stable.txt"},
                  L"Directory-move snapshot seed query returned an unexpected candidate set.");
    UNREFERENCED_PARAMETER(seedStats);

    if (! applyJournalDirectoryMoveMutations(state, caseRoot, outsideRoot))
    {
        return false;
    }

    const std::vector<LocalSearchIndexCore::SyntheticJournalRecordForTests> syntheticRecords = buildJournalDirectoryMoveSyntheticRecords(caseRoot, outsideRoot);
    LocalSearchIndexCore::QueryStats replayStats{};
    hr = repository.ApplySyntheticJournalForTests(caseRoot.wstring(), syntheticRecords, &replayStats);
    state.Require(SUCCEEDED(hr), std::format(L"Directory-move snapshot synthetic replay failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    LocalSearchIndexCore::QueryStats persistedStats{};
    std::vector<LocalSearchIndexCore::Candidate> persistedCandidates;
    hr = runPersistedIndexedNameQueryForTests(repository, caseRoot.wstring(), persistedStats, persistedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Directory-move snapshot persisted query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(persistedStats.snapshotLoaded, L"Directory-move snapshot persisted query should load the saved snapshot.");
    state.Require(requireJournalDirectoryMoveResults(state, caseRoot, replayStats, persistedCandidates, L"Directory-move snapshot synthetic replay"),
                  L"Directory-move snapshot synthetic replay should persist the expected directory-move result set.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_journal_replay_invalidated_rebuilds",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_journal_invalidated_rebuild", caseRoot),
                  L"Failed to prepare index_core_journal_invalidated_rebuild root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::Repository repository;
    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", seedStats, seedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Journal invalidation seed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(CollectIndexedCandidateNames(seedCandidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Journal invalidation seed query returned an unexpected candidate set.");

    LocalSearchIndexCore::QueryStats syntheticStats{};
    hr = repository.ApplySyntheticJournalForTests(caseRoot.wstring(), std::span<const LocalSearchIndexCore::SyntheticJournalRecordForTests>{}, &syntheticStats);
    state.Require(SUCCEEDED(hr), std::format(L"Journal invalidation synthetic ready marker failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(syntheticStats.journalId != 0u && syntheticStats.nextUsn != 0u,
                  L"Journal invalidation synthetic ready marker should assign a deterministic journal position.");

    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.txt", "beta"), L"Failed to create beta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }
    ::Sleep(25u);

    repository.SetNextJournalStateForTests(syntheticStats.journalId, 0u, syntheticStats.nextUsn + 8u);
    repository.SetNextJournalReplayReadFailureForTests(ERROR_JOURNAL_ENTRY_DELETED);
    const auto clearInjectedFailure = wil::scope_exit([&] noexcept { repository.SetNextJournalReplayReadFailureForTests(ERROR_SUCCESS); });

    LocalSearchIndexCore::QueryStats rebuildStats{};
    std::vector<LocalSearchIndexCore::Candidate> rebuildCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", rebuildStats, rebuildCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Journal invalidation query should rebuild instead of failing. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(rebuildStats.rebuiltJournalRangeInvalid, L"Journal invalidation query should report a journal-range rebuild after replay read failure.");
    state.Require(CollectIndexedCandidateNames(rebuildCandidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"Journal invalidation rebuild should return the fresh filesystem candidate set.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_snapshot_corruption_rebuild",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_snapshot_corruption", caseRoot), L"Failed to prepare index_core_snapshot_corruption root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");

    LocalSearchIndexCore::Repository repository;
    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", seedStats, seedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Seed indexed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (! state.failure.empty())
    {
        return false;
    }

    hr = repository.CorruptSnapshotForTests(caseRoot.wstring(), LocalSearchIndexCore::SnapshotCorruptionMode::InvalidMagic);
    state.Require(SUCCEEDED(hr), std::format(L"CorruptSnapshotForTests(InvalidMagic) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    hr = repository.DropCachedVolumeForTests(caseRoot.wstring());
    state.Require(SUCCEEDED(hr), std::format(L"DropCachedVolumeForTests failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    LocalSearchIndexCore::QueryStats invalidMagicStats{};
    std::vector<LocalSearchIndexCore::Candidate> invalidMagicCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", invalidMagicStats, invalidMagicCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Indexed query after invalid-magic corruption failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(invalidMagicStats.rebuiltSnapshotCorruption, L"Indexed query should report a rebuild after invalid snapshot magic corruption.");

    hr = repository.CorruptSnapshotForTests(caseRoot.wstring(), LocalSearchIndexCore::SnapshotCorruptionMode::EntryCountTooLarge);
    state.Require(SUCCEEDED(hr), std::format(L"CorruptSnapshotForTests(EntryCountTooLarge) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    hr = repository.DropCachedVolumeForTests(caseRoot.wstring());
    state.Require(SUCCEEDED(hr), std::format(L"DropCachedVolumeForTests failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    LocalSearchIndexCore::QueryStats hugeCountStats{};
    std::vector<LocalSearchIndexCore::Candidate> hugeCountCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", hugeCountStats, hugeCountCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Indexed query after huge-count corruption failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(hugeCountStats.rebuiltSnapshotCorruption, L"Indexed query should report a rebuild after impossible snapshot entry count.");
    state.Require(CollectIndexedCandidateNames(hugeCountCandidates) == CollectIndexedCandidateNames(seedCandidates),
                  L"Indexed query after huge-count rebuild should preserve the candidate set.");

    hr = repository.CorruptSnapshotForTests(caseRoot.wstring(), LocalSearchIndexCore::SnapshotCorruptionMode::NextUsnPastEnd);
    state.Require(SUCCEEDED(hr), std::format(L"CorruptSnapshotForTests(NextUsnPastEnd) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    hr = repository.DropCachedVolumeForTests(caseRoot.wstring());
    state.Require(SUCCEEDED(hr), std::format(L"DropCachedVolumeForTests failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    LocalSearchIndexCore::QueryStats rangeStats{};
    std::vector<LocalSearchIndexCore::Candidate> rangeCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", rangeStats, rangeCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Indexed query after NextUsn corruption failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (rangeStats.journalAvailable)
    {
        state.Require(rangeStats.rebuiltJournalRangeInvalid, L"Indexed query should rebuild when the snapshot NextUsn is beyond the live journal range.");
    }
    state.Require(CollectIndexedCandidateNames(rangeCandidates) == CollectIndexedCandidateNames(seedCandidates),
                  L"Indexed query after corruption rebuild should preserve the candidate set.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_usn_record_bounds_reject_corruption",
                  [&](SelfTest::CaseState& state) noexcept
{
    const size_t nameOffset = static_cast<size_t>(FIELD_OFFSET(USN_RECORD_V2, FileName));

    std::vector<std::byte> validRecordBytes(nameOffset + (2u * sizeof(wchar_t)));
    auto* validRecord           = reinterpret_cast<USN_RECORD_V2*>(validRecordBytes.data());
    validRecord->RecordLength   = static_cast<DWORD>(validRecordBytes.size());
    validRecord->MajorVersion   = 2u;
    validRecord->MinorVersion   = 0u;
    validRecord->FileNameOffset = static_cast<WORD>(nameOffset);
    validRecord->FileNameLength = static_cast<WORD>(2u * sizeof(wchar_t));
    std::memcpy(validRecordBytes.data() + nameOffset, L"ok", 2u * sizeof(wchar_t));
    state.Require(LocalSearchIndexCore::TryParseUsnRecordForTests(validRecordBytes.data(), validRecordBytes.size()), L"Valid USN_RECORD_V2 should parse.");

    std::vector<std::byte> oddLengthRecordBytes(nameOffset + sizeof(wchar_t));
    auto* oddLengthRecord           = reinterpret_cast<USN_RECORD_V2*>(oddLengthRecordBytes.data());
    oddLengthRecord->RecordLength   = static_cast<DWORD>(oddLengthRecordBytes.size());
    oddLengthRecord->MajorVersion   = 2u;
    oddLengthRecord->MinorVersion   = 0u;
    oddLengthRecord->FileNameOffset = static_cast<WORD>(nameOffset);
    oddLengthRecord->FileNameLength = 1u;
    state.Require(! LocalSearchIndexCore::TryParseUsnRecordForTests(oddLengthRecordBytes.data(), oddLengthRecordBytes.size()),
                  L"USN_RECORD_V2 with odd FileNameLength should be rejected.");

    std::vector<std::byte> overflowingNameRecordBytes(nameOffset + (4u * sizeof(wchar_t)) + 16u);
    auto* overflowingNameRecord           = reinterpret_cast<USN_RECORD_V2*>(overflowingNameRecordBytes.data());
    overflowingNameRecord->RecordLength   = static_cast<DWORD>(nameOffset + sizeof(wchar_t));
    overflowingNameRecord->MajorVersion   = 2u;
    overflowingNameRecord->MinorVersion   = 0u;
    overflowingNameRecord->FileNameOffset = static_cast<WORD>(nameOffset);
    overflowingNameRecord->FileNameLength = static_cast<WORD>(4u * sizeof(wchar_t));
    state.Require(! LocalSearchIndexCore::TryParseUsnRecordForTests(overflowingNameRecordBytes.data(), overflowingNameRecordBytes.size()),
                  L"USN_RECORD_V2 with FileNameOffset + FileNameLength beyond RecordLength should be rejected.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_broker_query_batch_rejects_oversized_count",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto makeBatchPayload = [](const uint32_t count) noexcept
    {
        std::array<std::byte, sizeof(uint32_t) * 2u> payload{};
        constexpr uint32_t kReserved = 0u;
        std::memcpy(payload.data(), &count, sizeof(count));
        std::memcpy(payload.data() + sizeof(count), &kReserved, sizeof(kReserved));
        return payload;
    };

    const auto emptyPayload = makeBatchPayload(0u);
    std::vector<LocalSearchIndexCore::Candidate> emptyCandidates;
    HRESULT hr = SearchServiceBroker::DecodeQueryBatchForTests(std::span<const std::byte>(emptyPayload.data(), emptyPayload.size()), emptyCandidates);
    state.Require(hr == S_OK, std::format(L"Empty QueryBatch should decode. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(emptyCandidates.empty(), L"Empty QueryBatch should not produce candidates.");

    const HRESULT protocolErrorHr = HRESULT_FROM_WIN32(RPC_S_PROTOCOL_ERROR);
    const auto hugeCountPayload   = makeBatchPayload((std::numeric_limits<uint32_t>::max)());
    std::vector<LocalSearchIndexCore::Candidate> hugeCountCandidates;
    hr = SearchServiceBroker::DecodeQueryBatchForTests(std::span<const std::byte>(hugeCountPayload.data(), hugeCountPayload.size()), hugeCountCandidates);
    state.Require(hr == protocolErrorHr,
                  std::format(L"Huge-count QueryBatch should decode as a failing protocol error. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(FAILED(hr), L"Huge-count QueryBatch protocol error should satisfy FAILED(hr).");
    state.Require(hugeCountCandidates.empty(), L"Huge-count QueryBatch should leave decoded candidates empty.");

    struct TestFrameHeader final
    {
        uint32_t magic           = 0x53535252u;
        uint32_t protocolVersion = SearchServiceBroker::kProtocolVersion;
        uint32_t messageType     = 0u;
        uint32_t payloadBytes    = 0u;
    };

    struct TestStatusResponsePayload final
    {
        uint32_t processId        = 0u;
        uint32_t pipeNameBytes    = 0u;
        uint32_t storageRootBytes = 0u;
        uint32_t reserved         = 0u;
    };

    const auto readExact = [](HANDLE handle, void* buffer, const uint32_t byteCount) noexcept -> HRESULT
    {
        uint8_t* destination = static_cast<uint8_t*>(buffer);
        uint32_t totalRead   = 0u;
        while (totalRead < byteCount)
        {
            DWORD chunkRead = 0u;
            if (::ReadFile(handle, destination + totalRead, byteCount - totalRead, &chunkRead, nullptr) == 0)
            {
                return HRESULT_FROM_WIN32(::GetLastError());
            }

            if (chunkRead == 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
            }

            totalRead += chunkRead;
        }

        return S_OK;
    };

    const auto writeExact = [](HANDLE handle, const void* buffer, const uint32_t byteCount) noexcept -> HRESULT
    {
        const uint8_t* source = static_cast<const uint8_t*>(buffer);
        uint32_t totalWritten = 0u;
        while (totalWritten < byteCount)
        {
            DWORD chunkWritten = 0u;
            if (::WriteFile(handle, source + totalWritten, byteCount - totalWritten, &chunkWritten, nullptr) == 0)
            {
                return HRESULT_FROM_WIN32(::GetLastError());
            }

            if (chunkWritten == 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
            }

            totalWritten += chunkWritten;
        }

        return S_OK;
    };

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0,
                  L"Failed to override the search service pipe for malformed QueryBatch test.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    const auto serveSingleFakeServiceFrame = [&](const uint32_t messageType, std::span<const std::byte> responsePayload) noexcept -> HRESULT
    {
        wil::unique_handle pipe(
            ::CreateNamedPipeW(pipeName.c_str(), PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT, 1u, 4096u, 4096u, 0u, nullptr));
        if (! pipe)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        if (::ConnectNamedPipe(pipe.get(), nullptr) == 0)
        {
            const DWORD error = ::GetLastError();
            if (error != ERROR_PIPE_CONNECTED)
            {
                return HRESULT_FROM_WIN32(error);
            }
        }

        TestFrameHeader requestHeader{};
        HRESULT localHr = readExact(pipe.get(), &requestHeader, static_cast<uint32_t>(sizeof(requestHeader)));
        if (FAILED(localHr))
        {
            return localHr;
        }

        std::vector<std::byte> requestPayload;
        requestPayload.resize(requestHeader.payloadBytes);
        if (requestHeader.payloadBytes != 0u)
        {
            localHr = readExact(pipe.get(), requestPayload.data(), requestHeader.payloadBytes);
            if (FAILED(localHr))
            {
                return localHr;
            }
        }

        TestFrameHeader responseHeader{};
        responseHeader.messageType  = messageType;
        responseHeader.payloadBytes = static_cast<uint32_t>(responsePayload.size_bytes());
        localHr                     = writeExact(pipe.get(), &responseHeader, static_cast<uint32_t>(sizeof(responseHeader)));
        if (SUCCEEDED(localHr) && ! responsePayload.empty())
        {
            localHr = writeExact(pipe.get(), responsePayload.data(), static_cast<uint32_t>(responsePayload.size_bytes()));
        }
        if (SUCCEEDED(localHr) && ::FlushFileBuffers(pipe.get()) == 0)
        {
            localHr = HRESULT_FROM_WIN32(::GetLastError());
        }

        return localHr;
    };

    wil::unique_handle serverPipe(
        ::CreateNamedPipeW(pipeName.c_str(), PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT, 1u, 4096u, 4096u, 0u, nullptr));
    state.Require(! ! serverPipe,
                  std::format(L"Failed to create malformed QueryBatch fake service pipe. error={}", static_cast<unsigned long>(::GetLastError())));
    if (! serverPipe)
    {
        return false;
    }

    std::atomic<long> serverHr{S_OK};
    std::jthread server([&, pipe = std::move(serverPipe)](std::stop_token) mutable noexcept
    {
        if (::ConnectNamedPipe(pipe.get(), nullptr) == 0)
        {
            const DWORD error = ::GetLastError();
            if (error != ERROR_PIPE_CONNECTED)
            {
                serverHr.store(static_cast<long>(HRESULT_FROM_WIN32(error)));
                return;
            }
        }

        TestFrameHeader requestHeader{};
        HRESULT localHr = readExact(pipe.get(), &requestHeader, static_cast<uint32_t>(sizeof(requestHeader)));
        if (FAILED(localHr))
        {
            serverHr.store(static_cast<long>(localHr));
            return;
        }

        std::vector<std::byte> requestPayload;
        requestPayload.resize(requestHeader.payloadBytes);
        if (requestHeader.payloadBytes != 0u)
        {
            localHr = readExact(pipe.get(), requestPayload.data(), requestHeader.payloadBytes);
            if (FAILED(localHr))
            {
                serverHr.store(static_cast<long>(localHr));
                return;
            }
        }

        TestFrameHeader batchHeader{};
        batchHeader.messageType  = 5u;
        batchHeader.payloadBytes = static_cast<uint32_t>(hugeCountPayload.size());
        localHr                  = writeExact(pipe.get(), &batchHeader, static_cast<uint32_t>(sizeof(batchHeader)));
        if (SUCCEEDED(localHr))
        {
            localHr = writeExact(pipe.get(), hugeCountPayload.data(), static_cast<uint32_t>(hugeCountPayload.size()));
        }
        if (SUCCEEDED(localHr) && ::FlushFileBuffers(pipe.get()) == 0)
        {
            localHr = HRESULT_FROM_WIN32(::GetLastError());
        }

        serverHr.store(static_cast<long>(localHr));
    });

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = L"C:\\";
    request.namePattern        = L"*";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.includeFiles       = true;
    request.includeDirectories = false;

    std::vector<LocalSearchIndexCore::Candidate> malformedCandidates;
    LocalSearchIndexCore::QueryStats stats{};
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, malformedCandidates, &stats);
    server.join();

    const HRESULT serverResult = static_cast<HRESULT>(serverHr.load());
    state.Require(SUCCEEDED(serverResult), std::format(L"Malformed QueryBatch fake server failed. hr=0x{:08X}", static_cast<unsigned long>(serverResult)));
    state.Require(hr == protocolErrorHr,
                  std::format(L"Broker Query should fail with RPC_S_PROTOCOL_ERROR for huge-count QueryBatch. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(malformedCandidates.empty(), L"Malformed QueryBatch should leave broker query candidates empty.");

    TestStatusResponsePayload statusPayloadPod{};
    statusPayloadPod.processId = ::GetCurrentProcessId();
    std::array<std::byte, sizeof(statusPayloadPod)> statusPayload{};
    std::memcpy(statusPayload.data(), &statusPayloadPod, sizeof(statusPayloadPod));

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! info || ! search)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"service\"}");
    state.Require(SUCCEEDED(setHr), std::format(L"Failed to configure service search. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    std::filesystem::path caseRoot;
    if (! PrepareSearchCaseRoot(root, L"search_service_malformed_querybatch_fallback", caseRoot))
    {
        state.Require(false, L"Failed to prepare search_service_malformed_querybatch_fallback root.");
        return false;
    }
    if (! SelfTest::WriteTextFile(caseRoot / L"fallback.txt", "fallback"))
    {
        state.Require(false, L"Failed to create fallback.txt.");
        return false;
    }

    std::wstring rootText    = caseRoot.wstring();
    std::wstring namePattern = L"*.txt";
    FileSystemSearchQuery fallbackQuery{};
    fallbackQuery.sizeBytes   = sizeof(FileSystemSearchQuery);
    fallbackQuery.rootPath    = rootText.c_str();
    fallbackQuery.namePattern = namePattern.c_str();
    fallbackQuery.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    fallbackQuery.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;

    std::atomic<long> fallbackServerHr{S_OK};
    std::jthread fallbackServer([&](std::stop_token) noexcept
    {
        HRESULT localHr = serveSingleFakeServiceFrame(2u, std::span<const std::byte>(statusPayload.data(), statusPayload.size()));
        if (SUCCEEDED(localHr))
        {
            localHr = serveSingleFakeServiceFrame(5u, std::span<const std::byte>(hugeCountPayload.data(), hugeCountPayload.size()));
        }

        fallbackServerHr.store(static_cast<long>(localHr));
    });

    RecordingSearchCallback callback;
    const HRESULT searchHr = search->Search(&fallbackQuery, &callback, nullptr);
    fallbackServer.join();

    const HRESULT fallbackServerResult = static_cast<HRESULT>(fallbackServerHr.load());
    state.Require(SUCCEEDED(fallbackServerResult),
                  std::format(L"Malformed QueryBatch fallback fake server failed. hr=0x{:08X}", static_cast<unsigned long>(fallbackServerResult)));
    state.Require(SUCCEEDED(searchHr),
                  std::format(L"Malformed QueryBatch service search should fall back and succeed. hr=0x{:08X}", static_cast<unsigned long>(searchHr)));

    const auto matches = callback.Matches();
    state.Require(matches.size() == 1u, std::format(L"Malformed QueryBatch fallback search expected 1 match. got={}", matches.size()));
    if (! matches.empty())
    {
        state.Require(matches[0].displayName == L"fallback.txt", L"Malformed QueryBatch fallback returned the wrong file.");
    }

    const auto progressSnapshots            = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Malformed QueryBatch fallback search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend != FILESYSTEM_SEARCH_BACKEND_SERVICE, L"Malformed QueryBatch fallback should not complete on the service backend.");
        state.Require((completed->warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) != 0u,
                      L"Malformed QueryBatch fallback should report the service-unavailable warning.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_broker_buffered_query_caps_unbounded_stream",
                  [&](SelfTest::CaseState& state) noexcept
{
    struct TestFrameHeader final
    {
        uint32_t magic           = 0x53535252u;
        uint32_t protocolVersion = SearchServiceBroker::kProtocolVersion;
        uint32_t messageType     = 0u;
        uint32_t payloadBytes    = 0u;
    };

    struct TestCandidateBatchHeader final
    {
        uint32_t count    = 0u;
        uint32_t reserved = 0u;
    };

    struct TestCandidateEntryHeader final
    {
        uint32_t fileAttributes     = 0u;
        uint32_t metadataFlags      = LocalSearchIndexCore::CANDIDATE_METADATA_NONE;
        uint32_t fullPathBytes      = 0u;
        uint32_t displayNameBytes   = 0u;
        int64_t creationTime100ns   = 0;
        int64_t lastAccessTime100ns = 0;
        int64_t endOfFile           = 0;
        int64_t lastWriteTime100ns  = 0;
        int64_t changeTime100ns     = 0;
        int64_t allocationSize      = 0;
    };

    struct TestQueryCompletePayload final
    {
        int32_t result                  = S_OK;
        uint32_t fileSystemKind         = 0u;
        uint32_t flags                  = 0u;
        uint64_t entryCount             = 0u;
        uint64_t fileCount              = 0u;
        uint64_t directoryCount         = 0u;
        uint64_t candidateCount         = 0u;
        uint64_t nextUsn                = 0u;
        uint64_t journalId              = 0u;
        uint64_t snapshotFileBytes      = 0u;
        uint64_t estimatedMemoryBytes   = 0u;
        uint32_t ensureReadyDurationMs  = 0u;
        uint32_t executeQueryDurationMs = 0u;
    };

    const auto readExact = [](HANDLE handle, void* buffer, const uint32_t byteCount) noexcept -> HRESULT
    {
        uint8_t* destination = static_cast<uint8_t*>(buffer);
        uint32_t totalRead   = 0u;
        while (totalRead < byteCount)
        {
            DWORD chunkRead = 0u;
            if (::ReadFile(handle, destination + totalRead, byteCount - totalRead, &chunkRead, nullptr) == 0)
            {
                return HRESULT_FROM_WIN32(::GetLastError());
            }

            if (chunkRead == 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
            }

            totalRead += chunkRead;
        }

        return S_OK;
    };

    const auto writeExact = [](HANDLE handle, const void* buffer, const uint32_t byteCount) noexcept -> HRESULT
    {
        const uint8_t* source = static_cast<const uint8_t*>(buffer);
        uint32_t totalWritten = 0u;
        while (totalWritten < byteCount)
        {
            DWORD chunkWritten = 0u;
            if (::WriteFile(handle, source + totalWritten, byteCount - totalWritten, &chunkWritten, nullptr) == 0)
            {
                return HRESULT_FROM_WIN32(::GetLastError());
            }

            if (chunkWritten == 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
            }

            totalWritten += chunkWritten;
        }

        return S_OK;
    };

    const auto appendBytes = [](std::vector<std::byte>& buffer, const void* data, const size_t byteCount)
    {
        const auto* bytes = static_cast<const std::byte*>(data);
        buffer.insert(buffer.end(), bytes, bytes + byteCount);
    };

    const auto appendUtf16 = [&](std::vector<std::byte>& buffer, const std::wstring& text) { appendBytes(buffer, text.data(), text.size() * sizeof(wchar_t)); };

    const auto makeBatchPayload = [&](const uint32_t firstIndex, const uint32_t count) -> std::vector<std::byte>
    {
        TestCandidateBatchHeader batchHeader{};
        batchHeader.count = count;

        std::vector<std::byte> payload;
        payload.reserve(sizeof(batchHeader) + (static_cast<size_t>(count) * sizeof(TestCandidateEntryHeader)));
        appendBytes(payload, &batchHeader, sizeof(batchHeader));

        for (uint32_t index = 0u; index < count; ++index)
        {
            const uint32_t candidateIndex  = firstIndex + index;
            const std::wstring displayName = std::format(L"candidate_{:05}.txt", candidateIndex);
            const std::wstring fullPath    = std::wstring(L"C:\\buffered\\") + displayName;

            TestCandidateEntryHeader entry{};
            entry.fileAttributes   = FILE_ATTRIBUTE_ARCHIVE;
            entry.fullPathBytes    = static_cast<uint32_t>(fullPath.size() * sizeof(wchar_t));
            entry.displayNameBytes = static_cast<uint32_t>(displayName.size() * sizeof(wchar_t));
            appendBytes(payload, &entry, sizeof(entry));
            appendUtf16(payload, fullPath);
            appendUtf16(payload, displayName);
        }

        return payload;
    };

    const auto isExpectedPipeDisconnect = [](const HRESULT hr) noexcept
    { return hr == HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE) || hr == HRESULT_FROM_WIN32(ERROR_NO_DATA) || hr == HRESULT_FROM_WIN32(ERROR_PIPE_NOT_CONNECTED); };

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0,
                  L"Failed to override the search service pipe for buffered-query cap test.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    wil::unique_handle serverPipe(
        ::CreateNamedPipeW(pipeName.c_str(), PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT, 1u, 64u * 1024u, 64u * 1024u, 0u, nullptr));
    state.Require(! ! serverPipe,
                  std::format(L"Failed to create buffered-query cap fake service pipe. error={}", static_cast<unsigned long>(::GetLastError())));
    if (! serverPipe)
    {
        return false;
    }

    constexpr uint32_t kBatchSize      = 256u;
    constexpr uint32_t kCandidateCount = 65'537u;
    std::atomic<long> serverHr{S_OK};
    std::atomic<uint32_t> candidatesSent{0u};
    std::jthread server([&, pipe = std::move(serverPipe)](std::stop_token) mutable noexcept
    {
        HRESULT localHr = S_OK;
        if (::ConnectNamedPipe(pipe.get(), nullptr) == 0)
        {
            const DWORD error = ::GetLastError();
            if (error != ERROR_PIPE_CONNECTED)
            {
                serverHr.store(static_cast<long>(HRESULT_FROM_WIN32(error)));
                return;
            }
        }

        TestFrameHeader requestHeader{};
        localHr = readExact(pipe.get(), &requestHeader, static_cast<uint32_t>(sizeof(requestHeader)));
        if (SUCCEEDED(localHr))
        {
            std::vector<std::byte> requestPayload;
            requestPayload.resize(requestHeader.payloadBytes);
            if (requestHeader.payloadBytes != 0u)
            {
                localHr = readExact(pipe.get(), requestPayload.data(), requestHeader.payloadBytes);
            }
        }

        for (uint32_t offset = 0u; SUCCEEDED(localHr) && offset < kCandidateCount; offset += kBatchSize)
        {
            const uint32_t count                      = (std::min)(kBatchSize, kCandidateCount - offset);
            const std::vector<std::byte> batchPayload = makeBatchPayload(offset, count);

            TestFrameHeader batchFrame{};
            batchFrame.messageType  = 5u;
            batchFrame.payloadBytes = static_cast<uint32_t>(batchPayload.size());
            localHr                 = writeExact(pipe.get(), &batchFrame, static_cast<uint32_t>(sizeof(batchFrame)));
            if (SUCCEEDED(localHr))
            {
                localHr = writeExact(pipe.get(), batchPayload.data(), static_cast<uint32_t>(batchPayload.size()));
            }

            if (FAILED(localHr) && isExpectedPipeDisconnect(localHr) && candidatesSent.load(std::memory_order_acquire) != 0u)
            {
                localHr = S_OK;
                break;
            }

            if (SUCCEEDED(localHr))
            {
                candidatesSent.fetch_add(count, std::memory_order_acq_rel);
            }
        }

        if (SUCCEEDED(localHr))
        {
            TestQueryCompletePayload complete{};
            complete.entryCount     = kCandidateCount;
            complete.fileCount      = kCandidateCount;
            complete.candidateCount = kCandidateCount;

            TestFrameHeader completeFrame{};
            completeFrame.messageType  = 6u;
            completeFrame.payloadBytes = static_cast<uint32_t>(sizeof(complete));
            localHr                    = writeExact(pipe.get(), &completeFrame, static_cast<uint32_t>(sizeof(completeFrame)));
            if (SUCCEEDED(localHr))
            {
                localHr = writeExact(pipe.get(), &complete, static_cast<uint32_t>(sizeof(complete)));
            }
            if (FAILED(localHr) && isExpectedPipeDisconnect(localHr) && candidatesSent.load(std::memory_order_acquire) >= kCandidateCount)
            {
                localHr = S_OK;
            }
        }

        serverHr.store(static_cast<long>(localHr));
    });

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = L"C:\\";
    request.namePattern        = L"*";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.includeFiles       = true;
    request.includeDirectories = false;

    std::vector<LocalSearchIndexCore::Candidate> candidates;
    LocalSearchIndexCore::QueryStats stats{};
    const HRESULT hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, candidates, &stats);
    server.join();

    const HRESULT serverResult = static_cast<HRESULT>(serverHr.load());
    state.Require(SUCCEEDED(serverResult), std::format(L"Buffered-query cap fake server failed. hr=0x{:08X}", static_cast<unsigned long>(serverResult)));
    const HRESULT overflowHr = HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    state.Require(
        hr == overflowHr,
        std::format(L"Broker Query should cap no-callback candidate accumulation with ERROR_BUFFER_OVERFLOW. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(candidates.empty(), std::format(L"Overflowing buffered broker query should clear candidates, got {}.", candidates.size()));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_source_allocation_and_folding_guard",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto contains = [](std::string_view source, std::string_view needle) noexcept { return source.find(needle) != std::string_view::npos; };

    const auto containsBefore = [](std::string_view source, std::string_view before, std::string_view after) noexcept
    {
        const size_t beforeOffset = source.find(before);
        const size_t afterOffset  = source.find(after);
        return beforeOffset != std::string_view::npos && afterOffset != std::string_view::npos && beforeOffset < afterOffset;
    };
    const auto compactWhitespace = [](std::string_view source)
    {
        std::string compact;
        compact.reserve(source.size());
        for (const char ch : source)
        {
            if (ch != ' ' && ch != '\t' && ch != '\r' && ch != '\n')
            {
                compact.push_back(ch);
            }
        }
        return compact;
    };
    const auto containsBeforeCompact = [&](std::string_view source, std::string_view before, std::string_view after)
    {
        const std::string compactSourceText = compactWhitespace(source);
        const std::string compactBefore     = compactWhitespace(before);
        const std::string compactAfter      = compactWhitespace(after);
        const size_t beforeOffset           = compactSourceText.find(compactBefore);
        const size_t afterOffset            = compactSourceText.find(compactAfter);
        return beforeOffset != std::string::npos && afterOffset != std::string::npos && beforeOffset < afterOffset;
    };

    const auto sourceBetween = [](std::string_view source, std::string_view beginMarker, std::string_view endMarker) noexcept -> std::string_view
    {
        const size_t beginOffset = source.find(beginMarker);
        if (beginOffset == std::string_view::npos)
        {
            return {};
        }

        const size_t endOffset = source.find(endMarker, beginOffset + beginMarker.size());
        if (endOffset == std::string_view::npos || endOffset <= beginOffset)
        {
            return {};
        }

        return source.substr(beginOffset, endOffset - beginOffset);
    };

    const auto findRepoRoot = []() noexcept -> std::filesystem::path
    {
        std::error_code ec;
        std::filesystem::path current = std::filesystem::current_path(ec);
        if (ec)
        {
            return {};
        }

        for (size_t depth = 0u; depth < 8u && ! current.empty(); ++depth)
        {
            std::error_code slnEc;
            std::error_code brokerEc;
            if (std::filesystem::exists(current / L"RedSalamander.sln", slnEc) &&
                std::filesystem::exists(current / L"Common" / L"SearchServiceBroker.cpp", brokerEc))
            {
                return current;
            }

            const std::filesystem::path parent = current.parent_path();
            if (parent.empty() || parent == current)
            {
                break;
            }
            current = parent;
        }

        return {};
    };

    const auto readSourceFile = [](const std::filesystem::path& path, std::string& outSource) noexcept
    {
        outSource.clear();
        std::ifstream stream(path, std::ios::binary);
        if (! stream)
        {
            return false;
        }

        outSource.assign(std::istreambuf_iterator<char>(stream), std::istreambuf_iterator<char>());
        return stream.eof() || stream.good();
    };

    const auto requireNoBroadCatch = [&](std::wstring_view fileName, std::string_view source) noexcept
    { state.Require(! contains(source, "catch (...)"), std::format(L"{} must not use catch (...).", fileName)); };

    const auto requireBadAllocCatchesTerminate = [&](std::wstring_view fileName, std::string_view source) noexcept
    {
        size_t searchOffset = 0u;
        while (true)
        {
            const size_t catchOffset = source.find("catch (const std::bad_alloc&", searchOffset);
            if (catchOffset == std::string_view::npos)
            {
                break;
            }

            const size_t bodyOpen = source.find('{', catchOffset);
            if (bodyOpen == std::string_view::npos)
            {
                state.Require(false, std::format(L"{} has an unparsable std::bad_alloc catch.", fileName));
                break;
            }

            size_t bodyClose  = bodyOpen + 1u;
            size_t braceDepth = 1u;
            bool foundBodyEnd = false;
            for (; bodyClose < source.size(); ++bodyClose)
            {
                if (source[bodyClose] == '{')
                {
                    ++braceDepth;
                }
                else if (source[bodyClose] == '}')
                {
                    --braceDepth;
                    if (braceDepth == 0u)
                    {
                        foundBodyEnd = true;
                        break;
                    }
                }
            }

            if (! foundBodyEnd)
            {
                state.Require(false, std::format(L"{} has an unterminated std::bad_alloc catch.", fileName));
                break;
            }

            const std::string_view body = source.substr(bodyOpen, bodyClose - bodyOpen + 1u);
            state.Require(contains(body, "std::terminate();"), std::format(L"{} must terminate inside every std::bad_alloc catch on search paths.", fileName));
            searchOffset = bodyClose + 1u;
        }
    };

    const std::filesystem::path repoRoot = findRepoRoot();
    state.Require(! repoRoot.empty(), L"Search source guard could not locate the repository root.");
    if (repoRoot.empty())
    {
        return false;
    }

    std::string brokerSource;
    std::string localIndexSource;
    std::string sqliteSource;
    std::string pluginSearchSource;
    std::string searchTextSource;
    std::string findWindowSource;
    std::string incrementalSearchSource;
    state.Require(readSourceFile(repoRoot / L"Common" / L"SearchServiceBroker.cpp", brokerSource), L"Failed to read SearchServiceBroker.cpp.");
    state.Require(readSourceFile(repoRoot / L"Common" / L"LocalSearchIndexCore.cpp", localIndexSource), L"Failed to read LocalSearchIndexCore.cpp.");
    state.Require(readSourceFile(repoRoot / L"Common" / L"SqliteIndexStore.cpp", sqliteSource), L"Failed to read SqliteIndexStore.cpp.");
    state.Require(readSourceFile(repoRoot / L"Plugins" / L"FileSystem" / L"FileSystem.Search.cpp", pluginSearchSource),
                  L"Failed to read FileSystem.Search.cpp.");
    state.Require(readSourceFile(repoRoot / L"Common" / L"SearchTextHelpers.cpp", searchTextSource), L"Failed to read SearchTextHelpers.cpp.");
    state.Require(readSourceFile(repoRoot / L"RedSalamander" / L"FindFilesWindow.cpp", findWindowSource), L"Failed to read FindFilesWindow.cpp.");
    state.Require(readSourceFile(repoRoot / L"RedSalamander" / L"FolderViewIncrementalSearch.h", incrementalSearchSource),
                  L"Failed to read FolderViewIncrementalSearch.h.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::string_view decodeBatch =
        sourceBetween(brokerSource, "HRESULT DecodeQueryBatchPayload", "[[nodiscard]] size_t ResolveClientBufferedCandidateLimit");
    state.Require(! decodeBatch.empty(), L"SearchServiceBroker.cpp must keep DecodeQueryBatchPayload discoverable for source guard.");
    state.Require(containsBefore(decodeBatch,
                                 "const size_t maxEntriesInPayload = remaining.size_bytes() / sizeof(CandidateEntryHeader);",
                                 "batchCandidates.reserve(batchHeader.count);"),
                  L"QueryBatch decode must compute the maximum payload entry count before reserving candidate storage.");
    state.Require(
        containsBefore(decodeBatch, "if (static_cast<size_t>(batchHeader.count) > maxEntriesInPayload)", "batchCandidates.reserve(batchHeader.count);"),
        L"QueryBatch decode must reject count > remaining payload entries before reserve.");
    state.Require(contains(decodeBatch, "return kProtocolErrorHr;"), L"QueryBatch external-count preflight must fail with the protocol-error HRESULT.");

    const std::string_view queryBody = sourceBetween(brokerSource, "HRESULT Query(const QueryRequest& request,", "HRESULT RequestRebuild(");
    state.Require(! queryBody.empty(), L"SearchServiceBroker.cpp must keep Query discoverable for source guard.");
    state.Require(contains(queryBody, "constexpr size_t kMaxClientBufferedCandidates") ||
                      contains(brokerSource, "constexpr size_t kMaxClientBufferedCandidates"),
                  L"Broker no-callback candidate accumulation must keep a configured count ceiling.");
    state.Require(contains(brokerSource, "constexpr uint64_t kMaxClientBufferedCandidateBytes"),
                  L"Broker no-callback candidate accumulation must keep a configured byte ceiling.");
    state.Require(containsBeforeCompact(queryBody, "CanAppendClientBufferedCandidates(outCandidates.size()", "outCandidates.insert("),
                  L"Broker no-callback path must validate count/bytes before accumulating candidates.");
    state.Require(contains(queryBody, "outCandidates.clear();") && contains(queryBody, "HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW)"),
                  L"Broker no-callback overflow must clear buffered candidates and return ERROR_BUFFER_OVERFLOW.");

    const std::string_view snapshotLoad = sourceBetween(
        localIndexSource, "class SnapshotVolumeStore final", "HRESULT Save(const PersistedVolumeState& state, QueryStats& stats) noexcept override");
    state.Require(! snapshotLoad.empty(), L"LocalSearchIndexCore.cpp must keep snapshot Load discoverable for source guard.");
    state.Require(containsBefore(snapshotLoad, "const uint64_t maxEntriesInBytes", "outState.entries.reserve(static_cast<size_t>(header.entryCount));"),
                  L"Snapshot load must derive a file-size entry ceiling before reserving snapshot entries.");
    state.Require(containsBefore(snapshotLoad,
                                 "if (header.entryCount > maxEntriesInBytes || header.entryCount > kMaxSizeT)",
                                 "outState.entries.reserve(static_cast<size_t>(header.entryCount));"),
                  L"Snapshot load must reject corrupt entryCount before reserve.");
    state.Require(contains(snapshotLoad, "HRESULT_FROM_WIN32(ERROR_INVALID_DATA)"),
                  L"Snapshot external-count preflight must reject corruption as ERROR_INVALID_DATA.");

    const std::string_view foldWideText = sourceBetween(searchTextSource, "[[nodiscard]] std::wstring FoldWideText", "struct RegexQuantifier final");
    state.Require(! foldWideText.empty(), L"SearchTextHelpers.cpp must keep FoldWideText discoverable for source guard.");
    state.Require(contains(foldWideText, "OrdinalString::FoldCaseInvariant(text)"),
                  L"SearchTextHelpers literal folding must route through the invariant fold helper.");
    state.Require(! contains(foldWideText, "CharLowerBuffW") && ! contains(foldWideText, "std::towlower"),
                  L"SearchTextHelpers literal folding must not use local lowercasing.");

    const std::string_view findResultKey =
        sourceBetween(findWindowSource, "[[nodiscard]] std::wstring ToLowerCopy", "[[nodiscard]] std::wstring FormatSearchStatusHint");
    state.Require(! findResultKey.empty(), L"FindFilesWindow.cpp must keep result-key folding discoverable for source guard.");
    state.Require(contains(findResultKey, "OrdinalString::FoldCaseInvariant(value)") && contains(findResultKey, "MakeResultKey"),
                  L"Find result identity keys must route through invariant folding.");
    state.Require(! contains(findResultKey, "std::towlower") && ! contains(findResultKey, "CharLowerBuffW"),
                  L"Find result identity keys must not use local lowercasing.");
    state.Require(contains(incrementalSearchSource, "OrdinalString::FindContainsFoldedInvariant") &&
                      contains(incrementalSearchSource, "OrdinalString::StartsWithFoldedInvariant") &&
                      contains(incrementalSearchSource, "const std::wstring foldedQuery = OrdinalString::FoldCaseInvariant(query);") &&
                      contains(incrementalSearchSource, "StartsWithPreFoldedNoCase(displayNameAt(index), foldedQuery)"),
                  L"FolderView incremental search folding helpers must use the shared invariant folded search helpers and pre-fold each prefix query once.");

    requireNoBroadCatch(L"SearchServiceBroker.cpp", brokerSource);
    requireNoBroadCatch(L"LocalSearchIndexCore.cpp", localIndexSource);
    requireNoBroadCatch(L"SqliteIndexStore.cpp", sqliteSource);
    requireNoBroadCatch(L"FileSystem.Search.cpp", pluginSearchSource);
    requireNoBroadCatch(L"SearchTextHelpers.cpp", searchTextSource);
    requireNoBroadCatch(L"FindFilesWindow.cpp", findWindowSource);
    requireBadAllocCatchesTerminate(L"SearchServiceBroker.cpp", brokerSource);
    requireBadAllocCatchesTerminate(L"LocalSearchIndexCore.cpp", localIndexSource);
    requireBadAllocCatchesTerminate(L"SqliteIndexStore.cpp", sqliteSource);
    requireBadAllocCatchesTerminate(L"FileSystem.Search.cpp", pluginSearchSource);
    requireBadAllocCatchesTerminate(L"FindFilesWindow.cpp", findWindowSource);

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_low_hardening_smoke",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto isHighSurrogate  = [](wchar_t ch) noexcept { return ch >= static_cast<wchar_t>(0xD800) && ch <= static_cast<wchar_t>(0xDBFF); };
    const auto isLowSurrogate   = [](wchar_t ch) noexcept { return ch >= static_cast<wchar_t>(0xDC00) && ch <= static_cast<wchar_t>(0xDFFF); };
    const auto hasLoneSurrogate = [&](std::wstring_view text) noexcept
    {
        for (size_t index = 0u; index < text.size(); ++index)
        {
            if (isHighSurrogate(text[index]) && (index + 1u >= text.size() || ! isLowSurrogate(text[index + 1u])))
            {
                return true;
            }
            if (isLowSurrogate(text[index]) && (index == 0u || ! isHighSurrogate(text[index - 1u])))
            {
                return true;
            }
        }
        return false;
    };

    std::wstring emoji;
    emoji.push_back(static_cast<wchar_t>(0xD83D));
    emoji.push_back(static_cast<wchar_t>(0xDE00));
    const std::wstring snippetText  = std::wstring(L"abc") + emoji + L"def";
    const std::wstring endSnippet   = SearchTextHelpers::BuildSnippet(snippetText, 0u, 1u, 4u);
    const std::wstring startSnippet = SearchTextHelpers::BuildSnippet(snippetText, 5u, 1u, 4u);
    state.Require(! hasLoneSurrogate(endSnippet), L"Search snippets must not split a surrogate pair at the end boundary.");
    state.Require(! hasLoneSurrogate(startSnippet), L"Search snippets must not split a surrogate pair at the start boundary.");
    state.Require(endSnippet.find(emoji) != std::wstring::npos && startSnippet.find(emoji) != std::wstring::npos,
                  L"Search snippets should preserve the complete surrogate pair when expanding boundaries.");

    const auto makeDirectoryInfoBuffer = [](std::wstring_view name) noexcept
    {
        const size_t byteCount = offsetof(FILE_FULL_DIR_INFO, FileName) + name.size() * sizeof(wchar_t);
        std::vector<std::byte> buffer(byteCount);
        auto* info           = reinterpret_cast<FILE_FULL_DIR_INFO*>(buffer.data());
        info->FileAttributes = FILE_ATTRIBUTE_ARCHIVE;
        info->FileNameLength = static_cast<ULONG>(name.size() * sizeof(wchar_t));
        if (! name.empty())
        {
            std::memcpy(info->FileName, name.data(), name.size() * sizeof(wchar_t));
        }
        return buffer;
    };

    std::vector<std::byte> validDirectoryInfo = makeDirectoryInfoBuffer(L"alpha.txt");
    state.Require(LocalSearchIndexCore::TryParseFileFullDirectoryInformationForTests(validDirectoryInfo.data(), validDirectoryInfo.size()),
                  L"Valid FILE_FULL_DIR_INFO should parse for the bounds-check smoke test.");

    std::vector<std::byte> oddNameLength                                        = validDirectoryInfo;
    reinterpret_cast<FILE_FULL_DIR_INFO*>(oddNameLength.data())->FileNameLength = 1u;
    state.Require(! LocalSearchIndexCore::TryParseFileFullDirectoryInformationForTests(oddNameLength.data(), oddNameLength.size()),
                  L"FILE_FULL_DIR_INFO with an odd byte FileNameLength must be rejected.");

    std::vector<std::byte> overrunName = validDirectoryInfo;
    reinterpret_cast<FILE_FULL_DIR_INFO*>(overrunName.data())->FileNameLength += static_cast<ULONG>(sizeof(wchar_t));
    state.Require(! LocalSearchIndexCore::TryParseFileFullDirectoryInformationForTests(overrunName.data(), overrunName.size()),
                  L"FILE_FULL_DIR_INFO with a filename beyond the entry bytes must be rejected.");

    std::vector<std::byte> shortNextEntry                                         = validDirectoryInfo;
    reinterpret_cast<FILE_FULL_DIR_INFO*>(shortNextEntry.data())->NextEntryOffset = static_cast<ULONG>(offsetof(FILE_FULL_DIR_INFO, FileName) - 1u);
    state.Require(! LocalSearchIndexCore::TryParseFileFullDirectoryInformationForTests(shortNextEntry.data(), shortNextEntry.size()),
                  L"FILE_FULL_DIR_INFO with a short NextEntryOffset must be rejected.");

    const auto contains         = [](std::string_view source, std::string_view needle) noexcept { return source.find(needle) != std::string_view::npos; };
    const auto countOccurrences = [](std::string_view source, std::string_view needle) noexcept
    {
        size_t count  = 0u;
        size_t offset = 0u;
        while (true)
        {
            offset = source.find(needle, offset);
            if (offset == std::string_view::npos)
            {
                return count;
            }
            ++count;
            offset += needle.size();
        }
    };
    const auto sourceBetween = [](std::string_view source, std::string_view beginMarker, std::string_view endMarker) noexcept -> std::string_view
    {
        const size_t beginOffset = source.find(beginMarker);
        if (beginOffset == std::string_view::npos)
        {
            return {};
        }

        const size_t endOffset = source.find(endMarker, beginOffset + beginMarker.size());
        if (endOffset == std::string_view::npos || endOffset <= beginOffset)
        {
            return {};
        }

        return source.substr(beginOffset, endOffset - beginOffset);
    };
    const auto containsBefore = [](std::string_view source, std::string_view before, std::string_view after) noexcept
    {
        const size_t beforePos = source.find(before);
        const size_t afterPos  = source.find(after);
        return beforePos != std::string_view::npos && afterPos != std::string_view::npos && beforePos < afterPos;
    };
    const auto compactWhitespace = [](std::string_view source)
    {
        std::string compact;
        compact.reserve(source.size());
        for (const char ch : source)
        {
            if (ch != ' ' && ch != '\t' && ch != '\r' && ch != '\n')
            {
                compact.push_back(ch);
            }
        }
        return compact;
    };
    const auto containsCompact = [&](std::string_view source, std::string_view needle)
    {
        const std::string compactSourceText = compactWhitespace(source);
        const std::string compactNeedle     = compactWhitespace(needle);
        return compactSourceText.find(compactNeedle) != std::string::npos;
    };
    const auto readSourceFile = [](const std::filesystem::path& path, std::string& outSource) noexcept
    {
        outSource.clear();
        std::ifstream stream(path, std::ios::binary);
        if (! stream)
        {
            return false;
        }

        outSource.assign(std::istreambuf_iterator<char>(stream), std::istreambuf_iterator<char>());
        return stream.eof() || stream.good();
    };

    const std::filesystem::path repoRoot = GetWorkspaceRootFromSourcePath();
    std::string pluginSearchSource;
    std::string fallbackSource;
    std::string sqliteSource;
    std::string localIndexSource;
    std::string localIndexHeaderSource;
    std::string searchServiceSource;
    std::string findWindowSource;
    std::string factorySource;
    std::string viewerWebSource;
    std::string compareEngineHeader;
    std::string compareEngineSource;
    std::string compareWindowSource;
    std::string dxUiControlsSource;
    std::string fileOpsDiagnosticsSource;
    state.Require(readSourceFile(repoRoot / L"Plugins" / L"FileSystem" / L"FileSystem.Search.cpp", pluginSearchSource),
                  L"Failed to read FileSystem.Search.cpp for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"RedSalamander" / L"SearchFallbackEngine.cpp", fallbackSource),
                  L"Failed to read SearchFallbackEngine.cpp for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"Common" / L"SqliteIndexStore.cpp", sqliteSource),
                  L"Failed to read SqliteIndexStore.cpp for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"Common" / L"LocalSearchIndexCore.cpp", localIndexSource),
                  L"Failed to read LocalSearchIndexCore.cpp for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"Common" / L"LocalSearchIndexCore.h", localIndexHeaderSource),
                  L"Failed to read LocalSearchIndexCore.h for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"Common" / L"SearchServiceBroker.cpp", searchServiceSource),
                  L"Failed to read SearchServiceBroker.cpp for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"RedSalamander" / L"FindFilesWindow.cpp", findWindowSource),
                  L"Failed to read FindFilesWindow.cpp for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"Common" / L"PlugInterfaces" / L"FactoryImpl.h", factorySource),
                  L"Failed to read FactoryImpl.h for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"Plugins" / L"ViewerWeb" / L"ViewerWeb.cpp", viewerWebSource),
                  L"Failed to read ViewerWeb.cpp for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"RedSalamander" / L"CompareDirectoriesEngine.h", compareEngineHeader),
                  L"Failed to read CompareDirectoriesEngine.h for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"RedSalamander" / L"CompareDirectoriesEngine.cpp", compareEngineSource),
                  L"Failed to read CompareDirectoriesEngine.cpp for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"RedSalamander" / L"CompareDirectoriesWindow.cpp", compareWindowSource),
                  L"Failed to read CompareDirectoriesWindow.cpp for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"Common" / L"DxUi" / L"DxUi.Controls.cpp", dxUiControlsSource),
                  L"Failed to read DxUi.Controls.cpp for low-hardening guard.");
    state.Require(readSourceFile(repoRoot / L"RedSalamander" / L"FolderWindow.FileOperations.State.Diagnostics.Part.cpp", fileOpsDiagnosticsSource),
                  L"Failed to read FolderWindow.FileOperations.State.Diagnostics.Part.cpp for low-hardening guard.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::string_view nativeMarkQueued =
        sourceBetween(pluginSearchSource, "[[nodiscard]] bool MarkQueuedDirectory", "[[nodiscard]] bool ShouldUseParallelDirectoryWalk");
    const std::string_view fallbackMarkQueued = sourceBetween(fallbackSource, "[[nodiscard]] bool MarkQueuedDirectory", "HRESULT CheckSearchCancelled");
    state.Require(contains(nativeMarkQueued, "runtime.queuedDirectories.erase(visitKey);"),
                  L"Native search must roll back the logical visit key when physical identity de-dup rejects a directory.");
    state.Require(contains(fallbackMarkQueued, "runtime.queuedDirectories.erase(visitKey);"),
                  L"Fallback search must roll back the logical visit key when physical identity de-dup rejects a directory.");
    state.Require(contains(nativeMarkQueued, "FAILED(TryGetDirectoryVisitIdentity(fullPath, identity))") &&
                      contains(nativeMarkQueued, "FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED") && contains(nativeMarkQueued, "return false;"),
                  L"Native followed-symlink traversal must skip a directory when physical identity cannot be probed.");
    state.Require(contains(pluginSearchSource, "MarkQueuedDirectory(runtime, fullPath, &warningFlags)") &&
                      contains(pluginSearchSource, "state.warningFlags.fetch_or(warningFlags"),
                  L"Native parallel search must fan identity-probe warnings through the parallel warning accumulator.");
    state.Require(contains(pluginSearchSource, "kMaxQueuedFollowSymlinkDirectories") &&
                      contains(nativeMarkQueued, "runtime.queuedDirectories.size() >= kMaxQueuedFollowSymlinkDirectories") &&
                      contains(nativeMarkQueued, "FILESYSTEM_SEARCH_WARNING_OVERFLOW"),
                  L"Native followed-symlink traversal must cap queued directory identity memory like the fallback engine.");
    state.Require(contains(fallbackMarkQueued, "FAILED(TryGetDirectoryVisitIdentity(fullPath, identity))") &&
                      contains(fallbackMarkQueued, "FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED") && contains(fallbackMarkQueued, "return false;"),
                  L"Fallback followed-symlink traversal must skip a directory when physical identity cannot be probed.");

    const std::string_view nativeParallelNameOnlyEntry =
        sourceBetween(pluginSearchSource, "void AccumulateParallelNameOnlyEntry", "[[nodiscard]] bool HasQueuedParallelDirectoryResultWork");
    const std::string_view nativeParallelNameOnlyMaterialize =
        sourceBetween(nativeParallelNameOnlyEntry, "if (! MaterializeEntryPaths(metadata, directoryFullPath", "ParallelDirectoryResult::Match match");
    state.Require(contains(nativeParallelNameOnlyMaterialize, "return;") && contains(nativeParallelNameOnlyMaterialize, "FILESYSTEM_SEARCH_WARNING_OVERFLOW") &&
                      ! contains(nativeParallelNameOnlyMaterialize, "result.status"),
                  L"Parallel name-only search must warn and skip malformed materialized entries instead of failing the whole directory walk.");
    const std::string_view nativeParallelDirectory = sourceBetween(pluginSearchSource, "void BuildParallelDirectoryResult", "void ParallelScanWorkerBody");
    const std::string_view nativeParallelDirectoryMaterialize =
        sourceBetween(nativeParallelDirectory, "if (! MaterializeEntryPaths(metadata, frame.fullPath", "EnqueueParallelDirectory");
    state.Require(contains(nativeParallelDirectoryMaterialize, "continue;") &&
                      contains(nativeParallelDirectoryMaterialize, "FILESYSTEM_SEARCH_WARNING_OVERFLOW") &&
                      ! contains(nativeParallelDirectoryMaterialize, "result.status"),
                  L"Parallel recursive directory enqueue must warn and skip malformed materialized entries instead of failing the whole directory walk.");
    const std::string_view nativeParallelHasQueued =
        sourceBetween(pluginSearchSource, "[[nodiscard]] bool HasQueuedParallelDirectoryResultWork", "void QueueParallelDirectoryResultChunk");
    state.Require(contains(nativeParallelHasQueued, "! result.matches.empty()") && ! contains(nativeParallelHasQueued, "SUCCEEDED(result.status)"),
                  L"Parallel name-only search must flush already accumulated matches when cancellation is observed mid-directory.");
    const std::string_view nativeParallelScanWorker =
        sourceBetween(pluginSearchSource, "void ParallelScanWorkerBody", "void CALLBACK ParallelScanWorkerCallback");
    const std::string_view nativeParallelWorkerContentMaterialize =
        sourceBetween(nativeParallelScanWorker, "if (! MaterializeEntryPaths(entry, chunk->directoryFullPath", "++result.candidateFiles");
    const std::string_view nativeParallelWorkerMatchMaterialize =
        sourceBetween(nativeParallelScanWorker, "uint32_t matchedBy = FILESYSTEM_SEARCH_MATCH_SOURCE_NONE", "result.matches.push_back");
    state.Require(contains(nativeParallelWorkerContentMaterialize, "FILESYSTEM_SEARCH_WARNING_OVERFLOW") &&
                      contains(nativeParallelWorkerContentMaterialize, "continue;") && ! contains(nativeParallelWorkerContentMaterialize, "return;") &&
                      contains(nativeParallelWorkerMatchMaterialize, "FILESYSTEM_SEARCH_WARNING_OVERFLOW") &&
                      contains(nativeParallelWorkerMatchMaterialize, "continue;") && ! contains(nativeParallelWorkerMatchMaterialize, "return;"),
                  L"Parallel chunk workers must warn and continue after malformed materialized entries instead of abandoning the chunk.");

    const std::string_view applyPragmas = sourceBetween(sqliteSource, "[[nodiscard]] HRESULT ApplyPragmas", "[[nodiscard]] HRESULT CreateSchemaVersion1");
    state.Require(! contains(applyPragmas, "\"VACUUM;\""), L"SQLite bootstrap pragmas must not run full VACUUM inline.");
    const std::string_view automaticMaintenance = sourceBetween(sqliteSource, "HRESULT RunAutomaticMaintenance", "HRESULT EnumerateVolume");
    state.Require(
        contains(automaticMaintenance, "const bool shouldRewriteAutoVacuum") && contains(automaticMaintenance, "\"VACUUM;\"") &&
            contains(automaticMaintenance, "RunQuickCheck(db.get(), L\"automatic auto-vacuum mode rewrite\")") &&
            contains(automaticMaintenance, "HRESULT_FROM_WIN32(ERROR_BUSY)") && containsCompact(automaticMaintenance, "result.ranVacuum = true"),
        L"Automatic SQLite maintenance must rewrite legacy non-incremental auto_vacuum stores outside bootstrap after quick_check and defer busy checkpoints.");
    const std::string_view manualMaintenance = sourceBetween(sqliteSource, "HRESULT RunManualMaintenance", "HRESULT RunAutomaticMaintenance");
    state.Require(contains(manualMaintenance, "RunQuickCheck(db.get(), L\"manual maintenance\")") && contains(manualMaintenance, "\"VACUUM;\""),
                  L"Manual SQLite maintenance must run quick_check before VACUUM.");

    const std::string_view snapshotSave = sourceBetween(
        localIndexSource, "HRESULT Save(const PersistedVolumeState& state, QueryStats& stats) noexcept override", "HRESULT Delete() noexcept override");
    const std::string_view snapshotSaveCleanup = sourceBetween(snapshotSave, "const auto deleteTempOnFailure", "SnapshotHeader header{}");
    state.Require(contains(snapshotSave, "OpenSnapshotTempFile(_snapshotPath, file, tempPath)") && contains(snapshotSave, "FlushFileBuffers(file.get())") &&
                      contains(snapshotSave, "const std::wstring extendedTempPath     = ToExtendedPath(tempPath)") &&
                      contains(snapshotSave, "const std::wstring extendedSnapshotPath = ToExtendedPath(_snapshotPath)") &&
                      contains(snapshotSave,
                               "MoveFileExW(extendedTempPath.c_str(), extendedSnapshotPath.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)") &&
                      ! contains(snapshotSave, "CREATE_ALWAYS"),
                  L"Snapshot persistence must write a flushed long-path-safe sibling temp and atomically replace the final snapshot.");
    state.Require(containsBefore(snapshotSaveCleanup, "file.reset();", "const std::wstring extendedTempPath = ToExtendedPath(tempPath)") &&
                      containsBefore(snapshotSaveCleanup,
                                     "const std::wstring extendedTempPath = ToExtendedPath(tempPath)",
                                     "static_cast<void>(::DeleteFileW(extendedTempPath.c_str()));"),
                  L"Snapshot temp cleanup must close the temp handle before deleting the long-path-safe sibling on failure.");
    const std::string_view snapshotDirectory = sourceBetween(localIndexSource, "HRESULT EnsureSnapshotDirectory", "HRESULT OpenSnapshotTempFile");
    const std::string_view snapshotTempOpen  = sourceBetween(localIndexSource, "HRESULT OpenSnapshotTempFile", "HRESULT WriteSnapshotBytes");
    const std::string_view snapshotLoad = sourceBetween(
        localIndexSource, "HRESULT Load(PersistedVolumeState& outState, QueryStats& stats) noexcept override", "HRESULT Save(const PersistedVolumeState");
    const std::string_view snapshotDelete = sourceBetween(localIndexSource, "HRESULT Delete() noexcept override", "#ifdef ENABLE_TESTS");
    const std::string_view snapshotCorrupt = sourceBetween(
        localIndexSource, "HRESULT CorruptForTests(SnapshotCorruptionMode mode) noexcept override", "#endif");
    state.Require(contains(snapshotDirectory, "ToExtendedPath(parent.native())") &&
                       contains(snapshotTempOpen, "extendedSnapshotPath = ToExtendedPath(snapshotPath)") &&
                       contains(snapshotTempOpen, "Common::Paths::CreateUniqueSiblingFile(extendedSnapshotPath") &&
                       contains(snapshotTempOpen, ".maximumAttempts = 16u") &&
                      contains(snapshotLoad, "extendedSnapshotPath = ToExtendedPath(_snapshotPath)") &&
                      contains(snapshotLoad, "CreateFileW(extendedSnapshotPath.c_str()") &&
                      contains(snapshotDelete, "DeleteFileW(extendedSnapshotPath.c_str())") &&
                      contains(snapshotCorrupt, "CreateFileW(extendedSnapshotPath.c_str()"),
                  L"Snapshot directory, temp-open, load, delete, and corruption operations must preserve extended-length paths.");
    state.Require(contains(localIndexSource, "stats.hardlinkAliasCoverageIncomplete = true") && contains(localIndexSource, "USN_REASON_HARD_LINK_CHANGE") &&
                      contains(localIndexHeaderSource, "hardlinkAliasCoverageIncomplete"),
                  L"Search indexing must report when hardlink/multi-path alias coverage is non-exhaustive.");
    state.Require(contains(searchServiceSource, "QUERY_COMPLETE_FLAG_HARDLINK_ALIAS_INCOMPLETE") &&
                      contains(searchServiceSource, "stats.hardlinkAliasCoverageIncomplete"),
                  L"Search service query stats must round-trip non-exhaustive hardlink alias coverage.");

    const std::string_view replaceVolume = sourceBetween(sqliteSource, "HRESULT ReplaceVolume", "HRESULT ApplyJournalDelta");
    state.Require(countOccurrences(replaceVolume, "UpdateVolume(db.get(), volumeId, request)") == 1u,
                  L"SqliteIndexStore::ReplaceVolume should update volume metadata exactly once per existing volume.");

    const std::string_view directoryInfoParser =
        sourceBetween(localIndexSource, "TryParseFileFullDirectoryInformationEntry", "[[nodiscard]] std::wstring ExtractVolumeRoot");
    state.Require(contains(directoryInfoParser, "entryBytes > remainingBytes") && contains(directoryInfoParser, "entryBytes - kFileNameOffset"),
                  L"FILE_FULL_DIR_INFO parsing must bound FileNameLength against the validated entry bytes before reading FileName.");

    const std::string_view warningSummary =
        sourceBetween(findWindowSource, "std::wstring FindFilesWindow::BuildWarningSummary", "UINT FindFilesWindow::BackendStringId");
    state.Require(contains(warningSummary, "FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED") && contains(warningSummary, "IDS_FIND_WARNING_REGEX_REJECTED"),
                  L"Find Files must surface rejected-regex warnings with a localized reason instead of only E_INVALIDARG.");
    state.Require(contains(warningSummary, "FILESYSTEM_SEARCH_WARNING_SERVICE_ROOT_REJECTED") &&
                      contains(warningSummary, "IDS_FIND_WARNING_SERVICE_ROOT_REJECTED"),
                  L"Find Files must distinguish a healthy-service root rejection from service transport unavailability.");

    const std::string_view factoryEnumerate = sourceBetween(factorySource, "[[nodiscard]] HRESULT FactoryEnumeratePlugins", "// FactoryCreate<TInterface>");
    state.Require(contains(factoryEnumerate, "entries.empty()") && contains(factoryEnumerate, "*count    = 0;") && contains(factoryEnumerate, "return S_OK;") &&
                      contains(factoryEnumerate, "entries[0]"),
                  L"FactoryEnumeratePlugins must handle an empty entries span before reading entries[0].");

    state.Require(countOccurrences(viewerWebSource, "<!doctype html><html><head><meta charset=\\\"utf-8\\\"") == 1u &&
                      countOccurrences(viewerWebSource, "kInternalHtmlHead") >= 5u,
                  L"ViewerWeb internal HTML head/CSP literal should stay hoisted to one shared constant.");

    state.Require(contains(compareEngineHeader, "_pendingContentCompareRetryAttempts") &&
                      contains(compareEngineHeader, "_pendingSubdirAggregateRetryAttempts") &&
                      contains(compareEngineSource, "kMaxPendingUpdateRetryAttempts") &&
                      contains(compareEngineSource, "if (! QueueContentUpdateFolderRetryLocked(folderKey, currentVersion))") &&
                      contains(compareEngineSource, "_pendingContentCompareRetryAttempts.erase(folderKey)") &&
                      contains(compareEngineSource, "_pendingSubdirAggregateRetryAttempts.erase(retryKey)"),
                  L"CompareDirectories pending content/subdir retries must be capped and cleared with their pending work.");
    state.Require(contains(compareWindowSource, "static const CompareDetailsTextStrings strings") && contains(compareWindowSource, "const auto& strings"),
                  L"CompareDirectories detail text resource strings must be cached and used by reference.");
    state.Require(contains(dxUiControlsSource, "const bool wasIndeterminate = _indeterminate") && contains(dxUiControlsSource, "if (wasIndeterminate)") &&
                      contains(dxUiControlsSource, "Invalidate(host);"),
                  L"DxUi checkbox keyboard handling must invalidate when clearing the indeterminate glyph without a value change.");
    state.Require(contains(fileOpsDiagnosticsSource, "const size_t overflow = _diagnosticsPendingFlush.size() - maxPendingFlush") &&
                      contains(fileOpsDiagnosticsSource, "_diagnosticsPendingFlush.erase(_diagnosticsPendingFlush.begin(),"),
                  L"FileOps diagnostics pending flush trimming must remove overflow entries in one bulk erase.");
    state.Require(contains(localIndexSource, "IsRedSalamanderStagedTempName") && contains(localIndexSource, ".rs_tmp_") &&
                      contains(localIndexSource, ".rs_copy_tmp_") && contains(localIndexSource, ".~rs-write-"),
                  L"Local search index hydration must skip RedSalamander staged temp sibling names.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_broker_stalled_query_honors_cancel",
                  [&](SelfTest::CaseState& state) noexcept
{
    struct TestFrameHeader final
    {
        uint32_t magic           = 0x53535252u;
        uint32_t protocolVersion = SearchServiceBroker::kProtocolVersion;
        uint32_t messageType     = 0u;
        uint32_t payloadBytes    = 0u;
    };

    const auto readExact = [](HANDLE handle, void* buffer, const uint32_t byteCount) noexcept -> HRESULT
    {
        uint8_t* destination = static_cast<uint8_t*>(buffer);
        uint32_t totalRead   = 0u;
        while (totalRead < byteCount)
        {
            DWORD chunkRead = 0u;
            if (::ReadFile(handle, destination + totalRead, byteCount - totalRead, &chunkRead, nullptr) == 0)
            {
                return HRESULT_FROM_WIN32(::GetLastError());
            }
            if (chunkRead == 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
            }
            totalRead += chunkRead;
        }
        return S_OK;
    };

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0,
                  L"Failed to override the search service pipe for stalled broker query test.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    wil::unique_event_nothrow requestReceived;
    requestReceived.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    wil::unique_event_nothrow releaseServer;
    releaseServer.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    wil::unique_event_nothrow clientDone;
    clientDone.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    state.Require(requestReceived && releaseServer && clientDone, L"Failed to create stalled broker synchronization events.");

    wil::unique_handle serverPipe(
        ::CreateNamedPipeW(pipeName.c_str(), PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT, 1u, 4096u, 4096u, 0u, nullptr));
    state.Require(! ! serverPipe, std::format(L"Failed to create stalled broker fake service pipe. error={}", static_cast<unsigned long>(::GetLastError())));
    if (! serverPipe || ! requestReceived || ! releaseServer || ! clientDone)
    {
        return false;
    }

    std::atomic<long> serverHr{S_OK};
    std::jthread server([&, pipe = std::move(serverPipe)](std::stop_token) mutable noexcept
    {
        if (::ConnectNamedPipe(pipe.get(), nullptr) == 0)
        {
            const DWORD error = ::GetLastError();
            if (error != ERROR_PIPE_CONNECTED)
            {
                serverHr.store(static_cast<long>(HRESULT_FROM_WIN32(error)));
                return;
            }
        }

        TestFrameHeader requestHeader{};
        HRESULT localHr = readExact(pipe.get(), &requestHeader, static_cast<uint32_t>(sizeof(requestHeader)));
        if (SUCCEEDED(localHr) && requestHeader.payloadBytes <= 64u * 1024u)
        {
            std::vector<std::byte> requestPayload(requestHeader.payloadBytes);
            if (requestHeader.payloadBytes != 0u)
            {
                localHr = readExact(pipe.get(), requestPayload.data(), requestHeader.payloadBytes);
            }
        }
        else if (SUCCEEDED(localHr))
        {
            localHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        serverHr.store(static_cast<long>(localHr));
        static_cast<void>(::SetEvent(requestReceived.get()));
        static_cast<void>(::WaitForSingleObject(releaseServer.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(5'000))));
    });

    std::atomic<bool> cancelRequested{false};
    std::atomic<long> clientHr{E_PENDING};
    const auto cancelThunk = [](void* cookie) noexcept -> HRESULT
    {
        const auto* cancelFlag = static_cast<const std::atomic<bool>*>(cookie);
        if (cancelFlag == nullptr)
        {
            return E_POINTER;
        }
        return cancelFlag->load(std::memory_order_acquire) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
    };

    std::jthread client([&](std::stop_token) noexcept
    {
        SearchServiceBroker::QueryRequest request{};
        request.rootPath           = L"C:\\";
        request.namePattern        = L"*";
        request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
        request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_INCLUDE_FILES);
        request.includeFiles       = true;
        request.includeDirectories = false;

        std::vector<LocalSearchIndexCore::Candidate> candidates;
        LocalSearchIndexCore::QueryStats stats{};
        const HRESULT hr = SearchServiceBroker::Query(request, nullptr, nullptr, cancelThunk, &cancelRequested, candidates, &stats);
        clientHr.store(static_cast<long>(hr), std::memory_order_release);
        static_cast<void>(::SetEvent(clientDone.get()));
    });

    state.Require(::WaitForSingleObject(requestReceived.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(3'000))) == WAIT_OBJECT_0,
                  L"Stalled broker fake service did not receive the query request.");
    const HRESULT serverResult = static_cast<HRESULT>(serverHr.load(std::memory_order_acquire));
    state.Require(SUCCEEDED(serverResult),
                  std::format(L"Stalled broker fake service failed while reading the request. hr=0x{:08X}", static_cast<unsigned long>(serverResult)));

    cancelRequested.store(true, std::memory_order_release);
    const DWORD cancelWait = ::WaitForSingleObject(clientDone.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(500)));
    state.Require(cancelWait == WAIT_OBJECT_0, L"Stalled broker query did not honor cancellation within the bounded wait.");

    static_cast<void>(::SetEvent(releaseServer.get()));
    if (cancelWait != WAIT_OBJECT_0)
    {
        static_cast<void>(::WaitForSingleObject(clientDone.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(5'000))));
    }
    client.join();
    server.join();

    if (cancelWait == WAIT_OBJECT_0)
    {
        const HRESULT queryHr = static_cast<HRESULT>(clientHr.load(std::memory_order_acquire));
        state.Require(queryHr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                      std::format(L"Stalled broker query expected ERROR_CANCELLED after cancellation. hr=0x{:08X}", static_cast<unsigned long>(queryHr)));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_broker_streaming_query_honors_cancel",
                  [&](SelfTest::CaseState& state) noexcept
{
    struct TestFrameHeader final
    {
        uint32_t magic           = 0x53535252u;
        uint32_t protocolVersion = SearchServiceBroker::kProtocolVersion;
        uint32_t messageType     = 0u;
        uint32_t payloadBytes    = 0u;
    };

    struct TestCandidateBatchHeader final
    {
        uint32_t count    = 0u;
        uint32_t reserved = 0u;
    };

    struct TestCandidateEntryHeader final
    {
        uint32_t fileAttributes     = 0u;
        uint32_t metadataFlags      = LocalSearchIndexCore::CANDIDATE_METADATA_NONE;
        uint32_t fullPathBytes      = 0u;
        uint32_t displayNameBytes   = 0u;
        int64_t creationTime100ns   = 0;
        int64_t lastAccessTime100ns = 0;
        int64_t endOfFile           = 0;
        int64_t lastWriteTime100ns  = 0;
        int64_t changeTime100ns     = 0;
        int64_t allocationSize      = 0;
    };

    struct BatchCancelState final
    {
        BatchCancelState()                                   = default;
        BatchCancelState(const BatchCancelState&)            = delete;
        BatchCancelState(BatchCancelState&&)                 = delete;
        BatchCancelState& operator=(const BatchCancelState&) = delete;
        BatchCancelState& operator=(BatchCancelState&&)      = delete;

        std::atomic<bool>* cancelRequested = nullptr;
        HANDLE batchConsumedEvent          = nullptr;
        std::atomic<uint32_t> batchesSeen{0u};
        std::atomic<uint32_t> candidatesSeen{0u};
    };

    const auto appendBytes = [](std::vector<std::byte>& buffer, const void* source, const size_t byteCount)
    {
        const auto* bytes = static_cast<const std::byte*>(source);
        buffer.insert(buffer.end(), bytes, bytes + byteCount);
    };

    const auto appendUtf16 = [&](std::vector<std::byte>& buffer, std::wstring_view text)
    {
        if (! text.empty())
        {
            appendBytes(buffer, text.data(), text.size() * sizeof(wchar_t));
        }
    };

    const auto makeQueryBatchPayload = [&](std::initializer_list<std::wstring_view> names)
    {
        std::vector<std::byte> payload;
        TestCandidateBatchHeader batchHeader{};
        batchHeader.count = static_cast<uint32_t>(names.size());
        appendBytes(payload, &batchHeader, sizeof(batchHeader));

        for (std::wstring_view name : names)
        {
            const std::wstring fullPath = std::wstring(L"C:\\broker-cancel\\") + std::wstring(name);
            TestCandidateEntryHeader entry{};
            entry.fileAttributes   = FILE_ATTRIBUTE_NORMAL;
            entry.fullPathBytes    = static_cast<uint32_t>(fullPath.size() * sizeof(wchar_t));
            entry.displayNameBytes = static_cast<uint32_t>(name.size() * sizeof(wchar_t));
            appendBytes(payload, &entry, sizeof(entry));
            appendUtf16(payload, fullPath);
            appendUtf16(payload, name);
        }

        return payload;
    };

    const auto readExact = [](HANDLE handle, void* buffer, const uint32_t byteCount) noexcept -> HRESULT
    {
        uint8_t* destination = static_cast<uint8_t*>(buffer);
        uint32_t totalRead   = 0u;
        while (totalRead < byteCount)
        {
            DWORD chunkRead = 0u;
            if (::ReadFile(handle, destination + totalRead, byteCount - totalRead, &chunkRead, nullptr) == 0)
            {
                return HRESULT_FROM_WIN32(::GetLastError());
            }
            if (chunkRead == 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
            }
            totalRead += chunkRead;
        }
        return S_OK;
    };

    const auto writeExact = [](HANDLE handle, const void* buffer, const uint32_t byteCount) noexcept -> HRESULT
    {
        const uint8_t* source = static_cast<const uint8_t*>(buffer);
        uint32_t totalWritten = 0u;
        while (totalWritten < byteCount)
        {
            DWORD chunkWritten = 0u;
            if (::WriteFile(handle, source + totalWritten, byteCount - totalWritten, &chunkWritten, nullptr) == 0)
            {
                return HRESULT_FROM_WIN32(::GetLastError());
            }
            if (chunkWritten == 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
            }
            totalWritten += chunkWritten;
        }
        return S_OK;
    };

    const auto sendFrame = [&](HANDLE handle, uint32_t messageType, std::span<const std::byte> payload) noexcept -> HRESULT
    {
        TestFrameHeader header{};
        header.messageType  = messageType;
        header.payloadBytes = static_cast<uint32_t>(payload.size_bytes());
        HRESULT hr          = writeExact(handle, &header, static_cast<uint32_t>(sizeof(header)));
        if (SUCCEEDED(hr) && ! payload.empty())
        {
            hr = writeExact(handle, payload.data(), static_cast<uint32_t>(payload.size_bytes()));
        }
        return hr;
    };

    const auto cancelThunk = [](void* cookie) noexcept -> HRESULT
    {
        auto* cancelChecks = static_cast<std::atomic<bool>*>(cookie);
        if (cancelChecks == nullptr)
        {
            return E_POINTER;
        }
        return cancelChecks->load(std::memory_order_acquire) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
    };

    const auto batchThunk = [](LocalSearchIndexCore::Candidate* candidates, size_t count, size_t* consumedCount, void* cookie) noexcept -> HRESULT
    {
        auto* batchState = static_cast<BatchCancelState*>(cookie);
        if (batchState == nullptr || batchState->cancelRequested == nullptr || batchState->batchConsumedEvent == nullptr || consumedCount == nullptr ||
            (count != 0u && candidates == nullptr))
        {
            return E_POINTER;
        }

        *consumedCount = count;
        batchState->batchesSeen.fetch_add(1u, std::memory_order_acq_rel);
        batchState->candidatesSeen.fetch_add(static_cast<uint32_t>(count), std::memory_order_acq_rel);
        batchState->cancelRequested->store(true, std::memory_order_release);
        static_cast<void>(::SetEvent(batchState->batchConsumedEvent));
        return S_OK;
    };

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0,
                  L"Failed to override the search service pipe for streaming broker cancel test.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    wil::unique_event_nothrow batchConsumed;
    batchConsumed.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    wil::unique_event_nothrow releaseServer;
    releaseServer.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    wil::unique_event_nothrow clientDone;
    clientDone.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    state.Require(batchConsumed && releaseServer && clientDone, L"Failed to create streaming broker cancel synchronization events.");

    wil::unique_handle serverPipe(
        ::CreateNamedPipeW(pipeName.c_str(), PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT, 1u, 4096u, 4096u, 0u, nullptr));
    state.Require(! ! serverPipe, std::format(L"Failed to create streaming broker fake service pipe. error={}", static_cast<unsigned long>(::GetLastError())));
    if (! serverPipe || ! batchConsumed || ! releaseServer || ! clientDone)
    {
        return false;
    }

    const std::vector<std::byte> firstBatchPayload = makeQueryBatchPayload({L"alpha.txt", L"beta.txt"});
    std::atomic<long> serverHr{S_OK};
    std::jthread server([&, pipe = std::move(serverPipe)](std::stop_token) mutable noexcept
    {
        if (::ConnectNamedPipe(pipe.get(), nullptr) == 0)
        {
            const DWORD error = ::GetLastError();
            if (error != ERROR_PIPE_CONNECTED)
            {
                serverHr.store(static_cast<long>(HRESULT_FROM_WIN32(error)), std::memory_order_release);
                return;
            }
        }

        TestFrameHeader requestHeader{};
        HRESULT localHr = readExact(pipe.get(), &requestHeader, static_cast<uint32_t>(sizeof(requestHeader)));
        if (SUCCEEDED(localHr) && requestHeader.payloadBytes <= 64u * 1024u)
        {
            std::vector<std::byte> requestPayload(requestHeader.payloadBytes);
            if (requestHeader.payloadBytes != 0u)
            {
                localHr = readExact(pipe.get(), requestPayload.data(), requestHeader.payloadBytes);
            }
        }
        else if (SUCCEEDED(localHr))
        {
            localHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        if (SUCCEEDED(localHr))
        {
            localHr = sendFrame(pipe.get(), 5u, std::span<const std::byte>(firstBatchPayload.data(), firstBatchPayload.size()));
        }
        if (SUCCEEDED(localHr) && ::FlushFileBuffers(pipe.get()) == 0)
        {
            localHr = HRESULT_FROM_WIN32(::GetLastError());
        }

        serverHr.store(static_cast<long>(localHr), std::memory_order_release);
        static_cast<void>(::WaitForSingleObject(releaseServer.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(5'000))));
    });

    std::atomic<bool> cancelRequested{false};
    BatchCancelState batchState{};
    batchState.cancelRequested    = &cancelRequested;
    batchState.batchConsumedEvent = batchConsumed.get();
    std::atomic<long> clientHr{E_PENDING};
    std::jthread client([&](std::stop_token) noexcept
    {
        SearchServiceBroker::QueryRequest request{};
        request.rootPath           = L"C:\\";
        request.namePattern        = L"*.txt";
        request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
        request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_INCLUDE_FILES);
        request.includeFiles       = true;
        request.includeDirectories = false;

        std::vector<LocalSearchIndexCore::Candidate> candidates;
        LocalSearchIndexCore::QueryStats stats{};
        const HRESULT hr = SearchServiceBroker::Query(request, nullptr, nullptr, cancelThunk, &cancelRequested, candidates, &stats, batchThunk, &batchState);
        clientHr.store(static_cast<long>(hr), std::memory_order_release);
        static_cast<void>(::SetEvent(clientDone.get()));
    });

    state.Require(::WaitForSingleObject(batchConsumed.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(3'000))) == WAIT_OBJECT_0,
                  L"Streaming broker fake service did not deliver the first batch to the client.");
    state.Require(::WaitForSingleObject(clientDone.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(500))) == WAIT_OBJECT_0,
                  L"Streaming broker query did not honor cancellation after the first batch.");

    static_cast<void>(::SetEvent(releaseServer.get()));
    client.join();
    server.join();

    const HRESULT serverResult = static_cast<HRESULT>(serverHr.load(std::memory_order_acquire));
    state.Require(SUCCEEDED(serverResult), std::format(L"Streaming broker fake service failed. hr=0x{:08X}", static_cast<unsigned long>(serverResult)));
    const HRESULT queryHr = static_cast<HRESULT>(clientHr.load(std::memory_order_acquire));
    state.Require(queryHr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Streaming broker query expected ERROR_CANCELLED after first batch. hr=0x{:08X}", static_cast<unsigned long>(queryHr)));
    state.Require(batchState.batchesSeen.load(std::memory_order_acquire) == 1u,
                  std::format(L"Streaming broker query should consume one batch before cancellation. batches={}",
                              batchState.batchesSeen.load(std::memory_order_acquire)));
    state.Require(batchState.candidatesSeen.load(std::memory_order_acquire) == 2u,
                  std::format(L"Streaming broker query should consume two candidates before cancellation. candidates={}",
                              batchState.candidatesSeen.load(std::memory_order_acquire)));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sqlite_index_store_bootstrap_creates_schema",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_index_store_bootstrap", caseRoot), L"Failed to prepare sqlite_index_store_bootstrap root.");

    const std::filesystem::path databasePath = caseRoot / L"index-v2.sqlite3";
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    const HRESULT bootstrapHr = SqliteIndexStore::EnsureBootstrap(databasePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(bootstrapHr), std::format(L"SqliteIndexStore::EnsureBootstrap failed. hr=0x{:08X}", static_cast<unsigned long>(bootstrapHr)));
    if (FAILED(bootstrapHr))
    {
        return false;
    }

    const uint32_t runtimeVersion = static_cast<uint32_t>(sqlite3_libversion_number());
    state.Require(
        runtimeVersion >= SqliteIndexStore::kMinimumWalSqliteVersionNumber,
        std::format(L"SQLite runtime version {} is below the accepted WAL minimum {}.", runtimeVersion, SqliteIndexStore::kMinimumWalSqliteVersionNumber));

    std::error_code existsEc;
    state.Require(std::filesystem::exists(databasePath, existsEc), std::format(L"SQLite database was not created: {}", databasePath.wstring()));
    state.Require(bootstrapInfo.schemaVersion == SqliteIndexStore::kSchemaVersion,
                  std::format(L"Unexpected SQLite schema version {}.", bootstrapInfo.schemaVersion));
    state.Require(bootstrapInfo.schemaReady, L"SQLite bootstrap should create the expected schema objects.");
    state.Require(bootstrapInfo.walEnabled, L"SQLite bootstrap should enable WAL mode.");
    state.Require(bootstrapInfo.foreignKeysEnabled, L"SQLite bootstrap should enable foreign keys.");
    state.Require(bootstrapInfo.incrementalAutoVacuumEnabled, L"SQLite bootstrap should enable incremental auto-vacuum.");
    state.Require(bootstrapInfo.databaseBytes > 0u, L"SQLite bootstrap should create a non-empty database file.");
    state.Require(EqualsIgnoreCase(bootstrapInfo.databasePath, databasePath.wstring()), L"SQLite bootstrap should report the requested database path.");
    state.Require(EqualsIgnoreCase(bootstrapInfo.writeAheadLogPath, databasePath.wstring() + L"-wal"),
                  L"SQLite bootstrap should report the matching WAL path.");

    const LocalSearchIndexCore::PersistentStoreInfo storeInfo =
        LocalSearchIndexCore::GetPersistentStoreInfo({.snapshotRootDirectory = caseRoot.wstring(),
                                                      .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                                                      .sqliteDatabasePath    = databasePath.wstring()});
    state.Require(storeInfo.kind == LocalSearchIndexCore::PersistentStoreKind::Sqlite, L"Persistent store info should report the SQLite backend.");
    state.Require(storeInfo.autoCheckpointEnabled, L"Persistent store info should advertise automatic checkpointing for SQLite.");
    state.Require(storeInfo.autoCompactionEnabled, L"Persistent store info should advertise automatic compaction for SQLite.");
    state.Require(storeInfo.autoCheckpointTargetBytes == LocalSearchIndexCore::GetDefaultSqliteMaintenancePolicy().autoCheckpointTargetBytes,
                  L"Persistent store info should reuse the default SQLite checkpoint target.");
    state.Require(storeInfo.autoCompactionFragmentationPercent == LocalSearchIndexCore::GetDefaultSqliteMaintenancePolicy().autoCompactionFragmentationPercent,
                  L"Persistent store info should reuse the default SQLite compaction threshold.");
    state.Require(storeInfo.autoCompactionMinBytes == LocalSearchIndexCore::GetDefaultSqliteMaintenancePolicy().autoCompactionMinBytes,
                  L"Persistent store info should reuse the default SQLite compaction reclaim threshold.");
    state.Require(EqualsIgnoreCase(storeInfo.primaryPath, bootstrapInfo.databasePath),
                  L"Persistent store info should point at the bootstrapped SQLite database.");

    SqliteIndexStore::StoreInfo inspectInfo{};
    const HRESULT inspectHr = SqliteIndexStore::InspectStore(databasePath.wstring(), inspectInfo);
    state.Require(SUCCEEDED(inspectHr), std::format(L"SqliteIndexStore::InspectStore failed. hr=0x{:08X}", static_cast<unsigned long>(inspectHr)));
    if (FAILED(inspectHr))
    {
        return false;
    }

    state.Require(inspectInfo.schemaReady, L"SQLite inspection should confirm the schema after bootstrap.");
    state.Require(inspectInfo.schemaVersion == SqliteIndexStore::kSchemaVersion, L"SQLite inspection should report the current schema version.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sqlite_index_store_manual_compaction_reclaims_space",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_index_store_manual_compaction", caseRoot),
                  L"Failed to prepare sqlite_index_store_manual_compaction root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path databasePath = caseRoot / L"index-v2.sqlite3";
    SqliteIndexStore::StoreInfo beforeInfo{};
    std::wstring prepareError;
    state.Require(PrepareSqliteMaintenanceStore(caseRoot, databasePath, beforeInfo, prepareError), prepareError);
    if (! state.failure.empty())
    {
        return false;
    }

    SqliteIndexStore::ManualMaintenanceResult maintenance{};
    const HRESULT maintenanceHr = SqliteIndexStore::RunManualMaintenance(databasePath.wstring(), &maintenance);
    state.Require(SUCCEEDED(maintenanceHr),
                  std::format(L"SqliteIndexStore::RunManualMaintenance failed. hr=0x{:08X}", static_cast<unsigned long>(maintenanceHr)));
    if (FAILED(maintenanceHr))
    {
        return false;
    }

    state.Require(maintenance.ranVacuum, L"Manual maintenance should run VACUUM.");
    state.Require(maintenance.before.freelistPageCount >= beforeInfo.freelistPageCount,
                  L"Manual maintenance should capture the fragmented pre-maintenance freelist state.");
    state.Require(maintenance.after.freelistPageCount < maintenance.before.freelistPageCount,
                  std::format(L"Expected manual maintenance to reduce freelist pages, before={} after={}.",
                              maintenance.before.freelistPageCount,
                              maintenance.after.freelistPageCount));
    state.Require(maintenance.after.databaseBytes <= maintenance.before.databaseBytes,
                  std::format(L"Expected manual maintenance to keep or shrink the database file, before={} after={}.",
                              maintenance.before.databaseBytes,
                              maintenance.after.databaseBytes));
    state.Require(maintenance.after.writeAheadLogBytes <= maintenance.before.writeAheadLogBytes,
                  std::format(L"Expected manual maintenance to keep or shrink the WAL file, before={} after={}.",
                              maintenance.before.writeAheadLogBytes,
                              maintenance.after.writeAheadLogBytes));
    state.Require(! maintenance.after.lastCheckpointUtc.empty(), L"Manual maintenance should record lastCheckpointUtc.");
    state.Require(! maintenance.after.lastCompactionUtc.empty(), L"Manual maintenance should record lastCompactionUtc.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sqlite_index_store_automatic_checkpoint_truncates_wal",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_index_store_automatic_checkpoint", caseRoot),
                  L"Failed to prepare sqlite_index_store_automatic_checkpoint root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path databasePath = caseRoot / L"index-v2.sqlite3";
    SqliteIndexStore::StoreInfo beforeInfo{};
    std::wstring prepareError;
    state.Require(PrepareSqliteMaintenanceStore(caseRoot, databasePath, beforeInfo, prepareError), prepareError);
    if (! state.failure.empty())
    {
        return false;
    }

    sqlite3* walSeedConnection = nullptr;
    state.Require(OpenSqliteWalSeedConnection(databasePath, walSeedConnection, prepareError), prepareError);
    if (! state.failure.empty())
    {
        return false;
    }
    const auto closeWalSeedConnection = wil::scope_exit([&]() noexcept
    {
        if (walSeedConnection != nullptr)
        {
            static_cast<void>(sqlite3_close(walSeedConnection));
        }
    });

    SqliteIndexStore::StoreInfo walSeedInfo{};
    const HRESULT walSeedInspectHr = SqliteIndexStore::InspectStore(databasePath.wstring(), walSeedInfo);
    state.Require(SUCCEEDED(walSeedInspectHr),
                  std::format(L"SqliteIndexStore::InspectStore after WAL seed failed. hr=0x{:08X}", static_cast<unsigned long>(walSeedInspectHr)));
    if (FAILED(walSeedInspectHr))
    {
        return false;
    }
    state.Require(walSeedInfo.writeAheadLogBytes != 0u,
                  std::format(L"Automatic checkpoint fixture should have WAL bytes before maintenance, got {}.", walSeedInfo.writeAheadLogBytes));
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::SqliteMaintenancePolicy policy{};
    policy.autoCheckpointTargetBytes          = 1u;
    policy.autoCompactionFragmentationPercent = std::numeric_limits<uint32_t>::max();
    policy.autoCompactionMinBytes             = std::numeric_limits<uint64_t>::max();

    SqliteIndexStore::AutomaticMaintenanceResult maintenance{};
    const HRESULT maintenanceHr = SqliteIndexStore::RunAutomaticMaintenance(databasePath.wstring(), policy, &maintenance);
    state.Require(SUCCEEDED(maintenanceHr),
                  std::format(L"SqliteIndexStore::RunAutomaticMaintenance (checkpoint) failed. hr=0x{:08X}", static_cast<unsigned long>(maintenanceHr)));
    if (FAILED(maintenanceHr))
    {
        return false;
    }

    state.Require(maintenance.maintenanceNeeded, L"Automatic checkpoint test should require maintenance.");
    state.Require(maintenance.ranCheckpoint, L"Automatic checkpoint test should run a WAL checkpoint.");
    state.Require(! maintenance.ranVacuum, L"Automatic checkpoint test should not run full VACUUM.");
    state.Require(! maintenance.ranIncrementalVacuum, L"Automatic checkpoint test should not run incremental vacuum.");
    state.Require(! maintenance.after.lastCheckpointUtc.empty(), L"Automatic checkpoint test should record lastCheckpointUtc.");
    state.Require(maintenance.after.lastCompactionUtc.empty(), L"Automatic checkpoint test should not record lastCompactionUtc.");
    state.Require(maintenance.after.writeAheadLogBytes <= maintenance.before.writeAheadLogBytes,
                  std::format(L"Automatic checkpoint should keep or shrink WAL, before={} after={}.",
                              maintenance.before.writeAheadLogBytes,
                              maintenance.after.writeAheadLogBytes));
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sqlite_index_store_automatic_compaction_is_bounded",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_index_store_automatic_compaction", caseRoot),
                  L"Failed to prepare sqlite_index_store_automatic_compaction root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path databasePath = caseRoot / L"index-v2.sqlite3";
    SqliteIndexStore::StoreInfo beforeInfo{};
    std::wstring prepareError;
    state.Require(PrepareSqliteMaintenanceStore(caseRoot, databasePath, beforeInfo, prepareError), prepareError);
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::SqliteMaintenancePolicy policy{};
    policy.autoCheckpointTargetBytes          = std::numeric_limits<uint64_t>::max();
    policy.autoCompactionFragmentationPercent = 1u;
    policy.autoCompactionMinBytes             = 1u;

    SqliteIndexStore::AutomaticMaintenanceResult maintenance{};
    const HRESULT maintenanceHr = SqliteIndexStore::RunAutomaticMaintenance(databasePath.wstring(), policy, &maintenance);
    state.Require(SUCCEEDED(maintenanceHr),
                  std::format(L"SqliteIndexStore::RunAutomaticMaintenance (compaction) failed. hr=0x{:08X}", static_cast<unsigned long>(maintenanceHr)));
    if (FAILED(maintenanceHr))
    {
        return false;
    }

    state.Require(maintenance.maintenanceNeeded, L"Automatic compaction test should require maintenance.");
    state.Require(! maintenance.ranVacuum, L"Automatic compaction test should not run full VACUUM for an already-incremental store.");
    state.Require(maintenance.ranIncrementalVacuum, L"Automatic compaction test should run incremental vacuum.");
    state.Require(maintenance.requestedVacuumPages != 0u, L"Automatic compaction test should request at least one vacuum page.");
    state.Require(maintenance.requestedVacuumPages <= 4096u,
                  std::format(L"Automatic compaction should stay bounded to one pass, got {} pages.", maintenance.requestedVacuumPages));
    state.Require(maintenance.reclaimedVacuumPages != 0u, L"Automatic compaction should reclaim at least one freelist page.");
    state.Require(maintenance.after.freelistPageCount < maintenance.before.freelistPageCount,
                  std::format(L"Automatic compaction should reduce freelist pages, before={} after={}.",
                              maintenance.before.freelistPageCount,
                              maintenance.after.freelistPageCount));
    state.Require(! maintenance.after.lastCompactionUtc.empty(), L"Automatic compaction should record lastCompactionUtc.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sqlite_index_store_automatic_rewrites_legacy_auto_vacuum",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_index_store_automatic_rewrites_auto_vacuum", caseRoot),
                  L"Failed to prepare sqlite_index_store_automatic_rewrites_auto_vacuum root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path databasePath = caseRoot / L"index-v2.sqlite3";
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    const HRESULT bootstrapHr = SqliteIndexStore::EnsureBootstrap(databasePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(bootstrapHr), std::format(L"SqliteIndexStore::EnsureBootstrap failed. hr=0x{:08X}", static_cast<unsigned long>(bootstrapHr)));
    state.Require(bootstrapInfo.incrementalAutoVacuumEnabled, L"Fresh SQLite maintenance fixture should start with incremental auto_vacuum enabled.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring sqliteError;
    constexpr std::string_view kMakeLegacyAutoVacuumStore = "PRAGMA journal_mode=DELETE;"
                                                            "PRAGMA auto_vacuum=NONE;"
                                                            "VACUUM;"
                                                            "PRAGMA journal_mode=WAL;";
    state.Require(ExecuteSqliteScript(databasePath, kMakeLegacyAutoVacuumStore, sqliteError), sqliteError);
    if (! state.failure.empty())
    {
        return false;
    }

    SqliteIndexStore::StoreInfo legacyInfo{};
    const HRESULT legacyInspectHr = SqliteIndexStore::InspectStore(databasePath.wstring(), legacyInfo);
    state.Require(
        SUCCEEDED(legacyInspectHr),
        std::format(L"SqliteIndexStore::InspectStore for legacy auto_vacuum fixture failed. hr=0x{:08X}", static_cast<unsigned long>(legacyInspectHr)));
    state.Require(! legacyInfo.incrementalAutoVacuumEnabled, L"Legacy auto_vacuum fixture should inspect as non-incremental before maintenance.");
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::SqliteMaintenancePolicy policy{};
    policy.autoCheckpointTargetBytes          = std::numeric_limits<uint64_t>::max();
    policy.autoCompactionFragmentationPercent = std::numeric_limits<uint32_t>::max();
    policy.autoCompactionMinBytes             = std::numeric_limits<uint64_t>::max();

    SqliteIndexStore::AutomaticMaintenanceResult maintenance{};
    const HRESULT maintenanceHr = SqliteIndexStore::RunAutomaticMaintenance(databasePath.wstring(), policy, &maintenance);
    state.Require(
        SUCCEEDED(maintenanceHr),
        std::format(L"SqliteIndexStore::RunAutomaticMaintenance (auto_vacuum rewrite) failed. hr=0x{:08X}", static_cast<unsigned long>(maintenanceHr)));
    if (FAILED(maintenanceHr))
    {
        return false;
    }

    state.Require(maintenance.maintenanceNeeded, L"Legacy auto_vacuum fixture should require automatic maintenance.");
    state.Require(! maintenance.before.incrementalAutoVacuumEnabled, L"Automatic maintenance should capture the legacy pre-rewrite auto_vacuum state.");
    state.Require(maintenance.ranVacuum, L"Automatic maintenance should run full VACUUM to persist incremental auto_vacuum mode.");
    state.Require(! maintenance.ranIncrementalVacuum, L"Automatic auto_vacuum rewrite should not also run incremental vacuum.");
    state.Require(maintenance.after.incrementalAutoVacuumEnabled, L"Automatic maintenance should persist incremental auto_vacuum mode.");
    state.Require(! maintenance.after.lastCompactionUtc.empty(), L"Automatic auto_vacuum rewrite should record lastCompactionUtc.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_rejects_device_root_and_continues",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_rejects_device_root", caseRoot), L"Failed to prepare search_service_rejects_device_root root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::wstring extraArgs            = std::format(L"--storage-root=\"{}\" --store-backend=snapshot", storageRoot.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, false, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = LR"(\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\)";
    request.namePattern        = L"*";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    std::vector<LocalSearchIndexCore::Candidate> candidates;
    LocalSearchIndexCore::QueryStats stats{};
    const HRESULT queryHr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, candidates, &stats);
    state.Require(FAILED(queryHr), L"Device-namespace search service query should be rejected.");
    state.Require(candidates.empty(), L"Rejected device-namespace query should not return candidates.");

    SearchServiceBroker::ServiceStatus status{};
    if (WaitForSearchServiceStatus(state, status, pipeName, L"device-root rejection recovery service"))
    {
        state.Require(status.pipeName == pipeName, L"Status after rejected device root should come from the isolated service.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_slow_partial_client_does_not_block_next_client",
                  [&](SelfTest::CaseState& state) noexcept
{
    struct TestFrameHeader final
    {
        uint32_t magic           = 0x53535252u;
        uint32_t protocolVersion = SearchServiceBroker::kProtocolVersion;
        uint32_t messageType     = 3u;
        uint32_t payloadBytes    = sizeof(uint32_t);
    };

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_slow_partial_client", caseRoot), L"Failed to prepare search_service_slow_partial_client root.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr wchar_t kServerFrameTimeoutEnvVar[] = L"REDSALAMANDER_SEARCH_SERVICE_SERVER_FRAME_TIMEOUT_MS";
    const std::wstring previousPipeOverride       = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousTimeoutOverride    = GetEnvVarTrimmed(kServerFrameTimeoutEnvVar);
    const std::wstring pipeName                   = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    state.Require(::SetEnvironmentVariableW(kServerFrameTimeoutEnvVar, L"100") != 0, L"Failed to override the search service server frame timeout.");
    const auto restoreOverrides = wil::scope_exit([&] noexcept
    {
        const wchar_t* restorePipeValue    = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        const wchar_t* restoreTimeoutValue = previousTimeoutOverride.empty() ? nullptr : previousTimeoutOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restorePipeValue));
        static_cast<void>(::SetEnvironmentVariableW(kServerFrameTimeoutEnvVar, restoreTimeoutValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::wstring extraArgs            = std::format(L"--storage-root=\"{}\" --store-backend=snapshot", storageRoot.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    wil::unique_handle slowClient(::CreateFileW(pipeName.c_str(), GENERIC_READ | GENERIC_WRITE, 0u, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    state.Require(! ! slowClient,
                  std::format(L"Failed to connect slow partial client to search service pipe. error={}", static_cast<unsigned long>(::GetLastError())));
    if (! slowClient)
    {
        return false;
    }

    TestFrameHeader header{};
    DWORD written = 0u;
    state.Require(::WriteFile(slowClient.get(), &header, static_cast<DWORD>(sizeof(header)), &written, nullptr) != 0 &&
                      written == static_cast<DWORD>(sizeof(header)),
                  L"Failed to write the slow partial client frame header.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::this_thread::sleep_for(std::chrono::milliseconds(350));

    SearchServiceBroker::ServiceStatus status{};
    if (! WaitForSearchServiceStatusWithProcessDiagnostics(state, status, pipeName, L"slow partial client recovery service", service))
    {
        return false;
    }
    state.Require(status.pipeName == pipeName, L"Status after slow partial client should come from the isolated service.");

    slowClient.reset();
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_legacy_auto_vacuum_queues_idle_maintenance",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_legacy_auto_vacuum_queue", caseRoot),
                  L"Failed to prepare search_service_sqlite_legacy_auto_vacuum_queue root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path databasePath = caseRoot / L"legacy-auto-vacuum.sqlite3";
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    const HRESULT bootstrapHr = SqliteIndexStore::EnsureBootstrap(databasePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(bootstrapHr), std::format(L"SqliteIndexStore::EnsureBootstrap failed. hr=0x{:08X}", static_cast<unsigned long>(bootstrapHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring sqliteError;
    constexpr std::string_view kMakeLegacyAutoVacuumStore = "PRAGMA journal_mode=DELETE;"
                                                            "PRAGMA auto_vacuum=NONE;"
                                                            "VACUUM;"
                                                            "PRAGMA journal_mode=WAL;";
    state.Require(ExecuteSqliteScript(databasePath, kMakeLegacyAutoVacuumStore, sqliteError), sqliteError);
    if (! state.failure.empty())
    {
        return false;
    }

    SqliteIndexStore::StoreInfo legacyInfo{};
    const HRESULT legacyInspectHr = SqliteIndexStore::InspectStore(databasePath.wstring(), legacyInfo);
    state.Require(
        SUCCEEDED(legacyInspectHr),
        std::format(L"SqliteIndexStore::InspectStore for legacy auto_vacuum queue fixture failed. hr=0x{:08X}", static_cast<unsigned long>(legacyInspectHr)));
    state.Require(! legacyInfo.incrementalAutoVacuumEnabled, L"Legacy auto_vacuum queue fixture should inspect as non-incremental.");
    state.Require(legacyInfo.freelistPageCount == 0u,
                  std::format(L"Legacy auto_vacuum queue fixture should not rely on freelist pressure. freelist={}", legacyInfo.freelistPageCount));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", databasePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, false, serviceError, false, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus queuedStatus{};
    if (! WaitForSearchServiceStatus(state, queuedStatus, pipeName, L"legacy auto_vacuum queued maintenance service"))
    {
        return false;
    }

    state.Require(queuedStatus.maintenanceQueued,
                  L"Legacy non-incremental auto_vacuum store should queue idle maintenance even without WAL/freelist pressure.");
    state.Require(! queuedStatus.maintenanceRunning, L"Legacy auto_vacuum maintenance should wait for the idle grace window before running.");

    std::this_thread::sleep_for(std::chrono::milliseconds{1'300});

    SearchServiceBroker::ServiceStatus completedStatus{};
    HRESULT lastStatusHr = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
    bool maintenanceCompleted = false;
    const auto completionDeadline = std::chrono::steady_clock::now() +
                                    std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(30'000))};
    while (std::chrono::steady_clock::now() < completionDeadline)
    {
        completedStatus = {};
        lastStatusHr     = SearchServiceBroker::GetStatus(completedStatus);
        if (SUCCEEDED(lastStatusHr) && ! completedStatus.maintenanceQueued && ! completedStatus.maintenanceRunning &&
            ! completedStatus.lastCompactionUtc.empty() && ! completedStatus.lastCheckpointUtc.empty())
        {
            maintenanceCompleted = true;
            break;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds{50});
    }
    state.Require(maintenanceCompleted,
                  std::format(L"Legacy auto_vacuum maintenance did not complete before the service-status deadline. hr=0x{:08X} queued={} running={} "
                              L"lastCompaction='{}' lastCheckpoint='{}'.",
                              static_cast<unsigned long>(lastStatusHr),
                              completedStatus.maintenanceQueued,
                              completedStatus.maintenanceRunning,
                              completedStatus.lastCompactionUtc,
                              completedStatus.lastCheckpointUtc));
    if (! maintenanceCompleted)
    {
        return false;
    }

    state.Require(! completedStatus.maintenanceQueued, L"Legacy auto_vacuum maintenance should clear the queued flag after completion.");
    state.Require(! completedStatus.lastCompactionUtc.empty(), L"Legacy auto_vacuum maintenance should record lastCompactionUtc.");
    state.Require(! completedStatus.lastCheckpointUtc.empty(), L"Legacy auto_vacuum maintenance should record lastCheckpointUtc.");

    SqliteIndexStore::StoreInfo completedInfo{};
    const HRESULT completedInspectHr = SqliteIndexStore::InspectStore(databasePath.wstring(), completedInfo);
    state.Require(SUCCEEDED(completedInspectHr),
                  std::format(L"SqliteIndexStore::InspectStore after legacy auto_vacuum maintenance failed. hr=0x{:08X}",
                              static_cast<unsigned long>(completedInspectHr)));
    state.Require(completedInfo.incrementalAutoVacuumEnabled, L"Legacy auto_vacuum maintenance should persist incremental auto_vacuum mode.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sqlite_index_store_upgrade_paths",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_index_store_upgrade_paths", caseRoot), L"Failed to prepare sqlite_index_store_upgrade_paths root.");

    const std::filesystem::path legacyPath = caseRoot / L"legacy.sqlite3";
    const std::wstring legacyRootPath      = (caseRoot / L"legacy-root").wstring();
    const std::u8string legacyRootPathUtf8 = std::filesystem::path(legacyRootPath).u8string();
    const std::string legacyRootPathSql(reinterpret_cast<const char*>(legacyRootPathUtf8.c_str()), legacyRootPathUtf8.size());
    std::wstring sqliteError;
    state.Require(
        ExecuteSqliteScript(
            legacyPath,
            std::format("PRAGMA user_version=1;"
                        "CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT NOT NULL);"
                        "CREATE TABLE IF NOT EXISTS volumes("
                        "    volume_id INTEGER PRIMARY KEY,"
                        "    root_path TEXT UNIQUE NOT NULL,"
                        "    fs_kind INTEGER NOT NULL,"
                        "    journal_id INTEGER NOT NULL,"
                        "    next_usn INTEGER NOT NULL,"
                        "    state INTEGER NOT NULL,"
                        "    entry_count INTEGER NOT NULL,"
                        "    last_seed_utc TEXT NULL,"
                        "    last_replay_utc TEXT NULL,"
                        "    last_error_hr INTEGER NOT NULL DEFAULT 0"
                        ");"
                        "CREATE TABLE IF NOT EXISTS entries("
                        "    volume_id INTEGER NOT NULL,"
                        "    file_id_low INTEGER NOT NULL,"
                        "    file_id_high INTEGER NOT NULL,"
                        "    parent_id_low INTEGER NOT NULL,"
                        "    parent_id_high INTEGER NOT NULL,"
                        "    full_path TEXT NOT NULL,"
                        "    full_path_folded TEXT NOT NULL,"
                        "    name TEXT NOT NULL,"
                        "    name_folded TEXT NOT NULL,"
                        "    extension_folded TEXT NOT NULL,"
                        "    attributes INTEGER NOT NULL,"
                        "    is_dir INTEGER NOT NULL,"
                        "    size_bytes INTEGER NOT NULL,"
                        "    write_time_100ns INTEGER NOT NULL,"
                        "    PRIMARY KEY(volume_id, file_id_low, file_id_high)"
                        ") WITHOUT ROWID;"
                        "CREATE INDEX IF NOT EXISTS idx_entries_name_folded ON entries(volume_id, name_folded);"
                        "CREATE INDEX IF NOT EXISTS idx_entries_extension_folded ON entries(volume_id, extension_folded, is_dir);"
                        "CREATE INDEX IF NOT EXISTS idx_entries_full_path_folded ON entries(volume_id, full_path_folded);"
                        "CREATE INDEX IF NOT EXISTS idx_entries_parent ON entries(volume_id, parent_id_low, parent_id_high);"
                        "INSERT INTO meta(key, value) VALUES('schema_version', '1');"
                        "INSERT INTO meta(key, value) VALUES('store_kind', 'sqlite-v1');"
                        "INSERT INTO volumes(volume_id, root_path, fs_kind, journal_id, next_usn, state, entry_count, last_error_hr)"
                        "VALUES(1, '{}', 1, 123, 456, {}, 1, 0);"
                        "INSERT INTO entries(volume_id, file_id_low, file_id_high, parent_id_low, parent_id_high, full_path, full_path_folded, "
                        "name, name_folded, extension_folded, attributes, is_dir, size_bytes, write_time_100ns)"
                        "VALUES(1, 10, 0, 0, 0, '{}\\\\alpha.txt', '{}\\\\alpha.txt', 'alpha.txt', 'alpha.txt', '.txt', 32, 0, 5, 133838640000000000);",
                        legacyRootPathSql,
                        SqliteIndexStore::kVolumeStateReady,
                        legacyRootPathSql,
                        legacyRootPathSql),
            sqliteError),
        sqliteError);
    if (! state.failure.empty())
    {
        return false;
    }

    SqliteIndexStore::StoreInfo upgradedInfo{};
    const HRESULT upgradeHr = SqliteIndexStore::EnsureBootstrap(legacyPath.wstring(), &upgradedInfo);
    state.Require(SUCCEEDED(upgradeHr),
                  std::format(L"SqliteIndexStore::EnsureBootstrap upgrade path failed. hr=0x{:08X}", static_cast<unsigned long>(upgradeHr)));
    if (FAILED(upgradeHr))
    {
        return false;
    }

    state.Require(upgradedInfo.schemaReady, L"Legacy SQLite store should upgrade to a ready schema.");
    state.Require(upgradedInfo.schemaVersion == SqliteIndexStore::kSchemaVersion, L"Legacy SQLite store should upgrade to the current schema version.");
    state.Require(upgradedInfo.volumeCount == 1u, std::format(L"Legacy SQLite upgrade should preserve one mirrored volume, got {}.", upgradedInfo.volumeCount));
    state.Require(upgradedInfo.legacyImportVolumeCount == 1u,
                  std::format(L"Legacy SQLite upgrade should mark ready v1 rows as pending legacy import, got {}.", upgradedInfo.legacyImportVolumeCount));

    const LocalSearchIndexCore::PersistentStoreInfo persistentStoreInfo = LocalSearchIndexCore::GetPersistentStoreInfo({
        .persistentStoreKind = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
        .sqliteDatabasePath  = legacyPath.wstring(),
    });
    state.Require(persistentStoreInfo.inspectionSucceeded, L"Upgraded legacy SQLite store should be inspectable through LocalSearchIndexCore.");
    state.Require(! persistentStoreInfo.readyForQueryCutover,
                  L"Upgraded legacy SQLite store should block query cutover until remirror/backfill clears the legacy-import state.");
    state.Require(persistentStoreInfo.legacyImportVolumeCount == 1u,
                  std::format(L"Expected LocalSearchIndexCore readiness inspection to report one pending legacy-import volume, got {}.",
                              persistentStoreInfo.legacyImportVolumeCount));

    const std::filesystem::path futurePath = caseRoot / L"future.sqlite3";
    state.Require(ExecuteSqliteScript(futurePath, "PRAGMA user_version=99;", sqliteError), sqliteError);
    if (! state.failure.empty())
    {
        return false;
    }

    SqliteIndexStore::StoreInfo futureInfo{};
    const HRESULT futureHr = SqliteIndexStore::EnsureBootstrap(futurePath.wstring(), &futureInfo);
    state.Require(
        futureHr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH),
        std::format(L"Expected future schema bootstrap to fail with ERROR_REVISION_MISMATCH, got hr=0x{:08X}.", static_cast<unsigned long>(futureHr)));
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sqlite_index_store_load_and_apply_journal_delta",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_index_store_load_and_apply_delta", caseRoot),
                  L"Failed to prepare sqlite_index_store_load_and_apply_delta root.");
    const std::filesystem::path databasePath = caseRoot / L"index-v2.sqlite3";

    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(databasePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::EnsureBootstrap for load/apply-delta test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const uint64_t rootId = 1u;
    SqliteIndexStore::ReplaceVolumeRequest replaceRequest{};
    replaceRequest.rootPath       = caseRoot.wstring();
    replaceRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replaceRequest.journalId      = 7u;
    replaceRequest.nextUsn        = 70u;
    replaceRequest.state          = SqliteIndexStore::kVolumeStateReady;
    replaceRequest.entries.push_back({
        .fileIdLow    = rootId,
        .fileIdHigh   = 0u,
        .parentIdLow  = 0u,
        .parentIdHigh = 0u,
        .fullPath     = caseRoot.wstring(),
        .name         = L"",
        .attributes   = FILE_ATTRIBUTE_DIRECTORY,
    });
    replaceRequest.entries.push_back({
        .fileIdLow           = 2u,
        .fileIdHigh          = 0u,
        .parentIdLow         = rootId,
        .parentIdHigh        = 0u,
        .fullPath            = (caseRoot / L"alpha.txt").wstring(),
        .name                = L"alpha.txt",
        .attributes          = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes           = 5u,
        .writeTime100ns      = 133838640000000000ull,
        .creationTime100ns   = 133838640000000000ull,
        .lastAccessTime100ns = 133838640000000000ull,
        .changeTime100ns     = 133838640000000000ull,
        .allocationSize      = 4096u,
    });

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(databasePath.wstring(), replaceRequest, &replaceResult);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::ReplaceVolume for load/apply-delta test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest loadedVolume{};
    hr = SqliteIndexStore::LoadVolume(databasePath.wstring(), caseRoot.wstring(), loadedVolume);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::LoadVolume failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(loadedVolume.rootPath == caseRoot.wstring(), L"LoadVolume should preserve the root path.");
    state.Require(loadedVolume.state == SqliteIndexStore::kVolumeStateReady, L"LoadVolume should preserve the ready state.");
    state.Require(loadedVolume.journalId == 7u, L"LoadVolume should preserve the journal id.");
    state.Require(loadedVolume.nextUsn == 70u, L"LoadVolume should preserve the next USN.");
    state.Require(loadedVolume.entries.size() == 2u, std::format(L"LoadVolume should hydrate 2 entries, got {}.", loadedVolume.entries.size()));

    SqliteIndexStore::ApplyJournalDeltaRequest deltaRequest{};
    deltaRequest.rootPath       = caseRoot.wstring();
    deltaRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    deltaRequest.journalId      = 7u;
    deltaRequest.nextUsn        = 71u;
    deltaRequest.state          = SqliteIndexStore::kVolumeStateReady;
    deltaRequest.deletedEntries.push_back({
        .fileIdLow  = 2u,
        .fileIdHigh = 0u,
    });
    deltaRequest.upsertEntries.push_back({
        .fileIdLow           = 3u,
        .fileIdHigh          = 0u,
        .parentIdLow         = rootId,
        .parentIdHigh        = 0u,
        .fullPath            = (caseRoot / L"beta.txt").wstring(),
        .name                = L"beta.txt",
        .attributes          = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes           = 4u,
        .writeTime100ns      = 133838640000010000ull,
        .creationTime100ns   = 133838640000010000ull,
        .lastAccessTime100ns = 133838640000010000ull,
        .changeTime100ns     = 133838640000010000ull,
        .allocationSize      = 4096u,
    });

    SqliteIndexStore::ApplyJournalDeltaResult deltaResult{};
    hr = SqliteIndexStore::ApplyJournalDelta(databasePath.wstring(), deltaRequest, &deltaResult);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::ApplyJournalDelta failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(deltaResult.deletedEntryCount == 1u,
                  std::format(L"ApplyJournalDelta should report one deleted entry, got {}.", deltaResult.deletedEntryCount));
    state.Require(deltaResult.upsertedEntryCount == 1u,
                  std::format(L"ApplyJournalDelta should report one upserted entry, got {}.", deltaResult.upsertedEntryCount));

    loadedVolume = {};
    hr           = SqliteIndexStore::LoadVolume(databasePath.wstring(), caseRoot.wstring(), loadedVolume);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::LoadVolume after delta failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(loadedVolume.nextUsn == 71u, std::format(L"ApplyJournalDelta should advance nextUsn to 71, got {}.", loadedVolume.nextUsn));
    state.Require(loadedVolume.entries.size() == 2u,
                  std::format(L"ApplyJournalDelta should leave 2 entries (root + beta), got {}.", loadedVolume.entries.size()));

    const auto hasAlpha = std::ranges::any_of(
        loadedVolume.entries, [&](const SqliteIndexStore::ImportedEntry& entry) noexcept { return EqualsIgnoreCase(entry.name, L"alpha.txt"); });
    const auto hasBeta = std::ranges::any_of(loadedVolume.entries,
                                             [&](const SqliteIndexStore::ImportedEntry& entry) noexcept { return EqualsIgnoreCase(entry.name, L"beta.txt"); });
    state.Require(! hasAlpha, L"ApplyJournalDelta should remove alpha.txt from the stored volume.");
    state.Require(hasBeta, L"ApplyJournalDelta should upsert beta.txt into the stored volume.");

    hr = SqliteIndexStore::DeleteVolume(databasePath.wstring(), caseRoot.wstring());
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::DeleteVolume before replay reseed failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ApplyJournalDeltaRequest replayAfterLossRequest{};
    replayAfterLossRequest.rootPath             = caseRoot.wstring();
    replayAfterLossRequest.fileSystemKind       = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replayAfterLossRequest.journalId            = 7u;
    replayAfterLossRequest.nextUsn              = 72u;
    replayAfterLossRequest.state                = SqliteIndexStore::kVolumeStateReady;
    replayAfterLossRequest.deletedEntries       = deltaRequest.deletedEntries;
    replayAfterLossRequest.upsertEntries        = deltaRequest.upsertEntries;
    replayAfterLossRequest.seedEntriesIfMissing = loadedVolume.entries;

    SqliteIndexStore::ApplyJournalDeltaResult replayAfterLossResult{};
    hr = SqliteIndexStore::ApplyJournalDelta(databasePath.wstring(), replayAfterLossRequest, &replayAfterLossResult);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::ApplyJournalDelta after volume loss failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(replayAfterLossResult.insertedNewVolume, L"ApplyJournalDelta should reseed a missing volume when a snapshot seed is provided.");

    loadedVolume = {};
    hr           = SqliteIndexStore::LoadVolume(databasePath.wstring(), caseRoot.wstring(), loadedVolume);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::LoadVolume after replay reseed failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(loadedVolume.nextUsn == 72u, std::format(L"Replay reseed should advance nextUsn to 72, got {}.", loadedVolume.nextUsn));
    state.Require(loadedVolume.entries.size() == 2u,
                  std::format(L"Replay reseed should restore 2 entries (root + beta), got {}.", loadedVolume.entries.size()));

    const auto hasAlphaAfterReplay = std::ranges::any_of(
        loadedVolume.entries, [&](const SqliteIndexStore::ImportedEntry& entry) noexcept { return EqualsIgnoreCase(entry.name, L"alpha.txt"); });
    const auto hasBetaAfterReplay = std::ranges::any_of(
        loadedVolume.entries, [&](const SqliteIndexStore::ImportedEntry& entry) noexcept { return EqualsIgnoreCase(entry.name, L"beta.txt"); });
    state.Require(! hasAlphaAfterReplay, L"Replay reseed should not restore deleted alpha.txt.");
    state.Require(hasBetaAfterReplay, L"Replay reseed should preserve beta.txt from the snapshot seed.");

    SqliteIndexStore::VolumeInfo volumeInfo{};
    hr = SqliteIndexStore::InspectVolume(databasePath.wstring(), caseRoot.wstring(), volumeInfo);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::InspectVolume after delta failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(volumeInfo.entryCount == 2u, std::format(L"InspectVolume should report 2 entries after delta, got {}.", volumeInfo.entryCount));
    state.Require(volumeInfo.nextUsn == 72u, std::format(L"InspectVolume should report nextUsn=72 after replay reseed, got {}.", volumeInfo.nextUsn));

    SqliteIndexStore::QueryRequest queryRequest{};
    queryRequest.rootPath     = caseRoot.wstring();
    queryRequest.namePattern  = L"*.txt";
    queryRequest.nameMode     = FILESYSTEM_SEARCH_NAME_WILDCARD;
    queryRequest.recursive    = true;
    queryRequest.includeFiles = true;

    std::vector<std::wstring> queryNames;
    SqliteIndexStore::QueryRuntimeStats queryStats{};
    hr = SqliteIndexStore::EnumerateVolume(databasePath.wstring(),
                                           queryRequest,
                                           nullptr,
                                           nullptr,
                                           [](LocalSearchIndexCore::Candidate* candidate, void* cookie) noexcept -> HRESULT
    {
        if (candidate == nullptr || cookie == nullptr)
        {
            return E_POINTER;
        }

        try
        {
            static_cast<std::vector<std::wstring>*>(cookie)->push_back(candidate->displayName);
            return S_OK;
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
        catch (const std::exception&)
        {
            // Mandatory: this callback is noexcept; report append failure instead of unwinding through SQLite enumeration.
            Debug::Error(L"CompareDirectories selftest: replay reseed EnumerateVolume callback failed with an unexpected std::exception.");
            return E_FAIL;
        }
    },
                                           &queryNames,
                                           &queryStats);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::EnumerateVolume after delta failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(queryNames == std::vector<std::wstring>{L"beta.txt"}, L"EnumerateVolume after delta should return only beta.txt.");
    state.Require(queryStats.nextUsn == 72u, std::format(L"EnumerateVolume should expose nextUsn=72 after delta, got {}.", queryStats.nextUsn));
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sqlite_index_store_root_lookup_case_insensitive",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_index_store_root_lookup_case", caseRoot),
                  L"Failed to prepare sqlite_index_store_root_lookup_case root.");
    const std::filesystem::path databasePath = caseRoot / L"index-v2.sqlite3";

    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(databasePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::EnsureBootstrap for root case test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring storedRoot = caseRoot.wstring();
    std::wstring queryRoot        = storedRoot;
    for (wchar_t& ch : queryRoot)
    {
        if (ch >= L'a' && ch <= L'z')
        {
            ch = static_cast<wchar_t>(ch - (L'a' - L'A'));
        }
        else if (ch >= L'A' && ch <= L'Z')
        {
            ch = static_cast<wchar_t>(ch + (L'a' - L'A'));
        }
    }
    state.Require(queryRoot != storedRoot, L"Root case test requires a path with ASCII letters.");

    SqliteIndexStore::ReplaceVolumeRequest replaceRequest{};
    replaceRequest.rootPath       = storedRoot;
    replaceRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replaceRequest.journalId      = 101u;
    replaceRequest.nextUsn        = 202u;
    replaceRequest.state          = SqliteIndexStore::kVolumeStateReady;
    replaceRequest.entries.push_back({
        .fileIdLow    = 1u,
        .fileIdHigh   = 0u,
        .parentIdLow  = 0u,
        .parentIdHigh = 0u,
        .fullPath     = storedRoot,
        .name         = std::filesystem::path(storedRoot).filename().wstring(),
        .attributes   = FILE_ATTRIBUTE_DIRECTORY,
    });
    replaceRequest.entries.push_back({
        .fileIdLow    = 2u,
        .fileIdHigh   = 0u,
        .parentIdLow  = 1u,
        .parentIdHigh = 0u,
        .fullPath     = (caseRoot / L"Alpha.txt").wstring(),
        .name         = L"Alpha.txt",
        .attributes   = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes    = 5u,
    });

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(databasePath.wstring(), replaceRequest, &replaceResult);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::ReplaceVolume for root case test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::VolumeInfo volumeInfo{};
    hr = SqliteIndexStore::InspectVolume(databasePath.wstring(), queryRoot, volumeInfo);
    state.Require(SUCCEEDED(hr), std::format(L"InspectVolume should find differently-cased root. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(volumeInfo.entryCount == 2u, std::format(L"InspectVolume should report 2 entries, got {}.", volumeInfo.entryCount));

    SqliteIndexStore::ReplaceVolumeRequest loadedVolume{};
    hr = SqliteIndexStore::LoadVolume(databasePath.wstring(), queryRoot, loadedVolume);
    state.Require(SUCCEEDED(hr), std::format(L"LoadVolume should find differently-cased root. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(loadedVolume.rootPath == queryRoot, L"LoadVolume should preserve the caller's display root casing in its response.");
    state.Require(loadedVolume.entries.size() == 2u, std::format(L"LoadVolume should return 2 entries, got {}.", loadedVolume.entries.size()));

    SqliteIndexStore::QueryRequest queryRequest{};
    queryRequest.rootPath     = queryRoot;
    queryRequest.namePattern  = L"*.txt";
    queryRequest.nameMode     = FILESYSTEM_SEARCH_NAME_WILDCARD;
    queryRequest.recursive    = true;
    queryRequest.includeFiles = true;

    std::vector<std::wstring> queryNames;
    hr = SqliteIndexStore::EnumerateVolume(databasePath.wstring(),
                                           queryRequest,
                                           nullptr,
                                           nullptr,
                                           [](LocalSearchIndexCore::Candidate* candidate, void* cookie) noexcept -> HRESULT
    {
        if (candidate == nullptr || cookie == nullptr)
        {
            return E_POINTER;
        }

        try
        {
            static_cast<std::vector<std::wstring>*>(cookie)->push_back(candidate->displayName);
            return S_OK;
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
        catch (const std::exception&)
        {
            // Mandatory: this callback is noexcept; report append failure instead of unwinding through SQLite enumeration.
            Debug::Error(L"CompareDirectories selftest: root-case EnumerateVolume callback failed with an unexpected std::exception.");
            return E_FAIL;
        }
    },
                                           &queryNames,
                                           nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"EnumerateVolume should find differently-cased root. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(queryNames == std::vector<std::wstring>{L"Alpha.txt"}, L"EnumerateVolume should return the stored file under a differently-cased root.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_sqlite_option_keeps_snapshot_runtime_store",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_sqlite_option_runtime_store", caseRoot),
                  L"Failed to prepare index_core_sqlite_option_runtime_store root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    ULARGE_INTEGER expectedWriteTime{};
    expectedWriteTime.QuadPart = 133838640000000000ull;
    FILETIME expectedWriteTimeFile{};
    expectedWriteTimeFile.dwLowDateTime  = expectedWriteTime.LowPart;
    expectedWriteTimeFile.dwHighDateTime = expectedWriteTime.HighPart;
    state.Require(SetFileLastWriteTime(caseRoot / L"alpha.txt", expectedWriteTimeFile), L"Failed to set alpha.txt last write time.");
    ObservedFileMetadata expectedMetadata{};
    state.Require(TryReadObservedFileMetadata(caseRoot / L"alpha.txt", expectedMetadata), L"Failed to read expected alpha.txt metadata.");

    const std::filesystem::path storeRoot  = caseRoot / L"store";
    const std::filesystem::path sqlitePath = caseRoot / L"index-v2.sqlite3";
    LocalSearchIndexCore::Repository repository({
        .snapshotRootDirectory = storeRoot.wstring(),
        .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
        .sqliteDatabasePath    = sqlitePath.wstring(),
    });

    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", seedStats, seedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Seed indexed query with sqlite-configured repository failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(seedCandidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Seed indexed query with sqlite-configured repository returned an unexpected candidate set.");
    state.Require(! seedStats.snapshotPath.empty(), L"Snapshot-backed runtime store should report a snapshot path.");
    state.Require(seedStats.snapshotPath.ends_with(L".bin"), std::format(L"Expected a snapshot binary path, got '{}'.", seedStats.snapshotPath));
    state.Require(! EqualsIgnoreCase(seedStats.snapshotPath, sqlitePath.wstring()),
                  L"Snapshot mutation store should remain distinct from the SQLite query store.");

    const std::filesystem::path snapshotPath(seedStats.snapshotPath);
    std::error_code snapshotExistsEc;
    state.Require(std::filesystem::exists(snapshotPath, snapshotExistsEc), std::format(L"Snapshot runtime store was not created: {}", snapshotPath.wstring()));
    state.Require(EqualsIgnoreCase(snapshotPath.parent_path().lexically_normal().wstring(), storeRoot.lexically_normal().wstring()),
                  L"Snapshot runtime store should use the configured snapshot root directory.");

    std::error_code sqliteExistsEc;
    state.Require(std::filesystem::exists(sqlitePath, sqliteExistsEc), L"SQLite-configured repository should bootstrap and mirror the SQLite sidecar.");

    if (! CanProveDirectSqliteCurrentness(seedStats))
    {
        return state.Skip(DirectSqliteCurrentnessSkipReason(seedStats));
    }

    state.Require(seedStats.usedSqliteStore, L"SQLite-configured repository should use SQLite for query enumeration once the mirrored store is ready.");
    state.Require(! seedStats.sqliteCutoverBlocked, L"SQLite-configured repository should not report a blocked query cutover for a clean mirrored store.");

    hr = repository.DropCachedVolumeForTests(caseRoot.wstring());
    state.Require(SUCCEEDED(hr), std::format(L"DropCachedVolumeForTests failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    LocalSearchIndexCore::QueryStats reloadStats{};
    std::vector<LocalSearchIndexCore::Candidate> reloadCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", reloadStats, reloadCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Reload indexed query with sqlite-configured repository failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(! reloadStats.snapshotLoaded, L"Warm SQLite-backed reload should not need to reload the snapshot runtime store.");
    state.Require(reloadStats.usedSqliteStore, L"Reload indexed query should still enumerate results from SQLite.");
    state.Require(! reloadStats.sqliteCutoverBlocked, L"Reload indexed query should keep the SQLite cutover open.");
    state.Require(CollectIndexedCandidateNames(reloadCandidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Reload indexed query with sqlite-configured repository returned an unexpected candidate set.");
    state.Require(reloadCandidates.size() == 1u, std::format(L"Reload indexed query expected 1 candidate, got {}.", reloadCandidates.size()));
    if (reloadCandidates.size() == 1u)
    {
        const auto& candidate = reloadCandidates.front();
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_CREATION_TIME) != 0u,
                      L"Warm SQLite-backed reload should return persisted creation-time metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_LAST_ACCESS_TIME) != 0u,
                      L"Warm SQLite-backed reload should return persisted last-access metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_END_OF_FILE) != 0u,
                      L"Warm SQLite-backed reload should return persisted size metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_LAST_WRITE_TIME) != 0u,
                      L"Warm SQLite-backed reload should return persisted last-write metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_CHANGE_TIME) != 0u,
                      L"Warm SQLite-backed reload should return persisted change-time metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_ALLOCATION_SIZE) != 0u,
                      L"Warm SQLite-backed reload should return persisted allocation-size metadata.");
        state.Require(
            candidate.creationTime100ns == expectedMetadata.creationTime,
            std::format(L"Warm SQLite-backed reload creation-time mismatch. got={} expected={}", candidate.creationTime100ns, expectedMetadata.creationTime));
        state.Require(
            candidate.lastAccessTime100ns == expectedMetadata.lastAccessTime,
            std::format(L"Warm SQLite-backed reload last-access mismatch. got={} expected={}", candidate.lastAccessTime100ns, expectedMetadata.lastAccessTime));
        state.Require(candidate.endOfFile == 5, std::format(L"Warm SQLite-backed reload size mismatch. got={}", candidate.endOfFile));
        state.Require(candidate.lastWriteTime100ns == static_cast<int64_t>(expectedWriteTime.QuadPart),
                      std::format(L"Warm SQLite-backed reload last-write mismatch. got={} expected={}",
                                  candidate.lastWriteTime100ns,
                                  static_cast<long long>(expectedWriteTime.QuadPart)));
        state.Require(
            candidate.changeTime100ns == expectedMetadata.changeTime,
            std::format(L"Warm SQLite-backed reload change-time mismatch. got={} expected={}", candidate.changeTime100ns, expectedMetadata.changeTime));
        state.Require(
            candidate.allocationSize == expectedMetadata.allocationSize,
            std::format(L"Warm SQLite-backed reload allocation-size mismatch. got={} expected={}", candidate.allocationSize, expectedMetadata.allocationSize));
    }
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_sqlite_cold_start_bypasses_snapshot_runtime_store",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_sqlite_cold_start", caseRoot), L"Failed to prepare index_core_sqlite_cold_start root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");

    ULARGE_INTEGER expectedWriteTime{};
    expectedWriteTime.QuadPart = 133838640000000000ull;
    FILETIME expectedWriteTimeFile{};
    expectedWriteTimeFile.dwLowDateTime  = expectedWriteTime.LowPart;
    expectedWriteTimeFile.dwHighDateTime = expectedWriteTime.HighPart;
    state.Require(SetFileLastWriteTime(caseRoot / L"alpha.txt", expectedWriteTimeFile), L"Failed to set alpha.txt last write time.");
    ObservedFileMetadata expectedMetadata{};
    state.Require(TryReadObservedFileMetadata(caseRoot / L"alpha.txt", expectedMetadata), L"Failed to read expected alpha.txt metadata.");

    const std::filesystem::path storeRoot  = caseRoot / L"store";
    const std::filesystem::path sqlitePath = caseRoot / L"index-v2.sqlite3";

    {
        LocalSearchIndexCore::Repository seedRepository({
            .snapshotRootDirectory = storeRoot.wstring(),
            .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
            .sqliteDatabasePath    = sqlitePath.wstring(),
            .sqliteAuthoritative   = true,
        });

        LocalSearchIndexCore::QueryStats seedStats{};
        std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
        HRESULT hr = RunIndexedNameQuery(seedRepository, caseRoot.wstring(), L"*.txt", seedStats, seedCandidates);
        state.Require(SUCCEEDED(hr), std::format(L"Seed cold-start SQLite query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        state.Require(! seedStats.snapshotLoaded, L"Seed cold-start SQLite query should not load a legacy snapshot runtime store.");
        state.Require(! seedStats.snapshotSaved, L"Seed cold-start SQLite query should not create a legacy snapshot runtime store.");
        state.Require(seedStats.snapshotPath.empty(), L"Seed cold-start SQLite query should not report a snapshot path when SQLite seeds directly.");
        state.Require(
            seedStats.snapshotFileBytes == 0u,
            std::format(L"Seed cold-start SQLite query should report snapshotFileBytes=0 when no snapshot exists. got={}", seedStats.snapshotFileBytes));
        if (! CanProveDirectSqliteCurrentness(seedStats))
        {
            return state.Skip(DirectSqliteCurrentnessSkipReason(seedStats));
        }
    }

    auto snapshotFiles = CollectDirectoryFilesByExtension(storeRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Seed cold-start SQLite query should not leave a compatibility snapshot runtime file behind.");

    LocalSearchIndexCore::Repository coldRepository({
        .snapshotRootDirectory = storeRoot.wstring(),
        .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
        .sqliteDatabasePath    = sqlitePath.wstring(),
        .sqliteAuthoritative   = true,
    });

    LocalSearchIndexCore::QueryStats coldStats{};
    std::vector<LocalSearchIndexCore::Candidate> coldCandidates;
    HRESULT hr = RunIndexedNameQuery(coldRepository, caseRoot.wstring(), L"*.txt", coldStats, coldCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Cold-start SQLite query after snapshot removal failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(coldStats.usedSqliteStore, L"Cold-start SQLite query should enumerate from the ready SQLite store.");
    state.Require(! coldStats.snapshotLoaded, L"Cold-start SQLite query should not reload a deleted snapshot runtime store.");
    state.Require(! coldStats.snapshotSaved, L"Cold-start SQLite query should not recreate the snapshot runtime store.");
    state.Require(coldStats.snapshotPath.empty(), L"Cold-start SQLite query should not report a snapshot path when SQLite answers directly.");
    state.Require(coldStats.snapshotFileBytes == 0u,
                  std::format(L"Cold-start SQLite query should report snapshotFileBytes=0 when no snapshot exists. got={}", coldStats.snapshotFileBytes));
    state.Require(CollectIndexedCandidateNames(coldCandidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Cold-start SQLite query returned an unexpected candidate set.");
    state.Require(coldCandidates.size() == 1u, std::format(L"Cold-start SQLite query expected 1 candidate, got {}.", coldCandidates.size()));
    if (coldCandidates.size() == 1u)
    {
        const auto& candidate = coldCandidates.front();
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_CREATION_TIME) != 0u,
                      L"Cold-start SQLite query should keep persisted creation-time metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_LAST_ACCESS_TIME) != 0u,
                      L"Cold-start SQLite query should keep persisted last-access metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_END_OF_FILE) != 0u,
                      L"Cold-start SQLite query should keep persisted size metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_LAST_WRITE_TIME) != 0u,
                      L"Cold-start SQLite query should keep persisted last-write metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_CHANGE_TIME) != 0u,
                      L"Cold-start SQLite query should keep persisted change-time metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_ALLOCATION_SIZE) != 0u,
                      L"Cold-start SQLite query should keep persisted allocation-size metadata.");
        state.Require(
            candidate.creationTime100ns == expectedMetadata.creationTime,
            std::format(L"Cold-start SQLite query creation-time mismatch. got={} expected={}", candidate.creationTime100ns, expectedMetadata.creationTime));
        state.Require(
            candidate.lastAccessTime100ns == expectedMetadata.lastAccessTime,
            std::format(L"Cold-start SQLite query last-access mismatch. got={} expected={}", candidate.lastAccessTime100ns, expectedMetadata.lastAccessTime));
        state.Require(candidate.endOfFile == 5, std::format(L"Cold-start SQLite query size mismatch. got={}", candidate.endOfFile));
        state.Require(candidate.lastWriteTime100ns == static_cast<int64_t>(expectedWriteTime.QuadPart),
                      std::format(L"Cold-start SQLite query last-write mismatch. got={} expected={}",
                                  candidate.lastWriteTime100ns,
                                  static_cast<long long>(expectedWriteTime.QuadPart)));
        state.Require(candidate.changeTime100ns == expectedMetadata.changeTime,
                      std::format(L"Cold-start SQLite query change-time mismatch. got={} expected={}", candidate.changeTime100ns, expectedMetadata.changeTime));
        state.Require(
            candidate.allocationSize == expectedMetadata.allocationSize,
            std::format(L"Cold-start SQLite query allocation-size mismatch. got={} expected={}", candidate.allocationSize, expectedMetadata.allocationSize));
    }

    snapshotFiles = CollectDirectoryFilesByExtension(storeRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Cold-start SQLite query should not create a compatibility snapshot runtime file.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_sqlite_authoritative_replays_without_snapshot_runtime_store",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_sqlite_authoritative_replay", caseRoot),
                  L"Failed to prepare index_core_sqlite_authoritative_replay root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path storeRoot  = caseRoot / L"store";
    const std::filesystem::path sqlitePath = caseRoot / L"authoritative.sqlite3";
    LocalSearchIndexCore::Repository repository({
        .snapshotRootDirectory = storeRoot.wstring(),
        .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
        .sqliteDatabasePath    = sqlitePath.wstring(),
        .sqliteAuthoritative   = true,
    });

    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", seedStats, seedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Authoritative sqlite seed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(seedCandidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Authoritative sqlite seed query returned an unexpected candidate set.");
    state.Require(! seedStats.snapshotLoaded, L"Authoritative sqlite seed query should not load a compatibility snapshot.");
    state.Require(! seedStats.snapshotSaved, L"Authoritative sqlite seed query should not save a compatibility snapshot.");
    state.Require(seedStats.snapshotPath.empty(), L"Authoritative sqlite seed query should not report a compatibility snapshot path.");
    state.Require(seedStats.snapshotFileBytes == 0u, L"Authoritative sqlite seed query should report snapshotFileBytes=0.");

    auto snapshotFiles = CollectDirectoryFilesByExtension(storeRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Authoritative sqlite seed query should not create a compatibility snapshot runtime file.");

    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.txt", "beta"), L"Failed to create beta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    hr = repository.DropCachedVolumeForTests(caseRoot.wstring());
    state.Require(SUCCEEDED(hr), std::format(L"DropCachedVolumeForTests for authoritative sqlite replay failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    LocalSearchIndexCore::QueryStats replayStats{};
    std::vector<LocalSearchIndexCore::Candidate> replayCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", replayStats, replayCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Authoritative sqlite replay query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(replayCandidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"Authoritative sqlite replay query returned an unexpected candidate set.");
    state.Require(! replayStats.snapshotLoaded, L"Authoritative sqlite replay query should not load a compatibility snapshot.");
    state.Require(! replayStats.snapshotSaved, L"Authoritative sqlite replay query should not save a compatibility snapshot.");
    state.Require(replayStats.snapshotPath.empty(), L"Authoritative sqlite replay query should not report a compatibility snapshot path.");
    state.Require(replayStats.snapshotFileBytes == 0u, L"Authoritative sqlite replay query should report snapshotFileBytes=0.");
    if (! CanProveDirectSqliteCurrentness(replayStats))
    {
        return state.Skip(DirectSqliteCurrentnessSkipReason(replayStats));
    }

    state.Require(replayStats.journalReplayApplied, L"Authoritative sqlite replay query should advance the sqlite store through journal replay.");

    SqliteIndexStore::VolumeInfo volumeInfo{};
    hr = SqliteIndexStore::InspectVolume(sqlitePath.wstring(), caseRoot.wstring(), volumeInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::InspectVolume after authoritative replay failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(volumeInfo.state == SqliteIndexStore::kVolumeStateReady,
                  std::format(L"Authoritative sqlite replay should leave the stored root READY, got {}.", volumeInfo.state));
    state.Require(volumeInfo.nextUsn == replayStats.nextUsn,
                  std::format(L"Authoritative sqlite replay should persist the advanced nextUsn. stored={} stats={}", volumeInfo.nextUsn, replayStats.nextUsn));

    snapshotFiles = CollectDirectoryFilesByExtension(storeRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Authoritative sqlite replay query should still avoid compatibility snapshot runtime files.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_sqlite_authoritative_replays_directory_moves",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_sqlite_authoritative_directory_moves", caseRoot),
                  L"Failed to prepare index_core_sqlite_authoritative_directory_moves root.");
    std::filesystem::path outsideRoot;
    if (! prepareJournalDirectoryMoveFixture(state, caseRoot, L"index_core_sqlite_authoritative_directory_moves_outside", outsideRoot))
    {
        return false;
    }

    const std::filesystem::path storeRoot  = caseRoot / L"store";
    const std::filesystem::path sqlitePath = caseRoot / L"authoritative.sqlite3";
    LocalSearchIndexCore::Repository repository({
        .snapshotRootDirectory = storeRoot.wstring(),
        .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
        .sqliteDatabasePath    = sqlitePath.wstring(),
        .sqliteAuthoritative   = true,
    });

    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", seedStats, seedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Authoritative sqlite directory-move seed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(CollectIndexedCandidateNames(seedCandidates) == std::vector<std::wstring>{L"inside-child.txt", L"moved-out-child.txt", L"stable.txt"},
                  L"Authoritative sqlite directory-move seed query returned an unexpected candidate set.");
    state.Require(! seedStats.snapshotLoaded, L"Authoritative sqlite directory-move seed should not load a compatibility snapshot.");
    state.Require(! seedStats.snapshotSaved, L"Authoritative sqlite directory-move seed should not create a compatibility snapshot.");

    if (! applyJournalDirectoryMoveMutations(state, caseRoot, outsideRoot))
    {
        return false;
    }

    const std::vector<LocalSearchIndexCore::SyntheticJournalRecordForTests> syntheticRecords = buildJournalDirectoryMoveSyntheticRecords(caseRoot, outsideRoot);
    LocalSearchIndexCore::QueryStats replayStats{};
    hr = repository.ApplySyntheticJournalForTests(caseRoot.wstring(), syntheticRecords, &replayStats);
    state.Require(SUCCEEDED(hr), std::format(L"Authoritative sqlite directory-move synthetic replay failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(! replayStats.snapshotLoaded, L"Authoritative sqlite directory-move replay should not load a compatibility snapshot.");
    state.Require(! replayStats.snapshotSaved, L"Authoritative sqlite directory-move replay should not create a compatibility snapshot.");
    state.Require(replayStats.snapshotPath.empty(), L"Authoritative sqlite directory-move replay should not report a compatibility snapshot path.");

    SqliteIndexStore::ReplaceVolumeRequest storedVolume{};
    hr = SqliteIndexStore::LoadVolume(sqlitePath.wstring(), caseRoot.wstring(), storedVolume);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::LoadVolume after directory-move replay failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(collectSqliteStoredFileNames(storedVolume) == expectedJournalDirectoryMoveNames(),
                  L"Authoritative sqlite directory-move delta should persist the replayed file set.");
    state.Require(hasSqliteEntryFullPath(storedVolume, caseRoot / L"renamed-inside-dir" / L"inside-child.txt"),
                  L"Authoritative sqlite directory-move delta should persist the in-tree renamed child at the new path.");
    state.Require(! hasSqliteEntryFullPath(storedVolume, caseRoot / L"inside-dir" / L"inside-child.txt"),
                  L"Authoritative sqlite directory-move delta should delete the old in-tree renamed path.");
    state.Require(hasSqliteEntryFullPath(storedVolume, caseRoot / L"moved-in-dir" / L"nested" / L"moved-in-nested.txt"),
                  L"Authoritative sqlite directory-move delta should persist hydrated moved-in children.");
    state.Require(hasSqliteEntryFullPath(storedVolume, caseRoot / L"unzipped-dir" / L"nested" / L"unzipped-nested.txt"),
                  L"Authoritative sqlite directory-move delta should persist hydrated unzipped children.");
    state.Require(! hasSqliteEntryFullPath(storedVolume, caseRoot / L"move-out-dir" / L"moved-out-child.txt"),
                  L"Authoritative sqlite directory-move delta should delete the moved-out subtree row.");

    LocalSearchIndexCore::QueryStats persistedStats{};
    std::vector<LocalSearchIndexCore::Candidate> persistedCandidates;
    hr = runPersistedIndexedNameQueryForTests(repository, caseRoot.wstring(), persistedStats, persistedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Authoritative sqlite directory-move persisted query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(
        requireJournalDirectoryMoveResults(state, caseRoot, replayStats, persistedCandidates, L"Authoritative sqlite directory-move synthetic replay"),
        L"Authoritative sqlite directory-move synthetic replay should persist the expected directory-move result set.");

    auto snapshotFiles = CollectDirectoryFilesByExtension(storeRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Authoritative sqlite directory-move flow should not create compatibility snapshot runtime files.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_sqlite_authoritative_reseeds_after_store_loss_during_replay",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_sqlite_authoritative_reseed_after_loss", caseRoot),
                  L"Failed to prepare index_core_sqlite_authoritative_reseed_after_loss root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path storeRoot  = caseRoot / L"store";
    const std::filesystem::path sqlitePath = caseRoot / L"authoritative.sqlite3";
    LocalSearchIndexCore::Repository repository({
        .snapshotRootDirectory = storeRoot.wstring(),
        .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
        .sqliteDatabasePath    = sqlitePath.wstring(),
        .sqliteAuthoritative   = true,
    });

    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", seedStats, seedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Authoritative sqlite reseed-after-loss seed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(seedCandidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Authoritative sqlite reseed-after-loss seed query returned an unexpected candidate set.");

    std::error_code ec;
    std::filesystem::remove(sqlitePath, ec);
    state.Require(! ec, std::format(L"Failed to delete authoritative SQLite database '{}'.", sqlitePath.wstring()));
    ec.clear();
    static_cast<void>(std::filesystem::remove(std::filesystem::path(sqlitePath.wstring() + L"-wal"), ec));
    state.Require(! ec, std::format(L"Failed to delete authoritative SQLite WAL '{}-wal'.", sqlitePath.wstring()));
    ec.clear();
    static_cast<void>(std::filesystem::remove(std::filesystem::path(sqlitePath.wstring() + L"-shm"), ec));
    state.Require(! ec, std::format(L"Failed to delete authoritative SQLite SHM '{}-shm'.", sqlitePath.wstring()));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.txt", "beta"), L"Failed to create beta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::this_thread::sleep_for(std::chrono::milliseconds(150));

    LocalSearchIndexCore::QueryStats replayStats{};
    std::vector<LocalSearchIndexCore::Candidate> replayCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", replayStats, replayCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Authoritative sqlite reseed-after-loss replay query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(replayCandidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"Authoritative sqlite reseed-after-loss replay query returned an unexpected candidate set.");
    state.Require(! replayStats.snapshotLoaded, L"Authoritative sqlite reseed-after-loss replay should not load a compatibility snapshot.");
    state.Require(! replayStats.snapshotSaved, L"Authoritative sqlite reseed-after-loss replay should not save a compatibility snapshot.");
    state.Require(replayStats.snapshotPath.empty(), L"Authoritative sqlite reseed-after-loss replay should not report a compatibility snapshot path.");
    state.Require(replayStats.snapshotFileBytes == 0u, L"Authoritative sqlite reseed-after-loss replay should report snapshotFileBytes=0.");
    if (! CanProveDirectSqliteCurrentness(replayStats))
    {
        return state.Skip(DirectSqliteCurrentnessSkipReason(replayStats));
    }

    state.Require(replayStats.journalReplayApplied, L"Authoritative sqlite reseed-after-loss replay should still use journal replay before the full reseed.");

    SqliteIndexStore::VolumeInfo volumeInfo{};
    hr = SqliteIndexStore::InspectVolume(sqlitePath.wstring(), caseRoot.wstring(), volumeInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::InspectVolume after authoritative SQLite store loss failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(volumeInfo.state == SqliteIndexStore::kVolumeStateReady,
                  std::format(L"Authoritative sqlite reseed-after-loss should leave the stored root READY, got {}.", volumeInfo.state));
    state.Require(volumeInfo.entryCount == 3u,
                  std::format(L"Authoritative sqlite reseed-after-loss should reseed 3 entries (root + alpha + beta), got {}.", volumeInfo.entryCount));

    hr = repository.DropCachedVolumeForTests(caseRoot.wstring());
    state.Require(SUCCEEDED(hr),
                  std::format(L"DropCachedVolumeForTests after authoritative SQLite store loss failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    LocalSearchIndexCore::QueryStats reloadStats{};
    std::vector<LocalSearchIndexCore::Candidate> reloadCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", reloadStats, reloadCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Authoritative sqlite reseed-after-loss reload query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(reloadCandidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"Authoritative sqlite reseed-after-loss reload query returned an unexpected candidate set.");
    state.Require(! reloadStats.snapshotLoaded, L"Authoritative sqlite reseed-after-loss reload should not load a compatibility snapshot.");
    state.Require(! reloadStats.snapshotSaved, L"Authoritative sqlite reseed-after-loss reload should not save a compatibility snapshot.");
    state.Require(reloadStats.snapshotPath.empty(), L"Authoritative sqlite reseed-after-loss reload should not report a compatibility snapshot path.");
    state.Require(reloadStats.snapshotFileBytes == 0u, L"Authoritative sqlite reseed-after-loss reload should report snapshotFileBytes=0.");

    auto snapshotFiles = CollectDirectoryFilesByExtension(storeRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Authoritative sqlite reseed-after-loss flow should not create a compatibility snapshot runtime file.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_sqlite_cold_start_stale_root_refreshes_before_query",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_sqlite_cold_start_stale_root", caseRoot),
                  L"Failed to prepare index_core_sqlite_cold_start_stale_root root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");

    const std::filesystem::path storeRoot  = caseRoot / L"store";
    const std::filesystem::path sqlitePath = caseRoot / L"index-v2.sqlite3";

    {
        LocalSearchIndexCore::Repository seedRepository({
            .snapshotRootDirectory = storeRoot.wstring(),
            .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
            .sqliteDatabasePath    = sqlitePath.wstring(),
            .sqliteAuthoritative   = true,
        });

        LocalSearchIndexCore::QueryStats seedStats{};
        std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
        HRESULT hr = RunIndexedNameQuery(seedRepository, caseRoot.wstring(), L"*.txt", seedStats, seedCandidates);
        state.Require(SUCCEEDED(hr), std::format(L"Seed stale-root SQLite query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        state.Require(! seedStats.snapshotSaved, L"Seed stale-root SQLite query should not create a compatibility snapshot runtime store.");
    }

    auto snapshotFiles = CollectDirectoryFilesByExtension(storeRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Seed stale-root SQLite query should not leave a compatibility snapshot runtime file behind.");

    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.txt", "beta"), L"Failed to create beta.txt after seed.");
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::Repository coldRepository({
        .snapshotRootDirectory = storeRoot.wstring(),
        .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
        .sqliteDatabasePath    = sqlitePath.wstring(),
        .sqliteAuthoritative   = true,
    });

    LocalSearchIndexCore::QueryStats coldStats{};
    std::vector<LocalSearchIndexCore::Candidate> coldCandidates;
    HRESULT hr = RunIndexedNameQuery(coldRepository, caseRoot.wstring(), L"*.txt", coldStats, coldCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Cold-start stale-root SQLite query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(coldCandidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"Cold-start stale-root SQLite query should rebuild/replay before answering from SQLite.");
    state.Require(coldStats.sqliteCutoverBlocked, L"Cold-start stale-root SQLite query should block direct SQLite before refreshing.");
    state.Require(! coldStats.snapshotSaved,
                  L"Cold-start stale-root SQLite query should refresh SQLite without recreating a compatibility snapshot runtime store.");
    state.Require(coldStats.snapshotPath.empty(), L"Cold-start stale-root SQLite query should not report a compatibility snapshot path.");
    snapshotFiles = CollectDirectoryFilesByExtension(storeRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Cold-start stale-root SQLite query should not create a compatibility snapshot runtime file.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_sqlite_direct_query_freshness_policy",
                  [&](SelfTest::CaseState& state) noexcept
{
    state.Require(! LocalSearchIndexCore::ShouldAllowDirectSqliteQueryForTests({}),
                  L"Direct SQLite query should be blocked when the stored volume is not ready.");
    state.Require(! LocalSearchIndexCore::ShouldAllowDirectSqliteQueryForTests({
                      .storedVolumeReady = true,
                  }),
                  L"Direct SQLite query should be blocked when no current journal state is available.");
    state.Require(! LocalSearchIndexCore::ShouldAllowDirectSqliteQueryForTests({
                      .storedVolumeReady       = true,
                      .currentJournalKnown     = true,
                      .currentJournalAvailable = false,
                  }),
                  L"Direct SQLite query should be blocked when the current journal state is unavailable.");
    state.Require(LocalSearchIndexCore::ShouldAllowDirectSqliteQueryForTests({
                      .storedVolumeReady       = true,
                      .currentJournalKnown     = true,
                      .currentJournalAvailable = true,
                      .storedJournalId         = 42u,
                      .storedNextUsn           = 900u,
                      .currentJournalId        = 42u,
                      .currentFirstUsn         = 800u,
                      .currentNextUsn          = 900u,
                  }),
                  L"Direct SQLite query should be allowed when the stored journal cursor matches the live journal state.");
    state.Require(! LocalSearchIndexCore::ShouldAllowDirectSqliteQueryForTests({
                      .storedVolumeReady       = true,
                      .currentJournalKnown     = true,
                      .currentJournalAvailable = true,
                      .storedJournalId         = 42u,
                      .storedNextUsn           = 900u,
                      .currentJournalId        = 43u,
                      .currentFirstUsn         = 800u,
                      .currentNextUsn          = 900u,
                  }),
                  L"Direct SQLite query should be blocked when the live journal id changed.");
    state.Require(! LocalSearchIndexCore::ShouldAllowDirectSqliteQueryForTests({
                      .storedVolumeReady       = true,
                      .currentJournalKnown     = true,
                      .currentJournalAvailable = true,
                      .storedJournalId         = 42u,
                      .storedNextUsn           = 899u,
                      .currentJournalId        = 42u,
                      .currentFirstUsn         = 800u,
                      .currentNextUsn          = 900u,
                  }),
                  L"Direct SQLite query should be blocked when the live journal advanced past the stored cursor.");
    state.Require(! LocalSearchIndexCore::ShouldAllowDirectSqliteQueryForTests({
                      .storedVolumeReady       = true,
                      .currentJournalKnown     = true,
                      .currentJournalAvailable = true,
                      .storedJournalId         = 42u,
                      .storedNextUsn           = 750u,
                      .currentJournalId        = 42u,
                      .currentFirstUsn         = 800u,
                      .currentNextUsn          = 900u,
                  }),
                  L"Direct SQLite query should be blocked when the stored cursor fell behind the current journal range.");
    state.Require(! LocalSearchIndexCore::ShouldAllowDirectSqliteQueryForTests({
                      .storedVolumeReady       = true,
                      .currentJournalKnown     = true,
                      .currentJournalAvailable = true,
                      .storedJournalId         = 42u,
                      .storedNextUsn           = 901u,
                      .currentJournalId        = 42u,
                      .currentFirstUsn         = 800u,
                      .currentNextUsn          = 900u,
                  }),
                  L"Direct SQLite query should be blocked when the stored cursor is ahead of the live journal tail.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_sqlite_cutover_blocked_by_pending_legacy_import",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_sqlite_cutover_blocked", caseRoot), L"Failed to prepare index_core_sqlite_cutover_blocked root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");

    const std::filesystem::path storeRoot  = caseRoot / L"store";
    const std::filesystem::path sqlitePath = caseRoot / L"index-v2.sqlite3";
    LocalSearchIndexCore::Repository repository({
        .snapshotRootDirectory = storeRoot.wstring(),
        .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
        .sqliteDatabasePath    = sqlitePath.wstring(),
    });

    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", seedStats, seedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Seed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    if (! CanProveDirectSqliteCurrentness(seedStats))
    {
        return state.Skip(DirectSqliteCurrentnessSkipReason(seedStats));
    }

    state.Require(seedStats.usedSqliteStore, L"Seed query should use SQLite once the mirrored store is ready.");

    SqliteIndexStore::ReplaceVolumeRequest pendingRequest{};
    pendingRequest.rootPath       = L"X:\\pending-legacy-root";
    pendingRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    pendingRequest.journalId      = 1u;
    pendingRequest.nextUsn        = 1u;
    pendingRequest.state          = SqliteIndexStore::kVolumeStateImportedLegacySnapshot;
    pendingRequest.entries.push_back({
        .fileIdLow      = 1u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 0u,
        .parentIdHigh   = 0u,
        .fullPath       = L"X:\\pending-legacy-root\\ghost.txt",
        .name           = L"ghost.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), pendingRequest, &replaceResult);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::ReplaceVolume for pending legacy volume failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    LocalSearchIndexCore::QueryStats cachedStats{};
    std::vector<LocalSearchIndexCore::Candidate> cachedCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", cachedStats, cachedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Cached-ready query after external pending insert failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(cachedStats.usedSqliteStore, L"Cached-ready query should continue using SQLite until the repository invalidates its cached store readiness.");
    state.Require(! cachedStats.sqliteCutoverBlocked, L"Cached-ready query should not report a blocked cutover before repository invalidation.");
    state.Require(CollectIndexedCandidateNames(cachedCandidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Cached-ready query returned an unexpected candidate set.");

    hr = repository.DropCachedVolumeForTests(caseRoot.wstring());
    state.Require(SUCCEEDED(hr), std::format(L"DropCachedVolumeForTests after pending insert failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    LocalSearchIndexCore::QueryStats blockedStats{};
    std::vector<LocalSearchIndexCore::Candidate> blockedCandidates;
    hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*.txt", blockedStats, blockedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Blocked-cutover query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(blockedStats.snapshotLoaded, L"Blocked-cutover query should fall back to the snapshot runtime path.");
    state.Require(! blockedStats.usedSqliteStore, L"Blocked-cutover query should not enumerate results from SQLite.");
    state.Require(blockedStats.sqliteCutoverBlocked, L"Blocked-cutover query should report the SQLite cutover gate as blocked.");
    state.Require(CollectIndexedCandidateNames(blockedCandidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Blocked-cutover query returned an unexpected candidate set.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_sqlite_sidecar_imports_legacy_snapshot",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_sqlite_sidecar_import", caseRoot), L"Failed to prepare index_core_sqlite_sidecar_import root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");

    const std::filesystem::path storeRoot  = caseRoot / L"store";
    const std::filesystem::path sqlitePath = storeRoot / L"index-v2.sqlite3";

    LocalSearchIndexCore::Repository snapshotRepository({
        .snapshotRootDirectory = storeRoot.wstring(),
        .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::SnapshotBinary,
    });

    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = RunIndexedNameQuery(snapshotRepository, caseRoot.wstring(), L"*.txt", seedStats, seedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Seed snapshot query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(! seedStats.snapshotPath.empty(), L"Seed snapshot query should create a snapshot path.");
    state.Require(seedStats.snapshotPath.ends_with(L".bin"), L"Seed snapshot query should persist a .bin snapshot.");

    LocalSearchIndexCore::Repository sqliteRepository({
        .snapshotRootDirectory = storeRoot.wstring(),
        .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
        .sqliteDatabasePath    = sqlitePath.wstring(),
    });

    LocalSearchIndexCore::QueryStats importStats{};
    std::vector<LocalSearchIndexCore::Candidate> importCandidates;
    hr = RunIndexedNameQuery(sqliteRepository, caseRoot.wstring(), L"*.txt", importStats, importCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"SQLite-configured repository import query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(importCandidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"SQLite-configured repository import query returned an unexpected candidate set.");

    SqliteIndexStore::StoreInfo storeInfo{};
    hr = SqliteIndexStore::InspectStore(sqlitePath.wstring(), storeInfo);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::InspectStore after sidecar import failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(storeInfo.volumeCount == 1u, std::format(L"Expected one imported volume, got {}.", storeInfo.volumeCount));
    state.Require(storeInfo.legacyImportVolumeCount == 0u,
                  std::format(L"Expected zero legacy-import volumes after metadata backfill, got {}.", storeInfo.legacyImportVolumeCount));
    state.Require(storeInfo.entryCount == importStats.entryCount + 1u,
                  std::format(L"Expected {} imported entries, got {}.", importStats.entryCount + 1u, storeInfo.entryCount));

    std::wstring sqliteError;
    uint64_t mirroredState = 0u;
    state.Require(QuerySqliteSingleInt64(sqlitePath, "SELECT state FROM volumes LIMIT 1;", mirroredState, sqliteError), sqliteError);
    const uint32_t expectedMirroredState =
        CanProveDirectSqliteCurrentness(importStats) ? SqliteIndexStore::kVolumeStateReady : SqliteIndexStore::kVolumeStateCurrentnessUnproven;
    state.Require(mirroredState == expectedMirroredState, std::format(L"Expected mirrored volume state {}, got {}.", expectedMirroredState, mirroredState));

    uint64_t alphaSizeBytes = 0u;
    state.Require(QuerySqliteSingleInt64(sqlitePath, "SELECT size_bytes FROM entries WHERE name = 'alpha.txt' LIMIT 1;", alphaSizeBytes, sqliteError),
                  sqliteError);
    state.Require(alphaSizeBytes == 5u, std::format(L"Expected alpha.txt size_bytes=5, got {}.", alphaSizeBytes));

    uint64_t alphaWriteTime = 0u;
    state.Require(QuerySqliteSingleInt64(sqlitePath, "SELECT write_time_100ns FROM entries WHERE name = 'alpha.txt' LIMIT 1;", alphaWriteTime, sqliteError),
                  sqliteError);
    state.Require(alphaWriteTime != 0u, L"Expected alpha.txt write_time_100ns to be backfilled.");

    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.txt", "beta"), L"Failed to create beta.txt.");

    LocalSearchIndexCore::QueryStats updatedStats{};
    std::vector<LocalSearchIndexCore::Candidate> updatedCandidates;
    hr = RunIndexedNameQuery(sqliteRepository, caseRoot.wstring(), L"*.txt", updatedStats, updatedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"SQLite-configured repository update query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(updatedCandidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"SQLite-configured repository update query returned an unexpected candidate set.");

    hr = SqliteIndexStore::InspectStore(sqlitePath.wstring(), storeInfo);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::InspectStore after sidecar update failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(storeInfo.volumeCount == 1u, std::format(L"Expected one mirrored volume after update, got {}.", storeInfo.volumeCount));
    state.Require(storeInfo.legacyImportVolumeCount == 0u,
                  std::format(L"Expected zero legacy-import volumes after update, got {}.", storeInfo.legacyImportVolumeCount));
    state.Require(storeInfo.entryCount == updatedStats.entryCount + 1u,
                  std::format(L"Expected {} mirrored entries after update, got {}.", updatedStats.entryCount + 1u, storeInfo.entryCount));

    uint64_t betaSizeBytes = 0u;
    state.Require(QuerySqliteSingleInt64(sqlitePath, "SELECT size_bytes FROM entries WHERE name = 'beta.txt' LIMIT 1;", betaSizeBytes, sqliteError),
                  sqliteError);
    state.Require(betaSizeBytes == 4u, std::format(L"Expected beta.txt size_bytes=4, got {}.", betaSizeBytes));

    uint64_t betaWriteTime = 0u;
    state.Require(QuerySqliteSingleInt64(sqlitePath, "SELECT write_time_100ns FROM entries WHERE name = 'beta.txt' LIMIT 1;", betaWriteTime, sqliteError),
                  sqliteError);
    state.Require(betaWriteTime != 0u, L"Expected beta.txt write_time_100ns to be backfilled.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_index_core_sqlite_prefilter_classifies_name_patterns",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"index_core_sqlite_prefilter", caseRoot), L"Failed to prepare index_core_sqlite_prefilter root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alphabet.txt", "alphabet"), L"Failed to create alphabet.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.log", "beta"), L"Failed to create beta.log.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"gamma.bin", "gamma"), L"Failed to create gamma.bin.");

    const std::filesystem::path storeRoot  = caseRoot / L"store";
    const std::filesystem::path sqlitePath = caseRoot / L"index-v2.sqlite3";
    LocalSearchIndexCore::Repository repository({
        .snapshotRootDirectory = storeRoot.wstring(),
        .persistentStoreKind   = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
        .sqliteDatabasePath    = sqlitePath.wstring(),
    });

    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = RunIndexedNameQuery(repository, caseRoot.wstring(), L"*", seedStats, seedCandidates);
    state.Require(SUCCEEDED(hr), std::format(L"Seed SQLite prefilter query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    if (! CanProveDirectSqliteCurrentness(seedStats))
    {
        return state.Skip(DirectSqliteCurrentnessSkipReason(seedStats));
    }

    struct PrefilterExpectation final
    {
        std::wstring pattern;
        FileSystemSearchNameMode mode = FILESYSTEM_SEARCH_NAME_WILDCARD;
        std::vector<std::wstring> expectedNames;
        bool expectPrefilter = false;
    };

    const std::array expectations{
        PrefilterExpectation{
            .pattern         = L"alpha.txt",
            .mode            = FILESYSTEM_SEARCH_NAME_LITERAL,
            .expectedNames   = {L"alpha.txt"},
            .expectPrefilter = true,
        },
        PrefilterExpectation{
            .pattern         = L"alphabet.txt",
            .mode            = FILESYSTEM_SEARCH_NAME_WILDCARD,
            .expectedNames   = {L"alphabet.txt"},
            .expectPrefilter = true,
        },
        PrefilterExpectation{
            .pattern         = L"alph*",
            .mode            = FILESYSTEM_SEARCH_NAME_WILDCARD,
            .expectedNames   = {L"alpha.txt", L"alphabet.txt"},
            .expectPrefilter = true,
        },
        PrefilterExpectation{
            .pattern         = L"*.log",
            .mode            = FILESYSTEM_SEARCH_NAME_WILDCARD,
            .expectedNames   = {L"beta.log"},
            .expectPrefilter = true,
        },
        PrefilterExpectation{
            .pattern         = L"a*ha.txt",
            .mode            = FILESYSTEM_SEARCH_NAME_WILDCARD,
            .expectedNames   = {L"alpha.txt"},
            .expectPrefilter = false,
        },
        PrefilterExpectation{
            .pattern         = L"^alpha.*\\.txt$",
            .mode            = FILESYSTEM_SEARCH_NAME_REGEX,
            .expectedNames   = {L"alpha.txt", L"alphabet.txt"},
            .expectPrefilter = false,
        },
    };

    for (const auto& expectation : expectations)
    {
        LocalSearchIndexCore::QueryStats queryStats{};
        std::vector<LocalSearchIndexCore::Candidate> candidates;
        hr = RunIndexedNameQuery(repository, caseRoot.wstring(), expectation.pattern, expectation.mode, queryStats, candidates);
        state.Require(SUCCEEDED(hr), std::format(L"SQLite prefilter query '{}' failed. hr=0x{:08X}", expectation.pattern, static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        state.Require(queryStats.usedSqliteStore, std::format(L"Warm SQLite prefilter query '{}' should enumerate from SQLite.", expectation.pattern));
        state.Require(queryStats.sqliteReadOnlyQuery,
                      std::format(L"Warm SQLite prefilter query '{}' should use a read-only SQLite connection.", expectation.pattern));
        state.Require(queryStats.usedNamePrefilter == expectation.expectPrefilter,
                      std::format(L"Warm SQLite prefilter query '{}' prefilter expectation mismatch. expected={} got={}",
                                  expectation.pattern,
                                  expectation.expectPrefilter,
                                  queryStats.usedNamePrefilter));
        state.Require(CollectIndexedCandidateNames(candidates) == expectation.expectedNames,
                      std::format(L"Warm SQLite prefilter query '{}' returned an unexpected candidate set.", expectation.pattern));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sqlite_index_store_read_queries_stay_readonly_under_writer_lock",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_readonly_under_writer_lock", caseRoot), L"Failed to prepare sqlite_readonly_under_writer_lock root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path sqlitePath = caseRoot / L"index-v2.sqlite3";
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::EnsureBootstrap failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest replaceRequest{};
    replaceRequest.rootPath       = caseRoot.wstring();
    replaceRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replaceRequest.journalId      = 51u;
    replaceRequest.nextUsn        = 510u;
    replaceRequest.state          = SqliteIndexStore::kVolumeStateReady;
    replaceRequest.entries.push_back({
        .fileIdLow  = 1u,
        .fileIdHigh = 0u,
        .fullPath   = caseRoot.wstring(),
        .name       = L"",
        .attributes = FILE_ATTRIBUTE_DIRECTORY,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 2u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 1u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"alpha.txt").wstring(),
        .name           = L"alpha.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 5u,
        .writeTime100ns = 133838640000000000ull,
    });

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), replaceRequest, &replaceResult);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::ReplaceVolume failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    sqlite3* rawWriter             = nullptr;
    const std::u8string sqliteUtf8 = sqlitePath.u8string();
    const int openResult = sqlite3_open_v2(reinterpret_cast<const char*>(sqliteUtf8.c_str()), &rawWriter, SQLITE_OPEN_READWRITE | SQLITE_OPEN_NOMUTEX, nullptr);
    state.Require(openResult == SQLITE_OK && rawWriter != nullptr,
                  std::format(L"sqlite3_open_v2 writer failed. code={} message='{}'",
                              openResult,
                              (rawWriter != nullptr) ? static_cast<const wchar_t*>(sqlite3_errmsg16(rawWriter)) : L"<unknown>"));
    if (openResult != SQLITE_OK || rawWriter == nullptr)
    {
        if (rawWriter != nullptr)
        {
            static_cast<void>(sqlite3_close(rawWriter));
        }
        return false;
    }

    const auto closeWriter = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(sqlite3_exec(rawWriter, "ROLLBACK;", nullptr, nullptr, nullptr));
        static_cast<void>(sqlite3_close(rawWriter));
    });

    const int beginResult = sqlite3_exec(rawWriter, "BEGIN IMMEDIATE;", nullptr, nullptr, nullptr);
    state.Require(
        beginResult == SQLITE_OK,
        std::format(L"BEGIN IMMEDIATE writer lock failed. code={} message='{}'", beginResult, static_cast<const wchar_t*>(sqlite3_errmsg16(rawWriter))));
    if (beginResult != SQLITE_OK)
    {
        return false;
    }

    SqliteIndexStore::VolumeInfo volumeInfo{};
    hr = SqliteIndexStore::InspectVolume(sqlitePath.wstring(), caseRoot.wstring(), volumeInfo);
    state.Require(SUCCEEDED(hr), std::format(L"InspectVolume must stay read-only under a held writer lock. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(volumeInfo.entryCount == 2u, std::format(L"InspectVolume should report 2 entries, got {}.", volumeInfo.entryCount));

    SqliteIndexStore::ReplaceVolumeRequest loadedVolume{};
    hr = SqliteIndexStore::LoadVolume(sqlitePath.wstring(), caseRoot.wstring(), loadedVolume);
    state.Require(SUCCEEDED(hr), std::format(L"LoadVolume must stay read-only under a held writer lock. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(loadedVolume.entries.size() == 2u, std::format(L"LoadVolume should hydrate 2 entries, got {}.", loadedVolume.entries.size()));

    SqliteIndexStore::QueryRequest queryRequest{};
    queryRequest.rootPath           = caseRoot.wstring();
    queryRequest.namePattern        = L"alpha*";
    queryRequest.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    queryRequest.recursive          = true;
    queryRequest.includeFiles       = true;
    queryRequest.includeDirectories = false;

    std::vector<std::wstring> names;
    SqliteIndexStore::QueryRuntimeStats queryStats{};
    hr = CollectSqliteEnumerateNames(sqlitePath, queryRequest, names, &queryStats);
    state.Require(SUCCEEDED(hr), std::format(L"EnumerateVolume must stay read-only under a held writer lock. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(queryStats.readOnlyConnection, L"EnumerateVolume should report a read-only SQLite connection.");
    state.Require(names == std::vector<std::wstring>{L"alpha.txt"}, L"EnumerateVolume under writer lock returned an unexpected result set.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sqlite_index_store_prefix_prefilter_uses_range_and_literal_metacharacters",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_prefix_range_literals", caseRoot), L"Failed to prepare sqlite_prefix_range_literals root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path sqlitePath = caseRoot / L"index-v2.sqlite3";
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::EnsureBootstrap failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest replaceRequest{};
    replaceRequest.rootPath       = caseRoot.wstring();
    replaceRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replaceRequest.journalId      = 61u;
    replaceRequest.nextUsn        = 610u;
    replaceRequest.state          = SqliteIndexStore::kVolumeStateReady;
    replaceRequest.entries.push_back({
        .fileIdLow  = 1u,
        .fileIdHigh = 0u,
        .fullPath   = caseRoot.wstring(),
        .name       = L"",
        .attributes = FILE_ATTRIBUTE_DIRECTORY,
    });

    const std::array<std::wstring_view, 5> fileNames{{L"literal%one.txt", L"literal_one.txt", L"literalAone.txt", L"literalxone.txt", L"other.txt"}};
    for (size_t index = 0u; index < fileNames.size(); ++index)
    {
        replaceRequest.entries.push_back({
            .fileIdLow      = static_cast<uint64_t>(index + 2u),
            .fileIdHigh     = 0u,
            .parentIdLow    = 1u,
            .parentIdHigh   = 0u,
            .fullPath       = (caseRoot / fileNames[index]).wstring(),
            .name           = std::wstring(fileNames[index]),
            .attributes     = FILE_ATTRIBUTE_ARCHIVE,
            .sizeBytes      = 5u,
            .writeTime100ns = 133838640000000000ull + static_cast<uint64_t>(index),
        });
    }

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), replaceRequest, &replaceResult);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::ReplaceVolume failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::QueryRequest percentRequest{};
    percentRequest.rootPath           = caseRoot.wstring();
    percentRequest.namePattern        = L"literal%*";
    percentRequest.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    percentRequest.recursive          = true;
    percentRequest.includeFiles       = true;
    percentRequest.includeDirectories = false;

    SqliteIndexStore::QueryPlanInspectionForTests plan{};
    hr = SqliteIndexStore::ExplainEnumerateVolumeForTests(sqlitePath.wstring(), percentRequest, plan);
    state.Require(SUCCEEDED(hr), std::format(L"ExplainEnumerateVolumeForTests failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(
        plan.prefixLowerBound == L"literal%" && plan.prefixUpperBound == L"literal&",
        std::format(
            L"Percent-literal prefix bounds mismatch. lower='{}' upper='{}' expanded='{}'", plan.prefixLowerBound, plan.prefixUpperBound, plan.expandedSql));

    std::vector<std::wstring> names;
    SqliteIndexStore::QueryRuntimeStats queryStats{};
    const auto percentQueryStarted = std::chrono::steady_clock::now();
    hr                             = CollectSqliteEnumerateNames(sqlitePath, percentRequest, names, &queryStats);
    const uint64_t percentQueryElapsedUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - percentQueryStarted).count());
    Debug::Perf::Emit(L"compare.selftest.sqlite_prefix_range_query_us",
                      L"literal-percent",
                      percentQueryElapsedUs,
                      queryStats.scannedRows,
                      static_cast<uint64_t>(names.size()),
                      hr);
    state.Require(SUCCEEDED(hr), std::format(L"Percent-literal prefix enumerate failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(queryStats.usedNamePrefilter, L"Percent-literal prefix query should use a SQLite prefilter.");
    state.Require(
        names == std::vector<std::wstring>{L"literal%one.txt"},
        std::format(L"Percent-literal prefix should only return the literal percent name, got {} candidate(s). lower='{}' upper='{}' expanded='{}' plan='{}'",
                    names.size(),
                    plan.prefixLowerBound,
                    plan.prefixUpperBound,
                    plan.expandedSql,
                    plan.detail));

    SqliteIndexStore::QueryRequest underscoreRequest = percentRequest;
    underscoreRequest.namePattern                    = L"literal_*";
    names.clear();
    queryStats = {};
    hr         = CollectSqliteEnumerateNames(sqlitePath, underscoreRequest, names, &queryStats);
    state.Require(SUCCEEDED(hr), std::format(L"Underscore-literal prefix enumerate failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(queryStats.usedNamePrefilter, L"Underscore-literal prefix query should use a SQLite prefilter.");
    state.Require(names == std::vector<std::wstring>{L"literal_one.txt"},
                  std::format(L"Underscore-literal prefix should only return the literal underscore name, got {} candidate(s).", names.size()));

    state.Require(plan.sql.find("LIKE") == std::string::npos, L"Prefix prefilter SQL must not use LIKE.");
    state.Require(plan.sql.find("name_folded >= ?") != std::string::npos && plan.sql.find("name_folded < ?") != std::string::npos,
                  L"Prefix prefilter SQL must use a half-open name_folded range.");
    state.Require(plan.detail.find(L"idx_entries_name_folded") != std::wstring::npos,
                  std::format(L"Prefix prefilter should use idx_entries_name_folded. plan='{}'", plan.detail));
    state.Require(plan.detail.find(L"name_folded>?") != std::wstring::npos && plan.detail.find(L"name_folded<?") != std::wstring::npos,
                  std::format(L"Prefix prefilter plan should constrain name_folded on both bounds. plan='{}'", plan.detail));
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_native_matches_host_fallback",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
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

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"auto\"}");
    state.Require(SUCCEEDED(setHr), std::format(L"Failed to reset isolated local file system configuration. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! search)
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring unavailablePipe      = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, unavailablePipe.c_str()) != 0,
                  L"Failed to override the search service pipe for local-index parity.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_native_matches_host_fallback", caseRoot),
                  L"Failed to prepare search_native_matches_host_fallback root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"sub"), L"Failed to create sub directory.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"report-one.txt", "alpha needle omega"), L"Failed to create report-one.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"sub" / L"report-two.txt", "prefix needle suffix"), L"Failed to create sub\\report-two.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"other.txt", "needle"), L"Failed to create other.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"sub" / L"report-three.log", "needle"), L"Failed to create sub\\report-three.log.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"sub" / L"report-empty.txt", "no match here"), L"Failed to create sub\\report-empty.txt.");

    std::wstring rootText       = caseRoot.wstring();
    std::wstring namePattern    = L"report*.txt";
    std::wstring contentPattern = L"needle";

    FileSystemSearchQuery query{};
    query.sizeBytes      = sizeof(FileSystemSearchQuery);
    query.rootPath       = rootText.c_str();
    query.namePattern    = namePattern.c_str();
    query.contentPattern = contentPattern.c_str();
    query.flags          = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES | FILESYSTEM_SEARCH_WANT_SNIPPETS |
                                                              FILESYSTEM_SEARCH_PREFER_INDEX);
    query.nameMode       = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode    = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;
    query.maxSnippetCharacters = 48u;

    RecordingSearchCallback nativeCallback;
    const HRESULT nativeHr = search->Search(&query, &nativeCallback, nullptr);
    state.Require(SUCCEEDED(nativeHr), std::format(L"Native plugin search failed. hr=0x{:08X}", static_cast<unsigned long>(nativeHr)));
    if (FAILED(nativeHr))
    {
        return false;
    }

    RecordingSearchCallback fallbackCallback;
    const HRESULT fallbackHr = SearchFallbackEngine::Execute(created.fileSystem.get(), &query, &fallbackCallback, nullptr);
    state.Require(SUCCEEDED(fallbackHr), std::format(L"Host fallback search failed. hr=0x{:08X}", static_cast<unsigned long>(fallbackHr)));
    if (FAILED(fallbackHr))
    {
        return false;
    }

    auto nativeMatches      = nativeCallback.Matches();
    auto fallbackMatches    = fallbackCallback.Matches();
    const auto compareMatch = [](const RecordedSearchMatch& left, const RecordedSearchMatch& right) noexcept
    {
        return std::tie(left.fullPath, left.relativePath, left.displayName, left.previewText, left.matchedBy, left.attributes) <
               std::tie(right.fullPath, right.relativePath, right.displayName, right.previewText, right.matchedBy, right.attributes);
    };
    std::sort(nativeMatches.begin(), nativeMatches.end(), compareMatch);
    std::sort(fallbackMatches.begin(), fallbackMatches.end(), compareMatch);

    state.Require(nativeMatches.size() == fallbackMatches.size(),
                  std::format(L"Native and fallback match counts differ. native={}, fallback={}.", nativeMatches.size(), fallbackMatches.size()));
    for (size_t index = 0; index < nativeMatches.size() && index < fallbackMatches.size(); ++index)
    {
        const RecordedSearchMatch& nativeMatch   = nativeMatches[index];
        const RecordedSearchMatch& fallbackMatch = fallbackMatches[index];
        state.Require(nativeMatch.fullPath == fallbackMatch.fullPath, std::format(L"Match {} fullPath mismatch.", index));
        state.Require(nativeMatch.relativePath == fallbackMatch.relativePath, std::format(L"Match {} relativePath mismatch.", index));
        state.Require(nativeMatch.displayName == fallbackMatch.displayName, std::format(L"Match {} displayName mismatch.", index));
        state.Require(nativeMatch.previewText == fallbackMatch.previewText, std::format(L"Match {} previewText mismatch.", index));
        state.Require(nativeMatch.matchedBy == fallbackMatch.matchedBy, std::format(L"Match {} matchedBy mismatch.", index));
        state.Require(nativeMatch.attributes == fallbackMatch.attributes, std::format(L"Match {} attributes mismatch.", index));
    }

    const auto nativeProgress                       = nativeCallback.ProgressSnapshots();
    const auto fallbackProgress                     = fallbackCallback.ProgressSnapshots();
    const RecordedSearchProgress* nativeCompleted   = FindRecordedSearchProgress(nativeProgress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    const RecordedSearchProgress* fallbackCompleted = FindRecordedSearchProgress(fallbackProgress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(nativeCompleted != nullptr, L"Native plugin search missing completed progress.");
    state.Require(fallbackCompleted != nullptr, L"Host fallback search missing completed progress.");
    if (nativeCompleted != nullptr && fallbackCompleted != nullptr)
    {
        state.Require(nativeCompleted->backend == FILESYSTEM_SEARCH_BACKEND_INDEX, L"Native plugin search should report index backend.");
        state.Require(fallbackCompleted->backend == FILESYSTEM_SEARCH_BACKEND_SCAN, L"Host fallback search should report scan backend.");
        state.Require((nativeCompleted->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) == 0u,
                      L"Native indexed search should not report degraded-no-index on supported local roots.");
        state.Require((fallbackCompleted->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u,
                      L"Host fallback search should still report degraded-no-index when PREFER_INDEX is requested.");
        state.Require(nativeCompleted->matchedEntries == fallbackCompleted->matchedEntries, L"Native and fallback matched entry counts should match.");
        state.Require(nativeCompleted->candidateFiles == fallbackCompleted->candidateFiles, L"Native and fallback candidate file counts should match.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_native_unicode_long_path_matches_host_fallback",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
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

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"local-index\"}");
    state.Require(SUCCEEDED(setHr),
                  std::format(L"Failed to configure isolated local file system for local-index parity. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! search)
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring unavailablePipe      = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, unavailablePipe.c_str()) != 0,
                  L"Failed to override the search service pipe for Unicode/long-path parity.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_native_unicode_long_path_matches_host_fallback", caseRoot),
                  L"Failed to prepare search_native_unicode_long_path_matches_host_fallback root.");

    const std::filesystem::path unicodeFile = caseRoot / L"rapport-\u00E9quipe.txt";
    state.Require(SelfTest::WriteTextFile(unicodeFile, "alpha needle unicode"), L"Failed to create rapport-\\u00E9quipe.txt.");
    const std::wstring unicodeFileName = unicodeFile.filename().wstring();
    state.Require(unicodeFileName.find(L"\u00E9") != std::wstring::npos && unicodeFileName.find(L"\u00C3") == std::wstring::npos,
                  L"Unicode/long-path fixture must contain a real U+00E9 filename, not mojibake.");

    std::filesystem::path deepDir = caseRoot / L"chemin-long";
    state.Require(SelfTest::EnsureDirectory(deepDir), L"Failed to create chemin-long.");

    bool deepPathCreated                   = true;
    const std::wstring_view longSegments[] = {
        L"segment-00-abcdefghijklmnopqrstuvwxyz",
        L"segment-01-abcdefghijklmnopqrstuvwxyz",
        L"segment-02-abcdefghijklmnopqrstuvwxyz",
        L"segment-03-abcdefghijklmnopqrstuvwxyz",
        L"segment-04-abcdefghijklmnopqrstuvwxyz",
        L"segment-05-abcdefghijklmnopqrstuvwxyz",
    };
    for (const std::wstring_view segment : longSegments)
    {
        deepDir /= std::filesystem::path(segment);
        if (! EnsureExtendedDirectoryForCompareSelfTest(deepDir))
        {
            deepPathCreated = false;
            break;
        }
    }

    bool longPathCreated = deepPathCreated && deepDir.wstring().size() > 260u;
    if (! deepPathCreated)
    {
        deepDir                                    = caseRoot / L"chemin-profond";
        deepPathCreated                            = EnsureExtendedDirectoryForCompareSelfTest(deepDir);
        const std::wstring_view fallbackSegments[] = {
            L"segment-a-abcdefghijklmnop",
            L"segment-b-abcdefghijklmnop",
            L"segment-c-abcdefghijklmnop",
        };
        for (const std::wstring_view segment : fallbackSegments)
        {
            if (! deepPathCreated)
            {
                break;
            }

            deepDir /= std::filesystem::path(segment);
            deepPathCreated = EnsureExtendedDirectoryForCompareSelfTest(deepDir);
        }
        longPathCreated = false;
    }

    state.Require(deepPathCreated, L"Failed to create a deep Unicode search path.");
    if (! deepPathCreated)
    {
        return false;
    }

    const std::filesystem::path deepUnicodeFile = deepDir / L"r\u00E9sultat-final.txt";
    state.Require(WriteTextFileExtendedForCompareSelfTest(deepUnicodeFile, "prefix needle suffix"), L"Failed to create r\\u00E9sultat-final.txt.");
    const std::wstring deepUnicodeFileName = deepUnicodeFile.filename().wstring();
    state.Require(deepUnicodeFileName.find(L"\u00E9") != std::wstring::npos && deepUnicodeFileName.find(L"\u00C3") == std::wstring::npos,
                  L"Unicode/long-path deep fixture must contain a real U+00E9 filename, not mojibake.");

    std::wstring rootText       = caseRoot.wstring();
    std::wstring namePattern    = L"*.txt";
    std::wstring contentPattern = L"needle";

    FileSystemSearchQuery query{};
    query.sizeBytes      = sizeof(FileSystemSearchQuery);
    query.rootPath       = rootText.c_str();
    query.namePattern    = namePattern.c_str();
    query.contentPattern = contentPattern.c_str();
    query.flags          = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES | FILESYSTEM_SEARCH_WANT_SNIPPETS |
                                                              FILESYSTEM_SEARCH_PREFER_INDEX);
    query.nameMode       = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode    = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;
    query.maxSnippetCharacters = 48u;

    RecordingSearchCallback nativeCallback;
    const HRESULT nativeHr = search->Search(&query, &nativeCallback, nullptr);
    state.Require(SUCCEEDED(nativeHr), std::format(L"Native Unicode/long-path search failed. hr=0x{:08X}", static_cast<unsigned long>(nativeHr)));
    if (FAILED(nativeHr))
    {
        return false;
    }

    RecordingSearchCallback fallbackCallback;
    const HRESULT fallbackHr = SearchFallbackEngine::Execute(created.fileSystem.get(), &query, &fallbackCallback, nullptr);
    state.Require(SUCCEEDED(fallbackHr), std::format(L"Host fallback Unicode/long-path search failed. hr=0x{:08X}", static_cast<unsigned long>(fallbackHr)));
    if (FAILED(fallbackHr))
    {
        return false;
    }

    auto nativeMatches      = nativeCallback.Matches();
    auto fallbackMatches    = fallbackCallback.Matches();
    const auto compareMatch = [](const RecordedSearchMatch& left, const RecordedSearchMatch& right) noexcept
    {
        return std::tie(left.fullPath, left.relativePath, left.displayName, left.previewText, left.matchedBy, left.attributes) <
               std::tie(right.fullPath, right.relativePath, right.displayName, right.previewText, right.matchedBy, right.attributes);
    };
    std::sort(nativeMatches.begin(), nativeMatches.end(), compareMatch);
    std::sort(fallbackMatches.begin(), fallbackMatches.end(), compareMatch);

    state.Require(
        nativeMatches.size() == fallbackMatches.size(),
        std::format(L"Unicode/long-path native and fallback match counts differ. native={}, fallback={}.", nativeMatches.size(), fallbackMatches.size()));
    for (size_t index = 0; index < nativeMatches.size() && index < fallbackMatches.size(); ++index)
    {
        const RecordedSearchMatch& nativeMatch   = nativeMatches[index];
        const RecordedSearchMatch& fallbackMatch = fallbackMatches[index];
        state.Require(nativeMatch.fullPath == fallbackMatch.fullPath, std::format(L"Unicode/long-path match {} fullPath mismatch.", index));
        state.Require(nativeMatch.relativePath == fallbackMatch.relativePath, std::format(L"Unicode/long-path match {} relativePath mismatch.", index));
        state.Require(nativeMatch.displayName == fallbackMatch.displayName, std::format(L"Unicode/long-path match {} displayName mismatch.", index));
        state.Require(nativeMatch.previewText == fallbackMatch.previewText, std::format(L"Unicode/long-path match {} previewText mismatch.", index));
        state.Require(nativeMatch.matchedBy == fallbackMatch.matchedBy, std::format(L"Unicode/long-path match {} matchedBy mismatch.", index));
        state.Require(nativeMatch.attributes == fallbackMatch.attributes, std::format(L"Unicode/long-path match {} attributes mismatch.", index));
    }

    const RecordedSearchMatch* unicodeMatch = FindRecordedSearchMatch(nativeMatches, L"rapport-\u00E9quipe.txt");
    state.Require(unicodeMatch != nullptr, L"Native Unicode/long-path search missing rapport-\\u00E9quipe.txt.");

    const RecordedSearchMatch* deepUnicodeMatch = FindRecordedSearchMatch(nativeMatches, L"r\u00E9sultat-final.txt");
    state.Require(deepUnicodeMatch != nullptr, L"Native Unicode/long-path search missing r\\u00E9sultat-final.txt.");
    if (deepUnicodeMatch != nullptr)
    {
        state.Require(deepUnicodeMatch->relativePath.find(L"r\u00E9sultat-final.txt") != std::wstring::npos,
                      L"Unicode/long-path deep match should preserve the nested relative path.");
        if (longPathCreated)
        {
            state.Require(deepUnicodeMatch->fullPath.size() > 260u,
                          std::format(L"Expected a long-path match longer than MAX_PATH, got {} characters.", deepUnicodeMatch->fullPath.size()));
        }
    }

    const auto nativeProgress                       = nativeCallback.ProgressSnapshots();
    const auto fallbackProgress                     = fallbackCallback.ProgressSnapshots();
    const RecordedSearchProgress* nativeCompleted   = FindRecordedSearchProgress(nativeProgress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    const RecordedSearchProgress* fallbackCompleted = FindRecordedSearchProgress(fallbackProgress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(nativeCompleted != nullptr, L"Native Unicode/long-path search missing completed progress.");
    state.Require(fallbackCompleted != nullptr, L"Host fallback Unicode/long-path search missing completed progress.");
    if (nativeCompleted != nullptr && fallbackCompleted != nullptr)
    {
        state.Require(nativeCompleted->backend == FILESYSTEM_SEARCH_BACKEND_INDEX, L"Native Unicode/long-path search should report index backend.");
        state.Require(fallbackCompleted->backend == FILESYSTEM_SEARCH_BACKEND_SCAN, L"Host fallback Unicode/long-path search should report scan backend.");
        state.Require((nativeCompleted->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) == 0u,
                      L"Native Unicode/long-path search should not degrade on supported local roots.");
        state.Require((fallbackCompleted->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u,
                      L"Host fallback Unicode/long-path search should report degraded-no-index when PREFER_INDEX is requested.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"test_support_child_process_runner_contract",
                  [&](SelfTest::CaseState& state) noexcept
{
    namespace TestSupport = RedSalamander::TestSupport;

    const std::filesystem::path servicePath = GetSiblingExecutablePath(L"RedSalamanderSearchService.exe");
    std::error_code existsEc;
    state.Require(! servicePath.empty() && std::filesystem::exists(servicePath, existsEc),
                  std::format(L"Service executable not found: {}", servicePath.wstring()));
    if (! state.failure.empty())
    {
        return false;
    }

    const TestSupport::ChildProcessResult largeStreams = TestSupport::RunChildProcess(
        {.executablePath = servicePath,
         .arguments      = {L"--test-support-child-probe=streams"},
         .timeout        = std::chrono::milliseconds(SelfTest::ScaleTimeout(10'000))});
    state.Require(largeStreams.Completed(), largeStreams.diagnostic);
    state.Require(largeStreams.exitCode == 0u, std::format(L"Large-stream probe exited with code {}.", largeStreams.exitCode));
    state.Require(largeStreams.stdoutBytes.size() == 2u * 1024u * 1024u &&
                      std::all_of(largeStreams.stdoutBytes.begin(), largeStreams.stdoutBytes.end(), [](char value) noexcept { return value == 'O'; }),
                  L"Concurrent stdout drain captures the full output beyond pipe capacity.");
    state.Require(largeStreams.stderrBytes.size() == 2u * 1024u * 1024u &&
                      std::all_of(largeStreams.stderrBytes.begin(), largeStreams.stderrBytes.end(), [](char value) noexcept { return value == 'E'; }),
                  L"Concurrent stderr drain captures the full output beyond pipe capacity.");

    const TestSupport::ChildProcessResult boundedStreams = TestSupport::RunChildProcess(
        {.executablePath = servicePath,
         .arguments      = {L"--test-support-child-probe=streams"},
         .timeout        = std::chrono::milliseconds(SelfTest::ScaleTimeout(10'000)),
         .maxStdoutBytes = 64u * 1024u,
         .maxStderrBytes = 96u * 1024u});
    state.Require(boundedStreams.Completed(), boundedStreams.diagnostic);
    state.Require(boundedStreams.stdoutBytes.size() == 64u * 1024u && boundedStreams.stdoutTruncated,
                  L"Stdout capture stays bounded while continuing to drain the child pipe.");
    state.Require(boundedStreams.stderrBytes.size() == 96u * 1024u && boundedStreams.stderrTruncated,
                  L"Stderr capture stays bounded while continuing to drain the child pipe.");

    const TestSupport::ChildProcessResult quotedArguments = TestSupport::RunChildProcess(
        {.executablePath = servicePath,
         .arguments = {L"--test-support-child-probe=arguments", L"space value", L"quote\"value", L"backslash\\tail\\", L"Unicode-\u6E2C\u8A66-\U0001F642"},
         .timeout   = std::chrono::milliseconds(SelfTest::ScaleTimeout(5'000))});
    state.Require(quotedArguments.Completed(), quotedArguments.diagnostic);
    state.Require(quotedArguments.exitCode == 0u && quotedArguments.stdoutBytes == "ARGUMENTS_OK",
                  L"Windows quoting preserves spaces, quotes, trailing backslashes, and Unicode arguments.");

    const TestSupport::ChildProcessResult nonzero = TestSupport::RunChildProcess(
        {.executablePath = servicePath,
         .arguments      = {L"--test-support-child-probe=exit"},
         .timeout        = std::chrono::milliseconds(SelfTest::ScaleTimeout(5'000))});
    state.Require(nonzero.Completed(), nonzero.diagnostic);
    state.Require(nonzero.exitCode == 37u, std::format(L"Nonzero probe exit code changed: {}.", nonzero.exitCode));
    state.Require(nonzero.stdoutBytes == "NONZERO_STDOUT" && nonzero.stderrBytes == "NONZERO_STDERR",
                  L"Nonzero child completion retains independent stdout and stderr.");

    SECURITY_ATTRIBUTES eventSecurity{};
    eventSecurity.nLength        = sizeof(eventSecurity);
    eventSecurity.bInheritHandle = TRUE;
    wil::unique_handle inheritedEvent(CreateEventW(&eventSecurity, TRUE, FALSE, nullptr));
    state.Require(inheritedEvent != nullptr, L"Inherited-handle contract creates an inheritable event.");
    const TestSupport::ChildProcessResult inheritedHandle = TestSupport::RunChildProcess(
        {.executablePath = servicePath,
         .arguments = {L"--test-support-child-probe=inherited-handle",
                       std::format(L"{}", reinterpret_cast<uintptr_t>(inheritedEvent.get()))},
         .timeout = std::chrono::milliseconds(SelfTest::ScaleTimeout(5'000))});
    state.Require(inheritedHandle.Completed(), inheritedHandle.diagnostic);
    state.Require(inheritedHandle.exitCode == 0u && inheritedHandle.stdoutBytes == "HANDLE_CLOSED" &&
                      WaitForSingleObject(inheritedEvent.get(), 0u) == WAIT_TIMEOUT,
                  L"PROC_THREAD_ATTRIBUTE_HANDLE_LIST prevents unrelated inheritable handles from reaching the child.");

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"test_support_child_process_runner", caseRoot),
                  L"Failed to prepare the child-process runner contract root.");
    const std::filesystem::path markerPath = caseRoot / L"descendant-survived.marker";
    const TestSupport::ChildProcessResult timedOut = TestSupport::RunChildProcess(
        {.executablePath = servicePath,
         .arguments      = {L"--test-support-child-probe=spawn-tree", markerPath.wstring()},
         .timeout        = std::chrono::milliseconds(SelfTest::ScaleTimeout(100))});
    state.Require(timedOut.status == TestSupport::ChildProcessStatus::TimedOut && timedOut.diagnostic.contains(L"timed out"),
                  L"Child-process timeout is explicit and diagnostic-rich.");
    std::this_thread::sleep_for(std::chrono::milliseconds(SelfTest::ScaleTimeout(1'000)));
    state.Require(! std::filesystem::exists(markerPath, existsEc),
                  L"Kill-on-close JobObject terminates the timed-out root and delayed descendant before marker creation.");

    std::stop_source cancellationSource;
    std::jthread cancelRequester([&] noexcept
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(SelfTest::ScaleTimeout(50)));
        static_cast<void>(cancellationSource.request_stop());
    });
    TestSupport::ChildProcessOptions cancellationOptions{
        .executablePath = servicePath,
        .arguments      = {L"--test-support-child-probe=sleep"},
        .timeout        = std::chrono::milliseconds(SelfTest::ScaleTimeout(5'000)),
    };
    cancellationOptions.cancellationToken = cancellationSource.get_token();
    const TestSupport::ChildProcessResult cancelled = TestSupport::RunChildProcess(cancellationOptions);
    state.Require(cancelled.status == TestSupport::ChildProcessStatus::Cancelled && cancelled.elapsed < std::chrono::seconds(2),
                  L"Cancellation stops and contains the child process before the timeout budget.");

    const TestSupport::ChildProcessResult launchFailure = TestSupport::RunChildProcess(
        {.executablePath = caseRoot / L"missing-child.exe", .timeout = std::chrono::milliseconds(SelfTest::ScaleTimeout(1'000))});
    state.Require(launchFailure.status == TestSupport::ChildProcessStatus::LaunchFailed && launchFailure.win32Error != ERROR_SUCCESS,
                  L"Launch failure returns a deterministic Win32 diagnostic without leaking capture workers.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_help_lists_cli_options",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::filesystem::path servicePath = GetSiblingExecutablePath(L"RedSalamanderSearchService.exe");
    std::error_code existsEc;
    state.Require(! servicePath.empty() && std::filesystem::exists(servicePath, existsEc),
                  std::format(L"Service executable not found: {}", servicePath.wstring()));
    if (! state.failure.empty())
    {
        return false;
    }

    CapturedProcessResult result{};
    std::wstring runError;
    state.Require(RunProcessAndCaptureOutput(
                      servicePath.wstring(), {L"--help"}, static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), result, runError),
                  runError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(result.exitCode == 0u, std::format(L"Expected --help exit code 0, got {}.", result.exitCode));
    state.Require(result.output.contains("--run-foreground"), L"--help output should mention --run-foreground.");
    state.Require(result.output.contains("--compact"), L"--help output should mention --compact.");
    state.Require(result.output.contains("--request-compact"), L"--help output should mention --request-compact.");
    state.Require(result.output.contains("--register"), L"--help output should mention --register.");
    state.Require(result.output.contains("--unregister"), L"--help output should mention --unregister.");
    state.Require(result.output.contains("--pipe-name="), L"--help output should mention --pipe-name.");
    state.Require(result.output.contains("--storage-root="), L"--help output should mention --storage-root.");
    state.Require(result.output.contains("--store-backend="), L"--help output should mention --store-backend.");
    state.Require(result.output.contains("--sqlite-path="), L"--help output should mention --sqlite-path.");
    state.Require(result.output.contains("VACUUM"), L"--help output should explain that --compact runs VACUUM.");
#ifdef _DEBUG
    state.Require(result.output.contains("RedSalamanderSearchService.Debug"), L"--help output should mention the Debug service identity.");
#else
    state.Require(result.output.contains("RedSalamanderSearchService"), L"--help output should mention the Release service identity.");
#endif
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_compact_cli_runs_manual_sqlite_maintenance",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_compact_cli", caseRoot), L"Failed to prepare search_service_compact_cli root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path servicePath = GetSiblingExecutablePath(L"RedSalamanderSearchService.exe");
    std::error_code existsEc;
    state.Require(! servicePath.empty() && std::filesystem::exists(servicePath, existsEc),
                  std::format(L"Service executable not found: {}", servicePath.wstring()));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path databasePath = caseRoot / L"manual.sqlite3";
    SqliteIndexStore::StoreInfo beforeInfo{};
    std::wstring prepareError;
    state.Require(PrepareSqliteMaintenanceStore(caseRoot, databasePath, beforeInfo, prepareError), prepareError);
    if (! state.failure.empty())
    {
        return false;
    }

    CapturedProcessResult result{};
    std::wstring runError;
    const std::vector<std::wstring> arguments{L"--compact", std::format(L"--sqlite-path={}", databasePath.wstring())};
    state.Require(RunProcessAndCaptureOutput(servicePath.wstring(), arguments, static_cast<DWORD>(SelfTest::ScaleTimeout(15'000)), result, runError),
                  runError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(result.exitCode == 0u, std::format(L"Expected --compact exit code 0, got {}.", result.exitCode));
    state.Require(result.output.contains("Compaction completed"), L"--compact output should confirm successful compaction.");
    state.Require(result.output.contains("free pages"), L"--compact output should summarize freelist compaction.");

    SqliteIndexStore::StoreInfo afterInfo{};
    const HRESULT inspectHr = SqliteIndexStore::InspectStore(databasePath.wstring(), afterInfo);
    state.Require(SUCCEEDED(inspectHr),
                  std::format(L"SqliteIndexStore::InspectStore after --compact failed. hr=0x{:08X}", static_cast<unsigned long>(inspectHr)));
    if (FAILED(inspectHr))
    {
        return false;
    }

    state.Require(afterInfo.freelistPageCount < beforeInfo.freelistPageCount,
                  std::format(L"Expected --compact to reduce freelist pages, before={} after={}.", beforeInfo.freelistPageCount, afterInfo.freelistPageCount));
    state.Require(! afterInfo.lastCheckpointUtc.empty(), L"--compact should record lastCheckpointUtc.");
    state.Require(! afterInfo.lastCompactionUtc.empty(), L"--compact should record lastCompactionUtc.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_compact_request_roundtrip",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_compact_request", caseRoot), L"Failed to prepare search_service_compact_request root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path servicePath = GetSiblingExecutablePath(L"RedSalamanderSearchService.exe");
    std::error_code existsEc;
    state.Require(! servicePath.empty() && std::filesystem::exists(servicePath, existsEc),
                  std::format(L"Service executable not found: {}", servicePath.wstring()));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path sqlitePath = caseRoot / L"live.sqlite3";
    constexpr size_t kExplicitCompactLargeEntryCount    = 2048u;
    constexpr size_t kExplicitCompactRetainedEntryCount = 1800u;

    SqliteIndexStore::StoreInfo beforeInfo{};
    std::wstring prepareError;
    state.Require(PrepareSqliteMaintenanceStore(caseRoot,
                                                sqlitePath,
                                                beforeInfo,
                                                prepareError,
                                                kExplicitCompactLargeEntryCount,
                                                kExplicitCompactRetainedEntryCount),
                  prepareError);
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus preCompactStatus{};
    if (! WaitForSearchServiceStatus(state, preCompactStatus, pipeName, L"live compact request preflight service"))
    {
        return false;
    }

    state.Require(preCompactStatus.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                  L"Live compaction preflight should report the SQLite backend.");
    state.Require(preCompactStatus.persistentStoreFreelistPageCount > 0u,
                  L"Live compaction preflight should retain reclaimable free pages for the explicit request.");
    state.Require(! preCompactStatus.maintenanceQueued,
                  L"Live compaction preflight should not let automatic idle maintenance race the explicit --request-compact path.");
    state.Require(! preCompactStatus.maintenanceRunning,
                  L"Live compaction preflight should not start while automatic idle maintenance is running.");
    if (! state.failure.empty())
    {
        return false;
    }

    CapturedProcessResult result{};
    std::wstring runError;
    const std::vector<std::wstring> requestCompactArguments{L"--request-compact", std::format(L"--pipe-name={}", pipeName)};
    state.Require(RunProcessAndCaptureOutput(
                      servicePath.wstring(), requestCompactArguments, static_cast<DWORD>(SelfTest::ScaleTimeout(15'000)), result, runError),
                  runError);
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring requestCompactOutput(result.output.begin(), result.output.end());
    state.Require(result.exitCode == 0u,
                  std::format(L"Expected --request-compact exit code 0, got {}. output='{}'.", result.exitCode, requestCompactOutput));
    state.Require(result.output.contains("Live compaction completed"), L"--request-compact output should confirm successful live compaction.");
    state.Require(result.output.contains("free pages="), L"--request-compact output should summarize the refreshed freelist state.");

    SearchServiceBroker::ServiceStatus status{};
    if (! WaitForSearchServiceStatus(state, status, pipeName, L"live compact request service"))
    {
        return false;
    }

    state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                  L"Live compaction service should report the SQLite backend.");
    state.Require(status.persistentStoreFreelistPageCount < beforeInfo.freelistPageCount,
                  std::format(L"Expected --request-compact to reduce freelist pages, before={} after={}.",
                              beforeInfo.freelistPageCount,
                              status.persistentStoreFreelistPageCount));
    state.Require(! status.lastCheckpointUtc.empty(), L"--request-compact should refresh lastCheckpointUtc through the running service.");
    state.Require(! status.lastCompactionUtc.empty(), L"--request-compact should refresh lastCompactionUtc through the running service.");
    state.Require(! status.maintenanceRunning, L"--request-compact should complete before returning status to the client.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::string serviceOutput;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), serviceOutput, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(serviceOutput.contains("Maintenance running"), L"Foreground live compaction output should report running maintenance.");
    state.Require(serviceOutput.contains("Maintenance completed"), L"Foreground live compaction output should report completed maintenance.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_status_reports_maintenance_history",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_status_reports_maintenance_history", caseRoot),
                  L"Failed to prepare search_service_sqlite_status_reports_maintenance_history root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path sqlitePath = caseRoot / L"service-maint.sqlite3";
    SqliteIndexStore::StoreInfo beforeInfo{};
    std::wstring prepareError;
    state.Require(PrepareSqliteMaintenanceStore(caseRoot, sqlitePath, beforeInfo, prepareError), prepareError);
    if (! state.failure.empty())
    {
        return false;
    }

    SqliteIndexStore::ManualMaintenanceResult maintenance{};
    HRESULT hr = SqliteIndexStore::RunManualMaintenance(sqlitePath.wstring(), &maintenance);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::RunManualMaintenance for service-status test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, false, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus status{};
    if (! WaitForSearchServiceStatus(state, status, pipeName, L"maintenance-history SQLite service"))
    {
        return false;
    }

    state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                  L"Maintenance-history service status should report the SQLite backend.");
    state.Require(
        status.persistentStorePageCount == maintenance.after.pageCount,
        std::format(L"Expected service status pageCount={} after maintenance, got {}.", maintenance.after.pageCount, status.persistentStorePageCount));
    state.Require(status.persistentStoreFreelistPageCount == maintenance.after.freelistPageCount,
                  std::format(L"Expected service status freelistPageCount={} after maintenance, got {}.",
                              maintenance.after.freelistPageCount,
                              status.persistentStoreFreelistPageCount));
    state.Require(status.lastCheckpointUtc == maintenance.after.lastCheckpointUtc, L"Service status should report the stored lastCheckpointUtc value.");
    state.Require(status.lastCompactionUtc == maintenance.after.lastCompactionUtc, L"Service status should report the stored lastCompactionUtc value.");
    state.Require(! status.maintenanceQueued, L"Maintenance-history status should not report queued maintenance without a scheduler.");
    state.Require(! status.maintenanceRunning, L"Maintenance-history status should not report running maintenance without a scheduler.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_idle_maintenance_queue_and_completion",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_idle_maintenance", caseRoot),
                  L"Failed to prepare search_service_sqlite_idle_maintenance root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path sqlitePath = caseRoot / L"idle-maint.sqlite3";
    SqliteIndexStore::StoreInfo beforeInfo{};
    std::wstring prepareError;
    state.Require(PrepareSqliteMaintenanceStore(caseRoot, sqlitePath, beforeInfo, prepareError), prepareError);
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus queuedStatus{};
    HRESULT hr = SearchServiceBroker::GetStatus(queuedStatus);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SearchServiceBroker::GetStatus for queued idle maintenance failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(queuedStatus.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                  L"Idle-maintenance service status should report the SQLite backend.");
    state.Require(queuedStatus.maintenanceQueued, L"Idle-maintenance service status should report queued maintenance before the idle grace window expires.");
    state.Require(! queuedStatus.maintenanceRunning,
                  L"Idle-maintenance service status should not report running maintenance before the idle grace window expires.");

    std::this_thread::sleep_for(std::chrono::milliseconds{1'300});

    SearchServiceBroker::ServiceStatus completedStatus{};
    hr = SearchServiceBroker::GetStatus(completedStatus);
    state.Require(SUCCEEDED(hr), std::format(L"SearchServiceBroker::GetStatus after idle maintenance failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(! completedStatus.maintenanceQueued, L"Idle-maintenance service status should clear the queued flag after maintenance completes.");
    state.Require(! completedStatus.maintenanceRunning, L"Idle-maintenance service status should not report running maintenance after completion.");
    state.Require(completedStatus.persistentStoreFreelistPageCount < beforeInfo.freelistPageCount,
                  std::format(L"Idle maintenance should reduce freelist pages, before={} after={}.",
                              beforeInfo.freelistPageCount,
                              completedStatus.persistentStoreFreelistPageCount));
    state.Require(! completedStatus.lastCheckpointUtc.empty(), L"Idle maintenance should record lastCheckpointUtc.");
    state.Require(! completedStatus.lastCompactionUtc.empty(), L"Idle maintenance should record lastCompactionUtc.");

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(10'000)), output, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(output.contains("Maintenance queued"), L"Foreground service log should report queued idle maintenance.");
    state.Require(output.contains("Maintenance running"), L"Foreground service log should report running idle maintenance.");
    state.Require(output.contains("Maintenance completed"), L"Foreground service log should report completed idle maintenance.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_delete_burst_maintenance_preserves_query_parity",
                  [&](SelfTest::CaseState& state) noexcept
{
    constexpr size_t kLargeDeleteBurstEntryCount = 32768u;
    constexpr size_t kCompactedEntryCount        = 32u;

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_delete_burst_maintenance", caseRoot),
                  L"Failed to prepare search_service_sqlite_delete_burst_maintenance root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path sqlitePath = caseRoot / L"delete-burst.sqlite3";
    SqliteIndexStore::StoreInfo beforeInfo{};
    std::wstring prepareError;
    state.Require(PrepareSqliteMaintenanceStore(caseRoot, sqlitePath, beforeInfo, prepareError, kLargeDeleteBurstEntryCount, kCompactedEntryCount, true),
                  prepareError);
    if (! state.failure.empty())
    {
        return false;
    }

    std::vector<std::wstring> expectedNames;
    {
        const auto compactedEntries = BuildSyntheticSqliteEntries(caseRoot, kCompactedEntryCount);
        expectedNames.reserve(compactedEntries.size());
        for (const auto& entry : compactedEntries)
        {
            expectedNames.push_back(entry.name);
        }
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    // Observe the queued state with this case's predicate instead of consuming it in
    // ForegroundSearchServiceProcess::Start. That readiness request can release the
    // service into a bounded compaction window where no pipe is listening.
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, false, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    const auto formatMaintenanceStatus = [](const SearchServiceBroker::ServiceStatus& status) noexcept
    {
        return std::format(L"freelist={} bytes={} wal={} queued={} running={} checkpoint='{}' compaction='{}'",
                           status.persistentStoreFreelistPageCount,
                           status.persistentStoreBytes,
                           status.writeAheadLogBytes,
                           status.maintenanceQueued ? L"yes" : L"no",
                           status.maintenanceRunning ? L"yes" : L"no",
                           status.lastCheckpointUtc,
                           status.lastCompactionUtc);
    };

    auto waitForStatus = [&](auto&& predicate, SearchServiceBroker::ServiceStatus& outStatus, std::wstring_view failureMessage) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() +
                              std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(
                                  SelfTest::ScaleTimeout(kForegroundSearchServiceReadinessTimeoutMs))};
        HRESULT lastHr = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        SearchServiceBroker::ServiceStatus lastStatus{};
        bool hasLastStatus = false;
        while (std::chrono::steady_clock::now() < deadline)
        {
            SearchServiceBroker::ServiceStatus candidate{};
            lastHr = SearchServiceBroker::GetStatus(candidate);
            if (SUCCEEDED(lastHr))
            {
                lastStatus    = candidate;
                hasLastStatus = true;
                if (predicate(candidate))
                {
                    outStatus = std::move(candidate);
                    return true;
                }
            }

            const std::wstring exitSummary = service.TryCaptureExitedOutputForFailure();
            if (! exitSummary.empty())
            {
                state.Require(false,
                              std::format(L"Foreground search service exited while waiting for {}. lastHr=0x{:08X}{}",
                                          failureMessage,
                                          static_cast<unsigned long>(lastHr),
                                          exitSummary));
                return false;
            }

            std::this_thread::sleep_for(std::chrono::milliseconds{200});
        }

        state.Require(false,
                      std::format(L"{} lastHr=0x{:08X}{}",
                                  failureMessage,
                                  static_cast<unsigned long>(lastHr),
                                  hasLastStatus ? std::format(L" lastStatus=({})", formatMaintenanceStatus(lastStatus)) :
                                                  service.TryCaptureExitedOutputForFailure()));
        return false;
    };

    SearchServiceBroker::ServiceStatus initialStatus{};
    state.Require(waitForStatus([&](const SearchServiceBroker::ServiceStatus& status) noexcept
    {
        return (status.maintenanceQueued && status.persistentStoreFreelistPageCount >= beforeInfo.freelistPageCount) ||
               status.persistentStoreFreelistPageCount < beforeInfo.freelistPageCount;
    },
                                initialStatus,
                                L"Timed out waiting for the delete-burst store to queue maintenance or complete its first compaction window."),
                  L"Timed out waiting for the delete-burst store to queue maintenance or complete its first compaction window.");
    if (! state.failure.empty())
    {
        return false;
    }
    const bool firstWindowAlreadyObserved = initialStatus.persistentStoreFreelistPageCount < beforeInfo.freelistPageCount;

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    auto queryNames = [&](std::vector<std::wstring>& outNames, std::wstring_view context) noexcept
    {
        outNames.clear();
        std::vector<LocalSearchIndexCore::Candidate> candidates;
        LocalSearchIndexCore::QueryStats stats{};
        const HRESULT hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, candidates, &stats);
        state.Require(SUCCEEDED(hr), std::format(L"SearchServiceBroker::Query {} failed. hr=0x{:08X}", context, static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        outNames = CollectIndexedCandidateNames(candidates);
        return true;
    };

    std::vector<std::wstring> baselineNames;
    state.Require(queryNames(baselineNames, L"before maintenance"), L"Baseline query failed.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(baselineNames == expectedNames, L"Delete-burst baseline query should match the compacted SQLite entry set.");

    SearchServiceBroker::ServiceStatus afterFirstWindow{};
    if (firstWindowAlreadyObserved)
    {
        afterFirstWindow = initialStatus;
    }
    else
    {
        state.Require(waitForStatus([&](const SearchServiceBroker::ServiceStatus& status) noexcept
        { return status.persistentStoreFreelistPageCount < initialStatus.persistentStoreFreelistPageCount; },
                                    afterFirstWindow,
                                    L"Timed out waiting for the first idle maintenance window to reclaim space."),
                      L"Timed out waiting for the first idle maintenance window to reclaim space.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(afterFirstWindow.persistentStoreFreelistPageCount < beforeInfo.freelistPageCount,
                  std::format(L"First idle maintenance window should reclaim freelist pages, before={} after={} initialStatus=({}).",
                              beforeInfo.freelistPageCount,
                              afterFirstWindow.persistentStoreFreelistPageCount,
                              formatMaintenanceStatus(initialStatus)));
    state.Require(afterFirstWindow.maintenanceQueued || afterFirstWindow.maintenanceRunning,
                  L"Delete-burst maintenance should still have more work queued after the first bounded compaction pass.");
    state.Require(afterFirstWindow.persistentStoreBytes < beforeInfo.databaseBytes,
                  std::format(L"First idle maintenance window should shrink the database file, before={} after={}.",
                              beforeInfo.databaseBytes,
                              afterFirstWindow.persistentStoreBytes));
    if (beforeInfo.writeAheadLogBytes != 0u)
    {
        state.Require(afterFirstWindow.writeAheadLogBytes < beforeInfo.writeAheadLogBytes,
                      std::format(L"First idle maintenance window should shrink the WAL, before={} after={}.",
                                  beforeInfo.writeAheadLogBytes,
                                  afterFirstWindow.writeAheadLogBytes));
    }
    else
    {
        state.Require(
            afterFirstWindow.writeAheadLogBytes == 0u,
            std::format(L"First idle maintenance window should keep an eagerly checkpointed WAL empty, got {}.", afterFirstWindow.writeAheadLogBytes));
    }

    std::vector<std::wstring> afterFirstNames;
    state.Require(queryNames(afterFirstNames, L"after first maintenance window"), L"First maintenance-window query failed.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(afterFirstNames == baselineNames, L"Delete-burst query parity should hold after the first idle maintenance window.");

    SearchServiceBroker::ServiceStatus afterSecondWindow{};
    state.Require(waitForStatus([&](const SearchServiceBroker::ServiceStatus& status) noexcept
    { return status.persistentStoreFreelistPageCount < afterFirstWindow.persistentStoreFreelistPageCount; },
                                afterSecondWindow,
                                L"Timed out waiting for the second idle maintenance window to reclaim more space."),
                  L"Timed out waiting for the second idle maintenance window to reclaim more space.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(afterSecondWindow.persistentStoreFreelistPageCount < afterFirstWindow.persistentStoreFreelistPageCount,
                  std::format(L"Second idle maintenance window should keep reducing freelist pages, first={} second={}.",
                              afterFirstWindow.persistentStoreFreelistPageCount,
                              afterSecondWindow.persistentStoreFreelistPageCount));
    state.Require(afterSecondWindow.writeAheadLogBytes <= afterFirstWindow.writeAheadLogBytes,
                  std::format(L"Second idle maintenance window should not regrow the WAL, first={} second={}.",
                              afterFirstWindow.writeAheadLogBytes,
                              afterSecondWindow.writeAheadLogBytes));

    std::vector<std::wstring> afterSecondNames;
    state.Require(queryNames(afterSecondNames, L"after second maintenance window"), L"Second maintenance-window query failed.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(afterSecondNames == baselineNames, L"Delete-burst query parity should hold after the second idle maintenance window.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_bootstrap_status_roundtrip",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_bootstrap", caseRoot), L"Failed to prepare search_service_sqlite_bootstrap root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    ULARGE_INTEGER expectedWriteTime{};
    expectedWriteTime.QuadPart = 133838640000000000ull;
    FILETIME expectedWriteTimeFile{};
    expectedWriteTimeFile.dwLowDateTime  = expectedWriteTime.LowPart;
    expectedWriteTimeFile.dwHighDateTime = expectedWriteTime.HighPart;
    state.Require(SetFileLastWriteTime(caseRoot / L"alpha.txt", expectedWriteTimeFile), L"Failed to set alpha.txt last write time.");
    ObservedFileMetadata expectedMetadata{};
    state.Require(TryReadObservedFileMetadata(caseRoot / L"alpha.txt", expectedMetadata), L"Failed to read expected alpha.txt metadata.");

    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath  = caseRoot / L"service-index.sqlite3";
    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, false, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus status{};
    HRESULT hr = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(hr), std::format(L"SearchServiceBroker::GetStatus for SQLite service failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite, L"SQLite service status should report the SQLite backend.");
    state.Require(EqualsIgnoreCase(status.storageRootDirectory, storageRoot.wstring()),
                  L"SQLite service status should report the overridden compatibility snapshot root.");
    state.Require(EqualsIgnoreCase(status.persistentStorePath, sqlitePath.wstring()), L"SQLite service status should report the configured database path.");
    state.Require(status.persistentStoreBytes > 0u, L"SQLite service status should report database bytes after bootstrap.");
    state.Require(EqualsIgnoreCase(status.writeAheadLogPath, sqlitePath.wstring() + L"-wal"), L"SQLite service status should report the configured WAL path.");
    state.Require(status.persistentStorePageCount != 0u, L"SQLite service status should report a non-zero SQLite page count after bootstrap.");
    state.Require(status.lastCheckpointUtc.empty(), L"Fresh SQLite bootstrap status should not report a lastCheckpointUtc value before maintenance.");
    state.Require(status.lastCompactionUtc.empty(), L"Fresh SQLite bootstrap status should not report a lastCompactionUtc value before maintenance.");
    state.Require(! status.maintenanceQueued, L"Fresh SQLite bootstrap status should not report queued maintenance.");
    state.Require(! status.maintenanceRunning, L"Fresh SQLite bootstrap status should not report running maintenance.");
    state.Require(status.persistentStoreInspectionSucceeded, L"SQLite service status should report successful store inspection after bootstrap.");
    state.Require(status.readyForQueryCutover, L"Empty SQLite service store should report readyForQueryCutover after bootstrap.");
    state.Require(status.indexedVolumeCount == 0u, std::format(L"Expected zero mirrored volumes before query, got {}.", status.indexedVolumeCount));
    state.Require(status.indexedEntryCount == 0u, std::format(L"Expected zero mirrored entries before query, got {}.", status.indexedEntryCount));
    state.Require(status.legacyImportVolumeCount == 0u,
                  std::format(L"Expected zero legacy-import volumes before query, got {}.", status.legacyImportVolumeCount));
    state.Require(status.autoCheckpointEnabled, L"SQLite service status should report automatic checkpointing.");
    state.Require(status.autoCompactionEnabled, L"SQLite service status should report automatic compaction.");
    state.Require(status.storeState == LocalSearchIndexCore::StoreState::Ready,
                  L"Fresh SQLite bootstrap status should report a ready store state before the first query.");
    state.Require(status.syncPhase == LocalSearchIndexCore::SyncPhase::Watching,
                  L"Fresh SQLite bootstrap status should report the watching sync phase before the first query.");
    state.Require(status.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::Unknown,
                  L"Fresh SQLite bootstrap status should not report a query execution mode before the first query.");
    state.Require(status.fallbackReason == LocalSearchIndexCore::FallbackReason::None,
                  L"Fresh SQLite bootstrap status should not report a fallback reason before the first query.");

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    LocalSearchIndexCore::QueryStats queryStats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, candidates, &queryStats);
    state.Require(SUCCEEDED(hr), std::format(L"SQLite-backed service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"SQLite-backed service query returned an unexpected candidate set.");
    state.Require(! queryStats.snapshotSaved,
                  L"First SQLite-backed service query should seed directly into SQLite without creating a compatibility snapshot runtime store.");
    state.Require(queryStats.snapshotPath.empty(), L"First SQLite-backed service query should not report a compatibility snapshot path.");
    state.Require(queryStats.snapshotFileBytes == 0u,
                  std::format(L"First SQLite-backed service query should report snapshotFileBytes=0 when no compatibility snapshot exists. got={}",
                              queryStats.snapshotFileBytes));
    state.Require(! queryStats.usedSqliteStore, L"First SQLite-backed no-wait service query should not use SQLite before the root is mirrored.");
    state.Require(queryStats.usedLiveScanFallback, L"First SQLite-backed no-wait service query should report live-scan fallback usage.");
    state.Require(queryStats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                  L"First SQLite-backed no-wait service query should report the live-scan fallback execution mode.");
    state.Require(queryStats.fallbackReason == LocalSearchIndexCore::FallbackReason::StoreMissing,
                  L"First SQLite-backed no-wait service query should report store-missing until warmup or rebuild mirrors the root.");
    state.Require(! queryStats.sqliteCutoverBlocked,
                  L"First SQLite-backed no-wait service query should distinguish a missing root from a stale-root cutover block.");

    LocalSearchIndexCore::QueryStats warmQueryStats{};
    std::vector<LocalSearchIndexCore::Candidate> warmCandidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, warmCandidates, &warmQueryStats);
    state.Require(SUCCEEDED(hr), std::format(L"Warm SQLite-backed service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(warmCandidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Warm SQLite-backed service query returned an unexpected candidate set.");
    state.Require(! warmQueryStats.snapshotLoaded, L"Warm SQLite-backed service query should not need to reload the snapshot runtime store.");
    state.Require(! warmQueryStats.usedSqliteStore, L"Warm SQLite-backed no-wait service query should not use SQLite before warmup mirrors the root.");
    state.Require(! warmQueryStats.usedNamePrefilter,
                  L"Warm SQLite-backed no-wait service query should not report a SQLite prefilter before direct SQLite cutover is open.");
    state.Require(warmQueryStats.usedLiveScanFallback, L"Warm SQLite-backed no-wait service query should report live-scan fallback usage.");
    state.Require(warmQueryStats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                  L"Warm SQLite-backed no-wait service query should report the live-scan fallback execution mode.");
    state.Require(warmQueryStats.fallbackReason == LocalSearchIndexCore::FallbackReason::StoreMissing,
                  L"Warm SQLite-backed no-wait service query should keep reporting store-missing until warmup or rebuild mirrors the root.");
    state.Require(! warmQueryStats.sqliteCutoverBlocked,
                  L"Warm SQLite-backed no-wait service query should distinguish a missing root from a stale-root cutover block.");
    state.Require(warmCandidates.size() == 1u, std::format(L"Warm SQLite-backed service query expected 1 candidate, got {}.", warmCandidates.size()));
    if (warmCandidates.size() == 1u && warmQueryStats.usedSqliteStore)
    {
        const auto& candidate = warmCandidates.front();
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_CREATION_TIME) != 0u,
                      L"Warm SQLite-backed service query should transport persisted creation-time metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_LAST_ACCESS_TIME) != 0u,
                      L"Warm SQLite-backed service query should transport persisted last-access metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_END_OF_FILE) != 0u,
                      L"Warm SQLite-backed service query should transport persisted size metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_LAST_WRITE_TIME) != 0u,
                      L"Warm SQLite-backed service query should transport persisted last-write metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_CHANGE_TIME) != 0u,
                      L"Warm SQLite-backed service query should transport persisted change-time metadata.");
        state.Require((candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_ALLOCATION_SIZE) != 0u,
                      L"Warm SQLite-backed service query should transport persisted allocation-size metadata.");
        state.Require(candidate.creationTime100ns == expectedMetadata.creationTime,
                      std::format(L"Warm SQLite-backed service query creation-time mismatch. got={} expected={}",
                                  candidate.creationTime100ns,
                                  expectedMetadata.creationTime));
        state.Require(candidate.lastAccessTime100ns == expectedMetadata.lastAccessTime,
                      std::format(L"Warm SQLite-backed service query last-access mismatch. got={} expected={}",
                                  candidate.lastAccessTime100ns,
                                  expectedMetadata.lastAccessTime));
        state.Require(candidate.endOfFile == 5, std::format(L"Warm SQLite-backed service query size mismatch. got={}", candidate.endOfFile));
        state.Require(candidate.lastWriteTime100ns == static_cast<int64_t>(expectedWriteTime.QuadPart),
                      std::format(L"Warm SQLite-backed service query last-write mismatch. got={} expected={}",
                                  candidate.lastWriteTime100ns,
                                  static_cast<long long>(expectedWriteTime.QuadPart)));
        state.Require(
            candidate.changeTime100ns == expectedMetadata.changeTime,
            std::format(L"Warm SQLite-backed service query change-time mismatch. got={} expected={}", candidate.changeTime100ns, expectedMetadata.changeTime));
        state.Require(candidate.allocationSize == expectedMetadata.allocationSize,
                      std::format(L"Warm SQLite-backed service query allocation-size mismatch. got={} expected={}",
                                  candidate.allocationSize,
                                  expectedMetadata.allocationSize));
    }

    hr = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(hr), std::format(L"SearchServiceBroker::GetStatus after SQLite query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(status.persistentStoreInspectionSucceeded, L"SQLite service status should keep reporting successful inspection after query.");
    state.Require(status.readyForQueryCutover, L"SQLite service status should keep an empty store ready for future query cutover.");
    state.Require(status.indexedVolumeCount == 0u, std::format(L"Expected zero mirrored volumes before warmup/rebuild, got {}.", status.indexedVolumeCount));
    state.Require(status.indexedEntryCount == 0u, std::format(L"Expected zero mirrored entries before warmup/rebuild, got {}.", status.indexedEntryCount));
    state.Require(status.legacyImportVolumeCount == 0u,
                  std::format(L"Expected zero legacy-import volumes after query, got {}.", status.legacyImportVolumeCount));

    SqliteIndexStore::StoreInfo storeInfo{};
    hr = SqliteIndexStore::InspectStore(sqlitePath.wstring(), storeInfo);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::InspectStore after service bootstrap failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(storeInfo.schemaReady, L"Service SQLite bootstrap should leave a ready schema on disk.");
    state.Require(storeInfo.schemaVersion == SqliteIndexStore::kSchemaVersion, L"Service SQLite bootstrap should stamp the expected schema version.");
    const auto snapshotFiles = CollectDirectoryFilesByExtension(storageRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"SQLite service bootstrap roundtrip should not create compatibility snapshot runtime files.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_cold_start_bypasses_snapshot_runtime_store",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_cold_start", caseRoot), L"Failed to prepare search_service_sqlite_cold_start root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");

    const std::filesystem::path storageRoot        = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath         = caseRoot / L"cold-start.sqlite3";
    const std::wstring previousPipeOverride        = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousWarmupDelayOverride = GetEnvVarTrimmed(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS");
    const std::wstring pipeName                    = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    state.Require(::SetEnvironmentVariableW(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS", L"5000") != 0,
                  L"Failed to override the startup warmup delay.");
    const auto restoreOverrides = wil::scope_exit([&] noexcept
    {
        const wchar_t* restorePipeValue  = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        const wchar_t* restoreDelayValue = previousWarmupDelayOverride.empty() ? nullptr : previousWarmupDelayOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restorePipeValue));
        static_cast<void>(::SetEnvironmentVariableW(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS", restoreDelayValue));
    });

    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    {
        ForegroundSearchServiceProcess service;
        std::wstring serviceError;
        state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
        if (! state.failure.empty())
        {
            return false;
        }

        SearchServiceBroker::ServiceStatus status{};
        HRESULT hr = SearchServiceBroker::GetStatus(status);
        state.Require(SUCCEEDED(hr),
                      std::format(L"SearchServiceBroker::GetStatus for cold-start SQLite service failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        state.Require(EqualsIgnoreCase(status.storageRootDirectory, storageRoot.wstring()),
                      L"Cold-start SQLite service should report the overridden storage root.");
        state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                      L"Cold-start SQLite service should report the SQLite backend.");

        BrokerProgressRecorder recorder;
        LocalSearchIndexCore::QueryStats coldStats{};
        std::vector<LocalSearchIndexCore::Candidate> coldCandidates;
        hr = SearchServiceBroker::Query(request, &BrokerProgressRecorder::ProgressThunk, &recorder, nullptr, nullptr, coldCandidates, &coldStats);
        state.Require(SUCCEEDED(hr), std::format(L"Cold-start SQLite service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        state.Require(CollectIndexedCandidateNames(coldCandidates) == std::vector<std::wstring>{L"alpha.txt"},
                      L"Cold-start SQLite service query returned an unexpected candidate set.");
        state.Require(! coldStats.usedSqliteStore, L"Cold-start SQLite service query should not block waiting for direct SQLite cutover.");
        state.Require(coldStats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                      std::format(L"Cold-start SQLite service query should answer from the live filesystem while sqlite is missing. mode={} "
                                  L"fallbackReason={} liveScan={}",
                                  static_cast<uint32_t>(coldStats.queryExecutionMode),
                                  static_cast<uint32_t>(coldStats.fallbackReason),
                                  coldStats.usedLiveScanFallback));
        state.Require(coldStats.fallbackReason == LocalSearchIndexCore::FallbackReason::StoreMissing,
                      L"Cold-start SQLite service query should report a missing-store fallback reason.");
        state.Require(coldStats.ensureReadyDurationMs == 0u, L"Cold-start SQLite service query should not spend time in EnsureReady on the request thread.");
        state.Require(! coldStats.snapshotLoaded, L"Cold-start SQLite service query should not load a legacy snapshot runtime store.");
        state.Require(! coldStats.snapshotSaved, L"Cold-start SQLite service query should not create a legacy snapshot runtime store.");
        state.Require(coldStats.snapshotPath.empty(), L"Cold-start SQLite service query should not report a legacy snapshot path.");
        state.Require(
            coldStats.snapshotFileBytes == 0u,
            std::format(L"Cold-start SQLite service query should report snapshotFileBytes=0 when no snapshot exists. got={}", coldStats.snapshotFileBytes));
        state.Require(coldStats.usedLiveScanFallback, L"Cold-start SQLite service live-scan path should report live-scan fallback usage.");

        const auto progress = recorder.Snapshots();
        state.Require(! progress.empty(), L"Cold-start SQLite service query should report broker progress.");
        const RecordedSearchProgress* enumerating = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING);
        state.Require(enumerating != nullptr, L"Cold-start SQLite service query should report an enumerating progress update.");
        if (enumerating != nullptr)
        {
            state.Require(enumerating->queryExecutionMode == coldStats.queryExecutionMode,
                          L"Cold-start SQLite service progress should report the same execution mode as query stats.");
            state.Require(enumerating->fallbackReason == LocalSearchIndexCore::FallbackReason::StoreMissing,
                          L"Cold-start SQLite service progress should report the missing-store fallback reason.");
            state.Require(EqualsIgnoreCase(enumerating->activeRoot, caseRoot.wstring()),
                          L"Cold-start SQLite service progress should report the active root being searched.");
            state.Require((enumerating->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u,
                          L"Cold-start SQLite service live-scan path should report degraded-no-index through broker progress.");
            state.Require(enumerating->storeState == LocalSearchIndexCore::StoreState::Invalid,
                          L"Cold-start SQLite service live-scan path should report an invalid/missing store state.");
            state.Require(enumerating->syncPhase == LocalSearchIndexCore::SyncPhase::Idle,
                          L"Cold-start SQLite service live-scan path should stay non-blocking while warmup is delayed.");
        }

        status = {};
        hr     = SearchServiceBroker::GetStatus(status);
        state.Require(SUCCEEDED(hr),
                      std::format(L"SearchServiceBroker::GetStatus after cold-start SQLite query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                      L"Cold-start SQLite status should keep reporting the SQLite backend after a degraded query.");
        state.Require(status.queryExecutionMode == coldStats.queryExecutionMode,
                      L"Cold-start SQLite status should report the same execution mode as query stats.");
        state.Require(status.fallbackReason == LocalSearchIndexCore::FallbackReason::StoreMissing,
                      L"Cold-start SQLite status should report a missing-store fallback reason after a degraded query.");
        state.Require(EqualsIgnoreCase(status.activeRoot, caseRoot.wstring()), L"Cold-start SQLite status should report the searched root as active.");
        state.Require(status.storeState == LocalSearchIndexCore::StoreState::Invalid,
                      L"Cold-start SQLite live-scan status should report an invalid/missing store.");
        state.Require(status.syncPhase == LocalSearchIndexCore::SyncPhase::Idle,
                      L"Cold-start SQLite live-scan status should remain non-blocking while warmup is delayed.");

        std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const auto snapshotFiles = CollectDirectoryFilesByExtension(storageRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Cold-start SQLite service query should not create a compatibility snapshot runtime file.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_ntfs_traversal_seed_stays_degraded",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_ntfs_traversal_seed", caseRoot),
                  L"Failed to prepare search_service_sqlite_ntfs_traversal_seed root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::Repository probeRepository;
    LocalSearchIndexCore::SupportInfo support{};
    HRESULT hr = probeRepository.ProbePath(caseRoot.wstring(), support);
    state.Require(SUCCEEDED(hr), std::format(L"ProbePath failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    if (support.fileSystemKind != LocalSearchIndexCore::FileSystemKind::Ntfs)
    {
        return state.Skip(L"NTFS traversal-seed degradation test requires an NTFS-backed self-test root.");
    }

    const std::filesystem::path storageRoot   = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath    = caseRoot / L"ntfs-traversal.sqlite3";
    const std::wstring previousPipeOverride   = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousRootOverride   = GetEnvVarTrimmed(SearchServiceBroker::kDiscoverRootsEnvVar);
    const std::wstring previousForceTraversal = GetEnvVarTrimmed(L"REDSALAMANDER_TEST_FORCE_NTFS_TRAVERSAL_SEED");
    const std::wstring pipeName               = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, caseRoot.wstring().c_str()) != 0,
                  L"Failed to override the search service startup roots.");
    state.Require(::SetEnvironmentVariableW(L"REDSALAMANDER_TEST_FORCE_NTFS_TRAVERSAL_SEED", L"1") != 0, L"Failed to force NTFS traversal seeding.");
    const auto restoreOverrides = wil::scope_exit([&] noexcept
    {
        const wchar_t* restorePipeValue      = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        const wchar_t* restoreRootsValue     = previousRootOverride.empty() ? nullptr : previousRootOverride.c_str();
        const wchar_t* restoreTraversalValue = previousForceTraversal.empty() ? nullptr : previousForceTraversal.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restorePipeValue));
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, restoreRootsValue));
        static_cast<void>(::SetEnvironmentVariableW(L"REDSALAMANDER_TEST_FORCE_NTFS_TRAVERSAL_SEED", restoreTraversalValue));
    });

    const std::wstring extraArgs = std::format(L"--storage-root=\"{}\" --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    auto waitForStatus = [&](auto&& predicate, SearchServiceBroker::ServiceStatus& outStatus, std::wstring_view failureMessage) noexcept
    {
        const auto deadline =
            std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(15'000))};
        HRESULT lastStatusHr = S_OK;
        while (std::chrono::steady_clock::now() < deadline)
        {
            SearchServiceBroker::ServiceStatus currentStatus{};
            const HRESULT statusHr = SearchServiceBroker::GetStatus(currentStatus);
            lastStatusHr           = statusHr;
            if (SUCCEEDED(statusHr))
            {
                outStatus = currentStatus;
            }
            if (SUCCEEDED(statusHr) && predicate(currentStatus))
            {
                return true;
            }

            std::this_thread::sleep_for(std::chrono::milliseconds{200});
        }

        state.Require(false,
                      std::format(L"{0} Last status: pid={1} storeKind={2} warmupEnabled={3} running={4} completed={5} failed={6} total={7} "
                                  L"currentRoot='{8}' indexedVolumes={9} indexedEntries={10} ready={11} inspected={12} storePath='{13}' "
                                  L"walBytes={14} storeState={15} syncPhase={16} activeRoot='{17}' discoveredRoots={18} lastFailure={19} "
                                  L"failureHr=0x{20:08X} failureRoot='{21}' lastGetStatusHr=0x{22:08X}.",
                                  failureMessage,
                                  outStatus.processId,
                                  static_cast<unsigned int>(outStatus.persistentStoreKind),
                                  outStatus.startupWarmupEnabled,
                                  outStatus.startupWarmupRunning,
                                  outStatus.startupWarmupCompletedRoots,
                                  outStatus.startupWarmupFailedRoots,
                                  outStatus.startupWarmupTotalRoots,
                                  outStatus.startupWarmupCurrentRoot,
                                  outStatus.indexedVolumeCount,
                                  outStatus.indexedEntryCount,
                                  outStatus.readyForQueryCutover,
                                  outStatus.persistentStoreInspectionSucceeded,
                                  outStatus.persistentStorePath,
                                  outStatus.writeAheadLogBytes,
                                  static_cast<unsigned int>(outStatus.storeState),
                                  static_cast<unsigned int>(outStatus.syncPhase),
                                  outStatus.activeRoot,
                                  outStatus.discoveredRoots.size(),
                                  outStatus.startupWarmupHasFailure,
                                  static_cast<unsigned long>(outStatus.startupWarmupLastFailureHr),
                                  outStatus.startupWarmupLastFailureRoot,
                                  static_cast<unsigned long>(lastStatusHr)));
        return false;
    };

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, false, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus status{};
    state.Require(waitForStatus(
                      [&](const SearchServiceBroker::ServiceStatus& currentStatus) noexcept
    {
        return currentStatus.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite && currentStatus.startupWarmupCompletedRoots >= 1u &&
               currentStatus.indexedVolumeCount >= 1u;
    },
                      status,
                      L"Forced NTFS traversal-seed service did not finish startup warmup."),
                  L"Forced NTFS traversal-seed service did not finish startup warmup.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(EqualsIgnoreCase(status.persistentStorePath, sqlitePath.wstring()),
                  L"Forced NTFS traversal-seed service should report the configured sqlite path.");

    SqliteIndexStore::VolumeInfo volumeInfo{};
    hr = SqliteIndexStore::InspectVolume(sqlitePath.wstring(), caseRoot.wstring(), volumeInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::InspectVolume for forced NTFS traversal seed failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(volumeInfo.state == SqliteIndexStore::kVolumeStateCurrentnessUnproven,
                  std::format(L"Forced NTFS traversal seed should persist CURRENTNESS_UNPROVEN, got {}.", volumeInfo.state));
    state.Require(volumeInfo.entryCount >= 2u,
                  std::format(L"Forced NTFS traversal seed should persist at least root + alpha.txt, got {} entries.", volumeInfo.entryCount));

    BrokerProgressRecorder recorder;
    LocalSearchIndexCore::QueryStats queryStats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    hr = SearchServiceBroker::Query(request, &BrokerProgressRecorder::ProgressThunk, &recorder, nullptr, nullptr, candidates, &queryStats);
    state.Require(SUCCEEDED(hr), std::format(L"Forced NTFS traversal-seed service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Forced NTFS traversal-seed service query returned an unexpected candidate set.");
    state.Require(queryStats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                  L"Forced NTFS traversal-seed service query should stay on the live-scan fallback path.");
    state.Require(queryStats.fallbackReason == LocalSearchIndexCore::FallbackReason::StoreStale,
                  L"Forced NTFS traversal-seed service query should report store-stale while currentness is unproven.");
    state.Require(queryStats.usedLiveScanFallback, L"Forced NTFS traversal-seed service query should report live-scan fallback usage.");
    state.Require(! queryStats.usedSqliteStore, L"Forced NTFS traversal-seed service query should not enumerate results directly from sqlite.");

    const auto progress = recorder.Snapshots();
    state.Require(! progress.empty(), L"Forced NTFS traversal-seed service query should report broker progress.");
    const RecordedSearchProgress* enumerating = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING);
    state.Require(enumerating != nullptr, L"Forced NTFS traversal-seed service query should report an enumerating progress update.");
    if (enumerating != nullptr)
    {
        state.Require((enumerating->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u,
                      L"Forced NTFS traversal-seed service progress should report degraded-no-index.");
        state.Require(enumerating->queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                      L"Forced NTFS traversal-seed service progress should report the live-scan fallback execution mode.");
        state.Require(enumerating->fallbackReason == LocalSearchIndexCore::FallbackReason::StoreStale,
                      L"Forced NTFS traversal-seed service progress should report a store-stale fallback reason.");
        state.Require(enumerating->storeState == LocalSearchIndexCore::StoreState::Syncing,
                      L"Forced NTFS traversal-seed service progress should report a syncing store state.");
        state.Require(enumerating->syncPhase == LocalSearchIndexCore::SyncPhase::Idle,
                      L"Forced NTFS traversal-seed service progress should remain idle on the request thread.");
    }

    status = {};
    hr     = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SearchServiceBroker::GetStatus after forced NTFS traversal-seed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(status.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                  L"Forced NTFS traversal-seed status should report the live-scan fallback execution mode.");
    state.Require(status.fallbackReason == LocalSearchIndexCore::FallbackReason::StoreStale,
                  L"Forced NTFS traversal-seed status should report a store-stale fallback reason.");
    state.Require(status.storeState == LocalSearchIndexCore::StoreState::Syncing, L"Forced NTFS traversal-seed status should report a syncing store state.");
    state.Require(status.syncPhase == LocalSearchIndexCore::SyncPhase::Idle, L"Forced NTFS traversal-seed status should remain idle after the degraded query.");

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    const auto snapshotFiles = CollectDirectoryFilesByExtension(storageRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Forced NTFS traversal-seed sqlite-authoritative service should not create compatibility snapshot files.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_cold_start_stale_root_refreshes_before_query",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_cold_start_stale_root", caseRoot),
                  L"Failed to prepare search_service_sqlite_cold_start_stale_root root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");

    const std::filesystem::path storageRoot        = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath         = caseRoot / L"stale-root.sqlite3";
    const std::wstring previousPipeOverride        = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousRootOverride        = GetEnvVarTrimmed(SearchServiceBroker::kDiscoverRootsEnvVar);
    const std::wstring previousWarmupDelayOverride = GetEnvVarTrimmed(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS");
    const std::wstring pipeName                    = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, caseRoot.wstring().c_str()) != 0,
                  L"Failed to override the search service startup roots.");
    state.Require(::SetEnvironmentVariableW(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS", nullptr) != 0,
                  L"Failed to clear the startup warmup delay override.");
    const auto restoreOverrides = wil::scope_exit([&] noexcept
    {
        const wchar_t* restorePipeValue  = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        const wchar_t* restoreRootsValue = previousRootOverride.empty() ? nullptr : previousRootOverride.c_str();
        const wchar_t* restoreDelayValue = previousWarmupDelayOverride.empty() ? nullptr : previousWarmupDelayOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restorePipeValue));
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, restoreRootsValue));
        static_cast<void>(::SetEnvironmentVariableW(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS", restoreDelayValue));
    });

    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    bool seedReadyForQueryCutover = false;
    {
        ForegroundSearchServiceProcess service;
        std::wstring serviceError;
        state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, false, extraArgs), serviceError);
        if (! state.failure.empty())
        {
            return false;
        }

        auto waitForStatus = [&](auto&& predicate, SearchServiceBroker::ServiceStatus& outStatus, std::wstring_view failureMessage) noexcept
        {
            const auto deadline =
                std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(60'000))};
            HRESULT lastStatusHr = S_OK;
            while (std::chrono::steady_clock::now() < deadline)
            {
                const HRESULT hr = SearchServiceBroker::GetStatus(outStatus);
                lastStatusHr     = hr;
                if (SUCCEEDED(hr) && predicate(outStatus))
                {
                    return true;
                }

                std::this_thread::sleep_for(std::chrono::milliseconds{200});
            }

            state.Require(false,
                          std::format(L"{0} Last status: pid={1} backend={2} ready={3} inspected={4} warmupEnabled={5} running={6} "
                                      L"completed={7} failed={8} total={9} discoveredRoots={10} indexedVolumes={11} indexedEntries={12} "
                                      L"storePath='{13}' activeRoot='{14}' lastGetStatusHr=0x{15:08X}.",
                                      failureMessage,
                                      outStatus.processId,
                                      static_cast<unsigned int>(outStatus.persistentStoreKind),
                                      outStatus.readyForQueryCutover,
                                      outStatus.persistentStoreInspectionSucceeded,
                                      outStatus.startupWarmupEnabled,
                                      outStatus.startupWarmupRunning,
                                      outStatus.startupWarmupCompletedRoots,
                                      outStatus.startupWarmupFailedRoots,
                                      outStatus.startupWarmupTotalRoots,
                                      outStatus.discoveredRoots.size(),
                                      outStatus.indexedVolumeCount,
                                      outStatus.indexedEntryCount,
                                      outStatus.persistentStorePath,
                                      outStatus.activeRoot,
                                      static_cast<unsigned long>(lastStatusHr)));
            return false;
        };

        SearchServiceBroker::ServiceStatus warmStatus{};
        state.Require(waitForStatus(
                          [&](const SearchServiceBroker::ServiceStatus& status) noexcept
        {
            return status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite && ! status.startupWarmupRunning &&
                   status.startupWarmupCompletedRoots == 1u && status.startupWarmupFailedRoots == 0u && status.indexedVolumeCount == 1u &&
                   status.discoveredRoots.size() == 1u &&
                   EqualsIgnoreCase(status.discoveredRoots.front(), caseRoot.wstring());
        },
                          warmStatus,
                          L"Seed stale-root SQLite service should finish warmup and mirror the overridden root."),
                      L"Seed stale-root SQLite service should finish warmup and mirror the overridden root.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(warmStatus.persistentStoreInspectionSucceeded, L"Seed stale-root SQLite service should report a successfully inspected SQLite store.");
        SqliteIndexStore::VolumeInfo seededVolumeInfo{};
        const HRESULT inspectSeedHr = SqliteIndexStore::InspectVolume(sqlitePath.wstring(), caseRoot.wstring(), seededVolumeInfo);
        state.Require(SUCCEEDED(inspectSeedHr),
                      std::format(L"Seed stale-root SQLite volume inspection failed. hr=0x{:08X}", static_cast<unsigned long>(inspectSeedHr)));
        if (FAILED(inspectSeedHr))
        {
            return false;
        }
        seedReadyForQueryCutover = seededVolumeInfo.state == SqliteIndexStore::kVolumeStateReady;
        state.Require(warmStatus.readyForQueryCutover == seedReadyForQueryCutover,
                      std::format(L"Seed warmup status readiness must match persisted volume state. statusReady={} volumeState={}.",
                                  warmStatus.readyForQueryCutover,
                                  seededVolumeInfo.state));
        state.Require(warmStatus.legacyImportVolumeCount ==
                          (seededVolumeInfo.state == SqliteIndexStore::kVolumeStateImportedLegacySnapshot ? 1u : 0u),
                      L"Seed warmup legacy-import count must remain separate from aggregate query readiness.");
    }

    auto snapshotFiles = CollectDirectoryFilesByExtension(storageRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Seed stale-root SQLite service query should not leave a compatibility snapshot runtime file behind.");

    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.txt", "beta"), L"Failed to create beta.txt after seed.");
    if (! state.failure.empty())
    {
        return false;
    }

    {
        ForegroundSearchServiceProcess service;
        std::wstring serviceError;
        state.Require(::SetEnvironmentVariableW(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS", L"5000") != 0,
                      L"Failed to delay the second startup warmup.");
        state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
        if (! state.failure.empty())
        {
            return false;
        }

        BrokerProgressRecorder recorder;
        LocalSearchIndexCore::QueryStats coldStats{};
        std::vector<LocalSearchIndexCore::Candidate> coldCandidates;
        const LocalSearchIndexCore::FallbackReason expectedColdFallbackReason =
            seedReadyForQueryCutover ? LocalSearchIndexCore::FallbackReason::StoreStale : LocalSearchIndexCore::FallbackReason::CutoverBlocked;
        HRESULT hr = SearchServiceBroker::Query(request, &BrokerProgressRecorder::ProgressThunk, &recorder, nullptr, nullptr, coldCandidates, &coldStats);
        state.Require(SUCCEEDED(hr), std::format(L"Cold-start stale-root SQLite service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        state.Require(CollectIndexedCandidateNames(coldCandidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                      L"Cold-start stale-root SQLite service query should fall back to a live scan and return fresh results.");
        state.Require(coldStats.sqliteCutoverBlocked, L"Cold-start stale-root SQLite service query should block direct SQLite while the store is stale.");
        state.Require(! coldStats.usedSqliteStore, L"Cold-start stale-root SQLite service query should not answer from stale SQLite state.");
        state.Require(coldStats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                      std::format(L"Cold-start stale-root SQLite service query should answer from the live filesystem while sqlite is stale. mode={} "
                                  L"fallbackReason={} liveScan={}",
                                  static_cast<uint32_t>(coldStats.queryExecutionMode),
                                  static_cast<uint32_t>(coldStats.fallbackReason),
                                  coldStats.usedLiveScanFallback));
        state.Require(coldStats.fallbackReason == expectedColdFallbackReason,
                      L"Cold-start stale-root SQLite service query should report the fallback implied by the persisted seed readiness.");
        state.Require(coldStats.ensureReadyDurationMs == 0u,
                      L"Cold-start stale-root SQLite service query should not spend time in EnsureReady on the request thread.");
        state.Require(! coldStats.snapshotSaved, L"Cold-start stale-root SQLite service query should not recreate a compatibility snapshot runtime store.");
        state.Require(coldStats.snapshotPath.empty(), L"Cold-start stale-root SQLite service query should not report a compatibility snapshot path.");
        state.Require(coldStats.usedLiveScanFallback, L"Cold-start stale-root SQLite live-scan path should report live-scan fallback usage.");

        const auto progress = recorder.Snapshots();
        state.Require(! progress.empty(), L"Cold-start stale-root SQLite service query should report broker progress.");
        const RecordedSearchProgress* enumerating = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING);
        state.Require(enumerating != nullptr, L"Cold-start stale-root SQLite service query should report an enumerating progress update.");
        if (enumerating != nullptr)
        {
            state.Require(enumerating->queryExecutionMode == coldStats.queryExecutionMode,
                          L"Cold-start stale-root SQLite service progress should report the same execution mode as query stats.");
            state.Require(enumerating->fallbackReason == expectedColdFallbackReason,
                          L"Cold-start stale-root SQLite service progress should report the fallback implied by the persisted seed readiness.");
            state.Require(EqualsIgnoreCase(enumerating->activeRoot, caseRoot.wstring()),
                          L"Cold-start stale-root SQLite service progress should report the searched root as active.");
            state.Require((enumerating->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u,
                          L"Cold-start stale-root SQLite live-scan path should report degraded-no-index through broker progress.");
            state.Require(enumerating->storeState == LocalSearchIndexCore::StoreState::Syncing,
                          L"Cold-start stale-root SQLite live-scan path should report a syncing store state.");
            state.Require(enumerating->syncPhase == LocalSearchIndexCore::SyncPhase::Idle,
                          L"Cold-start stale-root SQLite live-scan path should stay non-blocking while warmup is delayed.");
        }

        SearchServiceBroker::ServiceStatus status{};
        hr = SearchServiceBroker::GetStatus(status);
        state.Require(SUCCEEDED(hr),
                      std::format(L"SearchServiceBroker::GetStatus after stale-root SQLite query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                      L"Stale-root SQLite status should keep reporting the SQLite backend after a degraded query.");
        state.Require(status.queryExecutionMode == coldStats.queryExecutionMode,
                      L"Stale-root SQLite status should report the same execution mode as query stats.");
        state.Require(status.fallbackReason == expectedColdFallbackReason,
                      L"Stale-root SQLite status should report the fallback implied by the persisted seed readiness.");
        state.Require(EqualsIgnoreCase(status.activeRoot, caseRoot.wstring()), L"Stale-root SQLite status should report the searched root as active.");
        state.Require(status.storeState == LocalSearchIndexCore::StoreState::Syncing, L"Stale-root SQLite live-scan status should report a syncing store.");
        state.Require(status.syncPhase == LocalSearchIndexCore::SyncPhase::Idle,
                      L"Stale-root SQLite live-scan status should stay non-blocking while warmup is delayed.");

        std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
        if (! state.failure.empty())
        {
            return false;
        }
    }

    snapshotFiles = CollectDirectoryFilesByExtension(storageRoot, L".bin");
    state.Require(snapshotFiles.empty(), L"Cold-start stale-root SQLite service query should not create a compatibility snapshot runtime file.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_external_rotation_refreshes_without_retry",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_external_rotation", caseRoot),
                  L"Failed to prepare search_service_sqlite_external_rotation root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.txt", "beta"), L"Failed to create beta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath  = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-external-rotation.sqlite3");
    const std::wstring pendingLegacyRoot    = (caseRoot.parent_path() / L"pending-legacy-root").wstring();

    const auto addRootEntry = [&](SqliteIndexStore::ReplaceVolumeRequest& request) noexcept
    {
        request.entries.push_back({
            .fileIdLow      = 1u,
            .fileIdHigh     = 0u,
            .parentIdLow    = 0u,
            .parentIdHigh   = 0u,
            .fullPath       = caseRoot.wstring(),
            .name           = caseRoot.filename().wstring(),
            .attributes     = FILE_ATTRIBUTE_DIRECTORY,
            .sizeBytes      = 0u,
            .writeTime100ns = 0u,
        });
    };
    const auto addFileEntry = [&](SqliteIndexStore::ReplaceVolumeRequest& request, uint64_t fileId, std::wstring_view name) noexcept
    {
        request.entries.push_back({
            .fileIdLow      = fileId,
            .fileIdHigh     = 0u,
            .parentIdLow    = 1u,
            .parentIdHigh   = 0u,
            .fullPath       = (caseRoot / name).wstring(),
            .name           = std::wstring(name),
            .attributes     = FILE_ATTRIBUTE_ARCHIVE,
            .sizeBytes      = 5u,
            .writeTime100ns = 0u,
        });
    };

    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::EnsureBootstrap for external rotation test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest initialRequest{};
    initialRequest.rootPath       = caseRoot.wstring();
    initialRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    initialRequest.journalId      = 1u;
    initialRequest.nextUsn        = 1u;
    initialRequest.state          = SqliteIndexStore::kVolumeStateReady;
    addRootEntry(initialRequest);
    addFileEntry(initialRequest, 2u, L"alpha.txt");

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), initialRequest, &replaceResult);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Initial SqliteIndexStore::ReplaceVolume for external rotation test failed. hr=0x{:08X}",
                              static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest pendingRequest{};
    pendingRequest.rootPath       = pendingLegacyRoot;
    pendingRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    pendingRequest.journalId      = 1u;
    pendingRequest.nextUsn        = 1u;
    pendingRequest.state          = SqliteIndexStore::kVolumeStateImportedLegacySnapshot;
    pendingRequest.entries.push_back({
        .fileIdLow      = 1u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 0u,
        .parentIdHigh   = 0u,
        .fullPath       = pendingLegacyRoot + L"\\ghost.txt",
        .name           = L"ghost.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });

    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), pendingRequest, &replaceResult);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Pending legacy SqliteIndexStore::ReplaceVolume for external rotation test failed. hr=0x{:08X}",
                              static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    constexpr wchar_t kAllowDirectSqliteWithoutJournalEnvVar[] = L"REDSALAMANDER_TEST_ALLOW_DIRECT_SQLITE_WITHOUT_JOURNAL";
    const std::wstring previousDirectSqliteOverride            = GetEnvVarTrimmed(kAllowDirectSqliteWithoutJournalEnvVar);
    state.Require(::SetEnvironmentVariableW(kAllowDirectSqliteWithoutJournalEnvVar, L"1") != 0,
                  L"Failed to enable direct SQLite test currentness override.");
    const auto restoreDirectSqliteOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousDirectSqliteOverride.empty() ? nullptr : previousDirectSqliteOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(kAllowDirectSqliteWithoutJournalEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    LocalSearchIndexCore::QueryStats primingStats{};
    std::vector<LocalSearchIndexCore::Candidate> primingCandidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, primingCandidates, &primingStats);
    state.Require(SUCCEEDED(hr), std::format(L"External rotation priming query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(primingStats.fallbackReason == LocalSearchIndexCore::FallbackReason::CutoverBlocked,
                  L"External rotation priming query should cache the pending-legacy cutover block.");
    state.Require(primingStats.usedLiveScanFallback, L"External rotation priming query should fall back while a pending legacy volume exists.");

    SqliteIndexStore::ReplaceVolumeRequest rotatedRequest = initialRequest;
    addFileEntry(rotatedRequest, 3u, L"beta.txt");
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), rotatedRequest, &replaceResult);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Rotated SqliteIndexStore::ReplaceVolume for external rotation test failed. hr=0x{:08X}",
                              static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    hr = SqliteIndexStore::DeleteVolume(sqlitePath.wstring(), pendingLegacyRoot);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::DeleteVolume for external rotation pending legacy root failed. hr=0x{:08X}",
                              static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::StoreInfo rotatedInfo{};
    hr = SqliteIndexStore::InspectStore(sqlitePath.wstring(), rotatedInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::InspectStore after external rotation failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(rotatedInfo.legacyImportVolumeCount == 0u,
                  std::format(L"External rotation should remove pending legacy volumes, got {}.", rotatedInfo.legacyImportVolumeCount));

    const std::filesystem::path perfPath = SelfTest::GetPerfArtifactPath(L"perf_metrics.jsonl");
    std::error_code perfEc;
    const std::uintmax_t retryMetricStartOffset = std::filesystem::exists(perfPath, perfEc) ? std::filesystem::file_size(perfPath, perfEc) : 0u;

    LocalSearchIndexCore::QueryStats rotatedStats{};
    std::vector<LocalSearchIndexCore::Candidate> rotatedCandidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, rotatedCandidates, &rotatedStats);
    state.Require(SUCCEEDED(hr), std::format(L"External rotation refreshed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(rotatedCandidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"External rotation refreshed query should return the externally rotated SQLite contents.");
    state.Require(rotatedStats.usedSqliteStore, L"External rotation refreshed query should enumerate from SQLite.");
    state.Require(rotatedStats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::Sqlite,
                  L"External rotation refreshed query should report SQLite execution.");
    state.Require(rotatedStats.fallbackReason == LocalSearchIndexCore::FallbackReason::None,
                  L"External rotation refreshed query should not require a stale-store fallback.");
    state.Require(! rotatedStats.sqliteCutoverBlocked, L"External rotation refreshed query should not report a cutover block.");

    bool retryMetricSeen = false;
    std::ifstream perfFile(perfPath, std::ios::binary);
    if (perfFile)
    {
        if (retryMetricStartOffset != 0u)
        {
            perfFile.seekg(static_cast<std::streamoff>(retryMetricStartOffset), std::ios::beg);
        }

        std::string line;
        while (std::getline(perfFile, line))
        {
            if (line.find("\"metric\":\"search.backend.sqlite.retry_query_ms\"") != std::string::npos)
            {
                retryMetricSeen = true;
                break;
            }
        }
    }

    state.Require(! retryMetricSeen,
                  L"External store rotation should refresh cached PersistentStoreInfo before querying instead of using retry_query_ms.");

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_invalid_store_falls_back_live_scan",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_invalid_store", caseRoot),
                  L"Failed to prepare search_service_sqlite_invalid_store root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath  = caseRoot / L"invalid.sqlite3";
    const std::string invalidDatabase       = "not a sqlite database";
    const std::span<const char> invalidDatabaseBytes(invalidDatabase.data(), invalidDatabase.size());
    state.Require(SelfTest::WriteBinaryFile(sqlitePath, std::as_bytes(invalidDatabaseBytes)), L"Failed to create an invalid SQLite fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus status{};
    HRESULT hr = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SearchServiceBroker::GetStatus for invalid-store SQLite service failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                  L"Invalid-store SQLite service status should report the SQLite backend.");
    state.Require(EqualsIgnoreCase(status.storageRootDirectory, storageRoot.wstring()),
                  L"Invalid-store SQLite service status should report the overridden storage root.");
    state.Require(EqualsIgnoreCase(status.persistentStorePath, sqlitePath.wstring()),
                  L"Invalid-store SQLite service status should report the configured database path.");
    state.Require(status.persistentStoreBytes > 0u, L"Invalid-store SQLite service status should still report database bytes for the corrupt file.");
    state.Require(! status.persistentStoreInspectionSucceeded, L"Invalid-store SQLite service status should report failed store inspection.");
    state.Require(! status.readyForQueryCutover, L"Invalid-store SQLite service status should report readyForQueryCutover=false.");
    state.Require(status.storeState == LocalSearchIndexCore::StoreState::Invalid,
                  L"Invalid-store SQLite service status should report an invalid store state before the first query.");
    state.Require(status.syncPhase == LocalSearchIndexCore::SyncPhase::Idle, L"Invalid-store SQLite service status should remain idle before the first query.");
    state.Require(status.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::Unknown,
                  L"Invalid-store SQLite service status should not report a query execution mode before the first query.");
    state.Require(status.fallbackReason == LocalSearchIndexCore::FallbackReason::StoreInvalid,
                  L"Invalid-store SQLite service status should report a store-invalid fallback reason.");

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    BrokerProgressRecorder recorder;
    LocalSearchIndexCore::QueryStats queryStats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    hr = SearchServiceBroker::Query(request, &BrokerProgressRecorder::ProgressThunk, &recorder, nullptr, nullptr, candidates, &queryStats);
    state.Require(SUCCEEDED(hr), std::format(L"Invalid-store SQLite service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Invalid-store SQLite service query should fall back to the live filesystem and find alpha.txt.");
    state.Require(! queryStats.usedSqliteStore, L"Invalid-store SQLite service query should not use the configured SQLite store.");
    state.Require(queryStats.usedLiveScanFallback, L"Invalid-store SQLite service query should report live-scan fallback usage.");
    state.Require(queryStats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                  L"Invalid-store SQLite service query should report the live-scan fallback execution mode.");
    state.Require(queryStats.fallbackReason == LocalSearchIndexCore::FallbackReason::StoreInvalid,
                  L"Invalid-store SQLite service query should report a store-invalid fallback reason.");
    state.Require(queryStats.ensureReadyDurationMs == 0u, L"Invalid-store SQLite service query should not block in EnsureReady on the request thread.");

    const auto progress = recorder.Snapshots();
    state.Require(! progress.empty(), L"Invalid-store SQLite service query should report broker progress.");
    const RecordedSearchProgress* enumerating = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING);
    state.Require(enumerating != nullptr, L"Invalid-store SQLite service query should report an enumerating progress update.");
    if (enumerating != nullptr)
    {
        state.Require((enumerating->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u,
                      L"Invalid-store SQLite service query should report degraded-no-index through broker progress.");
        state.Require(enumerating->storeState == LocalSearchIndexCore::StoreState::Invalid,
                      L"Invalid-store SQLite service progress should report an invalid store state.");
        state.Require(enumerating->syncPhase == LocalSearchIndexCore::SyncPhase::Idle,
                      L"Invalid-store SQLite service progress should remain idle while answering from the live filesystem.");
        state.Require(enumerating->queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                      L"Invalid-store SQLite service progress should report the live-scan fallback execution mode.");
        state.Require(enumerating->fallbackReason == LocalSearchIndexCore::FallbackReason::StoreInvalid,
                      L"Invalid-store SQLite service progress should report a store-invalid fallback reason.");
        state.Require(EqualsIgnoreCase(enumerating->activeRoot, caseRoot.wstring()),
                      L"Invalid-store SQLite service progress should report the searched root as active.");
    }

    status = {};
    hr     = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SearchServiceBroker::GetStatus after invalid-store SQLite query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(status.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                  L"Invalid-store SQLite status should report the live-scan fallback execution mode after the degraded query.");
    state.Require(status.fallbackReason == LocalSearchIndexCore::FallbackReason::StoreInvalid,
                  L"Invalid-store SQLite status should report a store-invalid fallback reason after the degraded query.");
    state.Require(status.storeState == LocalSearchIndexCore::StoreState::Invalid,
                  L"Invalid-store SQLite status should continue reporting an invalid store after the degraded query.");
    state.Require(status.syncPhase == LocalSearchIndexCore::SyncPhase::Idle, L"Invalid-store SQLite status should remain idle after the degraded query.");
    state.Require(status.activeRoot.empty(),
                  L"Invalid-store SQLite status should clear generation-bound active-root state when the configured store cannot be inspected.");

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(output.contains("Store not inspected yet"),
                  L"Invalid-store SQLite foreground banner should surface that the store was not inspectable at startup.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_query_failure_falls_back_live_scan",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_query_failure", caseRoot),
                  L"Failed to prepare search_service_sqlite_query_failure root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    uint64_t currentJournalId = 0u;
    uint64_t currentNextUsn   = 0u;
    if (! TryPrepareDirectSqliteCursorForSelfTest(state, caseRoot, L"SQLite query-failure", currentJournalId, currentNextUsn))
    {
        return state.failure.empty();
    }

    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath  = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-query-failure.sqlite3");
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::EnsureBootstrap for SQLite query-failure test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest replaceRequest{};
    replaceRequest.rootPath       = caseRoot.wstring();
    replaceRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replaceRequest.journalId      = currentJournalId;
    replaceRequest.nextUsn        = currentNextUsn;
    replaceRequest.state          = SqliteIndexStore::kVolumeStateReady;
    replaceRequest.entries.push_back({
        .fileIdLow      = 2u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 1u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"alpha.txt").wstring(),
        .name           = L"alpha.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), replaceRequest, &replaceResult);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::ReplaceVolume for SQLite query-failure test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus status{};
    hr = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SearchServiceBroker::GetStatus for SQLite query-failure service failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                  L"SQLite query-failure service status should report the SQLite backend.");
    state.Require(status.persistentStoreInspectionSucceeded,
                  L"SQLite query-failure service status should report successful store inspection before the first query.");
    state.Require(status.readyForQueryCutover, L"SQLite query-failure service status should report readyForQueryCutover before the first query.");
    state.Require(status.storeState == LocalSearchIndexCore::StoreState::Ready,
                  L"SQLite query-failure service status should report a ready store before the first query.");
    state.Require(status.syncPhase == LocalSearchIndexCore::SyncPhase::Watching,
                  L"SQLite query-failure service status should report the watching sync phase before the first query.");
    state.Require(status.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::Unknown,
                  L"SQLite query-failure service status should not report a query execution mode before the first query.");
    state.Require(status.fallbackReason == LocalSearchIndexCore::FallbackReason::None,
                  L"SQLite query-failure service status should not report a fallback reason before the first query.");

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    BrokerProgressRecorder recorder;
    LocalSearchIndexCore::QueryStats queryStats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    hr = SearchServiceBroker::Query(request, &BrokerProgressRecorder::ProgressThunk, &recorder, nullptr, nullptr, candidates, &queryStats);
    state.Require(SUCCEEDED(hr), std::format(L"SQLite query-failure service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"SQLite query-failure service query should fall back to the live filesystem and find alpha.txt.");
    state.Require(! queryStats.usedSqliteStore, L"SQLite query-failure service query should not report SQLite usage after the failure.");
    state.Require(queryStats.usedLiveScanFallback, L"SQLite query-failure service query should report live-scan fallback usage.");
    state.Require(queryStats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                  L"SQLite query-failure service query should report the live-scan fallback execution mode.");
    state.Require(queryStats.fallbackReason == LocalSearchIndexCore::FallbackReason::SqliteFailure,
                  L"SQLite query-failure service query should report the sqlite-failure fallback reason.");
    state.Require(queryStats.ensureReadyDurationMs == 0u, L"SQLite query-failure service query should not block in EnsureReady on the request thread.");

    const auto progress = recorder.Snapshots();
    state.Require(! progress.empty(), L"SQLite query-failure service query should report broker progress.");
    const RecordedSearchProgress* enumerating = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING);
    state.Require(enumerating != nullptr, L"SQLite query-failure service query should report an enumerating progress update.");
    if (enumerating != nullptr)
    {
        state.Require((enumerating->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u,
                      L"SQLite query-failure service query should report degraded-no-index through broker progress.");
        state.Require(enumerating->storeState == LocalSearchIndexCore::StoreState::Recovering,
                      L"SQLite query-failure service progress should report a recovering store state.");
        state.Require(enumerating->syncPhase == LocalSearchIndexCore::SyncPhase::Idle,
                      L"SQLite query-failure service progress should remain idle while answering from the live filesystem.");
        state.Require(enumerating->queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                      L"SQLite query-failure service progress should report the live-scan fallback execution mode.");
        state.Require(enumerating->fallbackReason == LocalSearchIndexCore::FallbackReason::SqliteFailure,
                      L"SQLite query-failure service progress should report a sqlite-failure fallback reason.");
        state.Require(EqualsIgnoreCase(enumerating->activeRoot, caseRoot.wstring()),
                      L"SQLite query-failure service progress should report the searched root as active.");
    }

    status = {};
    hr     = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SearchServiceBroker::GetStatus after SQLite query-failure service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(status.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                  L"SQLite query-failure service status should report the live-scan fallback execution mode after the degraded query.");
    state.Require(status.fallbackReason == LocalSearchIndexCore::FallbackReason::SqliteFailure,
                  L"SQLite query-failure service status should report a sqlite-failure fallback reason after the degraded query.");
    state.Require(status.storeState == LocalSearchIndexCore::StoreState::Recovering,
                  L"SQLite query-failure service status should report a recovering store after the degraded query.");
    state.Require(status.syncPhase == LocalSearchIndexCore::SyncPhase::Idle,
                  L"SQLite query-failure service status should remain idle after the degraded query.");
    state.Require(EqualsIgnoreCase(status.activeRoot, caseRoot.wstring()), L"SQLite query-failure service status should report the searched root as active.");

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(output.contains("SQLite query fallback"), L"SQLite query-failure foreground output should record the SQLite fallback warning.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_midquery_failure_restarts_live_scan_without_duplicates",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_midquery_failure", caseRoot),
                  L"Failed to prepare search_service_sqlite_midquery_failure root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.txt", "beta"), L"Failed to create beta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    uint64_t currentJournalId = 0u;
    uint64_t currentNextUsn   = 0u;
    if (! TryPrepareDirectSqliteCursorForSelfTest(state, caseRoot, L"SQLite mid-query-failure", currentJournalId, currentNextUsn))
    {
        return state.failure.empty();
    }

    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath  = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-midquery-failure.sqlite3");
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::EnsureBootstrap for SQLite mid-query-failure test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest replaceRequest{};
    replaceRequest.rootPath       = caseRoot.wstring();
    replaceRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replaceRequest.journalId      = currentJournalId;
    replaceRequest.nextUsn        = currentNextUsn;
    replaceRequest.state          = SqliteIndexStore::kVolumeStateReady;
    replaceRequest.entries.push_back({
        .fileIdLow      = 1u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 0u,
        .parentIdHigh   = 0u,
        .fullPath       = caseRoot.wstring(),
        .name           = caseRoot.filename().wstring(),
        .attributes     = FILE_ATTRIBUTE_DIRECTORY,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 2u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 1u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"alpha.txt").wstring(),
        .name           = L"alpha.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 3u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 1u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"beta.txt").wstring(),
        .name           = L"beta.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), replaceRequest, &replaceResult);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::ReplaceVolume for SQLite mid-query-failure test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    constexpr wchar_t kInjectedFailAfterRowsEnvVar[] = L"REDSALAMANDER_TEST_SQLITE_FAIL_AFTER_EMITTED_ROWS";
    const std::wstring previousInjectedFailAfterRows = GetEnvVarTrimmed(kInjectedFailAfterRowsEnvVar);
    state.Require(::SetEnvironmentVariableW(kInjectedFailAfterRowsEnvVar, L"1") != 0, L"Failed to enable the injected SQLite mid-query failure.");
    const auto restoreInjectedFailAfterRows = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousInjectedFailAfterRows.empty() ? nullptr : previousInjectedFailAfterRows.c_str();
        static_cast<void>(::SetEnvironmentVariableW(kInjectedFailAfterRowsEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    BrokerProgressRecorder recorder;
    LocalSearchIndexCore::QueryStats queryStats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    hr = SearchServiceBroker::Query(request, &BrokerProgressRecorder::ProgressThunk, &recorder, nullptr, nullptr, candidates, &queryStats);
    state.Require(SUCCEEDED(hr), std::format(L"SQLite mid-query-failure service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(candidates.size() == 2u,
                  std::format(L"SQLite mid-query-failure service query should return exactly 2 unique candidates, got {}.", candidates.size()));
    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"SQLite mid-query-failure service query should restart as a live scan without duplicating early SQLite hits.");
    state.Require(queryStats.usedSqliteStore, L"SQLite mid-query-failure service query should report partial SQLite usage before the restart.");
    state.Require(queryStats.usedLiveScanFallback, L"SQLite mid-query-failure service query should report live-scan fallback usage.");
    state.Require(queryStats.sqliteReadOnlyQuery, L"SQLite mid-query-failure service query should report the read-only SQLite query attempt.");
    state.Require(queryStats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                  L"SQLite mid-query-failure service query should report the live-scan fallback execution mode.");
    state.Require(queryStats.fallbackReason == LocalSearchIndexCore::FallbackReason::SqliteFailure,
                  L"SQLite mid-query-failure service query should report the sqlite-failure fallback reason.");

    const auto progress = recorder.Snapshots();
    state.Require(! progress.empty(), L"SQLite mid-query-failure service query should report broker progress.");
    const RecordedSearchProgress* finalProgress = progress.empty() ? nullptr : &progress.back();
    state.Require(finalProgress != nullptr, L"SQLite mid-query-failure service query should report a final broker progress snapshot.");
    if (finalProgress != nullptr)
    {
        state.Require(finalProgress->candidateFiles == 2u,
                      std::format(L"SQLite mid-query-failure service progress should report 2 unique candidates, got {}.", finalProgress->candidateFiles));
        state.Require((finalProgress->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u,
                      L"SQLite mid-query-failure service progress should report degraded-no-index.");
        state.Require(finalProgress->queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                      L"SQLite mid-query-failure service progress should report the live-scan fallback execution mode.");
        state.Require(finalProgress->fallbackReason == LocalSearchIndexCore::FallbackReason::SqliteFailure,
                      L"SQLite mid-query-failure service progress should report a sqlite-failure fallback reason.");
    }

    SearchServiceBroker::ServiceStatus status{};
    hr = SearchServiceBroker::GetStatus(status);
    state.Require(
        SUCCEEDED(hr),
        std::format(L"SearchServiceBroker::GetStatus after SQLite mid-query-failure service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(status.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                  L"SQLite mid-query-failure service status should report the live-scan fallback execution mode after the degraded query.");
    state.Require(status.fallbackReason == LocalSearchIndexCore::FallbackReason::SqliteFailure,
                  L"SQLite mid-query-failure service status should report a sqlite-failure fallback reason after the degraded query.");
    state.Require(status.storeState == LocalSearchIndexCore::StoreState::Recovering,
                  L"SQLite mid-query-failure service status should report a recovering store after the degraded query.");

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(output.contains("injected enumerate failure"),
                  L"SQLite mid-query-failure foreground output should record the injected SQLite query failure.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_prefilter_roundtrip",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_prefilter", caseRoot), L"Failed to prepare search_service_sqlite_prefilter root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alphabet.txt", "alphabet"), L"Failed to create alphabet.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.log", "beta"), L"Failed to create beta.log.");

    uint64_t currentJournalId = 0u;
    uint64_t currentNextUsn   = 0u;
    if (! TryPrepareDirectSqliteCursorForSelfTest(state, caseRoot, L"SQLite service prefilter", currentJournalId, currentNextUsn))
    {
        return state.failure.empty();
    }

    const std::filesystem::path sqlitePath  = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-prefilter.sqlite3");
    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::EnsureBootstrap for SQLite service prefilter test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest replaceRequest{};
    replaceRequest.rootPath       = caseRoot.wstring();
    replaceRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replaceRequest.journalId      = currentJournalId;
    replaceRequest.nextUsn        = currentNextUsn;
    replaceRequest.state          = SqliteIndexStore::kVolumeStateReady;
    replaceRequest.entries.push_back({
        .fileIdLow      = 1u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 0u,
        .parentIdHigh   = 0u,
        .fullPath       = caseRoot.wstring(),
        .name           = caseRoot.filename().wstring(),
        .attributes     = FILE_ATTRIBUTE_DIRECTORY,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 2u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 1u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"alpha.txt").wstring(),
        .name           = L"alpha.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 5u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 3u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 1u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"alphabet.txt").wstring(),
        .name           = L"alphabet.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 8u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 4u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 1u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"beta.log").wstring(),
        .name           = L"beta.log",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 4u,
        .writeTime100ns = 0u,
    });

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), replaceRequest, &replaceResult);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::ReplaceVolume for SQLite service prefilter test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::QueryRequest seedRequest{};
    seedRequest.rootPath           = caseRoot.wstring();
    seedRequest.namePattern        = L"*";
    seedRequest.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    seedRequest.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    seedRequest.recursive          = true;
    seedRequest.includeFiles       = true;
    seedRequest.includeDirectories = false;

    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    hr = SearchServiceBroker::Query(seedRequest, nullptr, nullptr, nullptr, nullptr, seedCandidates, &seedStats);
    state.Require(SUCCEEDED(hr), std::format(L"Seed SQLite service prefilter query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(seedStats.usedSqliteStore, L"Seed SQLite service prefilter query should enumerate from the seeded SQLite fixture.");

    SearchServiceBroker::QueryRequest prefixRequest = seedRequest;
    prefixRequest.namePattern                       = L"alph*";

    LocalSearchIndexCore::QueryStats prefixStats{};
    std::vector<LocalSearchIndexCore::Candidate> prefixCandidates;
    hr = SearchServiceBroker::Query(prefixRequest, nullptr, nullptr, nullptr, nullptr, prefixCandidates, &prefixStats);
    state.Require(SUCCEEDED(hr), std::format(L"Prefix SQLite service prefilter query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(prefixStats.usedSqliteStore, L"Warm SQLite service prefix query should enumerate from SQLite.");
    state.Require(prefixStats.usedNamePrefilter, L"Warm SQLite service prefix query should report a broker-roundtripped name prefilter.");
    state.Require(CollectIndexedCandidateNames(prefixCandidates) == std::vector<std::wstring>{L"alpha.txt", L"alphabet.txt"},
                  L"Warm SQLite service prefix query returned an unexpected candidate set.");

    SearchServiceBroker::QueryRequest regexRequest = seedRequest;
    regexRequest.namePattern                       = L"^alpha.*\\.txt$";
    regexRequest.nameMode                          = FILESYSTEM_SEARCH_NAME_REGEX;

    LocalSearchIndexCore::QueryStats regexStats{};
    std::vector<LocalSearchIndexCore::Candidate> regexCandidates;
    hr = SearchServiceBroker::Query(regexRequest, nullptr, nullptr, nullptr, nullptr, regexCandidates, &regexStats);
    state.Require(SUCCEEDED(hr), std::format(L"Regex SQLite service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(regexStats.usedSqliteStore, L"Warm SQLite service regex query should still enumerate from SQLite.");
    state.Require(! regexStats.usedNamePrefilter, L"Warm SQLite service regex query should report prefilter=false so regex stays in the C++ matcher.");
    state.Require(CollectIndexedCandidateNames(regexCandidates) == std::vector<std::wstring>{L"alpha.txt", L"alphabet.txt"},
                  L"Warm SQLite service regex query returned an unexpected candidate set.");

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_status_reports_pending_legacy_import",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_pending_legacy_import", caseRoot),
                  L"Failed to prepare search_service_sqlite_pending_legacy_import root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");

    const std::filesystem::path sqlitePath = caseRoot / L"pending.sqlite3";
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::EnsureBootstrap for pending-import test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest request{};
    request.rootPath       = caseRoot.wstring();
    request.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    request.journalId      = 1u;
    request.nextUsn        = 1u;
    request.state          = SqliteIndexStore::kVolumeStateImportedLegacySnapshot;
    request.entries.push_back({
        .fileIdLow      = 1u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 0u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"alpha.txt").wstring(),
        .name           = L"alpha.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), request, &replaceResult);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::ReplaceVolume for pending-import test failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus status{};
    hr = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SearchServiceBroker::GetStatus for pending legacy-import SQLite service failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                  L"Pending legacy-import service status should report the SQLite backend.");
    state.Require(status.persistentStoreInspectionSucceeded, L"Pending legacy-import service status should report successful store inspection.");
    state.Require(! status.readyForQueryCutover, L"Pending legacy-import service status should report readyForQueryCutover=false.");
    state.Require(status.indexedVolumeCount == 1u,
                  std::format(L"Expected one mirrored volume in pending legacy-import status, got {}.", status.indexedVolumeCount));
    state.Require(status.indexedEntryCount == 1u,
                  std::format(L"Expected one mirrored entry in pending legacy-import status, got {}.", status.indexedEntryCount));
    state.Require(status.legacyImportVolumeCount == 1u,
                  std::format(L"Expected one legacy-import volume in pending legacy-import status, got {}.", status.legacyImportVolumeCount));
    state.Require(status.storeState == LocalSearchIndexCore::StoreState::Syncing, L"Pending legacy-import service status should report a syncing store state.");
    state.Require(status.syncPhase == LocalSearchIndexCore::SyncPhase::Idle,
                  L"Pending legacy-import service status should stay idle until a query or rebuild starts.");
    state.Require(status.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::Unknown,
                  L"Pending legacy-import service status should not report a query execution mode before the first query.");
    state.Require(status.fallbackReason == LocalSearchIndexCore::FallbackReason::CutoverBlocked,
                  L"Pending legacy-import service status should report a cutover-blocked fallback reason.");

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(output.contains("Readiness"), L"Pending legacy-import foreground output should include the readiness field.");
    state.Require(output.contains("Pending backfill/rebuild"), L"Pending legacy-import foreground output should report pending backfill/rebuild.");
    state.Require(output.contains("legacy-imports=1"), L"Pending legacy-import foreground output should report one legacy-import volume.");

    request.state = SqliteIndexStore::kVolumeStateCurrentnessUnproven;
    hr            = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), request, &replaceResult);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::ReplaceVolume for currentness-unproven status failed. hr=0x{:08X}",
                              static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::filesystem::path readyRoot = caseRoot / L"ready-root";
    state.Require(SelfTest::EnsureDirectory(readyRoot), L"Failed to create mixed-state ready root.");
    state.Require(SelfTest::WriteTextFile(readyRoot / L"beta.txt", "beta"), L"Failed to create mixed-state ready-root file.");
    if (! state.failure.empty())
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest readyRequest{};
    readyRequest.rootPath       = readyRoot.wstring();
    readyRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    readyRequest.journalId      = 2u;
    readyRequest.nextUsn        = 2u;
    readyRequest.state          = SqliteIndexStore::kVolumeStateReady;
    readyRequest.entries.push_back({
        .fileIdLow      = 2u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 0u,
        .parentIdHigh   = 0u,
        .fullPath       = (readyRoot / L"beta.txt").wstring(),
        .name           = L"beta.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), readyRequest, &replaceResult);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::ReplaceVolume for mixed ready status failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::StoreInfo currentnessInfo{};
    hr = SqliteIndexStore::InspectStore(sqlitePath.wstring(), currentnessInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::InspectStore for currentness-unproven status failed. hr=0x{:08X}",
                              static_cast<unsigned long>(hr)));
    state.Require(currentnessInfo.legacyImportVolumeCount == 0u,
                  L"Currentness-unproven volumes must not be counted as legacy-import migration work.");
    state.Require(currentnessInfo.queryUnreadyVolumeCount == 1u,
                  L"A mixed Ready + CurrentnessUnproven store must remain query-unready until currentness is proved.");
    state.Require(currentnessInfo.volumeCount == 2u, L"Mixed-state readiness fixture should contain both Ready and CurrentnessUnproven volumes.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    status = {};
    hr     = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SearchServiceBroker::GetStatus for currentness-unproven SQLite service failed. hr=0x{:08X}",
                              static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(EqualsIgnoreCase(status.pipeName, pipeName), L"Currentness-unproven status must come from the intended isolated pipe.");
    state.Require(status.persistentStoreInspectionSucceeded, L"Currentness-unproven status should report successful store inspection.");
    state.Require(! status.readyForQueryCutover, L"Currentness-unproven status must not report query cutover readiness.");
    state.Require(status.legacyImportVolumeCount == 0u,
                  L"Currentness-unproven status must keep legacy migration state separate from query readiness.");
    state.Require(status.indexedVolumeCount == 2u, L"Currentness-unproven status should retain the mixed Ready volume in its store summary.");
    state.Require(status.storeState == LocalSearchIndexCore::StoreState::Syncing,
                  L"Currentness-unproven status should report a syncing store before the first query.");
    state.Require(status.syncPhase == LocalSearchIndexCore::SyncPhase::Idle,
                  L"Currentness-unproven status should remain idle until a query or rebuild proves currentness.");
    state.Require(status.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::Unknown,
                  L"Currentness-unproven status should not report a query execution mode before the first query.");
    state.Require(status.fallbackReason == LocalSearchIndexCore::FallbackReason::CutoverBlocked,
                  L"Currentness-unproven status should report a cutover-blocked fallback reason.");

    SearchServiceBroker::QueryRequest liveRequest{};
    liveRequest.rootPath           = caseRoot.wstring();
    liveRequest.namePattern        = L"*.txt";
    liveRequest.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    liveRequest.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    liveRequest.recursive          = true;
    liveRequest.includeFiles       = true;
    liveRequest.includeDirectories = false;

    LocalSearchIndexCore::QueryStats liveStats{};
    std::vector<LocalSearchIndexCore::Candidate> liveCandidates;
    hr = SearchServiceBroker::Query(liveRequest, nullptr, nullptr, nullptr, nullptr, liveCandidates, &liveStats);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Currentness-unproven mixed-store query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(liveStats.usedLiveScanFallback, L"Currentness-unproven mixed store should query through live-scan fallback.");
    if (! state.failure.empty())
    {
        return false;
    }

    request.state = SqliteIndexStore::kVolumeStateReady;
    hr            = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), request, &replaceResult);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::ReplaceVolume for externally refreshed ready status failed. hr=0x{:08X}",
                              static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus refreshedStatus{};
    hr = SearchServiceBroker::GetStatus(refreshedStatus);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SearchServiceBroker::GetStatus after external ready rotation failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(refreshedStatus.readyForQueryCutover,
                  L"Fresh inspected Ready volumes must supersede the previous cutover-blocked runtime snapshot.");
    state.Require(refreshedStatus.storeState == LocalSearchIndexCore::StoreState::Ready,
                  L"External ready rotation should replace stale syncing runtime state with Ready.");
    state.Require(refreshedStatus.syncPhase == LocalSearchIndexCore::SyncPhase::Watching,
                  L"External ready rotation should replace stale idle runtime sync state with Watching.");
    state.Require(refreshedStatus.fallbackReason == LocalSearchIndexCore::FallbackReason::None,
                  L"External ready rotation should clear the stale cutover-blocked fallback reason.");
    state.Require(refreshedStatus.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::Unknown,
                  L"External ready rotation should clear the previous generation's query execution mode.");
    state.Require(refreshedStatus.indexedVolumeCount == 2u, L"External ready rotation should preserve both indexed volumes.");

    std::string currentnessOutput;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(2'000)), currentnessOutput, serviceError), serviceError);
    state.Require(currentnessOutput.contains("Currentness unproven"),
                  L"Currentness-unproven foreground output should distinguish currentness proof from legacy backfill work.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_binary_uses_console_subsystem",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::filesystem::path servicePath = GetSiblingExecutablePath(L"RedSalamanderSearchService.exe");
    std::error_code existsEc;
    state.Require(! servicePath.empty() && std::filesystem::exists(servicePath, existsEc),
                  std::format(L"Service executable not found: {}", servicePath.wstring()));
    if (! state.failure.empty())
    {
        return false;
    }

    WORD subsystem = 0u;
    std::wstring readError;
    state.Require(TryReadPortableExecutableSubsystem(servicePath.wstring(), subsystem, readError), readError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(subsystem == IMAGE_SUBSYSTEM_WINDOWS_CUI, std::format(L"Expected the service binary to use the console subsystem, got {}.", subsystem));
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_default_identity_matches_build",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring serviceName = SearchServiceBroker::kServiceName;
    const std::wstring pipeName    = SearchServiceBroker::GetDefaultPipeName();
    const std::wstring programData = GetEnvVarTrimmed(L"ProgramData");
    state.Require(! programData.empty(), L"ProgramData should be available when validating the default search-service identity.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring storageRoot = SearchServiceBroker::GetProgramDataSearchIndexRoot();

#ifdef _DEBUG
    state.Require(serviceName == L"RedSalamanderSearchService.Debug", std::format(L"Unexpected Debug service name '{}'.", serviceName));
    state.Require(pipeName == LR"(\\.\pipe\RedSalamander.SearchService.Debug.v3)", std::format(L"Unexpected Debug pipe name '{}'.", pipeName));
    const std::wstring expectedStorageRoot = (std::filesystem::path(programData) / L"RedSalamander" / L"SearchIndex.Debug").wstring();
    state.Require(EqualsIgnoreCase(storageRoot, expectedStorageRoot), std::format(L"Unexpected Debug storage root '{}'.", storageRoot));
#else
    state.Require(serviceName == L"RedSalamanderSearchService", std::format(L"Unexpected Release service name '{}'.", serviceName));
    state.Require(pipeName == LR"(\\.\pipe\RedSalamander.SearchService.v3)", std::format(L"Unexpected Release pipe name '{}'.", pipeName));
    const std::wstring expectedStorageRoot = (std::filesystem::path(programData) / L"RedSalamander" / L"SearchIndex").wstring();
    state.Require(EqualsIgnoreCase(storageRoot, expectedStorageRoot), std::format(L"Unexpected Release storage root '{}'.", storageRoot));
#endif

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_default_store_uses_build_specific_programdata_root",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_default_store_pd", caseRoot),
                  L"Failed to prepare search_service_sqlite_default_store_uses_build_specific_programdata_root root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path programDataRoot = caseRoot / L"program-data";
    state.Require(SelfTest::EnsureDirectory(programDataRoot), std::format(L"Failed to create '{}'.", programDataRoot.wstring()));
    const std::filesystem::path discoveredRoot = caseRoot / L"discovered-root";
    state.Require(SelfTest::EnsureDirectory(discoveredRoot), std::format(L"Failed to create '{}'.", discoveredRoot.wstring()));
    state.Require(SelfTest::WriteTextFile(discoveredRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for default-store startup discovery.");
    if (! state.failure.empty())
    {
        return false;
    }

#ifdef _DEBUG
    const std::filesystem::path expectedStorageRoot = programDataRoot / L"RedSalamander" / L"SearchIndex.Debug";
    const std::filesystem::path siblingStorageRoot  = programDataRoot / L"RedSalamander" / L"SearchIndex";
#else
    const std::filesystem::path expectedStorageRoot = programDataRoot / L"RedSalamander" / L"SearchIndex";
    const std::filesystem::path siblingStorageRoot  = programDataRoot / L"RedSalamander" / L"SearchIndex.Debug";
#endif

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousRootOverride = GetEnvVarTrimmed(SearchServiceBroker::kDiscoverRootsEnvVar);
    const std::wstring previousProgramData  = GetEnvVarTrimmed(L"ProgramData");
    const std::wstring pipeName             = MakeUniquePipeName();

    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, discoveredRoot.c_str()) != 0,
                  L"Failed to override the search service startup roots.");
    state.Require(::SetEnvironmentVariableW(L"ProgramData", programDataRoot.c_str()) != 0,
                  L"Failed to override ProgramData for the split-store compatibility test.");
    const auto restoreProgramData  = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousProgramData.empty() ? nullptr : previousProgramData.c_str();
        static_cast<void>(::SetEnvironmentVariableW(L"ProgramData", restoreValue));
    });
    const auto restoreRootOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousRootOverride.empty() ? nullptr : previousRootOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, restoreValue));
    });
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });
    if (! state.failure.empty())
    {
        return false;
    }

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, L"--store-backend=sqlite"), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus status{};
    if (! WaitForSearchServiceStatusWithProcessDiagnostics(state, status, pipeName, L"default-store SQLite service", service))
    {
        return false;
    }

    state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                  L"Default-store SQLite service status should report the SQLite backend.");
    state.Require(EqualsIgnoreCase(status.storageRootDirectory, expectedStorageRoot.wstring()),
                  std::format(L"Expected storage root '{}', got '{}'.", expectedStorageRoot.wstring(), status.storageRootDirectory));
    state.Require(status.persistentStoreInspectionSucceeded, L"Default-store SQLite service should inspect the active store during startup.");
    state.Require(! status.persistentStorePath.empty(), L"Default-store SQLite service should report a database path.");
    if (status.persistentStorePath.empty())
    {
        return false;
    }

    const std::filesystem::path reportedStorePath = status.persistentStorePath;
    state.Require(EqualsIgnoreCase(reportedStorePath.parent_path().wstring(), expectedStorageRoot.wstring()),
                  std::format(L"Expected SQLite database path '{}' to stay under '{}'.", reportedStorePath.wstring(), expectedStorageRoot.wstring()));
    state.Require(status.persistentStoreBytes > 0u, L"Default-store SQLite startup should create a non-empty database file.");

    std::error_code existsEc;
    const bool expectedRootExists = std::filesystem::exists(expectedStorageRoot, existsEc);
    state.Require(! existsEc && expectedRootExists,
                  std::format(L"Expected active storage root '{}' to exist after startup. error={}", expectedStorageRoot.wstring(), existsEc.value()));
    existsEc.clear();
    const bool storeFileExists = std::filesystem::exists(reportedStorePath, existsEc);
    state.Require(! existsEc && storeFileExists,
                  std::format(L"Expected SQLite database '{}' to exist after startup. error={}", reportedStorePath.wstring(), existsEc.value()));
    existsEc.clear();
    const bool siblingRootExists = std::filesystem::exists(siblingStorageRoot, existsEc);
    state.Require(! existsEc && ! siblingRootExists,
                  std::format(L"Sibling build storage root '{}' should remain untouched. error={}", siblingStorageRoot.wstring(), existsEc.value()));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_seeded_default_store_reuses_build_specific_programdata_root",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_seeded_store_pd", caseRoot),
                  L"Failed to prepare search_service_sqlite_seeded_default_store_reuses_build_specific_programdata_root root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path programDataRoot = caseRoot / L"program-data";
    state.Require(SelfTest::EnsureDirectory(programDataRoot), std::format(L"Failed to create '{}'.", programDataRoot.wstring()));
    const std::filesystem::path discoveredRoot = caseRoot / L"discovered-root";
    state.Require(SelfTest::EnsureDirectory(discoveredRoot), std::format(L"Failed to create '{}'.", discoveredRoot.wstring()));
    state.Require(SelfTest::WriteTextFile(discoveredRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for seeded default-store startup discovery.");
    if (! state.failure.empty())
    {
        return false;
    }

#ifdef _DEBUG
    const std::filesystem::path expectedStorageRoot = programDataRoot / L"RedSalamander" / L"SearchIndex.Debug";
    const std::filesystem::path siblingStorageRoot  = programDataRoot / L"RedSalamander" / L"SearchIndex";
#else
    const std::filesystem::path expectedStorageRoot = programDataRoot / L"RedSalamander" / L"SearchIndex";
    const std::filesystem::path siblingStorageRoot  = programDataRoot / L"RedSalamander" / L"SearchIndex.Debug";
#endif
    const std::filesystem::path sqlitePath = expectedStorageRoot / L"index-v2.sqlite3";

    state.Require(SelfTest::EnsureDirectory(expectedStorageRoot), std::format(L"Failed to create expected storage root '{}'.", expectedStorageRoot.wstring()));
    if (! state.failure.empty())
    {
        return false;
    }

    SqliteIndexStore::StoreInfo bootstrapInfo{};
    const HRESULT bootstrapHr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(bootstrapHr),
                  std::format(L"Failed to seed default SQLite store '{}'. hr=0x{:08X}", sqlitePath.wstring(), static_cast<unsigned long>(bootstrapHr)));
    if (FAILED(bootstrapHr))
    {
        return false;
    }

    state.Require(EqualsIgnoreCase(bootstrapInfo.databasePath, sqlitePath.wstring()),
                  std::format(L"Seeded SQLite store should stay at '{}', got '{}'.", sqlitePath.wstring(), bootstrapInfo.databasePath));
    state.Require(bootstrapInfo.databaseBytes > 0u, L"Seeded default SQLite store should create a non-empty database file.");

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousRootOverride = GetEnvVarTrimmed(SearchServiceBroker::kDiscoverRootsEnvVar);
    const std::wstring previousProgramData  = GetEnvVarTrimmed(L"ProgramData");
    const std::wstring pipeName             = MakeUniquePipeName();

    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, discoveredRoot.c_str()) != 0,
                  L"Failed to override the search service startup roots.");
    state.Require(::SetEnvironmentVariableW(L"ProgramData", programDataRoot.c_str()) != 0,
                  L"Failed to override ProgramData for the seeded split-store compatibility test.");
    const auto restoreProgramData  = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousProgramData.empty() ? nullptr : previousProgramData.c_str();
        static_cast<void>(::SetEnvironmentVariableW(L"ProgramData", restoreValue));
    });
    const auto restoreRootOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousRootOverride.empty() ? nullptr : previousRootOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, restoreValue));
    });
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });
    if (! state.failure.empty())
    {
        return false;
    }

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, L"--store-backend=sqlite"), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus status{};
    if (! WaitForSearchServiceStatusWithProcessDiagnostics(state, status, pipeName, L"seeded default-store SQLite service", service))
    {
        return false;
    }

    state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                  L"Seeded default-store SQLite service status should report the SQLite backend.");
    state.Require(EqualsIgnoreCase(status.storageRootDirectory, expectedStorageRoot.wstring()),
                  std::format(L"Expected seeded storage root '{}', got '{}'.", expectedStorageRoot.wstring(), status.storageRootDirectory));
    state.Require(status.persistentStoreInspectionSucceeded, L"Seeded default-store SQLite service should inspect the seeded store during startup.");
    state.Require(EqualsIgnoreCase(status.persistentStorePath, sqlitePath.wstring()),
                  std::format(L"Expected seeded SQLite database path '{}', got '{}'.", sqlitePath.wstring(), status.persistentStorePath));
    state.Require(
        status.persistentStoreBytes >= bootstrapInfo.databaseBytes,
        std::format(L"Expected reported database size {} to be at least the seeded size {}.", status.persistentStoreBytes, bootstrapInfo.databaseBytes));

    std::error_code existsEc;
    const bool expectedRootExists = std::filesystem::exists(expectedStorageRoot, existsEc);
    state.Require(! existsEc && expectedRootExists,
                  std::format(L"Expected seeded storage root '{}' to exist after startup. error={}", expectedStorageRoot.wstring(), existsEc.value()));
    existsEc.clear();
    const bool storeFileExists = std::filesystem::exists(sqlitePath, existsEc);
    state.Require(! existsEc && storeFileExists,
                  std::format(L"Expected seeded SQLite database '{}' to exist after startup. error={}", sqlitePath.wstring(), existsEc.value()));
    existsEc.clear();
    const bool siblingRootExists = std::filesystem::exists(siblingStorageRoot, existsEc);
    state.Require(! existsEc && ! siblingRootExists,
                  std::format(L"Sibling build storage root '{}' should remain untouched when reusing a seeded default store. error={}",
                              siblingStorageRoot.wstring(),
                              existsEc.value()));

    SqliteIndexStore::StoreInfo seededInfo{};
    const HRESULT inspectHr = SqliteIndexStore::InspectStore(sqlitePath.wstring(), seededInfo);
    state.Require(
        SUCCEEDED(inspectHr),
        std::format(L"Failed to inspect seeded SQLite store '{}' after startup. hr=0x{:08X}", sqlitePath.wstring(), static_cast<unsigned long>(inspectHr)));
    if (FAILED(inspectHr))
    {
        return false;
    }

    state.Require(seededInfo.schemaReady, L"Seeded default SQLite store should remain schema-ready after service startup.");
    state.Require(EqualsIgnoreCase(seededInfo.databasePath, sqlitePath.wstring()), L"Service startup should reuse the seeded default SQLite path in place.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_discovers_fixed_local_roots_on_start",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::vector<std::wstring> expectedRoots = LocalSearchIndexCore::DiscoverFixedLocalRoots();
    state.Require(! expectedRoots.empty(), L"Expected at least one fixed local root to be discoverable on this machine.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_discovers_fixed_local_roots", caseRoot),
                  L"Failed to prepare search_service_discovers_fixed_local_roots root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath  = caseRoot / L"discovery.sqlite3";
    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, false, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus status{};
    if (! WaitForSearchServiceStatus(state, status, pipeName, L"root-discovery service"))
    {
        return false;
    }

    state.Require(status.discoveredRootCount == expectedRoots.size(),
                  std::format(L"Expected {} discovered fixed roots, got {}.", expectedRoots.size(), status.discoveredRootCount));
    state.Require(status.discoveredRoots == expectedRoots, L"Service status should report the same discovered fixed-root set as LocalSearchIndexCore.");
    state.Require(status.indexedVolumeCount == 0u,
                  std::format(L"Discovery-only startup should not mirror SQLite volumes yet, got {} indexed volumes.", status.indexedVolumeCount));
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_startup_warms_overridden_roots",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_startup_warms_overridden_roots", caseRoot),
                  L"Failed to prepare search_service_sqlite_startup_warms_overridden_roots root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path rootA = caseRoot / L"root-a";
    const std::filesystem::path rootB = caseRoot / L"root-b";
    std::error_code createEc;
    std::filesystem::create_directories(rootA, createEc);
    state.Require(! createEc, std::format(L"Failed to create '{}'. error={}", rootA.wstring(), createEc.value()));
    createEc.clear();
    std::filesystem::create_directories(rootB, createEc);
    state.Require(! createEc, std::format(L"Failed to create '{}'. error={}", rootB.wstring(), createEc.value()));
    state.Require(SelfTest::WriteTextFile(rootA / L"alpha.txt", "alpha"), L"Failed to create root-a alpha.txt.");
    state.Require(SelfTest::WriteTextFile(rootB / L"beta.txt", "beta"), L"Failed to create root-b beta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::Repository expectedRepository;
    LocalSearchIndexCore::SupportInfo supportA{};
    LocalSearchIndexCore::SupportInfo supportB{};
    HRESULT hr = expectedRepository.ProbePath(rootA.wstring(), supportA);
    state.Require(SUCCEEDED(hr), std::format(L"ProbePath failed for root-a. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    hr = expectedRepository.ProbePath(rootB.wstring(), supportB);
    state.Require(SUCCEEDED(hr), std::format(L"ProbePath failed for root-b. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (! state.failure.empty())
    {
        return false;
    }

    std::vector<std::wstring> expectedRoots{supportA.normalizedRootPath, supportB.normalizedRootPath};
    std::sort(expectedRoots.begin(), expectedRoots.end(), [](const std::wstring& left, const std::wstring& right) noexcept {
        return OrdinalString::LessNoCase(left, right);
    });

    const std::filesystem::path storageRoot        = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath         = caseRoot / L"startup-index.sqlite3";
    const std::wstring previousPipeOverride        = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousRootOverride        = GetEnvVarTrimmed(SearchServiceBroker::kDiscoverRootsEnvVar);
    const std::wstring previousWarmupDelayOverride = GetEnvVarTrimmed(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS");
    const std::wstring pipeName                    = MakeUniquePipeName();
    const std::wstring rootOverride                = std::format(L"{};{}", rootA.wstring(), rootB.wstring());
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, rootOverride.c_str()) != 0,
                  L"Failed to override the search service startup roots.");
    state.Require(::SetEnvironmentVariableW(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS", L"250") != 0,
                  L"Failed to override the startup warmup delay.");
    const auto restoreOverrides = wil::scope_exit([&] noexcept
    {
        const wchar_t* restorePipeValue  = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        const wchar_t* restoreRootsValue = previousRootOverride.empty() ? nullptr : previousRootOverride.c_str();
        const wchar_t* restoreDelayValue = previousWarmupDelayOverride.empty() ? nullptr : previousWarmupDelayOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restorePipeValue));
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, restoreRootsValue));
        static_cast<void>(::SetEnvironmentVariableW(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS", restoreDelayValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus status{};
    hr = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Initial SearchServiceBroker::GetStatus for startup warmup service failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(status.discoveredRoots == expectedRoots, L"Startup warmup status should report the overridden discovered-root set while warmup is running.");
    state.Require(status.startupWarmupEnabled, L"Startup warmup status should report the warmup path as enabled.");
    state.Require(status.startupWarmupRunning, L"Startup warmup status should report a running warmup immediately after foreground startup.");
    state.Require(status.startupWarmupTotalRoots == expectedRoots.size(),
                  std::format(L"Expected {} startup warmup roots, got {}.", expectedRoots.size(), status.startupWarmupTotalRoots));
    state.Require(status.startupWarmupCompletedRoots < expectedRoots.size(),
                  std::format(L"Startup warmup should not report all roots completed yet. completed={} total={}.",
                              status.startupWarmupCompletedRoots,
                              expectedRoots.size()));
    state.Require(! status.startupWarmupCurrentRoot.empty(), L"Startup warmup status should report the current root while warmup is running.");
    state.Require(status.indexedVolumeCount < expectedRoots.size(),
                  std::format(L"Startup warmup should not have mirrored all roots yet. indexedVolumes={} expectedLessThan={}.",
                              status.indexedVolumeCount,
                              expectedRoots.size()));

    std::this_thread::sleep_for(std::chrono::milliseconds(800));

    status = {};
    hr     = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Follow-up SearchServiceBroker::GetStatus for startup warmup service failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(! status.startupWarmupRunning, L"Startup warmup status should report completion after the warmup window elapses.");
    state.Require(status.startupWarmupCompletedRoots == expectedRoots.size(),
                  std::format(L"Expected startup warmup to complete {} roots, got {}.", expectedRoots.size(), status.startupWarmupCompletedRoots));
    state.Require(status.startupWarmupFailedRoots == 0u,
                  std::format(L"Expected startup warmup to finish without failures, got {} failed roots.", status.startupWarmupFailedRoots));
    state.Require(status.startupWarmupCurrentRoot.empty(), L"Completed startup warmup should clear the current-root status field.");
    state.Require(
        status.indexedVolumeCount == expectedRoots.size(),
        std::format(L"Startup warmup did not mirror both overridden roots. indexedVolumes={} expected={}.", status.indexedVolumeCount, expectedRoots.size()));
    state.Require(status.legacyImportVolumeCount == 0u,
                  std::format(L"Startup warmup should not leave legacy-import roots behind. legacyImports={}.", status.legacyImportVolumeCount));
    state.Require(
        status.indexedEntryCount >= expectedRoots.size() * 2u,
        std::format(L"Startup warmup mirrored too few entries. indexedEntries={} expectedAtLeast={}.", status.indexedEntryCount, expectedRoots.size() * 2u));
    if (! state.failure.empty())
    {
        return false;
    }

    SqliteIndexStore::VolumeInfo volumeInfoA{};
    SqliteIndexStore::VolumeInfo volumeInfoB{};
    const HRESULT inspectVolumeAHr = SqliteIndexStore::InspectVolume(sqlitePath.wstring(), supportA.normalizedRootPath, volumeInfoA);
    const HRESULT inspectVolumeBHr = SqliteIndexStore::InspectVolume(sqlitePath.wstring(), supportB.normalizedRootPath, volumeInfoB);
    state.Require(SUCCEEDED(inspectVolumeAHr),
                  std::format(L"Startup-warmed SQLite volume A inspection failed. hr=0x{:08X}", static_cast<unsigned long>(inspectVolumeAHr)));
    state.Require(SUCCEEDED(inspectVolumeBHr),
                  std::format(L"Startup-warmed SQLite volume B inspection failed. hr=0x{:08X}", static_cast<unsigned long>(inspectVolumeBHr)));
    if (FAILED(inspectVolumeAHr) || FAILED(inspectVolumeBHr))
    {
        return false;
    }
    const bool persistentStoreReady = volumeInfoA.state == SqliteIndexStore::kVolumeStateReady &&
                                      volumeInfoB.state == SqliteIndexStore::kVolumeStateReady;
    state.Require(status.readyForQueryCutover == persistentStoreReady,
                  std::format(L"Startup-warm status readiness must match every persisted volume state. statusReady={} states={},{}.",
                              status.readyForQueryCutover,
                              volumeInfoA.state,
                              volumeInfoB.state));

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = supportA.normalizedRootPath;
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    LocalSearchIndexCore::QueryStats queryStats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, candidates, &queryStats);
    state.Require(SUCCEEDED(hr), std::format(L"Startup-warmed SQLite service query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Startup-warmed SQLite service query returned an unexpected candidate set.");
    if (persistentStoreReady && CanProveDirectSqliteCurrentness(queryStats))
    {
        state.Require(queryStats.usedSqliteStore, L"Current startup roots should use the direct SQLite query path.");
        state.Require(queryStats.sqliteReadOnlyQuery, L"Direct startup-root SQLite queries should remain read-only.");
    }
    else
    {
        state.Require(! queryStats.usedSqliteStore, L"Startup roots without proven currentness must not be queried directly from SQLite.");
        state.Require(queryStats.usedLiveScanFallback, L"Startup roots without proven currentness should use the live-scan fallback.");
        state.Require(queryStats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                      L"Startup roots without proven currentness should report the live-scan execution mode.");
        state.Require(queryStats.fallbackReason != LocalSearchIndexCore::FallbackReason::None,
                      L"Startup-root live fallback should report why SQLite currentness was not proven.");
        if (! persistentStoreReady)
        {
            state.Require(queryStats.sqliteCutoverBlocked, L"An unready aggregate startup store should report blocked SQLite cutover.");
            state.Require(queryStats.fallbackReason == LocalSearchIndexCore::FallbackReason::CutoverBlocked,
                          L"An unready aggregate startup store should report a cutover-blocked fallback reason.");
        }
    }
    state.Require(! queryStats.snapshotSaved, L"Startup-warmed SQLite service query should not create a compatibility snapshot runtime store.");
    state.Require(queryStats.snapshotPath.empty(), L"Startup-warmed SQLite service query should not report a compatibility snapshot path.");

    SearchServiceBroker::ServiceStatus afterQueryStatus{};
    hr = SearchServiceBroker::GetStatus(afterQueryStatus);
    state.Require(SUCCEEDED(hr), std::format(L"SearchServiceBroker::GetStatus after startup-warm query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(
        afterQueryStatus.indexedVolumeCount == expectedRoots.size(),
        std::format(L"Startup-warmed service should keep {} mirrored roots after query, got {}.", expectedRoots.size(), afterQueryStatus.indexedVolumeCount));
    state.Require(afterQueryStatus.discoveredRoots == expectedRoots, L"Startup-warmed service should preserve the overridden discovered-root set in status.");
    state.Require(afterQueryStatus.startupWarmupCompletedRoots == expectedRoots.size(),
                  std::format(L"Startup-warmed service should preserve completed warmup count {} after query, got {}.",
                              expectedRoots.size(),
                              afterQueryStatus.startupWarmupCompletedRoots));

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(output.contains("Warmup"), L"Foreground SQLite startup warmup output should include the warmup dashboard/log field.");
    state.Require(output.contains("Startup warmup"), L"Foreground SQLite startup warmup output should log the startup warmup event.");
    state.Require(output.contains("warmup=2/2 failed=0"), L"Foreground SQLite startup warmup output should report completed warmup counts.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_sqlite_startup_warmup_failure_status_roundtrip",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_sqlite_startup_warmup_failure_status_roundtrip", caseRoot),
                  L"Failed to prepare search_service_sqlite_startup_warmup_failure_status_roundtrip root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path rootA = caseRoot / L"root-a";
    const std::filesystem::path rootZ = caseRoot / L"root-z";
    std::error_code createEc;
    std::filesystem::create_directories(rootA, createEc);
    state.Require(! createEc, std::format(L"Failed to create '{}'. error={}", rootA.wstring(), createEc.value()));
    createEc.clear();
    std::filesystem::create_directories(rootZ, createEc);
    state.Require(! createEc, std::format(L"Failed to create '{}'. error={}", rootZ.wstring(), createEc.value()));
    state.Require(SelfTest::WriteTextFile(rootA / L"alpha.txt", "alpha"), L"Failed to create root-a alpha.txt.");
    state.Require(SelfTest::WriteTextFile(rootZ / L"zeta.txt", "zeta"), L"Failed to create root-z zeta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::Repository expectedRepository;
    LocalSearchIndexCore::SupportInfo supportA{};
    LocalSearchIndexCore::SupportInfo supportZ{};
    HRESULT hr = expectedRepository.ProbePath(rootA.wstring(), supportA);
    state.Require(SUCCEEDED(hr), std::format(L"ProbePath failed for root-a. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    hr = expectedRepository.ProbePath(rootZ.wstring(), supportZ);
    state.Require(SUCCEEDED(hr), std::format(L"ProbePath failed for root-z. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (! state.failure.empty())
    {
        return false;
    }

    std::vector<std::wstring> expectedRoots{supportA.normalizedRootPath, supportZ.normalizedRootPath};
    std::sort(expectedRoots.begin(), expectedRoots.end(), [](const std::wstring& left, const std::wstring& right) noexcept {
        return OrdinalString::LessNoCase(left, right);
    });

    const std::filesystem::path storageRoot        = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath         = caseRoot / L"startup-failure-index.sqlite3";
    const std::wstring previousPipeOverride        = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousRootOverride        = GetEnvVarTrimmed(SearchServiceBroker::kDiscoverRootsEnvVar);
    const std::wstring previousWarmupDelayOverride = GetEnvVarTrimmed(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS");
    const std::wstring pipeName                    = MakeUniquePipeName();
    const std::wstring rootOverride                = std::format(L"{};{}", rootA.wstring(), rootZ.wstring());
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, rootOverride.c_str()) != 0,
                  L"Failed to override the search service startup roots.");
    state.Require(::SetEnvironmentVariableW(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS", L"250") != 0,
                  L"Failed to override the startup warmup delay.");
    const auto restoreOverrides = wil::scope_exit([&] noexcept
    {
        const wchar_t* restorePipeValue  = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        const wchar_t* restoreRootsValue = previousRootOverride.empty() ? nullptr : previousRootOverride.c_str();
        const wchar_t* restoreDelayValue = previousWarmupDelayOverride.empty() ? nullptr : previousWarmupDelayOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restorePipeValue));
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, restoreRootsValue));
        static_cast<void>(::SetEnvironmentVariableW(L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS", restoreDelayValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs),
                  serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    std::error_code removeEc;
    std::filesystem::remove_all(rootZ, removeEc);
    state.Require(! removeEc, std::format(L"Failed to remove startup-failure root '{}'. error={}", rootZ.wstring(), removeEc.value()));
    if (! state.failure.empty())
    {
        return false;
    }

    std::this_thread::sleep_for(std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(2'500))});

    SearchServiceBroker::ServiceStatus status{};
    hr = SearchServiceBroker::GetStatus(status);
    state.Require(
        SUCCEEDED(hr),
        std::format(L"Initial SearchServiceBroker::GetStatus for startup warmup failure service failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(status.startupWarmupEnabled, L"Startup warmup failure test should still report warmup enabled.");
    state.Require(status.startupWarmupTotalRoots == expectedRoots.size(),
                  std::format(L"Expected {} startup warmup roots, got {}.", expectedRoots.size(), status.startupWarmupTotalRoots));
    state.Require(status.discoveredRoots == expectedRoots, L"Startup warmup failure test should preserve the overridden discovered-root set.");

    auto waitForStatus = [&](auto&& predicate, SearchServiceBroker::ServiceStatus& outStatus, std::wstring_view failureMessage) noexcept
    {
        const auto deadline =
            std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(15'000))};
        HRESULT lastStatusHr = S_OK;
        while (std::chrono::steady_clock::now() < deadline)
        {
            SearchServiceBroker::ServiceStatus currentStatus{};
            const HRESULT statusHr = SearchServiceBroker::GetStatus(currentStatus);
            lastStatusHr           = statusHr;
            if (SUCCEEDED(statusHr))
            {
                outStatus = currentStatus;
            }
            if (SUCCEEDED(statusHr) && predicate(currentStatus))
            {
                return true;
            }

            std::this_thread::sleep_for(std::chrono::milliseconds{200});
        }

        state.Require(false,
                      std::format(L"{0} Last status: running={1} completed={2} failed={3} total={4} indexedVolumes={5} ready={6} hasFailure={7} "
                                  L"failureHr=0x{8:08X} failureRoot='{9}' discoveredRoots={10} inspected={11} storePath='{12}' walBytes={13} "
                                  L"storeState={14} syncPhase={15} activeRoot='{16}' lastGetStatusHr=0x{17:08X}.",
                                  failureMessage,
                                  outStatus.startupWarmupRunning,
                                  outStatus.startupWarmupCompletedRoots,
                                  outStatus.startupWarmupFailedRoots,
                                  outStatus.startupWarmupTotalRoots,
                                  outStatus.indexedVolumeCount,
                                  outStatus.readyForQueryCutover,
                                  outStatus.startupWarmupHasFailure,
                                  static_cast<unsigned long>(outStatus.startupWarmupLastFailureHr),
                                  outStatus.startupWarmupLastFailureRoot,
                                  outStatus.discoveredRoots.size(),
                                  outStatus.persistentStoreInspectionSucceeded,
                                  outStatus.persistentStorePath,
                                  outStatus.writeAheadLogBytes,
                                  static_cast<unsigned int>(outStatus.storeState),
                                  static_cast<unsigned int>(outStatus.syncPhase),
                                  outStatus.activeRoot,
                                  static_cast<unsigned long>(lastStatusHr)));
        return false;
    };

    status = {};
    state.Require(waitForStatus(
                      [&](const SearchServiceBroker::ServiceStatus& currentStatus) noexcept
    {
        return ! currentStatus.startupWarmupRunning && currentStatus.startupWarmupCompletedRoots == 1u && currentStatus.startupWarmupFailedRoots == 1u &&
               currentStatus.startupWarmupHasFailure && FAILED(currentStatus.startupWarmupLastFailureHr) &&
               EqualsIgnoreCase(currentStatus.startupWarmupLastFailureRoot, supportZ.normalizedRootPath) && currentStatus.indexedVolumeCount == 1u;
    },
                      status,
                      L"Startup warmup failure test did not reach a stable completed status."),
                  L"Startup warmup failure test did not reach a stable completed status.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(! status.startupWarmupRunning, L"Startup warmup failure test should report a completed warmup state.");
    state.Require(status.startupWarmupCompletedRoots == 1u,
                  std::format(L"Startup warmup failure test should complete exactly one root, got {}.", status.startupWarmupCompletedRoots));
    state.Require(status.startupWarmupFailedRoots == 1u,
                  std::format(L"Startup warmup failure test should record exactly one failed root, got {}.", status.startupWarmupFailedRoots));
    state.Require(status.startupWarmupHasFailure, L"Startup warmup failure test should report the last failure through status.");
    state.Require(FAILED(status.startupWarmupLastFailureHr),
                  std::format(L"Startup warmup failure test should report a failing HRESULT, got 0x{:08X}.",
                              static_cast<unsigned long>(status.startupWarmupLastFailureHr)));
    state.Require(EqualsIgnoreCase(status.startupWarmupLastFailureRoot, supportZ.normalizedRootPath),
                  L"Startup warmup failure test should report the failing root through status.");
    state.Require(status.indexedVolumeCount == 1u,
                  std::format(L"Startup warmup failure test should mirror only the surviving root, got {} indexed volumes.", status.indexedVolumeCount));
    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = supportA.normalizedRootPath;
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    SearchServiceBroker::ServiceStatus queryStatus{};
    hr = SearchServiceBroker::GetStatus(queryStatus);
    state.Require(SUCCEEDED(hr), std::format(L"Startup warmup failure pre-query status failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::VolumeInfo preQueryVolumeInfo{};
    const HRESULT preQueryInspectHr = SqliteIndexStore::InspectVolume(sqlitePath.wstring(), supportA.normalizedRootPath, preQueryVolumeInfo);
    state.Require(SUCCEEDED(preQueryInspectHr),
                  std::format(L"Startup warmup failure pre-query SQLite volume inspection failed. hr=0x{:08X} root='{}' sqlitePath='{}'.",
                              static_cast<unsigned long>(preQueryInspectHr),
                              supportA.normalizedRootPath,
                              sqlitePath.wstring()));
    if (FAILED(preQueryInspectHr))
    {
        return false;
    }
    const bool survivingVolumeReady = preQueryVolumeInfo.state == SqliteIndexStore::kVolumeStateReady;
    state.Require(status.readyForQueryCutover == survivingVolumeReady,
                  std::format(L"Warmup-failure status readiness must match the surviving persisted volume. statusReady={} volumeState={}.",
                              status.readyForQueryCutover,
                              preQueryVolumeInfo.state));
    state.Require(status.legacyImportVolumeCount ==
                      (preQueryVolumeInfo.state == SqliteIndexStore::kVolumeStateImportedLegacySnapshot ? 1u : 0u),
                  L"Warmup-failure legacy-import count must remain separate from aggregate query readiness.");

    LocalSearchIndexCore::QueryStats queryStats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, candidates, &queryStats);
    state.Require(SUCCEEDED(hr), std::format(L"Startup warmup failure follow-up query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Startup warmup failure follow-up query returned an unexpected candidate set.");
    if (survivingVolumeReady && CanProveDirectSqliteCurrentness(queryStats))
    {
        state.Require(queryStats.usedSqliteStore,
                      L"A current surviving warmup root should use the direct SQLite query path when its live cursor is readable.");
        state.Require(queryStats.sqliteReadOnlyQuery, L"Direct surviving-root SQLite queries should remain read-only.");
    }
    else
    {
        state.Require(! queryStats.usedSqliteStore,
                      L"A surviving warmup root without proven currentness must not be queried directly from SQLite.");
        state.Require(queryStats.usedLiveScanFallback,
                      L"A surviving warmup root without proven currentness should use the live-scan fallback.");
        state.Require(queryStats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                      L"A surviving warmup root without proven currentness should report live-scan execution.");
        state.Require(queryStats.fallbackReason != LocalSearchIndexCore::FallbackReason::None,
                      L"A surviving-root live fallback should report why SQLite currentness was not proven.");
        if (! survivingVolumeReady)
        {
            state.Require(queryStats.sqliteCutoverBlocked, L"An unready surviving volume should report blocked SQLite cutover.");
            state.Require(queryStats.fallbackReason == LocalSearchIndexCore::FallbackReason::CutoverBlocked,
                          L"An unready surviving volume should report a cutover-blocked fallback reason.");
        }
    }

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(output.contains("Startup warmup failed"), L"Foreground SQLite startup warmup failure output should log the failure event.");
    state.Require(output.contains("lastFailure=0x"), L"Foreground SQLite startup warmup failure output should report the failing HRESULT.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_foreground_rejects_second_instance",
                  [&](SelfTest::CaseState& state) noexcept
{
    constexpr wchar_t kInstanceEventEnvVar[] = L"REDSALAMANDER_SEARCH_SERVICE_INSTANCE_EVENT";

    const std::filesystem::path servicePath = GetSiblingExecutablePath(L"RedSalamanderSearchService.exe");
    std::error_code existsEc;
    state.Require(! servicePath.empty() && std::filesystem::exists(servicePath, existsEc),
                  std::format(L"Service executable not found: {}", servicePath.wstring()));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring previousPipeOverride     = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousInstanceOverride = GetEnvVarTrimmed(kInstanceEventEnvVar);
    const std::wstring pipeName                 = MakeUniquePipeName();
    const std::wstring instanceEventName        = std::format(LR"(Global\RedSalamander.SearchService.Instance.Test.{})", MakeGuidText());
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    state.Require(::SetEnvironmentVariableW(kInstanceEventEnvVar, instanceEventName.c_str()) != 0,
                  L"Failed to override the search service single-instance event name.");
    const auto restoreOverrides = wil::scope_exit([&] noexcept
    {
        const wchar_t* restorePipeValue     = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        const wchar_t* restoreInstanceValue = previousInstanceOverride.empty() ? nullptr : previousInstanceOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restorePipeValue));
        static_cast<void>(::SetEnvironmentVariableW(kInstanceEventEnvVar, restoreInstanceValue));
    });
    if (! state.failure.empty())
    {
        return false;
    }

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    CapturedProcessResult result{};
    std::wstring runError;
    const std::vector<std::wstring> commandLine{L"--run-foreground", std::format(L"--pipe-name={}", pipeName)};
    state.Require(RunProcessAndCaptureOutput(servicePath.wstring(), commandLine, static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), result, runError),
                  runError);
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(result.exitCode != 0u, L"The second foreground launch should fail when another instance is already running.");
    state.Require(result.output.contains("already running"), L"The duplicate-instance failure should explain that another instance is already running.");
#ifdef _DEBUG
    state.Require(result.output.contains("RedSalamanderSearchService.Debug"), L"The duplicate-instance failure should mention the Debug service name.");
#else
    state.Require(result.output.contains("RedSalamanderSearchService"), L"The duplicate-instance failure should mention the Release service name.");
#endif
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_foreground_logs_request_status",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_foreground_logs_request_status", caseRoot),
                  L"Failed to prepare search_service_foreground_logs_request_status root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    const std::filesystem::path sqlitePath = caseRoot / L"foreground.sqlite3";

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    LocalSearchIndexCore::QueryStats queryStats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    HRESULT queryHr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, candidates, &queryStats);
    state.Require(SUCCEEDED(queryHr), std::format(L"Foreground redirected query failed. hr=0x{:08X}", static_cast<unsigned long>(queryHr)));
    if (FAILED(queryHr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Foreground redirected query returned an unexpected candidate set.");

    SearchServiceBroker::ServiceStatus status{};
    const HRESULT statusHr = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(statusHr), std::format(L"Foreground status probe failed. hr=0x{:08X}", static_cast<unsigned long>(statusHr)));
    if (FAILED(statusHr))
    {
        return false;
    }
    state.Require(EqualsIgnoreCase(status.pipeName, pipeName), L"Foreground readiness/status must come from the intended isolated pipe.");
    state.Require(status.readyForQueryCutover, L"Foreground status probe should report SQLite cutover readiness after the mirrored query.");
    state.Require(
        status.legacyImportVolumeCount == 0u,
        std::format(L"Foreground status probe should report zero legacy-import volumes after the mirrored query. got={}", status.legacyImportVolumeCount));

    std::string output;
    const auto shutdownStart = std::chrono::steady_clock::now();
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(2'000)), output, serviceError), serviceError);
    const auto shutdownDuration = std::chrono::steady_clock::now() - shutdownStart;
    state.Require(shutdownDuration < std::chrono::milliseconds{SelfTest::ScaleTimeout(2'000)},
                  L"Foreground service graceful shutdown should complete within the owner-controlled two-second budget.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(! output.empty(), L"Foreground redirected output should capture at least some lifecycle text.");
    state.Require(output.contains(" db=\""), L"Foreground redirected output should include the database runtime state when the dashboard is unavailable.");
    state.Require(output.contains(" sync=\""), L"Foreground redirected output should include synchronization progress when the dashboard is unavailable.");
    state.Require(output.contains(" search=\""),
                  L"Foreground redirected output should include the active query execution mode when the dashboard is unavailable.");
    state.Require(output.contains("sqlite") || output.contains("live-scan") || output.contains("in-memory-index"),
                  L"Foreground redirected output should surface a concrete query execution mode.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_query_reports_live_progress",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_query_reports_live_progress", caseRoot),
                  L"Failed to prepare search_service_query_reports_live_progress root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"sub"), L"Failed to create sub directory.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"sub" / L"beta.txt", "beta"), L"Failed to create beta.txt.");

    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--storage-root=\"{}\" --store-backend=snapshot", storageRoot.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    BrokerProgressRecorder recorder;
    LocalSearchIndexCore::QueryStats stats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    const HRESULT hr = SearchServiceBroker::Query(request, &BrokerProgressRecorder::ProgressThunk, &recorder, nullptr, nullptr, candidates, &stats);
    state.Require(SUCCEEDED(hr), std::format(L"SearchServiceBroker::Query with progress failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"Live-progress query returned an unexpected candidate set.");

    const auto progress = recorder.Snapshots();
    state.Require(! progress.empty(), L"Live-progress query should report progress callbacks.");
    const RecordedSearchProgress* indexLookup = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP);
    state.Require(indexLookup != nullptr, L"Live-progress query should include an index-lookup progress update.");
    const RecordedSearchProgress* enumerating = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING);
    state.Require(enumerating != nullptr, L"Live-progress query should include an enumerating progress update.");
    if (enumerating != nullptr)
    {
        state.Require(! enumerating->currentPath.empty(), L"Enumerating progress should include a current path.");
        state.Require(enumerating->backend == FILESYSTEM_SEARCH_BACKEND_SERVICE, L"Live-progress query should tag broker progress with the service backend.");
    }

    SearchServiceBroker::ServiceStatus status{};
    const HRESULT statusHr = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(statusHr), std::format(L"Final live-progress status request failed. hr=0x{:08X}", static_cast<unsigned long>(statusHr)));

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_filters_cached_descendants_denied_to_client",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto tryDenyListDirectoryToInteractiveUsers =
        [](const std::filesystem::path& path, PSECURITY_DESCRIPTOR& outOriginalSecurityDescriptor, PACL& outOriginalDacl) noexcept
    {
        outOriginalSecurityDescriptor = nullptr;
        outOriginalDacl               = nullptr;

        std::array<std::byte, SECURITY_MAX_SID_SIZE> sidBuffer{};
        DWORD sidSize = static_cast<DWORD>(sidBuffer.size());
        if (! ::CreateWellKnownSid(WinInteractiveSid, nullptr, sidBuffer.data(), &sidSize))
        {
            return false;
        }

        PACL existingDacl                       = nullptr;
        PSECURITY_DESCRIPTOR securityDescriptor = nullptr;
        const DWORD getSecurityError            = ::GetNamedSecurityInfoW(
            const_cast<wchar_t*>(path.c_str()), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nullptr, nullptr, &existingDacl, nullptr, &securityDescriptor);
        if (getSecurityError != ERROR_SUCCESS || securityDescriptor == nullptr)
        {
            return false;
        }
        auto freeOriginalOnFailure = wil::scope_exit([&] noexcept
        {
            if (securityDescriptor != nullptr)
            {
                static_cast<void>(::LocalFree(securityDescriptor));
            }
        });

        EXPLICIT_ACCESSW denyEntry{};
        denyEntry.grfAccessPermissions = FILE_LIST_DIRECTORY;
        denyEntry.grfAccessMode        = DENY_ACCESS;
        denyEntry.grfInheritance       = NO_INHERITANCE;
        denyEntry.Trustee.TrusteeForm  = TRUSTEE_IS_SID;
        denyEntry.Trustee.TrusteeType  = TRUSTEE_IS_WELL_KNOWN_GROUP;
        denyEntry.Trustee.ptstrName    = reinterpret_cast<wchar_t*>(sidBuffer.data());

        PACL newDacl                = nullptr;
        const DWORD setEntriesError = ::SetEntriesInAclW(1, &denyEntry, existingDacl, &newDacl);
        if (setEntriesError != ERROR_SUCCESS || newDacl == nullptr)
        {
            return false;
        }
        auto freeNewDacl = wil::scope_exit([&] noexcept
        {
            if (newDacl != nullptr)
            {
                static_cast<void>(::LocalFree(newDacl));
            }
        });

        const DWORD setSecurityError =
            ::SetNamedSecurityInfoW(const_cast<wchar_t*>(path.c_str()), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nullptr, nullptr, newDacl, nullptr);
        if (setSecurityError != ERROR_SUCCESS)
        {
            return false;
        }

        outOriginalSecurityDescriptor = securityDescriptor;
        outOriginalDacl               = existingDacl;
        securityDescriptor            = nullptr;
        return true;
    };

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_filters_cached_descendants_denied_to_client", caseRoot),
                  L"Failed to prepare search_service_filters_cached_descendants_denied_to_client root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"blocked"), L"Failed to create blocked directory.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"allowed.txt", "allowed"), L"Failed to create allowed.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"blocked" / L"hidden.txt", "hidden"), L"Failed to create blocked\\hidden.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path sqlitePath = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-acl-filter.sqlite3");
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::EnsureBootstrap for search service ACL filter failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest replaceRequest{};
    replaceRequest.rootPath       = caseRoot.wstring();
    replaceRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replaceRequest.journalId      = 1u;
    replaceRequest.nextUsn        = 1u;
    replaceRequest.state          = SqliteIndexStore::kVolumeStateReady;
    replaceRequest.entries.push_back({
        .fileIdLow      = 1u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 0u,
        .parentIdHigh   = 0u,
        .fullPath       = caseRoot.wstring(),
        .name           = caseRoot.filename().wstring(),
        .attributes     = FILE_ATTRIBUTE_DIRECTORY,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 2u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 1u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"allowed.txt").wstring(),
        .name           = L"allowed.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 7u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 3u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 1u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"blocked").wstring(),
        .name           = L"blocked",
        .attributes     = FILE_ATTRIBUTE_DIRECTORY,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 4u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 3u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"blocked" / L"hidden.txt").wstring(),
        .name           = L"hidden.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 6u,
        .writeTime100ns = 0u,
    });

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), replaceRequest, &replaceResult);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::ReplaceVolume for search service ACL filter failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    PSECURITY_DESCRIPTOR originalSecurityDescriptor = nullptr;
    PACL originalDacl                               = nullptr;
    state.Require(tryDenyListDirectoryToInteractiveUsers(caseRoot / L"blocked", originalSecurityDescriptor, originalDacl),
                  L"Failed to deny list access to the blocked search-service directory.");
    if (! state.failure.empty())
    {
        return false;
    }
    const auto restoreBlockedDacl = wil::scope_exit([&] noexcept
    {
        if (originalSecurityDescriptor != nullptr && originalDacl != nullptr)
        {
            static_cast<void>(::SetNamedSecurityInfoW(
                const_cast<wchar_t*>((caseRoot / L"blocked").c_str()), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nullptr, nullptr, originalDacl, nullptr));
        }
        if (originalSecurityDescriptor != nullptr)
        {
            static_cast<void>(::LocalFree(originalSecurityDescriptor));
        }
    });

    wil::unique_handle blockedProbe(::CreateFileW((caseRoot / L"blocked").c_str(),
                                                  FILE_LIST_DIRECTORY,
                                                  FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                  nullptr,
                                                  OPEN_EXISTING,
                                                  FILE_FLAG_BACKUP_SEMANTICS,
                                                  nullptr));
    state.Require(! blockedProbe, L"Blocked directory list probe should fail for the client token.");
    if (blockedProbe)
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    constexpr wchar_t kAllowDirectSqliteWithoutJournalEnvVar[] = L"REDSALAMANDER_TEST_ALLOW_DIRECT_SQLITE_WITHOUT_JOURNAL";
    const std::wstring previousDirectSqliteOverride            = GetEnvVarTrimmed(kAllowDirectSqliteWithoutJournalEnvVar);
    state.Require(::SetEnvironmentVariableW(kAllowDirectSqliteWithoutJournalEnvVar, L"1") != 0, L"Failed to enable direct SQLite test currentness override.");
    const auto restoreDirectSqliteOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousDirectSqliteOverride.empty() ? nullptr : previousDirectSqliteOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(kAllowDirectSqliteWithoutJournalEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    LocalSearchIndexCore::QueryStats stats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, candidates, &stats);
    state.Require(SUCCEEDED(hr), std::format(L"Search service ACL-filter query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }
    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"allowed.txt"},
                  L"Search service must filter cached descendants whose parent directory is denied to the client.");

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_transient_authorization_failure_is_incomplete_not_cached",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_transient_auth_not_cached", caseRoot),
                  L"Failed to prepare search_service_transient_auth_not_cached root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"allowed.txt", "allowed"), L"Failed to create allowed.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"transient" / L"retry-a.txt", "retry-a"), L"Failed to create transient\\retry-a.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"transient" / L"retry-b.txt", "retry-b"), L"Failed to create transient\\retry-b.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path sqlitePath = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-transient-auth.sqlite3");

    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::EnsureBootstrap for transient authorization cache test failed. hr=0x{:08X}",
                              static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest replaceRequest{};
    replaceRequest.rootPath       = caseRoot.wstring();
    replaceRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replaceRequest.journalId      = 1u;
    replaceRequest.nextUsn        = 1u;
    replaceRequest.state          = SqliteIndexStore::kVolumeStateReady;
    replaceRequest.entries.push_back({
        .fileIdLow      = 1u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 0u,
        .parentIdHigh   = 0u,
        .fullPath       = caseRoot.wstring(),
        .name           = caseRoot.filename().wstring(),
        .attributes     = FILE_ATTRIBUTE_DIRECTORY,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 2u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 1u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"allowed.txt").wstring(),
        .name           = L"allowed.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 7u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 3u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 1u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"transient").wstring(),
        .name           = L"transient",
        .attributes     = FILE_ATTRIBUTE_DIRECTORY,
        .sizeBytes      = 0u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 4u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 3u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"transient" / L"retry-a.txt").wstring(),
        .name           = L"retry-a.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 7u,
        .writeTime100ns = 0u,
    });
    replaceRequest.entries.push_back({
        .fileIdLow      = 5u,
        .fileIdHigh     = 0u,
        .parentIdLow    = 3u,
        .parentIdHigh   = 0u,
        .fullPath       = (caseRoot / L"transient" / L"retry-b.txt").wstring(),
        .name           = L"retry-b.txt",
        .attributes     = FILE_ATTRIBUTE_ARCHIVE,
        .sizeBytes      = 7u,
        .writeTime100ns = 0u,
    });

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), replaceRequest, &replaceResult);
    state.Require(SUCCEEDED(hr),
                  std::format(L"SqliteIndexStore::ReplaceVolume for transient authorization cache test failed. hr=0x{:08X}",
                              static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    wil::unique_handle transientDirectoryLock(::CreateFileW((caseRoot / L"transient").c_str(),
                                                            FILE_LIST_DIRECTORY | FILE_READ_ATTRIBUTES,
                                                            0u,
                                                            nullptr,
                                                            OPEN_EXISTING,
                                                            FILE_FLAG_BACKUP_SEMANTICS,
                                                            nullptr));
    state.Require(static_cast<bool>(transientDirectoryLock),
                  std::format(L"Failed to hold transient directory exclusively. error={}", ::GetLastError()));
    if (! transientDirectoryLock)
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    constexpr wchar_t kAllowDirectSqliteWithoutJournalEnvVar[] = L"REDSALAMANDER_TEST_ALLOW_DIRECT_SQLITE_WITHOUT_JOURNAL";
    const std::wstring previousDirectSqliteOverride            = GetEnvVarTrimmed(kAllowDirectSqliteWithoutJournalEnvVar);
    state.Require(::SetEnvironmentVariableW(kAllowDirectSqliteWithoutJournalEnvVar, L"1") != 0,
                  L"Failed to enable direct SQLite test currentness override.");
    const auto restoreDirectSqliteOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousDirectSqliteOverride.empty() ? nullptr : previousDirectSqliteOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(kAllowDirectSqliteWithoutJournalEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    LocalSearchIndexCore::QueryStats firstStats{};
    std::vector<LocalSearchIndexCore::Candidate> firstCandidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, firstCandidates, &firstStats);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Search service transient authorization first query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(firstCandidates) == std::vector<std::wstring>{L"allowed.txt"},
                  L"Transient authorization failure should skip only candidates under the locked directory on the first query.");
    state.Require((firstStats.warningFlags & FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED) != 0u,
                  std::format(L"Transient authorization failure should report ACCESS_DENIED_SKIPPED. warnings=0x{:08X}.",
                              static_cast<unsigned long>(firstStats.warningFlags)));

    transientDirectoryLock.reset();

    LocalSearchIndexCore::QueryStats secondStats{};
    std::vector<LocalSearchIndexCore::Candidate> secondCandidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, secondCandidates, &secondStats);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Search service transient authorization second query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(secondCandidates) == std::vector<std::wstring>{L"allowed.txt", L"retry-a.txt", L"retry-b.txt"},
                  L"Transient authorization failure must not cache a false directory authorization result for later queries.");
    state.Require((secondStats.warningFlags & FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED) == 0u,
                  std::format(L"Transient authorization second query should not retain the first query warning. warnings=0x{:08X}.",
                              static_cast<unsigned long>(secondStats.warningFlags)));

    std::string output;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), output, serviceError), serviceError);
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_candidate_impersonation_failure_is_incomplete_warning",
                  [&](SelfTest::CaseState& state) noexcept
{
#if ! defined(RS_SEARCH_TEST_HOOKS)
    return state.Skip(L"Search service test hooks are unavailable in this build configuration.");
#else
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_candidate_impersonation_failure", caseRoot),
                  L"Failed to prepare search_service_candidate_impersonation_failure root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"allowed.txt", "allowed"), L"Failed to create allowed.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"nested" / L"hidden.txt", "hidden"), L"Failed to create nested\\hidden.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path storageRoot = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-service-store");

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");

    const auto restoreOverrides = wil::scope_exit([&] noexcept
    {
        const wchar_t* restorePipeValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restorePipeValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=snapshot --test-fail-client-auth-impersonation-once", storageRoot.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    BrokerProgressRecorder recorder;
    LocalSearchIndexCore::QueryStats stats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    const HRESULT hr = SearchServiceBroker::Query(request, &BrokerProgressRecorder::ProgressThunk, &recorder, nullptr, nullptr, candidates, &stats);
    std::string serviceOutput;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), serviceOutput, serviceError), serviceError);
    const std::wstring serviceOutputText(serviceOutput.begin(), serviceOutput.end());
    state.Require(SUCCEEDED(hr),
                  std::format(L"Search service candidate impersonation failure should complete as incomplete, not fail. hr=0x{:08X}. service output='{}'",
                              static_cast<unsigned long>(hr),
                              serviceOutputText));
    if (FAILED(hr))
    {
        return false;
    }

    const std::vector<std::wstring> candidateNames = CollectIndexedCandidateNames(candidates);
    std::wstring candidateNamesText;
    for (const auto& name : candidateNames)
    {
        if (! candidateNamesText.empty())
        {
            candidateNamesText.append(L", ");
        }
        candidateNamesText.push_back(L'\'');
        candidateNamesText.append(name);
        candidateNamesText.push_back(L'\'');
    }
    state.Require(candidateNames == std::vector<std::wstring>{L"allowed.txt"},
                  std::format(L"Search service candidate impersonation failure should skip only the affected nested candidate. got={}",
                              candidateNamesText));
    state.Require((stats.warningFlags & FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED) != 0u,
                  std::format(L"Search service candidate impersonation failure should report ACCESS_DENIED_SKIPPED. warnings=0x{:08X}.",
                              static_cast<unsigned long>(stats.warningFlags)));

    const auto progress = recorder.Snapshots();
    state.Require(! progress.empty(), L"Search service candidate impersonation failure should report broker progress.");
    return state.failure.empty();
#endif
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_status_and_query_roundtrip",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_status_and_query", caseRoot), L"Failed to prepare search_service_status_and_query root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath  = storageRoot / L"index-v2.sqlite3";
    state.Require(SelfTest::EnsureDirectory(storageRoot), L"Failed to prepare search_service_status_and_query service storage root.");

    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr), std::format(L"SqliteIndexStore::EnsureBootstrap for status roundtrip failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, false, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus status{};
    if (! WaitForSearchServiceStatus(state, status, pipeName, L"status roundtrip service"))
    {
        return false;
    }

    state.Require(status.protocolVersion == SearchServiceBroker::kProtocolVersion, L"Search service status reported an unexpected protocol version.");
    state.Require(status.pipeName == pipeName, L"Search service status reported an unexpected pipe name.");
    state.Require(! status.storageRootDirectory.empty(), L"Search service status should report a storage root.");
    state.Require(EqualsIgnoreCase(status.storageRootDirectory, storageRoot.wstring()), L"Search service should report the overridden test storage root.");
    state.Require(status.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                  L"Search service status should default to the sqlite backend.");
    state.Require(! status.persistentStorePath.empty(), L"SQLite-default service status should report the database path.");
    state.Require(EqualsIgnoreCase(status.persistentStorePath, sqlitePath.wstring()),
                  L"SQLite-default service status should report the configured database path.");
    state.Require(status.persistentStorePath.ends_with(L"index-v2.sqlite3"), L"SQLite-default service status should report the default index-v2.sqlite3 path.");
    state.Require(! status.writeAheadLogPath.empty(), L"SQLite-default service status should report the WAL path.");
    state.Require(status.writeAheadLogPath.ends_with(L"index-v2.sqlite3-wal"), L"SQLite-default service status should report the default WAL filename.");
    state.Require(! status.maintenanceQueued, L"SQLite-default service status should not report queued maintenance.");
    state.Require(! status.maintenanceRunning, L"SQLite-default service status should not report running maintenance.");
    state.Require(status.autoCheckpointEnabled, L"SQLite-default service status should advertise automatic checkpointing.");
    state.Require(status.autoCompactionEnabled, L"SQLite-default service status should advertise automatic compaction.");
    state.Require(status.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::Unknown,
                  L"SQLite-default service status should not report a query execution mode before any query runs.");
    state.Require(status.readyForQueryCutover, L"SQLite-default service status should report a ready default store before any query runs.");
    state.Require(status.storeState == LocalSearchIndexCore::StoreState::Ready,
                  L"SQLite-default service status should report a ready store before any query runs.");
    state.Require(status.syncPhase == LocalSearchIndexCore::SyncPhase::Watching,
                  L"SQLite-default service status should report the watching sync phase before any query runs.");
    state.Require(status.fallbackReason == LocalSearchIndexCore::FallbackReason::None,
                  L"SQLite-default service status should not report a fallback reason before any query runs.");

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    LocalSearchIndexCore::QueryStats stats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, candidates, &stats);
    state.Require(SUCCEEDED(hr), std::format(L"SearchServiceBroker::Query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(candidates) == std::vector<std::wstring>{L"alpha.txt"},
                  L"Search service query returned an unexpected candidate set.");
    state.Require(stats.candidateCount == 1u, L"Search service query should report one candidate.");
    state.Require(stats.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback,
                  L"SQLite-default service query should answer from the live filesystem until the exact root is ready in sqlite.");
    state.Require(stats.fallbackReason == LocalSearchIndexCore::FallbackReason::StoreMissing,
                  L"SQLite-default service query should report store-missing for a root that has not been mirrored into sqlite.");
    state.Require(! stats.sqliteCutoverBlocked, L"SQLite-default service query should not report cutover-blocked for a missing mirrored root.");
    state.Require(stats.usedLiveScanFallback, L"SQLite-default service query should report live-scan fallback usage.");
    state.Require(! stats.snapshotLoaded, L"SQLite-default service query should not load a compatibility snapshot.");
    state.Require(! stats.snapshotSaved, L"SQLite-default service query should not save a compatibility snapshot.");
    state.Require(stats.snapshotPath.empty(), L"SQLite-default service query should not report a compatibility snapshot path.");
    state.Require(stats.snapshotFileBytes == 0u, L"SQLite-default service query should report snapshotFileBytes=0.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_service_single_request_uses_query",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousRootOverride = GetEnvVarTrimmed(SearchServiceBroker::kDiscoverRootsEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_single_request_query", caseRoot),
                  L"Failed to prepare search_service_single_request_query root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, caseRoot.c_str()) != 0,
                  L"Failed to override the search service startup roots.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });
    const auto restoreRootOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousRootOverride.empty() ? nullptr : previousRootOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kDiscoverRootsEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! info || ! search)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"auto\"}");
    state.Require(SUCCEEDED(setHr),
                  std::format(L"Failed to configure auto backend for single-request service test. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    std::wstring rootText    = caseRoot.wstring();
    std::wstring namePattern = L"*.txt";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    RecordingSearchCallback callback;
    const HRESULT hr = search->Search(&query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Single-request service search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const auto matches = callback.Matches();
    state.Require(matches.size() == 1u, std::format(L"Single-request service search expected 1 match, got {}.", matches.size()));
    state.Require(FindRecordedSearchMatch(matches, L"alpha.txt") != nullptr, L"Single-request service search missing alpha.txt.");

    const auto progressSnapshots            = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Single-request service search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SERVICE, L"Single-request service search should complete on the service backend.");
        state.Require((completed->warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) == 0u,
                      L"Single-request service search should keep using the service transport rather than falling back to the host.");
    }

    SearchServiceBroker::ServiceStatus status{};
    const HRESULT statusHr = SearchServiceBroker::GetStatus(status);
    state.Require(SUCCEEDED(statusHr),
                  std::format(L"Service should still accept one more request after the search. hr=0x{:08X}", static_cast<unsigned long>(statusHr)));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_service_content_early_stop_stats",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! info || ! search)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"auto\"}");
    state.Require(SUCCEEDED(setHr),
                  std::format(L"Failed to configure auto backend for service early-stop stats test. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_content_early_stop_stats", caseRoot),
                  L"Failed to prepare search_service_content_early_stop_stats root.");
    for (int index = 0; index < 300; ++index)
    {
        const std::filesystem::path path = caseRoot / std::format(L"report-{:03}.txt", index);
        state.Require(SelfTest::WriteTextFile(path, "alpha needle omega"), std::format(L"Failed to create {}.", path.filename().native()));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::filesystem::path sqlitePath  = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-content-early-stop.sqlite3");
    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs =
        std::format(L"--storage-root=\"{}\" --store-backend=sqlite --sqlite-path=\"{}\"", storageRoot.wstring(), sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, false, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring rootText       = caseRoot.wstring();
    std::wstring namePattern    = L"report-*.txt";
    std::wstring contentPattern = L"needle";

    FileSystemSearchQuery query{};
    query.sizeBytes      = sizeof(FileSystemSearchQuery);
    query.rootPath       = rootText.c_str();
    query.namePattern    = namePattern.c_str();
    query.contentPattern = contentPattern.c_str();
    query.flags          = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES | FILESYSTEM_SEARCH_WANT_SNIPPETS);
    query.nameMode       = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode    = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;
    query.maxResults     = 1u;
    query.maxSnippetCharacters = 32u;

    RecordingSearchCallback callback;
    const HRESULT hr = search->Search(&query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Service early-stop search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const auto matches = callback.Matches();
    state.Require(matches.size() == 1u, std::format(L"Service early-stop search expected 1 match, got {}.", matches.size()));

    const auto progressSnapshots            = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Service early-stop search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SERVICE, L"Service early-stop search should complete on the service backend.");
        state.Require(completed->matchedEntries == 1u,
                      std::format(L"Service early-stop search should report one matched entry, got {}.", completed->matchedEntries));
        state.Require(completed->candidateFiles == 1u,
                      std::format(L"Service early-stop search should report one consumed candidate, got {}.", completed->candidateFiles));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_service_indexed_name_latency_and_parity",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! info || ! search)
    {
        return false;
    }

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_indexed_name_latency", caseRoot),
                  L"Failed to prepare search_service_indexed_name_latency root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"sub"), L"Failed to create sub directory.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"sub" / L"beta.txt", "beta"), L"Failed to create sub\\beta.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"notes.md", "ignored"), L"Failed to create notes.md.");

    ULARGE_INTEGER expectedWriteTime{};
    expectedWriteTime.QuadPart = 133838640000000000ull;
    FILETIME expectedWriteTimeFile{};
    expectedWriteTimeFile.dwLowDateTime  = expectedWriteTime.LowPart;
    expectedWriteTimeFile.dwHighDateTime = expectedWriteTime.HighPart;
    state.Require(SetFileLastWriteTime(caseRoot / L"alpha.txt", expectedWriteTimeFile), L"Failed to set alpha.txt last write time.");
    state.Require(SetFileLastWriteTime(caseRoot / L"sub" / L"beta.txt", expectedWriteTimeFile), L"Failed to set sub\\beta.txt last write time.");

    const std::filesystem::path sqlitePath = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-service-index.sqlite3");
    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus readyStatus{};
    if (! WaitForSearchServiceStatusWithProcessDiagnostics(state, readyStatus, pipeName, L"indexed name latency service", service))
    {
        return false;
    }
    state.Require(readyStatus.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                  L"Indexed name latency service should report the SQLite backend before the warmup query.");

    SearchServiceBroker::QueryRequest brokerRequest{};
    brokerRequest.rootPath           = caseRoot.wstring();
    brokerRequest.namePattern        = L"*.txt";
    brokerRequest.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    brokerRequest.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    brokerRequest.recursive          = true;
    brokerRequest.includeFiles       = true;
    brokerRequest.includeDirectories = false;

    LocalSearchIndexCore::QueryStats warmupStats{};
    std::vector<LocalSearchIndexCore::Candidate> warmupCandidates;
    HRESULT hr = SearchServiceBroker::Query(brokerRequest, nullptr, nullptr, nullptr, nullptr, warmupCandidates, &warmupStats);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Warmup broker query failed. hr=0x{:08X}{}",
                              static_cast<unsigned long>(hr),
                              service.TryCaptureExitedOutputForFailure()));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(CollectIndexedCandidateNames(warmupCandidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"Warmup broker query returned an unexpected candidate set.");

    BrokerBatchRecorder batchRecorder;
    LocalSearchIndexCore::QueryStats warmStats{};
    std::vector<LocalSearchIndexCore::Candidate> ignoredCandidates;
    hr = SearchServiceBroker::Query(
        brokerRequest, nullptr, nullptr, nullptr, nullptr, ignoredCandidates, &warmStats, &BrokerBatchRecorder::CandidateBatchThunk, &batchRecorder);
    state.Require(SUCCEEDED(hr), std::format(L"Warm broker batch query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    if (CanProveDirectSqliteCurrentness(warmStats))
    {
        state.Require(warmStats.usedSqliteStore, L"Warm broker batch query should enumerate from SQLite.");
        state.Require(warmStats.sqliteReadOnlyQuery, L"Warm broker batch query should use a read-only SQLite connection.");
        state.Require(! warmStats.sqliteCutoverBlocked, L"Warm broker batch query should not report a blocked SQLite cutover.");
        state.Require(warmStats.usedNamePrefilter, L"Warm broker batch query should use the SQLite extension prefilter.");
    }
    else
    {
        state.Require(! warmStats.usedSqliteStore,
                      L"Warm broker batch query should fall back when no live journal cursor is available for direct SQLite validation.");
        state.Require(! warmStats.usedNamePrefilter, L"Warm broker batch query should not report a SQLite prefilter when direct SQLite remains blocked.");
        state.Require(warmStats.usedLiveScanFallback, L"Warm broker batch query should report live-scan fallback when direct SQLite remains blocked.");
        state.Require(warmStats.fallbackReason != LocalSearchIndexCore::FallbackReason::None,
                      L"Warm broker batch query should report why direct SQLite was not used.");
    }
    state.Require(CollectIndexedCandidateNames(batchRecorder.Candidates()) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"Warm broker batch query returned an unexpected candidate set.");
    state.Require(batchRecorder.BatchesSeen() >= 1u, L"Warm broker batch query should deliver at least one candidate batch.");
    const std::optional<uint32_t> firstBatchElapsedMs = batchRecorder.FirstBatchElapsedMs();
    state.Require(firstBatchElapsedMs.has_value(), L"Warm broker batch query should record first-batch latency.");
    if (firstBatchElapsedMs.has_value())
    {
        const uint32_t batchesSeen    = batchRecorder.BatchesSeen();
        const uint64_t candidateCount = static_cast<uint64_t>(batchRecorder.Candidates().size());
        const uint64_t firstBatchUs   = static_cast<uint64_t>(*firstBatchElapsedMs) * 1000u;
        const std::wstring perfDetail = std::format(L"{}|sqlite={}|prefilter={}|liveFallback={}|fallbackReason={}",
                                                    caseRoot.native(),
                                                    warmStats.usedSqliteStore,
                                                    warmStats.usedNamePrefilter,
                                                    warmStats.usedLiveScanFallback,
                                                    static_cast<int>(warmStats.fallbackReason));
        Debug::Perf::Emit(L"compare.selftest.local_search_service_indexed_first_batch_us", perfDetail, firstBatchUs, candidateCount, batchesSeen, S_OK);
        AppendCompareSelfTestTraceLine(std::format(
            L"Case: local_search_service_indexed_name_latency_and_parity firstBatchMs={} targetMs={} budgetMs={} batches={} candidates={} sqlite={} "
            L"prefilter={} liveFallback={} fallbackReason={}",
            *firstBatchElapsedMs,
            kWarmIndexedFirstBatchTargetMs,
            kWarmIndexedFirstBatchBudgetMs,
            batchesSeen,
            candidateCount,
            warmStats.usedSqliteStore,
            warmStats.usedNamePrefilter,
            warmStats.usedLiveScanFallback,
            static_cast<int>(warmStats.fallbackReason)));
        if (*firstBatchElapsedMs >= kWarmIndexedFirstBatchTargetMs)
        {
            AppendCompareSelfTestTraceLine(std::format(
                L"Case: local_search_service_indexed_name_latency_and_parity first-batch target exceeded under current load; got={} ms target={} ms "
                L"criticalBudget={} ms.",
                *firstBatchElapsedMs,
                kWarmIndexedFirstBatchTargetMs,
                kWarmIndexedFirstBatchBudgetMs));
        }
        state.Require(*firstBatchElapsedMs < kWarmIndexedFirstBatchBudgetMs,
                      std::format(L"Warm broker first batch exceeded critical budget. got={} ms budget={} ms target={} ms.",
                                  *firstBatchElapsedMs,
                                  kWarmIndexedFirstBatchBudgetMs,
                                  kWarmIndexedFirstBatchTargetMs));
    }

    std::wstring rootText    = caseRoot.wstring();
    std::wstring namePattern = L"*.txt";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    const HRESULT setAutoHr = info->SetConfiguration("{\"searchBackendPreference\":\"auto\"}");
    state.Require(SUCCEEDED(setAutoHr), std::format(L"Failed to configure service-backed auto search. hr=0x{:08X}", static_cast<unsigned long>(setAutoHr)));
    if (FAILED(setAutoHr))
    {
        return false;
    }

    RecordingSearchCallback serviceCallback;
    hr = search->Search(&query, &serviceCallback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Service-backed indexed name search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const auto serviceProgress                     = serviceCallback.ProgressSnapshots();
    const RecordedSearchProgress* serviceCompleted = FindRecordedSearchProgress(serviceProgress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(serviceCompleted != nullptr, L"Service-backed indexed name search missing completed progress.");
    if (serviceCompleted != nullptr)
    {
        state.Require(serviceCompleted->backend == FILESYSTEM_SEARCH_BACKEND_SERVICE,
                      L"Service-backed indexed name search should complete on the service backend.");
        state.Require((serviceCompleted->warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) == 0u,
                      L"Service-backed indexed name search should keep using the service transport rather than falling back to the host.");
    }

    const HRESULT setLocalIndexHr = info->SetConfiguration("{\"searchBackendPreference\":\"local-index\"}");
    state.Require(SUCCEEDED(setLocalIndexHr), std::format(L"Failed to configure local-index search. hr=0x{:08X}", static_cast<unsigned long>(setLocalIndexHr)));
    if (FAILED(setLocalIndexHr))
    {
        return false;
    }

    RecordingSearchCallback localIndexCallback;
    hr = search->Search(&query, &localIndexCallback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Local-index name search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const auto localIndexProgress                     = localIndexCallback.ProgressSnapshots();
    const RecordedSearchProgress* localIndexCompleted = FindRecordedSearchProgress(localIndexProgress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(localIndexCompleted != nullptr, L"Local-index name search missing completed progress.");
    if (localIndexCompleted != nullptr)
    {
        state.Require(localIndexCompleted->backend == FILESYSTEM_SEARCH_BACKEND_INDEX, L"Local-index name search should complete on the local-index backend.");
    }

    RecordingSearchCallback fallbackCallback;
    hr = SearchFallbackEngine::Execute(created.fileSystem.get(), &query, &fallbackCallback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Host fallback name search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    auto serviceMatches     = serviceCallback.Matches();
    auto localIndexMatches  = localIndexCallback.Matches();
    auto fallbackMatches    = fallbackCallback.Matches();
    const auto compareMatch = [](const RecordedSearchMatch& left, const RecordedSearchMatch& right) noexcept
    {
        return std::tie(left.fullPath,
                        left.relativePath,
                        left.displayName,
                        left.creationTime,
                        left.lastWriteTime,
                        left.changeTime,
                        left.endOfFile,
                        left.allocationSize,
                        left.matchedBy,
                        left.attributes) < std::tie(right.fullPath,
                                                    right.relativePath,
                                                    right.displayName,
                                                    right.creationTime,
                                                    right.lastWriteTime,
                                                    right.changeTime,
                                                    right.endOfFile,
                                                    right.allocationSize,
                                                    right.matchedBy,
                                                    right.attributes);
    };
    std::sort(serviceMatches.begin(), serviceMatches.end(), compareMatch);
    std::sort(localIndexMatches.begin(), localIndexMatches.end(), compareMatch);
    std::sort(fallbackMatches.begin(), fallbackMatches.end(), compareMatch);

    state.Require(serviceMatches.size() == 2u, std::format(L"Service-backed indexed name search expected 2 matches, got {}.", serviceMatches.size()));
    state.Require(localIndexMatches.size() == fallbackMatches.size(), L"Local-index name search should match host fallback result count.");
    state.Require(serviceMatches.size() == fallbackMatches.size(), L"Service-backed indexed name search should match host fallback result count.");

    for (size_t index = 0; index < serviceMatches.size() && index < fallbackMatches.size(); ++index)
    {
        state.Require(serviceMatches[index].fullPath == fallbackMatches[index].fullPath,
                      std::format(L"Service-backed indexed match {} fullPath mismatch.", index));
        state.Require(serviceMatches[index].relativePath == fallbackMatches[index].relativePath,
                      std::format(L"Service-backed indexed match {} relativePath mismatch.", index));
        state.Require(serviceMatches[index].displayName == fallbackMatches[index].displayName,
                      std::format(L"Service-backed indexed match {} displayName mismatch.", index));
        state.Require(serviceMatches[index].creationTime == fallbackMatches[index].creationTime,
                      std::format(L"Service-backed indexed match {} creationTime mismatch.", index));
        state.Require(serviceMatches[index].lastWriteTime == fallbackMatches[index].lastWriteTime,
                      std::format(L"Service-backed indexed match {} lastWriteTime mismatch.", index));
        state.Require(serviceMatches[index].changeTime == fallbackMatches[index].changeTime,
                      std::format(L"Service-backed indexed match {} changeTime mismatch.", index));
        state.Require(serviceMatches[index].endOfFile == fallbackMatches[index].endOfFile,
                      std::format(L"Service-backed indexed match {} size mismatch.", index));
        state.Require(serviceMatches[index].allocationSize == fallbackMatches[index].allocationSize,
                      std::format(L"Service-backed indexed match {} allocationSize mismatch.", index));
        state.Require(serviceMatches[index].matchedBy == fallbackMatches[index].matchedBy,
                      std::format(L"Service-backed indexed match {} matchedBy mismatch.", index));
        state.Require(serviceMatches[index].attributes == fallbackMatches[index].attributes,
                      std::format(L"Service-backed indexed match {} attributes mismatch.", index));
    }

    for (size_t index = 0; index < localIndexMatches.size() && index < fallbackMatches.size(); ++index)
    {
        state.Require(localIndexMatches[index].fullPath == fallbackMatches[index].fullPath, std::format(L"Local-index match {} fullPath mismatch.", index));
        state.Require(localIndexMatches[index].relativePath == fallbackMatches[index].relativePath,
                      std::format(L"Local-index match {} relativePath mismatch.", index));
        state.Require(localIndexMatches[index].displayName == fallbackMatches[index].displayName,
                      std::format(L"Local-index match {} displayName mismatch.", index));
        state.Require(localIndexMatches[index].creationTime == fallbackMatches[index].creationTime,
                      std::format(L"Local-index match {} creationTime mismatch.", index));
        state.Require(localIndexMatches[index].lastWriteTime == fallbackMatches[index].lastWriteTime,
                      std::format(L"Local-index match {} lastWriteTime mismatch.", index));
        state.Require(localIndexMatches[index].changeTime == fallbackMatches[index].changeTime,
                      std::format(L"Local-index match {} changeTime mismatch.", index));
        state.Require(localIndexMatches[index].endOfFile == fallbackMatches[index].endOfFile, std::format(L"Local-index match {} size mismatch.", index));
        state.Require(localIndexMatches[index].allocationSize == fallbackMatches[index].allocationSize,
                      std::format(L"Local-index match {} allocationSize mismatch.", index));
        state.Require(localIndexMatches[index].matchedBy == fallbackMatches[index].matchedBy, std::format(L"Local-index match {} matchedBy mismatch.", index));
        state.Require(localIndexMatches[index].attributes == fallbackMatches[index].attributes,
                      std::format(L"Local-index match {} attributes mismatch.", index));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_service_matches_host_fallback",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! info || ! search)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"auto\"}");
    state.Require(SUCCEEDED(setHr), std::format(L"Failed to configure service-backed auto search. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_matches_fallback", caseRoot), L"Failed to prepare search_service_matches_fallback root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"sub"), L"Failed to create sub directory.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"report-one.txt", "alpha needle omega"), L"Failed to create report-one.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"sub" / L"report-two.txt", "prefix needle suffix"), L"Failed to create sub\\report-two.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"other.txt", "needle"), L"Failed to create other.txt.");
    ULARGE_INTEGER expectedWriteTime{};
    expectedWriteTime.QuadPart = 133838640000000000ull;
    FILETIME expectedWriteTimeFile{};
    expectedWriteTimeFile.dwLowDateTime  = expectedWriteTime.LowPart;
    expectedWriteTimeFile.dwHighDateTime = expectedWriteTime.HighPart;
    state.Require(SetFileLastWriteTime(caseRoot / L"report-one.txt", expectedWriteTimeFile), L"Failed to set report-one.txt last write time.");
    state.Require(SetFileLastWriteTime(caseRoot / L"sub" / L"report-two.txt", expectedWriteTimeFile), L"Failed to set report-two.txt last write time.");

    const std::filesystem::path sqlitePath = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-service-index.sqlite3");
    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, false, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring rootText       = caseRoot.wstring();
    std::wstring namePattern    = L"report*.txt";
    std::wstring contentPattern = L"needle";

    FileSystemSearchQuery query{};
    query.sizeBytes      = sizeof(FileSystemSearchQuery);
    query.rootPath       = rootText.c_str();
    query.namePattern    = namePattern.c_str();
    query.contentPattern = contentPattern.c_str();
    query.flags          = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES | FILESYSTEM_SEARCH_WANT_SNIPPETS);
    query.nameMode       = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode    = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;
    query.maxSnippetCharacters = 48u;

    RecordingSearchCallback warmupCallback;
    const HRESULT warmupHr = search->Search(&query, &warmupCallback, nullptr);
    state.Require(SUCCEEDED(warmupHr), std::format(L"Warmup service-backed plugin search failed. hr=0x{:08X}", static_cast<unsigned long>(warmupHr)));
    if (FAILED(warmupHr))
    {
        return false;
    }

    RecordingSearchCallback nativeCallback;
    const HRESULT nativeHr = search->Search(&query, &nativeCallback, nullptr);
    state.Require(SUCCEEDED(nativeHr), std::format(L"Service-backed plugin search failed. hr=0x{:08X}", static_cast<unsigned long>(nativeHr)));
    if (FAILED(nativeHr))
    {
        return false;
    }

    RecordingSearchCallback fallbackCallback;
    const HRESULT fallbackHr = SearchFallbackEngine::Execute(created.fileSystem.get(), &query, &fallbackCallback, nullptr);
    state.Require(SUCCEEDED(fallbackHr), std::format(L"Host fallback search failed. hr=0x{:08X}", static_cast<unsigned long>(fallbackHr)));
    if (FAILED(fallbackHr))
    {
        return false;
    }

    auto nativeMatches      = nativeCallback.Matches();
    auto fallbackMatches    = fallbackCallback.Matches();
    const auto compareMatch = [](const RecordedSearchMatch& left, const RecordedSearchMatch& right) noexcept
    {
        return std::tie(left.fullPath,
                        left.relativePath,
                        left.displayName,
                        left.previewText,
                        left.creationTime,
                        left.lastWriteTime,
                        left.changeTime,
                        left.endOfFile,
                        left.allocationSize,
                        left.matchedBy,
                        left.attributes) < std::tie(right.fullPath,
                                                    right.relativePath,
                                                    right.displayName,
                                                    right.previewText,
                                                    right.creationTime,
                                                    right.lastWriteTime,
                                                    right.changeTime,
                                                    right.endOfFile,
                                                    right.allocationSize,
                                                    right.matchedBy,
                                                    right.attributes);
    };
    std::sort(nativeMatches.begin(), nativeMatches.end(), compareMatch);
    std::sort(fallbackMatches.begin(), fallbackMatches.end(), compareMatch);
    state.Require(nativeMatches.size() == fallbackMatches.size(), L"Service-backed native search should match host fallback result count.");
    for (size_t index = 0; index < nativeMatches.size() && index < fallbackMatches.size(); ++index)
    {
        state.Require(nativeMatches[index].fullPath == fallbackMatches[index].fullPath, std::format(L"Service-backed match {} fullPath mismatch.", index));
        state.Require(nativeMatches[index].relativePath == fallbackMatches[index].relativePath,
                      std::format(L"Service-backed match {} relativePath mismatch.", index));
        state.Require(nativeMatches[index].displayName == fallbackMatches[index].displayName,
                      std::format(L"Service-backed match {} displayName mismatch.", index));
        state.Require(nativeMatches[index].previewText == fallbackMatches[index].previewText,
                      std::format(L"Service-backed match {} previewText mismatch.", index));
        state.Require(nativeMatches[index].matchedBy == fallbackMatches[index].matchedBy, std::format(L"Service-backed match {} matchedBy mismatch.", index));
        state.Require(nativeMatches[index].attributes == fallbackMatches[index].attributes,
                      std::format(L"Service-backed match {} attributes mismatch.", index));
        state.Require(nativeMatches[index].creationTime == fallbackMatches[index].creationTime,
                      std::format(L"Service-backed match {} creationTime mismatch.", index));
        state.Require(nativeMatches[index].lastWriteTime == fallbackMatches[index].lastWriteTime,
                      std::format(L"Service-backed match {} lastWriteTime mismatch.", index));
        state.Require(nativeMatches[index].changeTime == fallbackMatches[index].changeTime,
                      std::format(L"Service-backed match {} changeTime mismatch.", index));
        state.Require(nativeMatches[index].endOfFile == fallbackMatches[index].endOfFile, std::format(L"Service-backed match {} size mismatch.", index));
        state.Require(nativeMatches[index].allocationSize == fallbackMatches[index].allocationSize,
                      std::format(L"Service-backed match {} allocationSize mismatch.", index));
    }

    const auto nativeProgress                     = nativeCallback.ProgressSnapshots();
    const RecordedSearchProgress* nativeCompleted = FindRecordedSearchProgress(nativeProgress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(nativeCompleted != nullptr, L"Service-backed native search missing completed progress.");
    if (nativeCompleted != nullptr)
    {
        state.Require(nativeCompleted->backend == FILESYSTEM_SEARCH_BACKEND_SERVICE, L"Service-backed native search should report the service backend.");
        state.Require((nativeCompleted->warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) == 0u,
                      L"Service-backed native search should keep using the service transport rather than falling back to the host.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_service_protocol_mismatch_falls_back_local_index",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion + 1u, false, serviceError), serviceError);
    std::this_thread::sleep_for(std::chrono::milliseconds(100));

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! info || ! search)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"auto\"}");
    state.Require(SUCCEEDED(setHr), std::format(L"Failed to configure auto search. hr=0x{:08X}", static_cast<unsigned long>(setHr)));

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_protocol_mismatch", caseRoot), L"Failed to prepare search_service_protocol_mismatch root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"pref.txt", "preference"), L"Failed to create pref.txt.");

    std::wstring rootText    = caseRoot.wstring();
    std::wstring namePattern = L"*.txt";
    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;

    RecordingSearchCallback callback;
    const HRESULT searchHr = search->Search(&query, &callback, nullptr);
    state.Require(SUCCEEDED(searchHr), std::format(L"Search with protocol-mismatched service failed. hr=0x{:08X}", static_cast<unsigned long>(searchHr)));

    const auto progressSnapshots            = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Search with protocol-mismatched service missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend != FILESYSTEM_SEARCH_BACKEND_SERVICE, L"Protocol mismatch should not leave the search on the service backend.");
        state.Require((completed->warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) != 0u,
                      L"Protocol mismatch fallback should report the service-unavailable warning.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_service_disconnect_falls_back_local_index",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    state.Require(service.Start(pipeName, 1u, SearchServiceBroker::kProtocolVersion, true, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! info || ! search)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"auto\"}");
    state.Require(SUCCEEDED(setHr), std::format(L"Failed to configure auto search. hr=0x{:08X}", static_cast<unsigned long>(setHr)));

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_disconnect", caseRoot), L"Failed to prepare search_service_disconnect root.");
    for (int index = 0; index < 400; ++index)
    {
        state.Require(SelfTest::WriteTextFile(caseRoot / std::format(L"file-{:03}.txt", index), "value"), L"Failed to create a disconnect-fallback test file.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring rootText    = caseRoot.wstring();
    std::wstring namePattern = L"*.txt";
    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;

    RecordingSearchCallback callback;
    const HRESULT searchHr = search->Search(&query, &callback, nullptr);
    state.Require(SUCCEEDED(searchHr), std::format(L"Search with disconnecting service failed. hr=0x{:08X}", static_cast<unsigned long>(searchHr)));
    if (FAILED(searchHr))
    {
        return false;
    }

    const auto matches = callback.Matches();
    state.Require(matches.size() == 400u, std::format(L"Disconnect fallback search expected 400 matches, got {}.", matches.size()));

    const auto progressSnapshots            = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progressSnapshots, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Disconnect fallback search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend != FILESYSTEM_SEARCH_BACKEND_SERVICE, L"Disconnecting service should not leave the search on the service backend.");
        state.Require((completed->warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) != 0u,
                      L"Disconnect fallback should report the service-unavailable warning.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_multi_client_and_rebuild_control",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_multi_client", caseRoot), L"Failed to prepare search_service_multi_client root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::wstring extraArgs            = std::format(L"--storage-root=\"{}\" --store-backend=snapshot", storageRoot.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, false, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus initialStatus{};
    if (! WaitForSearchServiceStatus(state, initialStatus, pipeName, L"multi-client rebuild service"))
    {
        return false;
    }

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    struct QueryOutcome final
    {
        HRESULT hr = E_FAIL;
        std::vector<LocalSearchIndexCore::Candidate> candidates;
    };

    std::array<QueryOutcome, 2> outcomes{};
    std::jthread clientA([&]() noexcept
    {
        LocalSearchIndexCore::QueryStats stats{};
        outcomes[0].hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, outcomes[0].candidates, &stats);
    });
    std::jthread clientB([&]() noexcept
    {
        LocalSearchIndexCore::QueryStats stats{};
        outcomes[1].hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, outcomes[1].candidates, &stats);
    });

    clientA.join();
    clientB.join();

    for (size_t index = 0; index < outcomes.size(); ++index)
    {
        state.Require(SUCCEEDED(outcomes[index].hr),
                      std::format(L"Service query {} failed. hr=0x{:08X}", index, static_cast<unsigned long>(outcomes[index].hr)));
        state.Require(CollectIndexedCandidateNames(outcomes[index].candidates) == std::vector<std::wstring>{L"alpha.txt"},
                      std::format(L"Service query {} returned an unexpected candidate set.", index));
    }

    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.txt", "beta"), L"Failed to create beta.txt.");
    const HRESULT rebuildHr = SearchServiceBroker::RequestRebuild(caseRoot.wstring());
    state.Require(SUCCEEDED(rebuildHr), std::format(L"SearchServiceBroker::RequestRebuild failed. hr=0x{:08X}", static_cast<unsigned long>(rebuildHr)));
    if (FAILED(rebuildHr))
    {
        return false;
    }

    LocalSearchIndexCore::QueryStats rebuiltStats{};
    std::vector<LocalSearchIndexCore::Candidate> rebuiltCandidates;
    const HRESULT queryHr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, rebuiltCandidates, &rebuiltStats);
    state.Require(SUCCEEDED(queryHr), std::format(L"Search service query after rebuild failed. hr=0x{:08X}", static_cast<unsigned long>(queryHr)));
    state.Require(CollectIndexedCandidateNames(rebuiltCandidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"Search service rebuild control should refresh the indexed candidate set.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_rebuild_deleted_root_purges_index",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_rebuild_deleted_root", caseRoot),
                  L"Failed to prepare search_service_rebuild_deleted_root root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"stale.txt", "stale"), L"Failed to create stale.txt.");

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::filesystem::path storageRoot = caseRoot / L"service-store";
    const std::wstring extraArgs            = std::format(L"--storage-root=\"{}\" --store-backend=snapshot", storageRoot.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArgs), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::ServiceStatus readyStatus{};
    if (! WaitForSearchServiceStatusWithProcessDiagnostics(state, readyStatus, pipeName, L"deleted-root rebuild service", service))
    {
        return false;
    }

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    LocalSearchIndexCore::QueryStats seedStats{};
    std::vector<LocalSearchIndexCore::Candidate> seedCandidates;
    HRESULT hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, seedCandidates, &seedStats);
    state.Require(SUCCEEDED(hr), std::format(L"Search service seed query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(CollectIndexedCandidateNames(seedCandidates) == std::vector<std::wstring>{L"stale.txt"},
                  L"Search service seed query should index stale.txt before the root is deleted.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::error_code removeEc;
    std::filesystem::remove_all(caseRoot, removeEc);
    state.Require(! removeEc, std::format(L"Failed to delete indexed root before rebuild. error={}", removeEc.value()));
    state.Require(! std::filesystem::exists(caseRoot, removeEc), L"Deleted indexed root should no longer exist before rebuild.");
    if (! state.failure.empty())
    {
        return false;
    }

    hr = SearchServiceBroker::RequestRebuild(caseRoot.wstring());
    state.Require(SUCCEEDED(hr), std::format(L"SearchServiceBroker::RequestRebuild should purge a deleted root. hr=0x{:08X}",
                                             static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    std::error_code createEc;
    std::filesystem::create_directories(caseRoot, createEc);
    state.Require(! createEc, std::format(L"Failed to recreate empty root after rebuild. error={}", createEc.value()));
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::QueryStats rebuiltStats{};
    std::vector<LocalSearchIndexCore::Candidate> rebuiltCandidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, rebuiltCandidates, &rebuiltStats);
    state.Require(SUCCEEDED(hr), std::format(L"Search service query after deleted-root rebuild failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(CollectIndexedCandidateNames(rebuiltCandidates).empty(),
                  L"Search service deleted-root rebuild should purge stale candidates under the rebuilt root.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_text_helpers_decoding_and_binary",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto makeBytes = [](std::initializer_list<uint8_t> values) noexcept
    {
        std::vector<std::byte> bytes;
        bytes.reserve(values.size());
        for (const uint8_t value : values)
        {
            bytes.push_back(static_cast<std::byte>(value));
        }
        return bytes;
    };

    SearchTextHelpers::DecodedTextResult decoded{};

    const auto utf8Bom = makeBytes({0xEFu, 0xBBu, 0xBFu, static_cast<uint8_t>('h'), static_cast<uint8_t>('i')});
    state.Require(SearchTextHelpers::TryDecodeSearchableText(utf8Bom, 0u, decoded), L"UTF-8 BOM text should decode.");
    state.Require(decoded.encoding == SearchTextHelpers::DecodedTextEncoding::Utf8, L"UTF-8 BOM should report UTF-8 encoding.");
    state.Require(decoded.text == L"hi", L"UTF-8 BOM decoded text mismatch.");

    const auto utf16Le = makeBytes({0xFFu, 0xFEu, static_cast<uint8_t>('o'), 0x00u, static_cast<uint8_t>('k'), 0x00u});
    state.Require(SearchTextHelpers::TryDecodeSearchableText(utf16Le, 0u, decoded), L"UTF-16 LE text should decode.");
    state.Require(decoded.encoding == SearchTextHelpers::DecodedTextEncoding::Utf16Le, L"UTF-16 LE should report UTF-16 LE encoding.");
    state.Require(decoded.text == L"ok", L"UTF-16 LE decoded text mismatch.");

    const auto utf32Le =
        makeBytes({0xFFu, 0xFEu, 0x00u, 0x00u, static_cast<uint8_t>('A'), 0x00u, 0x00u, 0x00u, static_cast<uint8_t>('B'), 0x00u, 0x00u, 0x00u});
    state.Require(SearchTextHelpers::TryDecodeSearchableText(utf32Le, 0u, decoded), L"UTF-32 LE text should decode.");
    state.Require(decoded.encoding == SearchTextHelpers::DecodedTextEncoding::Utf32Le, L"UTF-32 LE should report UTF-32 LE encoding.");
    state.Require(decoded.text == L"AB", L"UTF-32 LE decoded text mismatch.");

    const auto binary = makeBytes({static_cast<uint8_t>('A'), 0x00u, static_cast<uint8_t>('B'), static_cast<uint8_t>('C')});
    state.Require(! SearchTextHelpers::TryDecodeSearchableText(binary, 0u, decoded), L"Binary data should not decode as searchable text.");
    state.Require(decoded.binary, L"Binary data should set the binary flag.");

    const auto ansi = makeBytes({0xE9u});
    state.Require(SearchTextHelpers::TryDecodeSearchableText(ansi, 1252u, decoded), L"ANSI fallback should decode CP-1252 text.");
    state.Require(decoded.encoding == SearchTextHelpers::DecodedTextEncoding::Ansi, L"ANSI fallback should report ANSI encoding.");
    state.Require(decoded.usedFallbackCodePage, L"ANSI fallback should record fallback code page usage.");
    state.Require(decoded.text == L"\u00E9", L"ANSI fallback decoded text mismatch.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_invariant_case_folding_consistency",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto makeExpectedResultKey = [](std::wstring_view pluginId, std::wstring_view instanceContext, std::wstring_view fullPath) noexcept
    {
        std::wstring key = OrdinalString::FoldCaseInvariant(pluginId);
        key.push_back(L'|');
        key.append(OrdinalString::FoldCaseInvariant(instanceContext));
        key.push_back(L'|');
        key.append(OrdinalString::FoldCaseInvariant(fullPath));
        return key;
    };

    const std::wstring pluginId        = L"PLUGIN-I-\u0130";
    const std::wstring instanceContext = L"Instance-\u00C9";
    const std::wstring fullPath        = L"C:\\Search\\\u0130TEM\\\u00C9quipe.txt";
    const std::wstring expectedKey     = makeExpectedResultKey(pluginId, instanceContext, fullPath);

    state.Require(DebugMakeFindFilesResultKeyForTests(pluginId, instanceContext, fullPath) == expectedKey,
                  L"Find result identity key must use invariant folding for non-ASCII path/context text.");

    const char* previousLocaleRaw    = std::setlocale(LC_CTYPE, nullptr);
    const std::string previousLocale = previousLocaleRaw != nullptr ? std::string(previousLocaleRaw) : std::string();
    const auto restoreLocale         = wil::scope_exit([&]() noexcept
    {
        if (! previousLocale.empty())
        {
            static_cast<void>(std::setlocale(LC_CTYPE, previousLocale.c_str()));
        }
    });

    bool appliedTurkishLocale = false;
    const std::array<const char*, 3> turkishLocaleNames{{"Turkish_Turkey.1254", "tr-TR", "tr_TR"}};
    for (const char* localeName : turkishLocaleNames)
    {
        if (std::setlocale(LC_CTYPE, localeName) != nullptr)
        {
            appliedTurkishLocale = true;
            break;
        }
    }

    if (appliedTurkishLocale)
    {
        state.Require(DebugMakeFindFilesResultKeyForTests(pluginId, instanceContext, fullPath) == expectedKey,
                      L"Find result identity key must stay invariant under a Turkish C locale.");
    }
    else
    {
        Debug::Warning(L"CompareSelfTest: Turkish C locale unavailable; invariant key fixture still covers non-ASCII folding.");
    }

    size_t containsOffset = std::wstring_view::npos;
    state.Require(OrdinalString::FindContainsFoldedInvariant(L"prefix \u00C9quipe.txt", L"\u00E9QUIPE", containsOffset) && containsOffset == 7u,
                  std::format(L"Invariant contains helper offset mismatch. got={}", containsOffset));
    state.Require(FolderViewIncrementalSearch::StartsWithNoCase(L"\u00C9quipe.txt", L"\u00E9QUIPE"),
                  L"Folder incremental prefix search must use invariant folding.");

    const std::optional<UINT32> folderContainsOffset = FolderViewIncrementalSearch::FindContainsOffsetNoCase(L"prefix \u00C9quipe.txt", L"\u00E9QUIPE");
    state.Require(folderContainsOffset.has_value() && folderContainsOffset.value() == 7u,
                  std::format(L"Folder incremental contains offset mismatch. got={}", folderContainsOffset.value_or(UINT32_MAX)));

    size_t literalOffset = std::wstring_view::npos;
    state.Require(SearchTextHelpers::FindLiteralWithChunkOverlap(L"prefix \u00C9quipe suffix", L"\u00E9QUIPE", false, 4u, literalOffset) && literalOffset == 7u,
                  std::format(L"Literal content folding offset mismatch. got={}", literalOffset));

    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_invariant_case_folding_consistency", caseRoot),
                  L"Failed to prepare search_invariant_case_folding_consistency root.");
    const std::filesystem::path filePath = caseRoot / L"\u00C9quipe.txt";
    state.Require(SelfTest::WriteTextFile(filePath, "RAPPORT \xC3\x89QUIPE"), L"Failed to create invariant-fold search fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring rootText       = caseRoot.wstring();
    const std::wstring namePattern    = L"\u00E9QUIPE";
    const std::wstring contentPattern = L"\u00E9quipe";

    FileSystemSearchQuery query{};
    query.sizeBytes      = sizeof(FileSystemSearchQuery);
    query.rootPath       = rootText.c_str();
    query.namePattern    = namePattern.c_str();
    query.contentPattern = contentPattern.c_str();
    query.flags          = FILESYSTEM_SEARCH_INCLUDE_FILES;
    query.nameMode       = FILESYSTEM_SEARCH_NAME_LITERAL;
    query.contentMode    = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;

    RecordingSearchCallback callback;
    const HRESULT hr = SearchFallbackEngine::Execute(baseFs.get(), &query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Invariant fallback search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const auto matches = callback.Matches();
    state.Require(matches.size() == 1u, std::format(L"Invariant fallback search expected 1 match, got {}.", matches.size()));
    const RecordedSearchMatch* match = FindRecordedSearchMatch(matches, L"\u00C9quipe.txt");
    state.Require(match != nullptr, L"Invariant fallback search missing the accented file.");
    if (match != nullptr)
    {
        state.Require((match->matchedBy & FILESYSTEM_SEARCH_MATCH_SOURCE_NAME) != 0u, L"Invariant fallback search should match by name.");
        state.Require((match->matchedBy & FILESYSTEM_SEARCH_MATCH_SOURCE_CONTENT) != 0u, L"Invariant fallback search should match by content.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"search_text_helpers_chunk_overlap_literal_and_regex",
                  [&](SelfTest::CaseState& state) noexcept
{
    size_t literalPos       = std::wstring_view::npos;
    const bool literalFound = SearchTextHelpers::FindLiteralWithChunkOverlap(L"aaaaXYZneedlez", L"XYZneedle", true, 4u, literalPos);
    state.Require(literalFound, L"Chunk-overlap literal search should find cross-chunk matches.");
    state.Require(literalPos == 4u, std::format(L"Chunk-overlap literal match position mismatch. got={}", literalPos));

    SearchTextHelpers::DecodedTextResult decoded{};
    decoded.text     = L"prefix abc123 suffix";
    decoded.encoding = SearchTextHelpers::DecodedTextEncoding::Utf8;

    const std::wregex regex(L"abc\\d+");
    SearchTextHelpers::TextSearchPattern pattern{};
    pattern.mode          = FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX;
    pattern.pattern       = L"abc\\d+";
    pattern.compiledRegex = &regex;
    pattern.caseSensitive = true;

    SearchTextHelpers::TextSearchResult result{};
    state.Require(SearchTextHelpers::MatchDecodedText(decoded, pattern, 32u, true, result), L"Regex helper should match decoded text.");
    state.Require(result.matched, L"Regex helper should report a match.");
    state.Require(result.matchOffset == 7u, std::format(L"Regex helper offset mismatch. got={}", result.matchOffset));
    state.Require(result.matchLength == 6u, std::format(L"Regex helper length mismatch. got={}", result.matchLength));
    state.Require(result.previewText.find(L"abc123") != std::wstring::npos, L"Regex helper snippet should include the match.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"host_fallback_regex_rejected_single_completion",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"host_fallback_regex_rejected", caseRoot), L"Failed to prepare host_fallback_regex_rejected root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"aaa.txt", "aaa"), L"Failed to create aaa.txt.");

    const std::wstring rootText = caseRoot.wstring();

    const auto requireRejected = [&](std::wstring_view label, std::wstring_view regexPattern, bool contentRegex) noexcept
    {
        const std::wstring regexText(regexPattern);
        const std::wstring nameWildcard = L"*.txt";

        FileSystemSearchQuery query{};
        query.sizeBytes      = sizeof(FileSystemSearchQuery);
        query.rootPath       = rootText.c_str();
        query.namePattern    = contentRegex ? nameWildcard.c_str() : regexText.c_str();
        query.contentPattern = contentRegex ? regexText.c_str() : nullptr;
        query.flags          = FILESYSTEM_SEARCH_INCLUDE_FILES;
        query.nameMode       = contentRegex ? FILESYSTEM_SEARCH_NAME_WILDCARD : FILESYSTEM_SEARCH_NAME_REGEX;
        query.contentMode    = contentRegex ? FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX : FILESYSTEM_SEARCH_CONTENT_DISABLED;

        RecordingSearchCallback callback;
        const HRESULT hr = SearchFallbackEngine::Execute(baseFs.get(), &query, &callback, nullptr);
        state.Require(hr == E_INVALIDARG,
                      std::format(L"{} unsafe fallback regex should return E_INVALIDARG, got 0x{:08X}.", label, static_cast<unsigned long>(hr)));
        state.Require(callback.Matches().empty(), std::format(L"{} unsafe fallback regex should not emit matches.", label));
        state.Require(callback.ProgressCalls() == 1u,
                      std::format(L"{} unsafe fallback regex should emit exactly one progress callback, got {}.", label, callback.ProgressCalls()));

        const auto progress                     = callback.ProgressSnapshots();
        const RecordedSearchProgress* completed = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
        state.Require(completed != nullptr, std::format(L"{} unsafe fallback regex missing completed progress.", label));
        if (completed != nullptr)
        {
            state.Require((completed->warningFlags & FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED) != 0u,
                          std::format(L"{} unsafe fallback regex missing REGEX_REJECTED warning. warnings=0x{:08X}.", label, completed->warningFlags));
            state.Require(
                completed->statusHint == E_INVALIDARG,
                std::format(L"{} unsafe fallback regex statusHint mismatch. got=0x{:08X}.", label, static_cast<unsigned long>(completed->statusHint)));
        }
    };

    requireRejected(L"duplicate-alternative-name", L"(a|a)*", false);
    requireRejected(L"prefix-overlap-alternative-name", L"(a|ab)*x", false);
    requireRejected(L"duplicate-alternative-content", L"(a|a)*", true);
    requireRejected(L"prefix-overlap-alternative-content", L"(a|ab)*x", true);

    const std::wstring tooLongPattern(SearchTextHelpers::kMaxRegexPatternLength + 1u, L'a');
    requireRejected(L"too-long-name", tooLongPattern, false);
    requireRejected(L"too-long-content", tooLongPattern, true);

    std::wstring tooDeepPattern;
    tooDeepPattern.reserve((SearchTextHelpers::kMaxRegexGroupDepth + 1u) * 2u + 1u);
    for (size_t depth = 0u; depth <= SearchTextHelpers::kMaxRegexGroupDepth; ++depth)
    {
        tooDeepPattern.push_back(L'(');
    }
    tooDeepPattern.push_back(L'a');
    for (size_t depth = 0u; depth <= SearchTextHelpers::kMaxRegexGroupDepth; ++depth)
    {
        tooDeepPattern.push_back(L')');
    }
    requireRejected(L"too-deep-name", tooDeepPattern, false);
    requireRejected(L"too-deep-content", tooDeepPattern, true);

    const std::wstring boundedPattern = LR"(^a{2,4}\.txt$)";
    FileSystemSearchQuery boundedQuery{};
    boundedQuery.sizeBytes   = sizeof(FileSystemSearchQuery);
    boundedQuery.rootPath    = rootText.c_str();
    boundedQuery.namePattern = boundedPattern.c_str();
    boundedQuery.flags       = FILESYSTEM_SEARCH_INCLUDE_FILES;
    boundedQuery.nameMode    = FILESYSTEM_SEARCH_NAME_REGEX;
    boundedQuery.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    RecordingSearchCallback boundedCallback;
    const HRESULT boundedHr = SearchFallbackEngine::Execute(baseFs.get(), &boundedQuery, &boundedCallback, nullptr);
    state.Require(SUCCEEDED(boundedHr), std::format(L"Host fallback bounded regex should be accepted, got 0x{:08X}.", static_cast<unsigned long>(boundedHr)));
    if (SUCCEEDED(boundedHr))
    {
        state.Require(FindRecordedSearchMatch(boundedCallback.Matches(), L"aaa.txt") != nullptr, L"Host fallback bounded regex should match aaa.txt.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"host_fallback_regex_oversized_content_reports_overflow",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"host_fallback_regex_oversized", caseRoot), L"Failed to prepare host_fallback_regex_oversized root.");

    constexpr size_t kRegexHelperLimitForTest = 5u * 1024u * 1024u;
    std::string payload(kRegexHelperLimitForTest + 64u, 'a');
    payload.replace(0u, 6u, "needle");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"large.txt", payload), L"Failed to create large.txt.");

    const std::wstring rootText       = caseRoot.wstring();
    const std::wstring namePattern    = L"*.txt";
    const std::wstring contentPattern = L"needle";

    FileSystemSearchQuery query{};
    query.sizeBytes      = sizeof(FileSystemSearchQuery);
    query.rootPath       = rootText.c_str();
    query.namePattern    = namePattern.c_str();
    query.contentPattern = contentPattern.c_str();
    query.flags          = FILESYSTEM_SEARCH_INCLUDE_FILES;
    query.nameMode       = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode    = FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX;

    RecordingSearchCallback callback;
    const HRESULT hr = SearchFallbackEngine::Execute(baseFs.get(), &query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Host fallback oversized regex content search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(callback.Matches().empty(), L"Host fallback oversized regex content search should report no clean match under the overflow policy.");

    const auto progress                     = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Host fallback oversized regex content missing completed progress.");
    if (completed != nullptr)
    {
        state.Require((completed->warningFlags & FILESYSTEM_SEARCH_WARNING_OVERFLOW) != 0u,
                      std::format(L"Host fallback oversized regex content should report OVERFLOW. warnings=0x{:08X}.", completed->warningFlags));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"host_fallback_regex_cancel_before_after_name_candidate",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"host_fallback_regex_cancel_candidate", caseRoot),
                  L"Failed to prepare host_fallback_regex_cancel_candidate root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"target.txt", "target"), L"Failed to create target.txt.");

    const std::wstring rootText    = caseRoot.wstring();
    const std::wstring namePattern = LR"(^target.*\.txt$)";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = FILESYSTEM_SEARCH_INCLUDE_FILES;
    query.nameMode    = FILESYSTEM_SEARCH_NAME_REGEX;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    const auto requireCancelled = [&](std::wstring_view label, uint32_t cancelAfterChecks) noexcept
    {
        RecordingSearchCallback callback(RecordingSearchCallback::Mode::Success, cancelAfterChecks);
        const HRESULT hr = SearchFallbackEngine::Execute(baseFs.get(), &query, &callback, nullptr);
        state.Require(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                      std::format(L"{} regex cancellation should return ERROR_CANCELLED, got 0x{:08X}.", label, static_cast<unsigned long>(hr)));
        state.Require(callback.Matches().empty(), std::format(L"{} regex cancellation should not emit matches.", label));

        const auto progress                     = callback.ProgressSnapshots();
        const RecordedSearchProgress* completed = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
        state.Require(completed != nullptr, std::format(L"{} regex cancellation missing completed progress.", label));
        if (completed != nullptr)
        {
            state.Require(completed->statusHint == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                          std::format(L"{} regex cancellation statusHint mismatch. got=0x{:08X}.", label, static_cast<unsigned long>(completed->statusHint)));
        }
    };

    requireCancelled(L"before-candidate", 3u);
    requireCancelled(L"after-candidate", 4u);

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"host_fallback_search_local_plugin_path_root",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"host_fallback_local_path_root", caseRoot), L"Failed to prepare host_fallback_local_path_root root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"sub"), L"Failed to create sub directory.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"alpha.txt", "needle alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"beta.log", "needle beta"), L"Failed to create beta.log.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"sub" / L"gamma.txt", "needle gamma"), L"Failed to create sub\\gamma.txt.");

    const wil::com_ptr<IFileSystem> wrappedFs = CreatePluginPathMappedRootFileSystem(baseFs, caseRoot);
    state.Require(static_cast<bool>(wrappedFs), L"Failed to create plugin-path mapped filesystem wrapper.");
    if (! wrappedFs)
    {
        return false;
    }

    const std::wstring namePattern    = L"*.txt";
    const std::wstring contentPattern = L"needle";
    const std::wstring rootText       = L"/";

    FileSystemSearchQuery query{};
    query.sizeBytes      = sizeof(FileSystemSearchQuery);
    query.rootPath       = rootText.c_str();
    query.namePattern    = namePattern.c_str();
    query.contentPattern = contentPattern.c_str();
    query.flags          = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES | FILESYSTEM_SEARCH_WANT_SNIPPETS |
                                                              FILESYSTEM_SEARCH_PREFER_INDEX);
    query.nameMode       = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode    = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;
    query.maxSnippetCharacters = 48u;

    RecordingSearchCallback callback;
    const HRESULT hr = SearchFallbackEngine::Execute(wrappedFs.get(), &query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Host fallback local path-root search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    const auto matches = callback.Matches();
    state.Require(matches.size() == 2u, std::format(L"Host fallback local path-root search expected 2 matches, got {}.", matches.size()));

    const RecordedSearchMatch* alphaMatch = FindRecordedSearchMatch(matches, L"alpha.txt");
    state.Require(alphaMatch != nullptr, L"Host fallback local path-root search missing alpha.txt.");
    if (alphaMatch != nullptr)
    {
        state.Require(alphaMatch->previewText.find(L"needle") != std::wstring::npos, L"alpha.txt snippet should include the matched token.");
    }

    const RecordedSearchMatch* gammaMatch = FindRecordedSearchMatch(matches, L"gamma.txt");
    state.Require(gammaMatch != nullptr, L"Host fallback local path-root search missing gamma.txt.");
    if (gammaMatch != nullptr)
    {
        state.Require(gammaMatch->relativePath == L"sub/gamma.txt", L"gamma.txt relative path should preserve plugin separators.");
        state.Require((gammaMatch->matchedBy & FILESYSTEM_SEARCH_MATCH_SOURCE_NAME) != 0u, L"gamma.txt should include the name match source.");
        state.Require((gammaMatch->matchedBy & FILESYSTEM_SEARCH_MATCH_SOURCE_CONTENT) != 0u, L"gamma.txt should include the content match source.");
    }

    const auto progress                     = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Host fallback local path-root search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SCAN, L"Host fallback local path-root search should report scan backend.");
        state.Require((completed->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u,
                      L"Host fallback local path-root search should report degraded-no-index when PREFER_INDEX is requested.");
        state.Require(completed->matchedEntries == 2u, std::format(L"Completed progress matchedEntries mismatch. got={}", completed->matchedEntries));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"host_fallback_search_content_degraded_without_io",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"host_fallback_no_io", caseRoot), L"Failed to prepare host_fallback_no_io root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"plain.txt", "needle"), L"Failed to create plain.txt.");

    const wil::com_ptr<IFileSystem> wrappedFs = CreatePluginPathMappedRootFileSystemNoIO(baseFs, caseRoot);
    state.Require(static_cast<bool>(wrappedFs), L"Failed to create no-IO plugin-path mapped filesystem wrapper.");
    if (! wrappedFs)
    {
        return false;
    }

    const std::wstring rootText       = L"/";
    const std::wstring contentPattern = L"needle";

    FileSystemSearchQuery query{};
    query.sizeBytes      = sizeof(FileSystemSearchQuery);
    query.rootPath       = rootText.c_str();
    query.contentPattern = contentPattern.c_str();
    query.flags          = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode       = FILESYSTEM_SEARCH_NAME_DISABLED;
    query.contentMode    = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;

    RecordingSearchCallback callback;
    const HRESULT hr = SearchFallbackEngine::Execute(wrappedFs.get(), &query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Host fallback no-IO content search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(callback.Matches().empty(), L"Host fallback no-IO content search should not report matches.");

    const auto progress                     = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Host fallback no-IO content search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require((completed->warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT) != 0u,
                      L"Host fallback no-IO content search should report degraded-no-content.");
        state.Require(completed->statusHint == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                      std::format(L"Host fallback no-IO content search statusHint mismatch. hr=0x{:08X}", static_cast<unsigned long>(completed->statusHint)));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"host_fallback_search_access_denied_warning",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"host_fallback_access_denied", caseRoot), L"Failed to prepare host_fallback_access_denied root.");
    state.Require(SelfTest::EnsureDirectory(caseRoot / L"blocked"), L"Failed to create blocked directory.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"allowed.txt", "allowed"), L"Failed to create allowed.txt.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"blocked" / L"hidden.txt", "hidden"), L"Failed to create blocked\\hidden.txt.");

    ReadDirectoryTestBehavior behavior{};
    behavior.targetPath = caseRoot / L"blocked";
    behavior.forcedHr   = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);

    const wil::com_ptr<IFileSystem> wrappedFs = CreateReadDirectoryBehaviorFileSystem(baseFs, behavior);
    state.Require(static_cast<bool>(wrappedFs), L"Failed to create access-denied ReadDirectoryInfo wrapper.");
    if (! wrappedFs)
    {
        return false;
    }

    const std::wstring rootText    = caseRoot.wstring();
    const std::wstring namePattern = L"*.txt";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;

    RecordingSearchCallback callback;
    const HRESULT hr = SearchFallbackEngine::Execute(wrappedFs.get(), &query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Host fallback access-denied search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const auto matches = callback.Matches();
    state.Require(matches.size() == 1u, std::format(L"Host fallback access-denied search expected 1 match, got {}.", matches.size()));
    state.Require(FindRecordedSearchMatch(matches, L"allowed.txt") != nullptr, L"Host fallback access-denied search missing allowed.txt.");
    state.Require(FindRecordedSearchMatch(matches, L"hidden.txt") == nullptr, L"Host fallback access-denied search should skip blocked\\hidden.txt.");

    const auto progress                     = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Host fallback access-denied search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SCAN, L"Host fallback access-denied search should report scan backend.");
        state.Require((completed->warningFlags & FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED) != 0u,
                      L"Host fallback access-denied search should report access-denied-skipped.");
        state.Require(completed->matchedEntries == 1u,
                      std::format(L"Host fallback access-denied search matchedEntries mismatch. got={}", completed->matchedEntries));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"host_fallback_search_short_read_and_cancel",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"host_fallback_short_read_cancel", caseRoot), L"Failed to prepare host_fallback_short_read_cancel root.");

    std::string payload(512u * 1024u, 'a');
    payload.replace(260u * 1024u, 6u, "needle");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"big.txt", payload), L"Failed to create big.txt.");

    const wil::com_ptr<IFileSystem> wrappedFs = CreateShortReadFileSystem(baseFs, caseRoot, 1024u, static_cast<DWORD>(SelfTest::ScaleTimeout(1)));
    state.Require(static_cast<bool>(wrappedFs), L"Failed to create short-read filesystem wrapper.");
    if (! wrappedFs)
    {
        return false;
    }

    const std::wstring rootText       = caseRoot.wstring();
    const std::wstring contentPattern = L"needle";

    FileSystemSearchQuery query{};
    query.sizeBytes              = sizeof(FileSystemSearchQuery);
    query.rootPath               = rootText.c_str();
    query.contentPattern         = contentPattern.c_str();
    query.flags                  = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode               = FILESYSTEM_SEARCH_NAME_DISABLED;
    query.contentMode            = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;
    query.maxContentBytesPerFile = SearchTextHelpers::kDefaultContentBytesPerFile;

    RecordingSearchCallback callback(RecordingSearchCallback::Mode::Success, 4u);
    const HRESULT hr = SearchFallbackEngine::Execute(wrappedFs.get(), &query, &callback, nullptr);
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Host fallback short-read cancel search expected ERROR_CANCELLED. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(callback.CancelCalls() >= 4u,
                  std::format(L"Host fallback short-read cancel search expected repeated cancel polling. calls={}", callback.CancelCalls()));

    const auto progress                       = callback.ProgressSnapshots();
    const RecordedSearchProgress* contentScan = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_CONTENT_SCAN);
    state.Require(contentScan != nullptr, L"Host fallback short-read cancel search should enter content-scan phase.");
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Host fallback short-read cancel search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(
            completed->statusHint == HRESULT_FROM_WIN32(ERROR_CANCELLED),
            std::format(L"Host fallback short-read cancel search statusHint mismatch. hr=0x{:08X}", static_cast<unsigned long>(completed->statusHint)));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"host_fallback_search_dummy_name_only",
                  [&](SelfTest::CaseState& state) noexcept
{
    state.Require(dummyFs && dummyIo && dummyOps, L"Dummy filesystem setup is incomplete.");
    if (! dummyFs || ! dummyIo || ! dummyOps)
    {
        return false;
    }

    const std::filesystem::path rootPath = std::filesystem::path(L"/") / L"search_dummy_name_only";
    const std::filesystem::path filePath = rootPath / L"probe.txt";
    state.Require(EnsureDirectoryExistsFsOps(dummyOps, rootPath), L"Failed to create dummy search root.");
    state.Require(WriteFileTextFsIo(dummyIo, filePath, "dummy-search"), L"Failed to create dummy probe.txt.");

    const std::wstring rootText    = ToPluginPathText(rootPath);
    const std::wstring namePattern = L"probe.txt";

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = rootText.c_str();
    query.namePattern = namePattern.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_LITERAL;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;
    query.maxResults  = 1u;

    RecordingSearchCallback callback;
    const HRESULT hr = SearchFallbackEngine::Execute(dummyFs.get(), &query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Host fallback dummy name-only search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    const auto matches = callback.Matches();
    state.Require(matches.size() == 1u, std::format(L"Host fallback dummy name-only search expected 1 match, got {}.", matches.size()));
    const RecordedSearchMatch* probeMatch = FindRecordedSearchMatch(matches, L"probe.txt");
    state.Require(probeMatch != nullptr, L"Host fallback dummy name-only search missing probe.txt.");

    const auto progress                     = callback.ProgressSnapshots();
    const RecordedSearchProgress* completed = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(completed != nullptr, L"Host fallback dummy name-only search missing completed progress.");
    if (completed != nullptr)
    {
        state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SCAN, L"Host fallback dummy name-only search should report scan backend.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"host_fallback_search_7z_name_only",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance archiveCreated{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltin7zFileSystemId, {}, archiveCreated);
    state.Require(SUCCEEDED(createHr) && archiveCreated.fileSystem,
                  std::format(L"Host fallback 7z name-only search: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! archiveCreated.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemInitialize> init;
    const HRESULT initQiHr = archiveCreated.fileSystem->QueryInterface(IID_PPV_ARGS(init.put()));
    state.Require(SUCCEEDED(initQiHr) && init, L"Host fallback 7z name-only search: missing IFileSystemInitialize.");
    if (FAILED(initQiHr) || ! init)
    {
        return false;
    }

    const std::filesystem::path archivePath = GetWorkspaceRootFromSourcePath() / L"Plugins" / L"FileSystem7z" / L"Tests" / L"Tests.zip";
    state.Require(SelfTest::PathExists(archivePath), std::format(L"Host fallback 7z name-only search: fixture archive missing: {}", archivePath.wstring()));
    if (! SelfTest::PathExists(archivePath))
    {
        return false;
    }

    const HRESULT initHr = init->Initialize(archivePath.c_str(), nullptr);
    state.Require(SUCCEEDED(initHr), std::format(L"Host fallback 7z name-only search: Initialize failed. hr=0x{:08X}", static_cast<unsigned long>(initHr)));
    if (FAILED(initHr))
    {
        return false;
    }

    const auto filePathOpt = FindFirstRegularEntryPath(archiveCreated.fileSystem, L"/");
    state.Require(filePathOpt.has_value(), L"Host fallback 7z name-only search: no regular file found in fixture archive.");
    if (! filePathOpt.has_value())
    {
        return false;
    }

    const std::filesystem::path regularPath(filePathOpt.value());
    const std::wstring leafName = regularPath.filename().wstring();
    state.Require(! leafName.empty(), L"Host fallback 7z name-only search: failed to derive a leaf name from the fixture path.");
    if (leafName.empty())
    {
        return false;
    }

    FileSystemSearchQuery query{};
    query.sizeBytes   = sizeof(FileSystemSearchQuery);
    query.rootPath    = L"/";
    query.namePattern = leafName.c_str();
    query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    query.nameMode    = FILESYSTEM_SEARCH_NAME_LITERAL;
    query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;
    query.maxResults  = 1u;

    RecordingSearchCallback callback;
    const HRESULT hr = SearchFallbackEngine::Execute(archiveCreated.fileSystem.get(), &query, &callback, nullptr);
    state.Require(SUCCEEDED(hr), std::format(L"Host fallback 7z name-only search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    const auto matches = callback.Matches();
    state.Require(! matches.empty(), L"Host fallback 7z name-only search expected at least one match.");
    if (! matches.empty())
    {
        state.Require(std::wstring_view(matches.front().displayName) == leafName, L"Host fallback 7z name-only search matched the wrong leaf.");
    }

    return state.failure.empty();
});

const auto runRemoteFallbackNameOnlySmoke = [&](std::wstring_view caseName,
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
        const ResolvedRemoteProfile resolved = ResolveRemoteConnectionProfile(envVarName, defaultProfileName, pluginId);
        state.Require(resolved.profile != nullptr, L"Remote fallback smoke: connection profile not found after gates passed.");
        if (! resolved.profile)
        {
            return false;
        }

        CreatedFileSystemInstance remoteCreated{};
        const HRESULT createHr = TryCreateFileSystemInstance(pluginId, {}, remoteCreated);
        state.Require(SUCCEEDED(createHr) && remoteCreated.fileSystem,
                      std::format(L"Remote fallback smoke: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! remoteCreated.fileSystem)
        {
            return false;
        }

        const std::wstring initialPath = NormalizePluginPathForSelfTest(resolved.profile->initialPath);
        const std::wstring rootText    = MakeConnectionPathForSelfTest(resolved.profileName, initialPath);
        state.Require(! rootText.empty(), L"Remote fallback smoke: failed to build connection-scoped search root.");
        if (rootText.empty())
        {
            return false;
        }

        FileSystemSearchQuery query{};
        query.sizeBytes   = sizeof(FileSystemSearchQuery);
        query.rootPath    = rootText.c_str();
        query.namePattern = L"*";
        query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_INCLUDE_FILES | FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES);
        query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
        query.contentMode = FILESYSTEM_SEARCH_CONTENT_DISABLED;
        query.maxResults  = 1u;

        RecordingSearchCallback callback;
        const HRESULT hr = SearchFallbackEngine::Execute(remoteCreated.fileSystem.get(), &query, &callback, nullptr);
        state.Require(SUCCEEDED(hr), std::format(L"Remote fallback smoke search failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
        state.Require(callback.ProgressCalls() >= 1u, L"Remote fallback smoke search expected progress callbacks.");

        const auto progress                     = callback.ProgressSnapshots();
        const RecordedSearchProgress* completed = FindRecordedSearchProgress(progress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
        state.Require(completed != nullptr, L"Remote fallback smoke search missing completed progress.");
        if (completed != nullptr)
        {
            state.Require(completed->backend == FILESYSTEM_SEARCH_BACKEND_SCAN, L"Remote fallback smoke search should report scan backend.");
        }

        return state.failure.empty();
    });
};

runRemoteFallbackNameOnlySmoke(L"host_fallback_search_remote_ftp_name_only", L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kBuiltinFtpFileSystemId);

SelfTest::RunCase(options,
                  suite,
                  L"search_service_missing_pipe_retry_is_cancellable",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_missing_pipe_retry_cancel", caseRoot),
                  L"Failed to prepare search_service_missing_pipe_retry_cancel root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring previousPipeOverride  = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousRetryOverride = GetEnvVarTrimmed(SearchServiceBroker::kClientMissingPipeRetryMsEnvVar);
    const std::wstring pipeName              = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0,
                  L"Failed to override the missing search service pipe.");
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kClientMissingPipeRetryMsEnvVar, L"5000") != 0,
                  L"Failed to stretch the missing-pipe retry window.");
    const auto restoreOverrides = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(::SetEnvironmentVariableW(
            SearchServiceBroker::kPipeNameEnvVar, previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str()));
        static_cast<void>(::SetEnvironmentVariableW(
            SearchServiceBroker::kClientMissingPipeRetryMsEnvVar, previousRetryOverride.empty() ? nullptr : previousRetryOverride.c_str()));
    });

    struct CancelContext final
    {
        ULONGLONG cancelAtTick = 0u;
    } cancelContext{.cancelAtTick = ::GetTickCount64() + 50u};
    const auto cancelCheck = [](void* cookie) noexcept -> HRESULT
    {
        const auto* context = static_cast<const CancelContext*>(cookie);
        return context != nullptr && ::GetTickCount64() >= context->cancelAtTick ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
    };

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;

    std::vector<LocalSearchIndexCore::Candidate> candidates;
    LocalSearchIndexCore::QueryStats stats{};
    const auto started = std::chrono::steady_clock::now();
    const HRESULT hr = SearchServiceBroker::Query(request, nullptr, nullptr, cancelCheck, &cancelContext, candidates, &stats);
    const auto elapsedMs = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - started).count();
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Missing-pipe retry cancellation returned 0x{:08X} instead of ERROR_CANCELLED.", static_cast<unsigned long>(hr)));
    state.Require(elapsedMs < 1000, std::format(L"Missing-pipe retry cancellation took {}ms; expected prompt cancellation.", elapsedMs));
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_service_root_rejection_does_not_arm_transport_cooldown",
                  [&](SelfTest::CaseState& state) noexcept
{
#if ! defined(RS_SEARCH_TEST_HOOKS)
    return state.Skip(L"Search service test hooks are unavailable in this build configuration.");
#else
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_root_rejection_cooldown", caseRoot),
                  L"Failed to prepare search_service_root_rejection_cooldown root.");
    const std::filesystem::path rejectedRoot = caseRoot / L"service-reject-root-a";
    const std::filesystem::path validRoot    = caseRoot / L"valid-root-b";
    state.Require(SelfTest::WriteTextFile(rejectedRoot / L"rejected.txt", "rejected"), L"Failed to create rejected-root fixture.");
    state.Require(SelfTest::WriteTextFile(validRoot / L"valid.txt", "valid"), L"Failed to create valid-root fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    const auto restorePipeOverride = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(::SetEnvironmentVariableW(
            SearchServiceBroker::kPipeNameEnvVar, previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str()));
    });

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::filesystem::path storageRoot = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-service-store");
    const std::wstring extraArguments =
        std::format(L"--storage-root=\"{}\" --store-backend=snapshot --test-reject-marked-root", storageRoot.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, false, extraArguments), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem, L"Failed to create isolated local file system instance.");
    wil::com_ptr<IInformations> info;
    wil::com_ptr<IFileSystemSearch> search;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    state.Require(CreateFileSystemSearch(created.fileSystem, search), L"Isolated local file system instance missing IFileSystemSearch.");
    if (! info || ! search)
    {
        return false;
    }
    state.Require(SUCCEEDED(info->SetConfiguration("{\"searchBackendPreference\":\"service\"}")),
                  L"Failed to configure service-backed search.");

    const auto runSearch = [&](const std::filesystem::path& searchRoot, RecordingSearchCallback& callback) noexcept -> HRESULT
    {
        const std::wstring rootText = searchRoot.wstring();
        FileSystemSearchQuery query{};
        query.sizeBytes   = sizeof(FileSystemSearchQuery);
        query.rootPath    = rootText.c_str();
        query.namePattern = L"*.txt";
        query.flags       = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
        query.nameMode    = FILESYSTEM_SEARCH_NAME_WILDCARD;
        return search->Search(&query, &callback, nullptr);
    };

    RecordingSearchCallback rejectedCallback;
    const HRESULT rejectedHr = runSearch(rejectedRoot, rejectedCallback);
    state.Require(SUCCEEDED(rejectedHr),
                  std::format(L"Root-rejected service search should fall back successfully. hr=0x{:08X}", static_cast<unsigned long>(rejectedHr)));
    const auto rejectedProgress = rejectedCallback.ProgressSnapshots();
    const RecordedSearchProgress* rejectedCompleted = FindRecordedSearchProgress(rejectedProgress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(rejectedCompleted != nullptr, L"Root-rejected service search missing completion progress.");
    if (rejectedCompleted != nullptr)
    {
        state.Require(rejectedCompleted->backend != FILESYSTEM_SEARCH_BACKEND_SERVICE, L"Rejected root should complete on a fallback backend.");
        state.Require((rejectedCompleted->warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_ROOT_REJECTED) != 0u,
                      L"Rejected root should report SERVICE_ROOT_REJECTED.");
        state.Require((rejectedCompleted->warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) == 0u,
                      L"Healthy-service root rejection must not report SERVICE_UNAVAILABLE.");
    }

    RecordingSearchCallback validCallback;
    const HRESULT validHr = runSearch(validRoot, validCallback);
    state.Require(SUCCEEDED(validHr), std::format(L"Immediate valid-root service search failed. hr=0x{:08X}", static_cast<unsigned long>(validHr)));
    const auto validProgress = validCallback.ProgressSnapshots();
    const RecordedSearchProgress* validCompleted = FindRecordedSearchProgress(validProgress, FILESYSTEM_SEARCH_PHASE_COMPLETED);
    state.Require(validCompleted != nullptr, L"Immediate valid-root service search missing completion progress.");
    if (validCompleted != nullptr)
    {
        state.Require(validCompleted->backend == FILESYSTEM_SEARCH_BACKEND_SERVICE,
                      L"Immediate valid root B should still use the service backend after root A was rejected.");
        state.Require((validCompleted->warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) == 0u,
                      L"Immediate valid root B should not inherit a service-unavailable cooldown.");
    }
    return state.failure.empty();
#endif
});

SelfTest::RunCase(options,
                  suite,
                  L"search_service_transient_parent_failure_is_cached_and_warns_batch_and_completion",
                  [&](SelfTest::CaseState& state) noexcept
{
#if ! defined(RS_SEARCH_TEST_HOOKS)
    return state.Skip(L"Search service test hooks are unavailable in this build configuration.");
#else
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"search_service_transient_parent_batch_cache", caseRoot),
                  L"Failed to prepare search_service_transient_parent_batch_cache root.");
    const std::filesystem::path transientParent = caseRoot / L"transient-auth";
    const std::filesystem::path allowedParent   = caseRoot / L"z-allowed";
    state.Require(SelfTest::EnsureDirectory(transientParent), L"Failed to create transient authorization parent.");
    state.Require(SelfTest::EnsureDirectory(allowedParent), L"Failed to create allowed authorization parent.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path sqlitePath = caseRoot.parent_path() / (caseRoot.filename().wstring() + L".sqlite3");
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr), L"Failed to bootstrap transient authorization SQLite fixture.");
    if (FAILED(hr))
    {
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest replaceRequest{};
    replaceRequest.rootPath       = caseRoot.wstring();
    replaceRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replaceRequest.journalId      = 1u;
    replaceRequest.nextUsn        = 1u;
    replaceRequest.state          = SqliteIndexStore::kVolumeStateReady;
    replaceRequest.entries.push_back({.fileIdLow = 1u,
                                      .fullPath  = caseRoot.wstring(),
                                      .name      = caseRoot.filename().wstring(),
                                      .attributes = FILE_ATTRIBUTE_DIRECTORY});
    replaceRequest.entries.push_back({.fileIdLow = 2u,
                                      .parentIdLow = 1u,
                                      .fullPath  = transientParent.wstring(),
                                      .name      = transientParent.filename().wstring(),
                                      .attributes = FILE_ATTRIBUTE_DIRECTORY});
    replaceRequest.entries.push_back({.fileIdLow = 3u,
                                      .parentIdLow = 1u,
                                      .fullPath  = allowedParent.wstring(),
                                      .name      = allowedParent.filename().wstring(),
                                      .attributes = FILE_ATTRIBUTE_DIRECTORY});
    for (uint64_t index = 0u; index < 2u; ++index)
    {
        const std::wstring name = std::format(L"blocked-{}.txt", index);
        replaceRequest.entries.push_back({.fileIdLow = 10u + index,
                                          .parentIdLow = 2u,
                                          .fullPath  = (transientParent / name).wstring(),
                                          .name      = name,
                                          .attributes = FILE_ATTRIBUTE_ARCHIVE});
    }
    for (uint64_t index = 0u; index < 260u; ++index)
    {
        const std::wstring name = std::format(L"allowed-{:03}.txt", index);
        replaceRequest.entries.push_back({.fileIdLow = 100u + index,
                                          .parentIdLow = 3u,
                                          .fullPath  = (allowedParent / name).wstring(),
                                          .name      = name,
                                          .attributes = FILE_ATTRIBUTE_ARCHIVE});
    }
    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), replaceRequest, &replaceResult);
    state.Require(SUCCEEDED(hr), L"Failed to seed transient authorization SQLite fixture.");
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring previousDirectOverride = GetEnvVarTrimmed(L"REDSALAMANDER_TEST_ALLOW_DIRECT_SQLITE_WITHOUT_JOURNAL");
    const std::wstring pipeName = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0, L"Failed to override the search service pipe.");
    state.Require(::SetEnvironmentVariableW(L"REDSALAMANDER_TEST_ALLOW_DIRECT_SQLITE_WITHOUT_JOURNAL", L"1") != 0,
                  L"Failed to enable direct SQLite test currentness override.");
    const auto restoreOverrides = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(::SetEnvironmentVariableW(
            SearchServiceBroker::kPipeNameEnvVar, previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str()));
        static_cast<void>(::SetEnvironmentVariableW(L"REDSALAMANDER_TEST_ALLOW_DIRECT_SQLITE_WITHOUT_JOURNAL",
                                                    previousDirectOverride.empty() ? nullptr : previousDirectOverride.c_str()));
    });

    const std::filesystem::path perfPath = SelfTest::GetPerfArtifactPath(L"perf_metrics.jsonl");
    std::error_code perfEc;
    const std::uintmax_t metricOffset = std::filesystem::exists(perfPath, perfEc) ? std::filesystem::file_size(perfPath, perfEc) : 0u;

    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring extraArguments = std::format(
        L"--store-backend=sqlite --sqlite-path=\"{}\" --test-fail-marked-client-directory-open-bad-netpath", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 0u, SearchServiceBroker::kProtocolVersion, true, serviceError, true, extraArguments), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.namePattern        = L"*.txt";
    request.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    request.flags              = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_RECURSIVE | FILESYSTEM_SEARCH_INCLUDE_FILES);
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;
    LocalSearchIndexCore::QueryStats stats{};
    std::vector<LocalSearchIndexCore::Candidate> candidates;
    hr = SearchServiceBroker::Query(request, nullptr, nullptr, nullptr, nullptr, candidates, &stats);
    state.Require(SUCCEEDED(hr), std::format(L"Transient authorization query failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(candidates.size() == 260u, std::format(L"Expected 260 authorized candidates, got {}.", candidates.size()));
    state.Require((stats.warningFlags & FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED) != 0u,
                  L"Transient authorization query completion should report ACCESS_DENIED_SKIPPED.");

    std::string serviceOutput;
    state.Require(service.ShutdownAndWaitForExitAndCapture(static_cast<DWORD>(SelfTest::ScaleTimeout(5'000)), serviceOutput, serviceError), serviceError);
    const std::string warningToken = "warnings=0x00000004";
    const size_t batchPosition = serviceOutput.find("Query batch sent");
    const size_t completionPosition = serviceOutput.find("Query completed");
    state.Require(batchPosition != std::string::npos && serviceOutput.find(warningToken, batchPosition) != std::string::npos,
                  L"QUERY_BATCH_SENT should carry the accumulated ACCESS_DENIED_SKIPPED warning.");
    state.Require(completionPosition != std::string::npos && serviceOutput.find(warningToken, completionPosition) != std::string::npos,
                  L"Query completion should carry the accumulated ACCESS_DENIED_SKIPPED warning.");

    size_t transientAttemptMetrics = 0u;
    std::ifstream perfFile(perfPath, std::ios::binary);
    if (perfFile)
    {
        if (metricOffset != 0u)
        {
            perfFile.seekg(static_cast<std::streamoff>(metricOffset), std::ios::beg);
        }
        std::string line;
        while (std::getline(perfFile, line))
        {
            if (line.find("\"metric\":\"search.service.authorization.transient_parent_failures\"") != std::string::npos)
            {
                ++transientAttemptMetrics;
            }
        }
    }
    state.Require(transientAttemptMetrics == 1u,
                  std::format(L"Two candidates under one transient parent should perform and log one authorization attempt; metric count={}",
                              transientAttemptMetrics));
    return state.failure.empty();
#endif
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_sqlite_generation_probe_skips_steady_state_and_detects_bump",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"local_search_sqlite_generation_probe", caseRoot),
                  L"Failed to prepare local_search_sqlite_generation_probe root.");
    const std::filesystem::path sqlitePath = caseRoot.parent_path() / (caseRoot.filename().wstring() + L".sqlite3");
    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(sqlitePath.wstring(), &bootstrapInfo);
    state.Require(SUCCEEDED(hr), L"Failed to bootstrap generation-probe SQLite fixture.");
    if (FAILED(hr))
    {
        return false;
    }

    const auto makeRequest = [&](bool includeBeta) noexcept
    {
        SqliteIndexStore::ReplaceVolumeRequest request{};
        request.rootPath       = caseRoot.wstring();
        request.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
        request.journalId      = 1u;
        request.nextUsn        = includeBeta ? 2u : 1u;
        request.state          = SqliteIndexStore::kVolumeStateReady;
        request.entries.push_back({.fileIdLow = 1u,
                                   .fullPath  = caseRoot.wstring(),
                                   .name      = caseRoot.filename().wstring(),
                                   .attributes = FILE_ATTRIBUTE_DIRECTORY});
        request.entries.push_back({.fileIdLow = 2u,
                                   .parentIdLow = 1u,
                                   .fullPath  = (caseRoot / L"alpha.txt").wstring(),
                                   .name      = L"alpha.txt",
                                   .attributes = FILE_ATTRIBUTE_ARCHIVE});
        if (includeBeta)
        {
            request.entries.push_back({.fileIdLow = 3u,
                                       .parentIdLow = 1u,
                                       .fullPath  = (caseRoot / L"beta.txt").wstring(),
                                       .name      = L"beta.txt",
                                       .attributes = FILE_ATTRIBUTE_ARCHIVE});
        }
        return request;
    };

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    SqliteIndexStore::ReplaceVolumeRequest initialRequest = makeRequest(false);
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), initialRequest, &replaceResult);
    state.Require(SUCCEEDED(hr), L"Failed to seed initial generation-probe volume.");
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring previousDirectOverride = GetEnvVarTrimmed(L"REDSALAMANDER_TEST_ALLOW_DIRECT_SQLITE_WITHOUT_JOURNAL");
    state.Require(::SetEnvironmentVariableW(L"REDSALAMANDER_TEST_ALLOW_DIRECT_SQLITE_WITHOUT_JOURNAL", L"1") != 0,
                  L"Failed to enable direct SQLite test currentness override.");
    const auto restoreDirectOverride = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(::SetEnvironmentVariableW(L"REDSALAMANDER_TEST_ALLOW_DIRECT_SQLITE_WITHOUT_JOURNAL",
                                                    previousDirectOverride.empty() ? nullptr : previousDirectOverride.c_str()));
    });

    LocalSearchIndexCore::Repository repository({.persistentStoreKind = LocalSearchIndexCore::PersistentStoreKind::Sqlite,
                                                  .sqliteDatabasePath  = sqlitePath.wstring(),
                                                  .sqliteAuthoritative = true});
    LocalSearchIndexCore::QueryPlan plan{};
    plan.rootPath           = caseRoot.wstring();
    plan.namePattern        = L"*.txt";
    plan.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    plan.recursive          = true;
    plan.includeFiles       = true;
    plan.includeDirectories = false;

    LocalSearchIndexCore::QueryStats warmStats{};
    std::vector<LocalSearchIndexCore::Candidate> warmCandidates;
    hr = repository.Query(plan, nullptr, nullptr, warmCandidates, &warmStats);
    state.Require(SUCCEEDED(hr) && warmStats.usedSqliteStore, L"Generation-probe warm-up query should use SQLite.");

    const std::filesystem::path perfPath = SelfTest::GetPerfArtifactPath(L"perf_metrics.jsonl");
    const auto currentPerfOffset = [&]() noexcept -> std::uintmax_t
    {
        std::error_code ec;
        return std::filesystem::exists(perfPath, ec) ? std::filesystem::file_size(perfPath, ec) : 0u;
    };
    const auto countMetricAfter = [&](std::uintmax_t offset, std::string_view metric) noexcept -> size_t
    {
        size_t count = 0u;
        std::ifstream file(perfPath, std::ios::binary);
        if (file)
        {
            if (offset != 0u)
            {
                file.seekg(static_cast<std::streamoff>(offset), std::ios::beg);
            }
            std::string line;
            const std::string token = std::format("\"metric\":\"{}\"", metric);
            while (std::getline(file, line))
            {
                if (line.find(token) != std::string::npos)
                {
                    ++count;
                }
            }
        }
        return count;
    };

    const std::uintmax_t steadyOffset = currentPerfOffset();
    LocalSearchIndexCore::QueryStats steadyStats{};
    std::vector<LocalSearchIndexCore::Candidate> steadyCandidates;
    hr = repository.Query(plan, nullptr, nullptr, steadyCandidates, &steadyStats);
    state.Require(SUCCEEDED(hr), L"Steady-state generation-probe query failed.");
    state.Require(countMetricAfter(steadyOffset, "search.backend.sqlite.store_generation_probe_opens") == 0u,
                  L"Repeated steady-state query should perform zero SQLite generation-probe opens after warm-up.");
    state.Require(countMetricAfter(steadyOffset, "search.backend.sqlite.store_generation_probe_skips") == 1u,
                  L"Repeated steady-state query should record one unchanged-file stamp skip.");

    SqliteIndexStore::ReplaceVolumeRequest rotatedRequest = makeRequest(true);
    hr = SqliteIndexStore::ReplaceVolume(sqlitePath.wstring(), rotatedRequest, &replaceResult);
    state.Require(SUCCEEDED(hr), L"Failed to bump the generation-probe SQLite store.");
    if (FAILED(hr))
    {
        return false;
    }

    const std::uintmax_t bumpedOffset = currentPerfOffset();
    LocalSearchIndexCore::QueryStats bumpedStats{};
    std::vector<LocalSearchIndexCore::Candidate> bumpedCandidates;
    hr = repository.Query(plan, nullptr, nullptr, bumpedCandidates, &bumpedStats);
    state.Require(SUCCEEDED(hr), L"Generation-bump query failed.");
    state.Require(CollectIndexedCandidateNames(bumpedCandidates) == std::vector<std::wstring>{L"alpha.txt", L"beta.txt"},
                  L"Generation bump should be detected within one query.");
    state.Require(countMetricAfter(bumpedOffset, "search.backend.sqlite.store_generation_probe_opens") == 1u,
                  L"One changed-file validation should perform at most one generation-probe open per Enumerate.");
    state.Require(countMetricAfter(bumpedOffset, "search.backend.sqlite.store_generation_refreshes") == 1u,
                  L"Generation bump should refresh cached persistent-store metadata in the same query.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"local_search_junction_alias_hydration_indexes_alias_only_descendant_once",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"local_search_junction_alias_hydration", caseRoot),
                  L"Failed to prepare local_search_junction_alias_hydration root.");
    const std::filesystem::path targetPath = caseRoot / L"a-target";
    const std::filesystem::path aliasPath  = caseRoot / L"z-alias";
    state.Require(SelfTest::EnsureDirectory(targetPath), L"Failed to create junction target directory.");
    if (! state.failure.empty())
    {
        return false;
    }
    if (! TryCreateDirectoryJunction(aliasPath, targetPath))
    {
        const DWORD error = ::GetLastError();
        if (error == ERROR_PRIVILEGE_NOT_HELD)
        {
            return state.Skip(L"Directory junction creation requires a privilege unavailable on this machine.");
        }
        state.Require(false, std::format(L"Failed to create directory junction. error={}", error));
        return false;
    }

    const std::wstring previousForceTraversal = GetEnvVarTrimmed(L"REDSALAMANDER_TEST_FORCE_NTFS_TRAVERSAL_SEED");
    state.Require(::SetEnvironmentVariableW(L"REDSALAMANDER_TEST_FORCE_NTFS_TRAVERSAL_SEED", L"1") != 0,
                  L"Failed to force traversal seeding for the junction alias test.");
    const auto restoreTraversal = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(::SetEnvironmentVariableW(
            L"REDSALAMANDER_TEST_FORCE_NTFS_TRAVERSAL_SEED", previousForceTraversal.empty() ? nullptr : previousForceTraversal.c_str()));
    });

    LocalSearchIndexCore::Repository repository;
    LocalSearchIndexCore::QueryPlan plan{};
    plan.rootPath           = caseRoot.wstring();
    plan.namePattern        = L"needle_alias.txt";
    plan.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    plan.recursive          = true;
    plan.includeFiles       = true;
    plan.includeDirectories = false;
    LocalSearchIndexCore::QueryStats initialStats{};
    std::vector<LocalSearchIndexCore::Candidate> initialCandidates;
    HRESULT hr = repository.Query(plan, nullptr, nullptr, initialCandidates, &initialStats);
    state.Require(SUCCEEDED(hr), L"Initial junction alias index build failed.");
    state.Require(initialCandidates.empty(), L"Initial junction alias fixture should not contain needle_alias.txt.");
    state.Require(SelfTest::WriteTextFile(targetPath / L"needle_alias.txt", "alias"), L"Failed to create alias-only descendant after initial hydration.");
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::QueryStats hydratedStats{};
    std::vector<LocalSearchIndexCore::Candidate> hydratedCandidates;
    hr = repository.HydrateRootAndQueryForTests(plan, hydratedCandidates, &hydratedStats);
    state.Require(SUCCEEDED(hr), std::format(L"Junction alias rehydration failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(hydratedCandidates.size() == 1u,
                  std::format(L"Junction alias rehydration should index one physical descendant, got {}.", hydratedCandidates.size()));
    if (hydratedCandidates.size() == 1u)
    {
        state.Require(EqualsIgnoreCase(hydratedCandidates[0].fullPath, (targetPath / L"needle_alias.txt").wstring()),
                      L"Coverage-only alias hydration should retain the canonical path and avoid duplicate alias emission.");
    }
    state.Require(hydratedStats.hardlinkAliasCoverageIncomplete,
                  L"Coverage-only alias hydration should continue to report that alias-path representation is incomplete.");
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sqlite_maintenance_busy_checkpoint_and_callback_skip_observability",
                  [&](SelfTest::CaseState& state) noexcept
{
    std::filesystem::path caseRoot;
    state.Require(PrepareSearchCaseRoot(root, L"sqlite_busy_checkpoint_skip_observability", caseRoot),
                  L"Failed to prepare sqlite_busy_checkpoint_skip_observability root.");
    const std::filesystem::path databasePath = caseRoot / L"index-v2.sqlite3";
    SqliteIndexStore::StoreInfo preparedInfo{};
    std::wstring prepareError;
    state.Require(PrepareSqliteMaintenanceStore(caseRoot, databasePath, preparedInfo, prepareError), prepareError);
    if (! state.failure.empty())
    {
        return false;
    }

    sqlite3* walSeedConnection = nullptr;
    state.Require(OpenSqliteWalSeedConnection(databasePath, walSeedConnection, prepareError), prepareError);
    if (! state.failure.empty())
    {
        return false;
    }
    const auto closeWalSeed = wil::scope_exit([&]() noexcept
    {
        if (walSeedConnection != nullptr)
        {
            static_cast<void>(sqlite3_close(walSeedConnection));
        }
    });

    sqlite3* reader = nullptr;
    const std::u8string databaseUtf8 = databasePath.u8string();
    const int openResult = sqlite3_open_v2(reinterpret_cast<const char*>(databaseUtf8.c_str()), &reader, SQLITE_OPEN_READONLY | SQLITE_OPEN_NOMUTEX, nullptr);
    state.Require(openResult == SQLITE_OK && reader != nullptr, L"Failed to open checkpoint-blocking SQLite reader.");
    if (openResult != SQLITE_OK || reader == nullptr)
    {
        if (reader != nullptr)
        {
            static_cast<void>(sqlite3_close(reader));
        }
        return false;
    }
    const auto closeReader = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(sqlite3_exec(reader, "ROLLBACK;", nullptr, nullptr, nullptr));
        static_cast<void>(sqlite3_close(reader));
    });
    state.Require(sqlite3_exec(reader, "BEGIN; SELECT value FROM meta LIMIT 1;", nullptr, nullptr, nullptr) == SQLITE_OK,
                  L"Failed to hold a read snapshot that makes the final checkpoint busy.");

    LocalSearchIndexCore::SqliteMaintenancePolicy policy{};
    policy.autoCheckpointTargetBytes          = 1u;
    policy.autoCompactionFragmentationPercent = (std::numeric_limits<uint32_t>::max)();
    policy.autoCompactionMinBytes             = (std::numeric_limits<uint64_t>::max)();
    SqliteIndexStore::AutomaticMaintenanceResult maintenance{};
    const HRESULT maintenanceHr = SqliteIndexStore::RunAutomaticMaintenance(databasePath.wstring(), policy, &maintenance);
    state.Require(maintenanceHr == S_FALSE,
                  std::format(L"Busy automatic checkpoint should return S_FALSE. hr=0x{:08X}", static_cast<unsigned long>(maintenanceHr)));
    state.Require(maintenance.after.lastCheckpointUtc.empty(), L"Busy checkpoint must leave lastCheckpointUtc unchanged in the result.");
    SqliteIndexStore::StoreInfo afterBusy{};
    const HRESULT inspectHr = SqliteIndexStore::InspectStore(databasePath.wstring(), afterBusy);
    state.Require(SUCCEEDED(inspectHr), L"Failed to inspect SQLite store after busy checkpoint.");
    state.Require(afterBusy.lastCheckpointUtc.empty(), L"Busy checkpoint must leave persisted lastCheckpointUtc empty.");

    SqliteIndexStore::ReplaceVolumeRequest replaceRequest{};
    replaceRequest.rootPath       = caseRoot.wstring();
    replaceRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    replaceRequest.journalId      = 1u;
    replaceRequest.nextUsn        = 1u;
    replaceRequest.state          = SqliteIndexStore::kVolumeStateReady;
    replaceRequest.entries.push_back({.fileIdLow = 1u,
                                      .fullPath  = caseRoot.wstring(),
                                      .name      = caseRoot.filename().wstring(),
                                      .attributes = FILE_ATTRIBUTE_DIRECTORY});
    replaceRequest.entries.push_back({.fileIdLow = 2u,
                                      .parentIdLow = 1u,
                                      .fullPath  = (caseRoot / L"skip.txt").wstring(),
                                      .name      = L"skip.txt",
                                      .attributes = FILE_ATTRIBUTE_ARCHIVE});
    replaceRequest.entries.push_back({.fileIdLow = 3u,
                                      .parentIdLow = 1u,
                                      .fullPath  = (caseRoot / L"emit.txt").wstring(),
                                      .name      = L"emit.txt",
                                      .attributes = FILE_ATTRIBUTE_ARCHIVE});
    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    HRESULT hr = SqliteIndexStore::ReplaceVolume(databasePath.wstring(), replaceRequest, &replaceResult);
    state.Require(SUCCEEDED(hr), L"Failed to seed callback-skip SQLite fixture.");
    if (FAILED(hr))
    {
        return false;
    }

    struct SkipState final
    {
        uint64_t callbacks = 0u;
    } skipState;
    const auto callback = [](LocalSearchIndexCore::Candidate* candidate, void* cookie) noexcept -> HRESULT
    {
        if (candidate == nullptr || cookie == nullptr)
        {
            return E_POINTER;
        }
        auto& localState = *static_cast<SkipState*>(cookie);
        ++localState.callbacks;
        return localState.callbacks == 1u ? LocalSearchIndexCore::kSkipCandidateHr : S_OK;
    };
    SqliteIndexStore::QueryRequest request{};
    request.rootPath           = caseRoot.wstring();
    request.nameMode           = FILESYSTEM_SEARCH_NAME_DISABLED;
    request.recursive          = true;
    request.includeFiles       = true;
    request.includeDirectories = false;
    SqliteIndexStore::QueryRuntimeStats queryStats{};
    hr = SqliteIndexStore::EnumerateVolume(databasePath.wstring(), request, nullptr, nullptr, callback, &skipState, &queryStats);
    state.Require(SUCCEEDED(hr), L"SQLite callback-skip enumeration failed.");
    state.Require(skipState.callbacks == 2u, std::format(L"Expected two callback invocations, got {}.", skipState.callbacks));
    state.Require(queryStats.emittedRows == 1u,
                  std::format(L"One skipped callback plus one emitted row should report emittedRows == 1, got {}.", queryStats.emittedRows));
    return state.failure.empty();
});
