// I35 S6 (Specs/Plans/Done/FileOperations_CopyMoveUxAndReliabilityReview_2026-10-11.md, section 10): the
// per-file cost of the progress reports. A flat folder of small files is copied from the test root on the
// Local provider's own route (Local to Local, Verify off) and through the host bridge (Local to the
// Dummy provider, its simulated latency set to zero for the case), and the wall time of every copy is
// reported, so builds can be compared on the same machine (an A/B of the R10.1 phase reports against the
// build before them). Each result requires the expected route plus exact observed names and bytes;
// verification and missing/corrupt-output controls run after the measured copy interval.
//
// REDSALAMANDER_FILEOPS_SMALLFILE_COUNT: files in the folder (default 2,000).
// REDSALAMANDER_FILEOPS_SMALLFILE_REPEATS: copies per route, alternating the routes (default 1).
// REDSALAMANDER_FILEOPS_SMALLFILE_TAG: a name for the build, carried in the reported rows.

namespace
{
constexpr std::wstring_view kSmallFileCountEnv   = L"REDSALAMANDER_FILEOPS_SMALLFILE_COUNT";
constexpr std::wstring_view kSmallFileRepeatsEnv = L"REDSALAMANDER_FILEOPS_SMALLFILE_REPEATS";
constexpr std::wstring_view kSmallFileTagEnv     = L"REDSALAMANDER_FILEOPS_SMALLFILE_TAG";
constexpr std::wstring_view kSmallFileDummyRoot  = L"/small-file-cost";
constexpr DWORD kSmallFileBytes                  = 1024u;
constexpr auto kSmallFileCopyTimeout             = std::chrono::minutes(20);

[[nodiscard]] uint32_t ReadSmallFileSetting(std::wstring_view name, uint32_t fallback, uint32_t minimum, uint32_t maximum)
{
    const std::wstring text = GetEnvVarTrimmed(name);
    if (text.empty())
        return fallback;
    wchar_t* end              = nullptr;
    const unsigned long value = std::wcstoul(text.c_str(), &end, 10);
    if (end == nullptr || *end != L'\0')
        return fallback;
    return std::clamp(static_cast<uint32_t>(std::min<unsigned long>(value, std::numeric_limits<uint32_t>::max())), minimum, maximum);
}

// count files of kSmallFileBytes each, every file with its own content.
[[nodiscard]] bool CreateSmallFileCorpus(const std::filesystem::path& folder, uint32_t count) noexcept
{
    if (! RecreateEmptyDirectory(folder))
        return false;
    std::array<char, kSmallFileBytes> content{};
    for (uint32_t index = 0u; index < count; ++index)
    {
        for (size_t offset = 0u; offset < content.size(); ++offset)
            content[offset] = static_cast<char>((index * 131u + static_cast<uint32_t>(offset) * 7u) & 0xFFu);
        const std::filesystem::path path = folder / std::format(L"f{:06}.bin", index);
        wil::unique_hfile file(CreateFileW(path.c_str(), GENERIC_WRITE, 0u, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
        DWORD written = 0u;
        if (! file || WriteFile(file.get(), content.data(), kSmallFileBytes, &written, nullptr) == FALSE || written != kSmallFileBytes)
            return false;
    }
    return true;
}

// Benchmark-only policy: verify the complete observed listing and each file's exact source
// bytes after stopping the wall timer. Provider success/counters cannot stand in for output.
[[nodiscard]] bool VerifySmallFileCopy(IFileSystem* sourceFs,
                                       IFileSystem* destinationFs,
                                       const std::filesystem::path& source,
                                       const std::filesystem::path& destination,
                                       uint32_t expectedCount,
                                       uint32_t& observedCount) noexcept
{
    observedCount     = 0u;
    const auto reject = [&](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Small-file output verification: destination={} accepted={} reason={}", destination.native(), observedCount, reason));
        return false;
    };
    wil::com_ptr<IFileSystemIO> sourceIo;
    wil::com_ptr<IFileSystemIO> destinationIo;
    wil::com_ptr<IFilesInformation> listing;
    if (! sourceFs || ! destinationFs || FAILED(sourceFs->QueryInterface(IID_PPV_ARGS(sourceIo.put()))) ||
        FAILED(destinationFs->QueryInterface(IID_PPV_ARGS(destinationIo.put()))) ||
        FAILED(destinationFs->ReadDirectoryInfo(ToPluginPathText(destination).c_str(), listing.put())) || ! listing)
        return reject(L"provider IO or directory listing unavailable");
    FileInfo* buffer    = nullptr;
    unsigned long bytes = 0u;
    unsigned long count = 0u;
    if (FAILED(listing->GetBuffer(&buffer)) || FAILED(listing->GetBufferSize(&bytes)) || FAILED(listing->GetCount(&count)) || count != expectedCount)
        return reject(std::format(L"directory buffer/count invalid: observed={} expected={}", count, expectedCount));
    Common::Plugins::PackedFileInfoCursor cursor(buffer, bytes, count);
    std::vector<bool> seen(expectedCount, false);
    for (unsigned long entryIndex = 0u; entryIndex < count; ++entryIndex)
    {
        const FileInfo* entry = nullptr;
        if (cursor.Next(&entry) != S_OK || ! entry || (entry->FileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0u)
            return reject(L"invalid packed record or non-file entry");
        const std::wstring_view name(entry->FileName, entry->FileNameSize / sizeof(wchar_t));
        if (name.size() != 11u || name.front() != L'f' || name.substr(7u) != L".bin")
            return reject(std::format(L"unexpected filename: {}", name));
        uint32_t index = 0u;
        for (size_t digit = 1u; digit < 7u; ++digit)
        {
            if (name[digit] < L'0' || name[digit] > L'9')
                return false;
            index = index * 10u + static_cast<uint32_t>(name[digit] - L'0');
        }
        if (index >= expectedCount || seen[index])
            return false;
        seen[index] = true;
        const std::filesystem::path leaf(name);
        std::string actual;
        std::string expected;
        if (! ReadFileTextFsIo(sourceIo, source / leaf, expected) || expected.size() != kSmallFileBytes ||
            ! ReadFileTextFsIo(destinationIo, destination / leaf, actual) || actual != expected)
            return reject(std::format(L"content mismatch: name={} sourceBytes={} destinationBytes={}", name, expected.size(), actual.size()));
        ++observedCount;
    }
    return observedCount == expectedCount;
}

// Removes the case's folder from the Dummy provider (its files live in memory).
void RemoveSmallFileDummyRoot(IFileSystem* dummy) noexcept
{
    if (! dummy)
        return;
    FileSystemOptions options{};
    options.sizeBytes  = sizeof(options);
    options.linkPolicy = FILESYSTEM_LINK_PRESERVE;
    static_cast<void>(dummy->DeleteItem(std::wstring(kSmallFileDummyRoot).c_str(), FILESYSTEM_FLAG_RECURSIVE, &options, nullptr, nullptr));
}
} // namespace

[[nodiscard]] bool RunFileOpsSmallFileCopyCostChecks(SelfTestState& state, std::wstring& failure) noexcept
{
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! state.fsDummy || ! state.infoDummy || ! CancelOutstandingSelfTestOperations(operations))
    {
        failure = L"owner, Local and Dummy providers or initial worker quiet point unavailable";
        return false;
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
    {
        failure = L"no-activation guard unavailable";
        return false;
    }

    const uint32_t count             = ReadSmallFileSetting(kSmallFileCountEnv, 2'000u, 16u, 200'000u);
    const uint32_t repeats           = ReadSmallFileSetting(kSmallFileRepeatsEnv, 1u, 1u, 20u);
    const std::wstring tagText       = GetEnvVarTrimmed(kSmallFileTagEnv);
    const std::wstring tag           = tagText.empty() ? std::wstring(L"none") : tagText;
    const std::filesystem::path root = state.tempRoot / L"small-file-cost";
    auto& settings                   = EnsureFileOperationsSettingsForSelfTest();
    const bool previousVerify        = settings.verifyAfterCopy;
    const bool previousAutoDismiss   = operations->GetAutoDismissSuccess();
    const HWND originalPopup         = operations->GetPopupHwndForSelfTest();
    // The Dummy provider simulates a per-access latency for the suite; the bridge copy measures the host
    // and the bridge, so the latency is zero for the case and the suite's configuration comes back after.
    std::string dummyConfiguration;
    const bool dummyBackedUp = BackupPluginConfiguration(state.infoDummy.get(), dummyConfiguration);
    std::vector<uint64_t> tasks;
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        const bool quiet = CancelOutstandingSelfTestOperations(operations);
        for (const uint64_t taskId : tasks)
        {
            if (quiet)
            {
                operations->DismissCompletedTask(taskId);
                operations->DebugRemoveDiagnosticsForTask(taskId);
                state.completedTasks.erase(taskId);
            }
        }
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = previousVerify;
        operations->SetAutoDismissSuccess(previousAutoDismiss);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
        if (quiet && ! originalPopup && operations->GetPopupHwndForSelfTest())
            operations->DebugResetPopupForSelfTest();
        RemoveSmallFileDummyRoot(state.fsDummy.get());
        if (dummyBackedUp)
            static_cast<void>(SetPluginConfiguration(state.infoDummy.get(), dummyConfiguration));
        static_cast<void>(SelfTest::RemoveAll(root));
    });
    if (! dummyBackedUp)
    {
        failure = L"the Dummy provider's configuration could not be read";
        return false;
    }
    std::string noLatency = dummyConfiguration;
    for (const std::string_view key : {std::string_view("\"latencyMs\":"), std::string_view("\"streamChunkLatencyMs\":")})
    {
        const size_t at = noLatency.find(key);
        if (at == std::string::npos)
            continue;
        const size_t valueBegin = at + key.size();
        size_t valueEnd         = valueBegin;
        while (valueEnd < noLatency.size() && noLatency[valueEnd] >= '0' && noLatency[valueEnd] <= '9')
            ++valueEnd;
        noLatency.replace(valueBegin, valueEnd - valueBegin, "0");
    }
    if (! SetPluginConfiguration(state.infoDummy.get(), noLatency))
    {
        failure = L"the Dummy provider's latency could not be set to zero";
        return false;
    }
    operations->SetAutoDismissSuccess(false);
    settings.verifyAfterCopy = false;

    const std::filesystem::path corpus = root / L"source" / L"corpus";
    const auto createStarted           = std::chrono::steady_clock::now();
    if (! CreateSmallFileCorpus(corpus, count))
    {
        failure = std::format(L"the {}-file corpus could not be created under '{}'", count, corpus.native());
        return false;
    }
    AppendLog(std::format(L"Small-file cost: {} files of {} bytes created in {} ms (tag={}).",
                          count,
                          kSmallFileBytes,
                          std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - createStarted).count(),
                          tag));

    for (uint32_t repeat = 0u; repeat < repeats; ++repeat)
    {
        for (const bool bridge : {false, true})
        {
            const std::wstring_view route = bridge ? L"bridge" : L"native";
            std::filesystem::path destination;
            if (bridge)
            {
                RemoveSmallFileDummyRoot(state.fsDummy.get());
                destination = std::filesystem::path(std::format(L"{}/out-{:02}", kSmallFileDummyRoot, repeat));
                if (! EnsureDummyFolderExists(state.fsDummy.get(), kSmallFileDummyRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), destination.native()))
                {
                    failure = std::format(L"{} copy {}: the Dummy destination is unavailable", route, repeat);
                    return false;
                }
            }
            else
            {
                destination = root / std::format(L"out-{:02}-{}", repeat, route);
                if (! RecreateEmptyDirectory(destination))
                {
                    failure = std::format(L"{} copy {}: destination unavailable", route, repeat);
                    return false;
                }
            }
            const auto started  = std::chrono::steady_clock::now();
            const auto admitted = StartFileOperationAndGetId(operations,
                                                             FILESYSTEM_COPY,
                                                             FolderWindow::Pane::Left,
                                                             FolderWindow::Pane::Right,
                                                             state.fsLocal,
                                                             {corpus},
                                                             destination,
                                                             FILESYSTEM_FLAG_RECURSIVE,
                                                             false,
                                                             0u,
                                                             FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                             false,
                                                             bridge ? state.fsDummy : wil::com_ptr<IFileSystem>{});
            if (! admitted.has_value())
            {
                failure = std::format(L"{} copy {}: not admitted", route, repeat);
                return false;
            }
            const uint64_t taskId = admitted.value();
            tasks.push_back(taskId);
            const auto done = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                return state.completedTasks.contains(taskId) && operations->FindTask(taskId) == nullptr;
            }, {.timeout = SelfTest::Scale(std::chrono::duration_cast<std::chrono::milliseconds>(kSmallFileCopyTimeout)), .operationName = L"Small-file Copy"});
            const uint64_t wallUs =
                static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());
            if (! done.conditionMet)
            {
                failure = std::format(L"{} copy {}: {}", route, repeat, done.timeoutDiagnostic);
                return false;
            }
            const CompletedTaskInfo& info = state.completedTasks.at(taskId);
            uint32_t copied               = 0u;
            IFileSystem* destinationFs    = bridge ? state.fsDummy.get() : state.fsLocal.get();
            const bool outputVerified     = VerifySmallFileCopy(state.fsLocal.get(), destinationFs, corpus, destination / L"corpus", count, copied);
            const bool routeAsExpected    = bridge ? info.bridgeFileAdmissionCount != 0u : info.bridgeFileAdmissionCount == 0u;
            if (info.hr != S_OK || ! outputVerified || copied != count || ! routeAsExpected)
            {
                failure = std::format(L"{} copy {}: hr=0x{:08X}, {} of {} files, bridge admissions {}",
                                      route,
                                      repeat,
                                      static_cast<unsigned long>(info.hr),
                                      copied,
                                      count,
                                      info.bridgeFileAdmissionCount);
                return false;
            }
            if (repeat == 0u)
            {
                const auto firstOutput = destination / L"corpus" / L"f000000.bin";
                FileSystemOptions options{};
                options.sizeBytes     = sizeof(options);
                options.linkPolicy    = FILESYSTEM_LINK_PRESERVE;
                uint32_t controlCount = 0u;
                if (FAILED(destinationFs->DeleteItem(ToPluginPathText(firstOutput).c_str(), FILESYSTEM_FLAG_NONE, &options, nullptr, nullptr)) ||
                    VerifySmallFileCopy(state.fsLocal.get(), destinationFs, corpus, destination / L"corpus", count, controlCount) ||
                    ! DummyWriteTextFile(destinationFs, ToPluginPathText(firstOutput), std::string(kSmallFileBytes, 'X'), true) ||
                    VerifySmallFileCopy(state.fsLocal.get(), destinationFs, corpus, destination / L"corpus", count, controlCount))
                {
                    failure = std::format(L"{} output verifier accepted missing or corrupt output", route);
                    return false;
                }
            }
            const std::wstring detail =
                std::format(L"route={} files={} repeat={} tag={} concurrency={}", route, count, repeat, tag, info.configuredMaxConcurrency);
            AppendLog(std::format(L"Small-file cost: {} wallUs={} ({:.1f} us/file)", detail, wallUs, static_cast<double>(wallUs) / static_cast<double>(count)));
            Debug::Perf::Emit(L"FileOps.SelfTest.SmallFileCopy", detail, wallUs, count, info.bridgeFileAdmissionCount, S_OK);
            if (bridge)
                RemoveSmallFileDummyRoot(state.fsDummy.get());
            else
                static_cast<void>(SelfTest::RemoveAll(destination));
        }
    }
    return true;
}
