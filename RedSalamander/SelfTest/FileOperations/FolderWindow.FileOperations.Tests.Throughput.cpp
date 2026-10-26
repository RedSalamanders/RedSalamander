// I35 S6 (Specs/Plans/Done/FileOperations_CopyMoveUxAndReliabilityReview_2026-10-11.md, section 10):
// SMB throughput and the speed display. The throughput cases measure engine facts where every
// route's progress reports arrive (gaps between reports, silences and the phase each one was spent
// in, the variation of whole-second throughput) and report them for the archived A/B comparison;
// they never assert the display, whose rules have their own deterministic gates.

namespace
{
using ThroughputCadence = FolderWindow::FileOperationState::Task::ProgressCadenceStats;
using ThroughputPhase   = FolderWindow::FileOperationState::ProgressPhase;

// smoke (default): 64 MiB + 16 x 4 MiB, the routine Full-run size. full: the R10.5 corpus,
// 1 GiB + 64 x 16 MiB, requested for archived perf evidence. large: one 4 GiB file + 16 x 16 MiB,
// a file several times the client cache's dirty-page threshold (about 1.1 GiB on the reference
// machine), so a cached writer is held inside a write in the middle of the file, not only at commit.
constexpr std::wstring_view kThroughputCorpusEnv = L"REDSALAMANDER_FILEOPS_THROUGHPUT_CORPUS";
// "bridge" copies through the host bridge (the writer of a Managed Move or a Verify On copy) instead of
// the Local provider's own route.
constexpr std::wstring_view kThroughputRouteEnv = L"REDSALAMANDER_FILEOPS_THROUGHPUT_ROUTE";
// Comma-separated transfer variants run back to back on one corpus (default "adopted"): an A/B
// schedule such as "base,e,base,e" interleaves the baseline and one candidate.
constexpr std::wstring_view kThroughputVariantsEnv = L"REDSALAMANDER_FILEOPS_THROUGHPUT_VARIANTS";
// "loopback" copies to the loopback administrative-share alias even when an SMB share is registered.
constexpr std::wstring_view kThroughputSmbTargetEnv = L"REDSALAMANDER_FILEOPS_THROUGHPUT_SMB_TARGET";
constexpr std::wstring_view kThroughputSmbRootEnv   = L"REDSALAMANDER_SELFTEST_SMB_ROOT";
constexpr auto kThroughputCopyTimeout               = std::chrono::minutes(15);

struct ThroughputCorpus final
{
    std::wstring_view name;
    uint64_t largeBytes = 0;
    uint32_t smallCount = 0;
    uint64_t smallBytes = 0;

    [[nodiscard]] uint64_t TotalBytes() const noexcept
    {
        return largeBytes + static_cast<uint64_t>(smallCount) * smallBytes;
    }
};

constexpr uint64_t kThroughputMiB = 1024ull * 1024ull;
constexpr ThroughputCorpus kSmokeThroughputCorpus{L"smoke", 64ull * kThroughputMiB, 16u, 4ull * kThroughputMiB};
constexpr ThroughputCorpus kFullThroughputCorpus{L"full", 1024ull * kThroughputMiB, 64u, 16ull * kThroughputMiB};
constexpr ThroughputCorpus kLargeThroughputCorpus{L"large", 4096ull * kThroughputMiB, 16u, 16ull * kThroughputMiB};

[[nodiscard]] const ThroughputCorpus& SelectThroughputCorpus() noexcept
{
    const std::wstring requested = GetEnvVarTrimmed(kThroughputCorpusEnv);
    if (_wcsicmp(requested.c_str(), L"full") == 0)
        return kFullThroughputCorpus;
    return _wcsicmp(requested.c_str(), L"large") == 0 ? kLargeThroughputCorpus : kSmokeThroughputCorpus;
}

[[nodiscard]] bool ThroughputUsesBridgeRoute() noexcept
{
    return _wcsicmp(GetEnvVarTrimmed(kThroughputRouteEnv).c_str(), L"bridge") == 0;
}

// I35 R10.3: one transfer configuration. `tuning` replaces the adopted FileSystemTransferTuning
// defaults for the run; "adopted" keeps them.
struct ThroughputVariant final
{
    std::wstring_view name;
    std::optional<uint32_t> tuning;
};

constexpr uint32_t kThroughputNoTemporary = FILESYSTEM_TRANSFER_TUNING_NO_TEMPORARY_REMOTE | FILESYSTEM_TRANSFER_TUNING_NO_TEMPORARY_LOCAL;
const std::array<ThroughputVariant, 10> kThroughputVariants{{
    {L"adopted", std::nullopt},
    {L"base", FILESYSTEM_TRANSFER_TUNING_NONE},
    {L"e", FILESYSTEM_TRANSFER_TUNING_NO_TEMPORARY_REMOTE},
    {L"el", FILESYSTEM_TRANSFER_TUNING_NO_TEMPORARY_LOCAL},
    {L"ea", kThroughputNoTemporary},
    {L"b", FILESYSTEM_TRANSFER_TUNING_PREALLOCATE},
    {L"d", FILESYSTEM_TRANSFER_TUNING_SKIP_ABSENT_BACKUP_STREAMS},
    {L"c", FILESYSTEM_TRANSFER_TUNING_DESTINATION_CONCURRENCY},
    {L"ebd", kThroughputNoTemporary | FILESYSTEM_TRANSFER_TUNING_PREALLOCATE | FILESYSTEM_TRANSFER_TUNING_SKIP_ABSENT_BACKUP_STREAMS},
    {L"w", FILESYSTEM_TRANSFER_TUNING_WRITE_THROUGH_REMOTE},
}};

// The schedule the environment asks for; an unknown name refuses the whole case.
[[nodiscard]] std::optional<std::vector<ThroughputVariant>> ParseThroughputVariants(std::wstring& unknownName)
{
    std::wstring text = GetEnvVarTrimmed(kThroughputVariantsEnv);
    if (text.empty())
        text = L"adopted";
    std::vector<ThroughputVariant> schedule;
    size_t start = 0u;
    while (start <= text.size())
    {
        const size_t comma      = text.find(L',', start);
        std::wstring_view token = std::wstring_view(text).substr(start, comma == std::wstring::npos ? std::wstring::npos : comma - start);
        while (! token.empty() && token.front() == L' ')
            token.remove_prefix(1u);
        while (! token.empty() && token.back() == L' ')
            token.remove_suffix(1u);
        if (! token.empty())
        {
            const auto known = std::ranges::find_if(kThroughputVariants, [&](const ThroughputVariant& variant) noexcept {
                return variant.name.size() == token.size() && _wcsnicmp(variant.name.data(), token.data(), token.size()) == 0;
            });
            if (known == kThroughputVariants.end())
            {
                unknownName.assign(token);
                return std::nullopt;
            }
            schedule.push_back(*known);
        }
        if (comma == std::wstring::npos)
            break;
        start = comma + 1u;
    }
    if (schedule.empty())
        schedule.push_back(kThroughputVariants.front());
    return schedule;
}

// Incompressible content, different in every file: a destination that compresses or deduplicates
// must not make one variant look faster than another. xorshift64*: fast, and incompressible enough
// for any storage-side compressor.
class ThroughputContentStream final
{
public:
    explicit ThroughputContentStream(uint64_t seed) noexcept : _state(seed * 0x9E3779B97F4A7C15ull + 0xD1B54A32D192ED03ull)
    {
    }

    void Fill(std::span<uint64_t> words) noexcept
    {
        for (uint64_t& word : words)
        {
            _state ^= _state >> 12u;
            _state ^= _state << 25u;
            _state ^= _state >> 27u;
            word = _state * 0x2545F4914F6CDD1Dull;
        }
    }

private:
    uint64_t _state = 0u;
};

constexpr size_t kThroughputChunkWords = (4u * 1024u * 1024u) / sizeof(uint64_t);

[[nodiscard]] bool WriteThroughputCorpusFile(const std::filesystem::path& path, uint64_t bytes, uint64_t seed) noexcept
{
    wil::unique_handle file(CreateFileW(MakeExtendedPathForSelfTest(path).c_str(),
                                        GENERIC_WRITE,
                                        FILE_SHARE_READ | FILE_SHARE_DELETE,
                                        nullptr,
                                        CREATE_ALWAYS,
                                        FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                                        nullptr));
    if (! file)
        return false;
    std::vector<uint64_t> chunk(kThroughputChunkWords);
    ThroughputContentStream content(seed);
    uint64_t remaining = bytes;
    while (remaining > 0u)
    {
        content.Fill(chunk);
        const DWORD toWrite = static_cast<DWORD>(std::min<uint64_t>(remaining, chunk.size() * sizeof(uint64_t)));
        DWORD written       = 0u;
        if (WriteFile(file.get(), chunk.data(), toWrite, &written, nullptr) == FALSE || written != toWrite)
            return false;
        remaining -= toWrite;
    }
    return true;
}

// True when the file holds exactly the content WriteThroughputCorpusFile writes for `seed`.
[[nodiscard]] bool ThroughputFileMatches(const std::filesystem::path& path, uint64_t bytes, uint64_t seed) noexcept
{
    wil::unique_handle file(CreateFileW(MakeExtendedPathForSelfTest(path).c_str(),
                                        GENERIC_READ,
                                        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                        nullptr,
                                        OPEN_EXISTING,
                                        FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                                        nullptr));
    LARGE_INTEGER size{};
    if (! file || GetFileSizeEx(file.get(), &size) == FALSE || static_cast<uint64_t>(size.QuadPart) != bytes)
        return false;
    std::vector<uint64_t> expected(kThroughputChunkWords);
    std::vector<uint64_t> actual(kThroughputChunkWords);
    ThroughputContentStream content(seed);
    uint64_t remaining = bytes;
    while (remaining > 0u)
    {
        content.Fill(expected);
        const DWORD toRead = static_cast<DWORD>(std::min<uint64_t>(remaining, expected.size() * sizeof(uint64_t)));
        DWORD read         = 0u;
        if (ReadFile(file.get(), actual.data(), toRead, &read, nullptr) == FALSE || read != toRead || std::memcmp(actual.data(), expected.data(), toRead) != 0)
            return false;
        remaining -= toRead;
    }
    return true;
}

// The corpus files: <folder>\large.bin and <folder>\small-NN.bin, copied as top-level items. The set
// is flat because the registered NAS exposes no file identity: a nested directory's stage could not
// be published there (its publication outcome is unknown) and the run would measure nothing.
[[nodiscard]] std::vector<std::filesystem::path> ThroughputCorpusFiles(const std::filesystem::path& folder, const ThroughputCorpus& corpus)
{
    std::vector<std::filesystem::path> files;
    files.reserve(1u + corpus.smallCount);
    files.push_back(folder / L"large.bin");
    for (uint32_t index = 0u; index < corpus.smallCount; ++index)
        files.push_back(folder / std::format(L"small-{:02}.bin", index));
    return files;
}

[[nodiscard]] bool CreateThroughputCorpus(const std::filesystem::path& corpusFolder, const ThroughputCorpus& corpus)
{
    if (! RecreateEmptyDirectory(corpusFolder))
        return false;
    const std::vector<std::filesystem::path> files = ThroughputCorpusFiles(corpusFolder, corpus);
    for (size_t index = 0u; index < files.size(); ++index)
    {
        if (! WriteThroughputCorpusFile(files[index], index == 0u ? corpus.largeBytes : corpus.smallBytes, index == 0u ? 1u : 99u + index))
            return false;
    }
    return true;
}

// Every corpus file arrived with its exact size (the content itself is the provider's own proof).
[[nodiscard]] bool ThroughputCopyComplete(const std::filesystem::path& destination, const ThroughputCorpus& corpus)
{
    const std::vector<std::filesystem::path> files = ThroughputCorpusFiles(destination, corpus);
    for (size_t index = 0u; index < files.size(); ++index)
    {
        if (! FileSizeEquals(files[index], index == 0u ? corpus.largeBytes : corpus.smallBytes))
            return false;
    }
    return true;
}

// The registered SMB self-test root (path.smb.primary), revalidated here: a UNC directory at
// least three segments deep, one segment naming the self-test area, no dot segments. Empty when
// not configured or not acceptable; `refusal` explains a configured but refused root.
[[nodiscard]] std::filesystem::path ResolveThroughputSmbRoot(std::wstring& refusal)
{
    std::wstring text = GetEnvVarTrimmed(kThroughputSmbRootEnv);
    std::replace(text.begin(), text.end(), L'/', L'\\');
    while (text.size() > 2u && text.back() == L'\\')
        text.pop_back();
    if (text.empty())
        return {};
    size_t segments    = 0u;
    bool selfTestNamed = false;
    bool dotSegment    = false;
    for (size_t i = 2u; i < text.size();)
    {
        while (i < text.size() && text[i] == L'\\')
            ++i;
        const size_t start = i;
        while (i < text.size() && text[i] != L'\\')
            ++i;
        if (start == i)
            continue;
        const std::wstring_view segment = std::wstring_view(text).substr(start, i - start);
        ++segments;
        dotSegment = dotSegment || segment == L"." || segment == L"..";
        std::wstring lowered(segment);
        std::ranges::transform(lowered, lowered.begin(), [](wchar_t ch) noexcept { return static_cast<wchar_t>(towlower(ch)); });
        selfTestNamed = selfTestNamed || lowered.find(L"selftest") != std::wstring::npos;
    }
    if (! text.starts_with(L"\\\\") || segments < 3u || dotSegment || ! selfTestNamed)
    {
        refusal = L"the SMB root must be a UNC directory below the share, contain no dot segments and include a 'selftest' segment";
        return {};
    }
    return std::filesystem::path(text);
}

struct ThroughputSample final
{
    std::wstring_view target;
    std::wstring_view route;
    std::wstring_view variant;
    uint64_t bytes                 = 0;
    uint64_t wallUs                = 0;
    unsigned int concurrencyBudget = 0;
    ThroughputCadence cadence{};
};

[[nodiscard]] double ThroughputMegabytesPerSecond(uint64_t bytes, uint64_t us) noexcept
{
    return us == 0u ? 0.0 : (static_cast<double>(bytes) / (1024.0 * 1024.0)) / (static_cast<double>(us) / 1'000'000.0);
}

[[nodiscard]] std::wstring DescribeThroughputSample(const ThroughputSample& sample)
{
    const ThroughputCadence& cadence = sample.cadence;
    const auto silences              = [&cadence](ThroughputPhase phase) noexcept { return cadence.silenceCountByPhase[static_cast<size_t>(phase)]; };
    const double gib                 = static_cast<double>(sample.bytes) / (1024.0 * 1024.0 * 1024.0);
    return std::format(L"target={} route={} variant={} bytes={} wallMs={} MBps={:.1f} budget={} gaps={} p50Ms={:.1f} p95Ms={:.1f} p99Ms={:.1f} "
                       L"maxMs={:.1f} stalls={} stallMs={} silences={} silencesPerGiB={:.2f} silenceMs={} maxSilenceMs={:.1f} opening={} "
                       L"transferring={} finalizing={} publishing={} verifying={} unknown={} betweenItems={} seconds={} cv={:.3f}",
                       sample.target,
                       sample.route,
                       sample.variant,
                       sample.bytes,
                       sample.wallUs / 1'000u,
                       ThroughputMegabytesPerSecond(sample.bytes, sample.wallUs),
                       sample.concurrencyBudget,
                       cadence.gapCount,
                       static_cast<double>(cadence.GapPercentileUs(0.50)) / 1000.0,
                       static_cast<double>(cadence.GapPercentileUs(0.95)) / 1000.0,
                       static_cast<double>(cadence.GapPercentileUs(0.99)) / 1000.0,
                       static_cast<double>(cadence.maxGapUs) / 1000.0,
                       cadence.StallCount(),
                       cadence.StallUs() / 1'000u,
                       cadence.SilenceCount(),
                       gib > 0.0 ? static_cast<double>(cadence.SilenceCount()) / gib : 0.0,
                       cadence.SilenceUs() / 1'000u,
                       static_cast<double>(cadence.maxSilenceUs) / 1000.0,
                       silences(ThroughputPhase::Opening),
                       silences(ThroughputPhase::Transferring),
                       silences(ThroughputPhase::Finalizing),
                       silences(ThroughputPhase::Publishing),
                       silences(ThroughputPhase::Verifying),
                       silences(ThroughputPhase::Unknown),
                       cadence.betweenItemGapCount,
                       cadence.throughputBucketCount,
                       cadence.ThroughputCoefficientOfVariation());
}

// Owns the settings and popup state a throughput case changes, and every task it starts.
class ThroughputCaseScope final
{
public:
    ThroughputCaseScope(SelfTestState& state, FolderWindow::FileOperationState* operations) noexcept : _state(state), _operations(operations)
    {
        auto& settings = EnsureFileOperationsSettingsForSelfTest();
        _verify        = settings.verifyAfterCopy;
        _autoDismiss   = operations->GetAutoDismissSuccess();
        _originalPopup = operations->GetPopupHwndForSelfTest();
        operations->SetAutoDismissSuccess(false);
    }
    ThroughputCaseScope(const ThroughputCaseScope&)            = delete;
    ThroughputCaseScope& operator=(const ThroughputCaseScope&) = delete;
    ~ThroughputCaseScope()
    {
        const bool quiet = CancelOutstandingSelfTestOperations(_operations);
        for (const uint64_t taskId : tasks)
        {
            if (quiet)
            {
                _operations->DismissCompletedTask(taskId);
                _operations->DebugRemoveDiagnosticsForTask(taskId);
                _state.completedTasks.erase(taskId);
            }
        }
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = _verify;
        _operations->SetAutoDismissSuccess(_autoDismiss);
        static_cast<void>(_operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
        if (quiet && ! _originalPopup && _operations->GetPopupHwndForSelfTest())
            _operations->DebugResetPopupForSelfTest();
    }

    std::vector<uint64_t> tasks;

private:
    SelfTestState& _state;
    FolderWindow::FileOperationState* _operations = nullptr;
    HWND _originalPopup                           = nullptr;
    bool _verify                                  = false;
    bool _autoDismiss                             = false;
};

// One Copy of the corpus folder into `destinationParent` (Verify off), timed from admission to the
// retired task, on the Local provider's own route or, with `bridgeRoute`, through the host bridge (the
// same provider instance named as the destination). The destination copy is removed afterwards.
[[nodiscard]] std::optional<ThroughputSample> RunThroughputCopy(SelfTestState& state,
                                                                FolderWindow::FileOperationState* operations,
                                                                ThroughputCaseScope& scope,
                                                                const std::filesystem::path& corpusFolder,
                                                                const ThroughputCorpus& corpus,
                                                                const std::filesystem::path& destinationParent,
                                                                std::wstring_view target,
                                                                const ThroughputVariant& variant,
                                                                bool bridgeRoute,
                                                                std::wstring& failure)
{
    if (! RecreateEmptyDirectory(destinationParent))
    {
        failure = std::format(L"{}: destination '{}' unavailable", target, destinationParent.native());
        return std::nullopt;
    }
    const auto removeCopy = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(destinationParent)); });

    const auto startedAt = std::chrono::steady_clock::now();
    const auto admitted  = StartFileOperationAndGetId(operations,
                                                      FILESYSTEM_COPY,
                                                      FolderWindow::Pane::Left,
                                                      FolderWindow::Pane::Right,
                                                      state.fsLocal,
                                                      ThroughputCorpusFiles(corpusFolder, corpus),
                                                      destinationParent,
                                                      FILESYSTEM_FLAG_RECURSIVE,
                                                      false,
                                                      0u,
                                                      FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                      false,
                                                      bridgeRoute ? state.fsLocal : nullptr);
    if (! admitted.has_value())
    {
        failure = std::format(L"{}: the Copy was not admitted", target);
        return std::nullopt;
    }
    const uint64_t taskId = admitted.value();
    scope.tasks.push_back(taskId);
    const auto done       = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return state.completedTasks.contains(taskId) && operations->FindTask(taskId) == nullptr;
    }, {.timeout = SelfTest::Scale(std::chrono::duration_cast<std::chrono::milliseconds>(kThroughputCopyTimeout)), .operationName = L"Throughput Copy"});
    const uint64_t wallUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - startedAt).count());
    if (! done.conditionMet)
    {
        failure = std::format(L"{}: {}", target, done.timeoutDiagnostic);
        return std::nullopt;
    }
    const CompletedTaskInfo& info = state.completedTasks.at(taskId);
    if (info.hr != S_OK || ! ThroughputCopyComplete(destinationParent, corpus))
    {
        failure = std::format(L"{}: the Copy did not publish the whole corpus (hr=0x{:08X})", target, static_cast<unsigned long>(info.hr));
        return std::nullopt;
    }
    if (bridgeRoute && info.bridgeStageCreateCount < 1u + static_cast<uint64_t>(corpus.smallCount))
    {
        failure = std::format(
            L"{}: the Copy did not run through the host bridge ({} bridge stages for {} files)", target, info.bridgeStageCreateCount, 1u + corpus.smallCount);
        return std::nullopt;
    }

    ThroughputSample sample{};
    sample.target            = target;
    sample.route             = bridgeRoute ? std::wstring_view(L"bridge") : std::wstring_view(L"provider");
    sample.variant           = variant.name;
    sample.bytes             = corpus.TotalBytes();
    sample.wallUs            = wallUs;
    sample.concurrencyBudget = info.configuredMaxConcurrency;
    sample.cadence           = info.progressCadence;
    return sample;
}

void EmitThroughputSample(const ThroughputSample& sample, std::wstring_view corpusName)
{
    const std::wstring detail = std::format(L"corpus={} {}", corpusName, DescribeThroughputSample(sample));
    AppendLog(std::format(L"Throughput: {}", detail));
    Debug::Perf::Emit(L"FileOps.SelfTest.Throughput.Copy",
                      detail,
                      sample.wallUs,
                      sample.bytes,
                      static_cast<uint64_t>(std::llround(sample.cadence.ThroughputCoefficientOfVariation() * 1000.0)),
                      S_OK);
}

struct ThroughputCaseResult final
{
    bool passed  = false;
    bool skipped = false;
    std::wstring reason;
};

// Copies one corpus to one destination root for every scheduled variant, back to back.
[[nodiscard]] ThroughputCaseResult RunThroughputSchedule(SelfTestState& state,
                                                         std::wstring_view caseName,
                                                         std::wstring_view target,
                                                         const std::filesystem::path& destinationRoot,
                                                         std::vector<ThroughputSample>& samples)
{
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return {.reason = L"owner, Local provider or initial worker quiet point unavailable"};
    std::wstring unknownVariant;
    const auto schedule = ParseThroughputVariants(unknownVariant);
    if (! schedule.has_value())
        return {.reason = std::format(L"unknown transfer variant '{}' in {}", unknownVariant, kThroughputVariantsEnv)};
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return {.reason = L"no-activation guard unavailable"};

    const ThroughputCorpus& corpus           = SelectThroughputCorpus();
    const std::filesystem::path corpusParent = state.tempRoot / std::format(L"{}-source", caseName);
    const std::filesystem::path corpusFolder = corpusParent / L"corpus";
    const auto removeCorpus                  = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(corpusParent)); });
    if (! CreateThroughputCorpus(corpusFolder, corpus))
        return {.reason = std::format(L"the {} corpus could not be created under '{}'", corpus.name, corpusParent.native())};

    ThroughputCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    const auto restoreTuning                                  = wil::scope_exit([]() noexcept { SetFileOpsTransferTuningOverrideForSelfTest(std::nullopt); });
    const bool bridgeRoute                                    = ThroughputUsesBridgeRoute();
    for (size_t run = 0u; run < schedule->size(); ++run)
    {
        const ThroughputVariant& variant = schedule->at(run);
        SetFileOpsTransferTuningOverrideForSelfTest(variant.tuning);
        std::wstring failure;
        const auto sample = RunThroughputCopy(state,
                                              operations,
                                              scope,
                                              corpusFolder,
                                              corpus,
                                              destinationRoot / std::format(L"run-{:02}-{}", run, variant.name),
                                              target,
                                              variant,
                                              bridgeRoute,
                                              failure);
        if (! sample.has_value())
            return {.reason = std::format(L"run {} ({}): {}", run, variant.name, failure)};
        if (sample->cadence.gapCount == 0u || sample->cadence.lastBytes < corpus.TotalBytes())
            return {.reason = std::format(
                        L"run {} ({}): the host measured no report cadence for the Copy ({})", run, variant.name, DescribeThroughputSample(*sample))};
        // Engine fact (R10.5): every silence longer than 600 ms followed reports that named their phase.
        if (sample->cadence.silenceCountByPhase[static_cast<size_t>(ThroughputPhase::Unknown)] != 0u ||
            sample->cadence.stallCountByPhase[static_cast<size_t>(ThroughputPhase::Unknown)] != 0u)
            return {.reason = std::format(
                        L"run {} ({}): a silence over 600 ms followed a report without a phase ({})", run, variant.name, DescribeThroughputSample(*sample))};
        EmitThroughputSample(*sample, corpus.name);
        samples.push_back(*sample);
    }
    return {.passed = true};
}

// ---------------------------------------------------------------------------------------------
// R10.5 SMB cases: the target is the registered SMB self-test root (a unique per-run child) when it
// is configured and reachable, else the loopback administrative-share alias of the case root.

struct ThroughputShareTarget final
{
    std::filesystem::path root;
    std::wstring_view name;
    bool nas = false;
    std::optional<ThroughputCaseResult> unavailable;
};

[[nodiscard]] ThroughputShareTarget ResolveThroughputShareTarget(SelfTestState& state, std::wstring_view childPrefix, const std::filesystem::path& loopbackRoot)
{
    const auto reject = [](std::wstring_view reason) { return ThroughputShareTarget{.unavailable = ThroughputCaseResult{.reason = std::wstring(reason)}}; };
    const auto skip   = [](std::wstring reason)
    { return ThroughputShareTarget{.unavailable = ThroughputCaseResult{.skipped = true, .reason = std::move(reason)}}; };
    std::wstring refusal;
    const std::filesystem::path smbRoot = ResolveThroughputSmbRoot(refusal);
    if (! refusal.empty())
        return reject(refusal);
    const bool loopbackRequested = _wcsicmp(GetEnvVarTrimmed(kThroughputSmbTargetEnv).c_str(), L"loopback") == 0;
    if (! smbRoot.empty() && ! loopbackRequested)
    {
        const DWORD attributes = GetFileAttributesW(smbRoot.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
        {
            const DWORD error = attributes == INVALID_FILE_ATTRIBUTES ? GetLastError() : ERROR_DIRECTORY;
            return skip(std::format(L"the registered SMB self-test root is unreachable or not a directory (error={})", error));
        }
        return {.root = smbRoot / std::format(L"{}-{}-{}", childPrefix, GetCurrentProcessId(), GetTickCount64()), .name = L"nas", .nas = true};
    }
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");
    if (uncStatus == LoopbackShareStatus::Unreachable)
        return skip(L"no SMB self-test share is registered and the loopback administrative share is not reachable");
    return {.root = LoopbackShareAlias(loopbackRoot), .name = L"loopback"};
}

// One Move file: its new content at the source, and the content a destination already had (if any).
struct MoveFixtureFile final
{
    std::wstring_view relative;
    uint64_t newSeed = 0u;
    uint64_t oldSeed = 0u;
};

constexpr uint64_t kMoveNewBytes = 16ull * 1024ull * 1024ull; // several reports per file
constexpr uint64_t kMoveOldBytes = 4ull * 1024ull;
// Two files at the top and one in each of two nested folders: a created folder's children publish while the walker
// holds no handle on it (I35: a Samba server refuses every rename inside a folder held with DELETE access).
constexpr std::array<MoveFixtureFile, 4> kMoveFixtureFiles{{
    {L"a.bin", 1u, 11u},
    {L"b.bin", 2u, 12u},
    {L"nested\\c.bin", 3u, 13u},
    {L"nested\\deeper\\d.bin", 4u, 14u},
}};

[[nodiscard]] bool WriteMoveFixture(const std::filesystem::path& payload, bool oldContent, uint64_t seedBase) noexcept
{
    if (! RecreateEmptyDirectory(payload))
        return false;
    for (const MoveFixtureFile& file : kMoveFixtureFiles)
    {
        std::error_code error;
        std::filesystem::create_directories((payload / file.relative).parent_path(), error);
        if (error)
            return false;
        if (! WriteThroughputCorpusFile(
                payload / file.relative, oldContent ? kMoveOldBytes : kMoveNewBytes, seedBase + (oldContent ? file.oldSeed : file.newSeed)))
            return false;
    }
    return true;
}

// Every name under `root` (recursively) with a stage or backup name shape (a Possible artifact: the
// Local provider's .rs_copy_tmp_, the bridge's .rs_tmp_, ...); listing failures count as found.
[[nodiscard]] std::vector<std::wstring> FindStageNames(const std::filesystem::path& root)
{
    std::vector<std::wstring> found;
    std::error_code error;
    std::filesystem::recursive_directory_iterator it(root, std::filesystem::directory_options::none, error);
    for (; ! error && it != std::filesystem::recursive_directory_iterator(); it.increment(error))
    {
        if (FileOperationArtifacts::HasPossibleArtifactName(it->path().filename().native()))
            found.push_back(it->path().native());
    }
    if (error && error != std::errc::no_such_file_or_directory)
        found.push_back(std::format(L"<listing failed: {}>", error.value()));
    return found;
}

// A cancelled Move leaves each file either moved (the complete new content published, the source
// gone or kept) or untouched (the destination's own content, the source kept); never a partial file.
[[nodiscard]] std::optional<std::wstring> CheckMoveOutcome(const std::filesystem::path& sourcePayload,
                                                           const std::filesystem::path& destinationPayload,
                                                           bool destinationExisted,
                                                           uint64_t seedBase,
                                                           unsigned int& movedCount)
{
    for (const MoveFixtureFile& file : kMoveFixtureFiles)
    {
        const std::filesystem::path source      = sourcePayload / file.relative;
        const std::filesystem::path destination = destinationPayload / file.relative;
        std::error_code error;
        const bool destinationPresent = std::filesystem::exists(destination, error);
        const bool destinationNew     = destinationPresent && ThroughputFileMatches(destination, kMoveNewBytes, seedBase + file.newSeed);
        const bool destinationOld     = destinationPresent && destinationExisted && ThroughputFileMatches(destination, kMoveOldBytes, seedBase + file.oldSeed);
        if (destinationPresent ? ! (destinationNew || destinationOld) : destinationExisted)
            return std::format(L"'{}' is neither its new nor its previous content at the destination (present={})", file.relative, destinationPresent);
        const bool sourcePresent = std::filesystem::exists(source, error);
        if (sourcePresent && ! ThroughputFileMatches(source, kMoveNewBytes, seedBase + file.newSeed))
            return std::format(L"the source of '{}' changed", file.relative);
        if (! sourcePresent && ! destinationNew)
            return std::format(L"the source of '{}' was removed without its content published", file.relative);
        movedCount += ! sourcePresent ? 1u : 0u;
    }
    return std::nullopt;
}

// ---------------------------------------------------------------------------------------------
// R10.2 deterministic display gates: synthetic 100 ms telemetry ticks through the production
// per-task rate step (DebugRunFileOperationsRateModel), no clock and no task involved.

using RateTick        = FileOperationsPopupInternal::RateModelDebugTick;
using RateObservation = FileOperationsPopupInternal::RateModelDebugObservation;
using RateStall       = FileOperationsPopupInternal::RateStallKind;

constexpr ULONGLONG kRateTickMs = 100ull;
constexpr double kRateMiB       = 1024.0 * 1024.0;

// Builds tick sequences segment by segment: a segment delivers a constant amount per tick (zero is
// a silence) and every tick records the progress the task would have published by then.
class RateShape final
{
public:
    void Run(ULONGLONG durationMs, uint64_t bytesPerTick, RateStall stall = RateStall::WaitingForStorage, unsigned long itemsPerTick = 0u)
    {
        for (ULONGLONG elapsed = kRateTickMs; elapsed <= durationMs; elapsed += kRateTickMs)
        {
            _tick += kRateTickMs;
            if (bytesPerTick != 0u || itemsPerTick != 0u)
            {
                _bytes += bytesPerTick;
                _items += itemsPerTick;
                _lastCallbackTick = _tick;
                if (bytesPerTick != 0u)
                    _lastByteTick = _tick;
            }
            ticks.push_back(RateTick{.tick                     = _tick,
                                     .completedBytes           = _bytes,
                                     .totalBytes               = totalBytes,
                                     .completedItems           = _items,
                                     .lastProgressCallbackTick = _lastCallbackTick,
                                     .lastByteProgressTick     = _lastByteTick,
                                     .progressStateChangeTick  = _stateChangeTick,
                                     .stallKind                = stall,
                                     .speedLimitBytesPerSecond = speedLimitBytesPerSecond,
                                     .paused                   = paused,
                                     .waitingForDecision       = waitingForDecision,
                                     .stopping                 = stopping});
        }
    }

    // One report of `bytes` after `gapMs` without any (gapMs is a multiple of the tick).
    void Report(ULONGLONG gapMs, uint64_t bytes, RateStall stall = RateStall::WaitingForDestination)
    {
        if (gapMs > kRateTickMs)
            Run(gapMs - kRateTickMs, 0u, stall);
        Run(kRateTickMs, bytes, stall);
    }

    // Pause and resume are state changes the task publishes at that instant.
    void SetPaused(bool value) noexcept
    {
        paused           = value;
        _stateChangeTick = _tick;
    }

    std::vector<RateTick> ticks;
    uint64_t totalBytes               = 0u;
    uint64_t speedLimitBytesPerSecond = 0u;
    bool paused                       = false;
    bool waitingForDecision           = false;
    bool stopping                     = false;

private:
    ULONGLONG _tick             = 100'000ull;
    ULONGLONG _lastCallbackTick = 0ull;
    ULONGLONG _lastByteTick     = 0ull;
    ULONGLONG _stateChangeTick  = 0ull;
    uint64_t _bytes             = 0u;
    unsigned long _items        = 0u;
};

[[nodiscard]] std::vector<RateObservation> RunRateShape(const RateShape& shape, FileSystemOperation operation = FILESYSTEM_COPY)
{
    std::vector<RateObservation> observations;
    if (! DebugRunFileOperationsRateModel(operation, shape.ticks, observations) || observations.size() != shape.ticks.size())
        observations.clear();
    return observations;
}

[[nodiscard]] bool RateNear(double value, double expected, double relativeTolerance) noexcept
{
    return std::abs(value - expected) <= std::abs(expected) * relativeTolerance;
}

[[nodiscard]] std::wstring FormatRateMiB(double bytesPerSecond)
{
    return std::format(L"{:.3f} MiB/s", bytesPerSecond / kRateMiB);
}

// A 2.5 s silence inside a steady transfer: the speed falls honestly (the window average, never held),
// the stall is named after 600 ms and hatched from where the bytes stopped, and the recovery climbs
// back without a snap or an overshoot.
[[nodiscard]] std::optional<std::wstring> CheckRateShortStall()
{
    constexpr uint64_t kPerTick = 10ull * 1024ull * 1024ull; // 100 MiB/s
    constexpr double kRate      = 100.0 * kRateMiB;
    RateShape shape;
    shape.Run(10'000ull, kPerTick, RateStall::WaitingForDestination);
    const size_t steadyEnd = shape.ticks.size() - 1u;
    shape.Run(2'500ull, 0u, RateStall::WaitingForDestination);
    const size_t silenceEnd = shape.ticks.size() - 1u;
    shape.Run(6'000ull, kPerTick, RateStall::WaitingForDestination);
    const auto observed = RunRateShape(shape);
    if (observed.empty())
        return L"the rate model rejected the short-stall shape";

    const RateObservation& steady = observed[steadyEnd];
    if (! RateNear(steady.displayedBytesPerSec, kRate, 0.005) || steady.stallKind != RateStall::None || steady.stalledSampleCount != 0u ||
        steady.rateLabelText != DebugFormatFileOperationsRateLabelText(steady.displayedBytesPerSec, RateStall::None))
        return std::format(L"steady 100 MiB/s showed {} (stall={} band={} label='{}')",
                           FormatRateMiB(steady.displayedBytesPerSec),
                           static_cast<unsigned int>(steady.stallKind),
                           steady.stalledSampleCount,
                           steady.rateLabelText);
    const std::wstring stallText = LoadStringResource(nullptr, IDS_FILEOPS_STALL_WAITING_DESTINATION);
    for (size_t index = steadyEnd + 1u; index <= silenceEnd; ++index)
    {
        const RateObservation& now = observed[index];
        const ULONGLONG silentMs   = (index - steadyEnd) * kRateTickMs;
        if (now.displayedBytesPerSec >= observed[index - 1u].displayedBytesPerSec)
            return std::format(L"the speed was held during a silence ({} ms: {})", silentMs, FormatRateMiB(now.displayedBytesPerSec));
        const bool named = now.stallKind == RateStall::WaitingForDestination && now.rateLabelText.find(stallText) != std::wstring::npos;
        if ((silentMs > 600u) != named)
            return std::format(L"after {} ms without bytes the stall was {}named (label='{}')", silentMs, named ? L"" : L"not ", now.rateLabelText);
        if (now.newestSample < 0.45f * static_cast<float>(kRate))
            return std::format(L"a {} ms silence was plotted as a trough ({} plotted)", silentMs, FormatRateMiB(now.newestSample));
    }
    const RateObservation& stalled = observed[silenceEnd];
    if (! RateNear(stalled.displayedBytesPerSec, 0.5 * kRate, 0.015))
        return std::format(L"2.5 s of silence in a 5 s window should halve 100 MiB/s, saw {}", FormatRateMiB(stalled.displayedBytesPerSec));
    if (stalled.stalledSampleCount < 24u || stalled.stalledSampleCount > 26u || ! stalled.newestSampleStalled)
        return std::format(L"the stall band should cover the 25 silent buckets from the last byte, saw {}", stalled.stalledSampleCount);
    for (size_t index = silenceEnd + 1u; index < observed.size(); ++index)
    {
        const RateObservation& now = observed[index];
        const double rise          = now.displayedBytesPerSec - observed[index - 1u].displayedBytesPerSec;
        if (now.stallKind != RateStall::None || rise < -0.001 * kRate || rise > 0.025 * kRate || now.displayedBytesPerSec > kRate * 1.005)
            return std::format(L"the recovery after a short stall snapped, overshot or stayed named (tick {}: {} rise {})",
                               index - silenceEnd,
                               FormatRateMiB(now.displayedBytesPerSec),
                               FormatRateMiB(rise));
    }
    if (! RateNear(observed.back().displayedBytesPerSec, kRate, 0.005))
        return std::format(L"5 s after the stall the window should hold only the resumed rate, saw {}", FormatRateMiB(observed.back().displayedBytesPerSec));
    return std::nullopt;
}

// A stall longer than the window: the speed reaches zero (Rate: — with the stall named), and when bytes
// resume the measurement restarts at the stall end, so the new rate is shown as soon as it is measured,
// neither averaged with the empty window (an undershoot) nor taken over from before the stall (a snap).
[[nodiscard]] std::optional<std::wstring> CheckRateLongStallRestart()
{
    constexpr double kResumedRate = 50.0 * kRateMiB;
    RateShape shape;
    shape.totalBytes = 64ull * 1024ull * 1024ull * 1024ull;
    shape.Run(10'000ull, 10ull * 1024ull * 1024ull, RateStall::Finalizing);
    const size_t steadyEnd = shape.ticks.size() - 1u;
    shape.Run(8'000ull, 0u, RateStall::Finalizing);
    const size_t silenceEnd = shape.ticks.size() - 1u;
    shape.Run(6'000ull, 5ull * 1024ull * 1024ull, RateStall::Finalizing);
    const auto observed = RunRateShape(shape);
    if (observed.empty())
        return L"the rate model rejected the long-stall shape";

    const RateObservation& emptied = observed[steadyEnd + 50u];
    if (emptied.displayedBytesPerSec != 0.0 || emptied.hasSmoothedEta || emptied.stallKind != RateStall::Finalizing ||
        emptied.rateLabelText != DebugFormatFileOperationsRateLabelText(0.0, RateStall::Finalizing))
        return std::format(L"a window holding only a stall should show no speed, no ETA and the stall (saw {} eta={} label='{}')",
                           FormatRateMiB(emptied.displayedBytesPerSec),
                           emptied.hasSmoothedEta,
                           emptied.rateLabelText);
    for (size_t index = silenceEnd + 1u; index < observed.size(); ++index)
    {
        const RateObservation& now = observed[index];
        const ULONGLONG resumedMs  = (index - silenceEnd) * kRateTickMs;
        if (now.stallKind != RateStall::None)
            return std::format(L"the stall stayed named {} ms after bytes resumed", resumedMs);
        const bool measured = resumedMs >= 1'000u;
        if (! measured && now.displayedBytesPerSec != 0.0)
            return std::format(
                L"{} ms after a long stall the speed was {} before a full second was measured", resumedMs, FormatRateMiB(now.displayedBytesPerSec));
        if (measured && ! RateNear(now.displayedBytesPerSec, kResumedRate, 0.005))
            return std::format(L"{} ms after a long stall the speed was {}, expected the resumed 50 MiB/s", resumedMs, FormatRateMiB(now.displayedBytesPerSec));
    }
    return std::nullopt;
}

// The SMB sawtooth: a 16 MiB burst, then 2.5 s of silence, over and over. The old hold-and-decay
// display swung between ~11 % of the true rate and a full burst rate (red before R10.2); the window
// average stays within one burst of the true rate and is never plotted as a trough. Whether a silence
// is named depends on the stream's own cadence: when every report is one burst, 2.6 s between reports
// is the cadence and nothing is a stall; when a burst is several reports 100 ms apart, the 2.2 s
// silence after it is far longer than the cadence and is named.
[[nodiscard]] std::optional<std::wstring> CheckRateBurstSawtooth()
{
    constexpr uint64_t kBurstBytes = 16ull * 1024ull * 1024ull;
    constexpr ULONGLONG kPeriodMs  = 2'600ull;
    constexpr double kTrueRate     = static_cast<double>(kBurstBytes) / (static_cast<double>(kPeriodMs) / 1000.0);
    for (const unsigned int reportsPerBurst : {1u, 4u})
    {
        RateShape shape;
        for (size_t burst = 0u; burst < 24u; ++burst)
        {
            for (unsigned int report = 0u; report < reportsPerBurst; ++report)
                shape.Run(kRateTickMs, kBurstBytes / reportsPerBurst, RateStall::WaitingForDestination);
            shape.Run(kPeriodMs - kRateTickMs * reportsPerBurst, 0u, RateStall::WaitingForDestination);
        }
        const auto observed = RunRateShape(shape);
        if (observed.empty())
            return L"the rate model rejected the sawtooth shape";

        size_t namedTicks = 0u;
        for (size_t index = 60u; index < observed.size(); ++index)
        {
            const RateObservation& now = observed[index];
            const double ratio         = now.displayedBytesPerSec / kTrueRate;
            if (ratio < 0.45 || ratio > 1.10)
                return std::format(L"the sawtooth ({} report(s) a burst) displayed {} ({:.2f} x the true {}) at tick {}",
                                   reportsPerBurst,
                                   FormatRateMiB(now.displayedBytesPerSec),
                                   ratio,
                                   FormatRateMiB(kTrueRate),
                                   index);
            if (static_cast<double>(now.newestSample) < 0.45 * kTrueRate || static_cast<double>(now.newestSample) > 1.10 * kTrueRate)
                return std::format(L"the sawtooth ({} report(s) a burst) plotted {} at tick {} (true {})",
                                   reportsPerBurst,
                                   FormatRateMiB(now.newestSample),
                                   index,
                                   FormatRateMiB(kTrueRate));
            namedTicks += now.stallKind == RateStall::WaitingForDestination ? 1u : 0u;
        }
        if (reportsPerBurst == 1u && namedTicks != 0u)
            return std::format(L"a stream whose every report is a burst 2.6 s apart was named stalled on {} ticks (its cadence, not a stall)", namedTicks);
        if (reportsPerBurst > 1u && namedTicks == 0u)
            return L"the 2.2 s silences after bursts of 100 ms reports were never named as stalls";
    }
    return std::nullopt;
}

// A transfer held at its speed limit: the throttle releases a 1 MiB chunk once a second (the first
// limit preset), so reports are a second apart. That is the limit at work, never a stall (red before
// the fix: "Waiting for the destination..." every second and ~40 % of the graph hatched).
[[nodiscard]] std::optional<std::wstring> CheckRateThrottledChunks()
{
    constexpr uint64_t kChunk = 1024ull * 1024ull;
    RateShape shape;
    shape.speedLimitBytesPerSecond = kChunk;
    for (size_t chunk = 0u; chunk < 20u; ++chunk)
        shape.Report(1'000ull, kChunk);
    const auto observed = RunRateShape(shape);
    if (observed.empty())
        return L"the rate model rejected the throttled shape";
    for (size_t index = 0u; index < observed.size(); ++index)
    {
        if (observed[index].stallKind != RateStall::None || observed[index].stalledSampleCount != 0u)
            return std::format(L"a transfer held at its 1 MiB/s limit was named stalled {} ms in (label='{}', band={})",
                               (index + 1u) * kRateTickMs,
                               observed[index].rateLabelText,
                               observed[index].stalledSampleCount);
    }
    if (! RateNear(observed.back().displayedBytesPerSec, static_cast<double>(kChunk), 0.05))
        return std::format(L"a transfer held at 1 MiB/s showed {}", FormatRateMiB(observed.back().displayedBytesPerSec));
    return std::nullopt;
}

// A slow destination that reports each 1 MiB chunk 800-900 ms apart (about 1.2 MiB/s: a USB 2.0 drive,
// a share over a WAN). Reports coarser than the 600 ms stall threshold are its steady cadence, never a
// stall (red before the fix).
[[nodiscard]] std::optional<std::wstring> CheckRateSlowSteadyStream()
{
    constexpr uint64_t kChunk = 1024ull * 1024ull;
    RateShape shape;
    for (size_t chunk = 0u; chunk < 24u; ++chunk)
        shape.Report(chunk % 2u == 0u ? 800ull : 900ull, kChunk);
    const auto observed = RunRateShape(shape);
    if (observed.empty())
        return L"the rate model rejected the slow steady shape";
    for (size_t index = 0u; index < observed.size(); ++index)
    {
        if (observed[index].stallKind != RateStall::None || observed[index].stalledSampleCount != 0u)
            return std::format(L"a steady 1.2 MiB/s stream reporting every 850 ms was named stalled {} ms in (label='{}', band={})",
                               (index + 1u) * kRateTickMs,
                               observed[index].rateLabelText,
                               observed[index].stalledSampleCount);
    }
    if (! RateNear(observed.back().displayedBytesPerSec, static_cast<double>(kChunk) / 0.85, 0.12))
        return std::format(L"a steady stream of about 1.2 MiB/s showed {}", FormatRateMiB(observed.back().displayedBytesPerSec));
    return std::nullopt;
}

// A real silence still is a stall: 3 s without bytes after steady reports every 100 ms is named once it
// passes 600 ms, and a transfer held at its 1 MiB/s limit is named once its silence is clearly longer
// than its one-second cadence and its speed has fallen below the limit.
[[nodiscard]] std::optional<std::wstring> CheckRateRealSilenceAfterSteady()
{
    constexpr uint64_t kChunk = 1024ull * 1024ull;
    RateShape fast;
    fast.Run(5'000ull, kChunk, RateStall::WaitingForDestination);
    const size_t fastSteadyEnd = fast.ticks.size() - 1u;
    fast.Run(3'000ull, 0u, RateStall::WaitingForDestination);
    const auto fastObserved = RunRateShape(fast);
    if (fastObserved.empty())
        return L"the rate model rejected the fast silence shape";
    for (size_t index = fastSteadyEnd + 1u; index < fastObserved.size(); ++index)
    {
        const ULONGLONG silentMs = (index - fastSteadyEnd) * kRateTickMs;
        const bool named         = fastObserved[index].stallKind == RateStall::WaitingForDestination;
        if ((silentMs > 600u) != named)
            return std::format(L"after {} ms without bytes (steady reports every 100 ms before) the stall was {}named", silentMs, named ? L"" : L"not ");
    }

    RateShape throttled;
    throttled.speedLimitBytesPerSecond = kChunk;
    for (size_t chunk = 0u; chunk < 10u; ++chunk)
        throttled.Report(1'000ull, kChunk);
    const size_t lastChunk = throttled.ticks.size() - 1u;
    throttled.Run(4'000ull, 0u, RateStall::WaitingForDestination);
    const auto throttledObserved = RunRateShape(throttled);
    if (throttledObserved.empty())
        return L"the rate model rejected the throttled silence shape";
    ULONGLONG firstNamedMs = 0u;
    for (size_t index = lastChunk + 1u; index < throttledObserved.size(); ++index)
    {
        const ULONGLONG silentMs = (index - lastChunk) * kRateTickMs;
        const bool named         = throttledObserved[index].stallKind == RateStall::WaitingForDestination;
        if (firstNamedMs == 0u && named)
            firstNamedMs = silentMs;
        else if (firstNamedMs != 0u && ! named)
            return std::format(L"a throttled transfer's silence stopped being named after {} ms while still silent", silentMs);
    }
    if (firstNamedMs <= 1'500u || firstNamedMs > 3'000u)
        return std::format(L"a 4 s silence of a transfer held at 1 MiB/s was named after {} ms (expected past its 1.5 s cadence threshold, within 3 s)",
                           firstNamedMs);
    return std::nullopt;
}

// A stopping task waits on its own abort (a slow cleanup of what it started): its silence is never
// named "Waiting for the destination..." (red before the fix).
[[nodiscard]] std::optional<std::wstring> CheckRateStoppingNoStall()
{
    RateShape shape;
    shape.Run(5'000ull, 10ull * 1024ull * 1024ull, RateStall::WaitingForDestination);
    shape.stopping            = true;
    const size_t stoppingFrom = shape.ticks.size();
    shape.Run(4'000ull, 0u, RateStall::WaitingForDestination);
    const auto observed = RunRateShape(shape);
    if (observed.empty())
        return L"the rate model rejected the stopping shape";
    for (size_t index = stoppingFrom; index < observed.size(); ++index)
    {
        if (observed[index].stallKind != RateStall::None)
            return std::format(L"a stopping task was named stalled {} ms into its abort (label='{}')",
                               (index - stoppingFrom + 1u) * kRateTickMs,
                               observed[index].rateLabelText);
    }
    return std::nullopt;
}

// Pause, a question for the user and a resume: nothing moves by design, so nothing is shown and
// nothing is a stall; after the resume the speed is measured from the resume.
[[nodiscard]] std::optional<std::wstring> CheckRateBlockedStates()
{
    constexpr double kRate = 20.0 * kRateMiB;
    RateShape shape;
    shape.Run(5'000ull, 2ull * 1024ull * 1024ull);
    shape.SetPaused(true);
    const size_t pausedFrom = shape.ticks.size();
    shape.Run(3'000ull, 0u);
    shape.SetPaused(false);
    const size_t resumedFrom = shape.ticks.size();
    shape.Run(3'000ull, 2ull * 1024ull * 1024ull);
    shape.waitingForDecision  = true;
    const size_t decisionFrom = shape.ticks.size();
    shape.Run(10'000ull, 0u);
    const auto observed = RunRateShape(shape);
    if (observed.empty())
        return L"the rate model rejected the blocked-state shape";

    const std::wstring unavailable = LoadStringResource(nullptr, IDS_FILEOPS_RATE_UNAVAILABLE);
    for (size_t index = pausedFrom; index < resumedFrom; ++index)
    {
        if (observed[index].displayedBytesPerSec != 0.0 || observed[index].stallKind != RateStall::None || observed[index].rateLabelText != unavailable)
            return std::format(L"a paused task showed {} (stall={} label='{}')",
                               FormatRateMiB(observed[index].displayedBytesPerSec),
                               static_cast<unsigned int>(observed[index].stallKind),
                               observed[index].rateLabelText);
    }
    for (size_t index = resumedFrom; index < decisionFrom; ++index)
    {
        // Measuring restarts at the first tick that sees the resumed task.
        const ULONGLONG measuredMs = (index - resumedFrom) * kRateTickMs;
        const double shown         = observed[index].displayedBytesPerSec;
        if (measuredMs < 1'000u ? shown != 0.0 : ! RateNear(shown, kRate, 0.005))
            return std::format(L"{} ms into a resume the speed was {} (expected it measured from the resume)", measuredMs, FormatRateMiB(shown));
    }
    for (size_t index = decisionFrom; index < observed.size(); ++index)
    {
        if (observed[index].displayedBytesPerSec != 0.0 || observed[index].stallKind != RateStall::None)
            return std::format(L"a task waiting for a decision showed {} or a stall ({})",
                               FormatRateMiB(observed[index].displayedBytesPerSec),
                               static_cast<unsigned int>(observed[index].stallKind));
    }
    return std::nullopt;
}

// The 40 s rule (UI_FileOperationsPopup.md): a long silence ends at exactly zero with no ETA, and a
// deletion's item rate uses the same window but never names a transfer stall.
[[nodiscard]] std::optional<std::wstring> CheckRateLongSilenceAndItems()
{
    const double silentRate = DebugWindowedRateAfterSilence(100.0 * kRateMiB, 40'000ull);
    if (silentRate != 0.0)
        return std::format(L"40 s of silence should end at exactly zero, saw {}", FormatRateMiB(silentRate));
    if (! RateNear(DebugWindowedRateAfterSilence(100.0 * kRateMiB, 0ull), 100.0 * kRateMiB, 0.005))
        return L"a steady rate was not reported exactly";

    RateShape copy;
    copy.totalBytes = 64ull * 1024ull * 1024ull * 1024ull;
    copy.Run(10'000ull, 10ull * 1024ull * 1024ull);
    copy.Run(40'000ull, 0u);
    const auto copied = RunRateShape(copy);
    if (copied.empty() || copied.back().displayedBytesPerSec != 0.0 || copied.back().hasSmoothedEta)
        return L"after 40 s of silence a Copy kept a speed or an ETA";

    RateShape deletion;
    deletion.Run(6'000ull, 0u, RateStall::WaitingForStorage, 3u);
    deletion.Run(3'000ull, 0u, RateStall::WaitingForStorage, 0u);
    const auto deleted = RunRateShape(deletion, FILESYSTEM_DELETE);
    if (deleted.empty() || ! RateNear(deleted[59].displayedItemsPerSec, 30.0, 0.01))
        return std::format(L"30 deletions a second showed {:.2f} items/s", deleted.empty() ? -1.0 : deleted[59].displayedItemsPerSec);
    for (const RateObservation& now : deleted)
    {
        if (now.stallKind != RateStall::None)
            return L"a deletion named a transfer stall";
    }
    return std::nullopt;
}

// R10.1 engine fact: every file whose commit outlasts the 50 ms boundary gap reports again after its
// commit. With every Local commit slowed to 150 ms, each witness file repeats its Finalizing report at
// least twice (the post-commit boundary and its closing report); with the boundary reports suppressed
// (the negative control) only the closing report remains, and the gate must tell the two apart.
[[nodiscard]] std::optional<std::wstring> CheckPhaseBoundaryWitness(SelfTestState& state,
                                                                    const std::filesystem::path& destinationRoot,
                                                                    std::wstring_view target)
{
    constexpr std::wstring_view kSlowCommitEnv = L"REDSALAMANDER_FILEOPS_SLOW_COMMIT_MS";
    constexpr std::wstring_view kSuppressEnv   = L"REDSALAMANDER_FILEOPS_SUPPRESS_PHASE_BOUNDARIES";
    constexpr ThroughputCorpus kWitnessCorpus{L"witness", 4ull * kThroughputMiB, 2u, 4ull * kThroughputMiB};
    constexpr uint64_t kWitnessFiles = 3u;
    auto* operations                 = TryGetFileOps(TryGetFolderWindow(state.mainWindow));
    if (! operations || ! CancelOutstandingSelfTestOperations(operations))
        return L"the phase-boundary witness found no quiet owner";
    const std::filesystem::path corpusParent = state.tempRoot / L"phase-witness-source";
    const auto cleanup                       = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kSlowCommitEnv.data(), nullptr));
        static_cast<void>(SetEnvironmentVariableW(kSuppressEnv.data(), nullptr));
        static_cast<void>(SelfTest::RemoveAll(corpusParent));
    });
    if (! CreateThroughputCorpus(corpusParent / L"corpus", kWitnessCorpus))
        return L"the phase-boundary witness corpus could not be created";
    ThroughputCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    SetFileOpsTransferTuningOverrideForSelfTest(std::nullopt);
    static_cast<void>(SetEnvironmentVariableW(kSlowCommitEnv.data(), L"150"));
    std::array<uint64_t, 2> repeats{};
    for (const bool suppressed : {false, true})
    {
        static_cast<void>(SetEnvironmentVariableW(kSuppressEnv.data(), suppressed ? L"1" : nullptr));
        Sleep(250u); // the Local provider refreshes its self-test knobs at most every 100 ms
        std::wstring failure;
        const auto sample = RunThroughputCopy(state,
                                              operations,
                                              scope,
                                              corpusParent / L"corpus",
                                              kWitnessCorpus,
                                              destinationRoot / (suppressed ? L"witness-suppressed" : L"witness"),
                                              target,
                                              ThroughputVariant{L"adopted", std::nullopt},
                                              false,
                                              failure);
        if (! sample.has_value())
            return std::format(L"phase-boundary witness ({}): {}", suppressed ? L"suppressed" : L"normal", failure);
        repeats[suppressed ? 1u : 0u] = sample->cadence.finalizingRepeatCount;
    }
    AppendLog(std::format(
        L"Phase-boundary witness: target={} files={} finalizingRepeats={} (boundaries suppressed: {})", target, kWitnessFiles, repeats[0], repeats[1]));
    if (repeats[0] < 2u * kWitnessFiles)
        return std::format(
            L"with 150 ms commits the host saw {} Finalizing repeats for {} files (at least {} expected: a post-commit boundary and a closing report each)",
            repeats[0],
            kWitnessFiles,
            2u * kWitnessFiles);
    if (repeats[1] >= 2u * kWitnessFiles)
        return std::format(L"with the boundary reports suppressed the host still saw {} Finalizing repeats: the witness cannot see a missing boundary",
                           repeats[1]);
    return std::nullopt;
}
} // namespace

// R10.5: the corpus copied from the local test root to the registered SMB self-test share (a unique
// per-run child, removed afterwards), or to the loopback administrative-share alias when no share is
// registered. Reports the report cadence and whole-second throughput variation of every run.
[[nodiscard]] ThroughputCaseResult RunFileOpsSmbCopyThroughputVarianceChecks(SelfTestState& state) noexcept
{
    const ThroughputShareTarget target = ResolveThroughputShareTarget(state, L"throughput", state.tempRoot / L"throughput-loopback");
    if (target.unavailable.has_value())
        return *target.unavailable;
    const auto removeDestinationRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(target.root)); });
    std::vector<ThroughputSample> samples;
    ThroughputCaseResult result = RunThroughputSchedule(state, L"smb-throughput", target.name, target.root, samples);
    if (! result.passed)
        return result;
    AppendLog(std::format(L"SMB throughput: {} run(s) to the {} target.", samples.size(), target.name));
    if (const auto witness = CheckPhaseBoundaryWitness(state, target.root, target.name); witness.has_value())
        return {.reason = *witness};
    return result;
}

// R10.6 reference: the same corpus copied within the local test root, so a candidate that helps the
// share can be held to the local-throughput regression bound.
[[nodiscard]] ThroughputCaseResult RunFileOpsLocalCopyThroughputChecks(SelfTestState& state) noexcept
{
    const std::filesystem::path destinationRoot = state.tempRoot / L"local-throughput";
    const auto removeDestinationRoot            = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(destinationRoot)); });
    std::vector<ThroughputSample> samples;
    ThroughputCaseResult result = RunThroughputSchedule(state, L"local-throughput", L"local", destinationRoot, samples);
    if (! result.passed)
        return result;
    if (const auto witness = CheckPhaseBoundaryWitness(state, destinationRoot, L"local"); witness.has_value())
        return {.reason = *witness};
    return result;
}

// I35 S6: the host's progress accounting with synthetic reports (no task runs). A report without item
// bytes names no phase, so a silence after it is attributed to Unknown (and the R10.5 engine-fact gate
// can fail); a zero-delta Finalizing repeat is counted; and a file completed while discovery is open is
// counted once even when a report arrives after its in-flight entry expired (a commit over 300 ms).
[[nodiscard]] ThroughputCaseResult RunFileOpsProgressPhaseAccountingChecks(SelfTestState& state) noexcept
{
    auto* operations = TryGetFileOps(TryGetFolderWindow(state.mainWindow));
    if (! operations)
        return {.reason = L"owner unavailable"};
    const FileOpsProgressAccountingDebugResult result = DebugRunFileOpsProgressAccountingForSelfTest(*operations);
    AppendLog(std::format(L"Progress accounting: unsizedPhase={} unknownSilences={} unknownStalls={} finalizingRepeats={} completionsWhileOpen={}",
                          static_cast<unsigned int>(result.unsizedReportPhase),
                          result.unknownSilences,
                          result.unknownStalls,
                          result.finalizingRepeats,
                          result.completionsCountedWhileOpen));
    std::wstring failures;
    const auto fail = [&](std::wstring text) { failures += std::format(L"{}{}", failures.empty() ? L"" : L"; ", text); };
    if (result.unsizedReportPhase != ThroughputPhase::Unknown)
        fail(std::format(L"a report without item bytes announced phase {} (it names none)", static_cast<unsigned int>(result.unsizedReportPhase)));
    if (result.unknownSilences != 1u || result.unknownStalls != 1u)
        fail(std::format(L"the 700 ms silence after a report without item bytes was attributed to a named phase (Unknown silences {}, stalls {})",
                         result.unknownSilences,
                         result.unknownStalls));
    if (result.finalizingRepeats != 1u)
        fail(std::format(L"a zero-delta Finalizing repeat was counted {} times", result.finalizingRepeats));
    if (result.completionsCountedWhileOpen != 2u)
        fail(std::format(L"two files completed while discovery was open were counted {} times (a report after the 300 ms in-flight expiry counted one again)",
                         result.completionsCountedWhileOpen));
    if (! failures.empty())
        return {.reason = failures};
    return {.passed = true};
}

// I35 R10.2: the stall name never moves the graph. A rate label that grows by "· Waiting for the
// destination..." keeps the label rows and the plot where they were at every popup width, and when the
// labels have no row above the plot (the graph draws them inside it) the stall band never hatches them.
[[nodiscard]] ThroughputCaseResult RunFileOpsStallLabelLayoutChecks(SelfTestState& state) noexcept
{
    using FileOperationsPopupInternal::FileOperationsPopupState;
    using FileOperationsPopupInternal::PopupLayoutDebugSnapshot;
    using FileOperationsPopupInternal::TaskSnapshot;
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations)
        return {.reason = L"owner unavailable"};
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return {.reason = L"no-activation guard unavailable"};
    auto hostLease = std::make_shared<int>(1);
    wil::unique_hwnd popup(FileOperationsPopup::Create(operations, folder, state.mainWindow, hostLease));
    if (! popup)
        return {.reason = L"synthetic popup could not be created"};
    const auto destroyPopup = wil::scope_exit([&]() noexcept
    {
        popup.reset();
        hostLease.reset();
    });
    auto* popupState        = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA));
    if (! popupState)
        return {.reason = L"synthetic popup state is unavailable"};

    constexpr uint64_t kTaskId = 0xF26A202610120056ull;
    TaskSnapshot task{};
    task.taskId                   = kTaskId;
    task.operation                = FILESYSTEM_COPY;
    task.lifecyclePhase           = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Running);
    task.started                  = true;
    task.discoveryClosed          = true;
    task.hasProgressCallbacks     = true;
    task.lastProgressCallbackTick = GetTickCount64();
    task.operationStartTick       = task.lastProgressCallbackTick > 250u ? task.lastProgressCallbackTick - 250u : 0u;
    task.currentSourcePath        = L"C:\\Data\\Archive\\payload.bin";
    task.currentDestinationPath   = L"\\\\server\\share\\Archive\\payload.bin";
    task.totalBytes               = 64ull * 1024ull * 1024ull * 1024ull;
    task.completedBytes           = 16ull * 1024ull * 1024ull * 1024ull;
    task.totalItems               = 2u;
    task.completedItems           = 1u;
    popupState->DebugSetVisualScenario({task}, false, true);
    ShowWindow(popup.get(), SW_SHOWNOACTIVATE);

    const UINT dpi    = GetDpiForWindow(popup.get());
    const auto resize = [&](int widthDip) noexcept
    {
        RECT desired{0, 0, MulDiv(widthDip, static_cast<int>(dpi), 96), MulDiv(520, static_cast<int>(dpi), 96)};
        return AdjustWindowRectExForDpi(&desired,
                                        static_cast<DWORD>(GetWindowLongPtrW(popup.get(), GWL_STYLE)),
                                        FALSE,
                                        static_cast<DWORD>(GetWindowLongPtrW(popup.get(), GWL_EXSTYLE)),
                                        dpi) != FALSE &&
               SetWindowPos(
                   popup.get(), nullptr, 0, 0, desired.right - desired.left, desired.bottom - desired.top, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE) != FALSE;
    };
    // The fixture's speed with and without the stall name beside a long remaining time, so that across
    // the popup's widths (from its 480 DIP minimum) the plain label fits a row beside the ETA where the
    // stalled one would not.
    constexpr double kFixtureRate    = 64.0 * 1024.0 * 1024.0;
    const std::wstring plainRate     = DebugFormatFileOperationsRateLabelText(kFixtureRate, RateStall::None);
    const std::wstring stalledRate   = DebugFormatFileOperationsRateLabelText(kFixtureRate, RateStall::WaitingForDestination);
    const std::wstring longRemaining = L"Time remaining: about 1 hour and 20 minutes";
    const auto snapshotWith          = [&](RateStall stall, PopupLayoutDebugSnapshot& out)
    {
        popupState->DebugSetVisualStallKind(kTaskId, stall);
        popupState->debugGraphLabelText = {stall == RateStall::None ? plainRate : stalledRate, longRemaining};
        for (unsigned int frame = 0u; frame < 3u; ++frame)
            RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        out        = {};
        out.taskId = kTaskId;
        return DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), out);
    };
    const auto hasArea = [](const D2D1_RECT_F& rect) noexcept { return rect.right > rect.left && rect.bottom > rect.top; };
    // The band may hatch only below the labels, wherever they are.
    const auto bandClearOfLabels = [&](const PopupLayoutDebugSnapshot& layout) noexcept
    { return ! hasArea(layout.graphStallBandRect) || layout.graphStallBandRect.top + 0.5f >= layout.graphLabelRowBottomPx; };

    std::wstring failures;
    const auto fail = [&](std::wstring text)
    {
        if (failures.size() < 1'500u)
            failures += std::format(L"{}{}", failures.empty() ? L"" : L"; ", text);
    };
    size_t compared = 0u;
    for (int widthDip = 760; widthDip >= 480; widthDip -= 20)
    {
        PopupLayoutDebugSnapshot plain{};
        PopupLayoutDebugSnapshot stalled{};
        if (! resize(widthDip) || ! snapshotWith(RateStall::None, plain) || ! snapshotWith(RateStall::WaitingForDestination, stalled))
            return {.reason = std::format(L"the popup could not be laid out at {} DIP", widthDip)};
        if (! hasArea(plain.graphPlotRect))
            continue;
        ++compared;
        if (plain.graphPlotRect.top != stalled.graphPlotRect.top || plain.graphPlotRect.bottom != stalled.graphPlotRect.bottom ||
            plain.graphLabelsOutsidePlot != stalled.graphLabelsOutsidePlot)
            fail(std::format(L"at {} DIP the stall name moved the plot (top {:.1f} -> {:.1f}, labels above the plot {} -> {})",
                             widthDip,
                             plain.graphPlotRect.top,
                             stalled.graphPlotRect.top,
                             plain.graphLabelsOutsidePlot,
                             stalled.graphLabelsOutsidePlot));
        if (! bandClearOfLabels(stalled))
            fail(std::format(L"at {} DIP the stall band starts at {:.1f}, over the labels ending at {:.1f}",
                             widthDip,
                             stalled.graphStallBandRect.top,
                             stalled.graphLabelRowBottomPx));
    }
    if (compared == 0u)
        fail(L"no popup width showed the task's graph");

    // Ordinary Copy/Move cards now keep speed in the Throughput heading and ETA
    // beside the counts. Even long localized labels remain outside the plot;
    // the generic graph's in-plot fallback is a different presentation contract.
    std::wstring longRate;
    std::wstring longEta;
    for (unsigned int repeat = 0u; repeat < 6u; ++repeat)
    {
        longRate += L"Vitesse moyenne de transfert observée : 64,0 Mo par seconde · En attente de la destination. ";
        longEta += L"Temps restant estimé pour terminer la copie : environ une minute et vingt secondes. ";
    }
    const auto clearLabels = wil::scope_exit([&]() noexcept { popupState->debugGraphLabelText = {}; });
    PopupLayoutDebugSnapshot inside{};
    if (! resize(480))
        return {.reason = L"the popup could not be laid out with long labels"};
    popupState->DebugSetVisualStallKind(kTaskId, RateStall::WaitingForDestination);
    popupState->debugGraphLabelText = {std::move(longRate), std::move(longEta)};
    for (unsigned int frame = 0u; frame < 3u; ++frame)
        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    inside.taskId = kTaskId;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), inside))
        return {.reason = L"the popup could not be laid out with long labels"};
    if (! inside.graphLabelsOutsidePlot || ! hasArea(inside.graphRateLabelRect) || ! hasArea(inside.graphEtaLabelRect))
        fail(L"ordinary Copy/Move must retain both long localized labels in its task summary outside the plot");
    if (! bandClearOfLabels(inside))
        fail(std::format(L"with long task-summary labels the stall band starts at {:.1f}, over the labels ending at {:.1f}",
                         inside.graphStallBandRect.top,
                         inside.graphLabelRowBottomPx));
    if (! failures.empty())
        return {.reason = failures};
    return {.passed = true};
}

// R10.5: the native route's concurrency budget toward the loopback alias of the case root. The two
// sides report different storage preferences; with the destination consulted (R10.3 candidate c) the
// budget is min(source, destination) and records the destination's storage kind, as the bridge does;
// without it the source alone sets the budget, as it always did.
[[nodiscard]] ThroughputCaseResult RunFileOpsSmbNativeRouteConcurrencyBudgetChecks(SelfTestState& state) noexcept
{
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return {.reason = L"owner, Local provider or initial worker quiet point unavailable"};
    const auto reject                   = [](std::wstring_view reason) { return ThroughputCaseResult{.reason = std::wstring(reason)}; };
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(state.tempRoot);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the loopback administrative-share alias of the case root");
    if (uncStatus == LoopbackShareStatus::Unreachable)
        return {.skipped = true, .reason = L"the loopback administrative share is not reachable"};
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return {.reason = L"no-activation guard unavailable"};

    constexpr unsigned int kSourcePreferred      = 6u;
    constexpr unsigned int kDestinationPreferred = 2u;
    constexpr ThroughputCorpus kBudgetCorpus{L"budget", 1024ull * 1024ull, 12u, 256ull * 1024ull};
    const std::filesystem::path root         = state.tempRoot / L"native-concurrency";
    const std::filesystem::path sourceFolder = root / L"source" / L"corpus";
    const auto cleanup                       = wil::scope_exit([&]() noexcept
    {
        ClearFileOpsStoragePathConcurrencyOverridesForSelfTest();
        SetFileOpsTransferTuningOverrideForSelfTest(std::nullopt);
        static_cast<void>(SelfTest::RemoveAll(root));
    });
    if (! RecreateEmptyDirectory(root / L"destination") || ! CreateThroughputCorpus(sourceFolder, kBudgetCorpus))
        return {.reason = L"case root unavailable"};
    AddFileOpsStoragePathConcurrencyOverrideForSelfTest((root / L"source").native(), kSourcePreferred, FILESYSTEM_STORAGE_NVME);
    AddFileOpsStoragePathConcurrencyOverrideForSelfTest(LoopbackShareAlias(root).native(), kDestinationPreferred, FILESYSTEM_STORAGE_NETWORK_SHARE);

    struct BudgetExpectation final
    {
        ThroughputVariant variant;
        unsigned int budget      = 0u;
        uint32_t destinationKind = FILESYSTEM_STORAGE_UNKNOWN;
    };
    const std::array<BudgetExpectation, 2> expectations{{
        {{L"base", FILESYSTEM_TRANSFER_TUNING_NONE}, kSourcePreferred, FILESYSTEM_STORAGE_UNKNOWN},
        {{L"c", FILESYSTEM_TRANSFER_TUNING_DESTINATION_CONCURRENCY}, std::min(kSourcePreferred, kDestinationPreferred), FILESYSTEM_STORAGE_NETWORK_SHARE},
    }};
    ThroughputCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    for (size_t run = 0u; run < expectations.size(); ++run)
    {
        const BudgetExpectation& expected = expectations[run];
        SetFileOpsTransferTuningOverrideForSelfTest(expected.variant.tuning);
        std::wstring failure;
        const auto sample = RunThroughputCopy(state,
                                              operations,
                                              scope,
                                              sourceFolder,
                                              kBudgetCorpus,
                                              LoopbackShareAlias(root / L"destination" / std::format(L"run-{}", run)),
                                              L"loopback",
                                              expected.variant,
                                              false,
                                              failure);
        if (! sample.has_value())
            return {.reason = std::format(L"variant {}: {}", expected.variant.name, failure)};
        const CompletedTaskInfo& info = state.completedTasks.at(scope.tasks.back());
        AppendLog(std::format(L"Native route budget: variant={} budget={} destinationKind={} bridgeAdmissions={}",
                              expected.variant.name,
                              info.configuredMaxConcurrency,
                              info.autoConcurrencyDestinationStorageKind,
                              info.bridgeFileAdmissionCount));
        if (info.bridgeFileAdmissionCount != 0u || info.configuredMaxConcurrency != expected.budget ||
            info.autoConcurrencyDestinationStorageKind != expected.destinationKind)
            return {.reason = std::format(L"variant {}: the provider route ran with budget {} and destination storage kind {} (expected {} and {}; bridge "
                                          L"admissions {})",
                                          expected.variant.name,
                                          info.configuredMaxConcurrency,
                                          info.autoConcurrencyDestinationStorageKind,
                                          expected.budget,
                                          expected.destinationKind,
                                          info.bridgeFileAdmissionCount)};
    }
    return {.passed = true};
}

// A Move between two volumes served by one provider admits no destination provider of its own and runs
// through the host bridge (a Managed Move: a Local Move from a local volume to a NAS share is one). The
// source's provider then serves the destination, and the destination's storage still counts: the budget is
// min(source, destination) with the destination's storage kind recorded, and a stall is named after the
// destination, the only high-latency side. (The bridge used to consult a null destination provider: the
// source alone set the budget and a stall read "Waiting for storage…".) The destination is the loopback
// alias of the authorized alternate volume: another volume, so the Move cannot be a native rename (the
// alias of the case root's own volume is), and a network path, so it is the high-latency side.
[[nodiscard]] ThroughputCaseResult RunFileOpsSameProviderBridgeDestinationStorageChecks(SelfTestState& state) noexcept
{
    using StallSide  = FolderWindow::FileOperationState::ProgressStallSide;
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return {.reason = L"owner, Local provider or initial worker quiet point unavailable"};
    std::wstring alternateDetail;
    const std::optional<std::filesystem::path> alternate = TryCreateAlternateWritableVolumeSelfTestRoot(state.tempRoot, alternateDetail);
    if (! alternate.has_value())
        return {.skipped = true, .reason = std::format(L"no authorized alternate volume ({})", alternateDetail)};
    const auto removeAlternate          = wil::scope_exit([&]() noexcept
    {
        if (SelfTest::RemoveAll(alternate.value()))
            PruneEmptyAlternateVolumeSandboxParents(alternate.value());
    });
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(alternate.value());
    if (uncStatus == LoopbackShareStatus::Refused)
        return {.reason = L"the sandbox guard refused the loopback administrative-share alias of the alternate volume"};
    if (uncStatus == LoopbackShareStatus::Unreachable)
        return {.skipped = true, .reason = L"the loopback administrative share of the alternate volume is not reachable"};
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return {.reason = L"no-activation guard unavailable"};

    constexpr unsigned int kSourcePreferred      = 6u;
    constexpr unsigned int kDestinationPreferred = 2u;
    constexpr ThroughputCorpus kMoveCorpus{L"move", 1024ull * 1024ull, 4u, 256ull * 1024ull};
    const std::filesystem::path root             = state.tempRoot / L"same-provider-bridge";
    const std::filesystem::path sourceFolder     = root / L"corpus";
    const std::filesystem::path destinationLocal = alternate.value() / L"same-provider-bridge";
    const auto cleanup                           = wil::scope_exit([&]() noexcept
    {
        ClearFileOpsStoragePathConcurrencyOverridesForSelfTest();
        static_cast<void>(SelfTest::RemoveAll(root));
    });
    if (! RecreateEmptyDirectory(destinationLocal) || ! CreateThroughputCorpus(sourceFolder, kMoveCorpus))
        return {.reason = L"case roots unavailable"};
    AddFileOpsStoragePathConcurrencyOverrideForSelfTest(root.native(), kSourcePreferred, FILESYSTEM_STORAGE_NVME);
    AddFileOpsStoragePathConcurrencyOverrideForSelfTest(
        LoopbackShareAlias(alternate.value()).native(), kDestinationPreferred, FILESYSTEM_STORAGE_NETWORK_SHARE);

    ThroughputCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    const auto admitted                                       = StartFileOperationAndGetId(operations,
                                                                                           FILESYSTEM_MOVE,
                                                                                           FolderWindow::Pane::Left,
                                                                                           FolderWindow::Pane::Right,
                                                                                           state.fsLocal,
                                                                                           {sourceFolder},
                                                                                           LoopbackShareAlias(destinationLocal),
                                                                                           FILESYSTEM_FLAG_RECURSIVE,
                                                                                           false,
                                                                                           0u,
                                                                                           FolderWindow::FileOperationState::ExecutionMode::PerItem);
    if (! admitted.has_value())
        return {.reason = L"the Move was not admitted"};
    const uint64_t taskId = admitted.value();
    scope.tasks.push_back(taskId);
    const auto done = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return state.completedTasks.contains(taskId) && operations->FindTask(taskId) == nullptr;
    }, {.timeout = SelfTest::Scale(std::chrono::duration_cast<std::chrono::milliseconds>(kThroughputCopyTimeout)), .operationName = L"Managed Move"});
    if (! done.conditionMet)
        return {.reason = done.timeoutDiagnostic};
    const CompletedTaskInfo& info = state.completedTasks.at(taskId);
    const bool moved              = info.hr == S_OK && ThroughputCopyComplete(destinationLocal / L"corpus", kMoveCorpus) &&
                                    GetFileAttributesW(sourceFolder.c_str()) == INVALID_FILE_ATTRIBUTES;
    const bool managed            = ! info.sourceItemResults.empty() && info.sourceItemResults.front().has_value() &&
                                    info.sourceItemResults.front().value().strategy == FileOperations::OperationStrategy::Managed;
    AppendLog(std::format(L"Same-provider bridge storage: hr=0x{:08X} moved={} managed={} budget={} destinationKind={} stallSide={}",
                          static_cast<unsigned long>(info.hr),
                          moved,
                          managed,
                          info.configuredMaxConcurrency,
                          info.autoConcurrencyDestinationStorageKind,
                          static_cast<unsigned int>(info.progressStallSide)));
    if (! moved)
        return {.reason = std::format(L"the Move did not move the corpus (hr=0x{:08X})", static_cast<unsigned long>(info.hr))};
    if (! managed)
        return {.reason = L"the Move did not run as a Managed Move (host bridge): the fixture no longer exercises a same-provider bridge route"};
    if (info.configuredMaxConcurrency != std::min(kSourcePreferred, kDestinationPreferred) ||
        info.autoConcurrencyDestinationStorageKind != FILESYSTEM_STORAGE_NETWORK_SHARE)
        return {.reason = std::format(L"the Managed Move ran with budget {} and destination storage kind {} (expected {} and {})",
                                      info.configuredMaxConcurrency,
                                      info.autoConcurrencyDestinationStorageKind,
                                      std::min(kSourcePreferred, kDestinationPreferred),
                                      static_cast<uint32_t>(FILESYSTEM_STORAGE_NETWORK_SHARE))};
    if (info.progressStallSide != StallSide::Destination)
        return {.reason = std::format(L"a stall of the Managed Move would be named after side {} (expected the destination, {})",
                                      static_cast<unsigned int>(info.progressStallSide),
                                      static_cast<unsigned int>(StallSide::Destination))};
    return {.passed = true};
}

// R10.5: a Move of a folder of files from the local root to the share (the provider's managed route:
// copy, publish, then remove the exact source), cancelled inside each phase its reports announce, onto
// a fresh destination (direct-final route) and, on the loopback alias, onto existing files (staged
// overwrite). Whatever moment the cancel lands, no stage is left behind and each file is either moved
// or untouched. A planted stage on the share is then removable through the artifact touch guard
// (loopback alias only: on an identity-less NAS a leftover goes to the by-name consent of I35 R7.3,
// which the IdentityLess family covers). The folder holds two nested folders, so a cancel can also land
// inside a folder the Move created; the NAS takes them because no walker holds a created folder open
// (R0fSmb_ManagedMoveTreeToShare moves such a tree to the end).
[[nodiscard]] ThroughputCaseResult RunFileOpsSmbManagedMoveLeavesNoStageChecks(SelfTestState& state) noexcept
{
    if (! RunFileOpsBridgeStageDebtSelfTestForSelfTest())
        return {.reason = L"a sibling completion cleared unresolved bridge stage debt"};
    Debug::Perf::Emit(L"FileOps.SelfTest.StageDebtAggregation", L"counterfactual=Published;candidate=Unknown;strongerDebtPreserved=true", 0u, 12u, 2u, S_OK);
    using Task       = FolderWindow::FileOperationState::Task;
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return {.reason = L"owner, Local provider or initial worker quiet point unavailable"};
    const std::filesystem::path localRoot = state.tempRoot / L"managed-move";
    const ThroughputShareTarget target    = ResolveThroughputShareTarget(state, L"managed-move", localRoot / L"share");
    if (target.unavailable.has_value())
        return *target.unavailable;
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return {.reason = L"no-activation guard unavailable"};
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        SetFileOpsCancelAtProgressPhaseForSelfTest(std::nullopt, 0u);
        HostClearTestPromptResultOverride();
        static_cast<void>(SelfTest::RemoveAll(target.root));
        static_cast<void>(SelfTest::RemoveAll(localRoot));
    });
    if (! RecreateEmptyDirectory(localRoot / L"share") || ! RecreateEmptyDirectory(target.root))
        return {.reason = std::format(L"the {} target '{}' is unavailable", target.name, target.root.native())};

    // This synchronous provider fixture cancels between publication preparation and rename;
    // elapsed-time worker cancellation cannot deterministically select that exact boundary.
    class PublicationCancelControl final : public IFileSystemOperationControl
    {
    public:
        HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void*) noexcept override
        {
            if (abort == nullptr)
                return E_POINTER;
            *abort = ++checks >= 2u ? TRUE : FALSE;
            return S_OK;
        }
        HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void*) noexcept override
        {
            if (mode == nullptr)
                return E_POINTER;
            *mode = FILESYSTEM_DISCOVERY_AHEAD;
            return S_OK;
        }
        HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress*, void*) noexcept override
        {
            return S_OK;
        }
        unsigned int checks = 0u;
    };
    wil::com_ptr<IFileSystemObjectBinding> binding;
    if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(binding.put()))) || ! binding)
        return {.reason = L"the Local object-binding publication fixture is unavailable"};
    for (const bool directory : {false, true})
    {
        const std::filesystem::path stagePath = target.root / (directory ? L"cancel-before-rename-dir-stage" : L"cancel-before-rename-file-stage");
        const std::filesystem::path finalPath = target.root / (directory ? L"cancel-before-rename-dir-final" : L"cancel-before-rename-file-final");
        FileSystemOptions publicationOptions{};
        publicationOptions.sizeBytes = sizeof(publicationOptions);
        wil::com_ptr<IFileSystemBoundObject> stage;
        wil::com_ptr<IFileWriter> writer;
        const HRESULT createHr = directory ? binding->CreateExclusiveDirectory(stagePath.c_str(), &publicationOptions, stage.put())
                                           : binding->CreateExclusiveWriter(stagePath.c_str(), &publicationOptions, writer.put(), stage.put());
        if (FAILED(createHr) || ! stage || (! directory && (! writer || FAILED(writer->Commit()))))
            return {.reason = L"the pre-rename cancellation stage could not be prepared"};
        writer.reset();
        PublicationCancelControl control;
        publicationOptions.operationControl = &control;
        FileSystemConditionalMutationResult result{};
        result.sizeBytes = sizeof(result);
        wil::com_ptr<IFileSystemBoundObject> published;
        const HRESULT publishHr = stage->PublishAs(finalPath.c_str(), nullptr, FILESYSTEM_FLAG_NONE, &publicationOptions, &result, published.put());
        if (publishHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || control.checks != 2u || result.outcomeKnown == FALSE || result.mutationCommitted != FALSE ||
            result.originalStillPresent == FALSE || published || GetFileAttributesW(stagePath.c_str()) == INVALID_FILE_ATTRIBUTES ||
            GetFileAttributesW(finalPath.c_str()) != INVALID_FILE_ATTRIBUTES)
            return {.reason = std::format(L"pre-rename cancellation must be a known non-commit (directory={} checks={} hr=0x{:08X})",
                                          directory,
                                          control.checks,
                                          static_cast<unsigned long>(publishHr))};
        const FileSystemOptions cleanupOptions = MakeOwnedStageCleanupOptions(&publicationOptions);
        result                                 = {};
        result.sizeBytes                       = sizeof(result);
        const HRESULT abortHr                  = stage->AbortOwnedObject(&cleanupOptions, &result);
        if (FAILED(abortHr) || result.outcomeKnown == FALSE || result.mutationCommitted == FALSE || result.originalStillPresent != FALSE ||
            GetFileAttributesW(stagePath.c_str()) != INVALID_FILE_ATTRIBUTES)
            return {.reason = L"the known unpublished stage must remain exactly cleanable after cancellation"};
        Debug::Perf::Emit(L"FileOps.SelfTest.PublicationCancelCheckpoint", directory ? L"directory" : L"file", 0u, control.checks, 0u, S_OK);
    }

    // The opposite boundary: a real rename committed, but its reply cannot prove that. Force
    // retained-handle identity even on loopback; an absent stage must never manufacture success.
    constexpr wchar_t kIdentityPath[]     = L"REDSALAMANDER_FILEOPS_STAGE_IDENTITY_UNSUPPORTED_PATH";
    constexpr wchar_t kIdentityFire[]     = L"REDSALAMANDER_FILEOPS_STAGE_IDENTITY_UNSUPPORTED_FIRED";
    constexpr wchar_t kRenamePath[]       = L"REDSALAMANDER_FILEOPS_RENAME_COMMITTED_FAILURE_PATH";
    constexpr wchar_t kRenameFire[]       = L"REDSALAMANDER_FILEOPS_RENAME_COMMITTED_FAILURE_FIRED";
    const auto clearPublicationInjections = wil::scope_exit([&]() noexcept
    {
        for (const wchar_t* name : {kIdentityPath, kIdentityFire, kRenamePath, kRenameFire})
            static_cast<void>(SetEnvironmentVariableW(name, nullptr));
    });
    for (const bool directory : {false, true})
    {
        const std::filesystem::path stagePath = target.root / (directory ? L"unknown-dir-stage" : L"unknown-file-stage");
        const std::filesystem::path finalPath = target.root / (directory ? L"unknown-dir-final" : L"unknown-file-final");
        for (const wchar_t* name : {kIdentityFire, kRenameFire})
            static_cast<void>(SetEnvironmentVariableW(name, nullptr));
        if (SetEnvironmentVariableW(kIdentityPath, stagePath.c_str()) == FALSE || SetEnvironmentVariableW(kRenamePath, finalPath.c_str()) == FALSE)
            return {.reason = L"the uncertain-publication controls could not be armed"};
        FileSystemOptions options{};
        options.sizeBytes = sizeof(options);
        wil::com_ptr<IFileSystemBoundObject> stage;
        wil::com_ptr<IFileWriter> writer;
        const HRESULT createHr = directory ? binding->CreateExclusiveDirectory(stagePath.c_str(), &options, stage.put())
                                           : binding->CreateExclusiveWriter(stagePath.c_str(), &options, writer.put(), stage.put());
        if (FAILED(createHr) || ! stage || (! directory && (! writer || FAILED(writer->Commit()))))
            return {.reason = L"the uncertain-publication stage could not be prepared"};
        writer.reset();
        FileSystemConditionalMutationResult result{};
        result.sizeBytes = sizeof(result);
        wil::com_ptr<IFileSystemBoundObject> published;
        const HRESULT publishHr = stage->PublishAs(finalPath.c_str(), nullptr, FILESYSTEM_FLAG_NONE, &options, &result, published.put());
        if (publishHr != HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) || result.outcomeKnown != FALSE || result.mutationCommitted != FALSE || published ||
            GetEnvVarTrimmed(kIdentityFire) != L"1" || GetEnvVarTrimmed(kRenameFire) != L"1" ||
            GetFileAttributesW(stagePath.c_str()) != INVALID_FILE_ATTRIBUTES || GetFileAttributesW(finalPath.c_str()) == INVALID_FILE_ATTRIBUTES)
            return {.reason = std::format(L"an attempted uncertain rename must preserve Unknown and its final object (directory={} hr=0x{:08X})",
                                          directory,
                                          static_cast<unsigned long>(publishHr))};
        Debug::Perf::Emit(L"FileOps.SelfTest.PublicationAttemptedUnknown", directory ? L"directory" : L"file", 0u, 1u, 0u, S_OK);
    }

    struct CancelPoint final
    {
        ThroughputPhase phase   = ThroughputPhase::Unknown;
        unsigned int occurrence = 1u;
        std::wstring_view name;
        bool mustFire = true;
    };
    // The last chunk of a file is its first Finalizing report; the later ones are the boundaries
    // after the commit flush and the finalize metadata (or the next file's last chunk).
    constexpr std::array<CancelPoint, 5> kCancelPoints{{
        {ThroughputPhase::Opening, 1u, L"opening", true},
        {ThroughputPhase::Transferring, 1u, L"transferring", true},
        {ThroughputPhase::Finalizing, 1u, L"finalizing-1", true},
        {ThroughputPhase::Finalizing, 2u, L"finalizing-2", false},
        {ThroughputPhase::Finalizing, 3u, L"finalizing-3", false},
    }};
    ThroughputCaseScope scope(state, operations);
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    const auto timeout                                        = SelfTest::Scale(std::chrono::duration_cast<std::chrono::milliseconds>(kThroughputCopyTimeout));
    for (size_t pointIndex = 0u; pointIndex < kCancelPoints.size(); ++pointIndex)
    {
        for (const bool overwrite : {false, true})
        {
            // Replacing an existing file binds that file's identity, which an identity-less NAS cannot
            // expose (R7.3 deletes by name, it never replaces by name): the staged-overwrite runs use the
            // loopback alias only.
            if (overwrite && target.nas)
                continue;
            const CancelPoint& point                     = kCancelPoints[pointIndex];
            const uint64_t seedBase                      = 1000u + pointIndex * 100u + (overwrite ? 50u : 0u);
            const std::wstring runName                   = std::format(L"{}-{}", point.name, overwrite ? L"overwrite" : L"fresh");
            const std::filesystem::path sourcePayload    = localRoot / L"source" / runName / L"payload";
            const std::filesystem::path destinationRoot  = target.root / runName;
            const std::filesystem::path destinationFiles = destinationRoot / L"payload";
            if (! WriteMoveFixture(sourcePayload, false, seedBase) || ! RecreateEmptyDirectory(destinationRoot) ||
                (overwrite && ! WriteMoveFixture(destinationFiles, true, seedBase)))
                return {.reason = std::format(L"{}: the fixture could not be created", runName)};

            SetFileOpsCancelAtProgressPhaseForSelfTest(point.phase, point.occurrence);
            const auto admitted = StartFileOperationAndGetId(
                operations,
                FILESYSTEM_MOVE,
                FolderWindow::Pane::Left,
                FolderWindow::Pane::Right,
                state.fsLocal,
                {sourcePayload},
                destinationRoot,
                static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                false,
                0u,
                FolderWindow::FileOperationState::ExecutionMode::PerItem);
            if (! admitted.has_value())
                return {.reason = std::format(L"{}: the Move was not admitted", runName)};
            const uint64_t taskId = admitted.value();
            scope.tasks.push_back(taskId);
            const auto done = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                // The colliding files of the overwrite run ask once; Overwrite applies to all of them.
                if (Task* task = operations->FindTask(taskId))
                {
                    if (const auto prompt = TryGetConflictPromptCopy(task); prompt.has_value())
                        task->DebugSubmitCurrentConflictDecision(
                            PromptHasAction(*prompt, Task::ConflictAction::Overwrite) ? Task::ConflictAction::Overwrite : Task::ConflictAction::Cancel, true);
                }
                return state.completedTasks.contains(taskId) && operations->FindTask(taskId) == nullptr;
            },
                {.timeout = timeout, .operationName = L"Managed Move"});
            const unsigned int fired = TakeFileOpsCancelAtProgressPhaseFiredForSelfTest();
            SetFileOpsCancelAtProgressPhaseForSelfTest(std::nullopt, 0u);
            if (! done.conditionMet)
                return {.reason = std::format(L"{}: {}", runName, done.timeoutDiagnostic)};
            const CompletedTaskInfo& info = state.completedTasks.at(taskId);
            if (point.mustFire && fired == 0u)
                return {.reason = std::format(L"{}: no progress report announced the phase, so the cancellation never fired (hr=0x{:08X})",
                                              runName,
                                              static_cast<unsigned long>(info.hr))};
            const std::vector<std::wstring> stages = FindStageNames(destinationRoot);
            const auto summary                     = FindIdentityLessSummary(operations, taskId);
            const auto* item =
                info.sourceItemResults.size() == 1u && info.sourceItemResults.front().has_value() ? &info.sourceItemResults.front().value() : nullptr;
            if (item == nullptr || ! summary.has_value() ||
                (fired != 0u && (info.hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || item->completion != FileOperations::ItemCompletion::Canceled ||
                                 item->sourceDisposition != FileOperations::SourceDisposition::Retained)))
                return {.reason = std::format(L"{}: cancellation lost its typed root result or retained source", runName)};
            const bool unknown = item->publication == FileOperations::PublicationState::Unknown;
            if (! stages.empty() && (! unknown || item->ownedStageDisposition != FileOperations::OwnedStageDisposition::Unknown ||
                                     summary->unknownPublicationItemCount != 1u || summary->unknownOwnedStageCount != 1u))
                return {.reason = std::format(L"{}: unexplained retained stage without completed-card Unknown debt: {} (publication={} stage={})",
                                              runName,
                                              stages.front(),
                                              static_cast<unsigned int>(item->publication),
                                              static_cast<unsigned int>(item->ownedStageDisposition))};
            for (const std::wstring& path : stages)
            {
                // Capture the actual object, rather than accepting a name (or listing failure) as
                // evidence. An identity-less NAS leftover must remain inspectable and guarded.
                FileOperationArtifacts::Candidate candidate;
                if (path.starts_with(L"<listing failed:") ||
                    FAILED(FileOperationArtifacts::CaptureProviderObjectCandidate(state.fsLocal.get(), path, L"builtin/file-system", L"", candidate)))
                    return {.reason = std::format(L"{}: the retained stage cannot be inspected: {}", runName, path)};
                const auto projection = FileOperationArtifacts::ProjectCandidate(candidate);
                if (projection.classification != FileOperationArtifacts::Classification::Possible || ! projection.capabilities.inspect)
                    return {.reason = std::format(L"{}: the retained stage lost its Possible-artifact presentation", runName)};
                if (candidate.identityProbe.outcome == FileOperationArtifacts::IdentityProbeOutcome::IdentityUnsupported)
                {
                    const std::array candidates{candidate};
                    FileOperationArtifacts::TouchGuardRequest request;
                    FileOperationArtifacts::TouchGuardReceipt receipt;
                    if (FileOperationArtifacts::BuildTouchGuardRequest(candidates, request) != S_OK || ! request.containsPossible ||
                        FileOperationArtifacts::AcceptTouchGuard(request, receipt) != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || ! receipt.items.empty())
                        return {.reason = std::format(L"{}: an unidentifiable stage acquired recovery authority", runName)};
                }
            }
            unsigned int movedCount = 0u;
            if (const auto defect = CheckMoveOutcome(sourcePayload, destinationFiles, overwrite, seedBase, movedCount); defect.has_value())
                return {.reason = std::format(L"{}: {} (hr=0x{:08X}, cancelled={})", runName, *defect, static_cast<unsigned long>(info.hr), fired)};
            Debug::Perf::Emit(L"FileOps.SelfTest.ManagedMoveCancellationTruth", runName, 0u, stages.size(), unknown ? 1u : 0u, S_OK);
            AppendLog(std::format(L"Managed Move cancel: target={} run={} cancelled={} hr=0x{:08X} moved={}/{} retainedStages={} unknownPublication={}",
                                  target.name,
                                  runName,
                                  fired,
                                  static_cast<unsigned long>(info.hr),
                                  movedCount,
                                  kMoveFixtureFiles.size(),
                                  stages.size(),
                                  unknown));
        }
    }

    if (target.nas)
    {
        AppendLog(L"Managed Move: the planted-stage delete is checked on the loopback alias only (on the NAS a leftover goes to the by-name "
                  L"consent of I35 R7.3, covered by the IdentityLess family).");
        return {.passed = true};
    }
    // A leftover stage is a Possible artifact: deleting it asks the touch guard once and removes it.
    const std::filesystem::path planted = target.root / L"payload.bin.rs_tmp_0123456789ABCDEF0123456789ABCDEF_0";
    if (! WriteThroughputCorpusFile(planted, kMoveOldBytes, 7u))
        return {.reason = L"the planted stage could not be created"};
    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    const auto admitted = StartFileOperationAndGetId(
        operations, FILESYSTEM_DELETE, FolderWindow::Pane::Right, std::nullopt, state.fsLocal, {planted}, {}, static_cast<FileSystemFlags>(0u), false);
    if (! admitted.has_value())
        return {.reason = L"the Delete of the planted stage was not admitted"};
    scope.tasks.push_back(admitted.value());
    const auto deleted = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return state.completedTasks.contains(admitted.value()) && operations->FindTask(admitted.value()) == nullptr;
    }, {.timeout = timeout, .operationName = L"Planted stage Delete"});
    HostClearTestPromptResultOverride();
    std::error_code error;
    if (! deleted.conditionMet || state.completedTasks.at(admitted.value()).hr != S_OK || std::filesystem::exists(planted, error) ||
        HostGetTestPromptRequestCount() == 0u)
        return {.reason = std::format(L"the planted stage was not removed through the touch guard (completed={} hr=0x{:08X} exists={} prompts={})",
                                      deleted.conditionMet,
                                      deleted.conditionMet ? static_cast<unsigned long>(state.completedTasks.at(admitted.value()).hr) : 0ul,
                                      std::filesystem::exists(planted, error),
                                      HostGetTestPromptRequestCount())};
    AppendLog(std::format(
        L"Managed Move: the planted stage on the {} target was removed through the touch guard (prompts={}).", target.name, HostGetTestPromptRequestCount()));
    return {.passed = true};
}

// R10.2 deterministic gates for the speed display (5 s window, named stalls, stall band, restart
// after a stall longer than the window, the 40 s rule).
[[nodiscard]] ThroughputCaseResult RunFileOpsRateWindowDisplayChecks()
{
    using Check = std::optional<std::wstring> (*)();
    constexpr std::array<std::pair<std::wstring_view, Check>, 9> kChecks{{
        {L"short stall", &CheckRateShortStall},
        {L"long stall restart", &CheckRateLongStallRestart},
        {L"burst sawtooth", &CheckRateBurstSawtooth},
        {L"blocked states", &CheckRateBlockedStates},
        {L"long silence and items", &CheckRateLongSilenceAndItems},
        {L"throttled chunks", &CheckRateThrottledChunks},
        {L"slow steady stream", &CheckRateSlowSteadyStream},
        {L"real silence after steady reports", &CheckRateRealSilenceAfterSteady},
        {L"stopping task", &CheckRateStoppingNoStall},
    }};
    // Every rule is checked, so one run shows each that fails.
    std::wstring failures;
    for (const auto& [name, check] : kChecks)
    {
        if (const std::optional<std::wstring> failure = check(); failure.has_value())
            failures += std::format(L"{}{}: {}", failures.empty() ? L"" : L"; ", name, *failure);
    }
    if (! failures.empty())
        return {.reason = failures};
    return {.passed = true};
}
