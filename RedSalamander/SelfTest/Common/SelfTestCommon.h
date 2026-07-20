#pragma once

// SelfTestCommon - debug-only self-test infrastructure shared by all suites.
//
// Artifacts are written to:
//   %LOCALAPPDATA%\RedSalamander\SelfTest\last_run\   (current run)
//   %LOCALAPPDATA%\RedSalamander\SelfTest\previous_run\  (previous run, kept for diffing)
//
// When running selftests from a developer checkout, the harness also attempts to archive the
// key artifacts into the repo under:
//   Specs\TestRuns\<ComputerHashName>\<Area>\yyyy-MM-dd_HHmmss\
// so runs can be compared over time without relying on external scripts.
//
// Everything in this header is compiled only when ENABLE_TESTS is defined. Release builds without
// that define produce empty stub implementations so that call-sites do not need conditional compilation.

#ifdef ENABLE_TESTS

#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#include <algorithm>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <cwctype>
#include <filesystem>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include <wil/com.h>
#include <wil/resource.h>

struct IFileSystem;

namespace SelfTest
{

enum class SelfTestSuite
{
    CompareDirectories,
    FileOperations,
    Commands,
};

struct SelfTestOptions
{
    // Abort the run immediately after the first case failure.
    bool failFast = false;
    // Multiply every timeout by this bounded finite factor (use > 1.0 on slow CI machines).
    double timeoutScale = 1.0;
    // Write a results.json file to the suite artifact directory on completion.
    bool writeJsonSummary = true;
    // Enumerate declared cases without executing their bodies.
    bool listCasesOnly = false;
    // When set, run the exact matching case name (case-insensitive), every case with that
    // case-insensitive prefix when the filter ends in '_', or an exact comma-separated case list.
    std::wstring caseFilter;
    // Run every matched case this many times in-process. One is the normal single-pass mode.
    uint32_t repeatCount = 1;
    // One-based repeat attempt for explicit case-order dispatchers.
    uint32_t repeatIndex = 1;
    // When set, suites that support explicit ordering shuffle matched cases with this seed.
    std::optional<uint64_t> shuffleSeed;
    // Test-only crash proof hook: raise an access violation when this exact case starts.
    std::wstring crashCaseName;
    // Test-only classifier proof hooks. These intentionally fail only in selected suite contexts.
    std::wstring flakyProofCaseName;
    std::wstring orderProofCaseName;
    // Explicit-order dispatchers run one exact case at a time; keep classifier proof hooks aware
    // that those exact-case calls still represent suite/shuffle context, not isolated retries.
    bool classifierProofSuiteContext = false;
    bool classifierProofShuffleContext = false;
    // Optional JSON5 perf budget file used by focused performance selftests.
    std::filesystem::path perfBudgetPath;
    // When true, missing or non-applicable perf budgets fail the focused perf case.
    bool requirePerfBudgets = false;
};

struct SelfTestCaseResult
{
    std::wstring name;
    enum class Status
    {
        passed,
        failed,
        skipped,
        crashed,
    } status;

    uint64_t durationMs = 0;
    std::wstring reason;
    uint32_t repeatIndex = 1;
};

struct SelfTestSuiteResult
{
    SelfTestSuite suite{};
    uint64_t durationMs = 0;
    int passed          = 0;
    int failed          = 0;
    int skipped         = 0;
    std::vector<SelfTestCaseResult> cases;
    std::wstring failureMessage;
};

void AppendSelfTestTrace(std::wstring_view msg) noexcept;
void AppendSuiteTrace(SelfTestSuite suite, std::wstring_view msg) noexcept;
void BeginInFlightSelfTestCase(SelfTestSuite suite, std::wstring_view name) noexcept;
void EndInFlightSelfTestCase(SelfTestSuite suite, std::wstring_view name) noexcept;
void TriggerSelfTestCaseCrashInjection(SelfTestSuite suite, std::wstring_view name) noexcept;

struct CaseState
{
    std::wstring failure;
    std::wstring skipped;

    bool Require(bool condition, std::wstring_view message) noexcept
    {
        if (condition)
        {
            return true;
        }

        if (failure.empty())
        {
            failure.assign(message);
        }
        return false;
    }

    bool Skip(std::wstring_view reason) noexcept
    {
        if (failure.empty() && skipped.empty())
        {
            skipped.assign(reason.empty() ? std::wstring_view(L"skipped") : reason);
        }
        return true;
    }
};

[[nodiscard]] inline bool SelfTestCaseNameEquals(std::wstring_view lhs, std::wstring_view rhs) noexcept
{
    if (lhs.size() != rhs.size())
    {
        return false;
    }

    for (size_t i = 0; i < lhs.size(); ++i)
    {
        if (std::towlower(lhs[i]) != std::towlower(rhs[i]))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] inline bool CaseFilterMatches(std::wstring_view filter, std::wstring_view name) noexcept
{
    if (filter.empty())
    {
        return true;
    }

    const auto trim = [](std::wstring_view value) noexcept
    {
        while (! value.empty() && std::iswspace(static_cast<wint_t>(value.front())) != 0)
        {
            value.remove_prefix(1u);
        }
        while (! value.empty() && std::iswspace(static_cast<wint_t>(value.back())) != 0)
        {
            value.remove_suffix(1u);
        }
        return value;
    };

    if (filter.find(L',') != std::wstring_view::npos)
    {
        size_t start = 0u;
        while (start <= filter.size())
        {
            const size_t comma           = filter.find(L',', start);
            const std::wstring_view part = trim(filter.substr(start, comma == std::wstring_view::npos ? std::wstring_view::npos : comma - start));
            if (! part.empty() && SelfTestCaseNameEquals(part, name))
            {
                return true;
            }
            if (comma == std::wstring_view::npos)
            {
                break;
            }
            start = comma + 1u;
        }
        return false;
    }

    if (SelfTestCaseNameEquals(filter, name))
    {
        return true;
    }

    if (filter.back() != L'_' || filter.size() > name.size())
    {
        return false;
    }

    for (size_t i = 0; i < filter.size(); ++i)
    {
        if (std::towlower(filter[i]) != std::towlower(name[i]))
        {
            return false;
        }
    }

    return true;
}

struct SelfTestCaseExecution
{
    std::wstring name;
    uint32_t repeatIndex = 1;
};

[[nodiscard]] inline bool ShouldUseExplicitCaseExecutionOrder(const SelfTestOptions& options) noexcept
{
    return ! options.listCasesOnly && (options.repeatCount > 1u || options.shuffleSeed.has_value());
}

std::vector<SelfTestCaseExecution> BuildSelfTestCaseExecutionOrder(const SelfTestOptions& options, std::span<const std::wstring> declaredCases);

void AppendCaseResult(SelfTestSuiteResult& suite, SelfTestCaseResult result) noexcept;
void AppendCaseResult(
    SelfTestSuiteResult& suite, std::wstring_view name, SelfTestCaseResult::Status status, std::wstring_view reason, uint64_t durationMs) noexcept;

[[nodiscard]] inline bool TryInjectSelfTestClassifierProofFailure(
    const SelfTestOptions& options, SelfTestSuiteResult& suite, std::wstring_view name, uint32_t repeatIndex, SelfTestCaseResult& result) noexcept
{
    if (options.listCasesOnly)
    {
        return false;
    }

    const bool isolatedCaseRerun = ! options.classifierProofSuiteContext && ! options.caseFilter.empty() && SelfTestCaseNameEquals(options.caseFilter, name);
    std::wstring_view reason;
    if (SelfTestCaseNameEquals(options.orderProofCaseName, name) && ! isolatedCaseRerun)
    {
        reason = L"injected order-dependent classifier proof failure";
    }
    else if (SelfTestCaseNameEquals(options.flakyProofCaseName, name) && ! isolatedCaseRerun && ! options.classifierProofShuffleContext &&
             ((repeatIndex - 1u) % 3u) == 0u)
    {
        reason = L"injected flaky classifier proof failure";
    }

    if (reason.empty())
    {
        return false;
    }

    AppendSuiteTrace(suite.suite, std::format(L"Case classifier proof injection: {} reason='{}'", name, reason));
    AppendSelfTestTrace(std::format(L"Case classifier proof injection: {} reason='{}'", name, reason));
    result.status = SelfTestCaseResult::Status::failed;
    result.reason = std::wstring(reason);
    AppendCaseResult(suite, std::move(result));
    return true;
}

template <typename Func>
void RunCaseAttempt(const SelfTestOptions& options, SelfTestSuiteResult& suite, std::wstring_view name, uint32_t repeatIndex, Func&& func) noexcept
{
    if (! CaseFilterMatches(options.caseFilter, name))
    {
        return;
    }

    SelfTestCaseResult result{};
    result.name        = std::wstring(name);
    result.repeatIndex = repeatIndex;

    if (options.listCasesOnly)
    {
        result.status = SelfTestCaseResult::Status::skipped;
        result.reason = L"listed only";
        AppendCaseResult(suite, std::move(result));
        return;
    }

    if (options.failFast && suite.failed != 0)
    {
        result.status = SelfTestCaseResult::Status::skipped;
        result.reason = L"not executed (fail-fast)";
        AppendCaseResult(suite, std::move(result));
        return;
    }

    std::wstring caseLine;
    caseLine.reserve(6 + name.size());
    caseLine.append(L"Case: ");
    caseLine.append(name);
    AppendSuiteTrace(suite.suite, caseLine);
    AppendSelfTestTrace(caseLine);

    const auto startedAt = std::chrono::steady_clock::now();
    BeginInFlightSelfTestCase(suite.suite, name);
    if (SelfTestCaseNameEquals(options.crashCaseName, name))
    {
        TriggerSelfTestCaseCrashInjection(suite.suite, name);
    }
    if (TryInjectSelfTestClassifierProofFailure(options, suite, name, repeatIndex, result))
    {
        EndInFlightSelfTestCase(suite.suite, name);
        return;
    }
    CaseState state{};
    const bool ok      = std::forward<Func>(func)(state);
    const auto endedAt = std::chrono::steady_clock::now();

    AppendSuiteTrace(suite.suite,
                     std::format(L"Case returned: {} ok={} failed={} skipped={}",
                                 name,
                                 ok ? L"yes" : L"no",
                                 state.failure.empty() ? L"no" : L"yes",
                                 state.skipped.empty() ? L"no" : L"yes"));
    if (! state.failure.empty())
    {
        AppendSuiteTrace(suite.suite, std::format(L"Case failure: {} reason='{}'", name, state.failure));
        AppendSelfTestTrace(std::format(L"Case failure: {} reason='{}'", name, state.failure));
    }
    else if (! state.skipped.empty())
    {
        AppendSuiteTrace(suite.suite, std::format(L"Case skipped: {} reason='{}'", name, state.skipped));
        AppendSelfTestTrace(std::format(L"Case skipped: {} reason='{}'", name, state.skipped));
    }

    result.durationMs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(endedAt - startedAt).count());

    if (! ok || ! state.failure.empty())
    {
        result.status = SelfTestCaseResult::Status::failed;
        result.reason = state.failure.empty() ? L"case returned false without recording a failure reason" : state.failure;
        AppendCaseResult(suite, std::move(result));
        EndInFlightSelfTestCase(suite.suite, name);
        return;
    }

    if (! state.skipped.empty())
    {
        result.status = SelfTestCaseResult::Status::skipped;
        result.reason = state.skipped;
        AppendCaseResult(suite, std::move(result));
        EndInFlightSelfTestCase(suite.suite, name);
        return;
    }

    result.status = SelfTestCaseResult::Status::passed;
    AppendCaseResult(suite, std::move(result));
    EndInFlightSelfTestCase(suite.suite, name);
}

template <typename Func> void RunCase(const SelfTestOptions& options, SelfTestSuiteResult& suite, std::wstring_view name, Func&& func) noexcept
{
    const uint32_t repeatCount = options.listCasesOnly ? 1u : std::max(1u, options.repeatCount);
    for (uint32_t repeatIndex = 1u; repeatIndex <= repeatCount; ++repeatIndex)
    {
        RunCaseAttempt(options, suite, name, options.repeatIndex + repeatIndex - 1u, func);
    }
}

struct SelfTestRunResult
{
    std::wstring startedUtcIso;
    uint64_t durationMs = 0;
    bool failFast       = false;
    double timeoutScale = 1.0;
    std::wstring caseFilter;
    uint32_t repeatCount = 1;
    std::optional<uint64_t> shuffleSeed;
    std::vector<SelfTestSuiteResult> suites;
};

SelfTestOptions& GetSelfTestOptions() noexcept;
std::wstring_view GetSelfTestBuildFlavor() noexcept;
std::wstring GetSelfTestMachineHash() noexcept;
const std::filesystem::path& SelfTestRoot() noexcept;
std::filesystem::path GetSuiteRoot(SelfTestSuite suite);
std::filesystem::path GetSuiteArtifactPath(SelfTestSuite suite, std::wstring_view filename);
std::filesystem::path GetPerfArtifactPath(std::wstring_view filename);
wil::com_ptr<IFileSystem> GetFileSystem(std::wstring_view pluginId) noexcept;

struct TestSandbox
{
    std::filesystem::path root;

    [[nodiscard]] bool IsValid() const noexcept
    {
        return ! root.empty();
    }
};

[[nodiscard]] TestSandbox AcquireTestSandbox(SelfTestSuite suite, std::wstring_view caseName) noexcept;
[[nodiscard]] TestSandbox AcquireTestSandboxOnVolume(SelfTestSuite suite, std::wstring_view caseName, const std::filesystem::path& volumeRoot) noexcept;

void RotateSelfTestRuns();
void InitSelfTestRun(const SelfTestOptions& options);

void SetRunStartedUtcIso(std::wstring_view startedUtcIso) noexcept;
std::wstring_view GetRunStartedUtcIso() noexcept;

bool EnsureDirectory(const std::filesystem::path& path) noexcept;
bool RemoveAll(const std::filesystem::path& path) noexcept;
[[nodiscard]] bool WriteBinaryFile(const std::filesystem::path& path, std::span<const std::byte> bytes) noexcept;
[[nodiscard]] bool WriteTextFile(const std::filesystem::path& path, std::wstring_view text);
[[nodiscard]] bool WriteTextFile(const std::filesystem::path& path, std::string_view text);

// Shared filesystem-test primitives. Repository discovery intentionally accepts any one of the
// stable checkout markers so normal clones, git worktrees (.git is a file), and source snapshots
// can all resolve the same root. REDSALAMANDER_REPO_ROOT takes precedence when it names a valid root.
[[nodiscard]] std::filesystem::path TryFindRepoRoot() noexcept;
[[nodiscard]] std::filesystem::path GetLocalAppDataPath() noexcept;
[[nodiscard]] HRESULT EnsureDirectoryExists(const std::filesystem::path& path) noexcept;
[[nodiscard]] HRESULT WriteUtf8File(const std::filesystem::path& path, std::string_view text) noexcept;
[[nodiscard]] std::string NarrowAscii(std::wstring_view text);
[[nodiscard]] uint64_t StableDeviceHash(std::wstring_view value) noexcept;
[[nodiscard]] std::optional<uint64_t> ExtractJsonUInt(std::string_view json, std::string_view key) noexcept;

// Load and pin the configured MTP plugin while resolving a debug self-test export. The caller may
// retain the module when the returned object is implemented by the plugin.
[[nodiscard]] HRESULT LoadMtpPluginSelfTestExport(std::string_view exportName,
                                                  wil::unique_hmodule& module,
                                                  FARPROC& exportAddress) noexcept;

template<typename Fn, typename... Args>
[[nodiscard]] HRESULT CallMtpPluginExport(std::string_view exportName, Args&&... args) noexcept
{
#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL module owners are intentionally non-copyable; this helper keeps the pin local.
    wil::unique_hmodule module;
#pragma warning(pop)
    FARPROC exportAddress = nullptr;
    const HRESULT loadHr  = LoadMtpPluginSelfTestExport(exportName, module, exportAddress);
    if (FAILED(loadHr))
    {
        return loadHr;
    }

#pragma warning(push)
#pragma warning(disable : 4191) // GetProcAddress returns FARPROC; the named self-test export fixes the typed ABI.
    const auto function = reinterpret_cast<Fn>(exportAddress);
#pragma warning(pop)
    if (! function)
    {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    return function(std::forward<Args>(args)...);
}

std::filesystem::path GetTempRoot(SelfTestSuite suite);
bool PathExists(const std::filesystem::path& p);
// Multiply baseMs by the current timeoutScale factor (see SelfTestOptions).
// Use this whenever waiting for asynchronous work in a test case.
uint64_t ScaleTimeout(uint64_t baseMs);

inline std::chrono::milliseconds Scale(std::chrono::milliseconds base) noexcept
{
    if (base.count() <= 0)
    {
        return std::chrono::milliseconds{0};
    }

    const uint64_t scaled = ScaleTimeout(static_cast<uint64_t>(base.count()));
    const uint64_t maxMs  = static_cast<uint64_t>(std::chrono::milliseconds::max().count());
    return std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>((scaled > maxMs) ? maxMs : scaled)};
}

void WriteSuiteJson(const SelfTestSuiteResult& result, const std::filesystem::path& path);
void WriteRunJson(const SelfTestRunResult& result, const std::filesystem::path& path);
void MarkInFlightSelfTestCaseCrashed(SelfTestRunResult& runResult, std::wstring_view reason) noexcept;

// Copies the meaningful artifacts from %LOCALAPPDATA%\RedSalamander\SelfTest\last_run\
// into the repo under Specs\TestRuns\<ComputerHashName>\<Area>\yyyy-MM-dd_HHmmss\.
// Fresh run/suite JSON is written directly from the current in-memory result when available.
// If the repo root cannot be found (e.g. installed build), this is a no-op.
void TryArchiveLastRunToRepo(std::wstring_view area, int exitCode, uint64_t durationMs, const SelfTestRunResult* runResult = nullptr) noexcept;

} // namespace SelfTest

#endif // ENABLE_TESTS
