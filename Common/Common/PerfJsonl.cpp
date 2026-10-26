// PerfJsonl.cpp -- JSONL performance sink implementation.
// Compiled into Common.dll; exported via COMMON_API so all modules
// (RedSalamander, plugins, etc.) share a single set of JSONL state.

#include "Common.Tests.PerfJsonl.h"
#include "HandleIo.h"
#include "Helpers.h"

#include <atomic>
#include <charconv>
#include <cwchar>

namespace Debug::detail
{

// ---- module-local global state (single instance in Common.dll) ----

static std::mutex g_perfJsonlMutex;
// Configuration, serialization and append share one lock. No emitter may reopen an old path after
// ConfigureJsonlOutput/ClearJsonlOutput has returned, and rows cannot interleave across threads.
// Cached append handle (guarded by g_perfJsonlMutex) so the sink does not open+close the file on
// every row -- the previous per-row CreateFileW/CloseHandle was ~50-150us each and dominated any
// high-frequency metric. Reopened lazily when the path changes or the handle is not yet open; the OS
// closes it on process exit. In-process readers (the selftest presence scan) still see writes because
// WriteFile publishes to the shared OS file cache (no FlushFileBuffers needed for read visibility).
static wil::unique_hfile g_perfJsonlFile;
static std::filesystem::path g_perfJsonlPath;
static std::wstring g_perfJsonlScenario;
static std::wstring g_perfJsonlBuild;
static std::wstring g_perfJsonlBranch;
static std::wstring g_perfJsonlCommit;
static std::wstring g_perfJsonlMachineHash;
static std::wstring g_perfJsonlRunId;
static std::string g_perfJsonlMetadata;
static std::string g_perfJsonlLine;
constexpr size_t kMaxRetainedJsonlLineBytes = 64u * 1024u;
#ifdef ENABLE_TESTS
static thread_local uint64_t g_selfTestJsonlWriteCalls     = 0u;
static thread_local uint64_t g_selfTestJsonlWriteElapsedUs = 0u;
static std::atomic<uint64_t> g_selfTestJsonlMetadataPreparations{0u};
static std::atomic<size_t> g_selfTestJsonlRetainedLineCapacity{0u};
static_assert(std::atomic<uint64_t>::is_always_lock_free && std::atomic<size_t>::is_always_lock_free);
#endif
// The JSONL sink is configured exactly once from the environment (tests set it before launch). Resolving
// it once and caching the result in an atomic makes the overwhelmingly common no-sink production path a
// single relaxed load -- no global mutex, no GetEnvironmentVariableW syscall, and no string copies per
// metric emit (this sink is hit from the UI layout/render path and icon worker threads).
static bool g_perfJsonlEnvResolved = false;      // guarded by g_perfJsonlMutex; env queried at most once
static std::atomic<int> g_perfJsonlSinkState{0}; // 0 = unresolved, 1 = active sink, 2 = no sink (free fast-path)

// ---- environment variable constants ----

static constexpr wchar_t kPerfJsonlPathEnv[]        = L"REDSALAMANDER_PERF_JSONL_PATH";
static constexpr wchar_t kPerfJsonlScenarioEnv[]    = L"REDSALAMANDER_PERF_JSONL_SCENARIO";
static constexpr wchar_t kPerfJsonlBuildEnv[]       = L"REDSALAMANDER_PERF_JSONL_BUILD";
static constexpr wchar_t kPerfJsonlBranchEnv[]      = L"REDSALAMANDER_PERF_JSONL_BRANCH";
static constexpr wchar_t kPerfJsonlCommitEnv[]      = L"REDSALAMANDER_PERF_JSONL_COMMIT";
static constexpr wchar_t kPerfJsonlMachineHashEnv[] = L"REDSALAMANDER_PERF_JSONL_MACHINE_HASH";
static constexpr wchar_t kPerfJsonlRunIdEnv[]       = L"REDSALAMANDER_PERF_JSONL_RUN_ID";

// ---- internal helpers (anonymous namespace, not exported) ----

namespace
{

void AppendJsonEscapedUtf8(std::string& out, std::string_view utf8) noexcept
{
    const char* const data = utf8.data();
    size_t runStart        = 0u;
    for (size_t index = 0u; index < utf8.size(); ++index)
    {
        const unsigned char ch = static_cast<unsigned char>(data[index]);
        if (ch >= 0x20u && ch != '\\' && ch != '"')
        {
            continue;
        }
        out.append(data + runStart, index - runStart);
        switch (ch)
        {
            case '\\': out.append("\\\\"); break;
            case '"': out.append("\\\""); break;
            case '\b': out.append("\\b"); break;
            case '\f': out.append("\\f"); break;
            case '\n': out.append("\\n"); break;
            case '\r': out.append("\\r"); break;
            case '\t': out.append("\\t"); break;
            default:
                constexpr char kHex[] = "0123456789ABCDEF";
                const char escaped[]  = {'\\', 'u', '0', '0', kHex[ch >> 4u], kHex[ch & 0x0Fu]};
                out.append(escaped, sizeof(escaped));
                break;
        }
        runStart = index + 1u;
    }
    out.append(data + runStart, utf8.size() - runStart);
}

void AppendJsonEscaped(std::string& out, std::wstring_view text) noexcept
{
    // Metric names and short details are usually ASCII. Prove that bounded subset
    // before appending anything, avoiding a checked string proxy and two conversion
    // calls per field. Unicode and larger fields still use the canonical strict helper.
    std::array<char, 256> ascii;
    if (text.size() <= ascii.size())
    {
        const wchar_t* const source = text.data();
        char* const bytes           = ascii.data();
        size_t index                = 0u;
        while (index < text.size() && static_cast<uint32_t>(source[index]) <= 0x7Fu)
        {
            bytes[index] = static_cast<char>(source[index]);
            ++index;
        }
        if (index == text.size())
        {
            AppendJsonEscapedUtf8(out, std::string_view(bytes, index));
            return;
        }
    }
    const auto converted = Common::Strings::TryUtf8FromUtf16Strict(text);
    if (converted.has_value())
    {
        AppendJsonEscapedUtf8(out, converted.value());
    }
}

void AppendJsonStringField(std::string& out, std::string_view name, std::wstring_view value, bool& firstField) noexcept
{
    if (! firstField)
    {
        out.push_back(',');
    }
    firstField = false;
    out.push_back('"');
    out.append(name);
    out.append("\":\"");
    AppendJsonEscaped(out, value);
    out.push_back('"');
}

void AppendJsonNumericField(std::string& out, std::string_view name, uint64_t value, bool& firstField) noexcept
{
    if (! firstField)
    {
        out.push_back(',');
    }
    firstField = false;
    out.push_back('"');
    out.append(name);
    out.append("\":");
    char digits[20]; // Maximum decimal uint64_t; to_chars adds no terminator or allocation.
    const auto converted = std::to_chars(digits, digits + sizeof(digits), value);
    out.append(digits, static_cast<size_t>(converted.ptr - digits));
}

void AppendPaddedDecimal(std::string& out, WORD value, size_t minimumWidth) noexcept
{
    char digits[5]; // SYSTEMTIME members are WORDs; retain years wider than four digits.
    const auto converted = std::to_chars(digits, digits + sizeof(digits), value);
    const size_t count   = static_cast<size_t>(converted.ptr - digits);
    if (count < minimumWidth)
    {
        out.append(minimumWidth - count, '0');
    }
    out.append(digits, count);
}

void AppendCurrentUtcIsoTimestamp(std::string& out) noexcept
{
    SYSTEMTIME t{};
    GetSystemTime(&t);
    // This fixed ASCII wire format needs neither wide formatting nor UTF-8 conversion.
    AppendPaddedDecimal(out, t.wYear, 4u);
    out.push_back('-');
    AppendPaddedDecimal(out, t.wMonth, 2u);
    out.push_back('-');
    AppendPaddedDecimal(out, t.wDay, 2u);
    out.push_back('T');
    AppendPaddedDecimal(out, t.wHour, 2u);
    out.push_back(':');
    AppendPaddedDecimal(out, t.wMinute, 2u);
    out.push_back(':');
    AppendPaddedDecimal(out, t.wSecond, 2u);
    out.push_back('.');
    AppendPaddedDecimal(out, t.wMilliseconds, 3u);
    out.push_back('Z');
}

// Caller holds g_perfJsonlMutex. Metadata changes only at configuration boundaries;
// do not copy, transcode and escape the same six fields for every visible item.
void PrepareJsonlMetadataLocked() noexcept
{
#ifdef ENABLE_TESTS
    g_selfTestJsonlMetadataPreparations.fetch_add(1u, std::memory_order_relaxed);
#endif
    std::string metadata;
    bool firstField = false;
    AppendJsonStringField(metadata, "scenario", g_perfJsonlScenario, firstField);
    AppendJsonStringField(metadata, "build", g_perfJsonlBuild, firstField);
    AppendJsonStringField(metadata, "branch", g_perfJsonlBranch, firstField);
    AppendJsonStringField(metadata, "commit", g_perfJsonlCommit, firstField);
    AppendJsonStringField(metadata, "machineHash", g_perfJsonlMachineHash, firstField);
    AppendJsonStringField(metadata, "runId", g_perfJsonlRunId, firstField);
    g_perfJsonlMetadata = std::move(metadata);
}

[[nodiscard]] std::wstring GetEnvironmentVariableValue(const wchar_t* name) noexcept
{
    const DWORD length = GetEnvironmentVariableW(name, nullptr, 0u);
    if (length == 0u)
    {
        return {};
    }

    std::wstring value(length, L'\0');
    const DWORD copied = GetEnvironmentVariableW(name, value.data(), length);
    if (copied == 0u)
    {
        return {};
    }

    value.resize(copied);
    return value;
}

// Caller must hold g_perfJsonlMutex.
void TryInitializeJsonlOutputFromEnvironmentLocked() noexcept
{
    if (g_perfJsonlEnvResolved)
    {
        return;
    }
    g_perfJsonlEnvResolved = true;

    const std::wstring path = GetEnvironmentVariableValue(kPerfJsonlPathEnv);
    if (path.empty())
    {
        return;
    }

    g_perfJsonlPath        = path;
    g_perfJsonlScenario    = GetEnvironmentVariableValue(kPerfJsonlScenarioEnv);
    g_perfJsonlBuild       = GetEnvironmentVariableValue(kPerfJsonlBuildEnv);
    g_perfJsonlBranch      = GetEnvironmentVariableValue(kPerfJsonlBranchEnv);
    g_perfJsonlCommit      = GetEnvironmentVariableValue(kPerfJsonlCommitEnv);
    g_perfJsonlMachineHash = GetEnvironmentVariableValue(kPerfJsonlMachineHashEnv);
    g_perfJsonlRunId       = GetEnvironmentVariableValue(kPerfJsonlRunIdEnv);
    PrepareJsonlMetadataLocked();
}

} // anonymous namespace

// ---- exported functions ----

COMMON_API void WritePerfJsonl(std::wstring_view metric, std::wstring_view detail, uint64_t durationUs, uint64_t value0, uint64_t value1, HRESULT hr) noexcept
{
    // Fast path: once resolved with no sink configured, every emit is a single atomic load and return.
    if (g_perfJsonlSinkState.load(std::memory_order_acquire) == 2)
    {
        return;
    }
#ifdef ENABLE_TESTS
    const auto writeStartedAt = std::chrono::steady_clock::now();
    const auto observeWrite   = wil::scope_exit([&]
    {
        ++g_selfTestJsonlWriteCalls;
        g_selfTestJsonlWriteElapsedUs += Debug::Perf::ElapsedUs(writeStartedAt);
    });
#endif

    // Best-effort perf sink. Bad allocation should remain fatal; other runtime failures should not
    // interfere with normal execution or the ETW path.
    try
    {
        std::scoped_lock lock(g_perfJsonlMutex);
        TryInitializeJsonlOutputFromEnvironmentLocked();
        g_perfJsonlSinkState.store(g_perfJsonlPath.empty() ? 2 : 1, std::memory_order_release);
        if (g_perfJsonlPath.empty())
        {
            return;
        }

        // Reopen the cached append handle only when the path changes (or it is not yet open), so the
        // common case is a single WriteFile with no per-row open/close syscalls.
        if (! g_perfJsonlFile)
        {
            if (g_perfJsonlPath.has_parent_path())
            {
                std::error_code ec;
                std::filesystem::create_directories(g_perfJsonlPath.parent_path(), ec);
            }
            // Include FILE_SHARE_DELETE: this append handle is cached for the process lifetime, and
            // without delete-sharing it would block external/next-run rotation from renaming or removing
            // the perf file -- and, more importantly, its parent run directory -- for as long as the
            // process lives (matching the append-log convention used elsewhere, e.g. DxUi.Menu.cpp).
            g_perfJsonlFile.reset(CreateFileW(g_perfJsonlPath.c_str(),
                                              FILE_APPEND_DATA,
                                              FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                              nullptr,
                                              OPEN_ALWAYS,
                                              FILE_ATTRIBUTE_NORMAL,
                                              nullptr));
        }

        if (! g_perfJsonlFile)
        {
            return;
        }

        // Reuse one bounded retained row under the same lock as WriteFile. An unusually
        // large caller record is written intact, then its oversized capacity is released.
        std::string& line = g_perfJsonlLine;
        line.clear();
        line.reserve(512u);
        const auto releaseOversizedLine = wil::scope_exit([&]
        {
            if (line.capacity() > kMaxRetainedJsonlLineBytes)
            {
                std::string{}.swap(line);
            }
#ifdef ENABLE_TESTS
            g_selfTestJsonlRetainedLineCapacity.store(line.capacity(), std::memory_order_relaxed);
#endif
        });
        line.append("{\"timestamp\":\"");
        AppendCurrentUtcIsoTimestamp(line);
        line.push_back('"');
        bool firstField = false;
        AppendJsonNumericField(line, "process", static_cast<uint64_t>(GetCurrentProcessId()), firstField);
        AppendJsonNumericField(line, "thread", static_cast<uint64_t>(GetCurrentThreadId()), firstField);
        AppendJsonStringField(line, "metric", metric, firstField);
        AppendJsonStringField(line, "detail", detail, firstField);
        AppendJsonNumericField(line, "durationUs", durationUs, firstField);
        AppendJsonNumericField(line, "value0", value0, firstField);
        AppendJsonNumericField(line, "value1", value1, firstField);
        AppendJsonNumericField(line, "hr", static_cast<uint32_t>(hr), firstField);
        line.append(g_perfJsonlMetadata);

        if (durationUs != 0)
        {
            AppendJsonNumericField(line, "value", durationUs, firstField);
            AppendJsonStringField(line, "unit", L"us", firstField);
        }
        else
        {
            AppendJsonNumericField(line, "value", value0, firstField);
            AppendJsonStringField(line, "unit", L"count", firstField);
        }

        line.append("}\n");

        // Reuse the shared partial-write/large-buffer loop. The ordinary append still
        // needs one WriteFile; all chunks stay under the same row/configuration lock.
        static_cast<void>(Common::HandleIo::WriteAll(g_perfJsonlFile.get(), line.data(), line.size()));
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        return;
    }
}

COMMON_API bool HasPerfJsonlOutput() noexcept
{
    const int state = g_perfJsonlSinkState.load(std::memory_order_acquire);
    if (state != 0)
    {
        return state == 1;
    }
    std::scoped_lock lock(g_perfJsonlMutex);
    TryInitializeJsonlOutputFromEnvironmentLocked();
    g_perfJsonlSinkState.store(g_perfJsonlPath.empty() ? 2 : 1, std::memory_order_release);
    return ! g_perfJsonlPath.empty();
}

} // namespace Debug::detail

namespace Debug::Perf
{
#ifdef ENABLE_TESTS
COMMON_API JsonlDebugSnapshot GetJsonlDebugSnapshotForSelfTest() noexcept
{
    // Observation cannot wait for a writer before the request timer starts.
    // Shared resource values are published under the writer lock, then read atomically.
    return {.threadWriteCalls     = detail::g_selfTestJsonlWriteCalls,
            .threadWriteElapsedUs = detail::g_selfTestJsonlWriteElapsedUs,
            .metadataPreparations = detail::g_selfTestJsonlMetadataPreparations.load(std::memory_order_relaxed),
            .retainedLineCapacity = detail::g_selfTestJsonlRetainedLineCapacity.load(std::memory_order_relaxed)};
}
#endif

COMMON_API void ConfigureJsonlOutput(const std::filesystem::path& path,
                                     std::wstring_view scenario,
                                     std::wstring_view build,
                                     std::wstring_view branch,
                                     std::wstring_view commit,
                                     std::wstring_view machineHash,
                                     std::wstring_view runId) noexcept
{
    try
    {
        std::scoped_lock lock(detail::g_perfJsonlMutex);
        detail::g_perfJsonlPath        = path;
        detail::g_perfJsonlScenario    = std::wstring(scenario);
        detail::g_perfJsonlBuild       = std::wstring(build);
        detail::g_perfJsonlBranch      = std::wstring(branch);
        detail::g_perfJsonlCommit      = std::wstring(commit);
        detail::g_perfJsonlMachineHash = std::wstring(machineHash);
        detail::g_perfJsonlRunId       = std::wstring(runId);
        detail::g_perfJsonlEnvResolved = true;
        detail::PrepareJsonlMetadataLocked();
        // Runtime (re)configuration overrides whatever the cached sink-state latch resolved at startup. Without
        // this the WritePerfJsonl fast-path could stay stuck at "no sink" (2) -- latched by an earlier emit while
        // no path was configured -- and silently drop every metric even after a path is set (e.g. tests that
        // configure the JSONL sink mid-process via this entry point).
        detail::g_perfJsonlSinkState.store(detail::g_perfJsonlPath.empty() ? 2 : 1, std::memory_order_release);
        // Drop any cached append handle so the next emit reopens against the newly configured path/inode
        // even when the path string is unchanged (e.g. a test truncates+reconfigures the same file).
        detail::g_perfJsonlFile.reset();
        SetEnvironmentVariableW(detail::kPerfJsonlPathEnv, detail::g_perfJsonlPath.empty() ? nullptr : detail::g_perfJsonlPath.c_str());
        SetEnvironmentVariableW(detail::kPerfJsonlScenarioEnv, detail::g_perfJsonlScenario.empty() ? nullptr : detail::g_perfJsonlScenario.c_str());
        SetEnvironmentVariableW(detail::kPerfJsonlBuildEnv, detail::g_perfJsonlBuild.empty() ? nullptr : detail::g_perfJsonlBuild.c_str());
        SetEnvironmentVariableW(detail::kPerfJsonlBranchEnv, detail::g_perfJsonlBranch.empty() ? nullptr : detail::g_perfJsonlBranch.c_str());
        SetEnvironmentVariableW(detail::kPerfJsonlCommitEnv, detail::g_perfJsonlCommit.empty() ? nullptr : detail::g_perfJsonlCommit.c_str());
        SetEnvironmentVariableW(detail::kPerfJsonlMachineHashEnv, detail::g_perfJsonlMachineHash.empty() ? nullptr : detail::g_perfJsonlMachineHash.c_str());
        SetEnvironmentVariableW(detail::kPerfJsonlRunIdEnv, detail::g_perfJsonlRunId.empty() ? nullptr : detail::g_perfJsonlRunId.c_str());
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
    }
}

COMMON_API void ClearJsonlOutput() noexcept
{
    std::scoped_lock lock(detail::g_perfJsonlMutex);
    detail::g_perfJsonlPath.clear();
    detail::g_perfJsonlScenario.clear();
    detail::g_perfJsonlBuild.clear();
    detail::g_perfJsonlBranch.clear();
    detail::g_perfJsonlCommit.clear();
    detail::g_perfJsonlMachineHash.clear();
    detail::g_perfJsonlRunId.clear();
    std::string{}.swap(detail::g_perfJsonlMetadata);
    std::string{}.swap(detail::g_perfJsonlLine);
#ifdef ENABLE_TESTS
    detail::g_selfTestJsonlRetainedLineCapacity.store(detail::g_perfJsonlLine.capacity(), std::memory_order_relaxed);
#endif
    // No sink remains configured; latch the fast-path so subsequent emits return cheaply. A later
    // ConfigureJsonlOutput re-enables the sink by storing 1 again.
    detail::g_perfJsonlSinkState.store(2, std::memory_order_release);
    // Drop the cached append handle so a later reconfigure to the SAME path reopens against the current
    // file instead of appending to a stale (possibly deleted/recreated) inode, and so the open handle does
    // not keep the file locked across the clear.
    detail::g_perfJsonlFile.reset();
    SetEnvironmentVariableW(detail::kPerfJsonlPathEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlScenarioEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlBuildEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlBranchEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlCommitEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlMachineHashEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlRunIdEnv, nullptr);
}

} // namespace Debug::Perf
