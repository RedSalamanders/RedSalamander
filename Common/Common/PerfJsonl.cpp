// PerfJsonl.cpp -- JSONL performance sink implementation.
// Compiled into Common.dll; exported via COMMON_API so all modules
// (RedSalamander, plugins, etc.) share a single set of JSONL state.

#include "Helpers.h"

#include <cwchar>

namespace Debug::detail
{

// ---- module-local global state (single instance in Common.dll) ----

static std::mutex g_perfJsonlMutex;
static std::mutex g_perfJsonlWriteMutex;
static std::filesystem::path g_perfJsonlPath;
static std::wstring g_perfJsonlScenario;
static std::wstring g_perfJsonlBuild;
static std::wstring g_perfJsonlBranch;
static std::wstring g_perfJsonlCommit;
static std::wstring g_perfJsonlMachineHash;
static std::wstring g_perfJsonlRunId;

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

bool ConvertUtf8(std::wstring_view text, std::string& out) noexcept
{
    if (text.empty())
    {
        out.clear();
        return true;
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return false;
    }

    out.resize(static_cast<size_t>(required));
    const int written = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), out.data(), required, nullptr, nullptr);
    return written == required;
}

void AppendJsonEscaped(std::string& out, std::wstring_view text) noexcept
{
    std::string utf8;
    if (! ConvertUtf8(text, utf8))
    {
        return;
    }

    for (const char rawCh : utf8)
    {
        const unsigned char ch = static_cast<unsigned char>(rawCh);
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
                if (ch < 0x20)
                {
                    char buf[8];
                    const int n = std::snprintf(buf, sizeof(buf), "\\u%04X", static_cast<unsigned int>(ch));
                    if (n > 0)
                    {
                        out.append(buf, static_cast<size_t>(n));
                    }
                }
                else
                {
                    out.push_back(static_cast<char>(ch));
                }
                break;
        }
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
    out.append(std::to_string(value));
}

std::wstring CurrentUtcIsoTimestamp() noexcept
{
    SYSTEMTIME t{};
    GetSystemTime(&t);
    return std::format(L"{0:04}-{1:02}-{2:02}T{3:02}:{4:02}:{5:02}.{6:03}Z", t.wYear, t.wMonth, t.wDay, t.wHour, t.wMinute, t.wSecond, t.wMilliseconds);
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
    if (! g_perfJsonlPath.empty())
    {
        return;
    }

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
}

} // anonymous namespace

// ---- exported functions ----

COMMON_API void WritePerfJsonl(std::wstring_view metric, std::wstring_view detail, uint64_t durationUs, uint64_t value0, uint64_t value1, HRESULT hr) noexcept
{
    std::filesystem::path outputPath;
    std::wstring scenario;
    std::wstring build;
    std::wstring branch;
    std::wstring commit;
    std::wstring machineHash;
    std::wstring runId;
    {
        std::scoped_lock lock(g_perfJsonlMutex);
        TryInitializeJsonlOutputFromEnvironmentLocked();
        outputPath  = g_perfJsonlPath;
        scenario    = g_perfJsonlScenario;
        build       = g_perfJsonlBuild;
        branch      = g_perfJsonlBranch;
        commit      = g_perfJsonlCommit;
        machineHash = g_perfJsonlMachineHash;
        runId       = g_perfJsonlRunId;
    }

    if (outputPath.empty())
    {
        return;
    }

    // Best-effort perf sink. Bad allocation should remain fatal; other runtime failures should not
    // interfere with normal execution or the ETW path.
    try
    {
        std::scoped_lock writeLock(g_perfJsonlWriteMutex);

        if (outputPath.has_parent_path())
        {
            std::error_code ec;
            std::filesystem::create_directories(outputPath.parent_path(), ec);
        }

        std::string line;
        line.reserve(512);
        line.push_back('{');
        bool firstField = true;

        AppendJsonStringField(line, "timestamp", CurrentUtcIsoTimestamp(), firstField);
        AppendJsonNumericField(line, "process", static_cast<uint64_t>(GetCurrentProcessId()), firstField);
        AppendJsonNumericField(line, "thread", static_cast<uint64_t>(GetCurrentThreadId()), firstField);
        AppendJsonStringField(line, "metric", metric, firstField);
        AppendJsonStringField(line, "detail", detail, firstField);
        AppendJsonNumericField(line, "durationUs", durationUs, firstField);
        AppendJsonNumericField(line, "value0", value0, firstField);
        AppendJsonNumericField(line, "value1", value1, firstField);
        AppendJsonNumericField(line, "hr", static_cast<uint32_t>(hr), firstField);
        AppendJsonStringField(line, "scenario", scenario, firstField);
        AppendJsonStringField(line, "build", build, firstField);
        AppendJsonStringField(line, "branch", branch, firstField);
        AppendJsonStringField(line, "commit", commit, firstField);
        AppendJsonStringField(line, "machineHash", machineHash, firstField);
        AppendJsonStringField(line, "runId", runId, firstField);

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

        wil::unique_handle file(
            CreateFileW(outputPath.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! file)
        {
            return;
        }

        DWORD written = 0;
        static_cast<void>(WriteFile(file.get(), line.data(), static_cast<DWORD>(line.size()), &written, nullptr));
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
    std::scoped_lock lock(g_perfJsonlMutex);
    TryInitializeJsonlOutputFromEnvironmentLocked();
    return ! g_perfJsonlPath.empty();
}

} // namespace Debug::detail

namespace Debug::Perf
{

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
    SetEnvironmentVariableW(detail::kPerfJsonlPathEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlScenarioEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlBuildEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlBranchEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlCommitEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlMachineHashEnv, nullptr);
    SetEnvironmentVariableW(detail::kPerfJsonlRunIdEnv, nullptr);
}

} // namespace Debug::Perf
