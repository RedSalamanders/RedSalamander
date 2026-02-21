#ifdef _DEBUG

#include "SelfTestCommon.h"

#include <array>
#include <cstddef>
#include <filesystem>
#include <format>
#include <limits>
#include <span>
#include <string>

#include "FileSystemPluginManager.h"
#include <shlobj_core.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#include <wil/win32_helpers.h>

namespace
{
struct ForceWilTemplateInstantiations_SelfTestCommon
{
    wil::unique_handle handle;
    wil::unique_any<char*, decltype(&::free), ::free> mallocString;
};
} // namespace
#pragma warning(pop)

#include <yyjson.h>

namespace SelfTest
{
namespace
{
constexpr std::wstring_view kRootDirName{L"SelfTest"};
constexpr std::wstring_view kRedSalamanderDirName{L"RedSalamander"};
constexpr std::wstring_view kLastRunDirName{L"last_run"};
constexpr std::wstring_view kPreviousRunDirName{L"previous_run"};
constexpr std::wstring_view kCompareDirName{L"compare"};
constexpr std::wstring_view kFileOpsDirName{L"fileops"};
constexpr std::wstring_view kCommandsDirName{L"commands"};
constexpr std::wstring_view kTraceFileName{L"trace.txt"};
constexpr const char* kSuiteCompareName = "CompareDirectories";
constexpr const char* kSuiteFileOpsName = "FileOperations";
constexpr const char* kSuiteCommandsName = "Commands";

SelfTestOptions g_options{};
std::wstring g_runStartedUtcIso;

// Path helpers
// Resolves %LOCALAPPDATA% via SHGetKnownFolderPath first; falls back to
// GetEnvironmentVariableW in case the COM shell API is unavailable.
[[nodiscard]] std::filesystem::path GetLocalAppDataPath() noexcept
{
    wil::unique_cotaskmem_string localAppData;
    const HRESULT hr = SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, localAppData.put());
    if (SUCCEEDED(hr) && localAppData)
    {
        return std::filesystem::path(localAppData.get());
    }

    const DWORD required = GetEnvironmentVariableW(L"LOCALAPPDATA", nullptr, 0);
    if (required == 0)
    {
        return {};
    }

    std::wstring buffer(static_cast<size_t>(required), L'\0');
    const DWORD written = GetEnvironmentVariableW(L"LOCALAPPDATA", buffer.data(), required);
    if (written == 0 || written >= required)
    {
        return {};
    }

    buffer.resize(static_cast<size_t>(written));
    return std::filesystem::path(buffer);
}

[[nodiscard]] const char* SuiteName(SelfTestSuite suite) noexcept
{
    switch (suite)
    {
        case SelfTestSuite::CompareDirectories: return kSuiteCompareName;
        case SelfTestSuite::FileOperations: return kSuiteFileOpsName;
        case SelfTestSuite::Commands: return kSuiteCommandsName;
    }
    return "Unknown";
}

const char* CaseStatusName(SelfTestCaseResult::Status status) noexcept
{
    switch (status)
    {
        case SelfTestCaseResult::Status::passed: return "passed";
        case SelfTestCaseResult::Status::failed: return "failed";
        case SelfTestCaseResult::Status::skipped: return "skipped";
    }
    return "unknown";
}

// Trace / logging helpers (UTF-16 LE with BOM, one message per line)
void TruncateUtf16Log(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return;
    }

    wil::unique_handle file(CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return;
    }

    const wchar_t bom = 0xFEFF;
    DWORD written     = 0;
    static_cast<void>(WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr));
}

void AppendUtf16LogLine(const std::filesystem::path& path, std::wstring_view message) noexcept
{
    if (path.empty())
    {
        return;
    }

    wil::unique_handle file(CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return;
    }

    LARGE_INTEGER size{};
    if (! GetFileSizeEx(file.get(), &size))
    {
        return;
    }

    if (size.QuadPart == 0)
    {
        const wchar_t bom = 0xFEFF;
        DWORD written     = 0;
        static_cast<void>(WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr));
    }

    LARGE_INTEGER seek{};
    seek.QuadPart = 0;
    static_cast<void>(SetFilePointerEx(file.get(), seek, nullptr, FILE_END));

    if (! message.empty())
    {
        const size_t bytes = message.size() * sizeof(wchar_t);
        if (bytes <= static_cast<size_t>(std::numeric_limits<DWORD>::max()))
        {
            DWORD written = 0;
            static_cast<void>(WriteFile(file.get(), message.data(), static_cast<DWORD>(bytes), &written, nullptr));
        }
    }

    constexpr wchar_t newline[] = L"\r\n";
    DWORD written               = 0;
    static_cast<void>(WriteFile(file.get(), newline, static_cast<DWORD>(sizeof(newline) - sizeof(wchar_t)), &written, nullptr));
    static_cast<void>(FlushFileBuffers(file.get()));
}

// File I/O helpers
bool ConvertUtf8(const std::wstring_view text, std::string& out) noexcept
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
    if (written != required)
    {
        return false;
    }

    return true;
}

bool WriteJsonBlob(const std::filesystem::path& path, yyjson_mut_doc* doc) noexcept
{
    if (path.empty() || ! doc)
    {
        return false;
    }

    yyjson_mut_val* root = yyjson_mut_doc_get_root(doc);
    if (! root)
    {
        return false;
    }

    size_t jsonLen = 0;
    yyjson_write_err writeErr{};
    char* json = yyjson_mut_write_opts(doc, YYJSON_WRITE_PRETTY, nullptr, &jsonLen, &writeErr);
    if (! json)
    {
        return false;
    }

    auto freeJson = wil::unique_any<char*, decltype(&::free), &::free>(json);
    if (! freeJson)
    {
        return false;
    }

    return WriteBinaryFile(path, std::as_bytes(std::span<const char>(json, jsonLen)));
}

// JSON serialization helpers (yyjson mutable API, UTF-8 output)
void AddCaseJson(yyjson_mut_doc* doc, yyjson_mut_val* casesArray, const SelfTestCaseResult& testCase)
{
    std::string caseName;
    if (! ConvertUtf8(testCase.name, caseName))
    {
        return;
    }

    yyjson_mut_val* caseObj = yyjson_mut_obj(doc);
    if (! caseObj)
    {
        return;
    }

    yyjson_mut_obj_add_strncpy(doc, caseObj, "name", caseName.c_str(), caseName.size());
    yyjson_mut_obj_add_str(doc, caseObj, "status", CaseStatusName(testCase.status));
    yyjson_mut_obj_add_uint(doc, caseObj, "duration_ms", testCase.durationMs);

    if (! testCase.reason.empty())
    {
        std::string reason;
        if (ConvertUtf8(testCase.reason, reason))
        {
            yyjson_mut_obj_add_strncpy(doc, caseObj, "reason", reason.c_str(), reason.size());
        }
    }

    static_cast<void>(yyjson_mut_arr_add_val(casesArray, caseObj));
}

void AddSuiteJson(yyjson_mut_doc* doc, yyjson_mut_val* suitesArray, const SelfTestSuiteResult& result)
{
    yyjson_mut_val* suiteObj = yyjson_mut_obj(doc);
    if (! suiteObj)
    {
        return;
    }

    yyjson_mut_obj_add_str(doc, suiteObj, "suite", SuiteName(result.suite));

    std::string rootUtf8;
    if (ConvertUtf8(GetSuiteRoot(result.suite).wstring(), rootUtf8))
    {
        yyjson_mut_obj_add_strncpy(doc, suiteObj, "artifact_root", rootUtf8.c_str(), rootUtf8.size());
    }

    yyjson_mut_obj_add_uint(doc, suiteObj, "duration_ms", result.durationMs);
    yyjson_mut_obj_add_int(doc, suiteObj, "passed", result.passed);
    yyjson_mut_obj_add_int(doc, suiteObj, "failed", result.failed);
    yyjson_mut_obj_add_int(doc, suiteObj, "skipped", result.skipped);
    yyjson_mut_obj_add_bool(doc, suiteObj, "fail_fast", g_options.failFast);
    yyjson_mut_obj_add_real(doc, suiteObj, "timeout_scale", g_options.timeoutScale);

    if (! result.failureMessage.empty())
    {
        std::string failUtf8;
        if (ConvertUtf8(result.failureMessage, failUtf8))
        {
            yyjson_mut_obj_add_strncpy(doc, suiteObj, "failureMessage", failUtf8.c_str(), failUtf8.size());
        }
    }

    yyjson_mut_val* casesArray = yyjson_mut_arr(doc);
    if (! casesArray)
    {
        return;
    }

    for (const auto& item : result.cases)
    {
        AddCaseJson(doc, casesArray, item);
    }

    yyjson_mut_obj_add_val(doc, suiteObj, "cases", casesArray);
    static_cast<void>(yyjson_mut_arr_add_val(suitesArray, suiteObj));
}

} // namespace

SelfTestOptions& GetSelfTestOptions() noexcept
{
    return g_options;
}

const std::filesystem::path& SelfTestRoot() noexcept
{
    static const std::filesystem::path root = []
    {
        const std::filesystem::path base = GetLocalAppDataPath();
        if (base.empty())
        {
            return std::filesystem::path{};
        }

        return base / kRedSalamanderDirName / kRootDirName;
    }();
    return root;
}

std::filesystem::path GetSuiteRoot(SelfTestSuite suite)
{
    const std::filesystem::path root = SelfTestRoot();
    if (root.empty())
    {
        return {};
    }

    std::filesystem::path suiteDir;
    switch (suite)
    {
        case SelfTestSuite::CompareDirectories: suiteDir = kCompareDirName; break;
        case SelfTestSuite::FileOperations: suiteDir = kFileOpsDirName; break;
        case SelfTestSuite::Commands: suiteDir = kCommandsDirName; break;
        default: return {};
    }
    return root / kLastRunDirName / suiteDir;
}

std::filesystem::path GetSuiteArtifactPath(SelfTestSuite suite, std::wstring_view filename)
{
    const std::filesystem::path suiteRoot = GetSuiteRoot(suite);
    if (suiteRoot.empty() || filename.empty())
    {
        return {};
    }

    return suiteRoot / filename;
}

wil::com_ptr<IFileSystem> GetFileSystem(std::wstring_view pluginId) noexcept
{
    if (pluginId.empty())
    {
        return {};
    }

    if (pluginId.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    for (const FileSystemPluginManager::PluginEntry& entry : FileSystemPluginManager::GetInstance().GetPlugins())
    {
        if (! entry.fileSystem)
        {
            continue;
        }

        if (CompareStringOrdinal(entry.id.c_str(), -1, pluginId.data(), static_cast<int>(pluginId.size()), TRUE) == CSTR_EQUAL)
        {
            return entry.fileSystem;
        }
    }

    return {};
}

// Rotate artifacts: previous_run/ is discarded, last_run/ is renamed to previous_run/,
// and fresh empty directories are created under last_run/ ready for the new run.
void RotateSelfTestRuns()
{
    const std::filesystem::path root = SelfTestRoot();
    if (root.empty())
    {
        return;
    }

    const std::filesystem::path lastRun     = root / kLastRunDirName;
    const std::filesystem::path previousRun = root / kPreviousRunDirName;

    std::error_code ec;
    if (std::filesystem::exists(previousRun, ec))
    {
        std::filesystem::remove_all(previousRun, ec);
    }

    if (std::filesystem::exists(lastRun, ec))
    {
        std::filesystem::rename(lastRun, previousRun, ec);
        if (ec)
        {
            std::filesystem::remove_all(lastRun, ec);
        }
    }

    std::filesystem::create_directories(lastRun / kCompareDirName, ec);
    std::filesystem::create_directories(lastRun / kFileOpsDirName, ec);
    std::filesystem::create_directories(lastRun / kCommandsDirName, ec);

    TruncateUtf16Log(lastRun / kTraceFileName);
    TruncateUtf16Log(lastRun / kCompareDirName / kTraceFileName);
    TruncateUtf16Log(lastRun / kFileOpsDirName / kTraceFileName);
    TruncateUtf16Log(lastRun / kCommandsDirName / kTraceFileName);
}

void InitSelfTestRun(const SelfTestOptions& options)
{
    g_options      = options;
}

void AppendSelfTestTrace(std::wstring_view msg) noexcept
{
    AppendUtf16LogLine(SelfTestRoot() / kLastRunDirName / kTraceFileName, msg);
}

void AppendSuiteTrace(SelfTestSuite suite, std::wstring_view msg) noexcept
{
    AppendUtf16LogLine(GetSuiteArtifactPath(suite, kTraceFileName), msg);
}

void SetRunStartedUtcIso(std::wstring_view startedUtcIso) noexcept
{
    g_runStartedUtcIso.assign(startedUtcIso);
}

std::wstring_view GetRunStartedUtcIso() noexcept
{
    return g_runStartedUtcIso;
}

bool EnsureDirectory(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    std::error_code ec;
    std::filesystem::create_directories(path, ec);
    if (ec)
    {
        return false;
    }

    return PathExists(path);
}

bool WriteBinaryFile(const std::filesystem::path& path, std::span<const std::byte> bytes) noexcept
{
    if (path.empty())
    {
        return false;
    }

    if (path.has_parent_path())
    {
        EnsureDirectory(path.parent_path());
    }

    wil::unique_handle file(CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return false;
    }

    if (! bytes.empty())
    {
        constexpr size_t kChunkSize = 16ull * 1024ull * 1024ull; // 16 MiB
        size_t offset               = 0;
        while (offset < bytes.size())
        {
            const size_t remaining = bytes.size() - offset;
            const size_t chunkSize = (remaining > kChunkSize) ? kChunkSize : remaining;

            DWORD written = 0;
            if (! WriteFile(file.get(), bytes.data() + offset, static_cast<DWORD>(chunkSize), &written, nullptr))
            {
                return false;
            }

            if (written == 0)
            {
                return false;
            }

            offset += static_cast<size_t>(written);
        }
    }

    static_cast<void>(FlushFileBuffers(file.get()));
    return true;
}

bool WriteTextFile(const std::filesystem::path& path, std::wstring_view text)
{
    if (path.empty())
    {
        return false;
    }

    std::string utf8;
    if (! ConvertUtf8(text, utf8))
    {
        return false;
    }

    return WriteBinaryFile(path, std::as_bytes(std::span<const char>(utf8)));
}

bool WriteTextFile(const std::filesystem::path& path, std::string_view text)
{
    if (path.empty())
    {
        return false;
    }

    if (text.empty())
    {
        return WriteBinaryFile(path, {});
    }

    return WriteBinaryFile(path, std::as_bytes(std::span<const char>(text)));
}

std::filesystem::path GetTempRoot(SelfTestSuite suite)
{
    const std::filesystem::path suiteRoot = GetSuiteRoot(suite);
    if (suiteRoot.empty())
    {
        return {};
    }

    std::error_code ec;
    std::filesystem::create_directories(suiteRoot, ec);
    return suiteRoot;
}

bool PathExists(const std::filesystem::path& p)
{
    if (p.empty())
    {
        return false;
    }

    std::error_code ec;
    return std::filesystem::exists(p, ec) && ! ec;
}

uint64_t ScaleTimeout(uint64_t baseMs)
{
    const double scaled = static_cast<double>(baseMs) * GetSelfTestOptions().timeoutScale;
    if (scaled <= 0.0)
    {
        return 0;
    }

    if (scaled > static_cast<double>(std::numeric_limits<uint64_t>::max()))
    {
        return std::numeric_limits<uint64_t>::max();
    }

    return static_cast<uint64_t>(scaled);
}

void WriteSuiteJson(const SelfTestSuiteResult& result, const std::filesystem::path& path)
{
    if (! g_options.writeJsonSummary || path.empty())
    {
        return;
    }

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return;
    }

    auto freeDoc         = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });
    yyjson_mut_val* root = yyjson_mut_obj(doc);
    if (! root)
    {
        return;
    }

    yyjson_mut_doc_set_root(doc, root);

    if (! g_runStartedUtcIso.empty())
    {
        std::string started;
        if (ConvertUtf8(g_runStartedUtcIso, started))
        {
            yyjson_mut_obj_add_strncpy(doc, root, "run_started_utc", started.c_str(), started.size());
        }
    }

    yyjson_mut_obj_add_str(doc, root, "suite", SuiteName(result.suite));
    std::string pathUtf8;
    if (ConvertUtf8(GetSuiteRoot(result.suite).wstring(), pathUtf8))
    {
        yyjson_mut_obj_add_strncpy(doc, root, "artifact_root", pathUtf8.c_str(), pathUtf8.size());
    }

    yyjson_mut_obj_add_uint(doc, root, "duration_ms", result.durationMs);
    yyjson_mut_obj_add_int(doc, root, "passed", result.passed);
    yyjson_mut_obj_add_int(doc, root, "failed", result.failed);
    yyjson_mut_obj_add_int(doc, root, "skipped", result.skipped);

    if (! result.failureMessage.empty())
    {
        std::string failUtf8;
        if (ConvertUtf8(result.failureMessage, failUtf8))
        {
            yyjson_mut_obj_add_strncpy(doc, root, "failureMessage", failUtf8.c_str(), failUtf8.size());
        }
    }

    yyjson_mut_val* cases = yyjson_mut_arr(doc);
    if (! cases)
    {
        return;
    }

    for (const auto& item : result.cases)
    {
        AddCaseJson(doc, cases, item);
    }

    yyjson_mut_obj_add_val(doc, root, "cases", cases);
    yyjson_mut_obj_add_bool(doc, root, "fail_fast", g_options.failFast);
    yyjson_mut_obj_add_real(doc, root, "timeout_scale", g_options.timeoutScale);

    static_cast<void>(WriteJsonBlob(path, doc));
}

void WriteRunJson(const SelfTestRunResult& result, const std::filesystem::path& path)
{
    if (path.empty())
    {
        return;
    }

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return;
    }

    auto freeDoc         = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });
    yyjson_mut_val* root = yyjson_mut_obj(doc);
    if (! root)
    {
        return;
    }

    yyjson_mut_doc_set_root(doc, root);

    std::string started;
    if (ConvertUtf8(result.startedUtcIso, started))
    {
        yyjson_mut_obj_add_strncpy(doc, root, "run_started_utc", started.c_str(), started.size());
    }

    yyjson_mut_obj_add_uint(doc, root, "duration_ms", result.durationMs);
    yyjson_mut_obj_add_bool(doc, root, "fail_fast", result.failFast);
    yyjson_mut_obj_add_real(doc, root, "timeout_scale", result.timeoutScale);

    int passed  = 0;
    int failed  = 0;
    int skipped = 0;

    yyjson_mut_val* suites = yyjson_mut_arr(doc);
    if (! suites)
    {
        return;
    }

    for (const auto& suite : result.suites)
    {
        passed += suite.passed;
        failed += suite.failed;
        skipped += suite.skipped;
        AddSuiteJson(doc, suites, suite);
    }

    yyjson_mut_obj_add_val(doc, root, "suites", suites);
    yyjson_mut_obj_add_int(doc, root, "passed", passed);
    yyjson_mut_obj_add_int(doc, root, "failed", failed);
    yyjson_mut_obj_add_int(doc, root, "skipped", skipped);

    static_cast<void>(WriteJsonBlob(path, doc));
}

} // namespace SelfTest

#endif // _DEBUG
