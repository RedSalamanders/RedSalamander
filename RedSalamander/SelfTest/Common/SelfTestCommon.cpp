#ifdef ENABLE_TESTS

#include "SelfTestCommon.h"
#include "Helpers.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <exception>
#include <filesystem>
#include <format>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <random>
#include <span>
#include <string>
#include <vector>

#include "AppDataPaths.h"
#include "FileSystemPluginManager.h"
#include <shlobj_core.h>
#include <winreg.h>

#include <bcrypt.h>
#pragma comment(lib, "bcrypt.lib")

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

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

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
constexpr std::wstring_view kPerfDirName{L"perf"};
constexpr std::wstring_view kTraceFileName{L"trace.txt"};
constexpr std::wstring_view kPerfMetricsFileName{L"perf_metrics.jsonl"};
constexpr std::wstring_view kRepoSpecsDirName{L"Specs"};
constexpr std::wstring_view kRepoTestRunsDirName{L"TestRuns"};
constexpr std::wstring_view kRepoGitDirName{L".git"};
constexpr std::wstring_view kRunsDirName{L"runs"};
constexpr std::wstring_view kArtifactsDirName{L"artifacts"};
constexpr std::wstring_view kScratchDirName{L"scratch"};
constexpr std::wstring_view kSelfTestArtifactDirName{L"selftest"};
constexpr std::wstring_view kAlternateVolumeTestSandboxDirName{L"RedSalamanderTestSandbox"};
constexpr std::wstring_view kUnifiedTestRootEnvVar{L"REDSALAMANDER_TEST_ROOT"};
constexpr std::wstring_view kUnifiedTestRunIdEnvVar{L"REDSALAMANDER_TEST_RUN_ID"};
constexpr std::wstring_view kSelfTestRootOverrideEnvVar{L"REDSALAMANDER_SELFTEST_ROOT"};
#if defined(RS_ASAN_DEBUG_BUILD)
constexpr std::wstring_view kSelfTestBuildFlavor{L"ASan Debug"};
#elif defined(_DEBUG)
constexpr std::wstring_view kSelfTestBuildFlavor{L"Debug"};
#else
constexpr std::wstring_view kSelfTestBuildFlavor{L"Release"};
#endif
constexpr const char* kSuiteCompareName  = "CompareDirectories";
constexpr const char* kSuiteFileOpsName  = "FileOperations";
constexpr const char* kSuiteCommandsName = "Commands";
constexpr int kRepoRootParentWalkLimit   = 10;

SelfTestOptions g_options{};
std::wstring g_runStartedUtcIso;
std::filesystem::file_time_type g_runArtifactStartTime{};
bool g_runArtifactStartTimeInitialized = false;
std::mutex g_traceLogMutex;
std::mutex g_inFlightSelfTestCaseMutex;
constexpr auto kArchiveFreshnessSlack = std::chrono::seconds(5);

struct InFlightSelfTestCase
{
    SelfTestSuite suite{};
    std::wstring name;
    std::chrono::steady_clock::time_point startedAt{};
};

std::optional<InFlightSelfTestCase> g_inFlightSelfTestCase;

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
        case SelfTestCaseResult::Status::crashed: return "crashed";
    }
    return "unknown";
}

[[nodiscard]] std::wstring_view SuiteArtifactPrefix(SelfTestSuite suite) noexcept
{
    switch (suite)
    {
        case SelfTestSuite::CompareDirectories: return L"compare";
        case SelfTestSuite::FileOperations: return L"fileops";
        case SelfTestSuite::Commands: return L"commands";
    }
    return L"unknown";
}

[[nodiscard]] std::wstring GetEnvironmentString(std::wstring_view name) noexcept
{
    if (name.empty())
    {
        return {};
    }

    const DWORD required = GetEnvironmentVariableW(name.data(), nullptr, 0u);
    if (required == 0u)
    {
        return {};
    }

    std::wstring value(required, L'\0');
    const DWORD written = GetEnvironmentVariableW(name.data(), value.data(), required);
    if (written == 0u || written >= required)
    {
        return {};
    }

    value.resize(written);
    return value;
}

[[nodiscard]] std::filesystem::path NormalizeSelfTestRootPath(const std::filesystem::path& root) noexcept
{
    if (root.empty())
    {
        return {};
    }

    std::error_code ec;
    if (! root.is_absolute())
    {
        const std::filesystem::path absoluteRoot = std::filesystem::absolute(root, ec);
        if (! ec && ! absoluteRoot.empty())
        {
            return absoluteRoot.lexically_normal();
        }
    }

    return root.lexically_normal();
}

[[nodiscard]] std::filesystem::path ExtendSelfTestPathForIo(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return {};
    }

    std::wstring normalized = path.native();
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');
    if (normalized.rfind(L"\\\\?\\", 0u) == 0u || normalized.rfind(L"\\\\.\\", 0u) == 0u)
    {
        return std::filesystem::path(normalized);
    }

    const DWORD required = GetFullPathNameW(normalized.c_str(), 0u, nullptr, nullptr);
    if (required == 0u)
    {
        return path;
    }

    std::wstring absolute(static_cast<size_t>(required), L'\0');
    const DWORD written = GetFullPathNameW(normalized.c_str(), required, absolute.data(), nullptr);
    if (written == 0u || written >= required)
    {
        return path;
    }

    absolute.resize(written);
    std::replace(absolute.begin(), absolute.end(), L'/', L'\\');
    if (absolute.rfind(L"\\\\", 0u) == 0u)
    {
        return std::filesystem::path(std::wstring(L"\\\\?\\UNC\\") + absolute.substr(2u));
    }

    return std::filesystem::path(std::wstring(L"\\\\?\\") + absolute);
}

[[nodiscard]] bool IsSafeSelfTestRunId(std::wstring_view value) noexcept
{
    if (value.empty())
    {
        return false;
    }

    return std::ranges::all_of(value, [](const wchar_t ch) noexcept
    {
        return (ch >= L'0' && ch <= L'9') || (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z') || ch == L'-' || ch == L'_';
    });
}

[[nodiscard]] std::filesystem::path GetUnifiedSelfTestRootFromEnvironment() noexcept
{
    const std::wstring rootText = GetEnvironmentString(kUnifiedTestRootEnvVar);
    const std::wstring runId    = GetEnvironmentString(kUnifiedTestRunIdEnvVar);
    if (rootText.empty() || ! IsSafeSelfTestRunId(runId))
    {
        return {};
    }

    const std::filesystem::path root = NormalizeSelfTestRootPath(std::filesystem::path(rootText));
    if (root.empty())
    {
        return {};
    }

    return root / std::wstring(kRunsDirName) / runId / std::wstring(kArtifactsDirName) / std::wstring(kSelfTestArtifactDirName);
}

[[nodiscard]] std::filesystem::path GetUnifiedTestRunScratchRootFromEnvironment() noexcept
{
    const std::wstring rootText = GetEnvironmentString(kUnifiedTestRootEnvVar);
    const std::wstring runId    = GetEnvironmentString(kUnifiedTestRunIdEnvVar);
    if (rootText.empty() || ! IsSafeSelfTestRunId(runId))
    {
        return {};
    }

    const std::filesystem::path root = NormalizeSelfTestRootPath(std::filesystem::path(rootText));
    if (root.empty())
    {
        return {};
    }

    return root / std::wstring(kRunsDirName) / runId / std::wstring(kScratchDirName);
}

[[nodiscard]] std::filesystem::path GetUnifiedTestRunScratchRootForVolume(const std::filesystem::path& volumeRoot) noexcept
{
    const std::wstring runId = GetEnvironmentString(kUnifiedTestRunIdEnvVar);
    if (! IsSafeSelfTestRunId(runId))
    {
        return {};
    }

    const std::filesystem::path normalizedVolumeRoot = NormalizeSelfTestRootPath(volumeRoot);
    if (normalizedVolumeRoot.empty())
    {
        return {};
    }

    return normalizedVolumeRoot / std::wstring(kAlternateVolumeTestSandboxDirName) / std::wstring(kRunsDirName) / runId / std::wstring(kScratchDirName);
}

[[nodiscard]] std::filesystem::path GetSelfTestRootOverrideFromEnvironment() noexcept
{
    const std::wstring value = GetEnvironmentString(kSelfTestRootOverrideEnvVar);
    if (value.empty())
    {
        return {};
    }

    return NormalizeSelfTestRootPath(std::filesystem::path(value));
}

[[nodiscard]] std::wstring SanitizeTestSandboxSegment(std::wstring_view value)
{
    std::wstring segment;
    segment.reserve(value.size());
    for (const wchar_t ch : value)
    {
        const bool allowed = std::iswalnum(static_cast<wint_t>(ch)) != 0 || ch == L'-' || ch == L'_' || ch == L'.';
        segment.push_back(allowed ? ch : L'_');
    }

    if (segment.empty())
    {
        segment = L"unnamed";
    }
    return segment;
}

[[nodiscard]] std::filesystem::path GetLegacyTestRunScratchRoot(SelfTestSuite suite)
{
    const std::filesystem::path root = SelfTestRoot();
    if (root.empty())
    {
        return {};
    }

    return root / kLastRunDirName / std::filesystem::path(SuiteArtifactPrefix(suite)) / kScratchDirName;
}

[[nodiscard]] TestSandbox CreateTestSandboxAtScratchRoot(SelfTestSuite suite,
                                                         std::wstring_view caseName,
                                                         const std::filesystem::path& scratchRoot,
                                                         std::wstring_view traceKind)
{
    const std::wstring suiteSegment = std::wstring(SuiteArtifactPrefix(suite));
    const std::wstring caseSegment  = SanitizeTestSandboxSegment(caseName);

    if (scratchRoot.empty())
    {
        AppendSelfTestTrace(std::format(L"{}: suite={} case={} unavailable", traceKind, suiteSegment, caseSegment));
        return {};
    }

    std::filesystem::path sandboxRoot = scratchRoot / suiteSegment / caseSegment;
    sandboxRoot                       = sandboxRoot.lexically_normal();

    std::error_code ec;
    std::filesystem::create_directories(sandboxRoot, ec);
    if (ec)
    {
        AppendSelfTestTrace(std::format(L"{}: suite={} case={} create failed path='{}' error={}",
                                        traceKind,
                                        suiteSegment,
                                        caseSegment,
                                        sandboxRoot.wstring(),
                                        ec.value()));
        return {};
    }

    AppendSelfTestTrace(std::format(L"{}: suite={} case={} root='{}'", traceKind, suiteSegment, caseSegment, sandboxRoot.wstring()));
    return TestSandbox{.root = std::move(sandboxRoot)};
}

// Trace / logging helpers (UTF-16 LE with BOM, one message per line)
void TruncateUtf16Log(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return;
    }

    const std::scoped_lock lock(g_traceLogMutex);
    const std::filesystem::path extendedPath = ExtendSelfTestPathForIo(path);
    wil::unique_handle file(
        CreateFileW(extendedPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
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

    const std::scoped_lock lock(g_traceLogMutex);
    const std::filesystem::path extendedPath = ExtendSelfTestPathForIo(path);
    wil::unique_handle file(
        CreateFileW(extendedPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
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
    const std::optional<std::string> converted = Common::Strings::TryUtf8FromUtf16Strict(text);
    if (! converted.has_value())
    {
        return false;
    }
    out = converted.value();
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

[[nodiscard]] std::wstring GetEnvironmentString(const wchar_t* name) noexcept
{
    if (! name || ! *name)
    {
        return {};
    }

    const DWORD required = GetEnvironmentVariableW(name, nullptr, 0);
    if (required == 0)
    {
        return {};
    }

    std::wstring buffer(static_cast<size_t>(required), L'\0');
    const DWORD written = GetEnvironmentVariableW(name, buffer.data(), required);
    if (written == 0 || written >= required)
    {
        return {};
    }

    buffer.resize(static_cast<size_t>(written));
    return buffer;
}

[[nodiscard]] std::wstring GetComputerNameString() noexcept
{
    wchar_t buf[MAX_COMPUTERNAME_LENGTH + 1]{};
    DWORD len = static_cast<DWORD>(std::size(buf));
    if (GetComputerNameW(buf, &len) == FALSE || len == 0)
    {
        return {};
    }
    return std::wstring(buf, buf + len);
}

[[nodiscard]] std::wstring GetUserNameString() noexcept
{
    wchar_t buf[256]{};
    DWORD len = static_cast<DWORD>(std::size(buf));
    if (GetUserNameW(buf, &len) == FALSE || len <= 1)
    {
        return {};
    }
    // GetUserNameW includes null terminator in len.
    return std::wstring(buf, buf + (len - 1));
}

[[nodiscard]] std::filesystem::path GetModuleDirectory() noexcept
{
    std::array<wchar_t, 4096> buffer{};
    const DWORD written = GetModuleFileNameW(nullptr, buffer.data(), static_cast<DWORD>(buffer.size()));
    if (written == 0 || written >= buffer.size())
    {
        return {};
    }

    std::filesystem::path p(buffer.data());
    if (p.has_parent_path())
    {
        return p.parent_path();
    }
    return {};
}

[[nodiscard]] bool PathExistsNoThrow(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    std::error_code ec;
    return std::filesystem::exists(path, ec) && ! ec;
}

[[nodiscard]] bool IsRepoRootCandidate(const std::filesystem::path& candidate) noexcept
{
    // A worktree stores .git as a file, while source snapshots may preserve only the solution or
    // project. Accept the union so every self-test uses one checkout-layout policy.
    return PathExistsNoThrow(candidate / kRepoGitDirName) || PathExistsNoThrow(candidate / L"RedSalamander.sln") ||
           PathExistsNoThrow(candidate / L"RedSalamander" / L"RedSalamander.vcxproj");
}

[[nodiscard]] std::filesystem::path TryFindRepoRootImpl() noexcept
{
    const std::wstring envRoot = GetEnvironmentString(L"REDSALAMANDER_REPO_ROOT");
    if (! envRoot.empty())
    {
        const std::filesystem::path root = std::filesystem::path(envRoot);
        if (IsRepoRootCandidate(root))
        {
            return root;
        }
    }

    const auto findFrom = [](std::filesystem::path cursor) noexcept
    {
        for (int i = 0; i < kRepoRootParentWalkLimit && ! cursor.empty(); ++i)
        {
            if (IsRepoRootCandidate(cursor))
            {
                return cursor;
            }

            const std::filesystem::path parent = cursor.parent_path();
            if (parent.empty() || parent == cursor)
            {
                break;
            }
            cursor = parent;
        }
        return std::filesystem::path{};
    };

    std::error_code ec;
    const std::filesystem::path currentPath = std::filesystem::current_path(ec);
    if (! ec)
    {
        if (const std::filesystem::path root = findFrom(currentPath); ! root.empty())
        {
            return root;
        }
    }

    return findFrom(GetModuleDirectory());
}

[[nodiscard]] std::filesystem::path TryFindTestRunsRoot() noexcept
{
    const std::filesystem::path repoRoot = TryFindRepoRootImpl();
    if (repoRoot.empty())
    {
        return {};
    }

    return repoRoot / kRepoSpecsDirName / kRepoTestRunsDirName;
}

[[nodiscard]] std::wstring TryReadRegistryStringValue(HKEY root, const wchar_t* subKey, const wchar_t* valueName) noexcept
{
    if (! subKey || ! valueName)
    {
        return {};
    }

    DWORD type  = 0;
    DWORD bytes = 0;
    LONG rc     = RegGetValueW(root, subKey, valueName, RRF_RT_REG_SZ, &type, nullptr, &bytes);
    if (rc != ERROR_SUCCESS || bytes < sizeof(wchar_t))
    {
        return {};
    }

    std::wstring buffer;
    buffer.resize(bytes / sizeof(wchar_t), L'\0');
    rc = RegGetValueW(root, subKey, valueName, RRF_RT_REG_SZ, &type, buffer.data(), &bytes);
    if (rc != ERROR_SUCCESS || bytes < sizeof(wchar_t))
    {
        return {};
    }

    // Ensure null-terminated / trim trailing nulls.
    while (! buffer.empty() && buffer.back() == L'\0')
    {
        buffer.pop_back();
    }

    return buffer;
}

#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL unique_bcrypt_* are intentionally non-copyable.
[[nodiscard]] bool TryComputeSha256(std::span<const std::byte> data, std::array<std::byte, 32>& outHash) noexcept
{
    outHash.fill(std::byte{0});

    if (data.size() > static_cast<size_t>(std::numeric_limits<ULONG>::max()))
    {
        return false;
    }

    BCRYPT_ALG_HANDLE algHandleRaw = nullptr;
    const NTSTATUS openStatus      = BCryptOpenAlgorithmProvider(&algHandleRaw, BCRYPT_SHA256_ALGORITHM, nullptr, 0);
    if (! BCRYPT_SUCCESS(openStatus) || ! algHandleRaw)
    {
        return false;
    }

    wil::unique_bcrypt_algorithm closeAlg(algHandleRaw);

    DWORD objLen  = 0;
    DWORD cbData  = 0;
    NTSTATUS prop = BCryptGetProperty(closeAlg.get(), BCRYPT_OBJECT_LENGTH, reinterpret_cast<PUCHAR>(&objLen), sizeof(objLen), &cbData, 0);
    if (! BCRYPT_SUCCESS(prop) || objLen == 0)
    {
        return false;
    }

    DWORD hashLen = 0;
    cbData        = 0;
    prop          = BCryptGetProperty(closeAlg.get(), BCRYPT_HASH_LENGTH, reinterpret_cast<PUCHAR>(&hashLen), sizeof(hashLen), &cbData, 0);
    if (! BCRYPT_SUCCESS(prop) || hashLen != static_cast<DWORD>(outHash.size()))
    {
        return false;
    }

    std::vector<std::byte> hashObject(static_cast<size_t>(objLen));
    BCRYPT_HASH_HANDLE hashHandleRaw = nullptr;
    const NTSTATUS createStatus      = BCryptCreateHash(closeAlg.get(), &hashHandleRaw, reinterpret_cast<PUCHAR>(hashObject.data()), objLen, nullptr, 0, 0);
    if (! BCRYPT_SUCCESS(createStatus) || ! hashHandleRaw)
    {
        return false;
    }

    wil::unique_bcrypt_hash destroyHash(hashHandleRaw);

    const NTSTATUS hashStatus =
        BCryptHashData(destroyHash.get(), reinterpret_cast<PUCHAR>(const_cast<std::byte*>(data.data())), static_cast<ULONG>(data.size()), 0);
    if (! BCRYPT_SUCCESS(hashStatus))
    {
        return false;
    }

    const NTSTATUS finishStatus = BCryptFinishHash(destroyHash.get(), reinterpret_cast<PUCHAR>(outHash.data()), hashLen, 0);
    if (! BCRYPT_SUCCESS(finishStatus))
    {
        return false;
    }

    return true;
}
#pragma warning(pop)

[[nodiscard]] std::wstring GetComputerHashName() noexcept
{
    // Prefer MachineGuid for stability. If it's unavailable, fall back to computer name.
    std::wstring seed = TryReadRegistryStringValue(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Cryptography", L"MachineGuid");
    if (seed.empty())
    {
        seed = GetComputerNameString();
    }
    if (seed.empty())
    {
        return L"unknown";
    }

    std::string seedUtf8;
    if (! ConvertUtf8(seed, seedUtf8))
    {
        return L"unknown";
    }

    std::array<std::byte, 32> hash{};
    if (! TryComputeSha256(std::as_bytes(std::span<const char>(seedUtf8.data(), seedUtf8.size())), hash))
    {
        return L"unknown";
    }

    constexpr wchar_t kHex[] = L"0123456789abcdef";
    std::wstring out;
    out.reserve(12);
    for (size_t i = 0; i < 6; ++i)
    {
        const uint8_t b = std::to_integer<uint8_t>(hash[i]);
        out.push_back(kHex[(b >> 4) & 0xF]);
        out.push_back(kHex[b & 0xF]);
    }

    return out;
}

[[nodiscard]] std::wstring GetTimestampFolderNameLocal() noexcept
{
    SYSTEMTIME st{};
    GetLocalTime(&st);
    return std::format(L"{0:04}-{1:02}-{2:02}_{3:02}{4:02}{5:02}", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
}

[[nodiscard]] bool ReadSmallTextFile(const std::filesystem::path& path, std::string& out) noexcept
{
    out.clear();

    if (path.empty())
    {
        return false;
    }

    wil::unique_handle file(CreateFileW(
        path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return false;
    }

    LARGE_INTEGER size{};
    if (! GetFileSizeEx(file.get(), &size) || size.QuadPart <= 0 || size.QuadPart > 1024 * 1024)
    {
        return false;
    }

    out.resize(static_cast<size_t>(size.QuadPart));
    DWORD read = 0;
    if (! ReadFile(file.get(), out.data(), static_cast<DWORD>(out.size()), &read, nullptr) || read != out.size())
    {
        out.clear();
        return false;
    }

    return true;
}

[[nodiscard]] std::string TrimAscii(std::string_view s) noexcept
{
    size_t start = 0;
    while (start < s.size() && (s[start] == ' ' || s[start] == '\t' || s[start] == '\r' || s[start] == '\n'))
    {
        ++start;
    }

    size_t end = s.size();
    while (end > start && (s[end - 1] == ' ' || s[end - 1] == '\t' || s[end - 1] == '\r' || s[end - 1] == '\n'))
    {
        --end;
    }

    return std::string(s.substr(start, end - start));
}

[[nodiscard]] std::filesystem::path ResolveGitDir(const std::filesystem::path& repoRoot) noexcept
{
    std::error_code ec;
    const std::filesystem::path dotGit = repoRoot / L".git";
    if (std::filesystem::is_directory(dotGit, ec) && ! ec)
    {
        return dotGit;
    }

    if (! std::filesystem::is_regular_file(dotGit, ec) || ec)
    {
        return {};
    }

    std::string text;
    if (! ReadSmallTextFile(dotGit, text))
    {
        return {};
    }

    const std::string line = TrimAscii(text);
    constexpr std::string_view kPrefix{"gitdir:"};
    if (line.size() <= kPrefix.size() || line.rfind(kPrefix, 0) != 0)
    {
        return {};
    }

    std::string_view rest(line);
    rest.remove_prefix(kPrefix.size());
    while (! rest.empty() && (rest.front() == ' ' || rest.front() == '\t'))
    {
        rest.remove_prefix(1);
    }

    if (rest.empty())
    {
        return {};
    }

    const std::optional<std::wstring> convertedRest = Common::Strings::TryUtf16FromUtf8Strict(rest);
    if (! convertedRest.has_value())
    {
        return {};
    }
    const std::wstring& restW = convertedRest.value();

    std::filesystem::path gitDir(restW);
    if (! gitDir.is_absolute())
    {
        gitDir = repoRoot / gitDir;
    }

    if (! std::filesystem::is_directory(gitDir, ec) || ec)
    {
        return {};
    }

    return gitDir;
}

[[nodiscard]] bool TryReadGitHeadInfo(const std::filesystem::path& repoRoot, std::wstring& outBranch, std::wstring& outCommit) noexcept
{
    outBranch.clear();
    outCommit.clear();

    const std::filesystem::path gitDir = ResolveGitDir(repoRoot);
    if (gitDir.empty())
    {
        return false;
    }

    std::string headText;
    if (! ReadSmallTextFile(gitDir / L"HEAD", headText))
    {
        return false;
    }

    const std::string head = TrimAscii(headText);
    constexpr std::string_view kRefPrefix{"ref:"};
    if (head.size() > kRefPrefix.size() && head.rfind(kRefPrefix, 0) == 0)
    {
        std::string_view ref(head);
        ref.remove_prefix(kRefPrefix.size());
        while (! ref.empty() && (ref.front() == ' ' || ref.front() == '\t'))
        {
            ref.remove_prefix(1);
        }
        if (ref.empty())
        {
            return false;
        }

        const std::optional<std::wstring> refW = Common::Strings::TryUtf16FromUtf8Strict(ref);
        if (! refW.has_value())
        {
            return false;
        }
        outBranch = std::filesystem::path(refW.value()).filename().wstring();

        // Try loose ref first
        std::string commitText;
        const std::filesystem::path looseRefPath = gitDir / std::filesystem::path(refW.value());
        if (ReadSmallTextFile(looseRefPath, commitText))
        {
            const std::string commit = TrimAscii(commitText);
            if (! commit.empty())
            {
                outCommit.assign(commit.begin(), commit.end());
                return true;
            }
        }

        // Fall back to packed-refs
        std::string packedText;
        if (ReadSmallTextFile(gitDir / L"packed-refs", packedText))
        {
            size_t pos = 0;
            while (pos < packedText.size())
            {
                const size_t lineEnd = packedText.find('\n', pos);
                const std::string_view line =
                    (lineEnd == std::string::npos) ? std::string_view(packedText).substr(pos) : std::string_view(packedText).substr(pos, lineEnd - pos);

                pos = (lineEnd == std::string::npos) ? packedText.size() : (lineEnd + 1);

                if (line.empty() || line[0] == '#' || line[0] == '^')
                {
                    continue;
                }

                const size_t space = line.find(' ');
                if (space == std::string::npos)
                {
                    continue;
                }

                const std::string_view hash    = line.substr(0, space);
                const std::string_view refName = line.substr(space + 1);
                if (refName == ref)
                {
                    const std::string commit(hash);
                    outCommit.assign(commit.begin(), commit.end());
                    return true;
                }
            }
        }

        return ! outBranch.empty();
    }

    // Detached HEAD (hash in HEAD file)
    if (! head.empty())
    {
        outBranch = L"(detached)";
        outCommit.assign(head.begin(), head.end());
        return true;
    }

    return false;
}

void CopyIfExists(const std::filesystem::path& source, const std::filesystem::path& dest) noexcept
{
    if (source.empty() || dest.empty())
    {
        return;
    }

    std::error_code ec;
    if (! std::filesystem::exists(source, ec) || ec)
    {
        return;
    }

    if (dest.has_parent_path())
    {
        EnsureDirectory(dest.parent_path());
    }

    static_cast<void>(CopyFileW(source.c_str(), dest.c_str(), FALSE));
}

[[nodiscard]] bool IsFreshRunArtifact(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    std::error_code ec;
    const auto writeTime = std::filesystem::last_write_time(path, ec);
    if (ec)
    {
        return false;
    }

    if (! g_runArtifactStartTimeInitialized)
    {
        return true;
    }

    return writeTime >= (g_runArtifactStartTime - kArchiveFreshnessSlack);
}

void AppendArchiveNote(std::wstring& notes, std::wstring_view note)
{
    if (note.empty())
    {
        return;
    }

    notes.append(note);
    if (! notes.ends_with(L"\n"))
    {
        notes.push_back(L'\n');
    }
}

void CopyIfExistsFresh(const std::filesystem::path& source, const std::filesystem::path& dest, std::wstring_view label, std::wstring& notes) noexcept
{
    if (source.empty() || dest.empty())
    {
        return;
    }

    std::error_code ec;
    if (! std::filesystem::exists(source, ec) || ec)
    {
        return;
    }

    if (! IsFreshRunArtifact(source))
    {
        AppendArchiveNote(notes, std::format(L"skipped stale artifact: {} ({})", label, source.wstring()));
        return;
    }

    CopyIfExists(source, dest);
}

void CopyFreshDirectoryFilesIfExists(const std::filesystem::path& source,
                                     const std::filesystem::path& dest,
                                     std::wstring_view label,
                                     std::wstring& notes) noexcept
{
    if (source.empty() || dest.empty())
    {
        return;
    }

    std::error_code ec;
    if (! std::filesystem::exists(source, ec) || ! std::filesystem::is_directory(source, ec))
    {
        return;
    }

    bool copiedAny = false;
    for (std::filesystem::recursive_directory_iterator it(source, ec), end; ! ec && it != end; it.increment(ec))
    {
        if (! it->is_regular_file(ec))
        {
            ec.clear();
            continue;
        }

        const std::filesystem::path filePath = it->path();
        if (! IsFreshRunArtifact(filePath))
        {
            continue;
        }

        const std::filesystem::path relative = std::filesystem::relative(filePath, source, ec);
        if (ec || relative.empty())
        {
            ec.clear();
            continue;
        }

        const std::filesystem::path targetPath = dest / relative;
        if (targetPath.has_parent_path())
        {
            static_cast<void>(EnsureDirectory(targetPath.parent_path()));
        }

        static_cast<void>(CopyFileW(filePath.c_str(), targetPath.c_str(), FALSE));
        copiedAny = true;
    }

    if (! copiedAny)
    {
        AppendArchiveNote(notes, std::format(L"skipped stale artifact directory: {} ({})", label, source.wstring()));
    }
}

[[nodiscard]] std::filesystem::path CreateUniqueRunFolder(const std::filesystem::path& base) noexcept
{
    if (base.empty())
    {
        return {};
    }

    const std::wstring ts = GetTimestampFolderNameLocal();
    std::error_code ec;

    std::filesystem::path candidate = base / ts;
    if (! std::filesystem::exists(candidate, ec) && ! ec)
    {
        std::filesystem::create_directories(candidate, ec);
        if (! ec)
        {
            return candidate;
        }
    }

    for (int i = 1; i < 1000; ++i)
    {
        candidate = base / std::format(L"{0}_{1:03}", ts, i);
        ec.clear();
        if (! std::filesystem::exists(candidate, ec) && ! ec)
        {
            std::filesystem::create_directories(candidate, ec);
            if (! ec)
            {
                return candidate;
            }
        }
    }

    return {};
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
    yyjson_mut_obj_add_uint(doc, caseObj, "repeat_index", testCase.repeatIndex);

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
    yyjson_mut_obj_add_uint(doc, suiteObj, "repeat_count", g_options.repeatCount);
    if (g_options.shuffleSeed.has_value())
    {
        yyjson_mut_obj_add_uint(doc, suiteObj, "shuffle_seed", g_options.shuffleSeed.value());
    }
    if (! g_options.caseFilter.empty())
    {
        std::string caseFilterUtf8;
        if (ConvertUtf8(g_options.caseFilter, caseFilterUtf8))
        {
            yyjson_mut_obj_add_strncpy(doc, suiteObj, "case_filter", caseFilterUtf8.c_str(), caseFilterUtf8.size());
        }
    }

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

std::filesystem::path TryFindRepoRoot() noexcept
{
    return TryFindRepoRootImpl();
}

std::filesystem::path GetLocalAppDataPath() noexcept
{
    return AppDataPaths::GetLocalAppDataPath();
}

HRESULT EnsureDirectoryExists(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return E_INVALIDARG;
    }

    std::error_code ec;
    std::filesystem::create_directories(ExtendSelfTestPathForIo(path), ec);
    if (ec)
    {
        return HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value()));
    }

    return PathExistsNoThrow(ExtendSelfTestPathForIo(path)) ? S_OK : HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
}

HRESULT WriteUtf8File(const std::filesystem::path& path, std::string_view text) noexcept
{
    if (path.empty())
    {
        return E_INVALIDARG;
    }

    if (path.has_parent_path())
    {
        const HRESULT directoryHr = EnsureDirectoryExists(path.parent_path());
        if (FAILED(directoryHr))
        {
            return directoryHr;
        }
    }

    const std::filesystem::path extendedPath = ExtendSelfTestPathForIo(path);
    wil::unique_handle file(CreateFileW(extendedPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    size_t offset = 0;
    while (offset < text.size())
    {
        const size_t remaining = text.size() - offset;
        const DWORD chunk      = static_cast<DWORD>(std::min<size_t>(remaining, 16ull * 1024ull * 1024ull));
        DWORD written          = 0;
        if (WriteFile(file.get(), text.data() + offset, chunk, &written, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (written != chunk)
        {
            return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
        }
        offset += written;
    }

    return FlushFileBuffers(file.get()) != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
}

std::string NarrowAscii(std::wstring_view text)
{
    std::string result;
    result.reserve(text.size());
    for (const wchar_t ch : text)
    {
        result.push_back(static_cast<char>(ch));
    }
    return result;
}

uint64_t StableDeviceHash(std::wstring_view value) noexcept
{
    uint64_t hash = 1469598103934665603ull;
    for (const wchar_t ch : value)
    {
        const auto lower = static_cast<uint64_t>(::towlower(static_cast<wint_t>(ch)));
        hash ^= lower & 0xFFu;
        hash *= 1099511628211ull;
        hash ^= (lower >> 8u) & 0xFFu;
        hash *= 1099511628211ull;
    }
    return hash;
}

std::optional<uint64_t> ExtractJsonUInt(std::string_view json, std::string_view key) noexcept
{
    if (key.empty())
    {
        return std::nullopt;
    }

    size_t keyStart = 0;
    while ((keyStart = json.find(key, keyStart)) != std::string_view::npos)
    {
        const size_t keyEnd = keyStart + key.size();
        if (keyStart > 0 && keyEnd < json.size() && json[keyStart - 1] == '"' && json[keyEnd] == '"')
        {
            size_t valueStart = json.find(':', keyEnd + 1);
            if (valueStart == std::string_view::npos)
            {
                return std::nullopt;
            }
            ++valueStart;
            while (valueStart < json.size() &&
                   (json[valueStart] == ' ' || json[valueStart] == '\t' || json[valueStart] == '\r' || json[valueStart] == '\n'))
            {
                ++valueStart;
            }

            if (valueStart >= json.size() || json[valueStart] < '0' || json[valueStart] > '9')
            {
                return std::nullopt;
            }

            uint64_t value = 0;
            for (; valueStart < json.size() && json[valueStart] >= '0' && json[valueStart] <= '9'; ++valueStart)
            {
                const uint64_t digit = static_cast<uint64_t>(json[valueStart] - '0');
                if (value > ((std::numeric_limits<uint64_t>::max)() - digit) / 10u)
                {
                    return std::nullopt;
                }
                value = value * 10u + digit;
            }
            return value;
        }
        keyStart = keyEnd;
    }

    return std::nullopt;
}

HRESULT LoadMtpPluginSelfTestExport(std::string_view exportName, wil::unique_hmodule& module, FARPROC& exportAddress) noexcept
{
    module.reset();
    exportAddress = nullptr;
    if (exportName.empty() || exportName.size() >= 128u || exportName.find('\0') != std::string_view::npos)
    {
        return E_INVALIDARG;
    }

    FileSystemPluginManager& pluginManager                  = FileSystemPluginManager::GetInstance();
    const FileSystemPluginManager::PluginEntry* mtpEntry = nullptr;
    for (const FileSystemPluginManager::PluginEntry& entry : pluginManager.GetPlugins())
    {
        if (CompareStringOrdinal(entry.id.data(), static_cast<int>(entry.id.size()), L"builtin/file-system-mtp", -1, TRUE) == CSTR_EQUAL)
        {
            mtpEntry = &entry;
            break;
        }
    }

    if (! mtpEntry || mtpEntry->disabled || ! mtpEntry->loadable || mtpEntry->path.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }
    if (mtpEntry->unloadDeferred || pluginManager.IsPluginPathDeferred(mtpEntry->path))
    {
        return HRESULT_FROM_WIN32(ERROR_BUSY);
    }

    module.reset(LoadLibraryExW(mtpEntry->path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
    if (! module)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    std::array<char, 128> exportNameBuffer{};
    std::copy(exportName.begin(), exportName.end(), exportNameBuffer.begin());
    exportAddress = GetProcAddress(module.get(), exportNameBuffer.data());
    if (! exportAddress)
    {
        module.reset();
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    return S_OK;
}

SelfTestOptions& GetSelfTestOptions() noexcept
{
    return g_options;
}

std::wstring_view GetSelfTestBuildFlavor() noexcept
{
    return kSelfTestBuildFlavor;
}

std::wstring GetSelfTestMachineHash() noexcept
{
    return GetComputerHashName();
}

std::vector<SelfTestCaseExecution> BuildSelfTestCaseExecutionOrder(const SelfTestOptions& options, std::span<const std::wstring> declaredCases)
{
    std::vector<std::wstring> selected;
    selected.reserve(declaredCases.size());
    for (const std::wstring& name : declaredCases)
    {
        if (CaseFilterMatches(options.caseFilter, name))
        {
            selected.push_back(name);
        }
    }

    if (options.shuffleSeed.has_value())
    {
        std::mt19937_64 rng(options.shuffleSeed.value());
        std::shuffle(selected.begin(), selected.end(), rng);
    }

    const uint32_t repeatCount = std::max(1u, options.repeatCount);
    std::vector<SelfTestCaseExecution> order;
    order.reserve(selected.size() * static_cast<size_t>(repeatCount));
    for (uint32_t repeatIndex = 1u; repeatIndex <= repeatCount; ++repeatIndex)
    {
        for (const std::wstring& name : selected)
        {
            order.push_back(SelfTestCaseExecution{.name = name, .repeatIndex = repeatIndex});
        }
    }

    return order;
}

const std::filesystem::path& SelfTestRoot() noexcept
{
    static const std::filesystem::path root = []
    {
        const std::filesystem::path overrideRoot = GetSelfTestRootOverrideFromEnvironment();
        if (! overrideRoot.empty())
        {
            return overrideRoot;
        }

        const std::filesystem::path unifiedRoot = GetUnifiedSelfTestRootFromEnvironment();
        if (! unifiedRoot.empty())
        {
            return unifiedRoot;
        }

        const std::filesystem::path base = AppDataPaths::GetLocalAppDataPath();
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

std::filesystem::path GetPerfArtifactPath(std::wstring_view filename)
{
    const std::filesystem::path root = SelfTestRoot();
    if (root.empty() || filename.empty())
    {
        return {};
    }

    return root / kLastRunDirName / kPerfDirName / filename;
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

        if (wil::compare_string_ordinal(entry.id, pluginId, true) == wistd::weak_ordering::equivalent)
        {
            return entry.fileSystem;
        }
    }

    return {};
}

TestSandbox AcquireTestSandbox(SelfTestSuite suite, std::wstring_view caseName) noexcept
{
    try
    {
        std::filesystem::path scratchRoot = GetUnifiedTestRunScratchRootFromEnvironment();
        if (scratchRoot.empty())
        {
            scratchRoot = GetLegacyTestRunScratchRoot(suite);
        }

        return CreateTestSandboxAtScratchRoot(suite, caseName, scratchRoot, L"TestSandbox");
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory: self-test path acquisition is used from noexcept case bodies. Convert path failures into a test-visible empty sandbox.
        AppendSelfTestTrace(L"TestSandbox: path acquisition failed with std::exception.");
        return {};
    }
}

TestSandbox AcquireTestSandboxOnVolume(SelfTestSuite suite, std::wstring_view caseName, const std::filesystem::path& volumeRoot) noexcept
{
    try
    {
        const std::filesystem::path scratchRoot = GetUnifiedTestRunScratchRootForVolume(volumeRoot);
        return CreateTestSandboxAtScratchRoot(suite, caseName, scratchRoot, L"TestSandbox alternate-volume");
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory: self-test path acquisition is used from noexcept case bodies. Convert path failures into a test-visible empty sandbox.
        AppendSelfTestTrace(L"TestSandbox alternate-volume: path acquisition failed with std::exception.");
        return {};
    }
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
    std::filesystem::create_directories(lastRun / kPerfDirName, ec);

    TruncateUtf16Log(lastRun / kTraceFileName);
    TruncateUtf16Log(lastRun / kCompareDirName / kTraceFileName);
    TruncateUtf16Log(lastRun / kFileOpsDirName / kTraceFileName);
    TruncateUtf16Log(lastRun / kCommandsDirName / kTraceFileName);
    std::filesystem::remove(lastRun / kPerfDirName / kPerfMetricsFileName, ec);

    Debug::Perf::ClearJsonlOutput();
}

void InitSelfTestRun(const SelfTestOptions& options)
{
    g_options                         = options;
    g_runArtifactStartTime            = std::filesystem::file_time_type::clock::now();
    g_runArtifactStartTimeInitialized = true;

    const std::filesystem::path lastRunRoot = SelfTestRoot() / kLastRunDirName;
    const std::filesystem::path perfPath    = lastRunRoot / kPerfDirName / kPerfMetricsFileName;

    std::wstring gitBranch;
    std::wstring gitCommit;
    const std::filesystem::path testRunsRoot = TryFindTestRunsRoot();
    if (! testRunsRoot.empty())
    {
        std::filesystem::path repoRoot = testRunsRoot;
        if (repoRoot.has_parent_path())
        {
            repoRoot = repoRoot.parent_path();
        }
        if (repoRoot.has_parent_path())
        {
            repoRoot = repoRoot.parent_path();
        }
        static_cast<void>(TryReadGitHeadInfo(repoRoot, gitBranch, gitCommit));
    }

    const SYSTEMTIME st = []()
    {
        SYSTEMTIME t{};
        GetSystemTime(&t);
        return t;
    }();
    const std::wstring runId = std::format(L"{0:04}-{1:02}-{2:02}_{3:02}{4:02}{5:02}", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);

    Debug::Perf::ConfigureJsonlOutput(perfPath, L"SelfTest", kSelfTestBuildFlavor, gitBranch, gitCommit, GetComputerHashName(), runId);
}

void AppendSelfTestTrace(std::wstring_view msg) noexcept
{
    AppendUtf16LogLine(SelfTestRoot() / kLastRunDirName / kTraceFileName, msg);
}

void AppendSuiteTrace(SelfTestSuite suite, std::wstring_view msg) noexcept
{
    AppendUtf16LogLine(GetSuiteArtifactPath(suite, kTraceFileName), msg);
}

void BeginInFlightSelfTestCase(SelfTestSuite suite, std::wstring_view name) noexcept
{
    const std::scoped_lock lock(g_inFlightSelfTestCaseMutex);
    g_inFlightSelfTestCase = InFlightSelfTestCase{
        .suite     = suite,
        .name      = std::wstring(name),
        .startedAt = std::chrono::steady_clock::now(),
    };
}

void EndInFlightSelfTestCase(SelfTestSuite suite, std::wstring_view name) noexcept
{
    const std::scoped_lock lock(g_inFlightSelfTestCaseMutex);
    if (! g_inFlightSelfTestCase.has_value())
    {
        return;
    }

    if (g_inFlightSelfTestCase->suite == suite && g_inFlightSelfTestCase->name == name)
    {
        g_inFlightSelfTestCase.reset();
    }
}

void TriggerSelfTestCaseCrashInjection(SelfTestSuite suite, std::wstring_view name) noexcept
{
    AppendSuiteTrace(suite, std::format(L"Case crash injection: {}", name));
    AppendSelfTestTrace(std::format(L"Case crash injection: {}", name));
    RaiseException(EXCEPTION_ACCESS_VIOLATION, EXCEPTION_NONCONTINUABLE, 0, nullptr);
    std::terminate();
}

void RecalculateSuiteSummary(SelfTestSuiteResult& suite) noexcept
{
    suite.passed = 0;
    suite.failed = 0;
    suite.skipped = 0;
    suite.failureMessage.clear();

    for (const SelfTestCaseResult& item : suite.cases)
    {
        switch (item.status)
        {
            case SelfTestCaseResult::Status::passed: ++suite.passed; break;
            case SelfTestCaseResult::Status::failed:
            case SelfTestCaseResult::Status::crashed:
            {
                ++suite.failed;
                if (suite.failureMessage.empty() && ! item.reason.empty())
                {
                    suite.failureMessage = item.reason;
                }
                break;
            }
            case SelfTestCaseResult::Status::skipped: ++suite.skipped; break;
        }
    }
}

void FlushSuiteJsonAfterCase(const SelfTestSuiteResult& suite) noexcept
{
    if (! g_options.writeJsonSummary)
    {
        return;
    }

    const std::filesystem::path jsonPath = GetSuiteArtifactPath(suite.suite, L"results.json");
    WriteSuiteJson(suite, jsonPath);
}

void AppendCaseResult(SelfTestSuiteResult& suite, SelfTestCaseResult result) noexcept
{
    suite.cases.push_back(std::move(result));
    switch (suite.cases.back().status)
    {
        case SelfTestCaseResult::Status::passed: ++suite.passed; break;
        case SelfTestCaseResult::Status::failed:
        case SelfTestCaseResult::Status::crashed:
        {
            ++suite.failed;
            if (suite.failureMessage.empty())
            {
                suite.failureMessage = suite.cases.back().reason;
            }
            break;
        }
        case SelfTestCaseResult::Status::skipped: ++suite.skipped; break;
    }

    FlushSuiteJsonAfterCase(suite);
}

void AppendCaseResult(
    SelfTestSuiteResult& suite, std::wstring_view name, SelfTestCaseResult::Status status, std::wstring_view reason, uint64_t durationMs) noexcept
{
    SelfTestCaseResult result{};
    result.name       = std::wstring(name);
    result.status     = status;
    result.durationMs = durationMs;
    result.repeatIndex = 1u;
    result.reason     = std::wstring(reason);
    AppendCaseResult(suite, std::move(result));
}

void MarkInFlightSelfTestCaseCrashed(SelfTestRunResult& runResult, std::wstring_view reason) noexcept
{
    std::optional<InFlightSelfTestCase> crashedCase;
    {
        const std::scoped_lock lock(g_inFlightSelfTestCaseMutex);
        crashedCase = g_inFlightSelfTestCase;
    }

    if (! crashedCase.has_value() || crashedCase->name.empty())
    {
        return;
    }

    const auto now = std::chrono::steady_clock::now();
    uint64_t durationMs = 0;
    if (now >= crashedCase->startedAt)
    {
        durationMs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(now - crashedCase->startedAt).count());
    }

    SelfTestSuiteResult* targetSuite = nullptr;
    const auto existingSuite = std::find_if(runResult.suites.begin(),
                                            runResult.suites.end(),
                                            [&](const SelfTestSuiteResult& item) noexcept { return item.suite == crashedCase->suite; });
    if (existingSuite != runResult.suites.end())
    {
        targetSuite = std::addressof(*existingSuite);
    }
    else
    {
        SelfTestSuiteResult suite{};
        suite.suite = crashedCase->suite;
        runResult.suites.push_back(std::move(suite));
        targetSuite = std::addressof(runResult.suites.back());
    }

    const auto existingCase = std::find_if(targetSuite->cases.begin(),
                                           targetSuite->cases.end(),
                                           [&](const SelfTestCaseResult& item) noexcept { return item.name == crashedCase->name; });
    if (existingCase != targetSuite->cases.end())
    {
        existingCase->status = SelfTestCaseResult::Status::crashed;
        existingCase->durationMs = durationMs;
        existingCase->reason = std::wstring(reason.empty() ? std::wstring_view(L"case crashed") : reason);
        RecalculateSuiteSummary(*targetSuite);
        FlushSuiteJsonAfterCase(*targetSuite);
    }
    else
    {
        SelfTestCaseResult result{};
        result.name = crashedCase->name;
        result.status = SelfTestCaseResult::Status::crashed;
        result.durationMs = durationMs;
        result.reason = std::wstring(reason.empty() ? std::wstring_view(L"case crashed") : reason);
        AppendCaseResult(*targetSuite, std::move(result));
    }

    AppendSuiteTrace(crashedCase->suite, std::format(L"Case crashed: {} reason='{}'", crashedCase->name, reason));
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
    std::filesystem::create_directories(ExtendSelfTestPathForIo(path), ec);
    if (ec)
    {
        return false;
    }

    return PathExists(path);
}

bool RemoveAll(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    std::error_code ec;
    static_cast<void>(std::filesystem::remove_all(ExtendSelfTestPathForIo(path), ec));
    return ! ec;
}

bool WriteBinaryFile(const std::filesystem::path& path, std::span<const std::byte> bytes) noexcept
{
    if (path.empty())
    {
        return false;
    }

    if (path.has_parent_path())
    {
        if (! EnsureDirectory(path.parent_path()))
        {
            return false;
        }
    }

    const std::filesystem::path extendedPath = ExtendSelfTestPathForIo(path);
    wil::unique_handle file(CreateFileW(extendedPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
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

    return EnsureDirectory(suiteRoot) ? suiteRoot : std::filesystem::path{};
}

bool PathExists(const std::filesystem::path& p)
{
    if (p.empty())
    {
        return false;
    }

    return GetFileAttributesW(ExtendSelfTestPathForIo(p).c_str()) != INVALID_FILE_ATTRIBUTES;
}

uint64_t ScaleTimeout(uint64_t baseMs)
{
    if (baseMs == 0u)
    {
        return 0u;
    }

    const double scaled = static_cast<double>(baseMs) * GetSelfTestOptions().timeoutScale;
    if (! std::isfinite(scaled))
    {
        return std::numeric_limits<uint64_t>::max();
    }

    if (baseMs > 0u && scaled < 1.0)
    {
        return 1u;
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
    yyjson_mut_obj_add_uint(doc, root, "repeat_count", g_options.repeatCount);
    if (g_options.shuffleSeed.has_value())
    {
        yyjson_mut_obj_add_uint(doc, root, "shuffle_seed", g_options.shuffleSeed.value());
    }

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
    yyjson_mut_obj_add_uint(doc, root, "repeat_count", result.repeatCount);
    if (result.shuffleSeed.has_value())
    {
        yyjson_mut_obj_add_uint(doc, root, "shuffle_seed", result.shuffleSeed.value());
    }
    if (! result.caseFilter.empty())
    {
        std::string caseFilterUtf8;
        if (ConvertUtf8(result.caseFilter, caseFilterUtf8))
        {
            yyjson_mut_obj_add_strncpy(doc, root, "case_filter", caseFilterUtf8.c_str(), caseFilterUtf8.size());
        }
    }

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

void TryArchiveLastRunToRepo(std::wstring_view area, int exitCode, uint64_t durationMs, const SelfTestRunResult* runResult) noexcept
{
    if (area.empty())
    {
        area = L"SelfTest";
    }

    const std::filesystem::path testRunsRoot = TryFindTestRunsRoot();
    if (testRunsRoot.empty())
    {
        AppendSelfTestTrace(L"ArchiveToRepo: repo root not found; skipping.");
        return;
    }

    const std::wstring profile = GetComputerHashName();

    std::filesystem::path repoRoot = testRunsRoot;
    if (repoRoot.has_parent_path())
    {
        repoRoot = repoRoot.parent_path();
    }
    if (repoRoot.has_parent_path())
    {
        repoRoot = repoRoot.parent_path();
    }

    std::wstring gitBranch;
    std::wstring gitCommit;
    static_cast<void>(TryReadGitHeadInfo(repoRoot, gitBranch, gitCommit));

    std::error_code ec;
    const std::filesystem::path areaRoot = testRunsRoot / profile / std::filesystem::path(area);
    std::filesystem::create_directories(areaRoot, ec);

    const std::filesystem::path runRoot = CreateUniqueRunFolder(areaRoot);
    if (runRoot.empty())
    {
        AppendSelfTestTrace(L"ArchiveToRepo: could not create run folder; skipping.");
        return;
    }

    // Log before copying so the archived trace includes the archive destination.
    AppendSelfTestTrace(std::format(L"ArchiveToRepo: {}", runRoot.wstring()));

    std::wstring archiveNotes;

    // Copy meaningful artifacts (explicit allow-list; never copy secrets or large work/ dirs).
    const std::filesystem::path lastRunRoot = SelfTestRoot() / kLastRunDirName;
    if (runResult)
    {
        WriteRunJson(*runResult, runRoot / L"selftest_run_results.json");
        for (const SelfTestSuiteResult& suite : runResult->suites)
        {
            WriteSuiteJson(suite, runRoot / std::format(L"{}_results.json", SuiteArtifactPrefix(suite.suite)));
        }
    }
    else
    {
        CopyIfExistsFresh(lastRunRoot / L"results.json", runRoot / L"selftest_run_results.json", L"selftest results", archiveNotes);
    }
    CopyIfExistsFresh(lastRunRoot / kTraceFileName, runRoot / L"selftest_run_trace.txt", L"selftest trace", archiveNotes);
    CopyIfExistsFresh(lastRunRoot / L"discovery.txt", runRoot / L"discovery.txt", L"selftest discovery", archiveNotes);

    auto copySuite = [&](std::wstring_view suiteDir, std::wstring_view prefix)
    {
        const std::filesystem::path suiteRoot = lastRunRoot / std::filesystem::path(suiteDir);
        if (! runResult)
        {
            CopyIfExistsFresh(suiteRoot / L"results.json", runRoot / std::format(L"{}_results.json", prefix), std::format(L"{} results", prefix), archiveNotes);
        }
        CopyIfExistsFresh(suiteRoot / kTraceFileName, runRoot / std::format(L"{}_trace.txt", prefix), std::format(L"{} trace", prefix), archiveNotes);
        CopyIfExistsFresh(suiteRoot / L"discovery.txt", runRoot / std::format(L"{}_discovery.txt", prefix), std::format(L"{} discovery", prefix), archiveNotes);
    };

    copySuite(kFileOpsDirName, L"fileops");
    copySuite(kCompareDirName, L"compare");
    copySuite(kCommandsDirName, L"commands");
    CopyFreshDirectoryFilesIfExists(lastRunRoot / kPerfDirName, runRoot / kPerfDirName, L"perf", archiveNotes);

    // env.txt (UTF-8 key/value lines, safe to check in; no secrets)
    {
        const std::wstring computerName = GetComputerNameString();
        const std::wstring userName     = GetUserNameString();
        const std::wstring exeDir       = GetModuleDirectory().wstring();
        const std::wstring cmdLine      = GetCommandLineW() ? GetCommandLineW() : L"";

        const SYSTEMTIME st = []()
        {
            SYSTEMTIME t{};
            GetSystemTime(&t);
            return t;
        }();
        const std::wstring nowUtcIso =
            std::format(L"{0:04}-{1:02}-{2:02}T{3:02}:{4:02}:{5:02}Z", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);

        std::wstring envText;
        envText.reserve(1024);
        envText.append(std::format(L"date: {}\n", nowUtcIso));
        if (! gitBranch.empty())
        {
            envText.append(std::format(L"git_branch: {}\n", gitBranch));
        }
        if (! gitCommit.empty())
        {
            envText.append(std::format(L"git_commit: {}\n", gitCommit));
        }
        envText.append(std::format(L"profile: {}\n", profile));
        envText.append(std::format(L"computer_hash: {}\n", profile));
        if (! computerName.empty())
        {
            envText.append(std::format(L"machine: {}\n", computerName));
        }
        if (! userName.empty())
        {
            envText.append(std::format(L"user: {}\n", userName));
        }
        envText.append(std::format(L"area: {}\n", std::wstring(area)));
        if (! exeDir.empty())
        {
            envText.append(std::format(L"exe_dir: {}\n", exeDir));
        }
        envText.append(std::format(L"repo_root: {}\n", repoRoot.wstring()));
        envText.append(std::format(L"selftest_root: {}\n", lastRunRoot.wstring()));
        if (! cmdLine.empty())
        {
            envText.append(std::format(L"command_line: {}\n", cmdLine));
        }

        static_cast<void>(WriteTextFile(runRoot / L"env.txt", envText));
    }

    // selftest.txt (lightweight run summary)
    {
        std::wstring text;
        text.reserve(512);
        text.append(std::format(L"=== selftest start (utc): {} ===\n", g_runStartedUtcIso.empty() ? L"<unknown>" : g_runStartedUtcIso));
        text.append(std::format(L"duration_ms: {}\n", durationMs));
        text.append(std::format(L"exit_code: {}\n", exitCode));
        static_cast<void>(WriteTextFile(runRoot / L"selftest.txt", text));
    }

    if (! archiveNotes.empty())
    {
        static_cast<void>(WriteTextFile(runRoot / L"archive_notes.txt", archiveNotes));
    }
}

} // namespace SelfTest

#endif // ENABLE_TESTS
