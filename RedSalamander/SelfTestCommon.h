#pragma once

// SelfTestCommon — debug-only self-test infrastructure shared by all suites.
//
// Artifacts are written to:
//   %LOCALAPPDATA%\RedSalamander\SelfTest\last_run\   (current run)
//   %LOCALAPPDATA%\RedSalamander\SelfTest\previous_run\  (previous run, kept for diffing)
//
// Everything in this header is compiled only in _DEBUG builds.  Release builds produce
// empty stub implementations so that call-sites do not need conditional compilation.

#ifdef _DEBUG

#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <span>
#include <string>
#include <string_view>
#include <vector>

#include <wil/com.h>

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
    // Multiply every timeout by this factor (use > 1.0 on slow CI machines).
    double timeoutScale = 1.0;
    // Write a results.json file to the suite artifact directory on completion.
    bool writeJsonSummary = true;
};

struct SelfTestCaseResult
{
    std::wstring name;
    enum class Status
    {
        passed,
        failed,
        skipped,
    } status;

    uint64_t durationMs = 0;
    std::wstring reason;
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

struct SelfTestRunResult
{
    std::wstring startedUtcIso;
    uint64_t durationMs = 0;
    bool failFast       = false;
    double timeoutScale = 1.0;
    std::vector<SelfTestSuiteResult> suites;
};

SelfTestOptions& GetSelfTestOptions() noexcept;
const std::filesystem::path& SelfTestRoot() noexcept;
std::filesystem::path GetSuiteRoot(SelfTestSuite suite);
std::filesystem::path GetSuiteArtifactPath(SelfTestSuite suite, std::wstring_view filename);
wil::com_ptr<IFileSystem> GetFileSystem(std::wstring_view pluginId) noexcept;

void RotateSelfTestRuns();
void InitSelfTestRun(const SelfTestOptions& options, SelfTestSuite suite);

void AppendSelfTestTrace(std::wstring_view msg) noexcept;
void AppendSuiteTrace(SelfTestSuite suite, std::wstring_view msg) noexcept;

void SetRunStartedUtcIso(std::wstring_view startedUtcIso) noexcept;
std::wstring_view GetRunStartedUtcIso() noexcept;

bool EnsureDirectory(const std::filesystem::path& path) noexcept;
[[nodiscard]] bool WriteBinaryFile(const std::filesystem::path& path, std::span<const std::byte> bytes) noexcept;
[[nodiscard]] bool WriteTextFile(const std::filesystem::path& path, std::wstring_view text);
[[nodiscard]] bool WriteTextFile(const std::filesystem::path& path, std::string_view text);

std::filesystem::path GetTempRoot(SelfTestSuite suite);
bool PathExists(const std::filesystem::path& p);
// Multiply baseMs by the current timeoutScale factor (see SelfTestOptions).
// Use this whenever waiting for asynchronous work in a test case.
uint64_t ScaleTimeout(uint64_t baseMs);

void WriteSuiteJson(const SelfTestSuiteResult& result, const std::filesystem::path& path);
void WriteRunJson(const SelfTestRunResult& result, const std::filesystem::path& path);

} // namespace SelfTest

#endif // _DEBUG
