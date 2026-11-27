#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <cstddef>
#include <filesystem>
#include <optional>
#include <string>
#include <string_view>

namespace CrashHandler
{
struct PendingReport
{
    std::filesystem::path directory;
    std::wstring dumpPath;
};

// Initialize local crash state and all fatal entry points before ordinary startup.
[[nodiscard]] HRESULT Install() noexcept;
// The caller acquires the directory through the governed test sandbox first.
[[nodiscard]] HRESULT InstallForProbe(const std::filesystem::path& directory) noexcept;

// Production fatal paths write once and exit. Harness-only SEH filters can continue to their own diagnostics.
[[nodiscard]] int WriteDumpForException(EXCEPTION_POINTERS* exceptionPointers) noexcept;

// Interactive startup is the only production caller. Claim the marker once and prune owned reports.
[[nodiscard]] std::optional<PendingReport> TakePendingReport();

[[noreturn]] void TriggerCrashTest(bool overlapWorkers) noexcept;

#ifdef ENABLE_TESTS
// Test-only entry points for dump naming, retention and marker I/O in the owned test folder.
[[nodiscard]] std::filesystem::path GetCrashMarkerPathForTest();
[[nodiscard]] std::filesystem::path BuildDumpPathForTest(const std::filesystem::path& dir);
[[nodiscard]] size_t PruneOldReportsForTest(const std::filesystem::path& dir, std::wstring_view pendingDump = {});
[[nodiscard]] HRESULT WriteMarkerFileForTest(const std::filesystem::path& markerPath, std::wstring_view dumpPath) noexcept;
[[nodiscard]] std::wstring ReadMarkerDumpPathForTest(const std::filesystem::path& markerPath);
[[nodiscard]] std::optional<PendingReport> TakePendingReportForTest(const std::filesystem::path& directory, bool interactive);
#endif
} // namespace CrashHandler
