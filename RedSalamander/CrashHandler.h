#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#ifdef ENABLE_TESTS
#include <cstddef>
#include <filesystem>
#include <string>
#include <string_view>
#endif

namespace CrashHandler
{
// Installs a unified crash front door (best-effort):
// - SetUnhandledExceptionFilter (SEH)
// - std::terminate handler
// - CRT purecall/invalid-parameter handlers
void Install() noexcept;

// Writes a minidump + crash marker (best-effort).
// Intended for use in a top-level __except filter.
[[nodiscard]] int WriteDumpForException(EXCEPTION_POINTERS* exceptionPointers) noexcept;

// If a previous crash marker exists, shows a prompt and optionally opens the crash folder.
void ShowPreviousCrashUiIfPresent(HWND ownerWindow) noexcept;

// Deliberate crash path to validate the dump pipeline.
void TriggerCrashTest() noexcept;

#ifdef ENABLE_TESTS
// Test-only entry points for dump naming, retention and marker I/O in the owned test folder.
[[nodiscard]] std::filesystem::path GetCrashMarkerPathForTest() noexcept;
[[nodiscard]] std::filesystem::path BuildDumpPathForTest(const std::filesystem::path& dir) noexcept;
[[nodiscard]] size_t PruneOldReportsForTest(const std::filesystem::path& dir, std::wstring_view pendingDump = {});
[[nodiscard]] HRESULT WriteMarkerFileForTest(const std::filesystem::path& markerPath, std::wstring_view dumpPath) noexcept;
[[nodiscard]] std::wstring ReadMarkerDumpPathForTest(const std::filesystem::path& markerPath) noexcept;
#endif
} // namespace CrashHandler
