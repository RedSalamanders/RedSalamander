#pragma once

#include <windows.h>

#ifdef ENABLE_TESTS
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
// Test-only seams. Thin forwarders to the pure path/marker helpers in CrashHandler.cpp's anonymous
// namespace. They perform no SEH, no minidump, and no UI — they exist so the dump-naming and marker
// write/read round-trip logic can be unit-tested without invoking the fused crash/UI paths.
[[nodiscard]] std::filesystem::path GetCrashMarkerPathForTest() noexcept;
[[nodiscard]] std::filesystem::path BuildDumpPathForTest(const std::filesystem::path& dir) noexcept;
[[nodiscard]] HRESULT WriteMarkerFileForTest(const std::filesystem::path& markerPath, std::wstring_view dumpPath) noexcept;
[[nodiscard]] std::wstring ReadMarkerDumpPathForTest(const std::filesystem::path& markerPath) noexcept;
#endif
} // namespace CrashHandler
