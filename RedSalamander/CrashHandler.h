#pragma once

#include <windows.h>

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
} // namespace CrashHandler
