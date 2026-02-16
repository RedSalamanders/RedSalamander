#pragma once

#include <cstdint>
#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

namespace FileOperationsSelfTest
{
// Starts the self-test state machine (debug-only).
// The caller owns the timer; call Tick() periodically until it returns true.
void Start(HWND mainWindow) noexcept;

// Advances the self-test state machine.
// Returns true when the self-test is complete (success or failure).
bool Tick(HWND mainWindow) noexcept;

// Best-effort completion notification for host-driven file ops tasks.
void NotifyTaskCompleted(std::uint64_t taskId, HRESULT hr) noexcept;

// Returns true when the self-test has been started.
bool IsRunning() noexcept;

// Returns true if the self-test finished with a failure.
bool DidFail() noexcept;

// Returns the failure message when DidFail() is true (best-effort; empty otherwise).
std::wstring_view FailureMessage() noexcept;
} // namespace FileOperationsSelfTest
