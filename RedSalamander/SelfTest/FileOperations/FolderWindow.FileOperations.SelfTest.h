#pragma once

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#ifdef ENABLE_TESTS
#include "SelfTestCommon.h"
#endif

namespace FileOperationsSelfTest
{
// Starts the self-test state machine (debug-only).
// The caller owns the timer; call Tick() periodically until it returns true.
#ifdef ENABLE_TESTS
void Start(HWND mainWindow, const SelfTest::SelfTestOptions& options = {}) noexcept;
std::vector<std::wstring> BuildRunFilters(const SelfTest::SelfTestOptions& options);
std::vector<std::wstring> BuildExpectedCaseNames(const SelfTest::SelfTestOptions& options);
#endif

// Advances the self-test state machine.
// Returns true when the self-test is complete (success or failure).
bool Tick(HWND mainWindow) noexcept;

// Best-effort completion notification for host-driven file ops tasks.
void NotifyTaskCompleted(std::uint64_t taskId, HRESULT hr) noexcept;

// Returns true when the self-test has been started.
bool IsRunning() noexcept;
bool IsDone() noexcept;
#ifdef ENABLE_TESTS
SelfTest::SelfTestSuiteResult GetSuiteResult() noexcept;
#endif

// Returns true if the self-test finished with a failure.
bool DidFail() noexcept;

// Returns the failure message when DidFail() is true (best-effort; empty otherwise).
std::wstring_view FailureMessage() noexcept;
} // namespace FileOperationsSelfTest
