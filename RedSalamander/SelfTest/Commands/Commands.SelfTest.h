#pragma once

#ifdef ENABLE_TESTS

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <cstdint>
#include <string>
#include <vector>

#include "SelfTestCommon.h"

#ifdef ENABLE_TESTS
void DebugResetConnectionManagerConnectNavigation() noexcept;
[[nodiscard]] bool DebugGetConnectionManagerConnectNavigation(uint8_t& outPane, std::wstring& outName) noexcept;
#endif

namespace CommandsSelfTest
{
std::vector<std::wstring> ListCases(const SelfTest::SelfTestOptions& options = {}) noexcept;
[[nodiscard]] bool Run(HWND mainWindow, const SelfTest::SelfTestOptions& options = {}, SelfTest::SelfTestSuiteResult* outResult = nullptr) noexcept;
} // namespace CommandsSelfTest

#endif // ENABLE_TESTS
