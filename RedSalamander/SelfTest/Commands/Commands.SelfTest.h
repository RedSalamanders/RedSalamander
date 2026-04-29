#pragma once

#ifdef ENABLE_TESTS

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <cstdint>
#include <string>

#include "SelfTestCommon.h"

#ifdef ENABLE_TESTS
void DebugResetConnectionManagerConnectNavigation() noexcept;
[[nodiscard]] bool DebugGetConnectionManagerConnectNavigation(uint8_t& outPane, std::wstring& outName) noexcept;
#endif

namespace CommandsSelfTest
{
[[nodiscard]] bool Run(HWND mainWindow, const SelfTest::SelfTestOptions& options = {}, SelfTest::SelfTestSuiteResult* outResult = nullptr) noexcept;
}

#endif // ENABLE_TESTS
