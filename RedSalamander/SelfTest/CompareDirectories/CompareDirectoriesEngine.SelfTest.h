#pragma once

#ifdef ENABLE_TESTS

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <string>
#include <vector>

#include "SelfTestCommon.h"

namespace CompareDirectoriesSelfTest
{
std::vector<std::wstring> ListCases(const SelfTest::SelfTestOptions& options = {}) noexcept;
[[nodiscard]] bool Run(const SelfTest::SelfTestOptions& options = {}, SelfTest::SelfTestSuiteResult* outResult = nullptr) noexcept;
}

#endif // ENABLE_TESTS
