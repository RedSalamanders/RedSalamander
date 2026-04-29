#pragma once

#ifdef ENABLE_TESTS

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "SelfTestCommon.h"

namespace CompareDirectoriesSelfTest
{
[[nodiscard]] bool Run(const SelfTest::SelfTestOptions& options = {}, SelfTest::SelfTestSuiteResult* outResult = nullptr) noexcept;
}

#endif // ENABLE_TESTS
