#pragma once

#ifdef _DEBUG

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "SelfTestCommon.h"

namespace CompareDirectoriesSelfTest
{
[[nodiscard]] bool Run(const SelfTest::SelfTestOptions& options = {}, SelfTest::SelfTestSuiteResult* outResult = nullptr) noexcept;
}

#endif // _DEBUG
