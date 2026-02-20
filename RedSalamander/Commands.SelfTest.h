#pragma once

#ifdef _DEBUG

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "SelfTestCommon.h"

namespace CommandsSelfTest
{
[[nodiscard]] bool Run(HWND mainWindow, const SelfTest::SelfTestOptions& options = {}, SelfTest::SelfTestSuiteResult* outResult = nullptr) noexcept;
}

#endif // _DEBUG

