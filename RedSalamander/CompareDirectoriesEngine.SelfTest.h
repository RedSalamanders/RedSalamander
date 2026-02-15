#pragma once

#ifdef _DEBUG

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

namespace CompareDirectoriesSelfTest
{
[[nodiscard]] bool Run() noexcept;
}

#endif // _DEBUG

