// pch.h: This is a precompiled header file.
// Files listed below are compiled only once, improving build performance for future builds.
// This also affects IntelliSense performance, including code completion and many code browsing features.
// However, files listed here are ALL re-compiled if any one of them is updated between builds.
// Do not add files here that you will be updating frequently as this negates the performance advantage.

#ifndef PCH_H
#define PCH_H

// add headers that you want to pre-compile here
#include "CppUnitTest.h"
#include "TestSupport/TestSupport.h"
#include <Windows.h>
#include <filesystem>
#include <string>
#include <string_view>
#include <system_error>

namespace PerformanceTests2
{

inline constexpr std::wstring_view kPerformanceTests2HarnessSegment{L"performance-tests2"};

[[nodiscard]] inline std::filesystem::path AcquirePerformanceTestSandbox(std::wstring_view caseName, std::error_code& ec)
{
    return RedSalamander::TestSupport::AcquireTestDirectory(
        {.harnessSegment = kPerformanceTests2HarnessSegment, .leafSegment = caseName, .fallbackRunIdPrefix = L"direct", .emptyLeafFallback = L"default"}, ec);
}

} // namespace PerformanceTests2

#endif // PCH_H
