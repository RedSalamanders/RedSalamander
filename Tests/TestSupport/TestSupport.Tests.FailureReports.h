#pragma once

#include <crtdbg.h>
#include <cstdio>
#include <cstdlib>
#include <cwchar>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

// Failed standard assertions, Debug CRT assertions and STL checks in a test process go to stderr.
// Having an stderr stream lets the test runner retain the report. The failing check exits with code 3
// without opening a CRT or Windows Error Reporting dialog.
//
// Modeled on DxUi's Tests/Support/TestSupport.Tests.FailureReports.h. Every native test entry point calls it first
// (Tools/Tests/FailureReportRouting.Tests.ps1 enforces that), and the product executables call it only on their
// ENABLE_TESTS self-test paths, so an interactive Debug run keeps the dialogs a developer can attach a debugger to.
namespace RedSalamander::TestSupport
{
#if defined(_DEBUG)
inline int __cdecl EndOnFailureReport(int reportType, wchar_t* message, int* returnValue) noexcept
{
    if (reportType == _CRT_WARN)
    {
        return FALSE; // Warnings keep the CRT's own handling: the debugger output.
    }
    std::fputws(message ? message : L"(a CRT report without text)\n", stderr);
    std::fflush(stderr);
    *returnValue = 0;
    std::abort();
}
#endif

// Call first in the process entry point: before any test can fail a check.
inline void RouteFailureReportsAwayFromDialogs() noexcept
{
    static_cast<void>(_set_error_mode(_OUT_TO_STDERR));
    static_cast<void>(_set_abort_behavior(0u, _WRITE_ABORT_MSG | _CALL_REPORTFAULT));
    static_cast<void>(SetErrorMode(GetErrorMode() | SEM_FAILCRITICALERRORS | SEM_NOGPFAULTERRORBOX));
#if defined(_DEBUG)
    static_cast<void>(_CrtSetReportHookW2(_CRT_RPTHOOK_INSTALL, &EndOnFailureReport));
#endif
}
} // namespace RedSalamander::TestSupport
