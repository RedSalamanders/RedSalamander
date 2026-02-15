#pragma once

// Shared exception name lookup for SEH exceptions
// This header can be included by any project that needs exception handling
// No dependencies - pure Win32 API

#pragma warning(push)
#pragma warning(disable : 4514) // unreferenced inline function has been removed

#include <Windows.h>

namespace exception
{

// Convert exception code to human-readable name (noexcept for SEH safety)
inline const wchar_t* GetExceptionName(DWORD exceptionCode) noexcept
{
    switch (exceptionCode)
    {
        case EXCEPTION_ACCESS_VIOLATION: return L"Access Violation";
        case EXCEPTION_ARRAY_BOUNDS_EXCEEDED: return L"Array Bounds Exceeded";
        case EXCEPTION_BREAKPOINT: return L"Breakpoint";
        case EXCEPTION_DATATYPE_MISALIGNMENT: return L"Datatype Misalignment";
        case EXCEPTION_FLT_DENORMAL_OPERAND: return L"Float Denormal Operand";
        case EXCEPTION_FLT_DIVIDE_BY_ZERO: return L"Float Divide by Zero";
        case EXCEPTION_FLT_INEXACT_RESULT: return L"Float Inexact Result";
        case EXCEPTION_FLT_INVALID_OPERATION: return L"Float Invalid Operation";
        case EXCEPTION_FLT_OVERFLOW: return L"Float Overflow";
        case EXCEPTION_FLT_STACK_CHECK: return L"Float Stack Check";
        case EXCEPTION_FLT_UNDERFLOW: return L"Float Underflow";
        case EXCEPTION_ILLEGAL_INSTRUCTION: return L"Illegal Instruction";
        case EXCEPTION_IN_PAGE_ERROR: return L"In Page Error";
        case EXCEPTION_INT_DIVIDE_BY_ZERO: return L"Integer Divide by Zero";
        case EXCEPTION_INT_OVERFLOW: return L"Integer Overflow";
        case EXCEPTION_INVALID_DISPOSITION: return L"Invalid Disposition";
        case EXCEPTION_NONCONTINUABLE_EXCEPTION: return L"Noncontinuable Exception";
        case EXCEPTION_PRIV_INSTRUCTION: return L"Privileged Instruction";
        case EXCEPTION_SINGLE_STEP: return L"Single Step";
        case EXCEPTION_STACK_OVERFLOW: return L"Stack Overflow";
        case 0xE06D7363: return L"C++ Exception";
        case 0xC0000420: return L"Assertion Failure";
        default: return L"Unknown Exception";
    }
}

} // namespace exception

#pragma warning(pop)
