#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#ifdef ENABLE_TESTS
#include "Common.Tests.WindowActivationGuard.h"
#endif

namespace Common
{
inline BOOL RequestForegroundWindow(HWND window) noexcept
{
#ifdef ENABLE_TESTS
    // The no-activation contract forbids issuing this request even when a CBT
    // hook would reject it. Decline before entering User32; keep the hook as a
    // fallback for native controls and other non-cooperating callers.
    if (RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread())
    {
        return FALSE;
    }
#endif
    return SetForegroundWindow(window);
}
} // namespace Common
