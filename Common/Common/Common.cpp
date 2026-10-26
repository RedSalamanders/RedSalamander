// Common.cpp : Defines the exported functions for the DLL.
//

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "Common.h"
#include "SynchronousIoCancelWatch.h"

namespace Common::SynchronousIoCancelWatch
{
namespace
{
// I35 R7.4: one owned-stage cleanup deadline per thread, shared by every module of the process, so a
// window a provider opens on a host worker thread is seen by the host's watch too (see the header).
thread_local std::atomic<ULONGLONG> g_cleanupDeadlineTick{0u};
} // namespace

std::atomic<ULONGLONG>* CurrentThreadCleanupDeadline() noexcept
{
    return &g_cleanupDeadlineTick;
}
} // namespace Common::SynchronousIoCancelWatch

// This is an example of an exported variable
COMMON_API int nCommon = 0;

// This is an example of an exported function.
COMMON_API int fnCommon()
{
    return 0;
}

// This is the constructor of a class that has been exported.
CCommon::CCommon()
{
    return;
}
