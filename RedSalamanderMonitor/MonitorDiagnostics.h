#pragma once

#include "Helpers.h"

namespace RedSalamanderMonitor
{
[[nodiscard]] inline bool IsSelfDiagnosticsEnabled() noexcept
{
    return Debug::detail::IsRuntimeMonitorDiagnosticsEnabled();
}

[[nodiscard]] inline bool ShouldDisplayInitialMonitorStatus() noexcept
{
    return IsSelfDiagnosticsEnabled();
}

[[nodiscard]] inline bool ShouldAcceptEtwEventForDisplay(const Debug::InfoParam& info, DWORD monitorProcessId) noexcept
{
    if (IsSelfDiagnosticsEnabled())
    {
        return true;
    }

    return info.processID != monitorProcessId;
}
} // namespace RedSalamanderMonitor
