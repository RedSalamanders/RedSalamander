#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <cstdint>

#include <DxUi/Diagnostics.h>
#include <DxUi/DxUi.h>

namespace ViewerSpaceHosts
{
// DxUi's diagnostics are borrowed thread-local sinks within this plugin module.
// Forward only during an owned child callback, then restore the caller's sinks.
class ScopedHostDiagnostics final
{
public:
    ScopedHostDiagnostics() noexcept;
    ~ScopedHostDiagnostics() noexcept;
    ScopedHostDiagnostics(const ScopedHostDiagnostics&)            = delete;
    ScopedHostDiagnostics& operator=(const ScopedHostDiagnostics&) = delete;
    ScopedHostDiagnostics(ScopedHostDiagnostics&&)                 = delete;
    ScopedHostDiagnostics& operator=(ScopedHostDiagnostics&&)      = delete;

private:
    DxUi::Diagnostics::Sink _previousSink;
    DxUi::Diagnostics::PerformanceSink _previousPerformanceSink;
};

// A read-only, UI-thread snapshot of a nested DxUi host. Keep this limited to
// present-state facts useful to the governed ViewerSpace capture diagnostics.
struct HostDiagnostics final
{
    uint32_t clientWidthPx               = 0u;
    uint32_t clientHeightPx              = 0u;
    uint32_t windowBackgroundArgb        = 0u;
    uint32_t textArgb                    = 0u;
    uint64_t renderCount                 = 0u;
    uint64_t resizeCount                 = 0u;
    uint64_t resizeFailureCount          = 0u;
    uint64_t presentFailureCount         = 0u;
    uint64_t prepareD2DFlushFailureCount = 0u;
    uint64_t invalidateCount             = 0u;
    uint32_t presentationMode            = 0u;
    bool diagnosticsAvailable            = false;
    bool hasWindow                       = false;
    bool visible                         = false;
    bool hasRoot                         = false;
    bool hasD2DContext                   = false;
    bool dark                            = false;
    bool highContrast                    = false;
};

[[nodiscard]] HostDiagnostics CaptureHostDiagnostics(const DxUi::WindowHost& host, HWND hwnd) noexcept;
} // namespace ViewerSpaceHosts
