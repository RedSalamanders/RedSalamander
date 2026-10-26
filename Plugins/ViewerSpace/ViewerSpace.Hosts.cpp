#include "ViewerSpace.Hosts.h"

#include <algorithm>
#include <cmath>

#include "Helpers.h"

namespace ViewerSpaceHosts
{
namespace
{
void ForwardDiagnostic(std::wstring_view level, std::wstring_view message) noexcept
{
    if (level == L"Error" || level == L"error")
    {
        Debug::Error(L"ViewerSpace DxUi: {}", message);
    }
    else if (level == L"Warning" || level == L"warning")
    {
        Debug::Warning(L"ViewerSpace DxUi: {}", message);
    }
}

void ForwardPerformance(const DxUi::Diagnostics::PerformanceRecord& record) noexcept
{
    Debug::Perf::Emit(record.name == L"DxUi::Paint" ? L"viewer.space.dxui.paint_us" : record.name,
                      record.detail,
                      record.durationUs,
                      record.value0,
                      record.value1,
                      record.result);
}

[[nodiscard]] uint8_t ToByte(float component) noexcept
{
    if (! std::isfinite(component))
    {
        return 0u;
    }
    const float clamped = std::clamp(component, 0.0f, 1.0f);
    return static_cast<uint8_t>(clamped * 255.0f + 0.5f);
}

[[nodiscard]] uint32_t ToArgb(const D2D1_COLOR_F& color) noexcept
{
    return (static_cast<uint32_t>(ToByte(color.a)) << 24u) | (static_cast<uint32_t>(ToByte(color.r)) << 16u) | (static_cast<uint32_t>(ToByte(color.g)) << 8u) |
           static_cast<uint32_t>(ToByte(color.b));
}
} // namespace

ScopedHostDiagnostics::ScopedHostDiagnostics() noexcept : _previousSink(DxUi::Diagnostics::sink), _previousPerformanceSink(DxUi::Diagnostics::performanceSink)
{
    DxUi::Diagnostics::sink            = ForwardDiagnostic;
    DxUi::Diagnostics::performanceSink = Debug::Perf::IsCaptureEnabled() ? ForwardPerformance : _previousPerformanceSink;
}

ScopedHostDiagnostics::~ScopedHostDiagnostics() noexcept
{
    DxUi::Diagnostics::sink            = _previousSink;
    DxUi::Diagnostics::performanceSink = _previousPerformanceSink;
}

HostDiagnostics CaptureHostDiagnostics(const DxUi::WindowHost& host, HWND hwnd) noexcept
{
    HostDiagnostics diagnostics{};
    diagnostics.hasWindow = hwnd != nullptr && IsWindow(hwnd) != FALSE;
    if (diagnostics.hasWindow)
    {
        RECT client{};
        if (GetClientRect(hwnd, &client) != FALSE)
        {
            diagnostics.clientWidthPx  = static_cast<uint32_t>(std::max(0L, client.right - client.left));
            diagnostics.clientHeightPx = static_cast<uint32_t>(std::max(0L, client.bottom - client.top));
        }
        diagnostics.visible = IsWindowVisible(hwnd) != FALSE;
    }

    const DxUi::ThemePalette& palette = host.GetTheme();
    diagnostics.windowBackgroundArgb  = ToArgb(palette.windowBackground);
    diagnostics.textArgb              = ToArgb(palette.text);
    diagnostics.dark                  = palette.dark;
    diagnostics.highContrast          = palette.highContrast;
    diagnostics.hasRoot               = host.GetRoot() != nullptr;
    diagnostics.presentationMode      = static_cast<uint32_t>(host.GetPresentationMode());
#if DXUI_ENABLE_DIAGNOSTICS
    diagnostics.diagnosticsAvailable        = true;
    diagnostics.renderCount                 = host.DebugGetRenderCount();
    diagnostics.resizeCount                 = host.DebugGetResizeCount();
    diagnostics.resizeFailureCount          = host.DebugGetResizeFailureCount();
    diagnostics.presentFailureCount         = host.DebugGetPresentFailureCount();
    diagnostics.prepareD2DFlushFailureCount = host.DebugGetSwapChainPrepareD2DFlushFailureCount();
    diagnostics.invalidateCount             = host.DebugGetInvalidateCount();
    diagnostics.hasD2DContext               = host.DebugHasD2DContext();
#endif
    return diagnostics;
}
} // namespace ViewerSpaceHosts
