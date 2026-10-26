#pragma once

#include <DxUi/DxUi.h>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <wil/com.h>

namespace FileOperationsSelfTest
{
// Copied observations for the FileOps popup/Issues harness, not a graphics ownership API.
// Native DXGI allocation/driver peaks and shared device storage are deliberately not attributed.
struct SurfaceResources final
{
    bool logicalBytesAvailable        = false;
    uint32_t width                    = 0u;
    uint32_t height                   = 0u;
    DXGI_FORMAT format                = DXGI_FORMAT_UNKNOWN;
    uint64_t singleTargetLogicalBytes = 0u;
    size_t brushEntries               = 0u;
    size_t textFormatEntries          = 0u;
    uint64_t renderCount              = 0u;
    uint64_t resizeCount              = 0u;
};

// Called synchronously on the host's UI thread, outside timed paints. No reference escapes.
[[nodiscard]] inline SurfaceResources CaptureSurfaceResources(const DxUi::WindowHost& host) noexcept
{
    SurfaceResources result{};
    result.brushEntries      = host.DebugGetBrushCacheSize();
    result.textFormatEntries = host.DebugGetConfiguredTextFormatCount();
    result.renderCount       = host.DebugGetRenderCount();
    result.resizeCount       = host.DebugGetResizeCount();
    auto* const context      = host.GetDeviceContext();
    if (! context)
        return result;
    wil::com_ptr_nothrow<ID2D1Image> target;
    context->GetTarget(target.put());
    wil::com_ptr_nothrow<ID2D1Bitmap1> bitmap;
    if (! target || FAILED(target.query_to(bitmap.put())) || ! bitmap)
        return result;
    const auto size = bitmap->GetPixelSize();
    result.width    = size.width;
    result.height   = size.height;
    result.format   = bitmap->GetPixelFormat().format;
    if (result.width && result.height && result.format == DXGI_FORMAT_B8G8R8A8_UNORM &&
        static_cast<uint64_t>(result.width) <= (std::numeric_limits<uint64_t>::max)() / 4u / result.height)
    {
        result.singleTargetLogicalBytes = static_cast<uint64_t>(result.width) * result.height * 4u;
        result.logicalBytesAvailable    = true;
    }
    return result;
}
} // namespace FileOperationsSelfTest
