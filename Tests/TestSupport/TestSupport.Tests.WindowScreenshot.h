#pragma once

#include <cstdint>
#include <filesystem>
#include <span>
#ifndef NOMINMAX
#define NOMINMAX
#endif
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>

namespace RedSalamander::TestSupport
{
// A scenario-owned client-region or whole-frame color witness. Use a tolerance and minimum
// population rather than exact antialiasing pixels. Counts are returned even on rejection.
struct WindowScreenshotColorProbe final
{
    RECT clientBounds{};
    COLORREF expectedColor         = 0;
    BYTE tolerance                 = 12u;
    uint32_t minimumMatchingPixels = 1u;
    uint64_t matchedPixels         = 0u;
    // Whole-frame population witnesses need no UI-thread client-coordinate mapping.
    // When true, clientBounds is ignored and the complete native WGC frame is checked.
    bool wholeFrame = false;
};

// Captures only a window owned by the calling test process, including DirectComposition
// and non-client chrome. Does not activate, move, or send input to the window. The caller
// supplies an already-authorized test-sandbox path and a visible, non-minimized fixture.
// The capture/readback device uses WARP, independently of the application's renderer,
// so a hardware adapter is not a precondition for observing the actual composed HWND.
// Initializes and balances the Windows Runtime while preserving an existing COM
// STA/MTA; all capture objects are released before that initialization is balanced.
// Optional probes reject missing render layers with ERROR_INVALID_DATA after saving the
// diagnostic PNG. Rejected frames request a non-activating repaint while capture is active,
// with at most three event-driven frame attempts. A successful encoder alone cannot establish
// complete visual content. Client-region probes require the window's UI thread.
// Whole-frame probes may run on another owned-process thread; retries schedule
// repaint asynchronously there and never block that thread in a synchronous paint.
[[nodiscard]] HRESULT SaveWindowScreenshot(HWND window, const std::filesystem::path& path, std::span<WindowScreenshotColorProbe> probes = {}) noexcept;
} // namespace RedSalamander::TestSupport
