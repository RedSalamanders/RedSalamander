#pragma once

#include <filesystem>
#ifndef NOMINMAX
#define NOMINMAX
#endif
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>

namespace RedSalamander::TestSupport
{
// Captures only a window owned by the calling test process, including DirectComposition
// and non-client chrome. Does not activate, move, or send input to the window. The caller
// supplies an already-authorized test-sandbox path and a visible, non-minimized fixture.
[[nodiscard]] HRESULT SaveWindowScreenshot(HWND window, const std::filesystem::path& path) noexcept;
} // namespace RedSalamander::TestSupport
