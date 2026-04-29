#pragma once

#include "resource.h"

#include <string>
#include <string_view>

#include <windows.h>

[[nodiscard]] HWND GetAboutDialogHandle() noexcept;
[[nodiscard]] HWND GetFatalErrorDialogHandle() noexcept;

#ifdef ENABLE_TESTS
struct FatalErrorDialogDebugSnapshot
{
    bool usesDxUiHost              = false;
    size_t visibleChildWindowCount = 0u;
    size_t bodyFirstVisibleLine    = 0u;
    size_t bodyVisibleLineCount    = 0u;
    size_t bodyTotalLineCount      = 0u;
    bool bodyCanScrollVertically   = false;
    bool themeDark                 = false;
    bool themeHighContrast         = false;
    bool themeRainbow              = false;
    uint32_t bodyFillArgb          = 0u;
    uint32_t bodyTextArgb          = 0u;
    uint64_t renderCount           = 0u;
    uint64_t resizeCount           = 0u;
    uint64_t resizeFailureCount    = 0u;
    std::wstring messageText;
};

void DebugShowFatalErrorDialog(HWND ownerWindow, const wchar_t* caption, const wchar_t* message) noexcept;
[[nodiscard]] bool DebugGetFatalErrorDialogSnapshot(FatalErrorDialogDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugScrollFatalErrorDialogByWheelDetents(int detents) noexcept;
[[nodiscard]] std::wstring_view DebugGetRedSalamanderHelpText() noexcept;
[[nodiscard]] bool DebugIsRedSalamanderDiagnosticsEnabledByDefault() noexcept;
#endif
