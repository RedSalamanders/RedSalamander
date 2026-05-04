#pragma once

#include "resource.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include <windows.h>

namespace Common::Settings
{
struct Settings;
}

[[nodiscard]] HWND GetAboutDialogHandle() noexcept;
[[nodiscard]] HWND GetFatalErrorDialogHandle() noexcept;

#ifdef ENABLE_TESTS
struct RereadAssociationsDebugSnapshot
{
    bool attempted                            = false;
    bool loaded                               = false;
    HRESULT hr                                = S_FALSE;
    size_t viewerActionCount                  = 0u;
    size_t editorActionCount                  = 0u;
    size_t userMenuActionCount                = 0u;
    size_t viewerExtensionMappingCount        = 0u;
    size_t fileSystemExtensionMappingCount    = 0u;
    size_t associationIconCacheSizeBefore     = 0u;
    size_t associationIconCacheSizeAfterClear = 0u;
    uint64_t leftRefreshCountBefore           = 0u;
    uint64_t leftRefreshCountAfter            = 0u;
    uint64_t rightRefreshCountBefore          = 0u;
    uint64_t rightRefreshCountAfter           = 0u;
    bool dynamicFileActionMenusRebuilt        = false;
    bool userMenuRebuilt                      = false;
    bool pluginsRefreshed                     = false;
    bool runtimeFoldersPreserved              = false;
};

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

void DebugSetRereadAssociationsSettingsForTest(const Common::Settings::Settings* settings) noexcept;
void DebugResetRereadAssociationsSnapshot() noexcept;
[[nodiscard]] bool DebugGetRereadAssociationsSnapshot(RereadAssociationsDebugSnapshot& out) noexcept;
void DebugShowFatalErrorDialog(HWND ownerWindow, const wchar_t* caption, const wchar_t* message) noexcept;
[[nodiscard]] bool DebugGetFatalErrorDialogSnapshot(FatalErrorDialogDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugScrollFatalErrorDialogByWheelDetents(int detents) noexcept;
[[nodiscard]] std::wstring_view DebugGetRedSalamanderHelpText() noexcept;
[[nodiscard]] bool DebugIsRedSalamanderDiagnosticsEnabledByDefault() noexcept;
#endif
