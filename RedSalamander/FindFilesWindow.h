#pragma once

#include <cstdint>
#include <filesystem>
#include <string>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#pragma warning(pop)

#include "AppTheme.h"
#include "SettingsStore.h"

struct IFileSystem;

struct FindFilesPaneContext
{
    wil::com_ptr<IFileSystem> fileSystem;
    std::wstring pluginId;
    std::wstring pluginShortId;
    std::wstring instanceContext;
    std::filesystem::path rootPluginPath;
};

[[nodiscard]] bool ShowFindFilesWindow(HWND owner,
                                       Common::Settings::Settings& settings,
                                       const AppTheme& theme,
                                       FindFilesPaneContext context) noexcept;

void UpdateFindFilesWindowsTheme(const AppTheme& theme) noexcept;

[[nodiscard]] HWND GetFindFilesWindowHandle() noexcept;

#ifdef _DEBUG
enum class FindFilesDebugOperation : uint8_t
{
    Find,
    Append,
    Intersect,
    Subtract,
};

struct FindFilesDebugSnapshot
{
    bool searchActive = false;
    size_t resultCount = 0;
    HRESULT lastStatusHint = S_OK;
    uint32_t warningFlags = 0;
    std::vector<std::wstring> fullPaths;
    std::wstring statusText;
};

[[nodiscard]] bool DebugConfigureFindFilesWindow(std::wstring rootPath,
                                                 std::wstring namePattern,
                                                 std::wstring contentPattern,
                                                 Common::Settings::SearchNameMode nameMode,
                                                 Common::Settings::SearchContentMode contentMode) noexcept;

[[nodiscard]] bool DebugSetFindFilesWindowOptions(bool recursive,
                                                  bool includeFiles,
                                                  bool includeDirectories,
                                                  bool preferIndex,
                                                  bool wantSnippets) noexcept;

[[nodiscard]] bool DebugStartFindFilesWindowSearch(FindFilesDebugOperation operation) noexcept;

[[nodiscard]] bool DebugCancelFindFilesWindowSearch() noexcept;

[[nodiscard]] bool DebugGetFindFilesWindowSnapshot(FindFilesDebugSnapshot& out) noexcept;

[[nodiscard]] bool DebugWaitForFindFilesWindowIdle(uint32_t timeoutMs) noexcept;
#endif
