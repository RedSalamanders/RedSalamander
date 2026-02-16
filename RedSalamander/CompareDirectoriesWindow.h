#pragma once

#include <filesystem>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#pragma warning(pop)

#include "AppTheme.h"
#include "SettingsStore.h"

struct IFileSystem;
class ShortcutManager;

[[nodiscard]] bool ShowCompareDirectoriesWindow(HWND owner,
                                                Common::Settings::Settings& settings,
                                                const AppTheme& theme,
                                                const ShortcutManager* shortcuts,
                                                wil::com_ptr<IFileSystem> baseFileSystem,
                                                std::filesystem::path leftRoot,
                                                std::filesystem::path rightRoot) noexcept;

void UpdateCompareDirectoriesWindowsTheme(const AppTheme& theme) noexcept;
