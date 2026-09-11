#pragma once

#include "AppTheme.h"
#include "SettingsStore.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

void ShowCommandPaletteWindow(HWND owner,
                              const Common::Settings::ShortcutsSettings& shortcuts,
                              bool terminalContext,
                              const AppTheme& theme) noexcept;
void UpdateCommandPaletteWindowTheme(const AppTheme& theme) noexcept;
[[nodiscard]] HWND GetCommandPaletteWindowHandle() noexcept;

#ifdef ENABLE_TESTS
struct CommandPaletteDebugSnapshot final
{
    size_t rowCount = 0u;
    size_t selectedRow = static_cast<size_t>(-1);
    size_t selectedShortcutCount = 0u;
    bool terminalContext = false;
    bool selectedEnabled = false;
    std::wstring searchText;
    std::wstring selectedCommandId;
};
[[nodiscard]] bool DebugGetCommandPaletteSnapshot(CommandPaletteDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetCommandPaletteSearch(std::wstring_view text) noexcept;
#endif
