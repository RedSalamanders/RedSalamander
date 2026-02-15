#pragma once

#include "AppTheme.h"
#include "SettingsStore.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

class ShortcutManager;

void ShowShortcutsWindow(HWND owner,
                         Common::Settings::Settings& settings,
                         const Common::Settings::ShortcutsSettings& shortcuts,
                         const ShortcutManager& shortcutManager,
                         const AppTheme& theme) noexcept;

void UpdateShortcutsWindowTheme(const AppTheme& theme) noexcept;

void UpdateShortcutsWindowData(const Common::Settings::ShortcutsSettings& shortcuts, const ShortcutManager& shortcutManager) noexcept;

[[nodiscard]] HWND GetShortcutsWindowHandle() noexcept;
