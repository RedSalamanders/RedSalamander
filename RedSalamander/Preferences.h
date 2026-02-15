#pragma once

#include <string_view>

#include "AppTheme.h"
#include "SettingsStore.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

[[nodiscard]] bool ShowPreferencesDialog(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme);
[[nodiscard]] bool ShowPreferencesDialogPlugins(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme);

[[nodiscard]] HWND GetPreferencesDialogHandle() noexcept;
