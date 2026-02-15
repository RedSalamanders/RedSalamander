#pragma once

#include <string_view>

#include "AppTheme.h"
#include "SettingsStore.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

enum class PrefCategory : int;

namespace PreferencesDialog
{
[[nodiscard]] bool
Show(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme, PrefCategory initialCategory) noexcept;

[[nodiscard]] HWND GetHandle() noexcept;
} // namespace PreferencesDialog
