#pragma once

#include "SettingsStore.h"

#include <filesystem>
#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

// App-layer settings-file policy shared by Preferences and command dispatch.
// This depends on SettingsHotReload and ShellExecute, so it intentionally does
// not live in the policy-neutral Common library.
namespace SettingsFileLauncher
{
[[nodiscard]] HRESULT Open(HWND owner, std::wstring_view appId, Common::Settings::Settings settingsToCreate, std::filesystem::path& outPath) noexcept;
}
