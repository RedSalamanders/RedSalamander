#pragma once

#include <Windows.h>
#include <cstdint>
#include <string_view>

#include "AppTheme.h"
#include "SettingsStore.h"

enum class PluginType : uint8_t
{
    FileSystem,
    Viewer,
};

HRESULT ShowPluginConfigurationDialog(HWND owner,
                                      std::wstring_view appId,
                                      PluginType pluginType,
                                      std::wstring_view pluginId,
                                      std::wstring_view pluginName,
                                      Common::Settings::Settings& settings,
                                      const AppTheme& theme);

// Shows the plugin configuration dialog and updates `inOutWorkingSettings.plugins.configurationByPluginId` on OK.
// This does not apply changes to running plugins and does not persist settings to disk.
HRESULT EditPluginConfigurationDialog(HWND owner,
                                      PluginType pluginType,
                                      std::wstring_view pluginId,
                                      std::wstring_view pluginName,
                                      Common::Settings::Settings& baselineSettings,
                                      Common::Settings::Settings& inOutWorkingSettings,
                                      const AppTheme& theme);
