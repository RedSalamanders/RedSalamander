#pragma once

#include <Windows.h>
#include <cstddef>
#include <cstdint>
#include <string>
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

[[nodiscard]] HWND GetPluginConfigurationDialogHandle() noexcept;
void UpdatePluginConfigurationWindowsTheme(const AppTheme& theme) noexcept;

#ifdef ENABLE_TESTS
enum class PluginConfigurationDialogDebugFocusKind : uint8_t
{
    None,
    Panel,
    Edit,
    Combo,
    Toggle,
    Choice,
    CommandButton,
};

struct PluginConfigurationDialogDebugSnapshot
{
    bool usesDxUiCommandButtons                       = false;
    bool usesDxUiFormSurface                          = false;
    bool usesDxUiFormStatics                          = false;
    bool usesDxUiFormInputs                           = false;
    bool panelHasVerticalScrollbar                    = false;
    size_t legacyOwnerDrawCommandButtonCount          = 0u;
    size_t legacyOwnerDrawFormInputCount              = 0u;
    size_t visibleLegacyCommandButtonCount            = 0u;
    size_t visibleLegacyFormControlCount              = 0u;
    size_t visibleLegacyFormStaticCount               = 0u;
    size_t visibleLegacyFormInputCount                = 0u;
    size_t visibleDxCommandButtonHostCount            = 0u;
    size_t visibleDxFormHostCount                     = 0u;
    size_t visibleDxFormStaticHostCount               = 0u;
    size_t visibleDxFormInputHostCount                = 0u;
    int panelClientHeight                             = 0;
    int panelContentHeight                            = 0;
    int panelScrollPosY                               = 0;
    bool themeDark                                    = false;
    bool themeHighContrast                            = false;
    bool themeRainbow                                 = false;
    PluginConfigurationDialogDebugFocusKind focusKind = PluginConfigurationDialogDebugFocusKind::None;
    std::wstring focusLabel;
};

[[nodiscard]] bool DebugGetPluginConfigurationDialogSnapshot(PluginConfigurationDialogDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugScrollPluginConfigurationDialogByWheelDetents(int wheelDetents) noexcept;
[[nodiscard]] bool DebugFocusPluginConfigurationDialogFirstInput() noexcept;
[[nodiscard]] bool DebugGetPluginConfigurationDialogFirstVisibleToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetPluginConfigurationDialogVisibleToggleHostAndClientRectByLabel(std::wstring_view label, HWND& outHost, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetPluginConfigurationDialogFocusedHost(HWND& outHost) noexcept;
[[nodiscard]] bool DebugAdvancePluginConfigurationDialogTab(bool reverse) noexcept;
[[nodiscard]] bool DebugSetPluginConfigurationNextBrowsePath(std::wstring_view path) noexcept;
[[nodiscard]] bool DebugCancelPluginConfigurationNextBrowse() noexcept;
[[nodiscard]] bool DebugCancelPluginConfigurationDialog() noexcept;
#endif
