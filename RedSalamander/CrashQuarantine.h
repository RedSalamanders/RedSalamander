#pragma once

#ifdef ENABLE_TESTS
#include <string>
#include <string_view>
#include <vector>
#endif

namespace Common::Settings
{
struct Settings;
}

namespace CrashQuarantine
{
// Best-effort: if a crash marker exists, offers to disable the last-active filesystem plugin in settings.
// Must be called after settings load and before plugin initialization.
void OfferPluginDisableIfPreviousCrashDetected(Common::Settings::Settings& settings) noexcept;

#ifdef ENABLE_TESTS
// Pure: given whether a crash marker exists and the active FS plugin ids from the previous session,
// returns the ids that should be offered for disable (non-empty ids not already disabled in settings).
// No file I/O, no UI, no SessionState. Returns empty when there is no offer to make.
[[nodiscard]] std::vector<std::wstring> SelectPluginsToDisable(bool markerExists,
                                                               const std::vector<std::wstring>& activeFileSystemPluginIds,
                                                               const Common::Settings::Settings& settings) noexcept;

// Test-only exposure of the apply step (disables a plugin id in the settings struct, mirroring the
// production YES branch). No I/O, no UI.
void DisablePluginIdInSettingsForTest(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept;
#endif
} // namespace CrashQuarantine
