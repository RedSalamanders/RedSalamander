#pragma once

#include <cstdint>
#include <string>
#include <string_view>

#include "AppTheme.h"
#include "SettingsStore.h"

// Shows the Connection Manager dialog.
// Returns:
// - S_OK: user chose a connection (selectedConnectionNameOut set)
// - S_FALSE: user cancelled (selectedConnectionNameOut cleared)
// - failure HRESULT: unexpected error
HRESULT ShowConnectionManagerDialog(HWND owner,
                                    std::wstring_view appId,
                                    Common::Settings::Settings& settings,
                                    const AppTheme& theme,
                                    std::wstring_view filterPluginId,
                                    std::wstring& selectedConnectionNameOut) noexcept;

// Shows a modeless Connection Manager window (similar to Preferences).
// `targetPane` is an app-defined identifier (0=Left, 1=Right) used when the user clicks Connect.
[[nodiscard]] bool ShowConnectionManagerWindow(HWND owner,
                                               std::wstring_view appId,
                                               Common::Settings::Settings& settings,
                                               const AppTheme& theme,
                                               std::wstring_view filterPluginId,
                                               uint8_t targetPane) noexcept;

[[nodiscard]] HWND GetConnectionManagerDialogHandle() noexcept;
