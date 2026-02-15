#pragma once

#include "Preferences.Internal.h"

namespace PrefsPluginConfiguration
{
void Clear(PreferencesDialogState& state) noexcept;

[[nodiscard]] bool EnsureEditor(HWND parent, PreferencesDialogState& state, const PrefsPluginListItem& pluginItem) noexcept;

[[nodiscard]] bool HandleCommand(HWND host, PreferencesDialogState& state, UINT notifyCode, HWND hwndCtl) noexcept;

void LayoutCards(HWND host, PreferencesDialogState& state, int x, int& y, int width, HFONT dialogFont) noexcept;
} // namespace PrefsPluginConfiguration
