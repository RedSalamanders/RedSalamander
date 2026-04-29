#pragma once

#include "Preferences.Internal.h"

namespace PrefsPluginConfiguration
{
void Clear(PreferencesDialogState& state) noexcept;

void SetDetailsIdText(PreferencesDialogState& state, std::wstring_view text) noexcept;

void SetDetailsConfigErrorText(PreferencesDialogState& state, std::wstring_view text) noexcept;

void SetDetailsConfigEmptyStateText(PreferencesDialogState& state, std::wstring_view text) noexcept;

[[nodiscard]] bool EnsureEditor(HWND parent, PreferencesDialogState& state, const PrefsPluginListItem& pluginItem) noexcept;

void LayoutCards(HWND host, PreferencesDialogState& state, int x, int& y, int width, const PreferencesTypographyContext& typography) noexcept;
} // namespace PrefsPluginConfiguration
