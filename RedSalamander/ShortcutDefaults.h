#pragma once

#include "SettingsStore.h"

namespace ShortcutDefaults
{
[[nodiscard]] Common::Settings::ShortcutsSettings CreateDefaultShortcuts();

[[nodiscard]] bool AreShortcutsDefault(const Common::Settings::ShortcutsSettings& shortcuts);

void EnsureShortcutsInitialized(Common::Settings::Settings& settings);
} // namespace ShortcutDefaults
