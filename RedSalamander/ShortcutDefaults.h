#pragma once

#include "SettingsStore.h"

namespace ShortcutDefaults
{
[[nodiscard]] Common::Settings::ShortcutsSettings CreateDefaultShortcuts();

[[nodiscard]] bool AreShortcutsDefault(const Common::Settings::ShortcutsSettings& shortcuts);

[[nodiscard]] bool IsDefaultFunctionBarBinding(const Common::Settings::ShortcutBinding& binding);

[[nodiscard]] bool IsDefaultFolderViewBinding(const Common::Settings::ShortcutBinding& binding);

void EnsureShortcutsInitialized(Common::Settings::Settings& settings);
} // namespace ShortcutDefaults
