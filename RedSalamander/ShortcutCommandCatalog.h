#pragma once

#include "CommandRegistry.h"
#include "SettingsStore.h"

#include <string>
#include <vector>

enum class ShortcutCommandContext : uint8_t
{
    Folder,
    Terminal,
};

struct ShortcutCommandCatalogEntry final
{
    std::wstring commandId;
    std::wstring displayName;
    std::wstring description;
    std::wstring searchText;
    std::vector<std::wstring> shortcutTexts;
    CommandVisualId visualId = CommandVisualId::None;
    CommandStateSource stateSource = CommandStateSource::LegacyHost;
    bool paletteVisible = true;
};

// App-layer canonical command projection shared by the binding-centric Helper
// and the command-centric palette. It intentionally depends on localized app
// resources and CommandRegistry, so it is not a Common helper.
[[nodiscard]] std::vector<ShortcutCommandCatalogEntry> BuildShortcutCommandCatalog(
    const Common::Settings::ShortcutsSettings& shortcuts,
    ShortcutCommandContext context);
