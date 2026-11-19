#include "Framework.h"

#include "ShortcutCommandCatalog.h"

#include "Helpers.h"
#include "ShortcutManager.h"
#include "ShortcutText.h"

#include <algorithm>
#include <chrono>
#include <unordered_map>

std::vector<ShortcutCommandCatalogEntry> BuildShortcutCommandCatalog(const Common::Settings::ShortcutsSettings& shortcuts, ShortcutCommandContext context)
{
    const auto startedAt = std::chrono::steady_clock::now();
    const uint8_t eligibleMask =
        context == ShortcutCommandContext::Terminal
            ? static_cast<uint8_t>(CommandShortcutScopeMask(CommandShortcutScope::Application) | CommandShortcutScopeMask(CommandShortcutScope::FunctionBar) |
                                   CommandShortcutScopeMask(CommandShortcutScope::Terminal))
            : static_cast<uint8_t>(CommandShortcutScopeMask(CommandShortcutScope::Application) | CommandShortcutScopeMask(CommandShortcutScope::FunctionBar) |
                                   CommandShortcutScopeMask(CommandShortcutScope::FolderView));

    std::vector<ShortcutCommandCatalogEntry> result;
    std::unordered_map<std::wstring, size_t> indexByCommand;
    for (const CommandInfo& command : GetAllCommands())
    {
        if ((command.shortcutScopeMask & eligibleMask) == 0u)
        {
            continue;
        }
        ShortcutCommandCatalogEntry entry;
        entry.commandId   = command.id;
        entry.displayName = ShortcutText::GetCommandDisplayName(command.id);
        if (command.descriptionStringId != 0u)
        {
            entry.description = LoadStringResource(nullptr, command.descriptionStringId);
        }
        entry.searchText     = std::format(L"{} {} {} {}", entry.displayName, entry.description, command.id, command.searchKeywords);
        entry.visualId       = command.visualId;
        entry.stateSource    = command.stateSource;
        entry.paletteVisible = command.paletteVisible;
        indexByCommand.emplace(entry.commandId, result.size());
        result.push_back(std::move(entry));
    }

    std::unordered_map<uint32_t, const Common::Settings::ShortcutBinding*> effectiveBindings;
    const auto overlayBindings = [&](const std::vector<Common::Settings::ShortcutBinding>& bindings)
    {
        for (const auto& binding : bindings)
        {
            effectiveBindings[ShortcutManager::MakeChordKey(binding)] = &binding;
        }
    };
    if (context == ShortcutCommandContext::Terminal)
    {
        // Runtime precedence is Terminal, then Application, then Function Bar.
        overlayBindings(shortcuts.functionBar);
        overlayBindings(shortcuts.application);
        overlayBindings(shortcuts.terminal);
    }
    else
    {
        // Runtime precedence outside Terminal is Application, then Function Bar,
        // then Folder View.
        overlayBindings(shortcuts.folderView);
        overlayBindings(shortcuts.functionBar);
        overlayBindings(shortcuts.application);
    }

    const auto addEffectiveBindings = [&](const std::vector<Common::Settings::ShortcutBinding>& bindings)
    {
        for (const auto& binding : bindings)
        {
            const auto effective = effectiveBindings.find(ShortcutManager::MakeChordKey(binding));
            if (effective == effectiveBindings.end() || effective->second != &binding)
            {
                continue;
            }
            if (ShortcutIds::IsPseudoCommandId(binding.commandId))
            {
                continue;
            }
            const std::wstring canonical(CanonicalizeCommandId(binding.commandId));
            const auto found = indexByCommand.find(canonical);
            if (found == indexByCommand.end())
            {
                continue;
            }
            result[found->second].shortcutTexts.push_back(ShortcutText::FormatChordText(binding.keyPosition, binding.vk, binding.modifiers));
        }
    };
    if (context == ShortcutCommandContext::Terminal)
    {
        addEffectiveBindings(shortcuts.terminal);
        addEffectiveBindings(shortcuts.application);
        addEffectiveBindings(shortcuts.functionBar);
    }
    else
    {
        addEffectiveBindings(shortcuts.application);
        addEffectiveBindings(shortcuts.functionBar);
        addEffectiveBindings(shortcuts.folderView);
    }

    std::ranges::sort(result,
                      [](const auto& left, const auto& right) noexcept
    {
        const int nameOrder = CompareStringOrdinal(
            left.displayName.data(), static_cast<int>(left.displayName.size()), right.displayName.data(), static_cast<int>(right.displayName.size()), TRUE);
        return nameOrder == CSTR_LESS_THAN || (nameOrder == CSTR_EQUAL && left.commandId < right.commandId);
    });
    Debug::Perf::EmitDurationUs(L"shortcut.catalog.rebuild_us", Debug::Perf::ElapsedUs(startedAt), result.size(), 0u);
    return result;
}
