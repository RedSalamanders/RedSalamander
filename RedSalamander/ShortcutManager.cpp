#include "ShortcutManager.h"

#include <algorithm>

#include "CommandRegistry.h"
#include "SettingsStore.h"

namespace
{
void LoadBindings(const std::vector<Common::Settings::ShortcutBinding>& bindings,
                  std::unordered_map<uint32_t, std::wstring>& outMap,
                  std::unordered_map<std::wstring, uint32_t>& outReverseMap,
                  std::vector<uint32_t>& outConflicts)
{
    outMap.clear();
    outReverseMap.clear();
    outConflicts.clear();
    for (const auto& binding : bindings)
    {
        if (binding.commandId.empty())
        {
            continue;
        }
        const uint32_t key        = ShortcutManager::MakeChordKey(binding.vk, binding.modifiers);
        const auto [it, inserted] = outMap.try_emplace(key, binding.commandId);
        if (! inserted)
        {
            outConflicts.push_back(key);
            continue;
        }
        const std::wstring_view canonicalCommandId = CanonicalizeCommandId(binding.commandId);
        outReverseMap[std::wstring(canonicalCommandId)] = key;
    }
    std::sort(outConflicts.begin(), outConflicts.end());
    outConflicts.erase(std::unique(outConflicts.begin(), outConflicts.end()), outConflicts.end());
}
} // namespace

void ShortcutManager::Clear() noexcept
{
    _functionBar.clear();
    _folderView.clear();
    _functionBarReverse.clear();
    _folderViewReverse.clear();
    _functionBarConflicts.clear();
    _folderViewConflicts.clear();
}

void ShortcutManager::Load(const Common::Settings::ShortcutsSettings& shortcuts)
{
    LoadBindings(shortcuts.functionBar, _functionBar, _functionBarReverse, _functionBarConflicts);
    LoadBindings(shortcuts.folderView, _folderView, _folderViewReverse, _folderViewConflicts);
}

std::optional<std::wstring_view> ShortcutManager::FindFunctionBarCommand(uint32_t vk, uint32_t modifiers) const noexcept
{
    const uint32_t key = MakeChordKey(vk, modifiers);
    const auto it      = _functionBar.find(key);
    if (it == _functionBar.end())
    {
        return std::nullopt;
    }
    return it->second;
}

std::optional<std::wstring_view> ShortcutManager::FindFolderViewCommand(uint32_t vk, uint32_t modifiers) const noexcept
{
    const uint32_t key = MakeChordKey(vk, modifiers);
    const auto it      = _folderView.find(key);
    if (it == _folderView.end())
    {
        return std::nullopt;
    }
    return it->second;
}

const std::vector<uint32_t>& ShortcutManager::GetFunctionBarConflicts() const noexcept
{
    return _functionBarConflicts;
}

const std::vector<uint32_t>& ShortcutManager::GetFolderViewConflicts() const noexcept
{
    return _folderViewConflicts;
}

uint32_t ShortcutManager::MakeChordKey(uint32_t vk, uint32_t modifiers) noexcept
{
    const uint32_t clampedVk  = vk & 0xFFu;
    const uint32_t clampedMod = modifiers & 0x7u;
    return clampedVk | (clampedMod << 8);
}

std::optional<ShortcutManager::ShortcutChord> ShortcutManager::TryGetShortcutForCommand(std::wstring_view commandId) const noexcept
{
    if (commandId.empty())
    {
        return std::nullopt;
    }

    commandId = CanonicalizeCommandId(commandId);

    const auto findIn = [&](const std::unordered_map<std::wstring, uint32_t>& bindings) -> std::optional<ShortcutChord>
    {
        const auto it = bindings.find(std::wstring(commandId));
        if (it == bindings.end())
        {
            return std::nullopt;
        }
        const uint32_t key = it->second;
        ShortcutChord chord;
        chord.vk        = key & 0xFFu;
        chord.modifiers = (key >> 8) & 0x7u;
        return chord;
    };
    if (std::optional<ShortcutChord> found = findIn(_functionBarReverse))
    {
        return found;
    }
    if (std::optional<ShortcutChord> found = findIn(_folderViewReverse))
    {
        return found;
    }

    return std::nullopt;
}

