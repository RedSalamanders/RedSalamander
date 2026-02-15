#pragma once

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

namespace Common::Settings
{
struct ShortcutBinding;
struct ShortcutsSettings;
} // namespace Common::Settings

class ShortcutManager
{
public:
    static constexpr uint32_t kModCtrl  = 1u;
    static constexpr uint32_t kModAlt   = 2u;
    static constexpr uint32_t kModShift = 4u;

    struct ShortcutChord final
    {
        uint32_t vk        = 0;
        uint32_t modifiers = 0;
    };

    void Clear() noexcept;
    void Load(const Common::Settings::ShortcutsSettings& shortcuts);

    [[nodiscard]] std::optional<std::wstring_view> FindFunctionBarCommand(uint32_t vk, uint32_t modifiers) const noexcept;
    [[nodiscard]] std::optional<std::wstring_view> FindFolderViewCommand(uint32_t vk, uint32_t modifiers) const noexcept;
    [[nodiscard]] std::optional<ShortcutChord> TryGetShortcutForCommand(std::wstring_view commandId) const noexcept;

    [[nodiscard]] const std::vector<uint32_t>& GetFunctionBarConflicts() const noexcept;
    [[nodiscard]] const std::vector<uint32_t>& GetFolderViewConflicts() const noexcept;

    [[nodiscard]] static uint32_t MakeChordKey(uint32_t vk, uint32_t modifiers) noexcept;

private:
    std::unordered_map<uint32_t, std::wstring> _functionBar;
    std::unordered_map<uint32_t, std::wstring> _folderView;
    std::vector<uint32_t> _functionBarConflicts;
    std::vector<uint32_t> _folderViewConflicts;
};
