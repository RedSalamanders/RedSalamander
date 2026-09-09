#pragma once

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "Keyboard.h"

namespace Common::Settings
{
struct ShortcutBinding;
struct ShortcutsSettings;
} // namespace Common::Settings

namespace ShortcutIds
{
inline constexpr std::wstring_view kUnassignedCommandId  = L"cmd/shortcut/unassigned";
inline constexpr std::wstring_view kPassThroughCommandId = L"cmd/shortcut/passthrough";

[[nodiscard]] constexpr bool IsUnassignedCommandId(std::wstring_view commandId) noexcept
{
    return commandId == kUnassignedCommandId;
}

[[nodiscard]] constexpr bool IsPassThroughCommandId(std::wstring_view commandId) noexcept
{
    return commandId == kPassThroughCommandId;
}

[[nodiscard]] constexpr bool IsPseudoCommandId(std::wstring_view commandId) noexcept
{
    return IsUnassignedCommandId(commandId) || IsPassThroughCommandId(commandId);
}
} // namespace ShortcutIds

class ShortcutManager
{
public:
    static constexpr uint32_t kModCtrl  = 1u;
    static constexpr uint32_t kModAlt   = 2u;
    static constexpr uint32_t kModShift = 4u;

    struct ShortcutChord final
    {
        uint32_t vk                               = 0;
        uint32_t modifiers                        = 0;
        Common::Keyboard::KeyPosition keyPosition = Common::Keyboard::KeyPosition::None;
    };

    void Clear() noexcept;
    void Load(const Common::Settings::ShortcutsSettings& shortcuts);

    [[nodiscard]] std::optional<std::wstring_view> FindApplicationCommand(uint32_t vk,
                                                                          uint32_t modifiers,
                                                                          uint16_t scanCode = 0u,
                                                                          bool extended     = false) const noexcept;
    [[nodiscard]] std::optional<std::wstring_view> FindFunctionBarCommand(uint32_t vk, uint32_t modifiers) const noexcept;
    [[nodiscard]] std::optional<std::wstring_view> FindFolderViewCommand(uint32_t vk,
                                                                         uint32_t modifiers,
                                                                         uint16_t scanCode = 0u,
                                                                         bool extended     = false) const noexcept;
    [[nodiscard]] std::optional<std::wstring_view> FindTerminalCommand(uint32_t vk, uint32_t modifiers, uint16_t scanCode, bool extended) const noexcept;
    [[nodiscard]] std::optional<ShortcutChord> TryGetShortcutForCommand(std::wstring_view commandId) const noexcept;

    [[nodiscard]] const std::vector<uint32_t>& GetApplicationConflicts() const noexcept;
    [[nodiscard]] const std::vector<uint32_t>& GetFunctionBarConflicts() const noexcept;
    [[nodiscard]] const std::vector<uint32_t>& GetFolderViewConflicts() const noexcept;
    [[nodiscard]] const std::vector<uint32_t>& GetTerminalConflicts() const noexcept;

    [[nodiscard]] static uint32_t MakeChordKey(uint32_t vk, uint32_t modifiers) noexcept;
    [[nodiscard]] static uint32_t MakeChordKey(Common::Keyboard::KeyPosition keyPosition, uint32_t modifiers) noexcept;
    [[nodiscard]] static uint32_t MakeChordKey(const Common::Settings::ShortcutBinding& binding) noexcept;

private:
    std::unordered_map<uint32_t, std::wstring> _application;
    std::unordered_map<uint32_t, std::wstring> _functionBar;
    std::unordered_map<uint32_t, std::wstring> _folderView;
    std::unordered_map<uint32_t, std::wstring> _terminal;
    std::unordered_map<std::wstring, uint32_t> _applicationReverse;
    std::unordered_map<std::wstring, uint32_t> _functionBarReverse;
    std::unordered_map<std::wstring, uint32_t> _folderViewReverse;
    std::unordered_map<std::wstring, uint32_t> _terminalReverse;
    std::vector<uint32_t> _applicationConflicts;
    std::vector<uint32_t> _functionBarConflicts;
    std::vector<uint32_t> _folderViewConflicts;
    std::vector<uint32_t> _terminalConflicts;
};
