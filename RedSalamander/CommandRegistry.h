#pragma once

#include <optional>
#include <span>
#include <string_view>

struct CommandInfo
{
    std::wstring_view id;
    unsigned int displayNameStringId = 0;
    unsigned int descriptionStringId = 0;
    unsigned int wmCommandId         = 0;
};

[[nodiscard]] std::wstring_view CanonicalizeCommandId(std::wstring_view commandId) noexcept;

[[nodiscard]] const CommandInfo* FindCommandInfo(std::wstring_view commandId) noexcept;

[[nodiscard]] const CommandInfo* FindCommandInfoByWmCommandId(unsigned int wmCommandId) noexcept;

[[nodiscard]] std::span<const CommandInfo> GetAllCommands() noexcept;

[[nodiscard]] std::optional<unsigned int> TryGetWmCommandId(std::wstring_view commandId) noexcept;

[[nodiscard]] std::optional<unsigned int> TryGetCommandDisplayNameStringId(std::wstring_view commandId) noexcept;

[[nodiscard]] std::optional<unsigned int> TryGetCommandDescriptionStringId(std::wstring_view commandId) noexcept;
