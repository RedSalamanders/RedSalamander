#pragma once

#include <cstdint>
#include <string>
#include <string_view>

namespace ShortcutText
{
[[nodiscard]] std::wstring VkToDisplayText(uint32_t vk) noexcept;

[[nodiscard]] std::wstring GetCommandDisplayName(std::wstring_view commandId) noexcept;

[[nodiscard]] std::wstring FormatChordText(uint32_t vk, uint32_t modifiers) noexcept;
} // namespace ShortcutText
