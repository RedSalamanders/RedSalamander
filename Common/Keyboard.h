#pragma once

#include <cstdint>
#include <string_view>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

namespace Common::Keyboard
{
enum class KeyPosition : uint8_t
{
    None = 0u,
    NumberRowMinus,
    NumberRowPlus,
};

inline constexpr uint16_t kNumberRowMinusScanCode = 0x0Cu;
inline constexpr uint16_t kNumberRowPlusScanCode  = 0x0Du;

[[nodiscard]] constexpr uint16_t ScanCodeFromKeyMessageLParam(LPARAM lParam) noexcept
{
    return static_cast<uint16_t>((static_cast<ULONG_PTR>(lParam) >> 16u) & 0xFFu);
}

[[nodiscard]] constexpr bool IsExtendedKeyMessageLParam(LPARAM lParam) noexcept
{
    return (static_cast<ULONG_PTR>(lParam) & (1ull << 24u)) != 0u;
}

[[nodiscard]] constexpr uint16_t ScanCodeForPosition(KeyPosition position) noexcept
{
    switch (position)
    {
        case KeyPosition::NumberRowMinus: return kNumberRowMinusScanCode;
        case KeyPosition::NumberRowPlus: return kNumberRowPlusScanCode;
        case KeyPosition::None: break;
    }
    return 0u;
}

[[nodiscard]] constexpr std::string_view TokenForPosition(KeyPosition position) noexcept
{
    switch (position)
    {
        case KeyPosition::NumberRowMinus: return "numberRowMinus";
        case KeyPosition::NumberRowPlus: return "numberRowPlus";
        case KeyPosition::None: break;
    }
    return {};
}

[[nodiscard]] constexpr KeyPosition PositionFromToken(std::string_view token) noexcept
{
    if (token == "numberRowMinus")
    {
        return KeyPosition::NumberRowMinus;
    }
    if (token == "numberRowPlus")
    {
        return KeyPosition::NumberRowPlus;
    }
    return KeyPosition::None;
}

[[nodiscard]] constexpr bool MatchesPosition(KeyPosition position, uint16_t scanCode, bool extended) noexcept
{
    return position != KeyPosition::None && ! extended && ScanCodeForPosition(position) == scanCode;
}

[[nodiscard]] constexpr KeyPosition PositionFromScanCode(uint16_t scanCode, bool extended) noexcept
{
    if (extended)
    {
        return KeyPosition::None;
    }
    if (scanCode == kNumberRowMinusScanCode)
    {
        return KeyPosition::NumberRowMinus;
    }
    if (scanCode == kNumberRowPlusScanCode)
    {
        return KeyPosition::NumberRowPlus;
    }
    return KeyPosition::None;
}

[[nodiscard]] inline UINT VirtualKeyForPosition(KeyPosition position, HKL keyboardLayout) noexcept
{
    const UINT scanCode = ScanCodeForPosition(position);
    if (scanCode == 0u)
    {
        return 0u;
    }

    const UINT vk = MapVirtualKeyExW(scanCode, MAPVK_VSC_TO_VK_EX, keyboardLayout);
    return vk == 0u ? MapVirtualKeyW(scanCode, MAPVK_VSC_TO_VK_EX) : vk;
}
} // namespace Common::Keyboard
