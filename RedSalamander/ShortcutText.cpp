#include "Framework.h"

#include "ShortcutText.h"

#include <array>
#include <optional>
#include <vector>

#include "CommandRegistry.h"
#include "Helpers.h"
#include "resource.h"

namespace ShortcutText
{
std::wstring VkToDisplayText(uint32_t vk) noexcept
{
    if (vk >= VK_F1 && vk <= VK_F24)
    {
        return std::format(L"F{}", static_cast<unsigned>(vk - VK_F1 + 1));
    }

    if ((vk >= static_cast<uint32_t>('0') && vk <= static_cast<uint32_t>('9')) || (vk >= static_cast<uint32_t>('A') && vk <= static_cast<uint32_t>('Z')))
    {
        wchar_t buf[2]{};
        buf[0] = static_cast<wchar_t>(vk);
        buf[1] = L'\0';
        return buf;
    }

    UINT scanCode = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (scanCode == 0)
    {
        return std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
    }

    bool extended = false;
    switch (vk)
    {
        case VK_LEFT:
        case VK_UP:
        case VK_RIGHT:
        case VK_DOWN:
        case VK_PRIOR:
        case VK_NEXT:
        case VK_END:
        case VK_HOME:
        case VK_INSERT:
        case VK_DELETE: extended = true; break;
    }

    LPARAM lParam = static_cast<LPARAM>(scanCode) << 16;
    if (extended)
    {
        lParam |= (1 << 24);
    }

    wchar_t keyName[64]{};
    const int length = GetKeyNameTextW(static_cast<LONG>(lParam), keyName, static_cast<int>(std::size(keyName)));
    if (length > 0)
    {
        return std::wstring(keyName, static_cast<size_t>(length));
    }

    return std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
}

std::wstring GetCommandDisplayName(std::wstring_view commandId) noexcept
{
    constexpr std::wstring_view kGoDriveRootPrefix = L"cmd/pane/goDriveRoot/";
    if (commandId.starts_with(kGoDriveRootPrefix) && commandId.size() > kGoDriveRootPrefix.size())
    {
        wchar_t driveLetter = commandId[kGoDriveRootPrefix.size()];
        if (driveLetter >= L'a' && driveLetter <= L'z')
        {
            driveLetter = static_cast<wchar_t>(driveLetter - L'a' + L'A');
        }

        if (driveLetter >= L'A' && driveLetter <= L'Z')
        {
            const std::wstring display = FormatStringResource(nullptr, IDS_FMT_CMD_GO_DRIVE_ROOT_WITH_LETTER, driveLetter);
            if (! display.empty())
            {
                return display;
            }
        }
    }

    const std::optional<unsigned int> displayNameIdOpt = TryGetCommandDisplayNameStringId(commandId);
    if (displayNameIdOpt.has_value())
    {
        const std::wstring display = LoadStringResource(nullptr, displayNameIdOpt.value());
        if (! display.empty())
        {
            return display;
        }
    }

    return std::wstring(commandId);
}

std::wstring FormatChordText(uint32_t vk, uint32_t modifiers) noexcept
{
    std::vector<std::wstring> parts;
    parts.reserve(4);

    if ((modifiers & 1u) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_CTRL));
    }

    if ((modifiers & 2u) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_ALT));
    }

    if ((modifiers & 4u) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_SHIFT));
    }

    parts.push_back(VkToDisplayText(vk));

    std::wstring result;
    for (size_t i = 0; i < parts.size(); ++i)
    {
        if (parts[i].empty())
        {
            continue;
        }

        if (! result.empty())
        {
            result.append(L" + ");
        }
        result.append(parts[i]);
    }
    return result;
}
} // namespace ShortcutText
