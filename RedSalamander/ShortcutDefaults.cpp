#include "ShortcutDefaults.h"

#include <Windows.h>

#include <algorithm>
#include <string_view>
#include <tuple>

#include "ShortcutManager.h"

namespace
{
void AddBinding(std::vector<Common::Settings::ShortcutBinding>& dest, uint32_t vk, uint32_t modifiers, std::wstring_view commandId)
{
    Common::Settings::ShortcutBinding binding;
    binding.vk        = vk;
    binding.modifiers = modifiers;
    binding.commandId = std::wstring(commandId);
    dest.push_back(std::move(binding));
}

using NormalizedBinding = std::tuple<uint32_t, uint32_t, std::wstring>;

[[nodiscard]] std::vector<NormalizedBinding> NormalizeBindings(const std::vector<Common::Settings::ShortcutBinding>& bindings)
{
    std::vector<NormalizedBinding> result;
    result.reserve(bindings.size());
    for (const auto& binding : bindings)
    {
        if (binding.commandId.empty())
        {
            continue;
        }

        result.emplace_back(binding.vk, binding.modifiers & 0x7u, binding.commandId);
    }

    std::sort(result.begin(), result.end());
    return result;
}
} // namespace

Common::Settings::ShortcutsSettings ShortcutDefaults::CreateDefaultShortcuts()
{
    Common::Settings::ShortcutsSettings shortcuts;

    // Function bar bindings (F1..F12).
    AddBinding(shortcuts.functionBar, VK_F1, 0, L"cmd/app/showShortcuts");
    AddBinding(shortcuts.functionBar, VK_F1, ShortcutManager::kModAlt, L"cmd/app/openLeftDriveMenu");

    AddBinding(shortcuts.functionBar, VK_F2, 0, L"cmd/pane/rename");
    AddBinding(shortcuts.functionBar, VK_F2, ShortcutManager::kModCtrl, L"cmd/pane/sort/none");
    AddBinding(shortcuts.functionBar, VK_F2, ShortcutManager::kModAlt, L"cmd/app/openRightDriveMenu");
    AddBinding(shortcuts.functionBar, VK_F2, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/saveSelection");

    AddBinding(shortcuts.functionBar, VK_F3, 0, L"cmd/pane/view");
    AddBinding(shortcuts.functionBar, VK_F3, ShortcutManager::kModCtrl, L"cmd/pane/sort/name");
    AddBinding(shortcuts.functionBar, VK_F3, ShortcutManager::kModAlt, L"cmd/pane/alternateView");
    AddBinding(shortcuts.functionBar, VK_F3, ShortcutManager::kModShift, L"cmd/pane/openCurrentFolder");
    AddBinding(shortcuts.functionBar, VK_F3, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/app/viewWidth");

    AddBinding(shortcuts.functionBar, VK_F4, 0, L"cmd/pane/edit");
    AddBinding(shortcuts.functionBar, VK_F4, ShortcutManager::kModCtrl, L"cmd/pane/sort/extension");
    AddBinding(shortcuts.functionBar, VK_F4, ShortcutManager::kModAlt, L"cmd/app/exit");
    AddBinding(shortcuts.functionBar, VK_F4, ShortcutManager::kModShift, L"cmd/pane/editNew");
    AddBinding(shortcuts.functionBar, VK_F4, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/editWidth");

    AddBinding(shortcuts.functionBar, VK_F5, 0, L"cmd/pane/copyToOtherPane");
    AddBinding(shortcuts.functionBar, VK_F5, ShortcutManager::kModCtrl, L"cmd/pane/sort/time");
    AddBinding(shortcuts.functionBar, VK_F5, ShortcutManager::kModAlt, L"cmd/pane/pack");
    AddBinding(shortcuts.functionBar, VK_F5, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/saveSelection");

    AddBinding(shortcuts.functionBar, VK_F6, 0, L"cmd/pane/moveToOtherPane");
    AddBinding(shortcuts.functionBar, VK_F6, ShortcutManager::kModCtrl, L"cmd/pane/sort/size");
    AddBinding(shortcuts.functionBar, VK_F6, ShortcutManager::kModAlt, L"cmd/pane/unpack");
    AddBinding(shortcuts.functionBar, VK_F6, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/loadSelection");

    AddBinding(shortcuts.functionBar, VK_F7, 0, L"cmd/pane/createDirectory");
    AddBinding(shortcuts.functionBar, VK_F7, ShortcutManager::kModCtrl, L"cmd/pane/changeCase");
    AddBinding(shortcuts.functionBar, VK_F7, ShortcutManager::kModAlt, L"cmd/pane/find");
    AddBinding(shortcuts.functionBar, VK_F7, ShortcutManager::kModShift, L"cmd/pane/changeDirectory");

    AddBinding(shortcuts.functionBar, VK_F8, 0, L"cmd/pane/delete");
    AddBinding(shortcuts.functionBar, VK_F8, ShortcutManager::kModCtrl, L"cmd/pane/changeAttributes");
    AddBinding(shortcuts.functionBar, VK_F8, ShortcutManager::kModShift, L"cmd/pane/permanentDelete");

    AddBinding(shortcuts.functionBar, VK_F9, 0, L"cmd/pane/userMenu");
    AddBinding(shortcuts.functionBar, VK_F9, ShortcutManager::kModCtrl, L"cmd/pane/refresh");
    AddBinding(shortcuts.functionBar, VK_F9, ShortcutManager::kModAlt, L"cmd/pane/unpack");
    AddBinding(shortcuts.functionBar, VK_F9, ShortcutManager::kModShift, L"cmd/pane/hotPaths");
    AddBinding(shortcuts.functionBar, VK_F9, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/shares");

    AddBinding(shortcuts.functionBar, VK_F10, 0, L"cmd/pane/menu");
    AddBinding(shortcuts.functionBar, VK_F10, ShortcutManager::kModCtrl, L"cmd/app/compare");
    AddBinding(shortcuts.functionBar, VK_F10, ShortcutManager::kModAlt, L"cmd/pane/viewSpace");
    AddBinding(shortcuts.functionBar, VK_F10, ShortcutManager::kModShift, L"cmd/pane/contextMenu");
    AddBinding(shortcuts.functionBar, VK_F10, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/calculateDirectorySizes");
    AddBinding(shortcuts.functionBar, VK_F10, ShortcutManager::kModAlt | ShortcutManager::kModShift, L"cmd/pane/contextMenuCurrentDirectory");

    AddBinding(shortcuts.functionBar, VK_F11, 0, L"cmd/pane/connect");
    AddBinding(shortcuts.functionBar, VK_F11, ShortcutManager::kModCtrl, L"cmd/pane/zoomPanel");
    AddBinding(shortcuts.functionBar, VK_F11, ShortcutManager::kModAlt, L"cmd/pane/listOpenedFiles");
    AddBinding(shortcuts.functionBar, VK_F11, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/app/fullScreen");

    AddBinding(shortcuts.functionBar, VK_F12, 0, L"cmd/pane/disconnect");
    AddBinding(shortcuts.functionBar, VK_F12, ShortcutManager::kModCtrl, L"cmd/pane/filter");
    AddBinding(shortcuts.functionBar, VK_F12, ShortcutManager::kModAlt, L"cmd/pane/showFoldersHistory");

    // FolderView bindings (non-function-bar).
    AddBinding(shortcuts.folderView, VK_BACK, 0, L"cmd/pane/upOneDirectory");
    AddBinding(shortcuts.folderView, VK_TAB, 0, L"cmd/pane/switchPaneFocus");
    AddBinding(shortcuts.folderView, VK_TAB, ShortcutManager::kModShift, L"cmd/pane/switchPaneFocus");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('U'), ShortcutManager::kModCtrl, L"cmd/app/swapPanes");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('2'), ShortcutManager::kModAlt, L"cmd/pane/display/brief");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('3'), ShortcutManager::kModAlt, L"cmd/pane/display/detailed");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('4'), ShortcutManager::kModAlt, L"cmd/pane/display/extraDetailed");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('A'), ShortcutManager::kModCtrl, L"cmd/pane/selection/selectAll");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('C'), ShortcutManager::kModCtrl, L"cmd/pane/clipboardCopy");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('V'), ShortcutManager::kModCtrl, L"cmd/pane/clipboardPaste");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('L'), ShortcutManager::kModCtrl, L"cmd/pane/focusAddressBar");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('J'), ShortcutManager::kModCtrl, L"cmd/app/toggleFileOperationsFailedItems");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('D'), ShortcutManager::kModAlt, L"cmd/pane/focusAddressBar");
    AddBinding(shortcuts.folderView, VK_DOWN, ShortcutManager::kModAlt, L"cmd/pane/showFoldersHistory");
    AddBinding(shortcuts.folderView, VK_UP, ShortcutManager::kModAlt, L"cmd/pane/upOneDirectory");
    AddBinding(shortcuts.folderView, VK_OEM_2, ShortcutManager::kModAlt, L"cmd/app/about");
    AddBinding(shortcuts.folderView, VK_OEM_2, ShortcutManager::kModAlt | ShortcutManager::kModShift, L"cmd/app/about");

    for (wchar_t driveLetter = L'A'; driveLetter <= L'Z'; ++driveLetter)
    {
        std::wstring commandId = L"cmd/pane/goDriveRoot/";
        commandId.push_back(driveLetter);
        AddBinding(shortcuts.folderView, static_cast<uint32_t>(driveLetter), ShortcutManager::kModShift, commandId);
    }

    // Hot path shortcuts: Ctrl+1..Ctrl+9, Ctrl+0 and Ctrl+Shift+1..Ctrl+Shift+9, Ctrl+Shift+0
    for (int i = 0; i < 10; ++i)
    {
        const wchar_t digit                = (i < 9) ? static_cast<wchar_t>(L'1' + i) : L'0';
        const std::wstring goCommandId     = std::wstring(L"cmd/pane/hotPath/") + digit;
        const std::wstring setCommandId    = std::wstring(L"cmd/pane/setHotPath/") + digit;
        AddBinding(shortcuts.folderView, static_cast<uint32_t>(digit), ShortcutManager::kModCtrl, goCommandId);
        AddBinding(shortcuts.folderView, static_cast<uint32_t>(digit), ShortcutManager::kModCtrl | ShortcutManager::kModShift, setCommandId);
    }

    AddBinding(shortcuts.folderView, VK_RETURN, 0, L"cmd/pane/executeOpen");
    AddBinding(shortcuts.folderView, VK_RETURN, ShortcutManager::kModCtrl, L"cmd/pane/bringFilenameToCommandLine");
    AddBinding(shortcuts.folderView, VK_RETURN, ShortcutManager::kModAlt, L"cmd/pane/openProperties");
    AddBinding(shortcuts.folderView, VK_RETURN, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/bringFilenameToCommandLine");

    AddBinding(shortcuts.folderView, VK_SPACE, 0, L"cmd/pane/selectCalculateDirectorySizeNext");
    AddBinding(shortcuts.folderView, VK_SPACE, ShortcutManager::kModCtrl, L"cmd/pane/bringCurrentDirToCommandLine");
    AddBinding(shortcuts.folderView, VK_SPACE, ShortcutManager::kModAlt, L"cmd/pane/windowMenu");
    AddBinding(shortcuts.folderView, VK_SPACE, ShortcutManager::kModShift, L"cmd/pane/quickSearch");
    AddBinding(shortcuts.folderView, VK_SPACE, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/bringCurrentDirToCommandLine");

    AddBinding(shortcuts.folderView, VK_INSERT, 0, L"cmd/pane/selectNext");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModCtrl, L"cmd/pane/clipboardCopy");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModAlt, L"cmd/pane/copyPathAndNameAsText");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModShift, L"cmd/pane/clipboardPaste");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/copyPathAndFileName");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModCtrl | ShortcutManager::kModAlt, L"cmd/pane/copyPathAsText");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModAlt | ShortcutManager::kModShift, L"cmd/pane/copyNameAsText");

    AddBinding(shortcuts.folderView, VK_DELETE, 0, L"cmd/pane/moveToRecycleBin");
    AddBinding(shortcuts.folderView, VK_DELETE, ShortcutManager::kModShift, L"cmd/pane/permanentDeleteWithValidation");
    AddBinding(shortcuts.folderView, VK_DELETE, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/permanentDeleteWithValidation");

    return shortcuts;
}

bool ShortcutDefaults::AreShortcutsDefault(const Common::Settings::ShortcutsSettings& shortcuts)
{
    const Common::Settings::ShortcutsSettings defaults = CreateDefaultShortcuts();
    return NormalizeBindings(shortcuts.functionBar) == NormalizeBindings(defaults.functionBar) &&
           NormalizeBindings(shortcuts.folderView) == NormalizeBindings(defaults.folderView);
}

void ShortcutDefaults::EnsureShortcutsInitialized(Common::Settings::Settings& settings)
{
    if (! settings.shortcuts.has_value())
    {
        settings.shortcuts = CreateDefaultShortcuts();
        return;
    }

    Common::Settings::ShortcutsSettings& shortcuts = settings.shortcuts.value();

    auto findFunctionBarBinding = [&](uint32_t vk, uint32_t modifiers) noexcept -> Common::Settings::ShortcutBinding*
    {
        for (auto& binding : shortcuts.functionBar)
        {
            if (binding.vk == vk && (binding.modifiers & 0x7u) == modifiers)
            {
                return &binding;
            }
        }
        return nullptr;
    };

    const bool hasF1NoneBinding =
        std::any_of(shortcuts.functionBar.begin(),
                    shortcuts.functionBar.end(),
                    [](const Common::Settings::ShortcutBinding& binding) noexcept
                    { return binding.vk == static_cast<uint32_t>(VK_F1) && (binding.modifiers & 0x7u) == 0u && ! binding.commandId.empty(); });

    if (! hasF1NoneBinding)
    {
        AddBinding(shortcuts.functionBar, VK_F1, 0, L"cmd/app/showShortcuts");
    }

    Common::Settings::ShortcutBinding* ctrlF2 = findFunctionBarBinding(VK_F2, ShortcutManager::kModCtrl);
    if (! ctrlF2)
    {
        AddBinding(shortcuts.functionBar, VK_F2, ShortcutManager::kModCtrl, L"cmd/pane/sort/none");
    }
    else if (ctrlF2->commandId == L"cmd/pane/changeAttributes")
    {
        ctrlF2->commandId = L"cmd/pane/sort/none";
    }

    if (! findFunctionBarBinding(VK_F8, ShortcutManager::kModCtrl))
    {
        AddBinding(shortcuts.functionBar, VK_F8, ShortcutManager::kModCtrl, L"cmd/pane/changeAttributes");
    }

    auto findFolderViewBinding = [&](uint32_t vk, uint32_t modifiers) noexcept -> Common::Settings::ShortcutBinding*
    {
        for (auto& binding : shortcuts.folderView)
        {
            if (binding.vk == vk && (binding.modifiers & 0x7u) == modifiers)
            {
                return &binding;
            }
        }
        return nullptr;
    };

    if (! findFolderViewBinding(static_cast<uint32_t>('U'), ShortcutManager::kModCtrl))
    {
        AddBinding(shortcuts.folderView, static_cast<uint32_t>('U'), ShortcutManager::kModCtrl, L"cmd/app/swapPanes");
    }

    if (! findFolderViewBinding(static_cast<uint32_t>('2'), ShortcutManager::kModAlt))
    {
        AddBinding(shortcuts.folderView, static_cast<uint32_t>('2'), ShortcutManager::kModAlt, L"cmd/pane/display/brief");
    }

    if (! findFolderViewBinding(static_cast<uint32_t>('3'), ShortcutManager::kModAlt))
    {
        AddBinding(shortcuts.folderView, static_cast<uint32_t>('3'), ShortcutManager::kModAlt, L"cmd/pane/display/detailed");
    }

    if (! findFolderViewBinding(static_cast<uint32_t>('4'), ShortcutManager::kModAlt))
    {
        AddBinding(shortcuts.folderView, static_cast<uint32_t>('4'), ShortcutManager::kModAlt, L"cmd/pane/display/extraDetailed");
    }

    if (! findFolderViewBinding(static_cast<uint32_t>('A'), ShortcutManager::kModCtrl))
    {
        AddBinding(shortcuts.folderView, static_cast<uint32_t>('A'), ShortcutManager::kModCtrl, L"cmd/pane/selection/selectAll");
    }

    if (! findFolderViewBinding(static_cast<uint32_t>('C'), ShortcutManager::kModCtrl))
    {
        AddBinding(shortcuts.folderView, static_cast<uint32_t>('C'), ShortcutManager::kModCtrl, L"cmd/pane/clipboardCopy");
    }

    if (! findFolderViewBinding(static_cast<uint32_t>('V'), ShortcutManager::kModCtrl))
    {
        AddBinding(shortcuts.folderView, static_cast<uint32_t>('V'), ShortcutManager::kModCtrl, L"cmd/pane/clipboardPaste");
    }

    if (! findFolderViewBinding(static_cast<uint32_t>('L'), ShortcutManager::kModCtrl))
    {
        AddBinding(shortcuts.folderView, static_cast<uint32_t>('L'), ShortcutManager::kModCtrl, L"cmd/pane/focusAddressBar");
    }

    if (! findFolderViewBinding(static_cast<uint32_t>('J'), ShortcutManager::kModCtrl))
    {
        AddBinding(shortcuts.folderView, static_cast<uint32_t>('J'), ShortcutManager::kModCtrl, L"cmd/app/toggleFileOperationsFailedItems");
    }

    if (! findFolderViewBinding(static_cast<uint32_t>('D'), ShortcutManager::kModAlt))
    {
        AddBinding(shortcuts.folderView, static_cast<uint32_t>('D'), ShortcutManager::kModAlt, L"cmd/pane/focusAddressBar");
    }

    if (! findFolderViewBinding(VK_DOWN, ShortcutManager::kModAlt))
    {
        AddBinding(shortcuts.folderView, VK_DOWN, ShortcutManager::kModAlt, L"cmd/pane/showFoldersHistory");
    }

    if (! findFolderViewBinding(VK_UP, ShortcutManager::kModAlt))
    {
        AddBinding(shortcuts.folderView, VK_UP, ShortcutManager::kModAlt, L"cmd/pane/upOneDirectory");
    }

    if (! findFolderViewBinding(VK_OEM_2, ShortcutManager::kModAlt))
    {
        AddBinding(shortcuts.folderView, VK_OEM_2, ShortcutManager::kModAlt, L"cmd/app/about");
    }

    if (! findFolderViewBinding(VK_OEM_2, ShortcutManager::kModAlt | ShortcutManager::kModShift))
    {
        AddBinding(shortcuts.folderView, VK_OEM_2, ShortcutManager::kModAlt | ShortcutManager::kModShift, L"cmd/app/about");
    }

    for (wchar_t driveLetter = L'A'; driveLetter <= L'Z'; ++driveLetter)
    {
        if (findFolderViewBinding(static_cast<uint32_t>(driveLetter), ShortcutManager::kModShift))
        {
            continue;
        }

        std::wstring commandId = L"cmd/pane/goDriveRoot/";
        commandId.push_back(driveLetter);
        AddBinding(shortcuts.folderView, static_cast<uint32_t>(driveLetter), ShortcutManager::kModShift, commandId);
    }

    // Ensure hot path shortcuts exist.
    for (int i = 0; i < 10; ++i)
    {
        const wchar_t digit = (i < 9) ? static_cast<wchar_t>(L'1' + i) : L'0';
        const uint32_t vk   = static_cast<uint32_t>(digit);

        if (! findFolderViewBinding(vk, ShortcutManager::kModCtrl))
        {
            AddBinding(shortcuts.folderView, vk, ShortcutManager::kModCtrl, std::wstring(L"cmd/pane/hotPath/") + digit);
        }

        if (! findFolderViewBinding(vk, ShortcutManager::kModCtrl | ShortcutManager::kModShift))
        {
            AddBinding(shortcuts.folderView, vk, ShortcutManager::kModCtrl | ShortcutManager::kModShift, std::wstring(L"cmd/pane/setHotPath/") + digit);
        }
    }
}
