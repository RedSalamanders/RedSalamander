#include "ShortcutDefaults.h"

#include <algorithm>
#include <format>
#include <string_view>
#include <tuple>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "ShortcutManager.h"

namespace
{
constexpr uint32_t kCurrentShortcutMigrationVersion = 1u;

void AddBinding(std::vector<Common::Settings::ShortcutBinding>& dest, uint32_t vk, uint32_t modifiers, std::wstring_view commandId)
{
    Common::Settings::ShortcutBinding binding;
    binding.vk        = vk;
    binding.modifiers = modifiers;
    binding.commandId = std::wstring(commandId);
    dest.push_back(std::move(binding));
}

void AddPositionBinding(std::vector<Common::Settings::ShortcutBinding>& dest,
                        Common::Keyboard::KeyPosition keyPosition,
                        uint32_t modifiers,
                        std::wstring_view commandId)
{
    Common::Settings::ShortcutBinding binding;
    binding.modifiers   = modifiers;
    binding.commandId   = std::wstring(commandId);
    binding.keyPosition = keyPosition;
    dest.push_back(std::move(binding));
}

void MigrateDeprecatedShortcutCommandIds(std::vector<Common::Settings::ShortcutBinding>& bindings) noexcept
{
    for (Common::Settings::ShortcutBinding& binding : bindings)
    {
        if (binding.commandId == L"cmd/pane/permanentDeleteWithValidation")
        {
            binding.commandId = L"cmd/pane/permanentDelete";
        }
    }
}

[[nodiscard]] bool SameChord(const Common::Settings::ShortcutBinding& left, const Common::Settings::ShortcutBinding& right) noexcept
{
    return left.vk == right.vk && left.keyPosition == right.keyPosition && (left.modifiers & 0x7u) == (right.modifiers & 0x7u);
}

[[nodiscard]] Common::Settings::ShortcutBinding* FindBindingByChord(std::vector<Common::Settings::ShortcutBinding>& bindings,
                                                                    const Common::Settings::ShortcutBinding& candidate) noexcept
{
    for (Common::Settings::ShortcutBinding& binding : bindings)
    {
        if (SameChord(binding, candidate))
        {
            return &binding;
        }
    }
    return nullptr;
}

[[nodiscard]] Common::Settings::ShortcutBinding* FindBindingByChord(std::vector<Common::Settings::ShortcutBinding>& bindings,
                                                                    uint32_t vk,
                                                                    uint32_t modifiers) noexcept
{
    Common::Settings::ShortcutBinding candidate;
    candidate.vk        = vk;
    candidate.modifiers = modifiers;
    return FindBindingByChord(bindings, candidate);
}

[[nodiscard]] bool IsDefaultBindingIn(const std::vector<Common::Settings::ShortcutBinding>& defaults, const Common::Settings::ShortcutBinding& binding) noexcept
{
    for (const Common::Settings::ShortcutBinding& defaultBinding : defaults)
    {
        if (SameChord(binding, defaultBinding) && binding.commandId == defaultBinding.commandId)
        {
            return true;
        }
    }
    return false;
}

void RestoreMissingDefaultBindings(std::vector<Common::Settings::ShortcutBinding>& bindings, const std::vector<Common::Settings::ShortcutBinding>& defaults)
{
    for (const Common::Settings::ShortcutBinding& defaultBinding : defaults)
    {
        Common::Settings::ShortcutBinding* existing = FindBindingByChord(bindings, defaultBinding);
        if (! existing)
        {
            bindings.push_back(defaultBinding);
            continue;
        }

        if (existing->commandId.empty())
        {
            existing->commandId = defaultBinding.commandId;
        }
    }
}

using NormalizedBinding = std::tuple<uint32_t, uint32_t, uint32_t, std::wstring>;

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

        result.emplace_back(static_cast<uint32_t>(binding.keyPosition), binding.vk, binding.modifiers & 0x7u, binding.commandId);
    }

    std::sort(result.begin(), result.end());
    return result;
}

[[nodiscard]] bool GridLayoutEqual(const std::vector<Common::Settings::GridColumnLayoutEntry>& lhs,
                                   const std::vector<Common::Settings::GridColumnLayoutEntry>& rhs)
{
    if (lhs.size() != rhs.size())
    {
        return false;
    }

    for (size_t index = 0u; index < lhs.size(); ++index)
    {
        if (lhs[index].columnId != rhs[index].columnId || lhs[index].displayIndex != rhs[index].displayIndex || lhs[index].widthDip != rhs[index].widthDip)
        {
            return false;
        }
    }

    return true;
}
} // namespace

Common::Settings::ShortcutsSettings ShortcutDefaults::CreateDefaultShortcuts()
{
    Common::Settings::ShortcutsSettings shortcuts;
    shortcuts.migrationVersion = kCurrentShortcutMigrationVersion;

    // Application bindings are global and may be overridden by a context scope.
    AddBinding(shortcuts.application, VK_OEM_COMMA, ShortcutManager::kModCtrl, L"cmd/app/preferences");
    AddBinding(shortcuts.application, VK_OEM_COMMA, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/app/openSettingsFile");
    AddBinding(shortcuts.application, static_cast<uint32_t>('P'), ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/app/commandPalette");
    AddBinding(shortcuts.application,
               static_cast<uint32_t>('J'),
               ShortcutManager::kModCtrl | ShortcutManager::kModShift,
               L"cmd/app/showFileOperations");

    // Function bar bindings (F1..F12).
    AddBinding(shortcuts.functionBar, VK_F1, 0, L"cmd/app/showShortcuts");
    AddBinding(shortcuts.functionBar, VK_F1, ShortcutManager::kModAlt, L"cmd/app/openLeftDriveMenu");

    AddBinding(shortcuts.functionBar, VK_F2, 0, L"cmd/pane/rename");
    AddBinding(shortcuts.functionBar, VK_F2, ShortcutManager::kModCtrl, L"cmd/pane/sort/none");
    AddBinding(shortcuts.functionBar, VK_F2, ShortcutManager::kModAlt, L"cmd/app/openRightDriveMenu");

    AddBinding(shortcuts.functionBar, VK_F3, 0, L"cmd/pane/view");
    AddBinding(shortcuts.functionBar, VK_F3, ShortcutManager::kModCtrl, L"cmd/pane/sort/name");
    AddBinding(shortcuts.functionBar, VK_F3, ShortcutManager::kModAlt, L"cmd/pane/alternateView");
    AddBinding(shortcuts.functionBar, VK_F3, ShortcutManager::kModShift, L"cmd/pane/openCurrentFolder");
    AddBinding(shortcuts.functionBar, VK_F3, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/app/viewWidth");

    AddBinding(shortcuts.functionBar, VK_F4, 0, L"cmd/pane/edit");
    AddBinding(shortcuts.functionBar, VK_F4, ShortcutManager::kModCtrl, L"cmd/pane/sort/extension");
    AddBinding(shortcuts.functionBar, VK_F4, ShortcutManager::kModAlt, L"cmd/app/exit");
    AddBinding(shortcuts.functionBar, VK_F4, ShortcutManager::kModShift, L"cmd/pane/editNew");
    AddBinding(shortcuts.functionBar, VK_F4, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/alternateEdit");

    AddBinding(shortcuts.functionBar, VK_F5, 0, L"cmd/pane/copyToOtherPane");
    AddBinding(shortcuts.functionBar, VK_F5, ShortcutManager::kModShift, L"cmd/pane/copyToOtherPaneWithOptions");
    AddBinding(shortcuts.functionBar, VK_F5, ShortcutManager::kModCtrl, L"cmd/pane/sort/time");
    AddBinding(shortcuts.functionBar, VK_F5, ShortcutManager::kModAlt, L"cmd/pane/pack");
    AddBinding(shortcuts.functionBar, VK_F5, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/selection/save");

    AddBinding(shortcuts.functionBar, VK_F6, 0, L"cmd/pane/moveToOtherPane");
    AddBinding(shortcuts.functionBar, VK_F6, ShortcutManager::kModShift, L"cmd/pane/moveToOtherPaneWithOptions");
    AddBinding(shortcuts.functionBar, VK_F6, ShortcutManager::kModCtrl, L"cmd/pane/sort/size");
    AddBinding(shortcuts.functionBar, VK_F6, ShortcutManager::kModAlt, L"cmd/pane/unpack");
    AddBinding(shortcuts.functionBar, VK_F6, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/selection/restore");

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
    AddBinding(shortcuts.functionBar, VK_F10, ShortcutManager::kModAlt | ShortcutManager::kModShift, L"cmd/pane/contextMenuCurrentDirectory");

    AddBinding(shortcuts.functionBar, VK_F11, 0, L"cmd/pane/connect");
    AddBinding(shortcuts.functionBar, VK_F11, ShortcutManager::kModCtrl, L"cmd/pane/zoomPanel");
    AddBinding(shortcuts.functionBar, VK_F11, ShortcutManager::kModAlt, L"cmd/pane/listOpenedFiles");
    AddBinding(shortcuts.functionBar, VK_F11, ShortcutManager::kModShift, L"cmd/app/theme/selectPrev");
    AddBinding(shortcuts.functionBar, VK_F11, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/app/fullScreen");

    AddBinding(shortcuts.functionBar, VK_F12, 0, L"cmd/pane/disconnect");
    AddBinding(shortcuts.functionBar, VK_F12, ShortcutManager::kModCtrl, L"cmd/pane/filter");
    AddBinding(shortcuts.functionBar, VK_F12, ShortcutManager::kModAlt, L"cmd/pane/showFoldersHistory");
    AddBinding(shortcuts.functionBar, VK_F12, ShortcutManager::kModShift, L"cmd/app/theme/selectNext");

    // FolderView bindings (non-function-bar).
    AddBinding(shortcuts.folderView, VK_BACK, 0, L"cmd/pane/upOneDirectory");
    AddBinding(shortcuts.folderView, VK_BACK, ShortcutManager::kModShift, L"cmd/pane/goRootDirectory");
    AddBinding(shortcuts.folderView, VK_TAB, 0, L"cmd/pane/switchPaneFocus");
    AddBinding(shortcuts.folderView, VK_TAB, ShortcutManager::kModShift, L"cmd/pane/switchPaneFocus");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('U'), ShortcutManager::kModCtrl, L"cmd/app/swapPanes");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('2'), ShortcutManager::kModAlt, L"cmd/pane/display/brief");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('3'), ShortcutManager::kModAlt, L"cmd/pane/display/detailed");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('4'), ShortcutManager::kModAlt, L"cmd/pane/display/extraDetailed");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('5'), ShortcutManager::kModAlt, L"cmd/pane/viewOptions/toggleThumbnails");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('6'), ShortcutManager::kModAlt, L"cmd/pane/viewOptions/togglePreviewPane");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('7'), ShortcutManager::kModAlt, L"cmd/pane/openCommandShell");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('A'), ShortcutManager::kModCtrl, L"cmd/pane/selection/selectAll");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('F'), ShortcutManager::kModCtrl, L"cmd/pane/find");
    AddBinding(shortcuts.folderView, VK_ESCAPE, 0, L"cmd/pane/selection/unselectAll");
    AddPositionBinding(shortcuts.folderView,
                       Common::Keyboard::KeyPosition::NumberRowPlus,
                       ShortcutManager::kModCtrl,
                       L"cmd/pane/selection/selectDialog");
    AddPositionBinding(shortcuts.folderView,
                       Common::Keyboard::KeyPosition::NumberRowMinus,
                       ShortcutManager::kModCtrl,
                       L"cmd/pane/selection/unselectDialog");
    AddPositionBinding(shortcuts.folderView,
                       Common::Keyboard::KeyPosition::NumberRowPlus,
                       ShortcutManager::kModCtrl | ShortcutManager::kModShift,
                       L"cmd/pane/selection/selectSameExtension");
    AddPositionBinding(shortcuts.folderView,
                       Common::Keyboard::KeyPosition::NumberRowMinus,
                       ShortcutManager::kModCtrl | ShortcutManager::kModShift,
                       L"cmd/pane/selection/unselectSameExtension");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('C'), ShortcutManager::kModCtrl, L"cmd/pane/clipboardCopy");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('X'), ShortcutManager::kModCtrl, L"cmd/pane/clipboardCut");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('V'), ShortcutManager::kModCtrl, L"cmd/pane/clipboardPaste");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('L'), ShortcutManager::kModCtrl, L"cmd/pane/focusAddressBar");
    AddBinding(shortcuts.folderView, VK_OEM_PERIOD, ShortcutManager::kModCtrl, L"cmd/pane/setPathFromOtherPane");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('J'), ShortcutManager::kModCtrl, L"cmd/app/toggleFileOperationsFailedItems");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('D'), ShortcutManager::kModAlt, L"cmd/pane/focusAddressBar");
    AddBinding(shortcuts.folderView, VK_DOWN, ShortcutManager::kModAlt, L"cmd/pane/selection/goToNextSelectedName");
    AddBinding(shortcuts.folderView, VK_UP, ShortcutManager::kModAlt, L"cmd/pane/selection/goToPreviousSelectedName");
    AddBinding(shortcuts.folderView, VK_LEFT, ShortcutManager::kModAlt, L"cmd/pane/historyBack");
    AddBinding(shortcuts.folderView, VK_RIGHT, ShortcutManager::kModAlt, L"cmd/pane/historyForward");
    AddBinding(shortcuts.folderView,
               static_cast<uint32_t>('T'),
               ShortcutManager::kModCtrl | ShortcutManager::kModAlt,
               L"cmd/terminal/openFloatingWindow");
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
        const wchar_t digit             = (i < 9) ? static_cast<wchar_t>(L'1' + i) : L'0';
        const std::wstring goCommandId  = std::wstring(L"cmd/pane/hotPath/") + digit;
        const std::wstring setCommandId = std::wstring(L"cmd/pane/setHotPath/") + digit;
        AddBinding(shortcuts.folderView, static_cast<uint32_t>(digit), ShortcutManager::kModCtrl, goCommandId);
        AddBinding(shortcuts.folderView, static_cast<uint32_t>(digit), ShortcutManager::kModCtrl | ShortcutManager::kModShift, setCommandId);
    }

    AddBinding(shortcuts.folderView, VK_RETURN, 0, L"cmd/pane/executeOpen");
    AddBinding(shortcuts.folderView, VK_RETURN, ShortcutManager::kModCtrl, L"cmd/pane/bringFilenameToCommandLine");
    AddBinding(shortcuts.folderView, VK_RETURN, ShortcutManager::kModAlt, L"cmd/pane/openProperties");
    AddBinding(shortcuts.folderView, VK_RETURN, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/bringFullPathToTerminal");

    AddBinding(shortcuts.folderView, VK_SPACE, 0, L"cmd/pane/selectCalculateDirectorySizeNext");
    AddBinding(shortcuts.folderView, VK_SPACE, ShortcutManager::kModCtrl, L"cmd/pane/bringCurrentDirToCommandLine");
    AddBinding(shortcuts.folderView, VK_SPACE, ShortcutManager::kModAlt, L"cmd/pane/windowMenu");
    AddBinding(shortcuts.folderView, VK_SPACE, ShortcutManager::kModShift, L"cmd/pane/quickSearch");
    AddBinding(shortcuts.folderView, VK_SPACE, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/bringCurrentDirToCommandLine");

    AddBinding(shortcuts.folderView, VK_INSERT, 0, L"cmd/pane/selectNext");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModCtrl, L"cmd/pane/clipboardCopy");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModAlt, L"cmd/pane/copyPathAndNameAsText");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModShift, L"cmd/pane/clipboardPaste");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/copyUncPathAndNameAsText");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModCtrl | ShortcutManager::kModAlt, L"cmd/pane/copyPathAsText");
    AddBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModAlt | ShortcutManager::kModShift, L"cmd/pane/copyNameAsText");

    AddBinding(shortcuts.folderView, VK_DELETE, 0, L"cmd/pane/moveToRecycleBin");
    AddBinding(shortcuts.folderView, VK_DELETE, ShortcutManager::kModShift, L"cmd/pane/permanentDelete");
    AddBinding(shortcuts.folderView, VK_DELETE, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/permanentDelete");

    AddBinding(shortcuts.folderView, static_cast<uint32_t>('T'), ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/openCommandShell");
    AddBinding(shortcuts.folderView, VK_TAB, ShortcutManager::kModCtrl, L"cmd/pane/contentTab/next");
    AddBinding(shortcuts.folderView, VK_TAB, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/pane/contentTab/previous");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('1'), ShortcutManager::kModCtrl | ShortcutManager::kModAlt, L"cmd/pane/contentTab/select/1");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('2'), ShortcutManager::kModCtrl | ShortcutManager::kModAlt, L"cmd/pane/contentTab/select/2");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('3'), ShortcutManager::kModCtrl | ShortcutManager::kModAlt, L"cmd/pane/contentTab/select/3");
    AddBinding(shortcuts.folderView, static_cast<uint32_t>('9'), ShortcutManager::kModCtrl | ShortcutManager::kModAlt, L"cmd/pane/contentTab/last");

    // Terminal scope contains the 46 P1 bindings from the reviewed Windows Terminal merge.
    AddBinding(shortcuts.terminal, VK_F4, ShortcutManager::kModAlt, L"cmd/app/exit");
    AddBinding(shortcuts.terminal, VK_RETURN, ShortcutManager::kModAlt, L"cmd/app/fullScreen");
    AddBinding(shortcuts.terminal, VK_SPACE, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/sessionMenu");
    AddBinding(shortcuts.terminal, static_cast<uint32_t>('F'), ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/find");
    AddBinding(shortcuts.terminal, VK_SPACE, ShortcutManager::kModAlt, L"cmd/app/systemMenu");
    AddBinding(shortcuts.terminal, VK_OEM_PERIOD, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/suggestions");
    AddBinding(shortcuts.terminal, static_cast<uint32_t>('T'), ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/tab/new");
    AddBinding(shortcuts.terminal, static_cast<uint32_t>('N'), ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/openFloatingWindow");
    AddBinding(shortcuts.terminal, VK_TAB, ShortcutManager::kModCtrl, L"cmd/terminal/tab/next");
    AddBinding(shortcuts.terminal, VK_TAB, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/tab/previous");
    for (uint32_t tabIndex = 1u; tabIndex <= 8u; ++tabIndex)
    {
        AddBinding(shortcuts.terminal,
                   static_cast<uint32_t>('0') + tabIndex,
                   ShortcutManager::kModCtrl | ShortcutManager::kModAlt,
                   std::format(L"cmd/terminal/tab/select/{}", tabIndex));
    }
    AddBinding(shortcuts.terminal, static_cast<uint32_t>('9'), ShortcutManager::kModCtrl | ShortcutManager::kModAlt, L"cmd/terminal/tab/last");
    AddBinding(shortcuts.terminal, static_cast<uint32_t>('W'), ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/close");
    AddBinding(shortcuts.terminal, VK_LEFT, ShortcutManager::kModAlt | ShortcutManager::kModShift, L"cmd/pane/resizeSplitter/left");
    AddBinding(shortcuts.terminal, VK_RIGHT, ShortcutManager::kModAlt | ShortcutManager::kModShift, L"cmd/pane/resizeSplitter/right");
    AddBinding(shortcuts.terminal, VK_LEFT, ShortcutManager::kModAlt, L"cmd/pane/focus/left");
    AddBinding(shortcuts.terminal, VK_RIGHT, ShortcutManager::kModAlt, L"cmd/pane/focus/right");
    AddBinding(shortcuts.terminal, VK_LEFT, ShortcutManager::kModCtrl | ShortcutManager::kModAlt, L"cmd/pane/switchPaneFocus");
    AddBinding(shortcuts.terminal, static_cast<uint32_t>('C'), ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/copy");
    AddBinding(shortcuts.terminal, VK_INSERT, ShortcutManager::kModCtrl, L"cmd/terminal/copy");
    AddBinding(shortcuts.terminal, VK_RETURN, 0u, L"cmd/terminal/copySelectionOrPassthrough");
    AddBinding(shortcuts.terminal, static_cast<uint32_t>('C'), ShortcutManager::kModCtrl, L"cmd/terminal/copySelectionOrPassthrough");
    AddBinding(shortcuts.terminal, static_cast<uint32_t>('V'), ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/paste");
    AddBinding(shortcuts.terminal, VK_INSERT, ShortcutManager::kModShift, L"cmd/terminal/paste");
    AddBinding(shortcuts.terminal, static_cast<uint32_t>('A'), ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/selectAll");
    AddBinding(shortcuts.terminal, VK_APPS, 0u, L"cmd/terminal/contextMenu");
    AddBinding(shortcuts.terminal, VK_F10, ShortcutManager::kModShift, L"cmd/terminal/contextMenu");
    AddBinding(shortcuts.terminal, VK_DOWN, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/scroll/lineDown");
    AddBinding(shortcuts.terminal, VK_NEXT, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/scroll/pageDown");
    AddBinding(shortcuts.terminal, VK_UP, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/scroll/lineUp");
    AddBinding(shortcuts.terminal, VK_PRIOR, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/scroll/pageUp");
    AddBinding(shortcuts.terminal, VK_HOME, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/scroll/top");
    AddBinding(shortcuts.terminal, VK_END, ShortcutManager::kModCtrl | ShortcutManager::kModShift, L"cmd/terminal/scroll/bottom");
    AddPositionBinding(shortcuts.terminal,
                       Common::Keyboard::KeyPosition::NumberRowPlus,
                       ShortcutManager::kModCtrl,
                       L"cmd/terminal/font/increase");
    AddPositionBinding(shortcuts.terminal,
                       Common::Keyboard::KeyPosition::NumberRowMinus,
                       ShortcutManager::kModCtrl,
                       L"cmd/terminal/font/decrease");
    AddBinding(shortcuts.terminal, VK_ADD, ShortcutManager::kModCtrl, L"cmd/terminal/font/increase");
    AddBinding(shortcuts.terminal, VK_SUBTRACT, ShortcutManager::kModCtrl, L"cmd/terminal/font/decrease");
    AddBinding(shortcuts.terminal, static_cast<uint32_t>('0'), ShortcutManager::kModCtrl, L"cmd/terminal/font/reset");
    AddBinding(shortcuts.terminal, VK_NUMPAD0, ShortcutManager::kModCtrl, L"cmd/terminal/font/reset");

    return shortcuts;
}

bool ShortcutDefaults::AreShortcutsDefault(const Common::Settings::ShortcutsSettings& shortcuts)
{
    const Common::Settings::ShortcutsSettings defaults = CreateDefaultShortcuts();
    return NormalizeBindings(shortcuts.application) == NormalizeBindings(defaults.application) &&
           NormalizeBindings(shortcuts.functionBar) == NormalizeBindings(defaults.functionBar) &&
           NormalizeBindings(shortcuts.folderView) == NormalizeBindings(defaults.folderView) &&
           NormalizeBindings(shortcuts.terminal) == NormalizeBindings(defaults.terminal) &&
           shortcuts.applicationCollapsed == defaults.applicationCollapsed && shortcuts.functionBarCollapsed == defaults.functionBarCollapsed &&
           shortcuts.folderViewCollapsed == defaults.folderViewCollapsed && shortcuts.terminalCollapsed == defaults.terminalCollapsed &&
           shortcuts.migrationVersion == defaults.migrationVersion &&
           shortcuts.sortColumnId == defaults.sortColumnId && shortcuts.sortDescending == defaults.sortDescending &&
           GridLayoutEqual(shortcuts.gridLayout, defaults.gridLayout);
}

bool ShortcutDefaults::IsDefaultFunctionBarBinding(const Common::Settings::ShortcutBinding& binding)
{
    const Common::Settings::ShortcutsSettings defaults = CreateDefaultShortcuts();
    return IsDefaultBindingIn(defaults.functionBar, binding);
}

bool ShortcutDefaults::IsDefaultFolderViewBinding(const Common::Settings::ShortcutBinding& binding)
{
    const Common::Settings::ShortcutsSettings defaults = CreateDefaultShortcuts();
    return IsDefaultBindingIn(defaults.folderView, binding);
}

bool ShortcutDefaults::IsDefaultApplicationBinding(const Common::Settings::ShortcutBinding& binding)
{
    const Common::Settings::ShortcutsSettings defaults = CreateDefaultShortcuts();
    return IsDefaultBindingIn(defaults.application, binding);
}

bool ShortcutDefaults::IsDefaultTerminalBinding(const Common::Settings::ShortcutBinding& binding)
{
    const Common::Settings::ShortcutsSettings defaults = CreateDefaultShortcuts();
    return IsDefaultBindingIn(defaults.terminal, binding);
}

void ShortcutDefaults::EnsureShortcutsInitialized(Common::Settings::Settings& settings)
{
    if (! settings.shortcuts.has_value())
    {
        settings.shortcuts = CreateDefaultShortcuts();
        return;
    }

    Common::Settings::ShortcutsSettings& shortcuts = settings.shortcuts.value();
    MigrateDeprecatedShortcutCommandIds(shortcuts.application);
    MigrateDeprecatedShortcutCommandIds(shortcuts.functionBar);
    MigrateDeprecatedShortcutCommandIds(shortcuts.folderView);
    MigrateDeprecatedShortcutCommandIds(shortcuts.terminal);

    Common::Settings::ShortcutBinding* ctrlF2 = FindBindingByChord(shortcuts.functionBar, VK_F2, ShortcutManager::kModCtrl);
    if (! ctrlF2)
    {
        AddBinding(shortcuts.functionBar, VK_F2, ShortcutManager::kModCtrl, L"cmd/pane/sort/none");
    }
    else if (ctrlF2->commandId == L"cmd/pane/changeAttributes")
    {
        ctrlF2->commandId = L"cmd/pane/sort/none";
    }

    if (shortcuts.migrationVersion < kCurrentShortcutMigrationVersion)
    {
        Common::Settings::ShortcutBinding* ctrlAltT =
            FindBindingByChord(shortcuts.folderView, static_cast<uint32_t>('T'), ShortcutManager::kModCtrl | ShortcutManager::kModAlt);
        if (ctrlAltT && ctrlAltT->commandId == L"cmd/pane/openCommandShell")
        {
            ctrlAltT->commandId = L"cmd/terminal/openFloatingWindow";
        }
        shortcuts.migrationVersion = kCurrentShortcutMigrationVersion;
    }

    const Common::Settings::ShortcutsSettings defaults = CreateDefaultShortcuts();
    RestoreMissingDefaultBindings(shortcuts.application, defaults.application);
    RestoreMissingDefaultBindings(shortcuts.functionBar, defaults.functionBar);
    RestoreMissingDefaultBindings(shortcuts.folderView, defaults.folderView);
    RestoreMissingDefaultBindings(shortcuts.terminal, defaults.terminal);
}
