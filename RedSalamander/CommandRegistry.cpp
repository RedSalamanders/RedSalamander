#include "CommandRegistry.h"

#include <algorithm>
#include <array>

#include "resource.h"

namespace
{
constexpr std::array<CommandInfo, 125> kCommands = {
    CommandInfo{L"cmd/app/about", IDS_CMD_ABOUT, IDS_CMD_DESC_ABOUT, IDM_ABOUT},
    CommandInfo{L"cmd/app/compare", IDS_CMD_COMPARE, IDS_CMD_DESC_COMPARE, IDM_APP_COMPARE},
    CommandInfo{L"cmd/app/exit", IDS_CMD_EXIT, IDS_CMD_DESC_EXIT, IDM_EXIT},
    CommandInfo{L"cmd/app/fullScreen", IDS_CMD_FULL_SCREEN, IDS_CMD_DESC_FULL_SCREEN, IDM_APP_FULL_SCREEN},
    CommandInfo{L"cmd/app/openFileExplorerKnownFolder", IDS_CMD_OPEN_FILE_EXPLORER_KNOWN_FOLDER, IDS_CMD_DESC_OPEN_FILE_EXPLORER_KNOWN_FOLDER, 0},
    CommandInfo{L"cmd/app/openLeftDriveMenu", IDS_CMD_OPEN_LEFT_DRIVE_MENU, IDS_CMD_DESC_OPEN_LEFT_DRIVE_MENU, IDM_LEFT_CHANGE_DRIVE},
    CommandInfo{L"cmd/app/openRightDriveMenu", IDS_CMD_OPEN_RIGHT_DRIVE_MENU, IDS_CMD_DESC_OPEN_RIGHT_DRIVE_MENU, IDM_RIGHT_CHANGE_DRIVE},
    CommandInfo{L"cmd/app/plugins/configure", IDS_CMD_PLUGINS_CONFIGURE, IDS_CMD_DESC_PLUGINS_CONFIGURE, 0},
    CommandInfo{L"cmd/app/plugins/manage", IDS_CMD_PLUGINS_MANAGE, IDS_CMD_DESC_PLUGINS_MANAGE, IDM_VIEW_PLUGINS_MANAGE},
    CommandInfo{L"cmd/app/plugins/toggleEnabled", IDS_CMD_PLUGINS_TOGGLE_ENABLED, IDS_CMD_DESC_PLUGINS_TOGGLE_ENABLED, 0},
    CommandInfo{L"cmd/app/preferences", IDS_CMD_PREFERENCES, IDS_CMD_DESC_PREFERENCES, IDM_FILE_PREFERENCES},
    CommandInfo{L"cmd/app/rereadAssociations", IDS_CMD_REREAD_ASSOCIATIONS, IDS_CMD_DESC_REREAD_ASSOCIATIONS, IDM_APP_REREAD_ASSOCIATIONS},
    CommandInfo{L"cmd/app/showShortcuts", IDS_CMD_SHORTCUTS, IDS_CMD_DESC_SHORTCUTS, IDM_APP_SHOW_SHORTCUTS},
    CommandInfo{L"cmd/app/swapPanes", IDS_CMD_SWAP_PANES, IDS_CMD_DESC_SWAP_PANES, IDM_APP_SWAP_PANES},
    CommandInfo{L"cmd/app/theme/select", IDS_CMD_THEME_SELECT, IDS_CMD_DESC_THEME_SELECT, 0},
    CommandInfo{
        L"cmd/app/theme/systemHighContrastIndicator", IDS_CMD_THEME_SYSTEM_HIGH_CONTRAST_INDICATOR, IDS_CMD_DESC_THEME_SYSTEM_HIGH_CONTRAST_INDICATOR, 0},
    CommandInfo{L"cmd/app/toggleFileOperationsFailedItems",
                IDS_CMD_TOGGLE_FILEOPS_FAILED_ITEMS,
                IDS_CMD_DESC_TOGGLE_FILEOPS_FAILED_ITEMS,
                IDM_VIEW_FILEOPS_FAILED_ITEMS},
    CommandInfo{L"cmd/app/toggleFunctionBar", IDS_CMD_TOGGLE_FUNCTION_BAR, IDS_CMD_DESC_TOGGLE_FUNCTION_BAR, IDM_VIEW_FUNCTIONBAR},
    CommandInfo{L"cmd/app/toggleMenuBar", IDS_CMD_TOGGLE_MENU_BAR, IDS_CMD_DESC_TOGGLE_MENU_BAR, IDM_VIEW_MENUBAR},
    CommandInfo{L"cmd/app/viewWidth", IDS_CMD_VIEW_WIDTH, IDS_CMD_DESC_VIEW_WIDTH, IDM_APP_VIEW_WIDTH},
    CommandInfo{L"cmd/pane/alternateView", IDS_CMD_ALTERNATE_VIEW, IDS_CMD_DESC_ALTERNATE_VIEW, IDM_PANE_ALTERNATE_VIEW},
    CommandInfo{L"cmd/pane/bringCurrentDirToCommandLine",
                IDS_CMD_BRING_CURRENT_DIR_TO_COMMAND_LINE,
                IDS_CMD_DESC_BRING_CURRENT_DIR_TO_COMMAND_LINE,
                IDM_PANE_BRING_CURRENT_DIR_TO_COMMAND_LINE},
    CommandInfo{L"cmd/pane/bringFilenameToCommandLine",
                IDS_CMD_BRING_FILENAME_TO_COMMAND_LINE,
                IDS_CMD_DESC_BRING_FILENAME_TO_COMMAND_LINE,
                IDM_PANE_BRING_FILENAME_TO_COMMAND_LINE},
    CommandInfo{
        L"cmd/pane/calculateDirectorySizes", IDS_CMD_CALCULATE_DIRECTORY_SIZES, IDS_CMD_DESC_CALCULATE_DIRECTORY_SIZES, IDM_PANE_CALCULATE_DIRECTORY_SIZES},
    CommandInfo{L"cmd/pane/changeAttributes", IDS_CMD_CHANGE_ATTRIBUTES, IDS_CMD_DESC_CHANGE_ATTRIBUTES, IDM_PANE_CHANGE_ATTRIBUTES},
    CommandInfo{L"cmd/pane/changeCase", IDS_CMD_CHANGE_CASE, IDS_CMD_DESC_CHANGE_CASE, IDM_PANE_CHANGE_CASE},
    CommandInfo{L"cmd/pane/changeDirectory", IDS_CMD_CHANGE_DIRECTORY, IDS_CMD_DESC_CHANGE_DIRECTORY, IDM_PANE_CHANGE_DIRECTORY},
    CommandInfo{L"cmd/pane/clipboardCopy", IDS_CMD_CLIPBOARD_COPY, IDS_CMD_DESC_CLIPBOARD_COPY, IDM_PANE_CLIPBOARD_COPY},
    CommandInfo{L"cmd/pane/clipboardCut", IDS_CMD_CLIPBOARD_CUT, IDS_CMD_DESC_CLIPBOARD_CUT, IDM_PANE_CLIPBOARD_CUT},
    CommandInfo{L"cmd/pane/clipboardPaste", IDS_CMD_CLIPBOARD_PASTE, IDS_CMD_DESC_CLIPBOARD_PASTE, IDM_PANE_CLIPBOARD_PASTE},
    CommandInfo{L"cmd/pane/clipboardPasteShortcut", IDS_CMD_CLIPBOARD_PASTE_SHORTCUT, IDS_CMD_DESC_CLIPBOARD_PASTE_SHORTCUT, IDM_PANE_CLIPBOARD_PASTE_SHORTCUT},
    CommandInfo{L"cmd/pane/connect", IDS_CMD_CONNECT, IDS_CMD_DESC_CONNECT, IDM_PANE_CONNECT},
    CommandInfo{L"cmd/pane/connections", IDS_CMD_CONNECTIONS, IDS_CMD_DESC_CONNECTIONS, IDM_PANE_CONNECTION_MANAGER},
    CommandInfo{L"cmd/pane/contextMenu", IDS_CMD_CONTEXT_MENU, IDS_CMD_DESC_CONTEXT_MENU, IDM_PANE_CONTEXT_MENU},
    CommandInfo{L"cmd/pane/contextMenuCurrentDirectory",
                IDS_CMD_CONTEXT_MENU_CURRENT_DIRECTORY,
                IDS_CMD_DESC_CONTEXT_MENU_CURRENT_DIRECTORY,
                IDM_PANE_CONTEXT_MENU_CURRENT_DIRECTORY},
    CommandInfo{L"cmd/pane/copyNameAsText", IDS_CMD_COPY_NAME_AS_TEXT, IDS_CMD_DESC_COPY_NAME_AS_TEXT, IDM_PANE_COPY_NAME_AS_TEXT},
    CommandInfo{L"cmd/pane/copyPathAndFileName", IDS_CMD_COPY_PATH_AND_FILE_NAME, IDS_CMD_DESC_COPY_PATH_AND_FILE_NAME, IDM_PANE_COPY_PATH_AND_FILE_NAME},
    CommandInfo{
        L"cmd/pane/copyPathAndNameAsText", IDS_CMD_COPY_PATH_AND_NAME_AS_TEXT, IDS_CMD_DESC_COPY_PATH_AND_NAME_AS_TEXT, IDM_PANE_COPY_PATH_AND_NAME_AS_TEXT},
    CommandInfo{L"cmd/pane/copyPathAsText", IDS_CMD_COPY_PATH_AS_TEXT, IDS_CMD_DESC_COPY_PATH_AS_TEXT, IDM_PANE_COPY_PATH_AS_TEXT},
    CommandInfo{L"cmd/pane/copyToOtherPane", IDS_CMD_COPY, IDS_CMD_DESC_COPY, IDM_PANE_COPY_TO_OTHER},
    CommandInfo{L"cmd/pane/createDirectory", IDS_CMD_MAKE_DIRECTORY, IDS_CMD_DESC_MAKE_DIRECTORY, IDM_PANE_CREATE_DIR},
    CommandInfo{L"cmd/pane/delete", IDS_CMD_DELETE, IDS_CMD_DESC_DELETE, IDM_PANE_DELETE},
    CommandInfo{L"cmd/pane/disconnect", IDS_CMD_DISCONNECT, IDS_CMD_DESC_DISCONNECT, IDM_PANE_DISCONNECT},
    CommandInfo{L"cmd/pane/display/brief", IDS_CMD_DISPLAY_BRIEF, IDS_CMD_DESC_DISPLAY_BRIEF, IDM_PANE_DISPLAY_BRIEF},
    CommandInfo{L"cmd/pane/display/detailed", IDS_CMD_DISPLAY_DETAILED, IDS_CMD_DESC_DISPLAY_DETAILED, IDM_PANE_DISPLAY_DETAILED},
    CommandInfo{L"cmd/pane/display/extraDetailed", IDS_CMD_DISPLAY_EXTRA_DETAILED, IDS_CMD_DESC_DISPLAY_EXTRA_DETAILED, IDM_PANE_DISPLAY_EXTRA_DETAILED},
    CommandInfo{L"cmd/pane/edit", IDS_CMD_EDIT, IDS_CMD_DESC_EDIT, IDM_PANE_EDIT},
    CommandInfo{L"cmd/pane/editNew", IDS_CMD_EDIT_NEW, IDS_CMD_DESC_EDIT_NEW, IDM_PANE_EDIT_NEW},
    CommandInfo{L"cmd/pane/editWidth", IDS_CMD_EDIT_WIDTH, IDS_CMD_DESC_EDIT_WIDTH, IDM_PANE_EDIT_WIDTH},
    CommandInfo{L"cmd/pane/editWith", IDS_CMD_EDIT_WITH, IDS_CMD_DESC_EDIT_WITH, 0},
    CommandInfo{L"cmd/pane/executeOpen", IDS_CMD_EXECUTE_OPEN, IDS_CMD_DESC_EXECUTE_OPEN, IDM_PANE_EXECUTE_OPEN},
    CommandInfo{L"cmd/pane/filter", IDS_CMD_FILTER, IDS_CMD_DESC_FILTER, 0},
    CommandInfo{L"cmd/pane/find", IDS_CMD_FIND, IDS_CMD_DESC_FIND, IDM_PANE_FIND},
    CommandInfo{L"cmd/pane/focusAddressBar", IDS_CMD_FOCUS_ADDRESS_BAR, IDS_CMD_DESC_FOCUS_ADDRESS_BAR, 0},
    CommandInfo{L"cmd/pane/goDriveRoot", IDS_CMD_GO_DRIVE_ROOT, IDS_CMD_DESC_GO_DRIVE_ROOT, 0},
    CommandInfo{L"cmd/pane/goRootDirectory", IDS_CMD_GO_ROOT_DIRECTORY, IDS_CMD_DESC_GO_ROOT_DIRECTORY, 0},
    CommandInfo{L"cmd/pane/goToShortcutOrLinkTarget",
                IDS_CMD_GO_TO_SHORTCUT_OR_LINK_TARGET,
                IDS_CMD_DESC_GO_TO_SHORTCUT_OR_LINK_TARGET,
                IDM_PANE_GO_TO_SHORTCUT_OR_LINK_TARGET},
    CommandInfo{L"cmd/pane/historyBack", IDS_CMD_HISTORY_BACK, IDS_CMD_DESC_HISTORY_BACK, 0},
    CommandInfo{L"cmd/pane/historyForward", IDS_CMD_HISTORY_FORWARD, IDS_CMD_DESC_HISTORY_FORWARD, 0},
    CommandInfo{L"cmd/pane/hotPath", IDS_CMD_HOT_PATH_GO, IDS_CMD_DESC_HOT_PATH_GO, 0},
    CommandInfo{L"cmd/pane/hotPaths", IDS_CMD_HOT_PATHS, IDS_CMD_DESC_HOT_PATHS, 0},
    CommandInfo{L"cmd/pane/listOpenedFiles", IDS_CMD_LIST_OPENED_FILES, IDS_CMD_DESC_LIST_OPENED_FILES, IDM_PANE_LIST_OPENED_FILES},
    CommandInfo{L"cmd/pane/loadSelection", IDS_CMD_LOAD_SELECTION, IDS_CMD_DESC_LOAD_SELECTION, IDM_PANE_LOAD_SELECTION},
    CommandInfo{L"cmd/pane/makeFileList", IDS_CMD_MAKE_FILE_LIST, IDS_CMD_DESC_MAKE_FILE_LIST, IDM_PANE_MAKE_FILE_LIST},
    CommandInfo{L"cmd/pane/menu", IDS_CMD_MENU, IDS_CMD_DESC_MENU, IDM_PANE_MENU},
    CommandInfo{L"cmd/pane/moveToOtherPane", IDS_CMD_MOVE, IDS_CMD_DESC_MOVE, IDM_PANE_MOVE_TO_OTHER},
    CommandInfo{L"cmd/pane/moveToRecycleBin", IDS_CMD_MOVE_TO_RECYCLE_BIN, IDS_CMD_DESC_MOVE_TO_RECYCLE_BIN, IDM_PANE_MOVE_TO_RECYCLE_BIN},
    CommandInfo{L"cmd/pane/navigatePath", IDS_CMD_NAVIGATE_PATH, IDS_CMD_DESC_NAVIGATE_PATH, 0},
    CommandInfo{L"cmd/pane/newFromShellTemplate", IDS_CMD_NEW_FROM_TEMPLATE, IDS_CMD_DESC_NEW_FROM_TEMPLATE, 0},
    CommandInfo{L"cmd/pane/openCommandShell", IDS_CMD_OPEN_COMMAND_SHELL, IDS_CMD_DESC_OPEN_COMMAND_SHELL, IDM_PANE_OPEN_COMMAND_SHELL},
    CommandInfo{L"cmd/pane/openCurrentFolder", IDS_CMD_OPEN_CURRENT_FOLDER, IDS_CMD_DESC_OPEN_CURRENT_FOLDER, IDM_PANE_OPEN_CURRENT_FOLDER},
    CommandInfo{L"cmd/pane/openProperties", IDS_CMD_OPEN_PROPERTIES, IDS_CMD_DESC_OPEN_PROPERTIES, IDM_PANE_OPEN_PROPERTIES},
    CommandInfo{L"cmd/pane/openSecurity", IDS_CMD_OPEN_SECURITY, IDS_CMD_DESC_OPEN_SECURITY, IDM_PANE_OPEN_SECURITY},
    CommandInfo{L"cmd/pane/pack", IDS_CMD_PACK, IDS_CMD_DESC_PACK, IDM_PANE_PACK},
    CommandInfo{L"cmd/pane/permanentDelete", IDS_CMD_PERMANENT_DELETE, IDS_CMD_DESC_PERMANENT_DELETE, IDM_PANE_PERMANENT_DELETE},
    CommandInfo{L"cmd/pane/permanentDeleteWithValidation",
                IDS_CMD_PERMANENT_DELETE_WITH_VALIDATION,
                IDS_CMD_DESC_PERMANENT_DELETE_WITH_VALIDATION,
                IDM_PANE_PERMANENT_DELETE_WITH_VALIDATION},
    CommandInfo{L"cmd/pane/quickSearch", IDS_CMD_QUICK_SEARCH, IDS_CMD_DESC_QUICK_SEARCH, IDM_PANE_QUICK_SEARCH},
    CommandInfo{L"cmd/pane/refresh", IDS_CMD_REFRESH, IDS_CMD_DESC_REFRESH, 0},
    CommandInfo{L"cmd/pane/rename", IDS_CMD_RENAME, IDS_CMD_DESC_RENAME, IDM_PANE_RENAME},
    CommandInfo{L"cmd/pane/saveSelection", IDS_CMD_SAVE_SELECTION, IDS_CMD_DESC_SAVE_SELECTION, IDM_PANE_SAVE_SELECTION},
    CommandInfo{L"cmd/pane/selectCalculateDirectorySizeNext",
                IDS_CMD_SELECT_CALC_DIR_SIZE_NEXT,
                IDS_CMD_DESC_SELECT_CALC_DIR_SIZE_NEXT,
                IDM_PANE_SELECT_CALC_DIR_SIZE_NEXT},
    CommandInfo{L"cmd/pane/selectFileSystemPlugin", IDS_CMD_SELECT_FILE_SYSTEM_PLUGIN, IDS_CMD_DESC_SELECT_FILE_SYSTEM_PLUGIN, 0},
    CommandInfo{L"cmd/pane/selectNext", IDS_CMD_SELECT_NEXT, IDS_CMD_DESC_SELECT_NEXT, IDM_PANE_SELECT_NEXT},
    CommandInfo{L"cmd/pane/selection/goToNextSelectedName",
                IDS_CMD_SELECTION_GOTO_NEXT_SELECTED_NAME,
                IDS_CMD_DESC_SELECTION_GOTO_NEXT_SELECTED_NAME,
                IDM_PANE_SELECTION_GOTO_NEXT_SELECTED_NAME},
    CommandInfo{L"cmd/pane/selection/goToPreviousSelectedName",
                IDS_CMD_SELECTION_GOTO_PREV_SELECTED_NAME,
                IDS_CMD_DESC_SELECTION_GOTO_PREV_SELECTED_NAME,
                IDM_PANE_SELECTION_GOTO_PREV_SELECTED_NAME},
    CommandInfo{L"cmd/pane/selection/hideSelectedNames",
                IDS_CMD_SELECTION_HIDE_SELECTED_NAMES,
                IDS_CMD_DESC_SELECTION_HIDE_SELECTED_NAMES,
                IDM_PANE_SELECTION_HIDE_SELECTED_NAMES},
    CommandInfo{L"cmd/pane/selection/hideUnselectedNames",
                IDS_CMD_SELECTION_HIDE_UNSELECTED_NAMES,
                IDS_CMD_DESC_SELECTION_HIDE_UNSELECTED_NAMES,
                IDM_PANE_SELECTION_HIDE_UNSELECTED_NAMES},
    CommandInfo{L"cmd/pane/selection/invert", IDS_CMD_SELECTION_INVERT, IDS_CMD_DESC_SELECTION_INVERT, IDM_PANE_SELECTION_INVERT},
    CommandInfo{L"cmd/pane/selection/restore", IDS_CMD_SELECTION_RESTORE, IDS_CMD_DESC_SELECTION_RESTORE, IDM_PANE_SELECTION_RESTORE},
    CommandInfo{L"cmd/pane/selection/selectAll", IDS_CMD_SELECTION_SELECT_ALL, IDS_CMD_DESC_SELECTION_SELECT_ALL, IDM_PANE_SELECTION_SELECT_ALL},
    CommandInfo{L"cmd/pane/selection/selectDialog", IDS_CMD_SELECTION_SELECT_DIALOG, IDS_CMD_DESC_SELECTION_SELECT_DIALOG, IDM_PANE_SELECTION_SELECT_DIALOG},
    CommandInfo{L"cmd/pane/selection/selectSameExtension",
                IDS_CMD_SELECTION_SELECT_SAME_EXTENSION,
                IDS_CMD_DESC_SELECTION_SELECT_SAME_EXTENSION,
                IDM_PANE_SELECTION_SELECT_SAME_EXTENSION},
    CommandInfo{
        L"cmd/pane/selection/selectSameName", IDS_CMD_SELECTION_SELECT_SAME_NAME, IDS_CMD_DESC_SELECTION_SELECT_SAME_NAME, IDM_PANE_SELECTION_SELECT_SAME_NAME},
    CommandInfo{L"cmd/pane/selection/showHiddenNames",
                IDS_CMD_SELECTION_SHOW_HIDDEN_NAMES,
                IDS_CMD_DESC_SELECTION_SHOW_HIDDEN_NAMES,
                IDM_PANE_SELECTION_SHOW_HIDDEN_NAMES},
    CommandInfo{L"cmd/pane/selection/unselectAll", IDS_CMD_SELECTION_UNSELECT_ALL, IDS_CMD_DESC_SELECTION_UNSELECT_ALL, IDM_PANE_SELECTION_UNSELECT_ALL},
    CommandInfo{
        L"cmd/pane/selection/unselectDialog", IDS_CMD_SELECTION_UNSELECT_DIALOG, IDS_CMD_DESC_SELECTION_UNSELECT_DIALOG, IDM_PANE_SELECTION_UNSELECT_DIALOG},
    CommandInfo{L"cmd/pane/selection/unselectSameExtension",
                IDS_CMD_SELECTION_UNSELECT_SAME_EXTENSION,
                IDS_CMD_DESC_SELECTION_UNSELECT_SAME_EXTENSION,
                IDM_PANE_SELECTION_UNSELECT_SAME_EXTENSION},
    CommandInfo{L"cmd/pane/selection/unselectSameName",
                IDS_CMD_SELECTION_UNSELECT_SAME_NAME,
                IDS_CMD_DESC_SELECTION_UNSELECT_SAME_NAME,
                IDM_PANE_SELECTION_UNSELECT_SAME_NAME},
    CommandInfo{L"cmd/pane/setHotPath", IDS_CMD_SET_HOT_PATH, IDS_CMD_DESC_SET_HOT_PATH, 0},
    CommandInfo{L"cmd/pane/setPathFromOtherPane", IDS_CMD_SET_PATH_FROM_OTHER_PANE, IDS_CMD_DESC_SET_PATH_FROM_OTHER_PANE, 0},
    CommandInfo{L"cmd/pane/shares", IDS_CMD_SHARES, IDS_CMD_DESC_SHARES, IDM_PANE_SHARES},
    CommandInfo{L"cmd/pane/showFoldersHistory", IDS_CMD_SHOW_FOLDERS_HISTORY, IDS_CMD_DESC_SHOW_FOLDERS_HISTORY, IDM_PANE_SHOW_FOLDERS_HISTORY},
    CommandInfo{L"cmd/pane/sort/attributes", IDS_CMD_SORT_BY_ATTRIBUTES, IDS_CMD_DESC_SORT_BY_ATTRIBUTES, IDM_PANE_SORT_ATTRIBUTES},
    CommandInfo{L"cmd/pane/sort/extension", IDS_CMD_SORT_BY_EXTENSION, IDS_CMD_DESC_SORT_BY_EXTENSION, IDM_PANE_SORT_EXTENSION},
    CommandInfo{L"cmd/pane/sort/name", IDS_CMD_SORT_BY_NAME, IDS_CMD_DESC_SORT_BY_NAME, IDM_PANE_SORT_NAME},
    CommandInfo{L"cmd/pane/sort/none", IDS_CMD_SORT_NONE, IDS_CMD_DESC_SORT_NONE, IDM_PANE_SORT_NONE},
    CommandInfo{L"cmd/pane/sort/size", IDS_CMD_SORT_BY_SIZE, IDS_CMD_DESC_SORT_BY_SIZE, IDM_PANE_SORT_SIZE},
    CommandInfo{L"cmd/pane/sort/time", IDS_CMD_SORT_BY_TIME, IDS_CMD_DESC_SORT_BY_TIME, IDM_PANE_SORT_TIME},
    CommandInfo{L"cmd/pane/switchPaneFocus", IDS_CMD_SWITCH_PANE_FOCUS, IDS_CMD_DESC_SWITCH_PANE_FOCUS, IDM_VIEW_SWITCH_PANE_FOCUS},
    CommandInfo{L"cmd/pane/unpack", IDS_CMD_UNPACK, IDS_CMD_DESC_UNPACK, IDM_PANE_UNPACK},
    CommandInfo{L"cmd/pane/upOneDirectory", IDS_CMD_UP_ONE_DIRECTORY, IDS_CMD_DESC_UP_ONE_DIRECTORY, 0},
    CommandInfo{L"cmd/pane/userMenu", IDS_CMD_USER_MENU, IDS_CMD_DESC_USER_MENU, IDM_PANE_USER_MENU},
    CommandInfo{L"cmd/pane/view", IDS_CMD_VIEW, IDS_CMD_DESC_VIEW, IDM_PANE_VIEW},
    CommandInfo{L"cmd/pane/viewOptions/toggleFileExtensions",
                IDS_CMD_VIEWOPTIONS_TOGGLE_FILE_EXTENSIONS,
                IDS_CMD_DESC_VIEWOPTIONS_TOGGLE_FILE_EXTENSIONS,
                IDM_VIEW_PANE_FILE_EXTENSIONS},
    CommandInfo{
        L"cmd/pane/viewOptions/toggleFilterBar", IDS_CMD_VIEWOPTIONS_TOGGLE_FILTER_BAR, IDS_CMD_DESC_VIEWOPTIONS_TOGGLE_FILTER_BAR, IDM_VIEW_PANE_FILTER_BAR},
    CommandInfo{L"cmd/pane/viewOptions/toggleHiddenFiles",
                IDS_CMD_VIEWOPTIONS_TOGGLE_HIDDEN_FILES,
                IDS_CMD_DESC_VIEWOPTIONS_TOGGLE_HIDDEN_FILES,
                IDM_VIEW_PANE_HIDDEN_FILES},
    CommandInfo{L"cmd/pane/viewOptions/toggleNavigationBar", IDS_CMD_VIEWOPTIONS_TOGGLE_NAVIGATION_BAR, IDS_CMD_DESC_VIEWOPTIONS_TOGGLE_NAVIGATION_BAR, 0},
    CommandInfo{L"cmd/pane/viewOptions/togglePreviewPane",
                IDS_CMD_VIEWOPTIONS_TOGGLE_PREVIEW_PANE,
                IDS_CMD_DESC_VIEWOPTIONS_TOGGLE_PREVIEW_PANE,
                IDM_VIEW_PANE_PREVIEW_PANE},
    CommandInfo{L"cmd/pane/viewOptions/toggleStatusBar", IDS_CMD_VIEWOPTIONS_TOGGLE_STATUS_BAR, IDS_CMD_DESC_VIEWOPTIONS_TOGGLE_STATUS_BAR, 0},
    CommandInfo{L"cmd/pane/viewOptions/toggleSystemFiles",
                IDS_CMD_VIEWOPTIONS_TOGGLE_SYSTEM_FILES,
                IDS_CMD_DESC_VIEWOPTIONS_TOGGLE_SYSTEM_FILES,
                IDM_VIEW_PANE_SYSTEM_FILES},
    CommandInfo{
        L"cmd/pane/viewOptions/toggleThumbnails", IDS_CMD_VIEWOPTIONS_TOGGLE_THUMBNAILS, IDS_CMD_DESC_VIEWOPTIONS_TOGGLE_THUMBNAILS, IDM_VIEW_PANE_THUMBNAILS},
    CommandInfo{L"cmd/pane/viewSpace", IDS_CMD_SPACE_VIEW, IDS_CMD_DESC_SPACE_VIEW, IDM_PANE_VIEW_SPACE},
    CommandInfo{L"cmd/pane/viewWith", IDS_CMD_VIEW_WITH, IDS_CMD_DESC_VIEW_WITH, 0},
    CommandInfo{L"cmd/pane/windowMenu", IDS_CMD_WINDOW_MENU, IDS_CMD_DESC_WINDOW_MENU, IDM_VIEW_WINDOW_MENU},
    CommandInfo{L"cmd/pane/zoomPanel", IDS_CMD_ZOOM_PANEL, IDS_CMD_DESC_ZOOM_PANEL, 0},
};

static_assert(
    []() consteval
    {
        for (size_t i = 1; i < kCommands.size(); ++i)
        {
            if (kCommands[i - 1].id >= kCommands[i].id)
            {
                return false;
            }
        }
        return true;
    }(),
    "kCommands must be sorted by id.");
} // namespace

std::wstring_view CanonicalizeCommandId(std::wstring_view commandId) noexcept
{
    constexpr std::wstring_view kGoDriveRootPrefix = L"cmd/pane/goDriveRoot/";
    if (commandId.starts_with(kGoDriveRootPrefix))
    {
        return L"cmd/pane/goDriveRoot";
    }

    constexpr std::wstring_view kHotPathPrefix = L"cmd/pane/hotPath/";
    if (commandId.starts_with(kHotPathPrefix))
    {
        return L"cmd/pane/hotPath";
    }

    constexpr std::wstring_view kSetHotPathPrefix = L"cmd/pane/setHotPath/";
    if (commandId.starts_with(kSetHotPathPrefix))
    {
        return L"cmd/pane/setHotPath";
    }

    return commandId;
}

const CommandInfo* FindCommandInfo(std::wstring_view commandId) noexcept
{
    commandId = CanonicalizeCommandId(commandId);
    const auto it =
        std::lower_bound(kCommands.begin(), kCommands.end(), commandId, [](const CommandInfo& entry, std::wstring_view id) noexcept { return entry.id < id; });
    if (it == kCommands.end() || it->id != commandId)
    {
        return nullptr;
    }

    return &(*it);
}

const CommandInfo* FindCommandInfoByWmCommandId(unsigned int wmCommandId) noexcept
{
    if (wmCommandId == 0)
    {
        return nullptr;
    }

    for (const auto& entry : kCommands)
    {
        if (entry.wmCommandId == wmCommandId)
        {
            return &entry;
        }
    }

    return nullptr;
}

std::span<const CommandInfo> GetAllCommands() noexcept
{
    return kCommands;
}

std::optional<unsigned int> TryGetWmCommandId(std::wstring_view commandId) noexcept
{
    const CommandInfo* info = FindCommandInfo(commandId);
    if (! info || info->wmCommandId == 0)
    {
        return std::nullopt;
    }
    return info->wmCommandId;
}

std::optional<unsigned int> TryGetCommandDisplayNameStringId(std::wstring_view commandId) noexcept
{
    const CommandInfo* info = FindCommandInfo(commandId);
    if (! info || info->displayNameStringId == 0)
    {
        return std::nullopt;
    }
    return info->displayNameStringId;
}

std::optional<unsigned int> TryGetCommandDescriptionStringId(std::wstring_view commandId) noexcept
{
    const CommandInfo* info = FindCommandInfo(commandId);
    if (! info || info->descriptionStringId == 0)
    {
        return std::nullopt;
    }
    return info->descriptionStringId;
}
