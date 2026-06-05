#include "NavigationViewInternal.h"

#include <array>
#include <chrono>

#include <windowsx.h>

#include <shellapi.h>

#include "ConnectionSecrets.h"
#include "DirectoryInfoCache.h"
#include "FileSystemPluginManager.h"
#include "FluentIcons.h"
#include "Helpers.h"
#include "IconCache.h"
#include "MaskSyntax.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "SettingsStore.h"
#include "ShortcutText.h"
#include "resource.h"

namespace
{
struct MenuGlyphTag
{
    wchar_t glyph = 0;
};

const MenuGlyphTag kMenuGlyphConnections{FluentIcons::kConnections};
constexpr int kLeftNavigationControlId  = 1001;
constexpr int kRightNavigationControlId = 1003;
constexpr float kEmbeddedDestinationDropdownMaxHeightDip = 420.0f;
constexpr GUID kKnownFolderIdOneDrive   = {
    0xA52BBA46,
    0xE9E1,
    0x435F,
    {0xB3, 0xD9, 0x28, 0xDA, 0xA6, 0x48, 0xC0, 0xF6},
};

bool IsFilePluginShortId(std::wstring_view pluginShortId) noexcept
{
    return pluginShortId.empty() || EqualsNoCase(pluginShortId, L"file");
}

[[nodiscard]] bool IsConnectionProtocolShortId(std::wstring_view pluginShortId) noexcept
{
    return EqualsNoCase(pluginShortId, L"ftp") || EqualsNoCase(pluginShortId, L"sftp") || EqualsNoCase(pluginShortId, L"scp") ||
           EqualsNoCase(pluginShortId, L"imap");
}

[[nodiscard]] bool LooksLikeDriveRootPath(const wchar_t* path) noexcept
{
    if (path == nullptr)
    {
        return false;
    }

    const wchar_t driveLetter = path[0];
    if (! ((driveLetter >= L'A' && driveLetter <= L'Z') || (driveLetter >= L'a' && driveLetter <= L'z')))
    {
        return false;
    }

    return path[1] == L':' && (path[2] == L'\\' || path[2] == L'/') && path[3] == L'\0';
}

void ApplyEmbeddedDestinationDropdownOptions(RedSalamander::DxUi::ContextMenuSessionCallbacks& sessionCallbacks, bool embeddedDestinationMode) noexcept
{
    if (! embeddedDestinationMode)
    {
        return;
    }

    sessionCallbacks.rootVerticalPlacement = RedSalamander::DxUi::ContextMenuRootVerticalPlacement::Above;
    sessionCallbacks.maxRootHeightDip      = kEmbeddedDestinationDropdownMaxHeightDip;
}

[[nodiscard]] LONG ResolveDropdownAnchorY(const RECT& bounds, bool embeddedDestinationMode) noexcept
{
    return embeddedDestinationMode ? bounds.top : bounds.bottom;
}

[[nodiscard]] POINT MakeDropdownAnchorPoint(const RECT& bounds, bool alignEnd, bool embeddedDestinationMode) noexcept
{
    return POINT{alignEnd ? bounds.right : bounds.left, ResolveDropdownAnchorY(bounds, embeddedDestinationMode)};
}

[[nodiscard]] LONG ResolveBreadcrumbDropdownAnchorY(const D2D1_RECT_F& bounds, LONG sectionTop, bool embeddedDestinationMode) noexcept
{
    const float anchorY = embeddedDestinationMode ? bounds.top : bounds.bottom;
    return static_cast<LONG>(std::lround(anchorY + static_cast<float>(sectionTop)));
}

[[nodiscard]] std::wstring CompactChordTextForMenu(std::wstring text)
{
    for (size_t pos = 0; (pos = text.find(L" + ", pos)) != std::wstring::npos;)
    {
        text.replace(pos, 3u, L"+");
        pos += 1u;
    }

    return text;
}

struct MenuPresentationText
{
    std::wstring label;
    std::wstring accelerator;
};

struct MenuInfoLineText
{
    std::wstring label;
    std::wstring value;
};

[[nodiscard]] MenuPresentationText DecodeMenuPresentationText(std::wstring_view rawText)
{
    const size_t tabPos                     = rawText.find(L'\t');
    const std::wstring_view labelText       = tabPos == std::wstring_view::npos ? rawText : rawText.substr(0u, tabPos);
    const std::wstring_view acceleratorText = tabPos == std::wstring_view::npos ? std::wstring_view{} : rawText.substr(tabPos + 1u);

    MenuPresentationText result{};
    result.label.reserve(labelText.size());
    for (size_t i = 0u; i < labelText.size(); ++i)
    {
        const wchar_t ch = labelText[i];
        if (ch == L'&')
        {
            if (i + 1u < labelText.size() && labelText[i + 1u] == L'&')
            {
                result.label.push_back(L'&');
                ++i;
            }
            continue;
        }

        result.label.push_back(ch);
    }

    result.label       = TrimWhitespace(result.label);
    result.accelerator = TrimWhitespace(std::wstring(acceleratorText));
    return result;
}

[[nodiscard]] MenuInfoLineText SplitFormattedMenuInfoLine(std::wstring text)
{
    const size_t colonPos = text.find(L':');
    if (colonPos == std::wstring::npos)
    {
        return MenuInfoLineText{.label = std::move(text)};
    }

    MenuInfoLineText result{};
    result.label = text.substr(0u, colonPos + 1u);

    size_t valueStart = colonPos + 1u;
    while (valueStart < text.size() && (text[valueStart] == L' ' || text[valueStart] == L'\t'))
    {
        ++valueStart;
    }

    result.value = TrimWhitespace(text.substr(valueStart));
    return result;
}

[[nodiscard]] std::optional<std::wstring> TryExtractDriveRootPrefix(std::wstring_view text) noexcept
{
    if (text.size() < 3u)
    {
        return std::nullopt;
    }

    const wchar_t driveLetter = text[0];
    if (! ((driveLetter >= L'A' && driveLetter <= L'Z') || (driveLetter >= L'a' && driveLetter <= L'z')))
    {
        return std::nullopt;
    }

    if (text[1] != L':' || (text[2] != L'\\' && text[2] != L'/'))
    {
        return std::nullopt;
    }

    std::wstring root;
    root.reserve(3u);
    root.push_back(driveLetter);
    root.push_back(L':');
    root.push_back(L'\\');
    return root;
}

[[nodiscard]] std::optional<std::wstring> TryResolveDriveRootIconPath(const NavigationMenuItem& item)
{
    const MenuPresentationText presentation = DecodeMenuPresentationText(item.label ? item.label : L"");
    if (const auto labelRoot = TryExtractDriveRootPrefix(presentation.label); labelRoot.has_value())
    {
        return labelRoot;
    }

    const auto tryCandidate = [](const wchar_t* candidate) -> std::optional<std::wstring>
    {
        if (! candidate || candidate[0] == L'\0' || ! LooksLikeDriveRootPath(candidate))
        {
            return std::nullopt;
        }

        return std::wstring(candidate);
    };

    if (const auto iconRoot = tryCandidate(item.iconPath); iconRoot.has_value())
    {
        return iconRoot;
    }
    if (const auto pathRoot = tryCandidate(item.path); pathRoot.has_value())
    {
        return pathRoot;
    }

    return std::nullopt;
}

[[nodiscard]] std::wstring ResolveNavigationMenuIconGlyph(const NavigationMenuItem& item)
{
    if ((item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0 || (item.flags & NAV_MENU_ITEM_FLAG_HEADER) != 0)
    {
        return {};
    }

    switch (item.commandId)
    {
        case DRIVE_INFO_COMMAND_PROPERTIES: return std::wstring(1u, FluentIcons::kInfo);
        case DRIVE_INFO_COMMAND_CLEANUP: return std::wstring(1u, FluentIcons::kClear);
        default: break;
    }

    const wchar_t* iconPath = (item.iconPath && item.iconPath[0] != L'\0') ? item.iconPath : item.path;
    if (iconPath && iconPath[0] != L'\0')
    {
        if (LooksLikeDriveRootPath(iconPath))
        {
            return std::wstring(1u, FluentIcons::kHardDrive);
        }

        if (iconPath[0] == L'\\' && iconPath[1] == L'\\')
        {
            return std::wstring(1u, FluentIcons::kConnections);
        }

        return std::wstring(1u, FluentIcons::kOpenFile);
    }

    return {};
}

[[nodiscard]] const GUID* TryGetKnownFolderIdForNavigationMenuItem(const NavigationMenuItem& item)
{
    const std::wstring decodedLabel = DecodeMenuPresentationText(item.label ? item.label : L"").label;
    const auto labelMatches         = [&](const UINT resourceId) -> bool
    {
        const std::wstring expected = LoadStringResource(nullptr, resourceId);
        return ! expected.empty() &&
               CompareStringOrdinal(decodedLabel.data(), static_cast<int>(decodedLabel.size()), expected.data(), static_cast<int>(expected.size()), TRUE) ==
                   CSTR_EQUAL;
    };

    if (labelMatches(IDS_MENU_NAV_DESKTOP))
    {
        return &FOLDERID_Desktop;
    }
    if (labelMatches(IDS_MENU_NAV_DOCUMENTS))
    {
        return &FOLDERID_Documents;
    }
    if (labelMatches(IDS_MENU_NAV_DOWNLOADS))
    {
        return &FOLDERID_Downloads;
    }
    if (labelMatches(IDS_MENU_NAV_PICTURES))
    {
        return &FOLDERID_Pictures;
    }
    if (labelMatches(IDS_MENU_NAV_MUSIC))
    {
        return &FOLDERID_Music;
    }
    if (labelMatches(IDS_MENU_NAV_VIDEOS))
    {
        return &FOLDERID_Videos;
    }
    if (labelMatches(IDS_MENU_NAV_ONEDRIVE))
    {
        return &kKnownFolderIdOneDrive;
    }

    return nullptr;
}

[[nodiscard]] std::shared_ptr<RedSalamander::DxUi::MenuFlyoutItem::BitmapIcon> ResolveNavigationMenuBitmapIcon(const NavigationMenuItem& item, int iconSizePx)
{
    if ((item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0 || (item.flags & NAV_MENU_ITEM_FLAG_HEADER) != 0 || iconSizePx <= 0)
    {
        return {};
    }

    switch (item.commandId)
    {
        case DRIVE_INFO_COMMAND_PROPERTIES:
        case DRIVE_INFO_COMMAND_CLEANUP: return {};
        default: break;
    }

    const auto wrapBitmap = [&](wil::unique_hbitmap bitmap)
    {
        if (! bitmap)
        {
            return std::shared_ptr<RedSalamander::DxUi::MenuFlyoutItem::BitmapIcon>{};
        }

        return std::make_shared<RedSalamander::DxUi::MenuFlyoutItem::BitmapIcon>(
            std::move(bitmap), static_cast<UINT>(iconSizePx), static_cast<UINT>(iconSizePx));
    };
    const auto tryBitmapFromPath = [&](const wchar_t* path, DWORD fileAttributes = 0, bool useFileAttributes = false)
    { return wrapBitmap(IconCache::GetInstance().CreateMenuBitmapFromPath(path, iconSizePx, fileAttributes, useFileAttributes)); };
    const auto tryBitmapFromIconIndex = [&](const int iconIndex)
    { return wrapBitmap(IconCache::GetInstance().CreateMenuBitmapFromIconIndex(iconIndex, iconSizePx)); };

    if (const GUID* const knownFolderId = TryGetKnownFolderIdForNavigationMenuItem(item))
    {
        if (auto bitmapIcon = wrapBitmap(IconCache::GetInstance().CreateMenuBitmapFromKnownFolder(*knownFolderId, iconSizePx)))
        {
            return bitmapIcon;
        }
        if (const auto iconIndex = IconCache::GetInstance().QuerySysIconIndexForKnownFolder(*knownFolderId); iconIndex.has_value())
        {
            if (auto bitmapIcon = tryBitmapFromIconIndex(iconIndex.value()))
            {
                return bitmapIcon;
            }
        }
    }

    if (const auto driveRootPath = TryResolveDriveRootIconPath(item); driveRootPath.has_value())
    {
        if (auto bitmapIcon = tryBitmapFromPath(driveRootPath->c_str()))
        {
            return bitmapIcon;
        }
        if (const auto iconIndex = IconCache::GetInstance().QuerySysIconIndexForPath(driveRootPath->c_str(), 0, false); iconIndex.has_value())
        {
            if (auto bitmapIcon = tryBitmapFromIconIndex(iconIndex.value()))
            {
                return bitmapIcon;
            }
        }
    }

    const wchar_t* iconPath = (item.iconPath && item.iconPath[0] != L'\0') ? item.iconPath : item.path;
    if (iconPath && iconPath[0] != L'\0')
    {
        if (auto bitmapIcon = tryBitmapFromPath(iconPath))
        {
            return bitmapIcon;
        }
        if (auto bitmapIcon = tryBitmapFromPath(iconPath, FILE_ATTRIBUTE_DIRECTORY, true))
        {
            return bitmapIcon;
        }
        if (const auto iconIndex = IconCache::GetInstance().QuerySysIconIndexForPath(iconPath, 0, false); iconIndex.has_value())
        {
            if (auto bitmapIcon = tryBitmapFromIconIndex(iconIndex.value()))
            {
                return bitmapIcon;
            }
        }
    }

    return {};
}

void ApplyNavigationMenuIcon(RedSalamander::DxUi::MenuFlyoutItem& item, const NavigationMenuItem& sourceItem, int iconSizePx)
{
    item.iconBitmap = ResolveNavigationMenuBitmapIcon(sourceItem, iconSizePx);
    if (! item.iconBitmap)
    {
        item.iconGlyph = ResolveNavigationMenuIconGlyph(sourceItem);
    }
}

[[nodiscard]] std::optional<bool> TryResolveLeftPaneNavigationView(HWND navigationViewHwnd) noexcept
{
    if (! navigationViewHwnd || IsWindow(navigationViewHwnd) == FALSE)
    {
        return std::nullopt;
    }

    const int controlId = GetDlgCtrlID(navigationViewHwnd);
    if (controlId == kLeftNavigationControlId)
    {
        return true;
    }
    if (controlId == kRightNavigationControlId)
    {
        return false;
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<std::wstring> TryGetShortcutTextForCommandId(const Common::Settings::Settings& settings, std::wstring_view commandId) noexcept;

void TrimTrailingMenuSeparators(std::vector<RedSalamander::DxUi::MenuFlyoutItem>& items) noexcept
{
    while (! items.empty() && items.back().kind == RedSalamander::DxUi::MenuItemKind::Separator)
    {
        items.pop_back();
    }
}

template <typename RegisterWindowCommandFn, typename RegisterNavigatePathFn>
[[nodiscard]] bool TryBuildPaneGoToFlyoutItem(HWND navigationViewHwnd,
                                              Common::Settings::Settings* settings,
                                              const std::deque<std::filesystem::path>& pathHistory,
                                              const std::optional<std::filesystem::path>& currentPath,
                                              UINT& nextId,
                                              UINT maxId,
                                              RegisterWindowCommandFn&& registerWindowCommand,
                                              RegisterNavigatePathFn&& registerNavigatePath,
                                              RedSalamander::DxUi::MenuFlyoutItem& outItem)
{
    const std::optional<bool> isLeftPaneOpt = TryResolveLeftPaneNavigationView(navigationViewHwnd);
    if (! isLeftPaneOpt.has_value())
    {
        return false;
    }

    outItem.text = LoadStringResource(nullptr, IDS_MENU_GO_TO);
    if (outItem.text.empty())
    {
        return false;
    }

    const auto appendSeparatorIfNeeded = [](std::vector<RedSalamander::DxUi::MenuFlyoutItem>& items) noexcept
    {
        if (items.empty() || items.back().kind == RedSalamander::DxUi::MenuItemKind::Separator)
        {
            return;
        }

        items.push_back(RedSalamander::DxUi::MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Separator});
    };
    const auto appendDisabledItem = [&](UINT labelResourceId, std::wstring_view fallback) noexcept
    {
        std::wstring label = LoadStringResource(nullptr, labelResourceId);
        if (label.empty())
        {
            label.assign(fallback);
        }
        outItem.children.push_back(RedSalamander::DxUi::MenuFlyoutItem{.text = std::wstring(label), .enabled = false});
    };
    const auto appendWindowCommandItem =
        [&](UINT labelResourceId, std::wstring_view fallbackLabel, UINT commandId, std::wstring_view shortcutCommandId) noexcept -> bool
    {
        if (nextId > maxId)
        {
            Debug::Warning(L"[NavigationView] Drive menu Go To submenu truncated (max actionable items reached)");
            return false;
        }

        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.text = LoadStringResource(nullptr, labelResourceId);
        if (item.text.empty())
        {
            item.text.assign(fallbackLabel);
        }
        if (settings)
        {
            if (const std::optional<std::wstring> shortcut = TryGetShortcutTextForCommandId(*settings, shortcutCommandId))
            {
                item.acceleratorText = shortcut.value();
            }
        }
        item.commandId = static_cast<int>(nextId++);
        registerWindowCommand(static_cast<UINT>(item.commandId), commandId);
        outItem.children.push_back(std::move(item));
        return true;
    };
    const auto appendNavigatePathItem = [&](std::wstring label, std::wstring accelerator, const std::filesystem::path& path, bool checked) noexcept -> bool
    {
        if (path.empty())
        {
            return true;
        }

        if (nextId > maxId)
        {
            Debug::Warning(L"[NavigationView] Drive menu Go To submenu truncated (max actionable items reached)");
            return false;
        }

        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.text            = std::move(label);
        item.acceleratorText = std::move(accelerator);
        item.commandId       = static_cast<int>(nextId++);
        item.checked         = checked;
        if (checked)
        {
            item.kind = RedSalamander::DxUi::MenuItemKind::Toggle;
        }

        registerNavigatePath(static_cast<UINT>(item.commandId), path);
        outItem.children.push_back(std::move(item));
        return true;
    };

    const bool isLeftPane = isLeftPaneOpt.value();
    struct GoToFixedItemSpec
    {
        UINT labelResourceId;
        std::wstring_view fallbackLabel;
        UINT commandId;
        std::wstring_view shortcutCommandId;
    };
    const std::array<GoToFixedItemSpec, 6> fixedItems = {{
        {IDS_MENU_GO_TO_BACK, L"Back", static_cast<UINT>(isLeftPane ? IDM_LEFT_GO_TO_BACK : IDM_RIGHT_GO_TO_BACK), L"cmd/pane/historyBack"},
        {IDS_MENU_GO_TO_FORWARD, L"Forward", static_cast<UINT>(isLeftPane ? IDM_LEFT_GO_TO_FORWARD : IDM_RIGHT_GO_TO_FORWARD), L"cmd/pane/historyForward"},
        {IDS_MENU_GO_TO_PARENT_DIRECTORY,
         L"Parent Directory",
         static_cast<UINT>(isLeftPane ? IDM_LEFT_GO_TO_PARENT_DIRECTORY : IDM_RIGHT_GO_TO_PARENT_DIRECTORY),
         L"cmd/pane/upOneDirectory"},
        {IDS_MENU_GO_TO_ROOT_DIRECTORY,
         L"Root Directory",
         static_cast<UINT>(isLeftPane ? IDM_LEFT_GO_TO_ROOT_DIRECTORY : IDM_RIGHT_GO_TO_ROOT_DIRECTORY),
         L"cmd/pane/goRootDirectory"},
        {IDS_MENU_GO_TO_PATH_FROM_OTHER_PANEL,
         L"Path from Other Panel",
         static_cast<UINT>(isLeftPane ? IDM_LEFT_GO_TO_PATH_FROM_OTHER_PANE : IDM_RIGHT_GO_TO_PATH_FROM_OTHER_PANE),
         L"cmd/pane/setPathFromOtherPane"},
        {IDS_MENU_GO_TO_HOT_PATHS, L"Hot Paths...", static_cast<UINT>(isLeftPane ? IDM_LEFT_HOT_PATHS : IDM_RIGHT_HOT_PATHS), L"cmd/pane/hotPaths"},
    }};

    for (size_t index = 0; index < fixedItems.size(); ++index)
    {
        if (index == 5u)
        {
            appendSeparatorIfNeeded(outItem.children);
        }

        if (! appendWindowCommandItem(
                fixedItems[index].labelResourceId, fixedItems[index].fallbackLabel, fixedItems[index].commandId, fixedItems[index].shortcutCommandId))
        {
            return ! outItem.children.empty();
        }
    }

    size_t hotPathCount = 0u;
    if (settings && settings->hotPaths.has_value())
    {
        const auto& slots = settings->hotPaths.value().slots;
        for (size_t i = 0; i < slots.size(); ++i)
        {
            if (! slots[i].has_value() || slots[i].value().path.empty())
            {
                continue;
            }

            const auto& slot        = slots[i].value();
            const wchar_t digitChar = (i < 9u) ? static_cast<wchar_t>(L'1' + i) : L'0';
            std::wstring label      = ! slot.label.empty() ? std::format(L"{}: {}", digitChar, slot.label) : std::format(L"{}: {}", digitChar, slot.path);

            std::wstring accelerator;
            std::wstring commandId = L"cmd/pane/hotPath/";
            commandId.push_back(digitChar);
            if (const std::optional<std::wstring> shortcut = TryGetShortcutTextForCommandId(*settings, commandId))
            {
                accelerator = shortcut.value();
            }

            if (! appendNavigatePathItem(std::move(label), std::move(accelerator), std::filesystem::path(slot.path), false))
            {
                return ! outItem.children.empty();
            }

            ++hotPathCount;
        }
    }
    if (hotPathCount == 0u)
    {
        appendDisabledItem(IDS_MENU_EMPTY, L"(Empty)");
    }

    appendSeparatorIfNeeded(outItem.children);

    size_t historyCount                = 0u;
    const std::wstring currentPathText = currentPath.has_value() ? currentPath.value().wstring() : std::wstring{};
    for (const auto& entry : pathHistory)
    {
        if (entry.empty())
        {
            continue;
        }

        const bool checked = ! currentPathText.empty() && EqualsNoCase(entry.wstring(), currentPathText);
        if (! appendNavigatePathItem(entry.wstring(), {}, entry, checked))
        {
            return ! outItem.children.empty();
        }

        ++historyCount;
    }
    if (historyCount == 0u)
    {
        appendDisabledItem(IDS_MENU_EMPTY, L"(Empty)");
    }

    TrimTrailingMenuSeparators(outItem.children);
    return ! outItem.children.empty();
}

[[nodiscard]] int CompareTextNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    if (left.empty() && right.empty())
    {
        return 0;
    }

    const int compare = CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE);
    if (compare == CSTR_LESS_THAN)
    {
        return -1;
    }
    if (compare == CSTR_GREATER_THAN)
    {
        return 1;
    }
    return 0;
}

[[nodiscard]] std::optional<std::wstring> TryGetShortcutTextForCommandId(const Common::Settings::Settings& settings, std::wstring_view commandId) noexcept
{
    if (commandId.empty())
    {
        return std::nullopt;
    }

    if (! settings.shortcuts.has_value())
    {
        return std::nullopt;
    }

    const auto findBinding = [&](const std::vector<Common::Settings::ShortcutBinding>& bindings) noexcept -> std::optional<std::wstring>
    {
        for (const auto& binding : bindings)
        {
            if (binding.commandId.empty())
            {
                continue;
            }

            if (std::wstring_view(binding.commandId) != commandId)
            {
                continue;
            }

            std::wstring text = ShortcutText::FormatChordText(binding.vk, binding.modifiers);
            if (text.empty())
            {
                return std::wstring{};
            }

            return CompactChordTextForMenu(std::move(text));
        }

        return std::nullopt;
    };

    const Common::Settings::ShortcutsSettings& shortcuts = settings.shortcuts.value();
    if (std::optional<std::wstring> found = findBinding(shortcuts.functionBar))
    {
        return found;
    }

    if (std::optional<std::wstring> found = findBinding(shortcuts.folderView))
    {
        return found;
    }

    return std::nullopt;
}

struct NavigationMenuSnapshot
{
    wil::com_ptr<INavigationMenu> menu;
    const NavigationMenuItem* items = nullptr;
    unsigned int count              = 0;
};

std::optional<NavigationMenuSnapshot> TryGetFileSystemNavigationMenuItems() noexcept
{
    FileSystemPluginManager& manager = FileSystemPluginManager::GetInstance();
    const auto& plugins              = manager.GetPlugins();

    for (const auto& entry : plugins)
    {
        if (entry.shortId.empty() || ! EqualsNoCase(entry.shortId, L"file"))
        {
            continue;
        }

        if (! entry.fileSystem)
        {
            continue;
        }

        wil::com_ptr<INavigationMenu> menu;
        const HRESULT qiHr = entry.fileSystem->QueryInterface(__uuidof(INavigationMenu), menu.put_void());
        if (FAILED(qiHr) || ! menu)
        {
            continue;
        }

        const NavigationMenuItem* items = nullptr;
        unsigned int count              = 0;
        const HRESULT hr                = menu->GetMenuItems(&items, &count);
        if (FAILED(hr) || ! items || count == 0)
        {
            continue;
        }

        NavigationMenuSnapshot snapshot;
        snapshot.menu  = std::move(menu);
        snapshot.items = items;
        snapshot.count = count;
        return snapshot;
    }

    return std::nullopt;
}
} // namespace

bool NavigationView::ExecuteNavigationMenuAction(UINT menuId)
{
    for (const auto& action : _navigationMenuActions)
    {
        if (action.menuId != menuId)
        {
            continue;
        }

        RequestOwnerPaneFocus();

        if (action.type == MenuActionType::NavigatePath)
        {
            RequestPathChange(std::filesystem::path(action.path));
            return true;
        }

        if (action.type == MenuActionType::WindowCommand)
        {
            const HWND rootOwner = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
            if (rootOwner)
            {
                SendMessageW(rootOwner, WM_COMMAND, MAKEWPARAM(action.commandId, 0), 0);
            }
            return true;
        }

        if (_navigationMenu)
        {
            static_cast<void>(_navigationMenu->ExecuteMenuCommand(action.commandId));
        }
        return true;
    }

    return false;
}

void NavigationView::OpenDriveMenuFromCommand()
{
    if (! _hWnd)
    {
        return;
    }

    if (IsFilePluginShortId(_pluginShortId) && _showMenuSection && _navigationMenu)
    {
        PostMessageW(_hWnd.get(), WndMsg::kNavigationViewShowMenuDropdown, 0, 0);
        return;
    }

    PostMessageW(_hWnd.get(), WndMsg::kNavigationViewShowDriveMenuDropdown, 0, 0);
}

bool NavigationView::ExecuteDriveMenuAction(UINT menuId)
{
    for (const auto& action : _driveMenuActions)
    {
        if (action.menuId != menuId)
        {
            continue;
        }

        RequestOwnerPaneFocus();

        if (action.type == MenuActionType::NavigatePath)
        {
            RequestPathChange(std::filesystem::path(action.path));
            return true;
        }

        if (action.type == MenuActionType::WindowCommand)
        {
            const HWND rootOwner = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
            if (rootOwner)
            {
                SendMessageW(rootOwner, WM_COMMAND, MAKEWPARAM(action.commandId, 0), 0);
            }
            return true;
        }

        if (_driveInfo && _currentPluginPath)
        {
            const std::wstring pathText = _currentPluginPath.value().wstring();
            static_cast<void>(_driveInfo->ExecuteDriveMenuCommand(action.commandId, pathText.c_str()));
        }

        return true;
    }

    return false;
}

void NavigationView::ShowMenuDropdown(bool ignoreInitialLeftButtonUp, bool focusFirstNavigableItem)
{
    TraceNavigationViewMenuDiagnostics(L"navigation.menu-dropdown.enter",
                                       L"hwnd={:#x} showMenu={} hasMenu={} ignoreInitialLeftButtonUp={} focusFirst={} embedded={} focus={:#x} active={:#x} capture={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       _showMenuSection ? 1 : 0,
                                       _navigationMenu ? 1 : 0,
                                       ignoreInitialLeftButtonUp ? 1 : 0,
                                       focusFirstNavigableItem ? 1 : 0,
                                       _embeddedDestinationMode ? 1 : 0,
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetCapture()));
    if (! _showMenuSection || ! _navigationMenu)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.menu-dropdown.skip",
                                           L"hwnd={:#x} reason=unavailable showMenu={} hasMenu={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           _showMenuSection ? 1 : 0,
                                           _navigationMenu ? 1 : 0);
        return;
    }

    _navDropdownKind = ModernDropdownKind::Menu;
    _navDropdownPaths.clear();
    _navDropdownSelectedIndex     = -1;
    const auto clearDropdownState = wil::scope_exit([&]() noexcept
    {
        _navDropdownKind = ModernDropdownKind::None;
        _navDropdownPaths.clear();
        _navDropdownSelectedIndex = -1;
    });

    const NavigationMenuItem* items = nullptr;
    unsigned int count              = 0;
    const HRESULT hr                = _navigationMenu->GetMenuItems(&items, &count);
    if (FAILED(hr) || ! items || count == 0)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.menu-dropdown.skip",
                                           L"hwnd={:#x} reason=no-menu-items hr={:#x} hasItems={} count={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           static_cast<unsigned int>(hr),
                                           items ? 1 : 0,
                                           count);
        return;
    }

    _menuButtonPressed = true;
    RenderDriveSection();

    const auto resetPressedState = wil::scope_exit([&]() noexcept
    {
        _menuButtonPressed = false;
        RenderDriveSection();
    });

    _menuBitmaps.clear();
    _navigationMenuActions.clear();
    const auto clearActions = wil::scope_exit([&]() noexcept { _navigationMenuActions.clear(); });

    constexpr unsigned int kMaxActions = ID_NAV_MENU_MAX - ID_NAV_MENU_BASE + 1u;
    UINT nextId                        = ID_NAV_MENU_BASE;
    std::vector<RedSalamander::DxUi::MenuFlyoutItem> popupItems;
    popupItems.reserve(count + 8u);

    const bool isFilePluginShortId         = IsFilePluginShortId(_pluginShortId);
    const auto getConnectionsManagerTarget = [&]() -> std::wstring
    {
        if (! isFilePluginShortId && IsConnectionProtocolShortId(_pluginShortId))
        {
            std::wstring target;
            target.reserve(_pluginShortId.size() + 1u);
            target.append(_pluginShortId);
            target.push_back(L':');
            return target;
        }

        return L"nav:";
    };

    bool goToItemAdded        = false;
    bool connectionsItemAdded = false;
    std::optional<size_t> goToItemIndex;
    const auto registerWindowCommandAction = [&](const UINT menuId, const UINT commandId) noexcept
    {
        MenuAction action;
        action.menuId    = menuId;
        action.type      = MenuActionType::WindowCommand;
        action.commandId = commandId;
        _navigationMenuActions.push_back(std::move(action));
    };
    const auto reserveTailActionIds = [&](const unsigned int startIndex, const size_t extraReserve) noexcept -> UINT
    {
        size_t remainingActionable = extraReserve;
        for (unsigned int j = startIndex; j < count; ++j)
        {
            const NavigationMenuItem& item = items[j];
            if ((item.flags & (NAV_MENU_ITEM_FLAG_SEPARATOR | NAV_MENU_ITEM_FLAG_HEADER)) != 0)
            {
                continue;
            }

            const bool hasPath    = item.path && item.path[0] != L'\0';
            const bool hasCommand = item.commandId != 0;
            if (hasPath || hasCommand)
            {
                ++remainingActionable;
            }
        }

        if (remainingActionable == 0u)
        {
            return ID_NAV_MENU_MAX;
        }

        if (remainingActionable >= static_cast<size_t>(ID_NAV_MENU_MAX - ID_NAV_MENU_BASE + 1u))
        {
            return nextId > ID_NAV_MENU_BASE ? static_cast<UINT>(nextId - 1u) : ID_NAV_MENU_BASE;
        }

        const UINT cappedMax = static_cast<UINT>(ID_NAV_MENU_MAX - remainingActionable);
        return cappedMax >= nextId ? cappedMax : static_cast<UINT>(nextId - 1u);
    };
    const auto tryAppendGoToMenu = [&](UINT maxIdForGoTo) noexcept
    {
        if (goToItemAdded || nextId > maxIdForGoTo)
        {
            return;
        }

        RedSalamander::DxUi::MenuFlyoutItem submenu{};
        const size_t actionCountBefore        = _navigationMenuActions.size();
        const UINT nextIdBefore               = nextId;
        const auto registerNavigatePathAction = [&](const UINT menuId, const std::filesystem::path& path) noexcept
        {
            MenuAction action;
            action.menuId = menuId;
            action.type   = MenuActionType::NavigatePath;
            action.path   = path.wstring();
            _navigationMenuActions.push_back(std::move(action));
        };

        if (! TryBuildPaneGoToFlyoutItem(
                _hWnd.get(), _settings, _pathHistory, _currentPath, nextId, maxIdForGoTo, registerWindowCommandAction, registerNavigatePathAction, submenu))
        {
            _navigationMenuActions.resize(actionCountBefore);
            nextId = nextIdBefore;
            return;
        }

        goToItemIndex = popupItems.size();
        popupItems.push_back(std::move(submenu));
        goToItemAdded = true;
    };
    const auto tryAppendConnectionsMenu = [&]() noexcept
    {
        if (connectionsItemAdded || nextId > ID_NAV_MENU_MAX)
        {
            return;
        }

        const std::wstring connectionsLabel = LoadStringResource(nullptr, IDS_MENU_CONNECTIONS);
        if (connectionsLabel.empty())
        {
            return;
        }

        struct ConnectionMenuItem
        {
            std::wstring label;
            std::wstring navName;
            std::wstring actionPath;
        };

        std::vector<ConnectionMenuItem> connectionItems;
        connectionItems.reserve(_settings && _settings->connections ? _settings->connections->items.size() + 1u : 2u);

        // Quick Connect (session-only)
        {
            ConnectionMenuItem item;
            item.navName = std::wstring(RedSalamander::Connections::kQuickConnectConnectionName);

            Common::Settings::ConnectionProfile quick{};
            RedSalamander::Connections::GetQuickConnectProfile(quick);

            if (! quick.host.empty())
            {
                item.label      = quick.port != 0u ? std::format(L"{}:{}", quick.host, quick.port) : quick.host;
                item.actionPath = std::format(L"nav:{}", item.navName);
            }
            else
            {
                item.label = LoadStringResource(nullptr, IDS_CONNECTIONS_QUICK_CONNECT);
                if (item.label.empty())
                {
                    item.label = L"<Quick Connect>";
                }
                item.actionPath = getConnectionsManagerTarget();
            }
            connectionItems.push_back(std::move(item));
        }

        // Persisted profiles
        if (_settings && _settings->connections)
        {
            for (const auto& profile : _settings->connections->items)
            {
                if (profile.name.empty() || profile.pluginId.empty())
                {
                    continue;
                }
                if (RedSalamander::Connections::IsQuickConnectConnectionName(profile.name))
                {
                    continue;
                }

                ConnectionMenuItem item;
                item.label      = profile.name;
                item.navName    = profile.name;
                item.actionPath = std::format(L"nav:{}", item.navName);
                connectionItems.push_back(std::move(item));
            }
        }

        if (connectionItems.size() > 1u)
        {
            std::sort(connectionItems.begin() + 1u,
                      connectionItems.end(),
                      [](const ConnectionMenuItem& a, const ConnectionMenuItem& b)
            {
                const int labelCompare = CompareTextNoCase(a.label, b.label);
                if (labelCompare != 0)
                {
                    return labelCompare < 0;
                }

                return CompareTextNoCase(a.navName, b.navName) < 0;
            });
        }

        std::vector<RedSalamander::DxUi::MenuFlyoutItem> connectionChildren;
        if (! connectionItems.empty())
        {
            const std::wstring managerLabel = LoadStringResource(nullptr, IDS_MENU_CONNECTIONS_ELLIPSIS);
            if (! managerLabel.empty() && nextId <= ID_NAV_MENU_MAX)
            {
                RedSalamander::DxUi::MenuFlyoutItem managerItem{};
                managerItem.text      = managerLabel;
                managerItem.commandId = static_cast<int>(nextId++);
                managerItem.iconGlyph.assign(1u, kMenuGlyphConnections.glyph);
                connectionChildren.push_back(std::move(managerItem));

                MenuAction action;
                action.menuId = static_cast<UINT>(connectionChildren.back().commandId);
                action.type   = MenuActionType::NavigatePath;
                action.path   = getConnectionsManagerTarget();
                _navigationMenuActions.push_back(std::move(action));
            }

            connectionChildren.push_back(RedSalamander::DxUi::MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Separator});
        }

        if (connectionItems.empty())
        {
            const std::wstring emptyLabel = LoadStringResource(nullptr, IDS_MENU_EMPTY);
            connectionChildren.push_back(
                RedSalamander::DxUi::MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Header, .text = emptyLabel.empty() ? L"(Empty)" : emptyLabel});
        }
        else
        {
            for (const auto& item : connectionItems)
            {
                if (nextId > ID_NAV_MENU_MAX)
                {
                    break;
                }

                RedSalamander::DxUi::MenuFlyoutItem child{};
                child.text      = item.label;
                child.commandId = static_cast<int>(nextId++);
                connectionChildren.push_back(std::move(child));

                MenuAction action;
                action.menuId = static_cast<UINT>(connectionChildren.back().commandId);
                action.type   = MenuActionType::NavigatePath;
                action.path   = item.actionPath;
                _navigationMenuActions.push_back(std::move(action));
            }
        }

        RedSalamander::DxUi::MenuFlyoutItem submenu{};
        submenu.text = connectionsLabel;
        submenu.iconGlyph.assign(1u, kMenuGlyphConnections.glyph);
        submenu.children = std::move(connectionChildren);
        popupItems.push_back(std::move(submenu));
        connectionsItemAdded = true;
    };

    const auto appendSeparatorIfNeeded = [&](std::vector<RedSalamander::DxUi::MenuFlyoutItem>& targetItems) noexcept
    {
        if (targetItems.empty() || targetItems.back().kind == RedSalamander::DxUi::MenuItemKind::Separator)
        {
            return;
        }

        targetItems.push_back(RedSalamander::DxUi::MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Separator});
    };

    const auto appendNavigationItem = [&](std::vector<RedSalamander::DxUi::MenuFlyoutItem>& targetItems,
                                          const NavigationMenuItem& item,
                                          UINT maxId,
                                          std::wstring_view truncatedLabel) noexcept -> bool
    {
        if ((item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0)
        {
            appendSeparatorIfNeeded(targetItems);
            return true;
        }

        const bool isHeader   = (item.flags & NAV_MENU_ITEM_FLAG_HEADER) != 0;
        const bool isDisabled = (item.flags & NAV_MENU_ITEM_FLAG_DISABLED) != 0;
        const bool hasPath    = item.path && item.path[0] != L'\0';
        const bool hasCommand = item.commandId != 0;
        const bool actionable = ! isHeader && (hasPath || hasCommand);

        if (actionable && nextId > maxId)
        {
            Debug::Warning(L"[NavigationView] {} truncated (max {} actionable items)", truncatedLabel, kMaxActions);
            return false;
        }

        const MenuPresentationText presentation = DecodeMenuPresentationText(item.label ? item.label : L"");
        RedSalamander::DxUi::MenuFlyoutItem dxItem{};
        dxItem.kind            = isHeader ? RedSalamander::DxUi::MenuItemKind::Header : RedSalamander::DxUi::MenuItemKind::Standard;
        dxItem.text            = presentation.label;
        dxItem.acceleratorText = presentation.accelerator;
        ApplyNavigationMenuIcon(dxItem, item, _menuIconSize);
        dxItem.enabled = ! isDisabled && ! isHeader;

        if (actionable)
        {
            dxItem.commandId = static_cast<int>(nextId++);
            MenuAction action;
            action.menuId = static_cast<UINT>(dxItem.commandId);
            if (hasPath)
            {
                action.type = MenuActionType::NavigatePath;
                action.path = item.path;
            }
            else
            {
                action.type      = MenuActionType::Command;
                action.commandId = item.commandId;
            }
            _navigationMenuActions.push_back(std::move(action));
        }

        targetItems.push_back(std::move(dxItem));
        return true;
    };

    const auto tryAppendCommonFoldersMenu = [&]() noexcept
    {
        if (isFilePluginShortId || nextId > ID_NAV_MENU_MAX)
        {
            return;
        }

        const std::wstring label = LoadStringResource(nullptr, IDS_MENU_COMMON_FOLDERS);
        if (label.empty())
        {
            return;
        }

        const std::optional<NavigationMenuSnapshot> fileMenuOpt = TryGetFileSystemNavigationMenuItems();
        if (! fileMenuOpt.has_value())
        {
            return;
        }

        const size_t actionCountBefore = _navigationMenuActions.size();
        const UINT nextIdBefore        = nextId;
        std::vector<RedSalamander::DxUi::MenuFlyoutItem> commonFolderChildren;
        commonFolderChildren.reserve(std::min<unsigned int>(fileMenuOpt.value().count, 8u));

        const unsigned int fileCount = fileMenuOpt.value().count;
        for (unsigned int i = 0; i < fileCount; ++i)
        {
            const NavigationMenuItem& item = fileMenuOpt.value().items[i];
            if ((item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0)
            {
                break;
            }

            if ((item.flags & NAV_MENU_ITEM_FLAG_HEADER) != 0 || item.path == nullptr || item.path[0] == L'\0')
            {
                continue;
            }

            if (! appendNavigationItem(commonFolderChildren, item, ID_NAV_MENU_MAX, L"Common folders submenu"))
            {
                break;
            }
        }

        if (commonFolderChildren.empty())
        {
            _navigationMenuActions.resize(actionCountBefore);
            nextId = nextIdBefore;
            return;
        }

        RedSalamander::DxUi::MenuFlyoutItem submenu{};
        submenu.text     = label;
        submenu.children = std::move(commonFolderChildren);
        popupItems.push_back(std::move(submenu));
    };

    for (unsigned int i = 0; i < count; ++i)
    {
        const NavigationMenuItem& item = items[i];
        const bool isSeparator         = (item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0;
        if (isSeparator)
        {
            if (isFilePluginShortId && ! goToItemAdded)
            {
                const NavigationMenuItem* nextNonSeparator = nullptr;
                for (unsigned int j = i + 1u; j < count; ++j)
                {
                    if ((items[j].flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0)
                    {
                        continue;
                    }

                    nextNonSeparator = &items[j];
                    break;
                }

                if (nextNonSeparator != nullptr && LooksLikeDriveRootPath(nextNonSeparator->path))
                {
                    tryAppendGoToMenu(reserveTailActionIds(i + 1u, 0u));
                }
            }

            appendSeparatorIfNeeded(popupItems);
            continue;
        }

        if (! appendNavigationItem(popupItems, item, ID_NAV_MENU_MAX, L"Navigation menu"))
        {
            break;
        }
    }

    if (isFilePluginShortId)
    {
        if (! goToItemAdded)
        {
            tryAppendGoToMenu(ID_NAV_MENU_MAX);
        }

        if (! connectionsItemAdded)
        {
            const size_t connectionsIndex = popupItems.size();
            tryAppendConnectionsMenu();
            if (connectionsItemAdded && goToItemIndex.has_value() && connectionsIndex + 1u == popupItems.size() && connectionsIndex > goToItemIndex.value())
            {
                RedSalamander::DxUi::MenuFlyoutItem connectionItem = std::move(popupItems.back());
                popupItems.pop_back();
                const size_t insertIndex = std::min(goToItemIndex.value(), popupItems.size());
                popupItems.insert(popupItems.begin() + static_cast<std::vector<RedSalamander::DxUi::MenuFlyoutItem>::difference_type>(insertIndex),
                                  std::move(connectionItem));
                goToItemIndex = insertIndex + 1u;
            }
        }
    }
    else
    {
        tryAppendCommonFoldersMenu();

        if (! goToItemAdded)
        {
            tryAppendGoToMenu(ID_NAV_MENU_MAX);
        }

        if (! connectionsItemAdded)
        {
            tryAppendConnectionsMenu();
        }
    }

    if (! IsFilePluginShortId(_pluginShortId))
    {
        const std::optional<NavigationMenuSnapshot> fileMenuOpt = TryGetFileSystemNavigationMenuItems();
        if (fileMenuOpt.has_value())
        {
            const std::wstring label = LoadStringResource(nullptr, IDS_MENU_CHANGE_DRIVE);
            if (! label.empty())
            {
                std::vector<RedSalamander::DxUi::MenuFlyoutItem> changeDriveChildren;
                changeDriveChildren.reserve(fileMenuOpt.value().count);
                UINT fileId                  = nextId;
                const unsigned int fileCount = fileMenuOpt.value().count;
                for (unsigned int i = 0; i < fileCount; ++i)
                {
                    const NavigationMenuItem& item = fileMenuOpt.value().items[i];
                    const bool isSeparator         = (item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0;
                    if (isSeparator)
                    {
                        appendSeparatorIfNeeded(changeDriveChildren);
                        continue;
                    }

                    const bool isHeader   = (item.flags & NAV_MENU_ITEM_FLAG_HEADER) != 0;
                    const bool isDisabled = (item.flags & NAV_MENU_ITEM_FLAG_DISABLED) != 0;
                    const bool hasPath    = item.path && item.path[0] != L'\0';
                    const bool hasCommand = item.commandId != 0;
                    const bool actionable = ! isHeader && (hasPath || hasCommand);

                    if (actionable && fileId > ID_NAV_MENU_MAX)
                    {
                        break;
                    }

                    const MenuPresentationText presentation = DecodeMenuPresentationText(item.label ? item.label : L"");
                    RedSalamander::DxUi::MenuFlyoutItem child{};
                    child.kind            = isHeader ? RedSalamander::DxUi::MenuItemKind::Header : RedSalamander::DxUi::MenuItemKind::Standard;
                    child.text            = presentation.label;
                    child.acceleratorText = presentation.accelerator;
                    ApplyNavigationMenuIcon(child, item, _menuIconSize);
                    child.enabled = ! isDisabled && ! isHeader;

                    if (actionable)
                    {
                        child.commandId = static_cast<int>(fileId++);
                        MenuAction action;
                        action.menuId = static_cast<UINT>(child.commandId);
                        if (hasPath)
                        {
                            action.type = MenuActionType::NavigatePath;
                            action.path = item.path;
                            _navigationMenuActions.push_back(std::move(action));
                        }
                    }

                    changeDriveChildren.push_back(std::move(child));
                }

                if (! changeDriveChildren.empty())
                {
                    appendSeparatorIfNeeded(popupItems);
                    RedSalamander::DxUi::MenuFlyoutItem submenu{};
                    submenu.text     = label;
                    submenu.children = std::move(changeDriveChildren);
                    popupItems.push_back(std::move(submenu));
                    nextId = fileId;
                }
            }
        }
    }

    if (popupItems.empty())
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.menu-dropdown.skip",
                                           L"hwnd={:#x} reason=no-popup-items sourceCount={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           count);
        return;
    }

    POINT pt = MakeDropdownAnchorPoint(_sectionDriveRect, false, _embeddedDestinationMode);
    ClientToScreen(_hWnd.get(), &pt);
    const HWND popupOwner = GetAncestor(_hWnd.get(), GA_ROOT);
    if (! popupOwner)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.menu-dropdown.skip",
                                           L"hwnd={:#x} reason=no-popup-owner",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()));
        return;
    }

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.ignoreInitialLeftButtonUp = ignoreInitialLeftButtonUp;
    sessionCallbacks.focusFirstNavigableItem   = focusFirstNavigableItem;
    ApplyEmbeddedDestinationDropdownOptions(sessionCallbacks, _embeddedDestinationMode);

    const auto startedAt  = std::chrono::steady_clock::now();
    TraceNavigationViewMenuDiagnostics(L"navigation.menu-dropdown.show",
                                       L"hwnd={:#x} owner={:#x} point=({}, {}) items={} embedded={} active={:#x} foreground={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       reinterpret_cast<uintptr_t>(popupOwner),
                                       pt.x,
                                       pt.y,
                                       popupItems.size(),
                                       _embeddedDestinationMode ? 1 : 0,
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetForegroundWindow()));
    const auto selectedId = RedSalamander::DxUi::ContextMenu::Show(
        popupOwner, pt, popupItems, MakeAppThemeDxPalette(_appTheme, ColorToCOLORREF(_theme.background)), sessionCallbacks);
    const uint64_t elapsedUs = Debug::Perf::ElapsedUs(startedAt);
    TraceNavigationViewMenuDiagnostics(L"navigation.menu-dropdown.result",
                                       L"hwnd={:#x} selected={} durationUs={} items={} focus={:#x} active={:#x} foreground={:#x} capture={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       selectedId.has_value() ? selectedId.value() : 0,
                                       elapsedUs,
                                       popupItems.size(),
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                       reinterpret_cast<uintptr_t>(GetCapture()));
    Debug::Perf::Emit(L"navigation.ui.dropdown_popup_us",
                      L"menu",
                      elapsedUs,
                      static_cast<uint64_t>(popupItems.size()),
                      static_cast<uint64_t>(selectedId.has_value() ? selectedId.value() : 0));

    if (selectedId.has_value() && selectedId.value() != 0)
    {
        static_cast<void>(ExecuteNavigationMenuAction(static_cast<UINT>(selectedId.value())));
    }
}

void NavigationView::ShowFileSystemDriveMenuDropdown(bool ignoreInitialLeftButtonUp)
{
    const std::optional<NavigationMenuSnapshot> fileMenuOpt = TryGetFileSystemNavigationMenuItems();
    if (! fileMenuOpt.has_value())
    {
        return;
    }

    _navDropdownKind = ModernDropdownKind::Drive;
    _navDropdownPaths.clear();
    _navDropdownSelectedIndex     = -1;
    const auto clearDropdownState = wil::scope_exit([&]() noexcept
    {
        _navDropdownKind = ModernDropdownKind::None;
        _navDropdownPaths.clear();
        _navDropdownSelectedIndex = -1;
    });

    _menuButtonPressed = true;
    RenderDriveSection();

    const auto resetPressedState = wil::scope_exit([&]() noexcept
    {
        _menuButtonPressed = false;
        RenderDriveSection();
    });

    const HWND popupOwner = GetAncestor(_hWnd.get(), GA_ROOT);
    if (! popupOwner)
    {
        return;
    }

    _navigationMenuActions.clear();
    const auto clearActions = wil::scope_exit([&]() noexcept { _navigationMenuActions.clear(); });

    UINT nextId = ID_NAV_MENU_BASE;
    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(fileMenuOpt.value().count + 8u);

    const bool isFilePluginShortId         = IsFilePluginShortId(_pluginShortId);
    const auto getConnectionsManagerTarget = [&]() -> std::wstring
    {
        if (! isFilePluginShortId && IsConnectionProtocolShortId(_pluginShortId))
        {
            std::wstring target;
            target.reserve(_pluginShortId.size() + 1u);
            target.append(_pluginShortId);
            target.push_back(L':');
            return target;
        }

        return L"nav:";
    };

    const auto appendSeparatorIfNeeded = [&](std::vector<RedSalamander::DxUi::MenuFlyoutItem>& targetItems) noexcept
    {
        if (targetItems.empty() || targetItems.back().kind == RedSalamander::DxUi::MenuItemKind::Separator)
        {
            return;
        }

        targetItems.push_back(RedSalamander::DxUi::MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Separator});
    };

    const auto appendNavigationItem = [&](std::vector<RedSalamander::DxUi::MenuFlyoutItem>& targetItems,
                                          const NavigationMenuItem& item,
                                          std::vector<MenuAction>& actionSink,
                                          UINT maxId,
                                          std::wstring_view truncatedLabel) noexcept -> bool
    {
        if ((item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0)
        {
            appendSeparatorIfNeeded(targetItems);
            return true;
        }

        const bool isHeader   = (item.flags & NAV_MENU_ITEM_FLAG_HEADER) != 0;
        const bool isDisabled = (item.flags & NAV_MENU_ITEM_FLAG_DISABLED) != 0;
        const bool hasPath    = item.path && item.path[0] != L'\0';
        const bool hasCommand = item.commandId != 0;
        const bool actionable = ! isHeader && (hasPath || hasCommand);
        if (actionable && nextId > maxId)
        {
            Debug::Warning(L"[NavigationView] {} truncated (max actionable items reached)", truncatedLabel);
            return false;
        }

        RedSalamander::DxUi::MenuFlyoutItem dxItem{};
        const MenuPresentationText presentation = DecodeMenuPresentationText(item.label ? item.label : L"");
        dxItem.kind                             = isHeader ? RedSalamander::DxUi::MenuItemKind::Header : RedSalamander::DxUi::MenuItemKind::Standard;
        dxItem.text                             = presentation.label;
        dxItem.acceleratorText                  = presentation.accelerator;
        ApplyNavigationMenuIcon(dxItem, item, _menuIconSize);
        dxItem.enabled = ! isDisabled && ! isHeader;

        if (actionable)
        {
            dxItem.commandId = static_cast<int>(nextId++);
            MenuAction action;
            action.menuId = static_cast<UINT>(dxItem.commandId);
            if (hasPath)
            {
                action.type = MenuActionType::NavigatePath;
                action.path = item.path;
            }
            else
            {
                action.type      = MenuActionType::Command;
                action.commandId = item.commandId;
            }
            actionSink.push_back(std::move(action));
        }

        targetItems.push_back(std::move(dxItem));
        return true;
    };

    bool goToItemAdded        = false;
    bool connectionsItemAdded = false;
    std::optional<size_t> goToItemIndex;
    const auto registerWindowCommandAction = [&](const UINT menuId, const UINT commandId) noexcept
    {
        MenuAction action;
        action.menuId    = menuId;
        action.type      = MenuActionType::WindowCommand;
        action.commandId = commandId;
        _navigationMenuActions.push_back(std::move(action));
    };
    const auto reserveTailActionIds = [&](const unsigned int startIndex, const size_t extraReserve) noexcept -> UINT
    {
        size_t remainingActionable = extraReserve;
        for (unsigned int j = startIndex; j < fileMenuOpt.value().count; ++j)
        {
            const NavigationMenuItem& item = fileMenuOpt.value().items[j];
            if ((item.flags & (NAV_MENU_ITEM_FLAG_SEPARATOR | NAV_MENU_ITEM_FLAG_HEADER)) != 0)
            {
                continue;
            }

            const bool hasPath    = item.path && item.path[0] != L'\0';
            const bool hasCommand = item.commandId != 0;
            if (hasPath || hasCommand)
            {
                ++remainingActionable;
            }
        }

        if (remainingActionable == 0u)
        {
            return ID_NAV_MENU_MAX;
        }

        if (remainingActionable >= static_cast<size_t>(ID_NAV_MENU_MAX - ID_NAV_MENU_BASE + 1u))
        {
            return nextId > ID_NAV_MENU_BASE ? static_cast<UINT>(nextId - 1u) : ID_NAV_MENU_BASE;
        }

        const UINT cappedMax = static_cast<UINT>(ID_NAV_MENU_MAX - remainingActionable);
        return cappedMax >= nextId ? cappedMax : static_cast<UINT>(nextId - 1u);
    };

    const auto tryAppendGoToMenu = [&](UINT maxIdForGoTo) noexcept
    {
        if (goToItemAdded || nextId > maxIdForGoTo)
        {
            return;
        }

        RedSalamander::DxUi::MenuFlyoutItem submenu{};
        const size_t actionCountBefore        = _navigationMenuActions.size();
        const UINT nextIdBefore               = nextId;
        const auto registerNavigatePathAction = [&](const UINT menuId, const std::filesystem::path& path) noexcept
        {
            MenuAction action;
            action.menuId = menuId;
            action.type   = MenuActionType::NavigatePath;
            action.path   = path.wstring();
            _navigationMenuActions.push_back(std::move(action));
        };

        if (! TryBuildPaneGoToFlyoutItem(
                _hWnd.get(), _settings, _pathHistory, _currentPath, nextId, maxIdForGoTo, registerWindowCommandAction, registerNavigatePathAction, submenu))
        {
            _navigationMenuActions.resize(actionCountBefore);
            nextId = nextIdBefore;
            return;
        }

        goToItemIndex = items.size();
        items.push_back(std::move(submenu));
        goToItemAdded = true;
    };
    const auto tryAppendConnectionsMenu = [&]() noexcept
    {
        if (connectionsItemAdded || nextId > ID_NAV_MENU_MAX)
        {
            return;
        }

        const std::wstring connectionsLabel = LoadStringResource(nullptr, IDS_MENU_CONNECTIONS);
        if (connectionsLabel.empty())
        {
            return;
        }

        struct ConnectionMenuItem
        {
            std::wstring label;
            std::wstring navName;
            std::wstring actionPath;
        };

        std::vector<ConnectionMenuItem> connectionItems;
        connectionItems.reserve(_settings && _settings->connections ? _settings->connections->items.size() + 1u : 2u);

        // Quick Connect (session-only)
        {
            ConnectionMenuItem item;
            item.navName = std::wstring(RedSalamander::Connections::kQuickConnectConnectionName);

            Common::Settings::ConnectionProfile quick{};
            RedSalamander::Connections::GetQuickConnectProfile(quick);

            if (! quick.host.empty())
            {
                item.label      = quick.port != 0u ? std::format(L"{}:{}", quick.host, quick.port) : quick.host;
                item.actionPath = std::format(L"nav:{}", item.navName);
            }
            else
            {
                item.label = LoadStringResource(nullptr, IDS_CONNECTIONS_QUICK_CONNECT);
                if (item.label.empty())
                {
                    item.label = L"<Quick Connect>";
                }
                item.actionPath = getConnectionsManagerTarget();
            }
            connectionItems.push_back(std::move(item));
        }

        // Persisted profiles
        if (_settings && _settings->connections)
        {
            for (const auto& profile : _settings->connections->items)
            {
                if (profile.name.empty() || profile.pluginId.empty())
                {
                    continue;
                }
                if (RedSalamander::Connections::IsQuickConnectConnectionName(profile.name))
                {
                    continue;
                }

                ConnectionMenuItem item;
                item.label      = profile.name;
                item.navName    = profile.name;
                item.actionPath = std::format(L"nav:{}", item.navName);
                connectionItems.push_back(std::move(item));
            }
        }

        if (connectionItems.size() > 1u)
        {
            std::sort(connectionItems.begin() + 1u,
                      connectionItems.end(),
                      [](const ConnectionMenuItem& a, const ConnectionMenuItem& b)
            {
                const int labelCompare = CompareTextNoCase(a.label, b.label);
                if (labelCompare != 0)
                {
                    return labelCompare < 0;
                }

                return CompareTextNoCase(a.navName, b.navName) < 0;
            });
        }

        std::vector<RedSalamander::DxUi::MenuFlyoutItem> connectionChildren;
        if (! connectionItems.empty())
        {
            const std::wstring managerLabel = LoadStringResource(nullptr, IDS_MENU_CONNECTIONS_ELLIPSIS);
            if (! managerLabel.empty() && nextId <= ID_NAV_MENU_MAX)
            {
                RedSalamander::DxUi::MenuFlyoutItem managerItem{};
                managerItem.text      = managerLabel;
                managerItem.commandId = static_cast<int>(nextId++);
                managerItem.iconGlyph.assign(1u, kMenuGlyphConnections.glyph);
                connectionChildren.push_back(std::move(managerItem));

                MenuAction action;
                action.menuId = static_cast<UINT>(connectionChildren.back().commandId);
                action.type   = MenuActionType::NavigatePath;
                action.path   = getConnectionsManagerTarget();
                _navigationMenuActions.push_back(std::move(action));
            }

            appendSeparatorIfNeeded(connectionChildren);
        }

        if (connectionItems.empty())
        {
            const std::wstring emptyLabel = LoadStringResource(nullptr, IDS_MENU_EMPTY);
            connectionChildren.push_back(
                RedSalamander::DxUi::MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Header, .text = emptyLabel.empty() ? L"(Empty)" : emptyLabel});
        }
        else
        {
            for (const auto& item : connectionItems)
            {
                if (nextId > ID_NAV_MENU_MAX)
                {
                    break;
                }

                RedSalamander::DxUi::MenuFlyoutItem child{};
                child.text      = item.label;
                child.commandId = static_cast<int>(nextId++);
                connectionChildren.push_back(std::move(child));

                MenuAction action;
                action.menuId = static_cast<UINT>(connectionChildren.back().commandId);
                action.type   = MenuActionType::NavigatePath;
                action.path   = item.actionPath;
                _navigationMenuActions.push_back(std::move(action));
            }
        }

        RedSalamander::DxUi::MenuFlyoutItem submenu{};
        submenu.text = connectionsLabel;
        submenu.iconGlyph.assign(1u, kMenuGlyphConnections.glyph);
        submenu.children = std::move(connectionChildren);
        items.push_back(std::move(submenu));
        connectionsItemAdded = true;
    };

    for (unsigned int i = 0; i < fileMenuOpt.value().count; ++i)
    {
        const NavigationMenuItem& item = fileMenuOpt.value().items[i];
        const bool isSeparator         = (item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0;
        if (isSeparator)
        {
            if (! goToItemAdded)
            {
                const NavigationMenuItem* nextNonSeparator = nullptr;
                for (unsigned int j = i + 1u; j < fileMenuOpt.value().count; ++j)
                {
                    if ((fileMenuOpt.value().items[j].flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0)
                    {
                        continue;
                    }

                    nextNonSeparator = &fileMenuOpt.value().items[j];
                    break;
                }

                if (nextNonSeparator != nullptr && LooksLikeDriveRootPath(nextNonSeparator->path))
                {
                    tryAppendGoToMenu(reserveTailActionIds(i + 1u, 0u));
                }
            }

            appendSeparatorIfNeeded(items);
            continue;
        }

        if (! appendNavigationItem(items, item, _navigationMenuActions, ID_NAV_MENU_MAX, L"File-system drive menu"))
        {
            break;
        }
    }

    if (! goToItemAdded)
    {
        tryAppendGoToMenu(ID_NAV_MENU_MAX);
    }

    if (! connectionsItemAdded)
    {
        const size_t connectionsIndex = items.size();
        tryAppendConnectionsMenu();
        if (connectionsItemAdded && goToItemIndex.has_value() && connectionsIndex + 1u == items.size() && connectionsIndex > goToItemIndex.value())
        {
            RedSalamander::DxUi::MenuFlyoutItem connectionItem = std::move(items.back());
            items.pop_back();
            const size_t insertIndex = std::min(goToItemIndex.value(), items.size());
            items.insert(items.begin() + static_cast<std::vector<RedSalamander::DxUi::MenuFlyoutItem>::difference_type>(insertIndex),
                         std::move(connectionItem));
            goToItemIndex = insertIndex + 1u;
        }
    }

    // Append hot paths with showInMenu flag.
    if (_settings && _settings->hotPaths.has_value())
    {
        bool anyVisible = false;
        for (const auto& slot : _settings->hotPaths.value().slots)
        {
            if (slot.has_value() && slot.value().showInMenu && ! slot.value().path.empty())
            {
                anyVisible = true;
                break;
            }
        }

        if (anyVisible)
        {
            appendSeparatorIfNeeded(items);

            const auto& slots = _settings->hotPaths.value().slots;
            for (size_t i = 0; i < slots.size(); ++i)
            {
                if (! slots[i].has_value() || ! slots[i].value().showInMenu || slots[i].value().path.empty())
                {
                    continue;
                }

                if (nextId > ID_NAV_MENU_MAX)
                {
                    break;
                }

                const auto& slot        = slots[i].value();
                const UINT id           = nextId++;
                const wchar_t digitChar = (i < 9) ? static_cast<wchar_t>(L'1' + i) : L'0';

                std::wstring label;
                if (! slot.label.empty())
                {
                    label = std::format(L"{}: {}", digitChar, slot.label);
                }
                else
                {
                    label = std::format(L"{}: {}", digitChar, slot.path);
                }

                std::wstring accelerator;
                if (_settings)
                {
                    std::wstring commandId = L"cmd/pane/hotPath/";
                    commandId.push_back(digitChar);
                    if (const std::optional<std::wstring> shortcutOpt = TryGetShortcutTextForCommandId(*_settings, commandId))
                    {
                        accelerator = shortcutOpt.value();
                    }
                }

                RedSalamander::DxUi::MenuFlyoutItem item{};
                item.text            = label;
                item.acceleratorText = accelerator;
                item.commandId       = static_cast<int>(id);
                items.push_back(std::move(item));

                MenuAction action;
                action.menuId = id;
                action.type   = MenuActionType::NavigatePath;
                action.path   = slot.path;
                _navigationMenuActions.push_back(std::move(action));
            }
        }
    }

    if (items.empty())
    {
        return;
    }

    POINT pt = MakeDropdownAnchorPoint(_sectionDriveRect, false, _embeddedDestinationMode);
    ClientToScreen(_hWnd.get(), &pt);

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.ignoreInitialLeftButtonUp = ignoreInitialLeftButtonUp;
    ApplyEmbeddedDestinationDropdownOptions(sessionCallbacks, _embeddedDestinationMode);

    const auto startedAt = std::chrono::steady_clock::now();
    const auto selectedId =
        RedSalamander::DxUi::ContextMenu::Show(popupOwner, pt, items, MakeAppThemeDxPalette(_appTheme, ColorToCOLORREF(_theme.background)), sessionCallbacks);
    Debug::Perf::Emit(L"navigation.ui.dropdown_popup_us",
                      L"drive",
                      Debug::Perf::ElapsedUs(startedAt),
                      static_cast<uint64_t>(items.size()),
                      static_cast<uint64_t>(selectedId.has_value() ? selectedId.value() : 0));

    if (selectedId.has_value() && selectedId.value() != 0)
    {
        static_cast<void>(ExecuteNavigationMenuAction(static_cast<UINT>(selectedId.value())));
    }
}

void NavigationView::ShowHistoryDropdown(bool ignoreInitialLeftButtonUp, bool focusFirstNavigableItem)
{
    TraceNavigationViewMenuDiagnostics(L"navigation.history-dropdown.enter",
                                       L"hwnd={:#x} historyCount={} ignoreInitialLeftButtonUp={} focusFirst={} embedded={} currentPath='{}' focus={:#x} active={:#x} capture={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       _pathHistory.size(),
                                       ignoreInitialLeftButtonUp ? 1 : 0,
                                       focusFirstNavigableItem ? 1 : 0,
                                       _embeddedDestinationMode ? 1 : 0,
                                       _currentPath.has_value() ? _currentPath->wstring() : std::wstring{},
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetCapture()));
    if (_pathHistory.empty())
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.history-dropdown.skip",
                                           L"hwnd={:#x} reason=empty-history",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()));
        return;
    }

    if (! _hWnd)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.history-dropdown.skip", L"hwnd=null reason=no-window");
        return;
    }

    _navDropdownKind = ModernDropdownKind::History;
    _navDropdownPaths.assign(_pathHistory.begin(), _pathHistory.end());
    _navDropdownSelectedIndex = -1;

    auto clearDropdownState = wil::scope_exit([&]() noexcept
    {
        _navDropdownKind = ModernDropdownKind::None;
        _navDropdownPaths.clear();
        _navDropdownSelectedIndex = -1;
    });

    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(_navDropdownPaths.size());

    auto historyEntryHasActiveFilter = [&](const std::filesystem::path& historyPath) -> bool
    {
        if (historyPath.empty() || ! _settings || ! _settings->folders.has_value())
        {
            return false;
        }

        const auto& folders = _settings->folders.value();
        if (folders.historyFilters.empty())
        {
            return false;
        }

        const std::wstring_view historyText = historyPath.native();
        const auto it =
            std::find_if(folders.historyFilters.begin(), folders.historyFilters.end(), [&](const Common::Settings::FolderHistoryFilterState& state) noexcept {
            return state.enabled && ! state.path.empty() && EqualsNoCase(state.path.native(), historyText);
        });
        if (it == folders.historyFilters.end())
        {
            return false;
        }

        const MaskSyntax::WildcardMask mask = MaskSyntax::ParseWildcardMask(it->text);
        return ! mask.includePatterns.empty() || ! mask.excludePatterns.empty();
    };

    int selectedIndex = 0;
    for (size_t i = 0; i < _navDropdownPaths.size(); ++i)
    {
        const auto& path = _navDropdownPaths[i];

        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.kind      = RedSalamander::DxUi::MenuItemKind::Radio;
        item.text      = path.wstring();
        item.commandId = ID_HISTORY_BASE + static_cast<int>(i);
        if (historyEntryHasActiveFilter(path))
        {
            item.iconGlyph.assign(1, FluentIcons::kFilter);
        }

        if (_currentPath && wil::compare_string_ordinal(path.wstring(), _currentPath->wstring(), true) == wistd::weak_ordering::equivalent)
        {
            item.checked  = true;
            selectedIndex = static_cast<int>(i);
        }

        items.push_back(std::move(item));
    }

    if (items.empty())
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.history-dropdown.skip",
                                           L"hwnd={:#x} reason=no-items historyCount={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           _navDropdownPaths.size());
        return;
    }

    _navDropdownSelectedIndex = std::clamp(selectedIndex, 0, static_cast<int>(items.size()) - 1);

    POINT pt = MakeDropdownAnchorPoint(_sectionHistoryRect, true, _embeddedDestinationMode);
    ClientToScreen(_hWnd.get(), &pt);
    const HWND popupOwner = GetAncestor(_hWnd.get(), GA_ROOT);
    if (! popupOwner)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.history-dropdown.skip",
                                           L"hwnd={:#x} reason=no-popup-owner",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()));
        return;
    }

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.ignoreInitialLeftButtonUp = ignoreInitialLeftButtonUp;
    sessionCallbacks.focusFirstNavigableItem   = focusFirstNavigableItem;
    ApplyEmbeddedDestinationDropdownOptions(sessionCallbacks, _embeddedDestinationMode);
    sessionCallbacks.rootHorizontalAlignment   = RedSalamander::DxUi::ContextMenuRootHorizontalAlignment::End;

    const auto startedAt = std::chrono::steady_clock::now();
    TraceNavigationViewMenuDiagnostics(L"navigation.history-dropdown.show",
                                       L"hwnd={:#x} owner={:#x} point=({}, {}) items={} selectedIndex={} ignoreInitialLeftButtonUp={} focusFirst={} active={:#x} foreground={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       reinterpret_cast<uintptr_t>(popupOwner),
                                       pt.x,
                                       pt.y,
                                       items.size(),
                                       _navDropdownSelectedIndex,
                                       ignoreInitialLeftButtonUp ? 1 : 0,
                                       focusFirstNavigableItem ? 1 : 0,
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetForegroundWindow()));
    const auto selectedId =
        RedSalamander::DxUi::ContextMenu::Show(popupOwner, pt, items, MakeAppThemeDxPalette(_appTheme, ColorToCOLORREF(_theme.background)), sessionCallbacks);
    const uint64_t elapsedUs = Debug::Perf::ElapsedUs(startedAt);
    TraceNavigationViewMenuDiagnostics(L"navigation.history-dropdown.result",
                                       L"hwnd={:#x} selected={} durationUs={} items={} selectedIndex={} focus={:#x} active={:#x} foreground={:#x} capture={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       selectedId.has_value() ? selectedId.value() : 0,
                                       elapsedUs,
                                       items.size(),
                                       _navDropdownSelectedIndex,
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                       reinterpret_cast<uintptr_t>(GetCapture()));
    Debug::Perf::Emit(L"navigation.ui.dropdown_popup_us",
                      L"history",
                      elapsedUs,
                      static_cast<uint64_t>(items.size()),
                      static_cast<uint64_t>(_navDropdownSelectedIndex >= 0 ? _navDropdownSelectedIndex : 0));

    if (selectedId.has_value() && selectedId.value() >= ID_HISTORY_BASE && selectedId.value() <= ID_HISTORY_MAX)
    {
        const size_t historyIndex = static_cast<size_t>(selectedId.value() - ID_HISTORY_BASE);
        if (historyIndex < _navDropdownPaths.size())
        {
            RequestPathChange(_navDropdownPaths[historyIndex]);
        }
    }

    if (_requestFolderViewFocusCallback)
    {
        const HWND root = GetAncestor(_hWnd.get(), GA_ROOT);
        if (root && GetActiveWindow() != root)
        {
            SetActiveWindow(root);
        }
        _requestFolderViewFocusCallback();
    }
}

void NavigationView::ShowDiskInfoDropdown(bool ignoreInitialLeftButtonUp, bool focusFirstNavigableItem)
{
    if (! _showDiskInfoSection || ! _currentPluginPath || ! _driveInfo)
        return;

    _navDropdownKind = ModernDropdownKind::DiskInfo;
    _navDropdownPaths.clear();
    _navDropdownSelectedIndex     = -1;
    const auto clearDropdownState = wil::scope_exit([&]() noexcept
    {
        _navDropdownKind = ModernDropdownKind::None;
        _navDropdownPaths.clear();
        _navDropdownSelectedIndex = -1;
    });

    UpdateDiskInfo(false);

    uint64_t usedBytes = 0;
    bool hasUsedBytes  = false;
    if (_hasUsedBytes)
    {
        usedBytes    = _usedBytes;
        hasUsedBytes = true;
    }
    else if (_hasTotalBytes && _hasFreeBytes && _totalBytes >= _freeBytes)
    {
        usedBytes    = _totalBytes - _freeBytes;
        hasUsedBytes = true;
    }

    double usedPercent  = 0.0;
    bool hasUsedPercent = false;
    if (_hasTotalBytes && _totalBytes > 0 && hasUsedBytes)
    {
        usedPercent = static_cast<double>(usedBytes) * 100.0 / static_cast<double>(_totalBytes);
        if (usedPercent < 0.0)
        {
            usedPercent = 0.0;
        }
        if (usedPercent > 100.0)
        {
            usedPercent = 100.0;
        }
        hasUsedPercent = true;
    }

    std::wstring headerName;
    if (! _driveDisplayName.empty())
    {
        headerName = _driveDisplayName;
    }
    else
    {
        const bool isFilePlugin = _pluginShortId.empty() || EqualsNoCase(_pluginShortId, L"file");
        if (isFilePlugin)
        {
            const std::filesystem::path root = _currentPluginPath.value().root_path();
            headerName                       = root.empty() ? _currentPluginPath.value().wstring() : root.wstring();
        }
        else
        {
            headerName = L"/";
        }
    }
    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.push_back(RedSalamander::DxUi::MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Header,
                                                        .text = FormatStringResource(nullptr, IDS_FMT_DISK_INFO_HEADER, headerName)});

    const std::wstring pathText = _currentPluginPath.value().wstring();
    _driveMenuActions.clear();
    const auto clearActions = wil::scope_exit([&]() noexcept { _driveMenuActions.clear(); });

    const NavigationMenuItem* driveMenuItems = nullptr;
    unsigned int driveMenuCount              = 0;
    const HRESULT itemsHr                    = _driveInfo->GetDriveMenuItems(pathText.c_str(), &driveMenuItems, &driveMenuCount);
    const bool hasDriveMenuItems             = SUCCEEDED(itemsHr) && driveMenuItems && driveMenuCount > 0;

    const auto appendSeparatorIfNeeded = [&]
    {
        if (items.empty() || items.back().kind == RedSalamander::DxUi::MenuItemKind::Separator)
        {
            return;
        }
        items.push_back(RedSalamander::DxUi::MenuFlyoutItem{.kind = RedSalamander::DxUi::MenuItemKind::Separator});
    };

    const auto appendInfoLine = [&](std::wstring text)
    {
        MenuInfoLineText line = SplitFormattedMenuInfoLine(std::move(text));
        items.push_back(RedSalamander::DxUi::MenuFlyoutItem{
            .kind = RedSalamander::DxUi::MenuItemKind::Info, .text = std::move(line.label), .acceleratorText = std::move(line.value)});
    };

    const bool hasInfoLines = (! _volumeLabel.empty() || ! _fileSystem.empty());
    const bool hasSizeLines = (_hasTotalBytes || hasUsedBytes || _hasFreeBytes);

    if (hasInfoLines || hasSizeLines || hasUsedPercent || hasDriveMenuItems)
    {
        appendSeparatorIfNeeded();
    }

    if (! _volumeLabel.empty())
    {
        const std::wstring volumeLabel = FormatStringResource(nullptr, IDS_FMT_DISK_VOLUME_LABEL, _volumeLabel);
        appendInfoLine(volumeLabel);
    }
    if (! _fileSystem.empty())
    {
        const std::wstring fileSystem = FormatStringResource(nullptr, IDS_FMT_DISK_FILE_SYSTEM, _fileSystem);
        appendInfoLine(fileSystem);
    }

    if (hasSizeLines && hasInfoLines)
    {
        appendSeparatorIfNeeded();
    }

    if (_hasTotalBytes)
    {
        const std::wstring totalSpace = FormatStringResource(nullptr, IDS_FMT_DISK_TOTAL_SPACE, FormatBytesCompact(_totalBytes), _totalBytes);
        appendInfoLine(totalSpace);
    }
    if (hasUsedBytes)
    {
        const std::wstring usedSpace = FormatStringResource(nullptr, IDS_FMT_DISK_USED_SPACE, FormatBytesCompact(usedBytes), usedBytes);
        appendInfoLine(usedSpace);
    }
    if (_hasFreeBytes)
    {
        const std::wstring freeSpace = FormatStringResource(nullptr, IDS_FMT_DISK_FREE_SPACE, FormatBytesCompact(_freeBytes), _freeBytes);
        appendInfoLine(freeSpace);
    }

    if (hasUsedPercent && (hasInfoLines || hasSizeLines))
    {
        appendSeparatorIfNeeded();
        const std::wstring percentUsed = FormatStringResource(nullptr, IDS_FMT_DISK_USED_PERCENT, usedPercent);
        appendInfoLine(percentUsed);
    }

    if (hasDriveMenuItems)
    {
        appendSeparatorIfNeeded();

        constexpr unsigned int kMaxActions = ID_DRIVE_MENU_MAX - ID_DRIVE_MENU_BASE + 1u;

        UINT nextId = ID_DRIVE_MENU_BASE;
        for (unsigned int i = 0; i < driveMenuCount; ++i)
        {
            const NavigationMenuItem& item = driveMenuItems[i];
            const bool isSeparator         = (item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0;
            if (isSeparator)
            {
                appendSeparatorIfNeeded();
                continue;
            }

            const bool isHeader   = (item.flags & NAV_MENU_ITEM_FLAG_HEADER) != 0;
            const bool isDisabled = (item.flags & NAV_MENU_ITEM_FLAG_DISABLED) != 0;
            const bool hasPath    = item.path && item.path[0] != L'\0';
            const bool hasCommand = item.commandId != 0;
            const bool actionable = ! isHeader && (hasPath || hasCommand);

            if (actionable && nextId > ID_DRIVE_MENU_MAX)
            {
                Debug::Warning(L"[NavigationView] Drive menu truncated (max {} actionable items)", kMaxActions);
                break;
            }

            RedSalamander::DxUi::MenuFlyoutItem dxItem{};
            const MenuPresentationText presentation = DecodeMenuPresentationText(item.label ? item.label : L"");
            dxItem.kind                             = isHeader ? RedSalamander::DxUi::MenuItemKind::Header : RedSalamander::DxUi::MenuItemKind::Standard;
            dxItem.text                             = presentation.label;
            dxItem.acceleratorText                  = presentation.accelerator;
            ApplyNavigationMenuIcon(dxItem, item, _menuIconSize);
            dxItem.enabled = ! isDisabled && ! isHeader;

            if (actionable)
            {
                dxItem.commandId = static_cast<int>(nextId++);
                MenuAction action;
                action.menuId = static_cast<UINT>(dxItem.commandId);
                if (hasPath)
                {
                    action.type = MenuActionType::NavigatePath;
                    action.path = item.path;
                }
                else
                {
                    action.type      = MenuActionType::Command;
                    action.commandId = item.commandId;
                }
                _driveMenuActions.push_back(std::move(action));
            }

            items.push_back(std::move(dxItem));
        }
    }

    if (items.empty())
    {
        return;
    }

    RECT rc  = _sectionDiskInfoRect;
    POINT pt = MakeDropdownAnchorPoint(rc, true, _embeddedDestinationMode);
    ClientToScreen(_hWnd.get(), &pt);
    const HWND popupOwner = GetAncestor(_hWnd.get(), GA_ROOT);
    if (! popupOwner)
    {
        return;
    }

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.ignoreInitialLeftButtonUp = ignoreInitialLeftButtonUp;
    sessionCallbacks.focusFirstNavigableItem   = focusFirstNavigableItem;
    sessionCallbacks.rootHorizontalAlignment   = RedSalamander::DxUi::ContextMenuRootHorizontalAlignment::End;
    ApplyEmbeddedDestinationDropdownOptions(sessionCallbacks, _embeddedDestinationMode);

    const auto startedAt = std::chrono::steady_clock::now();
    const auto selectedId =
        RedSalamander::DxUi::ContextMenu::Show(popupOwner, pt, items, MakeAppThemeDxPalette(_appTheme, ColorToCOLORREF(_theme.background)), sessionCallbacks);
    Debug::Perf::Emit(L"navigation.ui.dropdown_popup_us",
                      L"disk-info",
                      Debug::Perf::ElapsedUs(startedAt),
                      static_cast<uint64_t>(items.size()),
                      static_cast<uint64_t>(selectedId.has_value() ? selectedId.value() : 0));
    if (selectedId.has_value() && selectedId.value() != 0)
    {
        static_cast<void>(ExecuteDriveMenuAction(static_cast<UINT>(selectedId.value())));
    }
}

bool NavigationView::TryGetSiblingFolders(const std::filesystem::path& parentPath, std::vector<std::filesystem::path>& siblings)
{
    siblings.clear();

    if (! _fileSystemPlugin)
    {
        return false;
    }

    const std::filesystem::path pluginParentPath = ToPluginPath(parentPath);
    auto borrowed = DirectoryInfoCache::GetInstance().BorrowDirectoryInfo(_fileSystemPlugin.get(), pluginParentPath, DirectoryInfoCache::BorrowMode::CacheOnly);
    IFilesInformation* info = borrowed.Get();
    if (borrowed.Status() != S_OK || ! info)
    {
        QueueSiblingPrefetchForParent(parentPath);
        return false;
    }

    FileInfo* entry  = nullptr;
    const HRESULT hr = info->GetBuffer(&entry);
    if (FAILED(hr) || entry == nullptr)
    {
        return true;
    }

    while (entry != nullptr)
    {
        if ((entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
            const std::wstring_view name(entry->FileName, nameChars);
            if (name != L"." && name != L"..")
            {
                siblings.push_back(parentPath / std::wstring(name));
            }
        }

        if (entry->NextEntryOffset == 0)
        {
            break;
        }

        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset);
    }

    std::sort(siblings.begin(), siblings.end(), [](const std::filesystem::path& a, const std::filesystem::path& b) {
        return _wcsicmp(a.filename().c_str(), b.filename().c_str()) < 0;
    });

    return true;
}

void NavigationView::ShowSiblingsDropdown(size_t separatorIndex)
{
    TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.enter",
                                       L"hwnd={:#x} index={} separators={} segments={} embedded={} focus={:#x} active={:#x} capture={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       separatorIndex,
                                       _separators.size(),
                                       _segments.size(),
                                       _embeddedDestinationMode ? 1 : 0,
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetCapture()));
    if (separatorIndex >= _separators.size())
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.skip",
                                           L"hwnd={:#x} reason=separator-out-of-range index={} separators={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           separatorIndex,
                                           _separators.size());
        return;
    }

    // Sibling dropdown is only valid for separators between two real segments.
    const auto& separator = _separators[separatorIndex];
    if (separator.leftSegmentIndex >= _segments.size() || separator.rightSegmentIndex >= _segments.size())
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.skip",
                                           L"hwnd={:#x} reason=segment-out-of-range index={} left={} right={} segments={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           separatorIndex,
                                           separator.leftSegmentIndex,
                                           separator.rightSegmentIndex,
                                           _segments.size());
        return;
    }

    const auto& leftSegment  = _segments[separator.leftSegmentIndex];
    const auto& rightSegment = _segments[separator.rightSegmentIndex];
    if (leftSegment.isEllipsis || rightSegment.isEllipsis)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.skip",
                                           L"hwnd={:#x} reason=ellipsis index={} leftEllipsis={} rightEllipsis={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           separatorIndex,
                                           leftSegment.isEllipsis ? 1 : 0,
                                           rightSegment.isEllipsis ? 1 : 0);
        return;
    }

    const auto& segment                               = rightSegment;
    const std::filesystem::path normalizedSegmentPath = NormalizeDirectoryPath(segment.fullPath);
    std::filesystem::path parentPath                  = normalizedSegmentPath.parent_path();
    if (parentPath.empty())
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.skip",
                                           L"hwnd={:#x} reason=no-parent index={} segment='{}'",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           separatorIndex,
                                           normalizedSegmentPath.wstring());
        return;
    }

    std::vector<std::filesystem::path> siblings;
    if (! TryGetSiblingFolders(parentPath, siblings) || siblings.empty())
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.skip",
                                           L"hwnd={:#x} reason=no-siblings index={} parent='{}'",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           separatorIndex,
                                           parentPath.wstring());
        return;
    }

    // Set active separator and start rotation animation
    _activeSeparatorIndex = static_cast<int>(separatorIndex);
    StartSeparatorAnimation(separatorIndex, 90.0f);
    RenderPathSection();

    const auto& bounds = separator.bounds;
    if (! _hWnd)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.skip", L"hwnd=null reason=no-window index={}", separatorIndex);
        StartSeparatorAnimation(separatorIndex, 0.0f);
        _activeSeparatorIndex = -1;
        RenderPathSection();
        return;
    }

    _navDropdownKind          = ModernDropdownKind::Siblings;
    _navDropdownPaths         = siblings;
    _navDropdownSelectedIndex = -1;

    const std::filesystem::path normalizedCurrentPath = NormalizeDirectoryPath(segment.fullPath);
    const std::wstring currentPathText                = normalizedCurrentPath.wstring();
    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(_navDropdownPaths.size());

    int selectedIndex = 0;
    for (size_t i = 0; i < _navDropdownPaths.size(); ++i)
    {
        const std::filesystem::path normalizedSiblingPath = NormalizeDirectoryPath(_navDropdownPaths[i]);
        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.kind      = RedSalamander::DxUi::MenuItemKind::Radio;
        item.text      = FilenameOrPath(normalizedSiblingPath);
        item.commandId = static_cast<int>(ID_SIBLING_BASE + i);

        if (wil::compare_string_ordinal(normalizedSiblingPath.wstring(), currentPathText, true) == wistd::weak_ordering::equivalent)
        {
            item.checked  = true;
            selectedIndex = static_cast<int>(i);
        }

        items.push_back(std::move(item));
    }

    if (items.empty())
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.skip",
                                           L"hwnd={:#x} reason=no-items index={} siblingCount={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           separatorIndex,
                                           _navDropdownPaths.size());
        _navDropdownKind = ModernDropdownKind::None;
        _navDropdownPaths.clear();
        _navDropdownSelectedIndex = -1;
        StartSeparatorAnimation(separatorIndex, 0.0f);
        _activeSeparatorIndex = -1;
        RenderPathSection();
        return;
    }

    _navDropdownSelectedIndex = std::clamp(selectedIndex, 0, static_cast<int>(items.size()) - 1);
    _menuOpenForSeparator     = static_cast<int>(separatorIndex);

    auto clearDropdownState = wil::scope_exit([&]() noexcept
    {
        _navDropdownKind = ModernDropdownKind::None;
        _navDropdownPaths.clear();
        _navDropdownSelectedIndex = -1;
        StartSeparatorAnimation(separatorIndex, 0.0f);
        _menuOpenForSeparator            = -1;
        _pendingSeparatorMenuSwitchIndex = -1;
        _activeSeparatorIndex            = -1;
        RenderPathSection();
    });

    POINT pt = {static_cast<LONG>(std::lround(bounds.left + static_cast<float>(_sectionPathRect.left))),
                ResolveBreadcrumbDropdownAnchorY(bounds, _sectionPathRect.top, _embeddedDestinationMode)};
    ClientToScreen(_hWnd.get(), &pt);
    const HWND popupOwner = GetAncestor(_hWnd.get(), GA_ROOT);
    if (! popupOwner)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.skip",
                                           L"hwnd={:#x} reason=no-popup-owner index={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           separatorIndex);
        return;
    }

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.ignoreInitialLeftButtonUp = true;
    ApplyEmbeddedDestinationDropdownOptions(sessionCallbacks, _embeddedDestinationMode);

    const auto startedAt = std::chrono::steady_clock::now();
    TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.show",
                                       L"hwnd={:#x} owner={:#x} point=({}, {}) index={} items={} selectedIndex={} embedded={} active={:#x} foreground={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       reinterpret_cast<uintptr_t>(popupOwner),
                                       pt.x,
                                       pt.y,
                                       separatorIndex,
                                       items.size(),
                                       _navDropdownSelectedIndex,
                                       _embeddedDestinationMode ? 1 : 0,
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetForegroundWindow()));
    const auto selectedId =
        RedSalamander::DxUi::ContextMenu::Show(popupOwner, pt, items, MakeAppThemeDxPalette(_appTheme, ColorToCOLORREF(_theme.background)), sessionCallbacks);
    const uint64_t elapsedUs = Debug::Perf::ElapsedUs(startedAt);
    TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.result",
                                       L"hwnd={:#x} selected={} durationUs={} index={} items={} selectedIndex={} focus={:#x} active={:#x} foreground={:#x} capture={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       selectedId.has_value() ? selectedId.value() : 0,
                                       elapsedUs,
                                       separatorIndex,
                                       items.size(),
                                       _navDropdownSelectedIndex,
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                       reinterpret_cast<uintptr_t>(GetCapture()));
    Debug::Perf::Emit(L"navigation.ui.dropdown_popup_us",
                      L"siblings",
                      elapsedUs,
                      static_cast<uint64_t>(items.size()),
                      static_cast<uint64_t>(_navDropdownSelectedIndex >= 0 ? _navDropdownSelectedIndex : 0));

    if (selectedId.has_value() && selectedId.value() >= ID_SIBLING_BASE)
    {
        const size_t siblingIndex = static_cast<size_t>(selectedId.value() - ID_SIBLING_BASE);
        if (siblingIndex < _navDropdownPaths.size())
        {
            RequestPathChange(_navDropdownPaths[siblingIndex]);
        }
    }

    if (_requestFolderViewFocusCallback)
    {
        const HWND root = GetAncestor(_hWnd.get(), GA_ROOT);
        if (root && GetActiveWindow() != root)
        {
            SetActiveWindow(root);
        }
        _requestFolderViewFocusCallback();
    }
}
