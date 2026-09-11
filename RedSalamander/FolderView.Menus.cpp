#include "FolderViewInternal.h"

#include "FluentIcons.h"
#include "LocalizationManager.h"
#include "RedSalamander.h"
#include "ShortcutManager.h"

namespace
{
int FindMenuItemPosById(HMENU menu, UINT id) noexcept
{
    if (! menu)
    {
        return -1;
    }

    const int count = GetMenuItemCount(menu);
    if (count < 0)
    {
        return -1;
    }

    for (int pos = 0; pos < count; ++pos)
    {
        if (GetMenuItemID(menu, pos) == id)
        {
            return pos;
        }
    }

    return -1;
}

bool IsMenuSeparatorAt(HMENU menu, int pos) noexcept
{
    if (! menu || pos < 0)
    {
        return false;
    }

    MENUITEMINFOW mii{};
    mii.cbSize = sizeof(mii);
    mii.fMask  = MIIM_FTYPE;
    if (! GetMenuItemInfoW(menu, static_cast<UINT>(pos), TRUE, &mii))
    {
        return false;
    }

    return (mii.fType & MFT_SEPARATOR) != 0;
}

bool MenuContainsCommandIdRecursive(HMENU menu, UINT commandId) noexcept
{
    if (! menu)
    {
        return false;
    }

    if (FindMenuItemPosById(menu, commandId) >= 0)
    {
        return true;
    }

    const int count = GetMenuItemCount(menu);
    if (count <= 0)
    {
        return false;
    }

    for (int pos = 0; pos < count; ++pos)
    {
        HMENU subMenu = GetSubMenu(menu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (MenuContainsCommandIdRecursive(subMenu, commandId))
        {
            return true;
        }
    }

    return false;
}

HMENU FindSubMenuContainingCommandId(HMENU menu, UINT commandId) noexcept
{
    if (! menu)
    {
        return nullptr;
    }

    const int count = GetMenuItemCount(menu);
    for (int pos = 0; pos < count; ++pos)
    {
        HMENU subMenu = GetSubMenu(menu, pos);
        if (! subMenu)
        {
            continue;
        }
        if (FindMenuItemPosById(subMenu, commandId) >= 0)
        {
            return subMenu;
        }
        if (HMENU nested = FindSubMenuContainingCommandId(subMenu, commandId))
        {
            return nested;
        }
    }
    return nullptr;
}

[[nodiscard]] std::wstring StripMenuMnemonicMarkers(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size());

    for (size_t i = 0; i < text.size(); ++i)
    {
        if (text[i] != L'&')
        {
            result.push_back(text[i]);
            continue;
        }

        if ((i + 1u) < text.size() && text[i + 1u] == L'&')
        {
            result.push_back(L'&');
            ++i;
        }
    }

    return result;
}

void RemoveOverlaySampleSubmenu(HMENU menu, UINT sampleErrorCommandId) noexcept
{
    if (! menu)
    {
        return;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return;
    }

    for (int pos = 0; pos < itemCount; ++pos)
    {
        HMENU subMenu = GetSubMenu(menu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (! MenuContainsCommandIdRecursive(subMenu, sampleErrorCommandId))
        {
            continue;
        }

        DeleteMenu(menu, static_cast<UINT>(pos), MF_BYPOSITION);
        if (pos > 0 && IsMenuSeparatorAt(menu, pos - 1))
        {
            DeleteMenu(menu, static_cast<UINT>(pos - 1), MF_BYPOSITION);
        }
        break;
    }
}

[[nodiscard]] std::wstring VkToMenuShortcutText(uint32_t vk) noexcept
{
    vk &= 0xFFu;

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

[[nodiscard]] std::wstring FormatMenuChordText(uint32_t vk, uint32_t modifiers) noexcept
{
    std::wstring result;
    auto appendPart = [&](std::wstring_view part)
    {
        if (part.empty())
        {
            return;
        }
        if (! result.empty())
        {
            result.append(L"+");
        }
        result.append(part);
    };

    const uint32_t maskedMods = modifiers & 0x7u;
    if ((maskedMods & ShortcutManager::kModCtrl) != 0)
    {
        appendPart(LoadEmbeddedStringResource(nullptr, IDS_MOD_CTRL));
    }
    if ((maskedMods & ShortcutManager::kModAlt) != 0)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_ALT));
    }
    if ((maskedMods & ShortcutManager::kModShift) != 0)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_SHIFT));
    }

    appendPart(VkToMenuShortcutText(vk));
    return result;
}

[[nodiscard]] std::optional<std::wstring_view> TryGetCommandIdForContextMenuItem(UINT menuCommandId) noexcept
{
    switch (menuCommandId)
    {
        case CmdOpen: return L"cmd/pane/executeOpen";
        case CmdViewSpace: return L"cmd/pane/viewSpace";
        case CmdDelete: return L"cmd/pane/delete";
        case CmdRename: return L"cmd/pane/rename";
        case CmdCut: return L"cmd/pane/clipboardCut";
        case CmdCopy: return L"cmd/pane/clipboardCopy";
        case CmdPaste: return L"cmd/pane/clipboardPaste";
        case CmdRefresh: return L"cmd/pane/refresh";
        case IDM_PANE_CREATE_DIR: return L"cmd/pane/createDirectory";
        case IDM_PANE_EDIT_NEW: return L"cmd/pane/editNew";
        case CmdProperties: return L"cmd/pane/openProperties";
    }

    return std::nullopt;
}

[[nodiscard]] bool IsItemTargetContextCommand(const UINT commandId) noexcept
{
    switch (commandId)
    {
        case CmdOpen:
        case CmdOpenWith:
        case CmdDelete:
        case CmdMove:
        case CmdRename:
        case CmdCut:
        case CmdCopy:
        case CmdArtifactInspect:
        case CmdArtifactReveal:
        case CmdProperties: return true;
        default: return false;
    }
}

void AppendArtifactContextMenu(HMENU menu,
                               const std::shared_ptr<const FileOperationArtifacts::Projection>& projection) noexcept
{
    if (! menu || ! projection || projection->classification == FileOperationArtifacts::Classification::Ordinary)
    {
        return;
    }

    wil::unique_hmenu artifactMenu(CreatePopupMenu());
    if (! artifactMenu)
    {
        return;
    }
    const auto appendCommand = [&](const UINT commandId, const UINT stringId) noexcept
    {
        const std::wstring text = LoadStringResource(nullptr, stringId);
        return ! text.empty() && AppendMenuW(artifactMenu.get(), MF_STRING, commandId, text.c_str()) != FALSE;
    };
    if (projection->capabilities.inspect)
    {
        static_cast<void>(appendCommand(CmdArtifactInspect, IDS_FILEOPS_ARTIFACT_ACTION_INSPECT));
    }
    if (projection->capabilities.reveal)
    {
        static_cast<void>(appendCommand(CmdArtifactReveal, IDS_FILEOPS_ARTIFACT_ACTION_REVEAL));
    }
    if (GetMenuItemCount(artifactMenu.get()) <= 0)
    {
        return;
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_FILEOPS_ARTIFACT_MENU_TITLE);
    const int propertiesPos  = FindMenuItemPosById(menu, IDM_FOLDERVIEW_CONTEXT_PROPERTIES);
    if (title.empty() || propertiesPos < 0 ||
        InsertMenuW(menu,
                    static_cast<UINT>(propertiesPos),
                    MF_BYPOSITION | MF_POPUP,
                    reinterpret_cast<UINT_PTR>(artifactMenu.get()),
                    title.c_str()) == FALSE)
    {
        return;
    }
    if (InsertMenuW(menu, static_cast<UINT>(propertiesPos + 1), MF_BYPOSITION | MF_SEPARATOR, 0u, nullptr) == FALSE)
    {
        static_cast<void>(RemoveMenu(menu, static_cast<UINT>(propertiesPos), MF_BYPOSITION));
        return;
    }
    static_cast<void>(artifactMenu.release());
}

[[nodiscard]] bool IsSelectionAwareContextCommand(const UINT commandId) noexcept
{
    return commandId == CmdDelete || commandId == CmdMove || commandId == CmdCut || commandId == CmdCopy;
}

[[nodiscard]] std::wstring GetIconGlyphForCommand(UINT commandId) noexcept
{
    switch (commandId)
    {
        case IDM_FOLDERVIEW_CONTEXT_OPEN: return std::wstring(1, FluentIcons::kOpenFile);
        case IDM_FOLDERVIEW_CONTEXT_CUT: return std::wstring(1, FluentIcons::kCut);
        case IDM_FOLDERVIEW_CONTEXT_COPY: return std::wstring(1, FluentIcons::kCopy);
        case IDM_FOLDERVIEW_CONTEXT_PASTE: return std::wstring(1, FluentIcons::kPaste);
        case IDM_FOLDERVIEW_CONTEXT_DELETE: return std::wstring(1, FluentIcons::kDelete);
        case IDM_FOLDERVIEW_CONTEXT_RENAME: return std::wstring(1, FluentIcons::kRename);
        case IDM_FOLDERVIEW_CONTEXT_PROPERTIES: return std::wstring(1, FluentIcons::kInfo);
    }
    return {};
}

[[nodiscard]] std::vector<RedSalamander::DxUi::MenuFlyoutItem> ConvertHMenuToFlyoutItems(HMENU menu, const ShortcutManager* shortcutManager) noexcept
{
    using RedSalamander::DxUi::MenuFlyoutItem;
    using RedSalamander::DxUi::MenuItemKind;

    std::vector<MenuFlyoutItem> result;
    if (! menu)
        return result;

    const int count = GetMenuItemCount(menu);
    if (count < 0)
        return result;

    for (UINT pos = 0; pos < static_cast<UINT>(count); ++pos)
    {
        MENUITEMINFOW mii{};
        mii.cbSize = sizeof(mii);
        mii.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(menu, pos, TRUE, &mii))
            continue;

        MenuFlyoutItem item;

        if ((mii.fType & MFT_SEPARATOR) != 0)
        {
            item.kind = MenuItemKind::Separator;
            result.push_back(std::move(item));
            continue;
        }

        wchar_t textBuffer[256]{};
        const int textLen = GetMenuStringW(menu, pos, textBuffer, static_cast<int>(std::size(textBuffer)), MF_BYPOSITION);
        if (textLen > 0)
        {
            std::wstring_view fullText(textBuffer, static_cast<size_t>(textLen));
            const size_t tabPos               = fullText.find(L'\t');
            const std::wstring_view labelText = tabPos != std::wstring_view::npos ? fullText.substr(0, tabPos) : fullText;
            item.text                         = StripMenuMnemonicMarkers(labelText);
        }

        item.commandId = static_cast<int>(mii.wID);
        item.enabled   = (mii.fState & MFS_GRAYED) == 0;
        item.checked   = (mii.fState & MFS_CHECKED) != 0;

        if ((mii.fState & MFS_CHECKED) != 0)
            item.kind = MenuItemKind::Toggle;

        // Icon glyph
        item.iconGlyph = GetIconGlyphForCommand(mii.wID);

        // Accelerator text from ShortcutManager
        if (shortcutManager && mii.wID != 0u)
        {
            const auto commandIdOpt = TryGetCommandIdForContextMenuItem(mii.wID);
            if (commandIdOpt.has_value())
            {
                const auto chordOpt = shortcutManager->TryGetShortcutForCommand(commandIdOpt.value());
                if (chordOpt.has_value())
                    item.acceleratorText = FormatMenuChordText(chordOpt.value().vk, chordOpt.value().modifiers);
            }
        }

        // Recurse for submenus
        if (mii.hSubMenu)
            item.children = ConvertHMenuToFlyoutItems(mii.hSubMenu, shortcutManager);

        result.push_back(std::move(item));
    }

    return result;
}

} // namespace

void FolderView::OnContextMenu(POINT screenPt, const bool keyboardInvocation)
{
    if (! _hWnd)
        return;

    bool allowItemTarget = keyboardInvocation && _focusedIndex < _items.size();
    if (! keyboardInvocation)
    {
        const POINT clientPt = ScreenToClientPoint(screenPt);
        const std::optional<size_t> hit = HitTest(clientPt);
        if (hit.has_value())
        {
            allowItemTarget = true;
            if (hit.value() < _items.size() && ! _items[hit.value()].selected)
            {
                SelectSingle(hit.value());
            }
            else
            {
                FocusItem(hit.value(), false);
            }
            _anchorIndex = hit.value();
        }
        else
        {
            ClearSelection();
            _anchorIndex = _focusedIndex < _items.size() ? _focusedIndex : static_cast<size_t>(-1);
            DisarmPotentialDrag();
        }
    }

    HMENU rootMenu = Localization::LoadMenuResource(
        GetModuleHandleW(nullptr), allowItemTarget ? IDR_FOLDERVIEW_ITEM_CONTEXT : IDR_FOLDERVIEW_BACKGROUND_CONTEXT);
    if (! rootMenu)
        return;

    auto menuCleanup = wil::scope_exit([&] { DestroyMenu(rootMenu); });

    HMENU menu = GetSubMenu(rootMenu, 0);
    if (! menu)
        return;

    if (! allowItemTarget)
    {
        if (HMENU newMenu = FindSubMenuContainingCommandId(menu, IDM_PANE_CREATE_DIR))
        {
            PopulateShellNewTemplateMenuForFolderView(newMenu, _hWnd.get());
        }
    }

    const std::vector<std::wstring> selectionTargetSnapshot =
        allowItemTarget ? GetSelectedOrFocusedDisplayNames() : std::vector<std::wstring>{};
    const std::optional<std::wstring> currentTargetSnapshot =
        allowItemTarget && _focusedIndex < _items.size() ? std::optional<std::wstring>{_items[_focusedIndex].displayName} : std::nullopt;

    UpdateContextMenuState(menu, allowItemTarget);
    AppendArtifactContextMenu(menu, allowItemTarget ? GetFocusedArtifactProjection() : nullptr);
    if (! IsOverlaySampleEnabled())
    {
        RemoveOverlaySampleSubmenu(menu, CmdOverlaySampleError);
    }

    // Convert the HMENU to DxUI MenuFlyoutItems and show via D2D menu
    auto flyoutItems   = ConvertHMenuToFlyoutItems(menu, _shortcutManager);
    const auto palette = MakeAppThemeDxPalette(_appTheme);
    auto result        = RedSalamander::DxUi::ContextMenu::Show(_hWnd.get(), screenPt, flyoutItems, palette);
    if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
    {
        SetFocus(_hWnd.get());
    }
    if (result.has_value())
    {
        const UINT commandId = static_cast<UINT>(result.value());
        const bool itemTargetCommand = IsItemTargetContextCommand(commandId);
        const auto selectionSnapshotStillMatches = [&]() noexcept
        {
            const std::vector<std::wstring> currentTargets = GetSelectedOrFocusedDisplayNames();
            if (currentTargets.size() != selectionTargetSnapshot.size())
            {
                return false;
            }
            return std::is_permutation(selectionTargetSnapshot.begin(),
                                       selectionTargetSnapshot.end(),
                                       currentTargets.begin(),
                                       currentTargets.end(),
                                       [this](const std::wstring& expected, const std::wstring& current) noexcept
                                       { return EquivalentProviderComponent(expected, current); });
        };
        const auto currentSnapshotStillMatches = [&]() noexcept
        {
            return currentTargetSnapshot.has_value() && _focusedIndex < _items.size() &&
                EquivalentProviderComponent(currentTargetSnapshot.value(), _items[_focusedIndex].displayName);
        };
        const bool targetStillMatches = ! itemTargetCommand ||
            (allowItemTarget && (IsSelectionAwareContextCommand(commandId) ? selectionSnapshotStillMatches() : currentSnapshotStillMatches()));
        if (targetStillMatches)
        {
            PostMessageW(_hWnd.get(), WM_COMMAND, MAKEWPARAM(static_cast<WORD>(commandId), 0), 0);
        }
    }
}

void FolderView::UpdateContextMenuState(HMENU menu, const bool allowItemTarget) const
{
    if (! menu)
    {
        return;
    }

    const auto invalidIndex    = static_cast<size_t>(-1);
    const bool hasFocus        = _focusedIndex != invalidIndex && _focusedIndex < _items.size();
    const size_t selectedCount = static_cast<size_t>(std::count_if(_items.begin(), _items.end(), [](const FolderItem& item) { return item.selected; }));

    size_t effectiveCount = allowItemTarget ? selectedCount : 0u;
    if (allowItemTarget && effectiveCount == 0 && hasFocus)
    {
        effectiveCount = 1;
    }

    const bool hasTarget    = effectiveCount > 0;
    const bool singleTarget = effectiveCount == 1;

    auto setEnabled = [&](UINT command, bool enabled) { EnableMenuItem(menu, command, MF_BYCOMMAND | (enabled ? MF_ENABLED : static_cast<UINT>(MF_GRAYED))); };

    setEnabled(CmdOpen, allowItemTarget && hasFocus);
    setEnabled(CmdOpenWith, singleTarget && hasFocus);
    bool canViewSpace = false;
    if (_currentFolder.has_value())
    {
        canViewSpace = ! _currentFolder.value().empty();
    }
    setEnabled(CmdViewSpace, canViewSpace);
    setEnabled(CmdDelete, hasTarget);
    setEnabled(CmdMove, hasTarget);
    setEnabled(CmdRename, singleTarget && hasFocus);
    setEnabled(CmdCut, hasTarget && SupportsFileDropCut());
    setEnabled(CmdCopy, hasTarget);
    setEnabled(CmdProperties, singleTarget && hasFocus);

    setEnabled(CmdPaste, CanPasteItemsFromClipboard());
}
