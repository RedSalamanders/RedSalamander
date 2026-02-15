#include "Framework.h"

#include "ManagePluginsDialog.h"

#include <algorithm>
#include <array>
#include <cerrno>
#include <cwchar>
#include <filesystem>
#include <format>
#include <limits>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "FileSystemPluginManager.h"
#include "Helpers.h"
#include "HostServices.h"
#include "SettingsSave.h"
#include "SettingsSchemaExport.h"
#include "ThemedControls.h"
#include "ThemedInputFrames.h"
#include "ViewerPluginManager.h"
#include "resource.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include <commctrl.h>
#include <uxtheme.h>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

namespace
{
using ThemedControls::ApplyThemeToComboBox;
using ThemedControls::BlendColor;
using ThemedControls::DrawThemedPushButton;
using ThemedControls::DrawThemedSwitchToggle;
using ThemedControls::EnableOwnerDrawButton;
using ThemedControls::GetControlSurfaceColor;
using ThemedControls::MeasureTextWidth;
using ThemedControls::ScaleDip;

void ShowDialogAlert(HWND dlg, HostAlertSeverity severity, const std::wstring& title, const std::wstring& message) noexcept
{
    if (! dlg || message.empty())
    {
        return;
    }

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = severity;
    request.targetWindow = dlg;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(HostShowAlert(request));
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }

    return result;
}

#if 0
struct PluginListItem
{
    PluginType type = PluginType::FileSystem;
    size_t index    = 0;
};

struct DialogState
{
    DialogState()                              = default;
    DialogState(const DialogState&)            = delete;
    DialogState& operator=(const DialogState&) = delete;

    Common::Settings::Settings* settings       = nullptr;
    std::wstring appId;
    AppTheme theme{};
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> backgroundBrush;
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> selectionBrush;
    HWND list = nullptr;
    std::vector<PluginListItem> listItems;
};

void PaintPluginsListHeader(HWND header, const DialogState& state) noexcept
{
    if (! header)
    {
        return;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(header, &ps);

    RECT client{};
    if (! GetClientRect(header, &client))
    {
        return;
    }

    const HWND root         = GetAncestor(header, GA_ROOT);
    const bool windowActive = root && GetActiveWindow() == root;
    const COLORREF bg       = BlendColor(state.theme.windowBackground, state.theme.menu.separator, 1, 12);
    COLORREF textColor      = windowActive ? state.theme.menu.headerText : state.theme.menu.headerTextDisabled;
    if (textColor == bg)
    {
        textColor = ChooseContrastingTextColor(bg);
    }

    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> bgBrush(CreateSolidBrush(bg));
    FillRect(hdc.get(), &ps.rcPaint, bgBrush.get());

    HFONT fontToUse = reinterpret_cast<HFONT>(SendMessageW(header, WM_GETFONT, 0, 0));
    if (! fontToUse)
    {
        fontToUse = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), fontToUse);

    SetBkMode(hdc.get(), TRANSPARENT);
    SetTextColor(hdc.get(), textColor);

    const int dpi            = GetDeviceCaps(hdc.get(), LOGPIXELSX);
    const int paddingX       = MulDiv(8, dpi, USER_DEFAULT_SCREEN_DPI);

    const COLORREF lineColor = state.theme.menu.separator;
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> lineBrush(CreateSolidBrush(lineColor));

    const int count = Header_GetItemCount(header);
    for (int i = 0; i < count; ++i)
    {
        RECT rc{};
        if (Header_GetItemRect(header, i, &rc) == FALSE)
        {
            continue;
        }

        wchar_t buf[128]{};
        HDITEMW item{};
        item.mask       = HDI_TEXT | HDI_FORMAT;
        item.pszText    = buf;
        item.cchTextMax = static_cast<int>(std::size(buf));
        if (Header_GetItem(header, i, &item) == FALSE)
        {
            continue;
        }

        RECT textRect  = rc;
        textRect.left  = std::min(textRect.right, textRect.left + paddingX);
        textRect.right = std::max(textRect.left, textRect.right - paddingX);

        UINT flags     = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX;
        if ((item.fmt & HDF_RIGHT) != 0)
        {
            flags |= DT_RIGHT;
        }
        else if ((item.fmt & HDF_CENTER) != 0)
        {
            flags |= DT_CENTER;
        }
        else
        {
            flags |= DT_LEFT;
        }

        DrawTextW(hdc.get(), buf, static_cast<int>(std::wcslen(buf)), &textRect, flags);

        RECT rightLine = rc;
        rightLine.left = std::max(rightLine.left, rightLine.right - 1);
        FillRect(hdc.get(), &rightLine, lineBrush.get());
    }

    RECT bottomLine = client;
    bottomLine.top  = std::max(bottomLine.top, bottomLine.bottom - 1);
    FillRect(hdc.get(), &bottomLine, lineBrush.get());
}

LRESULT OnPluginsHeaderNcDestroy(HWND hwnd, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass) noexcept;

LRESULT CALLBACK PluginsHeaderSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    auto* state = reinterpret_cast<DialogState*>(dwRefData);
    if (! state || state->theme.highContrast)
    {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    switch (msg)
    {
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: PaintPluginsListHeader(hwnd, *state); return 0;
        case WM_NCDESTROY: return OnPluginsHeaderNcDestroy(hwnd, wParam, lParam, uIdSubclass);
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT OnPluginsHeaderNcDestroy(HWND hwnd, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass) noexcept
{
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C"' function
    RemoveWindowSubclass(hwnd, PluginsHeaderSubclassProc, uIdSubclass);
#pragma warning(pop)
    return DefSubclassProc(hwnd, WM_NCDESTROY, wParam, lParam);
}

std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }

    return result;
}

template <typename TOrigin> int OriginToGroupId(TOrigin origin) noexcept
{
    switch (origin)
    {
        case TOrigin::Embedded: return kPluginsGroupEmbedded;
        case TOrigin::Optional: return kPluginsGroupOptional;
        case TOrigin::Custom: return kPluginsGroupCustom;
    }
    return kPluginsGroupCustom;
}

template <typename TPluginEntry> std::wstring FormatPluginName(const TPluginEntry& entry)
{
    std::wstring name = entry.name.empty() ? entry.path.filename().wstring() : entry.name;

    if (entry.disabled)
    {
        name.append(L" ");
        name.append(LoadStringResource(nullptr, IDS_PLUGIN_SUFFIX_DISABLED));
    }
    else if (! entry.loadable)
    {
        name.append(L" ");
        name.append(LoadStringResource(nullptr, IDS_PLUGIN_SUFFIX_UNAVAILABLE));
    }

    return name;
}

template <typename TPluginEntry> std::wstring FormatPluginShortId(const TPluginEntry& entry)
{
    if (! entry.loadError.empty())
    {
        return entry.loadError;
    }

    if (! entry.shortId.empty())
    {
        return entry.shortId;
    }

    return entry.loadable ? std::wstring(L"(none)") : std::wstring(L"(unavailable)");
}

std::wstring PluginTypeToText(PluginType type)
{
    switch (type)
    {
        case PluginType::FileSystem: return LoadStringResource(nullptr, IDS_PLUGIN_TYPE_FILESYSTEM);
        case PluginType::Viewer: return LoadStringResource(nullptr, IDS_PLUGIN_TYPE_VIEWER);
    }

    return {};
}

void EnsureListColumns(HWND list, UINT dpi)
{
    ListView_DeleteAllItems(list);

    while (ListView_DeleteColumn(list, 0))
    {
    }

    RECT listRect{};
    GetClientRect(list, &listRect);

    const int totalWidth         = std::max(0l, listRect.right - listRect.left);
    const int minNameWidth       = MulDiv(120, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
    const int minTypeWidth       = MulDiv(80, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
    const int minShortWidth      = MulDiv(80, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);

    const int fallbackNameWidth  = MulDiv(160, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
    const int fallbackTypeWidth  = MulDiv(90, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
    const int fallbackShortWidth = MulDiv(120, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);

    int nameWidth                = fallbackNameWidth;
    int typeWidth                = fallbackTypeWidth;
    int shortWidth               = fallbackShortWidth;

    if (totalWidth > 0)
    {
        typeWidth  = std::min(typeWidth, totalWidth);
        shortWidth = std::min(shortWidth, std::max(0, totalWidth - typeWidth));

        typeWidth  = std::max(typeWidth, minTypeWidth);
        shortWidth = std::max(shortWidth, minShortWidth);

        nameWidth  = std::max(0, totalWidth - typeWidth - shortWidth);

        if (nameWidth < minNameWidth)
        {
            int deficit           = minNameWidth - nameWidth;
            const int shrinkShort = std::min(deficit, std::max(0, shortWidth - minShortWidth));
            shortWidth            = std::max(minShortWidth, shortWidth - shrinkShort);
            deficit -= shrinkShort;

            const int shrinkType = std::min(deficit, std::max(0, typeWidth - minTypeWidth));
            typeWidth            = std::max(minTypeWidth, typeWidth - shrinkType);
            deficit -= shrinkType;

            nameWidth = std::max(0, totalWidth - typeWidth - shortWidth);
        }
    }

    LVCOLUMNW col{};
    col.mask             = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;

    col.iSubItem         = kListColumnName;
    std::wstring colName = LoadStringResource(nullptr, IDS_PLUGIN_COLUMN_NAME);
    col.pszText          = colName.data();
    col.cx               = nameWidth;
    ListView_InsertColumn(list, kListColumnName, &col);

    col.iSubItem         = kListColumnType;
    std::wstring colType = LoadStringResource(nullptr, IDS_PLUGIN_COLUMN_TYPE);
    col.pszText          = colType.data();
    col.cx               = typeWidth;
    ListView_InsertColumn(list, kListColumnType, &col);

    col.iSubItem          = kListColumnShortId;
    std::wstring colShort = LoadStringResource(nullptr, IDS_PLUGIN_COLUMN_SHORT_ID);
    col.pszText           = colShort.data();
    col.cx                = shortWidth;
    ListView_InsertColumn(list, kListColumnShortId, &col);
}

void EnsureListGroups(HWND list)
{
    ListView_RemoveAllGroups(list);
    ListView_EnableGroupView(list, TRUE);

    LVGROUP group{};
    group.cbSize        = sizeof(group);
    group.mask          = LVGF_HEADER | LVGF_GROUPID;

    std::wstring header = LoadStringResource(nullptr, IDS_PLUGIN_GROUP_EMBEDDED);
    group.iGroupId      = kPluginsGroupEmbedded;
    group.pszHeader     = header.data();
    ListView_InsertGroup(list, -1, &group);

    header          = LoadStringResource(nullptr, IDS_PLUGIN_GROUP_OPTIONAL);
    group.iGroupId  = kPluginsGroupOptional;
    group.pszHeader = header.data();
    ListView_InsertGroup(list, -1, &group);

    header          = LoadStringResource(nullptr, IDS_PLUGIN_GROUP_CUSTOM);
    group.iGroupId  = kPluginsGroupCustom;
    group.pszHeader = header.data();
    ListView_InsertGroup(list, -1, &group);
}

void PopulatePluginsList([[maybe_unused]] HWND dlg, DialogState& state)
{
    if (! state.settings || ! state.list)
    {
        return;
    }

    ListView_DeleteAllItems(state.list);

    auto& fileSystemManager = FileSystemPluginManager::GetInstance();
    static_cast<void>(fileSystemManager.Refresh(*state.settings));

    auto& viewerManager = ViewerPluginManager::GetInstance();
    static_cast<void>(viewerManager.Refresh(*state.settings));

    const auto& fileSystemPlugins = fileSystemManager.GetPlugins();
    const auto& viewerPlugins     = viewerManager.GetPlugins();

    state.listItems.clear();
    state.listItems.reserve(fileSystemPlugins.size() + viewerPlugins.size());

    for (size_t i = 0; i < fileSystemPlugins.size(); ++i)
    {
        PluginListItem item;
        item.type  = PluginType::FileSystem;
        item.index = i;
        state.listItems.push_back(std::move(item));
    }

    for (size_t i = 0; i < viewerPlugins.size(); ++i)
    {
        PluginListItem item;
        item.type  = PluginType::Viewer;
        item.index = i;
        state.listItems.push_back(std::move(item));
    }

    const auto sortName = [](const auto& entry) { return entry.name.empty() ? entry.path.filename().wstring() : entry.name; };

    const auto compare  = [&](const PluginListItem& a, const PluginListItem& b)
    {
        int aOrigin = 0;
        int bOrigin = 0;

        std::wstring aName;
        std::wstring bName;

        std::wstring_view aId;
        std::wstring_view bId;

        if (a.type == PluginType::FileSystem)
        {
            const auto& entry = fileSystemPlugins[a.index];
            aOrigin           = static_cast<int>(entry.origin);
            aName             = sortName(entry);
            aId               = entry.id;
        }
        else
        {
            const auto& entry = viewerPlugins[a.index];
            aOrigin           = static_cast<int>(entry.origin);
            aName             = sortName(entry);
            aId               = entry.id;
        }

        if (b.type == PluginType::FileSystem)
        {
            const auto& entry = fileSystemPlugins[b.index];
            bOrigin           = static_cast<int>(entry.origin);
            bName             = sortName(entry);
            bId               = entry.id;
        }
        else
        {
            const auto& entry = viewerPlugins[b.index];
            bOrigin           = static_cast<int>(entry.origin);
            bName             = sortName(entry);
            bId               = entry.id;
        }

        if (aOrigin != bOrigin)
        {
            return aOrigin < bOrigin;
        }

        const int nameCmp = _wcsicmp(aName.c_str(), bName.c_str());
        if (nameCmp != 0)
        {
            return nameCmp < 0;
        }

        if (aId != bId)
        {
            return aId < bId;
        }

        return static_cast<int>(a.type) < static_cast<int>(b.type);
    };

    std::sort(state.listItems.begin(), state.listItems.end(), compare);

    for (size_t i = 0; i < state.listItems.size(); ++i)
    {
        const PluginListItem& listItem = state.listItems[i];

        int groupId                    = kPluginsGroupCustom;
        std::wstring displayName;
        std::wstring shortIdText;
        std::wstring typeText = PluginTypeToText(listItem.type);

        if (listItem.type == PluginType::FileSystem)
        {
            if (listItem.index >= fileSystemPlugins.size())
            {
                continue;
            }

            const auto& plugin = fileSystemPlugins[listItem.index];
            groupId            = OriginToGroupId(plugin.origin);
            displayName        = FormatPluginName(plugin);
            shortIdText        = FormatPluginShortId(plugin);
        }
        else
        {
            if (listItem.index >= viewerPlugins.size())
            {
                continue;
            }

            const auto& plugin = viewerPlugins[listItem.index];
            groupId            = OriginToGroupId(plugin.origin);
            displayName        = FormatPluginName(plugin);
            shortIdText        = FormatPluginShortId(plugin);
        }

        LVITEMW item{};
        item.mask       = LVIF_TEXT | LVIF_PARAM | LVIF_GROUPID;
        item.iItem      = ListView_GetItemCount(state.list);
        item.iSubItem   = 0;
        item.iGroupId   = groupId;
        item.lParam     = static_cast<LPARAM>(i);
        item.pszText    = displayName.data();

        const int index = ListView_InsertItem(state.list, &item);
        if (index >= 0)
        {
            ListView_SetItemText(state.list, index, kListColumnType, typeText.data());
            ListView_SetItemText(state.list, index, kListColumnShortId, shortIdText.data());
        }
    }
}

struct SelectedPlugin
{
    PluginType type                                        = PluginType::FileSystem;
    const FileSystemPluginManager::PluginEntry* fileSystem = nullptr;
    const ViewerPluginManager::PluginEntry* viewer         = nullptr;
};

std::optional<SelectedPlugin> FindSelectedPlugin(const DialogState& state)
{
    if (! state.list)
    {
        return std::nullopt;
    }

    const int selected = ListView_GetNextItem(state.list, -1, LVNI_SELECTED);
    if (selected < 0)
    {
        return std::nullopt;
    }

    LVITEMW item{};
    item.mask  = LVIF_PARAM;
    item.iItem = selected;
    if (! ListView_GetItem(state.list, &item))
    {
        return std::nullopt;
    }

    const size_t listIndex = static_cast<size_t>(item.lParam);
    if (listIndex >= state.listItems.size())
    {
        return std::nullopt;
    }

    const PluginListItem& listItem = state.listItems[listIndex];

    SelectedPlugin result;
    result.type = listItem.type;

    if (listItem.type == PluginType::FileSystem)
    {
        auto& manager       = FileSystemPluginManager::GetInstance();
        const auto& plugins = manager.GetPlugins();
        if (listItem.index >= plugins.size())
        {
            return std::nullopt;
        }
        result.fileSystem = &plugins[listItem.index];
    }
    else
    {
        auto& manager       = ViewerPluginManager::GetInstance();
        const auto& plugins = manager.GetPlugins();
        if (listItem.index >= plugins.size())
        {
            return std::nullopt;
        }
        result.viewer = &plugins[listItem.index];
    }

    return result;
}

void UpdateButtons(HWND dlg, const DialogState& state)
{
    const std::optional<SelectedPlugin> selected = FindSelectedPlugin(state);

    bool hasSelection                            = false;
    bool loadable                                = false;
    bool canRemove                               = false;
    std::wstring removeText                      = LoadStringResource(nullptr, IDS_PLUGIN_BUTTON_REMOVE);

    if (selected.has_value())
    {
        hasSelection = true;

        if (selected.value().type == PluginType::FileSystem && selected.value().fileSystem)
        {
            const auto& plugin = *selected.value().fileSystem;
            loadable           = plugin.loadable && ! plugin.id.empty();

            if (plugin.origin == FileSystemPluginManager::PluginOrigin::Custom)
            {
                canRemove  = true;
                removeText = LoadStringResource(nullptr, IDS_PLUGIN_BUTTON_REMOVE);
            }
            else if (! plugin.id.empty())
            {
                canRemove  = true;
                removeText = plugin.disabled ? LoadStringResource(nullptr, IDS_PLUGIN_BUTTON_ENABLE) : LoadStringResource(nullptr, IDS_PLUGIN_BUTTON_DISABLE);
            }
        }
        else if (selected.value().type == PluginType::Viewer && selected.value().viewer)
        {
            const auto& plugin = *selected.value().viewer;
            loadable           = plugin.loadable && ! plugin.id.empty();

            if (plugin.origin == ViewerPluginManager::PluginOrigin::Custom)
            {
                canRemove  = true;
                removeText = LoadStringResource(nullptr, IDS_PLUGIN_BUTTON_REMOVE);
            }
            else if (! plugin.id.empty())
            {
                canRemove  = true;
                removeText = plugin.disabled ? LoadStringResource(nullptr, IDS_PLUGIN_BUTTON_ENABLE) : LoadStringResource(nullptr, IDS_PLUGIN_BUTTON_DISABLE);
            }
        }
    }

    HWND removeButton = GetDlgItem(dlg, IDC_PLUGINS_REMOVE);
    if (! removeButton)
    {
        Debug::Error(L"UpdateButtons: Remove button not found");
        return;
    }

    if (! removeText.empty())
    {
        SetWindowTextW(removeButton, removeText.c_str());
    }

    EnableWindow(removeButton, canRemove);
    EnableWindow(GetDlgItem(dlg, IDC_PLUGINS_CONFIGURE), loadable);
    EnableWindow(GetDlgItem(dlg, IDC_PLUGINS_TEST), loadable);
    EnableWindow(GetDlgItem(dlg, IDC_PLUGINS_ABOUT), hasSelection);
}

void RemovePathFromVector(std::vector<std::filesystem::path>& values, const std::filesystem::path& needle)
{
    values.erase(std::remove_if(values.begin(), values.end(), [&](const std::filesystem::path& v) { return v == needle; }), values.end());
}

#endif
enum class PluginConfigFieldType : uint8_t
{
    Text,
    Value,
    Bool,
    Option,
    Selection,
};

struct PluginConfigChoice
{
    std::wstring value;
    std::wstring label;
};

struct PluginConfigField
{
    PluginConfigFieldType type = PluginConfigFieldType::Text;
    std::wstring key;
    std::wstring label;
    std::wstring description;

    bool hasMin = false;
    bool hasMax = false;
    int64_t min = 0;
    int64_t max = 0;

    std::wstring defaultText;
    int64_t defaultInt = 0;
    bool defaultBool   = false;
    std::wstring defaultOption;
    std::vector<std::wstring> defaultSelection;
    std::vector<PluginConfigChoice> choices;

    // UI metadata (optional x-ui-* attributes from schema)
    std::wstring uiSection; // x-ui-section: group fields under section headers
    int uiOrder = 0;        // x-ui-order: display order within plugin config dialog
    std::wstring uiControl; // x-ui-control: override control type (e.g., "custom" for future extensibility)
};

struct PluginConfigFieldControls
{
    PluginConfigField field;
    HWND hLabel                 = nullptr;
    HWND hEditFrame             = nullptr;
    HWND hEdit                  = nullptr;
    HWND hComboFrame            = nullptr;
    HWND hCombo                 = nullptr;
    HWND hToggle                = nullptr;
    HWND hComment               = nullptr;
    HWND hDefaults              = nullptr;
    size_t toggleOnChoiceIndex  = 0;
    size_t toggleOffChoiceIndex = 0;
    std::vector<HWND> choiceButtons;
};

enum class PluginConfigCommitMode : uint8_t
{
    ApplyToPluginsAndPersist,
    UpdateSettingsOnly,
};

struct PluginConfigDialogState
{
    PluginConfigDialogState()                                          = default;
    PluginConfigDialogState(const PluginConfigDialogState&)            = delete;
    PluginConfigDialogState& operator=(const PluginConfigDialogState&) = delete;

    Common::Settings::Settings* settings = nullptr;
    std::wstring appId;
    AppTheme theme{};
    PluginType pluginType = PluginType::FileSystem;
    std::wstring pluginId;
    std::wstring pluginName;
    std::string schemaJsonUtf8;
    std::string configurationJsonUtf8;
    PluginConfigCommitMode commitMode = PluginConfigCommitMode::ApplyToPluginsAndPersist;

    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> backgroundBrush;
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> inputBrush;
    COLORREF inputBackgroundColor = RGB(255, 255, 255);
    wil::unique_hfont commentFont;
    wil::unique_hfont boldFont;
    HWND panel             = nullptr;
    int contentHeight      = 0;
    int scrollPosY         = 0;
    int fixedWindowWidthPx = 0;
    std::vector<PluginConfigFieldControls> controls;
};

INT_PTR OnPluginConfigDialogCtlColorStatic(PluginConfigDialogState* state, HDC hdc, HWND control);
INT_PTR OnPluginConfigDialogCtlColorButton(PluginConfigDialogState* state, HDC hdc, HWND control);
INT_PTR OnPluginConfigDialogCtlColorEdit(PluginConfigDialogState* state, HDC hdc);
INT_PTR OnPluginConfigDialogCtlColorListBox(PluginConfigDialogState* state, HDC hdc);

void PersistSettings(HWND owner, Common::Settings::Settings& settings, std::wstring_view appId) noexcept
{
    if (appId.empty())
    {
        return;
    }

    const Common::Settings::Settings settingsToSave = SettingsSave::PrepareForSave(settings);
    const HRESULT hr                                = Common::Settings::SaveSettings(appId, settingsToSave);
    if (SUCCEEDED(hr))
    {
        const HRESULT schemaHr = SaveAggregatedSettingsSchema(appId, settings);
        if (FAILED(schemaHr))
        {
            Debug::Error(L"Failed to write aggregated settings schema (hr=0x{:08X})", static_cast<unsigned long>(schemaHr));
        }
        return;
    }

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(appId);
    Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(hr), settingsPath.wstring());

    if (! owner)
    {
        return;
    }

    const std::wstring message = FormatStringResource(nullptr, IDS_FMT_SETTINGS_SAVE_FAILED, settingsPath.wstring(), static_cast<unsigned long>(hr));
    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
    ShowDialogAlert(owner, HOST_ALERT_ERROR, title, message);
}

PluginConfigFieldType ParseFieldType(std::string_view type) noexcept
{
    if (type == "text")
    {
        return PluginConfigFieldType::Text;
    }
    if (type == "value")
    {
        return PluginConfigFieldType::Value;
    }
    if (type == "bool" || type == "boolean")
    {
        return PluginConfigFieldType::Bool;
    }
    if (type == "option")
    {
        return PluginConfigFieldType::Option;
    }
    if (type == "selection")
    {
        return PluginConfigFieldType::Selection;
    }
    return PluginConfigFieldType::Text;
}

std::optional<std::string_view> TryGetUtf8String(yyjson_val* obj, const char* key) noexcept
{
    if (! obj || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v || ! yyjson_is_str(v))
    {
        return std::nullopt;
    }

    const char* s = yyjson_get_str(v);
    if (! s)
    {
        return std::nullopt;
    }

    return std::string_view(s);
}

bool TryGetInt64(yyjson_val* obj, const char* key, int64_t& out) noexcept
{
    if (! obj || ! key)
    {
        return false;
    }

    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v)
    {
        return false;
    }

    if (yyjson_is_sint(v))
    {
        out = yyjson_get_sint(v);
        return true;
    }

    if (yyjson_is_uint(v))
    {
        out = static_cast<int64_t>(std::min<uint64_t>(yyjson_get_uint(v), static_cast<uint64_t>(std::numeric_limits<int64_t>::max())));
        return true;
    }

    if (yyjson_is_real(v))
    {
        out = static_cast<int64_t>(yyjson_get_real(v));
        return true;
    }

    return false;
}

[[nodiscard]] std::optional<bool> TryParseBoolToggleToken(std::wstring_view token) noexcept;

bool TryGetBoolValue(yyjson_val* obj, const char* key, bool& out) noexcept
{
    if (! obj || ! key)
    {
        return false;
    }

    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v)
    {
        return false;
    }

    if (yyjson_is_bool(v))
    {
        out = yyjson_get_bool(v);
        return true;
    }

    if (yyjson_is_sint(v))
    {
        out = yyjson_get_sint(v) != 0;
        return true;
    }

    if (yyjson_is_uint(v))
    {
        out = yyjson_get_uint(v) != 0;
        return true;
    }

    if (yyjson_is_str(v))
    {
        const char* s = yyjson_get_str(v);
        if (! s)
        {
            return false;
        }

        const std::optional<bool> parsed = TryParseBoolToggleToken(Utf16FromUtf8(s));
        if (parsed.has_value())
        {
            out = parsed.value();
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() > static_cast<size_t>(std::numeric_limits<int>::max()) || b.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    return CompareStringOrdinal(a.data(), static_cast<int>(a.size()), b.data(), static_cast<int>(b.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] std::optional<bool> TryParseBoolToggleToken(const std::wstring_view token) noexcept
{
    if (EqualsNoCase(token, L"on") || EqualsNoCase(token, L"true") || token == L"1")
    {
        return true;
    }

    if (EqualsNoCase(token, L"off") || EqualsNoCase(token, L"false") || token == L"0")
    {
        return false;
    }

    return std::nullopt;
}

[[nodiscard]] bool TryGetBoolToggleChoiceIndices(const PluginConfigField& field, size_t& outOnIndex, size_t& outOffIndex) noexcept
{
    if (field.type != PluginConfigFieldType::Option)
    {
        return false;
    }

    if (field.choices.size() != 2)
    {
        return false;
    }

    std::optional<size_t> onIndex;
    std::optional<size_t> offIndex;

    for (size_t i = 0; i < field.choices.size(); ++i)
    {
        const PluginConfigChoice& choice = field.choices[i];

        std::optional<bool> parsed = TryParseBoolToggleToken(choice.label);
        if (! parsed.has_value())
        {
            parsed = TryParseBoolToggleToken(choice.value);
        }

        if (! parsed.has_value())
        {
            continue;
        }

        if (parsed.value())
        {
            onIndex = i;
        }
        else
        {
            offIndex = i;
        }
    }

    if (! onIndex.has_value() || ! offIndex.has_value() || onIndex.value() == offIndex.value())
    {
        return false;
    }

    outOnIndex  = onIndex.value();
    outOffIndex = offIndex.value();
    return true;
}

[[nodiscard]] std::wstring_view TryGetChoiceLabelForValue(const PluginConfigField& field, std::wstring_view value) noexcept
{
    for (const auto& choice : field.choices)
    {
        if (choice.value == value)
        {
            return choice.label.empty() ? std::wstring_view(choice.value) : std::wstring_view(choice.label);
        }
    }
    return {};
}

std::wstring BuildFieldDefaultsTextForDisplay(const PluginConfigField& field)
{
    std::wstring defaults;

    switch (field.type)
    {
        case PluginConfigFieldType::Text:
        {
            if (! field.defaultText.empty())
            {
                defaults = std::format(L"Default: {}", field.defaultText);
            }
            break;
        }
        case PluginConfigFieldType::Value:
        {
            defaults = std::format(LocaleFormatting::GetFormatLocale(), L"Default: {:L}", field.defaultInt);
            if (field.hasMin)
            {
                defaults.append(std::format(LocaleFormatting::GetFormatLocale(), L"   Min: {:L}", field.min));
            }
            if (field.hasMax)
            {
                defaults.append(std::format(LocaleFormatting::GetFormatLocale(), L"   Max: {:L}", field.max));
            }
            break;
        }
        case PluginConfigFieldType::Bool:
        {
            defaults = std::format(L"Default: {}", field.defaultBool ? L"True" : L"False");
            break;
        }
        case PluginConfigFieldType::Option:
        {
            if (! field.defaultOption.empty())
            {
                const std::wstring_view label = TryGetChoiceLabelForValue(field, field.defaultOption);
                if (! label.empty())
                {
                    defaults = std::format(L"Default: {}", label);
                }
                else
                {
                    defaults = std::format(L"Default: {}", field.defaultOption);
                }
            }
            break;
        }
        case PluginConfigFieldType::Selection:
        {
            if (! field.defaultSelection.empty())
            {
                std::wstring joined;
                for (size_t i = 0; i < field.defaultSelection.size(); ++i)
                {
                    if (i > 0)
                    {
                        joined.append(L", ");
                    }

                    const std::wstring_view label = TryGetChoiceLabelForValue(field, field.defaultSelection[i]);
                    joined.append(label.empty() ? field.defaultSelection[i] : std::wstring(label));
                }
                defaults = std::format(L"Default: {}", joined);
            }
            break;
        }
    }

    return defaults;
}

int MeasureInfoHeight(HWND dlg, HFONT font, int width, const std::wstring& text) noexcept
{
    if (text.empty() || width <= 0)
    {
        return 0;
    }

    auto hdc = wil::GetDC(dlg);
    if (! hdc)
    {
        return 0;
    }

    RECT rc{};
    rc.right = width;

    HFONT useFont = font ? font : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    HGDIOBJ old   = nullptr;
    if (useFont)
    {
        old = SelectObject(hdc.get(), useFont);
    }

    DrawTextW(hdc.get(), text.c_str(), static_cast<int>(text.size()), &rc, DT_LEFT | DT_WORDBREAK | DT_CALCRECT);

    if (old)
    {
        SelectObject(hdc.get(), old);
    }

    return std::max(0l, rc.bottom - rc.top);
}

void EnsurePluginConfigCommentFont(PluginConfigDialogState& state, HFONT baseFont) noexcept
{
    if (state.commentFont)
    {
        return;
    }

    LOGFONTW lf{};
    if (baseFont && GetObjectW(baseFont, sizeof(lf), &lf) == sizeof(lf))
    {
        lf.lfItalic = TRUE;
        state.commentFont.reset(CreateFontIndirectW(&lf));
    }
}

void EnsurePluginConfigBoldFont(PluginConfigDialogState& state, HFONT baseFont) noexcept
{
    if (state.boldFont)
    {
        return;
    }

    LOGFONTW lf{};
    if (baseFont && GetObjectW(baseFont, sizeof(lf), &lf) == sizeof(lf))
    {
        lf.lfWeight = FW_BOLD;
        state.boldFont.reset(CreateFontIndirectW(&lf));
    }
}

void LayoutPluginConfigDialog(HWND dlg, PluginConfigDialogState& state) noexcept
{
    if (! dlg)
    {
        return;
    }

    HWND panel  = state.panel ? state.panel : GetDlgItem(dlg, IDC_PLUGIN_CONFIG_PLACEHOLDER);
    HWND ok     = GetDlgItem(dlg, IDOK);
    HWND cancel = GetDlgItem(dlg, IDCANCEL);
    if (! panel || ! ok || ! cancel)
    {
        return;
    }

    RECT client{};
    GetClientRect(dlg, &client);

    const UINT dpi   = GetDpiForWindow(dlg);
    const int margin = ScaleDip(dpi, 8);
    const int gapX   = ScaleDip(dpi, 8);

    RECT okRect{};
    RECT cancelRect{};
    GetWindowRect(ok, &okRect);
    GetWindowRect(cancel, &cancelRect);
    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&okRect), 2);
    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&cancelRect), 2);

    const int okWidth      = std::max(0l, okRect.right - okRect.left);
    const int okHeight     = std::max(0l, okRect.bottom - okRect.top);
    const int cancelWidth  = std::max(0l, cancelRect.right - cancelRect.left);
    const int cancelHeight = std::max(0l, cancelRect.bottom - cancelRect.top);
    const int buttonHeight = std::max(okHeight, cancelHeight);

    const int cancelLeft = std::max(0, static_cast<int>(client.right) - margin - cancelWidth);
    const int buttonsTop = std::max(0, static_cast<int>(client.bottom) - margin - buttonHeight);
    const int okLeft     = std::max(0, cancelLeft - gapX - okWidth);

    SetWindowPos(cancel, nullptr, cancelLeft, buttonsTop, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    SetWindowPos(ok, nullptr, okLeft, buttonsTop, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);

    const int panelLeft   = margin;
    const int panelTop    = margin;
    const int panelWidth  = std::max(0, static_cast<int>(client.right) - (2 * margin));
    const int panelBottom = std::max(panelTop, buttonsTop - margin);
    const int panelHeight = std::max(0, panelBottom - panelTop);

    SetWindowPos(panel, nullptr, panelLeft, panelTop, panelWidth, panelHeight, SWP_NOZORDER | SWP_NOACTIVATE);
}

void ScrollPanelTo(HWND panel, PluginConfigDialogState& state, int newScrollPosY) noexcept
{
    if (! panel)
    {
        return;
    }

    RECT client{};
    GetClientRect(panel, &client);
    const int clientHeight = std::max(0l, client.bottom - client.top);
    const int maxScroll    = std::max(0, state.contentHeight - clientHeight);

    newScrollPosY   = std::clamp(newScrollPosY, 0, maxScroll);
    const int delta = newScrollPosY - state.scrollPosY;
    if (delta == 0)
    {
        return;
    }

    state.scrollPosY = newScrollPosY;

    ScrollWindowEx(panel, 0, -delta, nullptr, nullptr, nullptr, nullptr, SW_INVALIDATE | SW_ERASE | SW_SCROLLCHILDREN);
    SetScrollPos(panel, SB_VERT, state.scrollPosY, TRUE);
    UpdateWindow(panel);
}

void UpdatePanelScrollInfo(HWND panel, PluginConfigDialogState& state) noexcept
{
    if (! panel)
    {
        return;
    }

    RECT client{};
    GetClientRect(panel, &client);
    const int clientHeight = std::max(0l, client.bottom - client.top);

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin   = 0;
    si.nMax   = std::max(0, state.contentHeight - 1);
    si.nPage  = static_cast<UINT>(std::max(0, clientHeight));
    si.nPos   = std::clamp(state.scrollPosY, 0, std::max(0, state.contentHeight - clientHeight));
    SetScrollInfo(panel, SB_VERT, &si, TRUE);

    ShowScrollBar(panel, SB_VERT, state.contentHeight > clientHeight);

    if (state.scrollPosY != si.nPos)
    {
        ScrollPanelTo(panel, state, si.nPos);
    }
}

void EnsurePanelChildVisible(HWND panel, PluginConfigDialogState& state, HWND child) noexcept
{
    if (! panel || ! child)
    {
        return;
    }

    RECT client{};
    GetClientRect(panel, &client);
    const int clientHeight = std::max(0l, client.bottom - client.top);
    if (clientHeight <= 0)
    {
        return;
    }

    RECT childRect{};
    if (! GetWindowRect(child, &childRect))
    {
        return;
    }

    MapWindowPoints(nullptr, panel, reinterpret_cast<POINT*>(&childRect), 2);

    const UINT dpi       = GetDpiForWindow(panel);
    const int padY       = ScaleDip(dpi, 8);
    const int viewTop    = std::max(0, static_cast<int>(client.top) + padY);
    const int viewBottom = std::max(viewTop, static_cast<int>(client.bottom) - padY);

    if (childRect.top < viewTop)
    {
        const int delta = childRect.top - viewTop;
        ScrollPanelTo(panel, state, state.scrollPosY + delta);
        UpdatePanelScrollInfo(panel, state);
        return;
    }

    if (childRect.bottom > viewBottom)
    {
        const int delta = childRect.bottom - viewBottom;
        ScrollPanelTo(panel, state, state.scrollPosY + delta);
        UpdatePanelScrollInfo(panel, state);
        return;
    }
}

PluginConfigFieldControls* FindToggleControls(PluginConfigDialogState& state, HWND toggle) noexcept
{
    if (! toggle)
    {
        return nullptr;
    }

    for (auto& c : state.controls)
    {
        if (c.hToggle == toggle)
        {
            return std::addressof(c);
        }
    }

    return nullptr;
}

void DrawPluginConfigToggle(const DRAWITEMSTRUCT& dis, const PluginConfigDialogState& state, const PluginConfigFieldControls& controls) noexcept
{
    const auto onIndex  = controls.toggleOnChoiceIndex;
    const auto offIndex = controls.toggleOffChoiceIndex;

    std::wstring_view onLabel;
    std::wstring_view offLabel;
    if (onIndex < controls.field.choices.size())
    {
        onLabel = controls.field.choices[onIndex].label;
    }
    if (offIndex < controls.field.choices.size())
    {
        offLabel = controls.field.choices[offIndex].label;
    }

    if (onLabel.empty())
    {
        onLabel = (controls.field.type == PluginConfigFieldType::Bool) ? L"True" : L"On";
    }
    if (offLabel.empty())
    {
        offLabel = (controls.field.type == PluginConfigFieldType::Bool) ? L"False" : L"Off";
    }

    const bool toggledOn = GetWindowLongPtrW(dis.hwndItem, GWLP_USERDATA) != 0;

    const COLORREF surface = state.inputBrush ? state.inputBackgroundColor : GetControlSurfaceColor(state.theme);
    const HFONT boldFont   = state.boldFont ? state.boldFont.get() : nullptr;
    DrawThemedSwitchToggle(dis, state.theme, surface, boldFont, onLabel, offLabel, toggledOn);
}

LRESULT CALLBACK PluginConfigInputControlSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData)
{
    return ThemedInputFrames::InputControlSubclassProc(hwnd, msg, wp, lp, subclassId, refData);
}

LRESULT CALLBACK PluginConfigInputFrameSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData)
{
    auto* state = reinterpret_cast<PluginConfigDialogState*>(refData);
    if (! state)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    ThemedInputFrames::FrameStyle frameStyle{};
    frameStyle.theme                = &state->theme;
    frameStyle.backdropBrush        = state->backgroundBrush.get();
    frameStyle.inputBackgroundColor = state->inputBackgroundColor;
    frameStyle.inputDisabledBackgroundColor =
        ThemedControls::BlendColor(state->theme.windowBackground, state->inputBackgroundColor, state->theme.dark ? 70 : 40, 255);

    return ThemedInputFrames::InputFrameSubclassProc(hwnd, msg, wp, lp, subclassId, reinterpret_cast<DWORD_PTR>(&frameStyle));
}

LRESULT CALLBACK PluginConfigPanelSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR /*subclassId*/, DWORD_PTR refData)
{
    auto* state = reinterpret_cast<PluginConfigDialogState*>(refData);
    if (! state)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    switch (msg)
    {
        case WM_ERASEBKGND:
        {
            RECT rc{};
            GetClientRect(hwnd, &rc);
            FillRect(reinterpret_cast<HDC>(wp), &rc, state->backgroundBrush.get());
            return 1;
        }
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (hdc)
            {
                FillRect(hdc.get(), &ps.rcPaint, state->backgroundBrush.get());
            }
            return 0;
        }
        case WM_SETFOCUS:
        {
            // Move focus into the first input control when tabbing into the scroll panel.
            const bool backwards = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
            HWND first           = GetNextDlgTabItem(hwnd, nullptr, backwards ? TRUE : FALSE);
            if (first)
            {
                SetFocus(first);
            }
            return 0;
        }
        case WM_CTLCOLORSTATIC: return OnPluginConfigDialogCtlColorStatic(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLORBTN: return OnPluginConfigDialogCtlColorButton(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLOREDIT: return OnPluginConfigDialogCtlColorEdit(state, reinterpret_cast<HDC>(wp));
        case WM_CTLCOLORLISTBOX: return OnPluginConfigDialogCtlColorListBox(state, reinterpret_cast<HDC>(wp));
        case WM_DRAWITEM:
        {
            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lp);
            if (! dis)
            {
                break;
            }

            if (dis->CtlType != ODT_BUTTON)
            {
                break;
            }

            if (auto* controls = FindToggleControls(*state, dis->hwndItem))
            {
                DrawPluginConfigToggle(*dis, *state, *controls);
                return TRUE;
            }

            break;
        }
        case WM_COMMAND:
        {
            const UINT notify = HIWORD(wp);
            if (notify == BN_SETFOCUS || notify == EN_SETFOCUS || notify == CBN_SETFOCUS)
            {
                HWND focusedControl = reinterpret_cast<HWND>(lp);
                if (focusedControl)
                {
                    EnsurePanelChildVisible(hwnd, *state, focusedControl);
                }
            }

            if (HIWORD(wp) == BN_CLICKED)
            {
                HWND clicked = reinterpret_cast<HWND>(lp);
                if (auto* controls = FindToggleControls(*state, clicked))
                {
                    const LONG_PTR current = GetWindowLongPtrW(clicked, GWLP_USERDATA);
                    SetWindowLongPtrW(clicked, GWLP_USERDATA, current == 0 ? 1 : 0);
                    InvalidateRect(clicked, nullptr, TRUE);
                    return 0;
                }
            }
            break;
        }
        case WM_SIZE:
        {
            UpdatePanelScrollInfo(hwnd, *state);
            break;
        }
        case WM_VSCROLL:
        {
            const UINT action  = LOWORD(wp);
            const UINT dpi     = GetDpiForWindow(hwnd);
            const int lineStep = std::max(1, ScaleDip(dpi, 18));

            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_ALL;
            if (! GetScrollInfo(hwnd, SB_VERT, &si))
            {
                break;
            }

            int newPos = si.nPos;
            switch (action)
            {
                case SB_LINEUP: newPos -= lineStep; break;
                case SB_LINEDOWN: newPos += lineStep; break;
                case SB_PAGEUP: newPos -= static_cast<int>(si.nPage); break;
                case SB_PAGEDOWN: newPos += static_cast<int>(si.nPage); break;
                case SB_THUMBPOSITION:
                case SB_THUMBTRACK: newPos = si.nTrackPos; break;
                case SB_TOP: newPos = si.nMin; break;
                case SB_BOTTOM: newPos = si.nMax; break;
            }

            ScrollPanelTo(hwnd, *state, newPos);
            UpdatePanelScrollInfo(hwnd, *state);
            return 0;
        }
        case WM_MOUSEWHEEL:
        {
            const int wheelDelta = GET_WHEEL_DELTA_WPARAM(wp);
            if (wheelDelta == 0)
            {
                break;
            }

            UINT lines = 3;
            SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
            if (lines == WHEEL_PAGESCROLL)
            {
                lines = 3;
            }

            const UINT dpi     = GetDpiForWindow(hwnd);
            const int lineStep = std::max(1, ScaleDip(dpi, 18));
            const int steps    = wheelDelta / WHEEL_DELTA;
            if (steps != 0)
            {
                ScrollPanelTo(hwnd, *state, state->scrollPosY - (steps * static_cast<int>(lines) * lineStep));
                UpdatePanelScrollInfo(hwnd, *state);
                return 0;
            }
            break;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

std::vector<PluginConfigField> ParseConfigurationSchema(std::string_view schemaJsonUtf8)
{
    std::vector<PluginConfigField> fields;

    if (schemaJsonUtf8.empty())
    {
        return fields;
    }

    yyjson_doc* doc = yyjson_read(schemaJsonUtf8.data(), schemaJsonUtf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
    if (! doc)
    {
        return fields;
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return fields;
    }

    yyjson_val* fieldsArr = yyjson_obj_get(root, "fields");
    if (! fieldsArr || ! yyjson_is_arr(fieldsArr))
    {
        return fields;
    }

    const size_t count = yyjson_arr_size(fieldsArr);
    fields.reserve(count);

    for (size_t i = 0; i < count; ++i)
    {
        yyjson_val* item = yyjson_arr_get(fieldsArr, i);
        if (! item || ! yyjson_is_obj(item))
        {
            continue;
        }

        const auto keyUtf8  = TryGetUtf8String(item, "key");
        const auto typeUtf8 = TryGetUtf8String(item, "type");
        if (! keyUtf8.has_value() || ! typeUtf8.has_value())
        {
            continue;
        }

        PluginConfigField field;
        field.key = Utf16FromUtf8(keyUtf8.value());
        if (field.key.empty())
        {
            continue;
        }

        field.type = ParseFieldType(typeUtf8.value());

        const auto labelUtf8 = TryGetUtf8String(item, "label");
        field.label          = labelUtf8.has_value() ? Utf16FromUtf8(labelUtf8.value()) : field.key;
        if (field.label.empty())
        {
            field.label = field.key;
        }

        const auto descriptionUtf8 = TryGetUtf8String(item, "description");
        if (descriptionUtf8.has_value())
        {
            field.description = Utf16FromUtf8(descriptionUtf8.value());
        }

        // Parse x-ui-* attributes for UI customization
        const auto uiSectionUtf8 = TryGetUtf8String(item, "x-ui-section");
        if (uiSectionUtf8.has_value())
        {
            field.uiSection = Utf16FromUtf8(uiSectionUtf8.value());
        }

        int64_t uiOrderValue = 0;
        if (TryGetInt64(item, "x-ui-order", uiOrderValue))
        {
            field.uiOrder = static_cast<int>(uiOrderValue);
        }

        const auto uiControlUtf8 = TryGetUtf8String(item, "x-ui-control");
        if (uiControlUtf8.has_value())
        {
            field.uiControl = Utf16FromUtf8(uiControlUtf8.value());
        }

        int64_t minValue = 0;
        if (TryGetInt64(item, "min", minValue))
        {
            field.hasMin = true;
            field.min    = minValue;
        }

        int64_t maxValue = 0;
        if (TryGetInt64(item, "max", maxValue))
        {
            field.hasMax = true;
            field.max    = maxValue;
        }

        if (field.type == PluginConfigFieldType::Text)
        {
            const auto def    = TryGetUtf8String(item, "default");
            field.defaultText = def.has_value() ? Utf16FromUtf8(def.value()) : std::wstring();
        }
        else if (field.type == PluginConfigFieldType::Value)
        {
            int64_t defValue = 0;
            if (TryGetInt64(item, "default", defValue))
            {
                field.defaultInt = defValue;
            }
        }
        else if (field.type == PluginConfigFieldType::Bool)
        {
            bool def = false;
            if (TryGetBoolValue(item, "default", def))
            {
                field.defaultBool = def;
            }
        }
        else if (field.type == PluginConfigFieldType::Option)
        {
            const auto def      = TryGetUtf8String(item, "default");
            field.defaultOption = def.has_value() ? Utf16FromUtf8(def.value()) : std::wstring();

            yyjson_val* options = yyjson_obj_get(item, "options");
            if (options && yyjson_is_arr(options))
            {
                const size_t optCount = yyjson_arr_size(options);
                field.choices.reserve(optCount);
                for (size_t o = 0; o < optCount; ++o)
                {
                    yyjson_val* opt = yyjson_arr_get(options, o);
                    if (! opt || ! yyjson_is_obj(opt))
                    {
                        continue;
                    }

                    const auto valueUtf8 = TryGetUtf8String(opt, "value");
                    if (! valueUtf8.has_value())
                    {
                        continue;
                    }

                    PluginConfigChoice choice;
                    choice.value = Utf16FromUtf8(valueUtf8.value());
                    if (choice.value.empty())
                    {
                        continue;
                    }

                    const auto optLabelUtf8 = TryGetUtf8String(opt, "label");
                    choice.label            = optLabelUtf8.has_value() ? Utf16FromUtf8(optLabelUtf8.value()) : choice.value;
                    if (choice.label.empty())
                    {
                        choice.label = choice.value;
                    }

                    field.choices.push_back(std::move(choice));
                }
            }
        }
        else if (field.type == PluginConfigFieldType::Selection)
        {
            yyjson_val* options = yyjson_obj_get(item, "options");
            if (options && yyjson_is_arr(options))
            {
                const size_t optCount = yyjson_arr_size(options);
                field.choices.reserve(optCount);
                for (size_t o = 0; o < optCount; ++o)
                {
                    yyjson_val* opt = yyjson_arr_get(options, o);
                    if (! opt || ! yyjson_is_obj(opt))
                    {
                        continue;
                    }

                    const auto valueUtf8 = TryGetUtf8String(opt, "value");
                    if (! valueUtf8.has_value())
                    {
                        continue;
                    }

                    PluginConfigChoice choice;
                    choice.value = Utf16FromUtf8(valueUtf8.value());
                    if (choice.value.empty())
                    {
                        continue;
                    }

                    const auto optLabelUtf8 = TryGetUtf8String(opt, "label");
                    choice.label            = optLabelUtf8.has_value() ? Utf16FromUtf8(optLabelUtf8.value()) : choice.value;
                    if (choice.label.empty())
                    {
                        choice.label = choice.value;
                    }

                    field.choices.push_back(std::move(choice));
                }
            }

            yyjson_val* def = yyjson_obj_get(item, "default");
            if (def && yyjson_is_arr(def))
            {
                const size_t defCount = yyjson_arr_size(def);
                field.defaultSelection.reserve(defCount);
                for (size_t d = 0; d < defCount; ++d)
                {
                    yyjson_val* v = yyjson_arr_get(def, d);
                    if (! v || ! yyjson_is_str(v))
                    {
                        continue;
                    }

                    const char* s = yyjson_get_str(v);
                    if (! s)
                    {
                        continue;
                    }

                    std::wstring value = Utf16FromUtf8(s);
                    if (! value.empty())
                    {
                        field.defaultSelection.push_back(std::move(value));
                    }
                }
            }
        }

        fields.push_back(std::move(field));
    }

    // Sort fields by uiOrder (if specified), then by original order
    std::stable_sort(fields.begin(),
                     fields.end(),
                     [](const PluginConfigField& a, const PluginConfigField& b)
                     {
                         // Fields with explicit uiOrder come first, sorted by order value
                         // Fields without uiOrder (order=0) maintain original order via stable_sort
                         if (a.uiOrder != 0 && b.uiOrder != 0)
                         {
                             return a.uiOrder < b.uiOrder;
                         }
                         if (a.uiOrder != 0)
                         {
                             return true; // a has order, b doesn't → a comes first
                         }
                         if (b.uiOrder != 0)
                         {
                             return false; // b has order, a doesn't → b comes first
                         }
                         return false; // both have no order → maintain original order
                     });

    return fields;
}

yyjson_doc* ParseJsonToDoc(std::string_view textUtf8) noexcept
{
    if (textUtf8.empty())
    {
        return nullptr;
    }

    return yyjson_read(textUtf8.data(), textUtf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
}

std::string BuildConfigurationJson(const std::vector<PluginConfigFieldControls>& controls)
{
    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return {};
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    if (! root)
    {
        return {};
    }
    yyjson_mut_doc_set_root(doc, root);

    for (const auto& c : controls)
    {
        const std::string keyUtf8 = Utf8FromUtf16(c.field.key);
        if (keyUtf8.empty() && ! c.field.key.empty())
        {
            continue;
        }

        if (keyUtf8.empty())
        {
            continue;
        }

        yyjson_mut_val* key = yyjson_mut_strncpy(doc, keyUtf8.c_str(), keyUtf8.size());
        if (! key)
        {
            return {};
        }

        if (c.field.type == PluginConfigFieldType::Text)
        {
            std::wstring value;
            if (c.hEdit)
            {
                const int len = GetWindowTextLengthW(c.hEdit);
                if (len > 0)
                {
                    value.resize(static_cast<size_t>(len) + 1u);
                    GetWindowTextW(c.hEdit, value.data(), len + 1);
                    value.resize(static_cast<size_t>(len));
                }
            }

            const std::string utf8 = Utf8FromUtf16(value);
            yyjson_mut_val* val    = yyjson_mut_strncpy(doc, utf8.c_str(), utf8.size());
            if (! val)
            {
                return {};
            }
            if (! yyjson_mut_obj_add(root, key, val))
            {
                return {};
            }
        }
        else if (c.field.type == PluginConfigFieldType::Value)
        {
            int64_t v = c.field.defaultInt;
            if (c.hEdit)
            {
                std::array<wchar_t, 64> buffer{};
                GetWindowTextW(c.hEdit, buffer.data(), static_cast<int>(buffer.size()));

                wchar_t* end           = nullptr;
                errno                  = 0;
                const long long parsed = wcstoll(buffer.data(), &end, 10);
                if (errno == 0 && end != buffer.data())
                {
                    v = static_cast<int64_t>(parsed);
                }
            }

            if (c.field.hasMin)
            {
                v = std::max(v, c.field.min);
            }
            if (c.field.hasMax)
            {
                v = std::min(v, c.field.max);
            }

            yyjson_mut_val* val = yyjson_mut_int(doc, v);
            if (! val)
            {
                return {};
            }
            if (! yyjson_mut_obj_add(root, key, val))
            {
                return {};
            }
        }
        else if (c.field.type == PluginConfigFieldType::Bool)
        {
            bool v = c.field.defaultBool;
            if (c.hToggle)
            {
                const LONG_PTR style = GetWindowLongPtrW(c.hToggle, GWL_STYLE);
                const LONG_PTR type  = style & BS_TYPEMASK;
                if (type == BS_OWNERDRAW)
                {
                    v = GetWindowLongPtrW(c.hToggle, GWLP_USERDATA) != 0;
                }
                else
                {
                    v = SendMessageW(c.hToggle, BM_GETCHECK, 0, 0) == BST_CHECKED;
                }
            }
            else if (! c.choiceButtons.empty())
            {
                v = SendMessageW(c.choiceButtons.front(), BM_GETCHECK, 0, 0) == BST_CHECKED;
            }

            yyjson_mut_val* val = yyjson_mut_bool(doc, v ? true : false);
            if (! val)
            {
                return {};
            }
            if (! yyjson_mut_obj_add(root, key, val))
            {
                return {};
            }
        }
        else if (c.field.type == PluginConfigFieldType::Option)
        {
            std::wstring selected;
            if (c.hToggle)
            {
                const bool isOn    = GetWindowLongPtrW(c.hToggle, GWLP_USERDATA) != 0;
                const size_t index = isOn ? c.toggleOnChoiceIndex : c.toggleOffChoiceIndex;
                if (index < c.field.choices.size())
                {
                    selected = c.field.choices[index].value;
                }
            }
            else if (c.hCombo)
            {
                const LRESULT index = SendMessageW(c.hCombo, CB_GETCURSEL, 0, 0);
                if (index >= 0 && index <= static_cast<LRESULT>(std::numeric_limits<int>::max()))
                {
                    const size_t choiceIndex = static_cast<size_t>(index);
                    if (choiceIndex < c.field.choices.size())
                    {
                        selected = c.field.choices[choiceIndex].value;
                    }
                }
            }
            else
            {
                for (size_t i = 0; i < c.choiceButtons.size() && i < c.field.choices.size(); ++i)
                {
                    if (SendMessageW(c.choiceButtons[i], BM_GETCHECK, 0, 0) == BST_CHECKED)
                    {
                        selected = c.field.choices[i].value;
                        break;
                    }
                }
            }

            const std::string utf8 = Utf8FromUtf16(selected);
            yyjson_mut_val* val    = yyjson_mut_strncpy(doc, utf8.c_str(), utf8.size());
            if (! val)
            {
                return {};
            }
            if (! yyjson_mut_obj_add(root, key, val))
            {
                return {};
            }
        }
        else if (c.field.type == PluginConfigFieldType::Selection)
        {
            yyjson_mut_val* arr = yyjson_mut_arr(doc);
            if (! arr)
            {
                return {};
            }
            if (! yyjson_mut_obj_add(root, key, arr))
            {
                return {};
            }

            for (size_t i = 0; i < c.choiceButtons.size() && i < c.field.choices.size(); ++i)
            {
                if (SendMessageW(c.choiceButtons[i], BM_GETCHECK, 0, 0) != BST_CHECKED)
                {
                    continue;
                }

                const std::string utf8 = Utf8FromUtf16(c.field.choices[i].value);
                yyjson_mut_val* val    = yyjson_mut_strncpy(doc, utf8.c_str(), utf8.size());
                if (! val)
                {
                    return {};
                }
                if (! yyjson_mut_arr_add_val(arr, val))
                {
                    return {};
                }
            }
        }
    }

    yyjson_write_err err{};
    size_t len = 0;
    wil::unique_any<char*, decltype(&::free), ::free> json(yyjson_mut_write_opts(doc, YYJSON_WRITE_NOFLAG, nullptr, &len, &err));
    if (! json || len == 0)
    {
        Debug::Error(L"Failed to serialize plugin configuration to JSON: code: {}", err.code);
        return {};
    }

    return std::string(json.get(), len);
}

void ApplyFieldDefaultToControls(const PluginConfigField& field, PluginConfigFieldControls& out, yyjson_val* configRoot)
{
    out.field = field;

    const std::string keyUtf8 = Utf8FromUtf16(field.key);
    yyjson_val* current       = nullptr;
    if (configRoot && ! keyUtf8.empty())
    {
        current = yyjson_obj_get(configRoot, keyUtf8.c_str());
    }

    if (field.type == PluginConfigFieldType::Text)
    {
        std::wstring value = field.defaultText;
        if (current && yyjson_is_str(current))
        {
            const char* s = yyjson_get_str(current);
            if (s)
            {
                value = Utf16FromUtf8(s);
            }
        }

        out.field.defaultText = value;
    }
    else if (field.type == PluginConfigFieldType::Value)
    {
        int64_t value = field.defaultInt;
        if (current)
        {
            if (yyjson_is_sint(current))
            {
                value = yyjson_get_sint(current);
            }
            else if (yyjson_is_uint(current))
            {
                value = static_cast<int64_t>(std::min<uint64_t>(yyjson_get_uint(current), static_cast<uint64_t>(std::numeric_limits<int64_t>::max())));
            }
            else if (yyjson_is_real(current))
            {
                value = static_cast<int64_t>(yyjson_get_real(current));
            }
        }

        out.field.defaultInt = value;
    }
    else if (field.type == PluginConfigFieldType::Bool)
    {
        bool value = field.defaultBool;
        if (current)
        {
            if (yyjson_is_bool(current))
            {
                value = yyjson_get_bool(current);
            }
            else if (yyjson_is_sint(current))
            {
                value = yyjson_get_sint(current) != 0;
            }
            else if (yyjson_is_uint(current))
            {
                value = yyjson_get_uint(current) != 0;
            }
            else if (yyjson_is_str(current))
            {
                const char* s = yyjson_get_str(current);
                if (s)
                {
                    const std::optional<bool> parsed = TryParseBoolToggleToken(Utf16FromUtf8(s));
                    if (parsed.has_value())
                    {
                        value = parsed.value();
                    }
                }
            }
        }

        out.field.defaultBool = value;
    }
    else if (field.type == PluginConfigFieldType::Option)
    {
        std::wstring value = field.defaultOption;
        if (current && yyjson_is_str(current))
        {
            const char* s = yyjson_get_str(current);
            if (s)
            {
                value = Utf16FromUtf8(s);
            }
        }

        out.field.defaultOption = value;
    }
    else if (field.type == PluginConfigFieldType::Selection)
    {
        std::vector<std::wstring> values = field.defaultSelection;
        if (current && yyjson_is_arr(current))
        {
            values.clear();
            const size_t count = yyjson_arr_size(current);
            values.reserve(count);
            for (size_t i = 0; i < count; ++i)
            {
                yyjson_val* v = yyjson_arr_get(current, i);
                if (! v || ! yyjson_is_str(v))
                {
                    continue;
                }
                const char* s = yyjson_get_str(v);
                if (! s)
                {
                    continue;
                }
                std::wstring t = Utf16FromUtf8(s);
                if (! t.empty())
                {
                    values.push_back(std::move(t));
                }
            }
        }

        out.field.defaultSelection = std::move(values);
    }
}

INT_PTR OnPluginConfigDialogInit(HWND dlg, PluginConfigDialogState* state)
{
    if (! state)
    {
        return FALSE;
    }

    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));

    ApplyTitleBarTheme(dlg, state->theme, GetActiveWindow() == dlg);

    state->backgroundBrush.reset(CreateSolidBrush(state->theme.windowBackground));
    state->inputBackgroundColor = GetControlSurfaceColor(state->theme);
    state->inputBrush.reset();
    if (! state->theme.highContrast)
    {
        state->inputBrush.reset(CreateSolidBrush(state->inputBackgroundColor));
    }
    state->contentHeight = 0;
    state->scrollPosY    = 0;

    RECT dlgRect{};
    if (GetWindowRect(dlg, &dlgRect))
    {
        state->fixedWindowWidthPx = std::max(0l, dlgRect.right - dlgRect.left);
    }

    if (! state->pluginName.empty())
    {
        SetWindowTextW(dlg, state->pluginName.c_str());
    }

    if (! state->theme.highContrast)
    {
        EnableOwnerDrawButton(dlg, IDOK);
        EnableOwnerDrawButton(dlg, IDCANCEL);
    }

    state->panel = GetDlgItem(dlg, IDC_PLUGIN_CONFIG_PLACEHOLDER);
    if (state->panel)
    {
        LONG_PTR exStyle = GetWindowLongPtrW(state->panel, GWL_EXSTYLE);
        if ((exStyle & WS_EX_CONTROLPARENT) == 0)
        {
            exStyle |= WS_EX_CONTROLPARENT;
            SetWindowLongPtrW(state->panel, GWL_EXSTYLE, exStyle);
        }

        LONG_PTR style = GetWindowLongPtrW(state->panel, GWL_STYLE);
        if ((style & WS_TABSTOP) == 0 || (style & SS_NOTIFY) == 0)
        {
            style |= WS_TABSTOP | SS_NOTIFY;
            SetWindowLongPtrW(state->panel, GWL_STYLE, style);
            SetWindowPos(state->panel, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        }

        const bool darkBackground = ChooseContrastingTextColor(state->theme.windowBackground) == RGB(255, 255, 255);
        const wchar_t* panelTheme = state->theme.highContrast ? L"" : (darkBackground ? L"DarkMode_Explorer" : L"Explorer");
        SetWindowTheme(state->panel, panelTheme, nullptr);
        SendMessageW(state->panel, WM_THEMECHANGED, 0, 0);

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(state->panel, PluginConfigPanelSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(state));
#pragma warning(pop)
    }

    LayoutPluginConfigDialog(dlg, *state);

    const std::vector<PluginConfigField> fields = ParseConfigurationSchema(state->schemaJsonUtf8);

    yyjson_doc* configDoc  = ParseJsonToDoc(state->configurationJsonUtf8);
    yyjson_val* configRoot = nullptr;
    if (configDoc)
    {
        configRoot = yyjson_doc_get_root(configDoc);
        if (! configRoot || ! yyjson_is_obj(configRoot))
        {
            configRoot = nullptr;
        }
    }

    state->controls.clear();
    state->controls.reserve(fields.size());

    HWND panel = state->panel;
    if (! panel)
    {
        if (configDoc)
        {
            yyjson_doc_free(configDoc);
        }
        return TRUE;
    }

    const UINT dpi = GetDpiForWindow(dlg);

    int margin = ScaleDip(dpi, 8);

    int spacingY = ScaleDip(dpi, 10);

    int labelOffsetY = ScaleDip(dpi, 3);

    int labelGapX = ScaleDip(dpi, 10);

    int labelHeight = std::max(1, ScaleDip(dpi, 18));

    int editHeight = std::max(1, ScaleDip(dpi, 26));

    int optionHeight = std::max(1, ScaleDip(dpi, 20));

    int minControlWidth = ScaleDip(dpi, 80);

    RECT panelRect{};
    GetClientRect(panel, &panelRect);
    int panelWidth = std::max(0l, panelRect.right - panelRect.left);

    // Reserve space for the vertical scrollbar. ShowScrollBar does not shrink the client area for us
    // (see e.g. ViewerImgRaw), so controls would otherwise draw under the scrollbar when content overflows.
    const int scrollW = GetSystemMetricsForDpi(SM_CXVSCROLL, dpi);
    panelWidth        = std::max(0, panelWidth - scrollW);

    [[maybe_unused]] const int contentWidth = std::max(0, panelWidth - (2 * margin));

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(dlg, WM_GETFONT, 0, 0));

    EnsurePluginConfigCommentFont(*state, font);
    EnsurePluginConfigBoldFont(*state, font);

    const int left  = margin;
    const int top   = margin;
    const int right = std::max(0, panelWidth - margin);

    const int availableWidth = std::max(0, right - left);

    int labelWidth          = (availableWidth * 2) / 5;
    const int minLabelWidth = ScaleDip(dpi, 110);
    labelWidth              = std::clamp(labelWidth, minLabelWidth, std::max(0, availableWidth - minControlWidth));

    const int labelTextWidth = std::max(0, labelWidth - labelGapX);

    const int controlX     = left + labelWidth;
    const int controlWidth = std::max(minControlWidth, right - controlX);

    int y = top;

    for (const auto& field : fields)
    {
        PluginConfigFieldControls controls;
        ApplyFieldDefaultToControls(field, controls, configRoot);

        controls.hLabel = CreateWindowExW(0,
                                          L"Static",
                                          controls.field.label.c_str(),
                                          WS_CHILD | WS_VISIBLE | SS_NOPREFIX | SS_WORDELLIPSIS,
                                          left,
                                          y + labelOffsetY,
                                          labelTextWidth,
                                          labelHeight,
                                          panel,
                                          nullptr,
                                          GetModuleHandleW(nullptr),
                                          nullptr);
        if (controls.hLabel && font)
        {
            SendMessageW(controls.hLabel, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
        }

        if (controls.field.type == PluginConfigFieldType::Text || controls.field.type == PluginConfigFieldType::Value)
        {
            const int valueWidth     = std::min(controlWidth, ScaleDip(dpi, 140));
            const int editFrameWidth = (controls.field.type == PluginConfigFieldType::Value) ? valueWidth : controlWidth;

            const bool customFrames = ! state->theme.highContrast;
            const int framePadding  = ScaleDip(dpi, 2);
            const int textMargin    = ScaleDip(dpi, 6);

            DWORD editStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL;
            if (controls.field.type == PluginConfigFieldType::Value)
            {
                editStyle |= ES_NUMBER;
            }

            if (customFrames)
            {
                controls.hEditFrame = CreateWindowExW(
                    0, L"Static", L"", WS_CHILD | WS_VISIBLE, controlX, y, editFrameWidth, editHeight, panel, nullptr, GetModuleHandleW(nullptr), nullptr);

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
                if (controls.hEditFrame)
                {
                    SetWindowSubclass(controls.hEditFrame, PluginConfigInputFrameSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(state));
                }
#pragma warning(pop)

                controls.hEdit = CreateWindowExW(0,
                                                 L"Edit",
                                                 L"",
                                                 editStyle,
                                                 controlX + framePadding,
                                                 y + framePadding,
                                                 std::max(1, editFrameWidth - (2 * framePadding)),
                                                 std::max(1, editHeight - (2 * framePadding)),
                                                 panel,
                                                 nullptr,
                                                 GetModuleHandleW(nullptr),
                                                 nullptr);
                if (controls.hEdit)
                {
                    SetWindowLongPtrW(controls.hEditFrame, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(controls.hEdit));
                    SendMessageW(controls.hEdit, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(textMargin, textMargin));

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
                    SetWindowSubclass(controls.hEdit, PluginConfigInputControlSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(controls.hEditFrame));
#pragma warning(pop)
                }
            }
            else
            {
                controls.hEdit = CreateWindowExW(
                    WS_EX_CLIENTEDGE, L"Edit", L"", editStyle, controlX, y, editFrameWidth, editHeight, panel, nullptr, GetModuleHandleW(nullptr), nullptr);

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
                if (controls.hEdit)
                {
                    SetWindowSubclass(controls.hEdit, PluginConfigInputControlSubclassProc, 1u, 0);
                }
#pragma warning(pop)
            }
            if (controls.hEdit && font)
            {
                SendMessageW(controls.hEdit, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
            }

            if (controls.hEdit)
            {
                if (controls.field.type == PluginConfigFieldType::Text)
                {
                    SetWindowTextW(controls.hEdit, controls.field.defaultText.c_str());
                }
                else
                {
                    const std::wstring text = std::to_wstring(controls.field.defaultInt);
                    SetWindowTextW(controls.hEdit, text.c_str());
                }
            }

            y += editHeight + spacingY;
        }
        else if (controls.field.type == PluginConfigFieldType::Bool)
        {
            if (! state->theme.highContrast)
            {
                controls.toggleOnChoiceIndex  = 0;
                controls.toggleOffChoiceIndex = 1;

                const int paddingX   = ScaleDip(dpi, 6);
                const int gapX       = ScaleDip(dpi, 6);
                const int trackWidth = ScaleDip(dpi, 28);

                const std::wstring_view leftLabel  = L"True";
                const std::wstring_view rightLabel = L"False";

                const HFONT toggleFont  = state->boldFont ? state->boldFont.get() : font;
                const int leftWidth     = MeasureTextWidth(panel, toggleFont, leftLabel);
                const int rightWidth    = MeasureTextWidth(panel, toggleFont, rightLabel);
                const int slackWidth    = ScaleDip(dpi, 6);
                const int measuredWidth = std::max(minControlWidth, (2 * paddingX) + leftWidth + gapX + trackWidth + gapX + rightWidth + slackWidth);
                const int toggleWidth   = std::min(controlWidth, measuredWidth);

                controls.hToggle = CreateWindowExW(0,
                                                   L"Button",
                                                   L"",
                                                   WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                                   controlX,
                                                   y,
                                                   toggleWidth,
                                                   editHeight,
                                                   panel,
                                                   nullptr,
                                                   GetModuleHandleW(nullptr),
                                                   nullptr);
                if (controls.hToggle && font)
                {
                    SendMessageW(controls.hToggle, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
                }

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
                if (controls.hToggle)
                {
                    SetWindowSubclass(controls.hToggle, PluginConfigInputControlSubclassProc, 1u, 0);
                }
#pragma warning(pop)

                SetWindowLongPtrW(controls.hToggle, GWLP_USERDATA, controls.field.defaultBool ? 1 : 0);

                y += editHeight + spacingY;
            }
            else
            {
                const int buttonHeight = std::max(1, optionHeight);
                int optionY            = y;

                for (size_t i = 0; i < 2; ++i)
                {
                    DWORD style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_MULTILINE | BS_AUTORADIOBUTTON;
                    if (i == 0)
                    {
                        style |= WS_GROUP;
                    }

                    const wchar_t* text = (i == 0) ? L"True" : L"False";
                    HWND hButton        = CreateWindowExW(
                        0, L"Button", text, style, controlX, optionY, controlWidth, buttonHeight, panel, nullptr, GetModuleHandleW(nullptr), nullptr);
                    if (hButton && font)
                    {
                        SendMessageW(hButton, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
                    }

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
                    if (hButton)
                    {
                        SetWindowSubclass(hButton, PluginConfigInputControlSubclassProc, 1u, 0);
                    }
#pragma warning(pop)

                    if (hButton)
                    {
                        const bool checked = (i == 0) ? controls.field.defaultBool : (! controls.field.defaultBool);
                        SendMessageW(hButton, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
                    }

                    controls.choiceButtons.push_back(hButton);
                    optionY += buttonHeight;
                }

                y = optionY + spacingY;
            }
        }
        else if (controls.field.type == PluginConfigFieldType::Option && ! state->theme.highContrast && controls.field.choices.size() == 2)
        {
            size_t leftIndex  = 0;
            size_t rightIndex = 1;

            size_t onIndex  = 0;
            size_t offIndex = 0;
            if (TryGetBoolToggleChoiceIndices(controls.field, onIndex, offIndex))
            {
                leftIndex  = onIndex;
                rightIndex = offIndex;
            }

            controls.toggleOnChoiceIndex  = leftIndex;
            controls.toggleOffChoiceIndex = rightIndex;

            const int paddingX   = ScaleDip(dpi, 6);
            const int gapX       = ScaleDip(dpi, 6);
            const int trackWidth = ScaleDip(dpi, 28);

            std::wstring_view leftLabel;
            std::wstring_view rightLabel;
            if (leftIndex < controls.field.choices.size())
            {
                const auto& choice = controls.field.choices[leftIndex];
                leftLabel          = choice.label.empty() ? std::wstring_view(choice.value) : std::wstring_view(choice.label);
            }
            if (rightIndex < controls.field.choices.size())
            {
                const auto& choice = controls.field.choices[rightIndex];
                rightLabel         = choice.label.empty() ? std::wstring_view(choice.value) : std::wstring_view(choice.label);
            }

            const HFONT toggleFont  = state->boldFont ? state->boldFont.get() : font;
            const int leftWidth     = MeasureTextWidth(panel, toggleFont, leftLabel);
            const int rightWidth    = MeasureTextWidth(panel, toggleFont, rightLabel);
            const int slackWidth    = ScaleDip(dpi, 6);
            const int measuredWidth = std::max(minControlWidth, (2 * paddingX) + leftWidth + gapX + trackWidth + gapX + rightWidth + slackWidth);
            const int toggleWidth   = std::min(controlWidth, measuredWidth);

            controls.hToggle = CreateWindowExW(0,
                                               L"Button",
                                               L"",
                                               WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                               controlX,
                                               y,
                                               toggleWidth,
                                               editHeight,
                                               panel,
                                               nullptr,
                                               GetModuleHandleW(nullptr),
                                               nullptr);
            if (controls.hToggle && font)
            {
                SendMessageW(controls.hToggle, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
            }

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
            if (controls.hToggle)
            {
                SetWindowSubclass(controls.hToggle, PluginConfigInputControlSubclassProc, 1u, 0);
            }
#pragma warning(pop)

            bool isLeftActive = true;
            if (rightIndex < controls.field.choices.size() && controls.field.defaultOption == controls.field.choices[rightIndex].value)
            {
                isLeftActive = false;
            }
            else if (leftIndex < controls.field.choices.size() && controls.field.defaultOption == controls.field.choices[leftIndex].value)
            {
                isLeftActive = true;
            }
            SetWindowLongPtrW(controls.hToggle, GWLP_USERDATA, isLeftActive ? 1 : 0);

            y += editHeight + spacingY;
        }
        else if (controls.field.type == PluginConfigFieldType::Option && controls.field.choices.size() > 2)
        {
            const bool customFrames = ! state->theme.highContrast;
            const int framePadding  = ScaleDip(dpi, 2);

            if (customFrames)
            {
                controls.hComboFrame = CreateWindowExW(
                    0, L"Static", L"", WS_CHILD | WS_VISIBLE, controlX, y, controlWidth, editHeight, panel, nullptr, GetModuleHandleW(nullptr), nullptr);
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
                if (controls.hComboFrame)
                {
                    SetWindowSubclass(controls.hComboFrame, PluginConfigInputFrameSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(state));
                }
#pragma warning(pop)
            }

            const int comboX      = controlX + (customFrames ? framePadding : 0);
            const int comboY      = y + (customFrames ? framePadding : 0);
            const int comboWidth  = std::max(1, controlWidth - (customFrames ? 2 * framePadding : 0));
            const int comboHeight = std::max(1, editHeight - (customFrames ? 2 * framePadding : 0));

            if (customFrames)
            {
                controls.hCombo = ThemedControls::CreateModernComboBox(panel, 0, &state->theme);
                if (controls.hCombo)
                {
                    SetWindowPos(controls.hCombo, nullptr, comboX, comboY, comboWidth, comboHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                }
            }
            else
            {
                controls.hCombo = CreateWindowExW(0,
                                                  L"ComboBox",
                                                  L"",
                                                  WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
                                                  comboX,
                                                  comboY,
                                                  comboWidth,
                                                  comboHeight * 8,
                                                  panel,
                                                  nullptr,
                                                  GetModuleHandleW(nullptr),
                                                  nullptr);
            }
            if (controls.hCombo && font)
            {
                SendMessageW(controls.hCombo, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
            }

            if (controls.hComboFrame && controls.hCombo)
            {
                SetWindowLongPtrW(controls.hComboFrame, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(controls.hCombo));
            }

            if (controls.hCombo)
            {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
                SetWindowSubclass(controls.hCombo, PluginConfigInputControlSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(controls.hComboFrame));
#pragma warning(pop)
            }

            if (controls.hCombo)
            {
                int selectedIndex = 0;
                for (size_t i = 0; i < controls.field.choices.size(); ++i)
                {
                    const auto& choice            = controls.field.choices[i];
                    const std::wstring_view label = choice.label.empty() ? std::wstring_view(choice.value) : std::wstring_view(choice.label);
                    SendMessageW(controls.hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.data()));
                    if (! controls.field.defaultOption.empty() && controls.field.defaultOption == choice.value)
                    {
                        if (i <= static_cast<size_t>(std::numeric_limits<int>::max()))
                        {
                            selectedIndex = static_cast<int>(i);
                        }
                    }
                }

                SendMessageW(controls.hCombo, CB_SETCURSEL, static_cast<WPARAM>(selectedIndex), 0);
                ApplyThemeToComboBox(controls.hCombo, state->theme);
            }

            y += editHeight + spacingY;
        }
        else
        {
            const bool isRadio = controls.field.type == PluginConfigFieldType::Option;
            int optionY        = y;

            for (size_t i = 0; i < controls.field.choices.size(); ++i)
            {
                const auto& choice = controls.field.choices[i];

                DWORD style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_MULTILINE;
                style |= isRadio ? BS_AUTORADIOBUTTON : BS_AUTOCHECKBOX;
                if (i == 0)
                {
                    style |= WS_GROUP;
                }

                const int buttonHeight = std::max(1, optionHeight);

                HWND hButton = CreateWindowExW(0,
                                               L"Button",
                                               choice.label.c_str(),
                                               style,
                                               controlX,
                                               optionY,
                                               controlWidth,
                                               buttonHeight,
                                               panel,
                                               nullptr,
                                               GetModuleHandleW(nullptr),
                                               nullptr);
                if (hButton && font)
                {
                    SendMessageW(hButton, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
                }

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
                if (hButton)
                {
                    SetWindowSubclass(hButton, PluginConfigInputControlSubclassProc, 1u, 0);
                }
#pragma warning(pop)

                if (hButton)
                {
                    if (isRadio)
                    {
                        const bool checked = ! controls.field.defaultOption.empty() && controls.field.defaultOption == choice.value;
                        SendMessageW(hButton, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
                    }
                    else
                    {
                        const bool checked = std::find(controls.field.defaultSelection.begin(), controls.field.defaultSelection.end(), choice.value) !=
                                             controls.field.defaultSelection.end();
                        SendMessageW(hButton, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
                    }
                }

                controls.choiceButtons.push_back(hButton);
                optionY += buttonHeight;
            }

            y = optionY + spacingY;
        }

        const HFONT infoFont = state->commentFont ? state->commentFont.get() : font;
        const int infoX      = left;
        const int infoWidth  = availableWidth;

        if (! controls.field.description.empty())
        {
            const int commentHeight = MeasureInfoHeight(panel, infoFont, infoWidth, controls.field.description);

            controls.hComment = CreateWindowExW(0,
                                                L"Static",
                                                controls.field.description.c_str(),
                                                WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL,
                                                infoX,
                                                y,
                                                infoWidth,
                                                commentHeight,
                                                panel,
                                                nullptr,
                                                GetModuleHandleW(nullptr),
                                                nullptr);
            if (controls.hComment && infoFont)
            {
                SendMessageW(controls.hComment, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
            }

            y += commentHeight + ScaleDip(dpi, 4);
        }

        const std::wstring defaultsText = BuildFieldDefaultsTextForDisplay(controls.field);
        if (! defaultsText.empty())
        {
            const int defaultsHeight = MeasureInfoHeight(panel, infoFont, infoWidth, defaultsText);
            controls.hDefaults       = CreateWindowExW(0,
                                                 L"Static",
                                                 defaultsText.c_str(),
                                                 WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL,
                                                 infoX,
                                                 y,
                                                 infoWidth,
                                                 defaultsHeight,
                                                 panel,
                                                 nullptr,
                                                 GetModuleHandleW(nullptr),
                                                 nullptr);
            if (controls.hDefaults && infoFont)
            {
                SendMessageW(controls.hDefaults, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
            }

            y += defaultsHeight + spacingY;
        }

        state->controls.push_back(std::move(controls));
    }

    if (configDoc)
    {
        yyjson_doc_free(configDoc);
    }

    state->contentHeight = y + margin;
    UpdatePanelScrollInfo(panel, *state);
    return TRUE;
}

INT_PTR OnPluginConfigDialogCtlColorDialog(PluginConfigDialogState* state)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnPluginConfigDialogCtlColorStatic(PluginConfigDialogState* state, HDC hdc, HWND control)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    COLORREF textColor = state->theme.menu.text;
    if (control)
    {
        const LONG_PTR style = GetWindowLongPtrW(control, GWL_STYLE);
        if ((style & WS_DISABLED) != 0)
        {
            textColor = state->theme.menu.disabledText;
        }
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);

    // Combo box drop-down list controls often paint their selection field via a child static window; match the input background.
    HWND parent = control ? GetParent(control) : nullptr;
    if (parent)
    {
        std::array<wchar_t, 32> className{};
        const int len = GetClassNameW(parent, className.data(), static_cast<int>(className.size()));
        if (len > 0 && _wcsicmp(className.data(), L"ComboBox") == 0)
        {
            const COLORREF background = state->inputBrush ? state->inputBackgroundColor : state->theme.windowBackground;
            SetBkColor(hdc, background);
            return reinterpret_cast<INT_PTR>(state->inputBrush ? state->inputBrush.get() : state->backgroundBrush.get());
        }
    }

    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnPluginConfigDialogCtlColorButton(PluginConfigDialogState* state, HDC hdc, HWND control)
{
    if (! state || ! state->backgroundBrush || ! control)
    {
        return FALSE;
    }

    const LONG_PTR style = GetWindowLongPtrW(control, GWL_STYLE);
    const LONG_PTR type  = style & BS_TYPEMASK;

    const bool themed = type == BS_CHECKBOX || type == BS_AUTOCHECKBOX || type == BS_RADIOBUTTON || type == BS_AUTORADIOBUTTON || type == BS_3STATE ||
                        type == BS_AUTO3STATE || type == BS_GROUPBOX;

    if (! themed)
    {
        return FALSE;
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, (style & WS_DISABLED) != 0 ? state->theme.menu.disabledText : state->theme.menu.text);
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnPluginConfigDialogCtlColorEdit(PluginConfigDialogState* state, HDC hdc)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    const COLORREF background = state->inputBrush ? state->inputBackgroundColor : state->theme.windowBackground;
    SetBkColor(hdc, background);
    SetTextColor(hdc, state->theme.menu.text);
    return reinterpret_cast<INT_PTR>(state->inputBrush ? state->inputBrush.get() : state->backgroundBrush.get());
}

INT_PTR OnPluginConfigDialogCtlColorListBox(PluginConfigDialogState* state, HDC hdc)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    const COLORREF background = state->inputBrush ? state->inputBackgroundColor : state->theme.windowBackground;
    SetBkColor(hdc, background);
    SetTextColor(hdc, state->theme.menu.text);
    return reinterpret_cast<INT_PTR>(state->inputBrush ? state->inputBrush.get() : state->backgroundBrush.get());
}

INT_PTR OnPluginConfigDialogCommand(HWND dlg, PluginConfigDialogState* state, UINT commandId, UINT /*codeNotify*/, HWND /*hwndCtl*/)
{
    if (! state)
    {
        return FALSE;
    }

    if (commandId == IDOK)
    {
        const std::string configJson = BuildConfigurationJson(state->controls);
        if (configJson.empty())
        {
            EndDialog(dlg, IDCANCEL);
            return TRUE;
        }

        if (state->commitMode == PluginConfigCommitMode::UpdateSettingsOnly)
        {
            if (! state->settings || state->pluginId.empty())
            {
                EndDialog(dlg, IDCANCEL);
                return TRUE;
            }

            Common::Settings::JsonValue parsedValue;
            const HRESULT parseHr = Common::Settings::ParseJsonValue(configJson, parsedValue);
            if (FAILED(parseHr))
            {
                const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
                const std::wstring message = LoadStringResource(nullptr, IDS_MSG_PLUGIN_CONFIG_APPLY_FAILED);
                ShowDialogAlert(dlg, HOST_ALERT_ERROR, title, message);
                return TRUE;
            }

            bool clearValue = std::holds_alternative<std::monostate>(parsedValue.value);
            if (! clearValue)
            {
                const auto* obj = std::get_if<Common::Settings::JsonValue::ObjectPtr>(&parsedValue.value);
                clearValue      = obj && *obj && (*obj)->members.empty();
            }

            if (clearValue)
            {
                state->settings->plugins.configurationByPluginId.erase(state->pluginId);
            }
            else
            {
                state->settings->plugins.configurationByPluginId[state->pluginId] = std::move(parsedValue);
            }

            EndDialog(dlg, IDOK);
            return TRUE;
        }

        HRESULT hr = E_FAIL;
        if (state->pluginType == PluginType::FileSystem)
        {
            auto& manager = FileSystemPluginManager::GetInstance();
            hr            = manager.SetConfiguration(state->pluginId, configJson, *state->settings);
        }
        else
        {
            auto& manager = ViewerPluginManager::GetInstance();
            hr            = manager.SetConfiguration(state->pluginId, configJson, *state->settings);
        }
        if (FAILED(hr))
        {
            const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
            const std::wstring message = LoadStringResource(nullptr, IDS_MSG_PLUGIN_CONFIG_APPLY_FAILED);
            ShowDialogAlert(dlg, HOST_ALERT_ERROR, title, message);
            return TRUE;
        }

        if (state->settings)
        {
            PersistSettings(dlg, *state->settings, state->appId);
        }

        EndDialog(dlg, IDOK);
        return TRUE;
    }

    if (commandId == IDCANCEL)
    {
        EndDialog(dlg, IDCANCEL);
        return TRUE;
    }

    return FALSE;
}

INT_PTR CALLBACK PluginConfigDialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* state = reinterpret_cast<PluginConfigDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

    switch (msg)
    {
        case WM_INITDIALOG: return OnPluginConfigDialogInit(dlg, reinterpret_cast<PluginConfigDialogState*>(lp));
        case WM_CTLCOLORDLG: return OnPluginConfigDialogCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnPluginConfigDialogCtlColorStatic(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLORBTN: return OnPluginConfigDialogCtlColorButton(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLOREDIT: return OnPluginConfigDialogCtlColorEdit(state, reinterpret_cast<HDC>(wp));
        case WM_CTLCOLORLISTBOX: return OnPluginConfigDialogCtlColorListBox(state, reinterpret_cast<HDC>(wp));
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(dlg, state->theme, wp != FALSE);
            }
            return FALSE;
        case WM_DRAWITEM:
        {
            if (! state || state->theme.highContrast)
            {
                break;
            }

            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lp);
            if (! dis || dis->CtlType != ODT_BUTTON)
            {
                break;
            }

            DrawThemedPushButton(*dis, state->theme);
            return TRUE;
        }
        case WM_SIZING:
        {
            if (! state || state->fixedWindowWidthPx <= 0)
            {
                break;
            }

            auto* rc = reinterpret_cast<RECT*>(lp);
            if (! rc)
            {
                break;
            }

            switch (wp)
            {
                case WMSZ_LEFT:
                case WMSZ_TOPLEFT:
                case WMSZ_BOTTOMLEFT: rc->left = rc->right - state->fixedWindowWidthPx; break;
                case WMSZ_RIGHT:
                case WMSZ_TOPRIGHT:
                case WMSZ_BOTTOMRIGHT: rc->right = rc->left + state->fixedWindowWidthPx; break;
                default: rc->right = rc->left + state->fixedWindowWidthPx; break;
            }

            return TRUE;
        }
        case WM_SIZE:
        {
            if (state)
            {
                LayoutPluginConfigDialog(dlg, *state);
                if (state->panel)
                {
                    UpdatePanelScrollInfo(state->panel, *state);
                }
            }
            return TRUE;
        }
        case WM_COMMAND: return OnPluginConfigDialogCommand(dlg, state, LOWORD(wp), HIWORD(wp), reinterpret_cast<HWND>(lp));
    }

    return FALSE;
}

HRESULT
ShowPluginConfigurationDialogInternal(HWND owner,
                                      std::wstring_view appId,
                                      PluginType pluginType,
                                      std::wstring_view pluginId,
                                      std::wstring_view pluginName,
                                      Common::Settings::Settings& settings,
                                      const AppTheme& theme)
{
    if (pluginId.empty())
    {
        return E_INVALIDARG;
    }

    std::string schema;
    HRESULT schemaHr = E_FAIL;

    if (pluginType == PluginType::FileSystem)
    {
        auto& manager = FileSystemPluginManager::GetInstance();
        schemaHr      = manager.GetConfigurationSchema(pluginId, settings, schema);
    }
    else
    {
        auto& manager = ViewerPluginManager::GetInstance();
        schemaHr      = manager.GetConfigurationSchema(pluginId, settings, schema);
    }
    if (FAILED(schemaHr))
    {
        return schemaHr;
    }

    std::string current;

    // First, try to get configuration from settings.plugins.configurationByPluginId
    const std::wstring pluginIdText(pluginId);
    const auto it = settings.plugins.configurationByPluginId.find(pluginIdText);
    if (it != settings.plugins.configurationByPluginId.end() && ! std::holds_alternative<std::monostate>(it->second.value))
    {
        const HRESULT serializeHr = Common::Settings::SerializeJsonValue(it->second, current);
        if (FAILED(serializeHr))
        {
            return serializeHr;
        }
    }

    // If no configuration in settings, fall back to plugin's current configuration
    if (current.empty())
    {
        HRESULT configHr = E_FAIL;

        if (pluginType == PluginType::FileSystem)
        {
            auto& manager = FileSystemPluginManager::GetInstance();
            configHr      = manager.GetConfiguration(pluginId, settings, current);
        }
        else
        {
            auto& manager = ViewerPluginManager::GetInstance();
            configHr      = manager.GetConfiguration(pluginId, settings, current);
        }
        if (FAILED(configHr))
        {
            return configHr;
        }
    }

    PluginConfigDialogState state;
    state.settings              = &settings;
    state.appId                 = std::wstring(appId);
    state.theme                 = theme;
    state.pluginType            = pluginType;
    state.pluginId              = std::wstring(pluginId);
    state.pluginName            = pluginName.empty() ? std::wstring(pluginId) : std::wstring(pluginName);
    state.schemaJsonUtf8        = std::move(schema);
    state.configurationJsonUtf8 = std::move(current);

#pragma warning(push)
#pragma warning(disable : 5039) // pointer/reference to potentially throwing function passed to extern "C"
    const INT_PTR result =
        DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_PLUGIN_CONFIG), owner, PluginConfigDialogProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)

    return result == IDOK ? S_OK : S_FALSE;
}

#if 0
HRESULT AddPluginFromDialog(HWND dlg, DialogState& state)
{
    if (! state.settings)
    {
        return E_FAIL;
    }

    std::array<wchar_t, 2048> fileBuffer{};

    OPENFILENAMEW ofn{};
    ofn.lStructSize           = sizeof(ofn);
    ofn.hwndOwner             = dlg;
    ofn.lpstrFile             = fileBuffer.data();
    ofn.nMaxFile              = static_cast<DWORD>(fileBuffer.size());
    const std::wstring filter = LoadStringResource(nullptr, IDS_FILEDLG_FILTER_DLL);
    ofn.lpstrFilter           = filter.c_str();
    ofn.nFilterIndex          = 1;
    ofn.Flags                 = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_EXPLORER;

    if (! GetOpenFileNameW(&ofn))
    {
        return S_FALSE;
    }

    const std::filesystem::path selectedPath(fileBuffer.data());
    if (selectedPath.empty())
    {
        return E_INVALIDARG;
    }

    HRESULT hr              = E_FAIL;

    auto& fileSystemManager = FileSystemPluginManager::GetInstance();
    hr                      = fileSystemManager.AddCustomPluginPath(selectedPath, *state.settings);

    if (hr == E_NOINTERFACE)
    {
        auto& viewerManager = ViewerPluginManager::GetInstance();
        hr                  = viewerManager.AddCustomPluginPath(selectedPath, *state.settings);
    }
    if (FAILED(hr))
    {
        UINT messageId = IDS_MSG_PLUGIN_ADD_FAILED;
        if (hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
        {
            messageId = IDS_MSG_PLUGIN_ADD_DUPLICATE_ID;
        }
        else if (hr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND) || hr == E_NOINTERFACE)
        {
            messageId = IDS_MSG_PLUGIN_ADD_INVALID;
        }

        const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        const std::wstring message = LoadStringResource(nullptr, messageId);
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, title, message);
    }
    return hr;
}

INT_PTR OnNotify(HWND dlg, DialogState* state, WPARAM /*wp*/, LPARAM lp)
{
    const auto* hdr = reinterpret_cast<const NMHDR*>(lp);
    if (! hdr || ! state)
    {
        return FALSE;
    }

    if (hdr->idFrom == IDC_PLUGINS_LIST && hdr->code == LVN_ITEMCHANGED)
    {
        UpdateButtons(dlg, *state);
        return FALSE;
    }

    if (hdr->idFrom == IDC_PLUGINS_LIST && hdr->code == NM_DBLCLK)
    {
        const auto* activate = reinterpret_cast<const NMITEMACTIVATE*>(lp);
        if (! activate || activate->iItem < 0)
        {
            return FALSE;
        }

        const std::optional<SelectedPlugin> selected = FindSelectedPlugin(*state);
        if (! selected.has_value() || ! state->settings)
        {
            return FALSE;
        }

        std::wstring_view pluginId;
        std::wstring_view pluginName;
        PluginType pluginType = PluginType::FileSystem;
        bool loadable         = false;

        if (selected.value().type == PluginType::FileSystem && selected.value().fileSystem)
        {
            const auto& plugin = *selected.value().fileSystem;
            pluginId           = plugin.id;
            pluginName         = plugin.name.empty() ? plugin.id : plugin.name;
            pluginType         = PluginType::FileSystem;
            loadable           = plugin.loadable && ! plugin.id.empty();
        }
        else if (selected.value().type == PluginType::Viewer && selected.value().viewer)
        {
            const auto& plugin = *selected.value().viewer;
            pluginId           = plugin.id;
            pluginName         = plugin.name.empty() ? plugin.id : plugin.name;
            pluginType         = PluginType::Viewer;
            loadable           = plugin.loadable && ! plugin.id.empty();
        }

        if (! loadable || pluginId.empty())
        {
            return FALSE;
        }

        const HRESULT hr = ShowPluginConfigurationDialogInternal(dlg, state->appId, pluginType, pluginId, pluginName, *state->settings, state->theme);
        if (hr == S_OK)
        {
            PopulatePluginsList(dlg, *state);
            UpdateButtons(dlg, *state);
        }

        return TRUE;
    }

    if (hdr->idFrom == IDC_PLUGINS_LIST && hdr->code == NM_CUSTOMDRAW)
    {
        auto* cd = reinterpret_cast<NMLVCUSTOMDRAW*>(lp);
        if (! cd)
        {
            return FALSE;
        }

        if (cd->nmcd.dwDrawStage == CDDS_PREPAINT)
        {
            return CDRF_NOTIFYITEMDRAW;
        }

        if (cd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT || cd->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM))
        {
            const bool listFocused  = state->list && GetFocus() == state->list;
            const bool windowActive = GetActiveWindow() == dlg;

            if (cd->dwItemType == LVCDI_GROUP)
            {
                cd->clrTextBk = state->theme.windowBackground;
                COLORREF text = (windowActive ? state->theme.menu.headerText : state->theme.menu.headerTextDisabled);
                if (text == state->theme.windowBackground)
                {
                    text = ChooseContrastingTextColor(state->theme.windowBackground);
                }
                cd->clrText = text;
                return CDRF_NEWFONT;
            }

            if (cd->dwItemType != LVCDI_ITEM)
            {
                return CDRF_NEWFONT;
            }

            const int itemIndex = static_cast<int>(cd->nmcd.dwItemSpec);

            LVITEMW item{};
            item.mask  = LVIF_PARAM;
            item.iItem = itemIndex;
            if (ListView_GetItem(state->list, &item))
            {
                const size_t listIndex = static_cast<size_t>(item.lParam);
                if (listIndex < state->listItems.size())
                {
                    const PluginListItem& listItem = state->listItems[listIndex];

                    bool disabled                  = false;
                    if (listItem.type == PluginType::FileSystem)
                    {
                        auto& manager       = FileSystemPluginManager::GetInstance();
                        const auto& plugins = manager.GetPlugins();
                        if (listItem.index < plugins.size())
                        {
                            const auto& plugin = plugins[listItem.index];
                            disabled           = plugin.disabled || ! plugin.loadable;
                        }
                    }
                    else
                    {
                        auto& manager       = ViewerPluginManager::GetInstance();
                        const auto& plugins = manager.GetPlugins();
                        if (listItem.index < plugins.size())
                        {
                            const auto& plugin = plugins[listItem.index];
                            disabled           = plugin.disabled || ! plugin.loadable;
                        }
                    }

                    const bool selectedItem = (cd->nmcd.uItemState & CDIS_SELECTED) != 0;
                    COLORREF bg             = state->theme.windowBackground;
                    COLORREF text           = disabled ? state->theme.menu.disabledText : state->theme.menu.text;
                    if (selectedItem)
                    {
                        const COLORREF selBg = state->theme.menu.selectionBg;
                        if (windowActive && listFocused)
                        {
                            bg   = selBg;
                            text = state->theme.menu.selectionText;
                        }
                        else if (! state->theme.highContrast)
                        {
                            const int denom = state->theme.menu.darkBase ? 2 : 3;
                            bg              = BlendColor(state->theme.windowBackground, selBg, 1, denom);
                            text            = ChooseContrastingTextColor(bg);
                        }
                        else
                        {
                            bg   = selBg;
                            text = state->theme.menu.selectionText;
                        }
                    }

                    cd->clrTextBk = bg;
                    cd->clrText   = text;
                }
            }

            return CDRF_NEWFONT;
        }
    }

    return FALSE;
}

INT_PTR OnPluginsDialogInit(HWND dlg, DialogState* state)
{
    if (! state)
    {
        return FALSE;
    }

    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));

    ApplyTitleBarTheme(dlg, state->theme, GetActiveWindow() == dlg);

    state->backgroundBrush.reset(CreateSolidBrush(state->theme.windowBackground));
    state->selectionBrush.reset(CreateSolidBrush(state->theme.menu.selectionBg));

    state->list = GetDlgItem(dlg, IDC_PLUGINS_LIST);

    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC  = ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icc);

    if (state->list)
    {
        ListView_SetExtendedListViewStyle(state->list, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP);

        const UINT dpi = GetDpiForWindow(dlg);
        EnsureListColumns(state->list, dpi);
        EnsureListGroups(state->list);

        ListView_SetBkColor(state->list, state->theme.windowBackground);
        ListView_SetTextBkColor(state->list, state->theme.windowBackground);
        ListView_SetTextColor(state->list, state->theme.menu.text);

        if (state->theme.highContrast)
        {
            SetWindowTheme(state->list, L"", nullptr);
        }
        else
        {
            // Prefer deriving the list view theme from the effective background so custom theme overrides stay readable.
            const bool darkBackground = ChooseContrastingTextColor(state->theme.windowBackground) == RGB(255, 255, 255);
            const wchar_t* listTheme  = darkBackground ? L"DarkMode_Explorer" : L"Explorer";
            SetWindowTheme(state->list, listTheme, nullptr);

            const HWND header = ListView_GetHeader(state->list);
            if (header)
            {
                SetWindowTheme(header, listTheme, nullptr);
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
                SetWindowSubclass(header, PluginsHeaderSubclassProc, kPluginsHeaderSubclassId, reinterpret_cast<DWORD_PTR>(state));
#pragma warning(pop)
                InvalidateRect(header, nullptr, TRUE);
            }
            const HWND tooltips = ListView_GetToolTips(state->list);
            if (tooltips)
            {
                SetWindowTheme(tooltips, listTheme, nullptr);
            }
        }
    }

    if (! state->theme.highContrast)
    {
        EnableOwnerDrawButton(dlg, IDC_PLUGINS_ADD);
        EnableOwnerDrawButton(dlg, IDC_PLUGINS_REMOVE);
        EnableOwnerDrawButton(dlg, IDC_PLUGINS_CONFIGURE);
        EnableOwnerDrawButton(dlg, IDC_PLUGINS_TEST);
        EnableOwnerDrawButton(dlg, IDC_PLUGINS_TEST_ALL);
        EnableOwnerDrawButton(dlg, IDC_PLUGINS_ABOUT);
        EnableOwnerDrawButton(dlg, IDCANCEL);
    }

    PopulatePluginsList(dlg, *state);
    UpdateButtons(dlg, *state);
    return TRUE;
}

INT_PTR OnPluginsDialogCtlColorDialog(DialogState* state)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnPluginsDialogCtlColorStatic(DialogState* state, HDC hdc)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, state->theme.menu.text);
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnPluginsDialogActivate(DialogState* state) noexcept
{
    if (state && state->list)
    {
        InvalidateRect(state->list, nullptr, FALSE);
        const HWND header = ListView_GetHeader(state->list);
        if (header)
        {
            InvalidateRect(header, nullptr, FALSE);
        }
    }
    return FALSE;
}

INT_PTR OnPluginsDialogCommand(HWND dlg, DialogState* state, UINT commandId, UINT /*codeNotify*/, HWND /*hwndCtl*/)
{
    if (! state)
    {
        return FALSE;
    }

    switch (commandId)
    {
        case IDC_PLUGINS_ADD:
        {
            const HRESULT hr = AddPluginFromDialog(dlg, *state);
            if (hr == S_OK)
            {
                PopulatePluginsList(dlg, *state);
                if (state->settings)
                {
                    PersistSettings(dlg, *state->settings, state->appId);
                }
            }
        }
        break;
        case IDC_PLUGINS_REMOVE:
        {
            if (! state->settings)
            {
                break;
            }

            const std::optional<SelectedPlugin> selected = FindSelectedPlugin(*state);
            if (! selected.has_value())
            {
                break;
            }

            HRESULT hr = S_OK;

            if (selected.value().type == PluginType::FileSystem && selected.value().fileSystem)
            {
                auto& manager      = FileSystemPluginManager::GetInstance();
                const auto& plugin = *selected.value().fileSystem;

                if (plugin.origin == FileSystemPluginManager::PluginOrigin::Custom)
                {
                    if (plugin.id.empty())
                    {
                        RemovePathFromVector(state->settings->plugins.customPluginPaths, plugin.path);
                        hr = manager.Refresh(*state->settings);
                    }
                    else
                    {
                        hr = manager.RemoveCustomPlugin(plugin.id, *state->settings);
                    }
                }
                else if (! plugin.id.empty())
                {
                    hr = plugin.disabled ? manager.EnablePlugin(plugin.id, *state->settings) : manager.DisablePlugin(plugin.id, *state->settings);
                }
            }
            else if (selected.value().type == PluginType::Viewer && selected.value().viewer)
            {
                auto& manager      = ViewerPluginManager::GetInstance();
                const auto& plugin = *selected.value().viewer;

                if (plugin.origin == ViewerPluginManager::PluginOrigin::Custom)
                {
                    RemovePathFromVector(state->settings->plugins.customPluginPaths, plugin.path);
                    hr = manager.Refresh(*state->settings);
                }
                else if (! plugin.id.empty())
                {
                    hr = plugin.disabled ? manager.EnablePlugin(plugin.id, *state->settings) : manager.DisablePlugin(plugin.id, *state->settings);
                }
            }

            if (SUCCEEDED(hr))
            {
                PopulatePluginsList(dlg, *state);
                UpdateButtons(dlg, *state);
                if (state->settings)
                {
                    PersistSettings(dlg, *state->settings, state->appId);
                }
            }
        }
        break;
        case IDC_PLUGINS_CONFIGURE:
        {
            if (! state->settings)
            {
                break;
            }

            const std::optional<SelectedPlugin> selected = FindSelectedPlugin(*state);
            if (! selected.has_value())
            {
                break;
            }

            std::wstring_view pluginId;
            std::wstring_view pluginName;
            PluginType pluginType = PluginType::FileSystem;
            bool loadable         = false;

            if (selected.value().type == PluginType::FileSystem && selected.value().fileSystem)
            {
                const auto& plugin = *selected.value().fileSystem;
                pluginId           = plugin.id;
                pluginName         = plugin.name.empty() ? plugin.id : plugin.name;
                pluginType         = PluginType::FileSystem;
                loadable           = plugin.loadable && ! plugin.id.empty();
            }
            else if (selected.value().type == PluginType::Viewer && selected.value().viewer)
            {
                const auto& plugin = *selected.value().viewer;
                pluginId           = plugin.id;
                pluginName         = plugin.name.empty() ? plugin.id : plugin.name;
                pluginType         = PluginType::Viewer;
                loadable           = plugin.loadable && ! plugin.id.empty();
            }

            if (! loadable || pluginId.empty())
            {
                break;
            }

            const HRESULT hr = ShowPluginConfigurationDialogInternal(dlg, state->appId, pluginType, pluginId, pluginName, *state->settings, state->theme);
            if (hr == S_OK)
            {
                PopulatePluginsList(dlg, *state);
            }
        }
        break;
        case IDC_PLUGINS_TEST_ALL:
        {
            size_t okCount   = 0;
            size_t failCount = 0;

            {
                auto& manager = FileSystemPluginManager::GetInstance();
                for (const auto& entry : manager.GetPlugins())
                {
                    if (entry.id.empty())
                    {
                        continue;
                    }

                    const HRESULT hr = manager.TestPlugin(entry.id);
                    if (SUCCEEDED(hr))
                    {
                        ++okCount;
                    }
                    else
                    {
                        ++failCount;
                    }
                }
            }

            {
                auto& manager = ViewerPluginManager::GetInstance();
                for (const auto& entry : manager.GetPlugins())
                {
                    if (entry.id.empty())
                    {
                        continue;
                    }

                    const HRESULT hr = manager.TestPlugin(entry.id);
                    if (SUCCEEDED(hr))
                    {
                        ++okCount;
                    }
                    else
                    {
                        ++failCount;
                    }
                }
            }

            const std::wstring message = FormatStringResource(nullptr, IDS_FMT_PLUGIN_TEST_ALL_RESULT, okCount, failCount);
            const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER);
            ShowDialogAlert(dlg, HOST_ALERT_INFO, title, message);
        }
        break;
        case IDC_PLUGINS_TEST:
        {
            const std::optional<SelectedPlugin> selected = FindSelectedPlugin(*state);
            if (! selected.has_value())
            {
                break;
            }

            HRESULT hr = E_FAIL;

            if (selected.value().type == PluginType::FileSystem && selected.value().fileSystem && ! selected.value().fileSystem->id.empty())
            {
                auto& manager = FileSystemPluginManager::GetInstance();
                hr            = manager.TestPlugin(selected.value().fileSystem->id);
            }
            else if (selected.value().type == PluginType::Viewer && selected.value().viewer && ! selected.value().viewer->id.empty())
            {
                auto& manager = ViewerPluginManager::GetInstance();
                hr            = manager.TestPlugin(selected.value().viewer->id);
            }
            else
            {
                break;
            }

            const UINT textId                = SUCCEEDED(hr) ? IDS_MSG_PLUGIN_TEST_OK : IDS_MSG_PLUGIN_TEST_FAILED;
            const HostAlertSeverity severity = SUCCEEDED(hr) ? HOST_ALERT_INFO : HOST_ALERT_ERROR;
            const std::wstring title         = LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER);
            const std::wstring message       = LoadStringResource(nullptr, textId);
            ShowDialogAlert(dlg, severity, title, message);
        }
        break;
        case IDC_PLUGINS_ABOUT:
        {
            const std::optional<SelectedPlugin> selected = FindSelectedPlugin(*state);
            if (! selected.has_value())
            {
                break;
            }

            std::wstring_view description;
            std::wstring_view id;
            std::wstring_view version;
            std::filesystem::path path;

            if (selected.value().type == PluginType::FileSystem && selected.value().fileSystem)
            {
                const auto& plugin = *selected.value().fileSystem;
                description        = plugin.description;
                id                 = plugin.id;
                version            = plugin.version;
                path               = plugin.path;
            }
            else if (selected.value().type == PluginType::Viewer && selected.value().viewer)
            {
                const auto& plugin = *selected.value().viewer;
                description        = plugin.description;
                id                 = plugin.id;
                version            = plugin.version;
                path               = plugin.path;
            }
            else
            {
                break;
            }

            const std::wstring message = FormatStringResource(nullptr, IDS_FMT_PLUGIN_ABOUT, description, id, version, path.wstring());
            const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER);
            ShowDialogAlert(dlg, HOST_ALERT_INFO, title, message);
        }
        break;
        case IDCANCEL: EndDialog(dlg, IDCANCEL); return TRUE;
    }

    return FALSE;
}

INT_PTR CALLBACK PluginsDialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* state = reinterpret_cast<DialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

    switch (msg)
    {
        case WM_INITDIALOG: return OnPluginsDialogInit(dlg, reinterpret_cast<DialogState*>(lp));
        case WM_CTLCOLORDLG: return OnPluginsDialogCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnPluginsDialogCtlColorStatic(state, reinterpret_cast<HDC>(wp));
        case WM_ACTIVATE: return OnPluginsDialogActivate(state);
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(dlg, state->theme, wp != FALSE);
            }
            return FALSE;
        case WM_DRAWITEM:
        {
            if (! state || state->theme.highContrast)
            {
                break;
            }

            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lp);
            if (! dis || dis->CtlType != ODT_BUTTON)
            {
                break;
            }

            DrawThemedPushButton(*dis, state->theme);
            return TRUE;
        }
        case WM_NOTIFY: return OnNotify(dlg, state, wp, lp);
        case WM_COMMAND: return OnPluginsDialogCommand(dlg, state, LOWORD(wp), HIWORD(wp), reinterpret_cast<HWND>(lp));
    }

    return FALSE;
}
#endif
} // namespace

HRESULT ShowPluginConfigurationDialog(HWND owner,
                                      std::wstring_view appId,
                                      PluginType pluginType,
                                      std::wstring_view pluginId,
                                      std::wstring_view pluginName,
                                      Common::Settings::Settings& settings,
                                      const AppTheme& theme)
{
    return ShowPluginConfigurationDialogInternal(owner, appId, pluginType, pluginId, pluginName, settings, theme);
}

HRESULT EditPluginConfigurationDialog(HWND owner,
                                      PluginType pluginType,
                                      std::wstring_view pluginId,
                                      std::wstring_view pluginName,
                                      Common::Settings::Settings& baselineSettings,
                                      Common::Settings::Settings& inOutWorkingSettings,
                                      const AppTheme& theme)
{
    if (pluginId.empty())
    {
        return E_INVALIDARG;
    }

    std::string schema;
    HRESULT schemaHr = E_FAIL;

    if (pluginType == PluginType::FileSystem)
    {
        auto& manager = FileSystemPluginManager::GetInstance();
        schemaHr      = manager.GetConfigurationSchema(pluginId, baselineSettings, schema);
    }
    else
    {
        auto& manager = ViewerPluginManager::GetInstance();
        schemaHr      = manager.GetConfigurationSchema(pluginId, baselineSettings, schema);
    }
    if (FAILED(schemaHr))
    {
        return schemaHr;
    }

    std::string current;
    const std::wstring pluginIdText(pluginId);
    const auto it = inOutWorkingSettings.plugins.configurationByPluginId.find(pluginIdText);
    if (it != inOutWorkingSettings.plugins.configurationByPluginId.end() && ! std::holds_alternative<std::monostate>(it->second.value))
    {
        const HRESULT serializeHr = Common::Settings::SerializeJsonValue(it->second, current);
        if (FAILED(serializeHr))
        {
            return serializeHr;
        }
    }

    if (current.empty())
    {
        HRESULT configHr = E_FAIL;

        if (pluginType == PluginType::FileSystem)
        {
            auto& manager = FileSystemPluginManager::GetInstance();
            configHr      = manager.GetConfiguration(pluginId, baselineSettings, current);
        }
        else
        {
            auto& manager = ViewerPluginManager::GetInstance();
            configHr      = manager.GetConfiguration(pluginId, baselineSettings, current);
        }
        if (FAILED(configHr))
        {
            return configHr;
        }
    }

    PluginConfigDialogState state;
    state.settings              = &inOutWorkingSettings;
    state.theme                 = theme;
    state.pluginType            = pluginType;
    state.pluginId              = std::wstring(pluginId);
    state.pluginName            = pluginName.empty() ? std::wstring(pluginId) : std::wstring(pluginName);
    state.schemaJsonUtf8        = std::move(schema);
    state.configurationJsonUtf8 = std::move(current);
    state.commitMode            = PluginConfigCommitMode::UpdateSettingsOnly;

#pragma warning(push)
#pragma warning(disable : 5039) // pointer/reference to potentially throwing function passed to extern "C"
    const INT_PTR result =
        DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_PLUGIN_CONFIG), owner, PluginConfigDialogProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)

    return result == IDOK ? S_OK : S_FALSE;
}
