// Preferences.Plugins.cpp

#include "Framework.h"

#include "Preferences.Plugin.Configuration.h"
#include "Preferences.Plugins.h"

#include <algorithm>
#include <array>
#include <cerrno>
#include <cwchar>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include <commctrl.h>
#include <commdlg.h>
#include <shobjidl.h>
#include <yyjson.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "FileSystemPluginManager.h"
#include "Helpers.h"
#include "HostServices.h"
#include "ManagePluginsDialog.h"
#include "ThemedControls.h"
#include "ViewerPluginManager.h"
#include "WindowMessages.h"

#include "resource.h"

namespace
{
constexpr int kPluginsColumnName            = 0;
constexpr int kPluginsColumnType            = 1;
constexpr int kPluginsColumnOrigin          = 2;
constexpr int kPluginsColumnId              = 3;
constexpr int kPluginsCustomPathsColumnPath = 0;

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

[[nodiscard]] bool IsDllPath(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    std::error_code ec;
    if (! std::filesystem::is_regular_file(path, ec))
    {
        return false;
    }

    const std::wstring ext = path.extension().wstring();
    return _wcsicmp(ext.c_str(), L".dll") == 0;
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
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

[[nodiscard]] bool TryBrowseCustomPluginPath(HWND owner, std::filesystem::path& outPath) noexcept
{
    outPath.clear();

    std::array<wchar_t, 2048> fileBuffer{};
    fileBuffer[0] = L'\0';

    OPENFILENAMEW ofn{};
    ofn.lStructSize           = sizeof(ofn);
    ofn.hwndOwner             = owner;
    ofn.lpstrFile             = fileBuffer.data();
    ofn.nMaxFile              = static_cast<DWORD>(fileBuffer.size());
    const std::wstring filter = LoadStringResource(nullptr, IDS_FILEDLG_FILTER_DLL);
    ofn.lpstrFilter           = filter.c_str();
    ofn.nFilterIndex          = 1;
    ofn.lpstrDefExt           = L"dll";
    ofn.Flags                 = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_EXPLORER | OFN_NOCHANGEDIR | OFN_HIDEREADONLY;

    if (! GetOpenFileNameW(&ofn))
    {
        return false;
    }

    outPath = std::filesystem::path(fileBuffer.data());
    return ! outPath.empty();
}

[[nodiscard]] std::wstring_view GetPluginDisplayName(const PrefsPluginListItem& item) noexcept
{
    return PrefsPlugins::GetDisplayName(item);
}

[[nodiscard]] std::wstring_view GetPluginId(const PrefsPluginListItem& item) noexcept
{
    return PrefsPlugins::GetId(item);
}

[[nodiscard]] std::wstring_view GetPluginShortIdOrId(const PrefsPluginListItem& item) noexcept
{
    return PrefsPlugins::GetShortIdOrId(item);
}

[[nodiscard]] bool IsPluginLoadable(const PrefsPluginListItem& item) noexcept
{
    return PrefsPlugins::IsLoadable(item);
}

[[nodiscard]] int GetPluginOriginOrder(const PrefsPluginListItem& item) noexcept
{
    return PrefsPlugins::GetOriginOrder(item);
}

[[nodiscard]] std::wstring GetPluginOriginText(const PrefsPluginListItem& item) noexcept
{
    const int origin = GetPluginOriginOrder(item);
    switch (origin)
    {
        case 0: return LoadStringResource(nullptr, IDS_PREFS_PLUGINS_ORIGIN_EMBEDDED);
        case 1: return LoadStringResource(nullptr, IDS_PREFS_PLUGINS_ORIGIN_OPTIONAL);
        case 2: return LoadStringResource(nullptr, IDS_PREFS_PLUGINS_ORIGIN_CUSTOM);
        default: return LoadStringResource(nullptr, IDS_PREFS_PLUGINS_ORIGIN_CUSTOM);
    }
}

[[nodiscard]] std::optional<PrefsPluginListItem> TryGetSelectedPluginItem(const PreferencesDialogState& state) noexcept
{
    if (! state.pluginsList)
    {
        return std::nullopt;
    }

    const int selected = ListView_GetNextItem(state.pluginsList.get(), -1, LVNI_SELECTED);
    if (selected < 0)
    {
        return std::nullopt;
    }

    LVITEMW item{};
    item.mask  = LVIF_PARAM;
    item.iItem = selected;
    if (! ListView_GetItem(state.pluginsList.get(), &item))
    {
        return std::nullopt;
    }

    const size_t rowIndex = static_cast<size_t>(item.lParam);
    if (rowIndex >= state.pluginsListItems.size())
    {
        return std::nullopt;
    }

    return state.pluginsListItems[rowIndex];
}

[[nodiscard]] std::optional<PrefsPluginListItem> TryGetActivePluginItem(const PreferencesDialogState& state) noexcept
{
    if (state.pluginsSelectedPlugin.has_value())
    {
        return state.pluginsSelectedPlugin;
    }
    return TryGetSelectedPluginItem(state);
}

[[nodiscard]] HTREEITEM FindPluginChildTreeItem(const PreferencesDialogState& state, const PrefsPluginListItem& plugin) noexcept
{
    if (! state.categoryTree || ! state.pluginsTreeRoot)
    {
        return nullptr;
    }

    const LPARAM desired = PrefsNavTree::EncodePluginData(plugin.type, plugin.index);

    HTREEITEM current = TreeView_GetChild(state.categoryTree, state.pluginsTreeRoot);
    while (current)
    {
        TVITEMW item{};
        item.mask  = TVIF_PARAM;
        item.hItem = current;
        if (TreeView_GetItem(state.categoryTree, &item) && item.lParam == desired)
        {
            return current;
        }

        current = TreeView_GetNextSibling(state.categoryTree, current);
    }

    return nullptr;
}

void UpdatePluginsActionButtonsEnabled(const PreferencesDialogState& state) noexcept
{
    const bool showDetails = state.pluginsSelectedPlugin.has_value();

    bool hasSelection                                 = false;
    bool loadable                                     = false;
    const std::optional<PrefsPluginListItem> selected = TryGetActivePluginItem(state);
    if (selected.has_value())
    {
        const std::wstring_view pluginId = GetPluginId(selected.value());
        hasSelection                     = ! pluginId.empty();
        loadable                         = hasSelection && IsPluginLoadable(selected.value());
    }

    if (state.pluginsConfigureButton)
    {
        const bool enableConfigure = showDetails ? (loadable && state.settings != nullptr) : hasSelection;
        EnableWindow(state.pluginsConfigureButton.get(), enableConfigure ? TRUE : FALSE);
    }
    if (state.pluginsTestButton)
    {
        EnableWindow(state.pluginsTestButton.get(), loadable ? TRUE : FALSE);
    }
    if (state.pluginsTestAllButton)
    {
        EnableWindow(state.pluginsTestAllButton.get(), TRUE);
    }
}

[[nodiscard]] bool IsPluginDisabledInWorkingSettings(const PreferencesDialogState& state, std::wstring_view pluginId) noexcept
{
    for (const auto& disabledId : state.workingSettings.plugins.disabledPluginIds)
    {
        if (std::wstring_view(disabledId) == pluginId)
        {
            return true;
        }
    }
    return false;
}

void EnsurePluginsListColumns(HWND list, UINT dpi) noexcept
{
    if (! list)
    {
        return;
    }

    HWND header               = ListView_GetHeader(list);
    const int existingColumns = header ? Header_GetItemCount(header) : 0;
    if (existingColumns >= 4)
    {
        return;
    }

    const std::wstring colName   = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_COL_NAME);
    const std::wstring colType   = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_COL_TYPE);
    const std::wstring colOrigin = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_COL_ORIGIN);
    const std::wstring colId     = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_COL_ID);

    LVCOLUMNW col{};
    col.mask     = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
    col.iSubItem = kPluginsColumnName;
    col.cx       = std::max(1, ThemedControls::ScaleDip(dpi, 220));
    col.pszText  = const_cast<LPWSTR>(colName.c_str());
    ListView_InsertColumn(list, kPluginsColumnName, &col);

    col.iSubItem = kPluginsColumnType;
    col.cx       = std::max(1, ThemedControls::ScaleDip(dpi, 90));
    col.pszText  = const_cast<LPWSTR>(colType.c_str());
    ListView_InsertColumn(list, kPluginsColumnType, &col);

    col.iSubItem = kPluginsColumnOrigin;
    col.cx       = std::max(1, ThemedControls::ScaleDip(dpi, 90));
    col.pszText  = const_cast<LPWSTR>(colOrigin.c_str());
    ListView_InsertColumn(list, kPluginsColumnOrigin, &col);

    col.iSubItem = kPluginsColumnId;
    col.cx       = std::max(1, ThemedControls::ScaleDip(dpi, 160));
    col.pszText  = const_cast<LPWSTR>(colId.c_str());
    ListView_InsertColumn(list, kPluginsColumnId, &col);
}

void UpdatePluginsListColumnWidths(HWND list, UINT dpi) noexcept
{
    if (! list)
    {
        return;
    }

    EnsurePluginsListColumns(list, dpi);

    RECT rc{};
    GetClientRect(list, &rc);
    const int width = std::max(0l, rc.right - rc.left);
    if (width <= 0)
    {
        return;
    }

    const int typeWidth   = std::min(width, std::max(1, ThemedControls::ScaleDip(dpi, 90)));
    const int originWidth = std::min(width, std::max(1, ThemedControls::ScaleDip(dpi, 90)));
    const int idWidth     = std::min(width, std::max(1, ThemedControls::ScaleDip(dpi, 170)));
    const int nameWidth   = std::max(0, width - typeWidth - originWidth - idWidth);

    ListView_SetColumnWidth(list, kPluginsColumnName, nameWidth);
    ListView_SetColumnWidth(list, kPluginsColumnType, typeWidth);
    ListView_SetColumnWidth(list, kPluginsColumnOrigin, originWidth);
    ListView_SetColumnWidth(list, kPluginsColumnId, idWidth);
}

void EnsurePluginsCustomPathsListColumns(HWND list, UINT dpi) noexcept
{
    if (! list)
    {
        return;
    }

    const HWND header         = ListView_GetHeader(list);
    const int existingColumns = header ? Header_GetItemCount(header) : 0;
    if (existingColumns > 0)
    {
        return;
    }

    LVCOLUMNW col{};
    col.mask     = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
    col.iSubItem = kPluginsCustomPathsColumnPath;
    col.cx       = std::max(1, ThemedControls::ScaleDip(dpi, 220));
    col.pszText  = const_cast<LPWSTR>(L"");
    ListView_InsertColumn(list, kPluginsCustomPathsColumnPath, &col);
}

void UpdatePluginsCustomPathsListColumnWidths(HWND list) noexcept
{
    if (! list)
    {
        return;
    }

    RECT rc{};
    GetClientRect(list, &rc);
    const int width = std::max(0l, rc.right - rc.left);
    ListView_SetColumnWidth(list, kPluginsCustomPathsColumnPath, width);
}

void SetPluginsListRowEnabled(HWND list, int row, bool enabled) noexcept
{
    if (! list || row < 0)
    {
        return;
    }

    const UINT state = static_cast<UINT>(INDEXTOSTATEIMAGEMASK(enabled ? 2 : 1));
    ListView_SetItemState(list, row, state, static_cast<UINT>(LVIS_STATEIMAGEMASK));
}
} // namespace

bool PluginsPane::EnsureCreated(HWND pageHost) noexcept
{
    return PrefsPaneHost::EnsureCreated(pageHost, _hWnd);
}

void PluginsPane::ResizeToHostClient(HWND pageHost) noexcept
{
    PrefsPaneHost::ResizeToHostClient(pageHost, _hWnd.get());
}

void PluginsPane::Show(bool visible) noexcept
{
    PrefsPaneHost::Show(_hWnd.get(), visible);
}

bool PluginsPane::HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept
{
    if (! host)
    {
        return false;
    }

    if (! state.refreshingPluginsPage)
    {
        if (PrefsPluginConfiguration::HandleCommand(host, state, notifyCode, hwndCtl))
        {
            return true;
        }
    }

    if (commandId == IDC_PREFS_PLUGINS_CONFIGURE)
    {
        if (notifyCode != BN_CLICKED)
        {
            return true;
        }

        HWND dlg = GetParent(host);
        if (! dlg)
        {
            return true;
        }

        if (! state.pluginsSelectedPlugin.has_value())
        {
            const std::optional<PrefsPluginListItem> selected = TryGetSelectedPluginItem(state);
            if (! selected.has_value())
            {
                return true;
            }

            if (HTREEITEM item = FindPluginChildTreeItem(state, selected.value()))
            {
                TreeView_SelectItem(state.categoryTree, item);
                TreeView_EnsureVisible(state.categoryTree, item);
            }
            return true;
        }

        const std::optional<PrefsPluginListItem> selected = TryGetActivePluginItem(state);
        if (! selected.has_value())
        {
            return true;
        }

        const std::wstring_view pluginId = GetPluginId(selected.value());
        if (pluginId.empty() || ! IsPluginLoadable(selected.value()) || ! state.settings)
        {
            return true;
        }

        const std::wstring_view pluginName = GetPluginDisplayName(selected.value());
        const PluginType pluginType        = (selected.value().type == PrefsPluginType::FileSystem) ? PluginType::FileSystem : PluginType::Viewer;

        const HRESULT hr = EditPluginConfigurationDialog(dlg, pluginType, pluginId, pluginName, *state.settings, state.workingSettings, state.theme);
        if (hr == S_FALSE)
        {
            return true;
        }
        if (FAILED(hr))
        {
            const std::wstring nameText(pluginName.empty() ? pluginId : pluginName);
            ShowDialogAlert(dlg,
                            HOST_ALERT_ERROR,
                            LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                            FormatStringResource(nullptr, IDS_PREFS_PLUGINS_CONFIGURE_OPEN_FAILED_FMT, nameText, static_cast<unsigned long>(hr)));
            return true;
        }

        SetDirty(dlg, state);
        PluginsPane::Refresh(host, state);
        return true;
    }

    if (commandId == IDC_PLUGINS_TEST)
    {
        if (notifyCode != BN_CLICKED)
        {
            return true;
        }

        HWND dlg = GetParent(host);
        if (! dlg)
        {
            return true;
        }

        const std::optional<PrefsPluginListItem> selected = TryGetActivePluginItem(state);
        if (! selected.has_value())
        {
            return true;
        }

        const std::wstring_view pluginId = GetPluginId(selected.value());
        if (pluginId.empty() || ! IsPluginLoadable(selected.value()))
        {
            return true;
        }

        HRESULT hr = E_FAIL;
        if (selected.value().type == PrefsPluginType::FileSystem)
        {
            hr = FileSystemPluginManager::GetInstance().TestPlugin(pluginId);
        }
        else
        {
            hr = ViewerPluginManager::GetInstance().TestPlugin(pluginId);
        }

        const UINT textId                = SUCCEEDED(hr) ? IDS_MSG_PLUGIN_TEST_OK : IDS_MSG_PLUGIN_TEST_FAILED;
        const HostAlertSeverity severity = SUCCEEDED(hr) ? HOST_ALERT_INFO : HOST_ALERT_ERROR;
        ShowDialogAlert(dlg, severity, LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER), LoadStringResource(nullptr, textId));
        return true;
    }

    if (commandId == IDC_PLUGINS_TEST_ALL)
    {
        if (notifyCode != BN_CLICKED)
        {
            return true;
        }

        HWND dlg = GetParent(host);
        if (! dlg)
        {
            return true;
        }

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

        ShowDialogAlert(dlg,
                        HOST_ALERT_INFO,
                        LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER),
                        FormatStringResource(nullptr, IDS_FMT_PLUGIN_TEST_ALL_RESULT, okCount, failCount));
        return true;
    }

    if (commandId == IDC_PREFS_PLUGINS_SEARCH_EDIT)
    {
        if (notifyCode == EN_CHANGE)
        {
            PluginsPane::Refresh(host, state);
        }
        return true;
    }

    if (commandId == IDC_PREFS_PLUGINS_CUSTOM_PATHS_ADD)
    {
        if (notifyCode != BN_CLICKED)
        {
            return true;
        }

        HWND dlg = GetParent(host);
        if (! dlg)
        {
            return true;
        }

        std::filesystem::path selectedPath;
        if (! TryBrowseCustomPluginPath(dlg, selectedPath))
        {
            return true;
        }

        if (! IsDllPath(selectedPath))
        {
            ShowDialogAlert(
                dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_INVALID));
            return true;
        }

        auto& customPaths = state.workingSettings.plugins.customPluginPaths;
        if (std::find(customPaths.begin(), customPaths.end(), selectedPath) == customPaths.end())
        {
            customPaths.push_back(selectedPath);
        }

        PluginsPane::Refresh(host, state);
        if (state.pluginsCustomPathsList)
        {
            const std::wstring selectedText = selectedPath.wstring();
            for (int i = 0; i < ListView_GetItemCount(state.pluginsCustomPathsList.get()); ++i)
            {
                std::array<wchar_t, 2048> buffer{};
                buffer[0] = L'\0';
                ListView_GetItemText(state.pluginsCustomPathsList.get(), i, 0, buffer.data(), static_cast<int>(buffer.size()));
                if (_wcsicmp(buffer.data(), selectedText.c_str()) == 0)
                {
                    ListView_SetItemState(state.pluginsCustomPathsList.get(), i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
                    ListView_EnsureVisible(state.pluginsCustomPathsList.get(), i, FALSE);
                    break;
                }
            }
        }
        SetDirty(dlg, state);
        return true;
    }

    if (commandId == IDC_PREFS_PLUGINS_CUSTOM_PATHS_REMOVE)
    {
        if (notifyCode != BN_CLICKED)
        {
            return true;
        }

        HWND dlg = GetParent(host);
        if (! dlg || ! state.pluginsCustomPathsList)
        {
            return true;
        }

        const int selected = ListView_GetNextItem(state.pluginsCustomPathsList.get(), -1, LVNI_SELECTED);
        if (selected < 0)
        {
            return true;
        }

        LVITEMW item{};
        item.mask  = LVIF_PARAM;
        item.iItem = selected;
        if (! ListView_GetItem(state.pluginsCustomPathsList.get(), &item))
        {
            return true;
        }

        const size_t pathIndex = static_cast<size_t>(item.lParam);
        auto& customPaths      = state.workingSettings.plugins.customPluginPaths;
        if (pathIndex >= customPaths.size())
        {
            return true;
        }

        customPaths.erase(customPaths.begin() + static_cast<std::ptrdiff_t>(pathIndex));
        PluginsPane::Refresh(host, state);
        SetDirty(dlg, state);
        return true;
    }

    return false;
}

void PluginsPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    if (state.pluginsSelectedPlugin.has_value())
    {
        state.refreshingPluginsPage = true;
        auto clearRefreshFlag       = wil::scope_exit([&]() noexcept { state.refreshingPluginsPage = false; });

        const PrefsPluginListItem pluginItem = state.pluginsSelectedPlugin.value();
        const std::wstring_view pluginId     = GetPluginId(pluginItem);

        const HWND parent                   = state.pluginsConfigureButton ? GetParent(state.pluginsConfigureButton.get()) : nullptr;
        const std::wstring previousEditorId = state.pluginsDetailsConfigPluginId;
        const bool hadEditor                = ! state.pluginsDetailsConfigFields.empty();
        static_cast<void>(PrefsPluginConfiguration::EnsureEditor(parent, state, pluginItem));
        const bool hasEditor = ! state.pluginsDetailsConfigFields.empty();

        if (state.pluginsDetailsConfigEdit)
        {
            std::wstring configText;

            if (! pluginId.empty())
            {
                const std::wstring pluginIdText(pluginId);
                const auto it = state.workingSettings.plugins.configurationByPluginId.find(pluginIdText);
                if (it != state.workingSettings.plugins.configurationByPluginId.end() && ! std::holds_alternative<std::monostate>(it->second.value))
                {
                    std::string configUtf8;
                    const HRESULT hr = Common::Settings::SerializeJsonValue(it->second, configUtf8);
                    if (SUCCEEDED(hr))
                    {
                        configText = Utf16FromUtf8(configUtf8);
                        if (configText.empty() && ! configUtf8.empty())
                        {
                            configText = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_CONFIG_UNAVAILABLE);
                        }
                    }
                    else
                    {
                        configText = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_CONFIG_UNAVAILABLE);
                    }
                }

                if (configText.empty() && IsPluginLoadable(pluginItem))
                {
                    std::string configUtf8;
                    HRESULT configHr = E_FAIL;
                    if (pluginItem.type == PrefsPluginType::FileSystem)
                    {
                        configHr = FileSystemPluginManager::GetInstance().GetConfiguration(pluginId, state.baselineSettings, configUtf8);
                    }
                    else
                    {
                        configHr = ViewerPluginManager::GetInstance().GetConfiguration(pluginId, state.baselineSettings, configUtf8);
                    }

                    if (SUCCEEDED(configHr))
                    {
                        if (configUtf8.empty())
                        {
                            // Some plugins report "no configuration" via an empty string / nullptr.
                            // Show a valid JSON object so the user can still see a concrete representation.
                            configText = L"{}";
                        }
                        else
                        {
                            configText = Utf16FromUtf8(configUtf8);
                        }
                    }
                    else
                    {
                        configText = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_CONFIG_UNAVAILABLE);
                    }
                }
            }

            if (configText.empty())
            {
                configText = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_CONFIG_DEFAULT);
            }

            SetWindowTextW(state.pluginsDetailsConfigEdit.get(), configText.c_str());
            SendMessageW(state.pluginsDetailsConfigEdit.get(), EM_SETSEL, 0, 0);
            SendMessageW(state.pluginsDetailsConfigEdit.get(), EM_SCROLLCARET, 0, 0);
        }

        if (previousEditorId != state.pluginsDetailsConfigPluginId || hadEditor != hasEditor)
        {
            RECT client{};
            if (GetClientRect(host, &client))
            {
                const int w = std::max(0l, client.right - client.left);
                const int h = std::max(0l, client.bottom - client.top);
                SendMessageW(host, WM_SIZE, SIZE_RESTORED, MAKELPARAM(w, h));
            }
            else
            {
                SendMessageW(host, WM_SIZE, SIZE_RESTORED, 0);
            }
        }

        UpdatePluginsActionButtonsEnabled(state);
        return;
    }

    if (! state.pluginsList)
    {
        return;
    }

    PrefsPluginConfiguration::Clear(state);
    if (state.pluginsDetailsConfigError)
    {
        SetWindowTextW(state.pluginsDetailsConfigError.get(), L"");
    }

    state.refreshingPluginsPage = true;
    auto clearRefreshFlag       = wil::scope_exit([&]() noexcept { state.refreshingPluginsPage = false; });

    const UINT dpi = GetDpiForWindow(host);

    std::wstring filterText;
    std::wstring_view filter;
    if (state.pluginsSearchEdit)
    {
        filterText = PrefsUi::GetWindowTextString(state.pluginsSearchEdit.get());
        filter     = PrefsUi::TrimWhitespace(filterText);
    }

    ThemedControls::ApplyThemeToListView(state.pluginsList, state.theme);
    EnsurePluginsListColumns(state.pluginsList.get(), dpi);
    ListView_SetExtendedListViewStyle(state.pluginsList.get(), LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP | LVS_EX_CHECKBOXES);

    std::wstring selectedPluginId;
    const int selected = ListView_GetNextItem(state.pluginsList.get(), -1, LVNI_SELECTED);
    if (selected >= 0)
    {
        LVITEMW item{};
        item.mask  = LVIF_PARAM;
        item.iItem = selected;
        if (ListView_GetItem(state.pluginsList.get(), &item))
        {
            const size_t rowIndex = static_cast<size_t>(item.lParam);
            if (rowIndex < state.pluginsListItems.size())
            {
                selectedPluginId.assign(GetPluginId(state.pluginsListItems[rowIndex]));
            }
        }
    }

    state.pluginsListItems.clear();
    ListView_DeleteAllItems(state.pluginsList.get());

    const auto& fsPlugins = FileSystemPluginManager::GetInstance().GetPlugins();
    state.pluginsListItems.reserve(fsPlugins.size());
    for (size_t i = 0; i < fsPlugins.size(); ++i)
    {
        if (! fsPlugins[i].id.empty())
        {
            state.pluginsListItems.push_back(PrefsPluginListItem{PrefsPluginType::FileSystem, i});
        }
    }

    const auto& viewerPlugins = ViewerPluginManager::GetInstance().GetPlugins();
    state.pluginsListItems.reserve(state.pluginsListItems.size() + viewerPlugins.size());
    for (size_t i = 0; i < viewerPlugins.size(); ++i)
    {
        if (! viewerPlugins[i].id.empty())
        {
            state.pluginsListItems.push_back(PrefsPluginListItem{PrefsPluginType::Viewer, i});
        }
    }

    std::sort(state.pluginsListItems.begin(),
              state.pluginsListItems.end(),
              [](const PrefsPluginListItem& a, const PrefsPluginListItem& b) noexcept
    {
        if (a.type != b.type)
        {
            return a.type < b.type;
        }

        const int aOrigin = GetPluginOriginOrder(a);
        const int bOrigin = GetPluginOriginOrder(b);
        if (aOrigin != bOrigin)
        {
            return aOrigin < bOrigin;
        }

        const std::wstring_view aName = GetPluginDisplayName(a);
        const std::wstring_view bName = GetPluginDisplayName(b);
        if (aName.empty() || bName.empty())
        {
            return aName < bName;
        }
        return _wcsicmp(aName.data(), bName.data()) < 0;
    });

    const std::wstring typeFileSystem = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_TYPE_FILE_SYSTEM);
    const std::wstring typeViewer     = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_TYPE_VIEWER);

    int insertPos = 0;
    for (size_t i = 0; i < state.pluginsListItems.size(); ++i)
    {
        const auto& row = state.pluginsListItems[i];

        const std::wstring_view nameText = GetPluginDisplayName(row);
        const std::wstring_view idText   = GetPluginId(row);
        const std::wstring_view shortId  = GetPluginShortIdOrId(row);
        const std::wstring originText    = GetPluginOriginText(row);

        if (nameText.empty() || idText.empty())
        {
            continue;
        }

        const std::wstring_view typeText = (row.type == PrefsPluginType::FileSystem) ? std::wstring_view(typeFileSystem) : std::wstring_view(typeViewer);
        if (! filter.empty() && ! PrefsUi::ContainsCaseInsensitive(nameText, filter) && ! PrefsUi::ContainsCaseInsensitive(idText, filter) &&
            ! PrefsUi::ContainsCaseInsensitive(shortId, filter) && ! PrefsUi::ContainsCaseInsensitive(typeText, filter) &&
            ! PrefsUi::ContainsCaseInsensitive(originText, filter))
        {
            continue;
        }

        LVITEMW item{};
        item.mask          = LVIF_TEXT | LVIF_PARAM;
        item.iItem         = insertPos;
        item.iSubItem      = 0;
        item.pszText       = const_cast<LPWSTR>(nameText.data());
        item.lParam        = static_cast<LPARAM>(i);
        const int inserted = ListView_InsertItem(state.pluginsList.get(), &item);
        if (inserted < 0)
        {
            continue;
        }

        ++insertPos;

        ListView_SetItemText(state.pluginsList.get(), inserted, kPluginsColumnType, const_cast<LPWSTR>(typeText.data()));
        ListView_SetItemText(state.pluginsList.get(), inserted, kPluginsColumnOrigin, const_cast<LPWSTR>(originText.c_str()));
        ListView_SetItemText(state.pluginsList.get(), inserted, kPluginsColumnId, const_cast<LPWSTR>(shortId.data()));

        const bool enabled = ! IsPluginDisabledInWorkingSettings(state, idText);
        SetPluginsListRowEnabled(state.pluginsList.get(), inserted, enabled);
    }

    UpdatePluginsListColumnWidths(state.pluginsList.get(), dpi);

    if (! selectedPluginId.empty())
    {
        for (int i = 0; i < ListView_GetItemCount(state.pluginsList.get()); ++i)
        {
            LVITEMW item{};
            item.mask  = LVIF_PARAM;
            item.iItem = i;
            if (! ListView_GetItem(state.pluginsList.get(), &item))
            {
                continue;
            }

            const size_t rowIndex = static_cast<size_t>(item.lParam);
            if (rowIndex >= state.pluginsListItems.size())
            {
                continue;
            }

            if (std::wstring_view(GetPluginId(state.pluginsListItems[rowIndex])) == selectedPluginId)
            {
                ListView_SetItemState(state.pluginsList.get(), i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
                ListView_EnsureVisible(state.pluginsList.get(), i, FALSE);
                break;
            }
        }
    }

    if (state.pluginsCustomPathsList)
    {
        std::wstring selectedPathText;
        const int selectedPathIndex = ListView_GetNextItem(state.pluginsCustomPathsList.get(), -1, LVNI_SELECTED);
        if (selectedPathIndex >= 0)
        {
            std::array<wchar_t, 2048> buffer{};
            buffer[0] = L'\0';
            ListView_GetItemText(state.pluginsCustomPathsList.get(), selectedPathIndex, 0, buffer.data(), static_cast<int>(buffer.size()));
            selectedPathText.assign(buffer.data());
        }

        ThemedControls::ApplyThemeToListView(state.pluginsCustomPathsList, state.theme);
        EnsurePluginsCustomPathsListColumns(state.pluginsCustomPathsList.get(), dpi);
        ListView_SetExtendedListViewStyle(state.pluginsCustomPathsList.get(), LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP);

        ListView_DeleteAllItems(state.pluginsCustomPathsList.get());

        for (size_t i = 0; i < state.workingSettings.plugins.customPluginPaths.size(); ++i)
        {
            const std::wstring pathText = state.workingSettings.plugins.customPluginPaths[i].wstring();
            LVITEMW item{};
            item.mask     = LVIF_TEXT | LVIF_PARAM;
            item.iItem    = static_cast<int>(i);
            item.iSubItem = 0;
            item.pszText  = const_cast<LPWSTR>(pathText.c_str());
            item.lParam   = static_cast<LPARAM>(i);
            ListView_InsertItem(state.pluginsCustomPathsList.get(), &item);
        }

        UpdatePluginsCustomPathsListColumnWidths(state.pluginsCustomPathsList.get());

        if (! selectedPathText.empty())
        {
            for (int i = 0; i < ListView_GetItemCount(state.pluginsCustomPathsList.get()); ++i)
            {
                std::array<wchar_t, 2048> buffer{};
                buffer[0] = L'\0';
                ListView_GetItemText(state.pluginsCustomPathsList.get(), i, 0, buffer.data(), static_cast<int>(buffer.size()));
                if (_wcsicmp(buffer.data(), selectedPathText.c_str()) == 0)
                {
                    ListView_SetItemState(state.pluginsCustomPathsList.get(), i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
                    ListView_EnsureVisible(state.pluginsCustomPathsList.get(), i, FALSE);
                    break;
                }
            }
        }

        const bool hasSelection = ListView_GetNextItem(state.pluginsCustomPathsList.get(), -1, LVNI_SELECTED) >= 0;
        if (state.pluginsCustomPathsRemoveButton)
        {
            EnableWindow(state.pluginsCustomPathsRemoveButton.get(), hasSelection ? TRUE : FALSE);
        }
    }

    UpdatePluginsActionButtonsEnabled(state);
}

bool PluginsPane::HandleNotify(HWND host, PreferencesDialogState& state, const NMHDR* hdr, LRESULT& /*outResult*/) noexcept
{
    if (! host || ! hdr)
    {
        return false;
    }

    if (state.refreshingPluginsPage)
    {
        return true;
    }

    if (state.pluginsCustomPathsList && hdr->hwndFrom == state.pluginsCustomPathsList.get())
    {
        if (hdr->code == LVN_ITEMCHANGED)
        {
            const bool hasSelection = ListView_GetNextItem(state.pluginsCustomPathsList.get(), -1, LVNI_SELECTED) >= 0;
            if (state.pluginsCustomPathsRemoveButton)
            {
                EnableWindow(state.pluginsCustomPathsRemoveButton.get(), hasSelection ? TRUE : FALSE);
            }

            PrefsPaneHost::EnsureControlVisible(host, state, state.pluginsCustomPathsList.get());
            return true;
        }

        return false;
    }

    if (! state.pluginsList || hdr->hwndFrom != state.pluginsList.get())
    {
        return false;
    }

    if (hdr->code != LVN_ITEMCHANGED)
    {
        return false;
    }

    const auto* nmlv = reinterpret_cast<const NMLISTVIEW*>(hdr);
    if (! nmlv || nmlv->iItem < 0 || (nmlv->uChanged & LVIF_STATE) == 0)
    {
        return true;
    }

    if ((nmlv->uOldState & LVIS_SELECTED) != (nmlv->uNewState & LVIS_SELECTED))
    {
        UpdatePluginsActionButtonsEnabled(state);
    }

    if ((nmlv->uOldState & LVIS_STATEIMAGEMASK) == (nmlv->uNewState & LVIS_STATEIMAGEMASK))
    {
        return true;
    }

    LVITEMW item{};
    item.mask  = LVIF_PARAM;
    item.iItem = nmlv->iItem;
    if (! ListView_GetItem(state.pluginsList.get(), &item))
    {
        return true;
    }

    const size_t rowIndex = static_cast<size_t>(item.lParam);
    if (rowIndex >= state.pluginsListItems.size())
    {
        return true;
    }

    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return true;
    }

    const PrefsPluginListItem& row   = state.pluginsListItems[rowIndex];
    const std::wstring_view pluginId = GetPluginId(row);
    if (pluginId.empty())
    {
        return true;
    }

    const bool enabled = (ListView_GetCheckState(state.pluginsList.get(), nmlv->iItem) != FALSE);
    if (! enabled && row.type == PrefsPluginType::FileSystem && pluginId == state.workingSettings.plugins.currentFileSystemPluginId)
    {
        state.refreshingPluginsPage = true;
        SetPluginsListRowEnabled(state.pluginsList.get(), nmlv->iItem, true);
        state.refreshingPluginsPage = false;

        ShowDialogAlert(dlg,
                        HOST_ALERT_WARNING,
                        LoadStringResource(nullptr, IDS_CAPTION_WARNING),
                        LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CANNOT_DISABLE_ACTIVE_FILE_SYSTEM));
        return true;
    }

    auto& disabled = state.workingSettings.plugins.disabledPluginIds;
    if (enabled)
    {
        disabled.erase(std::remove_if(disabled.begin(), disabled.end(), [&](const std::wstring& id) noexcept { return std::wstring_view(id) == pluginId; }),
                       disabled.end());
    }
    else
    {
        const auto it = std::find_if(disabled.begin(), disabled.end(), [&](const std::wstring& id) noexcept { return std::wstring_view(id) == pluginId; });
        if (it == disabled.end())
        {
            disabled.push_back(std::wstring(pluginId));
        }
    }

    SetDirty(dlg, state);
    return true;
}

void PluginsPane::LayoutControls(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, int sectionY, HFONT dialogFont) noexcept
{
    using namespace PrefsLayoutConstants;

    if (! host)
    {
        return;
    }

    RECT hostClient{};
    GetClientRect(host, &hostClient);
    const int hostBottom        = std::max(0l, hostClient.bottom - hostClient.top);
    const int hostContentBottom = std::max(0, hostBottom - margin);

    const UINT dpi        = GetDpiForWindow(host);
    const int rowHeight   = std::max(1, ThemedControls::ScaleDip(dpi, kRowHeightDip));
    const int labelHeight = std::max(1, ThemedControls::ScaleDip(dpi, kTitleHeightDip));
    const int gapX        = ThemedControls::ScaleDip(dpi, kToggleGapXDip);

    const int buttonHeight = rowHeight;
    const int buttonPadX   = ThemedControls::ScaleDip(dpi, kCardPaddingXDip);

    const auto measureButtonWidth = [&](HWND button, int minWidthDip) noexcept
    {
        if (! button)
        {
            return 0;
        }

        HFONT font = reinterpret_cast<HFONT>(SendMessageW(button, WM_GETFONT, 0, 0));
        if (! font)
        {
            font = dialogFont ? dialogFont : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        }

        const std::wstring text = PrefsUi::GetWindowTextString(button);
        const int textW         = ThemedControls::MeasureTextWidth(host, font, text);
        return std::max(ThemedControls::ScaleDip(dpi, minWidthDip), textW + 2 * buttonPadX);
    };

    const auto setVisible = [&](HWND hwnd, bool visible) noexcept
    {
        if (hwnd)
        {
            ShowWindow(hwnd, visible ? SW_SHOW : SW_HIDE);
        }
    };

    const bool showDetails = state.pluginsSelectedPlugin.has_value();
    if (showDetails)
    {
        const PrefsPluginListItem pluginItem     = state.pluginsSelectedPlugin.value();
        const std::wstring_view selectedPluginId = GetPluginId(pluginItem);

        if (! selectedPluginId.empty() && ! state.pluginsDetailsConfigPluginId.empty() && state.pluginsDetailsConfigPluginId != selectedPluginId)
        {
            PrefsPluginConfiguration::Clear(state);
            if (state.pluginsDetailsConfigError)
            {
                SetWindowTextW(state.pluginsDetailsConfigError.get(), L"");
            }
        }

        if (state.pluginsDetailsIdLabel && ! selectedPluginId.empty())
        {
            std::wstring formatted;
            formatted = FormatStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_ID_FMT, std::wstring(selectedPluginId));

            const std::wstring current = PrefsUi::GetWindowTextString(state.pluginsDetailsIdLabel.get());
            if (current != formatted)
            {
                SetWindowTextW(state.pluginsDetailsIdLabel.get(), formatted.c_str());
            }
        }

        const bool hasEditor =
            ! selectedPluginId.empty() && state.pluginsDetailsConfigPluginId == selectedPluginId && ! state.pluginsDetailsConfigFields.empty();

        bool showConfigError = false;
        if (state.pluginsDetailsConfigError)
        {
            showConfigError = ! PrefsUi::GetWindowTextString(state.pluginsDetailsConfigError.get()).empty();
        }

        setVisible(state.pluginsNote.get(), false);
        setVisible(state.pluginsSearchLabel.get(), false);
        setVisible(state.pluginsSearchFrame.get(), false);
        setVisible(state.pluginsSearchEdit.get(), false);
        setVisible(state.pluginsList.get(), false);
        setVisible(state.pluginsCustomPathsHeader.get(), false);
        setVisible(state.pluginsCustomPathsNote.get(), false);
        setVisible(state.pluginsCustomPathsList.get(), false);
        setVisible(state.pluginsCustomPathsAddButton.get(), false);
        setVisible(state.pluginsCustomPathsRemoveButton.get(), false);

        setVisible(state.pluginsConfigureButton.get(), false);
        setVisible(state.pluginsTestButton.get(), false);
        setVisible(state.pluginsTestAllButton.get(), false);
        setVisible(state.pluginsDetailsHint.get(), false);
        setVisible(state.pluginsDetailsIdLabel.get(), true);
        setVisible(state.pluginsDetailsConfigLabel.get(), false);
        setVisible(state.pluginsDetailsConfigError.get(), showConfigError);
        setVisible(state.pluginsDetailsConfigFrame.get(), false);
        setVisible(state.pluginsDetailsConfigEdit.get(), false);

        const HFONT infoFont = state.italicFont ? state.italicFont.get() : dialogFont;

        const int cardPaddingX = ThemedControls::ScaleDip(dpi, kCardPaddingXDip);
        const int cardPaddingY = ThemedControls::ScaleDip(dpi, kCardPaddingYDip);
        const int cardSpacingY = ThemedControls::ScaleDip(dpi, kCardSpacingYDip);

        const auto pushCard = [&](const RECT& card) noexcept { state.pageSettingCards.push_back(card); };

        if (state.pluginsDetailsIdLabel)
        {
            const std::wstring idText = PrefsUi::GetWindowTextString(state.pluginsDetailsIdLabel.get());
            const int measuredHeight  = idText.empty() ? 0 : PrefsUi::MeasureStaticTextHeight(host, dialogFont, width, idText);
            const int idHeight        = std::max(labelHeight, std::max(0, measuredHeight));

            SetWindowPos(state.pluginsDetailsIdLabel.get(), nullptr, x, y, width, idHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(state.pluginsDetailsIdLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            y += idHeight + sectionY;
        }
        else
        {
            y += sectionY;
        }

        if (showConfigError && ! hasEditor && state.pluginsDetailsConfigError)
        {
            const std::wstring errorText = PrefsUi::GetWindowTextString(state.pluginsDetailsConfigError.get());
            const int textWidth          = std::max(0, width - 2 * cardPaddingX);
            const int textHeight         = errorText.empty() ? 0 : PrefsUi::MeasureStaticTextHeight(host, infoFont, textWidth, errorText);
            const int cardHeight         = std::max(rowHeight + 2 * cardPaddingY, std::max(0, textHeight) + 2 * cardPaddingY);

            RECT card{};
            card.left   = x;
            card.top    = y;
            card.right  = x + width;
            card.bottom = y + cardHeight;
            pushCard(card);

            SetWindowPos(state.pluginsDetailsConfigError.get(),
                         nullptr,
                         x + cardPaddingX,
                         y + cardPaddingY,
                         textWidth,
                         std::max(0, textHeight),
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(state.pluginsDetailsConfigError.get(), WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
            y += cardHeight + cardSpacingY;
            return;
        }

        if (hasEditor)
        {
            PrefsPluginConfiguration::LayoutCards(host, state, x, y, width, dialogFont);
            return;
        }
        return;
    }

    PrefsPluginConfiguration::Clear(state);
    setVisible(state.pluginsDetailsHint.get(), false);
    setVisible(state.pluginsDetailsIdLabel.get(), false);
    setVisible(state.pluginsDetailsConfigLabel.get(), false);
    setVisible(state.pluginsDetailsConfigError.get(), false);
    setVisible(state.pluginsDetailsConfigFrame.get(), false);
    setVisible(state.pluginsDetailsConfigEdit.get(), false);

    setVisible(state.pluginsNote.get(), true);
    setVisible(state.pluginsSearchLabel.get(), true);
    setVisible(state.pluginsSearchFrame.get(), true);
    setVisible(state.pluginsSearchEdit.get(), true);
    setVisible(state.pluginsList.get(), true);
    setVisible(state.pluginsConfigureButton.get(), true);
    setVisible(state.pluginsTestButton.get(), true);
    setVisible(state.pluginsTestAllButton.get(), true);
    setVisible(state.pluginsCustomPathsHeader.get(), true);
    setVisible(state.pluginsCustomPathsNote.get(), true);
    setVisible(state.pluginsCustomPathsList.get(), true);
    setVisible(state.pluginsCustomPathsAddButton.get(), true);
    setVisible(state.pluginsCustomPathsRemoveButton.get(), true);

    const std::wstring noteText = PrefsUi::GetWindowTextString(state.pluginsNote.get());
    if (state.pluginsNote)
    {
        const int noteHeight = noteText.empty() ? 0 : PrefsUi::MeasureStaticTextHeight(host, dialogFont, width, noteText);
        SetWindowPos(state.pluginsNote.get(), nullptr, x, y, width, std::max(0, noteHeight), SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.pluginsNote.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        y += std::max(0, noteHeight) + sectionY;
    }

    const int searchLabelWidth   = std::min(width, ThemedControls::ScaleDip(dpi, 52));
    const int searchEditWidth    = std::max(0, width - searchLabelWidth - gapX);
    const int searchEditX        = x + searchLabelWidth + gapX;
    const int searchFramePadding = (state.pluginsSearchFrame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;
    if (state.pluginsSearchLabel)
    {
        SetWindowPos(
            state.pluginsSearchLabel.get(), nullptr, x, y + (rowHeight - labelHeight) / 2, searchLabelWidth, labelHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.pluginsSearchLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    if (state.pluginsSearchFrame)
    {
        SetWindowPos(state.pluginsSearchFrame.get(), nullptr, searchEditX, y, searchEditWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (state.pluginsSearchEdit)
    {
        SetWindowPos(state.pluginsSearchEdit.get(),
                     nullptr,
                     searchEditX + searchFramePadding,
                     y + searchFramePadding,
                     std::max(1, searchEditWidth - 2 * searchFramePadding),
                     std::max(1, rowHeight - 2 * searchFramePadding),
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.pluginsSearchEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    y += rowHeight + gapY;

    const int configureButtonWidth = std::min(width, measureButtonWidth(state.pluginsConfigureButton.get(), 120));
    const int testButtonWidth      = std::min(width, measureButtonWidth(state.pluginsTestButton.get(), 70));
    const int testAllButtonWidth   = std::min(width, measureButtonWidth(state.pluginsTestAllButton.get(), 90));

    const int buttonsRowWidth =
        configureButtonWidth + (testButtonWidth > 0 ? (gapX + testButtonWidth) : 0) + (testAllButtonWidth > 0 ? (gapX + testAllButtonWidth) : 0);
    const bool buttonsSingleRow = buttonsRowWidth > 0 && buttonsRowWidth <= width;

    int buttonsRowCount = 0;
    if (configureButtonWidth > 0)
    {
        ++buttonsRowCount;
    }
    if (testButtonWidth > 0)
    {
        ++buttonsRowCount;
    }
    if (testAllButtonWidth > 0)
    {
        ++buttonsRowCount;
    }
    if (buttonsSingleRow && buttonsRowCount > 1)
    {
        buttonsRowCount = 1;
    }

    const int actionsBlockHeight = (buttonsRowCount > 0) ? (gapY + (buttonsRowCount * buttonHeight) + ((buttonsRowCount - 1) * gapY) + sectionY) : sectionY;

    int customAddWidth    = std::min(width, measureButtonWidth(state.pluginsCustomPathsAddButton.get(), 70));
    int customRemoveWidth = std::min(width, measureButtonWidth(state.pluginsCustomPathsRemoveButton.get(), 70));
    if (customAddWidth > 0 && customRemoveWidth > 0 && (customAddWidth + gapX + customRemoveWidth > width))
    {
        customRemoveWidth = std::max(0, width - customAddWidth - gapX);
    }

    const HFONT headerFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const HFONT infoFont   = state.italicFont ? state.italicFont.get() : dialogFont;

    const std::wstring customNoteText = PrefsUi::GetWindowTextString(state.pluginsCustomPathsNote.get());
    const int customNoteHeight        = customNoteText.empty() ? 0 : PrefsUi::MeasureStaticTextHeight(host, infoFont, width, customNoteText);

    int customBlockHeight = labelHeight + gapY;
    if (customNoteHeight > 0)
    {
        customBlockHeight += customNoteHeight + gapY;
    }
    const int customListHeight = std::max(1, ThemedControls::ScaleDip(dpi, 90));
    customBlockHeight += customListHeight + gapY;

    const int pinnedCustomBtnsTop  = hostContentBottom - buttonHeight;
    const int minPluginsListHeight = std::max(1, ThemedControls::ScaleDip(dpi, 120));

    const int pinnedPluginsHeight = pinnedCustomBtnsTop - y - customBlockHeight - actionsBlockHeight;
    const bool pinnedLayout       = pinnedCustomBtnsTop >= y && pinnedPluginsHeight >= minPluginsListHeight;

    const int pluginsListTop         = y;
    const int reservedForActions     = actionsBlockHeight;
    const int preferredPluginsHeight = std::max(0, hostContentBottom - y - reservedForActions);
    const int pluginsListHeight      = pinnedLayout ? pinnedPluginsHeight : std::max(minPluginsListHeight, preferredPluginsHeight);

    if (state.pluginsList)
    {
        SetWindowPos(state.pluginsList.get(), nullptr, x, pluginsListTop, width, pluginsListHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.pluginsList.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        UpdatePluginsListColumnWidths(state.pluginsList.get(), dpi);
    }

    y += pluginsListHeight;

    y += gapY;
    if (buttonsSingleRow)
    {
        int currentX = x;
        if (state.pluginsConfigureButton && configureButtonWidth > 0)
        {
            SetWindowPos(
                state.pluginsConfigureButton.get(), nullptr, currentX, y, std::max(0, configureButtonWidth), buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(state.pluginsConfigureButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            currentX += configureButtonWidth + gapX;
        }
        if (state.pluginsTestButton && testButtonWidth > 0)
        {
            SetWindowPos(state.pluginsTestButton.get(), nullptr, currentX, y, std::max(0, testButtonWidth), buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(state.pluginsTestButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            currentX += testButtonWidth + gapX;
        }
        if (state.pluginsTestAllButton && testAllButtonWidth > 0)
        {
            SetWindowPos(state.pluginsTestAllButton.get(), nullptr, currentX, y, std::max(0, testAllButtonWidth), buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(state.pluginsTestAllButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        y += buttonHeight + sectionY;
    }
    else
    {
        int rows               = 0;
        const auto layoutStack = [&](HWND button, int buttonWidth) noexcept
        {
            if (! button || buttonWidth <= 0)
            {
                return;
            }

            SetWindowPos(button, nullptr, x, y, buttonWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(button, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            y += buttonHeight + gapY;
            ++rows;
        };

        layoutStack(state.pluginsConfigureButton.get(), configureButtonWidth);
        layoutStack(state.pluginsTestButton.get(), testButtonWidth);
        layoutStack(state.pluginsTestAllButton.get(), testAllButtonWidth);

        if (rows > 0)
        {
            y -= gapY;
        }
        y += sectionY;
    }

    if (state.pluginsCustomPathsHeader)
    {
        SetWindowPos(state.pluginsCustomPathsHeader.get(), nullptr, x, y, width, labelHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.pluginsCustomPathsHeader.get(), WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
    }

    y += labelHeight + gapY;

    if (state.pluginsCustomPathsNote)
    {
        SetWindowPos(state.pluginsCustomPathsNote.get(), nullptr, x, y, width, std::max(0, customNoteHeight), SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.pluginsCustomPathsNote.get(), WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
    }
    y += std::max(0, customNoteHeight);
    if (customNoteHeight > 0)
    {
        y += gapY;
    }

    if (state.pluginsCustomPathsList)
    {
        SetWindowPos(state.pluginsCustomPathsList.get(), nullptr, x, y, width, customListHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.pluginsCustomPathsList.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    y += customListHeight + gapY;

    const int customButtonsTop = pinnedLayout ? pinnedCustomBtnsTop : y;
    if (state.pluginsCustomPathsAddButton)
    {
        SetWindowPos(
            state.pluginsCustomPathsAddButton.get(), nullptr, x, customButtonsTop, std::max(0, customAddWidth), buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.pluginsCustomPathsAddButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    if (state.pluginsCustomPathsRemoveButton)
    {
        const int removeX = x + customAddWidth + gapX;
        SetWindowPos(state.pluginsCustomPathsRemoveButton.get(),
                     nullptr,
                     removeX,
                     customButtonsTop,
                     std::max(0, customRemoveWidth),
                     buttonHeight,
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.pluginsCustomPathsRemoveButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    y = customButtonsTop + buttonHeight;
}

void PluginsPane::CreateControls(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    const DWORD baseStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    const bool customButtons    = ! state.theme.systemHighContrast;
    const DWORD buttonStyle     = WS_CHILD | WS_VISIBLE | WS_TABSTOP | (customButtons ? BS_OWNERDRAW : 0U);
    const DWORD wrapStyle       = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;
    const DWORD listExStyle     = state.theme.systemHighContrast ? WS_EX_CLIENTEDGE : 0;
    const DWORD listStyle       = WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS;

    state.pluginsConfigureButton.reset(CreateWindowExW(0,
                                                       L"Button",
                                                       LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CONFIGURE_ELLIPSIS).c_str(),
                                                       buttonStyle,
                                                       0,
                                                       0,
                                                       10,
                                                       10,
                                                       parent,
                                                       reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PLUGINS_CONFIGURE)),
                                                       GetModuleHandleW(nullptr),
                                                       nullptr));
    if (state.pluginsConfigureButton)
    {
        EnableWindow(state.pluginsConfigureButton.get(), FALSE);
    }

    state.pluginsTestButton.reset(CreateWindowExW(0,
                                                  L"Button",
                                                  LoadStringResource(nullptr, IDS_BTN_TEST).c_str(),
                                                  buttonStyle,
                                                  0,
                                                  0,
                                                  10,
                                                  10,
                                                  parent,
                                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PLUGINS_TEST)),
                                                  GetModuleHandleW(nullptr),
                                                  nullptr));
    if (state.pluginsTestButton)
    {
        EnableWindow(state.pluginsTestButton.get(), FALSE);
    }

    state.pluginsTestAllButton.reset(CreateWindowExW(0,
                                                     L"Button",
                                                     LoadStringResource(nullptr, IDS_BTN_TEST_ALL).c_str(),
                                                     buttonStyle,
                                                     0,
                                                     0,
                                                     10,
                                                     10,
                                                     parent,
                                                     reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PLUGINS_TEST_ALL)),
                                                     GetModuleHandleW(nullptr),
                                                     nullptr));

    state.pluginsNote.reset(CreateWindowExW(0,
                                            L"Static",
                                            LoadStringResource(nullptr, IDS_PREFS_PLUGINS_NOTE).c_str(),
                                            wrapStyle,
                                            0,
                                            0,
                                            10,
                                            10,
                                            parent,
                                            nullptr,
                                            GetModuleHandleW(nullptr),
                                            nullptr));

    state.pluginsSearchLabel.reset(CreateWindowExW(0,
                                                   L"Static",
                                                   LoadStringResource(nullptr, IDS_PREFS_COMMON_SEARCH).c_str(),
                                                   baseStaticStyle,
                                                   0,
                                                   0,
                                                   10,
                                                   10,
                                                   parent,
                                                   nullptr,
                                                   GetModuleHandleW(nullptr),
                                                   nullptr));
    PrefsInput::CreateFramedEditBox(
        state, parent, state.pluginsSearchFrame, state.pluginsSearchEdit, IDC_PREFS_PLUGINS_SEARCH_EDIT, WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);
    if (state.pluginsSearchEdit)
    {
        SendMessageW(state.pluginsSearchEdit.get(), EM_SETLIMITTEXT, 128, 0);
    }

    state.pluginsList.reset(CreateWindowExW(listExStyle,
                                            WC_LISTVIEWW,
                                            L"",
                                            listStyle,
                                            0,
                                            0,
                                            10,
                                            10,
                                            parent,
                                            reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PLUGINS_LIST)),
                                            GetModuleHandleW(nullptr),
                                            nullptr));

    if (state.pluginsList)
    {
        const UINT dpi = GetDpiForWindow(state.pluginsList.get());
        ThemedControls::ApplyThemeToListView(state.pluginsList, state.theme);
        EnsurePluginsListColumns(state.pluginsList.get(), dpi);
        ListView_SetExtendedListViewStyle(state.pluginsList.get(), LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP | LVS_EX_CHECKBOXES);
    }

    state.pluginsCustomPathsHeader.reset(CreateWindowExW(0,
                                                         L"Static",
                                                         LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_HEADER).c_str(),
                                                         baseStaticStyle,
                                                         0,
                                                         0,
                                                         10,
                                                         10,
                                                         parent,
                                                         nullptr,
                                                         GetModuleHandleW(nullptr),
                                                         nullptr));

    state.pluginsCustomPathsNote.reset(CreateWindowExW(0,
                                                       L"Static",
                                                       LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_NOTE).c_str(),
                                                       wrapStyle,
                                                       0,
                                                       0,
                                                       10,
                                                       10,
                                                       parent,
                                                       nullptr,
                                                       GetModuleHandleW(nullptr),
                                                       nullptr));

    state.pluginsCustomPathsList.reset(CreateWindowExW(listExStyle,
                                                       WC_LISTVIEWW,
                                                       L"",
                                                       listStyle | LVS_NOCOLUMNHEADER,
                                                       0,
                                                       0,
                                                       10,
                                                       10,
                                                       parent,
                                                       reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PLUGINS_CUSTOM_PATHS_LIST)),
                                                       GetModuleHandleW(nullptr),
                                                       nullptr));

    state.pluginsCustomPathsAddButton.reset(CreateWindowExW(0,
                                                            L"Button",
                                                            LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_ADD_ELLIPSIS).c_str(),
                                                            buttonStyle,
                                                            0,
                                                            0,
                                                            10,
                                                            10,
                                                            parent,
                                                            reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PLUGINS_CUSTOM_PATHS_ADD)),
                                                            GetModuleHandleW(nullptr),
                                                            nullptr));
    state.pluginsCustomPathsRemoveButton.reset(CreateWindowExW(0,
                                                               L"Button",
                                                               LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_REMOVE).c_str(),
                                                               buttonStyle,
                                                               0,
                                                               0,
                                                               10,
                                                               10,
                                                               parent,
                                                               reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PLUGINS_CUSTOM_PATHS_REMOVE)),
                                                               GetModuleHandleW(nullptr),
                                                               nullptr));
    if (state.pluginsCustomPathsRemoveButton)
    {
        EnableWindow(state.pluginsCustomPathsRemoveButton.get(), FALSE);
    }

    state.pluginsDetailsHint.reset(CreateWindowExW(0,
                                                   L"Static",
                                                   LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_HINT).c_str(),
                                                   wrapStyle,
                                                   0,
                                                   0,
                                                   10,
                                                   10,
                                                   parent,
                                                   nullptr,
                                                   GetModuleHandleW(nullptr),
                                                   nullptr));
    state.pluginsDetailsIdLabel.reset(CreateWindowExW(0, L"Static", L"", wrapStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.pluginsDetailsConfigLabel.reset(CreateWindowExW(0,
                                                          L"Static",
                                                          LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_CONFIG_LABEL).c_str(),
                                                          baseStaticStyle,
                                                          0,
                                                          0,
                                                          10,
                                                          10,
                                                          parent,
                                                          nullptr,
                                                          GetModuleHandleW(nullptr),
                                                          nullptr));
    state.pluginsDetailsConfigError.reset(CreateWindowExW(0, L"Static", L"", wrapStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    state.pluginsDetailsConfigFrame,
                                    state.pluginsDetailsConfigEdit,
                                    IDC_PREFS_PLUGINS_DETAILS_CONFIG_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | ES_NOHIDESEL);

    const std::array<HWND, 6> detailControls = {
        state.pluginsDetailsHint.get(),
        state.pluginsDetailsIdLabel.get(),
        state.pluginsDetailsConfigLabel.get(),
        state.pluginsDetailsConfigError.get(),
        state.pluginsDetailsConfigFrame.get(),
        state.pluginsDetailsConfigEdit.get(),
    };
    for (HWND hwnd : detailControls)
    {
        if (hwnd)
        {
            ShowWindow(hwnd, SW_HIDE);
        }
    }
}
