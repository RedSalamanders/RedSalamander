// Preferences.Plugins.cpp

#include "Framework.h"

#include "DxUi/DxUi.Typography.h"
#include "Preferences.Plugin.Configuration.h"
#include "Preferences.Plugins.h"

#include <algorithm>
#include <array>
#include <cerrno>
#include <cmath>
#include <cstdlib>
#include <cwchar>
#include <limits>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include <commdlg.h>
#include <shobjidl.h>

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "FileSystemPluginManager.h"
#include "FluentIcons.h"
#include "Helpers.h"
#include "HostServices.h"
#include "UiMetrics.h"
#include "ViewerPluginManager.h"
#include "WindowMessages.h"

#include "resource.h"

#include <unordered_map>

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::WindowHost;

constexpr int kPluginsColumnName            = 0;
constexpr int kPluginsColumnType            = 1;
constexpr int kPluginsColumnOrigin          = 2;
constexpr int kPluginsColumnId              = 3;
constexpr int kPluginsCustomPathsColumnPath = 0;
#ifdef ENABLE_TESTS
enum class DebugCustomPluginBrowseResultKind
{
    Path,
    Cancel,
};

struct DebugCustomPluginBrowseResult
{
    DebugCustomPluginBrowseResultKind kind = DebugCustomPluginBrowseResultKind::Path;
    std::filesystem::path path{};
};

std::mutex g_debugCustomPluginBrowseResultMutex;
std::optional<DebugCustomPluginBrowseResult> g_debugNextCustomPluginBrowseResult;
#endif
[[nodiscard]] uint64_t MakeStableRowId(std::wstring_view pluginId) noexcept
{
    constexpr uint64_t kFNVOffset = 1469598103934665603ull;
    constexpr uint64_t kFNVPrime  = 1099511628211ull;

    uint64_t value = kFNVOffset;
    for (const wchar_t ch : pluginId)
    {
        value ^= static_cast<uint64_t>(std::towlower(static_cast<wint_t>(ch)));
        value *= kFNVPrime;
    }
    return value;
}

struct PluginsDxHostBase
{
    PluginsDxHostBase() = default;
    ~PluginsDxHostBase() noexcept
    {
        DetachBase();
    }
    PluginsDxHostBase(const PluginsDxHostBase&)            = delete;
    PluginsDxHostBase& operator=(const PluginsDxHostBase&) = delete;
    PluginsDxHostBase(PluginsDxHostBase&&)                 = delete;
    PluginsDxHostBase& operator=(PluginsDxHostBase&&)      = delete;

    wil::unique_hwnd hostHwnd;
    WindowHost host;

    void DetachBase() noexcept
    {
        host.Detach();
        PrefsDxHost::ResetOwnedHostWindow(hostHwnd);
    }
};

struct PluginsLabelDx : PluginsDxHostBase
{
    PluginsLabelDx()                                 = default;
    PluginsLabelDx(const PluginsLabelDx&)            = delete;
    PluginsLabelDx& operator=(const PluginsLabelDx&) = delete;
    PluginsLabelDx(PluginsLabelDx&&)                 = delete;
    PluginsLabelDx& operator=(PluginsLabelDx&&)      = delete;

    Label* label = nullptr;

    void Detach() noexcept
    {
        label = nullptr;
        DetachBase();
    }
};

struct PluginsButtonDx : PluginsDxHostBase
{
    PluginsButtonDx()                                  = default;
    PluginsButtonDx(const PluginsButtonDx&)            = delete;
    PluginsButtonDx& operator=(const PluginsButtonDx&) = delete;
    PluginsButtonDx(PluginsButtonDx&&)                 = delete;
    PluginsButtonDx& operator=(PluginsButtonDx&&)      = delete;

    Button* button = nullptr;

    void Detach() noexcept
    {
        button = nullptr;
        DetachBase();
    }
};

struct PluginsTextFieldDx : PluginsDxHostBase
{
    PluginsTextFieldDx()                                     = default;
    PluginsTextFieldDx(const PluginsTextFieldDx&)            = delete;
    PluginsTextFieldDx& operator=(const PluginsTextFieldDx&) = delete;
    PluginsTextFieldDx(PluginsTextFieldDx&&)                 = delete;
    PluginsTextFieldDx& operator=(PluginsTextFieldDx&&)      = delete;

    TextField* textField = nullptr;
    HWND legacyFrame     = nullptr;

    void Detach() noexcept
    {
        textField   = nullptr;
        legacyFrame = nullptr;
        DetachBase();
    }
};

struct PluginsGridRow
{
    uint64_t stableId = 0u;
    std::wstring pluginId;
    std::wstring name;
    std::wstring typeText;
    std::wstring originText;
    std::wstring shortIdText;
    bool enabled = true;
};

class PluginsGridModel final : public IDxGridModel
{
public:
    PluginsGridModel()
    {
        _columns = {
            {L"name", LoadStringResource(nullptr, IDS_PREFS_PLUGINS_COL_NAME), 240.0f, 140.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
            {L"type", LoadStringResource(nullptr, IDS_PREFS_PLUGINS_COL_TYPE), 120.0f, 90.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
            {L"origin", LoadStringResource(nullptr, IDS_PREFS_PLUGINS_COL_ORIGIN), 120.0f, 90.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
            {L"id", LoadStringResource(nullptr, IDS_PREFS_PLUGINS_COL_ID), 220.0f, 120.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
        };
    }

    void SetRows(std::vector<PluginsGridRow> rows)
    {
        _rows = std::move(rows);
        _rowIndexByStableId.clear();
        _rowIndexByStableId.reserve(_rows.size());
        for (size_t rowIndex = 0u; rowIndex < _rows.size(); ++rowIndex)
        {
            _rowIndexByStableId[_rows[rowIndex].stableId] = rowIndex;
        }
    }

    [[nodiscard]] const std::vector<PluginsGridRow>& GetRows() const noexcept
    {
        return _rows;
    }

    [[nodiscard]] std::optional<size_t> FindRowIndexByPluginId(std::wstring_view pluginId) const noexcept
    {
        const auto it = std::find_if(_rows.begin(), _rows.end(), [&](const PluginsGridRow& row) noexcept { return row.pluginId == pluginId; });
        if (it == _rows.end())
        {
            return std::nullopt;
        }
        return static_cast<size_t>(std::distance(_rows.begin(), it));
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rows.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return _columns.at(columnIndex);
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        if (rowIndex >= _rows.size())
        {
            return;
        }

        const PluginsGridRow& row = _rows[rowIndex];
        switch (columnIndex)
        {
            case kPluginsColumnName:
                outCell.kind    = RedSalamander::DxUi::GridCellKind::Checkbox;
                outCell.checked = row.enabled;
                outCell.text    = row.name;
                break;
            case kPluginsColumnType: outCell.text = row.typeText; break;
            case kPluginsColumnOrigin: outCell.text = row.originText; break;
            case kPluginsColumnId: outCell.text = row.shortIdText; break;
            default: break;
        }
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        if (rowIndex >= _rows.size())
        {
            return 0u;
        }
        return _rows[rowIndex].stableId;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        const auto it = _rowIndexByStableId.find(rowId);
        if (it == _rowIndexByStableId.end())
        {
            return std::nullopt;
        }

        return it->second;
    }

private:
    std::vector<GridColumnDesc> _columns;
    std::vector<PluginsGridRow> _rows;
    std::unordered_map<uint64_t, size_t> _rowIndexByStableId;
};

struct PluginsCustomPathGridRow
{
    uint64_t stableId = 0u;
    std::wstring path;
};

class PluginsCustomPathsGridModel final : public IDxGridModel
{
public:
    PluginsCustomPathsGridModel()
    {
        _columns = {
            {L"path", L"", 320.0f, 180.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
        };
    }

    void SetRows(std::vector<PluginsCustomPathGridRow> rows)
    {
        _rows = std::move(rows);
        _rowIndexByStableId.clear();
        _rowIndexByStableId.reserve(_rows.size());
        for (size_t rowIndex = 0u; rowIndex < _rows.size(); ++rowIndex)
        {
            _rowIndexByStableId[_rows[rowIndex].stableId] = rowIndex;
        }
    }

    [[nodiscard]] const std::vector<PluginsCustomPathGridRow>& GetRows() const noexcept
    {
        return _rows;
    }

    [[nodiscard]] std::optional<size_t> FindRowIndexByPath(std::wstring_view path) const noexcept
    {
        const std::wstring target(path);
        const auto it = std::find_if(
            _rows.begin(), _rows.end(), [&](const PluginsCustomPathGridRow& row) noexcept { return _wcsicmp(row.path.c_str(), target.c_str()) == 0; });
        if (it == _rows.end())
        {
            return std::nullopt;
        }
        return static_cast<size_t>(std::distance(_rows.begin(), it));
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rows.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return _columns.at(columnIndex);
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        if (rowIndex >= _rows.size() || columnIndex != kPluginsCustomPathsColumnPath)
        {
            return;
        }

        outCell.text = _rows[rowIndex].path;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        if (rowIndex >= _rows.size())
        {
            return 0u;
        }
        return _rows[rowIndex].stableId;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        const auto it = _rowIndexByStableId.find(rowId);
        if (it == _rowIndexByStableId.end())
        {
            return std::nullopt;
        }

        return it->second;
    }

private:
    std::vector<GridColumnDesc> _columns;
    std::vector<PluginsCustomPathGridRow> _rows;
    std::unordered_map<uint64_t, size_t> _rowIndexByStableId;
};

struct PluginsDxPage
{
    PluginsDxPage()                                = default;
    PluginsDxPage(const PluginsDxPage&)            = delete;
    PluginsDxPage& operator=(const PluginsDxPage&) = delete;
    PluginsDxPage(PluginsDxPage&&)                 = delete;
    PluginsDxPage& operator=(PluginsDxPage&&)      = delete;

    Label* note           = nullptr;
    Label* searchLabel    = nullptr;
    TextField* searchEdit = nullptr;
    Grid* listControl     = nullptr;
    std::unique_ptr<IDxGridModel> listModelStorage;
    PluginsGridModel* listModel  = nullptr;
    Button* configureButton      = nullptr;
    Button* testButton           = nullptr;
    Button* testAllButton        = nullptr;
    Label* detailsIdLabel        = nullptr;
    Label* detailsConfigError    = nullptr;
    Label* customPathsHeader     = nullptr;
    Label* customPathsNote       = nullptr;
    Grid* customPathsListControl = nullptr;
    std::unique_ptr<IDxGridModel> customPathsListModelStorage;
    PluginsCustomPathsGridModel* customPathsListModel = nullptr;
    Button* customPathsAddButton                      = nullptr;
    Button* customPathsRemoveButton                   = nullptr;

    void Detach() noexcept
    {
        note        = nullptr;
        searchLabel = nullptr;
        searchEdit  = nullptr;
        listControl = nullptr;
        listModelStorage.reset();
        listModel              = nullptr;
        configureButton        = nullptr;
        testButton             = nullptr;
        testAllButton          = nullptr;
        detailsIdLabel         = nullptr;
        detailsConfigError     = nullptr;
        customPathsHeader      = nullptr;
        customPathsNote        = nullptr;
        customPathsListControl = nullptr;
        customPathsListModelStorage.reset();
        customPathsListModel    = nullptr;
        customPathsAddButton    = nullptr;
        customPathsRemoveButton = nullptr;
    }
};

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

[[nodiscard]] bool TryBrowseCustomPluginPath(HWND owner, std::filesystem::path& outPath) noexcept
{
    outPath.clear();

#ifdef ENABLE_TESTS
    {
        std::scoped_lock lock(g_debugCustomPluginBrowseResultMutex);
        if (g_debugNextCustomPluginBrowseResult.has_value())
        {
            const DebugCustomPluginBrowseResult result = *g_debugNextCustomPluginBrowseResult;
            g_debugNextCustomPluginBrowseResult.reset();
            if (result.kind == DebugCustomPluginBrowseResultKind::Cancel)
            {
                return false;
            }

            outPath = result.path;
            return ! outPath.empty();
        }
    }
#endif

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

#ifdef ENABLE_TESTS
bool DebugSetPreferencesPluginsNextCustomPathBrowsePathImpl(const std::wstring_view path) noexcept
{
    std::scoped_lock lock(g_debugCustomPluginBrowseResultMutex);
    if (path.empty())
    {
        g_debugNextCustomPluginBrowseResult.reset();
        return true;
    }

    g_debugNextCustomPluginBrowseResult = DebugCustomPluginBrowseResult{.kind = DebugCustomPluginBrowseResultKind::Path, .path = std::filesystem::path(path)};
    return true;
}

bool DebugCancelPreferencesPluginsNextCustomPathBrowseImpl() noexcept
{
    std::scoped_lock lock(g_debugCustomPluginBrowseResultMutex);
    g_debugNextCustomPluginBrowseResult = DebugCustomPluginBrowseResult{.kind = DebugCustomPluginBrowseResultKind::Cancel};
    return true;
}
#endif

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

[[nodiscard]] PreferencesEmptyStateSpec GetPluginNoSettingsEmptyState(const PrefsPluginListItem& item) noexcept
{
    return PreferencesEmptyStateSpec{
        .iconGlyph         = FluentIcons::kPuzzle,
        .fallbackIconGlyph = L'\u25A1',
        .tone              = PreferencesEmptyStateSpec::Tone::Neutral,
        .title             = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_EMPTY_TITLE),
        .body              = FormatStringResource(nullptr, IDS_PREFS_PLUGINS_EMPTY_BODY, std::wstring(GetPluginDisplayName(item))),
        .caption           = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_EMPTY_CAPTION),
    };
}

[[nodiscard]] PreferencesEmptyStateSpec GetPluginMessageState(const PrefsInlineMessageSeverity severity,
                                                              std::wstring title,
                                                              std::wstring body,
                                                              std::wstring caption = {}) noexcept
{
    PreferencesEmptyStateSpec spec{};
    spec.title   = std::move(title);
    spec.body    = std::move(body);
    spec.caption = std::move(caption);

    switch (severity)
    {
        case PrefsInlineMessageSeverity::Warning:
            spec.iconGlyph         = FluentIcons::kWarning;
            spec.fallbackIconGlyph = FluentIcons::kFallbackWarning;
            spec.tone              = PreferencesEmptyStateSpec::Tone::Warning;
            break;
        case PrefsInlineMessageSeverity::Error:
            spec.iconGlyph         = FluentIcons::kError;
            spec.fallbackIconGlyph = FluentIcons::kFallbackError;
            spec.tone              = PreferencesEmptyStateSpec::Tone::Error;
            break;
        case PrefsInlineMessageSeverity::Info:
        default:
            spec.iconGlyph         = FluentIcons::kInfo;
            spec.fallbackIconGlyph = L'i';
            spec.tone              = PreferencesEmptyStateSpec::Tone::Info;
            break;
    }

    return spec;
}

void ClearPluginsStatusMessage(PreferencesDialogState& state) noexcept
{
    state.pluginsStatusTitleText.clear();
    state.pluginsStatusBodyText.clear();
    state.pluginsStatusSeverity = PrefsInlineMessageSeverity::Info;
}

void SetPluginsStatusMessage(PreferencesDialogState& state, HostAlertSeverity severity, std::wstring title, std::wstring body) noexcept
{
    switch (severity)
    {
        case HOST_ALERT_WARNING: state.pluginsStatusSeverity = PrefsInlineMessageSeverity::Warning; break;
        case HOST_ALERT_ERROR: state.pluginsStatusSeverity = PrefsInlineMessageSeverity::Error; break;
        case HOST_ALERT_BUSY:
        case HOST_ALERT_INFO:
        default: state.pluginsStatusSeverity = PrefsInlineMessageSeverity::Info; break;
    }
    state.pluginsStatusTitleText = std::move(title);
    state.pluginsStatusBodyText  = std::move(body);
}

void FlushPluginsFeedbackUi(HWND host) noexcept
{
    const HWND dlg    = host ? GetParent(host) : nullptr;
    const HWND target = (dlg && IsWindow(dlg) != FALSE) ? dlg : host;
    if (target && IsWindow(target) != FALSE)
    {
        RedrawWindow(target, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN | RDW_UPDATENOW);
    }
}

[[nodiscard]] std::wstring JoinPluginNames(const std::vector<std::wstring>& names) noexcept
{
    std::wstring result;
    for (const std::wstring& name : names)
    {
        if (name.empty())
        {
            continue;
        }

        if (! result.empty())
        {
            result.append(L"\r\n");
        }
        result.append(name);
    }
    return result;
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
    // Look up by the stable selected plugin ID tracked in state.
    if (! state.pluginsSelectedPluginId.empty())
    {
        return PrefsPlugins::FindItemById(state.pluginsSelectedPluginId);
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<PrefsPluginListItem> TryGetStableSelectedPluginItem(const PreferencesDialogState& state) noexcept
{
    if (! state.pluginsSelectedPluginId.empty())
    {
        if (state.pluginsSelectedPlugin.has_value() && GetPluginId(state.pluginsSelectedPlugin.value()) == state.pluginsSelectedPluginId)
        {
            return state.pluginsSelectedPlugin;
        }

        return PrefsPlugins::FindItemById(state.pluginsSelectedPluginId);
    }

    if (state.pluginsSelectedPlugin.has_value() && GetPluginId(state.pluginsSelectedPlugin.value()).empty())
    {
        return std::nullopt;
    }

    return state.pluginsSelectedPlugin;
}

void RefreshSelectedPluginCache(PreferencesDialogState& state) noexcept
{
    const std::optional<PrefsPluginListItem> selected = TryGetStableSelectedPluginItem(state);
    if (! selected.has_value())
    {
        state.pluginsSelectedPlugin.reset();
        state.pluginsSelectedPluginId.clear();
        return;
    }

    state.pluginsSelectedPlugin = selected;
    if (state.pluginsSelectedPluginId.empty())
    {
        state.pluginsSelectedPluginId.assign(GetPluginId(selected.value()));
    }
    if (state.pluginsRetainedSelectedPluginId.empty())
    {
        state.pluginsRetainedSelectedPluginId.assign(GetPluginId(selected.value()));
    }
}

[[nodiscard]] std::optional<PrefsPluginListItem> TryGetActivePluginItem(const PreferencesDialogState& state) noexcept
{
    if (const std::optional<PrefsPluginListItem> selected = TryGetStableSelectedPluginItem(state); selected.has_value())
    {
        return selected;
    }
    return TryGetSelectedPluginItem(state);
}

void UpdatePluginsActionButtonsEnabled(const PreferencesDialogState& /*state*/) noexcept
{
    // Button enabled states are now managed through DxUi in SyncDxControlsFromState.
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

[[nodiscard]] bool TryApplyPluginEnabledStateChange(HWND dlg, PreferencesDialogState& state, const PrefsPluginListItem& row, const bool enabled) noexcept
{
    const std::wstring_view pluginId = GetPluginId(row);
    if (pluginId.empty())
    {
        return false;
    }

    if (! enabled && row.type == PrefsPluginType::FileSystem && pluginId == state.workingSettings.plugins.currentFileSystemPluginId)
    {
        if (dlg)
        {
            ShowDialogAlert(dlg,
                            HOST_ALERT_WARNING,
                            LoadStringResource(nullptr, IDS_CAPTION_WARNING),
                            LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CANNOT_DISABLE_ACTIVE_FILE_SYSTEM));
        }
        return false;
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

    if (dlg)
    {
        SetDirty(dlg, state);
    }
    return true;
}

} // namespace

struct PluginsPane::DxState
{
    DxState()                          = default;
    DxState(const DxState&)            = delete;
    DxState& operator=(const DxState&) = delete;
    DxState(DxState&&)                 = delete;
    DxState& operator=(DxState&&)      = delete;

    PluginsDxPage page;

    void Detach() noexcept
    {
        page.Detach();
    }
};

PluginsPane::PluginsPane() = default;

PluginsPane::~PluginsPane()
{
    DetachDxHosts();
}

void PluginsPane::Destroy(PreferencesDialogState& state) noexcept
{
    _state = &state;
    PrefsPluginConfiguration::Clear(state);
    DetachDxHosts();

    state.pluginsListItems.clear();
    _pageHost = nullptr;
}

void PluginsPane::SyncDxListSelectionFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _dxState->page.listControl || ! _dxState->page.listModel)
    {
        return;
    }

    std::wstring selectedPluginId;
    if (! state.pluginsSelectedPluginId.empty())
    {
        selectedPluginId = state.pluginsSelectedPluginId;
    }
    else if (state.pluginsSelectedPlugin.has_value())
    {
        selectedPluginId.assign(GetPluginId(state.pluginsSelectedPlugin.value()));
    }

    if (selectedPluginId.empty())
    {
        _dxState->page.listControl->GetSelectionModel().Clear();
        if (_pageHostDx)
        {
            _pageHostDx->Invalidate();
        }
        return;
    }

    const auto dxRowIndex = _dxState->page.listModel->FindRowIndexByPluginId(selectedPluginId);
    if (! dxRowIndex.has_value())
    {
        _dxState->page.listControl->GetSelectionModel().Clear();
        if (_pageHostDx)
        {
            _pageHostDx->Invalidate();
        }
        return;
    }

    _dxState->page.listControl->GetSelectionModel().SetSingle(_dxState->page.listModel->GetStableRowId(dxRowIndex.value()));
    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void PluginsPane::SyncDxCustomPathsSelectionFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _dxState->page.customPathsListControl || ! _dxState->page.customPathsListModel)
    {
        return;
    }

    const std::wstring selectedPath = state.pluginsSelectedCustomPathText;

    if (selectedPath.empty())
    {
        _dxState->page.customPathsListControl->GetSelectionModel().Clear();
        if (_pageHostDx)
        {
            _pageHostDx->Invalidate();
        }
        return;
    }

    const auto dxRowIndex = _dxState->page.customPathsListModel->FindRowIndexByPath(selectedPath);
    if (! dxRowIndex.has_value())
    {
        _dxState->page.customPathsListControl->GetSelectionModel().Clear();
        if (_pageHostDx)
        {
            _pageHostDx->Invalidate();
        }
        return;
    }

    _dxState->page.customPathsListControl->GetSelectionModel().SetSingle(_dxState->page.customPathsListModel->GetStableRowId(dxRowIndex.value()));
    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

bool PluginsPane::EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept
{
    _pageHostDx      = state.pageHostDxHost;
    _pageContentRoot = state.pageHostDxContentRootControl;
    if (! _pageHostDx || ! _pageContentRoot)
    {
        Debug::Error(L"Preferences.Plugins: Shared page-host DX surface is unavailable; DxUi controls cannot be created.");
        return false;
    }

    if (_dxState && PrefsUi::HasRetainedDxChildren(_pageContentRoot))
    {
        _state      = &state;
        _hostWindow = parent;
        ApplyDxTheme(state);
        SyncDxControlsFromState(state);
        return true;
    }

    // Children were cleared by ResetPreferencesSharedPageSurface — the
    // config panel pointer in shared state is dangling.  Null it before
    // ClearChildren (which is harmless on an already-empty root) so that
    // PrefsPluginConfiguration::Clear won't dereference freed memory.
    state.pluginsDetailsConfigDxPanel = nullptr;
    state.pluginsDetailsConfigFields.clear();
    state.pluginsDetailsConfigPluginId.clear();
    state.pluginsDetailsConfigErrorText.clear();
    ClearPluginsStatusMessage(state);

    auto dxState = std::make_unique<DxState>();
    _pageHostDx->ResetInteractionState();
    _pageContentRoot->ClearChildren();

    dxState->page.note                    = _pageContentRoot->AddChild<Label>();
    dxState->page.searchLabel             = _pageContentRoot->AddChild<Label>();
    dxState->page.searchEdit              = _pageContentRoot->AddChild<TextField>();
    dxState->page.listControl             = _pageContentRoot->AddChild<Grid>();
    dxState->page.configureButton         = _pageContentRoot->AddChild<Button>();
    dxState->page.testButton              = _pageContentRoot->AddChild<Button>();
    dxState->page.testAllButton           = _pageContentRoot->AddChild<Button>();
    dxState->page.detailsIdLabel          = _pageContentRoot->AddChild<Label>();
    dxState->page.detailsConfigError      = _pageContentRoot->AddChild<Label>();
    dxState->page.customPathsHeader       = _pageContentRoot->AddChild<Label>();
    dxState->page.customPathsNote         = _pageContentRoot->AddChild<Label>();
    dxState->page.customPathsListControl  = _pageContentRoot->AddChild<Grid>();
    dxState->page.customPathsAddButton    = _pageContentRoot->AddChild<Button>();
    dxState->page.customPathsRemoveButton = _pageContentRoot->AddChild<Button>();

    dxState->page.note->SetFontRole(FontRole::Small);
    dxState->page.note->SetMultiline(true);
    dxState->page.detailsIdLabel->SetMultiline(true);
    dxState->page.detailsConfigError->SetFontRole(FontRole::Small);
    dxState->page.detailsConfigError->SetMultiline(true);
    dxState->page.customPathsHeader->SetFontRole(FontRole::Header);
    dxState->page.customPathsNote->SetFontRole(FontRole::Small);
    dxState->page.customPathsNote->SetMultiline(true);

    dxState->page.listControl->SetDelegate(this);
    dxState->page.listControl->SetSelectionMode(GridSelectionMode::Single);
    dxState->page.listControl->SetHeaderHeightDip(30.0f);
    dxState->page.listControl->SetRowHeightDip(30.0f);
    dxState->page.listControl->SetLineClamp(1u);

    dxState->page.customPathsListControl->SetDelegate(this);
    dxState->page.customPathsListControl->SetSelectionMode(GridSelectionMode::Single);
    dxState->page.customPathsListControl->SetHeaderHeightDip(0.0f);
    dxState->page.customPathsListControl->SetRowHeightDip(28.0f);
    dxState->page.customPathsListControl->SetLineClamp(1u);
    dxState->page.customPathsListControl->SetEmptyStateText(LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_EMPTY));

    auto listModel          = std::make_unique<PluginsGridModel>();
    dxState->page.listModel = listModel.get();
    dxState->page.listControl->SetModel(dxState->page.listModel);
    dxState->page.listModelStorage = std::move(listModel);

    auto customPathsModel              = std::make_unique<PluginsCustomPathsGridModel>();
    dxState->page.customPathsListModel = customPathsModel.get();
    dxState->page.customPathsListControl->SetModel(dxState->page.customPathsListModel);
    dxState->page.customPathsListModelStorage = std::move(customPathsModel);

    dxState->page.searchEdit->SetOnTextChanged([this, parent](std::wstring_view text) noexcept
    {
        if (_syncingDxInputs || ! _state)
        {
            return;
        }

        _state->pluginsSearchText.assign(text);
        if (parent && IsWindow(parent) != FALSE)
        {
            static_cast<void>(PrefsUi::PostDeferredAction(parent, PreferencesDeferredActionKind::PluginsSearchChanged));
        }
    });
    dxState->page.configureButton->SetOnClick([this]() noexcept
    {
        if (! _hostWindow || IsWindow(_hostWindow) == FALSE)
            return;
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::PluginsConfigure));
    });
    dxState->page.testButton->SetOnClick([this]() noexcept
    {
        if (! _hostWindow || IsWindow(_hostWindow) == FALSE)
            return;
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::PluginsTest));
    });
    dxState->page.testAllButton->SetOnClick([this]() noexcept
    {
        if (! _hostWindow || IsWindow(_hostWindow) == FALSE)
            return;
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::PluginsTestAll));
    });
    dxState->page.customPathsAddButton->SetOnClick([this]() noexcept
    {
        if (! _state || ! _hostWindow || IsWindow(_hostWindow) == FALSE)
            return;
        OnCustomPathsAddButtonClick(_hostWindow, *_state);
    });
    dxState->page.customPathsRemoveButton->SetOnClick([this]() noexcept
    {
        if (! _state || ! _hostWindow || IsWindow(_hostWindow) == FALSE)
            return;
        OnCustomPathsRemoveButtonClick(_hostWindow, *_state);
    });

    _dxState    = std::move(dxState);
    _state      = &state;
    _hostWindow = parent;
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    return true;
}

void PluginsPane::DetachDxHosts() noexcept
{
    if (_state)
    {
        _state->pluginsDetailsConfigDxPanel = nullptr;
    }

    if (_pageContentRoot && _pageHostDx && _pageHost && IsWindow(_pageHost) != FALSE)
    {
        _pageHostDx->ResetInteractionState();
        _pageContentRoot->ClearChildren();
    }

    _pageHostDx      = nullptr;
    _pageContentRoot = nullptr;

    if (_dxState)
    {
        _dxState->Detach();
        _dxState.reset();
    }

    _syncingDxInputs    = false;
    _syncingDxSelection = false;
    _state              = nullptr;
    _hostWindow         = nullptr;
}

void PluginsPane::ApplyDxTheme(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return;
    }

    const ThemePalette palette = PrefsUi::MakeDxPalette(state.theme);
    _pageHostDx->SetTheme(palette);
}

void PluginsPane::SyncDxControlsFromState(PreferencesDialogState& state) noexcept
{
    if (! _dxState)
    {
        return;
    }

    _dxState->page.note->SetText(LoadStringResource(nullptr, IDS_PREFS_PLUGINS_NOTE));
    _dxState->page.searchLabel->SetText(LoadStringResource(nullptr, IDS_PREFS_COMMON_SEARCH));
    _dxState->page.searchLabel->SetMnemonicTarget(_dxState->page.searchEdit);
    _dxState->page.detailsIdLabel->SetText(state.pluginsDetailsIdText);
    _dxState->page.detailsConfigError->SetText(state.pluginsDetailsMessageKind == PrefsPluginDetailsMessageKind::EmptyState
                                                   ? state.pluginsDetailsConfigEmptyStateText
                                                   : state.pluginsDetailsConfigErrorText);
    _dxState->page.customPathsHeader->SetText(LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_HEADER));
    _dxState->page.customPathsNote->SetText(LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_NOTE));

    _syncingDxInputs = true;
    auto clearSync   = wil::scope_exit([this]() noexcept { _syncingDxInputs = false; });
    _dxState->page.searchEdit->SetText(state.pluginsSearchText);
    _dxState->page.searchEdit->SetEnabled(true);

    if (_dxState->page.listControl && _dxState->page.listModel)
    {
        std::vector<PluginsGridRow> rows;
        rows.reserve(state.pluginsListItems.size());
        for (const PrefsPluginListItem& item : state.pluginsListItems)
        {
            const std::wstring_view pluginId = GetPluginId(item);
            if (pluginId.empty())
            {
                continue;
            }

            PluginsGridRow row{};
            row.pluginId.assign(pluginId);
            row.stableId    = MakeStableRowId(row.pluginId);
            row.name        = std::wstring(GetPluginDisplayName(item));
            row.typeText    = (item.type == PrefsPluginType::FileSystem) ? LoadStringResource(nullptr, IDS_PREFS_PLUGINS_TYPE_FILE_SYSTEM)
                                                                         : LoadStringResource(nullptr, IDS_PREFS_PLUGINS_TYPE_VIEWER);
            row.originText  = GetPluginOriginText(item);
            row.shortIdText = std::wstring(GetPluginShortIdOrId(item));
            row.enabled     = ! IsPluginDisabledInWorkingSettings(state, pluginId);
            rows.push_back(std::move(row));
        }

        _dxState->page.listModel->SetRows(std::move(rows));
        SyncDxListSelectionFromState(state);
        _dxState->page.listControl->NotifyDataChanged();
    }

    if (_dxState->page.customPathsListControl && _dxState->page.customPathsListModel)
    {
        std::vector<PluginsCustomPathGridRow> rows;
        rows.reserve(state.workingSettings.plugins.customPluginPaths.size());
        for (const std::filesystem::path& path : state.workingSettings.plugins.customPluginPaths)
        {
            const std::wstring pathText = path.wstring();
            if (pathText.empty())
            {
                continue;
            }

            PluginsCustomPathGridRow row{};
            row.path     = pathText;
            row.stableId = MakeStableRowId(row.path);
            rows.push_back(std::move(row));
        }

        _dxState->page.customPathsListModel->SetRows(std::move(rows));
        SyncDxCustomPathsSelectionFromState(state);
        _dxState->page.customPathsListControl->NotifyDataChanged();
    }

    bool hasSelection = false;
    bool loadable     = false;
    if (const std::optional<PrefsPluginListItem> selected = TryGetActivePluginItem(state); selected.has_value())
    {
        const std::wstring_view pluginId = GetPluginId(selected.value());
        hasSelection                     = ! pluginId.empty();
        loadable                         = hasSelection && IsPluginLoadable(selected.value());
    }

    if (_dxState->page.configureButton)
    {
        _dxState->page.configureButton->SetText(LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CONFIGURE_ELLIPSIS));
        _dxState->page.configureButton->SetEnabled(hasSelection);
    }
    if (_dxState->page.testButton)
    {
        _dxState->page.testButton->SetText(LoadStringResource(nullptr, IDS_BTN_TEST));
        _dxState->page.testButton->SetEnabled(loadable);
    }
    if (_dxState->page.testAllButton)
    {
        _dxState->page.testAllButton->SetText(LoadStringResource(nullptr, IDS_BTN_TEST_ALL));
        _dxState->page.testAllButton->SetEnabled(true);
    }
    if (_dxState->page.customPathsAddButton)
    {
        _dxState->page.customPathsAddButton->SetText(LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_ADD_ELLIPSIS));
        _dxState->page.customPathsAddButton->SetEnabled(true);
    }
    if (_dxState->page.customPathsRemoveButton)
    {
        _dxState->page.customPathsRemoveButton->SetText(LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_REMOVE));
        _dxState->page.customPathsRemoveButton->SetEnabled(! state.pluginsSelectedCustomPathText.empty());
    }
    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void PluginsPane::LayoutDxHosts(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _pageHost || ! _pageHostDx || ! _pageContentRoot)
    {
        return;
    }

    const bool showDetails  = state.pluginsDetailsActive && state.pluginsSelectedPlugin.has_value();
    const bool showListMode = ! showDetails;
    _dxState->page.note->SetVisible(showListMode);
    _dxState->page.searchLabel->SetVisible(showListMode);
    _dxState->page.searchEdit->SetVisible(showListMode);
    _dxState->page.listControl->SetVisible(showListMode);
    _dxState->page.configureButton->SetVisible(showListMode);
    _dxState->page.testButton->SetVisible(showListMode);
    _dxState->page.testAllButton->SetVisible(showListMode);
    _dxState->page.customPathsHeader->SetVisible(showListMode);
    _dxState->page.customPathsNote->SetVisible(showListMode);
    _dxState->page.customPathsListControl->SetVisible(showListMode);
    _dxState->page.customPathsAddButton->SetVisible(showListMode);
    _dxState->page.customPathsRemoveButton->SetVisible(showListMode);
    _pageHostDx->Invalidate();
}

#ifdef ENABLE_TESTS
size_t PluginsPane::DebugMainListRowCount() const noexcept
{
    if (! _dxState || ! _dxState->page.listModel)
    {
        return 0u;
    }

    return _dxState->page.listModel->GetRowCount();
}

RedSalamander::DxUi::GridVisibleWorkMetrics PluginsPane::DebugMainListVisibleWorkMetrics() const noexcept
{
    if (! _dxState || ! _dxState->page.listControl)
    {
        return {};
    }

    return _dxState->page.listControl->GetVisibleWorkMetrics();
}

uint64_t PluginsPane::DebugMainListRenderCount() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return 0u;
    }

    return _pageHostDx->DebugGetRenderCount();
}

uint64_t PluginsPane::DebugMainListResizeCount() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return 0u;
    }

    return _pageHostDx->DebugGetResizeCount();
}

uint64_t PluginsPane::DebugMainListResizeFailureCount() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return 0u;
    }

    return _pageHostDx->DebugGetResizeFailureCount();
}

bool PluginsPane::DebugFindToggleableMainListRow(size_t& outRowIndex, bool& outEnabled) const noexcept
{
    if (! _dxState || ! _dxState->page.listModel || ! _state)
    {
        return false;
    }

    const auto& rows = _dxState->page.listModel->GetRows();
    for (size_t i = 0; i < rows.size(); ++i)
    {
        const PluginsGridRow& row                     = rows[i];
        const std::optional<PrefsPluginListItem> item = PrefsPlugins::FindItemById(row.pluginId);
        if (! item.has_value())
        {
            continue;
        }

        const bool disableBlocked =
            row.enabled && item->type == PrefsPluginType::FileSystem && row.pluginId == _state->workingSettings.plugins.currentFileSystemPluginId;
        if (disableBlocked)
        {
            continue;
        }

        outRowIndex = i;
        outEnabled  = row.enabled;
        return true;
    }

    return false;
}

bool PluginsPane::DebugFindLoadableMainListRow(size_t& outRowIndex) const noexcept
{
    if (! _dxState || ! _dxState->page.listModel)
    {
        return false;
    }

    const auto& rows = _dxState->page.listModel->GetRows();
    for (size_t i = 0; i < rows.size(); ++i)
    {
        const std::optional<PrefsPluginListItem> item = PrefsPlugins::FindItemById(rows[i].pluginId);
        if (! item.has_value() || ! IsPluginLoadable(item.value()))
        {
            continue;
        }

        outRowIndex = i;
        return true;
    }

    return false;
}

bool PluginsPane::DebugGetMainListRowEnabled(const size_t rowIndex, bool& outEnabled) const noexcept
{
    if (! _dxState || ! _dxState->page.listModel)
    {
        return false;
    }

    const auto& rows = _dxState->page.listModel->GetRows();
    if (rowIndex >= rows.size())
    {
        return false;
    }

    outEnabled = rows[rowIndex].enabled;
    return true;
}

bool PluginsPane::DebugGetMainListCheckboxClientRect(const size_t rowIndex, RECT& outRect) const noexcept
{
    outRect = RECT{};
    if (! _dxState || ! _dxState->page.listControl || ! _pageHostDx || ! _dxState->page.listModel)
    {
        return false;
    }

    const auto& rows = _dxState->page.listModel->GetRows();
    if (rowIndex >= rows.size())
    {
        return false;
    }

    const std::optional<D2D1_RECT_F> cellRect = _dxState->page.listControl->GetVisibleCellRect(rowIndex, kPluginsColumnName);
    if (! cellRect.has_value())
    {
        return false;
    }

    const float contentLeft        = cellRect->left + 8.0f;
    const float contentTop         = cellRect->top + 3.0f;
    const float contentBottom      = cellRect->bottom - 3.0f;
    const float contentHeight      = std::max(0.0f, contentBottom - contentTop);
    const float indicatorSize      = std::min(16.0f, std::max(12.0f, contentHeight));
    const float indicatorTop       = contentTop + std::max(0.0f, (contentHeight - indicatorSize) * 0.5f);
    const D2D1_RECT_F checkboxRect = D2D1::RectF(contentLeft, indicatorTop, contentLeft + indicatorSize, indicatorTop + indicatorSize);

    outRect.left   = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(checkboxRect.left)));
    outRect.top    = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(checkboxRect.top)));
    outRect.right  = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(checkboxRect.right)));
    outRect.bottom = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(checkboxRect.bottom)));
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

bool PluginsPane::DebugGetMainListHeaderClientRect(const size_t columnIndex, RECT& outRect) const noexcept
{
    outRect = RECT{};
    if (! _dxState || ! _dxState->page.listControl || ! _dxState->page.listModel || ! _pageHostDx || columnIndex >= _dxState->page.listModel->GetColumnCount())
    {
        return false;
    }

    const auto headerRect = _dxState->page.listControl->GetVisibleColumnHeaderRect(columnIndex);
    if (! headerRect.has_value())
    {
        return false;
    }

    outRect.left   = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->left)));
    outRect.top    = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->top)));
    outRect.right  = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->right)));
    outRect.bottom = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->bottom)));
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

size_t PluginsPane::DebugCustomPathsListRowCount() const noexcept
{
    if (! _dxState || ! _dxState->page.customPathsListModel)
    {
        return 0u;
    }

    return _dxState->page.customPathsListModel->GetRowCount();
}

RedSalamander::DxUi::GridVisibleWorkMetrics PluginsPane::DebugCustomPathsListVisibleWorkMetrics() const noexcept
{
    if (! _dxState || ! _dxState->page.customPathsListControl)
    {
        return {};
    }

    return _dxState->page.customPathsListControl->GetVisibleWorkMetrics();
}

uint64_t PluginsPane::DebugCustomPathsListRenderCount() const noexcept
{
    return DebugMainListRenderCount();
}

uint64_t PluginsPane::DebugCustomPathsListResizeCount() const noexcept
{
    return DebugMainListResizeCount();
}

uint64_t PluginsPane::DebugCustomPathsListResizeFailureCount() const noexcept
{
    return DebugMainListResizeFailureCount();
}

bool PluginsPane::DebugGetCustomPathsListHeaderClientRect(const size_t columnIndex, RECT& outRect) const noexcept
{
    outRect = RECT{};
    if (! _dxState || ! _dxState->page.customPathsListControl || ! _dxState->page.customPathsListModel || ! _pageHostDx ||
        columnIndex >= _dxState->page.customPathsListModel->GetColumnCount())
    {
        return false;
    }

    const auto headerRect = _dxState->page.customPathsListControl->GetVisibleColumnHeaderRect(columnIndex);
    if (! headerRect.has_value())
    {
        return false;
    }

    outRect.left   = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->left)));
    outRect.top    = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->top)));
    outRect.right  = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->right)));
    outRect.bottom = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->bottom)));
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

PreferencesPluginsDebugFocusTarget PluginsPane::DebugGetFocusTarget() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return PreferencesPluginsDebugFocusTarget::None;
    }

    RedSalamander::DxUi::Control* const focusedControl = _pageHostDx->GetFocusControl();
    if (! focusedControl)
    {
        return PreferencesPluginsDebugFocusTarget::None;
    }

    if (focusedControl == _dxState->page.searchEdit)
    {
        return PreferencesPluginsDebugFocusTarget::SearchField;
    }
    if (focusedControl == _dxState->page.listControl)
    {
        return PreferencesPluginsDebugFocusTarget::MainList;
    }
    if (focusedControl == _dxState->page.configureButton)
    {
        return PreferencesPluginsDebugFocusTarget::ConfigureButton;
    }
    if (focusedControl == _dxState->page.testButton)
    {
        return PreferencesPluginsDebugFocusTarget::TestButton;
    }
    if (focusedControl == _dxState->page.testAllButton)
    {
        return PreferencesPluginsDebugFocusTarget::TestAllButton;
    }
    if (focusedControl == _dxState->page.customPathsListControl)
    {
        return PreferencesPluginsDebugFocusTarget::CustomPathsList;
    }
    if (focusedControl == _dxState->page.customPathsAddButton)
    {
        return PreferencesPluginsDebugFocusTarget::CustomPathsAddButton;
    }
    if (focusedControl == _dxState->page.customPathsRemoveButton)
    {
        return PreferencesPluginsDebugFocusTarget::CustomPathsRemoveButton;
    }

    return PreferencesPluginsDebugFocusTarget::None;
}

bool PluginsPane::DebugSelectMainListRow(const size_t rowIndex) noexcept
{
    if (! _dxState || ! _dxState->page.listControl || ! _dxState->page.listModel)
    {
        return false;
    }

    const auto& rows = _dxState->page.listModel->GetRows();
    if (rowIndex >= rows.size())
    {
        return false;
    }

    _dxState->page.listControl->GetSelectionModel().SetSingle(rows[rowIndex].stableId);
    OnGridSelectionChanged(*_dxState->page.listControl);
    return true;
}

bool PluginsPane::DebugClickMainListRow(const size_t rowIndex) noexcept
{
    if (! _dxState || ! _dxState->page.listControl || ! _dxState->page.listModel)
    {
        return false;
    }

    const auto& rows = _dxState->page.listModel->GetRows();
    if (rowIndex >= rows.size())
    {
        return false;
    }

    _dxState->page.listControl->GetSelectionModel().SetSingle(rows[rowIndex].stableId);
    OnGridSelectionChanged(*_dxState->page.listControl);
    return true;
}

bool PluginsPane::DebugToggleMainListCheckbox(const size_t rowIndex) noexcept
{
    if (! _dxState || ! _dxState->page.listControl || ! _pageHostDx || ! _dxState->page.listModel)
    {
        return false;
    }

    return _dxState->page.listControl->RequestToggleCheckboxCell(*_pageHostDx, rowIndex, kPluginsColumnName);
}

bool PluginsPane::DebugSelectCustomPathsListRow(const size_t rowIndex) noexcept
{
    if (! _dxState || ! _dxState->page.customPathsListControl || ! _dxState->page.customPathsListModel)
    {
        return false;
    }

    const auto& rows = _dxState->page.customPathsListModel->GetRows();
    if (rowIndex >= rows.size())
    {
        return false;
    }

    _dxState->page.customPathsListControl->GetSelectionModel().SetSingle(rows[rowIndex].stableId);
    OnGridSelectionChanged(*_dxState->page.customPathsListControl);
    return true;
}

bool PluginsPane::DebugClearCustomPaths() noexcept
{
    PreferencesDialogState* state = _state;
    if (! state && _hostWindow)
    {
        state  = PrefsUi::GetDialogState(_hostWindow);
        _state = state;
    }

    if (! state)
    {
        return false;
    }

    state->workingSettings.plugins.customPluginPaths.clear();
    state->pluginsSelectedCustomPathText.clear();

    if (_hostWindow && IsWindow(_hostWindow) != FALSE)
    {
        Refresh(_hostWindow, *state);
        return true;
    }

    SyncDxControlsFromState(*state);
    return true;
}

bool PluginsPane::DebugSetSearchText(std::wstring_view text) noexcept
{
    if (! _state)
    {
        return false;
    }

    _state->pluginsSearchText.assign(text);
    if (_dxState && _dxState->page.searchEdit)
    {
        _dxState->page.searchEdit->SetText(std::wstring(text));
        if (_hostWindow && IsWindow(_hostWindow) != FALSE)
        {
            Refresh(_hostWindow, *_state);
        }
        if (_pageHostDx)
        {
            _pageHostDx->Invalidate();
        }
        return true;
    }

    return false;
}

bool PluginsPane::DebugFocusMainList() noexcept
{
    if (! _dxState || ! _dxState->page.listControl || ! _pageHostDx)
    {
        return false;
    }

    _pageHostDx->SetFocusControl(_dxState->page.listControl);
    return true;
}

bool PluginsPane::DebugFocusSearchField() noexcept
{
    if (! _dxState || ! _dxState->page.searchEdit || ! _pageHostDx)
    {
        return false;
    }

    _pageHostDx->SetFocusControl(_dxState->page.searchEdit);
    return true;
}

bool PluginsPane::DebugScrollMainListByWheelDetents(const int detents) noexcept
{
    if (! _dxState || ! _dxState->page.listControl || ! _pageHostDx)
    {
        return false;
    }

    _pageHostDx->SetFocusControl(_dxState->page.listControl);
    const float wheelDelta = detents >= 0 ? static_cast<float>(WHEEL_DELTA) : -static_cast<float>(WHEEL_DELTA);
    const int steps        = std::abs(detents);
    for (int i = 0; i < steps; ++i)
    {
        static_cast<void>(_dxState->page.listControl->OnMouseWheel(*_pageHostDx, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0u));
    }
    return true;
}

bool PluginsPane::DebugScrollCustomPathsListByWheelDetents(const int detents) noexcept
{
    if (! _dxState || ! _dxState->page.customPathsListControl || ! _pageHostDx)
    {
        return false;
    }

    _pageHostDx->SetFocusControl(_dxState->page.customPathsListControl);
    const float wheelDelta = detents >= 0 ? static_cast<float>(WHEEL_DELTA) : -static_cast<float>(WHEEL_DELTA);
    const int steps        = std::abs(detents);
    for (int i = 0; i < steps; ++i)
    {
        static_cast<void>(_dxState->page.customPathsListControl->OnMouseWheel(*_pageHostDx, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0u));
    }
    return true;
}
#endif

void PluginsPane::OnGridSelectionChanged(Grid& sender)
{
    PreferencesDialogState* state = _state;
    if (! state && _hostWindow)
    {
        state  = PrefsUi::GetDialogState(_hostWindow);
        _state = state;
    }

    if (! state || ! _dxState || _syncingDxSelection)
    {
        return;
    }

    if (&sender == _dxState->page.listControl && _dxState->page.listModel)
    {
        const auto selectedRowIds = _dxState->page.listControl->GetSelectionModel().GetOrderedSelection();
        std::wstring selectedPluginId;
        if (! selectedRowIds.empty())
        {
            const auto& rows    = _dxState->page.listModel->GetRows();
            const auto rowIndex = _dxState->page.listModel->FindRowByStableId(selectedRowIds.front());
            if (rowIndex.has_value() && rowIndex.value() < rows.size())
            {
                selectedPluginId = rows[rowIndex.value()].pluginId;
            }
        }

        if (! selectedPluginId.empty())
        {
            if (const std::optional<PrefsPluginListItem> selected = PrefsPlugins::FindItemById(selectedPluginId); selected.has_value())
            {
                state->pluginsSelectedPlugin = selected;
                state->pluginsSelectedPluginId.assign(selectedPluginId);
                state->pluginsRetainedSelectedPluginId.assign(selectedPluginId);
            }
            else
            {
                state->pluginsSelectedPlugin.reset();
                state->pluginsSelectedPluginId.clear();
                state->pluginsRetainedSelectedPluginId.clear();
                state->pluginsDetailsActive = false;
            }
        }
        else
        {
            state->pluginsSelectedPlugin.reset();
            state->pluginsSelectedPluginId.clear();
            state->pluginsDetailsActive = false;
        }
        UpdatePluginsActionButtonsEnabled(*state);
        if (_hostWindow && IsWindow(_hostWindow) != FALSE)
        {
            SyncDxControlsFromState(*state);
        }
        else
        {
            SyncDxControlsFromState(*state);
        }
        return;
    }

    if (&sender == _dxState->page.customPathsListControl && _dxState->page.customPathsListModel)
    {
        const auto selectedRowIds = _dxState->page.customPathsListControl->GetSelectionModel().GetOrderedSelection();
        std::wstring selectedPath;
        if (! selectedRowIds.empty())
        {
            const auto& rows    = _dxState->page.customPathsListModel->GetRows();
            const auto rowIndex = _dxState->page.customPathsListModel->FindRowByStableId(selectedRowIds.front());
            if (rowIndex.has_value() && rowIndex.value() < rows.size())
            {
                selectedPath = rows[rowIndex.value()].path;
            }
        }

        state->pluginsSelectedCustomPathText = selectedPath;
        if (_dxState->page.customPathsRemoveButton)
        {
            _dxState->page.customPathsRemoveButton->SetEnabled(! selectedPath.empty());
        }
        if (_pageHostDx)
        {
            _pageHostDx->Invalidate();
        }
    }
}

void PluginsPane::OnGridCheckboxToggled(Grid& sender, size_t rowIndex, size_t /*columnIndex*/, bool checked)
{
    PreferencesDialogState* state = _state;
    if (! state && _hostWindow)
    {
        state  = PrefsUi::GetDialogState(_hostWindow);
        _state = state;
    }

    if (! _dxState || &sender != _dxState->page.listControl || ! state || ! _dxState->page.listModel)
    {
        return;
    }

    const auto& rows = _dxState->page.listModel->GetRows();
    if (rowIndex >= rows.size())
    {
        return;
    }

    const PluginsGridRow& row                         = rows[rowIndex];
    const std::optional<PrefsPluginListItem> selected = PrefsPlugins::FindItemById(row.pluginId);
    if (! selected.has_value())
    {
        return;
    }

    HWND dlg = _hostWindow ? GetParent(_hostWindow) : nullptr;
    if (! dlg && state->categoryTreeWindow)
    {
        dlg = GetParent(state->categoryTreeWindow);
    }

    static_cast<void>(TryApplyPluginEnabledStateChange(dlg, *state, selected.value(), checked));
    UpdatePluginsActionButtonsEnabled(*state);
    SyncDxControlsFromState(*state);
}

void PluginsPane::OnVisibilityChanged(bool visible) noexcept
{
    if (! visible && _pageHostDx)
    {
        _pageHostDx->ResetInteractionState();
    }
}

void PluginsPane::OnConfigureButtonClick(HWND host, PreferencesDialogState& state) noexcept
{
    const std::optional<PrefsPluginListItem> selected = state.pluginsSelectedPlugin.has_value() ? state.pluginsSelectedPlugin : TryGetActivePluginItem(state);
    if (! selected.has_value())
    {
        return;
    }

    const std::wstring_view pluginId = GetPluginId(selected.value());
    if (pluginId.empty())
    {
        return;
    }

    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    ClearPluginsStatusMessage(state);
    state.initialCategory       = PrefCategory::Plugins;
    state.currentCategory       = PrefCategory::Plugins;
    state.pluginsSelectedPlugin = selected;
    state.pluginsSelectedPluginId.assign(pluginId);
    state.pluginsRetainedSelectedPluginId.assign(pluginId);
    state.pluginsDetailsActive = true;
    PluginsPane::Refresh(host, state);
    SendMessageW(dlg, WndMsg::kPreferencesSelectPluginDetails, static_cast<WPARAM>(selected->type), static_cast<LPARAM>(selected->index));
}

void PluginsPane::OnTestButtonClick(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    const std::optional<PrefsPluginListItem> selected = TryGetActivePluginItem(state);
    if (! selected.has_value())
    {
        return;
    }

    const std::wstring_view pluginId = GetPluginId(selected.value());
    if (pluginId.empty() || ! IsPluginLoadable(selected.value()))
    {
        return;
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

    const UINT textId                = SUCCEEDED(hr) ? IDS_MSG_PLUGIN_TEST_OK : IDS_FMT_PLUGIN_TEST_FAILED_NAMED;
    const HostAlertSeverity severity = SUCCEEDED(hr) ? HOST_ALERT_INFO : HOST_ALERT_ERROR;
    const std::wstring pluginName    = std::wstring(GetPluginDisplayName(selected.value()));
    const std::wstring message       = SUCCEEDED(hr) ? LoadStringResource(nullptr, textId) : FormatStringResource(nullptr, textId, pluginName);
    SetPluginsStatusMessage(state, severity, LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER), message);
    PluginsPane::Refresh(host, state);
    FlushPluginsFeedbackUi(host);
    ShowDialogAlert(dlg, severity, LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER), message);
}

void PluginsPane::OnTestAllButtonClick(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    size_t okCount   = 0;
    size_t failCount = 0;
    std::vector<std::wstring> failedPluginNames;

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
                failedPluginNames.push_back(entry.name.empty() ? entry.id : entry.name);
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
                failedPluginNames.push_back(entry.name.empty() ? entry.id : entry.name);
            }
        }
    }

    const std::wstring failedNames   = JoinPluginNames(failedPluginNames);
    const std::wstring resultText    = failedNames.empty()
                                           ? FormatStringResource(nullptr, IDS_FMT_PLUGIN_TEST_ALL_RESULT, okCount, failCount)
                                           : FormatStringResource(nullptr, IDS_FMT_PLUGIN_TEST_ALL_RESULT_WITH_FAILURES, okCount, failCount, failedNames);
    const HostAlertSeverity severity = failCount == 0 ? HOST_ALERT_INFO : HOST_ALERT_WARNING;
    SetPluginsStatusMessage(state, severity, LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER), resultText);
    PluginsPane::Refresh(host, state);
    FlushPluginsFeedbackUi(host);
    ShowDialogAlert(dlg, severity, LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER), resultText);
}

void PluginsPane::OnCustomPathsAddButtonClick(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    std::filesystem::path selectedPath;
    if (! TryBrowseCustomPluginPath(dlg, selectedPath))
    {
        return;
    }

    if (! IsDllPath(selectedPath))
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_INVALID));
        return;
    }

    auto& customPaths = state.workingSettings.plugins.customPluginPaths;
    if (std::find(customPaths.begin(), customPaths.end(), selectedPath) == customPaths.end())
    {
        customPaths.push_back(selectedPath);
    }

    state.pluginsSelectedCustomPathText = selectedPath.wstring();
    PluginsPane::Refresh(host, state);
    SetDirty(dlg, state);
}

void PluginsPane::OnCustomPathsRemoveButtonClick(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    auto& customPaths = state.workingSettings.plugins.customPluginPaths;
    std::optional<size_t> pathIndex;
    if (! state.pluginsSelectedCustomPathText.empty())
    {
        for (size_t i = 0; i < customPaths.size(); ++i)
        {
            if (OrdinalString::EqualsNoCase(customPaths[i].native(), state.pluginsSelectedCustomPathText))
            {
                pathIndex = i;
                break;
            }
        }
    }
    if (! pathIndex.has_value())
    {
        return;
    }

    customPaths.erase(customPaths.begin() + static_cast<std::ptrdiff_t>(pathIndex.value()));
    state.pluginsSelectedCustomPathText.clear();
    PluginsPane::Refresh(host, state);
    SetDirty(dlg, state);
}

bool PluginsPane::HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept
{
    _hostWindow = host;
    _state      = &state;

    if (! host || ! _dxState)
    {
        return false;
    }

#pragma warning(push)
#pragma warning(disable : 4061) // Not all enum values handled explicitly -- intentional; this pane only handles its own actions.
    switch (action)
    {
        case PreferencesDeferredActionKind::PluginsSearchChanged:
            // DxUi callbacks already updated the search text; preserve it and just rebuild the filtered view.
            PluginsPane::Refresh(host, state);
            return true;

        case PreferencesDeferredActionKind::PluginsConfigure: OnConfigureButtonClick(host, state); return true;

        case PreferencesDeferredActionKind::PluginsTest: OnTestButtonClick(host, state); return true;

        case PreferencesDeferredActionKind::PluginsTestAll: OnTestAllButtonClick(host, state); return true;
        default: return false;
    }
#pragma warning(pop)
}

void PluginsPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    if (! _dxState)
    {
        if (! EnsureDxHosts(_pageHost ? _pageHost : host, state))
        {
            Debug::Error(L"Preferences.Plugins: Failed to ensure DxUi hosts during Refresh.");
        }
    }
    if (_dxState)
    {
        ApplyDxTheme(state);
    }

    RefreshSelectedPluginCache(state);
    if (state.pluginsDetailsActive && state.pluginsSelectedPlugin.has_value())
    {
        state.refreshingPluginsPage = true;
        auto clearRefreshFlag       = wil::scope_exit([&]() noexcept { state.refreshingPluginsPage = false; });

        const PrefsPluginListItem pluginItem = state.pluginsSelectedPlugin.value();

        const HWND parent                   = _pageHost ? _pageHost : host;
        const std::wstring previousEditorId = state.pluginsDetailsConfigPluginId;
        const bool hadEditor                = ! state.pluginsDetailsConfigFields.empty();
        static_cast<void>(PrefsPluginConfiguration::EnsureEditor(parent, state, pluginItem));
        const bool hasEditor = ! state.pluginsDetailsConfigFields.empty();

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
                SyncDxControlsFromState(state);
            }
        }

        UpdatePluginsActionButtonsEnabled(state);
        SyncDxControlsFromState(state);
        return;
    }

    PrefsPluginConfiguration::Clear(state);
    PrefsPluginConfiguration::SetDetailsIdText(state, L"");

    state.refreshingPluginsPage = true;
    auto clearRefreshFlag       = wil::scope_exit([&]() noexcept { state.refreshingPluginsPage = false; });

    const std::wstring_view filter = PrefsUi::TrimWhitespace(state.pluginsSearchText);

    std::wstring selectedPluginId = state.pluginsSelectedPluginId;
    if (selectedPluginId.empty())
    {
        // Preserve the user's last stable row across filtered rebuilds, even when the
        // details pane is not active. Search round-trips temporarily clear the visible
        // selection when the chosen row is filtered out, then should restore it.
        selectedPluginId = state.pluginsRetainedSelectedPluginId;
    }

    state.pluginsListItems.clear();
    PrefsPlugins::BuildListItems(state.pluginsListItems);
    if (! filter.empty())
    {
        std::erase_if(state.pluginsListItems,
                      [filter](const PrefsPluginListItem& item) noexcept
        {
            const std::wstring_view pluginId    = GetPluginId(item);
            const std::wstring_view displayName = GetPluginDisplayName(item);
            const std::wstring_view shortId     = GetPluginShortIdOrId(item);
            const std::wstring originText       = GetPluginOriginText(item);
            const std::wstring typeText         = item.type == PrefsPluginType::FileSystem ? LoadStringResource(nullptr, IDS_PREFS_PLUGINS_TYPE_FILE_SYSTEM)
                                                                                           : LoadStringResource(nullptr, IDS_PREFS_PLUGINS_TYPE_VIEWER);
            return ! (PrefsUi::ContainsCaseInsensitive(pluginId, filter) || PrefsUi::ContainsCaseInsensitive(displayName, filter) ||
                      PrefsUi::ContainsCaseInsensitive(shortId, filter) || PrefsUi::ContainsCaseInsensitive(originText, filter) ||
                      PrefsUi::ContainsCaseInsensitive(typeText, filter));
        });
    }

    if (! selectedPluginId.empty() &&
        std::none_of(state.pluginsListItems.begin(), state.pluginsListItems.end(), [selectedPluginId](const PrefsPluginListItem& item) noexcept {
        return GetPluginId(item) == selectedPluginId;
    }))
    {
        state.pluginsSelectedPlugin.reset();
        state.pluginsSelectedPluginId.clear();
    }
    else if (! selectedPluginId.empty())
    {
        if (const std::optional<PrefsPluginListItem> selected = PrefsPlugins::FindItemById(selectedPluginId); selected.has_value())
        {
            state.pluginsSelectedPlugin = selected;
            state.pluginsSelectedPluginId.assign(selectedPluginId);
            state.pluginsRetainedSelectedPluginId.assign(selectedPluginId);
        }
        else
        {
            state.pluginsSelectedPlugin.reset();
            state.pluginsSelectedPluginId.clear();
            state.pluginsRetainedSelectedPluginId.clear();
        }
    }
    else if (! state.pluginsDetailsActive)
    {
        state.pluginsSelectedPlugin.reset();
        state.pluginsSelectedPluginId.clear();
    }

    UpdatePluginsActionButtonsEnabled(state);
    SyncDxControlsFromState(state);
}

void PluginsPane::LayoutDxPage(HWND host,
                               PreferencesDialogState& state,
                               int x,
                               int& y,
                               int width,
                               int margin,
                               int gapY,
                               int sectionY,
                               const PreferencesTypographyContext& typography) noexcept
{
    using namespace PrefsLayoutConstants;

    Debug::Perf::Scope layoutPerf(L"preferences.ui.plugins_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(static_cast<uint64_t>(typography.dpi));

    RECT hostClient{};
    GetClientRect(host, &hostClient);
    const int hostBottom        = std::max(0l, hostClient.bottom - hostClient.top);
    const int hostContentBottom = std::max(0, hostBottom - margin);

    const UINT dpi        = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);
    const auto pxToDip    = [dpi](const int pixels) noexcept { return (static_cast<float>(pixels) * 96.0f) / static_cast<float>(std::max<UINT>(1u, dpi)); };
    const int rowHeight   = std::max(1, UiMetrics::ScaleDip(dpi, kRowHeightDip));
    const int labelHeight = std::max(1, UiMetrics::ScaleDip(dpi, kTitleHeightDip));
    const int gapX        = UiMetrics::ScaleDip(dpi, kToggleGapXDip);

    const int buttonHeight = rowHeight;
    const int buttonPadX   = UiMetrics::ScaleDip(dpi, kCardPaddingXDip);

    const auto measureButtonWidth = [&](HWND button, std::wstring_view fallbackText, int minWidthDip) noexcept
    {
        std::wstring text(fallbackText);
        if (button)
        {
            text = PrefsUi::GetWindowTextString(button);
        }
        const int textW = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.body, text);
        return std::max(UiMetrics::ScaleDip(dpi, minWidthDip), textW + 2 * buttonPadX);
    };

    RefreshSelectedPluginCache(state);
    const bool showDetails = state.pluginsDetailsActive && state.pluginsSelectedPlugin.has_value();
    if (showDetails)
    {
        const PrefsPluginListItem pluginItem     = state.pluginsSelectedPlugin.value();
        const std::wstring_view selectedPluginId = GetPluginId(pluginItem);

        if (! selectedPluginId.empty() && ! state.pluginsDetailsConfigPluginId.empty() && state.pluginsDetailsConfigPluginId != selectedPluginId)
        {
            PrefsPluginConfiguration::Clear(state);
        }

        if (! selectedPluginId.empty())
        {
            PrefsPluginConfiguration::SetDetailsIdText(state, FormatStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_ID_FMT, std::wstring(selectedPluginId)));
        }

        const bool hasEditor =
            ! selectedPluginId.empty() && state.pluginsDetailsConfigPluginId == selectedPluginId && ! state.pluginsDetailsConfigFields.empty();

        const bool showConfigError     = ! state.pluginsDetailsConfigErrorText.empty() || ! state.pluginsDetailsConfigEmptyStateText.empty();
        const bool useDxDetailsStatics = _dxState != nullptr;
        const int cardSpacingY         = UiMetrics::ScaleDip(dpi, kCardSpacingYDip);

        const auto pushCard = [&](const RECT& card) noexcept { state.pageSettingCards.push_back(card); };

        if (useDxDetailsStatics && _dxState->page.detailsIdLabel)
        {
            const std::wstring& idText = state.pluginsDetailsIdText;
            const int measuredHeight   = idText.empty() ? 0 : PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, width, idText);
            const int idHeight         = std::max(labelHeight, std::max(0, measuredHeight));

            _dxState->page.detailsIdLabel->SetText(idText);
            _dxState->page.detailsIdLabel->SetVisible(! idText.empty());
            _dxState->page.detailsIdLabel->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + idHeight)));
            y += idHeight + sectionY;
        }
        else
        {
            y += sectionY;
        }

        if (showConfigError && ! hasEditor)
        {
            if (useDxDetailsStatics && _dxState->page.detailsConfigError)
            {
                _dxState->page.detailsConfigError->SetVisible(false);
                _dxState->page.detailsConfigError->SetBounds(D2D1::RectF());
                _dxState->page.detailsConfigError->SetText(L"");
            }

            PreferencesEmptyStateSpec spec{};
            if (state.pluginsDetailsMessageKind == PrefsPluginDetailsMessageKind::EmptyState)
            {
                spec = GetPluginNoSettingsEmptyState(pluginItem);
            }
            else
            {
                spec = GetPluginMessageState(
                    PrefsInlineMessageSeverity::Warning, LoadStringResource(nullptr, IDS_CAPTION_WARNING), state.pluginsDetailsConfigErrorText);
            }

            const int cardHeight = PrefsUi::ShowSharedPageEmptyState(host, state, spec, x, y, width, typography);
            if (cardHeight > 0)
            {
                RECT card{x, y, x + width, y + cardHeight};
                pushCard(card);
                y += cardHeight + cardSpacingY;
            }

            LayoutDxHosts(state);
            return;
        }

        if (hasEditor)
        {
            PrefsPluginConfiguration::LayoutCards(host, state, x, y, width, typography);
            LayoutDxHosts(state);
            return;
        }

        LayoutDxHosts(state);
        return;
    }

    PrefsPluginConfiguration::Clear(state);

    const std::wstring noteText = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_NOTE);
    const int noteHeight        = noteText.empty() ? 0 : PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, width, noteText);
    if (_dxState)
    {
        _dxState->page.detailsIdLabel->SetVisible(false);
        _dxState->page.detailsConfigError->SetVisible(false);
        _dxState->page.note->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + std::max(0, noteHeight))));
    }
    y += std::max(0, noteHeight) + sectionY;

    if (! state.pluginsStatusBodyText.empty())
    {
        const int statusCardHeight =
            PrefsUi::ShowSharedPageEmptyState(host,
                                              state,
                                              GetPluginMessageState(state.pluginsStatusSeverity, state.pluginsStatusTitleText, state.pluginsStatusBodyText),
                                              x,
                                              y,
                                              width,
                                              typography);
        if (statusCardHeight > 0)
        {
            state.pageSettingCards.push_back(RECT{x, y, x + width, y + statusCardHeight});
            y += statusCardHeight + UiMetrics::ScaleDip(dpi, kCardSpacingYDip);
        }
    }

    const int searchLabelWidth = std::min(width, UiMetrics::ScaleDip(dpi, 52));
    const int searchEditWidth  = std::max(0, width - searchLabelWidth - gapX);
    const int searchEditX      = x + searchLabelWidth + gapX;
    if (_dxState)
    {
        _dxState->page.searchLabel->SetBounds(D2D1::RectF(
            pxToDip(x), pxToDip(y + (rowHeight - labelHeight) / 2), pxToDip(x + searchLabelWidth), pxToDip(y + (rowHeight - labelHeight) / 2 + labelHeight)));
        _dxState->page.searchEdit->SetBounds(D2D1::RectF(pxToDip(searchEditX), pxToDip(y), pxToDip(searchEditX + searchEditWidth), pxToDip(y + rowHeight)));
    }
    y += rowHeight + gapY;

    const int configureButtonWidth = std::min(width, measureButtonWidth(nullptr, LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CONFIGURE_ELLIPSIS), 90));
    const int testButtonWidth      = std::min(width, measureButtonWidth(nullptr, LoadStringResource(nullptr, IDS_BTN_TEST), 70));
    const int testAllButtonWidth   = std::min(width, measureButtonWidth(nullptr, LoadStringResource(nullptr, IDS_BTN_TEST_ALL), 90));

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

    int customAddWidth    = std::min(width, measureButtonWidth(nullptr, LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_ADD_ELLIPSIS), 70));
    int customRemoveWidth = std::min(width, measureButtonWidth(nullptr, LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_REMOVE), 70));
    if (customAddWidth > 0 && customRemoveWidth > 0 && (customAddWidth + gapX + customRemoveWidth > width))
    {
        customRemoveWidth = std::max(0, width - customAddWidth - gapX);
    }

    const std::wstring customNoteText = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_NOTE);
    const int customNoteHeight        = customNoteText.empty() ? 0 : PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, width, customNoteText);

    int customBlockHeight = labelHeight + gapY;
    if (customNoteHeight > 0)
    {
        customBlockHeight += customNoteHeight + gapY;
    }
    const int customListHeight = std::max(1, UiMetrics::ScaleDip(dpi, 90));
    customBlockHeight += customListHeight + gapY;

    const int pinnedCustomBtnsTop  = hostContentBottom - buttonHeight;
    const int minPluginsListHeight = std::max(1, UiMetrics::ScaleDip(dpi, 120));

    const int pinnedPluginsHeight = pinnedCustomBtnsTop - y - customBlockHeight - actionsBlockHeight;
    const bool pinnedLayout       = pinnedCustomBtnsTop >= y && pinnedPluginsHeight >= minPluginsListHeight;

    const int pluginsListTop         = y;
    const int reservedForActions     = actionsBlockHeight;
    const int preferredPluginsHeight = std::max(0, hostContentBottom - y - reservedForActions);
    const int pluginsListHeight      = pinnedLayout ? pinnedPluginsHeight : std::max(minPluginsListHeight, preferredPluginsHeight);

    if (_dxState)
    {
        _dxState->page.listControl->SetBounds(
            D2D1::RectF(pxToDip(x), pxToDip(pluginsListTop), pxToDip(x + width), pxToDip(pluginsListTop + pluginsListHeight)));
    }

    y += pluginsListHeight;

    y += gapY;
    if (buttonsSingleRow)
    {
        int currentX = x;
        if (_dxState)
        {
            _dxState->page.configureButton->SetBounds(
                D2D1::RectF(pxToDip(currentX), pxToDip(y), pxToDip(currentX + configureButtonWidth), pxToDip(y + buttonHeight)));
            currentX += configureButtonWidth + gapX;
        }
        if (_dxState && testButtonWidth > 0)
        {
            _dxState->page.testButton->SetBounds(D2D1::RectF(pxToDip(currentX), pxToDip(y), pxToDip(currentX + testButtonWidth), pxToDip(y + buttonHeight)));
            currentX += testButtonWidth + gapX;
        }
        if (_dxState && testAllButtonWidth > 0)
        {
            _dxState->page.testAllButton->SetBounds(
                D2D1::RectF(pxToDip(currentX), pxToDip(y), pxToDip(currentX + testAllButtonWidth), pxToDip(y + buttonHeight)));
        }

        y += buttonHeight + sectionY;
    }
    else
    {
        int currentY = y;
        if (_dxState)
        {
            if (configureButtonWidth > 0)
            {
                _dxState->page.configureButton->SetBounds(
                    D2D1::RectF(pxToDip(x), pxToDip(currentY), pxToDip(x + configureButtonWidth), pxToDip(currentY + buttonHeight)));
                currentY += buttonHeight + gapY;
            }
            if (testButtonWidth > 0)
            {
                _dxState->page.testButton->SetBounds(
                    D2D1::RectF(pxToDip(x), pxToDip(currentY), pxToDip(x + testButtonWidth), pxToDip(currentY + buttonHeight)));
                currentY += buttonHeight + gapY;
            }
            if (testAllButtonWidth > 0)
            {
                _dxState->page.testAllButton->SetBounds(
                    D2D1::RectF(pxToDip(x), pxToDip(currentY), pxToDip(x + testAllButtonWidth), pxToDip(currentY + buttonHeight)));
            }
        }
        if (configureButtonWidth > 0 || testButtonWidth > 0 || testAllButtonWidth > 0)
        {
            currentY -= gapY;
        }
        y = currentY + sectionY;
    }

    if (_dxState)
    {
        _dxState->page.customPathsHeader->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + labelHeight)));
    }

    y += labelHeight + gapY;

    if (_dxState)
    {
        _dxState->page.customPathsNote->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + std::max(0, customNoteHeight))));
    }
    y += std::max(0, customNoteHeight);
    if (customNoteHeight > 0)
    {
        y += gapY;
    }

    if (_dxState)
    {
        _dxState->page.customPathsListControl->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + customListHeight)));
    }
    y += customListHeight + gapY;

    const int customButtonsTop = pinnedLayout ? pinnedCustomBtnsTop : y;
    if (_dxState && customAddWidth > 0)
    {
        _dxState->page.customPathsAddButton->SetBounds(
            D2D1::RectF(pxToDip(x), pxToDip(customButtonsTop), pxToDip(x + customAddWidth), pxToDip(customButtonsTop + buttonHeight)));
    }
    if (_dxState && customRemoveWidth > 0)
    {
        const int removeX = x + customAddWidth + gapX;
        _dxState->page.customPathsRemoveButton->SetBounds(
            D2D1::RectF(pxToDip(removeX), pxToDip(customButtonsTop), pxToDip(removeX + customRemoveWidth), pxToDip(customButtonsTop + buttonHeight)));
    }

    y = customButtonsTop + buttonHeight;

    LayoutDxHosts(state);
}

void PluginsPane::LayoutPage(HWND host,
                             PreferencesDialogState& state,
                             int x,
                             int& y,
                             int width,
                             int margin,
                             int gapY,
                             int sectionY,
                             const PreferencesTypographyContext& typography) noexcept
{
    if (! host)
    {
        return;
    }

    if (EnsureDxHosts(_pageHost ? _pageHost : host, state))
    {
        LayoutDxPage(host, state, x, y, width, margin, gapY, sectionY, typography);
        return;
    }

    Debug::Error(L"Preferences.Plugins: DxUi surface initialization failed; page will not render correctly.");
}

void PluginsPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageHost = parent;

    if (state.currentCategory == PrefCategory::Plugins)
    {
        if (! EnsureDxHosts(parent, state))
        {
            Debug::Error(L"Preferences.Plugins: Failed to initialize DxUi hosts in CreateControls.");
            DetachDxHosts();
            return;
        }
    }

    if (state.currentCategory != PrefCategory::Plugins)
    {
        return;
    }
}

#ifdef ENABLE_TESTS
bool DebugSetPreferencesPluginsNextCustomPathBrowsePath(const std::wstring_view path) noexcept
{
    return DebugSetPreferencesPluginsNextCustomPathBrowsePathImpl(path);
}

bool DebugCancelPreferencesPluginsNextCustomPathBrowse() noexcept
{
    return DebugCancelPreferencesPluginsNextCustomPathBrowseImpl();
}
#endif
