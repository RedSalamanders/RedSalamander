#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "RedConfigureRoot.h"

#include "RedConfigureApp.h"
#include "RedConfigureSession.h"
#include "SettingsStore.h"
#include "resource.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>
#include <filesystem>
#include <functional>
#include <iterator>
#include <limits>
#include <memory>
#include <optional>
#include <set>
#include <span>
#include <string>
#include <string_view>
#include <system_error>
#include <vector>

namespace RedConfigure::Ui
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ColorSwatch;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::Control;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridColumnKind;
using RedSalamander::DxUi::GridRowStyle;
using RedSalamander::DxUi::GridRowTone;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::GridSortSpec;
using RedSalamander::DxUi::IDxGridDelegate;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::MakeDefaultThemePalette;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::ScrollPanel;
using RedSalamander::DxUi::SortDirection;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::WindowHost;

constexpr float kThemePreviewContentHeightDip = 560.0f;
constexpr std::array<std::wstring_view, 26> kThemePreviewColorKeys = {{
    L"app.accent",
    L"window.background",
    L"window.text",
    L"window.subduedText",
    L"navigation.background",
    L"navigation.text",
    L"navigation.accent",
    L"menu.background",
    L"menu.text",
    L"menu.selectionBackground",
    L"menu.selectionText",
    L"menu.border",
    L"folderView.background",
    L"folderView.itemForeground",
    L"folderView.itemBackgroundHovered",
    L"folderView.itemBackgroundSelected",
    L"folderView.itemForegroundSelected",
    L"folderView.warningForeground",
    L"dialog.background",
    L"dialog.text",
    L"dialog.buttonBackground",
    L"dialog.buttonText",
    L"progress.background",
    L"progress.fill",
    L"diff.addedBackground",
    L"diff.removedBackground",
}};

[[nodiscard]] std::wstring LoadAppString(HINSTANCE instance, UINT resourceId)
{
    wchar_t buffer[1024]{};
    const int length = ::LoadStringW(instance, resourceId, buffer, static_cast<int>(std::size(buffer)));
    if (length <= 0)
    {
        return {};
    }

    return std::wstring(buffer, static_cast<size_t>(length));
}

[[nodiscard]] D2D1_COLOR_F ColorFromArgb(uint32_t argb) noexcept
{
    const float a = static_cast<float>((argb >> 24u) & 0xFFu) / 255.0f;
    const float r = static_cast<float>((argb >> 16u) & 0xFFu) / 255.0f;
    const float g = static_cast<float>((argb >> 8u) & 0xFFu) / 255.0f;
    const float b = static_cast<float>(argb & 0xFFu) / 255.0f;
    return D2D1::ColorF(r, g, b, a);
}

[[nodiscard]] uint32_t Channel(uint32_t argb, uint32_t shift) noexcept
{
    return (argb >> shift) & 0xFFu;
}

[[nodiscard]] uint32_t PackArgb(uint32_t a, uint32_t r, uint32_t g, uint32_t b) noexcept
{
    return ((a & 0xFFu) << 24u) | ((r & 0xFFu) << 16u) | ((g & 0xFFu) << 8u) | (b & 0xFFu);
}

[[nodiscard]] uint32_t MixChannel(uint32_t from, uint32_t to, double amount) noexcept
{
    const double mixed = static_cast<double>(from) + ((static_cast<double>(to) - static_cast<double>(from)) * amount);
    return static_cast<uint32_t>(std::clamp(mixed, 0.0, 255.0) + 0.5);
}

[[nodiscard]] uint32_t MixArgb(uint32_t from, uint32_t to, double amount) noexcept
{
    return PackArgb(MixChannel(Channel(from, 24u), Channel(to, 24u), amount),
                    MixChannel(Channel(from, 16u), Channel(to, 16u), amount),
                    MixChannel(Channel(from, 8u), Channel(to, 8u), amount),
                    MixChannel(Channel(from, 0u), Channel(to, 0u), amount));
}

[[nodiscard]] std::wstring ThemeKeyGroup(std::wstring_view key)
{
    const size_t dot = key.find(L'.');
    return std::wstring(dot == std::wstring_view::npos ? key : key.substr(0u, dot));
}

[[nodiscard]] uint32_t ColorOrDefault(const RedConfigure::Themes::ThemePreviewModel& model, std::wstring_view key, uint32_t fallback)
{
    return model.GetEffectiveColor(key).value_or(fallback);
}

BOOL CALLBACK CollectLocaleNameProc(LPWSTR localeName, DWORD, LPARAM userData) noexcept
{
    auto* names = reinterpret_cast<std::set<std::wstring>*>(userData);
    if (names && localeName && localeName[0] != L'\0')
    {
        names->insert(localeName);
    }
    return TRUE;
}

[[nodiscard]] std::vector<std::wstring> EnumerateOfficialCultureNames()
{
    std::set<std::wstring> names;
    static_cast<void>(::EnumSystemLocalesEx(CollectLocaleNameProc, LOCALE_ALL, reinterpret_cast<LPARAM>(&names), nullptr));
    if (names.empty())
    {
        names.insert(L"en-US");
        names.insert(L"fr-FR");
        names.insert(L"de-DE");
        names.insert(L"es-ES");
        names.insert(L"it-IT");
    }
    return std::vector<std::wstring>(names.begin(), names.end());
}

[[nodiscard]] std::wstring LocaleDisplayName(std::wstring_view culture)
{
    const std::wstring cultureText(culture);
    wchar_t buffer[256]{};
    int length = ::GetLocaleInfoEx(cultureText.c_str(), LOCALE_SENGLISHDISPLAYNAME, buffer, static_cast<int>(std::size(buffer)));
    if (length <= 1)
    {
        length = ::GetLocaleInfoEx(cultureText.c_str(), LOCALE_SLOCALIZEDDISPLAYNAME, buffer, static_cast<int>(std::size(buffer)));
    }
    if (length <= 1)
    {
        return cultureText;
    }
    return std::wstring(buffer, static_cast<size_t>(length - 1));
}

[[nodiscard]] std::wstring CultureDisplayText(std::wstring_view culture, std::wstring_view suffix)
{
    const std::wstring cultureText(culture);
    const std::wstring displayName = LocaleDisplayName(culture);
    if (displayName.empty() || displayName == cultureText)
    {
        return cultureText + L" (" + std::wstring(suffix) + L")";
    }
    return cultureText + L" - " + displayName + L" (" + std::wstring(suffix) + L")";
}

[[nodiscard]] std::set<std::wstring> DiscoverExistingCultures(const RedConfigure::Workspace::WorkspaceScanResult& workspace)
{
    std::set<std::wstring> cultures;
    for (const auto& owner : workspace.resourceOwners)
    {
        for (const std::filesystem::path& path : owner.satelliteResourcePaths)
        {
            const std::filesystem::path cultureFolder = path.parent_path();
            if (! cultureFolder.empty() && cultureFolder.parent_path().filename() == L"Lang")
            {
                cultures.insert(cultureFolder.filename().wstring());
            }
        }
    }
    return cultures;
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text)
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int required = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::wstring PlaceholderStatusText(HINSTANCE instance, RedConfigure::Localization::PlaceholderStatus status)
{
    switch (status)
    {
        case RedConfigure::Localization::PlaceholderStatus::Ok: return LoadAppString(instance, IDS_REDCONFIGURE_STATUS_OK);
        case RedConfigure::Localization::PlaceholderStatus::BarePlaceholder:
            return LoadAppString(instance, IDS_REDCONFIGURE_STATUS_BARE_PLACEHOLDER);
        case RedConfigure::Localization::PlaceholderStatus::UnindexedFormatSpec:
            return LoadAppString(instance, IDS_REDCONFIGURE_STATUS_UNINDEXED_FORMAT);
        case RedConfigure::Localization::PlaceholderStatus::PrintfPlaceholder:
            return LoadAppString(instance, IDS_REDCONFIGURE_STATUS_PRINTF_FORMAT);
        case RedConfigure::Localization::PlaceholderStatus::PlaceholderMismatch:
            return LoadAppString(instance, IDS_REDCONFIGURE_STATUS_PLACEHOLDER_MISMATCH);
        default: return {};
    }
}

[[nodiscard]] std::wstring LocalizableKindText(HINSTANCE instance, RedConfigure::Localization::RcLocalizableKind kind)
{
    switch (kind)
    {
        case RedConfigure::Localization::RcLocalizableKind::StringTable: return LoadAppString(instance, IDS_REDCONFIGURE_KIND_STRING);
        case RedConfigure::Localization::RcLocalizableKind::MenuPopup: return LoadAppString(instance, IDS_REDCONFIGURE_KIND_MENU_POPUP);
        case RedConfigure::Localization::RcLocalizableKind::MenuItem: return LoadAppString(instance, IDS_REDCONFIGURE_KIND_MENU_ITEM);
        case RedConfigure::Localization::RcLocalizableKind::DialogCaption: return LoadAppString(instance, IDS_REDCONFIGURE_KIND_DIALOG_CAPTION);
        case RedConfigure::Localization::RcLocalizableKind::DialogControl: return LoadAppString(instance, IDS_REDCONFIGURE_KIND_DIALOG_CONTROL);
        default: return {};
    }
}

class InventoryGridModel final : public IDxGridModel
{
public:
    InventoryGridModel(HINSTANCE instance, const RedConfigure::RedConfigureSession& session) : _instance(instance), _session(session)
    {
        _columns.push_back(GridColumnDesc{
            .id = L"kind", .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_KIND), .widthDip = 130.0f, .minWidthDip = 90.0f, .kind = GridColumnKind::Text, .sortable = false, .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"owner",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_OWNER),
            .widthDip = 150.0f,
            .minWidthDip = 90.0f,
            .kind = GridColumnKind::Text,
            .sortable = false,
            .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"id", .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_ID), .widthDip = 160.0f, .minWidthDip = 90.0f, .kind = GridColumnKind::Text, .sortable = false, .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"source",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_SOURCE),
            .widthDip = 360.0f,
            .minWidthDip = 160.0f,
            .kind = GridColumnKind::Text,
            .sortable = false,
            .multiline = false});
    }

    InventoryGridModel(const InventoryGridModel&)            = delete;
    InventoryGridModel& operator=(const InventoryGridModel&) = delete;
    InventoryGridModel(InventoryGridModel&&)                 = delete;
    InventoryGridModel& operator=(InventoryGridModel&&)      = delete;

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _session.GetInventoryEntries().size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return columnIndex < _columns.size() ? _columns[columnIndex] : GridColumnDesc{};
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        const auto rows = _session.GetInventoryEntries();
        if (rowIndex >= rows.size())
        {
            return;
        }

        const RedConfigure::InventoryEntry& row = rows[rowIndex];
        switch (columnIndex)
        {
            case 0u: outCell.text = LocalizableKindText(_instance, row.kind); break;
            case 1u: outCell.text = row.ownerName; break;
            case 2u: outCell.text = row.itemId.empty() ? row.resourceId : row.itemId; break;
            case 3u: outCell.text = row.sourceText; break;
            default: break;
        }
        outCell.tooltipText = outCell.text;
        outCell.multiline   = false;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return static_cast<uint64_t>(rowIndex) + 1u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId == 0u)
        {
            return std::nullopt;
        }

        const size_t rowIndex = static_cast<size_t>(rowId - 1u);
        return rowIndex < _session.GetInventoryEntries().size() ? std::optional<size_t>(rowIndex) : std::nullopt;
    }

private:
    HINSTANCE _instance = nullptr;
    const RedConfigure::RedConfigureSession& _session;
    std::vector<GridColumnDesc> _columns;
};

class TranslationGridModel final : public IDxGridModel
{
public:
    TranslationGridModel(HINSTANCE instance, const RedConfigure::RedConfigureSession& session) : _instance(instance), _session(session)
    {
        _columns.push_back(GridColumnDesc{
            .id = L"id", .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_ID), .widthDip = 160.0f, .minWidthDip = 90.0f, .kind = GridColumnKind::Text, .sortable = true, .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"source",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_SOURCE),
            .widthDip = 360.0f,
            .minWidthDip = 160.0f,
            .kind = GridColumnKind::Text,
            .sortable = true,
            .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"target",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_TARGET),
            .widthDip = 360.0f,
            .minWidthDip = 160.0f,
            .kind = GridColumnKind::Text,
            .sortable = true,
            .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"status",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_STATUS),
            .widthDip = 150.0f,
            .minWidthDip = 110.0f,
            .kind = GridColumnKind::Text,
            .sortable = true,
            .multiline = false});
    }

    TranslationGridModel(const TranslationGridModel&)            = delete;
    TranslationGridModel& operator=(const TranslationGridModel&) = delete;
    TranslationGridModel(TranslationGridModel&&)                 = delete;
    TranslationGridModel& operator=(TranslationGridModel&&)      = delete;

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _viewRows.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return columnIndex < _columns.size() ? _columns[columnIndex] : GridColumnDesc{};
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        const auto rows = _session.GetTranslations();
        const std::optional<size_t> sessionRow = ResolveSessionRow(rowIndex);
        if (! sessionRow || sessionRow.value() >= rows.size())
        {
            return;
        }

        const RedConfigure::TranslationEntry& row = rows[sessionRow.value()];
        switch (columnIndex)
        {
            case 0u: outCell.text = row.id; break;
            case 1u: outCell.text = row.sourceText; break;
            case 2u: outCell.text = row.targetText; break;
            case 3u: outCell.text = PlaceholderStatusText(_instance, row.validation.status); break;
            default: break;
        }
        outCell.tooltipText = outCell.text;
        outCell.multiline   = false;
    }

    [[nodiscard]] GridRowStyle GetRowStyle(size_t rowIndex) const override
    {
        const auto rows = _session.GetTranslations();
        const std::optional<size_t> sessionRow = ResolveSessionRow(rowIndex);
        if (! sessionRow || sessionRow.value() >= rows.size() || rows[sessionRow.value()].validation.status == RedConfigure::Localization::PlaceholderStatus::Ok)
        {
            return {};
        }

        return GridRowStyle{.tone = GridRowTone::Warning};
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        const std::optional<size_t> sessionRow = ResolveSessionRow(rowIndex);
        return sessionRow ? static_cast<uint64_t>(sessionRow.value()) + 1u : 0u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId == 0u)
        {
            return std::nullopt;
        }

        const size_t sessionRow = static_cast<size_t>(rowId - 1u);
        const auto it = std::find(_viewRows.begin(), _viewRows.end(), sessionRow);
        return it != _viewRows.end() ? std::optional<size_t>(static_cast<size_t>(std::distance(_viewRows.begin(), it))) : std::nullopt;
    }

    void SetViewRows(std::vector<size_t> viewRows)
    {
        _viewRows = std::move(viewRows);
    }

    [[nodiscard]] std::optional<size_t> ResolveSessionRow(size_t viewRow) const noexcept
    {
        return viewRow < _viewRows.size() ? std::optional<size_t>(_viewRows[viewRow]) : std::nullopt;
    }

private:
    HINSTANCE _instance = nullptr;
    const RedConfigure::RedConfigureSession& _session;
    std::vector<GridColumnDesc> _columns;
    std::vector<size_t> _viewRows;
};

class ThemeColorGridModel final : public IDxGridModel
{
public:
    ThemeColorGridModel(HINSTANCE instance, const RedConfigure::RedConfigureSession& session) : _instance(instance), _session(session)
    {
        _columns.push_back(GridColumnDesc{
            .id = L"key",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_COLOR_KEY),
            .widthDip = 210.0f,
            .minWidthDip = 140.0f,
            .kind = GridColumnKind::Text,
            .sortable = false,
            .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"effective",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_EFFECTIVE_VALUE),
            .widthDip = 120.0f,
            .minWidthDip = 90.0f,
            .kind = GridColumnKind::Text,
            .sortable = false,
            .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"authored",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_AUTHORED_VALUE),
            .widthDip = 190.0f,
            .minWidthDip = 120.0f,
            .kind = GridColumnKind::Text,
            .sortable = false,
            .multiline = false});
    }

    ThemeColorGridModel(const ThemeColorGridModel&)            = delete;
    ThemeColorGridModel& operator=(const ThemeColorGridModel&) = delete;
    ThemeColorGridModel(ThemeColorGridModel&&)                 = delete;
    ThemeColorGridModel& operator=(ThemeColorGridModel&&)      = delete;

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _keys.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return columnIndex < _columns.size() ? _columns[columnIndex] : GridColumnDesc{};
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        if (rowIndex >= _keys.size())
        {
            return;
        }

        const std::wstring& key = _keys[rowIndex];
        switch (columnIndex)
        {
            case 0u: outCell.text = key; break;
            case 1u:
                if (const std::optional<uint32_t> color = _session.GetThemePreviewModel().GetEffectiveColor(key))
                {
                    outCell.text           = Common::Settings::FormatColor(color.value());
                    outCell.hasSwatchValue = true;
                    outCell.swatchArgb     = color.value();
                }
                break;
            case 2u: outCell.text = _session.GetThemePreviewModel().GetAuthoredColorText(key); break;
            default: break;
        }

        outCell.tooltipText = outCell.text.empty() ? key : outCell.text;
        outCell.multiline   = false;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return static_cast<uint64_t>(rowIndex) + 1u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId == 0u)
        {
            return std::nullopt;
        }

        const size_t rowIndex = static_cast<size_t>(rowId - 1u);
        return rowIndex < _keys.size() ? std::optional<size_t>(rowIndex) : std::nullopt;
    }

    void SetKeys(std::vector<std::wstring> keys)
    {
        _keys = std::move(keys);
    }

    [[nodiscard]] std::optional<std::wstring> GetKeyAt(size_t rowIndex) const
    {
        return rowIndex < _keys.size() ? std::optional<std::wstring>(_keys[rowIndex]) : std::nullopt;
    }

    [[nodiscard]] std::optional<size_t> FindKey(std::wstring_view key) const noexcept
    {
        for (size_t index = 0u; index < _keys.size(); ++index)
        {
            if (_keys[index] == key)
            {
                return index;
            }
        }
        return std::nullopt;
    }

private:
    HINSTANCE _instance = nullptr;
    const RedConfigure::RedConfigureSession& _session;
    std::vector<GridColumnDesc> _columns;
    std::vector<std::wstring> _keys;
};

class ThemeExampleControl final : public Control
{
public:
    explicit ThemeExampleControl(HINSTANCE instance) :
        _navText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_NAV)),
        _menuText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_MENU)),
        _folderText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_FOLDER)),
        _hoverText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_HOVER)),
        _selectedText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_SELECTED)),
        _dialogText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_DIALOG)),
        _buttonText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_BUTTON)),
        _progressText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_PROGRESS)),
        _warningText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_WARNING)),
        _diffAddedText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_DIFF_ADDED)),
        _diffRemovedText(LoadAppString(instance, IDS_REDCONFIGURE_THEME_PREVIEW_DIFF_REMOVED))
    {
    }

    void SetModel(const RedConfigure::Themes::ThemePreviewModel* model) noexcept
    {
        _model = model;
        RequestInvalidate();
    }

    void Refresh() noexcept
    {
        RequestInvalidate();
    }

    void SetSelectedToken(std::wstring_view key)
    {
        _selectedTokenKey = std::wstring(key);
        RequestInvalidate();
    }

    void SetOnTokenSelected(std::function<void(std::wstring_view)> onTokenSelected)
    {
        _onTokenSelected = std::move(onTokenSelected);
    }

    bool OnMouseDown(WindowHost&, D2D1_POINT_2F point, bool rightButton, UINT) override
    {
        if (rightButton || ! _onTokenSelected)
        {
            return false;
        }

        const PreviewLayout layout = BuildLayout(GetBounds());
        const std::vector<RedConfigure::ThemePreviewHitCandidate> candidates = BuildHitCandidates(layout);
        const bool continuingCycle = IsSameClickPoint(point, _lastClickPoint);
        const std::wstring selectedKey =
            RedConfigure::SelectThemePreviewHitKey(candidates, point.x, point.y, continuingCycle ? std::wstring_view(_lastClickKey) : std::wstring_view{});
        if (! selectedKey.empty())
        {
            _lastClickPoint = point;
            _lastClickKey   = selectedKey;
            _onTokenSelected(selectedKey);
            return true;
        }

        return false;
    }

    void Paint(WindowHost& host) const override
    {
        auto* dc = host.GetDeviceContext();
        if (! dc || ! _model)
        {
            return;
        }

        const D2D1_RECT_F bounds = GetBounds();
        const PreviewLayout layout = BuildLayout(bounds);
        const ThemePalette& palette = host.GetTheme();

        const D2D1_COLOR_F windowColor        = ColorFromArgb(ColorOrDefault(*_model, L"window.background", 0xFFFFFFFFu));
        const D2D1_COLOR_F windowTextColor    = ColorFromArgb(ColorOrDefault(*_model, L"window.text", 0xFF111111u));
        const D2D1_COLOR_F navColor           = ColorFromArgb(ColorOrDefault(*_model, L"navigation.background", 0xFFEAF2FEu));
        const D2D1_COLOR_F navTextColor       = ColorFromArgb(ColorOrDefault(*_model, L"navigation.text", 0xFF1F2937u));
        const D2D1_COLOR_F navAccentColor     = ColorFromArgb(ColorOrDefault(*_model, L"app.accent", 0xFF0F6CBDu));
        const D2D1_COLOR_F menuColor          = ColorFromArgb(ColorOrDefault(*_model, L"menu.background", 0xFFFFFFFFu));
        const D2D1_COLOR_F menuTextColor      = ColorFromArgb(ColorOrDefault(*_model, L"menu.text", 0xFF111111u));
        const D2D1_COLOR_F menuSelectionColor = ColorFromArgb(ColorOrDefault(*_model, L"menu.selectionBackground", 0xFFE8F1FFu));
        const D2D1_COLOR_F menuBorderColor    = ColorFromArgb(ColorOrDefault(*_model, L"menu.border", 0xFFD8D8D8u));
        const D2D1_COLOR_F folderColor        = ColorFromArgb(ColorOrDefault(*_model, L"folderView.background", 0xFFFFFFFFu));
        const D2D1_COLOR_F folderTextColor    = ColorFromArgb(ColorOrDefault(*_model, L"folderView.itemForeground", 0xFF111111u));
        const D2D1_COLOR_F hoverColor         = ColorFromArgb(ColorOrDefault(*_model, L"folderView.itemBackgroundHovered", 0xFFF0F6FFu));
        const D2D1_COLOR_F selectedColor      = ColorFromArgb(ColorOrDefault(*_model, L"folderView.itemBackgroundSelected", 0xFFCFE8FFu));
        const D2D1_COLOR_F selectedTextColor  = ColorFromArgb(ColorOrDefault(*_model, L"folderView.itemForegroundSelected", 0xFF0F172Au));
        const D2D1_COLOR_F warningColor       = ColorFromArgb(ColorOrDefault(*_model, L"folderView.warningForeground", 0xFF8A4B00u));
        const D2D1_COLOR_F dialogColor        = ColorFromArgb(ColorOrDefault(*_model, L"dialog.background", 0xFFF7F7F7u));
        const D2D1_COLOR_F dialogTextColor    = ColorFromArgb(ColorOrDefault(*_model, L"dialog.text", 0xFF111111u));
        const D2D1_COLOR_F buttonColor        = ColorFromArgb(ColorOrDefault(*_model, L"dialog.buttonBackground", 0xFFFFFFFFu));
        const D2D1_COLOR_F buttonTextColor    = ColorFromArgb(ColorOrDefault(*_model, L"dialog.buttonText", 0xFF111111u));
        const D2D1_COLOR_F progressBgColor    = ColorFromArgb(ColorOrDefault(*_model, L"progress.background", 0xFFE5E7EBu));
        const D2D1_COLOR_F progressFillColor  = ColorFromArgb(ColorOrDefault(*_model, L"progress.fill", 0xFF2563EBu));
        const D2D1_COLOR_F diffAddedColor     = ColorFromArgb(ColorOrDefault(*_model, L"diff.addedBackground", 0xFFEAF8EFu));
        const D2D1_COLOR_F diffRemovedColor   = ColorFromArgb(ColorOrDefault(*_model, L"diff.removedBackground", 0xFFFFECEFu));

        if (auto* brush = host.GetSolidBrush(windowColor))
        {
            dc->FillRectangle(bounds, brush);
        }
        if (auto* brush = host.GetSolidBrush(palette.borderDefault))
        {
            dc->DrawRectangle(bounds, brush, 1.0f);
        }

        if (auto* brush = host.GetSolidBrush(navColor))
        {
            dc->FillRectangle(layout.navRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(navAccentColor))
        {
            dc->FillRectangle(layout.navAccentRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(menuColor))
        {
            dc->FillRectangle(layout.menuRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(menuSelectionColor))
        {
            dc->FillRectangle(layout.menuSelectionRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(folderColor))
        {
            dc->FillRectangle(layout.folderRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(hoverColor))
        {
            dc->FillRectangle(layout.hoverRowRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(selectedColor))
        {
            dc->FillRectangle(layout.selectedRowRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(dialogColor))
        {
            dc->FillRectangle(layout.dialogRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(buttonColor))
        {
            dc->FillRectangle(layout.buttonRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(progressBgColor))
        {
            dc->FillRectangle(layout.progressBgRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(progressFillColor))
        {
            dc->FillRectangle(layout.progressFillRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(diffAddedColor))
        {
            dc->FillRectangle(layout.diffAddedRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(diffRemovedColor))
        {
            dc->FillRectangle(layout.diffRemovedRect, brush);
        }
        if (auto* brush = host.GetSolidBrush(palette.borderDefault))
        {
            dc->DrawRectangle(layout.menuRect, brush, 1.0f);
            dc->DrawRectangle(layout.folderRect, brush, 1.0f);
            dc->DrawRectangle(layout.dialogRect, brush, 1.0f);
            dc->DrawRectangle(layout.buttonRect, brush, 1.0f);
        }
        if (auto* brush = host.GetSolidBrush(menuBorderColor))
        {
            dc->DrawLine(D2D1::Point2F(layout.menuRect.left, layout.menuRect.bottom),
                         D2D1::Point2F(layout.menuRect.right, layout.menuRect.bottom),
                         brush,
                         1.0f);
        }

        const auto drawText = [&](std::wstring_view text, const D2D1_RECT_F& rect, FontRole role, const D2D1_COLOR_F& color)
        {
            if (auto* brush = host.GetSolidBrush(color))
            {
                dc->DrawTextW(text.data(), static_cast<UINT32>(text.size()), host.GetTextFormat(role), rect, brush);
            }
        };

        drawText(_navText,
                 D2D1::RectF(layout.navRect.left + 14.0f, layout.navRect.top + 14.0f, layout.navRect.right - 14.0f, layout.navRect.top + 44.0f),
                 FontRole::BodyStrong,
                 navTextColor);
        drawText(_menuText,
                 D2D1::RectF(layout.menuRect.left + 12.0f, layout.menuRect.top + 8.0f, layout.menuRect.right - 12.0f, layout.menuRect.bottom),
                 FontRole::Body,
                 menuTextColor);
        drawText(_folderText,
                 D2D1::RectF(layout.folderRect.left + 12.0f, layout.folderRect.top + 10.0f, layout.folderRect.right - 12.0f, layout.folderRect.top + 34.0f),
                 FontRole::BodyStrong,
                 windowTextColor);
        drawText(_hoverText,
                 D2D1::RectF(layout.hoverRowRect.left + 10.0f, layout.hoverRowRect.top + 6.0f, layout.hoverRowRect.right, layout.hoverRowRect.bottom),
                 FontRole::Body,
                 folderTextColor);
        drawText(_selectedText,
                 D2D1::RectF(layout.selectedRowRect.left + 10.0f, layout.selectedRowRect.top + 6.0f, layout.selectedRowRect.right, layout.selectedRowRect.bottom),
                 FontRole::Body,
                 selectedTextColor);
        drawText(_warningText,
                 D2D1::RectF(layout.warningRowRect.left + 10.0f, layout.warningRowRect.top + 6.0f, layout.warningRowRect.right, layout.warningRowRect.bottom),
                 FontRole::Body,
                 warningColor);
        drawText(_dialogText,
                 D2D1::RectF(layout.dialogRect.left + 12.0f, layout.dialogRect.top + 10.0f, layout.dialogRect.right - 12.0f, layout.dialogRect.top + 38.0f),
                 FontRole::BodyStrong,
                 dialogTextColor);
        drawText(_buttonText,
                 D2D1::RectF(layout.buttonRect.left + 18.0f, layout.buttonRect.top + 7.0f, layout.buttonRect.right - 18.0f, layout.buttonRect.bottom),
                 FontRole::Body,
                 buttonTextColor);
        drawText(_progressText,
                 D2D1::RectF(layout.progressBgRect.left, layout.progressBgRect.top - 24.0f, layout.progressBgRect.right, layout.progressBgRect.top - 2.0f),
                 FontRole::Small,
                 folderTextColor);
        drawText(_diffAddedText, layout.diffAddedRect, FontRole::Small, folderTextColor);
        drawText(_diffRemovedText, layout.diffRemovedRect, FontRole::Small, folderTextColor);
        DrawSelectedTokenHighlight(host, layout);
    }

private:
    struct PreviewLayout
    {
        D2D1_RECT_F navRect{};
        D2D1_RECT_F navAccentRect{};
        D2D1_RECT_F menuRect{};
        D2D1_RECT_F menuSelectionRect{};
        D2D1_RECT_F folderRect{};
        D2D1_RECT_F hoverRowRect{};
        D2D1_RECT_F selectedRowRect{};
        D2D1_RECT_F warningRowRect{};
        D2D1_RECT_F dialogRect{};
        D2D1_RECT_F buttonRect{};
        D2D1_RECT_F progressBgRect{};
        D2D1_RECT_F progressFillRect{};
        D2D1_RECT_F diffAddedRect{};
        D2D1_RECT_F diffRemovedRect{};
    };

    struct PreviewHitRegion
    {
        D2D1_RECT_F rect{};
        std::wstring_view key;
    };

    [[nodiscard]] static std::array<PreviewHitRegion, 14> BuildHitRegions(const PreviewLayout& layout) noexcept
    {
        return {{
            {layout.navRect, L"navigation.background"},
            {layout.navAccentRect, L"app.accent"},
            {layout.menuRect, L"menu.background"},
            {layout.menuSelectionRect, L"menu.selectionBackground"},
            {layout.folderRect, L"folderView.background"},
            {layout.hoverRowRect, L"folderView.itemBackgroundHovered"},
            {layout.selectedRowRect, L"folderView.itemBackgroundSelected"},
            {layout.warningRowRect, L"folderView.warningForeground"},
            {layout.dialogRect, L"dialog.background"},
            {layout.buttonRect, L"dialog.buttonBackground"},
            {layout.progressBgRect, L"progress.background"},
            {layout.progressFillRect, L"progress.fill"},
            {layout.diffAddedRect, L"diff.addedBackground"},
            {layout.diffRemovedRect, L"diff.removedBackground"},
        }};
    }

    [[nodiscard]] static std::vector<RedConfigure::ThemePreviewHitCandidate> BuildHitCandidates(const PreviewLayout& layout)
    {
        const std::array<PreviewHitRegion, 14> regions = BuildHitRegions(layout);
        std::vector<RedConfigure::ThemePreviewHitCandidate> candidates;
        candidates.reserve(regions.size());
        for (const PreviewHitRegion& region : regions)
        {
            candidates.push_back(RedConfigure::ThemePreviewHitCandidate{
                .key = std::wstring(region.key),
                .left = region.rect.left,
                .top = region.rect.top,
                .right = region.rect.right,
                .bottom = region.rect.bottom});
        }
        return candidates;
    }

    [[nodiscard]] static bool IsSameClickPoint(const D2D1_POINT_2F& lhs, const D2D1_POINT_2F& rhs) noexcept
    {
        constexpr float thresholdDip = 3.0f;
        return std::abs(lhs.x - rhs.x) <= thresholdDip && std::abs(lhs.y - rhs.y) <= thresholdDip;
    }

    [[nodiscard]] static D2D1_RECT_F InflateLocalRect(const D2D1_RECT_F& rect, float amount) noexcept
    {
        return D2D1::RectF(rect.left - amount, rect.top - amount, rect.right + amount, rect.bottom + amount);
    }

    void DrawSelectedTokenHighlight(WindowHost& host, const PreviewLayout& layout) const
    {
        if (_selectedTokenKey.empty())
        {
            return;
        }

        auto* dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }

        const ThemePalette& palette = host.GetTheme();
        const std::array<PreviewHitRegion, 14> regions = BuildHitRegions(layout);
        for (const PreviewHitRegion& region : regions)
        {
            if (region.key != _selectedTokenKey)
            {
                continue;
            }

            const D2D1_RECT_F outer = InflateLocalRect(region.rect, 2.0f);
            if (auto* brush = host.GetSolidBrush(palette.focusStrokeOuter))
            {
                dc->DrawRectangle(outer, brush, 3.0f);
            }
            if (auto* brush = host.GetSolidBrush(palette.focusStroke))
            {
                dc->DrawRectangle(InflateLocalRect(region.rect, 0.5f), brush, 2.0f);
            }
        }
    }

    [[nodiscard]] static PreviewLayout BuildLayout(const D2D1_RECT_F& bounds) noexcept
    {
        const float inset = 16.0f;
        PreviewLayout layout{};
        layout.navRect       = D2D1::RectF(bounds.left + inset, bounds.top + inset, bounds.left + 172.0f, bounds.bottom - inset);
        layout.navAccentRect = D2D1::RectF(layout.navRect.left, layout.navRect.top, layout.navRect.left + 5.0f, layout.navRect.bottom);
        layout.menuRect      = D2D1::RectF(layout.navRect.right + 12.0f, bounds.top + inset, bounds.right - inset, bounds.top + inset + 36.0f);
        layout.menuSelectionRect = D2D1::RectF(layout.menuRect.left + 92.0f, layout.menuRect.top + 5.0f, layout.menuRect.left + 150.0f, layout.menuRect.bottom - 5.0f);
        layout.folderRect =
            D2D1::RectF(layout.navRect.right + 12.0f, layout.menuRect.bottom + 12.0f, bounds.right - inset, bounds.bottom - inset);

        const float rowLeft  = layout.folderRect.left + 12.0f;
        const float rowRight = std::max(rowLeft + 60.0f, layout.folderRect.right - 250.0f);
        float rowTop         = layout.folderRect.top + 48.0f;
        layout.hoverRowRect  = D2D1::RectF(rowLeft, rowTop, rowRight, rowTop + 30.0f);
        rowTop += 34.0f;
        layout.selectedRowRect = D2D1::RectF(rowLeft, rowTop, rowRight, rowTop + 30.0f);
        rowTop += 34.0f;
        layout.warningRowRect = D2D1::RectF(rowLeft, rowTop, rowRight, rowTop + 30.0f);

        layout.dialogRect =
            D2D1::RectF(std::max(rowRight + 16.0f, layout.folderRect.right - 230.0f), layout.folderRect.top + 48.0f, layout.folderRect.right - 12.0f, layout.folderRect.top + 154.0f);
        layout.buttonRect = D2D1::RectF(layout.dialogRect.left + 14.0f, layout.dialogRect.bottom - 42.0f, layout.dialogRect.left + 112.0f, layout.dialogRect.bottom - 12.0f);

        layout.progressBgRect =
            D2D1::RectF(rowLeft, std::max(layout.warningRowRect.bottom + 44.0f, layout.folderRect.bottom - 104.0f), rowRight, std::max(layout.warningRowRect.bottom + 56.0f, layout.folderRect.bottom - 92.0f));
        layout.progressFillRect = D2D1::RectF(layout.progressBgRect.left, layout.progressBgRect.top, layout.progressBgRect.left + ((layout.progressBgRect.right - layout.progressBgRect.left) * 0.62f), layout.progressBgRect.bottom);
        layout.diffAddedRect =
            D2D1::RectF(rowLeft, layout.progressBgRect.bottom + 22.0f, rowRight, layout.progressBgRect.bottom + 48.0f);
        layout.diffRemovedRect =
            D2D1::RectF(rowLeft, layout.diffAddedRect.bottom + 4.0f, rowRight, layout.diffAddedRect.bottom + 30.0f);
        return layout;
    }

    const RedConfigure::Themes::ThemePreviewModel* _model = nullptr;
    std::function<void(std::wstring_view)> _onTokenSelected;
    std::wstring _selectedTokenKey;
    D2D1_POINT_2F _lastClickPoint = D2D1::Point2F(-10000.0f, -10000.0f);
    std::wstring _lastClickKey;
    std::wstring _navText;
    std::wstring _menuText;
    std::wstring _folderText;
    std::wstring _hoverText;
    std::wstring _selectedText;
    std::wstring _dialogText;
    std::wstring _buttonText;
    std::wstring _progressText;
    std::wstring _warningText;
    std::wstring _diffAddedText;
    std::wstring _diffRemovedText;
};

class RedConfigureRoot final : public Panel, public RedConfigureRootController, public IDxGridDelegate
{
public:
    RedConfigureRoot(HINSTANCE instance, RedConfigure::RedConfigureSession& session, std::filesystem::path initialRoot) :
        _instance(instance),
        _session(session),
        _workspaceRoot(std::move(initialRoot)),
        _inventoryModel(instance, session),
        _translationModel(instance, session),
        _themeColorModel(instance, session)
    {
        BuildControls();
        SetPage(0u);
    }

    RedConfigureRoot(const RedConfigureRoot&)            = delete;
    RedConfigureRoot& operator=(const RedConfigureRoot&) = delete;
    RedConfigureRoot(RedConfigureRoot&&)                 = delete;
    RedConfigureRoot& operator=(RedConfigureRoot&&)      = delete;

    void ReloadWorkspaceFromFields()
    {
        _workspaceRoot = RedConfigure::ResolveWorkspaceRootForLaunchPath(std::filesystem::path(std::wstring(_workspaceRootField->GetText())));
        std::wstring culture = _cultureCombo ? std::wstring(_cultureCombo->GetSelectedValue()) : std::wstring(_session.GetCultureName());
        if (culture.empty())
        {
            culture = L"fr-FR";
        }
        const std::wstring previousOwner(_session.GetActiveResourceOwnerName());
        const HRESULT hr = _session.LoadWorkspace(_workspaceRoot, culture);
        if (SUCCEEDED(hr) && ! previousOwner.empty())
        {
            const auto& owners = _session.GetWorkspace().resourceOwners;
            for (size_t index = 0u; index < owners.size(); ++index)
            {
                if (owners[index].name == previousOwner)
                {
                    static_cast<void>(_session.SetActiveResourceOwner(index));
                    break;
                }
            }
        }
        _lastLoadSucceeded = SUCCEEDED(hr);
        _statusLabel->SetText(LoadAppString(_instance, _lastLoadSucceeded ? IDS_REDCONFIGURE_STATUS_WORKSPACE_READY : IDS_REDCONFIGURE_STATUS_WORKSPACE_FAILED));
        SyncFromSession();
    }

    void OnGridSelectionChanged(Grid& sender) override
    {
        if (&sender == _themeColorGrid)
        {
            if (_syncing)
            {
                return;
            }

            const std::optional<size_t> row = sender.GetPrimarySelectedRow();
            if (! row)
            {
                return;
            }

            if (const std::optional<std::wstring> key = _themeColorModel.GetKeyAt(row.value()))
            {
                SelectThemeColorKey(key.value());
            }
            return;
        }

        if (&sender != _translationGrid)
        {
            return;
        }

        const std::optional<size_t> row = sender.GetPrimarySelectedRow();
        if (row)
        {
            const std::optional<size_t> sessionRow = _translationModel.ResolveSessionRow(row.value());
            if (! sessionRow)
            {
                return;
            }
            _selectedTranslation = sessionRow.value();
            SyncTranslationEditor();
        }
    }

    void OnGridSelectionChanged() override {}

    void OnGridSortRequested(const GridSortSpec& sortSpec) override
    {
        _translationViewOptions.sortColumn = ColumnFromGridSort(sortSpec.columnIndex);
        _translationViewOptions.sortDirection = SortDirectionFromGrid(sortSpec.direction);
        RebuildTranslationView(true);
    }

protected:
    void OnBoundsChanged() noexcept override
    {
        LayoutControls();
    }

private:
    [[nodiscard]] static RedConfigure::LocalizationViewColumn ColumnFromGridSort(size_t columnIndex) noexcept
    {
        switch (columnIndex)
        {
            case 0u: return RedConfigure::LocalizationViewColumn::Id;
            case 1u: return RedConfigure::LocalizationViewColumn::Source;
            case 2u: return RedConfigure::LocalizationViewColumn::Target;
            case 3u: return RedConfigure::LocalizationViewColumn::Status;
            default: return RedConfigure::LocalizationViewColumn::Id;
        }
    }

    [[nodiscard]] static RedConfigure::LocalizationSortDirection SortDirectionFromGrid(SortDirection direction) noexcept
    {
        switch (direction)
        {
            case SortDirection::Ascending: return RedConfigure::LocalizationSortDirection::Ascending;
            case SortDirection::Descending: return RedConfigure::LocalizationSortDirection::Descending;
            case SortDirection::None: return RedConfigure::LocalizationSortDirection::None;
            default: return RedConfigure::LocalizationSortDirection::None;
        }
    }

    void BuildControls()
    {
        const auto pages = RedConfigure::GetPageDefinitions();
        _navButtons.reserve(pages.size());
        for (size_t index = 0u; index < pages.size(); ++index)
        {
            auto* button = AddChild<Button>(LoadAppString(_instance, pages[index].titleResourceId));
            button->SetOnClick([this, index] { SetPage(index); });
            _navButtons.push_back(button);
        }

        _titleLabel = AddChild<Label>();
        _titleLabel->SetFontRole(FontRole::Subtitle);
        _descriptionLabel = AddChild<Label>();
        _descriptionLabel->SetMultiline(true);
        _descriptionLabel->SetFontRole(FontRole::Small);
        _scopeLabel = AddChild<Label>();
        _scopeLabel->SetFontRole(FontRole::Small);
        _scopeLabel->SetMultiline(false);
        _statusLabel = AddChild<Label>();
        _statusLabel->SetFontRole(FontRole::Small);
        _statusLabel->SetMultiline(true);

        _workspaceRootLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_WORKSPACE_ROOT));
        _workspaceRootField = AddChild<TextField>(_workspaceRoot.wstring());
        _cultureLabel       = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_CULTURE));
        _cultureCombo       = AddChild<ComboBox>();
        _cultureCombo->SetOnSelectionChanged([this](size_t)
        {
            if (_syncing)
            {
                return;
            }
            ReloadWorkspaceFromFields();
        });
        _reloadButton       = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_RELOAD));
        _reloadButton->SetPrimary(true);
        _reloadButton->SetOnClick([this] { ReloadWorkspaceFromFields(); });
        _ownerCountLabel       = AddChild<Label>();
        _themeFileCountLabel   = AddChild<Label>();
        _scanErrorCountLabel   = AddChild<Label>();
        _ownerSelectorLabel    = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_ACTIVE_OWNER));
        _ownerCombo            = AddChild<ComboBox>();
        _ownerCombo->SetOnSelectionChanged([this](size_t selectedIndex)
        {
            if (_syncing)
            {
                return;
            }
            if (SUCCEEDED(_session.SetActiveResourceOwner(selectedIndex)))
            {
                _selectedTranslation = 0u;
                SyncFromSession();
            }
        });
        _activeOwnerLabel      = AddChild<Label>();
        _translationCountLabel = AddChild<Label>();
        _localizationSearchLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SEARCH));
        _localizationSearchField = AddChild<TextField>();
        _localizationSearchField->SetOnTextChanged([this](std::wstring_view text)
        {
            if (_syncing)
            {
                return;
            }
            _translationViewOptions.searchText.assign(text);
            RebuildTranslationView(true);
        });
        _localizationIdFilterLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_ID_FILTER));
        _localizationIdFilterField = AddChild<TextField>();
        _localizationIdFilterField->SetOnTextChanged([this](std::wstring_view text)
        {
            if (_syncing)
            {
                return;
            }
            _translationViewOptions.idFilterText.assign(text);
            RebuildTranslationView(true);
        });
        _localizationStatusFilterLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_COL_STATUS));
        _localizationStatusFilterCombo = AddChild<ComboBox>();
        std::vector<ComboBox::Item> statusFilterItems;
        statusFilterItems.push_back(ComboBox::Item{.value = L"all", .display = LoadAppString(_instance, IDS_REDCONFIGURE_FILTER_ALL)});
        statusFilterItems.push_back(ComboBox::Item{.value = L"ok", .display = LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_OK)});
        statusFilterItems.push_back(ComboBox::Item{.value = L"problems", .display = LoadAppString(_instance, IDS_REDCONFIGURE_FILTER_PROBLEMS)});
        _localizationStatusFilterCombo->SetItems(std::move(statusFilterItems));
        _localizationStatusFilterCombo->SetSelectedIndex(0u);
        _localizationStatusFilterCombo->SetOnSelectionChanged([this](size_t selectedIndex)
        {
            if (_syncing)
            {
                return;
            }
            _translationViewOptions.statusFilter = selectedIndex == 1u ? RedConfigure::LocalizationStatusFilter::Ok
                                                     : selectedIndex == 2u ? RedConfigure::LocalizationStatusFilter::Problems
                                                                          : RedConfigure::LocalizationStatusFilter::All;
            RebuildTranslationView(true);
        });

        _inventoryGrid = AddChild<Grid>();
        _inventoryGrid->SetModel(&_inventoryModel);
        _inventoryGrid->SetSelectionMode(GridSelectionMode::Single);
        _inventoryGrid->SetRowHeightDip(30.0f);
        _inventoryGrid->SetHeaderHeightDip(32.0f);
        _inventoryGrid->SetLineClamp(1u);
        _inventoryGrid->SetEmptyStateText(LoadAppString(_instance, IDS_REDCONFIGURE_LOCALIZATION_EMPTY));

        _translationGrid = AddChild<Grid>();
        _translationGrid->SetModel(&_translationModel);
        _translationGrid->SetDelegate(this);
        _translationGrid->SetSelectionMode(GridSelectionMode::Single);
        _translationGrid->SetRowHeightDip(30.0f);
        _translationGrid->SetHeaderHeightDip(32.0f);
        _translationGrid->SetLineClamp(1u);
        _translationGrid->SetEmptyStateText(LoadAppString(_instance, IDS_REDCONFIGURE_LOCALIZATION_EMPTY));

        _sourceTextLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SOURCE_TEXT));
        _sourceTextField = AddChild<TextField>();
        _sourceTextField->SetReadOnly(true);
        _sourceTextField->SetMultiline(true);
        _targetTextLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_TARGET_TEXT));
        _targetTextField = AddChild<TextField>();
        _targetTextField->SetMultiline(true);
        _targetTextField->SetOnTextChanged([this](std::wstring_view text) { OnTargetTextChanged(text); });
        _validationLabel = AddChild<Label>();
        _validationLabel->SetFontRole(FontRole::Small);
        _exportRcButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_EXPORT_RC));
        _exportRcButton->SetOnClick([this] { ExportLocalization(); });

        _themeSelectorLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_NAME));
        _themeCombo         = AddChild<ComboBox>();
        _themeCombo->SetOnSelectionChanged([this](size_t selectedIndex)
        {
            if (_syncing)
            {
                return;
            }
            if (_session.SetActiveTheme(selectedIndex))
            {
                SyncThemeLibraryLabels();
                SyncThemeColorKeyCombo({});
                SyncThemeColorEditor();
                SyncExportPathLabels();
                SyncExportPreviews();
                SyncScopeLabel();
                if (auto* host = GetHost())
                {
                    host->Invalidate();
                }
            }
        });
        _themeNameLabel  = AddChild<Label>();
        _themePathLabel  = AddChild<Label>();
        _themeErrorLabel = AddChild<Label>();

        _colorKeyFilterLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_KEY_FILTER));
        _colorKeyFilterField = AddChild<TextField>();
        _colorKeyFilterField->SetOnTextChanged([this](std::wstring_view)
        {
            if (_syncing)
            {
                return;
            }

            SyncThemeColorKeyCombo(std::wstring(_colorKeyCombo->GetSelectedValue()));
            SyncThemeColorEditor();
            if (auto* host = GetHost())
            {
                host->Invalidate();
            }
        });
        _colorKeyLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_COLOR_KEY));
        _colorKeyCombo = AddChild<ComboBox>();
        _colorKeyCombo->SetVisible(false);
        _themeColorGrid = AddChild<Grid>();
        _themeColorGrid->SetModel(&_themeColorModel);
        _themeColorGrid->SetDelegate(this);
        _themeColorGrid->SetSelectionMode(GridSelectionMode::Single);
        _themeColorGrid->SetRowHeightDip(28.0f);
        _themeColorGrid->SetHeaderHeightDip(30.0f);
        _themeColorGrid->SetLineClamp(1u);
        _themeColorGrid->SetEmptyStateText(LoadAppString(_instance, IDS_REDCONFIGURE_THEME_KEYS_EMPTY));
        SyncThemeColorKeyCombo({});
        _colorKeyCombo->SetOnSelectionChanged([this](size_t) { SyncThemeColorEditor(); });
        _previousColorLabel = AddChild<Label>();
        _previousColorLabel->SetFontRole(FontRole::Small);
        _previousColorSwatch = AddChild<ColorSwatch>();
        _colorValueLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_COLOR_VALUE));
        _colorSwatch     = AddChild<ColorSwatch>();
        _colorValueCombo = AddChild<ComboBox>();
        _colorValueCombo->SetEditable(true);
        _colorValueCombo->SetAutoOpenOnTextInput(true);
        _colorValueCombo->SetMaxVisibleItems(8u);
        _colorValueCombo->SetOnTextChanged([this](std::wstring_view text) { OnThemeColorTextChanged(text); });
        _themeColorStatusLabel = AddChild<Label>();
        _themeColorStatusLabel->SetFontRole(FontRole::Small);
        _themeExpressionHelpLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_EXPRESSION_EXAMPLES));
        _themeExpressionHelpLabel->SetFontRole(FontRole::Small);
        _themeExpressionHelpLabel->SetMultiline(true);
        _themePreviewScroll = AddChild<ScrollPanel>();
        _themePreviewScroll->SetScrollStepDip(56.0f);
        _themePreview = _themePreviewScroll->AddChild<ThemeExampleControl>(_instance);
        _themePreview->SetModel(&_session.GetThemePreviewModel());
        _themePreview->SetOnTokenSelected([this](std::wstring_view key) { SelectThemeColorKey(key); });
        _darkenGroupButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_DARKEN_GROUP));
        _darkenGroupButton->SetOnClick([this] { ApplyThemeGroupTransform(ThemeGroupTransform::Darken); });
        _blendAccentGroupButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_BLEND_ACCENT_GROUP));
        _blendAccentGroupButton->SetOnClick([this] { ApplyThemeGroupTransform(ThemeGroupTransform::BlendAccent); });
        _resetColorButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_RESET_COLOR));
        _resetColorButton->SetOnClick([this] { ResetSelectedThemeColor(); });
        _exportThemeButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_EXPORT_THEME));
        _exportThemeButton->SetOnClick([this] { ExportTheme(); });

        _defaultRcPathLabel    = AddChild<Label>();
        _defaultThemePathLabel = AddChild<Label>();
        _rcPreviewLabel        = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_RC_PREVIEW));
        _rcPreviewField        = AddChild<TextField>();
        _rcPreviewField->SetReadOnly(true);
        _rcPreviewField->SetMultiline(true);
        _themeJsonPreviewLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_PREVIEW));
        _themeJsonPreviewField = AddChild<TextField>();
        _themeJsonPreviewField->SetReadOnly(true);
        _themeJsonPreviewField->SetMultiline(true);
    }

    void SetPage(size_t pageIndex)
    {
        const auto pages = RedConfigure::GetPageDefinitions();
        if (pageIndex >= pages.size())
        {
            return;
        }

        _selectedPage = pageIndex;
        _titleLabel->SetText(LoadAppString(_instance, pages[pageIndex].titleResourceId));
        _descriptionLabel->SetText(LoadAppString(_instance, pages[pageIndex].descriptionResourceId));
        for (size_t index = 0u; index < _navButtons.size(); ++index)
        {
            _navButtons[index]->SetPrimary(index == pageIndex);
        }
        SyncScopeLabel();
        SyncVisibility();
        LayoutControls();
        if (auto* host = GetHost())
        {
            host->Invalidate();
        }
    }

    void SyncFromSession()
    {
        _syncing = true;
        _workspaceRootField->SetText(_workspaceRoot.wstring());
        _ownerCountLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_RESOURCE_OWNERS) + L": " +
                                  std::to_wstring(_session.GetWorkspace().resourceOwners.size()));
        _themeFileCountLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_FILES) + L": " +
                                      std::to_wstring(_session.GetThemeCatalog().themes.size()));
        _scanErrorCountLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SCAN_ERRORS) + L": " +
                                      std::to_wstring(_session.GetWorkspace().errors.size()));
        _activeOwnerLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_ACTIVE_OWNER) + L": " +
                                   std::wstring(_session.GetActiveResourceOwnerName()));
        _translationCountLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_TRANSLATION_COUNT) + L": " +
                                        std::to_wstring(_session.GetTranslations().size()));
        SyncScopeLabel();

        SyncCultureCombo();
        SyncOwnerCombo();
        SyncThemeCombo();
        SyncThemeLibraryLabels();
        SyncThemeColorKeyCombo({});
        SyncExportPathLabels();
        SyncExportPreviews();
        _inventoryGrid->NotifyDataChanged();
        RebuildTranslationView(true);
        SyncThemeColorEditor();
        _syncing = false;

        if (auto* host = GetHost())
        {
            host->Invalidate();
        }
    }

    void RebuildTranslationView(bool preserveSelection)
    {
        _translationViewRows = RedConfigure::BuildTranslationView(_session.GetTranslations(), _translationViewOptions);
        _translationModel.SetViewRows(_translationViewRows);
        _translationGrid->NotifyDataChanged();
        const size_t totalTranslations = _session.GetTranslations().size();
        std::wstring countText = LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_TRANSLATION_COUNT) + L": " +
                                 std::to_wstring(_translationViewRows.size());
        if (_translationViewRows.size() != totalTranslations)
        {
            countText += L" / " + std::to_wstring(totalTranslations);
        }
        _translationCountLabel->SetText(std::move(countText));

        if (_translationViewRows.empty())
        {
            _selectedTranslation = 0u;
            SyncTranslationEditor();
            return;
        }

        size_t viewRow = 0u;
        if (preserveSelection)
        {
            const auto it = std::find(_translationViewRows.begin(), _translationViewRows.end(), _selectedTranslation);
            if (it != _translationViewRows.end())
            {
                viewRow = static_cast<size_t>(std::distance(_translationViewRows.begin(), it));
            }
            else
            {
                _selectedTranslation = _translationViewRows.front();
            }
        }
        else
        {
            _selectedTranslation = _translationViewRows.front();
        }

        static_cast<void>(_translationGrid->RequestSelectRow(viewRow, 0u));
        SyncTranslationEditor();
    }

    void SyncThemeLibraryLabels()
    {
        const auto& catalog = _session.GetThemeCatalog();
        std::wstring nameText = LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_NAME) + L": ";
        std::wstring pathText = LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_PATH) + L": ";
        if (! catalog.themes.empty())
        {
            const size_t themeIndex = std::min(_session.GetActiveThemeIndex(), catalog.themes.size() - 1u);
            nameText += catalog.themes[themeIndex].definition.name;
            pathText += catalog.themes[themeIndex].path.wstring();
        }
        else
        {
            nameText += _session.GetThemePreviewModel().GetTheme().name;
        }

        _themeNameLabel->SetText(std::move(nameText));
        _themePathLabel->SetText(std::move(pathText));
        _themeErrorLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_ERRORS) + L": " + std::to_wstring(catalog.errors.size()));
    }

    void SyncCultureCombo()
    {
        std::wstring current(_session.GetCultureName());
        if (current.empty())
        {
            current = L"fr-FR";
        }

        const std::wstring existingSuffix = LoadAppString(_instance, IDS_REDCONFIGURE_CULTURE_EXISTING_SUFFIX);
        const std::wstring createSuffix   = LoadAppString(_instance, IDS_REDCONFIGURE_CULTURE_CREATE_SUFFIX);

        const std::set<std::wstring> existingCultures = DiscoverExistingCultures(_session.GetWorkspace());

        std::vector<ComboBox::Item> cultureItems;
        cultureItems.reserve(existingCultures.size() + 128u);
        std::set<std::wstring> added;
        for (const std::wstring& culture : existingCultures)
        {
            cultureItems.push_back(ComboBox::Item{.value = culture, .display = CultureDisplayText(culture, existingSuffix)});
            added.insert(culture);
        }

        for (const std::wstring& culture : EnumerateOfficialCultureNames())
        {
            if (added.insert(culture).second)
            {
                cultureItems.push_back(ComboBox::Item{.value = culture, .display = CultureDisplayText(culture, createSuffix)});
            }
        }
        if (added.insert(current).second)
        {
            cultureItems.push_back(ComboBox::Item{.value = current, .display = CultureDisplayText(current, createSuffix)});
        }

        size_t selectedIndex = 0u;
        for (size_t index = 0u; index < cultureItems.size(); ++index)
        {
            if (cultureItems[index].value == current)
            {
                selectedIndex = index;
                break;
            }
        }

        _cultureCombo->SetItems(std::move(cultureItems));
        _cultureCombo->SetSelectedIndex(selectedIndex);
    }

    void SyncOwnerCombo()
    {
        std::vector<ComboBox::Item> ownerItems;
        ownerItems.reserve(_session.GetWorkspace().resourceOwners.size());
        for (const auto& owner : _session.GetWorkspace().resourceOwners)
        {
            ownerItems.push_back(ComboBox::Item{.value = owner.name, .display = owner.name});
        }

        _ownerCombo->SetItems(std::move(ownerItems));
        if (_session.GetWorkspace().resourceOwners.empty())
        {
            _ownerCombo->SetSelectedIndex(std::nullopt);
        }
        else
        {
            _ownerCombo->SetSelectedIndex(std::min(_session.GetActiveResourceOwnerIndex(), _session.GetWorkspace().resourceOwners.size() - 1u));
        }
    }

    void SyncThemeCombo()
    {
        std::vector<ComboBox::Item> themeItems;
        const auto& catalog = _session.GetThemeCatalog();
        if (catalog.themes.empty())
        {
            const auto& fallback = _session.GetThemePreviewModel().GetTheme();
            themeItems.push_back(ComboBox::Item{.value = fallback.id, .display = fallback.name});
        }
        else
        {
            themeItems.reserve(catalog.themes.size());
            for (const auto& theme : catalog.themes)
            {
                themeItems.push_back(ComboBox::Item{.value = theme.definition.id, .display = theme.definition.name});
            }
        }

        _themeCombo->SetItems(std::move(themeItems));
        _themeCombo->SetSelectedIndex(catalog.themes.empty() ? std::optional<size_t>(0u)
                                                             : std::optional<size_t>(std::min(_session.GetActiveThemeIndex(), catalog.themes.size() - 1u)));
    }

    void SyncThemeColorKeyCombo(std::wstring_view preferredKey)
    {
        if (! _colorKeyCombo)
        {
            return;
        }

        std::wstring selectedKey(preferredKey);
        if (selectedKey.empty())
        {
            selectedKey = std::wstring(_colorKeyCombo->GetSelectedValue());
        }

        std::set<std::wstring> added;
        std::vector<std::wstring> keys;
        keys.reserve(kThemePreviewColorKeys.size() + _session.GetThemePreviewModel().GetTheme().colors.size());
        const auto addKey = [&added, &keys](std::wstring key)
        {
            if (added.insert(key).second)
            {
                keys.push_back(std::move(key));
            }
        };

        for (std::wstring_view key : kThemePreviewColorKeys)
        {
            addKey(std::wstring(key));
        }

        std::vector<std::wstring> authoredKeys;
        authoredKeys.reserve(_session.GetThemePreviewModel().GetTheme().colors.size());
        for (const auto& [key, _] : _session.GetThemePreviewModel().GetTheme().colors)
        {
            authoredKeys.push_back(key);
        }
        std::sort(authoredKeys.begin(), authoredKeys.end());
        for (std::wstring& key : authoredKeys)
        {
            addKey(std::move(key));
        }

        _themeColorKeys = std::move(keys);

        const std::wstring filterText = _colorKeyFilterField ? std::wstring(_colorKeyFilterField->GetText()) : std::wstring{};
        _filteredThemeColorKeys       = RedConfigure::FilterThemeColorKeys(_themeColorKeys, filterText);

        const bool previousSyncing = _syncing;
        _syncing                   = true;
        _themeColorModel.SetKeys(_filteredThemeColorKeys);
        if (_themeColorGrid)
        {
            _themeColorGrid->NotifyDataChanged();
        }

        std::vector<ComboBox::Item> colorItems;
        colorItems.reserve(_filteredThemeColorKeys.size());
        for (const std::wstring& key : _filteredThemeColorKeys)
        {
            colorItems.push_back(ComboBox::Item{.value = key, .display = key});
        }

        size_t selectedIndex = 0u;
        for (size_t index = 0u; index < _filteredThemeColorKeys.size(); ++index)
        {
            if (! selectedKey.empty() && _filteredThemeColorKeys[index] == selectedKey)
            {
                selectedIndex = index;
                break;
            }
        }

        _colorKeyCombo->SetItems(std::move(colorItems));
        _colorKeyCombo->SetSelectedIndex(_filteredThemeColorKeys.empty() ? std::optional<size_t>{} : std::optional<size_t>(selectedIndex));
        if (_themeColorGrid && ! _filteredThemeColorKeys.empty())
        {
            static_cast<void>(_themeColorGrid->RequestSelectRow(selectedIndex, 0u));
        }
        _syncing = previousSyncing;
    }

    void SyncExportPathLabels()
    {
        _defaultRcPathLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_DEFAULT_RC_PATH) + L": " +
                                     _session.GetDefaultLocalizationExportPath().wstring());
        _defaultThemePathLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_DEFAULT_THEME_PATH) + L": " +
                                        _session.GetDefaultThemeExportPath().wstring());
    }

    void SyncExportPreviews()
    {
        std::wstring rcPreview;
        if (SUCCEEDED(_session.BuildLocalizationExportText(rcPreview)))
        {
            _rcPreviewField->SetText(std::move(rcPreview));
        }
        else
        {
            _rcPreviewField->SetText({});
        }

        std::string themePreview;
        if (SUCCEEDED(_session.BuildThemeExportText(themePreview)))
        {
            _themeJsonPreviewField->SetText(Utf16FromUtf8(themePreview));
        }
        else
        {
            _themeJsonPreviewField->SetText({});
        }
    }

    void SyncTranslationEditor()
    {
        const auto rows = _session.GetTranslations();
        _syncing        = true;
        if (_selectedTranslation < rows.size())
        {
            const RedConfigure::TranslationEntry& row = rows[_selectedTranslation];
            _sourceTextField->SetText(row.sourceText);
            _targetTextField->SetText(row.targetText);
            SetValidationStatus(row.validation.status);
        }
        else
        {
            _sourceTextField->SetText({});
            _targetTextField->SetText({});
            SetValidationStatus(RedConfigure::Localization::PlaceholderStatus::Ok);
        }
        _syncing = false;
    }

    void SyncThemeColorEditor()
    {
        if (! _colorKeyCombo)
        {
            return;
        }

        const std::wstring key = std::wstring(_colorKeyCombo->GetSelectedValue());
        if (key.empty())
        {
            _colorSwatch->SetSwatchValue(std::nullopt);
            _previousColorSwatch->SetSwatchValue(std::nullopt);
            _previousColorLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_PREVIOUS_COLOR) + L":");
            _colorValueCombo->SetItems({});
            _colorValueCombo->SetText({});
            _themeColorStatusLabel->SetText({});
            _themePreview->SetSelectedToken({});
            return;
        }

        const std::wstring previousKey = (!_activeThemeColorKey.empty() && _activeThemeColorKey != key) ? _activeThemeColorKey : _previousThemeColorKey;
        _syncing = true;
        const std::optional<uint32_t> color = _session.GetThemePreviewModel().GetEffectiveColor(key);
        const std::wstring authoredText     = _session.GetThemePreviewModel().GetAuthoredColorText(key);
        const std::vector<std::wstring> suggestions = RedConfigure::BuildThemeColorSuggestions(key, previousKey, color);
        std::vector<ComboBox::Item> suggestionItems;
        suggestionItems.reserve(suggestions.size());
        for (const std::wstring& suggestion : suggestions)
        {
            suggestionItems.push_back(ComboBox::Item{.value = suggestion, .display = suggestion});
        }

        if (! previousKey.empty())
        {
            _previousColorSwatch->SetSwatchValue(_session.GetThemePreviewModel().GetEffectiveColor(previousKey));
            _previousColorLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_PREVIOUS_COLOR) + L": " + previousKey);
        }
        else
        {
            _previousColorSwatch->SetSwatchValue(std::nullopt);
            _previousColorLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_PREVIOUS_COLOR) + L":");
        }

        _colorSwatch->SetSwatchValue(color);
        _colorValueCombo->SetItems(std::move(suggestionItems));
        _colorValueCombo->SetText(! authoredText.empty() ? authoredText : (color ? Common::Settings::FormatColor(color.value()) : std::wstring{}));
        _themeColorStatusLabel->SetText({});
        _themePreview->SetSelectedToken(key);
        _themePreview->Refresh();
        if (_activeThemeColorKey != key)
        {
            _previousThemeColorKey = _activeThemeColorKey;
            _activeThemeColorKey   = key;
        }
        _syncing = false;
    }

    void OnTargetTextChanged(std::wstring_view text)
    {
        if (_syncing)
        {
            return;
        }

        const auto rows = _session.GetTranslations();
        if (_selectedTranslation >= rows.size())
        {
            return;
        }

        const auto validation = RedConfigure::Localization::ValidatePlaceholders(rows[_selectedTranslation].sourceText, text);
        SetValidationStatus(validation.status);
        if (_session.UpdateTranslation(_selectedTranslation, text))
        {
            RebuildTranslationView(true);
            SyncExportPathLabels();
            SyncExportPreviews();
        }
    }

    void OnThemeColorTextChanged(std::wstring_view text)
    {
        if (_syncing)
        {
            return;
        }

        const std::wstring key = std::wstring(_colorKeyCombo->GetSelectedValue());
        if (key.empty())
        {
            return;
        }

        if (_session.UpdateThemeColor(key, text))
        {
            const std::optional<uint32_t> color = _session.GetThemePreviewModel().GetEffectiveColor(key);
            _colorSwatch->SetSwatchValue(color);
            _themeColorStatusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_THEME_COLOR_VALID));
            _themeColorStatusLabel->SetTextColor(ResolvePalette().subduedText);
            _themeColorGrid->NotifyDataChanged();
            _themePreview->Refresh();
            SyncExportPathLabels();
            SyncExportPreviews();
        }
        else
        {
            _themeColorStatusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_THEME_COLOR_INVALID));
            _themeColorStatusLabel->SetTextColor(ResolvePalette().errorText);
        }
    }

    enum class ThemeGroupTransform : uint8_t
    {
        Darken,
        BlendAccent,
    };

    void SelectThemeColorKey(std::wstring_view key)
    {
        if (! _colorKeyCombo)
        {
            return;
        }

        for (size_t index = 0u; index < _filteredThemeColorKeys.size(); ++index)
        {
            if (_filteredThemeColorKeys[index] == key)
            {
                const bool previousSyncing = _syncing;
                _syncing                   = true;
                _colorKeyCombo->SetSelectedIndex(index);
                if (_themeColorGrid)
                {
                    static_cast<void>(_themeColorGrid->RequestSelectRow(index, 0u));
                }
                _syncing = previousSyncing;
                SyncThemeColorEditor();
                return;
            }
        }

        for (const std::wstring& availableKey : _themeColorKeys)
        {
            if (availableKey == key)
            {
                _syncing = true;
                _colorKeyFilterField->SetText({});
                _syncing = false;
                SyncThemeColorKeyCombo(key);
                SyncThemeColorEditor();
                return;
            }
        }
    }

    void ApplyThemeGroupTransform(ThemeGroupTransform transform)
    {
        const std::wstring selectedKey(_colorKeyCombo->GetSelectedValue());
        const std::wstring selectedGroup = ThemeKeyGroup(selectedKey);
        if (selectedGroup.empty())
        {
            return;
        }

        const std::optional<uint32_t> accent = _session.GetThemePreviewModel().GetEffectiveColor(L"app.accent");
        bool changed                         = false;
        for (const std::wstring& key : _themeColorKeys)
        {
            if (ThemeKeyGroup(key) != selectedGroup)
            {
                continue;
            }

            const std::optional<uint32_t> color = _session.GetThemePreviewModel().GetEffectiveColor(key);
            if (! color)
            {
                continue;
            }

            uint32_t transformed = color.value();
            switch (transform)
            {
                case ThemeGroupTransform::Darken: transformed = MixArgb(color.value(), 0xFF000000u, 0.10); break;
                case ThemeGroupTransform::BlendAccent:
                    if (! accent || key == L"app.accent")
                    {
                        continue;
                    }
                    transformed = MixArgb(color.value(), accent.value(), 0.16);
                    break;
                default: break;
            }

            changed = _session.UpdateThemeColor(key, Common::Settings::FormatColor(transformed)) || changed;
        }

        if (changed)
        {
            SyncThemeColorEditor();
            _themeColorGrid->NotifyDataChanged();
            SyncExportPathLabels();
            SyncExportPreviews();
            _themePreview->Refresh();
        }
    }

    void ResetSelectedThemeColor()
    {
        const std::wstring selectedKey(_colorKeyCombo->GetSelectedValue());
        if (selectedKey.empty())
        {
            return;
        }

        _session.GetThemePreviewModel().ResetOverride(selectedKey);
        SyncThemeColorEditor();
        _themeColorGrid->NotifyDataChanged();
        SyncExportPathLabels();
        SyncExportPreviews();
        _themePreview->Refresh();
    }

    void SetValidationStatus(RedConfigure::Localization::PlaceholderStatus status)
    {
        _validationLabel->SetText(PlaceholderStatusText(_instance, status));
        const ThemePalette palette = ResolvePalette();
        _validationLabel->SetTextColor(status == RedConfigure::Localization::PlaceholderStatus::Ok ? palette.subduedText : palette.errorText);
        SyncScopeLabel();
    }

    void SyncScopeLabel()
    {
        if (! _scopeLabel)
        {
            return;
        }

        std::wstring text = LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SCOPE) + L": ";
        if (_selectedPage == 0u)
        {
            text += LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_RESOURCE_OWNERS) + L" " +
                    std::to_wstring(_session.GetWorkspace().resourceOwners.size());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_FILES) + L" " +
                    std::to_wstring(_session.GetThemeCatalog().themes.size());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SCAN_ERRORS) + L" " +
                    std::to_wstring(_session.GetWorkspace().errors.size());
        }
        else
        {
            text += LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_CULTURE) + L" " + std::wstring(_session.GetCultureName());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_ACTIVE_OWNER) + L" " +
                    std::wstring(_session.GetActiveResourceOwnerName());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_NAME) + L" " + _session.GetThemePreviewModel().GetTheme().name;
        }
        _scopeLabel->SetText(std::move(text));
    }

    [[nodiscard]] ThemePalette ResolvePalette() const noexcept
    {
        return GetHost() ? GetHost()->GetTheme() : MakeDefaultThemePalette(false);
    }

    void ExportLocalization()
    {
        const std::filesystem::path path = _session.GetDefaultLocalizationExportPath();
        const HRESULT hr                = _session.ExportLocalization(path);
        _statusLabel->SetText(LoadAppString(_instance, SUCCEEDED(hr) ? IDS_REDCONFIGURE_STATUS_EXPORT_RC_DONE : IDS_REDCONFIGURE_STATUS_EXPORT_FAILED) +
                              L": " + path.wstring());
    }

    void ExportTheme()
    {
        const std::filesystem::path path = _session.GetDefaultThemeExportPath();
        const HRESULT hr                = _session.ExportTheme(path);
        _statusLabel->SetText(LoadAppString(_instance, SUCCEEDED(hr) ? IDS_REDCONFIGURE_STATUS_EXPORT_THEME_DONE : IDS_REDCONFIGURE_STATUS_EXPORT_FAILED) +
                              L": " + path.wstring());
    }

    void SyncVisibility() noexcept
    {
        auto setVisible = [](std::span<Control* const> controls, bool visible) noexcept
        {
            for (Control* control : controls)
            {
                if (control)
                {
                    control->SetVisible(visible);
                }
            }
        };

        Control* workspaceControls[] = {_workspaceRootLabel, _workspaceRootField, _reloadButton, _ownerCountLabel, _themeFileCountLabel, _scanErrorCountLabel};
        Control* localizationControls[] = {_cultureLabel, _cultureCombo, _ownerSelectorLabel, _ownerCombo, _translationCountLabel, _activeOwnerLabel,
                                           _localizationSearchLabel, _localizationSearchField, _localizationIdFilterLabel,
                                           _localizationIdFilterField, _localizationStatusFilterLabel, _localizationStatusFilterCombo,
                                           _translationGrid, _sourceTextLabel, _sourceTextField, _targetTextLabel, _targetTextField,
                                           _validationLabel};
        Control* themeDesignerControls[] = {
            _themeSelectorLabel, _themeCombo, _themeNameLabel, _themePathLabel, _themeErrorLabel, _colorKeyFilterLabel, _colorKeyFilterField,
            _colorKeyLabel, _themeColorGrid, _colorValueLabel,
            _previousColorLabel, _previousColorSwatch, _colorSwatch, _colorValueCombo, _themeColorStatusLabel, _themeExpressionHelpLabel,
            _darkenGroupButton, _blendAccentGroupButton, _resetColorButton, _themePreviewScroll, _themePreview};
        Control* exportControls[] = {
            _defaultRcPathLabel, _defaultThemePathLabel, _exportRcButton, _exportThemeButton, _rcPreviewLabel, _rcPreviewField,
            _themeJsonPreviewLabel, _themeJsonPreviewField};

        setVisible(workspaceControls, false);
        setVisible(localizationControls, false);
        setVisible(themeDesignerControls, false);
        setVisible(exportControls, false);

        switch (_selectedPage)
        {
            case 0u: setVisible(workspaceControls, true); break;
            case 1u: setVisible(localizationControls, true); break;
            case 2u: setVisible(themeDesignerControls, true); break;
            case 3u: setVisible(exportControls, true); break;
            default: break;
        }

        if (_colorKeyCombo)
        {
            _colorKeyCombo->SetVisible(false);
        }
    }

    void LayoutControls() noexcept
    {
        const D2D1_RECT_F bounds = GetBounds();
        const float width        = std::max(0.0f, bounds.right - bounds.left);
        const float height       = std::max(0.0f, bounds.bottom - bounds.top);
        const float margin       = 18.0f;
        const float gap          = 12.0f;
        const float navWidth     = 220.0f;
        const float navButtonH   = 36.0f;
        const float contentLeft  = margin + navWidth + 18.0f;
        const float contentRight = std::max(contentLeft, width - margin);
        const float contentW     = std::max(0.0f, contentRight - contentLeft);
        const float contentTop   = margin;
        const float bodyTop      = contentTop + 110.0f;
        const float statusTop    = std::max(bodyTop, height - margin - 42.0f);
        const float bodyBottom   = std::max(bodyTop, statusTop - gap);

        float navY = margin;
        for (Button* button : _navButtons)
        {
            button->SetBounds(D2D1::RectF(margin, navY, margin + navWidth, navY + navButtonH));
            navY += navButtonH + 8.0f;
        }

        _titleLabel->SetBounds(D2D1::RectF(contentLeft, contentTop, contentRight, contentTop + 34.0f));
        _descriptionLabel->SetBounds(D2D1::RectF(contentLeft, contentTop + 38.0f, contentRight, contentTop + 76.0f));
        _scopeLabel->SetBounds(D2D1::RectF(contentLeft, contentTop + 80.0f, contentRight, contentTop + 104.0f));
        _statusLabel->SetBounds(D2D1::RectF(contentLeft, statusTop, contentRight, height - margin));

        switch (_selectedPage)
        {
            case 0u: LayoutWorkspacePage(contentLeft, bodyTop, contentRight, bodyBottom); break;
            case 1u: LayoutLocalizationWorkbenchPage(contentLeft, bodyTop, contentRight, bodyBottom); break;
            case 2u: LayoutThemeWorkbenchPage(contentLeft, bodyTop, contentRight, bodyBottom); break;
            case 3u: LayoutExportPage(contentLeft, bodyTop, contentRight, bodyBottom); break;
            default: break;
        }

        (void)contentW;
    }

    void LayoutWorkspacePage(float left, float top, float right, float bottom) noexcept
    {
        const float labelW  = 140.0f;
        const float rowH    = 32.0f;
        const float buttonW = 120.0f;
        const float gapX    = 12.0f;
        float y             = top;
        _workspaceRootLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
        _workspaceRootField->SetBounds(D2D1::RectF(left + labelW, y, std::max(left + labelW, right - buttonW - gapX), y + rowH));
        _reloadButton->SetBounds(D2D1::RectF(right - 120.0f, y, right, y + rowH));
        y += rowH + 30.0f;
        _ownerCountLabel->SetBounds(D2D1::RectF(left, y, right, y + 28.0f));
        _themeFileCountLabel->SetBounds(D2D1::RectF(left, y + 34.0f, right, y + 62.0f));
        _scanErrorCountLabel->SetBounds(D2D1::RectF(left, y + 68.0f, right, std::min(bottom, y + 96.0f)));
    }

    void LayoutLocalizationWorkbenchPage(float left, float top, float right, float bottom) noexcept
    {
        const float rowH        = 32.0f;
        const float rowGap      = 12.0f;
        const float labelW      = 112.0f;
        const float contentW    = std::max(0.0f, right - left);
        const bool compact      = contentW < 760.0f;
        const float fieldGap    = 10.0f;
        float y                 = top;

        if (compact)
        {
            _cultureLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
            _cultureCombo->SetBounds(D2D1::RectF(left + labelW, y, right, y + rowH));
            y += rowH + 8.0f;
            _ownerSelectorLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
            _ownerCombo->SetBounds(D2D1::RectF(left + labelW, y, right, y + rowH));
            y += rowH + 10.0f;
            _activeOwnerLabel->SetBounds(D2D1::RectF(left, y, right, y + 24.0f));
            _translationCountLabel->SetBounds(D2D1::RectF(left, y + 26.0f, right, y + 50.0f));
            y += 58.0f;

            _localizationSearchLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + 70.0f, y + rowH));
            _localizationSearchField->SetBounds(D2D1::RectF(left + 72.0f, y, right, y + rowH));
            y += rowH + 8.0f;
            const float filterMid = left + ((contentW - rowGap) / 2.0f);
            _localizationIdFilterLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + 42.0f, y + rowH));
            _localizationIdFilterField->SetBounds(D2D1::RectF(left + 44.0f, y, filterMid, y + rowH));
            _localizationStatusFilterLabel->SetBounds(D2D1::RectF(filterMid + rowGap, y + 6.0f, filterMid + rowGap + 72.0f, y + rowH));
            _localizationStatusFilterCombo->SetBounds(D2D1::RectF(filterMid + rowGap + 74.0f, y, right, y + rowH));
            y += rowH + 12.0f;
        }
        else
        {
            const float halfW     = ((contentW - rowGap) / 2.0f);
            const float rightRowL = left + halfW + rowGap;
            _cultureLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
            _cultureCombo->SetBounds(D2D1::RectF(left + labelW, y, left + halfW, y + rowH));
            _ownerSelectorLabel->SetBounds(D2D1::RectF(rightRowL, y + 6.0f, rightRowL + labelW, y + rowH));
            _ownerCombo->SetBounds(D2D1::RectF(rightRowL + labelW, y, right, y + rowH));
            y += rowH + 10.0f;

            _activeOwnerLabel->SetBounds(D2D1::RectF(left, y, left + halfW, y + 24.0f));
            _translationCountLabel->SetBounds(D2D1::RectF(rightRowL, y, right, y + 24.0f));
            y += 32.0f;

            const float searchW = std::clamp(contentW * 0.36f, 240.0f, 520.0f);
            const float idW     = std::clamp(contentW * 0.22f, 180.0f, 320.0f);
            const float statusW = 170.0f;
            float x             = left;
            _localizationSearchLabel->SetBounds(D2D1::RectF(x, y + 6.0f, x + 70.0f, y + rowH));
            _localizationSearchField->SetBounds(D2D1::RectF(x + 72.0f, y, x + searchW, y + rowH));
            x += searchW + rowGap;
            _localizationIdFilterLabel->SetBounds(D2D1::RectF(x, y + 6.0f, x + 42.0f, y + rowH));
            _localizationIdFilterField->SetBounds(D2D1::RectF(x + 44.0f, y, x + idW, y + rowH));
            x += idW + rowGap;
            _localizationStatusFilterLabel->SetBounds(D2D1::RectF(x, y + 6.0f, x + 72.0f, y + rowH));
            _localizationStatusFilterCombo->SetBounds(D2D1::RectF(x + 74.0f, y, std::min(right, x + 74.0f + statusW), y + rowH));
            y += rowH + 12.0f;
        }

        const float gridTop       = y;
        const float validationH   = 24.0f;
        const float labelH        = 24.0f;
        const float minGridH      = 84.0f;
        const float maxEditorH    = compact ? 96.0f : 150.0f;
        const float availableH    = std::max(0.0f, bottom - gridTop);
        float editorFieldH        = std::clamp(availableH * 0.28f, 54.0f, maxEditorH);
        float gridH               = availableH - editorFieldH - labelH - validationH - 34.0f;
        if (gridH < minGridH)
        {
            gridH        = std::min(minGridH, std::max(32.0f, availableH * 0.46f));
            editorFieldH = std::max(42.0f, availableH - gridH - labelH - validationH - 34.0f);
        }

        const float gridBottom = std::min(bottom, gridTop + gridH);
        const float editorTop  = gridBottom + 10.0f;
        const float fieldTop   = editorTop + labelH + 4.0f;
        const float validationTop = std::max(fieldTop + 20.0f, bottom - validationH);
        const float fieldBottom   = std::max(fieldTop + 20.0f, validationTop - 8.0f);
        const float editorHalfW   = std::max(100.0f, (right - left - fieldGap) / 2.0f);

        _translationGrid->SetBounds(D2D1::RectF(left, gridTop, right, gridBottom));
        _sourceTextLabel->SetBounds(D2D1::RectF(left, editorTop, left + editorHalfW, editorTop + 24.0f));
        _targetTextLabel->SetBounds(D2D1::RectF(left + editorHalfW + fieldGap, editorTop, right, editorTop + 24.0f));
        _sourceTextField->SetBounds(D2D1::RectF(left, fieldTop, left + editorHalfW, fieldBottom));
        _targetTextField->SetBounds(D2D1::RectF(left + editorHalfW + fieldGap, fieldTop, right, fieldBottom));
        _validationLabel->SetBounds(D2D1::RectF(left, validationTop, right, bottom));
    }

    void LayoutTranslationEditorPage(float left, float top, float right, float bottom) noexcept
    {
        const float editorTop = std::max(top + 180.0f, bottom - 190.0f);
        const float halfW     = (right - left - 12.0f) / 2.0f;
        _translationGrid->SetBounds(D2D1::RectF(left, top, right, editorTop - 12.0f));
        _sourceTextLabel->SetBounds(D2D1::RectF(left, editorTop, left + halfW, editorTop + 24.0f));
        _targetTextLabel->SetBounds(D2D1::RectF(left + halfW + 12.0f, editorTop, right, editorTop + 24.0f));
        _sourceTextField->SetBounds(D2D1::RectF(left, editorTop + 28.0f, left + halfW, bottom - 42.0f));
        _targetTextField->SetBounds(D2D1::RectF(left + halfW + 12.0f, editorTop + 28.0f, right, bottom - 42.0f));
        _validationLabel->SetBounds(D2D1::RectF(left, bottom - 32.0f, right - 140.0f, bottom));
        _exportRcButton->SetBounds(D2D1::RectF(right - 120.0f, bottom - 34.0f, right, bottom - 2.0f));
    }

    void LayoutThemeLibraryPage(float left, float top, float right, float) noexcept
    {
        _themeSelectorLabel->SetBounds(D2D1::RectF(left, top + 6.0f, left + 100.0f, top + 32.0f));
        _themeCombo->SetBounds(D2D1::RectF(left + 100.0f, top, right, top + 32.0f));
        _themeNameLabel->SetBounds(D2D1::RectF(left, top + 48.0f, right, top + 74.0f));
        _themePathLabel->SetBounds(D2D1::RectF(left, top + 82.0f, right, top + 108.0f));
        _themeFileCountLabel->SetBounds(D2D1::RectF(left, top + 116.0f, right, top + 142.0f));
        _themeErrorLabel->SetBounds(D2D1::RectF(left, top + 150.0f, right, top + 176.0f));
    }

    void LayoutThemeWorkbenchPage(float left, float top, float right, float bottom) noexcept
    {
        const float labelW   = 108.0f;
        const float rowH     = 32.0f;
        const float rowGap   = 10.0f;
        const float contentW = std::max(0.0f, right - left);
        _themeSelectorLabel->SetBounds(D2D1::RectF(left, top + 6.0f, left + labelW, top + rowH));
        _themeCombo->SetBounds(D2D1::RectF(left + labelW, top, right, top + rowH));
        _themeNameLabel->SetBounds(D2D1::RectF(left, top + 42.0f, right, top + 68.0f));
        _themePathLabel->SetBounds(D2D1::RectF(left, top + 68.0f, right, top + 92.0f));
        _themeErrorLabel->SetBounds(D2D1::RectF(left, top + 94.0f, right, top + 118.0f));

        const bool sideBySide = contentW >= 820.0f;
        const float editorTop = top + 128.0f;
        const float editorRight =
            sideBySide ? std::min(left + 460.0f, left + std::max(390.0f, contentW * 0.46f)) : right;
        const float previewLeft = sideBySide ? editorRight + 18.0f : left;
        float previewTop        = sideBySide ? editorTop : bottom;

        float y = editorTop;
        _colorKeyFilterLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
        _colorKeyFilterField->SetBounds(D2D1::RectF(left + labelW, y, editorRight, y + rowH));
        y += rowH + rowGap;

        _colorKeyLabel->SetBounds(D2D1::RectF(left, y, editorRight, y + 24.0f));
        y += 26.0f;
        const float keyGridH = sideBySide ? 176.0f : 132.0f;
        _themeColorGrid->SetBounds(D2D1::RectF(left, y, editorRight, y + keyGridH));
        y += keyGridH + rowGap;

        _previousColorSwatch->SetBounds(D2D1::RectF(left + labelW, y + 4.0f, left + labelW + 24.0f, y + 28.0f));
        _previousColorLabel->SetBounds(D2D1::RectF(left + labelW + 34.0f, y + 6.0f, editorRight, y + rowH));
        y += rowH + rowGap;

        _colorValueLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
        _colorSwatch->SetBounds(D2D1::RectF(left + labelW, y, left + labelW + rowH, y + rowH));
        _colorValueCombo->SetBounds(D2D1::RectF(left + labelW + rowH + 10.0f, y, editorRight, y + rowH));
        y += rowH + 4.0f;
        _themeColorStatusLabel->SetBounds(D2D1::RectF(left + labelW, y, editorRight, y + 22.0f));
        y += 24.0f;

        const float buttonGap = 8.0f;
        const float buttonW   = std::max(88.0f, (editorRight - (left + labelW) - (buttonGap * 2.0f)) / 3.0f);
        const float buttonX   = left + labelW;
        _darkenGroupButton->SetBounds(D2D1::RectF(buttonX, y, buttonX + buttonW, y + rowH));
        _blendAccentGroupButton->SetBounds(D2D1::RectF(buttonX + buttonW + buttonGap, y, buttonX + (buttonW * 2.0f) + buttonGap, y + rowH));
        _resetColorButton->SetBounds(D2D1::RectF(buttonX + (buttonW * 2.0f) + (buttonGap * 2.0f), y, editorRight, y + rowH));
        y += rowH + rowGap;

        const float helpBottom = sideBySide ? std::min(bottom, y + 82.0f) : std::min(bottom, y + 64.0f);
        _themeExpressionHelpLabel->SetBounds(D2D1::RectF(left, y, editorRight, std::max(y + 28.0f, helpBottom)));
        if (! sideBySide)
        {
            previewTop = std::min(bottom, std::max(y + 36.0f, helpBottom + 8.0f));
        }

        const D2D1_RECT_F previewBounds = D2D1::RectF(previewLeft, previewTop, right, bottom);
        _themePreviewScroll->SetBounds(previewBounds);
        const float previewContentHeight = std::max(kThemePreviewContentHeightDip, previewBounds.bottom - previewBounds.top);
        _themePreviewScroll->SetContentHeight(previewContentHeight);
        const float previewContentRight = std::max(previewBounds.left + 120.0f, previewBounds.right - _themePreviewScroll->GetScrollbarThickness() - 2.0f);
        _themePreview->SetBounds(D2D1::RectF(previewBounds.left, previewBounds.top, previewContentRight, previewBounds.top + previewContentHeight));
    }

    void LayoutExportPage(float left, float top, float right, float) noexcept
    {
        _defaultRcPathLabel->SetBounds(D2D1::RectF(left, top, right, top + 42.0f));
        _exportRcButton->SetBounds(D2D1::RectF(left, top + 48.0f, left + 140.0f, top + 80.0f));
        _defaultThemePathLabel->SetBounds(D2D1::RectF(left + 170.0f, top, right, top + 42.0f));
        _exportThemeButton->SetBounds(D2D1::RectF(left + 170.0f, top + 48.0f, left + 310.0f, top + 80.0f));

        const float previewTop = top + 108.0f;
        const float gap        = 12.0f;
        const float mid        = left + ((right - left - gap) / 2.0f);
        const float fieldBottom = std::max(previewTop + 58.0f, GetBounds().bottom - 18.0f);
        _rcPreviewLabel->SetBounds(D2D1::RectF(left, previewTop, mid, previewTop + 24.0f));
        _rcPreviewField->SetBounds(D2D1::RectF(left, previewTop + 28.0f, mid, fieldBottom));
        _themeJsonPreviewLabel->SetBounds(D2D1::RectF(mid + gap, previewTop, right, previewTop + 24.0f));
        _themeJsonPreviewField->SetBounds(D2D1::RectF(mid + gap, previewTop + 28.0f, right, fieldBottom));
    }

    HINSTANCE _instance = nullptr;
    RedConfigure::RedConfigureSession& _session;
    std::filesystem::path _workspaceRoot;
    InventoryGridModel _inventoryModel;
    TranslationGridModel _translationModel;
    ThemeColorGridModel _themeColorModel;
    bool _syncing           = false;
    bool _lastLoadSucceeded = false;
    size_t _selectedPage    = 0u;
    size_t _selectedTranslation = 0u;
    RedConfigure::LocalizationViewOptions _translationViewOptions;
    std::vector<size_t> _translationViewRows;
    std::vector<std::wstring> _themeColorKeys;
    std::vector<std::wstring> _filteredThemeColorKeys;
    std::wstring _activeThemeColorKey;
    std::wstring _previousThemeColorKey;

    std::vector<Button*> _navButtons;
    Label* _titleLabel       = nullptr;
    Label* _descriptionLabel = nullptr;
    Label* _scopeLabel       = nullptr;
    Label* _statusLabel      = nullptr;

    Label* _workspaceRootLabel = nullptr;
    TextField* _workspaceRootField = nullptr;
    Label* _cultureLabel          = nullptr;
    ComboBox* _cultureCombo       = nullptr;
    Button* _reloadButton         = nullptr;
    Label* _ownerCountLabel       = nullptr;
    Label* _themeFileCountLabel   = nullptr;
    Label* _scanErrorCountLabel   = nullptr;
    Label* _ownerSelectorLabel    = nullptr;
    ComboBox* _ownerCombo         = nullptr;
    Label* _activeOwnerLabel      = nullptr;
    Label* _translationCountLabel = nullptr;
    Label* _localizationSearchLabel = nullptr;
    TextField* _localizationSearchField = nullptr;
    Label* _localizationIdFilterLabel = nullptr;
    TextField* _localizationIdFilterField = nullptr;
    Label* _localizationStatusFilterLabel = nullptr;
    ComboBox* _localizationStatusFilterCombo = nullptr;

    Grid* _inventoryGrid       = nullptr;
    Grid* _translationGrid     = nullptr;
    Label* _sourceTextLabel    = nullptr;
    TextField* _sourceTextField = nullptr;
    Label* _targetTextLabel    = nullptr;
    TextField* _targetTextField = nullptr;
    Label* _validationLabel    = nullptr;
    Button* _exportRcButton    = nullptr;

    Label* _themeSelectorLabel = nullptr;
    ComboBox* _themeCombo      = nullptr;
    Label* _themeNameLabel  = nullptr;
    Label* _themePathLabel  = nullptr;
    Label* _themeErrorLabel = nullptr;

    Label* _colorKeyFilterLabel       = nullptr;
    TextField* _colorKeyFilterField   = nullptr;
    Label* _colorKeyLabel            = nullptr;
    ComboBox* _colorKeyCombo         = nullptr;
    Grid* _themeColorGrid            = nullptr;
    Label* _previousColorLabel       = nullptr;
    ColorSwatch* _previousColorSwatch = nullptr;
    Label* _colorValueLabel          = nullptr;
    ColorSwatch* _colorSwatch        = nullptr;
    ComboBox* _colorValueCombo       = nullptr;
    Label* _themeColorStatusLabel    = nullptr;
    Label* _themeExpressionHelpLabel = nullptr;
    Button* _darkenGroupButton       = nullptr;
    Button* _blendAccentGroupButton  = nullptr;
    Button* _resetColorButton        = nullptr;
    ScrollPanel* _themePreviewScroll = nullptr;
    ThemeExampleControl* _themePreview = nullptr;
    Button* _exportThemeButton       = nullptr;

    Label* _defaultRcPathLabel    = nullptr;
    Label* _defaultThemePathLabel = nullptr;
    Label* _rcPreviewLabel        = nullptr;
    TextField* _rcPreviewField    = nullptr;
    Label* _themeJsonPreviewLabel = nullptr;
    TextField* _themeJsonPreviewField = nullptr;
};


RedConfigureRootCreateResult CreateRedConfigureRoot(HINSTANCE instance, RedConfigureSession& session, std::filesystem::path initialRoot)
{
    auto root = std::make_unique<RedConfigureRoot>(instance, session, std::move(initialRoot));
    auto* controller = static_cast<RedConfigureRootController*>(root.get());
    std::unique_ptr<Panel> control = std::move(root);
    RedConfigureRootCreateResult result;
    result.control    = std::move(control);
    result.controller = controller;
    return result;
}

} // namespace RedConfigure::Ui
