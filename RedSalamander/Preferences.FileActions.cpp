#include "Framework.h"

#include "Preferences.FileActions.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cmath>
#include <cwctype>
#include <filesystem>
#include <format>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <vector>

#include "FileActionResolver.h"
#include "Helpers.h"
#include "UiMetrics.h"
#include "ViewerPluginManager.h"
#include "resource.h"

using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::Checkbox;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridColumnKind;
using RedSalamander::DxUi::GridRowStyle;
using RedSalamander::DxUi::GridRowTone;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::GridSortSpec;
using RedSalamander::DxUi::GridVisibleWorkMetrics;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::SortDirection;
using RedSalamander::DxUi::TabControl;
using RedSalamander::DxUi::TextField;

namespace
{
namespace Settings = Common::Settings;

constexpr std::wstring_view kComboMatchDefault   = L"default";
constexpr std::wstring_view kComboMatchExtension = L"extension";
constexpr std::wstring_view kComboMatchPattern   = L"pattern";
constexpr std::wstring_view kComboActionExternal = L"externalProgram";
constexpr std::wstring_view kComboActionViewer   = L"viewerPlugin";

[[nodiscard]] std::wstring LoadRes(const UINT id)
{
    return LoadStringResource(nullptr, id);
}

[[nodiscard]] wchar_t ToLower(const wchar_t ch) noexcept
{
    return static_cast<wchar_t>(std::towlower(ch));
}

[[nodiscard]] bool EqualsNoCase(std::wstring_view lhs, std::wstring_view rhs) noexcept
{
    if (lhs.size() != rhs.size())
    {
        return false;
    }

    for (size_t index = 0u; index < lhs.size(); ++index)
    {
        if (ToLower(lhs[index]) != ToLower(rhs[index]))
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] int CompareNoCase(std::wstring_view lhs, std::wstring_view rhs) noexcept
{
    const size_t commonSize = std::min(lhs.size(), rhs.size());
    for (size_t index = 0u; index < commonSize; ++index)
    {
        const wchar_t lhsCh = ToLower(lhs[index]);
        const wchar_t rhsCh = ToLower(rhs[index]);
        if (lhsCh < rhsCh)
        {
            return -1;
        }
        if (lhsCh > rhsCh)
        {
            return 1;
        }
    }
    if (lhs.size() == rhs.size())
    {
        return 0;
    }
    return lhs.size() < rhs.size() ? -1 : 1;
}

[[nodiscard]] std::wstring ToLowerString(std::wstring text)
{
    for (wchar_t& ch : text)
    {
        ch = ToLower(ch);
    }
    return text;
}

[[nodiscard]] std::wstring Trim(std::wstring_view value)
{
    size_t first = 0u;
    while (first < value.size() && std::iswspace(value[first]) != 0)
    {
        ++first;
    }

    size_t last = value.size();
    while (last > first && std::iswspace(value[last - 1u]) != 0)
    {
        --last;
    }

    return std::wstring(value.substr(first, last - first));
}

[[nodiscard]] std::wstring NormalizeExtensionText(std::wstring text)
{
    text = Trim(text);
    if (! text.empty() && text.front() != L'.')
    {
        text.insert(text.begin(), L'.');
    }
    return ToLowerString(std::move(text));
}

[[nodiscard]] std::vector<std::wstring> SplitList(std::wstring_view text)
{
    std::vector<std::wstring> result;
    size_t start = 0u;
    for (size_t index = 0u; index <= text.size(); ++index)
    {
        const bool atEnd     = index == text.size();
        const bool delimiter = ! atEnd && (text[index] == L',' || text[index] == L';' || text[index] == L'\r' || text[index] == L'\n' || text[index] == L'\t');
        if (! atEnd && ! delimiter)
        {
            continue;
        }

        std::wstring item = Trim(text.substr(start, index - start));
        if (! item.empty())
        {
            result.push_back(std::move(item));
        }
        start = index + 1u;
    }
    return result;
}

[[nodiscard]] std::wstring JoinStrings(std::span<const std::wstring> values)
{
    std::wstring text;
    for (const std::wstring& value : values)
    {
        if (value.empty())
        {
            continue;
        }
        if (! text.empty())
        {
            text.append(L", ");
        }
        text.append(value);
    }
    return text;
}

[[nodiscard]] uint64_t FnvAppend(uint64_t hash, std::wstring_view text) noexcept
{
    constexpr uint64_t kPrime = 1099511628211ull;
    for (const wchar_t ch : text)
    {
        hash ^= static_cast<uint64_t>(ToLower(ch));
        hash *= kPrime;
    }
    return hash;
}

[[nodiscard]] uint64_t StableHash(std::initializer_list<std::wstring_view> parts) noexcept
{
    uint64_t hash = 14695981039346656037ull;
    for (std::wstring_view part : parts)
    {
        hash = FnvAppend(hash, part);
        hash ^= 0xffu;
        hash *= 1099511628211ull;
    }
    return hash;
}

[[nodiscard]] std::wstring FileActionDisplayName(const Settings::FileActionDefinition& action)
{
    return action.displayName.empty() ? action.id : action.displayName;
}

[[nodiscard]] std::wstring ViewerPluginDisplayName(const ViewerPluginManager::PluginEntry& plugin)
{
    return plugin.name.empty() ? plugin.id : plugin.name;
}

[[nodiscard]] std::wstring MatchKindValue(const Settings::FileActionMatchKind kind)
{
    switch (kind)
    {
        case Settings::FileActionMatchKind::Default: return std::wstring(kComboMatchDefault);
        case Settings::FileActionMatchKind::Extension: return std::wstring(kComboMatchExtension);
        case Settings::FileActionMatchKind::Pattern: return std::wstring(kComboMatchPattern);
    }
    return std::wstring(kComboMatchDefault);
}

[[nodiscard]] Settings::FileActionMatchKind MatchKindFromValue(std::wstring_view value) noexcept
{
    if (value == kComboMatchExtension)
    {
        return Settings::FileActionMatchKind::Extension;
    }
    if (value == kComboMatchPattern)
    {
        return Settings::FileActionMatchKind::Pattern;
    }
    return Settings::FileActionMatchKind::Default;
}

[[nodiscard]] std::wstring ActionKindValue(const Settings::FileActionKind kind)
{
    switch (kind)
    {
        case Settings::FileActionKind::ViewerPlugin: return std::wstring(kComboActionViewer);
        case Settings::FileActionKind::ExternalProgram: return std::wstring(kComboActionExternal);
    }
    return std::wstring(kComboActionExternal);
}

[[nodiscard]] Settings::FileActionKind ActionKindFromValue(std::wstring_view value) noexcept
{
    if (value == kComboActionViewer)
    {
        return Settings::FileActionKind::ViewerPlugin;
    }
    return Settings::FileActionKind::ExternalProgram;
}

[[nodiscard]] std::wstring ActionKindDisplay(const Settings::FileActionKind kind)
{
    switch (kind)
    {
        case Settings::FileActionKind::ViewerPlugin: return LoadRes(IDS_PREFS_FILE_ACTION_TYPE_VIEWER_PLUGIN);
        case Settings::FileActionKind::ExternalProgram: return LoadRes(IDS_PREFS_FILE_ACTION_TYPE_EXTERNAL_PROGRAM);
    }
    return LoadRes(IDS_PREFS_FILE_ACTION_TYPE_EXTERNAL_PROGRAM);
}

[[nodiscard]] std::wstring MatchDisplay(const Settings::FileActionMatch& match)
{
    if (match.kind == Settings::FileActionMatchKind::Default)
    {
        return L"*";
    }
    return match.value;
}

[[nodiscard]] std::wstring MatchInputDisplay(const Settings::FileActionMatch& match)
{
    if (match.kind == Settings::FileActionMatchKind::Default)
    {
        return {};
    }
    return match.value;
}

[[nodiscard]] Settings::FileActionMatch BuildMatch(Settings::FileActionMatchKind kind, std::wstring_view value)
{
    Settings::FileActionMatch match{};
    match.kind = kind;
    if (kind == Settings::FileActionMatchKind::Extension)
    {
        match.value = NormalizeExtensionText(std::wstring(value));
    }
    else if (kind == Settings::FileActionMatchKind::Pattern)
    {
        match.value = Trim(value);
    }
    return match;
}

[[nodiscard]] bool MatchesSameKey(const Settings::FileActionMatch& lhs, const Settings::FileActionMatch& rhs) noexcept
{
    return lhs.kind == rhs.kind && EqualsNoCase(lhs.value, rhs.value);
}

[[nodiscard]] std::wstring FormatMatches(std::span<const Settings::FileActionMatch> matches)
{
    if (matches.empty())
    {
        return L"*";
    }

    std::wstring text;
    for (const Settings::FileActionMatch& match : matches)
    {
        if (! text.empty())
        {
            text.append(L", ");
        }
        text.append(MatchDisplay(match));
    }
    return text.empty() ? L"*" : text;
}

[[nodiscard]] std::vector<Settings::FileActionMatch> ParseMatchesField(std::wstring_view text)
{
    std::vector<Settings::FileActionMatch> matches;
    const std::vector<std::wstring> tokens = SplitList(text);
    for (const std::wstring& token : tokens)
    {
        if (token == L"*")
        {
            matches.push_back(Settings::FileActionMatch{.kind = Settings::FileActionMatchKind::Default});
        }
        else if (! token.empty() && token.front() == L'.')
        {
            matches.push_back(Settings::FileActionMatch{.kind = Settings::FileActionMatchKind::Extension, .value = NormalizeExtensionText(token)});
        }
        else
        {
            matches.push_back(Settings::FileActionMatch{.kind = Settings::FileActionMatchKind::Pattern, .value = Trim(token)});
        }
    }

    if (matches.empty())
    {
        matches.push_back(Settings::FileActionMatch{.kind = Settings::FileActionMatchKind::Default});
    }
    return matches;
}

[[nodiscard]] const Settings::FileActionDefinition* FindActionById(const std::vector<Settings::FileActionDefinition>& actions,
                                                                   std::wstring_view actionId) noexcept
{
    if (actionId.empty())
    {
        return nullptr;
    }

    const auto it =
        std::find_if(actions.begin(), actions.end(), [&](const Settings::FileActionDefinition& action) noexcept { return EqualsNoCase(action.id, actionId); });
    return it == actions.end() ? nullptr : &(*it);
}

[[nodiscard]] std::wstring ActionDisplay(const std::vector<Settings::FileActionDefinition>& actions, std::wstring_view actionId)
{
    if (actionId.empty())
    {
        return LoadRes(IDS_PREFS_FILE_ACTION_NONE);
    }

    if (const Settings::FileActionDefinition* action = FindActionById(actions, actionId))
    {
        return FileActionDisplayName(*action);
    }

    return FormatStringResource(nullptr, IDS_PREFS_FILE_ACTION_MISSING_FMT, std::wstring(actionId));
}

[[nodiscard]] std::wstring ActionStatus(const std::vector<Settings::FileActionDefinition>& actions, std::initializer_list<std::wstring_view> actionIds)
{
    for (std::wstring_view actionId : actionIds)
    {
        if (actionId.empty())
        {
            continue;
        }

        const Settings::FileActionDefinition* action = FindActionById(actions, actionId);
        if (! action)
        {
            return FormatStringResource(nullptr, IDS_PREFS_FILE_ACTION_MISSING_FMT, std::wstring(actionId));
        }
        if (! action->enabled)
        {
            return LoadRes(IDS_PREFS_FILE_ACTION_DISABLED);
        }
    }
    return LoadRes(IDS_PREFS_FILE_ACTION_ENABLED);
}

[[nodiscard]] std::wstring CurrentComputerName()
{
    std::array<wchar_t, MAX_COMPUTERNAME_LENGTH + 1u> buffer{};
    DWORD size = static_cast<DWORD>(buffer.size());
    if (GetComputerNameW(buffer.data(), &size) == FALSE)
    {
        return {};
    }
    return std::wstring(buffer.data(), size);
}

[[nodiscard]] std::filesystem::path PreviewPathFromField(const TextField* field)
{
    if (field)
    {
        std::wstring text = Trim(field->GetText());
        if (! text.empty())
        {
            return std::filesystem::path(text);
        }
    }
    return std::filesystem::path(LoadRes(IDS_PREFS_FILE_ACTION_TEST_FILE_DEFAULT));
}

[[nodiscard]] std::wstring CommandDisplay(const FileActionResolver::Command command)
{
    switch (command)
    {
        case FileActionResolver::Command::View: return LoadRes(IDS_PREFS_FILE_ACTION_COL_VIEW);
        case FileActionResolver::Command::AlternateView: return LoadRes(IDS_PREFS_FILE_ACTION_COL_ALTERNATE_VIEW);
        case FileActionResolver::Command::Edit: return LoadRes(IDS_PREFS_FILE_ACTION_COL_EDIT);
        case FileActionResolver::Command::AlternateEdit: return LoadRes(IDS_PREFS_FILE_ACTION_COL_ALTERNATE_EDIT);
        case FileActionResolver::Command::EditNew: return LoadRes(IDS_PREFS_FILE_ACTION_COL_EDIT_NEW);
    }
    return {};
}

struct FileActionGridRow
{
    size_t sourceIndex = 0u;
    uint64_t stableId  = 0u;
    std::vector<std::wstring> cells;
    GridRowTone tone = GridRowTone::None;
};

} // namespace

class FileActionGridModel final : public IDxGridModel
{
public:
    void SetColumns(std::vector<GridColumnDesc> columns)
    {
        _columns = std::move(columns);
    }

    void SetRows(std::vector<FileActionGridRow> rows)
    {
        _rows = std::move(rows);
    }

    void SortRows(const GridSortSpec& sortSpec)
    {
        if (_rows.empty())
        {
            return;
        }

        const bool sortBySourceOrder =
            sortSpec.direction == SortDirection::None || sortSpec.columnIndex >= _columns.size() || ! _columns[sortSpec.columnIndex].sortable;
        if (sortBySourceOrder)
        {
            std::stable_sort(_rows.begin(),
                             _rows.end(),
                             [](const FileActionGridRow& lhs, const FileActionGridRow& rhs) noexcept
            {
                if (lhs.sourceIndex != rhs.sourceIndex)
                {
                    return lhs.sourceIndex < rhs.sourceIndex;
                }
                return lhs.stableId < rhs.stableId;
            });
            return;
        }

        const auto cellText = [&](const FileActionGridRow& row, const size_t columnIndex) noexcept -> std::wstring_view
        { return columnIndex < row.cells.size() ? std::wstring_view(row.cells[columnIndex]) : std::wstring_view{}; };

        std::stable_sort(_rows.begin(),
                         _rows.end(),
                         [&](const FileActionGridRow& lhs, const FileActionGridRow& rhs) noexcept
        {
            int comparison = CompareNoCase(cellText(lhs, sortSpec.columnIndex), cellText(rhs, sortSpec.columnIndex));
            if (comparison == 0 && sortSpec.columnIndex != 0u)
            {
                comparison = CompareNoCase(cellText(lhs, 0u), cellText(rhs, 0u));
            }
            if (comparison == 0 && lhs.sourceIndex != rhs.sourceIndex)
            {
                comparison = lhs.sourceIndex < rhs.sourceIndex ? -1 : 1;
            }
            if (comparison == 0)
            {
                return lhs.stableId < rhs.stableId;
            }
            return sortSpec.direction == SortDirection::Ascending ? comparison < 0 : comparison > 0;
        });
    }

    [[nodiscard]] const FileActionGridRow* GetRow(const size_t rowIndex) const noexcept
    {
        return rowIndex < _rows.size() ? &_rows[rowIndex] : nullptr;
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rows.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(const size_t columnIndex) const override
    {
        if (columnIndex >= _columns.size())
        {
            return {};
        }
        return _columns[columnIndex];
    }

    void GetCellData(const size_t rowIndex, const size_t columnIndex, GridCellData& outCell) const override
    {
        outCell                      = {};
        const FileActionGridRow* row = GetRow(rowIndex);
        if (! row || columnIndex >= row->cells.size())
        {
            return;
        }
        outCell.text        = row->cells[columnIndex];
        outCell.tooltipText = outCell.text;
        outCell.multiline   = false;
    }

    [[nodiscard]] GridRowStyle GetRowStyle(const size_t rowIndex) const override
    {
        GridRowStyle style{};
        if (const FileActionGridRow* row = GetRow(rowIndex))
        {
            style.tone = row->tone;
        }
        return style;
    }

    [[nodiscard]] uint64_t GetStableRowId(const size_t rowIndex) const noexcept override
    {
        const FileActionGridRow* row = GetRow(rowIndex);
        return row ? row->stableId : 0u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(const uint64_t rowId) const noexcept override
    {
        for (size_t index = 0u; index < _rows.size(); ++index)
        {
            if (_rows[index].stableId == rowId)
            {
                return index;
            }
        }
        return std::nullopt;
    }

private:
    std::vector<GridColumnDesc> _columns;
    std::vector<FileActionGridRow> _rows;
};

namespace
{
[[nodiscard]] bool TextContainsFilter(const FileActionGridRow& row, std::wstring_view filter) noexcept
{
    if (filter.empty())
    {
        return true;
    }

    for (const std::wstring& cell : row.cells)
    {
        if (cell.empty() || cell.size() < filter.size())
        {
            continue;
        }

        for (size_t index = 0u; index + filter.size() <= cell.size(); ++index)
        {
            bool match = true;
            for (size_t offset = 0u; offset < filter.size(); ++offset)
            {
                if (ToLower(cell[index + offset]) != ToLower(filter[offset]))
                {
                    match = false;
                    break;
                }
            }
            if (match)
            {
                return true;
            }
        }
    }
    return false;
}

[[nodiscard]] std::vector<GridColumnDesc> ViewerAssociationColumns()
{
    return {
        GridColumnDesc{.id          = L"match",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_MATCH),
                       .widthDip    = 96.0f,
                       .minWidthDip = 64.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"computer",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_COMPUTER),
                       .widthDip    = 110.0f,
                       .minWidthDip = 80.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"view",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_VIEW),
                       .widthDip    = 180.0f,
                       .minWidthDip = 120.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"alternateView",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_ALTERNATE_VIEW),
                       .widthDip    = 210.0f,
                       .minWidthDip = 130.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"status",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_STATUS),
                       .widthDip    = 120.0f,
                       .minWidthDip = 90.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
    };
}

[[nodiscard]] std::vector<GridColumnDesc> EditorAssociationColumns()
{
    return {
        GridColumnDesc{.id          = L"match",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_MATCH),
                       .widthDip    = 88.0f,
                       .minWidthDip = 64.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"computer",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_COMPUTER),
                       .widthDip    = 96.0f,
                       .minWidthDip = 80.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"edit",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_EDIT),
                       .widthDip    = 140.0f,
                       .minWidthDip = 100.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"alternateEdit",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_ALTERNATE_EDIT),
                       .widthDip    = 210.0f,
                       .minWidthDip = 130.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"editNew",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_EDIT_NEW),
                       .widthDip    = 170.0f,
                       .minWidthDip = 120.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"status",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_STATUS),
                       .widthDip    = 120.0f,
                       .minWidthDip = 90.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
    };
}

[[nodiscard]] std::vector<GridColumnDesc> ActionColumns()
{
    return {
        GridColumnDesc{.id          = L"name",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_NAME),
                       .widthDip    = 180.0f,
                       .minWidthDip = 120.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"type",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_TYPE),
                       .widthDip    = 140.0f,
                       .minWidthDip = 100.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"appliesTo",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_APPLIES_TO),
                       .widthDip    = 180.0f,
                       .minWidthDip = 120.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"computer",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_COMPUTER),
                       .widthDip    = 150.0f,
                       .minWidthDip = 100.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
        GridColumnDesc{.id          = L"status",
                       .title       = LoadRes(IDS_PREFS_FILE_ACTION_COL_STATUS),
                       .widthDip    = 110.0f,
                       .minWidthDip = 80.0f,
                       .kind        = GridColumnKind::Text,
                       .sortable    = true,
                       .multiline   = false},
    };
}

[[nodiscard]] std::vector<FileActionGridRow> BuildViewerAssociationRows(const Settings::ViewerFileActionsSettings& settings, std::wstring_view filter)
{
    std::vector<FileActionGridRow> rows;
    rows.reserve(settings.associations.size());
    for (size_t index = 0u; index < settings.associations.size(); ++index)
    {
        const Settings::ViewerAssociationRule& rule = settings.associations[index];
        FileActionGridRow row{};
        row.sourceIndex = index;
        row.stableId    = StableHash({L"viewer-association", MatchKindValue(rule.match.kind), rule.match.value, rule.computerName});
        row.cells       = {
            MatchDisplay(rule.match),
            rule.computerName.empty() ? LoadRes(IDS_PREFS_FILE_ACTION_ANY) : rule.computerName,
            ActionDisplay(settings.actions, rule.viewActionId),
            ActionDisplay(settings.actions, rule.alternateViewActionId),
            ActionStatus(settings.actions, {rule.viewActionId, rule.alternateViewActionId}),
        };
        if (row.cells.back() != LoadRes(IDS_PREFS_FILE_ACTION_ENABLED))
        {
            row.tone = GridRowTone::Warning;
        }
        if (TextContainsFilter(row, filter))
        {
            rows.push_back(std::move(row));
        }
    }
    return rows;
}

[[nodiscard]] std::vector<FileActionGridRow> BuildEditorAssociationRows(const Settings::EditorFileActionsSettings& settings, std::wstring_view filter)
{
    std::vector<FileActionGridRow> rows;
    rows.reserve(settings.associations.size());
    for (size_t index = 0u; index < settings.associations.size(); ++index)
    {
        const Settings::EditorAssociationRule& rule = settings.associations[index];
        FileActionGridRow row{};
        row.sourceIndex = index;
        row.stableId    = StableHash({L"editor-association", MatchKindValue(rule.match.kind), rule.match.value, rule.computerName});
        row.cells       = {
            MatchDisplay(rule.match),
            rule.computerName.empty() ? LoadRes(IDS_PREFS_FILE_ACTION_ANY) : rule.computerName,
            ActionDisplay(settings.actions, rule.editActionId),
            ActionDisplay(settings.actions, rule.alternateEditActionId),
            ActionDisplay(settings.actions, rule.editNewActionId),
            ActionStatus(settings.actions, {rule.editActionId, rule.alternateEditActionId, rule.editNewActionId}),
        };
        if (row.cells.back() != LoadRes(IDS_PREFS_FILE_ACTION_ENABLED))
        {
            row.tone = GridRowTone::Warning;
        }
        if (TextContainsFilter(row, filter))
        {
            rows.push_back(std::move(row));
        }
    }
    return rows;
}

[[nodiscard]] std::vector<FileActionGridRow> BuildActionRows(const std::vector<Settings::FileActionDefinition>& actions,
                                                             std::wstring_view filter,
                                                             std::wstring_view familyText)
{
    std::vector<FileActionGridRow> rows;
    rows.reserve(actions.size());
    for (size_t index = 0u; index < actions.size(); ++index)
    {
        const Settings::FileActionDefinition& action = actions[index];
        FileActionGridRow row{};
        row.sourceIndex = index;
        row.stableId    = StableHash({familyText, L"action", action.id});
        row.cells       = {
            FileActionDisplayName(action),
            ActionKindDisplay(action.kind),
            FormatMatches(action.appliesTo.matches),
            action.appliesTo.computerNames.empty() ? LoadRes(IDS_PREFS_FILE_ACTION_ANY) : JoinStrings(action.appliesTo.computerNames),
            action.enabled ? LoadRes(IDS_PREFS_FILE_ACTION_ENABLED) : LoadRes(IDS_PREFS_FILE_ACTION_DISABLED),
        };
        if (! action.enabled)
        {
            row.tone = GridRowTone::Info;
        }
        if (TextContainsFilter(row, filter))
        {
            rows.push_back(std::move(row));
        }
    }
    return rows;
}

template <typename Rule>
[[nodiscard]] std::optional<size_t> FindAssociationByKey(const std::vector<Rule>& rules,
                                                         const Settings::FileActionMatch& match,
                                                         std::wstring_view computerName) noexcept
{
    for (size_t index = 0u; index < rules.size(); ++index)
    {
        if (MatchesSameKey(rules[index].match, match) && EqualsNoCase(rules[index].computerName, computerName))
        {
            return index;
        }
    }
    return std::nullopt;
}

[[nodiscard]] std::optional<size_t> FindActionIndexById(const std::vector<Settings::FileActionDefinition>& actions, std::wstring_view actionId) noexcept
{
    if (actionId.empty())
    {
        return std::nullopt;
    }
    for (size_t index = 0u; index < actions.size(); ++index)
    {
        if (EqualsNoCase(actions[index].id, actionId))
        {
            return index;
        }
    }
    return std::nullopt;
}

[[nodiscard]] Settings::ViewerAssociationRule& EnsureDefaultViewerAssociation(Settings::ViewerFileActionsSettings& settings)
{
    const auto existing =
        FindAssociationByKey(settings.associations, Settings::FileActionMatch{.kind = Settings::FileActionMatchKind::Default}, std::wstring_view{});
    if (existing.has_value())
    {
        return settings.associations[existing.value()];
    }

    Settings::ViewerAssociationRule rule{};
    rule.match.kind = Settings::FileActionMatchKind::Default;
    settings.associations.push_back(std::move(rule));
    return settings.associations.back();
}

[[nodiscard]] Settings::EditorAssociationRule& EnsureDefaultEditorAssociation(Settings::EditorFileActionsSettings& settings)
{
    const auto existing =
        FindAssociationByKey(settings.associations, Settings::FileActionMatch{.kind = Settings::FileActionMatchKind::Default}, std::wstring_view{});
    if (existing.has_value())
    {
        return settings.associations[existing.value()];
    }

    Settings::EditorAssociationRule rule{};
    rule.match.kind = Settings::FileActionMatchKind::Default;
    settings.associations.push_back(std::move(rule));
    return settings.associations.back();
}

[[nodiscard]] ComboBox::Item MakeComboItem(const std::wstring& value, const std::wstring& display)
{
    return ComboBox::Item{.value = value, .display = display};
}

[[nodiscard]] std::optional<size_t> FindComboItem(std::span<const ComboBox::Item> items, std::wstring_view value) noexcept
{
    for (size_t index = 0u; index < items.size(); ++index)
    {
        if (EqualsNoCase(items[index].value, value))
        {
            return index;
        }
    }
    return std::nullopt;
}

void SelectComboValue(ComboBox* combo, std::wstring_view value) noexcept
{
    if (! combo)
    {
        return;
    }

    combo->SetSelectedIndex(FindComboItem(combo->GetItems(), value).value_or(0u));
}

void RebuildViewerPluginOptions(PreferencesDialogState& state)
{
    state.viewersPluginOptions.clear();

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    state.viewersPluginOptions.reserve(plugins.size());
    for (const ViewerPluginManager::PluginEntry& plugin : plugins)
    {
        if (! plugin.id.empty())
        {
            state.viewersPluginOptions.push_back(ViewerPluginOption{.id = plugin.id, .displayName = ViewerPluginDisplayName(plugin)});
        }
    }
}

[[nodiscard]] std::vector<ComboBox::Item> BuildViewerPluginComboItems(PreferencesDialogState& state, const std::vector<Settings::FileActionDefinition>& actions)
{
    RebuildViewerPluginOptions(state);

    std::vector<ComboBox::Item> items;
    items.reserve(state.viewersPluginOptions.size() + actions.size() + 1u);
    items.push_back(MakeComboItem({}, LoadRes(IDS_PREFS_FILE_ACTION_NONE)));

    for (const ViewerPluginOption& option : state.viewersPluginOptions)
    {
        if (! option.id.empty())
        {
            items.push_back(MakeComboItem(option.id, option.displayName.empty() ? option.id : option.displayName));
        }
    }

    for (const Settings::FileActionDefinition& action : actions)
    {
        if (action.kind != Settings::FileActionKind::ViewerPlugin || action.pluginId.empty() || FindComboItem(items, action.pluginId).has_value())
        {
            continue;
        }

        items.push_back(MakeComboItem(action.pluginId, FormatStringResource(nullptr, IDS_PREFS_FILE_ACTION_MISSING_FMT, action.pluginId)));
    }

    return items;
}

void SetTextNoNotify(TextField* field, std::wstring text) noexcept
{
    if (field)
    {
        field->SetText(std::move(text));
    }
}

void SetLabelBounds(Label* label, const float left, const float top, const float right, const float bottom) noexcept
{
    if (label)
    {
        label->SetBounds(D2D1::RectF(left, top, right, bottom));
    }
}

template <typename TControl> void SetControlBounds(TControl* control, const float left, const float top, const float right, const float bottom) noexcept
{
    if (control)
    {
        control->SetBounds(D2D1::RectF(left, top, right, bottom));
    }
}

#ifdef ENABLE_TESTS
[[nodiscard]] RECT DipRectToPx(const RedSalamander::DxUi::WindowHost& host, const D2D1_RECT_F& rect) noexcept
{
    RECT out{};
    out.left   = static_cast<LONG>(std::lround(host.DipsToPixels(rect.left)));
    out.top    = static_cast<LONG>(std::lround(host.DipsToPixels(rect.top)));
    out.right  = static_cast<LONG>(std::lround(host.DipsToPixels(rect.right)));
    out.bottom = static_cast<LONG>(std::lround(host.DipsToPixels(rect.bottom)));
    return out;
}
#endif
} // namespace

FileActionPreferencesPage::FileActionPreferencesPage(const FileActionPreferencesFamily family) noexcept : _family(family)
{
}

FileActionPreferencesPage::~FileActionPreferencesPage()
{
    DetachDxPageHost();
}

bool FileActionPreferencesPage::IsViewerFamily() const noexcept
{
    return _family == FileActionPreferencesFamily::Viewers;
}

bool FileActionPreferencesPage::IsEditorsFamily() const noexcept
{
    return _family == FileActionPreferencesFamily::Editors;
}

PrefCategory FileActionPreferencesPage::Category() const noexcept
{
    return IsViewerFamily() ? PrefCategory::Viewers : PrefCategory::Editors;
}

const wchar_t* FileActionPreferencesPage::MetricFamilyText() const noexcept
{
    return IsViewerFamily() ? L"viewers" : L"editors";
}

void FileActionPreferencesPage::OnVisibilityChanged(const bool visible) noexcept
{
    if (! visible && _pageHost)
    {
        _pageHost->ResetInteractionState();
    }
}

void FileActionPreferencesPage::Destroy(PreferencesDialogState& state) noexcept
{
    if (IsViewerFamily())
    {
        state.viewersExtensionKeys.clear();
        state.viewersPluginOptions.clear();
    }
    DetachDxPageHost();
}

void FileActionPreferencesPage::ResetControlPointers() noexcept
{
    _tabs                    = nullptr;
    _associationsPage        = nullptr;
    _actionsPage             = nullptr;
    _searchLabel             = nullptr;
    _searchField             = nullptr;
    _associationsGrid        = nullptr;
    _matchKindLabel          = nullptr;
    _matchKindCombo          = nullptr;
    _matchValueLabel         = nullptr;
    _matchValueField         = nullptr;
    _computerLabel           = nullptr;
    _computerField           = nullptr;
    _primaryActionLabel      = nullptr;
    _primaryActionCombo      = nullptr;
    _alternateActionLabel    = nullptr;
    _alternateActionCombo    = nullptr;
    _editNewActionLabel      = nullptr;
    _editNewActionCombo      = nullptr;
    _testFileLabel           = nullptr;
    _testFileField           = nullptr;
    _previewLabel            = nullptr;
    _associationSaveButton   = nullptr;
    _associationRemoveButton = nullptr;
    _associationResetButton  = nullptr;
    _actionsGrid             = nullptr;
    _actionIdLabel           = nullptr;
    _actionIdField           = nullptr;
    _actionNameLabel         = nullptr;
    _actionNameField         = nullptr;
    _actionKindLabel         = nullptr;
    _actionKindCombo         = nullptr;
    _actionEnabledCheckbox   = nullptr;
    _pluginIdLabel           = nullptr;
    _pluginIdCombo           = nullptr;
    _executableLabel         = nullptr;
    _executableField         = nullptr;
    _argumentsLabel          = nullptr;
    _argumentsField          = nullptr;
    _workingDirectoryLabel   = nullptr;
    _workingDirectoryField   = nullptr;
    _appliesToLabel          = nullptr;
    _appliesToField          = nullptr;
    _computersLabel          = nullptr;
    _computersField          = nullptr;
    _actionSaveButton        = nullptr;
    _actionRemoveButton      = nullptr;
    _associationsModel.reset();
    _actionsModel.reset();
}

void FileActionPreferencesPage::DetachDxPageHost() noexcept
{
    if (_associationsGrid)
    {
        _associationsGrid->SetModel(nullptr);
        _associationsGrid->SetDelegate(nullptr);
    }
    if (_actionsGrid)
    {
        _actionsGrid->SetModel(nullptr);
        _actionsGrid->SetDelegate(nullptr);
    }
    if (_pageContentRoot && _pageHost)
    {
        _pageHost->ResetInteractionState();
        _pageContentRoot->ClearChildren();
    }

    ResetControlPointers();
    _pageHost        = nullptr;
    _pageContentRoot = nullptr;
    _state           = nullptr;
    _hostWindow      = nullptr;
    _activeGrid      = ActiveGrid::Actions;
    _previewActionId.clear();
    _previewReason.clear();
    _usesDxUiTypographyContext = false;
    _usesDxUiTypographyMetrics = false;
    _syncing                   = false;
}

void FileActionPreferencesPage::ApplyTheme(const PreferencesDialogState& state) noexcept
{
    if (_pageHost)
    {
        _pageHost->SetTheme(PrefsUi::MakeDxPalette(state.theme));
    }
}

bool FileActionPreferencesPage::EnsureDxPageHost(HWND parent, PreferencesDialogState& state) noexcept
{
    static_cast<void>(parent);

    _pageHost        = state.pageHostDxHost;
    _pageContentRoot = state.pageHostDxContentRootControl;
    if (! _pageHost || ! _pageContentRoot)
    {
        return false;
    }

    bool ownsCurrentChildren = false;
    if (_tabs)
    {
        for (const std::unique_ptr<RedSalamander::DxUi::Control>& child : _pageContentRoot->GetChildren())
        {
            if (child.get() == _tabs)
            {
                ownsCurrentChildren = true;
                break;
            }
        }
    }
    if (ownsCurrentChildren)
    {
        return true;
    }

    _pageHost->ResetInteractionState();
    _pageContentRoot->ClearChildren();
    ResetControlPointers();

    _tabs             = _pageContentRoot->AddChild<TabControl>();
    _actionsPage      = _tabs->AddTab<Panel>(LoadRes(IDS_PREFS_FILE_ACTION_TAB_ACTIONS));
    _associationsPage = _tabs->AddTab<Panel>(LoadRes(IDS_PREFS_FILE_ACTION_TAB_ASSOCIATIONS));
    _activeGrid       = ActiveGrid::Actions;
    _tabs->SetOnSelectionChanged([this](const size_t index) noexcept
    {
        Debug::Perf::Scope callbackPerf(L"preferences.ui.fileactions.tab_switch_callback_us");
        Debug::Perf::Emit(L"preferences.ui.fileactions.tab_switch",
                          MetricFamilyText(),
                          0u,
                          static_cast<uint64_t>(index),
                          _tabs && _tabs->GetSelectedPage() == _actionsPage ? 1u : 0u,
                          S_OK);
        _activeGrid = (_tabs && _tabs->GetSelectedPage() == _actionsPage) ? ActiveGrid::Actions : ActiveGrid::Associations;
        callbackPerf.SetValue0(static_cast<uint64_t>(index));
        callbackPerf.SetValue1(_activeGrid == ActiveGrid::Actions ? 1u : 0u);
    });

    _searchLabel          = _associationsPage->AddChild<Label>();
    _searchField          = _associationsPage->AddChild<TextField>();
    _associationsGrid     = _associationsPage->AddChild<Grid>();
    _matchKindLabel       = _associationsPage->AddChild<Label>();
    _matchKindCombo       = _associationsPage->AddChild<ComboBox>();
    _matchValueLabel      = _associationsPage->AddChild<Label>();
    _matchValueField      = _associationsPage->AddChild<TextField>();
    _computerLabel        = _associationsPage->AddChild<Label>();
    _computerField        = _associationsPage->AddChild<TextField>();
    _primaryActionLabel   = _associationsPage->AddChild<Label>();
    _primaryActionCombo   = _associationsPage->AddChild<ComboBox>();
    _alternateActionLabel = _associationsPage->AddChild<Label>();
    _alternateActionCombo = _associationsPage->AddChild<ComboBox>();
    if (IsEditorsFamily())
    {
        _editNewActionLabel = _associationsPage->AddChild<Label>();
        _editNewActionCombo = _associationsPage->AddChild<ComboBox>();
    }
    _testFileLabel           = _associationsPage->AddChild<Label>();
    _testFileField           = _associationsPage->AddChild<TextField>(LoadRes(IDS_PREFS_FILE_ACTION_TEST_FILE_DEFAULT));
    _previewLabel            = _associationsPage->AddChild<Label>();
    _associationSaveButton   = _associationsPage->AddChild<Button>(LoadRes(IDS_PREFS_FILE_ACTION_BUTTON_SAVE_ASSOCIATION));
    _associationRemoveButton = _associationsPage->AddChild<Button>(LoadRes(IDS_PREFS_FILE_ACTION_BUTTON_REMOVE));
    _associationResetButton  = _associationsPage->AddChild<Button>(LoadRes(IDS_PREFS_FILE_ACTION_BUTTON_RESET_DEFAULTS));

    _actionsGrid           = _actionsPage->AddChild<Grid>();
    _actionIdLabel         = _actionsPage->AddChild<Label>();
    _actionIdField         = _actionsPage->AddChild<TextField>();
    _actionNameLabel       = _actionsPage->AddChild<Label>();
    _actionNameField       = _actionsPage->AddChild<TextField>();
    _actionKindLabel       = _actionsPage->AddChild<Label>();
    _actionKindCombo       = _actionsPage->AddChild<ComboBox>();
    _actionEnabledCheckbox = _actionsPage->AddChild<Checkbox>(LoadRes(IDS_PREFS_FILE_ACTION_CHECK_ENABLED));
    if (IsViewerFamily())
    {
        _pluginIdLabel = _actionsPage->AddChild<Label>();
        _pluginIdCombo = _actionsPage->AddChild<ComboBox>();
    }
    _executableLabel       = _actionsPage->AddChild<Label>();
    _executableField       = _actionsPage->AddChild<TextField>();
    _argumentsLabel        = _actionsPage->AddChild<Label>();
    _argumentsField        = _actionsPage->AddChild<TextField>();
    _workingDirectoryLabel = _actionsPage->AddChild<Label>();
    _workingDirectoryField = _actionsPage->AddChild<TextField>();
    _appliesToLabel        = _actionsPage->AddChild<Label>();
    _appliesToField        = _actionsPage->AddChild<TextField>();
    _computersLabel        = _actionsPage->AddChild<Label>();
    _computersField        = _actionsPage->AddChild<TextField>();
    _actionSaveButton      = _actionsPage->AddChild<Button>(LoadRes(IDS_PREFS_FILE_ACTION_BUTTON_SAVE_ACTION));
    _actionRemoveButton    = _actionsPage->AddChild<Button>(LoadRes(IDS_PREFS_FILE_ACTION_BUTTON_REMOVE));

    _associationsModel = std::make_unique<FileActionGridModel>();
    _actionsModel      = std::make_unique<FileActionGridModel>();
    _associationsGrid->SetModel(_associationsModel.get());
    _associationsGrid->SetDelegate(this);
    _associationsGrid->SetSelectionMode(GridSelectionMode::Single);
    _associationsGrid->SetHeaderHeightDip(30.0f);
    _associationsGrid->SetRowHeightDip(30.0f);
    _associationsGrid->SetLineClamp(1u);
    _associationsGrid->SetEmptyStateText(LoadRes(IDS_PREFS_FILE_ACTION_NO_ASSOCIATIONS));
    _actionsGrid->SetModel(_actionsModel.get());
    _actionsGrid->SetDelegate(this);
    _actionsGrid->SetSelectionMode(GridSelectionMode::Single);
    _actionsGrid->SetHeaderHeightDip(30.0f);
    _actionsGrid->SetRowHeightDip(30.0f);
    _actionsGrid->SetLineClamp(1u);
    _actionsGrid->SetEmptyStateText(LoadRes(IDS_PREFS_FILE_ACTION_NO_ACTIONS));

    _matchKindCombo->SetVariant(ComboBoxVariant::Window);
    _primaryActionCombo->SetVariant(ComboBoxVariant::Window);
    _alternateActionCombo->SetVariant(ComboBoxVariant::Window);
    if (_editNewActionCombo)
    {
        _editNewActionCombo->SetVariant(ComboBoxVariant::Window);
    }
    _actionKindCombo->SetVariant(ComboBoxVariant::Window);
    if (_pluginIdCombo)
    {
        _pluginIdCombo->SetVariant(ComboBoxVariant::Window);
    }
    _previewLabel->SetMultiline(true);
    _previewLabel->SetFontRole(RedSalamander::DxUi::FontRole::Small);

    _searchField->SetOnTextChanged([this](std::wstring_view text) noexcept
    {
        if (_syncing || ! _state)
        {
            return;
        }
        OnSearchChanged(*_state, text);
    });
    _testFileField->SetOnTextChanged([this](std::wstring_view) noexcept
    {
        if (! _syncing && _state)
        {
            UpdatePreview(*_state);
        }
    });
    _actionKindCombo->SetOnSelectionChanged([this](const size_t) noexcept { SyncActionFieldAvailability(); });
    _associationSaveButton->SetPrimary(true);
    _associationSaveButton->SetOnClick([this]() noexcept
    {
        if (_state)
        {
            SaveAssociation(*_state);
        }
    });
    _associationRemoveButton->SetOnClick([this]() noexcept
    {
        if (_state)
        {
            RemoveSelectedAssociation(*_state);
        }
    });
    _associationResetButton->SetOnClick([this]() noexcept
    {
        if (_state)
        {
            ResetAssociationsAndActions(*_state);
        }
    });
    _actionSaveButton->SetPrimary(true);
    _actionSaveButton->SetOnClick([this]() noexcept
    {
        if (_state)
        {
            SaveAction(*_state);
        }
    });
    _actionRemoveButton->SetOnClick([this]() noexcept
    {
        if (_state)
        {
            RemoveSelectedAction(*_state);
        }
    });

    SyncStaticText();
    ApplyTheme(state);
    SyncFromState(state);
    return true;
}

void FileActionPreferencesPage::SyncStaticText() noexcept
{
    if (_searchLabel)
    {
        _searchLabel->SetText(LoadRes(IDS_PREFS_COMMON_SEARCH));
        _searchLabel->SetMnemonicTarget(_searchField);
    }
    if (_matchKindLabel)
    {
        _matchKindLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_MATCH_KIND));
        _matchKindLabel->SetMnemonicTarget(_matchKindCombo);
    }
    if (_matchValueLabel)
    {
        _matchValueLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_MATCH_VALUE));
        _matchValueLabel->SetMnemonicTarget(_matchValueField);
    }
    if (_computerLabel)
    {
        _computerLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_COMPUTER));
        _computerLabel->SetMnemonicTarget(_computerField);
    }
    if (_primaryActionLabel)
    {
        _primaryActionLabel->SetText(IsViewerFamily() ? LoadRes(IDS_PREFS_FILE_ACTION_COL_VIEW) : LoadRes(IDS_PREFS_FILE_ACTION_COL_EDIT));
        _primaryActionLabel->SetMnemonicTarget(_primaryActionCombo);
    }
    if (_alternateActionLabel)
    {
        _alternateActionLabel->SetText(IsViewerFamily() ? LoadRes(IDS_PREFS_FILE_ACTION_COL_ALTERNATE_VIEW)
                                                        : LoadRes(IDS_PREFS_FILE_ACTION_COL_ALTERNATE_EDIT));
        _alternateActionLabel->SetMnemonicTarget(_alternateActionCombo);
    }
    if (_editNewActionLabel)
    {
        _editNewActionLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_COL_EDIT_NEW));
        _editNewActionLabel->SetMnemonicTarget(_editNewActionCombo);
    }
    if (_testFileLabel)
    {
        _testFileLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_TEST_FILE));
        _testFileLabel->SetMnemonicTarget(_testFileField);
    }
    if (_actionIdLabel)
    {
        _actionIdLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_ACTION_ID));
        _actionIdLabel->SetMnemonicTarget(_actionIdField);
    }
    if (_actionNameLabel)
    {
        _actionNameLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_ACTION_NAME));
        _actionNameLabel->SetMnemonicTarget(_actionNameField);
    }
    if (_actionKindLabel)
    {
        _actionKindLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_ACTION_TYPE));
        _actionKindLabel->SetMnemonicTarget(_actionKindCombo);
    }
    if (_pluginIdLabel)
    {
        _pluginIdLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_PLUGIN_ID));
        _pluginIdLabel->SetMnemonicTarget(_pluginIdCombo);
    }
    if (_executableLabel)
    {
        _executableLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_EXECUTABLE));
        _executableLabel->SetMnemonicTarget(_executableField);
    }
    if (_argumentsLabel)
    {
        _argumentsLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_ARGUMENTS));
        _argumentsLabel->SetMnemonicTarget(_argumentsField);
    }
    if (_workingDirectoryLabel)
    {
        _workingDirectoryLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_WORKING_DIR));
        _workingDirectoryLabel->SetMnemonicTarget(_workingDirectoryField);
    }
    if (_appliesToLabel)
    {
        _appliesToLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_APPLIES_TO));
        _appliesToLabel->SetMnemonicTarget(_appliesToField);
    }
    if (_computersLabel)
    {
        _computersLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_LABEL_COMPUTERS));
        _computersLabel->SetMnemonicTarget(_computersField);
    }
}

void FileActionPreferencesPage::SyncActionCombos(PreferencesDialogState& state) noexcept
{
    std::vector<ComboBox::Item> matchKinds;
    matchKinds.push_back(MakeComboItem(std::wstring(kComboMatchDefault), LoadRes(IDS_PREFS_FILE_ACTION_KIND_DEFAULT)));
    matchKinds.push_back(MakeComboItem(std::wstring(kComboMatchExtension), LoadRes(IDS_PREFS_FILE_ACTION_KIND_EXTENSION)));
    matchKinds.push_back(MakeComboItem(std::wstring(kComboMatchPattern), LoadRes(IDS_PREFS_FILE_ACTION_KIND_PATTERN)));
    if (_matchKindCombo)
    {
        _matchKindCombo->SetItems(std::move(matchKinds));
        _matchKindCombo->SetSelectedIndex(0u);
    }

    std::vector<ComboBox::Item> actionKinds;
    actionKinds.push_back(MakeComboItem(std::wstring(kComboActionExternal), LoadRes(IDS_PREFS_FILE_ACTION_TYPE_EXTERNAL_PROGRAM)));
    if (IsViewerFamily())
    {
        actionKinds.push_back(MakeComboItem(std::wstring(kComboActionViewer), LoadRes(IDS_PREFS_FILE_ACTION_TYPE_VIEWER_PLUGIN)));
    }
    if (_actionKindCombo)
    {
        _actionKindCombo->SetItems(std::move(actionKinds));
        _actionKindCombo->SetSelectedIndex(0u);
    }

    const std::vector<Settings::FileActionDefinition>& actions =
        IsViewerFamily() ? state.workingSettings.fileActions.viewers.actions : state.workingSettings.fileActions.editors.actions;

    std::vector<ComboBox::Item> actionItems;
    actionItems.reserve(actions.size() + 1u);
    actionItems.push_back(MakeComboItem({}, LoadRes(IDS_PREFS_FILE_ACTION_NONE)));
    for (const Settings::FileActionDefinition& action : actions)
    {
        if (! action.id.empty())
        {
            actionItems.push_back(MakeComboItem(action.id, FileActionDisplayName(action)));
        }
    }

    auto setActionItems = [&](ComboBox* combo) noexcept
    {
        if (! combo)
        {
            return;
        }
        combo->SetItems(actionItems);
        combo->SetSelectedIndex(0u);
    };

    setActionItems(_primaryActionCombo);
    setActionItems(_alternateActionCombo);
    setActionItems(_editNewActionCombo);

    if (_pluginIdCombo)
    {
        _pluginIdCombo->SetItems(BuildViewerPluginComboItems(state, actions));
        _pluginIdCombo->SetSelectedIndex(0u);
    }
}

void FileActionPreferencesPage::SyncActionFieldAvailability() noexcept
{
    const Settings::FileActionKind actionKind = ActionKindFromValue(_actionKindCombo ? _actionKindCombo->GetSelectedValue() : std::wstring_view{});
    const bool viewerPluginAction             = IsViewerFamily() && actionKind == Settings::FileActionKind::ViewerPlugin;
    const bool externalAction                 = ! viewerPluginAction;

    if (_pluginIdLabel)
    {
        _pluginIdLabel->SetVisible(IsViewerFamily());
        _pluginIdLabel->SetEnabled(viewerPluginAction);
    }
    if (_pluginIdCombo)
    {
        _pluginIdCombo->SetVisible(IsViewerFamily());
        _pluginIdCombo->SetEnabled(viewerPluginAction);
    }

    if (_executableLabel)
    {
        _executableLabel->SetEnabled(externalAction);
    }
    if (_executableField)
    {
        _executableField->SetEnabled(externalAction);
    }
}

void FileActionPreferencesPage::RebuildModels(PreferencesDialogState& state) noexcept
{
    if (! _associationsModel || ! _actionsModel)
    {
        return;
    }

    const std::wstring_view filter         = IsViewerFamily() ? std::wstring_view(state.viewersSearchText) : std::wstring_view(_editorSearchText);
    const GridSortSpec associationSortSpec = _associationsGrid ? _associationsGrid->GetSortSpec() : GridSortSpec{};
    const GridSortSpec actionSortSpec      = _actionsGrid ? _actionsGrid->GetSortSpec() : GridSortSpec{};
    if (IsViewerFamily())
    {
        _associationsModel->SetColumns(ViewerAssociationColumns());
        _associationsModel->SetRows(BuildViewerAssociationRows(state.workingSettings.fileActions.viewers, filter));
        _actionsModel->SetRows(BuildActionRows(state.workingSettings.fileActions.viewers.actions, filter, L"viewers"));
    }
    else
    {
        _associationsModel->SetColumns(EditorAssociationColumns());
        _associationsModel->SetRows(BuildEditorAssociationRows(state.workingSettings.fileActions.editors, filter));
        _actionsModel->SetRows(BuildActionRows(state.workingSettings.fileActions.editors.actions, filter, L"editors"));
    }
    _actionsModel->SetColumns(ActionColumns());
    _associationsModel->SortRows(associationSortSpec);
    _actionsModel->SortRows(actionSortSpec);

    if (_associationsGrid)
    {
        _associationsGrid->NotifyDataChanged();
        if (_associationsModel->GetRowCount() > 0u && ! _associationsGrid->GetPrimarySelectedRow().has_value())
        {
            static_cast<void>(_associationsGrid->RequestSelectRow(0u, 0u));
        }
    }
    if (_actionsGrid)
    {
        _actionsGrid->NotifyDataChanged();
        if (_actionsModel->GetRowCount() > 0u && ! _actionsGrid->GetPrimarySelectedRow().has_value())
        {
            static_cast<void>(_actionsGrid->RequestSelectRow(0u, 0u));
        }
    }
}

void FileActionPreferencesPage::SyncAssociationFormFromSelection(PreferencesDialogState& state) noexcept
{
    const FileActionGridRow* selectedRow = nullptr;
    if (_associationsGrid && _associationsModel)
    {
        if (const std::optional<size_t> row = _associationsGrid->GetPrimarySelectedRow())
        {
            selectedRow = _associationsModel->GetRow(row.value());
        }
    }

    Settings::FileActionMatch match{};
    std::wstring computerName;
    std::wstring primaryActionId;
    std::wstring alternateActionId;
    std::wstring editNewActionId;

    if (selectedRow)
    {
        if (IsViewerFamily() && selectedRow->sourceIndex < state.workingSettings.fileActions.viewers.associations.size())
        {
            const Settings::ViewerAssociationRule& rule = state.workingSettings.fileActions.viewers.associations[selectedRow->sourceIndex];
            match                                       = rule.match;
            computerName                                = rule.computerName;
            primaryActionId                             = rule.viewActionId;
            alternateActionId                           = rule.alternateViewActionId;
        }
        else if (IsEditorsFamily() && selectedRow->sourceIndex < state.workingSettings.fileActions.editors.associations.size())
        {
            const Settings::EditorAssociationRule& rule = state.workingSettings.fileActions.editors.associations[selectedRow->sourceIndex];
            match                                       = rule.match;
            computerName                                = rule.computerName;
            primaryActionId                             = rule.editActionId;
            alternateActionId                           = rule.alternateEditActionId;
            editNewActionId                             = rule.editNewActionId;
        }
    }
    if (IsViewerFamily())
    {
        state.viewersSelectedExtensionText = MatchInputDisplay(match);
    }

    if (IsViewerFamily())
    {
        state.viewersSelectedExtensionText = selectedRow ? MatchInputDisplay(match) : std::wstring{};
    }

    _syncing = true;
    SelectComboValue(_matchKindCombo, MatchKindValue(match.kind));
    SetTextNoNotify(_matchValueField, MatchInputDisplay(match));
    SetTextNoNotify(_computerField, computerName);
    SelectComboValue(_primaryActionCombo, primaryActionId);
    SelectComboValue(_alternateActionCombo, alternateActionId);
    SelectComboValue(_editNewActionCombo, editNewActionId);
    _syncing = false;
}

void FileActionPreferencesPage::SyncActionFormFromSelection(PreferencesDialogState& state) noexcept
{
    const FileActionGridRow* selectedRow = nullptr;
    if (_actionsGrid && _actionsModel)
    {
        if (const std::optional<size_t> row = _actionsGrid->GetPrimarySelectedRow())
        {
            selectedRow = _actionsModel->GetRow(row.value());
        }
    }

    Settings::FileActionDefinition action{};
    action.enabled = true;
    action.kind    = IsViewerFamily() ? Settings::FileActionKind::ViewerPlugin : Settings::FileActionKind::ExternalProgram;
    action.appliesTo.matches.push_back(Settings::FileActionMatch{.kind = Settings::FileActionMatchKind::Default});

    const std::vector<Settings::FileActionDefinition>& actions =
        IsViewerFamily() ? state.workingSettings.fileActions.viewers.actions : state.workingSettings.fileActions.editors.actions;
    if (selectedRow && selectedRow->sourceIndex < actions.size())
    {
        action = actions[selectedRow->sourceIndex];
    }

    _syncing = true;
    SetTextNoNotify(_actionIdField, action.id);
    SetTextNoNotify(_actionNameField, action.displayName);
    SelectComboValue(_actionKindCombo, ActionKindValue(action.kind));
    if (_actionEnabledCheckbox)
    {
        _actionEnabledCheckbox->SetChecked(action.enabled);
    }
    SelectComboValue(_pluginIdCombo, action.pluginId);
    SetTextNoNotify(_executableField, action.executablePath);
    SetTextNoNotify(_argumentsField, action.arguments);
    SetTextNoNotify(_workingDirectoryField, action.workingDirectory);
    SetTextNoNotify(_appliesToField, FormatMatches(action.appliesTo.matches));
    SetTextNoNotify(_computersField, JoinStrings(action.appliesTo.computerNames));
    SyncActionFieldAvailability();
    _syncing = false;
}

void FileActionPreferencesPage::UpdatePreview(PreferencesDialogState& state) noexcept
{
    if (! _previewLabel)
    {
        return;
    }

    const std::filesystem::path path = PreviewPathFromField(_testFileField);
    const std::wstring computerName  = CurrentComputerName();
    FileActionResolver::Request request{
        .command = IsViewerFamily() ? FileActionResolver::Command::View : FileActionResolver::Command::Edit, .filePath = path, .computerName = computerName};
    const FileActionResolver::Resolution resolution = IsViewerFamily()
                                                          ? FileActionResolver::ResolveViewerAction(state.workingSettings.fileActions.viewers, request)
                                                          : FileActionResolver::ResolveEditorAction(state.workingSettings.fileActions.editors, request);

    _previewActionId = resolution.actionId;
    _previewReason   = resolution.reasonText;

    const std::wstring commandText = CommandDisplay(request.command);
    if (resolution.IsResolved())
    {
        _previewLabel->SetText(
            FormatStringResource(nullptr, IDS_PREFS_FILE_ACTION_PREVIEW_FMT, commandText, FileActionDisplayName(*resolution.action), resolution.reasonText));
    }
    else
    {
        _previewLabel->SetText(FormatStringResource(nullptr, IDS_PREFS_FILE_ACTION_PREVIEW_UNRESOLVED_FMT, commandText, resolution.reasonText));
    }
    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

void FileActionPreferencesPage::SyncFromState(PreferencesDialogState& state) noexcept
{
    if (IsViewerFamily() && _searchField)
    {
        _syncing = true;
        _searchField->SetText(state.viewersSearchText);
        _syncing = false;
    }
    else if (IsEditorsFamily() && _searchField)
    {
        _syncing = true;
        _searchField->SetText(_editorSearchText);
        _syncing = false;
    }

    SyncActionCombos(state);
    RebuildModels(state);
    SyncAssociationFormFromSelection(state);
    SyncActionFormFromSelection(state);
    UpdatePreview(state);
    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

void FileActionPreferencesPage::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _state      = &state;
    _hostWindow = parent;
    if (state.currentCategory != Category())
    {
        return;
    }

    if (! EnsureDxPageHost(parent, state))
    {
        Debug::Error(L"Preferences.FileActions: Failed to initialize DxUi hosts.");
        DetachDxPageHost();
    }
}

void FileActionPreferencesPage::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    _state      = &state;
    _hostWindow = host;
    if (state.currentCategory == Category() && ! EnsureDxPageHost(host, state))
    {
        Debug::Error(L"Preferences.FileActions: Failed to refresh DxUi hosts.");
        return;
    }

    ApplyTheme(state);
    SyncFromState(state);
}

void FileActionPreferencesPage::LayoutPage(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept
{
    if (! host || ! EnsureDxPageHost(host, state))
    {
        return;
    }

    _hostWindow                = host;
    _state                     = &state;
    _usesDxUiTypographyContext = true;
    _usesDxUiTypographyMetrics = true;

    Debug::Perf::Scope layoutPerf(IsViewerFamily() ? L"preferences.ui.fileactions.viewers_layout_us" : L"preferences.ui.fileactions.editors_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(typography.dpi);

    const UINT dpi     = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);
    const auto pxToDip = [dpi](const int pixels) noexcept { return (static_cast<float>(pixels) * 96.0f) / static_cast<float>(dpi); };

    const int tabHeader = UiMetrics::ScaleDip(dpi, 36);
    const int rowHeight = UiMetrics::ScaleDip(dpi, 28);
    RECT hostClient{};
    static_cast<void>(GetClientRect(host, &hostClient));
    const int hostHeight     = std::max(0l, hostClient.bottom - hostClient.top);
    const int minGridHeight  = UiMetrics::ScaleDip(dpi, 220);
    const int contentInsetX  = UiMetrics::ScaleDip(dpi, 16);
    const int contentInsetY  = UiMetrics::ScaleDip(dpi, 12);
    const int contentX       = x + contentInsetX;
    const int contentWidth   = std::max(0, width - (contentInsetX * 2));
    const int labelWidth     = std::min(contentWidth / 3, UiMetrics::ScaleDip(dpi, 160));
    const int gapX           = UiMetrics::ScaleDip(dpi, 8);
    const int halfWidth      = std::max(labelWidth + UiMetrics::ScaleDip(dpi, 120), (contentWidth - gapX) / 2);
    const int fieldWidth     = std::max(UiMetrics::ScaleDip(dpi, 110), halfWidth - labelWidth - gapX);
    const int leftX          = contentX;
    const int rightX         = contentX + halfWidth + gapX;
    const int fullFieldWidth = std::max(0, contentWidth - labelWidth - gapX);

    const int associationRows            = IsViewerFamily() ? 6 : 7;
    const int actionRows                 = IsViewerFamily() ? 8 : 7;
    const int associationAboveGridHeight = tabHeader + contentInsetY + rowHeight + gapY;
    const int associationBelowGridHeight = gapY + (associationRows * (rowHeight + gapY)) + UiMetrics::ScaleDip(dpi, 58) + contentInsetY;
    const int actionAboveGridHeight      = tabHeader + contentInsetY;
    const int actionBelowGridHeight      = gapY + (actionRows * (rowHeight + gapY)) + margin + contentInsetY;
    const int fixedPageHeightWithoutGrid = std::max(associationAboveGridHeight + associationBelowGridHeight, actionAboveGridHeight + actionBelowGridHeight);
    const int availablePageHeight        = std::max(0, hostHeight - y - margin);
    const int gridHeight                 = std::max(minGridHeight, availablePageHeight - fixedPageHeightWithoutGrid);
    const int associationPageHeight      = associationAboveGridHeight + gridHeight + associationBelowGridHeight;
    const int actionPageHeight           = actionAboveGridHeight + gridHeight + actionBelowGridHeight;
    const int pageHeight                 = std::max(associationPageHeight, actionPageHeight);

    if (_tabs)
    {
        _tabs->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + pageHeight)));
    }

    int rowY = y + tabHeader + contentInsetY;
    SetLabelBounds(_searchLabel, pxToDip(leftX), pxToDip(rowY), pxToDip(leftX + labelWidth), pxToDip(rowY + rowHeight));
    SetControlBounds(_searchField, pxToDip(leftX + labelWidth + gapX), pxToDip(rowY), pxToDip(leftX + contentWidth), pxToDip(rowY + rowHeight));
    rowY += rowHeight + gapY;
    SetControlBounds(_associationsGrid, pxToDip(leftX), pxToDip(rowY), pxToDip(leftX + contentWidth), pxToDip(rowY + gridHeight));
    rowY += gridHeight + gapY;

    const auto layoutPair = [&](Label* label, auto* control, const int baseX, const int baseY, const int baseFieldWidth) noexcept
    {
        SetLabelBounds(label, pxToDip(baseX), pxToDip(baseY), pxToDip(baseX + labelWidth), pxToDip(baseY + rowHeight));
        SetControlBounds(
            control, pxToDip(baseX + labelWidth + gapX), pxToDip(baseY), pxToDip(baseX + labelWidth + gapX + baseFieldWidth), pxToDip(baseY + rowHeight));
    };

    layoutPair(_matchKindLabel, _matchKindCombo, leftX, rowY, fieldWidth);
    layoutPair(_matchValueLabel, _matchValueField, rightX, rowY, std::max(0, contentWidth - (rightX - contentX) - labelWidth - gapX));
    rowY += rowHeight + gapY;
    layoutPair(_computerLabel, _computerField, leftX, rowY, fieldWidth);
    layoutPair(_primaryActionLabel, _primaryActionCombo, rightX, rowY, std::max(0, contentWidth - (rightX - contentX) - labelWidth - gapX));
    rowY += rowHeight + gapY;
    layoutPair(_alternateActionLabel, _alternateActionCombo, leftX, rowY, fieldWidth);
    if (IsEditorsFamily())
    {
        layoutPair(_editNewActionLabel, _editNewActionCombo, rightX, rowY, std::max(0, contentWidth - (rightX - contentX) - labelWidth - gapX));
        rowY += rowHeight + gapY;
    }
    else
    {
        rowY += rowHeight + gapY;
    }
    layoutPair(_testFileLabel, _testFileField, leftX, rowY, fullFieldWidth);
    rowY += rowHeight + gapY;
    SetControlBounds(_associationSaveButton, pxToDip(leftX), pxToDip(rowY), pxToDip(leftX + UiMetrics::ScaleDip(dpi, 150)), pxToDip(rowY + rowHeight));
    SetControlBounds(_associationRemoveButton,
                     pxToDip(leftX + UiMetrics::ScaleDip(dpi, 158)),
                     pxToDip(rowY),
                     pxToDip(leftX + UiMetrics::ScaleDip(dpi, 270)),
                     pxToDip(rowY + rowHeight));
    SetControlBounds(_associationResetButton,
                     pxToDip(leftX + UiMetrics::ScaleDip(dpi, 278)),
                     pxToDip(rowY),
                     pxToDip(leftX + UiMetrics::ScaleDip(dpi, 430)),
                     pxToDip(rowY + rowHeight));
    rowY += rowHeight + gapY;
    SetControlBounds(_previewLabel, pxToDip(leftX), pxToDip(rowY), pxToDip(leftX + contentWidth), pxToDip(rowY + UiMetrics::ScaleDip(dpi, 52)));

    int actionY = y + tabHeader + contentInsetY;
    SetControlBounds(_actionsGrid, pxToDip(leftX), pxToDip(actionY), pxToDip(leftX + contentWidth), pxToDip(actionY + gridHeight));
    actionY += gridHeight + gapY;
    layoutPair(_actionIdLabel, _actionIdField, leftX, actionY, fieldWidth);
    layoutPair(_actionNameLabel, _actionNameField, rightX, actionY, std::max(0, contentWidth - (rightX - contentX) - labelWidth - gapX));
    actionY += rowHeight + gapY;
    layoutPair(_actionKindLabel, _actionKindCombo, leftX, actionY, fieldWidth);
    SetControlBounds(_actionEnabledCheckbox, pxToDip(rightX), pxToDip(actionY), pxToDip(rightX + UiMetrics::ScaleDip(dpi, 160)), pxToDip(actionY + rowHeight));
    actionY += rowHeight + gapY;
    if (IsViewerFamily())
    {
        layoutPair(_pluginIdLabel, _pluginIdCombo, leftX, actionY, fullFieldWidth);
        actionY += rowHeight + gapY;
    }
    layoutPair(_executableLabel, _executableField, leftX, actionY, fullFieldWidth);
    actionY += rowHeight + gapY;
    layoutPair(_argumentsLabel, _argumentsField, leftX, actionY, fullFieldWidth);
    actionY += rowHeight + gapY;
    layoutPair(_workingDirectoryLabel, _workingDirectoryField, leftX, actionY, fullFieldWidth);
    actionY += rowHeight + gapY;
    layoutPair(_appliesToLabel, _appliesToField, leftX, actionY, fullFieldWidth);
    actionY += rowHeight + gapY;
    layoutPair(_computersLabel, _computersField, leftX, actionY, fullFieldWidth);
    actionY += rowHeight + gapY;
    SetControlBounds(_actionSaveButton, pxToDip(leftX), pxToDip(actionY), pxToDip(leftX + UiMetrics::ScaleDip(dpi, 130)), pxToDip(actionY + rowHeight));
    SetControlBounds(_actionRemoveButton,
                     pxToDip(leftX + UiMetrics::ScaleDip(dpi, 138)),
                     pxToDip(actionY),
                     pxToDip(leftX + UiMetrics::ScaleDip(dpi, 250)),
                     pxToDip(actionY + rowHeight));

    y += pageHeight + margin;
    state.pageHostDirectContentBottomPx = std::max(state.pageHostDirectContentBottomPx, y);

    if (_associationsGrid)
    {
        const GridVisibleWorkMetrics metrics = _associationsGrid->GetVisibleWorkMetrics();
        Debug::Perf::Emit(L"preferences.ui.fileactions.associations_visible_rows",
                          MetricFamilyText(),
                          static_cast<uint64_t>(metrics.visibleRowCount),
                          static_cast<uint64_t>(_associationsModel ? _associationsModel->GetRowCount() : 0u),
                          static_cast<uint64_t>(metrics.visibleCellCount),
                          S_OK);
    }

    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

bool FileActionPreferencesPage::HandleDeferredAction(HWND host, PreferencesDialogState& state, const PreferencesDeferredActionKind action) noexcept
{
    if (action != PreferencesDeferredActionKind::ViewersSearchChanged)
    {
        return false;
    }
    if (! host || state.currentCategory != Category())
    {
        return true;
    }
    RebuildModels(state);
    SyncAssociationFormFromSelection(state);
    SyncActionFormFromSelection(state);
    UpdatePreview(state);
    return true;
}

void FileActionPreferencesPage::OnGridSortRequested(const GridSortSpec& sortSpec)
{
    const RedSalamander::DxUi::Control* focus = _pageHost ? _pageHost->GetFocusControl() : nullptr;
    const bool associationsFocused            = focus == _associationsGrid;
    const bool actionsFocused                 = focus == _actionsGrid;
    if (_associationsGrid && (associationsFocused || (! actionsFocused && _activeGrid == ActiveGrid::Associations)))
    {
        _activeGrid = ActiveGrid::Associations;
        if (_associationsModel)
        {
            _associationsModel->SortRows(sortSpec);
        }
        _associationsGrid->SetSortSpec(sortSpec);
        _associationsGrid->NotifyDataChanged();
        if (_state)
        {
            SyncAssociationFormFromSelection(*_state);
        }
    }
    else if (_actionsGrid)
    {
        _activeGrid = ActiveGrid::Actions;
        if (_actionsModel)
        {
            _actionsModel->SortRows(sortSpec);
        }
        _actionsGrid->SetSortSpec(sortSpec);
        _actionsGrid->NotifyDataChanged();
        if (_state)
        {
            SyncActionFormFromSelection(*_state);
        }
    }

    if (_state)
    {
        UpdatePreview(*_state);
    }
    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

void FileActionPreferencesPage::OnGridSelectionChanged(Grid& sender)
{
    if (&sender == _actionsGrid)
    {
        _activeGrid = ActiveGrid::Actions;
        OnActionSelectionChanged();
    }
    else
    {
        _activeGrid = ActiveGrid::Associations;
        OnAssociationSelectionChanged();
    }
}

void FileActionPreferencesPage::OnGridSelectionChanged()
{
    if (_activeGrid == ActiveGrid::Actions)
    {
        OnActionSelectionChanged();
    }
    else
    {
        OnAssociationSelectionChanged();
    }
}

void FileActionPreferencesPage::OnSearchChanged(PreferencesDialogState& state, std::wstring_view text) noexcept
{
    if (IsViewerFamily())
    {
        state.viewersSearchText.assign(text);
    }
    else
    {
        _editorSearchText.assign(text);
    }
    RebuildModels(state);
    SyncAssociationFormFromSelection(state);
    SyncActionFormFromSelection(state);
    UpdatePreview(state);
}

void FileActionPreferencesPage::OnAssociationSelectionChanged() noexcept
{
    if (_state)
    {
        SyncAssociationFormFromSelection(*_state);
        UpdatePreview(*_state);
    }
}

void FileActionPreferencesPage::OnActionSelectionChanged() noexcept
{
    if (_state)
    {
        SyncActionFormFromSelection(*_state);
    }
}

void FileActionPreferencesPage::ClearAssociationSelectionAfterMutation(PreferencesDialogState& state) noexcept
{
    if (_associationsGrid)
    {
        _associationsGrid->GetSelectionModel().Clear();
    }
    SyncAssociationFormFromSelection(state);
    UpdatePreview(state);
    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

void FileActionPreferencesPage::MarkDirty(PreferencesDialogState& state) noexcept
{
    if (_hostWindow && IsWindow(_hostWindow) != FALSE)
    {
        if (HWND dlg = GetParent(_hostWindow))
        {
            SetDirty(dlg, state);
        }
    }
}

void FileActionPreferencesPage::SaveAssociation(PreferencesDialogState& state) noexcept
{
    const Settings::FileActionMatchKind kind = MatchKindFromValue(_matchKindCombo ? _matchKindCombo->GetSelectedValue() : std::wstring_view{});
    const std::wstring value                 = _matchValueField ? Trim(_matchValueField->GetText()) : std::wstring{};
    if (kind != Settings::FileActionMatchKind::Default && value.empty())
    {
        if (_previewLabel)
        {
            _previewLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_VALIDATION_MATCH_REQUIRED));
        }
        return;
    }

    const Settings::FileActionMatch match     = BuildMatch(kind, value);
    const std::wstring computerName           = _computerField ? Trim(_computerField->GetText()) : std::wstring{};
    const std::wstring primaryId              = _primaryActionCombo ? std::wstring(_primaryActionCombo->GetSelectedValue()) : std::wstring{};
    const std::wstring alternateId            = _alternateActionCombo ? std::wstring(_alternateActionCombo->GetSelectedValue()) : std::wstring{};
    const std::wstring editNewId              = _editNewActionCombo ? std::wstring(_editNewActionCombo->GetSelectedValue()) : std::wstring{};
    const auto selectedAssociationSourceIndex = [&]() noexcept -> std::optional<size_t>
    {
        if (! _associationsGrid || ! _associationsModel)
        {
            return std::nullopt;
        }

        const std::optional<size_t> rowIndex = _associationsGrid->GetPrimarySelectedRow();
        if (! rowIndex.has_value())
        {
            return std::nullopt;
        }

        const FileActionGridRow* row = _associationsModel->GetRow(rowIndex.value());
        if (! row)
        {
            return std::nullopt;
        }

        return row->sourceIndex;
    }();

    if (IsViewerFamily())
    {
        Settings::ViewerFileActionsSettings& settings = state.workingSettings.fileActions.viewers;
        Settings::ViewerAssociationRule rule{};
        rule.match                           = match;
        rule.computerName                    = computerName;
        rule.viewActionId                    = primaryId;
        rule.alternateViewActionId           = alternateId;
        const std::optional<size_t> existing = FindAssociationByKey(settings.associations, rule.match, rule.computerName);
        if (selectedAssociationSourceIndex.has_value() && selectedAssociationSourceIndex.value() < settings.associations.size() &&
            (! existing.has_value() || existing.value() == selectedAssociationSourceIndex.value()))
        {
            settings.associations[selectedAssociationSourceIndex.value()] = std::move(rule);
        }
        else if (existing.has_value())
        {
            settings.associations[existing.value()] = std::move(rule);
        }
        else
        {
            settings.associations.push_back(std::move(rule));
        }
    }
    else
    {
        Settings::EditorFileActionsSettings& settings = state.workingSettings.fileActions.editors;
        Settings::EditorAssociationRule rule{};
        rule.match                           = match;
        rule.computerName                    = computerName;
        rule.editActionId                    = primaryId;
        rule.alternateEditActionId           = alternateId;
        rule.editNewActionId                 = editNewId;
        const std::optional<size_t> existing = FindAssociationByKey(settings.associations, rule.match, rule.computerName);
        if (selectedAssociationSourceIndex.has_value() && selectedAssociationSourceIndex.value() < settings.associations.size() &&
            (! existing.has_value() || existing.value() == selectedAssociationSourceIndex.value()))
        {
            settings.associations[selectedAssociationSourceIndex.value()] = std::move(rule);
        }
        else if (existing.has_value())
        {
            settings.associations[existing.value()] = std::move(rule);
        }
        else
        {
            settings.associations.push_back(std::move(rule));
        }
    }

    MarkDirty(state);
    SyncFromState(state);
}

void FileActionPreferencesPage::RemoveSelectedAssociation(PreferencesDialogState& state) noexcept
{
    if (! _associationsGrid || ! _associationsModel)
    {
        return;
    }

    const std::optional<size_t> rowIndex = _associationsGrid->GetPrimarySelectedRow();
    if (! rowIndex.has_value())
    {
        return;
    }

    const FileActionGridRow* row = _associationsModel->GetRow(rowIndex.value());
    if (! row)
    {
        return;
    }

    if (IsViewerFamily())
    {
        auto& rules = state.workingSettings.fileActions.viewers.associations;
        if (row->sourceIndex < rules.size())
        {
            rules.erase(rules.begin() + static_cast<std::ptrdiff_t>(row->sourceIndex));
        }
    }
    else
    {
        auto& rules = state.workingSettings.fileActions.editors.associations;
        if (row->sourceIndex < rules.size())
        {
            rules.erase(rules.begin() + static_cast<std::ptrdiff_t>(row->sourceIndex));
        }
    }

    MarkDirty(state);
    SyncFromState(state);
    ClearAssociationSelectionAfterMutation(state);
}

void FileActionPreferencesPage::ResetAssociationsAndActions(PreferencesDialogState& state) noexcept
{
    if (IsViewerFamily())
    {
        state.workingSettings.fileActions.viewers = Settings::DefaultViewerFileActionsSettings();
    }
    else
    {
        state.workingSettings.fileActions.editors = Settings::DefaultEditorFileActionsSettings();
    }

    MarkDirty(state);
    SyncFromState(state);
    ClearAssociationSelectionAfterMutation(state);
}

void FileActionPreferencesPage::SaveAction(PreferencesDialogState& state) noexcept
{
    Settings::FileActionDefinition action{};
    action.id = _actionIdField ? Trim(_actionIdField->GetText()) : std::wstring{};
    if (action.id.empty())
    {
        if (_previewLabel)
        {
            _previewLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_VALIDATION_ID_REQUIRED));
        }
        return;
    }

    action.displayName = _actionNameField ? Trim(_actionNameField->GetText()) : std::wstring{};
    action.enabled     = ! _actionEnabledCheckbox || _actionEnabledCheckbox->IsChecked();
    action.kind        = ActionKindFromValue(_actionKindCombo ? _actionKindCombo->GetSelectedValue() : std::wstring_view{});
    if (IsEditorsFamily())
    {
        action.kind = Settings::FileActionKind::ExternalProgram;
    }
    if (action.kind == Settings::FileActionKind::ViewerPlugin)
    {
        action.pluginId = _pluginIdCombo ? std::wstring(_pluginIdCombo->GetSelectedValue()) : std::wstring{};
        if (action.pluginId.empty())
        {
            if (_previewLabel)
            {
                _previewLabel->SetText(LoadRes(IDS_PREFS_FILE_ACTION_VALIDATION_PLUGIN_REQUIRED));
            }
            return;
        }
    }
    else
    {
        action.executablePath = _executableField ? Trim(_executableField->GetText()) : std::wstring{};
    }
    action.arguments               = _argumentsField ? std::wstring(_argumentsField->GetText()) : std::wstring{};
    action.workingDirectory        = _workingDirectoryField ? Trim(_workingDirectoryField->GetText()) : std::wstring{};
    action.appliesTo.matches       = ParseMatchesField(_appliesToField ? _appliesToField->GetText() : std::wstring_view{});
    action.appliesTo.computerNames = SplitList(_computersField ? _computersField->GetText() : std::wstring_view{});

    std::vector<Settings::FileActionDefinition>& actions =
        IsViewerFamily() ? state.workingSettings.fileActions.viewers.actions : state.workingSettings.fileActions.editors.actions;
    const std::optional<size_t> existing = FindActionIndexById(actions, action.id);
    if (existing.has_value())
    {
        actions[existing.value()] = std::move(action);
    }
    else
    {
        actions.push_back(std::move(action));
    }

    MarkDirty(state);
    SyncFromState(state);
}

void FileActionPreferencesPage::RemoveSelectedAction(PreferencesDialogState& state) noexcept
{
    if (! _actionsGrid || ! _actionsModel)
    {
        return;
    }

    const std::optional<size_t> rowIndex = _actionsGrid->GetPrimarySelectedRow();
    if (! rowIndex.has_value())
    {
        return;
    }

    const FileActionGridRow* row = _actionsModel->GetRow(rowIndex.value());
    if (! row)
    {
        return;
    }

    std::vector<Settings::FileActionDefinition>& actions =
        IsViewerFamily() ? state.workingSettings.fileActions.viewers.actions : state.workingSettings.fileActions.editors.actions;
    if (row->sourceIndex < actions.size())
    {
        actions.erase(actions.begin() + static_cast<std::ptrdiff_t>(row->sourceIndex));
    }

    MarkDirty(state);
    SyncFromState(state);
}

bool FileActionPreferencesPage::SelectDefaultAction(PreferencesDialogState& state, const bool alternate, std::wstring_view actionId) noexcept
{
    if (! actionId.empty())
    {
        const std::vector<Settings::FileActionDefinition>& actions =
            IsViewerFamily() ? state.workingSettings.fileActions.viewers.actions : state.workingSettings.fileActions.editors.actions;
        if (! FindActionById(actions, actionId))
        {
            if (_previewLabel)
            {
                _previewLabel->SetText(FormatStringResource(nullptr, IDS_PREFS_FILE_ACTION_VALIDATION_ACTION_MISSING_FMT, std::wstring(actionId)));
            }
            return false;
        }
    }

    if (IsViewerFamily())
    {
        Settings::ViewerAssociationRule& rule = EnsureDefaultViewerAssociation(state.workingSettings.fileActions.viewers);
        std::wstring& target                  = alternate ? rule.alternateViewActionId : rule.viewActionId;
        if (target == actionId)
        {
            return true;
        }
        target.assign(actionId);
    }
    else
    {
        Settings::EditorAssociationRule& rule = EnsureDefaultEditorAssociation(state.workingSettings.fileActions.editors);
        std::wstring& target                  = alternate ? rule.alternateEditActionId : rule.editActionId;
        if (target == actionId)
        {
            return true;
        }
        target.assign(actionId);
    }

    MarkDirty(state);
    SyncFromState(state);
    return true;
}

bool FileActionPreferencesPage::SelectDefaultEditNewAction(PreferencesDialogState& state, std::wstring_view actionId) noexcept
{
    if (! IsEditorsFamily())
    {
        return false;
    }
    if (! actionId.empty() && ! FindActionById(state.workingSettings.fileActions.editors.actions, actionId))
    {
        if (_previewLabel)
        {
            _previewLabel->SetText(FormatStringResource(nullptr, IDS_PREFS_FILE_ACTION_VALIDATION_ACTION_MISSING_FMT, std::wstring(actionId)));
        }
        return false;
    }

    Settings::EditorAssociationRule& rule = EnsureDefaultEditorAssociation(state.workingSettings.fileActions.editors);
    if (rule.editNewActionId == actionId)
    {
        return true;
    }

    rule.editNewActionId.assign(actionId);
    MarkDirty(state);
    SyncFromState(state);
    return true;
}

#ifdef ENABLE_TESTS
size_t FileActionPreferencesPage::DebugAssociationRowCount() const noexcept
{
    return _associationsModel ? _associationsModel->GetRowCount() : 0u;
}

size_t FileActionPreferencesPage::DebugActionRowCount() const noexcept
{
    return _actionsModel ? _actionsModel->GetRowCount() : 0u;
}

GridVisibleWorkMetrics FileActionPreferencesPage::DebugAssociationVisibleWorkMetrics() const noexcept
{
    return _associationsGrid ? _associationsGrid->GetVisibleWorkMetrics() : GridVisibleWorkMetrics{};
}

uint64_t FileActionPreferencesPage::DebugAssociationRenderCount() const noexcept
{
    return _associationsGrid ? _associationsGrid->DebugGetPaintCount() : 0u;
}

uint64_t FileActionPreferencesPage::DebugAssociationResizeCount() const noexcept
{
    return _associationsGrid ? _associationsGrid->DebugGetPointerState().resizeMoveCount : 0u;
}

uint64_t FileActionPreferencesPage::DebugAssociationResizeFailureCount() const noexcept
{
    return 0u;
}

PreferencesViewersDebugFocusTarget FileActionPreferencesPage::DebugGetViewersFocusTarget() const noexcept
{
    if (! _pageHost)
    {
        return PreferencesViewersDebugFocusTarget::None;
    }

    const RedSalamander::DxUi::Control* focus = _pageHost->GetFocusControl();
    if (! focus)
    {
        return PreferencesViewersDebugFocusTarget::None;
    }

    if (focus == _tabs)
    {
        return PreferencesViewersDebugFocusTarget::Tabs;
    }
    if (focus == _searchField)
    {
        return PreferencesViewersDebugFocusTarget::SearchField;
    }
    if (focus == _associationsGrid)
    {
        return PreferencesViewersDebugFocusTarget::MappingsGrid;
    }
    if (focus == _matchKindCombo)
    {
        return PreferencesViewersDebugFocusTarget::MatchKindCombo;
    }
    if (focus == _actionsGrid)
    {
        return PreferencesViewersDebugFocusTarget::ActionsGrid;
    }
    if (focus == _matchKindCombo)
    {
        return PreferencesViewersDebugFocusTarget::MatchKindCombo;
    }
    if (focus == _matchValueField)
    {
        return PreferencesViewersDebugFocusTarget::MatchValueField;
    }
    if (focus == _computerField)
    {
        return PreferencesViewersDebugFocusTarget::ComputerField;
    }
    if (focus == _computerField)
    {
        return PreferencesViewersDebugFocusTarget::ComputerField;
    }
    if (focus == _actionIdField)
    {
        return PreferencesViewersDebugFocusTarget::ActionIdField;
    }
    if (focus == _actionNameField)
    {
        return PreferencesViewersDebugFocusTarget::ActionNameField;
    }
    if (focus == _actionKindCombo)
    {
        return PreferencesViewersDebugFocusTarget::ActionKindCombo;
    }
    if (focus == _actionEnabledCheckbox)
    {
        return PreferencesViewersDebugFocusTarget::ActionEnabledCheckbox;
    }
    if (focus == _pluginIdCombo)
    {
        return PreferencesViewersDebugFocusTarget::PluginIdField;
    }
    if (focus == _executableField)
    {
        return PreferencesViewersDebugFocusTarget::ExecutableField;
    }
    if (focus == _argumentsField)
    {
        return PreferencesViewersDebugFocusTarget::ArgumentsField;
    }
    if (focus == _workingDirectoryField)
    {
        return PreferencesViewersDebugFocusTarget::WorkingDirectoryField;
    }
    if (focus == _appliesToField)
    {
        return PreferencesViewersDebugFocusTarget::AppliesToField;
    }
    if (focus == _computersField)
    {
        return PreferencesViewersDebugFocusTarget::ComputersField;
    }
    if (focus == _testFileField)
    {
        return PreferencesViewersDebugFocusTarget::TestFileField;
    }
    if (focus == _primaryActionCombo)
    {
        return PreferencesViewersDebugFocusTarget::PrimaryActionCombo;
    }
    if (focus == _alternateActionCombo)
    {
        return PreferencesViewersDebugFocusTarget::AlternateActionCombo;
    }
    if (_editNewActionCombo && focus == _editNewActionCombo)
    {
        return PreferencesViewersDebugFocusTarget::EditNewActionCombo;
    }
    if (focus == _associationSaveButton)
    {
        return PreferencesViewersDebugFocusTarget::SaveButton;
    }
    if (focus == _associationRemoveButton)
    {
        return PreferencesViewersDebugFocusTarget::RemoveButton;
    }
    if (focus == _associationResetButton)
    {
        return PreferencesViewersDebugFocusTarget::ResetButton;
    }
    if (focus == _actionSaveButton)
    {
        return PreferencesViewersDebugFocusTarget::ActionSaveButton;
    }
    if (focus == _actionRemoveButton)
    {
        return PreferencesViewersDebugFocusTarget::ActionRemoveButton;
    }
    return PreferencesViewersDebugFocusTarget::None;
}

bool FileActionPreferencesPage::DebugUsesDxUiTypographyContext() const noexcept
{
    return _usesDxUiTypographyContext;
}

bool FileActionPreferencesPage::DebugUsesDxUiTypographyMetrics() const noexcept
{
    return _usesDxUiTypographyMetrics;
}

void FileActionPreferencesPage::DebugShowAssociationsTab() noexcept
{
    if (! _tabs || ! _associationsPage)
    {
        return;
    }

    if (_tabs->GetSelectedPage() != _associationsPage)
    {
        _tabs->SetSelectedIndex(1u);
        _activeGrid = ActiveGrid::Associations;
        if (_pageHost)
        {
            _pageHost->Invalidate();
        }
    }
}

bool FileActionPreferencesPage::DebugGetAssociationRowClientRect(const size_t rowIndex, RECT& outRect) noexcept
{
    DebugShowAssociationsTab();
    if (! _associationsGrid || ! _pageHost)
    {
        return false;
    }

    const std::optional<D2D1_RECT_F> rect = _associationsGrid->GetVisibleRowRect(rowIndex);
    if (! rect.has_value())
    {
        return false;
    }

    outRect = DipRectToPx(*_pageHost, rect.value());
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

bool FileActionPreferencesPage::DebugGetAssociationHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    DebugShowAssociationsTab();
    if (! _associationsGrid || ! _associationsModel || ! _pageHost || columnIndex >= _associationsModel->GetColumnCount())
    {
        return false;
    }

    const std::optional<D2D1_RECT_F> rect = _associationsGrid->GetVisibleColumnHeaderRect(columnIndex);
    if (! rect.has_value())
    {
        return false;
    }

    outRect = DipRectToPx(*_pageHost, rect.value());
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

bool FileActionPreferencesPage::DebugHitTestAssociationClientPoint(
    const POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) const noexcept
{
    outZone         = 0u;
    outColumnIndex  = 0u;
    outHeaderResize = false;
    outHostHitsList = false;
    if (! _associationsGrid || ! _pageHost)
    {
        return false;
    }

    const D2D1_POINT_2F pointDip =
        D2D1::Point2F(_pageHost->PixelsToDip(static_cast<float>(clientPoint.x)), _pageHost->PixelsToDip(static_cast<float>(clientPoint.y)));
    outHostHitsList = _pageHost->DebugHitTestControl(pointDip) == _associationsGrid;
    Grid::GridDebugHitInfo hit{};
    if (! _associationsGrid->DebugHitTestPoint(RedSalamander::DxUi::MakePointDip(pointDip), hit))
    {
        return false;
    }

    outZone         = hit.zone;
    outColumnIndex  = hit.columnIndex;
    outHeaderResize = hit.isHeaderResize;
    return true;
}

bool FileActionPreferencesPage::DebugGetAssociationPointerState(PreferencesGridPointerDebugState& outState) const noexcept
{
    outState = {};
    if (! _associationsGrid)
    {
        return false;
    }

    const Grid::GridDebugPointerState gridState            = _associationsGrid->DebugGetPointerState();
    outState.headerResizeDownCount                         = gridState.headerResizeDownCount;
    outState.resizeMoveCount                               = gridState.resizeMoveCount;
    outState.resizeActive                                  = gridState.resizeActive;
    outState.lastResizeDeltaDip                            = gridState.lastResizeDeltaDip;
    outState.lastResizeWidthDip                            = gridState.lastResizeWidthDip;
    outState.pressedHeaderActive                           = gridState.pressedHeaderActive;
    outState.pressedHeaderColumn                           = gridState.pressedHeaderColumn;
    outState.reorderActive                                 = gridState.reorderActive;
    outState.reorderColumn                                 = gridState.reorderColumn;
    outState.reorderTargetDisplayIndex                     = gridState.reorderTargetDisplayIndex;
    outState.headerReorderStartCount                       = gridState.headerReorderStartCount;
    outState.headerReorderCommitCount                      = gridState.headerReorderCommitCount;
    outState.headerReorderNoOpCount                        = gridState.headerReorderNoOpCount;
    outState.lastHeaderReorderColumn                       = gridState.lastHeaderReorderColumn;
    outState.lastHeaderReorderFromDisplayIndex             = gridState.lastHeaderReorderFromDisplayIndex;
    outState.lastHeaderReorderRawTargetDisplayIndex        = gridState.lastHeaderReorderRawTargetDisplayIndex;
    outState.lastHeaderReorderNormalizedTargetDisplayIndex = gridState.lastHeaderReorderNormalizedTargetDisplayIndex;
    return true;
}

bool FileActionPreferencesPage::DebugGetTabClientRect(const size_t tabIndex, RECT& outRect) const noexcept
{
    if (! _tabs || ! _pageHost || tabIndex >= _tabs->GetTabCount())
    {
        return false;
    }

    outRect = DipRectToPx(*_pageHost, _tabs->DebugGetTabRect(tabIndex));
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

bool FileActionPreferencesPage::DebugGetSelectedTabIndex(size_t& outIndex) const noexcept
{
    if (! _tabs)
    {
        return false;
    }

    const std::optional<size_t> selectedIndex = _tabs->GetSelectedIndex();
    if (! selectedIndex.has_value())
    {
        return false;
    }

    outIndex = selectedIndex.value();
    return true;
}

bool FileActionPreferencesPage::DebugSelectAssociationRow(const size_t rowIndex) noexcept
{
    DebugShowAssociationsTab();
    return _associationsGrid && _associationsGrid->RequestSelectRow(rowIndex, 0u);
}

bool FileActionPreferencesPage::DebugSetSearchText(std::wstring_view text) noexcept
{
    DebugShowAssociationsTab();
    if (! _state)
    {
        return false;
    }
    OnSearchChanged(*_state, text);
    return true;
}

bool FileActionPreferencesPage::DebugSelectDefaultAction(const bool alternate, std::wstring_view actionId) noexcept
{
    if (! _state)
    {
        return false;
    }
    return SelectDefaultAction(*_state, alternate, actionId);
}

bool FileActionPreferencesPage::DebugSelectDefaultEditNewAction(std::wstring_view actionId) noexcept
{
    if (! _state)
    {
        return false;
    }
    return SelectDefaultEditNewAction(*_state, actionId);
}

bool FileActionPreferencesPage::DebugFocusSearchField() noexcept
{
    DebugShowAssociationsTab();
    if (! _pageHost || ! _searchField)
    {
        return false;
    }
    _pageHost->SetFocusControl(_searchField);
    return _pageHost->GetFocusControl() == _searchField;
}

bool FileActionPreferencesPage::DebugScrollAssociationByWheelDetents(const int detents) noexcept
{
    DebugShowAssociationsTab();
    if (detents == 0 || ! _associationsGrid || ! _pageHost)
    {
        return false;
    }

    const float wheelDelta = detents >= 0 ? static_cast<float>(WHEEL_DELTA) : -static_cast<float>(WHEEL_DELTA);
    const int steps        = std::abs(detents);
    for (int index = 0; index < steps; ++index)
    {
        static_cast<void>(_associationsGrid->OnMouseWheel(*_pageHost, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0u));
    }

    return true;
}

std::wstring FileActionPreferencesPage::DebugPreviewActionId() const
{
    return _previewActionId;
}

std::wstring FileActionPreferencesPage::DebugPreviewReason() const
{
    return _previewReason;
}
#endif
