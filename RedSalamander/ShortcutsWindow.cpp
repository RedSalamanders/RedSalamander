#include "Framework.h"

#include "ShortcutsWindow.h"

#include <algorithm>
#include <array>
#include <cwctype>
#include <iterator>
#include <optional>
#include <unordered_map>
#include <vector>

#include "CommandRegistry.h"
#include "DxUiThemePalette.h"
#include "Helpers.h"
#include "SettingsHotReload.h"
#include "ShortcutManager.h"
#include "ShortcutText.h"
#include "WindowMaximizeBehavior.h"
#include "WindowPlacementPersistence.h"
#include "resource.h"

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

[[nodiscard]] bool DispatchShortcutCommandFromWindow(HWND ownerWindow, std::wstring_view commandId) noexcept;

namespace
{
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridGroupDesc;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::GridSortSpec;
using RedSalamander::DxUi::IDxGridDelegate;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::SortDirection;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::WindowHost;

constexpr wchar_t kShortcutsWindowId[] = L"ShortcutsWindow";
constexpr wchar_t kSettingsAppId[]     = L"RedSalamander";
constexpr wchar_t kClassName[]         = L"RedSalamander.ShortcutsWindow";

constexpr uint64_t kGroupStableIdFunctionBar = 1u;
constexpr uint64_t kGroupStableIdFolderView  = 2u;

[[nodiscard]] int CompareNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    return CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE) - CSTR_EQUAL;
}

[[nodiscard]] std::vector<RedSalamander::DxUi::GridColumnLayoutEntry> ConvertColumnLayout(const std::vector<Common::Settings::GridColumnLayoutEntry>& layout)
{
    std::vector<RedSalamander::DxUi::GridColumnLayoutEntry> converted;
    converted.reserve(layout.size());
    for (const auto& entry : layout)
    {
        if (entry.columnId.empty())
        {
            continue;
        }

        converted.push_back(RedSalamander::DxUi::GridColumnLayoutEntry{
            .columnId     = entry.columnId,
            .displayIndex = entry.displayIndex,
            .widthDip     = entry.widthDip,
        });
    }
    return converted;
}

[[nodiscard]] std::vector<Common::Settings::GridColumnLayoutEntry> ConvertColumnLayout(const std::vector<RedSalamander::DxUi::GridColumnLayoutEntry>& layout)
{
    std::vector<Common::Settings::GridColumnLayoutEntry> converted;
    converted.reserve(layout.size());
    for (const auto& entry : layout)
    {
        if (entry.columnId.empty())
        {
            continue;
        }

        converted.push_back(Common::Settings::GridColumnLayoutEntry{
            .columnId     = entry.columnId,
            .displayIndex = static_cast<uint32_t>(entry.displayIndex),
            .widthDip     = entry.widthDip,
        });
    }
    return converted;
}

[[nodiscard]] HWND NormalizeOwnedWindow(HWND owner) noexcept
{
    if (! owner || IsWindow(owner) == FALSE)
    {
        return nullptr;
    }

    return GetAncestor(owner, GA_ROOT);
}

[[nodiscard]] std::wstring_view TrimWhitespace(std::wstring_view text) noexcept
{
    while (! text.empty() && std::iswspace(static_cast<wint_t>(text.front())) != 0)
    {
        text.remove_prefix(1);
    }

    while (! text.empty() && std::iswspace(static_cast<wint_t>(text.back())) != 0)
    {
        text.remove_suffix(1);
    }

    return text;
}

[[nodiscard]] bool ContainsNoCase(std::wstring_view haystack, std::wstring_view needle) noexcept
{
    if (needle.empty())
    {
        return true;
    }

    if (needle.size() > haystack.size())
    {
        return false;
    }

    const auto it = std::search(haystack.begin(), haystack.end(), needle.begin(), needle.end(), [](wchar_t a, wchar_t b) noexcept {
        return std::towupper(static_cast<wint_t>(a)) == std::towupper(static_cast<wint_t>(b));
    });
    return it != haystack.end();
}

#ifdef ENABLE_TESTS
[[nodiscard]] bool IsActuallyVisibleChildWindow(HWND hwnd) noexcept
{
    if (! hwnd || IsWindowVisible(hwnd) == FALSE)
    {
        return false;
    }

    // DxUi text bridges stay WS_VISIBLE for IME routing, but an empty region keeps them off-screen.
    wil::unique_hrgn region(CreateRectRgn(0, 0, 0, 0));
    if (region)
    {
        const int rgnType = GetWindowRgn(hwnd, region.get());
        if (rgnType == NULLREGION)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] size_t CountVisibleChildWindowsLocal(HWND parent) noexcept
{
    if (! parent || IsWindow(parent) == FALSE)
    {
        return 0u;
    }

    size_t count = 0u;
    EnumChildWindows(parent,
                     [](HWND hwnd, LPARAM lParam) noexcept -> BOOL
    {
        auto* count = reinterpret_cast<size_t*>(lParam);
        if (count && IsActuallyVisibleChildWindow(hwnd))
        {
            ++(*count);
        }
        return TRUE;
    },
                     reinterpret_cast<LPARAM>(&count));
    return count;
}
#endif

[[nodiscard]] std::wstring GetCommandDisplayName(std::wstring_view commandId) noexcept
{
    return ShortcutText::GetCommandDisplayName(commandId);
}

[[nodiscard]] std::wstring GetCommandDescription(std::wstring_view commandId) noexcept
{
    const std::optional<unsigned int> descIdOpt = TryGetCommandDescriptionStringId(commandId);
    if (! descIdOpt.has_value())
    {
        return {};
    }

    const std::wstring description = LoadStringResource(nullptr, descIdOpt.value());
    return description;
}

[[nodiscard]] std::wstring FormatChordText(uint32_t vk, uint32_t modifiers) noexcept
{
    std::vector<std::wstring> parts;
    parts.reserve(4u);

    if ((modifiers & ShortcutManager::kModCtrl) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_CTRL));
    }
    if ((modifiers & ShortcutManager::kModAlt) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_ALT));
    }
    if ((modifiers & ShortcutManager::kModShift) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_SHIFT));
    }

    parts.push_back(ShortcutText::VkToDisplayText(vk));

    std::wstring result;
    for (const std::wstring& part : parts)
    {
        if (part.empty())
        {
            continue;
        }

        if (! result.empty())
        {
            result.append(L" + ");
        }
        result.append(part);
    }

    return result;
}

[[nodiscard]] bool IsConflictChord(uint32_t chordKey, const std::vector<uint32_t>& conflicts) noexcept
{
    return std::binary_search(conflicts.begin(), conflicts.end(), chordKey);
}

[[nodiscard]] const std::wstring& GetConflictMark() noexcept
{
    static const std::wstring mark = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_CONFLICT_MARK);
    return mark;
}

[[nodiscard]] uint64_t MakeShortcutStableRowId(uint64_t groupStableId, size_t bindingIndex) noexcept
{
    return (groupStableId << 32u) | (static_cast<uint64_t>(bindingIndex) + 1u);
}

struct ShortcutRow final
{
    uint64_t stableId  = 0u;
    uint32_t vk        = 0u;
    uint32_t modifiers = 0u;
    std::wstring commandId;
    std::wstring commandText;
    std::wstring keyText;
    std::wstring tooltipText;
    uint64_t groupStableId = 0u;
    bool hasConflict       = false;
};

enum class ShortcutKeySortGroup : uint8_t
{
    Function = 0u,
    Number,
    Letter,
    Other,
};

struct ShortcutKeySortKey final
{
    ShortcutKeySortGroup group = ShortcutKeySortGroup::Other;
    uint32_t ordinal           = 0u;
    std::wstring baseText;
    std::wstring modifierText;
};

[[nodiscard]] std::wstring FormatModifierSortText(uint32_t modifiers) noexcept
{
    std::wstring result;
    const auto appendPart = [&](UINT stringId)
    {
        const std::wstring part = LoadStringResource(nullptr, stringId);
        if (part.empty())
        {
            return;
        }
        if (! result.empty())
        {
            result.append(L" + ");
        }
        result.append(part);
    };

    if ((modifiers & ShortcutManager::kModCtrl) != 0)
    {
        appendPart(IDS_MOD_CTRL);
    }
    if ((modifiers & ShortcutManager::kModAlt) != 0)
    {
        appendPart(IDS_MOD_ALT);
    }
    if ((modifiers & ShortcutManager::kModShift) != 0)
    {
        appendPart(IDS_MOD_SHIFT);
    }

    return result;
}

[[nodiscard]] ShortcutKeySortKey MakeShortcutKeySortKey(const ShortcutRow& row)
{
    ShortcutKeySortKey key{
        .modifierText = FormatModifierSortText(row.modifiers),
    };

    if (row.vk >= VK_F1 && row.vk <= VK_F24)
    {
        key.group   = ShortcutKeySortGroup::Function;
        key.ordinal = row.vk - VK_F1 + 1u;
        return key;
    }

    if (row.vk >= static_cast<uint32_t>(L'0') && row.vk <= static_cast<uint32_t>(L'9'))
    {
        key.group   = ShortcutKeySortGroup::Number;
        key.ordinal = row.vk - static_cast<uint32_t>(L'0');
        return key;
    }

    if (row.vk >= static_cast<uint32_t>(L'A') && row.vk <= static_cast<uint32_t>(L'Z'))
    {
        key.group   = ShortcutKeySortGroup::Letter;
        key.ordinal = row.vk - static_cast<uint32_t>(L'A');
        return key;
    }

    key.group    = ShortcutKeySortGroup::Other;
    key.baseText = ShortcutText::VkToDisplayText(row.vk);
    return key;
}

[[nodiscard]] int CompareShortcutKeySort(const ShortcutRow& left, const ShortcutRow& right)
{
    const ShortcutKeySortKey leftKey  = MakeShortcutKeySortKey(left);
    const ShortcutKeySortKey rightKey = MakeShortcutKeySortKey(right);

    if (leftKey.group != rightKey.group)
    {
        return static_cast<int>(leftKey.group) - static_cast<int>(rightKey.group);
    }

    if (leftKey.ordinal != rightKey.ordinal)
    {
        return leftKey.ordinal < rightKey.ordinal ? -1 : 1;
    }

    int comparison = CompareNoCase(leftKey.baseText, rightKey.baseText);
    if (comparison != 0)
    {
        return comparison;
    }

    comparison = CompareNoCase(leftKey.modifierText, rightKey.modifierText);
    if (comparison != 0)
    {
        return comparison;
    }

    return CompareNoCase(left.keyText, right.keyText);
}

class ShortcutsGridModel final : public IDxGridModel
{
public:
    ShortcutsGridModel()
    {
        _columns = {
            {L"command", LoadStringResource(nullptr, IDS_SHORTCUTS_COL_COMMAND), 460.0f, 260.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
            {L"key", LoadStringResource(nullptr, IDS_SHORTCUTS_COL_KEY), 220.0f, 140.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
        };
    }

    void SetRows(std::vector<ShortcutRow> rows)
    {
        _rows = std::move(rows);
        RebuildGroups();
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
        if (rowIndex >= _rows.size() || columnIndex >= _columns.size())
        {
            return;
        }

        const ShortcutRow& row = _rows[rowIndex];
        if (columnIndex == 0u)
        {
            outCell.kind        = row.hasConflict ? RedSalamander::DxUi::GridCellKind::IconText : RedSalamander::DxUi::GridCellKind::Text;
            outCell.iconText    = row.hasConflict ? GetConflictMark() : std::wstring{};
            outCell.text        = row.commandText;
            outCell.multiline   = true;
            outCell.tooltipText = row.tooltipText;
            return;
        }

        outCell.text          = row.keyText;
        outCell.textAlignment = DWRITE_TEXT_ALIGNMENT_TRAILING;
        outCell.tooltipText   = row.tooltipText;
    }

    [[nodiscard]] RedSalamander::DxUi::GridRowStyle GetRowStyle(size_t rowIndex) const override
    {
        if (rowIndex >= _rows.size())
        {
            return {};
        }

        const ShortcutRow& row = _rows[rowIndex];
        return RedSalamander::DxUi::GridRowStyle{
            .tone        = RedSalamander::DxUi::GridRowTone::None,
            .rainbowSeed = ! row.commandText.empty() ? row.commandText : row.keyText,
        };
    }

    [[nodiscard]] size_t GetGroupCount() const noexcept override
    {
        return _groups.size();
    }

    [[nodiscard]] GridGroupDesc GetGroup(size_t groupIndex) const override
    {
        return _groups.at(groupIndex);
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return rowIndex < _rows.size() ? _rows[rowIndex].stableId : 0u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        const auto it = std::find_if(_rows.begin(), _rows.end(), [&](const ShortcutRow& row) noexcept { return row.stableId == rowId; });
        if (it == _rows.end())
        {
            return std::nullopt;
        }

        return static_cast<size_t>(std::distance(_rows.begin(), it));
    }

    [[nodiscard]] std::wstring GetRowAccessibleName(size_t rowIndex) const
    {
        if (rowIndex >= _rows.size())
        {
            return {};
        }

        return _rows[rowIndex].commandText;
    }

    [[nodiscard]] const ShortcutRow* GetRow(size_t rowIndex) const noexcept
    {
        return rowIndex < _rows.size() ? &_rows[rowIndex] : nullptr;
    }

    [[nodiscard]] bool SetGroupCollapsed(uint64_t stableId, bool collapsed) noexcept
    {
        _collapsedByStableId[stableId] = collapsed;

        bool found = false;
        for (auto& group : _groups)
        {
            if (group.stableId != stableId)
            {
                continue;
            }

            group.collapsed = collapsed;
            found           = true;
            break;
        }

        return found;
    }

    [[nodiscard]] bool IsGroupCollapsed(uint64_t stableId) const noexcept
    {
        const auto it = _collapsedByStableId.find(stableId);
        return it != _collapsedByStableId.end() && it->second;
    }

    void SortRows(const GridSortSpec& sortSpec)
    {
        if (_rows.empty())
        {
            RebuildGroups();
            return;
        }

        const auto compareRows = [&](const ShortcutRow& left, const ShortcutRow& right) noexcept
        {
            if (sortSpec.direction == SortDirection::None)
            {
                return left.stableId < right.stableId;
            }

            int comparison = 0;
            switch (sortSpec.columnIndex)
            {
                case 1u:
                    comparison = CompareShortcutKeySort(left, right);
                    if (comparison == 0)
                    {
                        comparison = CompareNoCase(left.commandText, right.commandText);
                    }
                    break;
                case 0u:
                default:
                    comparison = CompareNoCase(left.commandText, right.commandText);
                    if (comparison == 0)
                    {
                        comparison = CompareNoCase(left.keyText, right.keyText);
                    }
                    break;
            }

            if (comparison == 0)
            {
                return left.stableId < right.stableId;
            }

            return sortSpec.direction == SortDirection::Ascending ? comparison < 0 : comparison > 0;
        };

        size_t groupBegin = 0u;
        while (groupBegin < _rows.size())
        {
            size_t groupEnd = groupBegin + 1u;
            while (groupEnd < _rows.size() && _rows[groupEnd].groupStableId == _rows[groupBegin].groupStableId)
            {
                ++groupEnd;
            }

            std::stable_sort(_rows.begin() + static_cast<ptrdiff_t>(groupBegin), _rows.begin() + static_cast<ptrdiff_t>(groupEnd), compareRows);
            groupBegin = groupEnd;
        }

        RebuildGroups();
    }

private:
    void RebuildGroups()
    {
        _groups.clear();
        if (_rows.empty())
        {
            return;
        }

        const auto appendGroup = [&](uint64_t stableId, const std::wstring& title) noexcept
        {
            const auto it = std::find_if(_rows.begin(), _rows.end(), [&](const ShortcutRow& row) noexcept { return row.groupStableId == stableId; });
            if (it == _rows.end())
            {
                return;
            }

            const size_t start = static_cast<size_t>(std::distance(_rows.begin(), it));
            size_t count       = 0u;
            for (size_t rowIndex = start; rowIndex < _rows.size(); ++rowIndex)
            {
                if (_rows[rowIndex].groupStableId != stableId)
                {
                    break;
                }
                ++count;
            }

            _groups.push_back(GridGroupDesc{
                .stableId      = stableId,
                .title         = title,
                .startRowIndex = start,
                .rowCount      = count,
                .collapsed     = IsGroupCollapsed(stableId),
            });
        };

        appendGroup(kGroupStableIdFunctionBar, LoadStringResource(nullptr, IDS_SHORTCUTS_GROUP_FUNCTION_BAR));
        appendGroup(kGroupStableIdFolderView, LoadStringResource(nullptr, IDS_SHORTCUTS_GROUP_FOLDER_VIEW));
    }

private:
    std::vector<GridColumnDesc> _columns;
    std::vector<ShortcutRow> _rows;
    std::vector<GridGroupDesc> _groups;
    std::unordered_map<uint64_t, bool> _collapsedByStableId;
};

class ShortcutsWindow final : public IDxGridDelegate
{
public:
    ShortcutsWindow() = default;
    ~ShortcutsWindow()
    {
        Destroy();
    }

    ShortcutsWindow(const ShortcutsWindow&)            = delete;
    ShortcutsWindow& operator=(const ShortcutsWindow&) = delete;
    ShortcutsWindow(ShortcutsWindow&&)                 = delete;
    ShortcutsWindow& operator=(ShortcutsWindow&&)      = delete;

    HWND Create(HWND owner,
                Common::Settings::Settings& settings,
                const Common::Settings::ShortcutsSettings& shortcuts,
                const ShortcutManager& shortcutManager,
                const AppTheme& theme) noexcept;

    void UpdateTheme(const AppTheme& theme) noexcept;
    void UpdateData(const Common::Settings::ShortcutsSettings& shortcuts, const ShortcutManager& shortcutManager) noexcept;

    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _hWnd.get();
    }

#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugGetSnapshot(ShortcutsWindowDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugScrollByWheelDetents(int detents) noexcept;
    [[nodiscard]] bool DebugSelectRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugSetSearchText(std::wstring_view text) noexcept;
    [[nodiscard]] bool DebugFocusSearch() noexcept;
    [[nodiscard]] bool DebugFocusGrid() noexcept;
    [[nodiscard]] bool DebugCycleGridSortByColumn(size_t columnIndex) noexcept;
    [[nodiscard]] bool DebugApplyGridLayout(const std::vector<Common::Settings::GridColumnLayoutEntry>& layout) noexcept;
    [[nodiscard]] bool DebugSetGroupCollapsed(size_t groupIndex, bool collapsed) noexcept;
#endif

    void OnGridSelectionChanged(Grid& sender) override
    {
        if (! _gridModel)
        {
            _selectedRowId.reset();
            return;
        }

        const std::optional<size_t> rowIndex = sender.GetPrimarySelectedRow();
        _selectedRowId                       = rowIndex.has_value() ? std::optional<uint64_t>(_gridModel->GetStableRowId(rowIndex.value())) : std::nullopt;
    }

    void OnGridSelectionChanged() override
    {
        if (! _grid || ! _gridModel)
        {
            _selectedRowId.reset();
            return;
        }

        const std::optional<size_t> rowIndex = _grid->GetPrimarySelectedRow();
        _selectedRowId                       = rowIndex.has_value() ? std::optional<uint64_t>(_gridModel->GetStableRowId(rowIndex.value())) : std::nullopt;
    }

    void OnGridGroupToggled(Grid& /*sender*/, uint64_t groupStableId, bool collapsed) override
    {
        if (! _gridModel)
        {
            return;
        }

        static_cast<void>(_gridModel->SetGroupCollapsed(groupStableId, collapsed));
    }

    void OnGridGroupToggled(uint64_t groupStableId, bool collapsed) override
    {
        if (! _gridModel)
        {
            return;
        }

        static_cast<void>(_gridModel->SetGroupCollapsed(groupStableId, collapsed));
    }

    void OnGridSortRequested(const GridSortSpec& sortSpec) override
    {
        if (! _gridModel || ! _grid)
        {
            return;
        }

        _gridModel->SortRows(sortSpec);
        _grid->SetSortSpec(sortSpec);
        _grid->NotifyDataChanged();
        RestoreSelection();
        _dxHost.Invalidate();
    }

    void OnGridRowActivated(Grid& /*sender*/, size_t rowIndex) override
    {
        OnGridRowActivated(rowIndex);
    }

    void OnGridRowActivated(size_t rowIndex) override
    {
        if (! _gridModel || rowIndex >= _gridModel->GetRowCount())
        {
            return;
        }

        const ShortcutRow* const row = _gridModel->GetRow(rowIndex);
        if (! row || row->commandId.empty())
        {
            return;
        }

        const HWND dispatchOwner = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? _ownerWindow : _hWnd.get();
        if (! dispatchOwner || IsWindow(dispatchOwner) == FALSE)
        {
            return;
        }

        static_cast<void>(DispatchShortcutCommandFromWindow(dispatchOwner, row->commandId));
    }

private:
    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    LRESULT WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept;
    void PersistSettingsForClose(HWND hwnd) noexcept;
    void OnNcDestroy(HWND hwnd) noexcept;
    void Destroy() noexcept;
    void BuildUi();
    void ApplyTheme() noexcept;
    void Layout() noexcept;
    void RebuildRows() noexcept;
    void RestoreSelection() noexcept;
    void ResizeWindowToDefault(HWND hwnd) noexcept;

private:
    wil::unique_hwnd _hWnd;
    HINSTANCE _instance                     = nullptr;
    Common::Settings::Settings* _settings   = nullptr;
    const ShortcutManager* _shortcutManager = nullptr;
    Common::Settings::ShortcutsSettings _shortcuts;
    AppTheme _theme{};
    std::wstring _searchQuery;
    std::optional<uint64_t> _selectedRowId;
    size_t _dispatchDepth           = 0u;
    bool _deletePending             = false;
    bool _settingsPersistedForClose = false;
    HWND _ownerWindow               = nullptr;

    WindowHost _dxHost;
    std::unique_ptr<Panel> _rootStorage;
    Panel* _root           = nullptr;
    Label* _subtitleLabel  = nullptr;
    TextField* _searchEdit = nullptr;
    Grid* _grid            = nullptr;
    std::unique_ptr<ShortcutsGridModel> _gridModelStorage;
    ShortcutsGridModel* _gridModel = nullptr;
};

ShortcutsWindow* g_shortcutsWindow = nullptr;

ATOM ShortcutsWindow::RegisterWndClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom != 0)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = ShortcutsWindow::WndProcThunk;
    wc.hInstance     = instance;
    wc.hIcon         = LoadIconW(instance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIconSm       = LoadIconW(instance, MAKEINTRESOURCEW(IDI_SMALL));
    wc.lpszClassName = kClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

HWND ShortcutsWindow::Create(HWND owner,
                             Common::Settings::Settings& settings,
                             const Common::Settings::ShortcutsSettings& shortcuts,
                             const ShortcutManager& shortcutManager,
                             const AppTheme& theme) noexcept
{
    _instance = GetModuleHandleW(nullptr);
    if (! RegisterWndClass(_instance))
    {
        return nullptr;
    }

    _settings        = &settings;
    _shortcuts       = shortcuts;
    _shortcutManager = &shortcutManager;
    _theme           = theme;

    const std::wstring title   = LoadStringResource(nullptr, IDS_CMD_SHORTCUTS);
    const HWND normalizedOwner = NormalizeOwnedWindow(owner);
    _ownerWindow               = normalizedOwner;
    const UINT dpi             = normalizedOwner ? GetDpiForWindow(normalizedOwner) : GetDpiForSystem();
    const int defaultWidth     = MulDiv(900, static_cast<int>(dpi == 0u ? 96u : dpi), 96);
    const int defaultHeight    = MulDiv(620, static_cast<int>(dpi == 0u ? 96u : dpi), 96);

    const bool hasSavedPlacement = _settings->windows.find(std::wstring(kShortcutsWindowId)) != _settings->windows.end();

    const HWND hwnd = CreateWindowExW(0,
                                      kClassName,
                                      title.c_str(),
                                      WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                      CW_USEDEFAULT,
                                      CW_USEDEFAULT,
                                      defaultWidth,
                                      defaultHeight,
                                      nullptr,
                                      nullptr,
                                      _instance,
                                      this);
    if (! hwnd)
    {
        return nullptr;
    }

    if (! _hWnd)
    {
        _hWnd.reset(hwnd);
    }

    if (! hasSavedPlacement)
    {
        ResizeWindowToDefault(hwnd);
    }

    const int showCmd = hasSavedPlacement ? WindowPlacementPersistence::Restore(*_settings, kShortcutsWindowId, hwnd) : SW_SHOWNORMAL;
    ShowWindow(hwnd, showCmd);
    SetForegroundWindow(hwnd);
    return hwnd;
}

void ShortcutsWindow::UpdateTheme(const AppTheme& theme) noexcept
{
    _theme = theme;
    ApplyTheme();

    if (_hWnd)
    {
        ApplyTitleBarTheme(_hWnd.get(), _theme, GetActiveWindow() == _hWnd.get());
        _dxHost.Invalidate();
    }
}

void ShortcutsWindow::UpdateData(const Common::Settings::ShortcutsSettings& shortcuts, const ShortcutManager& shortcutManager) noexcept
{
    _shortcuts       = shortcuts;
    _shortcutManager = &shortcutManager;
    if (_gridModel)
    {
        static_cast<void>(_gridModel->SetGroupCollapsed(kGroupStableIdFunctionBar, _shortcuts.functionBarCollapsed));
        static_cast<void>(_gridModel->SetGroupCollapsed(kGroupStableIdFolderView, _shortcuts.folderViewCollapsed));
    }
    RebuildRows();
}

LRESULT CALLBACK ShortcutsWindow::WndProcThunk(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
        auto* self     = static_cast<ShortcutsWindow*>(cs ? cs->lpCreateParams : nullptr);
        if (! self)
        {
            return FALSE;
        }

        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        if (! self->_hWnd)
        {
            self->_hWnd.reset(hwnd);
        }
        g_shortcutsWindow = self;
    }

    auto* self = reinterpret_cast<ShortcutsWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! self)
    {
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    ++self->_dispatchDepth;
    const auto finishDispatch = wil::scope_exit([self]() noexcept
    {
        if (self->_dispatchDepth > 0u)
        {
            --self->_dispatchDepth;
        }
        if (self->_dispatchDepth == 0u && self->_deletePending)
        {
            delete self;
        }
    });

    return self->WndProc(hwnd, message, wParam, lParam);
}

LRESULT ShortcutsWindow::WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    bool handled     = false;
    LRESULT dxResult = 0;
    if (message != WM_CREATE)
    {
        dxResult = _dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
    }

    if (handled)
    {
        if (message == WM_NCDESTROY)
        {
            OnNcDestroy(hwnd);
        }
        else if (message == WM_SIZE || message == WM_DPICHANGED)
        {
            Layout();
        }
        return dxResult;
    }

    switch (message)
    {
        case WM_CREATE: return OnCreate(hwnd) ? 0 : -1;
        case WM_SIZE: Layout(); return 0;
        case WM_DPICHANGED:
        {
            const auto* suggested = reinterpret_cast<const RECT*>(lParam);
            if (suggested)
            {
                SetWindowPos(hwnd,
                             nullptr,
                             suggested->left,
                             suggested->top,
                             suggested->right - suggested->left,
                             suggested->bottom - suggested->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            Layout();
            return 0;
        }
        case WM_GETMINMAXINFO:
        {
            auto* info = reinterpret_cast<MINMAXINFO*>(lParam);
            if (info)
            {
                const UINT dpi         = GetDpiForWindow(hwnd);
                info->ptMinTrackSize.x = std::max<LONG>(info->ptMinTrackSize.x, MulDiv(680, static_cast<int>(dpi), 96));
                info->ptMinTrackSize.y = std::max<LONG>(info->ptMinTrackSize.y, MulDiv(420, static_cast<int>(dpi), 96));
                static_cast<void>(WindowMaximizeBehavior::ApplyVerticalMaximize(hwnd, *info));
            }
            return 0;
        }
        case WM_ERASEBKGND: return 1;
        case WM_KEYDOWN:
            if (wParam == VK_ESCAPE)
            {
                DestroyWindow(hwnd);
                return 0;
            }
            break;
        case WM_SHOWWINDOW:
            if (wParam != FALSE && _grid && ! _grid->GetPrimarySelectedRow().has_value() && _gridModel && _gridModel->GetRowCount() != 0u)
            {
                RestoreSelection();
            }
            return 0;
        case WM_ACTIVATE:
            ApplyTitleBarTheme(hwnd, _theme, LOWORD(wParam) != WA_INACTIVE);
            if (LOWORD(wParam) != WA_INACTIVE && _grid && ! _grid->GetPrimarySelectedRow().has_value() && _gridModel && _gridModel->GetRowCount() != 0u)
            {
                RestoreSelection();
            }
            return 0;
        case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, _theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
        case WM_CLOSE:
            PersistSettingsForClose(hwnd);
            DestroyWindow(hwnd);
            return 0;
        case WM_NCDESTROY: OnNcDestroy(hwnd); break;
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

bool ShortcutsWindow::OnCreate(HWND hwnd) noexcept
{
    if (! _dxHost.Attach(hwnd))
    {
        return false;
    }
    _dxHost.SetOnEscape([this]() noexcept
    {
        const HWND hwnd = _hWnd.get();
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            return false;
        }
        return PostMessageW(hwnd, WM_CLOSE, 0, 0) != FALSE;
    });

    BuildUi();
    ApplyTheme();
    RebuildRows();
    Layout();
    if (_searchEdit)
    {
        _dxHost.SetFocusControl(_searchEdit);
        RestoreSelection();
    }
    _settingsPersistedForClose = false;
    return true;
}

void ShortcutsWindow::PersistSettingsForClose(HWND hwnd) noexcept
{
    if (_settingsPersistedForClose || ! _settings || ! _hWnd || _hWnd.get() != hwnd)
    {
        return;
    }

    Common::Settings::ShortcutsSettings settings = _settings->shortcuts.value_or(_shortcuts);
    settings.functionBar                         = _shortcuts.functionBar;
    settings.folderView                          = _shortcuts.folderView;
    settings.functionBarCollapsed                = _gridModel && _gridModel->IsGroupCollapsed(kGroupStableIdFunctionBar);
    settings.folderViewCollapsed                 = _gridModel && _gridModel->IsGroupCollapsed(kGroupStableIdFolderView);
    settings.sortColumnId.clear();
    settings.sortDescending = false;
    if (_grid && _gridModel)
    {
        const auto sortSpec = _grid->GetSortSpec();
        if (sortSpec.direction != SortDirection::None && sortSpec.columnIndex < _gridModel->GetColumnCount())
        {
            settings.sortColumnId   = _gridModel->GetColumn(sortSpec.columnIndex).id;
            settings.sortDescending = sortSpec.direction == SortDirection::Descending;
        }
    }
    settings.gridLayout.clear();
    if (_grid)
    {
        settings.gridLayout = ConvertColumnLayout(_grid->CaptureColumnLayout());
    }
    _settings->shortcuts = std::move(settings);

    WindowPlacementPersistence::Save(*_settings, kShortcutsWindowId, hwnd);
    const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(kSettingsAppId, *_settings);
    if (FAILED(saveHr))
    {
        Debug::Error(L"SaveSettings failed for Shortcuts window (hr=0x{:08X}).", static_cast<unsigned long>(saveHr));
    }

    _settingsPersistedForClose = true;
}

void ShortcutsWindow::OnNcDestroy(HWND hwnd) noexcept
{
    PersistSettingsForClose(hwnd);
    const HWND restoreOwner = (_ownerWindow && IsWindow(_ownerWindow) != FALSE) ? _ownerWindow : nullptr;

    _dxHost.Detach();
    if (_hWnd.get() == hwnd)
    {
        _hWnd.release();
    }
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    if (g_shortcutsWindow == this)
    {
        g_shortcutsWindow = nullptr;
    }
    _deletePending = true;
    if (_dispatchDepth == 0u)
    {
        delete this;
    }

    if (restoreOwner)
    {
        static_cast<void>(SetActiveWindow(restoreOwner));
        static_cast<void>(SetForegroundWindow(restoreOwner));
    }
}

void ShortcutsWindow::Destroy() noexcept
{
    _gridModel = nullptr;
    _gridModelStorage.reset();
    _grid          = nullptr;
    _searchEdit    = nullptr;
    _subtitleLabel = nullptr;
    _root          = nullptr;
    _rootStorage.reset();
    _dxHost.Detach();
    _shortcutManager = nullptr;
    _ownerWindow     = nullptr;
    _searchQuery.clear();
    _selectedRowId.reset();
    _hWnd.reset();
}

void ShortcutsWindow::BuildUi()
{
    if (_root)
    {
        return;
    }

    _rootStorage = std::make_unique<Panel>();
    _root        = _rootStorage.get();

    _subtitleLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_CMD_DESC_SHORTCUTS));
    _subtitleLabel->SetFontRole(FontRole::Small);
    _subtitleLabel->SetMultiline(false);

    _searchEdit = _root->AddChild<TextField>();
    _searchEdit->SetPlaceholder(LoadStringResource(nullptr, IDS_SHORTCUTS_SEARCH_CUE));
    _searchEdit->SetAccessibleName(LoadStringResource(nullptr, IDS_SHORTCUTS_SEARCH_CUE));
    _searchEdit->SetOnTextChanged([this](std::wstring_view text)
    {
        const std::wstring trimmed = std::wstring(TrimWhitespace(text));
        if (trimmed == _searchQuery)
        {
            return;
        }

        _searchQuery = trimmed;
        RebuildRows();
    });

    _grid = _root->AddChild<Grid>();
    _grid->SetDelegate(this);
    _grid->SetSelectionMode(GridSelectionMode::Single);
    _grid->SetHeaderHeightDip(30.0f);
    _grid->SetRowHeightDip(48.0f);
    _grid->SetLineClamp(2u);

    _gridModelStorage = std::make_unique<ShortcutsGridModel>();
    _gridModel        = _gridModelStorage.get();
    static_cast<void>(_gridModel->SetGroupCollapsed(kGroupStableIdFunctionBar, _shortcuts.functionBarCollapsed));
    static_cast<void>(_gridModel->SetGroupCollapsed(kGroupStableIdFolderView, _shortcuts.folderViewCollapsed));
    _grid->SetModel(_gridModel);
    if (_settings && _settings->shortcuts.has_value())
    {
        const auto layout = ConvertColumnLayout(_settings->shortcuts->gridLayout);
        if (! layout.empty())
        {
            _grid->ApplyColumnLayout(layout);
        }
        if (! _settings->shortcuts->sortColumnId.empty())
        {
            for (size_t columnIndex = 0u; columnIndex < _gridModel->GetColumnCount(); ++columnIndex)
            {
                if (_gridModel->GetColumn(columnIndex).id != _settings->shortcuts->sortColumnId)
                {
                    continue;
                }

                _grid->SetSortSpec(GridSortSpec{
                    .columnIndex = columnIndex,
                    .direction   = _settings->shortcuts->sortDescending ? SortDirection::Descending : SortDirection::Ascending,
                });
                break;
            }
        }
    }

    _dxHost.SetRoot(std::move(_rootStorage));
}

void ShortcutsWindow::ApplyTheme() noexcept
{
    _dxHost.SetTheme(MakeAppThemeDxPalette(_theme));
    if (_hWnd)
    {
        ApplyTitleBarTheme(_hWnd.get(), _theme, GetActiveWindow() == _hWnd.get());
        ApplyWindowBackdropTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool);
    }
}

void ShortcutsWindow::Layout() noexcept
{
    if (! _root)
    {
        return;
    }

    const D2D1_RECT_F bounds = _dxHost.GetClientBoundsDip();
    _root->SetBounds(bounds);

    const float outer        = 16.0f;
    const float gap          = 12.0f;
    const float titleHeight  = 20.0f;
    const float searchHeight = 36.0f;

    float top = outer;
    if (_subtitleLabel)
    {
        _subtitleLabel->SetBounds(D2D1::RectF(outer, top, bounds.right - outer, top + titleHeight));
        top += titleHeight + gap;
    }

    if (_searchEdit)
    {
        _searchEdit->SetBounds(D2D1::RectF(outer, top, bounds.right - outer, top + searchHeight));
        top += searchHeight + gap;
    }

    if (_grid)
    {
        _grid->SetBounds(D2D1::RectF(outer, top, bounds.right - outer, bounds.bottom - outer));
    }
}

void ShortcutsWindow::RebuildRows() noexcept
{
    if (! _gridModel || ! _shortcutManager)
    {
        return;
    }

    const std::optional<uint64_t> previousSelectedRowId = _selectedRowId;
    std::vector<ShortcutRow> rows;
    rows.reserve(_shortcuts.functionBar.size() + _shortcuts.folderView.size());

    const auto addScope =
        [&](const std::vector<Common::Settings::ShortcutBinding>& bindings, const std::vector<uint32_t>& conflicts, uint64_t groupStableId) noexcept
    {
        std::unordered_map<uint32_t, std::vector<size_t>> chordToRows;
        const size_t scopeStart = rows.size();

        for (size_t bindingIndex = 0u; bindingIndex < bindings.size(); ++bindingIndex)
        {
            const auto& binding = bindings[bindingIndex];
            if (binding.commandId.empty() || ShortcutIds::IsUnassignedCommandId(binding.commandId))
            {
                continue;
            }

            ShortcutRow row;
            row.stableId      = MakeShortcutStableRowId(groupStableId, bindingIndex);
            row.vk            = binding.vk;
            row.modifiers     = binding.modifiers;
            row.groupStableId = groupStableId;
            row.commandId     = binding.commandId;
            row.hasConflict   = IsConflictChord(ShortcutManager::MakeChordKey(binding.vk, binding.modifiers), conflicts);
            row.keyText       = FormatChordText(binding.vk, binding.modifiers);

            const std::wstring displayName = GetCommandDisplayName(binding.commandId);
            const std::wstring description = GetCommandDescription(binding.commandId);
            row.commandText                = displayName;
            if (! description.empty())
            {
                row.commandText.append(L"\n");
                row.commandText.append(description);
            }
            row.tooltipText = displayName;
            if (! description.empty())
            {
                row.tooltipText.append(L"\n");
                row.tooltipText.append(description);
            }

            const uint32_t chordKey = ShortcutManager::MakeChordKey(binding.vk, binding.modifiers);
            rows.push_back(std::move(row));
            chordToRows[chordKey].push_back(rows.size() - 1u);
        }

        for (const auto& [_, indices] : chordToRows)
        {
            if (indices.size() <= 1u)
            {
                continue;
            }

            for (size_t i = 0; i < indices.size(); ++i)
            {
                const size_t rowIndex   = indices[i];
                const size_t otherIndex = indices[(i + 1u) % indices.size()];
                if (rowIndex >= rows.size() || otherIndex >= rows.size())
                {
                    continue;
                }

                const std::wstring conflictText =
                    FormatStringResource(nullptr, IDS_FMT_SHORTCUT_CONFLICT, rows[otherIndex].tooltipText, rows[rowIndex].keyText);
                rows[rowIndex].tooltipText.append(L"\n");
                rows[rowIndex].tooltipText.append(conflictText);
            }
        }

        const std::wstring_view query = TrimWhitespace(_searchQuery);
        if (query.empty())
        {
            return;
        }

        std::vector<ShortcutRow> filtered;
        filtered.reserve(rows.size());
        for (size_t rowIndex = scopeStart; rowIndex < rows.size(); ++rowIndex)
        {
            const ShortcutRow& row = rows[rowIndex];
            if (ContainsNoCase(row.commandText, query) || ContainsNoCase(row.keyText, query) || ContainsNoCase(row.tooltipText, query))
            {
                filtered.push_back(row);
            }
        }

        rows.resize(scopeStart);
        rows.insert(rows.end(), std::make_move_iterator(filtered.begin()), std::make_move_iterator(filtered.end()));
    };

    addScope(_shortcuts.functionBar, _shortcutManager->GetFunctionBarConflicts(), kGroupStableIdFunctionBar);
    addScope(_shortcuts.folderView, _shortcutManager->GetFolderViewConflicts(), kGroupStableIdFolderView);

    _gridModel->SetRows(std::move(rows));
    if (_grid)
    {
        _gridModel->SortRows(_grid->GetSortSpec());
    }
    _grid->NotifyDataChanged();
    if (_gridModel->GetRowCount() == 0u && previousSelectedRowId.has_value())
    {
        _selectedRowId = previousSelectedRowId;
    }
    RestoreSelection();
    _dxHost.Invalidate();
}

void ShortcutsWindow::RestoreSelection() noexcept
{
    if (! _grid || ! _gridModel)
    {
        return;
    }

    if (_selectedRowId.has_value())
    {
        if (const std::optional<size_t> rowIndex = _gridModel->FindRowByStableId(_selectedRowId.value()); rowIndex.has_value())
        {
            if (_grid->RequestSelectRow(rowIndex.value(), 0u))
            {
                return;
            }
        }
    }

    if (_gridModel->GetRowCount() != 0u)
    {
        static_cast<void>(_grid->RequestSelectRow(0u, 0u));
    }
}

void ShortcutsWindow::ResizeWindowToDefault(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    HMONITOR monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    MONITORINFO info{};
    info.cbSize = sizeof(info);
    if (! GetMonitorInfoW(monitor, &info))
    {
        return;
    }

    const RECT& rcWork   = info.rcWork;
    const int workWidth  = std::max(0L, rcWork.right - rcWork.left);
    const int workHeight = std::max(0L, rcWork.bottom - rcWork.top);
    const UINT dpi       = GetDpiForWindow(hwnd);
    const int width      = std::min(workWidth, MulDiv(960, static_cast<int>(dpi), 96));
    const int height     = std::min(workHeight, MulDiv(680, static_cast<int>(dpi), 96));

    SetWindowPos(hwnd, nullptr, rcWork.left, rcWork.top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
}

#ifdef ENABLE_TESTS
bool ShortcutsWindow::DebugGetSnapshot(ShortcutsWindowDebugSnapshot& out) const noexcept
{
    out = {};

    const HWND hwnd = GetHwnd();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    out.usesDxUiHost            = true;
    out.visibleChildWindowCount = CountVisibleChildWindowsLocal(hwnd);
    out.themeDark               = _theme.dark;
    out.themeHighContrast       = _theme.highContrast;
    out.themeRainbow            = _theme.menu.rainbowMode;
    out.rowCount                = _gridModel ? _gridModel->GetRowCount() : 0u;
    out.groupCount              = _gridModel ? _gridModel->GetGroupCount() : 0u;
    out.functionBarCollapsed    = _gridModel && _gridModel->IsGroupCollapsed(kGroupStableIdFunctionBar);
    out.folderViewCollapsed     = _gridModel && _gridModel->IsGroupCollapsed(kGroupStableIdFolderView);
    out.collapsedGroupCount     = static_cast<size_t>(out.functionBarCollapsed) + static_cast<size_t>(out.folderViewCollapsed);
    if (_gridModel)
    {
        out.rowKeyTexts.reserve(_gridModel->GetRowCount());
        for (size_t rowIndex = 0u; rowIndex < _gridModel->GetRowCount(); ++rowIndex)
        {
            if (const ShortcutRow* const row = _gridModel->GetRow(rowIndex))
            {
                out.rowKeyTexts.push_back(row->keyText);
            }
        }
    }
    if (_grid)
    {
        const auto metrics          = _grid->GetVisibleWorkMetrics();
        out.visibleRowCount         = static_cast<size_t>(metrics.visibleRowCount);
        out.visibleGroupHeaderCount = static_cast<size_t>(metrics.visibleGroupHeaderCount);
        out.visibleColumnCount      = metrics.visibleColumnCount;
        out.visibleCellCount        = static_cast<size_t>(metrics.visibleCellCount);
        out.verticalScrollDip       = metrics.verticalScrollDip;
        out.horizontalScrollDip     = metrics.horizontalScrollDip;
        out.hasVerticalScrollbar    = metrics.hasVerticalScrollbar;
        out.hasHorizontalScrollbar  = metrics.hasHorizontalScrollbar;
        out.firstColumnHeaderRect   = _grid->GetVisibleDisplayColumnHeaderRect(0u).value_or(D2D1::RectF());
        out.secondColumnHeaderRect  = _grid->GetVisibleDisplayColumnHeaderRect(1u).value_or(D2D1::RectF());
        const auto columnLayout     = _grid->CaptureColumnLayout();
        out.displayColumnIds.clear();
        out.displayColumnWidthsDip.clear();
        out.displayColumnIds.reserve(columnLayout.size());
        out.displayColumnWidthsDip.reserve(columnLayout.size());
        for (const auto& entry : columnLayout)
        {
            out.displayColumnIds.push_back(entry.columnId);
            out.displayColumnWidthsDip.push_back(entry.widthDip);
        }
        if (! columnLayout.empty())
        {
            out.firstDisplayColumnId = columnLayout[0].columnId;
        }
        if (columnLayout.size() > 1u)
        {
            out.secondDisplayColumnId = columnLayout[1].columnId;
        }
        if (const auto firstVisibleRow = _grid->GetVisibleRowAt(0u); firstVisibleRow.has_value())
        {
            out.firstVisibleRowIndex = firstVisibleRow.value();
            if (const ShortcutRow* const row = _gridModel ? _gridModel->GetRow(firstVisibleRow.value()) : nullptr; row)
            {
                out.firstVisibleRowName    = row->commandText;
                out.firstVisibleRowKeyText = row->keyText;
            }
        }
        if (const std::optional<size_t> firstVisibleRow = _grid->GetVisibleRowAt(0u); firstVisibleRow.has_value())
        {
            out.firstVisibleRowRect           = _grid->GetVisibleRowRect(firstVisibleRow.value()).value_or(D2D1::RectF());
            out.firstVisibleCommandCellRect   = _grid->GetVisibleCellRect(firstVisibleRow.value(), 0u).value_or(D2D1::RectF());
            out.firstVisibleKeyCellRect       = _grid->GetVisibleCellRect(firstVisibleRow.value(), 1u).value_or(D2D1::RectF());
            out.firstVisibleCommandLayoutRect = _grid->GetCellLayoutMetrics(_dxHost, firstVisibleRow.value(), 0u).cellRect;
            out.firstVisibleKeyLayoutRect     = _grid->GetCellLayoutMetrics(_dxHost, firstVisibleRow.value(), 1u).cellRect;
        }
        if (_gridModel)
        {
            const std::optional<size_t> selectedRow = _grid->GetPrimarySelectedRow();
            if (selectedRow.has_value())
            {
                out.selectedRowName = _gridModel->GetRowAccessibleName(selectedRow.value());
                if (const ShortcutRow* const row = _gridModel->GetRow(selectedRow.value()))
                {
                    out.selectedRowKeyText = row->keyText;
                }
                out.selectedRowKeyCellRect        = _grid->GetVisibleCellRect(selectedRow.value(), 1u).value_or(D2D1::RectF());
                out.selectedRowCommandCellHasIcon = _grid->GetCellLayoutMetrics(_dxHost, selectedRow.value(), 0u).hasIcon;
                RedSalamander::DxUi::GridDebugRowVisualState rowVisualState{};
                if (_grid->DebugGetRowVisualState(_dxHost.GetTheme(), selectedRow.value(), rowVisualState))
                {
                    out.selectedRowFillArgb    = rowVisualState.fillArgb;
                    out.selectedRowTextArgb    = rowVisualState.textArgb;
                    out.selectedRowUsesRainbow = rowVisualState.usesRainbow;
                }
            }
        }
    }
    out.renderCount        = _dxHost.DebugGetRenderCount();
    out.resizeCount        = _dxHost.DebugGetResizeCount();
    out.resizeFailureCount = _dxHost.DebugGetResizeFailureCount();
    out.searchText         = _searchEdit ? _searchEdit->GetText() : _searchQuery;
    out.hasTooltip         = _dxHost.HasTooltip();
    out.tooltipBounds      = _dxHost.DebugGetTooltipBoundsDip();
    out.tooltipText        = std::wstring(_dxHost.GetTooltipText());
    if (RedSalamander::DxUi::Control* const focusedControl = _dxHost.GetFocusControl())
    {
        if (focusedControl == _searchEdit)
        {
            out.focusTarget = ShortcutsWindowDebugFocusTarget::SearchField;
        }
        else if (focusedControl == _grid)
        {
            out.focusTarget = ShortcutsWindowDebugFocusTarget::Grid;
        }
    }
    if (_grid)
    {
        const auto sortSpec = _grid->GetSortSpec();
        out.sortDirection   = sortSpec.direction == RedSalamander::DxUi::SortDirection::None ? 0xFFu : static_cast<uint8_t>(sortSpec.direction);
        out.sortColumnIndex = sortSpec.direction == RedSalamander::DxUi::SortDirection::None ? 0xFFu : static_cast<uint8_t>(sortSpec.columnIndex);
    }
    return true;
}

bool ShortcutsWindow::DebugScrollByWheelDetents(const int detents) noexcept
{
    const HWND hwnd = GetHwnd();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    if (! _grid || detents == 0)
    {
        return detents == 0;
    }

    _dxHost.SetFocusControl(_grid);
    const float wheelDelta = detents > 0 ? static_cast<float>(WHEEL_DELTA) : -static_cast<float>(WHEEL_DELTA);
    const int stepCount    = detents > 0 ? detents : -detents;
    for (int remaining = stepCount; remaining > 0; --remaining)
    {
        if (! _grid->OnMouseWheel(_dxHost, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0u))
        {
            return false;
        }
    }
    return true;
}

bool ShortcutsWindow::DebugSelectRow(const size_t rowIndex) noexcept
{
    const HWND hwnd = GetHwnd();
    if (! hwnd || IsWindow(hwnd) == FALSE || ! _grid || ! _gridModel)
    {
        return false;
    }

    if (rowIndex >= _gridModel->GetRowCount())
    {
        return false;
    }

    _dxHost.SetFocusControl(_grid);
    return _grid->RequestSelectRow(rowIndex, 0u);
}

bool ShortcutsWindow::DebugSetSearchText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetHwnd();
    if (! hwnd || IsWindow(hwnd) == FALSE || ! _searchEdit)
    {
        return false;
    }

    const std::wstring trimmed = std::wstring(TrimWhitespace(text));
    _searchEdit->SetTextAndNotify(trimmed);
    _dxHost.SyncTextInput(_searchEdit);
    if (trimmed != _searchQuery)
    {
        _searchQuery = trimmed;
        RebuildRows();
    }
    _dxHost.Invalidate();
    return true;
}

bool ShortcutsWindow::DebugFocusSearch() noexcept
{
    const HWND hwnd = GetHwnd();
    if (! hwnd || IsWindow(hwnd) == FALSE || ! _searchEdit)
    {
        return false;
    }

    _dxHost.SetFocusControl(_searchEdit);
    return true;
}

bool ShortcutsWindow::DebugFocusGrid() noexcept
{
    const HWND hwnd = GetHwnd();
    if (! hwnd || IsWindow(hwnd) == FALSE || ! _grid)
    {
        return false;
    }

    _dxHost.SetFocusControl(_grid);
    return true;
}

bool ShortcutsWindow::DebugCycleGridSortByColumn(const size_t columnIndex) noexcept
{
    const HWND hwnd = GetHwnd();
    if (! hwnd || IsWindow(hwnd) == FALSE || ! _grid || ! _gridModel || columnIndex >= _gridModel->GetColumnCount())
    {
        return false;
    }

    const GridSortSpec beforeSort = _grid->GetSortSpec();
    GridSortSpec nextSort{};
    nextSort.columnIndex = columnIndex;
    nextSort.direction   = beforeSort.columnIndex == columnIndex ? NextSortDirection(beforeSort.direction) : SortDirection::Ascending;
    _grid->SetSortSpec(nextSort);
    OnGridSortRequested(nextSort);
    const GridSortSpec afterSort = _grid->GetSortSpec();
    return beforeSort.columnIndex != afterSort.columnIndex || beforeSort.direction != afterSort.direction;
}

bool ShortcutsWindow::DebugApplyGridLayout(const std::vector<Common::Settings::GridColumnLayoutEntry>& layout) noexcept
{
    const HWND hwnd = GetHwnd();
    if (! hwnd || IsWindow(hwnd) == FALSE || ! _grid)
    {
        return false;
    }

    const auto converted = ConvertColumnLayout(layout);
    if (converted.empty())
    {
        return false;
    }

    _grid->ApplyColumnLayout(converted);
    _dxHost.Invalidate();
    return true;
}

bool ShortcutsWindow::DebugSetGroupCollapsed(const size_t groupIndex, const bool collapsed) noexcept
{
    const HWND hwnd = GetHwnd();
    if (! hwnd || IsWindow(hwnd) == FALSE || ! _grid || ! _gridModel)
    {
        return false;
    }

    if (groupIndex >= _gridModel->GetGroupCount())
    {
        return false;
    }

    const GridGroupDesc group = _gridModel->GetGroup(groupIndex);
    if (! _gridModel->SetGroupCollapsed(group.stableId, collapsed))
    {
        return false;
    }

    _grid->NotifyDataChanged();
    _dxHost.Invalidate();
    return true;
}
#endif

} // namespace

void ShowShortcutsWindow(HWND owner,
                         Common::Settings::Settings& settings,
                         const Common::Settings::ShortcutsSettings& shortcuts,
                         const ShortcutManager& shortcutManager,
                         const AppTheme& theme) noexcept
{
    if (g_shortcutsWindow && g_shortcutsWindow->GetHwnd())
    {
        g_shortcutsWindow->UpdateData(shortcuts, shortcutManager);
        g_shortcutsWindow->UpdateTheme(theme);

        const HWND hwnd = g_shortcutsWindow->GetHwnd();
        if (IsIconic(hwnd))
        {
            ShowWindow(hwnd, SW_RESTORE);
        }
        else
        {
            ShowWindow(hwnd, SW_SHOW);
        }
        SetForegroundWindow(hwnd);
        return;
    }

    auto window = std::make_unique<ShortcutsWindow>();
    if (window->Create(owner, settings, shortcuts, shortcutManager, theme))
    {
        static_cast<void>(window.release());
    }
}

void UpdateShortcutsWindowTheme(const AppTheme& theme) noexcept
{
    if (! g_shortcutsWindow)
    {
        return;
    }

    g_shortcutsWindow->UpdateTheme(theme);
}

void UpdateShortcutsWindowData(const Common::Settings::ShortcutsSettings& shortcuts, const ShortcutManager& shortcutManager) noexcept
{
    if (! g_shortcutsWindow)
    {
        return;
    }

    g_shortcutsWindow->UpdateData(shortcuts, shortcutManager);
}

HWND GetShortcutsWindowHandle() noexcept
{
    return g_shortcutsWindow ? g_shortcutsWindow->GetHwnd() : nullptr;
}

#ifdef ENABLE_TESTS
bool DebugGetShortcutsWindowSnapshot(ShortcutsWindowDebugSnapshot& out) noexcept
{
    return g_shortcutsWindow && g_shortcutsWindow->DebugGetSnapshot(out);
}

bool DebugScrollShortcutsWindowByWheelDetents(const int detents) noexcept
{
    return g_shortcutsWindow && g_shortcutsWindow->DebugScrollByWheelDetents(detents);
}

bool DebugSelectShortcutsWindowRow(const size_t rowIndex) noexcept
{
    return g_shortcutsWindow && g_shortcutsWindow->DebugSelectRow(rowIndex);
}

bool DebugSetShortcutsWindowSearchText(std::wstring_view text) noexcept
{
    return g_shortcutsWindow && g_shortcutsWindow->DebugSetSearchText(text);
}

bool DebugFocusShortcutsWindowSearch() noexcept
{
    return g_shortcutsWindow && g_shortcutsWindow->DebugFocusSearch();
}

bool DebugFocusShortcutsWindowGrid() noexcept
{
    return g_shortcutsWindow && g_shortcutsWindow->DebugFocusGrid();
}

bool DebugCycleShortcutsWindowGridSortByColumn(const size_t columnIndex) noexcept
{
    return g_shortcutsWindow && g_shortcutsWindow->DebugCycleGridSortByColumn(columnIndex);
}

bool DebugApplyShortcutsWindowGridLayout(const std::vector<Common::Settings::GridColumnLayoutEntry>& layout) noexcept
{
    return g_shortcutsWindow && g_shortcutsWindow->DebugApplyGridLayout(layout);
}

bool DebugSetShortcutsWindowGroupCollapsed(const size_t groupIndex, const bool collapsed) noexcept
{
    return g_shortcutsWindow && g_shortcutsWindow->DebugSetGroupCollapsed(groupIndex, collapsed);
}
#endif
