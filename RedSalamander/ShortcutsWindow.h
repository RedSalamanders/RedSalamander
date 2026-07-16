#pragma once

#include "AppTheme.h"
#include "SettingsStore.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

class ShortcutManager;

void ShowShortcutsWindow(HWND owner,
                         Common::Settings::Settings& settings,
                         const Common::Settings::ShortcutsSettings& shortcuts,
                         const ShortcutManager& shortcutManager,
                         const AppTheme& theme) noexcept;

void UpdateShortcutsWindowTheme(const AppTheme& theme) noexcept;

void UpdateShortcutsWindowData(const Common::Settings::ShortcutsSettings& shortcuts, const ShortcutManager& shortcutManager) noexcept;

[[nodiscard]] HWND GetShortcutsWindowHandle() noexcept;

#ifdef ENABLE_TESTS
enum class ShortcutsWindowDebugFocusTarget : uint8_t
{
    None = 0u,
    SearchField,
    Grid,
};

struct ShortcutsWindowDebugSnapshot
{
    bool usesDxUiHost                         = false;
    bool hasTooltip                           = false;
    bool themeDark                            = false;
    bool themeHighContrast                    = false;
    bool themeRainbow                         = false;
    size_t visibleChildWindowCount            = 0u;
    size_t rowCount                           = 0u;
    size_t groupCount                         = 0u;
    size_t collapsedGroupCount                = 0u;
    size_t visibleRowCount                    = 0u;
    size_t visibleColumnCount                 = 0u;
    size_t visibleCellCount                   = 0u;
    size_t visibleGroupHeaderCount            = 0u;
    float verticalScrollDip                   = 0.0f;
    float horizontalScrollDip                 = 0.0f;
    bool hasVerticalScrollbar                 = false;
    bool hasHorizontalScrollbar               = false;
    uint64_t renderCount                      = 0u;
    uint64_t resizeCount                      = 0u;
    uint64_t resizeFailureCount               = 0u;
    D2D1_RECT_F firstColumnHeaderRect         = D2D1::RectF();
    D2D1_RECT_F secondColumnHeaderRect        = D2D1::RectF();
    D2D1_RECT_F firstVisibleRowRect           = D2D1::RectF();
    D2D1_RECT_F firstVisibleCommandLayoutRect = D2D1::RectF();
    D2D1_RECT_F firstVisibleKeyLayoutRect     = D2D1::RectF();
    D2D1_RECT_F firstVisibleCommandCellRect   = D2D1::RectF();
    D2D1_RECT_F firstVisibleKeyCellRect       = D2D1::RectF();
    D2D1_RECT_F selectedRowKeyCellRect        = D2D1::RectF();
    D2D1_RECT_F tooltipBounds                 = D2D1::RectF();
    uint32_t selectedRowFillArgb              = 0u;
    uint32_t selectedRowTextArgb              = 0u;
    bool selectedRowUsesRainbow               = false;
    bool selectedRowCommandCellHasIcon        = false;
    bool functionBarCollapsed                 = false;
    bool folderViewCollapsed                  = false;
    std::wstring searchText;
    std::wstring selectedRowName;
    std::wstring selectedRowKeyText;
    std::wstring firstVisibleRowName;
    std::wstring firstVisibleRowKeyText;
    size_t firstVisibleRowIndex = static_cast<size_t>(-1);
    std::wstring firstDisplayColumnId;
    std::wstring secondDisplayColumnId;
    std::wstring tooltipText;
    std::vector<std::wstring> rowKeyTexts;
    std::vector<std::wstring> displayColumnIds;
    std::vector<float> displayColumnWidthsDip;
    ShortcutsWindowDebugFocusTarget focusTarget = ShortcutsWindowDebugFocusTarget::None;
    uint8_t sortColumnIndex                     = 0xFFu;
    uint8_t sortDirection                       = 0xFFu;
};

[[nodiscard]] bool DebugGetShortcutsWindowSnapshot(ShortcutsWindowDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugScrollShortcutsWindowByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugSelectShortcutsWindowRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugSetShortcutsWindowSearchText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugFocusShortcutsWindowSearch() noexcept;
[[nodiscard]] bool DebugFocusShortcutsWindowGrid() noexcept;
[[nodiscard]] bool DebugCopyShortcutsWindowSelection() noexcept;
[[nodiscard]] bool DebugCycleShortcutsWindowGridSortByColumn(size_t columnIndex) noexcept;
[[nodiscard]] bool DebugApplyShortcutsWindowGridLayout(const std::vector<Common::Settings::GridColumnLayoutEntry>& layout) noexcept;
[[nodiscard]] bool DebugSetShortcutsWindowGroupCollapsed(size_t groupIndex, bool collapsed) noexcept;
#endif
