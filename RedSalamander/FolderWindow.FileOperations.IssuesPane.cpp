#include "FolderWindow.FileOperations.IssuesPane.h"

#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "Helpers.h"
#include "WindowMaximizeBehavior.h"
#include "WindowSizing.h"

#include <algorithm>
#include <format>
#include <memory>
#include <unordered_set>
#include <vector>

namespace
{
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridRowStyle;
using RedSalamander::DxUi::GridRowTone;
using RedSalamander::DxUi::GridSortSpec;
using RedSalamander::DxUi::IDxGridDelegate;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::SortDirection;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::WindowHost;

constexpr wchar_t kFileOperationsIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
constexpr UINT_PTR kRefreshTimerId                     = 1;
constexpr UINT kRefreshTimerIntervalMs                 = 750;

int DipsToPixels(int dip, UINT dpi) noexcept
{
    return MulDiv(dip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
}

[[nodiscard]] std::wstring FormatTimeText(const SYSTEMTIME& localTime)
{
    return std::format(L"{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}",
                       static_cast<unsigned>(localTime.wYear),
                       static_cast<unsigned>(localTime.wMonth),
                       static_cast<unsigned>(localTime.wDay),
                       static_cast<unsigned>(localTime.wHour),
                       static_cast<unsigned>(localTime.wMinute),
                       static_cast<unsigned>(localTime.wSecond),
                       static_cast<unsigned>(localTime.wMilliseconds));
}

[[nodiscard]] std::wstring FormatStatusText(HRESULT hr) noexcept
{
    return FormatHResultMessage(hr);
}

[[nodiscard]] UINT OperationStringId(FileSystemOperation operation) noexcept
{
    switch (operation)
    {
        case FILESYSTEM_COPY: return IDS_CMD_COPY;
        case FILESYSTEM_MOVE: return IDS_CMD_MOVE;
        case FILESYSTEM_DELETE: return IDS_CMD_DELETE;
        case FILESYSTEM_RENAME: return IDS_CMD_RENAME;
        default: return IDS_FILEOPS_ISSUES_OPERATION_UNKNOWN;
    }
}

struct IssuesRow
{
    FolderWindow::FileOperationState::DiagnosticSeverity severity = FolderWindow::FileOperationState::DiagnosticSeverity::Warning;
    uint64_t stableId                                             = 0;
    uint64_t taskId                                               = 0;
    std::wstring timeText;
    std::wstring taskText;
    std::wstring operationText;
    std::wstring severityText;
    std::wstring statusText;
    std::wstring statusTextDetail;
    std::wstring categoryText;
    std::wstring messageText;
    std::wstring sourcePathText;
    std::wstring destinationPathText;
};

[[nodiscard]] bool RowsEqual(const std::vector<IssuesRow>& lhs, const std::vector<IssuesRow>& rhs) noexcept
{
    if (lhs.size() != rhs.size())
    {
        return false;
    }

    for (size_t i = 0; i < lhs.size(); ++i)
    {
        if (lhs[i].stableId != rhs[i].stableId || lhs[i].statusText != rhs[i].statusText || lhs[i].messageText != rhs[i].messageText)
        {
            return false;
        }
    }
    return true;
}


class IssuesGridModel final : public IDxGridModel
{
public:
    explicit IssuesGridModel(const AppTheme& theme) : _theme(&theme)
    {
        _columns = {
            {L"time", LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_TIME), 170.0f, 120.0f, RedSalamander::DxUi::GridColumnKind::Text, true, false},
            {L"task", LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_TASK), 70.0f, 56.0f, RedSalamander::DxUi::GridColumnKind::Text, true, false},
            {L"operation", LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_OPERATION), 90.0f, 80.0f, RedSalamander::DxUi::GridColumnKind::Text, true, false},
            {L"severity", LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_SEVERITY), 86.0f, 72.0f, RedSalamander::DxUi::GridColumnKind::Text, true, false},
            {L"hresult", LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_HRESULT), 104.0f, 100.0f, RedSalamander::DxUi::GridColumnKind::Text, true, false},
            {L"status", LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_STATUS_TEXT), 220.0f, 120.0f, RedSalamander::DxUi::GridColumnKind::Text, true, true},
            {L"category", LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_CATEGORY), 140.0f, 100.0f, RedSalamander::DxUi::GridColumnKind::Text, true, true},
            {L"message", LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_MESSAGE), 300.0f, 140.0f, RedSalamander::DxUi::GridColumnKind::Text, true, true},
            {L"source", LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_SOURCE), 300.0f, 160.0f, RedSalamander::DxUi::GridColumnKind::Text, true, true},
            {L"destination",
             LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_DESTINATION),
             300.0f,
             160.0f,
             RedSalamander::DxUi::GridColumnKind::Text,
             true,
             true},
        };
    }

    void SetRows(const std::vector<IssuesRow>* rows) noexcept
    {
        _rows = rows;
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rows ? _rows->size() : 0u;
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
        if (! _rows || rowIndex >= _rows->size())
        {
            outCell = {};
            return;
        }

        const IssuesRow& row = _rows->at(rowIndex);
        switch (columnIndex)
        {
            case 0: outCell.text = row.timeText; break;
            case 1: outCell.text = row.taskText; break;
            case 2: outCell.text = row.operationText; break;
            case 3: outCell.text = row.severityText; break;
            case 4: outCell.text = row.statusText; break;
            case 5: outCell.text = row.statusTextDetail; break;
            case 6: outCell.text = row.categoryText; break;
            case 7:
                outCell.text      = row.messageText;
                outCell.multiline = true;
                break;
            case 8:
                outCell.text      = row.sourcePathText;
                outCell.multiline = true;
                break;
            case 9:
                outCell.text      = row.destinationPathText;
                outCell.multiline = true;
                break;
            default: outCell = {}; break;
        }
    }

    [[nodiscard]] GridRowStyle GetRowStyle(size_t rowIndex) const override
    {
        GridRowStyle style{};
        if (! _rows || rowIndex >= _rows->size())
        {
            return style;
        }

        const IssuesRow& row = _rows->at(rowIndex);
        if (_theme && _theme->menu.rainbowMode)
        {
            style.rainbowSeed = ! row.messageText.empty() ? row.messageText : row.categoryText;
            return style;
        }

        style.tone = (row.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Error) ? GridRowTone::Error : GridRowTone::Warning;
        return style;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return (_rows && rowIndex < _rows->size()) ? _rows->at(rowIndex).stableId : 0u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (! _rows)
        {
            return std::nullopt;
        }

        for (size_t rowIndex = 0; rowIndex < _rows->size(); ++rowIndex)
        {
            if (_rows->at(rowIndex).stableId == rowId)
            {
                return rowIndex;
            }
        }
        return std::nullopt;
    }

private:
    const AppTheme* _theme              = nullptr;
    const std::vector<IssuesRow>* _rows = nullptr;
    std::vector<GridColumnDesc> _columns;
};

class FileOperationsIssuesPaneState final : public IDxGridDelegate
{
public:
    FileOperationsIssuesPaneState()                                                = default;
    FileOperationsIssuesPaneState(const FileOperationsIssuesPaneState&)            = delete;
    FileOperationsIssuesPaneState(FileOperationsIssuesPaneState&&)                 = delete;
    FileOperationsIssuesPaneState& operator=(const FileOperationsIssuesPaneState&) = delete;
    FileOperationsIssuesPaneState& operator=(FileOperationsIssuesPaneState&&)      = delete;

    FolderWindow::FileOperationState* fileOps = nullptr;
    FolderWindow* folderWindow                = nullptr;
    std::weak_ptr<void> hostLifetime;

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnGridSortRequested(const GridSortSpec& sortSpec) override
    {
        _sortSpec = sortSpec;
        SortRows();
        if (_grid)
        {
            _grid->SetSortSpec(_sortSpec);
            _grid->NotifyDataChanged();
        }
        SaveViewState();
        _dxHost.Invalidate();
    }

#ifdef ENABLE_TESTS
    void FillSelfTestSnapshot(FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) const noexcept;
    [[nodiscard]] bool SelfTestHitTestGridPoint(D2D1_POINT_2F pointDip, FileOperationsIssuesPane::SelfTestGridHit& outHit) const noexcept;
    [[nodiscard]] bool SelfTestSelectTask(uint64_t taskId) noexcept;
    [[nodiscard]] bool SelfTestFocusGrid() noexcept;
    [[nodiscard]] bool SelfTestRefresh(bool force) noexcept;
    [[nodiscard]] bool SelfTestScrollByWheelDetents(int detents) noexcept;
#endif

private:
    [[nodiscard]] bool HasLiveOwnerState() const noexcept
    {
        return fileOps && ! hostLifetime.expired();
    }

    LRESULT OnCreate(HWND hwnd) noexcept;
    LRESULT OnSize(HWND hwnd, UINT width, UINT height) noexcept;
    LRESULT OnTimer(HWND hwnd, UINT_PTR timerId) noexcept;
    LRESULT OnMove(HWND hwnd) noexcept;
    LRESULT OnExitSizeMove(HWND hwnd) noexcept;
    LRESULT OnShowWindow(HWND hwnd, BOOL visible) noexcept;
    LRESULT OnClose(HWND hwnd) noexcept;
    LRESULT OnDpiChanged(HWND hwnd, UINT dpi, const RECT* suggested) noexcept;
    LRESULT OnThemeChanged(HWND hwnd) noexcept;
    LRESULT OnNcDestroy(HWND hwnd) noexcept;

    void ApplyTheme(HWND hwnd) noexcept;
    void RefreshRows(bool force) noexcept;
    std::vector<IssuesRow> BuildRows() const;
    void Layout() noexcept;
    void SortRows() noexcept;
    void ApplySavedViewState() noexcept;
    void SaveViewState() noexcept;
#ifdef ENABLE_TESTS
    [[nodiscard]] std::optional<size_t> FindRowIndexForTask(uint64_t taskId) const noexcept;
#endif

    UINT _dpi                  = USER_DEFAULT_SCREEN_DPI;
    bool _inThemeChange        = false;
    bool _inTitleBarThemeApply = false;
    size_t _dispatchDepth      = 0u;
    bool _deletePending        = false;
    AppTheme _theme{};
    WindowHost _dxHost;
    std::unique_ptr<Panel> _rootStorage;
    Panel* _root = nullptr;
    Grid* _grid  = nullptr;
    IssuesGridModel _model{_theme};
    GridSortSpec _sortSpec{};
    std::vector<IssuesRow> _rows;
    uint64_t _refreshGeneration = 0;
};

ATOM RegisterFileOperationsIssuesPaneClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = FileOperationsIssuesPaneState::WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon         = LoadIconW(instance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
    wc.hIconSm       = LoadIconW(instance, MAKEINTRESOURCEW(IDI_SMALL));
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kFileOperationsIssuesPaneClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}
} // namespace

void FileOperationsIssuesPaneState::ApplyTheme(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (folderWindow && hostLifetime.lock())
    {
        _theme = folderWindow->GetTheme();
        if (! _inTitleBarThemeApply)
        {
            _inTitleBarThemeApply = true;
            ApplyWindowChromeTheme(hwnd, _theme, WindowBackdropTarget::Tool, GetActiveWindow() == hwnd);
            _inTitleBarThemeApply = false;
        }
    }

    _dxHost.SetTheme(MakeFolderContentDxPalette(_theme));
}

std::vector<IssuesRow> FileOperationsIssuesPaneState::BuildRows() const
{
    std::vector<IssuesRow> rows;
    if (! fileOps || ! hostLifetime.lock())
    {
        return rows;
    }

    std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> liveDiagnostics;
    fileOps->CollectDiagnostics(liveDiagnostics);
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completed;
    fileOps->CollectCompletedTasks(completed);

    std::unordered_set<std::wstring> seenKeys;
    const auto appendIssue = [&](const FolderWindow::FileOperationState::TaskDiagnosticEntry& issue) noexcept
    {
        if (issue.severity != FolderWindow::FileOperationState::DiagnosticSeverity::Warning &&
            issue.severity != FolderWindow::FileOperationState::DiagnosticSeverity::Error)
        {
            return;
        }

        std::wstring key = std::format(L"{}|{}|{}|{}|{}|{}|{}|{}|{}|{}|{}|{}",
                                       static_cast<unsigned long long>(issue.taskId),
                                       static_cast<unsigned>(issue.localTime.wMilliseconds),
                                       static_cast<unsigned>(issue.operation),
                                       static_cast<unsigned long>(issue.status),
                                       static_cast<unsigned>(static_cast<unsigned char>(issue.severity)),
                                       issue.category,
                                       issue.message,
                                       issue.sourcePath,
                                       issue.destinationPath,
                                       issue.concurrencyMode,
                                       issue.storageType,
                                       issue.destinationStorageType);
        if (! seenKeys.insert(key).second)
        {
            return;
        }

        const uint64_t stableId = (static_cast<uint64_t>(StableHash32(key)) << 32u) ^ static_cast<uint64_t>(issue.taskId);

        IssuesRow row{};
        row.severity      = issue.severity;
        row.stableId      = stableId;
        row.taskId        = issue.taskId;
        row.timeText      = FormatTimeText(issue.localTime);
        row.taskText      = std::to_wstring(static_cast<unsigned long long>(issue.taskId));
        row.operationText = LoadStringResource(nullptr, OperationStringId(issue.operation));
        row.severityText  = LoadStringResource(
            nullptr, issue.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Error ? IDS_CAPTION_ERROR : IDS_CAPTION_WARNING);
        row.statusText          = std::format(L"0x{:08X}", static_cast<unsigned long>(issue.status));
        row.statusTextDetail    = FormatStatusText(issue.status);
        row.categoryText        = issue.category.empty() ? L"-" : issue.category;
        row.messageText         = issue.message.empty() ? L"-" : issue.message;
        row.sourcePathText      = issue.sourcePath.empty() ? L"-" : issue.sourcePath;
        row.destinationPathText = issue.destinationPath.empty() ? L"-" : issue.destinationPath;
        rows.push_back(std::move(row));
    };

    for (const auto& issue : liveDiagnostics)
    {
        appendIssue(issue);
    }
    for (const auto& task : completed)
    {
        for (const auto& issue : task.issueDiagnostics)
        {
            appendIssue(issue);
        }
    }
    return rows;
}

void FileOperationsIssuesPaneState::ApplySavedViewState() noexcept
{
    if (! fileOps || ! _grid)
    {
        return;
    }

    std::wstring savedSortColumnId;
    bool savedSortDescending = false;
    std::vector<Common::Settings::GridColumnLayoutEntry> savedGridLayout;
    if (! fileOps->TryGetIssuesPaneViewState(savedSortColumnId, savedSortDescending, savedGridLayout))
    {
        return;
    }

    if (! savedGridLayout.empty())
    {
        std::vector<RedSalamander::DxUi::GridColumnLayoutEntry> layout;
        layout.reserve(savedGridLayout.size());
        for (const auto& entry : savedGridLayout)
        {
            if (entry.columnId.empty())
            {
                continue;
            }

            layout.push_back(RedSalamander::DxUi::GridColumnLayoutEntry{
                .columnId     = entry.columnId,
                .displayIndex = static_cast<size_t>(entry.displayIndex),
                .widthDip     = entry.widthDip,
            });
        }

        if (! layout.empty())
        {
            _grid->ApplyColumnLayout(layout);
        }
    }

    if (! savedSortColumnId.empty())
    {
        for (size_t columnIndex = 0; columnIndex < _model.GetColumnCount(); ++columnIndex)
        {
            if (_model.GetColumn(columnIndex).id == savedSortColumnId)
            {
                _sortSpec.columnIndex = columnIndex;
                _sortSpec.direction   = savedSortDescending ? SortDirection::Descending : SortDirection::Ascending;
                _grid->SetSortSpec(_sortSpec);
                break;
            }
        }
    }
}

void FileOperationsIssuesPaneState::SaveViewState() noexcept
{
    if (! fileOps || ! _grid)
    {
        return;
    }

    std::wstring sortColumnId;
    if (_sortSpec.direction != SortDirection::None && _sortSpec.columnIndex < _model.GetColumnCount())
    {
        sortColumnId = _model.GetColumn(_sortSpec.columnIndex).id;
    }

    std::vector<Common::Settings::GridColumnLayoutEntry> savedGridLayout;
    const auto capturedLayout = _grid->CaptureColumnLayout();
    savedGridLayout.reserve(capturedLayout.size());
    for (const auto& entry : capturedLayout)
    {
        if (entry.columnId.empty())
        {
            continue;
        }

        savedGridLayout.push_back(Common::Settings::GridColumnLayoutEntry{
            .columnId     = entry.columnId,
            .displayIndex = static_cast<uint32_t>(entry.displayIndex),
            .widthDip     = entry.widthDip,
        });
    }

    fileOps->SaveIssuesPaneViewState(sortColumnId, _sortSpec.direction == SortDirection::Descending, savedGridLayout);
}

void FileOperationsIssuesPaneState::SortRows() noexcept
{
    if (_sortSpec.direction == SortDirection::None)
    {
        std::sort(_rows.begin(), _rows.end(), [](const IssuesRow& lhs, const IssuesRow& rhs) noexcept { return lhs.stableId < rhs.stableId; });
        return;
    }

    const auto getColumnText = [&](const IssuesRow& row) noexcept -> std::wstring_view
    {
        switch (_sortSpec.columnIndex)
        {
            case 0: return row.timeText;
            case 1: return row.taskText;
            case 2: return row.operationText;
            case 3: return row.severityText;
            case 4: return row.statusText;
            case 5: return row.statusTextDetail;
            case 6: return row.categoryText;
            case 7: return row.messageText;
            case 8: return row.sourcePathText;
            case 9: return row.destinationPathText;
            default: return row.messageText;
        }
    };

    std::sort(_rows.begin(),
              _rows.end(),
              [&](const IssuesRow& lhs, const IssuesRow& rhs) noexcept
    {
        const int cmp = OrdinalString::Compare(getColumnText(lhs), getColumnText(rhs), true);
        if (cmp == 0)
        {
            return lhs.stableId < rhs.stableId;
        }
        return _sortSpec.direction == SortDirection::Ascending ? (cmp < 0) : (cmp > 0);
    });
}

void FileOperationsIssuesPaneState::RefreshRows(bool force) noexcept
{
    std::vector<IssuesRow> rows = BuildRows();
    if (! force && RowsEqual(rows, _rows))
    {
        return;
    }

    _rows = std::move(rows);
    SortRows();
    _model.SetRows(&_rows);
    ++_refreshGeneration;
    if (_grid)
    {
        _grid->NotifyDataChanged();
    }
    _dxHost.Invalidate();
}

#ifdef ENABLE_TESTS
std::optional<size_t> FileOperationsIssuesPaneState::FindRowIndexForTask(uint64_t taskId) const noexcept
{
    for (size_t rowIndex = 0; rowIndex < _rows.size(); ++rowIndex)
    {
        if (_rows[rowIndex].taskId == taskId)
        {
            return rowIndex;
        }
    }

    return std::nullopt;
}

void FileOperationsIssuesPaneState::FillSelfTestSnapshot(FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) const noexcept
{
    outSnapshot                      = {};
    outSnapshot.rowCount             = _rows.size();
    outSnapshot.refreshGeneration    = _refreshGeneration;
    outSnapshot.visibleWork          = _grid ? _grid->GetVisibleWorkMetrics() : RedSalamander::DxUi::GridVisibleWorkMetrics{};
    outSnapshot.themeDark            = _theme.dark;
    outSnapshot.themeHighContrast    = _theme.highContrast;
    outSnapshot.themeRainbow         = _theme.menu.rainbowMode;
    outSnapshot.dxRenderCount        = _dxHost.DebugGetRenderCount();
    outSnapshot.dxResizeCount        = _dxHost.DebugGetResizeCount();
    outSnapshot.dxResizeFailureCount = _dxHost.DebugGetResizeFailureCount();
    outSnapshot.gridFocused          = _dxHost.GetFocusControl() == _grid;
    if (! _rows.empty())
    {
        outSnapshot.firstVisibleTaskId = _rows.front().taskId;
    }
    outSnapshot.taskHeaderRect      = _grid ? _grid->GetVisibleColumnHeaderRect(1u).value_or(D2D1::RectF()) : D2D1::RectF();
    outSnapshot.operationHeaderRect = _grid ? _grid->GetVisibleColumnHeaderRect(2u).value_or(D2D1::RectF()) : D2D1::RectF();
    outSnapshot.messageHeaderRect   = _grid ? _grid->GetVisibleColumnHeaderRect(7u).value_or(D2D1::RectF()) : D2D1::RectF();
    outSnapshot.hasActiveSort       = _sortSpec.direction != SortDirection::None;
    outSnapshot.sortColumnIndex     = _sortSpec.columnIndex;
    outSnapshot.sortDescending      = _sortSpec.direction == SortDirection::Descending;

    if (! _grid)
    {
        return;
    }

    const auto selection       = _grid->GetSelectionModel().GetOrderedSelection();
    outSnapshot.selectionCount = selection.size();
    if (selection.empty())
    {
        return;
    }

    outSnapshot.primarySelectedRowId     = selection.front();
    const std::optional<size_t> rowIndex = _model.FindRowByStableId(selection.front());
    if (! rowIndex.has_value() || rowIndex.value() >= _rows.size())
    {
        return;
    }

    outSnapshot.primarySelectedTaskId = _rows[rowIndex.value()].taskId;
    RedSalamander::DxUi::GridDebugRowVisualState rowVisualState{};
    if (_grid->DebugGetRowVisualState(_dxHost.GetTheme(), rowIndex.value(), rowVisualState))
    {
        outSnapshot.selectedIssueRowFillArgb    = rowVisualState.fillArgb;
        outSnapshot.selectedIssueRowTextArgb    = rowVisualState.textArgb;
        outSnapshot.selectedIssueRowUsesRainbow = rowVisualState.usesRainbow;
    }
}

bool FileOperationsIssuesPaneState::SelfTestHitTestGridPoint(D2D1_POINT_2F pointDip, FileOperationsIssuesPane::SelfTestGridHit& outHit) const noexcept
{
    outHit = {};
    if (! _grid)
    {
        return false;
    }

    Grid::GridDebugHitInfo hit{};
    if (! _grid->DebugHitTestPoint(RedSalamander::DxUi::MakePointDip(pointDip), hit))
    {
        return false;
    }

    outHit.zone             = hit.zone;
    outHit.rowIndex         = hit.rowIndex;
    outHit.groupIndex       = hit.groupIndex;
    outHit.columnIndex      = hit.columnIndex;
    outHit.rectDip          = hit.rectDip;
    outHit.onScrollbarThumb = hit.onScrollbarThumb;
    outHit.isHeaderResize   = hit.isHeaderResize;
    return true;
}

bool FileOperationsIssuesPaneState::SelfTestSelectTask(uint64_t taskId) noexcept
{
    if (! _grid)
    {
        return false;
    }

    const std::optional<size_t> rowIndex = FindRowIndexForTask(taskId);
    if (! rowIndex.has_value())
    {
        return false;
    }

    _grid->GetSelectionModel().SetSingle(_rows[rowIndex.value()].stableId);
    _grid->NotifyDataChanged();
    _dxHost.Invalidate();
    return true;
}

bool FileOperationsIssuesPaneState::SelfTestFocusGrid() noexcept
{
    if (! _grid)
    {
        return false;
    }

    const HWND hwnd = _dxHost.GetHwnd();
    if (hwnd && IsWindow(hwnd) != FALSE)
    {
        if (IsIconic(hwnd) != FALSE)
        {
            ShowWindow(hwnd, SW_RESTORE);
        }
        static_cast<void>(SetForegroundWindow(hwnd));
        static_cast<void>(SetActiveWindow(hwnd));
        static_cast<void>(SetFocus(hwnd));
    }

    _dxHost.SetFocusControl(_grid);
    return _dxHost.GetFocusControl() == _grid;
}

bool FileOperationsIssuesPaneState::SelfTestRefresh(bool force) noexcept
{
    RefreshRows(force);
    return true;
}

bool FileOperationsIssuesPaneState::SelfTestScrollByWheelDetents(int detents) noexcept
{
    if (! _grid || detents == 0)
    {
        return detents == 0;
    }

    _dxHost.SetFocusControl(_grid);
    const float wheelDelta = detents > 0 ? static_cast<float>(WHEEL_DELTA) : -static_cast<float>(WHEEL_DELTA);
    const int stepCount    = detents > 0 ? detents : -detents;
    for (int remaining = stepCount; remaining > 0; --remaining)
    {
        if (! _grid->OnMouseWheel(_dxHost, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0))
        {
            return false;
        }
    }

    return true;
}
#endif

void FileOperationsIssuesPaneState::Layout() noexcept
{
    if (! _root || ! _grid || ! _dxHost.GetHwnd())
    {
        return;
    }

    RECT client{};
    GetClientRect(_dxHost.GetHwnd(), &client);
    const float widthDip  = _dxHost.PixelsToDip(static_cast<float>(std::max(0L, client.right - client.left)));
    const float heightDip = _dxHost.PixelsToDip(static_cast<float>(std::max(0L, client.bottom - client.top)));
    _root->SetBounds(D2D1::RectF(0.0f, 0.0f, widthDip, heightDip));
    const float padding = _dxHost.PixelsToDip(static_cast<float>(DipsToPixels(6, _dpi)));
    _grid->SetBounds(D2D1::RectF(padding, padding, std::max(padding + 1.0f, widthDip - padding), std::max(padding + 1.0f, heightDip - padding)));
}

LRESULT FileOperationsIssuesPaneState::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);
    if (! _dxHost.Attach(hwnd))
    {
        return -1;
    }

    _rootStorage = std::make_unique<Panel>();
    _root        = _rootStorage.get();
    _grid        = _root->AddChild<Grid>();
    _grid->SetDelegate(this);
    _grid->SetModel(&_model);
    _grid->SetHeaderHeightDip(30.0f);
    _grid->SetRowHeightDip(34.0f);
    _grid->SetLineClamp(2u);
    _dxHost.SetRoot(std::move(_rootStorage));
    ApplySavedViewState();

    ApplyTheme(hwnd);
    Layout();
    RefreshRows(true);
    SetTimer(hwnd, kRefreshTimerId, kRefreshTimerIntervalMs, nullptr);
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnSize(HWND hwnd, UINT width, UINT height) noexcept
{
    static_cast<void>(width);
    static_cast<void>(height);
    Layout();
    if (hwnd && HasLiveOwnerState())
    {
        fileOps->SaveIssuesPanePlacement(hwnd);
    }
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnTimer(HWND hwnd, UINT_PTR timerId) noexcept
{
    if (timerId != kRefreshTimerId)
    {
        return 0;
    }
    if (! hostLifetime.lock())
    {
        DestroyWindow(hwnd);
        return 0;
    }
    RefreshRows(false);
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnMove(HWND hwnd) noexcept
{
    if (HasLiveOwnerState())
    {
        fileOps->SaveIssuesPanePlacement(hwnd);
    }
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnExitSizeMove(HWND hwnd) noexcept
{
    if (HasLiveOwnerState())
    {
        fileOps->SaveIssuesPanePlacement(hwnd);
    }
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnShowWindow(HWND hwnd, BOOL visible) noexcept
{
    if (visible)
    {
        RefreshRows(true);
        ApplyTheme(hwnd);
    }
    else if (HasLiveOwnerState())
    {
        SaveViewState();
    }
    if (HasLiveOwnerState())
    {
        fileOps->SaveIssuesPanePlacement(hwnd);
    }
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnClose(HWND hwnd) noexcept
{
    const HWND focusedBeforeHide = GetFocus();
    if (HasLiveOwnerState())
    {
        SaveViewState();
        fileOps->SaveIssuesPanePlacement(hwnd);
    }
    ShowWindow(hwnd, SW_HIDE);
    if (folderWindow && hostLifetime.lock())
    {
        RestoreActivePaneFolderViewFocusIfWindowHadFocusBeforeHide(*folderWindow, hwnd, focusedBeforeHide);
    }
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnDpiChanged(HWND hwnd, UINT dpi, const RECT* suggested) noexcept
{
    _dpi = dpi == 0 ? USER_DEFAULT_SCREEN_DPI : dpi;
    if (suggested)
    {
        SetWindowPos(hwnd,
                     nullptr,
                     suggested->left,
                     suggested->top,
                     std::max(1L, suggested->right - suggested->left),
                     std::max(1L, suggested->bottom - suggested->top),
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }
    Layout();
    ApplyTheme(hwnd);
    if (HasLiveOwnerState())
    {
        fileOps->SaveIssuesPanePlacement(hwnd);
    }
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnThemeChanged(HWND hwnd) noexcept
{
    if (_inThemeChange)
    {
        return 0;
    }
    _inThemeChange = true;
    ApplyTheme(hwnd);
    _inThemeChange = false;
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnNcDestroy(HWND hwnd) noexcept
{
    KillTimer(hwnd, kRefreshTimerId);
    if (HasLiveOwnerState())
    {
        SaveViewState();
        fileOps->OnIssuesPaneDestroyed(hwnd);
    }
    _dxHost.Detach();
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    _deletePending = true;
    if (_dispatchDepth == 0u)
    {
        delete this;
    }
    return 0;
}

LRESULT FileOperationsIssuesPaneState::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    bool dxHandled         = false;
    const LRESULT dxResult = _dxHost.HandleMessage(hwnd, msg, wp, lp, dxHandled);
    if (dxHandled)
    {
        switch (msg)
        {
            case WM_LBUTTONUP:
            case WM_RBUTTONUP:
            case WM_MBUTTONUP:
            case WM_CAPTURECHANGED:
                if (HasLiveOwnerState())
                {
                    SaveViewState();
                }
                break;
        }
        return dxResult;
    }

    switch (msg)
    {
        case WM_CREATE: return OnCreate(hwnd);
        case WM_SIZE: return OnSize(hwnd, LOWORD(lp), HIWORD(lp));
        case WM_TIMER: return OnTimer(hwnd, wp);
        case WM_MOVE: return OnMove(hwnd);
        case WM_GETMINMAXINFO:
        {
            auto* info = reinterpret_cast<MINMAXINFO*>(lp);
            if (info)
            {
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hwnd, *info, 520, 320);
                static_cast<void>(WindowMaximizeBehavior::ApplyVerticalMaximize(hwnd, *info));
            }
            return 0;
        }
        case WM_EXITSIZEMOVE: return OnExitSizeMove(hwnd);
        case WM_SHOWWINDOW: return OnShowWindow(hwnd, wp != FALSE);
        case WM_DPICHANGED: return OnDpiChanged(hwnd, static_cast<UINT>(HIWORD(wp)), reinterpret_cast<const RECT*>(lp));
        case WM_THEMECHANGED:
        case WM_SETTINGCHANGE:
        case WM_SYSCOLORCHANGE: return OnThemeChanged(hwnd);
        case WM_NCACTIVATE:
            if (folderWindow && hostLifetime.lock() && ! _inTitleBarThemeApply)
            {
                _inTitleBarThemeApply = true;
                ApplyTitleBarTheme(hwnd, folderWindow->GetTheme(), wp != FALSE);
                _inTitleBarThemeApply = false;
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_CLOSE: return OnClose(hwnd);
        case WM_NCDESTROY: return OnNcDestroy(hwnd);
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK FileOperationsIssuesPaneState::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCCREATE)
    {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        state    = reinterpret_cast<FileOperationsIssuesPaneState*>(cs ? cs->lpCreateParams : nullptr);
        if (! state)
        {
            return FALSE;
        }
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
    }

    if (! state)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    ++state->_dispatchDepth;
    const auto finishDispatch = wil::scope_exit([state]() noexcept
    {
        if (state->_dispatchDepth > 0u)
        {
            --state->_dispatchDepth;
        }
        if (state->_dispatchDepth == 0u && state->_deletePending)
        {
            delete state;
        }
    });

    return state->WndProc(hwnd, msg, wp, lp);
}

HWND FileOperationsIssuesPane::Create(FolderWindow::FileOperationState* fileOps,
                                      FolderWindow* folderWindow,
                                      HWND ownerWindow,
                                      std::weak_ptr<void> hostLifetime) noexcept
{
    if (! fileOps || ! folderWindow || hostLifetime.expired() || ! RegisterFileOperationsIssuesPaneClass(GetModuleHandleW(nullptr)))
    {
        return nullptr;
    }

    auto statePtr          = std::make_unique<FileOperationsIssuesPaneState>();
    statePtr->fileOps      = fileOps;
    statePtr->folderWindow = folderWindow;
    statePtr->hostLifetime = std::move(hostLifetime);

    const UINT ownerDpi = ownerWindow ? GetDpiForWindow(ownerWindow) : USER_DEFAULT_SCREEN_DPI;
    RECT windowRect{};
    bool startMaximized = false;
    if (! fileOps->TryGetIssuesPanePlacement(windowRect, startMaximized, ownerDpi))
    {
        const DWORD style   = WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN;
        const DWORD exStyle = WS_EX_APPWINDOW;
        RECT desiredWindowRect{0, 0, DipsToPixels(1100, ownerDpi), DipsToPixels(560, ownerDpi)};
        AdjustWindowRectExForDpi(&desiredWindowRect, style, FALSE, exStyle, ownerDpi);

        const int width  = std::max(1L, desiredWindowRect.right - desiredWindowRect.left);
        const int height = std::max(1L, desiredWindowRect.bottom - desiredWindowRect.top);
        HMONITOR monitor = MonitorFromWindow(ownerWindow ? ownerWindow : folderWindow->GetHwnd(), MONITOR_DEFAULTTONEAREST);
        MONITORINFO mi{};
        mi.cbSize = sizeof(mi);
        if (! GetMonitorInfoW(monitor, &mi))
        {
            return nullptr;
        }

        const int maxX    = static_cast<int>(mi.rcWork.right - width);
        const int maxY    = static_cast<int>(mi.rcWork.bottom - height);
        const int centerX = static_cast<int>(mi.rcWork.left + (mi.rcWork.right - mi.rcWork.left - width) / 2);
        const int centerY = static_cast<int>(mi.rcWork.top + (mi.rcWork.bottom - mi.rcWork.top - height) / 2);
        windowRect.left   = maxX >= mi.rcWork.left ? std::clamp(centerX, static_cast<int>(mi.rcWork.left), maxX) : mi.rcWork.left;
        windowRect.top    = maxY >= mi.rcWork.top ? std::clamp(centerY, static_cast<int>(mi.rcWork.top), maxY) : mi.rcWork.top;
        windowRect.right  = windowRect.left + width;
        windowRect.bottom = windowRect.top + height;
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_PANE_TITLE);
    auto* state              = statePtr.release();
    HWND pane                = CreateWindowExW(WS_EX_APPWINDOW,
                                               kFileOperationsIssuesPaneClassName,
                                               title.c_str(),
                                               WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
                                               windowRect.left,
                                               windowRect.top,
                                               std::max(1L, windowRect.right - windowRect.left),
                                               std::max(1L, windowRect.bottom - windowRect.top),
                                               nullptr,
                                               nullptr,
                                               GetModuleHandleW(nullptr),
                                               state);
    if (! pane)
    {
        std::unique_ptr<FileOperationsIssuesPaneState> reclaim(state);
        return nullptr;
    }

    ShowWindow(pane, startMaximized ? SW_SHOWMAXIMIZED : SW_SHOWNORMAL);
    UpdateWindow(pane);
    return pane;
}

#ifdef ENABLE_TESTS
bool FileOperationsIssuesPane::TryGetSelfTestSnapshot(HWND hwnd, SelfTestSnapshot& outSnapshot) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        outSnapshot = {};
        return false;
    }

    const auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! state)
    {
        outSnapshot = {};
        return false;
    }

    state->FillSelfTestSnapshot(outSnapshot);
    return true;
}

bool FileOperationsIssuesPane::SelfTestSelectTask(HWND hwnd, uint64_t taskId) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return state && state->SelfTestSelectTask(taskId);
}

bool FileOperationsIssuesPane::SelfTestHitTestGridPoint(HWND hwnd, float xDip, float yDip, SelfTestGridHit& outHit) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        outHit = {};
        return false;
    }

    const auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! state)
    {
        outHit = {};
        return false;
    }

    return state->SelfTestHitTestGridPoint(D2D1::Point2F(xDip, yDip), outHit);
}

bool FileOperationsIssuesPane::SelfTestFocusGrid(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return state && state->SelfTestFocusGrid();
}

bool FileOperationsIssuesPane::SelfTestRefresh(HWND hwnd, bool force) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return state && state->SelfTestRefresh(force);
}

bool FileOperationsIssuesPane::SelfTestScrollByWheelDetents(HWND hwnd, int detents) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return state && state->SelfTestScrollByWheelDetents(detents);
}
#endif
