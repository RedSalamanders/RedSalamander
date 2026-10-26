#include "FolderWindow.FileOperations.IssuesPane.h"
#include "FolderWindow.FileOperations.IssuesProjection.h"
#include "WindowActivation.h"

#include "DxUiThemePalette.h"
#include "FluentIcons.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "Helpers.h"
#include "UnicodeClipboard.h"
#include "WindowMaximizeBehavior.h"
#include "WindowSizing.h"
#include <DxUi/DxUi.h>
#include <DxUi/Typography.h>
#include <array>
#include <atomic>
#include <cmath>
#include <functional>

#include <algorithm>
#include <format>
#include <memory>
#include <unordered_map>
#include <vector>

namespace
{
using DxUi::Grid;
using DxUi::GridCellData;
using DxUi::GridColumnDesc;
using DxUi::GridRowStyle;
using DxUi::GridRowTone;
using DxUi::GridSortSpec;
using DxUi::IGridDelegate;
using DxUi::IGridModel;
using DxUi::Panel;
using DxUi::SortDirection;
using DxUi::ThemePalette;
using DxUi::WindowHost;

constexpr wchar_t kFileOperationsIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
constexpr UINT_PTR kRefreshTimerId                     = 1;
constexpr UINT kRefreshTimerIntervalMs                 = 750;
constexpr float kIssuesHeaderHeightDip                 = 30.0f;
constexpr float kIssuesPrimaryRowHeightDip             = 88.0f;
constexpr float kIssuesTechnicalRowHeightDip           = 34.0f;
// Application reading order: subject first; technical facts last. Use it for the
// retained child tree, visual layout and complete copy so accessibility agrees.
constexpr std::array<size_t, 9> kIssuesDetailOrder{{8u, 0u, 1u, 2u, 3u, 4u, 5u, 6u, 7u}};
constexpr float kIssuesDetailIconGutterDip = 24.0f;
[[nodiscard]] constexpr DxUi::FontRole IssuesDetailFontRole(size_t index) noexcept
{
    return index == 8u ? DxUi::FontRole::Subtitle : (index == 0u || index == 2u) ? DxUi::FontRole::BodyStrong : DxUi::FontRole::Body;
}

#ifdef ENABLE_TESTS
std::atomic<uint64_t> g_liveIssuesPaneStates{0u};
std::atomic<bool> g_failNextIssuesPaneCreation{false};
#endif

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

[[nodiscard]] UINT OperationStringId(FileSystemOperation operation) noexcept
{
    switch (operation)
    {
        case FILESYSTEM_COPY: return IDS_CMD_COPY;
        case FILESYSTEM_MOVE: return IDS_CMD_MOVE;
        case FILESYSTEM_DELETE: return IDS_CMD_DELETE;
        case FILESYSTEM_RENAME: return IDS_CMD_RENAME;
        case FILESYSTEM_CREATE_DIRECTORY: return IDS_FILEOPS_ISSUES_OPERATION_UNKNOWN;
        default: return IDS_FILEOPS_ISSUES_OPERATION_UNKNOWN;
    }
}

// Application detail navigation: the generic ScrollPanel supplies retained scrolling;
// this pane defines keyboard paging and copying the complete selected diagnostic.
class IssuesDetailBody final : public DxUi::ScrollPanel
{
public:
    IssuesDetailBody()
    {
        SetFocusable(true);
    }
    IssuesDetailBody(const IssuesDetailBody&)            = delete;
    IssuesDetailBody(IssuesDetailBody&&)                 = delete;
    IssuesDetailBody& operator=(const IssuesDetailBody&) = delete;
    IssuesDetailBody& operator=(IssuesDetailBody&&)      = delete;
    struct IconDecoration final
    {
        DxUi::Label* label = nullptr; // Borrowed from this panel's retained children.
        wchar_t fluent     = 0;
        wchar_t fallback   = 0;
        D2D1_COLOR_F color = D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f);
    };
    void SetIcon(size_t slot, DxUi::Label* label, wchar_t fluent, wchar_t fallback, D2D1_COLOR_F color) noexcept
    {
        _icons[slot] = {label, fluent, fallback, color};
    }
    [[nodiscard]] wchar_t IconGlyph(size_t slot, const WindowHost& host) const noexcept
    {
        const auto& icon = _icons[slot];
        return icon.fluent != 0 && host.HasFluentIconFont() ? icon.fluent : icon.fallback;
    }
    [[nodiscard]] D2D1_RECT_F IconBounds(size_t slot) const noexcept
    {
        const auto* label = _icons[slot].label;
        if (! label)
            return D2D1::RectF();
        const auto bounds = label->GetBounds();
        return D2D1::RectF(bounds.left - kIssuesDetailIconGutterDip, bounds.top, bounds.left - 4.0f, bounds.top + std::min(20.0f, bounds.bottom - bounds.top));
    }
    [[nodiscard]] D2D1_COLOR_F IconColor(size_t slot, const WindowHost& host) const noexcept
    {
        return host.GetTheme().highContrast ? host.GetTheme().text : _icons[slot].color;
    }
    void Paint(WindowHost& host) const override
    {
        if (! IsVisible())
            return;
        DxUi::ScrollPanel::Paint(host); // Native labels paint their own complete text.
        auto* dc = host.GetDeviceContext();
        if (! dc)
            return;
        dc->PushAxisAlignedClip(GetViewportRect(), D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
        D2D1_MATRIX_3X2_F oldTransform{};
        dc->GetTransform(&oldTransform);
        const auto restorePaintState = wil::scope_exit([&]() noexcept
        {
            dc->SetTransform(oldTransform);
            dc->PopAxisAlignedClip();
        });
        auto scrolled                = oldTransform;
        scrolled._32 -= GetScrollOffset();
        dc->SetTransform(scrolled);
        for (size_t slot = 0u; slot < _icons.size(); ++slot)
        {
            const auto& icon = _icons[slot];
            if (! icon.label || ! icon.label->IsVisible())
                continue;
            const wchar_t glyph = IconGlyph(slot, host);
            if (glyph == 0)
                continue;
            const bool fluent = icon.fluent != 0 && host.HasFluentIconFont();
            auto* format      = host.GetTextFormat(
                fluent ? DxUi::FontRole::Icon : DxUi::FontRole::Body, DWRITE_TEXT_ALIGNMENT_CENTER, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false);
            if (! format)
                continue;
            const auto bounds = IconBounds(slot);
            if (auto* brush = host.GetSolidBrush(IconColor(slot, host)))
                dc->DrawTextW(&glyph, 1u, format, &bounds, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
        }
    }
    void PaintOverlay(DxUi::WindowHost& host) const override
    {
        DxUi::ScrollPanel::PaintOverlay(host);
        if (! HasFocus())
            return;
        DxUi::ButtonChromeDrawSpec spec{};
        spec.bounds = GetBounds();
        DxUi::ButtonChromeCustomStyle style{};
        style.showFill   = false;
        style.showFocus  = true;
        style.focus      = host.GetTheme().focusStrokeOuter;
        style.focusRing  = DxUi::ButtonChromeFocusRing::Standard;
        spec.customStyle = style;
        DxUi::DrawButtonChrome(host.GetDeviceContext(), host.GetSolidBrush(style.focus), nullptr, nullptr, host.GetTheme(), spec);
    }

    std::function<bool()> copy;
    bool OnCopy(DxUi::WindowHost&) override
    {
        return copy && copy();
    }
    bool OnKeyDown(DxUi::WindowHost& host, UINT key, UINT modifiers) override
    {
        if (DxUi::ModifiersContainCtrl(modifiers) && ! DxUi::ModifiersContainAlt(modifiers) && key == 'C')
            return OnCopy(host);
        if (DxUi::ModifiersContainAlt(modifiers) || DxUi::ModifiersContainCtrl(modifiers))
            return DxUi::ScrollPanel::OnKeyDown(host, key, modifiers);
        const auto viewport = GetViewportRect();
        const float page    = std::max(24.0f, viewport.bottom - viewport.top - 24.0f);
        float offset        = GetScrollOffset();
        switch (key)
        {
            case VK_UP: offset -= 24.0f; break;
            case VK_DOWN: offset += 24.0f; break;
            case VK_PRIOR: offset -= page; break;
            case VK_NEXT: offset += page; break;
            case VK_HOME: offset = 0.0f; break;
            case VK_END: offset = GetContentHeight(); break;
            default: return DxUi::ScrollPanel::OnKeyDown(host, key, modifiers);
        }
        SetScrollOffset(offset);
        host.Invalidate();
        return true;
    }

private:
    std::array<IconDecoration, 3> _icons{};
};

using IssuesRow = FileOperationsIssues::Row;
[[nodiscard]] std::wstring_view IssuesColumnText(const IssuesRow& row, size_t key) noexcept
{
    return FileOperationsIssues::Projection::ColumnText(row, key);
}

class IssuesGridModel final : public IGridModel
{
public:
    explicit IssuesGridModel(const AppTheme& theme) : _theme(&theme)
    {
        _columns = {
            {L"time", {}, 170.0f, 120.0f, DxUi::GridColumnKind::Text, true, false},
            {L"task", {}, 70.0f, 56.0f, DxUi::GridColumnKind::Text, true, false},
            {L"operation", {}, 90.0f, 80.0f, DxUi::GridColumnKind::Text, true, false},
            {L"severity", {}, 86.0f, 72.0f, DxUi::GridColumnKind::Text, true, false},
            {L"hresult", {}, 104.0f, 100.0f, DxUi::GridColumnKind::Text, true, false},
            {L"status", {}, 220.0f, 120.0f, DxUi::GridColumnKind::Text, true, true},
            {L"category", {}, 140.0f, 100.0f, DxUi::GridColumnKind::Text, true, true},
            {L"message", {}, 300.0f, 140.0f, DxUi::GridColumnKind::Text, true, true},
            {L"source", {}, 300.0f, 160.0f, DxUi::GridColumnKind::Text, true, true},
            {L"destination", {}, 300.0f, 160.0f, DxUi::GridColumnKind::Text, true, true},
            {L"subject", {}, 240.0f, 100.0f, DxUi::GridColumnKind::Text, true, true},
            {L"decision", {}, 140.0f, 112.0f, DxUi::GridColumnKind::Text, true, true},
            {L"result", {}, 140.0f, 100.0f, DxUi::GridColumnKind::Text, true, true},
            {L"problem", {}, 240.0f, 120.0f, DxUi::GridColumnKind::Text, true, true},
            {L"summary-operation", {}, 90.0f, 80.0f, DxUi::GridColumnKind::Text, true, true},
        };
        ReloadColumnTitles();
    }

    bool ReloadColumnTitles()
    {
        constexpr std::array<UINT, 15> resources{{IDS_FILEOPS_ISSUES_COL_TIME,
                                                  IDS_FILEOPS_ISSUES_COL_TASK,
                                                  IDS_FILEOPS_ISSUES_COL_OPERATION,
                                                  IDS_FILEOPS_ISSUES_COL_SEVERITY,
                                                  IDS_FILEOPS_ISSUES_COL_HRESULT,
                                                  IDS_FILEOPS_ISSUES_COL_STATUS_TEXT,
                                                  IDS_FILEOPS_ISSUES_COL_CATEGORY,
                                                  IDS_FILEOPS_ISSUES_COL_MESSAGE,
                                                  IDS_FILEOPS_ISSUES_COL_SOURCE,
                                                  IDS_FILEOPS_ISSUES_COL_DESTINATION,
                                                  IDS_FILEOPS_ISSUES_FILE,
                                                  IDS_FILEOPS_ISSUES_DECISION,
                                                  IDS_FILEOPS_ISSUES_RESULT,
                                                  IDS_FILEOPS_ISSUES_PROBLEM,
                                                  IDS_FILEOPS_ISSUES_COL_OPERATION}};
        bool changed = false;
        for (size_t i = 0u; i < resources.size(); ++i)
        {
            auto title = LoadStringResource(nullptr, resources[i]);
            if (_columns[i].title != title)
            {
                _columns[i].title = std::move(title);
                changed           = true;
            }
        }
        return changed;
    }

    void SetPresentation(bool technical, bool narrow)
    {
        _visible = technical ? std::vector<size_t>{0, 1, 2, 3, 4, 5, 6, 7, 8, 9}
                   : narrow  ? std::vector<size_t>{10, 13, 11, 12}
                             : std::vector<size_t>{10, 13, 14, 11, 12};
    }
    [[nodiscard]] size_t ColumnKey(size_t index) const noexcept
    {
        return index < _visible.size() ? _visible[index] : _columns.size();
    }
    [[nodiscard]] std::optional<size_t> VisibleIndex(size_t key) const noexcept
    {
        const auto found = std::find(_visible.begin(), _visible.end(), key);
        return found == _visible.end() ? std::nullopt : std::optional<size_t>{static_cast<size_t>(found - _visible.begin())};
    }
    [[nodiscard]] size_t KeyCount() const noexcept
    {
        return _columns.size();
    }
    [[nodiscard]] const GridColumnDesc& ColumnByKey(size_t key) const
    {
        return _columns.at(key);
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
        return _visible.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return _columns.at(ColumnKey(columnIndex));
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        if (! _rows || rowIndex >= _rows->size())
        {
            outCell = {};
            return;
        }

        const IssuesRow& row = _rows->at(rowIndex);
        outCell              = {};
        const size_t key     = ColumnKey(columnIndex);
        outCell.text         = IssuesColumnText(row, key);
        outCell.multiline    = key >= 7u;
        outCell.tooltipText  = outCell.text;
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

        // A row without its original problem records only a decision or result, and an Information note
        // (I35) is an expected outcome: both are neutral, not a warning.
        style.tone = ! row.hasOriginalProblem || row.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Info ? GridRowTone::None
                     : (row.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Error)                        ? GridRowTone::Error
                                                                                                                            : GridRowTone::Warning;
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
    std::vector<size_t> _visible{10, 13, 14, 11, 12};
};

class FileOperationsIssuesPaneState final : public IGridDelegate
{
public:
    FileOperationsIssuesPaneState() : _rows(_projection.Rows())
    {
#ifdef ENABLE_TESTS
        g_liveIssuesPaneStates.fetch_add(1u, std::memory_order_relaxed);
#endif
    }
    ~FileOperationsIssuesPaneState() override
    {
#ifdef ENABLE_TESTS
        g_liveIssuesPaneStates.fetch_sub(1u, std::memory_order_relaxed);
#endif
    }
    void CompleteCreation() noexcept
    {
        _creationInProgress = false;
    }
    bool IsPendingDelete() const noexcept
    {
        return _deletePending;
    }
    FileOperationsIssuesPaneState(const FileOperationsIssuesPaneState&)            = delete;
    FileOperationsIssuesPaneState(FileOperationsIssuesPaneState&&)                 = delete;
    FileOperationsIssuesPaneState& operator=(const FileOperationsIssuesPaneState&) = delete;
    FileOperationsIssuesPaneState& operator=(FileOperationsIssuesPaneState&&)      = delete;

    FolderWindow::FileOperationState* fileOps = nullptr;
    FolderWindow* folderWindow                = nullptr;
    std::weak_ptr<void> hostLifetime;

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnGridSelectionChanged(Grid&) override
    {
        if (! _notifyingRows)
        {
            CaptureSelection();
            UpdateDetail(false);
        }
    }
    void OnGridSelectionChanged() override
    {
        if (! _notifyingRows)
        {
            CaptureSelection();
            UpdateDetail(false);
        }
    }

    void OnGridSortRequested(const GridSortSpec& sortSpec) override
    {
        _sortSpec             = sortSpec;
        _sortSpec.columnIndex = _model.ColumnKey(sortSpec.columnIndex);
        static_cast<void>(_projection.SetSort({_sortSpec.columnIndex,
                                               _sortSpec.direction == DxUi::SortDirection::Ascending    ? FileOperationsIssues::SortDirection::Ascending
                                               : _sortSpec.direction == DxUi::SortDirection::Descending ? FileOperationsIssues::SortDirection::Descending
                                                                                                        : FileOperationsIssues::SortDirection::None}));
        if (_grid)
        {
            ApplyGridSort();
            _grid->NotifyDataChanged();
        }
        SaveViewState();
        Layout();
        _dxHost.Invalidate();
    }

    [[nodiscard]] bool SelectTask(uint64_t taskId, uint64_t promptGeneration) noexcept;

#ifdef ENABLE_TESTS
    void FillSelfTestSnapshot(FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot, bool includeDetail) const noexcept;
    bool SelfTestDetailKey(UINT key) noexcept
    {
        return _detailBody && _detailBody->OnKeyDown(_dxHost, key, 0u);
    }
    bool SelfTestSetTechnicalExpanded(bool expanded) noexcept
    {
        if (! _technicalButton)
            return false;
        if (_technicalExpanded != expanded)
        {
            _technicalExpanded = expanded;
            UpdateDetail(false);
        }
        return true;
    }

    [[nodiscard]] bool SelfTestSelectDiagnostic(uint64_t taskId, std::wstring_view category, HRESULT status) noexcept;

    [[nodiscard]] bool SelfTestHitTestGridPoint(D2D1_POINT_2F pointDip, FileOperationsIssuesPane::SelfTestGridHit& outHit) const noexcept;
    [[nodiscard]] bool SelfTestFocusGrid() noexcept;
    [[nodiscard]] bool SelfTestRequestSelectRow(size_t rowIndex, UINT modifiers) noexcept
    {
        return _grid && _grid->RequestSelectRow(rowIndex, modifiers);
    }
    [[nodiscard]] bool SelfTestRefresh(bool force) noexcept;
    [[nodiscard]] bool SelfTestScrollByWheelDetents(int detents) noexcept;
#endif

private:
    void CaptureSelection() noexcept
    {
        if (! _grid)
        {
            _projection.ClearSelection();
            return;
        }
        const auto& selection = _grid->GetSelectionModel().GetOrderedSelection();
        _projection.SetSelectedStableId(selection.empty() ? 0u : selection.front());
    }

    [[nodiscard]] bool HasLiveOwnerState() const noexcept
    {
        return fileOps && ! hostLifetime.expired();
    }

    LRESULT OnCreate(HWND hwnd) noexcept;
    LRESULT OnSize(HWND hwnd, UINT sizeType, UINT width, UINT height) noexcept;
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
    void UpdateRefreshTimer(HWND hwnd) noexcept;
    void Layout() noexcept;
    void UpdateDetail(bool resetScroll) noexcept;
    bool CopyDetail() noexcept;
    void ApplySavedViewState() noexcept;
    void SaveViewState() noexcept;
    void CaptureColumnPreferences() noexcept;
    void ApplyGridPresentation(float contentWidth) noexcept;
    void ApplyGridSort() noexcept;
    [[nodiscard]] std::optional<size_t> FindRowIndexForTask(uint64_t taskId, uint64_t promptGeneration) const noexcept;
    [[nodiscard]] bool SelectRowIndex(size_t rowIndex) noexcept;

    UINT _dpi                  = USER_DEFAULT_SCREEN_DPI;
    bool _inThemeChange        = false;
    bool _inTitleBarThemeApply = false;
    size_t _dispatchDepth      = 0u;
    bool _deletePending        = false;
    bool _creationInProgress   = true;
    bool _tearingDown          = false;
    bool _notifyingRows        = false;
    AppTheme _theme{};
    WindowHost _dxHost;
    std::unique_ptr<Panel> _rootStorage;
    Panel* _root                  = nullptr;
    Grid* _grid                   = nullptr;
    IssuesDetailBody* _detailBody = nullptr;
    std::array<DxUi::Label*, 9> _detailLabels{};
    DxUi::Button* _technicalButton = nullptr;
    DxUi::Button* _copyButton      = nullptr;
    DxUi::Label* _copyFeedback     = nullptr;
    DxUi::Label* _sortStatus       = nullptr;
    std::vector<DxUi::GridColumnLayoutEntry> _columnPreferences;
    std::vector<DxUi::GridColumnLayoutEntry> _appliedColumnLayout;
    bool _gridTechnical = false;
    bool _gridNarrow    = false;
    float _gridWidth    = -1.0f;
    std::wstring _detailCopyText;
    UINT _copyFeedbackResourceId = 0u;
    bool _technicalExpanded      = false;
    uint64_t _detailRowId        = 0u;
    IssuesGridModel _model{_theme};
    GridSortSpec _sortSpec{};
    FileOperationsIssues::Projection _projection;
    const std::vector<IssuesRow>& _rows;
    uint64_t _refreshGeneration      = 0;
    uint64_t _languageRevision       = 0;
    uint64_t _capturedIssuesRevision = 0u;
    bool _refreshTimerRunning        = false;
    // x-perf#6: labels, operation and severity names and system messages have tiny domains. Each
    // is formatted once per language, not once per retained row on every changed revision.
    bool _textCacheValid                = false;
    uint64_t _textCacheLanguageRevision = 0u;
    FileOperationsIssues::Text _textCache{};
    std::wstring _errorSeverityText;
    std::wstring _warningSeverityText;
    std::unordered_map<unsigned, std::wstring> _operationTextCache;
    std::unordered_map<HRESULT, std::wstring> _systemTextCache;
#ifdef ENABLE_TESTS
    uint64_t _systemMessageLookups   = 0u;
    uint64_t _presentationApplyCount = 0u;
    uint64_t _modelRebuildCount      = 0u;
#endif
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

void FileOperationsIssuesPaneState::RefreshRows(bool force) noexcept
{
    if (! fileOps || ! hostLifetime.lock())
        return;

    // I35 R5.3: a change of the informational-notes setting changes which rows are visible.
    const bool showInformationalNotes = fileOps->GetShowInformationalNotes();
    if (showInformationalNotes != _textCache.showInformationalNotes)
    {
        _textCache.showInformationalNotes = showInformationalNotes;
        force                             = true;
    }

    std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> diagnostics;
    std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> completedIssues;
    uint64_t capturedRevision   = _capturedIssuesRevision;
    const uint64_t lastRevision = force ? 0u : _capturedIssuesRevision;
    if (! fileOps->CollectIssuesSnapshotIfChanged(lastRevision, diagnostics, completedIssues, capturedRevision))
        return;

    // Preserve the actual grid selection before the projection reconciles it against the new revision.
    CaptureSelection();
    _capturedIssuesRevision         = capturedRevision;
    const uint64_t languageRevision = Localization::GetLanguagePreferenceRevision();
    if (! _textCacheValid || _textCacheLanguageRevision != languageRevision)
    {
        _textCache.infoSeverity      = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_SEVERITY_INFO);
        _textCache.unknownDecision   = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_DECISION_UNKNOWN);
        _textCache.automaticDecision = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_AUTOMATIC);
        _textCache.unknownResult     = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_RESULT_UNKNOWN);
        _textCache.completedResult   = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_COMPLETED);
        _textCache.skippedResult     = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_SKIPPED);
        _textCache.canceledResult    = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_CANCELED);
        _textCache.failedResult      = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_FAILED_SHORT);
        _errorSeverityText           = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        _warningSeverityText         = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        _operationTextCache.clear();
        _systemTextCache.clear();
        _textCacheLanguageRevision = languageRevision;
        _textCacheValid            = true;
    }
    const auto operationText = [&](FileSystemOperation operation) -> const std::wstring&
    {
        auto found = _operationTextCache.find(static_cast<unsigned>(operation));
        if (found == _operationTextCache.end())
            found = _operationTextCache.emplace(static_cast<unsigned>(operation), LoadStringResource(nullptr, OperationStringId(operation))).first;
        return found->second;
    };
    const auto systemText = [&](HRESULT status) -> const std::wstring&
    {
        auto found = _systemTextCache.find(status);
        if (found == _systemTextCache.end())
        {
            // Retained rows carry few distinct statuses; this bound only stops pathological growth.
            if (_systemTextCache.size() >= 256u)
                _systemTextCache.clear();
#ifdef ENABLE_TESTS
            ++_systemMessageLookups;
#endif
            found = _systemTextCache.emplace(status, TryFormatSystemMessage(GetSystemMessageIdFromHResult(status))).first;
        }
        return found->second;
    };

    std::vector<FileOperationsIssues::Observation> observations;
    observations.reserve(diagnostics.size() + completedIssues.size());
    const auto appendIssue = [&](const FolderWindow::FileOperationState::TaskDiagnosticEntry& issue)
    {
        const auto facts = FolderWindow::FileOperationState::QualifyIssueFacts(issue);
        if (! facts.ProjectsRow())
            return;

        FileOperationsIssues::Observation observation{};
        observation.issue         = &issue;
        observation.timeText      = FormatTimeText(issue.localTime);
        observation.operationText = operationText(issue.operation);
        observation.severityText  = issue.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Error     ? _errorSeverityText
                                    : issue.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Warning ? _warningSeverityText
                                                                                                                      : _textCache.infoSeverity;
        observation.statusText    = std::format(L"0x{:08X}", static_cast<unsigned long>(issue.status));
        // One system-message lookup per status and language: a status without system text reads
        // "HRESULT 0x..." only as a technical detail.
        const std::wstring& statusSystemText = systemText(issue.status);
        observation.statusTextDetail = statusSystemText.empty() ? std::format(L"HRESULT 0x{:08X}", static_cast<unsigned long>(issue.status)) : statusSystemText;
        // Q9: the Problem column shows typed localized text; the engine message stays technical.
        if (facts.originalProblem || facts.acceptedDecision)
            observation.problemText = FolderWindow::FileOperationState::FormatIssueProblemText(issue, std::wstring_view(statusSystemText));
        if (facts.acceptedDecision)
        {
            const auto& accepted              = *facts.acceptedDecision;
            observation.formattedDecisionText = FileOperations::FormatConflictActionLabel(accepted.action, accepted.scope.bucket, issue.operation);
            // Q10: one cached apply-to-all row per originating prompt, with its item count.
            if (accepted.origin == FileOperations::ConflictDecisionOrigin::CachedScope && accepted.appliedCount != 0u)
                observation.automaticDecisionText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ISSUES_AUTOMATIC_COUNT, accepted.appliedCount);
        }
        observations.push_back(std::move(observation));
    };
    for (const auto& issue : diagnostics)
        appendIssue(issue);
    for (const auto& issue : completedIssues)
        appendIssue(issue);

    if (! _projection.Update(observations, _textCache, _capturedIssuesRevision, force))
        return;

    _model.SetRows(&_rows);
    ++_refreshGeneration;
    if (_grid)
    {
        _notifyingRows = true;
        _grid->NotifyDataChanged();
        _notifyingRows          = false;
        const uint64_t selected = _projection.State().selectedStableId;
        if (selected == 0u)
            _grid->GetSelectionModel().Clear();
    }
    UpdateDetail(false);
    _dxHost.Invalidate();
}
void FileOperationsIssuesPaneState::ApplySavedViewState() noexcept
{
    // Complete preferences retain hidden columns. Display indexes belong to their
    // primary/technical group; IDs, not transient model indexes, are persistent.
    constexpr std::array<size_t, 15> order{{0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 13, 14, 11, 12}};
    for (size_t i = 0u; i < order.size(); ++i)
    {
        const auto& column = _model.ColumnByKey(order[i]);
        _columnPreferences.push_back({column.id, i < 10u ? i : i - 10u, column.widthDip});
    }
    std::wstring savedSort;
    bool descending = false;
    std::vector<Common::Settings::GridColumnLayoutEntry> saved;
    if (fileOps && fileOps->TryGetIssuesPaneViewState(savedSort, descending, saved))
    {
        for (auto& target : _columnPreferences)
        {
            const auto entry = std::find_if(saved.begin(), saved.end(), [&](const auto& value) { return value.columnId == target.columnId; });
            if (entry != saved.end())
            {
                target.displayIndex = entry->displayIndex;
                if (std::isfinite(entry->widthDip) && entry->widthDip > 0.0f)
                    target.widthDip = entry->widthDip;
            }
        }
        for (size_t key = 0u; key < _model.KeyCount(); ++key)
            if (_model.ColumnByKey(key).id == savedSort)
                _sortSpec = {key, descending ? SortDirection::Descending : SortDirection::Ascending};
    }
    // Normalize each complete group once, preserving saved order and missing defaults.
    for (const auto range : {std::pair<size_t, size_t>{0u, 10u}, {10u, 15u}})
    {
        auto first = _columnPreferences.begin() + static_cast<std::ptrdiff_t>(range.first);
        auto last  = _columnPreferences.begin() + static_cast<std::ptrdiff_t>(range.second);
        std::stable_sort(first, last, [](const auto& a, const auto& b) { return a.displayIndex < b.displayIndex; });
        size_t index = 0u;
        for (auto it = first; it != last; ++it)
            it->displayIndex = index++;
    }
    static_cast<void>(_projection.SetSort({_sortSpec.columnIndex,
                                           _sortSpec.direction == DxUi::SortDirection::Ascending    ? FileOperationsIssues::SortDirection::Ascending
                                           : _sortSpec.direction == DxUi::SortDirection::Descending ? FileOperationsIssues::SortDirection::Descending
                                                                                                    : FileOperationsIssues::SortDirection::None}));
}

void FileOperationsIssuesPaneState::CaptureColumnPreferences() noexcept
{
    if (! _grid || _appliedColumnLayout.empty())
        return;
    const auto current = _grid->CaptureColumnLayout();
    std::vector<std::wstring> previousOrder;
    std::vector<std::wstring> currentOrder;
    for (const auto& entry : _appliedColumnLayout)
        previousOrder.push_back(entry.columnId);
    for (const auto& entry : current)
    {
        currentOrder.push_back(entry.columnId);
        const auto previous =
            std::find_if(_appliedColumnLayout.begin(), _appliedColumnLayout.end(), [&](const auto& value) { return value.columnId == entry.columnId; });
        if (previous == _appliedColumnLayout.end() || std::abs(previous->widthDip - entry.widthDip) <= 0.5f)
            continue;
        for (auto& preference : _columnPreferences)
            if (preference.columnId == entry.columnId)
                preference.widthDip = entry.widthDip;
    }
    if (previousOrder != currentOrder)
    {
        // Replace only visible slots in the complete group. A hidden operation
        // retains its position while the remaining primary columns are reordered.
        std::vector<size_t> slots;
        for (const auto& preference : _columnPreferences)
            if (std::find(currentOrder.begin(), currentOrder.end(), preference.columnId) != currentOrder.end())
                slots.push_back(preference.displayIndex);
        std::sort(slots.begin(), slots.end());
        for (size_t i = 0u; i < currentOrder.size() && i < slots.size(); ++i)
            for (auto& preference : _columnPreferences)
                if (preference.columnId == currentOrder[i])
                    preference.displayIndex = slots[i];
    }
    _appliedColumnLayout = current;
}

void FileOperationsIssuesPaneState::ApplyGridSort() noexcept
{
    const auto visible = _model.VisibleIndex(_sortSpec.columnIndex);
    _grid->SetSortSpec(visible.has_value() ? GridSortSpec{visible.value(), _sortSpec.direction} : GridSortSpec{});
    if (_sortStatus)
    {
        const bool hidden = _sortSpec.direction != SortDirection::None && ! visible.has_value() && _sortSpec.columnIndex < _model.KeyCount();
        _sortStatus->SetText(hidden ? FormatStringResource(nullptr,
                                                           IDS_FILEOPS_ISSUES_SORTED_BY,
                                                           _model.ColumnByKey(_sortSpec.columnIndex).title,
                                                           _sortSpec.direction == SortDirection::Descending ? L"↓" : L"↑")
                                    : L"");
        _sortStatus->SetVisible(hidden);
    }
}

void FileOperationsIssuesPaneState::ApplyGridPresentation(float contentWidth) noexcept
{
    const bool narrow = contentWidth < 680.0f;
    // The technical presentation has no narrow form: a width change there rebuilds nothing (issues#5).
    const bool changed = _gridTechnical != _technicalExpanded || (! _technicalExpanded && _gridNarrow != narrow);
    if (! changed && std::abs(_gridWidth - contentWidth) < 0.5f)
        return;
#ifdef ENABLE_TESTS
    ++_presentationApplyCount;
#endif
    CaptureColumnPreferences();
    _gridTechnical = _technicalExpanded;
    _gridNarrow    = narrow;
    _gridWidth     = contentWidth;
    if (changed)
    {
#ifdef ENABLE_TESTS
        ++_modelRebuildCount;
#endif
        _model.SetPresentation(_gridTechnical, narrow);
        _notifyingRows = true;
        _grid->SetModel(&_model);
        _notifyingRows = false;
    }
    std::vector<DxUi::GridColumnLayoutEntry> layout;
    float minimum = 0.0f;
    float extra   = 0.0f;
    for (size_t i = 0u; i < _model.GetColumnCount(); ++i)
    {
        const auto column = _model.GetColumn(i);
        const auto preference =
            std::find_if(_columnPreferences.begin(), _columnPreferences.end(), [&](const auto& value) { return value.columnId == column.id; });
        auto entry     = preference == _columnPreferences.end() ? DxUi::GridColumnLayoutEntry{column.id, i, column.widthDip} : *preference;
        entry.widthDip = std::max(column.minWidthDip, entry.widthDip);
        minimum += column.minWidthDip;
        extra += entry.widthDip - column.minWidthDip;
        layout.push_back(std::move(entry));
    }
    if (! _gridTechnical)
    {
        const float available = std::max(minimum, contentWidth - 18.0f);
        const float ratio     = extra > 0.0f ? std::max(0.0f, available - minimum) / extra : 0.0f;
        for (size_t i = 0u; i < layout.size(); ++i)
        {
            const auto column  = _model.GetColumn(i);
            layout[i].widthDip = column.minWidthDip + (layout[i].widthDip - column.minWidthDip) * ratio;
        }
    }
    _grid->ApplyColumnLayout(layout);
    _appliedColumnLayout = _grid->CaptureColumnLayout();
    _grid->SetRowHeightDip(_gridTechnical ? kIssuesTechnicalRowHeightDip : kIssuesPrimaryRowHeightDip);
    _grid->SetLineClamp(_gridTechnical ? 2u : 4u);
    ApplyGridSort();
}

void FileOperationsIssuesPaneState::SaveViewState() noexcept
{
    if (! fileOps || ! _grid)
        return;
    CaptureColumnPreferences();
    std::vector<Common::Settings::GridColumnLayoutEntry> saved;
    for (const auto& entry : _columnPreferences)
        saved.push_back({entry.columnId, static_cast<uint32_t>(entry.displayIndex), entry.widthDip});
    const auto sort =
        _sortSpec.direction != SortDirection::None && _sortSpec.columnIndex < _model.KeyCount() ? _model.ColumnByKey(_sortSpec.columnIndex).id : std::wstring{};
    fileOps->SaveIssuesPaneViewState(sort, _sortSpec.direction == SortDirection::Descending, saved);
}

std::optional<size_t> FileOperationsIssuesPaneState::FindRowIndexForTask(uint64_t taskId, uint64_t promptGeneration) const noexcept
{
    return _projection.FindRowIndexForTask(taskId, promptGeneration);
}

bool FileOperationsIssuesPaneState::SelectTask(uint64_t taskId, uint64_t promptGeneration) noexcept
{
    if (! _grid)
        return false;
    RefreshRows(false);
    const std::optional<size_t> rowIndex = FindRowIndexForTask(taskId, promptGeneration);
    if (! rowIndex.has_value())
        return false;

    return SelectRowIndex(rowIndex.value());
}

bool FileOperationsIssuesPaneState::SelectRowIndex(size_t rowIndex) noexcept
{
    if (! _grid || rowIndex >= _rows.size())
        return false;
    _grid->GetSelectionModel().SetSingle(_rows[rowIndex].stableId);
    // SetSingle is a model setter; NotifyDataChanged sees this as the prior selection and does not notify the delegate.
    CaptureSelection();
    _grid->NotifyDataChanged();
    _grid->EnsureRowVisible(rowIndex);
    UpdateDetail(true);
    _dxHost.Invalidate();
    return true;
}

#ifdef ENABLE_TESTS
bool FileOperationsIssuesPaneState::SelfTestSelectDiagnostic(uint64_t taskId, std::wstring_view category, HRESULT status) noexcept
{
    if (! _grid || taskId == 0u || category.empty())
        return false;
    RefreshRows(false);
    const std::wstring expectedStatus = std::format(L"0x{:08X}", static_cast<unsigned long>(status));
    const auto found                  = std::ranges::find_if(
        _rows, [&](const IssuesRow& row) noexcept { return row.taskId == taskId && row.categoryText == category && row.statusText == expectedStatus; });
    return found != _rows.end() && SelectRowIndex(static_cast<size_t>(found - _rows.begin()));
}

void FileOperationsIssuesPaneState::FillSelfTestSnapshot(FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot, bool includeDetail) const noexcept
{
    outSnapshot                    = {};
    outSnapshot.detailPresent      = _detailBody != nullptr;
    outSnapshot.technicalExpanded  = _technicalExpanded;
    outSnapshot.detailScrollOffset = _detailBody ? _detailBody->GetScrollOffset() : 0.0f;
    outSnapshot.detailViewport     = _detailBody ? _detailBody->GetBounds() : D2D1::RectF();
    outSnapshot.copyDetailBounds   = _copyButton ? _copyButton->GetBounds() : D2D1::RectF();
    outSnapshot.technicalBounds    = _technicalButton ? _technicalButton->GetBounds() : D2D1::RectF();
    if (includeDetail)
    {
        outSnapshot.firstRowBounds       = _grid ? _grid->GetVisibleRowRect(0u).value_or(D2D1::RectF()) : D2D1::RectF();
        outSnapshot.detailCopyText       = _detailCopyText;
        outSnapshot.hiddenSortText       = _sortStatus ? _sortStatus->GetText() : std::wstring{};
        outSnapshot.copyCaption          = _copyButton ? _copyButton->GetText() : std::wstring{};
        outSnapshot.technicalCaption     = _technicalButton ? _technicalButton->GetText() : std::wstring{};
        outSnapshot.emptyStateText       = _grid ? _grid->GetEmptyStateText() : std::wstring_view{};
        outSnapshot.detailAccessibleName = _detailBody ? _detailBody->GetAccessibleName() : std::wstring_view{};
        for (size_t column = 0u; column < _model.GetColumnCount(); ++column)
        {
            outSnapshot.columnTitles.push_back(_model.GetColumn(column).title);
            outSnapshot.columnIds.push_back(_model.GetColumn(column).id);
            outSnapshot.columnBounds.push_back(_grid->GetVisibleColumnHeaderRect(column).value_or(D2D1::RectF()));
        }
        for (size_t i = 0u; i < _detailLabels.size(); ++i)
        {
            const auto* label = _detailLabels[i];
            if (! label)
                continue;
            outSnapshot.detailTexts[i]      = label->GetText();
            outSnapshot.detailTextBounds[i] = label->GetBounds();
            const size_t iconSlot           = i == 0u ? 0u : i == 2u ? 1u : i == 3u ? 2u : 3u;
            if (iconSlot < 3u)
            {
                outSnapshot.detailIconGlyphs[i] = _detailBody->IconGlyph(iconSlot, _dxHost);
                outSnapshot.detailIconBounds[i] = _detailBody->IconBounds(iconSlot);
                outSnapshot.detailIconColors[i] = _detailBody->IconColor(iconSlot, _dxHost);
            }
            if (! label->IsVisible())
                continue;
            const auto bounds = label->GetBounds();
            const auto role   = IssuesDetailFontRole(i);
            auto* format      = _dxHost.GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_NEAR, true);
            wil::com_ptr<IDWriteTextLayout> layout;
            const auto text = label->GetText();
            if (_dxHost.GetWriteFactory() && format &&
                SUCCEEDED(_dxHost.GetWriteFactory()->CreateTextLayout(
                    text.data(), static_cast<UINT32>(text.size()), format, std::max(1.0f, bounds.right - bounds.left), 65536.0f, layout.put())))
            {
                DWRITE_TEXT_METRICS metrics{};
                if (SUCCEEDED(layout->GetMetrics(&metrics)))
                    outSnapshot.detailRequiredHeights[i] = metrics.height;
            }
        }
    }
    outSnapshot.detailIconsUseFluentFont      = _dxHost.HasFluentIconFont();
    outSnapshot.rowCount                      = _rows.size();
    outSnapshot.refreshGeneration             = _refreshGeneration;
    outSnapshot.projectionBuildCount          = _projection.BuildCount();
    outSnapshot.projectionSortCount           = _projection.SortCount();
    outSnapshot.issuesSourceRevision          = _projection.SourceRevision();
    outSnapshot.refreshTimerRunning           = _refreshTimerRunning;
    outSnapshot.refreshTimerIntervalMs        = kRefreshTimerIntervalMs;
    outSnapshot.visibleWork                   = _grid ? _grid->GetVisibleWorkMetrics() : DxUi::GridVisibleWorkMetrics{};
    outSnapshot.themeDark                     = _theme.dark;
    outSnapshot.themeHighContrast             = _theme.highContrast;
    outSnapshot.themeRainbow                  = _theme.menu.rainbowMode;
    outSnapshot.dxRenderCount                 = _dxHost.DebugGetRenderCount();
    outSnapshot.surfaceResources              = FileOperationsSelfTest::CaptureSurfaceResources(_dxHost);
    outSnapshot.dxResizeCount                 = _dxHost.DebugGetResizeCount();
    outSnapshot.dxResizeFailureCount          = _dxHost.DebugGetResizeFailureCount();
    outSnapshot.gridFocused                   = _dxHost.GetFocusControl() == _grid;
    outSnapshot.detailFocused                 = _dxHost.GetFocusControl() == _detailBody;
    outSnapshot.technicalFocused              = _dxHost.GetFocusControl() == _technicalButton;
    outSnapshot.copyDetailFocused             = _dxHost.GetFocusControl() == _copyButton;
    outSnapshot.detailRowId                   = _detailRowId;
    outSnapshot.copyDetailEnabled             = _copyButton && _copyButton->IsEnabled();
    outSnapshot.gridSingleSelection           = _grid && _grid->GetSelectionMode() == DxUi::GridSelectionMode::Single;
    outSnapshot.gridNarrow                    = _gridNarrow;
    outSnapshot.presentationApplyCount        = _presentationApplyCount;
    outSnapshot.presentationRebuildCount      = _modelRebuildCount;
    outSnapshot.gridModelAssignmentGeneration = _grid ? _grid->GetModelAssignmentGeneration() : 0u;
    outSnapshot.systemMessageLookupCount      = _systemMessageLookups;
    if (includeDetail)
    {
        for (const auto& row : _rows)
        {
            outSnapshot.rowStableIds.push_back(row.stableId);
            outSnapshot.rowTaskIds.push_back(row.taskId);
        }
    }
    if (! _rows.empty())
    {
        outSnapshot.firstVisibleTaskId = _rows.front().taskId;
    }
    outSnapshot.taskHeaderRect =
        _grid ? _grid->GetVisibleColumnHeaderRect(_model.VisibleIndex(1u).value_or(_model.GetColumnCount())).value_or(D2D1::RectF()) : D2D1::RectF();
    outSnapshot.operationHeaderRect =
        _grid ? _grid->GetVisibleColumnHeaderRect(_model.VisibleIndex(2u).value_or(_model.GetColumnCount())).value_or(D2D1::RectF()) : D2D1::RectF();
    outSnapshot.messageHeaderRect =
        _grid ? _grid->GetVisibleColumnHeaderRect(_model.VisibleIndex(7u).value_or(_model.GetColumnCount())).value_or(D2D1::RectF()) : D2D1::RectF();
    outSnapshot.hasActiveSort   = _sortSpec.direction != SortDirection::None;
    outSnapshot.sortColumnIndex = _sortSpec.columnIndex;
    outSnapshot.sortDescending  = _sortSpec.direction == SortDirection::Descending;

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

    if (includeDetail)
    {
        outSnapshot.primarySelectedRowBounds = _grid->GetVisibleRowRect(rowIndex.value()).value_or(D2D1::RectF());
        for (size_t column = 0u; column < _model.GetColumnCount(); ++column)
        {
            GridCellData cell{};
            _model.GetCellData(rowIndex.value(), column, cell);
            outSnapshot.selectedCells.push_back(cell.text);
        }
    }
    outSnapshot.primarySelectedTaskId      = _rows[rowIndex.value()].taskId;
    const auto& selected                   = _rows[rowIndex.value()];
    outSnapshot.primaryHasAcceptedDecision = selected.acceptedDecision.has_value();
    outSnapshot.primaryHasOriginalProblem  = selected.hasOriginalProblem;
    outSnapshot.primaryHasConflictResult   = selected.conflictResult.has_value();
    outSnapshot.primaryResultIsAmbiguous   = selected.ambiguousResult;
    outSnapshot.primaryResultText          = selected.resultText;
    outSnapshot.primaryProblemText         = selected.problemText;
    outSnapshot.primaryDecisionAction      = selected.acceptedDecision.has_value() ? static_cast<uint32_t>(selected.acceptedDecision->action) : 0u;
    outSnapshot.primaryDecisionIsCached =
        selected.acceptedDecision.has_value() && selected.acceptedDecision->origin == FileOperations::ConflictDecisionOrigin::CachedScope;
    DxUi::GridDebugRowVisualState rowVisualState{};
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
    if (! _grid->DebugHitTestPoint(DxUi::MakePointDip(pointDip), hit))
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
        static_cast<void>(Common::RequestForegroundWindow(hwnd));
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

void FileOperationsIssuesPaneState::UpdateDetail(bool resetScroll) noexcept
{
    if (_tearingDown || ! _detailBody || ! _grid)
        return;
    const auto selection      = _grid->GetSelectionModel().GetOrderedSelection();
    const auto index          = selection.empty() ? std::optional<size_t>{} : _model.FindRowByStableId(selection.front());
    const IssuesRow* row      = index.has_value() ? &_rows[index.value()] : nullptr;
    const uint64_t selectedId = row ? row->stableId : 0u;
    if (resetScroll || selectedId != _detailRowId)
        _detailBody->ScrollToTop();
    if (selectedId != _detailRowId)
    {
        _copyFeedbackResourceId = 0u;
        _copyFeedback->SetText(L"");
    }
    _detailRowId = selectedId;
    _copyButton->SetEnabled(row != nullptr);
    _technicalButton->SetEnabled(true);
    const auto section = [](UINT labelId, std::wstring_view value) { return std::format(L"{}\n{}", LoadStringResource(nullptr, labelId), value); };
    std::array<std::wstring, 9> texts;
    if (! row)
        texts[0] = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_SELECT_DETAIL);
    else
    {
        if (! row->sourceName.empty() && ! row->destinationName.empty() && row->sourceName != row->destinationName)
            texts[8] = section(IDS_FILEOPS_ISSUES_COL_SOURCE, row->sourceName) + L"\n" + section(IDS_FILEOPS_ISSUES_COL_DESTINATION, row->destinationName);
        else
            texts[8] = ! row->sourceName.empty() ? row->sourceName : row->destinationName;
        texts[0] = section(IDS_FILEOPS_ISSUES_PROBLEM, row->problemText);
        texts[1] = section(IDS_FILEOPS_ISSUES_COL_OPERATION, row->operationText);
        // issues#9: an ordinary failure has no decision or result to report; its Problem is the
        // failure. Conflict rows keep explicit "not recorded" / "unknown" facts.
        std::wstring choice = row->hasConflictFacts ? LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_DECISION_UNKNOWN) : std::wstring{};
        if (row->acceptedDecision.has_value())
        {
            const auto& accepted = row->acceptedDecision.value();
            choice               = FileOperations::FormatConflictActionLabel(accepted.action, accepted.scope.bucket, row->operation);
            choice += L"\n" + LoadStringResource(nullptr,
                                                 accepted.origin == FileOperations::ConflictDecisionOrigin::CachedScope ? IDS_FILEOPS_ISSUES_ORIGIN_CACHED
                                                                                                                        : IDS_FILEOPS_ISSUES_ORIGIN_USER);
            if (accepted.origin == FileOperations::ConflictDecisionOrigin::CachedScope && accepted.appliedCount != 0u)
                choice += L"\n" + FormatStringResource(nullptr, IDS_FMT_FILEOPS_ISSUES_AUTOMATIC_COUNT, accepted.appliedCount);
            choice += L"\n" + LoadStringResource(nullptr, accepted.applyToAll ? IDS_FILEOPS_ISSUES_SCOPE_SIMILAR : IDS_FILEOPS_ISSUES_SCOPE_ITEM);
        }
        if (row->hasConflictFacts)
        {
            texts[2] = section(IDS_FILEOPS_ISSUES_DECISION, choice);
            texts[3] = section(IDS_FILEOPS_ISSUES_RESULT, row->resultText);
        }
        // A path the operation does not have (a Delete's destination) has no section, not "-".
        if (! row->sourceName.empty() || ! row->sourcePathText.empty())
            texts[4] = section(IDS_FILEOPS_ISSUES_COL_SOURCE, row->sourceName.empty() ? row->sourcePathText : row->sourceParentPath);
        if (! row->destinationName.empty() || ! row->destinationPathText.empty())
            texts[5] = section(IDS_FILEOPS_ISSUES_COL_DESTINATION, row->destinationName.empty() ? row->destinationPathText : row->destinationParentPath);
        // The engine's own message (English, technical) appears only here, in logs and in exports.
        texts[6] = std::format(L"{}: {}\n{}: {}\n{}: {}\n{}: {}\n{}: {}\n{}: {}\n{}: {}",
                               LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_TIME),
                               row->timeText,
                               LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_TASK),
                               row->taskText,
                               LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_SEVERITY),
                               row->severityText,
                               LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_HRESULT),
                               row->statusText,
                               LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_STATUS_TEXT),
                               row->statusTextDetail,
                               LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_CATEGORY),
                               row->categoryText,
                               LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_MESSAGE),
                               row->messageText);
        // Raw scope provenance is technical data, not a translated policy inferred from paths.
        if (row->acceptedDecision.has_value())
        {
            const auto& accepted = row->acceptedDecision.value();
            texts[7]             = FileOperations::FormatAcceptedConflictDecisionDiagnostic(accepted);
        }
        if (row->conflictResult.has_value())
        {
            if (! texts[7].empty())
                texts[7] += L"\n";
            texts[7] += FileOperations::FormatConflictResultDiagnostic(row->conflictResult.value());
        }
    }
    _detailLabels[0]->SetTextColor(
        row && row->hasOriginalProblem && row->severity != FolderWindow::FileOperationState::DiagnosticSeverity::Info && ! _theme.highContrast
            ? std::optional<D2D1_COLOR_F>{row->severity == FolderWindow::FileOperationState::DiagnosticSeverity::Error ? _theme.folderView.errorText
                                                                                                                       : _theme.folderView.warningText}
            : std::nullopt);
    // The informational palette color is paired with its own banner fill, not
    // this ordinary detail background. Keep accepted choices in the readable
    // normal foreground; strong type and explicit labels distinguish the state.
    _detailLabels[2]->SetTextColor(std::nullopt);
    const auto neutral        = _theme.folderView.textNormal;
    wchar_t problemGlyph      = 0;
    wchar_t problemFallback   = 0;
    D2D1_COLOR_F problemColor = neutral;
    if (row)
    {
        if (! row->hasOriginalProblem)
        {
            problemGlyph    = 0;
            problemFallback = L'?';
        }
        else if (row->severity == FolderWindow::FileOperationState::DiagnosticSeverity::Info)
        {
            // I35: an Information note uses the neutral information glyph.
            problemGlyph    = FluentIcons::kInfo;
            problemFallback = L'i';
        }
        else if (row->severity == FolderWindow::FileOperationState::DiagnosticSeverity::Error)
        {
            problemGlyph    = FluentIcons::kError;
            problemFallback = FluentIcons::kFallbackError;
            problemColor    = _theme.folderView.errorText;
        }
        else
        {
            problemGlyph    = FluentIcons::kWarning;
            problemFallback = FluentIcons::kFallbackWarning;
            problemColor    = _theme.folderView.warningText;
        }
    }
    _detailBody->SetIcon(0u, _detailLabels[0], problemGlyph, problemFallback, problemColor);
    _detailBody->SetIcon(1u,
                         _detailLabels[2],
                         row && row->acceptedDecision ? FluentIcons::kForward : 0,
                         row && row->acceptedDecision ? FluentIcons::kFallbackForward : L'?',
                         row && row->acceptedDecision ? _theme.accent : neutral);
    wchar_t resultGlyph      = 0;
    wchar_t resultFallback   = L'?';
    D2D1_COLOR_F resultColor = neutral;
    if (row && row->conflictResult)
    {
        const auto& result                = row->conflictResult.value();
        const bool incompleteVerification = result.verification == FileOperations::VerificationState::Failed ||
                                            result.verification == FileOperations::VerificationState::Unavailable ||
                                            result.verification == FileOperations::VerificationState::Canceled;
        const bool transfer               = row->operation == FILESYSTEM_COPY || row->operation == FILESYSTEM_MOVE || row->operation == FILESYSTEM_RENAME;
        const bool removalRequired        = row->operation == FILESYSTEM_MOVE || row->operation == FILESYSTEM_RENAME || row->operation == FILESYSTEM_DELETE;
        const bool removalConfirmed       = result.sourceDisposition == FileOperations::SourceDisposition::Removed;
        // Delete has no destination publication. Its typed source-removal fact is
        // the mutation evidence; unknown operation/outcome never gains a check.
        const bool mutationConfirmed =
            row->operation == FILESYSTEM_DELETE ? removalConfirmed : transfer && result.publication == FileOperations::PublicationState::Published;
        const bool incompleteMutation = (transfer && (result.publication == FileOperations::PublicationState::NotPublished ||
                                                      result.publication == FileOperations::PublicationState::NotAttempted)) ||
                                        (removalRequired && result.sourceDisposition == FileOperations::SourceDisposition::Retained);
        if (result.completion == FileOperations::ItemCompletion::Completed && mutationConfirmed && ! incompleteVerification &&
            (! removalRequired || removalConfirmed))
        {
            resultGlyph    = FluentIcons::kCheckMark;
            resultFallback = FluentIcons::kFallbackCheckMark;
            resultColor    = _theme.fileOperations.successText;
        }
        else if (result.completion == FileOperations::ItemCompletion::Failed)
        {
            resultGlyph    = FluentIcons::kError;
            resultFallback = FluentIcons::kFallbackError;
            resultColor    = _theme.folderView.errorText;
        }
        else if (result.completion == FileOperations::ItemCompletion::Skipped || result.completion == FileOperations::ItemCompletion::Canceled ||
                 incompleteVerification || (result.completion == FileOperations::ItemCompletion::Completed && incompleteMutation))
        {
            resultGlyph    = FluentIcons::kWarning;
            resultFallback = FluentIcons::kFallbackWarning;
            resultColor    = _theme.folderView.warningText;
        }
    }
    _detailBody->SetIcon(2u, _detailLabels[3], resultGlyph, resultFallback, resultColor);
    // issues#7: labels wrap on "\n"; the clipboard text uses CRLF throughout, including any line
    // break inside a provider path or engine message.
    const auto appendCopyText = [this](std::wstring_view text)
    {
        for (size_t index = 0u; index < text.size(); ++index)
        {
            if (text[index] == L'\n' && (index == 0u || text[index - 1u] != L'\r'))
                _detailCopyText.push_back(L'\r');
            _detailCopyText.push_back(text[index]);
        }
    };
    _detailCopyText.clear();
    for (const size_t i : kIssuesDetailOrder)
    {
        _detailLabels[i]->SetText(texts[i]);
        _detailLabels[i]->SetVisible(! texts[i].empty() && (i == 8u || i < 6u || _technicalExpanded));
        if (! texts[i].empty())
        {
            if (! _detailCopyText.empty())
                _detailCopyText += L"\r\n\r\n";
            // Display removes duplicate leaf names; copy retains each exact provider path.
            if (row && i == 4u)
                appendCopyText(section(IDS_FILEOPS_ISSUES_COL_SOURCE, row->sourcePathText));
            else if (row && i == 5u)
                appendCopyText(section(IDS_FILEOPS_ISSUES_COL_DESTINATION, row->destinationPathText));
            else
                appendCopyText(texts[i]);
        }
    }
    // issues#8: the focusable detail region names the selected issue, not the selection instruction.
    std::wstring accessibleName;
    if (! row)
        accessibleName = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_SELECT_DETAIL);
    else
    {
        const std::wstring_view subject = std::wstring_view(row->subjectText).substr(0u, row->subjectText.find(L'\n'));
        accessibleName = subject.empty()
                             ? row->problemText
                             : FormatEmbeddedStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_FACT_JOIN, subject, std::wstring_view(row->problemText));
    }
    if (_detailBody->GetAccessibleName() != accessibleName)
        _detailBody->SetAccessibleName(std::move(accessibleName));
    _technicalButton->SetText(LoadStringResource(nullptr, _technicalExpanded ? IDS_FILEOPS_ISSUES_HIDE_TECHNICAL : IDS_FILEOPS_ISSUES_TECHNICAL));
    Layout();
    _dxHost.Invalidate();
}

bool FileOperationsIssuesPaneState::CopyDetail() noexcept
{
    if (_tearingDown || ! _copyFeedback || _detailRowId == 0u)
        return false;
    const bool copied       = Common::Clipboard::TrySetUnicodeText(_dxHost.GetHwnd(), _detailCopyText, Common::Clipboard::EmptyUnicodeTextPolicy::Reject);
    _copyFeedbackResourceId = copied ? IDS_FILEOPS_ISSUES_COPY_SUCCEEDED : IDS_FILEOPS_ISSUES_COPY_FAILED;
    _copyFeedback->SetText(LoadStringResource(nullptr, _copyFeedbackResourceId));
    Layout();
    _dxHost.Invalidate();
    return copied;
}

void FileOperationsIssuesPaneState::Layout() noexcept
{
    if (! _root || ! _grid || ! _detailBody || ! _dxHost.GetHwnd())
        return;
    RECT client{};
    GetClientRect(_dxHost.GetHwnd(), &client);
    // issues#5: a minimized window has an empty client area. Laying it out would switch the grid
    // to the narrow presentation and rebuild its model on every minimize and restore.
    if (IsIconic(_dxHost.GetHwnd()) != FALSE || client.right <= client.left)
        return;
    const float widthDip  = _dxHost.PixelsToDip(static_cast<float>(std::max(0L, client.right - client.left)));
    const float heightDip = _dxHost.PixelsToDip(static_cast<float>(std::max(0L, client.bottom - client.top)));
    _root->SetBounds(D2D1::RectF(0.0f, 0.0f, widthDip, heightDip));
    constexpr float gap      = 8.0f;
    const float right        = std::max(gap + 1.0f, widthDip - gap);
    const float contentWidth = std::max(1.0f, right - gap);
    const auto measure       = [&](std::wstring_view text, float width, DxUi::FontRole role)
    {
        const int pixels = std::max(1, static_cast<int>(std::floor(_dxHost.DipsToPixels(width))));
        return _dxHost.PixelsToDip(static_cast<float>(DxUi::Typography::MeasureWrappedTextHeightPx(_dxHost.GetHwnd(), role, pixels, text)));
    };
    ApplyGridPresentation(contentWidth);
    const float buttonWidth  = std::max(1.0f, (contentWidth - gap) * 0.5f);
    const float buttonHeight = std::max(36.0f,
                                        std::max(measure(_technicalButton->GetText(), buttonWidth - 32.0f, DxUi::FontRole::Body),
                                                 measure(_copyButton->GetText(), buttonWidth - 24.0f, DxUi::FontRole::Body)) +
                                            16.0f);
    _copyFeedback->SetVisible(! _copyFeedback->GetText().empty());
    const float feedbackHeight = _copyFeedback->GetText().empty() ? 0.0f : measure(_copyFeedback->GetText(), contentWidth, DxUi::FontRole::Body) + gap;
    const float footerTop      = std::max(gap, heightDip - gap - buttonHeight - feedbackHeight);
    const float sortHeight     = _sortStatus->GetText().empty() ? 0.0f : measure(_sortStatus->GetText(), contentWidth, DxUi::FontRole::Body) + gap;
    // The proportional split must not cut the first summary row. Keep the
    // hidden-sort explanation above the header; complete selected detail can
    // scroll independently below the grid while the actions remain fixed.
    const float rowHeight         = _gridTechnical ? kIssuesTechnicalRowHeightDip : kIssuesPrimaryRowHeightDip;
    const float minimumGridBottom = gap + sortHeight + kIssuesHeaderHeightDip + rowHeight + 2.0f;
    const float gridBottom        = std::clamp(heightDip * 0.38f, minimumGridBottom, std::max(minimumGridBottom, footerTop - 72.0f));
    _sortStatus->SetBounds(D2D1::RectF(gap, gap, right, gap + sortHeight));
    _grid->SetBounds(D2D1::RectF(gap, gap + sortHeight, right, gridBottom));
    _detailBody->SetBounds(D2D1::RectF(gap, gridBottom + gap, right, std::max(gridBottom + gap + 1.0f, footerTop - gap)));
    const float detailWidth = std::max(1.0f, contentWidth - _detailBody->GetScrollbarThickness() - gap * 2.0f);
    // DxUi controls use host coordinates. ScrollPanel translates only the scroll
    // offset; it does not add its origin to child bounds.
    const auto detailBounds = _detailBody->GetBounds();
    float y                 = detailBounds.top + gap;
    for (const size_t i : kIssuesDetailOrder)
    {
        auto* label = _detailLabels[i];
        if (! label->IsVisible())
            continue;
        const auto role        = IssuesDetailFontRole(i);
        const float iconGutter = (i == 0u || i == 2u || i == 3u) && _detailBody->IconGlyph(i == 0u   ? 0u
                                                                                           : i == 2u ? 1u
                                                                                                     : 2u,
                                                                                           _dxHost) != 0
                                     ? kIssuesDetailIconGutterDip
                                     : 0.0f;
        const float textWidth  = std::max(1.0f, detailWidth - iconGutter);
        const float height     = std::max(20.0f, measure(label->GetText(), textWidth, role));
        label->SetBounds(D2D1::RectF(detailBounds.left + gap + iconGutter, y, detailBounds.left + gap + iconGutter + textWidth, y + height));
        y += height + gap * 2.0f;
    }
    _detailBody->SetContentHeight(y - detailBounds.top);
    _technicalButton->SetBounds(D2D1::RectF(gap, footerTop, gap + buttonWidth, footerTop + buttonHeight));
    _copyButton->SetBounds(D2D1::RectF(gap * 2.0f + buttonWidth, footerTop, right, footerTop + buttonHeight));
    _copyFeedback->SetBounds(D2D1::RectF(gap, footerTop + buttonHeight + gap, right, heightDip - gap));
}

LRESULT FileOperationsIssuesPaneState::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);
    if (! _dxHost.Attach(hwnd))
    {
        return -1;
    }

#ifdef ENABLE_TESTS
    if (g_failNextIssuesPaneCreation.exchange(false, std::memory_order_relaxed))
        return -1; // Exercise failed WM_CREATE after a real host attachment.
#endif

    _rootStorage = std::make_unique<Panel>();
    _root        = _rootStorage.get();
    _sortStatus  = _root->AddChild<DxUi::Label>();
    _sortStatus->SetMultiline(true);
    _grid = _root->AddChild<Grid>();
    if (_grid)
    {
        _grid->SetEmptyStateText(LoadStringResource(nullptr, IDS_DXUI_NO_DATA));
    }
    _grid->SetDelegate(this);
    // issues#1: the detail shows one issue, so the grid selects one row: its current row.
    _grid->SetSelectionMode(DxUi::GridSelectionMode::Single);
    _grid->SetModel(&_model);
    _grid->SetHeaderHeightDip(kIssuesHeaderHeightDip);
    _grid->SetRowHeightDip(kIssuesTechnicalRowHeightDip);
    _grid->SetLineClamp(2u);
    _detailBody = _root->AddChild<IssuesDetailBody>();
    _detailBody->SetAccessibleName(LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_SELECT_DETAIL));
    _detailBody->copy = [this] { return CopyDetail(); };
    for (const size_t i : kIssuesDetailOrder)
    {
        _detailLabels[i] = _detailBody->AddChild<DxUi::Label>();
        _detailLabels[i]->SetMultiline(true);
        _detailLabels[i]->SetFontRole(IssuesDetailFontRole(i));
    }
    _technicalButton = _root->AddChild<DxUi::Button>(LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_TECHNICAL));
    _technicalButton->SetMultiline(true);
    _technicalButton->SetVariant(DxUi::ButtonVariant::Standard);
    _technicalButton->SetOnClick([this]
    {
        _technicalExpanded = ! _technicalExpanded;
        UpdateDetail(false);
    });
    _copyButton = _root->AddChild<DxUi::Button>(LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COPY_DETAIL));
    _copyButton->SetMultiline(true);
    _copyButton->SetOnClick([this] { static_cast<void>(CopyDetail()); });
    _copyFeedback = _root->AddChild<DxUi::Label>();
    _copyFeedback->SetMultiline(true);
    _dxHost.SetRoot(std::move(_rootStorage));
    ApplySavedViewState();

    ApplyTheme(hwnd);
    Layout();
    RefreshRows(true);
    // issues#4: an empty first projection publishes no rows; the instruction and the disabled
    // Copy action must not depend on it.
    UpdateDetail(true);
    _languageRevision = Localization::GetLanguagePreferenceRevision();
    UpdateRefreshTimer(hwnd);
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnSize(HWND hwnd, UINT sizeType, UINT width, UINT height) noexcept
{
    static_cast<void>(width);
    static_cast<void>(height);
    if (sizeType != SIZE_MINIMIZED)
        Layout();
    if (sizeType == SIZE_MINIMIZED)
        UpdateRefreshTimer(hwnd);
    else if (IsWindowVisible(hwnd))
    {
        RefreshRows(false);
        UpdateRefreshTimer(hwnd);
    }
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
    if (! IsWindowVisible(hwnd) || IsIconic(hwnd))
    {
        UpdateRefreshTimer(hwnd);
        return 0;
    }
    RefreshRows(false);
    return 0;
}

void FileOperationsIssuesPaneState::UpdateRefreshTimer(HWND hwnd) noexcept
{
    if (! hwnd)
        return;
    const bool shouldRun = IsWindowVisible(hwnd) && ! IsIconic(hwnd) && ! _tearingDown;
    if (! shouldRun)
    {
        if (_refreshTimerRunning)
        {
            KillTimer(hwnd, kRefreshTimerId);
            _refreshTimerRunning = false;
        }
        return;
    }
    if (! _refreshTimerRunning)
    {
        _refreshTimerRunning = SetTimer(hwnd, kRefreshTimerId, kRefreshTimerIntervalMs, nullptr) != 0u;
    }
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
        RefreshRows(false);
        ApplyTheme(hwnd);
    }
    else if (HasLiveOwnerState())
    {
        SaveViewState();
    }
    if (! visible || IsIconic(hwnd) || _tearingDown)
    {
        if (_refreshTimerRunning)
        {
            KillTimer(hwnd, kRefreshTimerId);
            _refreshTimerRunning = false;
        }
    }
    else if (! _refreshTimerRunning)
        _refreshTimerRunning = SetTimer(hwnd, kRefreshTimerId, kRefreshTimerIntervalMs, nullptr) != 0u;
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
    _inThemeChange            = true;
    const auto clearThemeFlag = wil::scope_exit([&] { _inThemeChange = false; });
    ApplyTheme(hwnd);
    const uint64_t languageRevision = Localization::GetLanguagePreferenceRevision();
    if (_languageRevision == languageRevision)
    {
        UpdateDetail(false);
        return 0;
    }
    const auto title = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_PANE_TITLE);
    if (Win32Text::GetWindowTextString(hwnd) != title)
        SetWindowTextW(hwnd, title.c_str());
    if (_grid && _detailBody && _technicalButton && _copyButton && _copyFeedback && _sortStatus)
    {
        // Settings reload applies the new satellite before this notification.
        // Retain the model and column layout; SetModel would reset grid state.
        const bool headingsChanged = _model.ReloadColumnTitles();
        _grid->SetEmptyStateText(LoadStringResource(nullptr, IDS_DXUI_NO_DATA));
        // UpdateDetail below (directly or through RefreshRows) renames the detail region.
        _copyButton->SetText(LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COPY_DETAIL));
        if (_copyFeedbackResourceId != 0u)
            _copyFeedback->SetText(LoadStringResource(nullptr, _copyFeedbackResourceId));
        ApplyGridSort();
        const auto previousGeneration = _refreshGeneration;
        RefreshRows(true);
        if (_refreshGeneration == previousGeneration)
        {
            // Headings and UIA names can change even with an empty/unchanged model.
            if (headingsChanged)
                _grid->NotifyDataChanged();
            UpdateDetail(false);
        }
        _languageRevision = languageRevision;
    }
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnNcDestroy(HWND hwnd) noexcept
{
    _tearingDown = true;
    KillTimer(hwnd, kRefreshTimerId);
    if (HasLiveOwnerState())
    {
        SaveViewState();
        fileOps->OnIssuesPaneDestroyed(hwnd);
    }
    _dxHost.Detach();
    _root            = nullptr;
    _grid            = nullptr;
    _detailBody      = nullptr;
    _technicalButton = nullptr;
    _sortStatus      = nullptr;
    _copyButton      = nullptr;
    _copyFeedback    = nullptr;
    _detailLabels.fill(nullptr);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    _deletePending = true;
    if (_dispatchDepth == 0u && ! _creationInProgress)
    {
        delete this;
    }
    return 0;
}

LRESULT FileOperationsIssuesPaneState::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    // Snapshot application view state before Detach destroys the retained controls.
    // Detach itself drains host payloads and disconnects accessibility providers.
    if (msg == WM_NCDESTROY)
        return OnNcDestroy(hwnd);

    bool dxHandled         = false;
    const LRESULT dxResult = _dxHost.HandleMessage(hwnd, msg, wp, lp, dxHandled);
    if (dxHandled)
    {
        switch (msg)
        {
            case WM_SIZE: return OnSize(hwnd, static_cast<UINT>(wp), LOWORD(lp), HIWORD(lp));
            // The host already applied the suggested rectangle; the adapter still owns reflow.
            case WM_DPICHANGED: return OnDpiChanged(hwnd, static_cast<UINT>(HIWORD(wp)), nullptr);
            case WM_DPICHANGED_AFTERPARENT: return OnDpiChanged(hwnd, GetDpiForWindow(hwnd), nullptr);

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
        case WM_SIZE: return OnSize(hwnd, static_cast<UINT>(wp), LOWORD(lp), HIWORD(lp));
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
        if (state->_dispatchDepth == 0u && state->_deletePending && ! state->_creationInProgress)
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
    auto* state              = statePtr.get();
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
    if (! pane || statePtr->IsPendingDelete())
        return nullptr;
    // Until CreateWindowEx returns, the factory owns all failure-path cleanup.
    statePtr->CompleteCreation();
    static_cast<void>(statePtr.release());

    ShowWindow(pane, startMaximized ? SW_SHOWMAXIMIZED : SW_SHOWNORMAL);
    UpdateWindow(pane);
    return pane;
}

bool FileOperationsIssuesPane::SelectTask(HWND hwnd, uint64_t taskId, uint64_t promptGeneration) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
        return false;
    auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return state && state->SelectTask(taskId, promptGeneration);
}

#ifdef ENABLE_TESTS
uint64_t FileOperationsIssuesPane::SelfTestLiveStateCount() noexcept
{
    return g_liveIssuesPaneStates.load(std::memory_order_relaxed);
}
void FileOperationsIssuesPane::SelfTestFailNextCreation(bool fail) noexcept
{
    g_failNextIssuesPaneCreation.store(fail, std::memory_order_relaxed);
}
bool FileOperationsIssuesPane::TryGetSelfTestSnapshot(HWND hwnd, SelfTestSnapshot& outSnapshot, bool includeDetail) noexcept
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

    state->FillSelfTestSnapshot(outSnapshot, includeDetail);
    return true;
}

bool FileOperationsIssuesPane::SelfTestDetailKey(HWND hwnd, UINT key) noexcept
{
    if (! hwnd || ! IsWindow(hwnd))
        return false;
    auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return state && state->SelfTestDetailKey(key);
}

bool FileOperationsIssuesPane::SelfTestSetTechnicalExpanded(HWND hwnd, bool expanded) noexcept
{
    if (! hwnd || ! IsWindow(hwnd))
        return false;
    auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return state && state->SelfTestSetTechnicalExpanded(expanded);
}

bool FileOperationsIssuesPane::SelfTestSelectTask(HWND hwnd, uint64_t taskId, uint64_t promptGeneration) noexcept
{
    return SelectTask(hwnd, taskId, promptGeneration);
}

bool FileOperationsIssuesPane::SelfTestSelectDiagnostic(HWND hwnd, uint64_t taskId, std::wstring_view category, HRESULT status) noexcept
{
    if (! hwnd || ! IsWindow(hwnd))
        return false;
    auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return state && state->SelfTestSelectDiagnostic(taskId, category, status);
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

bool FileOperationsIssuesPane::SelfTestRequestSelectRow(HWND hwnd, size_t rowIndex, UINT modifiers) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
        return false;
    auto* state = reinterpret_cast<FileOperationsIssuesPaneState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return state && state->SelfTestRequestSelectRow(rowIndex, modifiers);
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
