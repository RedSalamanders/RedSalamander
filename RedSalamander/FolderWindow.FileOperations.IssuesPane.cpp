#include "FolderWindow.FileOperations.IssuesPane.h"

#include "FolderWindow.FileOperationsInternal.h"
#include "ThemedControls.h"
#include "WindowMaximizeBehavior.h"

#include <array>
#include <bit>
#include <commctrl.h>
#include <format>
#include <uxtheme.h>
#include <vector>
#include <windowsx.h>

namespace
{
constexpr wchar_t kFileOperationsIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
constexpr int kIssuesListControlId                     = 1;
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
    wchar_t buffer[512]{};
    DWORD messageId = std::bit_cast<DWORD>(hr);
    if (HRESULT_FACILITY(hr) == FACILITY_WIN32)
    {
        const DWORD code = HRESULT_CODE(static_cast<DWORD>(hr));
        if (code != 0)
        {
            messageId = code;
        }
    }

    const DWORD written = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                                         nullptr,
                                         messageId,
                                         MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                                         buffer,
                                         static_cast<DWORD>(std::size(buffer)),
                                         nullptr);
    if (written == 0)
    {
        return std::format(L"HRESULT 0x{:08X}", static_cast<unsigned long>(hr));
    }

    std::wstring result(buffer, written);
    while (! result.empty())
    {
        const wchar_t ch = result.back();
        if (ch != L'\r' && ch != L'\n' && ch != L' ' && ch != L'\t')
        {
            break;
        }
        result.pop_back();
    }
    if (result.empty())
    {
        return std::format(L"HRESULT 0x{:08X}", static_cast<unsigned long>(hr));
    }

    return result;
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
        if (lhs[i].severity != rhs[i].severity || lhs[i].taskId != rhs[i].taskId || lhs[i].timeText != rhs[i].timeText ||
            lhs[i].statusText != rhs[i].statusText || lhs[i].statusTextDetail != rhs[i].statusTextDetail || lhs[i].categoryText != rhs[i].categoryText ||
            lhs[i].messageText != rhs[i].messageText || lhs[i].sourcePathText != rhs[i].sourcePathText ||
            lhs[i].destinationPathText != rhs[i].destinationPathText)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] D2D1::ColorF RainbowRowColor(const AppTheme& theme, const IssuesRow& row) noexcept
{
    const std::wstring_view seed = ! row.messageText.empty() ? std::wstring_view(row.messageText) : std::wstring_view(row.categoryText);
    const uint32_t hash          = StableHash32(seed);
    const float hue              = static_cast<float>((hash + static_cast<uint32_t>(row.taskId & 0xFFFFu)) % 360u);
    const float saturation       = theme.dark ? 0.35f : 0.28f;
    const float value            = theme.dark ? 0.34f : 0.96f;
    return ColorFromHSV(hue, saturation, value, 1.0f);
}

class FileOperationsIssuesPaneState final
{
public:
    FileOperationsIssuesPaneState()                                                = default;
    FileOperationsIssuesPaneState(const FileOperationsIssuesPaneState&)            = delete;
    FileOperationsIssuesPaneState(FileOperationsIssuesPaneState&&)                 = delete;
    FileOperationsIssuesPaneState& operator=(const FileOperationsIssuesPaneState&) = delete;
    FileOperationsIssuesPaneState& operator=(FileOperationsIssuesPaneState&&)      = delete;

    FolderWindow::FileOperationState* fileOps = nullptr;
    FolderWindow* folderWindow                = nullptr;

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

private:
    void ApplyTheme(HWND hwnd) noexcept;
    void ApplyColumnLayout() noexcept;
    void EnsureColumns() noexcept;
    void RefreshRows(bool force) noexcept;
    std::vector<IssuesRow> BuildRows() const;
    void RebuildList() noexcept;

    LRESULT OnCreate(HWND hwnd) noexcept;
    LRESULT OnEraseBkgnd() noexcept;
    LRESULT OnPaint(HWND hwnd) noexcept;
    LRESULT OnSize(HWND hwnd, UINT width, UINT height) noexcept;
    LRESULT OnNotify(NMHDR* header) noexcept;
    LRESULT OnTimer(HWND hwnd, UINT_PTR timerId) noexcept;
    LRESULT OnMove(HWND hwnd) noexcept;
    LRESULT OnExitSizeMove(HWND hwnd) noexcept;
    LRESULT OnShowWindow(HWND hwnd, BOOL visible) noexcept;
    LRESULT OnClose(HWND hwnd) noexcept;
    LRESULT OnDpiChanged(HWND hwnd, UINT dpi, const RECT* suggested) noexcept;
    LRESULT OnThemeChanged(HWND hwnd) noexcept;
    LRESULT OnNcDestroy(HWND hwnd) noexcept;

    UINT _dpi                  = USER_DEFAULT_SCREEN_DPI;
    bool _inThemeChange        = false;
    bool _inTitleBarThemeApply = false;
    AppTheme _theme{};
    wil::unique_hbrush _backgroundBrush;
    wil::unique_hwnd _list;
    std::vector<IssuesRow> _rows;
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

    if (folderWindow)
    {
        _theme = folderWindow->GetTheme();
        if (! _inTitleBarThemeApply)
        {
            _inTitleBarThemeApply = true;
            ApplyTitleBarTheme(hwnd, _theme, GetActiveWindow() == hwnd);
            _inTitleBarThemeApply = false;
        }
    }

    _backgroundBrush.reset(CreateSolidBrush(_theme.windowBackground));

    if (_list)
    {
        ThemedControls::ApplyThemeToListView(_list.get(), _theme);
        InvalidateRect(_list.get(), nullptr, TRUE);
    }
    InvalidateRect(hwnd, nullptr, TRUE);
}

void FileOperationsIssuesPaneState::ApplyColumnLayout() noexcept
{
    if (! _list)
    {
        return;
    }

    const int kWidthsDip[] = {
        170,
        70,
        80,
        80,
        100,
        220,
        130,
        280,
        300,
        300,
    };

    for (int i = 0; i < static_cast<int>(std::size(kWidthsDip)); ++i)
    {
        ListView_SetColumnWidth(_list.get(), i, DipsToPixels(kWidthsDip[i], _dpi));
    }
}

void FileOperationsIssuesPaneState::EnsureColumns() noexcept
{
    if (! _list)
    {
        return;
    }

    ListView_DeleteAllItems(_list.get());

    while (ListView_DeleteColumn(_list.get(), 0) != FALSE)
    {
    }

    struct ColumnDef
    {
        UINT titleId;
        int widthDip;
    };

    static constexpr ColumnDef kColumns[] = {
        {IDS_FILEOPS_ISSUES_COL_TIME, 170},
        {IDS_FILEOPS_ISSUES_COL_TASK, 70},
        {IDS_FILEOPS_ISSUES_COL_OPERATION, 80},
        {IDS_FILEOPS_ISSUES_COL_SEVERITY, 80},
        {IDS_FILEOPS_ISSUES_COL_HRESULT, 100},
        {IDS_FILEOPS_ISSUES_COL_STATUS_TEXT, 220},
        {IDS_FILEOPS_ISSUES_COL_CATEGORY, 130},
        {IDS_FILEOPS_ISSUES_COL_MESSAGE, 280},
        {IDS_FILEOPS_ISSUES_COL_SOURCE, 300},
        {IDS_FILEOPS_ISSUES_COL_DESTINATION, 300},
    };

    for (int i = 0; i < static_cast<int>(std::size(kColumns)); ++i)
    {
        std::wstring title = LoadStringResource(nullptr, kColumns[i].titleId);
        LVCOLUMNW column{};
        column.mask    = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT;
        column.fmt     = LVCFMT_LEFT;
        column.cx      = DipsToPixels(kColumns[i].widthDip, _dpi);
        column.pszText = title.empty() ? const_cast<wchar_t*>(L"") : title.data();
        ListView_InsertColumn(_list.get(), i, &column);
    }
}

std::vector<IssuesRow> FileOperationsIssuesPaneState::BuildRows() const
{
    std::vector<IssuesRow> rows;
    if (! fileOps)
    {
        return rows;
    }

    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completed;
    fileOps->CollectCompletedTasks(completed);

    for (const auto& task : completed)
    {
        const std::wstring operationText = LoadStringResource(nullptr, OperationStringId(task.operation));

        for (const auto& issue : task.issueDiagnostics)
        {
            if (issue.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Info)
            {
                continue;
            }

            IssuesRow row{};
            row.severity      = issue.severity;
            row.taskId        = issue.taskId;
            row.timeText      = FormatTimeText(issue.localTime);
            row.taskText      = std::to_wstring(static_cast<unsigned long long>(issue.taskId));
            row.operationText = operationText;
            row.severityText  = LoadStringResource(
                nullptr, issue.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Error ? IDS_CAPTION_ERROR : IDS_CAPTION_WARNING);
            row.statusText          = std::format(L"0x{:08X}", static_cast<unsigned long>(issue.status));
            row.statusTextDetail    = FormatStatusText(issue.status);
            row.categoryText        = issue.category.empty() ? L"-" : issue.category;
            row.messageText         = issue.message.empty() ? L"-" : issue.message;
            row.sourcePathText      = issue.sourcePath.empty() ? L"-" : issue.sourcePath;
            row.destinationPathText = issue.destinationPath.empty() ? L"-" : issue.destinationPath;
            rows.push_back(std::move(row));
        }
    }

    return rows;
}

void FileOperationsIssuesPaneState::RebuildList() noexcept
{
    if (! _list)
    {
        return;
    }

    SendMessageW(_list.get(), WM_SETREDRAW, FALSE, 0);
    ListView_DeleteAllItems(_list.get());

    for (size_t i = 0; i < _rows.size(); ++i)
    {
        LVITEMW item{};
        item.mask          = LVIF_TEXT | LVIF_PARAM;
        item.iItem         = static_cast<int>(i);
        item.iSubItem      = 0;
        item.pszText       = const_cast<wchar_t*>(_rows[i].timeText.c_str());
        item.lParam        = static_cast<LPARAM>(i);
        const int rowIndex = ListView_InsertItem(_list.get(), &item);
        if (rowIndex < 0)
        {
            continue;
        }

        ListView_SetItemText(_list.get(), rowIndex, 1, const_cast<wchar_t*>(_rows[i].taskText.c_str()));
        ListView_SetItemText(_list.get(), rowIndex, 2, const_cast<wchar_t*>(_rows[i].operationText.c_str()));
        ListView_SetItemText(_list.get(), rowIndex, 3, const_cast<wchar_t*>(_rows[i].severityText.c_str()));
        ListView_SetItemText(_list.get(), rowIndex, 4, const_cast<wchar_t*>(_rows[i].statusText.c_str()));
        ListView_SetItemText(_list.get(), rowIndex, 5, const_cast<wchar_t*>(_rows[i].statusTextDetail.c_str()));
        ListView_SetItemText(_list.get(), rowIndex, 6, const_cast<wchar_t*>(_rows[i].categoryText.c_str()));
        ListView_SetItemText(_list.get(), rowIndex, 7, const_cast<wchar_t*>(_rows[i].messageText.c_str()));
        ListView_SetItemText(_list.get(), rowIndex, 8, const_cast<wchar_t*>(_rows[i].sourcePathText.c_str()));
        ListView_SetItemText(_list.get(), rowIndex, 9, const_cast<wchar_t*>(_rows[i].destinationPathText.c_str()));
    }

    SendMessageW(_list.get(), WM_SETREDRAW, TRUE, 0);
    InvalidateRect(_list.get(), nullptr, TRUE);
}

void FileOperationsIssuesPaneState::RefreshRows(bool force) noexcept
{
    std::vector<IssuesRow> rows = BuildRows();
    if (! force && RowsEqual(rows, _rows))
    {
        return;
    }

    _rows = std::move(rows);
    RebuildList();
}

LRESULT FileOperationsIssuesPaneState::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);

    HWND list = CreateWindowExW(0,
                                WC_LISTVIEWW,
                                nullptr,
                                WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SINGLESEL,
                                0,
                                0,
                                1,
                                1,
                                hwnd,
                                reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIssuesListControlId)),
                                GetModuleHandleW(nullptr),
                                nullptr);
    if (! list)
    {
        return -1;
    }

    _list.reset(list);
    ListView_SetExtendedListViewStyle(_list.get(), LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP | LVS_EX_INFOTIP);

    EnsureColumns();
    ApplyColumnLayout();
    ApplyTheme(hwnd);

    RECT rc{};
    GetClientRect(hwnd, &rc);
    OnSize(hwnd, static_cast<UINT>(std::max(0L, rc.right - rc.left)), static_cast<UINT>(std::max(0L, rc.bottom - rc.top)));

    RefreshRows(true);
    SetTimer(hwnd, kRefreshTimerId, kRefreshTimerIntervalMs, nullptr);
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnEraseBkgnd() noexcept
{
    return 1;
}

LRESULT FileOperationsIssuesPaneState::OnPaint(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return 0;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
    if (! hdc)
    {
        return 0;
    }

    HBRUSH brush = _backgroundBrush ? _backgroundBrush.get() : static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH));
    FillRect(hdc.get(), &ps.rcPaint, brush);
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnSize(HWND hwnd, UINT width, UINT height) noexcept
{
    if (_list)
    {
        const int padding = DipsToPixels(6, _dpi);
        const int x       = padding;
        const int y       = padding;
        const int w       = std::max(1, static_cast<int>(width) - (padding * 2));
        const int h       = std::max(1, static_cast<int>(height) - (padding * 2));
        MoveWindow(_list.get(), x, y, w, h, TRUE);
    }

    if (hwnd && fileOps)
    {
        fileOps->SaveIssuesPanePlacement(hwnd);
    }

    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnNotify(NMHDR* header) noexcept
{
    if (! header || ! _list || header->hwndFrom != _list.get() || header->idFrom != kIssuesListControlId || header->code != NM_CUSTOMDRAW)
    {
        return 0;
    }

    if (_theme.highContrast)
    {
        return CDRF_DODEFAULT;
    }

    auto* cd = reinterpret_cast<NMLVCUSTOMDRAW*>(header);
    if (! cd)
    {
        return CDRF_DODEFAULT;
    }

    if (cd->nmcd.dwDrawStage == CDDS_PREPAINT)
    {
        return CDRF_NOTIFYITEMDRAW;
    }

    if (cd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)
    {
        const size_t index = static_cast<size_t>(cd->nmcd.dwItemSpec);
        if (index >= _rows.size())
        {
            return CDRF_DODEFAULT;
        }

        const bool selected = (cd->nmcd.uItemState & CDIS_SELECTED) != 0;
        if (selected)
        {
            cd->clrText   = _theme.menu.selectionText;
            cd->clrTextBk = _theme.menu.selectionBg;
            return CDRF_NOTIFYPOSTPAINT;
        }

        if (_theme.menu.rainbowMode)
        {
            const COLORREF bg = ColorToCOLORREF(RainbowRowColor(_theme, _rows[index]));
            cd->clrTextBk     = bg;
            cd->clrText       = ChooseContrastingTextColor(bg);
            return CDRF_NOTIFYPOSTPAINT;
        }

        if (_rows[index].severity == FolderWindow::FileOperationState::DiagnosticSeverity::Error)
        {
            cd->clrTextBk = ColorToCOLORREF(_theme.folderView.errorBackground);
            cd->clrText   = ColorToCOLORREF(_theme.folderView.errorText);
            return CDRF_NOTIFYPOSTPAINT;
        }

        cd->clrTextBk = ColorToCOLORREF(_theme.folderView.warningBackground);
        cd->clrText   = ColorToCOLORREF(_theme.folderView.warningText);
        return CDRF_NOTIFYPOSTPAINT;
    }

    if (cd->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)
    {
        const size_t index = static_cast<size_t>(cd->nmcd.dwItemSpec);
        if (index >= _rows.size())
        {
            return CDRF_DODEFAULT;
        }

        const bool selected = (cd->nmcd.uItemState & CDIS_SELECTED) != 0;

        COLORREF rowBg = _theme.windowBackground;
        if (selected)
        {
            rowBg = _theme.menu.selectionBg;
        }
        else
        {
            if (_theme.menu.rainbowMode)
            {
                rowBg = ColorToCOLORREF(RainbowRowColor(_theme, _rows[index]));
            }
            else if (_rows[index].severity == FolderWindow::FileOperationState::DiagnosticSeverity::Error)
            {
                rowBg = ColorToCOLORREF(_theme.folderView.errorBackground);
            }
            else
            {
                rowBg = ColorToCOLORREF(_theme.folderView.warningBackground);
            }
        }

        const COLORREF lineColor = ThemedControls::BlendColor(rowBg, _theme.menu.separator, 1, 6);

        const RECT& rc = cd->nmcd.rc;
        const int y    = rc.bottom - 1;

        wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> pen(CreatePen(PS_SOLID, 1, lineColor));
        if (pen)
        {
            auto oldPen = wil::SelectObject(cd->nmcd.hdc, pen.get());
            MoveToEx(cd->nmcd.hdc, rc.left, y, nullptr);
            LineTo(cd->nmcd.hdc, rc.right, y);
        }

        return CDRF_DODEFAULT;
    }

    return CDRF_DODEFAULT;
}

LRESULT FileOperationsIssuesPaneState::OnTimer(HWND /*hwnd*/, UINT_PTR timerId) noexcept
{
    if (timerId != kRefreshTimerId)
    {
        return 0;
    }

    RefreshRows(false);
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnMove(HWND hwnd) noexcept
{
    if (fileOps)
    {
        fileOps->SaveIssuesPanePlacement(hwnd);
    }
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnExitSizeMove(HWND hwnd) noexcept
{
    if (fileOps)
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

    if (fileOps)
    {
        fileOps->SaveIssuesPanePlacement(hwnd);
    }

    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnClose(HWND hwnd) noexcept
{
    if (fileOps)
    {
        fileOps->SaveIssuesPanePlacement(hwnd);
    }

    ShowWindow(hwnd, SW_HIDE);
    return 0;
}

LRESULT FileOperationsIssuesPaneState::OnDpiChanged(HWND hwnd, UINT dpi, const RECT* suggested) noexcept
{
    _dpi = dpi == 0 ? USER_DEFAULT_SCREEN_DPI : dpi;

    if (suggested)
    {
        const int width  = static_cast<int>(std::max(0L, suggested->right - suggested->left));
        const int height = static_cast<int>(std::max(0L, suggested->bottom - suggested->top));
        SetWindowPos(hwnd, nullptr, suggested->left, suggested->top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    ApplyColumnLayout();
    ApplyTheme(hwnd);

    if (fileOps)
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

    if (fileOps)
    {
        fileOps->OnIssuesPaneDestroyed(hwnd);
    }

    _rows.clear();
    _list.reset();

    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    delete this;
    return 0;
}

LRESULT FileOperationsIssuesPaneState::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_CREATE: return OnCreate(hwnd);
        case WM_ERASEBKGND: return OnEraseBkgnd();
        case WM_PAINT: return OnPaint(hwnd);
        case WM_SIZE: return OnSize(hwnd, LOWORD(lp), HIWORD(lp));
        case WM_NOTIFY: return OnNotify(reinterpret_cast<NMHDR*>(lp));
        case WM_TIMER: return OnTimer(hwnd, wp);
        case WM_MOVE: return OnMove(hwnd);
        case WM_GETMINMAXINFO:
        {
            auto* info = reinterpret_cast<MINMAXINFO*>(lp);
            if (info)
            {
                static_cast<void>(WindowMaximizeBehavior::ApplyVerticalMaximize(hwnd, *info));
            }
            return 0;
        }
        case WM_EXITSIZEMOVE: return OnExitSizeMove(hwnd);
        case WM_SHOWWINDOW: return OnShowWindow(hwnd, wp != FALSE);
        case WM_DPICHANGED: return OnDpiChanged(hwnd, static_cast<UINT>(HIWORD(wp)), reinterpret_cast<const RECT*>(lp));
        case WM_THEMECHANGED: return OnThemeChanged(hwnd);
        case WM_SETTINGCHANGE:
        case WM_SYSCOLORCHANGE: return OnThemeChanged(hwnd);
        case WM_NCACTIVATE:
            if (folderWindow)
            {
                if (! _inTitleBarThemeApply)
                {
                    _inTitleBarThemeApply = true;
                    ApplyTitleBarTheme(hwnd, folderWindow->GetTheme(), wp != FALSE);
                    _inTitleBarThemeApply = false;
                }
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

    if (state)
    {
        return state->WndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

HWND FileOperationsIssuesPane::Create(FolderWindow::FileOperationState* fileOps, FolderWindow* folderWindow, HWND ownerWindow) noexcept
{
    if (! fileOps || ! folderWindow)
    {
        return nullptr;
    }

    if (! RegisterFileOperationsIssuesPaneClass(GetModuleHandleW(nullptr)))
    {
        return nullptr;
    }

    auto statePtr          = std::make_unique<FileOperationsIssuesPaneState>();
    statePtr->fileOps      = fileOps;
    statePtr->folderWindow = folderWindow;

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

        const RECT work      = mi.rcWork;
        const int workLeft   = static_cast<int>(work.left);
        const int workTop    = static_cast<int>(work.top);
        const int workRight  = static_cast<int>(work.right);
        const int workBottom = static_cast<int>(work.bottom);

        const int maxX    = workRight - width;
        const int maxY    = workBottom - height;
        const int centerX = workLeft + (workRight - workLeft - width) / 2;
        const int centerY = workTop + (workBottom - workTop - height) / 2;

        const int x = maxX >= workLeft ? std::clamp(centerX, workLeft, maxX) : workLeft;
        const int y = maxY >= workTop ? std::clamp(centerY, workTop, maxY) : workTop;

        windowRect.left   = x;
        windowRect.top    = y;
        windowRect.right  = x + width;
        windowRect.bottom = y + height;
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_PANE_TITLE);

    auto* state = statePtr.release();
    HWND pane   = CreateWindowExW(WS_EX_APPWINDOW,
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
        std::unique_ptr<FileOperationsIssuesPaneState> reclaimed(state);
        return nullptr;
    }

    ShowWindow(pane, startMaximized ? SW_SHOWMAXIMIZED : SW_SHOWNORMAL);
    UpdateWindow(pane);
    return pane;
}
