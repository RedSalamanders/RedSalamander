#include "FolderWindowInternal.h"

#include "FluentIcons.h"

namespace
{
constexpr int kStatusBarFocusLineHeightDip = 2;

wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> g_statusBarIconFont;
UINT g_statusBarIconFontDpi   = USER_DEFAULT_SCREEN_DPI;
bool g_statusBarIconFontValid = false;

[[nodiscard]] bool EnsureStatusBarIconFont(UINT dpi, HWND hwnd) noexcept
{
    if (dpi == 0)
    {
        dpi = USER_DEFAULT_SCREEN_DPI;
    }

    if (dpi != g_statusBarIconFontDpi || ! g_statusBarIconFont)
    {
        g_statusBarIconFont      = FluentIcons::CreateFontForDpi(dpi, FluentIcons::kDefaultSizeDip);
        g_statusBarIconFontDpi   = dpi;
        g_statusBarIconFontValid = false;

        if (g_statusBarIconFont)
        {
            auto hdc = wil::GetDC(hwnd);
            if (hdc)
            {
                g_statusBarIconFontValid = FluentIcons::FontHasGlyph(hdc.get(), g_statusBarIconFont.get(), FluentIcons::kSort);
            }
        }
    }

    return g_statusBarIconFontValid;
}

[[nodiscard]] bool ContainsPrivateUseAreaGlyph(std::wstring_view text) noexcept
{
    for (wchar_t ch : text)
    {
        if (ch >= 0xE000 && ch <= 0xF8FF)
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] COLORREF BlendColor(COLORREF base, COLORREF overlay, int overlayWeight, int denom) noexcept
{
    if (denom <= 0)
    {
        return base;
    }
    overlayWeight        = std::clamp(overlayWeight, 0, denom);
    const int baseWeight = denom - overlayWeight;

    const int r = (static_cast<int>(GetRValue(base)) * baseWeight + static_cast<int>(GetRValue(overlay)) * overlayWeight) / denom;
    const int g = (static_cast<int>(GetGValue(base)) * baseWeight + static_cast<int>(GetGValue(overlay)) * overlayWeight) / denom;
    const int b = (static_cast<int>(GetBValue(base)) * baseWeight + static_cast<int>(GetBValue(overlay)) * overlayWeight) / denom;
    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}

void PaintSortIndicatorGlyph(HDC hdc, const RECT& rc, HFONT iconFont, HFONT arrowFont, COLORREF color, const std::wstring& sortText) noexcept
{
    if (! hdc || ! iconFont || sortText.empty())
    {
        return;
    }

    const wchar_t icon    = sortText.back();
    const bool hasArrow   = sortText.size() >= 2;
    const wchar_t arrowCh = hasArrow ? sortText.front() : L'\0';

    RECT box         = rc;
    const int width  = std::max(0L, box.right - box.left);
    const int height = std::max(0L, box.bottom - box.top);
    const int size   = std::max(0, std::min(width, height));
    if (size <= 0)
    {
        return;
    }

    box.left = std::max(box.left, box.right - size);

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, color);

    {
        auto oldFont = wil::SelectObject(hdc, iconFont);
        DrawTextW(hdc, &icon, 1, &box, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }

    if (hasArrow && arrowCh != 0 && arrowFont)
    {
        RECT arrowRect  = box;
        const int inset = std::max(1, size / 3);
        arrowRect.left  = std::min(arrowRect.right, arrowRect.left + inset);
        arrowRect.top   = std::min(arrowRect.bottom, arrowRect.top + inset);

        auto oldFont = wil::SelectObject(hdc, arrowFont);
        DrawTextW(hdc, &arrowCh, 1, &arrowRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }
}

bool IsPointInStatusBarPart(HWND statusBar, int part, POINT clientPt) noexcept
{
    RECT rc{};
    if (! statusBar)
    {
        return false;
    }

    const auto result = SendMessageW(statusBar, SB_GETRECT, static_cast<WPARAM>(part), reinterpret_cast<LPARAM>(&rc));
    if (result == 0)
    {
        return false;
    }
    return PtInRect(&rc, clientPt) != 0;
}

[[nodiscard]] bool IsStatusBarActivePane(HWND statusBar, const FolderWindow& owner) noexcept
{
    const int id = GetDlgCtrlID(statusBar);
    if (id == static_cast<int>(kLeftStatusBarId))
    {
        return owner.GetActivePane() == FolderWindow::Pane::Left;
    }
    if (id == static_cast<int>(kRightStatusBarId))
    {
        return owner.GetActivePane() == FolderWindow::Pane::Right;
    }
    return false;
}

[[nodiscard]] int GetStatusBarFocusLineHeightPx(int dpi, const RECT& clientRect) noexcept
{
    const int clientHeight = std::max(0l, clientRect.bottom - clientRect.top);
    if (clientHeight <= 0)
    {
        return 0;
    }

    const int desired = MulDiv(kStatusBarFocusLineHeightDip, dpi, USER_DEFAULT_SCREEN_DPI);
    return std::clamp(desired, 1, clientHeight);
}

[[nodiscard]] COLORREF StatusBarFocusLineColor(const AppTheme& theme, bool activePane, const uint32_t* hueDegrees) noexcept
{
    if (! activePane)
    {
        return theme.menu.separator;
    }

    if (! theme.menu.rainbowMode)
    {
        return theme.menu.selectionBg;
    }

    const uint32_t hueDegreesValue = hueDegrees ? *hueDegrees : 0u;
    const float hue                = static_cast<float>(hueDegreesValue % 360u);
    const float saturation         = 0.85f;
    const float value              = theme.menu.darkBase ? 0.80f : 0.90f;
    return ColorToCOLORREF(ColorFromHSV(hue, saturation, value, 1.0f));
}

void PaintStatusBar(HWND hwnd, HDC hdc) noexcept
{
    if (! hwnd || ! hdc)
    {
        return;
    }

    auto* owner = reinterpret_cast<FolderWindow*>(GetPropW(hwnd, kStatusBarOwnerProp));
    if (! owner)
    {
        return;
    }

    const auto* selectionText = reinterpret_cast<const std::wstring*>(GetPropW(hwnd, kStatusBarSelectionTextProp));
    const auto* sortText      = reinterpret_cast<const std::wstring*>(GetPropW(hwnd, kStatusBarSortTextProp));
    if (! selectionText || ! sortText)
    {
        return;
    }

    RECT client{};
    if (! GetClientRect(hwnd, &client))
    {
        return;
    }

    const AppTheme& theme       = owner->GetTheme();
    const bool activePane       = IsStatusBarActivePane(hwnd, *owner);
    const auto* focusHueDegrees = reinterpret_cast<const uint32_t*>(GetPropW(hwnd, kStatusBarFocusHueProp));

    RECT part0{};
    RECT part1{};
    const bool has0 = SendMessageW(hwnd, SB_GETRECT, 0, reinterpret_cast<LPARAM>(&part0)) != 0;
    const bool has1 = SendMessageW(hwnd, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&part1)) != 0;

    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> bgBrush(CreateSolidBrush(theme.menu.background));
    FillRect(hdc, &client, bgBrush.get());

    const bool hot = GetPropW(hwnd, kStatusBarSortHotProp) != nullptr;
    if (hot && has1)
    {
        const COLORREF hotBg = BlendColor(theme.menu.background, theme.menu.selectionBg, 1, 2);
        wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> hotBrush(CreateSolidBrush(hotBg));
        FillRect(hdc, &part1, hotBrush.get());

        RECT frame              = part1;
        const int dpi           = GetDeviceCaps(hdc, LOGPIXELSX);
        const int focusLineSize = GetStatusBarFocusLineHeightPx(dpi, client);
        frame.top               = std::min(frame.bottom, frame.top + focusLineSize);
        frame.left              = std::min(frame.right, frame.left + 1);

        wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> frameBrush(CreateSolidBrush(theme.menu.separator));
        FrameRect(hdc, &frame, frameBrush.get());
    }

    const int dpi         = GetDeviceCaps(hdc, LOGPIXELSX);
    const int focusLinePx = GetStatusBarFocusLineHeightPx(dpi, client);
    const RECT topLine    = {client.left, client.top, client.right, std::min(client.bottom, client.top + focusLinePx)};
    if (topLine.bottom > topLine.top)
    {
        const COLORREF lineColor = StatusBarFocusLineColor(theme, activePane, focusHueDegrees);
        wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> lineBrush(CreateSolidBrush(lineColor));
        FillRect(hdc, &topLine, lineBrush.get());
    }

    if (has0)
    {
        RECT sepRect  = part0;
        sepRect.left  = std::max(part0.left, part0.right - 1);
        sepRect.right = part0.right;
        sepRect.top   = std::min(part0.bottom, part0.top + focusLinePx);
        if (sepRect.right > sepRect.left && sepRect.bottom > sepRect.top)
        {
            wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> sepBrush(CreateSolidBrush(theme.menu.separator));
            FillRect(hdc, &sepRect, sepBrush.get());
        }
    }

    SetBkMode(hdc, TRANSPARENT);

    const bool iconFontValid = EnsureStatusBarIconFont(static_cast<UINT>(dpi), hwnd);

    HFONT textFont = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    if (! textFont)
    {
        textFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    const int paddingX     = std::max(1, MulDiv(kStatusBarPaddingXDip, dpi, USER_DEFAULT_SCREEN_DPI));
    const int sortPaddingX = std::max(1, MulDiv(kStatusBarSortPaddingXDip, dpi, USER_DEFAULT_SCREEN_DPI));

    RECT rc0  = has0 ? part0 : client;
    rc0.left  = std::min(rc0.right, rc0.left + paddingX);
    rc0.right = std::max(rc0.left, rc0.right - paddingX);
    rc0.top   = std::min(rc0.bottom, rc0.top + focusLinePx);

    RECT rc1  = has1 ? part1 : client;
    rc1.left  = std::min(rc1.right, rc1.left + sortPaddingX);
    rc1.right = std::max(rc1.left, rc1.right - sortPaddingX);
    rc1.top   = std::min(rc1.bottom, rc1.top + focusLinePx);

    {
        auto oldFont = wil::SelectObject(hdc, textFont);
        SetTextColor(hdc, theme.menu.text);
        DrawTextW(hdc, selectionText->c_str(), static_cast<int>(selectionText->size()), &rc0, DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_LEFT);
    }

    const bool sortUsesIcons = iconFontValid && g_statusBarIconFont && ContainsPrivateUseAreaGlyph(*sortText);
    const COLORREF sortColor = hot ? theme.menu.selectionText : theme.menu.text;
    if (sortUsesIcons && g_statusBarIconFont)
    {
        PaintSortIndicatorGlyph(hdc, rc1, g_statusBarIconFont.get(), textFont, sortColor, *sortText);
    }
    else
    {
        auto oldFont = wil::SelectObject(hdc, textFont);
        SetTextColor(hdc, sortColor);
        DrawTextW(hdc, sortText->c_str(), static_cast<int>(sortText->size()), &rc1, DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_RIGHT);
    }
}

bool StatusBarCanCustomPaint(HWND hwnd) noexcept
{
    return GetPropW(hwnd, kStatusBarOwnerProp) != nullptr && GetPropW(hwnd, kStatusBarSelectionTextProp) != nullptr &&
           GetPropW(hwnd, kStatusBarSortTextProp) != nullptr;
}

void UpdateStatusBarSortHot(HWND hwnd, POINT pt) noexcept
{
    const bool hotNow = IsPointInStatusBarPart(hwnd, 1, pt);
    const bool hotWas = GetPropW(hwnd, kStatusBarSortHotProp) != nullptr;
    if (hotNow == hotWas)
    {
        return;
    }

    if (hotNow)
    {
        SetPropW(hwnd, kStatusBarSortHotProp, reinterpret_cast<HANDLE>(1));
    }
    else
    {
        RemovePropW(hwnd, kStatusBarSortHotProp);
    }
    InvalidateRect(hwnd, nullptr, FALSE);
}

void TrackStatusBarMouseLeave(HWND hwnd) noexcept
{
    TRACKMOUSEEVENT tme{};
    tme.cbSize    = sizeof(tme);
    tme.dwFlags   = TME_LEAVE;
    tme.hwndTrack = hwnd;
    TrackMouseEvent(&tme);
}

LRESULT StatusBarOnEraseBkgnd(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    if (StatusBarCanCustomPaint(hwnd))
    {
        return 1;
    }
    return DefSubclassProc(hwnd, WM_ERASEBKGND, wParam, lParam);
}

LRESULT StatusBarOnPaint(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    if (! StatusBarCanCustomPaint(hwnd))
    {
        return DefSubclassProc(hwnd, WM_PAINT, wParam, lParam);
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint paintDc = wil::BeginPaint(hwnd, &ps);
    PaintStatusBar(hwnd, paintDc.get());
    return 0;
}

LRESULT StatusBarOnSetCursor(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    POINT screenPt{};
    if (GetCursorPos(&screenPt))
    {
        ScreenToClient(hwnd, &screenPt);
        if (IsPointInStatusBarPart(hwnd, 1, screenPt))
        {
            SetCursor(LoadCursorW(nullptr, IDC_HAND));
            return TRUE;
        }
    }
    return DefSubclassProc(hwnd, WM_SETCURSOR, wParam, lParam);
}

LRESULT StatusBarOnMouseMove(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    UpdateStatusBarSortHot(hwnd, pt);
    TrackStatusBarMouseLeave(hwnd);
    return DefSubclassProc(hwnd, WM_MOUSEMOVE, wParam, lParam);
}

LRESULT StatusBarOnMouseLeave(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    if (RemovePropW(hwnd, kStatusBarSortHotProp))
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return DefSubclassProc(hwnd, WM_MOUSELEAVE, wParam, lParam);
}

LRESULT StatusBarOnSize(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    const LRESULT result = DefSubclassProc(hwnd, WM_SIZE, wParam, lParam);

    if (StatusBarCanCustomPaint(hwnd))
    {
        RECT client{};
        if (GetClientRect(hwnd, &client) != 0)
        {
            const int width            = std::max(0L, client.right - client.left);
            const int dpi              = static_cast<int>(GetDpiForWindow(hwnd));
            const int minSortPartWidth = MulDiv(kStatusBarSortMinPartWidthDip, dpi, USER_DEFAULT_SCREEN_DPI);
            const int clampedSortWidth = std::clamp(minSortPartWidth, 0, width);
            const int parts[]          = {std::max(0, width - clampedSortWidth), -1};
            SendMessageW(hwnd, SB_SETPARTS, 2, reinterpret_cast<LPARAM>(parts));
        }
    }

    InvalidateRect(hwnd, nullptr, FALSE);
    return result;
}

LRESULT StatusBarOnNcDestroy(HWND hwnd, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass) noexcept
{
    RemovePropW(hwnd, kStatusBarSortHotProp);
    RemovePropW(hwnd, kStatusBarOwnerProp);
    RemovePropW(hwnd, kStatusBarSelectionTextProp);
    RemovePropW(hwnd, kStatusBarSortTextProp);
    RemovePropW(hwnd, kStatusBarFocusHueProp);
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
    RemoveWindowSubclass(hwnd, StatusBarSubclassProc, uIdSubclass);
#pragma warning(pop)
    return DefSubclassProc(hwnd, WM_NCDESTROY, wParam, lParam);
}
} // namespace

LRESULT CALLBACK StatusBarSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR /*dwRefData*/)
{
    switch (msg)
    {
        case WM_ERASEBKGND: return StatusBarOnEraseBkgnd(hwnd, wParam, lParam);
        case WM_PAINT: return StatusBarOnPaint(hwnd, wParam, lParam);
        case WM_SETCURSOR: return StatusBarOnSetCursor(hwnd, wParam, lParam);
        case WM_MOUSEMOVE: return StatusBarOnMouseMove(hwnd, wParam, lParam);
        case WM_MOUSELEAVE: return StatusBarOnMouseLeave(hwnd, wParam, lParam);
        case WM_SIZE: return StatusBarOnSize(hwnd, wParam, lParam);
        case WM_NCDESTROY: return StatusBarOnNcDestroy(hwnd, wParam, lParam, uIdSubclass);
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

namespace
{

std::wstring FormatLocalTime(int64_t fileTime)
{
    if (fileTime <= 0)
    {
        return {};
    }

    ULARGE_INTEGER uli{};
    uli.QuadPart = static_cast<ULONGLONG>(fileTime);

    FILETIME ft{};
    ft.dwLowDateTime  = uli.LowPart;
    ft.dwHighDateTime = uli.HighPart;

    FILETIME local{};
    SYSTEMTIME st{};
    if (! FileTimeToLocalFileTime(&ft, &local) || ! FileTimeToSystemTime(&local, &st))
    {
        return {};
    }

    return std::format(L"{:04d}-{:02d}-{:02d} {:02d}:{:02d}", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute);
}

std::wstring FormatFileAttributes(DWORD attrs)
{
    std::wstring result;
    result.reserve(10);

    auto add = [&](DWORD flag, wchar_t ch)
    {
        if ((attrs & flag) != 0)
        {
            result.push_back(ch);
        }
    };

    add(FILE_ATTRIBUTE_READONLY, L'R');
    add(FILE_ATTRIBUTE_HIDDEN, L'H');
    add(FILE_ATTRIBUTE_SYSTEM, L'S');
    add(FILE_ATTRIBUTE_ARCHIVE, L'A');
    add(FILE_ATTRIBUTE_COMPRESSED, L'C');
    add(FILE_ATTRIBUTE_ENCRYPTED, L'E');
    add(FILE_ATTRIBUTE_TEMPORARY, L'T');
    add(FILE_ATTRIBUTE_OFFLINE, L'O');
    add(FILE_ATTRIBUTE_REPARSE_POINT, L'P');

    if (result.empty())
    {
        result = L"-";
    }

    return result;
}

std::wstring BuildSelectionSummaryText(const FolderView::SelectionStats& stats, std::wstring_view selectionSizeText)
{
    if (stats.selectedFiles == 0 && stats.selectedFolders == 0)
    {
        return LoadStringResource(nullptr, IDS_STATUS_NO_SELECTION);
    }

    if (stats.singleItem.has_value())
    {
        const FolderView::SelectionStats::SelectedItemDetails details = stats.singleItem.value();
        const std::wstring timeText                                   = FormatLocalTime(details.lastWriteTime);
        const std::wstring attrsText                                  = FormatFileAttributes(details.fileAttributes);
        if (details.isDirectory)
        {
            if (! timeText.empty())
            {
                return FormatStringResource(nullptr, IDS_FMT_STATUS_SELECTED_SINGLE_DIR_TIME_ATTRS, selectionSizeText, timeText, attrsText);
            }
            return FormatStringResource(nullptr, IDS_FMT_STATUS_SELECTED_SINGLE_DIR_ATTRS, selectionSizeText, attrsText);
        }

        const std::wstring sizeText = FormatBytesCompact(details.sizeBytes);
        if (! timeText.empty())
        {
            return FormatStringResource(nullptr, IDS_FMT_STATUS_SELECTED_SINGLE_FILE_SIZE_TIME_ATTRS, sizeText, timeText, attrsText);
        }
        return FormatStringResource(nullptr, IDS_FMT_STATUS_SELECTED_SINGLE_FILE_SIZE_ATTRS, sizeText, attrsText);
    }

    const std::wstring_view folderSuffix = stats.selectedFolders == 1 ? std::wstring_view{} : std::wstring_view{L"s"};
    const std::wstring_view fileSuffix   = stats.selectedFiles == 1 ? std::wstring_view{} : std::wstring_view{L"s"};

    if (stats.selectedFiles > 0 && stats.selectedFolders > 0)
    {
        return FormatStringResource(
            nullptr, IDS_FMT_STATUS_SELECTED_FOLDERS_FILES, stats.selectedFolders, folderSuffix, stats.selectedFiles, fileSuffix, selectionSizeText);
    }

    if (stats.selectedFiles > 0)
    {
        return FormatStringResource(nullptr, IDS_FMT_STATUS_SELECTED_FILES, stats.selectedFiles, fileSuffix, selectionSizeText);
    }

    return FormatStringResource(nullptr, IDS_FMT_STATUS_SELECTED_FOLDERS, stats.selectedFolders, folderSuffix, selectionSizeText);
}

std::wstring BuildSortIndicatorText(FolderView::SortBy sortBy, FolderView::SortDirection direction, bool useFluentIcons)
{
    if (sortBy == FolderView::SortBy::None)
    {
        if (useFluentIcons)
        {
            return std::wstring(1, FluentIcons::kSort);
        }

        std::wstring placeholder = LoadStringResource(nullptr, IDS_STATUS_SORT_INDICATOR);
        if (placeholder.empty())
        {
            placeholder.assign(1, FluentIcons::kFallbackSort);
        }
        return placeholder;
    }

    // Asc/Desc should use arrows (not chevrons) and we overlay it over the sort-by glyph in the status bar paint.
    const wchar_t arrow = direction == FolderView::SortDirection::Ascending ? L'\u2191' : L'\u2193';

    const wchar_t icon = [&]() noexcept -> wchar_t
    {
        if (useFluentIcons)
        {
            switch (sortBy)
            {
                case FolderView::SortBy::Name: return FluentIcons::kFont;
                case FolderView::SortBy::Extension: return FluentIcons::kDocument;
                case FolderView::SortBy::Time: return FluentIcons::kCalendar;
                case FolderView::SortBy::Size: return FluentIcons::kHardDrive;
                case FolderView::SortBy::Attributes: return FluentIcons::kTag;
                case FolderView::SortBy::None: break;
            }
        }
        else
        {
            switch (sortBy)
            {
                case FolderView::SortBy::Name: return L'\u2263';
                case FolderView::SortBy::Extension: return L'\u24D4';
                case FolderView::SortBy::Time: return L'\u23F1';
                case FolderView::SortBy::Size: return direction == FolderView::SortDirection::Ascending ? L'\u25F0' : L'\u25F2';
                case FolderView::SortBy::Attributes: return L'\u24B6';
                case FolderView::SortBy::None: break;
            }
        }

        return 0;
    }();

    std::wstring result;
    result.reserve(icon != 0 ? 2u : 1u);
    result.push_back(arrow);
    if (icon != 0)
    {
        result.push_back(icon);
    }
    return result;
}
} // namespace

void FolderWindow::UpdatePaneStatusBar(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.hStatusBar)
    {
        return;
    }

    const RECT& rect = pane == Pane::Left ? _leftStatusBarRect : _rightStatusBarRect;
    const int width  = std::max(0L, rect.right - rect.left);
    const int height = std::max(0L, rect.bottom - rect.top);

    const bool visible = state.statusBarVisible && width > 0 && height > 0;
    ShowWindow(state.hStatusBar.get(), visible ? SW_SHOWNA : SW_HIDE);
    if (! visible)
    {
        return;
    }

    std::wstring selectionSizeText;
    if (state.selectionStats.selectedFiles > 0 || state.selectionStats.selectedFolders > 0)
    {
        if (state.selectionStats.selectedFolders > 0)
        {
            if (state.selectionFolderBytesPending)
            {
                uint64_t totalBytesSoFar = state.selectionStats.selectedFileBytes;
                totalBytesSoFar += state.selectionFolderBytes;
                const std::wstring sizeText = FormatBytesCompact(totalBytesSoFar);
                selectionSizeText           = FormatStringResource(nullptr, IDS_FMT_STATUS_CALCULATING_SIZE_WITH_BYTES, sizeText);
                if (selectionSizeText.empty())
                {
                    selectionSizeText = LoadStringResource(nullptr, IDS_STATUS_CALCULATING_SIZE);
                }
            }
            else if (! state.selectionFolderBytesValid)
            {
                selectionSizeText = LoadStringResource(nullptr, IDS_STATUS_SIZE_UNKNOWN);
            }
            else
            {
                uint64_t totalBytes = state.selectionStats.selectedFileBytes;
                totalBytes += state.selectionFolderBytes;
                selectionSizeText = FormatBytesCompact(totalBytes);
            }
        }
        else
        {
            selectionSizeText = FormatBytesCompact(state.selectionStats.selectedFileBytes);
        }
    }

    if (state.folderView.IsIncrementalSearchActive())
    {
        const std::wstring queryText = std::wstring{state.folderView.GetIncrementalSearchQuery()};
        state.statusSelectionText    = FormatStringResource(nullptr, IDS_FMT_STATUS_INCREMENTAL_SEARCH, queryText);
        if (state.statusSelectionText.empty())
        {
            state.statusSelectionText = queryText;
        }
    }
    else
    {
        state.statusSelectionText = BuildSelectionSummaryText(state.selectionStats, selectionSizeText);
    }
    const bool useFluentIcons = EnsureStatusBarIconFont(_dpi, state.hStatusBar.get());
    state.statusSortText      = BuildSortIndicatorText(state.folderView.GetSortBy(), state.folderView.GetSortDirection(), useFluentIcons);

    const int dpi              = static_cast<int>(_dpi);
    const int minSortPartWidth = MulDiv(kStatusBarSortMinPartWidthDip, dpi, USER_DEFAULT_SCREEN_DPI);

    const int sortPartWidth = minSortPartWidth;

    const int clampedSortWidth = std::clamp(sortPartWidth, 0, width);
    const int parts[]          = {std::max(0, width - clampedSortWidth), -1};
    SendMessageW(state.hStatusBar.get(), SB_SETPARTS, 2, reinterpret_cast<LPARAM>(parts));

    SendMessageW(state.hStatusBar.get(), SB_SETTEXTW, MAKEWPARAM(0, SBT_NOBORDERS), reinterpret_cast<LPARAM>(state.statusSelectionText.c_str()));
    SendMessageW(state.hStatusBar.get(), SB_SETTEXTW, MAKEWPARAM(1, SBT_NOBORDERS), reinterpret_cast<LPARAM>(state.statusSortText.c_str()));

    const std::wstring sortTip = LoadStringResource(nullptr, IDS_TIP_STATUS_SORT);
    if (! sortTip.empty())
    {
        SendMessageW(state.hStatusBar.get(), SB_SETTIPTEXTW, 1, reinterpret_cast<LPARAM>(sortTip.c_str()));
    }

    InvalidateRect(state.hStatusBar.get(), nullptr, FALSE);
}
