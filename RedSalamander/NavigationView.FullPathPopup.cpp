#include "NavigationViewInternal.h"

#include <windowsx.h>

#include <winerror.h>

#include "DirectoryInfoCache.h"
#include "Helpers.h"
#include "IconCache.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "resource.h"

ATOM NavigationView::RegisterFullPathPopupWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = FullPathPopupWndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kFullPathPopupClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

LRESULT CALLBACK NavigationView::FullPathPopupWndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    NavigationView* self = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self    = reinterpret_cast<NavigationView*>(cs->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    }
    else
    {
        self = reinterpret_cast<NavigationView*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (self)
    {
        return self->FullPathPopupWndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT NavigationView::OnFullPathPopupCreate([[maybe_unused]] HWND hwnd)
{
    _fullPathPopupTrackingMouse                   = false;
    _fullPathPopupEditMode                        = false;
    _fullPathPopupActiveSeparatorIndex            = -1;
    _fullPathPopupMenuOpenForSeparator            = -1;
    _fullPathPopupPendingSeparatorMenuSwitchIndex = -1;
    _fullPathPopupHoveredSegmentIndex             = -1;
    _fullPathPopupHoveredSeparatorIndex           = -1;
    _fullPathPopupScrollY                         = 0.0f;

    _fullPathPopupHoverTimer = 0;
    return 0;
}

LRESULT NavigationView::OnFullPathPopupNcDestroy(HWND hwnd)
{
    const bool restoreFolderViewFocus              = _restoreFolderViewFocusAfterFullPathPopupClose;
    _restoreFolderViewFocusAfterFullPathPopupClose = false;
    _fullPathPopupEditMode                         = false;

    if (_fullPathPopupHoverTimer != 0)
    {
        KillTimer(hwnd, HOVER_TIMER_ID);
        _fullPathPopupHoverTimer = 0;
    }

    DiscardFullPathPopupD2DResources();
    _fullPathPopupEdit.release();
    _fullPathPopup.release();
    _fullPathPopupSegments.clear();
    _fullPathPopupSeparators.clear();
    _fullPathPopupActiveSeparatorIndex            = -1;
    _fullPathPopupMenuOpenForSeparator            = -1;
    _fullPathPopupPendingSeparatorMenuSwitchIndex = -1;

    if (restoreFolderViewFocus && _hWnd && IsWindow(_hWnd.get()) != FALSE)
    {
        PostMessageW(_hWnd.get(), WndMsg::kNavigationViewRestoreFolderFocus, 0, 0);
    }

    return 0;
}

LRESULT NavigationView::OnFullPathPopupTimer([[maybe_unused]] HWND hwnd, UINT_PTR timerId)
{
    if (timerId != HOVER_TIMER_ID || _fullPathPopupEditMode)
    {
        return 0;
    }

    return 0;
}

LRESULT NavigationView::OnFullPathPopupSize(HWND hwnd, UINT width, UINT height)
{
    _fullPathPopupClientSize.cx = static_cast<LONG>(width);
    _fullPathPopupClientSize.cy = static_cast<LONG>(height);

    if (_fullPathPopupTarget)
    {
        _fullPathPopupTarget->Resize(D2D1::SizeU(static_cast<UINT32>(_fullPathPopupClientSize.cx), static_cast<UINT32>(_fullPathPopupClientSize.cy)));
    }

    BuildFullPathPopupLayout(static_cast<float>(_fullPathPopupClientSize.cx));

    if (_fullPathPopupEdit && _fullPathPopupEditMode)
    {
        UpdateFullPathPopupEditHostLayout();
    }

    InvalidateRect(hwnd, nullptr, FALSE);
    return 0;
}

LRESULT NavigationView::OnFullPathPopupMouseMove(HWND hwnd, POINT pt)
{
    if (! _fullPathPopupTrackingMouse)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = hwnd;
        TrackMouseEvent(&tme);
        _fullPathPopupTrackingMouse = true;
    }

    if (_fullPathPopupEditMode)
    {
        return 0;
    }

    const float x = static_cast<float>(pt.x);
    const float y = static_cast<float>(pt.y) + _fullPathPopupScrollY;

    int newHoveredSegment = -1;
    for (size_t i = 0; i < _fullPathPopupSegments.size(); ++i)
    {
        const auto& bounds = _fullPathPopupSegments[i].bounds;
        if (bounds.left <= x && x <= bounds.right && bounds.top <= y && y <= bounds.bottom)
        {
            newHoveredSegment = static_cast<int>(i);
            break;
        }
    }

    int newHoveredSeparator = -1;
    for (size_t i = 0; i < _fullPathPopupSeparators.size(); ++i)
    {
        const auto& bounds = _fullPathPopupSeparators[i].bounds;
        if (bounds.left <= x && x <= bounds.right && bounds.top <= y && y <= bounds.bottom)
        {
            newHoveredSeparator = static_cast<int>(i);
            break;
        }
    }

    if (newHoveredSegment != _fullPathPopupHoveredSegmentIndex || newHoveredSeparator != _fullPathPopupHoveredSeparatorIndex)
    {
        _fullPathPopupHoveredSegmentIndex   = newHoveredSegment;
        _fullPathPopupHoveredSeparatorIndex = newHoveredSeparator;
        InvalidateRect(hwnd, nullptr, FALSE);
    }

    if (_fullPathPopupMenuOpenForSeparator != -1 && _fullPathPopupPendingSeparatorMenuSwitchIndex == -1)
    {
        const int targetIndex = newHoveredSeparator;
        if (targetIndex != -1 && targetIndex != _fullPathPopupMenuOpenForSeparator)
        {
            const size_t separatorIndex = static_cast<size_t>(targetIndex);
            bool eligibleForSiblings    = false;

            if (separatorIndex < _fullPathPopupSeparators.size())
            {
                const auto& separator = _fullPathPopupSeparators[separatorIndex];
                if (separator.rightSegmentIndex < _fullPathPopupSegments.size())
                {
                    const auto& segment          = _fullPathPopupSegments[separator.rightSegmentIndex];
                    const auto normalizedSegment = NormalizeDirectoryPath(segment.fullPath);
                    eligibleForSiblings          = ! normalizedSegment.parent_path().empty();
                }
            }

            if (eligibleForSiblings)
            {
                _fullPathPopupPendingSeparatorMenuSwitchIndex = targetIndex;
                SendMessageW(hwnd, WM_CANCELMODE, 0, 0);
                PostMessageW(hwnd, WndMsg::kNavigationMenuShowSiblingsDropdown, static_cast<WPARAM>(separatorIndex), 0);
            }
        }
    }

    return 0;
}

LRESULT NavigationView::OnFullPathPopupMouseLeave(HWND hwnd)
{
    _fullPathPopupTrackingMouse = false;
    if (_fullPathPopupHoveredSegmentIndex != -1 || _fullPathPopupHoveredSeparatorIndex != -1)
    {
        _fullPathPopupHoveredSegmentIndex   = -1;
        _fullPathPopupHoveredSeparatorIndex = -1;
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return 0;
}

LRESULT NavigationView::OnFullPathPopupLButtonDown(HWND hwnd, POINT pt)
{
    if (_fullPathPopupEditMode)
    {
        return 0;
    }

    const float x = static_cast<float>(pt.x);
    const float y = static_cast<float>(pt.y) + _fullPathPopupScrollY;

    for (const auto& segment : _fullPathPopupSegments)
    {
        if (segment.bounds.left <= x && x <= segment.bounds.right && segment.bounds.top <= y && y <= segment.bounds.bottom)
        {
            RequestPathChange(segment.fullPath);
            CloseFullPathPopup();
            return 0;
        }
    }

    for (size_t i = 0; i < _fullPathPopupSeparators.size(); ++i)
    {
        const auto& separator = _fullPathPopupSeparators[i];
        if (separator.bounds.left <= x && x <= separator.bounds.right && separator.bounds.top <= y && y <= separator.bounds.bottom)
        {
            ShowFullPathPopupSiblingsDropdown(hwnd, i);
            return 0;
        }
    }

    return 0;
}

LRESULT NavigationView::OnFullPathPopupLButtonDblClk(HWND /*hwnd*/, POINT pt)
{
    if (_fullPathPopupEditMode)
    {
        return 0;
    }

    const float x = static_cast<float>(pt.x);
    const float y = static_cast<float>(pt.y) + _fullPathPopupScrollY;

    for (const auto& segment : _fullPathPopupSegments)
    {
        if (segment.bounds.left <= x && x <= segment.bounds.right && segment.bounds.top <= y && y <= segment.bounds.bottom)
        {
            return 0;
        }
    }

    for (const auto& separator : _fullPathPopupSeparators)
    {
        if (separator.bounds.left <= x && x <= separator.bounds.right && separator.bounds.top <= y && y <= separator.bounds.bottom)
        {
            return 0;
        }
    }

    EnterFullPathPopupEditMode();
    return 0;
}

LRESULT NavigationView::OnFullPathPopupActivate(WORD state, HWND activatingWindow)
{
    if (state == WA_INACTIVE)
    {
        if (_fullPathPopupEditMode)
        {
            return 0;
        }
        if (_fullPathPopup && activatingWindow && (activatingWindow == _fullPathPopup.get() || IsChild(_fullPathPopup.get(), activatingWindow) != FALSE))
        {
            return 0;
        }
        CloseFullPathPopup();
    }
    return 0;
}

LRESULT NavigationView::OnFullPathPopupKeyDown(WPARAM key)
{
    if (key == VK_ESCAPE)
    {
        _restoreFolderViewFocusAfterFullPathPopupClose = true;
        CloseFullPathPopup();
        return 0;
    }

    const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    if (key == VK_F4 || (ctrl && (key == 'L' || key == 'l')))
    {
        EnterFullPathPopupEditMode();
        return 0;
    }

    return 0;
}

LRESULT NavigationView::OnFullPathPopupSysKeyDown(HWND hwnd, WPARAM key, LPARAM lParam)
{
    if (key == 'D' || key == 'd')
    {
        EnterFullPathPopupEditMode();
        return 0;
    }
    return DefWindowProcW(hwnd, WM_SYSKEYDOWN, key, lParam);
}

LRESULT NavigationView::OnFullPathPopupSysChar(HWND hwnd, WPARAM key, LPARAM lParam)
{
    if (key == 'D' || key == 'd')
    {
        return 0;
    }
    return DefWindowProcW(hwnd, WM_SYSCHAR, key, lParam);
}

LRESULT NavigationView::OnFullPathPopupMouseWheel(HWND hwnd, int delta)
{
    if (_fullPathPopupEditMode)
    {
        return 0;
    }

    const float lineHeight = static_cast<float>(_sectionPathRect.bottom - _sectionPathRect.top);
    const float step       = lineHeight > 0.0f ? lineHeight : 24.0f;
    _fullPathPopupScrollY  = std::clamp(_fullPathPopupScrollY - (static_cast<float>(delta) / static_cast<float>(WHEEL_DELTA)) * step,
                                        0.0f,
                                        std::max(0.0f, _fullPathPopupContentHeight - static_cast<float>(_fullPathPopupClientSize.cy)));
    InvalidateRect(hwnd, nullptr, FALSE);
    return 0;
}

LRESULT NavigationView::OnFullPathPopupCtlColorEdit(HWND hwnd, HDC hdc, HWND hwndControl)
{
    if (_fullPathPopupEdit && hwndControl == _fullPathPopupEdit->GetTextInputHwnd())
    {
        SetTextColor(hdc, ColorToCOLORREF(_theme.text));
        SetBkColor(hdc, _theme.gdiBackground);
        return reinterpret_cast<LRESULT>(_backgroundBrush.get());
    }

    return DefWindowProcW(hwnd, WM_CTLCOLOREDIT, reinterpret_cast<WPARAM>(hdc), reinterpret_cast<LPARAM>(hwndControl));
}

LRESULT NavigationView::OnFullPathPopupCommand(HWND hwnd, UINT id, UINT codeNotify, HWND hwndCtl)
{
    return DefWindowProcW(hwnd, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(id), static_cast<WORD>(codeNotify)), reinterpret_cast<LPARAM>(hwndCtl));
}

LRESULT NavigationView::OnShowFullPathPopupSiblingsDropdown(HWND popupHwnd, size_t separatorIndex)
{
    _fullPathPopupPendingSeparatorMenuSwitchIndex = -1;
    ShowFullPathPopupSiblingsDropdown(popupHwnd, separatorIndex);
    return 0;
}

LRESULT NavigationView::FullPathPopupWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_CREATE: return OnFullPathPopupCreate(hwnd);
        case WM_NCDESTROY: return OnFullPathPopupNcDestroy(hwnd);
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: RenderFullPathPopup(); return 0;
        case WM_TIMER: return OnFullPathPopupTimer(hwnd, static_cast<UINT_PTR>(wp));
        case WM_SIZE: return OnFullPathPopupSize(hwnd, LOWORD(lp), HIWORD(lp));
        case WM_MOUSEMOVE: return OnFullPathPopupMouseMove(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_MOUSELEAVE: return OnFullPathPopupMouseLeave(hwnd);
        case WM_LBUTTONDOWN: return OnFullPathPopupLButtonDown(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_LBUTTONDBLCLK: return OnFullPathPopupLButtonDblClk(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_ACTIVATE: return OnFullPathPopupActivate(LOWORD(wp), reinterpret_cast<HWND>(lp));
        case WM_KEYDOWN: return OnFullPathPopupKeyDown(wp);
        case WM_SYSKEYDOWN: return OnFullPathPopupSysKeyDown(hwnd, wp, lp);
        case WM_SYSCHAR: return OnFullPathPopupSysChar(hwnd, wp, lp);
        case WM_MOUSEWHEEL: return OnFullPathPopupMouseWheel(hwnd, GET_WHEEL_DELTA_WPARAM(wp));
        case WM_CTLCOLOREDIT: return OnFullPathPopupCtlColorEdit(hwnd, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_COMMAND: return OnFullPathPopupCommand(hwnd, LOWORD(wp), HIWORD(wp), reinterpret_cast<HWND>(lp));
        case WM_EXITMENULOOP: OnFullPathPopupExitMenuLoop(hwnd, static_cast<BOOL>(wp)); return 0;
        case WndMsg::kNavigationMenuShowSiblingsDropdown: return OnShowFullPathPopupSiblingsDropdown(hwnd, static_cast<size_t>(wp));
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void NavigationView::ShowFullPathPopupSiblingsDropdown(HWND popupHwnd, size_t separatorIndex)
{
    if (! popupHwnd || _fullPathPopupEditMode)
    {
        return;
    }

    if (separatorIndex >= _fullPathPopupSeparators.size())
    {
        return;
    }

    const auto& separator = _fullPathPopupSeparators[separatorIndex];
    if (separator.rightSegmentIndex >= _fullPathPopupSegments.size())
    {
        return;
    }

    const auto& segment                               = _fullPathPopupSegments[separator.rightSegmentIndex];
    const std::filesystem::path normalizedSegmentPath = NormalizeDirectoryPath(segment.fullPath);
    std::filesystem::path parentPath                  = normalizedSegmentPath.parent_path();
    if (parentPath.empty())
    {
        return;
    }

    std::vector<std::filesystem::path> siblings;
    if (! TryGetSiblingFolders(parentPath, siblings) || siblings.empty())
    {
        return;
    }

    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(siblings.size());

    const std::filesystem::path normalizedCurrentPath = NormalizeDirectoryPath(segment.fullPath);
    const std::wstring currentPathText                = normalizedCurrentPath.wstring();
    int selectedIndex                                 = 0;

    for (size_t i = 0; i < siblings.size(); ++i)
    {
        const std::filesystem::path normalizedSiblingPath = NormalizeDirectoryPath(siblings[i]);
        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.kind      = RedSalamander::DxUi::MenuItemKind::Radio;
        item.text      = FilenameOrPath(normalizedSiblingPath);
        item.commandId = static_cast<int>(ID_SIBLING_BASE + i);
        item.checked   = wil::compare_string_ordinal(normalizedSiblingPath.wstring(), currentPathText, true) == wistd::weak_ordering::equivalent;
        if (item.checked)
        {
            selectedIndex = static_cast<int>(i);
        }
        items.push_back(std::move(item));
    }

    _fullPathPopupActiveSeparatorIndex            = static_cast<int>(separatorIndex);
    _fullPathPopupMenuOpenForSeparator            = -1;
    _fullPathPopupPendingSeparatorMenuSwitchIndex = -1;
    _navDropdownKind                              = ModernDropdownKind::Siblings;
    _navDropdownPaths                             = siblings;
    _navDropdownSelectedIndex                     = selectedIndex;
    InvalidateRect(popupHwnd, nullptr, FALSE);

    auto clearPopupState = wil::scope_exit([&]() noexcept
    {
        _fullPathPopupActiveSeparatorIndex            = -1;
        _fullPathPopupMenuOpenForSeparator            = -1;
        _fullPathPopupPendingSeparatorMenuSwitchIndex = -1;
        _navDropdownKind                              = ModernDropdownKind::None;
        _navDropdownPaths.clear();
        _navDropdownSelectedIndex = -1;
        if (popupHwnd && IsWindow(popupHwnd) != FALSE)
        {
            InvalidateRect(popupHwnd, nullptr, FALSE);
        }
    });

    POINT pt = {static_cast<LONG>(separator.bounds.left), static_cast<LONG>(separator.bounds.bottom - _fullPathPopupScrollY)};
    ClientToScreen(popupHwnd, &pt);

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.ignoreInitialLeftButtonUp = true;

    const auto startedAt = std::chrono::steady_clock::now();
    const auto selectedId =
        RedSalamander::DxUi::ContextMenu::Show(popupHwnd, pt, items, MakeAppThemeDxPalette(_appTheme, ColorToCOLORREF(_theme.background)), sessionCallbacks);
    Debug::Perf::Emit(L"navigation.ui.dropdown_popup_us",
                      L"fullpath-siblings",
                      Debug::Perf::ElapsedUs(startedAt),
                      static_cast<uint64_t>(items.size()),
                      static_cast<uint64_t>(selectedIndex >= 0 ? selectedIndex : 0));

    if (selectedId.has_value() && selectedId.value() >= ID_SIBLING_BASE)
    {
        const size_t siblingIndex = static_cast<size_t>(selectedId.value() - ID_SIBLING_BASE);
        if (siblingIndex < siblings.size())
        {
            RequestPathChange(siblings[siblingIndex]);
            CloseFullPathPopup();
            return;
        }
    }
}

void NavigationView::RequestFullPathPopup(const D2D1_RECT_F& anchorBounds)
{
    if (! _hWnd)
    {
        return;
    }

    // Always align the popup with the path display area (Section Path).
    static_cast<void>(anchorBounds);

    POINT pt = {_sectionPathRect.left, _sectionPathRect.bottom};
    ClientToScreen(_hWnd.get(), &pt);

    _pendingFullPathPopupAnchor = pt;

    if (_menuOpenForSeparator != -1)
    {
        _pendingFullPathPopup = true;
        SendMessageW(_hWnd.get(), WM_CANCELMODE, 0, 0);
        PostMessageW(_hWnd.get(), WndMsg::kNavigationMenuShowFullPath, 0, 0);
        return;
    }

    _pendingFullPathPopup = true;
    ShowFullPathPopup();
}

void NavigationView::UpdateFullPathPopupWindow()
{
    if (! _hWnd || ! _currentPluginPath)
    {
        return;
    }

    EnsureD2DResources();
    if (! _d2dFactory || ! _dwriteFactory || ! _pathFormat || ! _separatorFormat)
    {
        return;
    }

    bool needsPathRedraw = false;
    if (_menuButtonHovered)
    {
        _menuButtonHovered = false;
        RenderDriveSection();
    }

    if (_historyButtonHovered)
    {
        _historyButtonHovered = false;
        RenderHistorySection();
    }

    if (_diskInfoHovered)
    {
        _diskInfoHovered = false;
        RenderDiskInfoSection();
    }

    if (_hoveredSegmentIndex != -1 || _hoveredSeparatorIndex != -1)
    {
        _hoveredSegmentIndex   = -1;
        _hoveredSeparatorIndex = -1;
        needsPathRedraw        = true;
    }

    if (needsPathRedraw)
    {
        RenderPathSection();
    }

    if (! RegisterFullPathPopupWndClass(_hInstance))
    {
        return;
    }

    const float paddingX       = DipsToPixels(kPathPaddingDip, _dpi);
    const float paddingY       = paddingX;
    const float separatorWidth = DipsToPixels(kPathSeparatorWidthDip, _dpi);
    const float spacing        = DipsToPixels(kPathSpacingDip, _dpi);
    const float lineHeight     = static_cast<float>(_sectionPathRect.bottom - _sectionPathRect.top);
    POINT anchor               = {_sectionPathRect.left, _sectionPathRect.bottom};
    ClientToScreen(_hWnd.get(), &anchor);
    _pendingFullPathPopupAnchor = anchor;

    auto parts = SplitPathComponents(_currentPluginPath.value());
    if (parts.empty())
    {
        return;
    }

    float sumWidths = 0.0f;
    for (const auto& part : parts)
    {
        sumWidths += MeasureTextWidth(_dwriteFactory.get(), _pathFormat.get(), part.text, kIntrinsicTextLayoutMaxWidth, lineHeight);
    }

    const size_t segmentCount          = parts.size();
    const float contentSingleLineWidth = sumWidths + spacing * static_cast<float>(segmentCount) + separatorWidth * static_cast<float>(segmentCount - 1);

    HMONITOR hMon = MonitorFromPoint(_pendingFullPathPopupAnchor, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (! GetMonitorInfoW(hMon, &mi))
    {
        return;
    }

    const RECT work             = mi.rcWork;
    const float maxClientWidth  = static_cast<float>(std::max(0L, work.right - work.left));
    const float maxClientHeight = static_cast<float>(std::max(0L, work.bottom - work.top));

    const DWORD style   = WS_POPUP | WS_BORDER;
    const DWORD exStyle = WS_EX_TOOLWINDOW;

    RECT nonClientRect = {0, 0, 0, 0};
    if (! AdjustWindowRectExForDpi(&nonClientRect, style, FALSE, exStyle, _dpi))
    {
        AdjustWindowRectEx(&nonClientRect, style, FALSE, exStyle);
    }

    const LONG nonClientWidth         = nonClientRect.right - nonClientRect.left;
    const LONG maxWindowWidthForX     = std::max(0L, work.right - anchor.x);
    const float maxAlignedClientWidth = std::max(1.0f, static_cast<float>(std::max(0L, maxWindowWidthForX - nonClientWidth)));

    const float desiredClientWidth = std::max(1.0f, std::min(contentSingleLineWidth + paddingX * 2.0f, std::min(maxClientWidth, maxAlignedClientWidth)));

    _fullPathPopupClientSize.cx = static_cast<LONG>(std::ceil(desiredClientWidth));
    _fullPathPopupClientSize.cy = 1;
    BuildFullPathPopupLayout(static_cast<float>(_fullPathPopupClientSize.cx));

    const float desiredClientHeight = std::min(std::max(lineHeight + paddingY * 2.0f, _fullPathPopupContentHeight), maxClientHeight);
    _fullPathPopupClientSize.cy     = static_cast<LONG>(std::ceil(std::max(1.0f, desiredClientHeight)));
    _fullPathPopupScrollY           = 0.0f;

    RECT windowRect = {0, 0, _fullPathPopupClientSize.cx, _fullPathPopupClientSize.cy};
    if (! AdjustWindowRectExForDpi(&windowRect, style, FALSE, exStyle, _dpi))
    {
        AdjustWindowRectEx(&windowRect, style, FALSE, exStyle);
    }

    const int winWidth  = windowRect.right - windowRect.left;
    const int winHeight = windowRect.bottom - windowRect.top;

    int x = anchor.x;
    int y = anchor.y;

    if (y + winHeight > work.bottom)
    {
        const int aboveY = anchor.y - winHeight;
        if (aboveY >= work.top)
        {
            y = aboveY;
        }
        else
        {
            y = std::max(static_cast<int>(work.top), static_cast<int>(work.bottom - winHeight));
        }
    }

    if (x + winWidth > work.right)
    {
        x = std::max(static_cast<int>(work.left), static_cast<int>(work.right - winWidth));
    }

    const int minX = static_cast<int>(work.left);
    const int minY = static_cast<int>(work.top);
    const int maxX = std::max(minX, static_cast<int>(work.right - winWidth));
    const int maxY = std::max(minY, static_cast<int>(work.bottom - winHeight));
    x              = std::clamp(x, minX, maxX);
    y              = std::clamp(y, minY, maxY);

    if (! _fullPathPopup)
    {
        HWND popup = CreateWindowExW(exStyle, kFullPathPopupClassName, L"", style, x, y, winWidth, winHeight, _hWnd.get(), nullptr, _hInstance, this);
        if (! popup)
        {
            return;
        }

        _fullPathPopup.reset(popup);
    }
    else
    {
        SetWindowPos(_fullPathPopup.get(), HWND_TOP, x, y, winWidth, winHeight, SWP_NOACTIVATE);
    }

    RECT clientRect{};
    GetClientRect(_fullPathPopup.get(), &clientRect);
    _fullPathPopupClientSize.cx = clientRect.right - clientRect.left;
    _fullPathPopupClientSize.cy = clientRect.bottom - clientRect.top;

    BuildFullPathPopupLayout(static_cast<float>(_fullPathPopupClientSize.cx));
}

void NavigationView::ShowFullPathPopup()
{
    if (! _pendingFullPathPopup)
    {
        return;
    }

    _pendingFullPathPopup = false;

    bool needsPathRedraw = false;
    if (_menuButtonHovered)
    {
        _menuButtonHovered = false;
        RenderDriveSection();
    }

    if (_historyButtonHovered)
    {
        _historyButtonHovered = false;
        RenderHistorySection();
    }

    if (_diskInfoHovered)
    {
        _diskInfoHovered = false;
        RenderDiskInfoSection();
    }

    if (_hoveredSegmentIndex != -1 || _hoveredSeparatorIndex != -1)
    {
        _hoveredSegmentIndex   = -1;
        _hoveredSeparatorIndex = -1;
        needsPathRedraw        = true;
    }

    if (needsPathRedraw)
    {
        RenderPathSection();
    }

    CloseFullPathPopup();
    UpdateFullPathPopupWindow();
    if (! _fullPathPopup)
    {
        return;
    }

    ShowWindow(_fullPathPopup.get(), SW_SHOW);
    SetForegroundWindow(_fullPathPopup.get());
    SetFocus(_fullPathPopup.get());
    InvalidateRect(_fullPathPopup.get(), nullptr, FALSE);
}

void NavigationView::CloseFullPathPopup()
{
    if (_fullPathPopupEdit)
    {
        if (_fullPathPopupEdit->field)
        {
            ClearEditValidationError(*_fullPathPopupEdit);
            _fullPathPopupEdit->field->SetOnTextChanged({});
            _fullPathPopupEdit->field->SetOnSubmitted({});
            _fullPathPopupEdit->field->SetOnBlur({});
        }
        _fullPathPopupEdit->host.SetOnEscape({});
        _fullPathPopupEdit->host.SetOnTabBoundary({});
        _fullPathPopupEdit->DeactivateForHideOrDestroy();
        if (_fullPathPopupEdit->hwnd && IsWindow(_fullPathPopupEdit->hwnd.get()) != FALSE)
        {
            ShowWindow(_fullPathPopupEdit->hwnd.get(), SW_HIDE);
        }
    }
    _fullPathPopupEditMode = false;

    if (_fullPathPopup)
    {
        _fullPathPopup.reset();
    }
}

void NavigationView::DiscardFullPathPopupD2DResources()
{
    _fullPathPopupBackgroundBrush = nullptr;
    _fullPathPopupAccentBrush     = nullptr;
    _fullPathPopupPressedBrush    = nullptr;
    _fullPathPopupHoverBrush      = nullptr;
    _fullPathPopupSeparatorBrush  = nullptr;
    _fullPathPopupTextBrush       = nullptr;
    _fullPathPopupTarget          = nullptr;
}

void NavigationView::EnsureFullPathPopupD2DResources()
{
    if (! _fullPathPopup)
    {
        return;
    }

    EnsureD2DResources();
    if (! _d2dFactory)
    {
        return;
    }

    if (! _fullPathPopupTarget)
    {
        D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
        props.dpiX                          = 96.0f;
        props.dpiY                          = 96.0f;

        D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(
            _fullPathPopup.get(), D2D1::SizeU(static_cast<UINT32>(_fullPathPopupClientSize.cx), static_cast<UINT32>(_fullPathPopupClientSize.cy)));

        wil::com_ptr<ID2D1HwndRenderTarget> target;
        if (FAILED(_d2dFactory->CreateHwndRenderTarget(props, hwndProps, target.addressof())))
        {
            return;
        }

        target->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
        _fullPathPopupTarget = std::move(target);
    }

    if (_fullPathPopupTarget && ! _fullPathPopupTextBrush)
    {
        _fullPathPopupTarget->CreateSolidColorBrush(_theme.text, _fullPathPopupTextBrush.addressof());
        _fullPathPopupTarget->CreateSolidColorBrush(_theme.separator, _fullPathPopupSeparatorBrush.addressof());
        _fullPathPopupTarget->CreateSolidColorBrush(_theme.hoverHighlight, _fullPathPopupHoverBrush.addressof());
        _fullPathPopupTarget->CreateSolidColorBrush(_theme.pressedHighlight, _fullPathPopupPressedBrush.addressof());
        _fullPathPopupTarget->CreateSolidColorBrush(_theme.accent, _fullPathPopupAccentBrush.addressof());
        _fullPathPopupTarget->CreateSolidColorBrush(_theme.background, _fullPathPopupBackgroundBrush.addressof());
    }
}

void NavigationView::BuildFullPathPopupLayout(float clientWidth)
{
    _fullPathPopupSegments.clear();
    _fullPathPopupSeparators.clear();
    _fullPathPopupHoveredSegmentIndex   = -1;
    _fullPathPopupHoveredSeparatorIndex = -1;

    _fullPathPopupContentHeight = 0.0f;

    if (! _currentPluginPath || ! _dwriteFactory || ! _pathFormat)
    {
        return;
    }

    auto parts = SplitPathComponents(_currentPluginPath.value());
    if (parts.empty())
    {
        return;
    }

    const float paddingX        = DipsToPixels(kPathPaddingDip, _dpi);
    const float paddingY        = paddingX;
    const float separatorWidth  = DipsToPixels(kPathSeparatorWidthDip, _dpi);
    const float spacing         = DipsToPixels(kPathSpacingDip, _dpi);
    const float lineHeight      = static_cast<float>(_sectionPathRect.bottom - _sectionPathRect.top);
    const float maxContentWidth = std::max(0.0f, clientWidth - paddingX * 2.0f);

    constexpr std::wstring_view ellipsisText = kEllipsisText;

    float x = paddingX;
    float y = paddingY;

    for (size_t i = 0; i < parts.size(); ++i)
    {
        const float maxSegmentWidth = (i == 0) ? maxContentWidth : std::max(0.0f, maxContentWidth - separatorWidth);

        std::wstring displayText = parts[i].text;
        wil::com_ptr<IDWriteTextLayout> layout;
        float segWidth = 0.0f;
        CreateTextLayoutAndWidth(_dwriteFactory.get(), _pathFormat.get(), displayText, kIntrinsicTextLayoutMaxWidth, lineHeight, layout, segWidth);
        if (segWidth > maxSegmentWidth && maxSegmentWidth > 0.0f)
        {
            displayText = TruncateTextToWidth(_dwriteFactory.get(), _pathFormat.get(), displayText, maxSegmentWidth, lineHeight, ellipsisText);
            CreateTextLayoutAndWidth(_dwriteFactory.get(), _pathFormat.get(), displayText, kIntrinsicTextLayoutMaxWidth, lineHeight, layout, segWidth);
        }

        if (i > 0)
        {
            const float lineLimit = paddingX + maxContentWidth;
            // Use spacing/2 because the last segment on a line does not need the trailing spacing.
            if (x + separatorWidth + segWidth + spacing / 2.0f > lineLimit && x > paddingX)
            {
                x = paddingX;
                y += lineHeight;
            }

            BreadcrumbSeparator sep;
            sep.bounds            = D2D1::RectF(x, y, x + separatorWidth, y + lineHeight);
            sep.leftSegmentIndex  = _fullPathPopupSegments.size() - 1;
            sep.rightSegmentIndex = _fullPathPopupSegments.size();
            _fullPathPopupSeparators.push_back(sep);
            x += separatorWidth;
        }

        PathSegment segment;
        segment.text       = std::move(displayText);
        segment.fullPath   = parts[i].fullPath;
        segment.isEllipsis = false;
        segment.layout     = std::move(layout);
        segment.bounds     = D2D1::RectF(x - spacing / 2.0f, y, x + segWidth + spacing / 2.0f, y + lineHeight);
        _fullPathPopupSegments.push_back(std::move(segment));

        x += segWidth + spacing;
    }

    _fullPathPopupContentHeight = y + lineHeight + paddingY;

    if (_fullPathPopupClientSize.cy > 0)
    {
        const float maxScroll = std::max(0.0f, _fullPathPopupContentHeight - static_cast<float>(_fullPathPopupClientSize.cy));
        _fullPathPopupScrollY = std::clamp(_fullPathPopupScrollY, 0.0f, maxScroll);
    }
    else
    {
        _fullPathPopupScrollY = 0.0f;
    }
}

void NavigationView::RenderFullPathPopup()
{
    if (! _fullPathPopup)
    {
        return;
    }

    PAINTSTRUCT ps;
    wil::unique_hdc_paint hdc = wil::BeginPaint(_fullPathPopup.get(), &ps);
    static_cast<void>(ps);
    static_cast<void>(hdc.get());

    EnsureFullPathPopupD2DResources();
    if (! _fullPathPopupTarget)
    {
        return;
    }

    _fullPathPopupTarget->BeginDraw();
    const auto endDraw = wil::scope_exit([&]
    {
        if (! _fullPathPopupTarget)
        {
            return;
        }

        const HRESULT hrEnd = _fullPathPopupTarget->EndDraw();
        if (FAILED(hrEnd))
        {
            if (hrEnd == D2DERR_RECREATE_TARGET)
            {
                DiscardFullPathPopupD2DResources();
                return;
            }
            Debug::Error(L"[NavigationView] NavigationView::RenderFullPathPopup EndDraw failed (hr=0x{:08X})", static_cast<unsigned long>(hrEnd));
            return;
        }
    });

    const D2D1_RECT_F clientRect = D2D1::RectF(0.0f, 0.0f, static_cast<float>(_fullPathPopupClientSize.cx), static_cast<float>(_fullPathPopupClientSize.cy));
    if (_fullPathPopupBackgroundBrush)
    {
        _fullPathPopupTarget->FillRectangle(clientRect, _fullPathPopupBackgroundBrush.get());
    }

    if (! _fullPathPopupEditMode)
    {
        const float textInsetX        = DipsToPixels(kPathTextInsetDip, _dpi);
        const float hoverInset        = DipsToPixels(kBreadcrumbHoverInsetDip, _dpi);
        const float hoverCornerRadius = DipsToPixels(kBreadcrumbHoverCornerRadiusDip, _dpi);

        _fullPathPopupTarget->PushAxisAlignedClip(clientRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
        _fullPathPopupTarget->SetTransform(D2D1::Matrix3x2F::Translation(0.0f, -_fullPathPopupScrollY));

        for (size_t i = 0; i < _fullPathPopupSegments.size(); ++i)
        {
            const auto& segment = _fullPathPopupSegments[i];
            if (_fullPathPopupHoveredSegmentIndex == static_cast<int>(i) && _fullPathPopupHoverBrush)
            {
                const D2D1_RECT_F hoverRect = InsetRectF(segment.bounds, hoverInset, hoverInset);
                _fullPathPopupTarget->FillRoundedRectangle(RoundedRect(hoverRect, hoverCornerRadius), _fullPathPopupHoverBrush.get());
            }

            ID2D1SolidColorBrush* brush = _fullPathPopupTextBrush.get();
            const bool lastSegment      = i == (_fullPathPopupSegments.size() - 1);
            if (lastSegment && _fullPathPopupAccentBrush)
            {
                brush = _fullPathPopupAccentBrush.get();
            }

            if (segment.layout && brush)
            {
                _fullPathPopupTarget->DrawTextLayout(
                    D2D1::Point2F(segment.bounds.left + textInsetX, segment.bounds.top), segment.layout.get(), brush, D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
            }
        }

        for (size_t i = 0; i < _fullPathPopupSeparators.size(); ++i)
        {
            const auto& separator = _fullPathPopupSeparators[i];
            if (_fullPathPopupHoveredSeparatorIndex == static_cast<int>(i) && _fullPathPopupHoverBrush)
            {
                const D2D1_RECT_F hoverRect = InsetRectF(separator.bounds, hoverInset, hoverInset);
                _fullPathPopupTarget->FillRoundedRectangle(RoundedRect(hoverRect, hoverCornerRadius), _fullPathPopupHoverBrush.get());
            }
            else if (_fullPathPopupActiveSeparatorIndex == static_cast<int>(i) && _fullPathPopupPressedBrush)
            {
                const D2D1_RECT_F pressedRect = InsetRectF(separator.bounds, hoverInset, hoverInset);
                _fullPathPopupTarget->FillRoundedRectangle(RoundedRect(pressedRect, hoverCornerRadius), _fullPathPopupPressedBrush.get());
            }

            if (_separatorFormat && _fullPathPopupSeparatorBrush)
            {
                _fullPathPopupTarget->DrawText(&_breadcrumbSeparatorGlyph, 1, _separatorFormat.get(), separator.bounds, _fullPathPopupSeparatorBrush.get());
            }
        }

        _fullPathPopupTarget->SetTransform(D2D1::Matrix3x2F::Identity());
        _fullPathPopupTarget->PopAxisAlignedClip();
    }
}

void NavigationView::OnFullPathPopupExitMenuLoop(HWND popupHwnd, [[maybe_unused]] bool isShortcut)
{
    if (_fullPathPopupMenuOpenForSeparator == -1 && _fullPathPopupActiveSeparatorIndex == -1)
    {
        return;
    }

    _fullPathPopupMenuOpenForSeparator            = -1;
    _fullPathPopupActiveSeparatorIndex            = -1;
    _fullPathPopupPendingSeparatorMenuSwitchIndex = -1;

    if (popupHwnd)
    {
        InvalidateRect(popupHwnd, nullptr, FALSE);
    }
}

void NavigationView::EnterFullPathPopupEditMode()
{
    if (! _fullPathPopup || ! _currentEditPath || _fullPathPopupEditMode)
    {
        return;
    }

    _fullPathPopupEditMode = true;

    if (! _fullPathPopupEdit)
    {
        if (! RegisterDxHostWndClass(_hInstance))
        {
            _fullPathPopupEditMode = false;
            return;
        }

        RECT rc{};
        GetClientRect(_fullPathPopup.get(), &rc);
        const int hostWidth  = static_cast<int>((std::max)(0L, rc.right - rc.left));
        const int hostHeight = static_cast<int>((std::max)(0L, rc.bottom - rc.top));

        auto hostState = std::make_unique<NavigationDxTextHost>();
        HWND hwnd      = CreateWindowExW(0,
                                         kDxHostClassName,
                                         L"",
                                         WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                         rc.left,
                                         rc.top,
                                         hostWidth,
                                         hostHeight,
                                         _fullPathPopup.get(),
                                         reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_PATH_EDIT)),
                                         _hInstance,
                                         &hostState->host);
        if (! hwnd)
        {
            _fullPathPopupEditMode = false;
            return;
        }

        hostState->hwnd.reset(hwnd);
        hostState->host.SetTextInputBackend(RedSalamander::DxUi::TextInputBackend::Native);
        hostState->host.SetTheme(MakeNavigationDxEditPalette(_appTheme, _theme));

        auto field       = std::make_unique<RedSalamander::DxUi::TextField>();
        hostState->field = field.get();
        hostState->field->SetMultiline(false);
        hostState->field->SetClearButtonEnabled(false);
        hostState->field->SetCaretColor(_theme.text);
        hostState->field->SetHorizontalTextPadding(2.0f, 2.0f);
        hostState->field->SetVerticalTextPadding(1.0f, 1.0f);
        hostState->host.SetRoot(std::move(field));
        _fullPathPopupEdit = std::move(hostState);
    }

    if (! _fullPathPopupEdit || ! _fullPathPopupEdit->field || ! _fullPathPopupEdit->hwnd)
    {
        _fullPathPopupEditMode = false;
        return;
    }

    _fullPathPopupEdit->field->SetOnTextChanged([this](std::wstring_view /*text*/)
    {
        if (_fullPathPopupEdit)
        {
            ClearEditValidationError(*_fullPathPopupEdit);
        }
    });
    _fullPathPopupEdit->field->SetOnSubmitted([this]() { ExitFullPathPopupEditMode(true); });
    _fullPathPopupEdit->field->SetOnBlur([this]() noexcept
    {
        if (_fullPathPopupEditMode)
        {
            ExitFullPathPopupEditMode(false);
        }
    });
    _fullPathPopupEdit->host.SetOnEscape([this]() -> bool
    {
        _restoreFolderViewFocusAfterFullPathPopupClose = true;
        ExitFullPathPopupEditMode(false);
        return true;
    });
    _fullPathPopupEdit->host.SetOnTabBoundary([this](bool reverse) -> bool
    {
        ExitFullPathPopupEditMode(false);
        if (_requestFolderViewFocusCallback)
        {
            _requestFolderViewFocusCallback();
            return true;
        }

        if (_hWnd)
        {
            SetFocus(_hWnd.get());
        }
        MoveFocus(! reverse);
        return true;
    });

    _fullPathPopupEdit->field->SetText(_currentEditPath.value().native());
    _fullPathPopupEdit->field->SetSelectionRange(0u, _currentEditPath.value().native().size());
    UpdateFullPathPopupEditHostLayout();
    ApplyDxEditHostThemes();
    ShowWindow(_fullPathPopupEdit->hwnd.get(), SW_SHOW);
    InstallEditHostHook(*_fullPathPopupEdit);
    _fullPathPopupEdit->host.SetFocusControl(_fullPathPopupEdit->field);
    SetFocus(_fullPathPopupEdit->hwnd.get());
    InvalidateRect(_fullPathPopup.get(), nullptr, FALSE);
}

void NavigationView::ExitFullPathPopupEditMode(bool accept)
{
    if (! _fullPathPopupEditMode)
    {
        return;
    }

    if (! accept)
    {
        CloseFullPathPopup();
        return;
    }

    if (! _fullPathPopupEdit || ! _fullPathPopupEdit->field)
    {
        CloseFullPathPopup();
        return;
    }

    const std::wstring buffer(_fullPathPopupEdit->field->GetText());

    if (ValidatePath(buffer))
    {
        std::filesystem::path newPath(NavigationLocation::NormalizeUserTypedLocationText(buffer));
        const bool isFilePlugin = _pluginShortId.empty() || EqualsNoCase(_pluginShortId, L"file");
        const std::wstring_view typedText(newPath.native());
        if (! isFilePlugin && ! _currentInstanceContext.empty() && typedText.find(L'|') != std::wstring_view::npos)
        {
            std::wstring_view typedPrefix;
            std::wstring_view typedRemainder;
            if (! TryParsePluginPrefix(typedText, typedPrefix, typedRemainder))
            {
                std::wstring canonical;
                canonical.reserve(_pluginShortId.size() + 1u + typedText.size());
                canonical.append(_pluginShortId);
                canonical.push_back(L':');
                canonical.append(typedText);
                newPath = std::filesystem::path(std::move(canonical));
            }
        }
        bool changed = true;
        if (_currentEditPath)
        {
            changed = wil::compare_string_ordinal(newPath.wstring(), _currentEditPath.value().wstring(), true) != wistd::weak_ordering::equivalent;
        }

        if (changed)
        {
            RequestPathChange(newPath);
            CloseFullPathPopup();
            return;
        }

        _fullPathPopupEditMode = false;
        _fullPathPopupEdit->DeactivateForHideOrDestroy();
        const HWND focused = GetFocus();
        if (focused && _fullPathPopupEdit->hwnd &&
            (focused == _fullPathPopupEdit->hwnd.get() || focused == _fullPathPopupEdit->GetTextInputHwnd() ||
             IsChild(_fullPathPopupEdit->hwnd.get(), focused) != FALSE))
        {
            SetFocus(_fullPathPopup.get());
        }
        ShowWindow(_fullPathPopupEdit->hwnd.get(), SW_HIDE);
        if (_fullPathPopup)
        {
            SetFocus(_fullPathPopup.get());
            InvalidateRect(_fullPathPopup.get(), nullptr, FALSE);
        }
        return;
    }

    const std::wstring message = FormatStringResource(nullptr, IDS_FMT_INVALID_PATH, buffer);
    ShowEditValidationError(*_fullPathPopupEdit, message);
    if (const HWND inputHwnd = _fullPathPopupEdit->GetTextInputHwnd(); inputHwnd && IsWindow(inputHwnd) != FALSE)
    {
        SetFocus(inputHwnd);
    }
    else if (_fullPathPopupEdit->hwnd)
    {
        SetFocus(_fullPathPopupEdit->hwnd.get());
    }
}
