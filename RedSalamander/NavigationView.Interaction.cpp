#include "NavigationViewInternal.h"

#include <windowsx.h>

#include <commctrl.h>

#include "DirectoryInfoCache.h"
#include "Helpers.h"
#include "IconCache.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "resource.h"

void NavigationView::OnLButtonDown(POINT pt)
{
    if (_editMode)
    {
        const auto chrome = ComputeEditChromeRects(_sectionPathRect, _dpi);
        if (PtInRect(&chrome.closeRect, pt))
        {
            ExitEditMode(false);
        }
        return;
    }

    if (_hWnd)
    {
        SetFocus(_hWnd.get());
    }

    // Check Section 1 (menu button) click
    if (_showMenuSection && PtInRect(&_sectionDriveRect, pt))
    {
        _focusedRegion = FocusRegion::Menu;
        ShowMenuDropdown();
        return;
    }

    // Check history button click (Section 2 area)
    if (PtInRect(&_sectionHistoryRect, pt))
    {
        _focusedRegion = FocusRegion::History;
        ShowHistoryDropdown();
        return;
    }

    // Check Section 3 (disk info) click
    if (_showDiskInfoSection && PtInRect(&_sectionDiskInfoRect, pt))
    {
        _focusedRegion = FocusRegion::DiskInfo;
        ShowDiskInfoDropdown();
        return;
    }

    // Check if click is in Section 2 (breadcrumbs)
    if (! PtInRect(&_sectionPathRect, pt))
        return;

    _focusedRegion = FocusRegion::Path;

    // Transform to Section 2 local coordinates
    float localX          = static_cast<float>(pt.x - _sectionPathRect.left);
    float localY          = static_cast<float>(pt.y - _sectionPathRect.top);
    D2D1_POINT_2F clickPt = D2D1::Point2F(localX, localY);

    // Check breadcrumb segments
    for (size_t i = 0; i < _segments.size(); ++i)
    {
        const auto& segment = _segments[i];

        if (segment.bounds.left <= clickPt.x && clickPt.x <= segment.bounds.right && segment.bounds.top <= clickPt.y && clickPt.y <= segment.bounds.bottom)
        {
            if (segment.isEllipsis)
            {
                RequestFullPathPopup(segment.bounds);
                return;
            }

            // Navigate to this segment's path
            RequestPathChange(segment.fullPath);
            return;
        }
    }

    // Check separator clicks for sibling navigation
    for (size_t i = 0; i < _separators.size(); i++)
    {
        const auto& separator = _separators[i];
        const auto& bounds    = separator.bounds;

        if (bounds.left <= clickPt.x && clickPt.x <= bounds.right && bounds.top <= clickPt.y && clickPt.y <= bounds.bottom)
        {
            const bool adjacentToEllipsis = (separator.leftSegmentIndex < _segments.size() && _segments[separator.leftSegmentIndex].isEllipsis) ||
                                            (separator.rightSegmentIndex < _segments.size() && _segments[separator.rightSegmentIndex].isEllipsis);
            if (adjacentToEllipsis)
            {
                RequestFullPathPopup(bounds);
                return;
            }

            // If a different separator menu is open, close it first
            if (_menuOpenForSeparator != -1 && _menuOpenForSeparator != static_cast<int>(i))
            {
                SendMessageW(_hWnd.get(), WM_CANCELMODE, 0, 0);
                PostMessageW(_hWnd.get(), WndMsg::kNavigationMenuShowSiblingsDropdown, static_cast<WPARAM>(i), 0);
            }
            else
            {
                ShowSiblingsDropdown(i);
            }
            return;
        }
    }
}

void NavigationView::OnLButtonDblClk(POINT pt)
{
    if (_editMode)
        return;

    if (_hWnd)
    {
        SetFocus(_hWnd.get());
    }

    // Check if click is in Section 2
    if (! PtInRect(&_sectionPathRect, pt))
        return;

    _focusedRegion = FocusRegion::Path;

    // Transform to Section 2 local coordinates
    float localX          = static_cast<float>(pt.x - _sectionPathRect.left);
    float localY          = static_cast<float>(pt.y - _sectionPathRect.top);
    D2D1_POINT_2F clickPt = D2D1::Point2F(localX, localY);

    // Check if double-click is on the last segment or in empty space
    bool onLastSegment = false;
    if (! _segments.empty())
    {
        const auto& lastSegment = _segments.back();

        if (lastSegment.bounds.left <= clickPt.x && clickPt.x <= lastSegment.bounds.right && lastSegment.bounds.top <= clickPt.y &&
            clickPt.y <= lastSegment.bounds.bottom)
        {
            onLastSegment = true;
        }
    }

    // Check if double-click is in whitespace (after all segments)
    bool inWhitespace = true;
    for (const auto& segment : _segments)
    {
        if (segment.bounds.left <= clickPt.x && clickPt.x <= segment.bounds.right && segment.bounds.top <= clickPt.y && clickPt.y <= segment.bounds.bottom)
        {
            inWhitespace = false;
            break;
        }
    }

    // Also check separators
    if (inWhitespace)
    {
        for (const auto& separator : _separators)
        {
            const auto& bounds = separator.bounds;
            if (bounds.left <= clickPt.x && clickPt.x <= bounds.right && bounds.top <= clickPt.y && clickPt.y <= bounds.bottom)
            {
                inWhitespace = false;
                break;
            }
        }
    }

    // Enter edit mode if double-clicked on last segment or in whitespace
    if (onLastSegment || inWhitespace)
    {
        EnterEditMode();
    }
}

void NavigationView::OnMouseMove(POINT pt)
{
    if (_fullPathPopup)
        return;

    if (_editMode)
        return;

    // Track mouse for leave notification
    if (! _trackingMouse)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = _hWnd.get();
        TrackMouseEvent(&tme);
        _trackingMouse = true;
    }

    // Check hover states - mutually exclusive priority order
    bool menuButtonHovered    = false;
    bool historyButtonHovered = false;
    bool diskInfoHovered      = false;

    // Priority 1: Check Section 1 (menu button) hover
    if (_showMenuSection && PtInRect(&_sectionDriveRect, pt))
    {
        menuButtonHovered = true;
    }
    // Priority 2: Check history button hover (inside Section 2)
    else if (PtInRect(&_sectionHistoryRect, pt))
    {
        historyButtonHovered = true;
    }
    // Priority 3: Check Section 3 (disk info) hover
    else if (_showDiskInfoSection && PtInRect(&_sectionDiskInfoRect, pt))
    {
        diskInfoHovered = true;
    }

    // Update Section 1 if hover state changed
    if (menuButtonHovered != _menuButtonHovered)
    {
        _menuButtonHovered = menuButtonHovered;
        RenderDriveSection();
    }

    // Update history button if hover state changed
    if (historyButtonHovered != _historyButtonHovered)
    {
        _historyButtonHovered = historyButtonHovered;
        RenderHistorySection(); // History button is rendered in Section 3
    }

    // Update Section 3 if hover state changed
    if (diskInfoHovered != _diskInfoHovered)
    {
        _diskInfoHovered = diskInfoHovered;
        RenderDiskInfoSection();
    }

    bool needsRedraw = false;

    // Transform mouse coordinates to Section 2 local space
    float localX         = static_cast<float>(pt.x - _sectionPathRect.left);
    float localY         = static_cast<float>(pt.y - _sectionPathRect.top);
    D2D1_POINT_2F movePt = D2D1::Point2F(localX, localY);

    // Track segment hover (using local coordinates)
    int newHoveredSegment = -1;
    for (size_t i = 0; i < _segments.size(); ++i)
    {
        const auto& bounds = _segments[i].bounds;
        if (bounds.left <= movePt.x && movePt.x <= bounds.right && bounds.top <= movePt.y && movePt.y <= bounds.bottom)
        {
            newHoveredSegment = static_cast<int>(i);
            break;
        }
    }

    if (newHoveredSegment != _hoveredSegmentIndex)
    {
        _hoveredSegmentIndex = newHoveredSegment;
        needsRedraw          = true;
    }

    // Track separator hover (using local coordinates)
    int oldHoveredSeparator = _hoveredSeparatorIndex;
    _hoveredSeparatorIndex  = -1;

    for (size_t i = 0; i < _separators.size(); ++i)
    {
        const auto& bounds = _separators[i].bounds;
        if (bounds.left <= movePt.x && movePt.x <= bounds.right && bounds.top <= movePt.y && movePt.y <= bounds.bottom)
        {
            _hoveredSeparatorIndex = static_cast<int>(i);
            break;
        }
    }

    if (oldHoveredSeparator != _hoveredSeparatorIndex)
    {
        needsRedraw = true;
    }

    if (needsRedraw)
    {
        RenderPathSection();
    }
}

void NavigationView::OnMouseLeave()
{
    _trackingMouse = false;

    // Clear drive menu hover
    if (_menuButtonHovered)
    {
        _menuButtonHovered = false;
        RenderDriveSection(); // Re-render with normal state
    }

    // Clear disk info hover
    if (_diskInfoHovered)
    {
        _diskInfoHovered = false;
        RenderDiskInfoSection(); // Re-render with normal state
    }

    // Clear history button hover
    if (_historyButtonHovered)
    {
        _historyButtonHovered = false;
        RenderHistorySection(); // Re-render with normal state;
    }

    // Clear separator hover
    bool hadHoveredSeparator = (_hoveredSeparatorIndex != -1);
    _hoveredSeparatorIndex   = -1;

    // Clear segment hover
    bool hadHoveredSegment = (_hoveredSegmentIndex != -1);
    _hoveredSegmentIndex   = -1;

    const bool hadEditCloseHovered = _editCloseHovered;
    _editCloseHovered              = false;

    bool needsRedraw = hadHoveredSeparator || hadHoveredSegment || hadEditCloseHovered;
    if (needsRedraw)
    {
        RenderPathSection();
    }
}

void NavigationView::OnSetCursor([[maybe_unused]] HWND hwnd, UINT hitTest, [[maybe_unused]] UINT mouseMsg)
{
    if (_fullPathPopup)
    {
        SetCursor(LoadCursor(nullptr, IDC_ARROW));
        return;
    }

    if (hitTest == HTCLIENT)
    {
        POINT pt;
        GetCursorPos(&pt);
        ScreenToClient(_hWnd.get(), &pt);

        // Check hover states - mutually exclusive priority order
        bool menuButtonHovered    = false;
        bool historyButtonHovered = false;
        bool diskInfoHovered      = false;

        // Priority 1: Check Section 1 (menu button) hover
        if (_showMenuSection && PtInRect(&_sectionDriveRect, pt))
        {
            menuButtonHovered = true;
        }
        // Priority 2: Check history button hover (inside Section 2)
        else if (PtInRect(&_sectionHistoryRect, pt))
        {
            historyButtonHovered = true;
        }
        // Priority 3: Check Section 3 (disk info) hover
        else if (_showDiskInfoSection && PtInRect(&_sectionDiskInfoRect, pt))
        {
            diskInfoHovered = true;
        }

        // Update Section 1 if hover state changed
        if (menuButtonHovered != _menuButtonHovered)
        {
            _menuButtonHovered = menuButtonHovered;
            RenderDriveSection();
        }

        // Update history button if hover state changed
        if (historyButtonHovered != _historyButtonHovered)
        {
            _historyButtonHovered = historyButtonHovered;
            RenderHistorySection();
        }

        // Update Section 3 if hover state changed
        if (diskInfoHovered != _diskInfoHovered)
        {
            _diskInfoHovered = diskInfoHovered;
            RenderDiskInfoSection();
        }

        // Check Section 2 segments and separators (hover tracking now in OnTimer)
        if (_editMode)
        {
            const auto chrome = ComputeEditChromeRects(_sectionPathRect, _dpi);
            if (PtInRect(&chrome.closeRect, pt))
            {
                SetCursor(LoadCursor(nullptr, IDC_HAND));
                return;
            }
        }

        if (PtInRect(&_sectionPathRect, pt))
        {
            SetCursor(LoadCursor(nullptr, IDC_ARROW));
            return;
        }

        SetCursor(LoadCursor(nullptr, IDC_ARROW));
    }
}

void NavigationView::OnTimer(UINT_PTR timerId)
{
    if (timerId != HOVER_TIMER_ID || _hoverTimer == 0)
    {
        return;
    }

    if (_fullPathPopup)
    {
        bool needsRedraw = false;

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
            needsRedraw            = true;
        }

        if (needsRedraw)
        {
            RenderPathSection();
        }

        return;
    }

    // Check if cursor is over our window
    POINT screenPt{};
    GetCursorPos(&screenPt);

    HWND windowAtPoint  = WindowFromPoint(screenPt);
    const bool overMenu = IsWin32MenuWindow(windowAtPoint);

    POINT pt = screenPt;
    ScreenToClient(_hWnd.get(), &pt);

    RECT clientRect;
    GetClientRect(_hWnd.get(), &clientRect);
    bool inClient = ! overMenu && (PtInRect(&clientRect, pt) != 0);

    // Check Drive Menu hover
    bool menuButtonHovered = _showMenuSection && inClient && PtInRect(&_sectionDriveRect, pt);
    if (menuButtonHovered != _menuButtonHovered)
    {
        _menuButtonHovered = menuButtonHovered;
        RenderDriveSection(); // Re-render with new hover state
    }

    // Check history button hover
    bool historyButtonHovered = inClient && PtInRect(&_sectionHistoryRect, pt);
    if (historyButtonHovered != _historyButtonHovered)
    {
        _historyButtonHovered = historyButtonHovered;
        RenderHistorySection(); // Re-render with new hover state
    }

    // Check Disk Info hover
    bool diskInfoHovered = _showDiskInfoSection && inClient && PtInRect(&_sectionDiskInfoRect, pt);
    if (diskInfoHovered != _diskInfoHovered)
    {
        _diskInfoHovered = diskInfoHovered;
        RenderDiskInfoSection(); // Re-render with new hover state
    }

    if (_editMode)
    {
        bool needsRedraw = false;

        const auto chrome       = ComputeEditChromeRects(_sectionPathRect, _dpi);
        const bool closeHovered = inClient && PtInRect(&chrome.closeRect, pt);
        if (closeHovered != _editCloseHovered)
        {
            _editCloseHovered = closeHovered;
            needsRedraw       = true;
        }

        if (_hoveredSegmentIndex != -1 || _hoveredSeparatorIndex != -1)
        {
            _hoveredSegmentIndex   = -1;
            _hoveredSeparatorIndex = -1;
            needsRedraw            = true;
        }

        if (needsRedraw)
        {
            RenderPathSection();
        }
        return;
    }

    // Check Path segments and separators
    bool inSection2  = inClient && PtInRect(&_sectionPathRect, pt);
    bool needsRedraw = false;

    if (inSection2)
    {
        float localX         = static_cast<float>(pt.x - _sectionPathRect.left);
        float localY         = static_cast<float>(pt.y - _sectionPathRect.top);
        D2D1_POINT_2F movePt = D2D1::Point2F(localX, localY);

        // Check segment hovers - use _hoveredSegmentIndex for consistent pattern
        int newHoveredSegment = -1;
        for (size_t i = 0; i < _segments.size(); ++i)
        {
            const auto& bounds = _segments[i].bounds;
            if (bounds.left <= movePt.x && movePt.x <= bounds.right && bounds.top <= movePt.y && movePt.y <= bounds.bottom)
            {
                newHoveredSegment = static_cast<int>(i);
                break;
            }
        }

        if (newHoveredSegment != _hoveredSegmentIndex)
        {
            _hoveredSegmentIndex = newHoveredSegment;
            needsRedraw          = true;
        }

        // Check separator hovers
        int oldHoveredSeparator = _hoveredSeparatorIndex;
        _hoveredSeparatorIndex  = -1;
        for (size_t i = 0; i < _separators.size(); ++i)
        {
            const auto& bounds = _separators[i].bounds;
            if (bounds.left <= movePt.x && movePt.x <= bounds.right && bounds.top <= movePt.y && movePt.y <= bounds.bottom)
            {
                _hoveredSeparatorIndex = static_cast<int>(i);
                break;
            }
        }
        if (oldHoveredSeparator != _hoveredSeparatorIndex)
            needsRedraw = true;
    }
    else
    {
        // Clear all Section 2 hover states when outside
        if (_hoveredSegmentIndex != -1 || _hoveredSeparatorIndex != -1)
        {
            _hoveredSegmentIndex   = -1;
            _hoveredSeparatorIndex = -1;
            needsRedraw            = true;
        }
    }

    if (needsRedraw)
        RenderPathSection();

    if (_menuOpenForSeparator != -1 && _pendingSeparatorMenuSwitchIndex == -1 && _hoveredSeparatorIndex != -1)
    {
        const int targetIndex = _hoveredSeparatorIndex;
        if (targetIndex != _menuOpenForSeparator)
        {
            const size_t separatorIndex = static_cast<size_t>(targetIndex);

            const bool separatorValid = separatorIndex < _separators.size();
            bool eligibleForSiblings  = false;

            if (separatorValid)
            {
                const auto& separator = _separators[separatorIndex];
                if (separator.leftSegmentIndex < _segments.size() && separator.rightSegmentIndex < _segments.size())
                {
                    const auto& leftSegment  = _segments[separator.leftSegmentIndex];
                    const auto& rightSegment = _segments[separator.rightSegmentIndex];
                    eligibleForSiblings      = ! leftSegment.isEllipsis && ! rightSegment.isEllipsis;
                }
            }

            if (eligibleForSiblings)
            {
                _pendingSeparatorMenuSwitchIndex = targetIndex;
                SendMessageW(_hWnd.get(), WM_CANCELMODE, 0, 0);
                PostMessageW(_hWnd.get(), WndMsg::kNavigationMenuShowSiblingsDropdown, static_cast<WPARAM>(separatorIndex), 0);
            }
        }
    }
}

void NavigationView::OnEnterMenuLoop([[maybe_unused]] bool isTrackPopupMenu)
{
    _inMenuLoop = true;
    UpdateHoverTimerState();
}

void NavigationView::OnExitMenuLoop([[maybe_unused]] bool isShortcut)
{
    _inMenuLoop = false;

    // Clear menu state and reverse rotation animation
    if (_menuOpenForSeparator != -1)
    {
        _pendingSeparatorMenuSwitchIndex = -1;
        StartSeparatorAnimation(static_cast<size_t>(_menuOpenForSeparator), 0.0f);
        _menuOpenForSeparator = -1;
        _activeSeparatorIndex = -1;

        // Invalidate Section 2 for visual update
        RenderPathSection();
    }

    if (_requestFolderViewFocusCallback && _hWnd)
    {
        const HWND root = GetAncestor(_hWnd.get(), GA_ROOT);
        if (root && GetActiveWindow() == root)
        {
            _requestFolderViewFocusCallback();
        }
    }

    UpdateHoverTimerState();
}

void NavigationView::OnSetFocus()
{
    if (_hWnd)
    {
        PostMessageW(GetParent(_hWnd.get()), WndMsg::kPaneFocusChanged, 0, 0);
    }
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::OnKillFocus(HWND newFocus)
{
    if (_hWnd)
    {
        PostMessageW(GetParent(_hWnd.get()), WndMsg::kPaneFocusChanged, 0, 0);
    }
    if (_pathEdit && newFocus == _pathEdit.get())
    {
        return;
    }

    if (_editMode)
    {
        ExitEditMode(false);
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

bool NavigationView::OnKeyDown(WPARAM key)
{
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

    if (key == VK_ESCAPE)
    {
        if (_editMode)
        {
            ExitEditMode(false);
        }

        if (_requestFolderViewFocusCallback)
        {
            _requestFolderViewFocusCallback();
        }
        return true;
    }

    if (key == VK_TAB)
    {
        if (_editMode)
        {
            ExitEditMode(false);
        }

        MoveFocus(! shift);
        return true;
    }

    if (key == VK_RETURN || key == VK_SPACE)
    {
        ActivateFocusedRegion();
        return true;
    }

    return false;
}

void NavigationView::MoveFocus(bool forward)
{
    std::array<FocusRegion, 4> order{};
    size_t count = 0;
    if (_showMenuSection)
    {
        order[count++] = FocusRegion::Menu;
    }
    order[count++] = FocusRegion::Path;
    order[count++] = FocusRegion::History;
    if (_showDiskInfoSection)
    {
        order[count++] = FocusRegion::DiskInfo;
    }

    if (count == 0)
    {
        return;
    }

    size_t index = 0;
    bool found   = false;
    for (size_t i = 0; i < count; ++i)
    {
        if (order[i] == _focusedRegion)
        {
            index = i;
            found = true;
            break;
        }
    }

    if (! found)
    {
        _focusedRegion = order[0];
    }
    else if (forward)
    {
        if (index + 1 < count)
        {
            _focusedRegion = order[index + 1];
        }
        else
        {
            if (_requestFolderViewFocusCallback)
            {
                _requestFolderViewFocusCallback();
            }
            return;
        }
    }
    else
    {
        if (index == 0)
        {
            if (_requestFolderViewFocusCallback)
            {
                _requestFolderViewFocusCallback();
            }
            return;
        }

        _focusedRegion = order[index - 1];
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::ActivateFocusedRegion()
{
    if (_editMode)
    {
        return;
    }

    NormalizeFocusRegion();
    switch (_focusedRegion)
    {
        case FocusRegion::Menu: ShowMenuDropdown(); break;
        case FocusRegion::Path: EnterEditMode(); break;
        case FocusRegion::History: ShowHistoryDropdown(); break;
        case FocusRegion::DiskInfo: ShowDiskInfoDropdown(); break;
    }
}

void NavigationView::NormalizeFocusRegion()
{
    if (! _showMenuSection && _focusedRegion == FocusRegion::Menu)
    {
        _focusedRegion = FocusRegion::Path;
    }

    if (! _showDiskInfoSection && _focusedRegion == FocusRegion::DiskInfo)
    {
        _focusedRegion = FocusRegion::History;
    }
}
