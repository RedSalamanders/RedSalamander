#include "NavigationViewInternal.h"

#include <windowsx.h>

#include "DirectoryInfoCache.h"
#include "Helpers.h"
#include "IconCache.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "resource.h"

void NavigationView::TraceNavigationInputState(std::wstring_view eventName, std::optional<POINT> clientPoint) const noexcept
{
    if (! RedSalamander::DxUi::IsContextMenuDiagnosticsEnabled())
    {
        return;
    }

    const HWND hwnd = _hWnd.get();

    POINT screenPt{};
    POINT resolvedClientPt{};
    bool haveScreenPt = false;
    bool haveClientPt = false;

    if (clientPoint.has_value())
    {
        resolvedClientPt = clientPoint.value();
        haveClientPt     = true;
        screenPt         = resolvedClientPt;
        if (hwnd && ClientToScreen(hwnd, &screenPt) != FALSE)
        {
            haveScreenPt = true;
        }
    }
    RECT clientRect{};
    const bool haveClientRect = hwnd && GetClientRect(hwnd, &clientRect) != FALSE;
    const bool inClient       = haveClientPt && haveClientRect && PtInRect(&clientRect, resolvedClientPt) != 0;

    const HWND windowAtPoint = haveScreenPt ? WindowFromPoint(screenPt) : nullptr;
    const HWND root          = hwnd ? GetAncestor(hwnd, GA_ROOT) : nullptr;
    const HWND pathEditHost  = (_pathEdit && _pathEdit->hwnd) ? _pathEdit->hwnd.get() : nullptr;
    const HWND pathEditInput = _pathEdit ? _pathEdit->GetTextInputHwnd() : nullptr;

    TraceNavigationViewMenuDiagnostics(L"navigation.input-state",
                                       L"event={} hwnd={:#x} root={:#x} clientPt=({}, {}) haveClient={} screenPt=({}, {}) haveScreen={} inClient={} "
                                       L"windowAtPoint={:#x} focus={:#x} active={:#x} foreground={:#x} capture={:#x} "
                                       L"editMode={} fullPathPopup={:#x} fullPathEditMode={} pathEditHost={:#x} pathEditInput={:#x} "
                                       L"inMenuLoop={} hoverTimer={} trackingMouse={} embedded={} "
                                       L"hoverMenu={} hoverHistory={} hoverDisk={} hoveredSegment={} hoveredSeparator={} "
                                       L"menuPressed={} menuOpenSeparator={} pendingSeparator={} "
                                       L"rectDrive=({}, {}, {}, {}) rectPath=({}, {}, {}, {}) rectHistory=({}, {}, {}, {}) rectDisk=({}, {}, {}, {})",
                                       eventName,
                                       reinterpret_cast<uintptr_t>(hwnd),
                                       reinterpret_cast<uintptr_t>(root),
                                       resolvedClientPt.x,
                                       resolvedClientPt.y,
                                       haveClientPt ? 1 : 0,
                                       screenPt.x,
                                       screenPt.y,
                                       haveScreenPt ? 1 : 0,
                                       inClient ? 1 : 0,
                                       reinterpret_cast<uintptr_t>(windowAtPoint),
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                       reinterpret_cast<uintptr_t>(GetCapture()),
                                       _editMode ? 1 : 0,
                                       reinterpret_cast<uintptr_t>(_fullPathPopup.get()),
                                       _fullPathPopupEditMode ? 1 : 0,
                                       reinterpret_cast<uintptr_t>(pathEditHost),
                                       reinterpret_cast<uintptr_t>(pathEditInput),
                                       _inMenuLoop ? 1 : 0,
                                       static_cast<unsigned long long>(_hoverTimer),
                                       _trackingMouse ? 1 : 0,
                                       _embeddedDestinationMode ? 1 : 0,
                                       _menuButtonHovered ? 1 : 0,
                                       _historyButtonHovered ? 1 : 0,
                                       _diskInfoHovered ? 1 : 0,
                                       _hoveredSegmentIndex,
                                       _hoveredSeparatorIndex,
                                       _menuButtonPressed ? 1 : 0,
                                       _menuOpenForSeparator,
                                       _pendingSeparatorMenuSwitchIndex,
                                       _sectionDriveRect.left,
                                       _sectionDriveRect.top,
                                       _sectionDriveRect.right,
                                       _sectionDriveRect.bottom,
                                       _sectionPathRect.left,
                                       _sectionPathRect.top,
                                       _sectionPathRect.right,
                                       _sectionPathRect.bottom,
                                       _sectionHistoryRect.left,
                                       _sectionHistoryRect.top,
                                       _sectionHistoryRect.right,
                                       _sectionHistoryRect.bottom,
                                       _sectionDiskInfoRect.left,
                                       _sectionDiskInfoRect.top,
                                       _sectionDiskInfoRect.right,
                                       _sectionDiskInfoRect.bottom);
}

void NavigationView::ClearHoverState(std::wstring_view source, bool renderChanges) noexcept
{
    bool pathChanged = false;

    if (_menuButtonHovered)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-clear",
                                           L"hwnd={:#x} source={} target=menu",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           source);
        _menuButtonHovered = false;
        if (renderChanges)
        {
            RenderDriveSection();
        }
    }

    if (_historyButtonHovered)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-clear",
                                           L"hwnd={:#x} source={} target=history",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           source);
        _historyButtonHovered = false;
        if (renderChanges)
        {
            RenderHistorySection();
        }
    }

    if (_diskInfoHovered)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-clear",
                                           L"hwnd={:#x} source={} target=disk",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           source);
        _diskInfoHovered = false;
        if (renderChanges)
        {
            RenderDiskInfoSection();
        }
    }

    if (_hoveredSegmentIndex != -1 || _hoveredSeparatorIndex != -1)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-clear",
                                           L"hwnd={:#x} source={} target=path segment={} separator={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           source,
                                           _hoveredSegmentIndex,
                                           _hoveredSeparatorIndex);
        _hoveredSegmentIndex   = -1;
        _hoveredSeparatorIndex = -1;
        pathChanged            = true;
    }

    if (pathChanged && renderChanges)
    {
        RenderPathSection();
    }
}

void NavigationView::OnLButtonDown(const RedSalamander::DxUi::PointerInputEvent& event)
{
    if (! ShouldAcceptPointerEvent(event))
    {
        return;
    }

    OnLButtonDown(event.clientPointPx);
}

void NavigationView::OnLButtonDown(POINT pt)
{
    TraceNavigationInputState(L"lbutton-down", pt);
    TraceNavigationViewMenuDiagnostics(L"navigation.lbutton-down",
                                       L"hwnd={:#x} pt=({}, {}) editMode={} embedded={} historyRect=({}, {}, {}, {}) historyCount={} focus={:#x} active={:#x} capture={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       pt.x,
                                       pt.y,
                                       _editMode ? 1 : 0,
                                       _embeddedDestinationMode ? 1 : 0,
                                       _sectionHistoryRect.left,
                                       _sectionHistoryRect.top,
                                       _sectionHistoryRect.right,
                                       _sectionHistoryRect.bottom,
                                       _pathHistory.size(),
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetCapture()));
    if (_editMode)
    {
        const bool clickOnNavigationButton = (_showMenuSection && PtInRect(&_sectionDriveRect, pt)) || PtInRect(&_sectionHistoryRect, pt) ||
                                             (_showDiskInfoSection && PtInRect(&_sectionDiskInfoRect, pt));
        if (! clickOnNavigationButton)
        {
            TraceNavigationViewMenuDiagnostics(L"navigation.lbutton-down.ignored",
                                               L"hwnd={:#x} reason=edit-mode pt=({}, {})",
                                               reinterpret_cast<uintptr_t>(_hWnd.get()),
                                               pt.x,
                                               pt.y);
            return;
        }

        TraceNavigationViewMenuDiagnostics(L"navigation.lbutton-down.exit-edit",
                                           L"hwnd={:#x} pt=({}, {})",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           pt.x,
                                           pt.y);
        ExitEditMode(false, L"nav-button-click");
    }

    // Check Section 1 (menu button) click
    if (_showMenuSection && PtInRect(&_sectionDriveRect, pt))
    {
        RequestOwnerPaneFocus();
        _focusedRegion = FocusRegion::Menu;
        TraceNavigationViewMenuDiagnostics(L"navigation.menu-click",
                                           L"hwnd={:#x} pt=({}, {}) embedded={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           pt.x,
                                           pt.y,
                                           _embeddedDestinationMode ? 1 : 0);
        ShowMenuDropdown(true);
        return;
    }

    // Check history button click (Section 2 area)
    if (PtInRect(&_sectionHistoryRect, pt))
    {
        RequestOwnerPaneFocus();
        _focusedRegion = FocusRegion::History;
        TraceNavigationViewMenuDiagnostics(L"navigation.history-click",
                                           L"hwnd={:#x} pt=({}, {}) historyCount={} embedded={} focus={:#x} active={:#x}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           pt.x,
                                           pt.y,
                                           _pathHistory.size(),
                                           _embeddedDestinationMode ? 1 : 0,
                                           reinterpret_cast<uintptr_t>(GetFocus()),
                                           reinterpret_cast<uintptr_t>(GetActiveWindow()));
        ShowHistoryDropdown(true);
        return;
    }

    // Check Section 3 (disk info) click
    if (_showDiskInfoSection && PtInRect(&_sectionDiskInfoRect, pt))
    {
        RequestOwnerPaneFocus();
        _focusedRegion = FocusRegion::DiskInfo;
        TraceNavigationViewMenuDiagnostics(L"navigation.disk-click",
                                           L"hwnd={:#x} pt=({}, {}) embedded={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           pt.x,
                                           pt.y,
                                           _embeddedDestinationMode ? 1 : 0);
        ShowDiskInfoDropdown(true);
        return;
    }

    // Check if click is in Section 2 (breadcrumbs)
    if (! PtInRect(&_sectionPathRect, pt))
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.lbutton-down.ignored",
                                           L"hwnd={:#x} reason=outside-path pt=({}, {}) pathRect=({}, {}, {}, {})",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           pt.x,
                                           pt.y,
                                           _sectionPathRect.left,
                                           _sectionPathRect.top,
                                           _sectionPathRect.right,
                                           _sectionPathRect.bottom);
        return;
    }

    _focusedRegion = FocusRegion::Path;
    if (! _embeddedDestinationMode)
    {
        RequestOwnerPaneFocus();
    }

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
            RequestOwnerPaneFocus();
            TraceNavigationViewMenuDiagnostics(L"navigation.segment-click",
                                               L"hwnd={:#x} pt=({}, {}) index={} last={} ellipsis={} local=({:.1f}, {:.1f})",
                                               reinterpret_cast<uintptr_t>(_hWnd.get()),
                                               pt.x,
                                               pt.y,
                                               i,
                                               (i + 1u == _segments.size()) ? 1 : 0,
                                               segment.isEllipsis ? 1 : 0,
                                               clickPt.x,
                                               clickPt.y);
            if (segment.isEllipsis)
            {
                RequestFullPathPopup(segment.bounds);
                return;
            }

            if (i + 1u == _segments.size())
            {
                return;
            }

            // Navigate to this ancestor segment's path.
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
            RequestOwnerPaneFocus();
            const bool adjacentToEllipsis = (separator.leftSegmentIndex < _segments.size() && _segments[separator.leftSegmentIndex].isEllipsis) ||
                                            (separator.rightSegmentIndex < _segments.size() && _segments[separator.rightSegmentIndex].isEllipsis);
            if (adjacentToEllipsis)
            {
                RequestFullPathPopup(bounds);
                return;
            }

            // If a different separator menu is open, close it first
            if (_hWnd)
            {
                if (_menuOpenForSeparator != -1 && _menuOpenForSeparator != static_cast<int>(i))
                {
                    SendMessageW(_hWnd.get(), WM_CANCELMODE, 0, 0);
                }
                TraceNavigationViewMenuDiagnostics(L"navigation.siblings-click",
                                                   L"hwnd={:#x} pt=({}, {}) index={} local=({:.1f}, {:.1f}) menuOpenForSeparator={} direct=1",
                                                   reinterpret_cast<uintptr_t>(_hWnd.get()),
                                                   pt.x,
                                                   pt.y,
                                                   i,
                                                   clickPt.x,
                                                   clickPt.y,
                                                   _menuOpenForSeparator);
                _pendingSeparatorMenuSwitchIndex = -1;
                ShowSiblingsDropdown(i);
            }
            return;
        }
    }

    TraceNavigationViewMenuDiagnostics(L"navigation.lbutton-down.ignored",
                                       L"hwnd={:#x} reason=no-breadcrumb-hit pt=({}, {}) local=({:.1f}, {:.1f}) segments={} separators={}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       pt.x,
                                       pt.y,
                                       clickPt.x,
                                       clickPt.y,
                                       _segments.size(),
                                       _separators.size());
}

void NavigationView::OnLButtonDblClk(const RedSalamander::DxUi::PointerInputEvent& event)
{
    if (! ShouldAcceptPointerEvent(event))
    {
        return;
    }

    OnLButtonDblClk(event.clientPointPx);
}

void NavigationView::OnLButtonDblClk(POINT pt)
{
    TraceNavigationInputState(L"lbutton-dblclk", pt);
    if (_editMode)
        return;

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
#ifdef ENABLE_TESTS
        ++_debugDoubleClickActivateCount;
        _debugLastDoubleClickOnLastSegment = onLastSegment;
        _debugLastDoubleClickInWhitespace  = inWhitespace;
        _debugLastDoubleClickPoint         = pt;
        _debugLastDoubleClickLocalX        = clickPt.x;
        _debugLastDoubleClickLocalY        = clickPt.y;
#endif
        RequestOwnerPaneFocus();
        EnterEditMode();
    }
}

void NavigationView::OnMouseMove(const RedSalamander::DxUi::PointerInputEvent& event)
{
    if (! ShouldAcceptPointerEvent(event))
    {
        return;
    }

    OnMouseMove(event.clientPointPx);
}

void NavigationView::OnMouseMove(POINT pt)
{
    TraceNavigationInputState(L"mouse-move.enter", pt);

    if (_fullPathPopup)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.mouse-move.ignored",
                                           L"hwnd={:#x} reason=full-path-popup popup={:#x}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           reinterpret_cast<uintptr_t>(_fullPathPopup.get()));
        return;
    }

    // Track mouse for leave notification
    if (! _trackingMouse)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = _hWnd.get();
        const BOOL tracked = TrackMouseEvent(&tme);
        const DWORD error  = tracked == FALSE ? GetLastError() : ERROR_SUCCESS;
        _trackingMouse     = tracked != FALSE;
        TraceNavigationViewMenuDiagnostics(L"navigation.track-mouse",
                                           L"hwnd={:#x} requested=leave result={} lastError={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           tracked != FALSE ? 1 : 0,
                                           error);
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

    TraceNavigationViewMenuDiagnostics(L"navigation.hover-candidates",
                                       L"hwnd={:#x} pt=({}, {}) candidateMenu={} candidateHistory={} candidateDisk={} currentMenu={} currentHistory={} currentDisk={} currentSegment={} currentSeparator={}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       pt.x,
                                       pt.y,
                                       menuButtonHovered ? 1 : 0,
                                       historyButtonHovered ? 1 : 0,
                                       diskInfoHovered ? 1 : 0,
                                       _menuButtonHovered ? 1 : 0,
                                       _historyButtonHovered ? 1 : 0,
                                       _diskInfoHovered ? 1 : 0,
                                       _hoveredSegmentIndex,
                                       _hoveredSeparatorIndex);

    // Update Section 1 if hover state changed
    if (menuButtonHovered != _menuButtonHovered)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-change",
                                           L"hwnd={:#x} target=menu old={} new={} render=drive",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           _menuButtonHovered ? 1 : 0,
                                           menuButtonHovered ? 1 : 0);
        _menuButtonHovered = menuButtonHovered;
        RenderDriveSection();
    }

    // Update history button if hover state changed
    if (historyButtonHovered != _historyButtonHovered)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-change",
                                           L"hwnd={:#x} target=history old={} new={} render=history",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           _historyButtonHovered ? 1 : 0,
                                           historyButtonHovered ? 1 : 0);
        _historyButtonHovered = historyButtonHovered;
        RenderHistorySection(); // History button is rendered in Section 3
    }

    // Update Section 3 if hover state changed
    if (diskInfoHovered != _diskInfoHovered)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-change",
                                           L"hwnd={:#x} target=disk old={} new={} render=disk",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           _diskInfoHovered ? 1 : 0,
                                           diskInfoHovered ? 1 : 0);
        _diskInfoHovered = diskInfoHovered;
        RenderDiskInfoSection();
    }

    if (_editMode)
    {
        if (_hoveredSegmentIndex != -1 || _hoveredSeparatorIndex != -1)
        {
            TraceNavigationViewMenuDiagnostics(L"navigation.hover-clear",
                                               L"hwnd={:#x} source=mouse-move-edit-mode target=path segment={} separator={}",
                                               reinterpret_cast<uintptr_t>(_hWnd.get()),
                                               _hoveredSegmentIndex,
                                               _hoveredSeparatorIndex);
            _hoveredSegmentIndex   = -1;
            _hoveredSeparatorIndex = -1;
            RenderPathSection();
        }

        TraceNavigationInputState(L"mouse-move.exit", pt);
        return;
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
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-change",
                                           L"hwnd={:#x} target=segment old={} new={} local=({}, {})",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           _hoveredSegmentIndex,
                                           newHoveredSegment,
                                           localX,
                                           localY);
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
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-change",
                                           L"hwnd={:#x} target=separator old={} new={} local=({}, {})",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           oldHoveredSeparator,
                                           _hoveredSeparatorIndex,
                                           localX,
                                           localY);
        needsRedraw = true;
    }

    if (needsRedraw)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-render",
                                           L"hwnd={:#x} source=mouse-move target=path segment={} separator={}",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           _hoveredSegmentIndex,
                                           _hoveredSeparatorIndex);
        RenderPathSection();
    }

    TraceNavigationInputState(L"mouse-move.exit", pt);
}

void NavigationView::OnMouseLeave()
{
    TraceNavigationInputState(L"mouse-leave.enter");
    _trackingMouse = false;
    ClearHoverState(L"mouse-leave", true);
    TraceNavigationInputState(L"mouse-leave.exit");
}

void NavigationView::OnSetCursor([[maybe_unused]] HWND hwnd, UINT hitTest, [[maybe_unused]] UINT mouseMsg)
{
    TraceNavigationInputState(L"set-cursor");
    if (hitTest == HTCLIENT)
    {
        SetCursor(LoadCursor(nullptr, IDC_ARROW));
    }
}

void NavigationView::OnTimer(UINT_PTR timerId)
{
    if (timerId != HOVER_TIMER_ID || _hoverTimer == 0)
    {
        return;
    }

    TraceNavigationInputState(L"hover-timer.tick");
    TraceNavigationViewMenuDiagnostics(L"navigation.hover-timer.noop",
                                       L"hwnd={:#x} editMode={} inMenuLoop={} fullPathPopup={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       _editMode ? 1 : 0,
                                       _inMenuLoop ? 1 : 0,
                                       reinterpret_cast<uintptr_t>(_fullPathPopup.get()));
}

void NavigationView::OnEnterMenuLoop([[maybe_unused]] bool isTrackPopupMenu)
{
    TraceNavigationInputState(L"enter-menu-loop.before");
    TraceNavigationViewMenuDiagnostics(L"navigation.enter-menu-loop",
                                       L"hwnd={:#x} isTrackPopupMenu={} menuOpenSeparator={} pendingSeparator={} focus={:#x} active={:#x} foreground={:#x} capture={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       isTrackPopupMenu ? 1 : 0,
                                       _menuOpenForSeparator,
                                       _pendingSeparatorMenuSwitchIndex,
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                       reinterpret_cast<uintptr_t>(GetCapture()));
    _inMenuLoop = true;
    UpdateHoverTimerState();
    TraceNavigationInputState(L"enter-menu-loop.after");
}

void NavigationView::OnExitMenuLoop([[maybe_unused]] bool isShortcut)
{
    TraceNavigationInputState(L"exit-menu-loop.before");
    TraceNavigationViewMenuDiagnostics(L"navigation.exit-menu-loop",
                                       L"hwnd={:#x} isShortcut={} menuOpenSeparator={} pendingSeparator={} focus={:#x} active={:#x} foreground={:#x} capture={:#x}",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       isShortcut ? 1 : 0,
                                       _menuOpenForSeparator,
                                       _pendingSeparatorMenuSwitchIndex,
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                       reinterpret_cast<uintptr_t>(GetCapture()));
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
    TraceNavigationInputState(L"exit-menu-loop.after");
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
    if (_pathEdit && _pathEdit->hwnd &&
        (newFocus == _pathEdit->hwnd.get() || newFocus == _pathEdit->GetTextInputHwnd() || IsChild(_pathEdit->hwnd.get(), newFocus) != FALSE))
    {
        return;
    }
    if (IsEditValidationPopupWindow(newFocus))
    {
        return;
    }

    if (_editMode)
    {
        ExitEditMode(false, L"navigation-kill-focus");
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
            ExitEditMode(false, L"escape-key");
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
            ExitEditMode(false, L"tab-key");
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
        case FocusRegion::Menu:
            if (_hWnd)
            {
                PostMessageW(_hWnd.get(), WndMsg::kNavigationViewShowMenuDropdown, 1, 0);
            }
            break;
        case FocusRegion::Path:
#ifdef ENABLE_TESTS
            ++_debugKeyboardActivateCount;
#endif
            EnterEditMode();
            break;
        case FocusRegion::History:
            if (_hWnd)
            {
                PostMessageW(_hWnd.get(), WndMsg::kNavigationViewShowHistoryDropdown, 1, 0);
            }
            break;
        case FocusRegion::DiskInfo:
            if (_hWnd)
            {
                PostMessageW(_hWnd.get(), WndMsg::kNavigationViewShowDiskInfoDropdown, 1, 0);
            }
            break;
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
