#include "FolderViewInternal.h"

#include <cwctype>
#include <limits>

void FolderView::OnMouseWheelMessage(UINT keyState, int delta)
{
    const bool horizontal = (keyState & MK_SHIFT) != 0;
    if (horizontal)
    {
        delta = -delta;
    }

    OnMouseWheel(delta, horizontal);
}

void FolderView::OnMouseLeave()
{
    if (_hoveredIndex != static_cast<size_t>(-1) && _hoveredIndex < _items.size())
    {
        RECT rc = ToPixelRect(OffsetRect(_items[_hoveredIndex].bounds, -_horizontalOffset, -_scrollOffset), _dpi);
        InvalidateRect(_hWnd.get(), &rc, FALSE);
    }
    _hoveredIndex = static_cast<size_t>(-1);

    bool hasOverlay = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        hasOverlay = _errorOverlay.has_value();
    }

    if (hasOverlay && _alertOverlay)
    {
        _alertOverlay->ClearHotState();
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
    }
}

void FolderView::OnKeyDownMessage(WPARAM key)
{
    const bool ctrl  = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    OnKeyDown(key, ctrl, shift);
}

bool FolderView::OnSysKeyDownMessage(WPARAM key)
{
    if (key != 'D' && key != VK_DOWN && key != VK_UP)
    {
        return false;
    }

    const bool ctrl  = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    OnKeyDown(key, ctrl, shift);
    return true;
}

LRESULT FolderView::OnSetFocusMessage() noexcept
{
    if (! _hWnd)
    {
        return 0;
    }

    const HWND parent = GetParent(_hWnd.get());
    if (parent)
    {
        PostMessageW(parent, WndMsg::kPaneFocusChanged, 0, 0);
    }
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
    return 0;
}

LRESULT FolderView::OnKillFocusMessage() noexcept
{
    ExitIncrementalSearch();

    if (! _hWnd)
    {
        return 0;
    }

    const HWND parent = GetParent(_hWnd.get());
    if (parent)
    {
        PostMessageW(parent, WndMsg::kPaneFocusChanged, 0, 0);
    }
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
    return 0;
}

void FolderView::OnContextMenuMessage(HWND hwnd, LPARAM lParam)
{
    SetFocus(hwnd);

    POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    if (pt.x == -1 && pt.y == -1)
    {
        RECT rc{};
        GetClientRect(hwnd, &rc);
        pt.x = (rc.left + rc.right) / 2;
        pt.y = (rc.top + rc.bottom) / 2;
        ClientToScreen(hwnd, &pt);
    }

    OnContextMenu(pt);
}

void FolderView::OnHScrollMessage(UINT scrollRequest)
{
    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_ALL;
    GetScrollInfo(_hWnd.get(), SB_HORZ, &si);

    const float columnStride = _tileWidthDip + kColumnSpacingDip;
    int newPos               = si.nPos;

    const float viewWidthDip = DipFromPx(_clientSize.cx);
    const int visibleColumns = std::max(1, static_cast<int>(std::floor(viewWidthDip / columnStride)));

    switch (scrollRequest)
    {
        case SB_LINELEFT: newPos -= PxFromDip(columnStride); break;
        case SB_LINERIGHT: newPos += PxFromDip(columnStride); break;
        case SB_PAGELEFT: newPos -= PxFromDip(static_cast<float>(visibleColumns) * columnStride); break;
        case SB_PAGERIGHT: newPos += PxFromDip(static_cast<float>(visibleColumns) * columnStride); break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION:
        {
            const float thumbDip    = DipFromPx(static_cast<int>(si.nTrackPos));
            const float columnIndex = std::round((thumbDip - kColumnSpacingDip) / columnStride);
            const float snappedDip  = kColumnSpacingDip + (columnIndex * columnStride);
            newPos                  = PxFromDip(snappedDip);
            break;
        }
        case SB_LEFT: newPos = si.nMin; break;
        case SB_RIGHT: newPos = si.nMax; break;
    }

    newPos            = std::clamp(newPos, si.nMin, si.nMax);
    _horizontalOffset = DipFromPx(newPos);

    const float maxHorizontalOffset = std::max(0.0f, _contentWidth - DipFromPx(_clientSize.cx));
    _horizontalOffset               = std::clamp(_horizontalOffset, 0.0f, maxHorizontalOffset);

    UpdateScrollMetrics();
    BoostIconLoadingForVisibleRange();
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void FolderView::OnCommandMessage(UINT commandId)
{
    ErrorOverlayState overlay{};
    bool hasOverlay = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (_errorOverlay)
        {
            overlay    = *_errorOverlay;
            hasOverlay = true;
        }
    }

    if (hasOverlay && overlay.blocksInput)
    {
        return;
    }

    switch (commandId)
    {
        case CmdOpen: ActivateFocusedItem(); break;
        case CmdOpenWith:
        {
            if (_focusedIndex == static_cast<size_t>(-1) || _focusedIndex >= _items.size())
            {
                break;
            }

            const std::filesystem::path fullPath = GetItemFullPath(_items[_focusedIndex]);
            ShellExecuteW(
                _hWnd.get(), L"open", L"rundll32.exe", std::format(L"shell32.dll,OpenAs_RunDLL \"{}\"", fullPath.c_str()).c_str(), nullptr, SW_SHOWNORMAL);
            break;
        }
        case CmdViewSpace:
        {
            if (_hWnd)
            {
                SetFocus(_hWnd.get());
                const HWND root = GetAncestor(_hWnd.get(), GA_ROOT);
                if (root && PostMessageW(root, WM_COMMAND, MAKEWPARAM(IDM_PANE_VIEW_SPACE, 0), 0) != 0)
                {
                    break;
                }
            }
            break;
        }
        case CmdDelete: CommandDelete(); break;
        case CmdRename: RenameFocusedItem(); break;
        case CmdCopy: CopySelectionToClipboard(); break;
        case CmdPaste: PasteItemsFromClipboard(); break;
        case CmdSelectAll: SelectAll(); break;
        case CmdUnselectAll: ClearSelection(); break;
        case CmdProperties: ShowProperties(); break;
        case CmdMove: MoveSelectedItems(); break;
        case CmdOverlaySampleError:
            if (IsOverlaySampleEnabled())
            {
                DebugShowOverlaySample(OverlaySeverity::Error);
            }
            break;
        case CmdOverlaySampleWarning:
            if (IsOverlaySampleEnabled())
            {
                DebugShowOverlaySample(OverlaySeverity::Warning);
            }
            break;
        case CmdOverlaySampleInformation:
            if (IsOverlaySampleEnabled())
            {
                DebugShowOverlaySample(OverlaySeverity::Information);
            }
            break;
        case CmdOverlaySampleBusy:
            if (IsOverlaySampleEnabled())
            {
                DebugShowOverlaySample(OverlaySeverity::Busy);
            }
            break;
        case CmdOverlaySampleHide:
            if (IsOverlaySampleEnabled())
            {
                DebugHideOverlaySample();
            }
            break;
        case CmdOverlaySampleErrorNonModal:
            if (IsOverlaySampleEnabled())
            {
                DebugShowOverlaySample(ErrorOverlayKind::Operation, OverlaySeverity::Error, false);
            }
            break;
        case CmdOverlaySampleWarningNonModal:
            if (IsOverlaySampleEnabled())
            {
                DebugShowOverlaySample(ErrorOverlayKind::Operation, OverlaySeverity::Warning, false);
            }
            break;
        case CmdOverlaySampleInformationNonModal:
            if (IsOverlaySampleEnabled())
            {
                DebugShowOverlaySample(ErrorOverlayKind::Operation, OverlaySeverity::Information, false);
            }
            break;
        case CmdOverlaySampleCanceled:
            if (IsOverlaySampleEnabled())
            {
                DebugShowCanceledOverlaySample();
            }
            break;
        case CmdOverlaySampleBusyWithCancel:
            if (IsOverlaySampleEnabled())
            {
                DebugShowOverlaySample(ErrorOverlayKind::Enumeration, OverlaySeverity::Busy, true);
            }
            break;
    }
}

void FolderView::OnMouseWheel(int delta, bool horizontal)
{
    const float columnStride = _tileWidthDip + kColumnSpacingDip;

    if (horizontal)
    {
        // Horizontal wheel: scroll by full columns
        const int wheelClicks   = delta / WHEEL_DELTA;
        const float deltaOffset = static_cast<float>(wheelClicks) * columnStride;

        _horizontalOffset += deltaOffset;
    }
    else
    {
        // Vertical wheel: treat as horizontal scroll in horizontal-only layout
        const int wheelClicks   = -delta / WHEEL_DELTA; // Invert for natural scrolling
        const float deltaOffset = static_cast<float>(wheelClicks) * columnStride;

        _horizontalOffset += deltaOffset;
    }

    // Clamp to valid range
    const float maxHorizontal = std::max(0.0f, _contentWidth - DipFromPx(_clientSize.cx));
    _horizontalOffset         = std::clamp(_horizontalOffset, 0.0f, maxHorizontal);

    // Snap to nearest column boundary for crisp alignment
    const float currentColumnIndex = std::round((_horizontalOffset - kColumnSpacingDip) / columnStride);
    _horizontalOffset              = kColumnSpacingDip + (currentColumnIndex * columnStride);

    // Clamp again after snapping
    _horizontalOffset = std::clamp(_horizontalOffset, 0.0f, maxHorizontal);

    UpdateScrollMetrics();
    BoostIconLoadingForVisibleRange();
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void FolderView::OnLButtonDown(POINT pt, WPARAM keys)
{
    ErrorOverlayState overlay{};
    bool hasOverlay = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (_errorOverlay)
        {
            overlay    = *_errorOverlay;
            hasOverlay = true;
        }
    }

    if (hasOverlay && _alertOverlay)
    {
        constexpr uint32_t kCancelButtonId        = 1;
        const float xDip                          = DipFromPx(pt.x);
        const float yDip                          = DipFromPx(pt.y);
        const RedSalamander::Ui::AlertHitTest hit = _alertOverlay->HitTest(D2D1::Point2F(xDip, yDip));
        if (hit.part == RedSalamander::Ui::AlertHitTest::Part::Close)
        {
            DismissAlertOverlay();
            return;
        }
        if (hit.part == RedSalamander::Ui::AlertHitTest::Part::Button)
        {
            if (hit.buttonId == kCancelButtonId && overlay.kind == ErrorOverlayKind::Enumeration && overlay.severity == OverlaySeverity::Busy)
            {
                CancelPendingEnumeration();
                const std::wstring title   = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_CANCELED);
                const std::wstring message = LoadStringResource(nullptr, IDS_OVERLAY_MSG_ENUMERATION_CANCELED);
                ShowAlertOverlay(
                    ErrorOverlayKind::Enumeration, OverlaySeverity::Information, title, message, HRESULT_FROM_WIN32(ERROR_CANCELLED), false, false);
                return;
            }
            DismissAlertOverlay();
            return;
        }

        if (overlay.blocksInput)
        {
            return;
        }
    }

    SetFocus(_hWnd.get());
    SetCapture(_hWnd.get());
    _drag.dragging   = true;
    _drag.startPoint = pt;

    auto hit = HitTest(pt);
    if (hit)
    {
        bool ctrl  = (keys & MK_CONTROL) != 0;
        bool shift = (keys & MK_SHIFT) != 0;

        if (shift)
        {
            if (_anchorIndex == static_cast<size_t>(-1))
            {
                _anchorIndex = *hit;
            }
            RangeSelect(*hit);
        }
        else if (ctrl)
        {
            ToggleSelection(*hit);
            _anchorIndex = *hit;
        }
        else
        {
            FocusItem(*hit, false);
            _anchorIndex = *hit;
        }
    }
    else
    {
        ClearSelection();
        _anchorIndex = _focusedIndex != static_cast<size_t>(-1) ? _focusedIndex : static_cast<size_t>(-1);
    }
}

void FolderView::OnLButtonDblClk(POINT pt, WPARAM /*keys*/)
{
    ExitIncrementalSearch();

    SetFocus(_hWnd.get());

    const std::optional<size_t> hit = HitTest(pt);
    if (! hit.has_value() || hit.value() >= _items.size())
    {
        return;
    }

    FocusItem(hit.value(), false);
    _anchorIndex = hit.value();
    ActivateFocusedItem();
}

void FolderView::OnLButtonUp(POINT /*pt*/)
{
    ReleaseCapture();
    _drag.dragging = false;
}

void FolderView::OnMouseMove(POINT pt, WPARAM keys)
{
    bool hasOverlay = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        hasOverlay = _errorOverlay.has_value();
    }

    if (hasOverlay && _alertOverlay)
    {
        const float xDip   = DipFromPx(pt.x);
        const float yDip   = DipFromPx(pt.y);
        const bool changed = _alertOverlay->UpdateHotState(D2D1::Point2F(xDip, yDip));
        if (changed && _hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }

        if (_hWnd)
        {
            TRACKMOUSEEVENT tme{};
            tme.cbSize    = sizeof(tme);
            tme.dwFlags   = TME_LEAVE;
            tme.hwndTrack = _hWnd.get();
            TrackMouseEvent(&tme);
        }

        return;
    }

    // Track hover state
    const std::optional<size_t> hitResult = HitTest(pt);
    const size_t newHoveredIndex          = hitResult.value_or(static_cast<size_t>(-1));

    if (newHoveredIndex != _hoveredIndex)
    {
        // Invalidate old hovered item
        if (_hoveredIndex != static_cast<size_t>(-1) && _hoveredIndex < _items.size())
        {
            RECT rc = ToPixelRect(OffsetRect(_items[_hoveredIndex].bounds, -_horizontalOffset, -_scrollOffset), _dpi);
            InvalidateRect(_hWnd.get(), &rc, FALSE);
        }

        // Update hover index
        _hoveredIndex = newHoveredIndex;

        // Invalidate new hovered item
        if (_hoveredIndex != static_cast<size_t>(-1) && _hoveredIndex < _items.size())
        {
            RECT rc = ToPixelRect(OffsetRect(_items[_hoveredIndex].bounds, -_horizontalOffset, -_scrollOffset), _dpi);
            InvalidateRect(_hWnd.get(), &rc, FALSE);

            // Track mouse leave to clear hover
            TRACKMOUSEEVENT tme{};
            tme.cbSize    = sizeof(tme);
            tme.dwFlags   = TME_LEAVE;
            tme.hwndTrack = _hWnd.get();
            TrackMouseEvent(&tme);
        }
    }

    if ((_drag.dragging) && (keys & MK_LBUTTON))
    {
        const int dx      = std::abs(pt.x - _drag.startPoint.x);
        const int dy      = std::abs(pt.y - _drag.startPoint.y);
        const int threshX = GetSystemMetrics(SM_CXDRAG);
        const int threshY = GetSystemMetrics(SM_CYDRAG);
        if (dx > threshX || dy > threshY)
        {
            BeginDragDrop();
            _drag.dragging = false;
        }
    }
}

void FolderView::OnKeyDown(WPARAM key, bool ctrl, bool shift)
{
    static_cast<void>(ctrl);

    ErrorOverlayState overlay{};
    bool hasOverlay = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        if (_errorOverlay)
        {
            overlay    = *_errorOverlay;
            hasOverlay = true;
        }
    }

    if (hasOverlay)
    {
        if (overlay.blocksInput)
        {
            if (key == VK_TAB)
            {
                if (_navigationRequestCallback)
                {
                    ExitIncrementalSearch();
                    _navigationRequestCallback(NavigationRequest::SwitchPane);
                }
                return;
            }

            if (key == VK_ESCAPE)
            {
                if (overlay.kind == ErrorOverlayKind::Enumeration && overlay.severity == OverlaySeverity::Busy)
                {
                    CancelPendingEnumeration();
                    const std::wstring title   = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_CANCELED);
                    const std::wstring message = LoadStringResource(nullptr, IDS_OVERLAY_MSG_ENUMERATION_CANCELED);
                    ShowAlertOverlay(
                        ErrorOverlayKind::Enumeration, OverlaySeverity::Information, title, message, HRESULT_FROM_WIN32(ERROR_CANCELLED), false, false);
                    return;
                }

                if (overlay.closable)
                {
                    DismissAlertOverlay();
                    return;
                }
            }
            return;
        }
    }

    auto navigateUp = [&]() noexcept
    {
        ExitIncrementalSearch();

        if (_currentFolder)
        {
            const auto isConnectionRoot = [](const std::filesystem::path& path) noexcept -> bool
            {
                std::wstring normalized = path.generic_wstring();
                if (normalized.empty())
                {
                    return false;
                }

                while (normalized.size() > 1u && normalized.back() == L'/')
                {
                    normalized.pop_back();
                }

                constexpr std::wstring_view kConnPrefix = L"/@conn:";
                if (normalized.starts_with(kConnPrefix))
                {
                    const size_t nameStart = kConnPrefix.size();
                    if (nameStart >= normalized.size())
                    {
                        return false;
                    }
                    return normalized.find(L'/', nameStart) == std::wstring::npos;
                }

                return false;
            };

            // Connection manager roots are terminal; don't navigate above them.
            if (isConnectionRoot(_currentFolder.value()))
            {
                if (_navigateUpFromRootRequestCallback)
                {
                    _navigateUpFromRootRequestCallback();
                }
                return;
            }

            const std::filesystem::path parent = _currentFolder->parent_path();
            if (! parent.empty() && parent != _currentFolder.value())
            {
                SetFolderPath(parent);
                return;
            }
        }

        if (_navigateUpFromRootRequestCallback)
        {
            _navigateUpFromRootRequestCallback();
        }
    };

    if (_incrementalSearch.active)
    {
        switch (key)
        {
            case VK_ESCAPE: ExitIncrementalSearch(); return;
            case VK_BACK:
                if (_items.empty())
                {
                    navigateUp();
                    return;
                }
                HandleIncrementalSearchBackspace();
                return;
            case VK_LEFT:
            case VK_UP: HandleIncrementalSearchNavigate(false); return;
            case VK_RIGHT:
            case VK_DOWN: HandleIncrementalSearchNavigate(true); return;
            case VK_HOME:
            case VK_END:
            case VK_PRIOR:
            case VK_NEXT:
            case VK_TAB:
            case VK_RETURN:
            case VK_DELETE:
            case VK_F2: ExitIncrementalSearch(); break;
        }
    }

    if (key == VK_TAB)
    {
        if (_navigationRequestCallback)
        {
            ExitIncrementalSearch();
            _navigationRequestCallback(NavigationRequest::SwitchPane);
        }
        return;
    }

    if (key == VK_ESCAPE)
    {
        ClearSelection();
        if (_focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
        {
            _anchorIndex = _focusedIndex;
        }
        else
        {
            _anchorIndex = static_cast<size_t>(-1);
        }
        return;
    }

    if (_items.empty())
    {
        if (key == VK_BACK)
        {
            navigateUp();
        }
        return;
    }

    const auto invalidIndex = static_cast<size_t>(-1);

    const bool hasFocus = _focusedIndex != invalidIndex && _focusedIndex < _items.size();

    // Lambda to move focus with optional range selection
    auto focusIndex = [&](size_t index) noexcept
    {
        if (shift)
        {
            if (_anchorIndex != invalidIndex)
            {
                RangeSelect(index);
            }
            else
            {
                SelectSingle(index);
                _anchorIndex = index;
            }
            return;
        }

        FocusItem(index, true);
        _anchorIndex = index;
    };

    // Lambda to convert column/row to linear index
    auto indexFromColumnRow = [&](int column, int row) noexcept -> std::optional<size_t>
    {
        // Combine early exit conditions
        if (column < 0 || row < 0 || column >= static_cast<int>(_columnCounts.size()))
        {
            return std::nullopt;
        }
        if (row >= _columnCounts[static_cast<size_t>(column)])
        {
            return std::nullopt;
        }

        // Start with row offset, accumulate columns
        size_t index = static_cast<size_t>(row);
        for (int c = 0; c < column; ++c)
        {
            index += static_cast<size_t>(_columnCounts[static_cast<size_t>(c)]);
        }

        // Single final check
        return (index < _items.size()) ? std::optional<size_t>(index) : std::nullopt;
    };

    // Lambda for horizontal navigation to reduce code duplication
    auto navigateHorizontal = [&](int columnDelta) noexcept
    {
        if (! hasFocus || _columnCounts.empty())
            return;

        const auto& focused    = _items[_focusedIndex];
        const int targetColumn = focused.column + columnDelta;

        if (targetColumn < 0 || targetColumn >= static_cast<int>(_columnCounts.size()))
            return;

        const int targetRow = std::min(focused.row, _columnCounts[static_cast<size_t>(targetColumn)] - 1);
        if (auto newIndex = indexFromColumnRow(targetColumn, targetRow))
        {
            focusIndex(*newIndex);
        }
    };

    switch (key)
    {
        case VK_HOME:
            ExitIncrementalSearch();
            if (! _items.empty())
            {
                focusIndex(0);
            }
            break;
        case VK_END:
            ExitIncrementalSearch();
            if (! _items.empty())
            {
                focusIndex(_items.size() - 1);
            }
            break;
        case VK_PRIOR: // Page Up
        {
            ExitIncrementalSearch();
            // Scroll left by visible columns
            const float columnStride = _tileWidthDip + kColumnSpacingDip;
            const float viewWidthDip = DipFromPx(_clientSize.cx);
            const int visibleColumns = std::max(1, static_cast<int>(std::floor(viewWidthDip / columnStride)));
            const float scrollDelta  = static_cast<float>(visibleColumns) * columnStride;

            _horizontalOffset -= scrollDelta;
            _horizontalOffset = std::max(0.0f, _horizontalOffset);

            // Snap to column boundary
            const float columnIndex = std::round((_horizontalOffset - kColumnSpacingDip) / columnStride);
            _horizontalOffset       = kColumnSpacingDip + (columnIndex * columnStride);
            _horizontalOffset       = std::max(0.0f, _horizontalOffset);

            // Move focus left by visible columns
            if (hasFocus && ! _columnCounts.empty())
            {
                const auto& focused    = _items[_focusedIndex];
                const int targetColumn = std::max(0, focused.column - visibleColumns);
                const int targetRow    = std::min(focused.row, _columnCounts[static_cast<size_t>(targetColumn)] - 1);

                if (auto newIndex = indexFromColumnRow(targetColumn, targetRow))
                {
                    focusIndex(*newIndex);
                }
            }

            UpdateScrollMetrics();
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
            break;
        }
        case VK_NEXT: // Page Down
        {
            ExitIncrementalSearch();
            // Scroll right by visible columns
            const float columnStride = _tileWidthDip + kColumnSpacingDip;
            const float viewWidthDip = DipFromPx(_clientSize.cx);
            const int visibleColumns = std::max(1, static_cast<int>(std::floor(viewWidthDip / columnStride)));
            const float scrollDelta  = static_cast<float>(visibleColumns) * columnStride;

            _horizontalOffset += scrollDelta;
            const float maxHorizontal = std::max(0.0f, _contentWidth - viewWidthDip);
            _horizontalOffset         = std::min(_horizontalOffset, maxHorizontal);

            // Snap to column boundary
            const float columnIndex = std::round((_horizontalOffset - kColumnSpacingDip) / columnStride);
            _horizontalOffset       = kColumnSpacingDip + (columnIndex * columnStride);
            _horizontalOffset       = std::clamp(_horizontalOffset, 0.0f, maxHorizontal);

            // Move focus right by visible columns
            if (hasFocus && ! _columnCounts.empty())
            {
                const auto& focused    = _items[_focusedIndex];
                const int maxColumn    = static_cast<int>(_columnCounts.size()) - 1;
                const int targetColumn = std::min(maxColumn, focused.column + visibleColumns);
                const int targetRow    = std::min(focused.row, _columnCounts[static_cast<size_t>(targetColumn)] - 1);

                if (auto newIndex = indexFromColumnRow(targetColumn, targetRow))
                {
                    focusIndex(*newIndex);
                }
            }

            UpdateScrollMetrics();
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
            break;
        }
        case VK_BACK: navigateUp(); break;
        case VK_SPACE:
            ExitIncrementalSearch();
            if (hasFocus)
            {
                ToggleSelection(_focusedIndex);

                if (_focusedIndex + 1 < _items.size())
                {
                    FocusItem(_focusedIndex + 1, true);
                    _anchorIndex = _focusedIndex;
                }

                if (_selectionSizeComputationRequestedCallback)
                {
                    _selectionSizeComputationRequestedCallback();
                }
            }
            break;
        case VK_INSERT:
            ExitIncrementalSearch();
            if (hasFocus)
            {
                ToggleSelection(_focusedIndex);
                if (_focusedIndex + 1 < _items.size())
                {
                    FocusItem(_focusedIndex + 1, true);
                    _anchorIndex = _focusedIndex;
                }
            }
            break;
        case VK_LEFT: navigateHorizontal(-1); break;
        case VK_RIGHT: navigateHorizontal(1); break;
        case VK_UP:
            ExitIncrementalSearch();
            if (hasFocus && _focusedIndex > 0)
            {
                const size_t newIndex = _focusedIndex - 1;
                focusIndex(newIndex);
            }
            break;
        case VK_DOWN:
            ExitIncrementalSearch();
            if (hasFocus && _focusedIndex + 1 < _items.size())
            {
                const size_t newIndex = _focusedIndex + 1;
                focusIndex(newIndex);
            }
            break;
        case VK_RETURN:
            ExitIncrementalSearch();
            ActivateFocusedItem();
            break;
        case VK_DELETE:
            ExitIncrementalSearch();
            CommandDelete();
            break;
        case VK_F2:
            ExitIncrementalSearch();
            RenameFocusedItem();
            break;
    }
}

void FolderView::OnCharMessage(wchar_t character)
{
    const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool alt  = (GetKeyState(VK_MENU) & 0x8000) != 0;

    if (ctrl || alt)
    {
        return;
    }

    if (character == L'\b')
    {
        if (_incrementalSearch.active)
        {
            HandleIncrementalSearchBackspace();
        }
        return;
    }

    if (! std::iswprint(static_cast<wint_t>(character)))
    {
        return;
    }

    if (! _incrementalSearch.active)
    {
        _incrementalSearch.active = true;
        _incrementalSearch.query.clear();
    }

    _incrementalSearch.query.push_back(character);
    NotifyIncrementalSearchChanged();
    UpdateIncrementalSearchIndicatorState(GetTickCount64(), true, _incrementalSearch.query);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    if (_items.empty())
    {
        return;
    }

    const auto invalidIndex = static_cast<size_t>(-1);
    const bool hasFocus     = _focusedIndex != invalidIndex && _focusedIndex < _items.size();

    if (hasFocus && FindIncrementalSearchMatchOffset(_items[_focusedIndex].displayName).has_value())
    {
        UpdateIncrementalSearchHighlightForFocusedItem();
        return;
    }

    auto findMatchIndex = [&](size_t startIndex, bool forward) -> std::optional<size_t>
    {
        const size_t itemCount = _items.size();
        if (itemCount == 0)
        {
            return std::nullopt;
        }

        for (size_t offset = 0; offset < itemCount; ++offset)
        {
            size_t index = 0;
            if (forward)
            {
                index = (startIndex + offset) % itemCount;
            }
            else
            {
                index = (startIndex + itemCount - offset) % itemCount;
            }

            if (FindIncrementalSearchMatchOffset(_items[index].displayName).has_value())
            {
                return index;
            }
        }
        return std::nullopt;
    };

    size_t startIndex = 0;
    if (hasFocus)
    {
        startIndex = (_focusedIndex + 1) % _items.size();
    }

    const std::optional<size_t> matchIndex = findMatchIndex(startIndex, true);
    if (matchIndex.has_value())
    {
        FocusItem(matchIndex.value(), true);
        _anchorIndex = matchIndex.value();
        return;
    }

    ClearIncrementalSearchHighlight();
}

void FolderView::NotifyIncrementalSearchChanged() const noexcept
{
    if (_incrementalSearchChangedCallback)
    {
        _incrementalSearchChangedCallback();
    }
}

void FolderView::ExitIncrementalSearch() noexcept
{
    if (! _incrementalSearch.active && _incrementalSearch.query.empty() && _incrementalSearch.highlightedIndex == static_cast<size_t>(-1))
    {
        return;
    }
    const std::wstring previousQuery = _incrementalSearch.query;
    _incrementalSearch.active        = false;
    _incrementalSearch.query.clear();
    ClearIncrementalSearchHighlight();
    NotifyIncrementalSearchChanged();
    UpdateIncrementalSearchIndicatorState(GetTickCount64(), false, previousQuery);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderView::UpdateIncrementalSearchIndicatorState(uint64_t nowTickMs, bool triggerPulse, std::wstring_view displayQuery) noexcept
{
    _incrementalSearchIndicatorDisplayQuery.assign(displayQuery);

    const float targetVisibility = _incrementalSearch.active ? 1.0f : 0.0f;
    if (_incrementalSearchIndicatorVisibilityTo != targetVisibility)
    {
        _incrementalSearchIndicatorVisibilityFrom  = _incrementalSearchIndicatorVisibility;
        _incrementalSearchIndicatorVisibilityTo    = targetVisibility;
        _incrementalSearchIndicatorVisibilityStart = nowTickMs;
    }

    if (triggerPulse && targetVisibility > 0.0f)
    {
        _incrementalSearchIndicatorTypingPulseStart = nowTickMs;
    }

    _incrementalSearchIndicatorLayout.reset();
    _incrementalSearchIndicatorLayoutText.clear();
    _incrementalSearchIndicatorLayoutMaxWidthDip = 0.0f;
    _incrementalSearchIndicatorLayoutMetrics     = {};

    StartOverlayAnimation();
}

void FolderView::HandleIncrementalSearchBackspace()
{
    if (! _incrementalSearch.active)
    {
        return;
    }

    if (! _incrementalSearch.query.empty())
    {
        _incrementalSearch.query.pop_back();
    }

    if (_incrementalSearch.query.empty())
    {
        ExitIncrementalSearch();
        return;
    }

    NotifyIncrementalSearchChanged();
    UpdateIncrementalSearchIndicatorState(GetTickCount64(), true, _incrementalSearch.query);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    UpdateIncrementalSearchHighlightForFocusedItem();
    if (_incrementalSearch.highlightedIndex != static_cast<size_t>(-1))
    {
        return;
    }

    HandleIncrementalSearchNavigate(true);
}

void FolderView::HandleIncrementalSearchNavigate(bool forward)
{
    if (! _incrementalSearch.active || _incrementalSearch.query.empty())
    {
        return;
    }

    if (_items.empty())
    {
        return;
    }

    const auto invalidIndex = static_cast<size_t>(-1);
    const bool hasFocus     = _focusedIndex != invalidIndex && _focusedIndex < _items.size();

    size_t startIndex = 0;
    if (! hasFocus)
    {
        startIndex = forward ? 0 : (_items.size() - 1);
    }
    else if (forward)
    {
        startIndex = (_focusedIndex + 1) % _items.size();
    }
    else
    {
        startIndex = (_focusedIndex + _items.size() - 1) % _items.size();
    }

    const size_t itemCount = _items.size();
    for (size_t offset = 0; offset < itemCount; ++offset)
    {
        size_t index = 0;
        if (forward)
        {
            index = (startIndex + offset) % itemCount;
        }
        else
        {
            index = (startIndex + itemCount - offset) % itemCount;
        }

        if (! FindIncrementalSearchMatchOffset(_items[index].displayName).has_value())
        {
            continue;
        }

        FocusItem(index, true);
        _anchorIndex = index;
        return;
    }
}

void FolderView::UpdateIncrementalSearchHighlightForFocusedItem()
{
    if (! _incrementalSearch.active || _incrementalSearch.query.empty())
    {
        ClearIncrementalSearchHighlight();
        return;
    }

    const auto invalidIndex = static_cast<size_t>(-1);
    if (_focusedIndex == invalidIndex || _focusedIndex >= _items.size())
    {
        ClearIncrementalSearchHighlight();
        return;
    }

    const FolderItem& item                  = _items[_focusedIndex];
    const std::optional<UINT32> matchOffset = FindIncrementalSearchMatchOffset(item.displayName);
    if (! matchOffset.has_value())
    {
        ClearIncrementalSearchHighlight();
        return;
    }

    if (_incrementalSearch.query.size() > static_cast<size_t>(std::numeric_limits<UINT32>::max()))
    {
        ClearIncrementalSearchHighlight();
        return;
    }

    DWRITE_TEXT_RANGE range{};
    range.startPosition = matchOffset.value();
    range.length        = static_cast<UINT32>(_incrementalSearch.query.size());
    ApplyIncrementalSearchHighlight(_focusedIndex, range);
}

void FolderView::ClearIncrementalSearchHighlight() noexcept
{
    const auto invalidIndex = static_cast<size_t>(-1);
    const size_t itemIndex  = _incrementalSearch.highlightedIndex;
    if (itemIndex == invalidIndex)
    {
        return;
    }

    if (itemIndex < _items.size())
    {
        FolderItem& item = _items[itemIndex];
        if (item.labelLayout && _incrementalSearch.highlightedRange.length > 0)
        {
            if (item.displayName.size() <= static_cast<size_t>(std::numeric_limits<UINT32>::max()))
            {
                const UINT32 textLength = static_cast<UINT32>(item.displayName.size());
                if (_incrementalSearch.highlightedRange.startPosition < textLength)
                {
                    DWRITE_TEXT_RANGE normalizedRange{};
                    normalizedRange.startPosition = _incrementalSearch.highlightedRange.startPosition;
                    normalizedRange.length =
                        std::min(_incrementalSearch.highlightedRange.length, textLength - _incrementalSearch.highlightedRange.startPosition);

                    static_cast<void>(item.labelLayout->SetDrawingEffect(nullptr, normalizedRange));
                }
            }
        }

        if (_hWnd)
        {
            RECT rc = ToPixelRect(OffsetRect(item.bounds, -_horizontalOffset, -_scrollOffset), _dpi);
            InvalidateRect(_hWnd.get(), &rc, FALSE);
        }
    }

    _incrementalSearch.highlightedIndex = invalidIndex;
    _incrementalSearch.highlightedRange = {};
}

void FolderView::ApplyIncrementalSearchHighlight(size_t itemIndex, const DWRITE_TEXT_RANGE& range) noexcept
{
    if (itemIndex >= _items.size())
    {
        return;
    }

    const auto invalidIndex    = static_cast<size_t>(-1);
    const size_t previousIndex = _incrementalSearch.highlightedIndex;
    const bool hasPrevious     = previousIndex != invalidIndex && previousIndex < _items.size();

    auto invalidateItem = [&](size_t index) noexcept
    {
        if (! _hWnd || index >= _items.size())
        {
            return;
        }

        const FolderItem& item = _items[index];
        RECT rc                = ToPixelRect(OffsetRect(item.bounds, -_horizontalOffset, -_scrollOffset), _dpi);
        InvalidateRect(_hWnd.get(), &rc, FALSE);
    };

    auto clearRangeFormatting = [&](FolderItem& itemToClear, const DWRITE_TEXT_RANGE& clearRange) noexcept
    {
        if (! itemToClear.labelLayout || clearRange.length == 0)
        {
            return;
        }

        if (itemToClear.displayName.size() > static_cast<size_t>(std::numeric_limits<UINT32>::max()))
        {
            return;
        }

        const UINT32 textLength = static_cast<UINT32>(itemToClear.displayName.size());
        if (clearRange.startPosition >= textLength)
        {
            return;
        }

        DWRITE_TEXT_RANGE normalizedRange{};
        normalizedRange.startPosition = clearRange.startPosition;

        const UINT32 availableLength = textLength - clearRange.startPosition;
        normalizedRange.length       = std::min(clearRange.length, availableLength);

        static_cast<void>(itemToClear.labelLayout->SetDrawingEffect(nullptr, normalizedRange));
    };

    if (hasPrevious)
    {
        FolderItem& previousItem = _items[previousIndex];
        clearRangeFormatting(previousItem, _incrementalSearch.highlightedRange);

        if (previousIndex != itemIndex)
        {
            invalidateItem(previousIndex);
        }
    }

    _incrementalSearch.highlightedIndex = itemIndex;
    _incrementalSearch.highlightedRange = range;

    FolderItem& item = _items[itemIndex];
    if (item.labelLayout)
    {
        if (item.displayName.size() <= static_cast<size_t>(std::numeric_limits<UINT32>::max()))
        {
            const UINT32 textLength = static_cast<UINT32>(item.displayName.size());
            if (range.startPosition < textLength)
            {
                DWRITE_TEXT_RANGE normalizedRange{};
                normalizedRange.startPosition = range.startPosition;

                const UINT32 availableLength = textLength - range.startPosition;
                normalizedRange.length       = std::min(range.length, availableLength);

                if (item.selected || ! _incrementalSearchHighlightBrush)
                {
                    static_cast<void>(item.labelLayout->SetDrawingEffect(nullptr, normalizedRange));
                }
                else
                {
                    static_cast<void>(item.labelLayout->SetDrawingEffect(_incrementalSearchHighlightBrush.get(), normalizedRange));
                }
            }
        }
    }

    invalidateItem(itemIndex);
}

std::optional<UINT32> FolderView::FindIncrementalSearchMatchOffset(std::wstring_view displayName) const noexcept
{
    if (! _incrementalSearch.active)
    {
        return std::nullopt;
    }

    const std::wstring_view query = _incrementalSearch.query;
    if (query.empty() || displayName.size() < query.size())
    {
        return std::nullopt;
    }

    if (query.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return std::nullopt;
    }

    const int queryLength          = static_cast<int>(query.size());
    const size_t lastStartPosition = displayName.size() - query.size();
    for (size_t startPosition = 0; startPosition <= lastStartPosition; ++startPosition)
    {
        const int compareResult = CompareStringOrdinal(displayName.data() + startPosition, queryLength, query.data(), queryLength, TRUE);
        if (compareResult != CSTR_EQUAL)
        {
            continue;
        }

        if (startPosition > static_cast<size_t>(std::numeric_limits<UINT32>::max()))
        {
            return std::nullopt;
        }

        return static_cast<UINT32>(startPosition);
    }

    return std::nullopt;
}
