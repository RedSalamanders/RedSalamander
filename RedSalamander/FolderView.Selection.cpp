#include "FolderViewInternal.h"

void FolderView::SelectSingle(size_t index)
{
    if (index >= _items.size())
        return;

    const size_t previousFocusedIndex = _focusedIndex;
    const int marginPx                = std::max(1, PxFromDip(kFocusStrokeThicknessDip));
    auto invalidateItem               = [&](size_t itemIndex) noexcept
    {
        const auto invalidIndex = static_cast<size_t>(-1);
        if (itemIndex == invalidIndex || itemIndex >= _items.size())
        {
            return;
        }

        RECT rc = ToPixelRect(OffsetRect(_items[itemIndex].bounds, -_horizontalOffset, -_scrollOffset), _dpi);
        InflateRect(&rc, marginPx, marginPx);
        InvalidateRect(_hWnd.get(), &rc, FALSE);
    };

    for (size_t i = 0; i < _items.size(); ++i)
    {
        bool shouldSelect = i == index;
        if (_items[i].selected != shouldSelect)
        {
            _items[i].selected = shouldSelect;
            invalidateItem(i);
        }
        _items[i].focused = false;
    }
    _items[index].focused = true;
    _focusedIndex         = index;
    invalidateItem(previousFocusedIndex);
    invalidateItem(index);
    _selectionStats = {};
    if (_items[index].isDirectory)
    {
        _selectionStats.selectedFolders = 1;
    }
    else
    {
        _selectionStats.selectedFiles     = 1;
        _selectionStats.selectedFileBytes = _items[index].sizeBytes;
    }
    {
        const FolderItem& item = _items[index];
        SelectionStats::SelectedItemDetails details{};
        details.isDirectory        = item.isDirectory;
        details.sizeBytes          = item.sizeBytes;
        details.lastWriteTime      = item.lastWriteTime;
        details.fileAttributes     = item.fileAttributes;
        _selectionStats.singleItem = details;
    }
    NotifySelectionChanged();
    EnsureVisible(index);
    UpdateIncrementalSearchHighlightForFocusedItem();
    RememberFocusedItemForDisplayedFolder();
}

void FolderView::ToggleSelection(size_t index)
{
    if (index >= _items.size())
        return;

    const auto invalidIndex = static_cast<size_t>(-1);
    const int marginPx      = std::max(1, PxFromDip(kFocusStrokeThicknessDip));
    auto invalidateItem     = [&](size_t itemIndex) noexcept
    {
        if (itemIndex == invalidIndex || itemIndex >= _items.size())
        {
            return;
        }

        RECT rc = ToPixelRect(OffsetRect(_items[itemIndex].bounds, -_horizontalOffset, -_scrollOffset), _dpi);
        InflateRect(&rc, marginPx, marginPx);
        InvalidateRect(_hWnd.get(), &rc, FALSE);
    };

    if (_focusedIndex != invalidIndex && _focusedIndex < _items.size() && _focusedIndex != index)
    {
        _items[_focusedIndex].focused = false;
        invalidateItem(_focusedIndex);
    }

    FolderItem& item = _items[index];
    item.selected    = ! item.selected;
    item.focused     = true;
    _focusedIndex    = index;

    RecomputeSelectionStats();
    NotifySelectionChanged();

    invalidateItem(index);
    UpdateIncrementalSearchHighlightForFocusedItem();
    RememberFocusedItemForDisplayedFolder();
}

void FolderView::RangeSelect(size_t index)
{
    if (index >= _items.size() || _anchorIndex >= _items.size())
        return;

    const size_t previousFocusedIndex = _focusedIndex;
    const int marginPx                = std::max(1, PxFromDip(kFocusStrokeThicknessDip));
    auto invalidateItem               = [&](size_t itemIndex) noexcept
    {
        const auto invalidIndex = static_cast<size_t>(-1);
        if (itemIndex == invalidIndex || itemIndex >= _items.size())
        {
            return;
        }

        RECT rc = ToPixelRect(OffsetRect(_items[itemIndex].bounds, -_horizontalOffset, -_scrollOffset), _dpi);
        InflateRect(&rc, marginPx, marginPx);
        InvalidateRect(_hWnd.get(), &rc, FALSE);
    };

    size_t minIndex = std::min(index, _anchorIndex);
    size_t maxIndex = std::max(index, _anchorIndex);
    SelectionStats stats{};
    const FolderItem* singleSelected = nullptr;
    uint32_t selectedTotal           = 0;
    for (size_t i = 0; i < _items.size(); ++i)
    {
        bool shouldSelect = (i >= minIndex && i <= maxIndex);
        if (_items[i].selected != shouldSelect)
        {
            _items[i].selected = shouldSelect;
            invalidateItem(i);
        }
        _items[i].focused = (i == index);
        if (_items[i].selected)
        {
            ++selectedTotal;
            if (selectedTotal == 1)
            {
                singleSelected = &_items[i];
            }
            else
            {
                singleSelected = nullptr;
            }
            if (_items[i].isDirectory)
            {
                ++stats.selectedFolders;
            }
            else
            {
                ++stats.selectedFiles;
                stats.selectedFileBytes += _items[i].sizeBytes;
            }
        }
    }
    if (selectedTotal == 1 && singleSelected)
    {
        SelectionStats::SelectedItemDetails details{};
        details.isDirectory    = singleSelected->isDirectory;
        details.sizeBytes      = singleSelected->sizeBytes;
        details.lastWriteTime  = singleSelected->lastWriteTime;
        details.fileAttributes = singleSelected->fileAttributes;
        stats.singleItem       = details;
    }
    _focusedIndex   = index;
    _selectionStats = stats;
    invalidateItem(previousFocusedIndex);
    invalidateItem(index);
    NotifySelectionChanged();
    EnsureVisible(index);
    UpdateIncrementalSearchHighlightForFocusedItem();
    RememberFocusedItemForDisplayedFolder();
}

void FolderView::ClearSelection()
{
    bool selectionChanged = false;
    for (auto& item : _items)
    {
        if (! item.selected)
        {
            continue;
        }

        selectionChanged = true;
        RECT rc          = ToPixelRect(OffsetRect(item.bounds, -_horizontalOffset, -_scrollOffset), _dpi);
        InvalidateRect(_hWnd.get(), &rc, FALSE);
        item.selected = false;
    }
    _selectionStats = {};
    if (selectionChanged)
    {
        NotifySelectionChanged();
    }
    UpdateIncrementalSearchHighlightForFocusedItem();
}

void FolderView::SelectAll()
{
    SelectionStats stats{};
    for (auto& item : _items)
    {
        item.selected = true;
        if (item.isDirectory)
        {
            ++stats.selectedFolders;
        }
        else
        {
            ++stats.selectedFiles;
            stats.selectedFileBytes += item.sizeBytes;
        }
    }

    if ((stats.selectedFiles + stats.selectedFolders) == 1 && ! _items.empty())
    {
        const FolderItem& item = _items[0];
        SelectionStats::SelectedItemDetails details{};
        details.isDirectory    = item.isDirectory;
        details.sizeBytes      = item.sizeBytes;
        details.lastWriteTime  = item.lastWriteTime;
        details.fileAttributes = item.fileAttributes;
        stats.singleItem       = details;
    }

    _selectionStats = stats;
    NotifySelectionChanged();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
    UpdateIncrementalSearchHighlightForFocusedItem();
}

void FolderView::InvertSelection()
{
    if (_items.empty())
    {
        _selectionStats = {};
        NotifySelectionChanged();
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
        UpdateIncrementalSearchHighlightForFocusedItem();
        return;
    }

    for (auto& item : _items)
    {
        item.selected = ! item.selected;
    }

    RecomputeSelectionStats();
    NotifySelectionChanged();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
    UpdateIncrementalSearchHighlightForFocusedItem();
}

void FolderView::SetSelectionByDisplayNamePredicate(const std::function<bool(std::wstring_view)>& shouldSelect, bool clearExistingSelection)
{
    if (_items.empty())
    {
        _selectionStats = {};
        NotifySelectionChanged();
        return;
    }

    bool changed = false;
    for (auto& item : _items)
    {
        const bool wantsSelect = shouldSelect ? shouldSelect(item.displayName) : false;
        const bool desired     = clearExistingSelection ? wantsSelect : (item.selected || wantsSelect);
        if (item.selected != desired)
        {
            item.selected = desired;
            changed       = true;
        }
    }

    if (! changed)
    {
        return;
    }

    RecomputeSelectionStats();
    NotifySelectionChanged();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
    UpdateIncrementalSearchHighlightForFocusedItem();
}

void FolderView::ClearSelectionByDisplayNamePredicate(const std::function<bool(std::wstring_view)>& shouldUnselect)
{
    if (_items.empty())
    {
        return;
    }

    bool changed = false;
    for (auto& item : _items)
    {
        if (! item.selected)
        {
            continue;
        }

        const bool wantsUnselect = shouldUnselect ? shouldUnselect(item.displayName) : false;
        if (! wantsUnselect)
        {
            continue;
        }

        item.selected = false;
        changed       = true;
    }

    if (! changed)
    {
        return;
    }

    RecomputeSelectionStats();
    NotifySelectionChanged();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
    UpdateIncrementalSearchHighlightForFocusedItem();
}

void FolderView::SelectSameExtension()
{
    const auto invalidIndex = static_cast<size_t>(-1);
    if (_focusedIndex == invalidIndex || _focusedIndex >= _items.size())
    {
        return;
    }

    const FolderItem& focused = _items[_focusedIndex];
    if (focused.isDirectory)
    {
        return;
    }

    const std::wstring_view extension = focused.GetExtension();

    bool changed = false;
    for (auto& item : _items)
    {
        if (item.isDirectory)
        {
            continue;
        }

        if (! OrdinalString::EqualsNoCase(item.GetExtension(), extension))
        {
            continue;
        }

        if (! item.selected)
        {
            item.selected = true;
            changed       = true;
        }
    }

    if (! changed)
    {
        return;
    }

    RecomputeSelectionStats();
    NotifySelectionChanged();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
    UpdateIncrementalSearchHighlightForFocusedItem();
}

void FolderView::UnselectSameExtension()
{
    const auto invalidIndex = static_cast<size_t>(-1);
    if (_focusedIndex == invalidIndex || _focusedIndex >= _items.size())
    {
        return;
    }

    const FolderItem& focused = _items[_focusedIndex];
    if (focused.isDirectory)
    {
        return;
    }

    const std::wstring_view extension = focused.GetExtension();

    bool changed = false;
    for (auto& item : _items)
    {
        if (item.isDirectory || ! item.selected)
        {
            continue;
        }

        if (! OrdinalString::EqualsNoCase(item.GetExtension(), extension))
        {
            continue;
        }

        item.selected = false;
        changed       = true;
    }

    if (! changed)
    {
        return;
    }

    RecomputeSelectionStats();
    NotifySelectionChanged();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
    UpdateIncrementalSearchHighlightForFocusedItem();
}

void FolderView::HideSelectedNames()
{
    if (_items.empty())
    {
        return;
    }

    std::vector<std::wstring> namesToHide;
    namesToHide.reserve(static_cast<size_t>(_selectionStats.selectedFolders + _selectionStats.selectedFiles));

    for (const auto& item : _items)
    {
        if (! item.selected || item.displayName.empty())
        {
            continue;
        }

        namesToHide.emplace_back(item.displayName);
    }

    if (namesToHide.empty())
    {
        return;
    }

    const auto hiddenNames = _hiddenNames.load(std::memory_order_acquire);
    auto updated           = hiddenNames ? std::make_shared<HiddenNamesFilter>(*hiddenNames) : std::make_shared<HiddenNamesFilter>();
    updated->names.reserve(updated->names.size() + namesToHide.size());
    for (auto& name : namesToHide)
    {
        updated->names.insert(std::move(name));
    }

    std::shared_ptr<const HiddenNamesFilter> updatedConst = std::move(updated);
    _hiddenNames.store(std::move(updatedConst), std::memory_order_release);

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    RequestRefreshFromCache();
}

void FolderView::HideUnselectedNames()
{
    if (_items.empty())
    {
        return;
    }

    bool hasSelected = false;
    std::vector<std::wstring> namesToHide;
    namesToHide.reserve(_items.size());

    for (const auto& item : _items)
    {
        hasSelected = hasSelected || item.selected;

        if (item.selected || item.displayName.empty())
        {
            continue;
        }

        namesToHide.emplace_back(item.displayName);
    }

    if (! hasSelected || namesToHide.empty())
    {
        return;
    }

    const auto hiddenNames = _hiddenNames.load(std::memory_order_acquire);
    auto updated           = hiddenNames ? std::make_shared<HiddenNamesFilter>(*hiddenNames) : std::make_shared<HiddenNamesFilter>();
    updated->names.reserve(updated->names.size() + namesToHide.size());
    for (auto& name : namesToHide)
    {
        updated->names.insert(std::move(name));
    }

    std::shared_ptr<const HiddenNamesFilter> updatedConst = std::move(updated);
    _hiddenNames.store(std::move(updatedConst), std::memory_order_release);

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    RequestRefreshFromCache();
}

void FolderView::ShowHiddenNames()
{
    const auto hiddenNames = _hiddenNames.load(std::memory_order_acquire);
    if (! hiddenNames || hiddenNames->names.empty())
    {
        return;
    }

    _hiddenNames.store(std::shared_ptr<const HiddenNamesFilter>{}, std::memory_order_release);

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    RequestRefreshFromCache();
}

void FolderView::RecomputeSelectionStats() noexcept
{
    SelectionStats stats{};
    const FolderItem* singleSelected = nullptr;
    uint32_t selectedTotal           = 0;
    for (const auto& item : _items)
    {
        if (! item.selected)
        {
            continue;
        }

        ++selectedTotal;
        if (selectedTotal == 1)
        {
            singleSelected = &item;
        }
        else
        {
            singleSelected = nullptr;
        }

        if (item.isDirectory)
        {
            ++stats.selectedFolders;
        }
        else
        {
            ++stats.selectedFiles;
            stats.selectedFileBytes += item.sizeBytes;
        }
    }

    if (selectedTotal == 1 && singleSelected)
    {
        SelectionStats::SelectedItemDetails details{};
        details.isDirectory    = singleSelected->isDirectory;
        details.sizeBytes      = singleSelected->sizeBytes;
        details.lastWriteTime  = singleSelected->lastWriteTime;
        details.fileAttributes = singleSelected->fileAttributes;
        stats.singleItem       = details;
    }

    _selectionStats = stats;
}

void FolderView::NotifySelectionChanged() const noexcept
{
    if (_selectionChangedCallback)
    {
        _selectionChangedCallback(_selectionStats);
    }
}

void FolderView::FocusItem(size_t index, bool ensureVisible)
{
    if (index >= _items.size())
        return;

    const auto invalidIndex = static_cast<size_t>(-1);
    const int marginPx      = std::max(1, PxFromDip(kFocusStrokeThicknessDip));
    auto invalidateItem     = [&](size_t itemIndex) noexcept
    {
        if (itemIndex == invalidIndex || itemIndex >= _items.size())
        {
            return;
        }

        RECT rc = ToPixelRect(OffsetRect(_items[itemIndex].bounds, -_horizontalOffset, -_scrollOffset), _dpi);
        InflateRect(&rc, marginPx, marginPx);
        InvalidateRect(_hWnd.get(), &rc, FALSE);
    };

    if (_focusedIndex != invalidIndex && _focusedIndex < _items.size())
    {
        _items[_focusedIndex].focused = false;
        invalidateItem(_focusedIndex);
    }

    _items[index].focused = true;
    _focusedIndex         = index;
    invalidateItem(index);
    if (ensureVisible)
    {
        EnsureVisible(index);
    }
    UpdateIncrementalSearchHighlightForFocusedItem();
    RememberFocusedItemForDisplayedFolder();
}

[[nodiscard]] bool FolderView::GoToPreviousSelectedName()
{
    return GoToSelectedName(/*forward*/ false);
}

[[nodiscard]] bool FolderView::GoToNextSelectedName()
{
    return GoToSelectedName(/*forward*/ true);
}

[[nodiscard]] bool FolderView::GoToSelectedName(bool forward)
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
        return false;
    }

    ExitIncrementalSearch();

    if (_items.empty())
    {
        return false;
    }

    const auto invalidIndex = static_cast<size_t>(-1);
    const bool hasFocus     = _focusedIndex != invalidIndex && _focusedIndex < _items.size();

    uint32_t selectedCount  = 0;
    size_t firstSelectedIdx = invalidIndex;
    size_t lastSelectedIdx  = invalidIndex;
    for (size_t i = 0; i < _items.size(); ++i)
    {
        if (! _items[i].selected)
        {
            continue;
        }

        ++selectedCount;
        if (firstSelectedIdx == invalidIndex)
        {
            firstSelectedIdx = i;
        }
        lastSelectedIdx = i;
    }

    if (selectedCount == 0u)
    {
        return false;
    }

    const auto focusIndex = [&](size_t index) noexcept -> bool
    {
        if (index == invalidIndex || index >= _items.size() || index == _focusedIndex)
        {
            return false;
        }

        FocusItem(index, true);
        _anchorIndex = index;
        return true;
    };

    if (! hasFocus)
    {
        const size_t target = forward ? firstSelectedIdx : lastSelectedIdx;
        return focusIndex(target);
    }

    if (selectedCount == 1u)
    {
        return focusIndex(firstSelectedIdx);
    }

    if (forward)
    {
        for (size_t i = _focusedIndex + 1; i < _items.size(); ++i)
        {
            if (_items[i].selected)
            {
                return focusIndex(i);
            }
        }

        for (size_t i = 0; i < _focusedIndex; ++i)
        {
            if (_items[i].selected)
            {
                return focusIndex(i);
            }
        }

        return false;
    }

    if (_focusedIndex > 0)
    {
        for (size_t i = _focusedIndex; i-- > 0;)
        {
            if (_items[i].selected)
            {
                return focusIndex(i);
            }
        }
    }

    for (size_t i = _items.size(); i-- > (_focusedIndex + 1);)
    {
        if (_items[i].selected)
        {
            return focusIndex(i);
        }
    }

    return false;
}

bool FolderView::PrepareForExternalCommand(std::wstring_view focusItemDisplayName) noexcept
{
    if (! _hWnd || focusItemDisplayName.empty() || _items.empty())
    {
        return false;
    }

    std::optional<size_t> match;
    for (size_t i = 0; i < _items.size(); ++i)
    {
        if (_items[i].displayName == focusItemDisplayName)
        {
            match = i;
            break;
        }
    }

    if (! match.has_value())
    {
        for (size_t i = 0; i < _items.size(); ++i)
        {
            if (OrdinalString::EqualsNoCase(_items[i].displayName, focusItemDisplayName))
            {
                match = i;
                break;
            }
        }
    }

    if (! match.has_value())
    {
        return false;
    }

    ClearSelection();
    FocusItem(match.value(), true);
    _anchorIndex = match.value();
    return true;
}

void FolderView::ActivateFocusedItem()
{
    if (_focusedIndex == static_cast<size_t>(-1) || _focusedIndex >= _items.size())
        return;

    const auto& item = _items[_focusedIndex];
    if (item.isDirectory)
    {
        SetFolderPath(GetItemFullPath(item));
    }
    else
    {
        const std::filesystem::path fullPath = GetItemFullPath(item);
        bool handled                         = false;
        if (_openFileRequestCallback)
        {
            handled = _openFileRequestCallback(fullPath);
        }

        if (! handled)
        {
            ShellExecuteW(_hWnd.get(), L"open", fullPath.c_str(), nullptr, _currentFolder ? _currentFolder->c_str() : nullptr, SW_SHOWNORMAL);
        }
    }
}

std::vector<std::filesystem::path> FolderView::GetSelectedPaths() const
{
    std::vector<std::filesystem::path> paths;
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            paths.push_back(GetItemFullPath(item));
        }
    }
    return paths;
}

std::vector<std::filesystem::path> FolderView::GetSelectedOrFocusedPaths() const
{
    std::vector<std::filesystem::path> paths = GetSelectedPaths();
    if (! paths.empty())
    {
        return paths;
    }

    if (_focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
    {
        paths.push_back(GetItemFullPath(_items[_focusedIndex]));
    }

    return paths;
}

std::vector<std::wstring> FolderView::GetSelectedOrFocusedDisplayNames() const
{
    std::vector<std::wstring> names;
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            names.emplace_back(item.displayName);
        }
    }

    if (! names.empty())
    {
        return names;
    }

    if (_focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
    {
        names.emplace_back(_items[_focusedIndex].displayName);
    }

    return names;
}

std::vector<FolderView::PathAttributes> FolderView::GetSelectedOrFocusedPathAttributes() const
{
    std::vector<PathAttributes> items;
    for (const auto& item : _items)
    {
        if (! item.selected)
        {
            continue;
        }
        PathAttributes info{};
        info.path           = GetItemFullPath(item);
        info.fileAttributes = item.fileAttributes;
        items.push_back(std::move(info));
    }
    if (! items.empty())
    {
        return items;
    }

    if (_focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
    {
        const FolderItem& item = _items[_focusedIndex];
        PathAttributes info{};
        info.path           = GetItemFullPath(item);
        info.fileAttributes = item.fileAttributes;
        items.push_back(std::move(info));
    }

    return items;
}

std::vector<std::filesystem::path> FolderView::GetSelectedDirectoryPaths() const
{
    std::vector<std::filesystem::path> paths;
    for (const auto& item : _items)
    {
        if (! item.selected || ! item.isDirectory)
        {
            continue;
        }

        if ((item.fileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0)
        {
            continue;
        }

        paths.push_back(GetItemFullPath(item));
    }
    return paths;
}

#ifdef _DEBUG
std::wstring_view FolderView::DebugGetFocusedDisplayName() const noexcept
{
    const auto invalidIndex = static_cast<size_t>(-1);
    if (_focusedIndex == invalidIndex || _focusedIndex >= _items.size())
    {
        return {};
    }

    return _items[_focusedIndex].displayName;
}

bool FolderView::DebugHasItemDisplayName(std::wstring_view displayName) const noexcept
{
    if (displayName.empty())
    {
        return false;
    }

    for (const auto& item : _items)
    {
        if (item.displayName == displayName)
        {
            return true;
        }
    }

    return false;
}

bool FolderView::DebugIsItemSelectedByDisplayName(std::wstring_view displayName) const noexcept
{
    if (displayName.empty())
    {
        return false;
    }

    for (const auto& item : _items)
    {
        if (item.displayName == displayName)
        {
            return item.selected;
        }
    }

    return false;
}

size_t FolderView::DebugGetSelectedItemCount() const noexcept
{
    size_t count = 0;
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            ++count;
        }
    }

    return count;
}

bool FolderView::DebugIsEmptyFolderStateActive() const noexcept
{
    return CanShowEmptyFolderState() && _emptyFolderState.has_value();
}

FolderView::FilterWatermarkVisualMode FolderView::DebugGetFilterWatermarkVisualMode() const noexcept
{
    if (! IsNameFilterActive())
    {
        return FilterWatermarkVisualMode::None;
    }

    if (! _items.empty())
    {
        return FilterWatermarkVisualMode::Background;
    }

    bool hasOverlay = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        hasOverlay = _errorOverlay.has_value();
    }

    const bool canShowEmptyUi = ! hasOverlay && ! _pendingBusyOverlay.has_value();
    if (canShowEmptyUi && ! _emptyStateMessage.empty())
    {
        return FilterWatermarkVisualMode::Badge;
    }

    if (canShowEmptyUi && CanShowEmptyFolderState() && _emptyFolderState.has_value())
    {
        return FilterWatermarkVisualMode::Badge;
    }

    return FilterWatermarkVisualMode::Background;
}

std::wstring_view FolderView::DebugGetEmptyFolderFunMessage() const noexcept
{
    if (! _emptyFolderState)
    {
        return {};
    }

    return _emptyFolderState->funMessage;
}
#endif
