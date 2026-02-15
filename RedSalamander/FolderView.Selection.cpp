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
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
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
