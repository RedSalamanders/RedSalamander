#include "FolderViewInternal.h"

#ifdef ENABLE_TESTS
#include "Common.Tests.Runner.h"
#endif

void FolderView::SelectSingle(size_t index)
{
    if (index >= _items.size())
        return;

    const size_t previousFocusedIndex = _focusedIndex;
    if (previousFocusedIndex != index)
    {
        ++_removalFocusOwnershipEpoch;
    }
    const int marginPx  = std::max(1, PxFromDip(kFocusStrokeThicknessDip));
    auto invalidateItem = [&](size_t itemIndex) noexcept
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
    const uint64_t pathGeneration        = _folderPathGeneration;
    const uint64_t providerEpoch         = _removalFocusProviderEpoch;
    const uint64_t enumerationGeneration = _enumerationGeneration.load(std::memory_order_acquire);
    NotifySelectionChanged();
    // A host selection callback may navigate or replace the model. The old index
    // must not reach EnsureVisible or focus-memory updates after that reentry.
    if (pathGeneration != _folderPathGeneration || providerEpoch != _removalFocusProviderEpoch ||
        enumerationGeneration != _enumerationGeneration.load(std::memory_order_acquire) || _focusedIndex != index || index >= _items.size() ||
        ! _items[index].selected)
        return;
    EnsureVisible(index);
    UpdateIncrementalSearchHighlightForFocusedItem();
    RememberFocusedItemForDisplayedFolder();
    if (previousFocusedIndex != _focusedIndex)
    {
        NotifyFocusedItemChanged();
    }
}

void FolderView::ToggleSelection(size_t index)
{
    if (index >= _items.size())
        return;

    const auto invalidIndex           = static_cast<size_t>(-1);
    const size_t previousFocusedIndex = _focusedIndex;
    if (previousFocusedIndex != index)
    {
        ++_removalFocusOwnershipEpoch;
    }
    const int marginPx  = std::max(1, PxFromDip(kFocusStrokeThicknessDip));
    auto invalidateItem = [&](size_t itemIndex) noexcept
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
    if (previousFocusedIndex != _focusedIndex)
    {
        NotifyFocusedItemChanged();
    }
}

void FolderView::RangeSelect(size_t index, const bool additive)
{
    if (index >= _items.size() || _anchorIndex >= _items.size())
        return;

    const size_t previousFocusedIndex = _focusedIndex;
    if (previousFocusedIndex != index)
    {
        ++_removalFocusOwnershipEpoch;
    }
    const int marginPx  = std::max(1, PxFromDip(kFocusStrokeThicknessDip));
    auto invalidateItem = [&](size_t itemIndex) noexcept
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
        const bool shouldSelect = additive ? (_items[i].selected || (i >= minIndex && i <= maxIndex)) : (i >= minIndex && i <= maxIndex);
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
    if (previousFocusedIndex != _focusedIndex)
    {
        NotifyFocusedItemChanged();
    }
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

void FolderView::SelectSameName()
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

    const std::wstring_view nameWithoutExtension = focused.GetNameWithoutExtension();

    bool changed = false;
    for (auto& item : _items)
    {
        if (item.isDirectory)
        {
            continue;
        }

        if (! OrdinalString::EqualsNoCase(item.GetNameWithoutExtension(), nameWithoutExtension))
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

void FolderView::UnselectSameName()
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

    const std::wstring_view nameWithoutExtension = focused.GetNameWithoutExtension();

    bool changed = false;
    for (auto& item : _items)
    {
        if (item.isDirectory || ! item.selected)
        {
            continue;
        }

        if (! OrdinalString::EqualsNoCase(item.GetNameWithoutExtension(), nameWithoutExtension))
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

void FolderView::NotifySelectionChanged(SelectionChangeKind kind) const noexcept
{
    if (_selectionChangedCallback)
    {
        _selectionChangedCallback(_selectionStats, kind);
    }
}

std::optional<FolderView::SelectionStats::SelectedItemDetails> FolderView::GetFocusedItemDetails() const noexcept
{
    const auto invalidIndex = static_cast<size_t>(-1);
    if (_focusedIndex == invalidIndex || _focusedIndex >= _items.size())
    {
        return std::nullopt;
    }

    const FolderItem& item = _items[_focusedIndex];
    SelectionStats::SelectedItemDetails details{};
    details.isDirectory    = item.isDirectory;
    details.sizeBytes      = item.sizeBytes;
    details.lastWriteTime  = item.lastWriteTime;
    details.fileAttributes = item.fileAttributes;
    return details;
}

std::shared_ptr<const FileOperationArtifacts::Projection> FolderView::GetFocusedArtifactProjection() const noexcept
{
    const auto invalidIndex = static_cast<size_t>(-1);
    if (_focusedIndex == invalidIndex || _focusedIndex >= _items.size())
    {
        return {};
    }
    return _items[_focusedIndex].artifactProjection;
}

std::vector<FolderView::PossibleArtifactRow> FolderView::GetPossibleArtifactRows() const
{
    std::vector<PossibleArtifactRow> rows;
    for (const FolderItem& item : _items)
    {
        if (! item.artifactProjection || item.artifactProjection->classification == FileOperationArtifacts::Classification::Ordinary)
        {
            continue;
        }
        rows.push_back(PossibleArtifactRow{
            .path            = GetItemFullPath(item),
            .sizeBytes       = item.sizeBytes,
            .lastWriteTime   = item.lastWriteTime,
            .fileAttributes  = item.fileAttributes,
            .isDirectory     = item.isDirectory,
            .mayHoldOriginal = item.artifactProjection->mayHoldOriginal,
        });
    }
    return rows;
}

bool FolderView::HasPossibleArtifactRows() const noexcept
{
    return std::ranges::any_of(_items, [](const FolderItem& item) noexcept {
        return item.artifactProjection && item.artifactProjection->classification != FileOperationArtifacts::Classification::Ordinary;
    });
}

void FolderView::NotifyFocusedItemChanged() const noexcept
{
    if (_focusedItemChangedCallback)
    {
        _focusedItemChangedCallback();
    }
}

void FolderView::FocusItem(size_t index, bool ensureVisible)
{
    if (index >= _items.size())
        return;

    const auto invalidIndex           = static_cast<size_t>(-1);
    const size_t previousFocusedIndex = _focusedIndex;
    if (previousFocusedIndex != index)
    {
        ++_removalFocusOwnershipEpoch;
    }
    const int marginPx  = std::max(1, PxFromDip(kFocusStrokeThicknessDip));
    auto invalidateItem = [&](size_t itemIndex) noexcept
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
    if (previousFocusedIndex != _focusedIndex)
    {
        NotifyFocusedItemChanged();
    }
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

std::optional<FolderView::InspectionVisit> FolderView::CaptureInspectionVisit()
{
    if (! _hWnd || GetWindowThreadProcessId(_hWnd.get(), nullptr) != GetCurrentThreadId() || ! _currentFolder)
        return std::nullopt;
    if (! _inspectionLifetime)
        _inspectionLifetime = std::make_shared<const uint8_t>(uint8_t{0});
    InspectionVisit visit;
    visit.lifetime              = _inspectionLifetime;
    visit.pathGeneration        = _folderPathGeneration;
    visit.enumerationGeneration = _enumerationGeneration.load(std::memory_order_acquire);
    visit.providerEpoch         = _removalFocusProviderEpoch;
    visit.focusOwnershipEpoch   = _removalFocusOwnershipEpoch;
    return visit;
}

bool FolderView::IsInspectionVisitCurrent(const InspectionVisit& visit, bool includeFocusOwnership) const noexcept
{
    return _hWnd && GetWindowThreadProcessId(_hWnd.get(), nullptr) == GetCurrentThreadId() && _inspectionLifetime &&
           visit.lifetime.lock() == _inspectionLifetime && visit.pathGeneration == _folderPathGeneration &&
           visit.enumerationGeneration == _enumerationGeneration.load(std::memory_order_acquire) && visit.providerEpoch == _removalFocusProviderEpoch &&
           (! includeFocusOwnership || visit.focusOwnershipEpoch == _removalFocusOwnershipEpoch);
}

bool FolderView::IsSameNavigationVisit(const InspectionVisit& original, const InspectionVisit& settled) noexcept
{
    const auto lifetime = original.lifetime.lock();
    return lifetime && lifetime == settled.lifetime.lock() && original.pathGeneration == settled.pathGeneration &&
           original.providerEpoch == settled.providerEpoch && original.focusOwnershipEpoch == settled.focusOwnershipEpoch &&
           settled.enumerationGeneration >= original.enumerationGeneration;
}

FolderView::InspectionSelectionResult FolderView::SelectItemForInspection(const InspectionVisit& visit,
                                                                          const FileSystemPathIdentity& admittedIdentity,
                                                                          std::wstring_view providerPath,
                                                                          std::optional<bool> itemIsDirectory) noexcept
{
    Debug::Perf::Scope perf(L"folder.inspection.select_us");
    if (! IsInspectionVisitCurrent(visit, true))
        return {InspectionSelectionStatus::Stale};
    if (_inspectionCompletedGeneration != visit.enumerationGeneration)
        return {InspectionSelectionStatus::Pending};
    if (FAILED(_inspectionEnumerationStatus))
        return {InspectionSelectionStatus::EnumerationFailed, _inspectionEnumerationStatus};
    // Existing queued commands/selections run after the enumeration callback. They
    // must not silently replace a selection that inspection reported as successful.
    if (! admittedIdentity.pathTextStableIdentity || ! IsCurrentFolderEnumerated() || _pendingExternalCommandAfterEnumeration ||
        _pendingExternalSelectionAfterEnumeration)
        return {InspectionSelectionStatus::Unavailable};
    std::wstring parent;
    std::wstring leaf;
    if (! TryGetFileSystemParentPath(admittedIdentity, providerPath, parent) || ! TryGetFileSystemLeafName(admittedIdentity, providerPath, leaf) ||
        leaf.empty() || ! EquivalentPath(admittedIdentity, parent, _currentFolder->native()) ||
        ! EquivalentPath(admittedIdentity, parent, _itemsFolder.native()))
        return {InspectionSelectionStatus::Unavailable};

    std::optional<size_t> match;
    std::optional<size_t> kindMatch;
    size_t nameMatches = 0u;
    size_t kindMatches = 0u;
    for (size_t index = 0u; index < _items.size(); ++index)
    {
        if (! EquivalentComponent(admittedIdentity, _items[index].displayName, leaf))
            continue;
        ++nameMatches;
        match = match.value_or(index);
        if (itemIsDirectory.has_value() && _items[index].isDirectory == itemIsDirectory.value())
        {
            ++kindMatches;
            kindMatch = kindMatch.value_or(index);
        }
    }
    if (! match.has_value())
        return {InspectionSelectionStatus::MissingOrFiltered};
    if (nameMatches > 1u)
    {
        // A provider may list a file and a folder under one name; the admitted kind breaks the tie.
        if (kindMatches != 1u)
            return {InspectionSelectionStatus::Ambiguous};
        match = kindMatch;
    }

    _anchorIndex = match.value();
    SelectSingle(match.value()); // Canonical selection, current-item and ensure-visible path; no OS focus or file opening.
    // Selection notifications can reenter the owner. Never report success against a
    // replacement model or a different current/selected item after those callbacks.
    if (! IsInspectionVisitCurrent(visit, false) || _focusedIndex >= _items.size() || ! _items[_focusedIndex].selected ||
        ! EquivalentComponent(admittedIdentity, _items[_focusedIndex].displayName, leaf) ||
        _selectionStats.selectedFiles + _selectionStats.selectedFolders != 1u)
        return {InspectionSelectionStatus::Stale};
    return {InspectionSelectionStatus::Selected};
}

#ifdef ENABLE_TESTS
bool FolderView::DebugValidateInspectionSelectionReentryForSelfTest(const FileSystemPathIdentity& identity, std::wstring_view providerPath)
{
    const auto visit = CaptureInspectionVisit();
    SelfTest::AppendSelfTestTrace(std::format(L"Inspection reentry: captured={} pathGeneration={} enumerationGeneration={} completedGeneration={} "
                                              L"providerEpoch={} focusEpoch={} pendingCommand={} pendingSelection={} items={} current='{}' model='{}'",
                                              visit.has_value(),
                                              _folderPathGeneration,
                                              _enumerationGeneration.load(std::memory_order_acquire),
                                              _inspectionCompletedGeneration,
                                              _removalFocusProviderEpoch,
                                              _removalFocusOwnershipEpoch,
                                              _pendingExternalCommandAfterEnumeration.has_value(),
                                              _pendingExternalSelectionAfterEnumeration.has_value(),
                                              _items.size(),
                                              _currentFolder ? _currentFolder->native() : std::wstring{},
                                              _itemsFolder.native()));
    if (! visit)
        return false;
    // Represent an enumeration result queued before the clear. Retain the COM
    // arena as well as the rows: their displayName fields borrow that storage.
    auto latePayload                          = std::make_unique<EnumerationPayload>();
    latePayload->generation                   = visit->enumerationGeneration;
    latePayload->items                        = _items;
    latePayload->arenaBuffer                  = _itemsArenaBuffer;
    latePayload->folder                       = _itemsFolder;
    SelectionChangedCallback previousCallback = std::move(_selectionChangedCallback);
    const auto restoreCallback                = wil::scope_exit([&] { _selectionChangedCallback = std::move(previousCallback); });
    bool redirected                           = false;
    size_t notificationCount                  = 0u;
    _selectionChangedCallback                 = [&](const SelectionStats&, SelectionChangeKind)
    {
        ++notificationCount;
        if (! redirected)
        {
            redirected = true;
            SetFolderPath(std::nullopt);
        }
    };
    const auto selection = SelectItemForInspection(visit.value(), identity, providerPath);
    const bool cleared   = selection.status == InspectionSelectionStatus::Stale && redirected && _items.empty() && ! _currentFolder;
    SelfTest::AppendSelfTestTrace(std::format(L"Inspection reentry: status={} notifications={} redirected={} cleared={} pathGeneration={} "
                                              L"enumerationGeneration={} items={} current='{}'",
                                              static_cast<unsigned>(selection.status),
                                              notificationCount,
                                              redirected,
                                              cleared,
                                              _folderPathGeneration,
                                              _enumerationGeneration.load(std::memory_order_acquire),
                                              _items.size(),
                                              _currentFolder ? _currentFolder->native() : std::wstring{}));
    if (! cleared)
        return false;
    ProcessEnumerationResult(std::move(latePayload));
    const bool rejectedLatePayload = _enumerationGeneration.load(std::memory_order_acquire) != visit->enumerationGeneration && _items.empty() &&
                                     ! _itemsArenaBuffer && _itemsFolder.empty() && ! _currentFolder;
    SelfTest::AppendSelfTestTrace(std::format(L"Inspection reentry: queued pre-clear payload rejected={}", rejectedLatePayload));
    return rejectedLatePayload;
}

std::optional<FolderView::DebugKindTwinRows> FolderView::DebugAppendKindTwinRowForSelfTest(const FileSystemPathIdentity& identity,
                                                                                           std::wstring_view providerPath)
{
    std::wstring leaf;
    if (! TryGetFileSystemLeafName(identity, providerPath, leaf) || leaf.empty())
        return std::nullopt;
    const auto fileRow = std::ranges::find_if(_items, [&](const FolderItem& item) noexcept { return EquivalentComponent(identity, item.displayName, leaf); });
    if (fileRow == _items.end() || fileRow->isDirectory)
        return std::nullopt;
    const size_t fileIndex = static_cast<size_t>(std::distance(_items.begin(), fileRow));
    // The synthetic row borrows the file row's arena name; the arena outlives the twin, which
    // DebugRemoveKindTwinRowForSelfTest removes before the model is replaced.
    FolderItem folderRow  = _items[fileIndex];
    folderRow.isDirectory = true;
    folderRow.selected    = false;
    folderRow.focused     = false;
    _items.push_back(std::move(folderRow));
    return DebugKindTwinRows{fileIndex, _items.size() - 1u};
}

void FolderView::DebugRemoveKindTwinRowForSelfTest(const DebugKindTwinRows& rows)
{
    if (rows.folderIndex != _items.size() - 1u || rows.fileIndex >= rows.folderIndex)
        return;
    if (_focusedIndex == rows.folderIndex || _items[rows.folderIndex].selected)
        SelectSingle(rows.fileIndex);
    _items.erase(_items.begin() + static_cast<std::ptrdiff_t>(rows.folderIndex));
}

std::wstring FolderView::DebugValidateInspectionKindTieBreakForSelfTest(const FileSystemPathIdentity& identity, std::wstring_view providerPath)
{
    const auto rows = DebugAppendKindTwinRowForSelfTest(identity, providerPath);
    if (! rows)
        return L"the settled model has no file row with the inspected name";
    const size_t fileIndex   = rows->fileIndex;
    const size_t folderIndex = rows->folderIndex;
    const auto restore       = wil::scope_exit([&] { DebugRemoveKindTwinRowForSelfTest(rows.value()); });
    const auto select        = [&](std::optional<bool> isDirectory) noexcept
    {
        const auto visit = CaptureInspectionVisit();
        return visit ? SelectItemForInspection(visit.value(), identity, providerPath, isDirectory) : InspectionSelectionResult{};
    };
    if (const auto unknownKind = select(std::nullopt); unknownKind.status != InspectionSelectionStatus::Ambiguous)
        return std::format(L"a file and a folder with one name and no admitted kind returned status {}, not Ambiguous",
                           static_cast<unsigned>(unknownKind.status));
    if (const auto folder = select(true); folder.status != InspectionSelectionStatus::Selected || _focusedIndex != folderIndex)
        return std::format(L"the folder kind returned status {} with current row {} instead of selecting the folder row {}",
                           static_cast<unsigned>(folder.status),
                           _focusedIndex,
                           folderIndex);
    if (const auto file = select(false); file.status != InspectionSelectionStatus::Selected || _focusedIndex != fileIndex)
        return std::format(L"the file kind returned status {} with current row {} instead of selecting the file row {}",
                           static_cast<unsigned>(file.status),
                           _focusedIndex,
                           fileIndex);
    // Two rows of the inspected kind stay ambiguous: the kind breaks a tie, it does not pick one.
    _items[folderIndex].isDirectory = false;
    if (const auto twoFiles = select(false); twoFiles.status != InspectionSelectionStatus::Ambiguous)
        return std::format(L"two file rows with one name returned status {}, not Ambiguous", static_cast<unsigned>(twoFiles.status));
    return {};
}
#endif

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

std::optional<std::filesystem::path> FolderView::GetFocusedPath() const
{
    const auto invalidIndex = static_cast<size_t>(-1);
    if (_focusedIndex == invalidIndex || _focusedIndex >= _items.size())
    {
        return std::nullopt;
    }

    return GetItemFullPath(_items[_focusedIndex]);
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

#ifdef ENABLE_TESTS
std::wstring_view FolderView::DebugGetFocusedDisplayName() const noexcept
{
    const auto invalidIndex = static_cast<size_t>(-1);
    if (_focusedIndex == invalidIndex || _focusedIndex >= _items.size())
    {
        return {};
    }

    return _items[_focusedIndex].displayName;
}

std::wstring FolderView::DebugGetFocusedVisualDisplayName() const
{
    const auto invalidIndex = static_cast<size_t>(-1);
    if (_focusedIndex == invalidIndex || _focusedIndex >= _items.size())
    {
        return {};
    }

    return std::wstring(GetVisualDisplayName(_items[_focusedIndex]));
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

size_t FolderView::DebugGetItemCount() const noexcept
{
    return _items.size();
}

size_t FolderView::DebugGetBitmapIconCount() const noexcept
{
    size_t count = 0;
    for (const auto& item : _items)
    {
        if (item.icon)
        {
            ++count;
        }
    }
    return count;
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

FolderView::DebugFocusSelectionStateSnapshot FolderView::DebugGetFocusSelectionStateSnapshot() const noexcept
{
    constexpr uint64_t kFnvOffset = 14695981039346656037ull;
    constexpr uint64_t kFnvPrime  = 1099511628211ull;
    const auto appendDigest       = [](uint64_t& digest, std::wstring_view text) noexcept
    {
        for (const wchar_t ch : text)
        {
            digest ^= static_cast<uint16_t>(ch);
            digest *= kFnvPrime;
        }
        digest ^= 0xffffu;
        digest *= kFnvPrime;
    };

    DebugFocusSelectionStateSnapshot snapshot{};
    snapshot.itemCount                 = _items.size();
    snapshot.currentIndex              = _focusedIndex;
    snapshot.anchorIndex               = _anchorIndex;
    snapshot.resolutionReason          = _lastCurrentResolutionReason;
    snapshot.focusOwnershipEpoch       = _removalFocusOwnershipEpoch;
    snapshot.selectedFileCount         = _selectionStats.selectedFiles;
    snapshot.selectedFolderCount       = _selectionStats.selectedFolders;
    snapshot.selectedFileBytes         = _selectionStats.selectedFileBytes;
    snapshot.emptyParentActionActive   = CanShowEmptyFolderState() && _emptyFolderState.has_value();
    snapshot.selectionDigest           = kFnvOffset;
    snapshot.potentialDragSourceDigest = kFnvOffset;

    if (_focusedIndex < _items.size())
    {
        snapshot.currentDisplayName.assign(_items[_focusedIndex].displayName);
    }

    for (const FolderItem& item : _items)
    {
        if (item.focused)
        {
            ++snapshot.focusedFlagCount;
        }
        if (item.selected)
        {
            ++snapshot.selectedCount;
            appendDigest(snapshot.selectionDigest, item.displayName);
        }
    }

    if (_drag.potential)
    {
        snapshot.potentialDragSourceCount = _drag.sourceDisplayNames.size();
        for (const std::wstring& displayName : _drag.sourceDisplayNames)
        {
            appendDigest(snapshot.potentialDragSourceDigest, displayName);
        }
    }

    snapshot.focusMemoryEntryCount                   = _focusMemory.size();
    snapshot.focusMemoryPayloadBytes                 = _focusMemoryPayloadBytes;
    snapshot.focusMemoryEvictionCount                = _focusMemoryEvictionCount;
    snapshot.focusMemoryHitCount                     = _focusMemoryHitCount;
    snapshot.focusMemoryMissCount                    = _focusMemoryMissCount;
    snapshot.focusMemoryNoCacheCount                 = _focusMemoryNoCacheCount;
    snapshot.focusMemoryLocationIdentityNoCacheCount = _focusMemoryLocationIdentityNoCacheCount;
    snapshot.focusMemoryItemIdentityNoCacheCount     = _focusMemoryItemIdentityNoCacheCount;
    snapshot.focusMemoryOversizedEntryNoCacheCount   = _focusMemoryOversizedEntryNoCacheCount;
    snapshot.focusMemoryLastNoCacheReason            = _focusMemoryLastNoCacheReason;
    return snapshot;
}

void FolderView::DebugClearFocusMemoryForSelfTest() noexcept
{
    ClearFocusMemory();
    _focusMemoryEvictionCount                = 0u;
    _focusMemoryHitCount                     = 0u;
    _focusMemoryMissCount                    = 0u;
    _focusMemoryNoCacheCount                 = 0u;
    _focusMemoryLocationIdentityNoCacheCount = 0u;
    _focusMemoryItemIdentityNoCacheCount     = 0u;
    _focusMemoryOversizedEntryNoCacheCount   = 0u;
    _focusMemoryLastNoCacheReason            = FocusMemoryNoCacheReason::None;
}

bool FolderView::DebugRememberFocusMemoryEntryForSelfTest(const std::filesystem::path& folder, const std::wstring_view itemDisplayName) noexcept
{
    return RecordFocusMemoryEntry(folder, itemDisplayName);
}

std::wstring FolderView::DebugLookupFocusMemoryEntryForSelfTest(const std::filesystem::path& folder) noexcept
{
    return GetRememberedFocusedItemPathForFolder(folder);
}

std::optional<POINT> FolderView::DebugGetEmptyBackgroundClientPointForSelfTest() const noexcept
{
    if (! _hWnd)
    {
        return std::nullopt;
    }

    RECT client{};
    if (GetClientRect(_hWnd.get(), &client) == FALSE)
    {
        return std::nullopt;
    }
    for (int y = client.bottom - 2; y >= client.top; y -= 4)
    {
        for (int x = client.right - 2; x >= client.left; x -= 4)
        {
            const POINT candidate{x, y};
            if (! HitTest(candidate).has_value())
            {
                return candidate;
            }
        }
    }
    return std::nullopt;
}

bool FolderView::DebugIsEmptyFolderStateActive() const noexcept
{
    return CanShowEmptyFolderState() && _emptyFolderState.has_value();
}

std::wstring_view FolderView::DebugGetEmptyStateMessage() const noexcept
{
    return _emptyStateMessage;
}

FolderView::DebugEmptyFolderItemMetrics FolderView::DebugGetEmptyFolderItemMetrics() const noexcept
{
    return DebugEmptyFolderItemMetrics{
        .active          = CanShowEmptyFolderState() && _emptyFolderState.has_value(),
        .clientWidthDip  = DipFromPx(_clientSize.cx),
        .clientHeightDip = DipFromPx(_clientSize.cy),
        .tileWidthDip    = _tileWidthDip,
        .tileHeightDip   = _tileHeightDip,
        .labelHeightDip  = _labelHeightDip,
    };
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

bool FolderView::DebugGetColumnLayoutSnapshot(DebugColumnLayoutSnapshot& out) const
{
    out                        = DebugColumnLayoutSnapshot{};
    out.displayMode            = static_cast<uint8_t>(_displayMode);
    out.rowsPerColumn          = _rowsPerColumn;
    out.horizontalOffsetDip    = _horizontalOffset;
    out.contentWidthDip        = _contentWidth;
    out.clientWidthDip         = DipFromPx(_clientSize.cx);
    out.clientHeightDip        = DipFromPx(_clientSize.cy);
    out.tileWidthDip           = _tileWidthDip;
    out.tileHeightDip          = _tileHeightDip;
    out.maxHorizontalOffsetDip = std::max(0.0f, _contentWidth - out.clientWidthDip);

    const auto [visibleStart, visibleEnd] = GetVisibleItemRange();
    out.firstVisibleIndex                 = std::min(visibleStart, _items.size());
    out.lastVisibleIndex                  = visibleEnd > visibleStart ? std::min(visibleEnd - 1u, _items.size() - 1u) : out.firstVisibleIndex;

    out.columns.reserve(_columnLayout.empty() ? _columnCounts.size() : _columnLayout.size());
    size_t runningStart = 0;
    for (size_t columnIndex = 0; columnIndex < (_columnLayout.empty() ? _columnCounts.size() : _columnLayout.size()); ++columnIndex)
    {
        const bool hasColumnLayout = columnIndex < _columnLayout.size();
        const size_t startIndex    = hasColumnLayout ? std::min(_columnLayout[columnIndex].startIndex, _items.size())
                                                     : (columnIndex < _columnPrefixSums.size() ? std::min(_columnPrefixSums[columnIndex], _items.size())
                                                                                               : std::min(runningStart, _items.size()));
        const size_t itemCount     = hasColumnLayout ? _columnLayout[columnIndex].itemCount : static_cast<size_t>(std::max(0, _columnCounts[columnIndex]));
        const size_t endIndex      = std::min(startIndex + itemCount, _items.size());

        DebugColumnLayoutEntry entry{};
        entry.startIndex = startIndex;
        entry.itemCount  = endIndex - startIndex;
        entry.leftDip =
            hasColumnLayout ? _columnLayout[columnIndex].leftDip : kColumnSpacingDip + (static_cast<float>(columnIndex) * (_tileWidthDip + kColumnSpacingDip));
        entry.widthDip = hasColumnLayout ? _columnLayout[columnIndex].widthDip : _tileWidthDip;
        entry.rightDip = entry.leftDip + entry.widthDip;

        for (size_t itemIndex = startIndex; itemIndex < endIndex; ++itemIndex)
        {
            const FolderItem& item = _items[itemIndex];
            const float labelWidth =
                item.labelMetrics.widthIncludingTrailingWhitespace > 0.0f ? item.labelMetrics.widthIncludingTrailingWhitespace : item.labelMetrics.width;
            const float detailsWidth =
                item.detailsMetrics.widthIncludingTrailingWhitespace > 0.0f ? item.detailsMetrics.widthIncludingTrailingWhitespace : item.detailsMetrics.width;
            const float metadataWidth = item.metadataMetrics.widthIncludingTrailingWhitespace > 0.0f ? item.metadataMetrics.widthIncludingTrailingWhitespace
                                                                                                     : item.metadataMetrics.width;
            const float itemWidth     = std::max(labelWidth, std::max(detailsWidth, metadataWidth));
            if (itemWidth >= entry.widestObservedWidthDip)
            {
                entry.widestObservedWidthDip = itemWidth;
                entry.widestDisplayName      = std::wstring(GetVisualDisplayName(item));
            }
        }

        out.columns.push_back(std::move(entry));
        runningStart = endIndex;
    }

    if (_hWnd && out.columns.size() > 1u)
    {
        const float rowSpacingDip = GetFolderViewRowSpacingDip(_appTheme);
        const float sampleYDip    = rowSpacingDip + 1.0f - _scrollOffset;
        for (size_t columnIndex = 0; columnIndex + 1u < out.columns.size(); ++columnIndex)
        {
            const auto& leftColumn  = out.columns[columnIndex];
            const auto& rightColumn = out.columns[columnIndex + 1u];
            if (rightColumn.leftDip <= leftColumn.rightDip)
            {
                continue;
            }

            const float sampleXDip = leftColumn.rightDip + ((rightColumn.leftDip - leftColumn.rightDip) * 0.5f) - _horizontalOffset;
            POINT pt{
                .x = PxFromDip(sampleXDip),
                .y = PxFromDip(sampleYDip),
            };
            if (HitTest(pt).has_value())
            {
                ++out.hitTestSpacingFalsePositiveCount;
            }
        }
    }

    if (visibleEnd > visibleStart && visibleStart < _items.size())
    {
        const size_t clampedEnd = std::min(visibleEnd, _items.size());
        out.visibleItems.reserve(clampedEnd - visibleStart);
        for (size_t itemIndex = visibleStart; itemIndex < clampedEnd; ++itemIndex)
        {
            const FolderItem& item = _items[itemIndex];
            out.visibleItems.push_back(DebugVisibleItemEntry{
                .index       = itemIndex,
                .displayName = std::wstring(GetVisualDisplayName(item)),
                .bounds      = item.bounds,
                .column      = item.column,
                .row         = item.row,
            });
        }
    }

    return true;
}
#endif
