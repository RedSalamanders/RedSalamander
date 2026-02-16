#include "FolderViewInternal.h"

void FolderView::QueueIconLoading()
{
    if (_items.empty() || ! _hWnd)
    {
        return;
    }

    // Icon loading needs a valid D2D context to convert HICONs into D2D bitmaps.
    // During startup we can enumerate folders before deferred DirectX init; in that case we'll queue again
    // from `FolderView::OnDeferredInit()` once resources exist.
    if (! _d2dContext)
    {
        return;
    }

    // Initialize telemetry (per-batch)
    const uint64_t batchId = _iconLoadStats.batchId.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    _iconLoadStats.totalRequests.store(0u, std::memory_order_release);
    _iconLoadStats.visibleRequests.store(0u, std::memory_order_release);
    _iconLoadStats.cacheHits.store(0u, std::memory_order_release);
    _iconLoadStats.uniqueIconsQueued.store(0u, std::memory_order_release);
    _iconLoadStats.extracted.store(0u, std::memory_order_release);
    _iconLoadStats.bitmapPosted.store(0u, std::memory_order_release);
    _iconLoadStats.bitmapPostFailed.store(0u, std::memory_order_release);
    _iconLoadStats.bitmapConverted.store(0u, std::memory_order_release);
    _iconLoadStats.bitmapConvertFailed.store(0u, std::memory_order_release);
    _iconLoadStats.bitmapConvertUsTotal.store(0u, std::memory_order_release);
    _iconLoadStats.bitmapConvertUsMax.store(0u, std::memory_order_release);
    _iconLoadStats.pendingBitmapCreates.store(0u, std::memory_order_release);
    _iconLoadStats.bitmapFirstPostQpc.store(0, std::memory_order_release);
    _iconLoadStats.bitmapSummaryEmitted.store(false, std::memory_order_release);
    QueryPerformanceCounter(&_iconLoadStats.startTime);

    const float viewLeft   = _horizontalOffset;
    const float viewRight  = _horizontalOffset + DipFromPx(_clientSize.cx);
    const float viewTop    = _scrollOffset;
    const float viewBottom = _scrollOffset + DipFromPx(_clientSize.cy);

    struct GroupBuild
    {
        bool hasVisibleItems         = false;
        size_t firstVisibleItemIndex = static_cast<size_t>(-1);
        std::vector<size_t> itemIndices;
    };

    std::unordered_map<int, GroupBuild> groups;
    groups.reserve(std::min<size_t>(_items.size(), 256u));

    uint64_t totalNeeded   = 0;
    uint64_t visibleNeeded = 0;
    size_t skippedNoIndex  = 0;
    size_t skippedHasIcon  = 0;

    for (size_t i = 0; i < _items.size(); ++i)
    {
        const auto& item = _items[i];
        if (item.iconIndex < 0)
        {
            ++skippedNoIndex;
            continue;
        }

        if (item.icon)
        {
            ++skippedHasIcon;
            continue;
        }

        ++totalNeeded;

        const bool isVisible = ! (item.bounds.right < viewLeft || item.bounds.left > viewRight || item.bounds.bottom < viewTop || item.bounds.top > viewBottom);
        if (isVisible)
        {
            ++visibleNeeded;
        }

        auto& group           = groups[item.iconIndex];
        group.hasVisibleItems = group.hasVisibleItems || isVisible;
        if (isVisible)
        {
            group.firstVisibleItemIndex = std::min(group.firstVisibleItemIndex, i);
        }
        group.itemIndices.push_back(i);
    }

    // Build grouped requests and stamp already-cached bitmaps immediately.
    std::vector<IconLoadRequest> visibleRequests;
    std::vector<IconLoadRequest> offscreenRequests;
    visibleRequests.reserve(std::min<size_t>(groups.size(), 128u));
    offscreenRequests.reserve(std::min<size_t>(groups.size(), 128u));

    uint64_t stampedFromCache = 0;

    for (auto& [iconIndex, group] : groups)
    {
        if (iconIndex < 0 || group.itemIndices.empty())
        {
            continue;
        }

        // If the bitmap already exists for our D2D device, apply it immediately (no background work).
        auto cachedBitmap = IconCache::GetInstance().GetCachedBitmap(iconIndex, _d2dContext.get());
        if (cachedBitmap)
        {
            for (const size_t itemIndex : group.itemIndices)
            {
                if (itemIndex >= _items.size())
                {
                    continue;
                }
                auto& item = _items[itemIndex];
                if (item.icon || item.iconIndex != iconIndex)
                {
                    continue;
                }
                item.icon = cachedBitmap;
                ++stampedFromCache;
            }
            continue;
        }

        IconLoadRequest request;
        request.iconIndex             = iconIndex;
        request.hasVisibleItems       = group.hasVisibleItems;
        request.firstVisibleItemIndex = group.firstVisibleItemIndex;
        request.itemIndices           = std::move(group.itemIndices);

        if (request.hasVisibleItems)
        {
            visibleRequests.push_back(std::move(request));
        }
        else
        {
            offscreenRequests.push_back(std::move(request));
        }
    }

    // Process visible groups in view order so placeholders resolve in a stable, predictable way.
    std::sort(visibleRequests.begin(),
              visibleRequests.end(),
              [](const IconLoadRequest& a, const IconLoadRequest& b)
              {
                  if (a.firstVisibleItemIndex != b.firstVisibleItemIndex)
                  {
                      return a.firstVisibleItemIndex < b.firstVisibleItemIndex;
                  }
                  return a.itemIndices.size() > b.itemIndices.size();
              });

    std::deque<IconLoadRequest> newQueue;
    newQueue.insert(newQueue.end(), std::make_move_iterator(visibleRequests.begin()), std::make_move_iterator(visibleRequests.end()));
    newQueue.insert(newQueue.end(), std::make_move_iterator(offscreenRequests.begin()), std::make_move_iterator(offscreenRequests.end()));

    const uint64_t uniqueIconsQueued = static_cast<uint64_t>(newQueue.size());

    {
        std::lock_guard lock(_enumerationMutex);
        _iconLoadQueue = std::move(newQueue);
    }

    _iconLoadStats.totalRequests.store(totalNeeded, std::memory_order_release);
    _iconLoadStats.visibleRequests.store(visibleNeeded, std::memory_order_release);
    _iconLoadStats.cacheHits.store(stampedFromCache, std::memory_order_release);
    _iconLoadStats.uniqueIconsQueued.store(uniqueIconsQueued, std::memory_order_release);

    if (uniqueIconsQueued > 0)
    {
        _iconLoadingActive.store(true, std::memory_order_release);
        _enumerationCv.notify_one();
    }

    Debug::Info(L"FolderView: Icon load queued - {} items ({} visible), {} cached, {} unique icons queued; skipped {} no-index, {} has-icon",
                totalNeeded,
                visibleNeeded,
                stampedFromCache,
                uniqueIconsQueued,
                skippedNoIndex,
                skippedHasIcon);

    static_cast<void>(batchId);
}

void FolderView::BoostIconLoadingForVisibleRange()
{
    if (_items.empty() || ! _hWnd || ! _d2dContext)
    {
        return;
    }

    const auto [visStart, visEnd] = GetVisibleItemRange();
    if (visStart >= _items.size() || visEnd <= visStart)
    {
        return;
    }

    // Include a small buffer around the visible range to reduce scroll-pop-in.
    constexpr size_t kBufferItems = 64;
    const size_t rangeStart       = (visStart > kBufferItems) ? (visStart - kBufferItems) : 0;
    const size_t rangeEnd         = std::min(visEnd + kBufferItems, _items.size());

    std::vector<int> neededIconIndices;
    neededIconIndices.reserve(std::min<size_t>(rangeEnd - rangeStart, 256u));

    // Fast-path: if the bitmap already exists for our device, stamp it immediately.
    for (size_t i = rangeStart; i < rangeEnd; ++i)
    {
        auto& item = _items[i];
        if (item.icon || item.iconIndex < 0)
        {
            continue;
        }

        if (auto cached = IconCache::GetInstance().GetCachedBitmap(item.iconIndex, _d2dContext.get()))
        {
            item.icon = std::move(cached);
            continue;
        }

        neededIconIndices.push_back(item.iconIndex);
    }

    if (neededIconIndices.empty())
    {
        return;
    }

    std::sort(neededIconIndices.begin(), neededIconIndices.end());
    neededIconIndices.erase(std::unique(neededIconIndices.begin(), neededIconIndices.end()), neededIconIndices.end());

    bool boosted     = false;
    bool shouldQueue = false;
    {
        std::lock_guard lock(_enumerationMutex);
        if (_iconLoadQueue.empty())
        {
            shouldQueue = true;
        }
        else
        {
            std::deque<IconLoadRequest> highPriority;
            std::deque<IconLoadRequest> lowPriority;

            while (! _iconLoadQueue.empty())
            {
                IconLoadRequest request = std::move(_iconLoadQueue.front());
                _iconLoadQueue.pop_front();

                const bool needed = std::binary_search(neededIconIndices.begin(), neededIconIndices.end(), request.iconIndex);
                if (needed)
                {
                    request.hasVisibleItems = true;
                    highPriority.push_back(std::move(request));
                    boosted = true;
                }
                else
                {
                    lowPriority.push_back(std::move(request));
                }
            }

            if (boosted)
            {
                _iconLoadQueue = std::move(highPriority);
                _iconLoadQueue.insert(_iconLoadQueue.end(), std::make_move_iterator(lowPriority.begin()), std::make_move_iterator(lowPriority.end()));
            }
            else
            {
                _iconLoadQueue = std::move(lowPriority);
            }
        }
    }

    if (boosted)
    {
        _enumerationCv.notify_one();
    }
    else if (shouldQueue)
    {
        QueueIconLoading();
    }
}

void FolderView::ProcessIconLoadQueue()
{
    const uint64_t batchId = _iconLoadStats.batchId.load(std::memory_order_acquire);
    Debug::Perf::Scope perf(L"FolderView.IconLoading.ProcessQueue");
    perf.SetDetail(_itemsFolder.native());
    perf.SetValue0(_iconLoadStats.totalRequests.load(std::memory_order_relaxed));

    while (_iconLoadingActive.load(std::memory_order_acquire))
    {
        if (_iconLoadStats.batchId.load(std::memory_order_acquire) != batchId)
        {
            break;
        }

        IconLoadRequest request{};

        {
            std::lock_guard lock(_enumerationMutex);
            if (_iconLoadQueue.empty())
            {
                _iconLoadingActive.store(false, std::memory_order_release);

                // Log completion statistics
                LARGE_INTEGER frequency{};
                QueryPerformanceFrequency(&frequency);
                LARGE_INTEGER endTime{};
                QueryPerformanceCounter(&endTime);
                auto elapsedMs = static_cast<double>(((endTime.QuadPart - _iconLoadStats.startTime.QuadPart) * 1000000) / frequency.QuadPart) / 1000.0f;

                const uint64_t totalRequests = _iconLoadStats.totalRequests.load(std::memory_order_relaxed);
                const uint64_t cacheHits     = _iconLoadStats.cacheHits.load(std::memory_order_relaxed);
                const uint64_t uniqueQueued  = _iconLoadStats.uniqueIconsQueued.load(std::memory_order_relaxed);
                const size_t cacheHitRate    = totalRequests > 0 ? static_cast<size_t>((cacheHits * 100u) / totalRequests) : 0;

                // Get cache memory usage
                const size_t cacheMemoryMB = IconCache::GetInstance().GetMemoryUsage() / (1024 * 1024);
                const auto cacheStats      = IconCache::GetInstance().GetStats();

                Debug::Info(L"FolderView: Icon loading complete - {} items ({} visible), {} cached ({}%), {} unique queued, {} extracted, ({:.3f}ms)",
                            totalRequests,
                            _iconLoadStats.visibleRequests.load(std::memory_order_relaxed),
                            cacheHits,
                            cacheHitRate,
                            uniqueQueued,
                            _iconLoadStats.extracted.load(std::memory_order_relaxed),
                            elapsedMs);
                Debug::Info(L"FolderView: IconCache stats - {} cached icons (~{} MB), {} hits, {} misses, {} LRU evictions",
                            cacheStats.cacheSize,
                            cacheMemoryMB,
                            cacheStats.hitCount,
                            cacheStats.missCount,
                            cacheStats.lruEvictions);
                break;
            }

            request = std::move(_iconLoadQueue.front());
            _iconLoadQueue.pop_front(); // O(1) with deque vs O(N) with vector
        }

        if (request.iconIndex < 0 || request.itemIndices.empty())
        {
            continue;
        }

        wil::com_ptr<ID2D1Device> d2dDeviceSnapshot;
        {
            std::lock_guard lock(_d2dDeviceMutex);
            d2dDeviceSnapshot = _d2dDevice;
        }
        const bool cachedForDevice = d2dDeviceSnapshot && IconCache::GetInstance().HasCachedIcon(request.iconIndex, d2dDeviceSnapshot.get());

        auto bitmapRequest             = std::make_unique<IconBitmapRequest>();
        bitmapRequest->iconLoadBatchId = batchId;
        bitmapRequest->iconIndex       = request.iconIndex;
        bitmapRequest->itemIndices     = std::move(request.itemIndices);

        // Background thread: extract once per iconIndex (unless already cached).
        if (! cachedForDevice)
        {
            wil::unique_hicon hIcon = IconCache::GetInstance().ExtractSystemIcon(request.iconIndex, _iconSizeDip);
            if (! hIcon)
            {
                continue;
            }
            _iconLoadStats.extracted.fetch_add(1u, std::memory_order_relaxed);
            bitmapRequest->hIcon = std::move(hIcon);
        }

        if (! _hWnd)
        {
            continue;
        }

        const bool posted = PostMessagePayload(_hWnd.get(), WndMsg::kFolderViewCreateIconBitmap, 0, std::move(bitmapRequest));
        if (posted)
        {
            _iconLoadStats.bitmapPosted.fetch_add(1u, std::memory_order_relaxed);
            _iconLoadStats.pendingBitmapCreates.fetch_add(1u, std::memory_order_relaxed);

            if (_iconLoadStats.bitmapFirstPostQpc.load(std::memory_order_relaxed) == 0)
            {
                LARGE_INTEGER qpc{};
                QueryPerformanceCounter(&qpc);
                int64_t expected = 0;
                static_cast<void>(
                    _iconLoadStats.bitmapFirstPostQpc.compare_exchange_strong(expected, qpc.QuadPart, std::memory_order_release, std::memory_order_relaxed));
            }
        }
        else
        {
            _iconLoadStats.bitmapPostFailed.fetch_add(1u, std::memory_order_relaxed);
        }

        // Yield occasionally to avoid hogging CPU on large off-screen batches.
        if (! request.hasVisibleItems && ((_iconLoadStats.bitmapPosted.load(std::memory_order_relaxed) % 25u) == 0u))
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(1));
        }
    }

    perf.SetValue1(_iconLoadStats.extracted.load(std::memory_order_relaxed));
}

void FolderView::OnCreateIconBitmap(std::unique_ptr<IconBitmapRequest> requestPtr)
{
    // This runs on UI thread - safe to use _d2dContext
    if (! requestPtr)
    {
        return;
    }

    const uint64_t batchId = _iconLoadStats.batchId.load(std::memory_order_acquire);
    if (requestPtr->iconLoadBatchId != batchId)
    {
        return;
    }

    const auto onExit = wil::scope_exit(
        [&]() noexcept
        {
            const uint64_t remaining = _iconLoadStats.pendingBitmapCreates.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
            static_cast<void>(remaining);
            MaybeEmitIconBitmapSummary(batchId);
        });

    if (! _d2dContext || requestPtr->iconIndex < 0 || requestPtr->itemIndices.empty())
    {
        return;
    }

    wil::com_ptr<ID2D1Bitmap1> bitmap;
    if (requestPtr->hIcon)
    {
        // Convert HICON to D2D bitmap on UI thread (thread-safe)
        const auto convertStart = std::chrono::steady_clock::now();
        bitmap                  = IconCache::GetInstance().ConvertIconToBitmapOnUIThread(requestPtr->hIcon.get(), requestPtr->iconIndex, _d2dContext.get());
        const auto convertEnd   = std::chrono::steady_clock::now();

        const uint64_t convertUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(convertEnd - convertStart).count());
        _iconLoadStats.bitmapConverted.fetch_add(1u, std::memory_order_relaxed);
        _iconLoadStats.bitmapConvertUsTotal.fetch_add(convertUs, std::memory_order_relaxed);

        uint64_t maxUs = _iconLoadStats.bitmapConvertUsMax.load(std::memory_order_relaxed);
        while (convertUs > maxUs && ! _iconLoadStats.bitmapConvertUsMax.compare_exchange_weak(maxUs, convertUs, std::memory_order_relaxed))
        {
        }

        if (! bitmap)
        {
            _iconLoadStats.bitmapConvertFailed.fetch_add(1u, std::memory_order_relaxed);
            return;
        }
    }
    else
    {
        // Already cached for our device; just retrieve it.
        bitmap = IconCache::GetInstance().GetCachedBitmap(requestPtr->iconIndex, _d2dContext.get());
        if (! bitmap)
        {
            return;
        }
    }

    size_t applied = 0;
    std::optional<size_t> firstAppliedIndex;
    for (const size_t itemIndex : requestPtr->itemIndices)
    {
        if (itemIndex >= _items.size())
        {
            continue;
        }

        auto& item = _items[itemIndex];

        // Verify icon index still matches (item might have changed)
        if (item.iconIndex != requestPtr->iconIndex || item.icon)
        {
            continue;
        }

        item.icon = bitmap;
        if (! firstAppliedIndex.has_value())
        {
            firstAppliedIndex = itemIndex;
        }
        ++applied;
    }

    if (applied == 0 || ! _hWnd)
    {
        return;
    }

    // For single-item updates, invalidate only that region. Otherwise invalidate the whole view.
    if (applied == 1 && firstAppliedIndex.has_value())
    {
        const auto idx = firstAppliedIndex.value();
        if (idx < _items.size())
        {
            const auto& item             = _items[idx];
            const D2D1_RECT_F viewBounds = OffsetRect(item.bounds, -_horizontalOffset, -_scrollOffset);
            RECT updateRect;
            updateRect.left   = PxFromDip(viewBounds.left);
            updateRect.top    = PxFromDip(viewBounds.top);
            updateRect.right  = PxFromDip(viewBounds.right);
            updateRect.bottom = PxFromDip(viewBounds.bottom);
            InvalidateRect(_hWnd.get(), &updateRect, FALSE);
            return;
        }
    }

    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void FolderView::OnBatchIconUpdate()
{
    if (_items.empty() || ! _d2dContext)
    {
        return;
    }

    Debug::Perf::Scope perf(L"FolderView.IconLoading.BatchUpdate");
    perf.SetDetail(_itemsFolder.native());
    perf.SetValue0(_items.size());

    size_t retrieved = 0;

    for (auto& item : _items)
    {
        // Skip if no valid icon index or already has icon
        if (item.iconIndex < 0 || item.icon)
        {
            continue;
        }

        // Try to get from cache
        auto bitmap = IconCache::GetInstance().GetCachedBitmap(item.iconIndex, _d2dContext.get());
        if (bitmap)
        {
            item.icon = bitmap;
            ++retrieved;
        }
    }

    // Invalidate entire view to redraw with new icons
    if (retrieved > 0 && _hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    perf.SetValue1(retrieved);
    MaybeEmitIconBitmapSummary(_iconLoadStats.batchId.load(std::memory_order_acquire));
}

void FolderView::MaybeEmitIconBitmapSummary(uint64_t batchId) noexcept
{
    if (_iconLoadStats.batchId.load(std::memory_order_acquire) != batchId)
    {
        return;
    }

    if (_iconLoadingActive.load(std::memory_order_acquire))
    {
        return;
    }

    if (_iconLoadStats.pendingBitmapCreates.load(std::memory_order_acquire) != 0)
    {
        return;
    }

    bool expected = false;
    if (! _iconLoadStats.bitmapSummaryEmitted.compare_exchange_strong(expected, true, std::memory_order_acq_rel))
    {
        return;
    }

    const int64_t firstPostQpc = _iconLoadStats.bitmapFirstPostQpc.load(std::memory_order_acquire);
    if (firstPostQpc == 0)
    {
        return;
    }

    LARGE_INTEGER frequency{};
    QueryPerformanceFrequency(&frequency);

    LARGE_INTEGER now{};
    QueryPerformanceCounter(&now);

    const int64_t freq    = frequency.QuadPart > 0 ? frequency.QuadPart : 1;
    const uint64_t wallUs = static_cast<uint64_t>(((now.QuadPart - firstPostQpc) * 1000000ll) / freq);

    const uint64_t converted   = _iconLoadStats.bitmapConverted.load(std::memory_order_relaxed);
    const uint64_t convertUs   = _iconLoadStats.bitmapConvertUsTotal.load(std::memory_order_relaxed);
    const uint64_t postFailed  = _iconLoadStats.bitmapPostFailed.load(std::memory_order_relaxed);
    const uint64_t convertFail = _iconLoadStats.bitmapConvertFailed.load(std::memory_order_relaxed);

    const HRESULT hr = (postFailed == 0 && convertFail == 0) ? S_OK : S_FALSE;
    Debug::Perf::Emit(L"FolderView.IconLoading.BitmapConversion", _itemsFolder.native(), wallUs, converted, convertUs, hr);
}

void FolderView::OnIconLoaded(size_t itemIndex)
{
    // This handles icons that were already cached (individual item notification)
    if (itemIndex >= _items.size() || ! _hWnd || ! _d2dContext)
    {
        return;
    }

    auto& item = _items[itemIndex];

    if (item.iconIndex < 0 || item.icon)
    {
        return; // Already has icon or invalid index
    }

    // Get from cache (already converted, just retrieve)
    auto bitmap = IconCache::GetInstance().GetCachedBitmap(item.iconIndex, _d2dContext.get());
    if (bitmap)
    {
        item.icon = bitmap;

        // Invalidate just the item's bounds for efficient redraw
        const D2D1_RECT_F viewBounds = OffsetRect(item.bounds, -_horizontalOffset, -_scrollOffset);
        RECT updateRect;
        updateRect.left   = PxFromDip(viewBounds.left);
        updateRect.top    = PxFromDip(viewBounds.top);
        updateRect.right  = PxFromDip(viewBounds.right);
        updateRect.bottom = PxFromDip(viewBounds.bottom);

        InvalidateRect(_hWnd.get(), &updateRect, FALSE);
    }
}
