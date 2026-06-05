#include "FolderViewInternal.h"
#ifdef ENABLE_TESTS
#include "SelfTestCommon.h"
#endif

#include <exception>

namespace
{
constexpr unsigned int kMaxIconLoadRetries      = 2u;
constexpr unsigned int kMaxThumbnailLoadRetries = 1u;
constexpr size_t kMaxThumbnailQueueItems        = 256u;
constexpr uint32_t kMaxThumbnailPixelSize       = 512u;
constexpr uint64_t kMaxThumbnailCacheBytes      = 64ull * 1024ull * 1024ull;

[[nodiscard]] uint64_t PerfElapsedUs(const std::chrono::steady_clock::time_point& start) noexcept
{
    const auto elapsed = std::chrono::steady_clock::now() - start;
    return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count());
}

void PerfEmitCounter(std::wstring_view name, uint64_t value) noexcept
{
    Debug::Perf::Emit(name, L"", 0, value, 0, S_OK);
}

void PerfEmitDuration(std::wstring_view name, uint64_t durationUs, uint64_t value0 = 0, uint64_t value1 = 0, HRESULT hr = S_OK) noexcept
{
    Debug::Perf::Emit(name, L"", durationUs, value0, value1, hr);
}

#ifdef ENABLE_TESTS
[[nodiscard]] wil::unique_hbitmap CreateSyntheticThumbnailBitmap(uint32_t targetPx, size_t itemIndex, bool wide) noexcept
{
    const uint32_t safeSize = std::clamp(targetPx, 1u, kMaxThumbnailPixelSize);
    const uint32_t width    = safeSize;
    const uint32_t height   = wide ? std::max(1u, safeSize / 2u) : safeSize;

    wil::unique_hdc_window screenDc{GetDC(nullptr)};
    if (! screenDc)
    {
        return nullptr;
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = static_cast<LONG>(width);
    bmi.bmiHeader.biHeight      = -static_cast<LONG>(height);
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    wil::unique_hbitmap bitmap{CreateDIBSection(screenDc.get(), &bmi, DIB_RGB_COLORS, &bits, nullptr, 0)};
    if (! bitmap || ! bits)
    {
        return nullptr;
    }

    auto* pixels          = static_cast<uint32_t*>(bits);
    const uint8_t red     = static_cast<uint8_t>(64u + ((itemIndex * 41u) % 160u));
    const uint8_t green   = static_cast<uint8_t>(72u + ((itemIndex * 29u) % 150u));
    const uint8_t blue    = static_cast<uint8_t>(96u + ((itemIndex * 17u) % 130u));
    const uint32_t fill   = 0xFF000000u | (static_cast<uint32_t>(red) << 16u) | (static_cast<uint32_t>(green) << 8u) | blue;
    const uint32_t accent = 0xFFFFFFFFu;

    for (uint32_t y = 0; y < height; ++y)
    {
        for (uint32_t x = 0; x < width; ++x)
        {
            const bool border                            = x == 0u || y == 0u || x + 1u == width || y + 1u == height;
            pixels[(static_cast<size_t>(y) * width) + x] = border ? accent : fill;
        }
    }

    return bitmap;
}
#endif

[[nodiscard]] HRESULT ExtractShellThumbnailBitmap(const std::filesystem::path& fullPath, uint32_t targetPx, wil::unique_hbitmap& outBitmap) noexcept
{
    outBitmap.reset();
    if (fullPath.empty() || targetPx == 0u)
    {
        return E_INVALIDARG;
    }

    wil::com_ptr<IShellItemImageFactory> imageFactory;
    HRESULT hr = SHCreateItemFromParsingName(fullPath.c_str(), nullptr, IID_PPV_ARGS(imageFactory.put()));
    if (FAILED(hr) || ! imageFactory)
    {
        return hr;
    }

    const uint32_t safeSize = std::clamp(targetPx, 1u, kMaxThumbnailPixelSize);
    const SIZE size{static_cast<LONG>(safeSize), static_cast<LONG>(safeSize)};
    HBITMAP rawBitmap      = nullptr;
    constexpr SIIGBF flags = static_cast<SIIGBF>(SIIGBF_THUMBNAILONLY | SIIGBF_BIGGERSIZEOK | SIIGBF_SCALEUP);
    hr                     = imageFactory->GetImage(size, flags, &rawBitmap);
    outBitmap.reset(rawBitmap);
    return hr;
}

[[nodiscard]] bool HasLikelyWicImageExtension(const std::filesystem::path& path) noexcept
{
    const std::wstring extension = path.extension().wstring();
    if (extension.empty())
    {
        return false;
    }

    return OrdinalString::EqualsNoCase(extension, L".bmp") || OrdinalString::EqualsNoCase(extension, L".dib") ||
           OrdinalString::EqualsNoCase(extension, L".gif") || OrdinalString::EqualsNoCase(extension, L".jpg") ||
           OrdinalString::EqualsNoCase(extension, L".jpeg") || OrdinalString::EqualsNoCase(extension, L".jpe") ||
           OrdinalString::EqualsNoCase(extension, L".png") || OrdinalString::EqualsNoCase(extension, L".tif") ||
           OrdinalString::EqualsNoCase(extension, L".tiff") || OrdinalString::EqualsNoCase(extension, L".webp") ||
           OrdinalString::EqualsNoCase(extension, L".heic") || OrdinalString::EqualsNoCase(extension, L".heif");
}

struct DecodedThumbnailPixels
{
    uint32_t widthPx     = 0;
    uint32_t heightPx    = 0;
    uint32_t strideBytes = 0;
    std::vector<uint8_t> bgra;

    [[nodiscard]] bool IsValid() const noexcept
    {
        return widthPx > 0u && heightPx > 0u && strideBytes >= widthPx * 4u && ! bgra.empty();
    }
};

[[nodiscard]] HRESULT DecodeWicThumbnailPixels(IWICImagingFactory* factory,
                                               const std::filesystem::path& fullPath,
                                               uint32_t targetPx,
                                               DecodedThumbnailPixels& outPixels) noexcept
{
    outPixels = {};
    if (! factory || fullPath.empty() || targetPx == 0u)
    {
        return E_INVALIDARG;
    }

    wil::com_ptr<IWICBitmapDecoder> decoder;
    HRESULT hr = factory->CreateDecoderFromFilename(fullPath.c_str(), nullptr, GENERIC_READ, WICDecodeMetadataCacheOnDemand, decoder.put());
    if (FAILED(hr) || ! decoder)
    {
        return hr;
    }

    wil::com_ptr<IWICBitmapFrameDecode> frame;
    hr = decoder->GetFrame(0u, frame.put());
    if (FAILED(hr) || ! frame)
    {
        return hr;
    }

    UINT sourceWidth  = 0u;
    UINT sourceHeight = 0u;
    hr                = frame->GetSize(&sourceWidth, &sourceHeight);
    if (FAILED(hr) || sourceWidth == 0u || sourceHeight == 0u)
    {
        return FAILED(hr) ? hr : WINCODEC_ERR_BADIMAGE;
    }

    const uint32_t safeTarget = std::clamp(targetPx, 1u, kMaxThumbnailPixelSize);
    const double scale        = std::min(
        1.0, std::min(static_cast<double>(safeTarget) / static_cast<double>(sourceWidth), static_cast<double>(safeTarget) / static_cast<double>(sourceHeight)));
    const UINT targetWidth  = std::max<UINT>(1u, static_cast<UINT>(std::round(static_cast<double>(sourceWidth) * scale)));
    const UINT targetHeight = std::max<UINT>(1u, static_cast<UINT>(std::round(static_cast<double>(sourceHeight) * scale)));

    wil::com_ptr<IWICBitmapSource> source;
    if (targetWidth != sourceWidth || targetHeight != sourceHeight)
    {
        wil::com_ptr<IWICBitmapScaler> scaler;
        hr = factory->CreateBitmapScaler(scaler.put());
        if (FAILED(hr) || ! scaler)
        {
            return hr;
        }

        hr = scaler->Initialize(frame.get(), targetWidth, targetHeight, WICBitmapInterpolationModeFant);
        if (FAILED(hr))
        {
            return hr;
        }
        source = scaler;
    }
    else
    {
        source = frame;
    }

    wil::com_ptr<IWICFormatConverter> converter;
    hr = factory->CreateFormatConverter(converter.put());
    if (FAILED(hr) || ! converter)
    {
        return hr;
    }

    hr = converter->Initialize(source.get(), GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, nullptr, 0.0, WICBitmapPaletteTypeCustom);
    if (FAILED(hr))
    {
        return hr;
    }

    const uint32_t stride      = targetWidth * 4u;
    const uint64_t byteCount64 = static_cast<uint64_t>(stride) * static_cast<uint64_t>(targetHeight);
    if (byteCount64 > static_cast<uint64_t>(std::numeric_limits<uint32_t>::max()))
    {
        return E_OUTOFMEMORY;
    }

    std::vector<uint8_t> pixels;
    pixels.resize(static_cast<size_t>(byteCount64));
    hr = converter->CopyPixels(nullptr, stride, static_cast<UINT>(pixels.size()), pixels.data());
    if (FAILED(hr))
    {
        return hr;
    }

    outPixels.widthPx     = targetWidth;
    outPixels.heightPx    = targetHeight;
    outPixels.strideBytes = stride;
    outPixels.bgra        = std::move(pixels);
    return S_OK;
}
} // namespace

void FolderView::QueueIconLoading()
{
#ifdef ENABLE_TESTS
    ++_debugQueueIconLoadingCallCount;
#endif
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

    Debug::Perf::Scope queuePerf(L"icons.queue_build_us");
    queuePerf.SetDetail(_itemsFolder.native());

    // Initialize telemetry (per-batch)
    const uint64_t batchId               = _iconLoadStats.batchId.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    const uint64_t enumerationGeneration = _enumerationGeneration.load(std::memory_order_acquire);
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
    uint64_t visibleGroups = 0;
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
        auto cachedBitmap = IconCache::GetInstance().GetCachedBitmap(iconIndex, _d2dContext.get(), _iconSizeDip);
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
        request.enumerationGeneration = enumerationGeneration;
        request.iconIndex             = iconIndex;
        request.targetDipSize         = _iconSizeDip;
        request.hasVisibleItems       = group.hasVisibleItems;
        request.firstVisibleItemIndex = group.firstVisibleItemIndex;
        request.enqueuedAt            = std::chrono::steady_clock::now();
        request.itemIndices           = std::move(group.itemIndices);

        if (request.hasVisibleItems)
        {
            ++visibleGroups;
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
    PerfEmitCounter(L"icons.queue_groups_total", static_cast<uint64_t>(groups.size()));
    PerfEmitCounter(L"icons.queue_visible_groups", visibleGroups);
    PerfEmitCounter(L"icons.cache_stamp_count", stampedFromCache);

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

        if (auto cached = IconCache::GetInstance().GetCachedBitmap(item.iconIndex, _d2dContext.get(), _iconSizeDip))
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

    if (_thumbnailsVisible)
    {
        QueueThumbnailLoading();
    }
}

void FolderView::ProcessIconLoadQueue()
{
#ifdef ENABLE_TESTS
    ++_debugProcessIconQueueCallCount;
#endif
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
                Debug::Info(L"FolderView: IconCache stats - {} cached icons (~{} MB), {} hits, {} misses, {} LRU evictions, {} path cached, {} path evictions",
                            cacheStats.cacheSize,
                            cacheMemoryMB,
                            cacheStats.hitCount,
                            cacheStats.missCount,
                            cacheStats.lruEvictions,
                            cacheStats.pathCacheSize,
                            cacheStats.pathLruEvictions);
                break;
            }

            request = std::move(_iconLoadQueue.front());
            _iconLoadQueue.pop_front(); // O(1) with deque vs O(N) with vector
        }

        if (request.iconIndex < 0 || request.itemIndices.empty())
        {
            continue;
        }

        if (request.enqueuedAt.time_since_epoch().count() > 0)
        {
            PerfEmitDuration(L"icons.queue_wait_to_dequeue_us",
                             PerfElapsedUs(request.enqueuedAt),
                             static_cast<uint64_t>(request.iconIndex),
                             static_cast<uint64_t>(request.itemIndices.size()),
                             S_OK);
        }

        wil::com_ptr<ID2D1Device> d2dDeviceSnapshot;
        {
            std::lock_guard lock(_d2dDeviceMutex);
            d2dDeviceSnapshot = _d2dDevice;
        }
        const bool cachedForDevice =
            d2dDeviceSnapshot && IconCache::GetInstance().HasCachedIcon(request.iconIndex, d2dDeviceSnapshot.get(), request.targetDipSize);

        auto bitmapRequest                   = std::make_unique<IconBitmapRequest>();
        bitmapRequest->iconLoadBatchId       = batchId;
        bitmapRequest->enumerationGeneration = request.enumerationGeneration;
        bitmapRequest->iconIndex             = request.iconIndex;
        bitmapRequest->targetDipSize         = request.targetDipSize;
        bitmapRequest->postedAt              = std::chrono::steady_clock::now();
        bitmapRequest->itemIndices           = std::move(request.itemIndices);

        // Background thread: extract once per iconIndex (unless already cached).
        if (! cachedForDevice)
        {
            const auto extractStart = std::chrono::steady_clock::now();
            wil::unique_hicon hIcon = IconCache::GetInstance().ExtractSystemIcon(request.iconIndex, request.targetDipSize);
            PerfEmitDuration(L"icons.extract_us",
                             PerfElapsedUs(extractStart),
                             static_cast<uint64_t>(request.iconIndex),
                             static_cast<uint64_t>(request.retryCount),
                             hIcon ? S_OK : S_FALSE);
            if (! hIcon)
            {
                if (request.retryCount < kMaxIconLoadRetries)
                {
                    ++request.retryCount;
                    PerfEmitCounter(L"icons.extract_retries", 1);
                    {
                        std::lock_guard lock(_enumerationMutex);
                        if (request.hasVisibleItems)
                        {
                            _iconLoadQueue.push_front(std::move(request));
                        }
                        else
                        {
                            _iconLoadQueue.push_back(std::move(request));
                        }
                    }
                    _enumerationCv.notify_one();
                }
                else
                {
                    PerfEmitCounter(L"icons.extract_failures", 1);
                    Debug::Warning(L"FolderView: Failed to extract icon index {} after {} attempts", request.iconIndex, request.retryCount + 1u);
                }
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

void FolderView::QueueThumbnailLoading()
{
    if (! _thumbnailsVisible || _items.empty() || ! _hWnd)
    {
        return;
    }

    Debug::Perf::Scope queuePerf(L"thumbnails.queue_build_us");
    queuePerf.SetDetail(_itemsFolder.native());

    const uint64_t batchId               = _thumbnailLoadStats.batchId.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    const uint64_t enumerationGeneration = _enumerationGeneration.load(std::memory_order_acquire);
    const float targetDip                = _iconSizeDip;
    const uint32_t targetPx              = static_cast<uint32_t>(std::max(1, PxFromDip(targetDip)));

    _thumbnailLoadStats.queued.store(0u, std::memory_order_release);
    _thumbnailLoadStats.completed.store(0u, std::memory_order_release);
    _thumbnailLoadStats.fallback.store(0u, std::memory_order_release);
    _thumbnailLoadStats.staleDrops.store(0u, std::memory_order_release);
    _thumbnailLoadStats.pendingBitmapCreates.store(0u, std::memory_order_release);
    _thumbnailLoadStats.cacheHits.store(0u, std::memory_order_release);
    _thumbnailLoadStats.shellSuccess.store(0u, std::memory_order_release);
    _thumbnailLoadStats.wicSuccess.store(0u, std::memory_order_release);
    _thumbnailLoadStats.decodeFailures.store(0u, std::memory_order_release);
    _thumbnailLoadStats.visibleApply.store(0u, std::memory_order_release);
    _thumbnailLoadStats.cacheBytes.store(0u, std::memory_order_release);
    _thumbnailLoadStats.cacheEvicted.store(0u, std::memory_order_release);
    _thumbnailLoadStats.targetDipX100.store(static_cast<uint64_t>(std::round(targetDip * 100.0f)), std::memory_order_release);

    const auto [visibleStartRaw, visibleEndRaw] = GetVisibleItemRange();
    const size_t visibleStart                   = std::min(visibleStartRaw, _items.size());
    const size_t visibleEnd                     = std::min(std::max(visibleEndRaw, visibleStart), _items.size());

    uint64_t cacheBytes = 0u;
    for (const FolderItem& item : _items)
    {
        if (! item.thumbnail)
        {
            continue;
        }

        const D2D1_SIZE_U pixelSize = item.thumbnail->GetPixelSize();
        cacheBytes += static_cast<uint64_t>(pixelSize.width) * static_cast<uint64_t>(pixelSize.height) * 4u;
    }

    uint64_t evictedCount = 0u;
    if (cacheBytes > kMaxThumbnailCacheBytes)
    {
        for (size_t index = 0; index < _items.size() && cacheBytes > kMaxThumbnailCacheBytes; ++index)
        {
            if (index >= visibleStart && index < visibleEnd)
            {
                continue;
            }

            FolderItem& item = _items[index];
            if (! item.thumbnail)
            {
                continue;
            }

            const D2D1_SIZE_U pixelSize = item.thumbnail->GetPixelSize();
            const uint64_t itemBytes    = static_cast<uint64_t>(pixelSize.width) * static_cast<uint64_t>(pixelSize.height) * 4u;
            item.thumbnail.reset();
            cacheBytes = itemBytes >= cacheBytes ? 0u : cacheBytes - itemBytes;
            ++evictedCount;
        }
    }
    _thumbnailLoadStats.cacheBytes.store(cacheBytes, std::memory_order_release);
    if (evictedCount > 0u)
    {
        _thumbnailLoadStats.cacheEvicted.fetch_add(evictedCount, std::memory_order_relaxed);
        PerfEmitCounter(L"thumbnails.cache_evicted_count", evictedCount);
    }
    PerfEmitCounter(L"thumbnails.cache_bytes", cacheBytes);

    std::deque<ThumbnailLoadRequest> newQueue;
    const size_t reserveLimit = std::min(kMaxThumbnailQueueItems, visibleEnd - visibleStart);
    static_cast<void>(reserveLimit);

    for (size_t index = visibleStart; index < visibleEnd && newQueue.size() < kMaxThumbnailQueueItems; ++index)
    {
        const FolderItem& item = _items[index];
        if (item.thumbnail)
        {
            _thumbnailLoadStats.cacheHits.fetch_add(1u, std::memory_order_relaxed);
            continue;
        }

        ThumbnailLoadRequest request;
        request.enumerationGeneration = enumerationGeneration;
        request.itemIndex             = index;
        request.fullPath              = GetItemFullPath(item);
        request.targetPx              = targetPx;
        request.hasVisibleItem        = true;
        request.enqueuedAt            = std::chrono::steady_clock::now();
        newQueue.push_back(std::move(request));
    }

    const uint64_t queuedCount = static_cast<uint64_t>(newQueue.size());
    {
        std::lock_guard lock(_enumerationMutex);
        _thumbnailLoadQueue = std::move(newQueue);
    }

    _thumbnailLoadStats.queued.store(queuedCount, std::memory_order_release);
    PerfEmitCounter(L"thumbnails.queue_count", queuedCount);
    PerfEmitCounter(L"thumbnails.queue_visible_count", queuedCount);
    PerfEmitCounter(L"thumbnails.cache_hit_count", _thumbnailLoadStats.cacheHits.load(std::memory_order_relaxed));
    queuePerf.SetValue0(queuedCount);
    queuePerf.SetValue1(static_cast<uint64_t>(targetPx));

    if (queuedCount > 0u)
    {
        _thumbnailLoadingActive.store(true, std::memory_order_release);
        _enumerationCv.notify_one();
    }

    static_cast<void>(batchId);
}

bool FolderView::HasMissingVisibleThumbnails() const
{
    if (! _thumbnailsVisible || _items.empty())
    {
        return false;
    }

    const auto [visibleStartRaw, visibleEndRaw] = GetVisibleItemRange();
    const size_t visibleStart                   = std::min(visibleStartRaw, _items.size());
    const size_t visibleEnd                     = std::min(std::max(visibleEndRaw, visibleStart), _items.size());
    for (size_t index = visibleStart; index < visibleEnd; ++index)
    {
        if (! _items[index].thumbnail)
        {
            return true;
        }
    }

    return false;
}

void FolderView::QueueMissingVisibleThumbnails()
{
    if (_thumbnailsVisible && _d2dContext && HasMissingVisibleThumbnails())
    {
        QueueThumbnailLoading();
    }
}

void FolderView::CancelThumbnailLoading() noexcept
{
    _thumbnailLoadStats.batchId.fetch_add(1u, std::memory_order_acq_rel);
    {
        std::lock_guard lock(_enumerationMutex);
        _thumbnailLoadQueue.clear();
    }

    _thumbnailLoadingActive.store(false, std::memory_order_release);
    _thumbnailLoadStats.pendingBitmapCreates.store(0u, std::memory_order_release);
    _thumbnailLoadStats.cancelCount.fetch_add(1u, std::memory_order_relaxed);
    _enumerationCv.notify_one();
}

HRESULT FolderView::EnsureThumbnailWicFactory(wil::com_ptr<IWICImagingFactory>& thumbnailWicFactory, IWICImagingFactory** outFactory) noexcept
{
    if (! outFactory)
    {
        return E_POINTER;
    }
    *outFactory = nullptr;

    if (! thumbnailWicFactory)
    {
        APTTYPE apartmentType{};
        APTTYPEQUALIFIER apartmentQualifier{};
        const HRESULT apartmentHr = CoGetApartmentType(&apartmentType, &apartmentQualifier);
        if (FAILED(apartmentHr))
        {
            Debug::Error(L"FolderView thumbnail WIC fallback requires COM on the worker thread: 0x{:08X}", apartmentHr);
            return apartmentHr;
        }
        if (apartmentType != APTTYPE_MTA)
        {
            Debug::Warning(L"FolderView thumbnail WIC fallback expected worker MTA, apartment={}", static_cast<int>(apartmentType));
        }

        wil::com_ptr<IWICImagingFactory> factory;
        HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.put()));
        if (FAILED(hr) || ! factory)
        {
            return hr;
        }

        thumbnailWicFactory = std::move(factory);
        _thumbnailLoadStats.wicFactoryCreate.fetch_add(1u, std::memory_order_relaxed);
    }

    *outFactory = thumbnailWicFactory.get();
    return S_OK;
}

void FolderView::ProcessThumbnailLoadQueue()
{
    const uint64_t batchId = _thumbnailLoadStats.batchId.load(std::memory_order_acquire);
    Debug::Perf::Scope perf(L"FolderView.ThumbnailLoading.ProcessQueue");
    perf.SetDetail(_itemsFolder.native());
    perf.SetValue0(_thumbnailLoadStats.queued.load(std::memory_order_relaxed));

    wil::com_ptr<IWICImagingFactory> thumbnailWicFactory;
    while (_thumbnailLoadingActive.load(std::memory_order_acquire))
    {
        if (_thumbnailLoadStats.batchId.load(std::memory_order_acquire) != batchId)
        {
            break;
        }

        ThumbnailLoadRequest request{};
        {
            std::lock_guard lock(_enumerationMutex);
            if (_thumbnailLoadQueue.empty())
            {
                _thumbnailLoadingActive.store(false, std::memory_order_release);
                break;
            }

            request = std::move(_thumbnailLoadQueue.front());
            _thumbnailLoadQueue.pop_front();
        }

        if (request.itemIndex == static_cast<size_t>(-1))
        {
            continue;
        }

        if (request.enumerationGeneration != _enumerationGeneration.load(std::memory_order_acquire))
        {
            _thumbnailLoadStats.staleDrops.fetch_add(1u, std::memory_order_relaxed);
            continue;
        }

        if (request.enqueuedAt.time_since_epoch().count() > 0)
        {
            PerfEmitDuration(L"thumbnails.queue_wait_to_dequeue_us",
                             PerfElapsedUs(request.enqueuedAt),
                             static_cast<uint64_t>(request.itemIndex),
                             static_cast<uint64_t>(request.targetPx),
                             S_OK);
        }

        auto bitmapRequest                   = std::make_unique<ThumbnailBitmapRequest>();
        bitmapRequest->thumbnailLoadBatchId  = batchId;
        bitmapRequest->enumerationGeneration = request.enumerationGeneration;
        bitmapRequest->itemIndex             = request.itemIndex;
        bitmapRequest->postedAt              = std::chrono::steady_clock::now();

        const auto extractStart = std::chrono::steady_clock::now();
        HRESULT hr              = S_FALSE;
        bool usedFallback       = false;
        bool allowWicFallback   = false;
        bool triedWicFallback   = false;

#ifdef ENABLE_TESTS
        const DebugThumbnailProviderMode providerMode = _debugThumbnailProviderMode.load(std::memory_order_acquire);
        if (providerMode == DebugThumbnailProviderMode::ForceFallback)
        {
            usedFallback = true;
            hr           = S_FALSE;
        }
        else if (providerMode == DebugThumbnailProviderMode::ForceSyntheticSuccess)
        {
            bitmapRequest->hBitmap    = CreateSyntheticThumbnailBitmap(request.targetPx, request.itemIndex, false);
            hr                        = bitmapRequest->hBitmap ? S_OK : S_FALSE;
            usedFallback              = ! bitmapRequest->hBitmap;
            bitmapRequest->sourceKind = bitmapRequest->hBitmap ? ThumbnailBitmapRequest::SourceKind::Synthetic : ThumbnailBitmapRequest::SourceKind::Fallback;
        }
        else if (providerMode == DebugThumbnailProviderMode::ForceSyntheticWideSuccess)
        {
            bitmapRequest->hBitmap    = CreateSyntheticThumbnailBitmap(request.targetPx, request.itemIndex, true);
            hr                        = bitmapRequest->hBitmap ? S_OK : S_FALSE;
            usedFallback              = ! bitmapRequest->hBitmap;
            bitmapRequest->sourceKind = bitmapRequest->hBitmap ? ThumbnailBitmapRequest::SourceKind::Synthetic : ThumbnailBitmapRequest::SourceKind::Fallback;
        }
        else if (providerMode == DebugThumbnailProviderMode::ForceShellFailureAllowWic)
        {
            hr               = HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
            allowWicFallback = true;
        }
        else
#endif
        {
            wil::unique_hbitmap shellBitmap;
            hr = ExtractShellThumbnailBitmap(request.fullPath, request.targetPx, shellBitmap);
            if (SUCCEEDED(hr) && shellBitmap)
            {
                bitmapRequest->hBitmap    = std::move(shellBitmap);
                bitmapRequest->sourceKind = ThumbnailBitmapRequest::SourceKind::Shell;
                _thumbnailLoadStats.shellSuccess.fetch_add(1u, std::memory_order_relaxed);
            }
            else
            {
                allowWicFallback = true;
            }
        }

        if (allowWicFallback && HasLikelyWicImageExtension(request.fullPath))
        {
            triedWicFallback    = true;
            const auto wicStart = std::chrono::steady_clock::now();
            DecodedThumbnailPixels pixels;
            IWICImagingFactory* wicFactory = nullptr;
            HRESULT wicHr                  = EnsureThumbnailWicFactory(thumbnailWicFactory, &wicFactory);
            if (SUCCEEDED(wicHr))
            {
                wicHr = DecodeWicThumbnailPixels(wicFactory, request.fullPath, request.targetPx, pixels);
            }
            PerfEmitDuration(
                L"thumbnails.wic_decode_us", PerfElapsedUs(wicStart), static_cast<uint64_t>(request.itemIndex), static_cast<uint64_t>(request.targetPx), wicHr);
            if (SUCCEEDED(wicHr) && pixels.IsValid())
            {
                bitmapRequest->pixels.widthPx     = pixels.widthPx;
                bitmapRequest->pixels.heightPx    = pixels.heightPx;
                bitmapRequest->pixels.strideBytes = pixels.strideBytes;
                bitmapRequest->pixels.bgra        = std::move(pixels.bgra);
                bitmapRequest->sourceKind         = ThumbnailBitmapRequest::SourceKind::Wic;
                _thumbnailLoadStats.wicSuccess.fetch_add(1u, std::memory_order_relaxed);
                hr           = S_OK;
                usedFallback = false;
            }
            else
            {
                hr           = FAILED(wicHr) ? wicHr : S_FALSE;
                usedFallback = true;
            }
        }
        else if (allowWicFallback)
        {
            usedFallback = true;
        }

        if (usedFallback)
        {
            bitmapRequest->sourceKind = ThumbnailBitmapRequest::SourceKind::Fallback;
            if (triedWicFallback)
            {
                _thumbnailLoadStats.decodeFailures.fetch_add(1u, std::memory_order_relaxed);
            }
        }

        bitmapRequest->hr           = hr;
        bitmapRequest->usedFallback = usedFallback;
        PerfEmitDuration(
            L"thumbnails.extract_us", PerfElapsedUs(extractStart), static_cast<uint64_t>(request.itemIndex), static_cast<uint64_t>(request.targetPx), hr);

        if (! bitmapRequest->hBitmap && ! bitmapRequest->pixels.IsValid() && ! usedFallback && request.retryCount < kMaxThumbnailLoadRetries)
        {
            ++request.retryCount;
            {
                std::lock_guard lock(_enumerationMutex);
                _thumbnailLoadQueue.push_back(std::move(request));
            }
            _enumerationCv.notify_one();
            continue;
        }

        if (! _hWnd)
        {
            _thumbnailLoadStats.completed.fetch_add(1u, std::memory_order_relaxed);
            if (usedFallback)
            {
                _thumbnailLoadStats.fallback.fetch_add(1u, std::memory_order_relaxed);
            }
            continue;
        }

        _thumbnailLoadStats.pendingBitmapCreates.fetch_add(1u, std::memory_order_acq_rel);
        const bool posted = PostMessagePayload(_hWnd.get(), WndMsg::kFolderViewCreateThumbnailBitmap, 0, std::move(bitmapRequest));
        if (! posted)
        {
            const uint64_t before = _thumbnailLoadStats.pendingBitmapCreates.load(std::memory_order_acquire);
            if (before > 0u)
            {
                _thumbnailLoadStats.pendingBitmapCreates.fetch_sub(1u, std::memory_order_acq_rel);
            }
            _thumbnailLoadStats.completed.fetch_add(1u, std::memory_order_relaxed);
            if (usedFallback)
            {
                _thumbnailLoadStats.fallback.fetch_add(1u, std::memory_order_relaxed);
            }
        }
    }

    perf.SetValue1(_thumbnailLoadStats.completed.load(std::memory_order_relaxed));
}

void FolderView::OnCreateIconBitmap(std::unique_ptr<IconBitmapRequest> requestPtr)
{
    // This runs on UI thread - safe to use _d2dContext
    if (! requestPtr)
    {
        return;
    }

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"FolderView::OnCreateIconBitmap: begin requestGeneration={} iconIndex={} itemCount={}",
                                              requestPtr->enumerationGeneration,
                                              requestPtr->iconIndex,
                                              requestPtr->itemIndices.size()));
#endif

    const uint64_t batchId = _iconLoadStats.batchId.load(std::memory_order_acquire);
    if (requestPtr->iconLoadBatchId != batchId)
    {
        return;
    }

    const uint64_t enumerationGeneration = _enumerationGeneration.load(std::memory_order_acquire);
    if (requestPtr->enumerationGeneration != enumerationGeneration)
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(
            std::format(L"FolderView::OnCreateIconBitmap: dropped stale payload requestGeneration={} currentGeneration={} iconIndex={}",
                        requestPtr->enumerationGeneration,
                        enumerationGeneration,
                        requestPtr->iconIndex));
#endif
        return;
    }

    const auto onExit = wil::scope_exit([&]() noexcept
    {
        const uint64_t remaining = _iconLoadStats.pendingBitmapCreates.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        static_cast<void>(remaining);
        MaybeEmitIconBitmapSummary(batchId);
    });

    if (! _d2dContext || requestPtr->iconIndex < 0 || requestPtr->itemIndices.empty())
    {
        return;
    }

    if (requestPtr->postedAt.time_since_epoch().count() > 0)
    {
        PerfEmitDuration(L"icons.post_message_latency_us",
                         PerfElapsedUs(requestPtr->postedAt),
                         static_cast<uint64_t>(requestPtr->iconIndex),
                         static_cast<uint64_t>(requestPtr->itemIndices.size()),
                         S_OK);
    }

    wil::com_ptr<ID2D1Bitmap1> bitmap;
    if (requestPtr->hIcon)
    {
        // Convert HICON to D2D bitmap on UI thread (thread-safe)
        const auto convertStart = std::chrono::steady_clock::now();
        try
        {
            bitmap = IconCache::GetInstance().ConvertIconToBitmapOnUIThread(
                requestPtr->hIcon.get(), requestPtr->iconIndex, _d2dContext.get(), requestPtr->targetDipSize);
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
        catch (const std::exception& ex)
        {
            // This runs from a posted Win32 callback; recoverable icon conversion failures must not escape WndProc.
            Debug::Warning(L"FolderView: icon bitmap conversion threw for icon index {}", requestPtr->iconIndex);
            static_cast<void>(ex);
#ifdef ENABLE_TESTS
            SelfTest::AppendSelfTestTrace(std::format(L"FolderView::OnCreateIconBitmap: conversion exception iconIndex={}", requestPtr->iconIndex));
#endif
            bitmap = nullptr;
        }
        const uint64_t convertUs = PerfElapsedUs(convertStart);
        PerfEmitDuration(L"icons.ui_convert_us",
                         convertUs,
                         static_cast<uint64_t>(requestPtr->iconIndex),
                         static_cast<uint64_t>(requestPtr->itemIndices.size()),
                         bitmap ? S_OK : S_FALSE);
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
        bitmap = IconCache::GetInstance().GetCachedBitmap(requestPtr->iconIndex, _d2dContext.get(), requestPtr->targetDipSize);
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

    PerfEmitCounter(L"icons.ui_apply_count", static_cast<uint64_t>(applied));

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
            PerfEmitCounter(L"icons.invalidate_single_count", 1);
            PerfEmitCounter(L"icons.invalidate_area_px",
                            static_cast<uint64_t>(std::max<LONG>(0, updateRect.right - updateRect.left)) *
                                static_cast<uint64_t>(std::max<LONG>(0, updateRect.bottom - updateRect.top)));
            InvalidateRect(_hWnd.get(), &updateRect, FALSE);
            return;
        }
    }

    PerfEmitCounter(L"icons.invalidate_full_count", 1);
    PerfEmitCounter(L"icons.invalidate_area_px", static_cast<uint64_t>(_clientSize.cx) * static_cast<uint64_t>(_clientSize.cy));
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"FolderView::OnCreateIconBitmap: end applied={} iconIndex={}", applied, requestPtr->iconIndex));
#endif
}

void FolderView::OnCreateThumbnailBitmap(std::unique_ptr<ThumbnailBitmapRequest> requestPtr)
{
    if (! requestPtr)
    {
        return;
    }

    const auto onExit = wil::scope_exit([&]() noexcept
    {
        const uint64_t before = _thumbnailLoadStats.pendingBitmapCreates.load(std::memory_order_acquire);
        if (before > 0u)
        {
            _thumbnailLoadStats.pendingBitmapCreates.fetch_sub(1u, std::memory_order_acq_rel);
        }
    });

    const uint64_t batchId = _thumbnailLoadStats.batchId.load(std::memory_order_acquire);
    if (requestPtr->thumbnailLoadBatchId != batchId)
    {
        _thumbnailLoadStats.staleDrops.fetch_add(1u, std::memory_order_relaxed);
        return;
    }

    const uint64_t enumerationGeneration = _enumerationGeneration.load(std::memory_order_acquire);
    if (requestPtr->enumerationGeneration != enumerationGeneration)
    {
        _thumbnailLoadStats.staleDrops.fetch_add(1u, std::memory_order_relaxed);
        return;
    }

    if (requestPtr->postedAt.time_since_epoch().count() > 0)
    {
        PerfEmitDuration(
            L"thumbnails.post_message_latency_us", PerfElapsedUs(requestPtr->postedAt), static_cast<uint64_t>(requestPtr->itemIndex), 1u, requestPtr->hr);
    }

    const auto completeAsFallback = [&]() noexcept
    {
        _thumbnailLoadStats.completed.fetch_add(1u, std::memory_order_relaxed);
        _thumbnailLoadStats.fallback.fetch_add(1u, std::memory_order_relaxed);
        PerfEmitCounter(L"thumbnails.fallback_count", 1u);
    };

    if (requestPtr->usedFallback || FAILED(requestPtr->hr) || (! requestPtr->hBitmap && ! requestPtr->pixels.IsValid()))
    {
        completeAsFallback();
        return;
    }

    if (! _d2dContext)
    {
        completeAsFallback();
        return;
    }

    const auto convertStart = std::chrono::steady_clock::now();
    wil::com_ptr<ID2D1Bitmap1> bitmap;
    HRESULT hr = S_OK;
    const D2D1_BITMAP_PROPERTIES1 bitmapProps =
        D2D1::BitmapProperties1(D2D1_BITMAP_OPTIONS_NONE, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), _dpi, _dpi);

    if (requestPtr->pixels.IsValid())
    {
        hr = _d2dContext->CreateBitmap(D2D1::SizeU(requestPtr->pixels.widthPx, requestPtr->pixels.heightPx),
                                       requestPtr->pixels.bgra.data(),
                                       requestPtr->pixels.strideBytes,
                                       &bitmapProps,
                                       bitmap.put());
    }
    else
    {
        if (! _wicFactory)
        {
            EnsureDeviceIndependentResources();
        }

        if (! _wicFactory)
        {
            completeAsFallback();
            return;
        }

        wil::com_ptr<IWICBitmap> wicBitmap;
        hr = _wicFactory->CreateBitmapFromHBITMAP(requestPtr->hBitmap.get(), nullptr, WICBitmapUsePremultipliedAlpha, wicBitmap.put());
        if (SUCCEEDED(hr) && wicBitmap)
        {
            hr = _d2dContext->CreateBitmapFromWicBitmap(wicBitmap.get(), &bitmapProps, bitmap.put());
        }
    }

    PerfEmitDuration(L"thumbnails.ui_convert_us", PerfElapsedUs(convertStart), static_cast<uint64_t>(requestPtr->itemIndex), 1u, hr);
    if (FAILED(hr) || ! bitmap)
    {
        completeAsFallback();
        return;
    }

    if (requestPtr->itemIndex >= _items.size())
    {
        _thumbnailLoadStats.staleDrops.fetch_add(1u, std::memory_order_relaxed);
        return;
    }

    _items[requestPtr->itemIndex].thumbnail = std::move(bitmap);
    _thumbnailLoadStats.completed.fetch_add(1u, std::memory_order_relaxed);
    _thumbnailLoadStats.visibleApply.fetch_add(1u, std::memory_order_relaxed);
    PerfEmitCounter(L"thumbnails.ui_apply_count", 1u);

    if (! _hWnd)
    {
        return;
    }

    const auto& item             = _items[requestPtr->itemIndex];
    const D2D1_RECT_F viewBounds = OffsetRect(item.bounds, -_horizontalOffset, -_scrollOffset);
    RECT updateRect;
    updateRect.left   = PxFromDip(viewBounds.left);
    updateRect.top    = PxFromDip(viewBounds.top);
    updateRect.right  = PxFromDip(viewBounds.right);
    updateRect.bottom = PxFromDip(viewBounds.bottom);
    InvalidateRect(_hWnd.get(), &updateRect, FALSE);
}

#ifdef ENABLE_TESTS
FolderView::ThumbnailDebugSnapshot FolderView::DebugGetThumbnailSnapshot() const noexcept
{
    ThumbnailDebugSnapshot snapshot{};
    snapshot.visible                            = _thumbnailsVisible;
    snapshot.targetDip                          = static_cast<float>(_thumbnailLoadStats.targetDipX100.load(std::memory_order_acquire)) / 100.0f;
    snapshot.queuedCount                        = _thumbnailLoadStats.queued.load(std::memory_order_acquire);
    snapshot.completedCount                     = _thumbnailLoadStats.completed.load(std::memory_order_acquire);
    snapshot.fallbackCount                      = _thumbnailLoadStats.fallback.load(std::memory_order_acquire);
    snapshot.staleDropCount                     = _thumbnailLoadStats.staleDrops.load(std::memory_order_acquire);
    snapshot.pendingCount                       = _thumbnailLoadStats.pendingBitmapCreates.load(std::memory_order_acquire);
    snapshot.cacheHitCount                      = _thumbnailLoadStats.cacheHits.load(std::memory_order_acquire);
    snapshot.shellSuccessCount                  = _thumbnailLoadStats.shellSuccess.load(std::memory_order_acquire);
    snapshot.wicSuccessCount                    = _thumbnailLoadStats.wicSuccess.load(std::memory_order_acquire);
    snapshot.wicFactoryCreateCount              = _thumbnailLoadStats.wicFactoryCreate.load(std::memory_order_acquire);
    snapshot.decodeFailureCount                 = _thumbnailLoadStats.decodeFailures.load(std::memory_order_acquire);
    snapshot.visibleApplyCount                  = _thumbnailLoadStats.visibleApply.load(std::memory_order_acquire);
    const auto [visibleStartRaw, visibleEndRaw] = GetVisibleItemRange();
    const size_t visibleStart                   = std::min(visibleStartRaw, _items.size());
    const size_t visibleEnd                     = std::min(std::max(visibleEndRaw, visibleStart), _items.size());
    snapshot.visibleItemCount                   = static_cast<uint64_t>(visibleEnd - visibleStart);
    for (size_t index = 0; index < _items.size(); ++index)
    {
        if (! _items[index].thumbnail)
        {
            continue;
        }
        ++snapshot.totalThumbnailCount;
        if (index >= visibleStart && index < visibleEnd)
        {
            ++snapshot.visibleThumbnailCount;
        }
    }
    uint64_t cacheBytes = _thumbnailLoadStats.cacheBytes.load(std::memory_order_acquire);
    if (cacheBytes == 0u)
    {
        for (const FolderItem& item : _items)
        {
            if (! item.thumbnail)
            {
                continue;
            }

            const D2D1_SIZE_U pixelSize = item.thumbnail->GetPixelSize();
            cacheBytes += static_cast<uint64_t>(pixelSize.width) * static_cast<uint64_t>(pixelSize.height) * 4u;
        }
    }
    snapshot.cacheBytes                 = cacheBytes;
    snapshot.cacheEvictedCount          = _thumbnailLoadStats.cacheEvicted.load(std::memory_order_acquire);
    snapshot.cancelCount                = _thumbnailLoadStats.cancelCount.load(std::memory_order_acquire);
    snapshot.lastDrawSawThumbnail       = _debugLastThumbnailDrawSawThumbnail;
    snapshot.lastDrawSourceWidthPx      = _debugLastThumbnailSourceWidthPx;
    snapshot.lastDrawSourceHeightPx     = _debugLastThumbnailSourceHeightPx;
    snapshot.lastDrawSlotRectDip        = _debugLastThumbnailSlotRectDip;
    snapshot.lastDrawRectDip            = _debugLastThumbnailDrawRectDip;
    snapshot.lastIconDrawSawIcon        = _debugLastIconDrawSawIcon;
    snapshot.lastIconDrawSourceWidthPx  = _debugLastIconDrawSourceWidthPx;
    snapshot.lastIconDrawSourceHeightPx = _debugLastIconDrawSourceHeightPx;
    snapshot.lastIconDrawSlotRectDip    = _debugLastIconDrawSlotRectDip;
    snapshot.lastIconDrawRectDip        = _debugLastIconDrawRectDip;
    return snapshot;
}

void FolderView::DebugSetThumbnailProviderMode(DebugThumbnailProviderMode mode) noexcept
{
    _debugThumbnailProviderMode.store(mode, std::memory_order_release);
}
#endif

void FolderView::OnBatchIconUpdate()
{
#ifdef ENABLE_TESTS
    ++_debugBatchIconUpdateCallCount;
#endif
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"FolderView::OnBatchIconUpdate: begin itemCount={}", _items.size()));
#endif
    if (_items.empty() || ! _d2dContext)
    {
        return;
    }

    Debug::Perf::Scope perf(L"FolderView.IconLoading.BatchUpdate");
    perf.SetDetail(_itemsFolder.native());
    perf.SetValue0(_items.size());

    const auto scanStart = std::chrono::steady_clock::now();
    size_t retrieved     = 0;

    for (auto& item : _items)
    {
        // Skip if no valid icon index or already has icon
        if (item.iconIndex < 0 || item.icon)
        {
            continue;
        }

        // Try to get from cache
        auto bitmap = IconCache::GetInstance().GetCachedBitmap(item.iconIndex, _d2dContext.get(), _iconSizeDip);
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

    PerfEmitDuration(L"icons.batch_update_scan_us", PerfElapsedUs(scanStart), static_cast<uint64_t>(_items.size()), static_cast<uint64_t>(retrieved), S_OK);
    PerfEmitCounter(L"icons.batch_update_retrieved", static_cast<uint64_t>(retrieved));
    perf.SetValue1(retrieved);
    MaybeEmitIconBitmapSummary(_iconLoadStats.batchId.load(std::memory_order_acquire));
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"FolderView::OnBatchIconUpdate: end retrieved={}", retrieved));
#endif
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
    auto bitmap = IconCache::GetInstance().GetCachedBitmap(item.iconIndex, _d2dContext.get(), _iconSizeDip);
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

        PerfEmitCounter(L"icons.invalidate_single_count", 1);
        PerfEmitCounter(L"icons.invalidate_area_px",
                        static_cast<uint64_t>(std::max<LONG>(0, updateRect.right - updateRect.left)) *
                            static_cast<uint64_t>(std::max<LONG>(0, updateRect.bottom - updateRect.top)));
        InvalidateRect(_hWnd.get(), &updateRect, FALSE);
    }
}
