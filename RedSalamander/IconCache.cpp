
#include <algorithm>
#include <span>
#include <vector>

#include <cwctype>

#define WINDOWS_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <CommCtrl.h>
#include <CommonControls.h>
#include <KnownFolders.h>
#include <ShlObj.h>
#pragma comment(lib, "Comctl32.lib")

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/result.h>
#pragma warning(pop)

#include <wincodec.h>

#include "Helpers.h"
#include "IconCache.h"
#ifdef ENABLE_TESTS
#include "SelfTest/Common/SelfTestLatencyHooks.h"
#endif
#include "WSLDistro.h"

namespace
{
constexpr std::wstring_view kDirectoryExtensionKey = L"<directory>";
constexpr std::wstring_view kWslLocalhostPrefix    = L"\\\\wsl.localhost\\";
constexpr std::wstring_view kWslDollarPrefix       = L"\\\\wsl$\\";
constexpr uint64_t kSlowIconCacheLockWaitUs        = 250u;
constexpr uint64_t kSlowIconCacheLockHoldUs        = 250u;
constexpr std::chrono::seconds kPathIconFailureTtl{5};
constexpr std::chrono::milliseconds kLivePathIconFailureInitialBackoff{250};
constexpr std::chrono::milliseconds kLivePathIconFailureMaxBackoff{4000};

[[nodiscard]] std::chrono::milliseconds LivePathIconFailureBackoff(uint32_t consecutiveFailureCount) noexcept
{
    constexpr uint32_t kMaxBackoffShift = 4u;
    const uint32_t failureIndex         = consecutiveFailureCount > 0u ? consecutiveFailureCount - 1u : 0u;
    const uint32_t shift                = std::min(failureIndex, kMaxBackoffShift);
    const auto scaledCount = kLivePathIconFailureInitialBackoff.count() * static_cast<int64_t>(uint64_t{1u} << shift);
    return std::chrono::milliseconds{std::min(scaledCount, kLivePathIconFailureMaxBackoff.count())};
}

struct AssociationQueryKey
{
    std::wstring extension;
    DWORD fileAttributes = 0;

    [[nodiscard]] bool operator==(const AssociationQueryKey& other) const noexcept = default;
};

struct AssociationQueryKeyHash
{
    [[nodiscard]] size_t operator()(const AssociationQueryKey& key) const noexcept
    {
        size_t hash = std::hash<std::wstring>{}(key.extension);
        hash ^= static_cast<size_t>(key.fileAttributes) + 0x9e3779b9u + (hash << 6) + (hash >> 2);
        return hash;
    }
};

struct AssociationIconCacheEntry
{
    int iconIndex         = 0;
    size_t lastAccessTime = 0;
};

struct AssociationLruEvictScanMetric
{
    bool shouldEmit       = false;
    uint64_t durationUs   = 0;
    uint64_t cacheSize    = 0;
    uint64_t removedCount = 0;
};

std::mutex g_associationCacheMutex;
std::unordered_map<AssociationQueryKey, AssociationIconCacheEntry, AssociationQueryKeyHash> g_associationToIconIndex;
size_t g_associationQueryAccessCounter = 0;
size_t g_associationLruEvictions       = 0;
constexpr size_t kAssociationCacheSize = 2000;

[[nodiscard]] bool ApplyMaskAlphaToIconPixels(HBITMAP maskBitmap, BYTE* pixels, int width, int height, HDC hdc)
{
    if (! maskBitmap || ! pixels || width <= 0 || height <= 0 || ! hdc)
    {
        return false;
    }

    BITMAP maskMetrics{};
    if (GetObjectW(maskBitmap, sizeof(maskMetrics), &maskMetrics) == 0 || maskMetrics.bmWidth < width || maskMetrics.bmHeight < height)
    {
        return false;
    }

    BITMAPINFO maskBmi{};
    maskBmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    maskBmi.bmiHeader.biWidth       = width;
    maskBmi.bmiHeader.biHeight      = -height;
    maskBmi.bmiHeader.biPlanes      = 1;
    maskBmi.bmiHeader.biBitCount    = 32;
    maskBmi.bmiHeader.biCompression = BI_RGB;

    const size_t pixelCount = static_cast<size_t>(width) * static_cast<size_t>(height);
    std::vector<BYTE> maskPixels(pixelCount * 4u);
    const int scanLines = GetDIBits(hdc, maskBitmap, 0, static_cast<UINT>(height), maskPixels.data(), &maskBmi, DIB_RGB_COLORS);
    if (scanLines != height)
    {
        return false;
    }

    for (size_t i = 0; i < pixelCount; ++i)
    {
        const size_t offset = i * 4u;
        // Windows icon AND masks render as black for opaque pixels and non-black for transparent pixels
        // when expanded through GetDIBits; preserve that contract for icons without real alpha.
        const bool transparent = maskPixels[offset + 0u] != 0 || maskPixels[offset + 1u] != 0 || maskPixels[offset + 2u] != 0;
        if (transparent)
        {
            pixels[offset + 0u] = 0;
            pixels[offset + 1u] = 0;
            pixels[offset + 2u] = 0;
            pixels[offset + 3u] = 0;
        }
        else
        {
            pixels[offset + 3u] = 255;
        }
    }
    return true;
}

void NormalizeIconBitmapAlphaForD2D(BYTE* pixels, size_t pixelCount, HBITMAP maskBitmap, int width, int height, HDC hdc)
{
    if (! pixels || pixelCount == 0u)
    {
        return;
    }

    bool sawAlpha = false;
    bool sawRgb   = false;
    for (size_t i = 0; i < pixelCount; ++i)
    {
        const size_t offset = i * 4u;
        const BYTE b        = pixels[offset + 0u];
        const BYTE g        = pixels[offset + 1u];
        const BYTE r        = pixels[offset + 2u];
        const BYTE a        = pixels[offset + 3u];

        sawAlpha = sawAlpha || a != 0;
        sawRgb   = sawRgb || b != 0 || g != 0 || r != 0;
        if (a > 0 && a < 255)
        {
            pixels[offset + 0u] = static_cast<BYTE>((static_cast<unsigned int>(b) * a + 127u) / 255u);
            pixels[offset + 1u] = static_cast<BYTE>((static_cast<unsigned int>(g) * a + 127u) / 255u);
            pixels[offset + 2u] = static_cast<BYTE>((static_cast<unsigned int>(r) * a + 127u) / 255u);
        }
    }

    if (! sawAlpha && sawRgb && ! ApplyMaskAlphaToIconPixels(maskBitmap, pixels, width, height, hdc))
    {
        for (size_t i = 0; i < pixelCount; ++i)
        {
            pixels[(i * 4u) + 3u] = 255;
        }
    }
}

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

void PerfEmitDurationWithDetail(
    std::wstring_view name, std::wstring_view detail, uint64_t durationUs, uint64_t value0 = 0, uint64_t value1 = 0, HRESULT hr = S_OK) noexcept
{
    Debug::Perf::Emit(name, detail, durationUs, value0, value1, hr);
}

[[nodiscard]] bool ShouldAvoidRecallForIconPathLookup(DWORD fileAttributes) noexcept
{
    constexpr DWORD recallAttributes = FILE_ATTRIBUTE_OFFLINE | FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS | FILE_ATTRIBUTE_RECALL_ON_OPEN;
    return (fileAttributes & recallAttributes) != 0u;
}

[[nodiscard]] DWORD NormalizeAttributesForShellAttributeIconLookup(DWORD fileAttributes) noexcept
{
    return (fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;
}

void PerfEmitSlowIconCacheLockWait(std::wstring_view detail, uint64_t durationUs) noexcept
{
    if (durationUs >= kSlowIconCacheLockWaitUs)
    {
        Debug::Perf::Emit(L"iconcache.lock_wait_slow_us", detail, durationUs, kSlowIconCacheLockWaitUs, 0u, S_OK);
    }
}

void PerfEmitSlowIconCacheLockHold(std::wstring_view detail, uint64_t durationUs) noexcept
{
    if (durationUs >= kSlowIconCacheLockHoldUs)
    {
        Debug::Perf::Emit(L"iconcache.lock_hold_slow_us", detail, durationUs, kSlowIconCacheLockHoldUs, 0u, S_OK);
    }
}

void EmitAssociationLruEvictScanMetric(const AssociationLruEvictScanMetric& metric) noexcept
{
    if (! metric.shouldEmit)
    {
        return;
    }

    PerfEmitDuration(L"iconcache.association_lru_evict_scan_us", metric.durationUs, metric.cacheSize, metric.removedCount, S_OK);
}

[[nodiscard]] AssociationLruEvictScanMetric EvictAssociationQueryBatch()
{
    if (g_associationToIconIndex.size() < kAssociationCacheSize)
    {
        return {};
    }

    const auto scanStart    = std::chrono::steady_clock::now();
    const size_t evictCount = std::max<size_t>(g_associationToIconIndex.size() / 10, 1);
    std::vector<size_t> times;
    times.reserve(g_associationToIconIndex.size());
    for (const auto& [key, entry] : g_associationToIconIndex)
    {
        times.push_back(entry.lastAccessTime);
    }

    std::nth_element(times.begin(), times.begin() + static_cast<ptrdiff_t>(evictCount - 1), times.end());
    const size_t threshold = times[evictCount - 1];

    size_t removed = 0;
    for (auto it = g_associationToIconIndex.begin(); it != g_associationToIconIndex.end();)
    {
        if (it->second.lastAccessTime <= threshold && removed < evictCount)
        {
            it = g_associationToIconIndex.erase(it);
            ++removed;
        }
        else
        {
            ++it;
        }
    }

    g_associationLruEvictions += removed;
    return AssociationLruEvictScanMetric{
        .shouldEmit   = true,
        .durationUs   = PerfElapsedUs(scanStart),
        .cacheSize    = static_cast<uint64_t>(g_associationToIconIndex.size()),
        .removedCount = static_cast<uint64_t>(removed),
    };
}

[[nodiscard]] std::wstring NormalizeExtensionKey(std::wstring_view extension)
{
    if (extension.empty())
    {
        return {};
    }

    if (wil::compare_string_ordinal(extension, kDirectoryExtensionKey, true) == wistd::weak_ordering::equivalent)
    {
        return std::wstring(kDirectoryExtensionKey);
    }

    std::wstring key;
    key.reserve(extension.size() + 1u);

    if (extension[0] != L'.' && extension[0] != L'<')
    {
        key.push_back(L'.');
    }

    for (const wchar_t ch : extension)
    {
        key.push_back(static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))));
    }

    return key;
}

[[nodiscard]] bool IsDirectoryExtensionKey(std::wstring_view extension) noexcept
{
    return wil::compare_string_ordinal(extension, kDirectoryExtensionKey, true) == wistd::weak_ordering::equivalent;
}

[[nodiscard]] std::wstring BuildAssociationQueryPath(std::wstring_view normalizedExtension)
{
    if (IsDirectoryExtensionKey(normalizedExtension))
    {
        return L"C:\\DummyFolder\\";
    }

    std::wstring queryPath = L"C:\\Dummy";
    queryPath.append(normalizedExtension);
    return queryPath;
}

[[nodiscard]] std::optional<int> QueryAssociationIconIndex(std::wstring_view normalizedExtension, DWORD fileAttributes) noexcept
{
    const std::wstring extensionKey(normalizedExtension);

    std::optional<int> cachedIconIndex;
    uint64_t lookupWaitUs = 0;
    uint64_t lookupHoldUs = 0;
    {
        const auto lockWaitStart = std::chrono::steady_clock::now();
        std::lock_guard lock(g_associationCacheMutex);
        lookupWaitUs             = PerfElapsedUs(lockWaitStart);
        const auto lockHoldStart = std::chrono::steady_clock::now();
        AssociationQueryKey cacheKey{extensionKey, fileAttributes};
        const auto cached = g_associationToIconIndex.find(cacheKey);
        if (cached != g_associationToIconIndex.end())
        {
            cached->second.lastAccessTime = ++g_associationQueryAccessCounter;
            cachedIconIndex               = cached->second.iconIndex;
        }
        lookupHoldUs = PerfElapsedUs(lockHoldStart);
    }
    PerfEmitSlowIconCacheLockWait(L"association_lookup", lookupWaitUs);
    PerfEmitSlowIconCacheLockHold(L"association_lookup", lookupHoldUs);
    if (cachedIconIndex.has_value())
    {
        return cachedIconIndex;
    }

    SHFILEINFOW sfi{};
    const std::wstring queryPath = BuildAssociationQueryPath(normalizedExtension);
    constexpr UINT flags         = SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES;
    const auto shellStart        = std::chrono::steady_clock::now();
    const DWORD_PTR result       = SHGetFileInfoW(queryPath.c_str(), fileAttributes, &sfi, sizeof(sfi), flags);
    PerfEmitDurationWithDetail(L"iconcache.shgetfileinfo_us",
                               L"association",
                               PerfElapsedUs(shellStart),
                               static_cast<uint64_t>(fileAttributes),
                               static_cast<uint64_t>(flags),
                               result == 0 || sfi.iIcon < 0 ? S_FALSE : S_OK);
    if (result == 0 || sfi.iIcon < 0)
    {
        return std::nullopt;
    }

    uint64_t storeWaitUs = 0;
    uint64_t storeHoldUs = 0;
    AssociationLruEvictScanMetric evictionMetric{};
    std::optional<int> storedIconIndex;
    {
        const auto lockWaitStart = std::chrono::steady_clock::now();
        std::lock_guard lock(g_associationCacheMutex);
        storeWaitUs              = PerfElapsedUs(lockWaitStart);
        const auto lockHoldStart = std::chrono::steady_clock::now();
        AssociationQueryKey cacheKey{extensionKey, fileAttributes};
        evictionMetric = EvictAssociationQueryBatch();

        AssociationIconCacheEntry entry{};
        entry.iconIndex                               = sfi.iIcon;
        entry.lastAccessTime                          = ++g_associationQueryAccessCounter;
        g_associationToIconIndex[std::move(cacheKey)] = entry;
        storedIconIndex                               = entry.iconIndex;
        storeHoldUs                                   = PerfElapsedUs(lockHoldStart);
    }
    PerfEmitSlowIconCacheLockWait(L"association_store", storeWaitUs);
    PerfEmitSlowIconCacheLockHold(L"association_store", storeHoldUs);
    EmitAssociationLruEvictScanMetric(evictionMetric);
    return storedIconIndex;
}

[[nodiscard]] std::optional<int> QueryCachedDefaultAssociationIconIndex() noexcept
{
    static std::mutex defaultAssociationIconMutex;
    static std::optional<int> defaultAssociationIconIndex;

    {
        std::lock_guard lock(defaultAssociationIconMutex);
        if (defaultAssociationIconIndex.has_value())
        {
            return defaultAssociationIconIndex;
        }
    }

    const auto iconIndex = QueryAssociationIconIndex(std::wstring_view{}, FILE_ATTRIBUTE_NORMAL);
    if (iconIndex.has_value())
    {
        std::lock_guard lock(defaultAssociationIconMutex);
        defaultAssociationIconIndex = iconIndex;
    }

    return iconIndex;
}

[[nodiscard]] int SelectImageListSizeForTargetPixels(float targetPixels) noexcept
{
    if (! (targetPixels > 16.0f))
    {
        return SHIL_SMALL;
    }

    if (targetPixels <= 32.0f)
    {
        return SHIL_LARGE;
    }

    if (targetPixels <= 48.0f)
    {
        return SHIL_EXTRALARGE;
    }

    return SHIL_JUMBO;
}

[[nodiscard]] bool StartsWithIgnoreCase(std::wstring_view value, std::wstring_view prefix) noexcept
{
    if (value.size() < prefix.size())
    {
        return false;
    }

    const std::wstring_view head = value.substr(0, prefix.size());
    return wil::compare_string_ordinal(head, prefix, true) == wistd::weak_ordering::equivalent;
}

[[nodiscard]] std::optional<std::wstring> TryExtractWslDistroName(std::wstring_view path) noexcept
{
    std::wstring_view remainder;

    if (StartsWithIgnoreCase(path, kWslLocalhostPrefix))
    {
        remainder = path.substr(kWslLocalhostPrefix.size());
    }
    else if (StartsWithIgnoreCase(path, kWslDollarPrefix))
    {
        remainder = path.substr(kWslDollarPrefix.size());
    }
    else
    {
        return std::nullopt;
    }

    if (remainder.empty())
    {
        return std::nullopt;
    }

    const size_t separatorIndex        = remainder.find_first_of(L"\\/");
    const std::wstring_view distroView = (separatorIndex == std::wstring_view::npos) ? remainder : remainder.substr(0, separatorIndex);

    if (distroView.empty())
    {
        return std::nullopt;
    }

    return std::wstring(distroView);
}

[[nodiscard]] wil::unique_hicon ExtractShellSmallIconForPath(const wchar_t* path, DWORD fileAttributes, bool useFileAttributes) noexcept
{
    if (! path || path[0] == L'\0')
    {
        return {};
    }

    SHFILEINFOW sfi{};
    UINT flags = SHGFI_ICON | SHGFI_SMALLICON;
    if (useFileAttributes)
    {
        flags |= SHGFI_USEFILEATTRIBUTES;
    }

    const DWORD_PTR result = SHGetFileInfoW(path, fileAttributes, &sfi, sizeof(sfi), flags);
    if (result == 0 || ! sfi.hIcon)
    {
        return {};
    }

    return wil::unique_hicon{sfi.hIcon};
}

[[nodiscard]] wil::unique_hicon ExtractShellSmallIconForPidl(PCIDLIST_ABSOLUTE pidl) noexcept
{
    if (! pidl)
    {
        return {};
    }

    SHFILEINFOW sfi{};
    const DWORD_PTR result = SHGetFileInfoW(reinterpret_cast<PCWSTR>(pidl), 0, &sfi, sizeof(sfi), SHGFI_PIDL | SHGFI_ICON | SHGFI_SMALLICON);
    if (result == 0 || ! sfi.hIcon)
    {
        return {};
    }

    return wil::unique_hicon{sfi.hIcon};
}
} // namespace

#if defined(ENABLE_TESTS)
bool DebugNormalizeIconBitmapAlphaForD2D(std::span<BYTE> pixels) noexcept
{
    if (pixels.size() % 4u != 0u)
    {
        return false;
    }

    NormalizeIconBitmapAlphaForD2D(pixels.data(), pixels.size() / 4u, nullptr, 0, 0, nullptr);
    return true;
}

bool DebugNormalizeIconBitmapAlphaForD2DWithMask(std::span<BYTE> pixels, HBITMAP maskBitmap, int width, int height, HDC hdc) noexcept
{
    if (pixels.size() % 4u != 0u || width <= 0 || height <= 0)
    {
        return false;
    }

    const size_t expectedPixels = static_cast<size_t>(width) * static_cast<size_t>(height);
    if (pixels.size() / 4u != expectedPixels)
    {
        return false;
    }

    NormalizeIconBitmapAlphaForD2D(pixels.data(), expectedPixels, maskBitmap, width, height, hdc);
    return true;
}

int DebugSelectIconCacheImageListSize(float targetDipSize, float dpi) noexcept
{
    const float targetPixels = targetDipSize * dpi / 96.0f;
    return SelectImageListSizeForTargetPixels(targetPixels);
}

size_t DebugGetAssociationIconCacheSize() noexcept
{
    std::lock_guard lock(g_associationCacheMutex);
    return g_associationToIconIndex.size();
}

uint32_t DebugGetLivePathIconFailureBackoffMs(uint32_t consecutiveFailureCount) noexcept
{
    return static_cast<uint32_t>(LivePathIconFailureBackoff(consecutiveFailureCount).count());
}
#endif

IconCache& IconCache::GetInstance()
{
    // Intentionally leaked to avoid shutdown UAF from static destruction order issues.
    static IconCache* instance = new IconCache();
    return *instance;
}

void IconCache::Shutdown() noexcept
{
    size_t bitmapCount = 0;
    size_t extCount    = 0;
    size_t pathCount   = 0;
    {
        std::lock_guard lock(_mutex);

        for (const auto& entry : _deviceCaches)
        {
            bitmapCount += entry.second.bitmaps.size();
        }
        extCount  = _extensionToIconIndex.size();
        pathCount = _pathToIconIndex.size();

        _deviceCaches.clear();
        _extensionToIconIndex.clear();
        _pathToIconIndex.clear();
        _pathQueryAccessCounter = 0;
        _pathLruEvictions       = 0;
        _extractionFailureCount.clear();

        _systemImageListJumbo.reset();
        _systemImageListXL.reset();
        _systemImageListLarge.reset();
        _systemImageListSmall.reset();
        _wicFactory.reset();

        _initialized.store(false, std::memory_order_release);
        _warmingCompleted.store(false, std::memory_order_release);
        _warmingInProgress.store(false, std::memory_order_release);
    }

    DBGOUT_INFO(L"IconCache: Shutdown (cleared {} cached icons, {} extension mappings, {} path mappings)", bitmapCount, extCount, pathCount);
}

void IconCache::Initialize(ID2D1DeviceContext* d2dContext, float dpi)
{
    if (! d2dContext)
    {
        return;
    }

    const auto lockWaitStart = std::chrono::steady_clock::now();
    std::lock_guard lock(_mutex);
    PerfEmitSlowIconCacheLockWait(L"initialize", PerfElapsedUs(lockWaitStart));
    const auto lockHoldStart = std::chrono::steady_clock::now();
    auto lockHold            = wil::scope_exit([&] { PerfEmitSlowIconCacheLockHold(L"initialize", PerfElapsedUs(lockHoldStart)); });
    _dpi.store(dpi, std::memory_order_relaxed);

    // Initialize WIC factory for high-quality icon conversion
    wil::com_ptr<IWICImagingFactory> wicFactory;
    HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&wicFactory));
    if (hr == REGDB_E_CLASSNOTREG)
    {
        hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&wicFactory));
    }
    if (SUCCEEDED(hr))
    {
        _wicFactory = std::move(wicFactory);
    }
    else
    {
        if (hr == CO_E_NOTINITIALIZED)
        {
            Debug::Warning(L"IconCache: Failed to create WIC factory (COM not initialized on this thread): 0x{:08X}", hr);
        }
        else
        {
            Debug::Warning(L"IconCache: Failed to create WIC factory: 0x{:08X}", hr);
        }
        // Continue without WIC - conversion will fail gracefully
    }

    // Initialize special folder paths for quick lookup
    std::call_once(_specialFoldersInitOnce, [] { InitializeSpecialFolders(); });

    // Get all three system image list sizes for fallback support
    // Always try SHIL_EXTRALARGE (48×48) first for best quality on high-DPI displays
    if (! _systemImageListJumbo)
    {
        const HRESULT hrJumbo = SHGetImageList(SHIL_JUMBO, IID_PPV_ARGS(_systemImageListJumbo.put()));
        if (SUCCEEDED(hrJumbo))
        {
            DBGOUT_INFO(L"IconCache: Initialized SHIL_JUMBO (256×256) at {:.0f} DPI", dpi);
        }
        else
        {
            Debug::Warning(L"IconCache: Failed to get SHIL_JUMBO image list: 0x{:08X}", hrJumbo);
        }
    }

    if (! _systemImageListXL)
    {
        HRESULT hrXL = SHGetImageList(SHIL_EXTRALARGE, IID_PPV_ARGS(_systemImageListXL.put()));
        if (SUCCEEDED(hrXL))
        {
            DBGOUT_INFO(L"IconCache: Initialized SHIL_EXTRALARGE (48×48) at {:.0f} DPI", dpi);
        }
        else
        {
            Debug::Warning(L"IconCache: Failed to get SHIL_EXTRALARGE image list: 0x{:08X}", hrXL);
        }
    }

    if (! _systemImageListLarge)
    {
        HRESULT hrLarge = SHGetImageList(SHIL_LARGE, IID_PPV_ARGS(_systemImageListLarge.put()));
        if (FAILED(hrLarge))
        {
            Debug::Warning(L"IconCache: Failed to get SHIL_LARGE image list: 0x{:08X}", hrLarge);
        }
    }

    if (! _systemImageListSmall)
    {
        HRESULT hrSmall = SHGetImageList(SHIL_SMALL, IID_PPV_ARGS(_systemImageListSmall.put()));
        if (FAILED(hrSmall))
        {
            Debug::Warning(L"IconCache: Failed to get SHIL_SMALL image list: 0x{:08X}", hrSmall);
        }
    }

    _initialized.store(true, std::memory_order_release);
}

void IconCache::SetDpi(float dpi)
{
    std::lock_guard lock(_mutex);

    // Detect DPI change and refresh cache if needed
    const float currentDpi = _dpi.load(std::memory_order_relaxed);
    if (std::abs(currentDpi - dpi) > 0.1f)
    {
        const float oldDpi = currentDpi;
        _dpi.store(dpi, std::memory_order_relaxed);

        Debug::Info(L"IconCache: DPI changed from {:.0f} to {:.0f}, clearing cache", oldDpi, dpi);

        // Clear cache (icons extracted at old DPI may not be optimal)
        _deviceCaches.clear();
        _extensionToIconIndex.clear();
        _warmingCompleted.store(false, std::memory_order_release);

        // Note: We don't re-initialize image lists here as they're size-independent
        // They'll be used for extraction at the new DPI during rendering
    }
    else
    {
        _dpi.store(dpi, std::memory_order_relaxed);
    }
}

int IconCache::MakeBitmapCacheSizeClass(float targetDipSize) const noexcept
{
    return SelectOptimalImageListSize(targetDipSize);
}

wil::com_ptr<ID2D1Bitmap1> IconCache::GetIconBitmap(int iconIndex, ID2D1DeviceContext* d2dContext, float targetDipSize)
{
    if (! _initialized.load(std::memory_order_acquire) || ! d2dContext || iconIndex < 0)
    {
        return nullptr;
    }

    wil::com_ptr<ID2D1Device> device;
    d2dContext->GetDevice(device.put());
    if (! device)
    {
        return nullptr;
    }

    const IconBitmapCacheKey cacheKey{.iconIndex = iconIndex, .imageListSize = MakeBitmapCacheSizeClass(targetDipSize)};

    // Check cache first
    {
        const auto lockWaitStart = std::chrono::steady_clock::now();
        std::lock_guard lock(_mutex);
        PerfEmitSlowIconCacheLockWait(L"bitmap_lookup", PerfElapsedUs(lockWaitStart));
        const auto lockHoldStart = std::chrono::steady_clock::now();
        auto lockHold            = wil::scope_exit([&] { PerfEmitSlowIconCacheLockHold(L"bitmap_lookup", PerfElapsedUs(lockHoldStart)); });
        auto deviceIt            = _deviceCaches.find(device.get());
        if (deviceIt != _deviceCaches.end())
        {
            auto it = deviceIt->second.bitmaps.find(cacheKey);
            if (it != deviceIt->second.bitmaps.end())
            {
                _hitCount++;
                PerfEmitCounter(L"iconcache.get_bitmap_hit", 1);
                it->second.lastAccessTime = ++deviceIt->second.accessCounter;
                return it->second.bitmap;
            }
        }
        _missCount++;
        PerfEmitCounter(L"iconcache.get_bitmap_miss", 1);
    }

    // Cache miss - extract icon from system image list
    const auto extractStart = std::chrono::steady_clock::now();
    wil::unique_hicon icon  = ExtractSystemIcon(iconIndex, targetDipSize);
    PerfEmitDuration(L"iconcache.miss_extract_us",
                     PerfElapsedUs(extractStart),
                     static_cast<uint64_t>(iconIndex),
                     static_cast<uint64_t>(cacheKey.imageListSize),
                     icon ? S_OK : S_FALSE);
    if (! icon)
    {
        return nullptr;
    }

    // Convert to D2D bitmap
    const auto convertStart = std::chrono::steady_clock::now();
    auto bitmap             = ConvertIconToBitmap(icon.get(), d2dContext);
    PerfEmitDuration(L"iconcache.miss_convert_us", PerfElapsedUs(convertStart), static_cast<uint64_t>(iconIndex), 0, bitmap ? S_OK : S_FALSE);

    if (bitmap)
    {
        const D2D1_SIZE_U pixelSize = bitmap->GetPixelSize();
        const size_t bytes          = static_cast<size_t>(pixelSize.width) * static_cast<size_t>(pixelSize.height) * 4u;

        // Store in cache with LRU tracking
        const auto storeStart = std::chrono::steady_clock::now();
        std::lock_guard lock(_mutex);
        PerfEmitSlowIconCacheLockWait(L"bitmap_store", PerfElapsedUs(storeStart));
        const auto storeHoldStart = std::chrono::steady_clock::now();
        auto storeHold            = wil::scope_exit([&] { PerfEmitSlowIconCacheLockHold(L"bitmap_store", PerfElapsedUs(storeHoldStart)); });
        auto& cache               = _deviceCaches[device.get()];
        if (! cache.device)
        {
            cache.device = device;
        }
        EvictLRUIfNeeded(cache);
        CacheEntry entry;
        entry.bitmap            = bitmap;
        entry.lastAccessTime    = ++cache.accessCounter;
        entry.bytes             = bytes;
        cache.bitmaps[cacheKey] = std::move(entry);
        PerfEmitDuration(L"iconcache.miss_store_us", PerfElapsedUs(storeStart), static_cast<uint64_t>(iconIndex), bytes, S_OK);
    }

    return bitmap;
}

bool IconCache::HasCachedIcon(int iconIndex, ID2D1Device* device, float targetDipSize) const
{
    if (iconIndex < 0 || ! device)
    {
        return false;
    }

    const IconBitmapCacheKey cacheKey{.iconIndex = iconIndex, .imageListSize = MakeBitmapCacheSizeClass(targetDipSize)};

    const auto lockWaitStart = std::chrono::steady_clock::now();
    std::lock_guard lock(_mutex);
    PerfEmitSlowIconCacheLockWait(L"has_cached_icon", PerfElapsedUs(lockWaitStart));
    const auto lockHoldStart = std::chrono::steady_clock::now();
    auto lockHold            = wil::scope_exit([&] { PerfEmitSlowIconCacheLockHold(L"has_cached_icon", PerfElapsedUs(lockHoldStart)); });
    const auto deviceIt      = _deviceCaches.find(device);
    if (deviceIt == _deviceCaches.end())
    {
        return false;
    }

    return deviceIt->second.bitmaps.find(cacheKey) != deviceIt->second.bitmaps.end();
}

wil::com_ptr<ID2D1Bitmap1> IconCache::GetCachedBitmap(int iconIndex, ID2D1DeviceContext* d2dContext, float targetDipSize) const
{
    if (iconIndex < 0 || ! d2dContext)
    {
        return nullptr;
    }

    wil::com_ptr<ID2D1Device> device;
    d2dContext->GetDevice(device.put());
    if (! device)
    {
        return nullptr;
    }

    const IconBitmapCacheKey cacheKey{.iconIndex = iconIndex, .imageListSize = MakeBitmapCacheSizeClass(targetDipSize)};

    const auto lockWaitStart = std::chrono::steady_clock::now();
    std::lock_guard lock(_mutex);
    PerfEmitSlowIconCacheLockWait(L"cached_bitmap_lookup", PerfElapsedUs(lockWaitStart));
    const auto lockHoldStart = std::chrono::steady_clock::now();
    auto lockHold            = wil::scope_exit([&] { PerfEmitSlowIconCacheLockHold(L"cached_bitmap_lookup", PerfElapsedUs(lockHoldStart)); });
    const auto deviceIt      = _deviceCaches.find(device.get());
    if (deviceIt == _deviceCaches.end())
    {
        return nullptr;
    }

    const auto it = deviceIt->second.bitmaps.find(cacheKey);
    if (it != deviceIt->second.bitmaps.end())
    {
        return it->second.bitmap;
    }
    return nullptr;
}

wil::unique_hicon IconCache::ExtractSystemIcon(int iconIndex, float targetDipSize)
{
    if (iconIndex < 0)
    {
        return {};
    }

    if (! _initialized.load(std::memory_order_acquire))
    {
        return {};
    }

#ifdef ENABLE_TESTS
    SelfTestLatency::Consume(SelfTestLatency::Point::IconExtractSystemIcon);
#endif

    auto tryExtract = [&](IImageList* imageList) -> wil::unique_hicon
    {
        if (! imageList)
        {
            return {};
        }

        HICON hIcon      = nullptr;
        const HRESULT hr = imageList->GetIcon(iconIndex, ILD_NORMAL, &hIcon);
        wil::unique_hicon icon{hIcon};
        if (SUCCEEDED(hr) && icon)
        {
            return icon;
        }

        return {};
    };

    // Determine optimal image list based on DPI and target display size.
    const int optimalSize = SelectOptimalImageListSize(targetDipSize);

    // Try optimal size first.
    if (optimalSize == SHIL_JUMBO)
    {
        if (auto icon = tryExtract(_systemImageListJumbo.get()))
        {
            return icon;
        }
    }
    else if (optimalSize == SHIL_EXTRALARGE)
    {
        if (auto icon = tryExtract(_systemImageListXL.get()))
        {
            return icon;
        }
    }
    else if (optimalSize == SHIL_LARGE)
    {
        if (auto icon = tryExtract(_systemImageListLarge.get()))
        {
            return icon;
        }
    }
    else
    {
        if (auto icon = tryExtract(_systemImageListSmall.get()))
        {
            return icon;
        }
    }

    // Fallback cascade: Try remaining sizes in order of preference.
    if (optimalSize != SHIL_EXTRALARGE)
    {
        if (auto icon = tryExtract(_systemImageListXL.get()))
        {
            return icon;
        }
    }

    if (optimalSize != SHIL_LARGE)
    {
        if (auto icon = tryExtract(_systemImageListLarge.get()))
        {
            return icon;
        }
    }

    if (optimalSize != SHIL_SMALL)
    {
        if (auto icon = tryExtract(_systemImageListSmall.get()))
        {
            return icon;
        }
    }

    if (optimalSize != SHIL_JUMBO)
    {
        if (auto icon = tryExtract(_systemImageListJumbo.get()))
        {
            return icon;
        }
    }

    // All extraction attempts failed - this is unusual and worth logging.
    Debug::Warning(L"IconCache: Failed to extract icon index {} from all sizes (Jumbo/XL/Large/Small)", iconIndex);
    return {};
}

wil::unique_hbitmap IconCache::CreateMenuBitmapFromIcon(HICON icon, int size)
{
    if (! icon || size <= 0)
    {
        return nullptr;
    }

    constexpr int kMaxMenuIconBitmapDimension = 512;
    if (size > kMaxMenuIconBitmapDimension)
    {
        Debug::Warning(L"IconCache: Refusing oversized menu HICON bitmap {}x{}", size, size);
        return nullptr;
    }

    wil::unique_hdc_window hdcScreen{GetDC(nullptr)};
    if (! hdcScreen)
    {
        Debug::Warning(L"IconCache: Failed to acquire screen DC for menu icon conversion");
        return nullptr;
    }

    wil::unique_hdc memoryDc{CreateCompatibleDC(hdcScreen.get())};
    if (! memoryDc)
    {
        Debug::Warning(L"IconCache: Failed to create memory DC for menu icon conversion");
        return nullptr;
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = size;
    bmi.bmiHeader.biHeight      = -size;
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    wil::unique_hbitmap bitmap{CreateDIBSection(hdcScreen.get(), &bmi, DIB_RGB_COLORS, &bits, nullptr, 0)};
    if (! bitmap || ! bits)
    {
        Debug::Warning(L"IconCache: Failed to create DIB section for menu icon");
        return nullptr;
    }

    const size_t pixelCount = static_cast<size_t>(size) * static_cast<size_t>(size);
    auto* pixels            = static_cast<BYTE*>(bits);
    std::fill_n(pixels, pixelCount * 4u, BYTE{0});

    auto oldBitmap        = wil::SelectObject(memoryDc.get(), bitmap.get());
    const auto drawStart  = std::chrono::steady_clock::now();
    const BOOL drawResult = DrawIconEx(memoryDc.get(), 0, 0, icon, size, size, 0, nullptr, DI_NORMAL);
    const DWORD drawError = drawResult ? ERROR_SUCCESS : GetLastError();
    oldBitmap.reset();
    PerfEmitDuration(L"iconcache.menu_draw_icon_us",
                     PerfElapsedUs(drawStart),
                     static_cast<uint64_t>(size),
                     static_cast<uint64_t>(size),
                     drawResult ? S_OK : HRESULT_FROM_WIN32(drawError == ERROR_SUCCESS ? ERROR_INVALID_DATA : drawError));
    if (! drawResult)
    {
        Debug::Warning(L"IconCache: Failed to draw HICON into menu DIB: 0x{:08X}",
                       HRESULT_FROM_WIN32(drawError == ERROR_SUCCESS ? ERROR_INVALID_DATA : drawError));
        return nullptr;
    }

    NormalizeIconBitmapAlphaForD2D(pixels, pixelCount, nullptr, 0, 0, nullptr);

    return bitmap;
}

wil::unique_hbitmap IconCache::CreateMenuBitmapFromIconIndex(int iconIndex, int size)
{
    if (iconIndex < 0 || size <= 0)
    {
        return nullptr;
    }

    // size is in physical pixels (GDI menus); derive an approximate DIP size for selecting the best source image list.
    const float dpi           = _dpi.load(std::memory_order_relaxed);
    const float effectiveDpi  = (dpi > 1.0f) ? dpi : static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    const float targetDipSize = static_cast<float>(size) * static_cast<float>(USER_DEFAULT_SCREEN_DPI) / effectiveDpi;

    wil::unique_hicon icon = ExtractSystemIcon(iconIndex, targetDipSize);
    if (! icon)
    {
        return nullptr;
    }

    return CreateMenuBitmapFromIcon(icon.get(), size);
}

wil::unique_hbitmap IconCache::CreateMenuBitmapFromPath(const wchar_t* path, int size, DWORD fileAttributes, bool useFileAttributes)
{
    if (! path || path[0] == L'\0' || size <= 0)
    {
        return nullptr;
    }

    const std::wstring_view pathView = path;

    if (const auto special = TryGetSpecialFolderForPathPrefix(pathView); special.has_value() && special.value().iconIndex >= 0)
    {
        wil::unique_hbitmap bitmap = CreateMenuBitmapFromIconIndex(special.value().iconIndex, size);
        if (bitmap)
        {
            return bitmap;
        }
    }

    if (const auto distroName = TryExtractWslDistroName(pathView); distroName.has_value())
    {
        wil::unique_hicon icon = WSLDistro::LoadDistributionIcon(distroName.value(), size);
        if (icon)
        {
            wil::unique_hbitmap bitmap = CreateMenuBitmapFromIcon(icon.get(), size);
            if (bitmap)
            {
                return bitmap;
            }
        }
    }

    // Prefer the per-path shell icon when available so known folders and shell-customized entries
    // keep their stock small bitmap instead of collapsing to the shared system-image-list fallback.
    if (wil::unique_hicon directIcon = ExtractShellSmallIconForPath(path, fileAttributes, useFileAttributes); directIcon)
    {
        if (wil::unique_hbitmap bitmap = CreateMenuBitmapFromIcon(directIcon.get(), size))
        {
            return bitmap;
        }
    }

    const auto iconIndex = QuerySysIconIndexForPath(path, fileAttributes, useFileAttributes);
    if (! iconIndex.has_value())
    {
        return nullptr;
    }

    return CreateMenuBitmapFromIconIndex(iconIndex.value(), size);
}

wil::unique_hbitmap IconCache::CreateMenuBitmapFromKnownFolder(const GUID& folderId, int size)
{
    if (size <= 0)
    {
        return nullptr;
    }

    PIDLIST_ABSOLUTE pidl = nullptr;
    const HRESULT hr      = SHGetKnownFolderIDList(folderId, 0, nullptr, &pidl);
    if (FAILED(hr) || ! pidl)
    {
        return nullptr;
    }

    const auto pidlCleanup = wil::scope_exit([&] { ILFree(pidl); });
    // Known folders can expose a richer direct shell icon than their shared image-list index.
    // Keep that direct small icon when possible, then fall back to the image list if extraction fails.
    if (wil::unique_hicon icon = ExtractShellSmallIconForPidl(pidl); icon)
    {
        if (wil::unique_hbitmap bitmap = CreateMenuBitmapFromIcon(icon.get(), size))
        {
            return bitmap;
        }
    }

    if (const auto iconIndex = QuerySysIconIndexForPidl(pidl); iconIndex.has_value())
    {
        return CreateMenuBitmapFromIconIndex(iconIndex.value(), size);
    }

    return nullptr;
}

std::optional<int> IconCache::QuerySysIconIndexForPath(const wchar_t* path, DWORD fileAttributes, bool useFileAttributes)
{
    if (! path || *path == L'\0')
    {
        return std::nullopt;
    }

    const bool recallAvoided = ! useFileAttributes && ShouldAvoidRecallForIconPathLookup(fileAttributes);
    if (recallAvoided)
    {
        useFileAttributes = true;
        fileAttributes    = NormalizeAttributesForShellAttributeIconLookup(fileAttributes);
        PerfEmitCounter(L"icons.recall_avoided_count", 1u);
    }

    PathQueryKey cacheKey{};
    cacheKey.path              = path;
    cacheKey.fileAttributes    = fileAttributes;
    cacheKey.useFileAttributes = useFileAttributes;

    bool failureCacheHit     = false;
    bool failureCacheExpired = false;
    {
        const auto lockWaitStart = std::chrono::steady_clock::now();
        std::lock_guard lock(_mutex);
        PerfEmitSlowIconCacheLockWait(L"path_lookup", PerfElapsedUs(lockWaitStart));
        const auto lockHoldStart = std::chrono::steady_clock::now();
        auto lockHold            = wil::scope_exit([&] { PerfEmitSlowIconCacheLockHold(L"path_lookup", PerfElapsedUs(lockHoldStart)); });
        const auto it            = _pathToIconIndex.find(cacheKey);
        if (it != _pathToIconIndex.end())
        {
            if (it->second.lookupFailed)
            {
                const auto now = std::chrono::steady_clock::now();
                const auto failureBackoff = useFileAttributes ? std::chrono::duration_cast<std::chrono::milliseconds>(kPathIconFailureTtl)
                                                               : LivePathIconFailureBackoff(it->second.consecutiveFailureCount);
                if (now - it->second.failureStamp < failureBackoff)
                {
                    it->second.lastAccessTime = ++_pathQueryAccessCounter;
                    failureCacheHit           = true;
                }
                else
                {
                    failureCacheExpired = true;
                }
            }
            else
            {
                it->second.lastAccessTime = ++_pathQueryAccessCounter;
                return it->second.iconIndex;
            }
        }
    }

    if (failureCacheHit)
    {
        PerfEmitCounter(L"iconcache.path_failed_lookup_cache_hit", 1u);
        if (! useFileAttributes)
        {
            PerfEmitCounter(L"iconcache.path_live_lookup_failure_cache_hit", 1u);
        }
        return std::nullopt;
    }
    if (failureCacheExpired)
    {
        PerfEmitCounter(L"iconcache.path_failed_lookup_cache_expired", 1u);
    }

    // SHGetFileInfoW is called outside the lock — concurrent threads may query the same path.
    // This is intentional: duplicate work is harmless (same result), and holding the lock across
    // a potentially slow shell call would serialize all icon lookups.
    UINT flags = SHGFI_SYSICONINDEX;
    if (useFileAttributes)
    {
        flags |= SHGFI_USEFILEATTRIBUTES;
    }

#ifdef ENABLE_TESTS
    HRESULT forcedLiveLookupFailure = S_OK;
    if (! useFileAttributes)
    {
        SelfTestLatency::Consume(SelfTestLatency::Point::IconPathLiveLookup);
        forcedLiveLookupFailure = SelfTestLatency::ConsumeFailure(SelfTestLatency::Point::IconPathLiveLookup);
    }
#endif

    SHFILEINFOW sfi{};
    const auto shellStart = std::chrono::steady_clock::now();
    DWORD_PTR result      = 0;
#ifdef ENABLE_TESTS
    if (SUCCEEDED(forcedLiveLookupFailure))
#endif
    {
        result = SHGetFileInfoW(path, fileAttributes, &sfi, sizeof(sfi), flags);
    }
#ifdef ENABLE_TESTS
    const HRESULT lookupStatus = FAILED(forcedLiveLookupFailure) ? forcedLiveLookupFailure : (result == 0 || sfi.iIcon < 0 ? S_FALSE : S_OK);
#else
    const HRESULT lookupStatus = result == 0 || sfi.iIcon < 0 ? S_FALSE : S_OK;
#endif
    PerfEmitDurationWithDetail(L"iconcache.shgetfileinfo_us",
                               useFileAttributes ? L"path_attributes" : L"path_live",
                               PerfElapsedUs(shellStart),
                               static_cast<uint64_t>(fileAttributes),
                               static_cast<uint64_t>(flags),
                               lookupStatus);
    if (FAILED(lookupStatus) || result == 0 || sfi.iIcon < 0)
    {
        if (! useFileAttributes)
        {
            PerfEmitCounter(L"iconcache.path_live_lookup_failed_uncached", 1u);
        }

        const auto failureStamp  = std::chrono::steady_clock::now();
        const auto lockWaitStart = std::chrono::steady_clock::now();
        std::lock_guard lock(_mutex);
        PerfEmitSlowIconCacheLockWait(L"path_failure_store", PerfElapsedUs(lockWaitStart));
        const auto lockHoldStart = std::chrono::steady_clock::now();
        auto lockHold            = wil::scope_exit([&] { PerfEmitSlowIconCacheLockHold(L"path_failure_store", PerfElapsedUs(lockHoldStart)); });
        EvictPathQueryBatch();
        const auto existing = _pathToIconIndex.find(cacheKey);
        if (existing == _pathToIconIndex.end())
        {
            static_cast<void>(_pathToIconIndex.emplace(
                std::move(cacheKey), PathIconCacheEntry{-1, ++_pathQueryAccessCounter, true, failureStamp, 1u}));
            PerfEmitCounter(L"iconcache.path_failed_lookup_cached", 1u);
        }
        else if (existing->second.lookupFailed)
        {
            existing->second.failureStamp = failureStamp;
            existing->second.consecutiveFailureCount =
                existing->second.consecutiveFailureCount == UINT32_MAX ? UINT32_MAX : existing->second.consecutiveFailureCount + 1u;
            existing->second.lastAccessTime = ++_pathQueryAccessCounter;
            PerfEmitCounter(L"iconcache.path_failed_lookup_cached", 1u);
        }
        else
        {
            existing->second.lastAccessTime = ++_pathQueryAccessCounter;
            PerfEmitCounter(L"iconcache.duplicate_path_query_race", 1u);
        }
        return std::nullopt;
    }

    bool duplicateRace       = false;
    const auto lockWaitStart = std::chrono::steady_clock::now();
    std::lock_guard lock(_mutex);
    PerfEmitSlowIconCacheLockWait(L"path_store", PerfElapsedUs(lockWaitStart));
    const auto lockHoldStart = std::chrono::steady_clock::now();
    auto lockHold            = wil::scope_exit([&] { PerfEmitSlowIconCacheLockHold(L"path_store", PerfElapsedUs(lockHoldStart)); });
    EvictPathQueryBatch();
    auto [it, inserted] = _pathToIconIndex.emplace(std::move(cacheKey), PathIconCacheEntry{sfi.iIcon, ++_pathQueryAccessCounter, false, {}, 0u});
    if (! inserted)
    {
        duplicateRace             = true;
        it->second.iconIndex      = sfi.iIcon;
        it->second.lookupFailed   = false;
        it->second.failureStamp   = {};
        it->second.consecutiveFailureCount = 0u;
        it->second.lastAccessTime = ++_pathQueryAccessCounter;
    }
    if (duplicateRace)
    {
        PerfEmitCounter(L"iconcache.duplicate_path_query_race", 1);
    }
    return sfi.iIcon;
}

std::optional<int> IconCache::QuerySysIconIndexForPidl(PCIDLIST_ABSOLUTE pidl) const
{
    if (! pidl)
    {
        return std::nullopt;
    }

    SHFILEINFOW sfi{};
    const DWORD_PTR result = SHGetFileInfoW(reinterpret_cast<LPCWSTR>(pidl), 0, &sfi, sizeof(sfi), SHGFI_PIDL | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
    if (result == 0 || sfi.iIcon < 0)
    {
        return std::nullopt;
    }

    return sfi.iIcon;
}

std::optional<int> IconCache::QuerySysIconIndexForKnownFolder(const GUID& folderId) const
{
    PIDLIST_ABSOLUTE pidl = nullptr;
    const HRESULT hr      = SHGetKnownFolderIDList(folderId, 0, nullptr, &pidl);
    if (FAILED(hr) || ! pidl)
    {
        return std::nullopt;
    }

    auto pidlCleanup = wil::scope_exit([&] { ILFree(pidl); });
    return QuerySysIconIndexForPidl(pidl);
}

std::optional<IconCache::SpecialFolderMatch> IconCache::TryGetSpecialFolderForPathPrefix(std::wstring_view path) const
{
    if (path.empty())
    {
        return std::nullopt;
    }

    std::call_once(_specialFoldersInitOnce, [] { InitializeSpecialFolders(); });

    const std::wstring_view fullPath = path;
    const std::wstring* bestPath     = nullptr;

    for (const auto& specialPath : _specialFolderPaths)
    {
        const std::wstring_view specialView = specialPath;
        if (fullPath.size() < specialView.size())
        {
            continue;
        }

        if (wil::compare_string_ordinal(fullPath.substr(0, specialView.size()), specialView, true) != wistd::weak_ordering::equivalent)
        {
            continue;
        }

        if (fullPath.size() != specialView.size())
        {
            const wchar_t next = fullPath[specialView.size()];
            if (next != L'\\' && next != L'/')
            {
                continue;
            }
        }

        if (! bestPath || specialView.size() > bestPath->size())
        {
            bestPath = &specialPath;
        }
    }

    if (! bestPath)
    {
        return std::nullopt;
    }

    SpecialFolderMatch match;
    match.rootPath = *bestPath;

    const auto it = _specialFolderIconCache.find(match.rootPath);
    if (it != _specialFolderIconCache.end())
    {
        match.iconIndex = it->second;
    }

    return match;
}

wil::com_ptr<ID2D1Bitmap1> IconCache::ConvertIconToBitmapOnUIThread(HICON icon, int iconIndex, ID2D1DeviceContext* d2dContext, float targetDipSize)
{
    if (! icon || ! d2dContext || iconIndex < 0)
    {
        return nullptr;
    }

    wil::com_ptr<ID2D1Device> device;
    d2dContext->GetDevice(device.put());
    if (! device)
    {
        return nullptr;
    }

    const IconBitmapCacheKey cacheKey{.iconIndex = iconIndex, .imageListSize = MakeBitmapCacheSizeClass(targetDipSize)};

    // Check if already cached (another thread might have added it)
    {
        const auto lockWaitStart = std::chrono::steady_clock::now();
        std::lock_guard lock(_mutex);
        PerfEmitSlowIconCacheLockWait(L"ui_convert_lookup", PerfElapsedUs(lockWaitStart));
        const auto lockHoldStart = std::chrono::steady_clock::now();
        auto lockHold            = wil::scope_exit([&] { PerfEmitSlowIconCacheLockHold(L"ui_convert_lookup", PerfElapsedUs(lockHoldStart)); });
        const auto deviceIt      = _deviceCaches.find(device.get());
        if (deviceIt != _deviceCaches.end())
        {
            const auto it = deviceIt->second.bitmaps.find(cacheKey);
            if (it != deviceIt->second.bitmaps.end())
            {
                PerfEmitCounter(L"iconcache.ui_convert_hit_after_race", 1);
                return it->second.bitmap;
            }
        }
    }

    // Convert HICON to D2D bitmap (UI thread only)
    auto bitmap = ConvertIconToBitmap(icon, d2dContext);

    if (bitmap)
    {
        const D2D1_SIZE_U pixelSize = bitmap->GetPixelSize();
        const size_t bytes          = static_cast<size_t>(pixelSize.width) * static_cast<size_t>(pixelSize.height) * 4u;

        // Store in cache with LRU tracking
        const auto lockWaitStart = std::chrono::steady_clock::now();
        std::lock_guard lock(_mutex);
        PerfEmitSlowIconCacheLockWait(L"ui_convert_store", PerfElapsedUs(lockWaitStart));
        const auto lockHoldStart = std::chrono::steady_clock::now();
        auto lockHold            = wil::scope_exit([&] { PerfEmitSlowIconCacheLockHold(L"ui_convert_store", PerfElapsedUs(lockHoldStart)); });
        auto& cache              = _deviceCaches[device.get()];
        if (! cache.device)
        {
            cache.device = device;
        }
        EvictLRUIfNeeded(cache);
        CacheEntry entry;
        entry.bitmap            = bitmap;
        entry.lastAccessTime    = ++cache.accessCounter;
        entry.bytes             = bytes;
        cache.bitmaps[cacheKey] = std::move(entry);
    }

    return bitmap;
}

void IconCache::Clear()
{
    std::lock_guard lock(_mutex);
    size_t iconCount = 0;
    for (const auto& entry : _deviceCaches)
    {
        iconCount += entry.second.bitmaps.size();
    }
    const size_t extCount  = _extensionToIconIndex.size();
    const size_t pathCount = _pathToIconIndex.size();
    _deviceCaches.clear();
    _extensionToIconIndex.clear();
    _pathToIconIndex.clear();
    _pathQueryAccessCounter = 0;
    _pathLruEvictions       = 0;

    DBGOUT_INFO(L"IconCache: Cleared {} cached icons, {} extension mappings, and {} path mappings", iconCount, extCount, pathCount);
}

void IconCache::ClearAssociationCache() noexcept
{
    const auto clearStart   = std::chrono::steady_clock::now();
    size_t associationCount = 0u;
    {
        std::lock_guard lock(g_associationCacheMutex);
        associationCount = g_associationToIconIndex.size();
        g_associationToIconIndex.clear();
        g_associationQueryAccessCounter = 0u;
        g_associationLruEvictions       = 0u;
    }

    PerfEmitDuration(L"iconcache.association_clear_us", PerfElapsedUs(clearStart), static_cast<uint64_t>(associationCount), 0u, S_OK);
}

void IconCache::ClearDeviceCache(ID2D1Device* device)
{
    if (! device)
    {
        return;
    }

    std::lock_guard lock(_mutex);
    const auto it = _deviceCaches.find(device);
    if (it == _deviceCaches.end())
    {
        return;
    }

    const size_t bitmapCount = it->second.bitmaps.size();
    _deviceCaches.erase(it);
    DBGOUT_INFO(L"IconCache: Cleared device cache ({} bitmaps)", bitmapCount);
}

void IconCache::WarmCommonExtensions()
{
    Debug::Perf::Scope perf(L"IconCache.WarmCommonExtensions");
    TRACER_CTX(L"----------------");

    if (_warmingCompleted.load(std::memory_order_acquire))
    {
        return;
    }

    {
        bool expected = false;
        if (! _warmingInProgress.compare_exchange_strong(expected, true, std::memory_order_acq_rel))
        {
            return; // Already warming
        }
    }
    auto clearWarmingInProgress = wil::scope_exit([&] { _warmingInProgress.store(false, std::memory_order_release); });

    DBGOUT_INFO(L"IconCache: Starting lazy cache warming...");

    // Common file extensions to pre-cache (most frequently encountered)
    static constexpr std::pair<std::wstring_view, DWORD> commonExtensions[] = {
        {L".txt", FILE_ATTRIBUTE_NORMAL},   {L".log", FILE_ATTRIBUTE_NORMAL},    {L".xml", FILE_ATTRIBUTE_NORMAL},  {L".json", FILE_ATTRIBUTE_NORMAL},
        {L".jsonl", FILE_ATTRIBUTE_NORMAL}, {L".ndjson", FILE_ATTRIBUTE_NORMAL}, {L".ini", FILE_ATTRIBUTE_NORMAL},  {L".cfg", FILE_ATTRIBUTE_NORMAL},
        {L".md", FILE_ATTRIBUTE_NORMAL},    {L".cpp", FILE_ATTRIBUTE_NORMAL},    {L".h", FILE_ATTRIBUTE_NORMAL},    {L".hpp", FILE_ATTRIBUTE_NORMAL},
        {L".c", FILE_ATTRIBUTE_NORMAL},     {L".cs", FILE_ATTRIBUTE_NORMAL},     {L".py", FILE_ATTRIBUTE_NORMAL},   {L".js", FILE_ATTRIBUTE_NORMAL},
        {L".ts", FILE_ATTRIBUTE_NORMAL},    {L".html", FILE_ATTRIBUTE_NORMAL},   {L".htm", FILE_ATTRIBUTE_NORMAL},  {L".css", FILE_ATTRIBUTE_NORMAL},
        {L".pdf", FILE_ATTRIBUTE_NORMAL},   {L".zip", FILE_ATTRIBUTE_NORMAL},    {L".rar", FILE_ATTRIBUTE_NORMAL},  {L".7z", FILE_ATTRIBUTE_NORMAL},
        {L".png", FILE_ATTRIBUTE_NORMAL},   {L".jpg", FILE_ATTRIBUTE_NORMAL},    {L".jpeg", FILE_ATTRIBUTE_NORMAL}, {L".gif", FILE_ATTRIBUTE_NORMAL},
        {L".bmp", FILE_ATTRIBUTE_NORMAL},   {L".ico", FILE_ATTRIBUTE_NORMAL},    {L".svg", FILE_ATTRIBUTE_NORMAL},  {L".mp3", FILE_ATTRIBUTE_NORMAL},
        {L".wav", FILE_ATTRIBUTE_NORMAL},   {L".mp4", FILE_ATTRIBUTE_NORMAL},    {L".avi", FILE_ATTRIBUTE_NORMAL},  {L".mkv", FILE_ATTRIBUTE_NORMAL},
        {L".doc", FILE_ATTRIBUTE_NORMAL},   {L".docx", FILE_ATTRIBUTE_NORMAL},   {L".xls", FILE_ATTRIBUTE_NORMAL},  {L".xlsx", FILE_ATTRIBUTE_NORMAL},
        {L".ppt", FILE_ATTRIBUTE_NORMAL},   {L".pptx", FILE_ATTRIBUTE_NORMAL},   {L".dll", FILE_ATTRIBUTE_NORMAL},  {L".sys", FILE_ATTRIBUTE_NORMAL},
        {L".bat", FILE_ATTRIBUTE_NORMAL},   {L".cmd", FILE_ATTRIBUTE_NORMAL},    {L".ps1", FILE_ATTRIBUTE_NORMAL},  {L"<directory>", FILE_ATTRIBUTE_DIRECTORY},
    };

    size_t warmed = 0;
    for (const auto& [ext, attrib] : commonExtensions)
    {
        std::wstring extKey(ext);
        {
            std::lock_guard lock(_mutex);
            if (_extensionToIconIndex.find(extKey) != _extensionToIconIndex.end())
            {
                continue;
            }
        }

        const auto iconIndex = QueryAssociationIconIndex(extKey, attrib);
        if (iconIndex.has_value())
        {
            std::lock_guard lock(_mutex);
            if (_extensionToIconIndex.find(extKey) == _extensionToIconIndex.end())
            {
                _extensionToIconIndex[std::move(extKey)] = iconIndex.value();
                ++warmed;
            }
        }
    }

    _warmingCompleted.store(true, std::memory_order_release);
    perf.SetValue0(warmed);

    DBGOUT_INFO(L"IconCache: Lazy warming completed - {} extensions cached", warmed);
}

size_t IconCache::PrewarmBitmaps(ID2D1DeviceContext* d2dContext)
{
    if (! d2dContext)
    {
        return 0;
    }

    TRACER_CTX(L"PrewarmBitmaps");

    wil::com_ptr<ID2D1Device> device;
    d2dContext->GetDevice(device.put());
    if (! device)
    {
        return 0;
    }

    if (! _warmingCompleted.load(std::memory_order_acquire))
    {
        WarmCommonExtensions();
    }

    // Collect unique icon indices to prewarm
    std::vector<int> iconIndices;
    {
        const auto lockWaitStart = std::chrono::steady_clock::now();
        std::lock_guard lock(_mutex);
        PerfEmitSlowIconCacheLockWait(L"prewarm_snapshot", PerfElapsedUs(lockWaitStart));
        const auto lockHoldStart = std::chrono::steady_clock::now();
        auto lockHold            = wil::scope_exit([&] { PerfEmitSlowIconCacheLockHold(L"prewarm_snapshot", PerfElapsedUs(lockHoldStart)); });
        std::unordered_set<int> uniqueIndices;
        for (const auto& [ext, iconIndex] : _extensionToIconIndex)
        {
            if (iconIndex >= 0)
            {
                uniqueIndices.insert(iconIndex);
            }
        }
        iconIndices.assign(uniqueIndices.begin(), uniqueIndices.end());
    }

    if (iconIndices.empty())
    {
        return 0;
    }

    DBGOUT_INFO(L"IconCache: Pre-warming {} D2D bitmaps...", iconIndices.size());

    size_t created = 0;
    for (int iconIndex : iconIndices)
    {
        // Check if already cached for this device
        if (HasCachedIcon(iconIndex, device.get(), 16.0f))
        {
            continue;
        }

        // Extract icon and convert to D2D bitmap
        wil::unique_hicon hIcon = ExtractSystemIcon(iconIndex, 16.0f);
        if (hIcon)
        {
            auto bitmap = ConvertIconToBitmapOnUIThread(hIcon.get(), iconIndex, d2dContext, 16.0f);
            if (bitmap)
            {
                ++created;
            }
        }
    }

    DBGOUT_INFO(L"IconCache: Pre-warmed {} D2D bitmaps", created);
    return created;
}

IconCache::Stats IconCache::GetStats() const
{
    std::lock_guard lock(_mutex);
    Stats stats;
    for (const auto& entry : _deviceCaches)
    {
        stats.cacheSize += entry.second.bitmaps.size();
    }
    stats.hitCount           = _hitCount;
    stats.missCount          = _missCount;
    stats.extensionCacheSize = _extensionToIconIndex.size();
    stats.pathCacheSize      = _pathToIconIndex.size();
    stats.lruEvictions       = _lruEvictions;
    stats.pathLruEvictions   = _pathLruEvictions;
    return stats;
}

std::optional<int> IconCache::GetIconIndexByExtension(std::wstring_view extension) const
{
    const std::wstring key = NormalizeExtensionKey(extension);
    std::lock_guard lock(_mutex);
    auto it = _extensionToIconIndex.find(key);
    if (it != _extensionToIconIndex.end())
    {
        return it->second;
    }
    return std::nullopt;
}

void IconCache::RegisterExtension(std::wstring_view extension, int iconIndex)
{
    if (iconIndex < 0)
    {
        return;
    }

    const std::wstring key = NormalizeExtensionKey(extension);
    std::lock_guard lock(_mutex);
    _extensionToIconIndex[key] = iconIndex;
}

std::optional<int> IconCache::GetOrQueryIconIndexByExtension(std::wstring_view extension, DWORD fileAttributes)
{
    if (const auto cachedIndex = GetIconIndexByExtension(extension); cachedIndex.has_value())
    {
        return cachedIndex;
    }

    if (RequiresPerFileLookup(extension))
    {
        return std::nullopt;
    }

    const std::wstring key = NormalizeExtensionKey(extension);
    const std::wstring_view keyView{key};
    const bool isFolder = IsDirectoryExtensionKey(keyView);
    auto iconIndex      = QueryAssociationIconIndex(keyView, fileAttributes);
    if (! iconIndex.has_value() && ! isFolder && ! keyView.empty())
    {
        iconIndex = QueryCachedDefaultAssociationIconIndex();
    }
    if (iconIndex.has_value())
    {
        RegisterExtension(keyView, iconIndex.value());
    }
    return iconIndex;
}

bool IconCache::RequiresPerFileLookup(std::wstring_view extension) const
{
    const std::wstring key = NormalizeExtensionKey(extension);
    const std::wstring_view keyView{key};
    if (keyView.empty() || keyView == kDirectoryExtensionKey)
    {
        return false;
    }

    // Per-file lookup: files with embedded or per-path icons (not stable by extension)
    static constexpr std::wstring_view kPerFileLookupExtensions[] = {
        L".exe",
        L".ico",
        L".lnk",
        L".url",
        L".dll",
        L".cpl",
        L".scr",
        L".msc",
        L".ocx",
    };

    for (const auto ext : kPerFileLookupExtensions)
    {
        if (keyView == ext)
        {
            return true;
        }
    }

    return false;
}

void IconCache::SetMaxCacheSize(size_t maxSize)
{
    std::lock_guard lock(_mutex);
    _maxCacheSize = maxSize;
}

void IconCache::EvictLRUIfNeeded(IconCache::DeviceCache& cache)
{
    // Must be called with _mutex locked
    if (cache.bitmaps.size() < _maxCacheSize)
    {
        return;
    }

    const auto scanStart = std::chrono::steady_clock::now();
    // Find oldest entry by access time
    IconBitmapCacheKey oldestKey{};
    bool haveOldest   = false;
    size_t oldestTime = SIZE_MAX;

    for (const auto& [key, entry] : cache.bitmaps)
    {
        if (entry.lastAccessTime < oldestTime)
        {
            oldestTime = entry.lastAccessTime;
            oldestKey  = key;
            haveOldest = true;
        }
    }

    if (haveOldest)
    {
        cache.bitmaps.erase(oldestKey);
        _lruEvictions++;
        DBGOUT_INFO(
            L"IconCache: Evicted icon index {} sizeClass {} (LRU), cache size now {}", oldestKey.iconIndex, oldestKey.imageListSize, cache.bitmaps.size());
    }

    PerfEmitDuration(L"iconcache.device_lru_evict_scan_us",
                     PerfElapsedUs(scanStart),
                     static_cast<uint64_t>(cache.bitmaps.size()),
                     static_cast<uint64_t>(haveOldest ? oldestKey.iconIndex : 0),
                     S_OK);
}

void IconCache::EvictPathQueryBatch()
{
    // Must be called with _mutex locked.
    // Evicts ~10% of entries (the oldest ones) to amortize the O(n) scan cost.
    if (_pathToIconIndex.size() < _maxCacheSize)
    {
        return;
    }

    const auto scanStart    = std::chrono::steady_clock::now();
    const size_t evictCount = std::max<size_t>(_pathToIconIndex.size() / 10, 1);

    // Find the evictCount-th smallest lastAccessTime as the threshold.
    // Collect all access times, partial-sort, then erase entries at or below the threshold.
    std::vector<size_t> times;
    times.reserve(_pathToIconIndex.size());
    for (const auto& [key, entry] : _pathToIconIndex)
    {
        times.push_back(entry.lastAccessTime);
    }

    std::nth_element(times.begin(), times.begin() + static_cast<ptrdiff_t>(evictCount - 1), times.end());
    const size_t threshold = times[evictCount - 1];

    size_t removed = 0;
    for (auto it = _pathToIconIndex.begin(); it != _pathToIconIndex.end();)
    {
        if (it->second.lastAccessTime <= threshold && removed < evictCount)
        {
            it = _pathToIconIndex.erase(it);
            ++removed;
        }
        else
        {
            ++it;
        }
    }

    _pathLruEvictions += removed;
    PerfEmitDuration(
        L"iconcache.path_lru_evict_scan_us", PerfElapsedUs(scanStart), static_cast<uint64_t>(_pathToIconIndex.size()), static_cast<uint64_t>(removed), S_OK);
}

wil::com_ptr<ID2D1Bitmap1> IconCache::ConvertIconToBitmap(HICON icon, ID2D1DeviceContext* d2dContext)
{
    if (! icon || ! d2dContext)
    {
        return nullptr;
    }

    ICONINFO iconInfo{};
    if (! GetIconInfo(icon, &iconInfo))
    {
        Debug::Warning(L"IconCache: Failed to read HICON metadata: 0x{:08X}", HRESULT_FROM_WIN32(GetLastError()));
        return nullptr;
    }

    wil::unique_hbitmap colorBitmap(iconInfo.hbmColor);
    wil::unique_hbitmap maskBitmap(iconInfo.hbmMask);

    const HBITMAP metricsBitmap = colorBitmap ? colorBitmap.get() : maskBitmap.get();
    BITMAP metrics{};
    if (! metricsBitmap || ! GetObjectW(metricsBitmap, sizeof(metrics), &metrics) || metrics.bmWidth <= 0 || metrics.bmHeight <= 0)
    {
        Debug::Warning(L"IconCache: Failed to resolve HICON dimensions");
        return nullptr;
    }

    const int width = metrics.bmWidth;
    int height      = metrics.bmHeight;
    if (! colorBitmap)
    {
        height = std::max(height / 2, 1);
    }

    constexpr int kMaxIconBitmapDimension = 512;
    if (width > kMaxIconBitmapDimension || height > kMaxIconBitmapDimension)
    {
        Debug::Warning(L"IconCache: Refusing oversized HICON bitmap {}x{}", width, height);
        return nullptr;
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = width;
    bmi.bmiHeader.biHeight      = -height;
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    wil::unique_hdc_window screenDc{GetDC(nullptr)};
    if (! screenDc)
    {
        Debug::Warning(L"IconCache: Failed to acquire screen DC for icon conversion");
        return nullptr;
    }

    wil::unique_hdc memoryDc{CreateCompatibleDC(screenDc.get())};
    if (! memoryDc)
    {
        Debug::Warning(L"IconCache: Failed to create memory DC for icon conversion");
        return nullptr;
    }

    void* bits = nullptr;
    wil::unique_hbitmap dib{CreateDIBSection(screenDc.get(), &bmi, DIB_RGB_COLORS, &bits, nullptr, 0)};
    if (! dib || ! bits)
    {
        Debug::Warning(L"IconCache: Failed to create DIB section for icon conversion");
        return nullptr;
    }

    const size_t pixelCount = static_cast<size_t>(width) * static_cast<size_t>(height);
    auto* pixels            = static_cast<BYTE*>(bits);
    std::fill_n(pixels, pixelCount * 4u, BYTE{0});

    auto oldBitmap        = wil::SelectObject(memoryDc.get(), dib.get());
    const auto drawStart  = std::chrono::steady_clock::now();
    const BOOL drawResult = DrawIconEx(memoryDc.get(), 0, 0, icon, width, height, 0, nullptr, DI_NORMAL);
    const DWORD drawError = drawResult ? ERROR_SUCCESS : GetLastError();
    oldBitmap.reset();
    PerfEmitDuration(L"iconcache.convert_draw_icon_us",
                     PerfElapsedUs(drawStart),
                     static_cast<uint64_t>(width),
                     static_cast<uint64_t>(height),
                     drawResult ? S_OK : HRESULT_FROM_WIN32(drawError == ERROR_SUCCESS ? ERROR_INVALID_DATA : drawError));
    if (! drawResult)
    {
        Debug::Warning(L"IconCache: Failed to draw HICON into DIB: 0x{:08X}", HRESULT_FROM_WIN32(drawError == ERROR_SUCCESS ? ERROR_INVALID_DATA : drawError));
        return nullptr;
    }

    NormalizeIconBitmapAlphaForD2D(pixels, pixelCount, maskBitmap.get(), width, height, screenDc.get());

    D2D1_BITMAP_PROPERTIES1 bitmapProps{};
    bitmapProps.pixelFormat.format    = DXGI_FORMAT_B8G8R8A8_UNORM;
    bitmapProps.pixelFormat.alphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED;
    const float dpi                   = _dpi.load(std::memory_order_relaxed);
    bitmapProps.dpiX                  = dpi;
    bitmapProps.dpiY                  = dpi;
    bitmapProps.bitmapOptions         = D2D1_BITMAP_OPTIONS_NONE;

    wil::com_ptr<ID2D1Bitmap1> d2dBitmap;
    const auto d2dCreateStart = std::chrono::steady_clock::now();
    const HRESULT hr          = d2dContext->CreateBitmap(
        D2D1::SizeU(static_cast<UINT32>(width), static_cast<UINT32>(height)), bits, static_cast<UINT32>(width) * 4u, &bitmapProps, d2dBitmap.put());
    PerfEmitDuration(L"iconcache.convert_create_d2d_bitmap_us", PerfElapsedUs(d2dCreateStart), static_cast<uint64_t>(width), static_cast<uint64_t>(height), hr);
    if (FAILED(hr))
    {
        Debug::Warning(L"IconCache: Failed to create D2D bitmap from HICON DIB: 0x{:08X}", hr);
        return nullptr;
    }

    return d2dBitmap;
}

int IconCache::SelectOptimalImageListSize(float targetDipSize) const
{
    // Calculate target size in physical pixels
    const float dpi          = _dpi.load(std::memory_order_relaxed);
    const float targetPixels = targetDipSize * dpi / 96.0f;

    // Prefer the smallest shell source that is at least as large as the physical target.
    // Downscaling a larger shell icon is cleaner than upscaling a 16px source on high-DPI displays.
    return SelectImageListSizeForTargetPixels(targetPixels);
}

// Static member definitions
std::unordered_map<std::wstring, int> IconCache::_specialFolderIconCache;
std::unordered_set<std::wstring> IconCache::_specialFolderPaths;
std::once_flag IconCache::_specialFoldersInitOnce;

size_t IconCache::GetMemoryUsage() const
{
    std::lock_guard lock(_mutex);
    size_t bytes = 0;
    for (const auto& entry : _deviceCaches)
    {
        for (const auto& [key, cacheEntry] : entry.second.bitmaps)
        {
            static_cast<void>(key);
            bytes += cacheEntry.bytes;
        }
    }
    return bytes;
}

bool IconCache::IsSpecialFolder(const std::wstring& path)
{
    std::call_once(_specialFoldersInitOnce, [] { InitializeSpecialFolders(); });

    // Use WIL's case-insensitive string comparison for Windows file system behavior
    // Linear search is acceptable for a small fixed set of special folders (O(n) where n is small)
    std::wstring_view pathView = path;
    for (const auto& specialPath : _specialFolderPaths)
    {
        if (wil::compare_string_ordinal(pathView, specialPath, true) == wistd::weak_ordering::equivalent)
        {
            return true;
        }
    }
    return false;
}

void IconCache::InitializeSpecialFolders()
{
    // Get known folder paths and cache them for O(1) lookup
    const KNOWNFOLDERID knownFolders[] = {
        FOLDERID_Desktop, FOLDERID_Documents, FOLDERID_Downloads, FOLDERID_Pictures, FOLDERID_Music, FOLDERID_Videos, FOLDERID_SkyDrive};

    for (const auto& folderId : knownFolders)
    {
        wil::unique_cotaskmem_string folderPath;
        const HRESULT hr = SHGetKnownFolderPath(folderId, 0, nullptr, folderPath.put());
        if (SUCCEEDED(hr) && folderPath)
        {
            std::wstring path = folderPath.get();
            _specialFolderPaths.insert(path);

            // Cache icon index for this known folder using PIDL for best fidelity.
            if (const auto iconIndex = IconCache::GetInstance().QuerySysIconIndexForKnownFolder(folderId); iconIndex.has_value())
            {
                _specialFolderIconCache.emplace(std::move(path), iconIndex.value());
            }
        }
    }

    Debug::Info(L"IconCache: Initialized {} special folder paths", _specialFolderPaths.size());
}
