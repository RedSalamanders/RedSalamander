#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include <atomic>
#include <cstddef>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>

#include <ShlObj.h>
#include <Windows.h>

#include <d2d1_1.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027
// (move assign deleted), C4820 (padding)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

// Forward declarations
struct IImageList;
struct IWICImagingFactory;

// Application-wide icon cache using Windows system image lists.
// Converts HICON to ID2D1Bitmap1 and caches by icon index (per Direct2D device) for sharing across views.
// Thread-safe for concurrent access from multiple windows and background threads.
class IconCache
{
public:
    // Get the singleton instance of the icon cache.
    static IconCache& GetInstance();

    // Initialize the cache (must be called before first use).
    // Contract:
    // - UI thread responsibility: call after COM is initialized as STA (CoInitializeEx(COINIT_APARTMENTTHREADED)).
    // - Caches system image list COM pointers and a WIC factory for process lifetime.
    // d2dContext: Any valid D2D device context (guard for initialization only).
    // dpi: Current DPI for bitmap creation.
    void Initialize(ID2D1DeviceContext* d2dContext, float dpi);

    // Warm cache with common file extensions on application startup.
    // This prefetches icon indices for frequently encountered file types.
    void WarmCommonExtensions();

    // Pre-create D2D bitmaps for all cached icon indices (call after WarmCommonExtensions).
    // This converts HICON to D2D bitmaps so they're ready when needed.
    // d2dContext: D2D device context for bitmap creation
    // Returns number of bitmaps created
    size_t PrewarmBitmaps(ID2D1DeviceContext* d2dContext);

    // Update DPI for future bitmap conversions (call on DPI change).
    // Note: Existing cached bitmaps are not updated; clear cache if needed.
    void SetDpi(float dpi);

    // Get or create a D2D bitmap for the given system icon index.
    // iconIndex: System image list icon index from SHGetFileInfo
    // d2dContext: D2D device context used to create/cache the bitmap (cache is per ID2D1Device)
    // Cached or newly created D2D bitmap, or nullptr on failure</returns>
    wil::com_ptr<ID2D1Bitmap1> GetIconBitmap(int iconIndex, ID2D1DeviceContext* d2dContext);

    // Check if icon is already cached for the given D2D device (thread-safe, no D2D calls).
    bool HasCachedIcon(int iconIndex, ID2D1Device* device) const;

    // Get cached bitmap without creating (returns nullptr if not cached).
    wil::com_ptr<ID2D1Bitmap1> GetCachedBitmap(int iconIndex, ID2D1DeviceContext* d2dContext) const;

    // Extract an icon from the system image list.
    // Contract:
    // - Requires IconCache::Initialize(...) to have run.
    // - Caller thread must have COM initialized. UI thread is STA; worker threads must initialize COM as MTA
    //   (e.g. `[[maybe_unused]] auto coInit = wil::CoInitializeEx(COINIT_MULTITHREADED);`).
    // targetDipSize: Target icon size in DIPs (e.g., 16.0f, 32.0f, 48.0f). Defaults to 16.0f (FolderView default).
    // Automatically selects the optimal system image list based on DPI scaling.
    wil::unique_hicon ExtractSystemIcon(int iconIndex, float targetDipSize = 16.0f);

    // Convert HICON to D2D bitmap on UI thread and cache it.
    // Must be called on UI thread with valid D2D context.
    wil::com_ptr<ID2D1Bitmap1> ConvertIconToBitmapOnUIThread(HICON icon, int iconIndex, ID2D1DeviceContext* d2dContext);

    // Create GDI HBITMAP from HICON using WIC for menu icons (UI thread only).
    // Returns 32-bit premultiplied BGRA bitmap suitable for SetMenuItemBitmaps.
    // Uses WIC pipeline with nearest-neighbor scaling and proper alpha channel handling.
    // Fixes black dots/borders by copying from converter output instead of original bitmap.
    // size: Pixel size of bitmap (typically GetSystemMetrics(SM_CXSMICON) at current DPI)
    // Returns nullptr on failure. Caller owns returned bitmap via wil::unique_hbitmap.
    wil::unique_hbitmap CreateMenuBitmapFromIcon(HICON icon, int size);

    // Create a menu bitmap directly from a system icon index (UI thread only).
    // Extracts an appropriately sized icon for the current DPI to avoid upscaling 16×16 on high DPI.
    wil::unique_hbitmap CreateMenuBitmapFromIconIndex(int iconIndex, int size);

    // Create a menu bitmap from a file system path (or UNC path).
    // Applies special-case icon resolution for known special folders and WSL distributions.
    // fileAttributes/useFileAttributes are forwarded to SHGetFileInfo fallback behavior.
    wil::unique_hbitmap CreateMenuBitmapFromPath(const wchar_t* path, int size, DWORD fileAttributes = 0, bool useFileAttributes = false);

    // Get icon index by file extension (fast cache lookup, no Shell API calls).
    // Returns nullopt if extension not cached.
    // extension: File extension including dot (e.g., L".txt") or special key for directories
    std::optional<int> GetIconIndexByExtension(std::wstring_view extension) const;

    // Register extension to icon index mapping for future lookups.
    // extension: File extension including dot (e.g., L".txt") or special key
    // iconIndex: System image list icon index from SHGetFileInfo
    void RegisterExtension(std::wstring_view extension, int iconIndex);

    // Get cached icon index or query + register it using a dummy path (thread-safe).
    // Uses SHGFI_USEFILEATTRIBUTES so the dummy path does not need to exist.
    // fileAttributes: FILE_ATTRIBUTE_DIRECTORY for L"<directory>", FILE_ATTRIBUTE_NORMAL for files.
    std::optional<int> GetOrQueryIconIndexByExtension(std::wstring_view extension, DWORD fileAttributes);

    // Check if extension requires per-file icon lookup (e.g., .exe, .ico, .lnk).
    // These file types have unique icons per file, not by extension.
    bool RequiresPerFileLookup(std::wstring_view extension) const;

    // Set maximum cache size for LRU eviction (default: 2000 unique icons; memory varies by extracted size).
    void SetMaxCacheSize(size_t maxSize);

    // Clear all cached bitmaps (call on device loss or when releasing resources).
    void Clear();

    // Clear cached bitmaps for a specific D2D device (call when discarding that device).
    // This prevents stale device caches from keeping old ID2D1Device instances alive.
    void ClearDeviceCache(ID2D1Device* device);

    // Get cache statistics for debugging/monitoring.
    struct Stats
    {
        size_t cacheSize          = 0;
        size_t hitCount           = 0;
        size_t missCount          = 0;
        size_t extensionCacheSize = 0;
        size_t lruEvictions       = 0;
    };
    Stats GetStats() const;
    // Get approximate bitmap memory usage in bytes (sum of width × height × 4 for cached entries).
    size_t GetMemoryUsage() const;

    // Check if path is a special folder (Desktop, Documents, etc.)
    static bool IsSpecialFolder(const std::wstring& path);

    // Query system image list icon index for a path (thread-safe).
    // fileAttributes is only used when useFileAttributes is true.
    std::optional<int> QuerySysIconIndexForPath(const wchar_t* path, DWORD fileAttributes, bool useFileAttributes) const;

    // Query system icon index for a PIDL (thread-safe).
    std::optional<int> QuerySysIconIndexForPidl(PCIDLIST_ABSOLUTE pidl) const;

    // Query system icon index for a known folder via PIDL (thread-safe).
    std::optional<int> QuerySysIconIndexForKnownFolder(const GUID& folderId) const;

    struct SpecialFolderMatch
    {
        std::wstring rootPath;
        int iconIndex = -1;
    };

    // Boundary-aware, case-insensitive prefix match against known special folders.
    // Returns the matched special folder root and its cached icon index when available.
    std::optional<SpecialFolderMatch> TryGetSpecialFolderForPathPrefix(std::wstring_view path) const;

private:
    IconCache()                            = default;
    ~IconCache()                           = default;
    IconCache(const IconCache&)            = delete;
    IconCache& operator=(const IconCache&) = delete;

    // Select optimal system image list size based on target DIP size and current DPI.
    // targetDipSize: Target display size in DIPs (e.g., 16.0f, 48.0f)
    // Returns: SHIL_JUMBO (256×256), SHIL_EXTRALARGE (48×48), SHIL_LARGE (32×32), or SHIL_SMALL (16×16)
    int SelectOptimalImageListSize(float targetDipSize) const;

    // Convert HICON to ID2D1Bitmap1 using WIC for superior quality. WIC-based conversion provides crisp icons without GDI quality degradation.
    // Returns nullptr on WIC conversion failures (logged via Debug::Warning).
    wil::com_ptr<ID2D1Bitmap1> ConvertIconToBitmap(HICON icon, ID2D1DeviceContext* d2dContext);

    // LRU cache entry with access time tracking
    struct CacheEntry
    {
        wil::com_ptr<ID2D1Bitmap1> bitmap;
        size_t lastAccessTime = 0;
        size_t bytes          = 0; // Approximate: width × height × 4 (BGRA)
    };

    struct DeviceCache
    {
        wil::com_ptr<ID2D1Device> device;
        std::unordered_map<int, CacheEntry> bitmaps;
        size_t accessCounter = 0;
    };

    // Evict least recently used icon if cache exceeds size limit.
    void EvictLRUIfNeeded(DeviceCache& cache);

    mutable std::mutex _mutex;
    std::unordered_map<ID2D1Device*, DeviceCache> _deviceCaches;
    std::unordered_map<std::wstring, int> _extensionToIconIndex;

    std::atomic<float> _dpi{96.0f};
    std::atomic<bool> _initialized{false};
    size_t _maxCacheSize = 2000;

    // Statistics
    mutable size_t _hitCount     = 0;
    mutable size_t _missCount    = 0;
    mutable size_t _lruEvictions = 0;

    // System image list COM objects (cached) - initialized once and treated as immutable for lock-free reads in hot paths.
    // NOTE: Clear() does not reset these; they remain valid for the lifetime of the process once acquired.
    wil::com_ptr<IImageList> _systemImageListJumbo; // 256×256 (SHIL_JUMBO)
    wil::com_ptr<IImageList> _systemImageListXL;    // 48×48 (SHIL_EXTRALARGE)
    wil::com_ptr<IImageList> _systemImageListLarge; // 32×32 (SHIL_LARGE)
    wil::com_ptr<IImageList> _systemImageListSmall; // 16×16 (SHIL_SMALL)

    // WIC factory for high-quality icon conversion (thread-safe COM object)
    wil::com_ptr<IWICImagingFactory> _wicFactory;

    // Special folder icon caching
    static std::unordered_map<std::wstring, int> _specialFolderIconCache;
    static std::unordered_set<std::wstring> _specialFolderPaths;
    static std::once_flag _specialFoldersInitOnce;

    // Per-session extraction failure tracking
    std::unordered_map<std::wstring, int> _extractionFailureCount;

    // Lazy warming flags (atomic so callers can check state without taking the cache mutex).
    std::atomic<bool> _warmingCompleted{false};
    std::atomic<bool> _warmingInProgress{false};

    // Initialize special folder paths cache
    static void InitializeSpecialFolders();
};
