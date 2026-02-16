
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
#include "WSLDistro.h"

namespace
{
constexpr std::wstring_view kDirectoryExtensionKey = L"<directory>";
constexpr std::wstring_view kWslLocalhostPrefix    = L"\\\\wsl.localhost\\";
constexpr std::wstring_view kWslDollarPrefix       = L"\\\\wsl$\\";

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
} // namespace

IconCache& IconCache::GetInstance()
{
    static IconCache instance;
    return instance;
}

void IconCache::Initialize(ID2D1DeviceContext* d2dContext, float dpi)
{
    if (! d2dContext)
    {
        return;
    }

    std::lock_guard lock(_mutex);
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

wil::com_ptr<ID2D1Bitmap1> IconCache::GetIconBitmap(int iconIndex, ID2D1DeviceContext* d2dContext)
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

    // Check cache first
    {
        std::lock_guard lock(_mutex);
        auto deviceIt = _deviceCaches.find(device.get());
        if (deviceIt != _deviceCaches.end())
        {
            auto it = deviceIt->second.bitmaps.find(iconIndex);
            if (it != deviceIt->second.bitmaps.end())
            {
                _hitCount++;
                it->second.lastAccessTime = ++deviceIt->second.accessCounter;
                return it->second.bitmap;
            }
        }
        _missCount++;
    }

    // Cache miss - extract icon from system image list
    wil::unique_hicon icon = ExtractSystemIcon(iconIndex);
    if (! icon)
    {
        return nullptr;
    }

    // Convert to D2D bitmap
    auto bitmap = ConvertIconToBitmap(icon.get(), d2dContext);

    if (bitmap)
    {
        const D2D1_SIZE_U pixelSize = bitmap->GetPixelSize();
        const size_t bytes          = static_cast<size_t>(pixelSize.width) * static_cast<size_t>(pixelSize.height) * 4u;

        // Store in cache with LRU tracking
        std::lock_guard lock(_mutex);
        auto& cache = _deviceCaches[device.get()];
        if (! cache.device)
        {
            cache.device = device;
        }
        EvictLRUIfNeeded(cache);
        CacheEntry entry;
        entry.bitmap             = bitmap;
        entry.lastAccessTime     = ++cache.accessCounter;
        entry.bytes              = bytes;
        cache.bitmaps[iconIndex] = std::move(entry);
    }

    return bitmap;
}

bool IconCache::HasCachedIcon(int iconIndex, ID2D1Device* device) const
{
    if (iconIndex < 0 || ! device)
    {
        return false;
    }

    std::lock_guard lock(_mutex);
    const auto deviceIt = _deviceCaches.find(device);
    if (deviceIt == _deviceCaches.end())
    {
        return false;
    }

    return deviceIt->second.bitmaps.find(iconIndex) != deviceIt->second.bitmaps.end();
}

wil::com_ptr<ID2D1Bitmap1> IconCache::GetCachedBitmap(int iconIndex, ID2D1DeviceContext* d2dContext) const
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

    std::lock_guard lock(_mutex);
    const auto deviceIt = _deviceCaches.find(device.get());
    if (deviceIt == _deviceCaches.end())
    {
        return nullptr;
    }

    const auto it = deviceIt->second.bitmaps.find(iconIndex);
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
    if (! icon || ! _wicFactory || size <= 0)
    {
        return nullptr;
    }

    // Step 1: HICON → WIC bitmap (preserves alpha)
    wil::com_ptr<IWICBitmap> wicBitmap;
    HRESULT hr = _wicFactory->CreateBitmapFromHICON(icon, &wicBitmap);
    if (FAILED(hr))
    {
        Debug::Warning(L"IconCache: Failed to create WIC bitmap from HICON for menu: 0x{:08X}", hr);
        return nullptr;
    }

    // Get source bitmap dimensions
    UINT srcWidth  = 0;
    UINT srcHeight = 0;
    wicBitmap->GetSize(&srcWidth, &srcHeight);

    // Step 2: Scale to exact target size if needed (prevents blurry icons)
    wil::com_ptr<IWICBitmapSource> scaledSource = wicBitmap;
    wil::com_ptr<IWICBitmapScaler> scaler;

    if (srcWidth != static_cast<UINT>(size) || srcHeight != static_cast<UINT>(size))
    {
        hr = _wicFactory->CreateBitmapScaler(&scaler);
        if (SUCCEEDED(hr))
        {
            hr = scaler->Initialize(wicBitmap.get(), static_cast<UINT>(size), static_cast<UINT>(size), WICBitmapInterpolationModeNearestNeighbor);
            if (SUCCEEDED(hr))
            {
                scaledSource = scaler;
            }
        }
    }

    // Step 3: Convert to premultiplied BGRA (required for GDI transparency)
    wil::com_ptr<IWICFormatConverter> converter;
    hr = _wicFactory->CreateFormatConverter(&converter);
    if (FAILED(hr))
    {
        Debug::Warning(L"IconCache: Failed to create WIC converter for menu: 0x{:08X}", hr);
        return nullptr;
    }

    hr = converter->Initialize(scaledSource.get(), GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, nullptr, 0.0f, WICBitmapPaletteTypeCustom);
    if (FAILED(hr))
    {
        Debug::Warning(L"IconCache: Failed to initialize WIC converter for menu: 0x{:08X}", hr);
        return nullptr;
    }

    // Step 4: Create DIB section for menu bitmap
    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = size;
    bmi.bmiHeader.biHeight      = -size; // Top-down DIB
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32; // 32-bit BGRA
    bmi.bmiHeader.biCompression = BI_RGB;

    void* pBits = nullptr;
    wil::unique_hdc_window hdcScreen{GetDC(nullptr)};
    wil::unique_hbitmap hBitmap{CreateDIBSection(hdcScreen.get(), &bmi, DIB_RGB_COLORS, &pBits, nullptr, 0)};

    if (! hBitmap || ! pBits)
    {
        Debug::Warning(L"IconCache: Failed to create DIB section for menu icon");
        return nullptr;
    }

    // Step 5: Copy pixels directly from converter (NOT from original bitmap!)
    // CopyPixels reads from the converter's output, giving us proper PBGRA pixels
    UINT stride     = static_cast<UINT>(size) * 4u;
    UINT bufferSize = stride * static_cast<UINT>(size);

    WICRect rect = {0, 0, size, size};
    hr           = converter->CopyPixels(&rect, stride, bufferSize, static_cast<BYTE*>(pBits));
    if (FAILED(hr))
    {
        Debug::Warning(L"IconCache: Failed to copy pixels from converter for menu: 0x{:08X}", hr);
        return nullptr;
    }

    return hBitmap;
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

    const auto iconIndex = QuerySysIconIndexForPath(path, fileAttributes, useFileAttributes);
    if (! iconIndex.has_value())
    {
        return nullptr;
    }

    return CreateMenuBitmapFromIconIndex(iconIndex.value(), size);
}

std::optional<int> IconCache::QuerySysIconIndexForPath(const wchar_t* path, DWORD fileAttributes, bool useFileAttributes) const
{
    if (! path || *path == L'\0')
    {
        return std::nullopt;
    }

    UINT flags = SHGFI_SYSICONINDEX;
    if (useFileAttributes)
    {
        flags |= SHGFI_USEFILEATTRIBUTES;
    }

    SHFILEINFOW sfi{};
    const DWORD_PTR result = SHGetFileInfoW(path, fileAttributes, &sfi, sizeof(sfi), flags);
    if (result == 0 || sfi.iIcon < 0)
    {
        return std::nullopt;
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

wil::com_ptr<ID2D1Bitmap1> IconCache::ConvertIconToBitmapOnUIThread(HICON icon, int iconIndex, ID2D1DeviceContext* d2dContext)
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

    // Check if already cached (another thread might have added it)
    {
        std::lock_guard lock(_mutex);
        const auto deviceIt = _deviceCaches.find(device.get());
        if (deviceIt != _deviceCaches.end())
        {
            const auto it = deviceIt->second.bitmaps.find(iconIndex);
            if (it != deviceIt->second.bitmaps.end())
            {
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
        std::lock_guard lock(_mutex);
        auto& cache = _deviceCaches[device.get()];
        if (! cache.device)
        {
            cache.device = device;
        }
        EvictLRUIfNeeded(cache);
        CacheEntry entry;
        entry.bitmap             = bitmap;
        entry.lastAccessTime     = ++cache.accessCounter;
        entry.bytes              = bytes;
        cache.bitmaps[iconIndex] = std::move(entry);
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
    const size_t extCount = _extensionToIconIndex.size();

    _deviceCaches.clear();
    _extensionToIconIndex.clear();

    DBGOUT_INFO(L"IconCache: Cleared {} cached icons and {} extension mappings", iconCount, extCount);
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
        {L".txt", FILE_ATTRIBUTE_NORMAL},  {L".log", FILE_ATTRIBUTE_NORMAL},
        {L".xml", FILE_ATTRIBUTE_NORMAL},  {L".json", FILE_ATTRIBUTE_NORMAL},
        {L".ini", FILE_ATTRIBUTE_NORMAL},  {L".cfg", FILE_ATTRIBUTE_NORMAL},
        {L".md", FILE_ATTRIBUTE_NORMAL},   {L".cpp", FILE_ATTRIBUTE_NORMAL},
        {L".h", FILE_ATTRIBUTE_NORMAL},    {L".hpp", FILE_ATTRIBUTE_NORMAL},
        {L".c", FILE_ATTRIBUTE_NORMAL},    {L".cs", FILE_ATTRIBUTE_NORMAL},
        {L".py", FILE_ATTRIBUTE_NORMAL},   {L".js", FILE_ATTRIBUTE_NORMAL},
        {L".ts", FILE_ATTRIBUTE_NORMAL},   {L".html", FILE_ATTRIBUTE_NORMAL},
        {L".htm", FILE_ATTRIBUTE_NORMAL},  {L".css", FILE_ATTRIBUTE_NORMAL},
        {L".pdf", FILE_ATTRIBUTE_NORMAL},  {L".zip", FILE_ATTRIBUTE_NORMAL},
        {L".rar", FILE_ATTRIBUTE_NORMAL},  {L".7z", FILE_ATTRIBUTE_NORMAL},
        {L".png", FILE_ATTRIBUTE_NORMAL},  {L".jpg", FILE_ATTRIBUTE_NORMAL},
        {L".jpeg", FILE_ATTRIBUTE_NORMAL}, {L".gif", FILE_ATTRIBUTE_NORMAL},
        {L".bmp", FILE_ATTRIBUTE_NORMAL},  {L".ico", FILE_ATTRIBUTE_NORMAL},
        {L".svg", FILE_ATTRIBUTE_NORMAL},  {L".mp3", FILE_ATTRIBUTE_NORMAL},
        {L".wav", FILE_ATTRIBUTE_NORMAL},  {L".mp4", FILE_ATTRIBUTE_NORMAL},
        {L".avi", FILE_ATTRIBUTE_NORMAL},  {L".mkv", FILE_ATTRIBUTE_NORMAL},
        {L".doc", FILE_ATTRIBUTE_NORMAL},  {L".docx", FILE_ATTRIBUTE_NORMAL},
        {L".xls", FILE_ATTRIBUTE_NORMAL},  {L".xlsx", FILE_ATTRIBUTE_NORMAL},
        {L".ppt", FILE_ATTRIBUTE_NORMAL},  {L".pptx", FILE_ATTRIBUTE_NORMAL},
        {L".dll", FILE_ATTRIBUTE_NORMAL},  {L".sys", FILE_ATTRIBUTE_NORMAL},
        {L".bat", FILE_ATTRIBUTE_NORMAL},  {L".cmd", FILE_ATTRIBUTE_NORMAL},
        {L".ps1", FILE_ATTRIBUTE_NORMAL},  {L"<directory>", FILE_ATTRIBUTE_DIRECTORY},
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

        // Query icon index
        // For folders: Use SHGFI_USEFILEATTRIBUTES with FILE_ATTRIBUTE_DIRECTORY
        // Note: Icon indices are size-agnostic; we extract from size-specific image lists later
        const bool isFolder = (ext == L"<directory>");
        std::wstring queryPath;
        if (isFolder)
        {
            queryPath = L"C:\\DummyFolder\\";
        }
        else
        {
            queryPath = L"C:\\Dummy";
            queryPath.append(ext);
        }

        SHFILEINFOW sfi{};
        const DWORD_PTR result = SHGetFileInfoW(queryPath.c_str(), attrib, &sfi, sizeof(sfi), SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES);
        if (result != 0 && sfi.iIcon >= 0)
        {
            std::lock_guard lock(_mutex);
            if (_extensionToIconIndex.find(extKey) == _extensionToIconIndex.end())
            {
                _extensionToIconIndex[std::move(extKey)] = sfi.iIcon;
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
        std::lock_guard lock(_mutex);
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
        if (HasCachedIcon(iconIndex, device.get()))
        {
            continue;
        }

        // Extract icon and convert to D2D bitmap
        wil::unique_hicon hIcon = ExtractSystemIcon(iconIndex, 16.0f);
        if (hIcon)
        {
            auto bitmap = ConvertIconToBitmapOnUIThread(hIcon.get(), iconIndex, d2dContext);
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
    stats.lruEvictions       = _lruEvictions;
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
    const bool isFolder = (keyView == kDirectoryExtensionKey);
    std::wstring queryPath;
    if (isFolder)
    {
        queryPath = L"C:\\DummyFolder\\";
    }
    else
    {
        queryPath = L"C:\\Dummy";
        queryPath.append(keyView);
    }

    const auto iconIndex = QuerySysIconIndexForPath(queryPath.c_str(), fileAttributes, true);
    if (iconIndex.has_value())
    {
        RegisterExtension(keyView, *iconIndex);
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

    // Find oldest entry by access time
    int oldestKey     = -1;
    size_t oldestTime = SIZE_MAX;

    for (const auto& [key, entry] : cache.bitmaps)
    {
        if (entry.lastAccessTime < oldestTime)
        {
            oldestTime = entry.lastAccessTime;
            oldestKey  = key;
        }
    }

    if (oldestKey >= 0)
    {
        cache.bitmaps.erase(oldestKey);
        _lruEvictions++;
        DBGOUT_INFO(L"IconCache: Evicted icon index {} (LRU), cache size now {}", oldestKey, cache.bitmaps.size());
    }
}

wil::com_ptr<ID2D1Bitmap1> IconCache::ConvertIconToBitmap(HICON icon, ID2D1DeviceContext* d2dContext)
{
    if (! icon || ! d2dContext)
    {
        return nullptr;
    }

    if (! _wicFactory)
    {
        Debug::Warning(L"IconCache: WIC factory not initialized, cannot convert icon");
        return nullptr;
    }

    // Step 1: Create WIC bitmap from HICON (preserves alpha channel)
    wil::com_ptr<IWICBitmap> wicBitmap;
    HRESULT hr = _wicFactory->CreateBitmapFromHICON(icon, &wicBitmap);
    if (FAILED(hr))
    {
        Debug::Warning(L"IconCache: Failed to create WIC bitmap from HICON: 0x{:08X}", hr);
        return nullptr;
    }

    // Step 2: Convert to premultiplied BGRA (required by Direct2D)
    wil::com_ptr<IWICFormatConverter> converter;
    hr = _wicFactory->CreateFormatConverter(&converter);
    if (FAILED(hr))
    {
        Debug::Warning(L"IconCache: Failed to create WIC format converter: 0x{:08X}", hr);
        return nullptr;
    }

    hr = converter->Initialize(wicBitmap.get(),
                               GUID_WICPixelFormat32bppPBGRA, // Premultiplied BGRA - Direct2D native format
                               WICBitmapDitherTypeNone,
                               nullptr,
                               0.0f,
                               WICBitmapPaletteTypeCustom);
    if (FAILED(hr))
    {
        Debug::Warning(L"IconCache: Failed to initialize WIC format converter: 0x{:08X}", hr);
        return nullptr;
    }

    // Step 3: Create Direct2D bitmap from WIC bitmap
    D2D1_BITMAP_PROPERTIES1 bitmapProps{};
    bitmapProps.pixelFormat.format    = DXGI_FORMAT_B8G8R8A8_UNORM;
    bitmapProps.pixelFormat.alphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED; // Must match PBGRA
    const float dpi                   = _dpi.load(std::memory_order_relaxed);
    bitmapProps.dpiX                  = dpi;
    bitmapProps.dpiY                  = dpi;
    bitmapProps.bitmapOptions         = D2D1_BITMAP_OPTIONS_NONE;

    wil::com_ptr<ID2D1Bitmap1> d2dBitmap;
    hr = d2dContext->CreateBitmapFromWicBitmap(converter.get(), &bitmapProps, &d2dBitmap);
    if (FAILED(hr))
    {
        Debug::Warning(L"IconCache: Failed to create D2D bitmap from WIC: 0x{:08X}", hr);
        return nullptr;
    }

    return d2dBitmap;
}

int IconCache::SelectOptimalImageListSize(float targetDipSize) const
{
    // Calculate target size in physical pixels
    const float dpi          = _dpi.load(std::memory_order_relaxed);
    const float targetPixels = targetDipSize * dpi / 96.0f;

    // Select closest matching size to minimize scaling artifacts
    // Prefer slightly larger source to avoid upscaling (which looks worse than downscaling)
    if (targetPixels >= 64.0f)
    {
        return SHIL_JUMBO; // 256×256 - best for very large icons or extreme DPI scaling
    }

    if (targetPixels >= 40.0f)
    {
        return SHIL_EXTRALARGE; // 48×48 - best for high-DPI or large icons
    }

    if (targetPixels >= 24.0f)
    {
        return SHIL_LARGE; // 32×32 - good for medium sizes
    }

    return SHIL_SMALL; // 16×16 - optimal for small icons at standard DPI
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
        for (const auto& [iconIndex, cacheEntry] : entry.second.bitmaps)
        {
            static_cast<void>(iconIndex);
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
                _specialFolderIconCache.emplace(std::move(path), *iconIndex);
            }
        }
    }

    Debug::Info(L"IconCache: Initialized {} special folder paths", _specialFolderPaths.size());
}
