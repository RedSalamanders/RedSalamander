#include "NavigationViewInternal.h"

#include <windowsx.h>

#include <commctrl.h>
#include <shellapi.h>

#include "DirectoryInfoCache.h"
#include "FluentIcons.h"
#include "Helpers.h"
#include "IconCache.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "resource.h"

namespace
{
struct NavigationViewSharedDeviceResources
{
    wil::com_ptr<ID2D1Factory1> d2dFactory;
    wil::com_ptr<IDWriteFactory> dwriteFactory;
    wil::com_ptr<ID3D11Device> d3dDevice;
    wil::com_ptr<ID3D11DeviceContext> d3dContext;
    wil::com_ptr<ID2D1Device> d2dDevice;
};

NavigationViewSharedDeviceResources& GetNavigationViewSharedDeviceResources() noexcept
{
    static NavigationViewSharedDeviceResources resources;
    return resources;
}

std::mutex g_navigationViewSharedDeviceResourcesMutex;

[[nodiscard]] bool EnsureNavigationViewSharedDeviceResources(NavigationViewSharedDeviceResources& resources) noexcept
{
    HRESULT hr = S_OK;

    if (! resources.d2dFactory)
    {
        D2D1_FACTORY_OPTIONS options{};
        hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, options, &resources.d2dFactory);
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] D2D1CreateFactory failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return false;
        }
    }

    if (! resources.dwriteFactory)
    {
        hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(resources.dwriteFactory.put()));
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] DWriteCreateFactory failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return false;
        }
    }

    if (! resources.d3dDevice)
    {
        D3D_FEATURE_LEVEL featureLevels[] = {
            D3D_FEATURE_LEVEL_11_1,
            D3D_FEATURE_LEVEL_11_0,
        };

        hr = D3D11CreateDevice(nullptr,
                               D3D_DRIVER_TYPE_HARDWARE,
                               nullptr,
                               D3D11_CREATE_DEVICE_BGRA_SUPPORT,
                               featureLevels,
                               ARRAYSIZE(featureLevels),
                               D3D11_SDK_VERSION,
                               resources.d3dDevice.put(),
                               nullptr,
                               resources.d3dContext.put());
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] D3D11CreateDevice failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return false;
        }
    }

    if (! resources.d2dDevice)
    {
        wil::com_ptr<IDXGIDevice> dxgiDevice;
        hr = resources.d3dDevice->QueryInterface(IID_PPV_ARGS(&dxgiDevice));
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] QueryInterface(IDXGIDevice) failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return false;
        }

        hr = resources.d2dFactory->CreateDevice(dxgiDevice.get(), resources.d2dDevice.put());
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] ID2D1Factory1::CreateDevice failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return false;
        }
    }

    return true;
}
} // namespace

void NavigationView::WarmSharedDeviceResources() noexcept
{
    Debug::Perf::Scope perf(L"NavigationView.WarmSharedDeviceResources");

    std::lock_guard lock(g_navigationViewSharedDeviceResourcesMutex);
    NavigationViewSharedDeviceResources& shared = GetNavigationViewSharedDeviceResources();
    perf.SetHr(EnsureNavigationViewSharedDeviceResources(shared) ? S_OK : E_FAIL);
}

void NavigationView::EnsureD2DResources()
{
    HRESULT hr = S_OK;

    {
        std::lock_guard lock(g_navigationViewSharedDeviceResourcesMutex);
        NavigationViewSharedDeviceResources& shared = GetNavigationViewSharedDeviceResources();
        if (! EnsureNavigationViewSharedDeviceResources(shared))
        {
            return;
        }

        _d2dFactory    = shared.d2dFactory;
        _dwriteFactory = shared.dwriteFactory;
        _d3dDevice     = shared.d3dDevice;
        _d3dContext    = shared.d3dContext;
        _d2dDevice     = shared.d2dDevice;
    }

    // Create D2D factory
    if (! _d2dFactory || ! _dwriteFactory || ! _d3dDevice || ! _d2dDevice)
    {
        Debug::Error(L"[NavigationView] EnsureD2DResources failed: shared device resources are null");
        return;
    }

    // Calculate bar height in physical pixels (fixed 24 DIP at 96 DPI)
    float barHeight      = static_cast<float>(MulDiv(kHeight, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI));
    float breadcrumbSize = barHeight * 0.6f; // ~14.4pt at 96 DPI
    float separatorSize  = static_cast<float>(MulDiv(FluentIcons::kDefaultSizeDip, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI));

    if (! _pathFormat)
    {
        hr = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, breadcrumbSize, L"", &_pathFormat);
        if (SUCCEEDED(hr))
        {
            _pathFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
            _pathFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
        }
    }

    if (! _separatorFormat)
    {
        hr = _dwriteFactory->CreateTextFormat(FluentIcons::kFontFamily.data(),
                                              nullptr,
                                              DWRITE_FONT_WEIGHT_NORMAL,
                                              DWRITE_FONT_STYLE_NORMAL,
                                              DWRITE_FONT_STRETCH_NORMAL,
                                              separatorSize,
                                              L"",
                                              &_separatorFormat);
        if (SUCCEEDED(hr))
        {
            _dwriteFluentIconsValid   = true;
            _breadcrumbSeparatorGlyph = FluentIcons::kChevronRightSmall;
            _historyChevronGlyph      = FluentIcons::kChevronDown;
        }
        else
        {
            hr = _dwriteFactory->CreateTextFormat(
                L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, separatorSize, L"", &_separatorFormat);
            if (SUCCEEDED(hr))
            {
                _dwriteFluentIconsValid   = false;
                _breadcrumbSeparatorGlyph = FluentIcons::kFallbackChevronRight;
                _historyChevronGlyph      = FluentIcons::kFallbackChevronDown;
            }
        }

        if (SUCCEEDED(hr) && _separatorFormat)
        {
            _separatorFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
            _separatorFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
        }
    }

    if (_pathFormat)
    {
        static_cast<void>(_pathFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
    }

    if (_separatorFormat)
    {
        static_cast<void>(_separatorFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
    }

    // Create D2D device context
    if (! _d2dContext)
    {
        hr = _d2dDevice->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE, &_d2dContext);
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] ID2D1Device::CreateDeviceContext failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return;
        }
    }

    // Create swap chain for full window
    if (! _swapChain && _hWnd)
    {
        UINT winWidth  = static_cast<UINT>(_clientSize.cx);
        UINT winHeight = static_cast<UINT>(_clientSize.cy);

        if (winWidth == 0 || winHeight == 0)
            return;

        wil::com_ptr<IDXGIDevice> dxgiDevice;
        hr = _d3dDevice->QueryInterface(IID_PPV_ARGS(&dxgiDevice));
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] QueryInterface(IDXGIDevice) for swap chain failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return;
        }

        wil::com_ptr<IDXGIAdapter> dxgiAdapter;
        hr = dxgiDevice->GetAdapter(&dxgiAdapter);
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] IDXGIDevice::GetAdapter failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return;
        }

        wil::com_ptr<IDXGIFactory2> dxgiFactory;
        hr = dxgiAdapter->GetParent(IID_PPV_ARGS(&dxgiFactory));
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] IDXGIAdapter::GetParent(IDXGIFactory2) failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return;
        }

        DXGI_SWAP_CHAIN_DESC1 desc{};
        desc.Width            = winWidth;
        desc.Height           = winHeight;
        desc.Format           = DXGI_FORMAT_B8G8R8A8_UNORM;
        desc.SampleDesc.Count = 1;
        desc.BufferUsage      = DXGI_USAGE_RENDER_TARGET_OUTPUT;
        desc.BufferCount      = 2;
        desc.SwapEffect       = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
        desc.AlphaMode        = DXGI_ALPHA_MODE_IGNORE;

        hr = dxgiFactory->CreateSwapChainForHwnd(_d3dDevice.get(), _hWnd.get(), &desc, nullptr, nullptr, &_swapChain);
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] CreateSwapChainForHwnd failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return;
        }
        else
        {
            _hasPresented = false; // Reset flag when swap chain is created
        }

        // Disable Alt+Enter
        hr = dxgiFactory->MakeWindowAssociation(_hWnd.get(), DXGI_MWA_NO_ALT_ENTER);
        if (FAILED(hr))
        {
            Debug::Warning(L"[NavigationView] MakeWindowAssociation failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
        }
    }

    // Create render target from swap chain
    if (! _d2dTarget)
    {
        if (! _swapChain)
        {
            Debug::Error(L"[NavigationView] Cannot create D2D target: swap chain is null");
            return;
        }

        wil::com_ptr<IDXGISurface> surface;
        hr = _swapChain->GetBuffer(0, IID_PPV_ARGS(&surface));
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] IDXGISwapChain::GetBuffer failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            return;
        }
        DXGI_SURFACE_DESC surfaceDesc = {};
        surface->GetDesc(&surfaceDesc);
        if (SUCCEEDED(hr))
        {
            D2D1_BITMAP_PROPERTIES1 props = D2D1::BitmapProperties1(D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW,
                                                                    D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED),
                                                                    static_cast<float>(USER_DEFAULT_SCREEN_DPI),
                                                                    static_cast<float>(USER_DEFAULT_SCREEN_DPI));

            // create bitmap for DirectX
            hr = _d2dContext->CreateBitmapFromDxgiSurface(surface.get(), &props, &_d2dTarget);
            if (FAILED(hr))
            {
                Debug::Error(L"[NavigationView] CreateBitmapFromDxgiSurface failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
            }
        }
    }

    // Create/update brushes from theme colors
    if (_d2dContext)
    {
        if (! _textBrush)
        {
            _d2dContext->CreateSolidColorBrush(_theme.text, &_textBrush);
        }
        else
        {
            _textBrush->SetColor(_theme.text);
        }

        if (! _separatorBrush)
        {
            _d2dContext->CreateSolidColorBrush(_theme.separator, &_separatorBrush);
        }
        else
        {
            _separatorBrush->SetColor(_theme.separator);
        }

        if (! _hoverBrush)
        {
            _d2dContext->CreateSolidColorBrush(_theme.hoverHighlight, &_hoverBrush);
        }
        else
        {
            _hoverBrush->SetColor(_theme.hoverHighlight);
        }

        if (! _pressedBrush)
        {
            _d2dContext->CreateSolidColorBrush(_theme.pressedHighlight, &_pressedBrush);
        }
        else
        {
            _pressedBrush->SetColor(_theme.pressedHighlight);
        }

        if (! _accentBrush)
        {
            _d2dContext->CreateSolidColorBrush(_theme.accent, &_accentBrush);
        }
        else
        {
            _accentBrush->SetColor(_theme.accent);
        }

        if (! _rainbowBrush)
        {
            _d2dContext->CreateSolidColorBrush(_theme.accent, &_rainbowBrush);
        }
        else
        {
            _rainbowBrush->SetColor(_theme.accent);
        }

        if (! _backgroundBrushD2D)
        {
            _d2dContext->CreateSolidColorBrush(_theme.background, &_backgroundBrushD2D);
        }
        else
        {
            _backgroundBrushD2D->SetColor(_theme.background);
        }
    }

    // Update menu icon bitmap when D2D resources are ready
    if (_d2dContext && _currentPluginPath && ! _menuIconBitmapD2D)
    {
        UpdateMenuIconBitmap();
    }
}

void NavigationView::DiscardD2DResources()
{
    InvalidateBreadcrumbLayoutCache();

    wil::com_ptr<ID2D1Device> oldD2DDevice = _d2dDevice;

    _d2dTarget          = nullptr;
    _swapChain          = nullptr;
    _hasPresented       = false; // Reset flag when discarding swap chain
    _textBrush          = nullptr;
    _separatorBrush     = nullptr;
    _hoverBrush         = nullptr;
    _pressedBrush       = nullptr;
    _accentBrush        = nullptr;
    _rainbowBrush       = nullptr;
    _backgroundBrushD2D = nullptr;
    _menuIconBitmapD2D  = nullptr; // Clear menu icon bitmap
    _d2dContext         = nullptr;
    _d2dDevice          = nullptr;
    _d3dContext         = nullptr;
    _d3dDevice          = nullptr;
    _pathFormat         = nullptr;
    _separatorFormat    = nullptr;
    _dwriteFactory      = nullptr;
    _d2dFactory         = nullptr;

    IconCache::GetInstance().ClearDeviceCache(oldD2DDevice.get());
}

void NavigationView::UpdateMenuIconBitmap()
{
    // Clear existing bitmap
    _menuIconBitmapD2D = nullptr;

    if (! _showMenuSection || ! _currentPluginPath || ! _d2dContext)
    {
        return;
    }

    // Extract icon index - coherent with FolderView/IconCache:
    // 1. For paths under special folders (e.g., Documents\foo\bar), use the special folder icon index.
    // 2. For drive paths (e.g., C:\foo\bar), use the drive root icon index (C:\).
    // 3. Otherwise, use the current folder icon index.

    int iconIndex = -1;

    const bool isFilePlugin = _pluginShortId.empty() || EqualsNoCase(_pluginShortId, L"file");
    if (! isFilePlugin)
    {
        if (! _currentInstanceContext.empty() && LooksLikeWindowsAbsolutePath(_currentInstanceContext))
        {
            const auto index = IconCache::GetInstance().QuerySysIconIndexForPath(_currentInstanceContext.c_str(), 0, false);
            iconIndex        = index.value_or(-1);
        }

        if (iconIndex < 0 && EqualsNoCase(_pluginShortId, L"fk"))
        {
            SHSTOCKICONINFO sii{};
            sii.cbSize       = sizeof(sii);
            const HRESULT hr = SHGetStockIconInfo(SIID_DRIVENET, SHGSI_SYSICONINDEX, &sii);
            if (SUCCEEDED(hr))
            {
                iconIndex = sii.iSysImageIndex;
            }
        }

        if (iconIndex < 0 && _fileSystemPlugin)
        {
            void** const vtbl = *reinterpret_cast<void***>(_fileSystemPlugin.get());
            if (vtbl != nullptr && vtbl[0] != nullptr)
            {
                HMODULE module = nullptr;
                if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                       reinterpret_cast<LPCWSTR>(vtbl[0]),
                                       &module) != 0 &&
                    module != nullptr)
                {
                    std::wstring modulePath;
                    modulePath.resize(MAX_PATH);

                    for (;;)
                    {
                        const DWORD length = GetModuleFileNameW(module, modulePath.data(), static_cast<DWORD>(modulePath.size()));
                        if (length == 0)
                        {
                            modulePath.clear();
                            break;
                        }

                        if (length < modulePath.size() - 1u)
                        {
                            modulePath.resize(static_cast<size_t>(length));
                            break;
                        }

                        if (modulePath.size() >= 32768u)
                        {
                            modulePath.resize(static_cast<size_t>(length));
                            break;
                        }

                        modulePath.resize(modulePath.size() * 2u);
                    }

                    if (! modulePath.empty())
                    {
                        const auto index = IconCache::GetInstance().QuerySysIconIndexForPath(modulePath.c_str(), 0, false);
                        iconIndex        = index.value_or(-1);
                    }
                }
            }
        }

        if (iconIndex < 0)
        {
            const auto folderIndex = IconCache::GetInstance().GetOrQueryIconIndexByExtension(L"<directory>", FILE_ATTRIBUTE_DIRECTORY);
            iconIndex              = folderIndex.value_or(-1);
        }
    }
    else
    {
        const auto& currentPath              = _currentPluginPath.value();
        const std::wstring currentPathString = currentPath.wstring();

        if (currentPath.has_root_path())
        {
            const auto special = IconCache::GetInstance().TryGetSpecialFolderForPathPrefix(currentPathString);
            if (special.has_value())
            {
                iconIndex = special.value().iconIndex;
            }

            if (iconIndex < 0)
            {
                const std::wstring rootPath = currentPath.root_path().wstring();
                const auto rootIndex        = IconCache::GetInstance().QuerySysIconIndexForPath(rootPath.c_str(), 0, false);
                iconIndex                   = rootIndex.value_or(-1);
            }
        }

        if (iconIndex < 0)
        {
            const auto pathIndex = IconCache::GetInstance().QuerySysIconIndexForPath(currentPathString.c_str(), 0, false);
            iconIndex            = pathIndex.value_or(-1);
        }
    }

    if (iconIndex >= 0)
    {
        _menuIconBitmapD2D = IconCache::GetInstance().GetIconBitmap(iconIndex, _d2dContext.get());
    }
}

void NavigationView::RenderDriveSection()
{
    if (! _showMenuSection)
    {
        return;
    }
    // Ensure D2D resources are initialized before rendering
    EnsureD2DResources();

    if (! _d2dContext || ! _swapChain)
    {
        return;
    }

    // Render Section Drive Menu (shows hamburger fallback)

    _d2dContext->BeginDraw();
    auto endDraw = wil::scope_exit(
        [&]
        {
            const HRESULT hrEnd = _d2dContext->EndDraw();
            if (FAILED(hrEnd))
            {
                if (hrEnd == D2DERR_RECREATE_TARGET)
                {
                    DiscardD2DResources();
                    return;
                }

                Debug::Error(L"[NavigationView] EndDraw failed (hr=0x{:08X})", static_cast<unsigned long>(hrEnd));
                return;
            }

            RECT dirtyRect = _sectionDriveRect;
            Present(&dirtyRect);
        });

    _d2dContext->SetTarget(_d2dTarget.get());

    // Determine background color based on button state
    D2D1_COLOR_F bgColor;
    if (_menuButtonPressed)
    {
        bgColor = _theme.backgroundPressed;
    }
    else if (_menuButtonHovered)
    {
        bgColor = _theme.backgroundHover;
    }
    else
    {
        bgColor = _theme.background;
    }

    // Create temporary brush for background
    wil::com_ptr<ID2D1SolidColorBrush> bgBrush;
    _d2dContext->CreateSolidColorBrush(bgColor, &bgBrush);

    // Convert Section 1 rect to D2D rect
    D2D1_RECT_F section1RectF = D2D1::RectF(static_cast<float>(_sectionDriveRect.left),
                                            static_cast<float>(_sectionDriveRect.top),
                                            static_cast<float>(_sectionDriveRect.right),
                                            static_cast<float>(_sectionDriveRect.bottom));

    // Fill background
    _d2dContext->FillRectangle(section1RectF, bgBrush.get());

    // Draw icon or hamburger fallback
    if (_menuIconBitmapD2D)
    {
        // Get bitmap size
        D2D1_SIZE_F bitmapSize = _menuIconBitmapD2D->GetSize();

        // Center the icon in Section 1
        float centerX = (static_cast<float>(_sectionDriveRect.left) + static_cast<float>(_sectionDriveRect.right)) / 2.0f;
        float centerY = (static_cast<float>(_sectionDriveRect.top) + static_cast<float>(_sectionDriveRect.bottom)) / 2.0f;
        float left    = centerX - bitmapSize.width / 2.0f;
        float top     = centerY - bitmapSize.height / 2.0f;

        D2D1_RECT_F destRect = D2D1::RectF(left, top, left + bitmapSize.width, top + bitmapSize.height);

        const float opacity = _paneFocused ? 1.0f : 0.55f;
        _d2dContext->DrawBitmap(_menuIconBitmapD2D.get(), destRect, opacity);
    }
    else
    {
        // Fallback: Draw hamburger icon (3 horizontal lines)
        ID2D1SolidColorBrush* lineBrush = _textBrush.get();
        wil::com_ptr<ID2D1SolidColorBrush> fallbackBrush;
        if (! lineBrush && _d2dContext)
        {
            _d2dContext->CreateSolidColorBrush(_theme.text, fallbackBrush.addressof());
            lineBrush = fallbackBrush.get();
        }

        float centerX   = (static_cast<float>(_sectionDriveRect.left) + static_cast<float>(_sectionDriveRect.right)) / 2.0f;
        float centerY   = (static_cast<float>(_sectionDriveRect.top) + static_cast<float>(_sectionDriveRect.bottom)) / 2.0f;
        float lineWidth = 2.0f;

        for (int i = -1; i <= 1; i++)
        {
            float y = centerY + static_cast<float>(i * 5);
            if (lineBrush)
            {
                _d2dContext->DrawLine(D2D1::Point2F(centerX - 6.0f, y), D2D1::Point2F(centerX + 7.0f, y), lineBrush, lineWidth);
            }
        }
    }

    if (! _editMode && _accentBrush && _hWnd && GetFocus() == _hWnd.get() && _focusedRegion == FocusRegion::Menu)
    {
        constexpr float inset = 1.0f;
        const D2D1_RECT_F focusRect =
            D2D1::RectF(section1RectF.left + inset, section1RectF.top + inset, section1RectF.right - inset, section1RectF.bottom - inset);
        const float cornerRadius        = DipsToPixels(kFocusRingCornerRadiusDip, _dpi);
        const D2D1_ROUNDED_RECT rounded = RoundedRect(focusRect, cornerRadius);
        _d2dContext->DrawRoundedRectangle(rounded, _accentBrush.get(), 2.0f);
    }
}

void NavigationView::RenderHistorySection()
{
    // Render: History button with Direct2D

    // Ensure D2D resources are initialized
    EnsureD2DResources();

    if (! _d2dContext || ! _d2dTarget)
    {
        return;
    }

    _d2dContext->BeginDraw();
    auto endDraw = wil::scope_exit(
        [&]
        {
            const HRESULT hrEnd = _d2dContext->EndDraw();
            if (FAILED(hrEnd))
            {
                if (hrEnd == D2DERR_RECREATE_TARGET)
                {
                    DiscardD2DResources();
                    return;
                }

                Debug::Error(L"[NavigationView] EndDraw failed (hr=0x{:08X})", static_cast<unsigned long>(hrEnd));
                return;
            }

            RECT dirtyRect = _sectionHistoryRect;
            Present(&dirtyRect);
        });

    _d2dContext->SetTarget(_d2dTarget.get());

    ID2D1SolidColorBrush* textBrush = _textBrush.get();

    // Render History Button
    D2D1_RECT_F historyRect = D2D1::RectF(static_cast<float>(_sectionHistoryRect.left),
                                          static_cast<float>(_sectionHistoryRect.top),
                                          static_cast<float>(_sectionHistoryRect.right),
                                          static_cast<float>(_sectionHistoryRect.bottom));

    const D2D1_COLOR_F historyBgColor = _historyButtonHovered ? _theme.backgroundHover : _theme.background;

    wil::com_ptr<ID2D1SolidColorBrush> historyBgBrush;
    _d2dContext->CreateSolidColorBrush(historyBgColor, &historyBgBrush);
    _d2dContext->FillRectangle(historyRect, historyBgBrush.get());

    // Draw down chevron
    if (_separatorFormat)
    {
        wil::com_ptr<IDWriteTextLayout> historyLayout;
        _dwriteFactory->CreateTextLayout(
            &_historyChevronGlyph, 1, _separatorFormat.get(), historyRect.right - historyRect.left, historyRect.bottom - historyRect.top, &historyLayout);
        // L"▿^▿∨∨⊽⨈⩔⩡▼"

        if (historyLayout)
        {
            historyLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
            historyLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

            D2D1_POINT_2F origin = D2D1::Point2F(historyRect.left, historyRect.top);
            if (textBrush)
            {
                _d2dContext->DrawTextLayout(origin, historyLayout.get(), textBrush);
            }
        }
    }

    if (! _editMode && _accentBrush && _hWnd && GetFocus() == _hWnd.get() && _focusedRegion == FocusRegion::History)
    {
        constexpr float inset           = 1.0f;
        const D2D1_RECT_F focusRect     = D2D1::RectF(historyRect.left + inset, historyRect.top + inset, historyRect.right - inset, historyRect.bottom - inset);
        const float cornerRadius        = DipsToPixels(kFocusRingCornerRadiusDip, _dpi);
        const D2D1_ROUNDED_RECT rounded = RoundedRect(focusRect, cornerRadius);
        _d2dContext->DrawRoundedRectangle(rounded, _accentBrush.get(), 2.0f);
    }
}

void NavigationView::RenderDiskInfoSection()
{
    if (! _showDiskInfoSection)
    {
        return;
    }
    // Render: disk info with Direct2D

    // Ensure D2D resources are initialized
    EnsureD2DResources();

    if (! _d2dContext || ! _d2dTarget)
    {
        return;
    }

    _d2dContext->BeginDraw();
    auto endDraw = wil::scope_exit(
        [&]
        {
            const HRESULT hrEnd = _d2dContext->EndDraw();
            if (FAILED(hrEnd))
            {
                if (hrEnd == D2DERR_RECREATE_TARGET)
                {
                    DiscardD2DResources();
                    return;
                }

                Debug::Error(L"[NavigationView] EndDraw failed (hr=0x{:08X})", static_cast<unsigned long>(hrEnd));
                return;
            }

            RECT dirtyRect = _sectionDiskInfoRect;
            Present(&dirtyRect);
        });
    _d2dContext->SetTarget(_d2dTarget.get());

    ID2D1SolidColorBrush* textBrush = _textBrush.get();
    wil::com_ptr<ID2D1SolidColorBrush> progressBrush;
    wil::com_ptr<ID2D1SolidColorBrush> barBgBrush;

    // Draw down chevron
    if (_separatorFormat)
    {
        wil::com_ptr<IDWriteTextLayout> historyLayout;
        _dwriteFactory->CreateTextLayout(&_historyChevronGlyph,
                                         1,
                                         _separatorFormat.get(),
                                         static_cast<float>(_sectionHistoryRect.right - _sectionHistoryRect.left),
                                         static_cast<float>(_sectionHistoryRect.bottom - _sectionHistoryRect.top),
                                         &historyLayout);

        if (historyLayout)
        {
            historyLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
            historyLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

            D2D1_POINT_2F origin = D2D1::Point2F(static_cast<float>(_sectionHistoryRect.left), static_cast<float>(_sectionHistoryRect.top));
            if (textBrush)
            {
                _d2dContext->DrawTextLayout(origin, historyLayout.get(), textBrush);
            }
        }
    }

    // Render Section 3: Disk Info
    D2D1_RECT_F section4Rect = D2D1::RectF(static_cast<float>(_sectionDiskInfoRect.left),
                                           static_cast<float>(_sectionDiskInfoRect.top),
                                           static_cast<float>(_sectionDiskInfoRect.right),
                                           static_cast<float>(_sectionDiskInfoRect.bottom));

    const D2D1_COLOR_F section4BgColor = _diskInfoHovered ? _theme.backgroundHover : _theme.background;

    wil::com_ptr<ID2D1SolidColorBrush> section4BgBrush;
    _d2dContext->CreateSolidColorBrush(section4BgColor, &section4BgBrush);
    _d2dContext->FillRectangle(section4Rect, section4BgBrush.get());

    // Draw disk space text (right-aligned) - only if we have a path and disk info
    if (_currentPluginPath && ! _diskSpaceText.empty())
    {
        wil::com_ptr<IDWriteTextLayout> diskTextLayout;
        _dwriteFactory->CreateTextLayout(_diskSpaceText.c_str(),
                                         static_cast<UINT32>(_diskSpaceText.length()),
                                         _pathFormat.get(),
                                         section4Rect.right - section4Rect.left - 8.0f, // Padding
                                         section4Rect.bottom - section4Rect.top - 6.0f, // Leave space for progress bar
                                         &diskTextLayout);

        if (diskTextLayout)
        {
            diskTextLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);
            diskTextLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

            D2D1_POINT_2F origin = D2D1::Point2F(section4Rect.left + 4.0f, section4Rect.top);
            if (textBrush)
            {
                _d2dContext->DrawTextLayout(origin, diskTextLayout.get(), textBrush);
            }
        }
    }

    // Draw progress bar at bottom - only if we have a path
    if (_currentPluginPath)
    {
        D2D1_RECT_F progressRect = D2D1::RectF(section4Rect.left + 4.0f, section4Rect.bottom - 3.0f, section4Rect.right - 4.0f, section4Rect.bottom);

        const bool hasUsageInfo = _hasUsedBytes || _hasFreeBytes;
        if (_totalBytes > 0 && hasUsageInfo)
        {
            // Calculate used space percentage
            uint64_t usedBytes = 0;
            if (_hasUsedBytes)
            {
                usedBytes = _usedBytes;
            }
            else if (_totalBytes >= _freeBytes)
            {
                usedBytes = _totalBytes - _freeBytes;
            }

            double usedPercent = static_cast<double>(usedBytes) / static_cast<double>(_totalBytes);
            if (usedPercent < 0.0)
            {
                usedPercent = 0.0;
            }
            if (usedPercent > 1.0)
            {
                usedPercent = 1.0;
            }

            // Progress color (theme-defined: ok if < 90% used, warn if >= 90% used)
            D2D1_COLOR_F progressColor;
            if (usedPercent < 0.9)
            {
                progressColor = _theme.progressOk;
            }
            else
            {
                progressColor = _theme.progressWarn;
            }

            _d2dContext->CreateSolidColorBrush(progressColor, &progressBrush);

            // Draw used portion
            D2D1_RECT_F fillRect = D2D1::RectF(progressRect.left,
                                               progressRect.top,
                                               progressRect.left + static_cast<float>((progressRect.right - progressRect.left) * usedPercent),
                                               progressRect.bottom);
            _d2dContext->FillRectangle(fillRect, progressBrush.get());

            // Draw free portion
            if (fillRect.right < progressRect.right)
            {
                D2D1_RECT_F bgRect = D2D1::RectF(fillRect.right, progressRect.top, progressRect.right, progressRect.bottom);
                _d2dContext->CreateSolidColorBrush(_theme.progressBackground, &barBgBrush);
                _d2dContext->FillRectangle(bgRect, barBgBrush.get());
            }
        }
        else
        {
            // No disk info available - show grey bar
            _d2dContext->CreateSolidColorBrush(_theme.progressBackground, &progressBrush);
            _d2dContext->FillRectangle(progressRect, progressBrush.get());
        }
    } // End if (_currentPluginPath) for progress bar

    if (! _editMode && _accentBrush && _hWnd && GetFocus() == _hWnd.get() && _focusedRegion == FocusRegion::DiskInfo)
    {
        constexpr float inset       = 1.0f;
        const D2D1_RECT_F focusRect = D2D1::RectF(section4Rect.left + inset, section4Rect.top + inset, section4Rect.right - inset, section4Rect.bottom - inset);
        const float cornerRadius    = DipsToPixels(kFocusRingCornerRadiusDip, _dpi);
        const D2D1_ROUNDED_RECT rounded = RoundedRect(focusRect, cornerRadius);
        _d2dContext->DrawRoundedRectangle(rounded, _accentBrush.get(), 2.0f);
    }
}

void NavigationView::Present(std::optional<RECT*> dirtyRect)
{
    if (! _swapChain)
    {
        Debug::Error(L"[NavigationView] Cannot present: swap chain is null");
        return;
    }

    if (_deferPresent)
    {
        if (! dirtyRect)
        {
            _queuedPresentFull = true;
            _queuedPresentDirtyRect.reset();
            return;
        }

        if (_queuedPresentFull)
        {
            return;
        }

        const RECT rect = *dirtyRect.value();
        if (_queuedPresentDirtyRect)
        {
            RECT merged{};
            RECT existing = *_queuedPresentDirtyRect;
            UnionRect(&merged, &existing, &rect);
            _queuedPresentDirtyRect = merged;
        }
        else
        {
            _queuedPresentDirtyRect = rect;
        }
        return;
    }

    HRESULT hr = S_OK;

    if (! _hasPresented)
    {
        // First present must be a regular Present without dirty rects
        hr            = _swapChain->Present(0, 0);
        _hasPresented = true;
        if (FAILED(hr))
        {
            Debug::Error(L"[NavigationView] Present failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
        }
        if (hr == D2DERR_RECREATE_TARGET)
        {
            Debug::Info(L"[NavigationView] Recreating D2D resources");
            DiscardD2DResources();
        }
        return;
    }

    // Subsequent presents can use Present1 with dirty rects
    DXGI_PRESENT_PARAMETERS params{};
    if (dirtyRect)
    {
        params.DirtyRectsCount = 1;
        params.pDirtyRects     = dirtyRect.value();
    }
    hr = _swapChain->Present1(0, 0, &params);
    if (FAILED(hr))
    {
        Debug::Error(L"[NavigationView] Present1 failed (hr=0x{:08X})", static_cast<unsigned long>(hr));
    }
    if (hr == D2DERR_RECREATE_TARGET)
    {
        Debug::Info(L"[NavigationView] Recreating D2D resources");
        DiscardD2DResources();
    }
}

void NavigationView::UpdateDiskInfo()
{
    _diskSpaceText.clear();
    _freeBytes     = 0;
    _totalBytes    = 0;
    _usedBytes     = 0;
    _hasTotalBytes = false;
    _hasFreeBytes  = false;
    _hasUsedBytes  = false;
    _volumeLabel.clear();
    _fileSystem.clear();
    _driveDisplayName.clear();

    if (! _currentPluginPath || ! _driveInfo)
    {
        return;
    }

    const std::wstring pathText = _currentPluginPath.value().wstring();
    DriveInfo info{};
    const HRESULT hr = _driveInfo->GetDriveInfo(pathText.c_str(), &info);
    if (FAILED(hr) || hr == S_FALSE)
    {
        return;
    }

    if ((info.flags & DRIVE_INFO_FLAG_HAS_DISPLAY_NAME) != 0 && info.displayName)
    {
        _driveDisplayName = info.displayName;
    }
    else
    {
        const bool isFilePlugin = _pluginShortId.empty() || EqualsNoCase(_pluginShortId, L"file");
        if (isFilePlugin)
        {
            const std::filesystem::path root = _currentPluginPath.value().root_path();
            _driveDisplayName                = root.empty() ? _currentPluginPath.value().wstring() : root.wstring();
        }
        else
        {
            _driveDisplayName = L"/";
        }
    }

    if ((info.flags & DRIVE_INFO_FLAG_HAS_VOLUME_LABEL) != 0 && info.volumeLabel)
    {
        _volumeLabel = info.volumeLabel;
    }

    if ((info.flags & DRIVE_INFO_FLAG_HAS_FILE_SYSTEM) != 0 && info.fileSystem)
    {
        _fileSystem = info.fileSystem;
    }

    if ((info.flags & DRIVE_INFO_FLAG_HAS_TOTAL_BYTES) != 0)
    {
        _totalBytes    = info.totalBytes;
        _hasTotalBytes = true;
    }

    if ((info.flags & DRIVE_INFO_FLAG_HAS_FREE_BYTES) != 0)
    {
        _freeBytes    = info.freeBytes;
        _hasFreeBytes = true;
    }

    if ((info.flags & DRIVE_INFO_FLAG_HAS_USED_BYTES) != 0)
    {
        _usedBytes    = info.usedBytes;
        _hasUsedBytes = true;
    }

    if (_hasTotalBytes)
    {
        if (_hasFreeBytes)
        {
            _diskSpaceText = FormatBytesCompact(_freeBytes) + L" ";
        }
        else
        {
            _diskSpaceText = FormatBytesCompact(_totalBytes) + L" ";
        }
    }
}
