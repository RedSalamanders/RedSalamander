#include "FolderViewInternal.h"

#include <cmath>

void FolderView::EnsureDeviceIndependentResources()
{
    if (! _wicFactory)
    {
        const HRESULT hrFactory = CoCreateInstance(CLSID_WICImagingFactory2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(_wicFactory.addressof()));
        if (! CheckHR(hrFactory, L"CoCreateInstance(CLSID_WICImagingFactory)"))
        {
            return;
        }
    }
    if (! _dwriteFactory)
    {
        const HRESULT hrDWrite =
            DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_dwriteFactory.addressof()));
        if (! CheckHR(hrDWrite, L"DWriteCreateFactory"))
        {
            return;
        }
    }
    if (! _labelFormat)
    {
        const HRESULT hrFormat = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 12.0f, L"en-us", _labelFormat.addressof());
        if (! CheckHR(hrFormat, L"IDWriteFactory::CreateTextFormat"))
        {
            return;
        }
        _labelFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
        _labelFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _labelFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
        DWRITE_TRIMMING trimming{};
        trimming.granularity = DWRITE_TRIMMING_GRANULARITY_CHARACTER;
        wil::com_ptr<IDWriteInlineObject> ellipsis;
        if (SUCCEEDED(_dwriteFactory->CreateEllipsisTrimmingSign(_labelFormat.get(), ellipsis.addressof())))
        {
            _labelFormat->SetTrimming(&trimming, ellipsis.get());
            _ellipsisSign = std::move(ellipsis);
        }
        else
        {
            _labelFormat->SetTrimming(&trimming, nullptr);
            _ellipsisSign.reset();
        }
    }

    if (! _detailsFormat)
    {
        const HRESULT hrFormat = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 10.0f, L"en-us", _detailsFormat.addressof());
        if (! CheckHR(hrFormat, L"IDWriteFactory::CreateTextFormat(details)"))
        {
            return;
        }

        _detailsFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
        _detailsFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _detailsFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);

        DWRITE_TRIMMING trimming{};
        trimming.granularity = DWRITE_TRIMMING_GRANULARITY_NONE;
        _detailsFormat->SetTrimming(&trimming, nullptr);
        _detailsEllipsisSign.reset();

        wil::com_ptr<IDWriteTextLayout> probe;
        constexpr wchar_t probeText[] = L"Ag";
        const HRESULT hrProbe         = _dwriteFactory->CreateTextLayout(
            probeText, static_cast<UINT32>(std::size(probeText) - 1), _detailsFormat.get(), 1000.0f, 1000.0f, probe.addressof());
        if (SUCCEEDED(hrProbe) && probe)
        {
            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(probe->GetMetrics(&metrics)))
            {
                _detailsLineHeightDip  = metrics.height;
                _metadataLineHeightDip = metrics.height;
            }
        }

        if (_detailsLineHeightDip <= 0.0f)
        {
            _detailsLineHeightDip = 12.0f;
        }
        if (_metadataLineHeightDip <= 0.0f)
        {
            _metadataLineHeightDip = _detailsLineHeightDip;
        }
    }

    // Alert overlay formats are owned by the shared RedSalamander::Ui::AlertOverlay component.
}

void FolderView::EnsureDeviceResources()
{
    if (_d3dDevice && _d2dContext && _d2dFactory)
    {
        return;
    }

    UINT creationFlags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
#if defined(_DEBUG)
    creationFlags |= D3D11_CREATE_DEVICE_DEBUG;
#endif

    D3D_FEATURE_LEVEL levels[] = {D3D_FEATURE_LEVEL_11_1, D3D_FEATURE_LEVEL_11_0, D3D_FEATURE_LEVEL_10_1};

    HRESULT hrDevice = D3D11CreateDevice(nullptr,
                                         D3D_DRIVER_TYPE_HARDWARE,
                                         nullptr,
                                         creationFlags,
                                         levels,
                                         std::size(levels),
                                         D3D11_SDK_VERSION,
                                         _d3dDevice.addressof(),
                                         &_featureLevel,
                                         _d3dContext.addressof());
    if (FAILED(hrDevice))
    {
        hrDevice = D3D11CreateDevice(nullptr,
                                     D3D_DRIVER_TYPE_WARP,
                                     nullptr,
                                     creationFlags,
                                     levels,
                                     std::size(levels),
                                     D3D11_SDK_VERSION,
                                     _d3dDevice.addressof(),
                                     &_featureLevel,
                                     _d3dContext.addressof());
        if (! CheckHR(hrDevice, L"D3D11CreateDevice (WARP)"))
        {
            return;
        }
    }
    else if (! CheckHR(hrDevice, L"D3D11CreateDevice"))
    {
        return;
    }
    else
    {
        Debug::Info(L"FolderView: D3D device created with feature level {:#06x}", static_cast<int>(_featureLevel));
    }

    wil::com_ptr<IDXGIDevice> dxgiDevice;
    const HRESULT hrDxgiDevice = _d3dDevice ? _d3dDevice->QueryInterface(IID_PPV_ARGS(dxgiDevice.addressof())) : E_POINTER;
    if (! CheckHR(hrDxgiDevice, L"ID3D11Device::QueryInterface IDXGIDevice"))
    {
        return;
    }

    D2D1_FACTORY_OPTIONS options{};
#if defined(_DEBUG)
    options.debugLevel = D2D1_DEBUG_LEVEL_INFORMATION;
#endif
    const HRESULT hrD2DFactory =
        D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, __uuidof(ID2D1Factory1), &options, reinterpret_cast<void**>(_d2dFactory.addressof()));
    if (! CheckHR(hrD2DFactory, L"D2D1CreateFactory"))
    {
        return;
    }
    wil::com_ptr<ID2D1Device> createdD2DDevice;
    const HRESULT hrCreateDevice = _d2dFactory->CreateDevice(dxgiDevice.get(), createdD2DDevice.addressof());
    if (! CheckHR(hrCreateDevice, L"ID2D1Factory1::CreateDevice"))
    {
        return;
    }
    {
        std::lock_guard lock(_d2dDeviceMutex);
        _d2dDevice = createdD2DDevice;
    }
    const HRESULT hrCreateContext = createdD2DDevice->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE, _d2dContext.addressof());
    if (! CheckHR(hrCreateContext, L"ID2D1Device::CreateDeviceContext"))
    {
        return;
    }
    _d2dContext->SetUnitMode(D2D1_UNIT_MODE_DIPS);
    _d2dContext->SetDpi(_dpi, _dpi);
    _d2dContext->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);

    RecreateThemeBrushes();

    // Create placeholder icon for async loading
    CreatePlaceholderIcon();

    if (! _incrementalSearchIndicatorStrokeStyle && _d2dFactory)
    {
        D2D1_STROKE_STYLE_PROPERTIES props{};
        props.startCap   = D2D1_CAP_STYLE_ROUND;
        props.endCap     = D2D1_CAP_STYLE_ROUND;
        props.dashCap    = D2D1_CAP_STYLE_ROUND;
        props.lineJoin   = D2D1_LINE_JOIN_ROUND;
        props.miterLimit = 10.0f;
        props.dashStyle  = D2D1_DASH_STYLE_SOLID;

        const HRESULT hrStrokeStyle = _d2dFactory->CreateStrokeStyle(props, nullptr, 0, _incrementalSearchIndicatorStrokeStyle.addressof());
        static_cast<void>(CheckHR(hrStrokeStyle, L"ID2D1Factory1::CreateStrokeStyle(incremental search indicator)"));
    }
}

void FolderView::RecreateThemeBrushes()
{
    if (! _d2dContext)
    {
        return;
    }

    // Reset existing brushes
    _backgroundBrush.reset();
    _textBrush.reset();
    _detailsTextBrush.reset();
    _metadataTextBrush.reset();
    _selectionBrush.reset();
    _focusedBackgroundBrush.reset();
    _focusBrush.reset();
    _incrementalSearchHighlightBrush.reset();
    _incrementalSearchIndicatorBackgroundBrush.reset();
    _incrementalSearchIndicatorBorderBrush.reset();
    _incrementalSearchIndicatorTextBrush.reset();
    _incrementalSearchIndicatorShadowBrush.reset();
    _incrementalSearchIndicatorAccentBrush.reset();

    // Create brushes from theme colors
    const HRESULT hrBgBrush = _d2dContext->CreateSolidColorBrush(_theme.backgroundColor, _backgroundBrush.addressof());
    if (! CheckHR(hrBgBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(background)"))
    {
        return;
    }

    const HRESULT hrTextBrush = _d2dContext->CreateSolidColorBrush(_theme.textNormal, _textBrush.addressof());
    if (! CheckHR(hrTextBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(text)"))
    {
        return;
    }

    D2D1::ColorF detailsColor    = _theme.textNormal;
    detailsColor.a               = std::clamp(detailsColor.a * kDetailsTextAlpha, 0.0f, 1.0f);
    const HRESULT hrDetailsBrush = _d2dContext->CreateSolidColorBrush(detailsColor, _detailsTextBrush.addressof());
    if (! CheckHR(hrDetailsBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(details text)"))
    {
        return;
    }

    D2D1::ColorF metadataColor    = _theme.textNormal;
    metadataColor.a               = std::clamp(metadataColor.a * kMetadataTextAlpha, 0.0f, 1.0f);
    const HRESULT hrMetadataBrush = _d2dContext->CreateSolidColorBrush(metadataColor, _metadataTextBrush.addressof());
    if (! CheckHR(hrMetadataBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(metadata text)"))
    {
        return;
    }

    const HRESULT hrSelBrush = _d2dContext->CreateSolidColorBrush(_theme.itemBackgroundSelected, _selectionBrush.addressof());
    if (! CheckHR(hrSelBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(selection)"))
    {
        return;
    }

    const HRESULT hrFocusedBgBrush = _d2dContext->CreateSolidColorBrush(_theme.itemBackgroundFocused, _focusedBackgroundBrush.addressof());
    if (! CheckHR(hrFocusedBgBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(focused background)"))
    {
        return;
    }

    const HRESULT hrFocusBrush = _d2dContext->CreateSolidColorBrush(_theme.focusBorder, _focusBrush.addressof());
    if (! CheckHR(hrFocusBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(focus)"))
    {
        return;
    }

    const HRESULT hrIncrementalSearchBrush = _d2dContext->CreateSolidColorBrush(_theme.textSelected, _incrementalSearchHighlightBrush.addressof());
    if (! CheckHR(hrIncrementalSearchBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(incremental search highlight text)"))
    {
        return;
    }

    auto clamp01 = [](float value) noexcept -> float { return std::clamp(value, 0.0f, 1.0f); };
    auto lerp    = [](float a, float b, float t) noexcept -> float { return a + (b - a) * t; };
    auto blend   = [&](D2D1::ColorF base, const D2D1::ColorF& tint, float t) noexcept -> D2D1::ColorF
    {
        base.r = clamp01(lerp(base.r, tint.r, t));
        base.g = clamp01(lerp(base.g, tint.g, t));
        base.b = clamp01(lerp(base.b, tint.b, t));
        base.a = 1.0f;
        return base;
    };

    D2D1::ColorF indicatorBackground = _theme.backgroundColor;
    const float bgNudge              = _theme.darkBase ? 0.06f : -0.03f;
    indicatorBackground.r            = clamp01(indicatorBackground.r + bgNudge);
    indicatorBackground.g            = clamp01(indicatorBackground.g + bgNudge);
    indicatorBackground.b            = clamp01(indicatorBackground.b + bgNudge);
    indicatorBackground              = blend(indicatorBackground, _theme.focusBorder, _theme.darkBase ? 0.16f : 0.08f);

    D2D1::ColorF indicatorText = _theme.textNormal;
    indicatorText.a            = 1.0f;

    D2D1::ColorF indicatorShadow = D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f);

    const HRESULT hrIndicatorBg = _d2dContext->CreateSolidColorBrush(indicatorBackground, _incrementalSearchIndicatorBackgroundBrush.addressof());
    static_cast<void>(CheckHR(hrIndicatorBg, L"ID2D1DeviceContext::CreateSolidColorBrush(incremental search indicator background)"));

    const HRESULT hrIndicatorBorder = _d2dContext->CreateSolidColorBrush(_theme.focusBorder, _incrementalSearchIndicatorBorderBrush.addressof());
    static_cast<void>(CheckHR(hrIndicatorBorder, L"ID2D1DeviceContext::CreateSolidColorBrush(incremental search indicator border)"));

    const HRESULT hrIndicatorText = _d2dContext->CreateSolidColorBrush(indicatorText, _incrementalSearchIndicatorTextBrush.addressof());
    static_cast<void>(CheckHR(hrIndicatorText, L"ID2D1DeviceContext::CreateSolidColorBrush(incremental search indicator text)"));

    const HRESULT hrIndicatorShadow = _d2dContext->CreateSolidColorBrush(indicatorShadow, _incrementalSearchIndicatorShadowBrush.addressof());
    static_cast<void>(CheckHR(hrIndicatorShadow, L"ID2D1DeviceContext::CreateSolidColorBrush(incremental search indicator shadow)"));

    const HRESULT hrIndicatorAccent = _d2dContext->CreateSolidColorBrush(_theme.focusBorder, _incrementalSearchIndicatorAccentBrush.addressof());
    static_cast<void>(CheckHR(hrIndicatorAccent, L"ID2D1DeviceContext::CreateSolidColorBrush(incremental search indicator accent)"));
}

void FolderView::EnsureSwapChain()
{
    if (_clientSize.cx <= 0 || _clientSize.cy <= 0)
    {
        return;
    }
    if (! _d3dDevice)
    {
        return;
    }

    if (! _swapChain && ! _swapChainLegacy)
    {
        wil::com_ptr<IDXGIDevice> dxgiDevice;
        const HRESULT hrDxgiDevice2 = _d3dDevice ? _d3dDevice->QueryInterface(IID_PPV_ARGS(dxgiDevice.addressof())) : E_POINTER;
        if (! CheckHR(hrDxgiDevice2, L"ID3D11Device::QueryInterface IDXGIDevice"))
        {
            return;
        }
        wil::com_ptr<IDXGIAdapter> adapter;
        const HRESULT hrGetAdapter = dxgiDevice->GetAdapter(adapter.addressof());
        if (! CheckHR(hrGetAdapter, L"IDXGIDevice::GetAdapter"))
        {
            return;
        }
        wil::com_ptr<IDXGIFactory2> factory;
        const HRESULT hrGetParent = adapter->GetParent(IID_PPV_ARGS(factory.addressof()));
        if (! CheckHR(hrGetParent, L"IDXGIAdapter::GetParent"))
        {
            return;
        }

        DXGI_SWAP_CHAIN_DESC1 desc{};
        desc.Width            = static_cast<UINT>(_clientSize.cx);
        desc.Height           = static_cast<UINT>(_clientSize.cy);
        desc.Format           = DXGI_FORMAT_B8G8R8A8_UNORM;
        desc.Stereo           = FALSE;
        desc.SampleDesc.Count = 1;
        desc.BufferUsage      = DXGI_USAGE_RENDER_TARGET_OUTPUT;
        desc.BufferCount      = kSwapChainBufferCount;
        desc.Scaling          = DXGI_SCALING_NONE;
        desc.SwapEffect       = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
        desc.AlphaMode        = DXGI_ALPHA_MODE_IGNORE;
        desc.Flags            = 0;

        HRESULT hrSwapChain = factory->CreateSwapChainForHwnd(_d3dDevice.get(), _hWnd.get(), &desc, nullptr, nullptr, _swapChain.addressof());
        if (hrSwapChain == DXGI_ERROR_INVALID_CALL)
        {
            // Some older drivers require STRETCH; retry with that mode.
            desc.Scaling = DXGI_SCALING_STRETCH;
            hrSwapChain  = factory->CreateSwapChainForHwnd(_d3dDevice.get(), _hWnd.get(), &desc, nullptr, nullptr, _swapChain.addressof());
        }
        if (SUCCEEDED(hrSwapChain))
        {
            Debug::Info(L"FolderView: Created flip-model swap chain {}x{}", desc.Width, desc.Height);
            _supportsPresent1 = true;
            _swapChainLegacy.reset();
            const HRESULT hrAssociate = factory->MakeWindowAssociation(_hWnd.get(), DXGI_MWA_NO_ALT_ENTER | DXGI_MWA_NO_WINDOW_CHANGES);
            if (! CheckHR(hrAssociate, L"IDXGIFactory2::MakeWindowAssociation"))
            {
                _swapChain.reset();
                return;
            }
        }
        else
        {
            bool fallbackSucceeded = false;
            if (hrSwapChain == DXGI_ERROR_ACCESS_DENIED || hrSwapChain == DXGI_ERROR_INVALID_CALL || hrSwapChain == E_ACCESSDENIED)
            {
                wil::com_ptr<IDXGIFactory> factoryLegacy;
                if (FAILED(factory->QueryInterface(IID_PPV_ARGS(factoryLegacy.addressof()))))
                {
                    adapter->GetParent(IID_PPV_ARGS(factoryLegacy.addressof()));
                }

                if (factoryLegacy)
                {
                    DXGI_SWAP_CHAIN_DESC legacyDesc{};
                    legacyDesc.BufferDesc.Width                   = desc.Width;
                    legacyDesc.BufferDesc.Height                  = desc.Height;
                    legacyDesc.BufferDesc.Format                  = desc.Format;
                    legacyDesc.BufferDesc.RefreshRate.Numerator   = 60;
                    legacyDesc.BufferDesc.RefreshRate.Denominator = 1;
                    legacyDesc.BufferUsage                        = DXGI_USAGE_RENDER_TARGET_OUTPUT;
                    legacyDesc.OutputWindow                       = _hWnd.get();
                    legacyDesc.SampleDesc.Count                   = 1;
                    legacyDesc.SampleDesc.Quality                 = 0;
                    legacyDesc.Windowed                           = TRUE;
                    legacyDesc.BufferCount                        = kSwapChainBufferCount;
                    legacyDesc.SwapEffect                         = DXGI_SWAP_EFFECT_DISCARD;
                    legacyDesc.Flags                              = DXGI_SWAP_CHAIN_FLAG_ALLOW_MODE_SWITCH;

                    wil::com_ptr<IDXGISwapChain> legacySwap;
                    HRESULT hrLegacy = factoryLegacy->CreateSwapChain(_d3dDevice.get(), &legacyDesc, legacySwap.addressof());
                    if (SUCCEEDED(hrLegacy))
                    {
                        Debug::Warning(L"FolderView: Falling back to legacy swap chain {}x{}", legacyDesc.BufferDesc.Width, legacyDesc.BufferDesc.Height);
                        const HRESULT hrLegacyAssoc = factoryLegacy->MakeWindowAssociation(_hWnd.get(), DXGI_MWA_NO_ALT_ENTER | DXGI_MWA_NO_WINDOW_CHANGES);
                        if (! CheckHR(hrLegacyAssoc, L"IDXGIFactory::MakeWindowAssociation"))
                        {
                            return;
                        }
                        _swapChainLegacy = std::move(legacySwap);
                        _swapChain.reset();
                        _supportsPresent1 = false;
                        fallbackSucceeded = true;
                    }
                    else
                    {
                        CheckHR(hrLegacy, L"IDXGIFactory::CreateSwapChain");
                    }
                }
            }

            if (! fallbackSucceeded)
            {
                CheckHR(hrSwapChain, L"IDXGIFactory2::CreateSwapChainForHwnd");
                return;
            }
        }
    }

    IDXGISwapChain* activeSwapChain = nullptr;
    if (_supportsPresent1 && _swapChain)
    {
        activeSwapChain = _swapChain.get();
    }
    else if (_swapChainLegacy)
    {
        activeSwapChain = _swapChainLegacy.get();
    }

    if (! _d2dContext || ! activeSwapChain)
    {
        return;
    }

    // Only create the render target if we don't have one yet
    if (! _d2dTarget)
    {
        wil::com_ptr<IDXGISurface> surface;
        const HRESULT hrBuffer = activeSwapChain->GetBuffer(0, IID_PPV_ARGS(surface.addressof()));
        if (! CheckHR(hrBuffer, L"IDXGISwapChain::GetBuffer"))
        {
            return;
        }

        D2D1_BITMAP_PROPERTIES1 properties{};
        properties.pixelFormat   = D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED);
        properties.dpiX          = _dpi;
        properties.dpiY          = _dpi;
        properties.bitmapOptions = D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW;

        const HRESULT hrTarget = _d2dContext->CreateBitmapFromDxgiSurface(surface.get(), &properties, _d2dTarget.addressof());
        if (! CheckHR(hrTarget, L"ID2D1DeviceContext::CreateBitmapFromDxgiSurface"))
        {
            return;
        }
        _d2dContext->SetTarget(_d2dTarget.get());
    }
}

void FolderView::PrepareForSwapChainChange()
{
    Debug::Info(L"FolderView::PrepareForSwapChainChange");

    // Detach the D2D render target if we have one
    if (_d2dContext && _d2dTarget)
    {
        _d2dContext->SetTarget(nullptr);
    }
    _d2dTarget.reset();

    // Flush the D3D11 device context to release all buffer references
    // This is sufficient to allow swap chain resize without needing D2D Flush
    if (_d3dContext)
    {
        _d3dContext->ClearState();
        _d3dContext->Flush();
    }
}

void FolderView::ReleaseSwapChain()
{
    Debug::Info(L"FolderView::ReleaseSwapChain");
    PrepareForSwapChainChange();
    _swapChain.reset();
    _swapChainLegacy.reset();
    _supportsPresent1 = true;
}

bool FolderView::TryResizeSwapChain(UINT width, UINT height)
{
    if (! _swapChain && ! _swapChainLegacy)
    {
        return false;
    }

    PrepareForSwapChainChange();

    const UINT safeWidth  = std::max(1U, width);
    const UINT safeHeight = std::max(1U, height);

    HRESULT hr = S_OK;
    if (_swapChain)
    {
        hr = _swapChain->ResizeBuffers(kSwapChainBufferCount, safeWidth, safeHeight, DXGI_FORMAT_B8G8R8A8_UNORM, 0);
    }
    else if (_swapChainLegacy)
    {
        hr = _swapChainLegacy->ResizeBuffers(kSwapChainBufferCount, safeWidth, safeHeight, DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_SWAP_CHAIN_FLAG_ALLOW_MODE_SWITCH);
    }

    if (FAILED(hr))
    {
        Debug::Error(L"TryResizeSwapChain failed: 0x{:08X}", hr);
        ReportError(L"IDXGISwapChain::ResizeBuffers", hr);
        return false;
    }

    return true;
}

void FolderView::DiscardDeviceResources()
{
    ReleaseSwapChain();

    // Clear per-item cached icons: ID2D1Bitmap1 instances are tied to the originating ID2D1Device.
    for (auto& item : _items)
    {
        item.icon.reset();
    }

    wil::com_ptr<ID2D1Device> oldD2DDevice;
    {
        std::lock_guard lock(_d2dDeviceMutex);
        oldD2DDevice = _d2dDevice;
    }

    _backgroundBrush.reset();
    _textBrush.reset();
    _selectionBrush.reset();
    _focusedBackgroundBrush.reset();
    _focusBrush.reset();
    _incrementalSearchHighlightBrush.reset();
    _incrementalSearchIndicatorBackgroundBrush.reset();
    _incrementalSearchIndicatorBorderBrush.reset();
    _incrementalSearchIndicatorTextBrush.reset();
    _incrementalSearchIndicatorShadowBrush.reset();
    _incrementalSearchIndicatorAccentBrush.reset();
    _detailsTextBrush.reset();

    _placeholderFolderIcon.reset();
    _placeholderFileIcon.reset();
    _shortcutOverlayIcon.reset();

    _labelFormat.reset();
    _detailsFormat.reset();
    _incrementalSearchIndicatorLayout.reset();
    _incrementalSearchIndicatorLayoutText.clear();
    _incrementalSearchIndicatorLayoutMaxWidthDip = 0.0f;
    _incrementalSearchIndicatorLayoutMetrics     = {};
    if (_alertOverlay)
    {
        _alertOverlay->ResetDeviceResources();
        _alertOverlay->ResetTextResources();
    }

    _ellipsisSign.reset();
    _detailsEllipsisSign.reset();

    _dwriteFactory.reset();
    _wicFactory.reset();
    _d2dContext.reset();
    {
        std::lock_guard lock(_d2dDeviceMutex);
        _d2dDevice.reset();
    }

    IconCache::GetInstance().ClearDeviceCache(oldD2DDevice.get());

    _incrementalSearchIndicatorStrokeStyle.reset();
    _d2dFactory.reset();
    _d3dContext.reset();
    _d3dDevice.reset();
}

void FolderView::CreatePlaceholderIcon()
{
    if (! _d2dContext || ! _d2dFactory)
    {
        return;
    }

    // Log DPI information for high-DPI validation
    Debug::Info(L"FolderView: Creating placeholder icons at DPI={} ({}% scaling)", _dpi, static_cast<int>((_dpi / 96.0f) * 100.0f + 0.5f));

    // Create 48×48 Fluent Design placeholder icons for folders and files
    constexpr float size = 48.0f;

    // Create folder placeholder (rounded rectangle with tab)
    {
        // Create compatible render target for offscreen rendering
        wil::com_ptr<ID2D1BitmapRenderTarget> folderTarget;
        const D2D1_SIZE_F targetSize        = D2D1::SizeF(size, size);
        const D2D1_PIXEL_FORMAT pixelFormat = D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED);
        HRESULT hr = _d2dContext->CreateCompatibleRenderTarget(&targetSize, nullptr, &pixelFormat, D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_NONE, &folderTarget);
        if (SUCCEEDED(hr))
        {
            {
                folderTarget->BeginDraw();
                auto endDraw = wil::scope_exit([&] { folderTarget->EndDraw(); });
                folderTarget->Clear(D2D1::ColorF(0, 0, 0, 0)); // Transparent background

                // Create gradient brush (light blue to blue - Windows 11 folder colors)
                wil::com_ptr<ID2D1LinearGradientBrush> gradientBrush;
                D2D1_GRADIENT_STOP gradientStops[2]{};
                gradientStops[0].color    = D2D1::ColorF(0.565f, 0.792f, 0.976f, 1.0f); // RGB(144, 202, 249)
                gradientStops[0].position = 0.0f;
                gradientStops[1].color    = D2D1::ColorF(0.259f, 0.647f, 0.961f, 1.0f); // RGB(66, 165, 245)
                gradientStops[1].position = 1.0f;

                wil::com_ptr<ID2D1GradientStopCollection> gradientStopCollection;
                hr = folderTarget->CreateGradientStopCollection(gradientStops, 2, &gradientStopCollection);
                if (SUCCEEDED(hr))
                {
                    hr = folderTarget->CreateLinearGradientBrush(
                        D2D1::LinearGradientBrushProperties(D2D1::Point2F(24, 8), D2D1::Point2F(24, 40)), gradientStopCollection.get(), &gradientBrush);
                }

                if (gradientBrush)
                {
                    // Draw folder body (rounded rectangle)
                    D2D1_ROUNDED_RECT folderBody = D2D1::RoundedRect(D2D1::RectF(6, 14, 42, 40), 3.0f, 3.0f);
                    folderTarget->FillRoundedRectangle(folderBody, gradientBrush.get());

                    // Draw folder tab
                    D2D1_ROUNDED_RECT folderTab = D2D1::RoundedRect(D2D1::RectF(6, 8, 26, 14), 2.0f, 2.0f);
                    folderTarget->FillRoundedRectangle(folderTab, gradientBrush.get());
                }
            }
            folderTarget->GetBitmap(_placeholderFolderIcon.addressof());
        }
    }

    // Create file placeholder (document with folded corner)
    {
        wil::com_ptr<ID2D1BitmapRenderTarget> fileTarget;
        const D2D1_SIZE_F targetSize        = D2D1::SizeF(size, size);
        const D2D1_PIXEL_FORMAT pixelFormat = D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED);
        HRESULT hr = _d2dContext->CreateCompatibleRenderTarget(&targetSize, nullptr, &pixelFormat, D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_NONE, &fileTarget);
        if (SUCCEEDED(hr))
        {
            {
                fileTarget->BeginDraw();
                auto endDraw = wil::scope_exit([&] { fileTarget->EndDraw(); });
                fileTarget->Clear(D2D1::ColorF(0, 0, 0, 0)); // Transparent background

                // Create brushes
                wil::com_ptr<ID2D1SolidColorBrush> fillBrush;
                wil::com_ptr<ID2D1SolidColorBrush> outlineBrush;
                fileTarget->CreateSolidColorBrush(D2D1::ColorF(0.980f, 0.980f, 0.980f, 1.0f), &fillBrush);    // RGB(250, 250, 250)
                fileTarget->CreateSolidColorBrush(D2D1::ColorF(0.741f, 0.741f, 0.741f, 1.0f), &outlineBrush); // RGB(189, 189, 189)

                if (fillBrush && outlineBrush)
                {
                    // Create path geometry for document with folded corner
                    wil::com_ptr<ID2D1PathGeometry> docPath;
                    hr = _d2dFactory->CreatePathGeometry(&docPath);
                    if (SUCCEEDED(hr))
                    {
                        wil::com_ptr<ID2D1GeometrySink> sink;
                        hr = docPath->Open(&sink);
                        if (SUCCEEDED(hr))
                        {
                            // Document outline with folded corner (8×8 fold)
                            sink->BeginFigure(D2D1::Point2F(10, 8), D2D1_FIGURE_BEGIN_FILLED);
                            sink->AddLine(D2D1::Point2F(30, 8));  // Top edge
                            sink->AddLine(D2D1::Point2F(38, 16)); // Folded corner diagonal
                            sink->AddLine(D2D1::Point2F(38, 40)); // Right edge
                            sink->AddLine(D2D1::Point2F(10, 40)); // Bottom edge
                            sink->EndFigure(D2D1_FIGURE_END_CLOSED);
                            sink->Close();

                            // Fill and outline
                            fileTarget->FillGeometry(docPath.get(), fillBrush.get());
                            fileTarget->DrawGeometry(docPath.get(), outlineBrush.get(), 1.0f);

                            // Draw fold line
                            fileTarget->DrawLine(D2D1::Point2F(30, 8), D2D1::Point2F(38, 16), outlineBrush.get(), 1.0f);
                        }
                    }
                }
            }
            fileTarget->GetBitmap(_placeholderFileIcon.addressof());
        }
    }

    // Create shortcut overlay icon (16×16 arrow)
    {
        // Extract shortcut arrow from system stock icon
        SHSTOCKICONINFO sii{};
        sii.cbSize = sizeof(sii);
        HRESULT hr = SHGetStockIconInfo(SIID_LINK, SHGSI_ICON | SHGSI_SMALLICON, &sii);
        if (SUCCEEDED(hr) && sii.hIcon)
        {
            // Convert HICON to D2D bitmap
            wil::unique_hicon icon(sii.hIcon);

            // Get icon dimensions
            ICONINFO iconInfo{};
            if (GetIconInfo(icon.get(), &iconInfo))
            {
                wil::unique_hbitmap colorBitmap(iconInfo.hbmColor);
                wil::unique_hbitmap maskBitmap(iconInfo.hbmMask);

                BITMAP bm{};
                if (GetObjectW(colorBitmap.get(), sizeof(bm), &bm))
                {
                    const int width  = bm.bmWidth;
                    const int height = bm.bmHeight;

                    // Create compatible DC and draw icon
                    wil::unique_hdc_window hdcScreen(GetDC(nullptr));
                    wil::unique_hdc hdcMem(CreateCompatibleDC(hdcScreen.get()));

                    if (hdcMem)
                    {
                        BITMAPINFO bmi{};
                        bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
                        bmi.bmiHeader.biWidth       = width;
                        bmi.bmiHeader.biHeight      = -height;
                        bmi.bmiHeader.biPlanes      = 1;
                        bmi.bmiHeader.biBitCount    = 32;
                        bmi.bmiHeader.biCompression = BI_RGB;

                        void* pBits = nullptr;
                        wil::unique_hbitmap hBitmap(CreateDIBSection(hdcMem.get(), &bmi, DIB_RGB_COLORS, &pBits, nullptr, 0));
                        if (hBitmap && pBits)
                        {
                            std::memset(pBits, 0, static_cast<size_t>(width) * static_cast<size_t>(height) * 4);
                            auto oldBitmap = wil::SelectObject(hdcMem.get(), hBitmap.get());
                            DrawIconEx(hdcMem.get(), 0, 0, icon.get(), width, height, 0, nullptr, DI_NORMAL);
                            oldBitmap.reset();

                            // Premultiply alpha
                            auto* pixels            = static_cast<BYTE*>(pBits);
                            const size_t pixelCount = static_cast<size_t>(width) * static_cast<size_t>(height);
                            for (size_t i = 0; i < pixelCount; ++i)
                            {
                                const size_t offset = i * 4;
                                BYTE b              = pixels[offset + 0];
                                BYTE g              = pixels[offset + 1];
                                BYTE r              = pixels[offset + 2];
                                BYTE a              = pixels[offset + 3];

                                if (a > 0 && a < 255)
                                {
                                    const float alpha  = a / 255.0f;
                                    pixels[offset + 0] = static_cast<BYTE>(b * alpha + 0.5f);
                                    pixels[offset + 1] = static_cast<BYTE>(g * alpha + 0.5f);
                                    pixels[offset + 2] = static_cast<BYTE>(r * alpha + 0.5f);
                                }
                            }

                            // Create D2D bitmap
                            D2D1_BITMAP_PROPERTIES props{};
                            props.pixelFormat.format    = DXGI_FORMAT_B8G8R8A8_UNORM;
                            props.pixelFormat.alphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED;
                            props.dpiX                  = _dpi;
                            props.dpiY                  = _dpi;

                            const UINT32 pitch = static_cast<UINT32>(width) * 4u;
                            _d2dContext->CreateBitmap(
                                D2D1::SizeU(static_cast<UINT32>(width), static_cast<UINT32>(height)), pBits, pitch, &props, _shortcutOverlayIcon.addressof());
                        }
                    }
                }
            }
        }
    }
}

void FolderView::Render(const RECT& invalidRect)
{
    // std::wstring rectInfo = std::format(L"Rect({},{},{},{}) Items:{}", invalidRect.left, invalidRect.top, invalidRect.right, invalidRect.bottom,
    // _items.size()); TRACER_CTX(rectInfo.c_str());

    EnsureDeviceIndependentResources();
    EnsureDeviceResources();
    EnsureSwapChain();

    if (! _d2dContext || ! _d2dTarget)
    {
        Debug::Warning(L"FolderView::Render skipped - no valid render target");
        return;
    }

    RECT paintRect = invalidRect;
    if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top)
    {
        paintRect.left   = 0;
        paintRect.top    = 0;
        paintRect.right  = _clientSize.cx;
        paintRect.bottom = _clientSize.cy;
    }

    paintRect.left   = std::max<LONG>(0, paintRect.left);
    paintRect.top    = std::max<LONG>(0, paintRect.top);
    paintRect.right  = std::min<LONG>(_clientSize.cx, paintRect.right);
    paintRect.bottom = std::min<LONG>(_clientSize.cy, paintRect.bottom);

    D2D1_RECT_F dirtyDip = D2D1::RectF(DipFromPx(paintRect.left), DipFromPx(paintRect.top), DipFromPx(paintRect.right), DipFromPx(paintRect.bottom));

    HRESULT hr               = S_OK;
    const uint64_t nowTickMs = GetTickCount64();
    {
        _d2dContext->BeginDraw();
        auto endDraw = wil::scope_exit([&] { hr = _d2dContext->EndDraw(); });
        _d2dContext->SetTransform(D2D1::Matrix3x2F::Identity());
        _d2dContext->PushAxisAlignedClip(dirtyDip, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
        _d2dContext->FillRectangle(dirtyDip, _backgroundBrush.get());

        const float layoutLeft   = dirtyDip.left + _horizontalOffset;
        const float layoutRight  = dirtyDip.right + _horizontalOffset;
        const float layoutTop    = dirtyDip.top + _scrollOffset;
        const float layoutBottom = dirtyDip.bottom + _scrollOffset;

        auto drawIfVisible = [&](FolderItem& item)
        {
            const D2D1_RECT_F viewBounds = OffsetRect(item.bounds, -_horizontalOffset, -_scrollOffset);
            if (viewBounds.right < dirtyDip.left || viewBounds.left > dirtyDip.right || viewBounds.bottom < dirtyDip.top || viewBounds.top > dirtyDip.bottom)
            {
                return;
            }
            DrawItem(item);
        };

        if (_columnCounts.empty())
        {
            for (auto& item : _items)
            {
                drawIfVisible(item);
            }
        }
        else
        {
            const float columnStride = _tileWidthDip + kColumnSpacingDip;
            const float rowStride    = _tileHeightDip + kRowSpacingDip;
            if (columnStride <= 0.0f || rowStride <= 0.0f)
            {
                for (auto& item : _items)
                {
                    drawIfVisible(item);
                }
            }
            else
            {
                const float firstColumnLeft = kColumnSpacingDip;
                const float firstRowTop     = kRowSpacingDip;
                size_t columnBaseIndex      = 0;
                for (int column = 0; column < static_cast<int>(_columnCounts.size()) && columnBaseIndex < _items.size(); ++column)
                {
                    const int rows          = _columnCounts[static_cast<size_t>(column)];
                    const float columnLeft  = firstColumnLeft + static_cast<float>(column) * columnStride;
                    const float columnRight = columnLeft + _tileWidthDip;
                    if (columnRight < layoutLeft || columnLeft > layoutRight)
                    {
                        columnBaseIndex += static_cast<size_t>(rows);
                        continue;
                    }
                    if (rows <= 0)
                    {
                        continue;
                    }

                    int firstRow = static_cast<int>(std::floor((layoutTop - firstRowTop) / rowStride));
                    int lastRow  = static_cast<int>(std::ceil((layoutBottom - firstRowTop) / rowStride));
                    firstRow     = std::max(firstRow, 0);
                    lastRow      = std::max(lastRow, 0);
                    lastRow      = std::min(lastRow, rows - 1);
                    if (firstRow > rows - 1 || firstRow > lastRow)
                    {
                        columnBaseIndex += static_cast<size_t>(rows);
                        continue;
                    }

                    size_t startIndex = columnBaseIndex + static_cast<size_t>(firstRow);
                    size_t endIndex   = columnBaseIndex + static_cast<size_t>(lastRow);
                    if (startIndex >= _items.size())
                    {
                        break;
                    }
                    endIndex = std::min(endIndex, _items.size() - 1);

                    for (size_t idx = startIndex; idx <= endIndex; ++idx)
                    {
                        drawIfVisible(_items[idx]);
                    }

                    columnBaseIndex += static_cast<size_t>(rows);
                }
            }
        }

        if (_items.empty() && _displayedFolder.has_value() && ! _emptyStateMessage.empty() && _dwriteFactory && (_detailsFormat || _labelFormat))
        {
            bool hasOverlay = false;
            {
                std::lock_guard lock(_errorOverlayMutex);
                hasOverlay = _errorOverlay.has_value();
            }

            if (! hasOverlay && _emptyStateMessage.size() <= static_cast<size_t>(std::numeric_limits<UINT32>::max()))
            {
                const UINT32 length         = static_cast<UINT32>(_emptyStateMessage.size());
                const float clientWidthDip  = std::max(1.0f, DipFromPx(_clientSize.cx));
                const float clientHeightDip = std::max(1.0f, DipFromPx(_clientSize.cy));

                wil::com_ptr<IDWriteTextLayout> layout;
                const HRESULT hrLayout = _dwriteFactory->CreateTextLayout(_emptyStateMessage.data(),
                                                                          length,
                                                                          _detailsFormat ? _detailsFormat.get() : _labelFormat.get(),
                                                                          clientWidthDip,
                                                                          clientHeightDip,
                                                                          layout.addressof());
                if (SUCCEEDED(hrLayout) && layout)
                {
                    layout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
                    layout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                    layout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);

                    ID2D1SolidColorBrush* brush = _detailsTextBrush ? _detailsTextBrush.get() : _textBrush.get();
                    if (brush)
                    {
                        _d2dContext->DrawTextLayout(D2D1::Point2F(0.0f, 0.0f), layout.get(), brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    }
                }
            }
        }

        DrawIncrementalSearchIndicator(nowTickMs);
        DrawErrorOverlay();

        _d2dContext->PopAxisAlignedClip();
    }

    if (FAILED(hr))
    {
        ReportError(L"ID2D1DeviceContext::EndDraw", hr);
        ReleaseSwapChain();
        EnsureSwapChain();
    }
    else if (_supportsPresent1 && _swapChain)
    {
        DXGI_PRESENT_PARAMETERS params{};
        params.DirtyRectsCount = 1;
        params.pDirtyRects     = &paintRect;
        params.pScrollRect     = nullptr;
        params.pScrollOffset   = nullptr;
        _swapChain->Present1(1, 0, &params);
        ClearErrorOverlay(ErrorOverlayKind::Rendering);
    }
    else if (_swapChainLegacy)
    {
        _swapChainLegacy->Present(1, 0);
        ClearErrorOverlay(ErrorOverlayKind::Rendering);
    }
}

void FolderView::DrawIncrementalSearchIndicator(uint64_t nowTickMs)
{
    if (! _d2dContext || ! _incrementalSearchIndicatorBackgroundBrush || ! _incrementalSearchIndicatorBorderBrush || ! _incrementalSearchIndicatorTextBrush ||
        ! _incrementalSearchIndicatorShadowBrush || ! _incrementalSearchIndicatorAccentBrush)
    {
        return;
    }

    const float visibility = std::clamp(_incrementalSearchIndicatorVisibility, 0.0f, 1.0f);
    if (visibility <= 0.001f)
    {
        return;
    }

    const float clientWidthDip  = std::max(0.0f, DipFromPx(_clientSize.cx));
    const float clientHeightDip = std::max(0.0f, DipFromPx(_clientSize.cy));
    if (clientWidthDip <= 0.0f || clientHeightDip <= 0.0f)
    {
        return;
    }

    constexpr float kMarginDip           = 10.0f;
    constexpr float kHeightDip           = 30.0f;
    constexpr float kPaddingXDip         = 12.0f;
    constexpr float kIconSizeDip         = 14.0f;
    constexpr float kPillIconTextGapDip  = 8.0f;
    constexpr uint64_t kPulseMs          = 260;
    constexpr uint64_t kIconSwayPeriodMs = 3200;
    constexpr float kPi                  = 3.14159265358979323846f;

    float pulse         = 0.0f;
    float pulseProgress = 0.0f;
    if (_incrementalSearchIndicatorTypingPulseStart != 0)
    {
        const uint64_t elapsed = nowTickMs >= _incrementalSearchIndicatorTypingPulseStart ? (nowTickMs - _incrementalSearchIndicatorTypingPulseStart) : 0;
        pulseProgress          = std::clamp(static_cast<float>(elapsed) / static_cast<float>(kPulseMs), 0.0f, 1.0f);
        pulse                  = std::sin(pulseProgress * kPi);
    }

    float iconSwayDip = 0.0f;
    if (kIconSwayPeriodMs > 0)
    {
        const uint64_t phaseMs = nowTickMs % kIconSwayPeriodMs;
        const float phase      = static_cast<float>(phaseMs) / static_cast<float>(kIconSwayPeriodMs);
        const float amplitude  = kIconSizeDip * 0.18f;
        iconSwayDip            = std::sin(phase * 2.0f * kPi) * amplitude * visibility;
    }

    const float maxPillWidthDip = std::max(0.0f, clientWidthDip - 2.0f * kMarginDip);
    const std::wstring_view queryText{_incrementalSearchIndicatorDisplayQuery};

    float textWidthDip  = 0.0f;
    float textHeightDip = 0.0f;

    if (! queryText.empty() && _dwriteFactory && _labelFormat && maxPillWidthDip > 0.0f)
    {
        const float maxTextWidthDip = std::max(1.0f, maxPillWidthDip - (kPaddingXDip * 2.0f) - kIconSizeDip - kPillIconTextGapDip);

        const bool layoutNeedsUpdate = ! _incrementalSearchIndicatorLayout || std::wstring_view{_incrementalSearchIndicatorLayoutText} != queryText ||
                                       std::abs(_incrementalSearchIndicatorLayoutMaxWidthDip - maxTextWidthDip) > 0.5f;

        if (layoutNeedsUpdate)
        {
            _incrementalSearchIndicatorLayoutText.assign(queryText);
            _incrementalSearchIndicatorLayoutMaxWidthDip = maxTextWidthDip;
            _incrementalSearchIndicatorLayout.reset();
            _incrementalSearchIndicatorLayoutMetrics = {};

            if (_incrementalSearchIndicatorLayoutText.size() <= static_cast<size_t>(std::numeric_limits<UINT32>::max()))
            {
                wil::com_ptr<IDWriteTextLayout> layout;
                const HRESULT hrLayout = _dwriteFactory->CreateTextLayout(_incrementalSearchIndicatorLayoutText.c_str(),
                                                                          static_cast<UINT32>(_incrementalSearchIndicatorLayoutText.size()),
                                                                          _labelFormat.get(),
                                                                          maxTextWidthDip,
                                                                          kHeightDip,
                                                                          layout.addressof());
                if (SUCCEEDED(hrLayout) && layout)
                {
                    DWRITE_TEXT_METRICS metrics{};
                    if (SUCCEEDED(layout->GetMetrics(&metrics)))
                    {
                        _incrementalSearchIndicatorLayoutMetrics = metrics;
                    }
                    _incrementalSearchIndicatorLayout = std::move(layout);
                }
            }
        }

        if (_incrementalSearchIndicatorLayout)
        {
            textWidthDip  = std::min(_incrementalSearchIndicatorLayoutMetrics.widthIncludingTrailingWhitespace, maxTextWidthDip);
            textHeightDip = _incrementalSearchIndicatorLayoutMetrics.height;
        }
    }

    float pillWidthDip = kPaddingXDip + kIconSizeDip + kPaddingXDip;
    if (textWidthDip > 0.0f)
    {
        pillWidthDip = kPaddingXDip + kIconSizeDip + kPillIconTextGapDip + textWidthDip + kPaddingXDip;
    }
    pillWidthDip = std::clamp(pillWidthDip, 0.0f, maxPillWidthDip);

    float x = clientWidthDip - kMarginDip - pillWidthDip;
    float y = kMarginDip;

    const float slide = 1.0f - visibility;
    x += slide * 18.0f;
    y -= slide * 10.0f;

    const float cornerRadiusDip = kHeightDip * 0.5f;

    D2D1_ROUNDED_RECT shadow =
        D2D1::RoundedRect(D2D1::RectF(x + 2.0f, y + 2.0f, x + pillWidthDip + 2.0f, y + kHeightDip + 2.0f), cornerRadiusDip, cornerRadiusDip);
    D2D1_ROUNDED_RECT pill = D2D1::RoundedRect(D2D1::RectF(x, y, x + pillWidthDip, y + kHeightDip), cornerRadiusDip, cornerRadiusDip);

    const float shadowOpacity = visibility * (_theme.darkBase ? 0.35f : 0.22f);
    _incrementalSearchIndicatorShadowBrush->SetOpacity(shadowOpacity);
    _d2dContext->FillRoundedRectangle(shadow, _incrementalSearchIndicatorShadowBrush.get());

    const float backgroundOpacity = visibility * (_theme.darkBase ? 0.80f : 0.92f);
    _incrementalSearchIndicatorBackgroundBrush->SetOpacity(backgroundOpacity);
    _d2dContext->FillRoundedRectangle(pill, _incrementalSearchIndicatorBackgroundBrush.get());

    const float borderOpacity = visibility * (0.55f + 0.25f * pulse);
    _incrementalSearchIndicatorBorderBrush->SetOpacity(borderOpacity);
    _d2dContext->DrawRoundedRectangle(pill, _incrementalSearchIndicatorBorderBrush.get(), 1.0f + 0.8f * pulse);

    const float iconCenterX  = x + kPaddingXDip + kIconSizeDip * 0.5f + iconSwayDip;
    const float iconCenterY  = y + kHeightDip * 0.5f;
    const float iconHalfSize = kIconSizeDip * 0.5f;
    const float iconBarHalfW = kIconSizeDip * 0.35f;

    const float iconStroke = 1.5f + 0.6f * pulse;
    auto* strokeStyle      = _incrementalSearchIndicatorStrokeStyle.get();

    _incrementalSearchIndicatorAccentBrush->SetOpacity(visibility);
    _d2dContext->DrawLine(D2D1::Point2F(iconCenterX - iconBarHalfW, iconCenterY - iconHalfSize),
                          D2D1::Point2F(iconCenterX + iconBarHalfW, iconCenterY - iconHalfSize),
                          _incrementalSearchIndicatorAccentBrush.get(),
                          iconStroke,
                          strokeStyle);
    _d2dContext->DrawLine(D2D1::Point2F(iconCenterX - iconBarHalfW, iconCenterY + iconHalfSize),
                          D2D1::Point2F(iconCenterX + iconBarHalfW, iconCenterY + iconHalfSize),
                          _incrementalSearchIndicatorAccentBrush.get(),
                          iconStroke,
                          strokeStyle);
    _d2dContext->DrawLine(D2D1::Point2F(iconCenterX, iconCenterY - iconHalfSize),
                          D2D1::Point2F(iconCenterX, iconCenterY + iconHalfSize),
                          _incrementalSearchIndicatorAccentBrush.get(),
                          iconStroke,
                          strokeStyle);

    if (_incrementalSearchIndicatorLayout)
    {
        const float textX = x + kPaddingXDip + kIconSizeDip + kPillIconTextGapDip;
        const float textY = y + (kHeightDip - textHeightDip) * 0.5f;

        _incrementalSearchIndicatorTextBrush->SetOpacity(visibility);
        _d2dContext->DrawTextLayout(
            D2D1::Point2F(textX, textY), _incrementalSearchIndicatorLayout.get(), _incrementalSearchIndicatorTextBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

        if (pulse > 0.0f && textWidthDip > 0.0f)
        {
            const float underlineLen     = textWidthDip * std::clamp(pulseProgress * 1.35f, 0.0f, 1.0f);
            const float underlineOpacity = visibility * (0.20f + 0.60f * pulse);
            _incrementalSearchIndicatorAccentBrush->SetOpacity(underlineOpacity);
            const float underlineY = y + kHeightDip - 6.0f;
            _d2dContext->DrawLine(D2D1::Point2F(textX, underlineY),
                                  D2D1::Point2F(textX + underlineLen, underlineY),
                                  _incrementalSearchIndicatorAccentBrush.get(),
                                  1.6f + 0.6f * pulse,
                                  strokeStyle);
        }
    }
}

void FolderView::DrawItem(FolderItem& item)
{
    // Ensure text layout is created lazily before rendering
    const float labelWidth = std::max(0.0f, _tileWidthDip - (kLabelHorizontalPaddingDip * 2.0f) - _iconSizeDip - kIconTextGapDip);
    EnsureItemTextLayout(item, labelWidth);

    D2D1_RECT_F bounds = OffsetRect(item.bounds, -_horizontalOffset, -_scrollOffset);

    // Determine item state for color selection
    const bool isHovered = (_hoveredIndex != static_cast<size_t>(-1) && std::addressof(item) == std::addressof(_items[_hoveredIndex]));

    const float itemWidth                 = std::max(0.0f, bounds.right - bounds.left);
    const float itemHeight                = std::max(0.0f, bounds.bottom - bounds.top);
    const float maxCornerRadius           = std::min(itemWidth, itemHeight) * 0.5f;
    const float cornerRadius              = std::min(kSelectionCornerRadiusDip, maxCornerRadius);
    const D2D1_ROUNDED_RECT roundedBounds = D2D1::RoundedRect(bounds, cornerRadius, cornerRadius);

    const bool selectionActive = _paneFocused;

    auto compositeOverBackground = [&](const D2D1::ColorF& overlay) noexcept -> D2D1::ColorF
    {
        const float alpha = std::clamp(overlay.a, 0.0f, 1.0f);
        D2D1::ColorF result(0.0f, 0.0f, 0.0f, 1.0f);
        result.r = overlay.r * alpha + _theme.backgroundColor.r * (1.0f - alpha);
        result.g = overlay.g * alpha + _theme.backgroundColor.g * (1.0f - alpha);
        result.b = overlay.b * alpha + _theme.backgroundColor.b * (1.0f - alpha);
        result.a = 1.0f;
        return result;
    };

    D2D1::ColorF selectionBackground            = _theme.itemBackgroundSelected;
    D2D1::ColorF selectionBackgroundForContrast = selectionBackground;
    if (item.selected)
    {
        if (_theme.rainbowMode)
        {
            const uint32_t hash = item.stableHash32;
            const float hue     = static_cast<float>(hash % 360u);
            const float sat     = 0.85f;
            const float val     = _theme.darkBase ? 0.75f : 0.90f;
            selectionBackground = ColorFromHSV(hue, sat, val);
            selectionBackground.a =
                selectionActive ? std::clamp(_theme.itemBackgroundSelected.a, 0.0f, 1.0f) : std::clamp(_theme.itemBackgroundSelectedInactive.a, 0.0f, 1.0f);
        }
        else
        {
            selectionBackground = selectionActive ? _theme.itemBackgroundSelected : _theme.itemBackgroundSelectedInactive;
        }
        selectionBackgroundForContrast = compositeOverBackground(selectionBackground);
    }

    // Draw background based on state
    if (item.selected)
    {
        if (_selectionBrush)
        {
            _selectionBrush->SetColor(selectionBackground);
            _d2dContext->FillRoundedRectangle(roundedBounds, _selectionBrush.get());
        }
    }
    else if (item.focused && _paneFocused)
    {
        if (_focusedBackgroundBrush)
        {
            if (_theme.rainbowMode)
            {
                const uint32_t hash = item.stableHash32;
                const float hue     = static_cast<float>(hash % 360u);
                const float sat     = 0.85f;
                const float val     = _theme.darkBase ? 0.70f : 0.92f;
                D2D1::ColorF color  = ColorFromHSV(hue, sat, val);
                color.a             = _theme.itemBackgroundFocused.a;
                _focusedBackgroundBrush->SetColor(color);
            }
            else
            {
                _focusedBackgroundBrush->SetColor(_theme.itemBackgroundFocused);
            }
            _d2dContext->FillRoundedRectangle(roundedBounds, _focusedBackgroundBrush.get());
        }
    }
    else if (isHovered)
    {
        // Create temporary brush for hover state
        wil::com_ptr<ID2D1SolidColorBrush> hoverBrush;
        const HRESULT hrHover = _d2dContext->CreateSolidColorBrush(_theme.itemBackgroundHovered, hoverBrush.addressof());
        if (SUCCEEDED(hrHover) && hoverBrush)
        {
            _d2dContext->FillRoundedRectangle(roundedBounds, hoverBrush.get());
        }
    }

    // Draw focus border
    if (item.focused)
    {
        if (_focusBrush)
        {
            const float strokeThickness = _paneFocused ? kFocusStrokeThicknessDip : kFocusStrokeThicknessUnfocusedDip;
            const float inset           = strokeThickness * 0.5f;

            D2D1_RECT_F focusBounds = D2D1::RectF(bounds.left + inset, bounds.top + inset, bounds.right - inset, bounds.bottom - inset);
            const float focusWidth  = std::max(0.0f, focusBounds.right - focusBounds.left);
            const float focusHeight = std::max(0.0f, focusBounds.bottom - focusBounds.top);
            if (focusWidth > 0.0f && focusHeight > 0.0f)
            {
                const float maxFocusCornerRadius           = std::min(focusWidth, focusHeight) * 0.5f;
                const float focusCornerRadius              = std::min(std::max(0.0f, cornerRadius - inset), maxFocusCornerRadius);
                const D2D1_ROUNDED_RECT focusRoundedBounds = D2D1::RoundedRect(focusBounds, focusCornerRadius, focusCornerRadius);

                D2D1::ColorF focusColor = _theme.focusBorder;
                if (item.selected)
                {
                    const COLORREF contrast = ChooseContrastingTextColor(ColorToCOLORREF(selectionBackgroundForContrast));
                    focusColor              = ColorFromCOLORREF(contrast);
                }
                else if (_theme.rainbowMode)
                {
                    const uint32_t hash = item.stableHash32;
                    const float hue     = static_cast<float>(hash % 360u);
                    const float sat     = 0.85f;
                    const float val     = _theme.darkBase ? 0.85f : 0.80f;
                    focusColor          = ColorFromHSV(hue, sat, val);
                }

                if (! _paneFocused)
                {
                    focusColor.a = std::clamp(focusColor.a * kFocusBorderOpacityUnfocused, 0.0f, 1.0f);
                }

                _focusBrush->SetColor(focusColor);
                _d2dContext->DrawRoundedRectangle(focusRoundedBounds, _focusBrush.get(), strokeThickness);
            }
        }
    }

    const float contentTop    = bounds.top + kLabelVerticalPaddingDip;
    const float contentBottom = bounds.bottom - kLabelVerticalPaddingDip;
    const float contentHeight = std::max(0.0f, contentBottom - contentTop);

    const float iconLeft = bounds.left + kLabelHorizontalPaddingDip;
    const float iconTop  = _displayMode == DisplayMode::Brief ? contentTop + std::max(0.0f, (contentHeight - _iconSizeDip) * 0.5f) : contentTop;
    D2D1_RECT_F iconRect = D2D1::RectF(iconLeft, iconTop, iconLeft + _iconSizeDip, iconTop + _iconSizeDip);
    if (item.icon)
    {
        // Render icon with nearest neighbor interpolation for crisp pixel-perfect rendering
        _d2dContext->DrawBitmap(item.icon.get(), iconRect, 1.0f, D2D1_INTERPOLATION_MODE_NEAREST_NEIGHBOR);

        // Render shortcut overlay if applicable
        if (item.isShortcut && _shortcutOverlayIcon)
        {
            // Position overlay at bottom-right corner of icon
            const float overlaySize = _iconSizeDip * 0.5f; // Half icon size for overlay
            D2D1_RECT_F overlayRect = D2D1::RectF(iconRect.right - overlaySize, iconRect.bottom - overlaySize, iconRect.right, iconRect.bottom);
            _d2dContext->DrawBitmap(_shortcutOverlayIcon.get(), &overlayRect, 1.0f, D2D1_INTERPOLATION_MODE_LINEAR);
        }
    }
    else
    {
        // Select appropriate placeholder based on item type
        auto& placeholder = item.isDirectory ? _placeholderFolderIcon : _placeholderFileIcon;
        if (placeholder)
        {
            // Draw placeholder with reduced opacity and linear interpolation
            _d2dContext->DrawBitmap(placeholder.get(), &iconRect, 0.4f, D2D1_INTERPOLATION_MODE_LINEAR);
        }
        else
        {
            // Fallback if placeholders not created
            _d2dContext->FillRectangle(iconRect, _backgroundBrush.get());
            _d2dContext->DrawRectangle(iconRect, _focusBrush.get(), 1.0f);
        }
    }

    const float labelLeft      = iconRect.right + kIconTextGapDip;
    const float labelRight     = bounds.right - kLabelHorizontalPaddingDip;
    const float availableWidth = std::max(0.0f, labelRight - labelLeft);

    // Select text brush based on selection state
    ID2D1SolidColorBrush* textBrush = _textBrush.get();
    wil::com_ptr<ID2D1SolidColorBrush> selectedTextBrush;
    if (item.selected)
    {
        D2D1::ColorF selectedTextColor = selectionActive ? _theme.textSelected : _theme.textSelectedInactive;
        if (_theme.rainbowMode)
        {
            const float luminance =
                0.2126f * selectionBackgroundForContrast.r + 0.7152f * selectionBackgroundForContrast.g + 0.0722f * selectionBackgroundForContrast.b;
            selectedTextColor = luminance > 0.60f ? D2D1::ColorF(D2D1::ColorF::Black) : D2D1::ColorF(D2D1::ColorF::White);
        }

        const HRESULT hrSelText = _d2dContext->CreateSolidColorBrush(selectedTextColor, selectedTextBrush.addressof());
        if (SUCCEEDED(hrSelText) && selectedTextBrush)
        {
            textBrush = selectedTextBrush.get();
        }
    }

    auto drawIncrementalSearchHighlight = [&](D2D1_POINT_2F origin, const DWRITE_TEXT_RANGE& highlightRange) noexcept
    {
        constexpr float kHighlightPaddingXDip     = 2.0f;
        constexpr float kHighlightPaddingYDip     = 1.0f;
        constexpr float kHighlightCornerRadiusDip = 2.0f;
        constexpr float kSelectedOverlayAlpha     = 0.25f;

        if (! _d2dContext || ! _selectionBrush || ! item.labelLayout)
        {
            return;
        }

        if (! _incrementalSearch.active || _incrementalSearch.query.empty() || highlightRange.length == 0)
        {
            return;
        }

        if (item.displayName.size() > static_cast<size_t>(std::numeric_limits<UINT32>::max()))
        {
            return;
        }

        const UINT32 textLength = static_cast<UINT32>(item.displayName.size());
        if (highlightRange.startPosition >= textLength)
        {
            return;
        }

        DWRITE_TEXT_RANGE range{};
        range.startPosition = highlightRange.startPosition;
        range.length        = std::min(highlightRange.length, textLength - range.startPosition);
        if (range.length == 0)
        {
            return;
        }

        D2D1::ColorF highlightColor = _paneFocused ? _theme.itemBackgroundSelected : _theme.itemBackgroundSelectedInactive;
        if (item.selected)
        {
            D2D1_COLOR_F textColor = _theme.textSelected;
            if (textBrush)
            {
                textColor = textBrush->GetColor();
            }
            const float textLuminance = 0.2126f * textColor.r + 0.7152f * textColor.g + 0.0722f * textColor.b;
            const bool textIsLight    = textLuminance > 0.60f;

            const float backgroundLuminance =
                0.2126f * selectionBackgroundForContrast.r + 0.7152f * selectionBackgroundForContrast.g + 0.0722f * selectionBackgroundForContrast.b;

            const float preferredOverlayLum = textIsLight ? 0.0f : 1.0f;
            const float deltaLumPreferred   = std::abs(backgroundLuminance - preferredOverlayLum);
            const float effectiveChange     = deltaLumPreferred * kSelectedOverlayAlpha;
            const bool usePreferredOverlay  = effectiveChange >= 0.08f;

            const bool useBlackOverlay = usePreferredOverlay ? textIsLight : ! textIsLight;
            highlightColor = useBlackOverlay ? D2D1::ColorF(0.0f, 0.0f, 0.0f, kSelectedOverlayAlpha) : D2D1::ColorF(1.0f, 1.0f, 1.0f, kSelectedOverlayAlpha);
        }

        std::array<DWRITE_HIT_TEST_METRICS, 4> hitTestMetrics{};
        UINT32 metricsCount = 0;
        HRESULT hr          = item.labelLayout->HitTestTextRange(
            range.startPosition, range.length, origin.x, origin.y, hitTestMetrics.data(), static_cast<UINT32>(hitTestMetrics.size()), &metricsCount);

        std::vector<DWRITE_HIT_TEST_METRICS> dynamicMetrics;
        if (hr == E_NOT_SUFFICIENT_BUFFER)
        {
            if (metricsCount == 0)
            {
                return;
            }
            dynamicMetrics.resize(metricsCount);
            hr = item.labelLayout->HitTestTextRange(
                range.startPosition, range.length, origin.x, origin.y, dynamicMetrics.data(), static_cast<UINT32>(dynamicMetrics.size()), &metricsCount);
        }

        if (FAILED(hr) || metricsCount == 0)
        {
            return;
        }

        _selectionBrush->SetColor(highlightColor);

        const auto metricsData  = dynamicMetrics.empty() ? hitTestMetrics.data() : dynamicMetrics.data();
        const UINT32 metricsMax = dynamicMetrics.empty() ? static_cast<UINT32>(hitTestMetrics.size()) : static_cast<UINT32>(dynamicMetrics.size());
        const UINT32 count      = std::min(metricsCount, metricsMax);
        for (UINT32 i = 0; i < count; ++i)
        {
            const DWRITE_HIT_TEST_METRICS& metrics = metricsData[i];

            D2D1_RECT_F rect = D2D1::RectF(metrics.left, metrics.top, metrics.left + metrics.width, metrics.top + metrics.height);
            rect.left        = rect.left - kHighlightPaddingXDip;
            rect.right       = rect.right + kHighlightPaddingXDip;
            rect.top         = rect.top - kHighlightPaddingYDip;
            rect.bottom      = rect.bottom + kHighlightPaddingYDip;

            const float rectWidth  = std::max(0.0f, rect.right - rect.left);
            const float rectHeight = std::max(0.0f, rect.bottom - rect.top);
            if (rectWidth <= 0.0f || rectHeight <= 0.0f)
            {
                continue;
            }

            const float maxRadius = std::min(rectWidth, rectHeight) * 0.5f;
            const float radius    = std::min(kHighlightCornerRadiusDip, maxRadius);
            _d2dContext->FillRoundedRectangle(D2D1::RoundedRect(rect, radius, radius), _selectionBrush.get());
        }
    };

    std::optional<DWRITE_TEXT_RANGE> incrementalSearchRange;
    if (item.labelLayout)
    {
        if (item.displayName.size() <= static_cast<size_t>(std::numeric_limits<UINT32>::max()))
        {
            const UINT32 textLength = static_cast<UINT32>(item.displayName.size());
            if (textLength > 0)
            {
                DWRITE_TEXT_RANGE clearRange{};
                clearRange.startPosition = 0;
                clearRange.length        = textLength;
                static_cast<void>(item.labelLayout->SetDrawingEffect(nullptr, clearRange));
            }

            const std::optional<UINT32> matchOffset = FindIncrementalSearchMatchOffset(item.displayName);
            if (matchOffset.has_value() && _incrementalSearch.query.size() <= static_cast<size_t>(std::numeric_limits<UINT32>::max()))
            {
                DWRITE_TEXT_RANGE range{};
                range.startPosition = matchOffset.value();

                if (range.startPosition < textLength)
                {
                    const UINT32 queryLength = static_cast<UINT32>(_incrementalSearch.query.size());
                    range.length             = std::min(queryLength, textLength - range.startPosition);

                    if (range.length > 0)
                    {
                        incrementalSearchRange = range;
                        if (! item.selected && _incrementalSearchHighlightBrush)
                        {
                            const D2D1::ColorF highlightTextColor = _paneFocused ? _theme.textSelected : _theme.textSelectedInactive;
                            _incrementalSearchHighlightBrush->SetColor(highlightTextColor);
                            static_cast<void>(item.labelLayout->SetDrawingEffect(_incrementalSearchHighlightBrush.get(), range));
                        }
                    }
                }
            }
        }
    }

    if (item.labelLayout)
    {
        if (_displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed)
        {
            const float nameHeight = item.labelMetrics.height > 0.0f ? item.labelMetrics.height : std::max(0.0f, contentHeight * 0.5f);
            D2D1_POINT_2F origin{labelLeft, contentTop};
            if (incrementalSearchRange.has_value())
            {
                drawIncrementalSearchHighlight(origin, incrementalSearchRange.value());
            }
            _d2dContext->DrawTextLayout(origin, item.labelLayout.get(), textBrush, D2D1_DRAW_TEXT_OPTIONS_CLIP);

            ID2D1SolidColorBrush* detailsBrush = item.selected ? textBrush : (_detailsTextBrush ? _detailsTextBrush.get() : textBrush);

            const float detailsTop = contentTop + nameHeight + kDetailsGapDip;
            if (item.detailsLayout)
            {
                D2D1_POINT_2F detailsOrigin{labelLeft, detailsTop};
                _d2dContext->DrawTextLayout(detailsOrigin, item.detailsLayout.get(), detailsBrush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }
            else if (! item.detailsText.empty() && _detailsFormat)
            {
                D2D1_RECT_F detailsRect = D2D1::RectF(labelLeft, detailsTop, labelLeft + availableWidth, contentBottom);
                _d2dContext->DrawTextW(item.detailsText.c_str(),
                                       static_cast<UINT32>(item.detailsText.length()),
                                       _detailsFormat.get(),
                                       detailsRect,
                                       detailsBrush,
                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }

            if (_displayMode == DisplayMode::ExtraDetailed)
            {
                const bool hasDetails = item.detailsLayout || (! item.detailsText.empty());
                const float detailsHeight =
                    hasDetails ? (item.detailsMetrics.height > 0.0f ? item.detailsMetrics.height : std::max(0.0f, _detailsLineHeightDip)) : 0.0f;
                const float metadataTop = hasDetails ? (detailsTop + std::max(0.0f, detailsHeight) + kDetailsGapDip) : detailsTop;

                ID2D1SolidColorBrush* metadataBrush = item.selected ? textBrush : (_metadataTextBrush ? _metadataTextBrush.get() : detailsBrush);
                if (item.metadataLayout)
                {
                    D2D1_POINT_2F metadataOrigin{labelLeft, metadataTop};
                    _d2dContext->DrawTextLayout(metadataOrigin, item.metadataLayout.get(), metadataBrush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
                }
                else if (! item.metadataText.empty() && _detailsFormat)
                {
                    D2D1_RECT_F metadataRect = D2D1::RectF(labelLeft, metadataTop, labelLeft + availableWidth, contentBottom);
                    _d2dContext->DrawTextW(item.metadataText.c_str(),
                                           static_cast<UINT32>(item.metadataText.length()),
                                           _detailsFormat.get(),
                                           metadataRect,
                                           metadataBrush,
                                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
                }
            }
        }
        else
        {
            const float metricsHeight = item.labelMetrics.height > 0.0f ? item.labelMetrics.height : contentHeight;
            const float offsetY       = std::max(0.0f, (contentHeight - metricsHeight) * 0.5f);
            D2D1_POINT_2F origin{labelLeft, contentTop + offsetY};
            if (incrementalSearchRange.has_value())
            {
                drawIncrementalSearchHighlight(origin, incrementalSearchRange.value());
            }
            _d2dContext->DrawTextLayout(origin, item.labelLayout.get(), textBrush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
        }
    }
    else
    {
        if (_displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed)
        {
            const float detailsHeight  = _detailsLineHeightDip > 0.0f ? _detailsLineHeightDip : 12.0f;
            const float metadataHeight = (_displayMode == DisplayMode::ExtraDetailed && _metadataLineHeightDip > 0.0f) ? _metadataLineHeightDip : 0.0f;
            const float nameBottom =
                std::max(contentTop, contentBottom - detailsHeight - kDetailsGapDip - (metadataHeight > 0.0f ? (metadataHeight + kDetailsGapDip) : 0.0f));

            D2D1_RECT_F labelRect = D2D1::RectF(labelLeft, contentTop, labelLeft + availableWidth, nameBottom);
            _d2dContext->DrawTextW(
                item.displayName.data(), static_cast<UINT32>(item.displayName.length()), _labelFormat.get(), labelRect, textBrush, D2D1_DRAW_TEXT_OPTIONS_CLIP);

            ID2D1SolidColorBrush* detailsBrush = item.selected ? textBrush : (_detailsTextBrush ? _detailsTextBrush.get() : textBrush);

            if (! item.detailsText.empty() && _detailsFormat)
            {
                D2D1_RECT_F detailsRect = D2D1::RectF(labelLeft, nameBottom + kDetailsGapDip, labelLeft + availableWidth, contentBottom);
                _d2dContext->DrawTextW(item.detailsText.c_str(),
                                       static_cast<UINT32>(item.detailsText.length()),
                                       _detailsFormat.get(),
                                       detailsRect,
                                       detailsBrush,
                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }

            if (_displayMode == DisplayMode::ExtraDetailed && ! item.metadataText.empty() && _detailsFormat)
            {
                const bool hasDetails               = ! item.detailsText.empty();
                const float detailsBottom           = nameBottom + kDetailsGapDip + (hasDetails ? detailsHeight : 0.0f);
                const float metadataTop             = hasDetails ? (detailsBottom + kDetailsGapDip) : detailsBottom;
                ID2D1SolidColorBrush* metadataBrush = item.selected ? textBrush : (_metadataTextBrush ? _metadataTextBrush.get() : detailsBrush);
                D2D1_RECT_F metadataRect            = D2D1::RectF(labelLeft, metadataTop, labelLeft + availableWidth, contentBottom);
                _d2dContext->DrawTextW(item.metadataText.c_str(),
                                       static_cast<UINT32>(item.metadataText.length()),
                                       _detailsFormat.get(),
                                       metadataRect,
                                       metadataBrush,
                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }
        }
        else
        {
            D2D1_RECT_F labelRect = D2D1::RectF(labelLeft, contentTop, labelLeft + availableWidth, contentBottom);
            _d2dContext->DrawTextW(
                item.displayName.data(), static_cast<UINT32>(item.displayName.length()), _labelFormat.get(), labelRect, textBrush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
        }
    }
}

D2D1_RECT_F FolderView::OffsetRect(const D2D1_RECT_F& rect, float dx, float dy) const
{
    return D2D1::RectF(rect.left + dx, rect.top + dy, rect.right + dx, rect.bottom + dy);
}

RECT FolderView::ToPixelRect(const D2D1_RECT_F& rect, float dpi)
{
    RECT r{};
    r.left   = static_cast<LONG>(std::floor(rect.left * dpi / 96.0f));
    r.top    = static_cast<LONG>(std::floor(rect.top * dpi / 96.0f));
    r.right  = static_cast<LONG>(std::ceil(rect.right * dpi / 96.0f));
    r.bottom = static_cast<LONG>(std::ceil(rect.bottom * dpi / 96.0f));
    return r;
}

bool FolderView::RectIntersects(const D2D1_RECT_F& rect, const RECT& pixelRect, float dpi)
{
    RECT item = ToPixelRect(rect, dpi);
    RECT intersection{};
    return IntersectRect(&intersection, &item, &pixelRect) != 0;
}
