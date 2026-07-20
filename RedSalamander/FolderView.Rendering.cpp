#include "FolderViewInternal.h"

#include "DxUi/DxUi.FrameRuntime.h"
#include "DxUi/DxUi.Typography.h"
#include "DxUi/DxUi.h"
#include "FluentIcons.h"
#include "FolderViewThumbnailGeometry.h"
#ifdef ENABLE_TESTS
#include "SelfTestCommon.h"
#endif

#include <cmath>

namespace
{
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
[[nodiscard]] bool ShouldForceFolderViewWarpDevice() noexcept
{
    return EnvironmentVariables::IsTruthyFlagSet(L"REDSALAMANDER_FOLDERVIEW_FORCE_WARP");
}
#endif

[[nodiscard]] D2D1::ColorF D2DColorFromArgb(uint32_t argb) noexcept
{
    return D2D1::ColorF(static_cast<float>((argb >> 16u) & 0xFFu) / 255.0f,
                        static_cast<float>((argb >> 8u) & 0xFFu) / 255.0f,
                        static_cast<float>(argb & 0xFFu) / 255.0f,
                        static_cast<float>((argb >> 24u) & 0xFFu) / 255.0f);
}

[[nodiscard]] D2D1_INTERPOLATION_MODE ResolveFolderViewIconBitmapInterpolation(D2D1_SIZE_U sourcePixelSize, float destinationSizeDip, float dpi) noexcept
{
    if (sourcePixelSize.width == 0u || sourcePixelSize.height == 0u || ! (destinationSizeDip > 0.0f))
    {
        return D2D1_INTERPOLATION_MODE_LINEAR;
    }

    const float effectiveDpi = dpi > 1.0f ? dpi : static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    const float targetPixels = destinationSizeDip * effectiveDpi / static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    const float sourcePixels = static_cast<float>(std::max(sourcePixelSize.width, sourcePixelSize.height));

    constexpr float kExactPixelTolerance = 0.25f;
    if (std::abs(sourcePixels - targetPixels) <= kExactPixelTolerance)
    {
        return D2D1_INTERPOLATION_MODE_NEAREST_NEIGHBOR;
    }

    return D2D1_INTERPOLATION_MODE_LINEAR;
}
} // namespace

#ifdef ENABLE_TESTS
D2D1_INTERPOLATION_MODE DebugResolveFolderViewIconBitmapInterpolation(D2D1_SIZE_U sourcePixelSize, float destinationSizeDip, float dpi) noexcept
{
    return ResolveFolderViewIconBitmapInterpolation(sourcePixelSize, destinationSizeDip, dpi);
}
#endif

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
        const HRESULT hrFormat = RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(12.0f), _labelFormat.addressof(), L"en-us");
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
        const HRESULT hrFormat = RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(10.0f), _detailsFormat.addressof(), L"en-us");
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

    if (! _filterWatermarkFormat)
    {
        const HRESULT hrFormat = RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiIconSpec(120.0f), _filterWatermarkFormat.addressof(), L"en-us");
        if (! CheckHR(hrFormat, L"IDWriteFactory::CreateTextFormat(filter watermark)"))
        {
            return;
        }

        _filterWatermarkFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
        _filterWatermarkFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        _filterWatermarkFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
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

    bool forceWarp = false;
#ifdef ENABLE_TESTS
    forceWarp = ShouldForceFolderViewWarpDevice();
#endif
    D3D_DRIVER_TYPE createdDriverType = D3D_DRIVER_TYPE_UNKNOWN;
    const HRESULT hrDevice            = RedSalamander::DxUi::CreateD3D11DeviceWithWarpFallback(
        creationFlags, levels, forceWarp, _d3dDevice.addressof(), &_featureLevel, _d3dContext.addressof(), &createdDriverType);
    if (! CheckHR(hrDevice, createdDriverType == D3D_DRIVER_TYPE_WARP ? L"D3D11CreateDevice (WARP)" : L"D3D11CreateDevice"))
    {
        return;
    }
    if (forceWarp)
    {
        Debug::Info(L"FolderView: D3D WARP device forced by REDSALAMANDER_FOLDERVIEW_FORCE_WARP with feature level {:#06x}", static_cast<int>(_featureLevel));
    }
    else if (createdDriverType == D3D_DRIVER_TYPE_WARP)
    {
        Debug::Warning(L"FolderView: D3D hardware device creation fell back to WARP with feature level {:#06x}", static_cast<int>(_featureLevel));
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

HRESULT FolderView::CreateFolderViewSolidColorBrush(const D2D1_COLOR_F& color, wil::com_ptr<ID2D1SolidColorBrush>& brush, SolidBrushLifetime lifetime) noexcept
{
    if (! _d2dContext)
    {
        return E_UNEXPECTED;
    }

    const HRESULT hr = _d2dContext->CreateSolidColorBrush(color, brush.put());
#ifdef ENABLE_TESTS
    if (SUCCEEDED(hr) && lifetime == SolidBrushLifetime::Transient && _debugDrawItemActive)
    {
        _debugDrawItemTransientBrushCreateCount.fetch_add(1u, std::memory_order_relaxed);
    }
#else
    static_cast<void>(lifetime);
#endif
    return hr;
}

void FolderView::RecreateThemeBrushes()
{
    if (! _d2dContext)
    {
        return;
    }

    // Reset existing brushes
    _backgroundBrush.reset();
    _filterWatermarkBrush.reset();
    _backgroundWatermarkBrush.reset();
    _textBrush.reset();
    _textUnfocusedBrush.reset();
    _detailsTextBrush.reset();
    _detailsTextUnfocusedBrush.reset();
    _metadataTextBrush.reset();
    _metadataTextUnfocusedBrush.reset();
    _selectionBrush.reset();
    _hoverBrush.reset();
    _selectedItemTextBrush.reset();
    _focusedBackgroundBrush.reset();
    _focusBrush.reset();
    _emptyFolderFocusCueBrush.reset();
    _incrementalSearchHighlightBrush.reset();
    _incrementalSearchIndicatorBackgroundBrush.reset();
    _incrementalSearchIndicatorBorderBrush.reset();
    _incrementalSearchIndicatorTextBrush.reset();
    _incrementalSearchIndicatorShadowBrush.reset();
    _incrementalSearchIndicatorAccentBrush.reset();

    // Create brushes from theme colors
    const HRESULT hrBgBrush = CreateFolderViewSolidColorBrush(_theme.backgroundColor, _backgroundBrush, SolidBrushLifetime::Cached);
    if (! CheckHR(hrBgBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(background)"))
    {
        return;
    }

    const HRESULT hrTextBrush = CreateFolderViewSolidColorBrush(_theme.textNormal, _textBrush, SolidBrushLifetime::Cached);
    if (! CheckHR(hrTextBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(text)"))
    {
        return;
    }

    {
        D2D1::ColorF textColor             = _theme.textNormal;
        textColor.a                        = FolderViewVisualState::ResolveNormalTextAlpha(textColor.a, false, false);
        const HRESULT hrUnfocusedTextBrush = CreateFolderViewSolidColorBrush(textColor, _textUnfocusedBrush, SolidBrushLifetime::Cached);
        if (! CheckHR(hrUnfocusedTextBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(unfocused text)"))
        {
            return;
        }
    }

    {
        D2D1::ColorF watermarkColor = _theme.textNormal;
        watermarkColor.a            = _theme.darkBase ? 0.07f : 0.04f;
        const HRESULT hrWatermark   = CreateFolderViewSolidColorBrush(watermarkColor, _filterWatermarkBrush, SolidBrushLifetime::Cached);
        static_cast<void>(CheckHR(hrWatermark, L"ID2D1DeviceContext::CreateSolidColorBrush(filter watermark)"));
    }

    {
        D2D1::ColorF watermarkColor = _theme.textNormal;
        watermarkColor.a            = 1.0f;
        const HRESULT hrWatermark   = CreateFolderViewSolidColorBrush(watermarkColor, _backgroundWatermarkBrush, SolidBrushLifetime::Cached);
        static_cast<void>(CheckHR(hrWatermark, L"ID2D1DeviceContext::CreateSolidColorBrush(background watermark)"));
    }

    D2D1::ColorF detailsColor    = _theme.textNormal;
    detailsColor.a               = std::clamp(detailsColor.a * kDetailsTextAlpha, 0.0f, 1.0f);
    const HRESULT hrDetailsBrush = CreateFolderViewSolidColorBrush(detailsColor, _detailsTextBrush, SolidBrushLifetime::Cached);
    if (! CheckHR(hrDetailsBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(details text)"))
    {
        return;
    }

    {
        D2D1::ColorF unfocusedDetailsColor    = detailsColor;
        unfocusedDetailsColor.a               = FolderViewVisualState::ResolveNormalTextAlpha(unfocusedDetailsColor.a, false, false);
        const HRESULT hrUnfocusedDetailsBrush = CreateFolderViewSolidColorBrush(unfocusedDetailsColor, _detailsTextUnfocusedBrush, SolidBrushLifetime::Cached);
        if (! CheckHR(hrUnfocusedDetailsBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(unfocused details text)"))
        {
            return;
        }
    }

    D2D1::ColorF metadataColor    = _theme.textNormal;
    metadataColor.a               = std::clamp(metadataColor.a * kMetadataTextAlpha, 0.0f, 1.0f);
    const HRESULT hrMetadataBrush = CreateFolderViewSolidColorBrush(metadataColor, _metadataTextBrush, SolidBrushLifetime::Cached);
    if (! CheckHR(hrMetadataBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(metadata text)"))
    {
        return;
    }

    {
        D2D1::ColorF unfocusedMetadataColor = metadataColor;
        unfocusedMetadataColor.a            = FolderViewVisualState::ResolveNormalTextAlpha(unfocusedMetadataColor.a, false, false);
        const HRESULT hrUnfocusedMetadataBrush =
            CreateFolderViewSolidColorBrush(unfocusedMetadataColor, _metadataTextUnfocusedBrush, SolidBrushLifetime::Cached);
        if (! CheckHR(hrUnfocusedMetadataBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(unfocused metadata text)"))
        {
            return;
        }
    }

    const HRESULT hrSelBrush = CreateFolderViewSolidColorBrush(_theme.itemBackgroundSelected, _selectionBrush, SolidBrushLifetime::Cached);
    if (! CheckHR(hrSelBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(selection)"))
    {
        return;
    }

    const HRESULT hrHoverBrush = CreateFolderViewSolidColorBrush(_theme.itemBackgroundHovered, _hoverBrush, SolidBrushLifetime::Cached);
    if (! CheckHR(hrHoverBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(hover)"))
    {
        return;
    }

    const HRESULT hrSelectedTextBrush = CreateFolderViewSolidColorBrush(_theme.textSelected, _selectedItemTextBrush, SolidBrushLifetime::Cached);
    if (! CheckHR(hrSelectedTextBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(selected item text)"))
    {
        return;
    }

    const HRESULT hrFocusedBgBrush = CreateFolderViewSolidColorBrush(_theme.itemBackgroundFocused, _focusedBackgroundBrush, SolidBrushLifetime::Cached);
    if (! CheckHR(hrFocusedBgBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(focused background)"))
    {
        return;
    }

    const HRESULT hrFocusBrush = CreateFolderViewSolidColorBrush(_theme.focusBorder, _focusBrush, SolidBrushLifetime::Cached);
    if (! CheckHR(hrFocusBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(focus)"))
    {
        return;
    }

    const HRESULT hrIncrementalSearchBrush = CreateFolderViewSolidColorBrush(_theme.textSelected, _incrementalSearchHighlightBrush, SolidBrushLifetime::Cached);
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

    const HRESULT hrIndicatorBg = CreateFolderViewSolidColorBrush(indicatorBackground, _incrementalSearchIndicatorBackgroundBrush, SolidBrushLifetime::Cached);
    static_cast<void>(CheckHR(hrIndicatorBg, L"ID2D1DeviceContext::CreateSolidColorBrush(incremental search indicator background)"));

    const HRESULT hrIndicatorBorder = CreateFolderViewSolidColorBrush(_theme.focusBorder, _incrementalSearchIndicatorBorderBrush, SolidBrushLifetime::Cached);
    static_cast<void>(CheckHR(hrIndicatorBorder, L"ID2D1DeviceContext::CreateSolidColorBrush(incremental search indicator border)"));

    const HRESULT hrIndicatorText = CreateFolderViewSolidColorBrush(indicatorText, _incrementalSearchIndicatorTextBrush, SolidBrushLifetime::Cached);
    static_cast<void>(CheckHR(hrIndicatorText, L"ID2D1DeviceContext::CreateSolidColorBrush(incremental search indicator text)"));

    const HRESULT hrIndicatorShadow = CreateFolderViewSolidColorBrush(indicatorShadow, _incrementalSearchIndicatorShadowBrush, SolidBrushLifetime::Cached);
    static_cast<void>(CheckHR(hrIndicatorShadow, L"ID2D1DeviceContext::CreateSolidColorBrush(incremental search indicator shadow)"));

    const HRESULT hrIndicatorAccent = CreateFolderViewSolidColorBrush(_theme.focusBorder, _incrementalSearchIndicatorAccentBrush, SolidBrushLifetime::Cached);
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
        _forceFullRenderOnNextPaint = true;
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
        ReportError(L"IDXGISwapChain::ResizeBuffers", hr);
        return false;
    }

    _forceFullRenderOnNextPaint = true;
    return true;
}

void FolderView::DiscardDeviceResources()
{
    ReleaseSwapChain();

    // Clear per-item cached icons: ID2D1Bitmap1 instances are tied to the originating ID2D1Device.
    for (auto& item : _items)
    {
        item.icon.reset();
        item.thumbnail.reset();
        item.thumbnailFallbackResolved = false;
        item.thumbnailFallbackTargetPx = 0u;
    }

    wil::com_ptr<ID2D1Device> oldD2DDevice;
    {
        std::lock_guard lock(_d2dDeviceMutex);
        oldD2DDevice = _d2dDevice;
    }

    _backgroundBrush.reset();
    _filterWatermarkBrush.reset();
    _backgroundWatermarkBrush.reset();
    _textBrush.reset();
    _textUnfocusedBrush.reset();
    _selectionBrush.reset();
    _hoverBrush.reset();
    _selectedItemTextBrush.reset();
    _focusedBackgroundBrush.reset();
    _focusBrush.reset();
    _emptyFolderFocusCueBrush.reset();
    _incrementalSearchHighlightBrush.reset();
    _incrementalSearchIndicatorBackgroundBrush.reset();
    _incrementalSearchIndicatorBorderBrush.reset();
    _incrementalSearchIndicatorTextBrush.reset();
    _incrementalSearchIndicatorShadowBrush.reset();
    _incrementalSearchIndicatorAccentBrush.reset();
    _detailsTextBrush.reset();
    _detailsTextUnfocusedBrush.reset();
    _metadataTextBrush.reset();
    _metadataTextUnfocusedBrush.reset();

    _placeholderFolderIcon.reset();
    _placeholderFileIcon.reset();
    _shortcutOverlayIcon.reset();

    _labelFormat.reset();
    _detailsFormat.reset();
    _filterWatermarkFormat.reset();
    _filterWatermarkLayout.reset();
    _emptyFolderFocusCueLayout.reset();
    _emptyFolderFocusCueLayoutWidthDip  = 0.0f;
    _emptyFolderFocusCueLayoutHeightDip = 0.0f;
    _emptyFolderFocusCueLayoutDpi       = 0.0f;
    _emptyFolderFocusCueFontSizeDip     = 0.0f;
    _filterWatermarkLayoutClientSizePx  = {};
    _filterWatermarkLayoutDpi           = 0.0f;
    _filterWatermarkLayoutFontSizeDip   = 0.0f;
    _filterWatermarkBadgeLayout.reset();
    _filterWatermarkBadgeLayoutClientSizePx = {};
    _filterWatermarkBadgeLayoutDpi          = 0.0f;
    _filterWatermarkBadgeLayoutFontSizeDip  = 0.0f;
    _backgroundWatermarkLayout.reset();
    _backgroundWatermarkLayoutClientSizePx = {};
    _backgroundWatermarkLayoutDpi          = 0.0f;
    _backgroundWatermarkLayoutText.clear();
    _backgroundWatermarkLayoutFontSizeDip = 0.0f;
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
#ifdef ENABLE_TESTS
    ++_debugRenderCallCount;
#endif
    // std::wstring rectInfo = std::format(L"Rect({},{},{},{}) Items:{}", invalidRect.left, invalidRect.top, invalidRect.right, invalidRect.bottom,
    // _items.size()); TRACER_CTX(rectInfo.c_str());

    EnsureDeviceIndependentResources();
    EnsureDeviceResources();
    EnsureSwapChain();

    if (! _d2dContext || ! _d2dTarget)
    {
        Debug::Warning(L"FolderView::Render skipped - no valid render target");
        ClearPendingPaintMetricsOnFailedFrame();
        return;
    }

    Debug::Perf::Scope framePerf(L"render.frame_us");
    framePerf.SetDetail(_itemsFolder.native());

    RedSalamander::DxUi::FrameClock frameClock;
    RedSalamander::DxUi::FrameStage frameStage = RedSalamander::DxUi::FrameStage::Idle;
    const auto frameStartedAt                  = frameClock.Now();
    uint64_t folderPresentUs                   = 0u;
    bool folderPresentAttempted                = false;
    HRESULT folderFrameHr                      = S_OK;

    if (_incrementalSearchLayoutEffectsDirty && (! _incrementalSearch.active || _incrementalSearch.query.empty()))
    {
        ClearIncrementalSearchLayoutEffects();
    }

    DrawItemPerfStats drawStats{};
    uint64_t itemsConsidered          = 0;
    uint64_t itemsDrawn               = 0;
    uint64_t layoutCreates            = 0;
    uint64_t dirtyAreaPx              = 0;
    _frameTextLayoutCreateUs          = 0; // Reset per-frame DirectWrite item text-layout instrumentation (render-path scope).
    _frameTextLayoutCreateCount       = 0;
    const auto beginToEndStart        = std::chrono::steady_clock::now();
    const auto emitFolderFrameMetrics = wil::scope_exit([&]
    {
        PerfEmitDuration(L"folder.frame.total_us", frameClock.ElapsedUs(frameStartedAt, frameClock.Now()), dirtyAreaPx, itemsDrawn, folderFrameHr);
        if (folderPresentAttempted)
        {
            PerfEmitDuration(L"folder.frame.present_us", folderPresentUs, dirtyAreaPx, itemsDrawn, folderFrameHr);
        }
        PerfEmitCounter(L"folder.frame.visible_work_count", itemsDrawn);
        PerfEmitCounter(L"folder.frame.dirty_rect_area_px", dirtyAreaPx);
        if (_frameTextLayoutCreateCount > 0)
        {
            PerfEmitDuration(L"dwrite.text_layout.frame_create_us", _frameTextLayoutCreateUs, _frameTextLayoutCreateCount, 0, folderFrameHr);
        }
        PerfEmitCounter(L"dwrite.text_layout.frame_create_count", _frameTextLayoutCreateCount);
    });

    auto recoverFromDeviceLoss = [&](std::wstring_view operation, HRESULT failureHr)
    {
        DiscardDeviceResources();
        _forceFullRenderOnNextPaint = true;
#ifdef ENABLE_TESTS
        ++_debugDeviceLossRecoveryCount;
        bool d2dDeviceDiscarded = false;
        {
            std::lock_guard lock(_d2dDeviceMutex);
            d2dDeviceDiscarded = _d2dDevice == nullptr;
        }
        const bool discardedResources =
            _d3dDevice == nullptr && _d3dContext == nullptr && _d2dFactory == nullptr && d2dDeviceDiscarded && _d2dContext == nullptr && _d2dTarget == nullptr;
        if (discardedResources)
        {
            ++_debugDeviceLossDiscardedResourcesCount;
        }
        SelfTest::AppendSelfTestTrace(std::format(L"FolderView::Render: recovered from device loss operation={} hr=0x{:08X} discarded={}",
                                                  operation,
                                                  static_cast<unsigned>(failureHr),
                                                  discardedResources ? L"yes" : L"no"));
#endif
        Debug::Perf::Emit(L"folder.render.device_loss_recovery_count", operation, 0, 1u, 0, failureHr);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
    };

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
#ifdef ENABLE_TESTS
    _debugLastRenderInvalidRectPx = paintRect;
    _debugLastRenderWasFullClient = paintRect.left == 0 && paintRect.top == 0 && paintRect.right == _clientSize.cx && paintRect.bottom == _clientSize.cy;
    if (_debugLastRenderWasFullClient)
    {
        ++_debugFullClientRenderCount;
    }
#endif
    dirtyAreaPx =
        static_cast<uint64_t>(std::max<LONG>(0, paintRect.right - paintRect.left)) * static_cast<uint64_t>(std::max<LONG>(0, paintRect.bottom - paintRect.top));

    D2D1_RECT_F dirtyDip = D2D1::RectF(DipFromPx(paintRect.left), DipFromPx(paintRect.top), DipFromPx(paintRect.right), DipFromPx(paintRect.bottom));

    HRESULT hr               = S_OK;
    const uint64_t nowTickMs = GetTickCount64();
    {
        RedSalamander::DxUi::FrameStageScope renderScope(frameStage, RedSalamander::DxUi::FrameStage::Render);
        _d2dContext->BeginDraw();
        auto endDraw = wil::scope_exit([&]
        {
            hr            = _d2dContext->EndDraw();
            folderFrameHr = hr;
            PerfEmitDuration(L"render.begin_to_enddraw_us", PerfElapsedUs(beginToEndStart), dirtyAreaPx, itemsDrawn, hr);
        });
        _d2dContext->SetTransform(D2D1::Matrix3x2F::Identity());
        _d2dContext->PushAxisAlignedClip(dirtyDip, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
        _d2dContext->FillRectangle(dirtyDip, _backgroundBrush.get());

        const bool canRenderWatermark = IsNameFilterActive() && ! _appTheme.highContrast && ! _appTheme.systemHighContrast && _dwriteFactory &&
                                        _filterWatermarkFormat && _filterWatermarkBrush;
        bool drawWatermarkBadge       = false;
        if (canRenderWatermark && _items.empty() && ! _emptyStateMessage.empty())
        {
            bool hasOverlay = false;
            {
                std::lock_guard lock(_errorOverlayMutex);
                hasOverlay = _errorOverlay.has_value();
            }
            drawWatermarkBadge = ! hasOverlay && ! _pendingBusyOverlay.has_value();
        }
        else if (canRenderWatermark && _items.empty() && _emptyStateMessage.empty())
        {
            bool hasOverlay = false;
            {
                std::lock_guard lock(_errorOverlayMutex);
                hasOverlay = _errorOverlay.has_value();
            }
            drawWatermarkBadge = ! hasOverlay && ! _pendingBusyOverlay.has_value() && CanShowEmptyFolderState() && _emptyFolderState.has_value();
        }

        if (canRenderWatermark && ! drawWatermarkBadge)
        {
            const float clientWidthDip  = std::max(0.0f, DipFromPx(_clientSize.cx));
            const float clientHeightDip = std::max(0.0f, DipFromPx(_clientSize.cy));
            if (clientWidthDip > 0.0f && clientHeightDip > 0.0f)
            {
                const float minDim      = std::min(clientWidthDip, clientHeightDip);
                const float fontSizeDip = std::clamp(minDim * 0.55f, 72.0f, 240.0f);

                const wchar_t glyphText[2]{FluentIcons::kFilter, 0};
                const bool needsLayoutRebuild = ! _filterWatermarkLayout || _filterWatermarkLayoutClientSizePx.cx != _clientSize.cx ||
                                                _filterWatermarkLayoutClientSizePx.cy != _clientSize.cy || _filterWatermarkLayoutDpi != _dpi;
                if (needsLayoutRebuild)
                {
                    ++layoutCreates;
                    _filterWatermarkLayout.reset();
                    const HRESULT hrLayout = _dwriteFactory->CreateTextLayout(
                        glyphText, 1u, _filterWatermarkFormat.get(), clientWidthDip, clientHeightDip, _filterWatermarkLayout.addressof());
                    if (SUCCEEDED(hrLayout) && _filterWatermarkLayout)
                    {
                        _filterWatermarkLayoutClientSizePx = _clientSize;
                        _filterWatermarkLayoutDpi          = _dpi;
                        _filterWatermarkLayoutFontSizeDip  = 0.0f;
                    }
                }

                if (_filterWatermarkLayout)
                {
                    const DWRITE_TEXT_RANGE range{0, 1};
                    if (_filterWatermarkLayoutFontSizeDip != fontSizeDip)
                    {
                        _filterWatermarkLayout->SetFontSize(fontSizeDip, range);
                        _filterWatermarkLayoutFontSizeDip = fontSizeDip;
                    }

                    constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
#ifdef ENABLE_TESTS
                    SelfTest::AppendSelfTestTrace(std::format(
                        L"FolderView::Render: before filter watermark draw itemCount={} client={}x{}", _items.size(), _clientSize.cx, _clientSize.cy));
#endif
                    _d2dContext->DrawTextLayout(D2D1::Point2F(0.0f, 0.0f), _filterWatermarkLayout.get(), _filterWatermarkBrush.get(), options);
#ifdef ENABLE_TESTS
                    SelfTest::AppendSelfTestTrace(L"FolderView::Render: after filter watermark draw");
#endif
                }
            }
        }
        else if (canRenderWatermark && drawWatermarkBadge)
        {
            const float clientWidthDip  = std::max(0.0f, DipFromPx(_clientSize.cx));
            const float clientHeightDip = std::max(0.0f, DipFromPx(_clientSize.cy));
            if (clientWidthDip > 0.0f && clientHeightDip > 0.0f)
            {
                const float minDim          = std::min(clientWidthDip, clientHeightDip);
                const float paddingDip      = std::clamp(minDim * 0.035f, 12.0f, 26.0f);
                const float fontSizeDip     = std::clamp(minDim * 0.12f, 24.0f, 56.0f);
                const float layoutWidthDip  = std::max(1.0f, clientWidthDip - paddingDip * 2.0f);
                const float layoutHeightDip = std::max(1.0f, clientHeightDip - paddingDip * 2.0f);

                const wchar_t glyphText[2]{FluentIcons::kFilter, 0};
                const bool needsLayoutRebuild = ! _filterWatermarkBadgeLayout || _filterWatermarkBadgeLayoutClientSizePx.cx != _clientSize.cx ||
                                                _filterWatermarkBadgeLayoutClientSizePx.cy != _clientSize.cy || _filterWatermarkBadgeLayoutDpi != _dpi;
                if (needsLayoutRebuild)
                {
                    ++layoutCreates;
                    _filterWatermarkBadgeLayout.reset();
                    const HRESULT hrLayout = _dwriteFactory->CreateTextLayout(
                        glyphText, 1u, _filterWatermarkFormat.get(), layoutWidthDip, layoutHeightDip, _filterWatermarkBadgeLayout.addressof());
                    if (SUCCEEDED(hrLayout) && _filterWatermarkBadgeLayout)
                    {
                        _filterWatermarkBadgeLayoutClientSizePx = _clientSize;
                        _filterWatermarkBadgeLayoutDpi          = _dpi;
                        _filterWatermarkBadgeLayoutFontSizeDip  = 0.0f;

                        _filterWatermarkBadgeLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);
                        _filterWatermarkBadgeLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
                        _filterWatermarkBadgeLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
                    }
                }

                if (_filterWatermarkBadgeLayout)
                {
                    const DWRITE_TEXT_RANGE range{0, 1};
                    if (_filterWatermarkBadgeLayoutFontSizeDip != fontSizeDip)
                    {
                        _filterWatermarkBadgeLayout->SetFontSize(fontSizeDip, range);
                        _filterWatermarkBadgeLayoutFontSizeDip = fontSizeDip;
                    }

                    constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
                    _d2dContext->DrawTextLayout(D2D1::Point2F(paddingDip, paddingDip), _filterWatermarkBadgeLayout.get(), _filterWatermarkBrush.get(), options);
                }
            }
        }

        const bool canRenderBackgroundWatermark =
            ! _backgroundWatermarkMessage.empty() && _dwriteFactory && (_detailsFormat || _labelFormat) && _backgroundWatermarkBrush;
        if (canRenderBackgroundWatermark)
        {
            const float clientWidthDip  = std::max(1.0f, DipFromPx(_clientSize.cx));
            const float clientHeightDip = std::max(1.0f, DipFromPx(_clientSize.cy));
            const float minDimDip       = std::min(clientWidthDip, clientHeightDip);

            const float fontSizeDip = std::clamp(minDimDip * 0.07f, 18.0f, 44.0f);

            const bool needsLayoutRebuild = ! _backgroundWatermarkLayout || _backgroundWatermarkLayoutClientSizePx.cx != _clientSize.cx ||
                                            _backgroundWatermarkLayoutClientSizePx.cy != _clientSize.cy || _backgroundWatermarkLayoutDpi != _dpi ||
                                            _backgroundWatermarkLayoutText != _backgroundWatermarkMessage;
            if (needsLayoutRebuild)
            {
                ++layoutCreates;
                _backgroundWatermarkLayout.reset();

                if (_backgroundWatermarkMessage.size() <= static_cast<size_t>(std::numeric_limits<UINT32>::max()))
                {
                    const UINT32 length    = static_cast<UINT32>(_backgroundWatermarkMessage.size());
                    const HRESULT hrLayout = _dwriteFactory->CreateTextLayout(_backgroundWatermarkMessage.data(),
                                                                              length,
                                                                              _detailsFormat ? _detailsFormat.get() : _labelFormat.get(),
                                                                              clientWidthDip,
                                                                              clientHeightDip,
                                                                              _backgroundWatermarkLayout.addressof());
                    if (SUCCEEDED(hrLayout) && _backgroundWatermarkLayout)
                    {
                        _backgroundWatermarkLayoutClientSizePx = _clientSize;
                        _backgroundWatermarkLayoutDpi          = _dpi;
                        _backgroundWatermarkLayoutText         = _backgroundWatermarkMessage;
                        _backgroundWatermarkLayoutFontSizeDip  = 0.0f;

                        _backgroundWatermarkLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
                        _backgroundWatermarkLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                        _backgroundWatermarkLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
                    }
                }
            }

            if (_backgroundWatermarkLayout)
            {
                const UINT32 length = static_cast<UINT32>(_backgroundWatermarkLayoutText.size());
                if (length > 0 && _backgroundWatermarkLayoutFontSizeDip != fontSizeDip)
                {
                    const DWRITE_TEXT_RANGE range{0, length};
                    _backgroundWatermarkLayout->SetFontSize(fontSizeDip, range);
                    _backgroundWatermarkLayoutFontSizeDip = fontSizeDip;
                }

                float opacity = _theme.darkBase ? 0.08f : 0.05f;
                if (_backgroundWatermarkAnimated)
                {
                    constexpr float kPi               = 3.14159265358979323846f;
                    constexpr float kTwoPi            = 2.0f * kPi;
                    constexpr uint64_t kPulsePeriodMs = 1400u;
                    const float t                     = static_cast<float>(nowTickMs % kPulsePeriodMs) / static_cast<float>(kPulsePeriodMs);
                    const float pulse                 = 0.75f + 0.25f * std::sin(t * kTwoPi);
                    opacity                           = std::clamp(opacity * pulse, 0.0f, 1.0f);
                }
                _backgroundWatermarkBrush->SetOpacity(opacity);

                constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
                _d2dContext->DrawTextLayout(D2D1::Point2F(0.0f, 0.0f), _backgroundWatermarkLayout.get(), _backgroundWatermarkBrush.get(), options);
            }
        }

        const float layoutLeft   = dirtyDip.left + _horizontalOffset;
        const float layoutRight  = dirtyDip.right + _horizontalOffset;
        const float layoutTop    = dirtyDip.top + _scrollOffset;
        const float layoutBottom = dirtyDip.bottom + _scrollOffset;
#ifdef ENABLE_TESTS
        if (IsNameFilterActive())
        {
            SelfTest::AppendSelfTestTrace(std::format(L"FolderView::Render: before filtered item draw itemCount={} dirty=({},{})->({},{})",
                                                      _items.size(),
                                                      static_cast<int>(dirtyDip.left),
                                                      static_cast<int>(dirtyDip.top),
                                                      static_cast<int>(dirtyDip.right),
                                                      static_cast<int>(dirtyDip.bottom)));
        }
#endif

        auto drawIfVisible = [&](FolderItem& item)
        {
            ++itemsConsidered;
            const D2D1_RECT_F viewBounds = OffsetRect(item.bounds, -_horizontalOffset, -_scrollOffset);
            if (viewBounds.right < dirtyDip.left || viewBounds.left > dirtyDip.right || viewBounds.bottom < dirtyDip.top || viewBounds.top > dirtyDip.bottom)
            {
                return;
            }
            ++itemsDrawn;
            DrawItem(item, &drawStats);
        };

        if (_columnLayout.empty())
        {
            for (auto& item : _items)
            {
                drawIfVisible(item);
            }
        }
        else
        {
            const float rowSpacingDip = GetFolderViewRowSpacingDip(_appTheme);
            const float rowStride     = _tileHeightDip + rowSpacingDip;
            if (rowStride <= 0.0f)
            {
                for (auto& item : _items)
                {
                    drawIfVisible(item);
                }
            }
            else
            {
                const float firstRowTop = rowSpacingDip;
                for (const auto& column : _columnLayout)
                {
                    const int rows          = static_cast<int>(column.itemCount);
                    const float columnLeft  = column.leftDip;
                    const float columnRight = column.RightDip();
                    if (columnRight < layoutLeft || columnLeft > layoutRight)
                    {
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
                        continue;
                    }

                    size_t startIndex = column.startIndex + static_cast<size_t>(firstRow);
                    size_t endIndex   = column.startIndex + static_cast<size_t>(lastRow);
                    if (startIndex >= _items.size())
                    {
                        break;
                    }
                    endIndex = std::min(endIndex, _items.size() - 1);

                    for (size_t idx = startIndex; idx <= endIndex; ++idx)
                    {
                        drawIfVisible(_items[idx]);
                    }
                }
            }
        }

        if (_items.empty() && _displayedFolder.has_value() && _dwriteFactory && (_detailsFormat || _labelFormat))
        {
            bool hasOverlay = false;
            {
                std::lock_guard lock(_errorOverlayMutex);
                hasOverlay = _errorOverlay.has_value();
            }

            if (! hasOverlay && ! _pendingBusyOverlay.has_value())
            {
                const float clientWidthDip  = std::max(1.0f, DipFromPx(_clientSize.cx));
                const float clientHeightDip = std::max(1.0f, DipFromPx(_clientSize.cy));
                const float minDimDip       = std::min(clientWidthDip, clientHeightDip);

                auto drawCenteredText = [&](std::wstring_view message)
                {
                    if (message.empty() || message.size() > static_cast<size_t>(std::numeric_limits<UINT32>::max()))
                    {
                        return;
                    }

                    const UINT32 length = static_cast<UINT32>(message.size());
                    wil::com_ptr<IDWriteTextLayout> layout;
                    ++layoutCreates;
                    const HRESULT hrLayout = _dwriteFactory->CreateTextLayout(message.data(),
                                                                              length,
                                                                              _detailsFormat ? _detailsFormat.get() : _labelFormat.get(),
                                                                              clientWidthDip,
                                                                              clientHeightDip,
                                                                              layout.addressof());
                    if (FAILED(hrLayout) || ! layout)
                    {
                        return;
                    }

                    layout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
                    layout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                    layout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);

                    const float messageFontSizeDip = std::clamp(minDimDip * 0.085f, 18.0f, 34.0f);
                    const DWRITE_TEXT_RANGE range{0, length};
                    static_cast<void>(layout->SetFontSize(messageFontSizeDip, range));

                    ID2D1SolidColorBrush* brush = _detailsTextBrush ? _detailsTextBrush.get() : _textBrush.get();
                    if (! brush)
                    {
                        return;
                    }

                    constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
                    _d2dContext->DrawTextLayout(D2D1::Point2F(0.0f, 0.0f), layout.get(), brush, options);
                };

                auto drawEmptyFolderFocusCue = [&]()
                {
                    if (! CanShowEmptyFolderState() || ! _emptyFolderState.has_value() || ! _focusBrush)
                    {
                        return;
                    }

                    const float cueLeft   = std::max(0.0f, kColumnSpacingDip);
                    const float cueTop    = std::max(0.0f, GetFolderViewRowSpacingDip(_appTheme));
                    const float cueWidth  = std::min(_tileWidthDip, std::max(0.0f, clientWidthDip - cueLeft - kColumnSpacingDip));
                    const float cueHeight = std::min(std::max(_tileHeightDip, 0.0f), std::max(0.0f, clientHeightDip - cueTop));
                    if (cueWidth <= 0.0f || cueHeight <= 0.0f)
                    {
                        return;
                    }

                    const D2D1_RECT_F cueBounds              = D2D1::RectF(cueLeft, cueTop, cueLeft + cueWidth, cueTop + cueHeight);
                    const float maxCornerRadius              = std::min(cueWidth, cueHeight) * 0.5f;
                    const float cornerRadius                 = std::min(kSelectionCornerRadiusDip, maxCornerRadius);
                    const D2D1_ROUNDED_RECT cueRoundedBounds = D2D1::RoundedRect(cueBounds, cornerRadius, cornerRadius);

                    if (_paneFocused && _focusedBackgroundBrush)
                    {
                        _focusedBackgroundBrush->SetColor(_theme.itemBackgroundFocused);
                        _d2dContext->FillRoundedRectangle(cueRoundedBounds, _focusedBackgroundBrush.get());
                    }
                    else if (_selectionBrush)
                    {
                        _selectionBrush->SetColor(_theme.itemBackgroundSelectedInactive);
                        _d2dContext->FillRoundedRectangle(cueRoundedBounds, _selectionBrush.get());
                    }

                    const float strokeThickness = _paneFocused ? kFocusStrokeThicknessDip : kFocusStrokeThicknessUnfocusedDip;
                    const float inset           = strokeThickness * 0.5f;
                    const D2D1_RECT_F focusBounds =
                        D2D1::RectF(cueBounds.left + inset, cueBounds.top + inset, cueBounds.right - inset, cueBounds.bottom - inset);
                    const float focusWidth  = std::max(0.0f, focusBounds.right - focusBounds.left);
                    const float focusHeight = std::max(0.0f, focusBounds.bottom - focusBounds.top);
                    if (focusWidth <= 0.0f || focusHeight <= 0.0f)
                    {
                        return;
                    }

                    D2D1::ColorF focusColor = _theme.focusBorder;
                    if (! _paneFocused)
                    {
                        focusColor.a = FolderViewVisualState::ResolveFocusBorderAlpha(focusColor.a, _paneFocused);
                    }
                    _focusBrush->SetColor(focusColor);

                    const float maxFocusCornerRadius = std::min(focusWidth, focusHeight) * 0.5f;
                    const float focusCornerRadius    = std::min(std::max(0.0f, cornerRadius - inset), maxFocusCornerRadius);
                    _d2dContext->DrawRoundedRectangle(D2D1::RoundedRect(focusBounds, focusCornerRadius, focusCornerRadius), _focusBrush.get(), strokeThickness);

                    const float textPaddingXDip      = kLabelHorizontalPaddingDip;
                    const float textPaddingYDip      = kLabelVerticalPaddingDip;
                    const float layoutWidthDip       = std::max(1.0f, cueWidth - (textPaddingXDip * 2.0f));
                    const float layoutHeightDip      = std::max(1.0f, cueHeight - (textPaddingYDip * 2.0f));
                    const float cueFontSizeDip       = std::clamp(cueHeight * 0.38f, 11.0f, 15.0f);
                    const bool needsCueLayoutRebuild = ! _emptyFolderFocusCueLayout || std::fabs(_emptyFolderFocusCueLayoutWidthDip - layoutWidthDip) > 0.5f ||
                                                       std::fabs(_emptyFolderFocusCueLayoutHeightDip - layoutHeightDip) > 0.5f ||
                                                       std::fabs(_emptyFolderFocusCueFontSizeDip - cueFontSizeDip) > 0.25f ||
                                                       _emptyFolderFocusCueLayoutDpi != _dpi;
                    if (needsCueLayoutRebuild)
                    {
                        _emptyFolderFocusCueLayout.reset();

                        std::wstring cueText = LoadStringResource(nullptr, IDS_EMPTY_FOLDER_PARENT_ROW);
                        if (cueText.empty())
                        {
                            cueText = LoadStringResource(nullptr, IDS_EMPTY_FOLDER_TITLE);
                        }
                        if (cueText.empty() || cueText.size() > static_cast<size_t>(std::numeric_limits<UINT32>::max()))
                        {
                            return;
                        }

                        const UINT32 cueLength = static_cast<UINT32>(cueText.size());
                        ++layoutCreates;
                        const HRESULT hrLayout = _dwriteFactory->CreateTextLayout(cueText.data(),
                                                                                  cueLength,
                                                                                  _labelFormat ? _labelFormat.get() : _detailsFormat.get(),
                                                                                  layoutWidthDip,
                                                                                  layoutHeightDip,
                                                                                  _emptyFolderFocusCueLayout.addressof());
                        if (FAILED(hrLayout) || ! _emptyFolderFocusCueLayout)
                        {
                            return;
                        }

                        _emptyFolderFocusCueLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
                        _emptyFolderFocusCueLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                        _emptyFolderFocusCueLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
                        const DWRITE_TEXT_RANGE cueRange{0, cueLength};
                        _emptyFolderFocusCueLayout->SetFontSize(cueFontSizeDip, cueRange);
                        _emptyFolderFocusCueLayoutWidthDip  = layoutWidthDip;
                        _emptyFolderFocusCueLayoutHeightDip = layoutHeightDip;
                        _emptyFolderFocusCueFontSizeDip     = cueFontSizeDip;
                        _emptyFolderFocusCueLayoutDpi       = _dpi;
                    }

                    D2D1::ColorF cueTextColor = _theme.textNormal;
                    if (! _paneFocused)
                    {
                        cueTextColor.a = FolderViewVisualState::ResolveNormalTextAlpha(cueTextColor.a, false, false);
                    }
                    if (! _emptyFolderFocusCueBrush &&
                        FAILED(CreateFolderViewSolidColorBrush(cueTextColor, _emptyFolderFocusCueBrush, SolidBrushLifetime::Cached)))
                    {
                        return;
                    }

                    if (! _emptyFolderFocusCueBrush || ! _emptyFolderFocusCueLayout)
                    {
                        return;
                    }

                    _emptyFolderFocusCueBrush->SetColor(cueTextColor);
                    constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
                    _d2dContext->DrawTextLayout(D2D1::Point2F(cueLeft + textPaddingXDip, cueTop + textPaddingYDip),
                                                _emptyFolderFocusCueLayout.get(),
                                                _emptyFolderFocusCueBrush.get(),
                                                options);
                };

                drawEmptyFolderFocusCue();

                if (! _emptyStateMessage.empty())
                {
                    const bool canDrawNoDifferences = _emptyStateMessageKind == EmptyStateMessageKind::CompareNoDifferences && ! _appTheme.highContrast &&
                                                      ! _appTheme.systemHighContrast && _filterWatermarkBrush && (_detailsTextBrush || _textBrush) &&
                                                      _filterWatermarkFormat;
                    if (canDrawNoDifferences)
                    {
                        const UINT messageId = _compareNoDifferencesState ? _compareNoDifferencesState->funMessageResourceId : 0;
                        const bool needsLayoutRebuild =
                            ! _emptyMessageIconLayout || ! _emptyMessageTitleLayout || _emptyMessageLayoutClientSizePx.cx != _clientSize.cx ||
                            _emptyMessageLayoutClientSizePx.cy != _clientSize.cy || _emptyMessageLayoutDpi != _dpi || _emptyMessageLayoutMessageId != messageId;
                        if (needsLayoutRebuild)
                        {
                            _emptyMessageIconLayout.reset();
                            _emptyMessageTitleLayout.reset();
                            _emptyMessageFunLayout.reset();
                            _emptyMessageIconMetrics  = {};
                            _emptyMessageTitleMetrics = {};
                            _emptyMessageFunMetrics   = {};

                            const float iconFontSizeDip    = std::clamp(minDimDip * 0.30f, 64.0f, 200.0f);
                            const float titleFontSizeDip   = std::clamp(minDimDip * 0.09f, 18.0f, 32.0f);
                            const float emojiFontSizeDip   = std::clamp(minDimDip * 0.11f, 28.0f, 52.0f);
                            const float messageFontSizeDip = std::clamp(minDimDip * 0.05f, 13.0f, 18.0f);

                            const wchar_t iconText[2]{FluentIcons::kLedLight, 0};
                            ++layoutCreates;
                            const HRESULT hrIconLayout = _dwriteFactory->CreateTextLayout(
                                iconText, 1u, _filterWatermarkFormat.get(), clientWidthDip, clientHeightDip, _emptyMessageIconLayout.addressof());
                            if (SUCCEEDED(hrIconLayout) && _emptyMessageIconLayout)
                            {
                                _emptyMessageIconLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
                                _emptyMessageIconLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
                                _emptyMessageIconLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);

                                const DWRITE_TEXT_RANGE range{0, 1};
                                static_cast<void>(_emptyMessageIconLayout->SetFontSize(iconFontSizeDip, range));
                                static_cast<void>(_emptyMessageIconLayout->GetMetrics(&_emptyMessageIconMetrics));
                                _emptyMessageIconFontSizeDip = iconFontSizeDip;
                            }

                            const std::wstring_view title = _emptyStateMessage;
                            const UINT32 titleLength =
                                static_cast<UINT32>(std::min<size_t>(title.size(), static_cast<size_t>(std::numeric_limits<UINT32>::max())));
                            if (titleLength > 0)
                            {
                                const float maxTextWidthDip = std::max(1.0f, clientWidthDip - std::clamp(minDimDip * 0.18f, 80.0f, 200.0f));
                                ++layoutCreates;
                                const HRESULT hrTitleLayout = _dwriteFactory->CreateTextLayout(title.data(),
                                                                                               titleLength,
                                                                                               _labelFormat ? _labelFormat.get() : _detailsFormat.get(),
                                                                                               maxTextWidthDip,
                                                                                               clientHeightDip,
                                                                                               _emptyMessageTitleLayout.addressof());
                                if (SUCCEEDED(hrTitleLayout) && _emptyMessageTitleLayout)
                                {
                                    _emptyMessageTitleLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
                                    _emptyMessageTitleLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
                                    _emptyMessageTitleLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);

                                    const DWRITE_TEXT_RANGE range{0, titleLength};
                                    static_cast<void>(_emptyMessageTitleLayout->SetFontSize(titleFontSizeDip, range));
                                    static_cast<void>(_emptyMessageTitleLayout->SetFontWeight(DWRITE_FONT_WEIGHT_SEMI_BOLD, range));
                                    static_cast<void>(_emptyMessageTitleLayout->GetMetrics(&_emptyMessageTitleMetrics));
                                }
                            }

                            std::wstring funText;
                            if (_compareNoDifferencesState && (! _compareNoDifferencesState->emoji.empty() || ! _compareNoDifferencesState->funMessage.empty()))
                            {
                                funText = _compareNoDifferencesState->emoji;
                                if (! funText.empty() && ! _compareNoDifferencesState->funMessage.empty())
                                {
                                    funText.append(L"\r\n");
                                }
                                funText.append(_compareNoDifferencesState->funMessage);
                            }

                            const UINT32 funLength =
                                static_cast<UINT32>(std::min<size_t>(funText.size(), static_cast<size_t>(std::numeric_limits<UINT32>::max())));
                            if (funLength > 0)
                            {
                                const float maxTextWidthDip = std::max(1.0f, clientWidthDip - std::clamp(minDimDip * 0.18f, 80.0f, 200.0f));
                                ++layoutCreates;
                                const HRESULT hrFunLayout = _dwriteFactory->CreateTextLayout(funText.data(),
                                                                                             funLength,
                                                                                             _labelFormat ? _labelFormat.get() : _detailsFormat.get(),
                                                                                             maxTextWidthDip,
                                                                                             clientHeightDip,
                                                                                             _emptyMessageFunLayout.addressof());
                                if (SUCCEEDED(hrFunLayout) && _emptyMessageFunLayout)
                                {
                                    _emptyMessageFunLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
                                    _emptyMessageFunLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
                                    _emptyMessageFunLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);

                                    if (_compareNoDifferencesState && ! _compareNoDifferencesState->emoji.empty())
                                    {
                                        const auto emojiLen = static_cast<UINT32>(std::min<size_t>(_compareNoDifferencesState->emoji.size(),
                                                                                                   static_cast<size_t>(std::numeric_limits<UINT32>::max())));
                                        if (emojiLen > 0 && emojiLen <= funLength)
                                        {
                                            const DWRITE_TEXT_RANGE emojiRange{0, emojiLen};
                                            _emptyMessageFunLayout->SetFontSize(emojiFontSizeDip, emojiRange);
                                            _emptyMessageFunLayout->SetFontFamilyName(L"Segoe UI Emoji", emojiRange);
                                        }

                                        const UINT32 messageStart = emojiLen;
                                        if (messageStart < funLength)
                                        {
                                            const DWRITE_TEXT_RANGE messageRange{messageStart, funLength - messageStart};
                                            _emptyMessageFunLayout->SetFontSize(messageFontSizeDip, messageRange);
                                        }
                                    }
                                    else
                                    {
                                        const DWRITE_TEXT_RANGE range{0, funLength};
                                        _emptyMessageFunLayout->SetFontSize(messageFontSizeDip, range);
                                    }

                                    static_cast<void>(_emptyMessageFunLayout->GetMetrics(&_emptyMessageFunMetrics));
                                }
                            }

                            _emptyMessageLayoutClientSizePx = _clientSize;
                            _emptyMessageLayoutDpi          = _dpi;
                            _emptyMessageLayoutMessageId    = messageId;
                        }

                        ID2D1SolidColorBrush* textBrush = _detailsTextBrush ? _detailsTextBrush.get() : _textBrush.get();
                        constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
                        if (_emptyMessageIconLayout && _filterWatermarkBrush && textBrush)
                        {
                            const float spacingIconToTitleDip = std::clamp(minDimDip * 0.03f, 8.0f, 16.0f);
                            const float spacingTitleToFunDip  = _emptyMessageFunLayout ? std::clamp(minDimDip * 0.025f, 6.0f, 14.0f) : 0.0f;

                            const float iconHeightDip  = std::max(0.0f, _emptyMessageIconMetrics.height);
                            const float titleHeightDip = std::max(0.0f, _emptyMessageTitleMetrics.height);
                            const float funHeightDip   = _emptyMessageFunLayout ? std::max(0.0f, _emptyMessageFunMetrics.height) : 0.0f;

                            const float groupHeightDip = iconHeightDip + spacingIconToTitleDip + titleHeightDip + spacingTitleToFunDip + funHeightDip;
                            float topDip               = (clientHeightDip - groupHeightDip) * 0.5f;
                            topDip                     = std::max(0.0f, topDip);

                            float yDip = topDip;
                            _d2dContext->DrawTextLayout(D2D1::Point2F(0.0f, yDip), _emptyMessageIconLayout.get(), _filterWatermarkBrush.get(), options);

                            yDip += iconHeightDip + spacingIconToTitleDip;
                            if (_emptyMessageTitleLayout)
                            {
                                const float maxTextWidthDip = std::max(1.0f, clientWidthDip - std::clamp(minDimDip * 0.18f, 80.0f, 200.0f));
                                const float leftDip         = (clientWidthDip - maxTextWidthDip) * 0.5f;
                                _d2dContext->DrawTextLayout(D2D1::Point2F(leftDip, yDip), _emptyMessageTitleLayout.get(), textBrush, options);
                            }

                            if (_emptyMessageFunLayout)
                            {
                                yDip += titleHeightDip + spacingTitleToFunDip;
                                const float maxTextWidthDip = std::max(1.0f, clientWidthDip - std::clamp(minDimDip * 0.18f, 80.0f, 200.0f));
                                const float leftDip         = (clientWidthDip - maxTextWidthDip) * 0.5f;
                                _d2dContext->DrawTextLayout(D2D1::Point2F(leftDip, yDip), _emptyMessageFunLayout.get(), textBrush, options);
                            }
                        }
                        else
                        {
                            drawCenteredText(_emptyStateMessage);
                        }
                    }
                    else
                    {
                        drawCenteredText(_emptyStateMessage);
                    }
                }
                else if (CanShowEmptyFolderState() && _emptyFolderState.has_value() && _filterWatermarkBrush && (_detailsTextBrush || _textBrush) &&
                         _filterWatermarkFormat)
                {
                    const UINT messageId          = _emptyFolderState ? _emptyFolderState->funMessageResourceId : 0;
                    const bool needsLayoutRebuild = ! _emptyFolderIconLayout || ! _emptyFolderTitleLayout || ! _emptyFolderFunLayout ||
                                                    _emptyFolderLayoutClientSizePx.cx != _clientSize.cx ||
                                                    _emptyFolderLayoutClientSizePx.cy != _clientSize.cy || _emptyFolderLayoutDpi != _dpi ||
                                                    _emptyFolderLayoutMessageId != messageId;
                    if (needsLayoutRebuild)
                    {
                        _emptyFolderIconLayout.reset();
                        _emptyFolderTitleLayout.reset();
                        _emptyFolderFunLayout.reset();
                        _emptyFolderIconMetrics  = {};
                        _emptyFolderTitleMetrics = {};
                        _emptyFolderFunMetrics   = {};

                        const float iconFontSizeDip    = std::clamp(minDimDip * 0.35f, 72.0f, 220.0f);
                        const float titleFontSizeDip   = std::clamp(minDimDip * 0.07f, 16.0f, 24.0f);
                        const float emojiFontSizeDip   = std::clamp(minDimDip * 0.11f, 28.0f, 52.0f);
                        const float messageFontSizeDip = std::clamp(minDimDip * 0.05f, 13.0f, 18.0f);

                        const wchar_t iconText[2]{FluentIcons::kPreview, 0};
                        ++layoutCreates;
                        const HRESULT hrIconLayout = _dwriteFactory->CreateTextLayout(
                            iconText, 1u, _filterWatermarkFormat.get(), clientWidthDip, clientHeightDip, _emptyFolderIconLayout.addressof());
                        if (SUCCEEDED(hrIconLayout) && _emptyFolderIconLayout)
                        {
                            _emptyFolderIconLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
                            _emptyFolderIconLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
                            _emptyFolderIconLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);

                            const DWRITE_TEXT_RANGE range{0, 1};
                            _emptyFolderIconLayout->SetFontSize(iconFontSizeDip, range);
                            static_cast<void>(_emptyFolderIconLayout->GetMetrics(&_emptyFolderIconMetrics));
                            _emptyFolderIconFontSizeDip = iconFontSizeDip;
                        }

                        std::wstring title = LoadStringResource(nullptr, IDS_EMPTY_FOLDER_TITLE);

                        const float maxTextWidthDip = std::max(1.0f, clientWidthDip - std::clamp(minDimDip * 0.18f, 80.0f, 200.0f));
                        const UINT32 titleLength = static_cast<UINT32>(std::min<size_t>(title.size(), static_cast<size_t>(std::numeric_limits<UINT32>::max())));
                        if (titleLength > 0)
                        {
                            ++layoutCreates;
                            const HRESULT hrTitleLayout = _dwriteFactory->CreateTextLayout(title.data(),
                                                                                           titleLength,
                                                                                           _labelFormat ? _labelFormat.get() : _detailsFormat.get(),
                                                                                           maxTextWidthDip,
                                                                                           clientHeightDip,
                                                                                           _emptyFolderTitleLayout.addressof());
                            if (SUCCEEDED(hrTitleLayout) && _emptyFolderTitleLayout)
                            {
                                _emptyFolderTitleLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
                                _emptyFolderTitleLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
                                _emptyFolderTitleLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);

                                const DWRITE_TEXT_RANGE range{0, titleLength};
                                _emptyFolderTitleLayout->SetFontSize(titleFontSizeDip, range);
                                _emptyFolderTitleLayout->SetFontWeight(DWRITE_FONT_WEIGHT_SEMI_BOLD, range);
                                static_cast<void>(_emptyFolderTitleLayout->GetMetrics(&_emptyFolderTitleMetrics));
                            }
                        }

                        std::wstring funText;
                        if (_emptyFolderState && (! _emptyFolderState->emoji.empty() || ! _emptyFolderState->funMessage.empty()))
                        {
                            funText = _emptyFolderState->emoji;
                            if (! funText.empty() && ! _emptyFolderState->funMessage.empty())
                            {
                                funText.append(L"\r\n");
                            }
                            funText.append(_emptyFolderState->funMessage);
                        }

                        const UINT32 funLength = static_cast<UINT32>(std::min<size_t>(funText.size(), static_cast<size_t>(std::numeric_limits<UINT32>::max())));
                        if (funLength > 0)
                        {
                            ++layoutCreates;
                            const HRESULT hrFunLayout = _dwriteFactory->CreateTextLayout(funText.data(),
                                                                                         funLength,
                                                                                         _labelFormat ? _labelFormat.get() : _detailsFormat.get(),
                                                                                         maxTextWidthDip,
                                                                                         clientHeightDip,
                                                                                         _emptyFolderFunLayout.addressof());
                            if (SUCCEEDED(hrFunLayout) && _emptyFolderFunLayout)
                            {
                                _emptyFolderFunLayout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
                                _emptyFolderFunLayout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
                                _emptyFolderFunLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);

                                if (_emptyFolderState && ! _emptyFolderState->emoji.empty())
                                {
                                    const auto emojiLen = static_cast<UINT32>(
                                        std::min<size_t>(_emptyFolderState->emoji.size(), static_cast<size_t>(std::numeric_limits<UINT32>::max())));
                                    if (emojiLen > 0 && emojiLen <= funLength)
                                    {
                                        const DWRITE_TEXT_RANGE emojiRange{0, emojiLen};
                                        _emptyFolderFunLayout->SetFontSize(emojiFontSizeDip, emojiRange);
                                        _emptyFolderFunLayout->SetFontFamilyName(L"Segoe UI Emoji", emojiRange);
                                    }

                                    const UINT32 messageStart = emojiLen;
                                    if (messageStart < funLength)
                                    {
                                        const DWRITE_TEXT_RANGE messageRange{messageStart, funLength - messageStart};
                                        _emptyFolderFunLayout->SetFontSize(messageFontSizeDip, messageRange);
                                    }
                                }
                                else
                                {
                                    const DWRITE_TEXT_RANGE range{0, funLength};
                                    _emptyFolderFunLayout->SetFontSize(messageFontSizeDip, range);
                                }

                                static_cast<void>(_emptyFolderFunLayout->GetMetrics(&_emptyFolderFunMetrics));
                            }
                        }

                        _emptyFolderLayoutClientSizePx = _clientSize;
                        _emptyFolderLayoutDpi          = _dpi;
                        _emptyFolderLayoutMessageId    = messageId;
                    }

                    ID2D1SolidColorBrush* textBrush = _detailsTextBrush ? _detailsTextBrush.get() : _textBrush.get();
                    constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
                    if (_emptyFolderIconLayout && _filterWatermarkBrush && textBrush)
                    {
                        const float spacingIconToTitleDip = std::clamp(minDimDip * 0.03f, 8.0f, 16.0f);
                        const float spacingTitleToFunDip  = std::clamp(minDimDip * 0.025f, 6.0f, 14.0f);

                        const float iconHeightDip  = std::max(0.0f, _emptyFolderIconMetrics.height);
                        const float titleHeightDip = std::max(0.0f, _emptyFolderTitleMetrics.height);
                        const float funHeightDip   = std::max(0.0f, _emptyFolderFunMetrics.height);

                        const float groupHeightDip = iconHeightDip + spacingIconToTitleDip + titleHeightDip + spacingTitleToFunDip + funHeightDip;
                        float topDip               = (clientHeightDip - groupHeightDip) * 0.5f;
                        topDip                     = std::max(0.0f, topDip);

                        float yDip = topDip;
                        _d2dContext->DrawTextLayout(D2D1::Point2F(0.0f, yDip), _emptyFolderIconLayout.get(), _filterWatermarkBrush.get(), options);

                        yDip += iconHeightDip + spacingIconToTitleDip;
                        if (_emptyFolderTitleLayout)
                        {
                            const float maxTextWidthDip = std::max(1.0f, clientWidthDip - std::clamp(minDimDip * 0.18f, 80.0f, 200.0f));
                            const float leftDip         = (clientWidthDip - maxTextWidthDip) * 0.5f;
                            _d2dContext->DrawTextLayout(D2D1::Point2F(leftDip, yDip), _emptyFolderTitleLayout.get(), textBrush, options);
                        }

                        yDip += titleHeightDip + spacingTitleToFunDip;
                        if (_emptyFolderFunLayout)
                        {
                            const float maxTextWidthDip = std::max(1.0f, clientWidthDip - std::clamp(minDimDip * 0.18f, 80.0f, 200.0f));
                            const float leftDip         = (clientWidthDip - maxTextWidthDip) * 0.5f;
                            _d2dContext->DrawTextLayout(D2D1::Point2F(leftDip, yDip), _emptyFolderFunLayout.get(), textBrush, options);
                        }
                    }
                }
            }
        }

        const bool canDrawIncrementalSearchIndicator =
            _d2dContext && _incrementalSearchIndicatorBackgroundBrush && _incrementalSearchIndicatorBorderBrush && _incrementalSearchIndicatorTextBrush &&
            _incrementalSearchIndicatorShadowBrush && _incrementalSearchIndicatorAccentBrush &&
            std::clamp(_incrementalSearchIndicatorVisibility, 0.0f, 1.0f) > 0.001f && DipFromPx(_clientSize.cx) > 0.0f && DipFromPx(_clientSize.cy) > 0.0f;
        bool hasErrorOverlay = false;
        {
            std::lock_guard lock(_errorOverlayMutex);
            hasErrorOverlay = _errorOverlay.has_value();
        }
        const bool canDrawErrorOverlay =
            hasErrorOverlay && _d2dContext && _dwriteFactory && _alertOverlay && DipFromPx(_clientSize.cx) > 0.0f && DipFromPx(_clientSize.cy) > 0.0f;

        DrawIncrementalSearchIndicator(nowTickMs);
        DrawErrorOverlay();
        if (canDrawIncrementalSearchIndicator || canDrawErrorOverlay)
        {
            PerfEmitCounter(L"folder.frame.overlay_dirty_rect_area_px", dirtyAreaPx);
        }

#ifdef ENABLE_TESTS
        if (IsNameFilterActive())
        {
            SelfTest::AppendSelfTestTrace(L"FolderView::Render: before PopAxisAlignedClip");
        }
#endif
        _d2dContext->PopAxisAlignedClip();
    }

    framePerf.SetValue0(dirtyAreaPx);
    framePerf.SetValue1(itemsDrawn);
    PerfEmitCounter(L"render.dirty_rect_area_px", dirtyAreaPx);
    PerfEmitCounter(L"render.items_considered", itemsConsidered);
    PerfEmitCounter(L"render.items_drawn", itemsDrawn);
    PerfEmitCounter(L"render.empty_state_layout_creates", layoutCreates);
    PerfEmitCounter(L"render.item_has_icon", drawStats.itemHasIcon);
    PerfEmitCounter(L"render.item_placeholder_icon", drawStats.itemPlaceholderIcon);
    PerfEmitCounter(L"render.item_textlayout_label", drawStats.itemTextLayoutLabel);
    PerfEmitCounter(L"render.item_textlayout_details", drawStats.itemTextLayoutDetails);
    PerfEmitCounter(L"render.item_textlayout_metadata", drawStats.itemTextLayoutMetadata);
    PerfEmitCounter(L"render.incremental_search_effect_updates", drawStats.incrementalSearchUpdates);
#ifdef ENABLE_TESTS
    _debugIncrementalSearchEffectUpdateCount += drawStats.incrementalSearchUpdates;
    if (const std::optional<HRESULT> forcedEndDraw = DebugConsumeNextRenderFailure(DebugRenderFailurePoint::EndDraw); forcedEndDraw.has_value())
    {
        hr            = forcedEndDraw.value();
        folderFrameHr = hr;
        SelfTest::AppendSelfTestTrace(std::format(L"FolderView::Render: forced EndDraw hr=0x{:08X}", static_cast<unsigned>(hr)));
    }
#endif

    if (FAILED(hr))
    {
        ClearPendingPaintMetricsOnFailedFrame();
        ReportError(L"ID2D1DeviceContext::EndDraw", hr);
        if (RedSalamander::DxUi::IsDeviceLossHResult(hr))
        {
            recoverFromDeviceLoss(L"ID2D1DeviceContext::EndDraw", hr);
            return;
        }
        ReleaseSwapChain();
        EnsureSwapChain();
        return;
    }
    else if (_supportsPresent1 && _swapChain)
    {
        DXGI_PRESENT_PARAMETERS params{};
        params.DirtyRectsCount            = 1;
        params.pDirtyRects                = &paintRect;
        params.pScrollRect                = nullptr;
        params.pScrollOffset              = nullptr;
        const auto presentStart           = std::chrono::steady_clock::now();
        const auto folderPresentStartedAt = frameClock.Now();
        HRESULT hrPresent                 = S_OK;
#ifdef ENABLE_TESTS
        const std::optional<HRESULT> forcedPresent = DebugConsumeNextRenderFailure(DebugRenderFailurePoint::Present);
        if (forcedPresent.has_value())
        {
            hrPresent = forcedPresent.value();
            SelfTest::AppendSelfTestTrace(std::format(L"FolderView::Render: forced Present1 hr=0x{:08X}", static_cast<unsigned>(hrPresent)));
        }
        else
#endif
        {
            RedSalamander::DxUi::FrameStageScope presentScope(frameStage, RedSalamander::DxUi::FrameStage::Present);
            hrPresent = _swapChain->Present1(1, 0, &params);
        }
        folderPresentUs        = frameClock.ElapsedUs(folderPresentStartedAt, frameClock.Now());
        folderPresentAttempted = true;
        folderFrameHr          = hrPresent;
        PerfEmitDuration(L"render.present_us", PerfElapsedUs(presentStart), dirtyAreaPx, itemsDrawn, hrPresent);
        if (FAILED(hrPresent))
        {
            ClearPendingPaintMetricsOnFailedFrame();
            ReportError(L"IDXGISwapChain1::Present1", hrPresent);
            if (RedSalamander::DxUi::IsDeviceLossHResult(hrPresent))
            {
                recoverFromDeviceLoss(L"IDXGISwapChain1::Present1", hrPresent);
                return;
            }
            ReleaseSwapChain();
            EnsureSwapChain();
            return;
        }
        _forceFullRenderOnNextPaint = false;
        ClearErrorOverlay(ErrorOverlayKind::Rendering);
        EmitPendingInputToPaintMetricAfterPresent();
        EmitPendingRefreshToPaintMetricAfterPresent();
    }
    else if (_swapChainLegacy)
    {
        const auto presentStart           = std::chrono::steady_clock::now();
        const auto folderPresentStartedAt = frameClock.Now();
        HRESULT hrPresent                 = S_OK;
#ifdef ENABLE_TESTS
        const std::optional<HRESULT> forcedPresent = DebugConsumeNextRenderFailure(DebugRenderFailurePoint::Present);
        if (forcedPresent.has_value())
        {
            hrPresent = forcedPresent.value();
            SelfTest::AppendSelfTestTrace(std::format(L"FolderView::Render: forced legacy Present hr=0x{:08X}", static_cast<unsigned>(hrPresent)));
        }
        else
#endif
        {
            RedSalamander::DxUi::FrameStageScope presentScope(frameStage, RedSalamander::DxUi::FrameStage::Present);
            hrPresent = _swapChainLegacy->Present(1, 0);
        }
        folderPresentUs        = frameClock.ElapsedUs(folderPresentStartedAt, frameClock.Now());
        folderPresentAttempted = true;
        folderFrameHr          = hrPresent;
        PerfEmitDuration(L"render.present_us", PerfElapsedUs(presentStart), dirtyAreaPx, itemsDrawn, hrPresent);
        if (FAILED(hrPresent))
        {
            ClearPendingPaintMetricsOnFailedFrame();
            ReportError(L"IDXGISwapChain::Present", hrPresent);
            if (RedSalamander::DxUi::IsDeviceLossHResult(hrPresent))
            {
                recoverFromDeviceLoss(L"IDXGISwapChain::Present", hrPresent);
                return;
            }
            ReleaseSwapChain();
            EnsureSwapChain();
            return;
        }
        _forceFullRenderOnNextPaint = false;
        ClearErrorOverlay(ErrorOverlayKind::Rendering);
        EmitPendingInputToPaintMetricAfterPresent();
        EmitPendingRefreshToPaintMetricAfterPresent();
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
        constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
        _d2dContext->DrawTextLayout(D2D1::Point2F(textX, textY), _incrementalSearchIndicatorLayout.get(), _incrementalSearchIndicatorTextBrush.get(), options);

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

void FolderView::DrawItem(FolderItem& item, DrawItemPerfStats* perfStats)
{
    const auto drawStart = std::chrono::steady_clock::now();
    auto emitPerf        = wil::scope_exit([&] { PerfEmitDuration(L"render.draw_item_us", PerfElapsedUs(drawStart)); });
#ifdef ENABLE_TESTS
    const bool previousDrawItemActive = _debugDrawItemActive;
    _debugDrawItemActive              = true;
    const auto restoreDrawItemActive  = wil::scope_exit([&] { _debugDrawItemActive = previousDrawItemActive; });
    if (IsNameFilterActive())
    {
        SelfTest::AppendSelfTestTrace(std::format(L"FolderView::DrawItem: begin name='{}' hasIcon={} selected={} focused={}",
                                                  item.displayName,
                                                  item.icon ? 1 : 0,
                                                  item.selected ? 1 : 0,
                                                  item.focused ? 1 : 0));
    }
#endif

    // Ensure text layout is created lazily before rendering
    const float labelWidth = GetItemTextLayoutWidth(item);
    EnsureItemTextLayout(item, labelWidth);
#ifdef ENABLE_TESTS
    if (IsNameFilterActive())
    {
        SelfTest::AppendSelfTestTrace(
            std::format(L"FolderView::DrawItem: after EnsureItemTextLayout name='{}' hasLabelLayout={} hasDetailsLayout={} hasMetadataLayout={}",
                        item.displayName,
                        item.labelLayout ? 1 : 0,
                        item.detailsLayout ? 1 : 0,
                        item.metadataLayout ? 1 : 0));
    }
#endif

    if (perfStats)
    {
        if (item.icon || (_thumbnailsVisible && item.thumbnail))
        {
            ++perfStats->itemHasIcon;
        }
        else
        {
            ++perfStats->itemPlaceholderIcon;
        }

        if (item.labelLayout)
        {
            ++perfStats->itemTextLayoutLabel;
        }
    }

    D2D1_RECT_F bounds = OffsetRect(item.bounds, -_horizontalOffset, -_scrollOffset);

    // Determine item state for color selection
    const bool isHovered =
        (_hoveredIndex != static_cast<size_t>(-1) && _hoveredIndex < _items.size() && std::addressof(item) == std::addressof(_items[_hoveredIndex]));

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
    bool selectionUsesRuntimeColor              = false;
    if (item.selected)
    {
        if (_theme.itemBackgroundSelectedDynamic.has_value())
        {
            const uint32_t argb = Common::Settings::EvaluateDynamicThemeColor(
                _theme.itemBackgroundSelectedDynamic.value(),
                Common::Settings::ThemeRuntimeContext{.seedHash32 = item.stableHash32, .highContrast = _appTheme.highContrast});
            selectionBackground = D2DColorFromArgb(argb);
            if (! selectionActive)
            {
                selectionBackground.a = std::clamp(_theme.itemBackgroundSelectedInactive.a, 0.0f, 1.0f);
            }
            selectionUsesRuntimeColor = true;
        }
        else if (_theme.itemBackgroundSelectedUsesInheritedRainbow)
        {
            const uint32_t hash = item.stableHash32;
            const float hue     = static_cast<float>(hash % 360u);
            const float sat     = 0.85f;
            const float val     = _theme.darkBase ? 0.75f : 0.90f;
            selectionBackground = ColorFromHSV(hue, sat, val);
            selectionBackground.a =
                selectionActive ? std::clamp(_theme.itemBackgroundSelected.a, 0.0f, 1.0f) : std::clamp(_theme.itemBackgroundSelectedInactive.a, 0.0f, 1.0f);
            selectionUsesRuntimeColor = true;
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
    else if (isHovered && _hoverBrush)
    {
        _hoverBrush->SetColor(_theme.itemBackgroundHovered);
        _d2dContext->FillRoundedRectangle(roundedBounds, _hoverBrush.get());
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
                    focusColor.a = FolderViewVisualState::ResolveFocusBorderAlpha(focusColor.a, _paneFocused);
                }

                _focusBrush->SetColor(focusColor);
                _d2dContext->DrawRoundedRectangle(focusRoundedBounds, _focusBrush.get(), strokeThickness);
            }
        }
    }

    const float contentTop    = bounds.top + kLabelVerticalPaddingDip;
    const float contentBottom = bounds.bottom - kLabelVerticalPaddingDip;
    const float contentHeight = std::max(0.0f, contentBottom - contentTop);
    const bool includeDetailsLine =
        _displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;
    const bool includeMetadataLine = _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;

    const float iconLeft = bounds.left + kLabelHorizontalPaddingDip;
    const float iconTop  = _displayMode == DisplayMode::Brief ? contentTop + std::max(0.0f, (contentHeight - _iconSizeDip) * 0.5f) : contentTop;
    D2D1_RECT_F iconRect = D2D1::RectF(iconLeft, iconTop, iconLeft + _iconSizeDip, iconTop + _iconSizeDip);
    float iconOpacity    = (item.fileAttributes & FILE_ATTRIBUTE_HIDDEN) != 0 ? 0.5f : 1.0f;
    if (! _paneFocused)
    {
        iconOpacity = FolderViewVisualState::ResolveNormalIconOpacity(iconOpacity, _paneFocused);
    }
    const bool drawingThumbnail = _thumbnailsVisible && item.thumbnail;
    ID2D1Bitmap1* bitmap        = drawingThumbnail ? item.thumbnail.get() : item.icon.get();
    if (bitmap)
    {
#ifdef ENABLE_TESTS
        if (IsNameFilterActive())
        {
            SelfTest::AppendSelfTestTrace(std::format(L"FolderView::DrawItem: before DrawBitmap name='{}'", item.displayName));
        }
#endif
        const D2D1_SIZE_U sourcePixelSize = bitmap->GetPixelSize();
        D2D1_RECT_F bitmapRect            = drawingThumbnail ? FitBitmapRectPreserveAspect(iconRect, sourcePixelSize) : iconRect;
#ifdef ENABLE_TESTS
        if (drawingThumbnail)
        {
            _debugLastThumbnailDrawSawThumbnail = true;
            _debugLastThumbnailSourceWidthPx    = sourcePixelSize.width;
            _debugLastThumbnailSourceHeightPx   = sourcePixelSize.height;
            _debugLastThumbnailSlotRectDip      = iconRect;
            _debugLastThumbnailDrawRectDip      = bitmapRect;
        }
        else
        {
            _debugLastIconDrawSawIcon        = true;
            _debugLastIconDrawSourceWidthPx  = sourcePixelSize.width;
            _debugLastIconDrawSourceHeightPx = sourcePixelSize.height;
            _debugLastIconDrawSlotRectDip    = iconRect;
            _debugLastIconDrawRectDip        = bitmapRect;
        }
#endif
        const D2D1_INTERPOLATION_MODE interpolationMode = ResolveFolderViewIconBitmapInterpolation(sourcePixelSize, _iconSizeDip, _dpi);
        _d2dContext->DrawBitmap(bitmap, bitmapRect, iconOpacity, interpolationMode);

        // Render shortcut overlay if applicable
        if (item.isShortcut && _shortcutOverlayIcon && ! drawingThumbnail)
        {
            // Position overlay at bottom-right corner of icon
            const float overlaySize = _iconSizeDip * 0.5f; // Half icon size for overlay
            D2D1_RECT_F overlayRect = D2D1::RectF(iconRect.right - overlaySize, iconRect.bottom - overlaySize, iconRect.right, iconRect.bottom);
            _d2dContext->DrawBitmap(_shortcutOverlayIcon.get(), &overlayRect, iconOpacity, D2D1_INTERPOLATION_MODE_LINEAR);
        }
    }
    else
    {
        // Select appropriate placeholder based on item type
        auto& placeholder = item.isDirectory ? _placeholderFolderIcon : _placeholderFileIcon;
        if (placeholder)
        {
            const float placeholderOpacity = FolderViewVisualState::ResolvePlaceholderIconOpacity(_paneFocused);
            _d2dContext->DrawBitmap(placeholder.get(), &iconRect, placeholderOpacity, D2D1_INTERPOLATION_MODE_LINEAR);
        }
        else
        {
            // Fallback if placeholders not created
            _d2dContext->FillRectangle(iconRect, _backgroundBrush.get());
            _d2dContext->DrawRectangle(iconRect, _focusBrush.get(), 1.0f);
        }
    }

    const float labelLeft                     = iconRect.right + kIconTextGapDip;
    const float labelRight                    = bounds.right - kLabelHorizontalPaddingDip;
    const float availableWidth                = std::max(0.0f, labelRight - labelLeft);
    const std::wstring_view visualDisplayName = GetVisualDisplayName(item);

    // Select text brush based on selection state
    ID2D1SolidColorBrush* textBrush = (! _paneFocused && _textUnfocusedBrush) ? _textUnfocusedBrush.get() : _textBrush.get();
    if (item.selected && _selectedItemTextBrush)
    {
        D2D1::ColorF selectedTextColor = selectionActive ? _theme.textSelected : _theme.textSelectedInactive;
        if (selectionUsesRuntimeColor)
        {
            const float luminance =
                0.2126f * selectionBackgroundForContrast.r + 0.7152f * selectionBackgroundForContrast.g + 0.0722f * selectionBackgroundForContrast.b;
            selectedTextColor = luminance > 0.60f ? D2D1::ColorF(D2D1::ColorF::Black) : D2D1::ColorF(D2D1::ColorF::White);
        }

        _selectedItemTextBrush->SetColor(selectedTextColor);
        textBrush = _selectedItemTextBrush.get();
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

        if (visualDisplayName.size() > static_cast<size_t>(std::numeric_limits<UINT32>::max()))
        {
            return;
        }

        const UINT32 textLength = static_cast<UINT32>(visualDisplayName.size());
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
    if (item.labelLayout && _incrementalSearch.active && ! _incrementalSearch.query.empty())
    {
        if (visualDisplayName.size() <= static_cast<size_t>(std::numeric_limits<UINT32>::max()))
        {
            const UINT32 textLength = static_cast<UINT32>(visualDisplayName.size());
            if (textLength > 0)
            {
                DWRITE_TEXT_RANGE clearRange{};
                clearRange.startPosition = 0;
                clearRange.length        = textLength;
                if (perfStats)
                {
                    ++perfStats->incrementalSearchUpdates;
                }
                static_cast<void>(item.labelLayout->SetDrawingEffect(nullptr, clearRange));
            }

            const std::optional<UINT32> matchOffset = FindIncrementalSearchMatchOffset(visualDisplayName);
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
                            if (perfStats)
                            {
                                ++perfStats->incrementalSearchUpdates;
                            }
                            static_cast<void>(item.labelLayout->SetDrawingEffect(_incrementalSearchHighlightBrush.get(), range));
                            _incrementalSearchLayoutEffectsDirty = true;
                        }
                    }
                }
            }
        }
    }

    if (item.labelLayout)
    {
        if (includeDetailsLine)
        {
            constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
            const float nameHeight = item.labelMetrics.height > 0.0f ? item.labelMetrics.height : std::max(0.0f, contentHeight * 0.5f);
            D2D1_POINT_2F origin{labelLeft, contentTop};
            if (incrementalSearchRange.has_value())
            {
                drawIncrementalSearchHighlight(origin, incrementalSearchRange.value());
            }
            _d2dContext->DrawTextLayout(origin, item.labelLayout.get(), textBrush, options);

            ID2D1SolidColorBrush* detailsBrush =
                item.selected ? textBrush
                              : (! _paneFocused && _detailsTextUnfocusedBrush ? _detailsTextUnfocusedBrush.get()
                                                                              : (_detailsTextBrush ? _detailsTextBrush.get() : textBrush));

            const float detailsTop = contentTop + nameHeight + kDetailsGapDip;
            if (item.detailsLayout)
            {
                if (perfStats)
                {
                    ++perfStats->itemTextLayoutDetails;
                }
                D2D1_POINT_2F detailsOrigin{labelLeft, detailsTop};
                _d2dContext->DrawTextLayout(detailsOrigin, item.detailsLayout.get(), detailsBrush, options);
            }
            else if (! item.detailsText.empty() && _detailsFormat)
            {
                if (perfStats)
                {
                    ++perfStats->itemTextLayoutDetails;
                }
                D2D1_RECT_F detailsRect = D2D1::RectF(labelLeft, detailsTop, labelLeft + availableWidth, contentBottom);
                _d2dContext->DrawTextW(
                    item.detailsText.c_str(), static_cast<UINT32>(item.detailsText.length()), _detailsFormat.get(), detailsRect, detailsBrush, options);
            }

            if (includeMetadataLine)
            {
                const bool hasDetails = item.detailsLayout || (! item.detailsText.empty());
                const float detailsHeight =
                    hasDetails ? (item.detailsMetrics.height > 0.0f ? item.detailsMetrics.height : std::max(0.0f, _detailsLineHeightDip)) : 0.0f;
                const float metadataTop = hasDetails ? (detailsTop + std::max(0.0f, detailsHeight) + kDetailsGapDip) : detailsTop;

                ID2D1SolidColorBrush* metadataBrush =
                    item.selected ? textBrush
                                  : (! _paneFocused && _metadataTextUnfocusedBrush ? _metadataTextUnfocusedBrush.get()
                                                                                   : (_metadataTextBrush ? _metadataTextBrush.get() : detailsBrush));
                if (item.metadataLayout)
                {
                    if (perfStats)
                    {
                        ++perfStats->itemTextLayoutMetadata;
                    }
                    D2D1_POINT_2F metadataOrigin{labelLeft, metadataTop};
                    _d2dContext->DrawTextLayout(metadataOrigin, item.metadataLayout.get(), metadataBrush, options);
                }
                else if (! item.metadataText.empty() && _detailsFormat)
                {
                    if (perfStats)
                    {
                        ++perfStats->itemTextLayoutMetadata;
                    }
                    D2D1_RECT_F metadataRect = D2D1::RectF(labelLeft, metadataTop, labelLeft + availableWidth, contentBottom);
                    _d2dContext->DrawTextW(
                        item.metadataText.c_str(), static_cast<UINT32>(item.metadataText.length()), _detailsFormat.get(), metadataRect, metadataBrush, options);
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
            constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
            _d2dContext->DrawTextLayout(origin, item.labelLayout.get(), textBrush, options);
        }
    }
    else
    {
        if (includeDetailsLine)
        {
            const float detailsHeight  = _detailsLineHeightDip > 0.0f ? _detailsLineHeightDip : 12.0f;
            const float metadataHeight = (includeMetadataLine && _metadataLineHeightDip > 0.0f) ? _metadataLineHeightDip : 0.0f;
            const float nameBottom =
                std::max(contentTop, contentBottom - detailsHeight - kDetailsGapDip - (metadataHeight > 0.0f ? (metadataHeight + kDetailsGapDip) : 0.0f));

            D2D1_RECT_F labelRect  = D2D1::RectF(labelLeft, contentTop, labelLeft + availableWidth, nameBottom);
            constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
            _d2dContext->DrawTextW(
                visualDisplayName.data(), static_cast<UINT32>(visualDisplayName.length()), _labelFormat.get(), labelRect, textBrush, options);

            ID2D1SolidColorBrush* detailsBrush =
                item.selected ? textBrush
                              : (! _paneFocused && _detailsTextUnfocusedBrush ? _detailsTextUnfocusedBrush.get()
                                                                              : (_detailsTextBrush ? _detailsTextBrush.get() : textBrush));

            if (! item.detailsText.empty() && _detailsFormat)
            {
                D2D1_RECT_F detailsRect = D2D1::RectF(labelLeft, nameBottom + kDetailsGapDip, labelLeft + availableWidth, contentBottom);
                _d2dContext->DrawTextW(
                    item.detailsText.c_str(), static_cast<UINT32>(item.detailsText.length()), _detailsFormat.get(), detailsRect, detailsBrush, options);
            }

            if (includeMetadataLine && ! item.metadataText.empty() && _detailsFormat)
            {
                const bool hasDetails     = ! item.detailsText.empty();
                const float detailsBottom = nameBottom + kDetailsGapDip + (hasDetails ? detailsHeight : 0.0f);
                const float metadataTop   = hasDetails ? (detailsBottom + kDetailsGapDip) : detailsBottom;
                ID2D1SolidColorBrush* metadataBrush =
                    item.selected ? textBrush
                                  : (! _paneFocused && _metadataTextUnfocusedBrush ? _metadataTextUnfocusedBrush.get()
                                                                                   : (_metadataTextBrush ? _metadataTextBrush.get() : detailsBrush));
                D2D1_RECT_F metadataRect = D2D1::RectF(labelLeft, metadataTop, labelLeft + availableWidth, contentBottom);
                _d2dContext->DrawTextW(
                    item.metadataText.c_str(), static_cast<UINT32>(item.metadataText.length()), _detailsFormat.get(), metadataRect, metadataBrush, options);
            }
        }
        else
        {
            D2D1_RECT_F labelRect  = D2D1::RectF(labelLeft, contentTop, labelLeft + availableWidth, contentBottom);
            constexpr auto options = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
            _d2dContext->DrawTextW(
                visualDisplayName.data(), static_cast<UINT32>(visualDisplayName.length()), _labelFormat.get(), labelRect, textBrush, options);
        }
    }
#ifdef ENABLE_TESTS
    if (IsNameFilterActive())
    {
        SelfTest::AppendSelfTestTrace(std::format(L"FolderView::DrawItem: end name='{}'", item.displayName));
    }
#endif
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
