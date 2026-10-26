#include "TestSupport.Tests.WindowScreenshot.h"
#include "Common.Tests.SandboxPath.h"

#ifdef ENABLE_TESTS
#pragma warning(push, 0) // Generated Windows SDK C++/WinRT headers are external implementation.
#include <d3d11.h>
#include <dwmapi.h>
#include <dxgi1_2.h>
#include <wil/com.h>
#include <wil/resource.h>
#include <wincodec.h>
#include <windows.graphics.capture.interop.h>
#include <windows.graphics.directx.direct3d11.interop.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Graphics.Capture.h>
#include <winrt/Windows.Graphics.DirectX.Direct3D11.h>
#pragma warning(pop)

#include <algorithm>
#include <cstdlib>
#include <limits>
#include <new>

// The SDK projection's factory/delegate templates emit aggregate and implicitly
// deleted-special-member warnings under /Wall. Scope this to the projection helper.
#pragma warning(push)
#pragma warning(disable : 5246 4265 4625 4626 5026 5027)

#pragma comment(lib, "windowsapp.lib")
#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "windowscodecs.lib")
#pragma comment(lib, "dwmapi.lib")

HRESULT RedSalamander::TestSupport::SaveWindowScreenshot(HWND window, const std::filesystem::path& path, std::span<WindowScreenshotColorProbe> probes) noexcept
{
    for (auto& probe : probes)
        probe.matchedPixels = 0u;
    DWORD processId          = 0;
    const DWORD windowThread = GetWindowThreadProcessId(window, &processId);
    if (processId != GetCurrentProcessId() || ! IsWindowVisible(window) || IsIconic(window))
    {
        return E_INVALIDARG;
    }
    const bool needsClientCoordinates = std::ranges::any_of(probes, [](const WindowScreenshotColorProbe& probe) noexcept { return ! probe.wholeFrame; });
    if (needsClientCoordinates && windowThread != GetCurrentThreadId())
        return E_INVALIDARG;
    std::error_code pathError;
    if (! Common::Testing::IsAuthorizedTestSandboxPath(path, {}, false, pathError))
    {
        return E_ACCESSDENIED;
    }
    // C++/WinRT projection APIs throw named HRESULT exceptions. This test boundary converts
    // them to a failed capture; it never substitutes a desktop image or a partial layer.
    try
    {
        namespace Capture  = winrt::Windows::Graphics::Capture;
        namespace Direct3D = winrt::Windows::Graphics::DirectX::Direct3D11;
        if (! Capture::GraphicsCaptureSession::IsSupported())
        {
            return E_NOTIMPL;
        }
        wil::com_ptr<ID3D11Device> device;
        wil::com_ptr<ID3D11DeviceContext> context;
        winrt::check_hresult(D3D11CreateDevice(nullptr,
                                               D3D_DRIVER_TYPE_HARDWARE,
                                               nullptr,
                                               D3D11_CREATE_DEVICE_BGRA_SUPPORT,
                                               nullptr,
                                               0u,
                                               D3D11_SDK_VERSION,
                                               device.put(),
                                               nullptr,
                                               context.put()));
        const auto dxgiDevice = device.query<IDXGIDevice>();
        wil::com_ptr<IInspectable> inspectable;
        winrt::check_hresult(CreateDirect3D11DeviceFromDXGIDevice(dxgiDevice.get(), inspectable.put()));
        Direct3D::IDirect3DDevice projected{nullptr};
        winrt::check_hresult(inspectable->QueryInterface(winrt::guid_of<Direct3D::IDirect3DDevice>(), winrt::put_abi(projected)));
        const auto interop = winrt::get_activation_factory<Capture::GraphicsCaptureItem, IGraphicsCaptureItemInterop>();
        Capture::GraphicsCaptureItem item{nullptr};
        winrt::check_hresult(interop->CreateForWindow(window, winrt::guid_of<Capture::GraphicsCaptureItem>(), winrt::put_abi(item)));
        wil::unique_event_nothrow frameArrived;
        winrt::check_hresult(frameArrived.create());
        auto pool = Capture::Direct3D11CaptureFramePool::CreateFreeThreaded(
            projected, winrt::Windows::Graphics::DirectX::DirectXPixelFormat::B8G8R8A8UIntNormalized, 2, item.Size());
        auto session = pool.CreateCaptureSession(item);
        auto close   = wil::scope_exit([&]
        {
            session.Close();
            pool.Close();
        });
        auto revoke  = pool.FrameArrived(winrt::auto_revoke, [event = frameArrived.get()](const auto&, const auto&) noexcept { SetEvent(event); });
        session.IsCursorCaptureEnabled(false);
        session.StartCapture();
        // A WGC session may initially expose stale HWND content while composition controls
        // are already current. A rejected semantic witness requests a fresh owned-window
        // paint with capture active. Bound retries; never accept an unverified partial layer.
        bool rejectedFrame = false;
        for (unsigned attempt = 0u; attempt < 3u; ++attempt)
        {
            if (attempt != 0u)
            {
                // The pool is bounded to two buffers. Drop already queued stale frames.
                for (unsigned queued = 0u; queued < 2u; ++queued)
                {
                    auto stale = pool.TryGetNextFrame();
                    if (! stale)
                        break;
                    stale.Close();
                }
                ResetEvent(frameArrived.get());
                UINT repaintFlags = RDW_INVALIDATE | RDW_ALLCHILDREN;
                if (windowThread == GetCurrentThreadId())
                    repaintFlags |= RDW_UPDATENOW;
                RedrawWindow(window, nullptr, nullptr, repaintFlags);
            }
            if (WaitForSingleObject(frameArrived.get(), 5000u) != WAIT_OBJECT_0)
            {
                return HRESULT_FROM_WIN32(rejectedFrame ? ERROR_INVALID_DATA : ERROR_TIMEOUT);
            }
            auto frame = pool.TryGetNextFrame();
            if (! frame)
                return E_UNEXPECTED;
            auto closeFrame   = wil::scope_exit([&] { frame.Close(); });
            const auto access = frame.Surface().as<::Windows::Graphics::DirectX::Direct3D11::IDirect3DDxgiInterfaceAccess>();
            wil::com_ptr<ID3D11Texture2D> texture;
            winrt::check_hresult(access->GetInterface(IID_PPV_ARGS(texture.put())));
            D3D11_TEXTURE2D_DESC description{};
            texture->GetDesc(&description);
            description.Usage          = D3D11_USAGE_STAGING;
            description.BindFlags      = 0u;
            description.MiscFlags      = 0u;
            description.CPUAccessFlags = D3D11_CPU_ACCESS_READ;
            wil::com_ptr<ID3D11Texture2D> staging;
            winrt::check_hresult(device->CreateTexture2D(&description, nullptr, staging.put()));
            context->CopyResource(staging.get(), texture.get());
            D3D11_MAPPED_SUBRESOURCE mapped{};
            winrt::check_hresult(context->Map(staging.get(), 0u, D3D11_MAP_READ, 0u, &mapped));
            auto unmap      = wil::scope_exit([&] { context->Unmap(staging.get(), 0u); });
            const auto size = frame.ContentSize();
            if (size.Width <= 0 || size.Height <= 0 || static_cast<UINT>(size.Width) > description.Width ||
                static_cast<UINT>(size.Height) > description.Height || mapped.RowPitch > std::numeric_limits<UINT>::max() / static_cast<UINT>(size.Height))
            {
                return E_UNEXPECTED;
            }
            bool probesMatched = true;
            if (! probes.empty())
            {
                RECT frameBounds{};
                POINT clientOrigin{};
                RECT client{};
                if (needsClientCoordinates)
                {
                    winrt::check_hresult(DwmGetWindowAttribute(window, DWMWA_EXTENDED_FRAME_BOUNDS, &frameBounds, sizeof(frameBounds)));
                    if (! ClientToScreen(window, &clientOrigin) || ! GetClientRect(window, &client))
                        return HRESULT_FROM_WIN32(GetLastError());
                }
                for (auto& probe : probes)
                {
                    probe.matchedPixels = 0u;
                    const RECT bounds   = probe.wholeFrame ? RECT{0, 0, size.Width, size.Height} : probe.clientBounds;
                    const RECT limit    = probe.wholeFrame ? RECT{0, 0, size.Width, size.Height} : client;
                    const POINT offset  = probe.wholeFrame ? POINT{} : POINT{clientOrigin.x - frameBounds.left, clientOrigin.y - frameBounds.top};
                    if (bounds.left < 0 || bounds.top < 0 || bounds.right > limit.right || bounds.bottom > limit.bottom || bounds.right <= bounds.left ||
                        bounds.bottom <= bounds.top || probe.minimumMatchingPixels == 0u)
                        return E_INVALIDARG;
                    const LONG left   = bounds.left + offset.x;
                    const LONG top    = bounds.top + offset.y;
                    const LONG right  = bounds.right + offset.x;
                    const LONG bottom = bounds.bottom + offset.y;
                    if (left < 0 || top < 0 || right > size.Width || bottom > size.Height)
                        return E_INVALIDARG;
                    const auto closeChannel = [&](BYTE actual, BYTE expected) noexcept
                    { return std::abs(static_cast<int>(actual) - static_cast<int>(expected)) <= probe.tolerance; };
                    for (LONG y = top; y < bottom; ++y)
                    {
                        const auto* row = static_cast<const BYTE*>(mapped.pData) + static_cast<size_t>(y) * mapped.RowPitch;
                        for (LONG x = left; x < right; ++x)
                        {
                            const auto* pixel = row + static_cast<size_t>(x) * 4u;
                            if (closeChannel(pixel[0], GetBValue(probe.expectedColor)) && closeChannel(pixel[1], GetGValue(probe.expectedColor)) &&
                                closeChannel(pixel[2], GetRValue(probe.expectedColor)))
                                ++probe.matchedPixels;
                        }
                    }
                    probesMatched = probesMatched && probe.matchedPixels >= probe.minimumMatchingPixels;
                }
            }
            wil::com_ptr<IWICImagingFactory> factory;
            winrt::check_hresult(CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.put())));
            wil::com_ptr<IWICStream> stream;
            winrt::check_hresult(factory->CreateStream(stream.put()));
            winrt::check_hresult(stream->InitializeFromFilename(path.c_str(), GENERIC_WRITE));
            wil::com_ptr<IWICBitmapEncoder> encoder;
            winrt::check_hresult(factory->CreateEncoder(GUID_ContainerFormatPng, nullptr, encoder.put()));
            winrt::check_hresult(encoder->Initialize(stream.get(), WICBitmapEncoderNoCache));
            wil::com_ptr<IWICBitmapFrameEncode> output;
            winrt::check_hresult(encoder->CreateNewFrame(output.put(), nullptr));
            winrt::check_hresult(output->Initialize(nullptr));
            winrt::check_hresult(output->SetSize(static_cast<UINT>(size.Width), static_cast<UINT>(size.Height)));
            WICPixelFormatGUID format = GUID_WICPixelFormat32bppBGRA;
            winrt::check_hresult(output->SetPixelFormat(&format));
            if (format != GUID_WICPixelFormat32bppBGRA)
            {
                return WINCODEC_ERR_UNSUPPORTEDPIXELFORMAT;
            }
            winrt::check_hresult(output->WritePixels(
                static_cast<UINT>(size.Height), mapped.RowPitch, mapped.RowPitch * static_cast<UINT>(size.Height), static_cast<BYTE*>(mapped.pData)));
            winrt::check_hresult(output->Commit());
            winrt::check_hresult(encoder->Commit());
            if (probesMatched)
                return S_OK;
            rejectedFrame = true;
        }
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const winrt::hresult_error& error)
    {
        return error.code();
    }
}
#pragma warning(pop)
#endif
