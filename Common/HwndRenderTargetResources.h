#pragma once

#include <algorithm>
#include <utility>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>
#include <d2d1.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#pragma warning(pop)

namespace Common::Rendering
{
// Owns only the device-independent factory and the HWND target-dependent
// resources shared by simple D2D chrome windows. Typography stays with the
// caller because font and trimming policy are surface-specific.
struct HwndRenderTargetResources
{
    wil::com_ptr<ID2D1Factory> d2dFactory;
    wil::com_ptr<ID2D1HwndRenderTarget> target;
    wil::com_ptr<ID2D1SolidColorBrush> solidBrush;

    void ResetTarget() noexcept
    {
        solidBrush.reset();
        target.reset();
    }

    [[nodiscard]] bool EnsureD2dFactory() noexcept
    {
        if (! d2dFactory)
        {
            const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, d2dFactory.addressof());
            if (FAILED(hr))
            {
                d2dFactory.reset();
                return false;
            }
        }

        return true;
    }

    [[nodiscard]] bool EnsureTarget(HWND hwnd, UINT dpi) noexcept
    {
        if (! hwnd || ! EnsureD2dFactory())
        {
            return false;
        }

        RECT client{};
        if (! GetClientRect(hwnd, &client))
        {
            return false;
        }

        const UINT width  = static_cast<UINT>(std::max(0L, client.right - client.left));
        const UINT height = static_cast<UINT>(std::max(0L, client.bottom - client.top));
        if (width == 0u || height == 0u)
        {
            return false;
        }

        if (! target)
        {
            const D2D1_RENDER_TARGET_PROPERTIES properties          = D2D1::RenderTargetProperties();
            const D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProperties = D2D1::HwndRenderTargetProperties(hwnd, D2D1::SizeU(width, height));

            wil::com_ptr<ID2D1HwndRenderTarget> createdTarget;
            const HRESULT hr = d2dFactory->CreateHwndRenderTarget(properties, hwndProperties, createdTarget.addressof());
            if (FAILED(hr) || ! createdTarget)
            {
                ResetTarget();
                return false;
            }

            createdTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
            createdTarget->SetDpi(static_cast<float>(dpi), static_cast<float>(dpi));
            target = std::move(createdTarget);
        }

        if (! solidBrush)
        {
            const HRESULT hr = target->CreateSolidColorBrush(D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f), solidBrush.addressof());
            if (FAILED(hr))
            {
                solidBrush.reset();
                return false;
            }
        }

        return target && solidBrush;
    }
};
} // namespace Common::Rendering
