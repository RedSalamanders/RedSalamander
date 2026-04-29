#include "Framework.h"

#include "D2DHdcPaint.h"
#include "Helpers.h"

#include <algorithm>
#include <mutex>

namespace
{
struct SharedFactoryState
{
    wil::com_ptr<ID2D1Factory> factory;
    DWORD ownerThreadId = 0;
};

[[nodiscard]] ID2D1Factory* SharedFactory() noexcept
{
    static std::mutex mutex;
    static SharedFactoryState state;

    std::lock_guard lock(mutex);
    const DWORD currentThreadId = GetCurrentThreadId();
    if (state.factory && state.ownerThreadId != currentThreadId)
    {
        _ASSERTE(state.ownerThreadId == currentThreadId);
        Debug::Error(L"D2DHdcPaint: single-threaded Direct2D factory accessed from a different thread.");
        return nullptr;
    }

    if (! state.factory)
    {
        // D2D1_FACTORY_TYPE_SINGLE_THREADED is intentional here: HDC paint sessions are a UI-thread
        // helper. Callers that need cross-thread painting must add synchronization or a separate factory.
        state.ownerThreadId = currentThreadId;
        if (FAILED(D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, state.factory.put())))
        {
            state.ownerThreadId = 0;
            state.factory.reset();
            return nullptr;
        }
    }
    return state.factory.get();
}

} // namespace

namespace D2DHdcPaint
{
D2D1_COLOR_F Color(COLORREF color, float alpha) noexcept
{
    alpha = std::clamp(alpha, 0.0f, 1.0f);
    return D2D1::ColorF(
        static_cast<float>(GetRValue(color)) / 255.0f, static_cast<float>(GetGValue(color)) / 255.0f, static_cast<float>(GetBValue(color)) / 255.0f, alpha);
}

Session::~Session() noexcept
{
    End();
}

bool Session::Begin(HDC hdc, const RECT& boundsPx) noexcept
{
    End();

    if (! hdc || boundsPx.right <= boundsPx.left || boundsPx.bottom <= boundsPx.top)
    {
        return false;
    }

    ID2D1Factory* factory = SharedFactory();
    if (! factory)
    {
        return false;
    }

    D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties(D2D1_RENDER_TARGET_TYPE_DEFAULT,
                                                                       D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE),
                                                                       static_cast<float>(USER_DEFAULT_SCREEN_DPI),
                                                                       static_cast<float>(USER_DEFAULT_SCREEN_DPI));
    if (FAILED(factory->CreateDCRenderTarget(&props, _target.put())) || ! _target)
    {
        _target.reset();
        return false;
    }

    if (FAILED(_target->BindDC(hdc, &boundsPx)))
    {
        _target.reset();
        return false;
    }

    _target->BeginDraw();
    _drawing    = true;
    _brushColor = CLR_INVALID;
    _brushAlpha = -1.0f;
    _originX    = boundsPx.left;
    _originY    = boundsPx.top;
    return true;
}

void Session::End() noexcept
{
    if (_drawing && _target)
    {
        const HRESULT hr = _target->EndDraw();
        if (FAILED(hr))
        {
            Debug::Warning(L"D2DHdcPaint: EndDraw failed while finishing HDC paint session: 0x{:08X}", hr);
        }
    }
    _drawing = false;
    _brush.reset();
    _target.reset();
    _brushColor = CLR_INVALID;
    _brushAlpha = -1.0f;
    _originX    = 0;
    _originY    = 0;
}

bool Session::SetBrush(COLORREF color, float alpha) noexcept
{
    if (! _target)
    {
        return false;
    }

    alpha = std::clamp(alpha, 0.0f, 1.0f);
    if (_brush && _brushColor == color && _brushAlpha == alpha)
    {
        return true;
    }

    _brush.reset();
    _brushColor = CLR_INVALID;
    _brushAlpha = -1.0f;
    if (FAILED(_target->CreateSolidColorBrush(Color(color, alpha), _brush.put())) || ! _brush)
    {
        return false;
    }

    _brushColor = color;
    _brushAlpha = alpha;
    return true;
}

D2D1_RECT_F Session::ToLocalRect(const RECT& rectPx) const noexcept
{
    return D2D1::RectF(static_cast<float>(rectPx.left - _originX),
                       static_cast<float>(rectPx.top - _originY),
                       static_cast<float>(rectPx.right - _originX),
                       static_cast<float>(rectPx.bottom - _originY));
}

D2D1_POINT_2F Session::ToLocalPoint(float x, float y) const noexcept
{
    return D2D1::Point2F(x - static_cast<float>(_originX), y - static_cast<float>(_originY));
}

void Session::FillRectangle(const RECT& rectPx, COLORREF color) noexcept
{
    if (! _target || rectPx.right <= rectPx.left || rectPx.bottom <= rectPx.top || ! SetBrush(color))
    {
        return;
    }

    _target->FillRectangle(ToLocalRect(rectPx), _brush.get());
}

void Session::FillRoundedRectangle(const RECT& rectPx, float radiusPx, COLORREF fill, COLORREF border, float strokeWidthPx) noexcept
{
    if (! _target || rectPx.right <= rectPx.left || rectPx.bottom <= rectPx.top)
    {
        return;
    }

    radiusPx                   = std::max(0.0f, radiusPx);
    const D2D1_RECT_F rect     = ToLocalRect(rectPx);
    const D2D1_ROUNDED_RECT rr = D2D1::RoundedRect(rect, radiusPx, radiusPx);
    if (SetBrush(fill))
    {
        _target->FillRoundedRectangle(rr, _brush.get());
    }
    if (SetBrush(border))
    {
        _target->DrawRoundedRectangle(rr, _brush.get(), std::max(1.0f, strokeWidthPx));
    }
}

void Session::DrawLine(float x1, float y1, float x2, float y2, COLORREF color, float strokeWidthPx) noexcept
{
    if (! _target || ! SetBrush(color))
    {
        return;
    }

    _target->DrawLine(ToLocalPoint(x1, y1), ToLocalPoint(x2, y2), _brush.get(), std::max(1.0f, strokeWidthPx));
}
} // namespace D2DHdcPaint
