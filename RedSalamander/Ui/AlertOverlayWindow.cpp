#include "AlertOverlayWindow.h"
#include "AnimationDispatcher.h"

#include <algorithm>
#include <cmath>

namespace RedSalamander::Ui
{
namespace
{
constexpr wchar_t kAlertOverlayWindowClassName[] = L"RedSalamander.AlertOverlayWindow";
constexpr uint64_t kShowAnimationMs              = 220;
constexpr BYTE kModelessLayerAlpha               = 245; // Slight transparency so the app remains visible below.
constexpr BYTE kModalLayerAlpha                  = 230; // More transparency for modal scrim effect.

POINT PointFromLParam(LPARAM lp) noexcept
{
    POINT pt{};
    pt.x = static_cast<LONG>(static_cast<short>(LOWORD(lp)));
    pt.y = static_cast<LONG>(static_cast<short>(HIWORD(lp)));
    return pt;
}

[[nodiscard]] bool SetWindowSubclassNoThrow(HWND hwnd, SUBCLASSPROC proc, UINT_PTR id, DWORD_PTR refData) noexcept
{
#pragma warning(push)
    // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
#pragma warning(disable : 5039)
    const BOOL ok = SetWindowSubclass(hwnd, proc, id, refData);
#pragma warning(pop)
    return ok != 0;
}

void RemoveWindowSubclassNoThrow(HWND hwnd, SUBCLASSPROC proc, UINT_PTR id) noexcept
{
#pragma warning(push)
    // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
#pragma warning(disable : 5039)
    static_cast<void>(RemoveWindowSubclass(hwnd, proc, id));
#pragma warning(pop)
}
} // namespace

AlertOverlayWindow::~AlertOverlayWindow()
{
    Destroy();
}

void AlertOverlayWindow::SetCallbacks(AlertOverlayWindowCallbacks callbacks) noexcept
{
    _callbacks = callbacks;
}

void AlertOverlayWindow::ClearCallbacks() noexcept
{
    _callbacks = {};
}

void AlertOverlayWindow::SetKeyBindings(std::optional<uint32_t> primaryButtonId, std::optional<uint32_t> escapeButtonId) noexcept
{
    _primaryButtonId = primaryButtonId;
    _escapeButtonId  = escapeButtonId;
}

HRESULT AlertOverlayWindow::ShowForParentClient(HWND parent, const AlertTheme& theme, AlertModel model, bool blocksInput) noexcept
{
    if (! parent || ! IsWindow(parent))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
    }

    AttachToParentClient(parent);
    return TransitionVisibility(true, &theme, &model, blocksInput);
}

HRESULT AlertOverlayWindow::ShowForAnchor(HWND anchor, const AlertTheme& theme, AlertModel model, bool blocksInput) noexcept
{
    if (! anchor || ! IsWindow(anchor))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
    }

    AttachToAnchor(anchor);
    if (! _hostParent || ! IsWindow(_hostParent))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
    }

    return TransitionVisibility(true, &theme, &model, blocksInput);
}

void AlertOverlayWindow::Hide() noexcept
{
    static_cast<void>(TransitionVisibility(false, nullptr, nullptr, false));
}

LRESULT CALLBACK AlertOverlayWindow::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        auto* self     = static_cast<AlertOverlayWindow*>(cs->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    auto* self = reinterpret_cast<AlertOverlayWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self)
    {
        return self->WndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT AlertOverlayWindow::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: OnPaint(); return 0;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_MOUSEMOVE: OnMouseMove(PointFromLParam(lp)); return 0;
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_LBUTTONDOWN: OnLButtonDown(PointFromLParam(lp)); return 0;
        case WM_KEYDOWN: OnKeyDown(wp); return 0;
        case WM_SETCURSOR: return OnSetCursor(reinterpret_cast<HWND>(wp), LOWORD(lp), HIWORD(lp));
        case WM_NCDESTROY:
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            StopAnimationTimer();
            _hwnd.release();
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        default: break;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void AlertOverlayWindow::OnPaint() noexcept
{
    if (! _hwnd)
    {
        return;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(_hwnd.get(), &ps);
    static_cast<void>(hdc);

    EnsureD2DResources();
    if (! _target || ! _dwriteFactory)
    {
        return;
    }

    const float widthDip  = DipFromPx(_clientSizePx.cx);
    const float heightDip = DipFromPx(_clientSizePx.cy);
    if (widthDip <= 0.0f || heightDip <= 0.0f)
    {
        return;
    }

    const uint64_t now = GetTickCount64();

    _target->BeginDraw();
    auto endDraw = wil::scope_exit(
        [&]
        {
            const HRESULT hr = _target->EndDraw();
            if (hr == D2DERR_RECREATE_TARGET)
            {
                DiscardD2DResources();
            }
        });

    _target->Clear(D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f));
    _overlay.Draw(_target.get(), _dwriteFactory.get(), widthDip, heightDip, now);

    ApplyRegionFromOverlay();
}

void AlertOverlayWindow::OnSize(UINT width, UINT height) noexcept
{
    _clientSizePx.cx = static_cast<LONG>(width);
    _clientSizePx.cy = static_cast<LONG>(height);

    if (_target && width > 0 && height > 0)
    {
        const HRESULT hr = _target->Resize(D2D1::SizeU(width, height));
        if (FAILED(hr))
        {
            DiscardD2DResources();
        }
    }

    _panelRegionPx.reset();
}

void AlertOverlayWindow::OnMouseMove(POINT pt) noexcept
{
    if (! _visible || ! _hwnd)
    {
        return;
    }

    if (! _trackingMouseLeave)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = _hwnd.get();
        if (TrackMouseEvent(&tme) != 0)
        {
            _trackingMouseLeave = true;
        }
    }

    const float xDip = DipFromPx(pt.x);
    const float yDip = DipFromPx(pt.y);
    if (_overlay.UpdateHotState(D2D1::Point2F(xDip, yDip)))
    {
        InvalidateRect(_hwnd.get(), nullptr, FALSE);
    }
}

void AlertOverlayWindow::OnMouseLeave() noexcept
{
    if (! _visible || ! _hwnd)
    {
        return;
    }

    _trackingMouseLeave = false;
    _overlay.ClearHotState();
    InvalidateRect(_hwnd.get(), nullptr, FALSE);
}

void AlertOverlayWindow::OnLButtonDown(POINT pt) noexcept
{
    if (! _visible || ! _hwnd)
    {
        return;
    }

    const float xDip       = DipFromPx(pt.x);
    const float yDip       = DipFromPx(pt.y);
    const AlertHitTest hit = _overlay.HitTest(D2D1::Point2F(xDip, yDip));
    if (hit.part == AlertHitTest::Part::Close)
    {
        InvokeDismiss();
        return;
    }

    if (hit.part == AlertHitTest::Part::Button)
    {
        InvokeButton(hit.buttonId);
        return;
    }
}

void AlertOverlayWindow::OnKeyDown(WPARAM key) noexcept
{
    if (! _visible || ! _hwnd)
    {
        return;
    }

    if (key == VK_ESCAPE)
    {
        InvokeDismiss();
        return;
    }

    if (key == VK_TAB)
    {
        const bool reverse = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        if (_overlay.FocusNextButton(reverse))
        {
            InvalidateRect(_hwnd.get(), nullptr, FALSE);
        }
        return;
    }

    if (key == VK_LEFT || key == VK_UP || key == VK_RIGHT || key == VK_DOWN)
    {
        const bool reverse = (key == VK_LEFT || key == VK_UP);
        if (_overlay.FocusNextButton(reverse))
        {
            InvalidateRect(_hwnd.get(), nullptr, FALSE);
        }
        return;
    }

    if (key == VK_RETURN || key == VK_SPACE)
    {
        std::optional<uint32_t> buttonId = _overlay.GetFocusedButtonId();
        if (! buttonId.has_value())
        {
            buttonId = _primaryButtonId;
        }
        if (! buttonId.has_value())
        {
            const auto& buttons = _overlay.GetModel().buttons;
            for (const auto& button : buttons)
            {
                if (button.primary)
                {
                    buttonId = button.id;
                    break;
                }
            }
            if (! buttonId.has_value() && ! buttons.empty())
            {
                buttonId = buttons.front().id;
            }
        }

        if (buttonId.has_value())
        {
            InvokeButton(buttonId.value());
        }
        return;
    }
}

LRESULT AlertOverlayWindow::OnSetCursor(HWND cursorWindow, UINT hitTest, UINT mouseMsg) noexcept
{
    static_cast<void>(cursorWindow);
    static_cast<void>(mouseMsg);

    if (! _hwnd)
    {
        return FALSE;
    }

    if (! _visible)
    {
        return DefWindowProcW(_hwnd.get(), WM_SETCURSOR, reinterpret_cast<WPARAM>(cursorWindow), MAKELPARAM(hitTest, mouseMsg));
    }

    if (hitTest == HTCLIENT)
    {
        POINT pt{};
        if (GetCursorPos(&pt) != 0)
        {
            ScreenToClient(_hwnd.get(), &pt);
            const float xDip       = DipFromPx(pt.x);
            const float yDip       = DipFromPx(pt.y);
            const AlertHitTest hit = _overlay.HitTest(D2D1::Point2F(xDip, yDip));
            if (hit.part == AlertHitTest::Part::Close || hit.part == AlertHitTest::Part::Button)
            {
                SetCursor(LoadCursorW(nullptr, IDC_HAND));
                return TRUE;
            }
        }
    }

    SetCursor(LoadCursorW(nullptr, IDC_ARROW));
    return TRUE;
}

void AlertOverlayWindow::InvokeButton(uint32_t buttonId) noexcept
{
    if (_callbacks.onButton)
    {
        _callbacks.onButton(_callbacks.context, buttonId);
        return;
    }

    Hide();
}

void AlertOverlayWindow::InvokeDismiss() noexcept
{
    if (_escapeButtonId.has_value())
    {
        InvokeButton(_escapeButtonId.value());
        return;
    }

    if (! _overlay.GetModel().closable)
    {
        return;
    }

    if (_callbacks.onDismissed)
    {
        _callbacks.onDismissed(_callbacks.context);
    }

    Hide();
}

void AlertOverlayWindow::StartAnimationTimer() noexcept
{
    if (! _hwnd || ! _visible)
    {
        StopAnimationTimer();
        return;
    }

    const uint64_t now        = GetTickCount64();
    const bool needsAnimation = _alwaysAnimate || now < _animateUntilTickMs;
    if (! needsAnimation)
    {
        StopAnimationTimer();
        return;
    }

    if (_animationSubscriptionId == 0)
    {
        _animationSubscriptionId = AnimationDispatcher::GetInstance().Subscribe(&AlertOverlayWindow::AnimationTickThunk, this);
    }
}

void AlertOverlayWindow::StopAnimationTimer() noexcept
{
    if (_animationSubscriptionId == 0)
    {
        return;
    }

    AnimationDispatcher::GetInstance().Unsubscribe(_animationSubscriptionId);
    _animationSubscriptionId = 0;
}

bool AlertOverlayWindow::AnimationTickThunk(void* context, uint64_t nowTickMs) noexcept
{
    auto* self = static_cast<AlertOverlayWindow*>(context);
    if (! self)
    {
        return false;
    }

    return self->OnAnimationTimer(nowTickMs);
}

bool AlertOverlayWindow::OnAnimationTimer(uint64_t nowTickMs) noexcept
{
    if (! _visible || ! _hwnd)
    {
        StopAnimationTimer();
        return false;
    }

    if (! _alwaysAnimate && nowTickMs >= _animateUntilTickMs)
    {
        StopAnimationTimer();
        return false;
    }

    InvalidateRect(_hwnd.get(), nullptr, FALSE);
    return true;
}

HRESULT AlertOverlayWindow::TransitionVisibility(bool show, const AlertTheme* theme, AlertModel* model, bool blocksInput) noexcept
{
    if (! show)
    {
        StopAnimationTimer();

        _visible            = false;
        _blocksInput        = false;
        _trackingMouseLeave = false;
        _alwaysAnimate      = false;
        _animateUntilTickMs = 0;
        _startTickMs        = 0;

        _overlay.ClearHotState();
        _panelRegionPx.reset();
        ClearRegion();

        if (_hwnd)
        {
            ShowWindow(_hwnd.get(), SW_HIDE);
        }

        if (_restoreFocus && IsWindow(_restoreFocus))
        {
            SetFocus(_restoreFocus);
        }
        _restoreFocus = nullptr;

        ClearCallbacks();
        _primaryButtonId.reset();
        _escapeButtonId.reset();
        return S_OK;
    }

    if (! theme || ! model)
    {
        return E_INVALIDARG;
    }

    if (! _hostParent || ! IsWindow(_hostParent))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
    }

    const HRESULT hrCreate = EnsureCreated(_hostParent);
    if (FAILED(hrCreate))
    {
        return hrCreate;
    }

    _blocksInput        = blocksInput;
    _visible            = true;
    _trackingMouseLeave = false;
    _panelRegionPx.reset();
    if (_hwnd)
    {
        const BYTE alpha = _blocksInput ? kModalLayerAlpha : kModelessLayerAlpha;
        static_cast<void>(SetLayeredWindowAttributes(_hwnd.get(), 0, alpha, LWA_ALPHA));
    }

    _overlay.SetTheme(*theme);
    _overlay.SetModel(std::move(*model));
    _overlay.ClearHotState();

    const uint64_t now  = GetTickCount64();
    _startTickMs        = now;
    _animateUntilTickMs = now + kShowAnimationMs;
    _overlay.SetStartTick(now);

    const AlertSeverity severity = _overlay.GetModel().severity;
    _alwaysAnimate               = severity == AlertSeverity::Busy;

    ClearRegion();
    UpdatePlacement();

    _restoreFocus = nullptr;
    if (_blocksInput)
    {
        _restoreFocus = GetFocus();
        ShowWindow(_hwnd.get(), SW_SHOW);
        SetFocus(_hwnd.get());
    }
    else
    {
        ShowWindow(_hwnd.get(), SW_SHOWNOACTIVATE);
    }

    StartAnimationTimer();
    InvalidateRect(_hwnd.get(), nullptr, FALSE);
    return S_OK;
}

HRESULT AlertOverlayWindow::EnsureCreated(HWND hostParent) noexcept
{
    if (_hwnd && IsWindow(_hwnd.get()))
    {
        return S_OK;
    }

    static ATOM atom = 0;
    if (atom == 0)
    {
        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.style         = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc   = &AlertOverlayWindow::WndProcThunk;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kAlertOverlayWindowClassName;
        atom             = RegisterClassExW(&wc);
    }

    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    RECT rc{};
    GetClientRect(hostParent, &rc);
    const int width  = std::max(0L, rc.right - rc.left);
    const int height = std::max(0L, rc.bottom - rc.top);

    const DWORD style   = WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN;
    const DWORD exStyle = WS_EX_LAYERED;
    HWND hwnd = CreateWindowExW(exStyle, kAlertOverlayWindowClassName, L"", style, 0, 0, width, height, hostParent, nullptr, GetModuleHandleW(nullptr), this);
    if (! hwnd)
    {
        const DWORD lastError = GetLastError();
        return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    _hwnd.reset(hwnd);

    RECT clientRect{};
    GetClientRect(hwnd, &clientRect);
    _clientSizePx.cx = std::max(0L, clientRect.right - clientRect.left);
    _clientSizePx.cy = std::max(0L, clientRect.bottom - clientRect.top);
    _dpi             = GetDpiForWindow(hwnd);
    return S_OK;
}

void AlertOverlayWindow::Destroy() noexcept
{
    Hide();
    ApplyAttachmentState(nullptr, nullptr, false, false);
    _d2dFactory.reset();
    _dwriteFactory.reset();
    DiscardD2DResources();

    _hwnd.reset();
}

void AlertOverlayWindow::AttachToParentClient(HWND parent) noexcept
{
    ApplyAttachmentState(parent, nullptr, true, false);
}

void AlertOverlayWindow::AttachToAnchor(HWND anchor) noexcept
{
    HWND hostParent = nullptr;
    if (anchor && IsWindow(anchor))
    {
        hostParent = GetParent(anchor);
        if (! hostParent || ! IsWindow(hostParent))
        {
            hostParent = anchor;
        }
    }

    ApplyAttachmentState(hostParent, anchor, false, true);
}

void AlertOverlayWindow::ApplyAttachmentState(HWND hostParent, HWND anchor, bool trackHostParent, bool trackAnchor) noexcept
{
    if (hostParent && ! IsWindow(hostParent))
    {
        hostParent = nullptr;
    }

    if (anchor && ! IsWindow(anchor))
    {
        anchor = nullptr;
    }

    if (_hostParent == hostParent && _anchor == anchor && _hostParentSubclassed == trackHostParent && _anchorSubclassed == trackAnchor)
    {
        return;
    }

    const UINT_PTR subclassId = _subclassId;
    if (_hostParentSubclassed && _hostParent && IsWindow(_hostParent))
    {
        RemoveWindowSubclassNoThrow(_hostParent, &AlertOverlayWindow::ParentSubclassProc, subclassId);
    }

    if (_anchorSubclassed && _anchor && IsWindow(_anchor))
    {
        RemoveWindowSubclassNoThrow(_anchor, &AlertOverlayWindow::AnchorSubclassProc, subclassId);
    }

    _hostParentSubclassed = false;
    _anchorSubclassed     = false;

    _hostParent = hostParent;
    _anchor     = anchor;
    _subclassId = reinterpret_cast<UINT_PTR>(this);

    if (trackHostParent && _hostParent && IsWindow(_hostParent))
    {
        if (SetWindowSubclassNoThrow(_hostParent, &AlertOverlayWindow::ParentSubclassProc, _subclassId, reinterpret_cast<DWORD_PTR>(this)))
        {
            _hostParentSubclassed = true;
        }
    }

    if (trackAnchor && _anchor && IsWindow(_anchor))
    {
        if (SetWindowSubclassNoThrow(_anchor, &AlertOverlayWindow::AnchorSubclassProc, _subclassId, reinterpret_cast<DWORD_PTR>(this)))
        {
            _anchorSubclassed = true;
        }
    }
}

void AlertOverlayWindow::UpdatePlacement() noexcept
{
    if (! _hwnd || ! _hostParent || ! IsWindow(_hostParent))
    {
        return;
    }

    RECT rc{};
    if (_anchor && _anchor != _hostParent && IsWindow(_anchor))
    {
        RECT anchorRect{};
        if (GetWindowRect(_anchor, &anchorRect) != 0)
        {
            POINT pts[2]{};
            pts[0].x = anchorRect.left;
            pts[0].y = anchorRect.top;
            pts[1].x = anchorRect.right;
            pts[1].y = anchorRect.bottom;

            MapWindowPoints(nullptr, _hostParent, pts, 2);
            rc.left   = pts[0].x;
            rc.top    = pts[0].y;
            rc.right  = pts[1].x;
            rc.bottom = pts[1].y;
        }
        else
        {
            GetClientRect(_hostParent, &rc);
        }
    }
    else
    {
        GetClientRect(_hostParent, &rc);
    }

    const int width  = std::max(0L, rc.right - rc.left);
    const int height = std::max(0L, rc.bottom - rc.top);

    UINT flags = SWP_NOACTIVATE;
    if (_visible)
    {
        flags |= SWP_SHOWWINDOW;
    }
    else
    {
        flags |= SWP_NOZORDER;
    }

    SetWindowPos(_hwnd.get(), HWND_TOP, rc.left, rc.top, width, height, flags);
}

LRESULT CALLBACK AlertOverlayWindow::ParentSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR id, DWORD_PTR refData) noexcept
{
    auto* self = reinterpret_cast<AlertOverlayWindow*>(refData);
    if (! self)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    if (msg == WM_SIZE || msg == WM_WINDOWPOSCHANGED)
    {
        self->UpdatePlacement();
    }

    if (msg == WM_NCDESTROY)
    {
        static_cast<void>(id);
        self->Destroy();
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

LRESULT CALLBACK AlertOverlayWindow::AnchorSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR id, DWORD_PTR refData) noexcept
{
    auto* self = reinterpret_cast<AlertOverlayWindow*>(refData);
    if (! self)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    if (msg == WM_SIZE || msg == WM_WINDOWPOSCHANGED)
    {
        self->UpdatePlacement();
    }

    if (msg == WM_NCDESTROY)
    {
        static_cast<void>(id);
        self->Hide();

        HWND hostParent = self->_hostParent;
        if (hostParent == hwnd)
        {
            hostParent = nullptr;
        }

        self->ApplyAttachmentState(hostParent, nullptr, false, false);
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

void AlertOverlayWindow::EnsureD2DResources() noexcept
{
    if (! _hwnd)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(_hwnd.get());
    if (dpi != 0 && dpi != _dpi)
    {
        _dpi = dpi;
        if (_target)
        {
            _target->SetDpi(static_cast<float>(_dpi), static_cast<float>(_dpi));
        }
        _panelRegionPx.reset();
    }

    if (! _d2dFactory)
    {
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, _d2dFactory.addressof());
        if (FAILED(hr))
        {
            _d2dFactory.reset();
        }
    }

    if (! _dwriteFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_dwriteFactory.addressof()));
        if (FAILED(hr))
        {
            _dwriteFactory.reset();
        }
    }

    if (! _d2dFactory || ! _dwriteFactory)
    {
        return;
    }

    if (! _target)
    {
        RECT clientRect{};
        GetClientRect(_hwnd.get(), &clientRect);
        _clientSizePx.cx = std::max(0L, clientRect.right - clientRect.left);
        _clientSizePx.cy = std::max(0L, clientRect.bottom - clientRect.top);

        D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
        props.dpiX                          = static_cast<float>(_dpi);
        props.dpiY                          = static_cast<float>(_dpi);

        const D2D1_SIZE_U size                       = D2D1::SizeU(static_cast<UINT32>(_clientSizePx.cx), static_cast<UINT32>(_clientSizePx.cy));
        D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(_hwnd.get(), size);

        wil::com_ptr<ID2D1HwndRenderTarget> target;
        if (FAILED(_d2dFactory->CreateHwndRenderTarget(props, hwndProps, target.addressof())))
        {
            return;
        }

        target->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
        _target = std::move(target);
    }
}

void AlertOverlayWindow::DiscardD2DResources() noexcept
{
    _target.reset();
    _overlay.ResetDeviceResources();
    _overlay.ResetTextResources();
}

void AlertOverlayWindow::ApplyRegionFromOverlay() noexcept
{
    if (! _hwnd)
    {
        return;
    }

    if (_blocksInput)
    {
        if (_panelRegionPx.has_value())
        {
            _panelRegionPx.reset();
            ClearRegion();
        }
        return;
    }

    if (! _overlay.HasLayout())
    {
        return;
    }

    const D2D1_RECT_F panelDip = _overlay.GetPanelRect();
    RECT panelPx{};
    panelPx.left   = PxFromDipFloor(panelDip.left);
    panelPx.top    = PxFromDipFloor(panelDip.top);
    panelPx.right  = PxFromDipCeil(panelDip.right);
    panelPx.bottom = PxFromDipCeil(panelDip.bottom);

    if (_panelRegionPx.has_value())
    {
        const RECT prev = _panelRegionPx.value();
        if (prev.left == panelPx.left && prev.top == panelPx.top && prev.right == panelPx.right && prev.bottom == panelPx.bottom)
        {
            return;
        }
    }

    const int radiusPx   = std::max(1, PxFromDipRound(12.0f));
    const int diameterPx = std::max(1, radiusPx * 2);

    wil::unique_any<HRGN, decltype(&::DeleteObject), ::DeleteObject> region;
    region.reset(CreateRoundRectRgn(panelPx.left, panelPx.top, panelPx.right, panelPx.bottom, diameterPx, diameterPx));
    if (! region)
    {
        return;
    }

    if (SetWindowRgn(_hwnd.get(), region.get(), TRUE) != 0)
    {
        region.release();
        _panelRegionPx = panelPx;
    }
}

void AlertOverlayWindow::ClearRegion() noexcept
{
    if (_hwnd)
    {
        SetWindowRgn(_hwnd.get(), nullptr, TRUE);
    }
}

float AlertOverlayWindow::DipFromPx(int px) const noexcept
{
    const float dpi = static_cast<float>(_dpi > 0 ? _dpi : 96u);
    return (static_cast<float>(px) * 96.0f) / dpi;
}

int AlertOverlayWindow::PxFromDipFloor(float dip) const noexcept
{
    const float dpi = static_cast<float>(_dpi > 0 ? _dpi : 96u);
    const float px  = (dip * dpi) / 96.0f;
    return static_cast<int>(std::floor(px));
}

int AlertOverlayWindow::PxFromDipCeil(float dip) const noexcept
{
    const float dpi = static_cast<float>(_dpi > 0 ? _dpi : 96u);
    const float px  = (dip * dpi) / 96.0f;
    return static_cast<int>(std::ceil(px));
}

int AlertOverlayWindow::PxFromDipRound(float dip) const noexcept
{
    const float dpi = static_cast<float>(_dpi > 0 ? _dpi : 96u);
    const float px  = (dip * dpi) / 96.0f;
    return static_cast<int>(std::lround(px));
}
} // namespace RedSalamander::Ui
