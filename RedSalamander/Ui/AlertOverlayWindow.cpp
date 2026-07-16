#include "AlertOverlayWindow.h"
#include "AnimationDispatcher.h"
#include "Win32CallbackHelpers.h"

#include <UIAutomation.h>
#include <algorithm>
#include <array>
#include <atomic>
#include <cmath>
#include <cwctype>
#include <format>
#include <new>
#include <oleauto.h>
#include <vector>
#include <windowsx.h>

#pragma comment(lib, "uiautomationcore.lib")

namespace RedSalamander::Ui
{
namespace
{
constexpr wchar_t kAlertOverlayWindowClassName[] = L"RedSalamander.AlertOverlayWindow";
constexpr uint64_t kShowAnimationMs              = 220;
constexpr wchar_t kParentOriginalWndProcProp[]   = L"RedSalamander.AlertOverlay.ParentOriginalWndProc";
constexpr wchar_t kParentStateProp[]             = L"RedSalamander.AlertOverlay.ParentState";
constexpr wchar_t kAnchorOriginalWndProcProp[]   = L"RedSalamander.AlertOverlay.AnchorOriginalWndProc";
constexpr wchar_t kAnchorStateProp[]             = L"RedSalamander.AlertOverlay.AnchorState";
constexpr UINT kUiaInvokeButtonMessage           = WM_APP + 0x72;
constexpr UINT kUiaInvokeDismissMessage          = WM_APP + 0x73;

POINT PointFromLParam(LPARAM lp) noexcept
{
    POINT pt{};
    pt.x = static_cast<LONG>(static_cast<short>(LOWORD(lp)));
    pt.y = static_cast<LONG>(static_cast<short>(HIWORD(lp)));
    return pt;
}

[[maybe_unused]] [[nodiscard]] WNDPROC GetStoredWndProc(HWND hwnd, const wchar_t* propName) noexcept
{
    return Win32Callback::GetStoredWndProc(hwnd, propName);
}

[[nodiscard]] bool InstallWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp, WNDPROC wndProc) noexcept
{
    return Win32Callback::InstallWndProcHook(hwnd, originalWndProcProp, wndProc);
}

void RestoreWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp, const wchar_t* stateProp) noexcept
{
    if (! hwnd)
    {
        return;
    }

    Win32Callback::RestoreWndProcHook(hwnd, originalWndProcProp);

    RemovePropW(hwnd, originalWndProcProp);
    RemovePropW(hwnd, stateProp);
}

LRESULT CallStoredWndProc(HWND hwnd, const wchar_t* originalWndProcProp, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    return Win32Callback::CallStoredWndProc(hwnd, originalWndProcProp, msg, wp, lp);
}
} // namespace

class AlertOverlayUiaProvider final : public IRawElementProviderSimple,
                                      public IRawElementProviderFragment,
                                      public IRawElementProviderFragmentRoot,
                                      public IInvokeProvider
{
public:
    enum class ElementKind : uint8_t
    {
        Root,
        Text,
        CloseButton,
        Button,
    };

    struct ElementId
    {
        ElementKind kind = ElementKind::Root;
        size_t index     = 0u;
    };

    AlertOverlayUiaProvider(HWND hwnd, ElementId id) noexcept : _hwnd(hwnd), _id(id)
    {
    }

    AlertOverlayUiaProvider(const AlertOverlayUiaProvider&)            = delete;
    AlertOverlayUiaProvider& operator=(const AlertOverlayUiaProvider&) = delete;
    AlertOverlayUiaProvider(AlertOverlayUiaProvider&&)                 = delete;
    AlertOverlayUiaProvider& operator=(AlertOverlayUiaProvider&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** object) noexcept override
    {
        if (! object)
        {
            return E_POINTER;
        }

        *object = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IRawElementProviderSimple))
        {
            *object = static_cast<IRawElementProviderSimple*>(this);
        }
        else if (riid == __uuidof(IRawElementProviderFragment))
        {
            *object = static_cast<IRawElementProviderFragment*>(this);
        }
        else if (_id.kind == ElementKind::Root && riid == __uuidof(IRawElementProviderFragmentRoot))
        {
            *object = static_cast<IRawElementProviderFragmentRoot*>(this);
        }
        else if (IsButtonLike() && riid == __uuidof(IInvokeProvider))
        {
            *object = static_cast<IInvokeProvider*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG remaining = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (remaining == 0u)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE get_ProviderOptions(ProviderOptions* options) noexcept override
    {
        if (! options)
        {
            return E_POINTER;
        }

        *options = ProviderOptions_ServerSideProvider;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetPatternProvider(PATTERNID patternId, IUnknown** provider) noexcept override
    {
        if (! provider)
        {
            return E_POINTER;
        }

        *provider = nullptr;
        if (patternId == UIA_InvokePatternId && IsButtonLike())
        {
            *provider = static_cast<IInvokeProvider*>(this);
            AddRef();
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetPropertyValue(PROPERTYID propertyId, VARIANT* value) noexcept override
    {
        if (! value)
        {
            return E_POINTER;
        }

        VariantInit(value);
        AlertOverlayWindow* window = ResolveWindow();
        if (! window)
        {
            if (propertyId == UIA_IsOffscreenPropertyId)
            {
                return SetBool(value, true);
            }
            return S_OK;
        }

        switch (propertyId)
        {
            case UIA_ControlTypePropertyId: return SetI4(value, ControlType());
            case UIA_NamePropertyId: return SetBstr(value, Name(*window));
            case UIA_AutomationIdPropertyId: return SetBstr(value, AutomationId(*window));
            case UIA_ClassNamePropertyId: return SetBstr(value, kAlertOverlayWindowClassName);
            case UIA_FrameworkIdPropertyId: return SetBstr(value, L"RedSalamander");
            case UIA_IsControlElementPropertyId: return SetBool(value, true);
            case UIA_IsContentElementPropertyId: return SetBool(value, _id.kind != ElementKind::Root);
            case UIA_IsEnabledPropertyId: return SetBool(value, true);
            case UIA_IsKeyboardFocusablePropertyId: return SetBool(value, IsButtonLike() || _id.kind == ElementKind::Root);
            case UIA_HasKeyboardFocusPropertyId: return SetBool(value, ::GetFocus() == _hwnd);
            case UIA_IsOffscreenPropertyId: return SetBool(value, IsWindowVisible(_hwnd) == FALSE);
            case UIA_NativeWindowHandlePropertyId:
                if (_id.kind == ElementKind::Root)
                {
                    return SetI4(value, static_cast<int>(reinterpret_cast<ULONG_PTR>(_hwnd) & 0x7FFFFFFFu));
                }
                return S_OK;
            default: return S_OK;
        }
    }

    HRESULT STDMETHODCALLTYPE get_HostRawElementProvider(IRawElementProviderSimple** provider) noexcept override
    {
        if (! provider)
        {
            return E_POINTER;
        }

        *provider = nullptr;
        if (_id.kind != ElementKind::Root)
        {
            return S_OK;
        }

        return UiaHostProviderFromHwnd(_hwnd, provider);
    }

    HRESULT STDMETHODCALLTYPE Navigate(NavigateDirection direction, IRawElementProviderFragment** provider) noexcept override
    {
        if (! provider)
        {
            return E_POINTER;
        }

        *provider                  = nullptr;
        AlertOverlayWindow* window = ResolveWindow();
        if (! window)
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }

        const std::vector<ElementId> children = Children(*window);
        switch (direction)
        {
            case NavigateDirection_FirstChild:
                if (_id.kind == ElementKind::Root && ! children.empty())
                {
                    return CreateFragment(_hwnd, children.front(), provider);
                }
                return S_OK;
            case NavigateDirection_LastChild:
                if (_id.kind == ElementKind::Root && ! children.empty())
                {
                    return CreateFragment(_hwnd, children.back(), provider);
                }
                return S_OK;
            case NavigateDirection_Parent:
                if (_id.kind != ElementKind::Root)
                {
                    return CreateFragment(_hwnd, ElementId{}, provider);
                }
                return S_OK;
            case NavigateDirection_NextSibling:
            case NavigateDirection_PreviousSibling:
                if (_id.kind == ElementKind::Root)
                {
                    return S_OK;
                }
                return NavigateSibling(children, direction, provider);
            default: return S_OK;
        }
    }

    HRESULT STDMETHODCALLTYPE GetRuntimeId(SAFEARRAY** runtimeId) noexcept override
    {
        if (! runtimeId)
        {
            return E_POINTER;
        }

        *runtimeId       = nullptr;
        SAFEARRAY* array = SafeArrayCreateVector(VT_I4, 0, 5u);
        if (! array)
        {
            return E_OUTOFMEMORY;
        }

        const ULONG_PTR hwndValue = reinterpret_cast<ULONG_PTR>(_hwnd);
        const LONG values[5]      = {UiaAppendRuntimeId,
                                     static_cast<LONG>(hwndValue & 0x7FFFFFFFu),
                                     static_cast<LONG>((hwndValue >> 32u) & 0x7FFFFFFFu),
                                     static_cast<LONG>(static_cast<int>(_id.kind)),
                                     static_cast<LONG>(_id.index & 0x7FFFFFFFu)};

        for (LONG i = 0; i < 5; ++i)
        {
            LONG item        = values[i];
            const HRESULT hr = SafeArrayPutElement(array, &i, &item);
            if (FAILED(hr))
            {
                SafeArrayDestroy(array);
                return hr;
            }
        }

        *runtimeId = array;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE get_BoundingRectangle(UiaRect* rect) noexcept override
    {
        if (! rect)
        {
            return E_POINTER;
        }

        RECT windowRect{};
        if (! _hwnd || GetWindowRect(_hwnd, &windowRect) == FALSE)
        {
            *rect = {};
            return UIA_E_ELEMENTNOTAVAILABLE;
        }

        rect->left   = static_cast<double>(windowRect.left);
        rect->top    = static_cast<double>(windowRect.top);
        rect->width  = static_cast<double>(std::max<LONG>(0, windowRect.right - windowRect.left));
        rect->height = static_cast<double>(std::max<LONG>(0, windowRect.bottom - windowRect.top));
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetEmbeddedFragmentRoots(SAFEARRAY** roots) noexcept override
    {
        if (! roots)
        {
            return E_POINTER;
        }

        *roots = nullptr;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetFocus() noexcept override
    {
        if (! _hwnd || IsWindow(_hwnd) == FALSE)
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }

        ::SetFocus(_hwnd);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE get_FragmentRoot(IRawElementProviderFragmentRoot** root) noexcept override
    {
        if (! root)
        {
            return E_POINTER;
        }

        *root = nullptr;
        return CreateRoot(_hwnd, root);
    }

    HRESULT STDMETHODCALLTYPE ElementProviderFromPoint(double /*x*/, double /*y*/, IRawElementProviderFragment** provider) noexcept override
    {
        if (! provider)
        {
            return E_POINTER;
        }

        *provider = nullptr;
        return CreateFragment(_hwnd, ElementId{}, provider);
    }

    HRESULT STDMETHODCALLTYPE GetFocus(IRawElementProviderFragment** provider) noexcept override
    {
        if (! provider)
        {
            return E_POINTER;
        }

        *provider                  = nullptr;
        AlertOverlayWindow* window = ResolveWindow();
        if (! window)
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }

        if (const std::optional<uint32_t> focusedButtonId = window->_overlay.GetFocusedButtonId(); focusedButtonId.has_value())
        {
            const auto& buttons = window->_overlay.GetModel().buttons;
            for (size_t index = 0; index < buttons.size(); ++index)
            {
                if (buttons[index].id == focusedButtonId.value())
                {
                    return CreateFragment(_hwnd, ElementId{ElementKind::Button, index}, provider);
                }
            }
        }

        return CreateFragment(_hwnd, ElementId{}, provider);
    }

    HRESULT STDMETHODCALLTYPE Invoke() noexcept override
    {
        AlertOverlayWindow* window = ResolveWindow();
        if (! window || ! IsButtonLike())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }

        if (_id.kind == ElementKind::CloseButton)
        {
            return PostInvokeMessage(_hwnd, kUiaInvokeDismissMessage, 0);
        }

        const auto& buttons = window->_overlay.GetModel().buttons;
        if (_id.index >= buttons.size())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }

        return PostInvokeMessage(_hwnd, kUiaInvokeButtonMessage, static_cast<WPARAM>(buttons[_id.index].id));
    }

private:
    static HRESULT PostInvokeMessage(HWND hwnd, UINT message, WPARAM wParam) noexcept
    {
        if (PostMessageW(hwnd, message, wParam, 0) != FALSE)
        {
            return S_OK;
        }

        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError == ERROR_SUCCESS ? ERROR_INVALID_HANDLE : lastError);
    }

    static HRESULT SetBool(VARIANT* value, bool data) noexcept
    {
        value->vt      = VT_BOOL;
        value->boolVal = data ? VARIANT_TRUE : VARIANT_FALSE;
        return S_OK;
    }

    static HRESULT SetI4(VARIANT* value, int data) noexcept
    {
        value->vt   = VT_I4;
        value->lVal = data;
        return S_OK;
    }

    static HRESULT SetBstr(VARIANT* value, std::wstring_view text) noexcept
    {
        BSTR bstr = SysAllocStringLen(text.data(), static_cast<UINT>(text.size()));
        if (! bstr && ! text.empty())
        {
            return E_OUTOFMEMORY;
        }

        value->vt      = VT_BSTR;
        value->bstrVal = bstr;
        return S_OK;
    }

    [[nodiscard]] bool IsButtonLike() const noexcept
    {
        return _id.kind == ElementKind::Button || _id.kind == ElementKind::CloseButton;
    }

    [[nodiscard]] AlertOverlayWindow* ResolveWindow() const noexcept
    {
        if (! _hwnd || IsWindow(_hwnd) == FALSE)
        {
            return nullptr;
        }

        auto* window = reinterpret_cast<AlertOverlayWindow*>(GetWindowLongPtrW(_hwnd, GWLP_USERDATA));
        if (! window || window->_hwnd.get() != _hwnd || ! window->_visible)
        {
            return nullptr;
        }

        return window;
    }

    [[nodiscard]] CONTROLTYPEID ControlType() const noexcept
    {
        switch (_id.kind)
        {
            case ElementKind::Root: return UIA_PaneControlTypeId;
            case ElementKind::Text: return UIA_TextControlTypeId;
            case ElementKind::CloseButton:
            case ElementKind::Button: return UIA_ButtonControlTypeId;
            default: return UIA_CustomControlTypeId;
        }
    }

    [[nodiscard]] std::wstring Name(const AlertOverlayWindow& window) const
    {
        const AlertModel& model = window._overlay.GetModel();
        switch (_id.kind)
        {
            case ElementKind::Root: return ! model.title.empty() ? model.title : L"Alert";
            case ElementKind::Text:
            {
                if (model.title.empty())
                {
                    return model.message;
                }
                if (model.message.empty())
                {
                    return model.title;
                }
                return std::format(L"{}\n{}", model.title, model.message);
            }
            case ElementKind::CloseButton: return L"Close";
            case ElementKind::Button:
                if (_id.index < model.buttons.size())
                {
                    return model.buttons[_id.index].label;
                }
                return {};
            default: return {};
        }
    }

    [[nodiscard]] std::wstring AutomationId(const AlertOverlayWindow& window) const
    {
        switch (_id.kind)
        {
            case ElementKind::Root: return L"AlertOverlay";
            case ElementKind::Text: return L"AlertOverlay.Text";
            case ElementKind::CloseButton: return L"AlertOverlay.Close";
            case ElementKind::Button:
            {
                const auto& buttons = window._overlay.GetModel().buttons;
                const uint32_t id   = _id.index < buttons.size() ? buttons[_id.index].id : 0u;
                return std::format(L"AlertOverlay.Button.{}", id);
            }
            default: return {};
        }
    }

    [[nodiscard]] static std::vector<ElementId> Children(const AlertOverlayWindow& window)
    {
        std::vector<ElementId> children;
        const AlertModel& model = window._overlay.GetModel();
        children.reserve(model.buttons.size() + 2u);
        if (! model.title.empty() || ! model.message.empty())
        {
            children.push_back(ElementId{ElementKind::Text, 0u});
        }
        if (model.closable)
        {
            children.push_back(ElementId{ElementKind::CloseButton, 0u});
        }
        for (size_t index = 0u; index < model.buttons.size(); ++index)
        {
            children.push_back(ElementId{ElementKind::Button, index});
        }
        return children;
    }

    [[nodiscard]] static bool SameElement(ElementId lhs, ElementId rhs) noexcept
    {
        return lhs.kind == rhs.kind && lhs.index == rhs.index;
    }

    static HRESULT CreateFragment(HWND hwnd, ElementId id, IRawElementProviderFragment** provider) noexcept
    {
        auto* created = new (std::nothrow) AlertOverlayUiaProvider(hwnd, id);
        if (! created)
        {
            return E_OUTOFMEMORY;
        }

        *provider = static_cast<IRawElementProviderFragment*>(created);
        return S_OK;
    }

    static HRESULT CreateRoot(HWND hwnd, IRawElementProviderFragmentRoot** root) noexcept
    {
        auto* created = new (std::nothrow) AlertOverlayUiaProvider(hwnd, ElementId{});
        if (! created)
        {
            return E_OUTOFMEMORY;
        }

        *root = static_cast<IRawElementProviderFragmentRoot*>(created);
        return S_OK;
    }

    HRESULT NavigateSibling(const std::vector<ElementId>& children, NavigateDirection direction, IRawElementProviderFragment** provider) noexcept
    {
        for (size_t index = 0u; index < children.size(); ++index)
        {
            if (! SameElement(children[index], _id))
            {
                continue;
            }

            if (direction == NavigateDirection_NextSibling && index + 1u < children.size())
            {
                return CreateFragment(_hwnd, children[index + 1u], provider);
            }
            if (direction == NavigateDirection_PreviousSibling && index > 0u)
            {
                return CreateFragment(_hwnd, children[index - 1u], provider);
            }
            return S_OK;
        }

        return S_OK;
    }

private:
    std::atomic<ULONG> _refCount{1u};
    HWND _hwnd = nullptr;
    ElementId _id{};
};

[[nodiscard]] LRESULT ReturnAlertOverlayUiaProvider(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    auto* provider = new (std::nothrow) AlertOverlayUiaProvider(hwnd, AlertOverlayUiaProvider::ElementId{});
    if (! provider)
    {
        return 0;
    }

    const LRESULT result = UiaReturnRawElementProvider(hwnd, wp, lp, static_cast<IRawElementProviderSimple*>(provider));
    static_cast<void>(provider->Release());
    return result;
}

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
        case WM_GETOBJECT:
            if (lp == static_cast<LPARAM>(UiaRootObjectId))
            {
                return ReturnAlertOverlayUiaProvider(hwnd, wp, lp);
            }
            break;
        case kUiaInvokeButtonMessage: InvokeButton(static_cast<uint32_t>(wp)); return 0;
        case kUiaInvokeDismissMessage: InvokeDismiss(); return 0;
        case WM_PAINT: OnPaint(); return 0;
        case WM_DPICHANGED_AFTERPARENT:
            _dpi = GetDpiForWindow(hwnd);
            if (_target)
            {
                _target->SetDpi(static_cast<float>(_dpi), static_cast<float>(_dpi));
            }
            _panelRegionPx.reset();
            UpdatePlacement();
            InvalidateRect(hwnd, nullptr, FALSE);
            return 0;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_MOUSEMOVE: OnMouseMove(PointFromLParam(lp)); return 0;
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_LBUTTONDOWN: OnLButtonDown(PointFromLParam(lp)); return 0;
        case WM_LBUTTONUP: OnLButtonUp(PointFromLParam(lp)); return 0;
        case WM_CAPTURECHANGED: OnCaptureChanged(reinterpret_cast<HWND>(lp)); return 0;
        case WM_MOUSEACTIVATE: return MA_ACTIVATE;
        case WM_KEYDOWN: OnKeyDown(wp); return 0;
        case WM_SYSCHAR:
            if (OnSysChar(wp))
            {
                return 0;
            }
            break;
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
    auto endDraw = wil::scope_exit([&]
    {
        const HRESULT hr = _target->EndDraw();
        if (hr == D2DERR_RECREATE_TARGET)
        {
            DiscardD2DResources();
        }
    });

    _target->Clear(D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f));
    if (_backdropBitmap)
    {
        _target->DrawBitmap(_backdropBitmap.get(), D2D1::RectF(0.0f, 0.0f, widthDip, heightDip), 1.0f, D2D1_BITMAP_INTERPOLATION_MODE_LINEAR);
    }
    _overlay.Draw(_target.get(), _dwriteFactory.get(), widthDip, heightDip, now);
#if defined(ENABLE_TESTS)
    const float lastDrawOpacity = _overlay.DebugGetLastDrawOpacityForTest();
    _debugMinimumDrawOpacity    = (_debugPaintCount == 0u) ? lastDrawOpacity : std::min(_debugMinimumDrawOpacity, lastDrawOpacity);
    ++_debugPaintCount;
#endif

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

    if (! EnsureOverlayLayoutForInput())
    {
        return;
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

#if defined(ENABLE_TESTS)
    ++_debugMouseDownCount;
    _debugLastMouseDownPointPx = pt;
    _debugLastMouseDownHitPart = -1;
#endif

    if (! EnsureOverlayLayoutForInput())
    {
        return;
    }

    const float xDip       = DipFromPx(pt.x);
    const float yDip       = DipFromPx(pt.y);
    const AlertHitTest hit = _overlay.HitTest(D2D1::Point2F(xDip, yDip));
#if defined(ENABLE_TESTS)
    _debugLastMouseDownHitPart = static_cast<int>(hit.part);
#endif
    if (hit.part == AlertHitTest::Part::None)
    {
        return;
    }

    _pressedHit = hit;
    SetCapture(_hwnd.get());
    if (_overlay.UpdateHotState(D2D1::Point2F(xDip, yDip)))
    {
        InvalidateRect(_hwnd.get(), nullptr, FALSE);
    }
}

void AlertOverlayWindow::OnLButtonUp(POINT pt) noexcept
{
    if (! _visible || ! _hwnd)
    {
        _pressedHit = {};
        return;
    }

#if defined(ENABLE_TESTS)
    ++_debugMouseUpCount;
    _debugLastMouseUpPointPx = pt;
    _debugLastMouseUpHitPart = -1;
#endif

    if (! EnsureOverlayLayoutForInput())
    {
        _pressedHit = {};
        return;
    }

    const AlertHitTest pressed = _pressedHit;
    _pressedHit                = {};
    if (GetCapture() == _hwnd.get())
    {
        ReleaseCapture();
    }

    const float xDip       = DipFromPx(pt.x);
    const float yDip       = DipFromPx(pt.y);
    const AlertHitTest hit = _overlay.HitTest(D2D1::Point2F(xDip, yDip));
#if defined(ENABLE_TESTS)
    _debugLastMouseUpHitPart = static_cast<int>(hit.part);
#endif
    // If activation or another transient capture path consumed the mouse-down, a
    // release on the close glyph still behaves like a standard dialog close.
    if (pressed.part == AlertHitTest::Part::None && hit.part == AlertHitTest::Part::Close)
    {
        InvokeDismiss();
        return;
    }

    if (pressed.part == AlertHitTest::Part::Close && hit.part == AlertHitTest::Part::Close)
    {
        InvokeDismiss();
        return;
    }

    if (pressed.part == AlertHitTest::Part::Button && hit.part == AlertHitTest::Part::Button && pressed.buttonId == hit.buttonId)
    {
        InvokeButton(hit.buttonId);
        return;
    }

    if (_overlay.UpdateHotState(D2D1::Point2F(xDip, yDip)))
    {
        InvalidateRect(_hwnd.get(), nullptr, FALSE);
    }
}

void AlertOverlayWindow::OnCaptureChanged(HWND newCapture) noexcept
{
    if (newCapture != _hwnd.get())
    {
        _pressedHit = {};
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

bool AlertOverlayWindow::OnSysChar(WPARAM key) noexcept
{
    if (! _visible || ! _hwnd)
    {
        return false;
    }

    const wchar_t mnemonic = static_cast<wchar_t>(key);
    if (mnemonic == L'\0')
    {
        return false;
    }

    const auto normalize = [](wchar_t ch) noexcept -> wchar_t { return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(ch))); };

    for (const auto& button : _overlay.GetModel().buttons)
    {
        for (const wchar_t ch : button.label)
        {
            if (ch == L'&')
            {
                continue;
            }

            if (std::iswspace(static_cast<wint_t>(ch)))
            {
                continue;
            }

            if (normalize(ch) == normalize(mnemonic))
            {
                InvokeButton(button.id);
                return true;
            }
            break;
        }
    }

    return false;
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
        static_cast<void>(EnsureOverlayLayoutForInput());
        const LPARAM messagePos = static_cast<LPARAM>(GetMessagePos());
        POINT pt{GET_X_LPARAM(messagePos), GET_Y_LPARAM(messagePos)};
        if (ScreenToClient(_hwnd.get(), &pt) != 0)
        {
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
#if defined(ENABLE_TESTS)
    ++_debugDismissCount;
#endif

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
        _pressedHit         = {};

        if (_hwnd && GetCapture() == _hwnd.get())
        {
            ReleaseCapture();
        }

        _overlay.ClearHotState();
        _panelRegionPx.reset();
        _backdropBitmap.reset();
        ClearRegion();

        HWND restoreFocus = nullptr;
        HWND restoreRoot  = nullptr;
        if (_restoreFocus && IsWindow(_restoreFocus))
        {
            restoreFocus = _restoreFocus;
            restoreRoot  = GetAncestor(_restoreFocus, GA_ROOT);
        }
        else if (_hostParent && IsWindow(_hostParent))
        {
            restoreRoot = GetAncestor(_hostParent, GA_ROOT);
        }

        if (_hwnd)
        {
            ShowWindow(_hwnd.get(), SW_HIDE);
        }

        if (restoreRoot && IsWindow(restoreRoot))
        {
            SetForegroundWindow(restoreRoot);
        }

        if (restoreFocus)
        {
            SetFocus(restoreFocus);
        }
        _restoreFocus = nullptr;

        // HostServices retains hidden overlays for reuse; release graphics/text resources
        // immediately so debug-layer shutdown cannot see them as live process-exit objects.
        DiscardD2DResources();
        _d2dFactory.reset();
        _dwriteFactory.reset();

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

    const bool wasVisible     = _visible && _hwnd && IsWindowVisible(_hwnd.get()) != FALSE;
    const bool wasBlocksInput = _blocksInput;

    _blocksInput        = blocksInput;
    _visible            = true;
    _trackingMouseLeave = false;
    _pressedHit         = {};
    _panelRegionPx.reset();
#if defined(ENABLE_TESTS)
    _debugPaintCount           = 0;
    _debugMinimumDrawOpacity   = 1.0f;
    _debugBackdropCaptureCount = 0;
    _debugBackdropSizePx       = {};
    _debugMouseDownCount       = 0;
    _debugMouseUpCount         = 0;
    _debugDismissCount         = 0;
    _debugLastMouseDownPointPx = {};
    _debugLastMouseUpPointPx   = {};
    _debugLastMouseDownHitPart = -1;
    _debugLastMouseUpHitPart   = -1;
#endif

    _overlay.SetTheme(*theme);
    _overlay.SetModel(std::move(*model));
    _overlay.ClearHotState();

    const uint64_t now = GetTickCount64();
    // Modal alerts must be readable on the first paint; otherwise they can appear blank
    // until another input message causes a later animation frame.
    const uint64_t visibleStartTick = _blocksInput ? ((now >= kShowAnimationMs) ? (now - kShowAnimationMs) : 0u) : now;
    _startTickMs                    = visibleStartTick;
    _animateUntilTickMs             = _blocksInput ? now : (now + kShowAnimationMs);
    _overlay.SetStartTick(visibleStartTick);

    const AlertSeverity severity = _overlay.GetModel().severity;
    _alwaysAnimate               = severity == AlertSeverity::Busy;

    _restoreFocus = nullptr;
    if (_blocksInput)
    {
        _restoreFocus = GetFocus();
    }

    ClearRegion();
    UpdatePlacement();
    if (EnsureOverlayLayoutForInput())
    {
        ApplyRegionFromOverlay();
    }
    const bool reuseVisibleBackdrop = wasVisible && wasBlocksInput && _blocksInput && _backdropBitmap;
    if (! reuseVisibleBackdrop)
    {
        CaptureBackdrop(wasVisible && _blocksInput);
    }

    if (_blocksInput)
    {
        SetWindowPos(_hwnd.get(), HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
        SetFocus(_hwnd.get());
    }
    else
    {
        SetWindowPos(_hwnd.get(), HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
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

    const DWORD style   = WS_POPUP | WS_CLIPSIBLINGS | WS_CLIPCHILDREN;
    const DWORD exStyle = WS_EX_TOOLWINDOW;
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
    DiscardD2DResources();
    _d2dFactory.reset();
    _dwriteFactory.reset();

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

    ApplyAttachmentState(hostParent, anchor, hostParent && hostParent != anchor, true);
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

    if (_hostParentSubclassed && _hostParent && IsWindow(_hostParent))
    {
        RestoreWndProcHook(_hostParent, kParentOriginalWndProcProp, kParentStateProp);
    }

    if (_anchorSubclassed && _anchor && IsWindow(_anchor))
    {
        RestoreWndProcHook(_anchor, kAnchorOriginalWndProcProp, kAnchorStateProp);
    }

    _hostParentSubclassed = false;
    _anchorSubclassed     = false;

    _hostParent = hostParent;
    _anchor     = anchor;
    if (trackHostParent && _hostParent && IsWindow(_hostParent))
    {
        if (SetPropW(_hostParent, kParentStateProp, this) != 0 &&
            InstallWndProcHook(_hostParent, kParentOriginalWndProcProp, &AlertOverlayWindow::ParentWndProc))
        {
            _hostParentSubclassed = true;
        }
        else
        {
            RemovePropW(_hostParent, kParentStateProp);
            RemovePropW(_hostParent, kParentOriginalWndProcProp);
        }
    }

    if (trackAnchor && _anchor && IsWindow(_anchor))
    {
        if (SetPropW(_anchor, kAnchorStateProp, this) != 0 && InstallWndProcHook(_anchor, kAnchorOriginalWndProcProp, &AlertOverlayWindow::AnchorWndProc))
        {
            _anchorSubclassed = true;
        }
        else
        {
            RemovePropW(_anchor, kAnchorStateProp);
            RemovePropW(_anchor, kAnchorOriginalWndProcProp);
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
        if (GetWindowRect(_anchor, &rc) == 0)
        {
            GetClientRect(_hostParent, &rc);
            POINT pts[2]{{rc.left, rc.top}, {rc.right, rc.bottom}};
            MapWindowPoints(_hostParent, nullptr, pts, 2);
            rc = RECT{pts[0].x, pts[0].y, pts[1].x, pts[1].y};
        }
    }
    else
    {
        GetClientRect(_hostParent, &rc);
        POINT pts[2]{{rc.left, rc.top}, {rc.right, rc.bottom}};
        MapWindowPoints(_hostParent, nullptr, pts, 2);
        rc = RECT{pts[0].x, pts[0].y, pts[1].x, pts[1].y};
    }

    const int width  = std::max(0L, rc.right - rc.left);
    const int height = std::max(0L, rc.bottom - rc.top);

    UINT flags = SWP_NOACTIVATE;
    if (! _visible)
    {
        flags |= SWP_NOZORDER | SWP_NOOWNERZORDER;
    }
    else if (! _blocksInput)
    {
        flags |= SWP_NOOWNERZORDER;
    }

    SetWindowPos(_hwnd.get(), HWND_TOP, rc.left, rc.top, width, height, flags);
    if (_visible && _blocksInput && IsWindowVisible(_hwnd.get()) != FALSE)
    {
        EnsureD2DResources();
        CaptureBackdrop(false);
        InvalidateRect(_hwnd.get(), nullptr, FALSE);
    }
}

bool AlertOverlayWindow::EnsureOverlayLayoutForInput() noexcept
{
    if (! _hwnd || ! _visible)
    {
        return false;
    }

    RECT clientRect{};
    if (GetClientRect(_hwnd.get(), &clientRect) != 0)
    {
        _clientSizePx.cx = std::max(0L, clientRect.right - clientRect.left);
        _clientSizePx.cy = std::max(0L, clientRect.bottom - clientRect.top);
    }

    if (_clientSizePx.cx <= 0 || _clientSizePx.cy <= 0)
    {
        return false;
    }

    EnsureD2DResources();
    if (! _dwriteFactory)
    {
        return false;
    }

    return _overlay.EnsureLayout(_dwriteFactory.get(), DipFromPx(_clientSizePx.cx), DipFromPx(_clientSizePx.cy));
}

LRESULT CALLBACK AlertOverlayWindow::ParentWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<AlertOverlayWindow*>(GetPropW(hwnd, kParentStateProp));
    if (! self)
    {
        return CallStoredWndProc(hwnd, kParentOriginalWndProcProp, msg, wp, lp);
    }

    if (msg == WM_SIZE || msg == WM_WINDOWPOSCHANGED)
    {
        self->UpdatePlacement();
    }

    if (msg == WM_NCDESTROY)
    {
        if (self->_hostParent == hwnd)
        {
            self->_hostParent           = nullptr;
            self->_hostParentSubclassed = false;
        }

        if (self->_anchor == hwnd)
        {
            self->_anchor           = nullptr;
            self->_anchorSubclassed = false;
        }

        self->Destroy();
        const LRESULT result = CallStoredWndProc(hwnd, kParentOriginalWndProcProp, msg, wp, lp);
        RemovePropW(hwnd, kParentStateProp);
        RemovePropW(hwnd, kParentOriginalWndProcProp);
        return result;
    }

    return CallStoredWndProc(hwnd, kParentOriginalWndProcProp, msg, wp, lp);
}

LRESULT CALLBACK AlertOverlayWindow::AnchorWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<AlertOverlayWindow*>(GetPropW(hwnd, kAnchorStateProp));
    if (! self)
    {
        return CallStoredWndProc(hwnd, kAnchorOriginalWndProcProp, msg, wp, lp);
    }

    if (msg == WM_SIZE || msg == WM_WINDOWPOSCHANGED)
    {
        self->UpdatePlacement();
    }

    if (msg == WM_NCDESTROY)
    {
        if (self->_anchor == hwnd)
        {
            self->_anchor           = nullptr;
            self->_anchorSubclassed = false;
        }

        self->Hide();

        HWND hostParent = self->_hostParent;
        if (hostParent == hwnd)
        {
            hostParent = nullptr;
        }

        self->ApplyAttachmentState(hostParent, nullptr, false, false);
        const LRESULT result = CallStoredWndProc(hwnd, kAnchorOriginalWndProcProp, msg, wp, lp);
        RemovePropW(hwnd, kAnchorStateProp);
        RemovePropW(hwnd, kAnchorOriginalWndProcProp);
        return result;
    }

    return CallStoredWndProc(hwnd, kAnchorOriginalWndProcProp, msg, wp, lp);
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
    _backdropBitmap.reset();
    _target.reset();
    _overlay.ResetDeviceResources();
    _overlay.ResetTextResources();
}

void AlertOverlayWindow::CaptureBackdrop(bool preserveExistingOnFailure) noexcept
{
    wil::com_ptr<ID2D1Bitmap> capturedBackdrop;
    const auto publishCapture = wil::scope_exit([&] noexcept
    {
        if (capturedBackdrop)
        {
            _backdropBitmap = std::move(capturedBackdrop);
        }
        else if (! preserveExistingOnFailure)
        {
            _backdropBitmap.reset();
        }
    });

    if (! _blocksInput || ! _hwnd)
    {
        return;
    }

    EnsureD2DResources();
    if (! _target)
    {
        return;
    }

    RECT screenRect{};
    if (GetWindowRect(_hwnd.get(), &screenRect) == FALSE)
    {
        return;
    }

    const LONG widthPx  = screenRect.right - screenRect.left;
    const LONG heightPx = screenRect.bottom - screenRect.top;
    if (widthPx <= 0 || heightPx <= 0)
    {
        return;
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = widthPx;
    bmi.bmiHeader.biHeight      = -heightPx;
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    wil::unique_hdc_window screenDc{GetDC(nullptr)};
    if (! screenDc)
    {
        return;
    }

    wil::unique_hdc memoryDc{CreateCompatibleDC(screenDc.get())};
    if (! memoryDc)
    {
        return;
    }

    wil::unique_hbitmap bitmap{CreateDIBSection(screenDc.get(), &bmi, DIB_RGB_COLORS, &bits, nullptr, 0)};
    if (! bitmap || ! bits)
    {
        return;
    }

    [[maybe_unused]] const auto oldBitmap = wil::SelectObject(memoryDc.get(), bitmap.get());
    bool capturedPixels                   = false;
    if (_anchor && IsWindow(_anchor) != FALSE)
    {
        RECT anchorRect{};
        if (GetWindowRect(_anchor, &anchorRect) != FALSE && anchorRect.left == screenRect.left && anchorRect.top == screenRect.top &&
            anchorRect.right == screenRect.right && anchorRect.bottom == screenRect.bottom)
        {
            capturedPixels = PrintWindow(_anchor, memoryDc.get(), PW_RENDERFULLCONTENT) != FALSE;
        }
    }

    if (! capturedPixels && _hostParent && IsWindow(_hostParent) != FALSE)
    {
        RECT hostClient{};
        if (GetClientRect(_hostParent, &hostClient) != FALSE)
        {
            std::array<POINT, 2> hostClientPoints{{
                POINT{hostClient.left, hostClient.top},
                POINT{hostClient.right, hostClient.bottom},
            }};
            if (MapWindowPoints(_hostParent, nullptr, hostClientPoints.data(), static_cast<UINT>(hostClientPoints.size())) != 0)
            {
                const RECT hostClientScreen{hostClientPoints[0].x, hostClientPoints[0].y, hostClientPoints[1].x, hostClientPoints[1].y};
                if (hostClientScreen.left == screenRect.left && hostClientScreen.top == screenRect.top && hostClientScreen.right == screenRect.right &&
                    hostClientScreen.bottom == screenRect.bottom)
                {
                    capturedPixels = PrintWindow(_hostParent, memoryDc.get(), PW_CLIENTONLY | PW_RENDERFULLCONTENT) != FALSE;
                }
            }
        }
    }

    if (! capturedPixels)
    {
        capturedPixels = BitBlt(memoryDc.get(), 0, 0, widthPx, heightPx, screenDc.get(), screenRect.left, screenRect.top, SRCCOPY | CAPTUREBLT) != FALSE;
    }

    if (! capturedPixels)
    {
        return;
    }

    const UINT32 width  = static_cast<UINT32>(widthPx);
    const UINT32 height = static_cast<UINT32>(heightPx);
    const D2D1_BITMAP_PROPERTIES props =
        D2D1::BitmapProperties(D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE), static_cast<float>(_dpi), static_cast<float>(_dpi));
    const HRESULT hr = _target->CreateBitmap(D2D1::SizeU(width, height), bits, width * 4u, props, capturedBackdrop.addressof());
    if (SUCCEEDED(hr))
    {
#if defined(ENABLE_TESTS)
        ++_debugBackdropCaptureCount;
        _debugBackdropSizePx = SIZE{widthPx, heightPx};
#endif
    }
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

#if defined(ENABLE_TESTS)
bool DebugGetAlertOverlayWindowSnapshot(HWND hwnd, AlertOverlayWindowDebugSnapshot& out) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto* window = reinterpret_cast<AlertOverlayWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! window || window->_hwnd.get() != hwnd)
    {
        return false;
    }

    out.visible                    = window->_visible;
    out.hasLayout                  = window->_overlay.HasLayout();
    out.hasBackdropBitmap          = window->_backdropBitmap != nullptr;
    out.paintCount                 = window->_debugPaintCount;
    out.lastDrawOpacity            = window->_overlay.DebugGetLastDrawOpacityForTest();
    out.lastDrawScrimOpacity       = window->_overlay.DebugGetLastDrawScrimOpacityForTest();
    out.minimumDrawOpacity         = window->_debugMinimumDrawOpacity;
    out.backdropCaptureCount       = window->_debugBackdropCaptureCount;
    out.clientSizePx               = window->_clientSizePx;
    out.backdropSizePx             = window->_debugBackdropSizePx;
    const D2D1_RECT_F closeRectDip = window->_overlay.DebugGetCloseRectForTest();
    out.closeRectPx.left           = window->PxFromDipFloor(closeRectDip.left);
    out.closeRectPx.top            = window->PxFromDipFloor(closeRectDip.top);
    out.closeRectPx.right          = window->PxFromDipCeil(closeRectDip.right);
    out.closeRectPx.bottom         = window->PxFromDipCeil(closeRectDip.bottom);
    out.usesSharedCloseChrome      = window->_overlay.DebugUsesSharedCloseChromeForTest();
    out.usesSharedButtonChrome     = window->_overlay.DebugUsesSharedButtonChromeForTest();
    out.mouseDownCount             = window->_debugMouseDownCount;
    out.mouseUpCount               = window->_debugMouseUpCount;
    out.dismissCount               = window->_debugDismissCount;
    out.lastMouseDownPointPx       = window->_debugLastMouseDownPointPx;
    out.lastMouseUpPointPx         = window->_debugLastMouseUpPointPx;
    out.lastMouseDownHitPart       = window->_debugLastMouseDownHitPart;
    out.lastMouseUpHitPart         = window->_debugLastMouseUpHitPart;
    return true;
}
#endif
} // namespace RedSalamander::Ui
