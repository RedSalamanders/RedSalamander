#include "AlertOverlayWindow.h"
#include "AnimationDispatcher.h"
#include "Helpers.h"
#include "Win32CallbackHelpers.h"
#include "WindowActivation.h"
#include "WindowMessages.h"
#include "resource.h"

#include <UIAutomation.h>
#include <algorithm>
#include <array>
#include <atomic>
#include <bit>
#include <cmath>
#include <cwctype>
#include <format>
#include <memory>
#include <new>
#include <oleauto.h>
#include <thread>
#include <vector>
#include <windowsx.h>

#pragma comment(lib, "uiautomationcore.lib")

namespace RedSalamander::Ui
{
namespace
{
// The alert records its existing DIP-based draw before host composition. This
// root only replays prepared content; it adds no competing input/UIA controls.
class AlertPresentationRoot final : public DxUi::Control
{
public:
    AlertPresentationRoot(const wil::com_ptr<ID2D1DeviceContext>& context, const wil::com_ptr<ID2D1CommandList>& frame) noexcept
        : _context(context),
          _frame(frame)
    {
    }
    AlertPresentationRoot(const AlertPresentationRoot&)            = delete;
    AlertPresentationRoot(AlertPresentationRoot&&)                 = delete;
    AlertPresentationRoot& operator=(const AlertPresentationRoot&) = delete;
    AlertPresentationRoot& operator=(AlertPresentationRoot&&)      = delete;

    void Paint(DxUi::WindowHost& host) const override
    {
        if (auto* context = host.GetDeviceContext(); context && context == _context.get() && _frame)
            context->DrawImage(_frame.get());
        else if (_frame)
            host.Invalidate(); // Retry only a prepared frame invalidated by device loss, never a persistent preparation failure.
    }

private:
    const wil::com_ptr<ID2D1DeviceContext>& _context;
    const wil::com_ptr<ID2D1CommandList>& _frame;
};

constexpr wchar_t kAlertOverlayWindowClassName[]    = L"RedSalamander.AlertOverlayWindow";
constexpr wchar_t kAlertOverlayWindowIdentityProp[] = L"RedSalamander.AlertOverlay.WindowIdentity";
constexpr uint64_t kShowAnimationMs                 = 220;
constexpr wchar_t kTrackedWindowHookProp[]          = L"RedSalamander.AlertOverlay.TrackedWindowHook";

#if defined(ENABLE_TESTS)
std::atomic<uint64_t> g_debugForegroundRequestCount{0u};

// The self-tests run without activation, so they can neither move the thread's real focus nor
// attach a UI Automation client: they stand in the focus window and capture the events raised.
thread_local bool g_debugFocusOverrideActive = false;
thread_local HWND g_debugFocusOverride       = nullptr;
thread_local bool g_debugUiaEventCapture     = false;
// Key messages a self-test sends carry no message time of their own; it stands one in for the key guard.
thread_local bool g_debugKeyMessageTimeActive = false;
thread_local DWORD g_debugKeyMessageTime      = 0u;
// A self-test without activation stands in the foreground window a hiding blocking alert checks.
thread_local bool g_debugForegroundOverrideActive = false;
thread_local HWND g_debugForegroundOverride       = nullptr;

struct DebugUiaEventRecord final
{
    EVENTID eventId     = 0;
    uint8_t elementKind = 0u;
    size_t elementIndex = 0u;
    uint64_t generation = 0u;
    HRESULT result      = E_PENDING;
};
thread_local std::vector<DebugUiaEventRecord> g_debugUiaEvents;
#endif

[[nodiscard]] HWND CurrentFocusWindow() noexcept
{
#if defined(ENABLE_TESTS)
    if (g_debugFocusOverrideActive)
        return g_debugFocusOverride;
#endif
    return ::GetFocus();
}

// The time the key message being handled was posted (GetMessageTime, milliseconds since system start on
// the GetTickCount clock).
[[nodiscard]] DWORD CurrentKeyMessageTime() noexcept
{
#if defined(ENABLE_TESTS)
    if (g_debugKeyMessageTimeActive)
        return g_debugKeyMessageTime;
#endif
    return static_cast<DWORD>(::GetMessageTime());
}

[[nodiscard]] DWORD CurrentTickCount32() noexcept
{
    return static_cast<DWORD>(::GetTickCount64());
}

[[nodiscard]] HWND CurrentForegroundWindow() noexcept
{
#if defined(ENABLE_TESTS)
    if (g_debugForegroundOverrideActive)
        return g_debugForegroundOverride;
#endif
    return ::GetForegroundWindow();
}

// Whether a key is down for this thread (its message queue state) or physically (a held key whose
// messages went to another window).
[[nodiscard]] bool IsKeyHeldDown(int virtualKey) noexcept
{
    return (::GetKeyState(virtualKey) & 0x8000) != 0 || (::GetAsyncKeyState(virtualKey) & 0x8000) != 0;
}

// The keys the guard of a decision overlay lets act (and the Ctrl of its Ctrl+Tab chord), one bit each.
constexpr std::array<WPARAM, 5u> kGuardKeys{{VK_RETURN, VK_ESCAPE, VK_F6, VK_TAB, VK_CONTROL}};

[[nodiscard]] uint8_t GuardKeyBit(WPARAM key) noexcept
{
    for (size_t index = 0u; index < kGuardKeys.size(); ++index)
    {
        if (kGuardKeys[index] == key)
            return static_cast<uint8_t>(1u << index);
    }
    return 0u;
}

// `candidate` is the overlay, a window inside it, or a window its owner chain ties to it (its own menu).
[[nodiscard]] bool IsWindowOfOverlay(HWND overlay, HWND candidate) noexcept
{
    if (! overlay || ! candidate)
        return false;
    if (candidate == overlay || IsChild(overlay, candidate) != FALSE)
        return true;
    for (HWND owner = GetWindow(GetAncestor(candidate, GA_ROOT), GW_OWNER); owner != nullptr; owner = GetWindow(owner, GW_OWNER))
    {
        if (owner == overlay)
            return true;
    }
    return false;
}

// lParam bit 30 of WM_KEYDOWN: the key was already down (an auto-repeat).
constexpr LPARAM kKeyPreviouslyDownFlag = static_cast<LPARAM>(1) << 30;

// The display text of a notification must be a null-terminated BSTR; the view need not be.
[[nodiscard]] wil::unique_bstr MakeNotificationBstr(std::wstring_view text) noexcept
{
    wil::unique_bstr value(SysAllocStringLen(text.data(), static_cast<UINT>(text.size())));
    return value;
}

[[nodiscard]] bool UiaEventsWanted() noexcept
{
#if defined(ENABLE_TESTS)
    if (g_debugUiaEventCapture)
        return true;
#endif
    return UiaClientsAreListening() != FALSE;
}

uint64_t NextAlertOverlayInstanceId() noexcept
{
    static std::atomic<uint64_t> next{1u};
    uint64_t identity = next.fetch_add(1u, std::memory_order_relaxed);
    if (identity == 0u)
        identity = next.fetch_add(1u, std::memory_order_relaxed);
    return identity;
}

struct AlertOverlayInvokePayload final
{
    std::weak_ptr<const uint64_t> instanceToken;
    uint64_t generation = 0u;
    uint32_t targetId   = 0u;
};

// A UI Automation call marshalled to the overlay's UI thread. It lives on the calling thread's stack
// while that thread is inside SendMessagePayload; never post or timeout it.
struct AlertUiaCall
{
    void* context                     = nullptr;
    HRESULT (*invoke)(void*) noexcept = nullptr;
    HRESULT result                    = static_cast<HRESULT>(UIA_E_ELEMENTNOTAVAILABLE);
};

// The registered payload behind the message's opaque token. Any window on the desktop can send the
// overlay a message with any LPARAM, so the LPARAM is never read as a pointer.
struct AlertUiaCallPayload final
{
    AlertUiaCall* call = nullptr;
};

[[nodiscard]] UINT AlertOverlayUiaCallMessage() noexcept
{
    static const UINT message = RegisterWindowMessageW(L"RedSalamander.AlertOverlay.UiaCall");
    return message;
}

LRESULT DispatchAlertUiaCall(HWND hwnd, UINT message, LPARAM token) noexcept
{
    // A token that yields no payload is stale, drained or forged: ignore it.
    const auto payload = TakeMessagePayload<AlertUiaCallPayload>(hwnd, message, token);
    if (! payload || ! payload->call || ! payload->call->invoke)
        return FALSE;
    payload->call->result = payload->call->invoke(payload->call->context);
    return TRUE;
}

POINT PointFromLParam(LPARAM lp) noexcept
{
    POINT pt{};
    pt.x = static_cast<LONG>(static_cast<short>(LOWORD(lp)));
    pt.y = static_cast<LONG>(static_cast<short>(HIWORD(lp)));
    return pt;
}

// Every overlay that follows a window (as its anchor or its host parent) shares one owned
// GWLP_WNDPROC link on that window and registers in it under its own id, so overlays on the same
// window chain independently: detaching one never unhooks another. The link is taken out only when
// no overlay remains and it is still the window's current procedure; under a later subclass it stays
// as a pass-through (restoring the original there would drop that subclass) until WM_NCDESTROY.
struct AlertTrackedWindowListener final
{
    uint64_t id                 = 0u;
    AlertOverlayWindow* overlay = nullptr;
    bool anchorRole             = false;
};

struct AlertTrackedWindowHook final
{
    WNDPROC originalWndProc = nullptr;
    std::vector<AlertTrackedWindowListener> listeners;
};

[[nodiscard]] AlertTrackedWindowHook* GetTrackedWindowHook(HWND hwnd) noexcept
{
    return static_cast<AlertTrackedWindowHook*>(GetPropW(hwnd, kTrackedWindowHookProp));
}

[[nodiscard]] uint64_t RegisterTrackedWindowListener(HWND hwnd, AlertOverlayWindow* overlay, bool anchorRole, WNDPROC hookWndProc) noexcept
{
    static std::atomic<uint64_t> nextId{0u};
    AlertTrackedWindowHook* hook = GetTrackedWindowHook(hwnd);
    if (! hook)
    {
        auto created = std::make_unique<AlertTrackedWindowHook>();
        if (Win32Callback::SetPropNoThrow(hwnd, kTrackedWindowHookProp, created.get()) == FALSE)
            return 0u;
        // A window always has a procedure, so a null previous one means the hook was refused.
        const auto previous = reinterpret_cast<WNDPROC>(Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(hookWndProc)));
        if (! previous)
        {
            RemovePropW(hwnd, kTrackedWindowHookProp);
            return 0u;
        }
        created->originalWndProc = previous;
        hook                     = created.release();
    }

    const uint64_t id = nextId.fetch_add(1u, std::memory_order_relaxed) + 1u;
    hook->listeners.push_back(AlertTrackedWindowListener{id, overlay, anchorRole});
    return id;
}

void UnregisterTrackedWindowListener(HWND hwnd, uint64_t id, WNDPROC hookWndProc) noexcept
{
    if (! hwnd || id == 0u || IsWindow(hwnd) == FALSE)
        return;
    AlertTrackedWindowHook* hook = GetTrackedWindowHook(hwnd);
    if (! hook)
        return;
    std::erase_if(hook->listeners, [id](const AlertTrackedWindowListener& listener) noexcept { return listener.id == id; });
    if (! hook->listeners.empty() || reinterpret_cast<WNDPROC>(GetWindowLongPtrW(hwnd, GWLP_WNDPROC)) != hookWndProc)
        return;

    static_cast<void>(Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(hook->originalWndProc)));
    RemovePropW(hwnd, kTrackedWindowHookProp);
    std::unique_ptr<AlertTrackedWindowHook> released(hook);
}
} // namespace

AlertOverlayWindow::AlertOverlayWindow() noexcept : _instanceToken(std::make_shared<const uint64_t>(NextAlertOverlayInstanceId()))
{
    _overlay.SetScrollingEnabled(true);
}

HWND AlertOverlayFocusWindow() noexcept
{
    return CurrentFocusWindow();
}

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
        Option,
        Button,
    };

    struct ElementId
    {
        ElementKind kind = ElementKind::Root;
        size_t index     = 0u;
    };

    AlertOverlayUiaProvider(HWND hwnd, ElementId id, std::weak_ptr<const uint64_t> instanceToken, uint64_t generation) noexcept
        : _hwnd(hwnd),
          _id(id),
          _instanceToken(std::move(instanceToken)),
          _generation(generation)
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
        if (GetWindowThreadProcessId(_hwnd, nullptr) != GetCurrentThreadId())
            return OnOwnerThread([&]() noexcept { return GetPropertyValue(propertyId, value); });
        if (! value)
        {
            return E_POINTER;
        }

        VariantInit(value);
        AlertOverlayWindow* window = ResolveWindow();
        if (! window)
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
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
            case UIA_IsEnabledPropertyId: return SetBool(value, IsEnabled(*window));
            case UIA_IsKeyboardFocusablePropertyId: return SetBool(value, IsEnabled(*window) && (IsButtonLike() || _id.kind == ElementKind::Root));
            case UIA_HasKeyboardFocusPropertyId:
            {
                const auto focused = window->_overlay.GetFocusedControl();
                const auto hit     = ControlHit(*window);
                return SetBool(value,
                               ::GetFocus() == _hwnd && (_id.kind == ElementKind::Root || (hit.part != AlertHitTest::Part::None && hit.part == focused.part &&
                                                                                           hit.buttonId == focused.buttonId)));
            }
            case UIA_IsOffscreenPropertyId:
            {
                const auto rect = VisibleRect(*window);
                return SetBool(value, IsWindowVisible(_hwnd) == FALSE || rect.right <= rect.left || rect.bottom <= rect.top);
            }
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
        if (GetWindowThreadProcessId(_hwnd, nullptr) != GetCurrentThreadId())
            return OnOwnerThread([&]() noexcept { return get_HostRawElementProvider(provider); });
        if (! provider)
        {
            return E_POINTER;
        }

        *provider = nullptr;
        if (_id.kind != ElementKind::Root)
        {
            return S_OK;
        }

        if (! ResolveWindow())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }

        return UiaHostProviderFromHwnd(_hwnd, provider);
    }

    HRESULT STDMETHODCALLTYPE Navigate(NavigateDirection direction, IRawElementProviderFragment** provider) noexcept override
    {
        if (GetWindowThreadProcessId(_hwnd, nullptr) != GetCurrentThreadId())
            return OnOwnerThread([&]() noexcept { return Navigate(direction, provider); });
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
                    return CreateFragment(_hwnd, children.front(), _instanceToken, _generation, provider);
                }
                return S_OK;
            case NavigateDirection_LastChild:
                if (_id.kind == ElementKind::Root && ! children.empty())
                {
                    return CreateFragment(_hwnd, children.back(), _instanceToken, _generation, provider);
                }
                return S_OK;
            case NavigateDirection_Parent:
                if (_id.kind != ElementKind::Root)
                {
                    return CreateFragment(_hwnd, ElementId{}, _instanceToken, _generation, provider);
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
        if (GetWindowThreadProcessId(_hwnd, nullptr) != GetCurrentThreadId())
            return OnOwnerThread([&]() noexcept { return GetRuntimeId(runtimeId); });
        if (! runtimeId)
        {
            return E_POINTER;
        }

        *runtimeId                 = nullptr;
        AlertOverlayWindow* window = ResolveWindow();
        if (! window)
            return UIA_E_ELEMENTNOTAVAILABLE;

        SAFEARRAY* array = SafeArrayCreateVector(VT_I4, 0, 9u);
        if (! array)
        {
            return E_OUTOFMEMORY;
        }

        const std::array<int, 9> values = RuntimeIdValues(_hwnd, *window->_instanceToken, _generation, _id);

        for (LONG i = 0; i < 9; ++i)
        {
            LONG item        = values[static_cast<size_t>(i)];
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
        if (GetWindowThreadProcessId(_hwnd, nullptr) != GetCurrentThreadId())
            return OnOwnerThread([&]() noexcept { return get_BoundingRectangle(rect); });
        if (! rect)
        {
            return E_POINTER;
        }

        *rect              = {};
        const auto* window = ResolveWindow();
        if (! window)
            return UIA_E_ELEMENTNOTAVAILABLE;
        const auto visible = VisibleRect(*window);
        if (visible.right <= visible.left || visible.bottom <= visible.top)
            return S_OK;
        POINT origin{};
        if (! ClientToScreen(_hwnd, &origin))
            return UIA_E_ELEMENTNOTAVAILABLE;
        rect->left   = static_cast<double>(origin.x + window->PxFromDipFloor(visible.left));
        rect->top    = static_cast<double>(origin.y + window->PxFromDipFloor(visible.top));
        rect->width  = static_cast<double>(window->PxFromDipCeil(visible.right) - window->PxFromDipFloor(visible.left));
        rect->height = static_cast<double>(window->PxFromDipCeil(visible.bottom) - window->PxFromDipFloor(visible.top));
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
        if (GetWindowThreadProcessId(_hwnd, nullptr) != GetCurrentThreadId())
            return OnOwnerThread([&]() noexcept { return SetFocus(); });
        if (! _hwnd || IsWindow(_hwnd) == FALSE)
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }

        auto* window = ResolveWindow();
        if (! window)
            return UIA_E_ELEMENTNOTAVAILABLE;
        if (_id.kind == ElementKind::Root)
        {
            ::SetFocus(_hwnd);
            return S_OK;
        }
        const auto hit = ControlHit(*window);
        if (hit.part == AlertHitTest::Part::None)
            return UIA_E_NOTSUPPORTED;
        if (! IsEnabled(*window))
            return UIA_E_ELEMENTNOTENABLED;
        return window->OnFocusControl(hit.part, hit.buttonId) ? S_OK : UIA_E_ELEMENTNOTAVAILABLE;
    }

    HRESULT STDMETHODCALLTYPE get_FragmentRoot(IRawElementProviderFragmentRoot** root) noexcept override
    {
        if (GetWindowThreadProcessId(_hwnd, nullptr) != GetCurrentThreadId())
            return OnOwnerThread([&]() noexcept { return get_FragmentRoot(root); });
        if (! root)
        {
            return E_POINTER;
        }

        *root = nullptr;
        if (! ResolveWindow())
            return UIA_E_ELEMENTNOTAVAILABLE;
        return CreateRoot(_hwnd, _instanceToken, _generation, root);
    }

    HRESULT STDMETHODCALLTYPE ElementProviderFromPoint(double x, double y, IRawElementProviderFragment** provider) noexcept override
    {
        if (GetWindowThreadProcessId(_hwnd, nullptr) != GetCurrentThreadId())
            return OnOwnerThread([&]() noexcept { return ElementProviderFromPoint(x, y, provider); });
        if (! provider)
        {
            return E_POINTER;
        }

        *provider          = nullptr;
        const auto* window = ResolveWindow();
        if (! window)
            return UIA_E_ELEMENTNOTAVAILABLE;
        POINT origin{};
        if (! ClientToScreen(_hwnd, &origin))
            return UIA_E_ELEMENTNOTAVAILABLE;
        const float dpiScale = static_cast<float>(window->_dpi) / 96.0f;
        const auto hit    = window->_overlay.HitTest(D2D1::Point2F(static_cast<float>(x - origin.x) / dpiScale, static_cast<float>(y - origin.y) / dpiScale));
        const auto& model = window->_overlay.GetModel();
        if (hit.part == AlertHitTest::Part::Close)
            return CreateFragment(_hwnd, {ElementKind::CloseButton, 0u}, _instanceToken, _generation, provider);
        for (size_t i = 0; i < model.options.size(); ++i)
            if (hit.part == AlertHitTest::Part::Option && model.options[i].id == hit.buttonId)
                return CreateFragment(_hwnd, {ElementKind::Option, i}, _instanceToken, _generation, provider);
        for (size_t i = 0; i < model.buttons.size(); ++i)
            if (hit.part == AlertHitTest::Part::Button && model.buttons[i].id == hit.buttonId)
                return CreateFragment(_hwnd, {ElementKind::Button, i}, _instanceToken, _generation, provider);
        return CreateFragment(_hwnd, ElementId{}, _instanceToken, _generation, provider);
    }

    HRESULT STDMETHODCALLTYPE GetFocus(IRawElementProviderFragment** provider) noexcept override
    {
        if (GetWindowThreadProcessId(_hwnd, nullptr) != GetCurrentThreadId())
            return OnOwnerThread([&]() noexcept { return GetFocus(provider); });
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

        return CreateFragment(_hwnd, FocusedElement(*window), _instanceToken, _generation, provider);
    }

    HRESULT STDMETHODCALLTYPE Invoke() noexcept override
    {
        if (GetWindowThreadProcessId(_hwnd, nullptr) != GetCurrentThreadId())
            return OnOwnerThread([&]() noexcept { return Invoke(); });
        AlertOverlayWindow* window = ResolveWindow();
        if (! window || ! IsButtonLike())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }

        if (_id.kind == ElementKind::CloseButton)
        {
            return PostInvokeMessage(*window, WndMsg::kAlertOverlayInvokeDismiss, 0u);
        }

        if (_id.kind == ElementKind::Option)
        {
            const auto& options = window->_overlay.GetModel().options;
            if (_id.index >= options.size())
            {
                return UIA_E_ELEMENTNOTAVAILABLE;
            }
            if (! IsEnabled(*window))
                return UIA_E_ELEMENTNOTENABLED;
            return PostInvokeMessage(*window, WndMsg::kAlertOverlayInvokeOption, options[_id.index].id);
        }

        const auto& buttons = window->_overlay.GetModel().buttons;
        if (_id.index >= buttons.size())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }

        return PostInvokeMessage(*window, WndMsg::kAlertOverlayInvokeButton, buttons[_id.index].id);
    }

    // The element the alert's keyboard focus is on (the control GetFocus reports, else the root).
    [[nodiscard]] static ElementId FocusedElement(const AlertOverlayWindow& window) noexcept
    {
        const AlertHitTest focused = window._overlay.GetFocusedControl();
        const AlertModel& model    = window._overlay.GetModel();
        if (focused.part == AlertHitTest::Part::Close)
            return {ElementKind::CloseButton, 0u};
        if (focused.part == AlertHitTest::Part::Button)
            for (size_t index = 0; index < model.buttons.size(); ++index)
                if (model.buttons[index].id == focused.buttonId)
                    return {ElementKind::Button, index};
        if (focused.part == AlertHitTest::Part::Option)
            for (size_t index = 0; index < model.options.size(); ++index)
                if (model.options[index].id == focused.buttonId)
                    return {ElementKind::Option, index};
        return {};
    }

    // A provider for one element of the window's current presentation (for raising events).
    [[nodiscard]] static wil::com_ptr_nothrow<IRawElementProviderSimple> ForCurrentPresentation(const AlertOverlayWindow& window, ElementId id) noexcept
    {
        wil::com_ptr_nothrow<IRawElementProviderSimple> provider;
        auto* created = new (std::nothrow) AlertOverlayUiaProvider(window._hwnd.get(), id, window._instanceToken, window._presentationGeneration);
        if (created)
            provider.attach(static_cast<IRawElementProviderSimple*>(created));
        return provider;
    }

    // Runtime IDs encode the HWND, the window instance, the presentation generation and the element.
    [[nodiscard]] static std::array<int, 9> RuntimeIdValues(HWND hwnd, uint64_t instance, uint64_t generation, ElementId id) noexcept
    {
        const auto lowWord       = [](uint64_t value) noexcept { return std::bit_cast<int>(static_cast<uint32_t>(value)); };
        const auto highWord      = [](uint64_t value) noexcept { return std::bit_cast<int>(static_cast<uint32_t>(value >> 32u)); };
        const uint64_t hwndValue = static_cast<uint64_t>(reinterpret_cast<ULONG_PTR>(hwnd));
        return std::to_array<int>({UiaAppendRuntimeId,
                                   lowWord(hwndValue),
                                   highWord(hwndValue),
                                   lowWord(instance),
                                   highWord(instance),
                                   lowWord(generation),
                                   highWord(generation),
                                   static_cast<int>(id.kind),
                                   lowWord(static_cast<uint64_t>(id.index))});
    }

private:
    HRESULT PostInvokeMessage(AlertOverlayWindow& window, UINT message, uint32_t targetId) const noexcept
    {
        if (! window.IsCurrentPresentation(_instanceToken, _generation))
            return UIA_E_ELEMENTNOTAVAILABLE;
        auto payload           = std::make_unique<AlertOverlayInvokePayload>();
        payload->instanceToken = _instanceToken;
        payload->generation    = _generation;
        payload->targetId      = targetId;
        if (PostMessagePayload(_hwnd, message, 0, std::move(payload)))
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
        return _id.kind == ElementKind::Button || _id.kind == ElementKind::Option || _id.kind == ElementKind::CloseButton;
    }

    [[nodiscard]] bool IsEnabled(const AlertOverlayWindow& window) const noexcept
    {
        if (_id.kind != ElementKind::Option)
            return true;
        const auto& options = window._overlay.GetModel().options;
        return _id.index < options.size() && window._overlay.IsOptionEnabled(options[_id.index].id);
    }

    [[nodiscard]] AlertOverlayWindow* ResolveWindow() const noexcept
    {
        const std::shared_ptr<const uint64_t> instanceToken = _instanceToken.lock();
        if (! instanceToken)
            return nullptr;
        if (! _hwnd || IsWindow(_hwnd) == FALSE)
        {
            return nullptr;
        }

        // A numeric HWND can be recycled for an unrelated window. Prove this is
        // the exact overlay instance before interpreting its app-owned userdata.
        if (GetPropW(_hwnd, kAlertOverlayWindowIdentityProp) != reinterpret_cast<HANDLE>(const_cast<uint64_t*>(instanceToken.get())))
            return nullptr;

        auto* window = reinterpret_cast<AlertOverlayWindow*>(GetWindowLongPtrW(_hwnd, GWLP_USERDATA));
        if (! window || window->_hwnd.get() != _hwnd || ! window->_visible || window->_instanceToken != instanceToken ||
            window->_presentationGeneration != _generation)
        {
            return nullptr;
        }

        return window;
    }

    template <typename Callback> HRESULT OnOwnerThread(Callback callback) noexcept
    {
        if (! _hwnd || ! IsWindow(_hwnd))
            return UIA_E_ELEMENTNOTAVAILABLE;
        const auto instanceToken = _instanceToken.lock();
        if (! instanceToken || GetPropW(_hwnd, kAlertOverlayWindowIdentityProp) != reinterpret_cast<HANDLE>(const_cast<uint64_t*>(instanceToken.get())))
            return UIA_E_ELEMENTNOTAVAILABLE;
        const UINT message = AlertOverlayUiaCallMessage();
        if (message == 0u)
            return UIA_E_ELEMENTNOTAVAILABLE;
        AlertUiaCall call{&callback, [](void* context) noexcept { return (*static_cast<Callback*>(context))(); }};
        auto payload  = std::make_unique<AlertUiaCallPayload>();
        payload->call = &call;
        return SendMessagePayload(_hwnd, message, 0, std::move(payload)) ? call.result : static_cast<HRESULT>(UIA_E_ELEMENTNOTAVAILABLE);
    }

    [[nodiscard]] AlertHitTest ControlHit(const AlertOverlayWindow& window) const noexcept
    {
        const auto& model = window._overlay.GetModel();
        if (_id.kind == ElementKind::CloseButton)
            return {AlertHitTest::Part::Close, 0u};
        if (_id.kind == ElementKind::Button && _id.index < model.buttons.size())
            return {AlertHitTest::Part::Button, model.buttons[_id.index].id};
        if (_id.kind == ElementKind::Option && _id.index < model.options.size())
            return {AlertHitTest::Part::Option, model.options[_id.index].id};
        return {};
    }

    [[nodiscard]] D2D1_RECT_F VisibleRect(const AlertOverlayWindow& window) const noexcept
    {
        if (_id.kind == ElementKind::Root)
            return window._overlay.GetPanelRect();
        if (_id.kind == ElementKind::Text)
            return window._overlay.GetVisibleTextRect();
        return window._overlay.GetVisibleControlRect(ControlHit(window));
    }

    [[nodiscard]] CONTROLTYPEID ControlType() const noexcept
    {
        switch (_id.kind)
        {
            case ElementKind::Root: return UIA_PaneControlTypeId;
            case ElementKind::Text: return UIA_TextControlTypeId;
            case ElementKind::CloseButton:
            case ElementKind::Option:
            case ElementKind::Button: return UIA_ButtonControlTypeId;
            default: return UIA_CustomControlTypeId;
        }
    }

    [[nodiscard]] std::wstring Name(const AlertOverlayWindow& window) const
    {
        const AlertModel& model = window._overlay.GetModel();
        switch (_id.kind)
        {
            case ElementKind::Root: return ! model.title.empty() ? model.title : LoadStringResource(nullptr, IDS_ALERT_OVERLAY_NAME);
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
            case ElementKind::CloseButton: return LoadStringResource(nullptr, IDS_ALERT_OVERLAY_CLOSE);
            case ElementKind::Option:
                if (_id.index < model.options.size())
                {
                    const auto& option = model.options[_id.index];
                    if (! option.choices.empty() && option.selectedIndex < option.choices.size())
                    {
                        return std::format(L"{}: {}", option.label, option.choices[option.selectedIndex].label);
                    }
                    return option.label;
                }
                return {};
            case ElementKind::Button:
                if (_id.index < model.buttons.size())
                {
                    return AlertButtonDisplayLabel(model.buttons[_id.index].label);
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
            case ElementKind::Option:
            {
                const auto& options = window._overlay.GetModel().options;
                const uint32_t id   = _id.index < options.size() ? options[_id.index].id : 0u;
                return std::format(L"AlertOverlay.Option.{}", id);
            }
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
        children.reserve(model.options.size() + model.buttons.size() + 2u);
        if (! model.title.empty() || ! model.message.empty())
        {
            children.push_back(ElementId{ElementKind::Text, 0u});
        }
        if (model.closable)
        {
            children.push_back(ElementId{ElementKind::CloseButton, 0u});
        }
        for (size_t index = 0u; index < model.options.size(); ++index)
        {
            children.push_back(ElementId{ElementKind::Option, index});
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

    static HRESULT CreateFragment(
        HWND hwnd, ElementId id, std::weak_ptr<const uint64_t> instanceToken, uint64_t generation, IRawElementProviderFragment** provider) noexcept
    {
        auto* created = new (std::nothrow) AlertOverlayUiaProvider(hwnd, id, std::move(instanceToken), generation);
        if (! created)
        {
            return E_OUTOFMEMORY;
        }

        *provider = static_cast<IRawElementProviderFragment*>(created);
        return S_OK;
    }

    static HRESULT CreateRoot(HWND hwnd, std::weak_ptr<const uint64_t> instanceToken, uint64_t generation, IRawElementProviderFragmentRoot** root) noexcept
    {
        auto* created = new (std::nothrow) AlertOverlayUiaProvider(hwnd, ElementId{}, std::move(instanceToken), generation);
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
                return CreateFragment(_hwnd, children[index + 1u], _instanceToken, _generation, provider);
            }
            if (direction == NavigateDirection_PreviousSibling && index > 0u)
            {
                return CreateFragment(_hwnd, children[index - 1u], _instanceToken, _generation, provider);
            }
            return S_OK;
        }

        return S_OK;
    }

private:
    std::atomic<ULONG> _refCount{1u};
    HWND _hwnd = nullptr;
    ElementId _id{};
    std::weak_ptr<const uint64_t> _instanceToken;
    uint64_t _generation = 0u;
};

namespace
{
void RecordUiaEvent(EVENTID eventId, AlertOverlayUiaProvider::ElementId element, uint64_t generation, HRESULT result) noexcept
{
#if defined(ENABLE_TESTS)
    if (g_debugUiaEventCapture)
        g_debugUiaEvents.push_back(DebugUiaEventRecord{eventId, static_cast<uint8_t>(element.kind), element.index, generation, result});
#else
    static_cast<void>(eventId);
    static_cast<void>(element);
    static_cast<void>(generation);
    static_cast<void>(result);
#endif
}
} // namespace

[[nodiscard]] LRESULT ReturnAlertOverlayUiaProvider(HWND hwnd, WPARAM wp, LPARAM lp, std::weak_ptr<const uint64_t> instanceToken, uint64_t generation) noexcept
{
    auto* provider = new (std::nothrow) AlertOverlayUiaProvider(hwnd, AlertOverlayUiaProvider::ElementId{}, std::move(instanceToken), generation);
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

void AlertOverlayWindow::SetTrackedWindowDestroyedCallback(void* context, void (*callback)(void* context, AlertOverlayWindow& overlay) noexcept) noexcept
{
    _trackedWindowDestroyedContext = context;
    _onTrackedWindowDestroyed      = callback;
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

HRESULT AlertOverlayWindow::ShowModelessForAnchor(HWND anchor, const AlertTheme& theme, AlertModel model, AlertModelessShowOptions options) noexcept
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

    const HRESULT hr = TransitionVisibility(true, &theme, &model, false, &options);
    if (SUCCEEDED(hr) && ! options.announcement.empty())
    {
        RaiseUiaAnnouncement(options.announcement);
    }
    return hr;
}

void AlertOverlayWindow::Hide() noexcept
{
    static_cast<void>(TransitionVisibility(false, nullptr, nullptr, false));
}

void AlertOverlayWindow::HideWithoutRestoringFocus() noexcept
{
    _suppressFocusRestoreOnHide = true;
    static_cast<void>(TransitionVisibility(false, nullptr, nullptr, false));
    _suppressFocusRestoreOnHide = false;
}

bool AlertOverlayWindow::HasKeyboardFocus() const noexcept
{
    return _hwnd && CurrentFocusWindow() == _hwnd.get();
}

std::optional<RECT> AlertOverlayWindow::GetButtonScreenRect(uint32_t buttonId) noexcept
{
    if (! _hwnd || ! _visible || ! EnsureOverlayLayoutForInput())
        return std::nullopt;
    const D2D1_RECT_F rectDip = _overlay.GetVisibleControlRect({AlertHitTest::Part::Button, buttonId});
    if (rectDip.right <= rectDip.left || rectDip.bottom <= rectDip.top)
        return std::nullopt;
    POINT corners[2]{{PxFromDipFloor(rectDip.left), PxFromDipFloor(rectDip.top)}, {PxFromDipCeil(rectDip.right), PxFromDipCeil(rectDip.bottom)}};
    SetLastError(ERROR_SUCCESS);
    if (MapWindowPoints(_hwnd.get(), nullptr, corners, 2) == 0 && GetLastError() != ERROR_SUCCESS)
        return std::nullopt;
    return RECT{corners[0].x, corners[0].y, corners[1].x, corners[1].y};
}

LRESULT CALLBACK AlertOverlayWindow::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        auto* self     = static_cast<AlertOverlayWindow*>(cs->lpCreateParams);
        if (! self || ! self->_instanceToken ||
            SetPropW(hwnd, kAlertOverlayWindowIdentityProp, reinterpret_cast<HANDLE>(const_cast<uint64_t*>(self->_instanceToken.get()))) == FALSE)
        {
            return FALSE;
        }
        InitPostedPayloadWindow(hwnd);
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
    if (const UINT uiaCallMessage = AlertOverlayUiaCallMessage(); uiaCallMessage != 0u && msg == uiaCallMessage)
        return DispatchAlertUiaCall(hwnd, msg, lp);

    switch (msg)
    {
        case WM_ERASEBKGND: return 1;
        case WM_GETOBJECT:
            if (lp == static_cast<LPARAM>(UiaRootObjectId))
            {
                return ReturnAlertOverlayUiaProvider(hwnd, wp, lp, _instanceToken, _presentationGeneration);
            }
            break;
        case WndMsg::kAlertOverlayInvokeButton: OnInvokeMessage(msg, lp); return 0;
        case WndMsg::kAlertOverlayInvokeOption: OnInvokeMessage(msg, lp); return 0;
        case WndMsg::kAlertOverlayInvokeDismiss: OnInvokeMessage(msg, lp); return 0;
        case WM_PAINT: OnPaint(); return 0;
        case WM_DPICHANGED:
        case WM_DPICHANGED_AFTERPARENT:
        {
            bool handled = false;
            static_cast<void>(_presentationHost.HandleMessage(hwnd, msg, wp, lp, handled));
            _dpi = GetDpiForWindow(hwnd);
            _overlay.ResetTextResources();
            _panelRegionPx.reset();
            UpdatePlacement(true);
            InvalidateRect(hwnd, nullptr, FALSE);
            return 0;
        }
        case WM_SHOWWINDOW:
        {
            bool handled = false;
            static_cast<void>(_presentationHost.HandleMessage(hwnd, msg, wp, lp, handled));
            break;
        }
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_MOUSEWHEEL: OnMouseWheel(GET_WHEEL_DELTA_WPARAM(wp), PointFromLParam(lp)); return 0;
        case WM_MOUSEMOVE: OnMouseMove(PointFromLParam(lp)); return 0;
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_LBUTTONDOWN: OnLButtonDown(PointFromLParam(lp)); return 0;
        case WM_LBUTTONUP: OnLButtonUp(PointFromLParam(lp)); return 0;
        case WM_RBUTTONDOWN:
        case WM_MBUTTONDOWN:
            // A click on the overlay is a deliberate interaction: it lifts the key guard.
            ReleaseKeyGuard();
            break;
        case WM_CAPTURECHANGED: OnCaptureChanged(reinterpret_cast<HWND>(lp)); return 0;
        case WM_MOUSEACTIVATE: return MA_ACTIVATE;
        case WM_KEYDOWN: OnKeyDown(wp, lp); return 0;
        case WM_KEYUP: OnKeyUp(wp); return 0;
        case WM_SETFOCUS: OnSetFocus(reinterpret_cast<HWND>(wp)); break;
        case WM_KILLFOCUS: OnKillFocus(reinterpret_cast<HWND>(wp)); break;
        case WM_CLOSE:
            if (OnCloseRequest())
            {
                return 0;
            }
            break;
        case WM_SYSCHAR:
            if (OnSysChar(wp))
            {
                return 0;
            }
            break;
        case WM_SETCURSOR: return OnSetCursor(reinterpret_cast<HWND>(wp), LOWORD(lp), HIWORD(lp));
        case WM_NCDESTROY:
        {
            // Destroyed with its owner even when no tracking hook saw that owner go: report hidden so
            // a prompt's nested loop exits. No focus or activation is restored during owner teardown.
            _visible      = false;
            _blocksInput  = false;
            _restoreFocus = nullptr;
            AdvancePresentationGeneration();
            static_cast<void>(DrainPostedPayloadsForWindow(hwnd));
            RemovePropW(hwnd, kAlertOverlayWindowIdentityProp);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            StopAnimationTimer();
            DiscardD2DResources();
            _presentationHost.Detach();
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
        return;
    EnsureD2DResources();
    const bool needsInitialFrame = ! _preparedFrame;
    HRESULT hr                   = E_FAIL;
    const float widthDip         = DipFromPx(_clientSizePx.cx);
    const float heightDip        = DipFromPx(_clientSizePx.cy);
    if (_target && _dwriteFactory && widthDip > 0.0f && heightDip > 0.0f)
    {
        wil::com_ptr<ID2D1CommandList> frame;
        hr = _target->CreateCommandList(frame.put());
        if (SUCCEEDED(hr))
        {
            // Borrow the host context only for preparation, restoring all modified state.
            wil::com_ptr<ID2D1Image> previousTarget;
            _target->GetTarget(previousTarget.put());
            D2D1_MATRIX_3X2_F previousTransform{};
            _target->GetTransform(&previousTransform);
            const auto previousTextAntialiasMode = _target->GetTextAntialiasMode();
            {
                const auto restore = wil::scope_exit([&]() noexcept
                {
                    _target->SetTarget(previousTarget.get());
                    _target->SetTransform(previousTransform);
                    _target->SetTextAntialiasMode(previousTextAntialiasMode);
                });
                _target->SetTarget(frame.get());
                _target->SetTransform(D2D1::Matrix3x2F::Identity());
                _target->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
                _target->BeginDraw();
                const auto endDraw = wil::scope_exit([&]() noexcept { hr = _target->EndDraw(); });
                _target->Clear(D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f));
                if (_backdropBitmap)
                    _target->DrawBitmap(_backdropBitmap.get(), D2D1::RectF(0.0f, 0.0f, widthDip, heightDip), 1.0f, D2D1_BITMAP_INTERPOLATION_MODE_LINEAR);
                _overlay.Draw(_target.get(), _dwriteFactory.get(), widthDip, heightDip, GetTickCount64());
            }
            if (SUCCEEDED(hr))
                hr = frame->Close();
            if (SUCCEEDED(hr))
                _preparedFrame = std::move(frame);
        }
    }
#if defined(ENABLE_TESTS)
    _debugLastEndDrawResult     = hr; // Recording result; the pixel witness separately verifies presented content.
    const float lastDrawOpacity = _overlay.DebugGetLastDrawOpacityForTest();
    _debugMinimumDrawOpacity    = (_debugPaintCount == 0u) ? lastDrawOpacity : std::min(_debugMinimumDrawOpacity, lastDrawOpacity);
    ++_debugPaintCount;
#endif
    if (FAILED(hr))
        DiscardD2DResources();
    ApplyRegionFromOverlay();
    // Nonmodal alerts clip their HWND to the panel. That makes WM_PAINT's first
    // dirty region partial, before the composition swap chain has a full frame.
    // Publish the prepared content once after creation/resize/device reset;
    // normal animation and subsequent paints keep their bounded dirty region.
    if (needsInitialFrame && _preparedFrame)
        static_cast<void>(_presentationHost.RenderInitialFrameForShow());
    bool handled = false;
    static_cast<void>(_presentationHost.HandleMessage(_hwnd.get(), WM_PAINT, 0, 0, handled));
    if (! handled)
    {
        PAINTSTRUCT ps{};
        const wil::unique_hdc_paint hdc = wil::BeginPaint(_hwnd.get(), &ps);
    }
}

void AlertOverlayWindow::OnSize(UINT width, UINT height) noexcept
{
    _clientSizePx.cx = static_cast<LONG>(width);
    _clientSizePx.cy = static_cast<LONG>(height);
    bool handled     = false;
    static_cast<void>(_presentationHost.HandleMessage(_hwnd.get(), WM_SIZE, SIZE_RESTORED, MAKELPARAM(width, height), handled));
    _preparedFrame.reset();
    _panelRegionPx.reset();
    if (_visible && EnsureOverlayLayoutForInput())
        _overlay.RevealFocusedControl();
}

LRESULT AlertOverlayWindow::OnFocusControl(AlertHitTest::Part part, uint32_t id) noexcept
{
    const AlertHitTest previous = _overlay.GetFocusedControl();
    if (! _visible || ! EnsureOverlayLayoutForInput() || ! _overlay.FocusControl({part, id}))
        return FALSE;
    SetFocus(_hwnd.get());
    // A UI Automation focus request is the user's deliberate move: it lifts the key guard (after the focus
    // gain, which arms it for any other origin).
    ReleaseKeyGuard();
    InvalidateRect(_hwnd.get(), nullptr, FALSE);
    if (previous.part != part || previous.buttonId != id)
        static_cast<void>(RaiseUiaFocusChanged());
    return TRUE;
}

void AlertOverlayWindow::OnMouseWheel(int delta, POINT screenPoint) noexcept
{
    if (! _visible || ! _hwnd || ! EnsureOverlayLayoutForInput() || ! ScreenToClient(_hwnd.get(), &screenPoint))
        return;
    const auto viewport = _overlay.GetBodyViewport();
    const float x       = DipFromPx(screenPoint.x);
    const float y       = DipFromPx(screenPoint.y);
    if (x < viewport.left || x > viewport.right + 12.0f || y < viewport.top || y > viewport.bottom)
        return;
    // Floating DIP movement preserves high-resolution partial wheel deltas.
    if (_overlay.ScrollBy(-static_cast<float>(delta) * 48.0f / WHEEL_DELTA))
    {
        _pressedHit = {};
        InvalidateRect(_hwnd.get(), nullptr, FALSE);
    }
}

void AlertOverlayWindow::OnMouseMove(POINT pt) noexcept
{
    if (_scrollDragging && _hwnd)
    {
        const auto track = _overlay.GetScrollTrack();
        if (_overlay.ScrollAt(D2D1::Point2F((track.left + track.right) * 0.5f, std::clamp(DipFromPx(pt.y), track.top, track.bottom))))
            InvalidateRect(_hwnd.get(), nullptr, FALSE);
        return;
    }
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
    // A click on the overlay is a deliberate interaction: it lifts the key guard (the click itself then
    // acts normally on button release).
    ReleaseKeyGuard();

#if defined(ENABLE_TESTS)
    ++_debugMouseDownCount;
    _debugLastMouseDownPointPx = pt;
    _debugLastMouseDownHitPart = -1;
#endif

    if (! EnsureOverlayLayoutForInput())
    {
        return;
    }

    const float xDip = DipFromPx(pt.x);
    const float yDip = DipFromPx(pt.y);
    const auto track = _overlay.GetScrollTrack();
    if (track.bottom > track.top && xDip >= track.left && xDip <= track.right && yDip >= track.top && yDip <= track.bottom)
    {
        _scrollDragging = true;
        SetCapture(_hwnd.get());
        static_cast<void>(_overlay.ScrollAt(D2D1::Point2F(xDip, yDip)));
        _pressedHit = {};
        InvalidateRect(_hwnd.get(), nullptr, FALSE);
        return;
    }
    const AlertHitTest hit = _overlay.HitTest(D2D1::Point2F(xDip, yDip));
#if defined(ENABLE_TESTS)
    _debugLastMouseDownHitPart = static_cast<int>(hit.part);
#endif
    if (hit.part == AlertHitTest::Part::None || (hit.part == AlertHitTest::Part::Option && ! _overlay.IsOptionEnabled(hit.buttonId)))
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
    if (_scrollDragging)
    {
        _scrollDragging = false;
        if (GetCapture() == _hwnd.get())
            ReleaseCapture();
        return;
    }
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

    if (pressed.part == AlertHitTest::Part::Option && hit.part == AlertHitTest::Part::Option && pressed.buttonId == hit.buttonId)
    {
        InvokeOption(hit.buttonId);
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
        _pressedHit     = {};
        _scrollDragging = false;
    }
}

void AlertOverlayWindow::OnKeyDown(WPARAM key, LPARAM flags) noexcept
{
    if (! _visible || ! _hwnd || ! AcceptsKeyboardInput())
    {
        return;
    }

    // A decision overlay never acts on an auto-repeated key: a key held for the pane underneath, or an
    // Escape or Enter held after it answered the previous decision, cannot answer this one. Other
    // non-blocking alerts (a plugin's non-modal alert) keep acting on it.
    if (_keysRequireFocus && (flags & kKeyPreviouslyDownFlag) != 0)
    {
        return;
    }

    if (_keysGuarded)
    {
        // The overlay got the focus without a deliberate interaction: the user may still be typing for the
        // pane. Only a fresh Enter (the default), Escape (the escape binding), F6 (the focus back to the
        // pane) or Ctrl+Tab (the guard lifted) acts.
        const bool liftChord = key == VK_TAB && (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        if ((key != VK_RETURN && key != VK_ESCAPE && key != VK_F6 && ! liftChord) || ! PassesKeyGuard(key, liftChord))
        {
            return;
        }
        if (key == VK_ESCAPE)
        {
            InvokeDismiss(DismissSource::Escape);
        }
        else if (key == VK_RETURN)
        {
            if (_primaryButtonId.has_value())
                InvokeButton(_primaryButtonId.value());
        }
        else if (key == VK_F6)
        {
            ReturnFocusToAnchor();
        }
        else
        {
            LiftKeyGuardToDefault();
        }
        return;
    }

    // F6 or Shift+F6 leaves a decision overlay for the pane it covers, without answering: application
    // shortcuts (the View command included) do not reach a focused overlay, so this is its keyboard way out.
    if (_keysRequireFocus && key == VK_F6)
    {
        ReturnFocusToAnchor();
        return;
    }

    if (key == VK_ESCAPE)
    {
        InvokeDismiss(DismissSource::Escape);
        return;
    }

    if (key == VK_PRIOR || key == VK_NEXT || key == VK_HOME || key == VK_END)
    {
        if (! EnsureOverlayLayoutForInput())
            return;
        const auto viewport = _overlay.GetBodyViewport();
        const float page    = std::max(1.0f, viewport.bottom - viewport.top - 24.0f);
        const float delta   = key == VK_HOME ? -_overlay.GetScrollMaximum() : key == VK_END ? _overlay.GetScrollMaximum() : key == VK_PRIOR ? -page : page;
        if (_overlay.ScrollBy(delta))
        {
            _pressedHit = {};
            InvalidateRect(_hwnd.get(), nullptr, FALSE);
        }
        return;
    }

    if (key == VK_TAB)
    {
        const bool reverse = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        static_cast<void>(EnsureOverlayLayoutForInput());
        // A modeless overlay has no other keyboard route to its close affordance: Tab reaches it last.
        if (_overlay.FocusNextInteractive(reverse, ! _blocksInput && _keysRequireFocus))
        {
            InvalidateRect(_hwnd.get(), nullptr, FALSE);
            static_cast<void>(RaiseUiaFocusChanged());
        }
        return;
    }

    if (key == VK_LEFT || key == VK_UP || key == VK_RIGHT || key == VK_DOWN)
    {
        const bool reverse = (key == VK_LEFT || key == VK_UP);
        if (const std::optional<uint32_t> optionId = _overlay.GetFocusedOptionId(); optionId.has_value())
        {
            InvokeOption(optionId.value(), reverse);
            return;
        }
        if (_overlay.FocusNextButton(reverse))
        {
            InvalidateRect(_hwnd.get(), nullptr, FALSE);
            static_cast<void>(RaiseUiaFocusChanged());
        }
        return;
    }

    if (key == VK_RETURN || key == VK_SPACE)
    {
        if (_overlay.GetFocusedControl().part == AlertHitTest::Part::Close)
        {
            InvokeDismiss();
            return;
        }
        if (const std::optional<uint32_t> optionId = _overlay.GetFocusedOptionId(); optionId.has_value())
        {
            InvokeOption(optionId.value());
            return;
        }
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

void AlertOverlayWindow::OnKeyUp(WPARAM key) noexcept
{
    // A guard key that was down when the guard was armed was meant for something else; once released, the
    // next press is the user's own.
    _guardHeldKeys = static_cast<uint8_t>(_guardHeldKeys & ~GuardKeyBit(key));
}

void AlertOverlayWindow::OnSetFocus(HWND previousFocus) noexcept
{
    // A focus the overlay gets without a click on it or a UI Automation request (an activation that returns
    // to it, a prompt that hands the focus back) may arrive while the user types for another window: guard it
    // again. A click lifts the guard right after (its button-down follows the activation); a focus coming
    // back from the overlay's own menu keeps the guard as it was.
    if (! _visible || _blocksInput || ! _guardKeyboardFocus || ! _hwnd || IsWindowOfOverlay(_hwnd.get(), previousFocus))
        return;
    ArmKeyGuard();
    // Guarded, Enter submits the default: the focus ring says so.
    FocusDefaultButton();
}

void AlertOverlayWindow::OnKillFocus(HWND newFocus) noexcept
{
    // The user moved the focus away from a modeless overlay that took it: a later hide must not take the
    // focus back to the window it came from. Its own controls and its own menu (an owned popup) keep it.
    if (_blocksInput || ! _modelessRestoresFocus || ! _hwnd || IsWindowOfOverlay(_hwnd.get(), newFocus))
        return;
    _restoreFocus          = nullptr;
    _modelessRestoresFocus = false;
}

bool AlertOverlayWindow::OnCloseRequest() noexcept
{
    // Alt+F4 on a focused decision overlay is meant for the window the user works in: pass it to the
    // owner's top-level window instead of destroying an overlay window its owner still holds. Other
    // non-blocking alerts keep the default close.
    if (! _hwnd || _blocksInput || ! _keysRequireFocus)
        return false;
    const HWND owner = GetWindow(_hwnd.get(), GW_OWNER);
    const HWND root  = owner ? GetAncestor(owner, GA_ROOT) : nullptr;
    if (! root || IsWindow(root) == FALSE)
        return false;
    static_cast<void>(PostMessageW(root, WM_CLOSE, 0, 0));
    return true;
}

void AlertOverlayWindow::ArmKeyGuard() noexcept
{
    _keysGuarded    = true;
    _guardArmedTick = CurrentTickCount32();
    _guardHeldKeys  = 0u;
    for (const WPARAM key : kGuardKeys)
    {
        if (IsKeyHeldDown(static_cast<int>(key)))
            _guardHeldKeys = static_cast<uint8_t>(_guardHeldKeys | GuardKeyBit(key));
    }
}

bool AlertOverlayWindow::PassesKeyGuard(WPARAM key, bool controlChord) const noexcept
{
    if (! _keysGuarded)
        return true;
    // Typed before the arming, or so soon after it that it was meant for the pane (typeahead).
    const LONG sinceArmed = static_cast<LONG>(CurrentKeyMessageTime() - _guardArmedTick);
    if (sinceArmed < static_cast<LONG>(kModelessKeyGuardTypeaheadMs))
        return false;
    // Still held from before the arming (no key-up seen since), the Ctrl of a chord included.
    if ((_guardHeldKeys & GuardKeyBit(key)) != 0u)
        return false;
    return ! controlChord || (_guardHeldKeys & GuardKeyBit(VK_CONTROL)) == 0u;
}

void AlertOverlayWindow::ReleaseKeyGuard() noexcept
{
    _keysGuarded   = false;
    _guardHeldKeys = 0u;
}

void AlertOverlayWindow::LiftKeyGuardToDefault() noexcept
{
    ReleaseKeyGuard();
    FocusDefaultButton();
}

void AlertOverlayWindow::FocusDefaultButton() noexcept
{
    if (! _hwnd || ! _primaryButtonId.has_value() || ! EnsureOverlayLayoutForInput())
        return;
    const AlertHitTest previous = _overlay.GetFocusedControl();
    if (previous.part == AlertHitTest::Part::Button && previous.buttonId == _primaryButtonId.value())
        return;
    if (! _overlay.FocusControl({AlertHitTest::Part::Button, _primaryButtonId.value()}))
        return;
    InvalidateRect(_hwnd.get(), nullptr, FALSE);
    static_cast<void>(RaiseUiaFocusChanged());
}

void AlertOverlayWindow::ReturnFocusToAnchor() noexcept
{
    HWND target = nullptr;
    if (_anchor && _restoreFocus && IsWindow(_restoreFocus) != FALSE && IsWindowVisible(_restoreFocus) != FALSE &&
        (_restoreFocus == _anchor || IsChild(_anchor, _restoreFocus) != FALSE))
        target = _restoreFocus;
    else if (_anchor && IsWindow(_anchor) != FALSE && IsWindowVisible(_anchor) != FALSE)
        target = _anchor;
    if (! target)
        return;
    // The user sent the focus away: a later hide takes it nowhere.
    _restoreFocus          = nullptr;
    _modelessRestoresFocus = false;
    MoveFocusTo(target);
    if (_callbacks.onFocusReturned)
        _callbacks.onFocusReturned(_callbacks.context);
}

HWND AlertOverlayWindow::ResolveFocusReturnTarget(HWND target) const noexcept
{
    if (! target || IsWindow(target) == FALSE)
        return nullptr;
    if (IsWindowVisible(target) != FALSE)
        return target;
    const HWND fallback = _anchor ? _anchor : _hostParent;
    return fallback && IsWindow(fallback) != FALSE && IsWindowVisible(fallback) != FALSE ? fallback : nullptr;
}

bool AlertOverlayWindow::OnSysChar(WPARAM key) noexcept
{
    // A guarded overlay has no mnemonics: an Alt chord meant for the pane cannot press a button.
    if (! _visible || ! _hwnd || ! AcceptsKeyboardInput() || _keysGuarded)
    {
        return false;
    }

    const wchar_t mnemonic = static_cast<wchar_t>(key);
    if (mnemonic == L'\0')
    {
        return false;
    }

    const auto normalize = [](wchar_t ch) noexcept -> wchar_t { return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(ch))); };

    // Only an explicit '&' marker names a mnemonic; a key two buttons share selects neither.
    std::optional<uint32_t> match;
    for (const auto& button : _overlay.GetModel().buttons)
    {
        const wchar_t marked = AlertButtonMnemonic(button.label);
        if (marked == L'\0' || normalize(marked) != normalize(mnemonic))
            continue;
        if (match.has_value())
            return false;
        match = button.id;
    }

    if (! match.has_value())
        return false;
    InvokeButton(match.value());
    return true;
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
            if (hit.part == AlertHitTest::Part::Close || hit.part == AlertHitTest::Part::Button ||
                (hit.part == AlertHitTest::Part::Option && _overlay.IsOptionEnabled(hit.buttonId)))
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

void AlertOverlayWindow::InvokeOption(uint32_t optionId, bool reverse) noexcept
{
    const AlertHitTest previous = _overlay.GetFocusedControl();
    if (_overlay.ActivateOption(optionId, reverse) && _hwnd)
    {
        // Pointer/UIA activation follows the activated option, never stale tab focus.
        static_cast<void>(_overlay.FocusControl({AlertHitTest::Part::Option, optionId}));
        if (EnsureOverlayLayoutForInput())
            _overlay.RevealFocusedControl();
        InvalidateRect(_hwnd.get(), nullptr, FALSE);
        // A focus move announces the option with its new value; otherwise report the new value.
        const bool focusMoved = previous.part != AlertHitTest::Part::Option || previous.buttonId != optionId;
        if (! focusMoved || ! RaiseUiaFocusChanged())
            RaiseUiaOptionValueChanged(optionId);
        // Last: the owner keeps its own state as the truth and may re-present this overlay from it.
        if (_callbacks.onOptionChanged)
            _callbacks.onOptionChanged(_callbacks.context, optionId, _overlay.GetOptionValue(optionId).value_or(0u));
    }
}

void AlertOverlayWindow::SetOptionValue(uint32_t optionId, uint64_t value) noexcept
{
    if (! _overlay.SelectOptionValue(optionId, value) || ! _hwnd)
        return;
    if (_visible)
    {
        static_cast<void>(EnsureOverlayLayoutForInput());
        InvalidateRect(_hwnd.get(), nullptr, FALSE);
        RaiseUiaOptionValueChanged(optionId);
    }
}

bool AlertOverlayWindow::RaiseUiaFocusChanged() noexcept
{
    // The HWND keeps the system focus while focus moves between the alert's controls, so no system
    // focus event reports those moves: raise one for the control that now has the focus.
    if (! _hwnd || ! _visible || CurrentFocusWindow() != _hwnd.get() || ! UiaEventsWanted())
        return false;
    const AlertOverlayUiaProvider::ElementId element = AlertOverlayUiaProvider::FocusedElement(*this);
    const auto provider                              = AlertOverlayUiaProvider::ForCurrentPresentation(*this, element);
    const HRESULT hr                                 = provider ? UiaRaiseAutomationEvent(provider.get(), UIA_AutomationFocusChangedEventId) : E_OUTOFMEMORY;
    RecordUiaEvent(UIA_AutomationFocusChangedEventId, element, _presentationGeneration, hr);
    return SUCCEEDED(hr);
}

void AlertOverlayWindow::RaiseUiaOptionValueChanged(uint32_t optionId) noexcept
{
    if (! _hwnd || ! _visible || ! UiaEventsWanted())
        return;
    const auto& options = _overlay.GetModel().options;
    for (size_t index = 0; index < options.size(); ++index)
    {
        if (options[index].id != optionId)
            continue;
        // An option's name carries its current value.
        const AlertOverlayUiaProvider::ElementId element{AlertOverlayUiaProvider::ElementKind::Option, index};
        const auto provider = AlertOverlayUiaProvider::ForCurrentPresentation(*this, element);
        wil::unique_variant previousName;
        wil::unique_variant currentName;
        HRESULT hr = provider ? provider->GetPropertyValue(UIA_NamePropertyId, currentName.reset_and_addressof()) : E_OUTOFMEMORY;
        if (SUCCEEDED(hr))
            hr = UiaRaiseAutomationPropertyChangedEvent(provider.get(), UIA_NamePropertyId, previousName, currentName);
        RecordUiaEvent(UIA_AutomationPropertyChangedEventId, element, _presentationGeneration, hr);
        return;
    }
}

void AlertOverlayWindow::RaiseUiaPresentationReplaced(bool textChanged) noexcept
{
    // A new model advances the presentation generation, which retires every element a client holds:
    // invalidate the root's children, read out changed text and report where the focus now is.
    if (! _hwnd || ! _visible || ! UiaEventsWanted())
        return;
    const AlertOverlayUiaProvider::ElementId rootElement{};
    const auto root = AlertOverlayUiaProvider::ForCurrentPresentation(*this, rootElement);
    if (! root)
        return;
    std::array<int, 9> runtimeId = AlertOverlayUiaProvider::RuntimeIdValues(_hwnd.get(), *_instanceToken, _presentationGeneration, rootElement);
    const HRESULT structureHr =
        UiaRaiseStructureChangedEvent(root.get(), StructureChangeType_ChildrenInvalidated, runtimeId.data(), static_cast<int>(runtimeId.size()));
    RecordUiaEvent(UIA_StructureChangedEventId, rootElement, _presentationGeneration, structureHr);

    const AlertModel& model = _overlay.GetModel();
    if (textChanged && (! model.title.empty() || ! model.message.empty()))
    {
        const auto text = AlertOverlayUiaProvider::ForCurrentPresentation(*this, {AlertOverlayUiaProvider::ElementKind::Text, 0u});
        wil::unique_variant textName;
        HRESULT hr = text ? text->GetPropertyValue(UIA_NamePropertyId, textName.reset_and_addressof()) : E_OUTOFMEMORY;
        if (SUCCEEDED(hr) && textName.vt != VT_BSTR)
            hr = E_UNEXPECTED;
        if (SUCCEEDED(hr))
        {
            const wil::unique_bstr activityId = wil::make_bstr_nothrow(L"RedSalamander.AlertOverlay.Replaced");
            if (! activityId)
                hr = E_OUTOFMEMORY;
            else
                hr = UiaRaiseNotificationEvent(
                    root.get(), NotificationKind_Other, NotificationProcessing_ImportantMostRecent, textName.bstrVal, activityId.get());
        }
        RecordUiaEvent(UIA_NotificationEventId, rootElement, _presentationGeneration, hr);
    }

    static_cast<void>(RaiseUiaFocusChanged());
}

void AlertOverlayWindow::RaiseUiaAnnouncement(std::wstring_view text) noexcept
{
    // A first show of a modeless alert raises no structure or focus event a screen reader would read:
    // announce the owner's text once, on the root, like a replaced model announces its new text.
    if (! _hwnd || ! _visible || text.empty())
        return;
#if defined(ENABLE_TESTS)
    _debugLastAnnouncement.assign(text);
#endif
    if (! UiaEventsWanted())
        return;
    const AlertOverlayUiaProvider::ElementId rootElement{};
    const auto root = AlertOverlayUiaProvider::ForCurrentPresentation(*this, rootElement);
    if (! root)
        return;
    const wil::unique_bstr displayString = MakeNotificationBstr(text);
    const wil::unique_bstr activityId    = wil::make_bstr_nothrow(L"RedSalamander.AlertOverlay.Shown");
    const HRESULT hr =
        displayString && activityId
            ? UiaRaiseNotificationEvent(root.get(), NotificationKind_Other, NotificationProcessing_ImportantMostRecent, displayString.get(), activityId.get())
            : E_OUTOFMEMORY;
    RecordUiaEvent(UIA_NotificationEventId, rootElement, _presentationGeneration, hr);
}

bool AlertOverlayWindow::AcceptsKeyboardInput() const noexcept
{
    // A modeless decision overlay acts on keys only while it has the focus itself: an Escape the
    // message loop routes to every visible alert, while the user types in a pane, must not decide.
    return ! _keysRequireFocus || CurrentFocusWindow() == _hwnd.get();
}

void AlertOverlayWindow::MoveFocusTo(HWND hwnd) noexcept
{
#if defined(ENABLE_TESTS)
    if (g_debugFocusOverrideActive)
    {
        g_debugFocusOverride = hwnd;
        return;
    }
#endif
    static_cast<void>(::SetFocus(hwnd));
}

void AlertOverlayWindow::InvokeDismiss(DismissSource source) noexcept
{
#if defined(ENABLE_TESTS)
    ++_debugDismissCount;
#endif

    // The close affordance of an owner that installed onClose only closes this surface (never a decision);
    // Escape keeps the escape binding first.
    if (source == DismissSource::Close && _callbacks.onClose)
    {
        _callbacks.onClose(_callbacks.context);
        return;
    }

    if (_escapeButtonId.has_value())
    {
        InvokeButton(_escapeButtonId.value());
        return;
    }

    if (_callbacks.onClose)
    {
        _callbacks.onClose(_callbacks.context);
        return;
    }

    if (! _overlay.GetModel().closable)
    {
        return;
    }

    Hide();
}

void AlertOverlayWindow::OnInvokeMessage(UINT message, LPARAM parameter) noexcept
{
    auto payload = TakeMessagePayload<AlertOverlayInvokePayload>(_hwnd.get(), message, parameter);
    if (! payload || ! IsCurrentPresentation(payload->instanceToken, payload->generation) || ! _visible)
        return;

    switch (message)
    {
        case WndMsg::kAlertOverlayInvokeButton: InvokeButton(payload->targetId); break;
        case WndMsg::kAlertOverlayInvokeOption: InvokeOption(payload->targetId); break;
        case WndMsg::kAlertOverlayInvokeDismiss: InvokeDismiss(); break;
        default: break;
    }
}

void AlertOverlayWindow::AdvancePresentationGeneration() noexcept
{
    ++_presentationGeneration;
    if (_presentationGeneration == 0u)
        ++_presentationGeneration;
}

bool AlertOverlayWindow::IsCurrentPresentation(const std::weak_ptr<const uint64_t>& instanceToken, uint64_t generation) const noexcept
{
    const std::shared_ptr<const uint64_t> lockedToken = instanceToken.lock();
    return lockedToken && _instanceToken == lockedToken && _presentationGeneration == generation;
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

HRESULT AlertOverlayWindow::TransitionVisibility(
    bool show, const AlertTheme* theme, AlertModel* model, bool blocksInput, AlertModelessShowOptions* modeless) noexcept
{
    if (! show)
    {
        // Only a blocking alert that captured the focus hands it back, and only while this overlay
        // or that focus's root still holds the foreground: a modeless alert, or one cleared while the
        // user works elsewhere or its owner is being torn down, never activates a window.
        // A focus-return target that is no longer visible (a decision mirror that went away while this
        // prompt was open) cannot take the focus back: the alert's own anchor takes it instead.
        HWND restoreFocus = _blocksInput ? ResolveFocusReturnTarget(_restoreFocus) : nullptr;
        HWND restoreRoot  = restoreFocus ? GetAncestor(restoreFocus, GA_ROOT) : nullptr;
        if (const HWND foreground = CurrentForegroundWindow(); ! restoreRoot || (foreground != restoreRoot && (! _hwnd || foreground != _hwnd.get())))
        {
            restoreFocus = nullptr;
            restoreRoot  = nullptr;
        }
        // I35 R8.13: a modeless decision overlay that took the focus gives it back to the window that
        // had it, only while it still has the focus itself; it never asks for the foreground.
        HWND modelessRestoreFocus = (! _blocksInput && _modelessRestoresFocus && _restoreFocus && IsWindow(_restoreFocus) != FALSE &&
                                     IsWindowVisible(_restoreFocus) != FALSE && _hwnd && CurrentFocusWindow() == _hwnd.get())
                                        ? _restoreFocus
                                        : nullptr;
        if (_suppressFocusRestoreOnHide)
        {
            restoreFocus         = nullptr;
            restoreRoot          = nullptr;
            modelessRestoreFocus = nullptr;
        }

        AdvancePresentationGeneration();
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

        if (_hwnd)
        {
            ShowWindow(_hwnd.get(), SW_HIDE);
        }

        if (restoreRoot && IsWindow(restoreRoot) && CurrentForegroundWindow() != restoreRoot)
        {
#if defined(ENABLE_TESTS)
            g_debugForegroundRequestCount.fetch_add(1u, std::memory_order_relaxed);
#endif
            Common::RequestForegroundWindow(restoreRoot);
        }

        if (restoreFocus)
        {
            MoveFocusTo(restoreFocus);
        }
        if (modelessRestoreFocus)
        {
            MoveFocusTo(modelessRestoreFocus);
        }
        _restoreFocus           = nullptr;
        _modelessRestoresFocus  = false;
        _keysRequireFocus       = false;
        _followAnchorVisibility = false;
        _guardKeyboardFocus     = false;
        ReleaseKeyGuard();

        // HostServices may retain hidden overlays for reuse; release graphics/text resources
        // immediately so debug-layer shutdown cannot see them as live process-exit objects.
        DiscardD2DResources();
        _presentationHost.Detach();
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

    EnsureD2DResources();
    if (! _target || ! _dwriteFactory)
    {
        DiscardD2DResources();
        _presentationHost.Detach();
        return E_FAIL; // Never show a blocking surface without a usable presentation device.
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
    _debugLastEndDrawResult    = E_PENDING;
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

    const bool textChanged = _overlay.GetModel().title != model->title || _overlay.GetModel().message != model->message;
    _overlay.SetTheme(*theme);
    AdvancePresentationGeneration();
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

    // A model shown on an alert that already blocks keeps the focus captured when it began to block:
    // by now the overlay itself has the focus, and it is never a window to hand focus back to. A
    // modeless decision overlay that took the focus keeps its focus-return target the same way.
    const bool wasModelessFocus     = wasVisible && ! wasBlocksInput && _modelessRestoresFocus;
    const bool wasKeyGuarded        = wasVisible && ! wasBlocksInput && _keysGuarded;
    const HWND previousRestoreFocus = (wasVisible && wasBlocksInput) || wasModelessFocus ? _restoreFocus : nullptr;
    _restoreFocus                   = nullptr;
    _modelessRestoresFocus          = false;
    _keysRequireFocus               = false;
    _followAnchorVisibility         = false;
    _guardKeyboardFocus             = false;
    bool takeModelessFocus          = false;
    if (_blocksInput)
    {
        ReleaseKeyGuard();
        const HWND focus = previousRestoreFocus ? previousRestoreFocus : CurrentFocusWindow();
        if (focus && focus != _hwnd.get() && IsChild(_hwnd.get(), focus) == FALSE)
            _restoreFocus = focus;
    }
    else if (modeless)
    {
        _keysRequireFocus       = modeless->keysRequireFocus;
        _followAnchorVisibility = modeless->followAnchorVisibility;
        _guardKeyboardFocus     = modeless->guardKeyboardFocus;
        if (wasModelessFocus)
        {
            _restoreFocus          = previousRestoreFocus;
            _modelessRestoresFocus = previousRestoreFocus != nullptr;
        }
        else if (modeless->takeFocus)
        {
            const HWND focus = CurrentFocusWindow();
            if (focus && focus != _hwnd.get() && IsChild(_hwnd.get(), focus) == FALSE)
            {
                _restoreFocus          = focus;
                _modelessRestoresFocus = true;
            }
            takeModelessFocus = CurrentFocusWindow() != _hwnd.get();
        }
        if (modeless->takeFocus && modeless->guardKeyboardFocus)
        {
            // Armed from this show: keys typed from now on for kModelessKeyGuardTypeaheadMs, and a guard key
            // the user is still holding, were meant for the window the focus came from.
            ArmKeyGuard();
        }
        else if (! wasKeyGuarded || ! _guardKeyboardFocus)
        {
            // A re-presentation keeps an armed guard until the user interacts deliberately.
            ReleaseKeyGuard();
        }
    }
    else
    {
        ReleaseKeyGuard();
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
        ApplyAnchorVisibility();
        // Taking the focus within the application the user is already typing in is not a foreground
        // request: the owner decided that the anchor had the focus in the foreground window.
        if (takeModelessFocus && IsWindowVisible(_hwnd.get()) != FALSE)
            MoveFocusTo(_hwnd.get());
    }

    StartAnimationTimer();
    InvalidateRect(_hwnd.get(), nullptr, FALSE);
    if (wasVisible)
        RaiseUiaPresentationReplaced(textChanged);
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
    _presentationHost.Detach();
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

    // The overlay is a popup owned by the anchor's top-level window, and owned windows do not move with
    // their owner: a move of that window reaches neither a child anchor nor its parent (I35 F11).
    HWND root = trackAnchor && anchor ? GetAncestor(anchor, GA_ROOT) : nullptr;
    if (root == anchor || root == hostParent)
    {
        root = nullptr;
    }

    if (_hostParent == hostParent && _anchor == anchor && _root == root && (_hostParentTrackingId != 0u) == trackHostParent &&
        (_anchorTrackingId != 0u) == trackAnchor && (_rootTrackingId != 0u) == (root != nullptr))
    {
        return;
    }

    UnregisterTrackedWindowListener(_hostParent, _hostParentTrackingId, &AlertOverlayWindow::TrackedWindowProc);
    UnregisterTrackedWindowListener(_anchor, _anchorTrackingId, &AlertOverlayWindow::TrackedWindowProc);
    UnregisterTrackedWindowListener(_root, _rootTrackingId, &AlertOverlayWindow::TrackedWindowProc);
    _hostParentTrackingId = 0u;
    _anchorTrackingId     = 0u;
    _rootTrackingId       = 0u;

    _hostParent = hostParent;
    _anchor     = anchor;
    _root       = root;
    if (trackHostParent && _hostParent && IsWindow(_hostParent))
    {
        _hostParentTrackingId = RegisterTrackedWindowListener(_hostParent, this, false, &AlertOverlayWindow::TrackedWindowProc);
    }

    if (trackAnchor && _anchor && IsWindow(_anchor))
    {
        _anchorTrackingId = RegisterTrackedWindowListener(_anchor, this, true, &AlertOverlayWindow::TrackedWindowProc);
    }

    if (_root && IsWindow(_root))
    {
        // Placement only: the root dies after the anchor, whose destruction already detached this overlay.
        _rootTrackingId = RegisterTrackedWindowListener(_root, this, false, &AlertOverlayWindow::TrackedWindowProc);
    }
}

void AlertOverlayWindow::UpdatePlacement(bool keepZOrder) noexcept
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
    if (! _visible || (keepZOrder && ! _blocksInput))
    {
        // Only a show raises a non-blocking overlay: a move or activation change of the window it tracks
        // keeps it where it is in the z-order (it never jumps above another application's window). A
        // blocking alert is raised again with its owner, so nothing of that owner ever covers it.
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

void AlertOverlayWindow::ApplyAnchorVisibility() noexcept
{
    // A modeless decision overlay sits over a pane's view: while another tab hides that view, or its
    // window is hidden, the overlay window hides too and comes back with it. It stays logically shown.
    if (! _hwnd || ! _visible || _blocksInput || ! _followAnchorVisibility || ! _anchor || IsWindow(_anchor) == FALSE)
        return;
    const bool anchorShown  = IsWindowVisible(_anchor) != FALSE;
    const bool overlayShown = IsWindowVisible(_hwnd.get()) != FALSE;
    if (anchorShown == overlayShown)
        return;
    ShowWindow(_hwnd.get(), anchorShown ? SW_SHOWNA : SW_HIDE);
    if (anchorShown)
        InvalidateRect(_hwnd.get(), nullptr, FALSE);
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

LRESULT CALLBACK AlertOverlayWindow::TrackedWindowProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    const AlertTrackedWindowHook* hook = GetTrackedWindowHook(hwnd);
    const WNDPROC original             = hook ? hook->originalWndProc : nullptr;
    if (hook && ! hook->listeners.empty() && (msg == WM_SIZE || msg == WM_WINDOWPOSCHANGED || msg == WM_NCDESTROY))
    {
        // A listener may detach, destroy or delete overlays (itself included) and even free the hook:
        // dispatch from a snapshot of ids, re-resolve each one, and never touch a listener after its call.
        std::vector<uint64_t> ids;
        ids.reserve(hook->listeners.size());
        for (const AlertTrackedWindowListener& listener : hook->listeners)
            ids.push_back(listener.id);
        for (const uint64_t id : ids)
        {
            const AlertTrackedWindowHook* current = GetTrackedWindowHook(hwnd);
            if (! current)
                break;
            const auto it = std::ranges::find(current->listeners, id, &AlertTrackedWindowListener::id);
            if (it != current->listeners.end())
                it->overlay->OnTrackedWindowMessage(hwnd, msg, it->anchorRole);
        }
    }

    const LRESULT result = original ? Win32Callback::CallWindowProcNoThrow(original, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    if (msg == WM_NCDESTROY)
    {
        if (AlertTrackedWindowHook* remaining = GetTrackedWindowHook(hwnd))
        {
            RemovePropW(hwnd, kTrackedWindowHookProp);
            std::unique_ptr<AlertTrackedWindowHook> released(remaining);
        }
    }
    return result;
}

void AlertOverlayWindow::OnTrackedWindowMessage(HWND hwnd, UINT msg, bool anchorRole) noexcept
{
    if (msg != WM_NCDESTROY)
    {
        UpdatePlacement(true);
        ApplyAnchorVisibility();
        return;
    }

    // The tracked window is going away: never hand focus or activation back into it.
    _restoreFocus = nullptr;
    if (anchorRole)
    {
        Hide();
        ApplyAttachmentState(_hostParent == hwnd ? nullptr : _hostParent, nullptr, false, false);
    }
    else
    {
        Destroy();
    }

    // Last statement: the owner may delete this overlay.
    if (_onTrackedWindowDestroyed)
        _onTrackedWindowDestroyed(_trackedWindowDestroyedContext, *this);
}

void AlertOverlayWindow::EnsureD2DResources() noexcept
{
    if (! _hwnd)
        return;
    if (! _presentationHost.GetHwnd())
    {
        DxUi::WindowHost::AttachOptions options{};
        options.presentationMode = DxUi::WindowHost::PresentationMode::CompositionSwapChain;
        if (! _presentationHost.Attach(_hwnd.get(), options))
            return;
        _presentationHost.SetRoot(std::make_unique<AlertPresentationRoot>(_target, _preparedFrame));
    }
    if (! _presentationHost.PrimeForShow())
    {
        DiscardD2DResources();
        return;
    }
    const UINT dpi = GetDpiForWindow(_hwnd.get());
    if (dpi != 0u && dpi != _dpi)
    {
        _dpi = dpi;
        _overlay.ResetTextResources();
        _panelRegionPx.reset();
    }
    auto* context = _presentationHost.GetDeviceContext();
    if (_target.get() != context)
    {
        DiscardD2DResources();
        _target = context;
    }
    _dwriteFactory = _presentationHost.GetWriteFactory();
}

void AlertOverlayWindow::DiscardD2DResources() noexcept
{
    _preparedFrame.reset();
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
bool DebugCaptureAlertOverlayWindowBitmap(HWND hwnd, DxUi::WindowHostBitmapCapture& out) noexcept
{
    out = {};
    if (! hwnd || GetWindowThreadProcessId(hwnd, nullptr) != GetCurrentThreadId())
        return false;
    auto* window = reinterpret_cast<AlertOverlayWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window && window->_hwnd.get() == hwnd && window->_presentationHost.DebugCaptureBitmap(out);
}

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

    out.visible                     = window->_visible;
    out.hasLayout                   = window->_overlay.HasLayout();
    out.hasBackdropBitmap           = window->_backdropBitmap != nullptr;
    out.paintCount                  = window->_debugPaintCount;
    out.lastEndDrawResult           = window->_debugLastEndDrawResult;
    out.renderTargetWindowState     = 0u; // Composition presentation does not expose the legacy HWND occlusion flag.
    out.usesCompositionPresentation = window->_presentationHost.GetHwnd() == hwnd;
    out.hasPreparedFrame            = static_cast<bool>(window->_preparedFrame);
    out.lastDrawOpacity             = window->_overlay.DebugGetLastDrawOpacityForTest();
    out.lastDrawScrimOpacity        = window->_overlay.DebugGetLastDrawScrimOpacityForTest();
    out.minimumDrawOpacity          = window->_debugMinimumDrawOpacity;
    out.backdropCaptureCount        = window->_debugBackdropCaptureCount;
    out.clientSizePx                = window->_clientSizePx;
    out.backdropSizePx              = window->_debugBackdropSizePx;
    const D2D1_RECT_F closeRectDip  = window->_overlay.DebugGetCloseRectForTest();
    out.closeRectPx.left            = window->PxFromDipFloor(closeRectDip.left);
    out.closeRectPx.top             = window->PxFromDipFloor(closeRectDip.top);
    out.closeRectPx.right           = window->PxFromDipCeil(closeRectDip.right);
    out.closeRectPx.bottom          = window->PxFromDipCeil(closeRectDip.bottom);
    out.usesSharedCloseChrome       = window->_overlay.DebugUsesSharedCloseChromeForTest();
    out.usesSharedButtonChrome      = window->_overlay.DebugUsesSharedButtonChromeForTest();
    out.mouseDownCount              = window->_debugMouseDownCount;
    out.mouseUpCount                = window->_debugMouseUpCount;
    out.dismissCount                = window->_debugDismissCount;
    out.lastMouseDownPointPx        = window->_debugLastMouseDownPointPx;
    out.lastMouseUpPointPx          = window->_debugLastMouseUpPointPx;
    out.lastMouseDownHitPart        = window->_debugLastMouseDownHitPart;
    out.lastMouseUpHitPart          = window->_debugLastMouseUpHitPart;
    const AlertModel& model         = window->_overlay.GetModel();
    out.presentation                = model.presentation;
    out.renderedIconPresentation    = window->_overlay.DebugGetLastDrawnIconPresentationForTest();
    out.renderedIconGlyph           = window->_overlay.DebugGetLastDrawnIconGlyphForTest();
    out.renderedCloseGlyph          = window->_overlay.DebugGetLastDrawnCloseGlyphForTest();
    out.title                       = model.title;
    out.optionValues.clear();
    out.optionValues.reserve(model.options.size());
    for (const AlertOption& option : model.options)
    {
        out.optionValues.push_back((! option.choices.empty() && option.selectedIndex < option.choices.size()) ? option.choices[option.selectedIndex].value
                                                                                                              : 0u);
    }
    out.bodyViewport     = window->_overlay.GetBodyViewport();
    out.scrollOffsetDip  = window->_overlay.GetScrollOffset();
    out.scrollMaximumDip = window->_overlay.GetScrollMaximum();
    out.visibleOptionRects.clear();
    out.optionIds.clear();
    for (const auto& option : model.options)
    {
        out.optionIds.push_back(option.id);
        out.visibleOptionRects.push_back(window->_overlay.GetVisibleControlRect({AlertHitTest::Part::Option, option.id}));
    }
    out.focusedControl = window->_overlay.GetFocusedControl();
    out.primaryButtonLabel.clear();
    for (const AlertButton& button : model.buttons)
    {
        if (button.primary)
        {
            out.primaryButtonLabel = button.label;
            break;
        }
    }
    return true;
}

bool DebugRunAlertOverlayIdentitySelfTest(HWND parent, AlertOverlayIdentitySelfTestResult& out) noexcept
{
    out = {};
    if (! parent || IsWindow(parent) == FALSE || GetWindowThreadProcessId(parent, nullptr) != GetCurrentThreadId())
        return false;

    struct CallbackState final
    {
        std::array<uint32_t, 4u> buttons{};
        size_t buttonCount = 0u;
    } callbacks;
    const AlertOverlayWindowCallbacks callbackSet{&callbacks,
                                                  [](void* context, uint32_t id) noexcept
    {
        auto& state = *static_cast<CallbackState*>(context);
        if (state.buttonCount < state.buttons.size())
            state.buttons[state.buttonCount++] = id;
    }};

    const auto readRuntimeId = [](IRawElementProviderFragment& provider) noexcept -> std::vector<LONG>
    {
        SAFEARRAY* array = nullptr;
        if (FAILED(provider.GetRuntimeId(&array)) || ! array)
            return {};
        const auto destroy = wil::scope_exit([&]() noexcept { SafeArrayDestroy(array); });
        LONG lower         = 0;
        LONG upper         = -1;
        if (FAILED(SafeArrayGetLBound(array, 1, &lower)) || FAILED(SafeArrayGetUBound(array, 1, &upper)) || upper < lower)
            return {};
        std::vector<LONG> result;
        result.reserve(static_cast<size_t>(upper - lower + 1));
        for (LONG index = lower; index <= upper; ++index)
        {
            LONG item = 0;
            if (FAILED(SafeArrayGetElement(array, &index, &item)))
                return {};
            result.push_back(item);
        }
        return result;
    };
    const auto makeFragment = [](AlertOverlayWindow& window, AlertOverlayUiaProvider::ElementId id) noexcept
    {
        wil::com_ptr_nothrow<IRawElementProviderFragment> provider;
        auto* created = new (std::nothrow) AlertOverlayUiaProvider(window._hwnd.get(), id, window._instanceToken, window._presentationGeneration);
        if (created)
            provider.attach(static_cast<IRawElementProviderFragment*>(created));
        return provider;
    };
    const auto makeRoot = [](AlertOverlayWindow& window) noexcept
    {
        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> provider;
        auto* created = new (std::nothrow)
            AlertOverlayUiaProvider(window._hwnd.get(), AlertOverlayUiaProvider::ElementId{}, window._instanceToken, window._presentationGeneration);
        if (created)
            provider.attach(static_cast<IRawElementProviderFragmentRoot*>(created));
        return provider;
    };
    const auto getInvoke = [](IRawElementProviderFragment& fragment, wil::com_ptr_nothrow<IInvokeProvider>& invoke) noexcept
    {
        wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
        wil::com_ptr_nothrow<IUnknown> pattern;
        return SUCCEEDED(fragment.QueryInterface(__uuidof(IRawElementProviderSimple), simple.put_void())) && simple &&
               SUCCEEDED(simple->GetPatternProvider(UIA_InvokePatternId, pattern.put())) && pattern && SUCCEEDED(pattern.query_to(invoke.put())) && invoke;
    };
    const auto dispatchQueued = [](HWND hwnd) noexcept
    {
        MSG message{};
        while (PeekMessageW(&message, hwnd, WndMsg::kAlertOverlayInvokeButton, WndMsg::kAlertOverlayInvokeOption, PM_REMOVE))
            static_cast<void>(DispatchMessageW(&message));
    };

    auto window             = std::make_unique<AlertOverlayWindow>();
    const auto detachWindow = wil::scope_exit([&]() noexcept
    {
        if (window)
            window->ApplyAttachmentState(nullptr, nullptr, false, false);
    });
    AlertTheme theme{};
    AlertModel first{};
    first.title = L"Identity A";
    first.options.push_back(AlertOption{.id = 101u, .label = L"Choice A", .choices = {{1u, L"One"}, {2u, L"Two"}}, .selectedIndex = 0u});
    first.options.push_back(AlertOption{.id = 111u, .label = L"Choice A2", .choices = {{11u, L"Eleven"}, {12u, L"Twelve"}}, .selectedIndex = 0u});
    first.buttons.push_back(AlertButton{.id = 200u, .label = L"Action A0", .primary = true});
    first.buttons.push_back(AlertButton{.id = 201u, .label = L"Action A1", .primary = false});
    if (FAILED(window->ShowForParentClient(parent, theme, std::move(first), false)))
        return false;
    window->SetCallbacks(callbackSet);

    const HWND firstHwnd = window->_hwnd.get();
    auto oldRoot         = makeRoot(*window);
    auto oldButton       = makeFragment(*window, {AlertOverlayUiaProvider::ElementKind::Button, 1u});
    auto oldOption       = makeFragment(*window, {AlertOverlayUiaProvider::ElementKind::Option, 0u});
    auto oldClose        = makeFragment(*window, {AlertOverlayUiaProvider::ElementKind::CloseButton, 0u});
    wil::com_ptr_nothrow<IInvokeProvider> oldButtonInvoke;
    wil::com_ptr_nothrow<IInvokeProvider> oldOptionInvoke;
    wil::com_ptr_nothrow<IInvokeProvider> oldCloseInvoke;
    if (! oldRoot || ! oldButton || ! oldOption || ! oldClose || ! getInvoke(*oldButton, oldButtonInvoke) || ! getInvoke(*oldOption, oldOptionInvoke) ||
        ! getInvoke(*oldClose, oldCloseInvoke))
        return false;
    wil::com_ptr_nothrow<IRawElementProviderFragment> oldRootFragment;
    if (FAILED(oldRoot.query_to(oldRootFragment.put())) || ! oldRootFragment)
        return false;
    const std::vector<LONG> firstRuntimeId = readRuntimeId(*oldRootFragment);
    if (FAILED(oldButtonInvoke->Invoke()) || FAILED(oldOptionInvoke->Invoke()) || FAILED(oldCloseInvoke->Invoke()))
        return false;

    window->ApplyAttachmentState(nullptr, nullptr, false, false);
    window->Hide();
    AlertModel second{};
    second.title    = L"Identity B";
    second.closable = false;
    second.options.push_back(AlertOption{.id = 102u, .label = L"Choice B", .choices = {{3u, L"Three"}, {4u, L"Four"}}, .selectedIndex = 0u});
    second.options.push_back(AlertOption{.id = 101u, .label = L"Choice A moved", .choices = {{1u, L"One"}, {2u, L"Two"}}, .selectedIndex = 0u});
    second.buttons.push_back(AlertButton{.id = 202u, .label = L"Action B", .primary = true});
    second.buttons.push_back(AlertButton{.id = 200u, .label = L"Action A0 moved", .primary = false});
    second.buttons.push_back(AlertButton{.id = 201u, .label = L"Action A1 moved", .primary = false});
    if (FAILED(window->ShowForParentClient(parent, theme, std::move(second), false)) || window->_hwnd.get() != firstHwnd)
        return false;
    window->SetCallbacks(callbackSet);

    wil::com_ptr_nothrow<IRawElementProviderFragment> staleChild;
    out.staleRootRejected   = oldRootFragment->Navigate(NavigateDirection_FirstChild, staleChild.put()) == UIA_E_ELEMENTNOTAVAILABLE;
    out.staleButtonRejected = oldButtonInvoke->Invoke() == UIA_E_ELEMENTNOTAVAILABLE;
    out.staleOptionRejected = oldOptionInvoke->Invoke() == UIA_E_ELEMENTNOTAVAILABLE;
    out.staleCloseRejected  = oldCloseInvoke->Invoke() == UIA_E_ELEMENTNOTAVAILABLE;
    dispatchQueued(firstHwnd);
    out.queuedButtonIgnoredAfterReplacement = callbacks.buttonCount == 0u;
    out.queuedOptionIgnoredAfterReplacement =
        window->GetOptionValue(102u) == std::optional<uint64_t>{3u} && window->GetOptionValue(101u) == std::optional<uint64_t>{1u};
    AlertOverlayWindowDebugSnapshot snapshot{};
    out.queuedCloseIgnoredAfterReplacement = DebugGetAlertOverlayWindowSnapshot(firstHwnd, snapshot) && snapshot.dismissCount == 0u;
    auto newRoot                           = makeRoot(*window);
    auto newButton                         = makeFragment(*window, {AlertOverlayUiaProvider::ElementKind::Button, 0u});
    auto newOption                         = makeFragment(*window, {AlertOverlayUiaProvider::ElementKind::Option, 0u});
    wil::com_ptr_nothrow<IRawElementProviderFragment> newRootFragment;
    wil::com_ptr_nothrow<IRawElementProviderSimple> newOptionSimple;
    wil::com_ptr_nothrow<IInvokeProvider> newButtonInvoke;
    wil::com_ptr_nothrow<IInvokeProvider> newOptionInvoke;
    if (! newRoot || ! newButton || ! newOption || FAILED(newRoot.query_to(newRootFragment.put())) || ! newRootFragment ||
        FAILED(newOption.query_to(newOptionSimple.put())) || ! newOptionSimple || ! getInvoke(*newButton, newButtonInvoke) ||
        ! getInvoke(*newOption, newOptionInvoke))
        return false;
    out.runtimeIdChanged              = ! firstRuntimeId.empty() && readRuntimeId(*newRootFragment) != firstRuntimeId;
    const HRESULT currentButtonResult = newButtonInvoke->Invoke();
    dispatchQueued(firstHwnd);
    out.currentButtonInvoked          = SUCCEEDED(currentButtonResult) && callbacks.buttonCount == 1u && callbacks.buttons[0] == 202u;
    const HRESULT currentOptionResult = newOptionInvoke->Invoke();
    dispatchQueued(firstHwnd);
    out.currentOptionInvoked = SUCCEEDED(currentOptionResult) && window->GetOptionValue(102u) == std::optional<uint64_t>{4u};
    wil::unique_variant updatedOptionName;
    const HRESULT updatedOptionProperty = newOptionSimple->GetPropertyValue(UIA_NamePropertyId, updatedOptionName.reset_and_addressof());
    out.currentProviderSurvivesSamePresentationUpdate =
        SUCCEEDED(updatedOptionProperty) && updatedOptionName.vt == VT_BSTR &&
        std::wstring_view(updatedOptionName.bstrVal, SysStringLen(updatedOptionName.bstrVal)) == L"Choice B: Four";

    const HWND foreignWindow =
        CreateWindowExW(0, L"STATIC", L"foreign HWND reuse witness", WS_POPUP, 0, 0, 1, 1, nullptr, nullptr, GetModuleHandleW(nullptr), nullptr);
    if (! foreignWindow)
        return false;
    const auto destroyForeignWindow = wil::scope_exit([&]() noexcept { DestroyWindow(foreignWindow); });
    SetWindowLongPtrW(foreignWindow, GWLP_USERDATA, static_cast<LONG_PTR>(1u));
    AlertOverlayUiaProvider foreignWindowProvider(
        foreignWindow, {AlertOverlayUiaProvider::ElementKind::Button, 0u}, window->_instanceToken, window->_presentationGeneration);
    wil::unique_variant foreignName;
    out.reusedForeignWindowRejected =
        foreignWindowProvider.GetPropertyValue(UIA_NamePropertyId, foreignName.reset_and_addressof()) == UIA_E_ELEMENTNOTAVAILABLE;

    wil::com_ptr_nothrow<IRawElementProviderFragment> providerAfterDestroy = newButton;
    const HRESULT queuedBeforeDestroy                                      = newButtonInvoke->Invoke();
    MSG drainedCandidate{};
    const bool hadQueuedCandidate =
        PeekMessageW(&drainedCandidate, firstHwnd, WndMsg::kAlertOverlayInvokeButton, WndMsg::kAlertOverlayInvokeButton, PM_REMOVE) != FALSE;
    window->ApplyAttachmentState(nullptr, nullptr, false, false);
    window.reset();
    auto replacement             = std::make_unique<AlertOverlayWindow>();
    const auto detachReplacement = wil::scope_exit([&]() noexcept
    {
        if (replacement)
            replacement->ApplyAttachmentState(nullptr, nullptr, false, false);
    });
    AlertModel third{};
    third.title = L"Identity C";
    third.buttons.push_back(AlertButton{.id = 301u, .label = L"Action C", .primary = true});
    if (FAILED(replacement->ShowForParentClient(parent, theme, std::move(third), false)))
        return false;
    replacement->SetCallbacks(callbackSet);
    wil::com_ptr_nothrow<IInvokeProvider> destroyedInvoke;
    out.destroyedProviderRejectedAfterReplacement = providerAfterDestroy &&
                                                    SUCCEEDED(providerAfterDestroy->QueryInterface(__uuidof(IInvokeProvider), destroyedInvoke.put_void())) &&
                                                    destroyedInvoke->Invoke() == UIA_E_ELEMENTNOTAVAILABLE;
    if (hadQueuedCandidate && SUCCEEDED(queuedBeforeDestroy))
    {
        static_cast<void>(SendMessageW(replacement->_hwnd.get(), drainedCandidate.message, drainedCandidate.wParam, drainedCandidate.lParam));
        out.drainedCommandIgnoredAfterReplacement = callbacks.buttonCount == 1u && callbacks.buttons[0] == 202u;
    }
    replacement->ApplyAttachmentState(nullptr, nullptr, false, false);
    replacement->Hide();
    return true;
}

uint64_t DebugGetAlertOverlayForegroundRequestCount() noexcept
{
    return g_debugForegroundRequestCount.load(std::memory_order_relaxed);
}

namespace
{
// A subclass another component installs on the anchor after the overlays' hook. UI thread only.
WNDPROC g_debugForeignBaseProc      = nullptr;
uint64_t g_debugForeignMessageCount = 0u;

LRESULT CALLBACK DebugForeignSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ++g_debugForeignMessageCount;
    return Win32Callback::CallWindowProcNoThrow(g_debugForeignBaseProc, hwnd, msg, wp, lp);
}
} // namespace

bool DebugRunAlertOverlayTrackingSelfTest(HWND placement, AlertOverlayTrackingSelfTestResult& out) noexcept
{
    out = {};
    RECT placementRect{};
    if (placement && IsWindow(placement) != FALSE)
        static_cast<void>(GetWindowRect(placement, &placementRect));
    const int originX               = placementRect.left + 48;
    const int originY               = placementRect.top + 48;
    const uint64_t foregroundBefore = DebugGetAlertOverlayForegroundRequestCount();
    const auto createAnchor         = [&]() noexcept
    { return CreateWindowExW(0, L"STATIC", L"", WS_POPUP | WS_VISIBLE, originX, originY, 360, 240, nullptr, nullptr, GetModuleHandleW(nullptr), nullptr); };
    const auto makeModel = [](const wchar_t* title)
    {
        AlertModel model{};
        model.title    = title;
        model.message  = L"Alert overlay tracking self-test";
        model.closable = true;
        return model;
    };
    const AlertTheme theme{};

    wil::unique_hwnd anchor(createAnchor());
    wil::unique_hwnd secondAnchor(createAnchor());
    if (! anchor || ! secondAnchor)
        return false;
    const auto originalProc  = reinterpret_cast<WNDPROC>(GetWindowLongPtrW(anchor.get(), GWLP_WNDPROC));
    const auto foreignProc   = &DebugForeignSubclassProc;
    const auto removeForeign = [&]() noexcept
    {
        if (anchor && reinterpret_cast<WNDPROC>(GetWindowLongPtrW(anchor.get(), GWLP_WNDPROC)) == foreignProc)
            static_cast<void>(Win32Callback::SetWindowLongPtrNoThrow(anchor.get(), GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(g_debugForeignBaseProc)));
    };
    const auto removeForeignOnExit = wil::scope_exit([&]() noexcept { removeForeign(); });
    const auto coversAnchorClient  = [&](const AlertOverlayWindow& overlay) noexcept
    {
        RECT client{};
        RECT overlayRect{};
        if (! overlay._hwnd || GetClientRect(anchor.get(), &client) == 0 || GetWindowRect(overlay._hwnd.get(), &overlayRect) == 0)
            return false;
        POINT corners[2]{{client.left, client.top}, {client.right, client.bottom}};
        MapWindowPoints(anchor.get(), nullptr, corners, 2);
        return overlayRect.left == corners[0].x && overlayRect.top == corners[0].y && overlayRect.right == corners[1].x && overlayRect.bottom == corners[1].y;
    };

    // A window-scope alert and a later overlay follow the same anchor, and another component then
    // subclasses it. Removing the first overlay must leave the second tracking and the subclass in place.
    auto first  = std::make_unique<AlertOverlayWindow>();
    auto second = std::make_unique<AlertOverlayWindow>();
    if (FAILED(first->ShowForAnchor(anchor.get(), theme, makeModel(L"First"), false)) ||
        FAILED(second->ShowForAnchor(anchor.get(), theme, makeModel(L"Second"), false)))
        return false;
    g_debugForeignMessageCount = 0u;
    g_debugForeignBaseProc =
        reinterpret_cast<WNDPROC>(Win32Callback::SetWindowLongPtrNoThrow(anchor.get(), GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(foreignProc)));
    if (! g_debugForeignBaseProc)
        return false;
    first.reset();
    SetWindowPos(anchor.get(), nullptr, originX + 16, originY + 8, 420, 280, SWP_NOZORDER | SWP_NOACTIVATE);
    out.siblingDetachKeepsTracking = second->IsVisible() && coversAnchorClient(*second);

    second.reset();
    const bool foreignStillCurrent = reinterpret_cast<WNDPROC>(GetWindowLongPtrW(anchor.get(), GWLP_WNDPROC)) == foreignProc;
    const uint64_t foreignBefore   = g_debugForeignMessageCount;
    std::array<wchar_t, 8u> text{};
    const bool textRoundTrip = SetWindowTextW(anchor.get(), L"chain") != FALSE &&
                               GetWindowTextW(anchor.get(), text.data(), static_cast<int>(text.size())) == 5 && std::wstring_view(text.data()) == L"chain";
    out.laterSubclassKept    = foreignStillCurrent && textRoundTrip && g_debugForeignMessageCount > foreignBefore;

    // Once the subclass is gone, the last overlay to detach puts the original procedure back.
    removeForeign();
    auto third = std::make_unique<AlertOverlayWindow>();
    if (FAILED(third->ShowForAnchor(anchor.get(), theme, makeModel(L"Third"), false)))
        return false;
    third.reset();
    out.originalProcedureRestored =
        reinterpret_cast<WNDPROC>(GetWindowLongPtrW(anchor.get(), GWLP_WNDPROC)) == originalProc && GetPropW(anchor.get(), kTrackedWindowHookProp) == nullptr;

    // The owner dies while no tracking hook sees it (as when a hook was lost): the overlay window is
    // destroyed with its owner and must stop reporting visible.
    auto orphan = std::make_unique<AlertOverlayWindow>();
    if (FAILED(orphan->ShowForAnchor(anchor.get(), theme, makeModel(L"Orphan"), false)))
        return false;
    orphan->ApplyAttachmentState(orphan->_hostParent, nullptr, false, false);
    anchor.reset();
    out.ownerDestroyedClearsVisibility = ! orphan->IsVisible() && ! orphan->_hwnd;
    orphan.reset();

    // Close (Escape) on a closable alert without an escape binding just hides it.
    auto closable = std::make_unique<AlertOverlayWindow>();
    if (FAILED(closable->ShowForAnchor(secondAnchor.get(), theme, makeModel(L"Closable"), false)) || ! closable->_hwnd)
        return false;
    static_cast<void>(SendMessageW(closable->_hwnd.get(), WM_KEYDOWN, VK_ESCAPE, 0));
    out.closeHidesWithoutEscapeBinding = ! closable->IsVisible();

    // The owner learns that the target died and deletes the overlay inside the notification, while a
    // second (hidden) overlay on the same target handles the same WM_NCDESTROY.
    auto owned = std::make_unique<AlertOverlayWindow>();
    struct DestroyedState final
    {
        std::unique_ptr<AlertOverlayWindow>* owned = nullptr;
        bool reported                              = false;
    } destroyed{&owned};
    if (FAILED(owned->ShowForAnchor(secondAnchor.get(), theme, makeModel(L"Owned"), false)))
        return false;
    owned->SetTrackedWindowDestroyedCallback(&destroyed,
                                             [](void* context, AlertOverlayWindow& overlay) noexcept
    {
        auto& state    = *static_cast<DestroyedState*>(context);
        state.reported = state.owned->get() == &overlay;
        state.owned->reset();
    });
    secondAnchor.reset();
    out.targetDestroyedCallbackReported = destroyed.reported && ! owned && ! closable->IsVisible();
    closable.reset();

    out.foregroundRequests = DebugGetAlertOverlayForegroundRequestCount() - foregroundBefore;
    return true;
}

bool DebugRunAlertOverlayUxAccessibilitySelfTest(HWND placement, AlertOverlayUxAccessibilitySelfTestResult& out) noexcept
{
    out = {};
    RECT placementRect{};
    UINT dpi = USER_DEFAULT_SCREEN_DPI;
    if (placement && IsWindow(placement) != FALSE)
    {
        static_cast<void>(GetWindowRect(placement, &placementRect));
        dpi = GetDpiForWindow(placement);
    }
    const uint64_t foregroundBefore = DebugGetAlertOverlayForegroundRequestCount();

    g_debugFocusOverrideActive = true;
    g_debugFocusOverride       = nullptr;
    g_debugUiaEventCapture     = true;
    g_debugUiaEvents.clear();
    const auto resetSeams = wil::scope_exit([]() noexcept
    {
        g_debugFocusOverrideActive = false;
        g_debugFocusOverride       = nullptr;
        g_debugUiaEventCapture     = false;
        g_debugUiaEvents.clear();
    });

    // Wide enough for the 80 DIP prompt icon (a 600 DIP panel), short enough that the body scrolls.
    wil::unique_hwnd anchor(CreateWindowExW(0,
                                            L"STATIC",
                                            L"",
                                            WS_POPUP | WS_VISIBLE,
                                            placementRect.left + 48,
                                            placementRect.top + 48,
                                            MulDiv(720, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI),
                                            MulDiv(300, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI),
                                            nullptr,
                                            nullptr,
                                            GetModuleHandleW(nullptr),
                                            nullptr));
    if (! anchor)
        return false;

    struct PressedButtons final
    {
        std::array<uint32_t, 8u> ids{};
        size_t count = 0u;
    } pressed;
    const AlertOverlayWindowCallbacks callbacks{&pressed,
                                                [](void* context, uint32_t id) noexcept
    {
        auto& state = *static_cast<PressedButtons*>(context);
        if (state.count < state.ids.size())
            state.ids[state.count++] = id;
    }};
    const AlertTheme theme{};
    auto window     = std::make_unique<AlertOverlayWindow>();
    const auto show = [&](AlertModel model, bool blocksInput) noexcept
    {
        if (FAILED(window->ShowForAnchor(anchor.get(), theme, std::move(model), blocksInput)) || ! window->_hwnd)
            return false;
        window->SetCallbacks(callbacks);
        return window->EnsureOverlayLayoutForInput();
    };
    const auto addButtons = [](AlertModel& model, const wchar_t* first, const wchar_t* second)
    {
        model.buttons.push_back(AlertButton{.id = 1u, .label = first, .primary = true});
        model.buttons.push_back(AlertButton{.id = 2u, .label = second, .primary = false});
    };
    const auto nameOf = [&](AlertOverlayUiaProvider::ElementId element) noexcept -> std::wstring
    {
        const auto provider = AlertOverlayUiaProvider::ForCurrentPresentation(*window, element);
        wil::unique_variant name;
        std::wstring result;
        if (provider && SUCCEEDED(provider->GetPropertyValue(UIA_NamePropertyId, name.reset_and_addressof())) && name.vt == VT_BSTR && name.bstrVal)
            result.assign(name.bstrVal, SysStringLen(name.bstrVal));
        return result;
    };
    const auto raised = [&](EVENTID eventId, AlertOverlayUiaProvider::ElementId element) noexcept
    {
        return std::ranges::any_of(g_debugUiaEvents,
                                   [&](const DebugUiaEventRecord& record) noexcept
        {
            return record.eventId == eventId && record.elementKind == static_cast<uint8_t>(element.kind) && record.elementIndex == element.index &&
                   record.generation == window->_presentationGeneration && SUCCEEDED(record.result);
        });
    };
    const auto raisedAny = [](EVENTID eventId) noexcept
    { return std::ranges::any_of(g_debugUiaEvents, [eventId](const DebugUiaEventRecord& record) noexcept { return record.eventId == eventId; }); };
    using Kind = AlertOverlayUiaProvider::ElementKind;

    // The icon of an over-tall prompt is centred on the visible body, not on the whole content.
    AlertModel tall{};
    tall.severity = AlertSeverity::Warning;
    tall.title    = L"Icon placement";
    for (int line = 1; line <= 60; ++line)
        tall.message.append(std::format(L"Line {} of a confirmation body taller than the alert.\n", line));
    addButtons(tall, L"OK", L"Cancel");
    if (! show(std::move(tall), false))
        return false;
    out.iconRect              = window->_overlay.DebugGetIconRectForTest();
    out.bodyViewport          = window->_overlay.GetBodyViewport();
    out.iconFixtureScrolls    = window->_overlay.GetScrollMaximum() > 0.0f && out.iconRect.bottom > out.iconRect.top;
    const float iconCentre    = (out.iconRect.top + out.iconRect.bottom) * 0.5f;
    const float bodyCentre    = (out.bodyViewport.top + out.bodyViewport.bottom) * 0.5f;
    out.iconInsideVisibleBody = out.iconFixtureScrolls && out.iconRect.top >= out.bodyViewport.top - 0.5f &&
                                out.iconRect.bottom <= out.bodyViewport.bottom + 0.5f && std::abs(iconCentre - bodyCentre) <= 1.0f;

    // Activating a visible option of a scrolling alert leaves the content where it is, so a second
    // click at the same point changes the same option.
    AlertModel settings{};
    settings.title   = L"Option activation";
    settings.message = L"Twelve settings below a short body.";
    for (uint32_t index = 0; index < 12u; ++index)
        settings.options.push_back(
            AlertOption{.id = 400u + index, .label = std::format(L"Setting {}", index), .choices = {{0u, L"Off"}, {1u, L"On"}}, .selectedIndex = 0u});
    addButtons(settings, L"OK", L"Cancel");
    if (! show(std::move(settings), false))
        return false;
    constexpr uint32_t kTarget = 401u;
    const AlertHitTest target{AlertHitTest::Part::Option, kTarget};
    const auto keepsPlace = [&](float offset, D2D1_RECT_F rect) noexcept
    {
        const bool layout = window->EnsureOverlayLayoutForInput();
        const auto now    = window->_overlay.GetVisibleControlRect(target);
        return layout && window->_overlay.GetScrollOffset() == offset && std::abs(now.top - rect.top) <= 0.5f && std::abs(now.bottom - rect.bottom) <= 0.5f;
    };
    const D2D1_RECT_F viewport = window->_overlay.GetBodyViewport();
    const D2D1_RECT_F initial  = window->_overlay.GetVisibleControlRect(target);
    out.optionScrollBeforeDip  = window->_overlay.GetScrollOffset();
    out.optionFixtureScrolls   = window->_overlay.GetScrollMaximum() > 0.0f && out.optionScrollBeforeDip == 0.0f && initial.top > viewport.top + 8.0f &&
                                 initial.bottom < viewport.bottom - 8.0f;
    window->InvokeOption(kTarget); // pointer click and UIA Invoke
    out.optionScrollAfterDip         = window->_overlay.GetScrollOffset();
    out.pointerActivationKeepsScroll = window->GetOptionValue(kTarget) == std::optional<uint64_t>{1u} && keepsPlace(0.0f, initial);
    static_cast<void>(window->_overlay.ScrollBy(24.0f));
    const float scrolled           = window->_overlay.GetScrollOffset();
    const D2D1_RECT_F scrolledRect = window->_overlay.GetVisibleControlRect(target);
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_KEYDOWN, VK_RIGHT, 0));
    out.arrowActivationKeepsScroll = scrolled > 0.0f && scrolledRect.top > viewport.top + 8.0f &&
                                     window->GetOptionValue(kTarget) == std::optional<uint64_t>{0u} && keepsPlace(scrolled, scrolledRect);
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_KEYDOWN, VK_SPACE, 0));
    out.spaceActivationKeepsScroll = window->GetOptionValue(kTarget) == std::optional<uint64_t>{1u} && keepsPlace(scrolled, scrolledRect);

    // Focus moves inside a focused alert raise UIA focus events; the option keeps the focus here.
    g_debugFocusOverride = window->_hwnd.get();
    g_debugUiaEvents.clear();
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_KEYDOWN, VK_TAB, 0));
    out.tabRaisesFocusEvent = raised(UIA_AutomationFocusChangedEventId, {Kind::Option, 2u});
    g_debugUiaEvents.clear();
    const auto firstButton = AlertOverlayUiaProvider::ForCurrentPresentation(*window, {Kind::Button, 0u});
    wil::com_ptr_nothrow<IRawElementProviderFragment> firstButtonFragment;
    out.uiaSetFocusRaisesFocusEvent = firstButton && SUCCEEDED(firstButton.query_to(firstButtonFragment.put())) && SUCCEEDED(firstButtonFragment->SetFocus()) &&
                                      raised(UIA_AutomationFocusChangedEventId, {Kind::Button, 0u});
    g_debugUiaEvents.clear();
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_KEYDOWN, VK_RIGHT, 0));
    out.arrowRaisesFocusEvent = raised(UIA_AutomationFocusChangedEventId, {Kind::Button, 1u});
    static_cast<void>(window->OnFocusControl(AlertHitTest::Part::Option, 400u));
    g_debugUiaEvents.clear();
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_KEYDOWN, VK_SPACE, 0));
    out.optionCycleRaisesNameChange = raised(UIA_AutomationPropertyChangedEventId, {Kind::Option, 0u}) && ! raisedAny(UIA_AutomationFocusChangedEventId);
    g_debugFocusOverride            = anchor.get();
    g_debugUiaEvents.clear();
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_KEYDOWN, VK_TAB, 0));
    out.unfocusedMoveRaisesNoFocusEvent = window->_overlay.GetFocusedControl().buttonId == 401u && ! raisedAny(UIA_AutomationFocusChangedEventId);

    // A new model on the visible, focused alert retires the old elements: say so and re-announce.
    g_debugFocusOverride = window->_hwnd.get();
    g_debugUiaEvents.clear();
    AlertModel replaced{};
    replaced.title   = L"Replaced alert";
    replaced.message = L"The alert now shows other text.";
    addButtons(replaced, L"Retry", L"Cancel");
    if (! show(std::move(replaced), false))
        return false;
    out.replacementRaisesStructureChange = raised(UIA_StructureChangedEventId, {Kind::Root, 0u});
    out.replacementAnnouncesText         = raised(UIA_NotificationEventId, {Kind::Root, 0u});
    out.replacementRaisesFocusEvent      = raised(UIA_AutomationFocusChangedEventId, {Kind::Button, 0u});
    window->Hide();

    // A blocking alert re-shown while it has the focus keeps the focus it first captured, and a
    // stale focus on the hidden overlay is never captured.
    AlertModel blocking{};
    blocking.title   = L"Blocking alert";
    blocking.message = L"Restore focus witness.";
    addButtons(blocking, L"OK", L"Cancel");
    g_debugFocusOverride        = anchor.get();
    const bool capturedFirst    = show(blocking, true) && window->_restoreFocus == anchor.get();
    g_debugFocusOverride        = window->_hwnd.get();
    out.reshowKeepsRestoreFocus = capturedFirst && show(blocking, true) && window->_restoreFocus == anchor.get();
    window->Hide();
    out.overlayNeverRecordedAsRestoreFocus = show(blocking, true) && window->_restoreFocus == nullptr;
    window->Hide();

    // Alt mnemonics: only '&' markers count, "&&" is a literal '&', a shared key selects nothing.
    AlertModel unmarked{};
    unmarked.title = L"Mnemonics";
    addButtons(unmarked, L"Copy", L"Cancel");
    pressed                          = {};
    out.unmarkedLabelsHaveNoMnemonic = show(std::move(unmarked), false) && ! window->OnSysChar(L'c') && ! window->OnSysChar(L'C') && pressed.count == 0u;
    AlertModel marked{};
    marked.title = L"Mnemonics";
    addButtons(marked, L"&Copy", L"C&ancel");
    marked.buttons.push_back(AlertButton{.id = 3u, .label = L"Save && &Quit", .primary = false});
    pressed                = {};
    const bool markedShown = show(std::move(marked), false);
    const bool markedKeys  = window->OnSysChar(L'a') && window->OnSysChar(L'C') && window->OnSysChar(L'q') && ! window->OnSysChar(L's');
    out.markedMnemonicInvokesItsButton =
        markedShown && markedKeys && pressed.count == 3u && pressed.ids[0] == 2u && pressed.ids[1] == 1u && pressed.ids[2] == 3u;
    out.markersHiddenFromName = markedShown && nameOf({Kind::Button, 0u}) == L"Copy" && nameOf({Kind::Button, 2u}) == L"Save & Quit";
    AlertModel ambiguous{};
    ambiguous.title = L"Mnemonics";
    addButtons(ambiguous, L"&Save", L"&Skip");
    pressed                      = {};
    out.ambiguousMnemonicIgnored = show(std::move(ambiguous), false) && ! window->OnSysChar(L's') && pressed.count == 0u;

    // The untitled alert and its close button are named from string resources.
    AlertModel untitled{};
    untitled.message  = L"An alert without a title.";
    untitled.closable = true;
    if (! show(std::move(untitled), false))
        return false;
    out.untitledRootName = nameOf({Kind::Root, 0u});
    out.closeName        = nameOf({Kind::CloseButton, 0u});

    // A message carrying a raw call pointer, on the retired WM_APP id or the registered one, runs nothing.
    const UINT uiaCallMessage = AlertOverlayUiaCallMessage();
    uint32_t forgedRuns       = 0u;
    AlertUiaCall forged{&forgedRuns,
                        [](void* context) noexcept
    {
        ++*static_cast<uint32_t*>(context);
        return S_OK;
    }};
    const LPARAM forgedPointer = reinterpret_cast<LPARAM>(&forged);
    const LRESULT retiredReply = SendMessageW(window->_hwnd.get(), WM_APP + 0x76, 0, forgedPointer);
    const LRESULT forgedReply  = uiaCallMessage != 0u ? SendMessageW(window->_hwnd.get(), uiaCallMessage, 0, forgedPointer) : TRUE;
    out.forgedUiaCallIgnored   = uiaCallMessage != 0u && retiredReply == 0 && forgedReply == 0 && forgedRuns == 0u;

    // A provider called from another thread still runs on the overlay's thread, which pumps sent messages here.
    wil::unique_event_nothrow workerDone;
    if (const auto root = AlertOverlayUiaProvider::ForCurrentPresentation(*window, {Kind::Root, 0u});
        root && SUCCEEDED(workerDone.create(wil::EventOptions::ManualReset)))
    {
        HRESULT workerResult = E_PENDING;
        std::wstring workerName;
        {
            std::jthread worker([&]() noexcept
            {
                wil::unique_variant name;
                workerResult = root->GetPropertyValue(UIA_NamePropertyId, name.reset_and_addressof());
                if (SUCCEEDED(workerResult) && name.vt == VT_BSTR && name.bstrVal)
                    workerName.assign(name.bstrVal, SysStringLen(name.bstrVal));
                workerDone.SetEvent();
            });
            const HANDLE done = workerDone.get();
            while (MsgWaitForMultipleObjectsEx(1u, &done, INFINITE, QS_SENDMESSAGE, 0u) != WAIT_OBJECT_0)
            {
                MSG sent{};
                static_cast<void>(PeekMessageW(&sent, nullptr, 0u, 0u, PM_NOREMOVE | PM_QS_SENDMESSAGE));
            }
        }
        out.crossThreadUiaCallMarshals = SUCCEEDED(workerResult) && ! workerName.empty() && workerName == out.untitledRootName;
    }

    // A synchronous payload nobody takes is reclaimed and deleted before SendMessagePayload returns.
    struct ReclaimWitness final
    {
        explicit ReclaimWitness(bool& flag) noexcept : destroyed(flag)
        {
        }
        ReclaimWitness(const ReclaimWitness&)            = delete;
        ReclaimWitness& operator=(const ReclaimWitness&) = delete;
        ReclaimWitness(ReclaimWitness&&)                 = delete;
        ReclaimWitness& operator=(ReclaimWitness&&)      = delete;
        ~ReclaimWitness()
        {
            destroyed = true;
        }
        bool& destroyed;
    };
    bool witnessDestroyed = false;
    const LRESULT ignoredReply =
        uiaCallMessage != 0u ? SendMessagePayload(anchor.get(), uiaCallMessage, 0, std::make_unique<ReclaimWitness>(witnessDestroyed)) : TRUE;
    out.untakenSyncPayloadReclaimed = ignoredReply == 0 && witnessDestroyed;
    window->ClearCallbacks();
    window->Hide();
    window.reset();

    out.foregroundRequests = DebugGetAlertOverlayForegroundRequestCount() - foregroundBefore;
    return true;
}

void DebugSetAlertOverlayFocusOverride(std::optional<HWND> focus) noexcept
{
    g_debugFocusOverrideActive = focus.has_value();
    g_debugFocusOverride       = focus.value_or(nullptr);
}

void DebugSetAlertOverlayKeyMessageTimeOverride(std::optional<DWORD> messageTime) noexcept
{
    g_debugKeyMessageTimeActive = messageTime.has_value();
    g_debugKeyMessageTime       = messageTime.value_or(0u);
}

void DebugSetAlertOverlayForegroundOverride(std::optional<HWND> foreground) noexcept
{
    g_debugForegroundOverrideActive = foreground.has_value();
    g_debugForegroundOverride       = foreground.value_or(nullptr);
}

namespace
{
[[nodiscard]] AlertOverlayWindow* DebugOverlayFromHwnd(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE || GetWindowThreadProcessId(hwnd, nullptr) != GetCurrentThreadId())
        return nullptr;
    wchar_t className[64]{};
    if (GetClassNameW(hwnd, className, static_cast<int>(std::size(className))) <= 0 || std::wstring_view(className) != kAlertOverlayWindowClassName)
        return nullptr;
    auto* window = reinterpret_cast<AlertOverlayWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window && window->GetHwnd() == hwnd ? window : nullptr;
}
} // namespace

bool DebugInvokeAlertOverlayButton(HWND hwnd, uint32_t buttonId) noexcept
{
    AlertOverlayWindow* window = DebugOverlayFromHwnd(hwnd);
    if (! window || ! window->_visible)
        return false;
    window->InvokeButton(buttonId);
    return true;
}

bool DebugInvokeAlertOverlayClose(HWND hwnd) noexcept
{
    AlertOverlayWindow* window = DebugOverlayFromHwnd(hwnd);
    if (! window || ! window->_visible)
        return false;
    window->InvokeDismiss(AlertOverlayWindow::DismissSource::Close);
    return true;
}

bool DebugInvokeAlertOverlayOption(HWND hwnd, uint32_t optionId) noexcept
{
    AlertOverlayWindow* window = DebugOverlayFromHwnd(hwnd);
    if (! window || ! window->_visible)
        return false;
    window->InvokeOption(optionId);
    return true;
}

bool DebugGetAlertOverlayModelessSnapshot(HWND hwnd, AlertOverlayModelessDebugSnapshot& out) noexcept
{
    out                        = {};
    AlertOverlayWindow* window = DebugOverlayFromHwnd(hwnd);
    if (! window)
        return false;
    out.visible                = window->_visible;
    out.windowVisible          = IsWindowVisible(hwnd) != FALSE;
    out.keysRequireFocus       = window->_keysRequireFocus;
    out.restoresFocus          = window->_modelessRestoresFocus;
    out.restoreFocus           = window->_restoreFocus;
    out.anchor                 = window->_anchor;
    out.lastAnnouncement       = window->_debugLastAnnouncement;
    out.primaryButtonId        = window->_primaryButtonId;
    out.escapeButtonId         = window->_escapeButtonId;
    out.focusedButtonId        = window->_overlay.GetFocusedButtonId();
    out.focusedClose           = window->_overlay.GetFocusedControl().part == AlertHitTest::Part::Close;
    out.guardsKeyboardFocus    = window->_guardKeyboardFocus;
    out.keysGuarded            = window->_keysGuarded;
    out.guardArmedTick         = window->_guardArmedTick;
    out.presentationGeneration = window->_presentationGeneration;
    const AlertModel& model    = window->_overlay.GetModel();
    out.title                  = model.title;
    out.message                = model.message;
    for (const AlertButton& button : model.buttons)
    {
        out.buttonIds.push_back(button.id);
        out.buttonLabels.push_back(AlertButtonDisplayLabel(button.label));
    }
    for (const AlertOption& option : model.options)
    {
        out.optionIds.push_back(option.id);
        out.optionValues.push_back(window->_overlay.GetOptionValue(option.id).value_or(0u));
    }
    if (const auto root = AlertOverlayUiaProvider::ForCurrentPresentation(*window, AlertOverlayUiaProvider::ElementId{}))
    {
        wil::unique_variant name;
        if (SUCCEEDED(root->GetPropertyValue(UIA_NamePropertyId, name.reset_and_addressof())) && name.vt == VT_BSTR && name.bstrVal)
            out.rootUiaName.assign(name.bstrVal, SysStringLen(name.bstrVal));
    }
    return true;
}

bool DebugRunAlertOverlayModelessSelfTest(HWND placement, AlertOverlayModelessSelfTestResult& out) noexcept
{
    out = {};
    RECT placementRect{};
    UINT dpi = USER_DEFAULT_SCREEN_DPI;
    if (placement && IsWindow(placement) != FALSE)
    {
        static_cast<void>(GetWindowRect(placement, &placementRect));
        dpi = GetDpiForWindow(placement);
    }
    const uint64_t foregroundBefore = DebugGetAlertOverlayForegroundRequestCount();

    g_debugFocusOverrideActive = true;
    g_debugFocusOverride       = nullptr;
    g_debugUiaEventCapture     = true;
    g_debugUiaEvents.clear();
    const auto resetSeams = wil::scope_exit([]() noexcept
    {
        g_debugFocusOverrideActive      = false;
        g_debugFocusOverride            = nullptr;
        g_debugUiaEventCapture          = false;
        g_debugKeyMessageTimeActive     = false;
        g_debugKeyMessageTime           = 0u;
        g_debugForegroundOverrideActive = false;
        g_debugForegroundOverride       = nullptr;
        g_debugUiaEvents.clear();
    });
    // Held keys and modifiers are stood in the thread's keyboard state, which is given back at the end.
    std::array<BYTE, 256> keyboardBefore{};
    const bool keyboardSaved   = GetKeyboardState(keyboardBefore.data()) != FALSE;
    const auto restoreKeyboard = wil::scope_exit([&]() noexcept
    {
        if (keyboardSaved)
            static_cast<void>(SetKeyboardState(keyboardBefore.data()));
    });
    const auto setKeyDown      = [](int virtualKey, bool down) noexcept
    {
        std::array<BYTE, 256> keyboard{};
        if (GetKeyboardState(keyboard.data()) == FALSE)
            return false;
        BYTE& state = keyboard[static_cast<size_t>(virtualKey)];
        state       = down ? static_cast<BYTE>(state | 0x80u) : static_cast<BYTE>(state & 0x7Fu);
        return SetKeyboardState(keyboard.data()) != FALSE;
    };

    // A top-level window holding a child "Folder window" with two child "panes", like the main window:
    // the overlay is owned by the top-level window and anchored to a pane two levels below it.
    const int scale  = static_cast<int>(dpi);
    const int left   = placementRect.left + 48;
    const int top    = placementRect.top + 48;
    const int width  = MulDiv(720, scale, USER_DEFAULT_SCREEN_DPI);
    const int height = MulDiv(420, scale, USER_DEFAULT_SCREEN_DPI);
    wil::unique_hwnd root(CreateWindowExW(
        0, L"STATIC", L"", WS_POPUP | WS_VISIBLE | WS_CLIPCHILDREN, left, top, width, height, nullptr, nullptr, GetModuleHandleW(nullptr), nullptr));
    if (! root)
        return false;
    wil::unique_hwnd host(CreateWindowExW(
        0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN, 0, 0, width, height, root.get(), nullptr, GetModuleHandleW(nullptr), nullptr));
    if (! host)
        return false;
    const auto createPane = [&](int x) noexcept
    {
        return CreateWindowExW(
            0, L"STATIC", L"", WS_CHILD | WS_VISIBLE, x, 8, width / 2 - 16, height - 16, host.get(), nullptr, GetModuleHandleW(nullptr), nullptr);
    };
    wil::unique_hwnd pane(createPane(8));
    wil::unique_hwnd otherPane(createPane(width / 2 + 8));
    if (! pane || ! otherPane)
        return false;

    struct CallbackState final
    {
        std::array<uint32_t, 8u> buttons{};
        size_t buttonCount    = 0u;
        uint32_t closes       = 0u;
        uint32_t optionId     = 0u;
        uint64_t optionValue  = 0u;
        uint32_t optionCalls  = 0u;
        uint32_t focusReturns = 0u;
    } pressed;
    const AlertOverlayWindowCallbacks callbacks{&pressed,
                                                [](void* context, uint32_t id) noexcept
    {
        auto& state = *static_cast<CallbackState*>(context);
        if (state.buttonCount < state.buttons.size())
            state.buttons[state.buttonCount++] = id;
    },
                                                [](void* context) noexcept { ++static_cast<CallbackState*>(context)->closes; },
                                                [](void* context, uint32_t optionId, uint64_t value) noexcept
    {
        auto& state       = *static_cast<CallbackState*>(context);
        state.optionId    = optionId;
        state.optionValue = value;
        ++state.optionCalls;
    },
                                                [](void* context) noexcept { ++static_cast<CallbackState*>(context)->focusReturns; }};
    constexpr uint32_t kDestructiveId = 10u;
    constexpr uint32_t kDefaultId     = 11u;
    constexpr uint32_t kEscapeId      = 99u; // bound to Escape, never drawn
    constexpr uint32_t kOptionId      = 50u;
    const auto makeModel              = [](bool withOption)
    {
        AlertModel model{};
        model.severity = AlertSeverity::Warning;
        model.title    = L"Copy";
        model.message  = L"An item with this name already exists.";
        model.closable = true;
        if (withOption)
            model.options.push_back(AlertOption{.id = kOptionId, .label = L"Apply to all", .choices = {{0u, L"Off"}, {1u, L"On"}}, .selectedIndex = 0u});
        model.buttons.push_back(AlertButton{.id = kDestructiveId, .label = L"Replace", .primary = false});
        model.buttons.push_back(AlertButton{.id = kDefaultId, .label = L"Cancel", .primary = true});
        return model;
    };
    const AlertTheme theme{};
    auto window     = std::make_unique<AlertOverlayWindow>();
    const auto show = [&](bool takeFocus, std::wstring announcement, bool withOption = false) noexcept
    {
        pressed = {};
        window->SetCallbacks(callbacks);
        window->SetKeyBindings(kDefaultId, kEscapeId);
        AlertModelessShowOptions options{};
        options.takeFocus              = takeFocus;
        options.guardKeyboardFocus     = true;
        options.keysRequireFocus       = true;
        options.followAnchorVisibility = true;
        options.announcement           = std::move(announcement);
        return SUCCEEDED(window->ShowModelessForAnchor(pane.get(), theme, makeModel(withOption), std::move(options))) && window->_hwnd;
    };
    // A key the user types `afterArmMs` after the guard was armed (negative: typed before it).
    const auto key = [&](UINT message, WPARAM vk, LONG afterArmMs, bool repeat = false) noexcept
    {
        g_debugKeyMessageTimeActive = true;
        g_debugKeyMessageTime       = window->_guardArmedTick + static_cast<DWORD>(afterArmMs);
        static_cast<void>(SendMessageW(window->_hwnd.get(), message, vk, repeat ? kKeyPreviouslyDownFlag : 0));
        g_debugKeyMessageTimeActive = false;
    };
    constexpr LONG kLate = static_cast<LONG>(kModelessKeyGuardTypeaheadMs) + 100;

    // 1. Shown while the pane had the focus: takes it, announces, and gives it back when it hides. The
    //    focus it took by itself is guarded: keys meant for the pane never reach the destructive button.
    g_debugFocusOverride = pane.get();
    g_debugUiaEvents.clear();
    if (! show(true, L"Waiting for your decision. An item with this name already exists."))
        return false;
    out.takeFocusCapturesAnchor = window->_modelessRestoresFocus && window->_restoreFocus == pane.get();
    out.takeFocusMovesFocus     = g_debugFocusOverride == window->_hwnd.get();
    out.firstShowAnnounces      = std::ranges::any_of(g_debugUiaEvents,
                                                      [&](const DebugUiaEventRecord& record) noexcept
    {
        return record.eventId == UIA_NotificationEventId && record.elementKind == static_cast<uint8_t>(AlertOverlayUiaProvider::ElementKind::Root) &&
               SUCCEEDED(record.result);
    }) && window->_debugLastAnnouncement == L"Waiting for your decision. An item with this name already exists.";
    // Tab, the arrows and Space (the pane's own navigation keys) are inert; the focus stays on the default.
    key(WM_KEYDOWN, VK_TAB, kLate);
    key(WM_KEYDOWN, VK_LEFT, kLate);
    key(WM_KEYDOWN, VK_RIGHT, kLate);
    key(WM_KEYDOWN, VK_SPACE, kLate);
    out.guardKeepsFocusOnDefault = window->_keysGuarded && window->_overlay.GetFocusedButtonId() == std::optional<uint32_t>{kDefaultId} &&
                                   pressed.buttonCount == 0u && pressed.closes == 0u;
    // An auto-repeated key (a key held for the pane) acts on nothing, guarded or not.
    key(WM_KEYDOWN, VK_ESCAPE, kLate, true);
    key(WM_KEYDOWN, VK_RETURN, kLate, true);
    out.guardIgnoresRepeat = pressed.buttonCount == 0u && window->IsVisible();
    // Typed before the show or within the typeahead window: meant for the pane.
    key(WM_KEYDOWN, VK_RETURN, -40);
    key(WM_KEYDOWN, VK_ESCAPE, 120);
    out.guardIgnoresTypeahead = pressed.buttonCount == 0u && window->IsVisible();
    // A fresh Enter after the window submits only the default, a fresh Escape only the escape binding.
    key(WM_KEYDOWN, VK_RETURN, kLate);
    out.guardEnterSubmitsDefault = pressed.buttonCount == 1u && pressed.buttons[0] == kDefaultId;
    key(WM_KEYDOWN, VK_ESCAPE, kLate);
    out.focusedEscapeRunsBinding = pressed.buttonCount == 2u && pressed.buttons[1] == kEscapeId && pressed.closes == 0u && window->IsVisible();
    window->Hide();
    out.hideRestoresFocus = g_debugFocusOverride == pane.get() && ! window->IsVisible() && ! window->_keysGuarded;

    // 1b. An Escape still held when the overlay took the focus acts only after it was released.
    {
        std::array<BYTE, 256> keyboard{};
        const bool keyboardRead = GetKeyboardState(keyboard.data()) != FALSE;
        const BYTE escapeState  = keyboard[VK_ESCAPE];
        keyboard[VK_ESCAPE]     = static_cast<BYTE>(keyboard[VK_ESCAPE] | 0x80u);
        const bool held         = keyboardRead && SetKeyboardState(keyboard.data()) != FALSE;
        g_debugFocusOverride    = pane.get();
        const bool shown        = show(true, {});
        keyboard[VK_ESCAPE]     = escapeState;
        static_cast<void>(SetKeyboardState(keyboard.data()));
        if (! shown)
            return false;
        key(WM_KEYDOWN, VK_ESCAPE, kLate);
        const bool ignoredWhileHeld = pressed.buttonCount == 0u;
        key(WM_KEYUP, VK_ESCAPE, kLate);
        key(WM_KEYDOWN, VK_ESCAPE, kLate + 50);
        out.guardIgnoresHeldKey = held && ignoredWhileHeld && pressed.buttonCount == 1u && pressed.buttons[0] == kEscapeId;
    }

    // 1c. A click on the overlay lifts the guard: then Tab moves the focus and reaches the close glyph last,
    //     while an auto-repeated key still acts on nothing.
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(2, 2)));
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_LBUTTONUP, 0, MAKELPARAM(2, 2)));
    out.clickLiftsGuard = ! window->_keysGuarded;
    bool reachedClose   = false;
    for (int press = 0; press < 4 && ! reachedClose; ++press)
    {
        key(WM_KEYDOWN, VK_TAB, kLate);
        reachedClose = window->_overlay.GetFocusedControl().part == AlertHitTest::Part::Close;
    }
    out.unguardedTabReachesClose = reachedClose;
    key(WM_KEYDOWN, VK_TAB, kLate, true);
    out.repeatIgnoredWhenUnguarded = window->_overlay.GetFocusedControl().part == AlertHitTest::Part::Close;

    // 1d. Losing the focus to the overlay's own menu keeps the focus-return target; losing it to another
    //     window forgets it, so a later hide does not take the focus back to the pane the user left.
    wil::unique_hwnd ownedMenu(
        CreateWindowExW(WS_EX_TOOLWINDOW, L"STATIC", L"", WS_POPUP, 0, 0, 8, 8, window->_hwnd.get(), nullptr, GetModuleHandleW(nullptr), nullptr));
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_KILLFOCUS, reinterpret_cast<WPARAM>(ownedMenu.get()), 0));
    out.ownedMenuKeepsReturn = ownedMenu && window->_modelessRestoresFocus && window->_restoreFocus == pane.get();
    ownedMenu.reset();
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_KILLFOCUS, reinterpret_cast<WPARAM>(otherPane.get()), 0));
    const bool forgotten = ! window->_modelessRestoresFocus && window->_restoreFocus == nullptr;
    g_debugFocusOverride = window->_hwnd.get(); // the user clicks back on the overlay later
    window->Hide();
    out.focusLossForgetsReturn = forgotten && g_debugFocusOverride != pane.get();

    // 1e. An option change reaches the owner with its new value; the owner's value is shown in place.
    g_debugFocusOverride = otherPane.get();
    if (! show(false, {}, true))
        return false;
    const uint64_t generationBefore = window->_presentationGeneration;
    window->InvokeOption(kOptionId);
    out.optionChangeReported = pressed.optionCalls == 1u && pressed.optionId == kOptionId && pressed.optionValue == 1u;
    window->SetOptionValue(kOptionId, 0u);
    out.optionValueSetInPlace =
        window->GetOptionValue(kOptionId) == std::optional<uint64_t>{0u} && window->_presentationGeneration == generationBefore && pressed.optionCalls == 1u;
    window->Hide();

    // 1f. F6 hands the focus back to the pane without answering (the overlay stays shown), but only a fresh
    //     F6: one held since the show, one typed within the typeahead window and an auto-repeated one act
    //     on nothing.
    {
        const bool held      = setKeyDown(VK_F6, true);
        g_debugFocusOverride = pane.get();
        const bool shown     = show(true, {});
        static_cast<void>(setKeyDown(VK_F6, false));
        if (! shown)
            return false;
        key(WM_KEYDOWN, VK_F6, kLate); // held since the show
        key(WM_KEYUP, VK_F6, kLate);
        key(WM_KEYDOWN, VK_F6, 120);         // within the typeahead window
        key(WM_KEYDOWN, VK_F6, kLate, true); // auto-repeated
        out.f6IgnoresRepeatTypeaheadAndHeld =
            held && window->_keysGuarded && g_debugFocusOverride == window->_hwnd.get() && pressed.focusReturns == 0u && pressed.buttonCount == 0u;
        key(WM_KEYDOWN, VK_F6, kLate + 20);
        out.f6ReturnsFocusToPane = g_debugFocusOverride == pane.get() && pressed.focusReturns == 1u && pressed.buttonCount == 0u && pressed.closes == 0u &&
                                   window->IsVisible() && ! window->_modelessRestoresFocus && window->_restoreFocus == nullptr;
    }

    // 1g. After a click lifted the guard and Tab moved the focus ring, a focus coming back from the overlay's
    //     own menu keeps the guard lifted; a focus handed to it later without a click (an activation that
    //     returns to it) arms the guard again and pins the focus ring to the default. Shift+F6 leaves like F6.
    {
        g_debugFocusOverride = window->_hwnd.get();
        static_cast<void>(SendMessageW(window->_hwnd.get(), WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(2, 2)));
        static_cast<void>(SendMessageW(window->_hwnd.get(), WM_LBUTTONUP, 0, MAKELPARAM(2, 2)));
        key(WM_KEYDOWN, VK_TAB, kLate);
        const bool movedUnguarded = ! window->_keysGuarded && window->_overlay.GetFocusedButtonId() != std::optional<uint32_t>{kDefaultId};
        wil::unique_hwnd menu(
            CreateWindowExW(WS_EX_TOOLWINDOW, L"STATIC", L"", WS_POPUP, 0, 0, 8, 8, window->_hwnd.get(), nullptr, GetModuleHandleW(nullptr), nullptr));
        static_cast<void>(SendMessageW(window->_hwnd.get(), WM_SETFOCUS, reinterpret_cast<WPARAM>(menu.get()), 0));
        out.ownedMenuFocusReturnKeepsGuard =
            menu && movedUnguarded && ! window->_keysGuarded && window->_overlay.GetFocusedButtonId() != std::optional<uint32_t>{kDefaultId};
        menu.reset();
        static_cast<void>(SendMessageW(window->_hwnd.get(), WM_SETFOCUS, reinterpret_cast<WPARAM>(otherPane.get()), 0));
        out.focusGainArmsGuard = movedUnguarded && window->_keysGuarded && window->_overlay.GetFocusedButtonId() == std::optional<uint32_t>{kDefaultId};
        const bool shifted     = setKeyDown(VK_SHIFT, true);
        key(WM_KEYDOWN, VK_F6, kLate);
        static_cast<void>(setKeyDown(VK_SHIFT, false));
        out.shiftF6ReturnsFocusToPane =
            shifted && g_debugFocusOverride == pane.get() && pressed.focusReturns == 2u && pressed.buttonCount == 0u && window->IsVisible();
    }

    // 1h. Ctrl+Tab lifts the guard and puts the focus ring on the default, but only a fresh chord: one typed
    //     within the typeahead window, an auto-repeated one and one whose Ctrl was held since the focus arrived
    //     act on nothing. Unguarded, Tab then moves the focus ring off the default.
    {
        const bool controlHeld = setKeyDown(VK_CONTROL, true);
        g_debugFocusOverride   = window->_hwnd.get();
        static_cast<void>(SendMessageW(window->_hwnd.get(), WM_SETFOCUS, reinterpret_cast<WPARAM>(otherPane.get()), 0));
        const bool armedWithControl = window->_keysGuarded && (window->_guardHeldKeys & GuardKeyBit(VK_CONTROL)) != 0u;
        key(WM_KEYDOWN, VK_TAB, kLate); // its Ctrl held since the focus arrived
        key(WM_KEYUP, VK_CONTROL, kLate);
        key(WM_KEYDOWN, VK_TAB, 120);         // within the typeahead window
        key(WM_KEYDOWN, VK_TAB, kLate, true); // auto-repeated
        out.ctrlTabIgnoresRepeatTypeaheadAndHeld =
            controlHeld && armedWithControl && window->_keysGuarded && g_debugFocusOverride == window->_hwnd.get() && pressed.buttonCount == 0u;
        key(WM_KEYDOWN, VK_TAB, kLate + 20);
        out.ctrlTabLiftsGuardToDefault =
            ! window->_keysGuarded && window->_overlay.GetFocusedButtonId() == std::optional<uint32_t>{kDefaultId} && pressed.buttonCount == 0u;
        static_cast<void>(setKeyDown(VK_CONTROL, false));
        key(WM_KEYDOWN, VK_TAB, kLate + 40);
        out.tabAfterCtrlTabLeavesDefault =
            out.ctrlTabLiftsGuardToDefault && window->_overlay.GetFocusedButtonId() != std::optional<uint32_t>{kDefaultId} && pressed.buttonCount == 0u;
    }
    window->Hide();

    // 1i. An overlay that does not require the focus for its keys (a plugin's non-modal alert) keeps acting on
    //     an auto-repeated key and lets Alt+F4 close its own window.
    {
        auto plain = std::make_unique<AlertOverlayWindow>();
        CallbackState plainPressed{};
        AlertOverlayWindowCallbacks plainCallbacks = callbacks;
        plainCallbacks.context                     = &plainPressed;
        plain->SetCallbacks(plainCallbacks);
        plain->SetKeyBindings(kDefaultId, kEscapeId);
        if (FAILED(plain->ShowForAnchor(otherPane.get(), theme, makeModel(false), false)) || ! plain->_hwnd)
            return false;
        const HWND plainWindow = plain->_hwnd.get();
        static_cast<void>(SendMessageW(plainWindow, WM_KEYDOWN, VK_RETURN, kKeyPreviouslyDownFlag));
        const bool repeatActed = plainPressed.buttonCount == 1u && plainPressed.buttons[0] == kDefaultId;
        static_cast<void>(SendMessageW(plainWindow, WM_CLOSE, 0, 0));
        MSG forwarded{};
        const bool posted                   = PeekMessageW(&forwarded, root.get(), WM_CLOSE, WM_CLOSE, PM_REMOVE) != FALSE;
        out.plainOverlayKeepsRepeatAndClose = repeatActed && ! posted && IsWindow(plainWindow) == FALSE && ! plain->IsVisible();
    }

    // 2. Shown while another window had the focus (the owner does not ask for it): the focus stays
    //    there, and an Escape routed to every visible alert, or Enter, is inert.
    g_debugFocusOverride = otherPane.get();
    if (! show(false, {}))
        return false;
    const bool keptAway = g_debugFocusOverride == otherPane.get() && ! window->_modelessRestoresFocus && window->_restoreFocus == nullptr;
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_KEYDOWN, VK_ESCAPE, 0));
    out.unfocusedEscapeIgnored = keptAway && pressed.buttonCount == 0u && pressed.closes == 0u && window->IsVisible();
    static_cast<void>(SendMessageW(window->_hwnd.get(), WM_KEYDOWN, VK_RETURN, 0));
    out.unfocusedEnterIgnored = pressed.buttonCount == 0u && window->IsVisible();

    // 3. The close glyph runs onClose, not the escape binding (close is never a decision).
    window->InvokeDismiss(AlertOverlayWindow::DismissSource::Close);
    out.closeRunsOnCloseNotBinding = pressed.closes == 1u && pressed.buttonCount == 0u;

    // 3b. A blocking alert whose focus-return target is hidden gives the focus to its own anchor instead: the
    //     helper, then a real blocking alert whose captured focus window was hidden while it was shown (its
    //     owner holds the foreground, so the hide hands the focus back without a foreground request).
    {
        wil::unique_hwnd hiddenTarget(CreateWindowExW(0, L"STATIC", L"", WS_CHILD, 0, 0, 8, 8, host.get(), nullptr, GetModuleHandleW(nullptr), nullptr));
        out.hiddenReturnFallsBackToAnchor = hiddenTarget && window->ResolveFocusReturnTarget(hiddenTarget.get()) == window->_anchor &&
                                            window->ResolveFocusReturnTarget(otherPane.get()) == otherPane.get();
    }
    {
        wil::unique_hwnd target(CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE, 0, 0, 8, 8, host.get(), nullptr, GetModuleHandleW(nullptr), nullptr));
        auto blocking                   = std::make_unique<AlertOverlayWindow>();
        g_debugFocusOverride            = target.get();
        g_debugForegroundOverrideActive = true;
        g_debugForegroundOverride       = root.get();
        const uint64_t requestsBefore   = DebugGetAlertOverlayForegroundRequestCount();
        const bool captured =
            target && SUCCEEDED(blocking->ShowForAnchor(otherPane.get(), theme, makeModel(false), true)) && blocking->_restoreFocus == target.get();
        ShowWindow(target.get(), SW_HIDE);
        blocking->Hide();
        out.blockingHideFallsBackToAnchor =
            captured && g_debugFocusOverride == otherPane.get() && DebugGetAlertOverlayForegroundRequestCount() == requestsBefore;
        g_debugForegroundOverrideActive = false;
        g_debugForegroundOverride       = nullptr;
        g_debugFocusOverride            = otherPane.get();
    }

    // 3c. Alt+F4 (WM_CLOSE) on the modeless overlay goes to its owner's top-level window.
    {
        static_cast<void>(SendMessageW(window->_hwnd.get(), WM_CLOSE, 0, 0));
        MSG forwarded{};
        const bool posted                = PeekMessageW(&forwarded, root.get(), WM_CLOSE, WM_CLOSE, PM_REMOVE) != FALSE;
        out.closeMessageForwardedToOwner = posted && window->_hwnd && IsWindow(window->_hwnd.get()) != FALSE && window->IsVisible();
    }

    // 4. A move of the top-level window moves the overlay with the pane and keeps its place in the z-order
    //    (an unrelated window raised above it stays above), while it raises a blocking alert again; a hidden
    //    pane hides the modeless overlay.
    const auto isAbove = [](HWND upper, HWND lower) noexcept
    {
        for (HWND candidate = lower ? GetWindow(lower, GW_HWNDPREV) : nullptr; candidate != nullptr; candidate = GetWindow(candidate, GW_HWNDPREV))
        {
            if (candidate == upper)
                return true;
        }
        return false;
    };
    wil::unique_hwnd above(CreateWindowExW(
        WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE, L"STATIC", L"", WS_POPUP, left, top, 16, 16, nullptr, nullptr, GetModuleHandleW(nullptr), nullptr));
    const auto raiseAbove = [&]() noexcept
    { return above && SetWindowPos(above.get(), HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW) != FALSE; };
    const bool modelessCovered = raiseAbove() && isAbove(above.get(), window->_hwnd.get());
    RECT before{};
    RECT after{};
    const bool placed = GetWindowRect(window->_hwnd.get(), &before) != FALSE;
    SetWindowPos(root.get(), nullptr, left + 40, top + 24, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    out.followsRootMove = placed && GetWindowRect(window->_hwnd.get(), &after) != FALSE && after.left == before.left + 40 && after.top == before.top + 24;
    out.trackingMoveKeepsModelessZOrder = modelessCovered && isAbove(above.get(), window->_hwnd.get());
    {
        auto blocking        = std::make_unique<AlertOverlayWindow>();
        g_debugFocusOverride = otherPane.get();
        const bool covered   = SUCCEEDED(blocking->ShowForAnchor(otherPane.get(), theme, makeModel(false), true)) && blocking->_hwnd && raiseAbove() &&
                               isAbove(above.get(), blocking->_hwnd.get());
        SetWindowPos(root.get(), nullptr, left + 56, top + 32, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        out.trackingMoveRaisesBlocking = covered && isAbove(blocking->_hwnd.get(), above.get());
        blocking->HideWithoutRestoringFocus();
        g_debugFocusOverride = otherPane.get();
    }
    above.reset();
    ShowWindow(pane.get(), SW_HIDE);
    out.hidesWithHiddenAnchor = window->IsVisible() && IsWindowVisible(window->_hwnd.get()) == FALSE;
    ShowWindow(pane.get(), SW_SHOWNA);
    out.showsWithShownAnchor = window->IsVisible() && IsWindowVisible(window->_hwnd.get()) != FALSE;

    // 5. The pane is destroyed under a visible modeless overlay: it hides and its owner is told.
    struct DestroyedState final
    {
        AlertOverlayWindow* overlay = nullptr;
        bool reported               = false;
    } destroyed{window.get()};
    window->SetTrackedWindowDestroyedCallback(&destroyed,
                                              [](void* context, AlertOverlayWindow& overlay) noexcept
    {
        auto& state    = *static_cast<DestroyedState*>(context);
        state.reported = state.overlay == &overlay;
    });
    pane.reset();
    out.destroyedAnchorHides    = ! window->IsVisible();
    out.destroyedAnchorReported = destroyed.reported;
    window->SetTrackedWindowDestroyedCallback(nullptr, nullptr);
    window.reset();

    out.foregroundRequests = DebugGetAlertOverlayForegroundRequestCount() - foregroundBefore;
    return true;
}
#endif
} // namespace RedSalamander::Ui
