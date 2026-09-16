#include "TerminalAccessibility.h"

#include <DxUi/AccessibilityTextUnits.h>

#include <algorithm>
#include <atomic>
#include <cwctype>
#include <limits>
#include <memory>
#include <mutex>
#include <new>
#include <string_view>

#include <UIAutomation.h>
#include <objbase.h>
#include <oleauto.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"
#include "resource.h"

#pragma comment(lib, "uiautomationcore")
#pragma comment(lib, "oleaut32")

extern HINSTANCE g_hInstance;

namespace
{
std::atomic_uint32_t g_accessibilityObjectCount{0u};

struct Snapshot final
{
    uint64_t generation = 0u;
    std::wstring text;
    bool focused = false;
};

struct State final
{
    State()                        = default;
    State(const State&)            = delete;
    State(State&&)                 = delete;
    State& operator=(const State&) = delete;
    State& operator=(State&&)      = delete;

    std::mutex mutex;
    HWND window             = nullptr;
    bool retired            = false;
    uint64_t nextGeneration = 1u;
    std::shared_ptr<const Snapshot> snapshot;
};

struct SafeArrayDeleter final
{
    void operator()(SAFEARRAY* value) const noexcept
    {
        if (value != nullptr)
        {
            static_cast<void>(SafeArrayDestroy(value));
        }
    }
};
using UniqueSafeArray = std::unique_ptr<SAFEARRAY, SafeArrayDeleter>;

[[nodiscard]] HRESULT SetStringVariant(VARIANT* output, std::wstring_view value) noexcept
{
    if (output == nullptr)
    {
        return E_POINTER;
    }
    VariantInit(output);
    BSTR text = SysAllocStringLen(value.data(), static_cast<UINT>(std::min<size_t>(value.size(), (std::numeric_limits<UINT>::max)())));
    if (text == nullptr && ! value.empty())
    {
        return E_OUTOFMEMORY;
    }
    output->vt      = VT_BSTR;
    output->bstrVal = text;
    return S_OK;
}

[[nodiscard]] HRESULT EmptyUnknownArray(SAFEARRAY** output) noexcept
{
    if (output == nullptr)
    {
        return E_POINTER;
    }
    *output = SafeArrayCreateVector(VT_UNKNOWN, 0, 0u);
    return *output != nullptr ? S_OK : E_OUTOFMEMORY;
}

[[nodiscard]] HRESULT EmptyDoubleArray(SAFEARRAY** output) noexcept
{
    if (output == nullptr)
    {
        return E_POINTER;
    }
    *output = SafeArrayCreateVector(VT_R8, 0, 0u);
    return *output != nullptr ? S_OK : E_OUTOFMEMORY;
}

[[nodiscard]] bool IsRetired(const std::shared_ptr<State>& state) noexcept
{
    std::scoped_lock lock(state->mutex);
    return state->retired;
}

[[nodiscard]] std::shared_ptr<const Snapshot> CaptureSnapshot(const std::shared_ptr<State>& state, HWND* window = nullptr) noexcept
{
    std::scoped_lock lock(state->mutex);
    if (state->retired || ! state->snapshot)
    {
        return {};
    }
    if (window != nullptr)
    {
        *window = state->window;
    }
    return state->snapshot;
}

MIDL_INTERFACE("18A4D4E9-424C-46C4-9BC1-5B7E7FDC7985")
ITerminalTextRangeIdentity : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE GetIdentity(const void** stateKey, uint64_t* generation, size_t* start, size_t* end) noexcept = 0;
};

class TerminalTextRange final : public ITextRangeProvider, public ITerminalTextRangeIdentity
{
public:
    TerminalTextRange(
        std::shared_ptr<State> state, std::shared_ptr<const Snapshot> snapshot, IRawElementProviderSimple* owner, size_t start, size_t end) noexcept
        : _state(std::move(state)),
          _snapshot(std::move(snapshot)),
          _owner(owner),
          _start(start),
          _end(end)
    {
        g_accessibilityObjectCount.fetch_add(1u, std::memory_order_acq_rel);
        clamp();
    }

    TerminalTextRange(const TerminalTextRange&)            = delete;
    TerminalTextRange(TerminalTextRange&&)                 = delete;
    TerminalTextRange& operator=(const TerminalTextRange&) = delete;
    TerminalTextRange& operator=(TerminalTextRange&&)      = delete;

    ~TerminalTextRange()
    {
        g_accessibilityObjectCount.fetch_sub(1u, std::memory_order_acq_rel);
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** object) noexcept override
    {
        if (object == nullptr)
        {
            return E_POINTER;
        }
        *object = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(ITextRangeProvider))
        {
            *object = static_cast<ITextRangeProvider*>(this);
        }
        else if (riid == __uuidof(ITerminalTextRangeIdentity))
        {
            *object = static_cast<ITerminalTextRangeIdentity*>(this);
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
        return _refCount.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG value = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (value == 0u)
        {
            delete this;
        }
        return value;
    }

    HRESULT STDMETHODCALLTYPE GetIdentity(const void** stateKey, uint64_t* generation, size_t* start, size_t* end) noexcept override
    {
        if (stateKey == nullptr || generation == nullptr || start == nullptr || end == nullptr)
        {
            return E_POINTER;
        }
        *stateKey   = _state.get();
        *generation = _snapshot->generation;
        *start      = _start;
        *end        = _end;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Clone(ITextRangeProvider** output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = nullptr;
        if (retired())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        auto* range = new (std::nothrow) TerminalTextRange(_state, _snapshot, _owner.get(), _start, _end);
        if (range == nullptr)
        {
            return E_OUTOFMEMORY;
        }
        *output = range;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Compare(ITextRangeProvider* range, BOOL* output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output              = FALSE;
        const void* stateKey = nullptr;
        uint64_t generation  = 0u;
        size_t start         = 0u;
        size_t end           = 0u;
        const HRESULT hr     = identity(range, stateKey, generation, start, end);
        if (FAILED(hr))
        {
            return hr;
        }
        *output = stateKey == _state.get() && generation == _snapshot->generation && start == _start && end == _end;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CompareEndpoints(TextPatternRangeEndpoint endpoint,
                                               ITextRangeProvider* targetRange,
                                               TextPatternRangeEndpoint targetEndpoint,
                                               int* output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output              = 0;
        const void* stateKey = nullptr;
        uint64_t generation  = 0u;
        size_t start         = 0u;
        size_t end           = 0u;
        const HRESULT hr     = identity(targetRange, stateKey, generation, start, end);
        if (FAILED(hr))
        {
            return hr;
        }
        if (stateKey != _state.get() || generation != _snapshot->generation)
        {
            return E_INVALIDARG;
        }
        const size_t left  = endpoint == TextPatternRangeEndpoint_Start ? _start : _end;
        const size_t right = targetEndpoint == TextPatternRangeEndpoint_Start ? start : end;
        *output            = left < right ? -1 : left > right ? 1 : 0;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ExpandToEnclosingUnit(TextUnit unit) noexcept override
    {
        if (retired())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        const DxUi::AccessibilityTextUnitSpan span = DxUi::GetEnclosingAccessibilityTextUnitSpan(_snapshot->text, _start, unit);
        _start                                     = span.start;
        _end                                       = span.end;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FindAttribute(TEXTATTRIBUTEID /*attributeId*/,
                                            VARIANT /*value*/,
                                            BOOL /*backward*/,
                                            ITextRangeProvider** output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = nullptr;
        return retired() ? UIA_E_ELEMENTNOTAVAILABLE : S_OK;
    }

    HRESULT STDMETHODCALLTYPE FindText(BSTR text, BOOL backward, BOOL ignoreCase, ITextRangeProvider** output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = nullptr;
        if (retired())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        if (text == nullptr)
        {
            return E_INVALIDARG;
        }
        const std::wstring_view needle(text, SysStringLen(text));
        const std::wstring_view haystack(_snapshot->text.data() + _start, _end - _start);
        const auto equal = [ignoreCase](wchar_t left, wchar_t right) noexcept
        { return ignoreCase == FALSE ? left == right : towlower(left) == towlower(right); };
        auto found = backward == FALSE ? std::search(haystack.begin(), haystack.end(), needle.begin(), needle.end(), equal)
                                       : std::find_end(haystack.begin(), haystack.end(), needle.begin(), needle.end(), equal);
        if (found == haystack.end())
        {
            return S_OK;
        }
        const size_t start = _start + static_cast<size_t>(std::distance(haystack.begin(), found));
        auto* range        = new (std::nothrow) TerminalTextRange(_state, _snapshot, _owner.get(), start, start + needle.size());
        if (range == nullptr)
        {
            return E_OUTOFMEMORY;
        }
        *output = range;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetAttributeValue(TEXTATTRIBUTEID /*attributeId*/, VARIANT* output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        VariantInit(output);
        return retired() ? UIA_E_ELEMENTNOTAVAILABLE : S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetBoundingRectangles(SAFEARRAY** output) noexcept override
    {
        if (retired())
        {
            if (output != nullptr)
            {
                *output = nullptr;
            }
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        return EmptyDoubleArray(output);
    }

    HRESULT STDMETHODCALLTYPE GetEnclosingElement(IRawElementProviderSimple** output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = nullptr;
        if (retired())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        if (! _owner)
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        *output = _owner.get();
        (*output)->AddRef();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetText(int maxLength, BSTR* output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = nullptr;
        if (retired())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        if (maxLength < -1)
        {
            return E_INVALIDARG;
        }
        const size_t available = _end - _start;
        const size_t requested = maxLength < 0 ? available : std::min(available, static_cast<size_t>(maxLength));
        *output                = SysAllocStringLen(_snapshot->text.data() + _start, static_cast<UINT>(requested));
        return *output != nullptr || requested == 0u ? S_OK : E_OUTOFMEMORY;
    }

    HRESULT STDMETHODCALLTYPE Move(TextUnit unit, int count, int* output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = 0;
        if (retired())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        const bool collapsed = _start == _end;
        const auto result    = DxUi::MoveAccessibilityTextPositionByUnit(_snapshot->text, _start, unit, count);
        if (result.moved != 0 || collapsed)
        {
            const auto span = DxUi::GetEnclosingAccessibilityTextUnitSpan(_snapshot->text, result.position, unit);
            _start          = collapsed ? result.position : span.start;
            _end            = collapsed ? result.position : span.end;
        }
        *output = result.moved;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE MoveEndpointByUnit(TextPatternRangeEndpoint endpoint, TextUnit unit, int count, int* output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = 0;
        if (retired())
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        size_t& value     = endpoint == TextPatternRangeEndpoint_Start ? _start : _end;
        const auto result = DxUi::MoveAccessibilityTextPositionByUnit(_snapshot->text, value, unit, count);
        value             = result.position;
        if (_start > _end)
        {
            if (endpoint == TextPatternRangeEndpoint_Start)
            {
                _end = _start;
            }
            else
            {
                _start = _end;
            }
        }
        *output = result.moved;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE MoveEndpointByRange(TextPatternRangeEndpoint endpoint,
                                                  ITextRangeProvider* targetRange,
                                                  TextPatternRangeEndpoint targetEndpoint) noexcept override
    {
        const void* stateKey = nullptr;
        uint64_t generation  = 0u;
        size_t start         = 0u;
        size_t end           = 0u;
        const HRESULT hr     = identity(targetRange, stateKey, generation, start, end);
        if (FAILED(hr))
        {
            return hr;
        }
        if (stateKey != _state.get() || generation != _snapshot->generation)
        {
            return E_INVALIDARG;
        }
        size_t& value = endpoint == TextPatternRangeEndpoint_Start ? _start : _end;
        value         = targetEndpoint == TextPatternRangeEndpoint_Start ? start : end;
        if (_start > _end)
        {
            if (endpoint == TextPatternRangeEndpoint_Start)
            {
                _end = _start;
            }
            else
            {
                _start = _end;
            }
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Select() noexcept override
    {
        return retired() ? UIA_E_ELEMENTNOTAVAILABLE : UIA_E_NOTSUPPORTED;
    }
    HRESULT STDMETHODCALLTYPE AddToSelection() noexcept override
    {
        return retired() ? UIA_E_ELEMENTNOTAVAILABLE : UIA_E_NOTSUPPORTED;
    }
    HRESULT STDMETHODCALLTYPE RemoveFromSelection() noexcept override
    {
        return retired() ? UIA_E_ELEMENTNOTAVAILABLE : UIA_E_NOTSUPPORTED;
    }
    HRESULT STDMETHODCALLTYPE ScrollIntoView(BOOL /*alignToTop*/) noexcept override
    {
        return retired() ? UIA_E_ELEMENTNOTAVAILABLE : UIA_E_NOTSUPPORTED;
    }
    HRESULT STDMETHODCALLTYPE GetChildren(SAFEARRAY** output) noexcept override
    {
        if (retired())
        {
            if (output != nullptr)
            {
                *output = nullptr;
            }
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        return EmptyUnknownArray(output);
    }

private:
    [[nodiscard]] bool retired() const noexcept
    {
        return IsRetired(_state);
    }

    [[nodiscard]] HRESULT identity(ITextRangeProvider* range, const void*& stateKey, uint64_t& generation, size_t& start, size_t& end) const noexcept
    {
        if (range == nullptr || retired())
        {
            return range == nullptr ? E_INVALIDARG : UIA_E_ELEMENTNOTAVAILABLE;
        }
        wil::com_ptr_nothrow<ITerminalTextRangeIdentity> identityProvider;
        const HRESULT hr = range->QueryInterface(__uuidof(ITerminalTextRangeIdentity), identityProvider.put_void());
        if (FAILED(hr) || ! identityProvider)
        {
            return E_INVALIDARG;
        }
        return identityProvider->GetIdentity(&stateKey, &generation, &start, &end);
    }

    void clamp() noexcept
    {
        _start = std::min(_start, _snapshot->text.size());
        _end   = std::clamp(_end, _start, _snapshot->text.size());
    }

    std::atomic<ULONG> _refCount{1u};
    std::shared_ptr<State> _state;
    std::shared_ptr<const Snapshot> _snapshot;
    wil::com_ptr_nothrow<IRawElementProviderSimple> _owner;
    size_t _start = 0u;
    size_t _end   = 0u;
};

class TerminalProvider final : public IRawElementProviderSimple, public ITextProvider
{
public:
    explicit TerminalProvider(std::shared_ptr<State> state) noexcept : _state(std::move(state))
    {
        g_accessibilityObjectCount.fetch_add(1u, std::memory_order_acq_rel);
    }
    TerminalProvider(const TerminalProvider&)            = delete;
    TerminalProvider(TerminalProvider&&)                 = delete;
    TerminalProvider& operator=(const TerminalProvider&) = delete;
    TerminalProvider& operator=(TerminalProvider&&)      = delete;
    ~TerminalProvider()
    {
        g_accessibilityObjectCount.fetch_sub(1u, std::memory_order_acq_rel);
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** object) noexcept override
    {
        if (object == nullptr)
        {
            return E_POINTER;
        }
        *object = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IRawElementProviderSimple))
        {
            *object = static_cast<IRawElementProviderSimple*>(this);
        }
        else if (riid == __uuidof(ITextProvider))
        {
            *object = static_cast<ITextProvider*>(this);
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
        return _refCount.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    }
    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG value = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (value == 0u)
        {
            delete this;
        }
        return value;
    }

    HRESULT STDMETHODCALLTYPE get_ProviderOptions(ProviderOptions* output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = ProviderOptions_ServerSideProvider;
        return IsRetired(_state) ? UIA_E_ELEMENTNOTAVAILABLE : S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetPatternProvider(PATTERNID patternId, IUnknown** output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = nullptr;
        if (IsRetired(_state))
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        if (patternId == UIA_TextPatternId)
        {
            *output = static_cast<ITextProvider*>(this);
            (*output)->AddRef();
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetPropertyValue(PROPERTYID propertyId, VARIANT* output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        VariantInit(output);
        HWND window         = nullptr;
        const auto snapshot = CaptureSnapshot(_state, &window);
        if (! snapshot)
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        switch (propertyId)
        {
            case UIA_NamePropertyId: return SetStringVariant(output, LoadStringResource(g_hInstance, IDS_TERMINAL_ACCESSIBILITY_NAME));
            case UIA_AutomationIdPropertyId: return SetStringVariant(output, L"RedSalamander.EmbeddedTerminal");
            case UIA_ClassNamePropertyId: return SetStringVariant(output, L"RedSalamander.Terminal.Plugin.Window");
            case UIA_ControlTypePropertyId:
                output->vt   = VT_I4;
                output->lVal = UIA_DocumentControlTypeId;
                break;
            case UIA_NativeWindowHandlePropertyId:
                output->vt   = VT_I4;
                output->lVal = HandleToLong(window);
                break;
            case UIA_IsControlElementPropertyId:
            case UIA_IsContentElementPropertyId:
            case UIA_IsEnabledPropertyId:
            case UIA_IsKeyboardFocusablePropertyId:
                output->vt      = VT_BOOL;
                output->boolVal = VARIANT_TRUE;
                break;
            case UIA_HasKeyboardFocusPropertyId:
                output->vt      = VT_BOOL;
                output->boolVal = snapshot->focused ? VARIANT_TRUE : VARIANT_FALSE;
                break;
            case UIA_IsPasswordPropertyId:
                output->vt      = VT_BOOL;
                output->boolVal = VARIANT_FALSE;
                break;
            default: break;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE get_HostRawElementProvider(IRawElementProviderSimple** output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output     = nullptr;
        HWND window = nullptr;
        if (! CaptureSnapshot(_state, &window) || window == nullptr)
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        return UiaHostProviderFromHwnd(window, output);
    }

    HRESULT STDMETHODCALLTYPE GetSelection(SAFEARRAY** output) noexcept override
    {
        if (IsRetired(_state))
        {
            if (output != nullptr)
            {
                *output = nullptr;
            }
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        return EmptyUnknownArray(output);
    }

    HRESULT STDMETHODCALLTYPE GetVisibleRanges(SAFEARRAY** output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output                   = nullptr;
        ITextRangeProvider* range = nullptr;
        HRESULT hr                = get_DocumentRange(&range);
        if (FAILED(hr))
        {
            return hr;
        }
        UniqueSafeArray array(SafeArrayCreateVector(VT_UNKNOWN, 0, 1u));
        if (! array)
        {
            range->Release();
            return E_OUTOFMEMORY;
        }
        LONG index = 0;
        hr         = SafeArrayPutElement(array.get(), &index, range);
        range->Release();
        if (FAILED(hr))
        {
            return hr;
        }
        *output = array.release();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE RangeFromChild(IRawElementProviderSimple* /*childElement*/, ITextRangeProvider** output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = nullptr;
        return IsRetired(_state) ? UIA_E_ELEMENTNOTAVAILABLE : E_INVALIDARG;
    }

    HRESULT STDMETHODCALLTYPE RangeFromPoint(UiaPoint /*point*/, ITextRangeProvider** output) noexcept override
    {
        const auto snapshot = CaptureSnapshot(_state);
        return createRange(snapshot, snapshot ? snapshot->text.size() : 0u, snapshot ? snapshot->text.size() : 0u, output);
    }

    HRESULT STDMETHODCALLTYPE get_DocumentRange(ITextRangeProvider** output) noexcept override
    {
        const auto snapshot = CaptureSnapshot(_state);
        return createRange(snapshot, 0u, snapshot ? snapshot->text.size() : 0u, output);
    }

    HRESULT STDMETHODCALLTYPE get_SupportedTextSelection(SupportedTextSelection* output) noexcept override
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = SupportedTextSelection_None;
        return IsRetired(_state) ? UIA_E_ELEMENTNOTAVAILABLE : S_OK;
    }

private:
    [[nodiscard]] HRESULT createRange(const std::shared_ptr<const Snapshot>& snapshot, size_t start, size_t end, ITextRangeProvider** output) noexcept
    {
        if (output == nullptr)
        {
            return E_POINTER;
        }
        *output = nullptr;
        if (! snapshot)
        {
            return UIA_E_ELEMENTNOTAVAILABLE;
        }
        auto* range = new (std::nothrow) TerminalTextRange(_state, snapshot, static_cast<IRawElementProviderSimple*>(this), start, end);
        if (range == nullptr)
        {
            return E_OUTOFMEMORY;
        }
        *output = range;
        return S_OK;
    }

    std::atomic<ULONG> _refCount{1u};
    std::shared_ptr<State> _state;
};
} // namespace

struct TerminalAccessibility::Impl final
{
    std::shared_ptr<State> state;
    wil::com_ptr_nothrow<IRawElementProviderSimple> provider;
};

TerminalAccessibility::TerminalAccessibility() : _impl(std::make_unique<Impl>())
{
}
TerminalAccessibility::~TerminalAccessibility()
{
    Retire();
}

HRESULT TerminalAccessibility::Initialize(HWND window) noexcept
{
    if (window == nullptr || _impl->state || _impl->provider)
    {
        return E_INVALIDARG;
    }
    _impl->state        = std::make_shared<State>();
    auto initial        = std::make_shared<Snapshot>();
    initial->generation = 1u;
    {
        std::scoped_lock lock(_impl->state->mutex);
        _impl->state->window         = window;
        _impl->state->nextGeneration = 2u;
        _impl->state->snapshot       = std::move(initial);
    }
    auto* provider = new (std::nothrow) TerminalProvider(_impl->state);
    if (provider == nullptr)
    {
        _impl->state.reset();
        return E_OUTOFMEMORY;
    }
    _impl->provider.attach(static_cast<IRawElementProviderSimple*>(provider));
    return S_OK;
}

void TerminalAccessibility::Publish(std::wstring text, bool focused) noexcept
{
    if (! _impl->state)
    {
        return;
    }
    auto next        = std::make_shared<Snapshot>();
    next->text       = std::move(text);
    next->focused    = focused;
    bool textChanged = false;
    {
        std::scoped_lock lock(_impl->state->mutex);
        if (_impl->state->retired)
        {
            return;
        }
        next->generation       = _impl->state->nextGeneration++;
        textChanged            = ! _impl->state->snapshot || _impl->state->snapshot->text != next->text;
        _impl->state->snapshot = std::move(next);
    }
    if (textChanged && _impl->provider)
    {
        static_cast<void>(UiaRaiseAutomationEvent(_impl->provider.get(), UIA_Text_TextChangedEventId));
    }
}

LRESULT TerminalAccessibility::HandleGetObject(WPARAM wParam, LPARAM lParam) noexcept
{
    if (! _impl->state || ! _impl->provider)
    {
        return 0;
    }
    HWND window = nullptr;
    if (! CaptureSnapshot(_impl->state, &window) || window == nullptr)
    {
        return 0;
    }
    return UiaReturnRawElementProvider(window, wParam, lParam, _impl->provider.get());
}

void TerminalAccessibility::Retire() noexcept
{
    if (! _impl || ! _impl->state)
    {
        return;
    }
    HWND window = nullptr;
    {
        std::scoped_lock lock(_impl->state->mutex);
        if (_impl->state->retired)
        {
            return;
        }
        _impl->state->retired = true;
        window                = _impl->state->window;
        _impl->state->window  = nullptr;
        _impl->state->snapshot.reset();
    }
    if (_impl->provider)
    {
        static_cast<void>(UiaDisconnectProvider(_impl->provider.get()));
    }
    if (window != nullptr)
    {
        static_cast<void>(UiaReturnRawElementProvider(window, 0u, 0, nullptr));
    }
    _impl->provider.reset();
    _impl->state.reset();
}

#if defined(ENABLE_TESTS)
HRESULT TerminalAccessibility::DebugGetName(std::wstring& name) noexcept
{
    name.clear();
    if (! _impl || ! _impl->provider)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
    }

    VARIANT value{};
    VariantInit(&value);
    const auto clearValue = wil::scope_exit([&value]() noexcept { static_cast<void>(VariantClear(&value)); });
    const HRESULT hr      = _impl->provider->GetPropertyValue(UIA_NamePropertyId, &value);
    if (FAILED(hr))
    {
        return hr;
    }
    if (value.vt != VT_BSTR || value.bstrVal == nullptr)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    name.assign(value.bstrVal, SysStringLen(value.bstrVal));
    return S_OK;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderTerminalAccessibilityLocalizationSelfTests(unsigned int* passedTests,
                                                                                                           unsigned int* failedTests) noexcept
{
    if (passedTests == nullptr || failedTests == nullptr)
    {
        return E_POINTER;
    }
    *passedTests     = 0u;
    *failedTests     = 0u;
    const auto check = [passedTests, failedTests](bool condition) noexcept
    {
        if (condition)
        {
            ++*passedTests;
        }
        else
        {
            ++*failedTests;
        }
    };

    TerminalAccessibility accessibility;
    const HRESULT initializeHr = accessibility.Initialize(GetDesktopWindow());
    check(initializeHr == S_OK);

    std::wstring providerName;
    const HRESULT propertyHr = SUCCEEDED(initializeHr) ? accessibility.DebugGetName(providerName) : initializeHr;
    check(propertyHr == S_OK && ! providerName.empty());
    check(providerName == LoadStringResource(g_hInstance, IDS_TERMINAL_ACCESSIBILITY_NAME));
    accessibility.Retire();
    check(CanUnloadTerminalAccessibilityProviders());
    return *failedTests == 0u ? S_OK : E_FAIL;
}
#endif

bool CanUnloadTerminalAccessibilityProviders() noexcept
{
    return g_accessibilityObjectCount.load(std::memory_order_acquire) == 0u;
}
