#include "DxUi.Internal.h"

#include <algorithm>
#include <atomic>
#include <limits>

#include <msctf.h>
#include <textstor.h>

namespace RedSalamander::DxUi
{
namespace
{
constexpr TsViewCookie kDxUiTextStoreView = 1u;

[[nodiscard]] LONG ToAcp(size_t value) noexcept
{
    return static_cast<LONG>((std::min)(value, static_cast<size_t>((std::numeric_limits<LONG>::max)())));
}

[[nodiscard]] bool IsTextStoreLayoutAvailable(const D2D1_RECT_F& bounds) noexcept
{
    return bounds.right > bounds.left && bounds.bottom > bounds.top;
}

[[nodiscard]] size_t ClampAcp(LONG value, size_t textLength, bool negativeMeansEnd = false) noexcept
{
    if (value < 0)
    {
        return negativeMeansEnd ? textLength : 0u;
    }

    return (std::min)(static_cast<size_t>(value), textLength);
}

struct AcpRange
{
    size_t start = 0u;
    size_t end   = 0u;
};

[[nodiscard]] AcpRange ClampAcpRange(LONG start, LONG end, size_t textLength, bool endMinusOneMeansEnd = true) noexcept
{
    AcpRange range{ClampAcp(start, textLength), ClampAcp(end, textLength, endMinusOneMeansEnd)};
    if (range.end < range.start)
    {
        std::swap(range.start, range.end);
    }
    return range;
}

[[nodiscard]] bool HasSelection(const TextInputState& state) noexcept
{
    return state.selectionAnchorIndex.has_value() && state.selectionAnchorIndex.value() != state.caretIndex;
}

[[nodiscard]] AcpRange GetSelectionRange(const TextInputState& state) noexcept
{
    if (! HasSelection(state))
    {
        const size_t caretIndex = (std::min)(state.caretIndex, state.text.size());
        return AcpRange{caretIndex, caretIndex};
    }

    return ClampAcpRange(ToAcp(state.selectionAnchorIndex.value()), ToAcp(state.caretIndex), state.text.size(), false);
}

[[nodiscard]] bool IsSameSelection(const TextInputState& left, const TextInputState& right) noexcept
{
    return left.caretIndex == right.caretIndex && left.selectionAnchorIndex == right.selectionAnchorIndex;
}

[[nodiscard]] bool IsSameRect(const D2D1_RECT_F& left, const D2D1_RECT_F& right) noexcept
{
    return left.left == right.left && left.top == right.top && left.right == right.right && left.bottom == right.bottom;
}

[[nodiscard]] RECT DipRectToScreenRect(WindowHost& host, const D2D1_RECT_F& rectDip) noexcept
{
    POINT topLeft{static_cast<LONG>(std::lround(host.DipsToPixels(rectDip.left))), static_cast<LONG>(std::lround(host.DipsToPixels(rectDip.top)))};
    POINT bottomRight{static_cast<LONG>(std::lround(host.DipsToPixels(rectDip.right))), static_cast<LONG>(std::lround(host.DipsToPixels(rectDip.bottom)))};
    ClientToScreen(host.GetHwnd(), &topLeft);
    ClientToScreen(host.GetHwnd(), &bottomRight);
    return RECT{topLeft.x, topLeft.y, bottomRight.x, bottomRight.y};
}

[[nodiscard]] D2D1_RECT_F ClipTextStoreRectToBounds(D2D1_RECT_F rect, const D2D1_RECT_F& bounds) noexcept
{
    rect.left   = std::clamp(rect.left, bounds.left, bounds.right);
    rect.right  = std::clamp(rect.right, rect.left, bounds.right);
    rect.top    = std::clamp(rect.top, bounds.top, bounds.bottom);
    rect.bottom = std::clamp(rect.bottom, rect.top, bounds.bottom);
    if (rect.right <= rect.left)
    {
        rect.right = (std::min)(bounds.right, rect.left + 1.0f);
    }
    if (rect.bottom <= rect.top)
    {
        rect.bottom = (std::min)(bounds.bottom, rect.top + 1.0f);
    }
    return rect;
}

[[nodiscard]] std::optional<D2D1_RECT_F> TryResolveMultilineTextStoreRangeRect(const WindowHost& host,
                                                                               const Control& control,
                                                                               const AcpRange& range,
                                                                               const D2D1_RECT_F& bounds) noexcept
{
    std::optional<D2D1_RECT_F> result;
    for (size_t index = range.start;; ++index)
    {
        const std::optional<D2D1_RECT_F> caretRect = control.TryGetTextInputCaretRect(host, index);
        if (! caretRect.has_value())
        {
            return std::nullopt;
        }

        if (! result.has_value())
        {
            result = caretRect.value();
        }
        else
        {
            D2D1_RECT_F rect = result.value();
            rect.left        = (std::min)(rect.left, caretRect->left);
            rect.top         = (std::min)(rect.top, caretRect->top);
            rect.right       = (std::max)(rect.right, caretRect->right);
            rect.bottom      = (std::max)(rect.bottom, caretRect->bottom);
            result           = rect;
        }

        if (index == range.end)
        {
            break;
        }
    }

    return ClipTextStoreRectToBounds(result.value(), bounds);
}

class DxUiTextStoreACP final : public ITextStoreACP, public ITextStoreACP2, public ITfContextOwnerCompositionSink
{
public:
    DxUiTextStoreACP(WindowHost& host, Control& control) noexcept : _host(&host), _control(&control)
    {
    }

    DxUiTextStoreACP(const DxUiTextStoreACP&)            = delete;
    DxUiTextStoreACP& operator=(const DxUiTextStoreACP&) = delete;
    DxUiTextStoreACP(DxUiTextStoreACP&&)                 = delete;
    DxUiTextStoreACP& operator=(DxUiTextStoreACP&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(ITextStoreACP))
        {
            *ppvObject = static_cast<ITextStoreACP*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(ITextStoreACP2))
        {
            *ppvObject = static_cast<ITextStoreACP2*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == __uuidof(ITfContextOwnerCompositionSink))
        {
            *ppvObject = static_cast<ITfContextOwnerCompositionSink*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _referenceCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG remaining = _referenceCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (remaining == 0u)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE OnStartComposition(ITfCompositionView* /*composition*/, BOOL* outAccepted) noexcept override
    {
        if (! outAccepted)
        {
            return E_POINTER;
        }

        *outAccepted = TRUE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnUpdateComposition(ITfCompositionView* /*composition*/, ITfRange* /*rangeNew*/) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnEndComposition(ITfCompositionView* /*composition*/) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE AdviseSink(REFIID riid, IUnknown* punk, DWORD dwMask) noexcept override
    {
        if (! punk)
        {
            return E_INVALIDARG;
        }
        if (riid != __uuidof(ITextStoreACPSink))
        {
            return E_NOINTERFACE;
        }
        if (_sink)
        {
            return E_FAIL;
        }

        ITextStoreACPSink* rawSink = nullptr;
        const HRESULT hr           = punk->QueryInterface(__uuidof(ITextStoreACPSink), reinterpret_cast<void**>(&rawSink));
        if (FAILED(hr))
        {
            return hr;
        }

        _sink.attach(rawSink);
        _sinkMask = dwMask;
        CaptureObservedState();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE UnadviseSink(IUnknown* punk) noexcept override
    {
        if (! punk)
        {
            return E_INVALIDARG;
        }
        if (! _sink)
        {
            return E_FAIL;
        }

        wil::com_ptr_nothrow<IUnknown> requestedIdentity;
        HRESULT hr = punk->QueryInterface(__uuidof(IUnknown), requestedIdentity.put_void());
        if (FAILED(hr) || ! requestedIdentity)
        {
            return E_FAIL;
        }

        wil::com_ptr_nothrow<IUnknown> advisedIdentity;
        hr = _sink->QueryInterface(__uuidof(IUnknown), advisedIdentity.put_void());
        if (FAILED(hr) || ! advisedIdentity || requestedIdentity.get() != advisedIdentity.get())
        {
            return E_FAIL;
        }

        _sink.reset();
        _sinkMask = 0u;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE RequestLock(DWORD dwLockFlags, HRESULT* phrSession) noexcept override
    {
        if (! phrSession)
        {
            return E_POINTER;
        }

        *phrSession = S_OK;
        if (_lockFlags != 0u)
        {
            *phrSession = TS_E_SYNCHRONOUS;
            return S_OK;
        }

        NotifyExternalChangesIfNeeded();
        _lockFlags                                   = dwLockFlags;
        wil::com_ptr_nothrow<ITextStoreACPSink> sink = _sink;
        const bool isReadWriteLock                   = (dwLockFlags & TS_LF_READWRITE) == TS_LF_READWRITE;
        bool editTransactionStarted                  = false;
        if (sink && isReadWriteLock)
        {
            const HRESULT startHr = sink->OnStartEditTransaction();
            if (FAILED(startHr))
            {
                _lockFlags  = 0u;
                *phrSession = startHr;
                return S_OK;
            }
            editTransactionStarted = true;
        }

        if (sink)
        {
            *phrSession = sink->OnLockGranted(dwLockFlags);
        }

        if (sink && editTransactionStarted)
        {
            const HRESULT endHr = sink->OnEndEditTransaction();
            if (SUCCEEDED(*phrSession) && FAILED(endHr))
            {
                *phrSession = endHr;
            }
        }
        _lockFlags = 0u;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetStatus(TS_STATUS* pdcs) noexcept override
    {
        if (! pdcs)
        {
            return E_POINTER;
        }

        pdcs->dwDynamicFlags = 0u;
        pdcs->dwStaticFlags  = TS_SS_NOHIDDENTEXT;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE QueryInsert(LONG acpTestStart, LONG acpTestEnd, ULONG cch, LONG* pacpResultStart, LONG* pacpResultEnd) noexcept override
    {
        if (! pacpResultStart || ! pacpResultEnd)
        {
            return E_POINTER;
        }

        TextInputState state{};
        if (! ReadState(state))
        {
            return TS_E_INVALIDPOS;
        }

        const AcpRange range = ClampAcpRange(acpTestStart, acpTestEnd, state.text.size());
        *pacpResultStart     = ToAcp(range.start);
        *pacpResultEnd       = ToAcp(range.start + cch);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetSelection(ULONG ulIndex, ULONG ulCount, TS_SELECTION_ACP* pSelection, ULONG* pcFetched) noexcept override
    {
        if (! HasReadLock())
        {
            return TS_E_NOLOCK;
        }
        if (! pSelection || ! pcFetched)
        {
            return E_POINTER;
        }

        *pcFetched = 0u;
        if (ulCount == 0u || (ulIndex != TS_DEFAULT_SELECTION && ulIndex != 0u))
        {
            return S_OK;
        }

        TextInputState state{};
        if (! ReadState(state))
        {
            return TS_E_INVALIDPOS;
        }

        const AcpRange range             = GetSelectionRange(state);
        pSelection[0].acpStart           = ToAcp(range.start);
        pSelection[0].acpEnd             = ToAcp(range.end);
        pSelection[0].style.ase          = TS_AE_END;
        pSelection[0].style.fInterimChar = FALSE;
        *pcFetched                       = 1u;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetSelection(ULONG ulCount, const TS_SELECTION_ACP* pSelection) noexcept override
    {
        if (! HasReadWriteLock())
        {
            return TS_E_NOLOCK;
        }
        if (ulCount == 0u)
        {
            return S_OK;
        }
        if (! pSelection)
        {
            return E_POINTER;
        }

        TextInputState state{};
        if (! ReadState(state))
        {
            return TS_E_INVALIDPOS;
        }

        const AcpRange range       = ClampAcpRange(pSelection[0].acpStart, pSelection[0].acpEnd, state.text.size(), false);
        state.caretIndex           = range.end;
        state.selectionAnchorIndex = range.start == range.end ? std::nullopt : std::optional<size_t>(range.start);
        if (! ApplyState(state, false))
        {
            return E_FAIL;
        }

        NotifySelectionChanged();
        NotifyLayoutChanged();
        CaptureObservedState();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetText(LONG acpStart,
                                      LONG acpEnd,
                                      WCHAR* pchPlain,
                                      ULONG cchPlainReq,
                                      ULONG* pcchPlainRet,
                                      TS_RUNINFO* prgRunInfo,
                                      ULONG cRunInfoReq,
                                      ULONG* pcRunInfoRet,
                                      LONG* pacpNext) noexcept override
    {
        if (! HasReadLock())
        {
            return TS_E_NOLOCK;
        }
        if (! pcchPlainRet || ! pcRunInfoRet || ! pacpNext)
        {
            return E_POINTER;
        }

        TextInputState state{};
        if (! ReadState(state))
        {
            return TS_E_INVALIDPOS;
        }

        const AcpRange range = ClampAcpRange(acpStart, acpEnd, state.text.size());
        const std::wstring_view value(state.text.data() + range.start, range.end - range.start);
        const size_t copied = (std::min)(value.size(), static_cast<size_t>(cchPlainReq));
        if (pchPlain && copied > 0u)
        {
            std::copy_n(value.data(), copied, pchPlain);
        }

        *pcchPlainRet = static_cast<ULONG>(copied);
        *pcRunInfoRet = 0u;
        if (prgRunInfo && cRunInfoReq > 0u && copied > 0u)
        {
            prgRunInfo[0].type   = TS_RT_PLAIN;
            prgRunInfo[0].uCount = static_cast<ULONG>(copied);
            *pcRunInfoRet        = 1u;
        }
        *pacpNext = ToAcp(range.start + copied);
        return copied == value.size() ? S_OK : S_FALSE;
    }

    HRESULT STDMETHODCALLTYPE SetText(DWORD dwFlags, LONG acpStart, LONG acpEnd, const WCHAR* pchText, ULONG cch, TS_TEXTCHANGE* pChange) noexcept override
    {
        if (! HasReadWriteLock())
        {
            return TS_E_NOLOCK;
        }

        return ReplaceTextRange(dwFlags, acpStart, acpEnd, pchText, cch, pChange);
    }

    HRESULT STDMETHODCALLTYPE GetFormattedText(LONG /*acpStart*/, LONG /*acpEnd*/, IDataObject** ppDataObject) noexcept override
    {
        if (! ppDataObject)
        {
            return E_POINTER;
        }

        *ppDataObject = nullptr;
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE GetEmbedded(LONG /*acpPos*/, REFGUID /*rguidService*/, REFIID /*riid*/, IUnknown** ppunk) noexcept override
    {
        if (! ppunk)
        {
            return E_POINTER;
        }

        *ppunk = nullptr;
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE QueryInsertEmbedded(const GUID* /*pguidService*/, const FORMATETC* /*pFormatEtc*/, BOOL* pfInsertable) noexcept override
    {
        if (! pfInsertable)
        {
            return E_POINTER;
        }

        *pfInsertable = FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE
    InsertEmbedded(DWORD /*dwFlags*/, LONG /*acpStart*/, LONG /*acpEnd*/, IDataObject* /*pDataObject*/, TS_TEXTCHANGE* /*pChange*/) noexcept override
    {
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE RequestSupportedAttrs(DWORD /*dwFlags*/, ULONG /*cFilterAttrs*/, const TS_ATTRID* /*paFilterAttrs*/) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE RequestAttrsAtPosition(LONG /*acpPos*/,
                                                     ULONG /*cFilterAttrs*/,
                                                     const TS_ATTRID* /*paFilterAttrs*/,
                                                     DWORD /*dwFlags*/) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE RequestAttrsTransitioningAtPosition(LONG /*acpPos*/,
                                                                  ULONG /*cFilterAttrs*/,
                                                                  const TS_ATTRID* /*paFilterAttrs*/,
                                                                  DWORD /*dwFlags*/) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FindNextAttrTransition(LONG acpStart,
                                                     LONG /*acpHalt*/,
                                                     ULONG /*cFilterAttrs*/,
                                                     const TS_ATTRID* /*paFilterAttrs*/,
                                                     DWORD /*dwFlags*/,
                                                     LONG* pacpNext,
                                                     BOOL* pfFound,
                                                     LONG* plFoundOffset) noexcept override
    {
        if (! pacpNext || ! pfFound || ! plFoundOffset)
        {
            return E_POINTER;
        }

        *pacpNext      = acpStart;
        *pfFound       = FALSE;
        *plFoundOffset = 0;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE RetrieveRequestedAttrs(ULONG /*ulCount*/, TS_ATTRVAL* /*paAttrVals*/, ULONG* pcFetched) noexcept override
    {
        if (! pcFetched)
        {
            return E_POINTER;
        }

        *pcFetched = 0u;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetEndACP(LONG* pacp) noexcept override
    {
        if (! HasReadLock())
        {
            return TS_E_NOLOCK;
        }
        if (! pacp)
        {
            return E_POINTER;
        }

        TextInputState state{};
        if (! ReadState(state))
        {
            return TS_E_INVALIDPOS;
        }

        *pacp = ToAcp(state.text.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetActiveView(TsViewCookie* pvcView) noexcept override
    {
        if (! pvcView)
        {
            return E_POINTER;
        }

        *pvcView = kDxUiTextStoreView;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetACPFromPoint(TsViewCookie vcView, const POINT* ptScreen, DWORD /*dwFlags*/, LONG* pacp) noexcept override
    {
        if (! HasReadLock())
        {
            return TS_E_NOLOCK;
        }
        if (! ptScreen || ! pacp)
        {
            return E_POINTER;
        }
        if (vcView != kDxUiTextStoreView)
        {
            return E_INVALIDARG;
        }

        TextInputState state{};
        if (! ReadState(state) || ! _host || ! _control)
        {
            return TS_E_INVALIDPOS;
        }

        const std::optional<PointDip> pointDip = _host->ScreenPointToDipPoint(*ptScreen);
        if (! pointDip)
        {
            return TS_E_INVALIDPOS;
        }

        const D2D1_RECT_F bounds = ResolveTextViewportBounds();
        if (! IsTextStoreLayoutAvailable(bounds))
        {
            *pacp = 0;
            return TS_E_NOLAYOUT;
        }
        if (! PointInRect(bounds, D2D1::Point2F(pointDip->x, pointDip->y)))
        {
            return TS_E_INVALIDPOS;
        }

        const D2D1_POINT_2F queryPoint = D2D1::Point2F(pointDip->x, pointDip->y);
        if (const std::optional<size_t> hitIndex = _control->TryHitTestTextInputPoint(*_host, queryPoint); hitIndex.has_value())
        {
            *pacp = ToAcp((std::min)(hitIndex.value(), state.text.size()));
            return S_OK;
        }

        const DWRITE_READING_DIRECTION readingDirection = ResolveReadingDirection(_control->GetFlowDirection());
        *pacp = ToAcp(HitTestCaretIndexDip(_host, state.text, FontRole::Body, bounds, 0.0f, queryPoint, readingDirection));
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetTextExt(TsViewCookie vcView, LONG acpStart, LONG acpEnd, RECT* prc, BOOL* pfClipped) noexcept override
    {
        if (! HasReadLock())
        {
            return TS_E_NOLOCK;
        }
        if (! prc || ! pfClipped)
        {
            return E_POINTER;
        }
        if (vcView != kDxUiTextStoreView)
        {
            return E_INVALIDARG;
        }

        TextInputState state{};
        if (! ReadState(state) || ! _host || ! _control)
        {
            return TS_E_INVALIDPOS;
        }

        const D2D1_RECT_F bounds = ResolveTextViewportBounds();
        if (! IsTextStoreLayoutAvailable(bounds))
        {
            *prc       = RECT{};
            *pfClipped = TRUE;
            return TS_E_NOLAYOUT;
        }
        const AcpRange range = ClampAcpRange(acpStart, acpEnd, state.text.size(), false);
        if (state.multiline)
        {
            const std::optional<D2D1_RECT_F> rectDip = TryResolveMultilineTextStoreRangeRect(*_host, *_control, range, bounds);
            if (! rectDip.has_value())
            {
                *prc       = RECT{};
                *pfClipped = TRUE;
                return TS_E_NOLAYOUT;
            }

            *prc       = DipRectToScreenRect(*_host, rectDip.value());
            *pfClipped = FALSE;
            return S_OK;
        }

        const float heightDip                           = (std::max)(1.0f, bounds.bottom - bounds.top);
        const float layoutWidth                         = (std::max)(1.0f, bounds.right - bounds.left);
        const DWRITE_READING_DIRECTION readingDirection = ResolveReadingDirection(_control->GetFlowDirection());
        const float startOffset = MeasureCaretOffsetDip(_host, state.text, FontRole::Body, range.start, heightDip, readingDirection, layoutWidth);
        const float endOffset   = MeasureCaretOffsetDip(_host, state.text, FontRole::Body, range.end, heightDip, readingDirection, layoutWidth);
        D2D1_RECT_F rectDip     = bounds;
        rectDip.left            = std::clamp(bounds.left + (std::min)(startOffset, endOffset), bounds.left, bounds.right);
        const float minRight    = (std::min)(rectDip.left + 1.0f, bounds.right);
        rectDip.right           = std::clamp(bounds.left + (std::max)(startOffset, endOffset), minRight, bounds.right);

        *prc       = DipRectToScreenRect(*_host, rectDip);
        *pfClipped = FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetScreenExt(TsViewCookie vcView, RECT* prc) noexcept override
    {
        if (! prc)
        {
            return E_POINTER;
        }
        if (vcView != kDxUiTextStoreView)
        {
            return E_INVALIDARG;
        }
        if (! _host || ! _control)
        {
            return TS_E_INVALIDPOS;
        }

        const D2D1_RECT_F bounds = ResolveTextViewportBounds();
        if (! IsTextStoreLayoutAvailable(bounds))
        {
            *prc = RECT{};
            return TS_E_NOLAYOUT;
        }
        POINT topLeft{static_cast<LONG>(std::lround(_host->DipsToPixels(bounds.left))), static_cast<LONG>(std::lround(_host->DipsToPixels(bounds.top)))};
        POINT bottomRight{static_cast<LONG>(std::lround(_host->DipsToPixels(bounds.right))),
                          static_cast<LONG>(std::lround(_host->DipsToPixels(bounds.bottom)))};
        ClientToScreen(_host->GetHwnd(), &topLeft);
        ClientToScreen(_host->GetHwnd(), &bottomRight);
        *prc = RECT{topLeft.x, topLeft.y, bottomRight.x, bottomRight.y};
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetWnd(TsViewCookie vcView, HWND* phwnd) noexcept override
    {
        if (! phwnd)
        {
            return E_POINTER;
        }
        if (vcView != kDxUiTextStoreView)
        {
            return E_INVALIDARG;
        }

        *phwnd = _host ? _host->GetHwnd() : nullptr;
        return *phwnd ? S_OK : TS_E_INVALIDPOS;
    }

    HRESULT STDMETHODCALLTYPE
    InsertTextAtSelection(DWORD dwFlags, const WCHAR* pchText, ULONG cch, LONG* pacpStart, LONG* pacpEnd, TS_TEXTCHANGE* pChange) noexcept override
    {
        if (! HasReadWriteLock())
        {
            return TS_E_NOLOCK;
        }
        if (! pacpStart || ! pacpEnd)
        {
            return E_POINTER;
        }

        TextInputState state{};
        if (! ReadState(state))
        {
            return TS_E_INVALIDPOS;
        }

        const AcpRange range = GetSelectionRange(state);
        *pacpStart           = ToAcp(range.start);
        *pacpEnd             = ToAcp(range.start + cch);
        if ((dwFlags & TS_IAS_QUERYONLY) != 0u)
        {
            if (pChange)
            {
                pChange->acpStart  = ToAcp(range.start);
                pChange->acpOldEnd = ToAcp(range.end);
                pChange->acpNewEnd = ToAcp(range.start + cch);
            }
            return S_OK;
        }

        return ReplaceTextRange(dwFlags, ToAcp(range.start), ToAcp(range.end), pchText, cch, pChange);
    }

    HRESULT STDMETHODCALLTYPE
    InsertEmbeddedAtSelection(DWORD /*dwFlags*/, IDataObject* /*pDataObject*/, LONG* pacpStart, LONG* pacpEnd, TS_TEXTCHANGE* /*pChange*/) noexcept override
    {
        if (! pacpStart || ! pacpEnd)
        {
            return E_POINTER;
        }

        *pacpStart = 0;
        *pacpEnd   = 0;
        return E_NOTIMPL;
    }

private:
    [[nodiscard]] bool HasReadLock() const noexcept
    {
        return (_lockFlags & TS_LF_READ) != 0u;
    }

    [[nodiscard]] bool HasReadWriteLock() const noexcept
    {
        return (_lockFlags & TS_LF_READWRITE) == TS_LF_READWRITE;
    }

    [[nodiscard]] bool ReadState(TextInputState& outState) const noexcept
    {
        if (! _host || ! _control)
        {
            return false;
        }

        NativeTextInputState nativeState{};
        if (_host->TryReadNativeTextInputState(_control, nativeState))
        {
            outState.text                 = nativeState.text;
            outState.selectionAnchorIndex = nativeState.selectionAnchorIndex;
            outState.caretIndex           = nativeState.caretIndex;
            outState.firstVisibleLine     = nativeState.firstVisibleLine;
            outState.readOnly             = nativeState.readOnly;
            outState.masked               = nativeState.masked;
            outState.multiline            = nativeState.multiline;
            return true;
        }

        if (const auto* textField = dynamic_cast<const TextField*>(_control))
        {
            outState.text       = textField->GetText();
            outState.readOnly   = textField->IsReadOnly();
            outState.masked     = textField->IsMasked();
            outState.caretIndex = outState.text.size();
            if (const std::optional<std::pair<size_t, size_t>> selection = textField->GetSelectionRange())
            {
                outState.selectionAnchorIndex = selection->first;
                outState.caretIndex           = selection->second;
            }
            return true;
        }

        if (const auto* comboBox = dynamic_cast<const ComboBox*>(_control); comboBox && comboBox->IsEditable())
        {
            outState.text       = comboBox->GetText();
            outState.caretIndex = outState.text.size();
            if (const std::optional<std::pair<size_t, size_t>> selection = comboBox->GetEditableSelectionRange())
            {
                outState.selectionAnchorIndex = selection->first;
                outState.caretIndex           = selection->second;
            }
            return true;
        }

        return false;
    }

    [[nodiscard]] bool ApplyState(const TextInputState& state, bool notifyChange) noexcept
    {
        if (! _host || ! _control || state.readOnly)
        {
            return false;
        }

        if (auto* textField = dynamic_cast<TextField*>(_control))
        {
            if (notifyChange)
            {
                textField->SetTextAndNotify(state.text);
            }
            else
            {
                textField->SetText(state.text);
            }

            const size_t selectionStart = state.selectionAnchorIndex.value_or(state.caretIndex);
            textField->SetSelectionRange(selectionStart, state.caretIndex);
            _host->SyncTextInput(_control);
            _host->Invalidate();
            return true;
        }

        if (auto* comboBox = dynamic_cast<ComboBox*>(_control); comboBox && comboBox->IsEditable())
        {
            if (notifyChange)
            {
                comboBox->SetTextAndNotify(state.text);
            }
            else
            {
                comboBox->SetText(state.text);
            }

            const size_t selectionStart = state.selectionAnchorIndex.value_or(state.caretIndex);
            comboBox->SetEditableSelectionRange(selectionStart, state.caretIndex);
            _host->SyncTextInput(_control);
            _host->Invalidate();
            return true;
        }

        return false;
    }

    HRESULT ReplaceTextRange(DWORD dwFlags, LONG acpStart, LONG acpEnd, const WCHAR* pchText, ULONG cch, TS_TEXTCHANGE* pChange) noexcept
    {
        TextInputState state{};
        if (! ReadState(state))
        {
            return TS_E_INVALIDPOS;
        }
        if (state.readOnly)
        {
            return E_ACCESSDENIED;
        }
        if (cch > 0u && ! pchText)
        {
            return E_POINTER;
        }

        const AcpRange range = ClampAcpRange(acpStart, acpEnd, state.text.size());
        if ((dwFlags & TS_IAS_QUERYONLY) != 0u)
        {
            if (pChange)
            {
                pChange->acpStart  = ToAcp(range.start);
                pChange->acpOldEnd = ToAcp(range.end);
                pChange->acpNewEnd = ToAcp(range.start + cch);
            }
            return S_OK;
        }

        const std::wstring_view replacement(pchText ? pchText : L"", static_cast<size_t>(cch));
        state.text.replace(range.start, range.end - range.start, replacement);
        state.caretIndex = range.start + replacement.size();
        state.selectionAnchorIndex.reset();
        if (! ApplyState(state, true))
        {
            return E_FAIL;
        }

        TS_TEXTCHANGE change{};
        change.acpStart  = ToAcp(range.start);
        change.acpOldEnd = ToAcp(range.end);
        change.acpNewEnd = ToAcp(state.caretIndex);
        if (pChange)
        {
            *pChange = change;
        }

        NotifyTextChanged(change);
        NotifySelectionChanged();
        NotifyLayoutChanged();
        CaptureObservedState();
        return S_OK;
    }

    void CaptureObservedState() noexcept
    {
        TextInputState state{};
        if (! ReadState(state))
        {
            _hasObservedState = false;
            return;
        }

        _observedState    = std::move(state);
        _observedViewport = ResolveTextViewportBounds();
        _hasObservedState = true;
    }

    void NotifyExternalChangesIfNeeded() noexcept
    {
        if (! _sink)
        {
            return;
        }

        TextInputState currentState{};
        if (! ReadState(currentState))
        {
            _hasObservedState = false;
            return;
        }

        const D2D1_RECT_F currentViewport = ResolveTextViewportBounds();
        if (! _hasObservedState)
        {
            _observedState    = std::move(currentState);
            _observedViewport = currentViewport;
            _hasObservedState = true;
            return;
        }

        const bool textChanged      = currentState.text != _observedState.text;
        const bool selectionChanged = ! IsSameSelection(currentState, _observedState);
        const bool layoutChanged    = textChanged || selectionChanged || currentState.firstVisibleLine != _observedState.firstVisibleLine ||
                                      currentState.masked != _observedState.masked || currentState.multiline != _observedState.multiline ||
                                      ! IsSameRect(currentViewport, _observedViewport);

        const TS_TEXTCHANGE textChange{0, ToAcp(_observedState.text.size()), ToAcp(currentState.text.size())};

        _observedState    = std::move(currentState);
        _observedViewport = currentViewport;

        if (textChanged)
        {
            NotifyTextChanged(textChange);
        }
        if (selectionChanged)
        {
            NotifySelectionChanged();
        }
        if (layoutChanged)
        {
            NotifyLayoutChanged();
        }
    }

    void NotifyTextChanged(const TS_TEXTCHANGE& change) noexcept
    {
        if (_sink && (_sinkMask & TS_AS_TEXT_CHANGE) != 0u)
        {
            static_cast<void>(_sink->OnTextChange(0u, &change));
        }
    }

    void NotifySelectionChanged() noexcept
    {
        if (_sink && (_sinkMask & TS_AS_SEL_CHANGE) != 0u)
        {
            static_cast<void>(_sink->OnSelectionChange());
        }
    }

    void NotifyLayoutChanged() noexcept
    {
        if (_sink && (_sinkMask & TS_AS_LAYOUT_CHANGE) != 0u)
        {
            static_cast<void>(_sink->OnLayoutChange(TS_LC_CHANGE, kDxUiTextStoreView));
        }
    }

    [[nodiscard]] D2D1_RECT_F ResolveTextViewportBounds() const noexcept
    {
        if (! _control)
        {
            return D2D1::RectF();
        }
        if (const std::optional<D2D1_RECT_F> viewport = _control->TryGetTextInputViewportRect(); viewport.has_value())
        {
            return viewport.value();
        }
        return _control->GetHitBounds();
    }

    std::atomic<ULONG> _referenceCount{1u};
    WindowHost* _host = nullptr;
    Control* _control = nullptr;
    DWORD _lockFlags  = 0u;
    DWORD _sinkMask   = 0u;
    wil::com_ptr_nothrow<ITextStoreACPSink> _sink;
    TextInputState _observedState;
    D2D1_RECT_F _observedViewport = D2D1::RectF();
    bool _hasObservedState        = false;
};
} // namespace

ITextStoreACP* CreateNativeTextInputTextStore(WindowHost& host, Control& control) noexcept
{
    auto* store = new (std::nothrow) DxUiTextStoreACP(host, control);
    return store ? static_cast<ITextStoreACP*>(store) : nullptr;
}

#if defined(ENABLE_TESTS)
ITextStoreACP* WindowHost::DebugCreateNativeTextInputTextStoreForTest() noexcept
{
    if (_textInputBackend != TextInputBackend::Native || ! _nativeTextInputControl || ! _nativeTextInputStateCacheValid)
    {
        return nullptr;
    }

    return CreateNativeTextInputTextStore(*this, *_nativeTextInputControl);
}
#endif
} // namespace RedSalamander::DxUi
