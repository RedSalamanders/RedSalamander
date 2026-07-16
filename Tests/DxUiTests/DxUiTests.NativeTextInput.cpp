#include "DxUiTestHelpers.h"

#include <array>
#include <atomic>
#include <fstream>
#include <functional>
#include <iterator>
#include <string>
#include <string_view>

#include <msctf.h>
#include <textstor.h>

namespace
{

void TestNativeTextInputTextStoreGetTextExtUsesRangeRectsBeforeCaretWalk()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.TextStoreACP.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "TextStoreACP source is readable for GetTextExt range-rect guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t helperFunction = source.find("TryResolveMultilineTextStoreRangeRect");
    const size_t storeClass     = source.find("class DxUiTextStoreACP", helperFunction);
    Require(helperFunction != std::string::npos && storeClass != std::string::npos && helperFunction < storeClass,
            "multiline TextStoreACP range helper source block is found");

    const std::string helperBlock = source.substr(helperFunction, storeClass - helperFunction);
    const size_t rangeRects       = helperBlock.find("TryGetTextInputRangeRects");
    const size_t caretRect        = helperBlock.find("TryGetTextInputCaretRect");
    Require(rangeRects != std::string::npos, "multiline TextStoreACP GetTextExt uses control-provided range rects");
    Require(caretRect != std::string::npos && rangeRects < caretRect,
            "multiline TextStoreACP GetTextExt tries range rects before falling back to caret walking");
}

void TestNativeTextInputTsfDeactivateRestoresFocusAssociationBeforePoppingDocument()
{
    const std::filesystem::path internalHeaderPath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Internal.h";
    std::ifstream internalHeaderInput(internalHeaderPath);
    Require(internalHeaderInput.good(), "DxUi internal header is readable for native text store teardown guard");
    const std::string internalHeader((std::istreambuf_iterator<char>(internalHeaderInput)), std::istreambuf_iterator<char>());
    Require(internalHeader.find("DisconnectNativeTextInputTextStore") != std::string::npos,
            "DxUi exposes an internal text-store disconnect helper for teardown");

    const std::filesystem::path textStorePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.TextStoreACP.cpp";
    std::ifstream textStoreInput(textStorePath);
    Require(textStoreInput.good(), "TextStoreACP source is readable for teardown guard");
    const std::string textStoreSource((std::istreambuf_iterator<char>(textStoreInput)), std::istreambuf_iterator<char>());
    Require(textStoreSource.find("void DisconnectNativeTextInputTextStore") != std::string::npos, "TextStoreACP implements a disconnect helper");

    const size_t disconnectMethod = textStoreSource.find("void Disconnect() noexcept");
    const size_t compositionStart = textStoreSource.find("HRESULT STDMETHODCALLTYPE OnStartComposition", disconnectMethod);
    Require(disconnectMethod != std::string::npos && compositionStart != std::string::npos && disconnectMethod < compositionStart,
            "TextStoreACP disconnect method block is found");
    const std::string disconnectBlock = textStoreSource.substr(disconnectMethod, compositionStart - disconnectMethod);
    Require(disconnectBlock.find("_host") != std::string::npos && disconnectBlock.find("nullptr") != std::string::npos &&
                disconnectBlock.find("_control") != std::string::npos,
            "TextStoreACP disconnect severs raw host/control pointers");
    Require(disconnectBlock.find("SecureWipe::SecureClear(_observedState.text)") != std::string::npos,
            "TextStoreACP disconnect securely clears observed text before dropping state");

    const std::filesystem::path nativeSourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.NativeTextInput.cpp";
    std::ifstream nativeSourceInput(nativeSourcePath);
    Require(nativeSourceInput.good(), "native text input source is readable for TSF teardown guard");
    const std::string nativeSource((std::istreambuf_iterator<char>(nativeSourceInput)), std::istreambuf_iterator<char>());

    const size_t activate = nativeSource.find("bool WindowHost::ActivateNativeTextInputTsf");
    const size_t deactivate = nativeSource.find("void WindowHost::DeactivateNativeTextInputTsf", activate);
    Require(activate != std::string::npos && deactivate != std::string::npos && activate < deactivate,
            "native TSF activate/deactivate source blocks are found");
    const std::string activateBlock = nativeSource.substr(activate, deactivate - activate);
    const size_t associateHostFocus = activateBlock.find("threadMgr->AssociateFocus(_hwnd, documentMgr.get()");
    const size_t focusDocument      = activateBlock.find("threadMgr->SetFocus(documentMgr.get())");
    const size_t capturePreviousFocusDocument =
        activateBlock.find("threadMgr->AssociateFocus(_hwnd, documentMgr.get(), previousFocusDocumentMgr.put())");
    const size_t storePreviousFocusDocument =
        activateBlock.find("previousFocusDocumentMgr.query_to(_nativeTextInputTsfPreviousFocusDocumentMgr.put())");
    Require(associateHostFocus != std::string::npos, "native TSF activate associates the host HWND with the active document manager");
    Require(capturePreviousFocusDocument != std::string::npos,
            "native TSF activate captures the previous HWND focus document manager");
    Require(storePreviousFocusDocument != std::string::npos,
            "native TSF activate keeps the previous HWND focus document manager alive until deactivation");
    Require(focusDocument != std::string::npos, "native TSF activate sets thread-manager focus to the active document manager");
    Require(associateHostFocus < focusDocument, "native TSF activate associates the host HWND before setting TSF focus");
    Require(capturePreviousFocusDocument < storePreviousFocusDocument && storePreviousFocusDocument < focusDocument,
            "native TSF activate stores the previous HWND focus document manager before setting TSF focus");

    const size_t nextMethod = nativeSource.find("\nvoid WindowHost::SyncNativeTextInputSession", deactivate);
    Require(deactivate != std::string::npos && nextMethod != std::string::npos && deactivate < nextMethod, "native TSF deactivate source block is found");
    const std::string deactivateBlock = nativeSource.substr(deactivate, nextMethod - deactivate);

    const size_t earlyDisconnect        = deactivateBlock.find("DisconnectNativeTextInputTextStore(textStoreToDisconnect.get())");
    const size_t lateDisconnect = deactivateBlock.find("DisconnectNativeTextInputTextStore(textStoreToDisconnect.get())", earlyDisconnect + 1u);
    const size_t pop                    = deactivateBlock.find("documentMgr->Pop");
    const size_t clearThreadMgrFocus    = deactivateBlock.find("threadMgr->SetFocus(nullptr)");
    const size_t loadPreviousFocus      = deactivateBlock.find("_nativeTextInputTsfPreviousFocusDocumentMgr.query_to(previousFocusDocumentMgr.put())");
    const size_t restoreHostFocus       = deactivateBlock.find("threadMgr->AssociateFocus(_hwnd, previousFocusDocumentMgr.get()");
    const size_t resetPreviousFocus     = deactivateBlock.find("_nativeTextInputTsfPreviousFocusDocumentMgr.reset()");
    const size_t directNullDisassociate = deactivateBlock.find("threadMgr->AssociateFocus(_hwnd, nullptr");
    const size_t holdStoreForDisconnect = deactivateBlock.find("textStoreToDisconnect");
    const size_t resetContext           = deactivateBlock.find("_nativeTextInputTsfContext.reset()");
    const size_t resetDocumentMgr       = deactivateBlock.find("_nativeTextInputTsfDocumentMgr.reset()");
    const size_t resetThreadMgr         = deactivateBlock.find("_nativeTextInputTsfThreadMgr.reset()");
    const size_t resetStore             = deactivateBlock.find("_nativeTextInputTsfTextStore.reset()");
    Require(earlyDisconnect != std::string::npos && lateDisconnect != std::string::npos,
            "native TSF deactivate has separate stale-control and live-control disconnect paths");
    Require(clearThreadMgrFocus != std::string::npos, "native TSF deactivate clears the thread-manager focus");
    Require(loadPreviousFocus != std::string::npos, "native TSF deactivate loads the previous HWND focus document manager");
    Require(restoreHostFocus != std::string::npos, "native TSF deactivate restores the previous HWND focus document manager");
    Require(resetPreviousFocus != std::string::npos, "native TSF deactivate releases the saved previous HWND focus document manager");
    Require(directNullDisassociate == std::string::npos,
            "native TSF deactivate restores the previous HWND focus document manager instead of unconditionally clearing it");
    Require(pop != std::string::npos, "native TSF deactivate pops the active document");
    Require(earlyDisconnect < pop, "native TSF deactivate disconnects a stale control before document Pop can call back into it");
    Require(clearThreadMgrFocus < pop, "native TSF deactivate clears the thread-manager focus before popping the document");
    Require(loadPreviousFocus < restoreHostFocus && restoreHostFocus < pop,
            "native TSF deactivate restores the previous HWND focus document manager before popping the document");
    Require(pop < resetPreviousFocus, "native TSF deactivate keeps the saved previous focus document manager through document Pop");
    Require(holdStoreForDisconnect != std::string::npos && holdStoreForDisconnect < resetContext,
            "native TSF deactivate holds a local text-store reference while releasing TSF-owned context objects");
    Require(resetContext != std::string::npos && resetDocumentMgr != std::string::npos && resetThreadMgr != std::string::npos,
            "native TSF deactivate releases TSF context/document/thread-manager objects");
    Require(pop < resetContext && resetContext < resetDocumentMgr && resetDocumentMgr < resetThreadMgr && resetThreadMgr < lateDisconnect,
            "native TSF deactivate releases TSF-owned context objects before severing a live text-store sink/host connection");
    Require(resetStore != std::string::npos && lateDisconnect < resetStore,
            "native TSF deactivate disconnects the live text store before dropping the host reference");

    const size_t shutdown            = nativeSource.find("void ShutdownNativeTextInputThreadManagerState");
    const size_t ensureThreadManager = nativeSource.find("[[nodiscard]] bool EnsureNativeTextInputThreadManager", shutdown);
    Require(shutdown != std::string::npos && ensureThreadManager != std::string::npos && shutdown < ensureThreadManager,
            "native TSF thread-manager shutdown source block is found");
    const std::string shutdownBlock = nativeSource.substr(shutdown, ensureThreadManager - shutdown);
    Require(shutdownBlock.find("state.threadMgr->Deactivate()") != std::string::npos,
            "native TSF thread-manager shutdown owns the thread-manager Deactivate call");
}

class NativeTextStoreTestSink final : public ITextStoreACPSink
{
public:
    NativeTextStoreTestSink()                                          = default;
    NativeTextStoreTestSink(const NativeTextStoreTestSink&)            = delete;
    NativeTextStoreTestSink& operator=(const NativeTextStoreTestSink&) = delete;
    NativeTextStoreTestSink(NativeTextStoreTestSink&&)                 = delete;
    NativeTextStoreTestSink& operator=(NativeTextStoreTestSink&&)      = delete;

    std::function<HRESULT(DWORD)> onLockGranted;
    std::function<HRESULT(const TS_TEXTCHANGE*)> onTextChange;
    DWORD lastLockFlags                = 0u;
    uint32_t textChangeCount           = 0u;
    uint32_t selectionChangeCount      = 0u;
    uint32_t layoutChangeCount         = 0u;
    uint32_t editTransactionStartCount = 0u;
    uint32_t editTransactionEndCount   = 0u;
    uint32_t editTransactionDepth      = 0u;
    TS_TEXTCHANGE lastTextChange{};

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(ITextStoreACPSink))
        {
            *ppvObject = static_cast<ITextStoreACPSink*>(this);
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
        return _referenceCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
    }

    HRESULT STDMETHODCALLTYPE OnTextChange(DWORD /*dwFlags*/, const TS_TEXTCHANGE* pChange) noexcept override
    {
        ++textChangeCount;
        if (pChange)
        {
            lastTextChange = *pChange;
        }
        return onTextChange ? onTextChange(pChange) : S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnSelectionChange() noexcept override
    {
        ++selectionChangeCount;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnLayoutChange(TsLayoutCode /*lcode*/, TsViewCookie /*vcView*/) noexcept override
    {
        ++layoutChangeCount;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnStatusChange(DWORD /*dwFlags*/) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnAttrsChange(LONG /*acpStart*/, LONG /*acpEnd*/, ULONG /*cAttrs*/, const TS_ATTRID* /*paAttrs*/) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnLockGranted(DWORD dwLockFlags) noexcept override
    {
        lastLockFlags = dwLockFlags;
        return onLockGranted ? onLockGranted(dwLockFlags) : S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnStartEditTransaction() noexcept override
    {
        ++editTransactionStartCount;
        ++editTransactionDepth;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnEndEditTransaction() noexcept override
    {
        ++editTransactionEndCount;
        if (editTransactionDepth > 0u)
        {
            --editTransactionDepth;
        }
        return S_OK;
    }

private:
    std::atomic<ULONG> _referenceCount{1u};
};

void SendNativeKey(HWND hwnd, UINT virtualKey)
{
    static_cast<void>(SendMessageW(hwnd, WM_KEYDOWN, static_cast<WPARAM>(virtualKey), 0));
    static_cast<void>(SendMessageW(hwnd, WM_KEYUP, static_cast<WPARAM>(virtualKey), 0));
}

void SendNativeCtrlKey(HWND hwnd, UINT virtualKey)
{
    static_cast<void>(SendMessageW(hwnd, WM_KEYDOWN, VK_CONTROL, 0));
    SendNativeKey(hwnd, virtualKey);
    static_cast<void>(SendMessageW(hwnd, WM_KEYUP, VK_CONTROL, 0));
}

void SendNativeShiftKey(HWND hwnd, UINT virtualKey)
{
    static_cast<void>(SendMessageW(hwnd, WM_KEYDOWN, VK_SHIFT, 0));
    SendNativeKey(hwnd, virtualKey);
    static_cast<void>(SendMessageW(hwnd, WM_KEYUP, VK_SHIFT, 0));
}

void SendNativeAltKey(HWND hwnd, UINT virtualKey)
{
    static_cast<void>(SendMessageW(hwnd, WM_SYSKEYDOWN, static_cast<WPARAM>(virtualKey), 0));
    static_cast<void>(SendMessageW(hwnd, WM_SYSKEYUP, static_cast<WPARAM>(virtualKey), 0));
}

void SendNativeClick(HWND hwnd, POINT point)
{
    const LPARAM lp = MAKELPARAM(point.x, point.y);
    static_cast<void>(SendMessageW(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, lp));
    static_cast<void>(SendMessageW(hwnd, WM_LBUTTONUP, 0, lp));
}

[[nodiscard]] POINT DipPointToScreenPoint(AttachedHostWindow& window, D2D1_POINT_2F pointDip)
{
    POINT point{static_cast<LONG>(std::lround(window.Host().DipsToPixels(pointDip.x))), static_cast<LONG>(std::lround(window.Host().DipsToPixels(pointDip.y)))};
    MapWindowPoints(window.Hwnd(), nullptr, &point, 1);
    return point;
}

[[nodiscard]] RECT DipRectToScreenRect(AttachedHostWindow& window, const D2D1_RECT_F& rectDip)
{
    POINT points[2]{
        {static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.left))), static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.top)))},
        {static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.right))),
         static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.bottom)))}};
    MapWindowPoints(window.Hwnd(), nullptr, points, 2);
    return RECT{points[0].x, points[0].y, points[1].x, points[1].y};
}

[[nodiscard]] RECT DipRectToClientRect(AttachedHostWindow& window, const D2D1_RECT_F& rectDip)
{
    RECT rect = DipRectToScreenRect(window, rectDip);
    MapWindowPoints(nullptr, window.Hwnd(), reinterpret_cast<POINT*>(&rect), 2);
    return rect;
}

[[nodiscard]] POINT GetNativeTextClientPointForCaretIndex(AttachedHostWindow& window,
                                                          RedSalamander::DxUi::TextField& field,
                                                          size_t caretIndex,
                                                          const char* context)
{
    field.SetSelectionRange(caretIndex, caretIndex);
    window.Host().SyncTextInput(&field);

    D2D1_RECT_F caretRectDip{};
    RECT caretRectScreenPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(caretRectDip, caretRectScreenPx), context);
    MapWindowPoints(nullptr, window.Hwnd(), reinterpret_cast<POINT*>(&caretRectScreenPx), 2);
    return POINT{caretRectScreenPx.left + 2, (caretRectScreenPx.top + caretRectScreenPx.bottom) / 2};
}

struct DirectWritePointerProbe
{
    D2D1_POINT_2F pointDip{};
    size_t expectedCaretIndex = 0u;
};

using DirectWritePointerSpan = std::pair<size_t, size_t>;

[[nodiscard]] float DirectWriteCaretOffsetDip(IDWriteTextLayout* layout, size_t textSize, size_t caretIndex, const char* context)
{
    Require(layout != nullptr, context);
    const size_t clampedCaret = (std::min)(caretIndex, textSize);
    float x                   = 0.0f;
    float y                   = 0.0f;
    DWRITE_HIT_TEST_METRICS positionMetrics{};
    const UINT32 textPosition = static_cast<UINT32>(clampedCaret == 0u ? 0u : clampedCaret - 1u);
    const BOOL trailingHit    = clampedCaret == 0u ? FALSE : TRUE;
    RequireSucceeded(layout->HitTestTextPosition(textPosition, trailingHit, &x, &y, &positionMetrics), context);
    return x;
}

[[nodiscard]] DirectWritePointerProbe CreateDirectWritePointerProbeForTextSpan(RedSalamander::DxUi::WindowHost& host,
                                                                               std::wstring_view text,
                                                                               const D2D1_RECT_F& textRect,
                                                                               RedSalamander::DxUi::FlowDirection flowDirection,
                                                                               size_t spanStart,
                                                                               size_t spanEnd,
                                                                               const char* context)
{
    using namespace RedSalamander::DxUi;

    Require(spanStart < spanEnd && spanEnd <= text.size(), context);
    IDWriteFactory* factory = host.GetWriteFactory();
    Require(factory != nullptr, context);
    const DWRITE_READING_DIRECTION readingDirection = ResolveReadingDirection(flowDirection);
    IDWriteTextFormat* format = host.GetTextFormat(FontRole::Body, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, readingDirection);
    Require(format != nullptr, context);

    const float layoutWidthDip  = std::max(1.0f, textRect.right - textRect.left);
    const float layoutHeightDip = std::max(1.0f, textRect.bottom - textRect.top);
    wil::com_ptr<IDWriteTextLayout> layout;
    RequireSucceeded(factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, layoutWidthDip, layoutHeightDip, layout.addressof()),
                     context);
    Require(layout != nullptr, context);

    const float spanStartX              = DirectWriteCaretOffsetDip(layout.get(), text.size(), spanStart, context);
    const float spanEndX                = DirectWriteCaretOffsetDip(layout.get(), text.size(), spanEnd, context);
    const float localX                  = std::clamp((spanStartX + spanEndX) * 0.5f, 0.0f, std::max(0.0f, layoutWidthDip - 1.0f));
    const float localY                  = layoutHeightDip * 0.5f;
    const D2D1_POINT_2F clickedPointDip = D2D1::Point2F(host.PixelsToDip(static_cast<float>(std::round(host.DipsToPixels(textRect.left + localX)))),
                                                        host.PixelsToDip(static_cast<float>(std::round(host.DipsToPixels(textRect.top + localY)))));
    const float clickedLocalX           = std::clamp(clickedPointDip.x - textRect.left, 0.0f, std::max(0.0f, layoutWidthDip - 1.0f));
    const float clickedLocalY           = std::clamp(clickedPointDip.y - textRect.top, 0.0f, std::max(1.0f, layoutHeightDip) - 1.0f);
    BOOL isTrailingHit                  = FALSE;
    BOOL isInside                       = FALSE;
    DWRITE_HIT_TEST_METRICS pointMetrics{};
    RequireSucceeded(layout->HitTestPoint(clickedLocalX, clickedLocalY, &isTrailingHit, &isInside, &pointMetrics), context);
    Require(isInside != FALSE, context);

    const size_t expectedIndex =
        SnapCaretIndexToTextElementBoundary(text, static_cast<size_t>(pointMetrics.textPosition) + (isTrailingHit ? pointMetrics.length : 0u));
    return DirectWritePointerProbe{clickedPointDip, expectedIndex};
}

void VerifyNativeTextInputPointerScenarioMatchesDirectWrite(std::wstring_view text,
                                                            RedSalamander::DxUi::FlowDirection flowDirection,
                                                            const std::array<DirectWritePointerSpan, 3u>& spans,
                                                            const char* context)
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root = std::make_unique<Panel>();
    root->SetFlowDirection(flowDirection);
    auto* field = root->AddChild<TextField>(std::wstring(text));
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 340.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, context);
    Require(window.Host().GetTextInputHwnd() == window.Hwnd(), context);

    TextFieldDebugSingleLinePaintState paint{};
    Require(field->DebugGetSingleLinePaintState(window.Host(), paint), context);

    std::array<DirectWritePointerProbe, 3u> probes{
        CreateDirectWritePointerProbeForTextSpan(window.Host(), text, paint.textRect, flowDirection, spans[0].first, spans[0].second, context),
        CreateDirectWritePointerProbeForTextSpan(window.Host(), text, paint.textRect, flowDirection, spans[1].first, spans[1].second, context),
        CreateDirectWritePointerProbeForTextSpan(window.Host(), text, paint.textRect, flowDirection, spans[2].first, spans[2].second, context),
    };
    std::sort(probes.begin(), probes.end(), [](const DirectWritePointerProbe& lhs, const DirectWritePointerProbe& rhs) noexcept {
        return lhs.pointDip.x < rhs.pointDip.x;
    });
    Require(probes[0].pointDip.x < probes[1].pointDip.x && probes[1].pointDip.x < probes[2].pointDip.x, context);

    for (size_t probeIndex = 0u; probeIndex < probes.size(); ++probeIndex)
    {
        const DirectWritePointerProbe& probe = probes[probeIndex];
        const POINT clickPoint{static_cast<LONG>(std::lround(window.Host().DipsToPixels(probe.pointDip.x))),
                               static_cast<LONG>(std::lround(window.Host().DipsToPixels(probe.pointDip.y)))};
        SendNativeClick(window.Hwnd(), clickPoint);

        NativeTextInputState state{};
        Require(window.Host().TryReadNativeTextInputState(field, state), context);
        Require(state.selectionAnchorIndex.value_or(state.caretIndex) == state.caretIndex, context);
        static_cast<void>(probeIndex);
        Require(state.caretIndex == probe.expectedCaretIndex, context);
    }
}

void RequireNativeTextSelection(
    RedSalamander::DxUi::WindowHost& host, RedSalamander::DxUi::TextField& field, size_t expectedStart, size_t expectedEnd, const char* context)
{
    const std::optional<std::pair<size_t, size_t>> selection = field.GetSelectionRange();
    Require(selection.has_value(), context);
    Require(selection.value().first == expectedStart && selection.value().second == expectedEnd, context);

    RedSalamander::DxUi::NativeTextInputState state{};
    Require(host.TryReadNativeTextInputState(&field, state), context);
    Require(state.selectionAnchorIndex.has_value(), context);
    const size_t stateSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t stateSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(stateSelectionStart == expectedStart && stateSelectionEnd == expectedEnd, context);
}

[[nodiscard]] std::wstring MakeWomanTechnologistTextElement()
{
    std::wstring text;
    text.push_back(static_cast<wchar_t>(0xD83D));
    text.push_back(static_cast<wchar_t>(0xDC69));
    text.push_back(static_cast<wchar_t>(0x200D));
    text.push_back(static_cast<wchar_t>(0xD83D));
    text.push_back(static_cast<wchar_t>(0xDCBB));
    return text;
}

[[nodiscard]] std::wstring MakeGrinningFaceTextElement()
{
    std::wstring text;
    text.push_back(static_cast<wchar_t>(0xD83D));
    text.push_back(static_cast<wchar_t>(0xDE00));
    return text;
}

[[nodiscard]] std::wstring MakeRainbowFlagTextElement()
{
    std::wstring text;
    text.push_back(static_cast<wchar_t>(0xD83C));
    text.push_back(static_cast<wchar_t>(0xDFF3));
    text.push_back(static_cast<wchar_t>(0xFE0F));
    text.push_back(static_cast<wchar_t>(0x200D));
    text.push_back(static_cast<wchar_t>(0xD83C));
    text.push_back(static_cast<wchar_t>(0xDF08));
    return text;
}

[[nodiscard]] std::wstring MakeUsFlagTextElement()
{
    std::wstring text;
    text.push_back(static_cast<wchar_t>(0xD83C));
    text.push_back(static_cast<wchar_t>(0xDDFA));
    text.push_back(static_cast<wchar_t>(0xD83C));
    text.push_back(static_cast<wchar_t>(0xDDF8));
    return text;
}

[[nodiscard]] std::wstring MakeHeartVariationTextElement()
{
    std::wstring text;
    text.push_back(static_cast<wchar_t>(0x2764));
    text.push_back(static_cast<wchar_t>(0xFE0F));
    return text;
}

[[nodiscard]] std::wstring MakeThumbsUpMediumSkinToneTextElement()
{
    std::wstring text;
    text.push_back(static_cast<wchar_t>(0xD83D));
    text.push_back(static_cast<wchar_t>(0xDC4D));
    text.push_back(static_cast<wchar_t>(0xD83C));
    text.push_back(static_cast<wchar_t>(0xDFFD));
    return text;
}

void TestWindowHostDefaultsToNativeTextInputBackend()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    Require(window.Host().GetTextInputBackend() == TextInputBackend::Native, "window host defaults to native text input backend");
    Require(! window.Host().DebugHasActiveNativeTextInputSession(), "native text input session is inactive by default");
}

void TestNativeTextInputBackendFocusesHostWithoutBridgeChild()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    if (! TryFocusDxUiTestWindow(window.Hwnd()))
    {
        SkipDxUiTest("native text input requires an interactive desktop for host focus assertions");
        return;
    }
    window.Host().SetFocusControl(field);
    window.PumpMessages();

    Require(window.Host().GetTextInputBackend() == TextInputBackend::Native, "window host keeps the native backend after text focus");
    Require(window.Host().GetFocusControl() == field, "native text input keeps retained focus on the text field");
    Require(GetFocus() == window.Hwnd(), "native text input focuses the host hwnd instead of a child bridge");
    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native text input does not create a hidden bridge child window");
    Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native text input reports the host hwnd as its active input hwnd");
    Require(window.Host().DebugHasActiveNativeTextInputSession(), "native text input session is active while the text field is focused");
}

void TestNativeTextInputBackendActivatesTsfDocumentOnFocus()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    const NativeTextInputEventCounters beforeCounters = window.Host().DebugGetNativeTextInputEventCounters();

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 260.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    window.PumpMessages();

    const NativeTextInputEventCounters counters = window.Host().DebugGetNativeTextInputEventCounters();
    Require(counters.tsfActivationAttemptCount > beforeCounters.tsfActivationAttemptCount,
            "native text input attempts TSF document activation when a text field gains focus");
    Require(counters.tsfActivationSuccessCount > beforeCounters.tsfActivationSuccessCount,
            "native text input activates a TSF document manager/context for the focused field");
    Require(window.Host().DebugHasActiveNativeTextInputTsfDocument(), "native text input keeps a TSF document/context alive while the field is focused");

    window.Host().SetFocusControl(nullptr);
    Require(! window.Host().DebugHasActiveNativeTextInputTsfDocument(), "native text input releases the TSF document/context when focus leaves the field");
    Require(window.Host().DebugGetNativeTextInputEventCounters().tsfDeactivationCount > counters.tsfDeactivationCount,
            "native text input counts TSF document deactivation when focus leaves the field");
}

void TestNativeTextInputBackendOwnsSystemCaretOnHostHwnd()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    if (! TryFocusDxUiTestWindow(window.Hwnd()))
    {
        SkipDxUiTest("native text input requires an interactive desktop for system caret assertions");
        return;
    }
    window.Host().SetFocusControl(field);
    window.PumpMessages();

    D2D1_RECT_F caretRectDip{};
    RECT caretRectScreenPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(caretRectDip, caretRectScreenPx), "native text input exposes a caret debug rect");
    RequireRectHasArea(caretRectDip, "native text input caret DIP rect has area");
    Require(caretRectScreenPx.right > caretRectScreenPx.left && caretRectScreenPx.bottom > caretRectScreenPx.top,
            "native text input caret screen rect has area");

    GUITHREADINFO guiThreadInfo{};
    guiThreadInfo.cbSize = sizeof(guiThreadInfo);
    Require(GetGUIThreadInfo(GetCurrentThreadId(), &guiThreadInfo) != FALSE, "native text input can read GUI thread caret state");
    Require(guiThreadInfo.hwndCaret == window.Hwnd(), "native text input creates the system caret on the host hwnd");
    RECT guiCaretScreenRectPx = guiThreadInfo.rcCaret;
    MapWindowPoints(guiThreadInfo.hwndCaret, nullptr, reinterpret_cast<POINT*>(&guiCaretScreenRectPx), 2);
    RequireRectNear(guiCaretScreenRectPx, caretRectScreenPx, "native text input system caret follows the debug screen rect");
}

void TestNativeTextInputBackendMovesSystemCaretAfterKeyInput()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    D2D1_RECT_F beforeRectDip{};
    RECT beforeRectScreenPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(beforeRectDip, beforeRectScreenPx), "native text input exposes the starting caret rect");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_LEFT, 0));

    D2D1_RECT_F afterRectDip{};
    RECT afterRectScreenPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(afterRectDip, afterRectScreenPx), "native text input exposes the moved caret rect");
    Require(afterRectDip.left < beforeRectDip.left, "native text input caret DIP rect moves left after left arrow");
    Require(afterRectScreenPx.left < beforeRectScreenPx.left, "native text input caret screen rect moves left after left arrow");
}

void TestNativeTextInputBackendClearsSessionWhenRootResets()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    Require(window.Host().DebugHasActiveNativeTextInputSession(), "native text input session is active before root reset");

    D2D1_RECT_F caretRectDip{};
    RECT caretRectScreenPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(caretRectDip, caretRectScreenPx), "native text input caret is active before root reset");

    window.Host().SetRoot(std::make_unique<Panel>());

    Require(! window.Host().DebugHasActiveNativeTextInputSession(), "native text input session clears after root reset");
    Require(! window.Host().DebugGetNativeTextInputCaretRect(caretRectDip, caretRectScreenPx), "native text input caret clears after root reset");
}

void TestNativeTextInputBackendClearsSessionWhenFocusedControlBecomesStale()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    Require(window.Host().DebugHasActiveNativeTextInputSession(), "native text input session is active before focused control is disabled");

    field->SetEnabled(false);
    window.Host().SetFocusControl(field);

    D2D1_RECT_F caretRectDip{};
    RECT caretRectScreenPx{};
    Require(! window.Host().DebugHasActiveNativeTextInputSession(), "native text input session clears after focused control is disabled");
    Require(! window.Host().DebugGetNativeTextInputCaretRect(caretRectDip, caretRectScreenPx),
            "native text input caret clears after focused control is disabled");
}

void TestNativeTextInputBackendUpdatesCaretWhenFocusedFieldMoves()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    D2D1_RECT_F beforeRectDip{};
    RECT beforeRectScreenPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(beforeRectDip, beforeRectScreenPx), "native text input exposes caret before focused field moves");

    field->SetBounds(D2D1::RectF(52.0f, 16.0f, 260.0f, 44.0f));

    D2D1_RECT_F afterRectDip{};
    RECT afterRectScreenPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(afterRectDip, afterRectScreenPx), "native text input exposes caret after focused field moves");
    Require(afterRectDip.left > beforeRectDip.left, "native text input caret DIP rect follows focused field movement");
    Require(afterRectScreenPx.left > beforeRectScreenPx.left, "native text input caret screen rect follows focused field movement");
}

void TestNativeTextInputBackendHostFocusLossControlsNativeSession()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    AttachedHostWindow externalWindow;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    Require(window.Host().GetFocusControl() == field, "native focus-loss test starts with retained text-field focus");
    Require(field->HasFocus(), "native focus-loss test starts with active focus visuals");
    Require(window.Host().DebugHasActiveNativeTextInputSession(), "native focus-loss test starts with an active native text session");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KILLFOCUS, reinterpret_cast<WPARAM>(window.Hwnd()), 0));
    Require(window.Host().GetFocusControl() == field, "native focus transfer inside the host keeps retained text-field focus");
    Require(field->HasFocus(), "native focus transfer inside the host keeps active focus visuals");
    Require(window.Host().DebugHasActiveNativeTextInputSession(), "native focus transfer inside the host keeps the native text session active");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KILLFOCUS, reinterpret_cast<WPARAM>(externalWindow.Hwnd()), 0));
    Require(window.Host().GetFocusControl() == field, "native external focus loss keeps retained logical text-field focus");
    Require(! field->HasFocus(), "native external focus loss clears active focus visuals");
    Require(! window.Host().DebugHasActiveNativeTextInputSession(), "native external focus loss deactivates the native text session");
    Require(! window.Host().HasActiveTextInput(), "native external focus loss clears backend-neutral active text input");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_SETFOCUS, 0, 0));
    Require(window.Host().GetFocusControl() == field, "native focus regain preserves the retained logical text-field focus");
    Require(field->HasFocus(), "native focus regain restores active focus visuals");
    Require(window.Host().DebugHasActiveNativeTextInputSession(), "native focus regain reactivates the native text session");
}

void TestNativeTextInputBackendTabMovesFocusToNextControl()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<TextField>(L"alpha");
    auto* button = root->AddChild<Button>(L"Next");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));
    button->SetBounds(D2D1::RectF(12.0f, 56.0f, 120.0f, 88.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));

    Require(window.Host().GetFocusControl() == button, "native text input tab moves focus to the next focusable control");
    Require(! window.Host().DebugHasActiveNativeTextInputSession(), "native text input session deactivates after tab leaves the text field");
}

void TestNativeTextInputBackendMultilineDialogKeysStayHostOwned()
{
    using namespace RedSalamander::DxUi;

    const auto runCase = [](bool wrapped)
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root            = std::make_unique<Panel>();
        auto* previousButton = root->AddChild<Button>(L"Previous");
        auto* field          = root->AddChild<TextField>(wrapped ? std::wstring(kWrappedMultilineClipboardTextForTest) : std::wstring(L"alpha"));
        auto* nextButton     = root->AddChild<Button>(L"Next");
        auto* cancelButton   = root->AddChild<Button>(L"Cancel");
        previousButton->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
        field->SetMultiline(true);
        field->SetBounds(wrapped ? D2D1::RectF(0.0f, 40.0f, 120.0f, 136.0f) : D2D1::RectF(0.0f, 40.0f, 220.0f, 136.0f));
        nextButton->SetBounds(D2D1::RectF(0.0f, 152.0f, 120.0f, 180.0f));
        cancelButton->SetBounds(D2D1::RectF(140.0f, 152.0f, 260.0f, 180.0f));

        size_t defaultCount = 0u;
        size_t cancelCount  = 0u;
        nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
        cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

        window.Host().SetRoot(std::move(root));
        window.Host().SetDefaultButton(nextButton);
        window.Host().SetCancelButton(cancelButton);
        window.Host().SetFocusControl(field);

        const std::wstring initialText(field->GetText());
        field->SetSelectionRange(field->GetText().size(), field->GetText().size());
        window.Host().SyncTextInput(field);

        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_RETURN, 0));
        Require(defaultCount == 0u,
                wrapped ? "native wrapped multiline return stays text-owned instead of invoking the default button"
                        : "native multiline return stays text-owned instead of invoking the default button");
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, static_cast<WPARAM>(L'\r'), 0));
        Require(field->GetText() == initialText + L"\n",
                wrapped ? "native wrapped multiline return inserts a logical newline" : "native multiline return inserts a logical newline");
        Require(window.Host().GetFocusControl() == field,
                wrapped ? "native wrapped multiline return keeps focus on the text field" : "native multiline return keeps focus on the text field");
        Require(window.Host().DebugHasActiveNativeTextInputSession(),
                wrapped ? "native wrapped multiline return keeps the native text session active"
                        : "native multiline return keeps the native text session active");

        SendNativeShiftKey(window.Hwnd(), VK_TAB);
        Require(window.Host().GetFocusControl() == previousButton,
                wrapped ? "native wrapped multiline shift+tab moves focus to the previous control"
                        : "native multiline shift+tab moves focus to the previous control");
        Require(! window.Host().DebugHasActiveNativeTextInputSession(),
                wrapped ? "native wrapped multiline shift+tab deactivates the native text session"
                        : "native multiline shift+tab deactivates the native text session");

        window.Host().SetFocusControl(field);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_ESCAPE, 0));
        Require(cancelCount == 1u, wrapped ? "native wrapped multiline escape invokes the cancel button" : "native multiline escape invokes the cancel button");
        Require(window.Host().GetFocusControl() == field,
                wrapped ? "native wrapped multiline escape keeps retained focus on the text field"
                        : "native multiline escape keeps retained focus on the text field");
        Require(window.Host().DebugHasActiveNativeTextInputSession(),
                wrapped ? "native wrapped multiline escape keeps the native text session active"
                        : "native multiline escape keeps the native text session active");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));
        Require(window.Host().GetFocusControl() == nextButton,
                wrapped ? "native wrapped multiline tab moves focus to the next control" : "native multiline tab moves focus to the next control");
        Require(! window.Host().DebugHasActiveNativeTextInputSession(),
                wrapped ? "native wrapped multiline tab deactivates the native text session" : "native multiline tab deactivates the native text session");
    };

    runCase(false);
    runCase(true);
}

void TestNativeTextInputBackendEnterInvokesDefaultButton()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    int invokeCount = 0;
    auto root       = std::make_unique<Panel>();
    auto* field     = root->AddChild<TextField>(L"alpha");
    auto* button    = root->AddChild<Button>(L"OK");
    button->SetOnClick([&] { ++invokeCount; });
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));
    button->SetBounds(D2D1::RectF(12.0f, 56.0f, 120.0f, 88.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(button);
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_RETURN, 0));

    Require(invokeCount == 1, "native text input enter invokes the default button for single-line fields");
    Require(field->GetText() == L"alpha", "native text input enter does not mutate single-line text");
}

void TestNativeTextInputBackendEscapeInvokesCancelButton()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    int invokeCount = 0;
    auto root       = std::make_unique<Panel>();
    auto* field     = root->AddChild<TextField>(L"alpha");
    auto* button    = root->AddChild<Button>(L"Cancel");
    button->SetOnClick([&] { ++invokeCount; });
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));
    button->SetBounds(D2D1::RectF(12.0f, 56.0f, 120.0f, 88.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetCancelButton(button);
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_ESCAPE, 0));

    Require(invokeCount == 1, "native text input escape invokes the cancel button for single-line fields");
}

void TestNativeTextInputBackendRevealedMaskedFieldRemasksBeforeEscapeCancel()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    int invokeCount = 0;
    auto root       = std::make_unique<Panel>();
    auto* field     = root->AddChild<TextField>(L"secret");
    field->SetMasked(true);
    field->SetPasswordRevealState(PasswordRevealState::Visible);
    auto* button = root->AddChild<Button>(L"Cancel");
    button->SetOnClick([&] { ++invokeCount; });
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));
    button->SetBounds(D2D1::RectF(12.0f, 56.0f, 120.0f, 88.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetCancelButton(button);
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_ESCAPE, 0));

    Require(invokeCount == 1, "native revealed masked field escape still invokes the cancel button");
    Require(field->GetPasswordRevealState() == PasswordRevealState::Hidden, "native revealed masked field remasks before escape cancel");

    NativeTextInputState state{};
    Require(window.Host().TryReadNativeTextInputState(field, state), "native revealed masked field exposes state after escape cancel");
    Require(state.passwordRevealState == PasswordRevealState::Hidden && state.secretVisibleDotCount == 6u,
            "native revealed masked field syncs hidden state after escape cancel");
}

void TestNativeTextInputBackendMenuKeyInvokesContextMenu()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    int contextMenuCount    = 0;
    bool keyboardInvocation = false;
    auto root               = std::make_unique<Panel>();
    auto* field             = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));
    field->SetOnContextMenu([&](POINT /*screenPoint*/, bool invokedByKeyboard)
    {
        ++contextMenuCount;
        keyboardInvocation = invokedByKeyboard;
    });

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_APPS, 0));

    Require(contextMenuCount == 1, "native text input menu key invokes the focused text field context menu");
    Require(keyboardInvocation, "native text input menu key reports keyboard context-menu invocation");
}

void TestNativeTextInputBackendMultilineContextMenuKeysStayOnHostHwnd()
{
    using namespace RedSalamander::DxUi;

    const auto runScenario = [](bool wrapped, bool shiftF10)
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        RecordingContextMenuInvocation contextMenu;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(wrapped ? std::wstring(kWrappedMultilineClipboardTextForTest) : std::wstring(L"alpha\nbeta"));
        field->SetMultiline(true);
        const D2D1_RECT_F fieldBounds = wrapped ? D2D1::RectF(12.0f, 16.0f, 132.0f, 112.0f) : D2D1::RectF(12.0f, 16.0f, 232.0f, 112.0f);
        field->SetBounds(fieldBounds);
        field->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native multiline context-menu route does not create a hidden bridge child");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native multiline context-menu route keeps the host hwnd as the input target");

        if (shiftF10)
        {
            static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_SHIFT, 0));
            static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_F10, 0));
            static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYUP, VK_SHIFT, 0));
        }
        else
        {
            static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_APPS, 0));
        }

        Require(contextMenu.count == 1u, "native multiline context-menu key invokes the focused field menu once");
        Require(contextMenu.lastKeyboardInvocation, "native multiline context-menu key reports keyboard invocation");
        const RECT fieldScreen = DipRectToScreenRect(window, fieldBounds);
        Require(contextMenu.lastPoint.x >= fieldScreen.left && contextMenu.lastPoint.x <= fieldScreen.right,
                "native multiline context-menu key anchor stays inside the field horizontally");
        Require(contextMenu.lastPoint.y >= fieldScreen.top && contextMenu.lastPoint.y <= fieldScreen.bottom,
                "native multiline context-menu key anchor stays inside the field vertically");
        Require(window.Host().GetFocusControl() == field, "native multiline context-menu key keeps focus on the text field");

        NativeTextInputState state{};
        Require(window.Host().TryReadNativeTextInputState(field, state), "native multiline context-menu route keeps session state readable");
        Require(state.multiline, "native multiline context-menu route keeps multiline session state");
    };

    runScenario(false, false);
    runScenario(false, true);
    runScenario(true, false);
    runScenario(true, true);
}

void TestNativeTextInputBackendImeStartEndUpdatesCompositionState()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable before ime composition starts");
    Require(! state.compositionStartIndex.has_value() && ! state.compositionEndIndex.has_value(),
            "native session state starts without an ime composition range");

    const size_t startingCaretIndex = state.caretIndex;
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after ime composition starts");
    Require(state.compositionStartIndex.has_value() && state.compositionStartIndex.value() == startingCaretIndex,
            "native ime start records the composition start at the current caret");
    Require(state.compositionEndIndex.has_value() && state.compositionEndIndex.value() == startingCaretIndex,
            "native ime start records an empty composition range at the current caret");
    Require(! state.conversionTargetStartIndex.has_value() && ! state.conversionTargetEndIndex.has_value(),
            "native ime start does not report a conversion target before composition attributes arrive");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_ENDCOMPOSITION, 0, 0));

    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after ime composition ends");
    Require(! state.compositionStartIndex.has_value() && ! state.compositionEndIndex.has_value(), "native ime end clears the composition range");
    Require(! state.conversionTargetStartIndex.has_value() && ! state.conversionTargetEndIndex.has_value(),
            "native ime end clears the conversion-target range");
}

void TestNativeTextInputBackendReadOnlySuppressesImeComposition()
{
    using namespace RedSalamander::DxUi;

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha");
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native read-only session state is readable before ime composition starts");
        Require(state.readOnly, "native read-only session state mirrors read-only text input");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

        Require(window.Host().TryReadNativeTextInputState(field, state), "native read-only session state is readable after ime composition start");
        Require(! state.compositionStartIndex.has_value() && ! state.compositionEndIndex.has_value(),
                "native read-only text input suppresses ime composition ranges");
        Require(! state.conversionTargetStartIndex.has_value() && ! state.conversionTargetEndIndex.has_value(),
                "native read-only text input suppresses ime conversion-target ranges");
    }

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha");
        field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native mutable session state is readable before read-only changes");
        Require(! state.readOnly && ! state.masked, "native mutable session starts editable and unmasked");

        field->SetReadOnly(true);
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after focused read-only toggle");
        Require(state.readOnly, "native session state mirrors focused read-only toggles before the next IME message");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));
        Require(window.Host().TryReadNativeTextInputState(field, state), "native focused read-only toggle test reads state after IME start");
        Require(! state.compositionStartIndex.has_value() && ! state.compositionEndIndex.has_value(),
                "native focused read-only toggles suppress the next IME start message");

        field->SetReadOnly(false);
        field->SetMasked(true);
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after focused masked toggle");
        Require(! state.readOnly && state.masked, "native session state mirrors focused masked toggles before the next native input message");
    }
}

void TestNativeTextInputBackendImeStartTracksSelectedCompositionRange()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 260.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(0u, 5u);
    window.Host().SyncTextInput(field);

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native selected session state is readable before ime composition");
    Require(state.selectionAnchorIndex.has_value() && state.selectionAnchorIndex.value() == 0u && state.caretIndex == 5u,
            "native selected session state mirrors the selected range before ime composition");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    Require(window.Host().TryReadNativeTextInputState(field, state), "native selected session state is readable after ime composition starts");
    Require(state.compositionStartIndex.has_value() && state.compositionStartIndex.value() == 0u,
            "native ime composition over selection starts at the selection start");
    Require(state.compositionEndIndex.has_value() && state.compositionEndIndex.value() == 5u,
            "native ime composition over selection ends at the selection end");
}

void TestNativeTextInputBackendImeNoPayloadWithoutActiveCompositionDoesNotStartRange()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 220.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, 0));

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after ime composition without payload");
    Require(! state.compositionStartIndex.has_value() && ! state.compositionEndIndex.has_value(),
            "native ime composition without payload does not create a composition range without an active composition");
    Require(! state.conversionTargetStartIndex.has_value() && ! state.conversionTargetEndIndex.has_value(),
            "native ime composition without payload does not create a conversion range without an active composition");
}

void TestNativeTextInputBackendImeWindowsTrackCaretRect()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(18.0f, 14.0f, 238.0f, 42.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(4u, 4u);
    window.Host().SyncTextInput(field);

    D2D1_RECT_F caretRectDip{};
    RECT caretRectClientPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(caretRectDip, caretRectClientPx), "native ime anchor test can read the native caret rect");
    MapWindowPoints(nullptr, window.Hwnd(), reinterpret_cast<POINT*>(&caretRectClientPx), 2);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    const auto compositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
    Require(compositionForm.has_value(), "native ime composition form is readable from the host hwnd");
    Require(compositionForm.value().dwStyle == CFS_FORCE_POSITION, "native ime composition form forces the caret position");
    RequirePointNear(compositionForm.value().ptCurrentPos,
                     POINT{caretRectClientPx.left, caretRectClientPx.top},
                     "native ime composition point tracks the host-owned caret rect");

    const auto candidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(candidateForm.has_value(), "native ime candidate form is readable from the host hwnd");
    Require(candidateForm.value().dwStyle == CFS_EXCLUDE, "native ime candidate form excludes the caret rect");
    RequirePointNear(candidateForm.value().ptCurrentPos,
                     POINT{caretRectClientPx.left, caretRectClientPx.bottom},
                     "native ime candidate point tracks the host-owned caret baseline");
    RequireRectNear(candidateForm.value().rcArea, caretRectClientPx, "native ime candidate exclusion rect tracks the host-owned caret rect");
}

void TestNativeTextInputBackendMultilineImeWindowsTrackCaretAcrossLines()
{
    using namespace RedSalamander::DxUi;

    const auto runCase = [](bool wrapped)
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(wrapped ? std::wstring(kWrappedMultilineClipboardTextForTest) : std::wstring(L"alpha\nbeta\ngamma"));
        field->SetMultiline(true);
        field->SetBounds(wrapped ? D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f) : D2D1::RectF(16.0f, 18.0f, 276.0f, 162.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        const size_t firstIndex = 1u;
        D2D1_RECT_F firstCaretDip{};
        Require(field->DebugGetCaretRect(window.Host(), firstIndex, firstCaretDip),
                wrapped ? "native wrapped ime anchor test can measure the first visual-line caret"
                        : "native multiline ime anchor test can measure the first-line caret");

        std::optional<size_t> laterIndex;
        if (wrapped)
        {
            const auto text = field->GetText();
            for (size_t candidateIndex = firstIndex + 1u; candidateIndex < text.size(); ++candidateIndex)
            {
                D2D1_RECT_F candidateCaretDip{};
                Require(field->DebugGetCaretRect(window.Host(), candidateIndex, candidateCaretDip),
                        "native wrapped ime anchor test can measure candidate wrapped carets");
                if (candidateCaretDip.top > firstCaretDip.top + 1.0f)
                {
                    laterIndex = candidateIndex;
                    break;
                }
            }
            Require(laterIndex.has_value(), "native wrapped ime anchor test can locate a caret on a later wrapped visual line");
        }
        else
        {
            laterIndex = field->GetText().find(L"gamma");
            Require(laterIndex.has_value() && laterIndex.value() != std::wstring::npos, "native multiline ime anchor test can locate a later-line token");
        }

        const auto requireImeFormsAtCaret = [&](size_t caretIndex, const char* context)
        {
            const std::optional<D2D1_RECT_F> expectedCaretDip = field->TryGetTextInputCaretRect(window.Host(), caretIndex);
            Require(expectedCaretDip.has_value(), context);
            const RECT expectedCaretClientPx = DipRectToClientRect(window, expectedCaretDip.value());

            const auto compositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
            Require(compositionForm.has_value(), context);
            RequirePointNear(compositionForm.value().ptCurrentPos, POINT{expectedCaretClientPx.left, expectedCaretClientPx.top}, context);

            const auto candidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
            Require(candidateForm.has_value(), context);
            RequirePointNear(candidateForm.value().ptCurrentPos, POINT{expectedCaretClientPx.left, expectedCaretClientPx.bottom}, context);
            RequireRectNear(candidateForm.value().rcArea, expectedCaretClientPx, context);
            return compositionForm.value().ptCurrentPos;
        };

        field->SetSelectionRange(laterIndex.value(), laterIndex.value());
        window.Host().SyncTextInput(field);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

        const POINT laterPoint = requireImeFormsAtCaret(laterIndex.value(),
                                                        wrapped ? "native wrapped ime composition/candidate forms track a later wrapped visual-line caret"
                                                                : "native multiline ime composition/candidate forms track a later logical-line caret");
        Require(laterPoint.y > static_cast<LONG>(firstCaretDip.top),
                wrapped ? "native wrapped ime anchor starts on a later visual line" : "native multiline ime anchor starts on a later logical line");

        field->SetSelectionRange(firstIndex, firstIndex);
        window.Host().SyncTextInput(field);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, 0));

        const POINT firstPoint =
            requireImeFormsAtCaret(firstIndex,
                                   wrapped ? "native wrapped ime composition/candidate forms update after caret moves during composition"
                                           : "native multiline ime composition/candidate forms update after caret moves during composition");
        Require(firstPoint.y < laterPoint.y,
                wrapped ? "native wrapped ime anchor returns to the first visual line during composition"
                        : "native multiline ime anchor returns to the first logical line during composition");
    };

    runCase(false);
    runCase(true);
}

void TestNativeTextInputBackendImeWindowsUpdateWhenFocusedFieldMoves()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(18.0f, 14.0f, 238.0f, 42.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(4u, 4u);
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    const auto beforeCompositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
    Require(beforeCompositionForm.has_value(), "native ime composition form is readable before the focused field moves");

    field->SetBounds(D2D1::RectF(58.0f, 28.0f, 278.0f, 56.0f));

    D2D1_RECT_F caretRectDip{};
    RECT caretRectClientPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(caretRectDip, caretRectClientPx),
            "native ime moved-field test can read the moved native caret rect");
    MapWindowPoints(nullptr, window.Hwnd(), reinterpret_cast<POINT*>(&caretRectClientPx), 2);

    const auto afterCompositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
    Require(afterCompositionForm.has_value(), "native ime composition form is readable after the focused field moves");
    RequirePointNear(afterCompositionForm.value().ptCurrentPos,
                     POINT{caretRectClientPx.left, caretRectClientPx.top},
                     "native ime composition point reanchors after the focused field moves");
    Require(afterCompositionForm.value().ptCurrentPos.x > beforeCompositionForm.value().ptCurrentPos.x,
            "native ime composition point moves right with the focused field");

    const auto afterCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(afterCandidateForm.has_value(), "native ime candidate form is readable after the focused field moves");
    RequirePointNear(afterCandidateForm.value().ptCurrentPos,
                     POINT{caretRectClientPx.left, caretRectClientPx.bottom},
                     "native ime candidate point reanchors after the focused field moves");
    RequireRectNear(afterCandidateForm.value().rcArea, caretRectClientPx, "native ime candidate exclusion rect reanchors after the focused field moves");
}

void TestNativeTextInputBackendMultilineImeWindowsUpdateWhenFocusedFieldMoves()
{
    using namespace RedSalamander::DxUi;

    const auto runCase = [](bool wrapped)
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(wrapped ? std::wstring(kWrappedMultilineClipboardTextForTest) : std::wstring(L"alpha\nbeta\ngamma"));
        field->SetMultiline(true);
        field->SetBounds(wrapped ? D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f) : D2D1::RectF(16.0f, 18.0f, 276.0f, 162.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        std::optional<size_t> laterIndex;
        if (wrapped)
        {
            laterIndex = field->GetText().find(L"foxtrot");
            Require(laterIndex.has_value() && laterIndex.value() != std::wstring::npos, "native wrapped ime move test can locate a later wrapped token");
        }
        else
        {
            laterIndex = field->GetText().find(L"gamma");
            Require(laterIndex.has_value() && laterIndex.value() != std::wstring::npos, "native multiline ime move test can locate a later-line token");
        }

        field->SetSelectionRange(laterIndex.value(), laterIndex.value());
        window.Host().SyncTextInput(field);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

        const auto beforeCompositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
        Require(beforeCompositionForm.has_value(),
                wrapped ? "native wrapped ime composition form is readable before the field moves"
                        : "native multiline ime composition form is readable before the field moves");

        field->SetBounds(wrapped ? D2D1::RectF(52.0f, 48.0f, 172.0f, 144.0f) : D2D1::RectF(52.0f, 48.0f, 312.0f, 192.0f));
        window.Host().SyncTextInput(field);

        const std::optional<D2D1_RECT_F> expectedCaretDip = field->TryGetTextInputCaretRect(window.Host(), laterIndex.value());
        Require(expectedCaretDip.has_value(),
                wrapped ? "native wrapped ime move test resolves the moved caret geometry"
                        : "native multiline ime move test resolves the moved caret geometry");
        const RECT expectedCaretClientPx = DipRectToClientRect(window, expectedCaretDip.value());

        const auto afterCompositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
        Require(afterCompositionForm.has_value(),
                wrapped ? "native wrapped ime composition form remains readable after the field moves"
                        : "native multiline ime composition form remains readable after the field moves");
        RequirePointNear(afterCompositionForm.value().ptCurrentPos,
                         POINT{expectedCaretClientPx.left, expectedCaretClientPx.top},
                         wrapped ? "native wrapped ime composition point reanchors after the field moves"
                                 : "native multiline ime composition point reanchors after the field moves");
        Require(afterCompositionForm.value().ptCurrentPos.x > beforeCompositionForm.value().ptCurrentPos.x,
                wrapped ? "native wrapped ime composition point moves right with the focused field"
                        : "native multiline ime composition point moves right with the focused field");
        Require(afterCompositionForm.value().ptCurrentPos.y > beforeCompositionForm.value().ptCurrentPos.y,
                wrapped ? "native wrapped ime composition point moves down with the focused field"
                        : "native multiline ime composition point moves down with the focused field");

        const auto afterCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
        Require(afterCandidateForm.has_value(),
                wrapped ? "native wrapped ime candidate form remains readable after the field moves"
                        : "native multiline ime candidate form remains readable after the field moves");
        RequirePointNear(afterCandidateForm.value().ptCurrentPos,
                         POINT{expectedCaretClientPx.left, expectedCaretClientPx.bottom},
                         wrapped ? "native wrapped ime candidate point reanchors after the field moves"
                                 : "native multiline ime candidate point reanchors after the field moves");
        RequireRectNear(afterCandidateForm.value().rcArea,
                        expectedCaretClientPx,
                        wrapped ? "native wrapped ime candidate exclusion rect reanchors after the field moves"
                                : "native multiline ime candidate exclusion rect reanchors after the field moves");
    };

    runCase(false);
    runCase(true);
}

void TestNativeTextInputBackendImeWindowsUpdateWhenEditableComboMoves()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetBounds(D2D1::RectF(18.0f, 14.0f, 238.0f, 46.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});
    combo->SetText(L"alpha beta");

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(combo);
    combo->SetEditableSelectionRange(4u, 4u);
    window.Host().SyncTextInput(combo);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    const auto beforeCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(beforeCandidateForm.has_value(), "native editable combo ime candidate form is readable before the combo moves");

    combo->SetBounds(D2D1::RectF(68.0f, 32.0f, 288.0f, 64.0f));

    D2D1_RECT_F caretRectDip{};
    RECT caretRectClientPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(caretRectDip, caretRectClientPx),
            "native editable combo ime move test can read the moved native caret rect");
    MapWindowPoints(nullptr, window.Hwnd(), reinterpret_cast<POINT*>(&caretRectClientPx), 2);

    const auto afterCompositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
    Require(afterCompositionForm.has_value(), "native editable combo ime composition form is readable after the combo moves");
    RequirePointNear(afterCompositionForm.value().ptCurrentPos,
                     POINT{caretRectClientPx.left, caretRectClientPx.top},
                     "native editable combo ime composition point reanchors after the combo moves");

    const auto afterCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(afterCandidateForm.has_value(), "native editable combo ime candidate form is readable after the combo moves");
    RequirePointNear(afterCandidateForm.value().ptCurrentPos,
                     POINT{caretRectClientPx.left, caretRectClientPx.bottom},
                     "native editable combo ime candidate point reanchors after the combo moves");
    RequireRectNear(afterCandidateForm.value().rcArea, caretRectClientPx, "native editable combo ime candidate exclusion rect reanchors after the combo moves");
    Require(afterCandidateForm.value().ptCurrentPos.x > beforeCandidateForm.value().ptCurrentPos.x,
            "native editable combo ime candidate point moves right with the focused combo");
}

void TestNativeTextInputBackendImeWindowsUpdateAfterProgrammaticTextFieldCaretMove()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta gamma");
    field->SetBounds(D2D1::RectF(18.0f, 14.0f, 320.0f, 42.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(0u, 0u);
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    const auto beforeCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(beforeCandidateForm.has_value(), "native ime candidate form is readable before programmatic TextField caret movement");

    const size_t targetCaret = field->GetText().size();
    field->SetSelectionRange(targetCaret, targetCaret);

    NativeTextInputState state{};
    Require(window.Host().TryReadNativeTextInputState(field, state), "native ime programmatic TextField caret test reads native state");
    Require(state.caretIndex == targetCaret, "native ime programmatic TextField caret movement refreshes native caret state");

    const std::optional<D2D1_RECT_F> expectedCaretDip = field->TryGetTextInputCaretRect(window.Host(), targetCaret);
    Require(expectedCaretDip.has_value(), "native ime programmatic TextField caret test resolves expected caret geometry");
    const RECT expectedCaretClientPx = DipRectToClientRect(window, expectedCaretDip.value());

    const auto afterCompositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
    Require(afterCompositionForm.has_value(), "native ime composition form is readable after programmatic TextField caret movement");
    RequirePointNear(afterCompositionForm.value().ptCurrentPos,
                     POINT{expectedCaretClientPx.left, expectedCaretClientPx.top},
                     "native ime composition point reanchors after programmatic TextField caret movement");

    const auto afterCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(afterCandidateForm.has_value(), "native ime candidate form is readable after programmatic TextField caret movement");
    RequirePointNear(afterCandidateForm.value().ptCurrentPos,
                     POINT{expectedCaretClientPx.left, expectedCaretClientPx.bottom},
                     "native ime candidate point reanchors after programmatic TextField caret movement");
    RequireRectNear(
        afterCandidateForm.value().rcArea, expectedCaretClientPx, "native ime candidate exclusion rect reanchors after programmatic TextField caret movement");
    Require(afterCandidateForm.value().ptCurrentPos.x > beforeCandidateForm.value().ptCurrentPos.x,
            "native ime candidate point moves right after programmatic TextField caret movement");
}

void TestNativeTextInputBackendImeWindowsUpdateAfterProgrammaticEditableComboCaretMove()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"alpha beta gamma");
    combo->SetBounds(D2D1::RectF(18.0f, 14.0f, 320.0f, 46.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(combo);
    combo->SetEditableSelectionRange(0u, 0u);
    window.Host().SyncTextInput(combo);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    const auto beforeCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(beforeCandidateForm.has_value(), "native editable combo ime candidate form is readable before programmatic caret movement");

    const size_t targetCaret = combo->GetText().size();
    combo->SetEditableSelectionRange(targetCaret, targetCaret);

    NativeTextInputState state{};
    Require(window.Host().TryReadNativeTextInputState(combo, state), "native editable combo ime programmatic caret test reads native state");
    Require(state.caretIndex == targetCaret, "native editable combo ime programmatic caret movement refreshes native caret state");

    const std::optional<D2D1_RECT_F> expectedCaretDip = combo->TryGetTextInputCaretRect(window.Host(), targetCaret);
    Require(expectedCaretDip.has_value(), "native editable combo ime programmatic caret test resolves expected caret geometry");
    const RECT expectedCaretClientPx = DipRectToClientRect(window, expectedCaretDip.value());

    const auto afterCompositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
    Require(afterCompositionForm.has_value(), "native editable combo ime composition form is readable after programmatic caret movement");
    RequirePointNear(afterCompositionForm.value().ptCurrentPos,
                     POINT{expectedCaretClientPx.left, expectedCaretClientPx.top},
                     "native editable combo ime composition point reanchors after programmatic caret movement");

    const auto afterCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(afterCandidateForm.has_value(), "native editable combo ime candidate form is readable after programmatic caret movement");
    RequirePointNear(afterCandidateForm.value().ptCurrentPos,
                     POINT{expectedCaretClientPx.left, expectedCaretClientPx.bottom},
                     "native editable combo ime candidate point reanchors after programmatic caret movement");
    RequireRectNear(afterCandidateForm.value().rcArea,
                    expectedCaretClientPx,
                    "native editable combo ime candidate exclusion rect reanchors after programmatic caret movement");
    Require(afterCandidateForm.value().ptCurrentPos.x > beforeCandidateForm.value().ptCurrentPos.x,
            "native editable combo ime candidate point moves right after programmatic caret movement");
}

void TestNativeTextInputBackendImeWindowsUpdateAfterFocusedTextFieldPaddingChange()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta gamma");
    field->SetBounds(D2D1::RectF(18.0f, 14.0f, 320.0f, 42.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(5u, 5u);
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    const auto beforeCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(beforeCandidateForm.has_value(), "native ime candidate form is readable before focused TextField padding changes");

    field->SetHorizontalTextPadding(26.0f, 8.0f);
    field->SetVerticalTextPadding(12.0f, 4.0f);

    const std::optional<D2D1_RECT_F> expectedCaretDip = field->TryGetTextInputCaretRect(window.Host(), 5u);
    Require(expectedCaretDip.has_value(), "native ime TextField padding test resolves expected caret geometry");
    const RECT expectedCaretClientPx = DipRectToClientRect(window, expectedCaretDip.value());

    const auto afterCompositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
    Require(afterCompositionForm.has_value(), "native ime composition form is readable after focused TextField padding changes");
    RequirePointNear(afterCompositionForm.value().ptCurrentPos,
                     POINT{expectedCaretClientPx.left, expectedCaretClientPx.top},
                     "native ime composition point reanchors after focused TextField padding changes");

    const auto afterCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(afterCandidateForm.has_value(), "native ime candidate form is readable after focused TextField padding changes");
    RequirePointNear(afterCandidateForm.value().ptCurrentPos,
                     POINT{expectedCaretClientPx.left, expectedCaretClientPx.bottom},
                     "native ime candidate point reanchors after focused TextField padding changes");
    RequireRectNear(
        afterCandidateForm.value().rcArea, expectedCaretClientPx, "native ime candidate exclusion rect reanchors after focused TextField padding changes");
    Require(afterCandidateForm.value().ptCurrentPos.x > beforeCandidateForm.value().ptCurrentPos.x,
            "native ime candidate point moves right after focused TextField padding changes");
}

void TestNativeTextInputBackendImeWindowsUpdateAfterFocusedEditableComboDensityChange()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"alpha beta gamma delta epsilon zeta");
    combo->SetBounds(D2D1::RectF(18.0f, 14.0f, 150.0f, 46.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(combo);
    combo->SetEditableSelectionRange(combo->GetText().size(), combo->GetText().size());
    window.Host().SyncTextInput(combo);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    const auto beforeCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(beforeCandidateForm.has_value(), "native editable combo ime candidate form is readable before density changes");

    ThemePalette compactTheme = window.Host().GetTheme();
    compactTheme.density      = Density::Compact;
    window.Host().SetTheme(compactTheme);

    const std::optional<D2D1_RECT_F> expectedCaretDip = combo->TryGetTextInputCaretRect(window.Host(), combo->GetText().size());
    Require(expectedCaretDip.has_value(), "native ime editable combo density test resolves expected caret geometry");
    const RECT expectedCaretClientPx = DipRectToClientRect(window, expectedCaretDip.value());

    const auto afterCompositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
    Require(afterCompositionForm.has_value(), "native editable combo ime composition form is readable after density changes");
    RequirePointNear(afterCompositionForm.value().ptCurrentPos,
                     POINT{expectedCaretClientPx.left, expectedCaretClientPx.top},
                     "native editable combo ime composition point reanchors after density changes");

    const auto afterCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(afterCandidateForm.has_value(), "native editable combo ime candidate form is readable after density changes");
    RequirePointNear(afterCandidateForm.value().ptCurrentPos,
                     POINT{expectedCaretClientPx.left, expectedCaretClientPx.bottom},
                     "native editable combo ime candidate point reanchors after density changes");
    RequireRectNear(
        afterCandidateForm.value().rcArea, expectedCaretClientPx, "native editable combo ime candidate exclusion rect reanchors after density changes");
    Require(afterCandidateForm.value().rcArea.right > beforeCandidateForm.value().rcArea.right,
            "native editable combo ime candidate exclusion rect widens with compact density text viewport");
}

void TestNativeTextInputBackendImeWindowsUpdateAfterDpiChange()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(18.0f, 14.0f, 238.0f, 42.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(4u, 4u);
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    D2D1_RECT_F beforeCaretRectDip{};
    RECT beforeCaretClientPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(beforeCaretRectDip, beforeCaretClientPx),
            "native ime dpi test can read the native caret before dpi change");
    MapWindowPoints(nullptr, window.Hwnd(), reinterpret_cast<POINT*>(&beforeCaretClientPx), 2);

    RECT windowRect{};
    Require(GetWindowRect(window.Hwnd(), &windowRect) != FALSE, "native ime dpi test can read the attached window rect");
    const LONG windowWidthPx  = std::max<LONG>(1, windowRect.right - windowRect.left);
    const LONG windowHeightPx = std::max<LONG>(1, windowRect.bottom - windowRect.top);
    RECT suggestedRect{windowRect.left,
                       windowRect.top,
                       windowRect.left + MulDiv(windowWidthPx, 144, USER_DEFAULT_SCREEN_DPI),
                       windowRect.top + MulDiv(windowHeightPx, 144, USER_DEFAULT_SCREEN_DPI)};

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_DPICHANGED, MAKELONG(144u, 144u), reinterpret_cast<LPARAM>(&suggestedRect), handled));
    Require(handled, "native ime dpi test handles WM_DPICHANGED");

    D2D1_RECT_F afterCaretRectDip{};
    RECT afterCaretClientPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(afterCaretRectDip, afterCaretClientPx),
            "native ime dpi test can read the native caret after dpi change");
    MapWindowPoints(nullptr, window.Hwnd(), reinterpret_cast<POINT*>(&afterCaretClientPx), 2);
    Require(afterCaretClientPx.left > beforeCaretClientPx.left, "native ime caret client anchor scales after dpi change");

    const auto afterCompositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
    Require(afterCompositionForm.has_value(), "native ime composition form is readable after dpi change");
    RequirePointNear(afterCompositionForm.value().ptCurrentPos,
                     POINT{afterCaretClientPx.left, afterCaretClientPx.top},
                     "native ime composition point reanchors after dpi change");

    const auto afterCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(afterCandidateForm.has_value(), "native ime candidate form is readable after dpi change");
    RequirePointNear(afterCandidateForm.value().ptCurrentPos,
                     POINT{afterCaretClientPx.left, afterCaretClientPx.bottom},
                     "native ime candidate point reanchors after dpi change");
    RequireRectNear(afterCandidateForm.value().rcArea, afterCaretClientPx, "native ime candidate exclusion rect reanchors after dpi change");
}

void TestNativeTextInputBackendImeWindowsUpdateAfterMultilineScroll()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"line 1\nline 2\nline 3\nline 4\nline 5\nline 6");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(18.0f, 14.0f, 238.0f, 74.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(field->GetText().size(), field->GetText().size());
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    const auto beforeCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(beforeCandidateForm.has_value(), "native ime candidate form is readable before multiline scroll");

    Require(field->OnMouseWheel(window.Host(), D2D1::Point2F(32.0f, 32.0f), -static_cast<float>(WHEEL_DELTA), 0u),
            "native multiline text field handles mouse-wheel scrolling during IME composition");

    D2D1_RECT_F afterCaretRectDip{};
    RECT afterCaretClientPx{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(afterCaretRectDip, afterCaretClientPx),
            "native ime multiline scroll test can read the scrolled native caret");
    MapWindowPoints(nullptr, window.Hwnd(), reinterpret_cast<POINT*>(&afterCaretClientPx), 2);

    const auto afterCompositionForm = ReadTextInputCompositionFormForTest(window.Hwnd());
    Require(afterCompositionForm.has_value(), "native ime composition form is readable after multiline scroll");
    RequirePointNear(afterCompositionForm.value().ptCurrentPos,
                     POINT{afterCaretClientPx.left, afterCaretClientPx.top},
                     "native ime composition point reanchors after multiline scroll");

    const auto afterCandidateForm = ReadTextInputCandidateFormForTest(window.Hwnd(), 0u);
    Require(afterCandidateForm.has_value(), "native ime candidate form is readable after multiline scroll");
    RequirePointNear(afterCandidateForm.value().ptCurrentPos,
                     POINT{afterCaretClientPx.left, afterCaretClientPx.bottom},
                     "native ime candidate point reanchors after multiline scroll");
    RequireRectNear(afterCandidateForm.value().rcArea, afterCaretClientPx, "native ime candidate exclusion rect reanchors after multiline scroll");
    Require(afterCandidateForm.value().ptCurrentPos.y != beforeCandidateForm.value().ptCurrentPos.y,
            "native ime candidate point changes when multiline scrolling shifts the caret viewport");
}

void TestNativeTextInputBackendImeResultPayloadCommitsSelectionReplacement()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 260.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(0u, 5u);
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputImePayload previewPayload;
    previewPayload.hasCompositionString = true;
    previewPayload.compositionString    = L"draft";
    window.Host().DebugSetNativeTextInputImePayloadForTest(previewPayload);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_COMPSTR));

    Require(field->GetText() == L"draft beta", "native ime preview replaces the active selected composition range");

    NativeTextInputImePayload payload;
    payload.hasResultString = true;
    payload.resultString    = L"kana";
    window.Host().DebugSetNativeTextInputImePayloadForTest(payload);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_RESULTSTR));

    NativeTextInputState state;
    Require(field->GetText() == L"kana beta", "native ime result payload replaces the active selected composition range");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native ime result payload leaves readable native session state");
    Require(state.text == field->GetText(), "native ime result payload syncs retained text into native state");
    Require(state.caretIndex == 4u, "native ime result payload places the caret after the committed result");
    Require(! state.selectionAnchorIndex.has_value(), "native ime result payload clears the replaced selection");
    Require(! state.compositionStartIndex.has_value() && ! state.compositionEndIndex.has_value(),
            "native ime result payload clears the active composition range");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_ENDCOMPOSITION, 0, 0));

    Require(field->GetText() == L"kana beta", "native ime end after result commit keeps the committed text");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native ime end after result commit leaves readable native session state");
    Require(state.text == field->GetText(), "native ime end after result commit keeps native state synchronized with retained text");
}

void TestNativeTextInputBackendImeCompositionPayloadPreviewsAndCancelRestoresBase()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 260.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(5u, 5u);
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputImePayload payload;
    payload.hasCompositionString  = true;
    payload.compositionString     = L"-ime";
    payload.compositionAttributes = {ATTR_INPUT, ATTR_TARGET_CONVERTED, ATTR_TARGET_CONVERTED, ATTR_INPUT};
    payload.compositionClauses    = {0u, 1u, 4u};
    payload.hasCursorPosition     = true;
    payload.cursorPosition        = 3u;
    window.Host().DebugSetNativeTextInputImePayloadForTest(payload);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_COMPSTR | GCS_COMPATTR | GCS_COMPCLAUSE | GCS_CURSORPOS));

    NativeTextInputState state;
    Require(field->GetText() == L"alpha-ime beta", "native ime composition payload previews inline in retained text");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native ime composition payload leaves readable native session state");
    Require(state.text == field->GetText(), "native ime composition payload syncs preview text into native state");
    Require(state.compositionStartIndex.has_value() && state.compositionStartIndex.value() == 5u,
            "native ime composition payload starts at the composition anchor");
    Require(state.compositionEndIndex.has_value() && state.compositionEndIndex.value() == 9u,
            "native ime composition payload extends by the composition string length");
    Require(state.conversionTargetStartIndex.has_value() && state.conversionTargetStartIndex.value() == 6u,
            "native ime composition attributes mark the first target-converted character");
    Require(state.conversionTargetEndIndex.has_value() && state.conversionTargetEndIndex.value() == 8u,
            "native ime composition attributes mark the target-converted range end");
    Require(state.compositionCursorIndex.has_value() && state.compositionCursorIndex.value() == 8u,
            "native ime composition exposes an absolute debug cursor index");
    const std::vector<size_t> expectedClauseBoundaries{5u, 6u, 9u};
    Require(state.compositionClauseBoundaries == expectedClauseBoundaries, "native ime composition maps clause offsets to absolute debug boundaries");
    Require(state.caretIndex == 8u, "native ime composition cursor position maps inside the preview string");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_ENDCOMPOSITION, 0, 0));

    Require(field->GetText() == L"alpha beta", "native ime cancelled composition restores the pre-composition base text");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native ime cancelled composition leaves readable native session state");
    Require(! state.compositionStartIndex.has_value() && ! state.compositionEndIndex.has_value(),
            "native ime cancelled composition clears the composition range");
    Require(! state.compositionCursorIndex.has_value() && state.compositionClauseBoundaries.empty(),
            "native ime cancelled composition clears debug cursor and clause diagnostics");
    Require(state.caretIndex == 5u, "native ime cancelled composition restores the base caret");
}

void TestNativeTextInputBackendImeCompositionPaintExposesStyledInlineRanges()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 260.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(5u, 5u);
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputImePayload payload;
    payload.hasCompositionString  = true;
    payload.compositionString     = L"-ime";
    payload.compositionAttributes = {ATTR_INPUT, ATTR_TARGET_CONVERTED, ATTR_TARGET_CONVERTED, ATTR_INPUT};
    window.Host().DebugSetNativeTextInputImePayloadForTest(payload);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_COMPSTR | GCS_COMPATTR));

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native ime styled paint test reads native session state");
    Require(state.compositionStartIndex.has_value() && state.compositionStartIndex.value() == 5u, "native ime styled paint test has a composition start");
    Require(state.compositionEndIndex.has_value() && state.compositionEndIndex.value() == 9u, "native ime styled paint test has a composition end");
    Require(state.conversionTargetStartIndex.has_value() && state.conversionTargetStartIndex.value() == 6u,
            "native ime styled paint test has a conversion target start");
    Require(state.conversionTargetEndIndex.has_value() && state.conversionTargetEndIndex.value() == 8u,
            "native ime styled paint test has a conversion target end");

    const std::optional<std::vector<D2D1_RECT_F>> expectedCompositionRects = field->TryGetTextInputRangeRects(window.Host(), 5u, 9u);
    const std::optional<std::vector<D2D1_RECT_F>> expectedConversionRects  = field->TryGetTextInputRangeRects(window.Host(), 6u, 8u);
    Require(expectedCompositionRects.has_value() && ! expectedCompositionRects->empty(), "native ime styled paint test resolves composition range geometry");
    Require(expectedConversionRects.has_value() && ! expectedConversionRects->empty(),
            "native ime styled paint test resolves conversion target range geometry");

    TextFieldDebugSingleLinePaintState paint{};
    Require(field->DebugGetSingleLinePaintState(window.Host(), paint), "native ime styled paint test reads single-line paint state");
    Require(paint.compositionUnderlineRects.size() == expectedCompositionRects->size(),
            "native ime styled paint exposes one underline rect per composition range rect");
    Require(paint.conversionTargetUnderlineRects.size() == expectedConversionRects->size(),
            "native ime styled paint exposes one underline rect per conversion-target range rect");

    const D2D1_RECT_F& compositionUnderline = paint.compositionUnderlineRects.front();
    const D2D1_RECT_F& compositionRect      = expectedCompositionRects->front();
    RequireRectHasArea(compositionUnderline, "native ime composition underline has area");
    RequireFloatNear(compositionUnderline.left, compositionRect.left, 1.0f, "native ime composition underline follows range left");
    RequireFloatNear(compositionUnderline.right, compositionRect.right, 1.0f, "native ime composition underline follows range right");
    Require(compositionUnderline.top >= compositionRect.top && compositionUnderline.bottom <= compositionRect.bottom,
            "native ime composition underline stays inside the text range vertical bounds");

    const D2D1_RECT_F& targetUnderline = paint.conversionTargetUnderlineRects.front();
    const D2D1_RECT_F& targetRect      = expectedConversionRects->front();
    RequireRectHasArea(targetUnderline, "native ime conversion-target underline has area");
    RequireFloatNear(targetUnderline.left, targetRect.left, 1.0f, "native ime conversion-target underline follows range left");
    RequireFloatNear(targetUnderline.right, targetRect.right, 1.0f, "native ime conversion-target underline follows range right");
    Require(targetUnderline.top >= targetRect.top && targetUnderline.bottom <= targetRect.bottom,
            "native ime conversion-target underline stays inside the text range vertical bounds");
    Require(targetUnderline.left > compositionUnderline.left && targetUnderline.right < compositionUnderline.right,
            "native ime conversion-target underline is nested inside the broader composition underline");
}

void TestNativeTextInputBackendImeCompositionPaintExposesEditableComboInlineRanges()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"alpha beta");
    combo->SetBounds(D2D1::RectF(12.0f, 16.0f, 260.0f, 48.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(combo);
    combo->SetEditableSelectionRange(5u, 5u);
    window.Host().SyncTextInput(combo);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputImePayload payload;
    payload.hasCompositionString  = true;
    payload.compositionString     = L"-ime";
    payload.compositionAttributes = {ATTR_INPUT, ATTR_TARGET_CONVERTED, ATTR_TARGET_CONVERTED, ATTR_INPUT};
    window.Host().DebugSetNativeTextInputImePayloadForTest(payload);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_COMPSTR | GCS_COMPATTR));

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(combo, state), "native editable combo ime styled paint test reads native session state");
    Require(state.compositionStartIndex.has_value() && state.compositionStartIndex.value() == 5u,
            "native editable combo ime styled paint test has a composition start");
    Require(state.compositionEndIndex.has_value() && state.compositionEndIndex.value() == 9u,
            "native editable combo ime styled paint test has a composition end");
    Require(state.conversionTargetStartIndex.has_value() && state.conversionTargetStartIndex.value() == 6u,
            "native editable combo ime styled paint test has a conversion target start");
    Require(state.conversionTargetEndIndex.has_value() && state.conversionTargetEndIndex.value() == 8u,
            "native editable combo ime styled paint test has a conversion target end");

    const std::optional<std::vector<D2D1_RECT_F>> expectedCompositionRects = combo->TryGetTextInputRangeRects(window.Host(), 5u, 9u);
    const std::optional<std::vector<D2D1_RECT_F>> expectedConversionRects  = combo->TryGetTextInputRangeRects(window.Host(), 6u, 8u);
    Require(expectedCompositionRects.has_value() && ! expectedCompositionRects->empty(),
            "native editable combo ime styled paint test resolves composition range geometry");
    Require(expectedConversionRects.has_value() && ! expectedConversionRects->empty(),
            "native editable combo ime styled paint test resolves conversion target range geometry");

    ComboBoxDebugEditablePaintState paint{};
    Require(combo->DebugGetEditablePaintState(window.Host(), paint), "native editable combo ime styled paint test reads editable paint state");
    Require(paint.compositionUnderlineRects.size() == expectedCompositionRects->size(),
            "native editable combo ime styled paint exposes one underline rect per composition range rect");
    Require(paint.conversionTargetUnderlineRects.size() == expectedConversionRects->size(),
            "native editable combo ime styled paint exposes one underline rect per conversion-target range rect");

    const D2D1_RECT_F& compositionUnderline = paint.compositionUnderlineRects.front();
    const D2D1_RECT_F& compositionRect      = expectedCompositionRects->front();
    RequireRectHasArea(compositionUnderline, "native editable combo ime composition underline has area");
    RequireFloatNear(compositionUnderline.left, compositionRect.left, 1.0f, "native editable combo ime composition underline follows range left");
    RequireFloatNear(compositionUnderline.right, compositionRect.right, 1.0f, "native editable combo ime composition underline follows range right");
    Require(compositionUnderline.top >= compositionRect.top && compositionUnderline.bottom <= compositionRect.bottom,
            "native editable combo ime composition underline stays inside the text range vertical bounds");

    const D2D1_RECT_F& targetUnderline = paint.conversionTargetUnderlineRects.front();
    const D2D1_RECT_F& targetRect      = expectedConversionRects->front();
    RequireRectHasArea(targetUnderline, "native editable combo ime conversion-target underline has area");
    RequireFloatNear(targetUnderline.left, targetRect.left, 1.0f, "native editable combo ime conversion-target underline follows range left");
    RequireFloatNear(targetUnderline.right, targetRect.right, 1.0f, "native editable combo ime conversion-target underline follows range right");
    Require(targetUnderline.top >= targetRect.top && targetUnderline.bottom <= targetRect.bottom,
            "native editable combo ime conversion-target underline stays inside the text range vertical bounds");
    Require(targetUnderline.left > compositionUnderline.left && targetUnderline.right < compositionUnderline.right,
            "native editable combo ime conversion-target underline is nested inside the broader composition underline");
}

void TestNativeTextInputBackendImeMultilineWrappedPreviewThenResultCommitsAtOriginalAnchor()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"first line\nalpha beta gamma wraps");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 124.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(11u, 11u);
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputImePayload previewPayload;
    previewPayload.hasCompositionString = true;
    previewPayload.compositionString    = L"kana-preview";
    previewPayload.hasCursorPosition    = true;
    previewPayload.cursorPosition       = previewPayload.compositionString.size();
    window.Host().DebugSetNativeTextInputImePayloadForTest(previewPayload);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_COMPSTR | GCS_CURSORPOS));

    NativeTextInputState state;
    Require(field->GetText() == L"first line\nkana-previewalpha beta gamma wraps", "native multiline ime preview inserts at the retained composition anchor");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native multiline ime preview leaves readable native state");
    Require(state.multiline, "native multiline ime preview keeps multiline session state");
    Require(state.compositionStartIndex.has_value() && state.compositionStartIndex.value() == 11u,
            "native multiline ime preview starts at the original anchor");
    Require(state.compositionEndIndex.has_value() && state.compositionEndIndex.value() == 23u, "native multiline ime preview exposes the active preview span");
    Require(state.caretIndex == 23u, "native multiline ime preview moves the caret to the preview cursor");

    NativeTextInputImePayload resultPayload;
    resultPayload.hasResultString = true;
    resultPayload.resultString    = L"done";
    window.Host().DebugSetNativeTextInputImePayloadForTest(resultPayload);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_RESULTSTR));

    Require(field->GetText() == L"first line\ndonealpha beta gamma wraps",
            "native multiline ime result replaces the original anchor, not a preview-length span in the base text");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native multiline ime result leaves readable native state");
    Require(state.text == field->GetText(), "native multiline ime result syncs retained text into native state");
    Require(state.caretIndex == 15u, "native multiline ime result places the caret after the committed result");
    Require(! state.selectionAnchorIndex.has_value(), "native multiline ime result leaves a collapsed selection");
    Require(! state.compositionStartIndex.has_value() && ! state.compositionEndIndex.has_value(),
            "native multiline ime result clears the active composition span");
}

void TestNativeTextInputBackendImeCompositionOwnsSpecialKeys()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(L"alpha");
    auto* nextButton   = root->AddChild<Button>(L"Next");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 40.0f, 120.0f, 68.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 40.0f, 260.0f, 68.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state) && state.compositionStartIndex.has_value(),
            "native ime composition activates on the host hwnd");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "native ime composition keeps return from invoking the host default button");
    Require(window.Host().GetFocusControl() == field, "native ime composition keeps focus on the text field after return");
    Require(window.Host().DebugHasActiveNativeTextInputSession(), "native ime composition keeps the native session active after return");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 0u, "native ime composition keeps escape from invoking the host cancel button");
    Require(window.Host().GetFocusControl() == field, "native ime composition keeps focus on the text field after escape");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == field, "native ime composition keeps tab from advancing host focus");
    Require(window.Host().DebugHasActiveNativeTextInputSession(), "native ime composition keeps the native session active after tab");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_ENDCOMPOSITION, 0, 0));

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 1u, "return resumes host default-button routing after native ime composition ends");

    window.Host().SetFocusControl(field);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 1u, "escape resumes host cancel-button routing after native ime composition ends");

    window.Host().SetFocusControl(field);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == nextButton, "tab resumes host focus traversal after native ime composition ends");
    Require(! window.Host().DebugHasActiveNativeTextInputSession(), "native session deactivates once tab resumes host focus traversal");
}

void TestNativeTextInputBackendImeCompositionLetsModifiedNavigationKeysRoute()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta gamma");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(field->GetText().size(), field->GetText().size());
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state) && state.compositionStartIndex.has_value(),
            "native ime modified-key routing test starts with active composition");
    Require(state.caretIndex == field->GetText().size(), "native ime modified-key routing starts at the end");

    SendNativeCtrlKey(window.Hwnd(), VK_LEFT);

    Require(window.Host().TryReadNativeTextInputState(field, state), "native ime modified ctrl+left leaves readable state");
    Require(state.caretIndex < field->GetText().size(), "native ime modified ctrl+left is routed to retained text navigation");
    Require(state.caretIndex == 11u, "native ime modified ctrl+left moves to the previous word boundary");
    Require(state.compositionStartIndex.has_value(), "native ime modified ctrl+left keeps the active composition state");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_LEFT, 0));
    Require(window.Host().TryReadNativeTextInputState(field, state), "native ime plain left leaves readable state");
    Require(state.caretIndex == 11u, "native ime plain left remains composition-owned and does not move the retained caret");
}

void TestNativeTextInputBackendMultilineImeCompositionOwnsSpecialKeys()
{
    using namespace RedSalamander::DxUi;

    const auto runScenario = [](bool wrapped)
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root          = std::make_unique<Panel>();
        auto* field        = root->AddChild<TextField>(wrapped ? std::wstring(kWrappedMultilineClipboardTextForTest) : std::wstring(L"alpha\nbeta"));
        auto* nextButton   = root->AddChild<Button>(L"Next");
        auto* cancelButton = root->AddChild<Button>(L"Cancel");
        field->SetMultiline(true);
        field->SetBounds(wrapped ? D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f) : D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
        nextButton->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));
        cancelButton->SetBounds(D2D1::RectF(140.0f, 112.0f, 260.0f, 140.0f));
        const std::wstring originalText(field->GetText());

        size_t defaultCount = 0u;
        size_t cancelCount  = 0u;
        nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
        cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

        window.Host().SetRoot(std::move(root));
        window.Host().SetDefaultButton(nextButton);
        window.Host().SetCancelButton(cancelButton);
        window.Host().SetFocusControl(field);

        static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state) && state.compositionStartIndex.has_value(),
                wrapped ? "native wrapped multiline ime composition activates on the host hwnd"
                        : "native multiline ime composition activates on the host hwnd");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_RETURN, 0));
        Require(defaultCount == 0u,
                wrapped ? "native wrapped multiline ime keeps return from invoking the default button"
                        : "native multiline ime keeps return from invoking the default button");
        Require(field->GetText() == originalText,
                wrapped ? "native wrapped multiline ime return does not insert a newline during active composition"
                        : "native multiline ime return does not insert a newline during active composition");
        Require(window.Host().GetFocusControl() == field,
                wrapped ? "native wrapped multiline ime keeps focus after return" : "native multiline ime keeps focus after return");
        Require(window.Host().DebugHasActiveNativeTextInputSession(),
                wrapped ? "native wrapped multiline ime keeps session active after return" : "native multiline ime keeps session active after return");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_ESCAPE, 0));
        Require(cancelCount == 0u,
                wrapped ? "native wrapped multiline ime keeps escape from invoking cancel" : "native multiline ime keeps escape from invoking cancel");
        Require(window.Host().GetFocusControl() == field,
                wrapped ? "native wrapped multiline ime keeps focus after escape" : "native multiline ime keeps focus after escape");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));
        Require(window.Host().GetFocusControl() == field,
                wrapped ? "native wrapped multiline ime keeps tab from advancing focus" : "native multiline ime keeps tab from advancing focus");
        Require(window.Host().DebugHasActiveNativeTextInputSession(),
                wrapped ? "native wrapped multiline ime keeps session active after tab" : "native multiline ime keeps session active after tab");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_ENDCOMPOSITION, 0, 0));

        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_RETURN, 0));
        Require(defaultCount == 0u,
                wrapped ? "native wrapped multiline return remains text-owned after composition ends"
                        : "native multiline return remains text-owned after composition ends");

        window.Host().SetFocusControl(field);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_ESCAPE, 0));
        Require(cancelCount == 1u,
                wrapped ? "native wrapped multiline escape resumes cancel routing after composition ends"
                        : "native multiline escape resumes cancel routing after composition ends");

        window.Host().SetFocusControl(field);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));
        Require(window.Host().GetFocusControl() == nextButton,
                wrapped ? "native wrapped multiline tab resumes focus traversal after composition ends"
                        : "native multiline tab resumes focus traversal after composition ends");
        Require(! window.Host().DebugHasActiveNativeTextInputSession(),
                wrapped ? "native wrapped multiline session deactivates after tab leaves field"
                        : "native multiline session deactivates after tab leaves field");
    };

    runScenario(false);
    runScenario(true);
}

void TestNativeTextInputBackendImeResultOnlyResumesHostKeyRouting()
{
    using namespace RedSalamander::DxUi;

    const auto runScenario = [](bool multiline, bool wrapped)
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root          = std::make_unique<Panel>();
        auto* field        = root->AddChild<TextField>(wrapped ? std::wstring(kWrappedMultilineClipboardTextForTest)
                                                               : (multiline ? std::wstring(L"alpha\nbeta") : std::wstring(L"alpha")));
        auto* nextButton   = root->AddChild<Button>(L"Next");
        auto* cancelButton = root->AddChild<Button>(L"Cancel");
        field->SetMultiline(multiline || wrapped);
        field->SetBounds(wrapped ? D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f)
                                 : (multiline ? D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f) : D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f)));
        nextButton->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));
        cancelButton->SetBounds(D2D1::RectF(140.0f, 112.0f, 260.0f, 140.0f));

        size_t defaultCount = 0u;
        size_t cancelCount  = 0u;
        nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
        cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

        window.Host().SetRoot(std::move(root));
        window.Host().SetDefaultButton(nextButton);
        window.Host().SetCancelButton(cancelButton);
        window.Host().SetFocusControl(field);

        static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

        NativeTextInputImePayload resultPayload;
        resultPayload.hasResultString = true;
        resultPayload.resultString    = L"R";
        window.Host().DebugSetNativeTextInputImePayloadForTest(resultPayload);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_RESULTSTR));

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native ime result-only test reads native session state");
        Require(! state.compositionStartIndex.has_value() && ! state.compositionEndIndex.has_value(), "native ime result-only clears composition ownership");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_RETURN, 0));
        if (multiline || wrapped)
        {
            Require(defaultCount == 0u,
                    wrapped ? "native wrapped multiline ime result-only keeps return text-owned" : "native multiline ime result-only keeps return text-owned");
            Require(window.Host().GetFocusControl() == field,
                    wrapped ? "native wrapped multiline ime result-only return keeps focus" : "native multiline ime result-only return keeps focus");
        }
        else
        {
            Require(defaultCount == 1u, "native single-line ime result-only return resumes default-button routing");
        }

        window.Host().SetFocusControl(field);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_ESCAPE, 0));
        Require(cancelCount == 1u,
                wrapped ? "native wrapped multiline ime result-only escape resumes cancel routing"
                        : (multiline ? "native multiline ime result-only escape resumes cancel routing"
                                     : "native single-line ime result-only escape resumes cancel routing"));

        window.Host().SetFocusControl(field);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));
        Require(window.Host().GetFocusControl() == nextButton,
                wrapped ? "native wrapped multiline ime result-only tab resumes focus traversal"
                        : (multiline ? "native multiline ime result-only tab resumes focus traversal"
                                     : "native single-line ime result-only tab resumes focus traversal"));
        Require(! window.Host().DebugHasActiveNativeTextInputSession(), "native ime result-only tab deactivates the native session");
    };

    runScenario(false, false);
    runScenario(true, false);
    runScenario(true, true);
}

void TestNativeTextInputBackendImeResultAndCompositionKeepsKeyOwnership()
{
    using namespace RedSalamander::DxUi;

    const auto runScenario = [](bool multiline, bool wrapped)
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root          = std::make_unique<Panel>();
        auto* field        = root->AddChild<TextField>(wrapped ? std::wstring(kWrappedMultilineClipboardTextForTest)
                                                               : (multiline ? std::wstring(L"alpha\nbeta") : std::wstring(L"alpha")));
        auto* nextButton   = root->AddChild<Button>(L"Next");
        auto* cancelButton = root->AddChild<Button>(L"Cancel");
        field->SetMultiline(multiline || wrapped);
        field->SetBounds(wrapped ? D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f)
                                 : (multiline ? D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f) : D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f)));
        nextButton->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));
        cancelButton->SetBounds(D2D1::RectF(140.0f, 112.0f, 260.0f, 140.0f));

        size_t defaultCount = 0u;
        size_t cancelCount  = 0u;
        nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
        cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

        window.Host().SetRoot(std::move(root));
        window.Host().SetDefaultButton(nextButton);
        window.Host().SetCancelButton(cancelButton);
        window.Host().SetFocusControl(field);

        static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

        NativeTextInputImePayload payload;
        payload.hasResultString      = true;
        payload.resultString         = L"R";
        payload.hasCompositionString = true;
        payload.compositionString    = L"ime";
        window.Host().DebugSetNativeTextInputImePayloadForTest(payload);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_RESULTSTR | GCS_COMPSTR));

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native ime continuing-composition test reads native session state");
        Require(state.compositionStartIndex.has_value() && state.compositionEndIndex.has_value(),
                "native ime continuing-composition keeps active composition ownership");

        const std::wstring textBeforeReturn(field->GetText());
        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_RETURN, 0));
        Require(defaultCount == 0u,
                wrapped ? "native wrapped multiline continuing ime keeps return owned"
                        : (multiline ? "native multiline continuing ime keeps return owned" : "native single-line continuing ime keeps return owned"));
        Require(field->GetText() == textBeforeReturn, "native continuing ime return does not mutate retained text");
        Require(window.Host().GetFocusControl() == field, "native continuing ime keeps focus after return");
        Require(window.Host().DebugHasActiveNativeTextInputSession(), "native continuing ime keeps session active after return");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_ESCAPE, 0));
        Require(cancelCount == 0u,
                wrapped ? "native wrapped multiline continuing ime keeps escape owned"
                        : (multiline ? "native multiline continuing ime keeps escape owned" : "native single-line continuing ime keeps escape owned"));
        Require(window.Host().GetFocusControl() == field, "native continuing ime keeps focus after escape");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));
        Require(window.Host().GetFocusControl() == field, "native continuing ime keeps tab from advancing focus");
        Require(window.Host().DebugHasActiveNativeTextInputSession(), "native continuing ime keeps session active after tab");
    };

    runScenario(false, false);
    runScenario(true, false);
    runScenario(true, true);
}

void TestNativeTextInputBackendSyncsPrintableCharIntoSessionState()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'Z', 0));

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after printable character input");
    Require(field->GetText() == L"alphaZ", "native host wm_char updates the retained text field");
    Require(state.text == L"alphaZ", "native session state mirrors retained text after printable character input");
    Require(state.caretIndex == 6u, "native session state mirrors caret after printable character input");
    Require(window.Host().DebugGetNativeTextInputEventCounters().synchronizationCount > 0u,
            "native text input counts synchronization after printable character input");
}

void TestNativeTextInputBackendSingleLineTabCharAndPasteReplacementSyncState()
{
    using namespace RedSalamander::DxUi;

    const bool edited = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha beta gamma");
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'\t', 0));
        NativeTextInputState state;
        if (field->GetText() != L"alpha beta gamma" || ! window.Host().TryReadNativeTextInputState(field, state) || state.text != L"alpha beta gamma" ||
            ! window.Host().DebugHasActiveNativeTextInputSession())
        {
            return false;
        }

        field->SetSelectionRange(6u, 10u);
        window.Host().SyncTextInput(field);
        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"OMEGA"), "clipboard prepared before native partial-selection paste");
        SendNativeCtrlKey(window.Hwnd(), 'V');

        if (! window.Host().TryReadNativeTextInputState(field, state))
        {
            return false;
        }
        return field->GetText() == L"alpha OMEGA gamma" && state.text == field->GetText() && state.caretIndex == 11u &&
               ! state.selectionAnchorIndex.has_value();
    });
    Require(edited, "native single-line host ignores tab characters and syncs partial-selection paste state");
}

void TestNativeTextInputBackendStateMirrorsInheritedFlowDirection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root = std::make_unique<Panel>();
    root->SetFlowDirection(FlowDirection::RightToLeft);
    auto* field = root->AddChild<TextField>(L"abc \x05D0\x05D1\x05D2 123");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after rtl inherited focus");
    Require(state.flowDirection == FlowDirection::RightToLeft, "native session state mirrors inherited rtl flow direction");
    Require(state.readingDirection == DWRITE_READING_DIRECTION_RIGHT_TO_LEFT, "native session state mirrors rtl DirectWrite reading direction");
}

void TestNativeTextInputBackendSyncsFocusedInheritedFlowDirectionChanges()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* panel = root.get();
    auto* field = root->AddChild<TextField>(L"abc \x05D0\x05D1\x05D2 123");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable before inherited flow change");
    Require(state.flowDirection == FlowDirection::LeftToRight, "native session state starts with ltr inherited flow direction");

    const uint64_t beforeSyncCount = window.Host().DebugGetNativeTextInputEventCounters().synchronizationCount;
    panel->SetFlowDirection(FlowDirection::RightToLeft);

    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after inherited flow change");
    Require(state.flowDirection == FlowDirection::RightToLeft, "native session state refreshes after inherited rtl flow direction change");
    Require(state.readingDirection == DWRITE_READING_DIRECTION_RIGHT_TO_LEFT, "native session state refreshes DirectWrite reading direction after flow change");
    Require(window.Host().DebugGetNativeTextInputEventCounters().synchronizationCount > beforeSyncCount,
            "native text input counts synchronization after inherited flow direction change");
}

void TestNativeTextInputBackendEditableComboSyncsTextAndFlowDirection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* panel = root.get();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});
    combo->SetText(L"al");

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(combo);

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(combo, state), "native editable combo session state is readable after focus");
    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native editable combo does not create a hidden bridge child window");
    Require(window.Host().DebugHasActiveNativeTextInputSession(), "native editable combo activates the native text input session");
    Require(state.text == L"al", "native editable combo session mirrors starting text");
    Require(state.caretIndex == 2u, "native editable combo session mirrors starting caret");
    Require(! state.multiline, "native editable combo session is single-line");
    Require(! state.masked, "native editable combo session is not masked");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'p', 0));
    Require(window.Host().TryReadNativeTextInputState(combo, state), "native editable combo session state is readable after character input");
    Require(combo->GetText() == L"alp", "native editable combo host wm_char updates retained text");
    Require(state.text == L"alp", "native editable combo session mirrors character input");
    Require(state.caretIndex == 3u, "native editable combo session mirrors caret after character input");

    panel->SetFlowDirection(FlowDirection::RightToLeft);
    Require(window.Host().TryReadNativeTextInputState(combo, state), "native editable combo session state is readable after inherited flow change");
    Require(state.flowDirection == FlowDirection::RightToLeft, "native editable combo session refreshes inherited rtl flow direction");
    Require(state.readingDirection == DWRITE_READING_DIRECTION_RIGHT_TO_LEFT,
            "native editable combo session refreshes DirectWrite reading direction after inherited flow change");
}

void TestNativeTextInputBackendEditableComboCommandsAndPopupSyncState()
{
    using namespace RedSalamander::DxUi;

    const bool edited = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* combo = root->AddChild<ComboBox>();
        combo->SetEditable(true);
        combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 32.0f));
        combo->SetItems(
            {ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}, ComboBox::Item{L"gamma", L"Gamma"}, ComboBox::Item{L"omega", L"Omega"}});
        combo->SetText(L"alpha beta");

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(combo);
        combo->SetEditableSelectionRange(6u, 10u);
        window.Host().SyncTextInput(combo);

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native editable combo copy");
        SendNativeCtrlKey(window.Hwnd(), 'C');
        std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"beta" || combo->GetText() != L"alpha beta")
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'X');
        clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"beta" || combo->GetText() != L"alpha ")
        {
            return false;
        }

        NativeTextInputState state;
        if (! window.Host().TryReadNativeTextInputState(combo, state) || state.text != L"alpha " || state.caretIndex != 6u ||
            state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Z');
        const std::optional<std::pair<size_t, size_t>> restoredSelection = combo->GetEditableSelectionRange();
        if (combo->GetText() != L"alpha beta" || ! restoredSelection || restoredSelection.value().first != 6u || restoredSelection.value().second != 10u)
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Y');
        if (combo->GetText() != L"alpha " || ! window.Host().TryReadNativeTextInputState(combo, state) || state.text != L"alpha " || state.caretIndex != 6u ||
            state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"omega\r\npsi\t"), "clipboard initialized before native editable combo paste");
        SendNativeCtrlKey(window.Hwnd(), 'V');
        if (combo->GetText() != L"alpha omegapsi" || ! window.Host().TryReadNativeTextInputState(combo, state) || state.text != L"alpha omegapsi" ||
            state.caretIndex != combo->GetText().size() || state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Z');
        if (combo->GetText() != L"alpha " || ! window.Host().TryReadNativeTextInputState(combo, state) || state.text != L"alpha " || state.caretIndex != 6u ||
            state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"XYZ"), "clipboard initialized before native editable combo shift+insert");
        SendNativeShiftKey(window.Hwnd(), VK_INSERT);
        if (combo->GetText() != L"alpha XYZ" || ! window.Host().TryReadNativeTextInputState(combo, state) || state.text != L"alpha XYZ" ||
            state.caretIndex != combo->GetText().size() || state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        combo->SetEditableSelectionRange(6u, 9u);
        window.Host().SyncTextInput(combo);
        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native editable combo shift+delete");
        SendNativeShiftKey(window.Hwnd(), VK_DELETE);
        clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"XYZ" || combo->GetText() != L"alpha ")
        {
            return false;
        }

        SendNativeAltKey(window.Hwnd(), VK_DOWN);
        if (! combo->DebugIsPopupOpen() || ! window.Host().TryReadNativeTextInputState(combo, state) || state.text != L"alpha ")
        {
            return false;
        }

        SendNativeKey(window.Hwnd(), VK_ESCAPE);
        return ! combo->DebugIsPopupOpen() && window.Host().DebugHasActiveNativeTextInputSession() && window.Host().TryReadNativeTextInputState(combo, state) &&
               state.text == L"alpha ";
    });
    Require(edited, "native editable combo command keys, undo/redo, and popup keys mutate retained text and sync native state");
}

void TestNativeTextInputBackendEditableComboExactMatchCommandsSyncSelection()
{
    using namespace RedSalamander::DxUi;

    const bool edited = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* combo = root->AddChild<ComboBox>();
        combo->SetEditable(true);
        combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 32.0f));
        combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"},
                         ComboBox::Item{L"alphax", L"Alphax"},
                         ComboBox::Item{L"alphabeta", L"Alphabeta"},
                         ComboBox::Item{L"beta", L"Beta"}});
        combo->SetText(L"alpha");

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(combo);
        combo->SetEditableSelectionRange(combo->GetText().size(), combo->GetText().size());
        window.Host().SyncTextInput(combo);

        if (! combo->GetSelectedIndex().has_value() || combo->GetSelectedIndex().value() != 0u)
        {
            return false;
        }

        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'x', 0));
        NativeTextInputState state{};
        if (combo->GetText() != L"alphax" || ! combo->GetSelectedIndex().has_value() || combo->GetSelectedIndex().value() != 1u ||
            ! window.Host().TryReadNativeTextInputState(combo, state) || state.text != L"alphax" || state.caretIndex != 6u)
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Z');
        if (combo->GetText() != L"alpha" || ! combo->GetSelectedIndex().has_value() || combo->GetSelectedIndex().value() != 0u)
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Y');
        if (combo->GetText() != L"alphax" || ! combo->GetSelectedIndex().has_value() || combo->GetSelectedIndex().value() != 1u)
        {
            return false;
        }

        combo->SetText(L"alpha");
        combo->SetEditableSelectionRange(combo->GetText().size(), combo->GetText().size());
        window.Host().SyncTextInput(combo);
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"beta"))
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'V');
        return combo->GetText() == L"alphabeta" && combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == 2u &&
               window.Host().TryReadNativeTextInputState(combo, state) && state.text == L"alphabeta" && state.caretIndex == combo->GetText().size();
    });

    Require(edited, "native editable combo exact-match typing, undo/redo, and paste keep selected item synchronized");
}

void TestNativeTextInputBackendEditableComboDeleteKeysAndPathWordDeleteSyncState()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));
    combo->SetText(L"alpha beta gamma");

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(combo);

    NativeTextInputState state{};
    combo->SetEditableSelectionRange(6u, 6u);
    window.Host().SyncTextInput(combo);
    SendNativeKey(window.Hwnd(), VK_DELETE);
    Require(combo->GetText() == L"alpha eta gamma", "native editable combo delete removes the next character");
    Require(window.Host().TryReadNativeTextInputState(combo, state), "native editable combo delete keeps state readable");
    Require(state.text == combo->GetText() && state.caretIndex == 6u && ! state.selectionAnchorIndex.has_value(),
            "native editable combo delete syncs collapsed caret state");

    combo->SetEditableSelectionRange(10u, 10u);
    window.Host().SyncTextInput(combo);
    SendNativeCtrlKey(window.Hwnd(), VK_DELETE);
    Require(combo->GetText() == L"alpha eta ", "native editable combo ctrl+delete removes the next word segment");
    Require(window.Host().TryReadNativeTextInputState(combo, state), "native editable combo ctrl+delete keeps state readable");
    Require(state.text == combo->GetText() && state.caretIndex == 10u && ! state.selectionAnchorIndex.has_value(),
            "native editable combo ctrl+delete syncs collapsed caret state");

    combo->SetText(L"select all candidate");
    window.Host().SyncTextInput(combo);
    SendNativeCtrlKey(window.Hwnd(), 'A');
    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'Z', 0));
    Require(combo->GetText() == L"Z", "native editable combo ctrl+a followed by typing replaces the full text");

    const std::wstring pathText = L"C:\\repo\\.build\\x64\\Debug";
    combo->SetText(pathText);
    combo->SetEditableSelectionRange(pathText.size(), pathText.size());
    window.Host().SyncTextInput(combo);
    SendNativeCtrlKey(window.Hwnd(), VK_BACK);
    Require(combo->GetText() == L"C:\\repo\\.build\\x64\\", "native editable combo ctrl+backspace deletes only the previous path segment");
    Require(combo->GetText().find(L'\x7F') == std::wstring_view::npos, "native editable combo ctrl+backspace does not insert the translated delete character");

    combo->SetText(pathText);
    combo->SetEditableSelectionRange(pathText.size(), pathText.size());
    window.Host().SyncTextInput(combo);
    SendNativeCtrlKey(window.Hwnd(), VK_BACK);
    SendNativeCtrlKey(window.Hwnd(), VK_BACK);
    Require(combo->GetText() == L"C:\\repo\\.build\\", "native editable combo repeated ctrl+backspace keeps deleting meaningful path segments");
    Require(window.Host().TryReadNativeTextInputState(combo, state), "native editable combo path deletion keeps state readable");
    Require(state.text == combo->GetText() && state.caretIndex == combo->GetText().size() && ! state.selectionAnchorIndex.has_value(),
            "native editable combo path deletion syncs native state");
}

void TestNativeTextInputBackendSyncsPrintableSysCharIntoSessionState()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_SYSCHAR, L'Z', 0));

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after printable system character input");
    Require(field->GetText() == L"alphaZ", "native host wm_syschar updates the retained text field");
    Require(state.text == L"alphaZ", "native session state mirrors retained text after printable system character input");
    Require(state.caretIndex == 6u, "native session state mirrors caret after printable system character input");
}

void TestNativeTextInputBackendSyncsKeySelectionIntoSessionState()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_LEFT, 0));

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after key input");
    Require(state.text == L"alpha", "native session state preserves text after left arrow");
    Require(state.caretIndex == 4u, "native session state mirrors caret after left arrow");
}

void TestNativeTextInputBackendExposesBackendNeutralTextInputState()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_LEFT, 0));

    TextInputState state;
    Require(window.Host().TryReadTextInputState(field, state), "native session is readable through backend-neutral text input state");
    Require(state.text == L"alpha", "backend-neutral text input state mirrors native text");
    Require(state.caretIndex == 4u, "backend-neutral text input state mirrors native caret");
    Require(! state.selectionAnchorIndex.has_value(), "backend-neutral text input state mirrors collapsed native selection");
}

void TestNativeTextInputBackendSurrogatePairDeletionSyncsState()
{
    using namespace RedSalamander::DxUi;

    auto makeText = []()
    {
        std::wstring text = L"A";
        text.push_back(static_cast<wchar_t>(0xD83D));
        text.push_back(static_cast<wchar_t>(0xDE00));
        text.push_back(L'B');
        return text;
    };

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(makeText());
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeKey(window.Hwnd(), VK_LEFT);
        SendNativeKey(window.Hwnd(), VK_BACK);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after surrogate-pair backspace");
        Require(field->GetText() == L"AB", "native backspace removes the full surrogate pair instead of one code unit");
        Require(state.text == L"AB", "native session state mirrors surrogate-pair backspace text");
        Require(state.caretIndex == 1u, "native session state mirrors surrogate-pair backspace caret");
    }

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(makeText());
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeKey(window.Hwnd(), VK_HOME);
        SendNativeKey(window.Hwnd(), VK_RIGHT);
        SendNativeKey(window.Hwnd(), VK_DELETE);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after surrogate-pair delete");
        Require(field->GetText() == L"AB", "native delete removes the full surrogate pair instead of one code unit");
        Require(state.text == L"AB", "native session state mirrors surrogate-pair delete text");
        Require(state.caretIndex == 1u, "native session state mirrors surrogate-pair delete caret");
    }
}

void TestNativeTextInputBackendEmojiZwJDeletionSyncsState()
{
    using namespace RedSalamander::DxUi;

    auto makeText = []()
    {
        std::wstring text = L"A";
        text += MakeWomanTechnologistTextElement();
        text.push_back(L'B');
        return text;
    };

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(makeText());
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeKey(window.Hwnd(), VK_LEFT);
        SendNativeKey(window.Hwnd(), VK_BACK);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after emoji ZWJ backspace");
        Require(field->GetText() == L"AB", "native backspace removes the full emoji ZWJ text element");
        Require(state.text == L"AB", "native session state mirrors emoji ZWJ backspace text");
        Require(state.caretIndex == 1u, "native session state mirrors emoji ZWJ backspace caret");
    }

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(makeText());
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeKey(window.Hwnd(), VK_HOME);
        SendNativeKey(window.Hwnd(), VK_RIGHT);
        SendNativeKey(window.Hwnd(), VK_DELETE);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after emoji ZWJ delete");
        Require(field->GetText() == L"AB", "native delete removes the full emoji ZWJ text element");
        Require(state.text == L"AB", "native session state mirrors emoji ZWJ delete text");
        Require(state.caretIndex == 1u, "native session state mirrors emoji ZWJ delete caret");
    }
}

void TestNativeTextInputBackendRegionalIndicatorFlagDeletionSyncsState()
{
    using namespace RedSalamander::DxUi;

    auto makeText = []()
    {
        std::wstring text = L"A";
        text += MakeUsFlagTextElement();
        text.push_back(L'B');
        return text;
    };

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(makeText());
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeKey(window.Hwnd(), VK_LEFT);
        SendNativeKey(window.Hwnd(), VK_BACK);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after regional-indicator flag backspace");
        Require(field->GetText() == L"AB", "native backspace removes the full regional-indicator flag text element");
        Require(state.text == L"AB", "native session state mirrors regional-indicator flag backspace text");
        Require(state.caretIndex == 1u, "native session state mirrors regional-indicator flag backspace caret");
    }

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(makeText());
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeKey(window.Hwnd(), VK_HOME);
        SendNativeKey(window.Hwnd(), VK_RIGHT);
        SendNativeKey(window.Hwnd(), VK_DELETE);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after regional-indicator flag delete");
        Require(field->GetText() == L"AB", "native delete removes the full regional-indicator flag text element");
        Require(state.text == L"AB", "native session state mirrors regional-indicator flag delete text");
        Require(state.caretIndex == 1u, "native session state mirrors regional-indicator flag delete caret");
    }
}

void TestNativeTextInputBackendEmojiSuffixDeletionSyncsState()
{
    using namespace RedSalamander::DxUi;

    const auto verifyBackspace = [](const std::wstring& textElement, const char* textMessage, const char* stateMessage)
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root         = std::make_unique<Panel>();
        std::wstring text = L"A";
        text += textElement;
        text.push_back(L'B');
        auto* field = root->AddChild<TextField>(text);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeKey(window.Hwnd(), VK_LEFT);
        SendNativeKey(window.Hwnd(), VK_BACK);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after emoji suffix backspace");
        Require(field->GetText() == L"AB", textMessage);
        Require(state.text == L"AB", stateMessage);
        Require(state.caretIndex == 1u, "native session state mirrors emoji suffix backspace caret");
    };

    const auto verifyDelete = [](const std::wstring& textElement, const char* textMessage, const char* stateMessage)
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root         = std::make_unique<Panel>();
        std::wstring text = L"A";
        text += textElement;
        text.push_back(L'B');
        auto* field = root->AddChild<TextField>(text);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeKey(window.Hwnd(), VK_HOME);
        SendNativeKey(window.Hwnd(), VK_RIGHT);
        SendNativeKey(window.Hwnd(), VK_DELETE);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after emoji suffix delete");
        Require(field->GetText() == L"AB", textMessage);
        Require(state.text == L"AB", stateMessage);
        Require(state.caretIndex == 1u, "native session state mirrors emoji suffix delete caret");
    };

    verifyBackspace(MakeHeartVariationTextElement(),
                    "native backspace removes the variation-selector emoji text element",
                    "native session state mirrors variation-selector emoji backspace text");
    verifyDelete(MakeHeartVariationTextElement(),
                 "native delete removes the variation-selector emoji text element",
                 "native session state mirrors variation-selector emoji delete text");
    verifyBackspace(MakeThumbsUpMediumSkinToneTextElement(),
                    "native backspace removes the skin-tone emoji text element",
                    "native session state mirrors skin-tone emoji backspace text");
    verifyDelete(MakeThumbsUpMediumSkinToneTextElement(),
                 "native delete removes the skin-tone emoji text element",
                 "native session state mirrors skin-tone emoji delete text");
}

void TestNativeTextInputBackendEmojiShiftSelectionSyncsState()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    std::wstring text = L"A";
    text += MakeWomanTechnologistTextElement();
    text.push_back(L'B');

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(text);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    SendNativeKey(window.Hwnd(), VK_LEFT);
    SendNativeShiftKey(window.Hwnd(), VK_LEFT);

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after emoji shift-selection");
    Require(state.caretIndex == 1u, "native shift-left lands at the start of the emoji text element");
    Require(state.selectionAnchorIndex.has_value() && state.selectionAnchorIndex.value() == 6u,
            "native shift-left anchors at the end of the emoji text element");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, static_cast<WPARAM>(L'X'), 0));

    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after emoji selection replacement");
    Require(field->GetText() == L"AXB", "native emoji shift-selection replacement removes the full text element");
    Require(state.text == L"AXB", "native session state mirrors emoji shift-selection replacement text");
    Require(state.caretIndex == 2u, "native session state mirrors emoji shift-selection replacement caret");
}

void TestNativeTextInputBackendCtrlWordDeletionSyncsState()
{
    using namespace RedSalamander::DxUi;

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha beta");
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeCtrlKey(window.Hwnd(), VK_BACK);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after ctrl+backspace");
        Require(field->GetText() == L"alpha ", "native ctrl+backspace deletes the previous word");
        Require(state.text == L"alpha ", "native session state mirrors ctrl+backspace text");
        Require(state.caretIndex == field->GetText().size(), "native session state mirrors ctrl+backspace caret");
    }

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha beta gamma");
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeKey(window.Hwnd(), VK_HOME);
        SendNativeCtrlKey(window.Hwnd(), VK_DELETE);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after ctrl+delete");
        Require(field->GetText() == L"beta gamma", "native ctrl+delete deletes the next word and following spacing");
        Require(state.text == L"beta gamma", "native session state mirrors ctrl+delete text");
        Require(state.caretIndex == 0u, "native session state mirrors ctrl+delete caret");
    }
}

void TestNativeTextInputBackendPointerCaretPlacementSyncsState()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(50, 12)));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, MAKELPARAM(50, 12)));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'X', 0));

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after pointer caret placement");
    const size_t insertionIndex = field->GetText().find(L'X');
    Require(insertionIndex >= 3u && insertionIndex <= 8u, "native pointer caret placement inserts text near the clicked location");
    Require(state.text == field->GetText(), "native session state mirrors pointer-placed character insertion text");
    Require(state.caretIndex == insertionIndex + 1u, "native session state mirrors pointer-placed character insertion caret");
}

void TestNativeTextInputBackendSingleLineDoubleClickSelectsWordOnHostHwnd()
{
    using namespace RedSalamander::DxUi;

    constexpr std::wstring_view kText = L"Contact alpha@example-domain.com now";
    const size_t expectedStart        = kText.find(L"domain");
    Require(expectedStart != std::wstring_view::npos, "native double-click word-selection test locates target word");
    const size_t expectedEnd = expectedStart + std::wstring_view(L"domain").size();

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kText));
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 420.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native double-click word-selection test does not create a hidden bridge child");
    Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native double-click word-selection test keeps input on the host hwnd");

    const POINT clickPoint =
        GetNativeTextClientPointForCaretIndex(window, *field, expectedStart + 1u, "native double-click word-selection test can locate the target caret");
    const LPARAM clickLp = MAKELPARAM(clickPoint.x, clickPoint.y);

    field->SetSelectionRange(0u, field->GetText().size());
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, clickLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, clickLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDBLCLK, MK_LBUTTON, clickLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, clickLp));

    RequireNativeTextSelection(
        window.Host(), *field, expectedStart, expectedEnd, "native host-HWND double-click selects exactly the punctuation-delimited word");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'X', 0));

    NativeTextInputState state{};
    Require(window.Host().TryReadNativeTextInputState(field, state), "native double-click replacement keeps session state readable");
    Require(field->GetText() == L"Contact alpha@example-X.com now", "native double-click selected word is replaced by typed text");
    Require(state.text == field->GetText(), "native double-click replacement syncs retained text into native state");
}

void TestNativeTextInputBackendSingleLineRepeatedClicksWithoutClassDoubleClicksSelectWord()
{
    using namespace RedSalamander::DxUi;

    constexpr std::wstring_view kText = L"Contact alpha@example-domain.com now";
    const size_t expectedStart        = kText.find(L"domain");
    Require(expectedStart != std::wstring_view::npos, "native repeated-click word-selection test locates target word");
    const size_t expectedEnd = expectedStart + std::wstring_view(L"domain").size();

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kText));
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 420.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require((GetClassLongPtrW(window.Hwnd(), GCL_STYLE) & CS_DBLCLKS) == 0, "native repeated-click test uses a host class without CS_DBLCLKS");
    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native repeated-click word-selection test does not create a hidden bridge child");
    Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native repeated-click word-selection test keeps input on the host hwnd");

    const POINT clickPoint =
        GetNativeTextClientPointForCaretIndex(window, *field, expectedStart + 1u, "native repeated-click word-selection test can locate the target caret");
    const LPARAM clickLp = MAKELPARAM(clickPoint.x, clickPoint.y);

    field->SetSelectionRange(0u, field->GetText().size());
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, clickLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, clickLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, clickLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, clickLp));

    RequireNativeTextSelection(
        window.Host(), *field, expectedStart, expectedEnd, "native repeated clicks synthesize double-click word selection on host classes without CS_DBLCLKS");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'X', 0));

    NativeTextInputState state{};
    Require(window.Host().TryReadNativeTextInputState(field, state), "native repeated-click replacement keeps session state readable");
    Require(field->GetText() == L"Contact alpha@example-X.com now", "native repeated-click selected word is replaced by typed text");
    Require(state.text == field->GetText(), "native repeated-click replacement syncs retained text into native state");
}

void TestNativeTextInputBackendSingleLineThirdClickSelectsAllOnHostHwnd()
{
    using namespace RedSalamander::DxUi;

    constexpr std::wstring_view kText = L"alpha beta";
    const size_t targetIndex          = kText.find(L"beta");
    Require(targetIndex != std::wstring_view::npos, "native third-click select-all test locates target word");

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kText));
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native third-click select-all test does not create a hidden bridge child");
    Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native third-click select-all test keeps input on the host hwnd");

    const POINT clickPoint =
        GetNativeTextClientPointForCaretIndex(window, *field, targetIndex + 1u, "native third-click select-all test can locate the target caret");
    const LPARAM clickLp = MAKELPARAM(clickPoint.x, clickPoint.y);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, clickLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, clickLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDBLCLK, MK_LBUTTON, clickLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, clickLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, clickLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, clickLp));

    RequireNativeTextSelection(window.Host(), *field, 0u, field->GetText().size(), "native host-HWND third click selects the full text");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'Q', 0));

    NativeTextInputState state{};
    Require(window.Host().TryReadNativeTextInputState(field, state), "native third-click replacement keeps session state readable");
    Require(field->GetText() == L"Q", "native third-click selected all text is replaced by typed text");
    Require(state.text == L"Q" && state.caretIndex == 1u && ! state.selectionAnchorIndex.has_value(),
            "native third-click replacement syncs collapsed caret state");
}

void TestNativeTextInputBackendSingleLineDragSelectionReplacesRangeOnHostHwnd()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native drag-selection test does not create a hidden bridge child");
    Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native drag-selection test keeps input on the host hwnd");

    POINT startPoint = GetNativeTextClientPointForCaretIndex(window, *field, 0u, "native drag-selection test can locate the start caret");
    POINT endPoint   = GetNativeTextClientPointForCaretIndex(window, *field, field->GetText().size(), "native drag-selection test can locate the end caret");
    startPoint.x -= 4;
    endPoint.x += 4;

    const LPARAM startLp = MAKELPARAM(startPoint.x, startPoint.y);
    const LPARAM endLp   = MAKELPARAM(endPoint.x, endPoint.y);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, startLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_MOUSEMOVE, MK_LBUTTON, endLp));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, endLp));

    RequireNativeTextSelection(window.Host(), *field, 0u, field->GetText().size(), "native host-HWND drag selects the full text range");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'Z', 0));

    NativeTextInputState state{};
    Require(window.Host().TryReadNativeTextInputState(field, state), "native drag-selection replacement keeps session state readable");
    Require(field->GetText() == L"Z", "native dragged selection is replaced by typed text");
    Require(state.text == L"Z" && state.caretIndex == 1u && ! state.selectionAnchorIndex.has_value(),
            "native drag-selection replacement syncs collapsed caret state");
}

void TestNativeTextInputBackendMixedBiDiDragSelectionCopiesLogicalOrderOnHostHwnd()
{
    using namespace RedSalamander::DxUi;

    const auto verifyScenario = [](FlowDirection flowDirection, bool logicalStartIsVisuallyBeforeLogicalEnd)
    {
        constexpr std::wstring_view mixedBiDiText = L"abc \x05D0\x05D1\x05D2";

        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root = std::make_unique<Panel>();
        root->SetFlowDirection(flowDirection);
        auto* field = root->AddChild<TextField>(std::wstring(mixedBiDiText));
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native mixed BiDi drag keeps the hidden bridge child absent");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native mixed BiDi drag keeps input on the host hwnd");

        POINT logicalStartPoint = GetNativeTextClientPointForCaretIndex(window, *field, 0u, "native mixed BiDi drag locates the logical start caret");
        POINT logicalEndPoint =
            GetNativeTextClientPointForCaretIndex(window, *field, field->GetText().size(), "native mixed BiDi drag locates the cross-script target caret");
        if (logicalStartIsVisuallyBeforeLogicalEnd)
        {
            Require(logicalStartPoint.x < logicalEndPoint.x, "native mixed BiDi ltr drag uses left-to-right visual caret order");
        }
        else
        {
            Require(logicalStartPoint.x > logicalEndPoint.x, "native mixed BiDi rtl drag uses right-to-left visual caret order");
        }

        TextFieldDebugSingleLinePaintState paint{};
        Require(field->DebugGetSingleLinePaintState(window.Host(), paint), "native mixed BiDi drag exposes the single-line text rectangle");
        const auto pointFromDip = [&window](float xDip, float yDip)
        { return POINT{static_cast<LONG>(std::lround(window.Host().DipsToPixels(xDip))), static_cast<LONG>(std::lround(window.Host().DipsToPixels(yDip)))}; };
        const float yDip = (paint.textRect.top + paint.textRect.bottom) * 0.5f;
        POINT dragStart =
            logicalStartIsVisuallyBeforeLogicalEnd ? pointFromDip(paint.textRect.left + 1.0f, yDip) : pointFromDip(paint.textRect.right - 1.0f, yDip);
        POINT dragEnd =
            logicalStartIsVisuallyBeforeLogicalEnd ? pointFromDip(paint.textRect.right - 1.0f, yDip) : pointFromDip(paint.textRect.left + 1.0f, yDip);

        const LPARAM startLp = MAKELPARAM(dragStart.x, dragStart.y);
        const LPARAM endLp   = MAKELPARAM(dragEnd.x, dragEnd.y);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, startLp));
        static_cast<void>(SendMessageW(window.Hwnd(), WM_MOUSEMOVE, MK_LBUTTON, endLp));
        static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, endLp));

        const std::optional<std::pair<size_t, size_t>> selection = field->GetSelectionRange();
        Require(selection.has_value(), "native mixed BiDi drag produces a retained selection");
        Require(selection.value().first < 4u, "native mixed BiDi drag starts before the rtl script run");
        Require(selection.value().second > 4u, "native mixed BiDi drag crosses into the rtl script run");
        Require(selection.value().second <= field->GetText().size(), "native mixed BiDi drag keeps selection within logical text");
        const std::wstring expectedSelectedText(mixedBiDiText.substr(selection.value().first, selection.value().second - selection.value().first));
        RequireNativeTextSelection(
            window.Host(), *field, selection.value().first, selection.value().second, "native mixed BiDi drag syncs cross-script selection to state");
        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native mixed BiDi drag copy");
        static_cast<void>(SendMessageW(window.Hwnd(), WM_COPY, 0, 0));

        const std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        Require(clipboardText.has_value(), "native mixed BiDi drag copy writes readable clipboard text");
        Require(clipboardText.value() != L"sentinel", "native mixed BiDi drag copy routes WM_COPY to the focused text field");
        Require(clipboardText.value().size() == expectedSelectedText.size(), "native mixed BiDi drag copy keeps logical text length");
        Require(clipboardText.value() == expectedSelectedText, "native mixed BiDi drag copy keeps logical UTF-16 order");
        return true;
    };

    Require(RetryClipboardSensitiveAction([&] { return verifyScenario(FlowDirection::LeftToRight, true); }),
            "native mixed BiDi ltr drag selection copies logical UTF-16 order");
    Require(RetryClipboardSensitiveAction([&] { return verifyScenario(FlowDirection::RightToLeft, false); }),
            "native mixed BiDi rtl drag selection copies logical UTF-16 order");
}

void TestNativeTextInputBackendMixedBiDiPointerHitTestMatchesDirectWriteVisualOrderOnHostHwnd()
{
    using namespace RedSalamander::DxUi;

    constexpr std::wstring_view mixedBiDiText = L"abc \x05D0\x05D1\x05D2 123";
    constexpr std::array<DirectWritePointerSpan, 3u> spans{
        DirectWritePointerSpan{0u, 1u}, DirectWritePointerSpan{5u, 6u}, DirectWritePointerSpan{mixedBiDiText.size() - 1u, mixedBiDiText.size()}};

    VerifyNativeTextInputPointerScenarioMatchesDirectWrite(
        mixedBiDiText, FlowDirection::LeftToRight, spans, "native mixed BiDi ltr pointer hit-test matches DirectWrite logical index");
    VerifyNativeTextInputPointerScenarioMatchesDirectWrite(
        mixedBiDiText, FlowDirection::RightToLeft, spans, "native mixed BiDi rtl pointer hit-test matches DirectWrite logical index");
}

void TestNativeTextInputBackendBiDiPointerScenarioMatrixMatchesDirectWriteOnHostHwnd()
{
    using namespace RedSalamander::DxUi;

    VerifyNativeTextInputPointerScenarioMatchesDirectWrite(
        L"abc 123 xyz",
        FlowDirection::LeftToRight,
        std::array<DirectWritePointerSpan, 3u>{DirectWritePointerSpan{0u, 1u}, DirectWritePointerSpan{4u, 5u}, DirectWritePointerSpan{10u, 11u}},
        "native ltr text pointer matrix matches DirectWrite");
    VerifyNativeTextInputPointerScenarioMatchesDirectWrite(
        L"\x05D0\x05D1\x05D2 \x05D3\x05D4\x05D5",
        FlowDirection::RightToLeft,
        std::array<DirectWritePointerSpan, 3u>{DirectWritePointerSpan{0u, 1u}, DirectWritePointerSpan{4u, 5u}, DirectWritePointerSpan{6u, 7u}},
        "native rtl text pointer matrix matches DirectWrite");
    VerifyNativeTextInputPointerScenarioMatchesDirectWrite(
        L"\x0627\x0628\x062C 123 abc",
        FlowDirection::RightToLeft,
        std::array<DirectWritePointerSpan, 3u>{DirectWritePointerSpan{0u, 1u}, DirectWritePointerSpan{5u, 6u}, DirectWritePointerSpan{10u, 11u}},
        "native arabic digits pointer matrix matches DirectWrite");

    std::wstring rtlWithSurrogate = L"\x05D0\x05D1 ";
    rtlWithSurrogate += MakeGrinningFaceTextElement();
    rtlWithSurrogate += L" \x05D2\x05D3";
    VerifyNativeTextInputPointerScenarioMatchesDirectWrite(
        rtlWithSurrogate,
        FlowDirection::RightToLeft,
        std::array<DirectWritePointerSpan, 3u>{DirectWritePointerSpan{0u, 1u}, DirectWritePointerSpan{3u, 5u}, DirectWritePointerSpan{7u, 8u}},
        "native rtl surrogate pointer matrix matches DirectWrite");

    const std::wstring pathLikeRtl = L"C:\\src\\\x05D0\x05D1\x05D2\\file.txt";
    VerifyNativeTextInputPointerScenarioMatchesDirectWrite(
        pathLikeRtl,
        FlowDirection::RightToLeft,
        std::array<DirectWritePointerSpan, 3u>{
            DirectWritePointerSpan{0u, 1u}, DirectWritePointerSpan{7u, 8u}, DirectWritePointerSpan{pathLikeRtl.size() - 1u, pathLikeRtl.size()}},
        "native rtl path-like pointer matrix matches DirectWrite");
}

void TestNativeTextInputBackendBiDiKeyboardLogicalBoundaryCommandsSyncState()
{
    using namespace RedSalamander::DxUi;

    constexpr std::wstring_view originalText = L"ab \x05D0\x05D1 cd";

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root = std::make_unique<Panel>();
    root->SetFlowDirection(FlowDirection::RightToLeft);
    auto* field = root->AddChild<TextField>(std::wstring(originalText));
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native BiDi keyboard test keeps the hidden bridge child absent");
    Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native BiDi keyboard test keeps input on the host hwnd");

    auto resetTextAndCaret = [&](size_t caretIndex)
    {
        field->SetText(std::wstring(originalText));
        field->SetSelectionRange(caretIndex, caretIndex);
        window.Host().SyncTextInput(field);
    };
    auto requireCollapsedState = [&](size_t expectedCaret, const char* context)
    {
        NativeTextInputState state{};
        Require(window.Host().TryReadNativeTextInputState(field, state), context);
        Require(state.flowDirection == FlowDirection::RightToLeft, context);
        Require(state.text == field->GetText(), context);
        Require(state.caretIndex == expectedCaret, context);
        Require(! state.selectionAnchorIndex.has_value(), context);
    };

    resetTextAndCaret(4u);
    SendNativeKey(window.Hwnd(), VK_HOME);
    requireCollapsedState(0u, "native BiDi keyboard Home moves to the logical start");

    resetTextAndCaret(4u);
    SendNativeCtrlKey(window.Hwnd(), VK_END);
    requireCollapsedState(originalText.size(), "native BiDi keyboard Ctrl+End moves to the logical end");

    resetTextAndCaret(originalText.size());
    SendNativeShiftKey(window.Hwnd(), VK_HOME);
    RequireNativeTextSelection(window.Host(), *field, 0u, originalText.size(), "native BiDi keyboard Shift+Home selects to logical start");

    SendNativeKey(window.Hwnd(), VK_RIGHT);
    requireCollapsedState(originalText.size(), "native BiDi keyboard right arrow collapses the full selection to the logical end");

    resetTextAndCaret(0u);
    SendNativeShiftKey(window.Hwnd(), VK_END);
    RequireNativeTextSelection(window.Host(), *field, 0u, originalText.size(), "native BiDi keyboard Shift+End selects to logical end");

    resetTextAndCaret(4u);
    SendNativeKey(window.Hwnd(), VK_LEFT);
    requireCollapsedState(3u, "native BiDi keyboard left arrow keeps current logical-order semantics");
    SendNativeKey(window.Hwnd(), VK_RIGHT);
    requireCollapsedState(4u, "native BiDi keyboard right arrow keeps current logical-order semantics");

    resetTextAndCaret(4u);
    SendNativeKey(window.Hwnd(), VK_BACK);
    requireCollapsedState(3u, "native BiDi keyboard Backspace syncs state after deleting before the script boundary");
    Require(field->GetText() == L"ab \x05D1 cd", "native BiDi keyboard Backspace deletes the previous logical Hebrew code unit");

    resetTextAndCaret(4u);
    SendNativeKey(window.Hwnd(), VK_DELETE);
    requireCollapsedState(4u, "native BiDi keyboard Delete syncs state after deleting at the script boundary");
    Require(field->GetText() == L"ab \x05D0 cd", "native BiDi keyboard Delete deletes the next logical Hebrew code unit");
}

void TestNativeTextInputBackendMixedBiDiEditTransactionsPreserveLogicalOrder()
{
    using namespace RedSalamander::DxUi;

    const bool edited = RetryClipboardSensitiveAction([]()
    {
        constexpr std::wstring_view originalText          = L"ab \x05D0\x05D1 cd/ef 123";
        constexpr std::wstring_view selectedHebrewText    = L"\x05D0\x05D1";
        constexpr std::wstring_view replacementHebrewText = L"\x05D2\x05D3";
        constexpr size_t hebrewStart                      = 3u;
        constexpr size_t hebrewEnd                        = 5u;

        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root = std::make_unique<Panel>();
        root->SetFlowDirection(FlowDirection::RightToLeft);
        auto* field = root->AddChild<TextField>(std::wstring(originalText));
        field->SetBounds(D2D1::RectF(16.0f, 16.0f, 300.0f, 48.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native mixed-BiDi transaction test keeps the hidden bridge child absent");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native mixed-BiDi transaction test keeps input on the host hwnd");

        field->SetSelectionRange(hebrewStart, hebrewEnd);
        window.Host().SyncTextInput(field);
        RequireNativeTextSelection(window.Host(), *field, hebrewStart, hebrewEnd, "native mixed-BiDi transaction starts with logical selection");

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native mixed-BiDi copy");
        SendNativeCtrlKey(window.Hwnd(), 'C');
        std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText.has_value() || clipboardText.value() != selectedHebrewText)
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'X');
        clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        NativeTextInputState state{};
        if (! clipboardText.has_value() || clipboardText.value() != selectedHebrewText || field->GetText() != L"ab  cd/ef 123" ||
            ! window.Host().TryReadNativeTextInputState(field, state) || state.text != field->GetText() || state.caretIndex != hebrewStart ||
            state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Z');
        if (field->GetText() != originalText)
        {
            return false;
        }
        RequireNativeTextSelection(window.Host(), *field, hebrewStart, hebrewEnd, "native mixed-BiDi undo restores logical selection after cut");

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), replacementHebrewText), "clipboard initialized before native mixed-BiDi paste");
        SendNativeCtrlKey(window.Hwnd(), 'V');
        if (field->GetText() != L"ab \x05D2\x05D3 cd/ef 123" || ! window.Host().TryReadNativeTextInputState(field, state) || state.text != field->GetText() ||
            state.caretIndex != hebrewEnd || state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Z');
        if (field->GetText() != originalText)
        {
            return false;
        }
        RequireNativeTextSelection(window.Host(), *field, hebrewStart, hebrewEnd, "native mixed-BiDi undo restores logical selection after paste");

        SendNativeCtrlKey(window.Hwnd(), 'Y');
        if (field->GetText() != L"ab \x05D2\x05D3 cd/ef 123" || ! window.Host().TryReadNativeTextInputState(field, state) || state.text != field->GetText() ||
            state.caretIndex != hebrewEnd || state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        field->SetText(std::wstring(originalText));
        field->SetSelectionRange(hebrewEnd, hebrewEnd);
        window.Host().SyncTextInput(field);
        SendNativeCtrlKey(window.Hwnd(), VK_BACK);
        if (field->GetText() != L"ab  cd/ef 123" || ! window.Host().TryReadNativeTextInputState(field, state) || state.text != field->GetText() ||
            state.caretIndex != hebrewStart || state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        field->SetText(std::wstring(originalText));
        field->SetSelectionRange(6u, 6u);
        window.Host().SyncTextInput(field);
        SendNativeCtrlKey(window.Hwnd(), VK_DELETE);
        return field->GetText() == L"ab \x05D0\x05D1 /ef 123" && window.Host().TryReadNativeTextInputState(field, state) && state.text == field->GetText() &&
               state.caretIndex == 6u && ! state.selectionAnchorIndex.has_value();
    });

    Require(edited, "native mixed-BiDi edit transactions preserve logical order and state across copy/cut/paste/undo/redo/word delete");
}

void TestNativeTextInputBackendPointerHitTestDoesNotSplitEmojiTextElements()
{
    using namespace RedSalamander::DxUi;

    const auto verifyPointerInsertion = [](const std::wstring& textElement, const char* message)
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        std::wstring text = L"A";
        text += textElement;
        text.push_back(L'B');

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(text);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 32.0f));
        field->SetClearButtonEnabled(false);

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeKey(window.Hwnd(), VK_HOME);
        SendNativeKey(window.Hwnd(), VK_RIGHT);
        D2D1_RECT_F elementStartRect{};
        RECT elementStartScreenRect{};
        Require(window.Host().DebugGetNativeTextInputCaretRect(elementStartRect, elementStartScreenRect),
                "native emoji pointer test can measure the element start caret");

        SendNativeKey(window.Hwnd(), VK_RIGHT);
        D2D1_RECT_F elementEndRect{};
        RECT elementEndScreenRect{};
        Require(window.Host().DebugGetNativeTextInputCaretRect(elementEndRect, elementEndScreenRect),
                "native emoji pointer test can measure the element end caret");

        const float clickX = (elementStartRect.left + elementEndRect.left) * 0.5f;
        const float clickY = (elementStartRect.top + elementStartRect.bottom) * 0.5f;
        SendNativeClick(window.Hwnd(), POINT{static_cast<LONG>(std::lround(clickX)), static_cast<LONG>(std::lround(clickY))});
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'X', 0));

        std::wstring beforeElement = L"AX";
        beforeElement += textElement;
        beforeElement.push_back(L'B');
        std::wstring afterElement = L"A";
        afterElement += textElement;
        afterElement.append(L"XB");
        Require(field->GetText() == beforeElement || field->GetText() == afterElement, message);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "native session state is readable after emoji pointer insertion");
        const size_t insertionIndex = field->GetText().find(L'X');
        Require(insertionIndex != std::wstring::npos, "native emoji pointer insertion writes the test character");
        Require(state.text == field->GetText(), "native session state mirrors emoji pointer insertion text");
        Require(state.caretIndex == insertionIndex + 1u, "native session state mirrors emoji pointer insertion caret");
    };

    verifyPointerInsertion(MakeWomanTechnologistTextElement(), "native pointer hit-test does not split a ZWJ emoji text element");
    verifyPointerInsertion(MakeHeartVariationTextElement(), "native pointer hit-test does not split a variation-selector emoji text element");
    verifyPointerInsertion(MakeThumbsUpMediumSkinToneTextElement(), "native pointer hit-test does not split a skin-tone emoji text element");
    verifyPointerInsertion(MakeUsFlagTextElement(), "native pointer hit-test does not split a regional-indicator flag text element");
}

void TestNativeTextInputBackendEditableComboPointerHitTestDoesNotSplitEmojiTextElement()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    const std::wstring textElement = MakeUsFlagTextElement();
    std::wstring text              = L"A";
    text += textElement;
    text.push_back(L'B');

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 32.0f));
    combo->SetText(text);

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(combo);

    SendNativeKey(window.Hwnd(), VK_HOME);
    SendNativeKey(window.Hwnd(), VK_RIGHT);
    D2D1_RECT_F elementStartRect{};
    RECT elementStartScreenRect{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(elementStartRect, elementStartScreenRect),
            "native editable combo emoji pointer test can measure the element start caret");

    SendNativeKey(window.Hwnd(), VK_RIGHT);
    D2D1_RECT_F elementEndRect{};
    RECT elementEndScreenRect{};
    Require(window.Host().DebugGetNativeTextInputCaretRect(elementEndRect, elementEndScreenRect),
            "native editable combo emoji pointer test can measure the element end caret");

    const float clickX = (elementStartRect.left + elementEndRect.left) * 0.5f;
    const float clickY = (elementStartRect.top + elementStartRect.bottom) * 0.5f;
    SendNativeClick(window.Hwnd(), POINT{static_cast<LONG>(std::lround(clickX)), static_cast<LONG>(std::lround(clickY))});
    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'X', 0));

    std::wstring beforeElement = L"AX";
    beforeElement += textElement;
    beforeElement.push_back(L'B');
    std::wstring afterElement = L"A";
    afterElement += textElement;
    afterElement.append(L"XB");
    Require(combo->GetText() == beforeElement || combo->GetText() == afterElement,
            "native editable combo pointer hit-test does not split a regional-indicator flag text element");

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(combo, state), "native editable combo session state is readable after emoji pointer insertion");
    const size_t insertionIndex = combo->GetText().find(L'X');
    Require(insertionIndex != std::wstring::npos, "native editable combo emoji pointer insertion writes the test character");
    Require(state.text == combo->GetText(), "native editable combo session state mirrors emoji pointer insertion text");
    Require(state.caretIndex == insertionIndex + 1u, "native editable combo session state mirrors emoji pointer insertion caret");
}

void TestNativeTextInputBackendNoSelectionCopyLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    const bool copied = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha");
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);
        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native no-selection ctrl+c");

        SendNativeCtrlKey(window.Hwnd(), 'C');

        const std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == L"sentinel" && field->GetText() == L"alpha" && ! field->GetSelectionRange().has_value();
    });
    Require(copied, "native ctrl+c without selection leaves clipboard and text unchanged");
}

void TestNativeTextInputBackendCtrlCopyCutPasteSyncsState()
{
    using namespace RedSalamander::DxUi;

    const bool edited = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha beta");
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);
        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"stale"), "clipboard initialized before native ctrl+c");

        SendNativeCtrlKey(window.Hwnd(), 'A');
        const std::optional<std::pair<size_t, size_t>> selectedAll = field->GetSelectionRange();
        if (! selectedAll || selectedAll.value().first != 0u || selectedAll.value().second != field->GetText().size())
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'C');
        std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"alpha beta")
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'X');
        clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"alpha beta" || ! field->GetText().empty())
        {
            return false;
        }

        NativeTextInputState state;
        if (! window.Host().TryReadNativeTextInputState(field, state) || ! state.text.empty() || state.caretIndex != 0u ||
            state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"omega\r\npsi\tzeta"), "clipboard initialized before native ctrl+v");
        SendNativeCtrlKey(window.Hwnd(), 'V');

        if (field->GetText() != L"omegapsizeta")
        {
            return false;
        }
        if (! window.Host().TryReadNativeTextInputState(field, state))
        {
            return false;
        }
        return state.text == L"omegapsizeta" && state.caretIndex == field->GetText().size() && ! state.selectionAnchorIndex.has_value();
    });
    Require(edited, "native ctrl+a/c/x/v copies, cuts, strips single-line paste control characters, and syncs session state");
}

void TestNativeTextInputBackendUndoRedoAndRedoClear()
{
    using namespace RedSalamander::DxUi;

    const bool edited = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha");
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeCtrlKey(window.Hwnd(), 'A');
        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"beta"), "clipboard initialized before native replacement paste");
        SendNativeCtrlKey(window.Hwnd(), 'V');
        if (field->GetText() != L"beta")
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Z');
        NativeTextInputState state;
        if (field->GetText() != L"alpha" || ! window.Host().TryReadNativeTextInputState(field, state) || state.text != L"alpha")
        {
            return false;
        }
        const std::optional<std::pair<size_t, size_t>> restoredSelection = field->GetSelectionRange();
        if (! restoredSelection || restoredSelection.value().first != 0u || restoredSelection.value().second != field->GetText().size())
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Y');
        if (field->GetText() != L"beta" || ! window.Host().TryReadNativeTextInputState(field, state) || state.text != L"beta")
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Z');
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'Z', 0));
        if (field->GetText() != L"Z")
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Y');
        if (field->GetText() != L"Z" || ! window.Host().TryReadNativeTextInputState(field, state))
        {
            return false;
        }
        return state.text == L"Z" && state.caretIndex == 1u && ! state.selectionAnchorIndex.has_value();
    });
    Require(edited, "native ctrl+z/y restores edit transactions and fresh edits clear redo history");
}

void TestNativeTextInputBackendEditTransactionsNotifyOnceAndIgnoreNoOps()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    uint32_t changeCount = 0u;
    std::wstring lastChangedText;
    field->SetOnTextChanged([&](std::wstring_view text)
    {
        ++changeCount;
        lastChangedText.assign(text);
    });

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_DELETE, 0));
    Require(changeCount == 0u && field->GetText() == L"alpha", "native no-op delete at end does not notify or mutate");

    SendNativeCtrlKey(window.Hwnd(), 'Z');
    Require(changeCount == 0u && field->GetText() == L"alpha", "native no-op delete does not create an undo transaction");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'Z', 0));
    NativeTextInputState state;
    Require(changeCount == 1u && lastChangedText == L"alphaZ", "native character edit notifies exactly once");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native edit transaction test reads state after char");
    Require(state.text == L"alphaZ" && state.caretIndex == 6u && ! state.selectionAnchorIndex.has_value(),
            "native character edit preserves collapsed caret after mutation");

    SendNativeCtrlKey(window.Hwnd(), 'Z');
    Require(changeCount == 2u && lastChangedText == L"alpha", "native undo notifies exactly once");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native edit transaction test reads state after undo");
    Require(state.text == L"alpha" && state.caretIndex == 5u && ! state.selectionAnchorIndex.has_value(), "native undo restores text and collapsed caret");

    SendNativeCtrlKey(window.Hwnd(), 'Y');
    Require(changeCount == 3u && lastChangedText == L"alphaZ", "native redo notifies exactly once");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native edit transaction test reads state after redo");
    Require(state.text == L"alphaZ" && state.caretIndex == 6u && ! state.selectionAnchorIndex.has_value(), "native redo restores text and collapsed caret");
}

void TestNativeTextInputBackendEmojiClipboardReplacementUndoRedo()
{
    using namespace RedSalamander::DxUi;

    const std::wstring emojiText    = MakeGrinningFaceTextElement() + MakeWomanTechnologistTextElement() + MakeRainbowFlagTextElement() +
                                      MakeThumbsUpMediumSkinToneTextElement() + MakeUsFlagTextElement();
    const std::wstring originalText = L"left " + emojiText + L" right";
    const std::wstring cutText      = L"left  right";
    const std::wstring replacement  = MakeRainbowFlagTextElement() + MakeGrinningFaceTextElement();
    const std::wstring replacedText = L"left " + replacement + L" right";

    const bool edited = RetryClipboardSensitiveAction([&]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(originalText);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        const size_t emojiStart = std::wstring_view(L"left ").size();
        const size_t emojiEnd   = emojiStart + emojiText.size();
        field->SetSelectionRange(emojiStart, emojiEnd);
        window.Host().SyncTextInput(field);

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native emoji copy");
        SendNativeCtrlKey(window.Hwnd(), 'C');
        std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText.has_value() || clipboardText.value() != emojiText)
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'X');
        clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText.has_value() || clipboardText.value() != emojiText || field->GetText() != cutText)
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Z');
        const std::optional<std::pair<size_t, size_t>> restoredSelection = field->GetSelectionRange();
        if (field->GetText() != originalText || ! restoredSelection.has_value() || restoredSelection.value().first != emojiStart ||
            restoredSelection.value().second != emojiEnd)
        {
            return false;
        }

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), replacement), "clipboard initialized before native emoji replacement paste");
        SendNativeCtrlKey(window.Hwnd(), 'V');
        if (field->GetText() != replacedText)
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Z');
        if (field->GetText() != originalText)
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Y');
        NativeTextInputState state;
        if (field->GetText() != replacedText || ! window.Host().TryReadNativeTextInputState(field, state))
        {
            return false;
        }
        if (state.text != replacedText || state.caretIndex != (emojiStart + replacement.size()) || state.selectionAnchorIndex.has_value())
        {
            return false;
        }
        return true;
    });
    Require(edited, "native emoji edit transactions copy/cut/paste, replace selections, round-trip clipboard text, and undo/redo");
}

void TestNativeTextInputBackendMaskedHiddenSuppressesCopyAndCut()
{
    using namespace RedSalamander::DxUi;

    const bool suppressed = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"secret");
        field->SetMasked(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);
        SendNativeCtrlKey(window.Hwnd(), 'A');

        NativeTextInputState state;
        if (! window.Host().TryReadNativeTextInputState(field, state) || ! state.masked)
        {
            return false;
        }

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native masked copy");
        SendNativeCtrlKey(window.Hwnd(), 'C');
        std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"sentinel")
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'X');
        clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"sentinel" || field->GetText() != L"secret")
        {
            return false;
        }

        const std::optional<std::pair<size_t, size_t>> selection = field->GetSelectionRange();
        return selection.has_value() && selection.value().first == 0u && selection.value().second == field->GetText().size();
    });
    Require(suppressed, "native hidden masked text suppresses copy and cut without disclosing or mutating the secret");
}

void TestNativeTextInputBackendMaskedExactPolicyCountsTextElements()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    std::wstring secret = L"A";
    secret.append(MakeWomanTechnologistTextElement());
    secret.append(MakeThumbsUpMediumSkinToneTextElement());
    secret.append(MakeUsFlagTextElement());

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(secret);
    field->SetMasked(true);
    field->SetPasswordMaskLengthPolicy(PasswordMaskLengthPolicy::Exact);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native masked exact policy exposes readable state");
    Require(state.masked, "native masked exact policy keeps the field masked");
    Require(state.maskLengthPolicy == PasswordMaskLengthPolicy::Exact, "native masked exact policy records exact masking");
    Require(state.secretVisibleDotCount == 4u, "native masked exact policy displays one dot per user-perceived text element");
}

void TestNativeTextInputBackendMaskedConcealedPolicyUsesStableBuckets()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"abc");
    field->SetMasked(true);
    field->SetPasswordMaskLengthPolicy(PasswordMaskLengthPolicy::Concealed);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native masked concealed policy exposes readable state");
    Require(state.maskLengthPolicy == PasswordMaskLengthPolicy::Concealed, "native masked concealed policy records concealed masking");
    const size_t firstBucketDots = state.secretVisibleDotCount;
    Require(firstBucketDots >= 4u && firstBucketDots <= 7u, "native masked concealed policy buckets a three-character secret into the first display range");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'd', 0));
    Require(window.Host().TryReadNativeTextInputState(field, state), "native masked concealed policy syncs after same-bucket edit");
    Require(field->GetText() == L"abcd", "native masked concealed policy accepts character input into retained secret text");
    Require(state.secretVisibleDotCount == firstBucketDots, "native masked concealed policy stays stable inside the same length bucket");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'e', 0));
    Require(window.Host().TryReadNativeTextInputState(field, state), "native masked concealed policy syncs after bucket-crossing edit");
    Require(field->GetText() == L"abcde", "native masked concealed policy keeps accepted character input masked in retained state");
    Require(state.secretVisibleDotCount >= 8u && state.secretVisibleDotCount <= 11u,
            "native masked concealed policy updates visible dots into the next display range only at a bucket boundary");
}

void TestNativeTextInputBackendMaskedConcealedPolicyRegeneratesEpochs()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"ab");
    field->SetMasked(true);
    field->SetPasswordMaskLengthPolicy(PasswordMaskLengthPolicy::Concealed);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 28.0f));

    auto* focusTarget = root->AddChild<Button>(L"next");
    focusTarget->SetBounds(D2D1::RectF(0.0f, 40.0f, 80.0f, 68.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native concealed epoch test reads initial state");
    const size_t firstEpochDots = state.secretVisibleDotCount;
    Require(firstEpochDots >= 4u && firstEpochDots <= 7u, "native concealed epoch dots stay inside the first privacy bucket display range");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'c', 0));
    Require(window.Host().TryReadNativeTextInputState(field, state), "native concealed epoch test reads same-bucket edit state");
    Require(state.secretVisibleDotCount == firstEpochDots, "native concealed dot count stays stable during same-bucket typing");

    field->SetText(L"xy");
    window.Host().SyncTextInput(field);
    Require(window.Host().TryReadNativeTextInputState(field, state), "native concealed epoch test reads full-reset state");
    const size_t resetEpochDots = state.secretVisibleDotCount;
    Require(resetEpochDots >= 4u && resetEpochDots <= 7u, "native concealed reset dots stay inside the first privacy bucket display range");
    Require(resetEpochDots != firstEpochDots, "native concealed dot count regenerates after an explicit full reset");

    window.Host().SetFocusControl(focusTarget);
    window.Host().SetFocusControl(field);
    Require(window.Host().TryReadNativeTextInputState(field, state), "native concealed epoch test reads refocused state");
    const size_t refocusEpochDots = state.secretVisibleDotCount;
    Require(refocusEpochDots >= 4u && refocusEpochDots <= 7u, "native concealed refocus dots stay inside the first privacy bucket display range");
    Require(refocusEpochDots != resetEpochDots, "native concealed dot count regenerates on a new focus epoch");
}

void TestNativeTextInputDeactivateSecureClearsCachedText()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.NativeTextInput.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "NativeTextInput source is readable for secure cache teardown guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t helperStart = source.find("void WindowHost::SecureClearNativeTextInputStateCache");
    Require(helperStart != std::string::npos, "WindowHost has a dedicated native text input cache secure-clear helper");
    const size_t helperEnd = source.find("\nvoid WindowHost::", helperStart + 1u);
    Require(helperEnd != std::string::npos && helperStart < helperEnd, "native text input cache secure-clear helper block is found");
    const std::string helperBlock = source.substr(helperStart, helperEnd - helperStart);
    Require(helperBlock.find("SecureWipe::SecureClear(_nativeTextInputStateCache.text)") != std::string::npos,
            "native text input cache secure-clear helper wipes the cached text copy");

    const size_t deactivateStart  = source.find("void WindowHost::DeactivateNativeTextInputSession");
    const size_t activateTsfStart = source.find("bool WindowHost::ActivateNativeTextInputTsf", deactivateStart);
    Require(deactivateStart != std::string::npos && activateTsfStart != std::string::npos && deactivateStart < activateTsfStart,
            "native text input deactivation block is found");
    const std::string deactivateBlock = source.substr(deactivateStart, activateTsfStart - deactivateStart);
    Require(deactivateBlock.find("SecureClearNativeTextInputStateCache()") != std::string::npos,
            "native text input deactivation secure-clears cached text when ending a session");
    Require(deactivateBlock.find("_nativeTextInputStateCacheValid = false") == std::string::npos,
            "native text input deactivation goes through the secure cache-clear helper instead of only invalidating the cache");
}

void TestNativeTextInputBackendConcealedEditingAffordancesAndPointerPolicy()
{
    using namespace RedSalamander::DxUi;

    const bool edited = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"secret");
        field->SetMasked(true);
        field->SetPasswordMaskLengthPolicy(PasswordMaskLengthPolicy::Concealed);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        NativeTextInputState state;
        if (! window.Host().TryReadNativeTextInputState(field, state) || state.secretVisibleDotCount < 8u || state.secretVisibleDotCount > 11u)
        {
            return false;
        }

        SendNativeClick(window.Hwnd(), POINT{12, 14});
        if (! window.Host().TryReadNativeTextInputState(field, state) || state.caretIndex != field->GetText().size() || state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        SendNativeKey(window.Hwnd(), VK_BACK);
        if (field->GetText() != L"secre")
        {
            return false;
        }

        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'X', 0));
        if (field->GetText() != L"secreX")
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'A');
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'Z', 0));
        if (field->GetText() != L"Z")
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Z');
        if (field->GetText() != L"secreX")
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'Y');
        if (field->GetText() != L"Z")
        {
            return false;
        }

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"paste"), "clipboard initialized before concealed paste");
        SendNativeCtrlKey(window.Hwnd(), 'A');
        SendNativeCtrlKey(window.Hwnd(), 'V');
        return field->GetText() == L"paste" && window.Host().TryReadNativeTextInputState(field, state) &&
               state.passwordRevealState == PasswordRevealState::Hidden && state.secretVisibleDotCount >= 8u && state.secretVisibleDotCount <= 11u;
    });
    Require(edited, "native concealed mode supports keyboard edits while pointer placement snaps to the logical end");
}

void TestNativeTextInputBackendRevealedMaskedFieldAllowsCopyAndCut()
{
    using namespace RedSalamander::DxUi;

    const bool revealedCut = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"secret");
        field->SetMasked(true);
        field->SetPasswordRevealState(PasswordRevealState::Visible);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);
        SendNativeCtrlKey(window.Hwnd(), 'A');

        NativeTextInputState state;
        if (! window.Host().TryReadNativeTextInputState(field, state) || ! state.masked || state.passwordRevealState != PasswordRevealState::Visible ||
            state.secretVisibleDotCount != 0u)
        {
            return false;
        }

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native revealed masked copy");
        SendNativeCtrlKey(window.Hwnd(), 'C');
        std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"secret")
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'X');
        clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == L"secret" && field->GetText().empty() &&
               window.Host().TryReadNativeTextInputState(field, state) && state.text.empty() && state.passwordRevealState == PasswordRevealState::Visible;
    });
    Require(revealedCut, "native revealed masked text allows copy and cut while preserving revealed state until remask");
}

void TestNativeTextInputBackendRevealedMaskedFieldRemasksOnBlurReadOnlyAndDisable()
{
    using namespace RedSalamander::DxUi;

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"secret");
        field->SetMasked(true);
        field->SetPasswordRevealState(PasswordRevealState::Visible);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);
        Require(field->GetSecretVisibleDotCount() == 0u, "visible masked field exposes plaintext without dots before blur");

        window.Host().SetFocusControl(nullptr);
        Require(field->GetPasswordRevealState() == PasswordRevealState::Hidden, "blur remasks a visible masked field");
        Require(field->GetSecretVisibleDotCount() == 6u, "blur restores exact masked dot count");
    }

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"secret");
        field->SetMasked(true);
        field->SetPasswordRevealState(PasswordRevealState::Visible);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        field->SetReadOnly(true);
        Require(field->GetPasswordRevealState() == PasswordRevealState::Hidden, "read-only transition remasks a visible masked field");
        Require(field->GetSecretVisibleDotCount() == 6u, "read-only transition restores exact masked dot count");
    }

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"secret");
        field->SetMasked(true);
        field->SetPasswordRevealState(PasswordRevealState::Visible);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        field->SetEnabled(false);
        Require(field->GetPasswordRevealState() == PasswordRevealState::Hidden, "disabled transition remasks a visible masked field");
        Require(field->GetSecretVisibleDotCount() == 6u, "disabled transition restores exact masked dot count");
    }
}

void TestNativeTextInputBackendImeCompositionClearsOnWindowDeactivate()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputState state{};
    Require(window.Host().TryReadNativeTextInputState(field, state) && state.compositionStartIndex.has_value(),
            "native ime deactivation test starts with an active composition");

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_ACTIVATE, MAKEWPARAM(WA_INACTIVE, 0), 0, handled));

    Require(handled, "native ime composition handles window deactivation");
    Require(window.Host().GetFocusControl() == field, "native ime window deactivation keeps retained logical focus");
    Require(! field->HasFocus(), "native ime window deactivation clears active focus visuals");
    Require(! window.Host().DebugHasActiveNativeTextInputSession(), "native ime window deactivation clears the native session");
    Require(! window.Host().TryReadNativeTextInputState(field, state), "native ime window deactivation clears readable session state");

    window.Host().SetFocusControl(field);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));
    Require(window.Host().TryReadNativeTextInputState(field, state) && state.compositionStartIndex.has_value(),
            "native ime deactivation test restarts composition before app deactivation");

    handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_ACTIVATEAPP, FALSE, 0, handled));

    Require(handled, "native ime composition handles app deactivation");
    Require(window.Host().GetFocusControl() == field, "native ime app deactivation keeps retained logical focus");
    Require(! field->HasFocus(), "native ime app deactivation clears active focus visuals");
    Require(! window.Host().DebugHasActiveNativeTextInputSession(), "native ime app deactivation clears the native session");
    Require(! window.Host().TryReadNativeTextInputState(field, state), "native ime app deactivation clears readable session state");
}

void TestNativeTextInputBackendRevealedMaskedFieldRemasksOnWindowDeactivate()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"secret");
    field->SetMasked(true);
    field->SetPasswordRevealState(PasswordRevealState::Visible);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_ACTIVATE, MAKEWPARAM(WA_INACTIVE, 0), 0, handled));

    Require(handled, "native masked field handles window deactivation");
    Require(field->GetPasswordRevealState() == PasswordRevealState::Hidden, "window deactivation remasks a visible masked field");

    NativeTextInputState state{};
    Require(! window.Host().TryReadNativeTextInputState(field, state), "window deactivation clears the native text input session");
    Require(field->GetSecretVisibleDotCount() == 6u, "window deactivation restores exact masked dot count");

    field->SetPasswordRevealState(PasswordRevealState::Visible);
    window.Host().SetFocusControl(field);
    Require(window.Host().TryReadNativeTextInputState(field, state), "native masked field reactivates before app deactivation");

    handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_ACTIVATEAPP, FALSE, 0, handled));

    Require(handled, "native masked field handles app deactivation");
    Require(field->GetPasswordRevealState() == PasswordRevealState::Hidden, "app deactivation remasks a visible masked field");
    Require(! window.Host().TryReadNativeTextInputState(field, state), "app deactivation clears the native text input session");
    Require(field->GetSecretVisibleDotCount() == 6u, "app deactivation restores exact masked dot count");
}

void TestNativeTextInputBackendMaskedRevealButtonRemasksOnCaptureLoss()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"secret");
    field->SetMasked(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    const LPARAM revealButtonPoint = MAKELPARAM(170, 14);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, revealButtonPoint));

    NativeTextInputState state{};
    Require(window.Host().TryReadNativeTextInputState(field, state), "native masked reveal capture-loss test reads state after press");
    Require(state.passwordRevealState == PasswordRevealState::Visible && state.secretVisibleDotCount == 0u,
            "masked reveal button press exposes plaintext before capture loss");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CAPTURECHANGED, 0, 0));

    Require(field->GetPasswordRevealState() == PasswordRevealState::Hidden, "capture loss remasks a pressed reveal button");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native masked reveal capture-loss test reads state after capture loss");
    Require(state.passwordRevealState == PasswordRevealState::Hidden && state.secretVisibleDotCount == 6u, "capture loss syncs hidden masked state");
}

void TestNativeTextInputBackendMaskedRevealButtonPeeksWithoutClearingSecret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"secret");
    field->SetMasked(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native masked reveal button test starts with readable state");
    Require(state.passwordRevealState == PasswordRevealState::Hidden && state.secretVisibleDotCount == 6u, "native masked reveal button test starts hidden");

    const D2D1_POINT_2F revealButtonPoint = D2D1::Point2F(170.0f, 14.0f);
    Require(field->OnMouseDown(window.Host(), revealButtonPoint, false, 0), "masked reveal button press is handled");
    Require(field->GetText() == L"secret", "masked reveal button press does not clear the secret");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native masked reveal button press syncs state");
    Require(state.passwordRevealState == PasswordRevealState::Visible && state.secretVisibleDotCount == 0u, "masked reveal button press peeks plaintext");

    Require(field->OnMouseUp(window.Host(), revealButtonPoint, false, 0), "masked reveal button release is handled");
    Require(field->GetText() == L"secret", "masked reveal button release preserves the secret");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native masked reveal button release syncs state");
    Require(state.passwordRevealState == PasswordRevealState::Hidden && state.secretVisibleDotCount == 6u, "masked reveal button release remasks plaintext");
}

void TestNativeTextInputBackendMaskedRevealButtonSupportsKeyboardPeek()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"secret");
    field->SetMasked(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    auto* next = root->AddChild<Button>(L"next");
    next->SetBounds(D2D1::RectF(0.0f, 40.0f, 80.0f, 68.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == field, "tab reaches the masked reveal affordance before leaving the field");

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native masked keyboard reveal test starts with readable state");
    Require(state.passwordRevealState == PasswordRevealState::Hidden && state.secretVisibleDotCount == 6u,
            "tab focus on reveal affordance does not disclose the secret");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_SPACE, 0));
    Require(window.Host().TryReadNativeTextInputState(field, state), "space press syncs native masked keyboard reveal state");
    Require(state.passwordRevealState == PasswordRevealState::Visible && state.secretVisibleDotCount == 0u,
            "space press peeks the keyboard-focused reveal affordance");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYUP, VK_SPACE, 0));
    Require(window.Host().TryReadNativeTextInputState(field, state), "space release syncs native masked keyboard reveal state");
    Require(state.passwordRevealState == PasswordRevealState::Hidden && state.secretVisibleDotCount == 6u,
            "space release remasks the keyboard-focused reveal affordance");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_RETURN, 0));
    Require(window.Host().TryReadNativeTextInputState(field, state), "enter press syncs native masked keyboard reveal state");
    Require(state.passwordRevealState == PasswordRevealState::Visible && state.secretVisibleDotCount == 0u,
            "enter press peeks the keyboard-focused reveal affordance");

    window.Host().SetFocusControl(next);
    Require(field->GetPasswordRevealState() == PasswordRevealState::Hidden, "blur remasks a keyboard-held reveal affordance");
    window.Host().SetFocusControl(field);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == next, "second tab leaves the reveal affordance for the next control");
}

void TestNativeTextInputBackendPasswordRevealModesControlAffordanceAndVisibility()
{
    using namespace RedSalamander::DxUi;

    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"secret");
        field->SetMasked(true);
        field->SetPasswordRevealMode(PasswordRevealMode::Hidden);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        auto* next = root->AddChild<Button>(L"next");
        next->SetBounds(D2D1::RectF(0.0f, 40.0f, 80.0f, 68.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        NativeTextInputState state;
        Require(window.Host().TryReadNativeTextInputState(field, state), "hidden reveal mode exposes initial native state");
        Require(state.passwordRevealState == PasswordRevealState::Hidden && state.secretVisibleDotCount == 6u, "hidden reveal mode starts masked");

        const D2D1_POINT_2F rightEdgePoint = D2D1::Point2F(170.0f, 14.0f);
        Require(field->OnMouseDown(window.Host(), rightEdgePoint, false, 0), "hidden reveal mode still handles a text-field click");
        Require(window.Host().TryReadNativeTextInputState(field, state), "hidden reveal mode syncs after right-edge click");
        Require(state.passwordRevealState == PasswordRevealState::Hidden && state.secretVisibleDotCount == 6u,
                "hidden reveal mode does not expose a pointer reveal affordance");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_TAB, 0));
        Require(window.Host().GetFocusControl() == next, "hidden reveal mode skips reveal affordance keyboard focus");
    }

    const bool visibleModeCopied = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"secret");
        field->SetMasked(true);
        field->SetPasswordRevealMode(PasswordRevealMode::Visible);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        auto* next = root->AddChild<Button>(L"next");
        next->SetBounds(D2D1::RectF(0.0f, 40.0f, 80.0f, 68.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        NativeTextInputState state;
        if (field->GetPasswordRevealMode() != PasswordRevealMode::Visible || field->GetPasswordRevealState() != PasswordRevealState::Visible ||
            field->GetSecretVisibleDotCount() != 0u || ! window.Host().TryReadNativeTextInputState(field, state) ||
            state.passwordRevealState != PasswordRevealState::Visible || state.secretVisibleDotCount != 0u)
        {
            return false;
        }

        SendNativeCtrlKey(window.Hwnd(), 'A');
        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before visible reveal mode copy");
        SendNativeCtrlKey(window.Hwnd(), 'C');
        const std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText.has_value() || clipboardText.value() != L"secret")
        {
            return false;
        }

        window.Host().SetFocusControl(next);
        if (field->GetPasswordRevealState() != PasswordRevealState::Visible || field->GetSecretVisibleDotCount() != 0u)
        {
            return false;
        }

        field->SetReadOnly(true);
        field->SetEnabled(false);
        return field->GetPasswordRevealState() == PasswordRevealState::Visible && field->GetSecretVisibleDotCount() == 0u;
    });
    Require(visibleModeCopied, "visible reveal mode keeps plaintext visible and copy-enabled across blur/read-only/disabled transitions");
}

void TestNativeTextInputBackendReadOnlyAllowsCopyAndSuppressesMutation()
{
    using namespace RedSalamander::DxUi;

    const bool suppressed = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha");
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);
        SendNativeCtrlKey(window.Hwnd(), 'A');

        NativeTextInputState state;
        if (! window.Host().TryReadNativeTextInputState(field, state) || ! state.readOnly)
        {
            return false;
        }

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native read-only copy");
        SendNativeCtrlKey(window.Hwnd(), 'C');
        std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"alpha")
        {
            return false;
        }

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"replacement"), "clipboard initialized before native read-only mutation attempts");
        SendNativeCtrlKey(window.Hwnd(), 'V');
        SendNativeCtrlKey(window.Hwnd(), 'X');
        SendNativeShiftKey(window.Hwnd(), VK_DELETE);
        SendNativeKey(window.Hwnd(), VK_BACK);
        SendNativeKey(window.Hwnd(), VK_DELETE);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'Z', 0));

        clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"replacement" || field->GetText() != L"alpha")
        {
            return false;
        }

        if (! window.Host().TryReadNativeTextInputState(field, state))
        {
            return false;
        }
        const std::optional<std::pair<size_t, size_t>> selection = field->GetSelectionRange();
        return state.text == L"alpha" && state.readOnly && selection.has_value() && selection.value().first == 0u &&
               selection.value().second == field->GetText().size();
    });
    Require(suppressed, "native read-only text allows copy and suppresses paste, cut, delete, and character mutation");
}

void TestNativeTextInputBackendMultilineCtrlCopyPastePreservesLogicalNewlines()
{
    using namespace RedSalamander::DxUi;

    const bool edited = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha\nbeta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        SendNativeCtrlKey(window.Hwnd(), 'A');
        SendNativeCtrlKey(window.Hwnd(), 'C');
        std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"alpha\nbeta")
        {
            return false;
        }

        field->SetSelectionRange(5u, 5u);
        window.Host().SyncTextInput(field);
        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"ONE\r\nTWO"), "clipboard initialized before native multiline ctrl+v");
        SendNativeCtrlKey(window.Hwnd(), 'V');

        NativeTextInputState state;
        if (! window.Host().TryReadNativeTextInputState(field, state))
        {
            return false;
        }
        return field->GetText() == L"alphaONE\nTWO\nbeta" && state.text == field->GetText() && state.caretIndex == 12u &&
               ! state.selectionAnchorIndex.has_value();
    });
    Require(edited, "native multiline ctrl+c/ctrl+v preserves logical LF text and normalizes pasted CRLF to LF");
}

void TestNativeTextInputBackendMultilineCharAndReturnReplacementSyncState()
{
    using namespace RedSalamander::DxUi;

    const auto runCase = [](bool wrapped)
    {
        const std::wstring initialText = wrapped ? std::wstring(kWrappedMultilineClipboardTextForTest) : std::wstring(kLogicalNewlineClipboardTextForTest);
        const size_t selectionStart    = wrapped ? kWrappedMultilineClipboardSelectionStartForTest : kLogicalNewlineClipboardSelectionStartForTest;
        const size_t selectionEnd      = wrapped ? kWrappedMultilineClipboardSelectionEndForTest : kLogicalNewlineClipboardSelectionEndForTest;

        const std::wstring charExpected   = initialText.substr(0u, selectionStart) + L"Z" + initialText.substr(selectionEnd);
        const std::wstring returnExpected = initialText.substr(0u, selectionStart) + L"\n" + initialText.substr(selectionEnd);

        {
            AttachedHostWindow window;
            window.Host().SetTextInputBackend(TextInputBackend::Native);

            auto root   = std::make_unique<Panel>();
            auto* field = root->AddChild<TextField>(initialText);
            field->SetMultiline(true);
            field->SetBounds(wrapped ? D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f) : D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

            window.Host().SetRoot(std::move(root));
            window.Host().SetFocusControl(field);
            field->SetSelectionRange(selectionStart, selectionEnd);
            window.Host().SyncTextInput(field);

            static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, static_cast<WPARAM>(L'Z'), 0));

            NativeTextInputState state{};
            Require(field->GetText() == charExpected,
                    wrapped ? "native wrapped multiline wm_char replaces the selected visible range"
                            : "native multiline wm_char replaces the selected logical newline-spanning range");
            Require(window.Host().TryReadNativeTextInputState(field, state),
                    wrapped ? "native wrapped multiline wm_char replacement syncs state" : "native multiline wm_char replacement syncs state");
            Require(state.text == field->GetText(),
                    wrapped ? "native wrapped multiline wm_char replacement keeps session text in sync"
                            : "native multiline wm_char replacement keeps session text in sync");
            Require(state.caretIndex == selectionStart + 1u && ! state.selectionAnchorIndex.has_value(),
                    wrapped ? "native wrapped multiline wm_char replacement leaves a collapsed caret after the inserted text"
                            : "native multiline wm_char replacement leaves a collapsed caret after the inserted text");
        }

        {
            AttachedHostWindow window;
            window.Host().SetTextInputBackend(TextInputBackend::Native);

            auto root    = std::make_unique<Panel>();
            auto* field  = root->AddChild<TextField>(initialText);
            auto* button = root->AddChild<Button>(L"Apply");
            field->SetMultiline(true);
            field->SetBounds(wrapped ? D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f) : D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
            button->SetBounds(wrapped ? D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f) : D2D1::RectF(0.0f, 132.0f, 120.0f, 160.0f));

            size_t invokeCount = 0u;
            button->SetOnClick([&invokeCount] { ++invokeCount; });

            window.Host().SetRoot(std::move(root));
            window.Host().SetDefaultButton(button);
            window.Host().SetFocusControl(field);

            field->SetSelectionRange(initialText.size(), initialText.size());
            window.Host().SyncTextInput(field);
            static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_RETURN, 0));
            Require(invokeCount == 0u,
                    wrapped ? "native wrapped multiline return does not invoke the default button"
                            : "native multiline return does not invoke the default button");
            static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, static_cast<WPARAM>(L'\r'), 0));
            Require(field->GetText() == initialText + L"\n",
                    wrapped ? "native wrapped multiline return inserts a logical newline at a collapsed caret"
                            : "native multiline return inserts a logical newline at a collapsed caret");

            field->SetText(initialText);
            field->SetSelectionRange(selectionStart, selectionEnd);
            window.Host().SyncTextInput(field);
            static_cast<void>(SendMessageW(window.Hwnd(), WM_KEYDOWN, VK_RETURN, 0));
            Require(invokeCount == 0u,
                    wrapped ? "native wrapped multiline return replacement does not invoke the default button"
                            : "native multiline return replacement does not invoke the default button");
            static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, static_cast<WPARAM>(L'\r'), 0));

            NativeTextInputState state{};
            Require(field->GetText() == returnExpected,
                    wrapped ? "native wrapped multiline return replaces the selected visible range with LF"
                            : "native multiline return replaces the selected logical newline-spanning range with LF");
            Require(window.Host().TryReadNativeTextInputState(field, state),
                    wrapped ? "native wrapped multiline return replacement syncs state" : "native multiline return replacement syncs state");
            Require(state.text == field->GetText() && state.caretIndex == selectionStart + 1u && ! state.selectionAnchorIndex.has_value(),
                    wrapped ? "native wrapped multiline return replacement leaves a collapsed caret after LF"
                            : "native multiline return replacement leaves a collapsed caret after LF");
        }
    };

    runCase(false);
    runCase(true);
}

void TestNativeTextInputBackendEditMessagesCopyPasteCutClearSelection()
{
    using namespace RedSalamander::DxUi;

    const bool edited = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 140.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        field->SetSelectionRange(0u, 10u);
        window.Host().SyncTextInput(field);
        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native wm_copy");
        static_cast<void>(SendMessageW(window.Hwnd(), WM_COPY, 0, 0));
        std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"alpha\nbeta")
        {
            return false;
        }

        field->SetSelectionRange(5u, 10u);
        window.Host().SyncTextInput(field);
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CLEAR, 0, 0));
        NativeTextInputState state;
        if (field->GetText() != L"alpha\ncharlie" || ! window.Host().TryReadNativeTextInputState(field, state) || state.text != field->GetText() ||
            state.caretIndex != 5u || state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"ONE\r\nTWO"), "clipboard initialized before native wm_paste");
        static_cast<void>(SendMessageW(window.Hwnd(), WM_PASTE, 0, 0));
        if (field->GetText() != L"alphaONE\nTWO\ncharlie")
        {
            return false;
        }

        field->SetSelectionRange(0u, 5u);
        window.Host().SyncTextInput(field);
        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native wm_cut");
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CUT, 0, 0));
        clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"alpha" || field->GetText() != L"ONE\nTWO\ncharlie")
        {
            return false;
        }

        static_cast<void>(SendMessageW(window.Hwnd(), WM_UNDO, 0, 0));
        return field->GetText() == L"alphaONE\nTWO\ncharlie" && window.Host().TryReadNativeTextInputState(field, state) && state.text == field->GetText();
    });
    Require(edited, "native edit messages copy, paste, cut, clear selection, and undo through the host-owned text session");
}

void TestNativeTextInputBackendEditMessagesRoundTripWin32Protocol()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta \xD83D\xDE00");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 140.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    const std::wstring expectedWin32Text = L"alpha\r\nbeta \xD83D\xDE00";
    Require(SendMessageW(window.Hwnd(), WM_GETTEXTLENGTH, 0, 0) == static_cast<LRESULT>(expectedWin32Text.size()),
            "native wm_gettextlength reports CRLF-normalized multiline text length");

    std::array<wchar_t, 64> buffer{};
    const LRESULT copied = SendMessageW(window.Hwnd(), WM_GETTEXT, buffer.size(), reinterpret_cast<LPARAM>(buffer.data()));
    Require(copied == static_cast<LRESULT>(expectedWin32Text.size()) && std::wstring(buffer.data()) == expectedWin32Text,
            "native wm_gettext returns CRLF-normalized multiline text and preserves surrogate pairs");

    const wchar_t replacementText[] = L"one\r\ntwo \xD83D\xDE00";
    Require(SendMessageW(window.Hwnd(), WM_SETTEXT, 0, reinterpret_cast<LPARAM>(replacementText)) == TRUE, "native wm_settext reports success");
    Require(field->GetText() == L"one\ntwo \xD83D\xDE00", "native wm_settext normalizes CRLF input to logical LF text");

    DWORD selectionStart = 0;
    DWORD selectionEnd   = 0;
    static_cast<void>(SendMessageW(window.Hwnd(), EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == 11u && selectionEnd == 11u, "native wm_settext leaves the Win32 selection collapsed at the logical end");

    Require(SendMessageW(window.Hwnd(), EM_SETSEL, 0, 3) != 0, "native em_setsel returns nonzero on success");
    Require(SendMessageW(window.Hwnd(), EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"ONE\r\nTWO")) == TRUE, "native em_replacesel reports success");
    Require(field->GetText() == L"ONE\nTWO\ntwo \xD83D\xDE00", "native em_replacesel normalizes CRLF replacement text");

    Require(SendMessageW(window.Hwnd(), EM_SETSEL, static_cast<WPARAM>(static_cast<UINT>(-1)), static_cast<LPARAM>(-1)) != 0,
            "native em_setsel accepts the 32-bit Win32 select-to-end sentinel");
    selectionStart = 0;
    selectionEnd   = 0;
    static_cast<void>(SendMessageW(window.Hwnd(), EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd && selectionEnd == 16u, "native em_setsel(-1, -1) collapses the Win32 selection at the end");
}

void TestNativeTextInputBackendEditMessagesSetTextClearsComposition()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));
    NativeTextInputState state{};
    Require(window.Host().TryReadNativeTextInputState(field, state) && state.compositionStartIndex.has_value(),
            "native edit-message wm_settext test starts with an active composition");

    Require(SendMessageW(window.Hwnd(), WM_SETTEXT, 0, reinterpret_cast<LPARAM>(L"omega\r\npsi")) == TRUE,
            "native wm_settext during composition reports success");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native wm_settext during composition leaves readable state");
    Require(field->GetText() == L"omega\r\npsi" || field->GetText() == L"omega\npsi", "native wm_settext during composition writes replacement text");
    Require(! state.compositionStartIndex.has_value() && ! state.compositionEndIndex.has_value(),
            "native wm_settext during composition clears the active composition range");
}

void TestNativeTextInputBackendEditMessagesFallBackWithoutTextInput()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    Require(SetWindowTextW(window.Hwnd(), L"fallback-title") != FALSE, "native edit-message fallback test sets a host window title");

    bool handled         = true;
    const LRESULT length = window.Host().HandleMessage(window.Hwnd(), WM_GETTEXTLENGTH, 0, 0, handled);
    Require(handled, "native edit-message shim dispatches default handling without a focused native text input");
    Require(length == 14, "native edit-message shim returns the default window-proc value when it declines a message");
}

void TestNativeTextInputBackendClearWithoutSelectionLeavesTextAndClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    const bool unchanged = RetryClipboardSensitiveAction([]()
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha\nbeta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 140.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);
        field->SetSelectionRange(5u, 5u);
        window.Host().SyncTextInput(field);
        Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before native wm_clear no-selection");

        static_cast<void>(SendMessageW(window.Hwnd(), WM_CLEAR, 0, 0));

        NativeTextInputState state;
        const std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return field->GetText() == L"alpha\nbeta" && clipboardText.has_value() && clipboardText.value() == L"sentinel" &&
               window.Host().TryReadNativeTextInputState(field, state) && state.text == field->GetText() && state.caretIndex == 5u &&
               ! state.selectionAnchorIndex.has_value();
    });
    Require(unchanged, "native wm_clear without selection leaves text and clipboard unchanged");
}

void TestNativeTextInputTextStoreRequiresLockAndExposesTextSelectionGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 260.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(0u, 5u);
    window.Host().SyncTextInput(field);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store is created for focused native text input");

    LONG endOutsideLock = 0;
    Require(store->GetEndACP(&endOutsideLock) == TS_E_NOLOCK, "native text store rejects GetEndACP without a document lock");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native text store accepts a sink");

    LONG endAcp = -1;
    TS_SELECTION_ACP selection{};
    ULONG fetchedSelection = 0u;
    std::wstring readText;
    ULONG fetchedChars      = 0u;
    ULONG fetchedRuns       = 0u;
    LONG nextAcp            = -1;
    TsViewCookie activeView = 0u;
    HWND textStoreHwnd      = nullptr;
    RECT screenExt{};
    RECT textExt{};
    BOOL textExtClipped = TRUE;
    LONG pointAcp       = -1;
    HRESULT callbackHr  = E_UNEXPECTED;

    sink.onLockGranted = [&](DWORD /*lockFlags*/) noexcept
    {
        callbackHr = store->GetEndACP(&endAcp);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        callbackHr = store->GetSelection(TS_DEFAULT_SELECTION, 1u, &selection, &fetchedSelection);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        wchar_t buffer[32]{};
        TS_RUNINFO runInfo{};
        callbackHr = store->GetText(0, -1, buffer, static_cast<ULONG>(std::size(buffer)), &fetchedChars, &runInfo, 1u, &fetchedRuns, &nextAcp);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }
        readText.assign(buffer, fetchedChars);

        callbackHr = store->GetActiveView(&activeView);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }
        callbackHr = store->GetWnd(activeView, &textStoreHwnd);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }
        callbackHr = store->GetScreenExt(activeView, &screenExt);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }
        callbackHr = store->GetTextExt(activeView, 0, 5, &textExt, &textExtClipped);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        POINT queryPoint{screenExt.left + 2, (screenExt.top + screenExt.bottom) / 2};
        callbackHr = store->GetACPFromPoint(activeView, &queryPoint, 0u, &pointAcp);
        return callbackHr;
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store read lock request dispatches to sink");
    RequireSucceeded(sessionHr, "native text store read lock callback succeeds");
    RequireSucceeded(callbackHr, "native text store read callback completed all queries");
    Require(sink.lastLockFlags == TS_LF_READ, "native text store grants a read lock");
    Require(endAcp == 10, "native text store reports the focused text end ACP");
    Require(fetchedSelection == 1u && selection.acpStart == 0 && selection.acpEnd == 5, "native text store exposes the retained selection range");
    Require(readText == L"alpha beta" && fetchedChars == 10u && fetchedRuns == 1u && nextAcp == 10, "native text store returns plain text and run info");
    Require(textStoreHwnd == window.Hwnd(), "native text store exposes the host HWND");
    Require(screenExt.right > screenExt.left && screenExt.bottom > screenExt.top, "native text store exposes a non-empty screen extent");
    Require(textExt.right > textExt.left && textExt.bottom > textExt.top && textExtClipped == FALSE, "native text store exposes a text extent");
    Require(pointAcp >= 0 && pointAcp <= endAcp, "native text store maps a screen point to an ACP");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store unadvises the sink");
}

void TestNativeTextInputTextStoreRejectsDestroyedControlDuringTeardown()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 260.0f, 44.0f));
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(CreateNativeTextInputTextStore(window.Host(), *field));
    Require(store != nullptr, "destroy-during-teardown test creates an independent native text store");

    TsViewCookie view = 0u;
    RequireSucceeded(store->GetActiveView(&view), "destroy-during-teardown test resolves the text-store view");
    RECT bounds{};
    RequireSucceeded(store->GetScreenExt(view, &bounds), "destroy-during-teardown test reads live control geometry");

    window.Host().SetRoot(std::make_unique<Panel>());

    Require(store->GetScreenExt(view, &bounds) == TS_E_INVALIDPOS,
            "native text store rejects a control destroyed during teardown instead of dereferencing stale state");
}

void TestNativeTextInputTextStoreExposesOwnerCompositionSink()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store composition sink test creates the text store");

    wil::com_ptr_nothrow<ITfContextOwnerCompositionSink> compositionSink;
    RequireSucceeded(store.query_to(compositionSink.put()), "native text store exposes ITfContextOwnerCompositionSink");

    BOOL compositionAccepted = FALSE;
    RequireSucceeded(compositionSink->OnStartComposition(nullptr, &compositionAccepted),
                     "native text store owner composition sink accepts composition start callbacks");
    Require(compositionAccepted == TRUE, "native text store owner composition sink allows TSF composition");
    RequireSucceeded(compositionSink->OnUpdateComposition(nullptr, nullptr), "native text store owner composition sink accepts composition update callbacks");
    RequireSucceeded(compositionSink->OnEndComposition(nullptr), "native text store owner composition sink accepts composition end callbacks");
}

void TestNativeTextInputTextStoreExposesAcp2Surface()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 260.0f, 44.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(6u, 10u);
    window.Host().SyncTextInput(field);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native ACP2 text store test creates the text store");

    wil::com_ptr_nothrow<ITextStoreACP2> store2;
    RequireSucceeded(store.query_to(store2.put()), "native text store exposes ITextStoreACP2");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store2->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native ACP2 text store accepts a sink");

    LONG endAcp = -1;
    TS_SELECTION_ACP selection{};
    ULONG fetchedSelection = 0u;
    std::wstring readText;
    ULONG fetchedChars      = 0u;
    ULONG fetchedRuns       = 0u;
    LONG nextAcp            = -1;
    TsViewCookie activeView = 0u;
    RECT screenExt{};
    RECT textExt{};
    BOOL textExtClipped = TRUE;
    LONG pointAcp       = -1;
    HRESULT callbackHr  = E_UNEXPECTED;

    sink.onLockGranted = [&](DWORD /*lockFlags*/) noexcept
    {
        callbackHr = store2->GetEndACP(&endAcp);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        callbackHr = store2->GetSelection(TS_DEFAULT_SELECTION, 1u, &selection, &fetchedSelection);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        wchar_t buffer[32]{};
        TS_RUNINFO runInfo{};
        callbackHr = store2->GetText(0, -1, buffer, static_cast<ULONG>(std::size(buffer)), &fetchedChars, &runInfo, 1u, &fetchedRuns, &nextAcp);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }
        readText.assign(buffer, fetchedChars);

        callbackHr = store2->GetActiveView(&activeView);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }
        callbackHr = store2->GetScreenExt(activeView, &screenExt);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }
        callbackHr = store2->GetTextExt(activeView, 6, 10, &textExt, &textExtClipped);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        POINT queryPoint{screenExt.left + 2, (screenExt.top + screenExt.bottom) / 2};
        callbackHr = store2->GetACPFromPoint(activeView, &queryPoint, 0u, &pointAcp);
        return callbackHr;
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store2->RequestLock(TS_LF_READ, &sessionHr), "native ACP2 text store read lock request dispatches to sink");
    RequireSucceeded(sessionHr, "native ACP2 text store read lock callback succeeds");
    RequireSucceeded(callbackHr, "native ACP2 text store read callback completed all queries");
    Require(endAcp == 10, "native ACP2 text store reports the focused text end ACP");
    Require(fetchedSelection == 1u && selection.acpStart == 6 && selection.acpEnd == 10, "native ACP2 text store exposes the retained selection range");
    Require(readText == L"alpha beta" && fetchedChars == 10u && fetchedRuns == 1u && nextAcp == 10, "native ACP2 text store returns plain text and run info");
    Require(screenExt.right > screenExt.left && screenExt.bottom > screenExt.top, "native ACP2 text store exposes a non-empty screen extent");
    Require(textExt.right > textExt.left && textExt.bottom > textExt.top && textExtClipped == FALSE, "native ACP2 text store exposes a text extent");
    Require(pointAcp >= 0 && pointAcp <= endAcp, "native ACP2 text store maps a screen point to an ACP");

    RequireSucceeded(store2->UnadviseSink(&sink), "native ACP2 text store unadvises the sink");
}

void TestNativeTextInputTextStoreMixedBiDiGeometryUsesTextViewport()
{
    using namespace RedSalamander::DxUi;

    constexpr std::wstring_view mixedBiDiText = L"abc \x05D0\x05D1\x05D2 123";
    constexpr std::array<DirectWritePointerSpan, 3u> spans{
        DirectWritePointerSpan{0u, 1u}, DirectWritePointerSpan{5u, 6u}, DirectWritePointerSpan{mixedBiDiText.size() - 1u, mixedBiDiText.size()}};

    for (const FlowDirection flowDirection : {FlowDirection::LeftToRight, FlowDirection::RightToLeft})
    {
        AttachedHostWindow window;
        window.Host().SetTextInputBackend(TextInputBackend::Native);

        auto root = std::make_unique<Panel>();
        root->SetFlowDirection(flowDirection);
        auto* field = root->AddChild<TextField>(std::wstring(mixedBiDiText));
        field->SetBounds(D2D1::RectF(20.0f, 18.0f, 380.0f, 50.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);
        field->SetSelectionRange(4u, 7u);
        window.Host().SyncTextInput(field);

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native text store mixed-BiDi geometry keeps the hidden bridge child absent");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native text store mixed-BiDi geometry keeps input on the host hwnd");

        TextFieldDebugSingleLinePaintState paint{};
        Require(field->DebugGetSingleLinePaintState(window.Host(), paint), "native text store mixed-BiDi geometry exposes the text viewport");
        const RECT expectedScreenExt = DipRectToScreenRect(window, paint.textRect);

        std::array<DirectWritePointerProbe, 3u> probes{
            CreateDirectWritePointerProbeForTextSpan(
                window.Host(), mixedBiDiText, paint.textRect, flowDirection, spans[0].first, spans[0].second, "native text store mixed-BiDi first span probe"),
            CreateDirectWritePointerProbeForTextSpan(
                window.Host(), mixedBiDiText, paint.textRect, flowDirection, spans[1].first, spans[1].second, "native text store mixed-BiDi script span probe"),
            CreateDirectWritePointerProbeForTextSpan(window.Host(),
                                                     mixedBiDiText,
                                                     paint.textRect,
                                                     flowDirection,
                                                     spans[2].first,
                                                     spans[2].second,
                                                     "native text store mixed-BiDi trailing span probe")};

        wil::com_ptr_nothrow<ITextStoreACP> store;
        store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
        Require(store != nullptr, "native text store mixed-BiDi geometry creates the text store");

        NativeTextStoreTestSink sink;
        RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_LAYOUT_CHANGE), "native text store mixed-BiDi geometry accepts a sink");

        TsViewCookie activeView = 0u;
        RECT screenExt{};
        std::array<RECT, spans.size()> textExts{};
        std::array<BOOL, spans.size()> clipped{};
        std::array<LONG, spans.size()> pointAcps{};
        HRESULT callbackHr = E_UNEXPECTED;

        sink.onLockGranted = [&](DWORD /*lockFlags*/) noexcept
        {
            callbackHr = store->GetActiveView(&activeView);
            if (FAILED(callbackHr))
            {
                return callbackHr;
            }

            callbackHr = store->GetScreenExt(activeView, &screenExt);
            if (FAILED(callbackHr))
            {
                return callbackHr;
            }

            for (size_t index = 0u; index < spans.size(); ++index)
            {
                callbackHr = store->GetTextExt(
                    activeView, static_cast<LONG>(spans[index].first), static_cast<LONG>(spans[index].second), &textExts[index], &clipped[index]);
                if (FAILED(callbackHr))
                {
                    return callbackHr;
                }

                const POINT queryPoint = DipPointToScreenPoint(window, probes[index].pointDip);
                callbackHr             = store->GetACPFromPoint(activeView, &queryPoint, 0u, &pointAcps[index]);
                if (FAILED(callbackHr))
                {
                    return callbackHr;
                }
            }

            return S_OK;
        };

        HRESULT sessionHr = E_UNEXPECTED;
        RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store mixed-BiDi geometry dispatches a read lock");
        RequireSucceeded(sessionHr, "native text store mixed-BiDi geometry lock callback succeeds");
        RequireSucceeded(callbackHr, "native text store mixed-BiDi geometry callback completed all queries");
        RequireRectNear(screenExt, expectedScreenExt, "native text store mixed-BiDi screen extent follows the visible text viewport");

        for (size_t index = 0u; index < spans.size(); ++index)
        {
            Require(clipped[index] == FALSE, "native text store mixed-BiDi text extent is not clipped");
            Require(textExts[index].right > textExts[index].left && textExts[index].bottom > textExts[index].top,
                    "native text store mixed-BiDi text extent has area");
            Require(pointAcps[index] == static_cast<LONG>(probes[index].expectedCaretIndex),
                    "native text store mixed-BiDi point maps to the DirectWrite logical ACP");
        }

        RequireSucceeded(store->UnadviseSink(&sink), "native text store mixed-BiDi geometry unadvises the sink");
    }
}

void TestNativeTextInputTextStoreMultilineTextExtUsesCaretLineGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"one\ntwo three\nfour");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 280.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    window.Host().SyncTextInput(field);

    D2D1_RECT_F expectedStartRect{};
    D2D1_RECT_F expectedEndRect{};
    Require(field->DebugGetCaretRect(window.Host(), 4u, expectedStartRect), "native text store multiline geometry can measure range start");
    Require(field->DebugGetCaretRect(window.Host(), 13u, expectedEndRect), "native text store multiline geometry can measure range end");
    const D2D1_RECT_F expectedDip = D2D1::RectF((std::min)(expectedStartRect.left, expectedEndRect.left),
                                                (std::min)(expectedStartRect.top, expectedEndRect.top),
                                                (std::max)(expectedStartRect.right, expectedEndRect.right),
                                                (std::max)(expectedStartRect.bottom, expectedEndRect.bottom));
    const RECT expectedScreen     = DipRectToScreenRect(window, expectedDip);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store multiline geometry creates the text store");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_LAYOUT_CHANGE), "native text store multiline geometry accepts a sink");

    TsViewCookie activeView = 0u;
    RECT textExt{};
    BOOL textExtClipped = TRUE;
    HRESULT callbackHr  = E_UNEXPECTED;
    sink.onLockGranted  = [&](DWORD /*lockFlags*/) noexcept
    {
        callbackHr = store->GetActiveView(&activeView);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        callbackHr = store->GetTextExt(activeView, 4, 13, &textExt, &textExtClipped);
        return callbackHr;
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store multiline geometry dispatches a read lock");
    RequireSucceeded(sessionHr, "native text store multiline geometry lock callback succeeds");
    RequireSucceeded(callbackHr, "native text store multiline geometry GetTextExt succeeds");
    Require(textExtClipped == FALSE, "native text store multiline text extent is not clipped");
    RequireRectNear(textExt, expectedScreen, "native text store multiline text extent follows line-level caret geometry");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store multiline geometry unadvises the sink");
}

void TestNativeTextInputTextStoreWrappedTextExtSpansVisualLineGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta gamma delta epsilon zeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 128.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    window.Host().SyncTextInput(field);

    D2D1_RECT_F firstCaretRect{};
    Require(field->DebugGetCaretRect(window.Host(), 0u, firstCaretRect), "native text store wrapped extent measures the first caret");

    std::optional<size_t> wrappedLineStart;
    D2D1_RECT_F previousLineCaretRect{};
    for (size_t index = 1u; index <= field->GetText().size(); ++index)
    {
        D2D1_RECT_F candidateRect{};
        Require(field->DebugGetCaretRect(window.Host(), index, candidateRect), "native text store wrapped extent measures candidate carets");
        if (candidateRect.top > firstCaretRect.top + 1.0f)
        {
            Require(index > 1u, "native text store wrapped extent found a non-empty first visual line");
            wrappedLineStart = index;
            Require(field->DebugGetCaretRect(window.Host(), index - 1u, previousLineCaretRect),
                    "native text store wrapped extent measures the previous visual-line caret");
            break;
        }
    }
    Require(wrappedLineStart.has_value(), "native text store wrapped extent found a caret on a wrapped visual line");

    const RECT expectedFirstLineScreen = DipRectToScreenRect(window,
                                                             D2D1::RectF((std::min)(firstCaretRect.left, previousLineCaretRect.left),
                                                                         (std::min)(firstCaretRect.top, previousLineCaretRect.top),
                                                                         (std::max)(firstCaretRect.right, previousLineCaretRect.right),
                                                                         (std::max)(firstCaretRect.bottom, previousLineCaretRect.bottom)));
    const LONG expectedMinimumWidth    = expectedFirstLineScreen.right - expectedFirstLineScreen.left;

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store wrapped extent creates the text store");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_LAYOUT_CHANGE), "native text store wrapped extent accepts a sink");

    TsViewCookie activeView = 0u;
    RECT textExt{};
    BOOL textExtClipped = TRUE;
    HRESULT callbackHr  = E_UNEXPECTED;
    sink.onLockGranted  = [&](DWORD /*lockFlags*/) noexcept
    {
        callbackHr = store->GetActiveView(&activeView);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        callbackHr = store->GetTextExt(activeView, 0, static_cast<LONG>(wrappedLineStart.value()), &textExt, &textExtClipped);
        return callbackHr;
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store wrapped extent dispatches a read lock");
    RequireSucceeded(sessionHr, "native text store wrapped extent lock callback succeeds");
    RequireSucceeded(callbackHr, "native text store wrapped extent GetTextExt succeeds");
    Require(textExtClipped == FALSE, "native text store wrapped extent is not clipped");
    Require(textExt.left <= expectedFirstLineScreen.left + 1, "native text store wrapped extent starts at the first visual-line caret");
    Require(textExt.right - textExt.left >= expectedMinimumWidth - 1, "native text store wrapped extent spans the selected first visual line");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store wrapped extent unadvises the sink");
}

void TestNativeTextInputTextStoreMultilinePointMapsToLineCaretAcp()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"one\ntwo three\nfour");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 280.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    window.Host().SyncTextInput(field);

    D2D1_RECT_F caretRect{};
    Require(field->DebugGetCaretRect(window.Host(), 8u, caretRect), "native text store multiline point can measure the target caret");
    const POINT queryPoint = DipPointToScreenPoint(window, D2D1::Point2F(caretRect.left, (caretRect.top + caretRect.bottom) * 0.5f));

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store multiline point creates the text store");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_LAYOUT_CHANGE), "native text store multiline point accepts a sink");

    TsViewCookie activeView = 0u;
    LONG pointAcp           = -1;
    HRESULT callbackHr      = E_UNEXPECTED;
    sink.onLockGranted      = [&](DWORD /*lockFlags*/) noexcept
    {
        callbackHr = store->GetActiveView(&activeView);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        callbackHr = store->GetACPFromPoint(activeView, &queryPoint, 0u, &pointAcp);
        return callbackHr;
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store multiline point dispatches a read lock");
    RequireSucceeded(sessionHr, "native text store multiline point lock callback succeeds");
    RequireSucceeded(callbackHr, "native text store multiline point GetACPFromPoint succeeds");
    Require(pointAcp == 8, "native text store multiline point maps to the line-level logical ACP");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store multiline point unadvises the sink");
}

void TestNativeTextInputTextStoreWrappedPointMapsToVisualLineCaretAcp()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta gamma delta epsilon");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 120.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    window.Host().SyncTextInput(field);

    D2D1_RECT_F firstRect{};
    Require(field->DebugGetCaretRect(window.Host(), 0u, firstRect), "native text store wrapped point can measure the first caret");

    std::optional<size_t> wrappedCaretIndex;
    D2D1_RECT_F wrappedCaretRect{};
    for (size_t index = 1u; index < field->GetText().size(); ++index)
    {
        D2D1_RECT_F candidateRect{};
        Require(field->DebugGetCaretRect(window.Host(), index, candidateRect), "native text store wrapped point can measure candidate carets");
        if (candidateRect.top > firstRect.top + 1.0f)
        {
            wrappedCaretIndex = index;
            wrappedCaretRect  = candidateRect;
            break;
        }
    }
    Require(wrappedCaretIndex.has_value(), "native text store wrapped point found a caret on a wrapped visual line");

    const POINT queryPoint = DipPointToScreenPoint(window, D2D1::Point2F(wrappedCaretRect.left, (wrappedCaretRect.top + wrappedCaretRect.bottom) * 0.5f));

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store wrapped point creates the text store");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_LAYOUT_CHANGE), "native text store wrapped point accepts a sink");

    TsViewCookie activeView = 0u;
    LONG pointAcp           = -1;
    HRESULT callbackHr      = E_UNEXPECTED;
    sink.onLockGranted      = [&](DWORD /*lockFlags*/) noexcept
    {
        callbackHr = store->GetActiveView(&activeView);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        callbackHr = store->GetACPFromPoint(activeView, &queryPoint, 0u, &pointAcp);
        return callbackHr;
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store wrapped point dispatches a read lock");
    RequireSucceeded(sessionHr, "native text store wrapped point lock callback succeeds");
    RequireSucceeded(callbackHr, "native text store wrapped point GetACPFromPoint succeeds");
    Require(pointAcp == static_cast<LONG>(wrappedCaretIndex.value()), "native text store wrapped point maps to the visual-line logical ACP");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store wrapped point unadvises the sink");
}

void TestNativeTextInputTextStoreReportsNoLayoutForEmptyBounds()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    window.Host().SyncTextInput(field);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 0.0f, 0.0f));

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store no-layout test creates a store");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native text store no-layout test advises a sink");

    HRESULT textExtHr   = E_UNEXPECTED;
    HRESULT screenExtHr = E_UNEXPECTED;
    HRESULT pointHr     = E_UNEXPECTED;
    sink.onLockGranted  = [&](DWORD /*lockFlags*/) noexcept
    {
        TsViewCookie activeView{};
        HRESULT callbackHr = store->GetActiveView(&activeView);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        RECT textExt{1, 2, 3, 4};
        BOOL textExtClipped = FALSE;
        textExtHr           = store->GetTextExt(activeView, 0, 5, &textExt, &textExtClipped);

        RECT screenExt{1, 2, 3, 4};
        screenExtHr = store->GetScreenExt(activeView, &screenExt);

        POINT queryPoint{0, 0};
        LONG pointAcp = -1;
        pointHr       = store->GetACPFromPoint(activeView, &queryPoint, 0u, &pointAcp);
        return S_OK;
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store no-layout read lock request dispatches to sink");
    RequireSucceeded(sessionHr, "native text store no-layout read lock callback succeeds");
    Require(textExtHr == TS_E_NOLAYOUT, "native text store reports no text layout for empty bounds");
    Require(screenExtHr == TS_E_NOLAYOUT, "native text store reports no screen layout for empty bounds");
    Require(pointHr == TS_E_NOLAYOUT, "native text store reports no point mapping for empty bounds");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store no-layout test unadvises the sink");
}

void TestNativeTextInputTextStoreInsertAtSelectionMutatesRetainedTextAndNotifiesSink()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(6u, 10u);
    window.Host().SyncTextInput(field);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store mutation test creates a store");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native text store mutation test advises a sink");

    LONG queryStart = -1;
    LONG queryEnd   = -1;
    TS_TEXTCHANGE queryChange{};
    LONG insertStart = -1;
    LONG insertEnd   = -1;
    TS_TEXTCHANGE change{};
    HRESULT callbackHr = E_UNEXPECTED;
    sink.onLockGranted = [&](DWORD /*lockFlags*/) noexcept
    {
        callbackHr = store->InsertTextAtSelection(TS_IAS_QUERYONLY, L"delta", 5u, &queryStart, &queryEnd, &queryChange);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        callbackHr = store->InsertTextAtSelection(0u, L"gamma", 5u, &insertStart, &insertEnd, &change);
        return callbackHr;
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READWRITE, &sessionHr), "native text store read-write lock request dispatches to sink");
    RequireSucceeded(sessionHr, "native text store read-write lock callback succeeds");
    RequireSucceeded(callbackHr, "native text store insert-at-selection succeeds");

    NativeTextInputState state;
    Require(field->GetText() == L"alpha gamma", "native text store mutation updates retained TextField text");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native text store mutation syncs native session state");
    Require(state.text == field->GetText() && state.caretIndex == 11u && ! state.selectionAnchorIndex.has_value(),
            "native text store mutation collapses native selection after inserted text");
    Require(queryStart == 6 && queryEnd == 11, "native text store query-only insert returns projected inserted ACP bounds");
    Require(queryChange.acpStart == 6 && queryChange.acpOldEnd == 10 && queryChange.acpNewEnd == 11,
            "native text store query-only insert reports the projected TS_TEXTCHANGE span");
    Require(insertStart == 6 && insertEnd == 11, "native text store insert-at-selection returns inserted ACP bounds");
    Require(change.acpStart == 6 && change.acpOldEnd == 10 && change.acpNewEnd == 11, "native text store reports the TS_TEXTCHANGE span");
    Require(sink.textChangeCount == 1u && sink.selectionChangeCount == 1u && sink.layoutChangeCount == 1u,
            "native text store notifies text, selection, and layout sinks after mutation");
    Require(sink.lastTextChange.acpStart == change.acpStart && sink.lastTextChange.acpOldEnd == change.acpOldEnd &&
                sink.lastTextChange.acpNewEnd == change.acpNewEnd,
            "native text store sink receives the mutation text-change span");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store mutation test unadvises the sink");
}

void TestNativeTextInputTextStoreEmojiRangeUsesLogicalUtf16Acp()
{
    using namespace RedSalamander::DxUi;

    const std::wstring emojiTextElement = MakeWomanTechnologistTextElement();
    const std::wstring replacement      = MakeUsFlagTextElement();
    const std::wstring originalText     = std::wstring(L"A") + emojiTextElement + L"Z";
    const size_t emojiStart             = 1u;
    const size_t emojiEnd               = emojiStart + emojiTextElement.size();
    const size_t replacementEnd         = emojiStart + replacement.size();

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(originalText);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(0u, 0u);
    window.Host().SyncTextInput(field);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store emoji range test creates a store");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native text store emoji range test advises a sink");

    LONG endAcp = -1;
    std::wstring readText;
    ULONG fetchedChars = 0u;
    TS_SELECTION_ACP emojiSelectionAfterSet{};
    ULONG fetchedSelection = 0u;
    LONG insertStart       = -1;
    LONG insertEnd         = -1;
    TS_TEXTCHANGE change{};
    HRESULT callbackHr = E_UNEXPECTED;
    sink.onLockGranted = [&](DWORD /*lockFlags*/) noexcept
    {
        callbackHr = store->GetEndACP(&endAcp);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        wchar_t buffer[32]{};
        TS_RUNINFO runInfo{};
        ULONG fetchedRuns = 0u;
        LONG nextAcp      = -1;
        callbackHr        = store->GetText(0, -1, buffer, static_cast<ULONG>(std::size(buffer)), &fetchedChars, &runInfo, 1u, &fetchedRuns, &nextAcp);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }
        readText.assign(buffer, fetchedChars);

        TS_SELECTION_ACP emojiSelection{};
        emojiSelection.acpStart           = static_cast<LONG>(emojiStart);
        emojiSelection.acpEnd             = static_cast<LONG>(emojiEnd);
        emojiSelection.style.ase          = TS_AE_END;
        emojiSelection.style.fInterimChar = FALSE;
        callbackHr                        = store->SetSelection(1u, &emojiSelection);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        callbackHr = store->GetSelection(TS_DEFAULT_SELECTION, 1u, &emojiSelectionAfterSet, &fetchedSelection);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        callbackHr = store->InsertTextAtSelection(0u, replacement.data(), static_cast<ULONG>(replacement.size()), &insertStart, &insertEnd, &change);
        return callbackHr;
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READWRITE, &sessionHr), "native text store emoji range read-write lock dispatches to sink");
    RequireSucceeded(sessionHr, "native text store emoji range callback succeeds");
    RequireSucceeded(callbackHr, "native text store emoji range replacement succeeds");

    const std::wstring expectedText = std::wstring(L"A") + replacement + L"Z";
    NativeTextInputState state;
    Require(endAcp == static_cast<LONG>(originalText.size()), "native text store emoji range reports logical UTF-16 end ACP");
    Require(readText == originalText && fetchedChars == originalText.size(), "native text store emoji range returns logical UTF-16 text");
    Require(fetchedSelection == 1u && emojiSelectionAfterSet.acpStart == static_cast<LONG>(emojiStart) &&
                emojiSelectionAfterSet.acpEnd == static_cast<LONG>(emojiEnd),
            "native text store emoji range selects the full emoji text element by ACP");
    Require(field->GetText() == expectedText, "native text store emoji range replacement updates retained TextField text");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native text store emoji range replacement syncs native session state");
    Require(state.text == expectedText && state.caretIndex == replacementEnd && ! state.selectionAnchorIndex.has_value(),
            "native text store emoji range replacement collapses native selection after the inserted emoji");
    Require(insertStart == static_cast<LONG>(emojiStart) && insertEnd == static_cast<LONG>(replacementEnd),
            "native text store emoji range replacement returns inserted ACP bounds");
    Require(change.acpStart == static_cast<LONG>(emojiStart) && change.acpOldEnd == static_cast<LONG>(emojiEnd) &&
                change.acpNewEnd == static_cast<LONG>(replacementEnd),
            "native text store emoji range replacement reports logical UTF-16 TS_TEXTCHANGE bounds");
    Require(sink.textChangeCount == 1u && sink.selectionChangeCount >= 1u && sink.layoutChangeCount >= 1u,
            "native text store emoji range replacement notifies text, selection, and layout sinks");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store emoji range test unadvises the sink");
}

void TestNativeTextInputTextStoreSetTextReplacesRangeAndNotifiesSink()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(6u, 10u);
    window.Host().SyncTextInput(field);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store creates a store for SetText replacement");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native text store SetText test advises a sink");

    TS_TEXTCHANGE change{};
    TS_SELECTION_ACP selectionAfter{};
    ULONG fetchedSelection = 0u;
    HRESULT callbackHr     = E_UNEXPECTED;
    sink.onLockGranted     = [&](DWORD /*lockFlags*/) noexcept
    {
        callbackHr = store->SetText(0u, 6, 10, L"delta", 5u, &change);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        return store->GetSelection(TS_DEFAULT_SELECTION, 1u, &selectionAfter, &fetchedSelection);
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READWRITE, &sessionHr), "native text store SetText read-write lock dispatches to sink");
    RequireSucceeded(sessionHr, "native text store SetText callback succeeds");
    RequireSucceeded(callbackHr, "native text store SetText mutation succeeds");

    NativeTextInputState state;
    Require(field->GetText() == L"alpha delta", "native text store SetText updates retained TextField text");
    Require(! field->GetSelectionRange().has_value(), "native text store SetText collapses retained TextField selection");
    Require(window.Host().TryReadNativeTextInputState(field, state), "native text store SetText syncs native session state");
    Require(state.text == field->GetText() && state.caretIndex == 11u && ! state.selectionAnchorIndex.has_value(),
            "native text store SetText collapses native selection after replacement");
    Require(fetchedSelection == 1u && selectionAfter.acpStart == 11 && selectionAfter.acpEnd == 11,
            "native text store SetText exposes the collapsed post-replacement selection");
    Require(change.acpStart == 6 && change.acpOldEnd == 10 && change.acpNewEnd == 11, "native text store SetText reports the replacement TS_TEXTCHANGE span");
    Require(sink.textChangeCount == 1u && sink.selectionChangeCount == 1u && sink.layoutChangeCount == 1u,
            "native text store SetText notifies text, selection, and layout sinks after replacement");
    Require(sink.lastTextChange.acpStart == change.acpStart && sink.lastTextChange.acpOldEnd == change.acpOldEnd &&
                sink.lastTextChange.acpNewEnd == change.acpNewEnd,
            "native text store SetText sink receives the replacement text-change span");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store SetText test unadvises the sink");
}

void TestNativeTextInputTextStoreExternalRetainedChangesNotifySink()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(0u, 5u);
    window.Host().SyncTextInput(field);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store creates a store for external retained-change notification");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native text store external-change test advises a sink");

    sink.onLockGranted = [](DWORD /*lockFlags*/) noexcept { return S_OK; };
    HRESULT sessionHr  = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store external-change baseline lock dispatches to sink");
    RequireSucceeded(sessionHr, "native text store external-change baseline lock succeeds");
    Require(sink.textChangeCount == 0u && sink.selectionChangeCount == 0u && sink.layoutChangeCount == 0u,
            "native text store baseline lock does not fabricate external changes");

    field->SetTextAndNotify(L"alpha gamma");
    field->SetSelectionRange(6u, 11u);
    window.Host().SyncTextInput(field);
    sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store observes an external retained text change");
    RequireSucceeded(sessionHr, "native text store external retained text lock succeeds");
    Require(sink.textChangeCount == 1u && sink.selectionChangeCount == 1u && sink.layoutChangeCount == 1u,
            "native text store notifies text, selection, and layout sinks after external retained text change");
    Require(sink.lastTextChange.acpStart == 0 && sink.lastTextChange.acpOldEnd == 10 && sink.lastTextChange.acpNewEnd == 11,
            "native text store external retained text change reports the previous and current document length");

    const uint32_t textChangesAfterText      = sink.textChangeCount;
    const uint32_t selectionChangesAfterText = sink.selectionChangeCount;
    const uint32_t layoutChangesAfterText    = sink.layoutChangeCount;
    field->SetSelectionRange(0u, 5u);
    window.Host().SyncTextInput(field);
    sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store observes an external retained selection change");
    RequireSucceeded(sessionHr, "native text store external retained selection lock succeeds");
    Require(sink.textChangeCount == textChangesAfterText && sink.selectionChangeCount == selectionChangesAfterText + 1u &&
                sink.layoutChangeCount == layoutChangesAfterText + 1u,
            "native text store notifies selection and layout sinks after external retained selection change");

    const uint32_t textChangesAfterSelection      = sink.textChangeCount;
    const uint32_t selectionChangesAfterSelection = sink.selectionChangeCount;
    const uint32_t layoutChangesAfterSelection    = sink.layoutChangeCount;
    field->SetBounds(D2D1::RectF(8.0f, 4.0f, 280.0f, 44.0f));
    window.Host().SyncTextInput(field);
    sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store observes an external retained layout change");
    RequireSucceeded(sessionHr, "native text store external retained layout lock succeeds");
    Require(sink.textChangeCount == textChangesAfterSelection && sink.selectionChangeCount == selectionChangesAfterSelection &&
                sink.layoutChangeCount == layoutChangesAfterSelection + 1u,
            "native text store notifies only the layout sink after external retained layout change");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store external-change test unadvises the sink");
}

void TestNativeTextInputTextStoreExternalChangeNotificationHandlesSinkRequestedLock()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(0u, 5u);
    window.Host().SyncTextInput(field);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store creates a store for reentrant external-change notification");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native text store reentrant external-change test advises a sink");

    sink.onLockGranted = [](DWORD /*lockFlags*/) noexcept { return S_OK; };
    HRESULT sessionHr  = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store reentrant external-change baseline lock dispatches to sink");
    RequireSucceeded(sessionHr, "native text store reentrant external-change baseline lock succeeds");
    Require(sink.textChangeCount == 0u && sink.selectionChangeCount == 0u && sink.layoutChangeCount == 0u,
            "native text store reentrant external-change baseline lock does not fabricate notifications");

    bool requestedNestedLock = false;
    HRESULT innerRequestHr   = E_UNEXPECTED;
    HRESULT innerSessionHr   = E_UNEXPECTED;
    sink.onTextChange        = [&](const TS_TEXTCHANGE* /*change*/) noexcept
    {
        if (! requestedNestedLock)
        {
            requestedNestedLock = true;
            innerRequestHr      = store->RequestLock(TS_LF_READ, &innerSessionHr);
        }
        return S_OK;
    };

    field->SetTextAndNotify(L"alpha gamma");
    field->SetSelectionRange(6u, 11u);
    window.Host().SyncTextInput(field);

    sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store observes a reentrant external retained text change");
    RequireSucceeded(sessionHr, "native text store reentrant external retained text lock succeeds");
    Require(requestedNestedLock, "native text store sink requested a nested lock from the external text-change notification");
    RequireSucceeded(innerRequestHr, "native text store accepts a nested sink-requested read lock during external notification");
    RequireSucceeded(innerSessionHr, "native text store nested sink-requested read lock succeeds during external notification");
    Require(sink.textChangeCount == 1u && sink.selectionChangeCount == 1u && sink.layoutChangeCount == 1u,
            "native text store emits each external retained-change notification once when the sink requests a lock");
    Require(sink.lastTextChange.acpStart == 0 && sink.lastTextChange.acpOldEnd == 10 && sink.lastTextChange.acpNewEnd == 11,
            "native text store reentrant external retained text change reports the previous and current document length");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store reentrant external-change test unadvises the sink");
}

void TestNativeTextInputTextStoreRepeatedEmojiExternalChangesStayBounded()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"seed");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 32.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(field->GetText().size(), field->GetText().size());
    window.Host().SyncTextInput(field);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store creates a store for repeated emoji external-change soak");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native text store repeated emoji external-change soak advises a sink");

    sink.onLockGranted = [](DWORD /*lockFlags*/) noexcept { return S_OK; };
    HRESULT sessionHr  = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store repeated emoji soak captures the baseline state");
    RequireSucceeded(sessionHr, "native text store repeated emoji soak baseline lock succeeds");
    Require(sink.textChangeCount == 0u && sink.selectionChangeCount == 0u && sink.layoutChangeCount == 0u,
            "native text store repeated emoji soak baseline does not fabricate notifications");

    uint32_t nestedLockCount = 0u;
    HRESULT innerRequestHr   = S_OK;
    HRESULT innerSessionHr   = S_OK;
    sink.onTextChange        = [&](const TS_TEXTCHANGE* /*change*/) noexcept
    {
        ++nestedLockCount;
        innerSessionHr = E_UNEXPECTED;
        innerRequestHr = store->RequestLock(TS_LF_READ, &innerSessionHr);
        return S_OK;
    };

    const std::wstring emojiChanges[] = {MakeGrinningFaceTextElement(),
                                         MakeWomanTechnologistTextElement(),
                                         MakeRainbowFlagTextElement(),
                                         MakeThumbsUpMediumSkinToneTextElement(),
                                         MakeHeartVariationTextElement()};
    std::wstring previousText(field->GetText());
    uint32_t expectedChangeCount = 0u;
    wchar_t suffix               = L'A';
    for (const std::wstring& emoji : emojiChanges)
    {
        std::wstring nextText = L"prefix ";
        nextText.append(emoji);
        nextText.append(L" suffix ");
        nextText.push_back(suffix);
        ++suffix;

        field->SetTextAndNotify(nextText);
        field->SetSelectionRange(nextText.size(), nextText.size());
        window.Host().SyncTextInput(field);

        sessionHr = E_UNEXPECTED;
        RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store repeated emoji soak observes the external change");
        RequireSucceeded(sessionHr, "native text store repeated emoji soak external-change lock succeeds");
        RequireSucceeded(innerRequestHr, "native text store repeated emoji soak nested sink-requested lock is accepted");
        RequireSucceeded(innerSessionHr, "native text store repeated emoji soak nested sink-requested lock succeeds");

        ++expectedChangeCount;
        Require(nestedLockCount == expectedChangeCount, "native text store repeated emoji soak grants one nested lock per text change");
        Require(sink.textChangeCount == expectedChangeCount && sink.selectionChangeCount == expectedChangeCount &&
                    sink.layoutChangeCount == expectedChangeCount,
                "native text store repeated emoji soak emits each notification family once per external change");
        Require(sink.lastTextChange.acpStart == 0 && sink.lastTextChange.acpOldEnd == static_cast<LONG>(previousText.size()) &&
                    sink.lastTextChange.acpNewEnd == static_cast<LONG>(nextText.size()),
                "native text store repeated emoji soak reports previous and current UTF-16 document lengths");

        previousText = std::move(nextText);
    }

    sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store repeated emoji soak accepts a final no-change lock");
    RequireSucceeded(sessionHr, "native text store repeated emoji soak final no-change lock succeeds");
    Require(sink.textChangeCount == expectedChangeCount && sink.selectionChangeCount == expectedChangeCount && sink.layoutChangeCount == expectedChangeCount &&
                nestedLockCount == expectedChangeCount,
            "native text store repeated emoji soak does not emit extra notifications after the final no-change lock");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store repeated emoji external-change soak unadvises the sink");
}

void TestNativeTextInputTextStoreUnadviseRequiresAdvisedSink()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    window.Host().SyncTextInput(field);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store creates a store for UnadviseSink identity validation");

    NativeTextStoreTestSink advisedSink;
    NativeTextStoreTestSink unrelatedSink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &advisedSink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native text store UnadviseSink test advises the primary sink");
    Require(FAILED(store->UnadviseSink(&unrelatedSink)), "native text store rejects UnadviseSink for an unrelated sink");

    advisedSink.onLockGranted = [](DWORD /*lockFlags*/) noexcept { return S_OK; };
    field->SetTextAndNotify(L"alpha gamma");
    window.Host().SyncTextInput(field);

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READ, &sessionHr), "native text store mismatched UnadviseSink keeps the advised sink connected");
    RequireSucceeded(sessionHr, "native text store mismatched UnadviseSink follow-up lock succeeds");
    Require(advisedSink.textChangeCount == 1u && unrelatedSink.textChangeCount == 0u,
            "native text store sends notifications only to the still-advised sink after mismatched UnadviseSink");

    RequireSucceeded(store->UnadviseSink(&advisedSink), "native text store UnadviseSink accepts the advised sink");
}

void TestNativeTextInputTextStoreReadWriteLockBracketsEditTransactionAndRejectsReentrantLock()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    window.Host().SyncTextInput(field);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store creates a store for read-write lock transaction coverage");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native text store transaction test advises a sink");

    HRESULT callbackHr     = E_UNEXPECTED;
    HRESULT innerRequestHr = E_UNEXPECTED;
    HRESULT innerSessionHr = E_UNEXPECTED;
    LONG endAcp            = -1;
    sink.onLockGranted     = [&](DWORD dwLockFlags) noexcept
    {
        if (dwLockFlags != TS_LF_READWRITE || sink.editTransactionDepth != 1u || sink.editTransactionStartCount != 1u || sink.editTransactionEndCount != 0u)
        {
            callbackHr = E_UNEXPECTED;
            return callbackHr;
        }

        innerRequestHr = store->RequestLock(TS_LF_READ, &innerSessionHr);
        if (FAILED(innerRequestHr) || innerSessionHr != TS_E_SYNCHRONOUS)
        {
            callbackHr = E_FAIL;
            return callbackHr;
        }

        callbackHr = store->GetEndACP(&endAcp);
        return callbackHr;
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READWRITE, &sessionHr), "native text store read-write lock dispatches to sink for transaction coverage");
    RequireSucceeded(sessionHr, "native text store read-write lock callback succeeds inside an edit transaction");
    RequireSucceeded(callbackHr, "native text store read-write transaction callback completed");
    Require(innerRequestHr == S_OK && innerSessionHr == TS_E_SYNCHRONOUS, "native text store rejects reentrant lock requests synchronously");
    Require(endAcp == 10, "native text store keeps the outer read-write lock usable after a rejected reentrant lock");
    Require(sink.editTransactionStartCount == 1u && sink.editTransactionEndCount == 1u && sink.editTransactionDepth == 0u,
            "native text store brackets read-write locks with a balanced edit transaction");

    TS_TEXTCHANGE change{};
    Require(store->SetText(0u, 0, 5, L"omega", 5u, &change) == TS_E_NOLOCK, "native text store releases the read-write lock after the callback returns");

    RequireSucceeded(store->UnadviseSink(&sink), "native text store transaction test unadvises the sink");
}

void TestNativeTextInputTextStoreEditableComboBoxSelectionAndMutation()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"alpha beta");
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(combo);
    combo->SetEditableSelectionRange(6u, 10u);
    window.Host().SyncTextInput(combo);

    wil::com_ptr_nothrow<ITextStoreACP> store;
    store.attach(window.Host().DebugCreateNativeTextInputTextStoreForTest());
    Require(store != nullptr, "native text store creates a store for focused editable ComboBox");

    NativeTextStoreTestSink sink;
    RequireSucceeded(store->AdviseSink(__uuidof(ITextStoreACPSink), &sink, TS_AS_TEXT_CHANGE | TS_AS_SEL_CHANGE | TS_AS_LAYOUT_CHANGE),
                     "native editable combo text store advises a sink");

    TS_SELECTION_ACP initialSelection{};
    ULONG fetchedSelection = 0u;
    LONG insertStart       = -1;
    LONG insertEnd         = -1;
    TS_TEXTCHANGE change{};
    HRESULT callbackHr = E_UNEXPECTED;
    sink.onLockGranted = [&](DWORD /*lockFlags*/) noexcept
    {
        callbackHr = store->GetSelection(TS_DEFAULT_SELECTION, 1u, &initialSelection, &fetchedSelection);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        TS_SELECTION_ACP alphaSelection{};
        alphaSelection.acpStart           = 0;
        alphaSelection.acpEnd             = 5;
        alphaSelection.style.ase          = TS_AE_END;
        alphaSelection.style.fInterimChar = FALSE;
        callbackHr                        = store->SetSelection(1u, &alphaSelection);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }

        callbackHr = store->InsertTextAtSelection(0u, L"omega", 5u, &insertStart, &insertEnd, &change);
        return callbackHr;
    };

    HRESULT sessionHr = E_UNEXPECTED;
    RequireSucceeded(store->RequestLock(TS_LF_READWRITE, &sessionHr), "native editable combo text store read-write lock dispatches to sink");
    RequireSucceeded(sessionHr, "native editable combo text store read-write callback succeeds");
    RequireSucceeded(callbackHr, "native editable combo text store selection and mutation callback succeeds");

    NativeTextInputState state;
    Require(fetchedSelection == 1u && initialSelection.acpStart == 6 && initialSelection.acpEnd == 10,
            "native editable combo text store exposes the retained selection range");
    Require(combo->GetText() == L"omega beta", "native editable combo text store mutation updates retained ComboBox text");
    Require(! combo->GetEditableSelectionRange().has_value(), "native editable combo text store mutation collapses retained ComboBox selection");
    Require(window.Host().TryReadNativeTextInputState(combo, state), "native editable combo text store mutation syncs native session state");
    Require(state.text == combo->GetText() && state.caretIndex == 5u && ! state.selectionAnchorIndex.has_value(),
            "native editable combo text store mutation collapses native selection after inserted text");
    Require(insertStart == 0 && insertEnd == 5, "native editable combo text store insert-at-selection returns inserted ACP bounds");
    Require(change.acpStart == 0 && change.acpOldEnd == 5 && change.acpNewEnd == 5, "native editable combo text store reports the TS_TEXTCHANGE span");
    Require(sink.textChangeCount == 1u && sink.selectionChangeCount == 2u && sink.layoutChangeCount == 2u,
            "native editable combo text store notifies text, selection, and layout sinks after selection plus mutation");

    RequireSucceeded(store->UnadviseSink(&sink), "native editable combo text store mutation test unadvises the sink");
}

void TestNativeTextInputBackendKeyToPaintMetricScenario()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 240.0f, 44.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, L'Z', 0));

    NativeTextInputState state;
    Require(window.Host().TryReadNativeTextInputState(field, state), "native key-to-paint scenario reads focused native state after key input");
    Require(state.text == L"alphaZ", "native key-to-paint scenario mutates retained text before paint");

    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

    WindowHostBitmapCapture capture;
    Require(window.Host().DebugCaptureBitmap(capture), "native key-to-paint scenario captures a rendered frame");
    Require(capture.widthPx > 0u && capture.heightPx > 0u && ! capture.bgraPixels.empty(), "native key-to-paint scenario has rendered pixels");
}

} // namespace

void RunNativeTextInputTests()
{
    TestNativeTextInputTsfDeactivateRestoresFocusAssociationBeforePoppingDocument();
    TestNativeTextInputTextStoreGetTextExtUsesRangeRectsBeforeCaretWalk();
    TestWindowHostDefaultsToNativeTextInputBackend();
    TestNativeTextInputBackendFocusesHostWithoutBridgeChild();
    TestNativeTextInputBackendActivatesTsfDocumentOnFocus();
    TestNativeTextInputBackendOwnsSystemCaretOnHostHwnd();
    TestNativeTextInputBackendMovesSystemCaretAfterKeyInput();
    TestNativeTextInputBackendClearsSessionWhenRootResets();
    TestNativeTextInputBackendClearsSessionWhenFocusedControlBecomesStale();
    TestNativeTextInputBackendUpdatesCaretWhenFocusedFieldMoves();
    TestNativeTextInputBackendHostFocusLossControlsNativeSession();
    TestNativeTextInputBackendTabMovesFocusToNextControl();
    TestNativeTextInputBackendMultilineDialogKeysStayHostOwned();
    TestNativeTextInputBackendEnterInvokesDefaultButton();
    TestNativeTextInputBackendEscapeInvokesCancelButton();
    TestNativeTextInputBackendRevealedMaskedFieldRemasksBeforeEscapeCancel();
    TestNativeTextInputBackendMenuKeyInvokesContextMenu();
    TestNativeTextInputBackendMultilineContextMenuKeysStayOnHostHwnd();
    TestNativeTextInputBackendSyncsPrintableCharIntoSessionState();
    TestNativeTextInputBackendSingleLineTabCharAndPasteReplacementSyncState();
    TestNativeTextInputBackendStateMirrorsInheritedFlowDirection();
    TestNativeTextInputBackendSyncsFocusedInheritedFlowDirectionChanges();
    TestNativeTextInputBackendEditableComboSyncsTextAndFlowDirection();
    TestNativeTextInputBackendEditableComboCommandsAndPopupSyncState();
    TestNativeTextInputBackendEditableComboExactMatchCommandsSyncSelection();
    TestNativeTextInputBackendEditableComboDeleteKeysAndPathWordDeleteSyncState();
    TestNativeTextInputBackendSyncsPrintableSysCharIntoSessionState();
    TestNativeTextInputTextStoreRequiresLockAndExposesTextSelectionGeometry();
    TestNativeTextInputTextStoreRejectsDestroyedControlDuringTeardown();
    TestNativeTextInputTextStoreExposesOwnerCompositionSink();
    TestNativeTextInputTextStoreExposesAcp2Surface();
    TestNativeTextInputTextStoreMixedBiDiGeometryUsesTextViewport();
    TestNativeTextInputTextStoreMultilineTextExtUsesCaretLineGeometry();
    TestNativeTextInputTextStoreWrappedTextExtSpansVisualLineGeometry();
    TestNativeTextInputTextStoreMultilinePointMapsToLineCaretAcp();
    TestNativeTextInputTextStoreWrappedPointMapsToVisualLineCaretAcp();
    TestNativeTextInputTextStoreReportsNoLayoutForEmptyBounds();
    TestNativeTextInputTextStoreInsertAtSelectionMutatesRetainedTextAndNotifiesSink();
    TestNativeTextInputTextStoreEmojiRangeUsesLogicalUtf16Acp();
    TestNativeTextInputTextStoreSetTextReplacesRangeAndNotifiesSink();
    TestNativeTextInputTextStoreExternalRetainedChangesNotifySink();
    TestNativeTextInputTextStoreExternalChangeNotificationHandlesSinkRequestedLock();
    TestNativeTextInputTextStoreRepeatedEmojiExternalChangesStayBounded();
    TestNativeTextInputTextStoreUnadviseRequiresAdvisedSink();
    TestNativeTextInputTextStoreReadWriteLockBracketsEditTransactionAndRejectsReentrantLock();
    TestNativeTextInputTextStoreEditableComboBoxSelectionAndMutation();
    TestNativeTextInputBackendKeyToPaintMetricScenario();
    TestNativeTextInputBackendImeStartEndUpdatesCompositionState();
    TestNativeTextInputBackendReadOnlySuppressesImeComposition();
    TestNativeTextInputBackendImeStartTracksSelectedCompositionRange();
    TestNativeTextInputBackendImeNoPayloadWithoutActiveCompositionDoesNotStartRange();
    TestNativeTextInputBackendImeWindowsTrackCaretRect();
    TestNativeTextInputBackendMultilineImeWindowsTrackCaretAcrossLines();
    TestNativeTextInputBackendImeWindowsUpdateWhenFocusedFieldMoves();
    TestNativeTextInputBackendMultilineImeWindowsUpdateWhenFocusedFieldMoves();
    TestNativeTextInputBackendImeWindowsUpdateWhenEditableComboMoves();
    TestNativeTextInputBackendImeWindowsUpdateAfterProgrammaticTextFieldCaretMove();
    TestNativeTextInputBackendImeWindowsUpdateAfterProgrammaticEditableComboCaretMove();
    TestNativeTextInputBackendImeWindowsUpdateAfterFocusedTextFieldPaddingChange();
    TestNativeTextInputBackendImeWindowsUpdateAfterFocusedEditableComboDensityChange();
    TestNativeTextInputBackendImeWindowsUpdateAfterDpiChange();
    TestNativeTextInputBackendImeWindowsUpdateAfterMultilineScroll();
    TestNativeTextInputBackendImeResultPayloadCommitsSelectionReplacement();
    TestNativeTextInputBackendImeCompositionPayloadPreviewsAndCancelRestoresBase();
    TestNativeTextInputBackendImeCompositionPaintExposesStyledInlineRanges();
    TestNativeTextInputBackendImeCompositionPaintExposesEditableComboInlineRanges();
    TestNativeTextInputBackendImeMultilineWrappedPreviewThenResultCommitsAtOriginalAnchor();
    TestNativeTextInputBackendImeCompositionOwnsSpecialKeys();
    TestNativeTextInputBackendImeCompositionLetsModifiedNavigationKeysRoute();
    TestNativeTextInputBackendMultilineImeCompositionOwnsSpecialKeys();
    TestNativeTextInputBackendImeResultOnlyResumesHostKeyRouting();
    TestNativeTextInputBackendImeResultAndCompositionKeepsKeyOwnership();
    TestNativeTextInputBackendSyncsKeySelectionIntoSessionState();
    TestNativeTextInputBackendExposesBackendNeutralTextInputState();
    TestNativeTextInputBackendSurrogatePairDeletionSyncsState();
    TestNativeTextInputBackendEmojiZwJDeletionSyncsState();
    TestNativeTextInputBackendRegionalIndicatorFlagDeletionSyncsState();
    TestNativeTextInputBackendEmojiSuffixDeletionSyncsState();
    TestNativeTextInputBackendEmojiShiftSelectionSyncsState();
    TestNativeTextInputBackendCtrlWordDeletionSyncsState();
    TestNativeTextInputBackendPointerCaretPlacementSyncsState();
    TestNativeTextInputBackendSingleLineDoubleClickSelectsWordOnHostHwnd();
    TestNativeTextInputBackendSingleLineRepeatedClicksWithoutClassDoubleClicksSelectWord();
    TestNativeTextInputBackendSingleLineThirdClickSelectsAllOnHostHwnd();
    TestNativeTextInputBackendSingleLineDragSelectionReplacesRangeOnHostHwnd();
    TestNativeTextInputBackendMixedBiDiDragSelectionCopiesLogicalOrderOnHostHwnd();
    TestNativeTextInputBackendMixedBiDiPointerHitTestMatchesDirectWriteVisualOrderOnHostHwnd();
    TestNativeTextInputBackendBiDiPointerScenarioMatrixMatchesDirectWriteOnHostHwnd();
    TestNativeTextInputBackendBiDiKeyboardLogicalBoundaryCommandsSyncState();
    TestNativeTextInputBackendMixedBiDiEditTransactionsPreserveLogicalOrder();
    TestNativeTextInputBackendPointerHitTestDoesNotSplitEmojiTextElements();
    TestNativeTextInputBackendEditableComboPointerHitTestDoesNotSplitEmojiTextElement();
    TestNativeTextInputBackendNoSelectionCopyLeavesClipboardUnchanged();
    TestNativeTextInputBackendCtrlCopyCutPasteSyncsState();
    TestNativeTextInputBackendUndoRedoAndRedoClear();
    TestNativeTextInputBackendEditTransactionsNotifyOnceAndIgnoreNoOps();
    TestNativeTextInputBackendEmojiClipboardReplacementUndoRedo();
    TestNativeTextInputBackendMaskedHiddenSuppressesCopyAndCut();
    TestNativeTextInputBackendMaskedExactPolicyCountsTextElements();
    TestNativeTextInputBackendMaskedConcealedPolicyUsesStableBuckets();
    TestNativeTextInputBackendMaskedConcealedPolicyRegeneratesEpochs();
    TestNativeTextInputDeactivateSecureClearsCachedText();
    TestNativeTextInputBackendConcealedEditingAffordancesAndPointerPolicy();
    TestNativeTextInputBackendRevealedMaskedFieldAllowsCopyAndCut();
    TestNativeTextInputBackendRevealedMaskedFieldRemasksOnBlurReadOnlyAndDisable();
    TestNativeTextInputBackendImeCompositionClearsOnWindowDeactivate();
    TestNativeTextInputBackendRevealedMaskedFieldRemasksOnWindowDeactivate();
    TestNativeTextInputBackendMaskedRevealButtonRemasksOnCaptureLoss();
    TestNativeTextInputBackendMaskedRevealButtonPeeksWithoutClearingSecret();
    TestNativeTextInputBackendMaskedRevealButtonSupportsKeyboardPeek();
    TestNativeTextInputBackendPasswordRevealModesControlAffordanceAndVisibility();
    TestNativeTextInputBackendReadOnlyAllowsCopyAndSuppressesMutation();
    TestNativeTextInputBackendMultilineCtrlCopyPastePreservesLogicalNewlines();
    TestNativeTextInputBackendMultilineCharAndReturnReplacementSyncState();
    TestNativeTextInputBackendEditMessagesCopyPasteCutClearSelection();
    TestNativeTextInputBackendEditMessagesRoundTripWin32Protocol();
    TestNativeTextInputBackendEditMessagesSetTextClearsComposition();
    TestNativeTextInputBackendEditMessagesFallBackWithoutTextInput();
    TestNativeTextInputBackendClearWithoutSelectionLeavesTextAndClipboardUnchanged();
}
