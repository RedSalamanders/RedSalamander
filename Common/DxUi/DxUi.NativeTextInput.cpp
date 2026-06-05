#include "DxUi.Internal.h"

#include "Helpers.h"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstring>
#include <limits>

#include <imm.h>
#include <msctf.h>
#include <textstor.h>

#pragma comment(lib, "imm32.lib")

namespace RedSalamander::DxUi
{
namespace
{
[[nodiscard]] NativeTextInputState ToNativeTextInputState(const Control& control, const TextInputState& controlState) noexcept
{
    NativeTextInputState state;
    state.text                 = controlState.text;
    state.selectionAnchorIndex = controlState.selectionAnchorIndex;
    state.caretIndex           = controlState.caretIndex;
    state.firstVisibleLine     = controlState.firstVisibleLine;
    state.readOnly             = controlState.readOnly;
    state.masked               = controlState.masked;
    state.multiline            = controlState.multiline;
    state.flowDirection        = control.GetFlowDirection();
    state.readingDirection     = ResolveReadingDirection(state.flowDirection);
    if (const auto* textField = dynamic_cast<const TextField*>(&control))
    {
        state.passwordRevealMode    = textField->GetPasswordRevealMode();
        state.passwordRevealState   = textField->GetPasswordRevealState();
        state.maskLengthPolicy      = textField->GetPasswordMaskLengthPolicy();
        state.secretVisibleDotCount = textField->GetSecretVisibleDotCount();
    }
    return state;
}

[[nodiscard]] RECT NormalizeCaretRectPx(const RECT& rect) noexcept
{
    RECT normalized = rect;
    if (normalized.right <= normalized.left)
    {
        normalized.right = normalized.left + 1;
    }
    if (normalized.bottom <= normalized.top)
    {
        normalized.bottom = normalized.top + std::max<LONG>(12, GetSystemMetrics(SM_CYCURSOR));
    }
    return normalized;
}

[[nodiscard]] bool HasTextSelection(const TextInputState& state) noexcept
{
    if (! state.selectionAnchorIndex.has_value())
    {
        return false;
    }

    const size_t selectionAnchorIndex = std::min(state.selectionAnchorIndex.value(), state.text.size());
    const size_t caretIndex           = std::min(state.caretIndex, state.text.size());
    return selectionAnchorIndex != caretIndex;
}

[[nodiscard]] std::wstring NormalizeWin32EditTextFromControlText(std::wstring_view text, bool multiline)
{
    if (! multiline || text.empty())
    {
        return std::wstring(text);
    }

    std::wstring normalized;
    normalized.reserve(text.size() + static_cast<size_t>(std::count(text.begin(), text.end(), L'\n')));
    for (wchar_t ch : text)
    {
        if (ch == L'\r')
        {
            continue;
        }

        if (ch == L'\n')
        {
            normalized.push_back(L'\r');
        }

        normalized.push_back(ch);
    }

    return normalized;
}

[[nodiscard]] std::wstring NormalizeControlTextFromWin32EditText(std::wstring_view text, bool multiline)
{
    if (! multiline || text.empty())
    {
        return std::wstring(text);
    }

    std::wstring normalized;
    normalized.reserve(text.size());
    for (size_t index = 0u; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch == L'\r')
        {
            if (index + 1u < text.size() && text[index + 1u] == L'\n')
            {
                normalized.push_back(L'\n');
                ++index;
            }
            else
            {
                normalized.push_back(L'\n');
            }
            continue;
        }

        normalized.push_back(ch);
    }

    return normalized;
}

[[nodiscard]] std::pair<size_t, size_t> GetTextInputSelectionRange(const TextInputState& state) noexcept
{
    const size_t anchorIndex = state.selectionAnchorIndex.value_or(state.caretIndex);
    return {(std::min)(anchorIndex, state.caretIndex), (std::max)(anchorIndex, state.caretIndex)};
}

void SetTextInputSelectionRange(TextInputState& state, size_t start, size_t end) noexcept
{
    start = (std::min)(start, state.text.size());
    end   = (std::min)(end, state.text.size());
    if (start >= end)
    {
        state.selectionAnchorIndex.reset();
        state.caretIndex = end;
        return;
    }

    state.selectionAnchorIndex = start;
    state.caretIndex           = end;
}

[[nodiscard]] size_t MapControlIndexToWin32EditIndex(std::wstring_view controlText, size_t controlIndex, bool multiline) noexcept
{
    const size_t clampedIndex = std::min(controlIndex, controlText.size());
    if (! multiline)
    {
        return clampedIndex;
    }

    size_t editIndex = 0u;
    for (size_t index = 0u; index < clampedIndex; ++index)
    {
        editIndex += (controlText[index] == L'\n') ? 2u : 1u;
    }
    return editIndex;
}

[[nodiscard]] size_t MapWin32EditIndexToControlIndex(std::wstring_view win32EditText, size_t editIndex, bool multiline) noexcept
{
    const size_t clampedIndex = std::min(editIndex, win32EditText.size());
    if (! multiline)
    {
        return clampedIndex;
    }

    size_t controlIndex = 0u;
    size_t index        = 0u;
    while (index < win32EditText.size() && index < clampedIndex)
    {
        if (win32EditText[index] == L'\r' && index + 1u < win32EditText.size() && win32EditText[index + 1u] == L'\n')
        {
            if (index + 2u <= clampedIndex)
            {
                index += 2u;
                ++controlIndex;
                continue;
            }
            break;
        }

        ++index;
        ++controlIndex;
    }

    return controlIndex;
}

[[nodiscard]] bool ReplaceTextInputSelection(TextInputState& state, std::wstring_view replacement)
{
    if (state.readOnly)
    {
        return false;
    }

    auto [selectionStart, selectionEnd] = GetTextInputSelectionRange(state);
    selectionStart                      = std::min(selectionStart, state.text.size());
    selectionEnd                        = std::min(selectionEnd, state.text.size());
    if (selectionEnd < selectionStart)
    {
        std::swap(selectionStart, selectionEnd);
    }

    state.text.replace(selectionStart, selectionEnd - selectionStart, replacement);
    state.caretIndex = selectionStart + replacement.size();
    state.selectionAnchorIndex.reset();
    return true;
}

[[nodiscard]] bool NativeTextInputSelectionChanged(const NativeTextInputState& previous, const NativeTextInputState& current) noexcept
{
    return previous.caretIndex != current.caretIndex || previous.selectionAnchorIndex != current.selectionAnchorIndex;
}

[[nodiscard]] bool NativeTextInputActiveTextPositionChanged(const NativeTextInputState& previous, const NativeTextInputState& current) noexcept
{
    return previous.caretIndex != current.caretIndex || previous.compositionCursorIndex != current.compositionCursorIndex;
}

[[nodiscard]] bool NativeTextInputCompositionChanged(const NativeTextInputState& previous, const NativeTextInputState& current) noexcept
{
    return previous.compositionStartIndex != current.compositionStartIndex || previous.compositionEndIndex != current.compositionEndIndex ||
           previous.compositionCursorIndex != current.compositionCursorIndex || previous.compositionClauseBoundaries != current.compositionClauseBoundaries;
}

[[nodiscard]] bool NativeTextInputConversionTargetChanged(const NativeTextInputState& previous, const NativeTextInputState& current) noexcept
{
    return previous.conversionTargetStartIndex != current.conversionTargetStartIndex || previous.conversionTargetEndIndex != current.conversionTargetEndIndex;
}

[[nodiscard]] std::pair<size_t, size_t> ResolveImeBaseCompositionRange(const TextInputState& state) noexcept
{
    const size_t textLength = state.text.size();
    const size_t caretIndex = std::min(state.caretIndex, textLength);
    if (! state.selectionAnchorIndex.has_value())
    {
        return {caretIndex, caretIndex};
    }

    const size_t selectionAnchorIndex = std::min(state.selectionAnchorIndex.value(), textLength);
    return {std::min(selectionAnchorIndex, caretIndex), std::max(selectionAnchorIndex, caretIndex)};
}

void ReplaceNativeTextInputRange(TextInputState& state, size_t rangeStart, size_t rangeEnd, std::wstring_view replacement)
{
    rangeStart = std::min(rangeStart, state.text.size());
    rangeEnd   = std::min(rangeEnd, state.text.size());
    if (rangeEnd < rangeStart)
    {
        std::swap(rangeStart, rangeEnd);
    }

    state.text.replace(rangeStart, rangeEnd - rangeStart, replacement);
    state.caretIndex = rangeStart + replacement.size();
    state.selectionAnchorIndex.reset();
}

[[nodiscard]] std::optional<std::wstring> ReadImeString(HIMC inputContext, DWORD index)
{
    const LONG byteCount = ImmGetCompositionStringW(inputContext, index, nullptr, 0u);
    if (byteCount < 0 || (byteCount % static_cast<LONG>(sizeof(wchar_t))) != 0)
    {
        return std::nullopt;
    }

    std::wstring value(static_cast<size_t>(byteCount) / sizeof(wchar_t), L'\0');
    if (byteCount == 0)
    {
        return value;
    }

    const LONG readCount = ImmGetCompositionStringW(inputContext, index, value.data(), static_cast<DWORD>(byteCount));
    if (readCount < 0 || (readCount % static_cast<LONG>(sizeof(wchar_t))) != 0)
    {
        return std::nullopt;
    }

    value.resize(static_cast<size_t>(readCount) / sizeof(wchar_t));
    return value;
}

[[nodiscard]] std::vector<uint8_t> ReadImeBytes(HIMC inputContext, DWORD index)
{
    const LONG byteCount = ImmGetCompositionStringW(inputContext, index, nullptr, 0u);
    if (byteCount <= 0)
    {
        return {};
    }

    std::vector<uint8_t> value(static_cast<size_t>(byteCount));
    const LONG readCount = ImmGetCompositionStringW(inputContext, index, value.data(), static_cast<DWORD>(value.size()));
    if (readCount <= 0)
    {
        return {};
    }

    value.resize(static_cast<size_t>(readCount));
    return value;
}

[[nodiscard]] std::vector<uint32_t> ReadImeDwords(HIMC inputContext, DWORD index)
{
    const std::vector<uint8_t> bytes = ReadImeBytes(inputContext, index);
    if (bytes.empty() || (bytes.size() % sizeof(uint32_t)) != 0u)
    {
        return {};
    }

    std::vector<uint32_t> value(bytes.size() / sizeof(uint32_t));
    std::memcpy(value.data(), bytes.data(), bytes.size());
    return value;
}

[[nodiscard]] bool IsImeTargetAttribute(uint8_t attribute) noexcept
{
    return attribute == ATTR_TARGET_CONVERTED || attribute == ATTR_TARGET_NOTCONVERTED;
}

[[nodiscard]] std::vector<size_t> ResolveImeClauseBoundaries(size_t compositionStartIndex, size_t compositionLength, const NativeTextInputImePayload& payload)
{
    std::vector<size_t> boundaries;
    boundaries.reserve(payload.compositionClauses.size());
    for (const uint32_t clauseOffset : payload.compositionClauses)
    {
        const size_t clampedOffset = (std::min)(static_cast<size_t>(clauseOffset), compositionLength);
        const size_t boundary      = compositionStartIndex + clampedOffset;
        if (boundaries.empty() || boundaries.back() != boundary)
        {
            boundaries.push_back(boundary);
        }
    }
    return boundaries;
}
} // namespace

WindowHost::NativeSystemCaret::~NativeSystemCaret() noexcept
{
    Reset();
}

bool WindowHost::NativeSystemCaret::Create(HWND ownerHwnd, int widthPx, int heightPx) noexcept
{
    Reset();
    if (! ownerHwnd || widthPx <= 0 || heightPx <= 0)
    {
        return false;
    }

    if (CreateCaret(ownerHwnd, nullptr, widthPx, heightPx) == FALSE)
    {
        return false;
    }

    _ownerHwnd = ownerHwnd;
    _created   = true;
    _visible   = false;
    _widthPx   = widthPx;
    _heightPx  = heightPx;
    return true;
}

bool WindowHost::NativeSystemCaret::Show() noexcept
{
    if (! _created || ! _ownerHwnd)
    {
        return false;
    }
    if (_visible)
    {
        return true;
    }

    if (ShowCaret(_ownerHwnd) == FALSE)
    {
        return false;
    }

    _visible = true;
    return true;
}

void WindowHost::NativeSystemCaret::Hide() noexcept
{
    if (_created && _visible && _ownerHwnd)
    {
        static_cast<void>(HideCaret(_ownerHwnd));
    }
    _visible = false;
}

void WindowHost::NativeSystemCaret::Reset() noexcept
{
    Hide();
    if (_created)
    {
        static_cast<void>(DestroyCaret());
    }
    _ownerHwnd = nullptr;
    _created   = false;
    _visible   = false;
    _widthPx   = 0;
    _heightPx  = 0;
}

void WindowHost::ActivateTextInput(Control* control) noexcept
{
    ActivateNativeTextInputSession(control);
}

void WindowHost::DeactivateTextInput(bool restoreHostFocus) noexcept
{
    DeactivateNativeTextInputSession(restoreHostFocus);
}

void WindowHost::SetTextInputBackend(TextInputBackend backend) noexcept
{
    if (_textInputBackend == backend)
    {
        return;
    }

    DeactivateTextInput(false);

    _textInputBackend = TextInputBackend::Native;

    if (! _focusedControl || ! _focusedControl->SupportsTextInput())
    {
        return;
    }

    ActivateTextInput(_focusedControl);
}

TextInputBackend WindowHost::GetTextInputBackend() const noexcept
{
    return _textInputBackend;
}

bool WindowHost::HasActiveNativeTextInputSession() const noexcept
{
    return _nativeTextInputControl != nullptr;
}

void WindowHost::ActivateNativeTextInputSession(Control* control) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    if (! control || ! control->SupportsTextInput())
    {
        DeactivateNativeTextInputSession(false);
        return;
    }

    TextInputState controlState;
    if (! control->ExportTextInputState(controlState))
    {
        DeactivateNativeTextInputSession(false);
        return;
    }

    const bool controlChanged = _nativeTextInputControl != control;
    if (controlChanged)
    {
        ClearNativeTextInputCompositionState();
    }
    _nativeTextInputControl         = control;
    _nativeTextInputStateCache      = ToNativeTextInputState(*control, controlState);
    _nativeTextInputStateCacheValid = true;
    ApplyNativeTextInputCompositionStateToCache();
    if (controlChanged)
    {
        ++_nativeTextInputEventCounters.activationCount;
    }

    if (_hwnd && GetFocus() != _hwnd)
    {
        SetFocus(_hwnd);
    }
    if (controlChanged || ! _nativeTextInputTsfActive)
    {
        static_cast<void>(ActivateNativeTextInputTsf(control));
    }
    UpdateNativeTextInputCaret();

    Debug::Perf::Emit(L"dxui.textinput.activate_us",
                      L"native-session",
                      Debug::Perf::ElapsedUs(startedAt),
                      _nativeTextInputStateCache.multiline ? 1u : 0u,
                      _nativeTextInputStateCache.masked ? 1u : 0u,
                      S_OK);
}

void WindowHost::DeactivateNativeTextInputSession(bool restoreHostFocus) noexcept
{
    DeactivateNativeTextInputTsf();

    if (! _nativeTextInputControl)
    {
        _nativeTextInputStateCacheValid = false;
        ClearNativeTextInputCompositionState();
        _pendingNativeTextInputPaintMetric.reset();
        return;
    }

    _nativeTextInputControl         = nullptr;
    _nativeTextInputStateCacheValid = false;
    ClearNativeTextInputCompositionState();
    _pendingNativeTextInputPaintMetric.reset();
    ++_nativeTextInputEventCounters.deactivationCount;
    DestroyNativeTextInputCaret();

    if (restoreHostFocus && _hwnd && GetFocus() != _hwnd)
    {
        SetFocus(_hwnd);
    }
}

bool WindowHost::ActivateNativeTextInputTsf(Control* control) noexcept
{
    DeactivateNativeTextInputTsf();
    ++_nativeTextInputEventCounters.tsfActivationAttemptCount;

    if (! control || ! control->SupportsTextInput() || ! _hwnd)
    {
        ++_nativeTextInputEventCounters.tsfActivationFailureCount;
        return false;
    }

    wil::com_ptr_nothrow<ITextStoreACP> textStore;
    textStore.attach(CreateNativeTextInputTextStore(*this, *control));
    if (! textStore)
    {
        ++_nativeTextInputEventCounters.tsfActivationFailureCount;
        return false;
    }

    wil::com_ptr_nothrow<ITfThreadMgr> threadMgr;
    HRESULT hr = CoCreateInstance(CLSID_TF_ThreadMgr, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(threadMgr.put()));
    if (hr == CO_E_NOTINITIALIZED)
    {
        const HRESULT coInitHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        if (FAILED(coInitHr))
        {
            ++_nativeTextInputEventCounters.tsfActivationFailureCount;
            return false;
        }

        _nativeTextInputTsfComInitialized = true;
        hr                                = CoCreateInstance(CLSID_TF_ThreadMgr, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(threadMgr.put()));
    }

    if (FAILED(hr) || ! threadMgr)
    {
        if (_nativeTextInputTsfComInitialized)
        {
            CoUninitialize();
            _nativeTextInputTsfComInitialized = false;
        }
        ++_nativeTextInputEventCounters.tsfActivationFailureCount;
        return false;
    }

    TfClientId clientId = 0u;
    hr                  = threadMgr->Activate(&clientId);
    if (FAILED(hr) || clientId == 0u)
    {
        if (_nativeTextInputTsfComInitialized)
        {
            CoUninitialize();
            _nativeTextInputTsfComInitialized = false;
        }
        ++_nativeTextInputEventCounters.tsfActivationFailureCount;
        return false;
    }

    _nativeTextInputTsfClientId = clientId;
    static_cast<void>(threadMgr.query_to(_nativeTextInputTsfThreadMgr.put()));

    auto failActivation = [this]() noexcept
    {
        DeactivateNativeTextInputTsf();
        ++_nativeTextInputEventCounters.tsfActivationFailureCount;
        return false;
    };

    wil::com_ptr_nothrow<ITfDocumentMgr> documentMgr;
    hr = threadMgr->CreateDocumentMgr(documentMgr.put());
    if (FAILED(hr) || ! documentMgr)
    {
        return failActivation();
    }

    wil::com_ptr_nothrow<ITfContext> context;
    TfEditCookie editCookie = 0u;
    hr                      = documentMgr->CreateContext(clientId, 0u, textStore.get(), context.put(), &editCookie);
    if (FAILED(hr) || ! context)
    {
        return failActivation();
    }

    hr = documentMgr->Push(context.get());
    if (FAILED(hr))
    {
        return failActivation();
    }

    static_cast<void>(threadMgr->SetFocus(documentMgr.get()));
    static_cast<void>(documentMgr.query_to(_nativeTextInputTsfDocumentMgr.put()));
    static_cast<void>(context.query_to(_nativeTextInputTsfContext.put()));
    static_cast<void>(textStore.query_to(_nativeTextInputTsfTextStore.put()));
    _nativeTextInputTsfActive = true;
    ++_nativeTextInputEventCounters.tsfActivationSuccessCount;
    return true;
}

void WindowHost::DeactivateNativeTextInputTsf() noexcept
{
    const bool wasActive = _nativeTextInputTsfActive;

    if (_nativeTextInputTsfDocumentMgr)
    {
        wil::com_ptr_nothrow<ITfDocumentMgr> documentMgr;
        if (SUCCEEDED(_nativeTextInputTsfDocumentMgr.query_to(documentMgr.put())) && documentMgr)
        {
            static_cast<void>(documentMgr->Pop(TF_POPF_ALL));
        }
    }

    if (_nativeTextInputTsfThreadMgr && _nativeTextInputTsfClientId != 0u)
    {
        wil::com_ptr_nothrow<ITfThreadMgr> threadMgr;
        if (SUCCEEDED(_nativeTextInputTsfThreadMgr.query_to(threadMgr.put())) && threadMgr)
        {
            static_cast<void>(threadMgr->Deactivate());
        }
    }

    _nativeTextInputTsfTextStore.reset();
    _nativeTextInputTsfContext.reset();
    _nativeTextInputTsfDocumentMgr.reset();
    _nativeTextInputTsfThreadMgr.reset();
    _nativeTextInputTsfClientId = 0u;
    _nativeTextInputTsfActive   = false;

    if (_nativeTextInputTsfComInitialized)
    {
        CoUninitialize();
        _nativeTextInputTsfComInitialized = false;
    }

    if (wasActive)
    {
        ++_nativeTextInputEventCounters.tsfDeactivationCount;
    }
}

void WindowHost::SyncNativeTextInputSession(Control* control) noexcept
{
    if (! control || control != _nativeTextInputControl)
    {
        return;
    }

    TextInputState controlState;
    if (! control->ExportTextInputState(controlState))
    {
        return;
    }

    const std::optional<NativeTextInputState> previousState =
        _nativeTextInputStateCacheValid ? std::optional<NativeTextInputState>{_nativeTextInputStateCache} : std::nullopt;
    _nativeTextInputStateCache      = ToNativeTextInputState(*control, controlState);
    _nativeTextInputStateCacheValid = true;
    ApplyNativeTextInputCompositionStateToCache();
    ++_nativeTextInputEventCounters.synchronizationCount;
    if (previousState.has_value())
    {
        RaiseNativeTextInputAccessibilityEvents(previousState.value());
    }
    UpdateNativeTextInputCaret();
}

void WindowHost::RaiseNativeTextInputAccessibilityEvent(TextInputAutomationEventKind kind) noexcept
{
    if (! RaiseWindowHostTextInputAutomationEvent(_hwnd, _nativeTextInputControl, kind))
    {
        return;
    }

    switch (kind)
    {
        case TextInputAutomationEventKind::TextChanged: ++_nativeTextInputEventCounters.uiaTextChangedCount; break;
        case TextInputAutomationEventKind::TextSelectionChanged: ++_nativeTextInputEventCounters.uiaTextSelectionChangedCount; break;
        case TextInputAutomationEventKind::ActiveTextPositionChanged: ++_nativeTextInputEventCounters.uiaActiveTextPositionChangedCount; break;
        case TextInputAutomationEventKind::TextEditCompositionChanged: ++_nativeTextInputEventCounters.uiaTextEditTextChangedCount; break;
        case TextInputAutomationEventKind::TextEditConversionTargetChanged: ++_nativeTextInputEventCounters.uiaTextEditConversionTargetChangedCount; break;
        default: break;
    }
}

void WindowHost::RaiseNativeTextInputAccessibilityEvents(const NativeTextInputState& previousState) noexcept
{
    if (! _nativeTextInputStateCacheValid)
    {
        return;
    }

    const NativeTextInputState& currentState = _nativeTextInputStateCache;
    if (previousState.text != currentState.text)
    {
        RaiseNativeTextInputAccessibilityEvent(TextInputAutomationEventKind::TextChanged);
    }
    if (NativeTextInputSelectionChanged(previousState, currentState))
    {
        RaiseNativeTextInputAccessibilityEvent(TextInputAutomationEventKind::TextSelectionChanged);
    }
    if (NativeTextInputActiveTextPositionChanged(previousState, currentState))
    {
        RaiseNativeTextInputAccessibilityEvent(TextInputAutomationEventKind::ActiveTextPositionChanged);
    }
    if (NativeTextInputCompositionChanged(previousState, currentState))
    {
        RaiseNativeTextInputAccessibilityEvent(TextInputAutomationEventKind::TextEditCompositionChanged);
    }
    if (NativeTextInputConversionTargetChanged(previousState, currentState))
    {
        RaiseNativeTextInputAccessibilityEvent(TextInputAutomationEventKind::TextEditConversionTargetChanged);
    }
}

void WindowHost::ApplyNativeTextInputCompositionStateToCache() noexcept
{
    if (! _nativeTextInputStateCacheValid)
    {
        return;
    }

    if (! _nativeTextInputImeComposing)
    {
        _nativeTextInputStateCache.compositionStartIndex.reset();
        _nativeTextInputStateCache.compositionEndIndex.reset();
        _nativeTextInputStateCache.conversionTargetStartIndex.reset();
        _nativeTextInputStateCache.conversionTargetEndIndex.reset();
        _nativeTextInputStateCache.compositionCursorIndex.reset();
        _nativeTextInputStateCache.compositionClauseBoundaries.clear();
        return;
    }

    _nativeTextInputStateCache.compositionStartIndex       = _nativeTextInputCompositionStartIndex;
    _nativeTextInputStateCache.compositionEndIndex         = _nativeTextInputCompositionEndIndex;
    _nativeTextInputStateCache.conversionTargetStartIndex  = _nativeTextInputConversionTargetStartIndex;
    _nativeTextInputStateCache.conversionTargetEndIndex    = _nativeTextInputConversionTargetEndIndex;
    _nativeTextInputStateCache.compositionCursorIndex      = _nativeTextInputCompositionCursorIndex;
    _nativeTextInputStateCache.compositionClauseBoundaries = _nativeTextInputCompositionClauseBoundaries;
}

void WindowHost::ClearNativeTextInputCompositionState() noexcept
{
    _nativeTextInputImeComposing = false;
    _nativeTextInputCompositionStartIndex.reset();
    _nativeTextInputCompositionEndIndex.reset();
    _nativeTextInputConversionTargetStartIndex.reset();
    _nativeTextInputConversionTargetEndIndex.reset();
    _nativeTextInputCompositionCursorIndex.reset();
    _nativeTextInputCompositionClauseBoundaries.clear();
    _nativeTextInputImeBaseState.reset();
    ApplyNativeTextInputCompositionStateToCache();
}

NativeTextInputImePayload WindowHost::ReadNativeTextInputImePayload(LPARAM compositionFlags) noexcept
{
    if (_debugNativeTextInputImePayload.has_value())
    {
        NativeTextInputImePayload payload = std::move(_debugNativeTextInputImePayload.value());
        _debugNativeTextInputImePayload.reset();
        return payload;
    }

    NativeTextInputImePayload payload;
    if (! _hwnd)
    {
        return payload;
    }

    HIMC inputContext = ImmGetContext(_hwnd);
    if (! inputContext)
    {
        return payload;
    }
    const auto releaseContext = wil::scope_exit([&] { ImmReleaseContext(_hwnd, inputContext); });

    if ((compositionFlags & GCS_RESULTSTR) != 0)
    {
        if (std::optional<std::wstring> resultString = ReadImeString(inputContext, GCS_RESULTSTR); resultString.has_value())
        {
            payload.hasResultString = true;
            payload.resultString    = std::move(resultString.value());
        }
    }
    if ((compositionFlags & GCS_COMPSTR) != 0)
    {
        if (std::optional<std::wstring> compositionString = ReadImeString(inputContext, GCS_COMPSTR); compositionString.has_value())
        {
            payload.hasCompositionString = true;
            payload.compositionString    = std::move(compositionString.value());
        }
    }
    if ((compositionFlags & GCS_COMPATTR) != 0)
    {
        payload.compositionAttributes = ReadImeBytes(inputContext, GCS_COMPATTR);
    }
    if ((compositionFlags & GCS_COMPCLAUSE) != 0)
    {
        payload.compositionClauses = ReadImeDwords(inputContext, GCS_COMPCLAUSE);
    }
    if ((compositionFlags & GCS_CURSORPOS) != 0)
    {
        const LONG cursorPosition = ImmGetCompositionStringW(inputContext, GCS_CURSORPOS, nullptr, 0u);
        if (cursorPosition >= 0)
        {
            payload.hasCursorPosition = true;
            payload.cursorPosition    = static_cast<size_t>(cursorPosition);
        }
    }
    return payload;
}

void WindowHost::UpdateNativeTextInputImeWindows() noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    if (! _hwnd || _textInputBackend != TextInputBackend::Native || ! _nativeTextInputControl || ! _nativeTextInputStateCacheValid)
    {
        return;
    }

    HIMC inputContext = ImmGetContext(_hwnd);
    if (! inputContext)
    {
        return;
    }
    const auto releaseContext = wil::scope_exit([&] { ImmReleaseContext(_hwnd, inputContext); });

    D2D1_RECT_F caretRectDip{};
    RECT caretClientRectPx{};
    RECT caretScreenRectPx{};
    if (! TryGetNativeTextInputCaretRects(caretRectDip, caretClientRectPx, caretScreenRectPx))
    {
        return;
    }

    COMPOSITIONFORM compositionForm{};
    compositionForm.dwStyle      = CFS_FORCE_POSITION;
    compositionForm.ptCurrentPos = POINT{caretClientRectPx.left, caretClientRectPx.top};
    static_cast<void>(ImmSetCompositionWindow(inputContext, &compositionForm));

    CANDIDATEFORM candidateForm{};
    candidateForm.dwIndex      = 0u;
    candidateForm.dwStyle      = CFS_EXCLUDE;
    candidateForm.ptCurrentPos = POINT{caretClientRectPx.left, caretClientRectPx.bottom};
    candidateForm.rcArea       = caretClientRectPx;
    static_cast<void>(ImmSetCandidateWindow(inputContext, &candidateForm));

    Debug::Perf::Emit(L"dxui.textinput.candidate_rect_update_us",
                      L"native-ime-window",
                      Debug::Perf::ElapsedUs(startedAt),
                      static_cast<uint64_t>(std::max<LONG>(0, caretClientRectPx.right - caretClientRectPx.left)),
                      static_cast<uint64_t>(std::max<LONG>(0, caretClientRectPx.bottom - caretClientRectPx.top)),
                      S_OK);
}

bool WindowHost::HandleNativeTextInputEditMessage(UINT msg, WPARAM wp, LPARAM lp, LRESULT& outResult) noexcept
{
    outResult = 0;
    if (_textInputBackend != TextInputBackend::Native || ! _focusedControl || ! _focusedControl->SupportsTextInput())
    {
        return false;
    }

    Control* const editTarget = _focusedControl;
    const auto inputStartedAt = std::chrono::steady_clock::now();

    const auto syncImportedState = [this, editTarget](const TextInputState& state, bool notifyChange, bool clearComposition) noexcept -> bool
    {
        if (clearComposition)
        {
            ClearNativeTextInputCompositionState();
        }

        if (! editTarget->ImportTextInputState(*this, state, notifyChange))
        {
            return false;
        }

        if (editTarget == _focusedControl)
        {
            SyncNativeTextInputSession(editTarget);
        }
        return true;
    };

    bool controlHandled = false;
    switch (msg)
    {
        case WM_GETTEXTLENGTH:
        {
            TextInputState state;
            if (! editTarget->ExportTextInputState(state))
            {
                return false;
            }

            outResult = static_cast<LRESULT>(NormalizeWin32EditTextFromControlText(state.text, state.multiline).size());
            return true;
        }
        case WM_GETTEXT:
        {
            if (wp == 0 || lp == 0)
            {
                return true;
            }

            TextInputState state;
            if (! editTarget->ExportTextInputState(state))
            {
                return false;
            }

            const std::wstring win32EditText = NormalizeWin32EditTextFromControlText(state.text, state.multiline);
            const size_t bufferChars         = static_cast<size_t>(wp);
            const size_t copyChars           = (bufferChars > 0u) ? (std::min)(win32EditText.size(), bufferChars - 1u) : 0u;
            auto* buffer                     = reinterpret_cast<wchar_t*>(lp);
            if (copyChars > 0u)
            {
                std::copy_n(win32EditText.data(), copyChars, buffer);
            }
            buffer[copyChars] = L'\0';
            outResult         = static_cast<LRESULT>(copyChars);
            return true;
        }
        case WM_SETTEXT:
        {
            TextInputState state;
            if (! editTarget->ExportTextInputState(state))
            {
                return false;
            }

            const wchar_t* const text = (lp != 0) ? reinterpret_cast<const wchar_t*>(lp) : L"";
            state.text                = NormalizeControlTextFromWin32EditText(text, state.multiline);
            state.selectionAnchorIndex.reset();
            state.caretIndex       = state.text.size();
            state.firstVisibleLine = 0u;
            outResult              = syncImportedState(state, true, true) ? TRUE : FALSE;
            return true;
        }
        case EM_GETSEL:
        {
            TextInputState state;
            if (! editTarget->ExportTextInputState(state))
            {
                return false;
            }

            const auto [selectionStart, selectionEnd] = GetTextInputSelectionRange(state);
            const DWORD startIndex = static_cast<DWORD>((std::min)(MapControlIndexToWin32EditIndex(state.text, selectionStart, state.multiline),
                                                                   static_cast<size_t>(std::numeric_limits<DWORD>::max())));
            const DWORD endIndex   = static_cast<DWORD>(
                (std::min)(MapControlIndexToWin32EditIndex(state.text, selectionEnd, state.multiline), static_cast<size_t>(std::numeric_limits<DWORD>::max())));
            if (wp != 0)
            {
                *reinterpret_cast<DWORD*>(wp) = startIndex;
            }
            if (lp != 0)
            {
                *reinterpret_cast<DWORD*>(lp) = endIndex;
            }

            outResult = MAKELRESULT(static_cast<WORD>(std::min<DWORD>(startIndex, std::numeric_limits<WORD>::max())),
                                    static_cast<WORD>(std::min<DWORD>(endIndex, std::numeric_limits<WORD>::max())));
            return true;
        }
        case EM_SETSEL:
        {
            TextInputState state;
            if (! editTarget->ExportTextInputState(state))
            {
                return false;
            }

            const std::wstring win32EditText = NormalizeWin32EditTextFromControlText(state.text, state.multiline);
            const auto mapSelectionIndex     = [&win32EditText, &state](uint64_t unsignedValue, LPARAM signedValue) noexcept
            {
                constexpr uint64_t kWin32EditEndSentinel = static_cast<uint64_t>(std::numeric_limits<UINT>::max());
                if (signedValue < 0 || unsignedValue == kWin32EditEndSentinel || unsignedValue == static_cast<uint64_t>((std::numeric_limits<WPARAM>::max)()))
                {
                    return state.text.size();
                }

                return MapWin32EditIndexToControlIndex(win32EditText, static_cast<size_t>(unsignedValue), state.multiline);
            };

            SetTextInputSelectionRange(
                state, mapSelectionIndex(static_cast<uint64_t>(wp), static_cast<LPARAM>(wp)), mapSelectionIndex(static_cast<uint64_t>(lp), lp));
            outResult = syncImportedState(state, false, false) ? TRUE : FALSE;
            return true;
        }
        case EM_REPLACESEL:
        {
            TextInputState state;
            if (! editTarget->ExportTextInputState(state))
            {
                return false;
            }

            const wchar_t* const replacementText     = (lp != 0) ? reinterpret_cast<const wchar_t*>(lp) : L"";
            const std::wstring normalizedReplacement = NormalizeControlTextFromWin32EditText(replacementText, state.multiline);
            if (! ReplaceTextInputSelection(state, normalizedReplacement))
            {
                outResult = FALSE;
                return true;
            }

            outResult = syncImportedState(state, true, true) ? TRUE : FALSE;
            return true;
        }
        case WM_COPY:
            SetInputModality(InputModality::Keyboard);
            controlHandled = editTarget->OnCopy(*this);
            break;
        case WM_CUT:
            SetInputModality(InputModality::Keyboard);
            controlHandled = editTarget->OnKeyDown(*this, 'X', MK_CONTROL);
            break;
        case WM_PASTE:
            SetInputModality(InputModality::Keyboard);
            controlHandled = editTarget->OnKeyDown(*this, 'V', MK_CONTROL);
            break;
        case WM_CLEAR:
        {
            SetInputModality(InputModality::Keyboard);
            TextInputState state;
            if (! editTarget->ExportTextInputState(state) || state.readOnly || ! HasTextSelection(state))
            {
                return false;
            }
            controlHandled = editTarget->OnKeyDown(*this, VK_DELETE, 0);
            break;
        }
        case WM_UNDO:
            SetInputModality(InputModality::Keyboard);
            controlHandled = editTarget->OnKeyDown(*this, 'Z', MK_CONTROL);
            break;
        default: return false;
    }

    if (controlHandled && editTarget == _focusedControl)
    {
        SyncNativeTextInputSession(editTarget);
        RecordNativeTextInputKeyToStateMetric(inputStartedAt, L"native-edit-message", static_cast<uint64_t>(msg), 0u, msg != WM_COPY);
        outResult = (msg == WM_UNDO) ? TRUE : 0;
    }
    return controlHandled;
}

bool WindowHost::HandleNativeTextInputImeMessage(UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    static_cast<void>(wp);
    if (_textInputBackend != TextInputBackend::Native || ! _focusedControl || ! _focusedControl->SupportsTextInput())
    {
        return false;
    }

    Control* const editTarget = _focusedControl;
    if (editTarget != _nativeTextInputControl)
    {
        ActivateNativeTextInputSession(editTarget);
    }
    if (! _nativeTextInputStateCacheValid)
    {
        SyncNativeTextInputSession(editTarget);
    }
    if (! _nativeTextInputStateCacheValid)
    {
        return false;
    }

    SetInputModality(InputModality::Keyboard);
    const auto inputStartedAt = std::chrono::steady_clock::now();
    if (_nativeTextInputStateCache.readOnly)
    {
        ClearNativeTextInputCompositionState();
        Debug::Perf::Emit(
            L"dxui.textinput.ime_update_us", L"native-ime-lifecycle", Debug::Perf::ElapsedUs(inputStartedAt), 0u, static_cast<uint64_t>(msg), S_OK);
        return true;
    }

    auto startCompositionFromCurrentRange = [this, editTarget]()
    {
        const NativeTextInputState previousState = _nativeTextInputStateCache;
        TextInputState baseState;
        if (editTarget->ExportTextInputState(baseState))
        {
            _nativeTextInputImeBaseState = baseState;
        }
        else
        {
            _nativeTextInputImeBaseState.reset();
        }

        const TextInputState& rangeState                        = _nativeTextInputImeBaseState.has_value()
                                                                      ? _nativeTextInputImeBaseState.value()
                                                                      : TextInputState{
                                                                            .text                 = _nativeTextInputStateCache.text,
                                                                            .selectionAnchorIndex = _nativeTextInputStateCache.selectionAnchorIndex,
                                                                            .caretIndex           = _nativeTextInputStateCache.caretIndex,
                                                                            .firstVisibleLine     = _nativeTextInputStateCache.firstVisibleLine,
                                                                            .readOnly             = _nativeTextInputStateCache.readOnly,
                                                                            .masked               = _nativeTextInputStateCache.masked,
                                                                            .multiline            = _nativeTextInputStateCache.multiline,
                                                                        };
        const auto [compositionStartIndex, compositionEndIndex] = ResolveImeBaseCompositionRange(rangeState);

        _nativeTextInputImeComposing          = true;
        _nativeTextInputCompositionStartIndex = compositionStartIndex;
        _nativeTextInputCompositionEndIndex   = compositionEndIndex;
        _nativeTextInputConversionTargetStartIndex.reset();
        _nativeTextInputConversionTargetEndIndex.reset();
        _nativeTextInputCompositionCursorIndex.reset();
        _nativeTextInputCompositionClauseBoundaries.clear();
        ApplyNativeTextInputCompositionStateToCache();
        RaiseNativeTextInputAccessibilityEvents(previousState);
    };

    auto ensureCompositionBaseState = [this, editTarget]() -> bool
    {
        if (_nativeTextInputImeBaseState.has_value())
        {
            return true;
        }

        TextInputState baseState;
        if (! editTarget->ExportTextInputState(baseState))
        {
            return false;
        }
        _nativeTextInputImeBaseState = baseState;
        return true;
    };

    auto updateConversionTargetFromAttributes =
        [this](size_t compositionStartIndex, size_t compositionLength, const NativeTextInputImePayload& payload) noexcept
    {
        _nativeTextInputConversionTargetStartIndex.reset();
        _nativeTextInputConversionTargetEndIndex.reset();

        const size_t attributeCount = std::min(payload.compositionAttributes.size(), compositionLength);
        std::optional<size_t> targetStart;
        size_t targetEnd = 0u;
        for (size_t index = 0u; index < attributeCount; ++index)
        {
            if (IsImeTargetAttribute(payload.compositionAttributes[index]))
            {
                if (! targetStart.has_value())
                {
                    targetStart = index;
                }
                targetEnd = index + 1u;
            }
            else if (targetStart.has_value())
            {
                break;
            }
        }

        if (targetStart.has_value())
        {
            _nativeTextInputConversionTargetStartIndex = compositionStartIndex + targetStart.value();
            _nativeTextInputConversionTargetEndIndex   = compositionStartIndex + targetEnd;
        }
    };

    auto applyCompositionPayload =
        [this, editTarget, &ensureCompositionBaseState, &updateConversionTargetFromAttributes](const NativeTextInputImePayload& payload) -> bool
    {
        if (! payload.hasCompositionString || ! ensureCompositionBaseState())
        {
            return false;
        }

        TextInputState previewState                             = _nativeTextInputImeBaseState.value();
        const auto [compositionStartIndex, compositionEndIndex] = ResolveImeBaseCompositionRange(previewState);
        ReplaceNativeTextInputRange(previewState, compositionStartIndex, compositionEndIndex, payload.compositionString);
        if (payload.hasCursorPosition)
        {
            previewState.caretIndex = compositionStartIndex + std::min(payload.cursorPosition, payload.compositionString.size());
        }

        if (! editTarget->ImportTextInputState(*this, previewState, false))
        {
            return false;
        }

        SyncNativeTextInputSession(editTarget);
        const NativeTextInputState previousState = _nativeTextInputStateCache;
        _nativeTextInputImeComposing             = true;
        _nativeTextInputCompositionStartIndex    = compositionStartIndex;
        _nativeTextInputCompositionEndIndex      = compositionStartIndex + payload.compositionString.size();
        updateConversionTargetFromAttributes(compositionStartIndex, payload.compositionString.size(), payload);
        _nativeTextInputCompositionCursorIndex =
            payload.hasCursorPosition ? std::optional<size_t>{compositionStartIndex + std::min(payload.cursorPosition, payload.compositionString.size())}
                                      : std::nullopt;
        _nativeTextInputCompositionClauseBoundaries = ResolveImeClauseBoundaries(compositionStartIndex, payload.compositionString.size(), payload);
        ApplyNativeTextInputCompositionStateToCache();
        RaiseNativeTextInputAccessibilityEvents(previousState);
        UpdateNativeTextInputCaret();
        const uint64_t conversionTargetLength =
            _nativeTextInputConversionTargetStartIndex.has_value() && _nativeTextInputConversionTargetEndIndex.has_value()
                ? static_cast<uint64_t>(_nativeTextInputConversionTargetEndIndex.value() - _nativeTextInputConversionTargetStartIndex.value())
                : 0u;
        Debug::Perf::Emit(L"dxui.textinput.composition_length",
                          L"native-ime-composition",
                          0u,
                          static_cast<uint64_t>(payload.compositionString.size()),
                          conversionTargetLength,
                          S_OK);
        return true;
    };

    auto applyResultPayload = [this, editTarget, &ensureCompositionBaseState](const NativeTextInputImePayload& payload) -> bool
    {
        if (! payload.hasResultString || ! ensureCompositionBaseState())
        {
            return false;
        }

        TextInputState committedState                           = _nativeTextInputImeBaseState.value();
        const auto [compositionStartIndex, compositionEndIndex] = ResolveImeBaseCompositionRange(committedState);
        ReplaceNativeTextInputRange(committedState, compositionStartIndex, compositionEndIndex, payload.resultString);
        if (! editTarget->ImportTextInputState(*this, committedState, true))
        {
            return false;
        }

        SyncNativeTextInputSession(editTarget);
        _nativeTextInputImeBaseState = committedState;
        Debug::Perf::Emit(L"dxui.textinput.composition_length", L"native-ime-result", 0u, static_cast<uint64_t>(payload.resultString.size()), 0u, S_OK);
        return true;
    };

    auto restoreCompositionBaseState = [this, editTarget]()
    {
        if (! _nativeTextInputImeBaseState.has_value())
        {
            return;
        }

        static_cast<void>(editTarget->ImportTextInputState(*this, _nativeTextInputImeBaseState.value(), false));
        SyncNativeTextInputSession(editTarget);
    };

    switch (msg)
    {
        case WM_IME_STARTCOMPOSITION:
            SyncNativeTextInputSession(editTarget);
            startCompositionFromCurrentRange();
            UpdateNativeTextInputImeWindows();
            break;
        case WM_IME_COMPOSITION:
        {
            constexpr LPARAM kCompositionPayloadMask = GCS_COMPSTR | GCS_COMPATTR | GCS_COMPCLAUSE | GCS_CURSORPOS;
            const bool hasResultPayload              = (lp & GCS_RESULTSTR) != 0;
            const bool hasCompositionPayload         = (lp & kCompositionPayloadMask) != 0;
            const NativeTextInputImePayload payload  = ReadNativeTextInputImePayload(lp);
            if (! hasResultPayload && ! hasCompositionPayload && ! _nativeTextInputImeComposing)
            {
                ClearNativeTextInputCompositionState();
            }
            else if (hasResultPayload && ! hasCompositionPayload)
            {
                static_cast<void>(applyResultPayload(payload));
                ClearNativeTextInputCompositionState();
            }
            else
            {
                if (! _nativeTextInputImeComposing)
                {
                    SyncNativeTextInputSession(editTarget);
                    startCompositionFromCurrentRange();
                }
                else
                {
                    ApplyNativeTextInputCompositionStateToCache();
                }
                if (hasResultPayload)
                {
                    static_cast<void>(applyResultPayload(payload));
                    startCompositionFromCurrentRange();
                }
                if (hasCompositionPayload)
                {
                    static_cast<void>(applyCompositionPayload(payload));
                }
                UpdateNativeTextInputImeWindows();
            }
            break;
        }
        case WM_IME_ENDCOMPOSITION:
            if (_nativeTextInputImeComposing)
            {
                restoreCompositionBaseState();
            }
            ClearNativeTextInputCompositionState();
            break;
        default: return false;
    }

    Debug::Perf::Emit(L"dxui.textinput.ime_update_us",
                      L"native-ime-lifecycle",
                      Debug::Perf::ElapsedUs(inputStartedAt),
                      _nativeTextInputImeComposing ? 1u : 0u,
                      static_cast<uint64_t>(msg),
                      S_OK);
    return true;
}

bool WindowHost::TryReadNativeTextInputState(const Control* control, NativeTextInputState& outState) const noexcept
{
    if (! control || control != _nativeTextInputControl || ! _nativeTextInputStateCacheValid)
    {
        return false;
    }

    outState = _nativeTextInputStateCache;
    return true;
}

bool WindowHost::TryGetNativeTextInputCaretRects(D2D1_RECT_F& outRectDip, RECT& outClientRectPx, RECT& outScreenRectPx) const noexcept
{
    if (! _hwnd || ! _nativeTextInputControl || ! _nativeTextInputStateCacheValid)
    {
        return false;
    }

    const std::optional<D2D1_RECT_F> caretRectDip = _nativeTextInputControl->GetTextInputCaretRect(*this, _nativeTextInputStateCache.caretIndex);
    if (! caretRectDip.has_value())
    {
        return false;
    }

    outRectDip      = caretRectDip.value();
    outClientRectPx = NormalizeCaretRectPx(RECT{static_cast<LONG>(std::lround(DipsToPixels(outRectDip.left))),
                                                static_cast<LONG>(std::lround(DipsToPixels(outRectDip.top))),
                                                static_cast<LONG>(std::lround(DipsToPixels(outRectDip.right))),
                                                static_cast<LONG>(std::lround(DipsToPixels(outRectDip.bottom)))});

    outScreenRectPx = outClientRectPx;
    MapWindowPoints(_hwnd, nullptr, reinterpret_cast<POINT*>(&outScreenRectPx), 2);
    return true;
}

void WindowHost::UpdateNativeTextInputCaret() noexcept
{
    D2D1_RECT_F caretRectDip{};
    RECT caretClientRectPx{};
    RECT caretScreenRectPx{};
    if (! TryGetNativeTextInputCaretRects(caretRectDip, caretClientRectPx, caretScreenRectPx))
    {
        DestroyNativeTextInputCaret();
        return;
    }

    _nativeTextInputCaretRectDip      = caretRectDip;
    _nativeTextInputCaretClientRectPx = caretClientRectPx;
    _nativeTextInputCaretScreenRectPx = caretScreenRectPx;
    _nativeTextInputCaretRectValid    = true;
    ++_nativeTextInputEventCounters.caretUpdateCount;

    if (! _hwnd || GetFocus() != _hwnd)
    {
        return;
    }

    const int caretWidthPx  = std::max(1, static_cast<int>(caretClientRectPx.right - caretClientRectPx.left));
    const int caretHeightPx = std::max(1, static_cast<int>(caretClientRectPx.bottom - caretClientRectPx.top));
    const bool recreateCaret =
        ! _nativeTextInputCaret.IsCreated() || caretWidthPx != _nativeTextInputCaret.WidthPx() || caretHeightPx != _nativeTextInputCaret.HeightPx();

    if (recreateCaret && ! _nativeTextInputCaret.Create(_hwnd, caretWidthPx, caretHeightPx))
    {
        return;
    }

    static_cast<void>(SetCaretPos(caretClientRectPx.left, caretClientRectPx.top));
    static_cast<void>(_nativeTextInputCaret.Show());
    if (_nativeTextInputImeComposing)
    {
        UpdateNativeTextInputImeWindows();
    }
}

void WindowHost::DestroyNativeTextInputCaret() noexcept
{
    _nativeTextInputCaret.Reset();
    _nativeTextInputCaretRectValid    = false;
    _nativeTextInputCaretRectDip      = D2D1::RectF();
    _nativeTextInputCaretClientRectPx = RECT{};
    _nativeTextInputCaretScreenRectPx = RECT{};
}

#if defined(ENABLE_TESTS)
bool WindowHost::DebugHasActiveNativeTextInputSession() const noexcept
{
    return HasActiveNativeTextInputSession();
}

bool WindowHost::DebugHasActiveNativeTextInputTsfDocument() const noexcept
{
    return _nativeTextInputTsfActive && _nativeTextInputTsfThreadMgr && _nativeTextInputTsfDocumentMgr && _nativeTextInputTsfContext &&
           _nativeTextInputTsfTextStore;
}

bool WindowHost::DebugGetNativeTextInputState(NativeTextInputState& outState) const noexcept
{
    if (! _nativeTextInputStateCacheValid)
    {
        return false;
    }

    outState = _nativeTextInputStateCache;
    return true;
}

bool WindowHost::DebugGetNativeTextInputCaretRect(D2D1_RECT_F& outRectDip, RECT& outScreenRectPx) const noexcept
{
    if (! _nativeTextInputCaretRectValid)
    {
        return false;
    }

    outRectDip      = _nativeTextInputCaretRectDip;
    outScreenRectPx = _nativeTextInputCaretScreenRectPx;
    return true;
}

NativeTextInputEventCounters WindowHost::DebugGetNativeTextInputEventCounters() const noexcept
{
    return _nativeTextInputEventCounters;
}

void WindowHost::DebugSetNativeTextInputImePayloadForTest(NativeTextInputImePayload payload)
{
    _debugNativeTextInputImePayload = std::move(payload);
}
#endif
} // namespace RedSalamander::DxUi
