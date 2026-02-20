#include "ColorTextView.h"
#include "resource.h"

#include "WindowMessages.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cmath>
#include <cwctype>
#include <fstream>
#include <iterator>
#include <memory>
#include <thread>
#include <utility>

#pragma warning(push)
#pragma warning(                                                                                                                                               \
    disable : 4625 4626 5026 5027) // WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#include <wil/resource.h>
#include <wil/result_macros.h>
#pragma warning(pop)

#ifndef GET_X_LPARAM
#define GET_X_LPARAM(lp) ((int)(short)LOWORD(lp))
#define GET_Y_LPARAM(lp) ((int)(short)HIWORD(lp))
#endif

#ifdef _DEBUG
// Debug background colors for spans - array of distinct colors
static const D2D1_COLOR_F debugColors[] = {
    D2D1::ColorF(1.0f, 0.8f, 0.8f, 0.3f), // Light red
    D2D1::ColorF(0.8f, 1.0f, 0.8f, 0.3f), // Light green
    D2D1::ColorF(0.8f, 0.8f, 1.0f, 0.3f), // Light blue
    D2D1::ColorF(1.0f, 1.0f, 0.8f, 0.3f), // Light yellow
    D2D1::ColorF(1.0f, 0.8f, 1.0f, 0.3f), // Light magenta
    D2D1::ColorF(0.8f, 1.0f, 1.0f, 0.3f), // Light cyan
    D2D1::ColorF(1.0f, 0.9f, 0.8f, 0.3f), // Light orange
    D2D1::ColorF(0.9f, 0.8f, 1.0f, 0.3f), // Light purple
};
static constexpr size_t debugColorCount = sizeof(debugColors) / sizeof(debugColors[0]);

static const D2D1_COLOR_F debugDirtyPalette[] = {
    D2D1::ColorF(D2D1::ColorF::LawnGreen, 0.35f),
    D2D1::ColorF(D2D1::ColorF::Orange, 0.35f),
    D2D1::ColorF(D2D1::ColorF::RoyalBlue, 0.35f),
    D2D1::ColorF(D2D1::ColorF::HotPink, 0.35f),
    D2D1::ColorF(D2D1::ColorF::SpringGreen, 0.35f),
    D2D1::ColorF(D2D1::ColorF::Tomato, 0.35f),
    D2D1::ColorF(D2D1::ColorF::MediumPurple, 0.35f),
    D2D1::ColorF(D2D1::ColorF::DeepSkyBlue, 0.35f),
    D2D1::ColorF(D2D1::ColorF::SandyBrown, 0.35f),
    D2D1::ColorF(D2D1::ColorF::Aquamarine, 0.35f),
    D2D1::ColorF(D2D1::ColorF::Firebrick, 0.35f),
    D2D1::ColorF(D2D1::ColorF::Gold, 0.35f),
};
static constexpr size_t debugDirtyPaletteSize = sizeof(debugDirtyPalette) / sizeof(debugDirtyPalette[0]);
#endif

// ===== Internal Helpers =====
static inline bool IsKeyDown(int vk)
{
    return (GetKeyState(vk) & 0x8000) != 0;
}

static LRESULT SafeCallWindowProcW(WNDPROC proc, HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    __try
    {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
        return CallWindowProcW(proc, hwnd, msg, wp, lp);
#pragma warning(pop)
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        return 0;
    }
}

static size_t FindCaseIncensitive(const std::wstring& hay, const std::wstring& needle, size_t from)
{
    if (needle.empty())
        return std::wstring::npos;

    // Pre-compute lowercase needle
    std::wstring lowerNeedle;
    lowerNeedle.reserve(needle.size());
    for (wchar_t ch : needle)
        lowerNeedle.push_back(towlower(ch));

    auto hayBegin = hay.begin() + static_cast<__int64>(from);
    auto hayEnd   = hay.end();

    auto it = std::search(hayBegin,
                          hayEnd,
                          lowerNeedle.begin(),
                          lowerNeedle.end(),
                          [](wchar_t hayChar, wchar_t needleChar)
                          {
                              return towlower(hayChar) == needleChar; // needleChar is already lowercase
                          });

    return (it == hayEnd) ? std::wstring::npos : (it - hay.begin());
}

// Returns the height of a horizontal scrollbar in DIPs for this window DPI
static float GetHorzScrollbarDip([[maybe_unused]] HWND hwnd, float dpi)
{
    UINT udpi    = static_cast<UINT>(dpi > 0 ? dpi : 96.0f);
    const int cy = GetSystemMetricsForDpi(SM_CYHSCROLL, udpi);
    return static_cast<float>(cy) * 96.f / static_cast<float>(udpi);
}

// ===== Public =====
ColorTextView::ColorTextView()
{
    // wil::critical_section (_etwQueueCS) initializes automatically via RAII
}

ColorTextView::~ColorTextView()
{
    // Clear atomic HWND first so ETW worker thread stops posting messages
    _hWndAtomic.store(nullptr, std::memory_order_release);
    // wil::critical_section (_etwQueueCS) cleans up automatically via RAII
}

ATOM ColorTextView::RegisterWndClass(HINSTANCE hinst)
{
    static ATOM s_atom = 0;
    if (s_atom)
        return s_atom;

    WNDCLASS wc{};
    wc.hInstance     = hinst;
    wc.lpfnWndProc   = &ColorTextView::WndProcThunk;
    wc.lpszClassName = L"ColorTextView";
    wc.hCursor       = LoadCursor(nullptr, IDC_IBEAM);
    wc.hbrBackground = nullptr; // we handle background
    s_atom           = RegisterClass(&wc);
    return s_atom;
}
HWND ColorTextView::Create(HWND parent, int x, int y, int w, int h)
{
    HINSTANCE hinst = (HINSTANCE)GetWindowLongPtr(parent, GWLP_HINSTANCE);
    RegisterWndClass(hinst);
    // Add scrollbar styles
    _hWnd = CreateWindowEx(
        0, L"ColorTextView", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_CLIPSIBLINGS | WS_VSCROLL | WS_HSCROLL, x, y, w, h, parent, nullptr, hinst, this);
    _hWndAtomic.store(_hWnd, std::memory_order_release);
    return _hWnd;
}

void ColorTextView::SetTheme(const Theme& t)
{
    // Validate theme colors and replace invalid ones with safe defaults
    auto isValidColor = [](const D2D1_COLOR_F& c)
    {
        return ! std::isnan(c.r) && ! std::isnan(c.g) && ! std::isnan(c.b) && ! std::isnan(c.a) && ! std::isinf(c.r) && ! std::isinf(c.g) &&
               ! std::isinf(c.b) && ! std::isinf(c.a);
    };

    Theme validTheme = t;

    // Replace invalid colors with safe defaults and log warnings
    if (! isValidColor(t.fg))
    {
#ifdef _DEBUG
        OutputDebugStringA("!!! SetTheme: Invalid foreground color, fallback to white\n");
#endif
        validTheme.fg = D2D1::ColorF(D2D1::ColorF::White);
    }
    if (! isValidColor(t.bg))
    {
#ifdef _DEBUG
        OutputDebugStringA("!!! SetTheme: Invalid background color, fallback to black\n");
#endif
        validTheme.bg = D2D1::ColorF(D2D1::ColorF::Black);
    }
    if (! isValidColor(t.selection))
        validTheme.selection = D2D1::ColorF(D2D1::ColorF::DodgerBlue, 0.5f);
    if (! isValidColor(t.caret))
        validTheme.caret = D2D1::ColorF(D2D1::ColorF::White);
    if (! isValidColor(t.gutterBg))
        validTheme.gutterBg = D2D1::ColorF(0.1f, 0.1f, 0.1f, 1.0f);
    if (! isValidColor(t.gutterFg))
        validTheme.gutterFg = D2D1::ColorF(D2D1::ColorF::Gray);
    if (! isValidColor(t.searchHighlight))
        validTheme.searchHighlight = D2D1::ColorF(D2D1::ColorF::Yellow, 0.5f);
    if (! isValidColor(t.metaError))
        validTheme.metaError = D2D1::ColorF(D2D1::ColorF::Red);
    if (! isValidColor(t.metaWarning))
        validTheme.metaWarning = D2D1::ColorF(D2D1::ColorF::Orange);
    if (! isValidColor(t.metaInfo))
        validTheme.metaInfo = D2D1::ColorF(D2D1::ColorF::Cyan);
    if (! isValidColor(t.metaDebug))
        validTheme.metaDebug = D2D1::ColorF(D2D1::ColorF::Gray);
    if (! isValidColor(t.metaText))
        validTheme.metaText = D2D1::ColorF(D2D1::ColorF::White);

    _theme = validTheme;

    // Clear brush cache to force recreation with validated colors
    _brushCache.clear();
    _brushAccessCounter = 0;

    // Optimization #7 - Pre-create all theme brushes to eliminate first-frame jank
    if (_d2dCtx)
    {
        GetBrush(_theme.bg);
        GetBrush(_theme.fg);
        GetBrush(_theme.caret);
        GetBrush(_theme.selection);
        GetBrush(_theme.searchHighlight);
        GetBrush(_theme.gutterBg);
        GetBrush(_theme.gutterFg);
        GetBrush(_theme.metaText);
        GetBrush(_theme.metaError);
        GetBrush(_theme.metaWarning);
        GetBrush(_theme.metaInfo);
        GetBrush(_theme.metaDebug);
    }

    ApplyColoringToLayout();
    ApplyColoringToTailLayout();
    UpdateFindBarTheme();
    InvalidateSliceBitmap();
    Invalidate();
}

void ColorTextView::EnableLineNumbers(bool enable)
{
    _displayLineNumbers = enable;
    // Layout width and transform change; cached slices are not reusable.
    _layoutCache.clear();
    InvalidateSliceBitmap();
    EnsureLayoutAsync();
    Invalidate();
}

void ColorTextView::EnableShowIds(bool enable)
{
    _document.EnableShowIds(enable);

    const UINT32 totalLen = static_cast<UINT32>(_document.TotalLength());
    _selStart             = std::min(_selStart, totalLen);
    _selEnd               = std::min(_selEnd, totalLen);
    _caretPos             = std::min(_caretPos, totalLen);

    // Prefix text changed for every line: invalidate all layouts and caches.
    _textLayout.reset();
    _tailLayout.reset();
    _fallbackLayout.reset();
    _layoutCache.clear();
    _sliceFilteredRuns.clear();
    _fallbackFilteredRuns.clear();
    _lineMetrics.clear();
    _tailLayoutValid      = false;
    _fallbackValid        = false;
    _sliceFirstLine       = 0;
    _sliceLastLine        = 0;
    _sliceFirstDisplayRow = 0;
    _sliceIsFiltered      = false;
    _sliceStartPos        = 0;
    _sliceEndPos          = 0;

    InvalidateSliceBitmap();
    RequestFullRedraw();

    EnsureWidthAsync();
    if (_renderMode == RenderMode::AUTO_SCROLL)
    {
        RebuildTailLayout();
    }
    else
    {
        EnsureLayoutAsync();
    }

    Invalidate();
}

void ColorTextView::SetFilterMask(uint32_t mask)
{
    // Try to preserve viewport context: find currently visible line before filter changes
    size_t anchorLine      = 0;
    const float lineHeight = GetLineHeight();
    if (_document.TotalLineCount() > 0 && lineHeight > 0.f)
    {
        const float viewTop        = std::max(0.f, _scrollY - _padding);
        const UINT32 topDisplayRow = static_cast<UINT32>(std::floor(viewTop / lineHeight));
        const size_t topVisIdx     = _document.VisibleIndexFromDisplayRow(topDisplayRow);
        if (topVisIdx < _document.VisibleLines().size())
            anchorLine = _document.VisibleLines()[topVisIdx].sourceIndex;
    }

    _document.SetFilterMask(mask);

    // Recalculate content height based on new visible line count
    const UINT32 displayRows = _document.TotalDisplayRows();
    _contentHeight           = (float)displayRows * GetLineHeight() + _padding * 2.f;

    // Adjust scroll position to keep anchor line in view if it's still visible
    if (_document.TotalLineCount() > 0 && anchorLine < _document.TotalLineCount())
    {
        if (_document.IsLineVisible(anchorLine))
        {
            // Anchor line still visible - scroll to keep it at same position
            const UINT32 newDisplayRow = _document.DisplayRowForSource(anchorLine);
            _scrollY                   = static_cast<float>(newDisplayRow) * lineHeight + _padding;
        }
        else
        {
            // Anchor line filtered out - find closest visible line
            size_t closestVisible = anchorLine;

            // Search forward first
            bool foundForward = false;
            for (size_t i = anchorLine; i < _document.TotalLineCount(); ++i)
            {
                if (_document.IsLineVisible(i))
                {
                    closestVisible = i;
                    foundForward   = true;
                    break;
                }
            }

            // If not found forward, search backward
            if (! foundForward && anchorLine > 0)
            {
                for (size_t i = anchorLine; i > 0; --i)
                {
                    if (_document.IsLineVisible(i - 1))
                    {
                        closestVisible = i - 1;
                        break;
                    }
                }
            }

            // Scroll to closest visible line
            const UINT32 newDisplayRow = _document.DisplayRowForSource(closestVisible);
            _scrollY                   = static_cast<float>(newDisplayRow) * lineHeight + _padding;
        }
    }

#ifdef _DEBUG
    auto msg = std::format("SetFilterMask: mask=0x{:02X}, displayRows={}, contentHeight={:.1f}, anchorLine={}, scrollY={:.1f}\n",
                           mask,
                           displayRows,
                           _contentHeight,
                           anchorLine,
                           _scrollY);
    OutputDebugStringA(msg.c_str());
#endif

    // Clamp scroll position to new valid range
    ClampScroll();

    // Update scrollbars to reflect new content height
    UpdateScrollBars();

    // CRITICAL: Clear ALL cached layouts - they contain ALL lines, not just visible ones
    _textLayout.reset();
    _tailLayout.reset();
    _fallbackLayout.reset();
    _layoutCache.clear();
    _sliceFirstLine  = 0;
    _sliceLastLine   = 0;
    _sliceStartPos   = 0;
    _sliceEndPos     = 0;
    _sliceIsFiltered = false;
    _sliceFilteredRuns.clear();
    _fallbackFilteredRuns.clear();

    // Invalidate cached state
    InvalidateSliceBitmap();
    _tailLayoutValid = false;
    _fallbackValid   = false;

    // Force immediate re-layout and redraw
    EnsureLayoutAsync();
    RequestFullRedraw();
    Invalidate();
}

size_t ColorTextView::GetVisibleLineCount() const
{
    return _document.VisibleLineCount();
}

void ColorTextView::SetFont(const wchar_t* family, float sizeDips)
{
    // Minimal: recreate formats
    if (! _dwriteFactory)
        CreateDeviceIndependentResources();
    _textFormat.reset();
    _dwriteFactory->CreateTextFormat(family ? family : L"Segoe UI",
                                     nullptr,
                                     DWRITE_FONT_WEIGHT_NORMAL,
                                     DWRITE_FONT_STYLE_NORMAL,
                                     DWRITE_FONT_STRETCH_NORMAL,
                                     sizeDips > 2 ? sizeDips : 16.f,
                                     L"en-us",
                                     _textFormat.addressof());
    _textFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
    _textFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
    // Ensure uniform line spacing so line height is predictable
    _fontSize = sizeDips > 2 ? sizeDips : 12.f;
    _textFormat->SetLineSpacing(DWRITE_LINE_SPACING_METHOD_UNIFORM, _fontSize + _linePaddingTop + _linePaddingBottom, _fontSize * 0.8f + _linePaddingTop);
    _avgCharWidthValid = false;

    // Text format changed: cached layouts are not reusable.
    _textLayout.reset();
    _tailLayout.reset();
    _fallbackLayout.reset();
    _layoutCache.clear();
    _sliceFilteredRuns.clear();
    _fallbackFilteredRuns.clear();
    _lineMetrics.clear();
    _tailLayoutValid      = false;
    _fallbackValid        = false;
    _sliceFirstLine       = 0;
    _sliceLastLine        = 0;
    _sliceFirstDisplayRow = 0;
    _sliceIsFiltered      = false;
    _sliceStartPos        = 0;
    _sliceEndPos          = 0;

    _document.MarkAllDirty();
    _lineWidthCache.assign(_document.TotalLineCount(), 0.f);
    _maxMeasuredWidth = 0.f;
    _maxMeasuredIndex = 0;
    // Recompute approximate content width and schedule precise measurement
    const size_t maxLen = _document.LongestLineChars();
    _approxContentWidth = GetAverageCharWidth() * static_cast<float>(maxLen);
    EnsureWidthAsync();
    if (_renderMode == RenderMode::AUTO_SCROLL)
    {
        RebuildTailLayout();
    }
    else
    {
        EnsureLayoutAsync();
    }
    InvalidateSliceBitmap();
    Invalidate();
}

void ColorTextView::SetText(const std::wstring& text)
{
    _document.SetText(text);
    _lineWidthCache.assign(_document.TotalLineCount(), 0.f);
    _maxMeasuredWidth = 0.f;
    _maxMeasuredIndex = 0;
    _selStart         = 0;
    _selEnd           = 0;
    _caretPos         = 0;
    _scrollY          = 0.0f;
    _matches.clear();
    _matchIndex = -1;

    _textLayout.reset();
    _tailLayout.reset();
    _fallbackLayout.reset();
    _layoutCache.clear();
    _sliceFilteredRuns.clear();
    _fallbackFilteredRuns.clear();
    _lineMetrics.clear();
    _tailLayoutValid      = false;
    _fallbackValid        = false;
    _sliceFirstLine       = 0;
    _sliceLastLine        = 0;
    _sliceFirstDisplayRow = 0;
    _sliceIsFiltered      = false;
    _sliceStartPos        = 0;
    _sliceEndPos          = 0;
    // Update approximate content width
    const size_t maxLen = _document.LongestLineChars();
    _approxContentWidth = GetAverageCharWidth() * static_cast<float>(maxLen);

    const UINT32 displayRows = _document.TotalDisplayRows();
    _contentHeight           = (float)displayRows * GetLineHeight() + _padding * 2.f;

    UpdateGutterWidth();
    EnsureLayoutAsync();
    EnsureWidthAsync();
    InvalidateSliceBitmap();
    Invalidate();
}
void ColorTextView::AppendText(const std::wstring& more)
{
    const size_t prevLineCount = _document.TotalLineCount();
    _document.AppendText(more);
    if (_lineWidthCache.size() != _document.TotalLineCount())
        _lineWidthCache.resize(_document.TotalLineCount(), 0.f);
    // Recompute approximate width simply
    const size_t maxLen = _document.LongestLineChars();
    _approxContentWidth = GetAverageCharWidth() * static_cast<float>(maxLen);

    const UINT32 displayRows = _document.TotalDisplayRows();
    _contentHeight           = (float)displayRows * GetLineHeight() + _padding * 2.f;

    UpdateGutterWidth();
    // Use adaptive timing based on how many lines were added
    const size_t linesAdded = _document.TotalLineCount() - prevLineCount;
    EnsureLayoutAdaptive(linesAdded);
    EnsureWidthAsync();
    InvalidateSliceBitmap();
    Invalidate();

    // Scroll to bottom if in auto-scroll mode
    if (_renderMode == RenderMode::AUTO_SCROLL)
        RequestScrollToBottom();
}

void ColorTextView::QueueEtwEvent(const Debug::InfoParam& info, const std::wstring& message)
{
    // Use atomic HWND for thread-safe cross-thread access (called from ETW worker thread)
    const HWND hwnd = _hWndAtomic.load(std::memory_order_acquire);
    if (! hwnd)
        return;

    bool shouldPost = false;
    {
        auto lock           = _etwQueueCS.lock();
        const bool wasEmpty = _etwEventQueue.empty();
        _etwEventQueue.push_back({info, message});
        // Only post message if queue was empty (prevents flooding)
        shouldPost = wasEmpty;
    }

    // Post message outside critical section to avoid potential deadlock
    if (shouldPost)
    {
        PostMessage(hwnd, WndMsg::kColorTextViewEtwBatch, 0, 0);
    }
}

void ColorTextView::AppendInfoLine(const Debug::InfoParam& info, const std::wstring& text, bool deferInvalidation)
{
    // Call appendInfoLine which acquires unique_lock internally
    // Then call other methods AFTER it returns (lock is released)
    // This avoids deadlock from trying to acquire shared_lock while holding unique_lock
    _document.AppendInfoLine(text, info);

    // When batching (deferInvalidation=true), skip per-event queries entirely.
    // The caller (OnAppEtwBatch) will query once after the entire batch is processed.
    // This avoids 3 lock acquisitions per event (TotalLineCount, LongestLineChars, TotalDisplayRows).
    if (deferInvalidation)
        return;

    // Get values AFTER appending (now lock is released)
    const size_t newLineCount = _document.TotalLineCount();
    if (_lineWidthCache.size() != newLineCount)
        _lineWidthCache.resize(newLineCount, 0.f);

    // Recompute approximate width simply
    const size_t maxLen = _document.LongestLineChars();
    _approxContentWidth = GetAverageCharWidth() * static_cast<float>(maxLen);

    const UINT32 displayRows = _document.TotalDisplayRows();
    _contentHeight           = (float)displayRows * GetLineHeight() + _padding * 2.f;

    UpdateGutterWidth();

    // Two-mode rendering: choose hot path or cold path
    if (ShouldUseAutoScrollMode())
    {
        // HOT PATH: AUTO-SCROLL mode
        // Fast synchronous tail layout update - no virtualization overhead
        if (_renderMode != RenderMode::AUTO_SCROLL)
        {
            SwitchToAutoScrollMode();
        }
        RebuildTailLayout();
        // No slice invalidation, no async workers, no bitmap caching
    }
    else
    {
        // COLD PATH: SCROLL-BACK mode
        // Full virtualization with async workers and bitmap caching
        if (_renderMode != RenderMode::SCROLL_BACK)
        {
            SwitchToScrollBackMode();
        }
        EnsureLayoutAdaptive(1);
        InvalidateSliceBitmap();
    }

    EnsureWidthAsync();
    Invalidate();
}

void ColorTextView::BeginBatchAppend()
{
    // Signal entering batch mode - could pause timers/workers in future
}

void ColorTextView::EndBatchAppend()
{
    // Finish batch: perform all deferred updates once
    UpdateGutterWidth();

    if (ShouldUseAutoScrollMode())
    {
        if (_renderMode != RenderMode::AUTO_SCROLL)
            SwitchToAutoScrollMode();
        RebuildTailLayout();
    }
    else
    {
        if (_renderMode != RenderMode::SCROLL_BACK)
            SwitchToScrollBackMode();
        EnsureLayoutAdaptive(1);
        InvalidateSliceBitmap();
    }

    EnsureWidthAsync();
    Invalidate();
}

void ColorTextView::ClearText()
{
    _document.Clear();
    _lineWidthCache.clear();
    _maxMeasuredWidth = 0.f;
    _maxMeasuredIndex = 0;
    _matches.clear();
    _matchIndex = -1;
    _textLayout.reset();
    _tailLayout.reset();
    _fallbackLayout.reset();
    _layoutCache.clear();
    _sliceFilteredRuns.clear();
    _fallbackFilteredRuns.clear();
    _lineMetrics.clear();
    _tailLayoutValid      = false;
    _fallbackValid        = false;
    _scrollY              = 0.0f;
    _contentHeight        = 0;
    _approxContentWidth   = 0;
    _sliceFirstLine       = 0;
    _sliceLastLine        = 0;
    _sliceFirstDisplayRow = 0;
    _sliceIsFiltered      = false;
    _sliceStartPos        = 0;
    _sliceEndPos          = 0;
    UpdateGutterWidth();
    InvalidateSliceBitmap();
    Invalidate();
}

void ColorTextView::AddColorRange(UINT32 start, UINT32 length, const D2D1_COLOR_F& color)
{
    if (start >= _document.TotalLength() || ! length)
        return;
    length = std::min<UINT32>(length, (UINT32)_document.TotalLength() - start);
    _document.AddColorRange(start, length, color);
    ApplyColoringToLayout();
    InvalidateSliceBitmap();
    Invalidate();
}

void ColorTextView::ColorizeWord(const std::wstring& word, const D2D1_COLOR_F& color, bool caseSensitive)
{
    if (word.empty() || _document.TotalLineCount() == 0)
        return;

    // per-line search
    UINT32 offset = 0;
    std::vector<Line::ColorSpan> ranges;
    ranges.reserve(64);
    for (size_t i = 0; i < _document.TotalLineCount(); ++i)
    {
        const auto& line = _document.GetSourceLine(i);
        size_t pos       = 0;
        while (true)
        {
            pos = caseSensitive ? line.text.find(word, pos) : FindCaseIncensitive(line.text, word, pos);
            if (pos == std::wstring::npos)
                break;
            // account for metadata prefix length
            const UINT32 plen = _document.PrefixLength(line);
            ranges.push_back(Line::ColorSpan{offset + plen + static_cast<UINT32>(pos), static_cast<UINT32>(word.size()), color});
            pos += word.size();
        }
        offset += _document.PrefixLength(line) + (UINT32)line.text.size() + 1;
    }
    for (const auto& r : ranges)
        _document.AddColorRange(r.start, r.length, r.color);
    ApplyColoringToLayout();
    InvalidateSliceBitmap();
    Invalidate();
}

void ColorTextView::ClearColoring()
{
    _document.ClearColoring();
    ApplyColoringToLayout();
    InvalidateSliceBitmap();
    Invalidate();
}

std::wstring ColorTextView::GetText() const
{
    const auto len = static_cast<UINT32>(_document.TotalLength());
    return _document.GetTextRange(0, len);
}

void ColorTextView::SetAutoScroll(bool enabled)
{
    // _renderMode is the single source of truth - just switch modes
    if (enabled)
    {
        if (_renderMode != RenderMode::AUTO_SCROLL)
            SwitchToAutoScrollMode();
    }
    else
    {
        if (_renderMode != RenderMode::SCROLL_BACK)
            SwitchToScrollBackMode();
    }
}

bool ColorTextView::GetAutoScroll() const
{
    // Auto-scroll is ON when in AUTO_SCROLL mode
    return _renderMode == RenderMode::AUTO_SCROLL;
}

bool ColorTextView::SaveTextToFile(const std::wstring& path) const
{
    return _document.SaveTextToFile(path);
}

void ColorTextView::CopySelection()
{
    CopySelectionToClipboard();
}

void ColorTextView::SetSearchQuery(const std::wstring& q, bool caseSensitive)
{
    _search              = q;
    _searchCaseSensitive = caseSensitive;
    RebuildMatches();
    Invalidate();
}

void ColorTextView::ShowFind()
{
    ShowFindBar();
}

void ColorTextView::GoToEnd(bool enableAutoScroll)
{
    if (_document.TotalLineCount() == 0)
        return;

    _caretPos = static_cast<UINT32>(_document.TotalLength());
    _selStart = _selEnd = _caretPos;

    if (enableAutoScroll)
    {
        SetAutoScroll(true);
    }
    else
    {
        EnsureCaretVisible();
    }

    _caretBlinkOn = true;
    Invalidate();
}
void ColorTextView::FindNext(bool backward)
{
    if (_matches.empty())
        return;

    if (_matchIndex >= 0)
    {
        const auto size = static_cast<__int64>(_matches.size());
        _matchIndex     = backward ? (_matchIndex - 1 + size) % size : (_matchIndex + 1) % size;
    }
    else
    {
        UINT32 anchor = _caretPos;
        switch (_findStartMode)
        {
            case FindStartMode::Top: anchor = 0; break;
            case FindStartMode::Bottom: anchor = static_cast<UINT32>(_document.TotalLength()); break;
            case FindStartMode::CurrentPosition:
            {
                if (_hasFocus)
                {
                    anchor = _caretPos;
                }
                else
                {
                    const auto vr = GetVisibleTextRange();
                    anchor        = backward ? vr.second : vr.first;
                }
                break;
            }
            default: anchor = _caretPos; break;
        }

        if (! backward)
        {
            auto it = std::lower_bound(_matches.begin(), _matches.end(), anchor, [](const Line::ColorSpan& span, UINT32 value) { return span.start < value; });
            if (it != _matches.begin())
            {
                const auto prev   = it - 1;
                const UINT32 end  = prev->start + prev->length;
                const bool inside = (anchor > prev->start) && (anchor < end);
                if (inside)
                    it = prev;
            }
            if (it == _matches.end())
                it = _matches.begin(); // wrap

            _matchIndex = static_cast<__int64>(std::distance(_matches.begin(), it));
        }
        else
        {
            auto it = std::upper_bound(_matches.begin(), _matches.end(), anchor, [](UINT32 value, const Line::ColorSpan& span) { return value < span.start; });
            if (it == _matches.begin())
                it = _matches.end(); // wrap
            --it;

            _matchIndex = static_cast<__int64>(std::distance(_matches.begin(), it));
        }
    }

    auto r    = _matches[static_cast<size_t>(_matchIndex)];
    _selStart = r.start;
    _selEnd   = r.start + r.length;
    _caretPos = _selEnd;
    EnsureCaretVisible();
    _caretBlinkOn = true;
    Invalidate();
}

void ColorTextView::UpdateScrollBars()
{
    if (! _hWnd)
        return;

    RECT clientRect;
    GetClientRect(_hWnd, &clientRect);
    float clientWidth  = static_cast<float>(clientRect.right - clientRect.left) * 96.f / static_cast<float>(_dpi);
    float clientHeight = static_cast<float>(clientRect.bottom - clientRect.top) * 96.f / static_cast<float>(_dpi);

    // Decide vertical scrollbar visibility first (page excludes horizontal scrollbar if visible)
    SCROLLINFO si{};
    si.cbSize         = sizeof(SCROLLINFO);
    si.fMask          = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin           = 0;
    si.nMax           = (int)(_contentHeight);
    float vertPageDip = clientHeight; // provisional, may subtract H-scroll later
    // We'll refine after deciding horizontal visibility below
    si.nPage = (UINT)std::max(0.f, vertPageDip);
    si.nPos  = (int)_scrollY;
    SetScrollInfo(_hWnd, SB_VERT, &si, TRUE);

    // Decide horizontal scrollbar visibility based on content width
    float availableWidth = clientWidth - (_padding * 2.f + (_displayLineNumbers ? _gutterDipW : 0.f));
    // If vertical scrollbar will be visible, it consumes width
    const bool vertVisible = _contentHeight > vertPageDip + 0.5f;
    if (vertVisible)
        availableWidth -= static_cast<float>(GetSystemMetricsForDpi(SM_CXVSCROLL, static_cast<UINT>(_dpi)) * 96.0 / _dpi);
    float contentWidth    = std::max(0.f, _approxContentWidth);
    const bool wantHorz   = contentWidth > availableWidth;
    _horzScrollbarVisible = wantHorz;

    // Now finalize vertical page after knowing horizontal visibility
    vertPageDip = clientHeight - (_horzScrollbarVisible ? GetHorzScrollbarDip(_hWnd, _dpi) : 0.f);
    si.nPage    = (UINT)std::max(0.f, vertPageDip);
    si.fMask    = SIF_PAGE | SIF_POS | SIF_RANGE;
    SetScrollInfo(_hWnd, SB_VERT, &si, TRUE);
    ShowScrollBar(_hWnd, SB_VERT, _contentHeight > vertPageDip + 0.5f);

    // Update horizontal scrollbar
    si.nMax  = (int)contentWidth;
    si.nPage = (UINT)availableWidth;
    si.nPos  = (int)_scrollX;
    SetScrollInfo(_hWnd, SB_HORZ, &si, TRUE);
    ShowScrollBar(_hWnd, SB_HORZ, _horzScrollbarVisible);
}

void ColorTextView::OnVScroll(UINT code, [[maybe_unused]] UINT pos)
{
    SCROLLINFO si{};
    si.cbSize = sizeof(SCROLLINFO);
    si.fMask  = SIF_ALL;
    GetScrollInfo(_hWnd, SB_VERT, &si);

    float oldScrollY         = _scrollY;
    bool userRequestedBottom = false;

    switch (code)
    {
        case SB_LINEUP: ScrollBy(-GetLineHeight()); break;
        case SB_LINEDOWN: ScrollBy(GetLineHeight()); break;
        case SB_PAGEUP: ScrollBy(-_clientDipH * 0.9f); break;
        case SB_PAGEDOWN: ScrollBy(+_clientDipH * 0.9f); break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: ScrollTo(static_cast<float>(si.nTrackPos)); break;
        case SB_TOP: ScrollTo(0.f); break;
        case SB_BOTTOM:
            userRequestedBottom = true;
            ScrollTo(_contentHeight);
            break;
    }

    if (oldScrollY != _scrollY)
    {
        // Handle auto-scroll state changes based on user action
        if (userRequestedBottom)
        {
            // User jumped to bottom (End key or SB_BOTTOM) -> enable auto-scroll
            if (_renderMode != RenderMode::AUTO_SCROLL)
            {
#ifdef _DEBUG
                OutputDebugStringA("OnVScroll: User jumped to BOTTOM, enabling auto-scroll\n");
#endif
                SwitchToAutoScrollMode();
            }
        }
        else if (_scrollY < oldScrollY)
        {
            // User scrolled UP -> disable auto-scroll
            if (_renderMode == RenderMode::AUTO_SCROLL)
            {
#ifdef _DEBUG
                OutputDebugStringA("OnVScroll: User scrolled UP, disabling auto-scroll\n");
#endif
                SwitchToScrollBackMode();
            }
        }

        UpdateScrollBars();
        MaybeRefreshVirtualSliceOnScroll();

        // Two-mode transition: check if we should switch modes based on scroll position
        if (ShouldUseAutoScrollMode())
        {
            if (_renderMode != RenderMode::AUTO_SCROLL)
            {
#ifdef _DEBUG
                OutputDebugStringA("OnVScroll: Switching to AUTO_SCROLL mode\n");
#endif
                SwitchToAutoScrollMode();
            }
        }
        else
        {
            if (_renderMode != RenderMode::SCROLL_BACK)
            {
#ifdef _DEBUG
                {
                    auto msg = std::format("OnVScroll: Switching to SCROLL_BACK mode, scrollY={:.1f}, contentHeight={:.1f}\n", _scrollY, _contentHeight);
                    OutputDebugStringA(msg.c_str());
                }
#endif
                SwitchToScrollBackMode();
            }
        }
    }
}
void ColorTextView::OnHScroll(UINT code, [[maybe_unused]] UINT pos)
{
    SCROLLINFO si{};
    si.cbSize = sizeof(SCROLLINFO);
    si.fMask  = SIF_ALL;
    GetScrollInfo(_hWnd, SB_HORZ, &si);

    float oldScrollX = _scrollX;
    float charWidth  = GetAverageCharWidth();

    switch (code)
    {
        case SB_LINELEFT: _scrollX -= charWidth; break;
        case SB_LINERIGHT: _scrollX += charWidth; break;
        case SB_PAGELEFT: _scrollX -= _clientDipW * 0.9f; break;
        case SB_PAGERIGHT: _scrollX += _clientDipW * 0.9f; break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: _scrollX = static_cast<float>(si.nTrackPos); break;
        case SB_LEFT: _scrollX = 0.f; break;
        case SB_RIGHT: _scrollX = (float)si.nMax; break;
    }

    ClampHorizontalScroll();

    if (oldScrollX != _scrollX)
    {
        UpdateScrollBars();
        Invalidate();
    }
}

float ColorTextView::GetLineHeight() const
{
    if (! _lineMetrics.empty())
        return _lineMetrics[0].height;
    return _fontSize + _linePaddingTop + _linePaddingBottom;
}
float ColorTextView::GetAverageCharWidth() const
{
    if (! _textFormat)
        return 8.0f; // fallback
    if (_avgCharWidthValid)
        return _avgCharWidth;

    // Measure a representative sample string
    static constexpr wchar_t kSample[] = L"ABCDEFGHabcdefgh0123456789";
    wil::com_ptr<IDWriteTextLayout> tl;
    if (SUCCEEDED(_dwriteFactory->CreateTextLayout(kSample, static_cast<UINT32>(wcslen(kSample)), _textFormat.get(), 1000.f, 1000.f, tl.addressof())))
    {
        DWRITE_TEXT_METRICS tm{};
        if (SUCCEEDED(tl->GetMetrics(&tm)) && tm.layoutWidth > 0 && wcslen(kSample) > 0)
        {
            _avgCharWidth      = tm.width / static_cast<float>(wcslen(kSample));
            _avgCharWidthValid = true;
            return _avgCharWidth;
        }
    }
    // Fallback if measurement failed
    _avgCharWidth      = _fontSize * 0.6f;
    _avgCharWidthValid = true;
    return _avgCharWidth;
}

void ColorTextView::ClampHorizontalScroll()
{
    float maxScrollX     = 0.f;
    float availableWidth = _clientDipW - (_padding * 2.f + (_displayLineNumbers ? _gutterDipW : 0.f));
    // Use approximate width derived from longest line and average char width
    maxScrollX = std::max(0.f, _approxContentWidth - availableWidth);
    _scrollX   = std::clamp(_scrollX, 0.f, maxScrollX);
}

UINT32 ColorTextView::GetCaretLine() const
{
    if (_document.TotalLineCount() == 0)
        return 0;
    auto [lineIndex, _] = _document.GetLineAndOffset(_caretPos);
    return (UINT32)std::min<size_t>(lineIndex, _document.TotalLineCount() - 1);
}
void ColorTextView::EnsureCaretVisible()
{
    if (_document.TotalLineCount() == 0)
        return;

    // Compute caret Y from display row mapping (handles filtering + embedded newlines).
    auto [lineIndex, off]  = _document.GetLineAndOffset(_caretPos);
    const auto& line       = _document.GetSourceLine(lineIndex);
    const UINT32 prefixLen = _document.PrefixLength(line);

    UINT32 rowInLine = 0;
    if (off > prefixLen && ! line.text.empty())
    {
        const size_t textOff = std::min<size_t>(static_cast<size_t>(off - prefixLen), line.text.size());
        const auto itEnd     = std::next(line.text.begin(), static_cast<std::wstring::difference_type>(textOff));
        rowInLine            = static_cast<UINT32>(std::count(line.text.begin(), itEnd, L'\n'));
    }

    const UINT32 caretDisplayRow = _document.DisplayRowForSource(lineIndex) + rowInLine;
    const float lineH            = GetLineHeight();
    const float caretTop         = static_cast<float>(caretDisplayRow) * lineH;
    const float caretBottom      = caretTop + lineH;
    float viewTop                = _scrollY;
    float viewBottom             = _scrollY + _clientDipH - _padding * 2;

    if (caretTop < viewTop)
        ScrollTo(caretTop - _padding);
    else if (caretBottom > viewBottom)
        ScrollTo(caretBottom - _clientDipH + _padding * 2);

    // Horizontal scrolling
    float caretLeft  = 0.f;
    float caretRight = 2.f; // caret width
    if (_textLayout)
    {
        DWRITE_HIT_TEST_METRICS metrics{};
        FLOAT x;
        FLOAT y;
        std::optional<UINT32> localPos;
        if (! _sliceIsFiltered)
        {
            if (_caretPos >= _sliceStartPos && _caretPos <= _sliceEndPos)
                localPos = _caretPos - _sliceStartPos;
        }
        else if (! _sliceFilteredRuns.empty())
        {
            auto it = std::upper_bound(_sliceFilteredRuns.begin(),
                                       _sliceFilteredRuns.end(),
                                       _caretPos,
                                       [](UINT32 value, const FilteredTextRun& run) { return value < run.sourceStart; });
            if (it != _sliceFilteredRuns.begin())
                --it;

            const UINT32 runStart = it->sourceStart;
            const UINT32 runEnd   = it->sourceStart + it->length;
            if (_caretPos >= runStart && _caretPos <= runEnd)
                localPos = it->layoutStart + (_caretPos - runStart);
        }

        if (localPos)
        {
            _textLayout->HitTestTextPosition(*localPos, FALSE, &x, &y, &metrics);
            caretLeft  = x;
            caretRight = x + 2.f;
        }
    }
    float availableWidth = _clientDipW - (_padding * 2.f + (_displayLineNumbers ? _gutterDipW : 0.f));
    float viewLeft       = _scrollX;
    float viewRight      = _scrollX + availableWidth;

    if (caretLeft < viewLeft)
        _scrollX = std::max(0.f, caretLeft - _padding);
    else if (caretRight > viewRight)
        _scrollX = caretRight - availableWidth + _padding;

    ClampScroll();
    ClampHorizontalScroll();
    UpdateScrollBars();
}

// ===== Win32 plumbing =====
LRESULT CALLBACK ColorTextView::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ColorTextView* self = nullptr;
    if (msg == WM_NCCREATE)
    {
        auto cs = reinterpret_cast<CREATESTRUCT*>(lp);
        self    = static_cast<ColorTextView*>(cs->lpCreateParams);
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)self);
        if (self)
        {
            self->_hWnd = hwnd;
            self->_hWndAtomic.store(hwnd, std::memory_order_release);
        }
    }
    else
    {
        self = reinterpret_cast<ColorTextView*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    }
    return self ? self->WndProc(hwnd, msg, wp, lp) : DefWindowProc(hwnd, msg, wp, lp);
}

void ColorTextView::CreateDeviceIndependentResources()
{
    HRESULT hr = S_OK;

    if (! _d2d1Factory)
    {
        D2D1_FACTORY_OPTIONS opts{};
#ifdef _DEBUG
        opts.debugLevel = D2D1_DEBUG_LEVEL_INFORMATION;
#endif
        hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_MULTI_THREADED, __uuidof(ID2D1Factory1), &opts, _d2d1Factory.put_void());
        if (FAILED(hr))
        {
            auto errorMsg = std::format(L"Failed to create D2D1 factory: HRESULT = 0x{:08X}\n", hr);
            OutputDebugStringW(errorMsg.c_str());
            return;
        }
    }
    if (! _dwriteFactory)
    {
        hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), _dwriteFactory.put_unknown());
        if (FAILED(hr))
        {
            auto errorMsg = std::format(L"Failed to create DirectWrite factory: HRESULT = 0x{:08X}\n", hr);
            OutputDebugStringW(errorMsg.c_str());
            return;
        }
    }

    if (! _textFormat)
    {
        hr = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 16.f, L"en-us", _textFormat.addressof());
        if (FAILED(hr))
        {
            auto errorMsg = std::format(L"Failed to create TextFormat Segoe: HRESULT = 0x{:08X}\n", hr);
            OutputDebugStringW(errorMsg.c_str());
            return;
        }
        _textFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
        _textFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
    }
    if (! _gutterTextFormat)
    {
        hr = _dwriteFactory->CreateTextFormat(L"Consolas",
                                              nullptr,
                                              DWRITE_FONT_WEIGHT_NORMAL,
                                              DWRITE_FONT_STYLE_NORMAL,
                                              DWRITE_FONT_STRETCH_NORMAL,
                                              12.f,
                                              L"en-us",
                                              _gutterTextFormat.addressof());
        if (FAILED(hr))
        {
            auto errorMsg = std::format(L"Failed to create TextFormat Consolas: HRESULT = 0x{:08X}\n", hr);
            OutputDebugStringW(errorMsg.c_str());
            return;
        }
    }
    _avgCharWidthValid = false;
}

void ColorTextView::CreateDeviceResources()
{
    // Ensure D2D/DWrite factories exist
    CreateDeviceIndependentResources();

    // Do nothing only if both the D2D context and the swap chain already exist
    if (_d2dCtx && _swapChain)
        return;

    HRESULT hr = S_OK;

    // Create D3D11 device
    UINT flags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
#if defined(_DEBUG)
    flags |= D3D11_CREATE_DEVICE_DEBUG;
#endif
    static const D3D_FEATURE_LEVEL levels[] = {D3D_FEATURE_LEVEL_11_1, D3D_FEATURE_LEVEL_11_0, D3D_FEATURE_LEVEL_10_1, D3D_FEATURE_LEVEL_10_0};
    D3D_FEATURE_LEVEL chosen                = D3D_FEATURE_LEVEL_11_0;

    // Try hardware first
    hr = D3D11CreateDevice(nullptr,
                           D3D_DRIVER_TYPE_HARDWARE,
                           nullptr,
                           flags,
                           levels,
                           ARRAYSIZE(levels),
                           D3D11_SDK_VERSION,
                           _d3dDevice.addressof(),
                           &chosen,
                           _d3dContext.addressof());
    if (FAILED(hr))
    {
#ifdef _DEBUG
        auto errorMsg = std::format(L"Failed to create D3D11 hardware device: HRESULT = 0x{:08X}, trying WARP...\n", hr);
        OutputDebugStringW(errorMsg.c_str());
#endif
        // Fallback to WARP
        hr = D3D11CreateDevice(nullptr,
                               D3D_DRIVER_TYPE_WARP,
                               nullptr,
                               flags,
                               levels,
                               ARRAYSIZE(levels),
                               D3D11_SDK_VERSION,
                               _d3dDevice.addressof(),
                               &chosen,
                               _d3dContext.addressof());
        if (FAILED(hr))
        {
#ifdef _DEBUG
            auto errorDeviceMsg = std::format(L"Failed to create D3D11 device (WARP): HRESULT = 0x{:08X}\n", hr);
            OutputDebugStringW(errorDeviceMsg.c_str());
#endif
            return;
        }
    }

    // Get current DPI and client size
    _dpi = static_cast<float>(GetDpiForWindow(_hWnd));
    RECT rc{};
    GetClientRect(_hWnd, &rc);
    const UINT width  = std::max(1u, static_cast<UINT>(rc.right - rc.left));
    const UINT height = std::max(1u, static_cast<UINT>(rc.bottom - rc.top));

#ifdef _DEBUG
    auto sizeMsg = std::format("Creating swap chain for {}x{} at {} DPI\n", width, height, _dpi);
    OutputDebugStringA(sizeMsg.c_str());
#endif

    // Create DXGI objects
    wil::com_ptr<IDXGIDevice> dxgiDevice;
    auto res = _d3dDevice.try_query_to(dxgiDevice.put());
    if (! res)
    {
#ifdef _DEBUG
        auto errorMsg = std::format(L"Failed to query DXGI device\n");
        OutputDebugStringW(errorMsg.c_str());
#endif
        return;
    }
    wil::com_ptr<IDXGIAdapter> dxgiAdapter;
    hr = dxgiDevice->GetAdapter(dxgiAdapter.addressof());
    if (FAILED(hr))
    {
#ifdef _DEBUG
        auto errorMsg = std::format(L"Failed to get DXGI adapter: HRESULT = 0x{:08X}\n", hr);
        OutputDebugStringW(errorMsg.c_str());
#endif
        return;
    }

    wil::com_ptr<IDXGIFactory2> dxgiFactory2;
    hr = dxgiAdapter->GetParent(IID_PPV_ARGS(dxgiFactory2.addressof()));
    if (FAILED(hr))
    {
        auto errorMsg = std::format(L"Failed to get DXGI factory: HRESULT = 0x{:08X}\n", hr);
        OutputDebugStringW(errorMsg.c_str());
        return;
    }

    // Create swap chain with validation
    DXGI_SWAP_CHAIN_DESC1 scd{};
    // Explicitly size buffers to current client size to avoid 0x0 edge cases
    scd.Width  = width;
    scd.Height = height;
    scd.Format = DXGI_FORMAT_B8G8R8A8_UNORM;
    // Allow render target output (for D2D target) and shader input (safe superset for interop)
    scd.BufferUsage      = DXGI_USAGE_RENDER_TARGET_OUTPUT | DXGI_USAGE_SHADER_INPUT;
    scd.SampleDesc.Count = 1;
    scd.BufferCount      = 2;
    scd.SwapEffect       = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
    // Do not allow composition to stretch the frame. We always resize the
    // swap chain buffers to exactly match the client area on WM_SIZE/WM_DPICHANGED
    // so text remains crisp and the image is never scaled.
    scd.Scaling   = DXGI_SCALING_NONE;
    scd.AlphaMode = DXGI_ALPHA_MODE_IGNORE;

    hr = dxgiFactory2->CreateSwapChainForHwnd(_d3dDevice.get(), _hWnd, &scd, nullptr, nullptr, _swapChain.addressof());
    if (FAILED(hr))
    {
        // Fallback: allow DXGI to choose scaling and buffer size
        scd.Width   = 0;
        scd.Height  = 0;
        scd.Scaling = DXGI_SCALING_STRETCH;
        hr          = dxgiFactory2->CreateSwapChainForHwnd(_d3dDevice.get(), _hWnd, &scd, nullptr, nullptr, _swapChain.addressof());
        if (FAILED(hr))
        {
            auto errorMsg = std::format(L"Failed to create swap chain (both modes): HRESULT = 0x{:08X}\n", hr);
            OutputDebugStringW(errorMsg.c_str());
            return;
        }
    }
    // Disable Alt+Enter
    wil::com_ptr<IDXGIFactory> dxgiFactory;
    dxgiFactory2.query_to(dxgiFactory.put());
    dxgiFactory->MakeWindowAssociation(_hWnd, DXGI_MWA_NO_ALT_ENTER);

    // Create D2D device/context
    wil::com_ptr<ID2D1Device> d2dDev;
    hr = _d2d1Factory->CreateDevice(dxgiDevice.get(), d2dDev.addressof());
    if (FAILED(hr))
    {
#ifdef _DEBUG
        auto errorMsg = std::format(L"Failed to create D2D device: HRESULT = 0x{:08X}\n", hr);
        OutputDebugStringW(errorMsg.c_str());
#endif
        return;
    }

    _d2dDevice = d2dDev;
    hr         = _d2dDevice->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_ENABLE_MULTITHREADED_OPTIMIZATIONS, _d2dCtx.addressof());
    if (FAILED(hr))
    {
#ifdef _DEBUG
        auto errorMsg = std::format(L"Failed to create D2D device context: HRESULT = 0x{:08X}\n", hr);
        OutputDebugStringW(errorMsg.c_str());
#endif
        return;
    }

    _d2dCtx->SetDpi(_dpi, _dpi);
    _d2dCtx->SetUnitMode(D2D1_UNIT_MODE_DIPS);

    // Create D2D target for backbuffer
    CreateSwapChainResources(width, height);

    // Invalidate any device-dependent caches created on a previous device
    _sliceBitmap.reset();
    _sliceCmd.reset();
    _brushCache.clear();
    _brushAccessCounter = 0;
    ClearTextLayoutEffects();
    EnsureFindBar();
    ApplyColoringToLayout();
}

void ColorTextView::DiscardDeviceResources()
{
    _brushCache.clear();
    _brushAccessCounter = 0;
    _sliceBitmap.reset();
    _sliceCmd.reset();

#ifdef _DEBUG
    _debugDirtyRectBrush.reset();
    _debugDirtyRectFillBrush.reset();
    _debugDirtyColorIndex = 0;
#endif

    if (_d2dCtx)
    {
        _d2dCtx->SetTarget(nullptr);
    }
    _d2dTargetBitmap.reset();

    // Force D3D context to release any deferred references
    if (_d3dContext)
    {
        _d3dContext->ClearState();
        _d3dContext->Flush();
    }

    // Release in proper order: D2D resources first, then D3D/DXGI
    _d2dCtx.reset();
    _d2dDevice.reset();
    _swapChain.reset();
    _d3dContext.reset();
    _d3dDevice.reset();
}

void ColorTextView::CreateSwapChainResources([[maybe_unused]] UINT w, [[maybe_unused]] UINT h)
{
    if (! _d2dCtx || ! _swapChain)
    {
        OutputDebugStringA("!!! CreateSwapChainResources: Missing required resources (_rt or _swapChain is null)\n");
        return;
    }

    // Reset previous target bitmap if it exists
    if (_d2dTargetBitmap)
    {
        _d2dCtx->SetTarget(nullptr);
        _d2dTargetBitmap.reset();
        OutputDebugStringA("!!! Cleared previous D2D target bitmap\n");
    }

    wil::com_ptr<IDXGISurface> surface;
    HRESULT hr = _swapChain->GetBuffer(0, IID_PPV_ARGS(surface.addressof()));
    if (FAILED(hr))
    {
        auto errorMsg = std::format(L"Failed to get swap chain buffer: HRESULT = 0x{:08X}\n", hr);
        OutputDebugStringW(errorMsg.c_str());
        return;
    }

#ifdef _DEBUG
    // Get and log surface description for diagnostics
    DXGI_SURFACE_DESC surfaceDesc{};
    hr = surface->GetDesc(&surfaceDesc);
    if (SUCCEEDED(hr))
    {
        RECT rc{};
        GetClientRect(_hWnd, &rc);
        auto surfaceMsg = std::format("Surface info: {}x{}, format: {}, sample count: {}\nClient Rect: {}x{}\n",
                                      surfaceDesc.Width,
                                      surfaceDesc.Height,
                                      static_cast<UINT>(surfaceDesc.Format),
                                      surfaceDesc.SampleDesc.Count,
                                      rc.right - rc.left,
                                      rc.bottom - rc.top);
        OutputDebugStringA(surfaceMsg.c_str());
    }
    // Log underlying D3D11 bind flags to verify TARGET eligibility
    if (auto tex2D = surface.try_query<ID3D11Texture2D>())
    {
        D3D11_TEXTURE2D_DESC td{};
        tex2D->GetDesc(&td);
        auto bindMsg = std::format("Backbuffer D3D11 BindFlags: 0x{:08X}\n", td.BindFlags);
        OutputDebugStringA(bindMsg.c_str());
    }
    // Validate DPI values
    if (_dpi <= 0.0f || _dpi > 1000.0f)
    {
        auto dpiMsg = std::format("Warning: Invalid DPI value: {}, using default 96.0\n", _dpi);
        OutputDebugStringA(dpiMsg.c_str());
        _dpi = 96.0f;
    }
#endif

    D2D1_BITMAP_PROPERTIES1 props{};
    props.pixelFormat = D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED);
    // Match the render target's DPI to ensure the bitmap DIP size matches the client DIP size
    props.dpiX = _dpi;
    props.dpiY = _dpi;
    // Create a drawable render target from the backbuffer; D2D requires TARGET and typically CANNOT_DRAW for swap chain surfaces
    props.bitmapOptions = D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW;
    hr                  = _d2dCtx->CreateBitmapFromDxgiSurface(surface.get(), &props, _d2dTargetBitmap.addressof());
    if (FAILED(hr))
    {
        auto errorMsg = std::format(L"Failed to create D2D bitmap from DXGI surface: HRESULT = 0x{:08X}\n", hr);
        OutputDebugStringW(errorMsg.c_str());

        // Additional diagnostics for specific error codes
        switch (hr)
        {
            case D2DERR_UNSUPPORTED_PIXEL_FORMAT: OutputDebugStringA("!!! Error: Unsupported pixel format\n"); break;
            case E_INVALIDARG: OutputDebugStringA("!!! Error: Invalid arguments (check DPI values and bitmap options)\n"); break;
            case D2DERR_INVALID_CALL: OutputDebugStringA("!!! Error: Invalid call state\n"); break;
            case E_OUTOFMEMORY: OutputDebugStringA("!!! Error: Out of memory\n"); break;
            default:
                auto specificMsg = std::format("!!! Error: Unknown error code 0x{:08X}\n", hr);
                OutputDebugStringA(specificMsg.c_str());
                break;
        }
        return;
    }
    _d2dCtx->SetTarget(_d2dTargetBitmap.get());
}

bool ColorTextView::RecreateSwapChain(UINT w, UINT h)
{
    if (! _d3dDevice)
        return false;

    // CRITICAL: Properly clean up existing swap chain to avoid "Only one flip model swap chain" error
    // Release all references to existing swap chain resources first
    if (_d2dCtx)
        _d2dCtx->SetTarget(nullptr);
    _d2dTargetBitmap.reset();
    _sliceBitmap.reset();

    // Force D3D context to release any deferred references
    if (_d3dContext)
    {
        _d3dContext->ClearState();
        _d3dContext->Flush();
    }

    // Release the swap chain completely before creating a new one
    _swapChain.reset();

    wil::com_ptr<IDXGIDevice> dxgiDevice;
    if (! _d3dDevice.try_query_to(dxgiDevice.put()))
        return false;
    wil::com_ptr<IDXGIAdapter> dxgiAdapter;
    if (FAILED(dxgiDevice->GetAdapter(dxgiAdapter.addressof())))
        return false;
    wil::com_ptr<IDXGIFactory2> dxgiFactory2;
    if (FAILED(dxgiAdapter->GetParent(IID_PPV_ARGS(dxgiFactory2.addressof()))))
        return false;

    DXGI_SWAP_CHAIN_DESC1 scd{};
    scd.Width            = w;
    scd.Height           = h;
    scd.Format           = DXGI_FORMAT_B8G8R8A8_UNORM;
    scd.BufferUsage      = DXGI_USAGE_RENDER_TARGET_OUTPUT | DXGI_USAGE_SHADER_INPUT;
    scd.SampleDesc.Count = 1;
    scd.BufferCount      = 2;
    scd.SwapEffect       = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
    scd.Scaling          = DXGI_SCALING_NONE;
    scd.AlphaMode        = DXGI_ALPHA_MODE_IGNORE;

    HRESULT hr = dxgiFactory2->CreateSwapChainForHwnd(_d3dDevice.get(), _hWnd, &scd, nullptr, nullptr, _swapChain.addressof());
    if (FAILED(hr))
    {
#ifdef _DEBUG
        auto errorMsg = std::format(L"RecreateSwapChain: CreateSwapChainForHwnd failed: 0x{:08X}\n", hr);
        OutputDebugStringW(errorMsg.c_str());
#endif
        return false;
    }

    // Disable Alt+Enter
    if (auto f = dxgiFactory2.try_query<IDXGIFactory>())
        f->MakeWindowAssociation(_hWnd, DXGI_MWA_NO_ALT_ENTER);

    CreateSwapChainResources(std::max(1u, w), std::max(1u, h));
    return _d2dTargetBitmap != nullptr;
}

void ColorTextView::EnsureBackbufferMatchesClient()
{
    if (! _swapChain || ! _d2dCtx)
    {
        OutputDebugStringA("!!! EnsureBackbufferMatchesClient: Missing required resources (_swapChain or _rt is null)\n");
        return;
    }

    RECT rc{};
    GetClientRect(_hWnd, &rc);
    const UINT clientW = std::max(1u, static_cast<UINT>(rc.right - rc.left));
    const UINT clientH = std::max(1u, static_cast<UINT>(rc.bottom - rc.top));

    // Current backbuffer size
    wil::com_ptr<ID3D11Texture2D> backbuffer;
    if (FAILED(_swapChain->GetBuffer(0, IID_PPV_ARGS(backbuffer.addressof()))))
        return;
    D3D11_TEXTURE2D_DESC td{};
    backbuffer->GetDesc(&td);

    // If the buffer already matches the client, nothing to do
    if (td.Width == clientW && td.Height == clientH)
    {
        _clientDipW = static_cast<float>(clientW) * 96.f / _dpi;
        _clientDipH = static_cast<float>(clientH) * 96.f / _dpi;
        return;
    }

#ifdef _DEBUG
    auto resizeMsg = std::format("ResizeBuffers {}x{} -> {}x{}\n", td.Width, td.Height, clientW, clientH);
    OutputDebugStringA(resizeMsg.c_str());
#endif

    // CRITICAL: Release ALL outstanding references to swap chain resources before ResizeBuffers
    // This is the fix for "cannot be resized unless all outstanding buffer references have been released"
    _d2dCtx->SetTarget(nullptr); // Remove D2D target first
    _d2dTargetBitmap.reset();    // Release the D2D bitmap wrapping the backbuffer
    _sliceBitmap.reset();        // Release slice bitmap to free memory
    backbuffer.reset();          // Release our reference to the backbuffer texture

    ResetPresentationState();
    RequestFullRedraw();

    // Force D3D context to release any deferred references
    if (_d3dContext)
    {
        _d3dContext->ClearState();
        _d3dContext->Flush();
    }

    HRESULT hr = _swapChain->ResizeBuffers(0, clientW, clientH, DXGI_FORMAT_UNKNOWN, 0);
    if (FAILED(hr))
    {
#ifdef _DEBUG
        auto errorMsg = std::format(L"Failed to resize swap chain buffers: HRESULT = 0x{:08X}\n", hr);
        OutputDebugStringW(errorMsg.c_str());
#endif

        // If resize failed, we need to recreate the entire swap chain
        if (! RecreateSwapChain(clientW, clientH))
            return;
    }
    else
    {
        CreateSwapChainResources(clientW, clientH);
    }

    _clientDipW = static_cast<float>(clientW) * 96.f / _dpi;
    _clientDipH = static_cast<float>(clientH) * 96.f / _dpi;
}

void ColorTextView::InvalidateSliceBitmap()
{
    _sliceBitmap.reset();
    _fallbackLayout.reset();
    _fallbackValid       = false;
    _fallbackStartLine   = 0;
    _fallbackEndLine     = 0;
    _fallbackLayoutWidth = 0.f;
    _fallbackFilteredRuns.clear();

    // Also invalidate tail layout for mode transitions
    _tailLayoutValid = false;
}

void ColorTextView::RebuildSliceBitmap()
{
    // TRACER;

    if (! _d2dCtx)
    {
        OutputDebugStringA("!!! RebuildSliceBitmap: Missing D2D context.\n");
        return;
    }
    if (! _textLayout)
    {
        OutputDebugStringA("!!! RebuildSliceBitmap: No text layout available; deferring rebuild.\n");
        return;
    }

    // Cache layout pointer to prevent re-entrant painting from invalidating it mid-use
    wil::com_ptr<IDWriteTextLayout> layoutToUse = _textLayout;
    if (! layoutToUse)
    {
        OutputDebugStringA("!!! RebuildSliceBitmap: Layout became null after check.\n");
        return;
    }

    // OPTIMIZATION: Incremental bitmap updates with dirty-region tracking
    // ===================================================================
    // When only a portion of the slice has changed (tracked via _sliceDirtyRegionValid),
    // we can update just that region instead of rebuilding the entire bitmap. This reduces
    // GPU upload overhead and CPU rendering time, especially beneficial at high log rates.
    // At 10k logs/sec with 256-line slices, this reduces rebuilds from ~2ms full to ~0.3ms partial.

    DWRITE_TEXT_METRICS tm{};
    if (FAILED(layoutToUse->GetMetrics(&tm)))
        return;
    const float wDip          = std::max(1.f, tm.widthIncludingTrailingWhitespace);
    const float hDip          = std::max(1.f, tm.height);
    const float viewportWidth = GetTextViewportWidthDip();
    const float cacheLimit    = std::min(kSliceBitmapMaxWidthDip, std::max(viewportWidth + kLayoutWidthSafetyMarginDip, kMinLayoutWidthDip));

    // Clamp width to cache limit instead of skipping - allows rendering of wide content (clipped)
    const bool isClipped     = (wDip > cacheLimit);
    const float clampedWidth = std::min(wDip, cacheLimit);

    if (isClipped)
    {
        auto msg = std::format("RebuildSliceBitmap: Clipping wide layout from {:.2f} dip to {:.2f} dip (will show overflow indicator)\n",
                               static_cast<double>(wDip),
                               static_cast<double>(clampedWidth));
        OutputDebugStringA(msg.c_str());
    }

    const UINT pxW = static_cast<UINT>(std::ceil(clampedWidth * (_dpi / 96.f)));
    const UINT pxH = static_cast<UINT>(std::ceil(hDip * (_dpi / 96.f)));

    // D3D11 texture dimension validation: Feature Level 11.x supports max 16384x16384
    // If slice is too large, skip bitmap caching and use direct layout rendering instead
    if (pxW > kMaxD3D11TextureDimension || pxH > kMaxD3D11TextureDimension)
    {
#ifdef _DEBUG
        auto msg = std::format("RebuildSliceBitmap: Slice too large ({}x{} px exceeds {}), using direct rendering\n", pxW, pxH, kMaxD3D11TextureDimension);
        OutputDebugStringA(msg.c_str());
#endif

        // Clear existing bitmap - will fall back to direct layout rendering in DrawScene
        _sliceBitmap.reset();
        _sliceDipW                   = clampedWidth;
        _sliceDipH                   = hDip;
        const UINT32 firstDisplayRow = _document.DisplayRowForSource(_sliceFirstLine);
        _sliceBitmapYBase            = static_cast<float>(firstDisplayRow) * GetLineHeight();
        return;
    }

    D2D1_BITMAP_PROPERTIES1 bp{};
    bp.pixelFormat = D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED);
    // Match the offscreen bitmap DPI to the current render DPI so
    // D2D/DWrite DIP calculations and bitmap sampling line up.
    // This avoids scale/size mismatches on per-monitor DPI (e.g., 150%).
    bp.dpiX          = _dpi;
    bp.dpiY          = _dpi;
    bp.bitmapOptions = D2D1_BITMAP_OPTIONS_TARGET;

    wil::com_ptr<ID2D1Bitmap1> newBmp;
    HRESULT hrCreate = _d2dCtx->CreateBitmap(D2D1::SizeU(pxW, pxH), nullptr, 0, &bp, newBmp.addressof());
    if (FAILED(hrCreate))
    {
        auto msg = std::format("!!! RebuildSliceBitmap: CreateBitmap failed hr=0x{:08X}, size={}x{} (fallback to direct rendering)\n",
                               static_cast<unsigned long>(hrCreate),
                               pxW,
                               pxH);
        OutputDebugStringA(msg.c_str());

        // Clear bitmap and rely on direct layout rendering
        _sliceBitmap.reset();
        _sliceDipW = clampedWidth;
        _sliceDipH = hDip;
        // Use display row offset for Y positioning (accounts for multi-line content)
        const UINT32 firstDisplayRow = _document.DisplayRowForSource(_sliceFirstLine);
        _sliceBitmapYBase            = static_cast<float>(firstDisplayRow) * GetLineHeight();
        return;
    }

    wil::com_ptr<ID2D1Image> prevTarget;
    _d2dCtx->GetTarget(prevTarget.addressof());
    D2D1_MATRIX_3X2_F prevXf{};
    _d2dCtx->GetTransform(&prevXf);
    _d2dCtx->SetTarget(newBmp.get());
    _d2dCtx->SetTransform(D2D1::Matrix3x2F::Identity());

    HRESULT hr = S_OK;
    {
        _d2dCtx->BeginDraw();
        auto endDraw = wil::scope_exit([&] { hr = _d2dCtx->EndDraw(); });
        _d2dCtx->Clear(D2D1::ColorF(0.f, 0.f, 0.f, 0.f));

        // Apply clipping rectangle if content is wider than cache limit
        if (isClipped)
        {
            const D2D1_RECT_F clipRect = D2D1::RectF(0.f, 0.f, clampedWidth, hDip);
            _d2dCtx->PushAxisAlignedClip(clipRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
        }

        auto* brush = GetBrush(_theme.fg);
        if (brush)
        {
            _d2dCtx->DrawTextLayout(D2D1::Point2F(0.f, 0.f), layoutToUse.get(), brush, D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
        }

        // Draw overflow indicator if content was clipped
        if (isClipped && brush)
        {
            // Draw orange underline across bottom of line
            const float underlineY = hDip - 2.f;
            auto* orangeBrush      = GetBrush(D2D1::ColorF(1.0f, 0.5f, 0.0f, 0.9f)); // Bright orange
            if (orangeBrush)
            {
                _d2dCtx->DrawLine(D2D1::Point2F(0.f, underlineY), D2D1::Point2F(clampedWidth, underlineY), orangeBrush, 2.f);
            }

            // Draw arrow indicator with dark background for visibility
            const float indicatorX = clampedWidth - 40.f;                             // 40 DIPs from right edge
            auto* darkBgBrush      = GetBrush(D2D1::ColorF(0.2f, 0.2f, 0.2f, 0.85f)); // Dark background
            if (darkBgBrush)
            {
                _d2dCtx->FillRectangle(D2D1::RectF(indicatorX, 0.f, clampedWidth, hDip), darkBgBrush);
            }

            auto* brightArrowBrush = GetBrush(D2D1::ColorF(1.0f, 0.8f, 0.0f, 1.0f)); // Bright yellow-orange
            if (brightArrowBrush && _textFormat)
            {
                wil::com_ptr<IDWriteTextLayout> indicator;
                const wchar_t* overflowText = L" →";
                if (SUCCEEDED(_dwriteFactory->CreateTextLayout(overflowText, 2, _textFormat.get(), 50.f, hDip, indicator.addressof())))
                {
                    _d2dCtx->DrawTextLayout(D2D1::Point2F(indicatorX + 5.f, 0.f), indicator.get(), brightArrowBrush, D2D1_DRAW_TEXT_OPTIONS_NONE);
                }
            }
        }

        if (isClipped)
        {
            _d2dCtx->PopAxisAlignedClip();
        }
    }

    _d2dCtx->SetTransform(prevXf);
    _d2dCtx->SetTarget(prevTarget.get());
    if (FAILED(hr))
    {
        auto msg = std::format("!!! RebuildSliceBitmap: EndDraw failed hr=0x{:08X}\n", static_cast<unsigned long>(hr));
        OutputDebugStringA(msg.c_str());
        return;
    }

    _sliceBitmap = std::move(newBmp);
    // Use display row offset for Y positioning (accounts for multi-line content)
    const UINT32 firstDisplayRow = _document.DisplayRowForSource(_sliceFirstLine);
    _sliceBitmapYBase            = static_cast<float>(firstDisplayRow) * GetLineHeight();
    _sliceDipW                   = clampedWidth; // Store clamped width, not original
    _sliceDipH                   = hDip;
    auto msg                     = std::format("RebuildSliceBitmap: Cached slice {:.2f}x{:.2f} dip at line {} {}\n",
                           static_cast<double>(_sliceDipW),
                           static_cast<double>(_sliceDipH),
                           static_cast<int>(_sliceFirstLine),
                           isClipped ? "(clipped with overflow indicator)" : "");
    OutputDebugStringA(msg.c_str());
}

void ColorTextView::OnCreate(const CREATESTRUCT* pCreateStruct)
{
    // DPI & size
    _dpi        = static_cast<float>(GetDpiForWindow(_hWnd));
    _clientDipW = static_cast<float>(pCreateStruct->cx) * 96.f / _dpi;
    _clientDipH = static_cast<float>(pCreateStruct->cy) * 96.f / _dpi;
    CreateDeviceIndependentResources();
    EnsureLayoutAsync();

    // Initialize scrollbars
    UpdateScrollBars();

    // Initialize gutter width cache
    UpdateGutterWidth();

    LogSystemInfo();
}

void ColorTextView::OnSize([[maybe_unused]] UINT width, [[maybe_unused]] UINT height)
{
    if (! _hWnd)
        return;
    // Requery physical client size and current DPI
    _dpi = static_cast<float>(GetDpiForWindow(_hWnd));
    if (_d2dCtx)
        _d2dCtx->SetDpi(_dpi, _dpi);

    // Track DIP client size only; defer swap-chain buffer work to OnPaint to avoid thrash
    RECT rc{};
    GetClientRect(_hWnd, &rc);
    const UINT pxW = std::max(1u, static_cast<UINT>(rc.right - rc.left));
    const UINT pxH = std::max(1u, static_cast<UINT>(rc.bottom - rc.top));
    _clientDipW    = static_cast<float>(pxW) * 96.f / _dpi;
    _clientDipH    = static_cast<float>(pxH) * 96.f / _dpi;

    // Viewport size changed; cached layouts are not reusable.
    _layoutCache.clear();
    _fallbackLayout.reset();
    _fallbackValid = false;
    _fallbackFilteredRuns.clear();
    InvalidateSliceBitmap();

    ClampHorizontalScroll();
    UpdateScrollBars();
    ClampScroll();
    UpdateScrollBars();

    if (_renderMode == RenderMode::AUTO_SCROLL)
    {
        _tailLayoutValid = false;
        RebuildTailLayout();
        ScrollToBottom();
    }
    else
    {
        EnsureLayoutAsync();
    }

    LayoutFindBar();
    Invalidate();
}

void ColorTextView::LogSystemInfo() const
{
#ifdef _DEBUG
    // Log adapter information
    if (_d3dDevice)
    {
        wil::com_ptr<IDXGIDevice> dxgiDevice;
        if (_d3dDevice.try_query_to(dxgiDevice.put()))
        {
            wil::com_ptr<IDXGIAdapter> adapter;
            if (SUCCEEDED(dxgiDevice->GetAdapter(adapter.addressof())))
            {
                DXGI_ADAPTER_DESC desc{};
                if (SUCCEEDED(adapter->GetDesc(&desc)))
                {
                    auto adapterMsg = std::format(L"Graphics Adapter: {}\n", std::wstring_view(desc.Description));
                    OutputDebugStringW(adapterMsg.c_str());

                    auto memMsg = std::format("Dedicated Video Memory: {} MB\n", desc.DedicatedVideoMemory / (1024 * 1024));
                    OutputDebugStringA(memMsg.c_str());
                }
            }
        }
    }

    // Log current window and DPI info
    RECT clientRect{};
    GetClientRect(_hWnd, &clientRect);
    auto windowMsg = std::format("Window size: {}x{}, DPI: {}\n", clientRect.right - clientRect.left, clientRect.bottom - clientRect.top, _dpi);
    OutputDebugStringA(windowMsg.c_str());
#endif
}

bool ColorTextView::ValidateDeviceState() const
{
    bool isValid = true;

    if (! _d3dDevice)
    {
        OutputDebugStringA("!!! Device validation: D3D11 device is null\n");
        isValid = false;
    }
    else
    {
        // Check if device is removed
        HRESULT reason = _d3dDevice->GetDeviceRemovedReason();
        if (FAILED(reason))
        {
            auto msg = std::format("!!! Device validation: D3D11 device removed, reason: 0x{:08X}\n", reason);
            OutputDebugStringA(msg.c_str());
            isValid = false;
        }
    }

    if (! _swapChain)
    {
        OutputDebugStringA("!!! Device validation: Swap chain is null\n");
        isValid = false;
    }

    if (! _d2dCtx)
    {
        OutputDebugStringA("!!! Device validation: D2D render target is null\n");
        isValid = false;
    }

    if (! _d2dTargetBitmap)
    {
        OutputDebugStringA("!!! Device validation: D2D target bitmap is null\n");
        isValid = false;
    }

    return isValid;
}

D2D1_RECT_F ColorTextView::GetViewRectDip() const
{
    float width  = _clientDipW;
    float height = _clientDipH;
    if (_horzScrollbarVisible)
        height = std::max(0.0f, height - GetHorzScrollbarDip(_hWnd, _dpi));
    return D2D1::RectF(0.0f, 0.0f, width, height);
}

RECT ColorTextView::GetViewRect() const
{
    const float scale         = _dpi / 96.0f;
    const D2D1_RECT_F dipRect = GetViewRectDip();
    RECT px{};
    px.left   = 0;
    px.top    = 0;
    px.right  = static_cast<LONG>(std::lround((dipRect.right - dipRect.left) * scale));
    px.bottom = static_cast<LONG>(std::lround((dipRect.bottom - dipRect.top) * scale));
    return px;
}

float ColorTextView::GetTextViewportWidthDip() const
{
    float width = _clientDipW - _padding * 2.f;
    if (_displayLineNumbers)
        width -= _gutterDipW;
    return std::max(width, 0.0f);
}

float ColorTextView::ComputeLayoutWidthDip() const
{
    const float viewport = GetTextViewportWidthDip();
    // Allow enough width to cover current scroll offset and the widest known content.
    const float scrolledExtent  = viewport + std::max(_scrollX, 0.0f);
    const float contentEstimate = std::max(_approxContentWidth, viewport);
    float target                = std::max({viewport, scrolledExtent, contentEstimate});
    target += kLayoutWidthSafetyMarginDip;
    target = std::min(target, kMaxLayoutWidthDip);
    return std::clamp(target, kMinLayoutWidthDip, kMaxLayoutWidthDip);
}

void ColorTextView::RequestFullRedraw()
{
    _needsFullRedraw  = true;
    _hasPendingDirty  = false;
    _hasPendingScroll = false;
    _pendingScrollDy  = 0;
    _pendingDirtyRect = RECT{};
}

void ColorTextView::ResetPresentationState()
{
    _presentInitialized = false;
    _hasPendingDirty    = false;
    _hasPendingScroll   = false;
    _pendingScrollDy    = 0;
    _pendingDirtyRect   = RECT{};
}

void ColorTextView::CreateFallbackLayoutIfNeeded(size_t visStartLine, size_t visEndLine)
{
    // TRACER;

    if (! _dwriteFactory || ! _textFormat || _document.TotalLineCount() == 0)
        return;

    // Add margin beyond visible range to prevent gaps during scrolling/updates
    // ARCHITECTURE NOTE: The fallback layout provides temporary rendering coverage
    // when the async slice worker hasn't caught up with scroll position changes.
    // The 32-line margin is in *visible-line space* to stay robust under heavy filtering.
    const size_t fallbackMarginVis = 32;
    const auto& visibleLines       = _document.VisibleLines();
    if (visibleLines.empty())
        return;

    const auto visBeginIt =
        std::lower_bound(visibleLines.begin(), visibleLines.end(), visStartLine, [](const VisibleLine& vl, size_t src) { return vl.sourceIndex < src; });
    const auto visEndIt =
        std::lower_bound(visibleLines.begin(), visibleLines.end(), visEndLine, [](const VisibleLine& vl, size_t src) { return vl.sourceIndex < src; });
    if (visBeginIt == visibleLines.end())
        return;

    const size_t startVisIdx = static_cast<size_t>(std::distance(visibleLines.begin(), visBeginIt));
    size_t endVisIdx         = static_cast<size_t>(std::distance(visibleLines.begin(), visEndIt));
    if (endVisIdx >= visibleLines.size())
        endVisIdx = visibleLines.size() - 1;

    const size_t startWithMarginVis = (startVisIdx > fallbackMarginVis) ? (startVisIdx - fallbackMarginVis) : 0;
    const size_t endWithMarginVis   = std::min(visibleLines.size() - 1, endVisIdx + fallbackMarginVis);
    const size_t clampedStart       = visibleLines[startWithMarginVis].sourceIndex;
    const size_t clampedEnd         = visibleLines[endWithMarginVis].sourceIndex;

    const float desiredWidth = ComputeLayoutWidthDip();

    // Check if existing fallback layout is still valid for current range
    const bool needsFallback = ! _fallbackValid || ! _fallbackLayout || _fallbackStartLine != clampedStart || _fallbackEndLine != clampedEnd ||
                               std::fabs(_fallbackLayoutWidth - desiredWidth) > 0.5f;

    if (! needsFallback)
        return;

    // Create new fallback layout synchronously (this is intentional - we need it immediately)
    _fallbackLayout.reset();
    _fallbackValid = false;
    _fallbackFilteredRuns.clear();

    // Build text - handle filtering
    std::wstring text;
    std::vector<FilteredTextRun> filteredRuns;
    if (_document.GetFilterMask() != Debug::InfoParam::Type::All)
    {
        // Build text from visible lines only
        const auto visBegin =
            std::lower_bound(visibleLines.begin(), visibleLines.end(), clampedStart, [](const VisibleLine& vl, size_t src) { return vl.sourceIndex < src; });
        const auto visEnd =
            std::upper_bound(visibleLines.begin(), visibleLines.end(), clampedEnd, [](size_t src, const VisibleLine& vl) { return src < vl.sourceIndex; });

        filteredRuns.reserve(static_cast<size_t>(std::distance(visBegin, visEnd)));
        for (auto it = visBegin; it != visEnd; ++it)
        {
            const size_t allIdx      = it->sourceIndex;
            const auto& displayText  = _document.GetDisplayTextRefAll(allIdx);
            const UINT32 layoutStart = static_cast<UINT32>(text.size());
            const UINT32 runLen      = static_cast<UINT32>(displayText.size()) + 1u; // +1 for '\n' (trimmed for last run below)
            const UINT32 sourceStart = _document.GetLineStartOffset(allIdx);
            filteredRuns.push_back(FilteredTextRun{
                .sourceLine  = allIdx,
                .layoutStart = layoutStart,
                .length      = runLen,
                .sourceStart = sourceStart,
            });
            text.append(displayText);
            text.append(L"\n");
        }

        // Remove trailing newline (layout uses separators between visible lines only)
        if (! text.empty())
        {
            text.pop_back();
            if (! filteredRuns.empty() && filteredRuns.back().length > 0)
                --filteredRuns.back().length;
        }
    }
    else
    {
        // No filtering - use position-based range
        const UINT32 startPos = _document.GetLineStartOffset(clampedStart);
        const auto& last      = _document.GetSourceLine(clampedEnd);
        const UINT32 endPos   = _document.GetLineStartOffset(clampedEnd) + _document.PrefixLength(last) + static_cast<UINT32>(last.text.size());
        const UINT32 length   = (endPos > startPos) ? (endPos - startPos) : 0;

        if (length > 0)
            text = _document.GetTextRange(startPos, length);
    }

    if (! text.empty())
    {
        wil::com_ptr<IDWriteTextLayout> temp;
        if (SUCCEEDED(
                _dwriteFactory->CreateTextLayout(text.c_str(), static_cast<UINT32>(text.size()), _textFormat.get(), desiredWidth, 1000000.f, temp.addressof())))
        {
            _fallbackLayout       = temp;
            _fallbackStartLine    = clampedStart;
            _fallbackEndLine      = clampedEnd;
            _fallbackLayoutWidth  = desiredWidth;
            _fallbackValid        = true;
            _fallbackFilteredRuns = std::move(filteredRuns);
        }
    }
}

void ColorTextView::DrawScene(bool clearTarget)
{
    // TRACER_CTX(L"DrawScene 1");

    if (clearTarget)
    {
        _d2dCtx->Clear(_theme.bg);
    }

    if (_document.TotalLineCount() == 0)
    {
        return; // Early exit
    }

    const float gutterDip = _displayLineNumbers ? _gutterDipW : 0.f;
    if (_displayLineNumbers)
    {
        // TRACER_CTX(L"displayLineNumbers");
        UpdateGutterWidth();
        const float viewDipH = _clientDipH - (_horzScrollbarVisible ? GetHorzScrollbarDip(_hWnd, _dpi) : 0.f);
        if (viewDipH > 0.f)
        {
            if (ID2D1SolidColorBrush* gutterBgBrush = GetBrush(_theme.gutterBg))
            {
                const D2D1_RECT_F gutterRect = D2D1::RectF(0.f, 0.f, _gutterDipW, viewDipH);
                _d2dCtx->FillRectangle(gutterRect, gutterBgBrush);
            }
        }
    }

    const float tx = _padding + gutterDip - _scrollX;
    const float ty = _padding - _scrollY;
    D2D1_MATRIX_3X2_F prev{};
    _d2dCtx->GetTransform(&prev);
    _d2dCtx->SetTransform(D2D1::Matrix3x2F::Translation(tx, ty));

    DrawHighlights();
    DrawSelection();
#ifdef _DEBUG
    DrawDebugSpans();
#endif

    // TWO-MODE RENDERING ARCHITECTURE
    // =================================
    // AUTO-SCROLL mode (hot path): Simple tail layout for last ~100 lines
    //   - No virtualization, no bitmap caching, no async workers
    //   - Synchronous layout updates on append for immediate visibility
    //   - Used when viewing the bottom of the log (99% use case)
    //
    // SCROLL-BACK mode (cold path): Full virtualization with bitmap caching
    //   - Slice-based rendering with async workers
    //   - Bitmap caching for performance during scrolling
    //   - Fallback layouts for gap coverage
    //   - Used when scrolling back through history (1% use case)

    if (_renderMode == RenderMode::AUTO_SCROLL && _tailLayoutValid && _tailLayout)
    {
        // TRACER_CTX(L"AutoScrollMode");

        // HOT PATH: Simple direct rendering of tail layout
        // No slice coverage checks, no fallbacks, no complexity
        wil::com_ptr<IDWriteTextLayout> layoutToUse = _tailLayout;
        const size_t firstLine                      = _tailFirstLine;

#ifdef _DEBUG
        static int autoScrollFrameCounter = 0;
        const bool shouldLog              = (++autoScrollFrameCounter % 60 == 0); // Log every 60th frame
        if (shouldLog)
        {
            auto msg =
                std::format("DrawScene: AUTO_SCROLL mode, tailFirstLine={}, scrollY={:.1f}, docLines={}\n", firstLine, _scrollY, _document.TotalLineCount());
            OutputDebugStringA(msg.c_str());
        }
#endif

        if (layoutToUse)
        {
            // Calculate layout metrics to position bottom-relative
            // This is critical when filtering is active - the tail layout contains only visible lines
            // but we need to position it at the document bottom, not at _tailFirstLine's absolute position
            DWRITE_TEXT_METRICS metrics{};
            layoutToUse->GetMetrics(&metrics);

            // Calculate how many display rows this layout occupies
            const float lineHeight         = GetLineHeight();
            const UINT32 layoutDisplayRows = static_cast<UINT32>(std::ceil(metrics.height / lineHeight));

            // Position layout so its bottom aligns with document bottom
            const UINT32 totalDisplayRows = _document.TotalDisplayRows();
            const float documentBottom    = static_cast<float>(totalDisplayRows) * lineHeight;
            const float yBase             = documentBottom - metrics.height;

#ifdef _DEBUG
            if (shouldLog)
            {
                const UINT32 firstDisplayRow = _document.DisplayRowForSource(firstLine);
                const float contentHeight    = _contentHeight;
                const float viewTop          = _scrollY - _padding;
                const float viewBottom       = viewTop + _clientDipH;
                auto msg = std::format("  AUTO_SCROLL: firstLine={}, firstDisplayRow={}, totalDisplayRows={}, layoutDisplayRows={}, layoutHeight={:.1f}, "
                                       "yBase={:.1f}, documentBottom={:.1f}, contentHeight={:.1f}, viewTop={:.1f}, viewBottom={:.1f}\n",
                                       firstLine,
                                       firstDisplayRow,
                                       totalDisplayRows,
                                       layoutDisplayRows,
                                       metrics.height,
                                       yBase,
                                       documentBottom,
                                       contentHeight,
                                       viewTop,
                                       viewBottom);
                OutputDebugStringA(msg.c_str());
            }
#endif

            auto* brush = GetBrush(_theme.fg);
            if (brush)
            {
                _d2dCtx->DrawTextLayout(D2D1::Point2F(0.f, yBase), layoutToUse.get(), brush, D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
            }
        }
    }
    else
    {
        // COLD PATH: Full virtualization mode (scroll-back through history)
        // TRACER_CTX(L"ScrollBackMode");

        auto [visStartLine, visEndLine] = GetVisibleLineRange();
        const bool sliceCoversView      = _textLayout && (_sliceFirstLine <= visStartLine) && (_sliceLastLine >= visEndLine);

#ifdef _DEBUG
        if (_document.TotalLineCount() > 0)
        {
            auto msg = std::format("DrawScene: mode=SCROLL_BACK, sliceCovers={}, visRange=[{},{}], sliceRange=[{},{}], docLines={}\n",
                                   sliceCoversView,
                                   visStartLine,
                                   visEndLine,
                                   _sliceFirstLine,
                                   _sliceLastLine,
                                   _document.TotalLineCount());
            OutputDebugStringA(msg.c_str());
        }
#endif

        if (! sliceCoversView)
        {
            // TRACER_CTX(L"SliceNotCoveringView");

            // Step 1: Create fallback layout if needed (separated for clarity and safety)
            CreateFallbackLayoutIfNeeded(visStartLine, visEndLine);

#ifdef _DEBUG
            {
                auto msg = std::format("DrawScene: Slice NOT covering view, fallbackValid={}, fallbackRange=[{},{}]\n",
                                       _fallbackValid ? 1 : 0,
                                       _fallbackStartLine,
                                       _fallbackEndLine);
                OutputDebugStringA(msg.c_str());
            }
#endif

            // Step 2: Render fallback layout (if available)
            if (_fallbackValid && _fallbackLayout)
            {
                // CRITICAL: Cache the layout pointer BEFORE any function calls to prevent
                // re-entrant painting (via GetLineHeight/GetBrush) from invalidating _fallbackLayout
                wil::com_ptr<IDWriteTextLayout> layoutToUse = _fallbackLayout;
                const size_t startLine                      = _fallbackStartLine;

                if (layoutToUse)
                {
                    // Use display row offset for Y positioning (accounts for multi-line content)
                    const UINT32 startDisplayRow = _document.DisplayRowForSource(startLine);
                    const float yBase            = static_cast<float>(startDisplayRow) * GetLineHeight();
                    auto* brush                  = GetBrush(_theme.fg);

#ifdef _DEBUG
                    {
                        auto msg = std::format("DrawScene: Rendering FALLBACK layout, startLine={}, startDisplayRow={}, yBase={:.1f}, scrollY={:.1f}\n",
                                               startLine,
                                               startDisplayRow,
                                               yBase,
                                               _scrollY);
                        OutputDebugStringA(msg.c_str());
                    }
#endif

                    if (brush)
                    {
                        _d2dCtx->DrawTextLayout(D2D1::Point2F(0.f, yBase), layoutToUse.get(), brush, D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
                    }
                }
            }

            // Step 3: Request async slice update for future frames
            EnsureLayoutAsync();
        }
        else
        {
            // TRACER_CTX(L"SliceCoveringView");

            _fallbackLayout.reset();
            _fallbackValid = false;
            // Draw either the cached slice OR the live layout, not both.
            // Drawing both led to doubled/overlapped glyphs when the slice already
            // covers the view (observed as "mixed up" text rendering).
            if (_sliceBitmap)
            {
                // TRACER_CTX(L"SliceBitmap");
                const float yBase     = _sliceBitmapYBase;
                const D2D1_RECT_F dst = D2D1::RectF(0.f, yBase, _sliceDipW, yBase + _sliceDipH);
                const D2D1_RECT_F src = D2D1::RectF(0.f, 0.f, _sliceDipW, _sliceDipH);

#ifdef _DEBUG
                {
                    auto msg = std::format(
                        "DrawScene: Rendering SLICE BITMAP, yBase={:.1f}, size={:.1f}x{:.1f}, scrollY={:.1f}\n", yBase, _sliceDipW, _sliceDipH, _scrollY);
                    OutputDebugStringA(msg.c_str());
                }
#endif

                _d2dCtx->DrawBitmap(_sliceBitmap.get(), dst, 1.f, D2D1_INTERPOLATION_MODE_NEAREST_NEIGHBOR, &src);
            }
            else if (_textLayout)
            {
                // TRACER_CTX(L"TextLayout");
                //  Cache layout pointer to prevent re-entrant painting issues
                wil::com_ptr<IDWriteTextLayout> layoutToUse = _textLayout;
                const size_t firstLine                      = _sliceFirstLine;

#ifdef _DEBUG
                {
                    DWRITE_TEXT_METRICS textMetrics{};
                    if (layoutToUse)
                    {
                        layoutToUse->GetMetrics(&textMetrics);
                    }
                    auto* brush            = GetBrush(_theme.fg);
                    const float lineHeight = GetLineHeight();
                    const float yBase      = static_cast<float>(_sliceFirstDisplayRow) * lineHeight;
                    auto msg =
                        std::format("DrawTextLayout: layout={:p} valid={}, firstLine={}, sliceFirstDisplayRow={}, yBase={:.1f}, lineHeight={:.1f}, brush={:p}, "
                                    "sliceRange=[{},{}], scrollY={:.1f}, ty={:.1f}, layoutSize={:.1f}x{:.1f}, lineCount={}\n",
                                    static_cast<const void*>(layoutToUse.get()),
                                    (layoutToUse ? 1 : 0),
                                    firstLine,
                                    _sliceFirstDisplayRow,
                                    yBase,
                                    lineHeight,
                                    static_cast<const void*>(brush),
                                    _sliceFirstLine,
                                    _sliceLastLine,
                                    _scrollY,
                                    static_cast<float>(_padding - _scrollY),
                                    textMetrics.width,
                                    textMetrics.height,
                                    textMetrics.lineCount);
                    OutputDebugStringA(msg.c_str());
                }
#endif

                if (layoutToUse)
                {
                    // Use pre-calculated display row offset for Y positioning
                    // This is critical when filtering: _sliceFirstDisplayRow accounts for all visible display rows
                    // up to _sliceFirstLine, while the layout contains only visible text from the slice
                    const float yBase = static_cast<float>(_sliceFirstDisplayRow) * GetLineHeight();
                    auto* brush       = GetBrush(_theme.fg);
                    if (brush)
                    {
                        _d2dCtx->DrawTextLayout(D2D1::Point2F(0.f, yBase), layoutToUse.get(), brush, D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
                    }
                }
            }
        }
    }

    _d2dCtx->SetTransform(prev);

    if (_displayLineNumbers)
    {
        DrawLineNumbers();
    }

    DrawCaret();
}

void ColorTextView::OnPaint()
{
    // TRACER;

    PAINTSTRUCT ps{};
    // scope for BeginPaint/EndPaint
    {
        wil::unique_hdc_paint paint_dc = wil::BeginPaint(_hWnd, &ps);
#ifdef _DEBUG
        static const LOGBRUSH logBrush{BS_SOLID, RGB(255, 0, 0), 0};
        static wil::unique_hbrush redBrush(CreateBrushIndirect(&logBrush));
        if (redBrush)
        {
            FillRect(paint_dc.get(), &ps.rcPaint, redBrush.get());
        }
#endif
    }

    CreateDeviceResources();
    if (! _d2dCtx || ! _swapChain || ! _d3dDevice)
    {
        OutputDebugStringA("!!! OnPaint: Missing device or swap chain after CreateDeviceResources, abort paint\n");
        return;
    }

    EnsureBackbufferMatchesClient();

    RECT clientRect{};
    GetClientRect(_hWnd, &clientRect);
    const UINT clientWidth  = static_cast<UINT>(clientRect.right - clientRect.left);
    const UINT clientHeight = static_cast<UINT>(clientRect.bottom - clientRect.top);
#ifdef _DEBUG
    if (_swapChain)
    {
        auto tex = wil::com_ptr<ID3D11Texture2D>();
        if (SUCCEEDED(_swapChain->GetBuffer(0, IID_PPV_ARGS(tex.addressof()))))
        {
            D3D11_TEXTURE2D_DESC td{};
            tex->GetDesc(&td);
            if (td.Width != clientWidth || td.Height != clientHeight)
            {
                auto msg = std::format("Paint: MISMATCH client {}x{} vs backbuffer {}x{}\n", clientWidth, clientHeight, td.Width, td.Height);
                OutputDebugStringA(msg.c_str());
            }
        }
    }
#endif

    if (! _d2dTargetBitmap)
    {
        OutputDebugStringA("!!! OnPaint: No D2D target bitmap after backbuffer resize\n");
        CreateSwapChainResources(clientWidth, clientHeight);

        if (! _d2dTargetBitmap)
        {
            OutputDebugStringA("!!! OnPaint: Failed to create D2D target bitmap, aborting paint\n");
            return;
        }
    }

    if (! _textLayout && _document.TotalLineCount() > 0)
        EnsureLayoutAsync();

    _d2dCtx->SetTarget(_d2dTargetBitmap.get());

    // Rebuild slice bitmap if missing and we have a layout
    if (! _sliceBitmap && _textLayout)
    {
        RebuildSliceBitmap();
    }

    const RECT viewRectPx   = GetViewRect();
    const LONG viewHeightPx = std::max<LONG>(0, viewRectPx.bottom - viewRectPx.top);
    const bool allowPartial = _presentInitialized && ! _needsFullRedraw;
    bool partialEligible    = allowPartial && _hasPendingDirty && _hasPendingScroll && (_pendingScrollDy != 0) && (viewHeightPx > 0);
    if (partialEligible)
    {
        // Without a ready layout (or any text), a partial invalidate can only paint
        // the debug dirty overlay, leaving the region blank. In that case fall
        // back to a full redraw so glyphs are refreshed in the same frame.
        const bool layoutReady = (_textLayout && (_sliceEndPos > _sliceStartPos));
        if (! layoutReady || _document.TotalLineCount() == 0)
        {
            partialEligible = false;
        }
    }

    RECT dirtyRectPx    = _pendingDirtyRect;
    LONG scrollAmountPx = 0;
    if (partialEligible)
    {
        scrollAmountPx = std::clamp(_pendingScrollDy, -viewHeightPx, viewHeightPx);
        if (scrollAmountPx == 0)
        {
            partialEligible = false;
        }
        else
        {
            dirtyRectPx          = viewRectPx;
            const LONG absScroll = scrollAmountPx >= 0 ? scrollAmountPx : -scrollAmountPx;
            if (scrollAmountPx > 0)
            {
                dirtyRectPx.top = std::max(dirtyRectPx.bottom - absScroll, dirtyRectPx.top);
            }
            else
            {
                dirtyRectPx.bottom = std::min(dirtyRectPx.top + absScroll, dirtyRectPx.bottom);
            }

            if (dirtyRectPx.bottom <= dirtyRectPx.top)
                partialEligible = false;
        }
    }

    bool drew        = false;
    bool usedPartial = false;

    HRESULT hr = S_OK;
    if (partialEligible)
    {
        {
            _d2dCtx->BeginDraw();
            auto endDraw               = wil::scope_exit([&] { hr = _d2dCtx->EndDraw(); });
            const float invScale       = (_dpi > 0.0f) ? 96.0f / _dpi : 1.0f;
            const D2D1_RECT_F dirtyDip = D2D1::RectF(static_cast<FLOAT>(static_cast<float>(dirtyRectPx.left) * invScale),
                                                     static_cast<FLOAT>(static_cast<float>(dirtyRectPx.top) * invScale),
                                                     static_cast<FLOAT>(static_cast<float>(dirtyRectPx.right) * invScale),
                                                     static_cast<FLOAT>(static_cast<float>(dirtyRectPx.bottom) * invScale));

            _d2dCtx->PushAxisAlignedClip(dirtyDip, D2D1_ANTIALIAS_MODE_ALIASED);
            if (ID2D1SolidColorBrush* bgBrush = GetBrush(_theme.bg))
            {
                _d2dCtx->FillRectangle(dirtyDip, bgBrush);
            }
#ifdef _DEBUG
            if (_d2dCtx)
            {
                if (! _debugDirtyRectFillBrush)
                {
                    const D2D1_COLOR_F initialColor = debugDirtyPaletteSize ? debugDirtyPalette[0] : D2D1::ColorF(1.f, 0.f, 0.f, 0.35f);
                    const HRESULT fillHr            = _d2dCtx->CreateSolidColorBrush(initialColor, _debugDirtyRectFillBrush.addressof());
                    if (FAILED(fillHr))
                    {
                        auto msg = std::format("!!! Failed to create debug dirty rect fill brush: HRESULT = 0x{:08X}\n", fillHr);
                        OutputDebugStringA(msg.c_str());
                        _debugDirtyRectFillBrush.reset();
                    }
                }
                if (_debugDirtyRectFillBrush && debugDirtyPaletteSize > 0)
                {
                    const size_t paletteIndex = _debugDirtyColorIndex % debugDirtyPaletteSize;
                    _debugDirtyRectFillBrush->SetColor(debugDirtyPalette[paletteIndex]);
                    _debugDirtyColorIndex = (_debugDirtyColorIndex + 1) % debugDirtyPaletteSize;
                    _d2dCtx->FillRectangle(dirtyDip, _debugDirtyRectFillBrush.get());
                }
            }
#endif
            DrawScene(false);

            _d2dCtx->PopAxisAlignedClip();

#ifdef _DEBUG
            if (_d2dCtx)
            {
                if (! _debugDirtyRectBrush)
                {
                    const auto outlineColor = D2D1::ColorF(0.f, 0.f, 0.f, 1.f);
                    const HRESULT brushHr   = _d2dCtx->CreateSolidColorBrush(outlineColor, _debugDirtyRectBrush.addressof());
                    if (FAILED(brushHr))
                    {
                        auto msg = std::format("!!! Failed to create debug dirty rect brush: HRESULT = 0x{:08X}\n", brushHr);
                        OutputDebugStringA(msg.c_str());
                        _debugDirtyRectBrush.reset();
                    }
                }
                if (_debugDirtyRectBrush)
                {
                    _d2dCtx->DrawRectangle(dirtyDip, _debugDirtyRectBrush.get(), 1.0f);
                }
            }
#endif
        }

        if (FAILED(hr))
        {
            auto errorMsg = std::format("!!! Partial EndDraw failed: HRESULT = 0x{:08X}\n", hr);
            OutputDebugStringA(errorMsg.c_str());
            if (hr == D2DERR_RECREATE_TARGET || hr == DXGI_ERROR_DEVICE_REMOVED || hr == DXGI_ERROR_DEVICE_RESET)
            {
                DiscardDeviceResources();
                ResetPresentationState();
            }
            RequestFullRedraw();
        }
        else
        {
            drew        = true;
            usedPartial = true;
        }
    }

    if (! drew)
    {
        if (! _d2dCtx)
        {
            RequestFullRedraw();
            InvalidateRect(_hWnd, nullptr, FALSE);
            return;
        }

        {
            _d2dCtx->BeginDraw();
            auto endDraw = wil::scope_exit([&] { hr = _d2dCtx->EndDraw(); });
            DrawScene(true);
        }

        if (FAILED(hr))
        {
            auto errorMsg = std::format("!!! EndDraw failed: HRESULT = 0x{:08X}\n", hr);
            OutputDebugStringA(errorMsg.c_str());

            if (hr == D2DERR_RECREATE_TARGET || hr == DXGI_ERROR_DEVICE_REMOVED || hr == DXGI_ERROR_DEVICE_RESET)
            {
                OutputDebugStringA("!!! Device lost during full render, discarding resources\n");
                DiscardDeviceResources();
                ResetPresentationState();
            }

            RequestFullRedraw();
            InvalidateRect(_hWnd, nullptr, FALSE);
            return;
        }

        drew        = true;
        usedPartial = false;
    }

    if (! _swapChain)
        return;

    DXGI_PRESENT_PARAMETERS params{};
    RECT scrollRectPx{};
    POINT scrollOffset{0, 0};
    if (usedPartial)
    {
        scrollRectPx         = viewRectPx;
        const LONG absScroll = scrollAmountPx >= 0 ? scrollAmountPx : -scrollAmountPx;
        if (scrollAmountPx > 0)
        {
            scrollRectPx.bottom = std::max(scrollRectPx.top, scrollRectPx.bottom - absScroll);
        }
        else
        {
            scrollRectPx.top = std::min(scrollRectPx.bottom, scrollRectPx.top + absScroll);
        }
        scrollOffset.y         = -scrollAmountPx;
        params.DirtyRectsCount = 1;
        params.pDirtyRects     = &dirtyRectPx;
        params.pScrollRect     = &scrollRectPx;
        params.pScrollOffset   = &scrollOffset;
    }

    const UINT syncInterval = _inSizeMove ? 0u : 1u;
    const HRESULT presentHr = _swapChain->Present1(syncInterval, 0, &params);
    if (FAILED(presentHr))
    {
        auto presentMsg = std::format("!!! Present1 failed: HRESULT = 0x{:08X}\n", presentHr);
        OutputDebugStringA(presentMsg.c_str());

        if (presentHr == DXGI_ERROR_DEVICE_REMOVED || presentHr == DXGI_ERROR_DEVICE_RESET)
        {
            OutputDebugStringA("!!! Present failed due to device removed/reset, discarding resources\n");
            DiscardDeviceResources();
        }

        ResetPresentationState();
        RequestFullRedraw();
        InvalidateRect(_hWnd, nullptr, FALSE);
        return;
    }

    _presentInitialized = true;
    _needsFullRedraw    = false;
    _hasPendingDirty    = false;
    _hasPendingScroll   = false;
    _pendingScrollDy    = 0;
    _pendingDirtyRect   = RECT{};
}

void ColorTextView::OnMouseWheel(short delta)
{
    bool shift = IsKeyDown(VK_SHIFT);

    if (shift)
    {
        // Horizontal scrolling with Shift+Mouse Wheel
        float charWidth = GetAverageCharWidth();
        _wheelRemainderX += delta;
        const int steps = _wheelRemainderX / WHEEL_DELTA;
        _wheelRemainderX -= steps * WHEEL_DELTA;
        if (steps == 0)
        {
            return;
        }

        float scrollAmount = -static_cast<float>(steps) * 3.0f * charWidth;

        float oldScrollX = _scrollX;
        _scrollX += scrollAmount;
        ClampHorizontalScroll();
        if (oldScrollX != _scrollX)
        {
            UpdateScrollBars();
            Invalidate();
        }
    }
    else
    {
        // Vertical scrolling
        UINT linesPerNotch = 3;
        SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &linesPerNotch, 0);

        float scrollAmount = 0.0f;
        _wheelRemainder += delta;
        const int steps = _wheelRemainder / WHEEL_DELTA;
        _wheelRemainder -= steps * WHEEL_DELTA;
        if (steps == 0)
        {
            return;
        }

        if (linesPerNotch == WHEEL_PAGESCROLL)
        {
            // Scroll by page
            scrollAmount = -static_cast<float>(steps) * _clientDipH;
        }
        else
        {
            // Scroll by lines
            // If linesPerNotch is 0 (unlikely but possible), we shouldn't scroll
            if (linesPerNotch > 0)
            {
                scrollAmount = -static_cast<float>(steps) * static_cast<float>(linesPerNotch) * GetLineHeight();
            }
        }

        if (scrollAmount != 0.0f)
        {
            // User is manually scrolling - if scrolling UP, disable auto-scroll
            if (scrollAmount < 0.0f && _renderMode == RenderMode::AUTO_SCROLL)
            {
#ifdef _DEBUG
                OutputDebugStringA("OnMouseWheel: User scrolled UP, disabling auto-scroll\n");
#endif
                SwitchToScrollBackMode();
            }

            ScrollBy(scrollAmount);

            // Two-mode transition: check if we should switch modes after wheel scroll
            if (ShouldUseAutoScrollMode())
            {
                if (_renderMode != RenderMode::AUTO_SCROLL)
                {
#ifdef _DEBUG
                    OutputDebugStringA("OnMouseWheel: Switching to AUTO_SCROLL mode\n");
#endif
                    SwitchToAutoScrollMode();
                }
            }
            else
            {
                if (_renderMode != RenderMode::SCROLL_BACK)
                {
#ifdef _DEBUG
                    {
                        auto msg = std::format("OnMouseWheel: Switching to SCROLL_BACK mode, scrollY={:.1f}, contentHeight={:.1f}\n", _scrollY, _contentHeight);
                        OutputDebugStringA(msg.c_str());
                    }
#endif
                    SwitchToScrollBackMode();
                }
            }
        }
    }
}

void ColorTextView::Invalidate()
{
    if (! _hWnd)
        return;
    RequestFullRedraw();
    InvalidateRect(_hWnd, nullptr, FALSE);
}
void ColorTextView::ScrollBy(float dy)
{
    float oldY = _scrollY;
    _scrollY += dy;
    ClampScroll();
    UpdateScrollBars();
    MaybeRefreshVirtualSliceOnScroll();
    InvalidateExposedArea(_scrollY - oldY);
}
void ColorTextView::ScrollTo(float y)
{
    float oldY = _scrollY;
    _scrollY   = y;
    ClampScroll();
    UpdateScrollBars();
    MaybeRefreshVirtualSliceOnScroll();
    InvalidateExposedArea(_scrollY - oldY);
}

void ColorTextView::ScrollToBottom()
{
    // Scroll to the bottom by requesting the maximum content height;
    // internal clamping will place the view at the bottom edge.
    ScrollTo(_contentHeight);
}

void ColorTextView::RequestScrollToBottom()
{
    _pendingScrollToBottom = true;
    ScrollToBottom();
}

void ColorTextView::SwitchToAutoScrollMode()
{
    if (_renderMode == RenderMode::AUTO_SCROLL)
        return; // Already in auto-scroll mode

    // TRACER;

    _renderMode      = RenderMode::AUTO_SCROLL;
    _tailLayoutValid = false;

    // Rebuild tail layout for immediate rendering
    RebuildTailLayout();

    // Scroll to bottom
    ScrollToBottom();

#ifdef _DEBUG
    OutputDebugStringA("ColorTextView: Switched to AUTO-SCROLL mode\n");
#endif
}

void ColorTextView::SwitchToScrollBackMode()
{
    if (_renderMode == RenderMode::SCROLL_BACK)
        return; // Already in scroll-back mode

    // TRACER;

    _renderMode = RenderMode::SCROLL_BACK;
    _tailLayout.reset();
    _tailLayoutValid = false;

    // Invalidate existing slice - it was created for auto-scroll and won't cover the new scroll position
    InvalidateSliceBitmap();

    // Trigger full virtualized layout
    EnsureLayoutAsync();

#ifdef _DEBUG
    OutputDebugStringA("ColorTextView: Switched to SCROLL-BACK mode\n");
#endif
}

bool ColorTextView::ShouldUseAutoScrollMode() const
{
    // _renderMode is the single source of truth for auto-scroll state
    // When in AUTO_SCROLL mode: tail optimization, scrolls to bottom on append
    // When in SCROLL_BACK mode: full virtualization, stays at current position

    // Mode is managed by:
    // - User toggling via SetAutoScroll() (called by main window menu)
    // - User scrolling UP -> switches to SCROLL_BACK mode
    // - User jumping to END -> switches to AUTO_SCROLL mode

    return _renderMode == RenderMode::AUTO_SCROLL;
}

void ColorTextView::RebuildTailLayout()
{
    // TRACER;

    if (! _dwriteFactory || ! _textFormat)
    {
        _tailLayoutValid = false;
        return;
    }

    const size_t lineCount = _document.TotalLineCount();
    if (lineCount == 0)
    {
        _tailLayout.reset();
        _tailLayoutValid = false;
        _tailFirstLine   = 0;
        return;
    }

    // Calculate tail window: must cover viewport + margin for smooth auto-scroll
    // Viewport can show many lines (e.g., 50-200 depending on window size/DPI)
    const float lineHeight    = GetLineHeight();
    const float viewHeight    = _clientDipH - (_horzScrollbarVisible ? GetHorzScrollbarDip(_hWnd, _dpi) : 0.f);
    const size_t visibleLines = static_cast<size_t>(std::ceil(viewHeight / std::max(lineHeight, 1.0f)));

    // Tail window = visible lines + margin (at least kTailLines, but scale with viewport)
    const size_t tailWindowSize = std::max(kTailLines, visibleLines + 50u);

    _tailFirstLine            = (lineCount > tailWindowSize) ? (lineCount - tailWindowSize) : 0;
    const size_t tailLastLine = lineCount - 1;

#ifdef _DEBUG
    static int tailRebuildCounter = 0;
    const bool shouldLog          = (++tailRebuildCounter % 10 == 0); // Log every 10th rebuild
    if (shouldLog)
    {
        auto msg = std::format("RebuildTailLayout #{}: lineCount={}, visibleLines={}, tailWindow={}, range=[{}, {}]\n",
                               tailRebuildCounter,
                               lineCount,
                               visibleLines,
                               tailWindowSize,
                               _tailFirstLine,
                               tailLastLine);
        OutputDebugStringA(msg.c_str());
    }
#endif

    // Build text for tail - handle filtering
    std::wstring tailText;
    Document::FilteredTailResult filteredTail; // Holds metadata for coloring pass
    if (_document.GetFilterMask() != Debug::InfoParam::Type::All)
    {
        // Build filtered text in a single locked scope (avoids per-line IsLineVisible + GetDisplayTextRefAll lock overhead)
        filteredTail       = _document.BuildFilteredTailText(_tailFirstLine, tailLastLine);
        tailText           = std::move(filteredTail.text);
        _tailFilteredLines = std::move(filteredTail.lines);

#ifdef _DEBUG
        if (shouldLog)
        {
            const UINT32 firstDisplayRow  = _document.DisplayRowForSource(_tailFirstLine);
            const UINT32 totalDisplayRows = _document.TotalDisplayRows();
            auto msg                      = std::format("  FILTERED: visibleInTail={}, firstDisplayRow={}, totalDisplayRows={}, textLength={}\n",
                                   filteredTail.visibleCount,
                                   firstDisplayRow,
                                   totalDisplayRows,
                                   tailText.length());
            OutputDebugStringA(msg.c_str());
        }
#endif
    }
    else
    {
        // No filtering - use position-based range
        _tailFilteredLines.clear();
        const UINT32 startPos = _document.GetLineStartOffset(_tailFirstLine);
        const auto& lastLine  = _document.GetSourceLine(tailLastLine);
        const UINT32 endPos   = _document.GetLineStartOffset(tailLastLine) + _document.PrefixLength(lastLine) + static_cast<UINT32>(lastLine.text.size());
        const UINT32 length   = (endPos > startPos) ? (endPos - startPos) : 0;

        if (length > 0)
            tailText = _document.GetTextRange(startPos, length);
    }

    if (tailText.empty())
    {
        _tailLayout.reset();
        _tailLayoutValid = false;
        return;
    }

    // Create layout synchronously (fast - only ~100 lines)
    const float layoutWidth = ComputeLayoutWidthDip();
    wil::com_ptr<IDWriteTextLayout> newLayout;
    HRESULT hr = _dwriteFactory->CreateTextLayout(tailText.c_str(),
                                                  static_cast<UINT32>(tailText.size()),
                                                  _textFormat.get(),
                                                  layoutWidth,
                                                  1000000.f, // Large height, layout will compute actual
                                                  newLayout.addressof());

    if (SUCCEEDED(hr) && newLayout)
    {
        _tailLayout      = newLayout;
        _tailLayoutValid = true;

        // Apply coloring to tail layout
        ApplyColoringToTailLayout();
    }
    else
    {
        _tailLayout.reset();
        _tailLayoutValid = false;

#ifdef _DEBUG
        auto msg2 = std::format("RebuildTailLayout: CreateTextLayout failed, hr=0x{:08X}\n", static_cast<unsigned long>(hr));
        OutputDebugStringA(msg2.c_str());
#endif
    }
}

void ColorTextView::ApplyColoringToTailLayout()
{
    if (! _tailLayout || ! _tailLayoutValid)
        return;

    // Apply color ranges to tail layout
    // Similar to ApplyColoringToLayout but for tail window
    const size_t lineCount = _document.TotalLineCount();
    if (lineCount == 0)
        return;

    const size_t tailLastLine = lineCount - 1;
    const bool isFiltered     = (_document.GetFilterMask() != Debug::InfoParam::Type::All);

    if (isFiltered)
    {
        // FILTERED MODE: Use cached _tailFilteredLines from RebuildTailLayout
        // This avoids per-line IsLineVisible() + GetSourceLine() lock overhead
        UINT32 layoutOffset = 0;

        for (const auto& info : _tailFilteredLines)
        {
            const UINT32 lineLen = info.prefixLen + info.textLen + 1; // +1 for newline separator

            if (info.hasMeta && info.prefixLen > 0)
            {
                D2D1_COLOR_F color = _theme.metaText;
                switch (info.type)
                {
                    case Debug::InfoParam::Error: color = _theme.metaError; break;
                    case Debug::InfoParam::Warning: color = _theme.metaWarning; break;
                    case Debug::InfoParam::Info: color = _theme.metaInfo; break;
                    case Debug::InfoParam::Debug: color = _theme.metaDebug; break;
                    case Debug::InfoParam::Text:
                    case Debug::InfoParam::All:
                    default: break;
                }

                DWRITE_TEXT_RANGE range{layoutOffset, info.prefixLen};
                auto* brush = GetBrush(color);
                if (brush)
                    _tailLayout->SetDrawingEffect(brush, range);
            }

            layoutOffset += lineLen;
        }
    }
    else
    {
        // NON-FILTERED MODE: Use document-absolute positions (original logic)
        const UINT32 tailStartPos = _document.GetLineStartOffset(_tailFirstLine);

        for (size_t i = _tailFirstLine; i <= tailLastLine; ++i)
        {
            const auto& line = _document.GetSourceLine(i);
            if (! line.hasMeta)
                continue;

            const UINT32 lineStartPos = _document.GetLineStartOffset(i);
            const UINT32 prefixLen    = _document.PrefixLength(line);
            const UINT32 localStart   = lineStartPos - tailStartPos;

            // Color the prefix based on severity
            D2D1_COLOR_F color = _theme.metaText;
            switch (line.meta.type)
            {
                case Debug::InfoParam::Error: color = _theme.metaError; break;
                case Debug::InfoParam::Warning: color = _theme.metaWarning; break;
                case Debug::InfoParam::Info: color = _theme.metaInfo; break;
                case Debug::InfoParam::Debug: color = _theme.metaDebug; break;
                case Debug::InfoParam::Text:
                case Debug::InfoParam::All: // Not a message type, use default
                default: break;
            }

            if (prefixLen > 0)
            {
                DWRITE_TEXT_RANGE range{localStart, prefixLen};
                auto* brush = GetBrush(color);
                if (brush)
                    _tailLayout->SetDrawingEffect(brush, range);
            }
        }
    }
}

void ColorTextView::EnsureLayoutAsync()
{
    CreateDeviceIndependentResources();
    if (! _dwriteFactory)
        return;
    // Debounce: coalesce rapid requests into a single layout
    if (! _layoutTimerArmed)
    {
        _layoutTimerArmed = true;
        SetTimer(_hWnd, kLayoutTimerId, kLayoutTimerDelayMs, nullptr);
    }
}

void ColorTextView::EnsureLayoutAdaptive(size_t changeSize)
{
    CreateDeviceIndependentResources();
    if (! _dwriteFactory)
        return;

    // Adaptive timing based on change magnitude:
    // - Small changes (<100 lines): synchronous layout for immediate feedback
    // - Medium changes (100-1000 lines): fast timer (4ms) for responsive updates
    // - Large changes (>1000 lines): standard timer (16ms) for batching efficiency

    if (changeSize < kSyncLayoutThresholdLines && changeSize > 0)
    {
        // Small change - process synchronously for immediate visual feedback
        // This avoids the 4-16ms timer lag during normal typing/single line additions
        if (_layoutTimerArmed)
        {
            KillTimer(_hWnd, kLayoutTimerId);
            _layoutTimerArmed = false;
        }

        const float layoutWidth = ComputeLayoutWidthDip();
        const UINT32 seq        = ++_layoutSeq;
        StartLayoutWorker(layoutWidth, seq);
    }
    else
    {
        // Medium/large change - use timer-based debouncing
        if (! _layoutTimerArmed)
        {
            _layoutTimerArmed = true;
            // Use fast timer for medium changes, standard timer for large changes
            const UINT delay = (changeSize < 1000) ? kFastLayoutTimerDelayMs : kLayoutTimerDelayMs;
            SetTimer(_hWnd, kLayoutTimerId, delay, nullptr);
        }
    }
}

void ColorTextView::MaybeRefreshVirtualSliceOnScroll()
{
    // Only applies to SCROLL-BACK mode with virtualized rendering
    if (_renderMode != RenderMode::SCROLL_BACK || ! _textLayout)
        return;

    if (_document.TotalLineCount() == 0 || _document.VisibleLines().empty())
        return;

    // Compute visible index range for current viewport (display-row based).
    const float H          = GetLineHeight();
    const float viewTop    = _scrollY - _padding;
    const float viewBottom = viewTop + (_clientDipH - (_horzScrollbarVisible ? GetHorzScrollbarDip(_hWnd, _dpi) : 0.f));
    const float rowH       = std::max(H, 1e-3f);
    const UINT32 topRow    = static_cast<UINT32>(std::floor(std::max(0.f, viewTop) / rowH));
    const UINT32 bottomRow = static_cast<UINT32>(std::floor(viewBottom / rowH));

    const UINT32 maxDisplayRow    = _document.TotalDisplayRows();
    const UINT32 clampedTopRow    = std::min(topRow, maxDisplayRow > 0 ? maxDisplayRow - 1 : 0);
    const UINT32 clampedBottomRow = std::min(bottomRow, maxDisplayRow > 0 ? maxDisplayRow - 1 : 0);

    const size_t visStartIdx = _document.VisibleIndexFromDisplayRow(clampedTopRow);
    const size_t visEndIdx   = _document.VisibleIndexFromDisplayRow(clampedBottomRow);

    const auto& visibleLines = _document.VisibleLines();
    const auto sliceBeginIt =
        std::lower_bound(visibleLines.begin(), visibleLines.end(), _sliceFirstLine, [](const VisibleLine& vl, size_t src) { return vl.sourceIndex < src; });
    const auto sliceEndIt =
        std::lower_bound(visibleLines.begin(), visibleLines.end(), _sliceLastLine, [](const VisibleLine& vl, size_t src) { return vl.sourceIndex < src; });

    if (sliceBeginIt == visibleLines.end() || sliceEndIt == visibleLines.end())
    {
        EnsureLayoutAsync();
        return;
    }

    const size_t sliceFirstVisIdx = static_cast<size_t>(std::distance(visibleLines.begin(), sliceBeginIt));
    const size_t sliceLastVisIdx  = static_cast<size_t>(std::distance(visibleLines.begin(), sliceEndIt));

    // Keep a margin so we don't thrash near the edges (in visible-line space).
    const size_t margin          = 8; // visible lines
    const bool comfortablyInside = (visStartIdx >= (sliceFirstVisIdx + margin)) && (visEndIdx + margin <= sliceLastVisIdx);
    if (! comfortablyInside)
    {
        EnsureLayoutAsync();
    }
}

void ColorTextView::StartLayoutWorker(float layoutWidth, UINT32 seq)
{
    // TRACER;
    const float safeWidth = std::clamp(layoutWidth, kMinLayoutWidthDip, kMaxLayoutWidthDip);

    // Decide slice range on UI thread for stability. Always virtualize
    size_t firstLine = 0;
    size_t lastLine  = _document.TotalLineCount() ? _document.TotalLineCount() - 1 : 0;
    if (_document.TotalLineCount() > 0)
    {
        // Map visible pixel coordinates to display rows, then to logical lines
        // IMPORTANT: Display rows != logical lines when lines contain embedded newlines
        const float lineHeight = GetLineHeight();
        const float viewTop    = std::max(0.f, _scrollY - _padding);
        const float viewBottom = viewTop + _clientDipH;

        // Calculate display row range (these are DISPLAY rows, not logical line indices)
        const UINT32 topRow    = static_cast<UINT32>(std::floor(viewTop / std::max(lineHeight, 1e-3f)));
        const UINT32 bottomRow = static_cast<UINT32>(std::floor(viewBottom / std::max(lineHeight, 1e-3f)));

        // Clamp display rows to valid range
        const UINT32 maxDisplayRow    = _document.TotalDisplayRows();
        const UINT32 clampedTopRow    = std::min(topRow, maxDisplayRow > 0 ? maxDisplayRow - 1 : 0);
        const UINT32 clampedBottomRow = std::min(bottomRow, maxDisplayRow > 0 ? maxDisplayRow - 1 : 0);

#ifdef _DEBUG
        {
            auto msg = std::format(
                "StartLayoutWorker: scrollY={:.1f}, viewTop={:.1f}, viewBottom={:.1f}, topRow={}->{} , bottomRow={}->{} , maxDisplayRow={}, lineCount={}\n",
                _scrollY,
                viewTop,
                viewBottom,
                topRow,
                clampedTopRow,
                bottomRow,
                clampedBottomRow,
                maxDisplayRow,
                _document.TotalLineCount());
            OutputDebugStringA(msg.c_str());
        }
#endif

        // Map display rows to logical line indices (handles multi-row lines correctly)
        const size_t startVisIdx = _document.VisibleIndexFromDisplayRow(clampedTopRow);
        const size_t endVisIdx   = _document.VisibleIndexFromDisplayRow(clampedBottomRow);

        // Safety check: ensure visible lines exist
        const auto& visibleLines = _document.VisibleLines();
        if (visibleLines.empty() || startVisIdx >= visibleLines.size() || endVisIdx >= visibleLines.size())
            return; // No visible lines to render

        const size_t visCount     = visibleLines.size();
        const size_t wantFirstVis = startVisIdx > kSlicePrefetchMargin ? (startVisIdx - kSlicePrefetchMargin) : 0;
        const size_t wantLastVis  = std::min(visCount - 1, endVisIdx + kSlicePrefetchMargin);
        // Align to block boundaries (in visible-line space) for better cache reuse under heavy filtering.
        const size_t alignedFirstVis = (wantFirstVis / kSliceBlockLines) * kSliceBlockLines;
        const size_t alignedLastVis  = std::min(visCount - 1, ((wantLastVis / kSliceBlockLines) * kSliceBlockLines) + kSliceBlockLines - 1);
        firstLine                    = visibleLines[alignedFirstVis].sourceIndex;
        lastLine                     = visibleLines[alignedLastVis].sourceIndex;
    }

    // Try cache first for virtual slices
    for (auto it = _layoutCache.begin(); it != _layoutCache.end(); ++it)
    {
        if (it->firstLine == firstLine && it->lastLine == lastLine && it->layout)
        {
            // Promote to most-recent (move to back)
            CachedSlice cached = std::move(*it);
            _layoutCache.erase(it);
            _layoutCache.push_back(std::move(cached));
            // Synthesize packet to avoid worker
            auto pkt                  = wil::make_unique_nothrow<LayoutPacket>();
            pkt->layout               = _layoutCache.back().layout;
            pkt->seq                  = seq;
            pkt->sliceStartPos        = _layoutCache.back().sliceStartPos;
            pkt->sliceEndPos          = _layoutCache.back().sliceEndPos;
            pkt->sliceFirstLine       = _layoutCache.back().firstLine;
            pkt->sliceLastLine        = _layoutCache.back().lastLine;
            pkt->sliceFirstDisplayRow = _layoutCache.back().firstDisplayRow;
            pkt->sliceIsFiltered      = _layoutCache.back().isFiltered;
            pkt->filteredRuns         = _layoutCache.back().filteredRuns;
            PostMessage(_hWnd, WndMsg::kColorTextViewLayoutReady, reinterpret_cast<WPARAM>(pkt.release()), 0);
            return;
        }
    }

    // CRITICAL FIX: Capture text data on UI thread BEFORE spawning worker
    std::wstring textCopy;
    UINT32 sliceStartPos        = 0;
    UINT32 sliceEndPos          = 0;
    UINT32 sliceFirstDisplayRow = 0;
    bool sliceIsFiltered        = false;
    std::vector<FilteredTextRun> filteredRuns;
    {
        // All document access happens here on UI thread with proper locking
        // CRITICAL: When filtering is active, build text from VISIBLE lines only
        // and track the display row offset for correct Y positioning.
        // ApplyColoringToLayout will compute layout-relative positions for filtered mode.

        if (_document.GetFilterMask() != Debug::InfoParam::Type::All)
        {
            // FILTERING MODE: Build text from visible lines only
            sliceIsFiltered          = true;
            const auto& visibleLines = _document.VisibleLines();
            const auto visBegin =
                std::lower_bound(visibleLines.begin(), visibleLines.end(), firstLine, [](const VisibleLine& vl, size_t src) { return vl.sourceIndex < src; });
            const auto visEnd =
                std::upper_bound(visibleLines.begin(), visibleLines.end(), lastLine, [](size_t src, const VisibleLine& vl) { return src < vl.sourceIndex; });

            if (visBegin == visEnd)
            {
                sliceFirstDisplayRow = _document.DisplayRowForSource(firstLine);
                sliceStartPos        = 0;
                sliceEndPos          = 0;
            }
            else
            {
                sliceFirstDisplayRow = visBegin->displayRowStart;
                sliceStartPos        = _document.GetLineStartOffset(visBegin->sourceIndex);
                const auto& lastVis  = *std::prev(visEnd);
                const auto& last     = _document.GetSourceLine(lastVis.sourceIndex);
                sliceEndPos          = _document.GetLineStartOffset(lastVis.sourceIndex) + _document.PrefixLength(last) + static_cast<UINT32>(last.text.size());

                filteredRuns.reserve(static_cast<size_t>(std::distance(visBegin, visEnd)));
                for (auto it = visBegin; it != visEnd; ++it)
                {
                    const size_t allIdx      = it->sourceIndex;
                    const auto& displayText  = _document.GetDisplayTextRefAll(allIdx);
                    const UINT32 layoutStart = static_cast<UINT32>(textCopy.size());
                    const UINT32 runLen      = static_cast<UINT32>(displayText.size()) + 1u; // +1 for '\n' (trimmed for last run below)
                    const UINT32 sourceStart = _document.GetLineStartOffset(allIdx);
                    filteredRuns.push_back(FilteredTextRun{
                        .sourceLine  = allIdx,
                        .layoutStart = layoutStart,
                        .length      = runLen,
                        .sourceStart = sourceStart,
                    });
                    textCopy.append(displayText);
                    textCopy.append(L"\n");
                }
            }

            // Remove trailing newline (layout uses separators between visible lines only)
            if (! textCopy.empty())
            {
                textCopy.pop_back();
                if (! filteredRuns.empty() && filteredRuns.back().length > 0)
                    --filteredRuns.back().length;
            }
        }
        else
        {
            // NO FILTERING: Use position-based range (original logic)
            sliceIsFiltered      = false;
            sliceFirstDisplayRow = _document.DisplayRowForSource(firstLine);
            sliceStartPos        = _document.GetLineStartOffset(firstLine);
            const auto& last     = _document.GetSourceLine(lastLine);
            sliceEndPos          = _document.GetLineStartOffset(lastLine) + _document.PrefixLength(last) + static_cast<UINT32>(last.text.size());
            textCopy             = _document.GetTextRange(sliceStartPos, sliceEndPos - sliceStartPos);
        }
        textCopy.erase(std::remove(textCopy.begin(), textCopy.end(), L'\r'), textCopy.end());
    }

    // Worker context now includes captured text
    struct Ctx
    {
        ColorTextView* self;
        float width;
        UINT32 seq;
        size_t firstLine;
        size_t lastLine;
        std::wstring text;    // Pre-captured text
        UINT32 sliceStartPos; // Pre-captured positions
        UINT32 sliceEndPos;
        UINT32 sliceFirstDisplayRow; // Display row offset for filtered rendering
        bool sliceIsFiltered;        // True if text contains only visible lines
        std::vector<FilteredTextRun> filteredRuns;
    };

    auto ctx = std::make_unique<Ctx>(Ctx{this,
                                         safeWidth,
                                         seq,
                                         firstLine,
                                         lastLine,
                                         std::move(textCopy),
                                         sliceStartPos,
                                         sliceEndPos,
                                         sliceFirstDisplayRow,
                                         sliceIsFiltered,
                                         std::move(filteredRuns)});

    auto rawCtx = ctx.release();

    auto worker = [](PTP_CALLBACK_INSTANCE, PVOID pCtxParam) noexcept -> void
    {
        std::unique_ptr<Ctx> ctx(static_cast<Ctx*>(pCtxParam));
        auto* self            = ctx->self;
        const float width     = std::clamp(ctx->width, kMinLayoutWidthDip, kMaxLayoutWidthDip);
        const UINT32 seqLocal = ctx->seq;
        auto res              = wil::CoInitializeEx();
        wil::com_ptr<IDWriteTextLayout> lay;

        if (self->_dwriteFactory && ! ctx->text.empty())
        {
            // Use pre-captured text - NO document access in worker thread
            self->_dwriteFactory->CreateTextLayout(
                ctx->text.c_str(), static_cast<UINT32>(ctx->text.size()), self->_textFormat.get(), width, 1000000.f, lay.addressof());

            auto pkt = wil::make_unique_nothrow<LayoutPacket>();
            pkt->layout.swap(lay);
            pkt->seq                  = seqLocal;
            pkt->sliceStartPos        = ctx->sliceStartPos;
            pkt->sliceEndPos          = ctx->sliceEndPos;
            pkt->sliceFirstLine       = ctx->firstLine;
            pkt->sliceLastLine        = ctx->lastLine;
            pkt->sliceFirstDisplayRow = ctx->sliceFirstDisplayRow;
            pkt->sliceIsFiltered      = ctx->sliceIsFiltered;
            pkt->filteredRuns         = std::move(ctx->filteredRuns);
            PostMessage(self->_hWnd, WndMsg::kColorTextViewLayoutReady, reinterpret_cast<WPARAM>(pkt.release()), 0);
            return;
        }

        // Failure path
        auto pkt                  = wil::make_unique_nothrow<LayoutPacket>();
        pkt->seq                  = seqLocal;
        pkt->sliceStartPos        = 0;
        pkt->sliceEndPos          = 0;
        pkt->sliceFirstLine       = 0;
        pkt->sliceLastLine        = 0;
        pkt->sliceFirstDisplayRow = 0;
        pkt->sliceIsFiltered      = false;
        pkt->filteredRuns.clear();
        pkt->layout = nullptr;
        PostMessage(self->_hWnd, WndMsg::kColorTextViewLayoutReady, reinterpret_cast<WPARAM>(pkt.release()), 0);
    };

    if (! TrySubmitThreadpoolCallback(worker, rawCtx, nullptr))
    {
        // Fallback: execute synchronously on calling thread
        worker(nullptr, rawCtx);
    }
}
namespace
{
inline D2D1_COLOR_F MetaColorForType(const ColorTextView::Theme& th, Debug::InfoParam::Type t)
{
    switch (t)
    {
        case Debug::InfoParam::Type::Error: return th.metaError;
        case Debug::InfoParam::Type::Warning: return th.metaWarning;
        case Debug::InfoParam::Type::Info: return th.metaInfo;
        case Debug::InfoParam::Type::Debug: return th.metaDebug;
        case Debug::InfoParam::Type::Text:
        case Debug::InfoParam::Type::All: // Not a message type, use default
        default: return th.metaText;
    }
}
} // namespace

void ColorTextView::ClearTextLayoutEffects()
{
    if (! _textLayout)
        return;
    // Compute text length from line metrics (DWRITE_TEXT_METRICS has no textLength)
    UINT32 lineCount = 0;
    _textLayout->GetLineMetrics(nullptr, 0, &lineCount);
    if (lineCount == 0)
        return;
    std::vector<DWRITE_LINE_METRICS> lm(lineCount);
    if (FAILED(_textLayout->GetLineMetrics(lm.data(), lineCount, &lineCount)))
        return;
    UINT32 totalLen = 0;
    for (UINT32 i = 0; i < lineCount; ++i)
        totalLen += lm[i].length;
    DWRITE_TEXT_RANGE r{0, totalLen};
    _textLayout->SetDrawingEffect(nullptr, r);
}

void ColorTextView::ApplyColoringToLayout()
{
    if (! _textLayout || ! _d2dCtx)
        return;

#ifdef _DEBUG
    // Clear previous debug spans
    ClearDebugSpans();
#endif

    // Clear previous drawing effects by resetting layout (cheap)
    // Re-apply default first
    // (DirectWrite has no "clear effects" call; we just overwrite spans)

    if (_sliceIsFiltered)
    {
        // FILTERED MODE: Layout contains only visible lines
        // Use the captured mapping runs to compute layout-relative positions.

#ifdef _DEBUG
        size_t spanIndex = 0;
#endif

        for (const auto& run : _sliceFilteredRuns)
        {
            if (run.length == 0 || run.sourceLine >= _document.TotalLineCount())
                continue;

            const auto& line            = _document.GetSourceLine(run.sourceLine);
            const UINT32 prefixLen      = _document.PrefixLength(line);
            const UINT32 textLen        = static_cast<UINT32>(line.text.size());
            const UINT32 lineContentLen = prefixLen + textLen;
            const UINT32 layoutLineEnd  = run.layoutStart + std::min(run.length, lineContentLen);

            // 1) Color metadata prefix at layout-relative position
            if (line.hasMeta && prefixLen > 0)
            {
                const auto color               = MetaColorForType(_theme, line.meta.type);
                ID2D1SolidColorBrush* rawBrush = GetBrush(color);
                if (rawBrush)
                {
                    const UINT32 prefixLenClamped = std::min(prefixLen, run.length);
                    if (prefixLenClamped > 0)
                    {
                        DWRITE_TEXT_RANGE r{run.layoutStart, prefixLenClamped};
                        _textLayout->SetDrawingEffect(rawBrush, r);
                    }
                }
            }

            // 2) Apply user color spans at layout-relative positions
            for (const auto& span : line.spans)
            {
                if (_d2dCtx)
                {
                    ID2D1SolidColorBrush* rawBrush = GetBrush(span.color);
                    if (rawBrush)
                    {
                        // Span positions are relative to line text (after prefix)
                        UINT32 spanStart = run.layoutStart + prefixLen + span.start;
                        UINT32 spanEnd   = spanStart + span.length;
                        spanEnd          = std::min(spanEnd, layoutLineEnd);
                        if (spanEnd > spanStart && spanStart < layoutLineEnd)
                        {
                            const UINT32 hitLen = spanEnd - spanStart;
                            DWRITE_TEXT_RANGE r{spanStart, hitLen};
                            _textLayout->SetDrawingEffect(rawBrush, r);

#ifdef _DEBUG
                            // Store debug background rectangle for this span
                            const D2D1_COLOR_F& debugColor = debugColors[spanIndex % debugColorCount];
                            DWRITE_HIT_TEST_METRICS hitMetrics[64];
                            UINT32 hitCount   = 0;
                            const float yBase = static_cast<float>(_sliceFirstDisplayRow) * GetLineHeight();
                            _textLayout->HitTestTextRange(spanStart, hitLen, 0, yBase, hitMetrics, 64, &hitCount);
                            hitCount = std::min(hitCount, 64u);
                            for (UINT32 i = 0; i < hitCount; ++i)
                            {
                                DebugSpanRect debugRect;
                                debugRect.rect = D2D1::RectF(
                                    hitMetrics[i].left, hitMetrics[i].top, hitMetrics[i].left + hitMetrics[i].width, hitMetrics[i].top + hitMetrics[i].height);
                                debugRect.color = debugColor;
                                _debugSpanRects.push_back(debugRect);
                            }
                            ++spanIndex;
#endif
                        }
                    }
                }
            }
        }
    }
    else
    {
        // NON-FILTERED MODE: Original document position-based logic
        size_t beginLine = _document.GetLineAndOffset(_sliceStartPos).first;
        size_t endLine   = _document.GetLineAndOffset(_sliceEndPos ? _sliceEndPos - 1 : 0).first;

#ifdef _DEBUG
        size_t spanIndex = 0;
#endif

        for (size_t lineIdx = beginLine; lineIdx <= endLine && lineIdx < _document.TotalLineCount(); ++lineIdx)
        {
            const auto& line  = _document.GetSourceLine(lineIdx);
            UINT32 lineOffset = _document.GetLineStartOffset(lineIdx);

            // 1) Color metadata prefix, if any
            if (line.hasMeta)
            {
                const UINT32 plen = _document.PrefixLength(line);
                if (plen)
                {
                    const auto color               = MetaColorForType(_theme, line.meta.type);
                    ID2D1SolidColorBrush* rawBrush = GetBrush(color);
                    UINT32 absStart                = lineOffset;
                    UINT32 absEnd                  = absStart + plen;
                    if (! (absEnd <= _sliceStartPos || absStart >= _sliceEndPos))
                    {
                        absStart = std::max(absStart, _sliceStartPos);
                        absEnd   = std::min(absEnd, _sliceEndPos);
                        if (absEnd > absStart)
                        {
                            DWRITE_TEXT_RANGE r{absStart - _sliceStartPos, absEnd - absStart};
                            _textLayout->SetDrawingEffect(rawBrush, r);
                        }
                    }
                }
            }

            // 2) Apply user color spans (if any)
            for (const auto& span : line.spans)
            {
                if (_d2dCtx)
                {
                    ID2D1SolidColorBrush* rawBrush = GetBrush(span.color);
                    // If using a virtualized slice, adjust range to local positions and clamp to slice
                    UINT32 absStart = lineOffset + span.start;
                    UINT32 absEnd   = absStart + span.length;
                    if (absEnd <= _sliceStartPos || absStart >= _sliceEndPos)
                        continue; // outside slice
                    absStart = std::max(absStart, _sliceStartPos);
                    absEnd   = std::min(absEnd, _sliceEndPos);
                    if (absEnd <= absStart)
                        continue;
                    DWRITE_TEXT_RANGE r{absStart - _sliceStartPos, absEnd - absStart};
                    _textLayout->SetDrawingEffect(rawBrush, r);

#ifdef _DEBUG
                    // Store debug background rectangle for this span
                    const D2D1_COLOR_F& debugColor = debugColors[spanIndex % debugColorCount];
                    DWRITE_HIT_TEST_METRICS hitMetrics[64];
                    UINT32 hitCount   = 0;
                    const float yBase = static_cast<float>(_sliceFirstDisplayRow) * GetLineHeight();
                    _textLayout->HitTestTextRange(absStart - _sliceStartPos, absEnd - absStart, 0, yBase, hitMetrics, 64, &hitCount);
                    hitCount = std::min(hitCount, 64u);
                    for (UINT32 i = 0; i < hitCount; ++i)
                    {
                        DebugSpanRect debugRect;
                        debugRect.rect = D2D1::RectF(
                            hitMetrics[i].left, hitMetrics[i].top, hitMetrics[i].left + hitMetrics[i].width, hitMetrics[i].top + hitMetrics[i].height);
                        debugRect.color = debugColor;
                        _debugSpanRects.push_back(debugRect);
                    }
                    ++spanIndex;
#endif
                }
            }
        } // end for loop over lines
    } // end else (non-filtered mode)
} // end ApplyColoringToLayout

void ColorTextView::DrawHighlights()
{
    if (! _d2dCtx || _matches.empty())
        return;

    IDWriteTextLayout* layout                = nullptr;
    float yBase                              = 0.f;
    bool isFiltered                          = false;
    UINT32 sourceStart                       = 0;
    UINT32 sourceEnd                         = 0;
    const std::vector<FilteredTextRun>* runs = nullptr;

    auto [visStartLine, visEndLine] = GetVisibleLineRange();
    const bool sliceCoversView      = _textLayout && (_sliceFirstLine <= visStartLine) && (_sliceLastLine >= visEndLine);

    if (sliceCoversView)
    {
        layout      = _textLayout.get();
        yBase       = static_cast<float>(_sliceFirstDisplayRow) * GetLineHeight();
        isFiltered  = _sliceIsFiltered;
        sourceStart = _sliceStartPos;
        sourceEnd   = _sliceEndPos;
        runs        = &_sliceFilteredRuns;
    }
    else
    {
        CreateFallbackLayoutIfNeeded(visStartLine, visEndLine);
        if (_fallbackValid && _fallbackLayout)
        {
            layout                  = _fallbackLayout.get();
            const UINT32 displayRow = _document.DisplayRowForSource(_fallbackStartLine);
            yBase                   = static_cast<float>(displayRow) * GetLineHeight();
            isFiltered              = (_document.GetFilterMask() != Debug::InfoParam::Type::All);
            sourceStart             = _document.GetLineStartOffset(_fallbackStartLine);
            const auto& last        = _document.GetSourceLine(_fallbackEndLine);
            sourceEnd               = _document.GetLineStartOffset(_fallbackEndLine) + _document.PrefixLength(last) + static_cast<UINT32>(last.text.size());
            runs                    = &_fallbackFilteredRuns;
        }
        else
        {
            layout      = _textLayout.get();
            yBase       = static_cast<float>(_sliceFirstDisplayRow) * GetLineHeight();
            isFiltered  = _sliceIsFiltered;
            sourceStart = _sliceStartPos;
            sourceEnd   = _sliceEndPos;
            runs        = &_sliceFilteredRuns;
        }
    }

    if (! layout)
        return;

    const auto highlightBrush = GetBrush(_theme.searchHighlight);
    if (! highlightBrush)
        return;

    ID2D1SolidColorBrush* activeBrush = highlightBrush;
    {
        D2D1_COLOR_F active = _theme.searchHighlight;
        active.a            = std::min(1.0f, active.a * 2.0f);
        if (active.a != _theme.searchHighlight.a)
        {
            if (auto* b = GetBrush(active))
                activeBrush = b;
        }
    }

    const bool hasActive     = (_matchIndex >= 0) && (static_cast<size_t>(_matchIndex) < _matches.size());
    const size_t activeIndex = hasActive ? static_cast<size_t>(_matchIndex) : static_cast<size_t>(-1);

    // Budget the amount of rectangles we fill per frame
    UINT32 rectBudget = 512;

    auto drawHitTestRange = [&](ID2D1SolidColorBrush* brush, UINT32 layoutStart, UINT32 length)
    {
        if (! brush || length == 0 || rectBudget == 0)
            return;

        std::array<DWRITE_HIT_TEST_METRICS, 64> buf{};
        UINT32 hitCount  = 0;
        const HRESULT hr = layout->HitTestTextRange(layoutStart, length, 0, yBase, buf.data(), static_cast<UINT32>(buf.size()), &hitCount);
        if (SUCCEEDED(hr))
        {
            hitCount = std::min(hitCount, static_cast<UINT32>(buf.size()));
            hitCount = std::min(hitCount, rectBudget);
            for (UINT32 i = 0; i < hitCount; ++i)
            {
                const auto rc = D2D1::RectF(buf[i].left, buf[i].top, buf[i].left + buf[i].width, buf[i].top + buf[i].height);
                _d2dCtx->FillRectangle(rc, brush);
            }
            rectBudget -= hitCount;
            return;
        }

        if (hitCount == 0 || hitCount > rectBudget)
            return;

        std::vector<DWRITE_HIT_TEST_METRICS> big(hitCount);
        UINT32 hitCount2 = 0;
        if (SUCCEEDED(layout->HitTestTextRange(layoutStart, length, 0, yBase, big.data(), hitCount, &hitCount2)))
        {
            hitCount2 = std::min(hitCount2, static_cast<UINT32>(big.size()));
            hitCount2 = std::min(hitCount2, rectBudget);
            for (UINT32 i = 0; i < hitCount2; ++i)
            {
                const auto rc = D2D1::RectF(big[i].left, big[i].top, big[i].left + big[i].width, big[i].top + big[i].height);
                _d2dCtx->FillRectangle(rc, brush);
            }
            rectBudget -= hitCount2;
        }
    };

    auto drawMatch = [&](const Line::ColorSpan& range, ID2D1SolidColorBrush* brush)
    {
        if (rectBudget == 0 || ! brush)
            return;

        const UINT32 matchStart = range.start;
        const UINT32 matchEnd   = range.start + range.length;

        if (! isFiltered)
        {
            const UINT32 rangeStart = std::max(matchStart, sourceStart);
            const UINT32 rangeEnd   = std::min(matchEnd, sourceEnd);
            if (rangeEnd <= rangeStart)
                return;

            const UINT32 localStart = rangeStart - sourceStart;
            drawHitTestRange(brush, localStart, rangeEnd - rangeStart);
            return;
        }

        if (! runs || runs->empty())
            return;

        for (const auto& run : *runs)
        {
            if (rectBudget == 0)
                break;

            const UINT32 runStart   = run.sourceStart;
            const UINT32 runEnd     = run.sourceStart + run.length;
            const UINT32 rangeStart = std::max(matchStart, runStart);
            const UINT32 rangeEnd   = std::min(matchEnd, runEnd);
            if (rangeEnd <= rangeStart)
                continue;

            const UINT32 layoutStart = run.layoutStart + (rangeStart - runStart);
            drawHitTestRange(brush, layoutStart, rangeEnd - rangeStart);
        }
    };

    // Draw active match first so it isn't starved by the rect budget.
    if (hasActive)
    {
        drawMatch(_matches[activeIndex], activeBrush);
    }

    for (size_t i = 0; i < _matches.size(); ++i)
    {
        if (rectBudget == 0)
            break;
        if (hasActive && i == activeIndex)
            continue;
        drawMatch(_matches[i], highlightBrush);
    }
}

void ColorTextView::DrawSelection()
{
    if (! _d2dCtx || _selStart == _selEnd)
        return;

    IDWriteTextLayout* layout                = nullptr;
    float yBase                              = 0.f;
    bool isFiltered                          = false;
    UINT32 sourceStart                       = 0;
    UINT32 sourceEnd                         = 0;
    const std::vector<FilteredTextRun>* runs = nullptr;

    auto [visStartLine, visEndLine] = GetVisibleLineRange();
    const bool sliceCoversView      = _textLayout && (_sliceFirstLine <= visStartLine) && (_sliceLastLine >= visEndLine);

    if (sliceCoversView)
    {
        layout      = _textLayout.get();
        yBase       = static_cast<float>(_sliceFirstDisplayRow) * GetLineHeight();
        isFiltered  = _sliceIsFiltered;
        sourceStart = _sliceStartPos;
        sourceEnd   = _sliceEndPos;
        runs        = &_sliceFilteredRuns;
    }
    else
    {
        CreateFallbackLayoutIfNeeded(visStartLine, visEndLine);
        if (_fallbackValid && _fallbackLayout)
        {
            layout                  = _fallbackLayout.get();
            const UINT32 displayRow = _document.DisplayRowForSource(_fallbackStartLine);
            yBase                   = static_cast<float>(displayRow) * GetLineHeight();
            isFiltered              = (_document.GetFilterMask() != Debug::InfoParam::Type::All);
            sourceStart             = _document.GetLineStartOffset(_fallbackStartLine);
            const auto& last        = _document.GetSourceLine(_fallbackEndLine);
            sourceEnd               = _document.GetLineStartOffset(_fallbackEndLine) + _document.PrefixLength(last) + static_cast<UINT32>(last.text.size());
            runs                    = &_fallbackFilteredRuns;
        }
        else
        {
            layout      = _textLayout.get();
            yBase       = static_cast<float>(_sliceFirstDisplayRow) * GetLineHeight();
            isFiltered  = _sliceIsFiltered;
            sourceStart = _sliceStartPos;
            sourceEnd   = _sliceEndPos;
            runs        = &_sliceFilteredRuns;
        }
    }

    if (! layout)
        return;

    UINT32 start      = std::min(_selStart, _selEnd);
    UINT32 end        = std::max(_selStart, _selEnd);
    const auto vr     = GetVisibleTextRange();
    UINT32 rangeStart = std::max(start, vr.first);
    UINT32 rangeEnd   = std::min(end, vr.second);
    if (rangeEnd <= rangeStart)
        return;

    const auto selectionBrush = GetBrush(_theme.selection);
    if (! selectionBrush)
        return;

    constexpr float kSelectionCornerRadiusDip = 2.0f;

    auto drawHitTestRange = [&](UINT32 layoutStart, UINT32 length, UINT32& rectBudget)
    {
        if (length == 0 || rectBudget == 0)
            return;

        std::array<DWRITE_HIT_TEST_METRICS, 64> buf{};
        UINT32 hitCount  = 0;
        const HRESULT hr = layout->HitTestTextRange(layoutStart, length, 0, yBase, buf.data(), static_cast<UINT32>(buf.size()), &hitCount);
        if (SUCCEEDED(hr))
        {
            hitCount = std::min(hitCount, static_cast<UINT32>(buf.size()));
            hitCount = std::min(hitCount, rectBudget);
            for (UINT32 i = 0; i < hitCount; ++i)
            {
                const auto rc                 = D2D1::RectF(buf[i].left, buf[i].top, buf[i].left + buf[i].width, buf[i].top + buf[i].height);
                const float rectWidth         = std::max(0.0f, rc.right - rc.left);
                const float rectHeight        = std::max(0.0f, rc.bottom - rc.top);
                const float maxCornerRadius   = std::min(rectWidth, rectHeight) * 0.5f;
                const float cornerRadius      = std::min(kSelectionCornerRadiusDip, maxCornerRadius);
                const D2D1_ROUNDED_RECT round = D2D1::RoundedRect(rc, cornerRadius, cornerRadius);
                _d2dCtx->FillRoundedRectangle(round, selectionBrush);
            }
            rectBudget -= hitCount;
            return;
        }

        if (hitCount == 0 || hitCount > rectBudget)
            return;

        std::vector<DWRITE_HIT_TEST_METRICS> big(hitCount);
        UINT32 hitCount2 = 0;
        if (SUCCEEDED(layout->HitTestTextRange(layoutStart, length, 0, yBase, big.data(), hitCount, &hitCount2)))
        {
            hitCount2 = std::min(hitCount2, static_cast<UINT32>(big.size()));
            hitCount2 = std::min(hitCount2, rectBudget);
            for (UINT32 i = 0; i < hitCount2; ++i)
            {
                const auto rc                 = D2D1::RectF(big[i].left, big[i].top, big[i].left + big[i].width, big[i].top + big[i].height);
                const float rectWidth         = std::max(0.0f, rc.right - rc.left);
                const float rectHeight        = std::max(0.0f, rc.bottom - rc.top);
                const float maxCornerRadius   = std::min(rectWidth, rectHeight) * 0.5f;
                const float cornerRadius      = std::min(kSelectionCornerRadiusDip, maxCornerRadius);
                const D2D1_ROUNDED_RECT round = D2D1::RoundedRect(rc, cornerRadius, cornerRadius);
                _d2dCtx->FillRoundedRectangle(round, selectionBrush);
            }
            rectBudget -= hitCount2;
        }
    };

    UINT32 rectBudget = 1024;
    if (! isFiltered)
    {
        const UINT32 clampedStart = std::max(rangeStart, sourceStart);
        const UINT32 clampedEnd   = std::min(rangeEnd, sourceEnd);
        if (clampedEnd <= clampedStart)
            return;

        drawHitTestRange(clampedStart - sourceStart, clampedEnd - clampedStart, rectBudget);
        return;
    }

    if (! runs || runs->empty())
        return;

    for (const auto& run : *runs)
    {
        if (rectBudget == 0)
            break;

        const UINT32 runStart = run.sourceStart;
        const UINT32 runEnd   = run.sourceStart + run.length;
        const UINT32 segStart = std::max(rangeStart, runStart);
        const UINT32 segEnd   = std::min(rangeEnd, runEnd);
        if (segEnd <= segStart)
            continue;

        const UINT32 layoutStart = run.layoutStart + (segStart - runStart);
        drawHitTestRange(layoutStart, segEnd - segStart, rectBudget);
    }
}

std::pair<size_t, size_t> ColorTextView::GetVisibleLineRange() const
{
    if (_document.TotalLineCount() == 0)
        return {0, 0};
    const float H          = GetLineHeight();
    const float viewTop    = _scrollY - _padding;
    const float viewBottom = viewTop + (_clientDipH - (_horzScrollbarVisible ? GetHorzScrollbarDip(_hWnd, _dpi) : 0.f));
    const float rowH       = std::max(H, 1e-3f);

    // Calculate display row range (not logical line range!)
    const UINT32 topRow    = static_cast<UINT32>(std::floor(std::max(0.f, viewTop) / rowH));
    const UINT32 bottomRow = static_cast<UINT32>(std::floor(viewBottom / rowH));

    // Clamp to valid display row range
    const UINT32 maxDisplayRow    = _document.TotalDisplayRows();
    const UINT32 clampedTopRow    = std::min(topRow, maxDisplayRow > 0 ? maxDisplayRow - 1 : 0);
    const UINT32 clampedBottomRow = std::min(bottomRow, maxDisplayRow > 0 ? maxDisplayRow - 1 : 0);

    // Map display rows to logical line indices (source indices)
    const size_t topVisIdx    = _document.VisibleIndexFromDisplayRow(clampedTopRow);
    const size_t bottomVisIdx = _document.VisibleIndexFromDisplayRow(clampedBottomRow);

    // Safety check: ensure visible lines exist
    if (_document.VisibleLines().empty() || topVisIdx >= _document.VisibleLines().size() || bottomVisIdx >= _document.VisibleLines().size())
        return {0, 0}; // Return empty range if no visible lines

    const size_t visStart = _document.VisibleLines()[topVisIdx].sourceIndex;
    const size_t visEnd   = _document.VisibleLines()[bottomVisIdx].sourceIndex;

#ifdef _DEBUG
    {
        auto msg = std::format("GetVisibleLineRange: scrollY={:.1f}, viewTop={:.1f}, viewBottom={:.1f}, topRow={}->{} , bottomRow={}->{} , maxDisplayRow={}, "
                               "visStart={}, visEnd={}\n",
                               _scrollY,
                               viewTop,
                               viewBottom,
                               topRow,
                               clampedTopRow,
                               bottomRow,
                               clampedBottomRow,
                               maxDisplayRow,
                               visStart,
                               visEnd);
        OutputDebugStringA(msg.c_str());
    }
#endif

    return {visStart, visEnd};
}

std::pair<UINT32, UINT32> ColorTextView::GetVisibleTextRange() const
{
    // virtualized: approximate by whole logical lines in view
    auto [visStart, visEnd] = GetVisibleLineRange();
    UINT32 startPos         = _document.TotalLineCount() ? _document.GetLineStartOffset(visStart) : 0u;
    UINT32 endPos           = 0u;
    if (_document.TotalLineCount() > 0)
    {
        const auto& line = _document.GetSourceLine(visEnd);
        endPos           = _document.GetLineStartOffset(visEnd) + _document.PrefixLength(line) + static_cast<UINT32>(line.text.size());
    }
    return {startPos, endPos};
}

ID2D1SolidColorBrush* ColorTextView::GetBrush(const D2D1_COLOR_F& color)
{
    // TRACER;

    if (! _d2dCtx)
        return nullptr;
    BrushCacheKey key{color.r, color.g, color.b, color.a};
    if (auto it = _brushCache.find(key); it != _brushCache.end())
    {
        it->second.lastAccess = ++_brushAccessCounter;
        return it->second.brush.get();
    }

    wil::com_ptr<ID2D1SolidColorBrush> b;
    if (FAILED(_d2dCtx->CreateSolidColorBrush(color, b.addressof())) || ! b)
        return nullptr;
    PruneBrushCacheIfNeeded();
    BrushCacheEntry entry{};
    entry.brush      = std::move(b);
    entry.lastAccess = ++_brushAccessCounter;
    auto [insIt, ok] = _brushCache.emplace(key, std::move(entry));
    return insIt->second.brush.get();
}

void ColorTextView::PruneBrushCacheIfNeeded()
{
    constexpr size_t kMaxBrushCacheSize = 256;
    if (_brushCache.size() < kMaxBrushCacheSize)
        return;
    // Build vector of (lastAccess, key)
    std::vector<std::pair<uint64_t, BrushCacheKey>> v;
    v.reserve(_brushCache.size());
    for (const auto& kv : _brushCache)
        v.emplace_back(kv.second.lastAccess, kv.first);
    // Sort ascending (oldest first)
    std::sort(v.begin(), v.end(), [](const auto& a, const auto& b) { return a.first < b.first; });
    // Remove oldest 25%
    const size_t removeCount = std::max<size_t>(1, kMaxBrushCacheSize / 4);
    for (size_t i = 0; i < removeCount && i < v.size(); ++i)
        _brushCache.erase(v[i].second);
}

void ColorTextView::InvalidateExposedArea(float deltaDipY)
{
    if (! _hWnd)
        return;
    const float dpiScale    = _dpi > 0.0f ? _dpi / 96.0f : 1.0f;
    const LONG deltaPx      = static_cast<LONG>(std::lround(deltaDipY * dpiScale));
    const RECT viewRectPx   = GetViewRect();
    const LONG viewHeightPx = viewRectPx.bottom - viewRectPx.top;

    const LONG absDelta = deltaPx >= 0 ? deltaPx : -deltaPx;
    if (deltaPx != 0 && viewHeightPx > 0 && absDelta < viewHeightPx)
    {
        RECT dirtyPx = viewRectPx;
        if (deltaPx > 0)
        {
            dirtyPx.top = std::max(dirtyPx.bottom - absDelta, dirtyPx.top);
        }
        else
        {
            dirtyPx.bottom = std::min(dirtyPx.top + absDelta, dirtyPx.bottom);
        }

        if (dirtyPx.bottom > dirtyPx.top)
        {
            _pendingDirtyRect = dirtyPx;
            _pendingScrollDy  = deltaPx;
            _hasPendingDirty  = true;
            _hasPendingScroll = true;
            _needsFullRedraw  = false;
            InvalidateRect(_hWnd, nullptr, FALSE);
            return;
        }
    }

    RequestFullRedraw();
    InvalidateRect(_hWnd, nullptr, FALSE);
}

void ColorTextView::UpdateGutterWidth()
{
    UINT32 targetLines = (UINT32)std::max<size_t>(1, _document.TotalLineCount());
    UINT32 digits      = 1;
    UINT32 t           = std::max(1u, targetLines);
    while (t >= 10)
    {
        ++digits;
        t /= 10;
    }
    if (digits != _gutterDigits)
    {
        _gutterDigits = digits;
        _gutterDipW   = 12.f * static_cast<float>(_gutterDigits) + 16.f;
    }
}

RECT ColorTextView::GetCaretRectPx() const
{
    RECT rc{};
    if (! _textLayout)
        return rc;

    const float dpiScale = _dpi / 96.f;
    const float tx       = _padding + (_displayLineNumbers ? _gutterDipW : 0.f) - _scrollX;
    const float ty       = _padding - _scrollY + static_cast<float>(_sliceFirstDisplayRow) * GetLineHeight();
    DWRITE_HIT_TEST_METRICS m{};
    FLOAT cx = 0;
    FLOAT cy = 0;

    std::optional<UINT32> localPos;
    if (! _sliceIsFiltered)
    {
        if (_caretPos >= _sliceStartPos && _caretPos <= _sliceEndPos)
            localPos = _caretPos - _sliceStartPos;
    }
    else if (! _sliceFilteredRuns.empty())
    {
        auto it = std::upper_bound(
            _sliceFilteredRuns.begin(), _sliceFilteredRuns.end(), _caretPos, [](UINT32 value, const FilteredTextRun& run) { return value < run.sourceStart; });
        if (it != _sliceFilteredRuns.begin())
            --it;

        const UINT32 runStart = it->sourceStart;
        const UINT32 runEnd   = it->sourceStart + it->length;
        if (_caretPos >= runStart && _caretPos <= runEnd)
            localPos = it->layoutStart + (_caretPos - runStart);
    }

    if (! localPos)
        return rc;

    _textLayout->HitTestTextPosition(*localPos, FALSE, &cx, &cy, &m);
    const float left   = (tx + cx) * dpiScale;
    const float top    = (ty + cy) * dpiScale;
    const float right  = (tx + cx + 2.f) * dpiScale;
    const float bottom = (ty + cy + m.height) * dpiScale;
    rc.left            = (int)std::floor(left);
    rc.top             = (int)std::floor(top);
    rc.right           = (int)std::ceil(right);
    rc.bottom          = (int)std::ceil(bottom);
    return rc;
}

void ColorTextView::InvalidateCaret()
{
    if (_hWnd)
    {
        RECT rc = GetCaretRectPx();
        if (! IsRectEmpty(&rc))
            InvalidateRect(_hWnd, &rc, FALSE);
    }
}

void ColorTextView::EnsureWidthAsync()
{
    if (! _dwriteFactory || ! _textFormat)
        return;
    const size_t lineCount = _document.TotalLineCount();
    if (lineCount == 0)
    {
        _lineWidthCache.clear();
        _maxMeasuredWidth   = 0.f;
        _maxMeasuredIndex   = 0;
        _approxContentWidth = 0.f;
        return;
    }

    if (_lineWidthCache.size() != lineCount)
        _lineWidthCache.resize(lineCount, 0.f);

    auto dirtyRange = _document.ExtractDirtyLineRange();
    if (! dirtyRange)
    {
        const float fallback = GetAverageCharWidth() * static_cast<float>(_document.LongestLineChars());
        const float widthDip = std::max(_maxMeasuredWidth, fallback);
        if (std::fabs(widthDip - _approxContentWidth) > 0.1f)
        {
            _approxContentWidth = widthDip;
            ClampHorizontalScroll();
            UpdateScrollBars();
        }
        return;
    }

    size_t first = std::min(dirtyRange->first, lineCount - 1);
    size_t last  = std::min(dirtyRange->second, lineCount - 1);
    if (first > last)
        std::swap(first, last);

    if (_maxMeasuredWidth > 0.f && _maxMeasuredIndex >= first && _maxMeasuredIndex <= last)
    {
        _maxMeasuredWidth = 0.f;
        _maxMeasuredIndex = 0;
    }

    std::vector<size_t> indices;
    std::vector<std::wstring> texts;
    const size_t count = last - first + 1;
    indices.reserve(count);
    texts.reserve(count);

    for (size_t idx = first; idx <= last; ++idx)
    {
        indices.push_back(idx);
        texts.emplace_back(_document.GetDisplayTextRefAll(idx));
        _lineWidthCache[idx] = 0.f;
    }

    const UINT32 seq = ++_widthSeq;

#pragma warning(push)
#pragma warning(disable : 4820) // Padding warning - trailing padding unavoidable for this struct
    struct WidthCtx
    {
        ColorTextView* self{};
        std::vector<size_t> indices;
        std::vector<std::wstring> texts;
        UINT32 seq{};
    };
#pragma warning(pop)

    auto ctx     = std::make_unique<WidthCtx>();
    ctx->self    = this;
    ctx->seq     = seq;
    ctx->indices = std::move(indices);
    ctx->texts   = std::move(texts);
    auto rawCtx  = ctx.release();

    auto worker = [](PTP_CALLBACK_INSTANCE, PVOID param) noexcept
    {
        std::unique_ptr<WidthCtx> widthCtx(static_cast<WidthCtx*>(param));
        auto* self = widthCtx->self;
        if (! self)
            return;

        auto res = wil::CoInitializeEx();

        auto pkt     = std::make_unique<ColorTextView::WidthPacket>();
        pkt->seq     = widthCtx->seq;
        pkt->indices = widthCtx->indices;
        pkt->widths.resize(widthCtx->texts.size(), 0.f);

        if (self->_dwriteFactory && self->_textFormat)
        {
            for (size_t i = 0; i < widthCtx->texts.size(); ++i)
            {
                const auto& text = widthCtx->texts[i];
                if (text.empty())
                    continue;
                wil::com_ptr<IDWriteTextLayout> tl;
                if (SUCCEEDED(self->_dwriteFactory->CreateTextLayout(
                        text.c_str(), static_cast<UINT32>(text.size()), self->_textFormat.get(), 1000000.f, 1000.f, tl.addressof())))
                {
                    DWRITE_TEXT_METRICS tm{};
                    if (SUCCEEDED(tl->GetMetrics(&tm)))
                        pkt->widths[i] = tm.widthIncludingTrailingWhitespace;
                }
            }
        }

        PostMessage(self->_hWnd, WndMsg::kColorTextViewWidthReady, reinterpret_cast<WPARAM>(pkt.release()), 0);
    };

    if (! TrySubmitThreadpoolCallback(worker, rawCtx, nullptr))
    {
        worker(nullptr, rawCtx);
    }
}
void ColorTextView::DrawLineNumbers()
{
    if (! _d2dCtx || _document.TotalLineCount() == 0)
        return;
    const auto gutterBrush = GetBrush(_theme.gutterFg);
    const float lineHeight = GetLineHeight();
    // Fast path: Compute visible logical lines from display rows.
    const float viewTop    = std::max(0.f, _scrollY - _padding);
    const float viewBottom = viewTop + _clientDipH;
    const UINT32 topRow    = static_cast<UINT32>(std::floor(viewTop / std::max(lineHeight, 1e-3f)));
    const UINT32 bottomRow = static_cast<UINT32>(std::floor(viewBottom / std::max(lineHeight, 1e-3f)));

    // Clamp display rows to valid range (fixes white lines bug)
    const UINT32 maxDisplayRow    = _document.TotalDisplayRows();
    const UINT32 clampedTopRow    = std::min(topRow, maxDisplayRow > 0 ? maxDisplayRow - 1 : 0);
    const UINT32 clampedBottomRow = std::min(bottomRow, maxDisplayRow > 0 ? maxDisplayRow - 1 : 0);

    const size_t startVisIdx = _document.VisibleIndexFromDisplayRow(clampedTopRow);
    const size_t endVisIdx   = _document.VisibleIndexFromDisplayRow(clampedBottomRow);

    // Safety check: ensure visible lines exist
    if (_document.VisibleLines().empty() || startVisIdx >= _document.VisibleLines().size() || endVisIdx >= _document.VisibleLines().size())
        return; // No visible lines to draw

    const auto& visibleLines  = _document.VisibleLines();
    const size_t clampedStart = std::min(startVisIdx, visibleLines.size() - 1);
    const size_t clampedEnd   = std::min(endVisIdx, visibleLines.size() - 1);
    const size_t startLineAll = visibleLines[clampedStart].sourceIndex;
    const size_t endLineAll   = visibleLines[clampedEnd].sourceIndex;

#ifdef _DEBUG
    static int debugCounter = 0;
    if (++debugCounter % 60 == 0) // Log every 60th frame
    {
        const UINT32 filterMask   = _document.GetFilterMask();
        const size_t visibleCount = _document.VisibleLineCount();
        const size_t totalCount   = _document.TotalLineCount();
        auto msg                  = std::format("DrawLineNumbers: topRow={}, bottomRow={}, startLine={}, endLine={}, visible={}/{}, mask=0x{:02X}\n",
                               clampedTopRow,
                               clampedBottomRow,
                               startLineAll,
                               endLineAll,
                               visibleCount,
                               totalCount,
                               filterMask);
        OutputDebugStringA(msg.c_str());
    }
#endif

    for (size_t visIdx = clampedStart; visIdx <= clampedEnd && visIdx < visibleLines.size(); ++visIdx)
    {
        const auto& vl          = visibleLines[visIdx];
        const UINT32 displayRow = vl.displayRowStart;

        // Only draw line number if the line's first display row is within visible viewport
        // This prevents line numbers from disappearing for multi-line entries whose start row
        // is above the viewport but whose text content is visible
        if (displayRow < clampedTopRow || displayRow > clampedBottomRow)
        {
#ifdef _DEBUG
            if (debugCounter % 60 == 0 && visIdx <= clampedStart + 2)
            {
                auto msg =
                    std::format("  Line {} start row {} is outside viewport [{}, {}]\n", vl.sourceIndex + 1, displayRow, clampedTopRow, clampedBottomRow);
                OutputDebugStringA(msg.c_str());
            }
#endif
            continue;
        }

        const float y          = _padding - _scrollY + static_cast<float>(displayRow) * lineHeight;
        const std::wstring txt = std::to_wstring(static_cast<UINT32>(vl.sourceIndex + 1));
        auto rc                = D2D1::RectF(2.f, y, _gutterDipW - 2.f, y + lineHeight + 2.f);

#ifdef _DEBUG
        if (debugCounter % 60 == 0 && visIdx <= clampedStart + 2)
        {
            auto msg = std::format("  Drawing line {} at y={:.1f} (displayRow={})\n", vl.sourceIndex + 1, y, displayRow);
            OutputDebugStringA(msg.c_str());
        }
#endif

        _d2dCtx->DrawText(txt.c_str(), static_cast<UINT32>(txt.size()), _gutterTextFormat.get(), rc, gutterBrush);
    }
    return;
}

void ColorTextView::DrawCaret()
{
    if (! _textLayout || ! _d2dCtx || ! _hasFocus || _selStart != _selEnd || ! _caretBlinkOn)
        return;
    // Draw caret using transform-based coordinates
    DWRITE_HIT_TEST_METRICS m{};
    FLOAT cx = 0, cy = 0;

    std::optional<UINT32> localPos;
    if (! _sliceIsFiltered)
    {
        if (_caretPos >= _sliceStartPos && _caretPos <= _sliceEndPos)
            localPos = _caretPos - _sliceStartPos;
    }
    else if (! _sliceFilteredRuns.empty())
    {
        auto it = std::upper_bound(
            _sliceFilteredRuns.begin(), _sliceFilteredRuns.end(), _caretPos, [](UINT32 value, const FilteredTextRun& run) { return value < run.sourceStart; });
        if (it != _sliceFilteredRuns.begin())
            --it;

        const UINT32 runStart = it->sourceStart;
        const UINT32 runEnd   = it->sourceStart + it->length;
        if (_caretPos >= runStart && _caretPos <= runEnd)
            localPos = it->layoutStart + (_caretPos - runStart);
    }

    if (! localPos)
        return;

    _textLayout->HitTestTextPosition(*localPos, FALSE, &cx, &cy, &m);
    const float yBase = static_cast<float>(_sliceFirstDisplayRow) * GetLineHeight();
    const float tx    = _padding + (_displayLineNumbers ? _gutterDipW : 0.f) - _scrollX;
    const float ty    = _padding - _scrollY;
    D2D1_MATRIX_3X2_F prev{};
    _d2dCtx->GetTransform(&prev);
    _d2dCtx->SetTransform(D2D1::Matrix3x2F::Translation(tx, ty));
    D2D1_RECT_F caret = D2D1::RectF(cx, yBase + cy, cx + 1.f, yBase + cy + m.height);
    _d2dCtx->FillRectangle(caret, GetBrush(_theme.caret));
    _d2dCtx->SetTransform(prev);
}

void ColorTextView::ClampScroll()
{
    const float viewDipH = _clientDipH - (_horzScrollbarVisible ? GetHorzScrollbarDip(_hWnd, _dpi) : 0.f);
    float maxY           = std::max(0.f, _contentHeight - viewDipH);
    _scrollY             = std::clamp(_scrollY, 0.f, maxY);
}

void ColorTextView::CopySelectionToClipboard()
{
    if (_selStart == _selEnd)
        return;

    UINT32 s = std::min(_selStart, _selEnd);
    UINT32 e = std::max(_selStart, _selEnd);

    if (e <= s)
        return;

    std::wstring sel;
    if (_document.GetFilterMask() == Debug::InfoParam::Type::All)
    {
        sel = _document.GetTextRange(s, e - s);
    }
    else
    {
        // Selection is stored in source-document coordinates (unfiltered space).
        // When filtering is active, that range may include hidden lines, so build the clipboard text
        // by intersecting the selection with the visible line set.
        const auto& visible = _document.VisibleLines();
        bool firstChunk     = true;
        for (const auto& vl : visible)
        {
            const size_t srcIndex = vl.sourceIndex;
            const UINT32 lineBase = _document.GetLineStartOffset(srcIndex);
            const auto& line      = _document.GetSourceLine(srcIndex);
            const UINT32 lineLen  = _document.PrefixLength(line) + static_cast<UINT32>(line.text.size());
            const UINT32 lineEnd  = lineBase + lineLen; // exclusive (no separator)

            const UINT32 segStart = std::max(s, lineBase);
            const UINT32 segEnd   = std::min(e, lineEnd);
            if (segEnd <= segStart)
                continue;

            const auto& display   = _document.GetDisplayTextRefAll(srcIndex); // prefix + text
            const UINT32 localOff = segStart - lineBase;
            const UINT32 localLen = segEnd - segStart;
            if (localOff >= display.size() || localLen == 0)
                continue;

            const UINT32 clampedLen = std::min(localLen, static_cast<UINT32>(display.size() - localOff));
            if (clampedLen == 0)
                continue;

            if (! firstChunk)
                sel.push_back(L'\n');
            firstChunk = false;
            sel.append(display, localOff, clampedLen);
        }
    }

    if (sel.empty())
        return;

    // WIL clipboard wrapper with RAII
    auto clipboard = wil::open_clipboard(_hWnd);

    // Clear clipboard first
    EmptyClipboard();

    // WIL global memory management
    wil::unique_hglobal globalMem{GlobalAlloc(GMEM_MOVEABLE, (sel.size() + 1) * sizeof(wchar_t))};
    if (globalMem)
    {
        // RAII lock wrapper
        const wil::unique_hglobal_locked lock{globalMem.get()};
        wcscpy_s(static_cast<wchar_t*>(lock.get()), sel.size() + 1, sel.c_str());
    }

    SetClipboardData(CF_UNICODETEXT, globalMem.release());
}

void ColorTextView::RebuildMatches()
{
    _matches.clear();
    _matchIndex = -1;

    if (_search.empty() || _document.TotalLineCount() == 0)
        return;

    const auto addMatchesForLine = [&](size_t sourceIndex, const Line& line)
    {
        const UINT32 lineStart = _document.GetLineStartOffset(sourceIndex);
        const UINT32 plen      = _document.PrefixLength(line);
        size_t pos             = 0;
        while (true)
        {
            pos = _searchCaseSensitive ? line.text.find(_search, pos) : FindCaseIncensitive(line.text, _search, pos);
            if (pos == std::wstring::npos)
                break;
            _matches.push_back(Line::ColorSpan{lineStart + plen + static_cast<UINT32>(pos), static_cast<UINT32>(_search.size()), _theme.searchHighlight});
            pos += _search.size();
        }
    };

    // In filtered mode, only search visible lines. Otherwise FindNext() can land on hidden lines
    // which makes caret/selection behavior confusing.
    if (_document.GetFilterMask() != Debug::InfoParam::Type::All)
    {
        for (const auto& visible : _document.VisibleLines())
        {
            const auto& line = _document.GetSourceLine(visible.sourceIndex);
            addMatchesForLine(visible.sourceIndex, line);
        }
        return;
    }

    for (size_t i = 0; i < _document.TotalLineCount(); ++i)
    {
        const auto& line = _document.GetSourceLine(i);
        addMatchesForLine(i, line);
    }
}

// ---- Message dispatch ----
void ColorTextView::OnLButtonDown(int x, int y)
{
    SetFocus(_hWnd);

    // Selecting text implies the user wants to inspect history; stop the hot-path auto-scroll mode.
    if (_renderMode == RenderMode::AUTO_SCROLL)
    {
        SwitchToScrollBackMode();
        // Kick layout immediately to make hit-testing/selection responsive.
        EnsureLayoutAdaptive(1);
    }

    _mouseDown = true;
    SetCapture(_hWnd);
    const float invScale = (_dpi > 0.0f) ? 96.0f / _dpi : 1.0f;
    UINT32 pos           = PosFromPoint(static_cast<float>(x) * invScale, static_cast<float>(y) * invScale);
    _matchIndex          = -1;
    _selStart = _selEnd = _caretPos = pos;
    _caretBlinkOn                   = true;
    Invalidate();
}

void ColorTextView::OnLButtonUp()
{
    if (_mouseDown)
    {
        _mouseDown = false;
        ReleaseCapture();
    }
}

void ColorTextView::OnMouseMove(int x, int y, WPARAM)
{
    if (! _mouseDown)
        return;
    const float invScale = (_dpi > 0.0f) ? 96.0f / _dpi : 1.0f;
    UINT32 pos           = PosFromPoint(static_cast<float>(x) * invScale, static_cast<float>(y) * invScale);
    _matchIndex          = -1;
    _selEnd              = pos;
    _caretPos            = pos;
    _caretBlinkOn        = true;
    Invalidate();
}

UINT32 ColorTextView::PosFromPoint(float x, float y)
{
    if (_document.TotalLineCount() == 0)
        return 0;

    IDWriteTextLayout* layout                = nullptr;
    float yBase                              = 0.f;
    bool isFiltered                          = false;
    UINT32 sourceBase                        = 0;
    const std::vector<FilteredTextRun>* runs = nullptr;

    auto [visStartLine, visEndLine] = GetVisibleLineRange();
    const bool sliceCoversView      = _textLayout && (_sliceFirstLine <= visStartLine) && (_sliceLastLine >= visEndLine);

    if (sliceCoversView)
    {
        layout     = _textLayout.get();
        yBase      = static_cast<float>(_sliceFirstDisplayRow) * GetLineHeight();
        isFiltered = _sliceIsFiltered;
        sourceBase = _sliceStartPos;
        runs       = &_sliceFilteredRuns;
    }
    else
    {
        CreateFallbackLayoutIfNeeded(visStartLine, visEndLine);
        if (_fallbackValid && _fallbackLayout)
        {
            layout                  = _fallbackLayout.get();
            const UINT32 displayRow = _document.DisplayRowForSource(_fallbackStartLine);
            yBase                   = static_cast<float>(displayRow) * GetLineHeight();
            isFiltered              = (_document.GetFilterMask() != Debug::InfoParam::Type::All);
            sourceBase              = _document.GetLineStartOffset(_fallbackStartLine);
            runs                    = &_fallbackFilteredRuns;
        }
        else if (_textLayout)
        {
            // Best-effort hit test against the last known slice.
            layout     = _textLayout.get();
            yBase      = static_cast<float>(_sliceFirstDisplayRow) * GetLineHeight();
            isFiltered = _sliceIsFiltered;
            sourceBase = _sliceStartPos;
            runs       = &_sliceFilteredRuns;
        }
    }

    if (! layout)
        return 0;

    const FLOAT ox = _padding + (_displayLineNumbers ? _gutterDipW : 0.f);
    const FLOAT oy = _padding - _scrollY;
    BOOL trailing = FALSE, inside = FALSE;
    DWRITE_HIT_TEST_METRICS m{};
    layout->HitTestPoint(x - ox + _scrollX, y - oy - yBase, &trailing, &inside, &m);

    const UINT32 layoutPos = m.textPosition + (trailing ? 1u : 0u);
    UINT32 pos             = 0;

    if (! isFiltered)
    {
        pos = sourceBase + layoutPos;
    }
    else if (runs && ! runs->empty())
    {
        // Find run by layout offset (runs are contiguous and sorted by layoutStart).
        auto it = std::upper_bound(runs->begin(), runs->end(), layoutPos, [](UINT32 value, const FilteredTextRun& run) { return value < run.layoutStart; });
        if (it != runs->begin())
            --it;

        const UINT32 runStart = it->layoutStart;
        const UINT32 runLen   = it->length;
        const UINT32 offset   = (layoutPos >= runStart) ? (layoutPos - runStart) : 0u;
        pos                   = it->sourceStart + std::min(offset, runLen);
    }
    else
    {
        // Filtered layout but no mapping (should be rare); fall back to best-effort base.
        pos = sourceBase;
    }

    if (pos > (UINT32)_document.TotalLength())
        pos = (UINT32)_document.TotalLength();
    return pos;
}

// ---- Find bar overlay ----
void ColorTextView::UpdateFindBarTheme()
{
    if (! _hFindPanel)
        return;

    const int dpi = static_cast<int>(std::lround(_dpi));
    auto ScalePx  = [&](int dip) { return MulDiv(dip, dpi, 96); };

    // Font: use system message font (scaled for current DPI).
    {
        // Prefer per-monitor metrics when available; otherwise scale from system DPI to the window DPI.
        using SystemParametersInfoForDpiPtr = BOOL(WINAPI*)(UINT, UINT, PVOID, UINT, UINT);
#pragma warning(push)
#pragma warning(disable : 4191) // C4191: 'reinterpret_cast': unsafe conversion from 'FARPROC'
        const auto spiForDpi = reinterpret_cast<SystemParametersInfoForDpiPtr>(GetProcAddress(GetModuleHandleW(L"user32.dll"), "SystemParametersInfoForDpi"));
#pragma warning(pop)
        const UINT systemDpi = GetDpiForSystem();

        NONCLIENTMETRICSW ncm{};
        ncm.cbSize        = sizeof(ncm);
        const bool gotNcm = spiForDpi ? spiForDpi(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0, static_cast<UINT>(dpi))
                                      : SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
        if (gotNcm)
        {
            LOGFONTW lf = ncm.lfMessageFont;
            if (! spiForDpi && systemDpi > 0)
            {
                lf.lfHeight = MulDiv(lf.lfHeight, dpi, static_cast<int>(systemDpi));
            }
            _findFont.reset(CreateFontIndirectW(&lf));
        }
        else
        {
            LOGFONTW lf{};
            wcscpy_s(lf.lfFaceName, L"Segoe UI");
            lf.lfHeight = -MulDiv(9, dpi, 72);
            _findFont.reset(CreateFontIndirectW(&lf));
        }
    }

    auto clampByte = [](float v) -> BYTE
    {
        v = std::clamp(v, 0.0f, 1.0f);
        return static_cast<BYTE>(std::lround(v * 255.0f));
    };
    auto ColorRefFromD2D = [&](const D2D1_COLOR_F& c) -> COLORREF { return RGB(clampByte(c.r), clampByte(c.g), clampByte(c.b)); };
    auto Blend           = [](const D2D1_COLOR_F& a, const D2D1_COLOR_F& b, float t) -> D2D1_COLOR_F
    {
        t = std::clamp(t, 0.0f, 1.0f);
        return D2D1::ColorF(a.r + (b.r - a.r) * t, a.g + (b.g - a.g) * t, a.b + (b.b - a.b) * t, 1.0f);
    };

    // Panel background slightly contrasted against the view background.
    const D2D1_COLOR_F panelBg = Blend(_theme.bg, _theme.fg, 0.08f);
    const D2D1_COLOR_F editBg  = Blend(_theme.bg, _theme.fg, 0.03f);
    const D2D1_COLOR_F border  = Blend(_theme.bg, _theme.fg, 0.20f);

    _findPanelBgColor = ColorRefFromD2D(panelBg);
    _findEditBgColor  = ColorRefFromD2D(editBg);
    _findTextColor    = ColorRefFromD2D(_theme.fg);

    _findPanelBgBrush.reset(CreateSolidBrush(_findPanelBgColor));
    _findEditBgBrush.reset(CreateSolidBrush(_findEditBgColor));
    _findBorderBrush.reset(CreateSolidBrush(ColorRefFromD2D(border)));

    // Cache a best-effort control height (prevents vertical clipping at larger font sizes).
    _findControlHeightPx = 0;
    if (_findFont)
    {
        wil::unique_hdc_window hdc(GetDC(_hFindPanel.get()));
        if (hdc)
        {
            const auto oldFont = wil::SelectObject(hdc.get(), _findFont.get());
            TEXTMETRICW tm{};
            if (GetTextMetricsW(hdc.get(), &tm))
            {
                _findControlHeightPx = std::max(ScalePx(22), static_cast<int>(tm.tmHeight) + ScalePx(8));
            }
        }
    }

    if (_findFont)
    {
        const WPARAM font = reinterpret_cast<WPARAM>(_findFont.get());
        SendMessageW(_hFindLabel.get(), WM_SETFONT, font, TRUE);
        SendMessageW(_hFindEdit.get(), WM_SETFONT, font, TRUE);
        SendMessageW(_hFindCase.get(), WM_SETFONT, font, TRUE);
        if (_hFindFrom)
        {
            SendMessageW(_hFindFrom.get(), WM_SETFONT, font, TRUE);
        }
    }

    if (_hFindEdit)
    {
        const int margin = ScalePx(4);
        SendMessageW(_hFindEdit.get(), EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(margin, margin));
    }

    if (_hFindFrom)
    {
        const int itemH = std::max(ScalePx(20), _findControlHeightPx > 0 ? _findControlHeightPx : 0);
        SendMessageW(_hFindFrom.get(), CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), itemH);
        SendMessageW(_hFindFrom.get(), CB_SETITEMHEIGHT, 0, itemH);
    }

    if (_hFindPanel && IsWindowVisible(_hFindPanel.get()))
    {
        LayoutFindBar();
    }

    InvalidateRect(_hFindPanel.get(), nullptr, TRUE);
}

void ColorTextView::QueueFindLiveUpdate()
{
    if (! _hWnd || ! _hFindPanel || ! _hFindEdit)
        return;
    if (! IsWindowVisible(_hFindPanel.get()))
        return;

    constexpr UINT kDelayMs = 120;
    SetTimer(_hWnd, kFindLiveTimerId, kDelayMs, nullptr);
    _findLiveTimerArmed = true;
}

void ColorTextView::PerformFindLiveUpdate()
{
    if (! _hFindEdit)
        return;

    const int len = GetWindowTextLengthW(_hFindEdit.get());
    std::wstring buffer;
    buffer.resize(static_cast<size_t>(std::max(0, len)) + 1u);
    GetWindowTextW(_hFindEdit.get(), buffer.data(), len + 1);
    buffer.resize(static_cast<size_t>(std::max(0, len)));

    const bool caseSensitive = _hFindCase && SendMessageW(_hFindCase.get(), BM_GETCHECK, 0, 0) == BST_CHECKED;
    if (buffer == _search && caseSensitive == _searchCaseSensitive)
        return;

    _search              = std::move(buffer);
    _searchCaseSensitive = caseSensitive;
    RebuildMatches();
    Invalidate();
}

void ColorTextView::UpdateFindStartModeFromUi()
{
    if (! _hFindFrom)
        return;

    const LRESULT sel = SendMessageW(_hFindFrom.get(), CB_GETCURSEL, 0, 0);
    if (sel >= 0 && sel <= 2)
    {
        _findStartMode = static_cast<FindStartMode>(sel);
    }
}

void ColorTextView::EnsureFindBar()
{
    if (_hFindPanel)
        return; // container panel

    HWND parent = GetParent(_hWnd);
    if (! parent)
        parent = _hWnd;

    const HINSTANCE instance = (HINSTANCE)GetWindowLongPtr(_hWnd, GWLP_HINSTANCE);

    _hFindPanel.reset(CreateWindowEx(0, L"STATIC", nullptr, WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, 0, 0, parent, nullptr, instance, nullptr));
    if (! _hFindPanel)
        return;

    SetWindowLongPtr(_hFindPanel.get(), GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));
    _prevFindPanelProc = reinterpret_cast<WNDPROC>(SetWindowLongPtr(_hFindPanel.get(), GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(&FindPanelProc)));

    const std::wstring findLabel   = LoadStringResource(instance, IDS_FIND_LABEL);
    const std::wstring findCase    = LoadStringResource(instance, IDS_FIND_CASE_LABEL);
    const std::wstring findCurrent = LoadStringResource(instance, IDS_FIND_FROM_CURRENT_POSITION);
    const std::wstring findTop     = LoadStringResource(instance, IDS_FIND_FROM_TOP);
    const std::wstring findBottom  = LoadStringResource(instance, IDS_FIND_FROM_BOTTOM);

    _hFindLabel.reset(CreateWindowEx(
        0, L"STATIC", findLabel.empty() ? L"Find:" : findLabel.c_str(), WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, _hFindPanel.get(), nullptr, instance, nullptr));
    _hFindEdit.reset(CreateWindowEx(
        WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL, 0, 0, 0, 0, _hFindPanel.get(), nullptr, instance, nullptr));
    _hFindCase.reset(CreateWindowEx(0,
                                    L"BUTTON",
                                    findCase.empty() ? L"Aa" : findCase.c_str(),
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX | BS_PUSHLIKE,
                                    0,
                                    0,
                                    0,
                                    0,
                                    _hFindPanel.get(),
                                    nullptr,
                                    instance,
                                    nullptr));
    _hFindFrom.reset(CreateWindowEx(0,
                                    L"COMBOBOX",
                                    nullptr,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
                                    0,
                                    0,
                                    0,
                                    0,
                                    _hFindPanel.get(),
                                    nullptr,
                                    instance,
                                    nullptr));
    if (_hFindFrom)
    {
        SendMessageW(_hFindFrom.get(), CB_ADDSTRING, 0, (LPARAM)(findCurrent.empty() ? L"Current Position" : findCurrent.c_str()));
        SendMessageW(_hFindFrom.get(), CB_ADDSTRING, 0, (LPARAM)(findTop.empty() ? L"Top" : findTop.c_str()));
        SendMessageW(_hFindFrom.get(), CB_ADDSTRING, 0, (LPARAM)(findBottom.empty() ? L"Bottom" : findBottom.c_str()));
        SendMessageW(_hFindFrom.get(), CB_SETCURSEL, static_cast<WPARAM>(_findStartMode), 0);
    }

    // Subclass edit to catch Enter/Escape
    _prevEditProc = reinterpret_cast<WNDPROC>(SetWindowLongPtr(_hFindEdit.get(), GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(&FindEditProc)));
    SetWindowLongPtr(_hFindEdit.get(), GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));

    UpdateFindBarTheme();
    HideFindBar();
}
void ColorTextView::ShowFindBar()
{
    EnsureFindBar();
    UpdateFindBarTheme();
    ShowWindow(_hFindPanel.get(), SW_SHOW);
    SetWindowPos(_hFindPanel.get(), HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
    LayoutFindBar();
    // prefill
    SetWindowTextW(_hFindEdit.get(), _search.c_str());
    SendMessageW(_hFindCase.get(), BM_SETCHECK, _searchCaseSensitive ? BST_CHECKED : BST_UNCHECKED, 0);
    if (_hFindFrom)
        SendMessageW(_hFindFrom.get(), CB_SETCURSEL, static_cast<WPARAM>(_findStartMode), 0);
    SetFocus(_hFindEdit.get());
}
void ColorTextView::HideFindBar()
{
    if (_hFindPanel)
    {
        ShowWindow(_hFindPanel.get(), SW_HIDE);
    }
    if (_findLiveTimerArmed)
    {
        KillTimer(_hWnd, kFindLiveTimerId);
        _findLiveTimerArmed = false;
    }
}
void ColorTextView::LayoutFindBar()
{
    if (! _hFindPanel)
        return; // place at top-right

    const int dpi = static_cast<int>(std::lround(_dpi));
    auto ScalePx  = [&](int dip) { return MulDiv(dip, dpi, 96); };

    HWND panelParent = GetParent(_hFindPanel.get());
    if (! panelParent)
        panelParent = _hWnd;

    RECT viewRectScreen{};
    GetWindowRect(_hWnd, &viewRectScreen);
    POINT tl{viewRectScreen.left, viewRectScreen.top};
    POINT br{viewRectScreen.right, viewRectScreen.bottom};
    ScreenToClient(panelParent, &tl);
    ScreenToClient(panelParent, &br);

    const int viewW  = std::max<int>(0, static_cast<int>(br.x - tl.x));
    const int pad    = ScalePx(8);
    const int cyPad  = ScalePx(5);
    const int ctrlH  = (_findControlHeightPx > 0) ? _findControlHeightPx : ScalePx(22);
    const int panelH = ctrlH + cyPad * 2;

    const int availableW = std::max(0, viewW - pad * 2);
    int panelW           = std::min(ScalePx(520), availableW);
    panelW               = std::max(ScalePx(120), panelW);

    const int x = tl.x + viewW - panelW - pad;
    const int y = tl.y + pad;
    MoveWindow(_hFindPanel.get(), x, y, panelW, panelH, TRUE);

    const int labelW = ScalePx(44);
    const int fromW  = ScalePx(140);
    const int caseW  = ScalePx(38);

    MoveWindow(_hFindLabel.get(), cyPad, cyPad, labelW, ctrlH, TRUE);
    MoveWindow(_hFindCase.get(), panelW - caseW - cyPad, cyPad, caseW, ctrlH, TRUE);
    if (_hFindFrom)
        MoveWindow(_hFindFrom.get(), panelW - caseW - fromW - cyPad * 2, cyPad, fromW, ctrlH, TRUE);

    const int editRightPad = cyPad + caseW + (_hFindFrom ? (fromW + cyPad) : 0);
    const int editW        = std::max(ScalePx(10), panelW - (labelW + cyPad * 3) - editRightPad);
    MoveWindow(_hFindEdit.get(), labelW + cyPad * 2, cyPad, editW, ctrlH, TRUE);
}

// ---- Message dispatch ----
LRESULT ColorTextView::WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) noexcept
{
    switch (msg)
    {
        case WM_CREATE: OnCreate(reinterpret_cast<CREATESTRUCT*>(lParam)); return 0;
        case WM_ENTERSIZEMOVE: OnEnterSizeMove(); return 0;
        case WM_EXITSIZEMOVE: OnExitSizeMove(); return 0;
        case WM_ERASEBKGND:
            // We fully paint with D2D; avoid GDI background erase to reduce flicker
            return 1;
        case WM_TIMER: return OnTimer(static_cast<UINT_PTR>(wParam));
        case WM_SIZE: OnSize(LOWORD(lParam), HIWORD(lParam)); return 0;
        case WM_PAINT: OnPaint(); return 0;
        case WM_DPICHANGED: OnDpiChanged(HIWORD(wParam), reinterpret_cast<RECT*>(lParam)); return 0;
        case WM_MOUSEWHEEL: OnMouseWheel(GET_WHEEL_DELTA_WPARAM(wParam)); return 0;
        case WM_VSCROLL: OnVScroll(LOWORD(wParam), HIWORD(wParam)); return 0;
        case WM_HSCROLL: OnHScroll(LOWORD(wParam), HIWORD(wParam)); return 0;
        case WM_LBUTTONDOWN: OnLButtonDown(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)); return 0;
        case WM_MOUSEMOVE: OnMouseMove(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), wParam); return 0;
        case WM_LBUTTONUP: OnLButtonUp(); return 0;
        case WM_SETFOCUS: OnSetFocus(); return 0;
        case WM_KILLFOCUS: OnKillFocus(); return 0;
        case WM_KEYDOWN: OnKeyDown(wParam); return 0;
        case WM_CHAR: OnChar(wParam); return 0;
        case WndMsg::kColorTextViewLayoutReady: return OnAppLayoutReady(reinterpret_cast<LayoutPacket*>(wParam));
        case WndMsg::kColorTextViewEtwBatch: return OnAppEtwBatch();
        case WndMsg::kColorTextViewWidthReady: return OnAppWidthReady(reinterpret_cast<WidthPacket*>(wParam));
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

void ColorTextView::OnEnterSizeMove()
{
    _inSizeMove = true;
}

void ColorTextView::OnExitSizeMove()
{
    _inSizeMove = false;

    EnsureBackbufferMatchesClient();
    Invalidate();
}

LRESULT ColorTextView::OnTimer(UINT_PTR timerId)
{
    if (timerId == kLayoutTimerId)
    {
        KillTimer(_hWnd, kLayoutTimerId);
        _layoutTimerArmed = false;

        const float layoutW = ComputeLayoutWidthDip();
        const UINT32 seq    = ++_layoutSeq;
        StartLayoutWorker(layoutW, seq);
        return 0;
    }

    if (timerId == kCaretTimerId)
    {
        _caretBlinkOn = ! _caretBlinkOn;
        InvalidateCaret();
        return 0;
    }

    if (timerId == kFindLiveTimerId)
    {
        KillTimer(_hWnd, kFindLiveTimerId);
        _findLiveTimerArmed = false;
        PerformFindLiveUpdate();
        return 0;
    }

    if (! _hWnd)
    {
        return 0;
    }

    return DefWindowProcW(_hWnd, WM_TIMER, static_cast<WPARAM>(timerId), 0);
}

LRESULT ColorTextView::OnAppLayoutReady(LayoutPacket* pkt)
{
    std::unique_ptr<LayoutPacket> holder(pkt);
    if (! pkt)
    {
        return 0;
    }

    if (pkt->seq != _layoutSeq.load())
    {
        return 0;
    }

    _textLayout.swap(pkt->layout);

    DWRITE_TEXT_METRICS tm{};
    if (_textLayout && SUCCEEDED(_textLayout->GetMetrics(&tm)))
    {
        const UINT32 displayRows = _document.TotalDisplayRows();
        _contentHeight           = static_cast<float>(displayRows) * GetLineHeight() + _padding * 2.f;
    }

    _sliceStartPos        = pkt->sliceStartPos;
    _sliceEndPos          = pkt->sliceEndPos;
    _sliceFirstLine       = pkt->sliceFirstLine;
    _sliceLastLine        = pkt->sliceLastLine;
    _sliceFirstDisplayRow = pkt->sliceFirstDisplayRow;
    _sliceIsFiltered      = pkt->sliceIsFiltered;
    _sliceFilteredRuns    = std::move(pkt->filteredRuns);

    _fallbackLayout.reset();
    _fallbackValid       = false;
    _fallbackStartLine   = 0;
    _fallbackEndLine     = 0;
    _fallbackLayoutWidth = 0.f;
    _fallbackFilteredRuns.clear();

    if (_textLayout)
    {
        UINT32 count = 0;
        _textLayout->GetLineMetrics(nullptr, 0, &count);
        _lineMetrics.resize(count);
        if (count != 0)
        {
            _textLayout->GetLineMetrics(_lineMetrics.data(), count, &count);
        }
    }

    ApplyColoringToLayout();
    InvalidateSliceBitmap();
    RebuildSliceBitmap();
    RebuildMatches();
    ClampScroll();
    UpdateScrollBars();

    if (_pendingScrollToBottom)
    {
        _pendingScrollToBottom = false;
        ScrollToBottom();
    }

    Invalidate();

    CachedSlice cs{};
    cs.firstLine       = _sliceFirstLine;
    cs.lastLine        = _sliceLastLine;
    cs.sliceStartPos   = _sliceStartPos;
    cs.sliceEndPos     = _sliceEndPos;
    cs.firstDisplayRow = _sliceFirstDisplayRow;
    cs.isFiltered      = _sliceIsFiltered;
    cs.filteredRuns    = _sliceFilteredRuns;
    cs.layout          = _textLayout;

    _layoutCache.erase(std::remove_if(_layoutCache.begin(),
                                      _layoutCache.end(),
                                      [&](const CachedSlice& s) { return s.firstLine == cs.firstLine && s.lastLine == cs.lastLine; }),
                       _layoutCache.end());
    _layoutCache.push_back(std::move(cs));
    if (_layoutCache.size() > kLayoutCacheMax)
    {
        _layoutCache.erase(_layoutCache.begin());
    }

    return 0;
}

LRESULT ColorTextView::OnAppEtwBatch()
{
    std::vector<ColorTextView::EtwEventEntry> batch;
    {
        auto lock = _etwQueueCS.lock();
        batch     = std::move(_etwEventQueue);
        _etwEventQueue.clear();
    }

    // Cap batch size to avoid blocking the UI thread for too long on mega-bursts.
    // Remaining events go back to the queue and we re-post to process them next pump cycle.
    constexpr size_t kMaxBatchSize = 200;
    if (batch.size() > kMaxBatchSize)
    {
        auto lock = _etwQueueCS.lock();
        // Prepend remaining events back to the front of the queue
        _etwEventQueue.insert(_etwEventQueue.begin(), std::make_move_iterator(batch.begin() + kMaxBatchSize), std::make_move_iterator(batch.end()));
        batch.resize(kMaxBatchSize);
        // Re-post to process remainder on next message pump cycle
        PostMessage(_hWnd, WndMsg::kColorTextViewEtwBatch, 0, 0);
    }

    for (const auto& entry : batch)
    {
        AppendInfoLine(entry.info, entry.message, true);
    }

    if (! batch.empty())
    {
        // Query document state once after the entire batch (instead of per-event).
        // This eliminates 3*N lock acquisitions for TotalLineCount, LongestLineChars, TotalDisplayRows.
        const size_t newLineCount = _document.TotalLineCount();
        if (_lineWidthCache.size() != newLineCount)
            _lineWidthCache.resize(newLineCount, 0.f);

        const size_t maxLen = _document.LongestLineChars();
        _approxContentWidth = GetAverageCharWidth() * static_cast<float>(maxLen);

        const UINT32 displayRows = _document.TotalDisplayRows();
        _contentHeight           = static_cast<float>(displayRows) * GetLineHeight() + _padding * 2.f;

        UpdateGutterWidth();

        if (ShouldUseAutoScrollMode())
        {
            if (_renderMode != RenderMode::AUTO_SCROLL)
            {
                SwitchToAutoScrollMode();
            }
            else
            {
                RebuildTailLayout();
                ScrollToBottom();
            }
        }
        else
        {
            if (_renderMode != RenderMode::SCROLL_BACK)
            {
                SwitchToScrollBackMode();
            }
            EnsureLayoutAdaptive(1);
            InvalidateSliceBitmap();
        }

        EnsureWidthAsync();
        Invalidate();
    }

    return 0;
}

LRESULT ColorTextView::OnAppWidthReady(WidthPacket* pkt)
{
    std::unique_ptr<ColorTextView::WidthPacket> holder(pkt);
    if (! pkt)
    {
        return 0;
    }

    if (pkt->seq != _widthSeq.load())
    {
        return 0;
    }

    if (_lineWidthCache.size() != _document.TotalLineCount())
    {
        _lineWidthCache.resize(_document.TotalLineCount(), 0.f);
    }

    for (size_t i = 0; i < pkt->indices.size() && i < pkt->widths.size(); ++i)
    {
        const size_t idx = pkt->indices[i];
        if (idx >= _lineWidthCache.size())
        {
            continue;
        }

        const float width    = pkt->widths[i];
        _lineWidthCache[idx] = width;

        if (width >= _maxMeasuredWidth)
        {
            _maxMeasuredWidth = width;
            _maxMeasuredIndex = idx;
        }
        else if (idx == _maxMeasuredIndex && width < _maxMeasuredWidth)
        {
            float localMax  = width;
            size_t localIdx = idx;
            for (size_t j = 0; j < _lineWidthCache.size(); ++j)
            {
                if (_lineWidthCache[j] > localMax)
                {
                    localMax = _lineWidthCache[j];
                    localIdx = j;
                }
            }
            _maxMeasuredWidth = localMax;
            _maxMeasuredIndex = localIdx;
        }
    }

    const float fallback = GetAverageCharWidth() * static_cast<float>(_document.LongestLineChars());
    const float widthDip = std::max(_maxMeasuredWidth, fallback);
    if (std::fabs(widthDip - _approxContentWidth) > 0.1f)
    {
        _approxContentWidth = widthDip;
        ClampHorizontalScroll();
        UpdateScrollBars();
        Invalidate();
    }

    return 0;
}

void ColorTextView::OnDpiChanged(UINT newDpi, const RECT* suggested)
{
    _dpi = static_cast<float>(newDpi);
    if (_d2dCtx)
        _d2dCtx->SetDpi(_dpi, _dpi);
    if (suggested)
    {
        SetWindowPos(_hWnd,
                     nullptr,
                     suggested->left,
                     suggested->top,
                     suggested->right - suggested->left,
                     suggested->bottom - suggested->top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }
    // Ensure swap chain matches new physical client size at this DPI
    RECT rc{};
    GetClientRect(_hWnd, &rc);
    const UINT pxW = std::max(1u, static_cast<UINT>(rc.right - rc.left));
    const UINT pxH = std::max(1u, static_cast<UINT>(rc.bottom - rc.top));
    if (_swapChain)
    {
        if (_d2dCtx)
            _d2dCtx->SetTarget(nullptr);
        _d2dTargetBitmap.reset();
        _sliceBitmap.reset();

        HRESULT hr = _swapChain->ResizeBuffers(0, pxW, pxH, DXGI_FORMAT_UNKNOWN, 0);
        if (FAILED(hr))
        {
#ifdef _DEBUG
            auto errorMsg = std::format(L"OnDpiChanged: ResizeBuffers failed: 0x{:08X}; recreating swap chain.\n", hr);
            OutputDebugStringW(errorMsg.c_str());
#endif

            if (! RecreateSwapChain(pxW, pxH))
                return;
        }
        else
        {
            CreateSwapChainResources(pxW, pxH);
        }
        _clientDipW = static_cast<float>(pxW) * 96.f / _dpi;
        _clientDipH = static_cast<float>(pxH) * 96.f / _dpi;
    }
    else
    {
        _clientDipW = static_cast<float>(pxW) * 96.f / _dpi;
        _clientDipH = static_cast<float>(pxH) * 96.f / _dpi;
    }

    _avgCharWidthValid = false;
    UpdateGutterWidth();
    // DPI affects DIP viewport size; cached layouts are not reusable.
    _layoutCache.clear();
    _fallbackLayout.reset();
    _fallbackValid = false;
    _fallbackFilteredRuns.clear();
    InvalidateSliceBitmap();

    ClampHorizontalScroll();
    UpdateScrollBars();
    ClampScroll();
    UpdateScrollBars();

    if (_renderMode == RenderMode::AUTO_SCROLL)
    {
        _tailLayoutValid = false;
        RebuildTailLayout();
        ScrollToBottom();
    }
    else
    {
        EnsureLayoutAsync();
    }
    UpdateFindBarTheme();
    LayoutFindBar();
    Invalidate();
}

void ColorTextView::OnSetFocus()
{
    _hasFocus     = true;
    _caretBlinkOn = true;
    // Start caret blink timer
    SetTimer(_hWnd, kCaretTimerId, kCaretBlinkDelayMs, nullptr);
    Invalidate();
}
void ColorTextView::OnKillFocus()
{
    _hasFocus = false;
    KillTimer(_hWnd, kCaretTimerId);
    Invalidate();
}

void ColorTextView::OnKeyDown(WPARAM vk)
{
    bool ctrl                = IsKeyDown(VK_CONTROL);
    bool shift               = IsKeyDown(VK_SHIFT);
    const UINT32 oldCaretPos = _caretPos;

    // Handle Escape to close find bar if it's visible
    if (vk == VK_ESCAPE && _hFindPanel && IsWindowVisible(_hFindPanel.get()))
    {
        HideFindBar();
        return;
    }

    switch (vk)
    {
        case 'C':
            if (ctrl && _selStart != _selEnd)
                CopySelectionToClipboard();
            break;
        case 'A':
            if (ctrl)
            {
                // Selection implies history inspection; stop hot-path auto-scroll.
                if (_renderMode == RenderMode::AUTO_SCROLL)
                {
                    SwitchToScrollBackMode();
                    EnsureLayoutAdaptive(1);
                }

                _selStart = 0;
                _selEnd   = (UINT32)_document.TotalLength();
                _caretPos = _selEnd;
                Invalidate();
            }
            break;
        case 'F':
            if (ctrl)
            {
                if (shift)
                {
                    _searchCaseSensitive = ! _searchCaseSensitive;
                }
                ShowFindBar();
            }
            break;
        case VK_F3: FindNext(shift); break;
        case VK_PRIOR: // Page Up
            if (ctrl)
                ScrollTo(0.f); // Ctrl+PageUp = top
            else
                ScrollBy(-_clientDipH * 0.9f);
            break;
        case VK_NEXT: // Page Down
            if (ctrl)
                ScrollTo(_contentHeight); // Ctrl+PageDown = bottom
            else
                ScrollBy(+_clientDipH * 0.9f);
            break;
        case VK_HOME:
            if (ctrl)
            {
                // Ctrl+Home - go to start of document
                _caretPos = 0;
                if (! shift)
                    _selStart = _selEnd = _caretPos;
                else
                    _selEnd = _caretPos;
                EnsureCaretVisible();
            }
            else
            {
                // Home - go to start of current line
                MoveCaretToLineStart(shift);
            }
            _caretBlinkOn = true;
            Invalidate();
            break;
        case VK_END:
            if (ctrl)
            {
                // Ctrl+End - go to end of document
                _caretPos = (UINT32)_document.TotalLength();
                if (! shift)
                    _selStart = _selEnd = _caretPos;
                else
                    _selEnd = _caretPos;
                EnsureCaretVisible();
            }
            else
            {
                // End - go to end of current line
                MoveCaretToLineEnd(shift);
            }
            _caretBlinkOn = true;
            Invalidate();
            break;
        case VK_LEFT:
            if (_caretPos > 0)
            {
                if (ctrl)
                    MoveCaretByWord(-1, shift); // Ctrl+Left - previous word
                else
                    _caretPos--; // Left arrow - previous character

                if (! shift)
                    _selStart = _selEnd = _caretPos;
                else
                    _selEnd = _caretPos;

                EnsureCaretVisible();
                _caretBlinkOn = true;
                Invalidate();
            }
            break;
        case VK_RIGHT:
            if (_caretPos < _document.TotalLength())
            {
                if (ctrl)
                    MoveCaretByWord(1, shift); // Ctrl+Right - next word
                else
                    _caretPos++; // Right arrow - next character

                if (! shift)
                    _selStart = _selEnd = _caretPos;
                else
                    _selEnd = _caretPos;

                EnsureCaretVisible();
                _caretBlinkOn = true;
                Invalidate();
            }
            break;
        case VK_UP:
        {
            UINT32 currentLine = GetCaretLine();
            if (currentLine > 0)
                MoveCaretToLine(currentLine - 1, shift);
        }
        break;
        case VK_DOWN:
        {
            const size_t totalLines = _document.TotalLineCount();
            if (totalLines == 0)
                break;

            const UINT32 currentLine = GetCaretLine();
            if (static_cast<size_t>(currentLine) + 1 < totalLines)
                MoveCaretToLine(currentLine + 1, shift);
        }
        break;
        default: break;
    }

    if (vk != VK_F3 && _caretPos != oldCaretPos)
        _matchIndex = -1;
}
void ColorTextView::OnChar([[maybe_unused]] WPARAM ch)
{
}

void ColorTextView::MoveCaretToLineStart(bool extendSelection)
{
    if (_document.TotalLineCount() == 0)
        return;

    UINT32 currentLine  = GetCaretLine();
    UINT32 lineStartPos = _document.GetLineStartOffset(currentLine);

    _caretPos = lineStartPos;

    if (extendSelection)
        _selEnd = _caretPos;
    else
        _selStart = _selEnd = _caretPos;

    EnsureCaretVisible();
}
void ColorTextView::MoveCaretToLineEnd(bool extendSelection)
{
    if (_document.TotalLineCount() == 0)
        return;

    UINT32 currentLine = GetCaretLine();
    if (currentLine < _document.TotalLineCount())
    {
        UINT32 lineStartPos = _document.GetLineStartOffset(currentLine);
        const auto& display = _document.GetDisplayTextRefAll(currentLine);
        _caretPos           = lineStartPos + static_cast<UINT32>(display.size());
    }

    if (extendSelection)
        _selEnd = _caretPos;
    else
        _selStart = _selEnd = _caretPos;

    EnsureCaretVisible();
}
void ColorTextView::MoveCaretByWord(int direction, bool extendSelection)
{
    if (_document.TotalLineCount() == 0)
        return;

    UINT32 newPos                  = _caretPos;
    auto [lineIndex, offsetInLine] = _document.GetLineAndOffset(newPos);
    if (lineIndex >= _document.TotalLineCount())
        lineIndex = _document.TotalLineCount() ? _document.TotalLineCount() - 1 : 0;
    const size_t totalLines = _document.TotalLineCount();

    if (direction > 0) // Move forward
    {
        const auto& display = _document.GetDisplayTextRefAll(lineIndex);
        if (offsetInLine > display.size())
            offsetInLine = static_cast<UINT32>(display.size());

        // Skip current word to its end
        while (offsetInLine < display.size() && ! iswspace(display[offsetInLine]))
            ++offsetInLine;
        // Skip whitespace after the word
        while (offsetInLine < display.size() && iswspace(display[offsetInLine]))
            ++offsetInLine;
        if (offsetInLine > display.size())
            offsetInLine = static_cast<UINT32>(display.size());

        // If at end of line, advance to the start of next line
        if (offsetInLine == display.size() && lineIndex + 1 < totalLines)
        {
            ++lineIndex;
            offsetInLine = 0;
        }
        newPos = _document.GetLineStartOffset(lineIndex) + static_cast<UINT32>(offsetInLine);
    }
    else // Move backward
    {
        if (newPos > 0)
        {
            if (offsetInLine == 0 && lineIndex > 0)
            {
                // Jump to end of previous line
                --lineIndex;
                const auto& prevDisplay = _document.GetDisplayTextRefAll(lineIndex);
                offsetInLine            = static_cast<UINT32>(prevDisplay.size());
            }

            const auto& display = _document.GetDisplayTextRefAll(lineIndex);
            if (offsetInLine > display.size())
                offsetInLine = static_cast<UINT32>(display.size());

            // Move left over whitespace
            while (offsetInLine > 0 && iswspace(display[offsetInLine - 1]))
                --offsetInLine;
            // Move left to the start of the previous word
            while (offsetInLine > 0 && ! iswspace(display[offsetInLine - 1]))
                --offsetInLine;
            newPos = _document.GetLineStartOffset(lineIndex) + static_cast<UINT32>(offsetInLine);
        }
    }

    _caretPos = newPos;

    if (extendSelection)
        _selEnd = _caretPos;
    else
        _selStart = _selEnd = _caretPos;

    EnsureCaretVisible();
}
void ColorTextView::MoveCaretToLine(UINT32 targetLine, bool extendSelection)
{
    if (_document.TotalLineCount() == 0)
        return;

    // Clamp target to document line range
    if (targetLine >= _document.TotalLineCount())
        targetLine = _document.TotalLineCount() ? static_cast<UINT32>(_document.TotalLineCount() - 1) : 0;

    // Choose a layout for hit-testing: prefer the current slice if it covers the view.
    IDWriteTextLayout* layout                = nullptr;
    bool isFiltered                          = false;
    UINT32 sourceBase                        = 0;
    const std::vector<FilteredTextRun>* runs = nullptr;
    std::optional<UINT32> localCaret         = std::nullopt;
    bool usingSlice                          = false;
    bool usingFallback                       = false;

    auto [visStartLine, visEndLine] = GetVisibleLineRange();
    const bool sliceCoversView      = _textLayout && (_sliceFirstLine <= visStartLine) && (_sliceLastLine >= visEndLine);

    if (sliceCoversView)
    {
        layout     = _textLayout.get();
        isFiltered = _sliceIsFiltered;
        sourceBase = _sliceStartPos;
        runs       = &_sliceFilteredRuns;
        usingSlice = true;
    }
    else
    {
        CreateFallbackLayoutIfNeeded(visStartLine, visEndLine);
        if (_fallbackValid && _fallbackLayout)
        {
            layout        = _fallbackLayout.get();
            isFiltered    = (_document.GetFilterMask() != Debug::InfoParam::Type::All);
            sourceBase    = _document.GetLineStartOffset(_fallbackStartLine);
            runs          = &_fallbackFilteredRuns;
            usingFallback = true;
        }
        else if (_textLayout)
        {
            layout     = _textLayout.get();
            isFiltered = _sliceIsFiltered;
            sourceBase = _sliceStartPos;
            runs       = &_sliceFilteredRuns;
            usingSlice = true;
        }
    }

    // If we have no layout yet, fall back to a simple move.
    if (! layout)
    {
        _caretPos = _document.GetLineStartOffset(targetLine);
        if (extendSelection)
            _selEnd = _caretPos;
        else
            _selStart = _selEnd = _caretPos;
        EnsureCaretVisible();
        EnsureLayoutAsync();
        Invalidate();
        return;
    }

    if (! isFiltered)
    {
        if (usingSlice && (targetLine < _sliceFirstLine || targetLine > _sliceLastLine))
        {
            _caretPos = _document.GetLineStartOffset(targetLine);
            if (extendSelection)
                _selEnd = _caretPos;
            else
                _selStart = _selEnd = _caretPos;
            EnsureCaretVisible();
            EnsureLayoutAsync();
            Invalidate();
            return;
        }

        if (usingFallback && (targetLine < _fallbackStartLine || targetLine > _fallbackEndLine))
        {
            _caretPos = _document.GetLineStartOffset(targetLine);
            if (extendSelection)
                _selEnd = _caretPos;
            else
                _selStart = _selEnd = _caretPos;
            EnsureCaretVisible();
            EnsureLayoutAsync();
            Invalidate();
            return;
        }
    }

    // Compute local caret position within the chosen layout (if possible), so we can preserve X.
    if (! isFiltered)
    {
        if (_caretPos >= sourceBase)
            localCaret = _caretPos - sourceBase;
    }
    else if (runs && ! runs->empty())
    {
        auto it = std::upper_bound(runs->begin(), runs->end(), _caretPos, [](UINT32 value, const FilteredTextRun& run) { return value < run.sourceStart; });
        if (it != runs->begin())
            --it;

        const UINT32 runStart = it->sourceStart;
        const UINT32 runEnd   = it->sourceStart + it->length;
        if (_caretPos >= runStart && _caretPos <= runEnd)
            localCaret = it->layoutStart + (_caretPos - runStart);
    }

    FLOAT currentX = 0.f;
    if (localCaret)
    {
        DWRITE_HIT_TEST_METRICS currentMetrics{};
        FLOAT currentY = 0.f;
        layout->HitTestTextPosition(*localCaret, FALSE, &currentX, &currentY, &currentMetrics);
    }

    // Compute target Y by hit-testing the start of the target line to get its layout row Y.
    std::optional<UINT32> targetLineLocalStart = std::nullopt;
    if (! isFiltered)
    {
        const UINT32 targetStart = _document.GetLineStartOffset(targetLine);
        if (targetStart >= sourceBase)
            targetLineLocalStart = targetStart - sourceBase;
    }
    else if (runs && ! runs->empty())
    {
        // Runs are in source order; binary search by sourceLine.
        auto it = std::lower_bound(runs->begin(), runs->end(), targetLine, [](const FilteredTextRun& run, UINT32 line) { return run.sourceLine < line; });
        if (it != runs->end() && it->sourceLine == targetLine)
            targetLineLocalStart = it->layoutStart;
    }

    if (! targetLineLocalStart)
    {
        // Target line not covered by this layout: jump and request a relayout.
        _caretPos = _document.GetLineStartOffset(targetLine);
        if (extendSelection)
            _selEnd = _caretPos;
        else
            _selStart = _selEnd = _caretPos;
        EnsureCaretVisible();
        EnsureLayoutAsync();
        Invalidate();
        return;
    }

    FLOAT startX = 0.f;
    FLOAT startY = 0.f;
    DWRITE_HIT_TEST_METRICS startMetrics{};
    layout->HitTestTextPosition(*targetLineLocalStart, FALSE, &startX, &startY, &startMetrics);
    const float lineH   = GetLineHeight();
    const FLOAT targetY = startY + (lineH * 0.5f);

    BOOL trailing = FALSE;
    BOOL inside   = FALSE;
    DWRITE_HIT_TEST_METRICS targetMetrics{};
    layout->HitTestPoint(currentX, targetY, &trailing, &inside, &targetMetrics);

    const UINT32 layoutPos = targetMetrics.textPosition + (trailing ? 1u : 0u);
    UINT32 newPos          = 0;
    if (! isFiltered)
    {
        newPos = sourceBase + layoutPos;
    }
    else if (runs && ! runs->empty())
    {
        auto it = std::upper_bound(runs->begin(), runs->end(), layoutPos, [](UINT32 value, const FilteredTextRun& run) { return value < run.layoutStart; });
        if (it != runs->begin())
            --it;

        const UINT32 runStart = it->layoutStart;
        const UINT32 runLen   = it->length;
        const UINT32 offset   = (layoutPos >= runStart) ? (layoutPos - runStart) : 0u;
        newPos                = it->sourceStart + std::min(offset, runLen);
    }
    else
    {
        newPos = _document.GetLineStartOffset(targetLine);
    }

    if (newPos > _document.TotalLength())
        newPos = static_cast<UINT32>(_document.TotalLength());

    _caretPos = newPos;

    if (extendSelection)
        _selEnd = _caretPos;
    else
        _selStart = _selEnd = _caretPos;

    EnsureCaretVisible();
    _caretBlinkOn = true;
    Invalidate();
}

// ---- Message dispatch ----
LRESULT CALLBACK ColorTextView::FindEditProc(HWND hEd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ColorTextView* self = (ColorTextView*)GetWindowLongPtr(hEd, GWLP_USERDATA);
    if (! self)
        return DefWindowProc(hEd, msg, wp, lp);
    switch (msg)
    {
        case WM_KEYDOWN:
            if (self->HandleFindEditKeyDown(hEd, wp))
            {
                return 0;
            }
            break;
        case WM_CHAR:
            // Let normal character input through
            break;
        case WM_KILLFOCUS:
            // Don't hide immediately on focus loss - user might click checkbox
            break;
    }

    __try
    {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
        return CallWindowProc(self->_prevEditProc, hEd, msg, wp, lp);
#pragma warning(pop)
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        return 0;
    }
}

LRESULT CALLBACK ColorTextView::FindPanelProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<ColorTextView*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    if (! self || ! self->_prevFindPanelProc)
        return DefWindowProcW(hwnd, msg, wp, lp);

    switch (msg)
    {
        case WM_COMMAND:
        {
            const UINT code = HIWORD(wp);
            const HWND ctl  = reinterpret_cast<HWND>(lp);
            if (ctl == self->_hFindEdit.get() && code == EN_CHANGE)
            {
                self->QueueFindLiveUpdate();
                return 0;
            }
            if (ctl == self->_hFindCase.get() && code == BN_CLICKED)
            {
                self->QueueFindLiveUpdate();
                return 0;
            }
            if (ctl == self->_hFindFrom.get() && code == CBN_SELCHANGE)
            {
                self->UpdateFindStartModeFromUi();
                return 0;
            }
            break;
        }
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLORBTN:
        {
            HDC hdc = reinterpret_cast<HDC>(wp);
            SetTextColor(hdc, self->_findTextColor);
            SetBkMode(hdc, TRANSPARENT);
            SetBkColor(hdc, self->_findPanelBgColor);
            return reinterpret_cast<INT_PTR>(self->_findPanelBgBrush.get());
        }
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORLISTBOX:
        {
            HDC hdc = reinterpret_cast<HDC>(wp);
            SetTextColor(hdc, self->_findTextColor);
            SetBkMode(hdc, OPAQUE);
            SetBkColor(hdc, self->_findEditBgColor);
            return reinterpret_cast<INT_PTR>(self->_findEditBgBrush.get());
        }
        case WM_ERASEBKGND:
        {
            HDC hdc = reinterpret_cast<HDC>(wp);
            RECT rc{};
            GetClientRect(hwnd, &rc);
            if (self->_findPanelBgBrush)
            {
                FillRect(hdc, &rc, self->_findPanelBgBrush.get());
                return 1;
            }
            break;
        }
        case WM_PAINT:
        {
            wil::unique_hdc_paint paintDc = wil::BeginPaint(hwnd);
            RECT rc{};
            GetClientRect(hwnd, &rc);
            if (self->_findPanelBgBrush)
            {
                FillRect(paintDc.get(), &rc, self->_findPanelBgBrush.get());
            }
            if (self->_findBorderBrush)
            {
                FrameRect(paintDc.get(), &rc, self->_findBorderBrush.get());
            }
            return 0;
        }
        case WM_NCDESTROY:
        {
            SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
            break;
        }
    }

    return SafeCallWindowProcW(self->_prevFindPanelProc, hwnd, msg, wp, lp);
}

bool ColorTextView::HandleFindEditKeyDown(HWND editControl, WPARAM key)
{
    if (key == VK_RETURN)
    {
        wchar_t buffer[512] = {};
        GetWindowTextW(editControl, buffer, 512);
        _search              = buffer;
        _searchCaseSensitive = SendMessageW(_hFindCase.get(), BM_GETCHECK, 0, 0) == BST_CHECKED;
        if (_hFindFrom)
        {
            const LRESULT sel = SendMessageW(_hFindFrom.get(), CB_GETCURSEL, 0, 0);
            if (sel >= 0 && sel <= 2)
                _findStartMode = static_cast<FindStartMode>(sel);
        }
        RebuildMatches();
        const bool initialBackward = (_findStartMode == FindStartMode::Bottom);
        FindNext(initialBackward);
        HideFindBar();
        if (_hWnd)
        {
            SetFocus(_hWnd);
        }
        return true;
    }

    if (key == VK_ESCAPE)
    {
        HideFindBar();
        if (_hWnd)
        {
            SetFocus(_hWnd);
        }
        return true;
    }

    return false;
}

void ColorTextView::SetLinePadding(float top, float bottom)
{
    _linePaddingTop    = top;
    _linePaddingBottom = bottom;
    // Apply to text format if available
    if (_textFormat)
    {
        _textFormat->SetLineSpacing(DWRITE_LINE_SPACING_METHOD_UNIFORM, _fontSize + _linePaddingTop + _linePaddingBottom, _fontSize * 0.8f + _linePaddingTop);
    }
    EnsureLayoutAsync();
    Invalidate();
}

#ifdef _DEBUG
// Debug span visualization
void ColorTextView::ClearDebugSpans()
{
    _debugSpanRects.clear();
}

void ColorTextView::DrawDebugSpans()
{
    if (! _d2dCtx || _debugSpanRects.empty())
        return;
    for (const auto& debugRect : _debugSpanRects)
    {
        ID2D1SolidColorBrush* debugBrush = GetBrush(debugRect.color);
        if (debugBrush)
            _d2dCtx->FillRectangle(debugRect.rect, debugBrush);
    }
}
#endif
