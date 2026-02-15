#pragma once
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include <algorithm>
#include <atomic>
#include <bit>
#include <cassert>
#include <cstdint>

#include <d2d1_1.h>
#include <d3d11.h>
#include <dwrite.h>
#include <dxgi1_3.h>

#include <optional>
#include <shared_mutex>
#include <string>
#include <tuple>
#include <unordered_map>
#include <utility>
#include <vector>

#include <windows.h>
#include <wrl/implements.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820) // WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027
                                                    // (move assign deleted), C4710 (not inlined), C4711 (auto inline), C4820 (padding)
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/wrl.h>
#pragma warning(pop)

#pragma comment(lib, "d2d1")
#pragma comment(lib, "dwrite")
#pragma comment(lib, "d3d11")
#pragma comment(lib, "dxgi")

#include "Document.h"
#include "Helpers.h"

#pragma warning(push)
#pragma warning(disable : 4820) // bytes padding added after data member

class ColorTextView
{
public:
    struct Theme
    {
        D2D1_COLOR_F bg              = D2D1::ColorF(D2D1::ColorF::White);
        D2D1_COLOR_F fg              = D2D1::ColorF(D2D1::ColorF::Black);
        D2D1_COLOR_F caret           = D2D1::ColorF(D2D1::ColorF::Black);
        D2D1_COLOR_F selection       = D2D1::ColorF(0.20f, 0.55f, 0.95f, 0.35f);
        D2D1_COLOR_F searchHighlight = D2D1::ColorF(1.00f, 0.85f, 0.05f, 0.35f);
        D2D1_COLOR_F gutterBg        = D2D1::ColorF(D2D1::ColorF::Gainsboro);
        D2D1_COLOR_F gutterFg        = D2D1::ColorF(D2D1::ColorF::DimGray);
        // Metadata prefix colors (by type)
        D2D1_COLOR_F metaText    = D2D1::ColorF(D2D1::ColorF::DimGray);
        D2D1_COLOR_F metaError   = D2D1::ColorF(D2D1::ColorF::Red);
        D2D1_COLOR_F metaWarning = D2D1::ColorF(D2D1::ColorF::Orange);
        D2D1_COLOR_F metaInfo    = D2D1::ColorF(D2D1::ColorF::DodgerBlue);
        D2D1_COLOR_F metaDebug   = D2D1::ColorF(D2D1::ColorF::MediumPurple);
    };

    ColorTextView();
    ~ColorTextView();

    // Explicitly delete copy constructor and copy assignment operator
    ColorTextView(const ColorTextView&)            = delete;
    ColorTextView& operator=(const ColorTextView&) = delete;

    // Creation & config
    static ATOM RegisterWndClass(HINSTANCE hinst);
    HWND Create(HWND parent, int x, int y, int w, int h);

    void SetTheme(const Theme& t);

    void EnableLineNumbers(bool enable);
    void EnableShowIds(bool enable);
    [[maybe_unused]] bool IsLineNumbersEnabled() const
    {
        return _displayLineNumbers;
    }
    [[maybe_unused]] bool IsShowIds() const
    {
        return _document.IsShowIds();
    }

    // Filter API - type-based message filtering
    void SetFilterMask(uint32_t mask);
    size_t GetVisibleLineCount() const;
    [[maybe_unused]] size_t GetTotalLineCount() const
    {
        return _document.TotalLineCount();
    }

    // Auto-scroll and rendering mode API
    void SetAutoScroll(bool enabled);
    bool GetAutoScroll() const;

    void SetLinePadding(float top, float bottom);
    void SetFont(const wchar_t* family, float sizeDips);

    [[maybe_unused]] HWND GetHWND()
    {
        return _hWnd;
    }

    // add text with or without metadata
    void AppendInfoLine(const Debug::InfoParam& info, const std::wstring& text, bool deferInvalidation = false);
    void BeginBatchAppend();
    void EndBatchAppend();
    void AppendText(const std::wstring& more);

    // ETW event batching - thread-safe queue for worker thread
    void QueueEtwEvent(const Debug::InfoParam& info, const std::wstring& message);

    // Content
    void SetText(const std::wstring& text);
    void ClearText();
    std::wstring GetText() const;
    bool SaveTextToFile(const std::wstring& path) const;

    // Editing Helpers
    void CopySelection();

    // Coloring
    void AddColorRange(UINT32 start, UINT32 length, const D2D1_COLOR_F& color);
    void ColorizeWord(const std::wstring& word, const D2D1_COLOR_F& color, bool caseSensitive = false);
    void ClearColoring();

    // Search API
    void SetSearchQuery(const std::wstring& q, bool caseSensitive = false);
    void FindNext(bool backward = false);
    void ShowFind();
    void GoToEnd(bool enableAutoScroll);

    // Request bottom-follow: immediate scroll and ensure after async layout
    void RequestScrollToBottom();

private:
    void MaybeRefreshVirtualSliceOnScroll();

    // Win32 plumbing
    static LRESULT CALLBACK WndProcThunk(HWND, UINT, WPARAM, LPARAM) noexcept;
    LRESULT WndProc(HWND, UINT, WPARAM, LPARAM) noexcept;

    struct LayoutPacket;
    struct WidthPacket;

    void OnCreate(const CREATESTRUCT* pCreateStruct);
    void OnEnterSizeMove();
    void OnExitSizeMove();
    LRESULT OnTimer(UINT_PTR timerId);
    void OnSize(UINT w, UINT h);
    void OnPaint();
    void OnDpiChanged(UINT newDpi, const RECT* suggested);
    void OnMouseWheel(short delta);
    void OnLButtonDown(int x, int y);
    void OnMouseMove(int x, int y, WPARAM mods);
    void OnLButtonUp();
    void OnSetFocus();
    void OnKillFocus();
    void OnKeyDown(WPARAM vk);
    void OnChar(WPARAM ch);
    void OnVScroll(UINT code, UINT pos);
    void OnHScroll(UINT code, UINT pos);
    LRESULT OnAppLayoutReady(LayoutPacket* pkt);
    LRESULT OnAppEtwBatch();
    LRESULT OnAppWidthReady(WidthPacket* pkt);

    // Scrolling
    void ScrollBy(float dy);
    void ScrollTo(float y);
    void ScrollToBottom();
    // Manage scrollers
    void UpdateScrollBars();
    float GetLineHeight() const;
    float GetAverageCharWidth() const;
    void ClampHorizontalScroll();
    UINT32 GetCaretLine() const;
    void MoveCaretToLine(UINT32 targetLine, bool extendSelection);
    void EnsureCaretVisible();
    void MoveCaretToLineStart(bool extendSelection);
    void MoveCaretToLineEnd(bool extendSelection);
    void MoveCaretByWord(int direction, bool extendSelection);

    // D2D/DWrite + D3D/DXGI
    void CreateDpiAwareDefaults();
    void CreateDeviceIndependentResources();
    void CreateDeviceResources();
    void CreateSwapChainResources(UINT w, UINT h);
    bool RecreateSwapChain(UINT w, UINT h);
    void DiscardDeviceResources();
    void EnsureBackbufferMatchesClient();
    void RebuildSliceBitmap();
    void InvalidateSliceBitmap();
    void EnsureLayoutAsync();
    void EnsureLayoutAdaptive(size_t changeSize);
    void CreateFallbackLayoutIfNeeded(size_t visStartLine, size_t visEndLine);
    void ApplyColoringToLayout();

    // Two-mode rendering Helpers
    void SwitchToAutoScrollMode();
    void SwitchToScrollBackMode();
    void RebuildTailLayout();
    void ApplyColoringToTailLayout();
    bool ShouldUseAutoScrollMode() const;
    void StartLayoutWorker(float layoutWidth, UINT32 seq);

    // Hit-testing & drawing
    UINT32 PosFromPoint(float x, float y);
    void DrawHighlights();
    void DrawSelection();
    void DrawLineNumbers();
    void DrawCaret();

    // Utilities
    void Invalidate();
    void ClampScroll();
    void CopySelectionToClipboard();
    void RebuildMatches();
    bool ValidateDeviceState() const;
    void LogSystemInfo() const;
    std::pair<size_t, size_t> GetVisibleLineRange() const;
    std::pair<UINT32, UINT32> GetVisibleTextRange() const;
    ID2D1SolidColorBrush* GetBrush(const D2D1_COLOR_F& color);
    void PruneBrushCacheIfNeeded();
    void UpdateGutterWidth();
    void InvalidateCaret();
    RECT GetCaretRectPx() const;
    void InvalidateExposedArea(float deltaDipY);
    void EnsureWidthAsync();

    // Filter Helpers
    void RebuildVisibleIndex() const;
    void InvalidateVisibleIndex();

    // Find bar UI (overlay)
    void EnsureFindBar();
    void ShowFindBar();
    void HideFindBar();
    void LayoutFindBar();
    void UpdateFindBarTheme();
    void QueueFindLiveUpdate();
    void PerformFindLiveUpdate();
    void UpdateFindStartModeFromUi();
    bool HandleFindEditKeyDown(HWND editControl, WPARAM key);
    static LRESULT CALLBACK FindEditProc(HWND, UINT, WPARAM, LPARAM) noexcept;
    static LRESULT CALLBACK FindPanelProc(HWND, UINT, WPARAM, LPARAM) noexcept;

    // Window handle
    HWND _hWnd = nullptr;

    // DPI
    float _dpi        = 96.0f;
    float _clientDipW = 0.0f;
    float _clientDipH = 0.0f;

    // Content
    Document _document;

    // Factories & resources
    wil::com_ptr<ID2D1Factory1> _d2d1Factory;
    wil::com_ptr<IDWriteFactory> _dwriteFactory;
    // D3D/DXGI device stack
    wil::com_ptr<ID3D11Device> _d3dDevice;
    wil::com_ptr<ID3D11DeviceContext> _d3dContext;
    wil::com_ptr<IDXGISwapChain1> _swapChain;
    // D2D device and context (we keep name _rt for minimal changes)
    wil::com_ptr<ID2D1Device> _d2dDevice;
    wil::com_ptr<ID2D1DeviceContext> _d2dCtx;
    wil::com_ptr<ID2D1Bitmap1> _d2dTargetBitmap; // backbuffer as D2D target
    // Offscreen slice cache bitmap
    wil::com_ptr<ID2D1Bitmap1> _sliceBitmap;
    float _sliceBitmapYBase = 0.f;
    float _sliceDipW        = 0.f;
    float _sliceDipH        = 0.f;
    // Optimization #1 - Incremental bitmap update tracking
    struct SliceDirtyRegion
    {
        size_t firstDirtyLine = 0;
        size_t lastDirtyLine  = 0;
        bool valid            = false;
    };
    SliceDirtyRegion _sliceDirtyRegion;
    // Offscreen slice cache (command list in DIPs to avoid DPI issues)
    wil::com_ptr<ID2D1CommandList> _sliceCmd;
    wil::com_ptr<IDWriteTextLayout> _fallbackLayout;
    size_t _fallbackStartLine  = 0;
    size_t _fallbackEndLine    = 0;
    float _fallbackLayoutWidth = 0.f;
    bool _fallbackValid        = false;
    // Brush cache: hashmap with timestamp-based pruning
    using BrushCacheKey = std::tuple<float, float, float, float>;
    struct BrushCacheEntry
    {
        wil::com_ptr<ID2D1SolidColorBrush> brush;
        uint64_t lastAccess = 0;
    };
    struct BrushKeyHash
    {
        [[maybe_unused]] size_t operator()(const BrushCacheKey& key) const noexcept
        {
            auto [r, g, b, a] = key;
            const auto hr     = std::bit_cast<uint32_t>(r);
            const auto hg     = std::bit_cast<uint32_t>(g);
            const auto hb     = std::bit_cast<uint32_t>(b);
            const auto ha     = std::bit_cast<uint32_t>(a);
            size_t h          = 1469598103934665603ull;
            h ^= hr + 0x9e3779b97f4a7c15ull + (h << 6) + (h >> 2);
            h ^= hg + 0x9e3779b97f4a7c15ull + (h << 6) + (h >> 2);
            h ^= hb + 0x9e3779b97f4a7c15ull + (h << 6) + (h >> 2);
            h ^= ha + 0x9e3779b97f4a7c15ull + (h << 6) + (h >> 2);
            return h;
        }
    };
    std::unordered_map<BrushCacheKey, BrushCacheEntry, BrushKeyHash> _brushCache;
    uint64_t _brushAccessCounter = 0;

    wil::com_ptr<IDWriteTextFormat> _textFormat;
    wil::com_ptr<IDWriteTextFormat> _gutterTextFormat;
    wil::com_ptr<IDWriteTextLayout> _textLayout;
    std::vector<DWRITE_LINE_METRICS> _lineMetrics;
    // Map logical document lines to display line indices (start) for precise gutter placement
    std::vector<UINT32> _docLineToDisplayStart;

    // Layout / rendering params
    float _fontSize          = 12.0f; // DIPs
    std::wstring _fontFamily = L"Segoe UI";

    // Layout/scroll
    float _padding            = 10.f;
    float _linePaddingTop     = 0.0f; // extra padding per line
    float _linePaddingBottom  = 0.0f;
    float _scrollY            = 0.f;
    float _scrollX            = 0.f; // Horizontal scroll position
    float _contentHeight      = 0.f;
    float _approxContentWidth = 0.f; // for horizontal scrollbar
    int _wheelRemainder       = 0;
    int _wheelRemainderX      = 0;
    bool _displayLineNumbers  = true;
    float _gutterDipW         = 50.f;
    UINT32 _gutterDigits      = 2; // cached digit count for gutter width
    // Scrollbar visibility/state
    bool _horzScrollbarVisible = false; // updated in UpdateScrollBars

    // Selection/caret
    UINT32 _selStart                         = 0;
    UINT32 _selEnd                           = 0;
    UINT32 _caretPos                         = 0;
    bool _mouseDown                          = false;
    bool _hasFocus                           = false;
    bool _caretBlinkOn                       = true;
    static constexpr UINT_PTR kCaretTimerId  = 2;
    static constexpr UINT kCaretBlinkDelayMs = 530;

    // Search state
    std::wstring _search;
    bool _searchCaseSensitive = false;
    std::vector<Line::ColorSpan> _matches;
    __int64 _matchIndex = -1;
    enum class FindStartMode : uint8_t
    {
        CurrentPosition = 0,
        Top             = 1,
        Bottom          = 2,
    };
    FindStartMode _findStartMode = FindStartMode::CurrentPosition;

    // Async
    std::atomic<UINT32> _layoutSeq{0};
    std::atomic<UINT32> _widthSeq{0};
    bool _layoutTimerArmed                    = false;
    static constexpr UINT_PTR kLayoutTimerId  = 1;
    static constexpr UINT kLayoutTimerDelayMs = 16; // Default timer delay (adaptive timing may use faster)
    size_t _sliceFirstLine                    = 0;
    size_t _sliceLastLine                     = 0;
    UINT32 _sliceFirstDisplayRow              = 0;     // Display row offset for filtered rendering
    bool _sliceIsFiltered                     = false; // True if slice text contains only visible lines (not all lines) // inclusive
    UINT32 _sliceStartPos                     = 0;
    UINT32 _sliceEndPos                       = 0; // exclusive

    // Filtered-mode mapping: layout offsets are not contiguous in the source document.
    // These runs map layout ranges back to source character positions (unfiltered document space).
    struct FilteredTextRun
    {
        size_t sourceLine  = 0;
        UINT32 layoutStart = 0;
        UINT32 length      = 0; // identical length in layout and source spaces
        UINT32 sourceStart = 0;
    };
    std::vector<FilteredTextRun> _sliceFilteredRuns;
    std::vector<FilteredTextRun> _fallbackFilteredRuns;

    // Cached metrics
    mutable bool _avgCharWidthValid = false;
    mutable float _avgCharWidth     = 0.f;
    std::vector<float> _lineWidthCache;
    float _maxMeasuredWidth  = 0.f;
    size_t _maxMeasuredIndex = 0;

    // Theme
    Theme _theme{};

    // Slice cache: reuse recent virtual layouts for smooth scroll
    struct CachedSlice
    {
        size_t firstLine       = 0;
        size_t lastLine        = 0; // inclusive
        UINT32 sliceStartPos   = 0;
        UINT32 sliceEndPos     = 0; // exclusive
        UINT32 firstDisplayRow = 0; // Display row offset for filtered rendering
        bool isFiltered        = false;
        std::vector<FilteredTextRun> filteredRuns;
        wil::com_ptr<IDWriteTextLayout> layout;
    };
    std::vector<CachedSlice> _layoutCache; // small LRU
    static constexpr size_t kSliceBlockLines = 256;
    static constexpr size_t kLayoutCacheMax  = 8;
    // Prefetch margin (logical lines) beyond the visible slice
    static constexpr size_t kSlicePrefetchMargin = 128;
    // Layout and slice sizing
    static constexpr float kMinLayoutWidthDip          = 256.f;
    static constexpr float kMaxLayoutWidthDip          = 8192.f;
    static constexpr float kLayoutWidthSafetyMarginDip = 512.f;
    static constexpr float kSliceBitmapMaxWidthDip     = 4096.f;
    static constexpr UINT kMaxD3D11TextureDimension    = 16384u; // D3D11 Feature Level 11.x limit

    // Two-mode rendering architecture
    // _renderMode is the single source of truth for auto-scroll state
    enum class RenderMode : uint8_t
    {
        AUTO_SCROLL, // Hot path: simple tail layout, direct rendering (auto-scroll ON)
        SCROLL_BACK  // Cold path: full virtualization, slice caching (auto-scroll OFF)
    };
    RenderMode _renderMode = RenderMode::AUTO_SCROLL;

    // AUTO-SCROLL mode state (hot path)
    wil::com_ptr<IDWriteTextLayout> _tailLayout; // Layout for last N lines
    size_t _tailFirstLine              = 0;      // First line in tail layout
    static constexpr size_t kTailLines = 100;    // Fixed window size for tail
    bool _tailLayoutValid              = false;

    // Adaptive layout timing thresholds
    static constexpr size_t kSyncLayoutThresholdLines = 100; // Lines below which we use sync layout
    static constexpr UINT kFastLayoutTimerDelayMs     = 4;   // Fast timer for small changes

    // structure for WndMsg::kColorTextViewLayoutReady
    struct LayoutPacket
    {
        UINT32 seq{};
        UINT32 sliceStartPos{};
        UINT32 sliceEndPos{};
        size_t sliceFirstLine{};
        size_t sliceLastLine{};
        UINT32 sliceFirstDisplayRow{}; // Display row offset for filtered rendering
        bool sliceIsFiltered{};        // True if text contains only visible lines
        std::vector<FilteredTextRun> filteredRuns;
        wil::com_ptr<IDWriteTextLayout> layout;
    };
    struct WidthPacket
    {
        UINT32 seq{};
        std::vector<size_t> indices;
        std::vector<float> widths;
    };

    // ETW event queue entry - accumulated from worker thread, batch-processed on UI thread
    struct EtwEventEntry
    {
        Debug::InfoParam info{};
        std::wstring message;
    };

    // ETW event queue with wil::critical_section for efficient batching (prevents message loop flooding)
    std::vector<EtwEventEntry> _etwEventQueue;
    wil::critical_section _etwQueueCS;

    // Find bar UI
    wil::unique_hwnd _hFindPanel;
    wil::unique_hwnd _hFindLabel;
    wil::unique_hwnd _hFindEdit;
    wil::unique_hwnd _hFindCase;
    wil::unique_hwnd _hFindFrom;
    WNDPROC _prevEditProc      = nullptr;
    WNDPROC _prevFindPanelProc = nullptr;
    wil::unique_hfont _findFont;
    wil::unique_hbrush _findPanelBgBrush;
    wil::unique_hbrush _findEditBgBrush;
    wil::unique_hbrush _findBorderBrush;
    COLORREF _findPanelBgColor                 = RGB(255, 255, 255);
    COLORREF _findEditBgColor                  = RGB(255, 255, 255);
    COLORREF _findTextColor                    = RGB(0, 0, 0);
    int _findControlHeightPx                   = 0;
    bool _findLiveTimerArmed                   = false;
    static constexpr UINT_PTR kFindLiveTimerId = 3;

    // Live-resize state: when true, avoid expensive resizes and vsync waits
    bool _inSizeMove = false;

    // Partial present state
    bool _needsFullRedraw    = true;
    bool _presentInitialized = false;
    bool _hasPendingDirty    = false;
    bool _hasPendingScroll   = false;
    LONG _pendingScrollDy    = 0;
    RECT _pendingDirtyRect{};

    // Layout optimization
    bool _pendingScrollToBottom = false; // ensure bottom after next layout

    // Helpers
    RECT GetViewRect() const;
    D2D1_RECT_F GetViewRectDip() const;
    float GetTextViewportWidthDip() const;
    float ComputeLayoutWidthDip() const;

    void DrawScene(bool clearTarget);
    void RequestFullRedraw();
    void ResetPresentationState();

    void ClearTextLayoutEffects();

#ifdef _DEBUG
    // Debug span visualization (always declared, only used in debug)
    struct DebugSpanRect
    {
        D2D1_RECT_F rect;
        D2D1_COLOR_F color;
    };
    wil::com_ptr<ID2D1SolidColorBrush> _debugDirtyRectBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _debugDirtyRectFillBrush;
    size_t _debugDirtyColorIndex = 0;
    std::vector<DebugSpanRect> _debugSpanRects;
    void DrawDebugSpans();
    void ClearDebugSpans();
#endif
};

#pragma warning(pop)
