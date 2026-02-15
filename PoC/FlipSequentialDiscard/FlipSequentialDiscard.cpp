// SwapChainPartialInvalidation.cpp
// Minimal demo: toggle between FLIP_SEQUENTIAL (dirty rects) and FLIP_DISCARD (full redraw).
// - Normal run: sequential (partial invalidation enabled)
// - Run with "--discard": discard (partial invalidation disabled)
// - Press 'D' during runtime to toggle (swap chain is recreated)

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include <algorithm>
#include <array>
#include <chrono>
#include <cmath>
#include <cstdint>

#include <d2d1_3.h>
#include <d3d11.h>
#include <dwrite.h>
#include <dxgi1_6.h>

#include <format>
#include <iterator>
#include <limits>
#include <random>
#include <string>
#include <string_view>
#include <vector>

#include <windows.h>
#include <windowsx.h>

#pragma warning(push)
// WRL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wrl.h>
#pragma warning(pop)

#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "dxgi.lib")
#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dwrite.lib")
#pragma comment(lib, "windowscodecs.lib")

using Microsoft::WRL::ComPtr;

static const wchar_t* kClassName   = L"SwapChainPartialInvalidationDemo";
static const wchar_t* kWindowTitle = L"Partial Swap Chain Invalidation Demo (Press D to toggle)";

enum class SwapMode : uint8_t
{
    Sequential,
    Discard
};

struct App
{
    // Win32
    HWND _hwnd = nullptr;
    RECT _client{};

    // DX11 / DXGI
    ComPtr<ID3D11Device> _d3dDevice;
    ComPtr<ID3D11DeviceContext> _d3dCtx;
    ComPtr<IDXGISwapChain1> _swapChain;
    DXGI_SWAP_EFFECT _swapEffect = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;

    // Direct2D / DWrite (on top of the DXGI device)
    ComPtr<ID2D1Factory1> _d2dFactory;
    ComPtr<ID2D1Device> _d2dDevice;
    ComPtr<ID2D1DeviceContext> _d2dCtx;
    ComPtr<ID2D1Bitmap1> _d2dTarget;
    ComPtr<IDWriteFactory> _dwFactory;
    ComPtr<IDWriteTextFormat> _textFormat;
    ComPtr<ID2D1SolidColorBrush> _brushText;
    ComPtr<ID2D1SolidColorBrush> _brushDirtyOutline; // added outline brush
    std::array<ComPtr<ID2D1SolidColorBrush>, 256> _brushBgPalette{};
    size_t _bgBrushIndex       = 0; // next (pending) background brush index
    size_t _activeBgBrushIndex = 0; // brush index currently visible on screen (last committed)

    // State
    SwapMode _mode         = SwapMode::Sequential;
    float _dpiScale        = 1.0f;
    const float _marginDip = 16.0f;

    std::vector<std::wstring> _lines;
    uint64_t _lineCounter = 0;
    std::mt19937_64 _randomGen{std::random_device{}()};
    std::uniform_int_distribution<int> _logLevelDist{0, 2};
    std::uniform_int_distribution<int> _lineLengthDist{48, 120};
    std::uniform_int_distribution<int> _hexDigitDist{0, 22};
    std::chrono::steady_clock::time_point _lastAppendTime = std::chrono::steady_clock::now();
    std::chrono::milliseconds _appendInterval{10};

    float _lineHeight      = 0.0f;
    float _baseline        = 0.0f;
    float _scrollOffset    = 0.0f;
    bool _autoScrollToTail = true;
    bool _needsFullRedraw  = true;
    RECT _pendingDirtyRect{};
    bool _hasPendingDirty         = false;
    bool _hasPendingScroll        = false;
    LONG _pendingScrollDy         = 0;
    size_t _lastRenderedLineCount = 0;
    UINT _lastAppendedCount       = 0;     // track number of lines appended this tick for partial drawing
    bool _presentInitialized      = false; // have we presented at least once (full) for current swap chain?
    int _scrollMaxCached          = std::numeric_limits<int>::min();
    UINT _scrollPageCached        = 0;
    int _scrollPosCached          = std::numeric_limits<int>::min();

    void ToggleMode()
    {
        _mode             = (_mode == SwapMode::Sequential) ? SwapMode::Discard : SwapMode::Sequential;
        _swapEffect       = (_mode == SwapMode::Sequential) ? DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL : DXGI_SWAP_EFFECT_FLIP_DISCARD;
        _needsFullRedraw  = true;
        _hasPendingDirty  = false;
        _hasPendingScroll = false;
        _pendingDirtyRect = RECT{};
        _pendingScrollDy  = 0;
        RecreateSwapChainAndTargets();
        SetTitle();
    }

    void ResetLogState()
    {
        _lines.clear();
        _lineCounter           = 0;
        _scrollOffset          = 0.0f;
        _autoScrollToTail      = true;
        _needsFullRedraw       = true;
        _pendingDirtyRect      = RECT{};
        _hasPendingDirty       = false;
        _hasPendingScroll      = false;
        _pendingScrollDy       = 0;
        _lastRenderedLineCount = 0;
        _lastAppendTime        = std::chrono::steady_clock::now();
        _bgBrushIndex          = 0;
        _activeBgBrushIndex    = 0;
        _presentInitialized    = false;
        _scrollMaxCached       = std::numeric_limits<int>::min();
        _scrollPageCached      = 0;
        _scrollPosCached       = std::numeric_limits<int>::min();
    }

    float Margin() const
    {
        return _marginDip * _dpiScale;
    }

    float ClientWidth() const
    {
        return static_cast<float>(_client.right - _client.left);
    }

    float ClientHeight() const
    {
        return static_cast<float>(_client.bottom - _client.top);
    }

    D2D1_RECT_F TextAreaRectF() const
    {
        const float m = Margin();
        return D2D1::RectF(m, m, ClientWidth() - m, ClientHeight() - m);
    }

    RECT TextAreaRect() const
    {
        const D2D1_RECT_F area = TextAreaRectF();
        RECT r{};
        r.left   = static_cast<LONG>(std::floor(area.left));
        r.top    = static_cast<LONG>(std::floor(area.top));
        r.right  = static_cast<LONG>(std::ceil(area.right));
        r.bottom = static_cast<LONG>(std::ceil(area.bottom));
        return r;
    }

    float ContentHeight() const
    {
        return static_cast<float>(_lines.size()) * _lineHeight;
    }

    void SnapLineMetricsToPixels()
    {
        if (_lineHeight > 0.0f)
        {
            _lineHeight = std::max(1.0f, std::round(_lineHeight));
        }

        if (_baseline > 0.0f)
        {
            _baseline = std::max(0.0f, std::round(_baseline));
        }
    }

    static D2D1_COLOR_F MakePaletteColor(size_t index)
    {
        static const std::array<D2D1_COLOR_F, 41> palette{D2D1::ColorF(D2D1::ColorF::OrangeRed),     D2D1::ColorF(D2D1::ColorF::MediumBlue),
                                                          D2D1::ColorF(D2D1::ColorF::LimeGreen),     D2D1::ColorF(D2D1::ColorF::DarkMagenta),
                                                          D2D1::ColorF(D2D1::ColorF::Gold),          D2D1::ColorF(D2D1::ColorF::DeepSkyBlue),
                                                          D2D1::ColorF(D2D1::ColorF::Crimson),       D2D1::ColorF(D2D1::ColorF::DarkTurquoise),
                                                          D2D1::ColorF(D2D1::ColorF::YellowGreen),   D2D1::ColorF(D2D1::ColorF::BlueViolet),
                                                          D2D1::ColorF(D2D1::ColorF::DarkOrange),    D2D1::ColorF(D2D1::ColorF::DodgerBlue),
                                                          D2D1::ColorF(D2D1::ColorF::Chartreuse),    D2D1::ColorF(D2D1::ColorF::MediumVioletRed),
                                                          D2D1::ColorF(D2D1::ColorF::Khaki),         D2D1::ColorF(D2D1::ColorF::SteelBlue),
                                                          D2D1::ColorF(D2D1::ColorF::Tomato),        D2D1::ColorF(D2D1::ColorF::Turquoise),
                                                          D2D1::ColorF(D2D1::ColorF::LawnGreen),     D2D1::ColorF(D2D1::ColorF::Indigo),
                                                          D2D1::ColorF(D2D1::ColorF::Orange),        D2D1::ColorF(D2D1::ColorF::RoyalBlue),
                                                          D2D1::ColorF(D2D1::ColorF::SpringGreen),   D2D1::ColorF(D2D1::ColorF::HotPink),
                                                          D2D1::ColorF(D2D1::ColorF::LightSeaGreen), D2D1::ColorF(D2D1::ColorF::SandyBrown),
                                                          D2D1::ColorF(D2D1::ColorF::SlateBlue),     D2D1::ColorF(D2D1::ColorF::MediumSpringGreen),
                                                          D2D1::ColorF(D2D1::ColorF::Firebrick),     D2D1::ColorF(D2D1::ColorF::CornflowerBlue),
                                                          D2D1::ColorF(D2D1::ColorF::Yellow),        D2D1::ColorF(D2D1::ColorF::MediumPurple),
                                                          D2D1::ColorF(D2D1::ColorF::Aquamarine),    D2D1::ColorF(D2D1::ColorF::DarkRed),
                                                          D2D1::ColorF(D2D1::ColorF::SeaGreen),      D2D1::ColorF(D2D1::ColorF::DeepPink),
                                                          D2D1::ColorF(D2D1::ColorF::DarkBlue),      D2D1::ColorF(D2D1::ColorF::PaleVioletRed),
                                                          D2D1::ColorF(D2D1::ColorF::DarkSeaGreen),  D2D1::ColorF(D2D1::ColorF::Magenta),
                                                          D2D1::ColorF(D2D1::ColorF::DarkGoldenrod)};

        return palette[index % palette.size()];
    }

    void CreateBrushPalette()
    {
        for (auto& brush : _brushBgPalette)
        {
            brush.Reset();
        }
        _bgBrushIndex = 0;

        for (size_t i = 0; i < _brushBgPalette.size(); ++i)
        {
            const auto color = MakePaletteColor(i);
            const HRESULT hr = _d2dCtx->CreateSolidColorBrush(color, _brushBgPalette[i].GetAddressOf());
            if (FAILED(hr))
            {
                _brushBgPalette[i].Reset();
            }
        }

        if (! _brushBgPalette[0])
        {
            const auto fallback = D2D1::ColorF(0.1f, 0.1f, 0.1f, 1.0f);
            _d2dCtx->CreateSolidColorBrush(fallback, _brushBgPalette[0].GetAddressOf());
        }
    }

    ID2D1SolidColorBrush* AcquireBackgroundBrush(size_t& indexOut)
    {
        return _brushBgPalette[_activeBgBrushIndex + 1].Get();

        // for (size_t attempt = 0; attempt < _brushBgPalette.size(); ++attempt)
        //{
        //     const size_t candidate = (_bgBrushIndex + attempt) % _brushBgPalette.size();
        //     if (_brushBgPalette[candidate])
        //     {
        //         indexOut = candidate;
        //         return _brushBgPalette[candidate].Get();
        //     }
        // }
        // indexOut = _activeBgBrushIndex; // fallback to currently active
        // return _brushText.Get();
    }

    void CommitBackgroundBrush(size_t indexUsed)
    {
        _activeBgBrushIndex = indexUsed;
        _bgBrushIndex       = (indexUsed + 1) % _brushBgPalette.size();
    }

    void UpdateScrollBar(float viewHeight)
    {
        const int maxVal = static_cast<int>(std::round(ContentHeight()));
        const UINT page  = static_cast<UINT>(std::max(viewHeight, _lineHeight));
        const int pos    = static_cast<int>(std::round(_scrollOffset));

        if (maxVal == _scrollMaxCached && page == _scrollPageCached && pos == _scrollPosCached)
        {
            return;
        }

        SCROLLINFO si{};
        si.cbSize = sizeof(si);
        si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
        si.nMin   = 0;
        si.nMax   = maxVal;
        si.nPage  = page;
        si.nPos   = pos;
        SetScrollInfo(_hwnd, SB_VERT, &si, TRUE);

        _scrollMaxCached  = maxVal;
        _scrollPageCached = page;
        _scrollPosCached  = pos;
    }

    void WarmUpLogs(size_t count)
    {
        _lines.reserve(_lines.size() + count);
        for (size_t i = 0; i < count; ++i)
        {
            _lines.emplace_back(MakeRandomLine());
        }
        _lastAppendTime  = std::chrono::steady_clock::now();
        _bgBrushIndex    = 0;
        _needsFullRedraw = true;
    }

    // Set append speed based on digit key (1..9, 0). 1 => 1000ms, 9 => ~111ms, 0 => no delay (one line per Tick)
    void SetAppendSpeedDigit(int digit)
    {
        if (digit == 0)
        {
            _appendInterval = std::chrono::milliseconds{0};
        }
        else
        {
            const int ms    = std::max(1, 1000 / digit);
            _appendInterval = std::chrono::milliseconds{ms};
        }
        _lastAppendTime = std::chrono::steady_clock::now();
    }

    std::wstring MakeRandomLine()
    {
        static constexpr std::array<std::wstring_view, 3> levels{L"INFO", L"WARN", L"ERROR"};
        static constexpr std::wstring_view digits = L"0123456789ABCDEF🥶😍😊にまจ็ال";

        const int levelIndex = _logLevelDist(_randomGen);
        const int len        = _lineLengthDist(_randomGen);

        std::wstring line;
        line.reserve(static_cast<size_t>(len) + 24);

        std::format_to(std::back_inserter(line), L"[{:06}] {:>5} ", ++_lineCounter, levels[static_cast<size_t>(levelIndex)]);

        const size_t payloadOffset = line.size();
        if (len > 0)
        {
            line.resize(payloadOffset + static_cast<size_t>(len));
            wchar_t* payload = line.data() + payloadOffset;
            for (int i = 0; i < len; ++i)
            {
                payload[i] = digits[static_cast<size_t>(_hexDigitDist(_randomGen))];
            }
        }
        else
        {
            line += L"Stupid Empty String";
        }

        return line;
    }

    UINT MaybeAppendLines(std::chrono::steady_clock::time_point now)
    {
        // Zero interval: one line per tick (avoid tight loop)
        if (_appendInterval.count() == 0)
        {
            _lines.emplace_back(MakeRandomLine());
            _lastAppendTime = now;
            return 1;
        }

        const auto elapsed = now - _lastAppendTime;
        if (elapsed < _appendInterval)
        {
            return 0;
        }

        const auto intervals = elapsed / _appendInterval;
        const auto clamped   = std::min(intervals, static_cast<std::chrono::steady_clock::duration::rep>(4));
        const UINT count     = static_cast<UINT>(clamped);
        if (count == 0)
        {
            return 0;
        }

        const size_t newSize = _lines.size() + static_cast<size_t>(count);
        if (_lines.capacity() < newSize)
        {
            const size_t growTo = std::max(newSize, std::max(_lines.capacity() * 2, static_cast<size_t>(64)));
            _lines.reserve(growTo);
        }

        for (UINT i = 0; i < count; ++i)
        {
            _lines.emplace_back(MakeRandomLine());
        }

        _lastAppendTime += _appendInterval * static_cast<std::chrono::steady_clock::duration::rep>(count);

        if (count == 4)
        {
            _lastAppendTime = now;
        }

        return count;
    }

    void DrawVisibleLines(const D2D1_RECT_F& textArea)
    {
        if (! _textFormat || _lineHeight <= 0.0f || _lines.empty())
        {
            return;
        }

        _d2dCtx->PushAxisAlignedClip(textArea, D2D1_ANTIALIAS_MODE_ALIASED);

        const float startIndexF = _scrollOffset / _lineHeight;
        size_t startIndex       = startIndexF > 0.0f ? static_cast<size_t>(startIndexF) : 0;
        if (startIndex >= _lines.size())
        {
            startIndex = _lines.size();
        }
        float y = textArea.top - (_scrollOffset - static_cast<float>(startIndex) * _lineHeight);
        y       = std::round(y);

        for (size_t i = startIndex; i < _lines.size() && y < textArea.bottom; ++i)
        {
            const float lineTop        = std::round(y);
            const float lineBottom     = lineTop + _lineHeight;
            const std::wstring& line   = _lines[i];
            const D2D1_RECT_F lineRect = D2D1::RectF(textArea.left, lineTop, textArea.right, lineBottom);
            _d2dCtx->DrawTextW(line.c_str(),
                               static_cast<UINT32>(line.size()),
                               _textFormat.Get(),
                               lineRect,
                               _brushText.Get(),
                               D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT,
                               DWRITE_MEASURING_MODE_NATURAL);
            y = lineBottom;
        }

        _d2dCtx->PopAxisAlignedClip();
    }

    bool RenderFull(const D2D1_RECT_F& textArea)
    {
        size_t brushIndex             = 0;
        ID2D1SolidColorBrush* bgBrush = AcquireBackgroundBrush(brushIndex);

        _d2dCtx->BeginDraw();
        const D2D1_RECT_F canvas = D2D1::RectF(0.0f, 0.0f, ClientWidth(), ClientHeight());
        _d2dCtx->FillRectangle(canvas, bgBrush);
        DrawVisibleLines(textArea);

        HRESULT hr = _d2dCtx->EndDraw();
        if (FAILED(hr))
        {
            if (hr == D2DERR_RECREATE_TARGET || hr == DXGI_ERROR_DEVICE_REMOVED || hr == DXGI_ERROR_DEVICE_RESET)
            {
                RecreateSwapChainAndTargets();
            }
            return false;
        }

        CommitBackgroundBrush(brushIndex); // commit only after successful draw
        return true;
    }

    bool RenderPartial(const D2D1_RECT_F& textArea)
    {
        if (! _hasPendingDirty)
            return false; // nothing to do

        size_t brushIndex             = 0;
        ID2D1SolidColorBrush* bgBrush = AcquireBackgroundBrush(brushIndex);

        const D2D1_RECT_F dirty = D2D1::RectF(static_cast<float>(_pendingDirtyRect.left),
                                              static_cast<float>(_pendingDirtyRect.top),
                                              static_cast<float>(_pendingDirtyRect.right),
                                              static_cast<float>(_pendingDirtyRect.bottom));

        _d2dCtx->BeginDraw();
        _d2dCtx->PushAxisAlignedClip(dirty, D2D1_ANTIALIAS_MODE_ALIASED);
        _d2dCtx->FillRectangle(dirty, bgBrush);
        DrawVisibleLines(textArea);
        _d2dCtx->PopAxisAlignedClip();

        // Draw outline on top (not clipped so it is fully visible)
        if (_brushDirtyOutline)
        {
            _d2dCtx->DrawRectangle(dirty, _brushDirtyOutline.Get(), 1.0f);
        }

        const HRESULT hr = _d2dCtx->EndDraw();
        if (FAILED(hr))
        {
            if (hr == D2DERR_RECREATE_TARGET || hr == DXGI_ERROR_DEVICE_REMOVED || hr == DXGI_ERROR_DEVICE_RESET)
            {
                RecreateSwapChainAndTargets();
            }
            return false;
        }
        CommitBackgroundBrush(brushIndex);
        return true;
    }

    void RenderSequential(const D2D1_RECT_F& textArea)
    {
        const RECT viewRect        = TextAreaRect();
        const LONG viewHeightPx    = viewRect.bottom - viewRect.top;
        const bool allowPartial    = _presentInitialized; // only after at least one full present
        const bool partialEligible = allowPartial && ! _needsFullRedraw && _hasPendingDirty && _hasPendingScroll && _pendingScrollDy > 0 && _lineHeight > 0.0f;

        LONG scrollAmount          = 0;
        bool requestFullRedrawNext = false;

        if (partialEligible)
        {
            scrollAmount = std::clamp(_pendingScrollDy, 0L, viewHeightPx);
            if (scrollAmount > 0)
            {
                _pendingDirtyRect.left   = viewRect.left;
                _pendingDirtyRect.right  = viewRect.right;
                _pendingDirtyRect.bottom = viewRect.bottom;
                _pendingDirtyRect.top    = std::max(viewRect.bottom - scrollAmount, viewRect.top);
            }
            else
            {
                _needsFullRedraw = true;
            }
        }

        bool drew        = false;
        bool usedPartial = false;

        if (partialEligible && scrollAmount > 0 && ! _needsFullRedraw)
        {
            drew        = RenderPartial(textArea);
            usedPartial = drew;
            if (! drew)
            {
                _needsFullRedraw = true; // fall back below
            }
        }

        if (_needsFullRedraw || ! _presentInitialized)
        {
            drew        = RenderFull(textArea);
            usedPartial = false;
        }

        if (! drew)
        {
            _hasPendingDirty  = false;
            _hasPendingScroll = false;
            _pendingScrollDy  = 0;
            return;
        }

        if (usedPartial)
        {
            const LONG backbufferW = static_cast<LONG>(ClientWidth());
            const LONG backbufferH = static_cast<LONG>(ClientHeight());

            RECT scrollRect{};
            scrollRect.left   = std::clamp(viewRect.left, 0L, backbufferW);
            scrollRect.right  = std::clamp(viewRect.right, scrollRect.left, backbufferW);
            scrollRect.top    = std::clamp(viewRect.top, 0L, backbufferH);
            scrollRect.bottom = std::clamp(viewRect.bottom - scrollAmount, scrollRect.top, backbufferH);

            RECT dirtyRect{};
            dirtyRect.left   = std::clamp(_pendingDirtyRect.left, 0L, backbufferW);
            dirtyRect.right  = std::clamp(_pendingDirtyRect.right, dirtyRect.left, backbufferW);
            dirtyRect.bottom = std::clamp(_pendingDirtyRect.bottom, 0L, backbufferH);
            dirtyRect.top    = std::clamp(_pendingDirtyRect.top, 0L, dirtyRect.bottom);

            const LONG sourceTop    = scrollRect.top + scrollAmount;
            const LONG sourceBottom = scrollRect.bottom + scrollAmount;
            const bool destValid    = scrollRect.left >= 0 && scrollRect.top >= 0 && scrollRect.left < scrollRect.right && scrollRect.top < scrollRect.bottom &&
                                   scrollRect.right <= backbufferW && scrollRect.bottom <= backbufferH;
            const bool sourceValid = sourceTop >= 0 && sourceTop < sourceBottom && sourceBottom <= backbufferH;

#if _DEBUG
            {
                const std::wstring msg =
                    std::format(L"[PartialPresent] attempt dy={} dest=({},{}-{},{}), src=({},{}-{},{}), dirty=({},{}-{},{}), destValid={}, srcValid={}\n",
                                scrollAmount,
                                scrollRect.left,
                                scrollRect.top,
                                scrollRect.right,
                                scrollRect.bottom,
                                scrollRect.left,
                                sourceTop,
                                scrollRect.right,
                                sourceBottom,
                                dirtyRect.left,
                                dirtyRect.top,
                                dirtyRect.right,
                                dirtyRect.bottom,
                                destValid,
                                sourceValid);
                OutputDebugStringW(msg.c_str());
            }
#endif

            if (destValid && sourceValid)
            {
                DXGI_PRESENT_PARAMETERS params{};
                params.DirtyRectsCount = 1;
                params.pDirtyRects     = &dirtyRect;
                params.pScrollRect     = &scrollRect;
                // Negative Y scrolls content up (matching how we append new lines at the bottom).
                POINT offset{0, -scrollAmount};
                params.pScrollOffset = &offset;
                const HRESULT hr     = _swapChain->Present1(1, 0, &params);
                if (FAILED(hr))
                {
#if _DEBUG
                    const std::wstring msg = std::format(L"[PartialPresent] Present1 failed hr=0x{:08X}\n", static_cast<unsigned long>(hr));
                    OutputDebugStringW(msg.c_str());
#endif
                    requestFullRedrawNext = true;
                }
                else
                {
#if _DEBUG
                    const std::wstring msg = std::format(L"[PartialPresent] dy={} dest=({},{}-{},{}), src=({},{}-{},{}), dirty=({},{}-{},{}), hr=0x{:08X}\n",
                                                         scrollAmount,
                                                         scrollRect.left,
                                                         scrollRect.top,
                                                         scrollRect.right,
                                                         scrollRect.bottom,
                                                         scrollRect.left,
                                                         sourceTop,
                                                         scrollRect.right,
                                                         sourceBottom,
                                                         dirtyRect.left,
                                                         dirtyRect.top,
                                                         dirtyRect.right,
                                                         dirtyRect.bottom,
                                                         static_cast<unsigned long>(hr));
                    OutputDebugStringW(msg.c_str());
#endif
                    _presentInitialized = true;
                }
            }
            else
            {
                DXGI_PRESENT_PARAMETERS pp{};
                _swapChain->Present1(1, 0, &pp);
#if _DEBUG
                const std::wstring msg = std::format(L"[PartialPresent] fallback full present (destValid={}, srcValid={})\n", destValid, sourceValid);
                OutputDebugStringW(msg.c_str());
#endif
                requestFullRedrawNext = true;
            }
        }
        else
        {
            DXGI_PRESENT_PARAMETERS pp{};
            _swapChain->Present1(1, 0, &pp);
            _presentInitialized = true;
        }

        _needsFullRedraw  = requestFullRedrawNext;
        _hasPendingDirty  = false;
        _hasPendingScroll = false;
        _pendingScrollDy  = 0;
    }

    void RenderDiscard(const D2D1_RECT_F& textArea)
    {
        if (RenderFull(textArea))
        {
            DXGI_PRESENT_PARAMETERS pp{};
            _swapChain->Present1(1, 0, &pp);
        }
        _needsFullRedraw  = false;
        _hasPendingDirty  = false;
        _hasPendingScroll = false;
        _pendingScrollDy  = 0;
    }

    void SetTitle()
    {
        std::wstring title = kWindowTitle;
        title += L"  [";
        title += (_mode == SwapMode::Sequential) ? L"FLIP_SEQUENTIAL + dirty rects" : L"FLIP_DISCARD";
        title += L"]";
        SetWindowTextW(_hwnd, title.c_str());
    }

    void Init(HWND h, bool useDiscard)
    {
        _hwnd       = h;
        _mode       = useDiscard ? SwapMode::Discard : SwapMode::Sequential;
        _swapEffect = useDiscard ? DXGI_SWAP_EFFECT_FLIP_DISCARD : DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
        GetClientRect(_hwnd, &_client);

        CreateD3D();
        CreateD2D();
        CreateTextResources();
        ResetLogState();
        WarmUpLogs(80);
        RecreateSwapChainAndTargets();

        SetTimer(_hwnd, 1, 16, nullptr); // ~60 FPS
        SetTitle();
    }

    void CreateD3D()
    {
        UINT flags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
#if _DEBUG
        flags |= D3D11_CREATE_DEVICE_DEBUG;
#endif
        D3D_FEATURE_LEVEL fls[] = {D3D_FEATURE_LEVEL_11_1, D3D_FEATURE_LEVEL_11_0, D3D_FEATURE_LEVEL_10_1, D3D_FEATURE_LEVEL_10_0};
        D3D_FEATURE_LEVEL flOut;
        HRESULT hr =
            D3D11CreateDevice(nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, flags, fls, ARRAYSIZE(fls), D3D11_SDK_VERSION, &_d3dDevice, &flOut, &_d3dCtx);
        if (FAILED(hr))
        {
            MessageBoxW(_hwnd, L"D3D11CreateDevice failed", L"Error", MB_ICONERROR);
        }
    }

    void CreateD2D()
    {
        D2D1_FACTORY_OPTIONS opts = {};
#if _DEBUG
        opts.debugLevel = D2D1_DEBUG_LEVEL_INFORMATION;
#endif
        HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_MULTI_THREADED, __uuidof(ID2D1Factory1), &opts, reinterpret_cast<void**>(_d2dFactory.GetAddressOf()));
        if (FAILED(hr))
        {
            MessageBoxW(_hwnd, L"D2D1CreateFactory failed", L"Error", MB_ICONERROR);
        }

        ComPtr<IDXGIDevice> dxgiDevice;
        _d3dDevice.As(&dxgiDevice);

        hr = _d2dFactory->CreateDevice(dxgiDevice.Get(), &_d2dDevice);
        if (FAILED(hr))
        {
            MessageBoxW(_hwnd, L"D2D1 CreateDevice failed", L"Error", MB_ICONERROR);
        }

        hr = _d2dDevice->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE, &_d2dCtx);
        if (FAILED(hr))
        {
            MessageBoxW(_hwnd, L"D2D CreateDeviceContext failed", L"Error", MB_ICONERROR);
        }

        hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_dwFactory.GetAddressOf()));
        if (FAILED(hr))
        {
            MessageBoxW(_hwnd, L"DWriteCreateFactory failed", L"Error", MB_ICONERROR);
        }

        UINT dpi  = GetDpiForWindow(_hwnd);
        _dpiScale = static_cast<float>(dpi) / 96.0f;
    }

    void CreateTextResources()
    {
        HRESULT hr = _dwFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 16.0f, L"en-us", &_textFormat);
        if (FAILED(hr))
        {
            MessageBoxW(_hwnd, L"CreateTextFormat failed", L"Error", MB_ICONERROR);
            return;
        }

        _textFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
        _textFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        // elipsis and no-wrap for single line measurements
        _textFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP); // single line, else wrapping first
        DWRITE_TRIMMING trim{};
        trim.granularity = DWRITE_TRIMMING_GRANULARITY_CHARACTER; // or WORD
        ComPtr<IDWriteInlineObject> ellipsis;
        _dwFactory->CreateEllipsisTrimmingSign(_textFormat.Get(), &ellipsis);
        _textFormat->SetTrimming(&trim, ellipsis.Get());

        ComPtr<IDWriteTextLayout> layout;
        hr = _dwFactory->CreateTextLayout(L"M", 1, _textFormat.Get(), 1024.0f, 1024.0f, &layout);
        if (SUCCEEDED(hr))
        {
            DWRITE_LINE_METRICS metrics{};
            UINT32 actual = 0;
            hr            = layout->GetLineMetrics(&metrics, 1, &actual);
            if (SUCCEEDED(hr) && actual > 0)
            {
                _lineHeight = metrics.height * _dpiScale;
                _baseline   = metrics.baseline * _dpiScale;
            }
        }

        if (_lineHeight <= 0.0f)
        {
            const float fontSize = _textFormat->GetFontSize();
            _lineHeight          = fontSize * 1.2f * _dpiScale;
            _baseline            = fontSize * _dpiScale;
        }
        SnapLineMetricsToPixels();
    }

    void RecreateSwapChainAndTargets()
    {
        if (_d2dCtx)
        {
            _d2dCtx->SetTarget(nullptr);
        }
        _d2dTarget.Reset();
        _swapChain.Reset();

        ComPtr<IDXGIDevice> dxgiDevice;
        _d3dDevice.As(&dxgiDevice);
        ComPtr<IDXGIAdapter> adapter;
        dxgiDevice->GetAdapter(&adapter);
        ComPtr<IDXGIFactory2> factory;
        adapter->GetParent(__uuidof(IDXGIFactory2), &factory);

        DXGI_SWAP_CHAIN_DESC1 desc = {};

        // If you specify the width as zero when you call the IDXGIFactory2::CreateSwapChainForHwnd method to
        // create a swap chain, the runtime obtains the width from the output window and assigns this width value
        // to the swap-chain description.
        desc.Width            = 0;
        desc.Height           = 0;
        desc.Format           = DXGI_FORMAT_B8G8R8A8_UNORM;
        desc.SampleDesc.Count = 1;
        desc.BufferUsage      = DXGI_USAGE_RENDER_TARGET_OUTPUT | DXGI_USAGE_SHADER_INPUT;
        desc.BufferCount      = 2;
        desc.Scaling          = DXGI_SCALING_NONE;
        desc.SwapEffect       = _swapEffect;
        desc.AlphaMode        = DXGI_ALPHA_MODE_IGNORE;
        desc.Flags            = 0;

        HRESULT hr = factory->CreateSwapChainForHwnd(_d3dDevice.Get(), _hwnd, &desc, nullptr, nullptr, &_swapChain);
        if (FAILED(hr))
        {
            MessageBoxW(_hwnd, L"CreateSwapChainForHwnd failed", L"Error", MB_ICONERROR);
            return;
        }

        ComPtr<IDXGISurface> backBuffer;
        hr = _swapChain->GetBuffer(0, __uuidof(IDXGISurface), &backBuffer);
        if (FAILED(hr))
        {
            MessageBoxW(_hwnd, L"GetBuffer failed", L"Error", MB_ICONERROR);
            return;
        }

        float dpi = 96.0f * _dpiScale;
        D2D1_BITMAP_PROPERTIES1 props =
            D2D1::BitmapProperties1(D2D1_BITMAP_OPTIONS_TARGET, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), dpi, dpi);
        hr = _d2dCtx->CreateBitmapFromDxgiSurface(backBuffer.Get(), &props, &_d2dTarget);
        if (FAILED(hr))
        {
            MessageBoxW(_hwnd, L"CreateBitmapFromDxgiSurface failed", L"Error", MB_ICONERROR);
            return;
        }
        _d2dCtx->SetTarget(_d2dTarget.Get());
        _d2dCtx->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_DEFAULT);

        _brushText.Reset();
        const auto textColor = D2D1::ColorF(D2D1::ColorF::White);
        _d2dCtx->CreateSolidColorBrush(textColor, &_brushText);
        _brushDirtyOutline.Reset();
        _d2dCtx->CreateSolidColorBrush(D2D1::ColorF(0.f, 0.f, 0.f, 1.f), &_brushDirtyOutline);
        CreateBrushPalette();

        _needsFullRedraw    = true;
        _hasPendingDirty    = false;
        _hasPendingScroll   = false;
        _pendingScrollDy    = 0;
        _presentInitialized = false; // force an initial full present before partial usage
        _scrollMaxCached    = std::numeric_limits<int>::min();
        _scrollPageCached   = 0;
        _scrollPosCached    = std::numeric_limits<int>::min();

        const float viewHeight = std::max(TextAreaRectF().bottom - TextAreaRectF().top, _lineHeight);
        UpdateScrollBar(viewHeight);
    }

    void OnResize(UINT w, UINT h)
    {
        _client.right  = _client.left + static_cast<LONG>(w);
        _client.bottom = _client.top + static_cast<LONG>(h);
        RecreateSwapChainAndTargets();

        const D2D1_RECT_F area = TextAreaRectF();
        const float viewHeight = std::max(area.bottom - area.top, _lineHeight);
        const float maxOffset  = std::max(ContentHeight() - viewHeight, 0.0f);
        if (_autoScrollToTail)
        {
            _scrollOffset = maxOffset;
        }
        else
        {
            _scrollOffset = std::clamp(_scrollOffset, 0.0f, maxOffset);
        }
        UpdateScrollBar(viewHeight);
        _needsFullRedraw = true;
    }

    void ScrollToTail()
    {
        const D2D1_RECT_F area = TextAreaRectF();
        const float viewHeight = std::max(area.bottom - area.top, _lineHeight);
        const float maxOffset  = std::max(ContentHeight() - viewHeight, 0.0f);
        _scrollOffset          = maxOffset;
        _autoScrollToTail      = true;
        UpdateScrollBar(viewHeight);
        _needsFullRedraw  = true;
        _hasPendingDirty  = false;
        _hasPendingScroll = false;
        _pendingScrollDy  = 0;
    }

    void OnScroll(UINT code, int pos)
    {
        const D2D1_RECT_F area = TextAreaRectF();
        const float viewHeight = std::max(area.bottom - area.top, _lineHeight);
        const float maxOffset  = std::max(ContentHeight() - viewHeight, 0.0f);
        float newOffset        = _scrollOffset;

        switch (code)
        {
            case SB_LINEUP: newOffset -= _lineHeight; break;
            case SB_LINEDOWN: newOffset += _lineHeight; break;
            case SB_PAGEUP: newOffset -= viewHeight; break;
            case SB_PAGEDOWN: newOffset += viewHeight; break;
            case SB_TOP: newOffset = 0.0f; break;
            case SB_BOTTOM: newOffset = maxOffset; break;
            case SB_THUMBPOSITION:
            case SB_THUMBTRACK: newOffset = static_cast<float>(pos); break;
            default: return;
        }

        newOffset         = std::clamp(newOffset, 0.0f, maxOffset);
        _autoScrollToTail = (maxOffset - newOffset) < (_lineHeight * 0.5f);
        if (std::fabs(newOffset - _scrollOffset) > 0.1f)
        {
            _scrollOffset    = newOffset;
            _needsFullRedraw = true;
        }
        UpdateScrollBar(viewHeight);
    }

    // Helper: prepare scroll + dirty rect for appended lines (multi-line capable)
    void PrepareScrollForAppended(UINT appended, const RECT& viewRect)
    {
        if (appended == 0 || _lineHeight <= 0.0f)
            return;
        // Snap scroll distance to nearest pixel to keep alignment stable
        const LONG lineHeightPx = static_cast<LONG>(std::lround(_lineHeight));
        const LONG dy           = lineHeightPx * static_cast<LONG>(appended);
        if (dy <= 0)
            return;
        _pendingDirtyRect = viewRect;
        // Clamp to at least 1 pixel high and within bounds
        _pendingDirtyRect.top = std::max(_pendingDirtyRect.bottom - dy, _pendingDirtyRect.top);
        _pendingScrollDy      = dy;
        _hasPendingDirty      = true;
        _hasPendingScroll     = true;
        _needsFullRedraw      = false;
    }

    void Tick()
    {
        if (! _d2dCtx || ! _swapChain)
            return;

        const auto now      = std::chrono::steady_clock::now();
        const UINT appended = MaybeAppendLines(now);
        _lastAppendedCount  = appended;

        const D2D1_RECT_F textArea = TextAreaRectF();
        const float viewHeight     = std::max(textArea.bottom - textArea.top, _lineHeight);
        UpdateScrollBar(viewHeight);

        const float previousOffset = _scrollOffset;
        const float maxOffset      = std::max(ContentHeight() - viewHeight, 0.0f);

        if (_autoScrollToTail)
        {
            _scrollOffset = maxOffset;
        }
        else
        {
            _scrollOffset = std::clamp(_scrollOffset, 0.0f, maxOffset);
        }

        const float delta   = _scrollOffset - previousOffset;
        const RECT viewRect = TextAreaRect();

        bool usePartial = false;
        if (appended > 0 && _autoScrollToTail && _mode == SwapMode::Sequential && _lineHeight > 0.0f)
        {
            // Expect delta roughly appended * lineHeight
            const LONG expectedDy = static_cast<LONG>(std::lround(static_cast<double>(appended) * _lineHeight));
            const LONG actualDy   = static_cast<LONG>(std::lround(delta));
            if (actualDy > 0 && std::abs(actualDy - expectedDy) <= 1)
            {
                PrepareScrollForAppended(appended, viewRect);
                usePartial = true;
            }
            else
            {
                _needsFullRedraw = true;
            }
        }
        else if (appended > 0)
        {
            _needsFullRedraw = true;
        }

        if (! usePartial)
        {
            _hasPendingDirty  = false;
            _hasPendingScroll = false;
            _pendingScrollDy  = 0;
        }

        switch (_mode)
        {
            case SwapMode::Sequential: RenderSequential(textArea); break;
            case SwapMode::Discard: RenderDiscard(textArea); break;
        }

        _lastRenderedLineCount = _lines.size();
    }
};

static App g_app;

static LRESULT OnMainWindowSize(UINT width, UINT height)
{
    if (g_app._swapChain)
    {
        g_app.OnResize(width, height);
    }
    return 0;
}

static LRESULT OnMainWindowTimer()
{
    g_app.Tick();
    return 0;
}

static LRESULT OnMainWindowVScroll(WORD request, WORD trackPos)
{
    g_app.OnScroll(request, trackPos);
    return 0;
}

static LRESULT OnMainWindowMouseWheel(int delta)
{
    const int steps = (delta >= 0 ? delta : -delta) / WHEEL_DELTA;
    for (int i = 0; i < steps; ++i)
    {
        g_app.OnScroll(delta > 0 ? SB_LINEUP : SB_LINEDOWN, 0);
    }
    return 0;
}

static LRESULT OnMainWindowKeyDown(WPARAM key)
{
    if (key == 'D')
    {
        g_app.ToggleMode();
        return 0;
    }

    if (key == VK_END)
    {
        g_app.ScrollToTail();
        return 0;
    }

    if (key >= '0' && key <= '9')
    {
        const int digit = static_cast<int>(key - '0');
        g_app.SetAppendSpeedDigit(digit);
        return 0;
    }

    return 0;
}

LRESULT CALLBACK WndProc(HWND h, UINT m, WPARAM w, LPARAM l)
{
    switch (m)
    {
        case WM_SIZE: return OnMainWindowSize(LOWORD(l), HIWORD(l));
        case WM_TIMER: return OnMainWindowTimer();
        case WM_VSCROLL: return OnMainWindowVScroll(LOWORD(w), HIWORD(w));
        case WM_MOUSEWHEEL: return OnMainWindowMouseWheel(GET_WHEEL_DELTA_WPARAM(w));
        case WM_KEYDOWN:
            if (w == VK_ESCAPE)
            {
                DestroyWindow(h);
                return 0;
            }
            return OnMainWindowKeyDown(w);
        case WM_DESTROY: PostQuitMessage(0); return 0;
    }
    return DefWindowProcW(h, m, w, l);
}

int APIENTRY wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR cmdLine, int show)
{
    bool startDiscard = (cmdLine && wcsstr(cmdLine, L"--discard") != nullptr);

    WNDCLASSEXW wc{sizeof(wc)};
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_OWNDC;
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInst;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = kClassName;
    RegisterClassExW(&wc);

    RECT r{0, 0, 900, 540};
    AdjustWindowRect(&r, WS_OVERLAPPEDWINDOW | WS_VSCROLL, FALSE);
    HWND hwnd = CreateWindowExW(0,
                                kClassName,
                                kWindowTitle,
                                WS_OVERLAPPEDWINDOW | WS_VSCROLL,
                                CW_USEDEFAULT,
                                CW_USEDEFAULT,
                                r.right - r.left,
                                r.bottom - r.top,
                                nullptr,
                                nullptr,
                                hInst,
                                nullptr);

    ShowWindow(hwnd, show);
    UpdateWindow(hwnd);

    g_app.Init(hwnd, startDiscard);

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0) > 0)
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    return 0;
}
