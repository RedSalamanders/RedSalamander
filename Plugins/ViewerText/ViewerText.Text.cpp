#include "ViewerText.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <format>
#include <limits>
#include <optional>
#include <string_view>
#include <vector>

#include <d2d1.h>
#include <d2d1helper.h>
#include <dwrite.h>

#include "ViewerText.ThemeHelpers.h"

#include "Helpers.h"

#include "resource.h"

extern HINSTANCE g_hInstance;

namespace
{
constexpr float kMonoFontSizeDip = 10.0f * 96.0f / 72.0f;

uint32_t StableHash32(std::wstring_view text) noexcept
{
    uint32_t hash = 2166136261u;
    for (wchar_t ch : text)
    {
        hash ^= static_cast<uint32_t>(ch);
        hash *= 16777619u;
    }
    return hash;
}

COLORREF ColorFromHSV(float hueDegrees, float saturation, float value) noexcept
{
    const float h = std::fmod(std::max(0.0f, hueDegrees), 360.0f);
    const float s = std::clamp(saturation, 0.0f, 1.0f);
    const float v = std::clamp(value, 0.0f, 1.0f);

    const float c = v * s;
    const float x = c * (1.0f - std::fabs(std::fmod(h / 60.0f, 2.0f) - 1.0f));
    const float m = v - c;

    float rf = 0.0f;
    float gf = 0.0f;
    float bf = 0.0f;

    if (h < 60.0f)
    {
        rf = c;
        gf = x;
        bf = 0.0f;
    }
    else if (h < 120.0f)
    {
        rf = x;
        gf = c;
        bf = 0.0f;
    }
    else if (h < 180.0f)
    {
        rf = 0.0f;
        gf = c;
        bf = x;
    }
    else if (h < 240.0f)
    {
        rf = 0.0f;
        gf = x;
        bf = c;
    }
    else if (h < 300.0f)
    {
        rf = x;
        gf = 0.0f;
        bf = c;
    }
    else
    {
        rf = c;
        gf = 0.0f;
        bf = x;
    }

    const auto toByte = [](float v01) noexcept
    {
        const float scaled = std::clamp(v01 * 255.0f, 0.0f, 255.0f);
        return static_cast<BYTE>(std::lround(scaled));
    };

    const BYTE r = toByte(rf + m);
    const BYTE g = toByte(gf + m);
    const BYTE b = toByte(bf + m);
    return RGB(r, g, b);
}

COLORREF ResolveAccentColor(const ViewerTheme& theme, std::wstring_view seed) noexcept
{
    if (theme.rainbowMode)
    {
        const uint32_t h = StableHash32(seed);
        const float hue  = static_cast<float>(h % 360u);
        const float sat  = theme.darkBase ? 0.70f : 0.55f;
        const float val  = theme.darkBase ? 0.95f : 0.85f;
        return ColorFromHSV(hue, sat, val);
    }

    return ColorRefFromArgb(theme.accentArgb);
}

float DipsFromPixels(int px, UINT dpi) noexcept
{
    if (dpi == 0)
    {
        return static_cast<float>(px);
    }

    return static_cast<float>(px) * 96.0f / static_cast<float>(dpi);
}

size_t DecimalDigits(uint64_t value) noexcept
{
    size_t digits = 1;
    while (value >= 10u)
    {
        value /= 10u;
        digits += 1;
    }

    return digits;
}

size_t LineNumberDigits(size_t lineCount) noexcept
{
    const uint64_t maxLine = (lineCount == 0) ? 1u : static_cast<uint64_t>(lineCount);
    return std::max<size_t>(3u, DecimalDigits(maxLine));
}

D2D1_COLOR_F ColorFFromColorRef(COLORREF color, float alpha = 1.0f) noexcept
{
    const float r = static_cast<float>(GetRValue(color)) / 255.0f;
    const float g = static_cast<float>(GetGValue(color)) / 255.0f;
    const float b = static_cast<float>(GetBValue(color)) / 255.0f;
    return D2D1::ColorF(r, g, b, alpha);
}

bool IsValidUtf8(const uint8_t* data, size_t size) noexcept
{
    if (! data || size == 0)
    {
        return true;
    }

    size_t i = 0;
    while (i < size)
    {
        const uint8_t b0 = data[i];
        if (b0 <= 0x7Fu)
        {
            i += 1;
            continue;
        }

        if (b0 < 0xC2u)
        {
            return false;
        }

        if (b0 <= 0xDFu)
        {
            if ((i + 1) >= size)
            {
                return true;
            }

            const uint8_t b1 = data[i + 1];
            if ((b1 & 0xC0u) != 0x80u)
            {
                return false;
            }

            i += 2;
            continue;
        }

        if (b0 <= 0xEFu)
        {
            if ((i + 2) >= size)
            {
                return true;
            }

            const uint8_t b1 = data[i + 1];
            const uint8_t b2 = data[i + 2];

            if ((b1 & 0xC0u) != 0x80u || (b2 & 0xC0u) != 0x80u)
            {
                return false;
            }

            if (b0 == 0xE0u && b1 < 0xA0u)
            {
                return false;
            }
            if (b0 == 0xEDu && b1 >= 0xA0u)
            {
                return false;
            }

            i += 3;
            continue;
        }

        if (b0 <= 0xF4u)
        {
            if ((i + 3) >= size)
            {
                return true;
            }

            const uint8_t b1 = data[i + 1];
            const uint8_t b2 = data[i + 2];
            const uint8_t b3 = data[i + 3];

            if ((b1 & 0xC0u) != 0x80u || (b2 & 0xC0u) != 0x80u || (b3 & 0xC0u) != 0x80u)
            {
                return false;
            }

            if (b0 == 0xF0u && b1 < 0x90u)
            {
                return false;
            }
            if (b0 == 0xF4u && b1 >= 0x90u)
            {
                return false;
            }

            i += 4;
            continue;
        }

        return false;
    }

    return true;
}

bool CopyUnicodeTextToClipboard(HWND hwnd, const std::wstring& text) noexcept
{
    if (! OpenClipboard(hwnd))
    {
        return false;
    }

    auto closeClipboard = wil::scope_exit([&] { CloseClipboard(); });
    if (EmptyClipboard() == 0)
    {
        return false;
    }

    const size_t bytes = (text.size() + 1) * sizeof(wchar_t);
    wil::unique_hglobal storage(GlobalAlloc(GMEM_MOVEABLE, bytes));
    if (! storage)
    {
        return false;
    }

    void* buffer = GlobalLock(storage.get());
    if (! buffer)
    {
        return false;
    }

    memcpy(buffer, text.c_str(), bytes);
    GlobalUnlock(storage.get());

    if (SetClipboardData(CF_UNICODETEXT, storage.get()) == nullptr)
    {
        return false;
    }

    storage.release();
    return true;
}
} // namespace

// Text viewer implementation moved from ViewerText.cpp.
LRESULT ViewerText::OnTextViewSize(HWND hwnd, UINT32 width, UINT32 height) noexcept
{
    if (_textViewTarget && width > 0 && height > 0)
    {
        const HRESULT hr = _textViewTarget->Resize(D2D1::SizeU(width, height));
        if (FAILED(hr))
        {
            DiscardTextViewDirect2D();
        }
    }

    RebuildTextVisualLines(hwnd);
    UpdateTextViewScrollBars(hwnd);
    InvalidateRect(hwnd, nullptr, TRUE);
    return 0;
}

LRESULT ViewerText::OnTextViewVScroll(HWND hwnd, UINT scrollCode) noexcept
{
    const uint64_t totalLines = _textVisualLineStarts.empty() ? 0u : static_cast<uint64_t>(_textVisualLineStarts.size());
    if (totalLines == 0)
    {
        return 0;
    }

    const uint64_t maxLine = totalLines - 1;

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_ALL;
    GetScrollInfo(hwnd, SB_VERT, &si);

    uint64_t top   = _textTopVisualLine;
    const int code = static_cast<int>(scrollCode);
    switch (code)
    {
        case SB_TOP: top = 0; break;
        case SB_BOTTOM: top = maxLine; break;
        case SB_LINEUP:
            if (top > 0)
            {
                top -= 1;
            }
            break;
        case SB_LINEDOWN:
            if (top < maxLine)
            {
                top += 1;
            }
            break;
        case SB_PAGEUP:
        {
            const uint64_t page = std::max<uint64_t>(1u, static_cast<uint64_t>(si.nPage));
            top                 = (top > page) ? (top - page) : 0;
            break;
        }
        case SB_PAGEDOWN:
        {
            const uint64_t page = std::max<uint64_t>(1u, static_cast<uint64_t>(si.nPage));
            top                 = std::min<uint64_t>(maxLine, top + page);
            break;
        }
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION:
        {
            const int pos = (code == SB_THUMBTRACK) ? si.nTrackPos : si.nPos;
            if (maxLine <= static_cast<uint64_t>(std::numeric_limits<int>::max()))
            {
                top = static_cast<uint64_t>(std::clamp(pos, 0, static_cast<int>(maxLine)));
            }
            else
            {
                constexpr int maxPos = std::numeric_limits<int>::max();
                const int clampedPos = std::clamp(pos, 0, maxPos);
                top                  = maxLine == 0 ? 0 : (static_cast<uint64_t>(clampedPos) * maxLine) / static_cast<uint64_t>(maxPos);
            }
            break;
        }
        default: break;
    }

    if (top > maxLine)
    {
        top = maxLine;
    }

    if (top == _textTopVisualLine)
    {
        if (_textStreamActive)
        {
            const bool scrollUp   = code == SB_LINEUP || code == SB_PAGEUP || code == SB_TOP;
            const bool scrollDown = code == SB_LINEDOWN || code == SB_PAGEDOWN || code == SB_BOTTOM;

            if (scrollUp && _textTopVisualLine == 0)
            {
                static_cast<void>(TryNavigateTextStream(GetAncestor(hwnd, GA_ROOT), true));
            }
            else if (scrollDown && _textTopVisualLine >= static_cast<uint32_t>(maxLine))
            {
                static_cast<void>(TryNavigateTextStream(GetAncestor(hwnd, GA_ROOT), false));
            }
        }

        return 0;
    }

    _textTopVisualLine = static_cast<uint32_t>(top);
    UpdateTextViewScrollBars(hwnd);
    InvalidateRect(hwnd, nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
    return 0;
}

LRESULT ViewerText::OnTextViewHScroll(HWND hwnd, UINT scrollCode) noexcept
{
    if (_wrap)
    {
        return 0;
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_ALL;
    GetScrollInfo(hwnd, SB_HORZ, &si);

    uint32_t left  = _textLeftColumn;
    const int code = static_cast<int>(scrollCode);
    switch (code)
    {
        case SB_LEFT: left = 0; break;
        case SB_RIGHT: left = _textMaxLineLength; break;
        case SB_LINELEFT:
            if (left > 0)
            {
                left -= 1;
            }
            break;
        case SB_LINERIGHT: left += 1; break;
        case SB_PAGELEFT:
        {
            const uint32_t page = std::max<uint32_t>(1u, static_cast<uint32_t>(si.nPage));
            left                = (left > page) ? (left - page) : 0;
            break;
        }
        case SB_PAGERIGHT:
        {
            const uint32_t page = std::max<uint32_t>(1u, static_cast<uint32_t>(si.nPage));
            left                = left + page;
            break;
        }
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION:
        {
            const int pos = (code == SB_THUMBTRACK) ? si.nTrackPos : si.nPos;
            left          = static_cast<uint32_t>(std::max(0, pos));
            break;
        }
        default: break;
    }

    left = std::min<uint32_t>(left, _textMaxLineLength);
    if (left == _textLeftColumn)
    {
        return 0;
    }

    _textLeftColumn = left;
    UpdateTextViewScrollBars(hwnd);
    InvalidateRect(hwnd, nullptr, TRUE);
    return 0;
}

LRESULT ViewerText::OnTextViewMouseWheel(HWND hwnd, int wheelDelta) noexcept
{
    if (wheelDelta == 0)
    {
        return 0;
    }

    UINT linesPerNotch = 3;
    static_cast<void>(SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &linesPerNotch, 0));
    if (linesPerNotch == 0 || linesPerNotch == WHEEL_PAGESCROLL)
    {
        linesPerNotch = 3;
    }

    const int steps = wheelDelta / WHEEL_DELTA;
    if (steps == 0)
    {
        return 0;
    }

    const int signedDelta     = -steps * static_cast<int>(linesPerNotch);
    const uint64_t totalLines = _textVisualLineStarts.empty() ? 0u : static_cast<uint64_t>(_textVisualLineStarts.size());
    if (totalLines == 0)
    {
        return 0;
    }

    uint64_t top = _textTopVisualLine;
    if (signedDelta < 0)
    {
        const uint64_t d = static_cast<uint64_t>(-signedDelta);
        top              = (top > d) ? (top - d) : 0;
    }
    else
    {
        const uint64_t maxLine = totalLines - 1;
        top                    = std::min<uint64_t>(maxLine, top + static_cast<uint64_t>(signedDelta));
    }

    if (top != _textTopVisualLine)
    {
        _textTopVisualLine = static_cast<uint32_t>(top);
        UpdateTextViewScrollBars(hwnd);
        InvalidateRect(hwnd, nullptr, TRUE);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
        }
    }
    else if (_textStreamActive)
    {
        if (signedDelta < 0 && _textTopVisualLine == 0)
        {
            static_cast<void>(TryNavigateTextStream(GetAncestor(hwnd, GA_ROOT), true));
        }
        else if (signedDelta > 0 && ! _textVisualLineStarts.empty() && _textTopVisualLine >= static_cast<uint32_t>(_textVisualLineStarts.size() - 1))
        {
            static_cast<void>(TryNavigateTextStream(GetAncestor(hwnd, GA_ROOT), false));
        }
    }

    return 0;
}

LRESULT ViewerText::OnTextViewLButtonDown(HWND hwnd, POINT pt) noexcept
{
    SetFocus(hwnd);

    SetCapture(hwnd);

    static_cast<void>(EnsureTextViewDirect2D(hwnd));
    const UINT dpi        = GetDpiForWindow(hwnd);
    const float xDip      = DipsFromPixels(pt.x, dpi);
    const float yDip      = DipsFromPixels(pt.y, dpi);
    const float marginDip = 6.0f;
    const float charW     = (_textCharWidthDip > 0.0f) ? _textCharWidthDip : 8.0f;
    const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
    float textStartX      = marginDip;
    if (_config.showLineNumbers && charW > 0.0f)
    {
        const size_t digits      = LineNumberDigits(_textLineStarts.size());
        const size_t gutterChars = digits + 2u;
        textStartX               = marginDip + static_cast<float>(gutterChars) * charW;
    }

    auto hitTestIndex = [&]() noexcept -> size_t
    {
        if (_textVisualLineStarts.empty() || _textVisualLineLogical.empty() || lineH <= 0.0f)
        {
            return 0;
        }

        const float relY      = std::max(0.0f, yDip - marginDip);
        const uint64_t row    = static_cast<uint64_t>(std::floor(relY / lineH));
        const uint64_t visual = std::min<uint64_t>(static_cast<uint64_t>(_textTopVisualLine) + row, static_cast<uint64_t>(_textVisualLineStarts.size() - 1));

        const uint32_t logical = _textVisualLineLogical[static_cast<size_t>(visual)];
        if (logical >= _textLineEnds.size())
        {
            return 0;
        }

        uint32_t segStart = _textVisualLineStarts[static_cast<size_t>(visual)];
        uint32_t segEnd   = _textLineEnds[logical];
        if ((visual + 1) < _textVisualLineStarts.size() && _textVisualLineLogical[static_cast<size_t>(visual + 1)] == logical)
        {
            segEnd = _textVisualLineStarts[static_cast<size_t>(visual + 1)];
        }

        if (! _wrap && segEnd >= segStart && _textLeftColumn != 0)
        {
            const uint32_t skip = std::min<uint32_t>(_textLeftColumn, segEnd - segStart);
            segStart += skip;
        }

        const float relX   = std::max(0.0f, xDip - textStartX);
        const uint32_t col = charW <= 0.0f ? 0u : static_cast<uint32_t>(std::floor(relX / charW));

        const uint32_t segLen     = segEnd >= segStart ? (segEnd - segStart) : 0u;
        const uint32_t colClamped = std::min<uint32_t>(col, segLen);
        const uint32_t idx32      = segStart + colClamped;
        const size_t idx          = std::min<size_t>(static_cast<size_t>(idx32), _textBuffer.size());
        return idx;
    };

    const size_t index = hitTestIndex();

    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    _textCaretIndex  = index;
    if (! shift)
    {
        _textSelAnchor = index;
    }
    _textSelActive = index;
    _textSelecting = true;

    InvalidateRect(hwnd, nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
    return 0;
}

LRESULT ViewerText::OnTextViewMouseMove(HWND hwnd, POINT pt) noexcept
{
    if (! _textSelecting || (GetKeyState(VK_LBUTTON) & 0x8000) == 0)
    {
        return 0;
    }

    static_cast<void>(EnsureTextViewDirect2D(hwnd));
    const UINT dpi        = GetDpiForWindow(hwnd);
    const float xDip      = DipsFromPixels(pt.x, dpi);
    const float yDip      = DipsFromPixels(pt.y, dpi);
    const float marginDip = 6.0f;
    const float charW     = (_textCharWidthDip > 0.0f) ? _textCharWidthDip : 8.0f;
    const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
    float textStartX      = marginDip;
    if (_config.showLineNumbers && charW > 0.0f)
    {
        const size_t digits      = LineNumberDigits(_textLineStarts.size());
        const size_t gutterChars = digits + 2u;
        textStartX               = marginDip + static_cast<float>(gutterChars) * charW;
    }

    if (_textVisualLineStarts.empty() || _textVisualLineLogical.empty() || lineH <= 0.0f)
    {
        return 0;
    }

    const float relY      = std::max(0.0f, yDip - marginDip);
    const uint64_t row    = static_cast<uint64_t>(std::floor(relY / lineH));
    const uint64_t visual = std::min<uint64_t>(static_cast<uint64_t>(_textTopVisualLine) + row, static_cast<uint64_t>(_textVisualLineStarts.size() - 1));

    const uint32_t logical = _textVisualLineLogical[static_cast<size_t>(visual)];
    if (logical >= _textLineEnds.size())
    {
        return 0;
    }

    uint32_t segStart = _textVisualLineStarts[static_cast<size_t>(visual)];
    uint32_t segEnd   = _textLineEnds[logical];
    if ((visual + 1) < _textVisualLineStarts.size() && _textVisualLineLogical[static_cast<size_t>(visual + 1)] == logical)
    {
        segEnd = _textVisualLineStarts[static_cast<size_t>(visual + 1)];
    }

    if (! _wrap && segEnd >= segStart && _textLeftColumn != 0)
    {
        const uint32_t skip = std::min<uint32_t>(_textLeftColumn, segEnd - segStart);
        segStart += skip;
    }

    const float relX   = std::max(0.0f, xDip - textStartX);
    const uint32_t col = charW <= 0.0f ? 0u : static_cast<uint32_t>(std::floor(relX / charW));

    const uint32_t segLen     = segEnd >= segStart ? (segEnd - segStart) : 0u;
    const uint32_t colClamped = std::min<uint32_t>(col, segLen);
    const uint32_t idx32      = segStart + colClamped;
    _textSelActive            = std::min<size_t>(static_cast<size_t>(idx32), _textBuffer.size());
    _textCaretIndex           = _textSelActive;

    InvalidateRect(hwnd, nullptr, TRUE);
    return 0;
}

LRESULT ViewerText::OnTextViewLButtonUp([[maybe_unused]] HWND hwnd) noexcept
{
    if (! _textSelecting)
    {
        return 0;
    }

    ReleaseCapture();
    _textSelecting = false;
    return 0;
}

LRESULT ViewerText::OnTextViewSetFocus(HWND hwnd) noexcept
{
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
    InvalidateRect(hwnd, nullptr, TRUE);
    return 0;
}

LRESULT ViewerText::OnTextViewKillFocus(HWND hwnd) noexcept
{
    InvalidateRect(hwnd, nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
    return 0;
}

LRESULT ViewerText::OnTextViewPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
    _allowEraseBkgndTextView  = false;
    static_cast<void>(hdc);

    if (EnsureTextViewDirect2D(hwnd) && _textViewTarget && _textViewBrush)
    {
        const UINT dpi    = GetDpiForWindow(hwnd);
        const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
        const COLORREF fg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

        HRESULT hr = S_OK;
        {
            _textViewTarget->BeginDraw();
            auto endDraw = wil::scope_exit([&] { hr = _textViewTarget->EndDraw(); });

            _textViewTarget->SetTransform(D2D1::Matrix3x2F::Identity());
            _textViewTarget->Clear(ColorFFromColorRef(bg));

            RECT rc{};
            GetClientRect(hwnd, &rc);

            const float widthDip  = DipsFromPixels(static_cast<int>(rc.right - rc.left), dpi);
            const float heightDip = DipsFromPixels(static_cast<int>(rc.bottom - rc.top), dpi);
            const float marginDip = 6.0f;
            const float charW     = (_textCharWidthDip > 0.0f) ? _textCharWidthDip : 8.0f;
            const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;

            float gutterWidthDip = 0.0f;
            float textStartX     = marginDip;
            if (_config.showLineNumbers && charW > 0.0f)
            {
                const size_t digits      = LineNumberDigits(_textLineStarts.size());
                const size_t gutterChars = digits + 2u;
                gutterWidthDip           = static_cast<float>(gutterChars) * charW;
                textStartX               = marginDip + gutterWidthDip;
            }

            _textViewBrush->SetColor(ColorFFromColorRef(fg));

            const uint64_t totalVisual = _textVisualLineStarts.empty() ? 0u : static_cast<uint64_t>(_textVisualLineStarts.size());
            const uint64_t topVisual   = static_cast<uint64_t>(_textTopVisualLine);

            const size_t selStartIndex = std::min(_textSelAnchor, _textSelActive);
            const size_t selEndIndex   = std::max(_textSelAnchor, _textSelActive);
            const bool hasSelection    = selStartIndex != selEndIndex;

            const bool hasFocus = GetFocus() == hwnd;

            const std::wstring seed      = _currentPath.empty() ? std::wstring(L"viewer") : _currentPath.filename().wstring();
            const COLORREF accent        = _hasTheme ? ResolveAccentColor(_theme, seed) : RGB(0, 120, 215);
            const uint8_t selectionAlpha = (_hasTheme && _theme.darkMode) ? 90u : 70u;
            const COLORREF selectionBg   = BlendColor(bg, accent, selectionAlpha);

            const bool hasSearchHighlights    = ! _searchQuery.empty() && ! _searchMatchStarts.empty();
            const size_t searchLen            = _searchQuery.size();
            const COLORREF searchAccent       = (_hasTheme && ! _theme.highContrast) ? ResolveAccentColor(_theme, L"search") : GetSysColor(COLOR_HIGHLIGHT);
            const uint8_t searchAlpha         = (_hasTheme && _theme.darkMode) ? 60u : 40u;
            const COLORREF searchBg           = BlendColor(bg, searchAccent, searchAlpha);
            const bool selectionIsSearchMatch = hasSelection && hasSearchHighlights && searchLen > 0 && (selEndIndex - selStartIndex == searchLen) &&
                                                std::binary_search(_searchMatchStarts.begin(), _searchMatchStarts.end(), selStartIndex);
            const uint8_t selectionFocusAlpha = (_hasTheme && _theme.darkMode) ? 140u : 120u;
            const COLORREF selectionFocusedBg = BlendColor(bg, accent, selectionFocusAlpha);

            const bool showLineNumbers    = _config.showLineNumbers && gutterWidthDip > 0.0f;
            const uint8_t lineNumberAlpha = (_hasTheme && _theme.darkMode) ? 160u : 140u;
            const COLORREF lineNumberFg   = BlendColor(bg, fg, lineNumberAlpha);

            if (showLineNumbers)
            {
                const uint8_t gutterAlpha = (_hasTheme && _theme.darkMode) ? 18u : 12u;
                const COLORREF gutterBg   = BlendColor(bg, accent, gutterAlpha);
                const float gutterRight   = std::min(widthDip, std::max(0.0f, textStartX));

                _textViewBrush->SetColor(ColorFFromColorRef(gutterBg));
                _textViewTarget->FillRectangle(D2D1::RectF(0.0f, 0.0f, gutterRight, heightDip), _textViewBrush.get());

                const COLORREF divider = BlendColor(bg, fg, (_hasTheme && _theme.darkMode) ? 40u : 20u);
                _textViewBrush->SetColor(ColorFFromColorRef(divider));
                const float sepX = std::min(widthDip, std::max(0.0f, textStartX - 1.0f));
                _textViewTarget->DrawLine(D2D1::Point2F(sepX, 0.0f), D2D1::Point2F(sepX, heightDip), _textViewBrush.get(), 1.0f);

                _textViewBrush->SetColor(ColorFFromColorRef(fg));
            }

            if (totalVisual > 0 && lineH > 0.0f && _textViewFormat && ! _textVisualLineLogical.empty())
            {
                const float usableH    = std::max(0.0f, heightDip - 2.0f * marginDip);
                const uint32_t maxRows = std::max<uint32_t>(1u, static_cast<uint32_t>(std::ceil(usableH / lineH)) + 1u);

                for (uint32_t row = 0; row < maxRows; ++row)
                {
                    const uint64_t visual = topVisual + static_cast<uint64_t>(row);
                    if (visual >= totalVisual)
                    {
                        break;
                    }

                    const uint32_t logical = _textVisualLineLogical[static_cast<size_t>(visual)];
                    if (logical >= _textLineStarts.size() || logical >= _textLineEnds.size())
                    {
                        break;
                    }

                    const uint32_t segStartRaw = _textVisualLineStarts[static_cast<size_t>(visual)];
                    uint32_t segStart          = segStartRaw;
                    uint32_t segEnd            = _textLineEnds[logical];
                    if ((visual + 1) < totalVisual && _textVisualLineLogical[static_cast<size_t>(visual + 1)] == logical)
                    {
                        segEnd = _textVisualLineStarts[static_cast<size_t>(visual + 1)];
                    }

                    if (! _wrap && segEnd >= segStart && _textLeftColumn != 0)
                    {
                        const uint32_t skip = std::min<uint32_t>(_textLeftColumn, segEnd - segStart);
                        segStart += skip;
                    }

                    const size_t startIndex = std::min<size_t>(static_cast<size_t>(segStart), _textBuffer.size());
                    const size_t endIndex   = std::min<size_t>(static_cast<size_t>(segEnd), _textBuffer.size());

                    const UINT32 len = endIndex - startIndex > static_cast<size_t>(std::numeric_limits<UINT32>::max())
                                           ? std::numeric_limits<UINT32>::max()
                                           : static_cast<UINT32>(endIndex - startIndex);

                    const float x = textStartX;
                    const float y = marginDip + static_cast<float>(row) * lineH;

                    const D2D1_RECT_F lineRc = D2D1::RectF(x, y, std::max(x, widthDip - marginDip), y + lineH);

                    if (showLineNumbers)
                    {
                        const bool isFirstSegment = (visual == 0) || (_textVisualLineLogical[static_cast<size_t>(visual - 1)] != logical);
                        if (isFirstSegment)
                        {
                            const std::wstring lineNumber  = std::to_wstring(static_cast<uint64_t>(logical) + 1u);
                            const float lineNumberRight    = std::max(marginDip, textStartX - charW);
                            const D2D1_RECT_F lineNumberRc = D2D1::RectF(marginDip, y, std::max(marginDip, lineNumberRight), y + lineH);

                            _textViewBrush->SetColor(ColorFFromColorRef(lineNumberFg));
                            _textViewTarget->DrawTextW(lineNumber.c_str(),
                                                       static_cast<UINT32>(std::min<size_t>(lineNumber.size(), std::numeric_limits<UINT32>::max())),
                                                       (_textViewFormatRight ? _textViewFormatRight.get() : _textViewFormat.get()),
                                                       lineNumberRc,
                                                       _textViewBrush.get(),
                                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            _textViewBrush->SetColor(ColorFFromColorRef(fg));
                        }
                    }

                    if (hasSearchHighlights && searchLen > 0 && endIndex >= startIndex)
                    {
                        const size_t visibleStart = startIndex;
                        const size_t visibleEnd   = endIndex;
                        const size_t scanStart    = (visibleStart > searchLen) ? (visibleStart - searchLen) : 0u;

                        auto it = std::lower_bound(_searchMatchStarts.begin(), _searchMatchStarts.end(), scanStart);
                        for (; it != _searchMatchStarts.end(); ++it)
                        {
                            const size_t matchStart = *it;
                            if (matchStart >= visibleEnd)
                            {
                                break;
                            }

                            const size_t matchEnd = matchStart + searchLen;
                            if (matchEnd <= visibleStart)
                            {
                                continue;
                            }

                            const size_t hlStart = std::max(matchStart, visibleStart);
                            const size_t hlEnd   = std::min(matchEnd, visibleEnd);
                            if (hlEnd <= hlStart || charW <= 0.0f)
                            {
                                continue;
                            }

                            const size_t colStart  = hlStart - visibleStart;
                            const size_t colLen    = hlEnd - hlStart;
                            const float hlX        = x + static_cast<float>(colStart) * charW;
                            const float hlW        = static_cast<float>(colLen) * charW;
                            const D2D1_RECT_F hlRc = D2D1::RectF(hlX, y, hlX + hlW, y + lineH);

                            _textViewBrush->SetColor(ColorFFromColorRef(searchBg));
                            _textViewTarget->FillRectangle(hlRc, _textViewBrush.get());
                            _textViewBrush->SetColor(ColorFFromColorRef(fg));
                        }
                    }

                    if (hasSelection && endIndex >= startIndex)
                    {
                        const size_t visibleStart = startIndex;
                        const size_t visibleEnd   = endIndex;
                        const size_t hlStart      = std::max(selStartIndex, visibleStart);
                        const size_t hlEnd        = std::min(selEndIndex, visibleEnd);
                        if (hlEnd > hlStart && charW > 0.0f)
                        {
                            const size_t colStart  = hlStart - visibleStart;
                            const size_t colLength = hlEnd - hlStart;
                            const float hlX        = x + static_cast<float>(colStart) * charW;
                            const float hlW        = static_cast<float>(colLength) * charW;
                            const D2D1_RECT_F hlRc = D2D1::RectF(hlX, y, hlX + hlW, y + lineH);

                            _textViewBrush->SetColor(ColorFFromColorRef(selectionIsSearchMatch ? selectionFocusedBg : selectionBg));
                            _textViewTarget->FillRectangle(hlRc, _textViewBrush.get());
                            _textViewBrush->SetColor(ColorFFromColorRef(fg));
                        }
                    }

                    if (len > 0)
                    {
                        _textViewTarget->DrawTextW(
                            _textBuffer.data() + startIndex, len, _textViewFormat.get(), lineRc, _textViewBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    }

                    if (hasFocus && charW > 0.0f && _textCaretIndex >= startIndex && _textCaretIndex <= endIndex)
                    {
                        const size_t caretCol     = _textCaretIndex - startIndex;
                        const float caretX        = x + static_cast<float>(caretCol) * charW;
                        const D2D1_RECT_F caretRc = D2D1::RectF(caretX, y, caretX + 1.0f, y + lineH);
                        _textViewTarget->FillRectangle(caretRc, _textViewBrush.get());
                    }
                }
            }

            DrawLoadingOverlay(_textViewTarget.get(), _textViewBrush.get(), widthDip, heightDip);
        }

        if (hr == D2DERR_RECREATE_TARGET)
        {
            DiscardTextViewDirect2D();
        }

        return 0;
    }

    FillRect(ps.hdc, &ps.rcPaint, _backgroundBrush.get());
    return 0;
}

LRESULT ViewerText::OnTextViewKeyDown(HWND hwnd, WPARAM vk, LPARAM lParam) noexcept
{
    const HWND root = GetAncestor(hwnd, GA_ROOT);
    if (HandleShortcutKey(root, vk))
    {
        return 0;
    }

    const bool ctrl  = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

    if (ctrl && (vk == 'C' || vk == 'c'))
    {
        const size_t a = std::min(_textSelAnchor, _textSelActive);
        const size_t b = std::max(_textSelAnchor, _textSelActive);
        if (a == b)
        {
            MessageBeep(MB_ICONINFORMATION);
            return 0;
        }

        const size_t end            = std::min(b, _textBuffer.size());
        const std::wstring selected = _textBuffer.substr(a, end - a);
        if (! CopyUnicodeTextToClipboard(root, selected))
        {
            MessageBeep(MB_ICONERROR);
        }
        return 0;
    }

    if (ctrl)
    {
        return DefWindowProcW(hwnd, WM_KEYDOWN, vk, lParam);
    }

    if (vk == VK_HOME)
    {
        CommandGoToTop(root, shift);
        return 0;
    }

    if (vk == VK_END)
    {
        CommandGoToBottom(root, shift);
        return 0;
    }

    if (_textVisualLineStarts.empty() || _textVisualLineLogical.empty() || _textLineStarts.empty() || _textLineEnds.empty())
    {
        return DefWindowProcW(hwnd, WM_KEYDOWN, vk, lParam);
    }

    static_cast<void>(EnsureTextViewDirect2D(hwnd));

    auto findVisualForIndex = [&]() noexcept -> uint32_t
    {
        const uint32_t idx32 = _textCaretIndex > static_cast<size_t>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max()
                                                                                                           : static_cast<uint32_t>(_textCaretIndex);
        auto it              = std::upper_bound(_textVisualLineStarts.begin(), _textVisualLineStarts.end(), idx32);
        if (it == _textVisualLineStarts.begin())
        {
            return 0;
        }

        size_t visual = static_cast<size_t>(std::distance(_textVisualLineStarts.begin(), it) - 1);
        if (visual >= _textVisualLineStarts.size())
        {
            visual = _textVisualLineStarts.size() - 1;
        }
        return static_cast<uint32_t>(visual);
    };

    auto getSegmentBounds = [&](uint32_t visual, uint32_t& outStart, uint32_t& outEnd) noexcept -> uint32_t
    {
        visual = std::min<uint32_t>(visual, static_cast<uint32_t>(_textVisualLineStarts.size() - 1));

        const uint32_t logical = std::min<uint32_t>(_textVisualLineLogical[visual], static_cast<uint32_t>(_textLineStarts.size() - 1));

        outStart = _textVisualLineStarts[visual];
        outEnd   = _textLineEnds[logical];
        if ((visual + 1) < _textVisualLineStarts.size() && _textVisualLineLogical[visual + 1] == logical)
        {
            outEnd = _textVisualLineStarts[visual + 1];
        }

        if (outEnd < outStart)
        {
            outEnd = outStart;
        }

        return logical;
    };

    auto ensureCaretVisible = [&]() noexcept
    {
        const uint32_t totalVisual = static_cast<uint32_t>(_textVisualLineStarts.size());
        if (totalVisual == 0)
        {
            return;
        }

        const uint32_t caretVisual = findVisualForIndex();

        SCROLLINFO si{};
        si.cbSize          = sizeof(si);
        si.fMask           = SIF_PAGE;
        const BOOL hasInfo = GetScrollInfo(hwnd, SB_VERT, &si);

        uint32_t page = 1;
        if (hasInfo != 0 && si.nPage > 0)
        {
            page = static_cast<uint32_t>(si.nPage);
        }

        if (caretVisual < _textTopVisualLine)
        {
            _textTopVisualLine = caretVisual;
        }
        else if (caretVisual >= _textTopVisualLine + page)
        {
            _textTopVisualLine = caretVisual - page + 1;
        }

        _textTopVisualLine = std::min<uint32_t>(_textTopVisualLine, totalVisual - 1);

        if (! _wrap && ! _textLineStarts.empty())
        {
            uint32_t segStart        = 0;
            uint32_t segEnd          = 0;
            const uint32_t logical   = getSegmentBounds(caretVisual, segStart, segEnd);
            const uint32_t lineStart = _textLineStarts[logical];
            const size_t caretIndex  = std::min<size_t>(_textCaretIndex, _textBuffer.size());

            uint32_t caretColumn = 0;
            if (caretIndex >= static_cast<size_t>(lineStart))
            {
                const size_t col = caretIndex - static_cast<size_t>(lineStart);
                caretColumn =
                    col > static_cast<size_t>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max() : static_cast<uint32_t>(col);
            }

            SCROLLINFO siH{};
            siH.cbSize          = sizeof(siH);
            siH.fMask           = SIF_PAGE;
            const BOOL hasInfoH = GetScrollInfo(hwnd, SB_HORZ, &siH);

            uint32_t pageCols = 1;
            if (hasInfoH != 0 && siH.nPage > 0)
            {
                pageCols = static_cast<uint32_t>(siH.nPage);
            }

            if (caretColumn < _textLeftColumn)
            {
                _textLeftColumn = caretColumn;
            }
            else if (caretColumn >= _textLeftColumn + pageCols)
            {
                _textLeftColumn = caretColumn - pageCols + 1;
            }

            _textLeftColumn = std::min<uint32_t>(_textLeftColumn, _textMaxLineLength);
        }

        UpdateTextViewScrollBars(hwnd);
    };

    auto setCaret = [&](size_t newCaret) noexcept
    {
        newCaret = std::min(newCaret, _textBuffer.size());

        _textCaretIndex = newCaret;
        if (! shift)
        {
            _textSelAnchor = newCaret;
        }
        _textSelActive = newCaret;
    };

    const uint32_t currentVisual = findVisualForIndex();
    uint32_t segStart            = 0;
    uint32_t segEnd              = 0;
    static_cast<void>(getSegmentBounds(currentVisual, segStart, segEnd));

    const size_t segStartSize = std::min<size_t>(static_cast<size_t>(segStart), _textBuffer.size());
    // const size_t segEndSize    = std::min<size_t>(static_cast<size_t>(segEnd), _textBuffer.size());
    _textPreferredColumn = (_textCaretIndex >= segStartSize) ? (_textCaretIndex - segStartSize) : 0u;

    const uint32_t totalVisual = static_cast<uint32_t>(_textVisualLineStarts.size());
    const uint32_t lastVisual  = totalVisual > 0 ? (totalVisual - 1) : 0;

    if (vk == VK_LEFT)
    {
        if (_textCaretIndex == 0 && _textStreamActive)
        {
            if (TryNavigateTextStream(root, true))
            {
                return 0;
            }
        }

        if (_textCaretIndex > 0)
        {
            setCaret(_textCaretIndex - 1);
            ensureCaretVisible();
            InvalidateRect(hwnd, nullptr, TRUE);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
            }
            return 0;
        }
    }
    else if (vk == VK_RIGHT)
    {
        if (_textCaretIndex >= _textBuffer.size() && _textStreamActive)
        {
            if (TryNavigateTextStream(root, false))
            {
                return 0;
            }
        }

        if (_textCaretIndex < _textBuffer.size())
        {
            setCaret(_textCaretIndex + 1);
            ensureCaretVisible();
            InvalidateRect(hwnd, nullptr, TRUE);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
            }
            return 0;
        }
    }
    else if (vk == VK_UP || vk == VK_DOWN || vk == VK_PRIOR || vk == VK_NEXT)
    {
        SCROLLINFO si{};
        si.cbSize = sizeof(si);
        si.fMask  = SIF_PAGE;
        static_cast<void>(GetScrollInfo(hwnd, SB_VERT, &si));
        const uint32_t page = std::max<uint32_t>(1u, static_cast<uint32_t>(si.nPage == 0 ? 1u : si.nPage));

        uint32_t targetVisual = currentVisual;
        if (vk == VK_UP)
        {
            if (currentVisual == 0 && _textStreamActive)
            {
                if (TryNavigateTextStream(root, true))
                {
                    return 0;
                }
            }
            targetVisual = currentVisual > 0 ? (currentVisual - 1) : 0;
        }
        else if (vk == VK_DOWN)
        {
            if (currentVisual >= lastVisual && _textStreamActive)
            {
                if (TryNavigateTextStream(root, false))
                {
                    return 0;
                }
            }
            targetVisual = std::min<uint32_t>(lastVisual, currentVisual + 1);
        }
        else if (vk == VK_PRIOR)
        {
            if (currentVisual == 0 && _textStreamActive)
            {
                if (TryNavigateTextStream(root, true))
                {
                    return 0;
                }
            }
            targetVisual = currentVisual > page ? (currentVisual - page) : 0;
        }
        else if (vk == VK_NEXT)
        {
            if (currentVisual >= lastVisual && _textStreamActive)
            {
                if (TryNavigateTextStream(root, false))
                {
                    return 0;
                }
            }
            targetVisual = std::min<uint32_t>(lastVisual, currentVisual + page);
        }

        uint32_t targetStart = 0;
        uint32_t targetEnd   = 0;
        static_cast<void>(getSegmentBounds(targetVisual, targetStart, targetEnd));

        const size_t targetStartSize = std::min<size_t>(static_cast<size_t>(targetStart), _textBuffer.size());
        const size_t targetEndSize   = std::min<size_t>(static_cast<size_t>(targetEnd), _textBuffer.size());
        const size_t targetLen       = targetEndSize >= targetStartSize ? (targetEndSize - targetStartSize) : 0u;

        const size_t col = std::min<size_t>(_textPreferredColumn, targetLen);
        setCaret(targetStartSize + col);

        ensureCaretVisible();
        InvalidateRect(hwnd, nullptr, TRUE);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
        }
        return 0;
    }

    return DefWindowProcW(hwnd, WM_KEYDOWN, vk, lParam);
}

LRESULT ViewerText::TextViewProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_ERASEBKGND: return _allowEraseBkgndTextView ? DefWindowProcW(hwnd, msg, wp, lp) : 1;
        case WM_PAINT: return OnTextViewPaint(hwnd);
        case WM_SIZE: return OnTextViewSize(hwnd, static_cast<UINT32>(LOWORD(lp)), static_cast<UINT32>(HIWORD(lp)));
        case WM_VSCROLL: return OnTextViewVScroll(hwnd, static_cast<UINT>(LOWORD(wp)));
        case WM_HSCROLL: return OnTextViewHScroll(hwnd, static_cast<UINT>(LOWORD(wp)));
        case WM_MOUSEWHEEL: return OnTextViewMouseWheel(hwnd, GET_WHEEL_DELTA_WPARAM(wp));
        case WM_LBUTTONDOWN:
            return OnTextViewLButtonDown(hwnd, {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))});
        case WM_MOUSEMOVE:
            return OnTextViewMouseMove(hwnd, {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))});
        case WM_LBUTTONUP: return OnTextViewLButtonUp(hwnd);
        case WM_CAPTURECHANGED: _textSelecting = false; return 0;
        case WM_KEYDOWN: return OnTextViewKeyDown(hwnd, wp, lp);
        case WM_SETFOCUS: return OnTextViewSetFocus(hwnd);
        case WM_KILLFOCUS: return OnTextViewKillFocus(hwnd);
        default: return DefWindowProcW(hwnd, msg, wp, lp);
    }
}

void ViewerText::RebuildTextLineIndex() noexcept
{
    _textLineStarts.clear();
    _textLineEnds.clear();
    _textMaxLineLength = 0;

    const size_t size = _textBuffer.size();
    size_t start      = 0;

    for (;;)
    {
        size_t pos = start;
        while (pos < size)
        {
            const wchar_t ch = _textBuffer[pos];
            if (ch == L'\n' || ch == L'\r')
            {
                break;
            }
            pos += 1;
        }

        const uint32_t start32 =
            start > static_cast<size_t>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max() : static_cast<uint32_t>(start);
        const uint32_t end32 =
            pos > static_cast<size_t>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max() : static_cast<uint32_t>(pos);

        _textLineStarts.push_back(start32);
        _textLineEnds.push_back(end32);

        if (end32 >= start32)
        {
            _textMaxLineLength = std::max(_textMaxLineLength, end32 - start32);
        }

        if (pos >= size)
        {
            break;
        }

        if (_textBuffer[pos] == L'\r' && (pos + 1) < size && _textBuffer[pos + 1] == L'\n')
        {
            start = pos + 2;
        }
        else
        {
            start = pos + 1;
        }

        if (start > size)
        {
            start = size;
        }
    }

    if (_textLineStarts.empty())
    {
        _textLineStarts.push_back(0);
        _textLineEnds.push_back(0);
    }
}

void ViewerText::RebuildTextVisualLines(HWND hwnd) noexcept
{
    _textVisualLineStarts.clear();
    _textVisualLineLogical.clear();
    _textWrapColumns = 0;

    if (_textLineStarts.empty())
    {
        _textVisualLineStarts.push_back(0);
        _textVisualLineLogical.push_back(0);
        return;
    }

    uint32_t maxCols = std::numeric_limits<uint32_t>::max();
    if (_wrap && hwnd)
    {
        if (_textCharWidthDip <= 0.0f || _textLineHeightDip <= 0.0f)
        {
            static_cast<void>(EnsureTextViewDirect2D(hwnd));
        }

        const float charW = (_textCharWidthDip > 0.0f) ? _textCharWidthDip : 8.0f;

        RECT client{};
        GetClientRect(hwnd, &client);
        const UINT dpi        = GetDpiForWindow(hwnd);
        const float widthDip  = std::max(0.0f, DipsFromPixels(static_cast<int>(client.right - client.left), dpi));
        const float marginDip = 6.0f;
        float availDip        = std::max(0.0f, widthDip - 2.0f * marginDip);
        if (_config.showLineNumbers && charW > 0.0f)
        {
            const size_t digits      = LineNumberDigits(_textLineStarts.size());
            const size_t gutterChars = digits + 2u;
            const float gutterDip    = static_cast<float>(gutterChars) * charW;
            availDip                 = std::max(0.0f, availDip - gutterDip);
        }
        const float colsF = availDip / charW;
        maxCols           = std::max<uint32_t>(1u, static_cast<uint32_t>(std::floor(colsF)));
        _textWrapColumns  = maxCols;
        _textLeftColumn   = 0;
    }

    if (! _wrap)
    {
        _textVisualLineStarts = _textLineStarts;
        _textVisualLineLogical.resize(_textLineStarts.size());
        for (uint32_t i = 0; i < static_cast<uint32_t>(_textLineStarts.size()); ++i)
        {
            _textVisualLineLogical[i] = i;
        }
        return;
    }

    for (uint32_t line = 0; line < static_cast<uint32_t>(_textLineStarts.size()); ++line)
    {
        const uint32_t start = _textLineStarts[line];
        const uint32_t end   = _textLineEnds.size() > line ? _textLineEnds[line] : start;
        const uint32_t len   = (end >= start) ? (end - start) : 0;

        if (len == 0)
        {
            _textVisualLineStarts.push_back(start);
            _textVisualLineLogical.push_back(line);
            continue;
        }

        for (uint32_t col = 0; col < len; col += maxCols)
        {
            const uint32_t segStart = start + col;
            _textVisualLineStarts.push_back(segStart);
            _textVisualLineLogical.push_back(line);
        }
    }

    if (_textVisualLineStarts.empty())
    {
        _textVisualLineStarts.push_back(0);
        _textVisualLineLogical.push_back(0);
    }
}

void ViewerText::UpdateTextViewScrollBars(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const uint64_t totalLines = _textVisualLineStarts.empty() ? 1u : static_cast<uint64_t>(_textVisualLineStarts.size());
    const uint64_t maxLine    = totalLines > 0 ? (totalLines - 1) : 0;

    RECT client{};
    GetClientRect(hwnd, &client);
    const UINT dpi           = GetDpiForWindow(hwnd);
    const float heightDip    = std::max(1.0f, DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi));
    const float lineH        = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
    const uint32_t pageLines = std::max<uint32_t>(1u, static_cast<uint32_t>(std::floor(heightDip / lineH)));

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS | SIF_DISABLENOSCROLL;
    si.nMin   = 0;

    if (maxLine <= static_cast<uint64_t>(std::numeric_limits<int>::max()))
    {
        si.nMax  = static_cast<int>(maxLine);
        si.nPos  = static_cast<int>(std::min<uint64_t>(_textTopVisualLine, maxLine));
        si.nPage = static_cast<UINT>(pageLines);
    }
    else
    {
        constexpr int maxPos = std::numeric_limits<int>::max();
        const uint64_t top   = std::min<uint64_t>(_textTopVisualLine, maxLine);
        const uint64_t pos64 = maxLine == 0 ? 0 : (top * static_cast<uint64_t>(maxPos)) / maxLine;
        si.nMax              = maxPos;
        si.nPos              = static_cast<int>(pos64);
        si.nPage             = static_cast<UINT>(pageLines);
    }

    SetScrollInfo(hwnd, SB_VERT, &si, TRUE);

    if (_wrap)
    {
        ShowScrollBar(hwnd, SB_HORZ, FALSE);
        return;
    }

    float widthDip    = std::max(1.0f, DipsFromPixels(static_cast<int>(client.right - client.left), dpi));
    const float charW = (_textCharWidthDip > 0.0f) ? _textCharWidthDip : 8.0f;
    if (_config.showLineNumbers && charW > 0.0f)
    {
        const size_t digits      = LineNumberDigits(_textLineStarts.size());
        const size_t gutterChars = digits + 2u;
        const float gutterDip    = static_cast<float>(gutterChars) * charW;
        widthDip                 = std::max(1.0f, widthDip - gutterDip);
    }
    const uint32_t pageCols = std::max<uint32_t>(1u, static_cast<uint32_t>(std::floor(widthDip / charW)));
    const uint32_t maxCol   = _textMaxLineLength;

    SCROLLINFO siH{};
    siH.cbSize = sizeof(siH);
    siH.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS | SIF_DISABLENOSCROLL;
    siH.nMin   = 0;
    siH.nMax   = static_cast<int>(std::min<uint32_t>(maxCol, static_cast<uint32_t>(std::numeric_limits<int>::max())));
    siH.nPage  = static_cast<UINT>(pageCols);
    siH.nPos   = static_cast<int>(std::min<uint32_t>(_textLeftColumn, maxCol));
    SetScrollInfo(hwnd, SB_HORZ, &siH, TRUE);
    ShowScrollBar(hwnd, SB_HORZ, TRUE);
}

bool ViewerText::EnsureTextViewDirect2D(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    const float dpiF = static_cast<float>(GetDpiForWindow(hwnd));

    if (! _d2dFactory)
    {
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, _d2dFactory.put());
        if (FAILED(hr) || ! _d2dFactory)
        {
            _d2dFactory.reset();
            return false;
        }
    }

    if (! _dwriteFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_dwriteFactory.put()));
        if (FAILED(hr) || ! _dwriteFactory)
        {
            _dwriteFactory.reset();
            return false;
        }
    }

    if (! _textViewTarget)
    {
        RECT client{};
        GetClientRect(hwnd, &client);

        const UINT32 width     = static_cast<UINT32>(std::max<LONG>(0, client.right - client.left));
        const UINT32 height    = static_cast<UINT32>(std::max<LONG>(0, client.bottom - client.top));
        const D2D1_SIZE_U size = D2D1::SizeU(width, height);

        D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
        props.dpiX                          = dpiF;
        props.dpiY                          = dpiF;

        const D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, size);

        const HRESULT hr = _d2dFactory->CreateHwndRenderTarget(props, hwndProps, _textViewTarget.put());
        if (FAILED(hr) || ! _textViewTarget)
        {
            _textViewTarget.reset();
            return false;
        }

        _textViewTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    }
    else
    {
        _textViewTarget->SetDpi(dpiF, dpiF);
    }

    if (! _textViewBrush)
    {
        const HRESULT hr = _textViewTarget->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::Black), _textViewBrush.put());
        if (FAILED(hr) || ! _textViewBrush)
        {
            _textViewBrush.reset();
            return false;
        }
    }

    if (! _textViewFormat)
    {
        const HRESULT hr = _dwriteFactory->CreateTextFormat(L"Consolas",
                                                            nullptr,
                                                            DWRITE_FONT_WEIGHT_NORMAL,
                                                            DWRITE_FONT_STYLE_NORMAL,
                                                            DWRITE_FONT_STRETCH_NORMAL,
                                                            kMonoFontSizeDip,
                                                            L"",
                                                            _textViewFormat.put());
        if (FAILED(hr) || ! _textViewFormat)
        {
            _textViewFormat.reset();
            return false;
        }

        static_cast<void>(_textViewFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));
        static_cast<void>(_textViewFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
        static_cast<void>(_textViewFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
    }

    if (! _textViewFormatRight)
    {
        const HRESULT hr = _dwriteFactory->CreateTextFormat(L"Consolas",
                                                            nullptr,
                                                            DWRITE_FONT_WEIGHT_NORMAL,
                                                            DWRITE_FONT_STYLE_NORMAL,
                                                            DWRITE_FONT_STRETCH_NORMAL,
                                                            kMonoFontSizeDip,
                                                            L"",
                                                            _textViewFormatRight.put());
        if (FAILED(hr) || ! _textViewFormatRight)
        {
            _textViewFormatRight.reset();
            return false;
        }

        static_cast<void>(_textViewFormatRight->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING));
        static_cast<void>(_textViewFormatRight->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
        static_cast<void>(_textViewFormatRight->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
    }

    if (_textCharWidthDip <= 0.0f || _textLineHeightDip <= 0.0f)
    {
        wil::com_ptr<IDWriteTextLayout> layout;
        const HRESULT hr = _dwriteFactory->CreateTextLayout(L"0", 1, _textViewFormat.get(), 1024.0f, 1024.0f, layout.put());
        if (SUCCEEDED(hr) && layout)
        {
            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(layout->GetMetrics(&metrics)))
            {
                _textCharWidthDip  = std::max(1.0f, metrics.widthIncludingTrailingWhitespace);
                _textLineHeightDip = std::max(1.0f, metrics.height);
            }
        }
    }

    return true;
}

void ViewerText::DiscardTextViewDirect2D() noexcept
{
    _textViewBrush.reset();
    _textViewFormat.reset();
    _textViewFormatRight.reset();
    _textViewTarget.reset();
    _textCharWidthDip  = 0.0f;
    _textLineHeightDip = 0.0f;
}

void ViewerText::SetShowLineNumbers(HWND hwnd, bool showLineNumbers) noexcept
{
    _config.showLineNumbers = showLineNumbers;
    _configurationJson      = std::format("{{\"textBufferMiB\":{},\"hexBufferMiB\":{},\"showLineNumbers\":\"{}\",\"wrapText\":\"{}\"}}",
                                     _config.textBufferMiB,
                                     _config.hexBufferMiB,
                                     _config.showLineNumbers ? "1" : "0",
                                     _config.wrapText ? "1" : "0");

    if (_hEdit)
    {
        RebuildTextVisualLines(_hEdit.get());
        if (! _textVisualLineStarts.empty())
        {
            _textTopVisualLine = std::min<uint32_t>(_textTopVisualLine, static_cast<uint32_t>(_textVisualLineStarts.size() - 1));
        }
        else
        {
            _textTopVisualLine = 0;
        }

        UpdateTextViewScrollBars(_hEdit.get());
        InvalidateRect(_hEdit.get(), nullptr, TRUE);
    }

    UpdateMenuChecks(hwnd);
}

void ViewerText::SetWrap(HWND hwnd, bool wrap) noexcept
{
    _wrap              = wrap;
    _config.wrapText   = wrap;
    _configurationJson = std::format("{{\"textBufferMiB\":{},\"hexBufferMiB\":{},\"showLineNumbers\":\"{}\",\"wrapText\":\"{}\"}}",
                                     _config.textBufferMiB,
                                     _config.hexBufferMiB,
                                     _config.showLineNumbers ? "1" : "0",
                                     _config.wrapText ? "1" : "0");
    if (_hEdit)
    {
        RebuildTextVisualLines(_hEdit.get());
        if (! _textVisualLineStarts.empty())
        {
            _textTopVisualLine = std::min<uint32_t>(_textTopVisualLine, static_cast<uint32_t>(_textVisualLineStarts.size() - 1));
        }
        else
        {
            _textTopVisualLine = 0;
        }

        UpdateTextViewScrollBars(_hEdit.get());
        InvalidateRect(_hEdit.get(), nullptr, TRUE);
    }

    UpdateMenuChecks(hwnd);
}

void ViewerText::CommandFindNext(HWND hwnd, bool backward)
{
    if (_viewMode == ViewMode::Hex)
    {
        CommandFindNextHex(hwnd, backward);
        return;
    }

    if (_viewMode != ViewMode::Text)
    {
        SetViewMode(hwnd, ViewMode::Text);
    }

    if (_searchQuery.empty())
    {
        CommandFind(hwnd);
        return;
    }

    if (! _hEdit)
    {
        return;
    }

    auto setStatusAfterFind = [&]() noexcept { _statusMessage.clear(); };

    auto findAndSelect = [&](size_t start) noexcept -> bool
    {
        if (_searchQuery.empty())
        {
            return false;
        }

        const size_t queryLen = _searchQuery.size();

        size_t found = std::wstring::npos;
        if (backward)
        {
            if (_textBuffer.empty())
            {
                return false;
            }

            const size_t startPos = std::min(start, _textBuffer.size() - 1);
            found                 = _textBuffer.rfind(_searchQuery, startPos);
        }
        else
        {
            found = _textBuffer.find(_searchQuery, start);
        }

        if (found == std::wstring::npos)
        {
            return false;
        }

        setStatusAfterFind();

        const size_t matchStart = found;
        const size_t matchEnd   = std::min(found + queryLen, _textBuffer.size());

        _textSelAnchor  = matchStart;
        _textSelActive  = matchEnd;
        _textCaretIndex = matchEnd;

        auto ensureCaretVisible = [&]() noexcept
        {
            if (_textVisualLineStarts.empty() || _textVisualLineLogical.empty())
            {
                return;
            }

            const uint32_t idx32 = _textCaretIndex > static_cast<size_t>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max()
                                                                                                               : static_cast<uint32_t>(_textCaretIndex);
            auto it              = std::upper_bound(_textVisualLineStarts.begin(), _textVisualLineStarts.end(), idx32);
            uint32_t caretVisual = 0;
            if (it != _textVisualLineStarts.begin())
            {
                size_t visual = static_cast<size_t>(std::distance(_textVisualLineStarts.begin(), it) - 1);
                if (visual >= _textVisualLineStarts.size())
                {
                    visual = _textVisualLineStarts.size() - 1;
                }
                caretVisual = static_cast<uint32_t>(visual);
            }

            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_PAGE;
            static_cast<void>(GetScrollInfo(_hEdit.get(), SB_VERT, &si));
            const uint32_t page = std::max<uint32_t>(1u, static_cast<uint32_t>(si.nPage == 0 ? 1u : si.nPage));

            if (caretVisual < _textTopVisualLine)
            {
                _textTopVisualLine = caretVisual;
            }
            else if (caretVisual >= _textTopVisualLine + page)
            {
                _textTopVisualLine = caretVisual - page + 1;
            }

            if (! _wrap)
            {
                const uint32_t logical   = _textVisualLineLogical[caretVisual];
                const uint32_t lineStart = _textLineStarts[logical];
                const size_t caretIndex  = std::min<size_t>(_textCaretIndex, _textBuffer.size());

                uint32_t caretColumn = 0;
                if (caretIndex >= static_cast<size_t>(lineStart))
                {
                    const size_t col = caretIndex - static_cast<size_t>(lineStart);
                    caretColumn =
                        col > static_cast<size_t>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max() : static_cast<uint32_t>(col);
                }

                SCROLLINFO siH{};
                siH.cbSize = sizeof(siH);
                siH.fMask  = SIF_PAGE;
                static_cast<void>(GetScrollInfo(_hEdit.get(), SB_HORZ, &siH));
                const uint32_t pageCols = std::max<uint32_t>(1u, static_cast<uint32_t>(siH.nPage == 0 ? 1u : siH.nPage));

                if (caretColumn < _textLeftColumn)
                {
                    _textLeftColumn = caretColumn;
                }
                else if (caretColumn >= _textLeftColumn + pageCols)
                {
                    _textLeftColumn = caretColumn - pageCols + 1;
                }

                _textLeftColumn = std::min<uint32_t>(_textLeftColumn, _textMaxLineLength);
            }

            UpdateTextViewScrollBars(_hEdit.get());
        };

        ensureCaretVisible();
        UpdateSearchHighlights();

        InvalidateRect(_hEdit.get(), nullptr, TRUE);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
        }

        return true;
    };

    auto tryFindFromSelection = [&]() noexcept -> bool
    {
        const size_t selStart = std::min(_textSelAnchor, _textSelActive);
        const size_t selEnd   = std::max(_textSelAnchor, _textSelActive);

        if (backward)
        {
            if (selStart == 0)
            {
                return findAndSelect(0);
            }
            return findAndSelect(selStart - 1);
        }

        return findAndSelect(selEnd);
    };

    if (! _textStreamActive)
    {
        if (tryFindFromSelection())
        {
            return;
        }

        const size_t wrapStart = backward ? (_textBuffer.empty() ? 0 : (_textBuffer.size() - 1)) : 0;
        if (findAndSelect(wrapStart))
        {
            _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_SEARCH_WRAPPED);
            InvalidateRect(_hWnd.get(), nullptr, TRUE);
            ShowInlineAlert(InlineAlertSeverity::Info, IDS_VIEWERTEXT_NAME, IDS_VIEWERTEXT_MSG_SEARCH_WRAPPED);
            return;
        }

        MessageBeep(MB_ICONINFORMATION);
        return;
    }

    bool wrapped = false;
    for (;;)
    {
        if (tryFindFromSelection())
        {
            return;
        }

        const bool hasMore = backward ? (_textStreamStartOffset > _textStreamSkipBytes) : (_textStreamEndOffset < _fileSize);
        if (hasMore)
        {
            if (TryNavigateTextStream(hwnd, backward))
            {
                UpdateSearchHighlights();
                continue;
            }
        }

        if (wrapped)
        {
            MessageBeep(MB_ICONINFORMATION);
            return;
        }

        wrapped        = true;
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_SEARCH_WRAPPED);
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
        ShowInlineAlert(InlineAlertSeverity::Info, IDS_VIEWERTEXT_NAME, IDS_VIEWERTEXT_MSG_SEARCH_WRAPPED);

        if (backward)
        {
            uint64_t lastStart        = _textStreamSkipBytes;
            const uint64_t chunkBytes = TextStreamChunkBytes();
            if (_fileSize > chunkBytes)
            {
                lastStart = _fileSize - chunkBytes;
            }
            lastStart = AlignTextStreamOffset(lastStart);
            static_cast<void>(LoadTextToEdit(hwnd, lastStart, true));
        }
        else
        {
            static_cast<void>(LoadTextToEdit(hwnd, _textStreamSkipBytes, false));
        }

        UpdateSearchHighlights();
    }
}

HRESULT ViewerText::LoadTextToEdit(HWND hwnd, uint64_t startOffset, bool scrollToEnd) noexcept
{
    if (! _hEdit)
    {
        Debug::Error(L"ViewerText: LoadTextToEdit failed because the DirectX text view is missing.");
        return E_FAIL;
    }

    _textBuffer.clear();
    _searchMatchStarts.clear();
    _textLineStarts.clear();
    _textLineEnds.clear();
    _textVisualLineStarts.clear();
    _textVisualLineLogical.clear();
    _textTopVisualLine   = 0;
    _textLeftColumn      = 0;
    _textCaretIndex      = 0;
    _textSelAnchor       = 0;
    _textSelActive       = 0;
    _textPreferredColumn = 0;
    _textSelecting       = false;

    if (! _fileReader)
    {
        Debug::Error(L"ViewerText: LoadTextToEdit failed because file reader is missing for '{}'.", _currentPath.c_str());
        return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
    }

    const uint64_t clampedStart = std::min<uint64_t>(std::max(startOffset, _textStreamSkipBytes), _fileSize);
    if (clampedStart > static_cast<uint64_t>(std::numeric_limits<__int64>::max()))
    {
        Debug::Error(L"ViewerText: LoadTextToEdit failed because start offset is out of range ({}).", static_cast<unsigned long long>(clampedStart));
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    uint64_t newPosition = 0;
    const HRESULT seekHr = _fileReader->Seek(static_cast<__int64>(clampedStart), FILE_BEGIN, &newPosition);
    if (FAILED(seekHr))
    {
        Debug::Error(L"ViewerText: Seek(FILE_BEGIN, {}) failed for '{}' (hr=0x{:08X}).",
                     static_cast<unsigned long long>(clampedStart),
                     _currentPath.c_str(),
                     static_cast<unsigned long>(seekHr));
        return seekHr;
    }

    const FileEncoding displayEncoding = DisplayEncodingFileEncoding();
    const UINT displayCodePage         = DisplayEncodingCodePage();

    const uint64_t maxChunkBytes  = TextStreamChunkBytes();
    const uint64_t availableBytes = (_fileSize > clampedStart) ? (_fileSize - clampedStart) : 0;
    uint64_t wantBytes64          = availableBytes;
    if (wantBytes64 > maxChunkBytes)
    {
        wantBytes64 = maxChunkBytes;
    }
    if (wantBytes64 > static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
    {
        wantBytes64 = static_cast<uint64_t>(std::numeric_limits<size_t>::max());
    }
    size_t wantBytes = static_cast<size_t>(wantBytes64);

    std::vector<uint8_t> bytes(wantBytes);
    size_t bytesReadTotal = 0;
    while (bytesReadTotal < bytes.size())
    {
        const size_t remaining   = bytes.size() - bytesReadTotal;
        const unsigned long want = remaining > static_cast<size_t>(std::numeric_limits<unsigned long>::max()) ? std::numeric_limits<unsigned long>::max()
                                                                                                              : static_cast<unsigned long>(remaining);

        unsigned long read   = 0;
        const HRESULT readHr = _fileReader->Read(bytes.data() + bytesReadTotal, want, &read);
        if (FAILED(readHr))
        {
            Debug::Error(L"ViewerText: Read failed for '{}' (hr=0x{:08X}).", _currentPath.c_str(), static_cast<unsigned long>(readHr));
            return readHr;
        }

        if (read == 0)
        {
            break;
        }

        bytesReadTotal += static_cast<size_t>(read);
    }

    bytes.resize(bytesReadTotal);

    auto utf8IncompleteTailSize = [](const uint8_t* data, size_t size) noexcept -> size_t
    {
        if (! data || size == 0)
        {
            return 0;
        }

        size_t start = size;
        for (size_t i = size; i > 0; --i)
        {
            const uint8_t b = data[i - 1];
            if ((b & 0xC0u) != 0x80u)
            {
                start = i - 1;
                break;
            }
        }

        if (start >= size)
        {
            return 0;
        }

        const uint8_t lead = data[start];
        size_t expected    = 1;
        if (lead <= 0x7Fu)
        {
            expected = 1;
        }
        else if (lead >= 0xC2u && lead <= 0xDFu)
        {
            expected = 2;
        }
        else if (lead >= 0xE0u && lead <= 0xEFu)
        {
            expected = 3;
        }
        else if (lead >= 0xF0u && lead <= 0xF4u)
        {
            expected = 4;
        }

        const size_t available = size - start;
        if (expected > 1 && available < expected)
        {
            return available;
        }

        return 0;
    };

    size_t carryBytes = 0;
    if (displayEncoding == FileEncoding::Utf16LE || displayEncoding == FileEncoding::Utf16BE)
    {
        carryBytes = bytes.size() % 2;
    }
    else if (displayEncoding == FileEncoding::Utf32LE || displayEncoding == FileEncoding::Utf32BE)
    {
        carryBytes = bytes.size() % 4;
    }
    else if (displayCodePage == CP_UTF8)
    {
        carryBytes = utf8IncompleteTailSize(bytes.data(), bytes.size());
    }

    if (carryBytes > bytes.size())
    {
        carryBytes = bytes.size();
    }

    const size_t convertBytes = bytes.size() - carryBytes;

    if (convertBytes > 0)
    {
        if ((displayEncoding == FileEncoding::Utf16LE || displayEncoding == FileEncoding::Utf16BE) && (convertBytes % 2) == 0)
        {
            const size_t wcharCount = convertBytes / 2;
            _textBuffer.resize(wcharCount);
            memcpy(_textBuffer.data(), bytes.data(), convertBytes);

            if (displayEncoding == FileEncoding::Utf16BE)
            {
                for (size_t i = 0; i < _textBuffer.size(); ++i)
                {
                    const wchar_t v = _textBuffer[i];
                    _textBuffer[i]  = static_cast<wchar_t>((static_cast<uint16_t>(v) >> 8) | (static_cast<uint16_t>(v) << 8));
                }
            }
        }
        else if ((displayEncoding == FileEncoding::Utf32LE || displayEncoding == FileEncoding::Utf32BE) && (convertBytes % 4) == 0)
        {
            const bool bigEndian = (displayEncoding == FileEncoding::Utf32BE);
            _textBuffer.clear();
            _textBuffer.reserve(convertBytes / 4);

            for (size_t i = 0; i + 3 < convertBytes; i += 4)
            {
                uint32_t cp = 0;
                if (bigEndian)
                {
                    cp = (static_cast<uint32_t>(bytes[i]) << 24) | (static_cast<uint32_t>(bytes[i + 1]) << 16) | (static_cast<uint32_t>(bytes[i + 2]) << 8) |
                         static_cast<uint32_t>(bytes[i + 3]);
                }
                else
                {
                    cp = static_cast<uint32_t>(bytes[i]) | (static_cast<uint32_t>(bytes[i + 1]) << 8) | (static_cast<uint32_t>(bytes[i + 2]) << 16) |
                         (static_cast<uint32_t>(bytes[i + 3]) << 24);
                }

                if (cp <= 0xFFFFu)
                {
                    if (cp >= 0xD800u && cp <= 0xDFFFu)
                    {
                        _textBuffer.push_back(static_cast<wchar_t>(0xFFFDu));
                    }
                    else
                    {
                        _textBuffer.push_back(static_cast<wchar_t>(cp));
                    }
                }
                else if (cp <= 0x10FFFFu)
                {
                    const uint32_t v = cp - 0x10000u;
                    _textBuffer.push_back(static_cast<wchar_t>(0xD800u + (v >> 10)));
                    _textBuffer.push_back(static_cast<wchar_t>(0xDC00u + (v & 0x3FFu)));
                }
                else
                {
                    _textBuffer.push_back(static_cast<wchar_t>(0xFFFDu));
                }
            }
        }
        else
        {
            if (convertBytes > static_cast<size_t>(std::numeric_limits<int>::max()))
            {
                Debug::Error(L"ViewerText: Text decode buffer too large ({} bytes).", static_cast<unsigned long long>(convertBytes));
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }

            const int srcLen       = static_cast<int>(convertBytes);
            const int requiredWide = MultiByteToWideChar(displayCodePage, 0, reinterpret_cast<LPCCH>(bytes.data()), srcLen, nullptr, 0);
            if (requiredWide <= 0)
            {
                const DWORD lastError = GetLastError();
                Debug::Error(L"ViewerText: MultiByteToWideChar failed for '{}' (cp={}, lastError={}).", _currentPath.c_str(), displayCodePage, lastError);
                return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_INVALID_DATA);
            }

            _textBuffer.resize(static_cast<size_t>(requiredWide));
            const int written = MultiByteToWideChar(displayCodePage, 0, reinterpret_cast<LPCCH>(bytes.data()), srcLen, _textBuffer.data(), requiredWide);
            if (written <= 0)
            {
                const DWORD lastError = GetLastError();
                Debug::Error(L"ViewerText: MultiByteToWideChar failed for '{}' (cp={}, lastError={}).", _currentPath.c_str(), displayCodePage, lastError);
                return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_INVALID_DATA);
            }

            _textBuffer.resize(static_cast<size_t>(written));
        }
    }

    _textStreamStartOffset = clampedStart;
    _textStreamEndOffset   = std::min<uint64_t>(clampedStart + static_cast<uint64_t>(bytesReadTotal) - static_cast<uint64_t>(carryBytes), _fileSize);
    _textStreamActive      = (_fileSize > _textStreamSkipBytes) && ((_fileSize - _textStreamSkipBytes) > maxChunkBytes);

    const UINT defaultCodePage = GetACP();
    _detectedCodePage          = 0;
    _detectedCodePageValid     = false;
    _detectedCodePageIsGuess   = false;

    switch (_encoding)
    {
        case FileEncoding::Utf8:
            _detectedCodePage        = CP_UTF8;
            _detectedCodePageValid   = true;
            _detectedCodePageIsGuess = false;
            break;
        case FileEncoding::Utf16LE:
            _detectedCodePage        = 1200u;
            _detectedCodePageValid   = true;
            _detectedCodePageIsGuess = false;
            break;
        case FileEncoding::Utf16BE:
            _detectedCodePage        = 1201u;
            _detectedCodePageValid   = true;
            _detectedCodePageIsGuess = false;
            break;
        case FileEncoding::Utf32LE:
            _detectedCodePage        = 12000u;
            _detectedCodePageValid   = true;
            _detectedCodePageIsGuess = false;
            break;
        case FileEncoding::Utf32BE:
            _detectedCodePage        = 12001u;
            _detectedCodePageValid   = true;
            _detectedCodePageIsGuess = false;
            break;
        case FileEncoding::Unknown:
        default:
        {
            _detectedCodePageIsGuess = true;
            if (! bytes.empty() && IsValidUtf8(bytes.data(), bytes.size()))
            {
                _detectedCodePage = CP_UTF8;
            }
            else
            {
                _detectedCodePage = defaultCodePage;
            }
            _detectedCodePageValid = true;
            break;
        }
    }

    RebuildTextLineIndex();
    UpdateTextStreamTotalLineCountAfterLoad();
    RebuildTextVisualLines(_hEdit.get());

    if (scrollToEnd && ! _textVisualLineStarts.empty())
    {
        _textTopVisualLine = static_cast<uint32_t>(_textVisualLineStarts.size() - 1);
        _textCaretIndex    = _textBuffer.size();
    }

    _textSelAnchor = _textCaretIndex;
    _textSelActive = _textCaretIndex;

    UpdateSearchHighlights();
    UpdateTextViewScrollBars(_hEdit.get());

    InvalidateRect(_hEdit.get(), nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }

    static_cast<void>(hwnd);
    return S_OK;
}

void ViewerText::UpdateTextStreamTotalLineCountAfterLoad() noexcept
{
    if (! _textStreamActive)
    {
        return;
    }

    if (_textTotalLineCount.has_value())
    {
        return;
    }

    if (_textStreamLineCountedEndOffset != _textStreamStartOffset)
    {
        return;
    }

    if (_textStreamEndOffset <= _textStreamStartOffset)
    {
        if (_textStreamEndOffset == _textStreamStartOffset && _textStreamStartOffset < _fileSize)
        {
            uint64_t totalLines = _textStreamLineCountedNewlines;
            if (totalLines < std::numeric_limits<uint64_t>::max())
            {
                totalLines += 1;
            }
            _textTotalLineCount = totalLines;
        }
        return;
    }

    uint64_t chunkNewlines = 0;
    if (! _textLineStarts.empty())
    {
        const uint64_t linesInChunk = static_cast<uint64_t>(_textLineStarts.size());
        if (linesInChunk > 0)
        {
            chunkNewlines = linesInChunk - 1;
        }
    }

    if (_textStreamLineCountLastWasCR && ! _textBuffer.empty() && _textBuffer.front() == L'\n')
    {
        if (chunkNewlines > 0)
        {
            chunkNewlines -= 1;
        }
    }

    if (chunkNewlines > 0)
    {
        constexpr uint64_t maxValue = std::numeric_limits<uint64_t>::max();
        if (_textStreamLineCountedNewlines > maxValue - chunkNewlines)
        {
            _textStreamLineCountedNewlines = maxValue;
        }
        else
        {
            _textStreamLineCountedNewlines += chunkNewlines;
        }
    }

    if (! _textBuffer.empty())
    {
        _textStreamLineCountLastWasCR = (_textBuffer.back() == L'\r');
    }

    _textStreamLineCountedEndOffset = _textStreamEndOffset;

    if (_textStreamLineCountedEndOffset >= _fileSize)
    {
        uint64_t totalLines = _textStreamLineCountedNewlines;
        if (totalLines < std::numeric_limits<uint64_t>::max())
        {
            totalLines += 1;
        }
        _textTotalLineCount = totalLines;
    }
}

uint64_t ViewerText::TextStreamChunkBytes() const noexcept
{
    uint64_t bytes = static_cast<uint64_t>(_config.textBufferMiB) * static_cast<uint64_t>(1024u) * static_cast<uint64_t>(1024u);
    bytes          = std::clamp<uint64_t>(bytes, 256u * 1024u, 256u * 1024u * 1024u);

    const FileEncoding encoding = DisplayEncodingFileEncoding();
    if (encoding == FileEncoding::Utf16LE || encoding == FileEncoding::Utf16BE)
    {
        bytes &= ~static_cast<uint64_t>(1);
        bytes = std::max<uint64_t>(bytes, 2u);
    }
    else if (encoding == FileEncoding::Utf32LE || encoding == FileEncoding::Utf32BE)
    {
        bytes &= ~static_cast<uint64_t>(3);
        bytes = std::max<uint64_t>(bytes, 4u);
    }

    return bytes;
}

uint64_t ViewerText::AlignTextStreamOffset(uint64_t offset) const noexcept
{
    uint64_t aligned            = offset;
    const FileEncoding encoding = DisplayEncodingFileEncoding();
    if (encoding == FileEncoding::Utf16LE || encoding == FileEncoding::Utf16BE)
    {
        aligned &= ~static_cast<uint64_t>(1);
    }
    else if (encoding == FileEncoding::Utf32LE || encoding == FileEncoding::Utf32BE)
    {
        aligned &= ~static_cast<uint64_t>(3);
    }

    aligned = std::max(aligned, _textStreamSkipBytes);
    aligned = std::min(aligned, _fileSize);
    return aligned;
}

bool ViewerText::TryNavigateTextStream(HWND hwnd, bool backward) noexcept
{
    if (! hwnd || ! _hEdit || ! _fileReader || _currentPath.empty() || _fileSize == 0)
    {
        return false;
    }

    const uint64_t chunkBytes = TextStreamChunkBytes();
    if (chunkBytes == 0)
    {
        return false;
    }

    uint64_t nextOffset = _textStreamStartOffset;
    bool scrollToEnd    = false;
    if (backward)
    {
        if (_textStreamStartOffset <= _textStreamSkipBytes)
        {
            return false;
        }

        const uint64_t delta = std::min<uint64_t>(_textStreamStartOffset - _textStreamSkipBytes, chunkBytes);
        nextOffset           = _textStreamStartOffset - delta;
        scrollToEnd          = true;
    }
    else
    {
        if (_textStreamEndOffset <= _textStreamStartOffset || _textStreamEndOffset >= _fileSize)
        {
            return false;
        }

        nextOffset  = _textStreamEndOffset;
        scrollToEnd = false;
    }

    nextOffset = AlignTextStreamOffset(nextOffset);
    if (nextOffset == _textStreamStartOffset)
    {
        return false;
    }

    const HRESULT hr = LoadTextToEdit(hwnd, nextOffset, scrollToEnd);
    if (FAILED(hr))
    {
        return false;
    }

    UpdateSearchHighlights();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }

    return true;
}
