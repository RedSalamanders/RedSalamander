#include "ViewerText.h"

#include <algorithm>
#include <chrono>
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
constexpr float kMonoFontSizeDip                    = 10.0f * 96.0f / 72.0f;
constexpr uint32_t kDiffViewportHydrationMarginRows = 32u;

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

D2D1_COLOR_F ColorFFromArgb(uint32_t argb) noexcept
{
    return ColorFFromColorRef(ColorRefFromArgb(argb), AlphaFromArgb(argb));
}

[[maybe_unused]] uint32_t ArgbFromColorRef(COLORREF rgb, uint8_t alpha = 0xFFu) noexcept
{
    const uint32_t r = static_cast<uint32_t>(GetRValue(rgb));
    const uint32_t g = static_cast<uint32_t>(GetGValue(rgb));
    const uint32_t b = static_cast<uint32_t>(GetBValue(rgb));
    return (static_cast<uint32_t>(alpha) << 24) | (r << 16) | (g << 8) | b;
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
    static_cast<void>(EnsureVisibleDiffViewportHydrated(hwnd));
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
        if (EnsureVisibleDiffViewportHydrated(hwnd))
        {
            UpdateTextViewScrollBars(hwnd);
            InvalidateRect(hwnd, nullptr, TRUE);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
            }
            return 0;
        }

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
    static_cast<void>(EnsureVisibleDiffViewportHydrated(hwnd));
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
    RefreshTextHorizontalViewport(hwnd);
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
        static_cast<void>(EnsureVisibleDiffViewportHydrated(hwnd));
        UpdateTextViewScrollBars(hwnd);
        InvalidateRect(hwnd, nullptr, TRUE);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
        }
    }
    else if (EnsureVisibleDiffViewportHydrated(hwnd))
    {
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

std::optional<ViewerText::TextViewHitTestResult> ViewerText::HitTestTextView(HWND hwnd, POINT pt) const noexcept
{
    if (! hwnd || _textVisualLineStarts.empty() || _textVisualLineLogical.empty() || _textVisualLineLayouts.empty())
    {
        return std::nullopt;
    }

    const UINT dpi        = GetDpiForWindow(hwnd);
    const float xDip      = DipsFromPixels(pt.x, dpi);
    const float yDip      = DipsFromPixels(pt.y, dpi);
    const float marginDip = RoundDipToDevicePixels(6.0f, dpi);
    const float charW     = (_textCharWidthDip > 0.0f) ? _textCharWidthDip : 8.0f;
    const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
    if (lineH <= 0.0f)
    {
        return std::nullopt;
    }

    float textStartX = marginDip;
    if (ShowTextLineNumbersInCurrentPresentation() && charW > 0.0f)
    {
        const size_t digits      = LineNumberDigits(_textLineStarts.size());
        const size_t gutterChars = digits + 2u;
        textStartX               = marginDip + static_cast<float>(gutterChars) * charW;
    }

    const float relY      = std::max(0.0f, yDip - marginDip);
    const uint64_t row    = static_cast<uint64_t>(std::floor(relY / lineH));
    const uint64_t visual = std::min<uint64_t>(static_cast<uint64_t>(_textTopVisualLine) + row, static_cast<uint64_t>(_textVisualLineStarts.size() - 1u));
    const TextVisualLineLayoutEntry& layout = _textVisualLineLayouts[static_cast<size_t>(visual)];
    if (layout.logicalLine >= _textLineEnds.size())
    {
        return std::nullopt;
    }

    size_t index = 0u;
    if (layout.splitPanes && charW > 0.0f)
    {
        const float leftPaneWidthDip  = static_cast<float>(layout.leftPaneColumns) * charW;
        const float separatorWidthDip = static_cast<float>(layout.separatorColumns) * charW;
        const float separatorLeft     = textStartX + leftPaneWidthDip;
        const float separatorRight    = separatorLeft + separatorWidthDip;

        if (xDip < separatorLeft)
        {
            const float relX      = std::max(0.0f, xDip - textStartX);
            const uint32_t col    = static_cast<uint32_t>(std::floor(relX / charW));
            const uint32_t segLen = layout.leftEndIndex >= layout.leftStartIndex ? (layout.leftEndIndex - layout.leftStartIndex) : 0u;
            const uint32_t idx32  = layout.leftStartIndex + std::min<uint32_t>(col, segLen);
            index                 = std::min<size_t>(static_cast<size_t>(idx32), _textBuffer.size());
        }
        else if (xDip < separatorRight)
        {
            index = std::min<size_t>(static_cast<size_t>(layout.separatorStartIndex), _textBuffer.size());
        }
        else
        {
            const float relX      = std::max(0.0f, xDip - separatorRight);
            const uint32_t col    = static_cast<uint32_t>(std::floor(relX / charW));
            const uint32_t segLen = layout.rightEndIndex >= layout.rightStartIndex ? (layout.rightEndIndex - layout.rightStartIndex) : 0u;
            const uint32_t idx32  = layout.rightStartIndex + std::min<uint32_t>(col, segLen);
            index                 = std::min<size_t>(static_cast<size_t>(idx32), _textBuffer.size());
        }
    }
    else
    {
        const float relX   = std::max(0.0f, xDip - textStartX);
        const uint32_t col = charW <= 0.0f ? 0u : static_cast<uint32_t>(std::floor(relX / charW));

        uint32_t segmentStart = layout.segmentStartIndex;
        uint32_t segmentEnd   = layout.segmentEndIndex;
        if (! _wrap && segmentEnd >= segmentStart && _textLeftColumn != 0u)
        {
            const uint32_t skip = std::min<uint32_t>(_textLeftColumn, segmentEnd - segmentStart);
            segmentStart += skip;
        }

        const uint32_t segLen     = segmentEnd >= segmentStart ? (segmentEnd - segmentStart) : 0u;
        const uint32_t colClamped = std::min<uint32_t>(col, segLen);
        const uint32_t idx32      = segmentStart + colClamped;
        index                     = std::min<size_t>(static_cast<size_t>(idx32), _textBuffer.size());
    }

    return TextViewHitTestResult{index, layout.logicalLine};
}

bool ViewerText::IsClickableHiddenDiffBannerLogicalLine(uint32_t logicalLine) const noexcept
{
    using DiffSemanticRowKind = ViewerText::DiffTextVariant::SemanticRowKind;

    if (! HasParsedDiffPresentation() || _config.diffContextMode != DiffContextMode::HunksOnly)
    {
        return false;
    }

    const auto* variant = CurrentDiffVariant();
    if (! variant || logicalLine >= variant->logicalRowStyles.size() || logicalLine >= variant->logicalRowRenderInfo.size())
    {
        return false;
    }

    return variant->logicalRowStyles[logicalLine].fullRow == DiffSemanticRowKind::HiddenContextBanner &&
           variant->logicalRowRenderInfo[logicalLine].clickableBanner;
}

bool ViewerText::DebugClickTextLogicalLine(HWND hwnd, uint32_t logicalLine) noexcept
{
    if (! hwnd || _viewMode != ViewMode::Text || logicalLine >= _textLineStarts.size() || _textVisualLineLogical.empty())
    {
        return false;
    }

    const auto findVisibleVisualLine = [&](uint32_t targetLogicalLine) noexcept -> std::optional<size_t>
    {
        const auto it =
            std::lower_bound(_textVisualLineLogical.begin(), _textVisualLineLogical.end(), targetLogicalLine, [](uint32_t left, uint32_t right) noexcept {
            return left < right;
        });
        if (it == _textVisualLineLogical.end() || *it != targetLogicalLine)
        {
            return std::nullopt;
        }

        const size_t visualLine = static_cast<size_t>(std::distance(_textVisualLineLogical.begin(), it));
        if (visualLine < _textTopVisualLine)
        {
            return std::nullopt;
        }

        RECT client{};
        GetClientRect(hwnd, &client);
        const UINT dpi          = GetDpiForWindow(hwnd);
        const float heightDip   = DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi);
        const float marginDip   = RoundDipToDevicePixels(6.0f, dpi);
        const float lineH       = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
        const float rowIndexDip = static_cast<float>(visualLine - static_cast<size_t>(_textTopVisualLine));
        const float yDip        = marginDip + rowIndexDip * lineH + lineH * 0.5f;
        if (yDip < 0.0f || yDip >= heightDip)
        {
            return std::nullopt;
        }

        return visualLine;
    };

    std::optional<size_t> visualLine = findVisibleVisualLine(logicalLine);
    if (! visualLine.has_value())
    {
        ScrollTextViewportToLogicalLine(hwnd, logicalLine);
        visualLine = findVisibleVisualLine(logicalLine);
    }

    if (! visualLine.has_value())
    {
        return false;
    }

    RECT client{};
    GetClientRect(hwnd, &client);
    const UINT dpi        = GetDpiForWindow(hwnd);
    const float widthDip  = DipsFromPixels(static_cast<int>(client.right - client.left), dpi);
    const float heightDip = DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi);
    const float marginDip = RoundDipToDevicePixels(6.0f, dpi);
    const float charW     = (_textCharWidthDip > 0.0f) ? _textCharWidthDip : 8.0f;
    const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;

    float textStartX = marginDip;
    if (ShowTextLineNumbersInCurrentPresentation() && charW > 0.0f)
    {
        const size_t digits      = LineNumberDigits(_textLineStarts.size());
        const size_t gutterChars = digits + 2u;
        textStartX += static_cast<float>(gutterChars) * charW;
    }

    const float rowIndexDip = static_cast<float>(visualLine.value() - static_cast<size_t>(_textTopVisualLine));
    const float xDip        = std::min(widthDip - marginDip, textStartX + std::max(charW, 12.0f));
    const float yDip        = marginDip + rowIndexDip * lineH + lineH * 0.5f;
    if (xDip < 0.0f || yDip < 0.0f || xDip >= widthDip || yDip >= heightDip)
    {
        return false;
    }

    const auto toPixels = [dpi](float dip) noexcept { return static_cast<LONG>(std::lround(dip * static_cast<float>(dpi) / 96.0f)); };

    const POINT pt{toPixels(xDip), toPixels(yDip)};
    static_cast<void>(OnTextViewLButtonDown(hwnd, pt));
    static_cast<void>(OnTextViewLButtonUp(hwnd));
    return true;
}

LRESULT ViewerText::OnTextViewLButtonDown(HWND hwnd, POINT pt) noexcept
{
    SetFocus(hwnd);

    const auto hit = HitTestTextView(hwnd, pt);
    if (hit.has_value() && IsClickableHiddenDiffBannerLogicalLine(hit->logicalLine))
    {
        if (const HWND root = GetAncestor(hwnd, GA_ROOT))
        {
            SetDiffContextMode(root, DiffContextMode::FullFileWhenAvailable);
        }
        return 0;
    }

    SetCapture(hwnd);
    const size_t index = hit.has_value() ? hit->bufferIndex : 0u;

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

    const auto hit = HitTestTextView(hwnd, pt);
    if (! hit.has_value())
    {
        return 0;
    }

    _textSelActive  = hit->bufferIndex;
    _textCaretIndex = _textSelActive;

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

LRESULT ViewerText::OnTextViewSetCursor(HWND hwnd, LPARAM lParam) noexcept
{
    if (! hwnd || LOWORD(lParam) != HTCLIENT)
    {
        return FALSE;
    }

    POINT pt{};
    if (GetCursorPos(&pt) == 0 || ScreenToClient(hwnd, &pt) == 0)
    {
        return FALSE;
    }

    const auto hit = HitTestTextView(hwnd, pt);
    if (! hit.has_value() || ! IsClickableHiddenDiffBannerLogicalLine(hit->logicalLine))
    {
        return FALSE;
    }

    SetCursor(LoadCursorW(nullptr, IDC_HAND));
    return TRUE;
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
        const auto paintStartedAt = std::chrono::steady_clock::now();
        const UINT dpi            = GetDpiForWindow(hwnd);
        const COLORREF bg         = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
        const COLORREF fg         = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

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
            const float marginDip = RoundDipToDevicePixels(6.0f, dpi);
            const float charW     = (_textCharWidthDip > 0.0f) ? _textCharWidthDip : 8.0f;
            const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
#ifdef _DEBUG
            _debugTextRenderCount += 1u;
            _debugTextVisibleRowCount        = 0u;
            _debugTextVisibleStyledRowCount  = 0u;
            _debugTextVisibleContextRowCount = 0u;
            _debugTextVisibleAddedRowCount   = 0u;
            _debugTextVisibleRemovedRowCount = 0u;
            _debugTextVisibleHeaderRowCount  = 0u;
            _debugTextVisibleBannerRowCount  = 0u;
            _debugTextVisibleGapHatchCount   = 0u;
            _debugTextVisibleSplitRowCount   = 0u;
#endif

            float gutterWidthDip = 0.0f;
            float textStartX     = marginDip;
            if (ShowTextLineNumbersInCurrentPresentation() && charW > 0.0f)
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

            const bool showLineNumbers                           = ShowTextLineNumbersInCurrentPresentation() && gutterWidthDip > 0.0f;
            const uint8_t lineNumberAlpha                        = (_hasTheme && _theme.darkMode) ? 160u : 140u;
            const COLORREF lineNumberFg                          = BlendColor(bg, fg, lineNumberAlpha);
            const ViewerText::DiffTextVariant* activeDiffVariant = HasParsedDiffPresentation() ? CurrentDiffVariant() : nullptr;
            using DiffSemanticRowKind                            = ViewerText::DiffTextVariant::SemanticRowKind;
            const auto findPlaceholderBandPlacement = [&](uint32_t logicalLine) noexcept -> std::optional<ViewerText::DiffTextVariant::PlaceholderBandPlacement>
            {
                if (! activeDiffVariant || activeDiffVariant->placeholderBands.empty())
                {
                    return std::nullopt;
                }

                const auto it = std::lower_bound(activeDiffVariant->placeholderBands.begin(),
                                                 activeDiffVariant->placeholderBands.end(),
                                                 logicalLine,
                                                 [](const ViewerText::DiffTextVariant::PlaceholderBandEntry& entry, uint32_t value) noexcept
                { return entry.logicalLine < value; });
                if (it == activeDiffVariant->placeholderBands.end() || it->logicalLine != logicalLine)
                {
                    return std::nullopt;
                }

                return it->placement;
            };
            const D2D1_COLOR_F baseTextBg    = ColorFFromColorRef(bg);
            const D2D1_COLOR_F diffAddedBg   = _hasTheme ? ColorFFromArgb(_theme.diffAddedBackgroundArgb) : ColorFFromColorRef(RGB(46, 160, 67), 0.16f);
            const D2D1_COLOR_F diffRemovedBg = _hasTheme ? ColorFFromArgb(_theme.diffRemovedBackgroundArgb) : ColorFFromColorRef(RGB(204, 51, 51), 0.16f);
            const D2D1_COLOR_F diffHeaderBg =
                _hasTheme ? ColorFFromArgb(_theme.diffHeaderBackgroundArgb) : ColorFFromColorRef(accent, (_hasTheme && _theme.darkMode) ? 0.16f : 0.10f);
            const D2D1_COLOR_F diffBannerBg =
                _hasTheme ? ColorFFromArgb(_theme.diffBannerBackgroundArgb) : ColorFFromColorRef(accent, (_hasTheme && _theme.darkMode) ? 0.22f : 0.16f);
            const D2D1_COLOR_F diffPlaceholderBg =
                _hasTheme ? ColorFFromArgb(_theme.diffPlaceholderBackgroundArgb) : ColorFFromColorRef(accent, (_hasTheme && _theme.darkMode) ? 0.18f : 0.12f);
            const D2D1_COLOR_F diffDividerBg = _hasTheme ? ColorFFromArgb(_theme.diffDividerArgb)
                                                         : ColorFFromColorRef(BlendColor(bg, accent, (_hasTheme && _theme.darkMode) ? 28u : 18u), 0.80f);
            const D2D1_COLOR_F diffStructuralBorderBg =
                ColorFFromColorRef(BlendColor(bg, fg, (_hasTheme && _theme.darkMode) ? 72u : 48u), (_hasTheme && _theme.darkMode) ? 0.92f : 0.82f);
            const D2D1_COLOR_F diffActiveHunkBorderBg = ColorFFromColorRef(BlendColor(bg, accent, (_hasTheme && _theme.darkMode) ? 168u : 120u), 0.96f);
            const COLORREF diffMarkerColorRef         = BlendColor(bg, fg, (_hasTheme && _theme.darkMode) ? 72u : 52u);
            const COLORREF diffGapHatchColorRef       = BlendColor(bg, fg, (_hasTheme && _theme.darkMode) ? 48u : 34u);
            const D2D1_COLOR_F diffMarkerFg           = ColorFFromColorRef(diffMarkerColorRef, (_hasTheme && _theme.darkMode) ? 0.88f : 0.78f);
            const D2D1_COLOR_F diffGapHatchFg         = ColorFFromColorRef(diffGapHatchColorRef, (_hasTheme && _theme.darkMode) ? 0.34f : 0.22f);
#ifdef _DEBUG
            _debugDiffMarkerArgb                = ArgbFromColorRef(diffMarkerColorRef, (_hasTheme && _theme.darkMode) ? 224u : 200u);
            _debugDiffGapHatchArgb              = ArgbFromColorRef(diffGapHatchColorRef, (_hasTheme && _theme.darkMode) ? 88u : 56u);
            _debugDiffContextUsesBaseBackground = activeDiffVariant != nullptr;
#endif
            const size_t activeHunkIndex      = activeDiffVariant ? CurrentDiffHunkIndex() : 0u;
            const auto isActiveHunkHeaderLine = [&](uint32_t logicalLine) noexcept
            {
                return activeDiffVariant && activeHunkIndex < activeDiffVariant->hunkNavigation.size() &&
                       activeDiffVariant->hunkNavigation[activeHunkIndex].startLogicalLine == logicalLine;
            };
            const auto semanticColorFor = [&](DiffSemanticRowKind kind) noexcept -> std::optional<D2D1_COLOR_F>
            {
                switch (kind)
                {
                    case DiffSemanticRowKind::Context: return std::nullopt;
                    case DiffSemanticRowKind::Added: return diffAddedBg;
                    case DiffSemanticRowKind::Removed: return diffRemovedBg;
                    case DiffSemanticRowKind::FileHeader: return diffHeaderBg;
                    case DiffSemanticRowKind::HunkHeader: return diffBannerBg;
                    case DiffSemanticRowKind::HiddenContextBanner: return diffBannerBg;
                    case DiffSemanticRowKind::Placeholder: return diffPlaceholderBg;
                    case DiffSemanticRowKind::None: break;
                }
                return std::nullopt;
            };
            size_t visibleRowsDrawn = 0u;

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

                    const TextVisualLineLayoutEntry* visualLayout =
                        static_cast<size_t>(visual) < _textVisualLineLayouts.size() ? &_textVisualLineLayouts[static_cast<size_t>(visual)] : nullptr;
                    const uint32_t segStartRaw = visualLayout ? visualLayout->segmentStartIndex : _textVisualLineStarts[static_cast<size_t>(visual)];
                    uint32_t segStart          = segStartRaw;
                    uint32_t segEnd            = visualLayout ? visualLayout->segmentEndIndex : _textLineEnds[logical];
                    if (! visualLayout && (visual + 1) < totalVisual && _textVisualLineLogical[static_cast<size_t>(visual + 1)] == logical)
                    {
                        segEnd = _textVisualLineStarts[static_cast<size_t>(visual + 1)];
                    }

                    if (! _wrap && (! visualLayout || ! visualLayout->splitPanes) && segEnd >= segStart && _textLeftColumn != 0)
                    {
                        const uint32_t skip = std::min<uint32_t>(_textLeftColumn, segEnd - segStart);
                        segStart += skip;
                    }

                    const size_t startIndex = std::min<size_t>(static_cast<size_t>(segStart), _textBuffer.size());
                    const size_t endIndex   = std::min<size_t>(static_cast<size_t>(segEnd), _textBuffer.size());

                    const float x                       = textStartX;
                    const float y                       = RoundDipToDevicePixels(marginDip + static_cast<float>(row) * lineH, dpi);
                    const float lineBottom              = RoundDipToDevicePixels(y + lineH, dpi);
                    const D2D1_RECT_F lineRc            = D2D1::RectF(x, y, std::max(x, widthDip - marginDip), lineBottom);
                    const bool isFirstSegmentForLogical = (visual == 0u) || (_textVisualLineLogical[static_cast<size_t>(visual - 1u)] != logical);
                    const size_t logicalStartIndex      = static_cast<size_t>(_textLineStarts[logical]);
                    const size_t logicalEndIndex        = static_cast<size_t>(_textLineEnds[logical]);
                    const std::wstring_view logicalLine = logicalEndIndex >= logicalStartIndex && logicalEndIndex <= _textBuffer.size()
                                                              ? std::wstring_view(_textBuffer.data() + logicalStartIndex, logicalEndIndex - logicalStartIndex)
                                                              : std::wstring_view{};
                    const bool splitPaneRow             = visualLayout && visualLayout->splitPanes && _documentKind == DocumentKind::Diff &&
                                                          _diffPresentation == DiffPresentationMode::SideBySide;
                    const size_t separatorColumn   = splitPaneRow ? static_cast<size_t>(visualLayout->separatorStartIndex >= _textLineStarts[logical]
                                                                                            ? (visualLayout->separatorStartIndex - _textLineStarts[logical])
                                                                                            : 0u)
                                                                  : logicalLine.find(L" | ");
                    const bool hasSeparator        = splitPaneRow || separatorColumn != std::wstring_view::npos;
                    const uint32_t lineStartColumn = segStart >= _textLineStarts[logical] ? (segStart - _textLineStarts[logical]) : 0u;
                    const uint32_t lineEndColumn   = segEnd >= _textLineStarts[logical] ? (segEnd - _textLineStarts[logical]) : lineStartColumn;
                    const uint32_t separatorStart  = static_cast<uint32_t>(separatorColumn);
                    const uint32_t separatorEnd = splitPaneRow
                                                      ? static_cast<uint32_t>(visualLayout->separatorEndIndex >= _textLineStarts[logical]
                                                                                  ? (visualLayout->separatorEndIndex - _textLineStarts[logical])
                                                                                  : separatorStart)
                                                      : static_cast<uint32_t>(hasSeparator ? std::min<size_t>(logicalLine.size(), separatorColumn + 3u) : 0u);
                    const uint32_t rightStartColumn = separatorEnd;
                    const size_t leftStartIndex     = splitPaneRow ? std::min<size_t>(visualLayout->leftStartIndex, _textBuffer.size()) : startIndex;
                    const size_t leftEndIndex       = splitPaneRow ? std::min<size_t>(visualLayout->leftEndIndex, _textBuffer.size()) : endIndex;
                    const size_t rightStartIndex    = splitPaneRow ? std::min<size_t>(visualLayout->rightStartIndex, _textBuffer.size()) : endIndex;
                    const size_t rightEndIndex      = splitPaneRow ? std::min<size_t>(visualLayout->rightEndIndex, _textBuffer.size()) : endIndex;
                    const auto* rowStyle =
                        (activeDiffVariant && logical < activeDiffVariant->logicalRowStyles.size()) ? &activeDiffVariant->logicalRowStyles[logical] : nullptr;
                    const auto* rowRender       = (activeDiffVariant && logical < activeDiffVariant->logicalRowRenderInfo.size())
                                                      ? &activeDiffVariant->logicalRowRenderInfo[logical]
                                                      : nullptr;
                    const D2D1_RECT_F fullRowRc = D2D1::RectF(0.0f, y, widthDip, lineBottom);
                    const float separatorWidthDip =
                        splitPaneRow ? static_cast<float>(visualLayout->separatorColumns) * charW : static_cast<float>(separatorEnd - separatorStart) * charW;
                    const float leftPaneWidthDip =
                        splitPaneRow ? static_cast<float>(visualLayout->leftPaneColumns) * charW : static_cast<float>(separatorStart) * charW;
                    const float leftPaneRight =
                        splitPaneRow ? std::min(lineRc.right, x + leftPaneWidthDip) : std::min(lineRc.right, x + static_cast<float>(separatorStart) * charW);
                    const float separatorLeft    = leftPaneRight;
                    const float separatorRight   = splitPaneRow
                                                       ? std::min(lineRc.right, separatorLeft + separatorWidthDip)
                                                       : std::min(lineRc.right, separatorLeft + static_cast<float>(separatorEnd - separatorStart) * charW);
                    const float rightPaneLeft    = splitPaneRow ? separatorRight : x + static_cast<float>(rightStartColumn - lineStartColumn) * charW;
                    const D2D1_RECT_F leftPaneRc = D2D1::RectF(x, y, std::max(x, leftPaneRight), lineBottom);
                    const D2D1_RECT_F separatorRc =
                        D2D1::RectF(std::max(x, separatorLeft), y, std::max(std::max(x, separatorLeft), separatorRight), lineBottom);
                    const D2D1_RECT_F rightPaneRc = D2D1::RectF(std::max(x, rightPaneLeft), y, std::max(std::max(x, rightPaneLeft), lineRc.right), lineBottom);
                    const auto fillVisibleColumns = [&](uint32_t columnStart, uint32_t columnEndExclusive, const D2D1_COLOR_F& fillColor) noexcept
                    {
                        if (columnEndExclusive <= columnStart || charW <= 0.0f)
                        {
                            return;
                        }

                        const uint32_t visibleStart = std::max(lineStartColumn, columnStart);
                        const uint32_t visibleEnd   = std::min(lineEndColumn, columnEndExclusive);
                        if (visibleEnd <= visibleStart)
                        {
                            return;
                        }

                        const float bandX = x + static_cast<float>(visibleStart - lineStartColumn) * charW;
                        const float bandW = static_cast<float>(visibleEnd - visibleStart) * charW;
                        _textViewBrush->SetColor(fillColor);
                        _textViewTarget->FillRectangle(D2D1::RectF(bandX, y, bandX + bandW, lineBottom), _textViewBrush.get());
                        _textViewBrush->SetColor(ColorFFromColorRef(fg));
                    };
                    const auto fillPaneRect = [&](const D2D1_RECT_F& rc, const D2D1_COLOR_F& fillColor) noexcept
                    {
                        if (rc.right <= rc.left)
                        {
                            return;
                        }

                        _textViewBrush->SetColor(fillColor);
                        _textViewTarget->FillRectangle(rc, _textViewBrush.get());
                        _textViewBrush->SetColor(ColorFFromColorRef(fg));
                    };
                    const auto fillTextArea = [&](const D2D1_COLOR_F& fillColor) noexcept
                    {
                        _textViewBrush->SetColor(fillColor);
                        _textViewTarget->FillRectangle(lineRc, _textViewBrush.get());
                        _textViewBrush->SetColor(ColorFFromColorRef(fg));
                    };
                    const auto fillFullRow = [&](const D2D1_COLOR_F& fillColor) noexcept
                    {
                        _textViewBrush->SetColor(fillColor);
                        _textViewTarget->FillRectangle(fullRowRc, _textViewBrush.get());
                        _textViewBrush->SetColor(ColorFFromColorRef(fg));
                    };
                    const auto drawRowEdges = [&](const D2D1_COLOR_F& edgeColor, float thicknessDip) noexcept
                    {
                        _textViewBrush->SetColor(edgeColor);
                        _textViewTarget->DrawLine(D2D1::Point2F(fullRowRc.left, y), D2D1::Point2F(fullRowRc.right, y), _textViewBrush.get(), thicknessDip);
                        _textViewTarget->DrawLine(
                            D2D1::Point2F(fullRowRc.left, lineBottom), D2D1::Point2F(fullRowRc.right, lineBottom), _textViewBrush.get(), thicknessDip);
                        _textViewBrush->SetColor(ColorFFromColorRef(fg));
                    };
                    const auto drawLeadingStripe = [&](const D2D1_COLOR_F& fillColor, float widthDip) noexcept
                    {
                        const float stripeLeft  = showLineNumbers ? std::max(marginDip, textStartX - RoundDipToDevicePixels(4.0f, dpi)) : marginDip;
                        const float stripeWidth = std::max(DevicePixelDip(dpi), widthDip);
                        _textViewBrush->SetColor(fillColor);
                        _textViewTarget->FillRectangle(D2D1::RectF(stripeLeft, y, std::min(widthDip, stripeLeft + stripeWidth), lineBottom),
                                                       _textViewBrush.get());
                        _textViewBrush->SetColor(ColorFFromColorRef(fg));
                    };
                    const auto drawBannerChip =
                        [&](const D2D1_COLOR_F& chipFill, const D2D1_COLOR_F& chipBorder, bool emphasizeBorder, bool centerChip = false) noexcept
                    {
                        if (! isFirstSegmentForLogical || logicalLine.empty())
                        {
                            return;
                        }

                        const float chipPadX     = RoundDipToDevicePixels(centerChip ? 10.0f : 8.0f, dpi);
                        const float chipPadY     = RoundDipToDevicePixels(centerChip ? 3.0f : 2.0f, dpi);
                        const float chipMinLeft  = std::max(marginDip, textStartX + RoundDipToDevicePixels(centerChip ? 6.0f : -2.0f, dpi));
                        const float maxChipWidth = std::max(0.0f, widthDip - chipMinLeft - marginDip);
                        if (maxChipWidth <= 0.0f)
                        {
                            return;
                        }

                        const float chipTextWidth =
                            std::min(maxChipWidth, static_cast<float>(std::min<size_t>(logicalLine.size(), 160u)) * charW + chipPadX * 2.0f);
                        const float chipLeft = centerChip ? std::max(chipMinLeft, x + std::max(0.0f, (lineRc.right - x - chipTextWidth) * 0.5f)) : chipMinLeft;
                        const float chipTop  = y + chipPadY;
                        const float chipBottom = std::max(chipTop + DevicePixelDip(dpi), lineBottom - chipPadY);
                        const float chipRadius = RoundDipToDevicePixels(emphasizeBorder ? 6.0f : 5.0f, dpi);
                        const D2D1_ROUNDED_RECT chipRect =
                            D2D1::RoundedRect(D2D1::RectF(chipLeft, chipTop, chipLeft + chipTextWidth, chipBottom), chipRadius, chipRadius);

                        _textViewBrush->SetColor(chipFill);
                        _textViewTarget->FillRoundedRectangle(chipRect, _textViewBrush.get());
                        _textViewBrush->SetColor(chipBorder);
                        _textViewTarget->DrawRoundedRectangle(chipRect, _textViewBrush.get(), emphasizeBorder ? 2.0f : 1.0f);
                        _textViewBrush->SetColor(ColorFFromColorRef(fg));
                    };
                    const auto drawGapHatch = [&](const D2D1_RECT_F& rc) noexcept
                    {
                        if (rc.right <= rc.left || rc.bottom <= rc.top)
                        {
                            return;
                        }

                        fillPaneRect(rc, baseTextBg);

                        const float stepDip      = std::max(DevicePixelDip(dpi) * 5.0f, RoundDipToDevicePixels(7.0f, dpi));
                        const float thicknessDip = std::max(DevicePixelDip(dpi), 1.0f);
                        const float diagonalDip  = rc.bottom - rc.top;
                        _textViewBrush->SetColor(diffGapHatchFg);
                        for (float startX = rc.left - diagonalDip; startX < rc.right + diagonalDip; startX += stepDip)
                        {
                            _textViewTarget->DrawLine(
                                D2D1::Point2F(startX, rc.bottom), D2D1::Point2F(startX + diagonalDip, rc.top), _textViewBrush.get(), thicknessDip);
                        }
                        _textViewBrush->SetColor(ColorFFromColorRef(fg));
#ifdef _DEBUG
                        _debugTextVisibleGapHatchCount += 1u;
#endif
                    };
                    const auto fillSeparatorBand = [&]() noexcept
                    {
                        if (! hasSeparator)
                        {
                            return;
                        }

                        if (splitPaneRow)
                        {
                            fillPaneRect(separatorRc, baseTextBg);
                            return;
                        }

                        fillVisibleColumns(separatorStart, separatorEnd, diffDividerBg);
                    };
                    const auto drawSeparatorEdges = [&]() noexcept
                    {
                        if (! hasSeparator || charW <= 0.0f)
                        {
                            return;
                        }

                        if (splitPaneRow)
                        {
                            const float centerX              = std::floor((separatorRc.left + separatorRc.right) * 0.5f);
                            const D2D1_COLOR_F softDividerBg = ColorFFromColorRef(BlendColor(bg, fg, (_hasTheme && _theme.darkMode) ? 24u : 18u),
                                                                                  (_hasTheme && _theme.darkMode) ? 0.24f : 0.18f);
                            _textViewBrush->SetColor(softDividerBg);
                            _textViewTarget->FillRectangle(separatorRc, _textViewBrush.get());
                            _textViewBrush->SetColor(diffStructuralBorderBg);
                            _textViewTarget->DrawLine(D2D1::Point2F(centerX, y), D2D1::Point2F(centerX, lineBottom), _textViewBrush.get(), 1.0f);
                            _textViewBrush->SetColor(ColorFFromColorRef(fg));
                            return;
                        }

                        const uint32_t visibleStartColumn = std::max(lineStartColumn, separatorStart);
                        const uint32_t visibleEndColumn   = std::min(lineEndColumn, separatorEnd);
                        if (visibleEndColumn <= visibleStartColumn)
                        {
                            return;
                        }

                        const float leftX  = x + static_cast<float>(visibleStartColumn - lineStartColumn) * charW;
                        const float rightX = x + static_cast<float>(visibleEndColumn - lineStartColumn) * charW;
                        _textViewBrush->SetColor(diffStructuralBorderBg);
                        _textViewTarget->DrawLine(D2D1::Point2F(leftX, y), D2D1::Point2F(leftX, lineBottom), _textViewBrush.get(), 1.0f);
                        _textViewTarget->DrawLine(D2D1::Point2F(rightX, y), D2D1::Point2F(rightX, lineBottom), _textViewBrush.get(), 1.0f);
                        _textViewBrush->SetColor(ColorFFromColorRef(fg));
                    };
                    const auto drawDimmedMarker =
                        [&](size_t markerIndex, size_t visibleStart, size_t visibleEnd, float spanLeft, const D2D1_RECT_F& spanRc) noexcept
                    {
                        if (! rowRender || markerIndex >= _textBuffer.size() || visibleEnd <= visibleStart || spanRc.right <= spanRc.left ||
                            markerIndex < visibleStart || markerIndex >= visibleEnd)
                        {
                            return;
                        }

                        const wchar_t marker = _textBuffer[markerIndex];
                        if (marker != L'+' && marker != L'-')
                        {
                            return;
                        }

                        const float markerX = spanLeft + static_cast<float>(markerIndex - visibleStart) * charW;
                        if (markerX >= spanRc.right)
                        {
                            return;
                        }

                        const D2D1_RECT_F markerRc = D2D1::RectF(markerX, y, std::min(spanRc.right, markerX + charW), lineBottom);
                        _textViewBrush->SetColor(diffMarkerFg);
                        _textViewTarget->DrawTextW(&marker, 1u, _textViewFormat.get(), markerRc, _textViewBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        _textViewBrush->SetColor(ColorFFromColorRef(fg));
                    };

#ifdef _DEBUG
                    if (isFirstSegmentForLogical && rowStyle)
                    {
                        if (splitPaneRow)
                        {
                            _debugTextVisibleSplitRowCount += 1u;
                        }

                        const auto countVisibleKind = [&](DiffSemanticRowKind kind) noexcept
                        {
                            switch (kind)
                            {
                                case DiffSemanticRowKind::Context: _debugTextVisibleContextRowCount += 1u; break;
                                case DiffSemanticRowKind::Added: _debugTextVisibleAddedRowCount += 1u; break;
                                case DiffSemanticRowKind::Removed: _debugTextVisibleRemovedRowCount += 1u; break;
                                case DiffSemanticRowKind::FileHeader: _debugTextVisibleHeaderRowCount += 1u; break;
                                case DiffSemanticRowKind::HunkHeader:
                                case DiffSemanticRowKind::HiddenContextBanner: _debugTextVisibleBannerRowCount += 1u; break;
                                case DiffSemanticRowKind::Placeholder:
                                case DiffSemanticRowKind::None: break;
                            }
                        };

                        bool styledRow = false;
                        if (rowStyle->fullRow != DiffSemanticRowKind::None)
                        {
                            styledRow = true;
                            countVisibleKind(rowStyle->fullRow);
                        }
                        else
                        {
                            if (rowStyle->leftPane != DiffSemanticRowKind::None)
                            {
                                styledRow = true;
                                countVisibleKind(rowStyle->leftPane);
                            }
                            if (rowStyle->rightPane != DiffSemanticRowKind::None)
                            {
                                styledRow = true;
                                countVisibleKind(rowStyle->rightPane);
                            }
                        }

                        if (styledRow)
                        {
                            _debugTextVisibleStyledRowCount += 1u;
                        }
                    }
#endif

                    if (rowStyle)
                    {
                        if (rowStyle->fullRow != DiffSemanticRowKind::None)
                        {
                            switch (rowStyle->fullRow)
                            {
                                case DiffSemanticRowKind::Context: break;
                                case DiffSemanticRowKind::Added: fillTextArea(diffAddedBg); break;
                                case DiffSemanticRowKind::Removed: fillTextArea(diffRemovedBg); break;
                                case DiffSemanticRowKind::FileHeader:
                                    fillFullRow(diffHeaderBg);
                                    drawRowEdges(diffStructuralBorderBg, 1.0f);
                                    drawLeadingStripe(diffDividerBg, RoundDipToDevicePixels(3.0f, dpi));
                                    break;
                                case DiffSemanticRowKind::HunkHeader:
                                {
                                    const bool activeHunkHeader = isActiveHunkHeaderLine(logical);
                                    fillFullRow(baseTextBg);
                                    drawRowEdges(activeHunkHeader ? diffActiveHunkBorderBg : diffStructuralBorderBg, activeHunkHeader ? 2.0f : 1.0f);
                                    drawLeadingStripe(activeHunkHeader ? diffActiveHunkBorderBg : diffDividerBg,
                                                      RoundDipToDevicePixels(activeHunkHeader ? 4.0f : 3.0f, dpi));
                                    drawBannerChip(diffBannerBg, activeHunkHeader ? diffActiveHunkBorderBg : diffDividerBg, activeHunkHeader);
                                    break;
                                }
                                case DiffSemanticRowKind::HiddenContextBanner:
                                {
                                    const bool clickableBanner = rowRender && rowRender->clickableBanner;
                                    fillFullRow(baseTextBg);
                                    drawRowEdges(diffStructuralBorderBg, 1.0f);
                                    drawBannerChip(diffBannerBg, clickableBanner ? diffActiveHunkBorderBg : diffStructuralBorderBg, clickableBanner, true);
                                    break;
                                }
                                case DiffSemanticRowKind::Placeholder:
                                    drawGapHatch(lineRc);
                                    drawRowEdges(diffStructuralBorderBg, 1.0f);
                                    break;
                                case DiffSemanticRowKind::None: break;
                            }
                        }
                        else if (hasSeparator)
                        {
                            if (splitPaneRow)
                            {
                                fillPaneRect(leftPaneRc, baseTextBg);
                                fillPaneRect(rightPaneRc, baseTextBg);
                            }
                            else
                            {
                                fillVisibleColumns(0u, separatorStart, baseTextBg);
                                fillVisibleColumns(rightStartColumn, static_cast<uint32_t>(logicalLine.size()), baseTextBg);
                            }
                            if (const auto leftFill = semanticColorFor(rowStyle->leftPane))
                            {
                                if (splitPaneRow)
                                {
                                    fillPaneRect(leftPaneRc, *leftFill);
                                }
                                else
                                {
                                    fillVisibleColumns(0u, separatorStart, *leftFill);
                                }
                            }
                            if (const auto rightFill = semanticColorFor(rowStyle->rightPane))
                            {
                                if (splitPaneRow)
                                {
                                    fillPaneRect(rightPaneRc, *rightFill);
                                }
                                else
                                {
                                    fillVisibleColumns(rightStartColumn, static_cast<uint32_t>(logicalLine.size()), *rightFill);
                                }
                            }
                            if (rowRender && rowRender->leftPaneAbsent)
                            {
                                drawGapHatch(leftPaneRc);
                            }
                            if (rowRender && rowRender->rightPaneAbsent)
                            {
                                drawGapHatch(rightPaneRc);
                            }
                        }
                    }

                    if (_documentKind == DocumentKind::Diff && _diffPresentation == DiffPresentationMode::SideBySide && hasSeparator)
                    {
                        fillSeparatorBand();
                        drawSeparatorEdges();
                    }

                    if (const auto placement = findPlaceholderBandPlacement(logical))
                    {
                        if (*placement == ViewerText::DiffTextVariant::PlaceholderBandPlacement::FullRow)
                        {
                            drawGapHatch(lineRc);
                        }
                        else if (! hasSeparator)
                        {
                            drawGapHatch(lineRc);
                        }
                        else if (*placement == ViewerText::DiffTextVariant::PlaceholderBandPlacement::LeftPane)
                        {
                            drawGapHatch(leftPaneRc);
                        }
                        else
                        {
                            drawGapHatch(rightPaneRc);
                        }
                    }

                    if (showLineNumbers)
                    {
                        if (isFirstSegmentForLogical)
                        {
                            const std::wstring lineNumber  = std::to_wstring(static_cast<uint64_t>(logical) + 1u);
                            const float lineNumberRight    = std::max(marginDip, textStartX - charW);
                            const D2D1_RECT_F lineNumberRc = D2D1::RectF(marginDip, y, std::max(marginDip, lineNumberRight), lineBottom);

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

                    const auto drawSearchHighlightsForSpan = [&](size_t visibleStart, size_t visibleEnd, float spanLeft, const D2D1_RECT_F& spanRc) noexcept
                    {
                        if (! hasSearchHighlights || searchLen == 0 || visibleEnd < visibleStart || charW <= 0.0f)
                        {
                            return;
                        }

                        const size_t scanStart = (visibleStart > searchLen) ? (visibleStart - searchLen) : 0u;
                        auto it                = std::lower_bound(_searchMatchStarts.begin(), _searchMatchStarts.end(), scanStart);
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
                            if (hlEnd <= hlStart)
                            {
                                continue;
                            }

                            const size_t colStart  = hlStart - visibleStart;
                            const size_t colLen    = hlEnd - hlStart;
                            const float hlX        = spanLeft + static_cast<float>(colStart) * charW;
                            const float hlW        = static_cast<float>(colLen) * charW;
                            const D2D1_RECT_F hlRc = D2D1::RectF(hlX, y, std::min(spanRc.right, hlX + hlW), lineBottom);

                            _textViewBrush->SetColor(ColorFFromColorRef(searchBg));
                            _textViewTarget->FillRectangle(hlRc, _textViewBrush.get());
                            _textViewBrush->SetColor(ColorFFromColorRef(fg));
                        }
                    };
                    const auto drawSelectionForSpan = [&](size_t visibleStart, size_t visibleEnd, float spanLeft, const D2D1_RECT_F& spanRc) noexcept
                    {
                        if (! hasSelection || visibleEnd < visibleStart || charW <= 0.0f)
                        {
                            return;
                        }

                        const size_t hlStart = std::max(selStartIndex, visibleStart);
                        const size_t hlEnd   = std::min(selEndIndex, visibleEnd);
                        if (hlEnd <= hlStart)
                        {
                            return;
                        }

                        const size_t colStart  = hlStart - visibleStart;
                        const size_t colLength = hlEnd - hlStart;
                        const float hlX        = spanLeft + static_cast<float>(colStart) * charW;
                        const float hlW        = static_cast<float>(colLength) * charW;
                        const D2D1_RECT_F hlRc = D2D1::RectF(hlX, y, std::min(spanRc.right, hlX + hlW), lineBottom);

                        _textViewBrush->SetColor(ColorFFromColorRef(selectionIsSearchMatch ? selectionFocusedBg : selectionBg));
                        _textViewTarget->FillRectangle(hlRc, _textViewBrush.get());
                        _textViewBrush->SetColor(ColorFFromColorRef(fg));
                    };
                    const auto drawTextSpan = [&](size_t visibleStart, size_t visibleEnd, const D2D1_RECT_F& spanRc) noexcept
                    {
                        if (visibleEnd <= visibleStart || spanRc.right <= spanRc.left)
                        {
                            return;
                        }

                        const size_t visibleLength = visibleEnd - visibleStart;
                        const UINT32 visibleLen = visibleLength > static_cast<size_t>(std::numeric_limits<UINT32>::max()) ? std::numeric_limits<UINT32>::max()
                                                                                                                          : static_cast<UINT32>(visibleLength);
                        if (visibleLen == 0u)
                        {
                            return;
                        }

                        _textViewTarget->DrawTextW(
                            _textBuffer.data() + visibleStart, visibleLen, _textViewFormat.get(), spanRc, _textViewBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    };
                    const auto drawCaretForSpan = [&](size_t visibleStart, size_t visibleEnd, float spanLeft, const D2D1_RECT_F& spanRc) noexcept
                    {
                        if (! hasFocus || charW <= 0.0f || _textCaretIndex < visibleStart || _textCaretIndex > visibleEnd || spanRc.right <= spanRc.left)
                        {
                            return;
                        }

                        const size_t caretCol     = _textCaretIndex - visibleStart;
                        const float caretX        = std::min(spanRc.right, spanLeft + static_cast<float>(caretCol) * charW);
                        const D2D1_RECT_F caretRc = D2D1::RectF(caretX, y, std::min(spanRc.right, caretX + 1.0f), lineBottom);
                        _textViewTarget->FillRectangle(caretRc, _textViewBrush.get());
                    };

                    if (splitPaneRow)
                    {
                        drawSearchHighlightsForSpan(leftStartIndex, leftEndIndex, leftPaneRc.left, leftPaneRc);
                        drawSearchHighlightsForSpan(rightStartIndex, rightEndIndex, rightPaneRc.left, rightPaneRc);
                        drawSelectionForSpan(leftStartIndex, leftEndIndex, leftPaneRc.left, leftPaneRc);
                        drawSelectionForSpan(rightStartIndex, rightEndIndex, rightPaneRc.left, rightPaneRc);
                        drawTextSpan(leftStartIndex, leftEndIndex, leftPaneRc);
                        drawTextSpan(rightStartIndex, rightEndIndex, rightPaneRc);
                        if (rowRender)
                        {
                            drawDimmedMarker(rowRender->leftMarkerIndex, leftStartIndex, leftEndIndex, leftPaneRc.left, leftPaneRc);
                            drawDimmedMarker(rowRender->rightMarkerIndex, rightStartIndex, rightEndIndex, rightPaneRc.left, rightPaneRc);
                        }
                        drawCaretForSpan(leftStartIndex, leftEndIndex, leftPaneRc.left, leftPaneRc);
                        drawCaretForSpan(rightStartIndex, rightEndIndex, rightPaneRc.left, rightPaneRc);
                    }
                    else
                    {
                        drawSearchHighlightsForSpan(startIndex, endIndex, x, lineRc);
                        drawSelectionForSpan(startIndex, endIndex, x, lineRc);
                        drawTextSpan(startIndex, endIndex, lineRc);
                        if (rowRender)
                        {
                            drawDimmedMarker(rowRender->fullMarkerIndex, startIndex, endIndex, x, lineRc);
                        }
                        drawCaretForSpan(startIndex, endIndex, x, lineRc);
                    }

                    visibleRowsDrawn += 1u;
                }
            }

#ifdef _DEBUG
            _debugTextVisibleRowCount = visibleRowsDrawn;
            _debugTextLastPaintUs =
                static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - paintStartedAt).count());
#endif

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

    if (_textVisualLineStarts.empty() || _textVisualLineLogical.empty() || _textVisualLineLayouts.empty() || _textLineStarts.empty() || _textLineEnds.empty())
    {
        return DefWindowProcW(hwnd, WM_KEYDOWN, vk, lParam);
    }

    static_cast<void>(EnsureTextViewDirect2D(hwnd));

    auto findVisualForIndex = [&]() noexcept -> uint32_t
    {
        uint32_t visual = 0u;
        if (! FindTextVisualLineForIndex(_textCaretIndex, visual))
        {
            return 0u;
        }
        return visual;
    };

    auto getSegmentBounds = [&](uint32_t visual, uint32_t& outStart, uint32_t& outEnd, bool preferRightPane) noexcept -> uint32_t
    {
        uint32_t logical = 0u;
        if (! GetTextVisualLineSegment(visual, logical, outStart, outEnd, preferRightPane))
        {
            outStart = 0u;
            outEnd   = 0u;
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
            uint32_t segStart = 0;
            uint32_t segEnd   = 0;
            const TextVisualLineLayoutEntry* caretLayout =
                static_cast<size_t>(caretVisual) < _textVisualLineLayouts.size() ? &_textVisualLineLayouts[static_cast<size_t>(caretVisual)] : nullptr;
            const bool preferRightPane = caretLayout && caretLayout->splitPanes &&
                                         (_textCaretIndex >= caretLayout->rightStartIndex ||
                                          (_textCaretIndex >= caretLayout->separatorStartIndex && caretLayout->rightEndIndex > caretLayout->rightStartIndex));
            const uint32_t logical     = getSegmentBounds(caretVisual, segStart, segEnd, preferRightPane);
            const uint32_t lineStart   = caretLayout && caretLayout->splitPanes ? segStart : _textLineStarts[logical];
            const size_t caretIndex    = std::min<size_t>(_textCaretIndex, _textBuffer.size());

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

        RefreshTextHorizontalViewport(hwnd);
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
    const TextVisualLineLayoutEntry* currentLayout =
        static_cast<size_t>(currentVisual) < _textVisualLineLayouts.size() ? &_textVisualLineLayouts[static_cast<size_t>(currentVisual)] : nullptr;
    const bool preferRightPane = currentLayout && currentLayout->splitPanes &&
                                 (_textCaretIndex >= currentLayout->rightStartIndex ||
                                  (_textCaretIndex >= currentLayout->separatorStartIndex && currentLayout->rightEndIndex > currentLayout->rightStartIndex));
    uint32_t segStart          = 0;
    uint32_t segEnd            = 0;
    static_cast<void>(getSegmentBounds(currentVisual, segStart, segEnd, preferRightPane));

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
        static_cast<void>(getSegmentBounds(targetVisual, targetStart, targetEnd, preferRightPane));

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
        case WM_SETCURSOR: return OnTextViewSetCursor(hwnd, lp);
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

bool ViewerText::HasPaneLocalSideBySideVisualLayout() const noexcept
{
    return _documentKind == DocumentKind::Diff && _diffParsedAvailable && _diffPresentation == DiffPresentationMode::SideBySide &&
           _textSideBySideLeftPaneColumns > 0u && _textSideBySideRightPaneColumns > 0u;
}

bool ViewerText::FindTextVisualLineForIndex(size_t index, uint32_t& visualLineOut) const noexcept
{
    if (_textVisualLineLayouts.empty())
    {
        visualLineOut = 0u;
        return false;
    }

    const uint32_t clampedIndex =
        index > static_cast<size_t>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max() : static_cast<uint32_t>(index);

    uint32_t fallbackVisual = 0u;
    for (size_t visualIndex = 0u; visualIndex < _textVisualLineLayouts.size(); ++visualIndex)
    {
        const TextVisualLineLayoutEntry& layout = _textVisualLineLayouts[visualIndex];
        if (clampedIndex >= layout.segmentStartIndex)
        {
            fallbackVisual = static_cast<uint32_t>(visualIndex);
        }

        if (! layout.splitPanes)
        {
            if (clampedIndex >= layout.segmentStartIndex && clampedIndex <= layout.segmentEndIndex)
            {
                visualLineOut = static_cast<uint32_t>(visualIndex);
                return true;
            }
            continue;
        }

        const bool inLeft      = clampedIndex >= layout.leftStartIndex && clampedIndex <= layout.leftEndIndex;
        const bool inSeparator = clampedIndex >= layout.separatorStartIndex && clampedIndex <= layout.separatorEndIndex;
        const bool inRight     = clampedIndex >= layout.rightStartIndex && clampedIndex <= layout.rightEndIndex;
        if (inLeft || inSeparator || inRight)
        {
            visualLineOut = static_cast<uint32_t>(visualIndex);
            return true;
        }
    }

    visualLineOut = fallbackVisual;
    return true;
}

bool ViewerText::GetTextVisualLineSegment(
    uint32_t visualLine, uint32_t& logicalLineOut, uint32_t& segmentStartOut, uint32_t& segmentEndOut, bool preferRightPane) const noexcept
{
    if (_textVisualLineLayouts.empty() || _textLineStarts.empty())
    {
        logicalLineOut  = 0u;
        segmentStartOut = 0u;
        segmentEndOut   = 0u;
        return false;
    }

    const size_t visualIndex                = std::min<size_t>(visualLine, _textVisualLineLayouts.size() - 1u);
    const TextVisualLineLayoutEntry& layout = _textVisualLineLayouts[visualIndex];
    logicalLineOut                          = std::min<uint32_t>(layout.logicalLine, static_cast<uint32_t>(_textLineStarts.size() - 1u));

    if (! layout.splitPanes)
    {
        segmentStartOut = layout.segmentStartIndex;
        segmentEndOut   = layout.segmentEndIndex;
        if (segmentEndOut < segmentStartOut)
        {
            segmentEndOut = segmentStartOut;
        }
        return true;
    }

    const uint32_t leftLength  = layout.leftEndIndex >= layout.leftStartIndex ? (layout.leftEndIndex - layout.leftStartIndex) : 0u;
    const uint32_t rightLength = layout.rightEndIndex >= layout.rightStartIndex ? (layout.rightEndIndex - layout.rightStartIndex) : 0u;
    const bool useRightPane    = preferRightPane && rightLength > 0u;

    if (useRightPane)
    {
        segmentStartOut = layout.rightStartIndex;
        segmentEndOut   = layout.rightEndIndex;
    }
    else if (leftLength > 0u || rightLength == 0u)
    {
        segmentStartOut = layout.leftStartIndex;
        segmentEndOut   = layout.leftEndIndex;
    }
    else
    {
        segmentStartOut = layout.rightStartIndex;
        segmentEndOut   = layout.rightEndIndex;
    }

    if (segmentEndOut < segmentStartOut)
    {
        segmentEndOut = segmentStartOut;
    }
    return true;
}

void ViewerText::RebuildTextVisualLines(HWND hwnd) noexcept
{
    _textVisualLineStarts.clear();
    _textVisualLineLogical.clear();
    _textVisualLineLayouts.clear();
    _textWrapColumns                = 0;
    _textSideBySideLeftPaneColumns  = 0u;
    _textSideBySideRightPaneColumns = 0u;
    _textSideBySideSeparatorColumns = 0u;

    if (_textLineStarts.empty())
    {
        _textVisualLineStarts.push_back(0);
        _textVisualLineLogical.push_back(0);
        _textVisualLineLayouts.push_back({});
        return;
    }

    uint32_t maxCols       = std::numeric_limits<uint32_t>::max();
    uint32_t availableCols = maxCols;
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
        const float marginDip = RoundDipToDevicePixels(6.0f, dpi);
        float availDip        = std::max(0.0f, widthDip - 2.0f * marginDip);
        if (ShowTextLineNumbersInCurrentPresentation() && charW > 0.0f)
        {
            const size_t digits      = LineNumberDigits(_textLineStarts.size());
            const size_t gutterChars = digits + 2u;
            const float gutterDip    = static_cast<float>(gutterChars) * charW;
            availDip                 = std::max(0.0f, availDip - gutterDip);
        }
        const float colsF = availDip / charW;
        maxCols           = std::max<uint32_t>(1u, static_cast<uint32_t>(std::floor(colsF)));
        availableCols     = maxCols;
        _textWrapColumns  = maxCols;
        _textLeftColumn   = 0;
    }
    else if (hwnd)
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
        const float marginDip = RoundDipToDevicePixels(6.0f, dpi);
        float availDip        = std::max(0.0f, widthDip - 2.0f * marginDip);
        if (ShowTextLineNumbersInCurrentPresentation() && charW > 0.0f)
        {
            const size_t digits      = LineNumberDigits(_textLineStarts.size());
            const size_t gutterChars = digits + 2u;
            const float gutterDip    = static_cast<float>(gutterChars) * charW;
            availDip                 = std::max(0.0f, availDip - gutterDip);
        }
        const float colsF = availDip / charW;
        availableCols     = std::max<uint32_t>(1u, static_cast<uint32_t>(std::floor(colsF)));
    }

    const DiffTextVariant* activeDiffVariant = CurrentDiffVariant();
    const bool hasPaneLocalLayout = _documentKind == DocumentKind::Diff && _diffParsedAvailable && _diffPresentation == DiffPresentationMode::SideBySide &&
                                    activeDiffVariant && activeDiffVariant->logicalRowPaneLayouts.size() == _textLineStarts.size();
    size_t reserveVisualCount     = _textLineStarts.size();
    if (_wrap && reserveVisualCount > 0u)
    {
        reserveVisualCount *= hasPaneLocalLayout ? 3u : 2u;
    }
    _textVisualLineStarts.reserve(reserveVisualCount);
    _textVisualLineLogical.reserve(reserveVisualCount);
    _textVisualLineLayouts.reserve(reserveVisualCount);
    if (hasPaneLocalLayout)
    {
        const uint32_t separatorColumns = 3u;
        const uint32_t paneBudget       = availableCols > separatorColumns ? (availableCols - separatorColumns) : 2u;
        _textSideBySideLeftPaneColumns  = std::max<uint32_t>(1u, paneBudget / 2u);
        _textSideBySideRightPaneColumns = std::max<uint32_t>(1u, paneBudget - _textSideBySideLeftPaneColumns);
        _textSideBySideSeparatorColumns = separatorColumns;
        if (_textSideBySideRightPaneColumns == 0u)
        {
            _textSideBySideRightPaneColumns = 1u;
        }
    }

    if (! _wrap)
    {
        uint32_t paneTextMaxColumns = _textMaxLineLength;
        for (uint32_t line = 0; line < static_cast<uint32_t>(_textLineStarts.size()); ++line)
        {
            const uint32_t start = _textLineStarts[line];
            const uint32_t end   = _textLineEnds.size() > line ? _textLineEnds[line] : start;

            if (hasPaneLocalLayout)
            {
                const auto& paneLayout = activeDiffVariant->logicalRowPaneLayouts[line];
                if (paneLayout.splitRow)
                {
                    const uint32_t leftStart      = start;
                    const uint32_t leftEnd        = std::min<uint32_t>(end, start + paneLayout.leftTextColumns);
                    const uint32_t separatorStart = leftEnd;
                    const uint32_t separatorEnd   = std::min<uint32_t>(end, separatorStart + paneLayout.separatorColumns);
                    const uint32_t rightStart     = separatorEnd;
                    const uint32_t rightEnd       = end;
                    const uint32_t leftOffset     = std::min<uint32_t>(_textLeftColumn, paneLayout.leftTextColumns);
                    const uint32_t rightOffset    = std::min<uint32_t>(_textLeftColumn, paneLayout.rightTextColumns);
                    const uint32_t leftVisible    = leftOffset < paneLayout.leftTextColumns
                                                        ? std::min<uint32_t>(_textSideBySideLeftPaneColumns, paneLayout.leftTextColumns - leftOffset)
                                                        : 0u;
                    const uint32_t rightVisible   = rightOffset < paneLayout.rightTextColumns
                                                        ? std::min<uint32_t>(_textSideBySideRightPaneColumns, paneLayout.rightTextColumns - rightOffset)
                                                        : 0u;

                    TextVisualLineLayoutEntry layout{};
                    layout.logicalLine       = line;
                    layout.segmentStartIndex = leftVisible > 0u ? (leftStart + leftOffset) : (rightVisible > 0u ? (rightStart + rightOffset) : leftEnd);
                    layout.segmentEndIndex =
                        rightVisible > 0u ? (rightStart + rightOffset + rightVisible) : (leftVisible > 0u ? (leftStart + leftOffset + leftVisible) : leftEnd);
                    layout.splitPanes          = true;
                    layout.leftStartIndex      = leftVisible > 0u ? (leftStart + leftOffset) : leftEnd;
                    layout.leftEndIndex        = layout.leftStartIndex + leftVisible;
                    layout.rightStartIndex     = rightVisible > 0u ? (rightStart + rightOffset) : rightEnd;
                    layout.rightEndIndex       = layout.rightStartIndex + rightVisible;
                    layout.separatorStartIndex = separatorStart;
                    layout.separatorEndIndex   = separatorEnd;
                    layout.leftPaneColumns     = _textSideBySideLeftPaneColumns;
                    layout.rightPaneColumns    = _textSideBySideRightPaneColumns;
                    layout.separatorColumns    = _textSideBySideSeparatorColumns;
                    _textVisualLineStarts.push_back(layout.segmentStartIndex);
                    _textVisualLineLogical.push_back(line);
                    _textVisualLineLayouts.push_back(layout);
                    paneTextMaxColumns = std::max<uint32_t>(paneTextMaxColumns, std::max(paneLayout.leftTextColumns, paneLayout.rightTextColumns));
                    continue;
                }
            }

            _textVisualLineStarts.push_back(start);
            _textVisualLineLogical.push_back(line);
            _textVisualLineLayouts.push_back(TextVisualLineLayoutEntry{
                .logicalLine       = line,
                .segmentStartIndex = start,
                .segmentEndIndex   = end,
            });
        }

        if (hasPaneLocalLayout)
        {
            _textMaxLineLength = paneTextMaxColumns;
        }
        return;
    }

    for (uint32_t line = 0; line < static_cast<uint32_t>(_textLineStarts.size()); ++line)
    {
        const uint32_t start = _textLineStarts[line];
        const uint32_t end   = _textLineEnds.size() > line ? _textLineEnds[line] : start;
        const uint32_t len   = (end >= start) ? (end - start) : 0;

        if (hasPaneLocalLayout)
        {
            const auto& paneLayout = activeDiffVariant->logicalRowPaneLayouts[line];
            if (paneLayout.splitRow)
            {
                const uint32_t leftStart      = start;
                const uint32_t leftEnd        = std::min<uint32_t>(end, start + paneLayout.leftTextColumns);
                const uint32_t separatorStart = leftEnd;
                const uint32_t separatorEnd   = std::min<uint32_t>(end, separatorStart + paneLayout.separatorColumns);
                const uint32_t rightStart     = separatorEnd;
                const uint32_t rightEnd       = end;
                const auto wrappedRows        = [](uint32_t textColumns, uint32_t paneColumns) noexcept
                {
                    if (textColumns == 0u)
                    {
                        return 1u;
                    }
                    return std::max<uint32_t>(1u, (textColumns + paneColumns - 1u) / paneColumns);
                };

                const uint32_t rowCount = std::max<uint32_t>(wrappedRows(paneLayout.leftTextColumns, _textSideBySideLeftPaneColumns),
                                                             wrappedRows(paneLayout.rightTextColumns, _textSideBySideRightPaneColumns));
                for (uint32_t wrapRow = 0u; wrapRow < rowCount; ++wrapRow)
                {
                    const uint32_t leftOffset   = wrapRow * _textSideBySideLeftPaneColumns;
                    const uint32_t rightOffset  = wrapRow * _textSideBySideRightPaneColumns;
                    const uint32_t leftVisible  = leftOffset < paneLayout.leftTextColumns
                                                      ? std::min<uint32_t>(_textSideBySideLeftPaneColumns, paneLayout.leftTextColumns - leftOffset)
                                                      : 0u;
                    const uint32_t rightVisible = rightOffset < paneLayout.rightTextColumns
                                                      ? std::min<uint32_t>(_textSideBySideRightPaneColumns, paneLayout.rightTextColumns - rightOffset)
                                                      : 0u;

                    TextVisualLineLayoutEntry layout{};
                    layout.logicalLine       = line;
                    layout.segmentStartIndex = leftVisible > 0u ? (leftStart + leftOffset) : (rightVisible > 0u ? (rightStart + rightOffset) : leftEnd);
                    layout.segmentEndIndex =
                        rightVisible > 0u ? (rightStart + rightOffset + rightVisible) : (leftVisible > 0u ? (leftStart + leftOffset + leftVisible) : leftEnd);
                    layout.splitPanes          = true;
                    layout.leftStartIndex      = leftVisible > 0u ? (leftStart + leftOffset) : leftEnd;
                    layout.leftEndIndex        = layout.leftStartIndex + leftVisible;
                    layout.rightStartIndex     = rightVisible > 0u ? (rightStart + rightOffset) : rightEnd;
                    layout.rightEndIndex       = layout.rightStartIndex + rightVisible;
                    layout.separatorStartIndex = separatorStart;
                    layout.separatorEndIndex   = separatorEnd;
                    layout.leftPaneColumns     = _textSideBySideLeftPaneColumns;
                    layout.rightPaneColumns    = _textSideBySideRightPaneColumns;
                    layout.separatorColumns    = _textSideBySideSeparatorColumns;
                    _textVisualLineStarts.push_back(layout.segmentStartIndex);
                    _textVisualLineLogical.push_back(line);
                    _textVisualLineLayouts.push_back(layout);
                }
                continue;
            }
        }

        if (len == 0)
        {
            _textVisualLineStarts.push_back(start);
            _textVisualLineLogical.push_back(line);
            _textVisualLineLayouts.push_back(TextVisualLineLayoutEntry{
                .logicalLine       = line,
                .segmentStartIndex = start,
                .segmentEndIndex   = start,
            });
            continue;
        }

        for (uint32_t col = 0; col < len; col += maxCols)
        {
            const uint32_t segStart = start + col;
            _textVisualLineStarts.push_back(segStart);
            _textVisualLineLogical.push_back(line);
            _textVisualLineLayouts.push_back(TextVisualLineLayoutEntry{
                .logicalLine       = line,
                .segmentStartIndex = segStart,
                .segmentEndIndex   = std::min<uint32_t>(end, segStart + maxCols),
            });
        }
    }

    if (_textVisualLineStarts.empty())
    {
        _textVisualLineStarts.push_back(0);
        _textVisualLineLogical.push_back(0);
        _textVisualLineLayouts.push_back({});
    }
}

std::optional<std::pair<uint32_t, uint32_t>> ViewerText::ComputeVisibleDiffHydrationLogicalRange(HWND hwnd) const noexcept
{
    if (_documentKind != DocumentKind::Diff || ! _diffParsedAvailable || _diffPresentation == DiffPresentationMode::RawText ||
        _config.diffContextMode != DiffContextMode::FullFileWhenAvailable)
    {
        return std::nullopt;
    }

    const HWND textWindow = hwnd ? hwnd : _hEdit.get();
    uint32_t pageRows     = 1u;
    if (textWindow)
    {
        RECT client{};
        GetClientRect(textWindow, &client);
        const UINT dpi        = GetDpiForWindow(textWindow);
        const float heightDip = std::max(1.0f, DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi));
        const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
        pageRows              = std::max<uint32_t>(1u, static_cast<uint32_t>(std::floor(heightDip / lineH)));
    }

    uint32_t topLogical    = 0u;
    uint32_t bottomLogical = pageRows > 0u ? (pageRows - 1u) : 0u;
    if (! _textVisualLineLogical.empty())
    {
        const size_t topVisual    = std::min<size_t>(_textTopVisualLine, _textVisualLineLogical.size() - 1u);
        const size_t bottomVisual = std::min<size_t>(topVisual + std::max<size_t>(1u, static_cast<size_t>(pageRows)) - 1u, _textVisualLineLogical.size() - 1u);
        topLogical                = _textVisualLineLogical[topVisual];
        bottomLogical             = _textVisualLineLogical[bottomVisual];
    }

    const uint32_t hydratedStart = topLogical > kDiffViewportHydrationMarginRows ? (topLogical - kDiffViewportHydrationMarginRows) : 0u;
    uint32_t hydratedEnd         = bottomLogical + 1u;
    if (hydratedEnd <= std::numeric_limits<uint32_t>::max() - kDiffViewportHydrationMarginRows)
    {
        hydratedEnd += kDiffViewportHydrationMarginRows;
    }
    else
    {
        hydratedEnd = std::numeric_limits<uint32_t>::max();
    }

    return std::pair<uint32_t, uint32_t>{hydratedStart, hydratedEnd};
}

void ViewerText::RefreshTextHorizontalViewport(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (! _wrap && HasPaneLocalSideBySideVisualLayout())
    {
        const uint32_t savedTopVisualLine = _textTopVisualLine;
        RebuildTextVisualLines(hwnd);
        if (! _textVisualLineStarts.empty())
        {
            _textTopVisualLine = std::min<uint32_t>(savedTopVisualLine, static_cast<uint32_t>(_textVisualLineStarts.size() - 1u));
        }
        else
        {
            _textTopVisualLine = 0u;
        }
    }

    UpdateTextViewScrollBars(hwnd);
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
    SyncFileComboSelection();

    if (_wrap)
    {
        ShowScrollBar(hwnd, SB_HORZ, FALSE);
        return;
    }

    float widthDip    = std::max(1.0f, DipsFromPixels(static_cast<int>(client.right - client.left), dpi));
    const float charW = (_textCharWidthDip > 0.0f) ? _textCharWidthDip : 8.0f;
    if (ShowTextLineNumbersInCurrentPresentation() && charW > 0.0f)
    {
        const size_t digits      = LineNumberDigits(_textLineStarts.size());
        const size_t gutterChars = digits + 2u;
        const float gutterDip    = static_cast<float>(gutterChars) * charW;
        widthDip                 = std::max(1.0f, widthDip - gutterDip);
    }
    const bool paneLocalHScroll = HasPaneLocalSideBySideVisualLayout() && _textSideBySideLeftPaneColumns > 0u && _textSideBySideRightPaneColumns > 0u;
    const uint32_t pageCols     = paneLocalHScroll ? std::max<uint32_t>(1u, std::min(_textSideBySideLeftPaneColumns, _textSideBySideRightPaneColumns))
                                                   : std::max<uint32_t>(1u, static_cast<uint32_t>(std::floor(widthDip / charW)));
    const uint32_t maxCol       = _textMaxLineLength;

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

bool ViewerText::EnsureVisibleDiffViewportHydrated(HWND hwnd) noexcept
{
    if (_documentKind != DocumentKind::Diff || ! _diffParsedAvailable || _diffPresentation == DiffPresentationMode::RawText ||
        _config.diffContextMode != DiffContextMode::FullFileWhenAvailable)
    {
        return false;
    }

    static_cast<void>(EnsureCurrentDiffVariantBuilt());
    const auto* variant = CurrentDiffVariant();
    if (! variant)
    {
        return false;
    }

    const auto requestedRange = ComputeVisibleDiffHydrationLogicalRange(hwnd);
    if (! requestedRange.has_value())
    {
        return false;
    }

    const bool needsDeferredRowHydration = variant->deferredContextRowCount != 0u && (requestedRange->first < variant->hydratedLogicalLineStart ||
                                                                                      requestedRange->second > variant->hydratedLogicalLineEndExclusive);
    const bool needsTailGrowth           = variant->hasExpandableTail && requestedRange->second >= variant->builtLogicalLineCount;
    if (! needsDeferredRowHydration && ! needsTailGrowth)
    {
        return false;
    }

    _diffInlineExpandedBuilt     = false;
    _diffSideBySideExpandedBuilt = false;
    ApplyCurrentTextPresentation(hwnd ? hwnd : _hEdit.get(), true);
    return true;
}

void ViewerText::ScrollTextViewportToLogicalLine(HWND hwnd, uint32_t targetLogicalLine) noexcept
{
    if (_textLineStarts.empty())
    {
        return;
    }

    targetLogicalLine = std::min<uint32_t>(targetLogicalLine, static_cast<uint32_t>(_textLineStarts.size() - 1u));

    uint32_t targetVisualLine = targetLogicalLine;
    if (! _textVisualLineLogical.empty())
    {
        const auto it = std::lower_bound(_textVisualLineLogical.begin(), _textVisualLineLogical.end(), targetLogicalLine);
        if (it != _textVisualLineLogical.end())
        {
            targetVisualLine = static_cast<uint32_t>(std::distance(_textVisualLineLogical.begin(), it));
        }
        else
        {
            targetVisualLine = static_cast<uint32_t>(_textVisualLineLogical.size() - 1u);
        }
    }

    _textTopVisualLine = targetVisualLine;
    _textLeftColumn    = 0u;

    const size_t caretIndex = std::min<size_t>(_textLineStarts[targetLogicalLine], _textBuffer.size());
    _textCaretIndex         = caretIndex;
    _textSelAnchor          = caretIndex;
    _textSelActive          = caretIndex;
    _textPreferredColumn    = 0u;
    _textSelecting          = false;

    const HWND textWindow = _hEdit ? _hEdit.get() : hwnd;
    static_cast<void>(EnsureVisibleDiffViewportHydrated(textWindow));
    if (textWindow)
    {
        UpdateTextViewScrollBars(textWindow);
        InvalidateRect(textWindow, nullptr, TRUE);
    }

    if (_hWnd)
    {
        SyncFileComboSelection();
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void ViewerText::ScrollToDiffSection(HWND hwnd, size_t sectionIndex) noexcept
{
    const auto* currentVariant = CurrentDiffVariant();
    if (_documentKind == DocumentKind::Diff && _diffParsedAvailable && _diffPresentation != DiffPresentationMode::RawText &&
        _config.diffContextMode == DiffContextMode::FullFileWhenAvailable && currentVariant && sectionIndex < currentVariant->sectionNavigation.size())
    {
        if (! _diffExpandedSectionIndex.has_value() || _diffExpandedSectionIndex.value() != sectionIndex)
        {
            _diffExpandedSectionIndex = sectionIndex;
            _diffReferenceCache.reset();
            _diffInlineExpandedBuilt     = false;
            _diffSideBySideExpandedBuilt = false;
            ApplyCurrentTextPresentation(hwnd ? hwnd : _hEdit.get());
        }
    }

    const auto* variant = CurrentDiffVariant();
    if (! variant)
    {
        if (_documentKind == DocumentKind::Diff && ! _diffParsedAvailable && _textStreamActive && sectionIndex < _diffStreamSections.size())
        {
            const uint64_t targetOffset = AlignTextStreamOffset(_diffStreamSections[sectionIndex].startOffset);
            if (SUCCEEDED(LoadTextToEdit(hwnd, targetOffset, false)))
            {
                SyncFileComboSelection();
            }
        }
        return;
    }

    if (sectionIndex >= variant->sectionNavigation.size() || _textLineStarts.empty())
    {
        return;
    }

    ScrollTextViewportToLogicalLine(hwnd, variant->sectionNavigation[sectionIndex].startLogicalLine);
}

void ViewerText::ScrollToDiffHunk(HWND hwnd, size_t hunkIndex) noexcept
{
    const auto* currentVariant = CurrentDiffVariant();
    if (_documentKind == DocumentKind::Diff && _diffParsedAvailable && _diffPresentation != DiffPresentationMode::RawText &&
        _config.diffContextMode == DiffContextMode::FullFileWhenAvailable && currentVariant && hunkIndex < currentVariant->hunkNavigation.size())
    {
        const size_t targetSectionIndex = currentVariant->hunkNavigation[hunkIndex].sectionIndex;
        if (! _diffExpandedSectionIndex.has_value() || _diffExpandedSectionIndex.value() != targetSectionIndex)
        {
            _diffExpandedSectionIndex = targetSectionIndex;
            _diffReferenceCache.reset();
            _diffInlineExpandedBuilt     = false;
            _diffSideBySideExpandedBuilt = false;
            ApplyCurrentTextPresentation(hwnd ? hwnd : _hEdit.get());
        }
    }

    const auto* variant = CurrentDiffVariant();
    if (! variant || hunkIndex >= variant->hunkNavigation.size())
    {
        return;
    }

    ScrollTextViewportToLogicalLine(hwnd, variant->hunkNavigation[hunkIndex].startLogicalLine);
}

bool ViewerText::NavigateDiffHunk(HWND hwnd, bool previous) noexcept
{
    const auto* variant = CurrentDiffVariant();
    if (! variant || variant->hunkNavigation.empty())
    {
        return false;
    }

    uint32_t topLogicalLine = 0u;
    if (! _textVisualLineLogical.empty())
    {
        const size_t topVisual = std::min<size_t>(_textTopVisualLine, _textVisualLineLogical.size() - 1u);
        topLogicalLine         = _textVisualLineLogical[topVisual];
    }

    size_t targetIndex = 0u;
    if (previous)
    {
        const auto it =
            std::lower_bound(variant->hunkNavigation.begin(),
                             variant->hunkNavigation.end(),
                             topLogicalLine,
                             [](const DiffTextVariant::HunkNavigationEntry& entry, uint32_t line) noexcept { return entry.startLogicalLine < line; });
        if (it == variant->hunkNavigation.begin())
        {
            return false;
        }

        targetIndex = static_cast<size_t>(std::distance(variant->hunkNavigation.begin(), it) - 1);
    }
    else
    {
        const auto it =
            std::upper_bound(variant->hunkNavigation.begin(),
                             variant->hunkNavigation.end(),
                             topLogicalLine,
                             [](uint32_t line, const DiffTextVariant::HunkNavigationEntry& entry) noexcept { return line < entry.startLogicalLine; });
        if (it == variant->hunkNavigation.end())
        {
            return false;
        }

        targetIndex = static_cast<size_t>(std::distance(variant->hunkNavigation.begin(), it));
    }

    ScrollToDiffHunk(hwnd, targetIndex);
    return true;
}

bool ViewerText::EnsureTextViewDirect2D(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    const UINT dpi                          = GetDpiForWindow(hwnd);
    const float dpiF                        = static_cast<float>(dpi);
    const MonoTextRenderMetrics monoMetrics = ComputeMonoTextRenderMetrics(kMonoFontSizeDip, dpi);

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
        static_cast<void>(_textViewFormat->SetLineSpacing(DWRITE_LINE_SPACING_METHOD_UNIFORM, monoMetrics.lineSpacingDip, monoMetrics.baselineDip));
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
        static_cast<void>(_textViewFormatRight->SetLineSpacing(DWRITE_LINE_SPACING_METHOD_UNIFORM, monoMetrics.lineSpacingDip, monoMetrics.baselineDip));
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
                _textCharWidthDip = std::max(1.0f, metrics.widthIncludingTrailingWhitespace);
            }
        }

        _textLineHeightDip = monoMetrics.lineSpacingDip;
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
    RefreshConfigurationJson();

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
    _wrap            = wrap;
    _config.wrapText = wrap;
    RefreshConfigurationJson();
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
            if (_textVisualLineStarts.empty() || _textVisualLineLogical.empty() || _textVisualLineLayouts.empty())
            {
                return;
            }

            uint32_t caretVisual = 0;
            if (! FindTextVisualLineForIndex(_textCaretIndex, caretVisual))
            {
                caretVisual = 0u;
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
                const TextVisualLineLayoutEntry* caretLayout =
                    static_cast<size_t>(caretVisual) < _textVisualLineLayouts.size() ? &_textVisualLineLayouts[static_cast<size_t>(caretVisual)] : nullptr;
                const bool preferRightPane = caretLayout && caretLayout->splitPanes &&
                                             (_textCaretIndex >= caretLayout->rightStartIndex || (_textCaretIndex >= caretLayout->separatorStartIndex &&
                                                                                                  caretLayout->rightEndIndex > caretLayout->rightStartIndex));
                uint32_t logical           = 0u;
                uint32_t segStart          = 0u;
                uint32_t segEnd            = 0u;
                static_cast<void>(GetTextVisualLineSegment(caretVisual, logical, segStart, segEnd, preferRightPane));
                const uint32_t lineStart = caretLayout && caretLayout->splitPanes ? segStart : _textLineStarts[logical];
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

            RefreshTextHorizontalViewport(_hEdit.get());
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
    _textVisualLineLayouts.clear();
    _textTopVisualLine              = 0;
    _textLeftColumn                 = 0;
    _textCaretIndex                 = 0;
    _textSelAnchor                  = 0;
    _textSelActive                  = 0;
    _textPreferredColumn            = 0;
    _textSelecting                  = false;
    _textSideBySideLeftPaneColumns  = 0u;
    _textSideBySideRightPaneColumns = 0u;
    _textSideBySideSeparatorColumns = 0u;

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
