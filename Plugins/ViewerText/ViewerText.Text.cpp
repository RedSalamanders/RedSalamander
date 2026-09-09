#include "ViewerText.h"

#include "ViewerText.SafetyHelpers.h"

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
#include "UnicodeClipboard.h"
#include "WindowMessages.h"
#include "WindowSizing.h"

#include "resource.h"

extern HINSTANCE g_hInstance;

namespace
{
constexpr float kMonoFontSizeDip                    = 10.0f * 96.0f / 72.0f;
constexpr uint32_t kDiffViewportHydrationMarginRows = 32u;
constexpr uint64_t kSparseWrapMaterializationLimit  = 65536u;
constexpr size_t kSparseViewportLayoutLimit         = 4096u;
constexpr size_t kSparseCheckpointCacheLimit        = 512u;
constexpr size_t kVerticalCaretHistoryLimit         = 256u;
constexpr int kTextStreamModuleAnchor               = 0;

COLORREF ResolveAccentColor(const ViewerTheme& theme, std::wstring_view seed) noexcept
{
    if (theme.rainbowMode)
    {
        const uint32_t h = StableVisualHash32Utf16V1(seed);
        const float hue  = static_cast<float>(h % 360u);
        const float sat  = theme.darkBase ? 0.70f : 0.55f;
        const float val  = theme.darkBase ? 0.95f : 0.85f;
        return ColorRefFromHsvClampedNegativeHueToZero(hue, sat, val);
    }

    return ColorRefFromArgb(theme.accentArgb);
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

HRESULT HitTestCaretPosition(IDWriteTextLayout* layout, size_t textLength, size_t textPosition, float& x, float& y, DWRITE_HIT_TEST_METRICS& metrics) noexcept
{
    x       = 0.0f;
    y       = 0.0f;
    metrics = {};
    if (! layout)
    {
        return E_INVALIDARG;
    }
    if (textLength == 0u)
    {
        return S_OK;
    }
    textPosition          = std::min(textPosition, textLength);
    const bool atEnd      = textPosition == textLength;
    const UINT32 position = static_cast<UINT32>(atEnd ? (textLength - 1u) : textPosition);
    return layout->HitTestTextPosition(position, atEnd ? TRUE : FALSE, &x, &y, &metrics);
}

HRESULT DecodeTextWindow(const std::vector<uint8_t>& bytes, size_t convertBytes, ViewerText::FileEncoding encoding, UINT codePage, std::wstring& text) noexcept
{
    text.clear();
    convertBytes = std::min(convertBytes, bytes.size());
    if (convertBytes == 0u)
    {
        return S_OK;
    }

    if ((encoding == ViewerText::FileEncoding::Utf16LE || encoding == ViewerText::FileEncoding::Utf16BE) && (convertBytes % 2u) == 0u)
    {
        text.resize(convertBytes / 2u);
        memcpy(text.data(), bytes.data(), convertBytes);
        if (encoding == ViewerText::FileEncoding::Utf16BE)
        {
            for (wchar_t& value : text)
            {
                value = static_cast<wchar_t>((static_cast<uint16_t>(value) >> 8u) | (static_cast<uint16_t>(value) << 8u));
            }
        }
        return S_OK;
    }

    if ((encoding == ViewerText::FileEncoding::Utf32LE || encoding == ViewerText::FileEncoding::Utf32BE) && (convertBytes % 4u) == 0u)
    {
        const bool bigEndian = encoding == ViewerText::FileEncoding::Utf32BE;
        text.reserve(convertBytes / 4u);
        for (size_t offset = 0u; offset + 3u < convertBytes; offset += 4u)
        {
            const uint32_t codePoint = bigEndian ? (static_cast<uint32_t>(bytes[offset]) << 24u) | (static_cast<uint32_t>(bytes[offset + 1u]) << 16u) |
                                                       (static_cast<uint32_t>(bytes[offset + 2u]) << 8u) | static_cast<uint32_t>(bytes[offset + 3u])
                                                 : static_cast<uint32_t>(bytes[offset]) | (static_cast<uint32_t>(bytes[offset + 1u]) << 8u) |
                                                       (static_cast<uint32_t>(bytes[offset + 2u]) << 16u) | (static_cast<uint32_t>(bytes[offset + 3u]) << 24u);
            if (codePoint <= 0xFFFFu)
            {
                text.push_back(codePoint >= 0xD800u && codePoint <= 0xDFFFu ? static_cast<wchar_t>(0xFFFDu) : static_cast<wchar_t>(codePoint));
            }
            else if (codePoint <= 0x10FFFFu)
            {
                const uint32_t pair = codePoint - 0x10000u;
                text.push_back(static_cast<wchar_t>(0xD800u + (pair >> 10u)));
                text.push_back(static_cast<wchar_t>(0xDC00u + (pair & 0x3FFu)));
            }
            else
            {
                text.push_back(static_cast<wchar_t>(0xFFFDu));
            }
        }
        return S_OK;
    }

    if (convertBytes > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }
    const int sourceLength = static_cast<int>(convertBytes);
    const int required     = MultiByteToWideChar(codePage, 0, reinterpret_cast<LPCCH>(bytes.data()), sourceLength, nullptr, 0);
    if (required <= 0)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_NO_UNICODE_TRANSLATION);
    }
    text.resize(static_cast<size_t>(required));
    const int written = MultiByteToWideChar(codePage, 0, reinterpret_cast<LPCCH>(bytes.data()), sourceLength, text.data(), required);
    if (written <= 0)
    {
        const DWORD error = GetLastError();
        text.clear();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_NO_UNICODE_TRANSLATION);
    }
    text.resize(static_cast<size_t>(written));
    return S_OK;
}

void BuildTextLineIndexForBuffer(std::wstring_view text, std::vector<uint32_t>& lineStarts, std::vector<uint32_t>& lineEnds, uint32_t& maxLineLength) noexcept
{
    lineStarts.clear();
    lineEnds.clear();
    maxLineLength = 0u;
    size_t start  = 0u;
    for (;;)
    {
        size_t end = start;
        while (end < text.size() && text[end] != L'\r' && text[end] != L'\n')
        {
            end += 1u;
        }
        const uint32_t start32 = static_cast<uint32_t>(std::min<size_t>(start, std::numeric_limits<uint32_t>::max()));
        const uint32_t end32   = static_cast<uint32_t>(std::min<size_t>(end, std::numeric_limits<uint32_t>::max()));
        lineStarts.push_back(start32);
        lineEnds.push_back(end32);
        maxLineLength = std::max(maxLineLength, end32 - start32);
        if (end >= text.size())
        {
            break;
        }
        start = end + ((text[end] == L'\r' && end + 1u < text.size() && text[end + 1u] == L'\n') ? 2u : 1u);
    }
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
    const uint64_t totalLines = TextVisualLineCount();
    if (totalLines == 0)
    {
        return 0;
    }

    const uint64_t maxLine = totalLines - 1;

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_ALL;
    GetScrollInfo(hwnd, SB_VERT, &si);
    if (_textSparseWrapActive)
    {
        const size_t requestedRows = std::min<size_t>(kSparseViewportLayoutLimit, std::max<size_t>(2u, static_cast<size_t>(std::max<UINT>(1u, si.nPage)) + 1u));
        RebuildSparseTextViewportLayouts(hwnd, requestedRows);
    }

    const uint32_t previousTop = _textTopVisualLine;
    bool rememberSparseTop     = false;
    uint64_t top               = _textTopVisualLine;
    const int code             = static_cast<int>(scrollCode);
    switch (code)
    {
        case SB_TOP:
            top = 0;
            _textSparseTopHistory.clear();
            break;
        case SB_BOTTOM:
            top = maxLine;
            _textSparseTopHistory.clear();
            break;
        case SB_LINEUP:
            if (_textSparseWrapActive && ! _textSparseTopHistory.empty())
            {
                top = _textSparseTopHistory.back();
                _textSparseTopHistory.pop_back();
            }
            else if (top > 0)
            {
                top -= 1;
            }
            break;
        case SB_LINEDOWN:
            rememberSparseTop = _textSparseWrapActive;
            if (_textSparseWrapActive && _textSparseViewportAnchors.size() > 1u)
            {
                top = _textSparseViewportAnchors[1u];
            }
            else if (top < maxLine)
            {
                top += 1;
            }
            break;
        case SB_PAGEUP:
        {
            if (_textSparseWrapActive && ! _textSparseTopHistory.empty())
            {
                top = _textSparseTopHistory.back();
                _textSparseTopHistory.pop_back();
            }
            else
            {
                const uint64_t page = std::max<uint64_t>(1u, static_cast<uint64_t>(si.nPage));
                top                 = (top > page) ? (top - page) : 0;
            }
            break;
        }
        case SB_PAGEDOWN:
        {
            rememberSparseTop   = _textSparseWrapActive;
            const uint64_t page = std::max<uint64_t>(1u, static_cast<uint64_t>(si.nPage));
            if (_textSparseWrapActive && _textSparseViewportAnchors.size() > 1u)
            {
                const size_t targetRow = std::min<size_t>(static_cast<size_t>(page), _textSparseViewportAnchors.size() - 1u);
                top                    = _textSparseViewportAnchors[targetRow];
            }
            else
            {
                top = std::min<uint64_t>(maxLine, top + page);
            }
            break;
        }
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION:
        {
            _textSparseTopHistory.clear();
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
    if (rememberSparseTop && top != previousTop)
    {
        constexpr size_t kSparseTopHistoryLimit = 4096u;
        if (_textSparseTopHistory.size() >= kSparseTopHistoryLimit)
        {
            _textSparseTopHistory.erase(_textSparseTopHistory.begin());
        }
        _textSparseTopHistory.push_back(previousTop);
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
    const uint64_t totalLines = TextVisualLineCount();
    if (totalLines == 0)
    {
        return 0;
    }

    uint64_t top = _textTopVisualLine;
    if (signedDelta < 0)
    {
        if (_textSparseWrapActive && ! _textSparseTopHistory.empty())
        {
            top = _textSparseTopHistory.back();
            _textSparseTopHistory.pop_back();
        }
        else
        {
            const uint64_t d = static_cast<uint64_t>(-signedDelta);
            top              = (top > d) ? (top - d) : 0;
        }
    }
    else
    {
        const uint64_t maxLine = totalLines - 1;
        if (_textSparseWrapActive)
        {
            const size_t requestedRows = std::min<size_t>(kSparseViewportLayoutLimit, static_cast<size_t>(signedDelta) + 1u);
            RebuildSparseTextViewportLayouts(hwnd, requestedRows);
            if (_textSparseViewportAnchors.size() > 1u)
            {
                const size_t targetRow = std::min<size_t>(static_cast<size_t>(signedDelta), _textSparseViewportAnchors.size() - 1u);
                top                    = _textSparseViewportAnchors[targetRow];
            }
            else
            {
                top = std::min<uint64_t>(maxLine, top + static_cast<uint64_t>(signedDelta));
            }
        }
        else
        {
            top = std::min<uint64_t>(maxLine, top + static_cast<uint64_t>(signedDelta));
        }
        if (_textSparseWrapActive && top != _textTopVisualLine)
        {
            constexpr size_t kSparseTopHistoryLimit = 4096u;
            if (_textSparseTopHistory.size() >= kSparseTopHistoryLimit)
            {
                _textSparseTopHistory.erase(_textSparseTopHistory.begin());
            }
            _textSparseTopHistory.push_back(_textTopVisualLine);
        }
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
        else if (signedDelta > 0 && totalLines > 0u &&
                 _textTopVisualLine >= static_cast<uint32_t>(std::min<uint64_t>(totalLines - 1u, std::numeric_limits<uint32_t>::max())))
        {
            static_cast<void>(TryNavigateTextStream(GetAncestor(hwnd, GA_ROOT), false));
        }
    }

    return 0;
}

void ViewerText::ClearTextLayoutCache() noexcept
{
    _textLayoutCache.clear();
    _textUncachedLayout.reset();
    _textLayoutCacheBytes = 0u;
    _textSparseViewportLayouts.clear();
    _textSparseViewportAnchors.clear();
    _textSparseViewportCheckpoints.clear();
    _textSparseCheckpointCache.clear();
    _textSparseTopHistory.clear();
    _textVerticalCaretHistory.clear();
    _textPreferredXValid             = false;
    _textSparseViewportTop           = 0u;
    _textSparseViewportRequestedRows = 0u;
    _textSparseViewportComplete      = false;
    if (_textLayoutGeneration == std::numeric_limits<uint64_t>::max())
    {
        _textLayoutGeneration = 1u;
    }
    else
    {
        _textLayoutGeneration += 1u;
    }
}

size_t ViewerText::NormalizeTextPosition(size_t index) const noexcept
{
    index = std::min(index, _textBuffer.size());
    if (index > 0u && index < _textBuffer.size())
    {
        const wchar_t previous = _textBuffer[index - 1u];
        const wchar_t current  = _textBuffer[index];
        if (previous >= static_cast<wchar_t>(0xD800u) && previous <= static_cast<wchar_t>(0xDBFFu) && current >= static_cast<wchar_t>(0xDC00u) &&
            current <= static_cast<wchar_t>(0xDFFFu))
        {
            index += 1u;
        }
    }
    return std::min(index, _textBuffer.size());
}

size_t ViewerText::NormalizeTextSegmentStart(size_t index) const noexcept
{
    index = std::min(index, _textBuffer.size());
    if (index > 0u && index < _textBuffer.size())
    {
        const wchar_t previous = _textBuffer[index - 1u];
        const wchar_t current  = _textBuffer[index];
        if (previous >= static_cast<wchar_t>(0xD800u) && previous <= static_cast<wchar_t>(0xDBFFu) && current >= static_cast<wchar_t>(0xDC00u) &&
            current <= static_cast<wchar_t>(0xDFFFu))
        {
            index -= 1u;
        }
    }
    return index;
}

size_t ViewerText::PreviousTextPosition(size_t index) const noexcept
{
    index = std::min(index, _textBuffer.size());
    if (index == 0u)
    {
        return 0u;
    }

    size_t previous = index - 1u;
    if (previous > 0u && _textBuffer[previous] >= static_cast<wchar_t>(0xDC00u) && _textBuffer[previous] <= static_cast<wchar_t>(0xDFFFu) &&
        _textBuffer[previous - 1u] >= static_cast<wchar_t>(0xD800u) && _textBuffer[previous - 1u] <= static_cast<wchar_t>(0xDBFFu))
    {
        previous -= 1u;
    }
    return previous;
}

size_t ViewerText::NextTextPosition(size_t index) const noexcept
{
    index = std::min(index, _textBuffer.size());
    if (index >= _textBuffer.size())
    {
        return _textBuffer.size();
    }

    size_t next = index + 1u;
    if (next < _textBuffer.size() && _textBuffer[index] >= static_cast<wchar_t>(0xD800u) && _textBuffer[index] <= static_cast<wchar_t>(0xDBFFu) &&
        _textBuffer[next] >= static_cast<wchar_t>(0xDC00u) && _textBuffer[next] <= static_cast<wchar_t>(0xDFFFu))
    {
        next += 1u;
    }
    return std::min(next, _textBuffer.size());
}

size_t ViewerText::FindTextSegmentEnd(size_t startIndex, size_t lineEndIndex, float widthDip) noexcept
{
    startIndex   = NormalizeTextSegmentStart(std::min(startIndex, _textBuffer.size()));
    lineEndIndex = NormalizeTextPosition(std::min(std::max(lineEndIndex, startIndex), _textBuffer.size()));
    if (startIndex >= lineEndIndex)
    {
        return startIndex;
    }

    const float boundedWidth                 = std::max(1.0f, widthDip);
    const float charWidth                    = std::max(1.0f, _textCharWidthDip);
    const double estimatedColumnsValue       = std::clamp(std::ceil(static_cast<double>(boundedWidth) / static_cast<double>(charWidth)), 1.0, 4096.0);
    const size_t estimatedColumns            = static_cast<size_t>(estimatedColumnsValue);
    constexpr size_t kCandidateCodeUnitLimit = 16u * 1024u;
    const size_t candidateCodeUnits          = std::min(kCandidateCodeUnitLimit, estimatedColumns * 4u + 16u);
    const size_t candidateEnd                = NormalizeTextPosition(std::min(lineEndIndex, startIndex + candidateCodeUnits));
    IDWriteTextLayout* layout                = GetTextSegmentLayout(startIndex, candidateEnd, boundedWidth);
    if (! layout)
    {
        return std::min(lineEndIndex, NextTextPosition(startIndex));
    }

    DWRITE_TEXT_METRICS textMetrics{};
    if (SUCCEEDED(layout->GetMetrics(&textMetrics)) && textMetrics.widthIncludingTrailingWhitespace <= boundedWidth + 0.25f)
    {
        return candidateEnd;
    }

    BOOL trailing = FALSE;
    BOOL inside   = FALSE;
    DWRITE_HIT_TEST_METRICS hitMetrics{};
    const HRESULT hitHr =
        layout->HitTestPoint(std::max(0.0f, boundedWidth - 0.25f), std::max(1.0f, _textLineHeightDip) * 0.5f, &trailing, &inside, &hitMetrics);
    if (FAILED(hitHr))
    {
        return std::min(lineEndIndex, NextTextPosition(startIndex));
    }

    size_t localEnd = static_cast<size_t>(hitMetrics.textPosition);
    if (hitMetrics.left + hitMetrics.width <= boundedWidth + 0.25f)
    {
        localEnd += static_cast<size_t>(hitMetrics.length);
    }
    size_t segmentEnd = NormalizeTextPosition(std::min(candidateEnd, startIndex + localEnd));
    if (segmentEnd <= startIndex)
    {
        segmentEnd = std::min(lineEndIndex, NextTextPosition(startIndex));
    }
    return segmentEnd;
}

IDWriteTextLayout* ViewerText::GetTextSegmentLayout(size_t startIndex, size_t endIndex, float widthDip) noexcept
{
    _textUncachedLayout.reset();
    if (! _dwriteFactory || ! _textViewFormat)
    {
        return nullptr;
    }

    startIndex          = NormalizeTextPosition(std::min(startIndex, _textBuffer.size()));
    endIndex            = NormalizeTextPosition(std::min(std::max(endIndex, startIndex), _textBuffer.size()));
    const size_t length = endIndex - startIndex;
    if (length > static_cast<size_t>(std::numeric_limits<UINT32>::max()))
    {
        return nullptr;
    }

    const float boundedWidth = std::max(1.0f, widthDip);
    const uint32_t widthKey  = static_cast<uint32_t>(
        std::min<double>(static_cast<double>(std::numeric_limits<uint32_t>::max()), std::round(static_cast<double>(boundedWidth) * 1000.0)));
    for (auto& entry : _textLayoutCache)
    {
        if (entry.generation == _textLayoutGeneration && entry.startIndex == static_cast<uint32_t>(startIndex) &&
            entry.endIndex == static_cast<uint32_t>(endIndex) && entry.widthMilliDip == widthKey && entry.layout)
        {
            entry.lastUse = ++_textLayoutUseCounter;
#ifdef _DEBUG
            _debugTextLayoutCacheHits += 1u;
#endif
            return entry.layout.get();
        }
    }

#ifdef _DEBUG
    _debugTextLayoutCacheMisses += 1u;
#endif
    wil::com_ptr<IDWriteTextLayout> layout;
    const HRESULT createHr = _dwriteFactory->CreateTextLayout(
        _textBuffer.data() + startIndex, static_cast<UINT32>(length), _textViewFormat.get(), boundedWidth, std::max(1.0f, _textLineHeightDip), layout.put());
    if (FAILED(createHr) || ! layout)
    {
        Debug::Error(L"ViewerText: CreateTextLayout failed for visible text segment (start={}, length={}, hr=0x{:08X}).",
                     startIndex,
                     length,
                     static_cast<unsigned long>(createHr));
        return nullptr;
    }

    static_cast<void>(layout->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
    static_cast<void>(layout->SetIncrementalTabStop(std::max(1.0f, _textCharWidthDip * 4.0f)));

    const size_t estimatedBytes    = sizeof(TextLayoutCacheEntry) + length * sizeof(wchar_t) + 4096u;
    const size_t effectiveMaxBytes = std::max<size_t>(8192u, _textLayoutCacheMaxBytes);
    if (estimatedBytes > effectiveMaxBytes)
    {
        _textUncachedLayout = std::move(layout);
        return _textUncachedLayout.get();
    }

    while (! _textLayoutCache.empty() &&
           (_textLayoutCache.size() >= std::max<size_t>(1u, _textLayoutCacheMaxEntries) || _textLayoutCacheBytes + estimatedBytes > effectiveMaxBytes))
    {
        const auto oldest = std::min_element(
            _textLayoutCache.begin(), _textLayoutCache.end(), [](const auto& left, const auto& right) noexcept { return left.lastUse < right.lastUse; });
        _textLayoutCacheBytes -= std::min(_textLayoutCacheBytes, oldest->estimatedBytes);
        _textLayoutCache.erase(oldest);
        _textLayoutCacheEvictions += 1u;
    }

    TextLayoutCacheEntry entry{};
    entry.startIndex     = static_cast<uint32_t>(startIndex);
    entry.endIndex       = static_cast<uint32_t>(endIndex);
    entry.widthMilliDip  = widthKey;
    entry.generation     = _textLayoutGeneration;
    entry.lastUse        = ++_textLayoutUseCounter;
    entry.estimatedBytes = estimatedBytes;
    entry.layout         = std::move(layout);
    _textLayoutCacheBytes += estimatedBytes;
    _textLayoutCache.push_back(std::move(entry));
    return _textLayoutCache.back().layout.get();
}

std::optional<ViewerText::TextViewHitTestResult> ViewerText::HitTestTextView(HWND hwnd, POINT pt) noexcept
{
    const uint64_t totalVisual = TextVisualLineCount();
    if (! hwnd || totalVisual == 0u)
    {
        return std::nullopt;
    }

    static_cast<void>(EnsureTextViewDirect2D(hwnd));

    const UINT dpi        = GetDpiForWindow(hwnd);
    const float xDip      = Common::WindowSizing::PixelToDip(static_cast<float>(pt.x), static_cast<float>(dpi));
    const float yDip      = Common::WindowSizing::PixelToDip(static_cast<float>(pt.y), static_cast<float>(dpi));
    const float marginDip = RoundDipToDevicePixels(6.0f, dpi);
    const float charW     = (_textCharWidthDip > 0.0f) ? _textCharWidthDip : 8.0f;
    const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
    RECT client{};
    GetClientRect(hwnd, &client);
    const float widthDip = Common::WindowSizing::PixelToDip(static_cast<float>(client.right - client.left), static_cast<float>(dpi));
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

    const float relY   = std::max(0.0f, yDip - marginDip);
    const uint64_t row = static_cast<uint64_t>(std::floor(relY / lineH));
    if (_textSparseWrapActive)
    {
        RebuildSparseTextViewportLayouts(hwnd, static_cast<size_t>(std::min<uint64_t>(row + 1u, kSparseViewportLayoutLimit)));
        if (row >= _textSparseViewportLayouts.size())
        {
            return std::nullopt;
        }
    }
    TextVisualLineLayoutEntry layout{};
    if (_textSparseWrapActive)
    {
        layout = _textSparseViewportLayouts[static_cast<size_t>(row)];
    }
    else
    {
        const uint64_t visual = std::min<uint64_t>(static_cast<uint64_t>(_textTopVisualLine) + row, totalVisual - 1u);
        if (! TryGetTextVisualLineLayout(visual, layout))
        {
            return std::nullopt;
        }
    }
    if (layout.logicalLine >= _textLineEnds.size())
    {
        return std::nullopt;
    }

    const auto hitSpan = [&](size_t start, size_t end, float left, float right) noexcept -> size_t
    {
        start                 = NormalizeTextPosition(std::min(start, _textBuffer.size()));
        end                   = NormalizeTextPosition(std::min(std::max(end, start), _textBuffer.size()));
        const float spanWidth = std::max(1.0f, right - left);
        if (! _wrap && end > start)
        {
            const size_t visibleCodeUnits = static_cast<size_t>(std::ceil(spanWidth / std::max(1.0f, charW))) * 4u + 16u;
            end                           = NormalizeTextPosition(std::min(end, start + visibleCodeUnits));
        }

        IDWriteTextLayout* textLayout = GetTextSegmentLayout(start, end, spanWidth);
        if (! textLayout)
        {
            return start;
        }

        BOOL trailing = FALSE;
        BOOL inside   = FALSE;
        DWRITE_HIT_TEST_METRICS metrics{};
        const HRESULT hitHr = textLayout->HitTestPoint(std::max(0.0f, xDip - left), lineH * 0.5f, &trailing, &inside, &metrics);
        if (FAILED(hitHr))
        {
            Debug::Error(L"ViewerText: HitTestPoint failed for visible text segment (hr=0x{:08X}).", static_cast<unsigned long>(hitHr));
            return start;
        }

        const size_t local = static_cast<size_t>(metrics.textPosition) + (trailing != FALSE ? static_cast<size_t>(metrics.length) : 0u);
        return NormalizeTextPosition(std::min(end, start + local));
    };

    size_t index = 0u;
    if (layout.splitPanes && charW > 0.0f)
    {
        const float leftPaneWidthDip  = static_cast<float>(layout.leftPaneColumns) * charW;
        const float separatorWidthDip = static_cast<float>(layout.separatorColumns) * charW;
        const float separatorLeft     = textStartX + leftPaneWidthDip;
        const float separatorRight    = separatorLeft + separatorWidthDip;

        if (xDip < separatorLeft)
        {
            index = hitSpan(layout.leftStartIndex, layout.leftEndIndex, textStartX, separatorLeft);
        }
        else if (xDip < separatorRight)
        {
            index = NormalizeTextPosition(std::min<size_t>(static_cast<size_t>(layout.separatorStartIndex), _textBuffer.size()));
        }
        else
        {
            index = hitSpan(layout.rightStartIndex, layout.rightEndIndex, separatorRight, separatorRight + static_cast<float>(layout.rightPaneColumns) * charW);
        }
    }
    else
    {
        uint32_t segmentStart = layout.segmentStartIndex;
        uint32_t segmentEnd   = layout.segmentEndIndex;
        if (! _wrap && segmentEnd >= segmentStart && _textLeftColumn != 0u)
        {
            const uint32_t skip = std::min<uint32_t>(_textLeftColumn, segmentEnd - segmentStart);
            segmentStart += skip;
        }

        index = hitSpan(segmentStart, segmentEnd, textStartX, std::max(textStartX + 1.0f, widthDip - marginDip));
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
    if (! hwnd || _viewMode != ViewMode::Text || logicalLine >= _textLineStarts.size() || (! _textSparseWrapActive && _textVisualLineLogical.empty()))
    {
        return false;
    }

    const auto findVisibleVisualLine = [&](uint32_t targetLogicalLine) noexcept -> std::optional<size_t>
    {
        if (_textSparseWrapActive)
        {
            RECT client{};
            GetClientRect(hwnd, &client);
            const UINT dpi        = GetDpiForWindow(hwnd);
            const float heightDip = std::max(1.0f, Common::WindowSizing::PixelToDip(static_cast<float>(client.bottom - client.top), static_cast<float>(dpi)));
            const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
            const size_t pageRows = std::max<size_t>(1u, static_cast<size_t>(std::ceil(heightDip / lineH)) + 1u);
            RebuildSparseTextViewportLayouts(hwnd, pageRows);
            for (size_t row = 0u; row < _textSparseViewportLayouts.size(); ++row)
            {
                if (_textSparseViewportLayouts[row].logicalLine == targetLogicalLine)
                {
                    if (row >= _textSparseViewportAnchors.size())
                    {
                        return std::nullopt;
                    }
                    return static_cast<size_t>(_textSparseViewportAnchors[row]);
                }
            }
            return std::nullopt;
        }

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
        const float heightDip   = Common::WindowSizing::PixelToDip(static_cast<float>(client.bottom - client.top), static_cast<float>(dpi));
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

    // The debug action is also used to address sparse rows deterministically. Put
    // the requested logical line at the viewport origin before deriving its hit
    // coordinates; otherwise an already-visible row retains the preceding sparse
    // anchor as the top row and cannot validate its own pane checkpoints.
    ScrollTextViewportToLogicalLine(hwnd, logicalLine);
    std::optional<size_t> visualLine = findVisibleVisualLine(logicalLine);

    if (! visualLine.has_value())
    {
        return false;
    }

    RECT client{};
    GetClientRect(hwnd, &client);
    const UINT dpi        = GetDpiForWindow(hwnd);
    const float widthDip  = Common::WindowSizing::PixelToDip(static_cast<float>(client.right - client.left), static_cast<float>(dpi));
    const float heightDip = Common::WindowSizing::PixelToDip(static_cast<float>(client.bottom - client.top), static_cast<float>(dpi));
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

    size_t visibleRowIndex = visualLine.value() - static_cast<size_t>(_textTopVisualLine);
    if (_textSparseWrapActive)
    {
        const auto anchor = std::find(_textSparseViewportAnchors.begin(), _textSparseViewportAnchors.end(), visualLine.value());
        if (anchor == _textSparseViewportAnchors.end())
        {
            return false;
        }
        visibleRowIndex = static_cast<size_t>(std::distance(_textSparseViewportAnchors.begin(), anchor));
    }

    const float rowIndexDip = static_cast<float>(visibleRowIndex);
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
    if (! _embeddedMode)
    {
        SetFocus(hwnd);
    }

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
    _textSelActive       = index;
    _textSelecting       = true;
    _textPreferredXValid = false;
    _textVerticalCaretHistory.clear();

    InvalidateRect(hwnd, nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
    return 0;
}

LRESULT ViewerText::OnTextViewMouseMove(HWND hwnd, POINT pt, WPARAM keyState) noexcept
{
#if defined(ENABLE_TESTS) && defined(_DEBUG)
    _debugHasLastTextViewMouseMoveClientPoint = true;
    _debugLastTextViewMouseMoveClientPoint    = pt;
    const auto debugHit                       = HitTestTextView(hwnd, pt);
    _debugLastTextViewMouseMoveHit            = debugHit.has_value();
    _debugLastTextViewMouseMoveLogicalLine    = debugHit.has_value() ? debugHit->logicalLine : static_cast<size_t>(-1);
#endif

    if (! _textSelecting || (keyState & MK_LBUTTON) == 0u)
    {
        return 0;
    }

    const auto hit = HitTestTextView(hwnd, pt);
    if (! hit.has_value())
    {
        return 0;
    }

    _textSelActive       = hit->bufferIndex;
    _textCaretIndex      = _textSelActive;
    _textPreferredXValid = false;
    _textVerticalCaretHistory.clear();

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

    const DWORD messagePos = GetMessagePos();
    POINT pt{static_cast<LONG>(static_cast<short>(LOWORD(messagePos))), static_cast<LONG>(static_cast<short>(HIWORD(messagePos)))};
    if (ScreenToClient(hwnd, &pt) == 0)
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

            const float widthDip  = Common::WindowSizing::PixelToDip(static_cast<float>(rc.right - rc.left), static_cast<float>(dpi));
            const float heightDip = Common::WindowSizing::PixelToDip(static_cast<float>(rc.bottom - rc.top), static_cast<float>(dpi));
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

            const uint64_t totalVisual = TextVisualLineCount();
            const uint64_t topVisual   = static_cast<uint64_t>(_textTopVisualLine);

            const size_t selStartIndex = std::min(_textSelAnchor, _textSelActive);
            const size_t selEndIndex   = std::max(_textSelAnchor, _textSelActive);
            const bool hasSelection    = selStartIndex != selEndIndex;

            const bool hasFocus = GetFocus() == hwnd;

            const std::wstring seed      = _currentPath.empty() ? std::wstring(L"viewer") : _currentPath.filename().wstring();
            const COLORREF accent        = _hasTheme ? ResolveAccentColor(_theme, seed) : RGB(0, 120, 215);
            const uint8_t selectionAlpha = (_hasTheme && _theme.darkMode) ? 90u : 70u;
            const COLORREF selectionBg   = BlendColorRefTruncate(bg, accent, selectionAlpha);

            const bool hasSearchHighlights    = ! _searchQuery.empty() && ! _searchMatchStarts.empty();
            const size_t searchLen            = _searchQuery.size();
            const COLORREF searchAccent       = (_hasTheme && ! _theme.highContrast) ? ResolveAccentColor(_theme, L"search") : GetSysColor(COLOR_HIGHLIGHT);
            const uint8_t searchAlpha         = (_hasTheme && _theme.darkMode) ? 60u : 40u;
            const COLORREF searchBg           = BlendColorRefTruncate(bg, searchAccent, searchAlpha);
            const bool selectionIsSearchMatch = hasSelection && hasSearchHighlights && searchLen > 0 && (selEndIndex - selStartIndex == searchLen) &&
                                                std::binary_search(_searchMatchStarts.begin(), _searchMatchStarts.end(), selStartIndex);
            const uint8_t selectionFocusAlpha = (_hasTheme && _theme.darkMode) ? 140u : 120u;
            const COLORREF selectionFocusedBg = BlendColorRefTruncate(bg, accent, selectionFocusAlpha);

            const bool showLineNumbers                           = ShowTextLineNumbersInCurrentPresentation() && gutterWidthDip > 0.0f;
            const uint8_t lineNumberAlpha                        = (_hasTheme && _theme.darkMode) ? 160u : 140u;
            const COLORREF lineNumberFg                          = BlendColorRefTruncate(bg, fg, lineNumberAlpha);
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
            const D2D1_COLOR_F diffDividerBg = _hasTheme
                                                   ? ColorFFromArgb(_theme.diffDividerArgb)
                                                   : ColorFFromColorRef(BlendColorRefTruncate(bg, accent, (_hasTheme && _theme.darkMode) ? 28u : 18u), 0.80f);
            const D2D1_COLOR_F diffStructuralBorderBg =
                ColorFFromColorRef(BlendColorRefTruncate(bg, fg, (_hasTheme && _theme.darkMode) ? 72u : 48u), (_hasTheme && _theme.darkMode) ? 0.92f : 0.82f);
            const D2D1_COLOR_F diffActiveHunkBorderBg =
                ColorFFromColorRef(BlendColorRefTruncate(bg, accent, (_hasTheme && _theme.darkMode) ? 168u : 120u), 0.96f);
            const COLORREF diffMarkerColorRef   = BlendColorRefTruncate(bg, fg, (_hasTheme && _theme.darkMode) ? 72u : 52u);
            const COLORREF diffGapHatchColorRef = BlendColorRefTruncate(bg, fg, (_hasTheme && _theme.darkMode) ? 48u : 34u);
            const D2D1_COLOR_F diffMarkerFg     = ColorFFromColorRef(diffMarkerColorRef, (_hasTheme && _theme.darkMode) ? 0.88f : 0.78f);
            const D2D1_COLOR_F diffGapHatchFg   = ColorFFromColorRef(diffGapHatchColorRef, (_hasTheme && _theme.darkMode) ? 0.34f : 0.22f);
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
                const COLORREF gutterBg   = BlendColorRefTruncate(bg, accent, gutterAlpha);
                const float gutterRight   = std::min(widthDip, std::max(0.0f, textStartX));

                _textViewBrush->SetColor(ColorFFromColorRef(gutterBg));
                _textViewTarget->FillRectangle(D2D1::RectF(0.0f, 0.0f, gutterRight, heightDip), _textViewBrush.get());

                const COLORREF divider = BlendColorRefTruncate(bg, fg, (_hasTheme && _theme.darkMode) ? 40u : 20u);
                _textViewBrush->SetColor(ColorFFromColorRef(divider));
                const float sepX = std::min(widthDip, std::max(0.0f, textStartX - 1.0f));
                _textViewTarget->DrawLine(D2D1::Point2F(sepX, 0.0f), D2D1::Point2F(sepX, heightDip), _textViewBrush.get(), 1.0f);

                _textViewBrush->SetColor(ColorFFromColorRef(fg));
            }

            if (totalVisual > 0 && lineH > 0.0f && _textViewFormat)
            {
                const float usableH    = std::max(0.0f, heightDip - 2.0f * marginDip);
                const uint32_t maxRows = std::max<uint32_t>(1u, static_cast<uint32_t>(std::ceil(usableH / lineH)) + 1u);
                if (_textSparseWrapActive)
                {
                    RebuildSparseTextViewportLayouts(hwnd, maxRows);
                }

                for (uint32_t row = 0; row < maxRows; ++row)
                {
                    if (_textSparseWrapActive && row >= _textSparseViewportLayouts.size())
                    {
                        break;
                    }
                    const uint64_t visual = _textSparseWrapActive ? _textSparseViewportAnchors[row] : topVisual + static_cast<uint64_t>(row);
                    if (visual >= totalVisual)
                    {
                        break;
                    }

                    TextVisualLineLayoutEntry resolvedLayout{};
                    if (_textSparseWrapActive)
                    {
                        resolvedLayout = _textSparseViewportLayouts[row];
                    }
                    else if (! TryGetTextVisualLineLayout(visual, resolvedLayout))
                    {
                        break;
                    }
                    const uint32_t logical = resolvedLayout.logicalLine;
                    if (logical >= _textLineStarts.size() || logical >= _textLineEnds.size())
                    {
                        break;
                    }

                    const TextVisualLineLayoutEntry* visualLayout = &resolvedLayout;
                    const uint32_t segStartRaw                    = visualLayout->segmentStartIndex;
                    uint32_t segStart                             = segStartRaw;
                    uint32_t segEnd                               = visualLayout ? visualLayout->segmentEndIndex : _textLineEnds[logical];
                    if (! _wrap && (! visualLayout || ! visualLayout->splitPanes) && segEnd >= segStart && _textLeftColumn != 0)
                    {
                        const uint32_t skip = std::min<uint32_t>(_textLeftColumn, segEnd - segStart);
                        segStart += skip;
                    }

                    const size_t startIndex = NormalizeTextPosition(std::min<size_t>(static_cast<size_t>(segStart), _textBuffer.size()));
                    size_t endIndex         = NormalizeTextPosition(std::min<size_t>(static_cast<size_t>(segEnd), _textBuffer.size()));
                    if (! _wrap && ! visualLayout->splitPanes && endIndex > startIndex)
                    {
                        const size_t visibleCodeUnits =
                            static_cast<size_t>(std::ceil(std::max(1.0f, widthDip - textStartX - marginDip) / std::max(1.0f, charW))) * 4u + 16u;
                        endIndex = NormalizeTextPosition(std::min(endIndex, startIndex + visibleCodeUnits));
                    }

                    const float x                       = textStartX;
                    const float y                       = RoundDipToDevicePixels(marginDip + static_cast<float>(row) * lineH, dpi);
                    const float lineBottom              = RoundDipToDevicePixels(y + lineH, dpi);
                    const D2D1_RECT_F lineRc            = D2D1::RectF(x, y, std::max(x, widthDip - marginDip), lineBottom);
                    const bool isFirstSegmentForLogical = visualLayout->splitPanes ? visualLayout->leftStartIndex == _textLineStarts[logical]
                                                                                   : visualLayout->segmentStartIndex == _textLineStarts[logical];
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
                            const D2D1_COLOR_F softDividerBg = ColorFFromColorRef(BlendColorRefTruncate(bg, fg, (_hasTheme && _theme.darkMode) ? 24u : 18u),
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

                        IDWriteTextLayout* spanLayout = GetTextSegmentLayout(visibleStart, visibleEnd, spanRc.right - spanRc.left);
                        if (! spanLayout)
                        {
                            return;
                        }
                        float markerLocalX = 0.0f;
                        float markerLocalY = 0.0f;
                        DWRITE_HIT_TEST_METRICS markerMetrics{};
                        if (FAILED(HitTestCaretPosition(
                                spanLayout, visibleEnd - visibleStart, markerIndex - visibleStart, markerLocalX, markerLocalY, markerMetrics)))
                        {
                            return;
                        }
                        const float markerX = spanLeft + markerLocalX;
                        if (markerX >= spanRc.right)
                        {
                            return;
                        }

                        const D2D1_RECT_F markerRc = D2D1::RectF(markerX, y, std::min(spanRc.right, markerX + std::max(1.0f, markerMetrics.width)), lineBottom);
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
                        if (! hasSearchHighlights || searchLen == 0 || visibleEnd < visibleStart || spanRc.right <= spanRc.left)
                        {
                            return;
                        }

                        IDWriteTextLayout* spanLayout = GetTextSegmentLayout(visibleStart, visibleEnd, spanRc.right - spanRc.left);
                        if (! spanLayout)
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

                            float startX = 0.0f;
                            float startY = 0.0f;
                            float endX   = 0.0f;
                            float endY   = 0.0f;
                            DWRITE_HIT_TEST_METRICS startMetrics{};
                            DWRITE_HIT_TEST_METRICS endMetrics{};
                            const HRESULT startHr =
                                HitTestCaretPosition(spanLayout, visibleEnd - visibleStart, hlStart - visibleStart, startX, startY, startMetrics);
                            const HRESULT endHr = HitTestCaretPosition(spanLayout, visibleEnd - visibleStart, hlEnd - visibleStart, endX, endY, endMetrics);
                            if (FAILED(startHr) || FAILED(endHr))
                            {
                                continue;
                            }
                            const float hlLeft     = spanLeft + std::min(startX, endX);
                            const float hlRight    = spanLeft + std::max(startX, endX);
                            const D2D1_RECT_F hlRc = D2D1::RectF(std::max(spanRc.left, hlLeft), y, std::min(spanRc.right, hlRight), lineBottom);

                            _textViewBrush->SetColor(ColorFFromColorRef(searchBg));
                            _textViewTarget->FillRectangle(hlRc, _textViewBrush.get());
                            _textViewBrush->SetColor(ColorFFromColorRef(fg));
                        }
                    };
                    const auto drawSelectionForSpan = [&](size_t visibleStart, size_t visibleEnd, float spanLeft, const D2D1_RECT_F& spanRc) noexcept
                    {
                        if (! hasSelection || visibleEnd < visibleStart || spanRc.right <= spanRc.left)
                        {
                            return;
                        }

                        const size_t hlStart = std::max(selStartIndex, visibleStart);
                        const size_t hlEnd   = std::min(selEndIndex, visibleEnd);
                        if (hlEnd <= hlStart)
                        {
                            return;
                        }

                        IDWriteTextLayout* spanLayout = GetTextSegmentLayout(visibleStart, visibleEnd, spanRc.right - spanRc.left);
                        if (! spanLayout)
                        {
                            return;
                        }
                        float startX = 0.0f;
                        float startY = 0.0f;
                        float endX   = 0.0f;
                        float endY   = 0.0f;
                        DWRITE_HIT_TEST_METRICS startMetrics{};
                        DWRITE_HIT_TEST_METRICS endMetrics{};
                        const HRESULT startHr =
                            HitTestCaretPosition(spanLayout, visibleEnd - visibleStart, hlStart - visibleStart, startX, startY, startMetrics);
                        const HRESULT endHr = HitTestCaretPosition(spanLayout, visibleEnd - visibleStart, hlEnd - visibleStart, endX, endY, endMetrics);
                        if (FAILED(startHr) || FAILED(endHr))
                        {
                            return;
                        }
                        const float hlLeft     = spanLeft + std::min(startX, endX);
                        const float hlRight    = spanLeft + std::max(startX, endX);
                        const D2D1_RECT_F hlRc = D2D1::RectF(std::max(spanRc.left, hlLeft), y, std::min(spanRc.right, hlRight), lineBottom);

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

                        IDWriteTextLayout* spanLayout = GetTextSegmentLayout(visibleStart, visibleEnd, spanRc.right - spanRc.left);
                        if (! spanLayout)
                        {
                            return;
                        }
                        _textViewTarget->DrawTextLayout(D2D1::Point2F(spanRc.left, spanRc.top), spanLayout, _textViewBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    };
                    const auto drawCaretForSpan = [&](size_t visibleStart, size_t visibleEnd, float spanLeft, const D2D1_RECT_F& spanRc) noexcept
                    {
                        if (! hasFocus || _textCaretIndex < visibleStart || _textCaretIndex > visibleEnd || spanRc.right <= spanRc.left)
                        {
                            return;
                        }

                        IDWriteTextLayout* spanLayout = GetTextSegmentLayout(visibleStart, visibleEnd, spanRc.right - spanRc.left);
                        if (! spanLayout)
                        {
                            return;
                        }
                        float localX = 0.0f;
                        float localY = 0.0f;
                        DWRITE_HIT_TEST_METRICS metrics{};
                        const HRESULT caretHr =
                            HitTestCaretPosition(spanLayout, visibleEnd - visibleStart, _textCaretIndex - visibleStart, localX, localY, metrics);
                        if (FAILED(caretHr))
                        {
                            return;
                        }
                        const float caretX        = std::min(spanRc.right, spanLeft + localX);
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
        if (! Common::Clipboard::TrySetUnicodeText(root, selected))
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
        _textPreferredXValid = false;
        _textVerticalCaretHistory.clear();
        CommandGoToTop(root, shift);
        return 0;
    }

    if (vk == VK_END)
    {
        _textPreferredXValid = false;
        _textVerticalCaretHistory.clear();
        CommandGoToBottom(root, shift);
        return 0;
    }

    if (TextVisualLineCount() == 0u || _textLineStarts.empty() || _textLineEnds.empty())
    {
        return DefWindowProcW(hwnd, WM_KEYDOWN, vk, lParam);
    }

    static_cast<void>(EnsureTextViewDirect2D(hwnd));
    if (_textSparseWrapActive)
    {
        RECT client{};
        GetClientRect(hwnd, &client);
        const float heightDip =
            std::max(1.0f, Common::WindowSizing::PixelToDip(static_cast<float>(client.bottom - client.top), static_cast<float>(GetDpiForWindow(hwnd))));
        const float lineH        = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
        const size_t visibleRows = std::max<size_t>(2u, static_cast<size_t>(std::ceil(heightDip / lineH)) + 1u);
        RebuildSparseTextViewportLayouts(hwnd, visibleRows);
    }

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
        const uint32_t totalVisual = static_cast<uint32_t>(std::min<uint64_t>(TextVisualLineCount(), std::numeric_limits<uint32_t>::max()));
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

        if (_textSparseWrapActive)
        {
            const auto anchor = std::find(_textSparseViewportAnchors.begin(), _textSparseViewportAnchors.end(), caretVisual);
            if (anchor == _textSparseViewportAnchors.end())
            {
                _textTopVisualLine = caretVisual;
            }
            else
            {
                const size_t caretRow = static_cast<size_t>(std::distance(_textSparseViewportAnchors.begin(), anchor));
                if (caretRow >= page)
                {
                    _textTopVisualLine = static_cast<uint32_t>(_textSparseViewportAnchors[caretRow - page + 1u]);
                }
            }
        }
        else if (caretVisual < _textTopVisualLine)
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
            TextVisualLineLayoutEntry caretLayoutValue{};
            const TextVisualLineLayoutEntry* caretLayout = TryGetTextVisualLineLayout(caretVisual, caretLayoutValue) ? &caretLayoutValue : nullptr;
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
        newCaret = NormalizeTextPosition(std::min(newCaret, _textBuffer.size()));

        _textCaretIndex = newCaret;
        if (! shift)
        {
            _textSelAnchor = newCaret;
        }
        _textSelActive = newCaret;
    };

    const uint32_t currentVisual = findVisualForIndex();
    TextVisualLineLayoutEntry currentLayoutValue{};
    const TextVisualLineLayoutEntry* currentLayout = TryGetTextVisualLineLayout(currentVisual, currentLayoutValue) ? &currentLayoutValue : nullptr;
    const bool preferRightPane = currentLayout && currentLayout->splitPanes &&
                                 (_textCaretIndex >= currentLayout->rightStartIndex ||
                                  (_textCaretIndex >= currentLayout->separatorStartIndex && currentLayout->rightEndIndex > currentLayout->rightStartIndex));
    uint32_t segStart          = 0;
    uint32_t segEnd            = 0;
    static_cast<void>(getSegmentBounds(currentVisual, segStart, segEnd, preferRightPane));

    const size_t segStartSize     = std::min<size_t>(static_cast<size_t>(segStart), _textBuffer.size());
    const size_t segEndSize       = std::min<size_t>(static_cast<size_t>(segEnd), _textBuffer.size());
    const bool verticalNavigation = vk == VK_UP || vk == VK_DOWN || vk == VK_PRIOR || vk == VK_NEXT;
    RECT keyClient{};
    GetClientRect(hwnd, &keyClient);
    const float keyWidthDip =
        std::max(1.0f, Common::WindowSizing::PixelToDip(static_cast<float>(keyClient.right - keyClient.left), static_cast<float>(GetDpiForWindow(hwnd))));
    const size_t keyVisibleCodeUnits = static_cast<size_t>(std::ceil(keyWidthDip / std::max(1.0f, _textCharWidthDip))) * 4u + 16u;
    if (! verticalNavigation || ! _textPreferredXValid)
    {
        _textPreferredColumn          = (_textCaretIndex >= segStartSize) ? (_textCaretIndex - segStartSize) : 0u;
        _textPreferredXDip            = static_cast<float>(_textPreferredColumn) * std::max(1.0f, _textCharWidthDip);
        const size_t currentLayoutEnd = _wrap ? segEndSize : NormalizeTextPosition(std::min(segEndSize, segStartSize + keyVisibleCodeUnits));
        if (_textCaretIndex >= segStartSize && _textCaretIndex <= currentLayoutEnd)
        {
            if (IDWriteTextLayout* currentTextLayout = GetTextSegmentLayout(segStartSize, currentLayoutEnd, keyWidthDip))
            {
                float caretX = 0.0f;
                float caretY = 0.0f;
                DWRITE_HIT_TEST_METRICS caretMetrics{};
                if (SUCCEEDED(
                        HitTestCaretPosition(currentTextLayout, currentLayoutEnd - segStartSize, _textCaretIndex - segStartSize, caretX, caretY, caretMetrics)))
                {
                    _textPreferredXDip = caretX;
                }
            }
        }
    }

    const uint32_t totalVisual = static_cast<uint32_t>(std::min<uint64_t>(TextVisualLineCount(), std::numeric_limits<uint32_t>::max()));
    const uint32_t lastVisual  = totalVisual > 0 ? (totalVisual - 1) : 0;

    if (vk == VK_LEFT)
    {
        _textPreferredXValid = false;
        _textVerticalCaretHistory.clear();
        if (_textCaretIndex == 0 && _textStreamActive)
        {
            if (TryNavigateTextStream(root, true))
            {
                return 0;
            }
        }

        if (_textCaretIndex > 0)
        {
            setCaret(PreviousTextPosition(_textCaretIndex));
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
        _textPreferredXValid = false;
        _textVerticalCaretHistory.clear();
        if (_textCaretIndex >= _textBuffer.size() && _textStreamActive)
        {
            if (TryNavigateTextStream(root, false))
            {
                return 0;
            }
        }

        if (_textCaretIndex < _textBuffer.size())
        {
            setCaret(NextTextPosition(_textCaretIndex));
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
        const uint32_t page   = std::max<uint32_t>(1u, static_cast<uint32_t>(si.nPage == 0 ? 1u : si.nPage));
        const bool movingDown = vk == VK_DOWN || vk == VK_NEXT;
        const bool pageMove   = vk == VK_PRIOR || vk == VK_NEXT;

        if (! _textVerticalCaretHistory.empty())
        {
            const TextVerticalCaretHistoryEntry previous = _textVerticalCaretHistory.back();
            if (previous.toCaret == _textCaretIndex && previous.pageMove == pageMove && previous.movingDown != movingDown)
            {
                _textVerticalCaretHistory.pop_back();
                setCaret(previous.fromCaret);
                _textPreferredXValid = true;
                ensureCaretVisible();
                InvalidateRect(hwnd, nullptr, TRUE);
                if (_hWnd)
                {
                    InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
                }
                return 0;
            }
            if (previous.toCaret != _textCaretIndex || previous.pageMove != pageMove)
            {
                _textVerticalCaretHistory.clear();
            }
        }

        uint32_t targetVisual   = currentVisual;
        const auto sparseAnchor = _textSparseWrapActive ? std::find(_textSparseViewportAnchors.begin(), _textSparseViewportAnchors.end(), currentVisual)
                                                        : _textSparseViewportAnchors.end();
        const size_t sparseRow =
            sparseAnchor != _textSparseViewportAnchors.end() ? static_cast<size_t>(std::distance(_textSparseViewportAnchors.begin(), sparseAnchor)) : 0u;
        if (_textSparseWrapActive && sparseAnchor != _textSparseViewportAnchors.end())
        {
            if (vk == VK_UP || vk == VK_PRIOR)
            {
                if (currentVisual == 0u && _textStreamActive && TryNavigateTextStream(root, true))
                {
                    return 0;
                }
                const size_t rows = vk == VK_UP ? 1u : static_cast<size_t>(page);
                if (sparseRow >= rows)
                {
                    targetVisual = static_cast<uint32_t>(_textSparseViewportAnchors[sparseRow - rows]);
                }
                else if (! _textSparseTopHistory.empty())
                {
                    targetVisual = _textSparseTopHistory.back();
                    _textSparseTopHistory.pop_back();
                }
                else
                {
                    targetVisual = currentVisual > rows ? (currentVisual - static_cast<uint32_t>(rows)) : 0u;
                }
            }
            else
            {
                if (currentVisual >= lastVisual && _textStreamActive && TryNavigateTextStream(root, false))
                {
                    return 0;
                }
                const size_t rows      = vk == VK_DOWN ? 1u : static_cast<size_t>(page);
                const size_t targetRow = std::min<size_t>(sparseRow + rows, _textSparseViewportAnchors.size() - 1u);
                targetVisual           = static_cast<uint32_t>(_textSparseViewportAnchors[targetRow]);
            }
        }
        else if (vk == VK_UP)
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

        size_t targetCaret           = targetStartSize + std::min<size_t>(_textPreferredColumn, targetLen);
        const size_t targetLayoutEnd = _wrap ? targetEndSize : NormalizeTextPosition(std::min(targetEndSize, targetStartSize + keyVisibleCodeUnits));
        if (IDWriteTextLayout* targetTextLayout = GetTextSegmentLayout(targetStartSize, targetLayoutEnd, keyWidthDip))
        {
            BOOL trailing = FALSE;
            BOOL inside   = FALSE;
            DWRITE_HIT_TEST_METRICS targetMetrics{};
            if (SUCCEEDED(targetTextLayout->HitTestPoint(_textPreferredXDip, std::max(1.0f, _textLineHeightDip) * 0.5f, &trailing, &inside, &targetMetrics)))
            {
                const size_t local = static_cast<size_t>(targetMetrics.textPosition) + (trailing != FALSE ? static_cast<size_t>(targetMetrics.length) : 0u);
                targetCaret        = std::min(targetLayoutEnd, targetStartSize + local);
            }
        }
        if (targetCaret != _textCaretIndex)
        {
            if (_textVerticalCaretHistory.size() >= kVerticalCaretHistoryLimit)
            {
                _textVerticalCaretHistory.erase(_textVerticalCaretHistory.begin());
            }
            _textVerticalCaretHistory.push_back(TextVerticalCaretHistoryEntry{
                .fromCaret  = _textCaretIndex,
                .toCaret    = targetCaret,
                .pageMove   = pageMove,
                .movingDown = movingDown,
            });
        }
        setCaret(targetCaret);
        _textPreferredXValid = true;

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
            return OnTextViewMouseMove(hwnd, {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))}, wp);
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

uint64_t ViewerText::TextVisualLineCount() const noexcept
{
    return _textSparseWrapActive ? _textSparseVisualLineCount : static_cast<uint64_t>(_textVisualLineLayouts.size());
}

void ViewerText::RebuildSparseTextViewportLayouts(HWND hwnd, size_t minimumRows) noexcept
{
    if (! _textSparseWrapActive || _textSparseVisualLines.empty() || _textSparseVisualLineCount == 0u)
    {
        _textSparseViewportLayouts.clear();
        _textSparseViewportAnchors.clear();
        _textSparseViewportCheckpoints.clear();
        _textSparseViewportRequestedRows = 0u;
        _textSparseViewportComplete      = false;
        return;
    }

    static_cast<void>(EnsureTextViewDirect2D(hwnd));
    minimumRows        = std::clamp(minimumRows, static_cast<size_t>(1u), kSparseViewportLayoutLimit);
    const uint64_t top = std::min<uint64_t>(_textTopVisualLine, _textSparseVisualLineCount - 1u);
    if (_textSparseViewportTop == top && _textSparseViewportRequestedRows >= minimumRows &&
        (_textSparseViewportLayouts.size() >= minimumRows || _textSparseViewportComplete))
    {
        return;
    }

    _textSparseViewportLayouts.clear();
    _textSparseViewportAnchors.clear();
    _textSparseViewportCheckpoints.clear();
    _textSparseViewportTop           = top;
    _textSparseViewportRequestedRows = minimumRows;
    _textSparseViewportComplete      = false;
    _textSparseViewportLayouts.reserve(minimumRows);
    _textSparseViewportAnchors.reserve(minimumRows);
    _textSparseViewportCheckpoints.reserve(minimumRows);

    const auto after    = std::upper_bound(_textSparseVisualLines.begin(),
                                           _textSparseVisualLines.end(),
                                           top,
                                           [](uint64_t value, const SparseTextVisualLineSummary& summary) noexcept { return value < summary.firstVisualLine; });
    size_t summaryIndex = static_cast<size_t>(
        std::distance(_textSparseVisualLines.begin(), after == _textSparseVisualLines.begin() ? _textSparseVisualLines.begin() : std::prev(after)));
    uint64_t rowInSummary = top - _textSparseVisualLines[summaryIndex].firstVisualLine;
    uint32_t plainCursor  = 0u;
    uint32_t leftCursor   = 0u;
    uint32_t rightCursor  = 0u;
    uint64_t rowOrdinal   = rowInSummary;

    const auto initializeCursors = [&](size_t targetSummaryIndex, const SparseTextVisualLineSummary& summary, uint64_t rowOffset) noexcept
    {
        rowOrdinal                  = rowOffset;
        const uint64_t targetVisual = summary.firstVisualLine + rowOffset;
        const auto cached =
            std::find_if(_textSparseCheckpointCache.begin(), _textSparseCheckpointCache.end(), [&](const SparseTextViewportCheckpoint& checkpoint) noexcept {
            return checkpoint.visualLine == targetVisual && checkpoint.summaryIndex == targetSummaryIndex;
        });
        if (cached != _textSparseCheckpointCache.end())
        {
            rowOrdinal  = cached->rowOrdinal;
            plainCursor = cached->plainCursor;
            leftCursor  = cached->leftCursor;
            rightCursor = cached->rightCursor;
            return;
        }

        if (summary.splitPanes)
        {
            const uint64_t leftLength  = static_cast<uint64_t>(summary.leftEndIndex - summary.leftStartIndex);
            const uint64_t rightLength = static_cast<uint64_t>(summary.rightEndIndex - summary.rightStartIndex);
            leftCursor                 = static_cast<uint32_t>(
                NormalizeTextSegmentStart(static_cast<size_t>(summary.leftStartIndex) + static_cast<size_t>(std::min(rowOffset, leftLength))));
            rightCursor = static_cast<uint32_t>(
                NormalizeTextSegmentStart(static_cast<size_t>(summary.rightStartIndex) + static_cast<size_t>(std::min(rowOffset, rightLength))));
        }
        else
        {
            const uint64_t length = static_cast<uint64_t>(summary.lineEndIndex - summary.lineStartIndex);
            plainCursor           = static_cast<uint32_t>(
                NormalizeTextSegmentStart(static_cast<size_t>(summary.lineStartIndex) + static_cast<size_t>(std::min(rowOffset, length))));
        }
    };
    initializeCursors(summaryIndex, _textSparseVisualLines[summaryIndex], rowInSummary);

    const auto rememberCheckpoint = [&](const SparseTextViewportCheckpoint& checkpoint)
    {
        const auto existing =
            std::find_if(_textSparseCheckpointCache.begin(), _textSparseCheckpointCache.end(), [&](const SparseTextViewportCheckpoint& candidate) noexcept {
            return candidate.visualLine == checkpoint.visualLine;
        });
        if (existing != _textSparseCheckpointCache.end())
        {
            *existing = checkpoint;
            return;
        }
        if (_textSparseCheckpointCache.size() >= kSparseCheckpointCacheLimit)
        {
            _textSparseCheckpointCache.erase(_textSparseCheckpointCache.begin());
        }
        _textSparseCheckpointCache.push_back(checkpoint);
    };

    const float charWidth = std::max(1.0f, _textCharWidthDip);
    while (_textSparseViewportLayouts.size() < minimumRows && summaryIndex < _textSparseVisualLines.size())
    {
        const SparseTextVisualLineSummary& summary = _textSparseVisualLines[summaryIndex];
        TextVisualLineLayoutEntry layout{};
        layout.logicalLine               = summary.logicalLine;
        bool summaryComplete             = false;
        const uint64_t boundedRowOrdinal = std::min<uint64_t>(rowOrdinal, static_cast<uint64_t>(summary.visualLineCount - 1u));
        const SparseTextViewportCheckpoint checkpoint{
            .visualLine   = summary.firstVisualLine + boundedRowOrdinal,
            .summaryIndex = summaryIndex,
            .rowOrdinal   = boundedRowOrdinal,
            .plainCursor  = plainCursor,
            .leftCursor   = leftCursor,
            .rightCursor  = rightCursor,
        };

        if (summary.splitPanes)
        {
            leftCursor  = std::clamp(leftCursor, summary.leftStartIndex, summary.leftEndIndex);
            rightCursor = std::clamp(rightCursor, summary.rightStartIndex, summary.rightEndIndex);
            const uint32_t leftSegmentEnd =
                static_cast<uint32_t>(FindTextSegmentEnd(leftCursor, summary.leftEndIndex, static_cast<float>(summary.leftPaneColumns) * charWidth));
            const uint32_t rightSegmentEnd =
                static_cast<uint32_t>(FindTextSegmentEnd(rightCursor, summary.rightEndIndex, static_cast<float>(summary.rightPaneColumns) * charWidth));
            layout.splitPanes          = true;
            layout.leftStartIndex      = leftCursor;
            layout.leftEndIndex        = std::max(leftCursor, leftSegmentEnd);
            layout.rightStartIndex     = rightCursor;
            layout.rightEndIndex       = std::max(rightCursor, rightSegmentEnd);
            layout.separatorStartIndex = summary.separatorStartIndex;
            layout.separatorEndIndex   = summary.separatorEndIndex;
            layout.leftPaneColumns     = summary.leftPaneColumns;
            layout.rightPaneColumns    = summary.rightPaneColumns;
            layout.separatorColumns    = summary.separatorColumns;
            const bool hasLeft         = layout.leftEndIndex > layout.leftStartIndex;
            const bool hasRight        = layout.rightEndIndex > layout.rightStartIndex;
            layout.segmentStartIndex   = hasLeft ? layout.leftStartIndex : (hasRight ? layout.rightStartIndex : summary.leftEndIndex);
            layout.segmentEndIndex     = hasRight ? layout.rightEndIndex : (hasLeft ? layout.leftEndIndex : summary.leftEndIndex);

            leftCursor      = layout.leftEndIndex;
            rightCursor     = layout.rightEndIndex;
            summaryComplete = leftCursor >= summary.leftEndIndex && rightCursor >= summary.rightEndIndex;
        }
        else
        {
            plainCursor               = std::clamp(plainCursor, summary.lineStartIndex, summary.lineEndIndex);
            const uint32_t segmentEnd = static_cast<uint32_t>(FindTextSegmentEnd(plainCursor, summary.lineEndIndex, _textWrapWidthDip));
            layout.segmentStartIndex  = plainCursor;
            layout.segmentEndIndex    = std::max(plainCursor, segmentEnd);
            plainCursor               = layout.segmentEndIndex;
            summaryComplete           = plainCursor >= summary.lineEndIndex;
        }

        _textSparseViewportAnchors.push_back(checkpoint.visualLine);
        _textSparseViewportCheckpoints.push_back(checkpoint);
        rememberCheckpoint(checkpoint);
        _textSparseViewportLayouts.push_back(layout);
        rowOrdinal = boundedRowOrdinal + 1u;

        if (! summaryComplete)
        {
            continue;
        }

        ++summaryIndex;
        rowInSummary = 0u;
        if (summaryIndex < _textSparseVisualLines.size())
        {
            initializeCursors(summaryIndex, _textSparseVisualLines[summaryIndex], rowInSummary);
        }
    }

    _textSparseViewportComplete = summaryIndex >= _textSparseVisualLines.size();
}

bool ViewerText::TryGetTextVisualLineLayout(uint64_t visualLine, TextVisualLineLayoutEntry& layoutOut) noexcept
{
    if (! _textSparseWrapActive)
    {
        if (_textVisualLineLayouts.empty())
        {
            return false;
        }
        const size_t index = static_cast<size_t>(std::min<uint64_t>(visualLine, _textVisualLineLayouts.size() - 1u));
        layoutOut          = _textVisualLineLayouts[index];
        return true;
    }

    if (_textSparseVisualLines.empty() || _textSparseVisualLineCount == 0u)
    {
        return false;
    }

    visualLine              = std::min<uint64_t>(visualLine, _textSparseVisualLineCount - 1u);
    const auto cachedAnchor = std::find(_textSparseViewportAnchors.begin(), _textSparseViewportAnchors.end(), visualLine);
    if (cachedAnchor != _textSparseViewportAnchors.end())
    {
        const size_t viewportRow = static_cast<size_t>(std::distance(_textSparseViewportAnchors.begin(), cachedAnchor));
        layoutOut                = _textSparseViewportLayouts[viewportRow];
        return true;
    }

    const auto after = std::upper_bound(_textSparseVisualLines.begin(),
                                        _textSparseVisualLines.end(),
                                        visualLine,
                                        [](uint64_t value, const SparseTextVisualLineSummary& summary) noexcept { return value < summary.firstVisualLine; });
    const auto it    = after == _textSparseVisualLines.begin() ? _textSparseVisualLines.begin() : std::prev(after);
    const uint64_t rowInLine = visualLine - it->firstVisualLine;
    if (it->splitPanes)
    {
        const uint32_t leftLength  = it->leftEndIndex - it->leftStartIndex;
        const uint32_t rightLength = it->rightEndIndex - it->rightStartIndex;
        const uint32_t leftOffset  = static_cast<uint32_t>(std::min<uint64_t>(rowInLine, leftLength));
        const uint32_t rightOffset = static_cast<uint32_t>(std::min<uint64_t>(rowInLine, rightLength));
        const uint32_t leftStart   = static_cast<uint32_t>(NormalizeTextSegmentStart(it->leftStartIndex + leftOffset));
        const uint32_t rightStart  = static_cast<uint32_t>(NormalizeTextSegmentStart(it->rightStartIndex + rightOffset));
        const uint32_t leftEnd =
            static_cast<uint32_t>(FindTextSegmentEnd(leftStart, it->leftEndIndex, static_cast<float>(it->leftPaneColumns) * std::max(1.0f, _textCharWidthDip)));
        const uint32_t rightEnd = static_cast<uint32_t>(
            FindTextSegmentEnd(rightStart, it->rightEndIndex, static_cast<float>(it->rightPaneColumns) * std::max(1.0f, _textCharWidthDip)));
        layoutOut.logicalLine         = it->logicalLine;
        layoutOut.segmentStartIndex   = leftEnd > leftStart ? leftStart : (rightEnd > rightStart ? rightStart : it->leftEndIndex);
        layoutOut.segmentEndIndex     = rightEnd > rightStart ? rightEnd : (leftEnd > leftStart ? leftEnd : it->leftEndIndex);
        layoutOut.splitPanes          = true;
        layoutOut.leftStartIndex      = leftStart;
        layoutOut.leftEndIndex        = leftEnd;
        layoutOut.rightStartIndex     = rightStart;
        layoutOut.rightEndIndex       = rightEnd;
        layoutOut.separatorStartIndex = it->separatorStartIndex;
        layoutOut.separatorEndIndex   = it->separatorEndIndex;
        layoutOut.leftPaneColumns     = it->leftPaneColumns;
        layoutOut.rightPaneColumns    = it->rightPaneColumns;
        layoutOut.separatorColumns    = it->separatorColumns;
        return true;
    }
    const uint64_t rawStart = static_cast<uint64_t>(it->lineStartIndex) + rowInLine;
    const uint32_t start    = static_cast<uint32_t>(std::min<uint64_t>(NormalizeTextSegmentStart(static_cast<size_t>(rawStart)), it->lineEndIndex));
    const uint32_t end      = static_cast<uint32_t>(FindTextSegmentEnd(start, it->lineEndIndex, _textWrapWidthDip));
    layoutOut               = TextVisualLineLayoutEntry{
        .logicalLine       = it->logicalLine,
        .segmentStartIndex = start,
        .segmentEndIndex   = std::max(start, end),
    };
    return true;
}

bool ViewerText::FindFirstTextVisualLineForLogical(uint32_t logicalLine, uint32_t& visualLineOut) const noexcept
{
    if (_textSparseWrapActive)
    {
        if (logicalLine >= _textSparseVisualLines.size())
        {
            visualLineOut = 0u;
            return false;
        }
        visualLineOut = static_cast<uint32_t>(std::min<uint64_t>(_textSparseVisualLines[logicalLine].firstVisualLine, std::numeric_limits<uint32_t>::max()));
        return true;
    }

    const auto it = std::lower_bound(_textVisualLineLogical.begin(), _textVisualLineLogical.end(), logicalLine);
    if (it == _textVisualLineLogical.end())
    {
        visualLineOut = _textVisualLineLogical.empty() ? 0u : static_cast<uint32_t>(_textVisualLineLogical.size() - 1u);
        return ! _textVisualLineLogical.empty();
    }
    visualLineOut = static_cast<uint32_t>(std::distance(_textVisualLineLogical.begin(), it));
    return true;
}

bool ViewerText::FindTextVisualLineForIndex(size_t index, uint32_t& visualLineOut) const noexcept
{
    if (_textSparseWrapActive)
    {
        if (_textSparseVisualLines.empty())
        {
            visualLineOut = 0u;
            return false;
        }

        for (size_t row = 0u; row < _textSparseViewportLayouts.size(); ++row)
        {
            const TextVisualLineLayoutEntry& layout = _textSparseViewportLayouts[row];
            const bool inPlain                      = ! layout.splitPanes && index >= layout.segmentStartIndex && index <= layout.segmentEndIndex;
            const bool inLeft                       = layout.splitPanes && index >= layout.leftStartIndex && index <= layout.leftEndIndex;
            const bool inRight                      = layout.splitPanes && index >= layout.rightStartIndex && index <= layout.rightEndIndex;
            const bool inSeparator                  = layout.splitPanes && index >= layout.separatorStartIndex && index <= layout.separatorEndIndex;
            if (inPlain || inLeft || inRight || inSeparator)
            {
                visualLineOut = static_cast<uint32_t>(std::min<uint64_t>(_textSparseViewportAnchors[row], std::numeric_limits<uint32_t>::max()));
                return true;
            }
        }

        index                       = NormalizeTextPosition(index);
        const uint32_t clampedIndex = static_cast<uint32_t>(std::min<size_t>(index, std::numeric_limits<uint32_t>::max()));
        const auto after = std::upper_bound(_textSparseVisualLines.begin(),
                                            _textSparseVisualLines.end(),
                                            clampedIndex,
                                            [](uint32_t value, const SparseTextVisualLineSummary& summary) noexcept { return value < summary.lineStartIndex; });
        const auto it    = after == _textSparseVisualLines.begin() ? _textSparseVisualLines.begin() : std::prev(after);
        uint64_t row     = 0u;
        if (it->splitPanes)
        {
            if (clampedIndex >= it->rightStartIndex)
            {
                row = static_cast<uint64_t>(clampedIndex - it->rightStartIndex);
            }
            else if (clampedIndex <= it->leftEndIndex)
            {
                row = static_cast<uint64_t>(clampedIndex - it->leftStartIndex);
            }
        }
        else
        {
            const uint64_t column = clampedIndex > it->lineStartIndex ? static_cast<uint64_t>(clampedIndex - it->lineStartIndex) : 0u;
            row                   = column;
        }
        row           = std::min<uint64_t>(it->visualLineCount - 1u, row);
        visualLineOut = static_cast<uint32_t>(std::min<uint64_t>(it->firstVisualLine + row, std::numeric_limits<uint32_t>::max()));
        return true;
    }

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
    uint32_t visualLine, uint32_t& logicalLineOut, uint32_t& segmentStartOut, uint32_t& segmentEndOut, bool preferRightPane) noexcept
{
    if (TextVisualLineCount() == 0u || _textLineStarts.empty())
    {
        logicalLineOut  = 0u;
        segmentStartOut = 0u;
        segmentEndOut   = 0u;
        return false;
    }

    TextVisualLineLayoutEntry layout{};
    if (! TryGetTextVisualLineLayout(visualLine, layout))
    {
        return false;
    }
    logicalLineOut = std::min<uint32_t>(layout.logicalLine, static_cast<uint32_t>(_textLineStarts.size() - 1u));

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
    ClearTextLayoutCache();
    _textVisualLineStarts.clear();
    _textVisualLineLogical.clear();
    _textVisualLineLayouts.clear();
    _textSparseVisualLines.clear();
    _textSparseVisualLineCount      = 0u;
    _textSparseWrapActive           = false;
    _textWrapColumns                = 0;
    _textWrapWidthDip               = 0.0f;
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
        const float widthDip  = std::max(0.0f, Common::WindowSizing::PixelToDip(static_cast<float>(client.right - client.left), static_cast<float>(dpi)));
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
        _textWrapWidthDip = std::max(1.0f, availDip);
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
        const float widthDip  = std::max(0.0f, Common::WindowSizing::PixelToDip(static_cast<float>(client.right - client.left), static_cast<float>(dpi)));
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
    if (hasPaneLocalLayout)
    {
        const uint32_t separatorColumns = 3u;
        const uint32_t paneBudget       = availableCols > separatorColumns ? (availableCols - separatorColumns) : 2u;
        _textSideBySideLeftPaneColumns  = std::max<uint32_t>(1u, paneBudget / 2u);
        _textSideBySideRightPaneColumns = std::max<uint32_t>(1u, paneBudget - _textSideBySideLeftPaneColumns);
        _textSideBySideSeparatorColumns = separatorColumns;
    }
    if (_wrap)
    {
        _textSparseVisualLines.reserve(_textLineStarts.size());
        uint64_t projectedVisualLines = 0u;
        for (uint32_t line = 0u; line < static_cast<uint32_t>(_textLineStarts.size()); ++line)
        {
            const uint32_t start = _textLineStarts[line];
            const uint32_t end   = _textLineEnds.size() > line ? _textLineEnds[line] : start;
            SparseTextVisualLineSummary summary{};
            summary.logicalLine     = line;
            summary.lineStartIndex  = start;
            summary.lineEndIndex    = end;
            summary.firstVisualLine = projectedVisualLines;
            if (hasPaneLocalLayout && activeDiffVariant->logicalRowPaneLayouts[line].splitRow)
            {
                const auto& pane              = activeDiffVariant->logicalRowPaneLayouts[line];
                summary.splitPanes            = true;
                summary.leftStartIndex        = start;
                summary.leftEndIndex          = std::min<uint32_t>(end, start + pane.leftTextColumns);
                summary.separatorStartIndex   = summary.leftEndIndex;
                summary.separatorEndIndex     = std::min<uint32_t>(end, summary.separatorStartIndex + pane.separatorColumns);
                summary.rightStartIndex       = summary.separatorEndIndex;
                summary.rightEndIndex         = end;
                summary.leftPaneColumns       = _textSideBySideLeftPaneColumns;
                summary.rightPaneColumns      = _textSideBySideRightPaneColumns;
                summary.separatorColumns      = _textSideBySideSeparatorColumns;
                const uint32_t leftCodeUnits  = summary.leftEndIndex - summary.leftStartIndex;
                const uint32_t rightCodeUnits = summary.rightEndIndex - summary.rightStartIndex;
                summary.visualLineCount       = std::max<uint32_t>(1u, std::max(leftCodeUnits, rightCodeUnits));
            }
            else
            {
                const uint32_t len      = end >= start ? (end - start) : 0u;
                summary.visualLineCount = std::max<uint32_t>(1u, len);
            }
            _textSparseVisualLines.push_back(summary);
            const uint32_t count = summary.visualLineCount;
            projectedVisualLines += count;
        }

        if (projectedVisualLines > kSparseWrapMaterializationLimit)
        {
            _textSparseVisualLineCount = projectedVisualLines;
            _textSparseWrapActive      = true;
            _textTopVisualLine         = static_cast<uint32_t>(std::min<uint64_t>(_textTopVisualLine, projectedVisualLines - 1u));
            return;
        }
        _textSparseVisualLines.clear();
    }
    size_t reserveVisualCount = _textLineStarts.size();
    if (_wrap && reserveVisualCount > 0u)
    {
        reserveVisualCount *= hasPaneLocalLayout ? 3u : 2u;
    }
    _textVisualLineStarts.reserve(reserveVisualCount);
    _textVisualLineLogical.reserve(reserveVisualCount);
    _textVisualLineLayouts.reserve(reserveVisualCount);

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

                    const uint32_t leftVisibleStart =
                        static_cast<uint32_t>(std::clamp<size_t>(NormalizeTextSegmentStart(static_cast<size_t>(leftStart) + leftOffset), leftStart, leftEnd));
                    const uint32_t leftVisibleEnd = static_cast<uint32_t>(
                        std::clamp<size_t>(NormalizeTextPosition(static_cast<size_t>(leftStart) + leftOffset + leftVisible), leftVisibleStart, leftEnd));
                    const uint32_t rightVisibleStart = static_cast<uint32_t>(
                        std::clamp<size_t>(NormalizeTextSegmentStart(static_cast<size_t>(rightStart) + rightOffset), rightStart, rightEnd));
                    const uint32_t rightVisibleEnd = static_cast<uint32_t>(
                        std::clamp<size_t>(NormalizeTextPosition(static_cast<size_t>(rightStart) + rightOffset + rightVisible), rightVisibleStart, rightEnd));
                    const bool hasLeftVisible  = leftVisibleEnd > leftVisibleStart;
                    const bool hasRightVisible = rightVisibleEnd > rightVisibleStart;

                    TextVisualLineLayoutEntry layout{};
                    layout.logicalLine         = line;
                    layout.segmentStartIndex   = hasLeftVisible ? leftVisibleStart : (hasRightVisible ? rightVisibleStart : leftEnd);
                    layout.segmentEndIndex     = hasRightVisible ? rightVisibleEnd : (hasLeftVisible ? leftVisibleEnd : leftEnd);
                    layout.splitPanes          = true;
                    layout.leftStartIndex      = hasLeftVisible ? leftVisibleStart : leftEnd;
                    layout.leftEndIndex        = hasLeftVisible ? leftVisibleEnd : leftEnd;
                    layout.rightStartIndex     = hasRightVisible ? rightVisibleStart : rightEnd;
                    layout.rightEndIndex       = hasRightVisible ? rightVisibleEnd : rightEnd;
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
                uint32_t leftCursor           = leftStart;
                uint32_t rightCursor          = rightStart;
                bool emittedRow               = false;
                do
                {
                    const uint32_t leftSegmentEnd = static_cast<uint32_t>(
                        FindTextSegmentEnd(leftCursor, leftEnd, static_cast<float>(_textSideBySideLeftPaneColumns) * std::max(1.0f, _textCharWidthDip)));
                    const uint32_t rightSegmentEnd = static_cast<uint32_t>(
                        FindTextSegmentEnd(rightCursor, rightEnd, static_cast<float>(_textSideBySideRightPaneColumns) * std::max(1.0f, _textCharWidthDip)));
                    const bool hasLeft  = leftSegmentEnd > leftCursor;
                    const bool hasRight = rightSegmentEnd > rightCursor;

                    TextVisualLineLayoutEntry layout{};
                    layout.logicalLine         = line;
                    layout.segmentStartIndex   = hasLeft ? leftCursor : (hasRight ? rightCursor : leftEnd);
                    layout.segmentEndIndex     = hasRight ? rightSegmentEnd : (hasLeft ? leftSegmentEnd : leftEnd);
                    layout.splitPanes          = true;
                    layout.leftStartIndex      = leftCursor;
                    layout.leftEndIndex        = std::max(leftCursor, leftSegmentEnd);
                    layout.rightStartIndex     = rightCursor;
                    layout.rightEndIndex       = std::max(rightCursor, rightSegmentEnd);
                    layout.separatorStartIndex = separatorStart;
                    layout.separatorEndIndex   = separatorEnd;
                    layout.leftPaneColumns     = _textSideBySideLeftPaneColumns;
                    layout.rightPaneColumns    = _textSideBySideRightPaneColumns;
                    layout.separatorColumns    = _textSideBySideSeparatorColumns;
                    _textVisualLineStarts.push_back(layout.segmentStartIndex);
                    _textVisualLineLogical.push_back(line);
                    _textVisualLineLayouts.push_back(layout);
                    leftCursor  = layout.leftEndIndex;
                    rightCursor = layout.rightEndIndex;
                    emittedRow  = true;
                } while (! emittedRow || leftCursor < leftEnd || rightCursor < rightEnd);
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

        uint32_t segStart = start;
        while (segStart < end)
        {
            const uint32_t segEnd = static_cast<uint32_t>(FindTextSegmentEnd(segStart, end, _textWrapWidthDip));
            _textVisualLineStarts.push_back(segStart);
            _textVisualLineLogical.push_back(line);
            _textVisualLineLayouts.push_back(TextVisualLineLayoutEntry{
                .logicalLine       = line,
                .segmentStartIndex = segStart,
                .segmentEndIndex   = std::max(segStart, segEnd),
            });
            segStart = static_cast<uint32_t>(segEnd > segStart ? segEnd : NextTextPosition(segStart));
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
        const float heightDip = std::max(1.0f, Common::WindowSizing::PixelToDip(static_cast<float>(client.bottom - client.top), static_cast<float>(dpi)));
        const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
        pageRows              = std::max<uint32_t>(1u, static_cast<uint32_t>(std::floor(heightDip / lineH)));
    }

    uint32_t topLogical    = 0u;
    uint32_t bottomLogical = pageRows > 0u ? (pageRows - 1u) : 0u;
    if (_textSparseWrapActive && ! _textSparseVisualLines.empty())
    {
        if (_textSparseViewportTop == _textTopVisualLine && ! _textSparseViewportLayouts.empty())
        {
            topLogical             = _textSparseViewportLayouts.front().logicalLine;
            const size_t bottomRow = std::min<size_t>(static_cast<size_t>(pageRows - 1u), _textSparseViewportLayouts.size() - 1u);
            bottomLogical          = _textSparseViewportLayouts[bottomRow].logicalLine;
        }
        else
        {
            const auto logicalForVisual = [&](uint64_t visual) noexcept
            {
                const auto after =
                    std::upper_bound(_textSparseVisualLines.begin(),
                                     _textSparseVisualLines.end(),
                                     visual,
                                     [](uint64_t value, const SparseTextVisualLineSummary& summary) noexcept { return value < summary.firstVisualLine; });
                const auto it = after == _textSparseVisualLines.begin() ? _textSparseVisualLines.begin() : std::prev(after);
                return it->logicalLine;
            };
            topLogical    = logicalForVisual(_textTopVisualLine);
            bottomLogical = logicalForVisual(std::min<uint64_t>(_textSparseVisualLineCount - 1u, static_cast<uint64_t>(_textTopVisualLine) + pageRows - 1u));
        }
    }
    else if (! _textVisualLineLogical.empty())
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

    const uint64_t totalLines = std::max<uint64_t>(1u, TextVisualLineCount());
    const uint64_t maxLine    = totalLines > 0 ? (totalLines - 1) : 0;

    RECT client{};
    GetClientRect(hwnd, &client);
    const UINT dpi           = GetDpiForWindow(hwnd);
    const float heightDip    = std::max(1.0f, Common::WindowSizing::PixelToDip(static_cast<float>(client.bottom - client.top), static_cast<float>(dpi)));
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

    float widthDip    = std::max(1.0f, Common::WindowSizing::PixelToDip(static_cast<float>(client.right - client.left), static_cast<float>(dpi)));
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
    static_cast<void>(FindFirstTextVisualLineForLogical(targetLogicalLine, targetVisualLine));

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
            if (StartAsyncTextStreamLoad(hwnd, targetOffset, false))
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
    if (_textSparseWrapActive && ! _textSparseVisualLines.empty())
    {
        if (_textSparseViewportTop == _textTopVisualLine && ! _textSparseViewportLayouts.empty())
        {
            topLogicalLine = _textSparseViewportLayouts.front().logicalLine;
        }
        else
        {
            const auto after =
                std::upper_bound(_textSparseVisualLines.begin(),
                                 _textSparseVisualLines.end(),
                                 static_cast<uint64_t>(_textTopVisualLine),
                                 [](uint64_t value, const SparseTextVisualLineSummary& summary) noexcept { return value < summary.firstVisualLine; });
            const auto it  = after == _textSparseVisualLines.begin() ? _textSparseVisualLines.begin() : std::prev(after);
            topLogicalLine = it->logicalLine;
        }
    }
    else if (! _textVisualLineLogical.empty())
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
    ClearTextLayoutCache();
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
        const uint64_t totalVisual = TextVisualLineCount();
        if (totalVisual > 0u)
        {
            _textTopVisualLine = static_cast<uint32_t>(std::min<uint64_t>(_textTopVisualLine, totalVisual - 1u));
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
        const uint64_t totalVisual = TextVisualLineCount();
        if (totalVisual > 0u)
        {
            _textTopVisualLine = static_cast<uint32_t>(std::min<uint64_t>(_textTopVisualLine, totalVisual - 1u));
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

        _textSelAnchor  = NormalizeTextPosition(matchStart);
        _textSelActive  = NormalizeTextPosition(matchEnd);
        _textCaretIndex = _textSelActive;

        auto ensureCaretVisible = [&]() noexcept
        {
            if (TextVisualLineCount() == 0u)
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
                TextVisualLineLayoutEntry caretLayoutValue{};
                const TextVisualLineLayoutEntry* caretLayout = TryGetTextVisualLineLayout(caretVisual, caretLayoutValue) ? &caretLayoutValue : nullptr;
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
            static_cast<void>(StartAsyncTextStreamLoad(hwnd, lastStart, true));
        }
        else
        {
            static_cast<void>(StartAsyncTextStreamLoad(hwnd, _textStreamSkipBytes, false));
        }

        return;
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
    _textSparseVisualLines.clear();
    _textSparseVisualLineCount = 0u;
    _textSparseWrapActive      = false;
    ClearTextLayoutCache();
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

    const HRESULT seekHr = ViewerTextSafety::SeekExact(_fileReader.get(), clampedStart);
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
    const HRESULT readHr = ViewerTextSafety::ReadExactly(_fileReader.get(), bytes);
    if (FAILED(readHr))
    {
        Debug::Error(L"ViewerText: exact read failed for '{}' (hr=0x{:08X}).", _currentPath.c_str(), static_cast<unsigned long>(readHr));
        return readHr;
    }
    const size_t bytesReadTotal = bytes.size();

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
        carryBytes = ViewerTextSafety::IncompleteUtf8TailSize(bytes.data(), bytes.size());
    }
    else if (ViewerTextSafety::UsesDbcsBoundaryCarry(displayCodePage))
    {
        carryBytes = ViewerTextSafety::IncompleteDbcsTailSize(bytes.data(), bytes.size(), displayCodePage);
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

    const uint64_t totalVisual = TextVisualLineCount();
    if (scrollToEnd && totalVisual > 0u)
    {
        _textTopVisualLine = static_cast<uint32_t>(std::min<uint64_t>(totalVisual - 1u, std::numeric_limits<uint32_t>::max()));
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

bool ViewerText::StartAsyncTextStreamLoad(HWND hwnd, uint64_t startOffset, bool scrollToEnd) noexcept
{
    if (! hwnd || ! _hEdit || ! _fileSystem || _currentPath.empty() || _fileSize == 0u || _windowIdentity == 0u)
    {
#ifdef _DEBUG
        _debugTextStreamRejectedCount += 1u;
#endif
        return false;
    }

    const uint64_t clampedStart = AlignTextStreamOffset(std::min<uint64_t>(std::max(startOffset, _textStreamSkipBytes), _fileSize));
    if (clampedStart > static_cast<uint64_t>(std::numeric_limits<__int64>::max()))
    {
#ifdef _DEBUG
        _debugTextStreamRejectedCount += 1u;
#endif
        return false;
    }

    const uint64_t requestId           = _asyncTextStreamRequestId.fetch_add(1u, std::memory_order_relaxed) + 1u;
    const uint64_t windowIdentity      = _windowIdentity;
    _activeAsyncTextStreamRequestId    = requestId;
    _pendingTextStreamStartOffset      = clampedStart;
    _textStreamLoadPending             = true;
    AsyncTextStreamFault injectedFault = AsyncTextStreamFault::None;
#if defined(_DEBUG) && defined(ENABLE_TESTS)
    injectedFault                  = _debugNextAsyncTextStreamFault;
    _debugNextAsyncTextStreamFault = AsyncTextStreamFault::None;
#endif
    const uint64_t chunkBytes          = TextStreamChunkBytes();
    const FileEncoding displayEncoding = DisplayEncodingFileEncoding();
    const UINT displayCodePage         = DisplayEncodingCodePage();

    struct AsyncTextStreamWorkItem final
    {
        AsyncTextStreamWorkItem()                                          = default;
        ~AsyncTextStreamWorkItem()                                         = default;
        AsyncTextStreamWorkItem(const AsyncTextStreamWorkItem&)            = delete;
        AsyncTextStreamWorkItem& operator=(const AsyncTextStreamWorkItem&) = delete;
        AsyncTextStreamWorkItem(AsyncTextStreamWorkItem&&)                 = delete;
        AsyncTextStreamWorkItem& operator=(AsyncTextStreamWorkItem&&)      = delete;

        wil::unique_hmodule moduleKeepAlive;
        ViewerText* viewer                             = nullptr;
        HWND hwnd                                      = nullptr;
        uint64_t windowIdentity                        = 0u;
        uint64_t fileSize                              = 0u;
        uint64_t chunkBytes                            = 0u;
        uint64_t streamSkipBytes                       = 0u;
        ViewerText::FileEncoding encoding              = ViewerText::FileEncoding::Unknown;
        UINT codePage                                  = 0u;
        ViewerText::AsyncTextStreamFault injectedFault = ViewerText::AsyncTextStreamFault::None;
        wil::com_ptr<IFileSystem> fileSystem;
        std::filesystem::path path;
        std::unique_ptr<AsyncTextStreamResult> result;
    };

    auto result = injectedFault == AsyncTextStreamFault::Allocation ? std::unique_ptr<AsyncTextStreamResult>{}
                                                                    : std::unique_ptr<AsyncTextStreamResult>(new (std::nothrow) AsyncTextStreamResult{});
    auto work   = injectedFault == AsyncTextStreamFault::Allocation ? std::unique_ptr<AsyncTextStreamWorkItem>{}
                                                                    : std::unique_ptr<AsyncTextStreamWorkItem>(new (std::nothrow) AsyncTextStreamWorkItem{});
    if (! result || ! work)
    {
#ifdef _DEBUG
        _debugTextStreamRejectedCount += 1u;
#endif
        OnAsyncTextStreamFailure(requestId, windowIdentity, E_OUTOFMEMORY);
        return false;
    }

    result->viewer         = this;
    result->requestId      = requestId;
    result->windowIdentity = windowIdentity;
    result->startOffset    = clampedStart;
    result->scrollToEnd    = scrollToEnd;
    result->hr             = E_FAIL;

    work->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kTextStreamModuleAnchor);
    if (! work->moduleKeepAlive)
    {
        const DWORD error = GetLastError();
#ifdef _DEBUG
        _debugTextStreamRejectedCount += 1u;
#endif
        OnAsyncTextStreamFailure(requestId, windowIdentity, HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_MOD_NOT_FOUND));
        return false;
    }

    work->viewer          = this;
    work->hwnd            = hwnd;
    work->windowIdentity  = windowIdentity;
    work->fileSize        = _fileSize;
    work->chunkBytes      = chunkBytes;
    work->streamSkipBytes = _textStreamSkipBytes;
    work->encoding        = displayEncoding;
    work->codePage        = displayCodePage;
    work->injectedFault   = injectedFault;
    work->fileSystem      = _fileSystem;
    work->path            = _currentPath;
    work->result          = std::move(result);

    AddRef();
    const BOOL submitted = injectedFault == AsyncTextStreamFault::Submit ? FALSE
                                                                         : TrySubmitThreadpoolCallback(
                                                                               [](PTP_CALLBACK_INSTANCE instance, void* context) noexcept
    {
        std::unique_ptr<AsyncTextStreamWorkItem> work(static_cast<AsyncTextStreamWorkItem*>(context));
        if (work && work->moduleKeepAlive)
        {
            TransferModulePinToCallbackReturn(instance, work->moduleKeepAlive);
        }
        ViewerText* viewer = work ? work->viewer : nullptr;
        auto releaseViewer = wil::scope_exit([&]() noexcept
        {
            if (viewer)
            {
                viewer->Release();
            }
        });
        if (! work || ! viewer || ! work->result)
        {
            return;
        }

        const auto startedAt = std::chrono::steady_clock::now();
        auto& result         = work->result;
        auto postTerminal    = wil::scope_exit([&]() noexcept
        {
            if (! result || ! work->hwnd || viewer->_windowIdentity != work->windowIdentity ||
                GetWindowLongPtrW(work->hwnd, GWLP_USERDATA) != reinterpret_cast<LONG_PTR>(viewer))
            {
                return;
            }

            result->elapsedUs =
                static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - startedAt).count());
            const uint64_t terminalRequestId = result->requestId;
            HRESULT terminalHr               = result->hr;
            const bool posted =
                work->injectedFault != AsyncTextStreamFault::PayloadPost &&
                PostMessagePayload(work->hwnd, WndMsg::kViewerTextAsyncStreamComplete, static_cast<WPARAM>(terminalRequestId), std::move(result));
            if (posted)
            {
                return;
            }

            result.reset();
            if (SUCCEEDED(terminalHr))
            {
                terminalHr = E_FAIL;
            }
            DWORD_PTR ignored = 0u;
            if (SendMessageTimeoutW(work->hwnd,
                                    WndMsg::kViewerTextAsyncStreamFailure,
                                    static_cast<WPARAM>(terminalRequestId),
                                    static_cast<LPARAM>(terminalHr),
                                    SMTO_ABORTIFHUNG | SMTO_BLOCK,
                                    2000u,
                                    &ignored) == 0)
            {
                Debug::Error(L"ViewerText: failed to deliver streamed text terminal result (request={}, hr=0x{:08X}, lastError={}).",
                             terminalRequestId,
                             static_cast<unsigned long>(terminalHr),
                             GetLastError());
            }
        });

        wil::com_ptr<IFileSystemIO> fileIo;
        const HRESULT ioHr = work->fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
        if (FAILED(ioHr) || ! fileIo)
        {
            result->hr = FAILED(ioHr) ? ioHr : E_NOINTERFACE;
            return;
        }
        wil::com_ptr<IFileReader> reader;
        const HRESULT openHr = fileIo->CreateFileReader(work->path.c_str(), reader.put());
        if (FAILED(openHr) || ! reader)
        {
            result->hr = FAILED(openHr) ? openHr : HRESULT_FROM_WIN32(ERROR_OPEN_FAILED);
            return;
        }

        const HRESULT seekHr = ViewerTextSafety::SeekExact(reader.get(), result->startOffset);
        if (FAILED(seekHr))
        {
            result->hr = seekHr;
            return;
        }

        const uint64_t available = work->fileSize > result->startOffset ? (work->fileSize - result->startOffset) : 0u;
        const size_t requested   = static_cast<size_t>(std::min<uint64_t>(std::min(available, work->chunkBytes), std::numeric_limits<size_t>::max()));
        std::vector<uint8_t> bytes(requested);
        const HRESULT readHr = ViewerTextSafety::ReadExactly(reader.get(), bytes);
        if (FAILED(readHr))
        {
            result->hr = readHr;
            return;
        }
        const size_t totalRead = bytes.size();

        size_t carryBytes = 0u;
        if (work->encoding == FileEncoding::Utf16LE || work->encoding == FileEncoding::Utf16BE)
        {
            carryBytes = bytes.size() % 2u;
        }
        else if (work->encoding == FileEncoding::Utf32LE || work->encoding == FileEncoding::Utf32BE)
        {
            carryBytes = bytes.size() % 4u;
        }
        else if (work->codePage == CP_UTF8)
        {
            carryBytes = ViewerTextSafety::IncompleteUtf8TailSize(bytes.data(), bytes.size());
        }
        else if (ViewerTextSafety::UsesDbcsBoundaryCarry(work->codePage))
        {
            carryBytes = ViewerTextSafety::IncompleteDbcsTailSize(bytes.data(), bytes.size(), work->codePage);
        }
        carryBytes = std::min(carryBytes, bytes.size());

        const HRESULT decodeHr = DecodeTextWindow(bytes, bytes.size() - carryBytes, work->encoding, work->codePage, result->textBuffer);
        if (FAILED(decodeHr))
        {
            result->hr = decodeHr;
            return;
        }
        BuildTextLineIndexForBuffer(result->textBuffer, result->textLineStarts, result->textLineEnds, result->textMaxLineLength);
        if (work->injectedFault == AsyncTextStreamFault::Worker)
        {
            result->hr = E_FAIL;
            return;
        }
        result->endOffset    = std::min<uint64_t>(work->fileSize, result->startOffset + static_cast<uint64_t>(totalRead) - static_cast<uint64_t>(carryBytes));
        result->streamActive = work->fileSize > work->streamSkipBytes && (work->fileSize - work->streamSkipBytes) > work->chunkBytes;
        result->hr           = S_OK;
    },
                                                                               work.get(),
                                                                               nullptr);
    if (submitted == FALSE)
    {
        Release();
#ifdef _DEBUG
        _debugTextStreamRejectedCount += 1u;
#endif
        OnAsyncTextStreamFailure(requestId, windowIdentity, HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY));
        return false;
    }

    static_cast<void>(work.release());
#ifdef _DEBUG
    _debugTextStreamAcceptedCount += 1u;
#endif
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
    return true;
}

void ViewerText::OnAsyncTextStreamComplete(std::unique_ptr<AsyncTextStreamResult> result) noexcept
{
    if (! result)
    {
        return;
    }
    if (result->windowIdentity != _windowIdentity || result->requestId != _activeAsyncTextStreamRequestId)
    {
#ifdef _DEBUG
        _debugTextStreamStaleCount += 1u;
#endif
        return;
    }
    if (FAILED(result->hr))
    {
        OnAsyncTextStreamFailure(result->requestId, result->windowIdentity, result->hr);
        return;
    }

#ifdef _DEBUG
    const auto uiApplyStartedAt = std::chrono::steady_clock::now();
#endif
    _textStreamLoadPending = false;
    _textBuffer            = std::move(result->textBuffer);
    _textLineStarts        = std::move(result->textLineStarts);
    _textLineEnds          = std::move(result->textLineEnds);
    _textMaxLineLength     = result->textMaxLineLength;
    _textStreamStartOffset = result->startOffset;
    _textStreamEndOffset   = result->endOffset;
    _textStreamActive      = result->streamActive;
    _textTopVisualLine     = 0u;
    _textLeftColumn        = 0u;
    _textCaretIndex        = 0u;
    _textSelAnchor         = 0u;
    _textSelActive         = 0u;
    _textSelecting         = false;
    RebuildTextVisualLines(_hEdit.get());
    if (result->scrollToEnd && TextVisualLineCount() > 0u)
    {
        _textTopVisualLine = static_cast<uint32_t>(std::min<uint64_t>(TextVisualLineCount() - 1u, std::numeric_limits<uint32_t>::max()));
        _textCaretIndex    = NormalizeTextPosition(_textBuffer.size());
        _textSelAnchor     = _textCaretIndex;
        _textSelActive     = _textCaretIndex;
    }
    UpdateTextStreamTotalLineCountAfterLoad();
    UpdateSearchHighlights();
    UpdateTextViewScrollBars(_hEdit.get());
    InvalidateRect(_hEdit.get(), nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
#ifdef _DEBUG
    _debugTextStreamTerminalCount += 1u;
    _debugTextStreamLastTerminalHr = result->hr;
    _debugTextStreamLastElapsedUs  = result->elapsedUs;
    _debugTextStreamLastUiApplyUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - uiApplyStartedAt).count());
#endif
}

void ViewerText::OnAsyncTextStreamFailure(uint64_t requestId, uint64_t windowIdentity, HRESULT hr) noexcept
{
    if (windowIdentity != _windowIdentity || requestId != _activeAsyncTextStreamRequestId)
    {
#ifdef _DEBUG
        _debugTextStreamStaleCount += 1u;
#endif
        return;
    }
    _textStreamLoadPending = false;
    Debug::Error(L"ViewerText: streamed text window request {} failed (hr=0x{:08X}).", requestId, static_cast<unsigned long>(hr));
#ifdef _DEBUG
    _debugTextStreamTerminalCount += 1u;
    _debugTextStreamLastTerminalHr = hr;
#endif
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
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

    const uint64_t navigationStart = _textStreamLoadPending ? _pendingTextStreamStartOffset : _textStreamStartOffset;
    const uint64_t navigationEnd   = _textStreamLoadPending
                                         ? navigationStart + std::min<uint64_t>(chunkBytes, _fileSize > navigationStart ? (_fileSize - navigationStart) : 0u)
                                         : _textStreamEndOffset;
    uint64_t nextOffset            = navigationStart;
    bool scrollToEnd               = false;
    if (backward)
    {
        if (navigationStart <= _textStreamSkipBytes)
        {
            return false;
        }

        const uint64_t delta = std::min<uint64_t>(navigationStart - _textStreamSkipBytes, chunkBytes);
        nextOffset           = navigationStart - delta;
        scrollToEnd          = true;
    }
    else
    {
        if (navigationEnd <= navigationStart || navigationEnd >= _fileSize)
        {
            return false;
        }

        nextOffset  = navigationEnd;
        scrollToEnd = false;
    }

    nextOffset = AlignTextStreamOffset(nextOffset);
    if (nextOffset == navigationStart)
    {
        return false;
    }

    return StartAsyncTextStreamLoad(hwnd, nextOffset, scrollToEnd);
}
