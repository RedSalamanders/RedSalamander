#include "ViewerText.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
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
constexpr uint64_t kMaxHexLoadBytes = 128u * 1024u * 1024u; // 128 MiB
constexpr float kMonoFontSizeDip    = 10.0f * 96.0f / 72.0f;

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

struct HexViewLayout
{
    float marginDip       = 6.0f;
    float headerPadYDip   = 2.0f;
    float headerY         = 0.0f;
    float headerH         = 0.0f;
    float dataStartY      = 0.0f;
    float xOffset         = 0.0f;
    float offsetTextRight = 0.0f;
    float xHex            = 0.0f;
    float hexTextRight    = 0.0f;
    float xText           = 0.0f;
};

HexViewLayout ComputeHexViewLayout(float lineH, float charW, uint64_t fileSize, size_t hexGroupSize) noexcept
{
    HexViewLayout layout{};

    const float padY     = std::clamp(std::floor(lineH * 0.15f), 2.0f, 6.0f);
    layout.headerPadYDip = padY;
    layout.headerY       = layout.marginDip;
    layout.headerH       = lineH + padY * 2.0f;
    layout.dataStartY    = layout.headerY + layout.headerH;
    layout.xOffset       = layout.marginDip;

    const size_t hexDigits    = (fileSize > 0xFFFFFFFFu) ? 16u : 8u;
    const uint64_t maxOffset  = (fileSize > 0) ? (fileSize - 1u) : 0u;
    const size_t decDigits    = DecimalDigits(maxOffset);
    const size_t offsetDigits = std::max<size_t>({12u, hexDigits, decDigits});

    constexpr float gapChars = 4.0f;
    const float gapDip       = gapChars * charW;

    layout.offsetTextRight = layout.xOffset + static_cast<float>(offsetDigits) * charW;
    layout.xHex            = layout.offsetTextRight + gapDip;

    const size_t groupSize  = std::max<size_t>(1u, hexGroupSize);
    const size_t groupCount = (ViewerText::kHexBytesPerLine + groupSize - 1u) / groupSize;
    const size_t hexChars   = groupCount * (groupSize * 2u + 1u);
    layout.hexTextRight     = layout.xHex + static_cast<float>(hexChars) * charW;
    layout.xText            = layout.hexTextRight + gapDip;

    return layout;
}

D2D1_COLOR_F ColorFFromColorRef(COLORREF color, float alpha = 1.0f) noexcept
{
    const float r = static_cast<float>(GetRValue(color)) / 255.0f;
    const float g = static_cast<float>(GetGValue(color)) / 255.0f;
    const float b = static_cast<float>(GetBValue(color)) / 255.0f;
    return D2D1::ColorF(r, g, b, alpha);
}

std::wstring CsvEscape(std::wstring_view value)
{
    if (value.find_first_of(L"\",\r\n") == std::wstring_view::npos)
    {
        return std::wstring(value);
    }

    std::wstring out;
    out.reserve(value.size() + 2);
    out.push_back(L'"');
    for (const wchar_t ch : value)
    {
        if (ch == L'"')
        {
            out.append(L"\"\"");
        }
        else
        {
            out.push_back(ch);
        }
    }
    out.push_back(L'"');
    return out;
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

std::optional<uint64_t> FindHexNeedleForwardInMemory(const std::vector<uint8_t>& hay, uint64_t startOffset, const std::vector<uint8_t>& needle) noexcept
{
    if (needle.empty())
    {
        return std::nullopt;
    }

    if (hay.empty() || startOffset >= static_cast<uint64_t>(hay.size()))
    {
        return std::nullopt;
    }

    const size_t start = static_cast<size_t>(startOffset);
    const auto it      = std::search(hay.begin() + static_cast<std::ptrdiff_t>(start), hay.end(), needle.begin(), needle.end());
    if (it == hay.end())
    {
        return std::nullopt;
    }

    const size_t index = static_cast<size_t>(std::distance(hay.begin(), it));
    return static_cast<uint64_t>(index);
}

std::optional<uint64_t>
FindHexNeedleBackwardInMemory(const std::vector<uint8_t>& hay, uint64_t startOffsetInclusive, const std::vector<uint8_t>& needle) noexcept
{
    if (needle.empty())
    {
        return std::nullopt;
    }

    if (hay.empty())
    {
        return std::nullopt;
    }

    const uint64_t lastIndex  = static_cast<uint64_t>(hay.size() - 1u);
    startOffsetInclusive      = std::min<uint64_t>(startOffsetInclusive, lastIndex);
    const size_t endExclusive = static_cast<size_t>(startOffsetInclusive + 1u);

    const auto it = std::find_end(hay.begin(), hay.begin() + static_cast<std::ptrdiff_t>(endExclusive), needle.begin(), needle.end());
    if (it == hay.begin() + static_cast<std::ptrdiff_t>(endExclusive))
    {
        return std::nullopt;
    }

    const size_t index = static_cast<size_t>(std::distance(hay.begin(), it));
    return static_cast<uint64_t>(index);
}

std::optional<uint64_t>
FindHexNeedleForward(IFileReader* reader, uint64_t fileSize, uint64_t startOffset, const std::vector<uint8_t>& needle, size_t chunkBytes) noexcept
{
    if (! reader || needle.empty())
    {
        return std::nullopt;
    }

    if (fileSize == 0 || startOffset >= fileSize)
    {
        return std::nullopt;
    }

    const size_t needleLen = needle.size();
    if (needleLen > fileSize)
    {
        return std::nullopt;
    }

    if (startOffset > fileSize - static_cast<uint64_t>(needleLen))
    {
        return std::nullopt;
    }

    const size_t overlap = (needleLen > 1u) ? (needleLen - 1u) : 0u;

    chunkBytes = std::max<size_t>(chunkBytes, needleLen);
    chunkBytes = std::clamp<size_t>(chunkBytes, 1u, static_cast<size_t>(std::numeric_limits<unsigned long>::max()));

    std::vector<uint8_t> buffer;
    buffer.resize(chunkBytes + overlap);

    if (startOffset > static_cast<uint64_t>(std::numeric_limits<__int64>::max()))
    {
        return std::nullopt;
    }

    uint64_t newPos      = 0;
    const HRESULT seekHr = reader->Seek(static_cast<__int64>(startOffset), FILE_BEGIN, &newPos);
    if (FAILED(seekHr))
    {
        return std::nullopt;
    }

    uint64_t readOffset = startOffset;
    size_t carry        = 0;

    while (readOffset < fileSize)
    {
        const uint64_t remaining   = fileSize - readOffset;
        const unsigned long toRead = static_cast<unsigned long>(std::min<uint64_t>(remaining, static_cast<uint64_t>(chunkBytes)));

        unsigned long bytesRead = 0;
        const HRESULT readHr    = reader->Read(buffer.data() + carry, toRead, &bytesRead);
        if (FAILED(readHr) || bytesRead == 0)
        {
            return std::nullopt;
        }

        const size_t totalBytes          = carry + static_cast<size_t>(bytesRead);
        const uint64_t bufferStartOffset = readOffset >= static_cast<uint64_t>(carry) ? (readOffset - static_cast<uint64_t>(carry)) : 0u;

        if (totalBytes >= needleLen)
        {
            const size_t maxStart = totalBytes - needleLen;
            for (size_t pos = 0; pos <= maxStart; ++pos)
            {
                const uint64_t matchOffset = bufferStartOffset + static_cast<uint64_t>(pos);
                if (matchOffset < startOffset)
                {
                    continue;
                }

                if (memcmp(buffer.data() + pos, needle.data(), needleLen) == 0)
                {
                    return matchOffset;
                }
            }
        }

        readOffset += static_cast<uint64_t>(bytesRead);

        const size_t newCarry = std::min(overlap, totalBytes);
        if (newCarry > 0)
        {
            memmove(buffer.data(), buffer.data() + (totalBytes - newCarry), newCarry);
        }
        carry = newCarry;
    }

    return std::nullopt;
}

std::optional<uint64_t>
FindHexNeedleBackward(IFileReader* reader, uint64_t fileSize, uint64_t startOffsetInclusive, const std::vector<uint8_t>& needle, size_t chunkBytes) noexcept
{
    if (! reader || needle.empty() || fileSize == 0)
    {
        return std::nullopt;
    }

    const size_t needleLen = needle.size();
    if (needleLen > fileSize)
    {
        return std::nullopt;
    }

    const uint64_t lastIndex = fileSize - 1u;
    startOffsetInclusive     = std::min<uint64_t>(startOffsetInclusive, lastIndex);
    if (startOffsetInclusive > fileSize - static_cast<uint64_t>(needleLen))
    {
        startOffsetInclusive = fileSize - static_cast<uint64_t>(needleLen);
    }

    const size_t overlap = (needleLen > 1u) ? (needleLen - 1u) : 0u;

    chunkBytes = std::max<size_t>(chunkBytes, needleLen);
    chunkBytes = std::clamp<size_t>(chunkBytes, 1u, static_cast<size_t>(std::numeric_limits<unsigned long>::max()));

    std::vector<uint8_t> buffer;
    buffer.resize(chunkBytes + overlap);

    std::vector<uint8_t> carryBytes;
    carryBytes.resize(overlap);
    size_t carry = 0;

    uint64_t blockEnd = startOffsetInclusive + static_cast<uint64_t>(needleLen);
    while (true)
    {
        const uint64_t blockStart    = (blockEnd > static_cast<uint64_t>(chunkBytes)) ? (blockEnd - static_cast<uint64_t>(chunkBytes)) : 0u;
        const uint64_t bytesToRead64 = blockEnd - blockStart;
        if (bytesToRead64 == 0)
        {
            break;
        }

        if (blockStart > static_cast<uint64_t>(std::numeric_limits<__int64>::max()) ||
            bytesToRead64 > static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()))
        {
            return std::nullopt;
        }

        uint64_t newPos      = 0;
        const HRESULT seekHr = reader->Seek(static_cast<__int64>(blockStart), FILE_BEGIN, &newPos);
        if (FAILED(seekHr))
        {
            return std::nullopt;
        }

        const unsigned long toRead = static_cast<unsigned long>(bytesToRead64);
        unsigned long bytesRead    = 0;
        const HRESULT readHr       = reader->Read(buffer.data(), toRead, &bytesRead);
        if (FAILED(readHr) || bytesRead == 0)
        {
            return std::nullopt;
        }

        const size_t readCount = static_cast<size_t>(bytesRead);
        if (carry > 0)
        {
            memcpy(buffer.data() + readCount, carryBytes.data(), carry);
        }

        const size_t totalBytes = readCount + carry;
        if (totalBytes >= needleLen && readCount > 0)
        {
            const size_t maxStartByTotal = totalBytes - needleLen;
            const size_t maxStart        = std::min<size_t>(readCount - 1u, maxStartByTotal);

            for (size_t pos = maxStart + 1u; pos-- > 0u;)
            {
                if (memcmp(buffer.data() + pos, needle.data(), needleLen) == 0)
                {
                    return blockStart + static_cast<uint64_t>(pos);
                }

                if (pos == 0)
                {
                    break;
                }
            }
        }

        if (blockStart == 0)
        {
            break;
        }

        carry = std::min(overlap, readCount);
        if (carry > 0)
        {
            memcpy(carryBytes.data(), buffer.data(), carry);
        }

        blockEnd = blockStart;
    }

    return std::nullopt;
}
} // namespace

// Hex viewer implementation moved from ViewerText.cpp.
LRESULT ViewerText::OnHexViewPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
    _allowEraseBkgndHexView   = false;
    static_cast<void>(hdc);

    if (EnsureHexViewDirect2D(hwnd) && _hexViewTarget && _hexViewBrush)
    {
        const UINT dpi    = GetDpiForWindow(hwnd);
        const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
        const COLORREF fg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

        HRESULT hr = S_OK;
        {
            _hexViewTarget->BeginDraw();
            auto endDraw = wil::scope_exit([&] { hr = _hexViewTarget->EndDraw(); });

            _hexViewTarget->SetTransform(D2D1::Matrix3x2F::Identity());
            _hexViewTarget->Clear(ColorFFromColorRef(bg));

            RECT rc{};
            GetClientRect(hwnd, &rc);

            const float widthDip  = DipsFromPixels(static_cast<int>(rc.right - rc.left), dpi);
            const float heightDip = DipsFromPixels(static_cast<int>(rc.bottom - rc.top), dpi);
            const float charW     = (_hexCharWidthDip > 0.0f) ? _hexCharWidthDip : 8.0f;
            const float lineH     = (_hexLineHeightDip > 0.0f) ? _hexLineHeightDip : 14.0f;

            const HexViewLayout layout  = ComputeHexViewLayout(lineH, charW, _fileSize, HexGroupSize());
            const float marginDip       = layout.marginDip;
            const float xOffset         = layout.xOffset;
            const float xHex            = layout.xHex;
            const float xText           = layout.xText;
            const float headerY         = layout.headerY;
            const float headerH         = layout.headerH;
            const float dataStartY      = layout.dataStartY;
            const float offsetTextRight = layout.offsetTextRight;
            const float hexTextRight    = layout.hexTextRight;

            const std::wstring seed = _currentPath.empty() ? std::wstring(L"viewer") : _currentPath.filename().wstring();
            const COLORREF accent   = _hasTheme ? ResolveAccentColor(_theme, seed) : RGB(0, 120, 215);
            COLORREF offsetAccent   = accent;
            COLORREF dataAccent     = accent;
            COLORREF textAccent     = accent;

            if (_hasTheme && _theme.rainbowMode)
            {
                const uint32_t h = StableHash32(seed);
                const float hue  = static_cast<float>(h % 360u);
                const float sat  = _theme.darkBase ? 0.70f : 0.55f;
                const float val  = _theme.darkBase ? 0.95f : 0.85f;

                offsetAccent = ColorFromHSV(hue, sat, val);
                dataAccent   = ColorFromHSV(hue + 120.0f, sat, val);
                textAccent   = ColorFromHSV(hue + 240.0f, sat, val);
            }

            if (_hexViewFormat && lineH > 0.0f)
            {
                const uint8_t alpha     = (_hasTheme && _theme.darkMode) ? 22u : 16u;
                const COLORREF headerBg = BlendColor(bg, accent, alpha);
                _hexViewBrush->SetColor(ColorFFromColorRef(headerBg));

                const D2D1_RECT_F headerRc = D2D1::RectF(0.0f, headerY, widthDip, headerY + headerH);
                _hexViewTarget->FillRectangle(headerRc, _hexViewBrush.get());

                if (_hexHeaderHot != HexHeaderHit::None)
                {
                    const uint8_t hoverAlpha = (_hasTheme && _theme.darkMode) ? 40u : 28u;
                    COLORREF hoverAccent     = accent;
                    switch (_hexHeaderHot)
                    {
                        case HexHeaderHit::Offset: hoverAccent = offsetAccent; break;
                        case HexHeaderHit::Data: hoverAccent = dataAccent; break;
                        case HexHeaderHit::Text: hoverAccent = textAccent; break;
                        case HexHeaderHit::None:
                        default: break;
                    }

                    const COLORREF hoverBg = BlendColor(bg, hoverAccent, hoverAlpha);
                    _hexViewBrush->SetColor(ColorFFromColorRef(hoverBg));

                    D2D1_RECT_F hotRc = headerRc;
                    switch (_hexHeaderHot)
                    {
                        case HexHeaderHit::Offset:
                            hotRc.left  = xOffset;
                            hotRc.right = std::min(widthDip, xHex);
                            break;
                        case HexHeaderHit::Data:
                            hotRc.left  = xHex;
                            hotRc.right = std::min(widthDip, xText);
                            break;
                        case HexHeaderHit::Text:
                            hotRc.left  = xText;
                            hotRc.right = widthDip;
                            break;
                        case HexHeaderHit::None:
                        default: break;
                    }

                    if (hotRc.right > hotRc.left)
                    {
                        _hexViewTarget->FillRectangle(hotRc, _hexViewBrush.get());
                    }
                }

                UINT offsetHeaderId = IDS_VIEWERTEXT_COL_OFFSET_HEX;
                if (_hexOffsetMode == HexOffsetMode::Decimal)
                {
                    offsetHeaderId = IDS_VIEWERTEXT_COL_OFFSET_DEC;
                }

                UINT dataHeaderId = IDS_VIEWERTEXT_COL_HEX;
                switch (_hexColumnMode)
                {
                    case HexColumnMode::Word: dataHeaderId = IDS_VIEWERTEXT_COL_HEX_WORD; break;
                    case HexColumnMode::Dword: dataHeaderId = IDS_VIEWERTEXT_COL_HEX_DWORD; break;
                    case HexColumnMode::Qword: dataHeaderId = IDS_VIEWERTEXT_COL_HEX_QWORD; break;
                    case HexColumnMode::Byte:
                    default: dataHeaderId = IDS_VIEWERTEXT_COL_HEX; break;
                }

                UINT textHeaderId = IDS_VIEWERTEXT_COL_TEXT_ANSI;
                switch (_hexTextMode)
                {
                    case HexTextMode::Utf8: textHeaderId = IDS_VIEWERTEXT_COL_TEXT_UTF8; break;
                    case HexTextMode::Utf16: textHeaderId = IDS_VIEWERTEXT_COL_TEXT_UTF16; break;
                    case HexTextMode::Ansi:
                    default: textHeaderId = IDS_VIEWERTEXT_COL_TEXT_ANSI; break;
                }

                const std::wstring offsetHeader = LoadStringResource(g_hInstance, offsetHeaderId);
                const std::wstring dataHeader   = LoadStringResource(g_hInstance, dataHeaderId);
                const std::wstring textHeader   = LoadStringResource(g_hInstance, textHeaderId);

                _hexViewBrush->SetColor(ColorFFromColorRef(fg));

                const float padX       = std::max(4.0f, charW * 0.5f);
                const float textTop    = headerY + layout.headerPadYDip;
                const float textBottom = std::max(textTop, headerY + headerH - layout.headerPadYDip);

                const D2D1_RECT_F offsetHeaderRc = D2D1::RectF(xOffset + padX, textTop, std::max(xOffset + padX, xHex - padX), textBottom);
                const D2D1_RECT_F dataHeaderRc   = D2D1::RectF(xHex + padX, textTop, std::max(xHex + padX, xText - padX), textBottom);
                const D2D1_RECT_F textHeaderRc =
                    D2D1::RectF(xText + padX, textTop, std::max(xText + padX, std::max(xText, widthDip - marginDip) - padX), textBottom);

                _hexViewTarget->DrawTextW(offsetHeader.c_str(),
                                          static_cast<UINT32>(std::min<size_t>(offsetHeader.size(), std::numeric_limits<UINT32>::max())),
                                          _hexViewFormat.get(),
                                          offsetHeaderRc,
                                          _hexViewBrush.get(),
                                          D2D1_DRAW_TEXT_OPTIONS_CLIP);
                _hexViewTarget->DrawTextW(dataHeader.c_str(),
                                          static_cast<UINT32>(std::min<size_t>(dataHeader.size(), std::numeric_limits<UINT32>::max())),
                                          _hexViewFormat.get(),
                                          dataHeaderRc,
                                          _hexViewBrush.get(),
                                          D2D1_DRAW_TEXT_OPTIONS_CLIP);
                _hexViewTarget->DrawTextW(textHeader.c_str(),
                                          static_cast<UINT32>(std::min<size_t>(textHeader.size(), std::numeric_limits<UINT32>::max())),
                                          _hexViewFormat.get(),
                                          textHeaderRc,
                                          _hexViewBrush.get(),
                                          D2D1_DRAW_TEXT_OPTIONS_CLIP);

                if (_hasTheme && _theme.rainbowMode)
                {
                    const float barH       = std::max(1.0f, std::min(3.0f, layout.headerPadYDip));
                    const float barTop     = std::max(headerY, (headerY + headerH) - barH);
                    const float barBottom  = headerY + headerH;
                    const uint8_t barAlpha = (_hasTheme && _theme.darkMode) ? 160u : 200u;

                    const float offsetBarLeft  = xOffset;
                    const float offsetBarRight = std::min(widthDip, xHex);
                    if (offsetBarRight > offsetBarLeft)
                    {
                        _hexViewBrush->SetColor(ColorFFromColorRef(BlendColor(bg, offsetAccent, barAlpha)));
                        _hexViewTarget->FillRectangle(D2D1::RectF(offsetBarLeft, barTop, offsetBarRight, barBottom), _hexViewBrush.get());
                    }

                    const float dataBarLeft  = std::min(xHex, widthDip);
                    const float dataBarRight = std::min(widthDip, xText);
                    if (dataBarRight > dataBarLeft)
                    {
                        _hexViewBrush->SetColor(ColorFFromColorRef(BlendColor(bg, dataAccent, barAlpha)));
                        _hexViewTarget->FillRectangle(D2D1::RectF(dataBarLeft, barTop, dataBarRight, barBottom), _hexViewBrush.get());
                    }

                    const float textBarLeft = std::min(xText, widthDip);
                    if (widthDip > textBarLeft)
                    {
                        _hexViewBrush->SetColor(ColorFFromColorRef(BlendColor(bg, textAccent, barAlpha)));
                        _hexViewTarget->FillRectangle(D2D1::RectF(textBarLeft, barTop, widthDip, barBottom), _hexViewBrush.get());
                    }
                }

                const COLORREF divider = BlendColor(bg, fg, (_hasTheme && _theme.darkMode) ? 40u : 20u);
                _hexViewBrush->SetColor(ColorFFromColorRef(divider));
                const D2D1_POINT_2F p0 = D2D1::Point2F(0.0f, headerY + headerH);
                const D2D1_POINT_2F p1 = D2D1::Point2F(widthDip, headerY + headerH);
                _hexViewTarget->DrawLine(p0, p1, _hexViewBrush.get(), 1.0f);
            }

            if (_fileSize > 0 && lineH > 0.0f && _hexViewFormat)
            {
                const float usableH    = std::max(0.0f, heightDip - headerH - 2.0f * marginDip);
                const uint32_t maxRows = std::max<uint32_t>(1u, static_cast<uint32_t>(std::ceil(usableH / lineH)) + 1u);

                std::wstring offsetText;
                std::wstring hexText;
                std::wstring asciiText;
                std::array<ByteSpan, kHexBytesPerLine> hexSpans{};
                std::array<ByteSpan, kHexBytesPerLine> textSpans{};
                size_t validBytes = 0;

                const bool hasSelection        = _hexSelectedOffset.has_value();
                uint64_t selectionStart        = 0;
                uint64_t selectionEndExclusive = 0;
                uint64_t activeOffset          = 0;
                if (hasSelection)
                {
                    activeOffset                         = _hexSelectedOffset.value();
                    const uint64_t anchorOffset          = _hexSelectionAnchorOffset.has_value() ? _hexSelectionAnchorOffset.value() : activeOffset;
                    selectionStart                       = std::min(anchorOffset, activeOffset);
                    const uint64_t selectionEndInclusive = std::max(anchorOffset, activeOffset);
                    selectionEndExclusive = selectionEndInclusive < std::numeric_limits<uint64_t>::max() ? (selectionEndInclusive + 1u) : selectionEndInclusive;
                }

                const bool hasSearch         = _hexSearchNeedleValid && ! _hexSearchNeedle.empty();
                const size_t searchNeedleLen = _hexSearchNeedle.size();
                std::vector<uint8_t> searchMask;
                std::vector<uint8_t> searchBytes;
                const uint8_t* searchBytesPtr  = nullptr;
                uint64_t searchMaskStartOffset = 0;
                COLORREF searchBg              = RGB(0, 0, 0);

                if (hasSearch && searchNeedleLen > 0)
                {
                    searchMaskStartOffset = _hexTopLine * static_cast<uint64_t>(kHexBytesPerLine);
                    if (searchMaskStartOffset < _fileSize)
                    {
                        const uint64_t maxVisibleBytes64 = static_cast<uint64_t>(maxRows) * static_cast<uint64_t>(kHexBytesPerLine);
                        const uint64_t remainingBytes64  = _fileSize - searchMaskStartOffset;
                        const uint64_t visibleBytes64    = std::min<uint64_t>(maxVisibleBytes64, remainingBytes64);

                        if (visibleBytes64 <= static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
                        {
                            size_t visibleBytes = static_cast<size_t>(visibleBytes64);
                            if (visibleBytes >= searchNeedleLen)
                            {
                                const COLORREF searchAccent =
                                    (_hasTheme && ! _theme.highContrast) ? ResolveAccentColor(_theme, L"search") : GetSysColor(COLOR_HIGHLIGHT);
                                const uint8_t alpha = (_hasTheme && _theme.darkMode) ? 60u : 40u;
                                searchBg            = BlendColor(bg, searchAccent, alpha);

                                searchMask.assign(visibleBytes, 0u);

                                if (! _hexBytes.empty() && searchMaskStartOffset < static_cast<uint64_t>(_hexBytes.size()))
                                {
                                    searchBytesPtr = _hexBytes.data() + static_cast<size_t>(searchMaskStartOffset);
                                }
                                else
                                {
                                    searchBytes.resize(visibleBytes);
                                    size_t readTotal = 0;
                                    while (readTotal < visibleBytes)
                                    {
                                        const size_t read = ReadHexBytes(
                                            searchMaskStartOffset + static_cast<uint64_t>(readTotal), searchBytes.data() + readTotal, visibleBytes - readTotal);
                                        if (read == 0)
                                        {
                                            break;
                                        }
                                        readTotal += read;
                                    }

                                    if (readTotal < visibleBytes)
                                    {
                                        searchBytes.resize(readTotal);
                                        searchMask.resize(readTotal);
                                    }

                                    if (! searchBytes.empty())
                                    {
                                        searchBytesPtr = searchBytes.data();
                                    }
                                }

                                if (searchBytesPtr && searchMask.size() >= searchNeedleLen)
                                {
                                    const size_t scanBytes = searchMask.size();
                                    for (size_t i = 0; i + searchNeedleLen <= scanBytes; ++i)
                                    {
                                        if (memcmp(searchBytesPtr + i, _hexSearchNeedle.data(), searchNeedleLen) == 0)
                                        {
                                            for (size_t j = 0; j < searchNeedleLen; ++j)
                                            {
                                                searchMask[i + j] = 1u;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                const bool focusSearchSelection = hasSelection && hasSearch && searchNeedleLen > 0 && selectionEndExclusive > selectionStart &&
                                                  (selectionEndExclusive - selectionStart) == static_cast<uint64_t>(searchNeedleLen);

                for (uint32_t row = 0; row < maxRows; ++row)
                {
                    const uint64_t line       = _hexTopLine + static_cast<uint64_t>(row);
                    const uint64_t lineOffset = line * static_cast<uint64_t>(kHexBytesPerLine);
                    if (lineOffset >= _fileSize)
                    {
                        break;
                    }

                    FormatHexLine(lineOffset, offsetText, hexText, asciiText, hexSpans, textSpans, validBytes);

                    const float y           = dataStartY + static_cast<float>(row) * lineH;
                    const D2D1_RECT_F rowRc = D2D1::RectF(0.0f, y, widthDip, y + lineH);

                    bool highlightRow            = false;
                    uint64_t overlapStart        = 0;
                    uint64_t overlapEndExclusive = 0;
                    if (hasSelection && validBytes > 0)
                    {
                        const uint64_t rowStart        = lineOffset;
                        const uint64_t rowEndExclusive = rowStart + static_cast<uint64_t>(validBytes);
                        if (selectionEndExclusive > rowStart && selectionStart < rowEndExclusive)
                        {
                            highlightRow        = true;
                            overlapStart        = std::max(selectionStart, rowStart);
                            overlapEndExclusive = std::min(selectionEndExclusive, rowEndExclusive);
                        }
                    }

                    if (highlightRow)
                    {
                        const uint8_t alpha  = 40u;
                        const COLORREF rowBg = BlendColor(bg, accent, alpha);
                        _hexViewBrush->SetColor(ColorFFromColorRef(rowBg));
                        _hexViewTarget->FillRectangle(rowRc, _hexViewBrush.get());
                    }

                    const D2D1_RECT_F offsetRc = D2D1::RectF(xOffset, y, std::max(xOffset, offsetTextRight), y + lineH);
                    const D2D1_RECT_F hexRc    = D2D1::RectF(xHex, y, std::max(xHex, hexTextRight), y + lineH);
                    const D2D1_RECT_F textRc   = D2D1::RectF(xText, y, std::max(xText, widthDip - marginDip), y + lineH);

                    if (! searchMask.empty() && validBytes > 0 && lineOffset >= searchMaskStartOffset)
                    {
                        const uint64_t base64 = lineOffset - searchMaskStartOffset;
                        if (base64 < static_cast<uint64_t>(searchMask.size()))
                        {
                            const size_t maskBase = static_cast<size_t>(base64);
                            _hexViewBrush->SetColor(ColorFFromColorRef(searchBg));
                            for (size_t byteIndex = 0; byteIndex < validBytes; ++byteIndex)
                            {
                                const size_t maskIndex = maskBase + byteIndex;
                                if (maskIndex >= searchMask.size())
                                {
                                    break;
                                }

                                if (searchMask[maskIndex] == 0u)
                                {
                                    continue;
                                }

                                const ByteSpan hexSpan = hexSpans[byteIndex];
                                if (hexSpan.length > 0)
                                {
                                    const float hlX        = xHex + static_cast<float>(hexSpan.start) * charW;
                                    const float hlW        = static_cast<float>(hexSpan.length) * charW;
                                    const D2D1_RECT_F hlRc = D2D1::RectF(hlX, y, hlX + hlW, y + lineH);
                                    _hexViewTarget->FillRectangle(hlRc, _hexViewBrush.get());
                                }

                                const ByteSpan textSpan = textSpans[byteIndex];
                                if (textSpan.length > 0)
                                {
                                    const float hlX        = xText + static_cast<float>(textSpan.start) * charW;
                                    const float hlW        = static_cast<float>(textSpan.length) * charW;
                                    const D2D1_RECT_F hlRc = D2D1::RectF(hlX, y, hlX + hlW, y + lineH);
                                    _hexViewTarget->FillRectangle(hlRc, _hexViewBrush.get());
                                }
                            }
                        }
                    }

                    if (highlightRow)
                    {
                        for (uint64_t selected = overlapStart; selected < overlapEndExclusive; ++selected)
                        {
                            const size_t byteIndex = static_cast<size_t>(selected - lineOffset);
                            if (byteIndex >= validBytes)
                            {
                                continue;
                            }

                            const uint8_t alpha = (selected == activeOffset) ? (focusSearchSelection ? 180u : 120u) : (focusSearchSelection ? 120u : 90u);
                            const COLORREF hlBg = BlendColor(bg, accent, alpha);
                            _hexViewBrush->SetColor(ColorFFromColorRef(hlBg));

                            const ByteSpan hexSpan = hexSpans[byteIndex];
                            if (hexSpan.length > 0)
                            {
                                const float hlX        = xHex + static_cast<float>(hexSpan.start) * charW;
                                const float hlW        = static_cast<float>(hexSpan.length) * charW;
                                const D2D1_RECT_F hlRc = D2D1::RectF(hlX, y, hlX + hlW, y + lineH);
                                _hexViewTarget->FillRectangle(hlRc, _hexViewBrush.get());
                            }

                            const ByteSpan textSpan = textSpans[byteIndex];
                            if (textSpan.length > 0)
                            {
                                const float hlX        = xText + static_cast<float>(textSpan.start) * charW;
                                const float hlW        = static_cast<float>(textSpan.length) * charW;
                                const D2D1_RECT_F hlRc = D2D1::RectF(hlX, y, hlX + hlW, y + lineH);
                                _hexViewTarget->FillRectangle(hlRc, _hexViewBrush.get());
                            }
                        }
                    }

                    _hexViewBrush->SetColor(ColorFFromColorRef(fg));
                    _hexViewTarget->DrawTextW(offsetText.c_str(),
                                              static_cast<UINT32>(std::min<size_t>(offsetText.size(), std::numeric_limits<UINT32>::max())),
                                              (_hexViewFormatRight ? _hexViewFormatRight.get() : _hexViewFormat.get()),
                                              offsetRc,
                                              _hexViewBrush.get(),
                                              D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    _hexViewTarget->DrawTextW(hexText.c_str(),
                                              static_cast<UINT32>(std::min<size_t>(hexText.size(), std::numeric_limits<UINT32>::max())),
                                              _hexViewFormat.get(),
                                              hexRc,
                                              _hexViewBrush.get(),
                                              D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    _hexViewTarget->DrawTextW(asciiText.c_str(),
                                              static_cast<UINT32>(std::min<size_t>(asciiText.size(), std::numeric_limits<UINT32>::max())),
                                              _hexViewFormat.get(),
                                              textRc,
                                              _hexViewBrush.get(),
                                              D2D1_DRAW_TEXT_OPTIONS_CLIP);
                }
            }

            DrawLoadingOverlay(_hexViewTarget.get(), _hexViewBrush.get(), widthDip, heightDip);
        }

        if (hr == D2DERR_RECREATE_TARGET)
        {
            DiscardHexViewDirect2D();
        }

        return 0;
    }

    FillRect(ps.hdc, &ps.rcPaint, _backgroundBrush.get());
    return 0;
}

LRESULT ViewerText::OnHexViewSize(HWND hwnd, UINT32 width, UINT32 height) noexcept
{
    if (_hexViewTarget && width > 0 && height > 0)
    {
        const HRESULT hr = _hexViewTarget->Resize(D2D1::SizeU(width, height));
        if (FAILED(hr))
        {
            DiscardHexViewDirect2D();
        }
    }

    UpdateHexViewScrollBars(hwnd);
    InvalidateRect(hwnd, nullptr, TRUE);
    return 0;
}

LRESULT ViewerText::OnHexViewVScroll(HWND hwnd, UINT scrollCode) noexcept
{
    const uint64_t totalLines = (_fileSize + (kHexBytesPerLine - 1)) / kHexBytesPerLine;
    if (totalLines == 0)
    {
        return 0;
    }

    const uint64_t maxLine = totalLines - 1;

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_ALL;
    GetScrollInfo(hwnd, SB_VERT, &si);

    uint64_t top   = _hexTopLine;
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
                const int clampedPos = std::clamp(pos, 0, std::numeric_limits<int>::max());
                top = maxLine == 0 ? 0 : (static_cast<uint64_t>(clampedPos) * maxLine) / static_cast<uint64_t>(std::numeric_limits<int>::max());
            }
            break;
        }
        default: break;
    }

    top = std::min<uint64_t>(top, maxLine);
    if (top == _hexTopLine)
    {
        return 0;
    }

    _hexTopLine = top;
    UpdateHexViewScrollBars(hwnd);
    InvalidateRect(hwnd, nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
    return 0;
}

LRESULT ViewerText::OnHexViewMouseWheel(HWND hwnd, int delta) noexcept
{
    if (_fileSize == 0)
    {
        return 0;
    }

    if (delta == 0)
    {
        return 0;
    }

    const int scrollLines = 3;
    const int absDelta    = delta >= 0 ? delta : -delta;
    const int notchCount  = std::max(1, absDelta / WHEEL_DELTA);
    const int stepLines   = notchCount * scrollLines;
    const int signedLines = (delta > 0) ? -stepLines : stepLines;

    const uint64_t totalLines = (_fileSize + (kHexBytesPerLine - 1)) / kHexBytesPerLine;
    const uint64_t maxLine    = totalLines > 0 ? (totalLines - 1) : 0;

    int64_t nextTop = static_cast<int64_t>(_hexTopLine) + static_cast<int64_t>(signedLines);
    if (nextTop < 0)
    {
        nextTop = 0;
    }

    uint64_t top = static_cast<uint64_t>(nextTop);
    top          = std::min<uint64_t>(top, maxLine);

    if (top != _hexTopLine)
    {
        _hexTopLine = top;
        UpdateHexViewScrollBars(hwnd);
        InvalidateRect(hwnd, nullptr, TRUE);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
        }
    }

    return 0;
}

LRESULT ViewerText::OnHexViewMouseMove(HWND hwnd, POINT pt) noexcept
{
    OnHexMouseMove(hwnd, static_cast<int>(pt.x), static_cast<int>(pt.y));
    if (! _hexTrackingMouseLeave)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = hwnd;
        if (TrackMouseEvent(&tme) != 0)
        {
            _hexTrackingMouseLeave = true;
        }
    }
    return 0;
}

LRESULT ViewerText::OnHexViewMouseLeave(HWND hwnd) noexcept
{
    _hexTrackingMouseLeave = false;
    if (_hexHeaderHot != HexHeaderHit::None)
    {
        _hexHeaderHot = HexHeaderHit::None;
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return 0;
}

LRESULT ViewerText::OnHexViewSetCursor(HWND /*hwnd*/, LPARAM lParam) noexcept
{
    if (LOWORD(lParam) == HTCLIENT && _hexHeaderHot != HexHeaderHit::None)
    {
        SetCursor(LoadCursorW(nullptr, IDC_HAND));
        return TRUE;
    }
    return 0;
}

LRESULT ViewerText::OnHexViewLButtonDown(HWND hwnd, POINT pt) noexcept
{
    OnHexMouseDown(hwnd, static_cast<int>(pt.x), static_cast<int>(pt.y));
    return 0;
}

LRESULT ViewerText::OnHexViewKeyDown(HWND hwnd, WPARAM vk, LPARAM lParam) noexcept
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
        CopyHexCsvToClipboard(hwnd);
        return 0;
    }

    if (_fileSize == 0)
    {
        return DefWindowProcW(hwnd, WM_KEYDOWN, vk, lParam);
    }

    if (vk == VK_UP || vk == VK_DOWN || vk == VK_LEFT || vk == VK_RIGHT || vk == VK_PRIOR || vk == VK_NEXT || vk == VK_HOME || vk == VK_END)
    {
        uint64_t offset = 0;
        if (_hexSelectedOffset.has_value())
        {
            offset = _hexSelectedOffset.value();
        }
        else
        {
            offset = _hexTopLine * static_cast<uint64_t>(kHexBytesPerLine);
            if (offset >= _fileSize)
            {
                offset = _fileSize - 1;
            }
        }

        uint64_t nextOffset = offset;
        if (vk == VK_HOME)
        {
            nextOffset = 0;
        }
        else if (vk == VK_END)
        {
            nextOffset = _fileSize - 1;
        }
        else
        {
            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_PAGE;
            static_cast<void>(GetScrollInfo(hwnd, SB_VERT, &si));
            const uint64_t pageLines = std::max<uint64_t>(1u, static_cast<uint64_t>(si.nPage == 0 ? 1u : si.nPage));

            int64_t delta = 0;
            if (vk == VK_LEFT)
            {
                delta = -1;
            }
            else if (vk == VK_RIGHT)
            {
                delta = 1;
            }
            else if (vk == VK_UP)
            {
                delta = -static_cast<int64_t>(kHexBytesPerLine);
            }
            else if (vk == VK_DOWN)
            {
                delta = static_cast<int64_t>(kHexBytesPerLine);
            }
            else if (vk == VK_PRIOR)
            {
                delta = -static_cast<int64_t>(kHexBytesPerLine) * static_cast<int64_t>(pageLines);
            }
            else if (vk == VK_NEXT)
            {
                delta = static_cast<int64_t>(kHexBytesPerLine) * static_cast<int64_t>(pageLines);
            }

            int64_t signedNext = static_cast<int64_t>(offset) + delta;
            if (signedNext < 0)
            {
                signedNext = 0;
            }

            nextOffset = static_cast<uint64_t>(signedNext);
            if (nextOffset >= _fileSize)
            {
                nextOffset = _fileSize - 1;
            }
        }

        if (shift)
        {
            if (! _hexSelectionAnchorOffset.has_value())
            {
                _hexSelectionAnchorOffset = offset;
            }
        }
        else
        {
            _hexSelectionAnchorOffset = nextOffset;
        }

        _hexSelectedOffset = nextOffset;

        const uint64_t line = nextOffset / static_cast<uint64_t>(kHexBytesPerLine);
        SCROLLINFO si{};
        si.cbSize = sizeof(si);
        si.fMask  = SIF_PAGE;
        static_cast<void>(GetScrollInfo(hwnd, SB_VERT, &si));
        const uint64_t pageLines = std::max<uint64_t>(1u, static_cast<uint64_t>(si.nPage == 0 ? 1u : si.nPage));

        if (line < _hexTopLine)
        {
            _hexTopLine = line;
        }
        else if (line >= _hexTopLine + pageLines)
        {
            _hexTopLine = line - pageLines + 1;
        }

        UpdateHexViewScrollBars(hwnd);
        InvalidateRect(hwnd, nullptr, TRUE);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
        }
        return 0;
    }

    return DefWindowProcW(hwnd, WM_KEYDOWN, vk, lParam);
}

LRESULT ViewerText::OnHexViewSetFocus(HWND hwnd) noexcept
{
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
    InvalidateRect(hwnd, nullptr, TRUE);
    return 0;
}

LRESULT ViewerText::OnHexViewKillFocus(HWND hwnd) noexcept
{
    InvalidateRect(hwnd, nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
    return 0;
}

LRESULT ViewerText::HexViewProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_ERASEBKGND: return _allowEraseBkgndHexView ? DefWindowProcW(hwnd, msg, wp, lp) : 1;
        case WM_PAINT: return OnHexViewPaint(hwnd);
        case WM_SIZE: return OnHexViewSize(hwnd, static_cast<UINT32>(LOWORD(lp)), static_cast<UINT32>(HIWORD(lp)));
        case WM_VSCROLL: return OnHexViewVScroll(hwnd, static_cast<UINT>(LOWORD(wp)));
        case WM_MOUSEWHEEL: return OnHexViewMouseWheel(hwnd, GET_WHEEL_DELTA_WPARAM(wp));
        case WM_MOUSEMOVE:
            return OnHexViewMouseMove(hwnd, {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))});
        case WM_MOUSELEAVE: return OnHexViewMouseLeave(hwnd);
        case WM_SETCURSOR:
            if (OnHexViewSetCursor(hwnd, lp) != 0)
            {
                return TRUE;
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_LBUTTONDOWN:
            return OnHexViewLButtonDown(hwnd, {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))});
        case WM_LBUTTONUP: OnHexMouseUp(hwnd); return 0;
        case WM_CAPTURECHANGED: _hexSelecting = false; return 0;
        case WM_KEYDOWN: return OnHexViewKeyDown(hwnd, wp, lp);
        case WM_SETFOCUS: return OnHexViewSetFocus(hwnd);
        case WM_KILLFOCUS: return OnHexViewKillFocus(hwnd);
        default: return DefWindowProcW(hwnd, msg, wp, lp);
    }
}

void ViewerText::ResetHexState() noexcept
{
    _hexBytes.clear();
    _hexSelectionAnchorOffset.reset();
    _hexSelectedOffset.reset();
    _hexCache.clear();
    _hexTopLine             = 0;
    _hexSelecting           = false;
    _hexHeaderHot           = HexHeaderHit::None;
    _hexTrackingMouseLeave  = false;
    _hexCacheOffset         = 0;
    _hexCacheValid          = 0;
    _hexLineCacheItem       = -1;
    _hexLineCacheValidBytes = 0;
    _hexLineCacheOffsetText.clear();
    _hexLineCacheHexText.clear();
    _hexLineCacheAsciiText.clear();
    for (auto& span : _hexLineCacheHexSpans)
    {
        span = {};
    }
    for (auto& span : _hexLineCacheTextSpans)
    {
        span = {};
    }
}

void ViewerText::UpdateHexViewScrollBars(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const uint64_t totalLines = (_fileSize + (kHexBytesPerLine - 1)) / kHexBytesPerLine;
    const uint64_t maxLine    = totalLines > 0 ? (totalLines - 1) : 0;

    RECT client{};
    GetClientRect(hwnd, &client);
    const UINT dpi             = GetDpiForWindow(hwnd);
    const float heightDip      = std::max(1.0f, DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi));
    const float lineH          = (_hexLineHeightDip > 0.0f) ? _hexLineHeightDip : 14.0f;
    const float charW          = (_hexCharWidthDip > 0.0f) ? _hexCharWidthDip : 8.0f;
    const HexViewLayout layout = ComputeHexViewLayout(lineH, charW, _fileSize, HexGroupSize());
    const float marginDip      = layout.marginDip;
    const float headerH        = layout.headerH;
    const float usableDip      = std::max(0.0f, heightDip - headerH - 2.0f * marginDip);
    const uint32_t pageLines   = std::max<uint32_t>(1u, static_cast<uint32_t>(std::floor(usableDip / lineH)));

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS | SIF_DISABLENOSCROLL;
    si.nMin   = 0;

    if (maxLine <= static_cast<uint64_t>(std::numeric_limits<int>::max()))
    {
        si.nMax  = static_cast<int>(maxLine);
        si.nPos  = static_cast<int>(std::min<uint64_t>(_hexTopLine, maxLine));
        si.nPage = static_cast<UINT>(pageLines);
    }
    else
    {
        constexpr int maxPos = std::numeric_limits<int>::max();
        const uint64_t top   = std::min<uint64_t>(_hexTopLine, maxLine);
        const uint64_t pos64 = maxLine == 0 ? 0 : (top * static_cast<uint64_t>(maxPos)) / maxLine;
        si.nMax              = maxPos;
        si.nPos              = static_cast<int>(pos64);
        si.nPage             = static_cast<UINT>(pageLines);
    }

    SetScrollInfo(hwnd, SB_VERT, &si, TRUE);
    ShowScrollBar(hwnd, SB_HORZ, FALSE);
}

bool ViewerText::EnsureHexViewDirect2D(HWND hwnd) noexcept
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

    if (! _hexViewTarget)
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

        const HRESULT hr = _d2dFactory->CreateHwndRenderTarget(props, hwndProps, _hexViewTarget.put());
        if (FAILED(hr) || ! _hexViewTarget)
        {
            _hexViewTarget.reset();
            return false;
        }

        _hexViewTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    }
    else
    {
        _hexViewTarget->SetDpi(dpiF, dpiF);
    }

    if (! _hexViewBrush)
    {
        const HRESULT hr = _hexViewTarget->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::Black), _hexViewBrush.put());
        if (FAILED(hr) || ! _hexViewBrush)
        {
            _hexViewBrush.reset();
            return false;
        }
    }

    if (! _hexViewFormat)
    {
        const HRESULT hr = _dwriteFactory->CreateTextFormat(
            L"Consolas", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, kMonoFontSizeDip, L"", _hexViewFormat.put());
        if (FAILED(hr) || ! _hexViewFormat)
        {
            _hexViewFormat.reset();
            return false;
        }

        static_cast<void>(_hexViewFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));
        static_cast<void>(_hexViewFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
        static_cast<void>(_hexViewFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
    }

    if (! _hexViewFormatRight)
    {
        const HRESULT hr = _dwriteFactory->CreateTextFormat(L"Consolas",
                                                            nullptr,
                                                            DWRITE_FONT_WEIGHT_NORMAL,
                                                            DWRITE_FONT_STYLE_NORMAL,
                                                            DWRITE_FONT_STRETCH_NORMAL,
                                                            kMonoFontSizeDip,
                                                            L"",
                                                            _hexViewFormatRight.put());
        if (FAILED(hr) || ! _hexViewFormatRight)
        {
            _hexViewFormatRight.reset();
            return false;
        }

        static_cast<void>(_hexViewFormatRight->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING));
        static_cast<void>(_hexViewFormatRight->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
        static_cast<void>(_hexViewFormatRight->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
    }

    if (_hexCharWidthDip <= 0.0f || _hexLineHeightDip <= 0.0f)
    {
        wil::com_ptr<IDWriteTextLayout> layout;
        const HRESULT hr = _dwriteFactory->CreateTextLayout(L"0", 1, _hexViewFormat.get(), 1024.0f, 1024.0f, layout.put());
        if (SUCCEEDED(hr) && layout)
        {
            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(layout->GetMetrics(&metrics)))
            {
                _hexCharWidthDip  = std::max(1.0f, metrics.widthIncludingTrailingWhitespace);
                _hexLineHeightDip = std::max(1.0f, metrics.height);
            }
        }
    }

    return true;
}

void ViewerText::DiscardHexViewDirect2D() noexcept
{
    _hexViewBrush.reset();
    _hexViewFormat.reset();
    _hexViewFormatRight.reset();
    _hexViewTarget.reset();
    _hexCharWidthDip  = 0.0f;
    _hexLineHeightDip = 0.0f;
}

void ViewerText::CommandFindNextHex(HWND hwnd, bool backward)
{
    if (_searchQuery.empty())
    {
        CommandFind(hwnd);
        return;
    }

    if (_viewMode != ViewMode::Hex)
    {
        SetViewMode(hwnd, ViewMode::Hex);
    }

    if (! _hHex)
    {
        return;
    }

    if (_fileSize == 0)
    {
        MessageBeep(MB_ICONINFORMATION);
        return;
    }

    if (! _hexSearchNeedleValid || _hexSearchNeedle.empty())
    {
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_SEARCH_HEX_INVALID);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
        }

        ShowInlineAlert(InlineAlertSeverity::Info, IDS_VIEWERTEXT_NAME, IDS_VIEWERTEXT_MSG_SEARCH_HEX_INVALID);
        return;
    }

    const size_t needleLen = _hexSearchNeedle.size();
    if (needleLen == 0 || needleLen > _fileSize)
    {
        MessageBeep(MB_ICONINFORMATION);
        return;
    }

    uint64_t selectionStart        = 0;
    uint64_t selectionEndExclusive = 0;
    bool hasSelection              = false;

    if (_hexSelectedOffset.has_value())
    {
        hasSelection                = true;
        const uint64_t active       = _hexSelectedOffset.value();
        const uint64_t anchor       = _hexSelectionAnchorOffset.has_value() ? _hexSelectionAnchorOffset.value() : active;
        selectionStart              = std::min(anchor, active);
        const uint64_t endInclusive = std::max(anchor, active);
        selectionEndExclusive       = endInclusive < std::numeric_limits<uint64_t>::max() ? (endInclusive + 1u) : endInclusive;
    }

    const uint64_t viewStartOffset = _hexTopLine * static_cast<uint64_t>(kHexBytesPerLine);
    if (! hasSelection)
    {
        selectionStart        = std::min<uint64_t>(viewStartOffset, _fileSize > 0 ? (_fileSize - 1u) : 0u);
        selectionEndExclusive = selectionStart;
    }

    const size_t chunkBytes = std::clamp<size_t>(static_cast<size_t>(_config.hexBufferMiB) * 1024u * 1024u, 256u * 1024u, 16u * 1024u * 1024u);

    bool wrapped = false;
    std::optional<uint64_t> matchStart;

    auto findForward = [&](uint64_t start) noexcept -> std::optional<uint64_t>
    {
        if (! _hexBytes.empty())
        {
            return FindHexNeedleForwardInMemory(_hexBytes, start, _hexSearchNeedle);
        }
        return FindHexNeedleForward(_fileReader.get(), _fileSize, start, _hexSearchNeedle, chunkBytes);
    };

    auto findBackward = [&](uint64_t startInclusive) noexcept -> std::optional<uint64_t>
    {
        if (! _hexBytes.empty())
        {
            return FindHexNeedleBackwardInMemory(_hexBytes, startInclusive, _hexSearchNeedle);
        }
        return FindHexNeedleBackward(_fileReader.get(), _fileSize, startInclusive, _hexSearchNeedle, chunkBytes);
    };

    if (backward)
    {
        if (selectionStart == 0)
        {
            matchStart = findBackward(0);
        }
        else
        {
            matchStart = findBackward(selectionStart - 1u);
        }

        if (! matchStart.has_value())
        {
            wrapped    = true;
            matchStart = findBackward(_fileSize - 1u);
        }
    }
    else
    {
        matchStart = findForward(selectionEndExclusive);
        if (! matchStart.has_value())
        {
            wrapped    = true;
            matchStart = findForward(0);
        }
    }

    if (! matchStart.has_value())
    {
        MessageBeep(MB_ICONINFORMATION);
        return;
    }

    const uint64_t matchOffset = matchStart.value();
    uint64_t matchEndInclusive = matchOffset + static_cast<uint64_t>(needleLen - 1u);
    if (matchEndInclusive >= _fileSize)
    {
        matchEndInclusive = _fileSize - 1u;
    }

    _hexSelectionAnchorOffset = matchEndInclusive;
    _hexSelectedOffset        = matchOffset;

    const uint64_t targetLine = matchOffset / static_cast<uint64_t>(kHexBytesPerLine);
    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_PAGE;
    static_cast<void>(GetScrollInfo(_hHex.get(), SB_VERT, &si));
    const uint64_t pageLines = std::max<uint64_t>(1u, static_cast<uint64_t>(si.nPage == 0 ? 1u : si.nPage));

    if (targetLine < _hexTopLine)
    {
        _hexTopLine = targetLine;
    }
    else if (targetLine >= _hexTopLine + pageLines)
    {
        _hexTopLine = targetLine - pageLines + 1u;
    }

    UpdateHexViewScrollBars(_hHex.get());

    if (wrapped)
    {
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_SEARCH_WRAPPED);
    }
    else
    {
        _statusMessage.clear();
    }

    InvalidateRect(_hHex.get(), nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }

    if (wrapped)
    {
        ShowInlineAlert(InlineAlertSeverity::Info, IDS_VIEWERTEXT_NAME, IDS_VIEWERTEXT_MSG_SEARCH_WRAPPED);
    }
}

void ViewerText::CommandGoToOffsetValue(HWND hwnd, uint64_t offset)
{
    if (_fileSize > 0 && offset >= _fileSize)
    {
        offset = _fileSize - 1;
    }

    SetViewMode(hwnd, ViewMode::Hex);
    if (! _hHex)
    {
        return;
    }

    if (_fileSize == 0)
    {
        return;
    }

    _hexSelectionAnchorOffset = offset;
    _hexSelectedOffset        = offset;

    const uint64_t targetLine = offset / static_cast<uint64_t>(kHexBytesPerLine);

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_PAGE;
    static_cast<void>(GetScrollInfo(_hHex.get(), SB_VERT, &si));
    const uint64_t pageLines = std::max<uint64_t>(1u, static_cast<uint64_t>(si.nPage == 0 ? 1u : si.nPage));

    if (targetLine < _hexTopLine)
    {
        _hexTopLine = targetLine;
    }
    else if (targetLine >= _hexTopLine + pageLines)
    {
        _hexTopLine = targetLine - pageLines + 1;
    }

    UpdateHexViewScrollBars(_hHex.get());
    InvalidateRect(_hHex.get(), nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
}

std::wstring ViewerText::FormatFileOffset(uint64_t offset) const
{
    if (_fileSize > 0xFFFFFFFFu)
    {
        return FormatStringResource(g_hInstance, IDS_VIEWERTEXT_OFFSET_STATUS_FORMAT_64, offset, offset);
    }

    return FormatStringResource(g_hInstance, IDS_VIEWERTEXT_OFFSET_STATUS_FORMAT_32, static_cast<uint32_t>(offset), static_cast<uint64_t>(offset));
}

HRESULT ViewerText::LoadHexData(HWND hwnd) noexcept
{
    if (_fileSize == 0)
    {
        ResetHexState();
        UpdateHexItemCount(hwnd);
        return S_OK;
    }

    if (! _fileReader)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
    }

    if (! _hexBytes.empty())
    {
        UpdateHexItemCount(hwnd);
        return S_OK;
    }

    if (_fileSize <= kMaxHexLoadBytes && _fileSize <= static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
    {
        _hexBytes.resize(static_cast<size_t>(_fileSize));

        uint64_t ignored     = 0;
        const HRESULT seekHr = _fileReader->Seek(0, FILE_BEGIN, &ignored);
        if (FAILED(seekHr))
        {
            _hexBytes.clear();
            return seekHr;
        }

        size_t offset = 0;
        while (offset < _hexBytes.size())
        {
            const unsigned long want = static_cast<unsigned long>(std::min<size_t>(256 * 1024, _hexBytes.size() - offset));
            unsigned long read       = 0;
            const HRESULT readHr     = _fileReader->Read(_hexBytes.data() + offset, want, &read);
            if (FAILED(readHr))
            {
                _hexBytes.clear();
                return readHr;
            }
            if (read == 0)
            {
                break;
            }
            offset += read;
        }
    }
    else
    {
        static_cast<void>(RefillHexCache(0));
    }

    if (_hHex)
    {
        UpdateHexViewScrollBars(_hHex.get());
        InvalidateRect(_hHex.get(), nullptr, TRUE);
    }
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
    static_cast<void>(hwnd);
    return S_OK;
}

void ViewerText::UpdateHexItemCount(HWND /*hwnd*/) noexcept
{
    if (! _hHex)
    {
        return;
    }

    UpdateHexViewScrollBars(_hHex.get());
    InvalidateRect(_hHex.get(), nullptr, TRUE);
}

void ViewerText::UpdateHexColumns(HWND hwnd) noexcept
{
    if (! _hHex)
    {
        return;
    }
    static_cast<void>(hwnd);
    UpdateHexViewScrollBars(_hHex.get());
    InvalidateRect(_hHex.get(), nullptr, TRUE);
}

size_t ViewerText::HexGroupSize() const noexcept
{
    switch (_hexColumnMode)
    {
        case HexColumnMode::Word: return 2u;
        case HexColumnMode::Dword: return 4u;
        case HexColumnMode::Qword: return 8u;
        case HexColumnMode::Byte:
        default: return 1u;
    }
}

void ViewerText::UpdateHexTextColumnHeader() noexcept
{
    if (! _hHex)
    {
        return;
    }
    _hexLineCacheItem       = -1;
    _hexLineCacheValidBytes = 0;
    InvalidateRect(_hHex.get(), nullptr, TRUE);
}

void ViewerText::UpdateHexColumnHeader() noexcept
{
    if (! _hHex)
    {
        return;
    }
    _hexLineCacheItem       = -1;
    _hexLineCacheValidBytes = 0;
    InvalidateRect(_hHex.get(), nullptr, TRUE);
}

void ViewerText::CycleHexColumnMode() noexcept
{
    switch (_hexColumnMode)
    {
        case HexColumnMode::Byte: _hexColumnMode = HexColumnMode::Word; break;
        case HexColumnMode::Word: _hexColumnMode = HexColumnMode::Dword; break;
        case HexColumnMode::Dword: _hexColumnMode = HexColumnMode::Qword; break;
        case HexColumnMode::Qword:
        default: _hexColumnMode = HexColumnMode::Byte; break;
    }

    _hexLineCacheItem       = -1;
    _hexLineCacheValidBytes = 0;
    UpdateHexColumnHeader();
    if (_hHex)
    {
        InvalidateRect(_hHex.get(), nullptr, TRUE);
    }
}

void ViewerText::CycleHexOffsetMode() noexcept
{
    _hexOffsetMode = (_hexOffsetMode == HexOffsetMode::Hex) ? HexOffsetMode::Decimal : HexOffsetMode::Hex;
    UpdateHexColumnHeader();
    if (_hHex)
    {
        InvalidateRect(_hHex.get(), nullptr, TRUE);
    }
}

void ViewerText::CycleHexTextMode() noexcept
{
    switch (_hexTextMode)
    {
        case HexTextMode::Ansi: _hexTextMode = HexTextMode::Utf8; break;
        case HexTextMode::Utf8: _hexTextMode = HexTextMode::Utf16; break;
        case HexTextMode::Utf16:
        default: _hexTextMode = HexTextMode::Ansi; break;
    }

    UpdateHexTextColumnHeader();
    if (_hHex)
    {
        InvalidateRect(_hHex.get(), nullptr, TRUE);
    }
}

bool ViewerText::HexBigEndian() const noexcept
{
    FileEncoding encoding = DisplayEncodingFileEncoding();
    if (encoding == FileEncoding::Unknown)
    {
        encoding = _encoding;
    }

    return encoding == FileEncoding::Utf16BE || encoding == FileEncoding::Utf32BE;
}

void ViewerText::EnsureHexLineCache(int item) noexcept
{
    if (item < 0)
    {
        _hexLineCacheItem       = -1;
        _hexLineCacheValidBytes = 0;
        _hexLineCacheOffsetText.clear();
        _hexLineCacheHexText.clear();
        _hexLineCacheAsciiText.clear();
        for (auto& span : _hexLineCacheHexSpans)
        {
            span = {};
        }
        for (auto& span : _hexLineCacheTextSpans)
        {
            span = {};
        }
        return;
    }

    if (item == _hexLineCacheItem)
    {
        return;
    }

    _hexLineCacheItem     = item;
    const uint64_t offset = static_cast<uint64_t>(item) * static_cast<uint64_t>(kHexBytesPerLine);
    FormatHexLine(
        offset, _hexLineCacheOffsetText, _hexLineCacheHexText, _hexLineCacheAsciiText, _hexLineCacheHexSpans, _hexLineCacheTextSpans, _hexLineCacheValidBytes);
}

size_t ViewerText::ReadHexBytes(uint64_t offset, uint8_t* dest, size_t destSize) noexcept
{
    if (! dest || destSize == 0)
    {
        return 0;
    }

    if (! _hexBytes.empty())
    {
        const uint64_t total = static_cast<uint64_t>(_hexBytes.size());
        if (offset >= total)
        {
            return 0;
        }

        const size_t start     = static_cast<size_t>(offset);
        const size_t available = _hexBytes.size() - start;
        const size_t take      = std::min(destSize, available);
        memcpy(dest, _hexBytes.data() + start, take);
        return take;
    }

    if (! _fileReader || _fileSize == 0)
    {
        return 0;
    }

    for (int attempt = 0; attempt < 2; ++attempt)
    {
        if (_hexCacheValid > 0 && offset >= _hexCacheOffset && offset < _hexCacheOffset + static_cast<uint64_t>(_hexCacheValid))
        {
            const uint64_t start64 = offset - _hexCacheOffset;
            const size_t start     = static_cast<size_t>(start64);
            if (start >= _hexCacheValid)
            {
                return 0;
            }

            const size_t available = _hexCacheValid - start;
            const size_t take      = std::min(destSize, available);
            memcpy(dest, _hexCache.data() + start, take);
            return take;
        }

        if (FAILED(RefillHexCache(offset)))
        {
            return 0;
        }
    }

    return 0;
}

HRESULT ViewerText::RefillHexCache(uint64_t offset) noexcept
{
    if (! _fileReader || _fileSize == 0)
    {
        return E_FAIL;
    }

    constexpr uint64_t kAlign = 4096u;

    uint64_t cacheBytes = static_cast<uint64_t>(_config.hexBufferMiB) * 1024u * 1024u;
    cacheBytes          = std::clamp<uint64_t>(cacheBytes, 256u * 1024u, 256u * 1024u * 1024u);

    const uint64_t aligned   = (offset / kAlign) * kAlign;
    const uint64_t remaining = (_fileSize > aligned) ? (_fileSize - aligned) : 0;
    const uint64_t want64    = std::min<uint64_t>(remaining, cacheBytes);
    const unsigned long want = want64 > static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()) ? std::numeric_limits<unsigned long>::max()
                                                                                                         : static_cast<unsigned long>(want64);

    _hexCacheOffset = aligned;
    _hexCacheValid  = 0;
    if (want == 0)
    {
        _hexCache.clear();
        return S_FALSE;
    }

    if (_hexCache.size() < static_cast<size_t>(want))
    {
        _hexCache.resize(static_cast<size_t>(want));
    }

    if (aligned > static_cast<uint64_t>(std::numeric_limits<__int64>::max()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    uint64_t ignored     = 0;
    const HRESULT seekHr = _fileReader->Seek(static_cast<__int64>(aligned), FILE_BEGIN, &ignored);
    if (FAILED(seekHr))
    {
        return seekHr;
    }

    unsigned long read   = 0;
    const HRESULT readHr = _fileReader->Read(_hexCache.data(), want, &read);
    if (FAILED(readHr))
    {
        return readHr;
    }

    _hexCacheValid = static_cast<size_t>(read);
    return (_hexCacheValid == 0) ? S_FALSE : S_OK;
}

void ViewerText::FormatHexLine(uint64_t offset,
                               std::wstring& outOffset,
                               std::wstring& outHex,
                               std::wstring& outAscii,
                               std::array<ByteSpan, kHexBytesPerLine>& hexSpans,
                               std::array<ByteSpan, kHexBytesPerLine>& textSpans,
                               size_t& validBytes) noexcept
{
    outOffset.clear();
    outHex.clear();
    outAscii.clear();
    validBytes = 0;

    for (auto& span : hexSpans)
    {
        span = {};
    }
    for (auto& span : textSpans)
    {
        span = {};
    }

    std::array<uint8_t, kHexBytesPerLine> bytes{};
    const size_t count = ReadHexBytes(offset, bytes.data(), bytes.size());
    validBytes         = count;

    if (_hexOffsetMode == HexOffsetMode::Decimal)
    {
        outOffset = FormatStringResource(g_hInstance, IDS_VIEWERTEXT_OFFSET_COL_DEC_FORMAT, offset);
    }
    else
    {
        if (_fileSize > 0xFFFFFFFFu)
        {
            outOffset = FormatStringResource(g_hInstance, IDS_VIEWERTEXT_OFFSET_COL_FORMAT_64, offset);
        }
        else
        {
            outOffset = FormatStringResource(g_hInstance, IDS_VIEWERTEXT_OFFSET_COL_FORMAT_32, static_cast<uint32_t>(offset));
        }
    }

    outHex.reserve(kHexBytesPerLine * 3);
    outAscii.reserve(kHexBytesPerLine);

    static constexpr wchar_t kHexDigits[] = L"0123456789ABCDEF";

    const size_t groupSize  = HexGroupSize();
    const size_t groupCount = (kHexBytesPerLine + groupSize - 1) / groupSize;
    outHex.reserve(groupCount * (groupSize * 2u + 1u));

    const bool bigEndian = HexBigEndian();

    for (size_t group = 0; group < groupCount; ++group)
    {
        const size_t groupStart     = group * groupSize;
        const size_t available      = (groupStart < count) ? std::min(groupSize, count - groupStart) : 0;
        const size_t groupCharStart = outHex.size();

        for (size_t pos = 0; pos < groupSize; ++pos)
        {
            const size_t byteIndex = bigEndian ? (groupStart + pos) : (groupStart + (groupSize - 1u - pos));
            if (byteIndex < groupStart + available && byteIndex < bytes.size())
            {
                const uint8_t b = bytes[byteIndex];
                outHex.push_back(kHexDigits[b >> 4]);
                outHex.push_back(kHexDigits[b & 0x0Fu]);
                hexSpans[byteIndex].start  = groupCharStart + pos * 2u;
                hexSpans[byteIndex].length = 2u;
            }
            else
            {
                outHex.append(L"  ");
            }
        }

        outHex.push_back(L' ');
    }

    std::array<wchar_t, kHexBytesPerLine> textChars{};
    textChars.fill(L' ');

    if (_hexTextMode == HexTextMode::Utf16)
    {
        for (size_t i = 0; i + 1 < count; i += 2)
        {
            const uint16_t value = bigEndian ? static_cast<uint16_t>(static_cast<uint16_t>(bytes[i] << 8u) | static_cast<uint16_t>(bytes[i + 1]))
                                             : static_cast<uint16_t>(static_cast<uint16_t>(bytes[i]) | static_cast<uint16_t>(bytes[i + 1] << 8u));
            wchar_t ch           = static_cast<wchar_t>(value);
            if (value >= 0xD800u && value <= 0xDFFFu)
            {
                ch = static_cast<wchar_t>(0xFFFDu);
            }

            if (ch < 32)
            {
                ch = L'.';
            }

            textChars[i] = ch;
            if ((i + 1) < textChars.size())
            {
                textChars[i + 1] = L' ';
            }
        }

        if ((count % 2u) == 1u)
        {
            const size_t i = count - 1;
            if (i < textChars.size())
            {
                textChars[i] = L'.';
            }
        }
    }
    else if (_hexTextMode == HexTextMode::Utf8)
    {
        auto decodeUtf8 = [](const uint8_t* data, size_t available, wchar_t& outChar, size_t& consumed) noexcept -> bool
        {
            outChar  = L'.';
            consumed = 1;

            if (! data || available == 0)
            {
                return false;
            }

            const uint8_t b0 = data[0];
            if (b0 <= 0x7Fu)
            {
                outChar  = static_cast<wchar_t>(b0);
                consumed = 1;
                return true;
            }

            if (b0 >= 0xC2u && b0 <= 0xDFu && available >= 2 && (data[1] & 0xC0u) == 0x80u)
            {
                const uint32_t cp = (static_cast<uint32_t>(b0 & 0x1Fu) << 6) | static_cast<uint32_t>(data[1] & 0x3Fu);
                outChar           = static_cast<wchar_t>(cp);
                consumed          = 2;
                return true;
            }

            if (b0 >= 0xE0u && b0 <= 0xEFu && available >= 3 && (data[1] & 0xC0u) == 0x80u && (data[2] & 0xC0u) == 0x80u)
            {
                if (b0 == 0xE0u && data[1] < 0xA0u)
                {
                    return false;
                }
                if (b0 == 0xEDu && data[1] >= 0xA0u)
                {
                    return false;
                }

                uint32_t cp =
                    (static_cast<uint32_t>(b0 & 0x0Fu) << 12) | (static_cast<uint32_t>(data[1] & 0x3Fu) << 6) | static_cast<uint32_t>(data[2] & 0x3Fu);

                if (cp >= 0xD800u && cp <= 0xDFFFu)
                {
                    cp = 0xFFFDu;
                }

                outChar  = static_cast<wchar_t>(cp);
                consumed = 3;
                return true;
            }

            if (b0 >= 0xF0u && b0 <= 0xF4u && available >= 4 && (data[1] & 0xC0u) == 0x80u && (data[2] & 0xC0u) == 0x80u && (data[3] & 0xC0u) == 0x80u)
            {
                uint32_t cp = (static_cast<uint32_t>(b0 & 0x07u) << 18) | (static_cast<uint32_t>(data[1] & 0x3Fu) << 12) |
                              (static_cast<uint32_t>(data[2] & 0x3Fu) << 6) | static_cast<uint32_t>(data[3] & 0x3Fu);

                if (b0 == 0xF0u && data[1] < 0x90u)
                {
                    return false;
                }
                if (b0 == 0xF4u && data[1] >= 0x90u)
                {
                    return false;
                }

                if (cp > 0x10FFFFu)
                {
                    cp = 0xFFFDu;
                }

                outChar  = (cp <= 0xFFFFu) ? static_cast<wchar_t>(cp) : static_cast<wchar_t>(0xFFFDu);
                consumed = 4;
                return true;
            }

            return false;
        };

        size_t i = 0;
        while (i < count)
        {
            wchar_t ch      = L'.';
            size_t consumed = 1;
            const bool ok   = decodeUtf8(bytes.data() + i, count - i, ch, consumed);
            if (! ok || consumed == 0)
            {
                consumed = 1;
                ch       = L'.';
            }

            if (ch < 32)
            {
                ch = L'.';
            }

            for (size_t j = 0; j < consumed && (i + j) < textChars.size(); ++j)
            {
                textChars[i + j] = (j == 0) ? ch : L' ';
            }

            i += consumed;
        }
    }
    else
    {
        UINT codePage = DisplayEncodingCodePage();
        if (codePage == CP_UTF8)
        {
            codePage = CP_ACP;
        }

        for (size_t i = 0; i < count; ++i)
        {
            const char src = static_cast<char>(bytes[i]);
            wchar_t wideBuf[2]{};
            wchar_t ch        = L'.';
            const int written = MultiByteToWideChar(codePage, MB_ERR_INVALID_CHARS, &src, 1, wideBuf, static_cast<int>(std::size(wideBuf)));
            if (written > 0)
            {
                ch = wideBuf[0];
            }
            else if (bytes[i] >= 32u && bytes[i] <= 126u)
            {
                ch = static_cast<wchar_t>(bytes[i]);
            }

            if (ch < 32)
            {
                ch = L'.';
            }

            textChars[i] = ch;
        }
    }

    for (size_t i = 0; i < kHexBytesPerLine; ++i)
    {
        textSpans[i].start = outAscii.size();
        if (i < count)
        {
            textSpans[i].length = 1u;
            outAscii.push_back(textChars[i]);
        }
        else
        {
            textSpans[i].length = 0;
            outAscii.push_back(L' ');
        }
    }
}

void ViewerText::OnHexMouseDown(HWND hwnd, int x, int y) noexcept
{
    if (! _hHex || hwnd != _hHex.get())
    {
        return;
    }

    if (_fileSize == 0)
    {
        MessageBeep(MB_ICONINFORMATION);
        return;
    }

    static_cast<void>(EnsureHexViewDirect2D(hwnd));

    const UINT dpi   = GetDpiForWindow(hwnd);
    const float xDip = DipsFromPixels(x, dpi);
    const float yDip = DipsFromPixels(y, dpi);

    const float marginDip      = 6.0f;
    const float charW          = (_hexCharWidthDip > 0.0f) ? _hexCharWidthDip : 8.0f;
    const float lineH          = (_hexLineHeightDip > 0.0f) ? _hexLineHeightDip : 14.0f;
    const HexViewLayout layout = ComputeHexViewLayout(lineH, charW, _fileSize, HexGroupSize());
    const float headerH        = layout.headerH;
    if (charW <= 0.0f || lineH <= 0.0f)
    {
        return;
    }

    const float xOffset      = layout.xOffset;
    const float xHex         = layout.xHex;
    const float xText        = layout.xText;
    const float hexTextRight = layout.hexTextRight;

    if (yDip >= marginDip && yDip < marginDip + headerH)
    {
        if (xDip >= xOffset && xDip < xHex)
        {
            CycleHexOffsetMode();
        }
        else if (xDip >= xHex && xDip < xText)
        {
            CycleHexColumnMode();
        }
        else if (xDip >= xText)
        {
            CycleHexTextMode();
        }

        SetFocus(hwnd);
        return;
    }

    const float relY      = std::max(0.0f, yDip - layout.dataStartY);
    const uint64_t row    = static_cast<uint64_t>(std::floor(relY / lineH));
    const uint64_t line   = _hexTopLine + row;
    const uint64_t offset = line * static_cast<uint64_t>(kHexBytesPerLine);
    if (offset >= _fileSize)
    {
        return;
    }

    const bool hitHexColumn  = xDip >= xHex && xDip < hexTextRight;
    const bool hitTextColumn = xDip >= xText;
    if (! hitHexColumn && ! hitTextColumn)
    {
        return;
    }

    std::array<uint8_t, kHexBytesPerLine> bytes{};
    const size_t validBytes = ReadHexBytes(offset, bytes.data(), bytes.size());
    if (validBytes == 0)
    {
        return;
    }

    const float relX       = hitHexColumn ? std::max(0.0f, xDip - xHex) : std::max(0.0f, xDip - xText);
    const size_t charIndex = static_cast<size_t>(std::floor(relX / charW));

    size_t found = 0;
    if (hitTextColumn)
    {
        found = std::min<size_t>(charIndex, validBytes - 1u);
    }
    else
    {
        const size_t groupSize  = HexGroupSize();
        const size_t groupUnit  = groupSize * 2u + 1u;
        const size_t group      = groupUnit > 0 ? (charIndex / groupUnit) : 0;
        const size_t within     = groupUnit > 0 ? (charIndex % groupUnit) : 0;
        const size_t groupStart = group * groupSize;
        const size_t posByte    = (within >= groupSize * 2u) ? (groupSize - 1u) : (within / 2u);

        const bool bigEndian   = HexBigEndian();
        const size_t byteIndex = bigEndian ? (groupStart + posByte) : (groupStart + (groupSize - 1u - posByte));
        found                  = std::min<size_t>(byteIndex, validBytes - 1u);
    }

    const uint64_t clickedOffset = offset + static_cast<uint64_t>(found);
    const bool shiftDown         = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    if (! shiftDown || ! _hexSelectedOffset.has_value())
    {
        _hexSelectionAnchorOffset = clickedOffset;
    }
    else if (! _hexSelectionAnchorOffset.has_value())
    {
        _hexSelectionAnchorOffset = _hexSelectedOffset.value();
    }

    _hexSelectedOffset = clickedOffset;
    _hexSelecting      = true;
    SetCapture(hwnd);

    SetFocus(hwnd);
    InvalidateRect(hwnd, nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
}

void ViewerText::OnHexMouseMove(HWND hwnd, int x, int y) noexcept
{
    if (! _hHex || hwnd != _hHex.get())
    {
        return;
    }

    static_cast<void>(EnsureHexViewDirect2D(hwnd));

    const UINT dpi   = GetDpiForWindow(hwnd);
    const float xDip = DipsFromPixels(x, dpi);
    const float yDip = DipsFromPixels(y, dpi);

    const float charW          = (_hexCharWidthDip > 0.0f) ? _hexCharWidthDip : 8.0f;
    const float lineH          = (_hexLineHeightDip > 0.0f) ? _hexLineHeightDip : 14.0f;
    const HexViewLayout layout = ComputeHexViewLayout(lineH, charW, _fileSize, HexGroupSize());
    if (charW <= 0.0f || lineH <= 0.0f)
    {
        return;
    }

    HexHeaderHit hotHeader = HexHeaderHit::None;
    if (yDip >= layout.headerY && yDip < layout.headerY + layout.headerH)
    {
        if (xDip >= layout.xOffset && xDip < layout.xHex)
        {
            hotHeader = HexHeaderHit::Offset;
        }
        else if (xDip >= layout.xHex && xDip < layout.xText)
        {
            hotHeader = HexHeaderHit::Data;
        }
        else if (xDip >= layout.xText)
        {
            hotHeader = HexHeaderHit::Text;
        }
    }

    if (hotHeader != _hexHeaderHot)
    {
        _hexHeaderHot = hotHeader;
        InvalidateRect(hwnd, nullptr, FALSE);
    }

    if (! _hexSelecting)
    {
        return;
    }

    if ((GetKeyState(VK_LBUTTON) & 0x8000) == 0)
    {
        _hexSelecting = false;
        ReleaseCapture();
        return;
    }

    if (_fileSize == 0)
    {
        return;
    }

    if (yDip < layout.dataStartY)
    {
        return;
    }

    const float xHex         = layout.xHex;
    const float xText        = layout.xText;
    const float hexTextRight = layout.hexTextRight;

    const bool hitHexColumn  = xDip >= xHex && xDip < hexTextRight;
    const bool hitTextColumn = xDip >= xText;
    if (! hitHexColumn && ! hitTextColumn)
    {
        return;
    }

    const float relY      = std::max(0.0f, yDip - layout.dataStartY);
    const uint64_t row    = static_cast<uint64_t>(std::floor(relY / lineH));
    const uint64_t line   = _hexTopLine + row;
    const uint64_t offset = line * static_cast<uint64_t>(kHexBytesPerLine);
    if (offset >= _fileSize)
    {
        return;
    }

    std::array<uint8_t, kHexBytesPerLine> bytes{};
    const size_t validBytes = ReadHexBytes(offset, bytes.data(), bytes.size());
    if (validBytes == 0)
    {
        return;
    }

    const float relX       = hitHexColumn ? std::max(0.0f, xDip - xHex) : std::max(0.0f, xDip - xText);
    const size_t charIndex = static_cast<size_t>(std::floor(relX / charW));

    size_t found = 0;
    if (hitTextColumn)
    {
        found = std::min<size_t>(charIndex, validBytes - 1u);
    }
    else
    {
        const size_t groupSize  = HexGroupSize();
        const size_t groupUnit  = groupSize * 2u + 1u;
        const size_t group      = groupUnit > 0 ? (charIndex / groupUnit) : 0;
        const size_t within     = groupUnit > 0 ? (charIndex % groupUnit) : 0;
        const size_t groupStart = group * groupSize;
        const size_t posByte    = (within >= groupSize * 2u) ? (groupSize - 1u) : (within / 2u);

        const bool bigEndian   = HexBigEndian();
        const size_t byteIndex = bigEndian ? (groupStart + posByte) : (groupStart + (groupSize - 1u - posByte));
        found                  = std::min<size_t>(byteIndex, validBytes - 1u);
    }

    const uint64_t newOffset = offset + static_cast<uint64_t>(found);
    if (_hexSelectedOffset.has_value() && _hexSelectedOffset.value() == newOffset)
    {
        return;
    }

    _hexSelectedOffset = newOffset;

    InvalidateRect(hwnd, nullptr, TRUE);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
    }
}

void ViewerText::OnHexMouseUp(HWND hwnd) noexcept
{
    if (! _hHex || hwnd != _hHex.get())
    {
        return;
    }

    if (! _hexSelecting)
    {
        return;
    }

    _hexSelecting = false;
    ReleaseCapture();
}

void ViewerText::CopyHexCsvToClipboard(HWND hwnd) noexcept
{
    if (! _hHex || hwnd != _hHex.get())
    {
        return;
    }

    if (_fileSize == 0)
    {
        MessageBeep(MB_ICONINFORMATION);
        return;
    }

    const uint64_t totalLines = (_fileSize + (kHexBytesPerLine - 1)) / kHexBytesPerLine;
    if (totalLines == 0)
    {
        MessageBeep(MB_ICONINFORMATION);
        return;
    }

    uint64_t startLine = 0;
    uint64_t endLine   = 0;
    if (_hexSelectedOffset.has_value())
    {
        const uint64_t active         = _hexSelectedOffset.value();
        const uint64_t anchor         = _hexSelectionAnchorOffset.has_value() ? _hexSelectionAnchorOffset.value() : active;
        const uint64_t selectionStart = std::min(anchor, active);
        const uint64_t selectionEnd   = std::max(anchor, active);
        startLine                     = selectionStart / static_cast<uint64_t>(kHexBytesPerLine);
        endLine                       = selectionEnd / static_cast<uint64_t>(kHexBytesPerLine);
    }
    else
    {
        static_cast<void>(EnsureHexViewDirect2D(hwnd));
        RECT client{};
        GetClientRect(hwnd, &client);

        const UINT dpi        = GetDpiForWindow(hwnd);
        const float heightDip = std::max(1.0f, DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi));
        const float marginDip = 6.0f;
        const float lineH     = (_hexLineHeightDip > 0.0f) ? _hexLineHeightDip : 14.0f;
        const float headerH   = lineH;
        const float usableDip = std::max(0.0f, heightDip - headerH - 2.0f * marginDip);
        const uint64_t rows   = std::max<uint64_t>(1u, static_cast<uint64_t>(std::ceil(usableDip / std::max(1.0f, lineH))));

        startLine = std::min<uint64_t>(_hexTopLine, totalLines - 1);
        endLine   = std::min<uint64_t>(totalLines - 1, startLine + rows - 1);
    }

    const std::wstring colOffset = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_COL_OFFSET);
    const std::wstring colHex    = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_COL_HEX);
    UINT colTextId               = IDS_VIEWERTEXT_COL_TEXT_ANSI;
    switch (_hexTextMode)
    {
        case HexTextMode::Utf8: colTextId = IDS_VIEWERTEXT_COL_TEXT_UTF8; break;
        case HexTextMode::Utf16: colTextId = IDS_VIEWERTEXT_COL_TEXT_UTF16; break;
        case HexTextMode::Ansi:
        default: colTextId = IDS_VIEWERTEXT_COL_TEXT_ANSI; break;
    }
    const std::wstring colText = LoadStringResource(g_hInstance, colTextId);

    std::wstring csv;
    const uint64_t lineCount           = (endLine >= startLine) ? (endLine - startLine + 2) : 2;
    constexpr uint64_t maxReserveLines = static_cast<uint64_t>(std::numeric_limits<size_t>::max() / 128u);
    const uint64_t reserveLines        = std::min<uint64_t>(lineCount, maxReserveLines);
    csv.reserve(static_cast<size_t>(reserveLines) * 128u);
    csv.append(CsvEscape(colOffset));
    csv.push_back(L',');
    csv.append(CsvEscape(colHex));
    csv.push_back(L',');
    csv.append(CsvEscape(colText));
    csv.append(L"\r\n");

    for (uint64_t line = startLine; line <= endLine; ++line)
    {
        const uint64_t offset = line * static_cast<uint64_t>(kHexBytesPerLine);

        std::wstring offsetText;
        std::wstring hexText;
        std::wstring asciiText;
        std::array<ByteSpan, kHexBytesPerLine> hexSpans{};
        std::array<ByteSpan, kHexBytesPerLine> textSpans{};
        size_t validBytes = 0;
        FormatHexLine(offset, offsetText, hexText, asciiText, hexSpans, textSpans, validBytes);
        static_cast<void>(validBytes);

        csv.append(CsvEscape(offsetText));
        csv.push_back(L',');
        csv.append(CsvEscape(hexText));
        csv.push_back(L',');
        csv.append(CsvEscape(asciiText));
        csv.append(L"\r\n");
    }

    if (! CopyUnicodeTextToClipboard(GetAncestor(hwnd, GA_ROOT), csv))
    {
        MessageBeep(MB_ICONERROR);
    }
}
