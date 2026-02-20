#include "ViewerText.h"

#include "ViewerText.ThemeHelpers.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <cerrno>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <format>
#include <functional>
#include <iterator>
#include <limits>
#include <memory>
#include <new>
#include <string_view>
#include <thread>

#include <commctrl.h>
#include <commdlg.h>
#include <d2d1.h>
#include <d2d1helper.h>
#include <dwmapi.h>
#include <dwrite.h>
#include <richedit.h>
#include <shobjidl_core.h>
#include <uxtheme.h>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#pragma comment(lib, "comctl32")
#pragma comment(lib, "d2d1")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "dwrite")
#pragma comment(lib, "uxtheme")

#include "Helpers.h"
#include "WindowMessages.h"

#include "resource.h"

extern HINSTANCE g_hInstance;

namespace
{
constexpr int kHeaderHeightDip           = 28;
constexpr int kStatusHeightDip           = 22;
constexpr float kWatermarkAngleDegrees   = -22.0f;
constexpr float kWatermarkFontSizeDip    = 56.0f;
constexpr float kWatermarkAngleRadians   = kWatermarkAngleDegrees * 0.01745329252f;
constexpr uint64_t kMaxHexLoadBytes      = 128u * 1024u * 1024u; // 128 MiB
constexpr UINT kAsyncOpenCompleteMessage = WndMsg::kViewerTextAsyncOpenComplete;
constexpr UINT kLoadingDelayMs           = 500u;
constexpr UINT kLoadingAnimIntervalMs    = 16u;
constexpr float kLoadingSpinnerDegPerSec = 90.0f;

static const int kViewerTextModuleAnchor = 0;

constexpr UINT_PTR kFileComboEscCloseSubclassId = 1u;

LRESULT CALLBACK
FileComboEscCloseSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, [[maybe_unused]] UINT_PTR subclassId, [[maybe_unused]] DWORD_PTR refData) noexcept
{
    if (msg == WM_KEYDOWN && wp == VK_ESCAPE)
    {
        const bool dropped = SendMessageW(hwnd, CB_GETDROPPEDSTATE, 0, 0) != 0;
        if (! dropped)
        {
            const HWND root = GetAncestor(hwnd, GA_ROOT);
            if (root)
            {
                PostMessageW(root, WM_CLOSE, 0, 0);
            }
            return 0;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

void InstallFileComboEscClose(HWND combo) noexcept
{
    if (! combo)
    {
        return;
    }

    static_cast<void>(SetWindowSubclass(combo, FileComboEscCloseSubclassProc, kFileComboEscCloseSubclassId, 0));
}

int HexNibbleValue(wchar_t ch) noexcept
{
    if (ch >= L'0' && ch <= L'9')
    {
        return static_cast<int>(ch - L'0');
    }
    if (ch >= L'a' && ch <= L'f')
    {
        return 10 + static_cast<int>(ch - L'a');
    }
    if (ch >= L'A' && ch <= L'F')
    {
        return 10 + static_cast<int>(ch - L'A');
    }
    return -1;
}

bool TryParseHexSearchNeedle(std::wstring_view query, std::vector<uint8_t>& outBytes) noexcept
{
    outBytes.clear();

    std::wstring digits;
    digits.reserve(query.size());

    for (size_t i = 0; i < query.size(); ++i)
    {
        const wchar_t ch = query[i];
        if (ch == L'0' && (i + 1) < query.size() && (query[i + 1] == L'x' || query[i + 1] == L'X'))
        {
            i += 1;
            continue;
        }

        if (iswspace(ch) != 0 || ch == L',' || ch == L';' || ch == L':' || ch == L'_')
        {
            continue;
        }

        if (HexNibbleValue(ch) >= 0)
        {
            digits.push_back(ch);
            continue;
        }

        return false;
    }

    if (digits.empty())
    {
        return false;
    }

    if ((digits.size() % 2u) == 1u)
    {
        digits.insert(digits.begin(), L'0');
    }

    outBytes.reserve(digits.size() / 2u);
    for (size_t i = 0; i + 1 < digits.size(); i += 2)
    {
        const int hi = HexNibbleValue(digits[i]);
        const int lo = HexNibbleValue(digits[i + 1]);
        if (hi < 0 || lo < 0)
        {
            outBytes.clear();
            return false;
        }

        outBytes.push_back(static_cast<uint8_t>((hi << 4u) | lo));
    }

    return ! outBytes.empty();
}

constexpr char kViewerTextSchemaJson[] = R"json({
    "version": 1,
    "title": "Text Viewer",
    "fields": [
        {
            "key": "textBufferMiB",
            "type": "value",
            "label": "Text buffer (MiB)",
            "description": "Approximate in-memory read buffer used by the streaming text renderer.",
            "default": 16,
            "min": 1,
            "max": 256
        },
        {
            "key": "hexBufferMiB",
            "type": "value",
            "label": "Hex buffer (MiB)",
            "description": "Approximate in-memory read buffer used by the streaming hex renderer.",
            "default": 8,
            "min": 1,
            "max": 256
        },
        {
            "key": "showLineNumbers",
            "type": "option",
            "label": "Line numbers",
            "description": "Show logical line numbers (newline-delimited).",
            "default": "0",
            "options": [
                { "value": "0", "label": "Off" },
                { "value": "1", "label": "On" }
            ]
        },
        {
            "key": "wrapText",
            "type": "option",
            "label": "Wrap",
            "description": "Wrap long lines in text mode.",
            "default": "1",
            "options": [
                { "value": "0", "label": "Off" },
                { "value": "1", "label": "On" }
            ]
        }
    ]
})json";

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

bool LooksLikeBinaryData(const uint8_t* data, size_t size) noexcept
{
    if (! data || size == 0)
    {
        return false;
    }

    constexpr size_t kMaxProbeBytes = 64u * 1024u;
    const size_t probeSize          = std::min(size, kMaxProbeBytes);

    size_t suspiciousControls = 0;
    for (size_t i = 0; i < probeSize; ++i)
    {
        const uint8_t b = data[i];
        if (b == 0)
        {
            return true;
        }

        if (b < 0x20u)
        {
            // Allow common whitespace/control used in text files.
            if (b == 0x09u || b == 0x0Au || b == 0x0Cu || b == 0x0Du)
            {
                continue;
            }
            suspiciousControls += 1;
            continue;
        }

        if (b == 0x7Fu)
        {
            suspiciousControls += 1;
            continue;
        }
    }

    const double ratio = static_cast<double>(suspiciousControls) / static_cast<double>(probeSize);
    return ratio > 0.25;
}

ViewerText::FileEncoding DisplayEncodingFileEncodingForSelection(UINT selection) noexcept
{
    switch (selection)
    {
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM: return ViewerText::FileEncoding::Utf8;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM: return ViewerText::FileEncoding::Utf16BE;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM: return ViewerText::FileEncoding::Utf16LE;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM: return ViewerText::FileEncoding::Utf32BE;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM: return ViewerText::FileEncoding::Utf32LE;
        default: return ViewerText::FileEncoding::Unknown;
    }
}

UINT CodePageForSelection(UINT selection) noexcept
{
    switch (selection)
    {
        case IDM_VIEWER_ENCODING_DISPLAY_ANSI: return CP_ACP;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF7: return 65000u;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM: return CP_UTF8;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM: return CP_ACP;
        default: break;
    }

    return selection;
}

uint64_t BytesToSkipForDisplayEncoding(UINT selection, ViewerText::FileEncoding encoding, uint64_t bomBytes) noexcept
{
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM && encoding == ViewerText::FileEncoding::Utf8 && bomBytes == 3)
    {
        return 3u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM && encoding == ViewerText::FileEncoding::Utf16LE && bomBytes == 2)
    {
        return 2u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM && encoding == ViewerText::FileEncoding::Utf16BE && bomBytes == 2)
    {
        return 2u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM && encoding == ViewerText::FileEncoding::Utf32LE && bomBytes == 4)
    {
        return 4u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM && encoding == ViewerText::FileEncoding::Utf32BE && bomBytes == 4)
    {
        return 4u;
    }

    return 0;
}

uint64_t TextStreamChunkBytes(uint32_t textBufferMiB, ViewerText::FileEncoding displayEncoding) noexcept
{
    uint64_t bytes = static_cast<uint64_t>(textBufferMiB) * 1024u * 1024u;
    bytes          = std::clamp<uint64_t>(bytes, 256u * 1024u, 256u * 1024u * 1024u);

    if (displayEncoding == ViewerText::FileEncoding::Utf16LE || displayEncoding == ViewerText::FileEncoding::Utf16BE)
    {
        bytes &= ~static_cast<uint64_t>(1);
        bytes = std::max<uint64_t>(bytes, 2u);
    }
    else if (displayEncoding == ViewerText::FileEncoding::Utf32LE || displayEncoding == ViewerText::FileEncoding::Utf32BE)
    {
        bytes &= ~static_cast<uint64_t>(3);
        bytes = std::max<uint64_t>(bytes, 4u);
    }

    return bytes;
}

size_t Utf8IncompleteTailSize(const uint8_t* data, size_t size) noexcept
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
}

void BuildTextLineIndex(std::wstring_view text, std::vector<uint32_t>& outLineStarts, std::vector<uint32_t>& outLineEnds, uint32_t& outMaxLineLength) noexcept
{
    outLineStarts.clear();
    outLineEnds.clear();
    outMaxLineLength = 0;

    const size_t size = text.size();
    size_t start      = 0;

    for (;;)
    {
        size_t pos = start;
        while (pos < size)
        {
            const wchar_t ch = text[pos];
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

        outLineStarts.push_back(start32);
        outLineEnds.push_back(end32);

        if (end32 >= start32)
        {
            outMaxLineLength = std::max(outMaxLineLength, end32 - start32);
        }

        if (pos >= size)
        {
            break;
        }

        if (text[pos] == L'\r' && (pos + 1) < size && text[pos + 1] == L'\n')
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

    if (outLineStarts.empty())
    {
        outLineStarts.push_back(0);
        outLineEnds.push_back(0);
    }
}

uint32_t MakeBgra(uint8_t r, uint8_t g, uint8_t b, uint8_t a) noexcept
{
    return static_cast<uint32_t>(b) | (static_cast<uint32_t>(g) << 8) | (static_cast<uint32_t>(r) << 16) | (static_cast<uint32_t>(a) << 24);
}

bool PointInRoundedRect(int x, int y, int left, int top, int right, int bottom, int radius) noexcept
{
    if (x < left || x >= right || y < top || y >= bottom)
    {
        return false;
    }

    const int r = std::max(0, radius);
    if (r == 0)
    {
        return true;
    }

    const int innerLeft   = left + r;
    const int innerTop    = top + r;
    const int innerRight  = right - r;
    const int innerBottom = bottom - r;

    if (x >= innerLeft && x < innerRight)
    {
        return true;
    }
    if (y >= innerTop && y < innerBottom)
    {
        return true;
    }

    const auto inCorner = [&](int cx, int cy) noexcept
    {
        const int dx = x - cx;
        const int dy = y - cy;
        return (dx * dx + dy * dy) <= (r * r);
    };

    if (x < innerLeft && y < innerTop)
    {
        return inCorner(innerLeft, innerTop);
    }
    if (x >= innerRight && y < innerTop)
    {
        return inCorner(innerRight - 1, innerTop);
    }
    if (x < innerLeft && y >= innerBottom)
    {
        return inCorner(innerLeft, innerBottom - 1);
    }
    if (x >= innerRight && y >= innerBottom)
    {
        return inCorner(innerRight - 1, innerBottom - 1);
    }

    return true;
}

wil::unique_hicon CreateViewerTextIcon(int sizePx) noexcept
{
    if (sizePx <= 0 || sizePx > 256)
    {
        return {};
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = sizePx;
    bmi.bmiHeader.biHeight      = -sizePx;
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    wil::unique_hbitmap color(CreateDIBSection(nullptr, &bmi, DIB_RGB_COLORS, &bits, nullptr, 0));
    if (! color || ! bits)
    {
        return {};
    }

    auto* pixels            = static_cast<uint32_t*>(bits);
    const size_t pixelCount = static_cast<size_t>(sizePx) * static_cast<size_t>(sizePx);
    std::fill_n(pixels, pixelCount, 0u);

    const COLORREF baseRef   = RGB(0, 120, 215);
    const COLORREF borderRef = RGB(0, 90, 160);
    const COLORREF lineRef   = RGB(255, 255, 255);

    const uint32_t lineRgb   = static_cast<uint32_t>(lineRef);
    const uint8_t lineR      = static_cast<uint8_t>(lineRgb & 0xFFu);
    const uint8_t lineG      = static_cast<uint8_t>((lineRgb >> 8) & 0xFFu);
    const uint8_t lineB      = static_cast<uint8_t>((lineRgb >> 16) & 0xFFu);
    const uint32_t linePixel = MakeBgra(lineR, lineG, lineB, 255u);

    const int margin = std::max(1, sizePx / 8);
    const int left   = margin;
    const int top    = margin;
    const int right  = sizePx - margin;
    const int bottom = sizePx - margin;
    const int radius = std::max(2, sizePx / 6);

    const int border      = std::max(1, sizePx / 16);
    const int innerLeft   = left + border;
    const int innerTop    = top + border;
    const int innerRight  = right - border;
    const int innerBottom = bottom - border;
    const int innerRadius = std::max(0, radius - border);

    for (int y = 0; y < sizePx; ++y)
    {
        for (int x = 0; x < sizePx; ++x)
        {
            if (! PointInRoundedRect(x, y, left, top, right, bottom, radius))
            {
                continue;
            }

            const bool inInner = PointInRoundedRect(x, y, innerLeft, innerTop, innerRight, innerBottom, innerRadius);
            const COLORREF c   = inInner ? baseRef : borderRef;
            pixels[static_cast<size_t>(y) * static_cast<size_t>(sizePx) + static_cast<size_t>(x)] =
                MakeBgra(static_cast<uint8_t>(GetRValue(c)), static_cast<uint8_t>(GetGValue(c)), static_cast<uint8_t>(GetBValue(c)), 255u);
        }
    }

    const int lineLeft   = innerLeft + std::max(1, sizePx / 8);
    const int lineRight  = innerRight - std::max(1, sizePx / 8);
    const int lineHeight = std::max(1, sizePx / 14);
    const int lineGap    = std::max(1, sizePx / 10);
    const int firstLineY = innerTop + std::max(1, sizePx / 6);

    for (int i = 0; i < 3; ++i)
    {
        const int y0 = firstLineY + i * (lineHeight + lineGap);
        for (int y = y0; y < y0 + lineHeight; ++y)
        {
            if (y < innerTop || y >= innerBottom)
            {
                continue;
            }

            for (int x = lineLeft; x < lineRight; ++x)
            {
                if (x < innerLeft || x >= innerRight)
                {
                    continue;
                }

                pixels[static_cast<size_t>(y) * static_cast<size_t>(sizePx) + static_cast<size_t>(x)] = linePixel;
            }
        }
    }

    const size_t maskStride = static_cast<size_t>(((sizePx + 31) / 32) * 4);
    std::vector<uint8_t> maskBits(maskStride * static_cast<size_t>(sizePx), 0u);
    wil::unique_hbitmap mask(CreateBitmap(sizePx, sizePx, 1, 1, maskBits.data()));
    if (! mask)
    {
        return {};
    }

    ICONINFO ii{};
    ii.fIcon    = TRUE;
    ii.hbmColor = color.get();
    ii.hbmMask  = mask.get();

    return wil::unique_hicon(CreateIconIndirect(&ii));
}

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

int PxFromDip(int dip, UINT dpi) noexcept
{
    return MulDiv(dip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
}

float DipsFromPixels(int px, UINT dpi) noexcept
{
    if (dpi == 0)
    {
        return static_cast<float>(px);
    }

    return static_cast<float>(px) * 96.0f / static_cast<float>(dpi);
}

D2D1_RECT_F RectFFromPixels(const RECT& rc, UINT dpi) noexcept
{
    const float left   = DipsFromPixels(static_cast<int>(rc.left), dpi);
    const float top    = DipsFromPixels(static_cast<int>(rc.top), dpi);
    const float right  = DipsFromPixels(static_cast<int>(rc.right), dpi);
    const float bottom = DipsFromPixels(static_cast<int>(rc.bottom), dpi);
    return D2D1::RectF(left, top, right, bottom);
}

D2D1_COLOR_F ColorFFromColorRef(COLORREF color, float alpha = 1.0f) noexcept
{
    const float r = static_cast<float>(GetRValue(color)) / 255.0f;
    const float g = static_cast<float>(GetGValue(color)) / 255.0f;
    const float b = static_cast<float>(GetBValue(color)) / 255.0f;
    return D2D1::ColorF(r, g, b, alpha);
}

void ClampRectNonNegative(RECT& rc) noexcept
{
    if (rc.right < rc.left)
    {
        rc.right = rc.left;
    }
    if (rc.bottom < rc.top)
    {
        rc.bottom = rc.top;
    }
}

struct FindDialogState
{
    ViewerText* viewer = nullptr;
    std::wstring initial;
    std::wstring result;
};

std::wstring ReadDialogItemText(HWND dlg, int controlId)
{
    const HWND control = GetDlgItem(dlg, controlId);
    if (! control)
    {
        return {};
    }

    const int length = GetWindowTextLengthW(control);
    if (length <= 0)
    {
        return {};
    }

    std::wstring text;
    text.resize(static_cast<size_t>(length) + 1u);
    GetWindowTextW(control, text.data(), length + 1);
    text.resize(static_cast<size_t>(length));
    return text;
}

INT_PTR OnFindDialogInit(HWND dlg, FindDialogState* state)
{
    SetWindowLongPtrW(dlg, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
    if (state)
    {
        SetDlgItemTextW(dlg, IDC_VIEWERTEXT_FIND_TEXT, state->initial.c_str());
        SendDlgItemMessageW(dlg, IDC_VIEWERTEXT_FIND_TEXT, EM_SETSEL, 0, -1);
        SetFocus(GetDlgItem(dlg, IDC_VIEWERTEXT_FIND_TEXT));
    }
    return FALSE;
}

INT_PTR OnFindDialogCommand(HWND dlg, UINT commandId)
{
    if (commandId == IDOK)
    {
        auto* state = reinterpret_cast<FindDialogState*>(GetWindowLongPtrW(dlg, GWLP_USERDATA));
        if (state)
        {
            state->result = ReadDialogItemText(dlg, IDC_VIEWERTEXT_FIND_TEXT);
        }
        EndDialog(dlg, IDOK);
        return TRUE;
    }

    if (commandId == IDCANCEL)
    {
        EndDialog(dlg, IDCANCEL);
        return TRUE;
    }

    return FALSE;
}

INT_PTR CALLBACK FindDlgProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_INITDIALOG: return OnFindDialogInit(dlg, reinterpret_cast<FindDialogState*>(lp));
        case WM_COMMAND: return OnFindDialogCommand(dlg, LOWORD(wp));
    }

    return FALSE;
}

struct GoToDialogState
{
    std::optional<uint64_t> offset;
};

INT_PTR OnGoToDialogInit(HWND dlg, GoToDialogState* state)
{
    SetWindowLongPtrW(dlg, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
    SetDlgItemInt(dlg, IDC_VIEWERTEXT_GOTO_OFFSET, 0, FALSE);
    SendDlgItemMessageW(dlg, IDC_VIEWERTEXT_GOTO_OFFSET, EM_SETSEL, 0, -1);
    SetFocus(GetDlgItem(dlg, IDC_VIEWERTEXT_GOTO_OFFSET));
    return FALSE;
}

bool TryParseOffset(std::wstring_view text, uint64_t& value) noexcept;

INT_PTR OnGoToDialogCommand(HWND dlg, UINT commandId)
{
    if (commandId == IDOK)
    {
        auto* state = reinterpret_cast<GoToDialogState*>(GetWindowLongPtrW(dlg, GWLP_USERDATA));
        if (state)
        {
            const std::wstring text = ReadDialogItemText(dlg, IDC_VIEWERTEXT_GOTO_OFFSET);
            uint64_t value          = 0;
            if (TryParseOffset(text, value))
            {
                state->offset = value;
            }
        }

        EndDialog(dlg, IDOK);
        return TRUE;
    }

    if (commandId == IDCANCEL)
    {
        EndDialog(dlg, IDCANCEL);
        return TRUE;
    }

    return FALSE;
}

bool TryParseOffset(std::wstring_view text, uint64_t& value) noexcept
{
    std::wstring t(text);
    if (t.empty())
    {
        return false;
    }

    wchar_t* start = t.data();
    while (*start != L'\0' && std::iswspace(static_cast<wint_t>(*start)) != 0)
    {
        ++start;
    }

    if (*start == L'\0')
    {
        return false;
    }

    errno                           = 0;
    wchar_t* end                    = nullptr;
    const unsigned long long parsed = wcstoull(start, &end, 0);
    if (end == start)
    {
        return false;
    }
    if (errno == ERANGE)
    {
        return false;
    }

    while (*end != L'\0' && std::iswspace(static_cast<wint_t>(*end)) != 0)
    {
        ++end;
    }

    if (*end != L'\0')
    {
        return false;
    }

    value = static_cast<uint64_t>(parsed);
    return true;
}

INT_PTR CALLBACK GoToDlgProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_INITDIALOG: return OnGoToDialogInit(dlg, reinterpret_cast<GoToDialogState*>(lp));
        case WM_COMMAND: return OnGoToDialogCommand(dlg, LOWORD(wp));
    }

    return FALSE;
}

HRESULT WriteAllHandle(HANDLE file, const void* data, size_t size) noexcept
{
    if (! file)
    {
        return E_INVALIDARG;
    }

    const uint8_t* bytes = reinterpret_cast<const uint8_t*>(data);
    size_t offset        = 0;
    while (offset < size)
    {
        const DWORD want = static_cast<DWORD>(std::min<size_t>(size - offset, std::numeric_limits<DWORD>::max()));
        DWORD written    = 0;
        if (WriteFile(file, bytes + offset, want, &written, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (written == 0)
        {
            return E_FAIL;
        }
        offset += written;
    }

    return S_OK;
}
} // namespace

ViewerText::ViewerText()
{
    _metaId          = L"builtin/viewer-text";
    _metaShortId     = L"read";
    _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_NAME);
    _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DESCRIPTION);

    _displayEncodingMenuSelection = IDM_VIEWER_ENCODING_DISPLAY_ANSI;
    _saveEncodingMenuSelection    = IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL;

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = nullptr;

    static_cast<void>(SetConfiguration(nullptr));
}

ViewerText::~ViewerText() = default;

void ViewerText::SetHost(IHost* host) noexcept
{
    _hostAlerts = nullptr;

    if (! host)
    {
        return;
    }

    wil::com_ptr<IHostAlerts> alerts;
    const HRESULT hr = host->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
    if (SUCCEEDED(hr) && alerts)
    {
        _hostAlerts = std::move(alerts);
    }
}

HRESULT STDMETHODCALLTYPE ViewerText::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    *ppvObject = nullptr;

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IViewer))
    {
        *ppvObject = static_cast<IViewer*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE ViewerText::AddRef() noexcept
{
    return _refCount.fetch_add(1) + 1;
}

ULONG STDMETHODCALLTYPE ViewerText::Release() noexcept
{
    const ULONG remaining = _refCount.fetch_sub(1) - 1;
    if (remaining == 0)
    {
        delete this;
    }
    return remaining;
}

HRESULT STDMETHODCALLTYPE ViewerText::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = nullptr;

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = kViewerTextSchemaJson;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    uint32_t textBufferMiB = 16;
    uint32_t hexBufferMiB  = 8;
    bool showLineNumbers   = false;
    bool wrapText          = true;

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        const std::string_view utf8(configurationJsonUtf8);
        if (! utf8.empty())
        {
            yyjson_doc* doc = yyjson_read(utf8.data(), utf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
            if (doc)
            {
                auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

                yyjson_val* root = yyjson_doc_get_root(doc);
                if (root && yyjson_is_obj(root))
                {
                    yyjson_val* textBuf = yyjson_obj_get(root, "textBufferMiB");
                    if (textBuf && yyjson_is_int(textBuf))
                    {
                        const int64_t value = yyjson_get_int(textBuf);
                        if (value > 0)
                        {
                            textBufferMiB = static_cast<uint32_t>(std::min<int64_t>(value, 256));
                        }
                    }

                    yyjson_val* hexBuf = yyjson_obj_get(root, "hexBufferMiB");
                    if (hexBuf && yyjson_is_int(hexBuf))
                    {
                        const int64_t value = yyjson_get_int(hexBuf);
                        if (value > 0)
                        {
                            hexBufferMiB = static_cast<uint32_t>(std::min<int64_t>(value, 256));
                        }
                    }

                    yyjson_val* lineNums = yyjson_obj_get(root, "showLineNumbers");
                    if (lineNums && yyjson_is_str(lineNums))
                    {
                        const char* value = yyjson_get_str(lineNums);
                        if (value != nullptr)
                        {
                            showLineNumbers = (strcmp(value, "1") == 0) || (strcmp(value, "true") == 0) || (strcmp(value, "on") == 0);
                        }
                    }

                    yyjson_val* wrap = yyjson_obj_get(root, "wrapText");
                    if (wrap && yyjson_is_str(wrap))
                    {
                        const char* value = yyjson_get_str(wrap);
                        if (value != nullptr)
                        {
                            wrapText = (strcmp(value, "1") == 0) || (strcmp(value, "true") == 0) || (strcmp(value, "on") == 0);
                        }
                    }
                }
            }
        }
    }

    _config.textBufferMiB   = textBufferMiB;
    _config.hexBufferMiB    = hexBufferMiB;
    _config.showLineNumbers = showLineNumbers;
    _config.wrapText        = wrapText;
    _wrap                   = wrapText;

    _configurationJson = std::format("{{\"textBufferMiB\":{},\"hexBufferMiB\":{},\"showLineNumbers\":\"{}\",\"wrapText\":\"{}\"}}",
                                     _config.textBufferMiB,
                                     _config.hexBufferMiB,
                                     _config.showLineNumbers ? "1" : "0",
                                     _config.wrapText ? "1" : "0");
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    if (_configurationJson.empty())
    {
        *configurationJsonUtf8 = nullptr;
        return S_OK;
    }

    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    const bool isDefault = _config.textBufferMiB == 16u && _config.hexBufferMiB == 8u && ! _config.showLineNumbers && _config.wrapText;
    *pSomethingToSave    = isDefault ? FALSE : TRUE;
    return S_OK;
}

namespace
{
struct ViewerTextClassBackgroundBrushState
{
    ViewerTextClassBackgroundBrushState()                                                      = default;
    ViewerTextClassBackgroundBrushState(const ViewerTextClassBackgroundBrushState&)            = delete;
    ViewerTextClassBackgroundBrushState& operator=(const ViewerTextClassBackgroundBrushState&) = delete;
    ViewerTextClassBackgroundBrushState(ViewerTextClassBackgroundBrushState&&)                 = delete;
    ViewerTextClassBackgroundBrushState& operator=(ViewerTextClassBackgroundBrushState&&)      = delete;

    wil::unique_hbrush activeBrush;
    COLORREF activeColor = CLR_INVALID;

    wil::unique_hbrush pendingBrush;
    COLORREF pendingColor = CLR_INVALID;

    bool viewerClassRegistered   = false;
    bool textViewClassRegistered = false;
    bool hexViewClassRegistered  = false;
};

ViewerTextClassBackgroundBrushState g_viewerTextClassBackgroundBrush;

HBRUSH GetActiveViewerTextClassBackgroundBrush() noexcept
{
    if (g_viewerTextClassBackgroundBrush.pendingBrush)
    {
        return g_viewerTextClassBackgroundBrush.pendingBrush.get();
    }

    if (! g_viewerTextClassBackgroundBrush.activeBrush)
    {
        const COLORREF fallback                      = GetSysColor(COLOR_WINDOW);
        g_viewerTextClassBackgroundBrush.activeColor = fallback;
        g_viewerTextClassBackgroundBrush.activeBrush.reset(CreateSolidBrush(fallback));
    }

    return g_viewerTextClassBackgroundBrush.activeBrush.get();
}

void RequestViewerTextClassBackgroundColor(COLORREF color) noexcept
{
    if (color == CLR_INVALID)
    {
        return;
    }

    if (g_viewerTextClassBackgroundBrush.pendingBrush && g_viewerTextClassBackgroundBrush.pendingColor == color)
    {
        return;
    }

    wil::unique_hbrush brush(CreateSolidBrush(color));
    if (! brush)
    {
        return;
    }

    g_viewerTextClassBackgroundBrush.pendingColor = color;
    g_viewerTextClassBackgroundBrush.pendingBrush = std::move(brush);
}

void ApplyPendingViewerTextClassBackgroundBrush(HWND viewerHwnd, HWND textViewHwnd, HWND hexViewHwnd) noexcept
{
    if (! g_viewerTextClassBackgroundBrush.pendingBrush)
    {
        return;
    }

    const bool needViewer = g_viewerTextClassBackgroundBrush.viewerClassRegistered;
    const bool needText   = g_viewerTextClassBackgroundBrush.textViewClassRegistered;
    const bool needHex    = g_viewerTextClassBackgroundBrush.hexViewClassRegistered;

    if (! needViewer && ! needText && ! needHex)
    {
        return;
    }

    if ((needViewer && ! viewerHwnd) || (needText && ! textViewHwnd) || (needHex && ! hexViewHwnd))
    {
        return;
    }

    const LONG_PTR newBrush = reinterpret_cast<LONG_PTR>(g_viewerTextClassBackgroundBrush.pendingBrush.get());
    if (needViewer)
    {
        SetClassLongPtrW(viewerHwnd, GCLP_HBRBACKGROUND, newBrush);
    }
    if (needText)
    {
        SetClassLongPtrW(textViewHwnd, GCLP_HBRBACKGROUND, newBrush);
    }
    if (needHex)
    {
        SetClassLongPtrW(hexViewHwnd, GCLP_HBRBACKGROUND, newBrush);
    }

    g_viewerTextClassBackgroundBrush.activeBrush  = std::move(g_viewerTextClassBackgroundBrush.pendingBrush);
    g_viewerTextClassBackgroundBrush.activeColor  = g_viewerTextClassBackgroundBrush.pendingColor;
    g_viewerTextClassBackgroundBrush.pendingColor = CLR_INVALID;
}
} // namespace

ATOM ViewerText::RegisterWndClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        g_viewerTextClassBackgroundBrush.viewerClassRegistered = true;
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = GetActiveViewerTextClassBackgroundBrush();
    wc.lpszClassName = kClassName;
    atom             = RegisterClassExW(&wc);
    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom                                                   = 1;
            g_viewerTextClassBackgroundBrush.viewerClassRegistered = true;
        }
        else
        {
            Debug::ErrorWithLastError(L"ViewerText: RegisterClassExW failed.");
        }
    }
    else
    {
        g_viewerTextClassBackgroundBrush.viewerClassRegistered = true;
    }
    return atom;
}

ATOM ViewerText::RegisterTextViewClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        g_viewerTextClassBackgroundBrush.textViewClassRegistered = true;
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = TextViewProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_IBEAM);
    wc.hbrBackground = GetActiveViewerTextClassBackgroundBrush();
    wc.lpszClassName = kTextViewClassName;
    atom             = RegisterClassExW(&wc);
    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom                                                     = 1;
            g_viewerTextClassBackgroundBrush.textViewClassRegistered = true;
        }
        else
        {
            Debug::ErrorWithLastError(L"ViewerText: RegisterClassExW failed for text view class.");
        }
    }
    else
    {
        g_viewerTextClassBackgroundBrush.textViewClassRegistered = true;
    }
    return atom;
}

ATOM ViewerText::RegisterHexViewClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom)
    {
        g_viewerTextClassBackgroundBrush.hexViewClassRegistered = true;
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = HexViewProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_IBEAM);
    wc.hbrBackground = GetActiveViewerTextClassBackgroundBrush();
    wc.lpszClassName = kHexViewClassName;
    atom             = RegisterClassExW(&wc);
    if (atom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom                                                    = 1;
            g_viewerTextClassBackgroundBrush.hexViewClassRegistered = true;
        }
        else
        {
            Debug::ErrorWithLastError(L"ViewerText: RegisterClassExW failed for hex view class.");
        }
    }
    else
    {
        g_viewerTextClassBackgroundBrush.hexViewClassRegistered = true;
    }
    return atom;
}

LRESULT CALLBACK ViewerText::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        auto* self     = static_cast<ViewerText*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            InitPostedPayloadWindow(hwnd);
        }
    }

    auto* self = reinterpret_cast<ViewerText*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self)
    {
        return self->WndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK ViewerText::TextViewProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        auto* self     = static_cast<ViewerText*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    auto* self = reinterpret_cast<ViewerText*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self)
    {
        return self->TextViewProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK ViewerText::HexViewProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        auto* self     = static_cast<ViewerText*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    auto* self = reinterpret_cast<ViewerText*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self)
    {
        return self->HexViewProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerText::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_DPICHANGED: OnDpiChanged(hwnd, static_cast<UINT>(LOWORD(wp)), reinterpret_cast<const RECT*>(lp)); return 0;
        case WM_COMMAND: OnCommand(hwnd, LOWORD(wp), HIWORD(wp), reinterpret_cast<HWND>(lp)); return 0;
        case WM_NOTIFY: return OnNotify(reinterpret_cast<const NMHDR*>(lp));
        case WM_MEASUREITEM: return OnMeasureItem(hwnd, reinterpret_cast<MEASUREITEMSTRUCT*>(lp));
        case WM_DRAWITEM: return OnDrawItem(reinterpret_cast<DRAWITEMSTRUCT*>(lp));
        case WM_KEYDOWN:
            if (HandleShortcutKey(hwnd, wp))
            {
                return 0;
            }
            break;
        case WM_CTLCOLORLISTBOX:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLORBTN: return OnCtlColor(msg, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_MOUSEMOVE:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            OnMouseMove(pt.x, pt.y);
            return 0;
        }
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_LBUTTONDOWN:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            OnLButtonDown(pt.x, pt.y);
            return 0;
        }
        case WM_LBUTTONUP:
        {
            const POINT pt = {static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))};
            OnLButtonUp(pt.x, pt.y);
            return 0;
        }
        case WM_TIMER: OnTimer(static_cast<UINT_PTR>(wp)); return 0;
        case WM_SETCURSOR:
            if (OnSetCursor(hwnd, lp))
            {
                return TRUE;
            }
            break;
        case kAsyncOpenCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncOpenResult>(lp);
            OnAsyncOpenComplete(std::move(result));
            return 0;
        }
        case WM_PAINT: OnPaint(); return 0;
        case WM_ERASEBKGND: return _allowEraseBkgnd ? DefWindowProcW(hwnd, msg, wp, lp) : 1;
        case WM_CLOSE: CommandExit(hwnd); return 0;
        case WM_NCACTIVATE: OnNcActivate(wp != FALSE); return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_NCDESTROY: return OnNcDestroy(hwnd, wp, lp);
        default: return DefWindowProcW(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void ViewerText::OnNcActivate(bool windowActive) noexcept
{
    ApplyTitleBarTheme(windowActive);
}

LRESULT ViewerText::OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    OnDestroy();
    static_cast<void>(DrainPostedPayloadsForWindow(hwnd));

    _hFileCombo.release();
    _hEdit.release();
    _hHex.release();
    _hWnd.release();
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);

    Release();
    return DefWindowProcW(hwnd, WM_NCDESTROY, wp, lp);
}

void ViewerText::OnCreate(HWND hwnd)
{
    const UINT dpi         = GetDpiForWindow(hwnd);
    const int uiHeightPx   = -MulDiv(9, static_cast<int>(dpi), 72);
    const int monoHeightPx = -MulDiv(10, static_cast<int>(dpi), 72);

    _allowEraseBkgnd         = true;
    _allowEraseBkgndTextView = true;
    _allowEraseBkgndHexView  = true;

    _uiFont.reset(CreateFontW(uiHeightPx,
                              0,
                              0,
                              0,
                              FW_NORMAL,
                              FALSE,
                              FALSE,
                              FALSE,
                              DEFAULT_CHARSET,
                              OUT_DEFAULT_PRECIS,
                              CLIP_DEFAULT_PRECIS,
                              CLEARTYPE_QUALITY,
                              DEFAULT_PITCH | FF_DONTCARE,
                              L"Segoe UI"));
    if (! _uiFont)
    {
        Debug::ErrorWithLastError(L"ViewerText: CreateFontW failed for UI font.");
    }
    _monoFont.reset(CreateFontW(monoHeightPx,
                                0,
                                0,
                                0,
                                FW_NORMAL,
                                FALSE,
                                FALSE,
                                FALSE,
                                DEFAULT_CHARSET,
                                OUT_DEFAULT_PRECIS,
                                CLIP_DEFAULT_PRECIS,
                                CLEARTYPE_QUALITY,
                                FIXED_PITCH | FF_MODERN,
                                L"Consolas"));
    if (! _monoFont)
    {
        Debug::ErrorWithLastError(L"ViewerText: CreateFontW failed for monospace font.");
    }

    const DWORD comboStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS;
    _hFileCombo.reset(CreateWindowExW(
        0, L"COMBOBOX", nullptr, comboStyle, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_VIEWERTEXT_FILE_COMBO)), g_hInstance, nullptr));
    if (! _hFileCombo)
    {
        Debug::ErrorWithLastError(L"ViewerText: CreateWindowExW failed for file combo.");
    }
    if (_hFileCombo && _uiFont)
    {
        SendMessageW(_hFileCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_uiFont.get()), TRUE);
    }
    if (_hFileCombo)
    {
        InstallFileComboEscClose(_hFileCombo.get());
    }
    if (_hFileCombo)
    {
        int itemHeight = PxFromDip(24, dpi);
        auto hdc       = wil::GetDC(hwnd);
        if (hdc)
        {
            HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
            static_cast<void>(oldFont);

            TEXTMETRICW tm{};
            if (GetTextMetricsW(hdc.get(), &tm) != 0)
            {
                itemHeight = tm.tmHeight + tm.tmExternalLeading + PxFromDip(6, dpi);
            }
        }

        itemHeight = std::max(itemHeight, 1);
        SendMessageW(_hFileCombo.get(), CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), static_cast<LPARAM>(itemHeight));
        SendMessageW(_hFileCombo.get(), CB_SETITEMHEIGHT, 0, static_cast<LPARAM>(itemHeight));
    }
    if (_hFileCombo)
    {
        COMBOBOXINFO info{};
        info.cbSize = sizeof(info);
        if (GetComboBoxInfo(_hFileCombo.get(), &info) != 0)
        {
            _hFileComboList = info.hwndList;
            _hFileComboItem = info.hwndItem;
        }
    }

    static_cast<void>(RegisterTextViewClass(g_hInstance));
    const DWORD textStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | WS_HSCROLL;
    _hEdit.reset(CreateWindowExW(0, kTextViewClassName, nullptr, textStyle, 0, 0, 0, 0, hwnd, nullptr, g_hInstance, this));
    if (! _hEdit)
    {
        Debug::ErrorWithLastError(L"ViewerText: CreateWindowExW failed for DirectX text view.");
    }

    static_cast<void>(RegisterHexViewClass(g_hInstance));
    const DWORD hexStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | WS_HSCROLL;
    _hHex.reset(CreateWindowExW(0, kHexViewClassName, nullptr, hexStyle, 0, 0, 0, 0, hwnd, nullptr, g_hInstance, this));
    if (! _hHex)
    {
        Debug::ErrorWithLastError(L"ViewerText: CreateWindowExW failed for DirectX hex view.");
    }

    ApplyTheme(hwnd);
    RefreshFileCombo(hwnd);
    Layout(hwnd);
    SetViewMode(hwnd, _viewMode);
    SetWrap(hwnd, _wrap);
}

void ViewerText::OnDestroy()
{
    EndLoadingUi();
    DiscardDirect2D();
    DiscardTextViewDirect2D();
    DiscardHexViewDirect2D();
    ResetHexState();
    _windowIconSmall.reset();
    _windowIconBig.reset();

    IViewerCallback* callback = _callback;
    void* cookie              = _callbackCookie;
    if (callback)
    {
        AddRef();
        static_cast<void>(callback->ViewerClosed(cookie));
        Release();
    }
}

void ViewerText::StartAsyncOpen(HWND hwnd, const std::filesystem::path& path, bool updateOtherFiles, UINT displayEncodingMenuSelection) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (path.empty())
    {
        Debug::Error(L"ViewerText: StartAsyncOpen called with an empty path.");
        return;
    }

    if (! _fileSystem)
    {
        Debug::Error(L"ViewerText: StartAsyncOpen failed because file system is missing.");
        return;
    }

    const bool pathChanged = (_currentPath != path);
    _currentPath           = path;

    if (updateOtherFiles)
    {
        _otherFiles.clear();
        _otherFiles.push_back(path);
        _otherIndex = 0;
        RefreshFileCombo(hwnd);
    }
    else
    {
        SyncFileComboSelection();
    }

    const std::wstring title = FormatStringResource(g_hInstance, IDS_VIEWERTEXT_TITLE_FORMAT, path.filename().wstring());
    if (! title.empty())
    {
        SetWindowTextW(hwnd, title.c_str());
    }

    _statusMessage.clear();
    _fileReader.reset();
    _fileSize              = 0;
    _encoding              = FileEncoding::Unknown;
    _bomBytes              = 0;
    _textStreamActive      = false;
    _textStreamSkipBytes   = 0;
    _textStreamStartOffset = 0;
    _textStreamEndOffset   = 0;
    _textTotalLineCount.reset();
    _textStreamLineCountedEndOffset = 0;
    _textStreamLineCountedNewlines  = 0;
    _textStreamLineCountLastWasCR   = false;
    _detectedCodePage               = 0;
    _detectedCodePageValid          = false;
    _detectedCodePageIsGuess        = false;

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
    _textMaxLineLength   = 0;

    ResetHexState();

    BeginLoadingUi();
    SetViewMode(hwnd, _viewMode);
    InvalidateRect(hwnd, nullptr, TRUE);

    const uint64_t requestId  = _asyncOpenRequestId.fetch_add(1, std::memory_order_relaxed) + 1u;
    _activeAsyncOpenRequestId = requestId;

    const ViewMode desiredViewMode              = _viewMode;
    const UINT previousDisplayEncodingSelection = _displayEncodingMenuSelection;
    const uint32_t textBufferMiB                = _config.textBufferMiB;
    const uint32_t hexBufferMiB                 = _config.hexBufferMiB;
    const bool allowHexFallback                 = static_cast<bool>(_hHex);

    wil::com_ptr<IFileSystem> fileSystem = _fileSystem;

    AddRef();

    struct AsyncOpenWorkItem final
    {
        AsyncOpenWorkItem()                                    = default;
        AsyncOpenWorkItem(const AsyncOpenWorkItem&)            = delete;
        AsyncOpenWorkItem& operator=(const AsyncOpenWorkItem&) = delete;

        wil::unique_hmodule moduleKeepAlive;
        std::function<void()> work;
    };

    auto ctx = std::unique_ptr<AsyncOpenWorkItem>(new (std::nothrow) AsyncOpenWorkItem{});

    ctx->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kViewerTextModuleAnchor);
    ctx->work            = [this,
                 hwnd,
                 requestId,
                 fileSystem = std::move(fileSystem),
                 path,
                 pathChanged,
                 desiredViewMode,
                 updateOtherFiles,
                 displayEncodingMenuSelection,
                 previousDisplayEncodingSelection,
                 textBufferMiB,
                 hexBufferMiB,
                 allowHexFallback]() mutable
    {
        auto releaseSelf = wil::scope_exit([&] { Release(); });

        // Sleep(15000); // Simulate long operation for testing purposes.
        std::unique_ptr<AsyncOpenResult> result(new (std::nothrow) AsyncOpenResult{});
        if (! result)
        {
            return;
        }

        result->viewer           = this;
        result->requestId        = requestId;
        result->path             = path;
        result->updateOtherFiles = updateOtherFiles;
        result->viewMode         = desiredViewMode;
        result->hr               = E_FAIL;

        wil::com_ptr<IFileSystemIO> fileIo;
        const HRESULT fileIoHr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
        if (FAILED(fileIoHr) || ! fileIo)
        {
            Debug::Error(L"ViewerText: Active filesystem does not implement IFileSystemIO (hr=0x{:08X}).", static_cast<unsigned long>(fileIoHr));
            result->hr = FAILED(fileIoHr) ? fileIoHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            return;
        }

        const HRESULT openReaderHr = fileIo->CreateFileReader(path.c_str(), result->fileReader.put());
        if (FAILED(openReaderHr) || ! result->fileReader)
        {
            Debug::Error(L"ViewerText: Failed to create file reader for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(openReaderHr));
            result->hr = FAILED(openReaderHr) ? openReaderHr : E_FAIL;
            return;
        }

        FileEncoding encoding = FileEncoding::Unknown;
        uint64_t bomBytes     = 0;
        uint64_t fileSize     = 0;

        uint64_t sizeBytes   = 0;
        const HRESULT sizeHr = result->fileReader->GetSize(&sizeBytes);
        if (FAILED(sizeHr))
        {
            Debug::Error(L"ViewerText: GetSize failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(sizeHr));
            result->hr = sizeHr;
            return;
        }

        fileSize = sizeBytes;

        BYTE bom[4]{};
        unsigned long read = 0;

        uint64_t pos         = 0;
        const HRESULT seekHr = result->fileReader->Seek(0, FILE_BEGIN, &pos);
        if (FAILED(seekHr))
        {
            Debug::Error(L"ViewerText: Seek(FILE_BEGIN, 0) failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(seekHr));
            result->hr = seekHr;
            return;
        }

        const HRESULT readHr = result->fileReader->Read(bom, static_cast<unsigned long>(std::size(bom)), &read);
        if (FAILED(readHr))
        {
            Debug::Error(L"ViewerText: Read failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(readHr));
            result->hr = readHr;
            return;
        }

        if (read >= 4 && bom[0] == 0xFF && bom[1] == 0xFE && bom[2] == 0x00 && bom[3] == 0x00)
        {
            encoding = FileEncoding::Utf32LE;
            bomBytes = 4;
        }
        else if (read >= 4 && bom[0] == 0x00 && bom[1] == 0x00 && bom[2] == 0xFE && bom[3] == 0xFF)
        {
            encoding = FileEncoding::Utf32BE;
            bomBytes = 4;
        }
        else if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
        {
            encoding = FileEncoding::Utf8;
            bomBytes = 3;
        }
        else if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
        {
            encoding = FileEncoding::Utf16LE;
            bomBytes = 2;
        }
        else if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
        {
            encoding = FileEncoding::Utf16BE;
            bomBytes = 2;
        }

        result->encoding = encoding;
        result->bomBytes = bomBytes;
        result->fileSize = fileSize;

        UINT selection = previousDisplayEncodingSelection;
        if (displayEncodingMenuSelection != 0 && IsEncodingMenuSelectionValid(displayEncodingMenuSelection))
        {
            selection = displayEncodingMenuSelection;
        }
        else if (pathChanged)
        {
            selection = IDM_VIEWER_ENCODING_DISPLAY_ANSI;
            switch (encoding)
            {
                case FileEncoding::Utf8: selection = IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM; break;
                case FileEncoding::Utf16LE: selection = IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM; break;
                case FileEncoding::Utf16BE: selection = IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM; break;
                case FileEncoding::Utf32LE: selection = IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM; break;
                case FileEncoding::Utf32BE: selection = IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM; break;
                case FileEncoding::Unknown:
                default:
                {
                    constexpr unsigned long kProbeSize = 64u * 1024u;
                    std::array<uint8_t, kProbeSize> probe{};
                    unsigned long probeRead = 0;

                    uint64_t probePos = 0;
                    if (SUCCEEDED(result->fileReader->Seek(0, FILE_BEGIN, &probePos)))
                    {
                        static_cast<void>(result->fileReader->Read(probe.data(), static_cast<unsigned long>(probe.size()), &probeRead));
                        if (probeRead != 0 && IsValidUtf8(probe.data(), static_cast<size_t>(probeRead)))
                        {
                            selection = IDM_VIEWER_ENCODING_DISPLAY_UTF8;
                        }
                    }
                    break;
                }
            }
        }

        if (! IsEncodingMenuSelectionValid(selection))
        {
            selection = IDM_VIEWER_ENCODING_DISPLAY_ANSI;
        }

        result->displayEncodingMenuSelection = selection;

        const uint64_t streamSkipBytes = ::BytesToSkipForDisplayEncoding(selection, encoding, bomBytes);
        result->textStreamSkipBytes    = streamSkipBytes;

        const uint64_t clampedStart   = std::min<uint64_t>(streamSkipBytes, fileSize);
        result->textStreamStartOffset = clampedStart;
        result->textStreamEndOffset   = clampedStart;
        result->textStreamActive      = false;

        const FileEncoding displayEncoding = DisplayEncodingFileEncodingForSelection(selection);
        const UINT displayCodePage         = CodePageForSelection(selection);
        const uint64_t maxChunkBytes       = ::TextStreamChunkBytes(textBufferMiB, displayEncoding);

        const uint64_t availableBytes = (fileSize > clampedStart) ? (fileSize - clampedStart) : 0;
        const uint64_t wantBytes64    = std::min<uint64_t>(availableBytes, maxChunkBytes);
        const size_t wantBytes        = static_cast<size_t>(std::min<uint64_t>(wantBytes64, static_cast<uint64_t>(std::numeric_limits<size_t>::max())));

        if (clampedStart > static_cast<uint64_t>(std::numeric_limits<__int64>::max()))
        {
            Debug::Error(L"ViewerText: File is too large to open (start offset 0x{:016X} exceeds maximum supported offset).", clampedStart);
            result->hr = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            return;
        }

        uint64_t newPosition     = 0;
        const HRESULT seekDataHr = result->fileReader->Seek(static_cast<__int64>(clampedStart), FILE_BEGIN, &newPosition);
        if (FAILED(seekDataHr))
        {
            Debug::Error(L"ViewerText: Seek to data start offset failed (0x{:016X}) for '{}' (hr=0x{:08X}).",
                         clampedStart,
                         path.c_str(),
                         static_cast<unsigned long>(seekDataHr));
            result->hr = seekDataHr;
            return;
        }

        std::vector<uint8_t> bytes(wantBytes);
        size_t bytesReadTotal = 0;
        while (bytesReadTotal < bytes.size())
        {
            const size_t remaining   = bytes.size() - bytesReadTotal;
            const unsigned long want = remaining > static_cast<size_t>(std::numeric_limits<unsigned long>::max()) ? std::numeric_limits<unsigned long>::max()
                                                                                                                  : static_cast<unsigned long>(remaining);

            unsigned long chunkRead = 0;
            const HRESULT chunkHr   = result->fileReader->Read(bytes.data() + bytesReadTotal, want, &chunkRead);
            if (FAILED(chunkHr))
            {
                Debug::Error(L"ViewerText: Read failed for '{}' at offset 0x{:016X} (hr=0x{:08X}).",
                             path.c_str(),
                             clampedStart + bytesReadTotal,
                             static_cast<unsigned long>(chunkHr));
                result->hr = chunkHr;
                return;
            }

            if (chunkRead == 0)
            {
                break;
            }

            bytesReadTotal += static_cast<size_t>(chunkRead);
        }

        bytes.resize(bytesReadTotal);

        ViewMode targetViewMode = desiredViewMode;
        if (targetViewMode == ViewMode::Text && allowHexFallback)
        {
            const bool unicodeDecode = (displayEncoding == FileEncoding::Utf16LE || displayEncoding == FileEncoding::Utf16BE ||
                                        displayEncoding == FileEncoding::Utf32LE || displayEncoding == FileEncoding::Utf32BE);
            if (! unicodeDecode && LooksLikeBinaryData(bytes.data(), bytes.size()))
            {
                targetViewMode = ViewMode::Hex;
            }
        }

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
            carryBytes = Utf8IncompleteTailSize(bytes.data(), bytes.size());
        }

        carryBytes                = std::min(carryBytes, bytes.size());
        const size_t convertBytes = bytes.size() - carryBytes;

        result->textBuffer.clear();
        if (convertBytes > 0)
        {
            if ((displayEncoding == FileEncoding::Utf16LE || displayEncoding == FileEncoding::Utf16BE) && (convertBytes % 2) == 0)
            {
                const size_t wcharCount = convertBytes / 2;
                result->textBuffer.resize(wcharCount);
                memcpy(result->textBuffer.data(), bytes.data(), convertBytes);

                if (displayEncoding == FileEncoding::Utf16BE)
                {
                    for (size_t i = 0; i < result->textBuffer.size(); ++i)
                    {
                        const wchar_t v       = result->textBuffer[i];
                        result->textBuffer[i] = static_cast<wchar_t>((static_cast<uint16_t>(v) >> 8) | (static_cast<uint16_t>(v) << 8));
                    }
                }
            }
            else if ((displayEncoding == FileEncoding::Utf32LE || displayEncoding == FileEncoding::Utf32BE) && (convertBytes % 4) == 0)
            {
                const bool bigEndian = (displayEncoding == FileEncoding::Utf32BE);
                result->textBuffer.reserve(convertBytes / 4);

                for (size_t i = 0; i + 3 < convertBytes; i += 4)
                {
                    uint32_t cp = 0;
                    if (bigEndian)
                    {
                        cp = (static_cast<uint32_t>(bytes[i]) << 24) | (static_cast<uint32_t>(bytes[i + 1]) << 16) |
                             (static_cast<uint32_t>(bytes[i + 2]) << 8) | static_cast<uint32_t>(bytes[i + 3]);
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
                            result->textBuffer.push_back(static_cast<wchar_t>(0xFFFDu));
                        }
                        else
                        {
                            result->textBuffer.push_back(static_cast<wchar_t>(cp));
                        }
                    }
                    else if (cp <= 0x10FFFFu)
                    {
                        const uint32_t v = cp - 0x10000u;
                        result->textBuffer.push_back(static_cast<wchar_t>(0xD800u + (v >> 10)));
                        result->textBuffer.push_back(static_cast<wchar_t>(0xDC00u + (v & 0x3FFu)));
                    }
                    else
                    {
                        result->textBuffer.push_back(static_cast<wchar_t>(0xFFFDu));
                    }
                }
            }
            else
            {
                if (convertBytes > static_cast<size_t>(std::numeric_limits<int>::max()))
                {
                    Debug::Error(L"ViewerText: File is too large to open (data size 0x{:016X} exceeds maximum supported size).", convertBytes);
                    result->hr = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                    return;
                }

                const int srcLen       = static_cast<int>(convertBytes);
                const int requiredWide = MultiByteToWideChar(displayCodePage, 0, reinterpret_cast<LPCCH>(bytes.data()), srcLen, nullptr, 0);
                if (requiredWide <= 0)
                {
                    auto lastError = Debug::ErrorWithLastError(
                        L"ViewerText: MultiByteToWideChar failed to calculate required buffer size for '{}' (hr=0x{:08X}).", path.c_str(), result->hr);
                    result->hr = HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_INVALID_DATA);
                    return;
                }

                result->textBuffer.resize(static_cast<size_t>(requiredWide));
                const int written =
                    MultiByteToWideChar(displayCodePage, 0, reinterpret_cast<LPCCH>(bytes.data()), srcLen, result->textBuffer.data(), requiredWide);
                if (written <= 0)
                {
                    auto lastError =
                        Debug::ErrorWithLastError(L"ViewerText: MultiByteToWideChar failed to convert data for '{}' (hr=0x{:08X}).", path.c_str(), result->hr);
                    result->hr = HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_INVALID_DATA);
                    return;
                }

                result->textBuffer.resize(static_cast<size_t>(written));
            }
        }

        result->textStreamStartOffset = clampedStart;
        if (bytesReadTotal >= carryBytes)
        {
            const uint64_t consumed     = static_cast<uint64_t>(bytesReadTotal - carryBytes);
            result->textStreamEndOffset = std::min<uint64_t>(clampedStart + consumed, fileSize);
        }
        else
        {
            result->textStreamEndOffset = clampedStart;
        }

        result->textStreamActive = (fileSize > streamSkipBytes) && ((fileSize - streamSkipBytes) > maxChunkBytes);

        const UINT defaultCodePage      = GetACP();
        result->detectedCodePage        = 0;
        result->detectedCodePageValid   = false;
        result->detectedCodePageIsGuess = false;

        switch (encoding)
        {
            case FileEncoding::Utf8:
                result->detectedCodePage        = CP_UTF8;
                result->detectedCodePageValid   = true;
                result->detectedCodePageIsGuess = false;
                break;
            case FileEncoding::Utf16LE:
                result->detectedCodePage        = 1200u;
                result->detectedCodePageValid   = true;
                result->detectedCodePageIsGuess = false;
                break;
            case FileEncoding::Utf16BE:
                result->detectedCodePage        = 1201u;
                result->detectedCodePageValid   = true;
                result->detectedCodePageIsGuess = false;
                break;
            case FileEncoding::Utf32LE:
                result->detectedCodePage        = 12000u;
                result->detectedCodePageValid   = true;
                result->detectedCodePageIsGuess = false;
                break;
            case FileEncoding::Utf32BE:
                result->detectedCodePage        = 12001u;
                result->detectedCodePageValid   = true;
                result->detectedCodePageIsGuess = false;
                break;
            case FileEncoding::Unknown:
            default:
            {
                result->detectedCodePageIsGuess = true;
                if (! bytes.empty() && IsValidUtf8(bytes.data(), bytes.size()))
                {
                    result->detectedCodePage = CP_UTF8;
                }
                else
                {
                    result->detectedCodePage = defaultCodePage;
                }
                result->detectedCodePageValid = true;
                break;
            }
        }

        BuildTextLineIndex(result->textBuffer, result->textLineStarts, result->textLineEnds, result->textMaxLineLength);

        result->viewMode = targetViewMode;

        const bool needHex = (targetViewMode == ViewMode::Hex);
        if (needHex && fileSize > 0)
        {
            if (fileSize <= kMaxHexLoadBytes && fileSize <= static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
            {
                result->hexBytes.resize(static_cast<size_t>(fileSize));

                uint64_t ignored        = 0;
                const HRESULT seekHexHr = result->fileReader->Seek(0, FILE_BEGIN, &ignored);
                if (FAILED(seekHexHr))
                {
                    Debug::Warning(
                        L"ViewerText: Seek(FILE_BEGIN, 0) failed for HEX preload of '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(seekHexHr));
                    result->hexBytes.clear();
                }
                else
                {
                    size_t offset = 0;
                    while (offset < result->hexBytes.size())
                    {
                        const unsigned long want = static_cast<unsigned long>(std::min<size_t>(256 * 1024, result->hexBytes.size() - offset));
                        unsigned long readHex    = 0;
                        const HRESULT readHexHr  = result->fileReader->Read(result->hexBytes.data() + offset, want, &readHex);
                        if (FAILED(readHexHr))
                        {
                            Debug::Warning(
                                L"ViewerText: Read failed for HEX preload of '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(readHexHr));
                            result->hexBytes.clear();
                            break;
                        }
                        if (readHex == 0)
                        {
                            break;
                        }
                        offset += static_cast<size_t>(readHex);
                    }
                }
            }
            else
            {
                uint64_t cacheBytes = static_cast<uint64_t>(hexBufferMiB) * 1024u * 1024u;
                cacheBytes          = std::clamp<uint64_t>(cacheBytes, 256u * 1024u, 256u * 1024u * 1024u);

                const uint64_t remaining = fileSize;
                const uint64_t want64    = std::min<uint64_t>(remaining, cacheBytes);
                const unsigned long want = want64 > static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()) ? std::numeric_limits<unsigned long>::max()
                                                                                                                     : static_cast<unsigned long>(want64);

                result->hexCacheOffset = 0;
                result->hexCacheValid  = 0;
                if (want > 0)
                {
                    result->hexCache.resize(static_cast<size_t>(want));

                    uint64_t ignored        = 0;
                    const HRESULT seekHexHr = result->fileReader->Seek(0, FILE_BEGIN, &ignored);
                    if (FAILED(seekHexHr))
                    {
                        Debug::Warning(L"ViewerText: Seek(FILE_BEGIN, 0) failed for HEX cache preload of '{}' (hr=0x{:08X}).",
                                       path.c_str(),
                                       static_cast<unsigned long>(seekHexHr));
                        result->hexCache.clear();
                    }
                    else
                    {
                        unsigned long readHex   = 0;
                        const HRESULT readHexHr = result->fileReader->Read(result->hexCache.data(), want, &readHex);
                        if (FAILED(readHexHr))
                        {
                            Debug::Warning(
                                L"ViewerText: Read failed for HEX cache preload of '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(readHexHr));
                            result->hexCache.clear();
                        }
                        else
                        {
                            result->hasHexCache   = true;
                            result->hexCacheValid = static_cast<size_t>(readHex);
                        }
                    }
                }
            }
        }

        result->hr = S_OK;

        if (FAILED(result->hr) && result->hr != E_OUTOFMEMORY && allowHexFallback && result->fileReader && result->fileSize > 0)
        {
            result->hexBytes.clear();
            result->hexCache.clear();
            result->hexCacheOffset = 0;
            result->hexCacheValid  = 0;
            result->hasHexCache    = false;

            const uint64_t fileSize = result->fileSize;
            if (fileSize <= kMaxHexLoadBytes && fileSize <= static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
            {
                result->hexBytes.resize(static_cast<size_t>(fileSize));

                uint64_t ignored        = 0;
                const HRESULT seekHexHr = result->fileReader->Seek(0, FILE_BEGIN, &ignored);
                if (FAILED(seekHexHr))
                {
                    result->hexBytes.clear();
                    Debug::Error(
                        L"ViewerText: Seek(FILE_BEGIN, 0) failed for HEX fallback of '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(seekHexHr));
                    return;
                }

                size_t offset = 0;
                while (offset < result->hexBytes.size())
                {
                    const unsigned long want = static_cast<unsigned long>(std::min<size_t>(256 * 1024, result->hexBytes.size() - offset));
                    unsigned long readHex    = 0;
                    const HRESULT readHexHr  = result->fileReader->Read(result->hexBytes.data() + offset, want, &readHex);
                    if (FAILED(readHexHr))
                    {
                        result->hexBytes.clear();
                        Debug::Error(L"ViewerText: Read failed for HEX fallback of '{}' at offset 0x{:016X} (hr=0x{:08X}).",
                                     path.c_str(),
                                     offset,
                                     static_cast<unsigned long>(readHexHr));
                        return;
                    }
                    if (readHex == 0)
                    {
                        break;
                    }
                    offset += static_cast<size_t>(readHex);
                }
            }
            else
            {
                uint64_t cacheBytes = static_cast<uint64_t>(hexBufferMiB) * 1024u * 1024u;
                cacheBytes          = std::clamp<uint64_t>(cacheBytes, 256u * 1024u, 256u * 1024u * 1024u);

                const uint64_t aligned   = 0;
                const uint64_t remaining = (fileSize > aligned) ? (fileSize - aligned) : 0;
                const uint64_t want64    = std::min<uint64_t>(remaining, cacheBytes);
                const unsigned long want = want64 > static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()) ? std::numeric_limits<unsigned long>::max()
                                                                                                                     : static_cast<unsigned long>(want64);

                result->hexCacheOffset = aligned;
                if (want > 0)
                {
                    result->hexCache.resize(static_cast<size_t>(want));

                    uint64_t ignored        = 0;
                    const HRESULT seekHexHr = result->fileReader->Seek(static_cast<__int64>(aligned), FILE_BEGIN, &ignored);
                    if (FAILED(seekHexHr))
                    {
                        Debug::Error(L"ViewerText: Seek to offset 0x{:016X} failed for HEX cache fallback of '{}' (hr=0x{:08X}).",
                                     aligned,
                                     path.c_str(),
                                     static_cast<unsigned long>(seekHexHr));
                        return;
                    }

                    unsigned long readHex   = 0;
                    const HRESULT readHexHr = result->fileReader->Read(result->hexCache.data(), want, &readHex);
                    if (FAILED(readHexHr))
                    {
                        Debug::Error(L"ViewerText: Read failed for HEX cache fallback of '{}' at offset 0x{:016X} (hr=0x{:08X}).",
                                     path.c_str(),
                                     aligned,
                                     static_cast<unsigned long>(readHexHr));
                        return;
                    }

                    result->hasHexCache   = true;
                    result->hexCacheValid = static_cast<size_t>(readHex);
                }
            }

            Debug::Warning(
                L"ViewerText: Failed to load '{}' as text (hr=0x{:08X}); falling back to HEX view.", path.c_str(), static_cast<unsigned long>(result->hr));
            result->viewMode = ViewMode::Hex;
            result->hr       = S_OK;
        }

        if (! hwnd || GetWindowLongPtrW(hwnd, GWLP_USERDATA) != reinterpret_cast<LONG_PTR>(this))
        {
            return;
        }

        static_cast<void>(PostMessagePayload(hwnd, kAsyncOpenCompleteMessage, 0, std::move(result)));
    };

    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
        {
            std::unique_ptr<AsyncOpenWorkItem> ctx(static_cast<AsyncOpenWorkItem*>(context));
            if (! ctx)
            {
                return;
            }

            static_cast<void>(ctx->moduleKeepAlive);
            if (ctx->work)
            {
                ctx->work();
            }
        },
        ctx.get(),
        nullptr);

    if (queued == 0)
    {
        Debug::Error(L"ViewerText: Failed to queue async open work item for '{}'.", path.c_str());
        return;
    }

    ctx.release();
}

void ViewerText::OnAsyncOpenComplete(std::unique_ptr<AsyncOpenResult> result) noexcept
{
    if (! result)
    {
        return;
    }
    if (result->viewer != this)
    {
        return;
    }

    if (result->requestId != _activeAsyncOpenRequestId)
    {
        return;
    }

    EndLoadingUi();

    if (FAILED(result->hr))
    {
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_ERR_OPEN_FAILED);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, TRUE);
        }

        ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_OPEN_FAILED);
        return;
    }

    _fileReader                   = std::move(result->fileReader);
    _fileSize                     = result->fileSize;
    _encoding                     = result->encoding;
    _bomBytes                     = result->bomBytes;
    _displayEncodingMenuSelection = result->displayEncodingMenuSelection;
    _detectedCodePage             = result->detectedCodePage;
    _detectedCodePageValid        = result->detectedCodePageValid;
    _detectedCodePageIsGuess      = result->detectedCodePageIsGuess;

    _statusMessage = result->statusMessage;

    _textStreamSkipBytes   = result->textStreamSkipBytes;
    _textStreamStartOffset = result->textStreamStartOffset;
    _textStreamEndOffset   = result->textStreamEndOffset;
    _textStreamActive      = result->textStreamActive;

    _textBuffer        = std::move(result->textBuffer);
    _textLineStarts    = std::move(result->textLineStarts);
    _textLineEnds      = std::move(result->textLineEnds);
    _textMaxLineLength = result->textMaxLineLength;

    _textTotalLineCount.reset();
    _textStreamLineCountedEndOffset = _textStreamStartOffset;
    _textStreamLineCountedNewlines  = 0;
    _textStreamLineCountLastWasCR   = false;
    UpdateTextStreamTotalLineCountAfterLoad();

    _textVisualLineStarts.clear();
    _textVisualLineLogical.clear();
    _textTopVisualLine   = 0;
    _textLeftColumn      = 0;
    _textCaretIndex      = 0;
    _textSelAnchor       = 0;
    _textSelActive       = 0;
    _textPreferredColumn = 0;
    _textSelecting       = false;
    _searchMatchStarts.clear();

    if (_hEdit)
    {
        RebuildTextVisualLines(_hEdit.get());
        UpdateTextViewScrollBars(_hEdit.get());
        UpdateSearchHighlights();
        InvalidateRect(_hEdit.get(), nullptr, TRUE);
    }

    _hexBytes = std::move(result->hexBytes);
    if (result->hasHexCache)
    {
        _hexCache       = std::move(result->hexCache);
        _hexCacheOffset = result->hexCacheOffset;
        _hexCacheValid  = result->hexCacheValid;
    }

    if (_hHex)
    {
        UpdateHexViewScrollBars(_hHex.get());
        InvalidateRect(_hHex.get(), nullptr, TRUE);
    }

    if (_hWnd)
    {
        SetViewMode(_hWnd.get(), result->viewMode);
    }
}

void ViewerText::BeginLoadingUi() noexcept
{
    ClearInlineAlert();

    _isLoading                = true;
    _showLoadingOverlay       = false;
    _loadingSpinnerAngleDeg   = 0.0f;
    _loadingSpinnerLastTickMs = GetTickCount64();

    _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_LOADING);

    if (! _hWnd)
    {
        return;
    }

    KillTimer(_hWnd.get(), kLoadingDelayTimerId);
    KillTimer(_hWnd.get(), kLoadingAnimTimerId);
    SetTimer(_hWnd.get(), kLoadingDelayTimerId, kLoadingDelayMs, nullptr);
}

void ViewerText::EndLoadingUi() noexcept
{
    if (_hWnd)
    {
        KillTimer(_hWnd.get(), kLoadingDelayTimerId);
        KillTimer(_hWnd.get(), kLoadingAnimTimerId);
    }

    _isLoading          = false;
    _showLoadingOverlay = false;
}

void ViewerText::UpdateLoadingSpinner() noexcept
{
    if (! _isLoading || ! _showLoadingOverlay)
    {
        return;
    }

    const ULONGLONG now       = GetTickCount64();
    const ULONGLONG last      = _loadingSpinnerLastTickMs;
    _loadingSpinnerLastTickMs = now;

    double deltaSec = 0.0;
    if (now > last)
    {
        deltaSec = static_cast<double>(now - last) / 1000.0;
    }

    _loadingSpinnerAngleDeg += static_cast<float>(deltaSec * static_cast<double>(kLoadingSpinnerDegPerSec));
    while (_loadingSpinnerAngleDeg >= 360.0f)
    {
        _loadingSpinnerAngleDeg -= 360.0f;
    }

    if (_viewMode == ViewMode::Text && _hEdit)
    {
        InvalidateRect(_hEdit.get(), nullptr, FALSE);
    }
    else if (_viewMode == ViewMode::Hex && _hHex)
    {
        InvalidateRect(_hHex.get(), nullptr, FALSE);
    }
}

void ViewerText::DrawLoadingOverlay(ID2D1HwndRenderTarget* target, ID2D1SolidColorBrush* brush, float widthDip, float heightDip) noexcept
{
    if (! _isLoading || ! _showLoadingOverlay || ! target || ! brush)
    {
        return;
    }

    if (widthDip <= 0.0f || heightDip <= 0.0f)
    {
        return;
    }

    const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

    const std::wstring seed = _currentPath.empty() ? std::wstring(L"viewer") : _currentPath.filename().wstring();
    const COLORREF accent   = _hasTheme ? ResolveAccentColor(_theme, seed) : RGB(0, 120, 215);

    if (! (_hasTheme && _theme.highContrast))
    {
        const uint8_t tintAlpha = (_hasTheme && _theme.darkMode) ? 28u : 18u;
        const COLORREF tint     = BlendColor(bg, accent, tintAlpha);
        const float overlayA    = (_hasTheme && _theme.darkMode) ? 0.85f : 0.75f;
        brush->SetColor(ColorFFromColorRef(tint, overlayA));
        target->FillRectangle(D2D1::RectF(0.0f, 0.0f, widthDip, heightDip), brush);
    }

    const float minDim = std::min(widthDip, heightDip);
    const float radius = std::clamp(minDim * 0.08f, 18.0f, 44.0f);
    const float stroke = std::clamp(radius * 0.20f, 3.0f, 6.0f);
    const float innerR = radius * 0.55f;
    const float outerR = radius;

    const float textHeightDip  = 34.0f;
    const float spacingDip     = 14.0f;
    const float groupHeightDip = outerR * 2.0f + spacingDip + textHeightDip;
    const float groupTopDip    = std::max(0.0f, (heightDip - groupHeightDip) * 0.5f);

    const float cx = widthDip * 0.5f;
    const float cy = groupTopDip + outerR;

    constexpr int kSegments = 12;
    constexpr float kPi     = 3.14159265358979323846f;
    const float baseRad     = (_loadingSpinnerAngleDeg - 90.0f) * (kPi / 180.0f);

    const bool rainbowSpinner = _hasTheme && ! _theme.highContrast && _theme.rainbowMode;
    float rainbowHue          = 0.0f;
    float rainbowSat          = 0.0f;
    float rainbowVal          = 0.0f;
    if (rainbowSpinner)
    {
        const uint32_t h = StableHash32(seed);
        rainbowHue       = static_cast<float>(h % 360u);
        rainbowSat       = _theme.darkBase ? 0.70f : 0.55f;
        rainbowVal       = _theme.darkBase ? 0.95f : 0.85f;
    }

    for (int i = 0; i < kSegments; ++i)
    {
        const float t     = static_cast<float>(i) / static_cast<float>(kSegments);
        const float alpha = 0.15f + 0.85f * (1.0f - t);
        const float angle = baseRad + t * (2.0f * kPi);
        const float s     = std::sin(angle);
        const float c     = std::cos(angle);

        const D2D1_POINT_2F p1 = D2D1::Point2F(cx + c * innerR, cy + s * innerR);
        const D2D1_POINT_2F p2 = D2D1::Point2F(cx + c * outerR, cy + s * outerR);

        COLORREF segmentColor = accent;
        if (rainbowSpinner)
        {
            const float hueStep    = 360.0f / static_cast<float>(kSegments);
            const float hueDegrees = rainbowHue + static_cast<float>(i) * hueStep;
            segmentColor           = ColorFromHSV(hueDegrees, rainbowSat, rainbowVal);
        }

        brush->SetColor(ColorFFromColorRef(segmentColor, alpha));
        target->DrawLine(p1, p2, brush, stroke);
    }

    std::wstring loadingText = _statusMessage;
    if (loadingText.empty())
    {
        loadingText = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_LOADING);
    }

    if (loadingText.empty())
    {
        return;
    }

    if (! _loadingOverlayFormat && _dwriteFactory)
    {
        wil::com_ptr<IDWriteTextFormat> format;
        const HRESULT hr = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 22.0f, L"", format.put());
        if (SUCCEEDED(hr) && format)
        {
            static_cast<void>(format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER));
            static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
            _loadingOverlayFormat = std::move(format);
        }
    }

    if (! _loadingOverlayFormat)
    {
        return;
    }

    const float textTopDip   = groupTopDip + outerR * 2.0f + spacingDip;
    const D2D1_RECT_F textRc = D2D1::RectF(0.0f, textTopDip, widthDip, std::min(heightDip, textTopDip + textHeightDip));

    brush->SetColor(ColorFFromColorRef(fg, 0.90f));
    const UINT32 len = static_cast<UINT32>(std::min<size_t>(loadingText.size(), std::numeric_limits<UINT32>::max()));
    target->DrawTextW(loadingText.c_str(), len, _loadingOverlayFormat.get(), textRc, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
}

void ViewerText::ShowInlineAlert(InlineAlertSeverity severity, UINT titleId, UINT messageId) noexcept
{
    if (! _hostAlerts)
    {
        return;
    }

    const std::wstring title   = LoadStringResource(g_hInstance, titleId);
    const std::wstring message = LoadStringResource(g_hInstance, messageId);
    if (message.empty())
    {
        return;
    }

    HostAlertSeverity hostSeverity = HOST_ALERT_ERROR;
    switch (severity)
    {
        case InlineAlertSeverity::Warning: hostSeverity = HOST_ALERT_WARNING; break;
        case InlineAlertSeverity::Info: hostSeverity = HOST_ALERT_INFO; break;
        case InlineAlertSeverity::Error:
        default: hostSeverity = HOST_ALERT_ERROR; break;
    }

    HWND targetWindow = nullptr;
    if (_viewMode == ViewMode::Hex && _hHex)
    {
        targetWindow = _hHex.get();
    }
    else if (_hEdit)
    {
        targetWindow = _hEdit.get();
    }
    else if (_hWnd)
    {
        targetWindow = _hWnd.get();
    }

    if (! targetWindow)
    {
        return;
    }

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODAL;
    request.severity     = hostSeverity;
    request.targetWindow = targetWindow;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(_hostAlerts->ShowAlert(&request, nullptr));
}

void ViewerText::ClearInlineAlert() noexcept
{
    if (! _hostAlerts)
    {
        return;
    }

    if (_hEdit)
    {
        static_cast<void>(_hostAlerts->ClearAlert(HOST_ALERT_SCOPE_WINDOW, reinterpret_cast<void*>(_hEdit.get())));
    }
    if (_hHex)
    {
        static_cast<void>(_hostAlerts->ClearAlert(HOST_ALERT_SCOPE_WINDOW, reinterpret_cast<void*>(_hHex.get())));
    }
    if (_hWnd)
    {
        static_cast<void>(_hostAlerts->ClearAlert(HOST_ALERT_SCOPE_WINDOW, reinterpret_cast<void*>(_hWnd.get())));
    }
}

void ViewerText::OnTimer(UINT_PTR timerId) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    if (timerId == kLoadingDelayTimerId)
    {
        KillTimer(_hWnd.get(), kLoadingDelayTimerId);
        if (! _isLoading)
        {
            return;
        }

        _showLoadingOverlay       = true;
        _loadingSpinnerAngleDeg   = 0.0f;
        _loadingSpinnerLastTickMs = GetTickCount64();
        SetTimer(_hWnd.get(), kLoadingAnimTimerId, kLoadingAnimIntervalMs, nullptr);

        if (_hEdit)
        {
            InvalidateRect(_hEdit.get(), nullptr, FALSE);
        }
        if (_hHex)
        {
            InvalidateRect(_hHex.get(), nullptr, FALSE);
        }
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
        return;
    }

    if (timerId == kLoadingAnimTimerId)
    {
        UpdateLoadingSpinner();
        return;
    }
}

void ViewerText::OnSize(UINT width, UINT height)
{
    if (! _hWnd)
    {
        return;
    }

    if (_d2dTarget && width > 0 && height > 0)
    {
        const HRESULT hr = _d2dTarget->Resize(D2D1::SizeU(width, height));
        if (FAILED(hr))
        {
            DiscardDirect2D();
        }
    }

    Layout(_hWnd.get());
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void ViewerText::OnDpiChanged(HWND hwnd, UINT newDpi, const RECT* suggested) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (suggested)
    {
        const int width  = std::max(1L, suggested->right - suggested->left);
        const int height = std::max(1L, suggested->bottom - suggested->top);
        SetWindowPos(hwnd, nullptr, suggested->left, suggested->top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    const int uiHeightPx   = -MulDiv(9, static_cast<int>(newDpi), 72);
    const int monoHeightPx = -MulDiv(10, static_cast<int>(newDpi), 72);

    _uiFont.reset(CreateFontW(uiHeightPx,
                              0,
                              0,
                              0,
                              FW_NORMAL,
                              FALSE,
                              FALSE,
                              FALSE,
                              DEFAULT_CHARSET,
                              OUT_DEFAULT_PRECIS,
                              CLIP_DEFAULT_PRECIS,
                              CLEARTYPE_QUALITY,
                              DEFAULT_PITCH | FF_DONTCARE,
                              L"Segoe UI"));
    if (! _uiFont)
    {
        Debug::ErrorWithLastError(L"ViewerText: CreateFontW failed for UI font (DPI change).");
    }

    _monoFont.reset(CreateFontW(monoHeightPx,
                                0,
                                0,
                                0,
                                FW_NORMAL,
                                FALSE,
                                FALSE,
                                FALSE,
                                DEFAULT_CHARSET,
                                OUT_DEFAULT_PRECIS,
                                CLIP_DEFAULT_PRECIS,
                                CLEARTYPE_QUALITY,
                                FIXED_PITCH | FF_MODERN,
                                L"Consolas"));
    if (! _monoFont)
    {
        Debug::ErrorWithLastError(L"ViewerText: CreateFontW failed for monospace font (DPI change).");
    }

    if (_hFileCombo && _uiFont)
    {
        SendMessageW(_hFileCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_uiFont.get()), TRUE);
    }
    if (_hEdit && _monoFont)
    {
        SendMessageW(_hEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_monoFont.get()), TRUE);
    }
    if (_hHex && _monoFont)
    {
        SendMessageW(_hHex.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_monoFont.get()), TRUE);
    }

    if (_hFileCombo)
    {
        int itemHeight = PxFromDip(24, newDpi);
        auto hdc       = wil::GetDC(hwnd);
        if (hdc)
        {
            HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
            static_cast<void>(oldFont);

            TEXTMETRICW tm{};
            if (GetTextMetricsW(hdc.get(), &tm) != 0)
            {
                itemHeight = tm.tmHeight + tm.tmExternalLeading + PxFromDip(6, newDpi);
            }
        }

        itemHeight = std::max(itemHeight, 1);
        SendMessageW(_hFileCombo.get(), CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), static_cast<LPARAM>(itemHeight));
        SendMessageW(_hFileCombo.get(), CB_SETITEMHEIGHT, 0, static_cast<LPARAM>(itemHeight));
    }

    UpdateHexColumns(hwnd);
    Layout(hwnd);
    InvalidateRect(hwnd, nullptr, TRUE);
}

void ViewerText::Layout(HWND hwnd) noexcept
{
    RECT client{};
    GetClientRect(hwnd, &client);

    const UINT dpi             = GetDpiForWindow(hwnd);
    const int edgeSizeY        = GetSystemMetricsForDpi(SM_CYEDGE, dpi);
    const int baseHeaderHeight = PxFromDip(kHeaderHeightDip, dpi);
    const int statusHeight     = PxFromDip(kStatusHeightDip, dpi);
    const int accentHeight     = std::max(1, PxFromDip(2, dpi));
    const int accentGap        = std::max(1, PxFromDip(1, dpi));
    const int minPadding       = PxFromDip(3, dpi);
    const int comboBorder      = std::max(0, edgeSizeY) * 2;

    const int minChromeHeight = PxFromDip(22, dpi) + accentHeight + accentGap + 2 * minPadding;

    const bool showCombo   = (_hFileCombo && _otherFiles.size() > 1);
    int desiredComboHeight = 0;
    if (showCombo)
    {
        int comboItemHeight           = 0;
        const LRESULT selectionHeight = SendMessageW(_hFileCombo.get(), CB_GETITEMHEIGHT, static_cast<WPARAM>(-1), 0);
        if (selectionHeight != CB_ERR && selectionHeight > 0)
        {
            comboItemHeight = static_cast<int>(selectionHeight);
        }

        if (comboItemHeight <= 0)
        {
            comboItemHeight = PxFromDip(24, dpi);
            auto hdc        = wil::GetDC(hwnd);
            if (hdc)
            {
                HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
                auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
                static_cast<void>(oldFont);

                TEXTMETRICW tm{};
                if (GetTextMetricsW(hdc.get(), &tm) != 0)
                {
                    comboItemHeight = tm.tmHeight + tm.tmExternalLeading + PxFromDip(6, dpi);
                }
            }
        }

        const int comboChromePadding = std::max(PxFromDip(4, dpi), comboBorder);
        desiredComboHeight           = std::max(1, comboItemHeight + comboChromePadding);
    }

    int headerHeight = baseHeaderHeight;
    headerHeight     = std::max(headerHeight, minChromeHeight);
    if (showCombo)
    {
        headerHeight = std::max(headerHeight, desiredComboHeight + accentHeight + accentGap + 2 * minPadding);
    }

    for (int pass = 0; pass < 2; ++pass)
    {
        _headerRect        = client;
        _headerRect.bottom = std::min(client.bottom, client.top + std::max(0, headerHeight));

        _statusRect     = client;
        _statusRect.top = std::max(client.top, client.bottom - std::max(0, statusHeight));

        _contentRect        = client;
        _contentRect.top    = _headerRect.bottom;
        _contentRect.bottom = _statusRect.top;

        ClampRectNonNegative(_headerRect);
        ClampRectNonNegative(_statusRect);
        ClampRectNonNegative(_contentRect);

        RECT headerContentRect{};
        headerContentRect        = _headerRect;
        headerContentRect.top    = std::min(headerContentRect.bottom, headerContentRect.top + minPadding);
        headerContentRect.bottom = std::max(headerContentRect.top, headerContentRect.bottom - accentHeight - accentGap - minPadding);

        const int headerContentH = std::max(0L, headerContentRect.bottom - headerContentRect.top);
        const int margin         = PxFromDip(10, dpi);
        const int buttonH        = std::min(headerContentH, PxFromDip(22, dpi));
        const int buttonW        = PxFromDip(72, dpi);
        const int buttonY        = headerContentRect.top + std::max(0, (headerContentH - buttonH) / 2);
        const int buttonX        = std::max<LONG>(headerContentRect.left, headerContentRect.right - margin - buttonW);
        _modeButtonRect.left     = buttonX;
        _modeButtonRect.top      = buttonY;
        _modeButtonRect.right    = std::min<LONG>(headerContentRect.right, buttonX + buttonW);
        _modeButtonRect.bottom   = std::min<LONG>(headerContentRect.bottom, buttonY + buttonH);

        int measuredComboHeight = 0;
        if (_hFileCombo)
        {
            ShowWindow(_hFileCombo.get(), showCombo ? SW_SHOW : SW_HIDE);
            EnableWindow(_hFileCombo.get(), showCombo ? TRUE : FALSE);

            if (showCombo)
            {
                int comboH = desiredComboHeight;
                comboH     = std::clamp(comboH, 1, std::max(1, headerContentH));

                const int comboX = headerContentRect.left + margin;
                const int comboW = std::max(0, static_cast<int>(_modeButtonRect.left) - margin - comboX);

                SetWindowPos(_hFileCombo.get(), nullptr, comboX, headerContentRect.top, comboW, comboH, SWP_NOZORDER | SWP_NOACTIVATE);

                RECT comboRc{};
                int actualComboH = comboH;
                if (GetWindowRect(_hFileCombo.get(), &comboRc) != 0)
                {
                    actualComboH = std::max(0L, comboRc.bottom - comboRc.top);
                }

                measuredComboHeight = actualComboH;

                int comboY = headerContentRect.top + std::max(0, (headerContentH - actualComboH) / 2);

                const int maxBottom = std::max(static_cast<int>(headerContentRect.top), static_cast<int>(headerContentRect.bottom));
                if (comboY + actualComboH > maxBottom)
                {
                    comboY = std::max(static_cast<int>(headerContentRect.top), maxBottom - actualComboH);
                }

                SetWindowPos(_hFileCombo.get(), nullptr, comboX, comboY, 0, 0, SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOSIZE);
            }
        }

        const int currentHeaderHeight = headerHeight;
        int requiredHeaderHeight      = currentHeaderHeight;
        if (showCombo && measuredComboHeight > 0)
        {
            requiredHeaderHeight = std::max(minChromeHeight, measuredComboHeight + accentHeight + accentGap + 2 * minPadding);
        }

        if (requiredHeaderHeight > currentHeaderHeight && pass == 0)
        {
            headerHeight = requiredHeaderHeight;
            continue;
        }

        break;
    }

    const int contentW = std::max(0L, _contentRect.right - _contentRect.left);
    const int contentH = std::max(0L, _contentRect.bottom - _contentRect.top);

    if (_hEdit)
    {
        SetWindowPos(_hEdit.get(), nullptr, _contentRect.left, _contentRect.top, contentW, contentH, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (_hHex)
    {
        SetWindowPos(_hHex.get(), nullptr, _contentRect.left, _contentRect.top, contentW, contentH, SWP_NOZORDER | SWP_NOACTIVATE);
    }
}

void ViewerText::RefreshFileCombo(HWND hwnd) noexcept
{
    if (! _hFileCombo)
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });

    SendMessageW(_hFileCombo.get(), CB_RESETCONTENT, 0, 0);

    if (_otherFiles.size() <= 1)
    {
        SendMessageW(_hFileCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(-1), 0);
        if (hwnd)
        {
            Layout(hwnd);
            InvalidateRect(hwnd, nullptr, TRUE);
        }
        return;
    }

    for (const auto& path : _otherFiles)
    {
        std::wstring itemText = path.filename().wstring();
        if (itemText.empty())
        {
            itemText = path.wstring();
        }

        SendMessageW(_hFileCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(itemText.c_str()));
    }

    if (_otherIndex >= _otherFiles.size())
    {
        _otherIndex = 0;
    }

    SendMessageW(_hFileCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(_otherIndex), 0);
    SendMessageW(_hFileCombo.get(), CB_SETMINVISIBLE, static_cast<WPARAM>(std::min<size_t>(_otherFiles.size(), 15)), 0);

    if (hwnd)
    {
        Layout(hwnd);
        InvalidateRect(hwnd, nullptr, TRUE);
    }
}

void ViewerText::SyncFileComboSelection() noexcept
{
    if (! _hFileCombo)
    {
        return;
    }

    if (_otherFiles.size() <= 1)
    {
        return;
    }

    if (_otherIndex >= _otherFiles.size())
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });

    SendMessageW(_hFileCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(_otherIndex), 0);
}

bool ViewerText::EnsureDirect2D(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    const UINT dpi   = GetDpiForWindow(hwnd);
    const float dpiF = static_cast<float>(dpi);

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

    if (! _d2dTarget)
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

        const HRESULT hr = _d2dFactory->CreateHwndRenderTarget(props, hwndProps, _d2dTarget.put());
        if (FAILED(hr) || ! _d2dTarget)
        {
            _d2dTarget.reset();
            return false;
        }

        _d2dTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    }
    else
    {
        _d2dTarget->SetDpi(dpiF, dpiF);
    }

    if (! _d2dBrush)
    {
        const HRESULT hr = _d2dTarget->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::Black), _d2dBrush.put());
        if (FAILED(hr) || ! _d2dBrush)
        {
            _d2dBrush.reset();
            return false;
        }
    }

    if (! _headerFormat)
    {
        const HRESULT hr = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 12.0f, L"", _headerFormat.put());
        if (FAILED(hr) || ! _headerFormat)
        {
            _headerFormat.reset();
            return false;
        }

        static_cast<void>(_headerFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));
        static_cast<void>(_headerFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    }

    if (! _headerFormatRight)
    {
        const HRESULT hr = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 12.0f, L"", _headerFormatRight.put());
        if (FAILED(hr) || ! _headerFormatRight)
        {
            _headerFormatRight.reset();
            return false;
        }

        static_cast<void>(_headerFormatRight->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING));
        static_cast<void>(_headerFormatRight->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    }

    if (! _modeButtonFormat)
    {
        const HRESULT hr = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 12.0f, L"", _modeButtonFormat.put());
        if (FAILED(hr) || ! _modeButtonFormat)
        {
            _modeButtonFormat.reset();
            return false;
        }

        static_cast<void>(_modeButtonFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER));
        static_cast<void>(_modeButtonFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    }

    if (! _statusFormat)
    {
        const HRESULT hr = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 11.0f, L"", _statusFormat.put());
        if (FAILED(hr) || ! _statusFormat)
        {
            _statusFormat.reset();
            return false;
        }

        static_cast<void>(_statusFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));
        static_cast<void>(_statusFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    }

    if (! _watermarkFormat)
    {
        const HRESULT hr = _dwriteFactory->CreateTextFormat(L"Segoe UI",
                                                            nullptr,
                                                            DWRITE_FONT_WEIGHT_SEMI_BOLD,
                                                            DWRITE_FONT_STYLE_NORMAL,
                                                            DWRITE_FONT_STRETCH_NORMAL,
                                                            kWatermarkFontSizeDip,
                                                            L"",
                                                            _watermarkFormat.put());
        if (FAILED(hr) || ! _watermarkFormat)
        {
            _watermarkFormat.reset();
            return false;
        }

        static_cast<void>(_watermarkFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER));
        static_cast<void>(_watermarkFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    }

    return true;
}

void ViewerText::DiscardDirect2D() noexcept
{
    _d2dBrush.reset();
    _headerFormat.reset();
    _headerFormatRight.reset();
    _modeButtonFormat.reset();
    _statusFormat.reset();
    _watermarkFormat.reset();
    _d2dTarget.reset();
}

void ViewerText::OnPaint()
{
    if (! _hWnd)
    {
        return;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(_hWnd.get(), &ps);
    _allowEraseBkgnd          = false;

    const int dpiInt = static_cast<int>(GetDpiForWindow(_hWnd.get()));

    if (EnsureDirect2D(_hWnd.get()) && _d2dTarget && _d2dBrush && _headerFormat && _headerFormatRight && _statusFormat && _watermarkFormat)
    {
        const UINT dpi = GetDpiForWindow(_hWnd.get());

        const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
        const COLORREF fg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

        COLORREF headerBg = bg;
        COLORREF statusBg = bg;
        if (_hasTheme && _theme.darkMode)
        {
            headerBg = RGB(std::max(0, GetRValue(bg) - 10), std::max(0, GetGValue(bg) - 10), std::max(0, GetBValue(bg) - 10));
            statusBg = RGB(std::min(255, GetRValue(bg) + 5), std::min(255, GetGValue(bg) + 5), std::min(255, GetBValue(bg) + 5));
        }
        else
        {
            headerBg = RGB(std::max(0, GetRValue(bg) - 5), std::max(0, GetGValue(bg) - 5), std::max(0, GetBValue(bg) - 5));
            statusBg = RGB(std::min(255, GetRValue(bg) + 5), std::min(255, GetGValue(bg) + 5), std::min(255, GetBValue(bg) + 5));
        }

        const std::wstring seed = _currentPath.empty() ? std::wstring(L"viewer") : _currentPath.filename().wstring();
        const COLORREF accent   = _hasTheme ? ResolveAccentColor(_theme, seed) : RGB(0, 120, 215);

        std::wstring titleText;
        if (! _currentPath.empty())
        {
            titleText = _currentPath.filename().wstring();
        }

        const UINT modeId           = _viewMode == ViewMode::Hex ? IDS_VIEWERTEXT_MODE_HEX : IDS_VIEWERTEXT_MODE_TEXT;
        const std::wstring modeText = LoadStringResource(g_hInstance, modeId);

        const std::wstring statusText = BuildStatusText();

        HRESULT hr = S_OK;
        {
            _d2dTarget->BeginDraw();
            auto endDraw = wil::scope_exit(
                [&]
                {
                    hr = _d2dTarget->EndDraw();
                    if (hr == D2DERR_RECREATE_TARGET)
                    {
                        DiscardDirect2D();
                    }
                });

            _d2dTarget->SetTransform(D2D1::Matrix3x2F::Identity());
            _d2dTarget->Clear(ColorFFromColorRef(bg));

            const D2D1_RECT_F headerRc = RectFFromPixels(_headerRect, dpi);
            const D2D1_RECT_F statusRc = RectFFromPixels(_statusRect, dpi);

            _d2dBrush->SetColor(ColorFFromColorRef(headerBg));
            _d2dTarget->FillRectangle(headerRc, _d2dBrush.get());

            _d2dBrush->SetColor(ColorFFromColorRef(statusBg));
            _d2dTarget->FillRectangle(statusRc, _d2dBrush.get());

            const int accentHeightPx = std::max(1, PxFromDip(2, dpi));
            RECT accentPx            = _headerRect;
            accentPx.top             = std::max(accentPx.top, accentPx.bottom - accentHeightPx);
            ClampRectNonNegative(accentPx);
            const D2D1_RECT_F accentRc = RectFFromPixels(accentPx, dpi);

            _d2dBrush->SetColor(ColorFFromColorRef(accent));
            _d2dTarget->FillRectangle(accentRc, _d2dBrush.get());

            const float marginDip    = 10.0f;
            D2D1_RECT_F headerTextRc = headerRc;
            headerTextRc.left += marginDip;
            headerTextRc.right -= marginDip;

            const D2D1_RECT_F modeButtonRc = RectFFromPixels(_modeButtonRect, dpi);
            const float radius             = 2.0f;

            float modeAlpha = 0.16f;
            if (_modeButtonPressed)
            {
                modeAlpha = 0.30f;
            }
            else if (_modeButtonHot)
            {
                modeAlpha = 0.22f;
            }

            _d2dBrush->SetColor(ColorFFromColorRef(accent, modeAlpha));
            _d2dTarget->FillRoundedRectangle(D2D1::RoundedRect(modeButtonRc, radius, radius), _d2dBrush.get());

            _d2dBrush->SetColor(ColorFFromColorRef(accent, 0.85f));
            _d2dTarget->DrawRoundedRectangle(D2D1::RoundedRect(modeButtonRc, radius, radius), _d2dBrush.get(), 1.0f);

            _d2dBrush->SetColor(ColorFFromColorRef(fg));
            _d2dTarget->DrawTextW(modeText.c_str(),
                                  static_cast<UINT32>(modeText.size()),
                                  _modeButtonFormat ? _modeButtonFormat.get() : _headerFormatRight.get(),
                                  modeButtonRc,
                                  _d2dBrush.get(),
                                  D2D1_DRAW_TEXT_OPTIONS_CLIP);

            if (_otherFiles.size() <= 1)
            {
                D2D1_RECT_F fileRc = headerTextRc;
                fileRc.right       = std::max(fileRc.left, modeButtonRc.left - marginDip);
                _d2dTarget->DrawTextW(
                    titleText.c_str(), static_cast<UINT32>(titleText.size()), _headerFormat.get(), fileRc, _d2dBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }

            D2D1_RECT_F statusTextRc = statusRc;
            statusTextRc.left += marginDip;
            statusTextRc.right -= marginDip;
            _d2dTarget->DrawTextW(
                statusText.c_str(), static_cast<UINT32>(statusText.size()), _statusFormat.get(), statusTextRc, _d2dBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

            if (! _isLoading && _fileReader && _fileSize == 0 && ! _currentPath.empty())
            {
                const std::wstring emptyText = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_EMPTY_WATERMARK);
                if (! emptyText.empty())
                {
                    const D2D1_RECT_F contentRc = RectFFromPixels(_contentRect, dpi);
                    const float centerX         = (contentRc.left + contentRc.right) / 2.0f;
                    const float centerY         = (contentRc.top + contentRc.bottom) / 2.0f;
                    const float alpha           = _hasTheme && _theme.darkMode ? 0.28f : 0.20f;
                    _d2dBrush->SetColor(ColorFFromColorRef(fg, alpha));

                    auto restoreTransform          = wil::scope_exit([&] { _d2dTarget->SetTransform(D2D1::Matrix3x2F::Identity()); });
                    const D2D1_MATRIX_3X2_F rotate = D2D1::Matrix3x2F::Rotation(kWatermarkAngleDegrees, D2D1::Point2F(centerX, centerY));
                    _d2dTarget->SetTransform(rotate * D2D1::Matrix3x2F::Identity());
                    _d2dTarget->DrawTextW(emptyText.c_str(),
                                          static_cast<UINT32>(emptyText.size()),
                                          _watermarkFormat.get(),
                                          contentRc,
                                          _d2dBrush.get(),
                                          D2D1_DRAW_TEXT_OPTIONS_CLIP);
                }
            }
        }

        if (hr == D2DERR_RECREATE_TARGET)
        {
            // Discarded in the EndDraw scope.
        }
        else if (FAILED(hr))
        {
            DiscardDirect2D();
        }
        else
        {
            return;
        }
    }

    FillRect(hdc.get(), &ps.rcPaint, _backgroundBrush.get());

    FillRect(hdc.get(), &_headerRect, _headerBrush.get());
    FillRect(hdc.get(), &_statusRect, _statusBrush.get());

    const COLORREF textColor = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
    const std::wstring seed  = _currentPath.empty() ? std::wstring(L"viewer") : _currentPath.filename().wstring();
    const COLORREF accent    = _hasTheme ? ResolveAccentColor(_theme, seed) : RGB(0, 120, 215);

    const int lineThickness = std::max(1, MulDiv(2, dpiInt, USER_DEFAULT_SCREEN_DPI));

    RECT line{};
    line.left   = 0;
    line.right  = ps.rcPaint.right;
    line.top    = _headerRect.bottom - lineThickness;
    line.bottom = _headerRect.bottom;

    wil::unique_hbrush accentBrush(CreateSolidBrush(accent));
    FillRect(hdc.get(), &line, accentBrush.get());

    SetBkMode(hdc.get(), TRANSPARENT);
    SetTextColor(hdc.get(), textColor);
    HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);

    RECT headerTextRc = _headerRect;
    headerTextRc.left += MulDiv(10, dpiInt, USER_DEFAULT_SCREEN_DPI);
    headerTextRc.right -= MulDiv(10, dpiInt, USER_DEFAULT_SCREEN_DPI);

    std::wstring titleText;
    if (! _currentPath.empty())
    {
        titleText = _currentPath.filename().wstring();
    }

    const UINT modeId           = _viewMode == ViewMode::Hex ? IDS_VIEWERTEXT_MODE_HEX : IDS_VIEWERTEXT_MODE_TEXT;
    const std::wstring modeText = LoadStringResource(g_hInstance, modeId);

    RECT modeRc = _modeButtonRect;
    ClampRectNonNegative(modeRc);

    uint8_t alpha = 40u;
    if (_modeButtonPressed)
    {
        alpha = 90u;
    }
    else if (_modeButtonHot)
    {
        alpha = 70u;
    }

    const COLORREF modeBg = BlendColor(_uiHeaderBg, accent, alpha);
    wil::unique_hbrush modeBrush(CreateSolidBrush(modeBg));
    FillRect(hdc.get(), &modeRc, modeBrush.get());
    FrameRect(hdc.get(), &modeRc, accentBrush.get());

    DrawTextW(hdc.get(), modeText.c_str(), static_cast<int>(modeText.size()), &modeRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

    if (_otherFiles.size() <= 1)
    {
        RECT fileRc  = headerTextRc;
        fileRc.right = std::max(fileRc.left, _modeButtonRect.left - MulDiv(10, dpiInt, USER_DEFAULT_SCREEN_DPI));
        DrawTextW(hdc.get(), titleText.c_str(), static_cast<int>(titleText.size()), &fileRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    }

    RECT statusTextRc = _statusRect;
    statusTextRc.left += MulDiv(10, dpiInt, USER_DEFAULT_SCREEN_DPI);
    statusTextRc.right -= MulDiv(10, dpiInt, USER_DEFAULT_SCREEN_DPI);

    const std::wstring statusText = BuildStatusText();
    DrawTextW(hdc.get(), statusText.c_str(), static_cast<int>(statusText.size()), &statusTextRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

    if (_fileSize == 0 && ! _currentPath.empty())
    {
        const std::wstring emptyText = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_EMPTY_WATERMARK);
        if (! emptyText.empty())
        {
            RECT contentRc    = _contentRect;
            const int centerX = (contentRc.left + contentRc.right) / 2;
            const int centerY = (contentRc.top + contentRc.bottom) / 2;

            XFORM xf{};
            xf.eM11 = static_cast<float>(std::cos(kWatermarkAngleRadians));
            xf.eM12 = static_cast<float>(std::sin(kWatermarkAngleRadians));
            xf.eM21 = -xf.eM12;
            xf.eM22 = xf.eM11;
            xf.eDx  = static_cast<float>(centerX);
            xf.eDy  = static_cast<float>(centerY);

            const int oldMode = SetGraphicsMode(hdc.get(), GM_ADVANCED);
            XFORM oldXf{};
            if (GetWorldTransform(hdc.get(), &oldXf) == 0)
            {
                oldXf.eM11 = 1.0f;
                oldXf.eM12 = 0.0f;
                oldXf.eM21 = 0.0f;
                oldXf.eM22 = 1.0f;
                oldXf.eDx  = 0.0f;
                oldXf.eDy  = 0.0f;
            }

            SetWorldTransform(hdc.get(), &xf);

            RECT drawRc = contentRc;
            OffsetRect(&drawRc, -centerX, -centerY);

            const int fontHeight = -MulDiv(static_cast<int>(kWatermarkFontSizeDip), dpiInt, USER_DEFAULT_SCREEN_DPI);
            wil::unique_hfont stampFont(CreateFontW(fontHeight,
                                                    0,
                                                    0,
                                                    0,
                                                    FW_SEMIBOLD,
                                                    FALSE,
                                                    FALSE,
                                                    FALSE,
                                                    DEFAULT_CHARSET,
                                                    OUT_DEFAULT_PRECIS,
                                                    CLIP_DEFAULT_PRECIS,
                                                    CLEARTYPE_QUALITY,
                                                    DEFAULT_PITCH | FF_DONTCARE,
                                                    L"Segoe UI"));
            if (stampFont)
            {
                auto oldStamp = wil::SelectObject(hdc.get(), stampFont.get());
                static_cast<void>(oldStamp);
            }

            const COLORREF stampColor = BlendColor(_uiBackground, _uiText, _hasTheme && _theme.darkMode ? 100u : 70u);
            SetTextColor(hdc.get(), stampColor);
            SetBkMode(hdc.get(), TRANSPARENT);
            DrawTextW(hdc.get(), emptyText.c_str(), static_cast<int>(emptyText.size()), &drawRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);

            SetWorldTransform(hdc.get(), &oldXf);
            if (oldMode != 0)
            {
                SetGraphicsMode(hdc.get(), oldMode);
            }
        }
    }
}

void ViewerText::OnCommand(HWND hwnd, UINT commandId, UINT notifyCode, HWND control) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (_hFileCombo && control == _hFileCombo.get() && commandId == IDC_VIEWERTEXT_FILE_COMBO)
    {
        if (notifyCode == CBN_DROPDOWN)
        {
            COMBOBOXINFO info{};
            info.cbSize = sizeof(info);
            if (GetComboBoxInfo(_hFileCombo.get(), &info) != 0)
            {
                _hFileComboList = info.hwndList;
                _hFileComboItem = info.hwndItem;
            }

            const wchar_t* winTheme = L"Explorer";
            if (_hasTheme && _theme.highContrast)
            {
                winTheme = L"";
            }
            else if (_hasTheme && _theme.darkMode)
            {
                winTheme = L"DarkMode_Explorer";
            }

            SetWindowTheme(_hFileCombo.get(), winTheme, nullptr);
            if (_hFileComboList)
            {
                SetWindowTheme(_hFileComboList, winTheme, nullptr);
                SendMessageW(_hFileComboList, WM_THEMECHANGED, 0, 0);
            }
            if (_hFileComboItem)
            {
                SetWindowTheme(_hFileComboItem, winTheme, nullptr);
                SendMessageW(_hFileComboItem, WM_THEMECHANGED, 0, 0);
            }
            SendMessageW(_hFileCombo.get(), WM_THEMECHANGED, 0, 0);

            return;
        }

        if (notifyCode == CBN_SELCHANGE && ! _syncingFileCombo)
        {
            const LRESULT sel = SendMessageW(_hFileCombo.get(), CB_GETCURSEL, 0, 0);
            if (sel >= 0 && static_cast<size_t>(sel) < _otherFiles.size())
            {
                _otherIndex = static_cast<size_t>(sel);
                static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
            }
        }

        return;
    }

    if (IsEncodingMenuSelectionValid(commandId))
    {
        SetDisplayEncodingMenuSelection(hwnd, commandId, true);
        return;
    }

    if (IsSaveEncodingMenuSelectionValid(commandId))
    {
        SetSaveEncodingMenuSelection(hwnd, commandId);
        return;
    }

    switch (commandId)
    {
        case IDM_VIEWER_FILE_OPEN: CommandOpen(hwnd); break;
        case IDM_VIEWER_FILE_SAVE_AS: CommandSaveAs(hwnd); break;
        case IDM_VIEWER_FILE_REFRESH: CommandRefresh(hwnd); break;
        case IDM_VIEWER_FILE_EXIT: CommandExit(hwnd); break;

        case IDM_VIEWER_OTHER_NEXT: CommandOtherNext(hwnd); break;
        case IDM_VIEWER_OTHER_PREVIOUS: CommandOtherPrevious(hwnd); break;
        case IDM_VIEWER_OTHER_FIRST: CommandOtherFirst(hwnd); break;
        case IDM_VIEWER_OTHER_LAST: CommandOtherLast(hwnd); break;

        case IDM_VIEWER_SEARCH_FIND: CommandFind(hwnd); break;
        case IDM_VIEWER_SEARCH_FIND_NEXT: CommandFindNext(hwnd, false); break;
        case IDM_VIEWER_SEARCH_FIND_PREVIOUS: CommandFindNext(hwnd, true); break;

        case IDM_VIEWER_VIEW_TEXT: SetViewMode(hwnd, ViewMode::Text); break;
        case IDM_VIEWER_VIEW_HEX: SetViewMode(hwnd, ViewMode::Hex); break;
        case IDM_VIEWER_VIEW_GOTO_TOP: CommandGoToTop(hwnd, false); break;
        case IDM_VIEWER_VIEW_GOTO_BOTTOM: CommandGoToBottom(hwnd, false); break;
        case IDM_VIEWER_VIEW_GOTO_OFFSET: CommandGoToOffset(hwnd); break;
        case IDM_VIEWER_VIEW_LINE_NUMBERS: SetShowLineNumbers(hwnd, ! _config.showLineNumbers); break;
        case IDM_VIEWER_VIEW_WRAP: SetWrap(hwnd, ! _wrap); break;

        case IDM_VIEWER_ENCODING_NEXT: CommandCycleDisplayEncoding(hwnd, false); break;
        case IDM_VIEWER_ENCODING_PREVIOUS: CommandCycleDisplayEncoding(hwnd, true); break;
    }
}

LRESULT ViewerText::OnNotify(const NMHDR* header)
{
    if (! header)
    {
        return 0;
    }

    if (_hHex)
    {
        const HWND listHeader = ListView_GetHeader(_hHex.get());
        if (listHeader && header->hwndFrom == listHeader && header->code == NM_CUSTOMDRAW)
        {
            if (! _hasTheme || _theme.highContrast)
            {
                return CDRF_DODEFAULT;
            }

            auto* cd = reinterpret_cast<NMCUSTOMDRAW*>(const_cast<NMHDR*>(header));
            if (! cd)
            {
                return CDRF_DODEFAULT;
            }

            if (cd->dwDrawStage == CDDS_PREPAINT)
            {
                RECT rc{};
                GetClientRect(listHeader, &rc);
                FillRect(cd->hdc, &rc, _headerBrush.get());
                return CDRF_NOTIFYITEMDRAW;
            }

            if (cd->dwDrawStage == CDDS_ITEMPREPAINT)
            {
                const UINT dpi    = GetDpiForWindow(listHeader);
                const int padding = PxFromDip(6, dpi);

                RECT rc = cd->rc;
                FillRect(cd->hdc, &rc, _headerBrush.get());

                const COLORREF bg     = _uiHeaderBg;
                const COLORREF fg     = _uiText;
                const COLORREF border = BlendColor(bg, fg, 80u);

                wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> pen(CreatePen(PS_SOLID, 1, border));
                auto oldPen = wil::SelectObject(cd->hdc, pen.get());
                static_cast<void>(oldPen);

                MoveToEx(cd->hdc, rc.left, rc.bottom - 1, nullptr);
                LineTo(cd->hdc, rc.right, rc.bottom - 1);
                MoveToEx(cd->hdc, rc.right - 1, rc.top, nullptr);
                LineTo(cd->hdc, rc.right - 1, rc.bottom);

                wchar_t textBuf[256]{};
                HDITEMW item{};
                item.mask       = HDI_TEXT;
                item.pszText    = textBuf;
                item.cchTextMax = static_cast<int>(std::size(textBuf));
                Header_GetItem(listHeader, static_cast<int>(cd->dwItemSpec), &item);

                HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
                auto oldFont    = wil::SelectObject(cd->hdc, fontToUse);
                static_cast<void>(oldFont);

                SetBkMode(cd->hdc, TRANSPARENT);
                SetTextColor(cd->hdc, fg);

                RECT textRc = rc;
                textRc.left += padding;
                textRc.right -= padding;
                DrawTextW(cd->hdc, textBuf, -1, &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

                return CDRF_SKIPDEFAULT;
            }

            return CDRF_DODEFAULT;
        }
    }

    if (_hHex && header->hwndFrom == _hHex.get())
    {
        if (header->code == LVN_GETDISPINFO)
        {
            auto* info = reinterpret_cast<NMLVDISPINFOW*>(const_cast<NMHDR*>(header));
            if (! info)
            {
                return 0;
            }

            EnsureHexLineCache(info->item.iItem);
            const wchar_t* src = nullptr;
            switch (info->item.iSubItem)
            {
                case 0: src = _hexLineCacheOffsetText.c_str(); break;
                case 1: src = _hexLineCacheHexText.c_str(); break;
                case 2: src = _hexLineCacheAsciiText.c_str(); break;
                default: src = L""; break;
            }

            if (info->item.pszText && info->item.cchTextMax > 0)
            {
                wcsncpy_s(info->item.pszText, static_cast<size_t>(info->item.cchTextMax), src, _TRUNCATE);
            }

            return 0;
        }

        if (header->code == LVN_COLUMNCLICK)
        {
            auto* info = reinterpret_cast<NMLISTVIEW*>(const_cast<NMHDR*>(header));
            if (! info)
            {
                return 0;
            }

            if (info->iSubItem == 0)
            {
                CycleHexOffsetMode();
                return 0;
            }

            if (info->iSubItem == 1)
            {
                CycleHexColumnMode();
                return 0;
            }

            if (info->iSubItem == 2)
            {
                CycleHexTextMode();
            }

            return 0;
        }

        if (header->code == NM_CUSTOMDRAW)
        {
            auto* cd = reinterpret_cast<NMLVCUSTOMDRAW*>(const_cast<NMHDR*>(header));
            if (! cd)
            {
                return CDRF_DODEFAULT;
            }

            if (cd->nmcd.dwDrawStage == CDDS_PREPAINT)
            {
                return CDRF_NOTIFYITEMDRAW;
            }

            if (cd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)
            {
                return CDRF_NOTIFYSUBITEMDRAW;
            }

            if (cd->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM))
            {
                const int item    = static_cast<int>(cd->nmcd.dwItemSpec);
                const int subItem = cd->iSubItem;

                RECT cell{};
                if (ListView_GetSubItemRect(_hHex.get(), item, subItem, LVIR_BOUNDS, &cell) == FALSE)
                {
                    return CDRF_DODEFAULT;
                }

                const UINT dpi    = _hWnd ? GetDpiForWindow(_hWnd.get()) : USER_DEFAULT_SCREEN_DPI;
                const int padding = PxFromDip(6, dpi);

                const COLORREF baseBg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
                const COLORREF baseFg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

                const std::wstring seed = _currentPath.empty() ? std::wstring(L"viewer") : _currentPath.filename().wstring();
                const COLORREF accent   = _hasTheme ? ResolveAccentColor(_theme, seed) : RGB(0, 120, 215);

                const UINT state    = ListView_GetItemState(_hHex.get(), item, LVIS_SELECTED);
                const bool selected = (state & LVIS_SELECTED) != 0;

                const COLORREF rowBg = selected ? accent : baseBg;
                const COLORREF rowFg = selected ? ContrastingTextColor(rowBg) : baseFg;

                SetDCBrushColor(cd->nmcd.hdc, rowBg);
                FillRect(cd->nmcd.hdc, &cell, static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));

                HFONT fontToUse = _monoFont ? _monoFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
                auto oldFont    = wil::SelectObject(cd->nmcd.hdc, fontToUse);
                static_cast<void>(oldFont);

                SetBkMode(cd->nmcd.hdc, TRANSPARENT);
                SetTextColor(cd->nmcd.hdc, rowFg);

                TEXTMETRICW tm{};
                GetTextMetricsW(cd->nmcd.hdc, &tm);
                const int cellH     = std::max(0L, cell.bottom - cell.top);
                const int y         = cell.top + std::max(0, (cellH - static_cast<int>(tm.tmHeight)) / 2);
                const int charWidth = std::max(1, static_cast<int>(tm.tmAveCharWidth));

                const int x0 = cell.left + padding;

                const uint64_t lineOffset = static_cast<uint64_t>(item) * static_cast<uint64_t>(kHexBytesPerLine);
                EnsureHexLineCache(item);

                const wchar_t* src = nullptr;
                size_t srcLen      = 0;
                switch (subItem)
                {
                    case 0:
                        src    = _hexLineCacheOffsetText.c_str();
                        srcLen = _hexLineCacheOffsetText.size();
                        break;
                    case 1:
                        src    = _hexLineCacheHexText.c_str();
                        srcLen = _hexLineCacheHexText.size();
                        break;
                    case 2:
                        src    = _hexLineCacheAsciiText.c_str();
                        srcLen = _hexLineCacheAsciiText.size();
                        break;
                    default:
                        src    = L"";
                        srcLen = 0;
                        break;
                }

                bool hasHighlight     = false;
                size_t highlightStart = 0;
                size_t highlightLen   = 0;
                if (_hexSelectedOffset.has_value() && (subItem == 1 || subItem == 2))
                {
                    const uint64_t sel = _hexSelectedOffset.value();
                    if (sel >= lineOffset)
                    {
                        const size_t lineBytes = _hexLineCacheValidBytes;
                        const size_t byteIndex = static_cast<size_t>(sel - lineOffset);
                        if (byteIndex < lineBytes)
                        {
                            const ByteSpan* spans = (subItem == 1) ? _hexLineCacheHexSpans.data() : _hexLineCacheTextSpans.data();
                            const ByteSpan span   = spans[byteIndex];
                            if (span.length > 0 && span.start < srcLen)
                            {
                                hasHighlight   = true;
                                highlightStart = span.start;
                                highlightLen   = span.length;
                            }
                        }
                    }
                }

                RECT clip = cell;
                clip.left = x0;

                if (! hasHighlight || highlightLen == 0 || highlightStart >= static_cast<size_t>(srcLen))
                {
                    ExtTextOutW(cd->nmcd.hdc, x0, y, ETO_CLIPPED, &clip, src, static_cast<UINT>(srcLen), nullptr);
                    return CDRF_SKIPDEFAULT;
                }

                const size_t start = highlightStart;
                const size_t end   = std::min(static_cast<size_t>(srcLen), start + highlightLen);

                const int highlightX = x0 + static_cast<int>(start) * charWidth;
                const int highlightW = std::max(0, static_cast<int>(end - start) * charWidth);

                const uint8_t alpha        = 160u;
                const COLORREF highlightBg = selected ? BlendColor(accent, baseBg, alpha) : BlendColor(baseBg, accent, alpha);
                const COLORREF highlightFg = ContrastingTextColor(highlightBg);

                const UINT preLen = static_cast<UINT>(start);
                if (preLen > 0)
                {
                    ExtTextOutW(cd->nmcd.hdc, x0, y, ETO_CLIPPED, &clip, src, preLen, nullptr);
                }

                RECT highlightRc  = cell;
                highlightRc.left  = highlightX;
                highlightRc.right = std::min(cell.right, static_cast<LONG>(highlightX + highlightW));
                if (highlightRc.right > highlightRc.left)
                {
                    SetDCBrushColor(cd->nmcd.hdc, highlightBg);
                    FillRect(cd->nmcd.hdc, &highlightRc, static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));

                    SetTextColor(cd->nmcd.hdc, highlightFg);
                    const int hlLenInt = static_cast<int>(end - start);
                    ExtTextOutW(cd->nmcd.hdc, highlightX, y, ETO_CLIPPED, &clip, src + start, static_cast<UINT>(hlLenInt), nullptr);
                    SetTextColor(cd->nmcd.hdc, rowFg);
                }

                const int postStart = static_cast<int>(end);
                const int postLen   = std::max(0, static_cast<int>(srcLen - postStart));
                if (postLen > 0)
                {
                    const int postX = x0 + postStart * charWidth;
                    ExtTextOutW(cd->nmcd.hdc, postX, y, ETO_CLIPPED, &clip, src + postStart, static_cast<UINT>(postLen), nullptr);
                }

                return CDRF_SKIPDEFAULT;
            }

            return CDRF_DODEFAULT;
        }
    }

    return 0;
}

LRESULT ViewerText::OnMeasureItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept
{
    if (! measure)
    {
        return FALSE;
    }

    if (measure->CtlType == ODT_MENU)
    {
        OnMeasureMenuItem(hwnd, measure);
        return TRUE;
    }

    if (measure->CtlType == ODT_COMBOBOX && measure->CtlID == IDC_VIEWERTEXT_FILE_COMBO)
    {
        OnMeasureFileComboItem(hwnd, measure);
        return TRUE;
    }

    return FALSE;
}

LRESULT ViewerText::OnDrawItem(DRAWITEMSTRUCT* draw) noexcept
{
    if (! draw)
    {
        return FALSE;
    }

    if (draw->CtlType == ODT_MENU)
    {
        OnDrawMenuItem(draw);
        return TRUE;
    }

    if (draw->CtlType == ODT_COMBOBOX && _hFileCombo && draw->hwndItem == _hFileCombo.get())
    {
        OnDrawFileComboItem(draw);
        return TRUE;
    }

    return FALSE;
}

void ViewerText::OnMeasureFileComboItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept
{
    if (! measure)
    {
        return;
    }

    const UINT dpi = hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;

    int height = PxFromDip(24, dpi);
    auto hdc   = wil::GetDC(hwnd);
    if (hdc)
    {
        HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
        static_cast<void>(oldFont);

        TEXTMETRICW tm{};
        if (GetTextMetricsW(hdc.get(), &tm) != 0)
        {
            height = tm.tmHeight + tm.tmExternalLeading + PxFromDip(6, dpi);
        }
    }

    measure->itemHeight = static_cast<UINT>(std::max(height, 1));
}

void ViewerText::OnDrawFileComboItem(DRAWITEMSTRUCT* draw) noexcept
{
    if (! draw || ! draw->hDC)
    {
        return;
    }

    const UINT dpi    = _hWnd ? GetDpiForWindow(_hWnd.get()) : USER_DEFAULT_SCREEN_DPI;
    const int padding = PxFromDip(6, dpi);

    const bool selected = (draw->itemState & ODS_SELECTED) != 0;
    const bool disabled = (draw->itemState & ODS_DISABLED) != 0;

    COLORREF baseBg = _uiHeaderBg;
    COLORREF baseFg = _uiText;
    COLORREF selBg  = _hasTheme && ! _theme.highContrast ? ResolveAccentColor(_theme, L"combo") : GetSysColor(COLOR_HIGHLIGHT);
    COLORREF selFg  = _hasTheme && ! _theme.highContrast ? ContrastingTextColor(selBg) : GetSysColor(COLOR_HIGHLIGHTTEXT);

    if (_theme.highContrast)
    {
        baseBg = GetSysColor(COLOR_WINDOW);
        baseFg = GetSysColor(COLOR_WINDOWTEXT);
    }

    COLORREF fillColor = selected ? selBg : baseBg;
    COLORREF textColor = selected ? selFg : baseFg;

    if (disabled)
    {
        textColor = BlendColor(fillColor, textColor, 120u);
    }

    wil::unique_hbrush bgBrush(CreateSolidBrush(fillColor));
    FillRect(draw->hDC, &draw->rcItem, bgBrush.get());

    int itemId = static_cast<int>(draw->itemID);
    if (itemId < 0 && _hFileCombo)
    {
        const LRESULT sel = SendMessageW(_hFileCombo.get(), CB_GETCURSEL, 0, 0);
        if (sel >= 0)
        {
            itemId = static_cast<int>(sel);
        }
    }

    std::wstring text;
    if (itemId >= 0 && _hFileCombo)
    {
        const LRESULT lenRes = SendMessageW(_hFileCombo.get(), CB_GETLBTEXTLEN, static_cast<WPARAM>(itemId), 0);
        const int len        = (lenRes > 0) ? static_cast<int>(lenRes) : 0;
        if (len > 0)
        {
            text.resize(static_cast<size_t>(len) + 1);
            SendMessageW(_hFileCombo.get(), CB_GETLBTEXT, static_cast<WPARAM>(itemId), reinterpret_cast<LPARAM>(text.data()));
            text.resize(wcsnlen(text.c_str(), text.size()));
        }
    }

    HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    auto oldFont    = wil::SelectObject(draw->hDC, fontToUse);
    static_cast<void>(oldFont);

    SetBkMode(draw->hDC, TRANSPARENT);
    SetTextColor(draw->hDC, textColor);

    RECT textRc = draw->rcItem;
    textRc.left += padding;
    textRc.right -= padding;
    DrawTextW(draw->hDC, text.c_str(), static_cast<int>(text.size()), &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

    if ((draw->itemState & ODS_FOCUS) != 0)
    {
        DrawFocusRect(draw->hDC, &draw->rcItem);
    }
}

void ViewerText::UpdateMenuChecks(HWND hwnd) noexcept
{
    HMENU menu = GetMenu(hwnd);
    if (! menu)
    {
        return;
    }

    const UINT selectedDisplay = EffectiveDisplayEncodingMenuSelection();
    HMENU encodingMenu         = nullptr;
    const int topCount         = GetMenuItemCount(menu);
    if (topCount > 0)
    {
        for (UINT pos = 0; pos < static_cast<UINT>(topCount); ++pos)
        {
            MENUITEMINFOW info{};
            info.cbSize = sizeof(info);
            info.fMask  = MIIM_SUBMENU;
            if (GetMenuItemInfoW(menu, pos, TRUE, &info) == 0)
            {
                continue;
            }

            if (! info.hSubMenu)
            {
                continue;
            }

            if (GetMenuState(info.hSubMenu, IDM_VIEWER_ENCODING_DISPLAY_ANSI, MF_BYCOMMAND) != static_cast<UINT>(-1))
            {
                encodingMenu = info.hSubMenu;
                break;
            }
        }
    }

    if (encodingMenu)
    {
        auto updateEncodingChecks = [&](auto&& self, HMENU currentMenu) noexcept -> void
        {
            if (! currentMenu)
            {
                return;
            }

            const int count = GetMenuItemCount(currentMenu);
            if (count <= 0)
            {
                Debug::Error(L"Encoding menu has no items");
                return;
            }

            for (UINT pos = 0; pos < static_cast<UINT>(count); ++pos)
            {
                MENUITEMINFOW info{};
                info.cbSize = sizeof(info);
                info.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_ID | MIIM_SUBMENU;
                if (GetMenuItemInfoW(currentMenu, pos, TRUE, &info) == 0)
                {
                    continue;
                }

                if (info.hSubMenu)
                {
                    self(self, info.hSubMenu);
                    continue;
                }

                if ((info.fType & MFT_SEPARATOR) != 0)
                {
                    continue;
                }

                if (! IsEncodingMenuSelectionValid(info.wID))
                {
                    continue;
                }

                info.fType |= MFT_RADIOCHECK;

                info.fState &= ~MFS_CHECKED;
                if (info.wID == selectedDisplay)
                {
                    info.fState |= MFS_CHECKED;
                }

                static_cast<void>(SetMenuItemInfoW(currentMenu, pos, TRUE, &info));
            }
        };

        updateEncodingChecks(updateEncodingChecks, encodingMenu);
    }

    const UINT selectedSave = EffectiveSaveEncodingMenuSelection();
    CheckMenuRadioItem(menu, IDM_VIEWER_ENCODING_SAVE_FIRST, IDM_VIEWER_ENCODING_SAVE_LAST, selectedSave, MF_BYCOMMAND);

    CheckMenuItem(menu, IDM_VIEWER_VIEW_TEXT, static_cast<UINT>(MF_BYCOMMAND | (_viewMode == ViewMode::Text ? MF_CHECKED : MF_UNCHECKED)));
    CheckMenuItem(menu, IDM_VIEWER_VIEW_HEX, static_cast<UINT>(MF_BYCOMMAND | (_viewMode == ViewMode::Hex ? MF_CHECKED : MF_UNCHECKED)));
    CheckMenuItem(menu, IDM_VIEWER_VIEW_LINE_NUMBERS, static_cast<UINT>(MF_BYCOMMAND | (_config.showLineNumbers ? MF_CHECKED : MF_UNCHECKED)));
    CheckMenuItem(menu, IDM_VIEWER_VIEW_WRAP, static_cast<UINT>(MF_BYCOMMAND | (_wrap ? MF_CHECKED : MF_UNCHECKED)));

    EnableMenuItem(menu, IDM_VIEWER_VIEW_LINE_NUMBERS, static_cast<UINT>(MF_BYCOMMAND | (_viewMode == ViewMode::Text ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, IDM_VIEWER_VIEW_WRAP, static_cast<UINT>(MF_BYCOMMAND | (_viewMode == ViewMode::Text ? MF_ENABLED : MF_GRAYED)));
}

LRESULT ViewerText::OnCtlColor([[maybe_unused]] UINT msg, HDC hdc, HWND control) noexcept
{
    if (! hdc || ! control || ! _hasTheme)
    {
        return 0;
    }

    if (_theme.highContrast)
    {
        return 0;
    }

    if (_hFileCombo && (control == _hFileCombo.get() || (_hFileComboList != nullptr && control == _hFileComboList) ||
                        (_hFileComboItem != nullptr && control == _hFileComboItem)))
    {
        SetBkMode(hdc, OPAQUE);
        SetTextColor(hdc, _uiText);
        SetBkColor(hdc, _uiHeaderBg);
        return reinterpret_cast<LRESULT>(_headerBrush.get());
    }

    return 0;
}

void ViewerText::OnMouseMove(int x, int y) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const POINT pt{.x = x, .y = y};
    const bool hot = PtInRect(&_modeButtonRect, pt) != 0;

    if (hot != _modeButtonHot)
    {
        _modeButtonHot = hot;
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
    }

    if (! _trackingMouseLeave)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = _hWnd.get();
        if (TrackMouseEvent(&tme) != 0)
        {
            _trackingMouseLeave = true;
        }
    }
}

void ViewerText::OnMouseLeave() noexcept
{
    _trackingMouseLeave = false;
    if (_modeButtonHot)
    {
        _modeButtonHot = false;
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        }
    }
}

void ViewerText::OnLButtonDown(int x, int y) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const POINT pt{.x = x, .y = y};
    if (PtInRect(&_modeButtonRect, pt) == 0)
    {
        return;
    }

    _modeButtonPressed = true;
    SetCapture(_hWnd.get());
    InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
}

void ViewerText::OnLButtonUp(int x, int y) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const HWND captured = GetCapture();
    if (captured == _hWnd.get())
    {
        ReleaseCapture();
    }

    const bool wasPressed = _modeButtonPressed;
    _modeButtonPressed    = false;

    if (wasPressed)
    {
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);

        const POINT pt{.x = x, .y = y};
        if (PtInRect(&_modeButtonRect, pt) != 0)
        {
            SetViewMode(_hWnd.get(), _viewMode == ViewMode::Hex ? ViewMode::Text : ViewMode::Hex);
        }
    }
}

bool ViewerText::OnSetCursor(HWND hwnd, LPARAM lParam) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    if (LOWORD(lParam) != HTCLIENT)
    {
        return false;
    }

    POINT pt{};
    if (GetCursorPos(&pt) == 0)
    {
        return false;
    }

    if (ScreenToClient(hwnd, &pt) == 0)
    {
        return false;
    }

    if (PtInRect(&_modeButtonRect, pt) == 0)
    {
        return false;
    }

    SetCursor(LoadCursorW(nullptr, IDC_HAND));
    return true;
}

void ViewerText::ApplyTheme(HWND hwnd) noexcept
{
    const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

    _backgroundBrush.reset(CreateSolidBrush(bg));

    COLORREF headerBg = bg;
    COLORREF statusBg = bg;
    if (_hasTheme && _theme.darkMode)
    {
        headerBg = RGB(std::max(0, GetRValue(bg) - 10), std::max(0, GetGValue(bg) - 10), std::max(0, GetBValue(bg) - 10));
        statusBg = RGB(std::min(255, GetRValue(bg) + 5), std::min(255, GetGValue(bg) + 5), std::min(255, GetBValue(bg) + 5));
    }
    else
    {
        headerBg = RGB(std::max(0, GetRValue(bg) - 5), std::max(0, GetGValue(bg) - 5), std::max(0, GetBValue(bg) - 5));
        statusBg = RGB(std::min(255, GetRValue(bg) + 5), std::min(255, GetGValue(bg) + 5), std::min(255, GetBValue(bg) + 5));
    }

    _headerBrush.reset(CreateSolidBrush(headerBg));
    _statusBrush.reset(CreateSolidBrush(statusBg));

    _uiBackground = bg;
    _uiText       = fg;
    _uiHeaderBg   = headerBg;
    _uiStatusBg   = statusBg;

    if (_hasTheme && _hWnd)
    {
        const bool windowActive = GetActiveWindow() == _hWnd.get();
        ApplyTitleBarTheme(windowActive);
    }

    const wchar_t* winTheme = L"Explorer";
    if (_hasTheme && _theme.highContrast)
    {
        winTheme = L"";
    }
    else if (_hasTheme && _theme.darkMode)
    {
        winTheme = L"DarkMode_Explorer";
    }

    if (_hEdit)
    {
        SetWindowTheme(_hEdit.get(), winTheme, nullptr);
        SendMessageW(_hEdit.get(), WM_THEMECHANGED, 0, 0);
    }
    if (_hHex)
    {
        SetWindowTheme(_hHex.get(), winTheme, nullptr);
        SendMessageW(_hHex.get(), WM_THEMECHANGED, 0, 0);
    }

    if (_hFileCombo)
    {
        SetWindowTheme(_hFileCombo.get(), winTheme, nullptr);
        SendMessageW(_hFileCombo.get(), WM_THEMECHANGED, 0, 0);
        if (_hFileComboList)
        {
            SetWindowTheme(_hFileComboList, winTheme, nullptr);
            SendMessageW(_hFileComboList, WM_THEMECHANGED, 0, 0);
        }
        if (_hFileComboItem)
        {
            SetWindowTheme(_hFileComboItem, winTheme, nullptr);
            SendMessageW(_hFileComboItem, WM_THEMECHANGED, 0, 0);
        }
    }

    ApplyMenuTheme(hwnd);
    UpdateMenuChecks(hwnd);

    if (_hEdit)
    {
        InvalidateRect(_hEdit.get(), nullptr, TRUE);
    }
    if (_hHex)
    {
        InvalidateRect(_hHex.get(), nullptr, TRUE);
    }
}

void ViewerText::ApplyTitleBarTheme(bool windowActive) noexcept
{
    if (! _hasTheme || ! _hWnd)
    {
        return;
    }

    static constexpr DWORD kDwmwaUseImmersiveDarkMode19 = 19u;
    static constexpr DWORD kDwmwaUseImmersiveDarkMode20 = 20u;
    static constexpr DWORD kDwmwaBorderColor            = 34u;
    static constexpr DWORD kDwmwaCaptionColor           = 35u;
    static constexpr DWORD kDwmwaTextColor              = 36u;
    static constexpr DWORD kDwmColorDefault             = 0xFFFFFFFFu;

    const BOOL darkMode = (_theme.darkMode && ! _theme.highContrast) ? TRUE : FALSE;
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaUseImmersiveDarkMode20, &darkMode, sizeof(darkMode));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaUseImmersiveDarkMode19, &darkMode, sizeof(darkMode));

    DWORD borderValue  = kDwmColorDefault;
    DWORD captionValue = kDwmColorDefault;
    DWORD textValue    = kDwmColorDefault;
    if (! _theme.highContrast && _theme.rainbowMode)
    {
        COLORREF accent = ResolveAccentColor(_theme, L"title");
        if (! windowActive)
        {
            static constexpr uint8_t kInactiveTitleBlendAlpha = 223u; // ~7/8 toward background
            const COLORREF bg                                 = ColorRefFromArgb(_theme.backgroundArgb);
            accent                                            = BlendColor(accent, bg, kInactiveTitleBlendAlpha);
        }

        const COLORREF text = ContrastingTextColor(accent);
        borderValue         = static_cast<DWORD>(accent);
        captionValue        = static_cast<DWORD>(accent);
        textValue           = static_cast<DWORD>(text);
    }

    DwmSetWindowAttribute(_hWnd.get(), kDwmwaBorderColor, &borderValue, sizeof(borderValue));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaCaptionColor, &captionValue, sizeof(captionValue));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaTextColor, &textValue, sizeof(textValue));
}

void ViewerText::SetViewMode(HWND hwnd, ViewMode mode) noexcept
{
    const ViewMode previous = _viewMode;
    _viewMode               = mode;

    const bool showContent = _isLoading || ! (_fileSize == 0 && ! _currentPath.empty());

    if (_hEdit)
    {
        ShowWindow(_hEdit.get(), showContent && _viewMode == ViewMode::Text ? SW_SHOW : SW_HIDE);
    }
    if (_hHex)
    {
        ShowWindow(_hHex.get(), showContent && _viewMode == ViewMode::Hex ? SW_SHOW : SW_HIDE);
    }

    if (! showContent)
    {
        if (_hWnd)
        {
            SetFocus(_hWnd.get());
        }
    }
    else if (_viewMode == ViewMode::Hex)
    {
        if (previous != _viewMode && ! _isLoading && _hexBytes.empty() && _hexCacheValid == 0)
        {
            static_cast<void>(LoadHexData(hwnd));
        }
        if (_hHex)
        {
            SetFocus(_hHex.get());
        }
    }
    else
    {
        if (_hEdit)
        {
            SetFocus(_hEdit.get());
        }
    }

    UpdateMenuChecks(hwnd);
    InvalidateRect(hwnd, nullptr, TRUE);
}

void ViewerText::CommandExit(HWND /*hwnd*/) noexcept
{
    static_cast<void>(Close());
}

std::optional<std::filesystem::path> ViewerText::ShowOpenDialog(HWND hwnd) noexcept
{
    wil::com_ptr<IFileOpenDialog> dialog;
    const HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.put()));
    if (FAILED(hr) || ! dialog)
    {
        return std::nullopt;
    }

    const std::wstring title = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DIALOG_OPEN_TITLE);
    if (! title.empty())
    {
        static_cast<void>(dialog->SetTitle(title.c_str()));
    }

    DWORD options = 0;
    static_cast<void>(dialog->GetOptions(&options));
    options |= FOS_FORCEFILESYSTEM | FOS_FILEMUSTEXIST | FOS_PATHMUSTEXIST;
    static_cast<void>(dialog->SetOptions(options));

    COMDLG_FILTERSPEC spec{};
    const std::wstring allFiles = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DIALOG_FILTER_ALL_FILES);
    spec.pszName                = allFiles.c_str();
    spec.pszSpec                = L"*.*";
    static_cast<void>(dialog->SetFileTypes(1, &spec));

    const HRESULT showHr = dialog->Show(hwnd);
    if (FAILED(showHr))
    {
        return std::nullopt;
    }

    wil::com_ptr<IShellItem> item;
    const HRESULT itemHr = dialog->GetResult(item.put());
    if (FAILED(itemHr) || ! item)
    {
        return std::nullopt;
    }

    wil::unique_cotaskmem_string path;
    const HRESULT nameHr = item->GetDisplayName(SIGDN_FILESYSPATH, path.put());
    if (FAILED(nameHr) || ! path)
    {
        return std::nullopt;
    }

    return std::filesystem::path(path.get());
}

std::optional<ViewerText::SaveAsResult> ViewerText::ShowSaveAsDialog(HWND hwnd) noexcept
{
    wil::com_ptr<IFileSaveDialog> dialog;
    const HRESULT hr = CoCreateInstance(CLSID_FileSaveDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.put()));
    if (FAILED(hr) || ! dialog)
    {
        return std::nullopt;
    }

    const std::wstring title = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DIALOG_SAVE_TITLE);
    if (! title.empty())
    {
        static_cast<void>(dialog->SetTitle(title.c_str()));
    }

    DWORD options = 0;
    static_cast<void>(dialog->GetOptions(&options));
    options |= FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_OVERWRITEPROMPT;
    static_cast<void>(dialog->SetOptions(options));

    if (! _currentPath.empty())
    {
        const std::wstring fileName = _currentPath.filename().wstring();
        if (! fileName.empty())
        {
            static_cast<void>(dialog->SetFileName(fileName.c_str()));
        }
    }

    UINT initialEncodingSelection = IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL;
    switch (EffectiveSaveEncodingMenuSelection())
    {
        case IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL: initialEncodingSelection = IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL; break;
        case IDM_VIEWER_ENCODING_SAVE_ANSI: initialEncodingSelection = IDM_VIEWER_ENCODING_DISPLAY_ANSI; break;
        case IDM_VIEWER_ENCODING_SAVE_UTF8: initialEncodingSelection = IDM_VIEWER_ENCODING_DISPLAY_UTF8; break;
        case IDM_VIEWER_ENCODING_SAVE_UTF8_BOM: initialEncodingSelection = IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM; break;
        case IDM_VIEWER_ENCODING_SAVE_UTF16BE_BOM: initialEncodingSelection = IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM; break;
        case IDM_VIEWER_ENCODING_SAVE_UTF16LE_BOM: initialEncodingSelection = IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM; break;
        default: initialEncodingSelection = IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL; break;
    }

    static constexpr DWORD kEncodingComboId = 6100u;

    wil::com_ptr<IFileDialogCustomize> customize;
    static_cast<void>(dialog->QueryInterface(IID_PPV_ARGS(customize.put())));
    if (customize)
    {
        const std::wstring encodingLabel = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_SAVEAS_ENCODING_LABEL);
        static_cast<void>(customize->AddComboBox(kEncodingComboId));
        if (! encodingLabel.empty())
        {
            static_cast<void>(customize->SetControlLabel(kEncodingComboId, encodingLabel.c_str()));
        }

        auto stripMenuText = [](std::wstring_view text) -> std::wstring
        {
            const size_t tabPos = text.find(L'\t');
            if (tabPos != std::wstring_view::npos)
            {
                text = text.substr(0, tabPos);
            }

            std::wstring result;
            result.reserve(text.size());

            for (size_t i = 0; i < text.size(); ++i)
            {
                const wchar_t ch = text[i];
                if (ch != L'&')
                {
                    result.push_back(ch);
                    continue;
                }

                if ((i + 1) < text.size() && text[i + 1] == L'&')
                {
                    result.push_back(L'&');
                    i += 1;
                }
            }

            while (! result.empty() && result.front() == L' ')
            {
                result.erase(result.begin());
            }
            while (! result.empty() && result.back() == L' ')
            {
                result.pop_back();
            }

            return result;
        };

        auto addMenuItemToCombo = [&](UINT commandId) noexcept
        {
            if (! hwnd)
            {
                return;
            }

            HMENU menu = GetMenu(hwnd);
            if (! menu)
            {
                return;
            }

            auto findMenuText = [&](auto&& self, HMENU currentMenu, UINT targetId) -> std::wstring
            {
                if (! currentMenu)
                {
                    return {};
                }

                const int count = GetMenuItemCount(currentMenu);
                if (count <= 0)
                {
                    Debug::Error(L"findMenuText: Menu has no items");
                    return {};
                }

                for (UINT pos = 0; pos < static_cast<UINT>(count); ++pos)
                {
                    MENUITEMINFOW info{};
                    info.cbSize = sizeof(info);
                    info.fMask  = MIIM_ID | MIIM_SUBMENU;
                    if (GetMenuItemInfoW(currentMenu, pos, TRUE, &info) == 0)
                    {
                        continue;
                    }

                    if (info.hSubMenu)
                    {
                        std::wstring sub = self(self, info.hSubMenu, targetId);
                        if (! sub.empty())
                        {
                            return sub;
                        }
                    }

                    if (info.wID != targetId)
                    {
                        continue;
                    }

                    wchar_t raw[256]{};
                    const int len = GetMenuStringW(currentMenu, pos, raw, static_cast<int>(std::size(raw)), MF_BYPOSITION);
                    if (len <= 0)
                    {
                        return {};
                    }

                    return stripMenuText(std::wstring_view(raw, static_cast<size_t>(len)));
                }

                return {};
            };

            std::wstring text = findMenuText(findMenuText, menu, commandId);
            if (text.empty())
            {
                return;
            }

            static_cast<void>(customize->AddControlItem(kEncodingComboId, commandId, text.c_str()));
        };

        addMenuItemToCombo(IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL);
        if (hwnd)
        {
            HMENU rootMenu = GetMenu(hwnd);
            if (rootMenu)
            {
                HMENU encodingMenu = nullptr;
                const int topCount = GetMenuItemCount(rootMenu);
                if (topCount <= 0)
                {
                    Debug::Error(L"addMenuItemToCombo: No top-level menu items");
                    return std::nullopt;
                }

                for (UINT pos = 0; pos < static_cast<UINT>(topCount); ++pos)
                {
                    MENUITEMINFOW info{};
                    info.cbSize = sizeof(info);
                    info.fMask  = MIIM_SUBMENU;
                    if (GetMenuItemInfoW(rootMenu, pos, TRUE, &info) == 0)
                    {
                        continue;
                    }

                    if (! info.hSubMenu)
                    {
                        continue;
                    }

                    if (GetMenuState(info.hSubMenu, IDM_VIEWER_ENCODING_DISPLAY_ANSI, MF_BYCOMMAND) != static_cast<UINT>(-1))
                    {
                        encodingMenu = info.hSubMenu;
                        break;
                    }
                }

                if (encodingMenu)
                {
                    auto addEncodingItems = [&](auto&& self, HMENU currentMenu) noexcept -> void
                    {
                        if (! currentMenu)
                        {
                            return;
                        }

                        const int count = GetMenuItemCount(currentMenu);
                        if (count <= 0)
                        {
                            Debug::Error(L"addMenuItemToCombo: Encoding menu has no items");
                            return;
                        }

                        for (UINT pos = 0; pos < static_cast<UINT>(count); ++pos)
                        {
                            MENUITEMINFOW info{};
                            info.cbSize = sizeof(info);
                            info.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_SUBMENU;
                            if (GetMenuItemInfoW(currentMenu, pos, TRUE, &info) == 0)
                            {
                                continue;
                            }

                            if (info.hSubMenu)
                            {
                                self(self, info.hSubMenu);
                                continue;
                            }

                            if ((info.fType & MFT_SEPARATOR) != 0)
                            {
                                continue;
                            }

                            if (! IsEncodingMenuSelectionValid(info.wID))
                            {
                                continue;
                            }

                            wchar_t raw[256]{};
                            const int len = GetMenuStringW(currentMenu, pos, raw, static_cast<int>(std::size(raw)), MF_BYPOSITION);
                            if (len <= 0)
                            {
                                continue;
                            }

                            std::wstring text = stripMenuText(std::wstring_view(raw, static_cast<size_t>(len)));
                            if (text.empty())
                            {
                                continue;
                            }

                            static_cast<void>(customize->AddControlItem(kEncodingComboId, info.wID, text.c_str()));
                        }
                    };

                    addEncodingItems(addEncodingItems, encodingMenu);
                }
            }
        }

        static_cast<void>(customize->SetSelectedControlItem(kEncodingComboId, initialEncodingSelection));
    }

    const HRESULT showHr = dialog->Show(hwnd);
    if (FAILED(showHr))
    {
        return std::nullopt;
    }

    wil::com_ptr<IShellItem> item;
    const HRESULT itemHr = dialog->GetResult(item.put());
    if (FAILED(itemHr) || ! item)
    {
        return std::nullopt;
    }

    wil::unique_cotaskmem_string path;
    const HRESULT nameHr = item->GetDisplayName(SIGDN_FILESYSPATH, path.put());
    if (FAILED(nameHr) || ! path)
    {
        return std::nullopt;
    }

    DWORD selectedEncoding = initialEncodingSelection;
    if (customize)
    {
        static_cast<void>(customize->GetSelectedControlItem(kEncodingComboId, &selectedEncoding));
    }

    SaveAsResult result{};
    result.path              = std::filesystem::path(path.get());
    result.encodingSelection = static_cast<UINT>(selectedEncoding);
    return result;
}

void ViewerText::CommandOpen(HWND hwnd)
{
    const auto path = ShowOpenDialog(hwnd);
    if (! path.has_value())
    {
        return;
    }

    static_cast<void>(OpenPath(hwnd, path.value(), true));
}

void ViewerText::CommandSaveAs(HWND hwnd)
{
    if (_currentPath.empty())
    {
        return;
    }

    const auto dest = ShowSaveAsDialog(hwnd);
    if (! dest.has_value())
    {
        return;
    }

    const UINT encodingSelection = dest.value().encodingSelection;
    if (encodingSelection == IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL)
    {
        if (! _fileSystem)
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }

        wil::unique_handle outFile(
            CreateFileW(dest.value().path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! outFile)
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }

        wil::com_ptr<IFileSystemIO> fileIo;
        if (FAILED(_fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void())) || ! fileIo)
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }

        wil::com_ptr<IFileReader> reader;
        const HRESULT openHr = fileIo->CreateFileReader(_currentPath.c_str(), reader.put());
        if (FAILED(openHr) || ! reader)
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }

        uint64_t ignored     = 0;
        const HRESULT seekHr = reader->Seek(0, FILE_BEGIN, &ignored);
        if (FAILED(seekHr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }

        std::vector<uint8_t> buffer(256u * 1024u);
        for (;;)
        {
            unsigned long read   = 0;
            const HRESULT readHr = reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &read);
            if (FAILED(readHr))
            {
                ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
                return;
            }

            if (read == 0)
            {
                break;
            }

            DWORD written = 0;
            if (WriteFile(outFile.get(), buffer.data(), read, &written, nullptr) == 0 || written != read)
            {
                ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
                return;
            }
        }
        return;
    }

    if (! _hEdit)
    {
        ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
        return;
    }

    wil::unique_handle outFile(CreateFileW(dest.value().path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! outFile)
    {
        ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
        return;
    }

    struct SaveEncoding
    {
        enum class Kind : uint8_t
        {
            CodePage,
            Utf16LE,
            Utf16BE,
            Utf32LE,
            Utf32BE,
        };

        Kind kind     = Kind::CodePage;
        UINT codePage = CP_UTF8;
        bool writeBom = false;
    };

    auto resolveSaveEncoding = [&](UINT selection) noexcept -> SaveEncoding
    {
        if (selection == IDM_VIEWER_ENCODING_SAVE_UTF16LE_BOM || selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::Utf16LE};
        }
        if (selection == IDM_VIEWER_ENCODING_SAVE_UTF16BE_BOM || selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::Utf16BE};
        }
        if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::Utf32LE};
        }
        if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::Utf32BE};
        }
        if (selection == IDM_VIEWER_ENCODING_SAVE_UTF8_BOM || selection == IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::CodePage, .codePage = CP_UTF8, .writeBom = true};
        }
        if (selection == IDM_VIEWER_ENCODING_SAVE_UTF8 || selection == IDM_VIEWER_ENCODING_DISPLAY_UTF8)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::CodePage, .codePage = CP_UTF8, .writeBom = false};
        }
        if (selection == IDM_VIEWER_ENCODING_SAVE_ANSI || selection == IDM_VIEWER_ENCODING_DISPLAY_ANSI)
        {
            return SaveEncoding{.kind = SaveEncoding::Kind::CodePage, .codePage = CP_ACP, .writeBom = false};
        }

        const UINT codePage = CodePageForMenuSelection(selection);
        return SaveEncoding{.kind = SaveEncoding::Kind::CodePage, .codePage = codePage, .writeBom = false};
    };

    const SaveEncoding saveEncoding = resolveSaveEncoding(encodingSelection);

    if (saveEncoding.kind == SaveEncoding::Kind::CodePage && saveEncoding.writeBom && saveEncoding.codePage == CP_UTF8)
    {
        static constexpr uint8_t kBom[] = {0xEFu, 0xBBu, 0xBFu};
        const HRESULT hr                = WriteAllHandle(outFile.get(), kBom, sizeof(kBom));
        if (FAILED(hr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }
    }
    else if (saveEncoding.kind == SaveEncoding::Kind::Utf16LE)
    {
        static constexpr uint8_t kBom[] = {0xFFu, 0xFEu};
        const HRESULT hr                = WriteAllHandle(outFile.get(), kBom, sizeof(kBom));
        if (FAILED(hr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }
    }
    else if (saveEncoding.kind == SaveEncoding::Kind::Utf16BE)
    {
        static constexpr uint8_t kBom[] = {0xFEu, 0xFFu};
        const HRESULT hr                = WriteAllHandle(outFile.get(), kBom, sizeof(kBom));
        if (FAILED(hr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }
    }
    else if (saveEncoding.kind == SaveEncoding::Kind::Utf32LE)
    {
        static constexpr uint8_t kBom[] = {0xFFu, 0xFEu, 0x00u, 0x00u};
        const HRESULT hr                = WriteAllHandle(outFile.get(), kBom, sizeof(kBom));
        if (FAILED(hr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }
    }
    else if (saveEncoding.kind == SaveEncoding::Kind::Utf32BE)
    {
        static constexpr uint8_t kBom[] = {0x00u, 0x00u, 0xFEu, 0xFFu};
        const HRESULT hr                = WriteAllHandle(outFile.get(), kBom, sizeof(kBom));
        if (FAILED(hr))
        {
            ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
            return;
        }
    }

    struct SaveCookie
    {
        SaveCookie()                             = default;
        SaveCookie(const SaveCookie&)            = delete;
        SaveCookie& operator=(const SaveCookie&) = delete;
        SaveCookie(SaveCookie&&)                 = delete;
        SaveCookie& operator=(SaveCookie&&)      = delete;
        ~SaveCookie()                            = default;

        HANDLE file = nullptr;
        SaveEncoding encoding{};
        HRESULT error = S_OK;
        std::optional<wchar_t> pendingHighSurrogate;
        std::vector<wchar_t> wideScratch;
        std::vector<uint8_t> byteScratch;

        static bool IsHighSurrogate(wchar_t ch) noexcept
        {
            return ch >= 0xD800u && ch <= 0xDBFFu;
        }

        static bool IsLowSurrogate(wchar_t ch) noexcept
        {
            return ch >= 0xDC00u && ch <= 0xDFFFu;
        }

        static HRESULT WriteChunk(SaveCookie& cookie, const wchar_t* data, size_t count) noexcept
        {
            if (! data || count == 0)
            {
                return S_OK;
            }

            if (cookie.encoding.kind == SaveEncoding::Kind::Utf16LE)
            {
                return WriteAllHandle(cookie.file, data, count * sizeof(wchar_t));
            }

            if (cookie.encoding.kind == SaveEncoding::Kind::Utf16BE)
            {
                cookie.byteScratch.resize(count * sizeof(wchar_t));
                for (size_t i = 0; i < count; ++i)
                {
                    const uint16_t value            = static_cast<uint16_t>(data[i]);
                    cookie.byteScratch[i * 2u + 0u] = static_cast<uint8_t>((value >> 8u) & 0xFFu);
                    cookie.byteScratch[i * 2u + 1u] = static_cast<uint8_t>(value & 0xFFu);
                }

                return WriteAllHandle(cookie.file, cookie.byteScratch.data(), cookie.byteScratch.size());
            }

            if (cookie.encoding.kind == SaveEncoding::Kind::Utf32LE || cookie.encoding.kind == SaveEncoding::Kind::Utf32BE)
            {
                cookie.byteScratch.clear();
                cookie.byteScratch.reserve(count * 4u);

                for (size_t i = 0; i < count; ++i)
                {
                    const wchar_t ch = data[i];

                    uint32_t cp = 0;
                    if (IsHighSurrogate(ch))
                    {
                        if ((i + 1) < count && IsLowSurrogate(data[i + 1]))
                        {
                            const uint32_t hi = static_cast<uint32_t>(ch) - 0xD800u;
                            const uint32_t lo = static_cast<uint32_t>(data[i + 1]) - 0xDC00u;
                            cp                = 0x10000u + ((hi << 10u) | lo);
                            i += 1;
                        }
                        else
                        {
                            cp = 0xFFFDu;
                        }
                    }
                    else if (IsLowSurrogate(ch))
                    {
                        cp = 0xFFFDu;
                    }
                    else
                    {
                        cp = static_cast<uint32_t>(ch);
                    }

                    if (cookie.encoding.kind == SaveEncoding::Kind::Utf32LE)
                    {
                        cookie.byteScratch.push_back(static_cast<uint8_t>(cp & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 8u) & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 16u) & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 24u) & 0xFFu));
                    }
                    else
                    {
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 24u) & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 16u) & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>((cp >> 8u) & 0xFFu));
                        cookie.byteScratch.push_back(static_cast<uint8_t>(cp & 0xFFu));
                    }
                }

                return WriteAllHandle(cookie.file, cookie.byteScratch.data(), cookie.byteScratch.size());
            }

            const UINT codePage = cookie.encoding.codePage;

            if (count > static_cast<size_t>(std::numeric_limits<int>::max()))
            {
                count = static_cast<size_t>(std::numeric_limits<int>::max());
            }

            const int srcLen   = static_cast<int>(count);
            const int required = WideCharToMultiByte(codePage, 0, data, srcLen, nullptr, 0, nullptr, nullptr);
            if (required <= 0)
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }

            cookie.byteScratch.resize(static_cast<size_t>(required));
            const int written = WideCharToMultiByte(codePage, 0, data, srcLen, reinterpret_cast<LPSTR>(cookie.byteScratch.data()), required, nullptr, nullptr);
            if (written <= 0)
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }

            return WriteAllHandle(cookie.file, cookie.byteScratch.data(), static_cast<size_t>(written));
        }

        static DWORD CALLBACK StreamOutCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG* pcb) noexcept
        {
            auto* cookie = reinterpret_cast<SaveCookie*>(dwCookie);
            if (! cookie || ! pbBuff || cb <= 0 || ! pcb)
            {
                return 1;
            }

            const size_t byteCount = static_cast<size_t>(cb);
            const size_t wideBytes = (byteCount / sizeof(wchar_t)) * sizeof(wchar_t);
            const size_t wideCount = wideBytes / sizeof(wchar_t);

            *pcb = static_cast<LONG>(wideBytes);
            if (wideCount == 0)
            {
                return 0;
            }

            const wchar_t* wide = reinterpret_cast<const wchar_t*>(pbBuff);

            cookie->wideScratch.clear();
            cookie->wideScratch.reserve(wideCount + 1);

            if (cookie->pendingHighSurrogate.has_value())
            {
                cookie->wideScratch.push_back(cookie->pendingHighSurrogate.value());
                cookie->pendingHighSurrogate.reset();
            }

            cookie->wideScratch.insert(cookie->wideScratch.end(), wide, wide + wideCount);

            if (! cookie->wideScratch.empty() && IsHighSurrogate(cookie->wideScratch.back()))
            {
                cookie->pendingHighSurrogate = cookie->wideScratch.back();
                cookie->wideScratch.pop_back();
            }

            const HRESULT hr = WriteChunk(*cookie, cookie->wideScratch.data(), cookie->wideScratch.size());
            if (FAILED(hr))
            {
                cookie->error = hr;
                return 1;
            }

            return 0;
        }
    };

    if (_textBuffer.empty() && _fileSize > _textStreamSkipBytes)
    {
        ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
        return;
    }

    SaveCookie cookie;
    cookie.file     = outFile.get();
    cookie.encoding = saveEncoding;

    cookie.error = SaveCookie::WriteChunk(cookie, _textBuffer.data(), _textBuffer.size());

    if (cookie.pendingHighSurrogate.has_value() && SUCCEEDED(cookie.error))
    {
        static constexpr wchar_t kReplacement = static_cast<wchar_t>(0xFFFDu);
        cookie.error                          = SaveCookie::WriteChunk(cookie, &kReplacement, 1);
        cookie.pendingHighSurrogate.reset();
    }

    if (FAILED(cookie.error))
    {
        ShowInlineAlert(InlineAlertSeverity::Error, IDS_VIEWERTEXT_CAPTION_ERROR, IDS_VIEWERTEXT_ERR_SAVE_FAILED);
        return;
    }

    if (_textStreamActive)
    {
        ShowInlineAlert(InlineAlertSeverity::Info, IDS_VIEWERTEXT_NAME, IDS_VIEWERTEXT_MSG_STREAM_TRUNCATED);
    }
}

void ViewerText::CommandRefresh(HWND hwnd)
{
    if (_currentPath.empty())
    {
        return;
    }

    static_cast<void>(OpenPath(hwnd, _currentPath, false));
}

void ViewerText::CommandOtherNext(HWND hwnd)
{
    if (_otherFiles.size() <= 1)
    {
        return;
    }

    _otherIndex = (_otherIndex + 1) % _otherFiles.size();
    static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
}

void ViewerText::CommandOtherPrevious(HWND hwnd)
{
    if (_otherFiles.size() <= 1)
    {
        return;
    }

    if (_otherIndex == 0)
    {
        _otherIndex = _otherFiles.size() - 1;
    }
    else
    {
        _otherIndex -= 1;
    }

    static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
}

void ViewerText::CommandOtherFirst(HWND hwnd)
{
    if (_otherFiles.empty())
    {
        return;
    }

    _otherIndex = 0;
    static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
}

void ViewerText::CommandOtherLast(HWND hwnd)
{
    if (_otherFiles.empty())
    {
        return;
    }

    _otherIndex = _otherFiles.size() - 1;
    static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
}

void ViewerText::CommandFind(HWND hwnd)
{
    FindDialogState state;
    state.viewer  = this;
    state.initial = _searchQuery;

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
    const INT_PTR res = DialogBoxParamW(g_hInstance, MAKEINTRESOURCEW(IDD_VIEWERTEXT_FIND), hwnd, FindDlgProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)
    if (res != IDOK)
    {
        return;
    }

    _searchQuery = state.result;
    UpdateSearchHighlights();
    if (_searchQuery.empty())
    {
        return;
    }

    CommandFindNext(hwnd, false);
}

void ViewerText::UpdateSearchHighlights() noexcept
{
    _searchMatchStarts.clear();

    _hexSearchNeedle.clear();
    _hexSearchNeedleValid = false;
    if (! _searchQuery.empty())
    {
        std::vector<uint8_t> needle;
        if (TryParseHexSearchNeedle(_searchQuery, needle))
        {
            if (! HexBigEndian())
            {
                std::reverse(needle.begin(), needle.end());
            }

            _hexSearchNeedle      = std::move(needle);
            _hexSearchNeedleValid = ! _hexSearchNeedle.empty();
        }
    }

    if (! _searchQuery.empty() && ! _textBuffer.empty() && _searchQuery.size() <= _textBuffer.size())
    {
        const size_t queryLen = _searchQuery.size();

        size_t pos = 0;
        while (pos < _textBuffer.size())
        {
            const size_t found = _textBuffer.find(_searchQuery, pos);
            if (found == std::wstring::npos)
            {
                break;
            }

            _searchMatchStarts.push_back(found);
            pos = found + queryLen;
        }
    }

    if (_hEdit)
    {
        InvalidateRect(_hEdit.get(), nullptr, TRUE);
    }
    if (_hHex)
    {
        InvalidateRect(_hHex.get(), nullptr, TRUE);
    }
}

void ViewerText::CommandGoToOffset(HWND hwnd)
{
    GoToDialogState state;
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
    const INT_PTR res = DialogBoxParamW(g_hInstance, MAKEINTRESOURCEW(IDD_VIEWERTEXT_GOTO), hwnd, GoToDlgProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)
    if (res != IDOK)
    {
        return;
    }

    if (! state.offset.has_value())
    {
        return;
    }

    CommandGoToOffsetValue(hwnd, state.offset.value());
}

void ViewerText::CommandGoToTop(HWND hwnd, bool extendSelection) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (_viewMode == ViewMode::Hex)
    {
        if (! _hHex || _fileSize == 0)
        {
            return;
        }

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

        const uint64_t nextOffset = 0;
        if (extendSelection)
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
        _hexTopLine        = 0;
        UpdateHexViewScrollBars(_hHex.get());
        InvalidateRect(_hHex.get(), nullptr, TRUE);
        InvalidateRect(hwnd, &_statusRect, FALSE);
        return;
    }

    if (! _hEdit)
    {
        return;
    }

    if (_textStreamActive && _textStreamStartOffset > _textStreamSkipBytes)
    {
        static_cast<void>(LoadTextToEdit(hwnd, _textStreamSkipBytes, false));
        return;
    }

    _textTopVisualLine = 0;
    _textLeftColumn    = 0;

    const size_t newCaret = 0;
    _textCaretIndex       = newCaret;
    if (! extendSelection)
    {
        _textSelAnchor = newCaret;
    }
    _textSelActive       = newCaret;
    _textPreferredColumn = 0;

    UpdateTextViewScrollBars(_hEdit.get());
    InvalidateRect(_hEdit.get(), nullptr, TRUE);
    InvalidateRect(hwnd, &_statusRect, FALSE);
}

void ViewerText::CommandGoToBottom(HWND hwnd, bool extendSelection) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (_viewMode == ViewMode::Hex)
    {
        if (! _hHex || _fileSize == 0)
        {
            return;
        }

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

        const uint64_t nextOffset = _fileSize - 1;
        if (extendSelection)
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

        const uint64_t targetLine = nextOffset / static_cast<uint64_t>(kHexBytesPerLine);

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
        InvalidateRect(hwnd, &_statusRect, FALSE);
        return;
    }

    if (! _hEdit)
    {
        return;
    }

    if (_textStreamActive && _fileSize > 0 && _textStreamEndOffset < _fileSize)
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

    if (_textVisualLineStarts.empty())
    {
        return;
    }

    RECT client{};
    GetClientRect(_hEdit.get(), &client);
    const UINT dpi        = GetDpiForWindow(_hEdit.get());
    const float heightDip = std::max(1.0f, DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi));
    const float marginDip = 6.0f;
    const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
    const float usableDip = std::max(0.0f, heightDip - 2.0f * marginDip);
    const uint32_t rows   = std::max<uint32_t>(1u, static_cast<uint32_t>(std::floor(usableDip / std::max(1.0f, lineH))));

    const uint32_t totalVisual = static_cast<uint32_t>(_textVisualLineStarts.size());
    const uint32_t lastVisual  = totalVisual > 0 ? (totalVisual - 1) : 0;
    const uint32_t desiredTop  = (totalVisual > rows) ? (totalVisual - rows) : 0;

    _textTopVisualLine = std::min<uint32_t>(desiredTop, lastVisual);

    const size_t newCaret = _textBuffer.size();
    _textCaretIndex       = newCaret;
    if (! extendSelection)
    {
        _textSelAnchor = newCaret;
    }
    _textSelActive       = newCaret;
    _textPreferredColumn = 0;

    UpdateTextViewScrollBars(_hEdit.get());
    InvalidateRect(_hEdit.get(), nullptr, TRUE);
    InvalidateRect(hwnd, &_statusRect, FALSE);
}

HRESULT ViewerText::DetectEncodingAndSize(const std::filesystem::path& path, FileEncoding& encoding, uint64_t& bomBytes, uint64_t& fileSize) noexcept
{
    encoding = FileEncoding::Unknown;
    bomBytes = 0;
    fileSize = 0;

    if (! _fileReader)
    {
        Debug::Error(L"ViewerText: DetectEncodingAndSize failed because file reader is missing for '{}'.", path.c_str());
        return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
    }

    uint64_t sizeBytes   = 0;
    const HRESULT sizeHr = _fileReader->GetSize(&sizeBytes);
    if (FAILED(sizeHr))
    {
        Debug::Error(L"ViewerText: GetSize failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(sizeHr));
        return sizeHr;
    }

    fileSize = sizeBytes;

    BYTE bom[4]{};
    unsigned long read = 0;

    uint64_t pos         = 0;
    const HRESULT seekHr = _fileReader->Seek(0, FILE_BEGIN, &pos);
    if (FAILED(seekHr))
    {
        Debug::Error(L"ViewerText: Seek(FILE_BEGIN, 0) failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(seekHr));
        return seekHr;
    }

    const HRESULT readHr = _fileReader->Read(bom, static_cast<unsigned long>(std::size(bom)), &read);
    if (FAILED(readHr))
    {
        Debug::Error(L"ViewerText: Read failed for '{}' (hr=0x{:08X}).", path.c_str(), static_cast<unsigned long>(readHr));
        return readHr;
    }

    if (read >= 4 && bom[0] == 0xFF && bom[1] == 0xFE && bom[2] == 0x00 && bom[3] == 0x00)
    {
        encoding = FileEncoding::Utf32LE;
        bomBytes = 4;
        return S_OK;
    }
    if (read >= 4 && bom[0] == 0x00 && bom[1] == 0x00 && bom[2] == 0xFE && bom[3] == 0xFF)
    {
        encoding = FileEncoding::Utf32BE;
        bomBytes = 4;
        return S_OK;
    }
    if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
    {
        encoding = FileEncoding::Utf8;
        bomBytes = 3;
        return S_OK;
    }
    if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
    {
        encoding = FileEncoding::Utf16LE;
        bomBytes = 2;
        return S_OK;
    }
    if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
    {
        encoding = FileEncoding::Utf16BE;
        bomBytes = 2;
        return S_OK;
    }

    encoding = FileEncoding::Unknown;
    bomBytes = 0;
    return S_OK;
}

std::wstring ViewerText::EncodingLabel() const
{
    UINT id = IDS_VIEWERTEXT_ENCODING_UNKNOWN;
    switch (_encoding)
    {
        case FileEncoding::Utf8: id = IDS_VIEWERTEXT_ENCODING_UTF8; break;
        case FileEncoding::Utf16LE: id = IDS_VIEWERTEXT_ENCODING_UTF16LE; break;
        case FileEncoding::Utf16BE: id = IDS_VIEWERTEXT_ENCODING_UTF16BE; break;
        case FileEncoding::Utf32LE: id = IDS_VIEWERTEXT_ENCODING_UTF32LE; break;
        case FileEncoding::Utf32BE: id = IDS_VIEWERTEXT_ENCODING_UTF32BE; break;
        case FileEncoding::Unknown: id = IDS_VIEWERTEXT_ENCODING_UNKNOWN; break;
        default: id = IDS_VIEWERTEXT_ENCODING_UNKNOWN; break;
    }

    return LoadStringResource(g_hInstance, id);
}

std::wstring ViewerText::BuildStatusText() const
{
    auto withStatusMessage = [&](std::wstring base) -> std::wstring
    {
        std::wstring combined = std::move(base);

        if (_viewMode == ViewMode::Text && _textStreamActive && ! _isLoading)
        {
            const std::wstring streamingMessage = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_MSG_STREAM_TRUNCATED);
            if (! streamingMessage.empty())
            {
                const std::wstring streamingCombined = FormatStringResource(g_hInstance, IDS_VIEWERTEXT_STATUS_WITH_MESSAGE_FORMAT, streamingMessage, combined);
                if (! streamingCombined.empty())
                {
                    combined = streamingCombined;
                }
            }
        }

        if (! _statusMessage.empty())
        {
            const std::wstring statusCombined = FormatStringResource(g_hInstance, IDS_VIEWERTEXT_STATUS_WITH_MESSAGE_FORMAT, _statusMessage, combined);
            if (! statusCombined.empty())
            {
                combined = statusCombined;
            }
        }

        return combined;
    };

    std::wstring detected;
    if (_encoding != FileEncoding::Unknown)
    {
        detected = EncodingLabel();

        if (_bomBytes > 0)
        {
            detected.append(LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DETECTED_SUFFIX_BOM));
        }
    }
    else if (_detectedCodePageValid)
    {
        if (_detectedCodePage == CP_UTF8)
        {
            detected = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_ENCODING_UTF8);
        }
        else
        {
            detected = FormatStringResource(g_hInstance, IDS_VIEWERTEXT_CODEPAGE_FORMAT, _detectedCodePage);
        }

        if (_detectedCodePageIsGuess)
        {
            detected.append(LoadStringResource(g_hInstance, IDS_VIEWERTEXT_DETECTED_SUFFIX_GUESS));
        }
    }
    else
    {
        detected = EncodingLabel();
    }

    auto stripMenuText = [](std::wstring_view text) -> std::wstring
    {
        const size_t tabPos = text.find(L'\t');
        if (tabPos != std::wstring_view::npos)
        {
            text = text.substr(0, tabPos);
        }

        std::wstring result;
        result.reserve(text.size());

        for (size_t i = 0; i < text.size(); ++i)
        {
            const wchar_t ch = text[i];
            if (ch != L'&')
            {
                result.push_back(ch);
                continue;
            }

            if ((i + 1) < text.size() && text[i + 1] == L'&')
            {
                result.push_back(L'&');
                i += 1;
            }
        }

        while (! result.empty() && result.front() == L' ')
        {
            result.erase(result.begin());
        }
        while (! result.empty() && result.back() == L' ')
        {
            result.pop_back();
        }

        return result;
    };

    const UINT selection = EffectiveDisplayEncodingMenuSelection();
    std::wstring active;
    if (_hWnd)
    {
        HMENU menu = GetMenu(_hWnd.get());
        if (menu)
        {
            wchar_t buffer[256]{};
            const int len = GetMenuStringW(menu, selection, buffer, static_cast<int>(std::size(buffer)), MF_BYCOMMAND);
            if (len > 0)
            {
                active.assign(buffer, static_cast<size_t>(len));
            }
        }
    }

    active = stripMenuText(active);

    const std::wstring sizeText = FormatBytesCompact(_fileSize);

    if (_viewMode == ViewMode::Hex)
    {
        uint64_t topOffset    = 0;
        uint64_t bottomOffset = 0;

        if (_fileSize > 0 && _hHex)
        {
            const uint64_t maxByte  = _fileSize - 1;
            const uint64_t topStart = _hexTopLine * static_cast<uint64_t>(kHexBytesPerLine);

            RECT client{};
            GetClientRect(_hHex.get(), &client);
            const UINT dpi        = GetDpiForWindow(_hHex.get());
            const float heightDip = std::max(1.0f, DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi));
            const float marginDip = 6.0f;
            const float lineH     = (_hexLineHeightDip > 0.0f) ? _hexLineHeightDip : 14.0f;
            const float headerH   = lineH;
            const float usableDip = std::max(0.0f, heightDip - headerH - 2.0f * marginDip);
            const uint32_t rows   = std::max<uint32_t>(1u, static_cast<uint32_t>(std::ceil(usableDip / std::max(1.0f, lineH))));

            const uint64_t bottomLine  = (_hexTopLine + static_cast<uint64_t>(rows) > 0) ? (_hexTopLine + static_cast<uint64_t>(rows) - 1) : 0;
            const uint64_t bottomStart = bottomLine * static_cast<uint64_t>(kHexBytesPerLine);

            topOffset    = std::min(topStart, maxByte);
            bottomOffset = std::min(maxByte, std::min(bottomStart, maxByte) + static_cast<uint64_t>(kHexBytesPerLine - 1));
        }

        return withStatusMessage(FormatStringResource(g_hInstance,
                                                      IDS_VIEWERTEXT_STATUS_HEX_FORMAT,
                                                      _fileSystemName,
                                                      detected,
                                                      active,
                                                      sizeText,
                                                      FormatFileOffset(topOffset),
                                                      FormatFileOffset(bottomOffset)));
    }

    int topLine    = 1;
    int bottomLine = 1;

    if (_hEdit && ! _textVisualLineStarts.empty() && ! _textVisualLineLogical.empty())
    {
        RECT client{};
        GetClientRect(_hEdit.get(), &client);
        const UINT dpi        = GetDpiForWindow(_hEdit.get());
        const float heightDip = std::max(1.0f, DipsFromPixels(static_cast<int>(client.bottom - client.top), dpi));
        const float marginDip = 6.0f;
        const float usableDip = std::max(0.0f, heightDip - 2.0f * marginDip);
        const float lineH     = (_textLineHeightDip > 0.0f) ? _textLineHeightDip : 14.0f;
        const uint32_t rows   = std::max<uint32_t>(1u, static_cast<uint32_t>(std::ceil(usableDip / std::max(1.0f, lineH))));

        const uint32_t totalVisual  = static_cast<uint32_t>(_textVisualLineStarts.size());
        const uint32_t topVisual    = std::min<uint32_t>(_textTopVisualLine, totalVisual - 1);
        const uint32_t bottomVisual = std::min<uint32_t>(totalVisual - 1, topVisual + rows - 1);

        const uint32_t topLogical    = std::min<uint32_t>(_textVisualLineLogical[topVisual], static_cast<uint32_t>(_textLineStarts.size() - 1));
        const uint32_t bottomLogical = std::min<uint32_t>(_textVisualLineLogical[bottomVisual], static_cast<uint32_t>(_textLineStarts.size() - 1));

        topLine    = static_cast<int>(topLogical) + 1;
        bottomLine = static_cast<int>(bottomLogical) + 1;
    }

    std::wstring totalLinesText = LoadStringResource(g_hInstance, IDS_VIEWERTEXT_UNKNOWN);
    if (! _isLoading)
    {
        if (_textTotalLineCount.has_value())
        {
            totalLinesText = std::format(L"{:L}", _textTotalLineCount.value());
        }
        else if (! _textStreamActive && ! _textLineStarts.empty())
        {
            const uint64_t totalLines = static_cast<uint64_t>(_textLineStarts.size());
            totalLinesText            = std::format(L"{:L}", totalLines);
        }
    }

    return withStatusMessage(
        FormatStringResource(g_hInstance, IDS_VIEWERTEXT_STATUS_TEXT_FORMAT, _fileSystemName, detected, active, sizeText, topLine, bottomLine, totalLinesText));
}

HRESULT ViewerText::OpenPath(HWND hwnd, const std::filesystem::path& path, bool updateOtherFiles) noexcept
{
    if (path.empty())
    {
        Debug::Error(L"ViewerText: OpenPath called with an empty path.");
        return E_INVALIDARG;
    }

    StartAsyncOpen(hwnd, path, updateOtherFiles, 0);
    return S_OK;
}

bool ViewerText::HandleShortcutKey(HWND hwnd, WPARAM vk) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    const bool ctrl  = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

    if (vk == VK_ESCAPE)
    {
        CommandExit(hwnd);
        return true;
    }

    if (vk == VK_SPACE)
    {
        CommandOtherNext(hwnd);
        return true;
    }

    if (vk == VK_BACK)
    {
        CommandOtherPrevious(hwnd);
        return true;
    }

    if (ctrl && vk == VK_RIGHT)
    {
        return false;
    }

    if (ctrl && vk == VK_LEFT)
    {
        return false;
    }

    if (ctrl && vk == VK_UP)
    {
        CommandOtherPrevious(hwnd);
        return true;
    }

    if (ctrl && vk == VK_DOWN)
    {
        CommandOtherNext(hwnd);
        return true;
    }

    if (ctrl && vk == VK_HOME)
    {
        CommandOtherFirst(hwnd);
        return true;
    }

    if (ctrl && vk == VK_END)
    {
        CommandOtherLast(hwnd);
        return true;
    }

    if (ctrl && (vk == 'F' || vk == 'f'))
    {
        CommandFind(hwnd);
        return true;
    }

    if (vk == VK_F3)
    {
        CommandFindNext(hwnd, shift);
        return true;
    }

    if (ctrl && (vk == 'G' || vk == 'g'))
    {
        CommandGoToOffset(hwnd);
        return true;
    }

    if (ctrl && (vk == 'O' || vk == 'o'))
    {
        CommandOpen(hwnd);
        return true;
    }

    if (ctrl && (vk == 'S' || vk == 's'))
    {
        CommandSaveAs(hwnd);
        return true;
    }

    if (vk == VK_F5)
    {
        CommandRefresh(hwnd);
        return true;
    }

    if (vk == VK_F8)
    {
        CommandCycleDisplayEncoding(hwnd, shift);
        return true;
    }

    return false;
}

HRESULT STDMETHODCALLTYPE ViewerText::Open(const ViewerOpenContext* context) noexcept
{
    if (! context || ! context->focusedPath || context->focusedPath[0] == L'\0')
    {
        Debug::Error(L"ViewerText: Open called with an invalid context (focusedPath missing).");
        return E_INVALIDARG;
    }

    if (! context->fileSystem)
    {
        Debug::Error(L"ViewerText: Open called with an invalid context (fileSystem missing).");
        return E_INVALIDARG;
    }

    _fileSystem = context->fileSystem;

    _fileSystemName.clear();
    if (context->fileSystemName && context->fileSystemName[0] != L'\0')
    {
        _fileSystemName = context->fileSystemName;
    }

    _selection.clear();
    if (context->selectionPaths && context->selectionCount > 0)
    {
        for (unsigned long i = 0; i < context->selectionCount; ++i)
        {
            const wchar_t* p = context->selectionPaths[i];
            if (p && p[0] != L'\0')
            {
                _selection.emplace_back(p);
            }
        }
    }

    _otherFiles.clear();
    if (context->otherFiles && context->otherFileCount > 0)
    {
        for (unsigned long i = 0; i < context->otherFileCount; ++i)
        {
            const wchar_t* p = context->otherFiles[i];
            if (p && p[0] != L'\0')
            {
                _otherFiles.emplace_back(p);
            }
        }
    }

    _otherIndex = 0;
    if (! _otherFiles.empty() && context->focusedOtherFileIndex < _otherFiles.size())
    {
        _otherIndex = static_cast<size_t>(context->focusedOtherFileIndex);
    }

    if ((context->flags & VIEWER_OPEN_FLAG_START_HEX) != 0)
    {
        _viewMode = ViewMode::Hex;
    }

    const std::filesystem::path path(context->focusedPath);

    if (! _hWnd)
    {
        if (! RegisterWndClass(g_hInstance))
        {
            return E_FAIL;
        }

        HWND ownerWindow = context->ownerWindow;

        RECT ownerRect{};
        if (ownerWindow && GetWindowRect(ownerWindow, &ownerRect) != 0)
        {
            const int w = ownerRect.right - ownerRect.left;
            const int h = ownerRect.bottom - ownerRect.top;

            wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(LoadMenuW(g_hInstance, MAKEINTRESOURCEW(IDR_VIEWERTEXT_MENU)));
            HWND window = CreateWindowExW(0,
                                          kClassName,
                                          L"",
                                          WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
                                          ownerRect.left,
                                          ownerRect.top,
                                          std::max(1, w),
                                          std::max(1, h),
                                          nullptr,
                                          menu.get(),
                                          g_hInstance,
                                          this);
            if (! window)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerText: CreateWindowExW failed.");
                return HRESULT_FROM_WIN32(lastError);
            }

            menu.release();

            _hWnd.reset(window);

            if (! _windowIconSmall)
            {
                _windowIconSmall = CreateViewerTextIcon(16);
            }
            if (! _windowIconBig)
            {
                _windowIconBig = CreateViewerTextIcon(32);
            }
            if (_windowIconSmall)
            {
                static_cast<void>(SendMessageW(_hWnd.get(), WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(_windowIconSmall.get())));
            }
            if (_windowIconBig)
            {
                static_cast<void>(SendMessageW(_hWnd.get(), WM_SETICON, ICON_BIG, reinterpret_cast<LPARAM>(_windowIconBig.get())));
            }

            ApplyTheme(_hWnd.get());
            ApplyPendingViewerTextClassBackgroundBrush(_hWnd.get(), _hEdit.get(), _hHex.get());

            AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
            ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
        else
        {
            wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(LoadMenuW(g_hInstance, MAKEINTRESOURCEW(IDR_VIEWERTEXT_MENU)));
            HWND window = CreateWindowExW(
                0, kClassName, L"", WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN, CW_USEDEFAULT, CW_USEDEFAULT, 900, 700, nullptr, menu.get(), g_hInstance, this);
            if (! window)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerText: CreateWindowExW failed.");
                return HRESULT_FROM_WIN32(lastError);
            }

            menu.release();
            _hWnd.reset(window);

            if (! _windowIconSmall)
            {
                _windowIconSmall = CreateViewerTextIcon(16);
            }
            if (! _windowIconBig)
            {
                _windowIconBig = CreateViewerTextIcon(32);
            }
            if (_windowIconSmall)
            {
                static_cast<void>(SendMessageW(_hWnd.get(), WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(_windowIconSmall.get())));
            }
            if (_windowIconBig)
            {
                static_cast<void>(SendMessageW(_hWnd.get(), WM_SETICON, ICON_BIG, reinterpret_cast<LPARAM>(_windowIconBig.get())));
            }

            ApplyTheme(_hWnd.get());
            ApplyPendingViewerTextClassBackgroundBrush(_hWnd.get(), _hEdit.get(), _hHex.get());

            AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
            ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }
    else
    {
        ApplyPendingViewerTextClassBackgroundBrush(_hWnd.get(), _hEdit.get(), _hHex.get());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        static_cast<void>(SetForegroundWindow(_hWnd.get()));
    }

    if (! _hWnd)
    {
        Debug::Error(L"ViewerText: Open failed because viewer window is missing after creation.");
        return E_FAIL;
    }

    StartAsyncOpen(_hWnd.get(), path, false, 0);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::Close() noexcept
{
    _hWnd.reset();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::SetTheme(const ViewerTheme* theme) noexcept
{
    if (! theme || theme->version != 2)
    {
        return E_INVALIDARG;
    }

    _theme    = *theme;
    _hasTheme = true;

    RequestViewerTextClassBackgroundColor(ColorRefFromArgb(_theme.backgroundArgb));
    ApplyPendingViewerTextClassBackgroundBrush(_hWnd.get(), _hEdit.get(), _hHex.get());

    if (_hWnd)
    {
        ApplyTheme(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
        if (_hEdit)
        {
            InvalidateRect(_hEdit.get(), nullptr, TRUE);
        }
        if (_hHex)
        {
            InvalidateRect(_hHex.get(), nullptr, TRUE);
        }
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerText::SetCallback(IViewerCallback* callback, void* cookie) noexcept
{
    _callback       = callback;
    _callbackCookie = cookie;
    return S_OK;
}
