#pragma once

#include <string_view>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027, C4820 (padding)
#pragma warning(disable : 4625 4626 5026 5027 4820)
#include <wil/resource.h>
#pragma warning(pop)

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

namespace FluentIcons
{
inline constexpr std::wstring_view kFontFamily = L"Segoe Fluent Icons";
inline constexpr int kDefaultSizeDip           = 15;

// Segoe Fluent Icons PUA glyphs (see https://learn.microsoft.com/windows/apps/design/style/segoe-fluent-icons-font)
inline constexpr wchar_t kChevronRight      = L'\uE76C';
inline constexpr wchar_t kChevronDown       = L'\uE70D';
inline constexpr wchar_t kChevronUp         = L'\uE70E';
inline constexpr wchar_t kChevronRightSmall = L'\uE970';
inline constexpr wchar_t kChevronDownSmall  = L'\uE96E';
inline constexpr wchar_t kChevronUpSmall    = L'\uE96D';
inline constexpr wchar_t kCheckMark         = L'\uE73E';
inline constexpr wchar_t kWarning           = L'\uE7BA';
inline constexpr wchar_t kError             = L'\uEA39';
inline constexpr wchar_t kSort              = L'\uE8CB';
inline constexpr wchar_t kSettings          = L'\uE713';
inline constexpr wchar_t kPuzzle            = L'\uEA86';
inline constexpr wchar_t kCopy              = L'\uE8C8';
inline constexpr wchar_t kPaste             = L'\uE77F';
inline constexpr wchar_t kCut               = L'\uE8C6';
inline constexpr wchar_t kDelete            = L'\uE74D';
inline constexpr wchar_t kRename            = L'\uE8AC';
inline constexpr wchar_t kOpenFile          = L'\uE8E5';
inline constexpr wchar_t kInfo              = L'\uE946';
inline constexpr wchar_t kCalendar          = L'\uE787';
inline constexpr wchar_t kHardDrive         = L'\uEDA2';
inline constexpr wchar_t kTag               = L'\uE8EC';
inline constexpr wchar_t kFont              = L'\uE8D2';
inline constexpr wchar_t kDocument          = L'\uE8A5';
inline constexpr wchar_t kClear             = L'\uE894';
inline constexpr wchar_t kMapDrive          = L'\uE8CE';
inline constexpr wchar_t kConnections       = L'\uED5C';
inline constexpr wchar_t kHistory           = L'\uE81C';
inline constexpr wchar_t kFind              = L'\uE721';
inline constexpr wchar_t kCommandPrompt     = L'\uE756';

// Fallback glyphs (standard Unicode) when Segoe Fluent Icons isn't installed.
inline constexpr wchar_t kFallbackChevronRight = L'\u203A'; // ›
inline constexpr wchar_t kFallbackChevronDown  = L'\u25BE'; // ▾
inline constexpr wchar_t kFallbackCheckMark    = L'\u2713'; // ✓
inline constexpr wchar_t kFallbackWarning      = L'\u26A0'; // ⚠
inline constexpr wchar_t kFallbackError        = L'\u2716'; // ✖
inline constexpr wchar_t kFallbackSort         = L'\u21C5'; // ⇅

[[nodiscard]] inline wil::unique_hfont CreateFontForDpi(UINT dpi, int sizeDip) noexcept
{
    const int heightPx = -MulDiv(sizeDip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);

    LOGFONTW lf{};
    lf.lfHeight         = heightPx;
    lf.lfWeight         = FW_NORMAL;
    lf.lfCharSet        = DEFAULT_CHARSET;
    lf.lfQuality        = CLEARTYPE_QUALITY;
    lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
    wcsncpy_s(lf.lfFaceName, kFontFamily.data(), _TRUNCATE);

    return wil::unique_hfont(CreateFontIndirectW(&lf));
}

[[nodiscard]] inline bool FontHasGlyph(HDC hdc, HFONT font, wchar_t ch) noexcept
{
    if (! hdc || ! font)
    {
        return false;
    }

    auto old = wil::SelectObject(hdc, font);

    WORD glyphIndex{};
    const wchar_t text[2]{ch, 0};
    if (GetGlyphIndicesW(hdc, text, 1, &glyphIndex, GGI_MARK_NONEXISTING_GLYPHS) == GDI_ERROR)
    {
        return false;
    }

    return glyphIndex != 0xFFFF;
}
} // namespace FluentIcons
