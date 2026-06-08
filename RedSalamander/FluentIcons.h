#pragma once

#include <string_view>

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
inline constexpr wchar_t kFolder            = L'\uE8B7';
inline constexpr wchar_t kInfo              = L'\uE946';
inline constexpr wchar_t kCalendar          = L'\uE787';
inline constexpr wchar_t kHardDrive         = L'\uEDA2';
inline constexpr wchar_t kTag               = L'\uE8EC';
inline constexpr wchar_t kFont              = L'\uE8D2';
inline constexpr wchar_t kDocument          = L'\uE8A5';
inline constexpr wchar_t kClear             = L'\uE894';
inline constexpr wchar_t kMapDrive          = L'\uE8CE';
inline constexpr wchar_t kSyncFolder        = L'\uE8F7';
inline constexpr wchar_t kPreview           = L'\uE8FF';
inline constexpr wchar_t kLedLight          = L'\uE781';
inline constexpr wchar_t kConnections       = L'\uED5C';
inline constexpr wchar_t kHistory           = L'\uE81C';
inline constexpr wchar_t kFind              = L'\uE721';
inline constexpr wchar_t kFilter            = L'\uE71C';
inline constexpr wchar_t kCommandPrompt     = L'\uE756';

// Fallback glyphs (standard Unicode) when Segoe Fluent Icons isn't installed.
inline constexpr wchar_t kFallbackChevronRight = L'\u203A'; // ›
inline constexpr wchar_t kFallbackChevronDown  = L'\u25BE'; // ▾
inline constexpr wchar_t kFallbackChevronUp    = L'\u25B4';
inline constexpr wchar_t kFallbackCheckMark    = L'\u2713'; // ✓
inline constexpr wchar_t kFallbackWarning      = L'\u26A0'; // ⚠
inline constexpr wchar_t kFallbackError        = L'\u2716'; // ✖
inline constexpr wchar_t kFallbackSort         = L'\u21C5'; // ⇅
} // namespace FluentIcons
