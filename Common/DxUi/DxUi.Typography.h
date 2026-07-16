#pragma once

#include <algorithm>
#include <atomic>
#include <cmath>
#include <cstdint>
#include <cwchar>
#include <limits>
#include <mutex>
#include <string>
#include <string_view>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <dwrite.h>
#include <windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#pragma warning(pop)

#include "DxUi.h"

namespace RedSalamander::DxUi::Typography
{
inline constexpr wchar_t kSegoeUiVariableSmallFamily[]   = L"Segoe UI Variable Small";
inline constexpr wchar_t kSegoeUiVariableTextFamily[]    = L"Segoe UI Variable Text";
inline constexpr wchar_t kSegoeUiVariableDisplayFamily[] = L"Segoe UI Variable Display";
inline constexpr wchar_t kSegoeUiFallbackFamily[]        = L"Segoe UI";
inline constexpr wchar_t kSegoeFluentIconsFamily[]       = L"Segoe Fluent Icons";
inline constexpr wchar_t kSegoeMdl2AssetsFamily[]        = L"Segoe MDL2 Assets";
inline constexpr wchar_t kSegoeUiEmojiFamily[]           = L"Segoe UI Emoji";
inline constexpr wchar_t kUiMonospaceFamily[]            = L"Consolas";
inline constexpr size_t kMaxDWriteFamilyNameLength       = 255u;

struct TypographySpec
{
    std::wstring_view familyName = kSegoeUiVariableTextFamily;
    DWRITE_FONT_WEIGHT weight    = DWRITE_FONT_WEIGHT_NORMAL;
    DWRITE_FONT_STYLE style      = DWRITE_FONT_STYLE_NORMAL;
    float sizeDip                = 13.0f;
};

struct TextPixelMetrics final
{
    int widthPx      = 0;
    int heightPx     = 0;
    int lineHeightPx = 0;
};

inline constexpr std::wstring_view kTypographyFamilyCacheMissMetric     = L"dxui.typography.family_cache_miss_count";
inline constexpr std::wstring_view kTypographyTextFormatCacheMissMetric = L"dxui.typography.text_format_cache_miss_count";

using TypographyPerfEmitter =
    void (*)(std::wstring_view metric, std::wstring_view detail, uint64_t durationUs, uint64_t value, uint64_t count, HRESULT hr) noexcept;

[[nodiscard]] inline std::atomic<TypographyPerfEmitter>& GetTypographyPerfEmitter() noexcept
{
    static std::atomic<TypographyPerfEmitter> emitter{nullptr};
    return emitter;
}

inline void SetTypographyPerfEmitter(TypographyPerfEmitter emitter) noexcept
{
    GetTypographyPerfEmitter().store(emitter, std::memory_order_release);
}

inline void EmitTypographyPerfCounter(
    std::wstring_view metric, std::wstring_view detail, uint64_t durationUs, uint64_t value, uint64_t count, HRESULT hr) noexcept
{
    if (TypographyPerfEmitter emitter = GetTypographyPerfEmitter().load(std::memory_order_acquire))
    {
        emitter(metric, detail, durationUs, value, count, hr);
    }
}

struct TypographyFontFamilyCacheEntry final
{
    wil::com_ptr<IDWriteFactory> factory;
    IDWriteFactory* factoryKey = nullptr;
    std::wstring familyName;
    bool available = false;
};

struct TypographyTextFormatCacheEntry final
{
    wil::com_ptr<IDWriteFactory> factory;
    IDWriteFactory* factoryKey = nullptr;
    FontRole role              = FontRole::Body;
    bool wrap                  = false;
    wil::com_ptr<IDWriteTextFormat> textFormat;
};

[[nodiscard]] inline std::wstring_view GetUiTextFamilyForSizeDip(float sizeDip) noexcept
{
    if (sizeDip <= 12.0f)
    {
        return kSegoeUiVariableSmallFamily;
    }
    if (sizeDip >= 32.0f)
    {
        return kSegoeUiVariableDisplayFamily;
    }
    return kSegoeUiVariableTextFamily;
}

[[nodiscard]] inline TypographySpec MakeUiTextSpec(float sizeDip,
                                                   DWRITE_FONT_WEIGHT weight = DWRITE_FONT_WEIGHT_NORMAL,
                                                   DWRITE_FONT_STYLE style   = DWRITE_FONT_STYLE_NORMAL) noexcept
{
    return TypographySpec{.familyName = GetUiTextFamilyForSizeDip(sizeDip), .weight = weight, .style = style, .sizeDip = sizeDip};
}

[[nodiscard]] inline TypographySpec MakeUiIconSpec(float sizeDip, DWRITE_FONT_WEIGHT weight = DWRITE_FONT_WEIGHT_NORMAL) noexcept
{
    return TypographySpec{.familyName = kSegoeFluentIconsFamily, .weight = weight, .sizeDip = sizeDip};
}

[[nodiscard]] inline TypographySpec MakeUiEmojiSpec(float sizeDip, DWRITE_FONT_WEIGHT weight = DWRITE_FONT_WEIGHT_NORMAL) noexcept
{
    return TypographySpec{.familyName = kSegoeUiEmojiFamily, .weight = weight, .sizeDip = sizeDip};
}

[[nodiscard]] inline TypographySpec MakeUiMonospaceSpec(float sizeDip, DWRITE_FONT_WEIGHT weight = DWRITE_FONT_WEIGHT_NORMAL) noexcept
{
    return TypographySpec{.familyName = kUiMonospaceFamily, .weight = weight, .sizeDip = sizeDip};
}

[[nodiscard]] inline TypographySpec GetDxUiTypographySpec(FontRole role) noexcept
{
    switch (role)
    {
        case FontRole::BodyStrong: return MakeUiTextSpec(14.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD);
        case FontRole::BodyLarge: return MakeUiTextSpec(18.0f, DWRITE_FONT_WEIGHT_NORMAL);
        case FontRole::ListItem: return MakeUiTextSpec(12.0f, DWRITE_FONT_WEIGHT_NORMAL);
        case FontRole::Title: return MakeUiTextSpec(24.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD);
        case FontRole::Subtitle: return MakeUiTextSpec(20.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD);
        case FontRole::TitleLarge: return MakeUiTextSpec(40.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD);
        case FontRole::Display: return MakeUiTextSpec(68.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD);
        case FontRole::Header: return MakeUiTextSpec(12.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD);
        case FontRole::Small: return MakeUiTextSpec(11.0f, DWRITE_FONT_WEIGHT_NORMAL);
        case FontRole::Icon: return MakeUiIconSpec(12.0f);
        case FontRole::HeroIcon: return MakeUiIconSpec(64.0f);
        case FontRole::Monospace: return MakeUiMonospaceSpec(12.0f);
        case FontRole::Body:
        default: return MakeUiTextSpec(13.0f, DWRITE_FONT_WEIGHT_NORMAL);
    }
}

template <size_t Capacity> [[nodiscard]] inline bool CopyNullTerminated(std::wstring_view text, wchar_t (&destination)[Capacity]) noexcept
{
    if (text.size() >= Capacity)
    {
        return false;
    }

    std::copy_n(text.data(), text.size(), destination);
    destination[text.size()] = L'\0';
    return true;
}

[[nodiscard]] inline bool IsNullTerminatedWithin(PCWSTR text, size_t maxLength) noexcept
{
    return text && wcsnlen_s(text, maxLength + 1u) <= maxLength;
}

[[nodiscard]] inline bool FontFamilyEquals(PCWSTR lhs, PCWSTR rhs) noexcept
{
    return lhs && rhs && std::wcscmp(lhs, rhs) == 0;
}

[[nodiscard]] inline PCWSTR GetFallbackFamilyName(PCWSTR preferredFamilyName) noexcept
{
    if (FontFamilyEquals(preferredFamilyName, kSegoeFluentIconsFamily))
    {
        return kSegoeMdl2AssetsFamily;
    }
    if (FontFamilyEquals(preferredFamilyName, kUiMonospaceFamily))
    {
        return kUiMonospaceFamily;
    }
    if (FontFamilyEquals(preferredFamilyName, kSegoeUiEmojiFamily))
    {
        return kSegoeUiEmojiFamily;
    }
    return kSegoeUiFallbackFamily;
}

[[nodiscard]] inline std::mutex& GetTypographyMeasurementCacheMutex() noexcept
{
    static std::mutex cacheMutex;
    return cacheMutex;
}

[[nodiscard]] inline std::vector<TypographyFontFamilyCacheEntry>& GetTypographyFontFamilyCache() noexcept
{
    static std::vector<TypographyFontFamilyCacheEntry> cache;
    return cache;
}

[[nodiscard]] inline std::vector<TypographyTextFormatCacheEntry>& GetTypographyTextFormatCache() noexcept
{
    static std::vector<TypographyTextFormatCacheEntry> cache;
    return cache;
}

[[nodiscard]] inline bool QueryFontFamilyAvailable(IDWriteFactory* dwriteFactory, PCWSTR familyName) noexcept
{
    if (! dwriteFactory || ! IsNullTerminatedWithin(familyName, kMaxDWriteFamilyNameLength) || familyName[0] == L'\0')
    {
        return false;
    }

    wil::com_ptr<IDWriteFontCollection> fontCollection;
    if (FAILED(dwriteFactory->GetSystemFontCollection(fontCollection.put())))
    {
        return false;
    }

    UINT32 familyIndex = 0u;
    BOOL familyExists  = FALSE;
    const HRESULT hr   = fontCollection->FindFamilyName(familyName, &familyIndex, &familyExists);
    return SUCCEEDED(hr) && familyExists == TRUE;
}

[[nodiscard]] inline bool IsFontFamilyAvailable(IDWriteFactory* dwriteFactory, PCWSTR familyName) noexcept
{
    if (! dwriteFactory || ! IsNullTerminatedWithin(familyName, kMaxDWriteFamilyNameLength) || familyName[0] == L'\0')
    {
        return false;
    }

    {
        std::scoped_lock lock(GetTypographyMeasurementCacheMutex());
        for (const TypographyFontFamilyCacheEntry& entry : GetTypographyFontFamilyCache())
        {
            if (entry.factoryKey == dwriteFactory && entry.familyName == familyName)
            {
                return entry.available;
            }
        }
    }

    const bool available = QueryFontFamilyAvailable(dwriteFactory, familyName);
    EmitTypographyPerfCounter(kTypographyFamilyCacheMissMetric, familyName, 0u, 1u, available ? 1u : 0u, S_OK);

    std::scoped_lock lock(GetTypographyMeasurementCacheMutex());
    for (const TypographyFontFamilyCacheEntry& entry : GetTypographyFontFamilyCache())
    {
        if (entry.factoryKey == dwriteFactory && entry.familyName == familyName)
        {
            return entry.available;
        }
    }

    TypographyFontFamilyCacheEntry entry;
    entry.factory    = dwriteFactory;
    entry.factoryKey = dwriteFactory;
    entry.familyName = familyName;
    entry.available  = available;
    GetTypographyFontFamilyCache().push_back(std::move(entry));
    return available;
}

[[nodiscard]] inline std::wstring ResolveCachedFontFamilyName(IDWriteFactory* dwriteFactory, PCWSTR preferredFamilyName) noexcept
{
    if (! IsNullTerminatedWithin(preferredFamilyName, kMaxDWriteFamilyNameLength) || preferredFamilyName[0] == L'\0')
    {
        return {};
    }

    PCWSTR fallbackFamily = GetFallbackFamilyName(preferredFamilyName);
    return IsFontFamilyAvailable(dwriteFactory, preferredFamilyName) ? std::wstring(preferredFamilyName) : std::wstring(fallbackFamily);
}

[[nodiscard]] inline HRESULT CreateTextFormat(IDWriteFactory* dwriteFactory,
                                              const TypographySpec& spec,
                                              IDWriteTextFormat** outFormat,
                                              PCWSTR localeName = L"") noexcept
{
    if (! dwriteFactory || ! outFormat)
    {
        return E_INVALIDARG;
    }

    *outFormat = nullptr;
    wchar_t preferredFamilyBuffer[kMaxDWriteFamilyNameLength + 1u]{};
    if (! CopyNullTerminated(spec.familyName, preferredFamilyBuffer) || ! IsNullTerminatedWithin(localeName, LOCALE_NAME_MAX_LENGTH - 1u))
    {
        return E_INVALIDARG;
    }

    PCWSTR fallbackFamily        = GetFallbackFamilyName(preferredFamilyBuffer);
    std::wstring resolvedFamily  = ResolveCachedFontFamilyName(dwriteFactory, preferredFamilyBuffer);
    PCWSTR resolvedFamilyPointer = resolvedFamily.empty() ? fallbackFamily : resolvedFamily.c_str();
    HRESULT hr                   = dwriteFactory->CreateTextFormat(
        resolvedFamilyPointer, nullptr, spec.weight, spec.style, DWRITE_FONT_STRETCH_NORMAL, spec.sizeDip, localeName, outFormat);
    if (FAILED(hr) && ! FontFamilyEquals(resolvedFamilyPointer, fallbackFamily))
    {
        hr = dwriteFactory->CreateTextFormat(fallbackFamily, nullptr, spec.weight, spec.style, DWRITE_FONT_STRETCH_NORMAL, spec.sizeDip, localeName, outFormat);
    }
    return hr;
}

[[nodiscard]] inline HRESULT CreateTextFormatWithStyle(IDWriteFactory* dwriteFactory,
                                                       PCWSTR preferredFamily,
                                                       DWRITE_FONT_WEIGHT weight,
                                                       DWRITE_FONT_STYLE style,
                                                       float sizeDip,
                                                       IDWriteTextFormat** outFormat,
                                                       PCWSTR localeName = L"") noexcept
{
    if (! dwriteFactory || ! outFormat)
    {
        return E_INVALIDARG;
    }

    *outFormat = nullptr;
    if (! IsNullTerminatedWithin(preferredFamily, kMaxDWriteFamilyNameLength) || ! IsNullTerminatedWithin(localeName, LOCALE_NAME_MAX_LENGTH - 1u))
    {
        return E_INVALIDARG;
    }

    PCWSTR fallbackFamily        = GetFallbackFamilyName(preferredFamily);
    std::wstring resolvedFamily  = ResolveCachedFontFamilyName(dwriteFactory, preferredFamily);
    PCWSTR resolvedFamilyPointer = resolvedFamily.empty() ? fallbackFamily : resolvedFamily.c_str();
    HRESULT hr = dwriteFactory->CreateTextFormat(resolvedFamilyPointer, nullptr, weight, style, DWRITE_FONT_STRETCH_NORMAL, sizeDip, localeName, outFormat);
    if (FAILED(hr) && ! FontFamilyEquals(resolvedFamilyPointer, fallbackFamily))
    {
        hr = dwriteFactory->CreateTextFormat(fallbackFamily, nullptr, weight, style, DWRITE_FONT_STRETCH_NORMAL, sizeDip, localeName, outFormat);
    }
    return hr;
}

[[nodiscard]] inline UINT GetEffectiveDpi(HWND hwnd) noexcept
{
    const UINT dpi = hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;
    return (std::max<UINT>)(dpi, USER_DEFAULT_SCREEN_DPI);
}

[[nodiscard]] inline float GetPixelsPerDip(UINT dpi) noexcept
{
    return static_cast<float>((std::max<UINT>)(dpi, USER_DEFAULT_SCREEN_DPI)) / static_cast<float>(USER_DEFAULT_SCREEN_DPI);
}

[[nodiscard]] inline int DipExtentToPixels(float extentDip, UINT dpi) noexcept
{
    if (! (extentDip > 0.0f))
    {
        return 0;
    }

    const float extentPx = extentDip * GetPixelsPerDip(dpi);
    return (std::max)(0, static_cast<int>(std::lround(extentPx)));
}

[[nodiscard]] inline IDWriteFactory* GetSharedMeasurementFactory() noexcept
{
    static const auto resources = []() noexcept
    {
        wil::com_ptr<IDWriteFactory> factory;
        static_cast<void>(DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(factory.put())));
        return factory;
    }();

    return resources.get();
}

[[nodiscard]] inline IDWriteTextFormat* GetCachedMeasurementTextFormat(IDWriteFactory* dwriteFactory, FontRole role, bool wrap) noexcept
{
    if (! dwriteFactory)
    {
        return nullptr;
    }

    {
        std::scoped_lock lock(GetTypographyMeasurementCacheMutex());
        for (const TypographyTextFormatCacheEntry& entry : GetTypographyTextFormatCache())
        {
            if (entry.factoryKey == dwriteFactory && entry.role == role && entry.wrap == wrap)
            {
                return entry.textFormat.get();
            }
        }
    }

    wil::com_ptr<IDWriteTextFormat> textFormat;
    if (FAILED(CreateTextFormat(dwriteFactory, GetDxUiTypographySpec(role), textFormat.put())) || ! textFormat)
    {
        return nullptr;
    }

    static_cast<void>(textFormat->SetWordWrapping(wrap ? DWRITE_WORD_WRAPPING_WRAP : DWRITE_WORD_WRAPPING_NO_WRAP));
    static_cast<void>(textFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
    static_cast<void>(textFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));

    EmitTypographyPerfCounter(kTypographyTextFormatCacheMissMetric, L"", 0u, 1u, static_cast<uint64_t>(role), S_OK);

    std::scoped_lock lock(GetTypographyMeasurementCacheMutex());
    for (const TypographyTextFormatCacheEntry& entry : GetTypographyTextFormatCache())
    {
        if (entry.factoryKey == dwriteFactory && entry.role == role && entry.wrap == wrap)
        {
            return entry.textFormat.get();
        }
    }

    TypographyTextFormatCacheEntry entry;
    entry.factory    = dwriteFactory;
    entry.factoryKey = dwriteFactory;
    entry.role       = role;
    entry.wrap       = wrap;
    entry.textFormat = std::move(textFormat);
    GetTypographyTextFormatCache().push_back(std::move(entry));
    return GetTypographyTextFormatCache().back().textFormat.get();
}

[[nodiscard]] inline TextPixelMetrics MeasureSingleLineTextMetrics(IDWriteFactory* dwriteFactory,
                                                                   IDWriteTextFormat* textFormat,
                                                                   UINT dpi,
                                                                   std::wstring_view text) noexcept
{
    TextPixelMetrics result{};
    if (! dwriteFactory || ! textFormat)
    {
        return result;
    }

    constexpr std::wstring_view kLineHeightSample = L"Ag";
    const bool useSampleOnly                      = text.empty();
    const std::wstring_view measureText           = useSampleOnly ? kLineHeightSample : text;
    if (measureText.size() > static_cast<size_t>((std::numeric_limits<UINT32>::max)()))
    {
        return result;
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    const float pixelsPerDip = GetPixelsPerDip(dpi);
    HRESULT hr               = dwriteFactory->CreateGdiCompatibleTextLayout(
        measureText.data(), static_cast<UINT32>(measureText.size()), textFormat, 16384.0f, 1024.0f, pixelsPerDip, nullptr, FALSE, layout.put());
    if (FAILED(hr) || ! layout)
    {
        hr = dwriteFactory->CreateTextLayout(measureText.data(), static_cast<UINT32>(measureText.size()), textFormat, 16384.0f, 1024.0f, layout.put());
        if (FAILED(hr) || ! layout)
        {
            return result;
        }
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return result;
    }

    const float widthDip = metrics.widthIncludingTrailingWhitespace > 0.0f ? metrics.widthIncludingTrailingWhitespace : metrics.width;
    result.widthPx       = useSampleOnly ? 0 : DipExtentToPixels(widthDip, dpi);
    result.heightPx      = DipExtentToPixels(metrics.height, dpi);
    result.lineHeightPx  = result.heightPx;
    return result;
}

[[nodiscard]] inline int MeasureSingleLineTextWidthPx(HWND hwnd, FontRole role, std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return 0;
    }

    IDWriteFactory* dwriteFactory = GetSharedMeasurementFactory();
    if (! dwriteFactory)
    {
        return 0;
    }

    IDWriteTextFormat* textFormat = GetCachedMeasurementTextFormat(dwriteFactory, role, false);
    if (! textFormat)
    {
        return 0;
    }

    return MeasureSingleLineTextMetrics(dwriteFactory, textFormat, GetEffectiveDpi(hwnd), text).widthPx;
}

[[nodiscard]] inline int MeasureWrappedTextHeightPx(HWND hwnd, FontRole role, int widthPx, std::wstring_view text) noexcept
{
    if (widthPx <= 0 || text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<UINT32>::max)()))
    {
        return 0;
    }

    IDWriteFactory* dwriteFactory = GetSharedMeasurementFactory();
    if (! dwriteFactory)
    {
        return 0;
    }

    IDWriteTextFormat* textFormat = GetCachedMeasurementTextFormat(dwriteFactory, role, true);
    if (! textFormat)
    {
        return 0;
    }

    const UINT dpi             = GetEffectiveDpi(hwnd);
    const float layoutWidthDip = (std::max)(1.0f, (static_cast<float>(widthPx) * static_cast<float>(USER_DEFAULT_SCREEN_DPI)) / static_cast<float>(dpi));
    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(dwriteFactory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), textFormat, layoutWidthDip, 4096.0f, layout.put())) || ! layout)
    {
        return 0;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return 0;
    }

    return DipExtentToPixels(metrics.height, dpi);
}

} // namespace RedSalamander::DxUi::Typography
