#ifndef NOMINMAX
#define NOMINMAX
#endif

#include "DxUi.Internal.h"

#include "Helpers.h"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdint>
#include <cwctype>

namespace RedSalamander::DxUi
{
[[nodiscard]] bool IsUtf16LeadSurrogate(wchar_t ch) noexcept
{
    return ch >= 0xD800 && ch <= 0xDBFF;
}

[[nodiscard]] bool IsUtf16TrailSurrogate(wchar_t ch) noexcept
{
    return ch >= 0xDC00 && ch <= 0xDFFF;
}

[[nodiscard]] bool HasCharTypeFlag(wchar_t ch, WORD mask) noexcept
{
    WORD charType = 0u;
    return GetStringTypeW(CT_CTYPE1, &ch, 1, &charType) != FALSE && (charType & mask) != 0u;
}

[[nodiscard]] bool IsWhitespaceCharacter(wchar_t ch) noexcept
{
    return HasCharTypeFlag(ch, C1_SPACE);
}

[[nodiscard]] size_t StepToPreviousCodePoint(std::wstring_view text, size_t caretIndex) noexcept
{
    const size_t caret = std::min(caretIndex, text.size());
    if (caret == 0u)
    {
        return 0u;
    }

    if (caret < text.size() && IsUtf16TrailSurrogate(text[caret]) && IsUtf16LeadSurrogate(text[caret - 1u]))
    {
        return caret - 1u;
    }

    const size_t previousIndex = caret - 1u;
    if (previousIndex > 0u && IsUtf16TrailSurrogate(text[previousIndex]) && IsUtf16LeadSurrogate(text[previousIndex - 1u]))
    {
        return previousIndex - 1u;
    }

    return previousIndex;
}

[[nodiscard]] size_t StepToNextCodePoint(std::wstring_view text, size_t caretIndex) noexcept
{
    const size_t caret = std::min(caretIndex, text.size());
    if (caret >= text.size())
    {
        return text.size();
    }

    if (caret > 0u && IsUtf16TrailSurrogate(text[caret]) && IsUtf16LeadSurrogate(text[caret - 1u]))
    {
        return std::min(text.size(), caret + 1u);
    }

    if ((caret + 1u) < text.size() && IsUtf16LeadSurrogate(text[caret]) && IsUtf16TrailSurrogate(text[caret + 1u]))
    {
        return caret + 2u;
    }

    return caret + 1u;
}

struct Utf16CodePoint
{
    uint32_t value{};
    size_t end{};
};

[[nodiscard]] Utf16CodePoint ReadCodePointAt(std::wstring_view text, size_t index) noexcept
{
    const size_t begin = std::min(index, text.size());
    if (begin >= text.size())
    {
        return {};
    }

    const wchar_t first = text[begin];
    if ((begin + 1u) < text.size() && IsUtf16LeadSurrogate(first) && IsUtf16TrailSurrogate(text[begin + 1u]))
    {
        const uint32_t high = static_cast<uint32_t>(first) - 0xD800u;
        const uint32_t low  = static_cast<uint32_t>(text[begin + 1u]) - 0xDC00u;
        return {0x10000u + ((high << 10u) | low), begin + 2u};
    }

    return {static_cast<uint32_t>(first), begin + 1u};
}

[[nodiscard]] bool IsVariationSelectorCodePoint(uint32_t value) noexcept
{
    return (value >= 0xFE00u && value <= 0xFE0Fu) || (value >= 0xE0100u && value <= 0xE01EFu);
}

[[nodiscard]] bool IsEmojiModifierCodePoint(uint32_t value) noexcept
{
    return value >= 0x1F3FBu && value <= 0x1F3FFu;
}

[[nodiscard]] bool IsRegionalIndicatorCodePoint(uint32_t value) noexcept
{
    return value >= 0x1F1E6u && value <= 0x1F1FFu;
}

[[nodiscard]] size_t ConsumeEmojiSuffix(std::wstring_view text, size_t index) noexcept
{
    size_t cursor = std::min(index, text.size());
    while (cursor < text.size())
    {
        const Utf16CodePoint suffix = ReadCodePointAt(text, cursor);
        if (! IsVariationSelectorCodePoint(suffix.value) && ! IsEmojiModifierCodePoint(suffix.value))
        {
            break;
        }

        cursor = suffix.end;
    }

    return cursor;
}

[[nodiscard]] size_t FindNextTextElementBoundary(std::wstring_view text, size_t elementStart) noexcept
{
    const size_t start = std::min(elementStart, text.size());
    if (start >= text.size())
    {
        return text.size();
    }

    const Utf16CodePoint first = ReadCodePointAt(text, start);
    size_t boundary            = ConsumeEmojiSuffix(text, first.end);

    if (IsRegionalIndicatorCodePoint(first.value) && boundary < text.size())
    {
        const Utf16CodePoint second = ReadCodePointAt(text, boundary);
        if (IsRegionalIndicatorCodePoint(second.value))
        {
            return ConsumeEmojiSuffix(text, second.end);
        }
    }

    while (boundary < text.size())
    {
        const Utf16CodePoint joiner = ReadCodePointAt(text, boundary);
        if (joiner.value != 0x200Du)
        {
            break;
        }

        boundary = joiner.end;
        if (boundary >= text.size())
        {
            break;
        }

        const Utf16CodePoint joined = ReadCodePointAt(text, boundary);
        boundary                    = ConsumeEmojiSuffix(text, joined.end);
    }

    return boundary > start ? boundary : StepToNextCodePoint(text, start);
}

[[nodiscard]] size_t StepToPreviousTextElement(std::wstring_view text, size_t caretIndex) noexcept
{
    const bool perfEnabled = Debug::Perf::IsCaptureEnabled();
    const auto startedAt   = perfEnabled ? std::chrono::steady_clock::now() : std::chrono::steady_clock::time_point{};
    const auto finish      = [perfEnabled, startedAt, text](size_t result) noexcept
    {
        if (perfEnabled)
        {
            Debug::Perf::Emit(L"dxui.textinput.grapheme_step_us", L"previous", Debug::Perf::ElapsedUs(startedAt), text.size(), result, S_OK);
        }
        return result;
    };

    const size_t caret = std::min(caretIndex, text.size());
    if (caret == 0u)
    {
        return finish(0u);
    }

    size_t cursor = 0u;
    while (cursor < text.size())
    {
        const size_t nextBoundary = FindNextTextElementBoundary(text, cursor);
        if (nextBoundary >= caret)
        {
            return finish(cursor);
        }

        cursor = nextBoundary;
    }

    return finish(0u);
}

[[nodiscard]] size_t StepToNextTextElement(std::wstring_view text, size_t caretIndex) noexcept
{
    const bool perfEnabled = Debug::Perf::IsCaptureEnabled();
    const auto startedAt   = perfEnabled ? std::chrono::steady_clock::now() : std::chrono::steady_clock::time_point{};
    const auto finish      = [perfEnabled, startedAt, text](size_t result) noexcept
    {
        if (perfEnabled)
        {
            Debug::Perf::Emit(L"dxui.textinput.grapheme_step_us", L"next", Debug::Perf::ElapsedUs(startedAt), text.size(), result, S_OK);
        }
        return result;
    };

    const size_t caret = std::min(caretIndex, text.size());
    if (caret >= text.size())
    {
        return finish(text.size());
    }

    size_t cursor = 0u;
    while (cursor < text.size())
    {
        const size_t nextBoundary = FindNextTextElementBoundary(text, cursor);
        if (caret < nextBoundary)
        {
            return finish(nextBoundary);
        }

        cursor = nextBoundary;
    }

    return finish(text.size());
}

[[nodiscard]] size_t SnapCaretIndexToTextElementBoundary(std::wstring_view text, size_t caretIndex) noexcept
{
    const size_t caret = std::min(caretIndex, text.size());
    if (caret == 0u || caret == text.size())
    {
        return caret;
    }

    size_t cursor = 0u;
    while (cursor < text.size())
    {
        const size_t nextBoundary = FindNextTextElementBoundary(text, cursor);
        if (caret == cursor || caret == nextBoundary)
        {
            return caret;
        }
        if (caret < nextBoundary)
        {
            const size_t distanceToStart = caret - cursor;
            const size_t distanceToEnd   = nextBoundary - caret;
            return distanceToStart < distanceToEnd ? cursor : nextBoundary;
        }

        cursor = nextBoundary;
    }

    return text.size();
}

[[nodiscard]] bool IsWordCharacter(wchar_t ch) noexcept
{
    return HasCharTypeFlag(ch, C1_ALPHA | C1_DIGIT) || ch == L'_';
}

[[nodiscard]] bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

[[nodiscard]] bool IsRtlOrBiDiControlCodeUnit(wchar_t ch) noexcept
{
    const uint32_t codeUnit = static_cast<uint32_t>(ch);
    return codeUnit == 0x061Cu || codeUnit == 0x200Fu || (codeUnit >= 0x0590u && codeUnit <= 0x08FFu) || (codeUnit >= 0x202Au && codeUnit <= 0x202Eu) ||
           (codeUnit >= 0x2066u && codeUnit <= 0x2069u) || (codeUnit >= 0xFB1Du && codeUnit <= 0xFDFFu) || (codeUnit >= 0xFE70u && codeUnit <= 0xFEFFu);
}

[[nodiscard]] bool ShouldEmitSingleLineBiDiTextMetric(std::wstring_view text, DWRITE_READING_DIRECTION readingDirection) noexcept
{
    if (readingDirection == DWRITE_READING_DIRECTION_RIGHT_TO_LEFT)
    {
        return true;
    }

    return std::any_of(text.begin(), text.end(), IsRtlOrBiDiControlCodeUnit);
}

[[nodiscard]] size_t FindPreviousWordBoundary(std::wstring_view text, size_t caretIndex) noexcept
{
    size_t caret = std::min(caretIndex, text.size());
    if (caret == 0u)
    {
        return 0u;
    }

    if (caret < text.size() && IsUtf16TrailSurrogate(text[caret]) && IsUtf16LeadSurrogate(text[caret - 1u]))
    {
        --caret;
    }

    size_t eraseFrom = caret;
    while (eraseFrom > 0u)
    {
        const size_t previousIndex = StepToPreviousCodePoint(text, eraseFrom);
        if (! IsWhitespaceCharacter(text[previousIndex]))
        {
            break;
        }

        eraseFrom = previousIndex;
    }

    if (eraseFrom > 0u)
    {
        const size_t previousIndex = StepToPreviousCodePoint(text, eraseFrom);
        const wchar_t previous     = text[previousIndex];
        if (IsPathSeparator(previous))
        {
            while (eraseFrom > 0u)
            {
                const size_t currentIndex = StepToPreviousCodePoint(text, eraseFrom);
                if (! IsPathSeparator(text[currentIndex]))
                {
                    break;
                }

                eraseFrom = currentIndex;
            }

            if (eraseFrom == 2u && text.size() >= 2u && text[1u] == L':')
            {
                return eraseFrom;
            }

            if (eraseFrom > 0u)
            {
                const size_t currentIndex = StepToPreviousCodePoint(text, eraseFrom);
                const wchar_t current     = text[currentIndex];
                if (IsWordCharacter(current))
                {
                    while (eraseFrom > 0u)
                    {
                        const size_t wordIndex = StepToPreviousCodePoint(text, eraseFrom);
                        if (! IsWordCharacter(text[wordIndex]))
                        {
                            break;
                        }

                        eraseFrom = wordIndex;
                    }
                }
                else if (! IsWhitespaceCharacter(current) && ! IsPathSeparator(current))
                {
                    while (eraseFrom > 0u)
                    {
                        const size_t punctuationIndex = StepToPreviousCodePoint(text, eraseFrom);
                        const wchar_t punctuation     = text[punctuationIndex];
                        if (IsWhitespaceCharacter(punctuation) || IsPathSeparator(punctuation) || IsWordCharacter(punctuation))
                        {
                            break;
                        }

                        eraseFrom = punctuationIndex;
                    }
                }
            }
        }
        else if (IsWordCharacter(previous))
        {
            while (eraseFrom > 0u)
            {
                const size_t currentIndex = StepToPreviousCodePoint(text, eraseFrom);
                if (! IsWordCharacter(text[currentIndex]))
                {
                    break;
                }

                eraseFrom = currentIndex;
            }
        }
        else
        {
            while (eraseFrom > 0u)
            {
                const size_t currentIndex = StepToPreviousCodePoint(text, eraseFrom);
                const wchar_t current     = text[currentIndex];
                if (IsWhitespaceCharacter(current) || IsPathSeparator(current) || IsWordCharacter(current))
                {
                    break;
                }

                eraseFrom = currentIndex;
            }
        }
    }

    return eraseFrom == caret ? StepToPreviousCodePoint(text, caret) : eraseFrom;
}

[[nodiscard]] size_t FindNextWordBoundary(std::wstring_view text, size_t caretIndex) noexcept
{
    size_t caret = std::min(caretIndex, text.size());
    if (caret > 0u && caret < text.size() && IsUtf16TrailSurrogate(text[caret]) && IsUtf16LeadSurrogate(text[caret - 1u]))
    {
        caret = std::min(text.size(), caret + 1u);
    }

    if (caret >= text.size())
    {
        return text.size();
    }

    while (caret < text.size() && IsWhitespaceCharacter(text[caret]))
    {
        caret = StepToNextCodePoint(text, caret);
    }

    if (caret >= text.size())
    {
        return text.size();
    }

    const wchar_t current = text[caret];
    if (IsPathSeparator(current))
    {
        while (caret < text.size() && IsPathSeparator(text[caret]))
        {
            caret = StepToNextCodePoint(text, caret);
        }
    }
    else if (IsWordCharacter(current))
    {
        while (caret < text.size() && IsWordCharacter(text[caret]))
        {
            caret = StepToNextCodePoint(text, caret);
        }
    }
    else
    {
        while (caret < text.size())
        {
            const wchar_t value = text[caret];
            if (IsWhitespaceCharacter(value) || IsPathSeparator(value) || IsWordCharacter(value))
            {
                break;
            }
            caret = StepToNextCodePoint(text, caret);
        }
    }

    while (caret < text.size() && IsWhitespaceCharacter(text[caret]))
    {
        caret = StepToNextCodePoint(text, caret);
    }

    return caret;
}

[[nodiscard]] float MeasureSingleLineTextWidthDip(
    const WindowHost* host, std::wstring_view text, FontRole role, float heightDip, DWRITE_READING_DIRECTION readingDirection) noexcept
{
    if (text.empty())
    {
        return 0.0f;
    }

    if (host)
    {
        auto* factory = host->GetWriteFactory();
        auto* format  = host->GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, readingDirection);
        if (factory && format)
        {
            wil::com_ptr<IDWriteTextLayout> layout;
            if (SUCCEEDED(factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, 32768.0f, heightDip, layout.addressof())) && layout)
            {
                DWRITE_TEXT_METRICS metrics{};
                if (SUCCEEDED(layout->GetMetrics(&metrics)))
                {
                    return metrics.widthIncludingTrailingWhitespace;
                }
            }
        }
    }

    return static_cast<float>(text.size()) * 7.0f;
}

[[nodiscard]] float ResolveSingleLineLayoutWidthDip(
    const WindowHost* host, std::wstring_view text, FontRole role, float heightDip, float minimumWidthDip, DWRITE_READING_DIRECTION readingDirection) noexcept
{
    const float measuredTextWidthDip = MeasureSingleLineTextWidthDip(host, text, role, heightDip, readingDirection);
    return std::max({1.0f, minimumWidthDip, measuredTextWidthDip + 32.0f});
}

[[nodiscard]] wil::com_ptr<IDWriteTextLayout> CreateSingleLineTextLayout(
    const WindowHost* host, std::wstring_view text, FontRole role, float widthDip, float heightDip, DWRITE_READING_DIRECTION readingDirection) noexcept
{
    if (! host)
    {
        return {};
    }

    auto* factory = host->GetWriteFactory();
    auto* format  = host->GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, readingDirection);
    if (! factory || ! format)
    {
        return {};
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(factory->CreateTextLayout(
            text.data(), static_cast<UINT32>(text.size()), format, std::max(1.0f, widthDip), std::max(1.0f, heightDip), layout.addressof())) ||
        ! layout)
    {
        return {};
    }
    return layout;
}

[[nodiscard]] float MeasureCaretOffsetDip(const WindowHost* host,
                                          std::wstring_view text,
                                          FontRole role,
                                          size_t caretIndex,
                                          float heightDip,
                                          DWRITE_READING_DIRECTION readingDirection,
                                          float layoutWidthDip) noexcept
{
    const size_t clampedCaret       = std::min(caretIndex, text.size());
    const float resolvedLayoutWidth = ResolveSingleLineLayoutWidthDip(host, text, role, heightDip, std::max(1.0f, layoutWidthDip), readingDirection);
    if (wil::com_ptr<IDWriteTextLayout> layout = CreateSingleLineTextLayout(host, text, role, resolvedLayoutWidth, heightDip, readingDirection))
    {
        float x = 0.0f;
        float y = 0.0f;
        DWRITE_HIT_TEST_METRICS metrics{};
        const UINT32 textPosition = static_cast<UINT32>(clampedCaret == 0u ? 0u : clampedCaret - 1u);
        const BOOL trailingHit    = clampedCaret == 0u ? FALSE : TRUE;
        if (SUCCEEDED(layout->HitTestTextPosition(textPosition, trailingHit, &x, &y, &metrics)))
        {
            return x;
        }
    }

    return static_cast<float>(clampedCaret) * 7.0f;
}

[[nodiscard]] size_t HitTestCaretIndexDip(const WindowHost* host,
                                          std::wstring_view text,
                                          FontRole role,
                                          const D2D1_RECT_F& textRect,
                                          float scrollDip,
                                          D2D1_POINT_2F point,
                                          DWRITE_READING_DIRECTION readingDirection) noexcept
{
    const bool perfEnabled    = Debug::Perf::IsCaptureEnabled();
    const auto startedAt      = perfEnabled ? std::chrono::steady_clock::now() : std::chrono::steady_clock::time_point{};
    const bool emitBiDiMetric = perfEnabled && ShouldEmitSingleLineBiDiTextMetric(text, readingDirection);
    const auto finish         = [perfEnabled, emitBiDiMetric, startedAt, text, readingDirection](size_t result) noexcept
    {
        if (perfEnabled)
        {
            const uint64_t elapsedUs = Debug::Perf::ElapsedUs(startedAt);
            Debug::Perf::Emit(L"dxui.textinput.hit_test_us", L"single-line", elapsedUs, text.size(), result, S_OK);
            if (emitBiDiMetric)
            {
                const std::wstring_view detail = readingDirection == DWRITE_READING_DIRECTION_RIGHT_TO_LEFT ? L"single-line-rtl" : L"single-line-ltr";
                Debug::Perf::Emit(L"dxui.textinput.bidi_hit_test_us", detail, elapsedUs, text.size(), result, S_OK);
            }
        }
        return result;
    };

    if (text.empty())
    {
        return finish(0u);
    }

    const float localX               = std::max(0.0f, point.x - textRect.left + scrollDip);
    const float localY               = std::clamp(point.y - textRect.top, 0.0f, std::max(1.0f, textRect.bottom - textRect.top) - 1.0f);
    const float heightDip            = std::max(1.0f, textRect.bottom - textRect.top);
    const float measuredTextWidthDip = MeasureSingleLineTextWidthDip(host, text, role, heightDip, readingDirection);
    const float layoutWidthDip       = std::max({1.0f, textRect.right - textRect.left + scrollDip, localX + 1.0f, measuredTextWidthDip + 32.0f});
    if (wil::com_ptr<IDWriteTextLayout> layout = CreateSingleLineTextLayout(host, text, role, layoutWidthDip, heightDip, readingDirection))
    {
        BOOL isTrailingHit = FALSE;
        BOOL isInside      = FALSE;
        DWRITE_HIT_TEST_METRICS metrics{};
        if (SUCCEEDED(layout->HitTestPoint(localX, localY, &isTrailingHit, &isInside, &metrics)))
        {
            if (! isInside)
            {
                if (readingDirection == DWRITE_READING_DIRECTION_RIGHT_TO_LEFT)
                {
                    const float textLeftDip = std::max(0.0f, layoutWidthDip - measuredTextWidthDip);
                    if (localX <= textLeftDip)
                    {
                        return finish(SnapCaretIndexToTextElementBoundary(text, text.size()));
                    }
                    if (localX >= layoutWidthDip - 1.0f)
                    {
                        return finish(0u);
                    }
                }
                else
                {
                    if (localX >= measuredTextWidthDip)
                    {
                        return finish(SnapCaretIndexToTextElementBoundary(text, text.size()));
                    }
                    if (localX <= 0.0f)
                    {
                        return finish(0u);
                    }
                }
            }
            const size_t textPosition    = static_cast<size_t>(metrics.textPosition);
            const size_t trailingAdvance = isTrailingHit ? static_cast<size_t>(metrics.length) : 0u;
            return finish(SnapCaretIndexToTextElementBoundary(text, textPosition + trailingAdvance));
        }
    }

    const double fallbackValue = std::clamp(std::floor(static_cast<double>(localX) / 7.0), 0.0, static_cast<double>(text.size()));
    return finish(SnapCaretIndexToTextElementBoundary(text, static_cast<size_t>(fallbackValue)));
}

void DrawSingleLineTextClipped(WindowHost& host,
                               std::wstring_view text,
                               const D2D1_RECT_F& rect,
                               FontRole role,
                               const D2D1_COLOR_F& color,
                               float scrollDip,
                               DWRITE_READING_DIRECTION readingDirection) noexcept
{
    auto* dc    = host.GetDeviceContext();
    auto* brush = host.GetSolidBrush(color);
    if (! dc || ! brush)
    {
        return;
    }

    const D2D1_RECT_F snappedRect = SnapRectToPixel(host, rect);
    const float heightDip         = std::max(1.0f, snappedRect.bottom - snappedRect.top);
    const float layoutWidthDip =
        ResolveSingleLineLayoutWidthDip(&host, text, role, heightDip, std::max(1.0f, snappedRect.right - snappedRect.left + scrollDip), readingDirection);
    if (wil::com_ptr<IDWriteTextLayout> layout = CreateSingleLineTextLayout(&host, text, role, layoutWidthDip, heightDip, readingDirection))
    {
        dc->PushAxisAlignedClip(snappedRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
        dc->DrawTextLayout(D2D1::Point2F(snappedRect.left - scrollDip, snappedRect.top), layout.get(), brush, kTextDrawOptions);
        dc->PopAxisAlignedClip();
        return;
    }

    if (auto* format = host.GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, readingDirection))
    {
        dc->PushAxisAlignedClip(snappedRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
        dc->DrawTextW(text.data(),
                      static_cast<UINT32>(text.size()),
                      format,
                      D2D1::RectF(snappedRect.left - scrollDip, snappedRect.top, snappedRect.right - scrollDip, snappedRect.bottom),
                      brush,
                      kTextDrawOptions,
                      DWRITE_MEASURING_MODE_NATURAL);
        dc->PopAxisAlignedClip();
    }
}

[[nodiscard]] std::optional<std::pair<size_t, size_t>> GetSingleLineSelectionRange(std::optional<size_t> anchorIndex, size_t caretIndex) noexcept
{
    if (! anchorIndex || anchorIndex.value() == caretIndex)
    {
        return std::nullopt;
    }

    return std::pair<size_t, size_t>(std::min(anchorIndex.value(), caretIndex), std::max(anchorIndex.value(), caretIndex));
}

void SetSingleLineCaretIndex(size_t& caretIndex, std::optional<size_t>& anchorIndex, size_t nextCaretIndex, bool extendSelection) noexcept
{
    if (extendSelection)
    {
        if (! anchorIndex)
        {
            anchorIndex = caretIndex;
        }
    }
    else
    {
        anchorIndex.reset();
    }

    caretIndex = nextCaretIndex;
    if (! extendSelection && anchorIndex && anchorIndex.value() == caretIndex)
    {
        anchorIndex.reset();
    }
}

[[nodiscard]] bool DeleteSingleLineSelection(std::wstring& text, size_t& caretIndex, std::optional<size_t>& anchorIndex) noexcept
{
    const std::optional<std::pair<size_t, size_t>> selectionRange = GetSingleLineSelectionRange(anchorIndex, caretIndex);
    if (! selectionRange)
    {
        return false;
    }

    const auto [selectionStart, selectionEnd] = selectionRange.value();
    text.erase(selectionStart, selectionEnd - selectionStart);
    caretIndex = selectionStart;
    anchorIndex.reset();
    return true;
}

void SelectAllSingleLineText(size_t textLength, size_t& caretIndex, std::optional<size_t>& anchorIndex) noexcept
{
    if (textLength == 0u)
    {
        caretIndex = 0u;
        anchorIndex.reset();
        return;
    }

    anchorIndex = 0u;
    caretIndex  = textLength;
}

void ResetSingleLineSelectionClickSequence(SingleLineSelectionClickSequence& sequence) noexcept
{
    sequence = {};
}

void ArmSingleLineSelectionClickSequence(SingleLineSelectionClickSequence& sequence, D2D1_POINT_2F pointDip) noexcept
{
    sequence.pointDip                    = pointDip;
    sequence.tickMs                      = ::GetTickCount64();
    sequence.promoteNextClickToSelectAll = true;
}

[[nodiscard]] bool ArePointsWithinSelectionClickBounds(const WindowHost& host, D2D1_POINT_2F firstPointDip, D2D1_POINT_2F secondPointDip) noexcept
{
    const float halfWidthDip  = std::max(1.0f, host.PixelsToDip(static_cast<float>(GetSystemMetrics(SM_CXDOUBLECLK))) * 0.5f);
    const float halfHeightDip = std::max(1.0f, host.PixelsToDip(static_cast<float>(GetSystemMetrics(SM_CYDOUBLECLK))) * 0.5f);
    return std::abs(firstPointDip.x - secondPointDip.x) <= halfWidthDip && std::abs(firstPointDip.y - secondPointDip.y) <= halfHeightDip;
}

bool ShouldPromoteSingleLineClickToSelectAll(const WindowHost& host, const SingleLineSelectionClickSequence& sequence, D2D1_POINT_2F pointDip) noexcept
{
    if (! sequence.promoteNextClickToSelectAll)
    {
        return false;
    }

    const uint64_t nowTickMs = ::GetTickCount64();
    if ((nowTickMs - sequence.tickMs) > static_cast<uint64_t>(::GetDoubleClickTime()))
    {
        return false;
    }

    return ArePointsWithinSelectionClickBounds(host, sequence.pointDip, pointDip);
}

[[nodiscard]] bool IsSelectionWhitespace(wchar_t value) noexcept
{
    return std::iswspace(static_cast<wint_t>(value)) != 0;
}

[[nodiscard]] int GetWordSelectionClass(wchar_t value) noexcept
{
    if (IsSelectionWhitespace(value))
    {
        return 0;
    }
    if (IsPathSeparator(value))
    {
        return 1;
    }
    if (IsWordCharacter(value))
    {
        return 2;
    }
    return 3;
}

void SelectSingleLineWordAt(std::wstring_view text, size_t hitIndex, size_t& caretIndex, std::optional<size_t>& anchorIndex) noexcept
{
    if (text.empty())
    {
        caretIndex = 0u;
        anchorIndex.reset();
        return;
    }

    size_t index             = std::min(hitIndex, text.size() - 1u);
    const int selectionClass = GetWordSelectionClass(text[index]);
    size_t selectionStart    = index;
    while (selectionStart > 0u && GetWordSelectionClass(text[selectionStart - 1u]) == selectionClass)
    {
        --selectionStart;
    }

    size_t selectionEnd = index + 1u;
    while (selectionEnd < text.size() && GetWordSelectionClass(text[selectionEnd]) == selectionClass)
    {
        ++selectionEnd;
    }

    anchorIndex = selectionStart;
    caretIndex  = selectionEnd;
    if (selectionStart == selectionEnd)
    {
        anchorIndex.reset();
    }
}

std::optional<D2D1_RECT_F> ComputeSingleLineSelectionPaintRect(const WindowHost& host,
                                                               std::wstring_view text,
                                                               const D2D1_RECT_F& rect,
                                                               FontRole role,
                                                               float scrollDip,
                                                               std::optional<std::pair<size_t, size_t>> selectionRange,
                                                               DWRITE_READING_DIRECTION readingDirection) noexcept
{
    if (! selectionRange)
    {
        return std::nullopt;
    }

    const auto [selectionStart, selectionEnd] = selectionRange.value();
    if (selectionStart >= selectionEnd)
    {
        return std::nullopt;
    }

    const D2D1_RECT_F snappedRect = SnapRectToPixel(host, rect);
    const float heightDip         = std::max(1.0f, snappedRect.bottom - snappedRect.top);
    const float layoutWidthDip    = std::max(1.0f, snappedRect.right - snappedRect.left + scrollDip);
    const float startOffsetDip    = MeasureCaretOffsetDip(&host, text, role, selectionStart, heightDip, readingDirection, layoutWidthDip);
    const float endOffsetDip      = MeasureCaretOffsetDip(&host, text, role, selectionEnd, heightDip, readingDirection, layoutWidthDip);
    const float selectionLeftDip  = std::min(startOffsetDip, endOffsetDip);
    const float selectionRightDip = std::max(startOffsetDip, endOffsetDip);
    D2D1_RECT_F selectionRect     = D2D1::RectF(
        snappedRect.left + selectionLeftDip - scrollDip, snappedRect.top + 1.0f, snappedRect.left + selectionRightDip - scrollDip, snappedRect.bottom - 1.0f);
    selectionRect       = SnapRectToPixel(host, selectionRect);
    selectionRect.left  = std::clamp(selectionRect.left, snappedRect.left, snappedRect.right);
    selectionRect.right = std::clamp(selectionRect.right, snappedRect.left, snappedRect.right);
    if (selectionRect.right <= selectionRect.left)
    {
        return std::nullopt;
    }

    return selectionRect;
}

void DrawSingleLineSelection(WindowHost& host,
                             std::wstring_view text,
                             const D2D1_RECT_F& rect,
                             FontRole role,
                             const D2D1_COLOR_F& textColor,
                             const D2D1_COLOR_F& selectionFill,
                             const D2D1_COLOR_F& selectionText,
                             float scrollDip,
                             std::optional<std::pair<size_t, size_t>> selectionRange,
                             DWRITE_READING_DIRECTION readingDirection) noexcept
{
    DrawSingleLineTextClipped(host, text, rect, role, textColor, scrollDip, readingDirection);

    if (! selectionRange)
    {
        return;
    }

    auto* dc        = host.GetDeviceContext();
    auto* fillBrush = host.GetSolidBrush(selectionFill);
    auto* textBrush = host.GetSolidBrush(selectionText);
    if (! dc || ! fillBrush || ! textBrush)
    {
        return;
    }

    const D2D1_RECT_F snappedRect                  = SnapRectToPixel(host, rect);
    const float heightDip                          = std::max(1.0f, snappedRect.bottom - snappedRect.top);
    const std::optional<D2D1_RECT_F> selectionRect = ComputeSingleLineSelectionPaintRect(host, text, rect, role, scrollDip, selectionRange, readingDirection);
    if (! selectionRect)
    {
        return;
    }

    dc->FillRectangle(selectionRect.value(), fillBrush);
    dc->PushAxisAlignedClip(selectionRect.value(), D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
    const float layoutWidthDip =
        ResolveSingleLineLayoutWidthDip(&host, text, role, heightDip, std::max(1.0f, snappedRect.right - snappedRect.left + scrollDip), readingDirection);
    if (wil::com_ptr<IDWriteTextLayout> layout = CreateSingleLineTextLayout(&host, text, role, layoutWidthDip, heightDip, readingDirection))
    {
        dc->DrawTextLayout(D2D1::Point2F(snappedRect.left - scrollDip, snappedRect.top), layout.get(), textBrush, kTextDrawOptions);
    }
    else if (auto* format = host.GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, readingDirection))
    {
        dc->DrawTextW(text.data(),
                      static_cast<UINT32>(text.size()),
                      format,
                      D2D1::RectF(snappedRect.left - scrollDip, snappedRect.top, snappedRect.right - scrollDip, snappedRect.bottom),
                      textBrush,
                      kTextDrawOptions,
                      DWRITE_MEASURING_MODE_NATURAL);
    }
    dc->PopAxisAlignedClip();
}
} // namespace RedSalamander::DxUi
