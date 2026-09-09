#include "DxUi.Internal.h"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cwctype>
#include <limits>

#include "Helpers.h"

namespace RedSalamander::DxUi
{
namespace
{
constexpr float kMultilineLayoutHeightDip = 32768.0f;
constexpr uint64_t kCaretBlinkPeriodMs    = 530u;

[[nodiscard]] std::wstring NormalizeMultilineLineEndings(std::wstring_view text, bool multiline);

struct ConcealedMaskBucket final
{
    size_t start       = 0u;
    size_t end         = 0u;
    size_t displayBase = 0u;
    size_t displaySpan = 1u;
};

[[nodiscard]] ConcealedMaskBucket GetConcealedMaskBucket(size_t exactCount) noexcept
{
    if (exactCount == 0u)
    {
        return {};
    }
    if (exactCount <= 4u)
    {
        return {.start = 1u, .end = 4u, .displayBase = 4u, .displaySpan = 4u};
    }
    if (exactCount <= 8u)
    {
        return {.start = 5u, .end = 8u, .displayBase = 8u, .displaySpan = 4u};
    }
    if (exactCount <= 12u)
    {
        return {.start = 9u, .end = 12u, .displayBase = 12u, .displaySpan = 4u};
    }

    const size_t bucketEnd = ((exactCount + 7u) / 8u) * 8u;
    return {.start = bucketEnd - 7u, .end = bucketEnd, .displayBase = bucketEnd, .displaySpan = 8u};
}

[[nodiscard]] size_t ScaleMaskedTextIndex(size_t index, size_t sourceLength, size_t targetLength) noexcept
{
    if (index == 0u || sourceLength == 0u || targetLength == 0u)
    {
        return 0u;
    }
    if (index >= sourceLength)
    {
        return targetLength;
    }

    const long double scaled = (static_cast<long double>(index) * static_cast<long double>(targetLength)) / static_cast<long double>(sourceLength);
    return (std::min)(targetLength, static_cast<size_t>(scaled + 0.5L));
}

} // namespace

[[nodiscard]] std::optional<std::pair<size_t, size_t>> ResolveNativeTextInputOptionalRange(const NativeTextInputState& state,
                                                                                           std::optional<size_t> startIndex,
                                                                                           std::optional<size_t> endIndex) noexcept
{
    if (! startIndex.has_value() || ! endIndex.has_value())
    {
        return std::nullopt;
    }

    const size_t start = (std::min)(startIndex.value(), state.text.size());
    const size_t end   = (std::min)(endIndex.value(), state.text.size());
    if (end <= start)
    {
        return std::nullopt;
    }

    return std::pair<size_t, size_t>{start, end};
}

[[nodiscard]] std::vector<D2D1_RECT_F> BuildTextInputUnderlineRects(
    const WindowHost& host, const Control& control, std::pair<size_t, size_t> range, float bottomInsetDip, float thicknessDip)
{
    std::vector<D2D1_RECT_F> underlineRects;
    const std::optional<std::vector<D2D1_RECT_F>> rangeRects = control.TryGetTextInputRangeRects(host, range.first, range.second);
    if (! rangeRects.has_value())
    {
        return underlineRects;
    }

    for (const D2D1_RECT_F& rangeRect : rangeRects.value())
    {
        const D2D1_RECT_F snappedRange = SnapRectToPixel(host, rangeRect);
        if (snappedRange.right <= snappedRange.left || snappedRange.bottom <= snappedRange.top)
        {
            continue;
        }

        const float thickness       = std::max(1.0f, thicknessDip);
        const float bottom          = std::min(snappedRange.bottom, std::max(snappedRange.top + thickness, snappedRange.bottom - bottomInsetDip));
        const D2D1_RECT_F underline = SnapRectToPixel(host, D2D1::RectF(snappedRange.left, bottom - thickness, snappedRange.right, bottom));
        if (underline.right > underline.left && underline.bottom > underline.top)
        {
            underlineRects.push_back(underline);
        }
    }

    return underlineRects;
}

[[nodiscard]] std::vector<D2D1_RECT_F> BuildNativeCompositionUnderlineRects(const WindowHost& host, const Control& control, const bool conversionTarget)
{
    NativeTextInputState state{};
    if (! host.TryReadNativeTextInputState(&control, state))
    {
        return {};
    }

    const std::optional<std::pair<size_t, size_t>> range =
        conversionTarget ? ResolveNativeTextInputOptionalRange(state, state.conversionTargetStartIndex, state.conversionTargetEndIndex)
                         : ResolveNativeTextInputOptionalRange(state, state.compositionStartIndex, state.compositionEndIndex);
    if (! range.has_value())
    {
        return {};
    }

    return BuildTextInputUnderlineRects(host, control, range.value(), conversionTarget ? 1.0f : 2.0f, conversionTarget ? 2.0f : 1.0f);
}

void DrawTextInputUnderlineRects(WindowHost& host, const std::vector<D2D1_RECT_F>& underlineRects, const D2D1_COLOR_F& color) noexcept
{
    auto* dc    = host.GetDeviceContext();
    auto* brush = host.GetSolidBrush(color);
    if (! dc || ! brush)
    {
        return;
    }

    for (const D2D1_RECT_F& rect : underlineRects)
    {
        dc->FillRectangle(rect, brush);
    }
}

namespace
{

[[nodiscard]] std::wstring NormalizeMultilineLineEndings(std::wstring_view text, bool multiline)
{
    if (! multiline || text.empty())
    {
        return std::wstring(text);
    }

    std::wstring normalized;
    normalized.reserve(text.size());
    for (size_t index = 0u; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch == L'\r')
        {
            if (index + 1u < text.size() && text[index + 1u] == L'\n')
            {
                ++index;
            }

            normalized.push_back(L'\n');
            continue;
        }

        normalized.push_back(ch);
    }

    return normalized;
}

[[nodiscard]] std::wstring NormalizePastedControlText(std::wstring_view text, bool multiline)
{
    if (text.empty())
    {
        return std::wstring(text);
    }

    if (multiline)
    {
        return NormalizeMultilineLineEndings(text, true);
    }

    std::wstring normalized;
    normalized.reserve(text.size());
    for (wchar_t ch : text)
    {
        const auto codeUnit = static_cast<unsigned int>(ch);
        if (codeUnit < 0x20u || (codeUnit >= 0x7Fu && codeUnit <= 0x9Fu))
        {
            continue;
        }

        normalized.push_back(ch);
    }

    return normalized;
}

[[nodiscard]] size_t FindLineStart(std::wstring_view text, size_t caretIndex) noexcept
{
    size_t index = std::min(caretIndex, text.size());
    while (index > 0u && text[index - 1u] != L'\n')
    {
        --index;
    }
    return index;
}

[[nodiscard]] size_t FindLineEnd(std::wstring_view text, size_t caretIndex) noexcept
{
    size_t index = std::min(caretIndex, text.size());
    while (index < text.size() && text[index] != L'\n')
    {
        ++index;
    }
    return index;
}

[[nodiscard]] wil::com_ptr<IDWriteTextLayout> CreateMultilineTextLayout(
    const WindowHost* host, std::wstring_view text, FontRole role, float widthDip, float heightDip) noexcept;
[[nodiscard]] std::vector<DWRITE_LINE_METRICS> GetMultilineLineMetrics(IDWriteTextLayout* layout) noexcept;
[[nodiscard]] std::optional<size_t> TryGetMultilineCaretLineIndex(IDWriteTextLayout* layout,
                                                                  const std::vector<DWRITE_LINE_METRICS>& metrics,
                                                                  size_t caretIndex,
                                                                  size_t textLength) noexcept;
struct WrappedLineTextRange
{
    size_t start = 0u;
    size_t end   = 0u;
};
[[nodiscard]] std::vector<WrappedLineTextRange> BuildWrappedLineTextRanges(const std::vector<DWRITE_LINE_METRICS>& metrics, size_t textLength) noexcept;
[[nodiscard]] std::optional<size_t> TryMoveCaretToWrappedLineBoundary(
    const WindowHost* host, std::wstring_view text, FontRole role, const D2D1_RECT_F& textRect, size_t caretIndex, bool moveToLineEnd) noexcept;

[[nodiscard]] size_t MoveCaretVerticallyByLogicalLine(std::wstring_view text,
                                                      size_t caretIndex,
                                                      bool moveDown,
                                                      std::optional<float>& preferredXOffsetDip) noexcept
{
    const size_t caret            = std::min(caretIndex, text.size());
    const size_t currentLineStart = FindLineStart(text, caret);
    const size_t currentLineEnd   = FindLineEnd(text, caret);
    const size_t currentColumn    = std::min(caret - currentLineStart, currentLineEnd - currentLineStart);
    const float targetXOffsetDip  = preferredXOffsetDip.value_or(static_cast<float>(currentColumn) * 7.0f);
    preferredXOffsetDip           = targetXOffsetDip;
    const size_t targetColumn     = static_cast<size_t>(std::max(0.0, std::floor(static_cast<double>(targetXOffsetDip) / 7.0)));

    if (moveDown)
    {
        if (currentLineEnd >= text.size())
        {
            return caret;
        }

        const size_t nextLineStart = currentLineEnd + 1u;
        const size_t nextLineEnd   = FindLineEnd(text, nextLineStart);
        return nextLineStart + std::min(targetColumn, nextLineEnd - nextLineStart);
    }

    if (currentLineStart == 0u)
    {
        return caret;
    }

    const size_t previousLineEnd   = currentLineStart - 1u;
    const size_t previousLineStart = FindLineStart(text, previousLineEnd);
    return previousLineStart + std::min(targetColumn, previousLineEnd - previousLineStart);
}

struct MultilineViewportMetrics
{
    size_t totalLineCount   = 1u;
    size_t visibleLineCount = 1u;
};

[[nodiscard]] size_t CountLogicalTextLines(std::wstring_view text) noexcept
{
    return 1u + static_cast<size_t>(std::count(text.begin(), text.end(), L'\n'));
}

[[nodiscard]] size_t ComputeMultilinePageLineCount(float viewportHeightDip, float lineHeightDip) noexcept
{
    const float effectiveLineHeightDip = std::max(1.0f, lineHeightDip);
    return std::max<size_t>(1u, static_cast<size_t>(std::floor(std::max(1.0f, viewportHeightDip) / effectiveLineHeightDip)));
}

[[nodiscard]] size_t ComputeMultilineWheelLineCount(size_t pageLineCount) noexcept
{
    UINT systemLines = 0u;
    if (SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &systemLines, 0) == FALSE)
    {
        systemLines = 3u;
    }

    if (systemLines == WHEEL_PAGESCROLL)
    {
        return std::max<size_t>(1u, pageLineCount);
    }

    if (systemLines == 0u)
    {
        return 0u;
    }

    return static_cast<size_t>(systemLines);
}

[[nodiscard]] size_t MoveCaretByPage(
    std::wstring_view text, size_t caretIndex, bool moveDown, std::optional<float>& preferredXOffsetDip, size_t lineCount) noexcept
{
    size_t nextCaretIndex = caretIndex;
    for (size_t step = 0u; step < lineCount; ++step)
    {
        const size_t movedIndex = MoveCaretVerticallyByLogicalLine(text, nextCaretIndex, moveDown, preferredXOffsetDip);
        if (movedIndex == nextCaretIndex)
        {
            break;
        }

        nextCaretIndex = movedIndex;
    }

    return nextCaretIndex;
}

[[nodiscard]] std::optional<size_t> TryMoveCaretVerticallyByWrappedLines(const WindowHost* host,
                                                                         std::wstring_view text,
                                                                         FontRole role,
                                                                         const D2D1_RECT_F& textRect,
                                                                         size_t caretIndex,
                                                                         int visualLineDelta,
                                                                         std::optional<float>& preferredXOffsetDip) noexcept
{
    if (! host || text.empty() || visualLineDelta == 0 || (textRect.right - textRect.left) <= 1.0f || (textRect.bottom - textRect.top) <= 1.0f)
    {
        return std::nullopt;
    }

    wil::com_ptr<IDWriteTextLayout> layout =
        CreateMultilineTextLayout(host, text, role, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
    if (! layout)
    {
        return std::nullopt;
    }

    const std::vector<DWRITE_LINE_METRICS> metrics = GetMultilineLineMetrics(layout.get());
    if (metrics.empty())
    {
        return std::nullopt;
    }

    float currentX = 0.0f;
    float currentY = 0.0f;
    DWRITE_HIT_TEST_METRICS currentHitMetrics{};
    if (FAILED(layout->HitTestTextPosition(static_cast<UINT32>(std::min(caretIndex, text.size())), FALSE, &currentX, &currentY, &currentHitMetrics)))
    {
        return std::nullopt;
    }

    const std::optional<size_t> currentLineIndex = TryGetMultilineCaretLineIndex(layout.get(), metrics, caretIndex, text.size());
    if (! currentLineIndex)
    {
        return std::nullopt;
    }

    const int maxLineIndex    = static_cast<int>(metrics.size()) - 1;
    const int targetLineIndex = std::clamp(static_cast<int>(currentLineIndex.value()) + visualLineDelta, 0, maxLineIndex);
    if (targetLineIndex == static_cast<int>(currentLineIndex.value()))
    {
        return caretIndex;
    }

    const float targetX = preferredXOffsetDip.value_or(currentX);
    preferredXOffsetDip = targetX;

    float targetTopDip = 0.0f;
    for (int index = 0; index < targetLineIndex; ++index)
    {
        targetTopDip += std::max(1.0f, metrics[static_cast<size_t>(index)].height);
    }

    const float targetHeightDip = std::max(1.0f, metrics[static_cast<size_t>(targetLineIndex)].height);
    const float targetY         = targetTopDip + std::min(targetHeightDip - 1.0f, targetHeightDip * 0.5f);

    BOOL isTrailingHit = FALSE;
    BOOL isInside      = FALSE;
    DWRITE_HIT_TEST_METRICS targetHitMetrics{};
    if (FAILED(layout->HitTestPoint(targetX, targetY, &isTrailingHit, &isInside, &targetHitMetrics)))
    {
        return std::nullopt;
    }

    const size_t textPosition    = static_cast<size_t>(targetHitMetrics.textPosition);
    const size_t trailingAdvance = isTrailingHit ? static_cast<size_t>(targetHitMetrics.length) : 0u;
    return std::min(text.size(), textPosition + trailingAdvance);
}

[[nodiscard]] size_t MoveMultilineCaretVertically(const WindowHost* host,
                                                  std::wstring_view text,
                                                  FontRole role,
                                                  const D2D1_RECT_F& textRect,
                                                  size_t caretIndex,
                                                  bool moveDown,
                                                  std::optional<float>& preferredXOffsetDip) noexcept
{
    if (const std::optional<size_t> wrappedCaretIndex =
            TryMoveCaretVerticallyByWrappedLines(host, text, role, textRect, caretIndex, moveDown ? 1 : -1, preferredXOffsetDip))
    {
        return wrappedCaretIndex.value();
    }

    return MoveCaretVerticallyByLogicalLine(text, caretIndex, moveDown, preferredXOffsetDip);
}

[[nodiscard]] size_t MoveMultilineCaretByPage(const WindowHost* host,
                                              std::wstring_view text,
                                              FontRole role,
                                              const D2D1_RECT_F& textRect,
                                              size_t caretIndex,
                                              bool moveDown,
                                              std::optional<float>& preferredXOffsetDip,
                                              size_t lineCount) noexcept
{
    if (lineCount == 0u)
    {
        return caretIndex;
    }

    if (const std::optional<size_t> wrappedCaretIndex = TryMoveCaretVerticallyByWrappedLines(
            host, text, role, textRect, caretIndex, moveDown ? static_cast<int>(lineCount) : -static_cast<int>(lineCount), preferredXOffsetDip))
    {
        return wrappedCaretIndex.value();
    }

    return MoveCaretByPage(text, caretIndex, moveDown, preferredXOffsetDip, lineCount);
}

[[nodiscard]] std::vector<WrappedLineTextRange> BuildWrappedLineTextRanges(const std::vector<DWRITE_LINE_METRICS>& metrics, size_t textLength) noexcept
{
    std::vector<WrappedLineTextRange> ranges;
    ranges.reserve(metrics.size());

    size_t lineStart = 0u;
    for (const DWRITE_LINE_METRICS& metric : metrics)
    {
        const size_t rawLineEnd = std::min(textLength, lineStart + static_cast<size_t>(metric.length));
        size_t visibleLineEnd   = rawLineEnd;
        if (metric.newlineLength > 0u)
        {
            const size_t newlineLength = static_cast<size_t>(metric.newlineLength);
            visibleLineEnd             = (visibleLineEnd >= newlineLength) ? visibleLineEnd - newlineLength : lineStart;
        }

        ranges.push_back({lineStart, std::max(lineStart, visibleLineEnd)});
        lineStart = rawLineEnd;
    }

    return ranges;
}

[[nodiscard]] std::optional<size_t> TryMoveCaretToWrappedLineBoundary(
    const WindowHost* host, std::wstring_view text, FontRole role, const D2D1_RECT_F& textRect, size_t caretIndex, bool moveToLineEnd) noexcept
{
    if (! host || text.empty() || (textRect.right - textRect.left) <= 1.0f || (textRect.bottom - textRect.top) <= 1.0f)
    {
        return std::nullopt;
    }

    const wil::com_ptr<IDWriteTextLayout> layout =
        CreateMultilineTextLayout(host, text, role, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
    if (! layout)
    {
        return std::nullopt;
    }

    const std::vector<DWRITE_LINE_METRICS> metrics = GetMultilineLineMetrics(layout.get());
    if (metrics.empty() || metrics.size() <= CountLogicalTextLines(text))
    {
        return std::nullopt;
    }

    const std::optional<size_t> currentLineIndex = TryGetMultilineCaretLineIndex(layout.get(), metrics, caretIndex, text.size());
    if (! currentLineIndex)
    {
        return std::nullopt;
    }

    const std::vector<WrappedLineTextRange> ranges = BuildWrappedLineTextRanges(metrics, text.size());
    if (currentLineIndex.value() >= ranges.size())
    {
        return std::nullopt;
    }

    const WrappedLineTextRange& range = ranges[currentLineIndex.value()];
    return moveToLineEnd ? range.end : range.start;
}

[[nodiscard]] wil::com_ptr<IDWriteTextLayout> CreateMultilineTextLayout(
    const WindowHost* host, std::wstring_view text, FontRole role, float widthDip, float heightDip) noexcept
{
    if (! host)
    {
        return {};
    }

    auto* factory = host->GetWriteFactory();
    auto* format  = host->GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_NEAR, true);
    if (! factory || ! format)
    {
        return {};
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(factory->CreateTextLayout(text.data(),
                                         static_cast<UINT32>(text.size()),
                                         format,
                                         std::max(1.0f, widthDip),
                                         std::max(kMultilineLayoutHeightDip, heightDip),
                                         layout.addressof())) ||
        ! layout)
    {
        return {};
    }
    return layout;
}

[[nodiscard]] std::vector<DWRITE_LINE_METRICS> GetMultilineLineMetrics(IDWriteTextLayout* layout) noexcept
{
    if (! layout)
    {
        return {};
    }

    UINT32 actualLineCount = 0u;
    HRESULT hr             = layout->GetLineMetrics(nullptr, 0u, &actualLineCount);
    if (FAILED(hr) && actualLineCount == 0u)
    {
        return {};
    }

    std::vector<DWRITE_LINE_METRICS> metrics(actualLineCount);
    if (actualLineCount == 0u)
    {
        return metrics;
    }

    if (FAILED(layout->GetLineMetrics(metrics.data(), actualLineCount, &actualLineCount)))
    {
        return {};
    }

    metrics.resize(actualLineCount);
    return metrics;
}

[[nodiscard]] float EstimateMultilineFallbackLineHeightDip(const WindowHost* host, FontRole role) noexcept
{
    if (! host)
    {
        return 20.0f;
    }

    auto* factory = host->GetWriteFactory();
    auto* format  = host->GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_NEAR, true);
    if (! factory || ! format)
    {
        return 20.0f;
    }

    static constexpr wchar_t kSampleText[] = L"Ag";
    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(factory->CreateTextLayout(
            kSampleText, static_cast<UINT32>(std::size(kSampleText) - 1u), format, 1024.0f, kMultilineLayoutHeightDip, layout.addressof())) ||
        ! layout)
    {
        return 20.0f;
    }

    const std::vector<DWRITE_LINE_METRICS> metrics = GetMultilineLineMetrics(layout.get());
    if (! metrics.empty())
    {
        return std::max(1.0f, metrics.front().height);
    }

    DWRITE_TEXT_METRICS textMetrics{};
    if (SUCCEEDED(layout->GetMetrics(&textMetrics)) && textMetrics.height > 0.0f)
    {
        return textMetrics.height;
    }

    return 20.0f;
}

[[nodiscard]] float MeasureWrappedLineOffsetDip(const std::vector<DWRITE_LINE_METRICS>& metrics, size_t firstVisibleLine) noexcept
{
    if (metrics.empty() || firstVisibleLine == 0u)
    {
        return 0.0f;
    }

    const size_t clampedFirstVisibleLine = std::min(firstVisibleLine, metrics.size());
    float offsetDip                      = 0.0f;
    for (size_t index = 0u; index < clampedFirstVisibleLine; ++index)
    {
        offsetDip += metrics[index].height;
    }
    return offsetDip;
}

[[nodiscard]] size_t CountVisibleWrappedLines(const std::vector<DWRITE_LINE_METRICS>& metrics, float viewportHeightDip, float fallbackLineHeightDip) noexcept
{
    if (metrics.empty())
    {
        return ComputeMultilinePageLineCount(viewportHeightDip, fallbackLineHeightDip);
    }

    const float clampedViewportHeightDip = std::max(1.0f, viewportHeightDip);
    float consumedHeightDip              = 0.0f;
    size_t visibleLineCount              = 0u;
    for (const DWRITE_LINE_METRICS& metric : metrics)
    {
        ++visibleLineCount;
        consumedHeightDip += std::max(1.0f, metric.height);
        if (consumedHeightDip >= clampedViewportHeightDip)
        {
            break;
        }
    }

    return std::max<size_t>(1u, visibleLineCount);
}

[[nodiscard]] MultilineViewportMetrics BuildMultilineViewportMetrics(std::wstring_view text,
                                                                     float viewportHeightDip,
                                                                     const std::vector<DWRITE_LINE_METRICS>& metrics,
                                                                     float fallbackLineHeightDip) noexcept
{
    MultilineViewportMetrics viewportMetrics;
    if (metrics.empty())
    {
        viewportMetrics.totalLineCount   = CountLogicalTextLines(text);
        viewportMetrics.visibleLineCount = ComputeMultilinePageLineCount(viewportHeightDip, fallbackLineHeightDip);
        return viewportMetrics;
    }

    viewportMetrics.totalLineCount   = std::max<size_t>(1u, metrics.size());
    viewportMetrics.visibleLineCount = CountVisibleWrappedLines(metrics, viewportHeightDip, fallbackLineHeightDip);
    return viewportMetrics;
}

[[nodiscard]] size_t ComputeMultilineMaxFirstVisibleLine(const MultilineViewportMetrics& viewportMetrics) noexcept
{
    // Preserve imported/view-driven top-line state up to the last content line.
    // The DxUI surface can legitimately show trailing blank space below the last line.
    return viewportMetrics.totalLineCount > 0u ? viewportMetrics.totalLineCount - 1u : 0u;
}

[[nodiscard]] size_t ClampMultilineFirstVisibleLine(size_t firstVisibleLine, const MultilineViewportMetrics& viewportMetrics) noexcept
{
    return (std::min)(firstVisibleLine, ComputeMultilineMaxFirstVisibleLine(viewportMetrics));
}

[[nodiscard]] std::optional<size_t> TryGetMultilineCaretLineIndex(IDWriteTextLayout* layout,
                                                                  const std::vector<DWRITE_LINE_METRICS>& metrics,
                                                                  size_t caretIndex,
                                                                  size_t textLength) noexcept
{
    if (! layout || metrics.empty())
    {
        return std::nullopt;
    }

    float x = 0.0f;
    float y = 0.0f;
    DWRITE_HIT_TEST_METRICS hitMetrics{};
    if (FAILED(layout->HitTestTextPosition(static_cast<UINT32>(std::min(caretIndex, textLength)), FALSE, &x, &y, &hitMetrics)))
    {
        return std::nullopt;
    }

    float lineTopDip = 0.0f;
    for (size_t index = 0u; index < metrics.size(); ++index)
    {
        lineTopDip += std::max(1.0f, metrics[index].height);
        if (y < lineTopDip || index + 1u == metrics.size())
        {
            return index;
        }
    }

    return std::nullopt;
}

[[nodiscard]] MultilineViewportMetrics ComputeMultilineViewportMetrics(const WindowHost* host,
                                                                       std::wstring_view text,
                                                                       FontRole role,
                                                                       const D2D1_RECT_F& textRect) noexcept
{
    const float viewportHeightDip     = std::max(1.0f, textRect.bottom - textRect.top);
    const auto multilineLayout        = CreateMultilineTextLayout(host, text, role, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
    const float fallbackLineHeightDip = EstimateMultilineFallbackLineHeightDip(host, role);
    return BuildMultilineViewportMetrics(text, viewportHeightDip, GetMultilineLineMetrics(multilineLayout.get()), fallbackLineHeightDip);
}

[[nodiscard]] size_t HitTestMultilineCaretIndexDip(
    const WindowHost* host, std::wstring_view text, FontRole role, const D2D1_RECT_F& textRect, float scrollDip, D2D1_POINT_2F point) noexcept
{
    if (text.empty())
    {
        return 0u;
    }

    const float localX = std::max(0.0f, point.x - textRect.left);
    const float localY = std::max(0.0f, point.y - textRect.top + scrollDip);
    if (wil::com_ptr<IDWriteTextLayout> layout =
            CreateMultilineTextLayout(host, text, role, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top)))
    {
        BOOL isTrailingHit = FALSE;
        BOOL isInside      = FALSE;
        DWRITE_HIT_TEST_METRICS metrics{};
        if (SUCCEEDED(layout->HitTestPoint(localX, localY, &isTrailingHit, &isInside, &metrics)))
        {
            const size_t textPosition    = static_cast<size_t>(metrics.textPosition);
            const size_t trailingAdvance = isTrailingHit ? static_cast<size_t>(metrics.length) : 0u;
            return std::min(text.size(), textPosition + trailingAdvance);
        }
    }

    size_t lineStart = 0u;
    size_t lineIndex = static_cast<size_t>(std::max(0.0, std::floor(static_cast<double>(localY) / 18.0)));
    while (lineIndex > 0u && lineStart < text.size())
    {
        const size_t nextBreak = text.find(L'\n', lineStart);
        if (nextBreak == std::wstring_view::npos)
        {
            return text.size();
        }
        lineStart = nextBreak + 1u;
        --lineIndex;
    }

    const size_t lineEnd        = text.find(L'\n', lineStart);
    const size_t clampedLineEnd = lineEnd == std::wstring_view::npos ? text.size() : lineEnd;
    const size_t column = static_cast<size_t>(std::clamp(std::floor(static_cast<double>(localX) / 7.0), 0.0, static_cast<double>(clampedLineEnd - lineStart)));
    return std::min(text.size(), lineStart + column);
}

[[nodiscard]] D2D1_RECT_F MeasureMultilineCaretRectDip(
    const WindowHost* host, std::wstring_view text, FontRole role, const D2D1_RECT_F& textRect, float scrollDip, size_t caretIndex) noexcept
{
    const D2D1_RECT_F fallbackRect = D2D1::RectF(textRect.left, textRect.top + 2.0f, textRect.left + 1.0f, textRect.top + 18.0f);
    if (wil::com_ptr<IDWriteTextLayout> layout =
            CreateMultilineTextLayout(host, text, role, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top)))
    {
        float x = 0.0f;
        float y = 0.0f;
        DWRITE_HIT_TEST_METRICS metrics{};
        if (SUCCEEDED(layout->HitTestTextPosition(static_cast<UINT32>(std::min(caretIndex, text.size())), FALSE, &x, &y, &metrics)))
        {
            const float caretTop    = textRect.top + y - scrollDip;
            const float caretHeight = std::max(12.0f, metrics.height);
            return D2D1::RectF(textRect.left + x, caretTop, textRect.left + x + 1.0f, caretTop + caretHeight);
        }
    }

    return fallbackRect;
}

[[nodiscard]] D2D1_RECT_F ClipTextInputRectToBounds(D2D1_RECT_F rect, const D2D1_RECT_F& bounds) noexcept
{
    rect.left   = std::clamp(rect.left, bounds.left, bounds.right);
    rect.right  = std::clamp(rect.right, rect.left, bounds.right);
    rect.top    = std::clamp(rect.top, bounds.top, bounds.bottom);
    rect.bottom = std::clamp(rect.bottom, rect.top, bounds.bottom);
    return rect;
}

void DrawMultilineSelection(WindowHost& host,
                            std::wstring_view text,
                            const D2D1_RECT_F& rect,
                            FontRole role,
                            const D2D1_COLOR_F& textColor,
                            const D2D1_COLOR_F& selectionFill,
                            const D2D1_COLOR_F& selectionText,
                            float scrollDip,
                            std::optional<std::pair<size_t, size_t>> selectionRange) noexcept
{
    auto* dc                = host.GetDeviceContext();
    auto* fillBrush         = host.GetSolidBrush(selectionFill);
    auto* textBrush         = host.GetSolidBrush(textColor);
    auto* selectedTextBrush = host.GetSolidBrush(selectionText);
    if (! dc || ! textBrush)
    {
        return;
    }

    const D2D1_RECT_F snappedRect          = SnapRectToPixel(host, rect);
    wil::com_ptr<IDWriteTextLayout> layout = CreateMultilineTextLayout(
        &host, text, role, std::max(1.0f, snappedRect.right - snappedRect.left), std::max(1.0f, snappedRect.bottom - snappedRect.top));
    if (! layout)
    {
        DrawCenteredText(host, text, snappedRect, role, textColor, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_NEAR, true);
        return;
    }

    dc->PushAxisAlignedClip(snappedRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
    dc->DrawTextLayout(D2D1::Point2F(snappedRect.left, snappedRect.top - scrollDip), layout.get(), textBrush, kTextDrawOptions);

    if (selectionRange && fillBrush && selectedTextBrush)
    {
        const auto [selectionStart, selectionEnd] = selectionRange.value();
        if (selectionStart < selectionEnd)
        {
            UINT32 actualCount = 0u;
            std::vector<DWRITE_HIT_TEST_METRICS> metrics(static_cast<size_t>(selectionEnd - selectionStart) + 4u);
            HRESULT hr = layout->HitTestTextRange(static_cast<UINT32>(selectionStart),
                                                  static_cast<UINT32>(selectionEnd - selectionStart),
                                                  snappedRect.left,
                                                  snappedRect.top - scrollDip,
                                                  metrics.data(),
                                                  static_cast<UINT32>(metrics.size()),
                                                  &actualCount);
            if (hr == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER))
            {
                metrics.resize(actualCount);
                hr = layout->HitTestTextRange(static_cast<UINT32>(selectionStart),
                                              static_cast<UINT32>(selectionEnd - selectionStart),
                                              snappedRect.left,
                                              snappedRect.top - scrollDip,
                                              metrics.data(),
                                              static_cast<UINT32>(metrics.size()),
                                              &actualCount);
            }

            if (SUCCEEDED(hr))
            {
                for (UINT32 index = 0u; index < actualCount; ++index)
                {
                    const auto& hit           = metrics[index];
                    const D2D1_RECT_F hitRect = SnapRectToPixel(host, D2D1::RectF(hit.left, hit.top + 1.0f, hit.left + hit.width, hit.top + hit.height - 1.0f));
                    if (hitRect.right > hitRect.left)
                    {
                        dc->FillRectangle(hitRect, fillBrush);
                    }
                }

                for (UINT32 index = 0u; index < actualCount; ++index)
                {
                    const auto& hit           = metrics[index];
                    const D2D1_RECT_F hitRect = SnapRectToPixel(host, D2D1::RectF(hit.left, hit.top, hit.left + hit.width, hit.top + hit.height));
                    if (hitRect.right <= hitRect.left)
                    {
                        continue;
                    }

                    dc->PushAxisAlignedClip(hitRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
                    dc->DrawTextLayout(D2D1::Point2F(snappedRect.left, snappedRect.top - scrollDip), layout.get(), selectedTextBrush, kTextDrawOptions);
                    dc->PopAxisAlignedClip();
                }
            }
        }
    }

    dc->PopAxisAlignedClip();
}

} // namespace

void WindowHost::CommitFocusedTextInput() noexcept
{
    SyncNativeTextInputSession(_focusedControl);
}

bool WindowHost::TryReadTextInputState(const Control* control, TextInputState& outState) const noexcept
{
    if (! control || control != _nativeTextInputControl || ! _nativeTextInputStateCacheValid)
    {
        return false;
    }

    outState.text                 = _nativeTextInputStateCache.text;
    outState.selectionAnchorIndex = _nativeTextInputStateCache.selectionAnchorIndex;
    outState.caretIndex           = _nativeTextInputStateCache.caretIndex;
    outState.firstVisibleLine     = _nativeTextInputStateCache.firstVisibleLine;
    outState.readOnly             = _nativeTextInputStateCache.readOnly;
    outState.masked               = _nativeTextInputStateCache.masked;
    outState.multiline            = _nativeTextInputStateCache.multiline;
    return true;
}

void WindowHost::SyncTextInput(Control* control) noexcept
{
    SyncNativeTextInputSession(control);
}

bool WindowHost::HasActiveTextInput() const noexcept
{
    return HasActiveNativeTextInputSession();
}

HWND WindowHost::GetTextInputHwnd() const noexcept
{
    return HasActiveNativeTextInputSession() ? _hwnd : nullptr;
}

TextField::TextField(std::wstring text) : _text(std::move(text))
{
    _caretIndex = _text.size();
    SetFocusable(true);
}

TextField::~TextField() noexcept
{
    SecureClearStorage();
}

void TextField::SecureClearStorage() noexcept
{
    SecureWipe::SecureClear(_text);
    for (EditHistoryState& state : _undoHistory)
    {
        SecureWipe::SecureClear(state.text);
    }
    for (EditHistoryState& state : _redoHistory)
    {
        SecureWipe::SecureClear(state.text);
    }
    SecureWipe::SecureClear(_cachedLayoutText);
    ClearSingleLineTextLayoutCache(_singleLineLayoutCache, true);
    _maskedSourceTextElementBoundaries.clear();
    _undoHistory.clear();
    _redoHistory.clear();
}

void TextField::RefreshAccessibilitySnapshot() const noexcept
{
    if (WindowHost* const host = GetHost())
    {
        RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
    }
}

void TextField::SetText(std::wstring text)
{
    _text       = std::move(text);
    _caretIndex = std::min(_caretIndex, _text.size());
    _selectionAnchorIndex.reset();
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    _preferredMultilineXOffsetDip.reset();
    _caretBlinkAnchorTickMs       = 0u;
    _caretVisible                 = true;
    _horizontalScrollDip          = 0.0f;
    _multilineFirstVisibleLine    = 0u;
    _multilineWheelDeltaRemainder = 0.0f;
    _dragSelecting                = false;
    _undoHistory.clear();
    _redoHistory.clear();
    BreakDirectEditMerge();
    RegenerateConcealedMaskEpoch();
    InvalidateSingleLineLayoutCache();
    InvalidateMultilineLayoutCache();
    if (auto* host = GetHost(); host && host->GetFocusControl() == this)
    {
        host->SyncTextInput(this);
    }
    RefreshAccessibilitySnapshot();
    RequestInvalidate();
}

std::wstring_view TextField::GetText() const noexcept
{
    return _text;
}

size_t TextField::GetCaretIndex() const noexcept
{
    return _caretIndex;
}

void TextField::SetSelectionRange(const size_t selectionStart, const size_t selectionEnd) noexcept
{
    BreakDirectEditMerge();
    const size_t clampedStart = std::min(selectionStart, _text.size());
    const size_t clampedEnd   = std::min(selectionEnd, _text.size());
    _preferredMultilineXOffsetDip.reset();
    _dragSelecting = false;
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    if (clampedStart == clampedEnd)
    {
        _selectionAnchorIndex.reset();
        _caretIndex = clampedEnd;
    }
    else
    {
        _selectionAnchorIndex = clampedStart;
        _caretIndex           = clampedEnd;
    }
    _caretBlinkAnchorTickMs = 0u;
    _caretVisible           = true;
    if (auto* host = GetHost(); host && host->GetFocusControl() == this)
    {
        host->SyncTextInput(this);
    }
    RefreshAccessibilitySnapshot();
    RequestInvalidate();
}

void TextField::ReplaceSelectionAndNotify(std::wstring_view replacement)
{
    const bool willMutate = HasSelection() || ! replacement.empty();
    if (! willMutate)
    {
        return;
    }

    RecordUndoStateForDirectEdit();
    _preferredMultilineXOffsetDip.reset();
    static_cast<void>(DeleteSelection());
    _text.insert(_caretIndex, replacement.data(), replacement.size());
    _caretIndex += replacement.size();
    _selectionAnchorIndex.reset();
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    _caretBlinkAnchorTickMs = 0u;
    _caretVisible           = true;
    InvalidateSingleLineLayoutCache();
    InvalidateMultilineLayoutCache();
    if (auto* host = GetHost(); host && host->GetFocusControl() == this)
    {
        host->SyncTextInput(this);
    }
    RefreshAccessibilitySnapshot();
    RequestInvalidate();
    static_cast<void>(NotifyChanged());
}

void TextField::SetTextAndNotify(std::wstring text)
{
    SetText(std::move(text));
    static_cast<void>(NotifyChanged());
}

void TextField::SetMasked(bool masked) noexcept
{
    if (_masked == masked)
    {
        return;
    }

    _masked = masked;
    RegenerateConcealedMaskEpoch();
    InvalidateSingleLineLayoutCache();
    if (WindowHost* const host = GetHost(); host && HasFocus())
    {
        host->SyncTextInput(this);
    }
    RefreshAccessibilitySnapshot();
    RequestInvalidate();
}

bool TextField::IsMasked() const noexcept
{
    return _masked;
}

void TextField::SetPasswordRevealMode(PasswordRevealMode mode) noexcept
{
    if (_passwordRevealMode == mode)
    {
        return;
    }

    _passwordRevealMode            = mode;
    _passwordRevealButtonHovered   = false;
    _passwordRevealButtonPressed   = false;
    _passwordRevealKeyboardFocused = false;

    if (_passwordRevealMode == PasswordRevealMode::Visible)
    {
        SetPasswordRevealState(PasswordRevealState::Visible);
    }
    else if (_passwordRevealMode == PasswordRevealMode::Hidden)
    {
        SetPasswordRevealState(PasswordRevealState::Hidden);
    }
    else if (_passwordRevealState == PasswordRevealState::Visible)
    {
        SetPasswordRevealState(PasswordRevealState::Hidden);
    }

    InvalidateSingleLineLayoutCache();
    RefreshAccessibilitySnapshot();
    RequestInvalidate();
}

PasswordRevealMode TextField::GetPasswordRevealMode() const noexcept
{
    return _passwordRevealMode;
}

void TextField::SetPasswordRevealState(PasswordRevealState state) noexcept
{
    if (_passwordRevealMode == PasswordRevealMode::Visible)
    {
        state = PasswordRevealState::Visible;
    }
    else if (_passwordRevealMode == PasswordRevealMode::Hidden)
    {
        state = PasswordRevealState::Hidden;
    }

    if (_passwordRevealState == state)
    {
        return;
    }

    const auto startedAt = std::chrono::steady_clock::now();
    _passwordRevealState = state;
    InvalidateSingleLineLayoutCache();
    RefreshAccessibilitySnapshot();
    RequestInvalidate();
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(
            L"dxui.textinput.reveal_toggle_us", L"textfield", Debug::Perf::ElapsedUs(startedAt), static_cast<uint64_t>(state), _masked ? 1u : 0u, S_OK);
    }
}

PasswordRevealState TextField::GetPasswordRevealState() const noexcept
{
    return _passwordRevealState;
}

void TextField::SetPasswordMaskLengthPolicy(PasswordMaskLengthPolicy policy) noexcept
{
    if (_maskLengthPolicy == policy)
    {
        return;
    }

    _maskLengthPolicy = policy;
    RegenerateConcealedMaskEpoch();
    InvalidateSingleLineLayoutCache();
    RequestInvalidate();
}

PasswordMaskLengthPolicy TextField::GetPasswordMaskLengthPolicy() const noexcept
{
    return _maskLengthPolicy;
}

size_t TextField::GetSecretVisibleDotCount() const noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    size_t result        = 0u;
    size_t exactCount    = 0u;
    if (! _masked || _text.empty() || _passwordRevealMode == PasswordRevealMode::Visible || _passwordRevealState == PasswordRevealState::Visible)
    {
        return result;
    }

    if (_maskLengthPolicy == PasswordMaskLengthPolicy::Exact)
    {
        exactCount = GetMaskedSourceTextElementBoundaries().size() - 1u;
        result     = exactCount;
    }
    else
    {
        exactCount = _text.size();
        result     = GetConcealedMaskVisibleDotCount(exactCount);
    }

    if (Debug::Perf::IsCaptureEnabled())
    {
        const std::wstring_view policyDetail = _maskLengthPolicy == PasswordMaskLengthPolicy::Exact ? L"exact" : L"concealed";
        Debug::Perf::Emit(L"dxui.textinput.secret_render_us", policyDetail, Debug::Perf::ElapsedUs(startedAt), _text.size(), result, S_OK);
        Debug::Perf::Emit(L"dxui.textinput.secret_display_dot_count", policyDetail, 0u, result, exactCount, S_OK);
    }
    return result;
}

void TextField::SetPasswordRevealAccessibleName(std::wstring name)
{
    _passwordRevealAccessibleName = std::move(name);
}

std::wstring_view TextField::GetPasswordRevealAccessibleName() const noexcept
{
    return _passwordRevealAccessibleName;
}

bool TextField::IsPasswordRevealButtonVisibleForAccessibility() const noexcept
{
    return IsPasswordRevealButtonVisible();
}

D2D1_RECT_F TextField::GetPasswordRevealButtonAccessibilityRect() const noexcept
{
    return GetPasswordRevealButtonRect();
}

bool TextField::InvokePasswordRevealButton(WindowHost& host)
{
    if (! IsPasswordRevealButtonVisible())
    {
        return false;
    }

    _passwordRevealButtonHovered   = false;
    _passwordRevealButtonPressed   = false;
    _passwordRevealKeyboardFocused = false;
    if (_passwordRevealState == PasswordRevealState::Visible)
    {
        RemaskPasswordReveal();
    }
    else
    {
        SetPasswordRevealState(PasswordRevealState::Visible);
    }
    ResetCaretBlink(host);
    host.SetFocusControl(this);
    host.SyncTextInput(this);
    Invalidate(host);
    return true;
}

void TextField::SetPlaceholder(std::wstring text)
{
    _placeholder = std::move(text);
}

void TextField::SetMultiline(bool multiline) noexcept
{
    _multiline = multiline;
    _preferredMultilineXOffsetDip.reset();
    _multilineFirstVisibleLine    = 0u;
    _multilineWheelDeltaRemainder = 0.0f;
    InvalidateSingleLineLayoutCache();
    InvalidateMultilineLayoutCache();
}

bool TextField::IsMultiline() const noexcept
{
    return _multiline;
}

void TextField::SetClearButtonEnabled(bool enabled) noexcept
{
    _clearButtonEnabled = enabled;
    if (! enabled)
    {
        _clearButtonHovered = false;
    }
}

bool TextField::IsClearButtonEnabled() const noexcept
{
    return _clearButtonEnabled;
}

void TextField::SetCaretColor(std::optional<D2D1_COLOR_F> caretColor) noexcept
{
    _caretColorOverride = caretColor;
}

void TextField::SetHorizontalTextPadding(float leftDip, float rightDip) noexcept
{
    const float clampedLeftDip  = std::max(0.0f, leftDip);
    const float clampedRightDip = std::max(0.0f, rightDip);
    if (_hasExplicitHorizontalTextPadding && _textPaddingLeftDip == clampedLeftDip && _textPaddingRightDip == clampedRightDip)
    {
        return;
    }

    _hasExplicitHorizontalTextPadding = true;
    _textPaddingLeftDip               = clampedLeftDip;
    _textPaddingRightDip              = clampedRightDip;
    InvalidateSingleLineLayoutCache();
    InvalidateMultilineLayoutCache();
    if (auto* host = GetHost(); host && host->GetFocusControl() == this)
    {
        host->SyncTextInput(this);
    }
    RequestInvalidate();
}

void TextField::SetVerticalTextPadding(float topDip, float bottomDip) noexcept
{
    const float clampedTopDip    = std::max(0.0f, topDip);
    const float clampedBottomDip = std::max(0.0f, bottomDip);
    if (_hasExplicitVerticalTextPadding && _textPaddingTopDip == clampedTopDip && _textPaddingBottomDip == clampedBottomDip)
    {
        return;
    }

    _hasExplicitVerticalTextPadding = true;
    _textPaddingTopDip              = clampedTopDip;
    _textPaddingBottomDip           = clampedBottomDip;
    InvalidateSingleLineLayoutCache();
    InvalidateMultilineLayoutCache();
    if (auto* host = GetHost(); host && host->GetFocusControl() == this)
    {
        host->SyncTextInput(this);
    }
    RequestInvalidate();
}

void TextField::SetReadOnly(bool readOnly) noexcept
{
    if (_readOnly == readOnly)
    {
        return;
    }

    _readOnly = readOnly;
    if (_readOnly)
    {
        _passwordRevealButtonHovered   = false;
        _passwordRevealButtonPressed   = false;
        _passwordRevealKeyboardFocused = false;
        RemaskPasswordReveal();
    }
    if (WindowHost* const host = GetHost(); host && HasFocus())
    {
        host->SyncTextInput(this);
    }
    RequestInvalidate();
}

bool TextField::IsReadOnly() const noexcept
{
    return _readOnly;
}

void TextField::SetOnTextChanged(std::function<void(std::wstring_view)> onTextChanged)
{
    _onTextChanged = std::move(onTextChanged);
}

void TextField::SetOnSubmitted(std::function<void()> onSubmitted)
{
    _onSubmitted = std::move(onSubmitted);
}

void TextField::SetOnPreviewKeyDown(std::function<bool(WindowHost& host, UINT virtualKey, UINT modifiers)> onPreviewKeyDown)
{
    _onPreviewKeyDown = std::move(onPreviewKeyDown);
}

void TextField::SetOnBlur(std::function<void()> onBlur)
{
    _onBlur = std::move(onBlur);
}

bool TextField::DebugGetMultilineState(const WindowHost& host, TextFieldDebugMultilineState& out) const noexcept
{
    out = {};
    if (! _multiline)
    {
        return false;
    }

    const D2D1_RECT_F textRect     = GetTextRect();
    const float viewportHeightDip  = std::max(1.0f, textRect.bottom - textRect.top);
    const std::wstring displayText = GetDisplayText();
    const auto multilineLayout =
        CreateMultilineTextLayout(&host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
    const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
    const float fallbackLineHeightDip                  = EstimateMultilineFallbackLineHeightDip(&host, FontRole::Body);
    const MultilineViewportMetrics viewportMetrics     = BuildMultilineViewportMetrics(displayText, viewportHeightDip, lineMetrics, fallbackLineHeightDip);
    const size_t clampedFirstVisibleLine               = ClampMultilineFirstVisibleLine(_multilineFirstVisibleLine, viewportMetrics);

    out.firstVisibleLine      = clampedFirstVisibleLine;
    out.visibleLineCount      = viewportMetrics.visibleLineCount;
    out.totalLineCount        = viewportMetrics.totalLineCount;
    out.canScrollVertically   = viewportMetrics.totalLineCount > viewportMetrics.visibleLineCount;
    out.cachedLayoutPresent   = static_cast<bool>(_cachedMultilineLayout);
    out.layoutDirty           = _multilineLayoutDirty;
    out.cachedLayoutWidthDip  = _cachedLayoutSize.width;
    out.cachedLayoutHeightDip = _cachedLayoutSize.height;
    return true;
}

bool TextField::DebugGetSingleLinePaintState(const WindowHost& host, TextFieldDebugSingleLinePaintState& out) const noexcept
{
    out = {};
    if (_multiline)
    {
        return false;
    }

    const D2D1_RECT_F textRect = GetTextRect();
    EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));

    out.textRect            = SnapRectToPixel(host, textRect);
    out.horizontalScrollDip = _horizontalScrollDip;
    if (IsClearButtonVisible())
    {
        out.trailingButtonRect    = SnapRectToPixel(host, GetClearButtonRect());
        out.hasTrailingButtonRect = true;
    }
    else if (IsPasswordRevealButtonVisible())
    {
        out.trailingButtonRect    = SnapRectToPixel(host, GetPasswordRevealButtonRect());
        out.hasTrailingButtonRect = true;
    }
    const DWRITE_READING_DIRECTION readingDirection               = ResolveReadingDirection(GetFlowDirection());
    const std::optional<std::pair<size_t, size_t>> selectionRange = GetSelectionRange();
    const std::optional<std::pair<size_t, size_t>> displaySelectionRange =
        selectionRange.has_value() ? ControlTextRangeToDisplayTextRange(selectionRange.value().first, selectionRange.value().second) : std::nullopt;
    if (const std::optional<D2D1_RECT_F> selectionPaintRect = ComputeSingleLineSelectionPaintRect(
            host, GetDisplayText(), textRect, FontRole::Body, _horizontalScrollDip, displaySelectionRange, readingDirection);
        selectionPaintRect.has_value())
    {
        out.selectionPaintRect    = selectionPaintRect.value();
        out.hasSelectionPaintRect = true;
    }
    out.compositionUnderlineRects      = BuildNativeCompositionUnderlineRects(host, *this, false);
    out.conversionTargetUnderlineRects = BuildNativeCompositionUnderlineRects(host, *this, true);
    return true;
}

bool TextField::DebugGetCaretRect(const WindowHost& host, size_t controlTextIndex, D2D1_RECT_F& outRect) const noexcept
{
    outRect                                    = {};
    const std::optional<D2D1_RECT_F> caretRect = GetTextInputCaretRect(host, controlTextIndex);
    if (! caretRect.has_value())
    {
        return false;
    }

    outRect = caretRect.value();
    return true;
}

void TextField::OnBoundsChanged() noexcept
{
    InvalidateSingleLineLayoutCache();
    InvalidateMultilineLayoutCache();
    if (! _multiline)
    {
        _horizontalScrollDip = 0.0f;
        if (WindowHost* const host = GetHost(); host && HasFocus())
        {
            host->SyncTextInput(this);
        }
        return;
    }

    WindowHost* const host = GetHost();
    if (! host)
    {
        _multilineFirstVisibleLine    = 0u;
        _multilineWheelDeltaRemainder = 0.0f;
        return;
    }

    const D2D1_RECT_F textRect     = GetTextRect();
    const std::wstring displayText = GetDisplayText();
    if (displayText.empty())
    {
        _multilineFirstVisibleLine    = 0u;
        _multilineWheelDeltaRemainder = 0.0f;
        return;
    }

    const float viewportHeightDip = std::max(1.0f, textRect.bottom - textRect.top);
    const auto multilineLayout =
        CreateMultilineTextLayout(host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
    const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
    const float fallbackLineHeightDip                  = EstimateMultilineFallbackLineHeightDip(host, FontRole::Body);
    const MultilineViewportMetrics viewportMetrics     = BuildMultilineViewportMetrics(displayText, viewportHeightDip, lineMetrics, fallbackLineHeightDip);
    _multilineFirstVisibleLine                         = ClampMultilineFirstVisibleLine(_multilineFirstVisibleLine, viewportMetrics);
    if (HasFocus() && ! _readOnly)
    {
        EnsureMultilineCaretVisible(host);
        host->SyncTextInput(this);
    }
}

void TextField::Paint(WindowHost& host) const
{
    const TextFieldVisualStyle style =
        ResolveTextFieldVisualStyle(host.GetTheme(), IsEnabled(), IsHovered(), HasFocus(), HasFocus() && host.IsKeyboardFocusVisible(), _caretColorOverride);
    constexpr float kTextFieldCornerRadiusDip = 4.0f;
    DrawRoundedRect(host, GetBounds(), style.fill, style.border, kTextFieldCornerRadiusDip);
    if (style.showFocus)
    {
        // WinUI accent bottom border: 2px accent line at bottom of control when focused
        // Inset horizontally by corner radius to avoid clipping into rounded corners
        if (auto* dc = host.GetDeviceContext())
        {
            const D2D1_RECT_F bounds = GetBounds();
            const D2D1_RECT_F accentBar =
                D2D1::RectF(bounds.left + kTextFieldCornerRadiusDip, bounds.bottom - 2.0f, bounds.right - kTextFieldCornerRadiusDip, bounds.bottom);
            dc->FillRectangle(&accentBar, host.GetSolidBrush(host.GetTheme().accent));
        }
    }

    const D2D1_RECT_F textRect                                    = GetTextRect();
    const std::wstring displayText                                = GetDisplayText();
    const bool usePlaceholder                                     = _text.empty() && ! _placeholder.empty();
    const DWRITE_READING_DIRECTION readingDirection               = ResolveReadingDirection(GetFlowDirection());
    const std::optional<std::pair<size_t, size_t>> selectionRange = GetSelectionRange();
    const std::optional<std::pair<size_t, size_t>> displaySelectionRange =
        selectionRange.has_value() ? ControlTextRangeToDisplayTextRange(selectionRange.value().first, selectionRange.value().second) : std::nullopt;
    wil::com_ptr<IDWriteTextLayout> paintSingleLineLayout;
    if (_multiline)
    {
        const auto multilineLayout =
            GetOrCreateMultilineLayout(&host, displayText, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
        const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
        const float multilineScrollDip                     = MeasureWrappedLineOffsetDip(lineMetrics, _multilineFirstVisibleLine);
        DrawMultilineSelection(
            host, displayText, textRect, FontRole::Body, style.text, style.selectionFill, style.selectionText, multilineScrollDip, displaySelectionRange);
    }
    else if (usePlaceholder)
    {
        DrawSingleLineTextClipped(host, std::wstring_view(_placeholder), textRect, FontRole::Body, style.placeholderText, 0.0f, readingDirection);
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
        paintSingleLineLayout = GetOrCreateSingleLineLayout(&host,
                                                            displayText,
                                                            std::max(1.0f, textRect.right - textRect.left + _horizontalScrollDip),
                                                            std::max(1.0f, textRect.bottom - textRect.top),
                                                            readingDirection);
        DrawSingleLineSelectionWithLayout(host,
                                          displayText,
                                          textRect,
                                          FontRole::Body,
                                          style.text,
                                          style.selectionFill,
                                          style.selectionText,
                                          _horizontalScrollDip,
                                          displaySelectionRange,
                                          readingDirection,
                                          paintSingleLineLayout.get());
    }

    if (! usePlaceholder)
    {
        DrawTextInputUnderlineRects(host, BuildNativeCompositionUnderlineRects(host, *this, false), style.text);
        DrawTextInputUnderlineRects(host, BuildNativeCompositionUnderlineRects(host, *this, true), host.GetTheme().accent);
    }

    if (HasFocus() && _caretVisible)
    {
        D2D1_RECT_F caretRect = D2D1::RectF();
        if (_multiline)
        {
            const auto multilineLayout =
                GetOrCreateMultilineLayout(&host, displayText, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
            const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
            const float multilineScrollDip                     = MeasureWrappedLineOffsetDip(lineMetrics, _multilineFirstVisibleLine);
            caretRect =
                MeasureMultilineCaretRectDip(&host, displayText, FontRole::Body, textRect, multilineScrollDip, ControlTextIndexToDisplayTextIndex(_caretIndex));
        }
        else
        {
            if (! paintSingleLineLayout)
            {
                paintSingleLineLayout = GetOrCreateSingleLineLayout(&host,
                                                                    displayText,
                                                                    std::max(1.0f, textRect.right - textRect.left + _horizontalScrollDip),
                                                                    std::max(1.0f, textRect.bottom - textRect.top),
                                                                    readingDirection);
            }
            const float caretOffset = MeasureCaretOffsetDip(paintSingleLineLayout.get(), displayText, ControlTextIndexToDisplayTextIndex(_caretIndex));
            const float caretMaxX   = std::max(textRect.left, textRect.right - 1.0f);
            const float caretX      = std::clamp(textRect.left + caretOffset - _horizontalScrollDip, textRect.left, caretMaxX);
            caretRect               = D2D1::RectF(caretX, textRect.top + 2.0f, caretX + 1.0f, textRect.bottom - 2.0f);
        }
        if (auto* dc = host.GetDeviceContext())
        {
            if (auto* brush = host.GetSolidBrush(style.caret))
            {
                const D2D1_RECT_F snappedCaretRect = SnapRectToPixel(host, caretRect);
                dc->DrawLine(
                    D2D1::Point2F(snappedCaretRect.left, snappedCaretRect.top), D2D1::Point2F(snappedCaretRect.left, snappedCaretRect.bottom), brush, 1.0f);
            }
        }
    }

    if (IsPasswordRevealButtonVisible())
    {
        const D2D1_RECT_F revealRect = GetPasswordRevealButtonRect();
        if (_passwordRevealButtonHovered || _passwordRevealButtonPressed || _passwordRevealKeyboardFocused)
        {
            const D2D1_ROUNDED_RECT hoverBg =
                D2D1::RoundedRect(D2D1::RectF(revealRect.left + 4.0f, revealRect.top + 4.0f, revealRect.right - 4.0f, revealRect.bottom - 4.0f), 4.0f, 4.0f);
            if (auto* dc = host.GetDeviceContext())
            {
                dc->FillRoundedRectangle(&hoverBg, host.GetSolidBrush(_passwordRevealButtonPressed ? host.GetTheme().pressedFill : host.GetTheme().hoverFill));
            }
        }
        DrawCenteredText(host, L"\xE7B3", revealRect, FontRole::Icon, style.text);
        if (_passwordRevealKeyboardFocused)
        {
            PaintFocusRing(host, D2D1::RectF(revealRect.left + 3.0f, revealRect.top + 3.0f, revealRect.right - 3.0f, revealRect.bottom - 3.0f), 4.0f);
        }
    }
    if (IsClearButtonVisible())
    {
        const D2D1_RECT_F clearRect = GetClearButtonRect();
        if (_clearButtonHovered)
        {
            const D2D1_ROUNDED_RECT hoverBg =
                D2D1::RoundedRect(D2D1::RectF(clearRect.left + 4.0f, clearRect.top + 4.0f, clearRect.right - 4.0f, clearRect.bottom - 4.0f), 4.0f, 4.0f);
            if (auto* dc = host.GetDeviceContext())
            {
                dc->FillRoundedRectangle(&hoverBg, host.GetSolidBrush(host.GetTheme().hoverFill));
            }
        }
        DrawCenteredText(host, L"\xE711", clearRect, FontRole::Icon, style.text);
    }
}

bool TextField::Tick(WindowHost& /*host*/, uint64_t nowTickMs)
{
    if (! HasFocus())
    {
        _caretVisible           = true;
        _caretBlinkAnchorTickMs = 0u;
        return false;
    }

    if (_caretBlinkAnchorTickMs == 0u)
    {
        _caretBlinkAnchorTickMs = nowTickMs;
        _caretVisible           = true;
    }
    else
    {
        _caretVisible = (((nowTickMs - _caretBlinkAnchorTickMs) / kCaretBlinkPeriodMs) % 2u) == 0u;
    }

    return true;
}

void TextField::OnFocusChanged(WindowHost& host, bool focused)
{
    Control::OnFocusChanged(host, focused);
    if (focused)
    {
        RegenerateConcealedMaskEpoch();
        if (_multiline)
        {
            if (! _readOnly)
            {
                EnsureMultilineCaretVisible(&host);
            }
        }
        ResetCaretBlink(host);
    }
    else
    {
        RemaskPasswordReveal();
        BreakDirectEditMerge();
        _caretBlinkAnchorTickMs = 0u;
        _caretVisible           = true;
        _dragSelecting          = false;
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        _preferredMultilineXOffsetDip.reset();
        _multilineWheelDeltaRemainder  = 0.0f;
        _passwordRevealButtonHovered   = false;
        _passwordRevealButtonPressed   = false;
        _passwordRevealKeyboardFocused = false;
        if (_onBlur)
        {
            _onBlur();
        }
        Invalidate(host);
    }
}

void TextField::OnEnabledChanged(bool enabled) noexcept
{
    if (! enabled)
    {
        _passwordRevealButtonHovered   = false;
        _passwordRevealButtonPressed   = false;
        _passwordRevealKeyboardFocused = false;
        RemaskPasswordReveal();
    }
}

bool TextField::OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    if (rightButton)
    {
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        return OnContextMenu(host, false, point);
    }

    if (IsPasswordRevealButtonVisible() && PointInRect(GetPasswordRevealButtonRect(), point))
    {
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        _dragSelecting                 = false;
        _passwordRevealKeyboardFocused = false;
        _passwordRevealButtonHovered   = true;
        _passwordRevealButtonPressed   = true;
        SetPasswordRevealState(PasswordRevealState::Visible);
        ResetCaretBlink(host);
        host.SyncTextInput(this);
        Invalidate(host);
        return true;
    }

    if (IsClearButtonVisible() && PointInRect(GetClearButtonRect(), point))
    {
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        _passwordRevealKeyboardFocused = false;
        SetTextAndNotify({});
        ResetCaretBlink(host);
        Invalidate(host);
        return true;
    }

    host.SetFocusControl(this);
    _passwordRevealKeyboardFocused = false;
    if (_masked && _maskLengthPolicy == PasswordMaskLengthPolicy::Concealed && _passwordRevealState != PasswordRevealState::Visible && ! _multiline)
    {
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        SetCaretIndex(_text.size(), false);
        _dragSelecting             = false;
        const D2D1_RECT_F textRect = GetTextRect();
        EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
        ResetCaretBlink(host);
        host.SyncTextInput(this);
        Invalidate(host);
        return true;
    }

    if (! _multiline && ! ModifiersContainShift(modifiers) && ShouldPromoteSingleLineClickToSelectAll(host, _selectionClickSequence, point))
    {
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        SelectAllText();
        _dragSelecting             = false;
        const D2D1_RECT_F textRect = GetTextRect();
        EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
        ResetCaretBlink(host);
        host.SyncTextInput(this);
        Invalidate(host);
        return true;
    }

    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    if (_multiline)
    {
        const D2D1_RECT_F textRect     = GetTextRect();
        const std::wstring displayText = GetDisplayText();
        const auto multilineLayout     = CreateMultilineTextLayout(
            &host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
        const size_t displayHitIndex =
            HitTestMultilineCaretIndexDip(&host,
                                          displayText,
                                          FontRole::Body,
                                          textRect,
                                          MeasureWrappedLineOffsetDip(GetMultilineLineMetrics(multilineLayout.get()), _multilineFirstVisibleLine),
                                          point);
        const size_t hitIndex = DisplayTextIndexToControlTextIndex(displayHitIndex);
        _preferredMultilineXOffsetDip.reset();
        if (ModifiersContainShift(modifiers))
        {
            SetCaretIndex(hitIndex, true);
        }
        else
        {
            BreakDirectEditMerge();
            _selectionAnchorIndex = hitIndex;
            _caretIndex           = hitIndex;
        }
        _dragSelecting = true;
        EnsureMultilineCaretVisible(&host);
    }
    else
    {
        const D2D1_RECT_F textRect     = GetTextRect();
        const std::wstring displayText = GetDisplayText();
        const size_t displayHitIndex =
            HitTestCaretIndexDip(&host, displayText, FontRole::Body, textRect, _horizontalScrollDip, point, ResolveReadingDirection(GetFlowDirection()));
        const size_t hitIndex = DisplayTextIndexToControlTextIndex(displayHitIndex);
        _preferredMultilineXOffsetDip.reset();
        if (ModifiersContainShift(modifiers))
        {
            SetCaretIndex(hitIndex, true);
        }
        else
        {
            BreakDirectEditMerge();
            _selectionAnchorIndex = hitIndex;
            _caretIndex           = hitIndex;
        }
        _dragSelecting = true;
        EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
    }
    ResetCaretBlink(host);
    host.SyncTextInput(this);
    Invalidate(host);
    return true;
}

bool TextField::OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    if (rightButton)
    {
        return OnMouseDown(host, point, rightButton, modifiers);
    }

    host.SetFocusControl(this);
    const D2D1_RECT_F textRect     = GetTextRect();
    const std::wstring displayText = GetDisplayText();
    size_t hitIndex                = 0u;
    if (_multiline)
    {
        const auto multilineLayout = CreateMultilineTextLayout(
            &host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
        const size_t displayHitIndex =
            HitTestMultilineCaretIndexDip(&host,
                                          displayText,
                                          FontRole::Body,
                                          textRect,
                                          MeasureWrappedLineOffsetDip(GetMultilineLineMetrics(multilineLayout.get()), _multilineFirstVisibleLine),
                                          point);
        hitIndex = DisplayTextIndexToControlTextIndex(displayHitIndex);
    }
    else
    {
        const size_t displayHitIndex =
            HitTestCaretIndexDip(&host, displayText, FontRole::Body, textRect, _horizontalScrollDip, point, ResolveReadingDirection(GetFlowDirection()));
        hitIndex = DisplayTextIndexToControlTextIndex(displayHitIndex);
    }
    _preferredMultilineXOffsetDip.reset();
    SelectWordAt(hitIndex);
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    if (! _multiline)
    {
        ArmSingleLineSelectionClickSequence(_selectionClickSequence, point);
    }
    _dragSelecting = false;
    if (_multiline)
    {
        EnsureMultilineCaretVisible(&host);
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
    }
    ResetCaretBlink(host);
    host.SyncTextInput(this);
    Invalidate(host);
    return true;
}

bool TextField::OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT /*modifiers*/)
{
    const bool wasClearHovered   = _clearButtonHovered;
    const bool wasRevealHovered  = _passwordRevealButtonHovered;
    _clearButtonHovered          = IsClearButtonVisible() && PointInRect(GetClearButtonRect(), point);
    _passwordRevealButtonHovered = IsPasswordRevealButtonVisible() && PointInRect(GetPasswordRevealButtonRect(), point);
    if (_clearButtonHovered != wasClearHovered || _passwordRevealButtonHovered != wasRevealHovered)
    {
        Invalidate(host);
    }

    if (! _dragSelecting)
    {
        return _clearButtonHovered || _passwordRevealButtonHovered || _passwordRevealButtonPressed;
    }

    const D2D1_RECT_F textRect     = GetTextRect();
    const std::wstring displayText = GetDisplayText();
    size_t hitIndex                = 0u;
    if (_multiline)
    {
        const auto multilineLayout = CreateMultilineTextLayout(
            &host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
        const size_t displayHitIndex =
            HitTestMultilineCaretIndexDip(&host,
                                          displayText,
                                          FontRole::Body,
                                          textRect,
                                          MeasureWrappedLineOffsetDip(GetMultilineLineMetrics(multilineLayout.get()), _multilineFirstVisibleLine),
                                          point);
        hitIndex = DisplayTextIndexToControlTextIndex(displayHitIndex);
    }
    else
    {
        const size_t displayHitIndex =
            HitTestCaretIndexDip(&host, displayText, FontRole::Body, textRect, _horizontalScrollDip, point, ResolveReadingDirection(GetFlowDirection()));
        hitIndex = DisplayTextIndexToControlTextIndex(displayHitIndex);
    }
    _preferredMultilineXOffsetDip.reset();
    SetCaretIndex(hitIndex, true);
    if (_multiline)
    {
        EnsureMultilineCaretVisible(&host);
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
    }
    ResetCaretBlink(host);
    host.SyncTextInput(this);
    Invalidate(host);
    return true;
}

bool TextField::OnMouseUp(WindowHost& host, D2D1_POINT_2F /*point*/, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton)
    {
        return false;
    }

    if (_passwordRevealButtonPressed)
    {
        _passwordRevealButtonPressed = false;
        RemaskPasswordReveal();
        ResetCaretBlink(host);
        host.SyncTextInput(this);
        Invalidate(host);
        return true;
    }

    const bool wasDragging = _dragSelecting;
    _dragSelecting         = false;
    if (wasDragging)
    {
        Invalidate(host);
    }
    return wasDragging;
}

void TextField::OnCaptureLost(WindowHost& host)
{
    const bool hadRevealPress    = _passwordRevealButtonPressed;
    const bool wasDragging       = _dragSelecting;
    _dragSelecting               = false;
    _passwordRevealButtonPressed = false;
    _passwordRevealButtonHovered = false;

    if (hadRevealPress)
    {
        RemaskPasswordReveal();
        ResetCaretBlink(host);
        host.SyncTextInput(this);
    }
    if (hadRevealPress || wasDragging)
    {
        Invalidate(host);
    }
}

bool TextField::OnMouseWheel(WindowHost& host, D2D1_POINT_2F /*point*/, float wheelDelta, UINT /*modifiers*/)
{
    if (! _multiline)
    {
        return false;
    }

    const D2D1_RECT_F textRect     = GetTextRect();
    const float viewportHeightDip  = std::max(1.0f, textRect.bottom - textRect.top);
    const std::wstring displayText = GetDisplayText();
    const auto multilineLayout =
        CreateMultilineTextLayout(&host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
    const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
    const float fallbackLineHeightDip                  = EstimateMultilineFallbackLineHeightDip(&host, FontRole::Body);
    const MultilineViewportMetrics viewportMetrics     = BuildMultilineViewportMetrics(displayText, viewportHeightDip, lineMetrics, fallbackLineHeightDip);
    const size_t wheelLineCount                        = ComputeMultilineWheelLineCount(viewportMetrics.visibleLineCount);
    if (wheelLineCount == 0u)
    {
        return true;
    }

    if (viewportMetrics.totalLineCount <= viewportMetrics.visibleLineCount)
    {
        _multilineFirstVisibleLine    = 0u;
        _multilineWheelDeltaRemainder = 0.0f;
        if (host.GetFocusControl() == this)
        {
            host.SyncTextInput(this);
        }
        return true;
    }

    _multilineWheelDeltaRemainder += wheelDelta;
    const int wheelStepCount = static_cast<int>(_multilineWheelDeltaRemainder / static_cast<float>(WHEEL_DELTA));
    if (wheelStepCount == 0)
    {
        return true;
    }
    _multilineWheelDeltaRemainder -= static_cast<float>(wheelStepCount * WHEEL_DELTA);

    const int direction              = wheelStepCount > 0 ? -1 : 1;
    const size_t deltaLines          = static_cast<size_t>(std::abs(wheelStepCount)) * (std::max)(static_cast<size_t>(1u), wheelLineCount);
    const size_t maxFirstVisibleLine = ComputeMultilineMaxFirstVisibleLine(viewportMetrics);
    if (_multilineFirstVisibleLine > maxFirstVisibleLine)
    {
        _multilineFirstVisibleLine = maxFirstVisibleLine;
    }
    const int64_t nextFirstVisibleLine =
        std::clamp<int64_t>(static_cast<int64_t>(_multilineFirstVisibleLine) + (static_cast<int64_t>(direction) * static_cast<int64_t>(deltaLines)),
                            0ll,
                            static_cast<int64_t>(maxFirstVisibleLine));

    if (_multilineFirstVisibleLine == static_cast<size_t>(nextFirstVisibleLine))
    {
        return true;
    }

    _multilineFirstVisibleLine = static_cast<size_t>(nextFirstVisibleLine);
    if (host.GetFocusControl() == this)
    {
        host.SyncTextInput(this);
    }
    Invalidate(host);
    return true;
}

bool TextField::OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers)
{
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    const auto textRect     = GetTextRect();
    const auto refreshCaret = [this, &host, &textRect]() noexcept
    {
        ResetCaretBlink(host);
        if (_multiline)
        {
            EnsureMultilineCaretVisible(&host);
        }
        else
        {
            EnsureCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
        }
        Invalidate(host);
    };

    if (virtualKey == VK_ESCAPE && _masked && _passwordRevealState != PasswordRevealState::Hidden)
    {
        _passwordRevealButtonPressed   = false;
        _passwordRevealKeyboardFocused = false;
        RemaskPasswordReveal();
        host.SyncTextInput(this);
        Invalidate(host);
        return false;
    }

    if (virtualKey == VK_TAB && ! ModifiersContainAlt(modifiers) && ! ModifiersContainCtrl(modifiers))
    {
        if (_passwordRevealKeyboardFocused)
        {
            _passwordRevealKeyboardFocused = false;
            _passwordRevealButtonPressed   = false;
            RemaskPasswordReveal();
            host.SyncTextInput(this);
            Invalidate(host);
            return ModifiersContainShift(modifiers);
        }

        if (! ModifiersContainShift(modifiers) && IsPasswordRevealButtonVisible())
        {
            _passwordRevealKeyboardFocused = true;
            _passwordRevealButtonPressed   = false;
            _passwordRevealButtonHovered   = false;
            Invalidate(host);
            return true;
        }

        return false;
    }

    if (_passwordRevealKeyboardFocused && (virtualKey == VK_SPACE || virtualKey == VK_RETURN) && ! ModifiersContainAlt(modifiers) &&
        ! ModifiersContainCtrl(modifiers))
    {
        _passwordRevealButtonPressed = true;
        SetPasswordRevealState(PasswordRevealState::Visible);
        ResetCaretBlink(host);
        host.SyncTextInput(this);
        Invalidate(host);
        return true;
    }

    if (_onPreviewKeyDown && _onPreviewKeyDown(host, virtualKey, modifiers))
    {
        return true;
    }

    if (ModifiersContainCtrl(modifiers))
    {
        if (virtualKey == 'Z' && ! ModifiersContainShift(modifiers) && ! _readOnly)
        {
            if (TryUndoDirectEdit())
            {
                if (! NotifyChanged())
                {
                    return true;
                }
                refreshCaret();
                host.SyncTextInput(this);
                return true;
            }
            return false;
        }
        if (virtualKey == 'Y' && ! ModifiersContainShift(modifiers) && ! _readOnly)
        {
            if (TryRedoDirectEdit())
            {
                if (! NotifyChanged())
                {
                    return true;
                }
                refreshCaret();
                host.SyncTextInput(this);
                return true;
            }
            return false;
        }
        if (virtualKey == 'A')
        {
            SelectAllText();
            refreshCaret();
            return true;
        }
        if (virtualKey == 'C')
        {
            return OnCopy(host);
        }
        if (virtualKey == 'X' && ! _readOnly)
        {
            if (OnCopy(host))
            {
                RecordUndoStateForDirectEdit();
                _preferredMultilineXOffsetDip.reset();
                if (DeleteSelection())
                {
                    if (! NotifyChanged())
                    {
                        return true;
                    }
                    refreshCaret();
                }
                return true;
            }
            return false;
        }
        if (virtualKey == VK_INSERT)
        {
            return OnCopy(host);
        }
        if (virtualKey == 'V' && ! _readOnly)
        {
            const auto clipboardText = host.ReadTextFromClipboard();
            if (clipboardText)
            {
                const std::wstring normalizedClipboardText = NormalizePastedControlText(clipboardText.value(), _multiline);
                const bool willMutate                      = HasSelection() || ! normalizedClipboardText.empty();
                if (willMutate)
                {
                    RecordUndoStateForDirectEdit();
                    static_cast<void>(DeleteSelection());
                    _text.insert(_caretIndex, normalizedClipboardText);
                    _caretIndex += normalizedClipboardText.size();
                    _selectionAnchorIndex.reset();
                    if (! NotifyChanged())
                    {
                        return true;
                    }
                    refreshCaret();
                }
                return true;
            }
        }
        if (virtualKey == VK_BACK && ! _readOnly)
        {
            const size_t eraseFrom = FindPreviousWordBoundary(_text, _caretIndex);
            if (HasSelection() || eraseFrom < _caretIndex)
            {
                RecordUndoStateForDirectEdit();
                _preferredMultilineXOffsetDip.reset();
            }
            if (DeleteSelection())
            {
                if (! NotifyChanged())
                {
                    return true;
                }
                refreshCaret();
                return true;
            }
            if (eraseFrom < _caretIndex)
            {
                _text.erase(eraseFrom, _caretIndex - eraseFrom);
                _caretIndex = eraseFrom;
                _selectionAnchorIndex.reset();
                if (! NotifyChanged())
                {
                    return true;
                }
                refreshCaret();
                return true;
            }
            return true;
        }
        if (virtualKey == VK_DELETE && ! _readOnly)
        {
            const size_t eraseTo = FindNextWordBoundary(_text, _caretIndex);
            if (HasSelection() || eraseTo > _caretIndex)
            {
                RecordUndoStateForDirectEdit();
                _preferredMultilineXOffsetDip.reset();
            }
            if (DeleteSelection())
            {
                if (! NotifyChanged())
                {
                    return true;
                }
                refreshCaret();
                return true;
            }
            if (eraseTo > _caretIndex)
            {
                _text.erase(_caretIndex, eraseTo - _caretIndex);
                _selectionAnchorIndex.reset();
                if (! NotifyChanged())
                {
                    return true;
                }
                refreshCaret();
                return true;
            }
            return true;
        }
    }

    if (virtualKey == VK_INSERT && ModifiersContainShift(modifiers) && ! _readOnly)
    {
        const auto clipboardText = host.ReadTextFromClipboard();
        if (clipboardText)
        {
            const std::wstring normalizedClipboardText = NormalizePastedControlText(clipboardText.value(), _multiline);
            const bool willMutate                      = HasSelection() || ! normalizedClipboardText.empty();
            if (willMutate)
            {
                RecordUndoStateForDirectEdit();
                static_cast<void>(DeleteSelection());
                _text.insert(_caretIndex, normalizedClipboardText);
                _caretIndex += normalizedClipboardText.size();
                _selectionAnchorIndex.reset();
                if (! NotifyChanged())
                {
                    return true;
                }
                refreshCaret();
            }
            return true;
        }
    }
    if (virtualKey == VK_DELETE && ModifiersContainShift(modifiers) && ! _readOnly)
    {
        if (OnCopy(host))
        {
            RecordUndoStateForDirectEdit();
            _preferredMultilineXOffsetDip.reset();
            if (DeleteSelection())
            {
                if (! NotifyChanged())
                {
                    return true;
                }
                refreshCaret();
            }
            return true;
        }
        return false;
    }

    if (virtualKey == VK_LEFT)
    {
        _preferredMultilineXOffsetDip.reset();
        const bool extendSelection = ModifiersContainShift(modifiers);
        if (ModifiersContainCtrl(modifiers))
        {
            SetCaretIndex(FindPreviousWordBoundary(_text, _caretIndex), extendSelection);
            refreshCaret();
            return true;
        }
        if (! extendSelection && HasSelection())
        {
            BreakDirectEditMerge();
            _caretIndex = GetSelectionRange().value().first;
            _selectionAnchorIndex.reset();
            refreshCaret();
            return true;
        }
        if (_caretIndex > 0u)
        {
            SetCaretIndex(StepToPreviousTextElement(_text, _caretIndex), extendSelection);
            refreshCaret();
        }
        return true;
    }
    if (virtualKey == VK_RIGHT)
    {
        _preferredMultilineXOffsetDip.reset();
        const bool extendSelection = ModifiersContainShift(modifiers);
        if (ModifiersContainCtrl(modifiers))
        {
            SetCaretIndex(FindNextWordBoundary(_text, _caretIndex), extendSelection);
            refreshCaret();
            return true;
        }
        if (! extendSelection && HasSelection())
        {
            BreakDirectEditMerge();
            _caretIndex = GetSelectionRange().value().second;
            _selectionAnchorIndex.reset();
            refreshCaret();
            return true;
        }
        SetCaretIndex(StepToNextTextElement(_text, _caretIndex), extendSelection);
        refreshCaret();
        return true;
    }
    if (! _multiline && virtualKey == VK_UP && ! ModifiersContainCtrl(modifiers) && ! ModifiersContainAlt(modifiers))
    {
        _preferredMultilineXOffsetDip.reset();
        SetCaretIndex(0u, ModifiersContainShift(modifiers));
        refreshCaret();
        return true;
    }
    if (! _multiline && virtualKey == VK_DOWN && ! ModifiersContainCtrl(modifiers) && ! ModifiersContainAlt(modifiers))
    {
        _preferredMultilineXOffsetDip.reset();
        SetCaretIndex(_text.size(), ModifiersContainShift(modifiers));
        refreshCaret();
        return true;
    }
    if (_multiline && virtualKey == VK_UP)
    {
        const bool extendSelection = ModifiersContainShift(modifiers);
        const size_t nextCaretIndex =
            MoveMultilineCaretVertically(&host, _text, FontRole::Body, GetTextRect(), _caretIndex, false, _preferredMultilineXOffsetDip);
        BreakDirectEditMerge();
        SetSingleLineCaretIndex(_caretIndex, _selectionAnchorIndex, nextCaretIndex, extendSelection);
        refreshCaret();
        return true;
    }
    if (_multiline && virtualKey == VK_DOWN)
    {
        const bool extendSelection = ModifiersContainShift(modifiers);
        const size_t nextCaretIndex =
            MoveMultilineCaretVertically(&host, _text, FontRole::Body, GetTextRect(), _caretIndex, true, _preferredMultilineXOffsetDip);
        BreakDirectEditMerge();
        SetSingleLineCaretIndex(_caretIndex, _selectionAnchorIndex, nextCaretIndex, extendSelection);
        refreshCaret();
        return true;
    }
    if (_multiline && virtualKey == VK_PRIOR)
    {
        const bool extendSelection                     = ModifiersContainShift(modifiers);
        const D2D1_RECT_F pageTextRect                 = GetTextRect();
        const MultilineViewportMetrics viewportMetrics = ComputeMultilineViewportMetrics(&host, _text, FontRole::Body, pageTextRect);
        const size_t nextCaretIndex                    = MoveMultilineCaretByPage(
            &host, _text, FontRole::Body, pageTextRect, _caretIndex, false, _preferredMultilineXOffsetDip, viewportMetrics.visibleLineCount);
        BreakDirectEditMerge();
        SetSingleLineCaretIndex(_caretIndex, _selectionAnchorIndex, nextCaretIndex, extendSelection);
        refreshCaret();
        return true;
    }
    if (_multiline && virtualKey == VK_NEXT)
    {
        const bool extendSelection                     = ModifiersContainShift(modifiers);
        const D2D1_RECT_F pageTextRect                 = GetTextRect();
        const MultilineViewportMetrics viewportMetrics = ComputeMultilineViewportMetrics(&host, _text, FontRole::Body, pageTextRect);
        const size_t nextCaretIndex                    = MoveMultilineCaretByPage(
            &host, _text, FontRole::Body, pageTextRect, _caretIndex, true, _preferredMultilineXOffsetDip, viewportMetrics.visibleLineCount);
        BreakDirectEditMerge();
        SetSingleLineCaretIndex(_caretIndex, _selectionAnchorIndex, nextCaretIndex, extendSelection);
        refreshCaret();
        return true;
    }
    if (virtualKey == VK_HOME)
    {
        _preferredMultilineXOffsetDip.reset();
        if (_multiline && ! ModifiersContainCtrl(modifiers))
        {
            const size_t nextCaretIndex =
                TryMoveCaretToWrappedLineBoundary(&host, _text, FontRole::Body, GetTextRect(), _caretIndex, false).value_or(FindLineStart(_text, _caretIndex));
            SetCaretIndex(nextCaretIndex, ModifiersContainShift(modifiers));
        }
        else
        {
            SetCaretIndex(0u, ModifiersContainShift(modifiers));
        }
        refreshCaret();
        return true;
    }
    if (virtualKey == VK_END)
    {
        _preferredMultilineXOffsetDip.reset();
        if (_multiline && ! ModifiersContainCtrl(modifiers))
        {
            const size_t nextCaretIndex =
                TryMoveCaretToWrappedLineBoundary(&host, _text, FontRole::Body, GetTextRect(), _caretIndex, true).value_or(FindLineEnd(_text, _caretIndex));
            SetCaretIndex(nextCaretIndex, ModifiersContainShift(modifiers));
        }
        else
        {
            SetCaretIndex(_text.size(), ModifiersContainShift(modifiers));
        }
        refreshCaret();
        return true;
    }
    if (virtualKey == VK_RETURN && _multiline)
    {
        return true;
    }
    if (virtualKey == VK_RETURN && ! _multiline)
    {
        if (_onSubmitted)
        {
            _onSubmitted();
            return true;
        }
        return false;
    }
    if (virtualKey == VK_BACK && ! _readOnly)
    {
        if (HasSelection() || (_caretIndex > 0 && ! _text.empty()))
        {
            RecordUndoStateForDirectEdit(HasSelection() ? DirectEditMergeKind::None : DirectEditMergeKind::BackspaceRun);
            _preferredMultilineXOffsetDip.reset();
        }
        if (DeleteSelection())
        {
            if (! NotifyChanged())
            {
                return true;
            }
            refreshCaret();
            return true;
        }
        if (_caretIndex > 0 && ! _text.empty())
        {
            const size_t eraseFrom = StepToPreviousTextElement(_text, _caretIndex);
            _text.erase(eraseFrom, _caretIndex - eraseFrom);
            _caretIndex = eraseFrom;
            _selectionAnchorIndex.reset();
            if (! NotifyChanged())
            {
                return true;
            }
            refreshCaret();
        }
        return true;
    }
    if (virtualKey == VK_DELETE && ! _readOnly)
    {
        if (HasSelection() || _caretIndex < _text.size())
        {
            RecordUndoStateForDirectEdit(HasSelection() ? DirectEditMergeKind::None : DirectEditMergeKind::DeleteRun);
            _preferredMultilineXOffsetDip.reset();
        }
        if (DeleteSelection())
        {
            if (! NotifyChanged())
            {
                return true;
            }
            refreshCaret();
            return true;
        }
        if (_caretIndex < _text.size())
        {
            const size_t eraseTo = StepToNextTextElement(_text, _caretIndex);
            _text.erase(_caretIndex, eraseTo - _caretIndex);
            _selectionAnchorIndex.reset();
            if (! NotifyChanged())
            {
                return true;
            }
            refreshCaret();
        }
        return true;
    }
    return false;
}

bool TextField::OnKeyUp(WindowHost& host, UINT virtualKey, UINT modifiers)
{
    if (! _passwordRevealKeyboardFocused || (virtualKey != VK_SPACE && virtualKey != VK_RETURN) || ModifiersContainAlt(modifiers) ||
        ModifiersContainCtrl(modifiers))
    {
        return false;
    }

    if (_passwordRevealButtonPressed)
    {
        _passwordRevealButtonPressed = false;
        RemaskPasswordReveal();
        ResetCaretBlink(host);
        host.SyncTextInput(this);
        Invalidate(host);
    }
    return true;
}

bool TextField::OnChar(WindowHost& host, wchar_t ch, UINT /*modifiers*/)
{
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    if (_passwordRevealKeyboardFocused && (ch == L' ' || ch == L'\r'))
    {
        return true;
    }
    if (_readOnly)
    {
        return false;
    }

    if (ch == L'\r')
    {
        if (! _multiline)
        {
            return true;
        }
        RecordUndoStateForDirectEdit(DirectEditMergeKind::InsertRun);
        _preferredMultilineXOffsetDip.reset();
        static_cast<void>(DeleteSelection());
        _text.insert(_caretIndex, 1u, L'\n');
        _caretIndex += 1u;
        _selectionAnchorIndex.reset();
    }
    else if (std::iswcntrl(static_cast<wint_t>(ch)) == 0 && (_multiline || ch != L'\t'))
    {
        const bool completesPendingSurrogatePair =
            IsUtf16TrailSurrogate(ch) && ! GetSelectionRange().has_value() && _caretIndex > 0u && IsUtf16LeadSurrogate(_text[_caretIndex - 1u]);
        if (! completesPendingSurrogatePair)
        {
            RecordUndoStateForDirectEdit(DirectEditMergeKind::InsertRun);
        }
        _preferredMultilineXOffsetDip.reset();
        static_cast<void>(DeleteSelection());
        _text.insert(_caretIndex, 1u, ch);
        _caretIndex += 1u;
        _selectionAnchorIndex.reset();
    }
    else
    {
        return false;
    }

    if (! NotifyChanged())
    {
        return true;
    }
    ResetCaretBlink(host);
    if (_multiline)
    {
        EnsureMultilineCaretVisible(&host);
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, GetTextRect().right - GetTextRect().left));
    }
    Invalidate(host);
    return true;
}

bool TextField::OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip)
{
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    host.SetFocusControl(this);
    host.SyncTextInput(this);
    ResetCaretBlink(host);
    Invalidate(host);
    return Control::OnContextMenu(host, keyboardInvocation, pointDip);
}

bool TextField::OnCopy(WindowHost& host)
{
    if (_masked && _passwordRevealMode != PasswordRevealMode::Visible && _passwordRevealState != PasswordRevealState::Visible)
    {
        return false;
    }

    if (const std::optional<std::pair<size_t, size_t>> selectionRange = GetSelectionRange())
    {
        const auto [selectionStart, selectionEnd] = selectionRange.value();
        return host.CopyTextToClipboard(_text.substr(selectionStart, selectionEnd - selectionStart));
    }

    return false;
}

bool TextField::OnSelectAll(WindowHost& host)
{
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    _preferredMultilineXOffsetDip.reset();
    SelectAllText();
    ResetCaretBlink(host);
    if (_multiline)
    {
        EnsureMultilineCaretVisible(&host);
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, GetTextRect().right - GetTextRect().left));
    }
    host.SyncTextInput(this);
    Invalidate(host);
    return true;
}

bool TextField::SupportsTextInput() const noexcept
{
    return true;
}

std::optional<D2D1_RECT_F> TextField::GetTextInputViewportRect() const noexcept
{
    const D2D1_RECT_F textRect = GetTextRect();
    if (textRect.right <= textRect.left || textRect.bottom <= textRect.top)
    {
        return std::nullopt;
    }
    return textRect;
}

std::optional<D2D1_RECT_F> TextField::GetTextInputCaretRect(const WindowHost& host, size_t controlTextIndex) const noexcept
{
    const D2D1_RECT_F textRect     = GetTextRect();
    const std::wstring displayText = GetDisplayText();
    const size_t displayCaretIndex = ControlTextIndexToDisplayTextIndex(controlTextIndex);
    if (textRect.right <= textRect.left || textRect.bottom <= textRect.top)
    {
        return std::nullopt;
    }

    if (_multiline)
    {
        const auto multilineLayout =
            GetOrCreateMultilineLayout(&host, displayText, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
        const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
        const float multilineScrollDip                     = MeasureWrappedLineOffsetDip(lineMetrics, _multilineFirstVisibleLine);
        return MeasureMultilineCaretRectDip(&host, displayText, FontRole::Body, textRect, multilineScrollDip, displayCaretIndex);
    }

    const bool perfEnabled                          = Debug::Perf::IsCaptureEnabled();
    const auto startedAt                            = perfEnabled ? std::chrono::steady_clock::now() : std::chrono::steady_clock::time_point{};
    const DWRITE_READING_DIRECTION readingDirection = ResolveReadingDirection(GetFlowDirection());
    const wil::com_ptr<IDWriteTextLayout> layout    = GetOrCreateSingleLineLayout(&host,
                                                                                  displayText,
                                                                                  std::max(1.0f, textRect.right - textRect.left + _horizontalScrollDip),
                                                                                  std::max(1.0f, textRect.bottom - textRect.top),
                                                                                  readingDirection);
    const float caretOffset                         = MeasureCaretOffsetDip(layout.get(), displayText, displayCaretIndex);
    const float caretMaxX                           = std::max(textRect.left, textRect.right - 1.0f);
    const float caretX                              = std::clamp(textRect.left + caretOffset - _horizontalScrollDip, textRect.left, caretMaxX);
    const D2D1_RECT_F result                        = D2D1::RectF(caretX, textRect.top + 2.0f, caretX + 1.0f, textRect.bottom - 2.0f);
    if (perfEnabled && ShouldEmitSingleLineBiDiTextMetric(displayText, readingDirection))
    {
        const std::wstring_view detail = readingDirection == DWRITE_READING_DIRECTION_RIGHT_TO_LEFT ? L"text-field-rtl" : L"text-field-ltr";
        Debug::Perf::Emit(L"dxui.textinput.bidi_caret_rect_us", detail, Debug::Perf::ElapsedUs(startedAt), displayCaretIndex, displayText.size(), S_OK);
    }
    return result;
}

std::optional<std::vector<D2D1_RECT_F>> TextField::GetTextInputRangeRects(const WindowHost& host,
                                                                          size_t controlTextStartIndex,
                                                                          size_t controlTextEndIndex) const
{
    const std::wstring displayText                              = GetDisplayText();
    const std::optional<std::pair<size_t, size_t>> displayRange = ControlTextRangeToDisplayTextRange(controlTextStartIndex, controlTextEndIndex);
    if (! displayRange.has_value())
    {
        return std::nullopt;
    }
    const size_t rangeStart = displayRange.value().first;
    const size_t rangeEnd   = displayRange.value().second;
    if (rangeStart >= rangeEnd || rangeStart > static_cast<size_t>(std::numeric_limits<UINT32>::max()) ||
        rangeEnd - rangeStart > static_cast<size_t>(std::numeric_limits<UINT32>::max()))
    {
        return std::nullopt;
    }

    const D2D1_RECT_F textRect = GetTextRect();
    if (textRect.right <= textRect.left || textRect.bottom <= textRect.top)
    {
        return std::nullopt;
    }

    if (! _multiline)
    {
        const DWRITE_READING_DIRECTION readingDirection = ResolveReadingDirection(GetFlowDirection());
        const float heightDip                           = std::max(1.0f, textRect.bottom - textRect.top);
        const wil::com_ptr<IDWriteTextLayout> layout =
            GetOrCreateSingleLineLayout(&host, displayText, std::max(1.0f, textRect.right - textRect.left + _horizontalScrollDip), heightDip, readingDirection);
        if (! layout)
        {
            return std::nullopt;
        }

        UINT32 actualCount = 0u;
        std::vector<DWRITE_HIT_TEST_METRICS> metrics((rangeEnd - rangeStart) + 4u);
        HRESULT hr = layout->HitTestTextRange(static_cast<UINT32>(rangeStart),
                                              static_cast<UINT32>(rangeEnd - rangeStart),
                                              textRect.left - _horizontalScrollDip,
                                              textRect.top,
                                              metrics.data(),
                                              static_cast<UINT32>(metrics.size()),
                                              &actualCount);
        if (hr == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER) && actualCount > 0u)
        {
            metrics.resize(actualCount);
            hr = layout->HitTestTextRange(static_cast<UINT32>(rangeStart),
                                          static_cast<UINT32>(rangeEnd - rangeStart),
                                          textRect.left - _horizontalScrollDip,
                                          textRect.top,
                                          metrics.data(),
                                          static_cast<UINT32>(metrics.size()),
                                          &actualCount);
        }
        if (FAILED(hr) || actualCount == 0u)
        {
            return std::nullopt;
        }

        std::vector<D2D1_RECT_F> rects;
        rects.reserve(actualCount);
        for (UINT32 index = 0u; index < actualCount; ++index)
        {
            const DWRITE_HIT_TEST_METRICS& hit = metrics[index];
            D2D1_RECT_F rect                   = D2D1::RectF(hit.left, hit.top, hit.left + hit.width, hit.top + hit.height);
            rect                               = ClipTextInputRectToBounds(rect, textRect);
            if (rect.right > rect.left && rect.bottom > rect.top)
            {
                rects.push_back(rect);
            }
        }

        if (rects.empty())
        {
            return std::nullopt;
        }
        return rects;
    }

    const auto multilineLayout =
        GetOrCreateMultilineLayout(&host, displayText, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
    if (! multilineLayout)
    {
        return std::nullopt;
    }

    const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
    const float multilineScrollDip                     = MeasureWrappedLineOffsetDip(lineMetrics, _multilineFirstVisibleLine);

    UINT32 actualCount = 0u;
    std::vector<DWRITE_HIT_TEST_METRICS> metrics((rangeEnd - rangeStart) + 4u);
    HRESULT hr = multilineLayout->HitTestTextRange(static_cast<UINT32>(rangeStart),
                                                   static_cast<UINT32>(rangeEnd - rangeStart),
                                                   textRect.left,
                                                   textRect.top - multilineScrollDip,
                                                   metrics.data(),
                                                   static_cast<UINT32>(metrics.size()),
                                                   &actualCount);
    if (hr == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER) && actualCount > 0u)
    {
        metrics.resize(actualCount);
        hr = multilineLayout->HitTestTextRange(static_cast<UINT32>(rangeStart),
                                               static_cast<UINT32>(rangeEnd - rangeStart),
                                               textRect.left,
                                               textRect.top - multilineScrollDip,
                                               metrics.data(),
                                               static_cast<UINT32>(metrics.size()),
                                               &actualCount);
    }
    if (FAILED(hr) || actualCount == 0u)
    {
        return std::nullopt;
    }

    std::vector<D2D1_RECT_F> rects;
    rects.reserve(actualCount);
    for (UINT32 index = 0u; index < actualCount; ++index)
    {
        const DWRITE_HIT_TEST_METRICS& hit = metrics[index];
        D2D1_RECT_F rect                   = D2D1::RectF(hit.left, hit.top, hit.left + hit.width, hit.top + hit.height);
        rect                               = ClipTextInputRectToBounds(rect, textRect);
        if (rect.right > rect.left && rect.bottom > rect.top)
        {
            rects.push_back(rect);
        }
    }

    if (rects.empty())
    {
        return std::nullopt;
    }
    return rects;
}

std::optional<size_t> TextField::HitTestTextInputPoint(const WindowHost& host, D2D1_POINT_2F point) const noexcept
{
    const D2D1_RECT_F textRect = GetTextRect();
    if (textRect.right <= textRect.left || textRect.bottom <= textRect.top)
    {
        return std::nullopt;
    }

    const std::wstring displayText = GetDisplayText();
    if (_multiline)
    {
        const auto multilineLayout =
            GetOrCreateMultilineLayout(&host, displayText, std::max(1.0f, textRect.right - textRect.left), std::max(1.0f, textRect.bottom - textRect.top));
        const size_t displayHitIndex =
            HitTestMultilineCaretIndexDip(&host,
                                          displayText,
                                          FontRole::Body,
                                          textRect,
                                          MeasureWrappedLineOffsetDip(GetMultilineLineMetrics(multilineLayout.get()), _multilineFirstVisibleLine),
                                          point);
        return DisplayTextIndexToControlTextIndex(displayHitIndex);
    }

    const size_t displayHitIndex =
        HitTestCaretIndexDip(&host, displayText, FontRole::Body, textRect, _horizontalScrollDip, point, ResolveReadingDirection(GetFlowDirection()));
    return DisplayTextIndexToControlTextIndex(displayHitIndex);
}

bool TextField::ExportTextInputState(TextInputState& outState) const
{
    outState.text                 = _text;
    outState.selectionAnchorIndex = _selectionAnchorIndex;
    outState.caretIndex           = _caretIndex;
    outState.firstVisibleLine     = _multilineFirstVisibleLine;
    outState.readOnly             = _readOnly;
    outState.masked               = _masked;
    outState.multiline            = _multiline;
    return true;
}

bool TextField::ImportTextInputState(WindowHost& host, const TextInputState& state, bool notifyChange)
{
    const std::wstring previousText       = _text;
    const size_t previousFirstVisibleLine = _multilineFirstVisibleLine;
    _undoHistory.clear();
    _redoHistory.clear();
    BreakDirectEditMerge();
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    _text       = state.text;
    _caretIndex = std::min(state.caretIndex, _text.size());
    InvalidateSingleLineLayoutCache();
    InvalidateMultilineLayoutCache();
    if (state.selectionAnchorIndex)
    {
        _selectionAnchorIndex = std::min(state.selectionAnchorIndex.value(), _text.size());
        if (_selectionAnchorIndex.value() == _caretIndex)
        {
            _selectionAnchorIndex.reset();
        }
    }
    else
    {
        _selectionAnchorIndex.reset();
    }

    const bool preserveReadOnlyMultilineViewport = _multiline && _readOnly && state.multiline;
    _multilineFirstVisibleLine = state.multiline ? (preserveReadOnlyMultilineViewport ? previousFirstVisibleLine : state.firstVisibleLine) : 0u;
    _preferredMultilineXOffsetDip.reset();
    _multilineWheelDeltaRemainder = 0.0f;
    _dragSelecting                = false;
    ResetCaretBlink(host);
    if (_multiline)
    {
        if (preserveReadOnlyMultilineViewport)
        {
            if (previousText != _text)
            {
                const D2D1_RECT_F textRect     = GetTextRect();
                const std::wstring displayText = GetDisplayText();
                if (displayText.empty())
                {
                    _multilineFirstVisibleLine = 0u;
                }
                else
                {
                    const float viewportHeightDip = std::max(1.0f, textRect.bottom - textRect.top);
                    const auto multilineLayout =
                        CreateMultilineTextLayout(&host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
                    const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
                    const float fallbackLineHeightDip                  = EstimateMultilineFallbackLineHeightDip(&host, FontRole::Body);
                    const MultilineViewportMetrics viewportMetrics =
                        BuildMultilineViewportMetrics(displayText, viewportHeightDip, lineMetrics, fallbackLineHeightDip);
                    _multilineFirstVisibleLine = ClampMultilineFirstVisibleLine(_multilineFirstVisibleLine, viewportMetrics);
                }
            }
        }
        else
        {
            EnsureMultilineCaretVisible(&host);
        }
    }
    else
    {
        EnsureCaretVisible(&host, std::max(1.0f, GetTextRect().right - GetTextRect().left));
    }
    Invalidate(host);
    if (notifyChange && previousText != _text)
    {
        static_cast<void>(NotifyChanged());
        return true;
    }
    return true;
}

TextField::EditHistoryState TextField::CaptureEditHistoryState() const
{
    return EditHistoryState{
        .text = _text, .caretIndex = _caretIndex, .selectionAnchorIndex = _selectionAnchorIndex, .firstVisibleLine = _multilineFirstVisibleLine};
}

void TextField::RestoreEditHistoryState(const EditHistoryState& state) noexcept
{
    _text       = state.text;
    _caretIndex = std::min(state.caretIndex, _text.size());
    if (state.selectionAnchorIndex)
    {
        _selectionAnchorIndex = std::min(state.selectionAnchorIndex.value(), _text.size());
        if (_selectionAnchorIndex.value() == _caretIndex)
        {
            _selectionAnchorIndex.reset();
        }
    }
    else
    {
        _selectionAnchorIndex.reset();
    }
    _multilineFirstVisibleLine = _multiline ? state.firstVisibleLine : 0u;
    _preferredMultilineXOffsetDip.reset();
    _multilineWheelDeltaRemainder = 0.0f;
    _dragSelecting                = false;
    _horizontalScrollDip          = 0.0f;
    InvalidateSingleLineLayoutCache();
    InvalidateMultilineLayoutCache();
}

void TextField::BeginEditTransactionMetric(const wchar_t* detail) noexcept
{
    if (! Debug::Perf::IsCaptureEnabled())
    {
        return;
    }

    _pendingEditTransactionStartedAt = std::chrono::steady_clock::now();
    _pendingEditTransactionDetail    = detail ? detail : L"";
}

void TextField::FinishEditTransactionMetric() noexcept
{
    if (! _pendingEditTransactionStartedAt.has_value())
    {
        return;
    }

    const auto startedAt        = _pendingEditTransactionStartedAt.value();
    const wchar_t* const detail = _pendingEditTransactionDetail ? _pendingEditTransactionDetail : L"";
    _pendingEditTransactionStartedAt.reset();
    _pendingEditTransactionDetail = L"";

    Debug::Perf::Emit(L"dxui.textinput.edit_transaction_us",
                      detail,
                      Debug::Perf::ElapsedUs(startedAt),
                      static_cast<uint64_t>(_text.size()),
                      static_cast<uint64_t>(_caretIndex),
                      S_OK);
    Debug::Perf::Emit(L"dxui.textinput.undo_depth", detail, 0u, static_cast<uint64_t>(_undoHistory.size()), static_cast<uint64_t>(_redoHistory.size()), S_OK);
}

void TextField::RegenerateConcealedMaskEpoch() const noexcept
{
    _concealedMaskVisibleDotCountValid = false;
}

size_t TextField::GetConcealedMaskVisibleDotCount(size_t exactCount) const noexcept
{
    const ConcealedMaskBucket bucket = GetConcealedMaskBucket(exactCount);
    if (bucket.displayBase == 0u)
    {
        _concealedMaskVisibleDotCountValid = false;
        _concealedMaskBucketStart          = 0u;
        _concealedMaskBucketEnd            = 0u;
        _concealedMaskVisibleDotCount      = 0u;
        return 0u;
    }

    if (_concealedMaskVisibleDotCountValid && exactCount >= _concealedMaskBucketStart && exactCount <= _concealedMaskBucketEnd)
    {
        return _concealedMaskVisibleDotCount;
    }

    const size_t previousDotCount = _concealedMaskVisibleDotCount;
    const size_t displaySpan      = std::max<size_t>(1u, bucket.displaySpan);
    const uint64_t mixedEpoch     = (_concealedMaskEpoch * 1103515245ull) + 12345ull + (static_cast<uint64_t>(bucket.displayBase) * 2654435761ull);
    size_t displayOffset          = static_cast<size_t>(mixedEpoch % displaySpan);
    size_t dotCount               = bucket.displayBase + displayOffset;
    if (displaySpan > 1u && dotCount == previousDotCount)
    {
        displayOffset = (displayOffset + 1u) % displaySpan;
        dotCount      = bucket.displayBase + displayOffset;
    }

    ++_concealedMaskEpoch;
    _concealedMaskBucketStart          = bucket.start;
    _concealedMaskBucketEnd            = bucket.end;
    _concealedMaskVisibleDotCount      = dotCount;
    _concealedMaskVisibleDotCountValid = true;
    return _concealedMaskVisibleDotCount;
}

void TextField::BreakDirectEditMerge() noexcept
{
    _directEditMergeKind = DirectEditMergeKind::None;
}

void TextField::RecordUndoStateForDirectEdit(DirectEditMergeKind mergeKind)
{
    BeginEditTransactionMetric(L"direct-edit");
    const bool mergeable = mergeKind != DirectEditMergeKind::None;
    const bool coalesced = mergeable && _directEditMergeKind == mergeKind && ! _undoHistory.empty() && ! HasSelection();
    if (! coalesced)
    {
        _undoHistory.push_back(CaptureEditHistoryState());
        if (_undoHistory.size() > kMaxEditHistoryEntries)
        {
            _undoHistory.erase(_undoHistory.begin());
        }
    }
    _directEditMergeKind = mergeable ? mergeKind : DirectEditMergeKind::None;
    _redoHistory.clear();
}

bool TextField::TryUndoDirectEdit() noexcept
{
    if (_undoHistory.empty())
    {
        return false;
    }

    BeginEditTransactionMetric(L"undo");
    BreakDirectEditMerge();
    _redoHistory.push_back(CaptureEditHistoryState());
    if (_redoHistory.size() > kMaxEditHistoryEntries)
    {
        _redoHistory.erase(_redoHistory.begin());
    }

    const EditHistoryState state = std::move(_undoHistory.back());
    _undoHistory.pop_back();
    RestoreEditHistoryState(state);
    return true;
}

bool TextField::TryRedoDirectEdit() noexcept
{
    if (_redoHistory.empty())
    {
        return false;
    }

    BeginEditTransactionMetric(L"redo");
    BreakDirectEditMerge();
    _undoHistory.push_back(CaptureEditHistoryState());
    if (_undoHistory.size() > kMaxEditHistoryEntries)
    {
        _undoHistory.erase(_undoHistory.begin());
    }

    const EditHistoryState state = std::move(_redoHistory.back());
    _redoHistory.pop_back();
    RestoreEditHistoryState(state);
    return true;
}

void TextField::RemaskPasswordReveal() noexcept
{
    if (_masked && _passwordRevealMode != PasswordRevealMode::Visible && _passwordRevealState != PasswordRevealState::Hidden)
    {
        SetPasswordRevealState(PasswordRevealState::Hidden);
    }
}

void TextField::SetCaretIndex(size_t caretIndex, bool extendSelection) noexcept
{
    _preferredMultilineXOffsetDip.reset();
    BreakDirectEditMerge();
    SetSingleLineCaretIndex(_caretIndex, _selectionAnchorIndex, std::min(caretIndex, _text.size()), extendSelection);
}

bool TextField::HasSelection() const noexcept
{
    return GetSelectionRange().has_value();
}

std::optional<std::pair<size_t, size_t>> TextField::GetSelectionRange() const noexcept
{
    return GetSingleLineSelectionRange(_selectionAnchorIndex, _caretIndex);
}

bool TextField::DeleteSelection() noexcept
{
    return DeleteSingleLineSelection(_text, _caretIndex, _selectionAnchorIndex);
}

void TextField::SelectAllText() noexcept
{
    BreakDirectEditMerge();
    SelectAllSingleLineText(_text.size(), _caretIndex, _selectionAnchorIndex);
}

void TextField::SelectWordAt(size_t hitIndex) noexcept
{
    BreakDirectEditMerge();
    SelectSingleLineWordAt(_text, hitIndex, _caretIndex, _selectionAnchorIndex);
}

D2D1_RECT_F TextField::GetTextRect() const noexcept
{
    const bool hasTrailingButton = IsClearButtonVisible() || IsPasswordRevealButtonVisible();
    const float rightInset       = hasTrailingButton ? std::max(38.0f, _textPaddingRightDip) : _textPaddingRightDip; // 30 DIP button + padding
    const float verticalPadding  = IsCompactDensity() ? 2.0f : 4.0f;
    const float topInset         = _hasExplicitVerticalTextPadding ? _textPaddingTopDip : verticalPadding;
    const float bottomInset      = _hasExplicitVerticalTextPadding ? _textPaddingBottomDip : verticalPadding;
    return D2D1::RectF(GetBounds().left + _textPaddingLeftDip, GetBounds().top + topInset, GetBounds().right - rightInset, GetBounds().bottom - bottomInset);
}

bool TextField::IsClearButtonVisible() const noexcept
{
    return _clearButtonEnabled && ! _masked && ! _readOnly && ! _multiline && ! _text.empty() && HasFocus();
}

D2D1_RECT_F TextField::GetClearButtonRect() const noexcept
{
    constexpr float kClearButtonWidthDip = 30.0f;
    return D2D1::RectF(GetBounds().right - kClearButtonWidthDip, GetBounds().top, GetBounds().right, GetBounds().bottom);
}

bool TextField::IsPasswordRevealButtonVisible() const noexcept
{
    return _passwordRevealMode == PasswordRevealMode::Peek && _masked && IsEnabled() && ! _readOnly && ! _multiline && ! _text.empty() && HasFocus();
}

D2D1_RECT_F TextField::GetPasswordRevealButtonRect() const noexcept
{
    constexpr float kRevealButtonWidthDip = 30.0f;
    return D2D1::RectF(GetBounds().right - kRevealButtonWidthDip, GetBounds().top, GetBounds().right, GetBounds().bottom);
}

std::wstring TextField::GetDisplayText() const
{
    if (! _masked || _text.empty() || _passwordRevealMode == PasswordRevealMode::Visible || _passwordRevealState == PasswordRevealState::Visible)
    {
        return _text;
    }

    const size_t dotCount = GetSecretVisibleDotCount();
    return std::wstring(dotCount, L'\u2022');
}

bool TextField::UsesMaskedDisplayText() const noexcept
{
    return _masked && ! _text.empty() && _passwordRevealMode != PasswordRevealMode::Visible && _passwordRevealState != PasswordRevealState::Visible;
}

const std::vector<size_t>& TextField::GetMaskedSourceTextElementBoundaries() const noexcept
{
    if (! _maskedSourceTextElementBoundaries.empty())
    {
        return _maskedSourceTextElementBoundaries;
    }

    const bool perfEnabled = Debug::Perf::IsCaptureEnabled();
    const auto startedAt   = perfEnabled ? std::chrono::steady_clock::now() : std::chrono::steady_clock::time_point{};
    _maskedSourceTextElementBoundaries.push_back(0u);
    size_t sourceIndex = 0u;
    while (sourceIndex < _text.size())
    {
        const size_t nextIndex = StepToNextTextElement(_text, sourceIndex);
        if (nextIndex <= sourceIndex)
        {
            _maskedSourceTextElementBoundaries.push_back(_text.size());
            break;
        }
        _maskedSourceTextElementBoundaries.push_back(nextIndex);
        sourceIndex = nextIndex;
    }
    if (perfEnabled)
    {
        Debug::Perf::Emit(L"dxui.textinput.masked_index_map_rebuild_us",
                          _maskLengthPolicy == PasswordMaskLengthPolicy::Exact ? L"exact" : L"concealed",
                          Debug::Perf::ElapsedUs(startedAt),
                          _text.size(),
                          _maskedSourceTextElementBoundaries.size() - 1u,
                          S_OK);
    }
    return _maskedSourceTextElementBoundaries;
}

size_t TextField::ControlTextIndexToDisplayTextIndex(size_t controlTextIndex) const noexcept
{
    const size_t clampedControlIndex = (std::min)(controlTextIndex, _text.size());
    if (! UsesMaskedDisplayText())
    {
        return clampedControlIndex;
    }

    const std::vector<size_t>& boundaries = GetMaskedSourceTextElementBoundaries();
    const size_t sourceElementCount       = boundaries.size() - 1u;
    const size_t snappedControlIndex      = SnapCaretIndexToTextElementBoundary(_text, clampedControlIndex);
    const auto boundary                   = std::lower_bound(boundaries.begin(), boundaries.end(), snappedControlIndex);
    const size_t sourceElementIndex       = boundary == boundaries.end() ? sourceElementCount : static_cast<size_t>(boundary - boundaries.begin());
    const size_t displayLength            = _maskLengthPolicy == PasswordMaskLengthPolicy::Exact ? sourceElementCount : GetSecretVisibleDotCount();
    return ScaleMaskedTextIndex(sourceElementIndex, sourceElementCount, displayLength);
}

size_t TextField::DisplayTextIndexToControlTextIndex(size_t displayTextIndex) const noexcept
{
    if (! UsesMaskedDisplayText())
    {
        return (std::min)(displayTextIndex, _text.size());
    }

    const std::vector<size_t>& boundaries = GetMaskedSourceTextElementBoundaries();
    const size_t sourceElementCount       = boundaries.size() - 1u;
    const size_t displayLength            = _maskLengthPolicy == PasswordMaskLengthPolicy::Exact ? sourceElementCount : GetSecretVisibleDotCount();
    const size_t sourceElementIndex       = ScaleMaskedTextIndex((std::min)(displayTextIndex, displayLength), displayLength, sourceElementCount);
    return boundaries[(std::min)(sourceElementIndex, sourceElementCount)];
}

std::optional<std::pair<size_t, size_t>> TextField::ControlTextRangeToDisplayTextRange(size_t controlTextStartIndex, size_t controlTextEndIndex) const noexcept
{
    if (controlTextStartIndex >= controlTextEndIndex)
    {
        return std::nullopt;
    }

    const size_t displayStart = ControlTextIndexToDisplayTextIndex(controlTextStartIndex);
    const size_t displayEnd   = ControlTextIndexToDisplayTextIndex(controlTextEndIndex);
    if (displayStart >= displayEnd)
    {
        return std::nullopt;
    }
    return std::pair<size_t, size_t>{displayStart, displayEnd};
}

wil::com_ptr<IDWriteTextLayout> TextField::GetOrCreateMultilineLayout(const WindowHost* host,
                                                                      std::wstring_view text,
                                                                      float widthDip,
                                                                      float heightDip) const noexcept
{
    if (! host || ! _multiline)
    {
        return CreateMultilineTextLayout(host, text, FontRole::Body, widthDip, heightDip);
    }

    const D2D1_SIZE_F desiredSize = D2D1::SizeF(widthDip, heightDip);
    const bool sizeChanged = std::abs(_cachedLayoutSize.width - desiredSize.width) > 0.5f || std::abs(_cachedLayoutSize.height - desiredSize.height) > 0.5f;
    const bool textChanged = text != _cachedLayoutText;

    if (_multilineLayoutDirty || sizeChanged || textChanged || ! _cachedMultilineLayout)
    {
        _cachedMultilineLayout = CreateMultilineTextLayout(host, text, FontRole::Body, widthDip, heightDip);
        _cachedLayoutText      = std::wstring(text);
        _cachedLayoutSize      = desiredSize;
        _multilineLayoutDirty  = false;
    }

    return _cachedMultilineLayout;
}

wil::com_ptr<IDWriteTextLayout> TextField::GetOrCreateSingleLineLayout(
    const WindowHost* host, std::wstring_view text, float minimumWidthDip, float heightDip, DWRITE_READING_DIRECTION readingDirection) const noexcept
{
    if (_multiline)
    {
        return GetOrCreateSingleLineTextLayout(host, nullptr, text, FontRole::Body, minimumWidthDip, heightDip, readingDirection);
    }

    return GetOrCreateSingleLineTextLayout(host, &_singleLineLayoutCache, text, FontRole::Body, minimumWidthDip, heightDip, readingDirection, true);
}

void TextField::InvalidateSingleLineLayoutCache() const noexcept
{
    ClearSingleLineTextLayoutCache(_singleLineLayoutCache, true);
    _maskedSourceTextElementBoundaries.clear();
}

void TextField::InvalidateMultilineLayoutCache() const noexcept
{
    _multilineLayoutDirty = true;
    _cachedMultilineLayout.reset();
    _cachedLayoutText.clear();
}

void TextField::EnsureCaretVisible(const WindowHost* host, float availableWidthDip) const noexcept
{
    if (_multiline || _text.empty())
    {
        _horizontalScrollDip = 0.0f;
        return;
    }

    const std::wstring displayText               = GetDisplayText();
    const wil::com_ptr<IDWriteTextLayout> layout = GetOrCreateSingleLineLayout(host,
                                                                               displayText,
                                                                               std::max(1.0f, availableWidthDip + _horizontalScrollDip),
                                                                               std::max(1.0f, GetTextRect().bottom - GetTextRect().top),
                                                                               ResolveReadingDirection(GetFlowDirection()));
    const float caretOffset                      = MeasureCaretOffsetDip(layout.get(), displayText, ControlTextIndexToDisplayTextIndex(_caretIndex));
    const float padding                          = 6.0f;
    if (caretOffset < _horizontalScrollDip + padding)
    {
        _horizontalScrollDip = std::max(0.0f, caretOffset - padding);
    }
    else if (caretOffset > _horizontalScrollDip + availableWidthDip - padding)
    {
        _horizontalScrollDip = std::max(0.0f, caretOffset - availableWidthDip + padding);
    }
}

void TextField::EnsureMultilineCaretVisible(const WindowHost* host) noexcept
{
    if (! _multiline)
    {
        _multilineFirstVisibleLine = 0u;
        return;
    }

    const D2D1_RECT_F textRect     = GetTextRect();
    const std::wstring displayText = GetDisplayText();
    if (displayText.empty())
    {
        _multilineFirstVisibleLine = 0u;
        return;
    }

    const float viewportHeightDip = std::max(1.0f, textRect.bottom - textRect.top);
    const auto multilineLayout =
        CreateMultilineTextLayout(host, displayText, FontRole::Body, std::max(1.0f, textRect.right - textRect.left), viewportHeightDip);
    const std::vector<DWRITE_LINE_METRICS> lineMetrics = GetMultilineLineMetrics(multilineLayout.get());
    const float fallbackLineHeightDip                  = EstimateMultilineFallbackLineHeightDip(host, FontRole::Body);
    const MultilineViewportMetrics viewportMetrics     = BuildMultilineViewportMetrics(displayText, viewportHeightDip, lineMetrics, fallbackLineHeightDip);
    const size_t displayCaretIndex                     = ControlTextIndexToDisplayTextIndex(_caretIndex);
    const size_t caretLine =
        TryGetMultilineCaretLineIndex(multilineLayout.get(), lineMetrics, displayCaretIndex, displayText.size())
            .value_or(static_cast<size_t>(
                std::count(displayText.begin(), std::next(displayText.begin(), static_cast<std::wstring::difference_type>(displayCaretIndex)), L'\n')));
    const size_t maxFirstVisibleLine = ComputeMultilineMaxFirstVisibleLine(viewportMetrics);
    size_t firstVisibleLine          = ClampMultilineFirstVisibleLine(_multilineFirstVisibleLine, viewportMetrics);
    if (caretLine < firstVisibleLine)
    {
        firstVisibleLine = caretLine;
    }
    else if (caretLine >= firstVisibleLine + viewportMetrics.visibleLineCount)
    {
        firstVisibleLine = caretLine - viewportMetrics.visibleLineCount + 1u;
    }

    _multilineFirstVisibleLine = (std::min)(firstVisibleLine, maxFirstVisibleLine);
}

void TextField::OnFlowDirectionChanged() noexcept
{
    InvalidateSingleLineLayoutCache();
    Control::OnFlowDirectionChanged();
}

void TextField::OnHostDpiChanged(WindowHost& host) noexcept
{
    Control::OnHostDpiChanged(host);
    InvalidateSingleLineLayoutCache();
    InvalidateMultilineLayoutCache();
    if (SupportsTextInput() && host.GetFocusControl() == this)
    {
        host.SyncTextInput(this);
    }
}

void TextField::OnDensityChanged() noexcept
{
    InvalidateSingleLineLayoutCache();
    InvalidateMultilineLayoutCache();
    Control::OnDensityChanged();
}

void TextField::ResetCaretBlink(WindowHost& host) noexcept
{
    _caretBlinkAnchorTickMs = ::GetTickCount64();
    _caretVisible           = true;
    host.RequestAnimation();
}

bool TextField::NotifyChanged()
{
    InvalidateSingleLineLayoutCache();
    InvalidateMultilineLayoutCache();
    FinishEditTransactionMetric();

    const std::function<void(std::wstring_view)> onTextChanged = _onTextChanged;
    if (! onTextChanged)
    {
        return true;
    }

    std::wstring textSnapshot             = _text;
    const bool secureSnapshot             = _masked;
    const std::weak_ptr<int> selfLifetime = GetLifetimeToken();
    onTextChanged(textSnapshot);
    if (secureSnapshot)
    {
        SecureWipe::SecureClear(textSnapshot);
    }
    return ! selfLifetime.expired();
}
} // namespace RedSalamander::DxUi
