#ifndef NOMINMAX
#define NOMINMAX
#endif

#include "DxUi.Internal.h"

#include <algorithm>
#include <cmath>
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

[[nodiscard]] bool IsWordCharacter(wchar_t ch) noexcept
{
    return HasCharTypeFlag(ch, C1_ALPHA | C1_DIGIT) || ch == L'_';
}

[[nodiscard]] bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
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

[[nodiscard]] float MeasureSingleLineTextWidthDip(const WindowHost* host, std::wstring_view text, FontRole role, float heightDip) noexcept
{
    if (text.empty())
    {
        return 0.0f;
    }

    if (host)
    {
        auto* factory = host->GetWriteFactory();
        auto* format  = host->GetTextFormat(role);
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

[[nodiscard]] wil::com_ptr<IDWriteTextLayout> CreateSingleLineTextLayout(
    const WindowHost* host, std::wstring_view text, FontRole role, float widthDip, float heightDip) noexcept
{
    if (! host)
    {
        return {};
    }

    auto* factory = host->GetWriteFactory();
    auto* format  = host->GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false);
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

[[nodiscard]] float MeasureCaretOffsetDip(const WindowHost* host, std::wstring_view text, FontRole role, size_t caretIndex, float heightDip) noexcept
{
    const size_t clampedCaret = std::min(caretIndex, text.size());
    if (wil::com_ptr<IDWriteTextLayout> layout = CreateSingleLineTextLayout(host, text, role, 32768.0f, heightDip))
    {
        float x = 0.0f;
        float y = 0.0f;
        DWRITE_HIT_TEST_METRICS metrics{};
        if (SUCCEEDED(layout->HitTestTextPosition(static_cast<UINT32>(clampedCaret), FALSE, &x, &y, &metrics)))
        {
            return x;
        }
    }

    return static_cast<float>(clampedCaret) * 7.0f;
}

[[nodiscard]] size_t HitTestCaretIndexDip(
    const WindowHost* host, std::wstring_view text, FontRole role, const D2D1_RECT_F& textRect, float scrollDip, D2D1_POINT_2F point) noexcept
{
    if (text.empty())
    {
        return 0u;
    }

    const float localX = std::max(0.0f, point.x - textRect.left + scrollDip);
    const float localY = std::clamp(point.y - textRect.top, 0.0f, std::max(1.0f, textRect.bottom - textRect.top) - 1.0f);
    if (wil::com_ptr<IDWriteTextLayout> layout =
            CreateSingleLineTextLayout(host, text, role, std::max(32768.0f, localX + 32.0f), std::max(1.0f, textRect.bottom - textRect.top)))
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

    const double fallbackValue = std::clamp(std::floor(static_cast<double>(localX) / 7.0), 0.0, static_cast<double>(text.size()));
    return static_cast<size_t>(fallbackValue);
}

void DrawSingleLineTextClipped(
    WindowHost& host, std::wstring_view text, const D2D1_RECT_F& rect, FontRole role, const D2D1_COLOR_F& color, float scrollDip) noexcept
{
    auto* dc    = host.GetDeviceContext();
    auto* brush = host.GetSolidBrush(color);
    if (! dc || ! brush)
    {
        return;
    }

    const D2D1_RECT_F snappedRect = SnapRectToPixel(host, rect);
    if (wil::com_ptr<IDWriteTextLayout> layout = CreateSingleLineTextLayout(
            &host, text, role, std::max(1.0f, MeasureSingleLineTextWidthDip(&host, text, role) + 32.0f), std::max(1.0f, snappedRect.bottom - snappedRect.top)))
    {
        dc->PushAxisAlignedClip(snappedRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
        dc->DrawTextLayout(D2D1::Point2F(snappedRect.left - scrollDip, snappedRect.top), layout.get(), brush, kTextDrawOptions);
        dc->PopAxisAlignedClip();
        return;
    }

    if (auto* format = host.GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false))
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
                                                               std::optional<std::pair<size_t, size_t>> selectionRange) noexcept
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
    const float startOffsetDip    = MeasureCaretOffsetDip(&host, text, role, selectionStart, heightDip);
    const float endOffsetDip      = MeasureCaretOffsetDip(&host, text, role, selectionEnd, heightDip);
    D2D1_RECT_F selectionRect     = D2D1::RectF(
        snappedRect.left + startOffsetDip - scrollDip, snappedRect.top + 1.0f, snappedRect.left + endOffsetDip - scrollDip, snappedRect.bottom - 1.0f);
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
                             std::optional<std::pair<size_t, size_t>> selectionRange) noexcept
{
    DrawSingleLineTextClipped(host, text, rect, role, textColor, scrollDip);

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
    const std::optional<D2D1_RECT_F> selectionRect = ComputeSingleLineSelectionPaintRect(host, text, rect, role, scrollDip, selectionRange);
    if (! selectionRect)
    {
        return;
    }

    dc->FillRectangle(selectionRect.value(), fillBrush);
    dc->PushAxisAlignedClip(selectionRect.value(), D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
    if (wil::com_ptr<IDWriteTextLayout> layout =
            CreateSingleLineTextLayout(&host, text, role, std::max(1.0f, MeasureSingleLineTextWidthDip(&host, text, role) + 32.0f), heightDip))
    {
        dc->DrawTextLayout(D2D1::Point2F(snappedRect.left - scrollDip, snappedRect.top), layout.get(), textBrush, kTextDrawOptions);
    }
    else if (auto* format = host.GetTextFormat(role, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false))
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
