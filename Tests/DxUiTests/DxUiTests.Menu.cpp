#include "DxUi/DxUiNativeMenuInterop.h"
#include "DxUiTestHelpers.h"
#include "FolderViewEmptyStateLayout.h"
#include "FolderViewIncrementalSearch.h"
#include "FolderViewVisualState.h"

#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <format>
#include <iterator>
#include <string>
#include <thread>
#include <vector>

namespace
{
[[nodiscard]] std::string NarrowAsciiForFailureMessage(std::wstring_view text)
{
    std::string result;
    result.reserve(text.size());
    for (const wchar_t ch : text)
    {
        const auto value = static_cast<unsigned int>(ch);
        result.push_back(value >= 0x20u && value <= 0x7eu ? static_cast<char>(value) : '?');
    }

    return result;
}

[[nodiscard]] std::string WideCodeUnitsForFailureMessage(std::wstring_view text)
{
    std::string result;
    for (const wchar_t ch : text)
    {
        std::format_to(std::back_inserter(result), "{:04X} ", static_cast<unsigned int>(ch));
    }

    return result;
}

void TestFolderViewIncrementalSearchKeepsContainsHighlightButUsesPrefixFocus()
{
    using namespace FolderViewIncrementalSearch;

    constexpr std::array<std::wstring_view, 4> names{{
        L"notes-a.txt",
        L"zeta.txt",
        L"ABC.txt",
        L"beta-a.txt",
    }};

    const auto displayNameAt = [&](size_t index) noexcept -> std::wstring_view { return index < names.size() ? names[index] : std::wstring_view{}; };

    const std::optional<UINT32> containsOffset = FindContainsOffsetNoCase(names[0], L"a");
    Require(containsOffset.has_value() && containsOffset.value() > 0u, "contains match remains available for incremental-search highlighting");
    Require(! StartsWithNoCase(names[0], L"a"), "a non-prefix contains match is not treated as a focus target");
    Require(StartsWithNoCase(names[2], L"a"), "prefix matching is case-insensitive");

    const std::optional<size_t> prefixIndex = FindNextPrefixMatchIndex(names.size(), 0u, true, displayNameAt, L"a");
    Require(prefixIndex.has_value() && prefixIndex.value() == 2u, "focus navigation prefers the next item whose name starts with the query");
}

void TestFolderViewInactiveVisualStateDimsNormalTextAndIcons()
{
    using namespace FolderViewVisualState;

    const auto same = [](float left, float right) noexcept { return std::fabs(left - right) <= 0.0001f; };

    Require(same(ResolveNormalTextAlpha(1.0f, true, false), 1.0f), "focused pane keeps normal text fully opaque");
    Require(same(ResolveNormalTextAlpha(1.0f, false, false), kUnfocusedPaneTextOpacity), "unfocused pane dims normal text");
    Require(same(ResolveNormalTextAlpha(1.0f, false, true), 1.0f), "selected text uses inactive-selection colors instead of normal text dimming");
    Require(same(ResolveNormalIconOpacity(1.0f, true), 1.0f), "focused pane keeps normal icons fully opaque");
    Require(same(ResolveNormalIconOpacity(1.0f, false), kUnfocusedPaneIconOpacity), "unfocused pane dims normal icons");
    Require(same(ResolveNormalIconOpacity(0.5f, false), 0.5f * kUnfocusedPaneIconOpacity), "hidden icons keep their hidden dim and get pane dimming");
    Require(same(ResolvePlaceholderIconOpacity(false), 0.4f * kUnfocusedPaneIconOpacity), "placeholder icons also dim in an unfocused pane");
    Require(same(ResolveFocusBorderAlpha(1.0f, false), kFocusBorderOpacityUnfocused), "unfocused current item keeps a dim focus border");
}

void TestFolderViewEmptyPlaceholderMetricsUseCurrentEmptyLayout()
{
    using namespace FolderViewEmptyStateLayout;

    constexpr PlaceholderItemMetricsInput brief{
        .clientWidthDip          = 640.0f,
        .clientHeightDip         = 360.0f,
        .iconSizeDip             = 16.0f,
        .estimatedCharWidthDip   = 8.0f,
        .estimatedLabelHeightDip = 18.0f,
        .detailsLineHeightDip    = 12.0f,
        .metadataLineHeightDip   = 10.0f,
        .titleLength             = 12u,
        .includeDetailsLine      = false,
        .includeMetadataLine     = false,
    };
    const auto expectedTileHeight = [](const PlaceholderItemMetricsInput& input) noexcept
    {
        constexpr float kLabelVerticalPaddingDip = 4.0f;
        constexpr float kDetailsGapDip           = 2.0f;

        float textBlockHeightDip = input.estimatedLabelHeightDip;
        if (input.includeDetailsLine)
        {
            textBlockHeightDip += kDetailsGapDip + input.detailsLineHeightDip;
        }
        if (input.includeMetadataLine)
        {
            textBlockHeightDip += kDetailsGapDip + input.metadataLineHeightDip;
        }

        return (std::max)(input.iconSizeDip, textBlockHeightDip) + (kLabelVerticalPaddingDip * 2.0f);
    };

    const PlaceholderItemMetrics briefMetrics = ResolvePlaceholderItemMetrics(brief);
    Require(briefMetrics.tileWidthDip == brief.clientWidthDip, "empty-folder placeholder uses the current client width as a full-view focus item");
    Require(briefMetrics.tileHeightDip == expectedTileHeight(brief), "empty-folder placeholder uses the current display-mode row height");

    PlaceholderItemMetricsInput narrow         = brief;
    narrow.clientWidthDip                      = 80.0f;
    const PlaceholderItemMetrics narrowMetrics = ResolvePlaceholderItemMetrics(narrow);
    Require(narrowMetrics.tileWidthDip == 80.0f, "empty-folder placeholder width follows the current client width");

    PlaceholderItemMetricsInput zeroWidth         = brief;
    zeroWidth.clientWidthDip                      = 0.0f;
    const PlaceholderItemMetrics zeroWidthMetrics = ResolvePlaceholderItemMetrics(zeroWidth);
    Require(zeroWidthMetrics.tileWidthDip == 0.0f && zeroWidthMetrics.tileHeightDip == 0.0f && zeroWidthMetrics.labelHeightDip == 0.0f,
            "empty-folder placeholder metrics clear when the current client width is zero");

    PlaceholderItemMetricsInput zeroHeight         = brief;
    zeroHeight.clientHeightDip                     = 0.0f;
    const PlaceholderItemMetrics zeroHeightMetrics = ResolvePlaceholderItemMetrics(zeroHeight);
    Require(zeroHeightMetrics.tileWidthDip == 0.0f && zeroHeightMetrics.tileHeightDip == 0.0f && zeroHeightMetrics.labelHeightDip == 0.0f,
            "empty-folder placeholder metrics clear when the current client height is zero");

    PlaceholderItemMetricsInput detailed         = brief;
    detailed.includeDetailsLine                  = true;
    const PlaceholderItemMetrics detailedMetrics = ResolvePlaceholderItemMetrics(detailed);
    Require(detailedMetrics.tileHeightDip == expectedTileHeight(detailed), "empty-folder placeholder height follows detailed display-mode row height");

    PlaceholderItemMetricsInput extraDetailed         = detailed;
    extraDetailed.includeMetadataLine                 = true;
    const PlaceholderItemMetrics extraDetailedMetrics = ResolvePlaceholderItemMetrics(extraDetailed);
    Require(extraDetailedMetrics.tileHeightDip == expectedTileHeight(extraDetailed),
            "empty-folder placeholder height follows extra-detailed display-mode row height");
}

class StripedBackdropControl final : public RedSalamander::DxUi::Control
{
public:
    explicit StripedBackdropControl(LONG stripeWidthPx = 12) noexcept : _stripeWidthPx((std::max)(1l, stripeWidthPx))
    {
    }

    void Paint(RedSalamander::DxUi::WindowHost& host) const override
    {
        auto* const dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }

        const D2D1_RECT_F bounds      = host.GetClientBoundsDip();
        const LONG left               = static_cast<LONG>(std::floor(bounds.left));
        const LONG right              = static_cast<LONG>(std::ceil(bounds.right));
        const D2D1_COLOR_F evenStripe = D2D1::ColorF(0.94f, 0.97f, 1.00f, 1.0f);
        const D2D1_COLOR_F oddStripe  = D2D1::ColorF(0.03f, 0.16f, 0.34f, 1.0f);

        for (LONG x = left; x < right; ++x)
        {
            const LONG stripeIndex = (x - left) / _stripeWidthPx;
            if (auto* const brush = host.GetSolidBrush((stripeIndex & 1) == 0 ? evenStripe : oddStripe))
            {
                dc->FillRectangle(D2D1::RectF(static_cast<float>(x), bounds.top, static_cast<float>(x + 1), bounds.bottom), brush);
            }
        }
    }

private:
    LONG _stripeWidthPx = 12;
};

template <typename TPredicate>
bool WaitForContextMenuPopupState(HWND popupHwnd,
                                  TPredicate&& predicate,
                                  RedSalamander::DxUi::ContextMenuPopupDebugState& outState,
                                  std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popupHwnd, outState) && predicate(outState))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

bool WaitForContextMenuPopupItemRect(HWND popupHwnd,
                                     size_t itemIndex,
                                     D2D1_RECT_F& outRectDip,
                                     std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugGetContextMenuPopupItemRect(popupHwnd, itemIndex, outRectDip))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

bool WaitForContextMenuPopupItemLayout(HWND popupHwnd,
                                       size_t itemIndex,
                                       RedSalamander::DxUi::ContextMenuPopupItemLayoutDebugState& outLayout,
                                       std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugGetContextMenuPopupItemLayout(popupHwnd, itemIndex, outLayout))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

bool WaitForContextMenuPopupItemPaint(HWND popupHwnd,
                                      size_t itemIndex,
                                      RedSalamander::DxUi::ContextMenuPopupItemPaintDebugState& outPaint,
                                      std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugGetContextMenuPopupItemPaint(popupHwnd, itemIndex, outPaint))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

bool WaitForContextMenuPopupItemText(HWND popupHwnd,
                                     size_t itemIndex,
                                     std::wstring& outText,
                                     std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugGetContextMenuPopupItemText(popupHwnd, itemIndex, outText))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    outText.clear();
    return false;
}

bool WaitForContextMenuPopupBitmapCapture(HWND popupHwnd,
                                          RedSalamander::DxUi::WindowHostBitmapCapture& outCapture,
                                          std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugCaptureContextMenuPopupBitmap(popupHwnd, outCapture) && outCapture.widthPx > 0u && outCapture.heightPx > 0u &&
            ! outCapture.bgraPixels.empty())
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

[[nodiscard]] uint8_t GetCapturePixelAlpha(const RedSalamander::DxUi::WindowHostBitmapCapture& capture, UINT xPx, UINT yPx) noexcept
{
    if (xPx >= capture.widthPx || yPx >= capture.heightPx)
    {
        return 0u;
    }

    const size_t base = (static_cast<size_t>(yPx) * static_cast<size_t>(capture.widthPx) + static_cast<size_t>(xPx)) * 4u;
    if ((base + 3u) >= capture.bgraPixels.size())
    {
        return 0u;
    }

    return capture.bgraPixels[base + 3u];
}

[[nodiscard]] uint32_t GetCapturePixelBgra(const RedSalamander::DxUi::WindowHostBitmapCapture& capture, UINT xPx, UINT yPx) noexcept
{
    if (xPx >= capture.widthPx || yPx >= capture.heightPx)
    {
        return 0u;
    }

    const size_t base = (static_cast<size_t>(yPx) * static_cast<size_t>(capture.widthPx) + static_cast<size_t>(xPx)) * 4u;
    if ((base + 3u) >= capture.bgraPixels.size())
    {
        return 0u;
    }

    return static_cast<uint32_t>(capture.bgraPixels[base + 0u]) | (static_cast<uint32_t>(capture.bgraPixels[base + 1u]) << 8u) |
           (static_cast<uint32_t>(capture.bgraPixels[base + 2u]) << 16u) | (static_cast<uint32_t>(capture.bgraPixels[base + 3u]) << 24u);
}

[[nodiscard]] UINT DipToPixelForPopup(float dip, UINT dpi) noexcept
{
    return static_cast<UINT>((std::max)(0l, std::lround(static_cast<double>(dip) * static_cast<double>(dpi) / 96.0)));
}

[[nodiscard]] POINT ClientScreenPointForTest(HWND hwnd, LONG x, LONG y, const char* context)
{
    POINT point{x, y};
    Require(ClientToScreen(hwnd, &point) != FALSE, context);
    return point;
}

RedSalamander::DxUi::WindowHostBitmapCapture CropWindowHostBitmapCaptureForTest(const RedSalamander::DxUi::WindowHostBitmapCapture& source,
                                                                                const RECT& cropRect)
{
    RedSalamander::DxUi::WindowHostBitmapCapture capture{};
    const LONG widthPx        = cropRect.right - cropRect.left;
    const LONG heightPx       = cropRect.bottom - cropRect.top;
    const LONG sourceWidthPx  = static_cast<LONG>(source.widthPx);
    const LONG sourceHeightPx = static_cast<LONG>(source.heightPx);
    const bool cropIsInBounds =
        widthPx > 0 && heightPx > 0 && cropRect.left >= 0 && cropRect.top >= 0 && cropRect.right <= sourceWidthPx && cropRect.bottom <= sourceHeightPx;
    const std::string cropBoundsMessage = std::format("popup backdrop crop stays inside the owner capture crop=({},{} {}x{}) source={}x{}",
                                                      cropRect.left,
                                                      cropRect.top,
                                                      widthPx,
                                                      heightPx,
                                                      sourceWidthPx,
                                                      sourceHeightPx);
    Require(cropIsInBounds, cropBoundsMessage.c_str());

    capture.widthPx  = static_cast<UINT>(widthPx);
    capture.heightPx = static_cast<UINT>(heightPx);
    capture.bgraPixels.resize(static_cast<size_t>(capture.widthPx) * static_cast<size_t>(capture.heightPx) * 4u);

    const size_t destinationStride = static_cast<size_t>(capture.widthPx) * 4u;
    const size_t sourceStride      = static_cast<size_t>(source.widthPx) * 4u;
    for (LONG y = 0; y < heightPx; ++y)
    {
        const size_t sourceOffset      = (static_cast<size_t>(cropRect.top + y) * sourceStride) + (static_cast<size_t>(cropRect.left) * 4u);
        const size_t destinationOffset = static_cast<size_t>(y) * destinationStride;
        std::copy_n(source.bgraPixels.data() + sourceOffset, destinationStride, capture.bgraPixels.data() + destinationOffset);
    }

    return capture;
}

[[nodiscard]] RECT FindOpaqueBoundsInCapture(const RedSalamander::DxUi::WindowHostBitmapCapture& capture, uint8_t alphaThreshold = 8u) noexcept
{
    RECT bounds{static_cast<LONG>(capture.widthPx), static_cast<LONG>(capture.heightPx), 0, 0};
    bool foundOpaque = false;

    for (UINT y = 0u; y < capture.heightPx; ++y)
    {
        for (UINT x = 0u; x < capture.widthPx; ++x)
        {
            if (GetCapturePixelAlpha(capture, x, y) <= alphaThreshold)
            {
                continue;
            }

            bounds.left   = (std::min)(bounds.left, static_cast<LONG>(x));
            bounds.top    = (std::min)(bounds.top, static_cast<LONG>(y));
            bounds.right  = (std::max)(bounds.right, static_cast<LONG>(x + 1u));
            bounds.bottom = (std::max)(bounds.bottom, static_cast<LONG>(y + 1u));
            foundOpaque   = true;
        }
    }

    return foundOpaque ? bounds : RECT{};
}

[[nodiscard]] RECT ComputeRightStripSampleRect(const RECT& opaqueBounds) noexcept
{
    const LONG widthPx  = opaqueBounds.right - opaqueBounds.left;
    const LONG heightPx = opaqueBounds.bottom - opaqueBounds.top;
    if (widthPx <= 24 || heightPx <= 24)
    {
        return RECT{};
    }

    const LONG horizontalInset = (std::max)(10l, widthPx / 12l);
    const LONG verticalInset   = (std::max)(10l, heightPx / 6l);
    const LONG sampleLeft      = opaqueBounds.left + (widthPx * 5l / 8l);
    const LONG sampleRight     = opaqueBounds.right - horizontalInset;
    const LONG sampleTop       = opaqueBounds.top + verticalInset;
    const LONG sampleBottom    = opaqueBounds.bottom - verticalInset;

    return RECT{sampleLeft, sampleTop, (std::max)(sampleLeft + 2l, sampleRight), (std::max)(sampleTop + 2l, sampleBottom)};
}

[[nodiscard]] uint64_t ComputeAverageAdjacentRgbDelta(const RedSalamander::DxUi::WindowHostBitmapCapture& capture, const RECT& sampleRect) noexcept
{
    if (sampleRect.right - sampleRect.left < 2 || sampleRect.bottom - sampleRect.top < 1)
    {
        return 0u;
    }

    uint64_t accumulatedDelta = 0u;
    uint64_t sampleCount      = 0u;

    const LONG clampedLeft   = (std::clamp)(sampleRect.left, 0l, static_cast<LONG>(capture.widthPx));
    const LONG clampedTop    = (std::clamp)(sampleRect.top, 0l, static_cast<LONG>(capture.heightPx));
    const LONG clampedRight  = (std::clamp)(sampleRect.right, clampedLeft, static_cast<LONG>(capture.widthPx));
    const LONG clampedBottom = (std::clamp)(sampleRect.bottom, clampedTop, static_cast<LONG>(capture.heightPx));

    for (LONG y = clampedTop; y < clampedBottom; ++y)
    {
        for (LONG x = clampedLeft + 1; x < clampedRight; ++x)
        {
            const size_t currentBase  = (static_cast<size_t>(y) * static_cast<size_t>(capture.widthPx) + static_cast<size_t>(x)) * 4u;
            const size_t previousBase = currentBase - 4u;
            accumulatedDelta += static_cast<uint64_t>(
                std::abs(static_cast<int>(capture.bgraPixels[currentBase + 0u]) - static_cast<int>(capture.bgraPixels[previousBase + 0u])));
            accumulatedDelta += static_cast<uint64_t>(
                std::abs(static_cast<int>(capture.bgraPixels[currentBase + 1u]) - static_cast<int>(capture.bgraPixels[previousBase + 1u])));
            accumulatedDelta += static_cast<uint64_t>(
                std::abs(static_cast<int>(capture.bgraPixels[currentBase + 2u]) - static_cast<int>(capture.bgraPixels[previousBase + 2u])));
            sampleCount += 3u;
        }
    }

    return sampleCount == 0u ? 0u : (accumulatedDelta / sampleCount);
}

[[nodiscard]] uint64_t ComputeAverageAbsoluteRgbDeltaBetweenCaptures(const RedSalamander::DxUi::WindowHostBitmapCapture& lhs,
                                                                     const RedSalamander::DxUi::WindowHostBitmapCapture& rhs,
                                                                     const RECT& sampleRect) noexcept
{
    if (lhs.widthPx != rhs.widthPx || lhs.heightPx != rhs.heightPx)
    {
        return 0u;
    }

    uint64_t accumulatedDelta = 0u;
    uint64_t sampleCount      = 0u;

    const LONG clampedLeft   = (std::clamp)(sampleRect.left, 0l, static_cast<LONG>(lhs.widthPx));
    const LONG clampedTop    = (std::clamp)(sampleRect.top, 0l, static_cast<LONG>(lhs.heightPx));
    const LONG clampedRight  = (std::clamp)(sampleRect.right, clampedLeft, static_cast<LONG>(lhs.widthPx));
    const LONG clampedBottom = (std::clamp)(sampleRect.bottom, clampedTop, static_cast<LONG>(lhs.heightPx));

    for (LONG y = clampedTop; y < clampedBottom; ++y)
    {
        for (LONG x = clampedLeft; x < clampedRight; ++x)
        {
            const size_t base = (static_cast<size_t>(y) * static_cast<size_t>(lhs.widthPx) + static_cast<size_t>(x)) * 4u;
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(lhs.bgraPixels[base + 0u]) - static_cast<int>(rhs.bgraPixels[base + 0u])));
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(lhs.bgraPixels[base + 1u]) - static_cast<int>(rhs.bgraPixels[base + 1u])));
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(lhs.bgraPixels[base + 2u]) - static_cast<int>(rhs.bgraPixels[base + 2u])));
            sampleCount += 3u;
        }
    }

    return sampleCount == 0u ? 0u : (accumulatedDelta / sampleCount);
}

[[nodiscard]] uint64_t ComputeAverageAbsoluteRgbDeltaToColor(const RedSalamander::DxUi::WindowHostBitmapCapture& capture,
                                                             const RECT& sampleRect,
                                                             const D2D1_COLOR_F& color) noexcept
{
    uint64_t accumulatedDelta = 0u;
    uint64_t sampleCount      = 0u;

    const LONG clampedLeft   = (std::clamp)(sampleRect.left, 0l, static_cast<LONG>(capture.widthPx));
    const LONG clampedTop    = (std::clamp)(sampleRect.top, 0l, static_cast<LONG>(capture.heightPx));
    const LONG clampedRight  = (std::clamp)(sampleRect.right, clampedLeft, static_cast<LONG>(capture.widthPx));
    const LONG clampedBottom = (std::clamp)(sampleRect.bottom, clampedTop, static_cast<LONG>(capture.heightPx));

    const int expectedBlue  = static_cast<int>((std::clamp)(color.b, 0.0f, 1.0f) * 255.0f + 0.5f);
    const int expectedGreen = static_cast<int>((std::clamp)(color.g, 0.0f, 1.0f) * 255.0f + 0.5f);
    const int expectedRed   = static_cast<int>((std::clamp)(color.r, 0.0f, 1.0f) * 255.0f + 0.5f);

    for (LONG y = clampedTop; y < clampedBottom; ++y)
    {
        for (LONG x = clampedLeft; x < clampedRight; ++x)
        {
            const size_t base = (static_cast<size_t>(y) * static_cast<size_t>(capture.widthPx) + static_cast<size_t>(x)) * 4u;
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(capture.bgraPixels[base + 0u]) - expectedBlue));
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(capture.bgraPixels[base + 1u]) - expectedGreen));
            accumulatedDelta += static_cast<uint64_t>(std::abs(static_cast<int>(capture.bgraPixels[base + 2u]) - expectedRed));
            sampleCount += 3u;
        }
    }

    return sampleCount == 0u ? 0u : (accumulatedDelta / sampleCount);
}

[[nodiscard]] D2D1_COLOR_F ResolveExpectedAcrylicMenuSlabColor(const RedSalamander::DxUi::ThemePalette& theme) noexcept
{
    return RedSalamander::DxUi::BlendColor(theme.overlayBackground, theme.headerHovered, theme.dark ? 0.14f : 0.10f);
}

HWND FindOwnedContextMenuPopupWindow(HWND ownerHwnd)
{
    HWND popupHwnd               = nullptr;
    const DWORD currentProcessId = GetCurrentProcessId();
    while ((popupHwnd = FindWindowExW(nullptr, popupHwnd, L"DxUi_ContextMenu", nullptr)) != nullptr)
    {
        DWORD popupProcessId = 0;
        static_cast<void>(GetWindowThreadProcessId(popupHwnd, &popupProcessId));
        if (popupProcessId != currentProcessId)
        {
            continue;
        }

        if (GetWindow(popupHwnd, GW_OWNER) == ownerHwnd)
        {
            return popupHwnd;
        }
    }

    return nullptr;
}

HWND WaitForOwnedContextMenuPopupWindow(HWND ownerHwnd, std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (HWND popupHwnd = FindOwnedContextMenuPopupWindow(ownerHwnd))
        {
            return popupHwnd;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return nullptr;
}

void DismissOwnedContextMenuPopupChain(HWND ownerHwnd) noexcept
{
    for (int attempt = 0; attempt < 4; ++attempt)
    {
        if (HWND popupHwnd = FindOwnedContextMenuPopupWindow(ownerHwnd))
        {
            PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            std::this_thread::sleep_for(std::chrono::milliseconds(30));
            continue;
        }

        break;
    }
}

HWND FindOwnedContextMenuPopupWindowByFirstItemText(HWND ownerHwnd, std::wstring_view firstItemText)
{
    HWND popupHwnd               = nullptr;
    const DWORD currentProcessId = GetCurrentProcessId();
    while ((popupHwnd = FindWindowExW(nullptr, popupHwnd, L"DxUi_ContextMenu", nullptr)) != nullptr)
    {
        DWORD popupProcessId = 0;
        static_cast<void>(GetWindowThreadProcessId(popupHwnd, &popupProcessId));
        if (popupProcessId != currentProcessId)
        {
            continue;
        }

        if (GetWindow(popupHwnd, GW_OWNER) != ownerHwnd)
        {
            continue;
        }

        std::wstring popupFirstItemText;
        if (RedSalamander::DxUi::DebugGetContextMenuPopupItemText(popupHwnd, 0u, popupFirstItemText) && popupFirstItemText == firstItemText)
        {
            return popupHwnd;
        }
    }

    return nullptr;
}

HWND WaitForOwnedContextMenuPopupWindowByFirstItemText(HWND ownerHwnd,
                                                       std::wstring_view firstItemText,
                                                       std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (HWND popupHwnd = FindOwnedContextMenuPopupWindowByFirstItemText(ownerHwnd, firstItemText))
        {
            return popupHwnd;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return nullptr;
}

bool WaitForWindowDestroyed(HWND hwnd, std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (IsWindow(hwnd) == FALSE)
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return IsWindow(hwnd) == FALSE;
}

bool WaitForFocusedWindow(HWND hwnd, std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (GetFocus() == hwnd)
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return GetFocus() == hwnd;
}

void TestEmbeddedViewerContextMenuNativeConversionFiltersStandaloneCommands()
{
    using namespace RedSalamander::DxUi;

    wil::unique_hmenu menu{CreatePopupMenu()};
    Require(menu != nullptr, "embedded viewer context-menu conversion creates a native popup menu");

    wil::unique_hmenu standaloneSubmenu{CreatePopupMenu()};
    Require(standaloneSubmenu != nullptr, "embedded viewer context-menu conversion creates a standalone-only submenu");
    Require(AppendMenuW(standaloneSubmenu.get(), MF_STRING, 91701u, L"&Standalone action\tAlt+S") != FALSE,
            "embedded viewer context-menu conversion populates a submenu that should become empty after filtering");

    Require(AppendMenuW(menu.get(), MF_SEPARATOR, 0, nullptr) != FALSE, "embedded viewer context-menu conversion can start with a separator");
    Require(AppendMenuW(menu.get(), MF_STRING, 91702u, L"&Open standalone\tCtrl+O") != FALSE,
            "embedded viewer context-menu conversion populates a standalone-only command");
    Require(AppendMenuW(menu.get(), MF_SEPARATOR, 0, nullptr) != FALSE, "embedded viewer context-menu conversion can trim interior separators");
    Require(AppendMenuW(menu.get(), MF_STRING, 91703u, L"&Copy path\tCtrl+C") != FALSE,
            "embedded viewer context-menu conversion populates an embedded-safe command");
    Require(AppendMenuW(menu.get(), MF_POPUP | MF_STRING, reinterpret_cast<UINT_PTR>(standaloneSubmenu.get()), L"&Standalone") != FALSE,
            "embedded viewer context-menu conversion attaches the standalone-only submenu");
    static_cast<void>(standaloneSubmenu.release());
    Require(AppendMenuW(menu.get(), MF_SEPARATOR, 0, nullptr) != FALSE, "embedded viewer context-menu conversion can trim trailing separators");
    Require(AppendMenuW(menu.get(), MF_STRING | MF_CHECKED, 91704u, L"Show &metadata\tF4") != FALSE,
            "embedded viewer context-menu conversion preserves checked embedded-safe commands");

    static constexpr std::array<int, 2> kPreviewContextMenuExcludedCommandIds{{91701, 91702}};
    const NativeMenuFlyoutOptions options{
        .includeAcceleratorText = false,
        .omitEmptySubmenus      = true,
        .trimSeparators         = true,
        .excludedCommandIds     = kPreviewContextMenuExcludedCommandIds,
    };

    const std::vector<MenuFlyoutItem> items = ConvertNativeHMenuToFlyoutItems(menu.get(), options);
    Require(items.size() == 3u, "embedded viewer context-menu conversion removes standalone-only commands, empty submenus, and outer separator runs");
    Require(items[0].text == L"Copy path", "embedded viewer context-menu conversion strips mnemonics from retained command text");
    Require(items[0].commandId == 91703, "embedded viewer context-menu conversion preserves retained command IDs");
    Require(items[0].acceleratorText.empty(), "embedded viewer context-menu conversion suppresses standalone keyboard shortcut text");
    Require(items[1].kind == MenuItemKind::Separator, "embedded viewer context-menu conversion preserves one intentional separator between retained groups");
    Require(items[2].text == L"Show metadata", "embedded viewer context-menu conversion keeps the second embedded-safe command");
    Require(items[2].kind == MenuItemKind::Toggle && items[2].checked, "embedded viewer context-menu conversion preserves checked state");
}

void RunMenuDismissalKeyScenario(UINT message, WPARAM virtualKey, const char* appearExpectation, const char* dismissExpectation, const char* focusExpectation)
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOW);
    ownerWindow.PumpMessages();

    wil::unique_hwnd focusedChild{CreateWindowExW(0,
                                                  L"BUTTON",
                                                  L"DismissalTarget",
                                                  WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                                                  12,
                                                  80,
                                                  140,
                                                  28,
                                                  ownerWindow.Hwnd(),
                                                  nullptr,
                                                  GetModuleHandleW(nullptr),
                                                  nullptr)};
    Require(focusedChild != nullptr, "dismissal validation creates a focusable child control");

    wil::unique_hmenu menu{CreateMenu()};
    Require(menu != nullptr, "dismissal validation creates a top-level native menu");

    wil::unique_hmenu filePopup{CreatePopupMenu()};
    Require(filePopup != nullptr, "dismissal validation creates a native popup menu");
    Require(AppendMenuW(filePopup.get(), MF_STRING, 3701u, L"&Open") != FALSE, "dismissal validation populates the native popup menu");
    Require(AppendMenuW(filePopup.get(), MF_STRING, 3702u, L"&Close") != FALSE, "dismissal validation populates a second native popup item");
    Require(AppendMenuW(menu.get(), MF_POPUP, reinterpret_cast<UINT_PTR>(filePopup.get()), L"&File") != FALSE,
            "dismissal validation attaches the popup menu to the native menu bar");
    static_cast<void>(filePopup.release());

    NativeMenuBarHost menuBarHost;
    Require(menuBarHost.Attach(GetModuleHandleW(nullptr), ownerWindow.Hwnd(), menu.get()), "dismissal validation attaches a native menu bar host");
    ownerWindow.PumpMessages();

    static_cast<void>(SetActiveWindow(ownerWindow.Hwnd()));
    static_cast<void>(SetFocus(focusedChild.get()));
    Require(WaitForFocusedWindow(focusedChild.get()), "dismissal validation starts with focus on the child control");

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = appearExpectation;
            return;
        }

        PostMessageW(popupHwnd, message, virtualKey, 0);
        if (! WaitForWindowDestroyed(popupHwnd))
        {
            driverFailure = dismissExpectation;
        }
    });

    Require(menuBarHost.FocusFirstItem(), "dismissal validation enters native menu-bar mode");
    static_cast<void>(SendMessageW(menuBarHost.GetHwnd(), WM_KEYDOWN, VK_DOWN, 0));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(WaitForFocusedWindow(focusedChild.get()), focusExpectation);
}

std::shared_ptr<RedSalamander::DxUi::MenuFlyoutItem::BitmapIcon> CreateSyntheticMenuBitmapIcon(UINT sizePx)
{
    if (sizePx == 0u)
    {
        return {};
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = static_cast<LONG>(sizePx);
    bmi.bmiHeader.biHeight      = -static_cast<LONG>(sizePx);
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    wil::unique_hdc_window screenDc{GetDC(nullptr)};
    wil::unique_hbitmap bitmap{CreateDIBSection(screenDc.get(), &bmi, DIB_RGB_COLORS, &bits, nullptr, 0)};
    if (! bitmap || ! bits)
    {
        return {};
    }

    const auto pixelCount = static_cast<size_t>(sizePx) * static_cast<size_t>(sizePx);
    auto* const pixels    = static_cast<uint32_t*>(bits);
    std::fill_n(pixels, pixelCount, 0xFF1F7AE0u);

    const UINT inset = sizePx > 6u ? 2u : 1u;
    for (UINT y = inset; y < (sizePx - inset); ++y)
    {
        for (UINT x = inset; x < (sizePx - inset); ++x)
        {
            pixels[static_cast<size_t>(y) * sizePx + x] = 0xFFFFFFFFu;
        }
    }

    return std::make_shared<RedSalamander::DxUi::MenuFlyoutItem::BitmapIcon>(std::move(bitmap), sizePx, sizePx);
}

RedSalamander::DxUi::WindowHostBitmapCapture CaptureMenuPopupBitmapForTheme(const RedSalamander::DxUi::ThemePalette& theme,
                                                                            const std::vector<RedSalamander::DxUi::MenuFlyoutItem>& items)
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    WindowHostBitmapCapture capture{};
    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for material capture validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
        {
            driverFailure = "menu popup bitmap capture succeeds for material validation";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the material-capture menu with Escape returns no invoked command");
    return capture;
}

[[nodiscard]] RECT GetPrimaryMonitorWorkArea() noexcept
{
    MONITORINFO monitorInfo{};
    monitorInfo.cbSize     = sizeof(monitorInfo);
    const HMONITOR monitor = MonitorFromPoint(POINT{0, 0}, MONITOR_DEFAULTTOPRIMARY);
    if (monitor && GetMonitorInfoW(monitor, &monitorInfo) != FALSE)
    {
        return monitorInfo.rcWork;
    }

    return RECT{0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN)};
}

void TestMenuKeyboardNavigationSkipsInfoRows()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Info, .text = L"Used Space:", .acceleratorText = L"561 GB"},
        {.kind = MenuItemKind::Standard, .text = L"Disk Properties...", .commandId = 3001},
        {.kind = MenuItemKind::Info, .text = L"Free Space:", .acceleratorText = L"1.27 TB"},
        {.kind = MenuItemKind::Standard, .text = L"Disk Cleanup...", .commandId = 3002},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for info-row keyboard navigation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupDebugState popupState{};
        PostMessageW(popupHwnd, WM_KEYDOWN, VK_DOWN, 0);
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 1u;
        }, popupState))
        {
            driverFailure = "first VK_DOWN skips the leading info row and focuses the first actionable item";
            return;
        }

        PostMessageW(popupHwnd, WM_KEYDOWN, VK_DOWN, 0);
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 3u;
        }, popupState))
        {
            driverFailure = "second VK_DOWN skips the middle info row and focuses the next actionable item";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the info-row navigation menu with Escape returns no invoked command");
}

void TestMenuKeyboardRightArrowMatchesWindowsMenuLoop()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<std::vector<MenuFlyoutItem>> rootMenus = {
        {
            {.text = L"A1", .commandId = 3101},
            {.text = L"A2", .commandId = 3102},
        },
        {
            {.text      = L"B1",
             .commandId = 3201,
             .children =
                 {
                     MenuFlyoutItem{.text = L"B11", .commandId = 3211},
                     MenuFlyoutItem{.text = L"B12", .commandId = 3212},
                 }},
            {.text = L"B2", .commandId = 3202},
        },
        {
            {.text = L"C1", .commandId = 3301},
            {.text = L"C2", .commandId = 3302},
            {.text = L"C3", .commandId = 3303},
        },
    };

    size_t activeRootIndex = 0u;
    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromDirection = [&](bool forward) -> std::optional<ContextMenuRootSwitchRequest>
    {
        activeRootIndex = forward ? ((activeRootIndex + 1u) % rootMenus.size()) : ((activeRootIndex + rootMenus.size() - 1u) % rootMenus.size());

        ContextMenuRootSwitchRequest request{};
        request.screenPoint = POINT{180, 180};
        request.items       = rootMenus[activeRootIndex];
        return request;
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND aPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"A1");
        if (! aPopupHwnd)
        {
            driverFailure = "menu popup window appears for right-arrow root switching validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        ContextMenuPopupDebugState popupState{};
        PostMessageW(aPopupHwnd, WM_KEYDOWN, VK_HOME, 0);
        if (! WaitForContextMenuPopupState(aPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Home selects the first A root item before directional switching";
            return;
        }

        PostMessageW(aPopupHwnd, WM_KEYDOWN, VK_DOWN, 0);
        if (! WaitForContextMenuPopupState(aPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 1u;
        }, popupState))
        {
            driverFailure = "Down selects A2 before directional switching";
            return;
        }

        PostMessageW(aPopupHwnd, WM_KEYDOWN, VK_RIGHT, 0);
        const HWND bPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! bPopupHwnd)
        {
            driverFailure = "Right arrow on A2 opens the B root popup";
            return;
        }

        if (! WaitForContextMenuPopupState(bPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Right arrow on A2 focuses B1 in the next root popup";
            return;
        }

        PostMessageW(bPopupHwnd, WM_KEYDOWN, VK_RIGHT, 0);
        const HWND bSubmenuHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B11");
        if (! bSubmenuHwnd)
        {
            driverFailure = "Right arrow on B1 opens the B submenu";
            return;
        }

        if (! WaitForContextMenuPopupState(bSubmenuHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Right arrow on B1 focuses B11 in the submenu";
            return;
        }

        PostMessageW(bSubmenuHwnd, WM_KEYDOWN, VK_RIGHT, 0);
        const HWND cPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"C1");
        if (! cPopupHwnd)
        {
            driverFailure = "Right arrow on B11 switches from the leaf submenu to the C root popup";
            return;
        }

        if (! WaitForContextMenuPopupState(cPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Right arrow on B11 focuses C1 after switching to the next root popup";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the right-arrow menu loop validation popup with Escape returns no invoked command");
}

void TestStationaryMouseDoesNotOverrideKeyboardRootSwitch()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 420, 260, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<std::vector<MenuFlyoutItem>> rootMenus = {
        {
            {.text = L"A1", .commandId = 3651},
            {.text = L"A2", .commandId = 3652},
        },
        {
            {.text = L"B1", .commandId = 3661},
            {.text = L"B2", .commandId = 3662},
        },
        {
            {.text = L"C1", .commandId = 3671},
            {.text = L"C2", .commandId = 3672},
        },
    };
    const std::array<POINT, 3> rootPopupPoints = {POINT{170, 180}, POINT{250, 180}, POINT{330, 180}};

    size_t activeRootIndex  = 1u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromDirection = [&](bool forward) -> std::optional<ContextMenuRootSwitchRequest>
    {
        activeRootIndex = forward ? ((activeRootIndex + 1u) % rootMenus.size()) : ((activeRootIndex + rootMenus.size() - 1u) % rootMenus.size());
        return buildRequest(activeRootIndex);
    };
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE)
        {
            return std::nullopt;
        }
        if (clientPoint.y < 0 || clientPoint.y >= 40)
        {
            return std::nullopt;
        }

        size_t hitIndex = rootMenus.size();
        if (clientPoint.x >= 40 && clientPoint.x < 120)
        {
            hitIndex = 0u;
        }
        else if (clientPoint.x >= 140 && clientPoint.x < 220)
        {
            hitIndex = 1u;
        }
        else if (clientPoint.x >= 240 && clientPoint.x < 320)
        {
            hitIndex = 2u;
        }

        if (hitIndex >= rootMenus.size() || hitIndex == activeRootIndex)
        {
            return std::nullopt;
        }

        activeRootIndex = hitIndex;
        return buildRequest(activeRootIndex);
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND bPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! bPopupHwnd)
        {
            driverFailure = "menu popup window appears for stationary-mouse keyboard switching validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        ContextMenuPopupDebugState popupState{};
        PostMessageW(bPopupHwnd, WM_KEYDOWN, VK_HOME, 0);
        if (! WaitForContextMenuPopupState(bPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Home selects B1 before stationary-mouse keyboard switching validation";
            return;
        }

        PostMessageW(ownerWindow.Hwnd(), WM_MOUSEMOVE, 0, MAKELPARAM(160, 16));

        PostMessageW(bPopupHwnd, WM_KEYDOWN, VK_RIGHT, 0);
        const HWND cPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"C1");
        if (! cPopupHwnd)
        {
            driverFailure = "Right arrow switches from B to C before stationary-mouse validation";
            return;
        }

        if (! WaitForContextMenuPopupState(cPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "switching from B to C focuses C1 before stationary-mouse validation";
            return;
        }

        PostMessageW(ownerWindow.Hwnd(), WM_MOUSEMOVE, 0, MAKELPARAM(160, 16));
        if (WaitForWindowDestroyed(cPopupHwnd, std::chrono::milliseconds(120)))
        {
            driverFailure = "a stationary mouse over B does not pull keyboard navigation back from C";
            return;
        }

        PostMessageW(cPopupHwnd, WM_KEYDOWN, VK_LEFT, 0);
        const HWND bPopupAfterLeftHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! bPopupAfterLeftHwnd)
        {
            driverFailure = "Left arrow returns from C to B before stationary-mouse previous-root validation";
            return;
        }

        PostMessageW(bPopupAfterLeftHwnd, WM_KEYDOWN, VK_LEFT, 0);
        const HWND aPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"A1");
        if (! aPopupHwnd)
        {
            driverFailure = "a second Left arrow switches from B to A before stationary-mouse previous-root validation";
            return;
        }

        if (! WaitForContextMenuPopupState(aPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "switching from B to A focuses A1 before stationary-mouse previous-root validation";
            return;
        }

        PostMessageW(ownerWindow.Hwnd(), WM_MOUSEMOVE, 0, MAKELPARAM(160, 16));
        if (WaitForWindowDestroyed(aPopupHwnd, std::chrono::milliseconds(120)))
        {
            driverFailure = "a stationary mouse over B does not pull keyboard navigation back from A";
        }
    });

    const ThemePalette theme = MakeDefaultThemePalette(true);
    const std::optional<int> result =
        ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the stationary-mouse keyboard switching validation popup with Escape returns no invoked command");
}

void TestMenuPointerOverSiblingRootSwitchesOpenMenu()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 420, 260, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<std::vector<MenuFlyoutItem>> rootMenus = {
        {
            {.text = L"View one", .commandId = 36701},
            {.text = L"View two", .commandId = 36702},
        },
        {
            {.text = L"Plugins one", .commandId = 36711},
            {.text = L"Plugins two", .commandId = 36712},
        },
    };
    const std::array<POINT, 2> rootPopupPoints = {POINT{270, 180}, POINT{190, 180}};

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE)
        {
            return std::nullopt;
        }
        if (clientPoint.y < 0 || clientPoint.y >= 40)
        {
            return std::nullopt;
        }

        size_t hitIndex = rootMenus.size();
        if (clientPoint.x >= 140 && clientPoint.x < 220)
        {
            hitIndex = 1u;
        }
        else if (clientPoint.x >= 240 && clientPoint.x < 320)
        {
            hitIndex = 0u;
        }

        if (hitIndex >= rootMenus.size() || hitIndex == activeRootIndex)
        {
            return std::nullopt;
        }

        activeRootIndex = hitIndex;
        return buildRequest(activeRootIndex);
    };

    POINT viewMenuClientPoint{280, 20};
    POINT viewMenuScreenPoint = viewMenuClientPoint;
    Require(ClientToScreen(ownerWindow.Hwnd(), &viewMenuScreenPoint) != FALSE, "View root test point converts to screen coordinates for idle cursor polling validation");
    SetCursorPos(viewMenuScreenPoint.x, viewMenuScreenPoint.y);

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before pointer root-switch validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        POINT pluginsMenuClientPoint{180, 20};
        POINT pluginsMenuScreenPoint = pluginsMenuClientPoint;
        if (ClientToScreen(ownerWindow.Hwnd(), &pluginsMenuScreenPoint) == FALSE)
        {
            driverFailure = "Plugins root test point converts to screen coordinates";
            return;
        }

        RECT viewPopupRect{};
        if (GetWindowRect(viewPopupHwnd, &viewPopupRect) == FALSE)
        {
            driverFailure = "View root popup exposes a screen rect before pointer root-switch validation";
            return;
        }

        SetCursorPos(pluginsMenuScreenPoint.x, pluginsMenuScreenPoint.y);
        const LPARAM capturedMouseMovePoint = MAKELPARAM(pluginsMenuScreenPoint.x - viewPopupRect.left, pluginsMenuScreenPoint.y - viewPopupRect.top);
        PostMessageW(viewPopupHwnd, WM_MOUSEMOVE, 0, capturedMouseMovePoint);

        const HWND pluginsPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one");
        if (! pluginsPopupHwnd)
        {
            driverFailure = "moving the pointer from an open View root to the Plugins root opens the Plugins popup";
            return;
        }

        ContextMenuPopupDebugState pluginsPopupState{};
        if (! DebugGetContextMenuPopupState(pluginsPopupHwnd, pluginsPopupState) || pluginsPopupState.rootSwitchImmediateRenderCount == 0u ||
            pluginsPopupState.renderCount == 0u)
        {
            driverFailure = "root-switched Plugins popup is painted before the menu loop accepts another pointer move";
            return;
        }

        if (IsWindow(viewPopupHwnd) != FALSE)
        {
            driverFailure = "switching to the Plugins root closes the previous View popup";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the pointer root-switch validation popup with Escape returns no invoked command");
}

void TestMenuBarHoverMessageSwitchesRootWhenCursorOutsidePopup()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 420, 260, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::array<std::vector<MenuFlyoutItem>, 2> rootMenus = {
        std::vector<MenuFlyoutItem>{{.text = L"View one", .commandId = 36741}, {.text = L"View two", .commandId = 36742}},
        std::vector<MenuFlyoutItem>{{.text = L"Plugins one", .commandId = 36751}, {.text = L"Plugins two", .commandId = 36752}},
    };
    const std::array<POINT, 2> rootPopupPoints = {
        ClientScreenPointForTest(ownerWindow.Hwnd(), 270, 60, "View synthetic hover root popup point converts to screen coordinates"),
        ClientScreenPointForTest(ownerWindow.Hwnd(), 190, 60, "Plugins synthetic hover root popup point converts to screen coordinates"),
    };

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    std::atomic<int> pendingMenuBarHoverRootSwitch{-1};
    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromMenuBarHover = [&]() -> std::optional<ContextMenuRootSwitchRequest>
    {
        const int hoverIndex = pendingMenuBarHoverRootSwitch.exchange(-1);
        if (hoverIndex < 0)
        {
            return std::nullopt;
        }

        activeRootIndex = static_cast<size_t>(hoverIndex);
        return buildRequest(activeRootIndex);
    };

    POINT idleScreenPoint = ClientScreenPointForTest(ownerWindow.Hwnd(), 20, 220, "synthetic hover idle point converts to screen coordinates");
    SetCursorPos(idleScreenPoint.x, idleScreenPoint.y);

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before direct synthetic menu-bar hover validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        pendingMenuBarHoverRootSwitch.store(1);
        if (PostMessageW(viewPopupHwnd, WndMsg::kDxUiContextMenuRootHoverChanged, 1u, 0) == 0)
        {
            driverFailure = "View popup receives the direct synthetic menu-bar hover switch message";
            return;
        }

        const HWND pluginsPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one");
        if (! pluginsPopupHwnd)
        {
            driverFailure = "direct synthetic menu-bar hover message switches from View to Plugins while cursor is outside the popup";
            return;
        }

        ContextMenuPopupDebugState pluginsPopupState{};
        if (! DebugGetContextMenuPopupState(pluginsPopupHwnd, pluginsPopupState) || pluginsPopupState.rootSwitchImmediateRenderCount == 0u ||
            pluginsPopupState.renderCount == 0u)
        {
            driverFailure = "synthetic menu-bar hover root switch paints the replacement Plugins popup immediately";
            return;
        }

        if (IsWindow(viewPopupHwnd) != FALSE)
        {
            driverFailure = "synthetic menu-bar hover root switch closes the previous View popup";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(activeRootIndex == 1u, "direct synthetic menu-bar hover validation records the accepted root index");
    Require(! result.has_value(), "closing the direct synthetic menu-bar hover popup with Escape returns no invoked command");
}

void TestMenuPointerInsideOverlappingPopupDoesNotSwitchRoot()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 420, 260, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::array<std::vector<MenuFlyoutItem>, 2> rootMenus = {
        std::vector<MenuFlyoutItem>{{.text = L"View one", .commandId = 36721}, {.text = L"View two", .commandId = 36722}},
        std::vector<MenuFlyoutItem>{{.text = L"Plugins one", .commandId = 36731}, {.text = L"Plugins two", .commandId = 36732}},
    };
    const std::array<POINT, 2> rootPopupPoints = {
        ClientScreenPointForTest(ownerWindow.Hwnd(), 90, 0, "View overlapping root popup point converts to screen coordinates"),
        ClientScreenPointForTest(ownerWindow.Hwnd(), 190, 0, "Plugins overlapping root popup point converts to screen coordinates"),
    };
    constexpr LONG kMainMenuStripHeightPx = 96;

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    std::atomic<int> pendingMenuBarHoverRootSwitch{-1};
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE || clientPoint.y < 0 || clientPoint.y >= kMainMenuStripHeightPx)
        {
            return std::nullopt;
        }
        if (activeRootIndex == 1u)
        {
            return std::nullopt;
        }

        activeRootIndex = 1u;
        return buildRequest(activeRootIndex);
    };
    sessionCallbacks.switchRootFromMenuBarHover = [&]() -> std::optional<ContextMenuRootSwitchRequest>
    {
        const int hoverIndex = pendingMenuBarHoverRootSwitch.exchange(-1);
        if (hoverIndex < 0)
        {
            return std::nullopt;
        }

        activeRootIndex = static_cast<size_t>(hoverIndex);
        return buildRequest(activeRootIndex);
    };

    POINT idleClientPoint{20, 140};
    POINT idleScreenPoint = idleClientPoint;
    Require(ClientToScreen(ownerWindow.Hwnd(), &idleScreenPoint) != FALSE, "idle cursor point converts to screen coordinates");
    SetCursorPos(idleScreenPoint.x, idleScreenPoint.y);

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before overlapping-popup validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(viewPopupHwnd, [](const ContextMenuPopupDebugState&) noexcept { return true; }, popupState))
        {
            driverFailure = "overlapping View popup exposes debug state";
            return;
        }

        D2D1_RECT_F firstItemRectDip{};
        if (! WaitForContextMenuPopupItemRect(viewPopupHwnd, 0u, firstItemRectDip))
        {
            driverFailure = "overlapping View popup exposes its first item rect";
            return;
        }

        const float itemMidXDip = (firstItemRectDip.left + firstItemRectDip.right) * 0.5f;
        const float itemMidYDip = (firstItemRectDip.top + firstItemRectDip.bottom) * 0.5f;
        const POINT popupItemScreenPoint{
            popupState.surfaceRectPx.left + static_cast<LONG>(DipToPixelForPopup(itemMidXDip, popupState.dpi)),
            popupState.surfaceRectPx.top + static_cast<LONG>(DipToPixelForPopup(itemMidYDip, popupState.dpi)),
        };

        POINT ownerClientPoint = popupItemScreenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &ownerClientPoint) == FALSE)
        {
            driverFailure = "overlapping popup sample point converts to owner client coordinates";
            return;
        }
        if (ownerClientPoint.y < 0 || ownerClientPoint.y >= kMainMenuStripHeightPx)
        {
            driverFailure = "overlapping popup sample point sits over the owner main-menu strip";
            return;
        }

        RECT popupWindowRect{};
        if (GetWindowRect(viewPopupHwnd, &popupWindowRect) == FALSE)
        {
            driverFailure = "overlapping View popup exposes a screen rect";
            return;
        }

        PostMessageW(viewPopupHwnd,
                     WM_MOUSEMOVE,
                     0,
                     MAKELPARAM(popupItemScreenPoint.x - popupWindowRect.left, popupItemScreenPoint.y - popupWindowRect.top));

        const HWND pluginsPopupHwnd =
            WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one", std::chrono::milliseconds(180));
        if (pluginsPopupHwnd)
        {
            driverFailure = "moving inside a frontmost popup must not switch the root menu underneath it";
            return;
        }

        if (SetCursorPos(popupItemScreenPoint.x, popupItemScreenPoint.y) == FALSE)
        {
            driverFailure = "overlapping popup sample point remains available as the live cursor position before synthetic hover validation";
            return;
        }
        pendingMenuBarHoverRootSwitch.store(1);
        if (PostMessageW(viewPopupHwnd, WndMsg::kDxUiContextMenuRootHoverChanged, 1u, 0) == 0)
        {
            driverFailure = "overlapping popup can receive the synthetic menu-bar hover switch message";
            return;
        }
        const HWND pluginsPopupFromHoverHwnd =
            WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one", std::chrono::milliseconds(180));
        if (pluginsPopupFromHoverHwnd)
        {
            driverFailure = "a menu-bar hover message must not switch root while the live cursor is inside a frontmost popup";
            return;
        }

        if (IsWindow(viewPopupHwnd) == FALSE)
        {
            driverFailure = "View popup remains alive after moving inside its overlapping surface";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the overlapping popup validation with Escape returns no invoked command");
}

void TestMenuRootSwitchIgnoresStaleMouseMoveAfterCursorSwitch()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::array<std::vector<MenuFlyoutItem>, 2> rootMenus = {
        std::vector<MenuFlyoutItem>{{.text = L"View one", .commandId = 36801}, {.text = L"View two", .commandId = 36802}},
        std::vector<MenuFlyoutItem>{{.text = L"Plugins one", .commandId = 36811}, {.text = L"Plugins two", .commandId = 36812}},
    };
    const std::array<POINT, 2> rootPopupPoints = {POINT{270, 180}, POINT{190, 180}};

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE || clientPoint.y < 0 || clientPoint.y >= 40)
        {
            return std::nullopt;
        }

        size_t hitIndex = rootMenus.size();
        if (clientPoint.x >= 140 && clientPoint.x < 220)
        {
            hitIndex = 1u;
        }
        else if (clientPoint.x >= 240 && clientPoint.x < 320)
        {
            hitIndex = 0u;
        }

        if (hitIndex >= rootMenus.size() || hitIndex == activeRootIndex)
        {
            return std::nullopt;
        }

        activeRootIndex = hitIndex;
        return buildRequest(activeRootIndex);
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before stale mouse-move validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        POINT pluginsMenuClientPoint{180, 20};
        POINT pluginsMenuScreenPoint = pluginsMenuClientPoint;
        if (ClientToScreen(ownerWindow.Hwnd(), &pluginsMenuScreenPoint) == FALSE)
        {
            driverFailure = "Plugins root test point converts to screen coordinates for stale mouse-move validation";
            return;
        }

        POINT viewMenuClientPoint{280, 20};
        POINT viewMenuScreenPoint = viewMenuClientPoint;
        if (ClientToScreen(ownerWindow.Hwnd(), &viewMenuScreenPoint) == FALSE)
        {
            driverFailure = "View root stale point converts to screen coordinates";
            return;
        }

        RECT viewPopupRect{};
        if (GetWindowRect(viewPopupHwnd, &viewPopupRect) == FALSE)
        {
            driverFailure = "View root popup exposes a screen rect before stale mouse-move validation";
            return;
        }

        SetCursorPos(pluginsMenuScreenPoint.x, pluginsMenuScreenPoint.y);
        const LPARAM pluginsMovePoint = MAKELPARAM(pluginsMenuScreenPoint.x - viewPopupRect.left, pluginsMenuScreenPoint.y - viewPopupRect.top);
        PostMessageW(viewPopupHwnd, WM_MOUSEMOVE, 0, pluginsMovePoint);

        const HWND pluginsPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one");
        if (! pluginsPopupHwnd)
        {
            driverFailure = "moving the live cursor to Plugins opens the Plugins popup before stale mouse-move validation";
            return;
        }

        POINT staleViewClientPoint = viewMenuScreenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &staleViewClientPoint) == FALSE)
        {
            driverFailure = "stale View screen point converts to owner client coordinates";
            return;
        }
        PostMessageW(ownerWindow.Hwnd(), WM_MOUSEMOVE, 0, MAKELPARAM(staleViewClientPoint.x, staleViewClientPoint.y));

        const HWND staleViewPopupHwnd =
            WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one", std::chrono::milliseconds(180));
        if (staleViewPopupHwnd && staleViewPopupHwnd != viewPopupHwnd)
        {
            driverFailure = "a stale View mouse-move message switched away from the live Plugins cursor";
            return;
        }

        if (IsWindow(pluginsPopupHwnd) == FALSE)
        {
            driverFailure = "Plugins popup remains alive after the stale View mouse-move message";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the stale mouse-move validation popup with Escape returns no invoked command");
}

void TestMenuRootSwitchDoesNotPollCursorWhileIdle()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::array<std::vector<MenuFlyoutItem>, 2> rootMenus = {
        std::vector<MenuFlyoutItem>{{.text = L"View one", .commandId = 36901}, {.text = L"View two", .commandId = 36902}},
        std::vector<MenuFlyoutItem>{{.text = L"Plugins one", .commandId = 36911}, {.text = L"Plugins two", .commandId = 36912}},
    };
    const std::array<POINT, 2> rootPopupPoints = {POINT{270, 180}, POINT{190, 180}};

    size_t activeRootIndex  = 0u;
    const auto buildRequest = [&](size_t rootIndex) -> ContextMenuRootSwitchRequest
    {
        ContextMenuRootSwitchRequest request{};
        request.screenPoint = rootPopupPoints[rootIndex];
        request.items       = rootMenus[rootIndex];
        return request;
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromPointer = [&](POINT screenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        POINT clientPoint = screenPoint;
        if (ScreenToClient(ownerWindow.Hwnd(), &clientPoint) == FALSE || clientPoint.y < 0 || clientPoint.y >= 40)
        {
            return std::nullopt;
        }

        size_t hitIndex = rootMenus.size();
        if (clientPoint.x >= 140 && clientPoint.x < 220)
        {
            hitIndex = 1u;
        }
        else if (clientPoint.x >= 240 && clientPoint.x < 320)
        {
            hitIndex = 0u;
        }

        if (hitIndex >= rootMenus.size() || hitIndex == activeRootIndex)
        {
            return std::nullopt;
        }

        activeRootIndex = hitIndex;
        return buildRequest(activeRootIndex);
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND viewPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"View one");
        if (! viewPopupHwnd)
        {
            driverFailure = "View root popup opens before idle cursor polling validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const HWND pluginsPopupHwnd =
            WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Plugins one", std::chrono::milliseconds(180));
        if (pluginsPopupHwnd)
        {
            driverFailure = "waiting without a mouse-move message must not switch root menus by idle polling";
            return;
        }

        if (IsWindow(viewPopupHwnd) == FALSE)
        {
            driverFailure = "View popup remains alive when no mouse-move message is delivered";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), rootPopupPoints[activeRootIndex], rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the idle cursor polling validation popup with Escape returns no invoked command");
}

void TestMenuHoveringSiblingClosesOpenSubmenuAfterDelay()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text      = L"B1",
         .commandId = 3681,
         .children =
             {
                 MenuFlyoutItem{.text = L"B11", .commandId = 36811},
                 MenuFlyoutItem{.text = L"B12", .commandId = 36812},
             }},
        {.text = L"B2", .commandId = 3682},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND rootPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! rootPopupHwnd)
        {
            driverFailure = "menu popup window appears for delayed submenu-close validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        D2D1_RECT_F firstItemRectDip  = D2D1::RectF();
        D2D1_RECT_F secondItemRectDip = D2D1::RectF();
        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(rootPopupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           popupState) ||
            ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 0u, firstItemRectDip) || ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 1u, secondItemRectDip))
        {
            driverFailure = "menu popup exposes geometry for delayed submenu-close validation";
            return;
        }

        const float scale = static_cast<float>(popupState.dpi) / 96.0f;
        const LONG b1X    = static_cast<LONG>(std::lround(((firstItemRectDip.left + firstItemRectDip.right) * 0.5f) * scale));
        const LONG b1Y    = static_cast<LONG>(std::lround(((firstItemRectDip.top + firstItemRectDip.bottom) * 0.5f) * scale));
        const LONG b2X    = static_cast<LONG>(std::lround(((secondItemRectDip.left + secondItemRectDip.right) * 0.5f) * scale));
        const LONG b2Y    = static_cast<LONG>(std::lround(((secondItemRectDip.top + secondItemRectDip.bottom) * 0.5f) * scale));

        PostMessageW(rootPopupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(b1X, b1Y));
        const HWND submenuHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B11", std::chrono::milliseconds(1200));
        if (! submenuHwnd)
        {
            driverFailure = "hovering B1 long enough opens its submenu before delayed close validation";
            return;
        }

        PostMessageW(rootPopupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(b2X, b2Y));
        if (! WaitForContextMenuPopupState(rootPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 1u;
        }, popupState))
        {
            driverFailure = "hovering B2 updates the root-popup hover target before delayed submenu close";
            return;
        }

        if (! WaitForWindowDestroyed(submenuHwnd, std::chrono::milliseconds(1200)))
        {
            driverFailure = "hovering B2 long enough closes the already-open submenu from B1";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the delayed submenu-close validation popup with Escape returns no invoked command");
}

void TestMenuHoveringSiblingWithChildrenReplacesOpenSubmenuAfterDelay()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text      = L"B1",
         .commandId = 3691,
         .children =
             {
                 MenuFlyoutItem{.text = L"B11", .commandId = 36911},
                 MenuFlyoutItem{.text = L"B12", .commandId = 36912},
             }},
        {.text      = L"B2",
         .commandId = 3692,
         .children =
             {
                 MenuFlyoutItem{.text = L"B21", .commandId = 36921},
                 MenuFlyoutItem{.text = L"B22", .commandId = 36922},
             }},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND rootPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! rootPopupHwnd)
        {
            driverFailure = "menu popup window appears for delayed submenu-replacement validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        D2D1_RECT_F firstItemRectDip  = D2D1::RectF();
        D2D1_RECT_F secondItemRectDip = D2D1::RectF();
        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(rootPopupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           popupState) ||
            ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 0u, firstItemRectDip) || ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 1u, secondItemRectDip))
        {
            driverFailure = "menu popup exposes geometry for delayed submenu-replacement validation";
            return;
        }

        const float scale = static_cast<float>(popupState.dpi) / 96.0f;
        const LONG b1X    = static_cast<LONG>(std::lround(((firstItemRectDip.left + firstItemRectDip.right) * 0.5f) * scale));
        const LONG b1Y    = static_cast<LONG>(std::lround(((firstItemRectDip.top + firstItemRectDip.bottom) * 0.5f) * scale));
        const LONG b2X    = static_cast<LONG>(std::lround(((secondItemRectDip.left + secondItemRectDip.right) * 0.5f) * scale));
        const LONG b2Y    = static_cast<LONG>(std::lround(((secondItemRectDip.top + secondItemRectDip.bottom) * 0.5f) * scale));

        PostMessageW(rootPopupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(b1X, b1Y));
        const HWND firstSubmenuHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B11", std::chrono::milliseconds(1200));
        if (! firstSubmenuHwnd)
        {
            driverFailure = "hovering B1 long enough opens its submenu before delayed replacement validation";
            return;
        }

        PostMessageW(rootPopupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(b2X, b2Y));
        if (! WaitForContextMenuPopupState(rootPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 1u;
        }, popupState))
        {
            driverFailure = "hovering B2 updates the root-popup hover target before delayed submenu replacement";
            return;
        }

        const HWND replacementSubmenuHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B21", std::chrono::milliseconds(1200));
        if (! replacementSubmenuHwnd)
        {
            driverFailure = "hovering B2 long enough opens B2's replacement submenu";
            return;
        }

        if (! WaitForWindowDestroyed(firstSubmenuHwnd, std::chrono::milliseconds(1200)))
        {
            driverFailure = "opening B2's submenu closes the older submenu from B1";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the delayed submenu-replacement validation popup with Escape returns no invoked command");
}

void TestMenuPointerInsideSubmenuAndParentItemCancelPendingCloseDelay()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text      = L"B1",
         .commandId = 3701,
         .children =
             {
                 MenuFlyoutItem{.text = L"B11", .commandId = 37011},
                 MenuFlyoutItem{.text = L"B12", .commandId = 37012},
             }},
        {.text = L"B2", .commandId = 3702},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND rootPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! rootPopupHwnd)
        {
            driverFailure = "menu popup window appears for submenu hover-retention validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        D2D1_RECT_F firstItemRectDip  = D2D1::RectF();
        D2D1_RECT_F secondItemRectDip = D2D1::RectF();
        ContextMenuPopupDebugState rootState{};
        if (! WaitForContextMenuPopupState(rootPopupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           rootState) ||
            ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 0u, firstItemRectDip) || ! WaitForContextMenuPopupItemRect(rootPopupHwnd, 1u, secondItemRectDip))
        {
            driverFailure = "root popup exposes geometry for submenu hover-retention validation";
            return;
        }

        const float rootScale = static_cast<float>(rootState.dpi) / 96.0f;
        const LONG b1X        = static_cast<LONG>(std::lround(((firstItemRectDip.left + firstItemRectDip.right) * 0.5f) * rootScale));
        const LONG b1Y        = static_cast<LONG>(std::lround(((firstItemRectDip.top + firstItemRectDip.bottom) * 0.5f) * rootScale));
        const LONG b2X        = static_cast<LONG>(std::lround(((secondItemRectDip.left + secondItemRectDip.right) * 0.5f) * rootScale));
        const LONG b2Y        = static_cast<LONG>(std::lround(((secondItemRectDip.top + secondItemRectDip.bottom) * 0.5f) * rootScale));

        PostMessageW(rootPopupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(b1X, b1Y));
        const HWND submenuHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B11", std::chrono::milliseconds(1200));
        if (! submenuHwnd)
        {
            driverFailure = "hovering B1 long enough opens its submenu before hover-retention validation";
            return;
        }

        PostMessageW(rootPopupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(b2X, b2Y));
        if (! WaitForContextMenuPopupState(rootPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 1u;
        }, rootState))
        {
            driverFailure = "hovering B2 starts the delayed close path before entering the submenu";
            return;
        }

        D2D1_RECT_F submenuItemRectDip = D2D1::RectF();
        ContextMenuPopupDebugState submenuState{};
        if (! WaitForContextMenuPopupState(submenuHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           submenuState) ||
            ! WaitForContextMenuPopupItemRect(submenuHwnd, 0u, submenuItemRectDip))
        {
            driverFailure = "submenu exposes geometry before hover-retention validation";
            return;
        }

        const float submenuScale = static_cast<float>(submenuState.dpi) / 96.0f;
        const LONG submenuX      = static_cast<LONG>(std::lround(((submenuItemRectDip.left + submenuItemRectDip.right) * 0.5f) * submenuScale));
        const LONG submenuY      = static_cast<LONG>(std::lround(((submenuItemRectDip.top + submenuItemRectDip.bottom) * 0.5f) * submenuScale));
        PostMessageW(submenuHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(submenuX, submenuY));

        if (WaitForWindowDestroyed(submenuHwnd, std::chrono::milliseconds(700)))
        {
            driverFailure = "moving into an open submenu cancels the parent delayed-close timer";
            return;
        }

        PostMessageW(rootPopupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(b1X, b1Y));
        if (WaitForWindowDestroyed(submenuHwnd, std::chrono::milliseconds(700)))
        {
            driverFailure = "moving from an open submenu back to its parent menu item keeps the submenu open";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the submenu hover-retention validation popup with Escape returns no invoked command");
}

void TestMenuKeyboardTabExitsMenuLoop()
{
    RunMenuDismissalKeyScenario(WM_KEYDOWN,
                                VK_TAB,
                                "menu popup window appears for Tab dismissal validation",
                                "Tab dismisses the active menu popup chain",
                                "dismissing the menu popup with Tab preserves or restores focus to the previously focused control");
}

void TestMenuKeyboardF10ExitsMenuLoop()
{
    RunMenuDismissalKeyScenario(WM_SYSKEYDOWN,
                                VK_F10,
                                "menu popup window appears for F10 dismissal validation",
                                "F10 dismisses the active menu popup chain",
                                "dismissing the menu popup with F10 preserves or restores focus to the previously focused control");
}

void TestMenuKeyboardAltExitsMenuLoop()
{
    RunMenuDismissalKeyScenario(WM_SYSKEYDOWN,
                                VK_MENU,
                                "menu popup window appears for Alt dismissal validation",
                                "Alt dismisses the active menu popup chain",
                                "dismissing the menu popup with Alt preserves or restores focus to the previously focused control");
}

void TestMenuKeyboardLeftArrowMatchesWindowsMenuLoop()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<std::vector<MenuFlyoutItem>> rootMenus = {
        {
            {.text = L"A1", .commandId = 3401},
            {.text = L"A2", .commandId = 3402},
        },
        {
            {.text      = L"B1",
             .commandId = 3501,
             .children =
                 {
                     MenuFlyoutItem{.text = L"B11", .commandId = 3511},
                     MenuFlyoutItem{.text = L"B12", .commandId = 3512},
                 }},
            {.text = L"B2", .commandId = 3502},
        },
        {
            {.text = L"C1", .commandId = 3601},
            {.text = L"C2", .commandId = 3602},
        },
    };

    size_t activeRootIndex = 1u;
    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.switchRootFromDirection = [&](bool forward) -> std::optional<ContextMenuRootSwitchRequest>
    {
        activeRootIndex = forward ? ((activeRootIndex + 1u) % rootMenus.size()) : ((activeRootIndex + rootMenus.size() - 1u) % rootMenus.size());

        ContextMenuRootSwitchRequest request{};
        request.screenPoint = POINT{180, 180};
        request.items       = rootMenus[activeRootIndex];
        return request;
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND bPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B1");
        if (! bPopupHwnd)
        {
            driverFailure = "menu popup window appears for left-arrow menu loop validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        ContextMenuPopupDebugState popupState{};
        PostMessageW(bPopupHwnd, WM_KEYDOWN, VK_HOME, 0);
        if (! WaitForContextMenuPopupState(bPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Home selects B1 before left-arrow submenu validation";
            return;
        }

        PostMessageW(bPopupHwnd, WM_KEYDOWN, VK_RIGHT, 0);
        const HWND bSubmenuHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"B11");
        if (! bSubmenuHwnd)
        {
            driverFailure = "Right arrow on B1 opens the B submenu before left-arrow validation";
            return;
        }

        if (! WaitForContextMenuPopupState(bSubmenuHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "Right arrow on B1 focuses B11 before left-arrow validation";
            return;
        }

        PostMessageW(bSubmenuHwnd, WM_KEYDOWN, VK_LEFT, 0);
        if (! WaitForWindowDestroyed(bSubmenuHwnd))
        {
            driverFailure = "Left arrow closes the active submenu before switching roots";
            return;
        }

        if (! WaitForContextMenuPopupState(bPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "closing the submenu restores keyboard focus to B1";
            return;
        }

        PostMessageW(bPopupHwnd, WM_KEYDOWN, VK_LEFT, 0);
        const HWND aPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"A1");
        if (! aPopupHwnd)
        {
            driverFailure = "a second Left arrow on B switches to the previous root popup";
            return;
        }

        if (! WaitForContextMenuPopupState(aPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.keyboardIndex.has_value() && state.keyboardIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "switching to the previous root popup focuses A1";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, rootMenus[activeRootIndex], theme, sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the left-arrow menu loop validation popup with Escape returns no invoked command");
}

void TestNativeMenuBarRestoresFocusAfterMenuDismiss()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 360, 240, SWP_NOZORDER);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOW);
    ownerWindow.PumpMessages();

    wil::unique_hwnd focusedChild{CreateWindowExW(
        0, L"BUTTON", L"Pane", WS_CHILD | WS_VISIBLE | WS_TABSTOP, 12, 80, 120, 28, ownerWindow.Hwnd(), nullptr, GetModuleHandleW(nullptr), nullptr)};
    Require(focusedChild != nullptr, "focus-restoration validation creates a focusable child control");

    wil::unique_hmenu menu{CreateMenu()};
    Require(menu != nullptr, "focus-restoration validation creates a top-level native menu");

    wil::unique_hmenu filePopup{CreatePopupMenu()};
    Require(filePopup != nullptr, "focus-restoration validation creates a native popup menu");
    Require(AppendMenuW(filePopup.get(), MF_STRING, 3801u, L"&Open") != FALSE, "focus-restoration validation populates the native popup menu");
    Require(AppendMenuW(menu.get(), MF_POPUP, reinterpret_cast<UINT_PTR>(filePopup.get()), L"&File") != FALSE,
            "focus-restoration validation attaches the popup menu to the native menu bar");
    static_cast<void>(filePopup.release());

    NativeMenuBarHost menuBarHost;
    Require(menuBarHost.Attach(GetModuleHandleW(nullptr), ownerWindow.Hwnd(), menu.get()), "native menu bar host attaches for focus-restoration validation");
    ownerWindow.PumpMessages();

    static_cast<void>(SetActiveWindow(ownerWindow.Hwnd()));
    static_cast<void>(SetFocus(focusedChild.get()));
    Require(WaitForFocusedWindow(focusedChild.get()), "focus-restoration validation starts with focus on the child control");

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "native menu bar opens a popup for focus-restoration validation";
            return;
        }

        PostMessageW(popupHwnd, WM_KEYDOWN, VK_TAB, 0);
        if (! WaitForWindowDestroyed(popupHwnd))
        {
            driverFailure = "Tab dismisses the native menu bar popup during focus-restoration validation";
        }
    });

    Require(menuBarHost.FocusFirstItem(), "native menu bar host enters menu mode for focus-restoration validation");
    static_cast<void>(SendMessageW(menuBarHost.GetHwnd(), WM_KEYDOWN, VK_DOWN, 0));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(WaitForFocusedWindow(focusedChild.get()), "dismissing the native menu bar popup restores focus to the previously focused child control");
}

void TestMenuInfoRowsDoNotDismissOnClick()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Info, .text = L"Volume Label:", .acceleratorText = L"Home"},
        {.kind = MenuItemKind::Standard, .text = L"Disk Properties...", .commandId = 3001},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for info-row click validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        D2D1_RECT_F infoRectDip = D2D1::RectF();
        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           popupState) ||
            ! WaitForContextMenuPopupItemRect(popupHwnd, 0u, infoRectDip))
        {
            driverFailure = "info-row popup exposes geometry and debug state";
            return;
        }

        const float scale       = static_cast<float>(popupState.dpi) / 96.0f;
        const LONG clientX      = static_cast<LONG>(std::lround(((infoRectDip.left + infoRectDip.right) * 0.5f) * scale));
        const LONG clientY      = static_cast<LONG>(std::lround(((infoRectDip.top + infoRectDip.bottom) * 0.5f) * scale));
        const LPARAM clickPoint = MAKELPARAM(clientX, clientY);

        PostMessageW(popupHwnd, WM_MOUSEMOVE, 0, clickPoint);
        PostMessageW(popupHwnd, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        PostMessageW(popupHwnd, WM_LBUTTONUP, 0, clickPoint);

        std::this_thread::sleep_for(std::chrono::milliseconds(80));
        if (IsWindow(popupHwnd) == FALSE)
        {
            driverFailure = "clicking an info row does not dismiss the popup";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "clicking an info row does not invoke a command");
}

void TestMenuPopupPositionClampsAcrossDpiMatrix()
{
    using namespace RedSalamander::DxUi;

    const RECT workArea         = GetPrimaryMonitorWorkArea();
    constexpr UINT kDpiMatrix[] = {96u, 120u, 144u, 168u, 192u};

    for (const UINT dpi : kDpiMatrix)
    {
        RECT popupRect{};
        Require(DebugComputeContextMenuPopupPosition(POINT{workArea.left - 80, workArea.top - 60},
                                                     240.0f,
                                                     180.0f,
                                                     dpi,
                                                     false,
                                                     nullptr,
                                                     nullptr,
                                                     ContextMenuRootHorizontalAlignment::Start,
                                                     popupRect),
                std::format("root popup position resolves for {} DPI near the top-left edge", dpi).c_str());
        Require(popupRect.left == workArea.left, std::format("root popup clamps left to the work area at {} DPI", dpi).c_str());
        Require(popupRect.top == workArea.top, std::format("root popup clamps top to the work area at {} DPI", dpi).c_str());
        Require(popupRect.right <= workArea.right && popupRect.bottom <= workArea.bottom,
                std::format("root popup stays inside the work area after top-left clamping at {} DPI", dpi).c_str());

        popupRect = RECT{};
        Require(DebugComputeContextMenuPopupPosition(POINT{workArea.right - 4, workArea.bottom - 4},
                                                     240.0f,
                                                     180.0f,
                                                     dpi,
                                                     false,
                                                     nullptr,
                                                     nullptr,
                                                     ContextMenuRootHorizontalAlignment::Start,
                                                     popupRect),
                std::format("root popup position resolves for {} DPI near the bottom-right edge", dpi).c_str());
        Require(popupRect.right == workArea.right, std::format("root popup clamps right to the work area at {} DPI", dpi).c_str());
        Require(popupRect.bottom == workArea.bottom, std::format("root popup clamps bottom to the work area at {} DPI", dpi).c_str());
        Require(popupRect.left >= workArea.left && popupRect.top >= workArea.top,
                std::format("root popup stays inside the work area after bottom-right clamping at {} DPI", dpi).c_str());
    }
}

void TestMenuPopupPositionSupportsRightAlignedRootAnchors()
{
    using namespace RedSalamander::DxUi;

    const RECT workArea         = GetPrimaryMonitorWorkArea();
    constexpr UINT kDpiMatrix[] = {96u, 120u, 144u, 168u, 192u};

    for (const UINT dpi : kDpiMatrix)
    {
        RECT popupRect{};
        const int anchorRightPx = workArea.right - 24;
        Require(DebugComputeContextMenuPopupPosition(
                    POINT{anchorRightPx, workArea.top + 48}, 220.0f, 160.0f, dpi, false, nullptr, nullptr, ContextMenuRootHorizontalAlignment::End, popupRect),
                std::format("right-aligned root popup position resolves at {} DPI", dpi).c_str());
        Require(popupRect.right == anchorRightPx, std::format("right-aligned root popup keeps its visible right edge on the anchor at {} DPI", dpi).c_str());
        Require(popupRect.left >= workArea.left && popupRect.top >= workArea.top && popupRect.bottom <= workArea.bottom,
                std::format("right-aligned root popup still stays inside the work area at {} DPI", dpi).c_str());

        const int anchorLeftPx = workArea.left + 1;
        Require(DebugComputeContextMenuPopupPosition(
                    POINT{anchorLeftPx, workArea.top + 92}, 220.0f, 160.0f, dpi, false, nullptr, nullptr, ContextMenuRootHorizontalAlignment::End, popupRect),
                std::format("right-aligned root popup resolves near the left edge at {} DPI", dpi).c_str());
        Require(popupRect.left == anchorLeftPx,
                std::format("right-aligned root popup falls back to a start-aligned placement when trailing alignment would cross the left edge at {} DPI "
                            "(left={}, right={}, anchor={})",
                            dpi,
                            popupRect.left,
                            popupRect.right,
                            anchorLeftPx)
                    .c_str());
        Require(popupRect.right <= workArea.right && popupRect.top >= workArea.top && popupRect.bottom <= workArea.bottom,
                std::format("right-aligned root popup fallback still stays inside the work area at {} DPI", dpi).c_str());
    }
}

void TestMenuPopupPositionSupportsAboveRightAlignedRootAnchors()
{
    using namespace RedSalamander::DxUi;

    const RECT workArea         = GetPrimaryMonitorWorkArea();
    constexpr UINT kDpiMatrix[] = {96u, 120u, 144u, 168u, 192u};

    for (const UINT dpi : kDpiMatrix)
    {
        RECT popupRect{};
        const int anchorRightPx = workArea.left + ((workArea.right - workArea.left) / 2);
        const int anchorTopPx   = workArea.bottom - 96;
        Require(DebugComputeContextMenuPopupPosition(POINT{anchorRightPx, anchorTopPx},
                                                     220.0f,
                                                     160.0f,
                                                     dpi,
                                                     false,
                                                     nullptr,
                                                     nullptr,
                                                     ContextMenuRootHorizontalAlignment::End,
                                                     ContextMenuRootVerticalPlacement::Above,
                                                     popupRect),
                std::format("above/right-aligned root popup position resolves at {} DPI", dpi).c_str());
        Require(popupRect.right == anchorRightPx,
                std::format("above/right-aligned root popup keeps its visible right edge on the anchor at {} DPI", dpi).c_str());
        Require(popupRect.bottom == anchorTopPx,
                std::format("above/right-aligned root popup keeps its visible bottom edge above the anchor at {} DPI", dpi).c_str());
        Require(popupRect.left >= workArea.left && popupRect.top >= workArea.top && popupRect.right <= workArea.right,
                std::format("above/right-aligned root popup stays inside the work area at {} DPI", dpi).c_str());
    }
}

void TestMenuInfoRowsUseMeasuredValueColumnWidth()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Info, .text = L"Total Space:", .acceleratorText = L"1.82 TB (1999377526784 bytes)"},
        {.kind = MenuItemKind::Info, .text = L"Free Space:", .acceleratorText = L"1.27 TB"},
        {.kind = MenuItemKind::Standard, .text = L"Disk Properties...", .commandId = 3001},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for measured info-row column validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupItemLayoutDebugState firstLayout{};
        ContextMenuPopupItemLayoutDebugState secondLayout{};
        if (! WaitForContextMenuPopupItemLayout(popupHwnd, 0u, firstLayout) || ! WaitForContextMenuPopupItemLayout(popupHwnd, 1u, secondLayout))
        {
            driverFailure = "info-row popup exposes measured text and value-column layout";
            return;
        }

        const float firstValueWidth  = firstLayout.acceleratorRectDip.right - firstLayout.acceleratorRectDip.left;
        const float secondValueWidth = secondLayout.acceleratorRectDip.right - secondLayout.acceleratorRectDip.left;
        if (firstValueWidth <= 120.5f || secondValueWidth <= 120.5f)
        {
            driverFailure = "info rows expand the measured value column beyond the old fixed-width accelerator slot";
            return;
        }

        if (std::fabs(firstLayout.acceleratorRectDip.left - secondLayout.acceleratorRectDip.left) > 0.5f ||
            std::fabs(firstLayout.acceleratorRectDip.right - secondLayout.acceleratorRectDip.right) > 0.5f)
        {
            driverFailure = "info rows share a stable aligned value column";
            return;
        }

        if (firstLayout.textRectDip.right > firstLayout.acceleratorRectDip.left || secondLayout.textRectDip.right > secondLayout.acceleratorRectDip.left)
        {
            driverFailure = "info-row labels stay separated from the right-aligned value column";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the measured info-row layout menu with Escape returns no invoked command");
}

void TestMenuStandardRowsDeriveAndAlignShortcutColumnFromTabbedText()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Standard, .text = L"&Open\tCtrl+O", .commandId = 3001},
        {.kind = MenuItemKind::Standard, .text = L"E&xit\t  Alt+F4  ", .commandId = 3002},
        {.kind = MenuItemKind::Standard, .text = L"Preferences...", .commandId = 3003},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for tabbed shortcut layout validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupItemLayoutDebugState firstLayout{};
        ContextMenuPopupItemLayoutDebugState secondLayout{};
        if (! WaitForContextMenuPopupItemLayout(popupHwnd, 0u, firstLayout) || ! WaitForContextMenuPopupItemLayout(popupHwnd, 1u, secondLayout))
        {
            driverFailure = "standard rows expose shared shortcut-column layout";
            return;
        }

        const float firstShortcutWidth  = firstLayout.acceleratorRectDip.right - firstLayout.acceleratorRectDip.left;
        const float secondShortcutWidth = secondLayout.acceleratorRectDip.right - secondLayout.acceleratorRectDip.left;
        if (firstShortcutWidth <= 12.0f || secondShortcutWidth <= 12.0f)
        {
            driverFailure = "tabbed standard rows allocate a visible right-aligned shortcut column";
            return;
        }

        if (std::fabs(firstLayout.acceleratorRectDip.left - secondLayout.acceleratorRectDip.left) > 0.5f ||
            std::fabs(firstLayout.acceleratorRectDip.right - secondLayout.acceleratorRectDip.right) > 0.5f)
        {
            driverFailure = "tabbed standard rows share a stable aligned shortcut column";
            return;
        }

        if (firstLayout.textRectDip.right > firstLayout.acceleratorRectDip.left || secondLayout.textRectDip.right > secondLayout.acceleratorRectDip.left)
        {
            driverFailure = "tabbed standard row labels stay separated from the derived shortcut column";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the tabbed standard-row layout menu with Escape returns no invoked command");
}

void TestMenuBitmapIconsReachPopupLayout()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    std::vector<MenuFlyoutItem> items;
    items.push_back(MenuFlyoutItem{.text = L"Downloads", .iconBitmap = CreateSyntheticMenuBitmapIcon(16u), .commandId = 3001});

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for bitmap-icon validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupItemLayoutDebugState layout{};
        if (! WaitForContextMenuPopupItemLayout(popupHwnd, 0u, layout))
        {
            driverFailure = "menu popup exposes layout for bitmap-icon item";
            return;
        }

        if (! layout.hasBitmapIcon)
        {
            driverFailure = "bitmap icon payload is preserved for popup rendering";
            return;
        }

        if (layout.iconRectDip.left < (layout.itemRectDip.left + 3.5f))
        {
            driverFailure = "bitmap icon slot begins inside the hovered row instead of before it";
            return;
        }

        if (layout.iconRectDip.right > (layout.textRectDip.left - 6.0f))
        {
            driverFailure = "bitmap icon slot leaves spacing before the menu text column";
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the bitmap-icon menu with Escape returns no invoked command");
}

void TestMenuPopupMaterialsProduceDistinctCaptures()
{
    using namespace RedSalamander::DxUi;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open", .commandId = 3001},
        {.text = L"Rename", .commandId = 3002},
        {.kind = MenuItemKind::Separator},
        {.text = L"Properties", .commandId = 3003},
    };

    ThemePalette micaTheme    = MakeDefaultThemePalette(true);
    micaTheme.overlayMaterial = OverlayMaterial::Mica;

    ThemePalette micaAltTheme    = micaTheme;
    micaAltTheme.overlayMaterial = OverlayMaterial::MicaAlt;

    ThemePalette acrylicTheme    = micaTheme;
    acrylicTheme.overlayMaterial = OverlayMaterial::Acrylic;

    const WindowHostBitmapCapture micaCapture    = CaptureMenuPopupBitmapForTheme(micaTheme, items);
    const WindowHostBitmapCapture micaAltCapture = CaptureMenuPopupBitmapForTheme(micaAltTheme, items);
    const WindowHostBitmapCapture acrylicCapture = CaptureMenuPopupBitmapForTheme(acrylicTheme, items);

    const BitmapComparisonStats micaVsMicaAlt = CompareWindowHostBitmapCapturesForTest(micaCapture, micaAltCapture, 4u);
    Require(micaVsMicaAlt.totalPixels > 0u, "menu material captures share the same geometry for comparison");
    Require(micaVsMicaAlt.DifferenceRatio() > 0.01, "Mica and Mica Alt popup materials produce visibly different menu captures");

    const BitmapComparisonStats micaVsAcrylic = CompareWindowHostBitmapCapturesForTest(micaCapture, acrylicCapture, 4u);
    Require(micaVsAcrylic.totalPixels > 0u, "menu acrylic capture shares the same geometry for comparison");
    Require(micaVsAcrylic.DifferenceRatio() > 0.01, "Mica and Acrylic popup materials produce visibly different menu captures");
}

void TestMenuRainbowHoverUsesSeededHighlightContrast()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.rainbowMode     = true;
    theme.overlayMaterial = OverlayMaterial::Solid;

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Radio, .text = L"Rainbow", .checked = true, .commandId = 8051},
        {.kind = MenuItemKind::Radio, .text = L"Forest Mist", .checked = false, .commandId = 8052},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    ContextMenuPopupDebugState popupState{};
    ContextMenuPopupItemPaintDebugState paintState{};
    D2D1_RECT_F firstRowRectDip = D2D1::RectF();
    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for rainbow hover highlight validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        if (! WaitForContextMenuPopupItemRect(popupHwnd, 0u, firstRowRectDip))
        {
            driverFailure = "menu popup exposes row bounds for rainbow hover highlight validation";
            return;
        }

        const LONG hoverX = static_cast<LONG>(std::lround((firstRowRectDip.left + firstRowRectDip.right) * 0.5f));
        const LONG hoverY = static_cast<LONG>(std::lround((firstRowRectDip.top + firstRowRectDip.bottom) * 0.5f));
        PostMessageW(popupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(hoverX, hoverY));

        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "moving the pointer onto the rainbow row sets popup hover";
            return;
        }

        if (! WaitForContextMenuPopupItemPaint(popupHwnd, 0u, paintState))
        {
            driverFailure = "menu popup exposes hovered paint state for rainbow highlight validation";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the rainbow hover highlight validation menu with Escape returns no invoked command");

    const D2D1_COLOR_F expectedFill       = RainbowMenuSelectionTint(L"Rainbow", theme.dark);
    const D2D1_COLOR_F expectedForeground = ChooseContrastingTextColor(expectedFill);
    Require(paintState.hovered, "rainbow hover highlight paint state reports the row as hovered");
    Require(paintState.usesHighlightFill, "rainbow hover highlight paint state reports the highlight backplate");
    Require(paintState.usesRainbowHighlight, "rainbow hover highlight paint state reports the seeded rainbow backplate");
    RequireColorNear(paintState.fillColor, expectedFill, "rainbow hover highlight uses the stable seeded menu rainbow fill");
    RequireColorNear(paintState.compositeFillColor, expectedFill, "opaque rainbow hover highlight keeps its composite fill identical to the seeded color");
    RequireColorNear(paintState.textColor, expectedForeground, "rainbow hover highlight text uses the contrasting foreground");
    RequireColorNear(paintState.checkColor, expectedForeground, "rainbow hover highlight radio indicator uses the contrasting foreground");
}

void TestMenuHoverContrastAppliesToGlyphsAcrossThemes()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme      = MakeDefaultThemePalette(false);
    theme.rainbowMode       = false;
    theme.highContrast      = false;
    theme.overlayMaterial   = OverlayMaterial::Acrylic;
    theme.overlayBackground = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    theme.headerHovered     = D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f);
    theme.accent            = D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f);
    theme.text              = D2D1::ColorF(0.98f, 0.98f, 0.98f, 1.0f);
    theme.subduedText       = theme.text;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open With\tCtrl+Enter", .iconGlyph = L"\uE8A7", .commandId = 8151, .children = {MenuFlyoutItem{.text = L"Viewer", .commandId = 8152}}},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    ContextMenuPopupDebugState popupState{};
    ContextMenuPopupItemPaintDebugState paintState{};
    D2D1_RECT_F firstRowRectDip = D2D1::RectF();
    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for hover foreground contrast validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        if (! WaitForContextMenuPopupItemRect(popupHwnd, 0u, firstRowRectDip))
        {
            driverFailure = "menu popup exposes row bounds for hover foreground contrast validation";
            return;
        }

        const LONG hoverX = static_cast<LONG>(std::lround((firstRowRectDip.left + firstRowRectDip.right) * 0.5f));
        const LONG hoverY = static_cast<LONG>(std::lround((firstRowRectDip.top + firstRowRectDip.bottom) * 0.5f));
        PostMessageW(popupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(hoverX, hoverY));

        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 0u;
        }, popupState))
        {
            driverFailure = "moving the pointer onto the custom-theme row sets popup hover";
            return;
        }

        if (! WaitForContextMenuPopupItemPaint(popupHwnd, 0u, paintState))
        {
            driverFailure = "menu popup exposes hovered paint state for hover foreground contrast validation";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the hover foreground contrast validation menu with Escape returns no invoked command");

    const D2D1_COLOR_F expectedForeground = ChooseContrastingTextColor(paintState.compositeFillColor);
    RequireColorDifferent(theme.text, expectedForeground, "custom non-rainbow hover validation requires a foreground different from the base text token");
    Require(! paintState.usesRainbowHighlight, "non-rainbow hover validation keeps the shared highlight path out of rainbow mode");
    RequireColorNear(paintState.textColor, expectedForeground, "non-rainbow menu hover text uses the contrasting foreground");
    RequireColorNear(paintState.acceleratorColor, expectedForeground, "non-rainbow menu hover accelerator uses the contrasting foreground");
    RequireColorNear(paintState.iconColor, expectedForeground, "non-rainbow menu hover icon glyph uses the contrasting foreground");
    RequireColorNear(paintState.chevronColor, expectedForeground, "non-rainbow menu hover submenu chevron uses the contrasting foreground");
}

void TestMenuRainbowCheckedItemUsesAccentIndicator()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.rainbowMode     = true;
    theme.overlayMaterial = OverlayMaterial::Solid;

    const std::vector<MenuFlyoutItem> items = {
        {.kind = MenuItemKind::Radio, .text = L"Rainbow", .checked = true, .commandId = 8101},
        {.kind = MenuItemKind::Radio, .text = L"High Contrast", .checked = false, .commandId = 8102},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    WindowHostBitmapCapture capture{};
    ContextMenuPopupItemLayoutDebugState checkedLayout{};
    ContextMenuPopupItemLayoutDebugState uncheckedLayout{};
    ContextMenuPopupDebugState popupState{};
    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for rainbow checked-row validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState&) noexcept { return true; }, popupState))
        {
            driverFailure = "menu popup exposes debug state for rainbow checked-row validation";
            return;
        }
        if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
        {
            driverFailure = "menu popup bitmap capture succeeds for rainbow checked-row validation";
            return;
        }
        if (! WaitForContextMenuPopupItemLayout(popupHwnd, 0u, checkedLayout))
        {
            driverFailure = "menu popup exposes checked-row layout for rainbow checked-row validation";
            return;
        }
        if (! WaitForContextMenuPopupItemLayout(popupHwnd, 1u, uncheckedLayout))
        {
            driverFailure = "menu popup exposes unchecked-row layout for rainbow checked-row validation";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the rainbow checked-row menu with Escape returns no invoked command");

    const float iconCenterXDip = (checkedLayout.iconRectDip.left + checkedLayout.iconRectDip.right) * 0.5f;
    const UINT sampleX         = (std::min)(capture.widthPx - 1u, DipToPixelForPopup(iconCenterXDip, popupState.dpi));
    const UINT checkedY =
        (std::min)(capture.heightPx - 1u, DipToPixelForPopup((checkedLayout.iconRectDip.top + checkedLayout.iconRectDip.bottom) * 0.5f, popupState.dpi));
    const UINT uncheckedY =
        (std::min)(capture.heightPx - 1u, DipToPixelForPopup((uncheckedLayout.iconRectDip.top + uncheckedLayout.iconRectDip.bottom) * 0.5f, popupState.dpi));

    Require(GetCapturePixelBgra(capture, sampleX, checkedY) != GetCapturePixelBgra(capture, sampleX, uncheckedY),
            "rainbow checked menu rows render an accent-colored check indicator distinct from unchecked rows");
}

void TestMenuCheckedRowsDoNotPaintSecondFullRowSelection()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.rainbowMode     = true;
    theme.overlayMaterial = OverlayMaterial::Solid;

    const std::vector<MenuFlyoutItem> uncheckedItems = {
        {.kind = MenuItemKind::Radio, .text = L"Rainbow", .checked = false, .commandId = 8201},
        {.kind = MenuItemKind::Radio, .text = L"Forest Mist", .checked = false, .commandId = 8202},
        {.kind = MenuItemKind::Radio, .text = L"Neon Tokyo", .checked = false, .commandId = 8203},
    };
    const std::vector<MenuFlyoutItem> checkedItems = {
        {.kind = MenuItemKind::Radio, .text = L"Rainbow", .checked = true, .commandId = 8201},
        {.kind = MenuItemKind::Radio, .text = L"Forest Mist", .checked = false, .commandId = 8202},
        {.kind = MenuItemKind::Radio, .text = L"Neon Tokyo", .checked = false, .commandId = 8203},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const auto captureTrailingPixel = [&](const std::vector<MenuFlyoutItem>& items, uint32_t& trailingPixel, std::string& failure) noexcept
    {
        WindowHostBitmapCapture capture{};
        D2D1_RECT_F firstRowRectDip = D2D1::RectF();
        D2D1_RECT_F hoverRowRectDip = D2D1::RectF();
        ContextMenuPopupDebugState popupState{};
        std::thread driver([&]
        {
            const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
            if (! popupHwnd)
            {
                failure = "menu popup window appears for checked-row selection validation";
                return;
            }

            const auto dismissPopup = wil::scope_exit([&]() noexcept
            {
                if (IsWindow(popupHwnd) != FALSE)
                {
                    PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
                }
            });

            if (! WaitForContextMenuPopupItemRect(popupHwnd, 0u, firstRowRectDip) || ! WaitForContextMenuPopupItemRect(popupHwnd, 1u, hoverRowRectDip))
            {
                failure = "menu popup exposes row bounds for checked-row selection validation";
                return;
            }

            const LONG hoverX = static_cast<LONG>(std::lround((hoverRowRectDip.left + hoverRowRectDip.right) * 0.5f));
            const LONG hoverY = static_cast<LONG>(std::lround((hoverRowRectDip.top + hoverRowRectDip.bottom) * 0.5f));
            PostMessageW(popupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(hoverX, hoverY));

            if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
                return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 1u;
            }, popupState))
            {
                failure = "moving the pointer to a plain row transfers popup hover there";
                return;
            }

            if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
            {
                failure = "menu popup bitmap capture succeeds for checked-row selection validation";
            }
        });

        const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
        driver.join();

        if (! failure.empty())
        {
            return;
        }

        if (result.has_value())
        {
            failure = "closing the checked-row selection validation menu with Escape returns no invoked command";
            return;
        }

        const UINT sampleX = (std::min)(capture.widthPx - 1u, DipToPixelForPopup(firstRowRectDip.right - 18.0f, popupState.dpi));
        const UINT sampleY = (std::min)(capture.heightPx - 1u, DipToPixelForPopup((firstRowRectDip.top + firstRowRectDip.bottom) * 0.5f, popupState.dpi));
        trailingPixel      = GetCapturePixelBgra(capture, sampleX, sampleY);
    };

    uint32_t uncheckedTrailingPixel = 0u;
    uint32_t checkedTrailingPixel   = 0u;
    std::string driverFailure;
    captureTrailingPixel(uncheckedItems, uncheckedTrailingPixel, driverFailure);
    Require(driverFailure.empty(), driverFailure.c_str());
    captureTrailingPixel(checkedItems, checkedTrailingPixel, driverFailure);
    Require(driverFailure.empty(), driverFailure.c_str());

    Require(uncheckedTrailingPixel == checkedTrailingPixel,
            "checked menu rows do not paint a second full-width selection backplate when another row is hovered");
}

void TestMenuCheckedRowsDoNotPaintLeadingCheckedBox()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.overlayMaterial = OverlayMaterial::Solid;

    const std::vector<MenuFlyoutItem> uncheckedItems = {
        {.kind = MenuItemKind::Toggle, .text = L"Show Hidden Files", .checked = false, .commandId = 8301},
        {.kind = MenuItemKind::Toggle, .text = L"Show System Files", .checked = false, .commandId = 8302},
    };
    const std::vector<MenuFlyoutItem> checkedItems = {
        {.kind = MenuItemKind::Toggle, .text = L"Show Hidden Files", .checked = true, .commandId = 8301},
        {.kind = MenuItemKind::Toggle, .text = L"Show System Files", .checked = false, .commandId = 8302},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 240, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const auto captureLeadingPixel = [&](const std::vector<MenuFlyoutItem>& items, uint32_t& leadingPixel, std::string& failure) noexcept
    {
        WindowHostBitmapCapture capture{};
        ContextMenuPopupItemLayoutDebugState firstLayout{};
        D2D1_RECT_F hoverRowRectDip = D2D1::RectF();
        ContextMenuPopupDebugState popupState{};
        std::thread driver([&]
        {
            const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
            if (! popupHwnd)
            {
                failure = "menu popup window appears for checked-indicator box validation";
                return;
            }

            const auto dismissPopup = wil::scope_exit([&]() noexcept
            {
                if (IsWindow(popupHwnd) != FALSE)
                {
                    PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
                }
            });

            if (! WaitForContextMenuPopupItemLayout(popupHwnd, 0u, firstLayout) || ! WaitForContextMenuPopupItemRect(popupHwnd, 1u, hoverRowRectDip))
            {
                failure = "menu popup exposes leading indicator geometry for checked-indicator box validation";
                return;
            }

            const LONG hoverX = static_cast<LONG>(std::lround((hoverRowRectDip.left + hoverRowRectDip.right) * 0.5f));
            const LONG hoverY = static_cast<LONG>(std::lround((hoverRowRectDip.top + hoverRowRectDip.bottom) * 0.5f));
            PostMessageW(popupHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(hoverX, hoverY));

            if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
                return state.hoveredIndex.has_value() && state.hoveredIndex.value() == 1u;
            }, popupState))
            {
                failure = "moving the pointer away from the checked row transfers popup hover to the plain row";
                return;
            }

            if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
            {
                failure = "menu popup bitmap capture succeeds for checked-indicator box validation";
            }
        });

        const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
        driver.join();

        if (! failure.empty())
        {
            return;
        }

        if (result.has_value())
        {
            failure = "closing the checked-indicator box validation menu with Escape returns no invoked command";
            return;
        }

        const UINT sampleX = (std::min)(capture.widthPx - 1u, DipToPixelForPopup(firstLayout.iconRectDip.left + 2.0f, popupState.dpi));
        const UINT sampleY =
            (std::min)(capture.heightPx - 1u, DipToPixelForPopup((firstLayout.iconRectDip.top + firstLayout.iconRectDip.bottom) * 0.5f, popupState.dpi));
        leadingPixel = GetCapturePixelBgra(capture, sampleX, sampleY);
    };

    uint32_t uncheckedLeadingPixel = 0u;
    uint32_t checkedLeadingPixel   = 0u;
    std::string driverFailure;
    captureLeadingPixel(uncheckedItems, uncheckedLeadingPixel, driverFailure);
    Require(driverFailure.empty(), driverFailure.c_str());
    captureLeadingPixel(checkedItems, checkedLeadingPixel, driverFailure);
    Require(driverFailure.empty(), driverFailure.c_str());

    Require(uncheckedLeadingPixel == checkedLeadingPixel, "checked menu rows do not paint a rounded leading box behind the check indicator");
}

void TestMenuPopupCompositionHostUsesTransparentShadowMargins()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Desktop", .commandId = 3001},
        {.text = L"Documents", .commandId = 3002},
        {.text = L"Downloads", .commandId = 3003},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for composition shadow-margin validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupDebugState popupState{};
        WindowHostBitmapCapture capture{};
        if (! WaitForContextMenuPopupState(popupHwnd,
                                           [](const ContextMenuPopupDebugState& state) noexcept
        { return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f; },
                                           popupState) ||
            ! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
        {
            driverFailure = "composition popup exposes geometry and bitmap capture";
            return;
        }

        const double scale         = static_cast<double>(popupState.dpi) / 96.0;
        const UINT visibleWidthPx  = static_cast<UINT>((std::max)(1ll, std::llround(static_cast<double>(popupState.visibleWidthDip) * scale)));
        const UINT visibleHeightPx = static_cast<UINT>((std::max)(1ll, std::llround(static_cast<double>(popupState.visibleHeightDip) * scale)));
        if (capture.widthPx <= visibleWidthPx || capture.heightPx <= visibleHeightPx)
        {
            driverFailure = "composition popup capture includes transparent window margins outside the visible menu surface";
            return;
        }

        const uint8_t topLeftAlpha       = GetCapturePixelAlpha(capture, 0u, 0u);
        const uint8_t topRightAlpha      = GetCapturePixelAlpha(capture, capture.widthPx - 1u, 0u);
        const uint8_t centerAlpha        = GetCapturePixelAlpha(capture, capture.widthPx / 2u, capture.heightPx / 2u);
        const UINT visibleLeftPx         = static_cast<UINT>((std::max)(1ll, std::llround(10.0 * scale)));
        const UINT visibleTopPx          = static_cast<UINT>((std::max)(1ll, std::llround(8.0 * scale)));
        const UINT visibleCornerProbeX   = (std::min)(capture.widthPx - 1u, visibleLeftPx + 2u);
        const UINT visibleCornerProbeY   = (std::min)(capture.heightPx - 1u, visibleTopPx + 2u);
        const uint8_t visibleCornerAlpha = GetCapturePixelAlpha(capture, visibleCornerProbeX, visibleCornerProbeY);
        if (topLeftAlpha > 8u || topRightAlpha > 8u)
        {
            driverFailure = "composition popup keeps outer window corners transparent";
            return;
        }
        if (centerAlpha < 32u)
        {
            driverFailure = "composition popup keeps the menu surface itself visibly opaque inside the transparent window";
            return;
        }
        if (visibleCornerAlpha < 64u)
        {
            driverFailure = "composition popup keeps a RoundSmall-style near-corner body pixel inside the visible menu surface instead of clipping too deeply";
            return;
        }

        RECT popupRect{};
        if (GetWindowRect(popupHwnd, &popupRect) == FALSE)
        {
            driverFailure = "composition popup exposes a valid window rectangle for region validation";
            return;
        }

        const int popupWidthPx  = popupRect.right - popupRect.left;
        const int popupHeightPx = popupRect.bottom - popupRect.top;
        if (popupWidthPx <= 0 || popupHeightPx <= 0)
        {
            driverFailure = "composition popup keeps a positive-sized window rectangle";
            return;
        }

        wil::unique_hrgn windowRegion(CreateRectRgn(0, 0, 0, 0));
        if (! windowRegion)
        {
            driverFailure = "composition popup test can allocate a region handle";
            return;
        }

        if (GetWindowRgn(popupHwnd, windowRegion.get()) == ERROR)
        {
            driverFailure = "composition popup exposes a non-rectangular host window region";
            return;
        }

        if (PtInRegion(windowRegion.get(), 0, 0) != FALSE || PtInRegion(windowRegion.get(), popupWidthPx - 1, 0) != FALSE)
        {
            driverFailure = "composition popup window region clips the top window corners";
            return;
        }

        if (PtInRegion(windowRegion.get(), 1, popupHeightPx / 2) == FALSE || PtInRegion(windowRegion.get(), popupWidthPx - 2, popupHeightPx / 2) == FALSE ||
            PtInRegion(windowRegion.get(), popupWidthPx / 2, 1) == FALSE)
        {
            driverFailure = "composition popup window region keeps the outer host margins available for the popup shadow";
            return;
        }

        const int topShadowShoulderX = (std::max)(1, static_cast<int>(std::llround(3.0 * scale)));
        const int topShadowShoulderY = (std::max)(1, static_cast<int>(std::llround(8.0 * scale)));
        if (PtInRegion(windowRegion.get(), topShadowShoulderX, topShadowShoulderY) == FALSE ||
            PtInRegion(windowRegion.get(), popupWidthPx - 1 - topShadowShoulderX, topShadowShoulderY) == FALSE)
        {
            driverFailure =
                "composition popup window region preserves the thinner top shadow shoulder instead of over-rounding it to match the thicker bottom margin";
            return;
        }

        if (PtInRegion(windowRegion.get(), popupWidthPx / 2, popupHeightPx / 2) == FALSE)
        {
            driverFailure = "composition popup window region still contains the menu body";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, MakeDefaultThemePalette(true));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the composition-shadow menu with Escape returns no invoked command");
}

void TestMenuPopupWindowClassDoesNotUseNativeDropShadow()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open", .commandId = 4101},
        {.text = L"Copy", .commandId = 4102},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for native-drop-shadow validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        const ULONG_PTR classStyle = GetClassLongPtrW(popupHwnd, GCL_STYLE);
        if ((classStyle & CS_DROPSHADOW) != 0u)
        {
            driverFailure = "composition popup window class does not opt into the native CS_DROPSHADOW effect";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, MakeDefaultThemePalette(true));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the native-drop-shadow validation menu with Escape returns no invoked command");
}

void TestMenuPopupAcrylicLightVisualBaseline()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme    = MakeDefaultThemePalette(false);
    theme.overlayMaterial = OverlayMaterial::Acrylic;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open", .commandId = 4201},
        {.text = L"Copy", .commandId = 4202},
        {.text = L"Move", .commandId = 4203},
        {.text = L"Properties", .commandId = 4204},
    };

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ownerWindow.Host().SetRoot(std::make_unique<StripedBackdropControl>(10));
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();
    WindowHostBitmapCapture ownerBackdropCapture{};
    Require(ownerWindow.Host().DebugCaptureBitmap(ownerBackdropCapture), "acrylic visual-baseline owner backdrop renders before popup capture");
    const POINT menuPoint = ClientScreenPointForTest(ownerWindow.Hwnd(), 72, 48, "acrylic baseline popup anchor maps to screen coordinates");

    WindowHostBitmapCapture capture{};
    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for acrylic visual-baseline capture";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.usesAppBackdropBlur && state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f;
        }, popupState))
        {
            driverFailure = "menu popup exposes acrylic geometry for visual-baseline capture";
            return;
        }

        POINT popupClientOrigin{menuPoint.x, menuPoint.y};
        if (ScreenToClient(ownerWindow.Hwnd(), &popupClientOrigin) == FALSE)
        {
            driverFailure = "acrylic visual-baseline popup origin maps to owner client coordinates";
            return;
        }

        const UINT surfaceWidthPx  = DipToPixelForPopup(popupState.visibleWidthDip, popupState.dpi);
        const UINT surfaceHeightPx = DipToPixelForPopup(popupState.visibleHeightDip, popupState.dpi);
        const RECT backdropCropRect{
            popupClientOrigin.x,
            popupClientOrigin.y,
            popupClientOrigin.x + static_cast<LONG>(surfaceWidthPx),
            popupClientOrigin.y + static_cast<LONG>(surfaceHeightPx),
        };
        const WindowHostBitmapCapture deterministicBackdrop = CropWindowHostBitmapCaptureForTest(ownerBackdropCapture, backdropCropRect);
        if (! DebugSetContextMenuPopupBackdropCapture(popupHwnd, deterministicBackdrop))
        {
            driverFailure = "menu popup accepts deterministic acrylic backdrop for visual-baseline capture";
            return;
        }

        if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, capture))
        {
            driverFailure = "menu popup bitmap capture succeeds for acrylic visual-baseline capture";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), menuPoint, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the acrylic visual-baseline menu with Escape returns no invoked command");
    VerifyOrUpdateBaselineForTest("acrylic light menu popup baseline matches", L"menu_popup_acrylic_light.png", capture, 0.06, 10u);
}

void TestMenuPopupKeepsSystemBackdropDisabledForAppRenderedMaterials()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.overlayMaterial = OverlayMaterial::Acrylic;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open", .commandId = 3001},
        {.text = L"Rename", .commandId = 3002},
        {.text = L"Properties", .commandId = 3003},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for popup-backdrop validation";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f;
        }, popupState))
        {
            driverFailure = "menu popup exposes debug state for popup-backdrop validation";
            return;
        }

        if (popupState.usesSystemBackdrop)
        {
            driverFailure = "app-rendered popup materials keep the popup HWND free of DWM system backdrops";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the popup-backdrop validation menu with Escape returns no invoked command");
}

void TestMenuPopupSubmenuFlipsLeftNearRightEdge()
{
    using namespace RedSalamander::DxUi;

    const RECT workArea   = GetPrimaryMonitorWorkArea();
    const RECT parentRect = {
        workArea.right - 220,
        workArea.top + 80,
        workArea.right - 24,
        workArea.top + 260,
    };
    const RECT parentItemRect = {
        workArea.right - 34,
        workArea.top + 112,
        workArea.right - 12,
        workArea.top + 146,
    };

    RECT popupRect{};
    Require(DebugComputeContextMenuPopupPosition(POINT{parentItemRect.right, parentItemRect.top},
                                                 200.0f,
                                                 160.0f,
                                                 144u,
                                                 true,
                                                 &parentRect,
                                                 &parentItemRect,
                                                 ContextMenuRootHorizontalAlignment::Start,
                                                 popupRect),
            "submenu popup position resolves near the right screen edge");
    Require(popupRect.right <= parentItemRect.left, "submenu popup flips to the left side when opening right would overflow the work area");
    Require(popupRect.left >= workArea.left && popupRect.top >= workArea.top && popupRect.bottom <= workArea.bottom,
            "flipped submenu popup remains fully inside the monitor work area");
}

void TestMenuMnemonicHonorsExplicitAmpersandLabels()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Save && Close\tCtrl+S", .commandId = 3001},
        {.text = L"E&xit\tAlt+F4", .commandId = 3002},
        {.text = L"Rock &", .commandId = 3003},
    };
    std::wstring normalizedFirstItemText;
    const bool normalizedFirstItem = DebugGetContextMenuItemDisplayText(items[0], normalizedFirstItemText);
    if (! normalizedFirstItem || normalizedFirstItemText != L"Save & Close")
    {
        const std::string failure =
            std::format("menu item parser preserves escaped ampersands while stripping tabbed accelerator text actual='{}' codeUnits='{}'",
                        NarrowAsciiForFailureMessage(normalizedFirstItemText),
                        WideCodeUnitsForFailureMessage(normalizedFirstItemText));
        Require(false, failure.c_str());
    }

    std::string driverFailure;
    std::thread driver([&]
    {
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Save & Close");
        if (! popupHwnd)
        {
            std::wstring popupFirstItemText;
            if (HWND fallbackPopupHwnd = FindOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
                fallbackPopupHwnd && RedSalamander::DxUi::DebugGetContextMenuPopupItemText(fallbackPopupHwnd, 0u, popupFirstItemText))
            {
                driverFailure = std::format("menu popup window appears for explicit ampersand mnemonic validation actual='{}' codeUnits='{}'",
                                            NarrowAsciiForFailureMessage(popupFirstItemText),
                                            WideCodeUnitsForFailureMessage(popupFirstItemText));
            }
            else
            {
                driverFailure = "menu popup window appears for explicit ampersand mnemonic validation";
            }
            return;
        }

        std::wstring firstItemText;
        std::wstring secondItemText;
        std::wstring thirdItemText;
        if (! WaitForContextMenuPopupItemText(popupHwnd, 0u, firstItemText) || ! WaitForContextMenuPopupItemText(popupHwnd, 1u, secondItemText) ||
            ! WaitForContextMenuPopupItemText(popupHwnd, 2u, thirdItemText))
        {
            driverFailure = "menu popup exposes item text for explicit ampersand mnemonic validation";
            return;
        }

        if (firstItemText != L"Save & Close")
        {
            driverFailure = std::format("menu popup preserves escaped ampersands while stripping tabbed accelerator text actual='{}' codeUnits='{}'",
                                        NarrowAsciiForFailureMessage(firstItemText),
                                        WideCodeUnitsForFailureMessage(firstItemText));
            return;
        }

        if (secondItemText != L"Exit")
        {
            driverFailure = std::format("menu popup strips explicit ampersand markers from displayed item text actual='{}' codeUnits='{}'",
                                        NarrowAsciiForFailureMessage(secondItemText),
                                        WideCodeUnitsForFailureMessage(secondItemText));
            return;
        }

        if (thirdItemText != L"Rock &")
        {
            driverFailure = std::format("menu popup preserves a trailing literal ampersand in displayed item text actual='{}' codeUnits='{}'",
                                        NarrowAsciiForFailureMessage(thirdItemText),
                                        WideCodeUnitsForFailureMessage(thirdItemText));
            return;
        }

        PostMessageW(popupHwnd, WM_KEYDOWN, 'X', 0);
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, MakeDefaultThemePalette(true));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(result.has_value() && result.value() == 3002, "explicit ampersand mnemonic invokes the matching popup item");
}

void TestMenuOpeningPointerUpCanBeIgnoredOutsideVisibleSurface()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Ignore Release", .commandId = 3001},
        {.text = L"Close", .commandId = 3002},
    };

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.ignoreInitialLeftButtonUp = true;

    std::string driverFailure;
    std::thread driver([&]
    {
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Ignore Release");
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for initial pointer-up validation";
            return;
        }

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f;
        }, popupState))
        {
            driverFailure = "menu popup exposes geometry for initial pointer-up validation";
            return;
        }

        PostMessageW(popupHwnd, WM_LBUTTONUP, 0, MAKELPARAM(1, 1));
        std::this_thread::sleep_for(std::chrono::milliseconds(80));

        if (IsWindow(popupHwnd) == FALSE)
        {
            driverFailure = "menu popup ignores the opening button-up before applying light-dismiss rules";
            return;
        }

        PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, MakeDefaultThemePalette(true), sessionCallbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the initial pointer-up validation menu returns no invoked command");
}

void TestMenuShadowMarginMouseUpLightDismissesAfterInitialRelease()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Shadow Desktop", .commandId = 3001},
        {.text = L"Documents", .commandId = 3002},
        {.text = L"Downloads", .commandId = 3003},
    };

    std::string driverFailure;
    std::thread driver([&]
    {
        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), L"Shadow Desktop");
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for shadow-margin light-dismiss validation";
            return;
        }

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
            return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f;
        }, popupState))
        {
            driverFailure = "menu popup exposes geometry for shadow-margin light-dismiss validation";
            return;
        }

        PostMessageW(popupHwnd, WM_LBUTTONUP, 0, MAKELPARAM(1, 1));
        if (! WaitForWindowDestroyed(popupHwnd))
        {
            driverFailure = "menu popup treats transparent shadow margins as outside the visible menu surface for light-dismiss";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, MakeDefaultThemePalette(true));
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "shadow-margin light-dismiss returns no invoked command");
}

void TestMenuAcrylicBackdropScenarioEmitsMetrics()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 96, 96, 480, 360, SWP_NOZORDER | SWP_NOACTIVATE);
    ownerWindow.Host().SetRoot(std::make_unique<StripedBackdropControl>(12));
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    UpdateWindow(ownerWindow.Hwnd());
    ownerWindow.PumpMessages();
    WindowHostBitmapCapture ownerBackdropBeforePopup{};
    Require(ownerWindow.Host().DebugCaptureBitmap(ownerBackdropBeforePopup), "owner backdrop renders before acrylic metric capture");
    const POINT menuPoint = ClientScreenPointForTest(ownerWindow.Hwnd(), 96, 72, "acrylic metric popup anchor maps to screen coordinates");

    ThemePalette theme    = MakeDefaultThemePalette(true);
    theme.overlayMaterial = OverlayMaterial::Acrylic;

    const std::vector<MenuFlyoutItem> items = {
        {.text = L"Open", .commandId = 3001},
        {.text = L"Copy", .commandId = 3002},
        {.text = L"Move", .commandId = 3003},
    };

    std::string driverFailure;
    WindowHostBitmapCapture popupCapture{};
    uint64_t openToCaptureUs = 0u;
    const auto startedAt     = std::chrono::steady_clock::now();

    std::thread driver([&]
    {
        const HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "menu popup window appears for acrylic backdrop metric capture";
            return;
        }

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) { return state.usesAppBackdropBlur; }, popupState))
        {
            driverFailure = "menu popup enables app-rendered backdrop blur for acrylic backdrop metric capture";
            return;
        }

        POINT popupClientOrigin{menuPoint.x, menuPoint.y};
        if (ScreenToClient(ownerWindow.Hwnd(), &popupClientOrigin) == FALSE)
        {
            driverFailure = "acrylic backdrop metric popup origin maps to owner client coordinates";
            return;
        }

        const UINT surfaceWidthPx  = DipToPixelForPopup(popupState.visibleWidthDip, popupState.dpi);
        const UINT surfaceHeightPx = DipToPixelForPopup(popupState.visibleHeightDip, popupState.dpi);
        const RECT backdropCropRect{
            popupClientOrigin.x,
            popupClientOrigin.y,
            popupClientOrigin.x + static_cast<LONG>(surfaceWidthPx),
            popupClientOrigin.y + static_cast<LONG>(surfaceHeightPx),
        };
        const WindowHostBitmapCapture deterministicBackdrop = CropWindowHostBitmapCaptureForTest(ownerBackdropBeforePopup, backdropCropRect);
        if (! DebugSetContextMenuPopupBackdropCapture(popupHwnd, deterministicBackdrop))
        {
            driverFailure = "menu popup accepts deterministic acrylic backdrop for metric capture";
            return;
        }

        if (! WaitForContextMenuPopupBitmapCapture(popupHwnd, popupCapture))
        {
            driverFailure = "menu popup bitmap capture succeeds for acrylic backdrop metric capture";
            return;
        }

        openToCaptureUs = Debug::Perf::ElapsedUs(startedAt);
        PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), menuPoint, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "closing the acrylic backdrop metric menu with Escape returns no invoked command");
    Require(popupCapture.widthPx > 0u && popupCapture.heightPx > 0u && ! popupCapture.bgraPixels.empty(),
            "acrylic backdrop metric scenario captures the popup bitmap");

    const RECT opaqueBounds = FindOpaqueBoundsInCapture(popupCapture);
    Require(opaqueBounds.right > opaqueBounds.left && opaqueBounds.bottom > opaqueBounds.top, "acrylic backdrop metric scenario resolves opaque popup bounds");
    const RECT sampleRect = ComputeRightStripSampleRect(opaqueBounds);
    Require(sampleRect.right > sampleRect.left && sampleRect.bottom > sampleRect.top,
            "acrylic backdrop metric scenario resolves a valid right-strip sample region");

    std::this_thread::sleep_for(std::chrono::milliseconds(30));
    InvalidateRect(ownerWindow.Hwnd(), nullptr, FALSE);
    UpdateWindow(ownerWindow.Hwnd());
    ownerWindow.PumpMessages();
    WindowHostBitmapCapture ownerBackdropCapture{};
    Require(ownerWindow.Host().DebugCaptureBitmap(ownerBackdropCapture), "owner backdrop capture succeeds for acrylic metrics");
    POINT popupClientOrigin{menuPoint.x, menuPoint.y};
    Require(ScreenToClient(ownerWindow.Hwnd(), &popupClientOrigin) != FALSE, "popup origin maps to owner client coordinates for acrylic metrics");
    const RECT rawCaptureRect = {
        popupClientOrigin.x - opaqueBounds.left,
        popupClientOrigin.y - opaqueBounds.top,
        popupClientOrigin.x - opaqueBounds.left + static_cast<LONG>(popupCapture.widthPx),
        popupClientOrigin.y - opaqueBounds.top + static_cast<LONG>(popupCapture.heightPx),
    };
    const WindowHostBitmapCapture rawCapture = CropWindowHostBitmapCaptureForTest(ownerBackdropCapture, rawCaptureRect);
    Require(rawCapture.widthPx == popupCapture.widthPx && rawCapture.heightPx == popupCapture.heightPx,
            "raw backdrop capture matches the popup capture geometry for acrylic metrics");

    const uint64_t popupAdjacentDelta = ComputeAverageAdjacentRgbDelta(popupCapture, sampleRect);
    const uint64_t rawAdjacentDelta   = ComputeAverageAdjacentRgbDelta(rawCapture, sampleRect);
    const uint64_t popupVsRawDelta    = ComputeAverageAbsoluteRgbDeltaBetweenCaptures(popupCapture, rawCapture, sampleRect);
    const uint64_t popupVsSlabDelta   = ComputeAverageAbsoluteRgbDeltaToColor(popupCapture, sampleRect, ResolveExpectedAcrylicMenuSlabColor(theme));
    const uint64_t popupToRawDeltaPermille =
        rawAdjacentDelta == 0u ? 0u : static_cast<uint64_t>((popupAdjacentDelta * 1000u + (rawAdjacentDelta / 2u)) / rawAdjacentDelta);
    const uint64_t minStrongBlurDelta = rawAdjacentDelta <= 1u ? 48u : 56u;

    Debug::Perf::EmitValue(L"dxui.menu.selftest.acrylic_popup_adjacent_rgb_delta", popupAdjacentDelta, S_OK);
    Debug::Perf::EmitValue(L"dxui.menu.selftest.acrylic_raw_adjacent_rgb_delta", rawAdjacentDelta, S_OK);
    Debug::Perf::EmitValue(L"dxui.menu.selftest.acrylic_popup_vs_raw_rgb_delta", popupVsRawDelta, S_OK);
    Debug::Perf::EmitValue(L"dxui.menu.selftest.acrylic_popup_vs_slab_rgb_delta", popupVsSlabDelta, S_OK);
    Debug::Perf::EmitValue(L"dxui.menu.selftest.acrylic_popup_to_raw_delta_permille", popupToRawDeltaPermille, S_OK);

    Require(rawAdjacentDelta > 0u, "acrylic backdrop metric scenario records non-zero raw backdrop variance");
    Require(popupVsRawDelta > 0u, "acrylic backdrop metric scenario materially changes the captured backdrop sample");
    Require(popupAdjacentDelta < rawAdjacentDelta, "acrylic backdrop metric scenario blurs the popup backdrop below the raw screen variance");
    Require(popupVsRawDelta >= minStrongBlurDelta,
            "acrylic backdrop metric scenario applies a visibly strong blur instead of a barely-changed transparent tint");
    Require(popupVsSlabDelta >= 24u,
            "acrylic backdrop metric scenario stays visually tied to the captured raw backdrop instead of collapsing into an opaque tint");

    Debug::Perf::Emit(L"dxui.menu.selftest.acrylic_open_to_capture_us",
                      L"",
                      openToCaptureUs,
                      static_cast<uint64_t>(popupCapture.widthPx),
                      static_cast<uint64_t>(popupCapture.heightPx),
                      S_OK);
}

} // namespace

void RunMenuTests()
{
    auto runTest = [](const char* name, void (*fn)())
    {
        std::cerr << "  [START] " << name << '\n' << std::flush;
        fn();
        std::cerr << "  [DONE] " << name << '\n' << std::flush;
    };

    runTest("TestFolderViewIncrementalSearchKeepsContainsHighlightButUsesPrefixFocus", TestFolderViewIncrementalSearchKeepsContainsHighlightButUsesPrefixFocus);
    runTest("TestFolderViewInactiveVisualStateDimsNormalTextAndIcons", TestFolderViewInactiveVisualStateDimsNormalTextAndIcons);
    runTest("TestFolderViewEmptyPlaceholderMetricsUseCurrentEmptyLayout", TestFolderViewEmptyPlaceholderMetricsUseCurrentEmptyLayout);
    runTest("TestMenuMnemonicHonorsExplicitAmpersandLabels", TestMenuMnemonicHonorsExplicitAmpersandLabels);
    runTest("TestMenuOpeningPointerUpCanBeIgnoredOutsideVisibleSurface", TestMenuOpeningPointerUpCanBeIgnoredOutsideVisibleSurface);
    runTest("TestEmbeddedViewerContextMenuNativeConversionFiltersStandaloneCommands", TestEmbeddedViewerContextMenuNativeConversionFiltersStandaloneCommands);
    runTest("TestMenuShadowMarginMouseUpLightDismissesAfterInitialRelease", TestMenuShadowMarginMouseUpLightDismissesAfterInitialRelease);
    runTest("TestMenuKeyboardNavigationSkipsInfoRows", TestMenuKeyboardNavigationSkipsInfoRows);
    runTest("TestMenuKeyboardRightArrowMatchesWindowsMenuLoop", TestMenuKeyboardRightArrowMatchesWindowsMenuLoop);
    runTest("TestStationaryMouseDoesNotOverrideKeyboardRootSwitch", TestStationaryMouseDoesNotOverrideKeyboardRootSwitch);
    runTest("TestMenuPointerOverSiblingRootSwitchesOpenMenu", TestMenuPointerOverSiblingRootSwitchesOpenMenu);
    runTest("TestMenuBarHoverMessageSwitchesRootWhenCursorOutsidePopup", TestMenuBarHoverMessageSwitchesRootWhenCursorOutsidePopup);
    runTest("TestMenuPointerInsideOverlappingPopupDoesNotSwitchRoot", TestMenuPointerInsideOverlappingPopupDoesNotSwitchRoot);
    runTest("TestMenuRootSwitchIgnoresStaleMouseMoveAfterCursorSwitch", TestMenuRootSwitchIgnoresStaleMouseMoveAfterCursorSwitch);
    runTest("TestMenuRootSwitchDoesNotPollCursorWhileIdle", TestMenuRootSwitchDoesNotPollCursorWhileIdle);
    runTest("TestMenuHoveringSiblingClosesOpenSubmenuAfterDelay", TestMenuHoveringSiblingClosesOpenSubmenuAfterDelay);
    runTest("TestMenuHoveringSiblingWithChildrenReplacesOpenSubmenuAfterDelay", TestMenuHoveringSiblingWithChildrenReplacesOpenSubmenuAfterDelay);
    runTest("TestMenuPointerInsideSubmenuAndParentItemCancelPendingCloseDelay", TestMenuPointerInsideSubmenuAndParentItemCancelPendingCloseDelay);
    runTest("TestMenuKeyboardTabExitsMenuLoop", TestMenuKeyboardTabExitsMenuLoop);
    runTest("TestMenuKeyboardF10ExitsMenuLoop", TestMenuKeyboardF10ExitsMenuLoop);
    runTest("TestMenuKeyboardAltExitsMenuLoop", TestMenuKeyboardAltExitsMenuLoop);
    runTest("TestMenuKeyboardLeftArrowMatchesWindowsMenuLoop", TestMenuKeyboardLeftArrowMatchesWindowsMenuLoop);
    runTest("TestNativeMenuBarRestoresFocusAfterMenuDismiss", TestNativeMenuBarRestoresFocusAfterMenuDismiss);
    runTest("TestMenuInfoRowsDoNotDismissOnClick", TestMenuInfoRowsDoNotDismissOnClick);
    runTest("TestMenuPopupPositionClampsAcrossDpiMatrix", TestMenuPopupPositionClampsAcrossDpiMatrix);
    runTest("TestMenuPopupPositionSupportsRightAlignedRootAnchors", TestMenuPopupPositionSupportsRightAlignedRootAnchors);
    runTest("TestMenuPopupPositionSupportsAboveRightAlignedRootAnchors", TestMenuPopupPositionSupportsAboveRightAlignedRootAnchors);
    runTest("TestMenuInfoRowsUseMeasuredValueColumnWidth", TestMenuInfoRowsUseMeasuredValueColumnWidth);
    runTest("TestMenuStandardRowsDeriveAndAlignShortcutColumnFromTabbedText", TestMenuStandardRowsDeriveAndAlignShortcutColumnFromTabbedText);
    runTest("TestMenuBitmapIconsReachPopupLayout", TestMenuBitmapIconsReachPopupLayout);
    runTest("TestMenuPopupMaterialsProduceDistinctCaptures", TestMenuPopupMaterialsProduceDistinctCaptures);
    runTest("TestMenuRainbowHoverUsesSeededHighlightContrast", TestMenuRainbowHoverUsesSeededHighlightContrast);
    runTest("TestMenuHoverContrastAppliesToGlyphsAcrossThemes", TestMenuHoverContrastAppliesToGlyphsAcrossThemes);
    runTest("TestMenuRainbowCheckedItemUsesAccentIndicator", TestMenuRainbowCheckedItemUsesAccentIndicator);
    runTest("TestMenuCheckedRowsDoNotPaintSecondFullRowSelection", TestMenuCheckedRowsDoNotPaintSecondFullRowSelection);
    runTest("TestMenuCheckedRowsDoNotPaintLeadingCheckedBox", TestMenuCheckedRowsDoNotPaintLeadingCheckedBox);
    runTest("TestMenuPopupCompositionHostUsesTransparentShadowMargins", TestMenuPopupCompositionHostUsesTransparentShadowMargins);
    runTest("TestMenuPopupWindowClassDoesNotUseNativeDropShadow", TestMenuPopupWindowClassDoesNotUseNativeDropShadow);
    runTest("TestMenuPopupSubmenuFlipsLeftNearRightEdge", TestMenuPopupSubmenuFlipsLeftNearRightEdge);
    runTest("TestMenuPopupAcrylicLightVisualBaseline", TestMenuPopupAcrylicLightVisualBaseline);
    runTest("TestMenuPopupKeepsSystemBackdropDisabledForAppRenderedMaterials", TestMenuPopupKeepsSystemBackdropDisabledForAppRenderedMaterials);
    runTest("TestMenuAcrylicBackdropScenarioEmitsMetrics", TestMenuAcrylicBackdropScenarioEmitsMetrics);
}
