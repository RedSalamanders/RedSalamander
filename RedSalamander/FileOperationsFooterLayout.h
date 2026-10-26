#pragma once

#include <DxUi/DxUi.h>

#include <algorithm>
#include <array>
#include <cmath>

// Application policy only: keep the settings cluster together and right aligned.
// DxUi owns the allocation-free measured row placement. Coordinates are local DIPs
// relative to the footer's command area; text measurement remains with the host.
namespace FileOperationsFooterLayout
{
inline constexpr int kMinimumClientWidthDip = 480;
inline constexpr float kMarginDip           = 10.0f;
inline constexpr float kButtonHeightDip     = 28.0f;
inline constexpr float kGapDip              = 8.0f;
inline constexpr float kSummaryTopDip       = 18.0f;
inline constexpr float kSummaryHeightDip    = 20.0f;
inline constexpr float kSummaryGapDip       = 10.0f;
inline constexpr float kCommandTopDip       = kSummaryTopDip + kSummaryHeightDip + kSummaryGapDip;
inline constexpr float kBottomDip           = 12.0f;
inline constexpr float kDefaultHeightDip    = kCommandTopDip + kButtonHeightDip + kBottomDip;

struct Layout final
{
    D2D1_RECT_F pauseResume{};
    D2D1_RECT_F cancel{};
    D2D1_RECT_F queue{};
    D2D1_RECT_F options{};
    D2D1_RECT_F details{};
    float heightDip     = kDefaultHeightDip;
    float commandTopDip = kCommandTopDip;
    size_t rowCount     = 0u;
};

// Zero omits a transfer action. Nonzero widths must be measured to fit the
// available command width. Failure leaves the caller's previous layout intact.
[[nodiscard]] inline HRESULT Arrange(
    float availableWidthDip, float pauseResumeWidthDip, float cancelWidthDip, Layout& result, float summaryHeightDip = kSummaryHeightDip) noexcept
{
    if (! std::isfinite(summaryHeightDip) || summaryHeightDip < 0.0f)
        return E_INVALIDARG;
    constexpr float settingsChromeDip  = 2.0f * (kButtonHeightDip + kGapDip);
    constexpr float minimumSettingsDip = 112.0f + settingsChromeDip;
    constexpr float idealSettingsDip   = 168.0f + settingsChromeDip;
    const std::array sizes{D2D1::SizeF(pauseResumeWidthDip, pauseResumeWidthDip == 0.0f ? 0.0f : kButtonHeightDip),
                           D2D1::SizeF(cancelWidthDip, cancelWidthDip == 0.0f ? 0.0f : kButtonHeightDip),
                           D2D1::SizeF(minimumSettingsDip, kButtonHeightDip)};
    std::array<D2D1_RECT_F, 3> groups{};
    DxUi::MeasuredActionLayout measured{};
    const HRESULT hr = DxUi::ArrangeMeasuredActions(sizes, availableWidthDip, D2D1::SizeF(kGapDip, kGapDip), groups, measured);
    if (FAILED(hr))
        return hr;

    // The third group is indivisible. Grow Queue within its row's spare room,
    // preserving the transfer controls and the two square settings affordances.
    const float settingsWidth = (std::min)(idealSettingsDip, availableWidthDip - groups[2].left);
    const float left          = availableWidthDip - settingsWidth;
    const float top           = groups[2].top;
    Layout arranged{};
    arranged.pauseResume   = groups[0];
    arranged.cancel        = groups[1];
    arranged.queue         = D2D1::RectF(left, top, left + settingsWidth - settingsChromeDip, top + kButtonHeightDip);
    arranged.options       = D2D1::RectF(arranged.queue.right + kGapDip, top, arranged.queue.right + kGapDip + kButtonHeightDip, top + kButtonHeightDip);
    arranged.details       = D2D1::RectF(availableWidthDip - kButtonHeightDip, top, availableWidthDip, top + kButtonHeightDip);
    arranged.commandTopDip = kSummaryTopDip + (std::max)(kSummaryHeightDip, summaryHeightDip) + kSummaryGapDip;
    arranged.heightDip     = arranged.commandTopDip + measured.heightDip + kBottomDip;
    if (! std::isfinite(arranged.heightDip))
        return E_INVALIDARG;
    arranged.rowCount = measured.rowCount;
    result            = arranged;
    return S_OK;
}
} // namespace FileOperationsFooterLayout
