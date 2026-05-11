#pragma once

#include <algorithm>

namespace RedSalamanderMonitor
{
struct ColorTextScrollBarInputs
{
    float clientWidthDip               = 0.0f;
    float clientHeightDip              = 0.0f;
    float contentWidthDip              = 0.0f;
    float contentHeightDip             = 0.0f;
    float textChromeWidthDip           = 0.0f;
    float verticalScrollbarWidthDip    = 0.0f;
    float horizontalScrollbarHeightDip = 0.0f;
    bool currentVerticalVisible        = false;
    bool currentHorizontalVisible      = false;
};

struct ColorTextScrollBarState
{
    bool verticalVisible      = false;
    bool horizontalVisible    = false;
    float viewportWidthDip    = 0.0f;
    float viewportHeightDip   = 0.0f;
    float verticalPageDip     = 0.0f;
    float horizontalPageDip   = 0.0f;
    float textAvailableDip    = 0.0f;
};

[[nodiscard]] inline ColorTextScrollBarState ComputeColorTextViewScrollBars(const ColorTextScrollBarInputs& inputs) noexcept
{
    const float noScrollWidthDip =
        (std::max)(0.0f, inputs.clientWidthDip + (inputs.currentVerticalVisible ? inputs.verticalScrollbarWidthDip : 0.0f));
    const float noScrollHeightDip =
        (std::max)(0.0f, inputs.clientHeightDip + (inputs.currentHorizontalVisible ? inputs.horizontalScrollbarHeightDip : 0.0f));

    bool verticalVisible   = false;
    bool horizontalVisible = false;

    for (int pass = 0; pass < 4; ++pass)
    {
        const float viewportWidthDip  = (std::max)(0.0f, noScrollWidthDip - (verticalVisible ? inputs.verticalScrollbarWidthDip : 0.0f));
        const float viewportHeightDip = (std::max)(0.0f, noScrollHeightDip - (horizontalVisible ? inputs.horizontalScrollbarHeightDip : 0.0f));
        const float textAvailableDip  = (std::max)(0.0f, viewportWidthDip - inputs.textChromeWidthDip);

        const bool nextVerticalVisible   = inputs.contentHeightDip > viewportHeightDip + 0.5f;
        const bool nextHorizontalVisible = inputs.contentWidthDip > textAvailableDip + 0.5f;
        if (nextVerticalVisible == verticalVisible && nextHorizontalVisible == horizontalVisible)
        {
            break;
        }

        verticalVisible   = nextVerticalVisible;
        horizontalVisible = nextHorizontalVisible;
    }

    ColorTextScrollBarState state{};
    state.verticalVisible   = verticalVisible;
    state.horizontalVisible = horizontalVisible;
    state.viewportWidthDip  = (std::max)(0.0f, noScrollWidthDip - (verticalVisible ? inputs.verticalScrollbarWidthDip : 0.0f));
    state.viewportHeightDip = (std::max)(0.0f, noScrollHeightDip - (horizontalVisible ? inputs.horizontalScrollbarHeightDip : 0.0f));
    state.verticalPageDip   = state.viewportHeightDip;
    state.textAvailableDip  = (std::max)(0.0f, state.viewportWidthDip - inputs.textChromeWidthDip);
    state.horizontalPageDip = state.textAvailableDip;
    return state;
}
} // namespace RedSalamanderMonitor
