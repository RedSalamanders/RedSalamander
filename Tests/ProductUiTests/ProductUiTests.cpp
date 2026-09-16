#define REDSAL_DEFINE_TRACE_PROVIDER
#include "DxUiThemePalette.h"
#include "FileMetadataFormatting.h"
#include "FolderViewEmptyStateLayout.h"
#include "FolderViewIncrementalSearch.h"
#include "HandleIo.h"
#include "Helpers.h"
#include "PaginationGuard.h"
#include "PaneVisualState.h"
#include "PathUtils.h"
#include "SettingsStore.h"
#include "TestSupport/TestSupport.h"
#include "ThroughputParsing.h"
#include "UriEncoding.h"
#include "ViewerDxUiTheme.h"
#include "ViewerFileComboHost.h"
#include "ViewerTitleBarTheme.h"
#include "WindowMessages.h"
#include "WindowSizing.h"
#include "YyjsonHelpers.h"
#include <DxUi/DxUi.h>
#include <atomic>
#include <chrono>
#include <cstdlib>
#include <iostream>
#include <thread>

namespace
{
inline void Require(bool condition, const char* message)
{
    if (! condition)
    {
        std::cerr << "FAILED: " << message << '\n';
        std::exit(1);
    }
}

inline void RequireColorNear(const D2D1_COLOR_F& actual, const D2D1_COLOR_F& expected, const char* message)
{
    const auto nearlyEqual = [](float a, float b) noexcept { return std::fabs(a - b) <= 0.0001f; };

    if (! nearlyEqual(actual.r, expected.r) || ! nearlyEqual(actual.g, expected.g) || ! nearlyEqual(actual.b, expected.b) ||
        ! nearlyEqual(actual.a, expected.a))
    {
        std::cerr << "FAILED: " << message << '\n';
        std::exit(1);
    }
}

inline void RequireColorDifferent(const D2D1_COLOR_F& lhs, const D2D1_COLOR_F& rhs, const char* message)
{
    const auto nearlyEqual = [](float a, float b) noexcept { return std::fabs(a - b) <= 0.0001f; };

    if (nearlyEqual(lhs.r, rhs.r) && nearlyEqual(lhs.g, rhs.g) && nearlyEqual(lhs.b, rhs.b) && nearlyEqual(lhs.a, rhs.a))
    {
        std::cerr << "FAILED: " << message << '\n';
        std::exit(1);
    }
}

inline void RequireFloatNear(float actual, float expected, float epsilon, const char* message)
{
    if (std::fabs(actual - expected) > epsilon)
    {
        std::cerr << "FAILED: " << message << '\n';
        std::exit(1);
    }
}

inline D2D1_COLOR_F BlendForTest(const D2D1_COLOR_F& a, const D2D1_COLOR_F& b, float t) noexcept
{
    const float clamped = std::clamp(t, 0.0f, 1.0f);
    return D2D1::ColorF(a.r + ((b.r - a.r) * clamped), a.g + ((b.g - a.g) * clamped), a.b + ((b.b - a.b) * clamped), a.a + ((b.a - a.a) * clamped));
}

struct PostedPayloadDrainStressWindowState final
{
    PostedPayloadDrainStressWindowState()                                                      = default;
    PostedPayloadDrainStressWindowState(const PostedPayloadDrainStressWindowState&)            = delete;
    PostedPayloadDrainStressWindowState(PostedPayloadDrainStressWindowState&&)                 = delete;
    PostedPayloadDrainStressWindowState& operator=(const PostedPayloadDrainStressWindowState&) = delete;
    PostedPayloadDrainStressWindowState& operator=(PostedPayloadDrainStressWindowState&&)      = delete;

    std::atomic<uint32_t> deliveredCount{0};
    std::atomic<uint32_t> drainedCount{0};
    std::atomic<uint32_t> staleTokenCount{0};
    std::atomic<uint32_t> staleTokenRejectionCount{0};
    std::atomic<uint64_t> drainDurationUs{0};
    std::atomic<bool> payloadQueuedBeforeDrain{false};
    std::atomic<bool> payloadQueuedAfterDrain{false};
};

struct PostedPayloadDrainStressPayload final
{
    std::atomic<uint32_t>* destroyedCount = nullptr;
    std::vector<uint32_t> values;

    ~PostedPayloadDrainStressPayload()
    {
        if (destroyedCount)
        {
            destroyedCount->fetch_add(1u, std::memory_order_acq_rel);
        }
    }
};

LRESULT CALLBACK PostedPayloadDrainStressWndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    auto* state = reinterpret_cast<PostedPayloadDrainStressWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (message)
    {
        case WM_NCCREATE:
        {
            const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            state              = static_cast<PostedPayloadDrainStressWindowState*>(create ? create->lpCreateParams : nullptr);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
            InitPostedPayloadWindow(hwnd);
            return TRUE;
        }

        case WndMsg::kFolderViewEnumerateComplete:
        {
            auto payload = TakeMessagePayload<PostedPayloadDrainStressPayload>(lParam);
            if (state && payload)
            {
                state->deliveredCount.fetch_add(1u, std::memory_order_acq_rel);
            }
            return 0;
        }

        case WM_NCDESTROY:
        {
            MSG queuedMessage{};
            if (state)
            {
                state->payloadQueuedBeforeDrain.store(
                    PeekMessageW(&queuedMessage, hwnd, WndMsg::kFolderViewEnumerateComplete, WndMsg::kFolderViewEnumerateComplete, PM_NOREMOVE) != 0,
                    std::memory_order_release);
            }
            const auto drainStarted = std::chrono::steady_clock::now();
            const size_t drained    = DrainPostedPayloadsForWindow(hwnd);
            const auto drainDurationUs =
                static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - drainStarted).count());
            if (state)
            {
                state->drainedCount.store(static_cast<uint32_t>(drained), std::memory_order_release);
                state->drainDurationUs.store(drainDurationUs, std::memory_order_release);
                state->payloadQueuedAfterDrain.store(
                    PeekMessageW(&queuedMessage, hwnd, WndMsg::kFolderViewEnumerateComplete, WndMsg::kFolderViewEnumerateComplete, PM_NOREMOVE) != 0,
                    std::memory_order_release);

                uint32_t staleTokenCount          = 0u;
                uint32_t staleTokenRejectionCount = 0u;
                while (PeekMessageW(&queuedMessage, hwnd, WndMsg::kFolderViewEnumerateComplete, WndMsg::kFolderViewEnumerateComplete, PM_REMOVE) != 0)
                {
                    ++staleTokenCount;
                    if (! TakeMessagePayload<PostedPayloadDrainStressPayload>(queuedMessage.lParam))
                    {
                        ++staleTokenRejectionCount;
                    }
                }
                state->staleTokenCount.store(staleTokenCount, std::memory_order_release);
                state->staleTokenRejectionCount.store(staleTokenRejectionCount, std::memory_order_release);
            }
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            break;
        }
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

[[nodiscard]] wil::unique_hwnd CreatePostedPayloadDrainStressWindow(PostedPayloadDrainStressWindowState& state)
{
    constexpr wchar_t kClassName[] = L"RedSalamander.DxUiTests.PostedPayloadDrainStressWindow";

    HINSTANCE instance = GetModuleHandleW(nullptr);
    WNDCLASSW existing{};
    if (GetClassInfoW(instance, kClassName, &existing) == FALSE)
    {
        WNDCLASSW wc{};
        wc.lpfnWndProc   = PostedPayloadDrainStressWndProc;
        wc.hInstance     = instance;
        wc.lpszClassName = kClassName;
        Require(RegisterClassW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS, "posted-payload drain stress window class registers");
    }

    return wil::unique_hwnd(CreateWindowExW(0, kClassName, L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, instance, &state));
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
    using namespace Common::PaneVisualState;

    const auto same = [](float left, float right) noexcept { return std::fabs(left - right) <= 0.0001f; };

    Require(same(ResolveNormalTextAlpha(1.0f, true, false), 1.0f), "focused pane keeps normal text fully opaque");
    Require(same(ResolveNormalTextAlpha(1.0f, false, false), kUnfocusedPaneTextOpacity), "unfocused pane dims normal text");
    Require(same(ResolveNormalTextAlpha(1.0f, false, true), 1.0f), "selected text uses inactive-selection colors instead of normal text dimming");
    Require(same(ResolveNormalIconOpacity(1.0f, true), 1.0f), "focused pane keeps normal icons fully opaque");
    Require(same(ResolveNormalIconOpacity(1.0f, false), kUnfocusedPaneIconOpacity), "unfocused pane dims normal icons");
    Require(same(ResolveNormalIconOpacity(0.5f, false), 0.5f * kUnfocusedPaneIconOpacity), "hidden icons keep their hidden dim and get pane dimming");
    Require(same(ResolvePlaceholderIconOpacity(false), 0.4f * kUnfocusedPaneIconOpacity), "placeholder icons also dim in an unfocused pane");
    Require(same(ResolveFocusBorderAlpha(1.0f, false), kFocusBorderOpacityUnfocused), "unfocused current item keeps a dim focus border");
    Require(same(ResolveInactiveContentOverlayAlpha(true, false), 0.0f), "focused pane does not apply an inactive-content overlay");
    Require(same(ResolveInactiveContentOverlayAlpha(false, false), 1.0f - kUnfocusedPaneTextOpacity),
            "unfocused pane overlay matches the standard text dimming opacity");
    Require(same(ResolveInactiveContentOverlayAlpha(false, true), 0.0f), "High Contrast suppresses the inactive-content overlay");
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

void TestSharedColorRefArgbAndTruncatingBlendGoldenValues()
{
    using namespace Common::Colors;

    Require(ColorRefFromArgb(0x7F123456u) == RGB(0x12, 0x34, 0x56), "ARGB-to-COLORREF conversion drops alpha and preserves RGB byte order");
    Require(BlendColorRefTruncate(RGB(10, 20, 30), RGB(110, 220, 230), 128u) == RGB(60, 120, 130),
            "COLORREF blend preserves the existing truncating 8-bit channel policy");
    Require(BlendColorRefTruncate(RGB(1, 2, 3), RGB(200, 201, 202), 0u) == RGB(1, 2, 3), "COLORREF blend alpha zero preserves the under color");
    Require(BlendColorRefTruncate(RGB(1, 2, 3), RGB(200, 201, 202), 255u) == RGB(200, 201, 202), "COLORREF blend alpha 255 selects the over color");
    Require(BlendColorRefWeightedTruncate(RGB(10, 20, 30), RGB(110, 220, 230), 1, 2) == RGB(60, 120, 130),
            "weighted COLORREF blend preserves the existing arbitrary-denominator truncation policy");
    Require(BlendColorRefWeightedTruncate(RGB(10, 20, 30), RGB(110, 220, 230), -1, 10) == RGB(10, 20, 30),
            "weighted COLORREF blend clamps a negative overlay weight to zero");
    Require(BlendColorRefWeightedTruncate(RGB(10, 20, 30), RGB(110, 220, 230), 20, 10) == RGB(110, 220, 230),
            "weighted COLORREF blend clamps an oversized overlay weight to the denominator");
    Require(BlendColorRefWeightedTruncate(RGB(10, 20, 30), RGB(110, 220, 230), 1, 0) == RGB(10, 20, 30),
            "weighted COLORREF blend preserves the base color for an invalid denominator");
}

void TestHsvColorRefNegativeHuePoliciesRemainExplicitAndDistinct()
{
    using namespace Common::Colors;

    Require(ColorRefFromHsvClampedNegativeHueToZero(-60.0f, 1.0f, 1.0f) == RGB(255, 0, 0), "majority viewer HSV policy clamps negative hue to red");
    Require(ColorRefFromHsvWrappedHue(-60.0f, 1.0f, 1.0f) == RGB(255, 0, 255), "ViewerWeb HSV policy wraps negative hue to magenta");
    Require(ColorRefFromHsvClampedNegativeHueToZero(120.0f, 2.0f, 2.0f) == RGB(0, 255, 0), "majority viewer HSV policy clamps saturation and value");
    Require(ColorRefFromHsvClampedNegativeHueToZero(240.0f, 1.0f, 1.0f) == RGB(0, 0, 255), "majority viewer HSV policy preserves primary blue");
}

void TestSharedLuminanceMathSeparatesLinearizedAndEncodedPolicies()
{
    using namespace Common::Colors;

    Require(std::abs(RelativeLuminanceFromArgb(0xFF000000u)) < 0.0000001, "relative luminance maps black to zero");
    Require(std::abs(RelativeLuminanceFromArgb(0xFFFFFFFFu) - 1.0) < 0.0000001, "relative luminance maps white to one");
    Require(std::abs(RelativeLuminanceFromArgb(0xFFFF0000u) - 0.2126) < 0.0000001, "relative luminance preserves the linear red coefficient");
    Require(std::abs(RelativeLuminanceFromColorRef(RGB(128, 128, 128)) - 0.2158605001) < 0.0000001, "relative luminance linearizes an encoded mid-gray");
    Require(std::abs(WeightedSrgbLuminanceWithoutLinearization(0.5, 0.5, 0.5) - 0.5) < 0.0000001,
            "encoded-channel luminance remains a separately named threshold policy");
    Require(std::abs(ContrastRatioFromRelativeLuminance(1.0, 0.0) - 21.0) < 0.0000001,
            "relative-luminance contrast ratio preserves the black/white golden value");
}

void TestWindowSizingDpiRoundingAndPopupGeometryGoldenValues()
{
    using namespace Common::WindowSizing;

    Require(DipToPixelRounded(96u, 8) == 8, "integer DIP conversion preserves values at 96 DPI");
    Require(DipToPixelRounded(120u, 8) == 10, "integer DIP conversion scales values at 120 DPI");
    Require(DipToPixelRounded(144u, 8) == 12, "integer DIP conversion scales values at 144 DPI");
    Require(DipToPixelRounded(192u, 8) == 16, "integer DIP conversion scales values at 192 DPI");
    Require(DipToPixelRounded(120u, -8) == -10, "integer DIP conversion preserves negative coordinates");
    Require(DipToPixelRounded(0.5f, 96u) == 1 && DipToPixelRounded(-0.5f, 96u) == -1, "fractional DIP conversion rounds half away from zero");
    Require(DipToPixelRounded((std::numeric_limits<float>::max)(), 192u) == (std::numeric_limits<int>::max)(),
            "fractional DIP conversion saturates positive overflow");
    Require(std::abs(PixelToDip(15.0f, 144.0f) - 10.0f) < 0.0001f, "pixel-to-DIP conversion preserves fractional monitor scaling");
    Require(std::abs(PixelToDip(15.0f, 0.0f) - 15.0f) < 0.0001f, "pixel-to-DIP conversion treats zero DPI as 96 DPI");

    Require(ComputeBoundedListPopupHeightPx(0u, 8u, 24, 10, 96u) == 34, "bounded popup height keeps one visible row for an empty list host");
    Require(ComputeBoundedListPopupHeightPx(100u, 8u, 24, 10, 192u) == 404, "bounded popup height caps visible rows before DPI scaling");

    const RECT work{-50, -40, 250, 160};
    const RECT clamped = ClampRectOriginToBounds(RECT{-200, -100, -100, 0}, work);
    Require(clamped.left == -50 && clamped.top == -40 && clamped.right == 50 && clamped.bottom == 60,
            "work-area clamping preserves size across negative monitor coordinates");
    const RECT oversized = ClampRectOriginToBounds(RECT{0, 0, 500, 300}, work);
    Require(oversized.left == -50 && oversized.top == -40 && oversized.right == 450 && oversized.bottom == 260,
            "oversized popup clamping pins the origin to the work-area origin");

    const RECT centered = CenterRectOnOwner(RECT{0, 0, 100, 50}, RECT{-300, -200, 300, 200});
    Require(centered.left == -50 && centered.top == -25 && centered.right == 50 && centered.bottom == 25,
            "owner centering preserves the window size across negative monitor coordinates");
    const RECT oddCentered = CenterRectOnOwner(RECT{10, 20, 111, 71}, RECT{0, 0, 400, 300});
    Require(oddCentered.left == 149 && oddCentered.top == 124 && oddCentered.right == 250 && oddCentered.bottom == 175,
            "owner centering preserves the existing integer midpoint rounding policy");

    const RECT verticalWork{-100, -50, 500, 450};
    Require(ResolveBoundedPopupTop(100, 120, 100, 4, verticalWork) == 124, "bounded popup placement prefers below when it fits");
    Require(ResolveBoundedPopupTop(400, 420, 100, 4, verticalWork) == 296, "bounded popup placement flips above near the bottom edge");
    Require(ResolveBoundedPopupTop(100, 120, 600, 4, verticalWork) == -50, "oversized popup placement clamps to the work-area top");
}

void TestViewerFileComboPopupInputAndHeightGoldenValues()
{
    using namespace RedSalamander::ViewerFileComboHost;

    Require(MessageMayOpenWindowComboPopup(WM_LBUTTONDOWN, 0u), "file combo opens for a primary-button press");
    Require(MessageMayOpenWindowComboPopup(WM_LBUTTONDBLCLK, 0u), "file combo opens for a primary-button double click");
    Require(MessageMayOpenWindowComboPopup(WM_SYSKEYDOWN, VK_DOWN) && MessageMayOpenWindowComboPopup(WM_SYSKEYDOWN, VK_UP),
            "file combo preserves Alt+Down and Alt+Up popup input");
    Require(MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_SPACE) && MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_RETURN) &&
                MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_F4) && MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_DOWN) &&
                MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_UP),
            "file combo preserves every keyboard popup input");
    Require(! MessageMayOpenWindowComboPopup(WM_KEYDOWN, VK_ESCAPE) && ! MessageMayOpenWindowComboPopup(WM_KEYUP, VK_F4) &&
                ! MessageMayOpenWindowComboPopup(WM_RBUTTONDOWN, 0u),
            "file combo rejects close keys, key-up, and secondary-button messages");
    Require(ComputeStandaloneComboPopupHeightPx(0u, 96u) == 34, "standalone file combo keeps one row plus chrome for an empty list");
    Require(ComputeStandaloneComboPopupHeightPx(100u, 192u) == 404, "standalone file combo caps its popup at eight visible rows before DPI scaling");
}

void TestViewerTitleBarThemeGoldenValues()
{
    using namespace RedSalamander::ViewerChrome;

    ViewerTheme theme{.sizeBytes = sizeof(ViewerTheme)};
    theme.darkMode       = TRUE;
    theme.darkBase       = TRUE;
    theme.rainbowMode    = TRUE;
    theme.backgroundArgb = 0xFF101820u;
    theme.accentArgb     = 0xFF336699u;

    const COLORREF accent           = ResolveViewerThemeAccent(theme, L"title-golden");
    const TitleBarAttributes active = ResolveTitleBarAttributes(theme, true, L"title-golden");
    Require(active.useDarkMode == TRUE && active.borderColor == accent && active.captionColor == accent,
            "rainbow viewer title bar applies its deterministic accent while active");
    const uint32_t luma =
        (static_cast<uint32_t>(GetRValue(accent)) * 299u + static_cast<uint32_t>(GetGValue(accent)) * 587u + static_cast<uint32_t>(GetBValue(accent)) * 114u) /
        1000u;
    Require(active.textColor == (luma < 128u ? RGB(255, 255, 255) : RGB(0, 0, 0)), "rainbow viewer title bar preserves the integer contrast threshold");

    const TitleBarAttributes inactive = ResolveTitleBarAttributes(theme, false, L"title-golden");
    const COLORREF expectedInactive   = Common::Colors::BlendColorRefTruncate(accent, Common::Colors::ColorRefFromArgb(theme.backgroundArgb), 223u);
    Require(inactive.captionColor == expectedInactive && inactive.borderColor == expectedInactive,
            "inactive rainbow viewer title bar preserves the 223/255 background blend");

    theme.rainbowMode              = FALSE;
    const TitleBarAttributes plain = ResolveTitleBarAttributes(theme, true, L"ignored");
    Require(plain.useDarkMode == TRUE && plain.captionColor == kDwmColorDefault && plain.borderColor == kDwmColorDefault && plain.textColor == kDwmColorDefault,
            "non-rainbow viewer title bar resets custom DWM colors while retaining dark mode");

    theme.highContrast                    = TRUE;
    const TitleBarAttributes highContrast = ResolveTitleBarAttributes(theme, true, L"ignored");
    Require(highContrast.useDarkMode == FALSE && highContrast.captionColor == kDwmColorDefault,
            "high-contrast viewer title bar delegates all colors and dark-mode policy to the system");
}

void TestUtfConversionPoliciesRemainExplicitAndPreserveEmbeddedNulls()
{
    using namespace Common::Strings;

    const std::string malformedUtf8{"\xC3\x28", 2u};
    Require(! IsValidUtf8Strict(malformedUtf8), "strict UTF-8 validation rejects an invalid continuation byte without conversion");
    Require(! TryUtf16FromUtf8Strict(malformedUtf8).has_value(), "strict UTF-8 conversion rejects an invalid continuation byte");
    Require(Utf16FromUtf8ReplacingInvalid(malformedUtf8) == std::wstring{L"\uFFFD("},
            "replacement UTF-8 conversion retains displayable text after an invalid sequence");
    Require(! TryUtf16FromUtf8Strict(std::string{"\xF0\x9F\x98", 3u}).has_value(), "strict UTF-8 conversion rejects a truncated sequence");
    Require(! IsValidUtf8Strict(std::string{"\xF0\x9F\x98", 3u}), "strict UTF-8 validation rejects a truncated sequence");

    const std::string utf8WithNull{"A\0\xF0\x9F\x98\x80", 6u};
    const std::optional<std::wstring> utf16WithNull = TryUtf16FromUtf8Strict(utf8WithNull);
    Require(utf16WithNull.has_value() && utf16WithNull.value().size() == 4u && utf16WithNull.value()[0] == L'A' && utf16WithNull.value()[1] == L'\0' &&
                utf16WithNull.value()[2] == L'\xD83D' && utf16WithNull.value()[3] == L'\xDE00',
            "strict UTF-8 conversion preserves embedded NUL and supplementary characters");
    const std::optional<std::string> roundTrip = TryUtf8FromUtf16Strict(utf16WithNull.value());
    Require(roundTrip.has_value() && roundTrip.value() == utf8WithNull, "strict UTF conversion round-trips embedded NUL and supplementary characters");

    const std::wstring unpairedSurrogate{L'\xD800'};
    Require(! TryUtf8FromUtf16Strict(unpairedSurrogate).has_value(), "strict UTF-16 conversion rejects an unpaired surrogate");
    Require(Utf8FromUtf16ReplacingInvalid(unpairedSurrogate) == std::string{"\xEF\xBF\xBD", 3u},
            "replacement UTF-16 conversion emits U+FFFD for an unpaired surrogate");
    Require(TryUtf16FromUtf8Strict({}).has_value() && TryUtf16FromUtf8Strict({}).value().empty(),
            "strict UTF conversion distinguishes valid empty input from malformed input");
    Require(IsValidUtf8Strict({}) && IsValidUtf8Strict("ASCII \xE7\xAB\xAF\xE6\x9C\xAB"), "strict UTF-8 validation accepts empty, ASCII, and multibyte input");

    std::wstring reusable(512u, L'x');
    const wchar_t* const storage = reusable.data();
    const size_t capacity        = reusable.capacity();
    for (const std::string_view text : {std::string_view{"long-enough-to-need-heap-storage.bin"},
                                        std::string_view{utf8WithNull},
                                        std::string_view{"short"},
                                        std::string_view{"\xE7\xAB\xAF\xE6\x9C\xAB"}})
    {
        const auto expected = TryUtf16FromUtf8Strict(text);
        Require(TryUtf16FromUtf8Strict(text, reusable) && expected.has_value() && reusable == expected.value(),
                "reusable strict conversion preserves complete ASCII, embedded NUL, supplementary and multibyte values");
        Require(reusable.data() == storage && reusable.capacity() == capacity, "strict conversion reuses sufficient caller storage");
    }
    Require(! TryUtf16FromUtf8Strict(malformedUtf8, reusable) && reusable.empty(), "failed reusable conversion clears stale output");
    Require(! TryUtf16FromUtf8Strict(std::string_view{"\xF0\x9F\x98", 3u}, reusable) && reusable.empty(),
            "reusable conversion rejects truncated supplementary input");
    Require(TryUtf16FromUtf8Strict({}, reusable) && reusable.empty(), "reusable conversion distinguishes valid empty input from failure");
    Require(TryUtf16FromUtf8Strict("recovered", reusable) && reusable == L"recovered" && reusable.data() == storage,
            "invalid and empty conversions do not poison subsequent storage reuse");
}

void TestOrdinalTrimAndEnvironmentPoliciesRemainExplicit()
{
    Require(OrdinalString::EqualsNoCase(L"FILE", L"file"), "ordinal identifier comparison is case-insensitive without locale folding");
    Require(! OrdinalString::EqualsNoCase(L"I", L"\u0131"), "ordinal identifier comparison does not apply Turkish dotless-I locale rules");
    Require(StringUtils::TrimWhitespace(L" \t\r\nvalue\u2003") == L"value", "shared whitespace trim preserves the existing iswspace character set");
    Require(StringUtils::TrimWhitespace(L"\u00A0value\u00A0") == L"value", "shared whitespace trim handles non-breaking-space boundaries");

    constexpr wchar_t kVariable[]              = L"REDSALAMANDER_LIGHTHOUSE_ENV_POLICY_TEST";
    const std::optional<std::wstring> previous = EnvironmentVariables::Read(kVariable);
    const auto restore = wil::scope_exit([&] { static_cast<void>(SetEnvironmentVariableW(kVariable, previous.has_value() ? previous->c_str() : nullptr)); });

    static_cast<void>(SetEnvironmentVariableW(kVariable, nullptr));
    Require(! EnvironmentVariables::Read(kVariable).has_value(), "environment reads distinguish a missing variable");
    static_cast<void>(SetEnvironmentVariableW(kVariable, L""));
    const std::optional<std::wstring> empty = EnvironmentVariables::Read(kVariable);
    Require(empty.has_value() && empty->empty(), "environment reads distinguish an explicitly empty variable");
    static_cast<void>(SetEnvironmentVariableW(kVariable, L"true"));
    Require(EnvironmentVariables::Read(kVariable).value_or(L"") == L"true", "environment reads preserve the complete value");
    Require(EnvironmentVariables::IsTruthyFlagSet(kVariable), "truthy environment flags accept the established t/y/1 prefixes");
    static_cast<void>(SetEnvironmentVariableW(kVariable, L"false"));
    Require(! EnvironmentVariables::IsTruthyFlagSet(kVariable), "truthy environment flags reject non-truthy prefixes");
}

void TestWindowsPathClassesKeepDriveRelativeAndFullyAbsolutePoliciesDistinct()
{
    using enum Common::Paths::WindowsPathClass;
    using Common::Paths::ClassifyWindowsPath;

    Require(ClassifyWindowsPath(L"folder\\leaf") == Relative, "relative Windows paths remain distinct");
    Require(ClassifyWindowsPath(L"\\folder\\leaf") == Rooted, "rooted current-drive paths are not fully absolute");
    Require(ClassifyWindowsPath(L"C:") == DriveRelative && ClassifyWindowsPath(L"C:leaf") == DriveRelative,
            "drive-qualified paths without a root separator remain drive-relative");
    Require(ClassifyWindowsPath(L"C:\\leaf") == DriveAbsolute && ClassifyWindowsPath(L"d:/leaf") == DriveAbsolute,
            "drive-absolute classification accepts both Windows separator variants");
    Require(ClassifyWindowsPath(L"\\\\server\\share\\leaf") == Unc && ClassifyWindowsPath(L"//server/share/leaf") == Unc,
            "UNC classification accepts slash variants without treating them as drive paths");
    Require(ClassifyWindowsPath(L"\\\\?\\C:\\leaf") == ExtendedDriveAbsolute && ClassifyWindowsPath(L"//?/C:/leaf") == ExtendedDriveAbsolute,
            "extended drive paths remain a named class");
    Require(ClassifyWindowsPath(L"\\\\?\\UNC\\server\\share\\leaf") == ExtendedUnc, "extended UNC paths remain distinct from ordinary UNC paths");
    Require(ClassifyWindowsPath(L"\\\\?\\Volume{guid}\\leaf") == ExtendedOther,
            "non-drive extended namespaces are not silently treated as fully absolute filesystem paths");
    Require(ClassifyWindowsPath(L"\\\\.\\PhysicalDrive0") == Device, "device namespace paths remain explicitly classified");

    Require(Common::Paths::IsDriveQualifiedWindowsPath(L"C:") && ! Common::Paths::IsDriveAbsoluteWindowsPath(L"C:"),
            "drive-qualified and drive-absolute predicates remain separate");
    Require(Common::Paths::IsFullyAbsoluteWindowsPath(L"C:\\leaf") && Common::Paths::IsFullyAbsoluteWindowsPath(L"\\\\server\\share") &&
                Common::Paths::IsFullyAbsoluteWindowsPath(L"\\\\?\\UNC\\server\\share\\leaf"),
            "fully absolute policy accepts rooted drives and complete UNC server/share roots");
    Require(! Common::Paths::IsFullyAbsoluteWindowsPath(L"C:leaf") && ! Common::Paths::IsFullyAbsoluteWindowsPath(L"\\\\server") &&
                ! Common::Paths::IsFullyAbsoluteWindowsPath(L"\\\\.\\PhysicalDrive0"),
            "fully absolute policy rejects drive-relative, incomplete UNC, and device paths");
    Require(ClassifyWindowsPath(L"C:\\dir\\..\\CON. ") == DriveAbsolute, "classification remains normalization-free and does not weaken leaf validation");
    Require(Common::Paths::ToExtendedWin32Path(L"C:\\leaf") == LR"(\\?\C:\leaf)", "normalization-free extended conversion prefixes drive-absolute paths");
    Require(Common::Paths::ToExtendedWin32Path(L"\\\\server\\share\\leaf") == LR"(\\?\UNC\server\share\leaf)",
            "normalization-free extended conversion prefixes UNC paths");
    Require(Common::Paths::ToExtendedWin32Path(L"C:leaf") == L"C:leaf" &&
                Common::Paths::ToExtendedWin32Path(L"\\\\.\\PhysicalDrive0") == L"\\\\.\\PhysicalDrive0",
            "extended conversion leaves drive-relative and device paths unchanged");
}

void TestBinaryThroughputGrammarPreservesBoundaryPoliciesAndRoundTrips()
{
    using Common::Parsing::FormatBinaryThroughputText;
    using Common::Parsing::ThroughputBoundaryWhitespacePolicy;
    using Common::Parsing::TryParseBinaryThroughputText;

    constexpr ThroughputBoundaryWhitespacePolicy ascii = ThroughputBoundaryWhitespacePolicy::AsciiWhitespace;
    constexpr ThroughputBoundaryWhitespacePolicy c0    = ThroughputBoundaryWhitespacePolicy::ControlCharactersThroughSpace;

    uint64_t parsed = 99u;
    Require(TryParseBinaryThroughputText(L"", ascii, parsed) && parsed == 0u, "empty throughput text preserves the unlimited-zero policy");
    Require(TryParseBinaryThroughputText(L" \t1.5 MiB/s\r\n", ascii, parsed) && parsed == 1572864u,
            "throughput parsing accepts fractional values, ASCII boundary whitespace, and the per-second suffix");
    Require(TryParseBinaryThroughputText(L"1,5 mIB/S", ascii, parsed) && parsed == 1572864u,
            "throughput parsing accepts the locale-independent comma decimal and case-insensitive binary alias");
    Require(TryParseBinaryThroughputText(L".5", ascii, parsed) && parsed == 512u, "bare fractional throughput values default to KiB/s");
    Require(TryParseBinaryThroughputText(L"0.0005 KiB", ascii, parsed) && parsed == 1u, "fractional throughput rounds to the nearest whole byte");
    Require(TryParseBinaryThroughputText(L"2 b", ascii, parsed) && parsed == 2u && TryParseBinaryThroughputText(L"2 K", ascii, parsed) && parsed == 2048u &&
                TryParseBinaryThroughputText(L"2 MB", ascii, parsed) && parsed == 2097152u && TryParseBinaryThroughputText(L"2 g", ascii, parsed) &&
                parsed == 2147483648u && TryParseBinaryThroughputText(L"2 TiB", ascii, parsed) && parsed == 2199023255552u &&
                TryParseBinaryThroughputText(L"2 p", ascii, parsed) && parsed == 2251799813685248u,
            "throughput parsing preserves every established byte-through-PiB alias");
    Require(TryParseBinaryThroughputText(L"999999999999999999999999999 PiB", ascii, parsed) && parsed == (std::numeric_limits<uint64_t>::max)(),
            "throughput parsing saturates positive overflow");

    parsed = 99u;
    Require(! TryParseBinaryThroughputText(L"1 XB", ascii, parsed) && parsed == 0u, "unknown throughput units fail and clear the output");
    Require(! TryParseBinaryThroughputText(L"1.2.3 MiB", ascii, parsed), "multiple decimal separators remain invalid");
    Require(! TryParseBinaryThroughputText(L"-1 MiB", ascii, parsed), "signed throughput values remain invalid");

    std::wstring controlWrapped = L"1 MiB";
    controlWrapped.insert(controlWrapped.begin(), static_cast<wchar_t>(0x001F));
    controlWrapped.push_back(static_cast<wchar_t>(0x001F));
    Require(TryParseBinaryThroughputText(controlWrapped, c0, parsed) && parsed == 1048576u,
            "Preferences preserves its historical C0-through-space boundary trimming");
    Require(! TryParseBinaryThroughputText(controlWrapped, ascii, parsed), "the popup preserves its narrower six-character ASCII whitespace boundary");

    constexpr std::array<uint64_t, 7u> roundTripValues{{0u, 1u, 1024u, 1048576u, 1073741824u, 1048577u, (std::numeric_limits<uint64_t>::max)()}};
    for (const uint64_t value : roundTripValues)
    {
        const std::wstring formatted = FormatBinaryThroughputText(value);
        Require(TryParseBinaryThroughputText(formatted, c0, parsed) && parsed == value,
                "the File Operations throughput formatter round-trips through the shared parser");
    }
}

void TestCompactFileMetadataFormattingPreservesTimeAndAttributePolicies()
{
    using Common::FileMetadata::DisplayProfile;
    using Common::FileMetadata::FormatDisplayFields;
    using Common::FileMetadata::NormalizedMetadata;

    const auto empty = FormatDisplayFields({}, DisplayProfile::CompactDetails);
    Require(empty.localTime.empty() && empty.attributes == L"-", "missing time and zero attributes preserve the compact empty/dash representation");

    const auto attributes =
        FormatDisplayFields({.fileAttributes = FILE_ATTRIBUTE_READONLY | FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM | FILE_ATTRIBUTE_ARCHIVE |
                                               FILE_ATTRIBUTE_COMPRESSED | FILE_ATTRIBUTE_ENCRYPTED | FILE_ATTRIBUTE_TEMPORARY | FILE_ATTRIBUTE_OFFLINE |
                                               FILE_ATTRIBUTE_REPARSE_POINT | FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_SPARSE_FILE},
                            DisplayProfile::CompactDetails);
    Require(attributes.attributes == L"RHSACETOP", "compact metadata preserves RHSACETOP ordering and intentionally omits directory and sparse flags");
    const auto sparseCompressed =
        FormatDisplayFields({.fileAttributes = FILE_ATTRIBUTE_SPARSE_FILE | FILE_ATTRIBUTE_COMPRESSED}, DisplayProfile::CompactDetails);
    Require(sparseCompressed.attributes == L"C", "compressed remains visible while sparse remains outside the compact attribute profile");

    Require(FormatDisplayFields({.lastWriteTime100nsSince1601 = -1}, DisplayProfile::CompactDetails).localTime.empty(),
            "negative metadata times remain unknown");
    Require(FormatDisplayFields({.lastWriteTime100nsSince1601 = (std::numeric_limits<int64_t>::max)()}, DisplayProfile::CompactDetails).localTime.empty(),
            "invalid metadata times fail closed to an empty display field");

    SYSTEMTIME utc{};
    utc.wYear   = 2024u;
    utc.wMonth  = 1u;
    utc.wDay    = 15u;
    utc.wHour   = 12u;
    utc.wMinute = 34u;
    utc.wSecond = 56u;
    FILETIME utcFileTime{};
    Require(SystemTimeToFileTime(&utc, &utcFileTime) != FALSE, "metadata test creates a valid UTC FILETIME");
    ULARGE_INTEGER ticks{};
    ticks.LowPart  = utcFileTime.dwLowDateTime;
    ticks.HighPart = utcFileTime.dwHighDateTime;

    FILETIME localFileTime{};
    SYSTEMTIME localSystemTime{};
    Require(FileTimeToLocalFileTime(&utcFileTime, &localFileTime) != FALSE && FileTimeToSystemTime(&localFileTime, &localSystemTime) != FALSE,
            "metadata test resolves the current Windows time-zone conversion");
    const std::wstring expectedLocal = std::format(L"{:04d}-{:02d}-{:02d} {:02d}:{:02d}",
                                                   localSystemTime.wYear,
                                                   localSystemTime.wMonth,
                                                   localSystemTime.wDay,
                                                   localSystemTime.wHour,
                                                   localSystemTime.wMinute);
    Require(FormatDisplayFields({.lastWriteTime100nsSince1601 = static_cast<int64_t>(ticks.QuadPart)}, DisplayProfile::CompactDetails).localTime ==
                expectedLocal,
            "compact metadata uses Windows local time, fixed date ordering, and minute precision independent of display locale");

    SYSTEMTIME preUnixUtc{};
    preUnixUtc.wYear  = 1969u;
    preUnixUtc.wMonth = 12u;
    preUnixUtc.wDay   = 31u;
    preUnixUtc.wHour  = 12u;
    FILETIME preUnixFileTime{};
    Require(SystemTimeToFileTime(&preUnixUtc, &preUnixFileTime) != FALSE, "metadata test creates a valid pre-Unix-epoch FILETIME");
    ticks.LowPart  = preUnixFileTime.dwLowDateTime;
    ticks.HighPart = preUnixFileTime.dwHighDateTime;
    Require(! FormatDisplayFields({.lastWriteTime100nsSince1601 = static_cast<int64_t>(ticks.QuadPart)}, DisplayProfile::CompactDetails).localTime.empty(),
            "valid pre-Unix-epoch Windows times remain displayable");

    Require(FormatBytesCompact(0u) == L"0 B", "the existing shared size formatter preserves exact zero-size display");
}

void TestPercentEncodingAndHandleIoPoliciesCoverForwardProgressAndAtomicSiblings()
{
    using Common::Uri::SlashPolicy;

    Require(Common::Uri::PercentEncodeBytes("AZaz09-._~") == "AZaz09-._~", "RFC3986 unreserved bytes remain literal");
    Require(Common::Uri::PercentEncodeBytes(" /?#%") == "%20%2F%3F%23%25", "reserved and unsafe bytes use uppercase percent escapes");
    Require(Common::Uri::PercentEncodeBytes("a/b c", SlashPolicy::Preserve) == "a/b%20c", "slash preservation is an explicit protocol policy");
    Require(Common::Uri::PercentEncodeBytes(std::string_view("\xC3\xA9", 2u)) == "%C3%A9", "UTF-8 is encoded byte-by-byte");
    std::wstring encodedWide;
    Require(Common::Uri::TryPercentEncodeUtf8ToWide(L"\U0001F600/x", SlashPolicy::Preserve, encodedWide) && encodedWide == L"%F0%9F%98%80/x",
            "wide protocol input is strictly converted to UTF-8 before byte encoding");
    const wchar_t unpaired[] = {static_cast<wchar_t>(0xD800), L'\0'};
    Require(! Common::Uri::TryPercentEncodeUtf8ToWide(unpaired, SlashPolicy::Encode, encodedWide),
            "strict protocol encoding rejects an unpaired UTF-16 surrogate");

    const std::array<std::byte, 5u> payload{};
    unsigned int partialCalls = 0u;
    const HRESULT partialHr =
        Common::HandleIo::Detail::TransferAll<const std::byte>(payload,
                                                               ERROR_WRITE_FAULT,
                                                               [&partialCalls](const std::byte* data, DWORD requested, DWORD& transferred) noexcept
    {
        static_cast<void>(data);
        ++partialCalls;
        transferred = (std::min)(requested, DWORD{2u});
        return S_OK;
    });
    Require(partialHr == S_OK && partialCalls == 3u, "total transfer retries partial forward progress until the buffer is complete");
    const HRESULT zeroProgressHr =
        Common::HandleIo::Detail::TransferAll<const std::byte>(payload,
                                                               ERROR_WRITE_FAULT,
                                                               [](const std::byte* data, DWORD requested, DWORD& transferred) noexcept
    {
        static_cast<void>(data);
        static_cast<void>(requested);
        transferred = 0u;
        return S_OK;
    });
    Require(zeroProgressHr == HRESULT_FROM_WIN32(ERROR_WRITE_FAULT), "successful zero-byte writes fail instead of spinning forever");
    const HRESULT errorHr = Common::HandleIo::Detail::TransferAll<const std::byte>(payload,
                                                                                   ERROR_WRITE_FAULT,
                                                                                   [](const std::byte* data, DWORD requested, DWORD& transferred) noexcept
    {
        static_cast<void>(data);
        static_cast<void>(requested);
        transferred = 0u;
        return HRESULT_FROM_WIN32(ERROR_DISK_FULL);
    });
    Require(errorHr == HRESULT_FROM_WIN32(ERROR_DISK_FULL), "total transfer preserves the underlying I/O error");

    const std::wstring sibling = []
    {
        std::error_code artifactError;
        const auto directory = RedSalamander::TestSupport::AcquireTestDirectory(
            {.harnessSegment = L"product-ui", .leafSegment = L"atomic-siblings", .fallbackRunIdPrefix = L"product-ui"}, artifactError);
        Require(! artifactError && ! directory.empty(), "atomic sibling fixture acquires its governed test directory");
        return (directory / L"lighthouse-target.bin").wstring();
    }();
    const Common::Paths::UniqueSiblingFileOptions options{.prefix = L"lighthouse-", .suffix = L".tmp", .desiredAccess = GENERIC_READ | GENERIC_WRITE};
    std::wstring firstPath;
    std::wstring secondPath;
    wil::unique_hfile firstFile;
    wil::unique_hfile secondFile;
    Require(SUCCEEDED(Common::Paths::CreateUniqueSiblingFile(sibling, options, firstPath, firstFile)) && firstFile,
            "atomic sibling creation returns ownership of the CREATE_NEW reservation");
    Require(SUCCEEDED(Common::Paths::CreateUniqueSiblingFile(sibling, options, secondPath, secondFile)) && secondFile && secondPath != firstPath,
            "independent atomic sibling reservations cannot publish the same path");
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        firstFile.reset();
        secondFile.reset();
        if (! firstPath.empty())
        {
            static_cast<void>(DeleteFileW(firstPath.c_str()));
        }
        if (! secondPath.empty())
        {
            static_cast<void>(DeleteFileW(secondPath.c_str()));
        }
    });

    constexpr std::array<std::byte, 5u> bytes{{std::byte{'R'}, std::byte{'S'}, std::byte{'I'}, std::byte{'O'}, std::byte{'!'}}};
    Require(Common::HandleIo::WriteAll(firstFile.get(), bytes) == S_OK, "shared handle writer writes the complete payload");
    uint64_t fileSize = 0u;
    Require(Common::HandleIo::GetFileSizeBounded(firstFile.get(), bytes.size(), fileSize) == S_OK && fileSize == bytes.size(),
            "bounded handle size accepts a file at the configured maximum");
    Require(Common::HandleIo::GetFileSizeBounded(firstFile.get(), bytes.size() - 1u, fileSize) == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE),
            "bounded handle size rejects oversized files at the domain boundary");
    Require(Common::HandleIo::Rewind(firstFile.get()) == S_OK, "shared handle rewind restores the beginning position");
    std::array<std::byte, bytes.size()> readBack{};
    Require(Common::HandleIo::ReadExact(firstFile.get(), readBack) == S_OK && readBack == bytes, "shared exact reader returns the entire payload after rewind");

    std::wstring stalePath = L"stale";
    wil::unique_hfile staleFile;
    const Common::Paths::UniqueSiblingFileOptions invalidOptions{.prefix = L"bad\\prefix", .suffix = L".tmp"};
    Require(Common::Paths::CreateUniqueSiblingFile(sibling, invalidOptions, stalePath, staleFile) == E_INVALIDARG && stalePath.empty() && ! staleFile,
            "failed atomic sibling validation clears unpublished path and handle outputs");
}

void TestCloudPaginationGuardBoundsProgressAndCancellation()
{
    using Common::Paging::Limits;
    using Common::Paging::Utf8ContinuationGuard;

    bool cancel                  = false;
    const auto cancellationProbe = [](void* cookie) noexcept -> HRESULT
    { return *static_cast<const bool*>(cookie) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK; };

    Limits limits{
        .maxPages           = 2u,
        .maxItems           = 3u,
        .maxBytes           = 12u,
        .maxTokenChars      = 8u,
        .deadlineTickMs     = 110u,
        .cancellationProbe  = cancellationProbe,
        .cancellationCookie = &cancel,
    };
    Utf8ContinuationGuard guard(limits);
    Require(guard.BeginFirstPage(100u) == S_OK, "pagination guard admits the bounded first page");
    Require(guard.CompletePage(2u, 5u, true, "next", 101u) == S_OK, "pagination guard records page items/bytes and a non-empty continuation");
    Require(guard.BeginContinuation("next", 102u) == S_OK, "pagination guard admits a new continuation once");
    Require(guard.CompletePage(1u, 7u, false, {}, 103u) == S_OK && guard.PageCount() == 2u && guard.ItemCount() == 3u && guard.ByteCount() == 12u,
            "pagination guard preserves exact bounded totals");

    Utf8ContinuationGuard repeated(limits);
    Require(repeated.BeginFirstPage(100u) == S_OK && repeated.CompletePage(0u, 0u, true, "same", 101u) == S_OK &&
                repeated.BeginContinuation("same", 102u) == S_OK && repeated.CompletePage(0u, 0u, true, "same", 103u) == S_OK &&
                repeated.BeginContinuation("same", 104u) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
            "pagination guard rejects a repeated continuation before another request");

    Utf8ContinuationGuard empty(limits);
    Require(empty.BeginFirstPage(100u) == S_OK && empty.CompletePage(0u, 0u, true, {}, 101u) == HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
            "pagination guard rejects an empty token when the provider says another page exists");

    Limits tiny         = limits;
    tiny.maxPages       = 1u;
    tiny.deadlineTickMs = 0u;
    Utf8ContinuationGuard pageCap(tiny);
    Require(pageCap.BeginFirstPage(100u) == S_OK && pageCap.CompletePage(0u, 0u, true, "next", 101u) == S_OK &&
                pageCap.BeginContinuation("next", 102u) == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE),
            "pagination guard enforces the page ceiling");

    Utf8ContinuationGuard deadline(limits);
    Require(deadline.BeginFirstPage(110u) == HRESULT_FROM_WIN32(ERROR_TIMEOUT), "pagination guard enforces the operation deadline");

    cancel                         = true;
    Limits cancelledLimits         = limits;
    cancelledLimits.deadlineTickMs = 0u;
    Utf8ContinuationGuard cancelled(cancelledLimits);
    Require(cancelled.BeginFirstPage(100u) == HRESULT_FROM_WIN32(ERROR_CANCELLED), "pagination guard propagates cancellation at request boundaries");
}

void TestYyjsonMemberPoliciesDistinguishPresenceTypeRangeAndCoercion()
{
    using namespace Common::Json;

    std::string json =
        R"json({"text":"value","nullValue":null,"wrong":true,"signed":-7,"unsigned":9,"tooLarge":18446744073709551615,"numericString":"42","unknownFuture":{"x":1}})json";
    UniqueDocument document{yyjson_read(json.data(), json.size(), 0)};
    Require(document != nullptr, "shared yyjson ownership accepts a valid document");
    const yyjson_val* root = yyjson_doc_get_root(document.get());

    const auto text = GetStringMember(root, "text", MemberRequirement::Required);
    Require(text.HasValue() && text.value == "value", "strict string member access returns the borrowed value with its explicit length");
    Require(GetStringMember(root, "missing", MemberRequirement::Optional).status == MemberStatus::MissingOptional,
            "optional member access distinguishes a missing optional member");
    Require(GetStringMember(root, "missing", MemberRequirement::Required).status == MemberStatus::MissingRequired,
            "required member access distinguishes a missing required member");
    Require(GetStringMember(root, "nullValue", MemberRequirement::Optional).status == MemberStatus::Null, "member access distinguishes JSON null from missing");
    Require(GetStringMember(root, "wrong", MemberRequirement::Required).status == MemberStatus::WrongType,
            "strict string member access rejects the wrong JSON type");

    const auto signedValue = GetInt64Member(root, "signed", MemberRequirement::Required);
    Require(signedValue.HasValue() && signedValue.value == -7, "strict signed member access accepts a signed integer");
    Require(GetInt64Member(root, "tooLarge", MemberRequirement::Required).status == MemberStatus::OutOfRange, "signed member access reports unsigned overflow");
    Require(GetInt64Member(root, "numericString", MemberRequirement::Required).status == MemberStatus::WrongType,
            "strict integer access rejects numeric strings");
    const auto coerced = GetInt64Member(root, "numericString", MemberRequirement::Required, NumericStringPolicy::Allow);
    Require(coerced.HasValue() && coerced.value == 42, "numeric-string coercion occurs only when requested at the call site");
    Require(GetUInt64Member(root, "signed", MemberRequirement::Required).status == MemberStatus::OutOfRange,
            "unsigned member access reports negative values as out of range");
    Require(
        GetUInt64Member(root, "unsigned", MemberRequirement::Required, NumericStringPolicy::Reject, UnsignedIntegerPolicy::RequireUnsignedStorage).HasValue(),
        "unsigned-storage member access accepts a JSON unsigned integer");
    Require(GetUInt64Member(root, "signed", MemberRequirement::Required, NumericStringPolicy::Reject, UnsignedIntegerPolicy::RequireUnsignedStorage).status ==
                MemberStatus::WrongType,
            "unsigned-storage member access rejects signed storage even before range conversion");
    Require(GetBoolMember(root, "wrong", MemberRequirement::Required).HasValue(), "strict bool member access accepts only a JSON boolean");
    Require(GetBoolMember(root, "unsigned", MemberRequirement::Required).status == MemberStatus::WrongType,
            "strict bool member access rejects integer coercion by default");
    const auto coercedBool = GetBoolMember(root, "unsigned", MemberRequirement::Required, BooleanIntegerPolicy::AllowZeroAndNonzero);
    Require(coercedBool.HasValue() && coercedBool.value, "integer-to-bool coercion occurs only when requested at the call site");
}

void TestAppThemeDxPaletteRefreshesAccentVariantsAfterAssigningAccent()
{
    const D2D1::ColorF customAccent(0.82f, 0.22f, 0.52f, 1.0f);
    const AppTheme theme                    = ResolveAppTheme(ThemeMode::Rainbow, L"dx-palette-accent-refresh", customAccent);
    const DxUi::ThemePalette palette        = MakeAppThemeDxPalette(theme);
    const DxUi::ThemePalette defaultPalette = DxUi::MakeDefaultThemePalette(theme.dark);

    RequireColorNear(palette.accent, customAccent, "app Dx palette keeps the app-assigned accent");
    RequireColorDifferent(palette.accentHover, defaultPalette.accentHover, "app Dx palette accent hover is not the stale default variant");
    RequireColorDifferent(palette.accentPressed, defaultPalette.accentPressed, "app Dx palette accent pressed is not the stale default variant");
    RequireColorDifferent(palette.headerHovered,
                          BlendForTest(palette.overlayBackground, defaultPalette.accentHover, theme.dark ? 0.44f : 0.28f),
                          "rainbow app Dx palette header hover uses the refreshed accent-hover variant");
    RequireColorDifferent(palette.headerPressed,
                          BlendForTest(palette.overlayBackground, defaultPalette.accentPressed, theme.dark ? 0.62f : 0.42f),
                          "rainbow app Dx palette header pressed uses the refreshed accent-pressed variant");
}

void TestFolderContentDxPalettePreservesItsNamedProfile()
{
    const AppTheme theme             = ResolveAppTheme(ThemeMode::Rainbow, L"folder-content-palette-golden");
    const DxUi::ThemePalette palette = MakeFolderContentDxPalette(theme);

    Require(palette.dark == theme.dark && palette.highContrast == theme.highContrast && palette.rainbowMode == theme.menu.rainbowMode,
            "folder-content palette preserves app mode flags");
    RequireColorNear(palette.accent, theme.accent, "folder-content palette preserves the resolved app accent");
    RequireColorNear(palette.surfaceBackground, theme.folderView.backgroundColor, "folder-content palette uses the folder-view surface");
    RequireColorNear(palette.inputFill, theme.folderView.backgroundColor, "folder-content palette keeps inputs on the folder-view surface");
    RequireColorNear(palette.gridLine, theme.folderView.gridLines, "folder-content palette uses folder-view grid lines");
    RequireColorNear(palette.text, theme.folderView.textNormal, "folder-content palette uses folder-view text");
    RequireColorNear(palette.headerPressed,
                     BlendForTest(palette.headerBackground, palette.accent, theme.dark ? 0.32f : 0.18f),
                     "folder-content palette preserves its stronger pressed-header profile");
}

void TestAppThemeSelectionResolutionIsSharedAndDeterministic()
{
    Require(ThemeModeFromThemeId(L"builtin/light") == ThemeMode::Light && ThemeModeFromThemeId(L"builtin/dark") == ThemeMode::Dark &&
                ThemeModeFromThemeId(L"builtin/rainbow") == ThemeMode::Rainbow && ThemeModeFromThemeId(L"builtin/highContrast") == ThemeMode::HighContrast &&
                ThemeModeFromThemeId(L"user/unknown") == ThemeMode::System,
            "theme-id mapping keeps the built-in identifiers and system fallback stable");

    Common::Settings::ThemeDefinition custom;
    custom.id          = L"user/dxui-selection-resolution";
    custom.name        = L"DxUi selection resolution";
    custom.baseThemeId = L"builtin/dark";
    custom.colors.emplace(L"app.accent", Common::Settings::ThemeColorSource(0x80402010u));
    custom.colors.emplace(L"window.background", Common::Settings::ThemeColorSource(0xFF112233u));
    custom.colors.emplace(L"navigation.background", Common::Settings::ThemeColorSource(0xFF203040u));
    custom.colors.emplace(L"navigation.separator", Common::Settings::ThemeColorSource(0xFF506070u));
    custom.colors.emplace(L"navigation.progressOk", Common::Settings::ThemeColorSource(0xFF109020u));

    const AppThemeSelectionResolution resolved = ResolveAppThemeSelection(custom.id, &custom, L"shared-theme-resolution-test");
    Require(resolved.customDefinitionResolved && resolved.resolvedColors.has_value() && resolved.baseMode == ThemeMode::Dark,
            "custom theme selection resolves once against its declared base mode");
    RequireColorNear(resolved.theme.accent,
                     D2D1::ColorF(0x40 / 255.0f, 0x20 / 255.0f, 0x10 / 255.0f, 0x80 / 255.0f),
                     "custom theme selection preserves ARGB accent channels and alpha");
    Require(resolved.theme.windowBackground == RGB(0x11, 0x22, 0x33), "custom theme selection applies the window background override");
    Require(resolved.theme.navigationView.gdiBackground == RGB(0x20, 0x30, 0x40) && resolved.theme.navigationView.gdiBorder == RGB(0x20, 0x30, 0x40) &&
                resolved.theme.navigationView.gdiBorderPen == RGB(0x50, 0x60, 0x70),
            "custom theme selection keeps navigation GDI colors synchronized with resolved D2D overrides");
    RequireColorNear(resolved.theme.fileOperations.progressTotal,
                     D2D1::ColorF(0x10 / 255.0f, 0x90 / 255.0f, 0x20 / 255.0f),
                     "custom theme selection derives file-operation colors from the effective navigation theme");

    Common::Settings::ThemeDefinition invalid = custom;
    invalid.id                                = L"user/dxui-selection-invalid";
    Common::Settings::ThemeColorSource missingReference;
    missingReference.kind         = Common::Settings::ThemeColorSourceKind::Reference;
    missingReference.references   = {L"missing.semantic.color"};
    invalid.colors[L"app.accent"] = std::move(missingReference);

    const AppThemeSelectionResolution rejected = ResolveAppThemeSelection(invalid.id, &invalid, L"shared-theme-resolution-test");
    Require(! rejected.customDefinitionResolved && ! rejected.resolvedColors.has_value() && rejected.baseMode == ThemeMode::Dark,
            "invalid custom theme selection reports failure without retaining partial resolved colors");
    const AppTheme darkBase = ResolveAppTheme(ThemeMode::Dark, L"shared-theme-resolution-test");
    RequireColorNear(rejected.theme.accent, darkBase.accent, "invalid custom theme selection falls back to the unmodified base theme");
}

void TestContiguousPostedPayloadCoalescingPreservesQueueOrderAndOperationKeys()
{
    constexpr UINT kPayloadMessage    = WM_APP + 0x370u;
    constexpr UINT kCompletionMessage = WM_APP + 0x371u;
    constexpr WPARAM kFirstOperation  = 41u;
    constexpr WPARAM kSecondOperation = 42u;

    PostedPayloadDrainStressWindowState state;
    std::atomic<uint32_t> destroyedCount{0};
    wil::unique_hwnd hwnd = CreatePostedPayloadDrainStressWindow(state);
    Require(hwnd != nullptr, "contiguous payload test window is created");

    const auto postPayload = [&](WPARAM operationKey, std::initializer_list<uint32_t> values)
    {
        auto payload            = std::make_unique<PostedPayloadDrainStressPayload>();
        payload->destroyedCount = &destroyedCount;
        payload->values.assign(values);
        Require(PostMessagePayload(hwnd.get(), kPayloadMessage, operationKey, std::move(payload)), "test payload posts successfully");
    };
    const auto removePayloadMessage = [&]()
    {
        MSG message{};
        Require(PeekMessageW(&message, hwnd.get(), kPayloadMessage, kPayloadMessage, PM_REMOVE) != 0, "expected payload message is queued");
        return message;
    };
    const auto takeQueuedPayload = [&]()
    {
        const MSG message = removePayloadMessage();
        return TakeMessagePayload<PostedPayloadDrainStressPayload>(message.lParam);
    };
    const auto appendValues = [](std::unique_ptr<PostedPayloadDrainStressPayload>& current, std::unique_ptr<PostedPayloadDrainStressPayload> newer) noexcept
    { current->values.insert(current->values.end(), newer->values.begin(), newer->values.end()); };

    postPayload(kFirstOperation, {1u});
    postPayload(kFirstOperation, {2u});
    postPayload(kSecondOperation, {100u});
    postPayload(kFirstOperation, {3u});
    MSG current             = removePayloadMessage();
    auto differentOperation = TakeAndCoalesceContiguousPostedPayloads<PostedPayloadDrainStressPayload>(
        hwnd.get(), kPayloadMessage, kFirstOperation, current.lParam, [](const PostedPayloadDrainStressPayload&, uint64_t) noexcept {
        return true;
    }, appendValues);
    Require(differentOperation.payload && differentOperation.payload->values == std::vector<uint32_t>{1u, 2u},
            "only contiguous payloads for the current operation are reduced");
    Require(differentOperation.drainedPayloadCount == 1u, "same-operation contiguous payload count is reported");
    Require(differentOperation.stoppedAtQueuedMessage && differentOperation.queuedMessage.wParam == kSecondOperation,
            "a different operation key remains at the queue head");
    differentOperation.payload.reset();
    takeQueuedPayload().reset();
    takeQueuedPayload().reset();

    postPayload(kFirstOperation, {4u});
    Require(PostMessageW(hwnd.get(), kCompletionMessage, kFirstOperation, 0) != 0, "completion marker posts successfully");
    postPayload(kFirstOperation, {5u});
    current                 = removePayloadMessage();
    auto completionBoundary = TakeAndCoalesceContiguousPostedPayloads<PostedPayloadDrainStressPayload>(
        hwnd.get(), kPayloadMessage, kFirstOperation, current.lParam, [](const PostedPayloadDrainStressPayload&, uint64_t) noexcept {
        return true;
    }, appendValues);
    Require(completionBoundary.payload && completionBoundary.payload->values == std::vector<uint32_t>{4u}, "completion behind progress is not bypassed");
    Require(completionBoundary.drainedPayloadCount == 0u && completionBoundary.stoppedAtQueuedMessage &&
                completionBoundary.queuedMessage.message == kCompletionMessage,
            "completion remains the next queue message");
    completionBoundary.payload.reset();
    MSG completion{};
    Require(PeekMessageW(&completion, nullptr, 0, 0, PM_REMOVE) != 0 && completion.message == kCompletionMessage,
            "completion marker retains its queue position");
    takeQueuedPayload().reset();

    postPayload(kFirstOperation, {6u});
    postPayload(kFirstOperation, {7u});
    postPayload(kFirstOperation, {8u});
    current       = removePayloadMessage();
    auto budgeted = TakeAndCoalesceContiguousPostedPayloads<PostedPayloadDrainStressPayload>(
        hwnd.get(), kPayloadMessage, kFirstOperation, current.lParam, [](const PostedPayloadDrainStressPayload&, uint64_t drainedPayloadCount) noexcept {
        return drainedPayloadCount < 1u;
    }, appendValues);
    Require(budgeted.payload && budgeted.payload->values == std::vector<uint32_t>{6u, 7u} && budgeted.drainedPayloadCount == 1u,
            "caller budget stops coalescing without losing the reduced payload");
    budgeted.payload.reset();
    auto finalPayload = takeQueuedPayload();
    Require(finalPayload && finalPayload->values == std::vector<uint32_t>{8u}, "the final snapshot remains queued after the budget boundary");
    finalPayload.reset();

    postPayload(kFirstOperation, {9u});
    postPayload(kFirstOperation, {10u});
    current        = removePayloadMessage();
    auto cancelled = TakeAndCoalesceContiguousPostedPayloads<PostedPayloadDrainStressPayload>(
        hwnd.get(), kPayloadMessage, kFirstOperation, current.lParam, [](const PostedPayloadDrainStressPayload&, uint64_t) noexcept {
        return false;
    }, appendValues);
    Require(cancelled.payload && cancelled.payload->values == std::vector<uint32_t>{9u} && cancelled.drainedPayloadCount == 0u,
            "caller cancellation predicate leaves later payloads untouched");
    cancelled.payload.reset();
    takeQueuedPayload().reset();

    hwnd.reset();
    Require(state.drainedCount.load(std::memory_order_acquire) == 0u, "all coalescing-test payload messages are explicitly consumed");
    Require(destroyedCount.load(std::memory_order_acquire) == 11u, "coalesced and queued payloads are each destroyed exactly once");
}

void TestSharedTestSupportPreservesSandboxAndEnvironmentPolicies()
{
    namespace TestSupport = RedSalamander::TestSupport;

    constexpr std::wstring_view kEnvironmentName{L"REDSALAMANDER_TEST_SUPPORT_SCOPE_CONTRACT"};
    static_cast<void>(SetEnvironmentVariableW(kEnvironmentName.data(), nullptr));
    {
        TestSupport::ScopedEnvironmentVariable scope(kEnvironmentName, L"temporary");
        const TestSupport::EnvironmentValue current = TestSupport::ReadEnvironmentValue(kEnvironmentName);
        Require(current.error == ERROR_SUCCESS && current.value && current.value.value() == L"temporary", "environment scope installs its temporary value");
    }
    const TestSupport::EnvironmentValue missing = TestSupport::ReadEnvironmentValue(kEnvironmentName);
    Require(missing.error == ERROR_SUCCESS && ! missing.value, "environment scope restores an originally missing value");

    Require(SetEnvironmentVariableW(kEnvironmentName.data(), L"original") != FALSE, "environment scope test installs an original value");
    {
        TestSupport::ScopedEnvironmentVariable scope(kEnvironmentName, std::nullopt);
        const TestSupport::EnvironmentValue current = TestSupport::ReadEnvironmentValue(kEnvironmentName);
        Require(current.error == ERROR_SUCCESS && ! current.value, "environment scope can temporarily remove a value");
    }
    const TestSupport::EnvironmentValue restored = TestSupport::ReadEnvironmentValue(kEnvironmentName);
    Require(restored.error == ERROR_SUCCESS && restored.value && restored.value.value() == L"original", "environment scope restores the exact original value");
    static_cast<void>(SetEnvironmentVariableW(kEnvironmentName.data(), nullptr));

    std::error_code ec;
    const std::filesystem::path outer =
        TestSupport::AcquireTestDirectory({.harnessSegment = L"dxui", .leafSegment = L"test-support-contract", .fallbackRunIdPrefix = L"dxui"}, ec);
    Require(! ec && ! outer.empty(), "test-support contract acquires an outer sandbox");

    {
        const std::wstring outerText = outer.wstring();
        TestSupport::ScopedEnvironmentVariable rootScope(TestSupport::kTestRootEnvironmentVariable, outerText);
        TestSupport::ScopedEnvironmentVariable runScope(TestSupport::kTestRunIdEnvironmentVariable, L"test-support-contract-run");

        const TestSupport::TestDirectoryOptions cleanOptions{
            .harnessSegment      = L"contract-harness",
            .leafSegment         = L"bad/leaf",
            .fallbackRunIdPrefix = L"unused",
        };
        const std::filesystem::path cleanDirectory = TestSupport::AcquireTestDirectory(cleanOptions, ec);
        Require(! ec && cleanDirectory.filename() == L"bad_leaf", "sandbox acquisition sanitizes unsafe leaf characters");
        std::filesystem::create_directories(cleanDirectory / L"sentinel", ec);
        Require(! ec, "sandbox clean-policy sentinel is created");
        const std::filesystem::path cleanedDirectory = TestSupport::AcquireTestDirectory(cleanOptions, ec);
        Require(! ec && cleanedDirectory == cleanDirectory && ! std::filesystem::exists(cleanDirectory / L"sentinel", ec),
                "clean-before-create removes prior case contents");

        TestSupport::TestDirectoryOptions retainedOptions = cleanOptions;
        retainedOptions.leafSegment                       = L"retained";
        retainedOptions.cleanExisting                     = false;
        const std::filesystem::path retainedDirectory     = TestSupport::AcquireTestDirectory(retainedOptions, ec);
        std::filesystem::create_directories(retainedDirectory / L"sentinel", ec);
        Require(! ec, "sandbox retain-policy sentinel is created");
        static_cast<void>(TestSupport::AcquireTestDirectory(retainedOptions, ec));
        Require(! ec && std::filesystem::exists(retainedDirectory / L"sentinel", ec), "no-clean acquisition preserves prior case contents");

        TestSupport::TestDirectoryOptions emptyLeafOptions = cleanOptions;
        emptyLeafOptions.leafSegment                       = L"";
        emptyLeafOptions.emptyLeafFallback                 = L"default";
        const std::filesystem::path emptyLeafDirectory     = TestSupport::AcquireTestDirectory(emptyLeafOptions, ec);
        Require(! ec && emptyLeafDirectory.filename() == L"default", "sandbox acquisition preserves the caller's empty-leaf fallback");

        TestSupport::TestDirectoryOptions traversalOptions = cleanOptions;
        traversalOptions.harnessSegment                    = L"..";
        traversalOptions.leafSegment                       = L"..";
        const std::filesystem::path traversalDirectory     = TestSupport::AcquireTestDirectory(traversalOptions, ec);
        Require(! ec && ! traversalDirectory.empty() && Common::Testing::IsSameOrDescendantTestSandboxPath(traversalDirectory, outer) &&
                    TestSupport::SanitizeTestSandboxSegment(L"..") == L".._",
                "sandbox acquisition neutralizes traversal-only harness and leaf segments");

        {
            TestSupport::ScopedEnvironmentVariable invalidRunScope(TestSupport::kTestRunIdEnvironmentVariable, L"..");
            const std::filesystem::path invalidRunDirectory = TestSupport::AcquireTestDirectory(cleanOptions, ec);
            Require(invalidRunDirectory.empty() && ec == std::errc::invalid_argument,
                    "an invalid configured run ID fails closed instead of selecting a fallback path");
        }

        const std::filesystem::path artifacts = TestSupport::AcquireTestDirectory({.harnessSegment      = L"contract-artifacts",
                                                                                   .fallbackRunIdPrefix = L"unused",
                                                                                   .kind                = TestSupport::TestDirectoryKind::Artifacts,
                                                                                   .includeLeafSegment  = false,
                                                                                   .cleanExisting       = false},
                                                                                  ec);
        Require(! ec && artifacts.filename() == L"contract-artifacts" && artifacts.parent_path().filename() == L"artifacts",
                "artifact acquisition omits the scratch leaf when requested");
    }

    {
        TestSupport::ScopedEnvironmentVariable rootScope(TestSupport::kTestRootEnvironmentVariable, L"C:\\Windows");
        TestSupport::ScopedEnvironmentVariable runScope(TestSupport::kTestRunIdEnvironmentVariable, L"test-support-contract-run");
        const std::filesystem::path unauthorized =
            TestSupport::AcquireTestDirectory({.harnessSegment = L"dxui", .leafSegment = L"unauthorized-root", .fallbackRunIdPrefix = L"unused"}, ec);
        Require(unauthorized.empty() && ec == std::errc::permission_denied, "an arbitrary absolute test root is rejected before test directories are created");
    }

    std::filesystem::remove_all(outer, ec);
    Require(! ec, "test-support contract cleans its outer sandbox");
}

void TestSharedTestSupportPumpsMessagesAndBoundsSnapshotPolling()
{
    namespace TestSupport = RedSalamander::TestSupport;
    using namespace std::chrono_literals;

    wil::unique_hwnd messageWindow(CreateWindowExW(0u, L"STATIC", L"waiting", 0u, 0, 0, 0, 0, HWND_MESSAGE, nullptr, GetModuleHandleW(nullptr), nullptr));
    Require(messageWindow != nullptr, "message-pump contract creates a message-only window");
    constexpr UINT kTestMessage = WM_APP + 73u;
    Require(PostMessageW(messageWindow.get(), kTestMessage, 0u, 0) != FALSE, "message-pump contract queues a window message");

    const TestSupport::MessagePumpWaitResult pumped = TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        MSG pending{};
        return PeekMessageW(&pending, messageWindow.get(), kTestMessage, kTestMessage, PM_NOREMOVE) == FALSE;
    },
        {.timeout = 500ms, .pollInterval = 1ms, .operationName = L"message-only window update"});
    Require(pumped.conditionMet && pumped.dispatchedMessageCount >= 1u, "message-pump wait dispatches queued UI work instead of starving it");
    Require(pumped.timeoutDiagnostic.empty(), "successful message-pump waits do not report a timeout");

    const TestSupport::MessagePumpWaitResult timedOut =
        TestSupport::PumpMessagesUntil([]() noexcept { return false; }, {.timeout = 25ms, .pollInterval = 1ms, .operationName = L"bounded timeout contract"});
    Require(! timedOut.conditionMet && timedOut.elapsed >= 20ms && timedOut.elapsed < 500ms, "message-pump timeout stays bounded near its declared budget");
    Require(timedOut.timeoutDiagnostic.find(L"bounded timeout contract") != std::wstring::npos &&
                timedOut.timeoutDiagnostic.find(L"budget 25 ms") != std::wstring::npos,
            "message-pump timeout reports the operation and declared budget");

    struct Snapshot final
    {
        uint32_t sequence = 0u;
    } lastSnapshot{};
    uint32_t sequence = 0u;
    std::wstring timeoutDiagnostic;
    const bool snapshotReady = TestSupport::WaitForSnapshot<Snapshot>(
        [&](Snapshot& snapshot) noexcept
    {
        snapshot.sequence = ++sequence;
        return true;
    },
        [](const Snapshot& snapshot) noexcept { return snapshot.sequence >= 3u; },
        {.timeout = 250ms, .pollInterval = 1ms, .operationName = L"typed snapshot contract"},
        &lastSnapshot,
        &timeoutDiagnostic);
    Require(snapshotReady && lastSnapshot.sequence == 3u, "typed snapshot polling returns the matching snapshot");
    Require(timeoutDiagnostic.empty(), "successful typed snapshot polling leaves no timeout diagnostic");

    sequence                    = 0u;
    const bool snapshotTimedOut = TestSupport::WaitForSnapshot<Snapshot>(
        [&](Snapshot& snapshot) noexcept
    {
        snapshot.sequence = ++sequence;
        return true;
    },
        [](const Snapshot&) noexcept { return false; },
        {.timeout = 10ms, .pollInterval = 1ms, .operationName = L"typed snapshot timeout"},
        &lastSnapshot,
        &timeoutDiagnostic);
    Require(! snapshotTimedOut && lastSnapshot.sequence > 0u, "typed snapshot timeout preserves the most recently observed snapshot");
    Require(timeoutDiagnostic.find(L"typed snapshot timeout") != std::wstring::npos, "typed snapshot timeout reports its operation name");
}
void TestViewerThemeAccentPressedUsesBaseThemePolarity()
{
    using namespace DxUi;
    using RedSalamander::MakeThemePaletteFromViewerTheme;

    ViewerTheme viewerTheme{.sizeBytes = sizeof(ViewerTheme)};
    viewerTheme.backgroundArgb          = 0xFFF2F2F2u;
    viewerTheme.textArgb                = 0xFF202020u;
    viewerTheme.selectionBackgroundArgb = 0xFFD8E8FFu;
    viewerTheme.selectionTextArgb       = 0xFF102030u;
    viewerTheme.accentArgb              = 0xFF808080u;
    viewerTheme.darkMode                = TRUE;
    viewerTheme.darkBase                = FALSE;

    const ThemePalette palette = MakeThemePaletteFromViewerTheme(viewerTheme);
    const auto luminance       = [](const D2D1_COLOR_F& color) noexcept { return (0.2126f * color.r) + (0.7152f * color.g) + (0.0722f * color.b); };

    Require(palette.dark, "viewer theme can request dark UI mode with a light base theme");
    Require(luminance(palette.accentPressed) < luminance(palette.accent),
            "light-base viewer theme derives a darker pressed accent even when UI dark mode is enabled");
}

void TestViewerThemeAdapterCopiesValuesAndOwnsItsRecordSize()
{
    ViewerTheme input{.sizeBytes = sizeof(ViewerTheme) + 128u};
    input.dpi                           = 144u;
    input.backgroundArgb                = 0xff030202u;
    input.textArgb                      = 0xff060403u;
    input.selectionBackgroundArgb       = 0xff090604u;
    input.selectionTextArgb             = 0xff0c0805u;
    input.accentArgb                    = 0xff0f0a06u;
    input.alertErrorBackgroundArgb      = 0xff120c07u;
    input.alertErrorTextArgb            = 0xff150e08u;
    input.alertWarningBackgroundArgb    = 0xff181009u;
    input.alertWarningTextArgb          = 0xff1b120au;
    input.alertInfoBackgroundArgb       = 0xff1e140bu;
    input.alertInfoTextArgb             = 0xff21160cu;
    input.darkMode                      = 0u;
    input.highContrast                  = 1u;
    input.rainbowMode                   = 0u;
    input.darkBase                      = 1u;
    input.diffAddedBackgroundArgb       = 0xff302011u;
    input.diffRemovedBackgroundArgb     = 0xff332212u;
    input.diffContextBackgroundArgb     = 0xff362413u;
    input.diffHeaderBackgroundArgb      = 0xff392614u;
    input.diffBannerBackgroundArgb      = 0xff3c2815u;
    input.diffPlaceholderBackgroundArgb = 0xff3f2a16u;
    input.diffDividerArgb               = 0xff422c17u;
    const auto colors                   = RedSalamander::MakeDxUiThemeColors(input);
    Require(colors.sizeBytes == sizeof(DxUi::ThemeColors), "neutral size does not inherit a plugin ABI tail");
    Require(colors.dpi == input.dpi, "viewer adapter preserves dpi");
    Require(colors.backgroundArgb == input.backgroundArgb, "viewer adapter preserves backgroundArgb");
    Require(colors.textArgb == input.textArgb, "viewer adapter preserves textArgb");
    Require(colors.selectionBackgroundArgb == input.selectionBackgroundArgb, "viewer adapter preserves selectionBackgroundArgb");
    Require(colors.selectionTextArgb == input.selectionTextArgb, "viewer adapter preserves selectionTextArgb");
    Require(colors.accentArgb == input.accentArgb, "viewer adapter preserves accentArgb");
    Require(colors.alertErrorBackgroundArgb == input.alertErrorBackgroundArgb, "viewer adapter preserves alertErrorBackgroundArgb");
    Require(colors.alertErrorTextArgb == input.alertErrorTextArgb, "viewer adapter preserves alertErrorTextArgb");
    Require(colors.alertWarningBackgroundArgb == input.alertWarningBackgroundArgb, "viewer adapter preserves alertWarningBackgroundArgb");
    Require(colors.alertWarningTextArgb == input.alertWarningTextArgb, "viewer adapter preserves alertWarningTextArgb");
    Require(colors.alertInfoBackgroundArgb == input.alertInfoBackgroundArgb, "viewer adapter preserves alertInfoBackgroundArgb");
    Require(colors.alertInfoTextArgb == input.alertInfoTextArgb, "viewer adapter preserves alertInfoTextArgb");
    Require(colors.darkMode == input.darkMode, "viewer adapter preserves darkMode");
    Require(colors.highContrast == input.highContrast, "viewer adapter preserves highContrast");
    Require(colors.rainbowMode == input.rainbowMode, "viewer adapter preserves rainbowMode");
    Require(colors.darkBase == input.darkBase, "viewer adapter preserves darkBase");
    Require(colors.diffAddedBackgroundArgb == input.diffAddedBackgroundArgb, "viewer adapter preserves diffAddedBackgroundArgb");
    Require(colors.diffRemovedBackgroundArgb == input.diffRemovedBackgroundArgb, "viewer adapter preserves diffRemovedBackgroundArgb");
    Require(colors.diffContextBackgroundArgb == input.diffContextBackgroundArgb, "viewer adapter preserves diffContextBackgroundArgb");
    Require(colors.diffHeaderBackgroundArgb == input.diffHeaderBackgroundArgb, "viewer adapter preserves diffHeaderBackgroundArgb");
    Require(colors.diffBannerBackgroundArgb == input.diffBannerBackgroundArgb, "viewer adapter preserves diffBannerBackgroundArgb");
    Require(colors.diffPlaceholderBackgroundArgb == input.diffPlaceholderBackgroundArgb, "viewer adapter preserves diffPlaceholderBackgroundArgb");
    Require(colors.diffDividerArgb == input.diffDividerArgb, "viewer adapter preserves diffDividerArgb");
}

} // namespace

int wmain()
{
    TestFolderViewIncrementalSearchKeepsContainsHighlightButUsesPrefixFocus();
    std::cout << "PASSED: TestFolderViewIncrementalSearchKeepsContainsHighlightButUsesPrefixFocus\n";
    TestFolderViewInactiveVisualStateDimsNormalTextAndIcons();
    std::cout << "PASSED: TestFolderViewInactiveVisualStateDimsNormalTextAndIcons\n";
    TestFolderViewEmptyPlaceholderMetricsUseCurrentEmptyLayout();
    std::cout << "PASSED: TestFolderViewEmptyPlaceholderMetricsUseCurrentEmptyLayout\n";
    TestSharedColorRefArgbAndTruncatingBlendGoldenValues();
    std::cout << "PASSED: TestSharedColorRefArgbAndTruncatingBlendGoldenValues\n";
    TestHsvColorRefNegativeHuePoliciesRemainExplicitAndDistinct();
    std::cout << "PASSED: TestHsvColorRefNegativeHuePoliciesRemainExplicitAndDistinct\n";
    TestSharedLuminanceMathSeparatesLinearizedAndEncodedPolicies();
    std::cout << "PASSED: TestSharedLuminanceMathSeparatesLinearizedAndEncodedPolicies\n";
    TestWindowSizingDpiRoundingAndPopupGeometryGoldenValues();
    std::cout << "PASSED: TestWindowSizingDpiRoundingAndPopupGeometryGoldenValues\n";
    TestViewerFileComboPopupInputAndHeightGoldenValues();
    std::cout << "PASSED: TestViewerFileComboPopupInputAndHeightGoldenValues\n";
    TestViewerTitleBarThemeGoldenValues();
    std::cout << "PASSED: TestViewerTitleBarThemeGoldenValues\n";
    TestUtfConversionPoliciesRemainExplicitAndPreserveEmbeddedNulls();
    std::cout << "PASSED: TestUtfConversionPoliciesRemainExplicitAndPreserveEmbeddedNulls\n";
    TestOrdinalTrimAndEnvironmentPoliciesRemainExplicit();
    std::cout << "PASSED: TestOrdinalTrimAndEnvironmentPoliciesRemainExplicit\n";
    TestWindowsPathClassesKeepDriveRelativeAndFullyAbsolutePoliciesDistinct();
    std::cout << "PASSED: TestWindowsPathClassesKeepDriveRelativeAndFullyAbsolutePoliciesDistinct\n";
    TestBinaryThroughputGrammarPreservesBoundaryPoliciesAndRoundTrips();
    std::cout << "PASSED: TestBinaryThroughputGrammarPreservesBoundaryPoliciesAndRoundTrips\n";
    TestCompactFileMetadataFormattingPreservesTimeAndAttributePolicies();
    std::cout << "PASSED: TestCompactFileMetadataFormattingPreservesTimeAndAttributePolicies\n";
    TestPercentEncodingAndHandleIoPoliciesCoverForwardProgressAndAtomicSiblings();
    std::cout << "PASSED: TestPercentEncodingAndHandleIoPoliciesCoverForwardProgressAndAtomicSiblings\n";
    TestCloudPaginationGuardBoundsProgressAndCancellation();
    std::cout << "PASSED: TestCloudPaginationGuardBoundsProgressAndCancellation\n";
    TestYyjsonMemberPoliciesDistinguishPresenceTypeRangeAndCoercion();
    std::cout << "PASSED: TestYyjsonMemberPoliciesDistinguishPresenceTypeRangeAndCoercion\n";
    TestAppThemeDxPaletteRefreshesAccentVariantsAfterAssigningAccent();
    std::cout << "PASSED: TestAppThemeDxPaletteRefreshesAccentVariantsAfterAssigningAccent\n";
    TestFolderContentDxPalettePreservesItsNamedProfile();
    std::cout << "PASSED: TestFolderContentDxPalettePreservesItsNamedProfile\n";
    TestAppThemeSelectionResolutionIsSharedAndDeterministic();
    std::cout << "PASSED: TestAppThemeSelectionResolutionIsSharedAndDeterministic\n";
    TestContiguousPostedPayloadCoalescingPreservesQueueOrderAndOperationKeys();
    std::cout << "PASSED: TestContiguousPostedPayloadCoalescingPreservesQueueOrderAndOperationKeys\n";
    TestSharedTestSupportPreservesSandboxAndEnvironmentPolicies();
    std::cout << "PASSED: TestSharedTestSupportPreservesSandboxAndEnvironmentPolicies\n";
    TestSharedTestSupportPumpsMessagesAndBoundsSnapshotPolling();
    std::cout << "PASSED: TestSharedTestSupportPumpsMessagesAndBoundsSnapshotPolling\n";
    TestViewerThemeAccentPressedUsesBaseThemePolarity();
    std::cout << "PASSED: TestViewerThemeAccentPressedUsesBaseThemePolarity\n";
    TestViewerThemeAdapterCopiesValuesAndOwnsItsRecordSize();
    std::cout << "PASSED: TestViewerThemeAdapterCopiesValuesAndOwnsItsRecordSize\n";
    return 0;
}
