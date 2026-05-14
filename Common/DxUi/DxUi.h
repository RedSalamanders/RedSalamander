#pragma once

#include <array>
#include <atomic>
#include <cstdint>
#include <functional>
#include <limits>
#include <memory>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <d2d1_1.h>
#include <d3d11.h>
#include <dcomp.h>
#include <dwrite.h>
#include <dxgi1_2.h>
#include <windows.h>

#include "PlugInterfaces/Viewer.h"

struct IRawElementProviderFragmentRoot;

namespace RedSalamander::DxUi
{
enum class SortDirection : uint8_t
{
    None,
    Ascending,
    Descending,
};

enum class GridSelectionMode : uint8_t
{
    Single,
    Extended,
};

enum class ComboBoxVariant : uint8_t
{
    Window,
    Modern,
    Edit,
};

enum class GridRowTone : uint8_t
{
    None,
    Info,
    Warning,
    Error,
};

enum class GridCellKind : uint8_t
{
    Text,
    Checkbox,
    IconText,
    ColorSwatch,
    Spinner,
    Marquee,
};

enum class GridColumnKind : uint8_t
{
    Text,
    Checkbox,
    StateImage,
};

enum class AdornmentTone : uint8_t
{
    Accent,
    Info,
    Warning,
    Error,
};

enum class OverlayMaterial : uint8_t
{
    Solid,
    Mica,
    MicaAlt,
    Acrylic,
};

enum class FontRole : uint8_t
{
    Body,
    BodyStrong, // 14/20 Semibold
    BodyLarge,  // 18/24 Regular
    Title,
    Subtitle,   // 20/28 Semibold
    TitleLarge, // 40/52 Semibold
    Display,    // 68/92 Semibold
    Header,     // DEPRECATED — use Subtitle for new code; kept for backward compat
    Small,      // maps to Caption in WinUI
    Icon,
    HeroIcon,
    Monospace,
};

// ---------------------------------------------------------------------------
// Button variant types
// ---------------------------------------------------------------------------

enum class ButtonVariant : uint8_t
{
    Standard,  // Default button
    DropDown,  // Button with chevron, opens flyout
    Split,     // Left action + right dropdown divider
    Hyperlink, // Text-only, accent color, underline on hover
    IconOnly,  // Square aspect, icon-only (no text label)
    Repeat,    // Fires continuously while held
};

enum class EasingCurve : uint8_t
{
    Linear,
    FastDecelerate,
    PointToPoint,
};

// ---------------------------------------------------------------------------
// Menu system types
// ---------------------------------------------------------------------------

enum class MenuItemKind : uint8_t
{
    Standard,  // Normal action item
    Toggle,    // Checkable on/off item
    Radio,     // Radio-exclusive item (within a group)
    Separator, // Visual divider
    Header,    // Non-interactive group label
    Info,      // Non-interactive body row with optional right-aligned detail text
    Slider,    // Discrete slider row with command-backed stops
};

struct MenuFlyoutItem
{
    MenuItemKind kind = MenuItemKind::Standard;
    std::wstring text;
    std::wstring acceleratorText;
    std::wstring iconGlyph;
    struct BitmapIcon final
    {
        BitmapIcon() = default;
        BitmapIcon(wil::unique_hbitmap sourceBitmapValue, UINT widthPxValue, UINT heightPxValue) noexcept
            : sourceBitmap(std::move(sourceBitmapValue)),
              widthPx(widthPxValue),
              heightPx(heightPxValue)
        {
        }

        BitmapIcon(const BitmapIcon&)            = delete;
        BitmapIcon& operator=(const BitmapIcon&) = delete;
        BitmapIcon(BitmapIcon&&)                 = default;
        BitmapIcon& operator=(BitmapIcon&&)      = default;

        wil::unique_hbitmap sourceBitmap;
        UINT widthPx  = 0u;
        UINT heightPx = 0u;
        mutable wil::com_ptr<ID2D1Bitmap1> cachedBitmap;
        mutable wil::com_ptr<ID2D1Device> cachedDevice;
    };
    std::shared_ptr<BitmapIcon> iconBitmap;
    struct SliderStop final
    {
        std::wstring text;
        int commandId = 0;
    };
    std::vector<SliderStop> sliderStops;
    uint32_t sliderValue = 0u;
    bool enabled  = true;
    bool checked  = false;
    int commandId = 0;
    std::vector<MenuFlyoutItem> children; // Non-empty → has submenu
};

struct MenuBarItem
{
    std::wstring text;
    wchar_t mnemonic    = L'\0';
    bool enabled        = true;
    bool rightJustified = false;
    size_t sourceIndex  = static_cast<size_t>(-1);
};

struct PointDip final
{
    float x = 0.0f;
    float y = 0.0f;

    [[nodiscard]] D2D1_POINT_2F AsD2D() const noexcept
    {
        return D2D1_POINT_2F{x, y};
    }
};

[[nodiscard]] inline PointDip MakePointDip(float x, float y) noexcept
{
    return PointDip{.x = x, .y = y};
}

[[nodiscard]] inline PointDip MakePointDip(D2D1_POINT_2F point) noexcept
{
    return MakePointDip(point.x, point.y);
}

struct SingleLineSelectionClickSequence final
{
    D2D1_POINT_2F pointDip           = D2D1::Point2F();
    uint64_t tickMs                  = 0u;
    bool promoteNextClickToSelectAll = false;
};

struct ContextMenuRootSwitchRequest
{
    POINT screenPoint{};
    std::vector<MenuFlyoutItem> items;
};

enum class ContextMenuRootHorizontalAlignment
{
    Start,
    End,
};

enum class ContextMenuRootVerticalPlacement
{
    Below,
    Above,
};

struct ContextMenuSessionCallbacks
{
    std::function<std::optional<ContextMenuRootSwitchRequest>(POINT screenPoint)> switchRootFromPointer;
    std::function<std::optional<ContextMenuRootSwitchRequest>(bool forward)> switchRootFromDirection;
    std::function<std::optional<ContextMenuRootSwitchRequest>()> switchRootFromMenuBarHover;
    ContextMenuRootHorizontalAlignment rootHorizontalAlignment = ContextMenuRootHorizontalAlignment::Start;
    ContextMenuRootVerticalPlacement rootVerticalPlacement     = ContextMenuRootVerticalPlacement::Below;
    bool focusFirstNavigableItem                               = false;
    bool ignoreInitialLeftButtonUp                             = false;
    bool ignoreInitialRightButtonUp                            = false;
};

// Shows a context menu at the given screen coordinates.
// Blocking: runs a nested message loop (like TrackPopupMenu).
// Returns the commandId of the invoked item, or nullopt if dismissed.
// (ThemePalette is defined later in this header; forward declaration suffices for reference parameter.)
struct ThemePalette;

class ContextMenu final
{
public:
    [[nodiscard]] static std::optional<int> Show(HWND ownerHwnd,
                                                 POINT screenPoint,
                                                 std::span<const MenuFlyoutItem> items,
                                                 const ThemePalette& theme,
                                                 const ContextMenuSessionCallbacks& sessionCallbacks = {});
};

#if defined(ENABLE_TESTS)
struct ContextMenuPopupDebugState
{
    bool hasScrollbar                 = false;
    bool usesSystemBackdrop           = false;
    bool usesAppBackdropBlur          = false;
    UINT dpi                          = USER_DEFAULT_SCREEN_DPI;
    float visibleWidthDip             = 0.0f;
    float visibleHeightDip            = 0.0f;
    float contentHeightDip            = 0.0f;
    float scrollOffsetDip             = 0.0f;
    D2D1_RECT_F viewportRectDip       = D2D1::RectF();
    D2D1_RECT_F scrollbarTrackRectDip = D2D1::RectF();
    D2D1_RECT_F scrollbarThumbRectDip = D2D1::RectF();
    RECT surfaceRectPx{};
    RECT windowRectPx{};
    std::optional<size_t> hoveredIndex;
    std::optional<size_t> keyboardIndex;
    std::vector<std::wstring> itemTexts;
    std::vector<MenuItemKind> itemKinds;
    std::vector<uint32_t> sliderValues;
    std::vector<uint32_t> sliderStopCounts;
    uint64_t rootPointerSwitchCount = 0;
    uint64_t rootSwitchImmediateRenderCount = 0;
    uint64_t renderCount = 0;
};

struct ContextMenuPopupItemLayoutDebugState
{
    D2D1_RECT_F itemRectDip        = D2D1::RectF();
    D2D1_RECT_F iconRectDip        = D2D1::RectF();
    D2D1_RECT_F textRectDip        = D2D1::RectF();
    D2D1_RECT_F acceleratorRectDip = D2D1::RectF();
    D2D1_RECT_F chevronRectDip     = D2D1::RectF();
    bool hasBitmapIcon             = false;
};

struct ContextMenuPopupItemPaintDebugState
{
    bool hovered                    = false;
    bool disabled                   = false;
    bool usesHighlightFill          = false;
    bool usesRainbowHighlight       = false;
    D2D1_COLOR_F fillColor          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F compositeFillColor = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F textColor          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F acceleratorColor   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F iconColor          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F checkColor         = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F chevronColor       = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct WindowHostBitmapCapture;

[[nodiscard]] bool DebugGetContextMenuPopupState(HWND hwnd, ContextMenuPopupDebugState& outState) noexcept;
[[nodiscard]] bool DebugGetContextMenuPopupItemRect(HWND hwnd, size_t itemIndex, D2D1_RECT_F& outRectDip) noexcept;
[[nodiscard]] bool DebugGetContextMenuItemDisplayText(const MenuFlyoutItem& item, std::wstring& outText);
[[nodiscard]] bool DebugGetContextMenuPopupItemText(HWND hwnd, size_t itemIndex, std::wstring& outText) noexcept;
[[nodiscard]] bool DebugGetContextMenuPopupItemLayout(HWND hwnd, size_t itemIndex, ContextMenuPopupItemLayoutDebugState& outState) noexcept;
[[nodiscard]] bool DebugGetContextMenuPopupItemPaint(HWND hwnd, size_t itemIndex, ContextMenuPopupItemPaintDebugState& outState) noexcept;
[[nodiscard]] bool DebugSetContextMenuPopupBackdropCapture(HWND hwnd, const WindowHostBitmapCapture& capture) noexcept;
[[nodiscard]] bool DebugCaptureContextMenuPopupBitmap(HWND hwnd, WindowHostBitmapCapture& outCapture) noexcept;
[[nodiscard]] bool DebugComputeContextMenuPopupPosition(POINT screenPoint,
                                                        float widthDip,
                                                        float heightDip,
                                                        UINT dpi,
                                                        bool isSubmenu,
                                                        const RECT* parentRect,
                                                        const RECT* parentItemRect,
                                                        ContextMenuRootHorizontalAlignment rootHorizontalAlignment,
                                                        RECT& outRect) noexcept;
[[nodiscard]] bool DebugComputeContextMenuPopupPosition(POINT screenPoint,
                                                        float widthDip,
                                                        float heightDip,
                                                        UINT dpi,
                                                        bool isSubmenu,
                                                        const RECT* parentRect,
                                                        const RECT* parentItemRect,
                                                        ContextMenuRootHorizontalAlignment rootHorizontalAlignment,
                                                        ContextMenuRootVerticalPlacement rootVerticalPlacement,
                                                        RECT& outRect) noexcept;
#endif

enum class InputModality : uint8_t
{
    Pointer,
    Keyboard,
};

enum class FlowDirection : uint8_t
{
    LeftToRight,
    RightToLeft,
};

enum class Density : uint8_t
{
    Standard,
    Compact,
};

struct TextInputBridgeState
{
    std::wstring text;
    std::optional<size_t> selectionAnchorIndex;
    size_t caretIndex       = 0u;
    size_t firstVisibleLine = 0u;
    bool readOnly           = false;
    bool masked             = false;
    bool multiline          = false;
};

LRESULT CALLBACK TextInputBridgeWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

struct TextFieldDebugMultilineState
{
    size_t firstVisibleLine     = 0u;
    size_t visibleLineCount     = 0u;
    size_t totalLineCount       = 0u;
    bool canScrollVertically    = false;
    bool cachedLayoutPresent    = false;
    bool layoutDirty            = true;
    float cachedLayoutWidthDip  = 0.0f;
    float cachedLayoutHeightDip = 0.0f;
};

struct TextFieldDebugSingleLinePaintState
{
    D2D1_RECT_F textRect{};
    D2D1_RECT_F selectionPaintRect{};
    float horizontalScrollDip  = 0.0f;
    bool hasSelectionPaintRect = false;
};

struct VisibleSpan
{
    uint64_t beginIndex = 0;
    uint64_t endIndex   = 0;
    float offsetDip     = 0.0f;
};

struct GridVisibleWorkMetrics
{
    uint64_t visibleRowCount         = 0;
    uint64_t visibleGroupHeaderCount = 0;
    size_t visibleColumnCount        = 0;
    uint64_t visibleCellCount        = 0;
    uint64_t visibleIconCellCount    = 0;
    uint64_t visibleBadgeCellCount   = 0;
    float verticalScrollDip          = 0.0f;
    float horizontalScrollDip        = 0.0f;
    bool hasVerticalScrollbar        = false;
    bool hasHorizontalScrollbar      = false;
};

struct GridDebugRowVisualState
{
    uint32_t fillArgb          = 0u;
    uint32_t textArgb          = 0u;
    uint32_t iconArgb          = 0u;
    uint32_t busyArgb          = 0u;
    uint32_t progressTrackArgb = 0u;
    uint32_t progressFillArgb  = 0u;
    bool usesRainbow           = false;
    bool selected              = false;
};

struct GridDebugCellVisualState
{
    uint32_t checkboxIndicatorFillArgb   = 0u;
    uint32_t checkboxIndicatorBorderArgb = 0u;
    uint32_t checkboxCheckArgb           = 0u;
    uint32_t swatchFillArgb              = 0u;
    uint32_t swatchBorderArgb            = 0u;
    uint32_t badgeFillArgb               = 0u;
    uint32_t badgeTextArgb               = 0u;
    bool hasCheckbox                     = false;
    bool hasSwatch                       = false;
    bool hasBadge                        = false;
    bool selected                        = false;
};

struct TreeDebugRowVisualState
{
    uint32_t fillArgb      = 0u;
    uint32_t textArgb      = 0u;
    uint32_t iconArgb      = 0u;
    uint32_t expanderArgb  = 0u;
    uint32_t badgeFillArgb = 0u;
    uint32_t badgeTextArgb = 0u;
    uint32_t focusArgb     = 0u;
    bool showFocus         = false;
    bool usesRainbow       = false;
    bool selected          = false;
};

struct GridSortGlyphVisualState
{
    SortDirection currentDirection  = SortDirection::None;
    float currentAlpha              = 0.0f;
    SortDirection previousDirection = SortDirection::None;
    float previousAlpha             = 0.0f;
    bool animating                  = false;
    bool reservesSpace              = false;
};

struct ScrollbarAnimationState
{
    float trackProgress = 0.0f;
    float thumbProgress = 0.0f;
    float targetTrack   = 0.0f;
    float targetThumb   = 0.0f;
    uint64_t lastTickMs = 0u;
    bool anchored       = false;
    bool active         = false;
};

struct GridScrollbarVisualState
{
    D2D1_RECT_F verticalTrackRect    = D2D1::RectF();
    D2D1_RECT_F verticalThumbRect    = D2D1::RectF();
    D2D1_RECT_F horizontalTrackRect  = D2D1::RectF();
    D2D1_RECT_F horizontalThumbRect  = D2D1::RectF();
    uint32_t verticalTrackArgb       = 0u;
    uint32_t verticalThumbArgb       = 0u;
    uint32_t horizontalTrackArgb     = 0u;
    uint32_t horizontalThumbArgb     = 0u;
    float verticalTrackHotProgress   = 0.0f;
    float verticalThumbHotProgress   = 0.0f;
    float horizontalTrackHotProgress = 0.0f;
    float horizontalThumbHotProgress = 0.0f;
    bool hasVerticalScrollbar        = false;
    bool hasHorizontalScrollbar      = false;
    bool verticalTrackHovered        = false;
    bool verticalThumbHovered        = false;
    bool horizontalTrackHovered      = false;
    bool horizontalThumbHovered      = false;
    bool verticalThumbDragging       = false;
    bool horizontalThumbDragging     = false;
};

struct TreeScrollbarVisualState
{
    D2D1_RECT_F verticalTrackRect  = D2D1::RectF();
    D2D1_RECT_F verticalThumbRect  = D2D1::RectF();
    uint32_t verticalTrackArgb     = 0u;
    uint32_t verticalThumbArgb     = 0u;
    float verticalTrackHotProgress = 0.0f;
    float verticalThumbHotProgress = 0.0f;
    bool hasVerticalScrollbar      = false;
    bool verticalTrackHovered      = false;
    bool verticalThumbHovered      = false;
    bool verticalThumbDragging     = false;
};

struct TreeItemData
{
    uint64_t id = 0u;
    std::optional<uint64_t> parentId;
    std::wstring text;
    std::wstring iconText;
    std::wstring badgeText;
    std::wstring tooltipText;
    uint32_t depth          = 0u;
    bool hasChildren        = false;
    bool expanded           = false;
    AdornmentTone badgeTone = AdornmentTone::Accent;
};

struct TreeItemLayoutMetrics
{
    bool hasExpander         = false;
    bool hasIcon             = false;
    bool hasBadge            = false;
    D2D1_RECT_F rowRect      = D2D1::RectF();
    D2D1_RECT_F expanderRect = D2D1::RectF();
    D2D1_RECT_F iconRect     = D2D1::RectF();
    D2D1_RECT_F textRect     = D2D1::RectF();
    D2D1_RECT_F badgeRect    = D2D1::RectF();
};

struct ThemePalette
{
    // ── Environment ──────────────────────────────────────────────────
    bool dark           = false;
    bool highContrast   = false;
    bool reducedMotion  = false;
    bool rainbowMode    = false;
    Density density     = Density::Standard;
    D2D1_COLOR_F accent = D2D1::ColorF(0.0f, 0.47f, 0.84f, 1.0f);

    // ── Surfaces ─────────────────────────────────────────────────────
    D2D1_COLOR_F windowBackground   = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    D2D1_COLOR_F surfaceBackground  = D2D1::ColorF(0.98f, 0.98f, 0.98f, 1.0f);
    D2D1_COLOR_F cardBackground     = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    D2D1_COLOR_F overlayBackground  = D2D1::ColorF(0.98f, 0.98f, 0.98f, 1.0f);
    D2D1_COLOR_F overlayBorder      = D2D1::ColorF(0.804f, 0.804f, 0.804f, 1.0f);
    OverlayMaterial overlayMaterial = OverlayMaterial::Solid;
    D2D1_COLOR_F smokeOverlay       = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.30f);

    // ── Header ───────────────────────────────────────────────────────
    D2D1_COLOR_F headerBackground = D2D1::ColorF(0.95f, 0.95f, 0.95f, 1.0f);
    D2D1_COLOR_F headerHovered    = D2D1::ColorF(0.92f, 0.95f, 1.0f, 1.0f);
    D2D1_COLOR_F headerPressed    = D2D1::ColorF(0.86f, 0.91f, 0.98f, 1.0f);

    // ── Chrome ───────────────────────────────────────────────────────
    D2D1_COLOR_F border        = D2D1::ColorF(0.74f, 0.74f, 0.74f, 1.0f);
    D2D1_COLOR_F gridLine      = D2D1::ColorF(0.88f, 0.88f, 0.88f, 1.0f);
    D2D1_COLOR_F borderDefault = D2D1::ColorF(0.804f, 0.804f, 0.804f, 1.0f);
    D2D1_COLOR_F borderStrong  = D2D1::ColorF(0.678f, 0.678f, 0.678f, 1.0f);

    // ── Typography ───────────────────────────────────────────────────
    D2D1_COLOR_F text         = D2D1::ColorF(0.08f, 0.08f, 0.08f, 1.0f);
    D2D1_COLOR_F subduedText  = D2D1::ColorF(0.40f, 0.40f, 0.40f, 1.0f);
    D2D1_COLOR_F disabledText = D2D1::ColorF(0.58f, 0.58f, 0.58f, 1.0f);

    // ── Selection ────────────────────────────────────────────────────
    D2D1_COLOR_F selectionFill         = D2D1::ColorF(0.0f, 0.47f, 0.84f, 1.0f);
    D2D1_COLOR_F selectionText         = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    D2D1_COLOR_F selectionInactiveFill = D2D1::ColorF(0.0f, 0.47f, 0.84f, 0.45f);
    D2D1_COLOR_F toggleKnobFill        = D2D1::ColorF(0.98f, 0.98f, 0.98f, 1.0f);
    D2D1_COLOR_F toggleKnobCheckedFill = D2D1::ColorF(0.06f, 0.06f, 0.06f, 1.0f);

    // ── Interaction states ───────────────────────────────────────────
    D2D1_COLOR_F focusStroke   = D2D1::ColorF(0.0f, 0.47f, 0.84f, 1.0f);
    D2D1_COLOR_F hoverFill     = D2D1::ColorF(0.0f, 0.47f, 0.84f, 0.10f);
    D2D1_COLOR_F pressedFill   = D2D1::ColorF(0.0f, 0.47f, 0.84f, 0.16f);
    D2D1_COLOR_F accentHover   = D2D1::ColorF(0.0f, 0.42f, 0.74f, 1.0f);
    D2D1_COLOR_F accentPressed = D2D1::ColorF(0.0f, 0.35f, 0.62f, 1.0f);

    // ── Focus ────────────────────────────────────────────────────────
    D2D1_COLOR_F focusStrokeOuter = D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f);
    D2D1_COLOR_F focusStrokeInner = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);

    // ── Button ───────────────────────────────────────────────────────
    D2D1_COLOR_F buttonFill        = D2D1::ColorF(0.98f, 0.98f, 0.98f, 1.0f);
    D2D1_COLOR_F buttonBorder      = D2D1::ColorF(0.74f, 0.74f, 0.74f, 1.0f);
    D2D1_COLOR_F buttonHotFill     = D2D1::ColorF(0.96f, 0.97f, 1.0f, 1.0f);
    D2D1_COLOR_F buttonPressedFill = D2D1::ColorF(0.92f, 0.95f, 1.0f, 1.0f);

    // ── Input ────────────────────────────────────────────────────────
    D2D1_COLOR_F inputFill   = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    D2D1_COLOR_F inputBorder = D2D1::ColorF(0.72f, 0.72f, 0.72f, 1.0f);

    // ── Scrollbar ────────────────────────────────────────────────────
    D2D1_COLOR_F scrollbarTrack    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.06f);
    D2D1_COLOR_F scrollbarThumb    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.22f);
    D2D1_COLOR_F scrollbarThumbHot = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.32f);

    // ── Tooltip ──────────────────────────────────────────────────────
    D2D1_COLOR_F tooltipBackground = D2D1::ColorF(0.10f, 0.10f, 0.10f, 0.96f);
    D2D1_COLOR_F tooltipText       = D2D1::ColorF(0.98f, 0.98f, 0.98f, 1.0f);

    // ── Status indicators ────────────────────────────────────────────
    D2D1_COLOR_F infoFill    = D2D1::ColorF(0.90f, 0.95f, 1.0f, 1.0f);
    D2D1_COLOR_F infoText    = D2D1::ColorF(0.0f, 0.39f, 0.71f, 1.0f);
    D2D1_COLOR_F warningFill = D2D1::ColorF(1.0f, 0.98f, 0.90f, 1.0f);
    D2D1_COLOR_F warningText = D2D1::ColorF(0.65f, 0.38f, 0.0f, 1.0f);
    D2D1_COLOR_F errorFill   = D2D1::ColorF(1.0f, 0.95f, 0.95f, 1.0f);
    D2D1_COLOR_F errorText   = D2D1::ColorF(0.75f, 0.0f, 0.0f, 1.0f);
};

struct ButtonVisualStyle
{
    D2D1_COLOR_F fill    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F border  = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F focus   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    bool showBorder      = false;
    bool showFocus       = false;
    float textOffsetXDip = 0.0f;
    float textOffsetYDip = 0.0f;
};

struct LabelVisualStyle
{
    D2D1_COLOR_F text = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct CheckboxVisualStyle
{
    D2D1_COLOR_F indicatorFill   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F indicatorBorder = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text            = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F check           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F focus           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F hoverFill       = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    bool showHoverFill           = false;
    bool showFocus               = false;
};

struct RadioButtonVisualStyle
{
    D2D1_COLOR_F circleBorder = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F circleFill   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F dotFill      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text         = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F hoverFill    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    float dotDiameterDip      = 8.0f;
    bool showHoverFill        = false;
    bool showFocus            = false;
};

struct ProgressBarVisualStyle
{
    D2D1_COLOR_F trackFill    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F progressFill = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct ToolbarVisualStyle
{
    D2D1_COLOR_F background    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F bottomBorder  = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F separatorLine = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct ColorSwatchVisualStyle
{
    D2D1_COLOR_F fill         = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F border       = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F swatchFill   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F swatchBorder = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F focus        = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    bool showFocus            = false;
};

struct CardPanelVisualStyle
{
    D2D1_COLOR_F fill   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F border = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct TooltipVisualStyle
{
    D2D1_COLOR_F fill   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F border = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct ToggleVisualStyle
{
    D2D1_COLOR_F rowFill     = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text        = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F trackFill   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F trackBorder = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F knobFill    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F knobBorder  = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F focus       = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    float knobDiameter       = 12.0f;
    bool showRowFill         = false;
    bool showFocus           = false;
};

struct ToggleLayoutMetrics
{
    bool compactSwitchOnly     = false;
    D2D1_RECT_F backgroundRect = D2D1::RectF();
    D2D1_RECT_F focusRect      = D2D1::RectF();
    D2D1_RECT_F textRect       = D2D1::RectF();
    D2D1_RECT_F trackRect      = D2D1::RectF();
    D2D1_RECT_F knobRect       = D2D1::RectF();
};

struct ComboBoxVisualStyle
{
    D2D1_COLOR_F fieldFill              = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F fieldBorder            = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F buttonFill             = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F buttonBorder           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F splitStroke            = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F focusAccent            = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F caret                  = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F glyph                  = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text                   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F placeholderText        = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F selectionFill          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F selectionText          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F popupFill              = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F popupBorder            = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F popupSelectedFill      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F popupActiveFill        = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F popupText              = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F popupSelectedText      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F popupActiveText        = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F popupEmptyText         = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F popupScrollbarTrack    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F popupScrollbarThumb    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F popupScrollbarThumbHot = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    bool showOuterBorder                = true;
    bool showButtonBackground           = true;
    bool showButtonSplit                = false;
    bool showLeftFocusAccent            = false;
    float cornerRadiusDip               = 5.0f;
};

struct TextFieldVisualStyle
{
    D2D1_COLOR_F fill            = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F border          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F focus           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text            = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F placeholderText = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F selectionFill   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F selectionText   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F caret           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    bool showFocus               = false;
};

struct GridSurfaceVisualStyle
{
    D2D1_COLOR_F fill            = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F border          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F headerFill      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F headerBorder    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F rowSeparator    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F columnSeparator = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F emptyText       = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct TreeSurfaceVisualStyle
{
    D2D1_COLOR_F fill      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F border    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F emptyText = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct GridHeaderVisualStyle
{
    D2D1_COLOR_F fill           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F hoveredFill    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F pressedFill    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F separator      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F titleText      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F busyGlyph      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F sortGlyph      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F groupFill      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F groupSeparator = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F groupText      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F groupGlyph     = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct GridSwatchVisualStyle
{
    D2D1_COLOR_F fill   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F border = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct GridCheckboxVisualStyle
{
    D2D1_COLOR_F indicatorFill   = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F indicatorBorder = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F check           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct GridBadgeVisualStyle
{
    D2D1_COLOR_F fill = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct TreeBadgeVisualStyle
{
    D2D1_COLOR_F fill = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct GridProgressVisualStyle
{
    D2D1_COLOR_F track = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F fill  = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

[[nodiscard]] D2D1_COLOR_F ResolveListIconColor(const ThemePalette& theme, const D2D1_COLOR_F& textColor, bool selected) noexcept;
[[nodiscard]] D2D1_COLOR_F ResolveGridBusyColor(const ThemePalette& theme, const D2D1_COLOR_F& textColor, bool selected) noexcept;
[[nodiscard]] GridProgressVisualStyle ResolveGridProgressVisualStyle(const ThemePalette& theme,
                                                                     const D2D1_COLOR_F& rowFill,
                                                                     const D2D1_COLOR_F& rowText,
                                                                     bool selected) noexcept;

struct GridSortSpec
{
    size_t columnIndex      = 0;
    SortDirection direction = SortDirection::None;
};

struct GridColumnLayoutEntry
{
    std::wstring columnId;
    size_t displayIndex = 0u;
    float widthDip      = 0.0f;
};

struct GridGroupLayoutEntry
{
    uint64_t groupStableId = 0u;
    bool collapsed         = false;
};

struct GridColumnDesc
{
    std::wstring id;
    std::wstring title;
    float widthDip                      = 120.0f;
    float minWidthDip                   = 64.0f;
    GridColumnKind kind                 = GridColumnKind::Text;
    bool sortable                       = true;
    bool multiline                      = true;
    DWRITE_TEXT_ALIGNMENT textAlignment = DWRITE_TEXT_ALIGNMENT_LEADING;
};

struct GridCellData
{
    std::wstring text;
    std::wstring iconText;
    std::wstring badgeText;
    std::wstring tooltipText;
    GridCellKind kind                   = GridCellKind::Text;
    DWRITE_TEXT_ALIGNMENT textAlignment = DWRITE_TEXT_ALIGNMENT_LEADING;
    bool multiline                      = false;
    bool checked                        = false;
    bool enabled                        = true;
    bool hasSwatchValue                 = false;
    uint32_t swatchArgb                 = 0u;
    float progress                      = 0.0f;
    AdornmentTone badgeTone             = AdornmentTone::Accent;
};

struct GridCellLayoutMetrics
{
    bool hasCheckbox         = false;
    bool hasIcon             = false;
    bool hasSwatch           = false;
    bool hasBadge            = false;
    D2D1_RECT_F cellRect     = D2D1::RectF();
    D2D1_RECT_F checkboxRect = D2D1::RectF();
    D2D1_RECT_F iconRect     = D2D1::RectF();
    D2D1_RECT_F swatchRect   = D2D1::RectF();
    D2D1_RECT_F textRect     = D2D1::RectF();
    D2D1_RECT_F badgeRect    = D2D1::RectF();
};

struct GridRowStyle
{
    GridRowTone tone = GridRowTone::None;
    std::wstring rainbowSeed;
};

struct GridGroupDesc
{
    uint64_t stableId = 0u;
    std::wstring title;
    size_t startRowIndex = 0u;
    size_t rowCount      = 0u;
    bool collapsed       = false;
};

class IDxGridModel
{
public:
    virtual ~IDxGridModel() = default;

    [[nodiscard]] virtual size_t GetRowCount() const noexcept                                  = 0;
    [[nodiscard]] virtual size_t GetColumnCount() const noexcept                               = 0;
    [[nodiscard]] virtual GridColumnDesc GetColumn(size_t columnIndex) const                   = 0;
    virtual void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const = 0;

    [[nodiscard]] virtual GridRowStyle GetRowStyle(size_t rowIndex) const;
    [[nodiscard]] virtual size_t GetGroupCount() const noexcept;
    [[nodiscard]] virtual GridGroupDesc GetGroup(size_t groupIndex) const;
    [[nodiscard]] virtual uint64_t GetStableRowId(size_t rowIndex) const noexcept;
    [[nodiscard]] virtual std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept = 0;
};

class Grid;

class IDxGridDelegate
{
public:
    virtual ~IDxGridDelegate() = default;

    virtual void OnGridSortRequested(const GridSortSpec& sortSpec);
    virtual void OnGridSelectionChanged(Grid& sender);
    virtual void OnGridSelectionChanged();
    virtual void OnGridCheckboxToggled(Grid& sender, size_t rowIndex, size_t columnIndex, bool checked);
    virtual void OnGridCheckboxToggled(size_t rowIndex, size_t columnIndex, bool checked);
    virtual void OnGridRowActivated(Grid& sender, size_t rowIndex);
    virtual void OnGridRowActivated(size_t rowIndex);
    virtual void OnGridContextMenu(Grid& sender, size_t rowIndex, POINT screenPoint);
    virtual void OnGridContextMenu(size_t rowIndex, POINT screenPoint);
    virtual void OnGridGroupToggled(Grid& sender, uint64_t groupStableId, bool collapsed);
    virtual void OnGridGroupToggled(uint64_t groupStableId, bool collapsed);
};

class IDxTreeModel
{
public:
    virtual ~IDxTreeModel() = default;

    [[nodiscard]] virtual size_t GetVisibleItemCount() const noexcept             = 0;
    virtual void GetVisibleItem(size_t visibleIndex, TreeItemData& outItem) const = 0;
    [[nodiscard]] virtual std::optional<size_t> FindVisibleItemById(uint64_t itemId) const noexcept;
};

class IDxTreeDelegate
{
public:
    virtual ~IDxTreeDelegate() = default;

    virtual void OnTreeSelectionChanged(uint64_t itemId);
    virtual void OnTreeItemInvoked(uint64_t itemId);
    virtual void OnTreeToggleExpanded(uint64_t itemId, bool expanded);
    virtual void OnTreeContextMenu(uint64_t itemId, POINT screenPoint);
};

class GridSelectionModel final
{
public:
    void Clear() noexcept;
    void SetSingle(uint64_t rowId) noexcept;
    void Toggle(uint64_t rowId) noexcept;
    void SetRange(const std::vector<uint64_t>& orderedRowIds, uint64_t anchorRowId, uint64_t currentRowId);
    void PreserveOrdered(const std::vector<uint64_t>& orderedRowIds);

    [[nodiscard]] bool IsSelected(uint64_t rowId) const noexcept;
    [[nodiscard]] std::optional<uint64_t> GetAnchor() const noexcept;
    [[nodiscard]] size_t GetCount() const noexcept;
    [[nodiscard]] std::span<const uint64_t> GetOrderedSelection() const noexcept;

private:
    std::vector<uint64_t> _selectedRowIds;
    std::optional<uint64_t> _anchorRowId;
};

class WindowHost;
class Panel;

struct PageHostDebugState
{
    bool active                         = false;
    bool hasConnectedAnimation          = false;
    uint64_t startTickMs                = 0u;
    float linearProgress                = 1.0f;
    float easedPageProgress             = 1.0f;
    float easedConnectedProgress        = 1.0f;
    float incomingOpacity               = 1.0f;
    float outgoingOpacity               = 0.0f;
    float incomingOffsetXDip            = 0.0f;
    float outgoingOffsetXDip            = 0.0f;
    D2D1_RECT_F currentConnectedRectDip = D2D1::RectF();
};

struct WindowHostBitmapCapture
{
    UINT widthPx  = 0u;
    UINT heightPx = 0u;
    std::vector<uint8_t> bgraPixels;
};

enum class WindowHostCursorKind : uint8_t
{
    Default,
    HorizontalResize,
};

class Control
{
public:
    virtual ~Control() = default;

    void SetBounds(const D2D1_RECT_F& bounds) noexcept;
    [[nodiscard]] D2D1_RECT_F GetBounds() const noexcept;
    [[nodiscard]] virtual D2D1_RECT_F GetHitBounds() const noexcept;

    void SetVisible(bool visible) noexcept;
    [[nodiscard]] bool IsVisible() const noexcept;
    void SetEnabled(bool enabled) noexcept;
    [[nodiscard]] bool IsEnabled() const noexcept;
    void SetFocusable(bool focusable) noexcept;
    [[nodiscard]] bool IsFocusable() const noexcept;
    [[nodiscard]] bool HasFocus() const noexcept;
    [[nodiscard]] bool IsHovered() const noexcept;

    virtual void Paint(WindowHost& host) const = 0;
    virtual void PaintOverlay(WindowHost& host) const;
    virtual bool Tick(WindowHost& host, uint64_t nowTickMs);

    virtual bool OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers);
    virtual bool OnMouseLeave(WindowHost& host);
    virtual bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers);
    virtual bool OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers);
    virtual bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers);
    virtual bool OnMouseWheel(WindowHost& host, D2D1_POINT_2F point, float wheelDelta, UINT modifiers);
    virtual bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers);
    virtual bool OnChar(WindowHost& host, wchar_t ch, UINT modifiers);
    virtual bool OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip);
    virtual bool OnCopy(WindowHost& host);
    virtual bool OnSelectAll(WindowHost& host);
    virtual bool OnMnemonic(WindowHost& host);
    [[nodiscard]] virtual size_t GetLogicalChildCount() const noexcept;
    [[nodiscard]] virtual Control* GetLogicalChild(size_t index) noexcept;
    [[nodiscard]] virtual const Control* GetLogicalChild(size_t index) const noexcept;

    void SetMnemonic(wchar_t mnemonic) noexcept;
    [[nodiscard]] wchar_t GetMnemonic() const noexcept;
    void SetFlowDirection(FlowDirection direction) noexcept;
    void ClearFlowDirection() noexcept;
    [[nodiscard]] bool HasExplicitFlowDirection() const noexcept;
    [[nodiscard]] FlowDirection GetFlowDirection() const noexcept;
    [[nodiscard]] bool IsRightToLeft() const noexcept;
    void SetDensity(Density density) noexcept;
    void ClearDensity() noexcept;
    [[nodiscard]] bool HasExplicitDensity() const noexcept;
    [[nodiscard]] Density GetDensity() const noexcept;
    [[nodiscard]] bool IsCompactDensity() const noexcept;
    void SetConnectedAnimationKey(std::wstring key);
    [[nodiscard]] std::wstring_view GetConnectedAnimationKey() const noexcept;
    void SetAccessibleName(std::wstring name);
    [[nodiscard]] std::wstring_view GetAccessibleName() const noexcept;
    void SetOnContextMenu(std::function<void(POINT screenPoint, bool keyboardInvocation)> onContextMenu);

protected:
    friend class Panel;
    friend class PageHost;
    friend class ScrollPanel;
    friend class WindowHost;

    [[nodiscard]] virtual Control* HitTest(D2D1_POINT_2F point);
    [[nodiscard]] virtual const Control* HitTest(D2D1_POINT_2F point) const;
    [[nodiscard]] virtual Control* HitTestOverlay(D2D1_POINT_2F point);
    [[nodiscard]] virtual const Control* HitTestOverlay(D2D1_POINT_2F point) const;
    [[nodiscard]] virtual POINT ResolveContextMenuAnchor(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip) const noexcept;
    [[nodiscard]] virtual WindowHostCursorKind ResolveCursorKind(WindowHost& host, D2D1_POINT_2F pointDip) const noexcept;

    [[nodiscard]] WindowHost* GetHost() const noexcept;
    void Invalidate(WindowHost& host) const;
    void RequestInvalidate() const noexcept;
    virtual void PropagateHost(WindowHost* host) noexcept;
    void SetParent(Panel* parent) noexcept;
    [[nodiscard]] Panel* GetParent() const noexcept;
    virtual void OnBoundsChanged() noexcept;
    virtual void OnFlowDirectionChanged() noexcept;
    virtual void OnDensityChanged() noexcept;
    virtual void OnHostDpiChanged(WindowHost& host) noexcept;
    virtual void OnFocusChanged(WindowHost& host, bool focused);
    virtual void OnHoverChanged(WindowHost& host, bool hovered);
    virtual void OnCaptureLost(WindowHost& host);
    [[nodiscard]] virtual bool SupportsTextInputBridge() const noexcept;
    [[nodiscard]] virtual std::optional<D2D1_RECT_F> GetTextInputBridgeViewportRect() const noexcept;
    [[nodiscard]] virtual std::optional<D2D1_RECT_F> GetTextInputBridgeCaretRect(const WindowHost& host, size_t controlTextIndex) const noexcept;
    virtual bool ExportTextInputBridgeState(TextInputBridgeState& outState) const;
    virtual bool ImportTextInputBridgeState(WindowHost& host, const TextInputBridgeState& state, bool notifyChange);

private:
    Panel* _parent      = nullptr;
    WindowHost* _host   = nullptr;
    D2D1_RECT_F _bounds = D2D1::RectF();
    bool _visible       = true;
    bool _enabled       = true;
    bool _focusable     = false;
    bool _hasFocus      = false;
    bool _hovered       = false;
    wchar_t _mnemonic   = L'\0';
    std::optional<FlowDirection> _explicitFlowDirection;
    std::optional<Density> _explicitDensity;
    std::wstring _connectedAnimationKey;
    std::wstring _accessibleName;
    std::function<void(POINT screenPoint, bool keyboardInvocation)> _onContextMenu;
};

class Panel : public Control
{
public:
    Panel() = default;

    template <typename TControl, typename... TArgs> TControl* AddChild(TArgs&&... args)
    {
        auto child = std::make_unique<TControl>(std::forward<TArgs>(args)...);
        auto* raw  = child.get();
        raw->SetParent(this);
        _children.push_back(std::move(child));
        return raw;
    }

    virtual void ClearChildren() noexcept;
    [[nodiscard]] size_t DebugChildCount() const noexcept
    {
        return _children.size();
    }
    [[nodiscard]] std::span<std::unique_ptr<Control>> GetChildren() noexcept;
    [[nodiscard]] std::span<const std::unique_ptr<Control>> GetChildren() const noexcept;
    [[nodiscard]] size_t GetLogicalChildCount() const noexcept override;
    [[nodiscard]] Control* GetLogicalChild(size_t index) noexcept override;
    [[nodiscard]] const Control* GetLogicalChild(size_t index) const noexcept override;

    void Paint(WindowHost& host) const override;
    void PaintOverlay(WindowHost& host) const override;
    bool Tick(WindowHost& host, uint64_t nowTickMs) override;

protected:
    [[nodiscard]] Control* HitTest(D2D1_POINT_2F point) override;
    [[nodiscard]] const Control* HitTest(D2D1_POINT_2F point) const override;
    [[nodiscard]] Control* HitTestOverlay(D2D1_POINT_2F point) override;
    [[nodiscard]] const Control* HitTestOverlay(D2D1_POINT_2F point) const override;
    void PropagateHost(WindowHost* host) noexcept override;
    void OnFlowDirectionChanged() noexcept override;
    void OnDensityChanged() noexcept override;
    void OnHostDpiChanged(WindowHost& host) noexcept override;
    [[nodiscard]] std::vector<std::unique_ptr<Control>>& AccessChildren() noexcept;
    [[nodiscard]] const std::vector<std::unique_ptr<Control>>& AccessChildren() const noexcept;

private:
    std::vector<std::unique_ptr<Control>> _children;
};

class PageHost final : public Control
{
public:
    PageHost()                           = default;
    PageHost(const PageHost&)            = delete;
    PageHost(PageHost&&)                 = delete;
    PageHost& operator=(const PageHost&) = delete;
    PageHost& operator=(PageHost&&)      = delete;

    void SetPage(std::unique_ptr<Control> page, std::wstring connectedAnimationKey = {});
    [[nodiscard]] Control* GetPage() noexcept;
    [[nodiscard]] const Control* GetPage() const noexcept;
    [[nodiscard]] bool HasActiveTransition() const noexcept;

    void Paint(WindowHost& host) const override;
    void PaintOverlay(WindowHost& host) const override;
    bool Tick(WindowHost& host, uint64_t nowTickMs) override;
    [[nodiscard]] size_t GetLogicalChildCount() const noexcept override;
    [[nodiscard]] Control* GetLogicalChild(size_t index) noexcept override;
    [[nodiscard]] const Control* GetLogicalChild(size_t index) const noexcept override;
#if defined(ENABLE_TESTS)
    [[nodiscard]] PageHostDebugState DebugGetTransitionState(uint64_t nowTickMs) const noexcept;
    void DebugFreezeTransitionProgress(float linearProgress) noexcept;
    void DebugUnfreezeTransitionProgress() noexcept;
#endif

protected:
    [[nodiscard]] Control* HitTest(D2D1_POINT_2F point) override;
    [[nodiscard]] const Control* HitTest(D2D1_POINT_2F point) const override;
    [[nodiscard]] Control* HitTestOverlay(D2D1_POINT_2F point) override;
    [[nodiscard]] const Control* HitTestOverlay(D2D1_POINT_2F point) const override;
    void PropagateHost(WindowHost* host) noexcept override;
    void OnBoundsChanged() noexcept override;
    void OnFlowDirectionChanged() noexcept override;
    void OnDensityChanged() noexcept override;
    void OnHostDpiChanged(WindowHost& host) noexcept override;

private:
    struct TransitionState final
    {
        bool active                = false;
        bool hasConnectedAnimation = false;
        uint64_t startTickMs       = 0u;
        uint64_t lastTickMs        = 0u;
        float linearProgress       = 1.0f;
        D2D1_RECT_F sourceRectDip  = D2D1::RectF();
        D2D1_RECT_F targetRectDip  = D2D1::RectF();
        // Keep test-only transition state in the layout for all builds so
        // DxUi static-library consumers cannot diverge on PageHost storage.
        bool debugFrozen          = false;
        float debugFrozenProgress = 1.0f;
    };

    void SyncChildBounds() noexcept;
    void FinishTransition() noexcept;
    [[nodiscard]] PageHostDebugState ResolveDebugState(const ThemePalette& theme, uint64_t nowTickMs) const noexcept;
    void PaintPage(WindowHost& host, const Control* page, float opacity, float offsetXDip) const;

    static constexpr uint64_t kPageTransitionDurationMs = 250u;
    static constexpr float kIncomingOffsetXDip          = 24.0f;
    static constexpr float kOutgoingOffsetXDip          = -12.0f;

    std::unique_ptr<Control> _currentPage;
    std::unique_ptr<Control> _outgoingPage;
    TransitionState _transition;
};

class CardPanel : public Panel
{
public:
    CardPanel() = default;

    void SetCornerRadius(float cornerRadiusDip) noexcept;
    [[nodiscard]] float GetCornerRadius() const noexcept;

    void Paint(WindowHost& host) const override;

private:
    float _cornerRadiusDip = 4.0f; // 4 DIP inline card, 8 DIP overlay card
};

class Label : public Control
{
public:
    explicit Label(std::wstring text = {});

    void SetText(std::wstring text);
    [[nodiscard]] std::wstring_view GetText() const noexcept;
    [[nodiscard]] Control* GetMnemonicTarget() const noexcept;
    void SetFontRole(FontRole fontRole) noexcept;
    void SetAlignment(DWRITE_TEXT_ALIGNMENT alignment) noexcept;
    void SetMultiline(bool multiline) noexcept;
    void SetTextColor(std::optional<D2D1_COLOR_F> color) noexcept;
    void SetMnemonicTarget(Control* target) noexcept;

    void Paint(WindowHost& host) const override;
    bool OnMnemonic(WindowHost& host) override;

private:
    std::wstring _text;
    FontRole _fontRole               = FontRole::Body;
    DWRITE_TEXT_ALIGNMENT _alignment = DWRITE_TEXT_ALIGNMENT_LEADING;
    bool _multiline                  = false;
    std::optional<D2D1_COLOR_F> _textColor;
    Control* _mnemonicTarget = nullptr;
};

class Button : public Control
{
public:
    explicit Button(std::wstring text = {});

    void SetText(std::wstring text);
    [[nodiscard]] std::wstring_view GetText() const noexcept;
    void SetPrimary(bool primary) noexcept;
    [[nodiscard]] bool IsPrimary() const noexcept;
    void SetVariant(ButtonVariant variant) noexcept;
    [[nodiscard]] ButtonVariant GetVariant() const noexcept;
    void SetTooltipText(std::wstring tooltipText);
    [[nodiscard]] std::wstring_view GetTooltipText() const noexcept;
    void SetOnClick(std::function<void()> onClick);
    bool Invoke(WindowHost& host, bool focusSelf);
    [[nodiscard]] float DebugGetHoverAnimationProgress() const noexcept;
    [[nodiscard]] float DebugGetFocusAnimationProgress() const noexcept;

    void Paint(WindowHost& host) const override;
    bool Tick(WindowHost& host, uint64_t nowTickMs) override;
    bool OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers) override;
    bool OnMouseLeave(WindowHost& host) override;
    bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;
    bool OnMnemonic(WindowHost& host) override;

protected:
    [[nodiscard]] bool IsPressed() const noexcept;
    void SetPressed(bool pressed) noexcept;
    void OnFocusChanged(WindowHost& host, bool focused) override;
    void OnHoverChanged(WindowHost& host, bool hovered) override;
    void OnCaptureLost(WindowHost& host) override;
    [[nodiscard]] float ResolveHoverAnimationProgress(const WindowHost& host) const noexcept;
    [[nodiscard]] float ResolveFocusAnimationProgress(const WindowHost& host) const noexcept;

private:
    struct InteractionTransitionState final
    {
        float progress      = 0.0f;
        float target        = 0.0f;
        uint64_t lastTickMs = 0u;
        bool anchored       = false;
        bool active         = false;
    };

    void UpdateInteractionTransition(WindowHost& host, InteractionTransitionState& transition, float target) noexcept;
    [[nodiscard]] bool AdvanceInteractionTransition(const WindowHost& host, InteractionTransitionState& transition, uint64_t nowTickMs) noexcept;

    static constexpr uint64_t _interactionAnimationDurationMs = 140u;
    std::wstring _text;
    std::wstring _tooltipText;
    std::function<void()> _onClick;
    InteractionTransitionState _hoverTransition{};
    InteractionTransitionState _focusTransition{};
    bool _pressed          = false;
    bool _primary          = false;
    ButtonVariant _variant = ButtonVariant::Standard;
};

class Toggle : public Button
{
public:
    explicit Toggle(std::wstring text = {});

    void SetChecked(bool checked) noexcept;
    [[nodiscard]] bool IsChecked() const noexcept;
    void SetStateLabels(std::wstring uncheckedText, std::wstring checkedText);
    [[nodiscard]] std::wstring_view GetActiveStateLabel() const noexcept;
    [[nodiscard]] std::wstring_view GetDisplayedText() const noexcept;
    void SetOnToggled(std::function<void(bool)> onToggled);
    [[nodiscard]] ToggleLayoutMetrics GetLayoutMetrics() const noexcept;

    void Paint(WindowHost& host) const override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;
    bool OnMnemonic(WindowHost& host) override;

private:
    std::wstring _uncheckedText;
    std::wstring _checkedText;
    std::function<void(bool)> _onToggled;
    bool _checked = false;
};

class Checkbox final : public Toggle
{
public:
    explicit Checkbox(std::wstring text = {});

    void SetIndeterminate(bool indeterminate) noexcept;
    [[nodiscard]] bool IsIndeterminate() const noexcept;

    void Paint(WindowHost& host) const override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;

private:
    bool _indeterminate = false;
};

class RadioButtons;

class RadioButton final : public Button
{
public:
    explicit RadioButton(std::wstring text = {});

    void SetChecked(bool checked) noexcept;
    [[nodiscard]] bool IsChecked() const noexcept;
    void SetOnSelected(std::function<void()> onSelected);
    void SetGroup(RadioButtons* group) noexcept;

    void Paint(WindowHost& host) const override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;
    bool OnMnemonic(WindowHost& host) override;

private:
    std::function<void()> _onSelected;
    RadioButtons* _group = nullptr;
    bool _checked        = false;
};

class RadioButtons final : public Panel
{
public:
    RadioButtons() = default;

    RadioButton* AddItem(std::wstring text);
    void SetSelectedIndex(int index) noexcept;
    [[nodiscard]] int GetSelectedIndex() const noexcept;
    void SetOnSelectionChanged(std::function<void(int)> onSelectionChanged);
    void SetHeader(std::wstring text);

    void Paint(WindowHost& host) const override;

private:
    friend class RadioButton;
    void SelectItem(RadioButton* item);

    std::function<void(int)> _onSelectionChanged;
    std::wstring _header;
    int _selectedIndex = -1;
};

class ProgressBar final : public Control
{
public:
    ProgressBar() = default;

    void SetValue(double value) noexcept;
    [[nodiscard]] double GetValue() const noexcept;
    void SetMinimum(double minimum) noexcept;
    [[nodiscard]] double GetMinimum() const noexcept;
    void SetMaximum(double maximum) noexcept;
    [[nodiscard]] double GetMaximum() const noexcept;
    void SetIndeterminate(bool indeterminate) noexcept;
    [[nodiscard]] bool IsIndeterminate() const noexcept;

    void Paint(WindowHost& host) const override;
    bool Tick(WindowHost& host, uint64_t nowTickMs) override;

private:
    double _value                 = 0.0;
    double _minimum               = 0.0;
    double _maximum               = 100.0;
    bool _indeterminate           = false;
    mutable float _animationPhase = 0.0f;
    uint64_t _lastTickMs          = 0;
};

enum class SliderOrientation : uint8_t
{
    Horizontal,
    Vertical,
};

class Slider final : public Control
{
public:
    Slider();

    void SetOrientation(SliderOrientation orientation) noexcept;
    [[nodiscard]] SliderOrientation GetOrientation() const noexcept;
    void SetMinimum(double minimum) noexcept;
    [[nodiscard]] double GetMinimum() const noexcept;
    void SetMaximum(double maximum) noexcept;
    [[nodiscard]] double GetMaximum() const noexcept;
    void SetValue(double value) noexcept;
    [[nodiscard]] double GetValue() const noexcept;
    void SetStep(double step) noexcept;
    [[nodiscard]] double GetStep() const noexcept;
    void SetLargeStep(double step) noexcept;
    [[nodiscard]] double GetLargeStep() const noexcept;
    void SetTickMarks(std::vector<double> tickMarks);
    [[nodiscard]] std::span<const double> GetTickMarks() const noexcept;
    void SetOnValueChanged(std::function<void(double)> onValueChanged);

    void Paint(WindowHost& host) const override;
    bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers) override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;
    void OnCaptureLost(WindowHost& host) override;

#if defined(ENABLE_TESTS)
    [[nodiscard]] D2D1_RECT_F DebugGetTrackRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F DebugGetThumbRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F DebugGetFillRect() const noexcept;
#endif

private:
    [[nodiscard]] double ClampValue(double value) const noexcept;
    [[nodiscard]] double GetNormalizedValue() const noexcept;
    void SetValueInternal(WindowHost* host, double value, bool notifyChanged) noexcept;
    [[nodiscard]] D2D1_RECT_F GetTrackRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetThumbRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetFillRect() const noexcept;
    void UpdateValueFromPoint(WindowHost& host, D2D1_POINT_2F point) noexcept;

    std::function<void(double)> _onValueChanged;
    std::vector<double> _tickMarks;
    double _minimum                  = 0.0;
    double _maximum                  = 100.0;
    double _value                    = 0.0;
    double _step                     = 1.0;
    double _largeStep                = 10.0;
    SliderOrientation _orientation   = SliderOrientation::Horizontal;
    bool _dragging                   = false;
    float _dragThumbPointerOffsetDip = 0.0f;
};

class Toolbar final : public Panel
{
public:
    Toolbar() = default;

    Button* AddButton(std::wstring tooltip, std::wstring iconGlyph);
    Toggle* AddToggleButton(std::wstring tooltip, std::wstring iconGlyph);
    void AddSeparator();

    void Paint(WindowHost& host) const override;
};

class MenuBar final : public Control
{
public:
    using OpenItemCallback = std::function<void(size_t index, POINT screenPoint, bool keyboardInvocation)>;
    using HoverChangedCallback = std::function<void(std::optional<size_t> hoveredIndex)>;

    MenuBar();

    void SetItems(std::vector<MenuBarItem> items);
    [[nodiscard]] std::span<const MenuBarItem> GetItems() const noexcept;
    void SetOnOpenItem(OpenItemCallback onOpenItem);
    void SetOnHoverChanged(HoverChangedCallback onHoverChanged);
    void SetSelectedIndex(std::optional<size_t> index) noexcept;
    void SetRetainSelectedIndexOnFocusLost(bool retain) noexcept;
    [[nodiscard]] std::optional<size_t> GetSelectedIndex() const noexcept;
    [[nodiscard]] std::optional<size_t> GetHoveredIndex() const noexcept;
    [[nodiscard]] std::optional<size_t> GetVisualHighlightIndex() const noexcept;
    [[nodiscard]] size_t GetVisualHighlightCount() const noexcept;
    [[nodiscard]] bool ActivateMnemonic(WindowHost& host, wchar_t mnemonic);
    [[nodiscard]] bool ActivateItem(WindowHost& host, size_t index, bool keyboardInvocation);
    [[nodiscard]] bool ActivateSelected(WindowHost& host, bool keyboardInvocation);
    [[nodiscard]] std::optional<size_t> HitTestPoint(const WindowHost& host, PointDip pointDip) const noexcept;
    [[nodiscard]] bool TryGetItemScreenRect(const WindowHost& host, size_t index, RECT& rectPx) const noexcept;

    void Paint(WindowHost& host) const override;
    bool OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers) override;
    bool OnMouseLeave(WindowHost& host) override;
    bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;
    bool OnMnemonic(WindowHost& host) override;

protected:
    void OnFocusChanged(WindowHost& host, bool focused) override;
    void OnCaptureLost(WindowHost& host) override;

private:
    [[nodiscard]] std::optional<size_t> HitTestItem(const WindowHost& host, PointDip pointDip) const noexcept;
    [[nodiscard]] D2D1_RECT_F GetItemRect(const WindowHost& host, size_t index) const noexcept;
    [[nodiscard]] float MeasureItemWidth(const WindowHost& host, const MenuBarItem& item) const noexcept;
    [[nodiscard]] std::optional<size_t> FindMnemonicItem(wchar_t mnemonic) const noexcept;
    void InvalidateIfInteractive(WindowHost& host) noexcept;

    std::vector<MenuBarItem> _items;
    OpenItemCallback _onOpenItem;
    HoverChangedCallback _onHoverChanged;
    std::optional<size_t> _hoveredIndex;
    std::optional<size_t> _selectedIndex;
    std::optional<size_t> _pressedIndex;
    bool _retainSelectedIndexOnFocusLost = false;
};

class TabControl final : public Panel
{
public:
    struct TabItem
    {
        std::wstring title;
        std::wstring tooltipText;
        bool closable = false;
    };

    TabControl();

    template <typename TControl, typename... TArgs> TControl* AddTab(std::wstring title, TArgs&&... args)
    {
        auto* child = AddChild<TControl>(std::forward<TArgs>(args)...);
        _tabs.push_back(TabItem{.title = std::move(title)});
        if (! _selectedIndex.has_value())
        {
            _selectedIndex = 0u;
        }
        SyncLayout();
        return child;
    }

    void RemoveTab(size_t index) noexcept;
    void SetTabTitle(size_t index, std::wstring title);
    [[nodiscard]] std::wstring_view GetTabTitle(size_t index) const noexcept;
    void SetTabTooltip(size_t index, std::wstring tooltipText);
    [[nodiscard]] std::wstring_view GetTabTooltip(size_t index) const noexcept;
    void SetTabClosable(size_t index, bool closable) noexcept;
    [[nodiscard]] bool IsTabClosable(size_t index) const noexcept;
    [[nodiscard]] size_t GetTabCount() const noexcept;
    void SetSelectedIndex(std::optional<size_t> index) noexcept;
    [[nodiscard]] std::optional<size_t> GetSelectedIndex() const noexcept;
    [[nodiscard]] Control* GetSelectedPage() noexcept;
    [[nodiscard]] const Control* GetSelectedPage() const noexcept;
    void SetOnSelectionChanged(std::function<void(size_t)> onSelectionChanged);
    void SetOnTabCloseRequested(std::function<bool(size_t)> onTabCloseRequested);
    void SetOnTabClosed(std::function<void(size_t)> onTabClosed);

    void Paint(WindowHost& host) const override;
    bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers) override;
    bool OnMouseLeave(WindowHost& host) override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseWheel(WindowHost& host, D2D1_POINT_2F point, float wheelDelta, UINT modifiers) override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;
    void OnCaptureLost(WindowHost& host) override;

#if defined(ENABLE_TESTS)
    [[nodiscard]] D2D1_RECT_F DebugGetTabRect(size_t index) const noexcept;
    [[nodiscard]] D2D1_RECT_F DebugGetCloseButtonRect(size_t index) const noexcept;
    [[nodiscard]] bool DebugIsCloseButtonVisible(size_t index) const noexcept;
    [[nodiscard]] float DebugGetHeaderScrollOffsetDip() const noexcept;
    [[nodiscard]] bool DebugHasOverflowButtons() const noexcept;
    [[nodiscard]] D2D1_RECT_F DebugGetHeaderDividerRect() const noexcept;
    [[nodiscard]] bool DebugHasHeaderDivider() const noexcept;
    [[nodiscard]] D2D1_RECT_F DebugGetBackButtonRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F DebugGetForwardButtonRect() const noexcept;
#endif

protected:
    [[nodiscard]] Control* HitTest(D2D1_POINT_2F point) override;
    [[nodiscard]] const Control* HitTest(D2D1_POINT_2F point) const override;
    void OnBoundsChanged() noexcept override;

private:
    enum class HeaderPart : uint8_t
    {
        None,
        BackButton,
        ForwardButton,
        Tab,
        CloseButton,
    };

    struct HeaderHitInfo
    {
        HeaderPart part = HeaderPart::None;
        size_t index    = 0u;
    };

    void SyncLayout() noexcept;
    void EnsureSelectedTabVisible() noexcept;
    [[nodiscard]] D2D1_RECT_F GetHeaderRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetContentRect() const noexcept;
    [[nodiscard]] bool NeedsOverflowButtons() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetHeaderDividerRect() const noexcept;
    [[nodiscard]] bool HasHeaderDividerPaintSegment() const noexcept;
    [[nodiscard]] float GetHeaderViewportLeft() const noexcept;
    [[nodiscard]] float GetHeaderViewportRight() const noexcept;
    [[nodiscard]] float MeasureTabWidthDip(size_t index) const noexcept;
    [[nodiscard]] float GetTotalTabWidthDip() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetBackButtonRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetForwardButtonRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetTabRect(size_t index) const noexcept;
    [[nodiscard]] D2D1_RECT_F GetCloseButtonRect(size_t index) const noexcept;
    [[nodiscard]] bool IsCloseButtonVisible(size_t index) const noexcept;
    [[nodiscard]] HeaderHitInfo HitTestHeader(D2D1_POINT_2F point) const noexcept;
    void ScrollHeaderBy(float deltaDip) noexcept;
    void SelectTab(WindowHost& host, size_t index, bool focusSelf) noexcept;
    void CloseTab(WindowHost& host, size_t index) noexcept;
    void ReorderTab(size_t fromIndex, size_t toIndex) noexcept;
    void UpdateDragReorder(WindowHost& host, D2D1_POINT_2F point) noexcept;
    void UpdateVisiblePageBounds() noexcept;
    void PaintHeaderDivider(WindowHost& host) const noexcept;

    std::vector<TabItem> _tabs;
    std::function<void(size_t)> _onSelectionChanged;
    std::function<bool(size_t)> _onTabCloseRequested;
    std::function<void(size_t)> _onTabClosed;
    std::optional<size_t> _selectedIndex;
    std::optional<size_t> _hoveredTabIndex;
    std::optional<size_t> _pressedTabIndex;
    std::optional<size_t> _closePressedIndex;
    std::optional<size_t> _draggingTabIndex;
    D2D1_POINT_2F _dragStartPoint = D2D1::Point2F();
    float _headerScrollOffsetDip  = 0.0f;
    bool _dragReordering          = false;
};

class ColorSwatch final : public Control
{
public:
    explicit ColorSwatch(std::optional<uint32_t> swatchArgb = std::nullopt);

    void SetSwatchValue(std::optional<uint32_t> swatchArgb) noexcept;
    [[nodiscard]] std::optional<uint32_t> GetSwatchValue() const noexcept;
    void SetOnClick(std::function<void()> onClick);

    void Paint(WindowHost& host) const override;
    bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;
    void OnCaptureLost(WindowHost& host) override;

private:
    std::optional<uint32_t> _swatchArgb;
    std::function<void()> _onClick;
    bool _pressed = false;
};

class TextField : public Control
{
public:
    explicit TextField(std::wstring text = {});

    void SetText(std::wstring text);
    void SetTextAndNotify(std::wstring text);
    [[nodiscard]] std::wstring_view GetText() const noexcept;
    void SetSelectionRange(size_t selectionStart, size_t selectionEnd) noexcept;
    [[nodiscard]] std::optional<std::pair<size_t, size_t>> GetSelectionRange() const noexcept;
    void SetMasked(bool masked) noexcept;
    [[nodiscard]] bool IsMasked() const noexcept;
    void SetPlaceholder(std::wstring text);
    void SetMultiline(bool multiline) noexcept;
    void SetClearButtonEnabled(bool enabled) noexcept;
    [[nodiscard]] bool IsClearButtonEnabled() const noexcept;
    void SetCaretColor(std::optional<D2D1_COLOR_F> caretColor) noexcept;
    void SetHorizontalTextPadding(float leftDip, float rightDip) noexcept;
    void SetVerticalTextPadding(float topDip, float bottomDip) noexcept;
    void SetReadOnly(bool readOnly) noexcept;
    [[nodiscard]] bool IsReadOnly() const noexcept;
    void SetOnTextChanged(std::function<void(std::wstring_view)> onTextChanged);
    void SetOnSubmitted(std::function<void()> onSubmitted);
    void SetOnBlur(std::function<void()> onBlur);
    [[nodiscard]] bool DebugGetMultilineState(const WindowHost& host, TextFieldDebugMultilineState& out) const noexcept;
    [[nodiscard]] bool DebugGetSingleLinePaintState(const WindowHost& host, TextFieldDebugSingleLinePaintState& out) const noexcept;

    void Paint(WindowHost& host) const override;
    bool Tick(WindowHost& host, uint64_t nowTickMs) override;
    bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers) override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseWheel(WindowHost& host, D2D1_POINT_2F point, float wheelDelta, UINT modifiers) override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;
    bool OnChar(WindowHost& host, wchar_t ch, UINT modifiers) override;
    bool OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip) override;
    bool OnCopy(WindowHost& host) override;
    bool OnSelectAll(WindowHost& host) override;

protected:
    void OnBoundsChanged() noexcept override;
    void OnFocusChanged(WindowHost& host, bool focused) override;
    void OnHostDpiChanged(WindowHost& host) noexcept override;
    [[nodiscard]] bool SupportsTextInputBridge() const noexcept override;
    [[nodiscard]] std::optional<D2D1_RECT_F> GetTextInputBridgeViewportRect() const noexcept override;
    [[nodiscard]] std::optional<D2D1_RECT_F> GetTextInputBridgeCaretRect(const WindowHost& host, size_t controlTextIndex) const noexcept override;
    bool ExportTextInputBridgeState(TextInputBridgeState& outState) const override;
    bool ImportTextInputBridgeState(WindowHost& host, const TextInputBridgeState& state, bool notifyChange) override;

private:
    struct EditHistoryState
    {
        std::wstring text;
        size_t caretIndex = 0u;
        std::optional<size_t> selectionAnchorIndex;
        size_t firstVisibleLine = 0u;
    };

    [[nodiscard]] std::wstring GetDisplayText() const;
    [[nodiscard]] D2D1_RECT_F GetTextRect() const noexcept;
    [[nodiscard]] bool IsClearButtonVisible() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetClearButtonRect() const noexcept;
    [[nodiscard]] wil::com_ptr<IDWriteTextLayout> GetOrCreateMultilineLayout(const WindowHost* host,
                                                                             std::wstring_view text,
                                                                             float widthDip,
                                                                             float heightDip) const noexcept;
    void InvalidateMultilineLayoutCache() const noexcept;
    void EnsureCaretVisible(const WindowHost* host, float availableWidthDip) const noexcept;
    void EnsureMultilineCaretVisible(const WindowHost* host) noexcept;
    void ResetCaretBlink(WindowHost& host) noexcept;
    void SetCaretIndex(size_t caretIndex, bool extendSelection) noexcept;
    [[nodiscard]] bool HasSelection() const noexcept;
    [[nodiscard]] bool DeleteSelection() noexcept;
    [[nodiscard]] EditHistoryState CaptureEditHistoryState() const;
    void RestoreEditHistoryState(const EditHistoryState& state) noexcept;
    void RecordUndoStateForDirectEdit();
    [[nodiscard]] bool TryUndoDirectEdit() noexcept;
    [[nodiscard]] bool TryRedoDirectEdit() noexcept;
    void SelectAllText() noexcept;
    void SelectWordAt(size_t hitIndex) noexcept;
    void NotifyChanged() const;

    std::wstring _text;
    std::wstring _placeholder;
    std::function<void(std::wstring_view)> _onTextChanged;
    std::function<void()> _onSubmitted;
    std::function<void()> _onBlur;
    size_t _caretIndex = 0;
    std::optional<size_t> _selectionAnchorIndex;
    std::optional<float> _preferredMultilineXOffsetDip;
    uint64_t _caretBlinkAnchorTickMs   = 0u;
    bool _caretVisible                 = true;
    mutable float _horizontalScrollDip = 0.0f;
    mutable wil::com_ptr<IDWriteTextLayout> _cachedMultilineLayout;
    mutable std::wstring _cachedLayoutText;
    mutable D2D1_SIZE_F _cachedLayoutSize = D2D1::SizeF(0.0f, 0.0f);
    mutable bool _multilineLayoutDirty    = true;
    size_t _multilineFirstVisibleLine     = 0u;
    float _multilineWheelDeltaRemainder   = 0.0f;
    bool _multiline                       = false;
    bool _readOnly                        = false;
    bool _masked                          = false;
    bool _dragSelecting                   = false;
    bool _clearButtonHovered              = false;
    bool _clearButtonEnabled              = true;
    std::optional<D2D1_COLOR_F> _caretColorOverride;
    float _textPaddingLeftDip   = 8.0f;
    float _textPaddingRightDip  = 8.0f;
    float _textPaddingTopDip    = 4.0f;
    float _textPaddingBottomDip = 4.0f;
    SingleLineSelectionClickSequence _selectionClickSequence;
    std::vector<EditHistoryState> _undoHistory;
    std::vector<EditHistoryState> _redoHistory;

    static constexpr size_t kMaxEditHistoryEntries = 64u;
};

class ComboBox : public Control
{
public:
    struct Item
    {
        std::wstring value;
        std::wstring display;
    };

    ComboBox();

    void SetVariant(ComboBoxVariant variant) noexcept;
    [[nodiscard]] ComboBoxVariant GetVariant() const noexcept;
    // Override the default kComboBoxMaxVisibleItems cap for this instance.
    // Pass 0 to restore the default.
    void SetMaxVisibleItems(size_t maxItems) noexcept;
    void SetEditable(bool editable) noexcept;
    [[nodiscard]] bool IsEditable() const noexcept;
    void SetAutoOpenOnTextInput(bool autoOpen) noexcept;
    [[nodiscard]] bool GetAutoOpenOnTextInput() const noexcept;
    void SetItems(std::vector<Item> items);
    [[nodiscard]] std::span<const Item> GetItems() const noexcept;
    void SetSelectedIndex(std::optional<size_t> selectedIndex) noexcept;
    [[nodiscard]] std::optional<size_t> GetSelectedIndex() const noexcept;
    [[nodiscard]] std::wstring_view GetSelectedValue() const noexcept;
    [[nodiscard]] std::wstring_view GetDisplayedText() const noexcept;
    void SetText(std::wstring text);
    void SetTextAndNotify(std::wstring text);
    [[nodiscard]] std::wstring_view GetText() const noexcept;
    void SetPlaceholder(std::wstring text);
    void SetOnTextChanged(std::function<void(std::wstring_view)> onTextChanged);
    void SetOnSelectionChanged(std::function<void(size_t)> onSelectionChanged);
    void SetOnSubmitted(std::function<void()> onSubmitted);
    void SetOnPopupRequested(std::function<bool()> onPopupRequested);
    [[nodiscard]] bool DebugIsPopupOpen() const noexcept;
    [[nodiscard]] std::optional<size_t> DebugGetHoveredPopupIndex() const noexcept;
    [[nodiscard]] D2D1_RECT_F DebugGetPopupBounds() const noexcept;
    [[nodiscard]] D2D1_RECT_F DebugGetPopupItemRect(size_t popupListIndex, const WindowHost* host = nullptr) const noexcept;
    [[nodiscard]] D2D1_RECT_F DebugGetPopupItemTextRect(size_t popupListIndex, const WindowHost* host = nullptr) const noexcept;
    [[nodiscard]] D2D1_RECT_F DebugGetEditableTextRect() const noexcept;

    void Paint(WindowHost& host) const override;
    void PaintOverlay(WindowHost& host) const override;
    bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers) override;
    bool OnMouseLeave(WindowHost& host) override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    void OnCaptureLost(WindowHost& host) override;
    bool OnMouseWheel(WindowHost& host, D2D1_POINT_2F point, float wheelDelta, UINT modifiers) override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;
    bool OnChar(WindowHost& host, wchar_t ch, UINT modifiers) override;
    bool OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip) override;
    bool Tick(WindowHost& host, uint64_t nowTickMs) override;
    bool OnCopy(WindowHost& host) override;
    bool OnSelectAll(WindowHost& host) override;

    [[nodiscard]] D2D1_RECT_F GetHitBounds() const noexcept override;

protected:
    void OnFocusChanged(WindowHost& host, bool focused) override;
    [[nodiscard]] Control* HitTestOverlay(D2D1_POINT_2F point) override;
    [[nodiscard]] const Control* HitTestOverlay(D2D1_POINT_2F point) const override;
    [[nodiscard]] bool SupportsTextInputBridge() const noexcept override;
    [[nodiscard]] std::optional<D2D1_RECT_F> GetTextInputBridgeViewportRect() const noexcept override;
    [[nodiscard]] std::optional<D2D1_RECT_F> GetTextInputBridgeCaretRect(const WindowHost& host, size_t controlTextIndex) const noexcept override;
    bool ExportTextInputBridgeState(TextInputBridgeState& outState) const override;
    bool ImportTextInputBridgeState(WindowHost& host, const TextInputBridgeState& state, bool notifyChange) override;

private:
    [[nodiscard]] std::optional<size_t> GetHighlightedPopupIndex() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetEditableTextRect() const noexcept;
    void EnsureEditableCaretVisible(const WindowHost* host, float availableWidthDip) const noexcept;
    void ResetEditableCaretBlink(WindowHost& host) noexcept;
    void SetEditableCaretIndex(size_t caretIndex, bool extendSelection) noexcept;
    [[nodiscard]] bool HasEditableSelection() const noexcept;
    [[nodiscard]] std::optional<std::pair<size_t, size_t>> GetEditableSelectionRange() const noexcept;
    [[nodiscard]] bool DeleteEditableSelection() noexcept;
    void SelectAllEditableText() noexcept;
    void SelectEditableWordAt(size_t hitIndex) noexcept;
    void NotifyTextChanged() const;
    void OpenPopup(WindowHost& host) noexcept;
    void ClosePopup() noexcept;
    bool RequestPopup(WindowHost& host);
    void MaybeAutoOpenPopup(WindowHost& host) noexcept;
    void CapturePopupBackdrop(WindowHost& host) noexcept;
    void RebuildPopupItems(const WindowHost* host = nullptr) noexcept;
    void CommitSelection(WindowHost& host, size_t itemIndex, bool closePopup);
    void SyncEditableSelectionFromText() noexcept;
    void EnsurePopupSelectionVisible(const WindowHost* host = nullptr) noexcept;
    void ScrollPopupBy(int deltaItems, const WindowHost* host = nullptr) noexcept;
    void ResetPopupLayout() noexcept;
    void UpdatePopupLayout(const WindowHost* host = nullptr) const noexcept;
    [[nodiscard]] float ComputePopupWidthDip(const WindowHost* host = nullptr) const noexcept;
    [[nodiscard]] bool HasPopupScrollbar() const noexcept;
    [[nodiscard]] size_t GetPopupItemCount() const noexcept;
    [[nodiscard]] size_t GetPopupVisibleItemCount() const noexcept;
    [[nodiscard]] size_t GetPopupRenderRowCount() const noexcept;
    [[nodiscard]] std::optional<size_t> GetPopupItemIndexAt(size_t popupListIndex) const noexcept;
    [[nodiscard]] std::optional<size_t> FindPopupListIndexForItem(size_t itemIndex) const noexcept;
    [[nodiscard]] std::optional<size_t> FindTypeaheadMatch(std::wstring_view prefix) const noexcept;
    [[nodiscard]] std::optional<size_t> HitTestPopupItem(D2D1_POINT_2F point) const noexcept;
    [[nodiscard]] D2D1_RECT_F GetPopupBounds() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetPopupItemRect(size_t popupListIndex, const WindowHost* host = nullptr) const noexcept;
    [[nodiscard]] D2D1_RECT_F GetPopupItemTextRect(size_t popupListIndex, const WindowHost* host = nullptr) const noexcept;
    [[nodiscard]] D2D1_RECT_F GetPopupScrollbarRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetPopupScrollbarThumbRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetPopupScrollbarThumbHitRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetDropButtonRect() const noexcept;
    void DragPopupScrollbarThumb(D2D1_POINT_2F point) noexcept;

    std::vector<Item> _items;
    std::vector<size_t> _popupItemIndices;
    std::wstring _text;
    std::wstring _placeholder;
    std::function<void(std::wstring_view)> _onTextChanged;
    std::function<void(size_t)> _onSelectionChanged;
    std::function<void()> _onSubmitted;
    std::function<bool()> _onPopupRequested;
    std::wstring _typeaheadBuffer;
    std::optional<size_t> _selectedIndex;
    std::optional<size_t> _activePopupIndex;
    std::optional<size_t> _hoveredPopupIndex;
    size_t _caretIndex = 0u;
    std::optional<size_t> _selectionAnchorIndex;
    size_t _popupScrollIndex                   = 0u;
    uint64_t _caretBlinkAnchorTickMs           = 0u;
    uint64_t _lastTypeaheadTickMs              = 0u;
    bool _caretVisible                         = true;
    mutable float _editableHorizontalScrollDip = 0.0f;
    mutable D2D1_RECT_F _popupBounds           = D2D1::RectF();
    mutable size_t _popupVisibleItemCount      = 0u;
    mutable WindowHostBitmapCapture _popupBackdropCapture{};
    mutable wil::com_ptr<ID2D1Bitmap1> _popupBackdropBitmap;
    mutable wil::com_ptr<ID2D1Device> _popupBackdropDevice;
    bool _dragPopupScrollbar    = false;
    bool _dragSelecting         = false;
    bool _popupUsesBackdropBlur = false;
    SingleLineSelectionClickSequence _selectionClickSequence;
    float _popupScrollbarDragOffsetDip = 0.0f;
    ComboBoxVariant _variant           = ComboBoxVariant::Modern;
    bool _editable                     = false;
    bool _autoOpenOnTextInput          = false;
    bool _open                         = false;
    size_t _maxVisibleItemsOverride    = 0u;
};

class StatusStrip final : public Control
{
public:
    explicit StatusStrip(std::wstring text = {});

    struct Section
    {
        std::wstring text;
        float widthDip = 0.0f; // 0 = stretch to fill remaining space
    };

    void SetText(std::wstring text);
    [[nodiscard]] std::wstring_view GetText() const noexcept;
    void SetFontRole(FontRole fontRole) noexcept;
    void SetSections(std::vector<Section> sections);
    void SetSectionText(size_t index, std::wstring text);
    [[nodiscard]] size_t GetSectionCount() const noexcept;
    [[nodiscard]] std::wstring_view GetSectionText(size_t index) const noexcept;

    void Paint(WindowHost& host) const override;

private:
    std::wstring _text;
    std::vector<Section> _sections;
    FontRole _fontRole = FontRole::Small;
};

class PopupLayer final : public Panel
{
public:
    void Paint(WindowHost& host) const override;
    void PaintOverlay(WindowHost& host) const override;

protected:
    [[nodiscard]] Control* HitTestOverlay(D2D1_POINT_2F point) override;
    [[nodiscard]] const Control* HitTestOverlay(D2D1_POINT_2F point) const override;
};

enum class StackOrientation : uint8_t
{
    Vertical,
    Horizontal,
};

struct LayoutConstraint
{
    float minExtent       = 0.0f;
    float preferredExtent = 0.0f;
    float maxExtent       = (std::numeric_limits<float>::max)();
};

[[nodiscard]] float ResolveConstrainedExtent(const LayoutConstraint& constraint, float availableExtent) noexcept;

// Layout helper that computes and applies SetBounds for children automatically.
// Children are arranged sequentially in the specified orientation. Each child gets
// the full cross-axis extent and a configurable main-axis extent (fixed height or
// auto-measured). Call ApplyLayout() after adding/removing children or when bounds change.
//
// Usage:
//   auto* stack = root->AddChild<StackPanel>();
//   stack->SetOrientation(StackOrientation::Vertical);
//   stack->SetGap(8.0f);
//   stack->SetPadding(12.0f, 12.0f, 12.0f, 12.0f);
//   auto* label = stack->AddChild<Label>(L"Title");
//   stack->SetChildExtent(label, 24.0f);  // 24 DIP tall
//   auto* edit = stack->AddChild<TextField>();
//   stack->SetChildExtent(edit, 32.0f);
//   stack->ApplyLayout();
//
class StackPanel : public Panel
{
public:
    StackPanel() = default;

    void SetOrientation(StackOrientation orientation) noexcept;
    [[nodiscard]] StackOrientation GetOrientation() const noexcept;
    void SetGap(float gapDip) noexcept;
    [[nodiscard]] float GetGap() const noexcept;
    void SetPadding(float left, float top, float right, float bottom) noexcept;

    // Set the main-axis extent for a child (height for Vertical, width for Horizontal).
    // Must be called for each child before ApplyLayout().
    void SetChildExtent(const Control* child, float extentDip);

    // Compute and apply SetBounds on all children based on current bounds, orientation,
    // gap, padding, and child extents. Call after modifying children or this panel's bounds.
    void ApplyLayout();

    // Returns the total content extent (main axis) needed to fit all children + gaps + padding.
    [[nodiscard]] float GetContentExtent() const noexcept;

private:
    void OnFlowDirectionChanged() noexcept override;
    void OnDensityChanged() noexcept override;

    StackOrientation _orientation = StackOrientation::Vertical;
    float _gapDip                 = 0.0f;
    float _padLeft                = 0.0f;
    float _padTop                 = 0.0f;
    float _padRight               = 0.0f;
    float _padBottom              = 0.0f;
    std::vector<std::pair<const Control*, float>> _childExtents;
};

// Scrollable container with a vertical scrollbar. Children are positioned in "content space"
// starting at the panel's top. When content exceeds the viewport, a scrollbar appears and the
// visible portion is clipped. ScrollPanel intercepts all mouse events and translates coordinates
// to content space before dispatching to children.
//
// Usage:
//   auto* scroll = root->AddChild<ScrollPanel>();
//   auto* stack = scroll->AddChild<StackPanel>();
//   stack->SetOrientation(StackOrientation::Vertical);
//   stack->SetGap(8.0f);
//   auto* label = stack->AddChild<Label>(L"Title");
//   stack->SetChildExtent(label, 24.0f);
//   // After layout:
//   scroll->SetBounds(viewportRect);
//   D2D1_RECT_F contentRect = { viewportRect.left, viewportRect.top,
//       viewportRect.right - scroll->GetScrollbarThickness(), viewportRect.top + 99999.0f };
//   stack->SetBounds(contentRect);
//   stack->ApplyLayout();
//   scroll->SetContentHeight(stack->GetContentExtent());
//
class ScrollPanel : public Panel
{
public:
    ScrollPanel() = default;

    void ClearChildren() noexcept override;

    void SetContentHeight(float heightDip) noexcept;
    [[nodiscard]] float GetContentHeight() const noexcept;

    [[nodiscard]] float GetScrollOffset() const noexcept;
    void SetScrollOffset(float offsetDip) noexcept;
    void ScrollToTop() noexcept;

    void SetScrollStepDip(float stepDip) noexcept;
    void SetInternalScrollbarEnabled(bool enabled) noexcept;
    void SetOnScrollChanged(std::function<void(float)> callback) noexcept;

    [[nodiscard]] bool NeedsScrollbar() const noexcept;
    [[nodiscard]] float GetScrollbarThickness() const noexcept;
    [[nodiscard]] bool IsInternalScrollbarEnabled() const noexcept;
    [[nodiscard]] bool DebugGetScrollbarThumbHitRect(D2D1_RECT_F& out) const noexcept;

    // Control overrides
    void Paint(WindowHost& host) const override;
    void PaintOverlay(WindowHost& host) const override;
    Control* HitTest(D2D1_POINT_2F point) override;
    const Control* HitTest(D2D1_POINT_2F point) const override;
    Control* HitTestOverlay(D2D1_POINT_2F point) override;
    const Control* HitTestOverlay(D2D1_POINT_2F point) const override;
    bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers) override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseWheel(WindowHost& host, D2D1_POINT_2F point, float wheelDelta, UINT modifiers) override;
    bool OnMouseLeave(WindowHost& host) override;
    void OnCaptureLost(WindowHost& host) override;

private:
    float _contentHeightDip = 0.0f;
    float _scrollOffsetDip  = 0.0f;
    float _scrollStepDip    = 40.0f;
    bool _internalScrollbarEnabled = true;

    bool _dragThumb           = false;
    float _dragThumbOffsetDip = 0.0f;

    enum class HotPart : uint8_t
    {
        None,
        Track,
        Thumb
    };
    HotPart _scrollbarHotPart = HotPart::None;

    Control* _innerHoveredChild = nullptr;
    Control* _innerCapturedChild = nullptr;

    [[nodiscard]] float GetScrollableExtent() const noexcept;
    [[nodiscard]] bool UsesInternalScrollbar() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetViewportRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetScrollbarTrackRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetScrollbarThumbRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetScrollbarThumbHitRect() const noexcept;
    void ClampScrollOffset() noexcept;
    void NotifyScrollChanged(float previousOffsetDip) noexcept;
    [[nodiscard]] D2D1_POINT_2F ToContentSpace(D2D1_POINT_2F viewportPoint) const noexcept;
    Control* FindChildAtContent(D2D1_POINT_2F contentPoint);
    Control* FindOverlayChildAtContent(D2D1_POINT_2F contentPoint);
    [[nodiscard]] const Control* FindOverlayChildAtContent(D2D1_POINT_2F contentPoint) const;
    void UpdateInnerHover(WindowHost& host, D2D1_POINT_2F viewportPoint);

    std::function<void(float)> _onScrollChanged;
};

class TooltipLayer final : public Control
{
public:
    bool SetTooltip(std::wstring text, const D2D1_POINT_2F& originDip);
    bool SetTooltipDelayed(std::wstring text, const D2D1_POINT_2F& originDip, uint64_t nowTickMs, uint64_t delayMs);
    bool BeginHideDelay(uint64_t nowTickMs, uint64_t delayMs = 100u) noexcept;
    bool CancelHideDelay() noexcept;
    bool Clear() noexcept;
    [[nodiscard]] bool HasTooltip() const noexcept;
    [[nodiscard]] std::wstring_view GetTooltipText() const noexcept;
#if defined(ENABLE_TESTS)
    [[nodiscard]] std::wstring_view DebugGetPendingTooltipText() const noexcept;
    [[nodiscard]] D2D1_RECT_F DebugGetBoundsDip(const WindowHost& host) const noexcept;
#endif

    void Paint(WindowHost& host) const override;
    bool Tick(WindowHost& host, uint64_t nowTickMs) override;
    void OnHostDpiChanged(WindowHost& host) noexcept override;

private:
    struct LayoutCache final
    {
        wil::com_ptr<IDWriteTextLayout> layout;
        D2D1_RECT_F bounds       = D2D1::RectF(0.0f, 0.0f, 0.0f, 0.0f);
        D2D1_RECT_F clientBounds = D2D1::RectF(0.0f, 0.0f, 0.0f, 0.0f);
        bool valid               = false;
    };

    void InvalidateLayoutCache() noexcept;
    [[nodiscard]] bool EnsureLayoutCache(const WindowHost& host) const noexcept;
    [[nodiscard]] D2D1_RECT_F ComputeBoundsDip(const WindowHost& host) const noexcept;
    std::wstring _text;
    std::wstring _pendingText;
    D2D1_POINT_2F _originDip = D2D1::Point2F();
    D2D1_POINT_2F _pendingOriginDip = D2D1::Point2F();
    mutable LayoutCache _layoutCache{};
    bool _showScheduled  = false;
    uint64_t _showTickMs = 0u;
    bool _hideScheduled  = false;
    uint64_t _hideTickMs = 0u;
};

class Tree final : public Control
{
public:
    Tree();

    // Non-owning model pointer. Caller manages model lifetime from SetModel() until Tree destruction.
    // Model state accessed only on UI thread — no synchronization needed.
    void SetModel(IDxTreeModel* model) noexcept;
    void SetDelegate(IDxTreeDelegate* delegate) noexcept;
    void SetRowHeightDip(float rowHeightDip) noexcept;
    void SetIndentDip(float indentDip) noexcept;
    void NotifyDataChanged();
    [[nodiscard]] IDxTreeModel* GetModel() const noexcept
    {
        return _model;
    }
    void SetSelectedItemId(std::optional<uint64_t> itemId) noexcept;
    [[nodiscard]] std::optional<uint64_t> GetSelectedItemId() const noexcept;
    bool RequestSelectVisibleItem(size_t visibleIndex) noexcept;
    bool RequestExpandedState(size_t visibleIndex, bool expanded) noexcept;
    [[nodiscard]] TreeItemLayoutMetrics GetItemLayoutMetrics(const WindowHost& host, size_t visibleIndex) const;
#if defined(ENABLE_TESTS)
    [[nodiscard]] bool DebugGetRowVisualState(const ThemePalette& theme,
                                              size_t visibleIndex,
                                              bool keyboardFocusVisible,
                                              TreeDebugRowVisualState& out) const noexcept;
    [[nodiscard]] TreeScrollbarVisualState DebugGetScrollbarVisualState(const ThemePalette& theme) const noexcept;
    [[nodiscard]] float DebugGetVerticalScrollDip() const noexcept;
    [[nodiscard]] size_t DebugGetFirstVisibleIndex() const noexcept;
    [[nodiscard]] std::optional<size_t> DebugGetSelectedVisibleIndex() const noexcept;
#endif

    void Paint(WindowHost& host) const override;
    bool Tick(WindowHost& host, uint64_t nowTickMs) override;
    bool OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers) override;
    bool OnMouseLeave(WindowHost& host) override;
    bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    void OnCaptureLost(WindowHost& host) override;
    bool OnMouseWheel(WindowHost& host, D2D1_POINT_2F point, float wheelDelta, UINT modifiers) override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;
    bool OnChar(WindowHost& host, wchar_t ch, UINT modifiers) override;
    bool OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip) override;
    void OnDensityChanged() noexcept override;

private:
    enum class HitZone : uint8_t
    {
        None,
        Item,
        Expander,
        VerticalScrollbar,
    };

    struct HitInfo
    {
        HitZone zone          = HitZone::None;
        size_t visibleIndex   = 0u;
        D2D1_RECT_F rectDip   = D2D1::RectF();
        bool onScrollbarThumb = false;
    };

    enum class ScrollbarHotPart : uint8_t
    {
        None,
        Track,
        Thumb,
    };

    [[nodiscard]] HitInfo HitTestPoint(PointDip pointDip) const noexcept;
    [[nodiscard]] float GetVerticalScrollableExtent() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetContentRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetVerticalScrollbarRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetVerticalThumbRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetVerticalThumbHitRect() const noexcept;
    [[nodiscard]] TreeItemLayoutMetrics ComputeItemLayoutMetrics(const WindowHost& host, size_t visibleIndex, const TreeItemData& item) const noexcept;
    [[nodiscard]] TreeItemLayoutMetrics ComputeItemLayoutMetrics(const WindowHost& host, float rowTopDip, const TreeItemData& item) const noexcept;
    void ClampScrollOffset() noexcept;
    void UpdateScrollbarHotState(const HitInfo& hit) noexcept;
    void SyncScrollbarAnimation(WindowHost& host) noexcept;
    [[nodiscard]] std::optional<size_t> FindSelectedVisibleIndex() const noexcept;
    void EnsureVisibleIndex(size_t visibleIndex) noexcept;
    void SelectVisibleIndex(size_t visibleIndex, bool notifyDelegate);
    void ToggleExpanded(size_t visibleIndex);
    [[nodiscard]] float ComputeExpanderProgress(uint64_t itemId, bool expanded, uint64_t nowTickMs) const noexcept;
    void StartExpanderAnimation(uint64_t itemId, bool fromExpanded, bool toExpanded) noexcept;
    [[nodiscard]] std::optional<size_t> FindNextTypeaheadMatch(std::wstring_view prefix) const noexcept;
    [[nodiscard]] float GetExpanderProgress(uint64_t itemId, bool expanded, uint64_t nowTickMs) const noexcept;
    [[nodiscard]] bool HasActiveExpanderAnimation(uint64_t itemId, bool expanded, uint64_t nowTickMs) const noexcept;
    [[nodiscard]] std::vector<TreeItemData> CaptureVisibleItems() const;
    void BeginTreeExpansionAnimation(
        uint64_t itemId, bool toExpanded, std::vector<TreeItemData>&& beforeItems, std::vector<TreeItemData>&& afterItems, uint64_t nowTickMs) noexcept;
    void ClearTreeExpansionAnimation() noexcept;
    [[nodiscard]] bool HasActiveTreeExpansionAnimation(uint64_t nowTickMs) const noexcept;
    [[nodiscard]] float GetTreeExpansionProgress(uint64_t nowTickMs) const noexcept;

    // Non-owning. Caller manages model lifetime. Valid from SetModel() until Tree destruction.
    // Invalidation validated at message entry by PruneStaleInteractionState().
    IDxTreeModel* _model       = nullptr;
    IDxTreeDelegate* _delegate = nullptr;
    std::optional<uint64_t> _selectedItemId;
    std::optional<size_t> _hoveredVisibleIndex;
    float _rowHeightBaseDip    = 28.0f;
    float _rowHeightDip        = 28.0f;
    float _indentDip           = 16.0f;
    float _verticalScrollDip   = 0.0f;
    float _wheelDeltaRemainder = 0.0f;
    float _dragThumbOffsetDip  = 0.0f;
    std::wstring _typeaheadBuffer;
    uint64_t _lastTypeaheadTickMs                               = 0u;
    static constexpr uint64_t _treeExpanderAnimationDurationMs  = 240u;
    static constexpr uint64_t _treeExpansionAnimationDurationMs = 320u;
    struct ExpanderAnimationState final
    {
        uint64_t itemId      = 0u;
        bool fromExpanded    = false;
        bool toExpanded      = false;
        float fromProgress   = 0.0f;
        float toProgress     = 1.0f;
        uint64_t startTickMs = 0u;
        bool active          = false;
    };
    struct TreeExpansionAnimationState final
    {
        uint64_t itemId      = 0u;
        bool toExpanded      = false;
        bool active          = false;
        uint64_t startTickMs = 0u;
        std::vector<TreeItemData> beforeItems;
        std::vector<TreeItemData> afterItems;
    };
    std::vector<ExpanderAnimationState> _expanderAnimations{};
    std::optional<TreeExpansionAnimationState> _treeExpansionAnimation;
    ScrollbarHotPart _verticalScrollbarHotPart = ScrollbarHotPart::None;
    ScrollbarAnimationState _verticalScrollbarAnimation{};
    bool _dragVerticalThumb = false;
};

class Grid final : public Control
{
public:
    Grid();

    // Non-owning model pointer. Caller manages model lifetime from SetModel() until Grid destruction.
    // Model state accessed only on UI thread — no synchronization needed.
    void SetModel(IDxGridModel* model) noexcept;
    void SetDelegate(IDxGridDelegate* delegate) noexcept;
    void SetSelectionMode(GridSelectionMode mode) noexcept;
    [[nodiscard]] GridSelectionMode GetSelectionMode() const noexcept
    {
        return _selectionMode;
    }
    void SetRowHeightDip(float rowHeightDip) noexcept;
    void SetHeaderHeightDip(float headerHeightDip) noexcept;
    void SetLineClamp(uint32_t lineClamp) noexcept;
    void SetEmptyStateText(std::wstring text);
    [[nodiscard]] std::wstring_view GetEmptyStateText() const noexcept;
    void SetHeaderBusy(bool busy) noexcept;
    void SetHeaderBusyColumn(std::optional<size_t> columnIndex) noexcept;
    void SetSortSpec(const GridSortSpec& sortSpec) noexcept;
    [[nodiscard]] GridSortSpec GetSortSpec() const noexcept;
    void ApplyColumnLayout(std::span<const GridColumnLayoutEntry> layout) noexcept;
    [[nodiscard]] std::vector<GridColumnLayoutEntry> CaptureColumnLayout() const;
    void ApplyGroupLayout(std::span<const GridGroupLayoutEntry> layout) noexcept;
    [[nodiscard]] std::vector<GridGroupLayoutEntry> CaptureGroupLayout() const;
    void NotifyDataChanged();
    [[nodiscard]] IDxGridModel* GetModel() const noexcept
    {
        return _model;
    }

    [[nodiscard]] GridSelectionModel& GetSelectionModel() noexcept;
    [[nodiscard]] const GridSelectionModel& GetSelectionModel() const noexcept;
    [[nodiscard]] GridVisibleWorkMetrics GetVisibleWorkMetrics() const;
    [[nodiscard]] GridCellLayoutMetrics GetCellLayoutMetrics(const WindowHost& host, size_t rowIndex, size_t columnIndex) const;
    [[nodiscard]] size_t GetVisibleRowCount() const;
    [[nodiscard]] std::optional<size_t> GetVisibleRowAt(size_t visibleRowIndex) const;
    [[nodiscard]] std::optional<size_t> FindVisibleRowOrdinal(size_t rowIndex) const;
    void EnsureRowVisible(size_t rowIndex) noexcept;
    [[nodiscard]] size_t GetVisibleColumnCount() const;
    [[nodiscard]] std::optional<size_t> GetVisibleColumnAt(size_t visibleColumnIndex) const;
    [[nodiscard]] std::optional<size_t> FindVisibleColumnOrdinal(size_t columnIndex) const;
    [[nodiscard]] std::optional<size_t> FindHeaderColumnAtPoint(PointDip pointDip) const noexcept;
    [[nodiscard]] std::optional<size_t> FindRowAtPoint(PointDip pointDip) const noexcept;
    [[nodiscard]] std::optional<std::pair<size_t, size_t>> FindCellAtPoint(PointDip pointDip) const noexcept;
    [[nodiscard]] std::optional<D2D1_RECT_F> GetVisibleColumnHeaderRect(size_t columnIndex) const;
    [[nodiscard]] std::optional<D2D1_RECT_F> GetVisibleDisplayColumnHeaderRect(size_t displayIndex) const;
    [[nodiscard]] std::optional<D2D1_RECT_F> GetVisibleRowRect(size_t rowIndex) const;
    [[nodiscard]] std::optional<D2D1_RECT_F> GetVisibleCellRect(size_t rowIndex, size_t columnIndex) const;
    [[nodiscard]] bool IsRowSelected(size_t rowIndex) const noexcept;
    [[nodiscard]] std::optional<size_t> GetPrimarySelectedRow() const noexcept;
#if defined(ENABLE_TESTS)
    struct GridDebugHitInfo
    {
        uint32_t zone         = 0u;
        size_t rowIndex       = 0u;
        size_t groupIndex     = 0u;
        size_t columnIndex    = 0u;
        D2D1_RECT_F rectDip   = D2D1::RectF();
        bool onScrollbarThumb = false;
        bool isHeaderResize   = false;
    };
    struct GridDebugPointerState
    {
        uint64_t headerResizeDownCount = 0u;
        uint64_t resizeMoveCount       = 0u;
        bool resizeActive              = false;
        float lastResizeDeltaDip       = 0.0f;
        float lastResizeWidthDip       = 0.0f;
        bool pressedHeaderActive       = false;
        size_t pressedHeaderColumn     = 0u;
        bool reorderActive             = false;
        size_t reorderColumn           = 0u;
        size_t reorderTargetDisplayIndex = 0u;
        uint64_t headerReorderStartCount = 0u;
        uint64_t headerReorderCommitCount = 0u;
        uint64_t headerReorderNoOpCount = 0u;
        size_t lastHeaderReorderColumn = 0u;
        size_t lastHeaderReorderFromDisplayIndex = 0u;
        size_t lastHeaderReorderRawTargetDisplayIndex = 0u;
        size_t lastHeaderReorderNormalizedTargetDisplayIndex = 0u;
    };

    [[nodiscard]] bool DebugGetRowVisualState(const ThemePalette& theme, size_t rowIndex, GridDebugRowVisualState& out) const noexcept;
    [[nodiscard]] bool DebugGetCellVisualState(const ThemePalette& theme, size_t rowIndex, size_t columnIndex, GridDebugCellVisualState& out) const noexcept;
    [[nodiscard]] GridSortGlyphVisualState DebugGetSortGlyphVisualState(const ThemePalette& theme, size_t columnIndex, uint64_t nowTickMs) const noexcept;
    [[nodiscard]] GridScrollbarVisualState DebugGetScrollbarVisualState(const ThemePalette& theme) const noexcept;
    [[nodiscard]] uint64_t DebugGetPaintCount() const noexcept;
    [[nodiscard]] bool DebugHitTestPoint(PointDip pointDip, GridDebugHitInfo& out) const noexcept;
    [[nodiscard]] GridDebugPointerState DebugGetPointerState() const noexcept;
    void DebugSetScrollOffsets(float verticalScrollDip, float horizontalScrollDip) noexcept;
#endif
    bool RequestSelectRow(size_t rowIndex, UINT modifiers);
    bool RequestRemoveRowSelection(size_t rowIndex);
    bool RequestToggleCheckboxCell(WindowHost& host, size_t rowIndex, size_t columnIndex);

    void Paint(WindowHost& host) const override;
    bool Tick(WindowHost& host, uint64_t nowTickMs) override;
    bool OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers) override;
    bool OnMouseLeave(WindowHost& host) override;
    bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    bool OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override;
    void OnCaptureLost(WindowHost& host) override;
    bool OnMouseWheel(WindowHost& host, D2D1_POINT_2F point, float wheelDelta, UINT modifiers) override;
    bool OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers) override;
    bool OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip) override;
    bool OnCopy(WindowHost& host) override;
    bool OnSelectAll(WindowHost& host) override;
    [[nodiscard]] WindowHostCursorKind ResolveCursorKind(WindowHost& host, D2D1_POINT_2F pointDip) const noexcept override;
    void OnDensityChanged() noexcept override;

private:
    enum class HitZone : uint8_t
    {
        None,
        Header,
        HeaderResize,
        GroupHeader,
        Cell,
        VerticalScrollbar,
        HorizontalScrollbar,
    };

    struct HitInfo
    {
        HitZone zone          = HitZone::None;
        size_t rowIndex       = 0;
        size_t groupIndex     = 0;
        size_t columnIndex    = 0;
        D2D1_RECT_F rectDip   = D2D1::RectF();
        bool onScrollbarThumb = false;
    };

    enum class ScrollbarHotPart : uint8_t
    {
        None,
        Track,
        Thumb,
    };

    struct VisibleBodyItem
    {
        enum class Kind : uint8_t
        {
            GroupHeader,
            Row,
        };

        Kind kind           = Kind::Row;
        size_t rowIndex     = 0u;
        size_t groupIndex   = 0u;
        D2D1_RECT_F rectDip = D2D1::RectF();
    };

    struct VisibleColumnSpan
    {
        size_t beginIndex = 0;
        size_t endIndex   = 0;
        float beginXDip   = 0.0f;
    };

    [[nodiscard]] HitInfo HitTestPoint(PointDip pointDip) const noexcept;
    void ClampScrollOffsets(bool normalizeVertical = true) noexcept;
    void EnsureColumnWidths() const;
    void SelectRow(size_t rowIndex, UINT modifiers);
    [[nodiscard]] std::optional<size_t> ResolveCheckboxToggleColumn(size_t rowIndex) const;
    [[nodiscard]] bool ToggleCheckboxCell(WindowHost& host, size_t rowIndex, size_t columnIndex);
    [[nodiscard]] std::wstring BuildSelectionTsv() const;
    [[nodiscard]] std::optional<size_t> FindNearestVisibleRow(std::span<const GridGroupDesc> groups, size_t preferredRowIndex) const noexcept;
    void ReconcileSelectionForVisibleRows(std::span<const GridGroupDesc> groups);
    [[nodiscard]] size_t CountGroupHeadersBeforeRow(std::span<const GridGroupDesc> groups, size_t rowIndex) const noexcept;
    [[nodiscard]] size_t CountCollapsedRowsBeforeRow(std::span<const GridGroupDesc> groups, size_t rowIndex) const noexcept;
    [[nodiscard]] float GetRowTopDip(std::span<const GridGroupDesc> groups, size_t rowIndex) const noexcept;
    [[nodiscard]] float GetBodyContentHeight(std::span<const GridGroupDesc> groups) const noexcept;
    [[nodiscard]] float GetRawVerticalScrollableExtent(std::span<const GridGroupDesc> groups) const noexcept;
    [[nodiscard]] float AlignVerticalScrollExtentToVisibleItemBoundary(float rawExtentDip, std::span<const GridGroupDesc> groups) const noexcept;
    [[nodiscard]] float NormalizeVerticalScrollOffset(float offsetDip) const noexcept;
    [[nodiscard]] float NormalizeVerticalScrollOffset(float offsetDip, std::span<const GridGroupDesc> groups) const noexcept;
    [[nodiscard]] std::vector<VisibleBodyItem> BuildVisibleBodyItems(std::span<const GridGroupDesc> groups) const;
    [[nodiscard]] float GetVerticalScrollableExtent() const;
    [[nodiscard]] float GetVerticalScrollableExtent(std::span<const GridGroupDesc> groups) const;
    [[nodiscard]] float GetHorizontalScrollableExtent() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetContentRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetVerticalScrollbarRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetHorizontalScrollbarRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetVerticalThumbRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetHorizontalThumbRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetVerticalThumbHitRect() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetHorizontalThumbHitRect() const noexcept;
    [[nodiscard]] float GetColumnLeftDip(size_t columnIndex) const noexcept;
    [[nodiscard]] GridCellLayoutMetrics ComputeCellLayoutMetrics(const WindowHost& host,
                                                                 const D2D1_RECT_F& cellRect,
                                                                 const GridColumnDesc& columnDesc,
                                                                 const GridCellData& cellData) const noexcept;
    [[nodiscard]] VisibleColumnSpan ComputeVisibleColumnSpan(float clipRightDip) const noexcept;
    [[nodiscard]] bool HasAnimatedVisibleCells() const;
    [[nodiscard]] GridSortGlyphVisualState ResolveSortGlyphVisualState(const ThemePalette& theme, size_t columnIndex, uint64_t nowTickMs) const noexcept;
    [[nodiscard]] float ComputeSortGlyphTransitionProgress(uint64_t nowTickMs) const noexcept;
    [[nodiscard]] std::optional<size_t> ResolveHeaderBusyColumn() const noexcept;
    bool UpdateScrollbarHotState(const HitInfo& hit) noexcept;
    void SyncScrollbarAnimation(WindowHost& host) noexcept;
    [[nodiscard]] size_t GetModelColumnIndexForDisplayIndex(size_t displayIndex) const noexcept;
    [[nodiscard]] size_t ResolveHeaderReorderTargetDisplayIndex(float xDip) const noexcept;
    void MoveColumnToDisplayIndex(size_t columnIndex, size_t targetDisplayIndex) noexcept;
    void RebuildColumnDisplayIndexLookup() const noexcept;

    // Non-owning. Caller manages model lifetime. Valid from SetModel() until Grid destruction.
    // Invalidation validated at message entry by PruneStaleInteractionState().
    IDxGridModel* _model       = nullptr;
    IDxGridDelegate* _delegate = nullptr;
    mutable std::vector<float> _columnWidths;
    mutable std::vector<size_t> _columnDisplayOrder;
    mutable std::vector<size_t> _columnDisplayIndexByModel;
    mutable std::vector<GridGroupDesc> _cachedGroups;
    mutable std::vector<VisibleBodyItem> _cachedVisibleItems;
    GridSelectionModel _selectionModel;
    GridSelectionMode _selectionMode = GridSelectionMode::Extended;
    GridSortSpec _sortSpec{};
    std::optional<size_t> _headerBusyColumn;
    float _rowHeightBaseDip         = 28.0f;
    float _groupHeaderHeightBaseDip = 28.0f;
    float _headerHeightBaseDip      = 32.0f;
    float _rowHeightDip             = 28.0f;
    float _groupHeaderHeightDip     = 28.0f;
    float _headerHeightDip          = 32.0f;
    float _verticalScrollDip        = 0.0f;
    float _horizontalScrollDip      = 0.0f;
    uint32_t _lineClamp             = 2;
    std::optional<size_t> _hoveredRow;
    std::optional<size_t> _hoveredColumn;
    std::optional<size_t> _activeColumn;
    std::optional<size_t> _pressedHeaderColumn;
    std::optional<size_t> _dragReorderColumn;
    size_t _dragReorderTargetDisplayIndex = 0u;
    std::optional<size_t> _resizeColumn;
    float _pressedHeaderOriginXDip               = 0.0f;
    float _resizeOriginXDip                      = 0.0f;
    float _resizeInitialWidthDip                 = 0.0f;
    bool _dragVerticalThumb                      = false;
    bool _dragHorizontalThumb                    = false;
    ScrollbarHotPart _verticalScrollbarHotPart   = ScrollbarHotPart::None;
    ScrollbarHotPart _horizontalScrollbarHotPart = ScrollbarHotPart::None;
    ScrollbarAnimationState _verticalScrollbarAnimation{};
    ScrollbarAnimationState _horizontalScrollbarAnimation{};
    float _dragThumbOffsetDip = 0.0f;
    bool _headerBusy          = false;
    std::wstring _emptyStateText;
    // Keep test counters in the layout for all builds so ENABLE_TESTS only
    // affects helper APIs, not the object ABI across project boundaries.
    mutable uint64_t _debugPaintCount    = 0u;
    uint64_t _debugHeaderResizeDownCount = 0u;
    uint64_t _debugResizeMoveCount       = 0u;
    float _debugLastResizeDeltaDip       = 0.0f;
    float _debugLastResizeWidthDip       = 0.0f;
    uint64_t _debugHeaderReorderStartCount = 0u;
    uint64_t _debugHeaderReorderCommitCount = 0u;
    uint64_t _debugHeaderReorderNoOpCount = 0u;
    size_t _debugLastHeaderReorderColumn = 0u;
    size_t _debugLastHeaderReorderFromDisplayIndex = 0u;
    size_t _debugLastHeaderReorderRawTargetDisplayIndex = 0u;
    size_t _debugLastHeaderReorderNormalizedTargetDisplayIndex = 0u;
    struct SortGlyphTransitionState final
    {
        GridSortSpec from{};
        GridSortSpec to{};
        uint64_t startTickMs = 0u;
        bool active          = false;
    } _sortGlyphTransition{};
};

ThemePalette MakeDefaultThemePalette(bool dark) noexcept;
ThemePalette MakeThemePaletteFromViewerTheme(const ViewerTheme& viewerTheme) noexcept;

[[nodiscard]] ButtonVisualStyle ResolveButtonVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, bool primary) noexcept;
[[nodiscard]] ButtonVisualStyle ResolveButtonVisualStyle(const ThemePalette& theme,
                                                         bool enabled,
                                                         bool hovered,
                                                         bool pressed,
                                                         bool focused,
                                                         bool keyboardFocused,
                                                         bool primary,
                                                         float hoverStrength,
                                                         float focusStrength) noexcept;
[[nodiscard]] LabelVisualStyle ResolveLabelVisualStyle(const ThemePalette& theme, const std::optional<D2D1_COLOR_F>& textColorOverride) noexcept;
[[nodiscard]] ToggleVisualStyle ResolveToggleVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, bool checked) noexcept;
[[nodiscard]] ToggleVisualStyle ResolveToggleVisualStyle(const ThemePalette& theme,
                                                         bool enabled,
                                                         bool hovered,
                                                         bool pressed,
                                                         bool focused,
                                                         bool keyboardFocused,
                                                         bool checked,
                                                         float hoverStrength,
                                                         float focusStrength) noexcept;
[[nodiscard]] CheckboxVisualStyle ResolveCheckboxVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, bool checked) noexcept;
[[nodiscard]] CheckboxVisualStyle ResolveCheckboxVisualStyle(const ThemePalette& theme,
                                                             bool enabled,
                                                             bool hovered,
                                                             bool pressed,
                                                             bool focused,
                                                             bool keyboardFocused,
                                                             bool checked,
                                                             float hoverStrength,
                                                             float focusStrength) noexcept;
[[nodiscard]] RadioButtonVisualStyle ResolveRadioButtonVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, bool checked) noexcept;
[[nodiscard]] RadioButtonVisualStyle ResolveRadioButtonVisualStyle(const ThemePalette& theme,
                                                                   bool enabled,
                                                                   bool hovered,
                                                                   bool pressed,
                                                                   bool focused,
                                                                   bool keyboardFocused,
                                                                   bool checked,
                                                                   float hoverStrength,
                                                                   float focusStrength) noexcept;
[[nodiscard]] ProgressBarVisualStyle ResolveProgressBarVisualStyle(const ThemePalette& theme) noexcept;
[[nodiscard]] ToolbarVisualStyle ResolveToolbarVisualStyle(const ThemePalette& theme) noexcept;
[[nodiscard]] ColorSwatchVisualStyle ResolveColorSwatchVisualStyle(const ThemePalette& theme,
                                                                   bool enabled,
                                                                   bool hovered,
                                                                   bool pressed,
                                                                   bool focused,
                                                                   bool keyboardFocused                      = false,
                                                                   const std::optional<uint32_t>& swatchArgb = std::nullopt) noexcept;
[[nodiscard]] CardPanelVisualStyle ResolveCardPanelVisualStyle(const ThemePalette& theme) noexcept;
[[nodiscard]] TooltipVisualStyle ResolveTooltipVisualStyle(const ThemePalette& theme) noexcept;
[[nodiscard]] TextFieldVisualStyle ResolveTextFieldVisualStyle(const ThemePalette& theme,
                                                               bool enabled,
                                                               bool hovered,
                                                               bool focused,
                                                               bool keyboardFocused,
                                                               std::optional<D2D1_COLOR_F> caretColorOverride = std::nullopt) noexcept;
[[nodiscard]] GridSurfaceVisualStyle ResolveGridSurfaceVisualStyle(const ThemePalette& theme) noexcept;
[[nodiscard]] TreeSurfaceVisualStyle ResolveTreeSurfaceVisualStyle(const ThemePalette& theme) noexcept;
[[nodiscard]] GridHeaderVisualStyle ResolveGridHeaderVisualStyle(const ThemePalette& theme) noexcept;
[[nodiscard]] GridCheckboxVisualStyle ResolveGridCheckboxVisualStyle(
    const ThemePalette& theme, const D2D1_COLOR_F& rowFill, const D2D1_COLOR_F& rowText, bool enabled, bool hovered, bool selected, bool checked) noexcept;
[[nodiscard]] GridSwatchVisualStyle ResolveGridSwatchVisualStyle(
    const ThemePalette& theme, const D2D1_COLOR_F& rowFill, const D2D1_COLOR_F& rowText, bool selected, const GridCellData& cellData) noexcept;
[[nodiscard]] GridBadgeVisualStyle ResolveGridBadgeVisualStyle(
    const ThemePalette& theme, const D2D1_COLOR_F& rowFill, const D2D1_COLOR_F& rowText, bool selected, AdornmentTone tone) noexcept;
[[nodiscard]] TreeBadgeVisualStyle ResolveTreeBadgeVisualStyle(const ThemePalette& theme, AdornmentTone tone) noexcept;
[[nodiscard]] ComboBoxVisualStyle ResolveComboBoxVisualStyle(
    const ThemePalette& theme, ComboBoxVariant variant, bool enabled, bool hovered, bool popupOpen, bool focused, bool keyboardFocused) noexcept;
[[nodiscard]] float EvaluateEasing(EasingCurve curve, float t) noexcept;
[[nodiscard]] SortDirection NextSortDirection(SortDirection current) noexcept;
[[nodiscard]] VisibleSpan ComputeVisibleSpan(uint64_t totalItems, float itemExtentDip, float scrollOffsetDip, float viewportExtentDip) noexcept;
[[nodiscard]] std::optional<size_t> FindMnemonicTextIndex(std::wstring_view text, wchar_t mnemonic) noexcept;
void ShutdownAllWindowHostsForProcessExit() noexcept;
#if defined(ENABLE_TESTS)
[[nodiscard]] size_t DebugGetAttachedWindowHostCount() noexcept;
[[nodiscard]] uint32_t DebugGetSharedWindowHostAttachmentCountForThread(DWORD threadId) noexcept;
void DebugClearClipboardFallbackText() noexcept;
[[nodiscard]] bool DebugSetClipboardFallbackText(std::wstring_view text) noexcept;
[[nodiscard]] std::optional<std::wstring> DebugReadClipboardFallbackText() noexcept;
[[nodiscard]] bool DebugWriteClipboardUnicodeText(HWND ownerWindow, std::wstring_view text) noexcept;
#endif

class WindowHost final
{
public:
    enum class PresentationMode : uint8_t
    {
        HwndSwapChain,
        CompositionSwapChain,
    };

    struct AttachOptions final
    {
        PresentationMode presentationMode = PresentationMode::HwndSwapChain;
    };

    WindowHost() = default;
    ~WindowHost();

    WindowHost(const WindowHost&)            = delete;
    WindowHost(WindowHost&&)                 = delete;
    WindowHost& operator=(const WindowHost&) = delete;
    WindowHost& operator=(WindowHost&&)      = delete;

    [[nodiscard]] bool Attach(HWND hwnd) noexcept;
    [[nodiscard]] bool Attach(HWND hwnd, const AttachOptions& options) noexcept;
    void Detach() noexcept;

    void SetTheme(const ThemePalette& palette) noexcept;
    [[nodiscard]] const ThemePalette& GetTheme() const noexcept;

    void SetRoot(std::unique_ptr<Control> root);
    [[nodiscard]] Control* GetRoot() noexcept;
    [[nodiscard]] const Control* GetRoot() const noexcept;

    bool SetTooltip(std::wstring text, const D2D1_POINT_2F& originDip);
    bool SetTooltipDelayed(std::wstring text, const D2D1_POINT_2F& originDip);
    bool BeginTooltipHideDelay(uint64_t delayMs = 100u) noexcept;
    bool ClearTooltip() noexcept;
    [[nodiscard]] bool HasTooltip() const noexcept;
    [[nodiscard]] std::wstring_view GetTooltipText() const noexcept;
#if defined(ENABLE_TESTS)
    [[nodiscard]] std::wstring_view DebugGetPendingTooltipText() const noexcept;
    [[nodiscard]] bool DebugAdvanceTooltipDelayForTest() noexcept;
#endif

    void SetSmokeOverlayVisible(bool visible) noexcept;
    [[nodiscard]] bool IsSmokeOverlayVisible() const noexcept;

    // Mica/Acrylic backdrop hint — sets DWM system backdrop type on Windows 11 22H2+.
    // Returns true if the backdrop was applied (host HWND exists and API succeeds).
    enum class BackdropType : int
    {
        None    = 1, // DWMSBT_NONE
        Mica    = 2, // DWMSBT_MAINWINDOW
        Acrylic = 3, // DWMSBT_TRANSIENTWINDOW
        MicaAlt = 4  // DWMSBT_TABBEDWINDOW
    };
    bool SetSystemBackdrop(BackdropType type) noexcept;

    void Invalidate() const noexcept;
    [[nodiscard]] bool PrimeForShow() noexcept;
    void RequestAnimation() noexcept;
    void SetDefaultButton(Button* button) noexcept;
    [[nodiscard]] Button* GetDefaultButton() const noexcept;
    void SetCancelButton(Button* button) noexcept;
    [[nodiscard]] Button* GetCancelButton() const noexcept;
    void SetOnTabBoundary(std::function<bool(bool reverse)> onTabBoundary);
    void SetOnEscape(std::function<bool()> onEscape);
    void SetOnFocusChanged(std::function<void(Control* control)> onFocusChanged);
    void ResetInteractionState() noexcept;
    void SetFocusControl(Control* control) noexcept;
    [[nodiscard]] Control* GetFocusControl() const noexcept;
    [[nodiscard]] bool HandleMnemonic(wchar_t mnemonic) noexcept;
    void CommitFocusedTextInputBridge(bool notifyChange) noexcept;
    [[nodiscard]] bool TryReadTextInputBridgeState(const Control* control, TextInputBridgeState& outState) const noexcept;
    void SyncTextInputBridge(Control* control) noexcept;
    [[nodiscard]] bool HasActiveTextInputBridge() const noexcept;
    [[nodiscard]] HWND GetTextInputBridgeHwnd() const noexcept;
    void CaptureMouse(Control* control) noexcept;
    void ReleaseMouseCapture() noexcept;
    [[nodiscard]] HWND GetHwnd() const noexcept;
    [[nodiscard]] D2D1_RECT_F GetClientBoundsDip() const noexcept;
    [[nodiscard]] float GetDpi() const noexcept;
    [[nodiscard]] float PixelsToDip(float pixels) const noexcept;
    [[nodiscard]] float DipsToPixels(float dips) const noexcept;
    [[nodiscard]] PointDip PixelsToDipPoint(POINT pointPx) const noexcept;
    [[nodiscard]] std::optional<PointDip> ScreenPointToDipPoint(POINT screenPointPx) const noexcept;
    [[nodiscard]] POINT DipPointToScreenPoint(D2D1_POINT_2F pointDip) const noexcept;
    [[nodiscard]] InputModality GetInputModality() const noexcept;
    [[nodiscard]] bool IsKeyboardFocusVisible() const noexcept;
    [[nodiscard]] PresentationMode GetPresentationMode() const noexcept
    {
        return _presentationMode;
    }
    [[nodiscard]] ID2D1DeviceContext* GetDeviceContext() const noexcept;
    [[nodiscard]] IDWriteFactory* GetWriteFactory() const noexcept;
    [[nodiscard]] IDWriteTextFormat* GetTextFormat(FontRole role) const noexcept;
    [[nodiscard]] IDWriteTextFormat* GetTextFormat(FontRole role,
                                                   DWRITE_TEXT_ALIGNMENT alignment,
                                                   DWRITE_PARAGRAPH_ALIGNMENT paragraphAlignment,
                                                   bool wrap,
                                                   DWRITE_READING_DIRECTION readingDirection = DWRITE_READING_DIRECTION_LEFT_TO_RIGHT) const noexcept;
    [[nodiscard]] bool HasFluentIconFont() const noexcept;
    [[nodiscard]] ID2D1SolidColorBrush* GetSolidBrush(const D2D1_COLOR_F& color) const;

    [[nodiscard]] bool CopyTextToClipboard(std::wstring_view text) const noexcept;
    [[nodiscard]] std::optional<std::wstring> ReadTextFromClipboard() const noexcept;

    LRESULT HandleMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept;
#if defined(ENABLE_TESTS)
    [[nodiscard]] uint64_t DebugGetInvalidateCount() const noexcept;
    [[nodiscard]] uint64_t DebugGetRenderCount() const noexcept;
    [[nodiscard]] uint64_t DebugGetResizeCount() const noexcept;
    [[nodiscard]] uint64_t DebugGetResizeFailureCount() const noexcept;
    [[nodiscard]] uint64_t DebugGetPresentFailureCount() const noexcept;
    [[nodiscard]] bool DebugHasActiveAnimationSubscription() const noexcept;
    [[nodiscard]] IRawElementProviderFragmentRoot* DebugCreateAccessibilityProvider() const noexcept;
    void DebugSimulateDeviceLoss() noexcept;
    [[nodiscard]] size_t DebugGetBrushCacheSize() const noexcept;
    [[nodiscard]] bool DebugHasFallbackBrush() const noexcept;
    [[nodiscard]] bool DebugHasD2DContext() const noexcept;
    [[nodiscard]] size_t DebugGetConfiguredTextFormatCount() const noexcept;
    [[nodiscard]] UINT DebugGetModifierState() const noexcept;
    [[nodiscard]] bool DebugGetNonVisibleTextServiceBridgeFont(LOGFONTW& outLogFont) const noexcept;
    void DebugSetForceNullSolidBrushes(bool force) noexcept;
    [[nodiscard]] const Control* DebugHitTestControl(D2D1_POINT_2F pointDip) noexcept;
    [[nodiscard]] WindowHostCursorKind DebugResolveCursorKindForPoint(D2D1_POINT_2F pointDip) noexcept;
    [[nodiscard]] UINT DebugGetWidthPx() const noexcept
    {
        return _widthPx;
    }
    [[nodiscard]] UINT DebugGetHeightPx() const noexcept
    {
        return _heightPx;
    }
    [[nodiscard]] D2D1_RECT_F DebugGetTooltipBoundsDip() const noexcept;
    [[nodiscard]] bool DebugCaptureBitmap(WindowHostBitmapCapture& out) noexcept;
#endif

private:
    friend LRESULT CALLBACK TextInputBridgeWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    [[nodiscard]] bool EnsureDeviceIndependentResources() const noexcept;
    [[nodiscard]] bool EnsureDeviceResources() noexcept;
    [[nodiscard]] bool EnsureSizeDependentResources(bool allowHidden = false) noexcept;
    void PrepareForSwapChainResize() noexcept;
    void DiscardSizeDependentResources(std::wstring_view reason = {}) noexcept;
    void DiscardDeviceResources() noexcept;
    void RecreateBrushCache() const;
    void Render(const RECT* dirtyRectPx = nullptr) noexcept;
#if defined(ENABLE_TESTS)
    void Render(const RECT* dirtyRectPx, WindowHostBitmapCapture* capture) noexcept;
    [[nodiscard]] bool CaptureCurrentBackBuffer(WindowHostBitmapCapture& out) noexcept;
#endif
    void OnDpiChanged(HWND hwnd, UINT newDpi, const RECT* suggestedRect) noexcept;
    void OnSize(UINT widthPx, UINT heightPx) noexcept;
    void OnSetFocus() noexcept;
    void OnKillFocus(bool clearRetainedFocus) noexcept;
    [[nodiscard]] bool EnsureTextInputBridge(bool multiline) noexcept;
    void DestroyTextInputBridge() noexcept;
    void UpdateTextInputBridgeBounds(Control* control, bool multiline) noexcept;
    void ActivateTextInputBridge(Control* control) noexcept;
    void DeactivateTextInputBridge(bool restoreHostFocus) noexcept;
    void ApplyTextInputBridgeState(const TextInputBridgeState& state) noexcept;
    [[nodiscard]] std::optional<TextInputBridgeState> ReadTextInputBridgeState() const noexcept;
    [[nodiscard]] LRESULT HandleTextInputBridgeWindowMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    void SyncFocusedControlFromTextInputBridge(bool notifyChange) noexcept;
    void HandleTextInputBridgeCommand(UINT notifyCode) noexcept;
    void PruneStaleInteractionState() noexcept;
    void ResetRootInteractionState() noexcept;
    [[nodiscard]] bool HandleTabNavigation(bool reverse) noexcept;
    [[nodiscard]] Control* HitTestControl(D2D1_POINT_2F pointDip) noexcept;
    [[nodiscard]] WindowHostCursorKind ResolveCursorKindForPoint(D2D1_POINT_2F pointDip) noexcept;
    [[nodiscard]] HCURSOR ResolveCursorHandle(WindowHostCursorKind cursorKind) const noexcept;
    void UpdateHover(D2D1_POINT_2F pointDip, UINT modifiers) noexcept;
    void ClearPendingPointerDoubleClick() noexcept;
    [[nodiscard]] bool ShouldTreatButtonDownAsDoubleClick(Control* target, UINT buttonDownMessage, LPARAM lp) const noexcept;
    void RememberPointerButtonDown(Control* target, UINT buttonDownMessage, LPARAM lp) noexcept;
    void UpdateModifierStateForKey(UINT virtualKey, bool keyDown, bool systemKey) noexcept;
    void SetInputModality(InputModality modality) noexcept;
    [[nodiscard]] D2D1_POINT_2F PointFromLParam(LPARAM lp) const noexcept;
    [[nodiscard]] UINT GetModifierState() const noexcept;
    static bool AnimationTickThunk(void* context, uint64_t nowTickMs) noexcept;
    bool OnAnimationTick(uint64_t nowTickMs) noexcept;

    HWND _hwnd            = nullptr;
    UINT _dpi             = USER_DEFAULT_SCREEN_DPI;
    ThemePalette _palette = MakeDefaultThemePalette(false);

    mutable wil::com_ptr<IDWriteFactory> _dwriteFactory;
    mutable wil::com_ptr<IDWriteTextFormat> _bodyTextFormat;
    mutable wil::com_ptr<IDWriteTextFormat> _bodyStrongTextFormat;
    mutable wil::com_ptr<IDWriteTextFormat> _bodyLargeTextFormat;
    mutable wil::com_ptr<IDWriteTextFormat> _titleTextFormat;
    mutable wil::com_ptr<IDWriteTextFormat> _subtitleTextFormat;
    mutable wil::com_ptr<IDWriteTextFormat> _titleLargeTextFormat;
    mutable wil::com_ptr<IDWriteTextFormat> _displayTextFormat;
    mutable wil::com_ptr<IDWriteTextFormat> _headerTextFormat;
    mutable wil::com_ptr<IDWriteTextFormat> _smallTextFormat;
    mutable wil::com_ptr<IDWriteTextFormat> _iconTextFormat;
    mutable wil::com_ptr<IDWriteTextFormat> _heroIconTextFormat;
    mutable wil::com_ptr<IDWriteTextFormat> _monoTextFormat;
    mutable std::unordered_map<uint64_t, wil::com_ptr<IDWriteTextFormat>> _configuredTextFormats;
    mutable bool _fluentIconFontAvailabilityChecked = false;
    mutable bool _fluentIconFontAvailable           = false;

    wil::com_ptr<ID3D11Device> _d3dDevice;
    wil::com_ptr<ID3D11DeviceContext> _d3dContext;
    wil::com_ptr<ID2D1Factory1> _d2dFactory;
    wil::com_ptr<ID2D1Device> _d2dDevice;
    wil::com_ptr<ID2D1DeviceContext> _d2dContext;
    wil::com_ptr<IDXGISwapChain1> _swapChain;
    wil::com_ptr<IDCompositionDesktopDevice> _dcompDevice;
    wil::com_ptr<IDCompositionTarget> _dcompTarget;
    wil::com_ptr<IDCompositionVisual2> _dcompVisual;
    wil::com_ptr<ID2D1Bitmap1> _targetBitmap;
    D3D_FEATURE_LEVEL _featureLevel    = D3D_FEATURE_LEVEL_11_0;
    UINT _widthPx                      = 0;
    UINT _heightPx                     = 0;
    UINT _swapChainWidthPx             = 0;
    UINT _swapChainHeightPx            = 0;
    uint64_t _sharedGraphicsGeneration = 0u;

    mutable std::unordered_map<uint32_t, wil::com_ptr<ID2D1SolidColorBrush>> _brushCache;
    mutable wil::com_ptr<ID2D1SolidColorBrush> _fallbackBrush;
    mutable std::vector<uint32_t> _brushFailureLogKeys;
    // Keep test-only storage unconditional so WindowHost layout stays stable
    // even when ENABLE_TESTS differs between a DxUi producer and consumer.
    bool _debugForceNullSolidBrushes = false;
    wil::unique_hmodule _textInputBridgeModule;
    wil::unique_hwnd _textInputBridgeEdit;
    wil::unique_hfont _nonVisibleTextServiceBridgeFont;
    UINT _nonVisibleTextServiceBridgeFontDpi = 0u;
    bool _textInputBridgeMultiline           = false;
    TextInputBridgeState _textInputBridgeStateCache;
    bool _textInputBridgeStateCacheValid          = false;
    bool _textInputBridgeSelectionLogicalNewlines = true;
    bool _textInputBridgeImeComposing             = false;
    bool _textInputBridgeImeHasVisibleText        = false;
    std::optional<TextInputBridgeState> _textInputBridgeImeBaseState;
    std::vector<TextInputBridgeState> _textInputBridgeUndoHistory;
    std::vector<TextInputBridgeState> _textInputBridgeRedoHistory;

    std::unique_ptr<Control> _root;
    TooltipLayer _tooltipLayer;
    mutable uint64_t _debugInvalidateCount     = 0u;
    mutable uint64_t _debugRenderCount         = 0u;
    mutable uint64_t _debugResizeCount         = 0u;
    mutable uint64_t _debugResizeFailureCount  = 0u;
    mutable uint64_t _debugPresentFailureCount = 0u;
    // Non-owning observers into the current retained control tree. Ownership stays with
    // `_root` / `Panel::_children`, so these pointers must be cleared or ignored before
    // the observed tree is replaced or destroyed.
    struct PointerDoubleClickCandidate
    {
        Control* target  = nullptr;
        UINT downMessage = 0u;
        POINT pointPx{};
        uint64_t tickMs = 0u;
    };

    Control* _hoveredControl         = nullptr;
    Control* _capturedControl        = nullptr;
    Control* _focusedControl         = nullptr;
    Button* _defaultButton           = nullptr;
    Button* _cancelButton            = nullptr;
    Control* _textInputBridgeControl = nullptr;
    bool _textInputBridgeSyncing     = false;
    PointerDoubleClickCandidate _pendingPointerDoubleClick;
    std::function<bool(bool reverse)> _onTabBoundary;
    std::function<bool()> _onEscape;
    std::function<void(Control* control)> _onFocusChanged;
    UINT _modifierState                = 0u;
    InputModality _inputModality       = InputModality::Pointer;
    PresentationMode _presentationMode = PresentationMode::HwndSwapChain;
    bool _smokeOverlayVisible          = false;
    std::atomic_bool _detachInProgress{false};
    DWORD _attachmentOwnerThreadId    = 0u;
    uint64_t _animationSubscriptionId = 0;
    uint64_t _lastAnimationTickMs     = 0;
};
} // namespace RedSalamander::DxUi
