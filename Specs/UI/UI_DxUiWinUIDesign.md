# DxUI WinUI/Fluent Design Specification

**Author:** Ripley (Lead / Reviewer)
**Date:** 2026-04-04
**Last Updated:** 2026-04-29
**Status:** Authoritative design and behavior contract for `DxUi`; rollout closure and archived validation history live in `Specs/Plans/Done/UI_DxUiWinUIDesignAlignmentPlan.md`
**Scope:** Design system tokens, control specifications, retained-host behavior, and verification requirements
**Inspiration:** WinUI 3 / Windows 11 Fluent Design System

---

## 1. Design Philosophy

DxUI should feel like a **native Windows 11 application** — calm, focused, and approachable. The framework already has strong foundations (D2D rendering, theme system, accessibility). This spec defines the Fluent-aligned behavior and visual contract that migrated DxUI surfaces are expected to follow.

### Core Principles (from Fluent Design)

1. **Calm & Focused** — Color is used sparingly; accent color highlights interactive elements and selection only. Neutral surfaces dominate.
2. **Layered Hierarchy** — Base layer (navigation/chrome) vs. content layer, with elevation expressed through shadow and contour.
3. **Progressive Rounding** — Rounded corners at 4px (controls) and 8px (overlays) create a soft, modern feel.
4. **Purposeful Motion** — Animations are fast, direct, and context-appropriate. Never decorative.
5. **Accessible by Default** — 4.5:1 contrast ratio minimum, reduced-motion support, full keyboard navigation, UI Automation.

---

## 2. Design Tokens

### 2.1 Corner Radius

WinUI defines three tiers. DxUI currently uses `2 DIP` uniformly — this should be updated.

| Token | Value | Usage |
|-------|-------|-------|
| `ControlCornerRadius` | **4 DIP** | Buttons, checkboxes, text fields, combo boxes, list backplates, scrollbar thumbs, progress bars, sliders, radio button hover backplate |
| `OverlayCornerRadius` | **8 DIP** | Dialogs, teaching tips, generic flyouts/popups, overlay cards |
| `PopupRoundSmallCornerRadius` | **4 DIP** | Menu flyouts, ComboBox dropdown popups, and popup surfaces that intentionally emulate the Windows 11 `ROUNDSMALL` silhouette |
| `BarCornerRadius` | **4 DIP** | Progress bar track, scrollbar track, slider rail |
| `CircleRadius` | **full (50%)** | Radio button outer circle, toggle switch knob |
| `0 DIP` | — | Edges that touch container boundaries, maximized windows, split-button joins |

*Exception: ToolTip uses 4 DIP due to small size (per WinUI spec).*

**Nesting rule:** When a rounded element is inset inside another rounded element, the inner radius = outer radius minus the padding. E.g., an 8px-radius popup with 4px padding → inner items get 4px radius.

**Migration from 2 DIP:**
- Selection/hover backplates in FolderView: 2 → 4 DIP
- Focus borders: 2 → 4 DIP
- NavigationView breadcrumb hover: 2 → 4 DIP
- CardPanel: already configurable, default should become 4 DIP for inline, 8 DIP for overlay cards
- Menu flyouts and ComboBox dropdown popups: → 4 DIP (`PopupRoundSmallCornerRadius`)
- Other overlay popups/flyouts: → 8 DIP (`OverlayCornerRadius`)

### 2.2 Typography (Type Ramp)

DxUI's `FontRole` enum should be expanded to match the WinUI type ramp. On the Windows 11 baseline supported by RedSalamander, the runtime typography contract is:

- `Segoe UI Variable Small` for 11-12 DIP caption/header-scale text
- `Segoe UI Variable Text` for 13-31 DIP body, subtitle, title, and standard control text
- `Segoe UI Variable Display` for 32+ DIP large-display text

All app-owned visible text surfaces should route through the shared typography helper instead of hardcoding `Segoe UI`, `CreateFontW`, `CreateFontIndirectW`, or `DEFAULT_GUI_FONT`.

Current implementation status on 2026-04-26:

- the shared helper in `Common/DxUi/DxUi.Typography.h` now drives DxUi text roles and app-owned DirectWrite measurement/rendering without exposing native `HFONT` creation or HFONT-derived measurement APIs,
- `RedSalamander/Preferences.Dialog.cpp`, `RedSalamander/ConnectionManagerDialog.cpp`, and `RedSalamander/ManagePluginsDialog.cpp` no longer consume the legacy `ThemedControls` visible owner-draw button/toggle, modern combo, list, or list-header rendering APIs; `RedSalamander/ThemedControls.h/.cpp` are deleted, `RedSalamander/Win32UiHelpers.h/.cpp` are deleted, and the surviving pure color/DPI helpers live in `RedSalamander/UiMetrics.h/.cpp`,
- `RedSalamander/ConnectionManagerDialog.cpp` and `RedSalamander/ManagePluginsDialog.cpp` no longer carry app-owned dialog font propagation or broad `STATIC` DxUi host windows; their DxUi host/frame scaffolding uses explicit custom host window classes,
- `Common/DxUi/DxUi.Grid.cpp` treats checkbox double-clicks as two checkbox clicks, so pointer-double-click detection cannot turn a rapid checkbox toggle sequence into a no-op or row activation,
- `Common/DxUi/DxUi.Grid.cpp` and `Common/DxUi/DxUi.Tree.cpp` clamp interactive text rows to the shared 20 DIP Segoe UI Variable Body line-height minimum so compact density cannot collapse hit-test cadence below readable body text metrics,
- `Common/DxUi/DxUi.Accessibility.cpp` now honors explicit control-level accessible-name overrides, and hosts must use that path for visible text inputs that have no durable header/label; `RedSalamander/ShortcutsWindow.cpp` uses it for the Shortcuts search field so the empty search box keeps a non-empty UI Automation name,
- the focused `cmd_preferences_dialog_plugins_`, `cmd_connection_manager_window_`, and `cmd_plugin_configuration_dialog_` families are archived green on 2026-04-23 after that retirement,
- `Plugins/ViewerText/ViewerText.cpp` no longer keeps a visible GDI/HFONT root-shell fallback; its active shell contract is now guarded by focused `ViewerPETests` coverage and the `viewer_text_diff_perf` perf scenario archives from 2026-04-22,
- `RedSalamander/NavigationViewInternal.h` no longer carries a shared-font visible layout seam, and the refreshed NavigationView command selftests now validate the DxUi edit-host snapshot contract rather than searching for descendant native `Edit` children,
- `RedSalamander/FunctionBar.cpp` no longer creates/selects `HFONT` or uses GDI text measurement for visible Function Bar text/chrome; it measures through `DxUi.Typography` / DirectWrite and paints text/chrome through a Direct2D target bound only to the paint DC, with focused command selftests guarding against blank child strips after chrome toggles,
- the stale former `ThemedControls::CenterEditTextVertically()` hook used only by the inactive Compare Directories legacy edit fallback is retired, and the full `cmd_compare_directories_options_` family stays green on the DxUi edit-host contract,
- `RedSalamander/CompareDirectoriesWindow.cpp`, `RedSalamander/CompareDirectoriesWindow.Menu.cpp`, `RedSalamander/CompareDirectoriesWindow.Progress.cpp`, and `RedSalamander/CompareDirectoriesWindow.Options.cpp` now render Compare banner title/progress text and options body/footer typography through DxUi/DirectWrite instead of `Win32Ui::MeasureTextWidth(...)`, a local GDI `DrawTextW`/`GetDC` bridge, or native font propagation; the debug snapshots guard `usesDxUiBannerText == true`, `visibleDxBannerTextHostCount >= 1`, `visibleLegacyBannerTextCount == 0`, `hasNativeUiFontState == false`, `visibleNativeBodyControlCount == 0`, and `usesDxUiTypographyMetrics == true`, and no `CompareDirectoriesWindow*` source matches `HFONT` or `WM_SETFONT`,
- `RedSalamander/Preferences.General.*`, `RedSalamander/Preferences.Panes.*`, and `RedSalamander/Preferences.Viewers.*` now receive a `PreferencesTypographyContext` for page layout and measure General/Panes toggle labels, combo state labels, wrapped card descriptions, Viewers hint text, and viewer combo option widths through DxUi/DirectWrite rather than `Win32Ui::MeasureTextWidth(...)` or `PrefsUi::MeasureStaticTextHeight(...)`; command selftest snapshots guard `generalUsesDxUiTypographyContext == true`, `generalUsesDxUiTypographyMetrics == true`, `panesUsesDxUiTypographyContext == true`, `panesUsesDxUiTypographyMetrics == true`, `viewersUsesDxUiTypographyContext == true`, and `viewersUsesDxUiTypographyMetrics == true`,
- all Preferences page layout signatures now receive `PreferencesTypographyContext`; Preferences shell dialog font propagation is retired, and the former Preferences HDC paint/bitmap audit seams are now behind the shared Direct2D-on-HDC bridge rather than native font state,
- `Common/DxUi/DxUi.Typography.h` owns HFONT-free DirectWrite measurement (`MeasureSingleLineTextMetrics`, `MeasureSingleLineTextWidthPx`, and `MeasureWrappedTextHeightPx`) and no longer exposes measurement helpers that derive DirectWrite formats from caller-owned `HFONT` or helper APIs that create native fonts,
- `Common/DxUi/DxUi.TextInput.cpp` keeps the text-service HWND font bridge as an explicit Path B residual: it is non-visible, zero-region OS text-input interop only, has no visible typography or layout authority, and is surfaced to tests through `DebugGetNonVisibleTextServiceBridgeFont(...)`,
- `Plugins/ViewerWeb/ViewerWeb.cpp` renders its status message with Direct2D/DirectWrite instead of creating an app-owned `HFONT` or using GDI `DrawTextW`,
- `Plugins/ViewerPE/ViewerPE.cpp` and `Plugins/ViewerImgRaw/ViewerImgRaw.cpp` no longer keep unused native `_uiFont` handles,
- `RedSalamander/FolderWindow.FileSystem.cpp` no longer carries the dead fenced change-case and selection-mask dialog implementations or their unused hook helpers; the live route for those commands is the owned DxUi prompt-window path, with fresh focused archives on 2026-04-22,
- remaining visible-typography exceptions are tracked in `Specs/UI/UI_VisibleTypographyAudit.md`,
- the strict active-menu source audit is clean for `NONCLIENTMETRICS`, `lfMenuFont`, `GetTextExtentPoint32W`, `GetTextMetricsW`, and `CreateMenuFontForDpi`, and the broad remaining-Win32-UI audit has zero unallowlisted findings as of `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_012857_remaining_win32_ui_dependency_post_closeout_recheck/`. Former app-owned HDC paint seams in Compare, Connection Manager, FolderWindow splitter, NavigationView, Preferences, and themed input frames now use the shared `D2DHdcPaint::Session` bridge, and the last Compare Options legacy `STATIC` fallback was replaced by a custom Dx host. The former broad ShortcutsWindow grouped-collapse / reorder / persisted-layout / search-state failures were tracked separately and closed in `Specs/Plans/Done/DxUI_MigrationRoadmap.md` as DxUi behavior follow-on work rather than visible typography or GDI blockers.

| Token / FontRole | Size (epx) | Line Height (epx) | Weight | DxUI FontRole |
|------------------|------------|-------------------|--------|---------------|
| `Caption` | 12 | 16 | Regular (400) | `Small` |
| `Body` | 14 | 20 | Regular (400) | `Body` |
| `BodyStrong` | 14 | 20 | Semibold (600) | `BodyStrong` *(new)* |
| `BodyLarge` | 18 | 24 | Regular (400) | `BodyLarge` *(new)* |
| `Subtitle` | 20 | 28 | Semibold (600) | `Subtitle` *(new)* |
| `Title` | 28 | 36 | Semibold (600) | `Title` |
| `TitleLarge` | 40 | 52 | Semibold (600) | `TitleLarge` *(new)* |
| `Display` | 68 | 92 | Semibold (600) | `Display` *(new)* |

**Supplementary roles (keep):**
- `Icon` — `Segoe Fluent Icons`, 16 DIP (was 15 DIP — align to WinUI 16x16 icon grid). Verify `IconCache` glyph metrics are updated.
- `HeroIcon` — `Segoe Fluent Icons`, 24+ DIP
- `Emoji` — `Segoe UI Emoji`
- `Monospace` — `Consolas` for the current shared fallback path (a future switch to Cascadia Mono must still go through the shared helper)

**FontRole migration from current enum:**
- `Header` (current) → **`Subtitle`** (20/28 Semibold). Existing controls using `Header` should migrate to `Subtitle`. The `Header` enum value is deprecated but retained as an alias during the transition.
- `Small` → `Caption` (rename for WinUI alignment)
- `Body`, `Title`, `Icon`, `HeroIcon`, `Monospace` — unchanged

**Typography rules:**
- Minimum text: 14px Semibold or 12px Regular (legibility in all languages)
- Interactive Grid and Tree text rows must not shrink below the `Body` line-height token: 20 DIP at standard scale. Compact density may reduce rows only down to that shared minimum so hit testing, hover, selection, and tooltip tracking stay aligned with Segoe UI Variable body metrics.
- Visible DxUi text rendering MUST enable DirectWrite color-font drawing (`D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT`) through the shared DxUi draw options, including labels, buttons, tooltips, ComboBox edit text, TextField normal text, TextField selection redraws, and multiline edit text. Emoji and other color glyphs must render in their native color instead of being forced through the current text brush.
- Alignment: left-flush by default (RTL: right-flush)
- Casing: sentence case everywhere, including titles
- Line length: 50-60 characters optimal; never < 20 or > 60
- Truncation: ellipsis by default; clipping only when container has clear visual boundary

### 2.3 Color System

#### 2.3.1 Semantic Color Palette

The existing `ThemePalette` struct maps well to WinUI concepts. Proposed additions/renames for Fluent alignment:

| Semantic Role | Light Value | Dark Value | Current ThemePalette Field | Action |
|---------------|------------|------------|---------------------------|--------|
| **Window Background** | `#FFFFFF` | `#191A1C` | `windowBackground` | Keep |
| **Surface Background** | `#F3F3F3` | `#2D2D30` | `surfaceBackground` | Keep |
| **Card Background** | `#FFFFFF` | `#2D2D30` | — | **Add** `cardBackground` |
| **Smoke (Modal Overlay)** | `#000000` @ 30% | `#000000` @ 30% | — | **Add** `smokeOverlay` |
| **Text Primary** | `#141414` | `#EBEBF2` | `text` | Keep |
| **Text Secondary** | `#5D5D5D` | `#9A9AA8` | `subduedText` | Keep |
| **Text Disabled** | `#A0A0A0` | `#5D5D68` | `disabledText` | Keep |
| **Accent** | `#0078D4` | `#60CDFF` | `accent` | Keep (user-customizable) |
| **Accent Hover** | `#006CBD` | `#73D6FF` | — | **Add** `accentHover` |
| **Accent Pressed** | `#005A9E` | `#4CC2FF` | — | **Add** `accentPressed` |
| **Accent Disabled** | `#0078D4` @ 40% | `#60CDFF` @ 40% | — | **Derive** from accent |
| **Selection Fill** | `#0078D4` | `#2196F3` | `selectionFill` | Keep |
| **Selection Text** | `#FFFFFF` | `#FFFFFF` | `selectionText` | Keep |
| **Border (Subtle)** | `#E5E5E5` | `#3A3A3F` | `border` | Update values |
| **Border (Default)** | `#CDCDCD` | `#4A4A52` | — | **Add** `borderDefault` |
| **Border (Strong)** | `#ADADAD` | `#6A6A75` | — | **Add** `borderStrong` |
| **Focus Stroke Outer** | `#000000` | `#FFFFFF` | — | **Add** `focusStrokeOuter` |
| **Focus Stroke Inner** | `#FFFFFF` | `#000000` | — | **Add** `focusStrokeInner` |

#### 2.3.2 Accent Color Principles

- Accent color indicates **interactivity** and **selection state** — use sparingly.
- Accent values must be auto-generated for both light and dark modes to meet contrast requirements.
- Accent color on interactive controls: fill for primary/accent buttons, bottom border on focused text fields, selection highlights.
- Never use accent as a large background area (except accent buttons, which are small).

#### 2.3.3 Accent Color Derivation

When the accent color is user-customizable, hover/pressed/disabled variants must be derived automatically:

| Variant | Light Theme | Dark Theme |
|---------|------------|------------|
| **Accent** | Base value | Base value |
| **Accent Hover** | Lighten +8% in Oklab L* | Lighten +8% in Oklab L* |
| **Accent Pressed** | Darken -12% in Oklab L* | Lighten +16% in Oklab L* |
| **Accent Disabled** | Base @ 40% opacity | Base @ 40% opacity |

Use the Oklab color space for perceptually uniform lightness shifts. After derivation, verify that all accent variants meet 4.5:1 contrast against their backgrounds.

**Implementation:** `DxUi.Theme.cpp` should implement a `DeriveAccentVariants(D2D1_COLOR_F baseAccent, bool dark)` function that produces hover, pressed, and disabled variants.

#### 2.3.4 High Contrast

- In high contrast mode, **all** colors must come from system HC tokens. No hardcoded values.
- `ThemePalette::highContrast` flag gates this; all per-control visual style resolvers must respect it.
- Focus indicators: always visible, 2px minimum stroke width.

#### 2.3.5 Theme JSON5 Migration

Six custom themes exist in `Specs/Themes/` (Ugly, SolarFlare, RetroTerminal, PaperAndInk, NeonTokyo, ForestMist). These need migration for new palette fields:

**Required new keys in each theme JSON5:**
- `cardBackground`, `smokeOverlay`
- `accentHover`, `accentPressed` (or omit to auto-derive from `accent`)
- `borderDefault`, `borderStrong`
- `focusStrokeOuter`, `focusStrokeInner`

**Auto-derivation rules:** If a theme JSON5 omits a new key, the theme loader should derive it:
- `cardBackground` → same as `surfaceBackground`
- `smokeOverlay` → `#000000` @ 30% (fixed)
- `accentHover` / `accentPressed` → derived from `accent` per §2.3.3
- `borderDefault` → interpolate between `border` and `borderStrong` (50%)
- `borderStrong` → darken `border` by 20% (light) or lighten by 20% (dark)
- `focusStrokeOuter` → `#000000` (light) / `#FFFFFF` (dark)
- `focusStrokeInner` → inverse of outer

**Files:** `DxUi.Theme.cpp` (loader + derivation), `Specs/Themes/*.json5` (add explicit values where auto-derivation doesn't match the theme's aesthetic intent).

### 2.4 Elevation & Layering

WinUI uses a two-layer app model plus overlay levels:

| Surface | Elevation Value | Shadow | Corner Radius | Background |
|---------|----------------|--------|---------------|------------|
| **Layer (base)** | 1 | None | 0 | `windowBackground` or Mica |
| **Card** | 8 | Subtle | 4-8 DIP | `cardBackground` + 1px `border` |
| **Control** | 2 | Subtle (rest), reduced (pressed) | 4 DIP | Per-control fill |
| **Tooltip** | 16 | Small | 4 DIP | `tooltipBackground` + 1px border |
| **Generic flyout/popup** | 32 | Medium | 8 DIP | Solid or lightly translucent + 1px border |
| **Menu / Combo popup** | 32 | Medium | 4 DIP | Shared popup material family, `ROUNDSMALL`-equivalent silhouette |
| **Dialog** | 128 | Large | 8 DIP | Solid + Smoke behind |
| **Window** | 128 | System | 8 DIP (system) | Mica/Solid |

**Shadow specification (for D2D implementation):**

| Level | Offset Y | Blur Radius | Color (Light) | Color (Dark) |
|-------|----------|-------------|---------------|-------------|
| Subtle (Card/Control) | 1 DIP | 2 DIP | `#000000` @ 6% | `#000000` @ 12% |
| Small (Tooltip) | 2 DIP | 4 DIP | `#000000` @ 8% | `#000000` @ 16% |
| Medium (Flyout) | 4 DIP | 8 DIP | `#000000` @ 12% | `#000000` @ 24% |
| Large (Dialog) | 8 DIP | 16 DIP | `#000000` @ 14% | `#000000` @ 28% |

**Implementation note:** Use `ID2D1Shadow` effect with `D2D1_SHADOW_OPTIMIZATION_QUALITY` for flyout/dialog shadows — these surfaces are rendered once per show, so the GPU cost is acceptable. For controls that animate frequently (e.g., hover cards), consider pre-rendered 9-slice shadow bitmaps to avoid per-frame effect recomputation. Contour-only (1px border) is acceptable for lower-elevation surfaces when a shadow would be excessive.

### 2.5 Materials

DxUI adopts a hybrid material model:

- Whole-window backdrops may use DWM-provided Mica/Acrylic hints through `WindowHost::SetSystemBackdrop(...)`.
- When a whole-window backdrop is active, top-level windows MUST leave caption/border/text colors at the DWM system default so the system material can remain visible in the title bar. App-specific title-bar colors are only valid when the effective backdrop is `None`.
- DxUI-owned popup windows such as menus, flyouts, and popups MUST keep the popup HWND free of DWM system-backdrop hints. Popup material is app-rendered on the transparent composition host so custom RedSalamander themes do not inherit system Mica/Acrylic colors.
- DxUI-owned popup windows such as menus, flyouts, and popups MUST NOT opt into legacy native window-class shadows (for example `CS_DROPSHADOW`). Popup shadow is app-rendered inside the transparent composition host so the shadow shape stays aligned with the painted popup material.
- Menu flyouts and ComboBox dropdown popups MUST emulate the Windows 11 `ROUNDSMALL` corner silhouette in the DxUI paint path and popup-region math. This is a visual contract, not a requirement to call the DWM window-corner-preference API on popup hosts.
- Popup material blur, when present, MUST come from an app-side backdrop snapshot/effect stage rendered inside the popup surface. The popup `HWND` does not opt into DWM Mica/Acrylic tinting.
- That popup surface MUST remain self-contained: backdrop snapshot blur (optional by material) + base fill + glaze + contour + shadow. It may use popup-window alpha to reveal underlying content through the composition host, but it does not rely on DWM backdrop tinting to define the material.
- `Mica`, `Mica Alt`, and `Acrylic` popup treatments MUST be visibly distinct in the DxUI paint path, not only via palette-token tint changes.
- Acrylic-family popup treatments SHOULD bias toward a stronger backdrop blur with restrained tint/glaze so the user primarily perceives blurred background content, not an opaque colored slab.

| Material | Description | Usage | DxUI Priority |
|----------|------------|-------|---------------|
| **Mica** | Opaque, desktop-tinted background | App window base layer | P2 (optional, DWM handles) |
| **Mica Alt** | Stronger header/accent-tinted popup material | Popup/flyout surface variant | P1 |
| **Acrylic** | DWM backdrop hint for whole windows plus a stronger translucent/glazed popup treatment | Whole-window backdrop and transient popup surfaces | P1 |
| **Smoke** | Translucent black overlay (`#000000` @ 30%) | Behind modal dialogs | **P1** — easy to implement |

Shared popup blur constants:

| Popup material | Backdrop blur |
|----------------|---------------|
| **Mica** | **28 DIP** |
| **Mica Alt** | **34 DIP** |
| **Acrylic** | **40 DIP** |

These constants are shared by menu flyouts and ComboBox dropdowns. Popup-material tuning MUST update this table and the corresponding DxUi constants together so future adjustments have one normative baseline.

**Smoke implementation:** When showing a modal dialog, render a full-window `FillRectangle` with `D2D1::ColorF(0, 0, 0, 0.3f)` over the base content before painting the dialog.

### 2.6 Spacing System

Based on WinUI's 4px base grid:

| Token | Value | Usage |
|-------|-------|-------|
| `SpaceXS` | 2 DIP | Tight inner gaps (icon-to-text in compact mode) |
| `SpaceSM` | 4 DIP | Inner padding (checkbox indicator to label, menu icon to text) |
| `SpaceMD` | 8 DIP | Standard padding (button padding, card inner padding) |
| `SpaceLG` | 12 DIP | Group separation, between controls in a form |
| `SpaceXL` | 16 DIP | Section separation, dialog padding |
| `SpaceXXL` | 24 DIP | Page margins, large content padding |

**Touch target minimum:** All interactive elements must have a minimum touch target of **44×44 DIP**, even if the visual element is smaller. Use transparent padding to extend hit areas. This applies to: compact buttons (24 DIP height), scrollbar thumb (2-8 DIP width), checkbox/radio box (20 DIP visual), clear/close buttons (30 DIP visual).

### 2.7 Motion / Animation

Current DxUI timings (140ms button, 240ms tree expand) are close to WinUI. Formalize:

| Purpose | Easing | Duration | Usage |
|---------|--------|----------|-------|
| **Direct Entrance** | `cubic-bezier(0, 0, 0, 1)` — fast decelerate | 167 / 250 / 333ms | Popups appearing, page enter |
| **Point-to-Point** | `cubic-bezier(0.55, 0.55, 0, 1)` | 167 / 250 / 333ms | Reposition, resize, scroll |
| **Direct Exit** | `cubic-bezier(0, 0, 0, 1)` + fade | 167ms | Popup dismiss, tooltip hide |
| **Fade In/Out** | Linear | 83ms | Opacity-only transitions |
| **Control Hover** | Linear or fast decelerate | 100-150ms | Button fill, hover backplate |
| **Toggle/Expand** | `cubic-bezier(0.55, 0.55, 0, 1)` | 200-250ms | Tree expand, toggle slide, accordion |

**Reduced motion:** When `ThemePalette::reducedMotion` is true, all animations complete instantly (duration = 0). Opacity transitions can remain at <=83ms.

**Reduced motion affects all controls:** Any control that defines an auto-hide behavior (scrollbar §3.7: 1500ms auto-hide delay) must remain permanently visible when `reducedMotion` is true. Any control with animated size changes (toggle switch knob, radio button inner dot) should transition instantly.

**Easing implementation:** DxUI's `AnimationDispatcher` operates on 16ms ticks (~60fps). Implement cubic-bezier easing as a per-frame evaluation: given elapsed fraction `t` (0→1), evaluate the cubic polynomial using De Casteljau's algorithm or a pre-sampled lookup table (64 entries is sufficient). Expose as:

```cpp
float EvaluateEasing(EasingCurve curve, float t); // t in [0,1] → eased value
```

Standard curves:
- `EasingCurve::FastDecelerate` — `cubic-bezier(0, 0, 0, 1)`
- `EasingCurve::PointToPoint` — `cubic-bezier(0.55, 0.55, 0, 1)`
- `EasingCurve::Linear` — identity

### 2.8 RTL / BiDi Layout

DxUI now supports inherited `FlowDirection` for the core control/layout surfaces covered by this plan. The implemented framework behavior includes:

- inherited `Control::FlowDirection` (`LeftToRight` / `RightToLeft`)
- mirrored horizontal `StackPanel` layout
- right-to-left-aware text rendering where the control already uses DirectWrite layout helpers
- scrollbar-side flipping for scrollable surfaces that reserve a vertical bar
- right-to-left item ordering and keyboard navigation for `MenuBar` and `TabControl`
- mirrored horizontal `Slider` geometry and keyboard semantics

Current design constraints and follow-up rules remain:

- Never hardcode left/right assumptions in control layout. Use `Leading` / `Trailing` semantics where possible.
- Scrollbar positioning: always configurable (not hardcoded to right side).
- Text alignment: use DirectWrite's `DWRITE_TEXT_ALIGNMENT` which already supports RTL via `DWRITE_READING_DIRECTION_RIGHT_TO_LEFT`.
- Icon positioning: icons on the leading edge, not hardcoded left.

**Future follow-ups outside the current contract:**
- broader `Grid` column-order mirroring where the application needs it
- BiDi text editing in `TextField` (DirectWrite supports this natively)

Current controls should continue using relative positioning terms in their layout code so broader BiDi work stays incremental rather than invasive.

### 2.9 Per-Monitor DPI Handling

Per-monitor DPI awareness is required for all DxUI hosts and for hybrid shell windows that participate in a migrated DxUI surface.

**Required behavior:**
1. `WindowHost` MUST route `WM_DPICHANGED` and `WM_DPICHANGED_AFTERPARENT` through one shared handler that updates `_dpi`, refreshes the D2D context DPI, invalidates retained DPI-sensitive state, reapplies root layout, re-synchronizes the active text-input bridge, and repaints.
2. DPI-sensitive retained caches such as multiline text layouts, tooltip layouts, and popup/window metrics MUST invalidate on that path and rebuild at the new DPI rather than reusing stale measurements.
3. Child/custom windows that own local DPI-sensitive chrome or measurements MUST listen for `WM_DPICHANGED_AFTERPARENT` and refresh their local DPI state when the parent window crosses monitors.
4. `NavigationView` edit-suggest and full-path popups MUST survive a live DPI transition and recompute their geometry instead of requiring a close/reopen cycle.

**Design principle:** The retained layout model stays DIP-based. The important DPI work is shared cache invalidation, popup/window metric refresh, and crisp rerasterization at the new scale rather than introducing manual pixel-space layout branches.

---

## 3. Control Specifications

### 3.1 Button

**Existing DxUI:** Button with primary/secondary style, 140ms hover animation. Needs alignment.

#### Visual States

| State | Accent (Primary) Fill | Accent Text | Standard Fill | Standard Text | Standard Border |
|-------|----------------------|-------------|---------------|---------------|----------------|
| **Rest** | `accent` | `#FFFFFF` | `buttonFill` (subtle gray) | `text` | `buttonBorder` (1px) |
| **Hover** | `accentHover` (lighter) | `#FFFFFF` | `buttonHotFill` (lighter) | `text` | `buttonBorder` |
| **Pressed** | `accentPressed` (darker) | `#FFFFFF` | `buttonPressedFill` (darker) | `text` | `buttonBorder` |
| **Disabled** | `accent` @ 40% | `#FFFFFF` @ 60% | `buttonFill` @ 40% | `disabledText` | `border` @ 40% |
| **Focus** | Same as rest + focus ring | — | Same as rest + focus ring | — | — |

#### Measurements

| Property | Value |
|----------|-------|
| Height | 32 DIP (default), 24 DIP (compact) |
| Min width | 120 DIP |
| Horizontal padding | 12 DIP (text), 8 DIP (icon-only) |
| Corner radius | 4 DIP |
| Border | 1 DIP |
| Icon size | 16x16 DIP |
| Icon-to-text gap | 8 DIP |
| Font | `Body` (14/20, Regular) |

#### Focus Indicator (all controls)

WinUI uses a **double-stroke focus ring** for maximum contrast:
- **Outer stroke:** 2 DIP, `focusStrokeOuter` (black in light / white in dark)
- **Inner stroke:** 1 DIP, `focusStrokeInner` (white in light / black in dark)
- **Offset:** 2 DIP outside the control bounds
- **Corner radius:** control radius + offset (e.g., 4 + 2 = 6 DIP for buttons)

This double-ring ensures visibility on any background. Replace current single-stroke focus.

#### Button Variants (New)

| Variant | Description | DxUI Status |
|---------|-------------|-------------|
| **Standard** | Neutral fill, border | Exists (secondary style) |
| **Accent** | Accent fill, white text | Exists (primary style) |
| **DropDown** | Standard + chevron glyph, opens flyout | **New** — add `SetFlyout()` |
| **Split** | Left action + right dropdown, divider line | **New** — compound control |
| **Toggle** | On/off state with checked appearance | Exists (Toggle inherits Button) |
| **Hyperlink** | Text-only, accent color, underline on hover | **New** — minimal, add style enum |
| **Icon** | Icon-only, no text, square aspect | **New** — add icon-only layout mode |
| **Repeat** | Fires continuously while held | **New** — add repeat timer mode |

### 3.2 Menu System

DxUI now owns the retained menu/flyout system used by the main menu bar, context menus, the migrated navigation drive/history dropdown surfaces, and the migrated viewer top menu bars.

Windows that still source commands from a resource `HMENU` MAY keep that `HMENU` as a hidden command/model source, but the visible top-level chrome MUST come from a `DxUi::MenuBar` host child. Once attached, the window MUST clear its live `GetMenu(hwnd)` handle so runtime validation can detect regressions back to native menu-bar chrome.

#### Components

| Component | Description |
|-----------|-------------|
| **DxContextMenu** | Popup menu shown on right-click or programmatic invoke |
| **DxMenuBar** | Horizontal bar of top-level menus (File, Edit, View, Help) |
| **DxMenuFlyout** | Reusable menu content (items + submenus) |

#### Menu Item Types

| Type | Description | Visual |
|------|-------------|--------|
| **MenuFlyoutItem** | Standard action item | Icon area (16px bitmap or glyph) + Text + Accelerator text |
| **MenuFlyoutSubItem** | Opens cascading submenu | Same as item + chevron-right indicator |
| **ToggleMenuFlyoutItem** | Checkable item (on/off) | Check mark (✓) in icon area when checked |
| **RadioMenuFlyoutItem** | Radio-exclusive item | Filled dot (●) in icon area when selected |
| **MenuFlyoutSeparator** | Visual divider | Horizontal line, 1px, `border` color, 4px vertical margin |
| **MenuFlyoutHeader** | Non-interactive group label | Subdued text, `Caption` font role, 8px bottom margin |

#### Menu Flyout Measurements

| Property | Value |
|----------|-------|
| Min width | 128 DIP |
| Max width | 456 DIP |
| Item height | 30 DIP (default), 24 DIP (compact) |
| Icon area width | 28 DIP column; visible icon slot starts at 5 DIP and uses a 24 DIP draw area |
| Bitmap icon sizing | Native 16x16 stock shell bitmap preferred; popup rendering MUST not upscale bitmap icons beyond their native menu size |
| Text left padding | 36 DIP (after icon area) |
| Accelerator text right padding | 16 DIP from right edge |
| Submenu chevron right padding | 12 DIP from right edge |
| Corner radius | **4 DIP** (`PopupRoundSmallCornerRadius`, Windows `ROUNDSMALL`-equivalent) |
| Border | 1 DIP, `border` |
| Shadow | Medium (4 DIP Y, 8 DIP blur) |
| Inner padding (top/bottom) | 4 DIP |
| Cascading submenu offset | 0 DIP overlap, 4 DIP vertical offset |
| Cascading hover delay | 400ms |

#### Menu Item Visual States

| State | Background | Text | Icon |
|-------|-----------|------|------|
| **Rest** | Transparent | `text` | `text` |
| **Hover** | Inset rounded highlight backplate. Standard themes derive this from the popup material family; Rainbow mode uses a stable seeded rainbow fill keyed from the visible item label. | Foreground MUST be chosen from the actual highlighted fill so label text, accelerator text, icon glyphs, check marks, radio dots, and submenu chevrons remain readable. | Same contrasting foreground as text while highlighted |
| **Pressed** | `headerPressed` | `text` | `text` |
| **Disabled** | Transparent | `disabledText` | `disabledText` |
| **Checked (Toggle)** | No extra full-row idle backplate beyond the normal row chrome | `text` | Check mark in the accent family at rest; when highlighted it adopts the same contrasting foreground as the row text |
| **Checked (Radio)** | No extra full-row idle backplate beyond the normal row chrome | `text` | Filled dot in the accent family at rest; when highlighted it adopts the same contrasting foreground as the row text |

#### Menu Popup Behavior

- **Light dismiss:** Click outside or press Escape closes the menu.
- **App deactivation:** Menu popup windows are transient owned flyouts, not global topmost overlays. Losing app activation dismisses the full popup chain.
- **Keyboard:** Up/Down arrows navigate items (skip separators, headers, and non-interactive info rows). Enter/Space invokes. Home/End jump to first/last item. Keyboard-opened roots and submenus focus their first navigable item. Right arrow opens the highlighted submenu; on a highlighted leaf inside a top-level menu session it switches to the next enabled root menu and focuses that popup's first navigable item. Left arrow closes only the current submenu; once the root popup is active, Left switches to the previous enabled root menu and focuses that popup's first navigable item. Tab/Shift+Tab, Alt, and F10 exit the popup chain and return focus to the control that owned focus before menu mode started. Stationary cursor placement alone does not override keyboard root switching or highlighted-item ownership; pointer takeover requires actual mouse movement.
- **Submenu hover timing:** Cascading submenus follow standard Windows menu behavior. Entering a child submenu cancels the parent's pending close/open hover timer so the child stays open while hovered. Leaving that child submenu starts the delayed close for that child chain even if the pointer moves back onto the parent/root menu. Settling on a sibling leaf closes the old child chain after the cascade hover delay; settling on a sibling with children replaces the old child chain after the same delay.
- **Mnemonics:** Underlined characters in item text, activated by pressing the character key.
- **Submenu cascade:** On hover (400ms delay) or Right arrow key. Submenu appears to the right; if insufficient space, appears to the left.
- **Opening click behavior:** Opening a menu from the menu bar or a migrated navigation dropdown MUST leave the popup open after the opening pointer-up. The opening click/release path must not immediately light-dismiss the popup.
- **Chrome rollout contract:** Menu-style dropdowns in app chrome such as drive/history/navigation menus, pane sort menus, File Operations task option menus, and migrated plugin context menus MUST use this DxUI popup contract instead of `TrackPopupMenu`.
- **Popup host geometry:** Popup menus MUST use a transparent composition-backed host window that is larger than the visible menu surface so shadow and material can render outside the interactive surface. Hit-testing and dismissal are based on the visible surface rect, not the full transparent popup window bounds.
- **Root placement:** Context-menu roots support start/end horizontal anchoring and below/above vertical placement. App chrome controls at the bottom edge, such as pane status-bar sort indicators, MUST use end + above placement so the visible popup surface grows into the pane area instead of covering the bottom chrome.
- **Menu surface material:** The visible menu surface uses the WinUI-style rounded overlay card inside that transparent popup host. `Mica`, `Mica Alt`, and `Acrylic` MUST produce distinct backdrop-blur/tint/glaze/shadow treatments in the painted surface, while still preserving transparent outer corners in the host capture.
- **Backdrop readability contract:** Acrylic-style menu and combo popups must keep the underlying scene visibly blurred. Stronger blur is desirable, but popup rendering must remain visually tied to the captured backdrop rather than replacing it with a mostly opaque tint.
- **Highlight contrast contract:** Any highlighted interactive row MUST resolve foreground colors from the effective highlight fill, not only from the base theme text tokens. This applies to label text, accelerators, check marks, radio dots, submenu chevrons, and monochrome glyph icons. In Rainbow mode the highlight fill MUST use the stable seeded menu rainbow algorithm so the same visible label always maps to the same hue.
- **Informational rows:** Context menus may include non-interactive body rows with a leading label and optional trailing value. These rows do not hover, do not take keyboard selection, and are appropriate for disk/file-system summary data.

#### MenuBar Measurements

| Property | Value |
|----------|-------|
| Height | 30 DIP (default), 24 DIP (compact) |
| Item horizontal padding | 10 DIP (default), 6 DIP (compact) |
| Font | `Body` (14/20, Regular); compact uses `Small` (11/16) to keep descenders inside the shared 24 DIP chrome |
| Item hover background | `headerHovered` |
| Opens on click, then hover-opens adjacent menus | Standard menu bar behavior |

Additional layout rule:

- MenuBar items flagged as right-justified by the owning menu definition (for example `Help`) MUST remain anchored to the trailing edge instead of participating in the leading item flow.
- MenuBar height and MenuFlyout row height MUST come from the same shared menu-density tokens so top-level menus and popup menus stay aligned when density changes.

### 3.3 Radio Button

#### Visual Structure

```
[  (○)  Label Text  ]     ← full row is hit-target
       ↑
  Outer circle: 20x20 DIP, 1px border
  Inner dot (selected): 8 DIP diameter, accent fill
```

#### Measurements

| Property | Value |
|----------|-------|
| Outer circle diameter | 20 DIP |
| Inner dot diameter (selected) | 8 DIP (rest), 10 DIP (hover), 6 DIP (pressed) |
| Border | 1 DIP |
| Circle-to-label gap | 8 DIP |
| Row height | 32 DIP |
| Inter-item spacing | 4 DIP (within RadioButtons group) |
| Corner radius (hover backplate) | 4 DIP |

#### Visual States

| State | Circle Border | Circle Fill | Inner Dot | Label |
|-------|--------------|-------------|-----------|-------|
| **Rest (unchecked)** | `borderStrong` | Transparent | None | `text` |
| **Hover (unchecked)** | `borderStrong` | Transparent | None | `text` |
| **Pressed (unchecked)** | `borderStrong` | Transparent | None | `text` |
| **Rest (checked)** | `accent` | `accent` | White, 8 DIP | `text` |
| **Hover (checked)** | `accentHover` | `accentHover` | White, 10 DIP | `text` |
| **Pressed (checked)** | `accentPressed` | `accentPressed` | White, 6 DIP | `text` |
| **Disabled (unchecked)** | `border` @ 40% | Transparent | None | `disabledText` |
| **Disabled (checked)** | `accent` @ 40% | `accent` @ 40% | White @ 60%, 8 DIP | `disabledText` |
| **Focus** | Any above + double focus ring around row | | | |

#### RadioButtons Group Control

- Container that manages mutual exclusivity.
- Header text (optional, `BodyStrong` font role).
- `MaxColumns` property for multi-column layout (default 1).
- Selection follows focus (arrow keys select; Ctrl+Arrow moves focus without selecting, Space confirms).
- Exposes `SelectedIndex`, `SelectedItem`, `SelectionChanged`.

### 3.4 Text Box

#### Visual States (updated to match WinUI)

| State | Border | Fill | Bottom Accent | Placeholder | Clear Button |
|-------|--------|------|--------------|-------------|-------------|
| **Rest (empty)** | 1px `borderDefault` (bottom), 1px `border` (sides/top) | `inputFill` | None | `subduedText`, italic | Hidden |
| **Hover (empty)** | 1px `borderStrong` (bottom), 1px `borderDefault` (sides/top) | `inputFill` | None | `subduedText`, italic | Hidden |
| **Focused (empty)** | 2px `accent` (bottom), 1px `border` (sides/top) | `inputFill` (brighter) | **2px accent bottom border** | `subduedText`, italic | Hidden |
| **Rest (text entered)** | Same as rest empty | `inputFill` | None | — | Hidden |
| **Focused (text entered)** | 2px `accent` (bottom) | `inputFill` (brighter) | **2px accent bottom border** | — | Visible (X) |
| **Focused (text selected)** | 2px `accent` (bottom) | `inputFill` (brighter) | **2px accent bottom border** | — | Visible (X) |
| **Disabled** | 1px `border` @ 40% | `inputFill` @ 40% | None | `disabledText` | Hidden |

**Key WinUI feature — accent bottom border on focus:** When the text field gains focus, the bottom border animates from 1px neutral to 2px accent. This is the primary visual indicator of focus state for text inputs.

#### Measurements

| Property | Value |
|----------|-------|
| Height | 32 DIP (single-line) |
| Corner radius | 4 DIP |
| Horizontal padding | 12 DIP |
| Vertical padding | 4 DIP |
| Clear button width | 30 DIP |
| Clear button icon | X, 12 DIP |
| Header gap | 4 DIP (header text above field) |
| Font | `Body` (14/20) |
| Placeholder font | `Body` (14/20), subdued color |

#### Clear Button Behavior

- Shown only when: editable, single-line, has text, and has focus.
- NOT shown when: read-only, multi-line (`AcceptsReturn`), or empty.
- Click clears all text and keeps focus.

#### Single-Line Selection Clipping

- Long single-line selections MUST clip both the selection fill and selected text to the editable text viewport. Horizontal scroll may keep the caret/end of the selection visible, but the highlight and text MUST NOT paint outside the input border, under the clear button, or into surrounding dialog chrome.
- This contract applies to shared single-line text rendering in `TextField` and editable `ComboBox` surfaces, including prompt fields such as Rename.

#### Accessibility

- `TextField` must expose a non-empty UI Automation `Name`.
- When the field has no durable visible header or adjacent label, the host must provide an explicit accessible name through the control API; placeholder/cue text may mirror that label, but it must not be the only name source.
- Search fields must continue to expose `Edit` + `ValuePattern` semantics even when empty.

#### Password / Masked Mode

DxUI's `TextField` already supports masked input (`SetMasked(true)`, renders `●` U+25CF via `EM_SETPASSWORDCHAR` bridge). The spec must define WinUI-aligned visual behavior for this mode.

**Masked-mode rendering:**
- Each character displays as `●` (black circle, 14 DIP body font). No ligatures or kerning — fixed-width per glyph.
- Caret and selection work normally (selection highlight spans masked glyphs).
- Clipboard: `Ctrl+C` / `Ctrl+X` are **disabled** in masked mode (prevent password copy). `Ctrl+V` paste is allowed.
- Context menu: Copy/Cut items are hidden or disabled when masked.
- Masking is presentation-only: the control keeps the real text in memory for editing/submission, so this mode is not a secure-memory or zeroization feature by itself.

**Reveal button (new — WinUI PasswordBox feature):**
- Position: right-aligned, inside the field (same slot as Clear button — they are mutually exclusive).
- Icon: eye glyph (Segoe Fluent Icons `\xE7B3`), 16 DIP.
- Size: 30 DIP wide hit target.
- Behavior: **press-and-hold** reveals plaintext; release re-masks. This is the WinUI default (`PasswordRevealMode::Peek`).
- Visual states: transparent (rest), `hoverFill` (hover), `pressedFill` (pressed/revealing).
- Shown only when: masked mode, has text, and has focus. Hidden when empty or disabled.
- `PasswordRevealMode` options: `Peek` (hold to reveal, default), `Visible` (always plaintext — for "show password" checkbox workflows), `Hidden` (no reveal button).

**Measurements (additions to TextField):**

| Property | Value |
|----------|-------|
| Reveal button width | 30 DIP |
| Reveal button icon | Eye glyph, 16 DIP |
| Masked glyph | `●` U+25CF |
| Glyph spacing | Tracking +0.5 DIP (slightly wider than normal body text for readability) |

**Accessibility:**
- UIA: `IsPassword = true` when masked. Screen readers announce "password field" but do not read content.
- Reveal button: UIA `Invoke` pattern, announced as "Show password".

### 3.5 Toggle Switch

#### Visual Structure

```
[  Header Label                                    ]
[  On/Off Label [=====(●)====]                     ]
       ↑ knob      ↑ track       
```

#### Measurements

| Property | Value |
|----------|-------|
| Track width | 40 DIP |
| Track height | 20 DIP |
| Track corner radius | 10 DIP (full round) |
| Knob diameter | 12 DIP (rest), 14 DIP (hover), 10 DIP (pressed) |
| Knob corner radius | 50% (circle) |
| Track-to-label gap | 12 DIP |
| Row height | 36 DIP |
| Header above gap | 4 DIP |

#### Visual States

| State | Track Fill | Track Border | Knob Fill | Knob Size |
|-------|-----------|-------------|-----------|-----------|
| **Off - Rest** | Transparent | 1px `borderStrong` | `borderStrong` | 12 DIP |
| **Off - Hover** | Transparent | 1px `borderStrong` | `borderStrong` | 14 DIP |
| **Off - Pressed** | Transparent | 1px `borderStrong` | `borderStrong` | 10 DIP (compressed) |
| **On - Rest** | `accent` | `accent` | `#FFFFFF` | 12 DIP |
| **On - Hover** | `accentHover` | `accentHover` | `#FFFFFF` | 14 DIP |
| **On - Pressed** | `accentPressed` | `accentPressed` | `#FFFFFF` | 10 DIP (compressed) |
| **Disabled (either)** | Per above @ 40% opacity | Per above @ 40% | Per above @ 60% | 12 DIP |

**Animation:** Knob slides from left (off) to right (on) over 150-200ms with decelerate easing. Knob size changes animate over 100ms.

### 3.6 Checkbox

#### Visual Structure

```
[  [☑]  Label Text  ]     ← full row is hit-target
    ↑
  Box: 20x20 DIP, 4 DIP corner radius
  Check mark: 12 DIP, 1.5 DIP stroke weight
```

#### Visual States

| State | Box Fill | Box Border | Check Mark |
|-------|---------|-----------|-----------|
| **Unchecked - Rest** | Transparent | 1px `borderStrong` | None |
| **Unchecked - Hover** | Transparent | 1px `borderStrong` | None |
| **Unchecked - Pressed** | Transparent | 1px `borderStrong` | None |
| **Checked - Rest** | `accent` | `accent` | White ✓ |
| **Checked - Hover** | `accentHover` | `accentHover` | White ✓ |
| **Checked - Pressed** | `accentPressed` | `accentPressed` | White ✓ |
| **Indeterminate** | `accent` | `accent` | White — (horizontal line) |
| **Disabled** | As above @ 40% | As above @ 40% | As above @ 60% |

#### Measurements

| Property | Value |
|----------|-------|
| Box size | 20 DIP |
| Corner radius | 4 DIP |
| Check mark stroke | 1.5 DIP |
| Box-to-label gap | 8 DIP |
| Row height | 32 DIP |

### 3.7 Scrollbar

Extract from Grid/Tree/ScrollPanel and make standalone.

#### Measurements

| Property | Value |
|----------|-------|
| Track width (collapsed) | 2 DIP |
| Track width (expanded/hover) | 6 DIP |
| Track width (dragging) | 8 DIP |
| Thumb corner radius | 4 DIP (bar element) |
| Min thumb height | 30 DIP |
| Track color | `scrollbarTrack` |
| Thumb color (rest) | `scrollbarThumb` |
| Thumb color (hover) | `scrollbarThumbHot` |
| Thumb color (pressed) | `scrollbarThumbHot` (slightly darker) |
| Expand animation | 150ms, decelerate easing |
| Auto-hide delay | 1500ms after last scroll interaction |

**Behavior:** Thin-mode by default (overlay, 2 DIP). Expands to full width on hover/scroll. Auto-hides after inactivity (respects reduced-motion: stays visible if enabled).

**Scroll physics:**
- Mouse wheel: 3 lines per notch (respecting `SystemParametersInfo(SPI_GETWHEELSCROLLLINES)`). Apply smooth scroll animation over 200ms with `PointToPoint` easing.
- Page scroll (click on track): scroll one viewport height minus one line. Smooth animate over 250ms.
- Thumb drag: direct 1:1 mapping, no animation.
- Keyboard (Page Up/Down): same as page scroll. Arrow keys: single line scroll with 100ms ease.
- Trackpad/precision scroll: pixel-accurate delta, no stepping. Optional inertia (decelerate easing, 300ms).
- **Scroll anchoring:** When content above the viewport changes size (e.g., dynamic row heights), maintain the visual position of the first visible item.

### 3.8 Progress Bar

#### Determinate

| Property | Value |
|----------|-------|
| Track height | 2 DIP (rest), 4 DIP (indeterminate) |
| Track fill | `border` @ 40% |
| Progress fill | `accent` |
| Corner radius | 4 DIP (bar element) |
| Track width | Stretches to container |

#### Indeterminate

- Animated segment (~40% of track width) slides back and forth continuously.
- Uses `cubic-bezier(0.55, 0.55, 0, 1)` easing, 2000ms loop.
- Reduced motion: static centered segment, no animation.

### 3.9 Toolbar

For RedSalamander Monitor.

| Property | Value |
|----------|-------|
| Height | 42 DIP (default), 36 DIP (compact) |
| Button size | 32 DIP (default), 28 DIP (compact, icon-only) |
| Button padding | 4 DIP between buttons |
| Icon size | 16x16 DIP |
| Separator | 1px vertical line, 8 DIP horizontal margin, 60% height |
| Background | `surfaceBackground` |
| Bottom border | 1px `border` |

**Items:** Icon buttons (standard button visual states), toggle buttons (checked = accent fill), separators.

Compact-mode rule:

- The host chrome metrics for Monitor toolbar wrappers MUST shrink with DxUI compact density rather than only restyling the inner controls.

### 3.10 Status Strip

Multi-section support for RedSalamander Monitor.

| Property | Value |
|----------|-------|
| Height | 24 DIP (default), 20 DIP (compact host wrapper) |
| Section padding | 8 DIP horizontal |
| Section separator | 1px vertical, `border`, 60% height |
| Font | `Caption` (12/16) |
| Background | `surfaceBackground` |
| Top border | 1px `border` |

### 3.11 ComboBox

Align the existing 3-variant ComboBox to WinUI spec:

| Property | Value |
|----------|-------|
| Field height | 32 DIP (default), 24 DIP (compact) |
| Corner radius | 4 DIP (field), 4 DIP (dropdown popup, `PopupRoundSmallCornerRadius`) |
| Dropdown button width | 22 DIP (default), 20 DIP (compact) |
| Chevron icon | 12 DIP |
| Bottom accent border on focus | 2px `accent` (same as TextField) |
| Dropdown item height | 30 DIP (default), 24 DIP (compact); dropdown rows share the same density tokens and popup material system as menu flyouts |
| Dropdown max visible items | 8 (then scrollbar) |
| Dropdown shadow | Medium (flyout level) |

#### Visual States

| State | Field Fill | Field Border | Bottom Border | Chevron | Text |
|-------|-----------|-------------|--------------|---------|------|
| **Rest** | `inputFill` | 1px `borderDefault` (sides/top), 1px `border` (bottom) | None | `subduedText` | `text` |
| **Hover** | `inputFill` | 1px `borderStrong` (sides/top), 1px `borderDefault` (bottom) | None | `text` | `text` |
| **Focused / Open** | `inputFill` (brighter) | 1px `border` (sides/top) | **2px `accent`** (bottom) | `text` | `text` |
| **Disabled** | `inputFill` @ 40% | 1px `border` @ 40% | None | `disabledText` | `disabledText` |

**Variant behavior:** All three variants (Window, Modern, Edit) share the same visual states. The Edit variant additionally shows a text cursor when focused and allows typing to filter the dropdown list.

Dropdown surface rule:

- ComboBox dropdown popups MUST use the same overlay material system as menu flyouts (`overlayBackground` / `overlayBorder`, shared Acrylic/Mica/Mica Alt blur constants, shared 30/24 row metrics) so Preferences and other popup-heavy surfaces read as one family.
- ComboBox dropdown popups MUST use the same `PopupRoundSmallCornerRadius` token and the same `ROUNDSMALL`-equivalent silhouette as menu flyouts so standalone dropdowns and menu-family popups read as one Windows-aligned surface family.

Compact geometry rule:

- Compact mode MUST shrink the field and popup row geometry without clipping text, collapsing the leading text inset, or crowding the popup rows against the left edge.

### 3.12 Dialog

| Property | Value |
|----------|-------|
| Corner radius | 8 DIP |
| Shadow | Large (8 DIP Y, 16 DIP blur) |
| Smoke overlay behind | `#000000` @ 30% |
| Min width | 320 DIP |
| Max width | 600 DIP |
| Content padding | 24 DIP |
| Button area | Right-aligned, 8 DIP gap between buttons |
| Primary button | Accent style |
| Secondary/Close buttons | Standard style |
| Title font | `Subtitle` (20/28 Semibold) |
| Body font | `Body` (14/20) |

### 3.13 Tooltip

| Property | Value |
|----------|-------|
| Corner radius | **4 DIP** (exception: small element) |
| Shadow | Small (2 DIP Y, 4 DIP blur) |
| Padding | 8 DIP horizontal, 4 DIP vertical |
| Font | `Caption` (12/16) |
| Max width | 320 DIP |
| Show delay | 500ms |
| Display duration | 5000ms |
| Background | `tooltipBackground` |
| Border | 1px `border` |
| Text | `tooltipText` |

### 3.14 Tab Control

Required for connection manager and viewer tabs (see `DxUI_MigrationRoadmap.md`). Spec to be defined in a separate document. Key requirements:
- Horizontal tab strip with overflow (scroll arrows or dropdown when tabs exceed width)
- Close button per tab (optional)
- Drag-to-reorder tabs
- Content area swaps on tab selection
- WinUI `TabView` as reference

### 3.15 TreeView Visual Styling

The existing `Tree` control is already feature-complete (model/delegate, icons, badges, typeahead, accessibility). Define WinUI-aligned visual tokens:

| Property | Value |
|----------|-------|
| Item height | 32 DIP (standard), 24 DIP (compact) |
| Indent per level | 16 DIP |
| Expand/collapse icon | Segoe Fluent Icons chevron, 12 DIP, rotates 90° on expand |
| Expand animation | 200ms, `PointToPoint` easing (rotate + content reveal) |
| Icon size | 16×16 DIP |
| Icon-to-text gap | 8 DIP |
| Selection backplate | Full row width, 4 DIP corner radius, `selectionFill` |
| Hover backplate | Full row width, 4 DIP corner radius, `hoverFill` |
| Connecting lines | None (WinUI TreeView omits them; use indentation only) |

### 3.16 Grid / DataGrid Visual Styling

The existing `Grid` control is the application's primary data display. Define WinUI-aligned visual tokens:

| Property | Value |
|----------|-------|
| Header height | 32 DIP |
| Header font | `BodyStrong` (14/20, Semibold) |
| Header background | `headerBackground` |
| Header border (bottom) | 1px `gridLine` |
| Sort indicator | Segoe Fluent Icons chevron, 10 DIP, `subduedText`, animated rotation on sort change (150ms) |
| Row height | 26 DIP (default, configurable) |
| Alternating rows | Optional `surfaceBackground` on even rows (off by default) |
| Group header height | 28 DIP |
| Group header font | `BodyStrong` (14/20, Semibold) |
| Group header expand icon | Same as TreeView expand/collapse |
| Column resize grip | 4 DIP wide, cursor changes to `SizeWE` |
| Frozen columns | Vertical separator: 1px `borderDefault`, subtle shadow (1 DIP blur) |

Checkbox cells must toggle on every accepted left-button down, including the second down in a pointer double-click sequence. A double-click on a checkbox cell must not activate the row or leave the checkbox in the one-click state.

### 3.17 Slider

Implemented in DxUI for framework/application use such as Preferences (font size, opacity). Key requirements:
- Horizontal and vertical orientation
- Track: 2 DIP height, `border` @ 40% fill, 4 DIP corner radius
- Filled track (left of thumb): `accent` fill
- Thumb: 20 DIP diameter circle, white fill, `accent` border (checked style)
- Thumb hover: 22 DIP. Pressed: 18 DIP.
- Range: min/max values, optional tick marks
- Keyboard: Left/Right (or Up/Down for vertical) change by step, Page Up/Down by large step

### 3.18 Tracking Tooltip

Implemented for tracking-hover scenarios such as grid/tree-style metadata hover and future data-point surfaces.

| Property | Value |
|----------|-------|
| Follows cursor | Yes — repositions continuously while pointer moves |
| Corner radius | 4 DIP |
| Shadow | Small (2 DIP Y, 4 DIP blur) |
| Padding | 8 DIP horizontal, 4 DIP vertical |
| Font | `Caption` (12/16) |
| Offset from cursor | 16 DIP below, 8 DIP right (flips if near screen edge) |
| Show delay | 0ms (immediate, since triggered by continuous hover) |
| Hide delay | 100ms after pointer leaves the tracking region |
| Background | `tooltipBackground` |
| Multi-line support | Yes — content can include label/value pairs |

---

## 4. Compound Patterns

### 4.1 Card Pattern

Cards group related content on the content layer.

| Property | Value |
|----------|-------|
| Corner radius | 4 DIP (inline card), 8 DIP (overlay card) |
| Border | 1 DIP `border` |
| Background | `cardBackground` |
| Inner padding | 12-16 DIP |
| Shadow | Subtle (optional, 1 DIP Y, 2 DIP blur) |

### 4.2 InfoBar / Validation Banner

For status messages (info, warning, error).

| Severity | Fill | Border (left 3px accent) | Icon | Text |
|----------|------|--------------------------|------|------|
| Info | `infoFill` | `accent` | Info circle | `infoText` |
| Warning | `warningFill` | `#F7630C` (orange) | Warning triangle | `warningText` |
| Error | `errorFill` | `#C42B1C` (red) | Error circle | `errorText` |
| Success | `#DFF6DD` / `#1E3A1E` | `#0F7B0F` (green) | Checkmark | Green text |

**Dismiss button (optional):**
- Shown when `IsClosable` is true (default: true).
- Position: right-aligned, vertically centered.
- Size: 32×32 DIP hit target, 16 DIP X icon.
- Visual states: transparent (rest), `hoverFill` (hover), `pressedFill` (pressed).
- Clicking dismisses the InfoBar with a 167ms fade-out animation.

**Measurements:**
| Property | Value |
|----------|-------|
| Height | 40 DIP (single-line), auto (multi-line) |
| Left accent border | 3 DIP width |
| Icon size | 16 DIP |
| Icon-to-text gap | 8 DIP |
| Inner padding | 12 DIP (horizontal), 8 DIP (vertical) |
| Corner radius | 4 DIP |

### 4.3 Popup Positioning Engine

All overlay surfaces (context menus, flyouts, tooltips, combo dropdowns, teaching tips) share a generic positioning algorithm.

**Anchor model:**
- **Anchor element:** The control or point that triggered the popup.
- **Preferred placement:** One of `Top`, `Bottom`, `Left`, `Right`, `TopLeft`, `TopRight`, `BottomLeft`, `BottomRight`.
- **Alignment:** `Center` (default), `Start`, `End` along the cross-axis.

**Flip behavior:** If the popup would extend beyond the monitor work area:
1. Flip to the opposite side of the anchor.
2. If still clipped, try adjacent placements.
3. As a last resort, constrain to the monitor bounds and allow overflow (with scrollbar if needed).

**Multi-monitor:** Use `MonitorFromRect()` to determine which monitor owns the anchor point. Constrain the popup to that monitor's work area.

**Nesting:** When a popup opens another popup (submenu cascade), the child popup registers with a popup stack. Light-dismiss events walk the stack: clicking outside closes from the topmost popup downward. `Escape` closes only the topmost popup.

**Implementation:** Centralize in a `PopupPositioner::Calculate(anchor, preferredPlacement, popupSize, monitorBounds)` function. All overlay surfaces call this instead of implementing their own positioning.

### 4.4 Focus Management Architecture

**Tab navigation:**
- `WindowHost::HandleTabNavigation()` traverses controls via `FindAdjacentFocusable()`.
- Tab order follows the visual tree (depth-first, top-to-bottom). A future `TabIndex` property may override this.
- `Shift+Tab` reverses direction.

**Focus scopes:**
- A focus scope is a container that manages internal focus independently. Examples: `RadioButtons` group, dialog button area, menu.
- `Tab` into a focus scope selects the previously-focused item (or the first item). `Tab` out leaves the scope.
- Arrow keys navigate within the scope. `Tab` does not.

**Focus trap (modal):**
- When a modal dialog is shown (with smoke overlay), focus is trapped: `Tab` cycles within the dialog only.
- `Escape` dismisses the dialog and returns focus to the element that opened it.
- Initial focus: first focusable control in the dialog, or the primary button if marked.

**Focus ring rendering (centralized):**
- The double-stroke focus ring (§3.1) is rendered by `Control::PaintFocusRing()` in the base class, not per-control.
- Input: control bounds, corner radius. Output: outer ring at bounds + 2 DIP offset with outer radius, inner ring at bounds + 1 DIP offset.
- Only drawn when `_inputModality == Keyboard` (not on pointer focus).

### 4.5 Keyboard Command Routing

**Dispatch priority (highest to lowest):**
1. **Focused control** — `OnKeyDown()` handles control-specific keys (Enter for button, arrows for tree).
2. **Focus scope** — If the focused control doesn't handle it, the enclosing focus scope tries (arrow key navigation within a radio group).
3. **Window-level accelerators** — Registered keyboard shortcuts (Ctrl+C, Ctrl+V, Alt+F4). These fire regardless of which control has focus.
4. **Menu mnemonics** — Alt+letter activates menu bar items.

**Accelerator registration:**
- `WindowHost::RegisterAccelerator(modifiers, key, commandId)` registers a global shortcut.
- Accelerator text displayed in menu items comes from the same registry.
- Conflicts: if a focused control handles the key, the accelerator does not fire (control wins).

### 4.6 Compact / Dense Mode

Several controls define both default and compact heights (Button: 32/24, Menu item: 30/24). Compact mode is controlled at the container level.

**Mechanism:** Each container (Panel, CardPanel, dialog) can set `Density` to `Standard` (default) or `Compact`. Child controls inherit the density and adjust their heights accordingly.

**Compact adjustments:**
| Control | Standard Height | Compact Height | Other Changes |
|---------|----------------|---------------|---------------|
| Button | 32 DIP | 24 DIP | Padding: 12→8 DIP |
| Menu bar | 30 DIP | 24 DIP | Right-justified items still anchor to the trailing edge; compact switches the text role to `Small` and shrinks horizontal padding 10→6 DIP |
| Menu item | 30 DIP | 24 DIP | Header rows: 24→20 DIP |
| ComboBox field | 32 DIP | 24 DIP | Preserve leading text inset; no clipping |
| ComboBox popup item | 30 DIP | 24 DIP | Shares menu flyout density/material tokens and preserves left/right text padding |
| TextField | 32 DIP | 24 DIP | Vertical padding: 4→2 DIP |
| Toggle row | 36 DIP | 28 DIP | — |
| Checkbox/Radio row | 32 DIP | 24 DIP | — |
| Toolbar host (Monitor) | 42 DIP | 36 DIP | Button height shrinks 32→28 DIP |
| Status strip host (Monitor) | 24 DIP | 20 DIP | Section widths remain host-defined |
| Grid/list surfaces | Host-defined | Host-defined | Header/row metrics must shrink when the surface opts into compact density |
| Tree item | 32 DIP | 24 DIP | — |

**When to use:** Compact mode is appropriate for dense data forms (Preferences pages), toolbars, status bars, search/list results, and menu-heavy chrome. Standard mode is default for dialogs and content areas.

---

## 5. Verification

### Visual Verification

- Compare rendered controls side-by-side with WinUI 3 Gallery app screenshots
- Test all states: rest, hover, pressed, focused, disabled
- Verify light theme, dark theme, high-contrast (black and white), and rainbow mode
- Verify at 100%, 125%, 150%, 200% DPI scales
- Verify reduced-motion mode (all animations should be instant or ≤83ms)
- **Automated baseline comparison:** Implemented in `DxUiTests.Rendering.cpp` via `AttachedHostWindow` + `WindowHost::DebugCaptureBitmap(...)`, with golden PNGs stored in `Tests/DxUiTests/Baselines/`. Current acceptance threshold: ≤2% differing pixels with per-channel tolerance 8 to absorb subpixel variation.

### Performance Verification

- **Frame budget:** All control rendering must complete within **16ms** (60fps target). Measure via ETW TraceLogging events in `WindowHost::Paint()`.
- **Menu popup latency:** Context menu must appear within **100ms** of right-click. Measure from `WM_CONTEXTMENU` to first `Paint()` of the menu surface.
- **Animation smoothness:** No frame drops during standard animations (hover transitions, tree expand). Monitor via `AnimationDispatcher` tick timing — flag any tick > 20ms.
- **Text layout caching:** `IDWriteTextLayout` objects must be cached and reused. Re-creation only on text change, font change, or DPI change. Verify via a layout-creation counter in debug builds.
- **Add ETW TraceLogging events** for: `DxUI::Paint` (per-frame), `DxUI::Layout` (per-measure), `DxUI::HitTest` (per-pointer-move), `DxUI::PopupShow`, `DxUI::FocusChange`.

### Functional Verification

- Menu keyboard navigation: arrow keys, Enter, Escape, mnemonics, Home/End
- Radio button group: arrow keys select, Ctrl+arrow moves focus, Tab enters/exits group
- Focus ring visible on all interactive controls when using keyboard
- Tab navigation visits all focusable controls in visual order
- Modal dialog traps focus (Tab cycles within dialog only)
- Popup light dismiss: click outside closes, Escape closes topmost

### Accessibility Verification

UI Automation must report correct control types and patterns for every new control:

| Control | UIA ControlType | Required Patterns | Required Properties |
|---------|----------------|-------------------|---------------------|
| Button | `Button` | Invoke | Name, IsEnabled, IsKeyboardFocusable |
| Toggle | `Button` | Toggle | Name, ToggleState |
| Checkbox | `CheckBox` | Toggle | Name, ToggleState |
| RadioButton | `RadioButton` | SelectionItem | Name, IsSelected |
| TextField | `Edit` | Value | Name, Value, IsReadOnly |
| ComboBox | `ComboBox` | ExpandCollapse, Selection, Value | Name, ExpandCollapseState |
| Menu | `Menu` | — | Name |
| MenuItem | `MenuItem` | Invoke (or Toggle) | Name, IsEnabled, AccessKey |
| Tree | `Tree` | Selection | Name |
| TreeItem | `TreeItem` | ExpandCollapse, SelectionItem | Name, IsExpanded |
| Grid | `DataGrid` | Grid, Table, Selection | RowCount, ColumnCount |
| ProgressBar | `ProgressBar` | RangeValue | Name, Value, Minimum, Maximum |
| Slider | `Slider` | RangeValue | Name, Value, Minimum, Maximum |
| ToggleSwitch | `Button` | Toggle | Name, ToggleState |
| Dialog | `Window` | Window | Name, IsModal |
| Tooltip | `ToolTip` | — | Name |

### Theme Validation Checklist

Each control must be verified across all theme combinations:

| | Light | Dark | HC Black | HC White | Rainbow | Custom (×6) |
|---|---|---|---|---|---|---|
| Rest | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Hover | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Pressed | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Focused | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Disabled | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |

High contrast: all colors from system HC tokens, no hardcoded values. Focus indicators: always visible, ≥2px stroke width.

### Regression

- Run existing DxUI test suite (`Tests/DxUiTests/` — 14 test modules), including submenu delayed close/replacement and child-submenu hover timer regressions.
- Run self-test commands that exercise Preferences pages (already DxUI)
- Run RedSalamander self-tests (`Tools/Run-AllTests.ps1` — Commands/309, Compare/141, FileOps/68 cases)
- Manual smoke test of all migrated dialogs
- Visual baseline comparison (automated, see above)

---

## 6. Reference

- [WinUI Design Principles](https://learn.microsoft.com/en-us/windows/apps/design/design-principles)
- [WinUI Typography](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/typography)
- [WinUI Color](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/color)
- [WinUI Geometry (Rounded Corners)](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/geometry)
- [WinUI Motion](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/motion)
- [WinUI Materials](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/materials)
- [WinUI Layering & Elevation](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/layering)
- [WinUI Buttons](https://learn.microsoft.com/en-us/windows/apps/design/controls/buttons)
- [WinUI Menus](https://learn.microsoft.com/en-us/windows/apps/design/controls/menus)
- [WinUI Radio Buttons](https://learn.microsoft.com/en-us/windows/apps/design/controls/radio-button)
- [WinUI Toggle Switch](https://learn.microsoft.com/en-us/windows/apps/design/controls/toggles)
- [WinUI Text Box](https://learn.microsoft.com/en-us/windows/apps/design/controls/text-box)
- [Fluent 2 Design System](https://fluent2.microsoft.design/)
- [Windows UI Kit for Figma](https://aka.ms/WinUI/3.0-figma-toolkit)
- Related: `Specs/UI/UI_DxUiSharedGrid.md`, `Specs/UI/UI_VisualStyle.md`, `Specs/UI/UI_PreferencesDialog.md`, `Specs/UI/UI_NavigationView.md`
