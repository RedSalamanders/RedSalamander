# Operation Aurora — Keyboard, Theme-Menu, and Function-Bar DXUI Cycle Overlay

**Status:** Done — implemented, verified, documented, and archived on 2026-07-21
**Date:** 2026-07-21  
**Priority:** P2  
**Effort:** M  
**Risk:** MEDIUM — new transient DirectComposition surface on the synchronous theme-switch path  
**Category:** UI feedback / keyboard and menu workflow / accessibility / rendering  
**Planned at:** commit `148f270dec`  
**Primary owner:** RedSalamander main-window theme-cycle commands and the new transient overlay surface  
**Authoritative closeout targets:** `Specs/UI/UI_CommandMenuKeyboard.md`, a new
`Specs/UI/UI_ThemeCycleOverlay.md`, `docs/Themes.md`, `docs/UserGuide.md`, and `docs/DxUi.md`

> **Implementation instruction:** Follow the decisions and gates in this plan in order. This plan specifies
> behavior; it does not authorize redesigning the theme order, changing shortcut defaults, adding settings,
> or adding theme color keys. Preserve every unrelated worktree edit.
>
> **Mandatory first check:**
>
> ```powershell
> git status --short
> git diff -- RedSalamander/RedSalamander.cpp Common/DxUi RedSalamander/Ui `
>   RedSalamander/RedSalamander.rc RedSalamander/resource.h Specs/Plans/WIP/README.md
> ```
>
> At plan creation the worktree already contained unrelated FileSystem bridge/S3/File Operations work and
> two plan files dated 2026-07-20/21. Do not stage, rewrite, revert, or otherwise absorb those changes.

---

## 1. Goal

When a user changes the application theme with the keyboard **Previous Theme** / **Next Theme** commands, selects
a theme from **View -> Theme**, or clicks a theme-cycle command on the Function Bar, show a polished transient
DXUI popup centered over RedSalamander:

- the newly active theme name is large and visually dominant in the middle;
- the theme immediately before it in the cycle is smaller in the top-left;
- the theme immediately after it in the cycle is smaller in the bottom-right;
- pressing the same shortcut again immediately advances in that direction and refreshes the same popup;
- the popup animates into view, animates spatially between adjacent themes, remains fully visible for a defined
  delay after the latest selection, and then animates out instead of disappearing abruptly;
- a primary click/tap anywhere on the visible popup dismisses it immediately without activating it or invoking
  content underneath;
- the popup never takes focus, blocks work, or delays the actual theme switch; it consumes only an explicit
  dismissal activation that begins on its visible surface;
- the surface is drawn using the newly active theme and remains legible in Light, Dark, Rainbow, custom,
  system High Contrast, every DPI, reduced motion, and long/localized theme-name cases.

This is feedback for fast keyboard, menu, and Function Bar workflows, not a second theme picker. It tells the user
where they are in the existing ordered ring and what the neighboring choices are after a supported interaction.

## 2. Product Contract in One View

```text
       owned, no-activate DXUI popup centered over the main client

    +----------------------------------------------------------+
    |  <  Light                                                |
    |                                                          |
    |                         DARK                             |
    |                                                          |
    |                                              Rainbow  >  |
    +----------------------------------------------------------+

       top-left: previous       center: active       bottom-right: next
```

The diagram shows spatial relationships, not literal English or casing. Actual text uses each theme's display
name, the active UI language, the resolved flow direction, and DirectWrite casing as authored. Theme names must
never be transformed to uppercase in code.

### 2.1 Required behavior

| Event | Theme changes? | Overlay behavior |
| --- | --- | --- |
| Keyboard binding invokes `cmd/app/theme/selectNext` | Yes | Show/update with the next theme active. |
| Keyboard binding invokes `cmd/app/theme/selectPrev` | Yes | Show/update with the previous theme active. |
| Key auto-repeat invokes either command | Yes, once per accepted repeat | Reuse one HWND, replace the snapshot, cancel disappearance, then restart the full 900 ms delay after the newest transition. |
| View -> Theme -> Next/Previous | Yes | Show/update with the selected neighbor active and directional motion. |
| View -> Theme selects System/Light/Dark/Rainbow/App High Contrast | Yes, when different | Show the selected theme with its ring neighbors and a neutral transition. |
| View -> Theme selects a file or inline custom theme | Yes, when different | Show the selected theme with its ring neighbors and a neutral transition. |
| View -> Theme selects the already active theme | No effective change | Do not show or restart the overlay. |
| Function Bar Previous/Next theme command is clicked | Yes | Show/update exactly like the matching keyboard cycle command. |
| Function Bar command resolves to the already active direct theme | No effective change | Do not show or restart the overlay. |
| A parameterized direct-theme keyboard binding or unsupported non-menu command selects a theme | Yes | Do not show the overlay. |
| Preferences previews/applies a theme | Yes | Hide an existing overlay; do not show a new one. |
| Settings hot reload or external theme-file reload | Maybe | Hide an existing overlay; do not show a new one. |
| `WM_THEMECHANGED`, system tone, accent, or High Contrast changes | Maybe | Hide or retheme as specified in section 10; never present it as a user selection. |
| Startup restores the configured theme | Yes | Do not show the overlay. |
| Primary click/tap lands on the visible popup | No | Consume that activation and hide the popup immediately. |
| Click lands outside the visible popup | No | Route normally to the application; do not dismiss early. |
| Overlay creation/rendering fails | Yes | Keep the theme change; omit the overlay and log one bounded diagnostic. |

### 2.2 Animation and disappearance at a glance

The visible timing contract is mandatory:

```text
latest accepted theme selection
        |
        +-- 140 ms appearance or 160 ms theme-to-theme transition
        |
        +-- 900 ms fully visible disappearance delay
        |
        +-- 160 ms disappearance animation
        |
        +-- hidden; animation callback and dismissal timer stopped
```

- The **disappearance delay is 900 ms after the newest content transition has completed**, not 900 ms from
  application startup, popup creation, or an older key press.
- Every accepted keyboard repeat or different View -> Theme/Function Bar selection cancels a pending
  disappearance, updates the content, completes its newest transition, and starts a fresh 900 ms delay.
- A first appearance therefore remains on screen for about 1,200 ms total: 140 ms enter + 900 ms fully visible
  + 160 ms exit. An already-visible adjacent update uses 160 + 900 + 160 ms from that latest update.
- Reduced motion removes appearance/update/disappearance movement and fading but preserves the same 900 ms
  readable delay before an immediate hide.
- A primary click/tap on the popup is an explicit dismissal and bypasses both the remaining delay and the exit
  animation; this interaction must feel immediate.
- Exact interpolation, reversal, and state-machine rules are normative in section 11.

### 2.3 Explicit non-goals

- No clickable theme picker, mouse wheel navigation, theme-selection buttons, close glyph, context menu, or
  keyboard focus. The entire visible surface has only one action: dismiss.
- No change to the existing theme-cycle order or wraparound semantics.
- No change to the default `Shift+F11` Previous Theme and `Shift+F12` Next Theme bindings.
- No persistent setting for enablement, duration, placement, typography, opacity, or animation.
- No popup for Preferences/theme editing, parameterized direct-theme keyboard bindings, programmatic selection,
  application startup, hot reload, or system color changes. Theme choices made inside View -> Theme or from a
  clicked Function Bar theme command are in scope.
- No new theme semantic color tokens. The overlay consumes the existing resolved `AppTheme` ->
  `DxUi::ThemePalette` mapping.
- No full-window smoke/scrim layer and no dimming of application content.
- No native `MessageBox`, tooltip, toast notification API, GDI text, layered-window bitmap renderer, or XAML.
- No background worker. Selection, snapshot construction, resource updates, rendering requests, and teardown
  remain on the application UI thread.

## 3. Current-State Evidence

The implementation must start from these existing facts rather than creating a second theme model:

1. `RedSalamander/RedSalamander.cpp:8471` owns `BuildThemeCycleIds()`. Its order is:
   `builtin/system`, `builtin/light`, `builtin/dark`, `builtin/rainbow`,
   `builtin/highContrast`, then file themes, then inline settings themes.
2. Custom file themes and settings themes are already sorted by name, then ID, in
   `CollectCustomThemeGroups()`. The overlay must use the exact vector returned for command dispatch; it must
   not independently enumerate or sort themes.
3. `SelectAdjacentTheme(HWND, int)` computes wraparound and calls `ApplyThemeId(...)`.
4. `ApplyThemeId(...)` updates `g_settings.theme.currentThemeId`, derives `g_themeMode`, and calls
   `ApplyAppTheme(...)`.
5. `ApplyAppTheme(...)` applies the resolved theme to the main title/backdrop, FolderWindow, open tool windows,
   menu state, and main DXUI menu host before redrawing.
6. `ExecuteCommandById(...)` currently converts registered Previous/Next command IDs to `WM_COMMAND`, so the
   main handler cannot distinguish keyboard, menu, mouse function-bar, and programmatic invocation.
7. Default bindings currently assign `Shift+F11` to `cmd/app/theme/selectPrev` and `Shift+F12` to
   `cmd/app/theme/selectNext` in `RedSalamander/ShortcutDefaults.cpp`.
8. Built-in localized display names already exist as `IDS_PREFS_THEMES_BASE_SYSTEM`, `_LIGHT`, `_DARK`,
   `_RAINBOW`, and `_HIGH_CONTRAST`. Custom definitions already carry `name` and `id`.
9. `MakeAppThemeDxPalette(...)` already maps the effective theme, reduced-motion override, High Contrast,
   Rainbow state, overlay material, text, subdued text, border, background, and accent into one DXUI palette.
10. `DxUi::WindowHost` already supports a premultiplied-alpha `CompositionSwapChain`, initial-frame priming,
    app animation dispatch, DPI invalidation, device-loss recreation, shared brushes/text formats, UIA snapshot
    publication, and test bitmap capture.

The existing `Ui::AlertOverlayWindow` is not the right visual or behavioral primitive. It is an alert/prompt
surface with scrim, closability, buttons, and optional modality. Do not add a theme-specific mode to it. The new
surface is a passive transient status overlay and must not inherit alert semantics.

## 4. Terminology and State Model

### 4.1 Terms

- **Cycle ring:** the immutable ordered theme-ID vector captured for one accepted Previous/Next invocation.
- **Previous:** `(currentIndex + count - 1) % count` in the ring.
- **Current:** the theme that was successfully selected by the invocation.
- **Next:** `(currentIndex + 1) % count` in the ring.
- **Direction:** `-1` for Previous Theme, `0` for a direct View -> Theme/Function Bar choice, and `+1` for Next
  Theme.
  Direction affects transition motion only; it does not change the meaning or placement of the
  previous/current/next fields.
- **Snapshot:** an owned, self-contained value carrying IDs and full display names for the three positions.
- **Generation:** a monotonically increasing UI-thread sequence number assigned to each accepted keyboard,
  View -> Theme, or Function Bar selection that changes the configured theme.

### 4.2 Required owned model

Use an owned value with equivalent semantics to:

```cpp
enum class ThemeCycleDirection : int8_t
{
    Previous = -1,
    Direct   = 0,
    Next     = 1,
};

enum class CommandInvocationSource : uint8_t
{
    KeyboardShortcut,
    ThemeMenu,
    FunctionBarPointer,
    OtherWmCommand,
    Programmatic,
    SelfTest,
};

struct ThemeCycleOverlaySnapshot final
{
    uint64_t generation = 0;
    ThemeCycleDirection direction = ThemeCycleDirection::Next;
    std::wstring previousThemeId;
    std::wstring previousDisplayName;
    std::wstring currentThemeId;
    std::wstring currentDisplayName;
    std::wstring nextThemeId;
    std::wstring nextDisplayName;
};
```

The final names may differ, but these invariants are mandatory:

- Store owned `std::wstring` values. Never retain `wstring_view`, pointers into `g_settings`, theme-definition
  addresses, or resource buffers across a theme apply, hot reload, animation tick, or message dispatch.
- Keep the state UI-thread-confined. Do not add a lock or atomics to make a UI object callable from workers.
- Do not post raw heap pointers. This feature needs no cross-thread message. If implementation later introduces
  an asynchronous producer, it must use `PostMessagePayload(...)` / `TakeMessagePayload<T>(...)`, initialize and
  drain the payload window, and add teardown stress coverage.
- The overlay snapshot is derived from the same captured cycle ring and selected index used to apply the theme.
  Never rebuild the ring after applying and accidentally show neighbors from a different hot-reloaded order.

### 4.3 Edge cases

- `count == 0`: do not change theme and hide any stale overlay.
- `count == 1`: the current name is shown; both neighbor fields are hidden and excluded from UIA.
- `count == 2`: previous and next legitimately name the same other theme; show it in both positions because each
  position communicates a different direction.
- Current ID absent from the ring: preserve the existing fallback to `ThemeIdFromThemeMode(g_themeMode)`, then
  index `0` if that fallback is also absent. The snapshot must reflect the actual target chosen from that rule.
- Duplicate IDs should already be rejected upstream. If a corrupted in-memory ring contains duplicates, use the
  first exact ID match, matching `std::find` behavior; do not add an overlay-only repair policy.
- An empty custom display name falls back to its full ID. Do not synthesize `Unnamed Theme`.
- A display name containing embedded control characters is sanitized only through the canonical theme-definition
  validation/loading path. The overlay must not create a second sanitizer with different semantics.

## 5. Invocation-Origin Contract

The popup is specifically for keyboard Previous/Next, View -> Theme choices, and Function Bar theme clicks. Do not
infer origin from focus, recent key state, or a process-global/thread-local latch around `SendMessageW`; those
approaches are reentrancy-prone. Preserve origin explicitly from the keyboard, menu, or Function Bar invocation
seam into theme selection.

### 5.1 Required command routing

1. Extend the narrow internal command-dispatch API to carry an explicit `CommandInvocationSource`.
2. Actual `WM_KEYDOWN` / `WM_SYSKEYDOWN` shortcut paths pass `KeyboardShortcut`.
3. Debug/selftest shortcut dispatch passes `SelfTest` and can opt into keyboard behavior explicitly.
4. The mouse/touch Function Bar invocation path passes `FunctionBarPointer`; theme Previous/Next and direct theme
   commands from that source are eligible for the overlay when they change the configured theme.
5. View -> Theme command dispatch passes `ThemeMenu`. This covers the native/DXUI menu adapter's Previous,
   Next, built-in, and dynamic custom-theme rows. If `WM_COMMAND` is the authoritative menu boundary, classify
   the documented menu notification shape (`HIWORD(wParam) == 0` and `lParam == 0`) there once and immediately
   convert it to the typed source; do not inspect recent input later in theme code.
6. Other main `WM_COMMAND` calls pass `OtherWmCommand`; programmatic callers pass `Programmatic`.
7. Before the generic command-ID-to-`WM_COMMAND` bridge, handle canonical
   `cmd/app/theme/selectNext` and `cmd/app/theme/selectPrev` directly when the source is a keyboard shortcut.
8. Route View -> Theme Previous/Next through the same selection helper with `ThemeMenu` and directional motion.
9. Route View -> Theme built-in/custom rows through a target-ID helper with `ThemeMenu` and `Direct` motion. The
   dynamic custom-theme command-ID map must preserve menu origin; do not bypass the overlay only for custom rows.
10. A direct theme command invoked by the Function Bar uses `Direct` motion and the same no-op check as a direct
    View -> Theme choice.
11. Direct parameterized selection `cmd/app/theme/select/<themeId>` continues through `ApplyThemeId(...)` and does
    not show this overlay when invoked by an ordinary keyboard binding or programmatic caller; the approved direct
    selection surfaces are View -> Theme and the Function Bar.

This change must remain narrow. Do not migrate every command to a new public dispatch framework as part of this
feature.

### 5.2 Ordering

For one accepted keyboard, View -> Theme, or Function Bar selection:

1. Capture the ring and determine the target index. Previous/Next computes the adjacent index; a direct menu or
   Function Bar choice finds its selected target ID in the same ring.
2. Build the owned previous/current/next snapshot.
3. Apply the target theme normally.
4. Confirm the main window is still valid, visible, enabled, and not minimized.
5. Resolve the newly effective `AppTheme` and corresponding DXUI palette.
6. Show or update the overlay with that palette and snapshot.
7. Return control to the message loop.

If a direct menu or Function Bar target already equals `g_settings.theme.currentThemeId`, return as a no-op
without creating, showing, updating, or extending the overlay. Otherwise the active theme must be committed before
the overlay is painted, so popup and application agree. The overlay cannot veto, roll back, or delay the theme
change.

## 6. Display-Name Resolution

Define one shared main-application helper for cycle display names; do not duplicate the built-in-ID switch in
paint, tests, accessibility, and command routing.

| Theme ID | Visible name source |
| --- | --- |
| `builtin/system` | `IDS_PREFS_THEMES_BASE_SYSTEM` |
| `builtin/light` | `IDS_PREFS_THEMES_BASE_LIGHT` |
| `builtin/dark` | `IDS_PREFS_THEMES_BASE_DARK` |
| `builtin/rainbow` | `IDS_PREFS_THEMES_BASE_RAINBOW` |
| `builtin/highContrast` | `IDS_PREFS_THEMES_BASE_HIGH_CONTRAST` |
| File or inline custom theme | `ThemeDefinition::name`, falling back to `ThemeDefinition::id` when empty |
| Unknown defensive fallback | Full ID |

Rules:

- Use the active language resources at invocation time.
- `builtin/system` remains named **System**, even when the effective Windows tone is currently dark or Windows
  High Contrast is active. Do not relabel it Light/Dark based on resolution.
- `builtin/highContrast` remains the application High Contrast entry and uses the existing localized
  “High Contrast (App)” name.
- Do not show source paths, `user/` prefixes removed by ad hoc code, or base-theme names instead of custom names.
- Preserve Unicode exactly. DirectWrite handles UTF-16, combining marks, surrogate pairs, bidi, and fallback
  fonts. Tests must include Latin, Japanese, accented Latin, emoji/supplementary characters, and an RTL name.

## 7. Window and Lifetime Architecture

### 7.1 Application ownership

Add one `ThemeCycleOverlayWindow`-equivalent object to the main application composition root. Do not add a leaked
singleton or independent global lifetime. It is created lazily on the first eligible keyboard, theme-menu, or
Function Bar selection, reused for the process lifetime, and explicitly destroyed before the main window and
shared DXUI graphics shutdown.

Expected source boundary:

- `RedSalamander/Ui/ThemeCycleOverlayWindow.h`
- `RedSalamander/Ui/ThemeCycleOverlayWindow.cpp`
- minimal ownership and routing changes in `RedSalamander/RedSalamander.cpp`

The implementation may choose a nearby established file split if the project layout requires it, but the
transient window must not be embedded as hundreds of lines in `RedSalamander.cpp`.

### 7.2 Native window contract

The overlay is one borderless owned popup HWND:

- style: `WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS`;
- extended style: `WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_NOREDIRECTIONBITMAP`;
- owner: the main RedSalamander top-level HWND;
- never `WS_EX_APPWINDOW`, never topmost, never taskbar-visible;
- show through `ShowWindow(..., SW_SHOWNOACTIVATE)` and `SetWindowPos(..., SWP_NOACTIVATE | ...)`;
- return `MA_NOACTIVATE` for `WM_MOUSEACTIVATE`;
- return `HTCLIENT` for points inside the actual rounded visible surface and `HTTRANSPARENT` for transparent
  shadow/animation gutters and pixels outside rounded corners;
- do not call `SetFocus`, `SetForegroundWindow`, `SetActiveWindow`, raw `SetCapture`, or register accelerators;
- do not disable the owner;
- do not create a nested/modal message loop.

For click dismissal, route pointer handling through the attached `DxUi::WindowHost`/root control and its balanced
`CaptureMouse(...)` /
`ReleaseMouseCapture()` state so the down/up pair cannot leak to underlying content. The root control is
non-focusable and exposes only the dismiss action.

Use `wil::unique_hwnd` for ownership. Close with `_hwnd.reset()`, never
`DestroyWindow(_hwnd.get())`. COM/DX resources use `wil::com_ptr<T>`. Animation subscription is a scalar token
that is always unsubscribed before window destruction.

### 7.3 DXUI host

- Attach one `DxUi::WindowHost` in `PresentationMode::CompositionSwapChain` so unused pixels remain transparent.
- Use the shared DXUI device/factory/cache path; do not create a private D3D11 device per popup.
- Prime layout/resources and render a complete initial frame with `PrimeForShow()` and
  `RenderInitialFrameForShow()` before exposing the window. A first frame that appears blank until another mouse,
  paint, move, or timer message is a correctness failure.
- Handle `WM_SIZE`, `WM_DPICHANGED`, `WM_DISPLAYCHANGE`, `WM_SETTINGCHANGE`, and device-loss recreation through
  the existing `WindowHost` paths.
- `WM_ERASEBKGND` returns nonzero; visible pixels are entirely Direct2D/DirectWrite.
- On `WM_NCDESTROY`, detach `WindowHost`, clear `GWLP_USERDATA`, release the wrapper-owned HWND correctly, and
  stop animation before any callback can touch the destroyed object.

### 7.4 Shared popup-surface rendering

Before adding paint helpers, search `Common/DxUi` again for a landed shared transient/popup material primitive.
At plan time menu and ComboBox popup material resolution is private to their implementation files, while
`CardPanel` is an inline-card visual and not the elevated overlay material.

If no canonical primitive has landed, add one reusable DXUI popup-surface primitive (for example a
`TransientSurface`/`PopupSurface` control or a shared material draw helper) under `Common/DxUi`. It must own:

- rounded surface fill from `ThemePalette::overlayBackground` and `overlayMaterial`;
- `overlayBorder` outer/inner rim treatment;
- the existing DXUI overlay material policy for Solid/Mica/MicaAlt/Acrylic;
- soft shadow in ordinary themes;
- solid/no-blur High Contrast behavior;
- opacity and scale/translation supplied by the caller;
- composition-safe premultiplied alpha;
- device-dependent brush/effect recreation and device-loss behavior.

Do not copy the menu/ComboBox material math into a RedSalamander-local function under a new name. If the shared
implementation consolidates existing menu/ComboBox material code, preserve their exact visuals with DxUi tests
and treat that consolidation as a separately reviewable change within the same branch.

Update `docs/DxUi.md` and the relevant Visual/UI DXUI specification at closeout when adding this shared primitive.

## 8. Geometry and Placement

All design dimensions are DIPs and are converted using the popup's current per-monitor DPI.

### 8.1 Anchor area

- Anchor to the main application's **client rectangle in screen coordinates**, not the monitor work area and not
  the currently focused child pane.
- The title bar is naturally excluded because it is non-client. The in-app menu row remains part of the client;
  centering uses the full client so the popup represents the application, not one pane.
- Center by arithmetic on the full rectangle, then round physical coordinates consistently so left/right and
  top/bottom differ by at most one pixel.
- Keep the surface wholly inside the main client with 16 DIP minimum edge clearance whenever the client allows.
- Clamp the final popup window to the monitor work area as a defensive fallback for unusual owner coordinates.
- Reposition without activation on owner `WM_MOVE`, `WM_SIZE`, `WM_DPICHANGED`, restored/maximized transitions,
  and display topology changes while visible.

### 8.2 Size

The transparent popup HWND includes a 12 DIP shadow/animation gutter around the visible surface.

| Dimension | Preferred | Responsive rule |
| --- | ---: | --- |
| Visible surface width | 560 DIP | `clamp(clientWidth * 0.58, 360, 640)`, then at most `clientWidth - 32`. |
| Visible surface height | 236 DIP | `clamp(clientHeight * 0.32, 180, 260)`, then at most `clientHeight - 32`. |
| Shadow/animation gutter | 12 DIP each side | May reduce to 4 DIP in a tiny client; never clip text. |
| Corner radius | 18 DIP | 0 in High Contrast if the platform's contrast policy requires rectangular borders. |
| Outer content inset | 24 DIP | May reduce to 16 DIP below 420 DIP surface width. |

If the main client is smaller than 192 x 128 DIP, apply the theme but suppress the overlay; a crushed unreadable
popup is worse than no popup. This is a rendering fallback, not a theme-selection failure.

### 8.3 Text rectangles

- **Previous:** top-left, leading aligned, top aligned; maximum 44% of content width and 42 DIP height.
- **Current:** horizontally centered and vertically centered in the surface; use the content width minus 64 DIP,
  leaving clear separation from both neighbor rectangles.
- **Next:** bottom-right, trailing aligned, bottom aligned; maximum 44% of content width and 42 DIP height.
- Neighbor rectangles must not overlap the current rectangle at the minimum supported size.
- Flow direction mirrors leading/trailing placement. In RTL UI, logical Previous remains at the leading top
  corner and logical Next at the trailing bottom corner; directional chevrons also mirror.

## 9. Typography and Text Fitting

### 9.1 Current theme

- Default role: `DxUi::FontRole::TitleLarge` (40/52 Semibold).
- One visual line only.
- Center alignment and center paragraph alignment.
- Measure against the actual content rectangle before paint.
- If it does not fit, step through bounded shared roles/sizes down to `Subtitle` (20/28 Semibold).
- If it still does not fit at the minimum role, use DirectWrite end ellipsis.
- Never horizontally scale glyphs, marquee text, reduce below 20 DIP, or allocate/recreate a text format every
  animation frame.

### 9.2 Neighbor themes

- Default role: `BodyStrong` (14/20 Semibold); `Small` is permitted only in the tiny responsive layout.
- Use `ThemePalette::subduedText` in ordinary themes and the resolved High Contrast text role in High Contrast.
- One line with end ellipsis.
- Prepend/append a small shared Fluent chevron glyph when the Fluent font is available; fall back to the Unicode
  single-angle glyph. The glyph is decorative in UIA and follows flow direction.
- Do not display literal, hardcoded “Previous” or “Next” captions in the visible surface. Position plus chevron
  supplies the visual relation; localized relation text belongs in accessibility names.

### 9.3 Contrast

- The center name uses `ThemePalette::text` unless the shared popup material resolver selects a stronger
  contrast-corrected overlay text color.
- All three text runs must meet at least WCAG 2.1 AA 4.5:1 against the effective composite surface. The large
  current label should also meet 4.5:1 even though large-text AA would permit less.
- High Contrast uses resolved system/app High Contrast colors at full opacity. No acrylic, blur, translucent
  text, rainbow gradient, or low-opacity shadow may override those colors.
- Rainbow may use a restrained accent rim or top highlight from the shared overlay material. The current theme
  name itself remains a stable solid high-contrast color; do not animate hue through glyphs.

## 10. Material and Theme-Change Behavior

### 10.1 Newly applied palette

The popup is themed from `ResolveConfiguredTheme()` **after** the selection has been applied, then converted by
`MakeAppThemeDxPalette(...)`. This guarantees that switching from Light to Dark produces a Dark popup, not a
Light popup announcing Dark.

Use the palette's existing `overlayMaterial` policy:

- Solid: opaque/tinted overlay background, border, and shadow.
- Mica/MicaAlt/Acrylic: shared DXUI material treatment, with ordinary fallback to Solid if DWM/backdrop/effect
  application is unsupported or fails.
- High Contrast: forced Solid, full-opacity surface/text/border, no backdrop blur, no decorative transparency.
- Reduced motion does not change material; it changes transition behavior.

Do not add `themeCycle.*` semantic colors in this feature. A future product requirement for independently
themeable notification colors must be a separate theme-system change covering schema, runtime, all themes,
documentation, galleries, and resolver performance.

### 10.2 External change while visible

- An unsupported/external theme change, Preferences apply/preview, settings hot reload, theme-file reload,
  language reload, or ring membership/order change hides the overlay immediately. A supported direct View ->
  Theme or Function Bar choice instead replaces the snapshot through the neutral transition. Do not leave stale
  neighbors visible.
- A pure owner activation/geometry change does not rebuild names.
- A system color/accent update while a stable `builtin/system` popup is visible may retheme the existing surface
  only if the snapshot generation and configured theme ID still match. Otherwise hide it.
- On language change, hide instead of mixing old snapshot names with new accessibility strings.

## 11. Animation, Disappearance Delay, and Timing

Animation is a product contract, not optional polish. Use the shared `Ui::AnimationDispatcher` only while pixels
are changing. Only one animation subscription may exist for this surface. During the fully visible delay, stop
the frame subscription and arm one balanced HWND dismissal timer; do not wake at animation cadence for 900 ms
just to compare a deadline.

Equivalent constants must be centralized and shared by runtime tests:

```cpp
constexpr uint64_t kAppearDurationMs              = 140u;
constexpr uint64_t kAdjacentTransitionDurationMs  = 160u;
constexpr uint64_t kDirectTransitionDurationMs    = 140u;
constexpr uint64_t kDisappearDelayMs              = 900u;
constexpr uint64_t kDisappearDurationMs           = 160u;
constexpr uint64_t kExitCancelRecoveryMs          = 80u;
```

Do not scatter literal timing values across paint, timer, UIA, and selftest code. Timing is monotonic and uses the
shared frame-runtime clock; wall-clock changes cannot shorten or extend the overlay.

### 11.1 State machine

```text
Hidden
  | accepted selection
  v
Appearing --140 ms--> FullyVisible --900 ms--> Disappearing --160 ms--> Hidden
  ^                        |                            |
  |                        | newer selection            | newer selection
  |                        v                            v
  +---------------- ContentTransition <----------------+
                           |
                           +--160 ms adjacent / 140 ms direct--> FullyVisible
```

Required states are `Hidden`, `Appearing`, `ContentTransition`, `FullyVisible`, and `Disappearing`, or an
equivalent explicit enum. Do not infer phase from opacity alone. The debug snapshot and deterministic tests must
expose the state, current transition progress, steady-visible start tick, disappearance deadline, and generation.

### 11.2 First appearance animation — 140 ms

Before `SW_SHOWNOACTIVATE`, prime the complete snapshot at enter progress zero, then render the first nonzero enter
progress before exposing the HWND so no opaque-black, uninitialized, or textless frame can be composited. The
mathematical transition still begins at opacity zero. Animate these values with `FastDecelerate` easing:

| Element | Start | End |
| --- | --- | --- |
| Whole surface opacity | 0.0 | 1.0 |
| Whole surface scale about its center | 0.94 | 1.00 |
| Whole surface vertical translation | +8 DIP | 0 DIP |
| Current-name opacity/scale | 0.0 / 0.97 | 1.0 / 1.00 |
| Previous-name offset/opacity | -8 logical DIP / 0.0 | 0 DIP / 1.0 |
| Next-name offset/opacity | +8 logical DIP / 0.0 | 0 DIP / 1.0 |

Neighbor-name animation begins 24 ms into the 140 ms enter phase so the center theme is perceived first. It still
finishes at 140 ms. “Logical” offsets mirror in RTL. The visible surface must remain inside the transparent
12 DIP animation gutter throughout the transform.

At 140 ms, snap all resolved values to their exact endpoints, stop the animation subscription, mark the surface
`FullyVisible`, and arm the 900 ms disappearance delay.

### 11.3 Adjacent theme transition — 160 ms

Previous/Next animation must explain movement through the ring instead of performing an unrelated fade:

- **Next:** the old current label moves from the center to the previous/top-leading position while scaling from
  `TitleLarge` toward `BodyStrong`; the old next label moves from bottom-trailing to the center while scaling up
  to `TitleLarge`; the old previous fades away; the new next fades/enters into bottom-trailing.
- **Previous:** mirror the choreography: the old current moves to next/bottom-trailing, the old previous moves to
  the center, the old next fades away, and the new previous enters at top-leading.
- Surface geometry, material, and HWND remain stationary. Only content transforms and crossfades.
- Use `PointToPoint` easing for moving labels and `FastDecelerate` for opacity/scale.
- Cache the outgoing and incoming DirectWrite layouts once at transition start. Per-frame work applies transforms,
  clips, and opacity; it must not recreate text layouts or interpolate font objects every frame.
- At the final frame, paint only the newest snapshot at canonical bounds/roles and release all outgoing layouts.

If an implementation cannot animate a cached layout between font roles without raster distortion, crossfade the
two cached endpoint layouts while translating them. Do not horizontally scale glyphs or create a new layout for
each intermediate size.

### 11.4 Direct menu/Function Bar selection transition — 140 ms

A direct View -> Theme or Function Bar choice may jump across several ring entries and therefore must not imply
adjacency:

- keep the surface stationary and fully opaque;
- fade the old three-label group from opacity 1 to 0 while translating it 4 DIP upward;
- fade the new group from 0 to 1 while translating it from +4 DIP to 0;
- scale only the new current label from 0.98 to 1.00;
- use `FastDecelerate` easing and release the outgoing layouts at completion.

The first eligible direct menu/Function Bar selection from `Hidden` uses the ordinary 140 ms appearance
animation, not a double enter plus content crossfade.

### 11.5 Disappearance delay — 900 ms after the latest settled content

- Start the delay only when the newest matching generation has completed its appearance/content transition and
  has been successfully presented at canonical full-opacity bounds.
- Record `steadyVisibleStartTickMs` and set `disappearDeadlineTickMs = steadyVisibleStartTickMs + 900`.
- Arm one HWND timer for the remaining duration. Treat it as one-shot: kill it immediately on `WM_TIMER` before
  changing state.
- Every new eligible selection kills an armed dismissal timer before updating content.
- A stale timer message must compare the captured/current deadline and generation; it cannot dismiss a newer
  snapshot.
- `SetTimer` failure logs once for that generation and falls back to the shared animation dispatcher as a
  deadline watcher without invalidating or rendering during the wait.
- There is no countdown, progress line, shrinking border, or opacity drift during the delay. The popup remains
  completely steady and readable.

### 11.6 Disappearance animation — 160 ms

At the valid disappearance deadline, enter `Disappearing` and animate the whole composed surface:

| Property | Start | End |
| --- | --- | --- |
| Opacity | 1.0 | 0.0 |
| Scale about center | 1.00 | 0.985 |
| Vertical translation | 0 DIP | -4 DIP |

Use `FastDecelerate` easing. Do not animate labels independently on exit. At 160 ms, snap opacity to zero, hide the
HWND without activation, stop the animation subscription, kill any dismissal timer, release outgoing transition
layouts, and enter `Hidden`. The lazily created HWND and device-independent resources may remain cached for reuse.

### 11.7 New input during appearance, transition, delay, or disappearance

- Update content on every accepted cycle; never queue snapshots.
- Keep the same HWND, DXUI host, shared device, and at most one animation subscription.
- During `Appearing`, preserve the current surface transform/opacity and transition directly to the newest content;
  do not jump back to opacity zero or replay the full enter animation.
- During `ContentTransition`, sample current visual progress, discard any older destination, and converge to the
  newest snapshot. Retain at most the currently visible outgoing state plus newest incoming state.
- During `FullyVisible`, kill the dismissal timer, perform the appropriate adjacent/direct content transition,
  then start a new 900 ms disappearance delay.
- During `Disappearing`, cancel the exit without hiding or recreating the HWND. Recover from the current opacity,
  scale, and translation to full visibility in at most 80 ms while beginning the newest content transition.
- A timer or animation callback carrying an older generation is stale and must be ignored.
- At every frame, the newest snapshot owns UIA and final state. Assistive technology must not be notified of an
  outgoing stale generation after newer input has been accepted.

Holding Next/Previous therefore keeps the popup visible continuously. Disappearance begins only 900 ms after
the final accepted repeat has finished its last 160 ms transition.

### 11.8 Reduced motion

When the resolved DXUI palette has `reducedMotion == true`:

- show the first complete frame immediately at full opacity, scale, and canonical position;
- replace repeated, menu-selected, or Function-Bar-selected content atomically with no translation, scale, fade,
  or crossfade;
- immediately mark the newest presented generation `FullyVisible` and restart the 900 ms disappearance delay;
- hide immediately at the deadline with no disappearance animation;
- use the balanced one-shot dismissal timer, not an animation subscription, during the delay;
- keep UIA notifications and all state semantics identical to ordinary motion.

### 11.9 Timer and animation lifetime

- Animation subscription exists only in `Appearing`, `ContentTransition`, or `Disappearing`.
- The dismissal timer exists only in `FullyVisible`.
- Entering one mechanism always stops the other first.
- `WM_NCDESTROY`, explicit hide, external invalidation, and application shutdown kill the dismissal timer and
  unsubscribe animation before clearing HWND/object state.
- A timer creation failure, stale `WM_TIMER`, skipped animation tick, or long UI-thread stall must converge from
  monotonic elapsed time to the correct current/final state; do not advance by a fixed amount per tick.

### 11.10 Suspension and immediate dismissal

Hide immediately and stop animation when:

- the main window is minimized, hidden, disabled for a modal owner, destroyed, or loses application activation;
- session shutdown/end begins;
- the app enters a fatal-error path;
- an external/non-user-initiated theme or language change invalidates the snapshot;
- popup render recovery reaches a terminal failure.

Escape does not dismiss the overlay because it remains non-focusable and Escape belongs to the active application
surface. Pointer input outside the visible rounded popup continues to the application underneath; pointer input
inside follows the explicit dismissal contract below.

### 11.11 Click/tap dismissal

The entire rounded popup surface is one dismiss target.

- A primary mouse click, touch tap, or pen primary activation whose press and release both resolve inside the
  visible rounded surface dismisses the popup.
- Transparent shadow/animation gutters and pixels outside rounded corners return `HTTRANSPARENT`; they are not
  part of the dismiss target and route normally to the application beneath.
- Inside the surface, return `HTCLIENT`, keep `WS_EX_NOACTIVATE`, return `MA_NOACTIVATE`, and show `IDC_HAND` so
  dismissibility is discoverable without taking focus.
- On primary down, set only a subtle pressed visual, capture through `DxUi::WindowHost`, and consume the event.
- On primary up inside the surface, release capture, cancel the disappearance timer/animation, hide immediately,
  and consume the release. Do not play the 160 ms exit animation after explicit dismissal.
- On primary up outside, capture loss, `WM_CANCELMODE`, owner deactivation, or teardown, release/cancel capture and
  clear the pressed visual. A canceled click does not invoke content beneath.
- Hiding on down is forbidden because the later button-up could reach a destructive underlying control. The
  complete down/up sequence belongs to the popup before it disappears.
- Right/middle/X-button clicks and wheel input do not invoke theme commands or show a context menu. They are
  consumed while over the visible surface so they cannot accidentally operate on obscured content.
- Clicking outside the popup does not dismiss it early; it interacts normally with the application while the
  popup continues toward its 900 ms automatic deadline.
- Explicit dismissal does not alter the selected theme, cycle index, focus HWND, active HWND, or foreground HWND.

The root UIA element exposes the same dismiss action through Invoke for assistive technology, while remaining out
of Tab order. Pointer interaction and UIA Invoke share one `DismissImmediately(Explicit)` path so cleanup cannot
drift.

## 12. Accessibility and Localization

### 12.1 UI Automation

The overlay is passive but not invisible to assistive technology:

- Expose one non-focusable root with an automation ID such as `ThemeCycleOverlay`, control type `Status`, and an
  Invoke action that dismisses the notification immediately.
- Expose logical Previous, Current, and Next text children when present, each with stable automation IDs.
- Root and children are not keyboard-focusable and do not enter Tab order. Only the root exposes Invoke; text
  children expose no interactive pattern.
- Publish a polite/most-recent live notification after the newly themed first frame is ready:
  `Theme changed to {0}. Previous: {1}. Next: {2}.`
- Use UIA notification processing semantics that coalesce rapid repeat into the most recent value rather than
  forcing a screen reader to speak every intermediate theme.
- Do not raise focus-changed events.
- The accessible names retain full, unellipsized display names.
- Root HelpText explains that clicking, tapping, or invoking dismisses the theme notification.
- UIA Invoke routes through the same immediate-dismiss path as a completed primary click and does not wait for or
  play the automatic disappearance animation.
- When `count == 1`, use the current-only notification variant and omit absent children.

If `DxUi::WindowHost` lacks a canonical status/live-region helper, add the narrow reusable helper to Common/DxUi
rather than implementing a private raw UIA provider that duplicates the host snapshot tree. Coordinate with the
active `DxUi_Uia_ContinuationBaton_2026-06-29.md` owner before modifying shared UIA internals.

### 12.2 Resource strings

Add source and satellite STRINGTABLE resources for accessibility/diagnostic grammar, for example:

```rc
IDS_THEME_CYCLE_OVERLAY_UIA_FULL
    "Theme changed to {0}. Previous: {1}. Next: {2}."
IDS_THEME_CYCLE_OVERLAY_UIA_CURRENT_ONLY
    "Theme changed to {0}."
IDS_THEME_CYCLE_OVERLAY_PREVIOUS_NAME
    "Previous theme: {0}"
IDS_THEME_CYCLE_OVERLAY_CURRENT_NAME
    "Current theme: {0}"
IDS_THEME_CYCLE_OVERLAY_NEXT_NAME
    "Next theme: {0}"
IDS_THEME_CYCLE_OVERLAY_DISMISS_ACTION
    "Dismiss theme notification"
IDS_THEME_CYCLE_OVERLAY_DISMISS_HELP
    "Click or tap anywhere on this notification to dismiss it."
```

Exact IDs may follow the resource file's active allocation block. Requirements:

- No user-facing hardcoded English in C++.
- Every variable resource uses positional placeholders in source argument order.
- Satellites preserve exact placeholder tokens and may reorder only for grammar.
- Decorative chevrons are not embedded in localized resource strings.
- Run `ResourceLocalizationContracts.Tests.ps1` after all source/satellite changes.
- Do not claim linguistic completion for Czech, Japanese, or Slovak without competent review; if source-identical
  fallback strings are temporarily required, route review to Operation Rosetta Lantern and keep the plan open.

### 12.3 RTL and bidi

- Follow the resolved UI flow direction from localization/DXUI.
- Keep each custom theme name as one isolated bidi text run so it cannot reorder neighboring decoration.
- Mirror leading/trailing corners and chevrons in RTL.
- UIA relation strings remain grammatical resources; do not build them by concatenation.

## 13. Rendering Failure and Recovery

Theme selection is primary; overlay feedback is best-effort.

- `D2DERR_RECREATE_TARGET`, DXGI device removed/reset, or lost shared device: discard device-dependent popup
  resources, let `WindowHost` recreate them, and retry one first-frame render while the request generation remains
  current.
- A transient failure does not show an alert and does not undo the theme.
- A second terminal failure for the same generation hides the overlay and emits one `Debug::Warning`/`Error`
  with HRESULT and stage. Do not log per animation frame.
- Creation/attach failure enters a suppressed-until-next-theme-change state. The next eligible cycle may retry;
  do not spin or recreate repeatedly inside one key repeat burst.
- Never fall back to GDI, a native tooltip, or an opaque black rectangle.
- The popup render path must not parse theme expressions, read files, enumerate themes, acquire settings locks,
  perform I/O, or allocate per frame. Theme resolution and text-layout creation occur on snapshot/palette change.

## 14. Performance Contract

This feature touches input-to-visible latency, synchronous theme application, DirectWrite layout, animation, and
shared rendering. Performance evidence is mandatory from the first implementation slice.

### 14.1 Protected scenario

**Scenario name:** warm keyboard theme-cycle burst with the centered overlay visible.

The deterministic fixture cycles a fixed built-in ring at least 240 accepted keyboard invocations in a test-enabled
Release build, keeps the main window at a fixed size/DPI, includes both directions and wraparound, and separately
runs a 1,000-update retention stress without file-backed tracing inside the measured paint/present path.

Keyboard repeat remains the percentile/performance stress because it is the highest-frequency path. Add a
separate deterministic one-shot UI fixture covering direct built-in/custom menu selection, menu Previous/Next,
Function Bar click, and explicit click dismissal; it is correctness/first-visible evidence, not a source of
synthetic rapid pointer repetition.

### 14.2 Metrics

Emit bounded event-level metrics under one `theme.cycle.*` family:

| Metric | Unit/grain | Meaning |
| --- | --- | --- |
| `theme.cycle.apply_us` | us per accepted command | Existing settings ID update plus `ApplyAppTheme`, excluding overlay work. |
| `theme.cycle.overlay.update_us` | us per accepted command | Snapshot handoff, text layout/update, geometry, and synchronous initial/update render request. |
| `theme.cycle.input_to_visible_us` | us per accepted command | Supported keyboard/menu/Function Bar selection dispatch timestamp through successful present of the matching generation. |
| `theme.cycle.overlay.window_create_count` | count per case | Number of native popup-window creations. |
| `theme.cycle.overlay.reuse_count` | count per case | Updates served by the existing popup HWND. |
| `theme.cycle.overlay.layout_create_count` | count per command or aggregate case | DirectWrite layouts created after a snapshot change; must be zero per animation-only frame. |
| `theme.cycle.overlay.render_count` | count per command/case | Frames rendered for appearance/content/disappearance animations; the steady delay must add none. |
| `theme.cycle.overlay.dismiss_delay_ms` | ms per completed visible lifetime | Observed steady-visible time before disappearance begins; target is exactly 900 ms subject to timer scheduling lateness, never early. |
| `theme.cycle.overlay.dismiss_timer_arm_count` | count per case | Balanced disappearance-delay arms; repeat must replace rather than accumulate timers. |
| `theme.cycle.overlay.stale_timer_ignored_count` | count per case | Old timer deliveries rejected by deadline/generation checks. |
| `theme.cycle.overlay.dropped_generation_count` | count per case | Intermediate snapshots superseded before present during burst; final generation must still present. |

Metric emission is gated through the existing `Debug::Perf` capture switch. Do not write per-frame selftest trace
lines or one JSONL row per glyph/draw call in the measured path.

### 14.3 Gates

Final performance claims use same-machine, same-suite, test-enabled Release evidence with at least 200 samples for
each p95 latency metric:

- `theme.cycle.apply_us` candidate p95 is no worse than the instrumentation-only baseline by more than the larger
  of 10% or 2 ms.
- Warm `theme.cycle.overlay.update_us` p95 is at most 8 ms on the evidence machine.
- Warm `theme.cycle.input_to_visible_us` p95 is at most 100 ms and max at most 250 ms on the deterministic fixture.
- One burst creates at most one popup HWND after warmup; subsequent accepted commands increment reuse, not create.
- No DirectWrite layout is created on an animation frame without a snapshot, DPI, language, or palette change.
- No animation frames are requested or rendered during the 900 ms fully visible disappearance delay.
- A completed ordinary lifetime never begins disappearance before 900 ms of steady visibility; scheduler delay
  may make it late but cannot make it early.
- Repeat updates leave at most one armed dismissal timer, and teardown leaves none.
- After the 1,000-update stress and final hide, animation subscriptions return to zero, popup count is at most one
  retained hidden HWND, USER/GDI handles do not grow with iteration count, and retained snapshot history is bounded
  to current plus at most one outgoing transition.

If the environment cannot meet an absolute timing gate for reasons outside the feature, archive the evidence and
record the exact blocker rather than weakening the gate silently. Do not claim a win from Debug timings.

### 14.4 Baseline sequence

1. Land or temporarily apply instrumentation without showing the overlay.
2. Run the deterministic theme-cycle perf case and archive the baseline.
3. Implement the popup without changing metric definitions.
4. Run and archive the candidate on the same machine/build/suite.
5. Analyze with `Tools/Show-PerfRuns.ps1`, using `-FailOnQuality` for p95 claims.
6. Preserve `results.json`, `trace.txt`, `perf_metrics.jsonl`, build label, and environment metadata under
   `Specs/TestRuns/<MachineHash>/Commands/<RunId>/`.

## 15. Deterministic Test Contract

### 15.1 Pure/state tests

Add focused coverage for:

- exact built-in order and custom file/settings append order;
- Next and Previous from every built-in entry;
- forward and backward wraparound;
- missing-current fallback;
- owned snapshot strings surviving source-vector/settings mutation;
- built-in localized name lookup and custom empty-name ID fallback;
- one/two-theme defensive geometry semantics;
- long UTF-16, Japanese, RTL, combining-mark, and supplementary-character names;
- content-transition state under rapid generation replacement;
- reduced-motion enter/update/exit state;
- deterministic fake-clock timing checkpoints for first appearance: `Appearing` before 140 ms, `FullyVisible` at
  140 ms, no disappearance before 1,040 ms, `Disappearing` at/after the valid 1,040 ms deadline, and `Hidden` at
  1,200 ms after the first presented enter frame when callbacks are delivered exactly on time;
- repeat timing: an update accepted during the delay kills the old timer, completes its 160 ms adjacent transition,
  remains steady for a new full 900 ms, and cannot be hidden by delivery of the old timer message;
- exit cancellation: input during the 160 ms disappearance samples current opacity/transform, returns to full
  visibility within 80 ms, presents the newest generation, and receives a fresh delay;
- disappearance-timer failure fallback and stale generation/deadline rejection;
- deterministic geometry at 96, 144, 192, and 240 DPI.

### 15.2 Commands selftests

Add named cases with equivalent coverage to:

- `theme_cycle_overlay_keyboard_shows_current_and_neighbors`
- `theme_cycle_overlay_previous_wraps_and_mirrors_transition`
- `theme_cycle_overlay_reuses_window_during_repeat`
- `theme_cycle_overlay_appearance_delay_and_disappearance_timeline`
- `theme_cycle_overlay_new_input_cancels_disappearance_without_flash`
- `theme_cycle_overlay_menu_previous_next_show_directional_state`
- `theme_cycle_overlay_menu_direct_builtin_and_custom_show_neutral_state`
- `theme_cycle_overlay_function_bar_click_shows_and_same_theme_is_noop`
- `theme_cycle_overlay_click_or_tap_dismisses_without_activation_or_clickthrough`
- `theme_cycle_overlay_suppresses_noop_and_non_user_selection_sources`
- `theme_cycle_overlay_preserves_focus_activation_and_pointer_routing`
- `theme_cycle_overlay_repositions_and_rethemes_across_dpi`
- `theme_cycle_overlay_honors_reduced_motion_and_high_contrast`
- `theme_cycle_overlay_recovers_from_device_loss`
- `theme_cycle_overlay_uia_reports_most_recent_generation`
- `theme_cycle_overlay_perf`
- `theme_cycle_overlay_teardown_stress`

Each live-window test must assert, as applicable:

- configured ID, resolved current name, previous name, next name, and direction;
- same HWND reused and monotonically increasing generation;
- popup rectangle centered over the main client within one physical pixel;
- popup remains within client/work-area clamps;
- a complete snapshot is primed before show and the first exposed enter frame has nonzero material/text content
  without mouse movement;
- popup palette matches the newly selected theme;
- current typography is larger than neighbor typography and rectangles match the corner/center contract;
- first appearance uses the 140 ms opacity/scale/translation contract and reaches exact endpoint values;
- adjacent Next/Previous uses the 160 ms ring-position choreography; direct menu/Function Bar selection uses the
  neutral 140 ms crossfade;
- the popup stays fully visible for at least 900 ms after the newest transition, requests no animation frames
  during that delay, and then completes the 160 ms disappearance animation;
- repeat, menu, or Function Bar input during the delay/exit cancels disappearance, reuses the HWND, and cannot
  flash hidden;
- focus HWND, active HWND, foreground HWND, and focused FolderView selection are unchanged;
- hit testing is `HTCLIENT` only over the visible rounded surface and transparent over gutters/corners;
- a primary down/up inside the surface is captured, consumed, and dismisses immediately without focus/activation,
  while a canceled/outside release does not dismiss or invoke obscured application content;
- View -> Theme and Function Bar built-in/custom/Previous/Next choices show the correct snapshot, while same-theme
  no-op, parameterized ordinary shortcut, programmatic, Preferences, and hot-reload changes suppress or hide;
- High Contrast uses solid full-opacity visuals and no blur;
- reduced motion has no intermediate opacity/translation/scale state, preserves the full 900 ms delay, and
  hides immediately when that delay expires;
- UIA has stable IDs, full names, no focusability/Tab entry, one root dismiss Invoke action, and most-recent live
  notification;
- destroy/minimize/deactivate/unload stops all animation callbacks and leaves no stale HWND access.

Provide an `ENABLE_TESTS` debug snapshot with owned strings, IDs, generation, visibility/phase, transition
progress, `steadyVisibleStartTickMs`, `disappearDeadlineTickMs`, current opacity/scale/translation, dismissal-timer
armed state/count, pointer-hover/pressed/captured state, explicit-dismiss count/reason, HWND creation/reuse counts,
rects in DIP and pixels, DPI, palette flags, text roles, ellipsis flags, animation subscription state,
render/present counts, and last failure HRESULT/stage. Tests must not scrape pixels for facts that can be exposed
as deterministic semantic state.

### 15.3 Visual/golden coverage

Closeout uses deterministic semantic and bitmap assertions rather than a new dedicated PNG baseline:

- the Commands rendering case captures stable appearance, an adjacent transition, High Contrast, and device-loss
  recovery, and asserts content geometry, transparent gutters, responsive placement, and layout stability;
- pure placement coverage exercises 96, 144, 192, and 240 DPI;
- the shared DXUI Rendering suite covers Solid, Mica, Mica Alt, Acrylic, pressed, and High Contrast pixel policy;
- the repository's existing Light, Dark, High Contrast, and popup visual baselines remain green.

No new theme-overlay PNG golden is committed. This is an explicit closeout caveat, not an accidental omission:
semantic state and owned bitmap assertions carry feature facts, while the existing shared visual baselines carry
renderer-wide regression coverage.

### 15.4 Localization contracts

Run the resource placeholder/coverage tests and validate that all shipped satellites contain placeholder-equivalent
entries. Any unreviewed source-identical translations remain an explicit closeout blocker or routed linguistic
review; they are not silently declared translated.

## 16. Implementation Phases

### Phase 0 — Drift, ownership, and baseline

- Re-run the current-state anchors and note drift from commit `148f270dec`.
- Coordinate before touching shared UIA files owned by the active DxUi UIA baton.
- Confirm no shared transient-surface helper has landed since this plan.
- Add the metric seam and deterministic baseline fixture without showing a popup.
- Archive test-enabled Release baseline evidence.

**Gate:** current theme-cycle behavior is unchanged, existing `theme_cycle_commands` remains green, and baseline
metrics/archive exist.

### Phase 1 — Pure cycle snapshot and invocation origin

- Introduce the narrow invocation-source enum/plumbing.
- Route keyboard Previous/Next, View -> Theme commands, and Function Bar clicks with explicit origin.
- Route direct built-in/custom menu or Function Bar choices through the same snapshot/show helper with neutral
  direction.
- Build the owned snapshot from the same ring/index used for selection.
- Centralize display-name resolution.
- Add pure and command-origin tests.

**Gate:** keyboard, theme-menu, function-bar pointer, other `WM_COMMAND`, parameterized direct selection, and
programmatic paths are distinguishable without recent-key heuristics or thread-local latches; no popup exists yet.

### Phase 2 — Shared DXUI transient material

- Reuse an existing shared primitive if available; otherwise add the narrow shared popup-surface primitive.
- Cover Solid, Mica, MicaAlt, Acrylic, Rainbow, High Contrast, opacity, transform, DPI, and device loss.
- Preserve existing menu/ComboBox visuals if consolidating material code.

**Gate:** focused DxUi tests and existing popup/menu/ComboBox gallery tests are green.

### Phase 3 — Passive overlay window and static layout

- Add composition-root ownership and lazy `ThemeCycleOverlayWindow` creation.
- Implement no-activate styles, selective rounded-surface hit testing, click/tap dismissal, transparent gutters,
  and owner-relative placement.
- Build previous/current/next DXUI content and initial-frame priming.
- Apply the newly resolved palette and responsive layout.
- Add first-frame, focus, input, geometry, DPI, High Contrast, long-name, and device-loss tests.

**Gate:** one keyboard cycle, one direct theme-menu choice, and one Function Bar click show correct complete frames
without changing application focus; explicit popup dismissal consumes only its own primary activation.

### Phase 4 — Repeat, animation, and dismissal

- Add generation replacement, bounded outgoing state, the 140 ms appearance, 160 ms adjacent choreography,
  140 ms direct-selection transition, 900 ms balanced disappearance timer, 160 ms disappearance, exit reversal,
  immediate explicit dismissal, balanced pointer capture, shared animation subscription, reduced-motion path,
  owner deactivation/minimize handling, and external-change invalidation.
- Add repeat and teardown stress tests.

**Gate:** rapid forward/backward auto-repeat always converges to the newest configured theme, uses one HWND and at
most one subscription/one dismissal timer, never begins disappearance before the newest 900 ms steady delay,
consumes a dismissal click without activating or invoking obscured content, and leaves no capture, timer, or
callback after teardown.

### Phase 5 — Accessibility and localization

- Add status/live-region semantics and most-recent notification.
- Add positional accessible names and full unellipsized strings.
- Add source/satellite resources and RTL behavior.
- Run UIA, resource-contract, and competent linguistic review gates.

**Gate:** Narrator/UIA automation receives the newest current/previous/next state without focus changes or a speech
backlog during repeat.

### Phase 6 — Performance, full verification, and closeout

- Run the 240-sample Release candidate and 1,000-update retention stress.
- Compare to the Phase 0 same-machine baseline and enforce quality gates.
- Run focused Debug, focused test-enabled Release, and final Full suite.
- Update durable UI/DXUI/theme/user documentation.
- Move this file to `Specs/Plans/Done/` and update/remove its WIP queue row only after every acceptance item closes.

## 17. Expected File Impact

The implementation should remain within this boundary unless drift proves a documented need:

- `RedSalamander/RedSalamander.cpp` — explicit invocation source, snapshot selection, overlay ownership calls,
  owner geometry/activation/teardown notifications.
- `RedSalamander/Ui/ThemeCycleOverlayWindow.h/.cpp` — passive owned popup, state, layout, animation, debug snapshot.
- `RedSalamander/RedSalamander.vcxproj` and `.filters` — explicit source/header entries if required.
- `Common/DxUi/DxUi.h` plus a focused implementation file — shared transient surface and possibly a narrow live-
  notification helper, only when no canonical helper exists.
- `Common/Common.vcxproj` and `.filters` — explicit shared source/header entries if required.
- `RedSalamander/RedSalamander.rc`, `resource.h`, and all shipped satellite `.rc` files — localized accessible
  strings.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp` or a focused new Commands selftest part —
  state, live-window, UIA, perf, and teardown cases.
- `Tests/DxUiTests/*` — shared material, capture, transform, High Contrast, DPI, and device-loss tests.
- `Specs/Testing/*` only if a new metric budget/analyzer contract is added.
- `Specs/TestRuns/*` — archived baseline/candidate evidence.
- closeout documentation named at the top of this plan.

No settings schema, theme JSON5 format, `AppTheme` color field, plugin ABI, file-system code, or shortcut-default
change is expected.

## 18. Verification Commands

Use the exact active runner syntax at implementation time. The intended gates are:

```powershell
./build.ps1 -ProjectName RedSalamander

./Tools/Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast `
  -CaseFilter theme_cycle_commands,theme_cycle_overlay

./Tools/Tests/ResourceLocalizationContracts.Tests.ps1

./Tools/Run-AllTests.ps1 -Suite Full
```

For final performance evidence, use a test-enabled Release build, the focused perf case, the runner-owned archive,
and `Tools/Show-PerfRuns.ps1 -CompareRun ... -FailOnQuality`. Record the actual baseline/candidate commands and
archive paths in this plan during implementation; do not leave placeholder evidence at closeout.

## 19. Acceptance Checklist

### Product and interaction

- [x] Keyboard Previous/Next and View -> Theme choices show the popup.
- [x] Function Bar theme clicks show the popup when they change the theme.
- [x] Direct built-in/custom menu choices use a neutral transition and the selected theme's exact ring neighbors.
- [x] Selecting the already active menu/Function Bar theme is a no-op and does not show/restart the popup.
- [x] The newly active theme is large and centered.
- [x] The exact previous neighbor is small at logical top-leading.
- [x] The exact next neighbor is small at logical bottom-trailing.
- [x] Forward and backward wraparound are correct.
- [x] First show completes the specified 140 ms appearance animation.
- [x] Adjacent Previous/Next transitions spatially promote the correct neighbor over 160 ms.
- [x] Direct menu/Function Bar selection uses a neutral 140 ms content transition.
- [x] The newest content remains completely steady and visible for 900 ms before disappearance begins.
- [x] Disappearance completes the specified 160 ms opacity/scale/translation animation.
- [x] Repeated shortcut/menu selections cancel pending disappearance, restart the full delay after their newest
  transition, and reuse one HWND without a hidden-frame flash.
- [x] A primary click/tap anywhere inside the visible rounded surface consumes the complete press/release and
  dismisses immediately without playing the automatic exit animation.
- [x] Transparent gutters/corners remain click-through; right/middle/wheel input over the visible surface cannot
  operate obscured application content.
- [x] The popup never takes focus, activation, or a Tab position; pointer capture is balanced only for dismissal.
- [x] Non-theme/unsupported Function Bar commands, parameterized direct-theme shortcuts, programmatic selection,
  startup, Preferences, hot reload, and system changes do not incorrectly show it.

### Visual and rendering

- [x] The surface is centered over the main client and responsive at small sizes.
- [x] Light, Dark, Rainbow, custom, and High Contrast visuals pass contrast review.
- [x] The popup uses the newly applied theme.
- [x] Long/localized/RTL names fit or ellipsize without overlap; UIA retains the full name.
- [x] Initial frame is complete before show.
- [x] DPI/monitor/resize changes reposition and relayout while visible.
- [x] Reduced motion eliminates motion/crossfade while retaining the 900 ms readable delay.
- [x] Device loss recovers or suppresses feedback without undoing the theme.
- [x] No GDI or blocking I/O enters the paint path.

### Accessibility and localization

- [x] Status/live-region UIA reports current, previous, and next using most-recent coalescing.
- [x] No focus event or Tab stop is exposed; the non-focusable root dismiss `Invoke` is the only interactive UIA
  pattern.
- [x] All grammar-bearing text is in STRINGTABLE resources with positional placeholders.
- [x] Satellite placeholder contracts pass; linguistic review status is honest.
- [x] RTL placement and chevrons mirror while logical previous/next semantics remain correct.

### Performance and lifetime

- [x] Instrumentation-only same-machine Release baseline is archived.
- [x] Candidate Release archive contains at least 200 quality samples for p95 claims.
- [x] `theme.cycle.apply_us` regression gate passes.
- [x] Overlay update and input-to-visible budgets pass or carry an explicit unresolved blocker.
- [x] Delay telemetry proves no early disappearance and no animation-frame rendering during the steady delay.
- [x] The 1,000-update stress proves bounded snapshots/layouts/handles and zero subscription after hide/destroy.
- [x] The Full attempt's unrelated failures are classified below as a reproducible SettingsStore/current-branch
  regression and downstream UI-automation fallout. Aurora's 11 cases passed inside that attempt and again in a
  clean focused run; the unrelated regression is accepted as outside this plan's ownership.

### Documentation closeout

- [x] Durable behavior is merged into authoritative UI/DXUI/theme/user docs.
- [x] Actual commands, archives, metrics, screenshots, and caveats replace planning placeholders.
- [x] The plan is moved to `Specs/Plans/Done/` and the WIP index is reconciled.

## 20. Stop Conditions

Stop and request a maintainer decision instead of silently changing the design if:

- product wants the overlay for additional theme-change sources beyond keyboard Previous/Next, View -> Theme, and
  clicked Function Bar theme commands;
- product wants the popup to be interactive or configurable;
- a setting or new theme semantic color key becomes necessary;
- a shared DXUI/UIA change conflicts with the active `DxUi_Uia_ContinuationBaton_2026-06-29.md` work;
- the only proposed implementation requires focus activation, a modal loop, GDI fallback, or a private D3D device;
- ring ordering must change or file/settings custom-theme ordering cannot remain authoritative;
- competent translation review is unavailable but completion would require claiming the new strings translated;
- baseline/candidate metric definitions are not comparable;
- existing unrelated worktree edits overlap a required file in a way that cannot be preserved safely.

These conditions block completion of the affected decision; they do not authorize dropping tests, accessibility,
performance evidence, or authoritative closeout updates.

## 21. Implementation Checkpoint — 2026-07-21

This section is a durable resume point. The implementation is intentionally still WIP: do not move this plan to
Done until the remaining verification and documentation gates below are complete.

### Implemented and saved in the working tree

- Added `ThemeCycleOverlayWindow` as one reusable, owned, non-activating DXUI composition popup. It implements the
  adjacent/direct phase machine, 140/160/140 ms transitions, the 900 ms fully steady interval, the 160 ms automatic
  exit, reduced-motion behavior, responsive placement, newly applied theme styling, cached DirectWrite layouts,
  no-op preservation, and monotonic generation/deadline validation.
- Routed keyboard Previous/Next, View -> Theme direct choices, and clicked Function Bar theme commands through an
  explicit invocation-source contract. A direct selection of the active theme remains a true no-op. Other theme
  application sources suppress an existing popup instead of producing false feedback.
- Implemented whole-surface primary press/capture/release dismissal without focus or activation. Transparent
  gutters/corners are hit-test transparent; secondary pointer input over the surface is consumed.
- Added localized accessibility resources in English, Czech, French, Japanese, and Slovak. The popup exposes a
  non-focusable Status root, stable automation IDs, three semantic labels, a non-focusing Invoke dismissal action,
  and a coalesced UIA notification after content settles.
- Extended DXUI narrowly with Status-role semantics, stable automation IDs, custom non-focusing Invoke support, a
  Hand cursor, and a live-notification helper. The Alert Overlay was reviewed but not migrated: it is a blocking,
  backdrop-owning custom renderer rather than a `WindowHost` transient status surface, so migration would be a
  separate behavioral rewrite and is not required by this feature.
- Added focused Commands selftests for exact timing boundaries, menu no-op/HWND reuse, Function Bar invocation and
  real mouse dismissal, UIA, reduced motion, 1,000-update retention, and a 240-command performance scenario. Added
  a DXUI accessibility test for Status-root child exposure and non-focusing Invoke.
- Added telemetry for window creation, text-layout creation, presentation, input-to-visible latency, dismiss delay,
  and the existing theme-apply duration. Animated opacity reuses stable theme brushes via layers so animation does
  not grow the brush cache.

### Evidence already collected

- Instrumentation-only Release baseline: `Specs/TestRuns/7d3a1247382a/Commands/2026-07-21_141058`.
- Baseline `theme.cycle.apply_us`, 240 quality samples: p50 1,148 us, p95 2,154 us, p99 3,996 us, max 4,346 us;
  quality PASS.
- The focused Release DXUI Accessibility suite passed after the Status-fragment navigation fix.
- Earlier focused application runs passed keyboard timing, menu no-op/HWND reuse, Function Bar dismissal, and the
  1,000-update retention case. These must be rerun against the final binary because the click test and DXUI
  accessibility implementation changed afterward.
- A Release application rebuild ended and produced `.build/x64/Release/RedSalamander.exe` at 2026-07-21 15:13:15,
  but its final command result was lost during session compaction. Treat the build as unverified and rebuild once.

### Exact resume sequence

1. Rebuild the test-enabled Release application and require zero errors and warnings.
2. Run every `theme_cycle_overlay_*` Commands case. In particular, confirm the new real
   `WM_LBUTTONDOWN`/`WM_LBUTTONUP` capture path, the application UIA case after relinking DXUI, and reduced motion.
3. If a `SetTimer` failure fallback has not yet been completed, add a dispatcher-backed deadline watcher that does
   not invalidate/render during the 900 ms steady interval; add a focused test for that fallback.
4. Add/confirm direct custom-theme coverage and count-one/count-two ring edge coverage. Add visual/capture assertions
   for long names, responsive geometry, and stable layout/resource counts if not already present.
5. Run the candidate 240-command perf case, archive it through the standard runner, compare it with the baseline
   using `Show-PerfRuns.ps1 -FailOnQuality`, and record apply, update/present, input-to-visible, and delay results here.
6. Run the resource-localization contract, focused DXUI tests, and the full repository suite. Classify any unrelated
   failures with evidence; do not check off a gate based on an older binary.
7. Merge the durable behavior into the authoritative UI/DXUI/theme documentation, record actual commands/archive
   paths/caveats, reconcile the WIP index, and only then move this plan to `Specs/Plans/Done/`.

No commit was created at this checkpoint. All source, resource, project, selftest, DXUI test, and plan edits are
saved locally in the current working tree.

## 22. Pause Checkpoint — 2026-07-21, After Timer/Render/Shared-Surface Work

This checkpoint supersedes the build and resume details in section 21. The user asked to save all progress and
continue later. The plan deliberately remains in `Specs/Plans/WIP/`; the latest source edits are saved locally but
have not yet passed a rebuild, final candidate measurement, or closeout review.

### Additional implementation now saved

- Added the shared DXUI `TransientSurfaceOptions`, `ResolveTransientSurfaceBackdrop(...)`, and
  `PaintTransientSurface(...)` primitive. The theme-cycle popup now uses it for fill, border, shadow, pressed
  feedback, High Contrast policy, and system-backdrop selection. Focused shared rendering tests and durable DXUI
  documentation are still required.
- Expanded the overlay debug snapshot with owned IDs/names, content rectangles, DPI, High Contrast, capture,
  layout/render/window-reuse counts, timer state, stale-deadline counts, dropped-generation counts, and device-loss
  hooks. Added bitmap capture, forced dismiss-timer failure, device-loss simulation, and direct-theme debug entry
  points for focused tests.
- Extracted the ring snapshot builder into the UI module as a pure `std::span`-based function. The main theme ring
  now copies both IDs and resolved display names before applying the theme, protecting the overlay snapshot from
  later settings/theme mutations. A final compile-safety pass added `<cstddef>` and a display-name helper forward
  declaration; this last refactor has not yet been rebuilt.
- Made every accepted update synchronously render a complete generation before recording input-to-visible time.
  Added exact create/reuse, render, layout-create, dropped-generation, update-duration, and deadline-arm telemetry.
  Removed the animation-settle layout invalidation so animation-only frames do not recreate text layouts.
- Added a physical-deadline fallback for `SetTimer` failure using a WIL-owned thread-pool timer. Its callback only
  posts the centralized `WndMsg::kThemeCycleOverlayFallbackDeadline` cookie; teardown cancels and waits for the
  callback, and the UI thread revalidates cookie, generation, and monotonic deadline before acting.
- Added work-area clamping, `WS_CLIPCHILDREN | WS_CLIPSIBLINGS`, logical top-leading/bottom-trailing RTL geometry,
  and shared material/backdrop use. Added real capture/click dismissal and retained one non-activating popup HWND.

### Last green evidence before the latest saved edits

- Full required Release rebuild succeeded with zero errors using:

  ```powershell
  $env:CL='/MP4 /Zf'
  ./build.ps1 -Configuration Release -Platform x64 -Rebuild -MaxCpuCount 1
  ```

  The log is `.build/logs/msbuild-20260721_175514_132.log`. It reported 13 pre-existing `/Wall` unused-helper/test
  warnings outside this feature.
- Test-enabled Release application build succeeded with zero warnings and zero errors using:

  ```powershell
  $env:RSBuildEnableTests='true'
  $env:CL='/MP4 /Zf'
  ./build.ps1 -Configuration Release -ProjectName RedSalamander -MaxCpuCount 1
  ```

  The log is `.build/logs/msbuild-20260721_181116_080.log`.
- Against that test-enabled binary, all seven then-existing focused Release cases passed individually between run
  timestamps `20260721T162406` and `20260721T162434`:
  `theme_cycle_overlay_keyboard_timing`, `theme_cycle_overlay_menu_noop_reuse`,
  `theme_cycle_overlay_function_bar_dismiss`, `theme_cycle_overlay_accessibility`,
  `theme_cycle_overlay_reduced_motion`, `theme_cycle_overlay_retention`, and `theme_cycle_overlay_perf`.
  These passes predate the latest shared-surface, pure-ring, synchronous-render, telemetry, and timer-fallback edits,
  so they are evidence of the earlier implementation only and must not be used as final sign-off.
- The archived instrumentation baseline remains
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-21_141058`: 240 quality samples for
  `theme.cycle.apply_us`, p50 1,148 us, p95 2,154 us, p99 3,996 us, max 4,346 us, quality PASS.
- `git diff --check` passed immediately before this checkpoint. Line-ending notices are informational CRLF
  normalization warnings. No commit or staging operation was performed.

### Exact state to assume when resuming

- The working tree intentionally contains modified tracked files plus the untracked overlay header/source and
  focused Commands selftest. Preserve all of them; do not reset or discard unrelated user changes.
- The last two compile-safety edits were not built. First inspect the diff, then build before extending tests.
- Alert Overlay review is complete for this feature: no migration is required. Revisit it only as a separate
  behavior/design task, not as a prerequisite for this popup.
- Do not move this plan to Done until the authoritative UI/DXUI documentation, candidate perf archive/comparison,
  focused and full tests, localization contract, and all relevant acceptance checks are complete.

### Resume in this order

1. Inspect `git status`, `git diff --check`, the pure ring builder, and the timer-fallback lifetime/cookie paths.
   Build the test-enabled Release application; fix compile errors before adding more behavior.
2. Add deterministic coverage for zero/one/two-entry rings, wraparound, owned Unicode names, custom themes and empty
   name fallback, forced `SetTimer` failure, stale deadlines, no layout creation on animation-only frames, no render
   during the 900 ms steady interval, RTL rectangles, capture/hit testing, High Contrast, responsive placement,
   device loss, and bitmap pixels/transparent gutters.
3. Add shared DXUI tests for transient-surface backdrop mapping and ordinary/High Contrast capture. Rerun DXUI
   Accessibility and the applicable rendering suite.
4. Extend the perf case so the candidate archive contains at least 240 update and input-to-visible quality samples
   plus aggregate create/reuse/render/layout/timer/stale/dropped counters. Run the candidate separately and compare
   it to the archived baseline with `Tools/Show-PerfRuns.ps1 -FailOnQuality`.
5. Run every `theme_cycle_overlay_*` case, the resource-localization contract, and the full repository suite on the
   final binary. Investigate the previously accumulating disk-audit issue count rather than treating it as green.
6. Write the authoritative theme-cycle-overlay UI contract, update DXUI/theme/user documentation, replace all
   placeholder evidence in this plan, reconcile the WIP index, and move this file to `Specs/Plans/Done/` only after
   every applicable acceptance gate is genuinely satisfied.

## 23. Final Closeout — 2026-07-21

Sections 21 and 22 are retained as historical resume checkpoints. This section supersedes their outstanding-work
lists: Aurora is implemented, its owned gates are complete, durable behavior is authoritative, and the plan is
archived under `Specs/Plans/Done/`.

### Final implementation

- `RedSalamander/Ui/ThemeCycleOverlayWindow.*` owns one lazy, reused, non-activating DXUI composition popup with
  responsive client-centered/work-area-clamped placement, exact previous/current/next ring content, RTL mirroring,
  newly applied theme material, High Contrast policy, cached layouts, device-loss recovery, and balanced capture.
- Keyboard Previous/Next, direct View -> Theme choices, and clicked Function Bar theme commands carry an explicit
  invocation source. Same-theme direct selection is a true no-op; programmatic, startup, Preferences, hot-reload,
  and system changes suppress feedback.
- First appearance is 140 ms, adjacent promotion is 160 ms, direct selection is a neutral 140 ms transition, the
  settled surface stays fully steady for 900 ms, and automatic exit is 160 ms. New input retargets from the current
  visual state; reduced motion keeps the 900 ms readable interval without motion.
- Primary press/release anywhere on the visible rounded surface dismisses immediately. Transparent gutters remain
  click-through, secondary input cannot activate obscured content, and the popup never activates or enters Tab
  order.
- Narrow shared DXUI work added transient-surface material policy, Status semantics, stable automation IDs,
  non-focusing Invoke, live notifications, Hand cursor support, and capture-loss reset notification. Alert Overlay
  was reviewed and deliberately not migrated because its blocking/backdrop-owning behavior is a different contract.

### Final build and correctness evidence

- Final clean test-enabled Release rebuild:
  `.build/logs/msbuild-20260721_215037_077.log` — success, 0 errors, 10 pre-existing ViewerPETests unused-code
  warnings, 15:40.
- Final focused Commands run:
  run id `20260721T200658Z-39976-375f210ef00440f598c2f0974b6a9b99` and archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-21_220734` — 11 passed, 0 failed, 0 skipped, zero disk-audit issues.
- Affected direct DXUI suites `WindowHost`, `Theme`, `Rendering`, and `Accessibility` all passed against the final
  binary. The shared rendering suite includes the repository's existing visual baselines.
- `Tools/Tests/ResourceLocalizationContracts.Tests.ps1` passed all five contracts for English and shipped
  satellites.
- Shutdown AV diagnosis and the fixed external-HWND-destroy/recreate contract are archived under
  `Specs/TestRuns/7d3a1247382a/Continuation/2026-07-21_theme-cycle-overlay-shutdown-av/`.
- `git diff --check` is part of final handoff validation. No commit or staging operation is part of this closeout.

### Performance evidence

- Instrumentation-only baseline:
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-21_141058`.
- Final candidate:
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-21_195731`.
- `theme.cycle.apply_us` comparison used 240/240 quality samples: p95 2,154 us -> 1,812 us (-15.9%), quality PASS.
  Candidate raw distribution was p50 1,280 us, p95 1,821 us, p99 2,058 us; the 8 ms p95 budget passed.
- `theme.cycle.input_to_visible_us` used 240 quality samples: p50 2,911 us, p95 4,329 us, p99 4,914 us; the 100 ms
  p95 and 250 ms maximum budgets passed.
- Candidate aggregate behavior: one HWND creation, 239 reuses, 720 renders, 1,437 layout creations, five timer arms,
  zero stale deadlines, and zero dropped generations. The steady 900 ms phase performs no animation-frame render.

### Repository-wide Full-suite limitation

- A clean `-Suite Full -SkipBuild -Configuration Release` attempt used run id
  `20260721T185108Z-73008-3e7a41dd588248e2ac5c8d5298b0f781`. Its Commands result reached 367 passed,
  23 failed, and one skipped. Every one of Aurora's 11 cases passed within that run.
- The earliest failure was `settings_store_search_roundtrip`:
  “Failed to save dirty search settings for sanitization test.” It reproduced alone after a second clean rebuild
  under run id `20260721T200615Z-70768-2b1d3e04e6ea4a98b25d954981ff29fc` and archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-21_220626`. No Aurora diff touches SettingsStore or search
  serialization; archived repository evidence shows the same case passed before the current branch regression.
- Downstream SettingsStore/Preferences failures caused native Error dialogs with `0x8007051A`
  (`ERROR_REVISION_MISMATCH`) and a leaked Make File List Save As dialog. The dialogs were dismissed and the
  already-invalid Full operation was stopped to protect the interactive desktop. The supported contamination
  recovery rebuild then passed.
- This unrelated current-branch SettingsStore/test-automation regression is explicitly outside Aurora ownership.
  It is recorded rather than concealed; it does not invalidate the feature-owned green build, focused runtime,
  DXUI, localization, accessibility, lifecycle, or performance evidence above.

### Authoritative closeout

Durable behavior now lives in `Specs/UI/UI_ThemeCycleOverlay.md` and is cross-referenced from
`Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/UI/UI_DxUiWinUIDesign.md`,
`Specs/Core/Core_SharedHelpers.md`, `docs/Themes.md`, `docs/UserGuide.md`, and `docs/DxUi.md`.
The WIP index no longer routes Aurora as executable work.
