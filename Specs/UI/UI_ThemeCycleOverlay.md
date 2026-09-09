# Theme Cycle Overlay Specification

**Status:** Authoritative
**Effective:** 2026-07-21
**Owner:** RedSalamander theme command dispatch and `Ui::ThemeCycleOverlayWindow`

## Purpose and scope

RedSalamander gives immediate, passive feedback when a user changes themes from one of these approved surfaces:

- keyboard or Function Bar `Previous Theme` / `Next Theme`;
- **View -> Theme -> Previous Theme** or **Next Theme**;
- a different built-in, file, or inline user theme selected from **View -> Theme**;
- a different direct-theme command invoked by clicking the Function Bar.

The overlay is not a picker and never changes the selection itself. Startup, Preferences preview/apply, settings
or theme-file hot reload, system color changes, programmatic selection, other `WM_COMMAND` sources, and an ordinary
parameterized direct-theme shortcut do not show it. Selecting the active direct theme is a true no-op: it neither
shows the overlay nor restarts an existing lifetime.

## Content and theme ring

The popup shows one owned snapshot:

- the active theme name, large and centered;
- the exact previous ring neighbor, smaller at logical top-leading;
- the exact next ring neighbor, smaller at logical bottom-trailing.

The ring is the same ring used by command dispatch: System, Light, Dark, Rainbow, High Contrast App, file themes
sorted by name/ID, then inline settings themes sorted by name/ID. It wraps in both directions. One-entry defensive
rings omit both neighbors; two-entry rings intentionally show the other entry in both neighbor positions.

Built-in display names come from localized resources. A custom theme uses its complete authored name, or its full
ID when the name is empty. Snapshots own all strings and remain valid if settings or theme vectors subsequently
change. Display text may ellipsize visually but UI Automation always retains the complete name.

## Window, placement, and rendering

The implementation owns one lazily created `WS_POPUP` composition window with `WS_EX_TOOLWINDOW`,
`WS_EX_NOACTIVATE`, and `WS_EX_NOREDIRECTIONBITMAP`. It is owned by the main window, never enters the taskbar or
Alt+Tab list, never takes focus or activation, and reuses its HWND across updates.

The popup is centered over the complete main client. Placement is DPI-aware, responsive, clamped to the monitor
work area, and suppressed below the defensive 192 x 128 DIP owner minimum. Previous/next placement and DirectWrite
reading direction mirror in RTL while the semantic ring direction remains unchanged.

The newly applied `DxUi::ThemePalette` is authoritative. The popup HWND always disables the DWM system backdrop.
The visible card uses the shared `DxUi::CaptureTransientSurfaceBackdrop(...)`, `TransientSurfaceBackdrop`, and
`PaintTransientSurface(...)` app-rendered material policy. Ordinary themes retain a rounded surface, contour,
inner rim, and shadow; Solid/Mica/MicaAlt/Acrylic keep distinct opacity/backdrop policies. High Contrast forces an
opaque rectangular fill and border with no shadow or backdrop. Transparent host gutters and rounded corners remain
transparent and hit-test through to underlying UI.

For a non-Solid ordinary material, the first show captures only the stable card rectangle before exposing the
popup. Visible replacements and redundant owner-geometry notifications reuse that clean snapshot so the popup can
never capture itself. Hide and owner-geometry changes that alter the card's screen rectangle release the snapshot and
its device cache; a capture/effect failure falls back to the themed card fill.

The initial complete frame is primed before the window is exposed. Animation frames reuse cached theme brushes and
DirectWrite layouts; layout creation occurs only for snapshot, DPI, flow/language, or palette changes. Device loss
recreates resources and repaints the current generation without reverting the selected theme. Rendering failure is
best-effort feedback failure: apply the theme, hide the popup, and emit one bounded diagnostic.

Each frame computes one animated card rectangle from the current opacity, scale, and
translation. Paint, visible-surface hit testing, cursor selection, capture start, release
dismissal, and transparent-gutter routing MUST all use that same rectangle. Input must
never target the untransformed endpoint while the card is entering, transitioning, or
exiting. The deterministic edge matrix covers the four animated boundaries and the
transparent corner/gutter cases.

## Motion and lifetime

Timing is measured from the newest accepted generation:

| Stage | Duration | Contract |
| --- | ---: | --- |
| First appearance | 140 ms | Fade/scale/translate to the stable endpoint. |
| Adjacent Previous/Next update | 160 ms | Promote the corresponding neighbor with directional motion. |
| Direct menu/Function Bar update | 140 ms | Neutral content transition. |
| Fully visible steady interval | 900 ms | No animation subscription work, animation request, or rendered frame. |
| Automatic exit | 160 ms | Fade/scale/translate, then hide. |

Every accepted replacement cancels the prior deadline, preserves one HWND, settles the newest content, and starts
a fresh 900 ms steady interval. Input during exit preserves the sampled opacity/scale/translation and recovers to
full visibility within 80 ms; it must not flash hidden. Reduced motion presents each snapshot atomically, preserves
the full 900 ms readable interval, and hides immediately when that interval expires.

The primary deadline is a window timer. If `SetTimer` fails, a WIL-owned thread-pool timer only posts an opaque
cookie back to the UI thread. Cookie, generation, and monotonic deadline are revalidated before exit begins.
Teardown cancels and waits for fallback callbacks. `WM_NCDESTROY` clears the retained-control view and releases the
owned HWND wrapper so owner-driven destruction cannot leave stale window or root pointers.

## Input behavior

A primary press that begins on the visible surface is consumed and captured. Release on the surface dismisses
immediately, bypassing the automatic exit; release outside or `WM_CANCELMODE` cancels the press without dismissal.
Clicking anywhere on the visible card has the same dismiss action. Transparent gutters/corners route through.
Secondary, middle, X-button, and wheel input over the card is consumed but never changes theme or dismisses it.

The entire card uses the hand cursor but has no keyboard focus, close glyph, context menu, wheel navigation, or
theme-selection action. Focus, active window, foreground window, and pane selection remain unchanged.

## Accessibility and localization

The retained root has stable automation ID `ThemeCycleOverlay`, UIA Status control type, no focusability, and the
only Invoke pattern; Invoke performs the same immediate dismissal as a completed primary click. Stable Previous,
Current, and Next text children expose complete localized names and no interactive patterns.

After a matching generation is visibly presented, `RaiseWindowHostAccessibilityNotification(...)` publishes a
most-recent/coalescing notification. Rapid repeat therefore announces the newest state rather than building a
speech backlog. The root raises no focus event and never enters Tab order.

All grammar-bearing names, notification text, dismiss action, and HelpText are STRINGTABLE resources. Variable
strings use positional placeholders and every shipped satellite preserves the source tokens. RTL uses logical
leading/trailing geometry and isolated DirectWrite runs.

## Performance and validation

Protected Release metrics are `theme.cycle.apply_us`, `theme.cycle.overlay.update_us`,
`theme.cycle.input_to_visible_us`, `theme.cycle.overlay.backdrop_capture_us`, and the `theme.cycle.overlay.*`
creation/reuse/backdrop/layout/render/timer counters.
Required gates are:

- at least 200 samples for p95 claims;
- apply p95 no worse than baseline by more than the larger of 10% or 2 ms;
- update p95 at most 8 ms;
- input-to-visible p95 at most 100 ms and max at most 250 ms;
- one popup creation per burst, bounded reuse/resources, no early deadline, no steady-interval frames, and no
  retained capture/timer/animation after hide or destruction.

`theme_cycle_overlay_perf` is the deterministic metric recorder for those gates, and
`theme_cycle_overlay_retention` guards repeated updates plus explicit external-HWND teardown/recreation. Any
performance or closeout claim must cite a fresh same-machine archived run; unavailable historical run directories
and incident artifacts are not normative proof.

## Shared DXUI and Alert Overlay decision

The feature required narrow shared DXUI improvements: Status-root UIA semantics, stable automation IDs,
non-focusing custom Invoke, most-recent notifications, the canonical transient-surface painter/backdrop resolver,
and correct capture-loss notification from `WindowHost::ResetInteractionState()`.

`Ui::AlertOverlayWindow` remains separate. It owns alert/prompt behavior, a scrim, closability, buttons, and
optional modality; migrating it would change behavior and renderer ownership rather than reuse this passive status
surface. No Alert Overlay migration is required for the theme-cycle overlay.
