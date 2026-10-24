# ViewerText Diff Polish And Perf Plan

Last updated: 2026-04-04

Status: Implemented

## References

- Current implementation:
  - `Plugins/ViewerText/ViewerText.cpp`
  - `Plugins/ViewerText/ViewerText.h`
  - `Plugins/ViewerText/ViewerText.Text.cpp`
  - `Common/WindowMessages.h`
  - `Tests/ViewerPETests/ViewerPETests.cpp`
  - `RedSalamander/Commands.SelfTest.PluginConfig.cpp`
- Theming integration points:
  - `RedSalamander/AppTheme.h`
  - `RedSalamander/AppTheme.cpp`
  - `RedSalamander/RedSalamander.cpp`
  - `Specs/Themes/*.theme.json5`
- Related specs:
  - `Specs/Plugins/Plugins_ViewerText.md`
  - `Specs/Plugins/Plugins_ViewerPlugins.md`
  - `Specs/Testing/Testing_PerformanceValidation.md`
  - `Specs/Testing/Testing_TestCoverage.md`
  - `Specs/TestRuns/README.md`
- Prior implementation plan:
  - `Specs/Plans/Done/ViewerText_DiffViewerPlan.md`

## Progress Checklist

- [x] Phase 0. UX target and screenshot-fit alignment
- [x] Phase 1. Theme-aware semantic diff palette and tokens
- [x] Phase 2. Rich diff row rendering and visual hierarchy
- [x] Phase 3. Hunk navigation and hidden-context affordances
- [x] Phase 4. Side-by-side polish and interaction refinement
- [x] Phase 5. Settings, persistence, and command surface review
- [x] Phase 6. Debug snapshot, deterministic validation, and accessibility
- [x] Phase 7. Perf instrumentation, tuning, and archived baselines
- [x] Phase 8. Normative docs and cleanup
- [x] Post-plan UX follow-up. Pane-local side-by-side text layout

## Why A Follow-Up Plan Exists

The first diff viewer landing is functionally complete, but it still reads too much like formatted text instead of a purpose-built diff surface.

The user-provided screenshots clarify the desired direction:

- changed rows should read immediately as `added` or `deleted` through row-level background treatment, not only through textual markers,
- the side-by-side split should feel like two coordinated panes with a clearer gutter and calmer context rows,
- hidden unchanged regions should read like compact diff UI (`82 hidden lines`) rather than plain placeholder text,
- the visual language must adapt to the current theme instead of hardcoding a single green/red look,
- this polish must not regress scroll smoothness, rehydration behavior, or bounded referenced-file I/O.

This plan covers that second stage.

## Goal

Upgrade `ViewerText` diff mode from a text-shaped presentation into a more recognizably diff-native UI while preserving the already-landed bounded parsing, referenced-file hydration, and perf protections.

The follow-up must:

- introduce theme-aware row semantics for `added`, `removed`, `context`, `header`, `placeholder`, and `hidden context` rows,
- make hidden-context regions and hunk boundaries legible at a glance,
- add direct hunk navigation affordances beyond file-section navigation,
- keep inline and side-by-side layouts visually consistent,
- preserve accessibility and high-contrast fallback,
- extend perf coverage so richer rendering stays bounded and measurable.

## Non-Goals

Out of scope for this follow-up:

- word-level diff refinement inside changed lines,
- patch application or editing,
- merge/conflict resolution UI,
- minimap/overview ruler unless needed for hunk navigation later,
- fully replacing the text surface with a separate custom editor control.

## UX Target

### 1. Visual hierarchy

The rendered diff should have three clear layers:

- structural rows: file headers, hunk headers, hidden-context banners,
- semantic content rows: added, removed, unchanged, placeholder,
- supporting chrome: gutters, line numbers, pane divider, active focus.

The current landing already has the information; this follow-up is about making the semantics visible without making the view noisy.

### 2. Screenshot-fit expectations

Target feel, derived from the screenshots:

- `added` rows use a theme-aware tinted background across the row or pane cell,
- `removed` rows use a distinct theme-aware tinted background across the row or pane cell,
- unchanged context rows stay quiet and readable,
- hidden-context rows collapse into a compact banner-style row with explicit counts,
- hunk headers read like separators or chips, not ordinary text lines,
- line numbers and markers stay aligned and easy to scan in both panes.

Do not clone another editor exactly, but aim for the same clarity.

### 3. Theme behavior

The row semantics must adapt to:

- light themes,
- dark themes,
- built-in `ThemeMode::Rainbow` / `rainbowMode`,
- custom JSON5 themes,
- high-contrast mode.

The colors should come from semantic theme tokens with fallbacks, not from hardcoded RGB values inside the diff renderer.

## Settings And Persistence

### 1. Persisted settings

Avoid adding many new knobs. Recommended persisted additions only if needed after implementation spikes:

- `diffVisualStyle`
  - values: `"semanticRows"` or `"textOnlyLegacy"`
  - default: `"semanticRows"`
  - only add if we need a real escape hatch during rollout
- `diffContextPresentation`
  - values: `"banner"` or `"expandedOnly"`
  - default: `"banner"`
  - only add if hidden-context banners become user-toggleable

Default preference is to ship the polished UI as the standard experience without multiplying settings.

### 2. Non-persisted session state

Keep these session-local:

- active hunk index,
- temporary hidden-context expansion state for one banner,
- temporary inline vs side-by-side switch,
- horizontal pane offsets,
- debug-only visual overlays.

### 3. Theme settings contract

Do not store add/remove row colors inside `builtin/viewer-text` plugin config.

Instead:

- define semantic theme keys in the app theme contract,
- provide defaults in `AppTheme`,
- treat built-in `ThemeMode::Rainbow` as a first-class semantic diff palette source rather than a special-case viewer override,
- allow shipped theme JSON5 files to override them,
- keep ViewerText responsible for using semantic tokens, not inventing per-theme policy.

## Recommended Architecture

### 1. Semantic diff row styling

Extend the diff row model so rendering can distinguish:

- `Context`,
- `Added`,
- `Removed`,
- `HunkHeader`,
- `FileHeader`,
- `HiddenContextBanner`,
- `PlaceholderFull`,
- `PlaceholderLeft`,
- `PlaceholderRight`.

The current text materialization can remain, but paint must stop treating all rows as visually equivalent.

### 2. Row background and gutter painting

Add paint-time support for:

- row background fills,
- pane-scoped fills in side-by-side mode,
- hunk/header separator bands,
- compact hidden-context banners with centered or gutter-attached labels,
- active hunk focus treatment.

This should layer cleanly with selection and search highlighting instead of fighting them.

### 3. Navigation model

Keep file-section navigation in the combo, and add hunk navigation through commands:

- `Diff / Next hunk`
- `Diff / Previous hunk`
- optional header buttons if the chrome can support them cleanly

The active hunk should be inferable from the current viewport and highlighted subtly.

### 4. Hidden-context interaction

Hidden unchanged regions should become compact banner rows that:

- display counts such as `82 hidden lines`,
- can optionally expand the local range on demand,
- still preserve bounded referenced-file hydration semantics.

Local expand/collapse is a UX affordance, not permission to preload the entire file.

## Planned Phases

### Phase 0. UX target and screenshot-fit alignment

- [x] Capture the current UI gaps explicitly in the plan/specs.
- [x] Define the target visual semantics for headers, hunks, added rows, removed rows, context rows, and hidden banners.
- [x] Decide whether the first rollout needs a temporary legacy visual fallback.
- [x] Document high-contrast and theme-adaptation rules up front.

### Phase 1. Theme-aware semantic diff palette and tokens

- [x] Define semantic theme tokens for diff added/removed/context/header/banner/placeholder surfaces.
- [x] Add defaults in `AppTheme`.
- [x] Wire theme override loading and schema/spec updates for the new tokens.
- [x] Update shipped themes with reasonable defaults and contrast-safe fallbacks.
- [x] Treat built-in `ThemeMode::Rainbow` / `rainbowMode` as an explicit contract in both specs and runtime validation.

### Phase 2. Rich diff row rendering and visual hierarchy

- [x] Add row- and pane-level background painting for semantic diff rows.
- [x] Render file headers and hunk headers as structural bands instead of plain text-only rows.
- [x] Keep selection, search, and caret/focus layering readable on top of semantic backgrounds.
- [x] Preserve syntax/text coloring where it still helps, but make row semantics the primary signal.

### Phase 3. Hunk navigation and hidden-context affordances

- [x] Add next/previous hunk commands and keyboard routing.
- [x] Track the active hunk in the viewport and expose it in the debug snapshot.
- [x] Replace plain hidden-context rows with compact banner-style rows.
- [x] This landing keeps hidden-context banners non-interactive, so no extra local expand/collapse state is introduced.

### Phase 4. Side-by-side polish and interaction refinement

- [x] Strengthen pane divider, gutter, and alignment cues in side-by-side mode.
- [x] Ensure added/removed rows color the correct pane cell without muddying the opposite pane.
- [x] Make inline and side-by-side layouts share the same semantic palette and hidden-banner treatment.
- [x] Review copy/find behavior so new banner/header rows remain truthful and predictable.
- [x] Keep the screenshot-fit side-by-side alignment scoped to pane-local visual layout so hunks-only rows line up cleanly without padding the synthesized text buffer.

### Phase 5. Settings, persistence, and command surface review

- [x] Decide whether any new persisted ViewerText keys are actually justified.
- [x] Add commands/menu labels for hunk navigation if implemented.
- [x] Keep defaults clean in `SomethingToSave()`.
- [x] Avoid persisting transient banner/hunk/session state.

### Phase 6. Debug snapshot, deterministic validation, and accessibility

- [x] Extend the ViewerText debug snapshot with active hunk index, banner count, and row-style counters as needed.
- [x] Add `ViewerPETests` coverage for theme-aware diff semantics at the metadata/debug level.
- [x] Expose enough theme-state metadata to verify active rainbow-mode behavior without image diffing.
- [x] Add runtime tests for hunk navigation and hidden-context banner interactions.
- [x] Verify readable fallbacks in high-contrast mode.

### Phase 7. Perf instrumentation, tuning, and archived baselines

- [x] Define new protected scenarios before implementation lands:
  - `viewer.diff.semantic_row_paint_us`
  - `viewer.diff.hunk_jump_to_visible_us`
  - `viewer.diff.hidden_banner_toggle_us`
  - `viewer.diff.viewport_backtrack_us` regression retention
  - `viewer.diff.theme_switch_repaint_us` if theme switching touches the active diff
- [x] Add visible-work counters for styled diff rows and banner rows.
- [x] Measure rainbow-mode theme-switch repaint cost on an already-open parsed diff.
- [x] Confirm richer row painting does not regress bounded referenced-file hydration.
- [x] Archive same-machine runs under `Specs/TestRuns/<MachineHash>/Commands/...`.
- [x] Compare candidate runs against the current `viewer_text_diff_perf` baseline honestly.
- [x] Re-check the scoped divider-alignment pass against the previous follow-up baseline before keeping it.

### Phase 8. Normative docs and cleanup

- [x] Update `Specs/Plugins/Plugins_ViewerText.md` with the polished diff UX contract.
- [x] Update `Specs/Testing/Testing_PerformanceValidation.md` with the new metric family members.
- [x] Update `Specs/Testing/Testing_TestCoverage.md` with new ViewerPETests and Commands coverage.
- [x] Remove any temporary legacy visual path if it is only used during rollout.

## Current Landing Notes

- AppTheme now owns semantic parsed-diff surfaces through `viewer.diff.*` tokens, including the new quiet `viewer.diff.contextBackground` pane surface, and shipped theme JSON5 files override those tokens directly.
- Built-in `ThemeMode::Rainbow` now participates directly in the semantic diff palette contract instead of being treated as a generic custom-theme approximation.
- ViewerText debug snapshots expose semantic row counters, visible styled-row counts, visible context-row counts, paint timing, the active `themeRainbow` flag, and the currently applied semantic diff colors so runtime tests and Commands perf runs can verify theme changes without image diffing.
- Rainbow mode and custom themes are first-class requirements for the semantic diff palette; the viewer consumes the resolved application theme rather than persisting viewer-local add/remove colors.
- Hunk-only parsed diff presentations now surface skipped unchanged regions as explicit `N hidden lines` banner rows instead of silently collapsing them.
- Parsed diff mode now exposes `Diff / Next hunk` and `Diff / Previous hunk` through the View menu and `F7` / `Shift+F7`, while debug snapshots and Commands perf runs track the active hunk index and `viewer.diff.hunk_jump_to_visible_us`.
- Parsed side-by-side rendering now paints quiet context surfaces in both panes, overlays add/remove/placeholder colors only on the affected pane cell, renders file headers as structural bands, and renders hunk headers / hidden-context rows as chip-style banner rows with active-hunk emphasis.
- Hunk-only side-by-side sections now keep a stable divider column across sibling split rows, while expanded unchanged-text sections intentionally keep unpadded synthesis so long context lines do not dominate wrap and viewport-hydration cost.
- ViewerText configuration remains intentionally small: the selftest contract now proves no persisted `diffVisualStyle`, `diffContextPresentation`, `activeDiffHunkIndex`, or `activeDiffSectionIndex` keys are introduced by the polish pass.
- Latest same-machine perf evidence for this follow-up is archived under `Specs/TestRuns/4cb089111a23/Commands/2026-04-04_115942/`, and the same-machine baseline comparison uses `Specs/TestRuns/4cb089111a23/Commands/2026-04-04_111534/` as the pre-polish reference.
- Baseline vs candidate summary on the same machine: large parsed side-by-side open improved from `324144 us` to `288091 us`, semantic paint from `342 us` to `333 us`, theme-switch repaint from `12409 us` to `9478 us`, and scroll repaint from `6838 us` to `4952 us` while visible styled rows increased from `15` to `25` and visible context rows are now reported as `20`.
- For expanded-context diffs on the same machine, hunk jump improved from `8605 us` to `7020 us`, expand-context from `16805 us` to `15807 us`, viewport rehydrate from `106415 us` to `90007 us`, and viewport backtrack from `93349 us` to `44481 us`, while bounded referenced-byte behavior remained identical at `32768` bytes after expand, `65536` after viewport advance, and `0` bytes reread on backtrack.
- The unresolved-placeholder scenario regressed slightly in open-to-first-visible latency (`107663 us` to `110895 us`) while improving semantic paint (`282 us` to `262 us`); this is small enough to keep, but it remains the main caveat from the same-machine comparison.
- Latest post-plan UX evidence is archived under `Specs/TestRuns/4cb089111a23/Commands/2026-04-04_123508/` for `viewer_text_diff_perf` and `Specs/TestRuns/4cb089111a23/Commands/2026-04-04_123523/` for `settings_viewer_text_plugin_roundtrip`.
- Compared with the previous follow-up baseline at `Specs/TestRuns/4cb089111a23/Commands/2026-04-04_115942/`, the scoped divider pass is mixed but bounded: large hunks-only side-by-side open moved from `288091 us` to `295297 us`, semantic paint from `333 us` to `443 us`, theme-switch repaint from `9478 us` to `9294 us`, and scroll repaint from `4952 us` to `7069 us`; expanded-context open improved from `110620 us` to `91947 us` and hunk jump from `7020 us` to `6290 us`, while expand-context moved from `15807 us` to `21344 us`, viewport rehydrate from `90007 us` to `130997 us`, and viewport backtrack from `44481 us` to `65276 us`. Bounded referenced-byte behavior stayed identical at `32768` after expand, `65536` after viewport advance, and `0` reread on backtrack.
- Latest pane-local side-by-side layout evidence is archived under `Specs/TestRuns/4cb089111a23/Commands/2026-04-04_131259/` for `viewer_text_diff_perf` and `Specs/TestRuns/4cb089111a23/Commands/2026-04-04_131303/` for `settings_viewer_text_plugin_roundtrip`.
- The deeper pane-local layout pass keeps the side-by-side buffer truthful while moving alignment into paint-time row metadata, and debug snapshots now expose `paneLocalSideBySideLayout`, `visibleSplitRowCount`, and the active pane column widths so runtime tests and perf runs can verify that contract directly.
- Compared with the previous same-machine post-plan baseline at `Specs/TestRuns/4cb089111a23/Commands/2026-04-04_123508/`, the pane-local layout pass is mixed: large parsed side-by-side open moved from `295297 us` to `298990 us`, semantic paint improved from `443 us` to `380 us`, theme-switch repaint moved from `9294 us` to `12119 us`, and scroll repaint moved from `7069 us` to `10542 us`, while the large scenario now reports `visibleSplitRowCount=13` with pane columns `36 / 36 / 3`.
- For expanded-context diffs on the same machine, open moved from `91947 us` to `96118 us`, hunk jump from `6290 us` to `7875 us`, expand-context from `21344 us` to `25087 us`, viewport rehydrate from `130997 us` to `156573 us`, and viewport backtrack from `65276 us` to `82051 us`. Bounded referenced-byte behavior stayed identical at `32768` after expand, `65536` after viewport advance, and `0` reread on backtrack.
- The unresolved-placeholder scenario improved in open-to-first-visible latency from `111478 us` to `94588 us` while semantic paint moved from `247 us` to `479 us`; this pane-local pass is therefore acceptable as a UX-driven tradeoff, but it is not a pure perf win.

## Validation Targets

The follow-up is not complete until all of the following are true:

- added and removed rows are visually distinguishable by background treatment in both inline and side-by-side layouts,
- colors adapt correctly across light, dark, custom, and high-contrast themes,
- colors adapt correctly across built-in rainbow mode and shipped/custom theme definitions,
- hidden unchanged regions read like explicit diff UI rather than plain text placeholders,
- users can navigate between hunks without manually scrolling for every section,
- rich rendering does not regress the bounded file-read and viewport-hydration behavior already landed,
- deterministic tests cover navigation, semantic row metadata, and hidden-context behavior,
- archived perf evidence exists for the richer rendering path.

## Recommended First Landing Order

To reduce risk, land this follow-up in the following order:

1. Theme tokens and paint primitives for semantic diff rows.
2. Hidden-context banner rows and hunk navigation.
3. Side-by-side pane polish and active-hunk cues.
4. Perf instrumentation, archived runs, and spec finalization.
