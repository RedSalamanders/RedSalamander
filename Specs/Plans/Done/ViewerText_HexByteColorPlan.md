# ViewerText Hex Byte Color Plan

Last updated: 2026-04-02

Status: Done

## References

- External design reference:
  - Alice Pellerin, "your hex editor should color-code bytes", published 2026-03-31:
    - https://simonomi.dev/blog/color-code-your-bytes/
- Current implementation:
  - `Plugins/ViewerText/ViewerText.cpp`
  - `Plugins/ViewerText/ViewerText.h`
  - `Plugins/ViewerText/ViewerText.Hex.cpp`
  - `RedSalamander/FolderWindow.Viewers.cpp`
  - `RedSalamander/ViewerPluginManager.cpp`
  - `RedSalamander/Preferences.Plugin.Configuration.cpp`
  - `Tests/ViewerPETests/ViewerPETests.cpp`
  - `RedSalamander/Commands.SelfTest.PluginConfig.cpp`
- Related specs:
  - `Specs/Plugins/Plugins_ViewerPlugins.md`
  - `Specs/UI/UI_PreferencesDialog.md`
  - `Specs/UI/UI_ManagePluginsDialog.md`
  - `Specs/Core/Core_SettingsStore.md`
  - `Specs/Testing/Testing_SelfTests.md`
  - `Specs/Testing/Testing_PerformanceValidation.md`
  - `Specs/TestRuns/README.md`

## Progress Checklist

- [x] Phase 0. Spec and contract alignment
- [x] Phase 1. Persisted ViewerText configuration wiring
- [x] Phase 2. Hex byte palette and contrast rules
- [x] Phase 3. Hex paint-path implementation
- [x] Phase 4. Stale-config overwrite protection for open viewers
- [x] Phase 5. Debug snapshot and deterministic tests
- [x] Phase 6. Perf instrumentation and archived runs
- [x] Phase 7. Docs, validation, and cleanup

## Current Validation Status

Validated on 2026-04-02.

Passing validation runs:

- `.\build.ps1 -ProjectName RedSalamander`
- `MSBuild.exe Tests\ViewerPETests\ViewerPETests.vcxproj /p:Configuration=Debug;Platform=x64`
- `.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=settings_viewer_text_plugin_roundtrip`
- `.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=viewer_text_hex_byte_color_perf`
- `.build\x64\Debug\ViewerPETests.exe TestViewerTextHexByteColorsFollowConfigAndHighContrastFallback`

Archived passing runs:

- ViewerText config round-trip: `Specs/TestRuns/4cb089111a23/Commands/2026-04-02_210241`
- ViewerText hex byte color perf scenario: `Specs/TestRuns/4cb089111a23/Commands/2026-04-02_210245`

Observed perf notes from the archived same-machine candidate run:

- `hexByteColorMode:"off"` initial visible paint emitted `viewer.hex.paint_us=9987`, `viewer.hex.visible_rows=2`, `viewer.hex.visible_bytes=32`, `viewer.hex.colorized_bytes=0`, `viewer.hex.color_runs=0`
- `hexByteColorMode:"off"` scroll repaint emitted `viewer.hex.paint_us=2096`, `viewer.hex.visible_rows=1`, `viewer.hex.visible_bytes=16`, `viewer.hex.colorized_bytes=0`, `viewer.hex.color_runs=0`
- `hexByteColorMode:"leadingNibble"` initial visible paint emitted `viewer.hex.paint_us=2661`, `viewer.hex.visible_rows=2`, `viewer.hex.visible_bytes=32`, `viewer.hex.colorized_bytes=32`, `viewer.hex.color_runs=61`
- `hexByteColorMode:"leadingNibble"` scroll repaint emitted `viewer.hex.paint_us=2164`, `viewer.hex.visible_rows=1`, `viewer.hex.visible_bytes=16`, `viewer.hex.colorized_bytes=16`, `viewer.hex.color_runs=32`

Interpretation notes:

- the off and leading-nibble scenarios were captured sequentially in one same-machine candidate run,
- these metrics are suitable as an archived baseline for future regressions,
- they are not a pre-change versus post-change proof of improvement.

## Goal

Add a persisted, user-controlled option for `builtin/viewer-text` so the Hex view can color-code bytes without breaking the current large-file and streaming behavior.

The v1 feature must:

- live in the existing viewer-plugin configuration path, not in a new host-owned settings section,
- be editable from `Preferences -> Plugins -> Text Viewer` and the reusable schema-driven plugin configuration dialog,
- be saved under `settings.plugins.configurationByPluginId["builtin/viewer-text"]`,
- keep the current monochrome behavior available as an explicit fallback,
- preserve search highlighting, byte selection, theme support, DPI behavior, and large-file hex scrolling,
- include deterministic test coverage and archived performance evidence.

## Current Repo-Fit Summary

`ViewerText` already has the right persistence surface for this feature:

- `GetConfigurationSchema()` exposes schema-driven plugin settings.
- `SetConfiguration()` / `GetConfiguration()` round-trip a JSON payload already persisted by `ViewerPluginManager`.
- `Preferences.Plugin.Configuration.cpp` and `ManagePluginsDialog.cpp` already render `option` fields from plugin schema with no custom UI work required.

The current Hex rendering path is still monochrome:

- `ViewerText::OnHexViewPaint()` draws the offset, hex, and text columns with a single `_hexViewBrush`.
- `FormatHexLine()` already computes per-byte `ByteSpan` ranges for both hex and text columns, which is the right hook for per-byte foreground coloring.
- search and selection are currently background overlays, so byte colors must compose with those fills rather than replacing them.

There is one existing integration hazard that this work should address:

- `FolderWindow::OnViewerClosed()` persists the viewer instance's full configuration on close.
- If Preferences changes `builtin/viewer-text` while an older viewer window is still open, that stale window can currently write the old configuration back into settings when it closes.

## Recommended Settings Shape

Use a string option, not a boolean:

- key: `hexByteColorMode`
- type: schema `option`
- default: `"leadingNibble"`
- values:
  - `"off"`: current monochrome hex rendering
  - `"leadingNibble"`: color groups based on the byte's high nibble, with special handling for `00` and `FF`

Rationale:

- the existing ViewerText config schema already stores user toggles as string option values,
- a named mode is clearer in JSON and in the Preferences UI than a bare boolean,
- this leaves room for later modes without a schema rename or migration.

Expected persisted JSON example:

```json
{
  "textBufferMiB": 16,
  "hexBufferMiB": 8,
  "showLineNumbers": "0",
  "wrapText": "1",
  "hexByteColorMode": "leadingNibble"
}
```

## Recommended Visual Contract

V1 should follow the article's main idea closely:

- use 18 buckets:
  - one bucket for `00`,
  - one bucket for `FF`,
  - one bucket for each leading nibble `0X` through `FX`, with `00` and `FF` overriding their generic nibble bucket,
- color both:
  - the hex digits for each visible byte,
  - the corresponding glyph in the text column when that byte owns a visible `textSpan`,
- keep the offset column monochrome.

Theme and accessibility rules:

- byte colors must be stable by byte value, not seeded by file name or current path,
- palette generation may depend on theme mode (`darkMode`, `darkBase`, `rainbowMode`, `highContrast`) but the mapping from bucket to color must stay deterministic,
- `00` should be visibly more subtle than normal bytes,
- `FF` should be visibly emphasized,
- group separators and spacing stay neutral,
- in `highContrast`, v1 should fall back to monochrome even when the saved mode is `"leadingNibble"`; accessibility wins over decoration.

Selection and search rules:

- existing search and selection background fills remain the primary highlight mechanism,
- byte colors are foreground-only,
- if a chosen byte color becomes unreadable over selection or search fill, selected glyphs should fall back to a contrast-safe text color for that byte draw.

## Recommended Implementation Shape

### Phase 0. Spec and contract alignment

- [x] Create or update the authoritative viewer spec so `ViewerText` has a documented configuration contract.
- [x] Prefer adding `Specs/Plugins/Plugins_ViewerText.md` and linking it from `Specs/Plugins/Plugins_ViewerPlugins.md`.
- [x] Document the new config key in the ViewerText spec and add a short note in `Specs/Core/Core_SettingsStore.md` that `builtin/viewer-text` now has a documented plugin-defined configuration payload.

### Phase 1. Persisted ViewerText configuration wiring

- [x] Add `HexByteColorMode` to `ViewerText.h`.
- [x] Extend `ViewerTextConfig` with `hexByteColorMode`.
- [x] Add the new field to `kViewerTextSchemaJson` in `ViewerText.cpp`.
- [x] Parse `"hexByteColorMode"` in `SetConfiguration()`.
- [x] Emit `"hexByteColorMode"` in canonical order from `_configurationJson`.
- [x] Extend `SomethingToSave()` so `"leadingNibble"` is part of the default clean state.

Recommended UI text:

- label: `Hex byte colors`
- description: `Color-code visible bytes in hex mode to make binary patterns easier to spot.`
- options:
  - `Off`
  - `Leading nibble (00/FF emphasized)`

### Phase 2. Hex byte palette and contrast rules

- [x] Add a helper that classifies a byte into a palette bucket.
- [x] Add a helper that builds the 18-color palette from the current viewer theme.
- [x] Keep palette generation independent from `_currentPath` and `ResolveAccentColor(seed)`.
- [x] Reuse `BlendColor`, `ContrastingTextColor`, and the existing theme helpers where they already fit.

Recommended palette behavior:

- light theme: lower saturation than terminal-rainbow defaults,
- dark theme: slightly brighter values so colored glyphs survive on dark backgrounds,
- rainbow mode: allowed to use stronger saturation, but still preserve the same bucket mapping,
- `00`: blended toward background so null padding recedes,
- `FF`: stronger contrast than the generic `FX` bucket.

### Phase 3. Hex paint-path implementation

- [x] Keep `FormatHexLine()` as the source of visible-byte layout/spans.
- [x] Add per-byte bucket capture for the visible row during paint.
- [x] Do not preprocess the whole file; classification must stay bounded to visible bytes only.
- [x] Keep the current background overlay order:
  - row highlight,
  - search highlight,
  - selection highlight,
  - then text draw.

Recommended draw strategy for v1:

- draw the full hex/text row once in the normal foreground color so spacing and separators remain correct,
- overlay colored glyph runs for visible bytes only,
- coalesce adjacent bytes into runs when possible to reduce draw calls,
- avoid per-row heap churn and avoid per-row `IDWriteTextLayout` allocation unless profiling proves the simpler run-overlay path is too slow.

This keeps the change local to `ViewerText.Hex.cpp` and preserves the current hit-testing and selection math.

### Phase 4. Stale-config overwrite protection for open viewers

- [x] Extend `FolderWindow::ViewerInstance` to remember the configuration that was applied when the viewer was created.
- [x] On viewer close, compare the current config returned by `IInformations::GetConfiguration()` with the initial config captured at open time.
- [x] Only persist back into settings when the viewer instance actually changed its own config during that session.

Why this matters:

- the new setting is expected to be changed from Preferences,
- `ViewerText` also changes some settings from inside the viewer window (`wrapText`, `showLineNumbers`),
- without a session-delta guard, closing an older window can revert a newer Preferences change.

This is a small host hardening change, but it prevents a real settings regression for this feature.

### Phase 5. Debug snapshot and deterministic tests

- [x] Add a debug-only `ViewerText` snapshot message in `Common/WindowMessages.h`, following the `ViewerSqlite` pattern.
- [x] Expose enough state to verify color mode without pixel scraping.

Recommended debug snapshot fields:

- current `ViewMode`,
- current `HexByteColorMode`,
- render count,
- visible row count,
- visible byte count,
- visible colorized byte count,
- visible unique color bucket count,
- whether high-contrast fallback disabled colorization.

Recommended deterministic tests:

- [x] `RedSalamander/Commands.SelfTest.PluginConfig.cpp`
  - add a `builtin/viewer-text` configuration round-trip case,
  - assert the schema exposes `hexByteColorMode`,
  - assert default state is clean,
  - assert `"leadingNibble"` becomes dirty and round-trips,
  - assert reset to `"off"` returns to clean.
- [x] `Tests/ViewerPETests/ViewerPETests.cpp`
  - add a real-window `ViewerText` hex-mode test using a crafted binary fixture with `00`, `0F`, `10`, `7F`, `80`, and `FF`,
  - verify the debug snapshot reports color mode on,
  - verify visible colorized bytes are non-zero,
  - verify more than one color bucket is visible,
  - verify `highContrast` forces the reported fallback when applicable.

## Phase 6. Perf instrumentation and archived runs

This feature touches a repeated paint path and therefore falls under the mandatory perf-validation contract.

Protected scenarios:

- `ViewerText hex initial visible paint with byte colors disabled`
- `ViewerText hex initial visible paint with byte colors enabled`
- `ViewerText hex wheel-scroll repaint with byte colors enabled on streamed binary data`

Recommended metric family:

- `viewer.hex.paint_us`
- `viewer.hex.visible_rows`
- `viewer.hex.visible_bytes`
- `viewer.hex.colorized_bytes`
- `viewer.hex.color_runs`
- `viewer.hex.high_contrast_fallback_count`

Recommended implementation:

- [x] add a `Debug::Perf::Scope` around the visible-row hex paint section,
- [x] emit bounded counters for rows, bytes, and color runs per paint,
- [x] keep metric emission limited to visible work only.

Recommended deterministic perf path:

- [x] add a focused `--commands-selftest` case that opens a crafted binary file in `builtin/viewer-text` with `VIEWER_OPEN_FLAG_START_HEX`,
- [x] run the same scenario twice on the same machine:
  - once with `"hexByteColorMode":"off"`,
  - once with `"hexByteColorMode":"leadingNibble"`,
- [x] archive the resulting `Commands` run under `Specs/TestRuns/<MachineHash>/Commands/<timestamp>/`,
- [x] compare `viewer.hex.paint_us` and `viewer.hex.color_runs` before accepting the feature.

The feature is not done until the archived run exists or the blocking reason is documented explicitly.

### Phase 7. Docs, validation, and cleanup

- [x] Update the authoritative viewer spec with the final behavior and config key.
- [x] Update any user-facing viewer/plugin documentation that enumerates built-in viewer settings.
- [x] Build the touched targets.
- [x] Run the focused plugin config test coverage.
- [x] Run the ViewerText viewer-harness coverage.
- [x] Run the commands selftest perf scenario and archive the run.

## Non-Goals For V1

- user-editable custom palettes,
- live propagation of the new setting into already-open ViewerText windows,
- coloring the offset column,
- file-type-aware semantic coloring,
- replacing selection/search backgrounds with colorized byte backgrounds.

## Exit Criteria

The plan is complete only when all of the following are true:

- `builtin/viewer-text` exposes a persisted `hexByteColorMode` option,
- Preferences and the schema-driven plugin configuration dialog can edit it,
- Hex view can render byte colors from visible bytes only,
- `highContrast` safely falls back to monochrome,
- stale viewer windows no longer overwrite newer persisted viewer settings on close,
- deterministic config round-trip and runtime viewer tests are green,
- archived same-machine perf evidence exists under `Specs/TestRuns/`.
