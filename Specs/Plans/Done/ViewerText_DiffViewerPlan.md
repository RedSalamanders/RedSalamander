# ViewerText Diff Viewer Plan

Last updated: 2026-04-03

Status: Implemented

Follow-up planning for visual polish, hunk navigation, and additional perf tuning is captured in `ViewerText_DiffPolishAndPerfPlan.md`.

## References

- Current implementation:
  - `Plugins/ViewerText/ViewerText.cpp`
  - `Plugins/ViewerText/ViewerText.h`
  - `Plugins/ViewerText/ViewerText.Text.cpp`
  - `RedSalamander/FolderWindow.Viewers.cpp`
  - `RedSalamander/ViewerPluginManager.cpp`
  - `Common/SettingsStore.h`
  - `Common/PlugInterfaces/Viewer.h`
  - `Common/PlugInterfaces/FileSystem.h`
  - `Tests/ViewerPETests/ViewerPETests.cpp`
  - `RedSalamander/Commands.SelfTest.PluginConfig.cpp`
- Related specs:
  - `Specs/Plugins/Plugins_ViewerPlugins.md`
  - `Specs/Plugins/Plugins_ViewerText.md`
  - `Specs/Core/Core_SettingsStore.md`
  - `Specs/Testing/Testing_SelfTests.md`
  - `Specs/Testing/Testing_PerformanceValidation.md`
  - `Specs/TestRuns/README.md`

## Progress Checklist

- [x] Phase 0. Scope, UX contract, and repo-fit alignment
- [x] Phase 1. Diff document model and streaming parser
- [x] Phase 2. Referenced-file resolution and full-text expansion
- [x] Phase 3. Inline and side-by-side render surfaces
- [x] Phase 4. Settings, persistence, and host integration
- [x] Phase 5. Debug snapshot and deterministic validation
- [x] Phase 6. Perf instrumentation and archived runs
- [x] Phase 7. Normative docs and cleanup

## Current Landing Notes

- 2026-04-03 first implementation slice landed in `ViewerText`.
- Parsed diff rendering currently reuses the existing text surface by materializing bounded in-memory `inline` and `side-by-side` text variants.
- Parsed diff parsing is still bounded, but recognized diff documents can now promote from the initial text chunk to a full bounded reread up to the fully buffered parse cap even when `textBufferMiB` is smaller, and oversized streamed multi-file diffs now keep a bounded raw-text section index so the header combo can still jump between file entries without full parsed materialization.
- Unchanged-text expansion now defers referenced-file reads and expanded-variant materialization until the user actually requests `Show unchanged text`, scopes those referenced-file reads to the active parsed diff file section while evicting the previous section cache on section navigation, and only materializes unchanged rows for the visible logical window plus margin inside that active section.
- Referenced-file hydration inside the active parsed diff section now reads only the byte ranges needed to decode the requested logical lines, grows incrementally as the viewport advances, and keeps tail growth bounded instead of preloading the full left/right files.
- Parsed diff opens now retain the parsed diff document model in memory so later unchanged-text expansion and layout switches can reuse it without reparsing the raw patch text.
- Expanded diff layouts now also reuse cached referenced-file content within the same document session, so switching from one expanded layout to the other does not reread the same referenced files.
- Parsed diff presentations now prepend explicit `old path:` / `new path:` summary rows so compared file paths remain readable in the rendered text itself.
- Parsed diff presentations with multiple file sections now repurpose the header file combo to navigate patch sections directly; sibling-file next/previous navigation remains unchanged.
- Parsed diff auto-open now covers medium-sized diffs that fully fit inside the active text buffer, and recognized diff payloads can promote beyond that normal buffer size up to the fully buffered parse cap. Larger streamed diffs still fall back to raw text, but recognized multi-file patches up to the bounded streamed-index cap now retain file-section combo navigation in that raw-text fallback.
- Diff-like text files discovered by header sniffing now get the same bounded reread promotion as extension-recognized diff files, so parsed diff open can still win when the first chunk proves the payload is a unified diff.
- Placeholder messaging still emits explicit placeholder rows in the rendered diff text, and those rows now draw with themed background bands on the parsed diff surface so they read as intentional empty-region UI instead of missing paint.
- `ViewerPETests` now cover parsed multi-file diffs, create/delete/rename metadata retention, malformed-hunk raw-text fallback, path-summary text visibility, resolved unchanged expansion, parsed-document reuse, referenced-file content reuse across expanded layouts, active-section-only unchanged-text hydration with on-demand section jumps, viewport-windowed unchanged-row rehydration on scroll, range-bounded referenced-file reads for viewport-nearby unchanged context, unresolved placeholder rows with background-band metadata, large-diff promotion beyond the normal text buffer size for both extension and header-sniffed detection, and raw streamed multi-file section indexing/navigation beyond the fully buffered parse cap.
- `ViewerPETests` now also prove that scrolling back to an already hydrated viewport range reuses cached referenced-file content without rereading the same file prefix.
- Diff-specific perf baseline coverage now ships via `viewer_text_diff_perf`, with the latest same-machine archived evidence under `Specs/TestRuns/4cb089111a23/Commands/2026-04-03_225829/`, and now records placeholder rows, placeholder bands, deferred unchanged rows, referenced-file bytes read, viewport rehydration latency, and zero-delta backtrack reuse for the protected scenarios.
- Post-plan stabilization tightened `viewer_text_diff_perf` so the resolved expansion scenario now waits for actual additional referenced-file byte growth after page-down navigation and archives both the expand-time and post-viewport byte counts in the JSON artifact.
- The perf stream now also emits `viewer.diff.viewport_referenced_bytes_read` and `viewer.diff.viewport_referenced_bytes_delta` so bounded post-scroll growth is comparable without opening the JSON artifact.
- The authoritative Commands perf scenario now also page-ups back into an already hydrated range and emits `viewer.diff.viewport_backtrack_us` plus `viewer.diff.viewport_backtrack_referenced_bytes_delta` to prove cached reuse during backtracking.

## Goal

Add a `.diff` / `.patch` document mode inside `builtin/viewer-text` so the built-in text viewer can present unified diff files as a real diff instead of only raw patch text.

The v1 feature must:

- recognize unified diff / git diff text documents and open them in a diff presentation by default,
- support both `side-by-side` and `inline` layouts,
- preserve the current `ViewerText` strengths: async open, streamed reading, theme support, keyboard navigation, search, other-file navigation, and raw text / hex escape hatches,
- optionally expand omitted unchanged lines by reading the referenced files through the active `IFileSystem`,
- render a clear themed placeholder in the empty background when unchanged text cannot be expanded because referenced files are unavailable or unresolvable,
- keep all persisted defaults under `settings.plugins.configurationByPluginId["builtin/viewer-text"]`,
- wire default extension mappings through `settings.extensions.openWithViewerByExtension`,
- include deterministic tests and archived perf evidence from the start.

## Scope And Non-Goals

In scope for v1:

- unified diff text,
- git-style headers (`diff --git`, `index`, `---`, `+++`, `@@`),
- create / delete / rename metadata display,
- multiple file sections in one patch,
- lazy full-text expansion when referenced files are readable through the current filesystem.

Out of scope for v1:

- binary patch application or binary delta visualization,
- three-way merge conflict UI,
- word-level diff refinement inside changed lines,
- persisted per-document path remapping tables,
- manual patch application or file editing.

## Repo-Fit Summary

`ViewerText` is already the right home for this feature:

- it already owns the file-system-backed async open path via `IFileSystemIO::CreateFileReader`,
- it already persists plugin-defined configuration through `GetConfigurationSchema()`, `SetConfiguration()`, `GetConfiguration()`, and `SomethingToSave()`,
- it already has a real runtime test harness in `Tests/ViewerPETests/ViewerPETests.cpp`,
- the host already routes extension-based viewer defaults through `Common::Settings::ExtensionsSettings::openWithViewerByExtension`.

The main design constraint is that `ViewerText` is currently centered on two display paths:

- text mode with streamed chunk loading,
- hex mode with bounded visible work.

The diff viewer should therefore be implemented as a new document/view mode layered on the same async reader and not as a separate plugin.

## Recommended UX Contract

### 1. Presentation modes

Add a diff-aware presentation with three user-facing ways to inspect the same `.diff` document:

- `Diff / Side by side`
- `Diff / Inline`
- `Raw text`

`Hex` remains available through the existing viewer path, but it is an escape hatch, not the default `.diff` experience.

Recommended default:

- recognized `.diff` / `.patch` / `.rej` files open in `Diff / Side by side`,
- parse failure or unsupported diff dialect falls back to `Raw text` with a non-modal info banner.

### 2. Full-text expansion

The diff view has two content modes:

- `Changed hunks only`
- `Show unchanged text when referenced files are available`

When full-text expansion is enabled:

- unchanged gaps should be hydrated lazily from the left and/or right referenced file,
- hydration must stay bounded to the visible region plus a small cache margin,
- missing referenced content must not silently collapse to blank rows.

Instead, the viewer should draw a themed placeholder message in the empty background of the affected pane, for example:

- `Unchanged lines hidden: referenced file is not available in this filesystem.`
- `No left-side file: patch adds this file.`
- `No right-side file: patch deletes this file.`

This placeholder is informational UI, not copied text.

### 3. Navigation model

The diff view should keep `ViewerText`-style navigation but adapt it to diff semantics:

- file combo navigates between patch file sections, not only between sibling `.diff` files,
- existing other-file navigation continues to switch between other `.diff` files in the folder,
- `Find` searches the rendered diff text for the current presentation,
- `Go To` should target diff file index / hunk index in v1, while `Raw text` keeps the existing line-based behavior,
- keyboard scrolling and page movement stay smooth and synchronized in both layouts.

### 4. Line-number contract

Do not reuse `showLineNumbers` as the diff-mode control.

Recommended v1 behavior:

- diff mode always shows left and right source line numbers because they are intrinsic to a diff,
- raw text mode continues to respect `showLineNumbers`,
- no extra persisted diff line-number toggle is required in v1.

## Recommended Architecture

### 1. Separate document kind from presentation

The current `ViewMode` (`Text` / `Hex`) is too narrow for diff documents.

Recommended split:

- `DocumentKind`
  - `PlainText`
  - `Diff`
- `PresentationMode`
  - `Text`
  - `Hex`
  - `DiffInline`
  - `DiffSideBySide`

This avoids overloading the existing text/hex state and makes raw-text fallback explicit.

### 2. Diff document model

Add a neutral in-memory model that can drive both inline and side-by-side layouts:

- `DiffDocument`
  - ordered `DiffFileSection` entries,
  - each section stores original path, new path, metadata flags, and ordered hunks,
- `DiffHunk`
  - hunk header metadata and ordered lines,
- `DiffLine`
  - `Context`, `Added`, `Removed`, `NoNewlineMarker`, `Header`, or `Placeholder`,
  - original line number and new line number when present,
  - rendered text payload.

For side-by-side rendering, do not create two unrelated trees. Instead, materialize a shared row model:

- `DiffRenderRow`
  - left cell text + state,
  - right cell text + state,
  - shared classification for focus, placeholder, and hunk boundaries.

### 3. Streaming parser contract

The `.diff` file itself must still respect the current streamed-open philosophy.

Recommended path:

- first pass builds a lightweight patch index from the diff file in file order,
- index records byte ranges or logical line ranges for each `DiffFileSection`,
- initial open only needs enough parsed state for the first visible section,
- remaining sections can continue indexing in background,
- visible rows are materialized on demand from the indexed section data.

Do not require the entire patch file to be decoded into one monolithic `std::wstring` before the first paint.

### 4. Raw text fallback stays first-class

Keep the existing raw text reader path available for:

- unsupported patch dialects,
- malformed headers,
- binary diffs,
- exact patch-text inspection,
- validation and debugging.

This reduces risk and preserves all current text-view capabilities.

## Referenced File Resolution

### 1. Resolution rules

Full-text expansion is best-effort and should use only the current `IFileSystem` instance.

For each diff file section:

- use `---` as the left-side source when not `/dev/null`,
- use `+++` as the right-side source when not `/dev/null`,
- strip common git prefixes such as `a/` and `b/` for resolution attempts, but keep original labels for display,
- try the exact path first when it looks absolute in the active filesystem,
- otherwise try paths relative to the `.diff` file parent path.

### 2. Resolution limits

Do not persist a user-specific repository-root mapping table in v1.

Reason:

- it adds cross-filesystem ambiguity,
- it complicates settings migration,
- stale mappings can silently show the wrong file contents.

If automatic resolution fails, the viewer should remain truthful and show placeholders instead of guessing.

### 3. Expansion policy

When `Show unchanged text` is enabled:

- hydrate only the ranges needed for the current viewport and nearby rows,
- keep a bounded cache per visible `DiffFileSection`,
- evict old expansion buffers when the user navigates away.

Current landing:
- referenced-file hydration is bounded to the active parsed diff file section and old section caches are dropped when the user jumps to another section,
- unchanged-row materialization inside that active section is now bounded to the visible logical window plus nearby rows and can rehydrate on scroll,
- referenced files inside that active section are opened lazily and decoded incrementally, reading only enough bytes to satisfy the currently requested logical lines plus bounded tail growth.

Do not load entire left and right files just to draw a few unchanged gaps.

## Render Design

### 1. Inline view

Inline view should be the lower-risk first render path:

- one vertical surface,
- diff markers and dual line numbers in the gutter,
- changed lines colored by diff class,
- placeholder rows drawn as themed background bands between hunks when full-text expansion is unavailable.

This mode can reuse more of the current text-view measurement and selection assumptions.

### 2. Side-by-side view

Side-by-side view should use a shared vertical row model:

- one synchronized vertical scroll position,
- left and right text cells aligned by logical diff row,
- per-pane placeholder background when only one side can expand unchanged text,
- selection/search focus remains row-stable across both sides.

Recommended v1 simplification:

- shared vertical scrolling,
- independent horizontal offsets are allowed but do not need to be persisted,
- word wrap remains off in side-by-side diff mode.

### 3. Placeholder rendering

Placeholder rows should not look like missing paint.

Recommended visual contract:

- draw the normal themed background first,
- add a subtle band or panel in the empty region,
- render a short explanatory sentence centered or left-aligned in that band,
- use subdued theme-aware colors,
- keep the message visible without obscuring real changed lines.

## Settings And Persistence

### 1. Persisted settings

Persist global diff defaults in the existing ViewerText plugin configuration object.

Recommended new keys:

- `diffDefaultLayout`
  - values: `"sideBySide"` or `"inline"`
  - default: `"sideBySide"`
- `diffContextMode`
  - values: `"hunksOnly"` or `"fullFileWhenAvailable"`
  - default: `"hunksOnly"`
- `diffAutoOpenMode`
  - values: `"parsed"` or `"rawText"`
  - default: `"parsed"`

Recommended canonical JSON example:

```json
{
  "textBufferMiB": 16,
  "hexBufferMiB": 8,
  "showLineNumbers": "0",
  "wrapText": "1",
  "hexByteColorMode": "leadingNibble",
  "diffDefaultLayout": "sideBySide",
  "diffContextMode": "hunksOnly",
  "diffAutoOpenMode": "parsed"
}
```

`SomethingToSave()` should be clean only when all diff settings are at their defaults too.

### 2. Non-persisted session state

Do not persist these in v1:

- current active patch file section,
- current hunk focus,
- current temporary raw-text versus diff presentation switch,
- horizontal split offsets,
- referenced-file resolution results,
- cached unchanged text buffers.

These are document-session details, not stable user preferences.

### 3. UI surface for settings

Expose the persisted diff defaults through the existing schema-driven plugin configuration UI:

- `Default diff layout`
- `Unchanged text`
- `Open recognized diff files as`

Also expose quick in-view toggles for:

- `Side by side`
- `Inline`
- `Show unchanged text`
- `Raw text`

In-view changes to the persisted defaults should update the configuration JSON the same way `wrapText` and `hexByteColorMode` already do.

### 4. Extension mapping persistence

Add default mappings in `Common::Settings::ExtensionsSettings::openWithViewerByExtension` for:

- `.diff`
- `.patch`
- `.rej`

All should point to `builtin/viewer-text`.

These stay host-owned settings, not plugin-owned settings.

## Planned Phases

### Phase 0. Scope, UX contract, and repo-fit alignment

- [x] Confirm v1 supports unified diff / git diff text only.
- [x] Define `DocumentKind` and `PresentationMode` terminology in code and docs.
- [x] Decide final user-facing labels for layout and unchanged-text toggles.
- [x] Confirm placeholder wording and accessibility expectations.

### Phase 1. Diff document model and streaming parser

- [x] Add diff format detection on open based on extension plus header sniffing.
- [x] Add a lightweight parser for git/unified diff file sections and hunks.
- [x] Parse fully buffered medium-sized diffs up to the bounded fully buffered parse cap instead of forcing raw text at the old small-file threshold.
- [x] Allow recognized diff documents to promote from the normal text chunk to a bounded full reread up to the fully buffered parse cap when the initial chunk proves the payload is a unified diff.
- [x] Retain the parsed diff document model after a successful parsed open so later variant materialization can reuse it.
- [x] Build a bounded index so the diff file can still open without full eager materialization.
- [x] Keep malformed or unsupported sections recoverable enough to fall back to raw text.

### Phase 2. Referenced-file resolution and full-text expansion

- [x] Add best-effort left/right file resolution through the current `IFileSystemIO`.
- [x] Strip `a/` and `b/` prefixes for resolution attempts only.
- [x] Defer referenced-file reads and expanded diff variant materialization until unchanged text is actually requested.
- [x] Reuse resolved referenced-file content across expanded diff layouts during the same parsed document session.
- [x] Scope unchanged-text hydration and referenced-file caching to the active parsed diff file section, and rehydrate on section jumps.
- [x] Bound unchanged-row materialization to the visible logical window plus nearby rows inside the active section, and rehydrate on scroll.
- [x] Add true referenced-file range hydration within the active section instead of loading full active-section files.
- [x] Add explicit placeholder states for unresolved left/right content and create/delete cases.

### Phase 3. Inline and side-by-side render surfaces

- [x] Add a diff presentation surface that can render inline rows.
- [x] Add the shared-row side-by-side renderer with synchronized vertical scrolling.
- [x] Draw themed placeholder background bands for unresolved placeholder rows on the parsed diff text surface.
- [x] Add diff-aware hit-testing, search, copy, and focus behavior.
- [x] Add header commands for `Inline`, `Side by side`, `Show unchanged text`, and `Raw text`.
- [x] Use the header file combo to navigate parsed diff file sections when a patch contains multiple file entries.

### Phase 4. Settings, persistence, and host integration

- [x] Extend `ViewerText` schema JSON with the diff defaults.
- [x] Parse, canonicalize, and round-trip the new keys in `SetConfiguration()` and `GetConfiguration()`.
- [x] Extend `SomethingToSave()` to treat default diff settings as clean.
- [x] Add default extension mappings for `.diff`, `.patch`, and `.rej`.
- [x] Ensure close-time persistence still only writes when the viewer instance changed its own configuration during the session.

### Phase 5. Debug snapshot and deterministic validation

- [x] Extend the debug snapshot with diff-specific state:
  - document kind,
  - current presentation mode,
  - current file-section count,
  - parsed diff parse count,
  - visible diff rows,
  - placeholder row count,
  - placeholder band count,
  - whether unchanged expansion is enabled,
  - whether left/right referenced files resolved.
- [x] Add parser-focused unit or selftest coverage for multi-file patches, create/delete, rename, and malformed hunks.
- [x] Add plugin-config round-trip coverage for the new diff settings.
- [x] Add real-window `ViewerText` tests in `Tests/ViewerPETests/ViewerPETests.cpp` for inline, side-by-side, raw-text fallback, and placeholder rendering.

### Phase 6. Perf instrumentation and archived runs

- [x] Define the protected scenarios:
  - `viewer.diff.open_to_first_visible_us` for parsed diff open to first visible side-by-side render.
  - `viewer.diff.scroll_repaint_us` for side-by-side scroll repaint after the first render.
  - `viewer.diff.expand_context_us` for unchanged-text expansion when referenced files resolve.
  - `viewer.diff.placeholder_rows` for unresolved-reference placeholder materialization.
  - `viewer.diff.placeholder_bands` for unresolved-reference placeholder-band materialization.
- [x] Add or reuse metrics such as:
  - `viewer.diff.open_to_first_visible_us`
  - `viewer.diff.visible_rows`
  - `viewer.diff.scroll_repaint_us`
  - `viewer.diff.expand_context_us`
  - `viewer.diff.placeholder_rows`
  - `viewer.diff.placeholder_bands`
- [x] Add deterministic selftest coverage for at least one large multi-file patch and one unresolved-reference scenario.
- [x] Archive same-machine runs under `Specs/TestRuns/<MachineHash>/Commands/...` or the area chosen by the existing perf harness.

### Phase 7. Normative docs and cleanup

- [x] Update `Specs/Plugins/Plugins_ViewerText.md` with the diff-mode configuration contract and UX rules.
- [x] Update `Specs/Plugins/Plugins_ViewerPlugins.md` so `builtin/viewer-text` is documented as the built-in diff viewer too.
- [x] Update `Specs/Core/Core_SettingsStore.md` to document the new `builtin/viewer-text` keys.
- [x] Add any localized resources and menu text needed for the new commands.

## Validation Targets

The feature is not complete until all of the following are true:

- opening a `.diff`, `.patch`, or `.rej` file launches `builtin/viewer-text` in parsed diff mode by default,
- the user can switch between inline and side-by-side without reopening the document,
- raw text remains available for exact patch inspection,
- enabling unchanged-text expansion reads referenced files only when available through the active filesystem,
- unresolved references show clear placeholder background text rather than silent blank gaps,
- large patches still open and scroll without preloading the entire patch or both referenced files,
- plugin configuration round-trips the diff defaults correctly,
- deterministic tests cover resolved and unresolved expansion paths,
- archived perf evidence exists or the block is documented explicitly.

## Recommended First Landing Order

To reduce risk, land this in the following order:

1. Diff detection, parser, inline view, and raw-text fallback.
2. Persisted diff defaults and extension mappings.
3. Referenced-file resolution plus placeholder-backed unchanged expansion.
4. Side-by-side renderer.
5. Perf baseline, archived runs, and normative spec updates.
