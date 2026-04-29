# ViewerText Specification

## Overview

`builtin/viewer-text` is the built-in Text/Hex viewer used as the fallback viewer for files without a more specific viewer association. It also owns the built-in parsed diff experience for `.diff`, `.patch`, and `.rej` documents.

The viewer supports:
- text mode with optional wrapping and logical line numbers,
- parsed diff mode with `side-by-side`, `inline`, and `raw text` presentations for unified diff / git diff text,
- optional unchanged-text expansion from referenced files resolved through the active `IFileSystem`,
- hex mode with byte-accurate selection and search highlighting,
- streamed loading for large files,
- theme-aware rendering for normal, built-in rainbow-mode, dark, custom JSON5, and high-contrast themes,
- schema-driven plugin configuration persisted under `settings.plugins.configurationByPluginId["builtin/viewer-text"]`.

## Persisted Configuration

`builtin/viewer-text` stores a plugin-defined JSON object. The canonical configuration shape is:

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

Keys:
- `textBufferMiB` (integer): approximate in-memory streaming buffer for text mode. Default `16`.
- `hexBufferMiB` (integer): approximate in-memory streaming buffer for hex mode. Default `8`.
- `showLineNumbers` (`"0"` or `"1"`): text-mode logical line numbers. Default `"0"`.
- `wrapText` (`"0"` or `"1"`): text-mode wrapping. Default `"1"`.
- `hexByteColorMode` (`"off"` or `"leadingNibble"`): hex-mode foreground coloring. Default `"leadingNibble"`.
- `diffDefaultLayout` (`"sideBySide"` or `"inline"`): parsed diff layout chosen by default when a recognized diff opens in parsed mode. Default `"sideBySide"`.
- `diffContextMode` (`"hunksOnly"` or `"fullFileWhenAvailable"`): whether parsed diffs render only changed hunks or also expand unchanged text from referenced files when those files are readable. Default `"hunksOnly"`.
- `diffAutoOpenMode` (`"parsed"` or `"rawText"`): whether recognized diff files open in parsed diff mode or raw text by default. Default `"parsed"`.

`SomethingToSave()` is clean only when all values are at their defaults.

## Diff Mode

Recognized unified diff / git diff text documents open as parsed diffs by default when:
- the extension is `.diff`, `.patch`, or `.rej`, or
- the decoded text starts with standard unified-diff markers such as `diff --git`, `---`, `+++`, or `@@`.

For the first landing, parsed diff materialization is intentionally bounded:
- the viewer eagerly parses recognized diff files that fully fit within the current fully buffered parse cap, even when they exceed the normal text buffer size, by promoting from the initial text chunk to a bounded full reread when the file extension or decoded header proves the payload is a unified diff,
- larger streamed diffs, malformed diffs, or unsupported diff dialects fall back to raw text with a non-modal status message, and recognized streamed multi-file diffs up to the bounded section-index cap still build a raw file-section index so the header combo can jump between patch entries without full parsed materialization,
- parsed diff rendering reuses the existing streaming text surface by synthesizing formatted diff text for `inline` and `side-by-side` layouts, while side-by-side paint uses pane-local visual layout for the left cell, separator, and right cell instead of depending on padded divider text in the synthesized buffer,
- hunks-only parsed opens do not probe referenced left/right files, expanded presentations materialize referenced files on first use instead of during open, later parsed diff layout/context work reuses the cached parsed diff document instead of reparsing the raw patch text, expanded layouts keep referenced-file caching bounded to the active diff file section while reusing that active section cache across sibling layouts within the same document session, unchanged-row text inside the active expanded section is only materialized for the visible logical window plus a small margin while far-away rows stay deferred, and referenced left/right files are decoded incrementally so viewport-nearby context reads stay bounded instead of forcing full active-section file loads.

Parsed diff commands exposed from `View`:
- `Diff / Side by side`
- `Diff / Inline`
- `Diff / Show unchanged text`
- `Diff / Next hunk`
- `Diff / Previous hunk`

Behavior:
- switching between `Side by side` and `Inline` updates the persisted `diffDefaultLayout`,
- toggling `Show unchanged text` updates the persisted `diffContextMode`,
- `View / Text` switches parsed diff documents to raw diff text for the current session, and the chrome mode button cycles `Diff -> Raw -> Hex`,
- the first `Show unchanged text` request for a parsed diff may materialize the selected expanded presentation on demand, but it only hydrates unchanged text for the active diff file section instead of eagerly expanding every file entry in the patch,
- when unchanged text is visible, the viewer materializes unchanged rows only for the visible logical window plus a small margin inside the current diff section; scrolling can rehydrate a later logical window while distant rows remain deferred,
- once a diff was parsed successfully during open, later parsed diff variant materialization reuses that cached parsed document model,
- once referenced files were loaded for the active expanded diff section, switching to the other expanded parsed layout reuses those cached referenced-file contents for the same open document,
- referenced-file loading inside the active expanded section is range-bounded: the viewer only reads enough bytes to decode the logical lines needed for the current hydrated window and any bounded tail growth near the viewport,
- when the viewport returns to an already hydrated logical range in the same active diff section, the viewer reuses cached referenced-file content instead of rereading the same file prefix,
- switching to `Raw text` is session-local and does not overwrite `diffAutoOpenMode`,
- parsed diff presentations repurpose the header file combo to navigate patch file sections when a diff contains multiple file entries,
- when unchanged text is hidden, skipped unchanged ranges render as compact clickable `Show N hidden lines` banner chips instead of disappearing silently,
- parsed diff presentations expose `Diff / Next hunk` and `Diff / Previous hunk` through the View menu and `F7` / `Shift+F7`, and those commands navigate between parsed hunks without reparsing the patch text,
- switching parsed diff layouts or unchanged-text mode keeps the current diff section selected, and jumping to another diff section rehydrates that section on demand while dropping the previous section's referenced-file cache,
- existing next/previous other-file navigation continues to target sibling `.diff` / text documents rather than diff sections,
- parsed diff presentations always embed source-side line numbers and therefore suppress the extra `showLineNumbers` gutter while they are active,
- parsed diff presentations keep horizontal scrolling available when wrapping is disabled: `side-by-side` applies the same left-column state across both visible panes, `inline` uses the normal single-surface text viewport, and switching between those presentations recomputes the horizontal scrollbar geometry while resetting stale carry-over offsets,
- parsed diff presentations prepend explicit `old path:` / `new path:` text lines for each file section before the raw diff metadata so the compared paths stay readable in text mode,
- parsed diff presentations do not render unified-diff hunk anchor rows such as `@@ -1,19 +1,21 @@` as visible document rows,
- unresolved referenced files surface explicit placeholder rows instead of silently rendering blank gaps, and those placeholder rows plus absent panes draw on the base text background with subtle hatched treatment,
- unchanged parsed-diff content rows use the normal text background so side-by-side mode reads like two coordinated panes instead of layered context fills,
- side-by-side parsed diffs paint the base text background in both panes first, then overlay `added`, `removed`, and `placeholder` treatments only on the affected pane cell, while a calmer divider lane plus a subtle center rule keep the split legible without adding synthetic text noise,
- hunks-only side-by-side parsed diffs keep sibling split rows aligned through pane-local visual layout within each file section, while the synthesized text stays truthful for copy/find/debug preview purposes,
- expanded unchanged-text side-by-side diffs keep the same pane-local semantic pane painting and truthful row synthesis while preserving bounded wrapping and viewport-rehydration work for long context rows,
- leading diff markers (`+` / `-`) stay visible but render dimmed toward the current background so the changed text remains dominant,
- file headers render as full-width structural bands, hunk headers stay structural-only in parsed mode, and hidden-context rows render as compact centered clickable `Show N hidden lines` banner chips instead of flat text rows,
- parsed diff row semantics use app-theme-owned tokens rather than viewer-local persisted colors:
  `viewer.diff.addedBackground`, `viewer.diff.removedBackground`, `viewer.diff.contextBackground`, `viewer.diff.headerBackground`,
  `viewer.diff.bannerBackground`, `viewer.diff.placeholderBackground`, and `viewer.diff.divider`,
- debug/runtime validation for parsed diffs must expose the total parsed hunk count plus the active viewport hunk index so command routing and viewport-derived hunk selection remain deterministic without image-based assertions,
- those semantic diff tokens must adapt cleanly across light, dark, built-in `ThemeMode::Rainbow`, custom JSON5 themes, and high-contrast themes,
- built-in `ThemeMode::Rainbow` must resolve a distinct semantic diff palette through the application theme system even when no custom theme file overrides `viewer.diff.*`,
- changing the application theme after a diff is already open must repaint the active parsed diff with the newly resolved semantic diff colors.

## Performance Validation

ViewerText diff mode treats these as protected performance scenarios:
- parsed side-by-side open for a large multi-file patch,
- side-by-side scroll repaint after the first visible render,
- visible split-row activation and pane-local layout on the first parsed side-by-side viewport,
- unchanged-text expansion when referenced files resolve,
- scroll-driven rehydration of a later unchanged-text viewport window,
- backtracking to an already hydrated unchanged-text viewport window,
- placeholder-row materialization when referenced files do not resolve.

The authoritative Commands selftest entrypoint is `viewer_text_diff_perf`. The baseline metric family emitted by that case is `viewer.diff.*`, including `viewer.diff.open_to_first_visible_us`, `viewer.diff.visible_rows`, `viewer.diff.semantic_row_paint_us`, `viewer.diff.visible_styled_rows`, `viewer.diff.visible_context_rows`, `viewer.diff.visible_banner_rows`, `viewer.diff.theme_switch_repaint_us`, `viewer.diff.scroll_repaint_us`, `viewer.diff.hunk_jump_to_visible_us`, `viewer.diff.expand_context_us`, `viewer.diff.viewport_rehydrate_us`, `viewer.diff.viewport_backtrack_us`, `viewer.diff.deferred_rows`, `viewer.diff.referenced_bytes_read`, `viewer.diff.viewport_referenced_bytes_read`, `viewer.diff.viewport_referenced_bytes_delta`, `viewer.diff.viewport_backtrack_referenced_bytes_delta`, `viewer.diff.placeholder_rows`, and `viewer.diff.placeholder_bands`. The protected expand-context path should exercise clickable hidden-banner reveal when that banner is visible, while scroll repaint plus viewport rehydrate/backtrack evidence remains the review surface for combo-sync and bounded rebuild work.

## Hex Byte Colors

When `hexByteColorMode` is `"leadingNibble"`:
- the offset column stays monochrome,
- visible bytes in the hex column are color-coded by leading nibble,
- the corresponding text-column glyphs are color-coded for the same visible bytes,
- `00` uses a subdued foreground bucket,
- `FF` uses an emphasized foreground bucket,
- only visible bytes are classified and painted; the viewer must not preprocess the whole file.

Foreground colorization must preserve the existing background highlight order:
1. row selection background
2. search-match background
3. active byte/selection background
4. text draw

If a byte color would be unreadable over the effective background, the glyph falls back to a contrast-safe foreground color.

## In-Viewer Menu

The visible top menu bar is rendered through the shared `RedSalamander.DxNativeMenuBar` host while the underlying command model continues to come from the hidden native viewer menu. `Alt`, `F10`, and menu mnemonics continue to work through that DxUi surface.

The viewer exposes the same option through `View -> Hex byte colors`:
- `Leading nibble (00/FF emphasized)`
- `Off`

Changing the menu selection updates the in-memory viewer configuration immediately, repaints the hex view, and persists through the normal viewer-instance configuration save path.

## Accessibility And Theme Rules

- Byte-to-color mapping is deterministic by byte value and must not depend on file path.
- Palette generation may adapt to theme flags such as `darkMode`, `darkBase`, and `rainbowMode`.
- In `highContrast`, hex byte coloring falls back to monochrome even when the saved mode is `"leadingNibble"`.

## Host Persistence Contract

The host may persist `builtin/viewer-text` configuration changes from:
- Preferences or the schema-driven plugin configuration dialog,
- in-view commands that modify `showLineNumbers`, `wrapText`, `hexByteColorMode`, `diffDefaultLayout`, or `diffContextMode`.

When a viewer instance closes, the host must persist configuration only if that instance changed its own configuration during the session. Closing an older viewer window must not overwrite newer persisted ViewerText settings.
Transient parsed-diff UI state such as active hunk/section selection, banner visibility, and theme-driven row paint decisions must not be serialized into the persisted ViewerText configuration.
