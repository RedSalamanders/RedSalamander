# RedConfigure UI Specification

## Shape

RedConfigure uses task modes rather than implementation pages:

1. `Start`
2. `Localization`
3. `Themes`
4. `Review & Export`

The left rail switches modes. When the window is narrow, the rail collapses to icon-only Segoe/Fluent glyph buttons with tooltips. The header shows the current mode title and scope line only; pages must not spend vertical space on a descriptive subtitle. The scope line keeps the current workspace breadth visible, such as checked owners/languages in Localization and scan health on Start.

Below the scope line, every mode shares a command bar with global search, Validate, Undo, Redo, Validation, and Review & Export actions. Global search routes to resource text in Localization and token/group text in Themes. Review & Export navigates to the review page; it never writes directly from another mode. A collapsible, modeless validation drawer remains available above the status line in every mode. Normal findings use that drawer/status area and must not open blocking message boxes.

On `Start`, the scope line shows scan health only; culture and active owner are not shown there because they are edited in `Localization`.

RedConfigure shows a lightweight splash immediately at process startup while the main window discovers and loads the workspace. The splash must be owned by a separate UI thread so its animation continues while the main thread performs synchronous startup work. It closes automatically after the main window has prepared its first visible frame, and close requests before the splash window is created must suppress the pending open.

## Start

`Start` is the workspace loading page. It keeps setup narrow so the user can first confirm that RedConfigure found the repository and scanned it correctly.

- workspace root
- discovered resource/theme counts
- scan errors

When launched from a build output folder such as `.build\x64\Debug`, the workspace root field must automatically normalize to the repository root when a parent contains the solution or Git marker. The user can still type another root and reload.

## Localization

`Localization` is a single workbench containing:

- owner tag picker with removable tags and an `All owners` command
- target-language tag picker with removable tags and an `All languages` command
- search and filters
- one primary localization review table
- read-only English source editor with selected-cell context in the same header row
- stacked target editors for every selected non-English language
- placeholder validation status shown beside the English source header only when there is a problem
- ordered/pinnable language controls, previous/next problem navigation, rectangular matrix paste, batch preview/apply, and a live resource example

Owner and language selectors use the shared DxUi tag picker control. Selected tags are badges inside the picker input frame, before the editable text/drop-down area. If the badges do not fit on one line, the picker grows to two or more internal rows and wraps badges before the text/drop-down area instead of placing badges outside the control or hiding them. The drop-down starts with `All owners` or `All languages`, followed by the available deduplicated options. Typing in the picker filters/autosuggests matching options. Arrow keys move the highlighted suggestion without adding a badge; Enter or mouse selection commits the highlighted suggestion. Committing a concrete suggestion adds a tag, removes that option from the suggestion list, and hides the all-option while any concrete tag is selected. Removing a tag puts that option back into the suggestion list, and the all-option returns when no concrete tag remains. Selecting all options collapses the display to the single all-tag; while the all-tag is active, concrete suggestions remain available and committing one replaces the all-tag with that concrete selection. Each visible tag has an `x` affordance that removes that selection. Owner options must be deduplicated by display name so repeated plugin/application owner names do not create duplicate selector entries. Language options must never include `en-US`: English is the read-only source language, so offering it would create a dead option that silently reverts.

English source text is read-only. The selected-cell owner/ID context must be displayed next to the `English source` label so the text and its row identity stay local to each other. The English edit box must align its left edge with the target-language edit boxes, leaving the culture-label gutter consistent across source and target rows. Editing target text updates the in-memory model immediately for the edited non-English language cell. All selected target languages are editable at the same time through stacked full-width editor rows. Source and target editors are multiline, and each row must size from explicit `\n` line breaks and width-based wrapping so newline-bearing or wrapped resources are visible without being forced into a single-line-height control.

Activating a matrix row transfers focus to its first visible target editor, providing direct commit/cancel editing through standard DxUi text-field behavior. The inspector shows selected owner/ID, source, neighboring visible targets, placeholder tokens, accelerator, validation status, and a live example. Examples substitute deterministic sample argument `42` for `{0}` and update with each accepted target edit. Matrix copy uses the grid's tab/newline representation; Paste matrix applies a rectangular TSV block starting at the selected row and culture.

Localization batch operations require two invocations: the first shows a change count plus representative before/after value; the second applies atomically. Supported operations are copy English, copy culture to culture, clear, find/replace (`find=replace` from the command field), normalize placeholder whitespace, preserve source accelerators, and mark reviewed.

Placeholder validation is problem-first. A clean selected cell must not show an `OK` validation line. When the selected source/target set has a validation problem, the status text appears beside the English source header, uses the theme error color, and renders with a bold/strong font.

The localization review table is the main workflow surface. It shows owner, ID, English source, one column for each checked target language, and status in one place. Review rows must be tall enough for at least two lines of source/target information, and source/target cells must wrap to at least two visible lines. Clean rows leave the status cell blank; problem rows show the validation status; rows whose visible target cells have no existing translation show the `Missing translation` status with a distinct row tone (validation problems take precedence). Header clicks cycle sorting for sortable columns. Search matches owner, ID, English source, and visible target-language text; the ID filter narrows resource IDs; and the status filter can show all rows, OK rows, or problem rows — missing-translation rows count as problem rows. The visible translation count must show filtered rows versus total rows when a filter is active.

The Localization layout must have a compact mode for high-DPI or narrow first-open sizes. In compact mode, owner selector, language selector, and filters stack into reserved rows, and the review table, stacked editor labels, edit boxes, and problem-only validation status occupy reserved vertical bands without overlap. The workbench body must have a page-level vertical scroller whenever the filters, two-line review table, and fully sized editor rows exceed the available viewport; the table must not collapse below its two-row minimum to avoid page scrolling.

The earlier separate inventory table is not part of the user workflow. Unsupported or not-yet-editable resource forms may remain in the internal inventory model, but the visible Localization page should not show two unrelated tables.

## Themes

`Themes` is a single workbench containing:

- active theme selector
- active theme identity
- theme token filter
- theme token grid
- color or expression editor
- color status
- expression examples
- batch group edits
- live preview
- built-in/file/user library labels, import/duplicate/reset, scene selection, alpha control, copy-effective/copy-override, token metadata/contrast, and recipe preview/apply

At the initial window size, the theme editor and preview must not overlap each other or the status line. When there is enough horizontal room, the editor sits beside the live preview so the preview has meaningful vertical space immediately. The preview surface is vertically scrollable whenever its examples exceed the available viewport.

The color editor accepts direct colors and expression text:

- `#RRGGBB`
- `#AARRGGBB`
- `ref(key)`
- `darken(key,amount)`
- `lighten(key,amount)`
- `blend(firstKey,secondKey,amount)`
- `alpha(key,amount)`
- `contrast(key)`
- `perceptualTone(key,tone)`
- `ensureContrast(foregroundKey,backgroundKey,ratio)`
- `harmonize(key,targetKey,amount)`
- `systemAccent()` and `systemColor(role)`
- `tone(lightKey,darkKey)`
- `seededRainbow(runtime.seed,saturation,value,alpha,hueOffset)`
- `seededChoice(runtime.seed,key1,key2[,key3...key8])`

Amounts may use `0.0` to `1.0` or percentage syntax such as `20%`.

Formatter/export output uses the canonical camelCase function spellings above; parsing remains case-insensitive for
compatibility. `ensureContrast` measures the foreground as rendered over its background. The background must be
opaque because the expression has no surface-backdrop input. Foreground alpha is preserved only when a candidate at
that alpha can meet the requested ratio; otherwise the edit is invalid and the previous valid preview remains.

The theme editor exposes version 2 palette entries separately from semantic tokens. Authors can add, rename, remove, and reference palette entries. Creating a palette entry from a repeated direct literal rewrites matching sources to references. Rename rewrites references, while delete is blocked until displayed dependents are removed. Batch darken/blend actions remain authored as palette-backed functions rather than flattened colors. The dependency display identifies palette and semantic references, reports missing references and cycles, and labels load-time, event-time, and paint-time sources. Expressions are not nested; the UI guides authors to create a named palette entry for an intermediate result.

Valid edits update the live preview immediately. Invalid edits show an error state and keep the last valid preview.

The value editor provides guided suggestions while typing. Suggestions must include the current direct color when known, every supported version 2 function, palette and semantic references, and contextual suggestions based on the previously selected token. The previous token is shown beside a small swatch so users can derive or blend from it without remembering the key.

The token grid must expose the previewed theme settings plus any authored color keys from the active theme, including app/window, navigation, menu, folder-view, dialog, progress, and diff sample colors. The grid shows key, group, effective value with a swatch, authored value/expression, source type, dependent-use count, and known foreground/background contrast ratio with AA status. A key filter narrows the grid by case-insensitive group or substring, such as `menu`, `accent`, or `background`.

Clicking a visible preview region selects the corresponding token so the user can edit from the example instead of hunting through the grid. Preview hit regions must map to the key that visibly drives that region; for example, the preview accent stripe selects `app.accent`. When preview regions overlap, the smallest visible region wins first, and repeated clicks at the same point cycle through containing regions. The selected token must be highlighted in the live preview while it is being edited.

Batch controls may apply direct color transforms to the current token group. A group is the part before the first dot, such as `menu` or `folderView`.

The scene selector provides App Shell, Folder View, Menu Popup, Dialogs, File Operations, Monitor Log, and Viewer Diff views and narrows the token navigator to that semantic group while the composite sample remains available. The authored value combo is the hex/expression field and includes recent/copied values. The alpha slider supplies the set-alpha recipe. Recipe/mass changes use two-step review and include dark/light variants, accent recolor, softened selections, increased contrast, semantic status colors, alpha, reference replacement, solid-to-palette conversion, and override removal.

Theme and localization mass changes use the shared typed approval states `NoChanges`, `Ready`, `Stale`,
`Invalid`, and `Applied`. A ready preview is bound to the complete request and the complete theme snapshot or
stable localization `(owner, resource ID, culture)` identity plus its recorded before-value. Approval preflights
every change before mutation. Any request change, intervening edit, reload, filter/sort rebuild, invalid candidate,
or failed apply clears the pending approval. A stale or invalid preview changes nothing; a successful batch is one
session Undo step.

Reference replacement edits exact parsed reference nodes only and never partial names. Solid-to-palette conversion
accepts direct sources only and requires an existing palette target. All ten recipes validate the complete
candidate theme before the preview is shown. Theme numeric grammar is locale-invariant: `.` is the only decimal
separator, input must be finite and fully consumed, and comma-decimal forms are rejected under every process locale.

The deterministic Debug contract keeps the 6-owner/1,500-row/10-theme scan below 500 ms, validation below 250 ms,
a single theme edit below 16 ms, and a validated 512-token explicit mass preview below 100 ms. Curated Track 19
evidence is archived under `Specs/TestRuns/4cb089111a23/RedConfigure/2026-07-17_1548_observatory_track19/`.

When Themes is inactive, its preview detaches from the model so an inactive page does not retain an active heavy preview surface. Returning to Themes reconnects the model and repaints current state.

## Review & Export

`Review & Export` shows:

- default localization output path
- default theme output path
- localization `.rc` preview
- theme `.theme.json5` preview
- explicit export actions

The page must make output paths and generated text visible before writing existing files.

The two previews are output-file cards in the review basket. Combined errors block export. Warnings require the user to see the issue list and invoke export again. Successful export shows the written path; each writer has already reopened and reparsed its written file before success is reported.

Validation carries typed category, code, severity, and structured arguments. Workflow correctness never searches
localized diagnostic text. Category/message formatting happens only at the presentation boundary through resource
strings.

Theme preview and export display the same authored version 2 JSON5 representation, including `formatVersion`, `palette`, and expressions. Export never replaces expressions with resolved hex values.

The localization `.rc` preview lists exactly the changed target-language satellite files that the export action will write: one generated file per dirty owner/culture pair. Export writes only those changed target-language satellite files and must not rewrite embedded English resources.

## Localization

All RedConfigure user-facing labels, descriptions, menu text, validation labels, and help text must live in `RedConfigure/RedConfigure.rc`.

Theme origin labels, dirty-scope text, duplicate-theme naming, validation categories/messages, token descriptions,
and source/contrast labels follow the same resource rule. Duplicate-theme creation returns a typed result, reserves
space for suffixes within the 64-character ID/name limits, and retries only identifier collisions.

Static menus must be `.rc` menu resources.
