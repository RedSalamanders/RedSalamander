# RedConfigure UI Specification

## Shape

RedConfigure uses task modes rather than implementation pages:

1. `Start`
2. `Localization`
3. `Themes`
4. `Review & Export`

The left rail switches modes. When the window is narrow, the rail collapses to icon-only Segoe/Fluent glyph buttons with tooltips. The header shows the current mode title and scope line only; pages must not spend vertical space on a descriptive subtitle. The scope line keeps the current workspace breadth visible, such as checked owners/languages in Localization and scan health on Start.

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

Owner and language selectors use the shared DxUi tag picker control. Selected tags are badges inside the picker input frame, before the editable text/drop-down area. If the badges do not fit on one line, the picker grows to two or more internal rows and wraps badges before the text/drop-down area instead of placing badges outside the control or hiding them. The drop-down starts with `All owners` or `All languages`, followed by the available deduplicated options. Typing in the picker filters/autosuggests matching options. Arrow keys move the highlighted suggestion without adding a badge; Enter or mouse selection commits the highlighted suggestion. Committing a concrete suggestion adds a tag, removes that option from the suggestion list, and hides the all-option while any concrete tag is selected. Removing a tag puts that option back into the suggestion list, and the all-option returns when no concrete tag remains. Selecting all options collapses the display to the single all-tag; while the all-tag is active, concrete suggestions remain available and committing one replaces the all-tag with that concrete selection. Each visible tag has an `x` affordance that removes that selection. Owner options must be deduplicated by display name so repeated plugin/application owner names do not create duplicate selector entries. Language options must never include `en-US`: English is the read-only source language, so offering it would create a dead option that silently reverts.

English source text is read-only. The selected-cell owner/ID context must be displayed next to the `English source` label so the text and its row identity stay local to each other. The English edit box must align its left edge with the target-language edit boxes, leaving the culture-label gutter consistent across source and target rows. Editing target text updates the in-memory model immediately for the edited non-English language cell. All selected target languages are editable at the same time through stacked full-width editor rows. Source and target editors are multiline, and each row must size from explicit `\n` line breaks and width-based wrapping so newline-bearing or wrapped resources are visible without being forced into a single-line-height control.

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

Amounts may use `0.0` to `1.0` or percentage syntax such as `20%`.

Valid edits update the live preview immediately. Invalid edits show an error state and keep the last valid preview.

The value editor provides guided suggestions while typing. Suggestions must include the current direct color when known, common commands such as `ref`, `darken`, `lighten`, `alpha`, `contrast`, and `blend`, and contextual suggestions based on the previously selected token. The previous token is shown beside a small swatch so users can derive or blend from it without remembering the key.

The token grid must expose the previewed theme settings plus any authored color keys from the active theme, including app/window, navigation, menu, folder-view, dialog, progress, and diff sample colors. The grid shows at least the key, effective value with a swatch, and authored value or expression. A key filter narrows the grid by case-insensitive group or substring, such as `menu`, `accent`, or `background`.

Clicking a visible preview region selects the corresponding token so the user can edit from the example instead of hunting through the grid. Preview hit regions must map to the key that visibly drives that region; for example, the preview accent stripe selects `app.accent`. When preview regions overlap, the smallest visible region wins first, and repeated clicks at the same point cycle through containing regions. The selected token must be highlighted in the live preview while it is being edited.

Batch controls may apply direct color transforms to the current token group. A group is the part before the first dot, such as `menu` or `folderView`.

## Review & Export

`Review & Export` shows:

- default localization output path
- default theme output path
- localization `.rc` preview
- theme `.theme.json5` preview
- explicit export actions

The page must make output paths and generated text visible before writing existing files.

The localization `.rc` preview lists exactly the changed target-language satellite files that the export action will write: one generated file per dirty owner/culture pair. Export writes only those changed target-language satellite files and must not rewrite embedded English resources.

## Localization

All RedConfigure user-facing labels, descriptions, menu text, validation labels, and help text must live in `RedConfigure/RedConfigure.rc`.

Static menus must be `.rc` menu resources.
