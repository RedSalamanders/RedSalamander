# RedConfigure UI Specification

## Shape

RedConfigure uses task modes rather than implementation pages:

1. `Start`
2. `Localization`
3. `Themes`
4. `Review & Export`

The left rail switches modes. The header shows the current mode title and purpose. A scope line keeps the active culture, owner, and theme visible while editing.

On `Start`, the scope line shows scan health only; culture and active owner are not shown there because they are edited in `Localization`.

## Start

`Start` is the workspace loading page. It keeps setup narrow so the user can first confirm that RedConfigure found the repository and scanned it correctly.

- workspace root
- discovered resource/theme counts
- scan errors

When launched from a build output folder such as `.build\x64\Debug`, the workspace root field must automatically normalize to the repository root when a parent contains the solution or Git marker. The user can still type another root and reload.

## Localization

`Localization` is a single workbench containing:

- culture selector
- active resource owner selector
- search and filters
- one primary resources table
- focused source/target editor
- placeholder validation status

The culture selector lists existing cultures discovered from `Lang\<culture>\` satellite resources first, then official Windows locale names as `new` targets. Culture entries must show both the culture code and a readable language/region name, for example `fr-FR - French (France)`. Selecting a `new` target creates that culture in the working model; the output file is created only on export.

Editing target text updates the in-memory model immediately. Invalid placeholder edits must be rejected and must leave the previous valid target text in the model.

The resource table is the main workflow surface. It shows ID, source, target, and status in one place. Header clicks cycle sorting for sortable columns. Search matches ID/source/target text, the ID filter narrows resource IDs, and the status filter can show all rows, OK rows, or problem rows. The visible translation count must show filtered rows versus total rows when a filter is active.

The Localization layout must have a compact mode for high-DPI or narrow first-open sizes. In compact mode, culture/owner controls and filters stack into separate rows, and the resource table, editor labels, edit boxes, and validation/status line occupy reserved vertical bands without overlap.

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

## Localization

All RedConfigure user-facing labels, descriptions, menu text, validation labels, and help text must live in `RedConfigure/RedConfigure.rc`.

Static menus must be `.rc` menu resources.
