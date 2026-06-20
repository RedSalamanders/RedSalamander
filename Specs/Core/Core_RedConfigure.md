# RedConfigure Core Specification

## Goal

`RedConfigure.exe` is a first-party authoring tool for RedSalamander localization resources and theme files.

It edits a working model in memory, previews the result, validates it, and writes explicit output files only when the user exports.

## Outputs

- Localization output is one or more target-language satellite `.rc` files.
- Theme output is a `.theme.json5` file.
- The tool must not silently modify live user settings.
- The tool must not commit, push, publish, or perform external side effects.

## Workspace Discovery

At startup, RedConfigure resolves the workspace root from the executable location. If the app is launched from a build folder, it walks parent folders until it finds `RedSalamander.sln` or `.git`, then uses that parent as the editable workspace root.

RedConfigure discovers resource owners from first-party `.vcxproj` files by reading `ResourceCompile` items. Generated or non-source trees are pruned before recursive traversal, including `.build`, `.git`, `.vs`, `vcpkg_installed`, architecture output folders such as `x64`, and archived `Specs/TestRuns` evidence.

Startup UX is protected by an immediate splash screen on a worker UI thread. The splash may show status text for workspace detection, workspace loading, and first-frame preparation, but it must not move the expensive workspace scan onto the splash thread. The main window remains responsible for closing the splash after `PrimeForShow()` has prepared the first frame. Startup instrumentation uses:

- `redconfigure.startup.create_window_us`
- `redconfigure.startup.on_create_us`
- `redconfigure.startup.workspace_reload_us`
- `redconfigure.startup.splash.visible_us`
- `redconfigure.workspace.discover_us`
- `redconfigure.theme_catalog.load_us`
- `redconfigure.localization_review.load_us`
- `redconfigure.localization_review.owner_load_us`
- `redconfigure.localization_review.worker_count`
- `redconfigure.localization_active_owner.load_us`

Each resource owner has:

- owner name
- project path
- embedded English `.rc` path
- optional satellite `.rc` paths under `Lang\<culture>\`

Theme files are discovered from selected theme folders and from `*.theme.json5` files that are not inside build output directories.

## Localization Model

The localization model supports:

- `STRINGTABLE` entries
- static `MENU`, `POPUP`, and `MENUITEM` captions as inventory
- dialog captions and static/control text as inventory when safely parsed

The first writable output scope is `STRINGTABLE`. Unsupported or not-yet-writable resource forms must remain visible as inventory and must not be rewritten silently.

English embedded `.rc` resources are the localization master. RedConfigure may display English source text, but it must not edit or export embedded English source files from the localization workbench.

Target-language selection is part of the localization workflow. Existing cultures are discovered from satellite resource paths under `Lang\<culture>\`. The session can ensure a new non-English culture in memory; export writes the satellite `.rc` only when requested. UI language filters operate on the loaded review culture names. `en-US` is the embedded source language and must never be offered as a review target culture: culture discovery excludes a `Lang\en-US` satellite if one exists, and `EnsureLocalizationReviewCulture` rejects `en-US` as a backstop.

Input `.rc` files must be decoded the same way the Windows resource compiler accepts first-party resources: UTF-16 LE with BOM, UTF-16 LE without BOM, UTF-16 BE with BOM, valid UTF-8, then the system ANSI code page as the final compatibility fallback. BOM-less UTF-16 detection must happen before UTF-8 or ANSI fallback so RC files emitted without a BOM are not parsed as NUL-interleaved text.

Localization loading is best-effort at the file level. A resource file that cannot be read or parsed must append a path-specific diagnostic to the workspace errors list and the session must continue loading the remaining owners/cultures. A bad satellite file must not prevent the embedded source inventory for the active owner from loading. A bad embedded source file skips that owner for review/inventory data but must not fail the whole workspace load.

The all-owner localization review load may parse owners in bounded parallel worker chunks during startup. Worker tasks must keep parse results and diagnostics thread-local and merge them on the caller thread in resource-owner order so row order, exported file ordering, and workspace error reporting remain deterministic.

The localization review model is keyed by resource owner plus string ID. It can load all discovered application and plugin resource owners at once. Each review row contains:

- owner name
- string resource ID
- read-only English source text
- one editable target cell per selected non-English culture

The selected-cell context is owner name plus string resource ID. The UI must keep that context local to the read-only English source display, and English source text must remain read-only.

The localization table view is a projection over review rows. It supports:

- owner filtering by deduplicated owner display name, including an all-owners state; selecting a duplicated name includes all underlying owners with that name
- language filtering by selected target-language names, including an all-target-languages state
- free-text search across owner, ID, English source, and visible target-language text
- ID column filtering
- status filtering for all, OK, or problem rows; a row has problems when any visible target-language cell fails placeholder validation or has no existing translation (the cell still displays the English fallback text)
- stable source order when no sort is active
- stable sortable projections by owner, ID, English source, visible target language, or status

Clean review rows may render an empty status cell; problem rows must render the validation problem text. Rows whose visible target cells have no existing translation and no pending edit must render a distinct `Missing translation` status (validation problem text wins when both apply) and a row tone distinct from the validation-warning tone, so untranslated strings are visible in the review instead of silently showing the English fallback as fine. The review table must support at least two visible lines of source/target information per row.

The target editing surface displays selected non-English cultures as stacked editable rows. Source and target editor row heights are derived from explicit newline counts and width-based wrapping in the parsed resource strings so `\n` content and wrapped long text are not collapsed into a single-line control. When selected target languages exceed the available editor viewport, the UI must expose vertical scrolling instead of shrinking rows below their content-derived height.

Placeholder validation is export-blocking for:

- bare `{}`
- unindexed format specs such as `{:08X}`
- printf-style placeholders such as `%s`
- target positional placeholders whose indexes or per-index counts do not match the source placeholders

Localization review export writes target-language satellite `.rc` files only for dirty owner/culture pairs and must never rewrite embedded English source resources. Each changed owner/culture export contains the complete generated satellite string table for that owner/culture so existing unchanged translations are preserved. By design, cells without an existing translation export the English fallback text; the `Missing translation` review status is a review-time signal only and does not change export behavior. Writers must preserve deterministic ordering so output diffs remain reviewable.

## Theme Model

Theme files use the current `ThemeDefinition` JSON5 shape:

```json5
{
  "id": "user/example",
  "name": "Example",
  "baseThemeId": "builtin/dark",
  "colors": {
    "app.accent": "#2ECC71"
  }
}
```

RedConfigure may accept richer authored theme values in the editor, including:

- `ref(app.accent)`
- `darken(app.accent,20%)`
- `lighten(app.accent,20%)`
- `blend(menu.background,app.accent,16%)`
- `alpha(folderView.itemBackgroundSelected,45%)`
- `contrast(folderView.background)`

Until durable runtime support for a `colorExpressions` schema lands, RedConfigure exports expression-authored colors as flattened direct `colors` values in `.theme.json5`.

Expression evaluation rules:

- references resolve against authored overrides, other expressions, and base/default preview colors
- missing references are invalid
- dependency cycles are invalid
- invalid edits keep the previous valid preview color
- direct color edits replace any expression for the same key

Theme preview tokens must include enough defaults for app/window, navigation, menu, folder-view, dialog, progress, and diff examples to update even when the active theme omits a key. The editor color-key grid is the union of those preview defaults and the active theme's authored `colors` keys. The UI can filter this key list by case-insensitive substring, but batch group transforms operate against the complete key list for the selected group. Preview-hit regions map back to color keys shown in the editor. When preview regions overlap, hit resolution prefers the smallest region first and can cycle through containing regions on repeated clicks.

Theme authoring UI must offer expression/value suggestions derived from the active token, the current effective color, the accent token, and the previously selected token. These suggestions are editor assistance only; stored/exported theme values still follow the theme model and export rules above.

Batch group edits are direct-color transforms applied to the current color-key group. They must update the in-memory theme model and export preview immediately.

## Validation

Export is blocked by errors and allowed with warnings after the user has seen the warning list.

Writers must use deterministic ordering so output diffs remain reviewable.

## Icon

RedConfigure owns an application icon resource and must register it with the main window class. If icon loading fails, the app may fall back to the default Windows application icon.
