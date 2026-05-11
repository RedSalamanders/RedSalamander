# RedConfigure Core Specification

## Goal

`RedConfigure.exe` is a first-party authoring tool for RedSalamander localization resources and theme files.

It edits a working model in memory, previews the result, validates it, and writes explicit output files only when the user exports.

## Outputs

- Localization output is a satellite `.rc` file.
- Theme output is a `.theme.json5` file.
- The tool must not silently modify live user settings.
- The tool must not commit, push, publish, or perform external side effects.

## Workspace Discovery

At startup, RedConfigure resolves the workspace root from the executable location. If the app is launched from a build folder, it walks parent folders until it finds `RedSalamander.sln` or `.git`, then uses that parent as the editable workspace root.

RedConfigure discovers resource owners from first-party `.vcxproj` files by reading `ResourceCompile` items. Build output directories such as `.build` are ignored.

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

Culture selection is part of the localization workflow. Existing cultures are discovered from satellite resource paths under `Lang\<culture>\`; new target cultures are chosen from the official Windows locale list and remain in-memory until export writes the satellite `.rc`. UI culture choices must include the culture code and a readable Windows locale display name.

Input `.rc` files must be decoded the same way the Windows resource compiler accepts first-party resources: UTF-16 LE with BOM, UTF-16 LE without BOM, UTF-16 BE with BOM, valid UTF-8, then the system ANSI code page as the final compatibility fallback. BOM-less UTF-16 detection must happen before UTF-8 or ANSI fallback so RC files emitted without a BOM are not parsed as NUL-interleaved text.

The localization table view is a projection over translation rows. It supports:

- free-text search across ID, source, and target text
- ID column filtering
- status filtering for all, OK, or problem rows
- stable source order when no sort is active
- stable sortable projections by ID, source, target, or status

Placeholder validation is export-blocking for:

- bare `{}`
- unindexed format specs such as `{:08X}`
- printf-style placeholders such as `%s`
- target positional placeholders whose indexes or per-index counts do not match the source placeholders

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
